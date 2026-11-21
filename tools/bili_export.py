#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""导出 B 站收藏夹到 CSV / Markdown（本地运行，不上传任何数据）。

用法（PowerShell，用本机 DSH 自带 Python）:
    $env:BILI_SESSDATA = Read-Host "粘贴 SESSDATA"
    & "C:\\Users\\ww7373\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe" "D:\\ai1\\a\\tools\\bili_export.py"
    Remove-Item Env:BILI_SESSDATA

SESSDATA 获取方式:
    浏览器登录 bilibili.com -> F12 -> Application(应用) -> Cookies ->
    https://www.bilibili.com -> 复制 SESSDATA 一行的 Value。

输出:
    D:\\ai1\\a\\bili_export\\<时间戳>\\favorites.csv   （Excel 可直接打开）
    D:\\ai1\\a\\bili_export\\<时间戳>\\summary.md      （汇总与重复/失效清单）

安全: 凭证只在本进程内存中使用，不写盘、不上传。
"""

import argparse
import csv
import getpass
import hashlib
import json
import os
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime

API = "https://api.bilibili.com"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)
REFERER = "https://www.bilibili.com/"

# B 站 wbi 签名使用的混淆表
MIXIN_TAB = [
    46, 47, 18, 2, 53, 8, 23, 32, 15, 50, 10, 31, 58, 3, 45, 35,
    27, 43, 5, 49, 33, 9, 42, 19, 29, 28, 14, 39, 12, 38, 41, 13,
    37, 48, 7, 16, 24, 55, 40, 61, 26, 17, 0, 1, 60, 51, 30, 4,
    22, 25, 54, 21, 56, 59, 6, 63, 57, 62, 11, 36, 20, 34, 44, 52,
]


class BiliError(Exception):
    pass


def _get_json(url, cookie, retries=3, timeout=20):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": UA, "Referer": REFERER, "Cookie": cookie},
    )
    last = None
    for i in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", "replace")
            try:
                data = json.loads(raw)
            except json.JSONDecodeError:
                raise BiliError("返回内容不是 JSON（可能被风控拦截），请稍后重试")
            code = data.get("code")
            if code == 0:
                return data.get("data")
            msg = data.get("message") or data.get("msg") or ""
            if code == -101:
                raise BiliError("登录态无效：SESSDATA 过期或错误，请重新从浏览器复制")
            if code == -352:
                raise BiliError("触发风控（-352）：请等待几分钟后再试，或更换网络")
            if code == -403:
                raise BiliError("无权限（-403）：该收藏夹不是你的，或不可见")
            last = BiliError(f"接口返回 code={code} msg={msg}")
        except BiliError:
            raise
        except Exception as exc:  # 网络错误重试
            last = exc
        time.sleep(0.6 * (i + 1))
    raise BiliError(f"请求失败：{last}")


def _api(path, params, cookie):
    q = urllib.parse.urlencode(params or {})
    url = API + path + (("?" + q) if q else "")
    return _get_json(url, cookie)


def get_wbi_mixin(cookie):
    data = _api("/x/web-interface/nav", {}, cookie)
    img = str(data["wbi_img"]["img_url"]).rsplit("/", 1)[-1].split(".")[0]
    sub = str(data["wbi_img"]["sub_url"]).rsplit("/", 1)[-1].split(".")[0]
    orig = img + sub
    return "".join(orig[i] for i in MIXIN_TAB)[:32]


def wbi_sign(params, mixin):
    p = dict(params)
    p["wts"] = int(time.time())
    parts = []
    for key in sorted(p):
        val = "".join(ch for ch in str(p[key]) if ch not in "!'()*")
        parts.append(
            urllib.parse.quote(str(key), safe="")
            + "="
            + urllib.parse.quote(val, safe="")
        )
    query = "&".join(parts)
    p["w_rid"] = hashlib.md5((query + mixin).encode("utf-8")).hexdigest()
    return p


def get_buvid(cookie):
    try:
        data = _api("/x/frontend/finger/spi", {}, cookie)
        return data.get("b_3"), data.get("b_4")
    except BiliError:
        return None, None


def fmt_duration(sec):
    try:
        sec = int(sec)
    except (TypeError, ValueError):
        return ""
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def ts_to_str(ts):
    try:
        return datetime.fromtimestamp(int(ts)).strftime("%Y-%m-%d %H:%M")
    except (TypeError, ValueError, OSError):
        return ""


def list_folders(mid, cookie, mixin):
    data = _api(
        "/x/v3/fav/folder/created/list-all",
        wbi_sign({"up_mid": mid}, mixin),
        cookie,
    )
    return (data or {}).get("list") or []


def list_resources(media_id, cookie, mixin, page_size=20, delay=0.5):
    items = []
    pn = 1
    while True:
        params = wbi_sign(
            {"media_id": media_id, "pn": pn, "ps": page_size, "platform": "web"},
            mixin,
        )
        data = _api("/x/v3/fav/resource/list", params, cookie)
        for m in (data or {}).get("medias") or []:
            items.append(m)
        if not (data or {}).get("has_more"):
            break
        pn += 1
        time.sleep(delay)
    return items


def export(cookie, out_dir, max_per_folder=0):
    nav = _api("/x/web-interface/nav", {}, cookie)
    mid = nav.get("mid")
    uname = nav.get("uname") or "unknown"
    mixin = get_wbi_mixin(cookie)

    folders = list_folders(mid, cookie, mixin)
    if not folders:
        raise BiliError("没有找到任何收藏夹（或全部为空）")

    rows = []
    stats = []
    for folder in folders:
        fid = folder.get("id")
        title = folder.get("title") or str(fid)
        print(f"  读取收藏夹：{title} ...", end="", flush=True)
        try:
            medias = list_resources(fid, cookie, mixin)
        except BiliError as exc:
            print(f" 失败（{exc}）")
            continue
        if max_per_folder:
            medias = medias[:max_per_folder]
        invalid = 0
        for m in medias:
            v_title = (m.get("title") or "").strip()
            bvid = m.get("bvid") or ""
            is_invalid = (not bvid) or v_title in ("已失效视频", "已删除视频")
            if is_invalid:
                invalid += 1
            rows.append(
                {
                    "收藏夹": title,
                    "标题": v_title or "（无标题）",
                    "UP主": ((m.get("upper") or {}).get("name") or ""),
                    "时长": fmt_duration(m.get("duration")),
                    "链接": f"https://www.bilibili.com/video/{bvid}" if bvid else "",
                    "收藏时间": ts_to_str(m.get("fav_time")),
                    "发布时间": ts_to_str(m.get("pubtime")),
                    "状态": "失效" if is_invalid else "正常",
                }
            )
        stats.append({"title": title, "count": len(medias), "invalid": invalid})
        print(f" {len(medias)} 条（失效 {invalid}）")

    os.makedirs(out_dir, exist_ok=True)
    csv_path = os.path.join(out_dir, "favorites.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(
            fh, fieldnames=["收藏夹", "标题", "UP主", "时长", "链接", "收藏时间", "发布时间", "状态"]
        )
        writer.writeheader()
        writer.writerows(rows)

    # 重复检测（跨收藏夹按 bvid）
    seen = {}
    for r in rows:
        link = r["链接"]
        if link:
            seen.setdefault(link, []).append(r["收藏夹"])
    dups = {k: v for k, v in seen.items() if len(v) > 1}

    md_path = os.path.join(out_dir, "summary.md")
    with open(md_path, "w", encoding="utf-8") as fh:
        fh.write("# B 站收藏夹导出报告\n\n")
        fh.write(f"- 导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        fh.write(f"- 账号：{uname}（UID {mid}）\n")
        fh.write(f"- 收藏夹数：{len(stats)}\n")
        fh.write(f"- 视频总数：{len(rows)}\n")
        fh.write(f"- 失效视频：{sum(s['invalid'] for s in stats)}\n")
        fh.write(f"- 跨收藏夹重复：{len(dups)} 个\n\n")
        fh.write("## 各收藏夹\n\n| 收藏夹 | 视频数 | 失效 |\n|---|---:|---:|\n")
        for s in stats:
            fh.write(f"| {s['title']} | {s['count']} | {s['invalid']} |\n")
        if dups:
            fh.write("\n## 重复视频（出现在多个收藏夹）\n\n")
            for link, names in sorted(dups.items(), key=lambda kv: -len(kv[1])):
                fh.write(f"- {link} — 出现在：{'、'.join(names)}\n")

    return csv_path, md_path, stats, rows, dups


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="导出 B 站收藏夹到 CSV/Markdown")
    parser.add_argument("--out", default=None, help="输出目录，默认 bili_export/<时间戳>")
    parser.add_argument("--max", type=int, default=0, help="每个收藏夹最多导出条数，0=全部")
    args = parser.parse_args()

    sess = os.environ.get("BILI_SESSDATA", "").strip()
    if not sess:
        try:
            sess = getpass.getpass("请输入 SESSDATA（输入不回显）: ").strip()
        except (EOFError, KeyboardInterrupt):
            sess = ""
    if not sess:
        print("未提供 SESSDATA。获取方法见脚本头部注释。")
        return 2

    cookie = f"SESSDATA={sess}"
    try:
        b3, b4 = get_buvid(cookie)
        if b3:
            cookie += f"; buvid3={b3}"
        if b4:
            cookie += f"; buvid4={b4}"

        base = os.path.dirname(os.path.abspath(__file__))
        out_dir = args.out or os.path.join(
            base, "..", "bili_export", datetime.now().strftime("%Y%m%d-%H%M%S")
        )
        out_dir = os.path.abspath(out_dir)

        print("正在读取收藏夹（只读，不会修改任何内容）...")
        csv_path, md_path, stats, rows, dups = export(cookie, out_dir, args.max)
    except BiliError as exc:
        print(f"导出失败：{exc}")
        return 1

    print("\n完成：")
    print(f"  明细：{csv_path}")
    print(f"  汇总：{md_path}")
    print(f"  共 {len(rows)} 条，跨收藏夹重复 {len(dups)} 个")
    print("\n提示：SESSDATA 仅本次进程使用，已随进程结束失效。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
