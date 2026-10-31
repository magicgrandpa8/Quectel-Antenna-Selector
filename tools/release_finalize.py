#!/usr/bin/env python3
"""
release_finalize.py — 新版 Excel 的第二階段：合併 datasheet 連結、擷取照片、建置 PWA、檢查新天線

用法 (在專案根目錄)：
    python3 tools/release_finalize.py --xlsx <已寫入 Histories 的 Excel> --version 3.26.0 \
        [--links <datasheet_links_RevYY.MM.csv>] [--pdf-dir <新 datasheet PDF 資料夾>] [--diff <release_prepare 的 json>]

步驟：
  1. --links：把 CSV (PN, Type, Datasheet_File, OneDrive_URL) 合併進 data/datasheet_links.csv
     (同料號以新連結取代；空白連結不覆蓋既有連結)
  2. --pdf-dir：用 tools/extract_datasheet_photos.py 擷取產品照片 (合併進 photo_datasheet.json)
  3. 執行 build_data.py 產生 index.html / photos.json / sw.js / version.json
  4. 檢查：每款新天線是否有 datasheet 照片與 Datasheet 連結，並列出照片解析度
"""
from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LINKS = ROOT / "data" / "datasheet_links.csv"


def merge_links(src: Path) -> list[str]:
    raw = LINKS.read_bytes()
    bom, crlf = raw.startswith(b"\xef\xbb\xbf"), b"\r\n" in raw
    rows = list(csv.reader(open(LINKS, encoding="utf-8-sig", newline="")))
    hdr, body = rows[0], rows[1:]
    idx = {r[0].strip(): i for i, r in enumerate(body)}
    changed = []
    for r in csv.DictReader(open(src, encoding="utf-8-sig", newline="")):
        pn, url = (r.get("PN") or "").strip(), (r.get("OneDrive_URL") or "").strip()
        if not pn:
            continue
        chk = (r.get("Anonymous_Check") or "").strip().upper()
        if url and chk and chk != "OK":                  # 未登入無法開啟：不放上工具
            print(f"  略過 {pn}：未登入驗證 {chk}，連結不寫入工具 (請確認分享設定為「任何人」)")
            url = ""
        row = [pn, (r.get("Type") or "").strip(), (r.get("Datasheet_File") or "").strip(), url]
        if pn in idx:
            if url and body[idx[pn]][3].strip() != url:      # 空白連結不覆蓋既有連結
                body[idx[pn]] = row
                changed.append(pn)
        else:
            body.append(row)
            idx[pn] = len(body) - 1
            changed.append(pn)
    with open(LINKS, "w", encoding="utf-8-sig" if bom else "utf-8", newline="") as f:
        csv.writer(f, lineterminator="\r\n" if crlf else "\n").writerows([hdr] + body)
    return changed


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--version", required=True)
    ap.add_argument("--links", type=Path)
    ap.add_argument("--pdf-dir", type=Path)
    ap.add_argument("--diff", type=Path, help="release_prepare.py 輸出的 json (用來檢查新天線)")
    ap.add_argument("--photo-xlsx", type=Path, help="(選用) 其他 Excel 的清晰照片")
    a = ap.parse_args()

    if a.links:
        ch = merge_links(a.links)
        print(f"datasheet 連結：合併 {len(ch)} 筆 {ch}")
    if a.pdf_dir:
        r = subprocess.run([sys.executable, str(ROOT / "tools" / "extract_datasheet_photos.py"), str(a.pdf_dir)],
                           cwd=ROOT, capture_output=True, text=True)
        print("照片擷取：" + "\n".join(l for l in r.stdout.splitlines() if l.startswith("PDF")))
        if r.returncode:
            sys.exit(r.stderr)
        for f in ("photo_datasheet_review.png", "photo_datasheet_report.csv"):
            p = ROOT / f
            if p.exists():
                p.replace(Path("/tmp") / f)                 # 檢查用檔案移出 repo (保留在 /tmp 供檢視)
    cmd = [sys.executable, str(ROOT / "build_data.py"), "--version", a.version, "--xlsx", str(a.xlsx)]
    if a.photo_xlsx:
        cmd += ["--photo-xlsx", str(a.photo_xlsx)]
    r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True)
    print("\n".join(l for l in r.stdout.splitlines() if l.strip()))
    if r.returncode:
        sys.exit(r.stderr)

    # 檢查新天線
    if a.diff and a.diff.exists():
        diff = json.loads(a.diff.read_text(encoding="utf8"))
        from PIL import Image
        s = (ROOT / "index.html").read_text(encoding="utf8")
        i = s.index("const DATA = ") + 13
        items = {p["oc"]: p for p in json.loads(s[i:s.index("\n", i)].rstrip(";"))["items"]}
        ds = json.loads((ROOT / "photo_datasheet.json").read_text(encoding="utf8"))
        ph = json.loads((ROOT / "photos.json").read_text(encoding="utf8"))
        print("\n新天線檢查 (料號｜datasheet 照片｜照片像素｜Datasheet 按鈕)：")
        problems = []
        for oc in diff.get("added", []):
            px = max(Image.open(io.BytesIO(base64.b64decode(ph[oc].split(",", 1)[1]))).size) if oc in ph else 0
            has_ds, has_btn = oc in ds, bool(items.get(oc, {}).get("ds"))
            print(f"  {oc:16} {'✓' if has_ds else '✗'}  {px:>4}px  {'✓' if has_btn else '✗'}")
            if not (has_ds and has_btn):
                problems.append(oc)
        print("全部完成" if not problems else f"需要處理：{problems}")


if __name__ == "__main__":
    main()
