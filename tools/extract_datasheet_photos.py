#!/usr/bin/env python3
"""
extract_datasheet_photos.py — 從天線 datasheet (PDF) 擷取產品照片，供工具使用

用法 (在專案根目錄執行，需先完成一次 build_data.py)：
    pip install pymupdf pillow
    python3 tools/extract_datasheet_photos.py <PDF 資料夾或 .zip> [--only-blurry]

流程：
    1. 依檔名比對料號 (例如 Quectel_YC0001CA_Datasheet_V1.2.pdf)；檔名沒有時改讀第 1 頁文字
    2. 在前 3 頁中找出產品照片：略過太小的圖 (logo)、在多頁重複出現的圖 (頁首/頁尾)、
       幾乎滿版的背景圖，再取面積最大的一張 (優先第 1 頁)
    3. 透明背景一律補成白底，裁掉白邊
    4. 結果寫入 photo_datasheet.json ({料號: WebP data URI，最長邊 800px})，build_data.py 會自動使用
       data/photo_datasheet_exclude.csv 中的料號一律略過 (例如封面是評估板而非天線本體)
       (優先順序：photo_overrides/ 手動照片 > datasheet 照片 > --photo-xlsx > TELEC Excel)
    5. 產生 photo_datasheet_report.csv 與 photo_datasheet_review.png (新舊照片對照) 供人工確認

選項：
    --only-blurry   只處理目前照片最長邊 < 200px 的料號
    --min-px N      擷取到的照片最長邊至少 N px 才採用 (預設 300；未達但比目前照片清晰時仍採用)；
                    datasheet 為官方照片，達此解析度即採用 (比目前照片小時會在報告中標註)
"""
from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

import pymupdf
from PIL import Image, ImageChops, ImageDraw, ImageStat

ROOT = Path(__file__).resolve().parent.parent
BLURRY_PX = 200
MAX_PX = 800             # 存檔時最長邊上限 (App 顯示不超過此尺寸)
EXCLUDE = ROOT / "data" / "photo_datasheet_exclude.csv"   # 不採用 datasheet 照片的料號


def load_app_data():
    """讀取 index.html 內嵌的料號清單與目前照片大小。"""
    html = (ROOT / "index.html").read_text(encoding="utf8")
    i = html.index("const DATA = ") + len("const DATA = ")
    data = json.loads(html[i:html.index("\n", i)].rstrip(";"))
    sizes = {}
    photos_path = ROOT / "photos.json"
    if photos_path.exists():
        for oc, uri in json.loads(photos_path.read_text(encoding="utf8")).items():
            sizes[oc] = max(Image.open(io.BytesIO(base64.b64decode(uri.split(",", 1)[1]))).size)
    return [p["oc"] for p in data["items"]], sizes


def find_oc(pdf_path: Path, doc, ocs_sorted: list[str]) -> str | None:
    """先比對檔名，再比對第 1 頁文字；取最長的符合料號 (避免 YC0002AA 誤配成 YC0002A)。"""
    name = pdf_path.stem.upper()
    for oc in ocs_sorted:
        if re.search(rf"(?<![A-Z0-9]){re.escape(oc)}(?![A-Z0-9])", name):
            return oc
    text = doc[0].get_text().upper() if len(doc) else ""
    for oc in ocs_sorted:
        if re.search(rf"(?<![A-Z0-9]){re.escape(oc)}(?![A-Z0-9])", text):
            return oc
    return None


def to_white_rgb(pix: pymupdf.Pixmap) -> Image.Image:
    """Pixmap (可能含透明或 CMYK) → 白底 RGB。"""
    if pix.n - pix.alpha >= 4:                     # CMYK → RGB
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    mode = "RGBA" if pix.alpha else "RGB"
    im = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    if mode == "RGBA":
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im)
    return im.convert("RGB")


def crop_white(im: Image.Image) -> Image.Image:
    bbox = ImageChops.difference(im, Image.new("RGB", im.size, (255, 255, 255))).convert("L") \
        .point(lambda v: 255 if v > 14 else 0).getbbox()
    return im.crop(bbox) if bbox else im


def extract_photo(doc) -> tuple[Image.Image, int] | None:
    """回傳 (產品照片, 頁碼)；找不到時回傳 None。"""
    pages = range(min(3, len(doc)))
    seen: dict[int, int] = {}                      # xref → 出現頁數
    for pno in range(len(doc)):
        for img in doc[pno].get_images(full=True):
            seen[img[0]] = seen.get(img[0], 0) + 1
    best = None
    for pno in pages:
        page = doc[pno]
        page_area = page.rect.width * page.rect.height
        for img in page.get_images(full=True):
            xref, smask = img[0], img[1]
            if seen.get(xref, 0) >= max(3, len(doc) // 2):
                continue                           # 頁首/頁尾等在大部分頁面重複出現的圖 (產品照可能在後面頁面再出現一次，不算)
            try:
                rects = page.get_image_rects(xref)
                if rects and max(r.width * r.height for r in rects) > 0.85 * page_area:
                    continue                       # 滿版背景
                pix = pymupdf.Pixmap(doc, xref)
                if smask:
                    # 套用透明遮罩 (smask)：先去掉可能已存在但未套用遮罩的 alpha，再與遮罩合成
                    base = pymupdf.Pixmap(pix, 0) if pix.alpha else pix
                    if base.n - base.alpha >= 4:
                        base = pymupdf.Pixmap(pymupdf.csRGB, base)
                    pix = pymupdf.Pixmap(base, pymupdf.Pixmap(doc, smask))
                im = crop_white(to_white_rgb(pix))
            except Exception:
                continue
            w, h = im.size
            if min(w, h) < 20 or max(w, h) < 150 or max(w, h) / max(1, min(w, h)) > 16:
                continue                           # logo、細線 (細長的鞭狀天線仍保留，單色色條由下一條規則排除)
            if max(ImageStat.Stat(im.convert("L")).stddev) < 12:
                continue                           # 幾乎單一顏色的裝飾圖 (例如側邊色條)
            score = w * h * (1.25 if pno == 0 else 1.0)
            if not best or score > best[0]:
                best = (score, im, pno + 1)
    return (best[1], best[2]) if best else None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("source", type=Path, help="datasheet PDF 資料夾或 .zip")
    ap.add_argument("--only-blurry", action="store_true")
    ap.add_argument("--min-px", type=int, default=300)
    a = ap.parse_args()

    ocs, cur_sizes = load_app_data()
    excluded = set()
    if EXCLUDE.exists():
        excluded = {r["PN"].strip() for r in csv.DictReader(open(EXCLUDE, encoding="utf-8-sig")) if r.get("PN")}
    ocs_sorted = sorted(ocs, key=len, reverse=True)
    tmp = None
    src = a.source
    if src.suffix.lower() == ".zip":
        tmp = tempfile.TemporaryDirectory()
        zipfile.ZipFile(src).extractall(tmp.name)
        src = Path(tmp.name)
    pdfs = sorted(p for p in src.rglob("*") if p.suffix.lower() == ".pdf")

    out_path = ROOT / "photo_datasheet.json"
    bundle = json.loads(out_path.read_text(encoding="utf8")) if out_path.exists() else {}
    report, review = [], []
    for pdf in pdfs:
        try:
            doc = pymupdf.open(pdf)
        except Exception as e:
            report.append([pdf.name, "", "無法開啟", str(e)[:80], "", ""]); continue
        oc = find_oc(pdf, doc, ocs_sorted)
        if not oc:
            report.append([pdf.name, "", "找不到料號", "", "", ""]); continue
        if oc in excluded:
            report.append([pdf.name, oc, "略過 (排除清單)", "", cur_sizes.get(oc, 0), ""]); continue
        cur = cur_sizes.get(oc, 0)
        if a.only_blurry and cur >= BLURRY_PX:
            report.append([pdf.name, oc, "略過 (目前照片已清晰)", "", cur, ""]); continue
        res = extract_photo(doc)
        if not res:
            report.append([pdf.name, oc, "找不到產品照片", "", cur, ""]); continue
        im, page = res
        size = max(im.size)
        if size < a.min_px and size <= cur:
            report.append([pdf.name, oc, f"未採用 (小於 {a.min_px}px 且不比目前清晰)", page, cur, size]); continue
        # datasheet 為官方照片：達最低解析度即採用；比目前照片小時在報告中標註，供人工確認
        status = "採用" if size >= cur else "採用 (比目前照片小，請確認)"
        im.thumbnail((MAX_PX, MAX_PX), Image.LANCZOS)
        buf = io.BytesIO(); im.save(buf, "WEBP", quality=90, method=6)          # WebP：約為 PNG 的 1/6
        bundle[oc] = "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode()
        report.append([pdf.name, oc, status, page, cur, size])
        review.append((oc, cur, size, im))

    out_path.write_text(json.dumps(bundle, separators=(",", ":")), encoding="utf8")
    with open(ROOT / "photo_datasheet_report.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["PDF 檔名", "料號", "結果", "頁碼", "目前照片 (px)", "datasheet 照片 (px)"])
        w.writerows(report)
    if review:                                      # 對照圖：每格顯示新照片，標註新舊尺寸
        cols, cw, ch = 4, 300, 260
        sheet = Image.new("RGB", (cols * cw, ((len(review) + cols - 1) // cols) * ch), "white")
        d = ImageDraw.Draw(sheet)
        for k, (oc, cur, size, im) in enumerate(review):
            x, y = (k % cols) * cw, (k // cols) * ch
            t = im.copy(); t.thumbnail((cw - 20, ch - 40))
            sheet.paste(t, (x + (cw - t.width) // 2, y + 30 + (ch - 40 - t.height) // 2))
            d.text((x + 8, y + 8), f"{oc}  {cur}px -> {size}px", fill="black")
        sheet.save(ROOT / "photo_datasheet_review.png")
    ok = sum(1 for r in report if str(r[2]).startswith("採用"))
    print(f"PDF {len(pdfs)} 份；採用 {ok} 張；photo_datasheet.json 共 {len(bundle)} 款")
    print("請檢查 photo_datasheet_review.png 與 photo_datasheet_report.csv，再執行 build_data.py")
    if tmp:
        tmp.cleanup()


if __name__ == "__main__":
    main()
