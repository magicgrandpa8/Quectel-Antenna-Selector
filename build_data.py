#!/usr/bin/env python3
"""
build_data.py — 將「Quectel Antenna Product and TELEC Certification」Excel
轉為 PWA 內嵌資料，並產生 index.html / version.json，同步 sw.js 快取版本。

用法:
    python3 build_data.py --version 3.0.0 \
        --xlsx Quectel_Antenna_Product_and_TELEC_Certification_V3_7_20260916_Rev26_Q3_for_CU.xlsx

資料來源 (僅使用可見工作表):
    "Antenna vs Cert" : 天線規格 + 模組 TELEC 認證矩陣 (☑)
    "Name"            : 模組清單、認證日期、類別 (類別欄為合併儲存格，向下填補)
    "Histories"       : 版次資訊 (取最新一筆的 Version 與 Updated Date)

刻意不使用:
    - veryHidden 的各模組原始認證工作表 (內部資料)
    - "Histories" 的更新人員姓名

輸出:
    index.html    內嵌資料、圖片、字型、版本號的 App
    version.json  供 App 偵測伺服器新版本
    sw.js         CACHE_VERSION 自動改為 antenna-selector-v<版本號>

相依套件: openpyxl, Pillow (含 WebP 支援)
"""
from __future__ import annotations

import argparse
import base64
import datetime as dt
import io
import json
import re
import sys
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image

THUMB_PX = 160          # 縮圖最長邊 (px)
THUMB_QUALITY = 72      # WebP 品質
MAIN_SHEET = "Antenna vs Cert"
MODULE_SHEET = "Name"
HISTORY_SHEET = "Histories"
CERT_MARKS = {"☑", "√", "✓", "✔", "V", "v", "Y", "YES"}
# 規格欄最後一欄；其後皆為模組認證矩陣
LAST_SPEC_COL = "Compatible with Japan Market"

EMPTY = {None, "", "None", "N/A", "n/a", "NA", "-"}


# --------------------------------------------------------------------------- #
# 通用工具
# --------------------------------------------------------------------------- #
def clean(v):
    """將儲存格值正規化為字串；空值回傳 None。"""
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace("\\n", "\n").replace("\u2011", "-").strip()
    s = re.sub(r"[ \t]+", " ", s)
    return None if s in EMPTY else s


def clean_url(v):
    """只接受 https 網址；其餘回傳 None。"""
    s = clean(v)
    if not s:
        return None
    s = s.split()[0]
    return s if s.lower().startswith("https://") else None


def fmt_date(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return clean(v)


_RANGE_RE = re.compile(
    r"(DC|\d+(?:\.\d+)?)\s*(GHz|MHz)?\s*(?:[–—\-~]|to)\s*(\d+(?:\.\d+)?)\s*(GHz|MHz)?",
    re.I,
)
_SINGLE_RE = re.compile(r"(\d{3,5}(?:\.\d+)?)\s*(GHz|MHz)?", re.I)


def parse_ranges(text: str | None) -> list[list[float]]:
    """
    從頻率描述文字解析出 [low, high] (MHz) 區間並合併重疊。
    假設：未標單位的數值視為 MHz；數值 < 10 且未標單位則視為 GHz。
    """
    if not text:
        return []
    out: list[list[float]] = []
    consumed = []
    for m in _RANGE_RE.finditer(text):
        lo_s, u1, hi_s, u2 = m.groups()
        unit = (u2 or u1 or "").lower()
        lo = 0.0 if lo_s.upper() == "DC" else float(lo_s)
        hi = float(hi_s)
        mul_hi = 1000 if unit == "ghz" or (not unit and hi < 10) else 1
        mul_lo = 1000 if (u1 or "").lower() == "ghz" or (not u1 and unit == "ghz") or (not unit and 0 < lo < 10) else 1
        lo, hi = lo * mul_lo, hi * mul_hi
        if hi > lo:
            out.append([round(lo, 2), round(hi, 2)])
            consumed.append(m.span())
    masked = list(text)
    for a, b in consumed:
        for i in range(a, b):
            masked[i] = " "
    for m in _SINGLE_RE.finditer("".join(masked)):
        f = float(m.group(1)) * (1000 if (m.group(2) or "").lower() == "ghz" else 1)
        if 100 <= f <= 100000:
            out.append([f - 1, f + 1])
    out.sort()
    merged: list[list[float]] = []
    for lo, hi in out:
        if merged and lo <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], hi)
        else:
            merged.append([lo, hi])
    return merged


def norm_dim(text: str | None):
    """
    尺寸正規化，回傳 (顯示字串, 最長邊 mm)。
    最長邊取括號前主要外形中的最大數值 (疊層 patch 取最大那層)。假設單位皆為 mm。
    """
    if not text:
        return None, None
    t = re.sub(r"\s*mm\b", "", text, flags=re.I)
    t = t.replace("Ф", "Φ").replace("Ø", "Φ").replace("φ", "Φ")
    t = re.sub(r"\s*[×xX*]\s*", " × ", t)
    t = re.sub(r"Φ\s*", "Φ ", t)
    t = re.sub(r"\s*\+\s*", " + ", t)
    t = re.sub(r"\s+", " ", t).strip(" ,;")
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", t.split("(")[0])]
    return t, (max(nums) if nums else None)


# --------------------------------------------------------------------------- #
# 圖片擷取：解析 drawing XML，以錨點列號對應圖片
# --------------------------------------------------------------------------- #
class SheetImages:
    def __init__(self, xlsx_path: Path):
        self.z = zipfile.ZipFile(xlsx_path)
        wb_xml = self.z.read("xl/workbook.xml").decode("utf8")
        rels = self._rels("xl/_rels/workbook.xml.rels")
        self.sheet_file = {}
        for name, rid in re.findall(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb_xml):
            self.sheet_file[name.replace("&amp;", "&")] = "xl/" + rels[rid].lstrip("/").replace("xl/", "")

    def _rels(self, path):
        try:
            xml = self.z.read(path).decode("utf8")
        except KeyError:
            return {}
        res = {}
        for tag in re.findall(r"<Relationship [^>]+>", xml):
            res[re.search(r'Id="([^"]+)"', tag).group(1)] = re.search(r'Target="([^"]+)"', tag).group(1)
        return res

    def row_images(self, sheet: str) -> dict[int, bytes]:
        """回傳 {0-based 列號: 圖片 bytes}，同列多張僅取第一張。"""
        sf = self.sheet_file.get(sheet)
        if not sf:
            return {}
        base = sf.rsplit("/", 1)
        drawing = next((t for t in self._rels(f"{base[0]}/_rels/{base[1]}.rels").values() if "drawings/" in t), None)
        if not drawing:
            return {}
        dpath = "xl/drawings/" + drawing.split("drawings/")[-1]
        dname = dpath.rsplit("/", 1)
        drels = self._rels(f"{dname[0]}/_rels/{dname[1]}.rels")
        xml = self.z.read(dpath).decode("utf8")
        result: dict[int, bytes] = {}
        for anchor in re.findall(r"<xdr:(?:twoCellAnchor|oneCellAnchor)\b.*?</xdr:(?:twoCellAnchor|oneCellAnchor)>", xml, re.S):
            row = re.search(r"<xdr:from>.*?<xdr:row>(\d+)</xdr:row>", anchor, re.S)
            emb = re.search(r'r:embed="([^"]+)"', anchor)
            if not row or not emb or emb.group(1) not in drels:
                continue
            r = int(row.group(1))
            if r in result:
                continue
            try:
                result[r] = self.z.read("xl/media/" + drels[emb.group(1)].split("media/")[-1])
            except KeyError:
                pass
        return result


def to_thumb(raw: bytes) -> str | None:
    """轉為 WebP 縮圖 data URI (裁白邊)；失敗回傳 None。"""
    try:
        im = Image.open(io.BytesIO(raw)).convert("RGBA")
        flat = Image.alpha_composite(Image.new("RGBA", im.size, (255, 255, 255, 255)), im).convert("RGB")
        bbox = Image.eval(flat, lambda p: 255 - p).point(lambda p: 255 if p > 12 else 0).getbbox()
        if bbox:
            flat = flat.crop(bbox)
        flat.thumbnail((THUMB_PX, THUMB_PX), Image.LANCZOS)
        buf = io.BytesIO()
        flat.save(buf, "WEBP", quality=THUMB_QUALITY, method=6)
        return "data:image/webp;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return None


# --------------------------------------------------------------------------- #
# 讀取工作表
# --------------------------------------------------------------------------- #
def read_modules(wb) -> list[dict]:
    """Name 工作表：模組名稱、認證日期、類別 (合併儲存格向下填補)。"""
    rows = list(wb[MODULE_SHEET].iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if r and "Module Name" in [clean(c) for c in r])
    hdr = [clean(c) for c in rows[hdr_i]]
    ci = {h: i for i, h in enumerate(hdr) if h}
    out, cat = [], None
    for r in rows[hdr_i + 1:]:
        name = clean(r[ci["Module Name"]]) if ci["Module Name"] < len(r) else None
        if not name:
            continue
        cat = clean(r[ci["Category"]]) or cat
        out.append({"name": name, "date": fmt_date(r[ci["Cert Date"]]), "cat": cat or "Other"})
    return out


def read_revision(wb) -> dict:
    """Histories：取最上方 (最新) 一筆的版次與日期；不讀取人員姓名。"""
    rows = list(wb[HISTORY_SHEET].iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if r and "Version" in [clean(c) for c in r])
    hdr = [clean(c) for c in rows[hdr_i]]
    ci = {h: i for i, h in enumerate(hdr) if h}
    for r in rows[hdr_i + 1:]:
        ver = clean(r[ci["Version"]])
        if ver:
            return {"rev": ver, "date": fmt_date(r[ci.get("Updated Date")]) if "Updated Date" in ci else None,
                    "remark": clean(r[ci["Remark"]]) if "Remark" in ci else None}
    return {}


def read_products(wb, imgs: dict[int, bytes], images: dict, module_names: list[str]):
    rows = list(wb[MAIN_SHEET].iter_rows(values_only=True))
    hdr = [clean(h) for h in rows[0]]
    idx = {h: i for i, h in enumerate(hdr) if h}
    first_mod = idx[LAST_SPEC_COL] + 1
    mod_cols = {i: hdr[i] for i in range(first_mod, len(hdr)) if hdr[i]}
    unknown = [m for m in mod_cols.values() if m not in module_names]
    if unknown:
        print(f"警告：矩陣中的模組不在 Name 工作表：{unknown}", file=sys.stderr)

    products = []
    # rows[1] 為各欄統計列，資料自 rows[2] 起
    for rnum, r in enumerate(rows[2:], start=2):
        get = lambda k: clean(r[idx[k]]) if k in idx and idx[k] < len(r) else None
        oc, name = get("OC"), get("Product Name")
        if not oc or not name:
            continue
        techs = []
        for n in range(1, 5):
            t = get(f"Technology{n}")
            if t:
                techs.append({"t": t, "f": get(f"Frequency Range{n}"), "q": get(f"Quantity{n}"),
                              "eff": get(f"Efficiency{n}"), "gain": get(f"Peak Gain{n}"),
                              "pat": get(f"Radiation Pattern{n}"), "pol": get(f"Polarization{n}")})
        freq_text = get("Frequency Range") or "\n".join(t["f"] or "" for t in techs)
        dim, dim_l = norm_dim(get("Dimensions(mm)"))
        certs = [m for i, m in mod_cols.items() if i < len(r) and (clean(r[i]) or "") in CERT_MARKS]
        p = {
            "oc": oc, "type": get("Type"), "brochure": get("Brochure Type"),
            "name": name, "desc": get("Detailed Description"),
            "ptype": get("Product Type"), "form": get("Form Factor"),
            "qty": get("Antenna Quantity"), "mount": get("Mounting Type"),
            "dim": dim, "dimL": dim_l, "atype": get("Antenna Type"),
            "freq": freq_text, "ranges": parse_ranges(freq_text),
            "techs": techs, "lna": get("LNA Gain(dB)"),
            "cables": [c for c in ({"len": get("Cable Length1"), "type": get("Cable Type1")},
                                   {"len": get("Cable Length2"), "type": get("Cable Type2")}) if c["len"] or c["type"]],
            "conn": [c for c in (get("Connector Type1"), get("Connector Type2")) if c],
            "evb": get("SMD EVB OC"), "evbDim": norm_dim(get("EVB Dimensions(mm)"))[0],
            "ip": get("IP Rating"), "ik": get("IK Rating"), "flame": get("Flame Rating"),
            "uv": get("UV Resistant"), "env": get("Environmental"), "temp": get("Operation Temperature"),
            "jp": (get(LAST_SPEC_COL) or "").upper() == "YES",
            "replace": get("To Replace"),
            "ds": clean_url(r[idx["Datasheet link"]]) if "Datasheet link" in idx else None,
            "certs": certs,
        }
        if rnum in imgs and oc not in images:
            t = to_thumb(imgs[rnum])
            if t:
                images[oc] = t
        products.append({k: v for k, v in p.items() if v not in (None, [], "")})
    return products, list(mod_cols.values())


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True, type=Path, help="Antenna Product and TELEC Certification Excel")
    ap.add_argument("--version", required=True, help="App 版本號，例如 3.0.0")
    ap.add_argument("--template", default=Path("template.html"), type=Path)
    ap.add_argument("--out", default=Path("index.html"), type=Path)
    ap.add_argument("--sw", default=Path("sw.js"), type=Path)
    a = ap.parse_args()
    if not re.fullmatch(r"\d+\.\d+\.\d+", a.version):
        sys.exit("--version 格式需為 X.Y.Z，例如 3.0.0")
    built = dt.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %z")

    wb = load_workbook(a.xlsx, read_only=True, data_only=True)
    for s in (MAIN_SHEET, MODULE_SHEET, HISTORY_SHEET):
        if s not in wb.sheetnames:
            sys.exit(f"Excel 缺少工作表：{s}")

    modules = read_modules(wb)
    images: dict[str, str] = {}
    products, matrix_mods = read_products(
        wb, SheetImages(a.xlsx).row_images(MAIN_SHEET), images, [m["name"] for m in modules])

    # 模組清單以 Name 工作表順序為準，附上認證天線數
    for m in modules:
        m["n"] = sum(1 for p in products if m["name"] in p.get("certs", []))
    for name in matrix_mods:
        if name not in {m["name"] for m in modules}:
            modules.append({"name": name, "date": None, "cat": "Other",
                            "n": sum(1 for p in products if name in p.get("certs", []))})

    data = {
        "version": a.version, "built": built, "source": a.xlsx.name,
        "revision": read_revision(wb),
        "items": products, "modules": modules, "images": images,
    }
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

    # ---- 模板：資料 + 內嵌字型 ----
    tpl = a.template.read_text(encoding="utf8")
    if "/*__DATA__*/null" not in tpl:
        sys.exit("template.html 缺少 /*__DATA__*/null 佔位符")
    font_css = []
    for f in sorted((a.template.parent / "fonts").glob("ibm-plex-sans-latin-*-normal.woff2")):
        weight = re.search(r"-(\d{3})-normal", f.name).group(1)
        font_css.append(
            "@font-face{font-family:'IBM Plex Sans';font-style:normal;font-display:swap;"
            f"font-weight:{weight};src:url(data:font/woff2;base64,{base64.b64encode(f.read_bytes()).decode()}) format('woff2');"
            "unicode-range:U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,"
            "U+2000-206F,U+2074,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD;}")
    if not font_css:
        print("警告：找不到 fonts/*.woff2，將使用系統字型", file=sys.stderr)
    html = tpl.replace("/*__FONTS__*/", "\n".join(font_css))
    html = html.replace("/*__DATA__*/null", payload.replace("</", "<\\/"))
    a.out.write_text(html, encoding="utf8")

    # ---- version.json / sw.js ----
    a.out.with_name("version.json").write_text(
        json.dumps({"version": a.version, "built": built}, ensure_ascii=False), encoding="utf8")
    if a.sw.exists():
        sw, n = re.subn(r"const CACHE_VERSION = '[^']*';",
                        f"const CACHE_VERSION = 'antenna-selector-v{a.version}';", a.sw.read_text(encoding="utf8"))
        if n != 1:
            sys.exit("sw.js 找不到 CACHE_VERSION 宣告")
        a.sw.write_text(sw, encoding="utf8")

    certified = sum(1 for p in products if p.get("certs"))
    print(f"版本: v{a.version}  建置時間: {built}  Excel 版次: {data['revision'].get('rev')}")
    print(f"天線: {len(products)}  (有 TELEC 認證: {certified})  模組: {len(modules)}  "
          f"認證組合 (☑): {sum(len(p.get('certs', [])) for p in products)}  圖片: {len(images)}")
    print(f"Datasheet 連結: {sum(1 for p in products if p.get('ds'))}/{len(products)}  "
          f"尺寸可解析: {sum(1 for p in products if p.get('dimL'))}/{len(products)}")
    print(f"輸出: {a.out} ({a.out.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
