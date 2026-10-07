#!/usr/bin/env python3
"""
export_guides.py — 將工具內建的使用說明匯出為 PDF (中 / 英 / 日 / 韓)

使用說明本身由 index.html 即時產生 (index.html?guide=1&lang=xx)，
每次以 build_data.py 更新工具後，使用說明即自動同步；此腳本只是把它存成 PDF 檔。

用法 (在專案根目錄執行，需先完成 build_data.py)：
    pip install playwright && python -m playwright install chromium
    python3 tools/export_guides.py                 # 輸出到 guides/
    python3 tools/export_guides.py --out ./pdf --lang en ja

一般使用者也可以直接在工具中按「使用說明 → 列印 / 存成 PDF」，結果相同。
"""
from __future__ import annotations

import argparse
import functools
import http.server
import json
import socketserver
import threading
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
LANGS = ["zh", "en", "ja", "ko"]
# 頁尾：左側為使用說明標題與版本 (取自頁面)，右側為頁碼
FOOT = ('<div style="width:100%;font-size:7.5px;color:#5B6B76;padding:0 13mm;display:flex;justify-content:space-between;'
        'font-family:sans-serif"><span>{title}</span><span><span class="pageNumber"></span> / <span class="totalPages"></span></span></div>')


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    """不輸出存取紀錄的靜態檔案伺服器。"""
    def log_message(self, *args):
        pass


def serve(root: Path):
    """以背景執行緒啟動本機 HTTP 伺服器，回傳 (server, port)。"""
    handler = functools.partial(QuietHandler, directory=str(root))
    srv = socketserver.TCPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_address[1]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", type=Path, default=ROOT / "guides")
    ap.add_argument("--lang", nargs="*", default=LANGS, choices=LANGS)
    a = ap.parse_args()
    version = json.loads((ROOT / "version.json").read_text(encoding="utf8"))["version"]
    a.out.mkdir(parents=True, exist_ok=True)
    srv, port = serve(ROOT)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            for lang in a.lang:
                page = browser.new_page()
                page.goto(f"http://127.0.0.1:{port}/index.html?guide=1&lang={lang}")
                # 等待所有 iPhone 畫面就緒
                page.wait_for_function("window.__guideReady === true", timeout=60000)
                page.wait_for_timeout(300)
                title = page.evaluate("document.querySelector('.g-foot')?.textContent || document.title")
                title = title.replace("&", "&amp;").replace("<", "&lt;")
                out = a.out / f"Quectel_Antenna_Selector_TELEC_User_Guide_{lang.upper()}_v{version}.pdf"
                page.pdf(path=str(out), format="A4", print_background=True, prefer_css_page_size=True,
                         display_header_footer=True, header_template="<div></div>", footer_template=FOOT.format(title=title))
                print(f"{lang}: {out}")
                page.close()
            browser.close()
    finally:
        srv.shutdown()


if __name__ == "__main__":
    main()
