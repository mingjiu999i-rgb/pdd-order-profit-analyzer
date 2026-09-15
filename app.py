from __future__ import annotations

import json
import io
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import webbrowser
import zipfile
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from pdd_analyzer.db import connect
from pdd_analyzer.engine import audit_summary, combined_summaries, daily_summary, day_detail
from pdd_analyzer.exporter import build_xlsx
from pdd_analyzer.importer import Importer, ShopRequiredError, sha256

BASE = Path(__file__).resolve().parent
WORKSPACE = Path(os.environ.get("PDD_WORKSPACE", BASE)).resolve()
DB_PATH = Path(os.environ.get("PDD_DB_PATH", WORKSPACE / ".data" / "pdd_analysis.db"))
INBOX = WORKSPACE / "导入报表"
EXPORT_DIR = WORKSPACE / "导出结果"
ARCHIVE_DIR = DB_PATH.parent / "导入归档"
PENDING_PATH = DB_PATH.parent / "pending_imports.json"
HOST, PORT = "127.0.0.1", int(os.environ.get("PDD_PORT", "8765"))
APP_VERSION = "2026.09.15.5"


def open_page(url):
    """Open a visible Chrome tab on macOS; fall back to the platform browser."""
    try:
        if sys.platform == "darwin":
            script = '''
on run argv
  set targetURL to item 1 of argv
  tell application "Google Chrome"
    activate
    if (count of windows) = 0 then
      make new window
      set URL of active tab of front window to targetURL
    else
      tell front window
        make new tab at end of tabs with properties {URL:targetURL}
        set active tab index to (count of tabs)
        set minimized to false
      end tell
    end if
  end tell
  tell application "System Events" to set frontmost of process "Google Chrome" to true
end run
'''
            subprocess.run(["osascript", "-e", script, "--", url], check=True)
            return
    except Exception:
        pass
    try: webbrowser.open(url)
    except Exception: pass


def safe_part(value):
    return re.sub(r'[\\/:*?"<>|]', "_", str(value or "待确认"))


def archive_import(path: Path, shop_name: str):
    """Move a processed inbox file out of sight while preserving the original."""
    folder = ARCHIVE_DIR / datetime.now().strftime("%Y-%m-%d") / safe_part(shop_name)
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / path.name
    if target.exists():
        target = folder / f"{path.stem}_{sha256(path.read_bytes())[:8]}{path.suffix}"
    shutil.move(str(path), str(target))


def read_pending():
    try: return json.loads(PENDING_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError): return []


def retry_pending_files(exclude_name=""):
    """After one shop is confirmed, retry the remaining files against new order history."""
    remaining = []
    with connect(DB_PATH) as conn:
        importer = Importer(conn)
        for item in read_pending():
            name = item.get("file_name", "")
            if not name or name == exclude_name:
                continue
            source = INBOX / Path(name).name
            if not source.is_file():
                continue
            try:
                results = importer.import_path(source)
                shops = sorted({r.get("shop_name") for r in results if r.get("shop_name")})
                archive_import(source, shops[0] if len(shops) == 1 else "多个店铺")
            except ShopRequiredError as exc:
                remaining.append({"file_name": source.name, "report_type": exc.report_type, "reason": str(exc)})
            except Exception as exc:
                remaining.append({"file_name": source.name, "report_type": "error", "reason": str(exc)})
    PENDING_PATH.write_text(json.dumps(remaining, ensure_ascii=False, indent=2), encoding="utf-8")


def auto_import_first_run():
    """Import inbox reports, prove their shop, then archive successful originals."""
    with connect(DB_PATH) as conn:
        source_dir = INBOX if INBOX.is_dir() else WORKSPACE
        candidates = [p for p in source_dir.iterdir() if p.is_file() and p.suffix.lower() in {".csv", ".xlsx", ".zip"}]
        known_hashes = {r[0] for r in conn.execute("SELECT file_sha256 FROM import_batches")}
        manifest_path = DB_PATH.parent / "auto_imported.json"
        try: auto_hashes = set(json.loads(manifest_path.read_text(encoding="utf-8")))
        except (FileNotFoundError, ValueError): auto_hashes = set()
        def priority(path):
            name = path.name
            if "订单" in name: return 0
            if "退款" in name or "售后" in name: return 1
            if "资金" in name or path.suffix.lower() == ".zip": return 2
            return 3
        importer = Importer(conn)
        pending = []
        for path in sorted(candidates, key=priority):
            digest = sha256(path.read_bytes())
            if digest in known_hashes | auto_hashes:
                row = conn.execute("SELECT shop_name FROM import_batches WHERE file_sha256=? ORDER BY id DESC LIMIT 1", (digest,)).fetchone()
                archive_import(path, row["shop_name"] if row else "已导入")
                continue
            try:
                results = importer.import_path(path)
                auto_hashes.add(digest)
                manifest_path.write_text(json.dumps(sorted(auto_hashes)), encoding="utf-8")
                for result in results:
                    print(f"自动导入：{result['file_name']}（{result.get('shop_name') or '未识别店铺'}）")
                shops = sorted({r.get("shop_name") for r in results if r.get("shop_name")})
                archive_import(path, shops[0] if len(shops) == 1 else "多个店铺")
            except ShopRequiredError as exc:
                pending.append({"file_name": path.name, "report_type": exc.report_type, "reason": str(exc)})
            except Exception as exc:
                print(f"跳过无法识别的文件 {path.name}：{exc}")
                pending.append({"file_name": path.name, "report_type": "error", "reason": str(exc)})
        PENDING_PATH.parent.mkdir(parents=True, exist_ok=True)
        PENDING_PATH.write_text(json.dumps(pending, ensure_ascii=False, indent=2), encoding="utf-8")


class Handler(BaseHTTPRequestHandler):
    def send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers(); self.wfile.write(data)

    def send_file(self, data: bytes, name: str, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
        from urllib.parse import quote
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(name)}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers(); self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        query = parse_qs(parsed.query)
        start_date = query.get("start", [""])[0]
        end_date = query.get("end", [""])[0]
        shop_filter = query.get("shop", [""])[0]
        for value in (start_date, end_date):
            if value and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
                self.send_json({"error": "日期格式不正确"}, 400); return
        if start_date and end_date and start_date > end_date:
            self.send_json({"error": "开始日期不能晚于截止日期"}, 400); return
        if path == "/api/summary":
            with connect(DB_PATH) as conn:
                days = daily_summary(conn, start_date or None, end_date or None, shop_filter or None)
                shops = [r["shop_name"] for r in conn.execute("SELECT DISTINCT shop_name FROM orders WHERE shop_name IS NOT NULL AND shop_name<>'' ORDER BY shop_name")]
                bounds = conn.execute("SELECT MIN(pay_date) first_date,MAX(pay_date) last_date FROM orders WHERE pay_date IS NOT NULL").fetchone()
                self.send_json({"days": days, "totals": combined_summaries(days), "audit": audit_summary(conn, shop_filter or None), "pending": read_pending(), "shops": shops, "bounds": dict(bounds)})
            return
        if path == "/api/version":
            self.send_json({"version": APP_VERSION})
            return
        if path == "/api/export.xlsx":
            with connect(DB_PATH) as conn:
                if shop_filter:
                    shops = [shop_filter]
                else:
                    shops = sorted({d["shop_name"] for d in daily_summary(conn, start_date or None, end_date or None)})
                stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                EXPORT_DIR.mkdir(parents=True, exist_ok=True)
                files = []
                for shop in shops:
                    safe_shop = re.sub(r'[\\/:*?"<>|]', "_", shop)
                    range_name = f"_{start_date or '最早'}至{end_date or '最新'}" if start_date or end_date else ""
                    name = f"{safe_shop}_订单经营分析{range_name}_{stamp}.xlsx"
                    data = build_xlsx(conn, shop, start_date or None, end_date or None)
                    (EXPORT_DIR / name).write_bytes(data)
                    files.append((name, data))
            if len(files) == 1:
                self.send_file(files[0][1], files[0][0])
            elif files:
                stream = io.BytesIO()
                with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
                    for name, data in files: archive.writestr(name, data)
                self.send_file(stream.getvalue(), f"各店铺订单经营分析_{stamp}.zip", "application/zip")
            else:
                with connect(DB_PATH) as conn:
                    self.send_file(build_xlsx(conn, None, start_date or None, end_date or None), "拼多多店铺_订单经营分析.xlsx")
            return
        if path.startswith("/api/day/"):
            shop = parse_qs(urlparse(self.path).query).get("shop", [None])[0]
            with connect(DB_PATH) as conn: self.send_json(day_detail(conn, path.rsplit("/", 1)[-1], shop))
            return
        target = BASE / "static" / ("index.html" if path == "/" else path.lstrip("/"))
        if not target.is_file() or BASE / "static" not in target.parents:
            self.send_error(404); return
        data = target.read_bytes(); self.send_response(200)
        self.send_header("Content-Type", mimetypes.guess_type(target.name)[0] or "application/octet-stream")
        self.send_header("Content-Length", str(len(data))); self.end_headers(); self.wfile.write(data)

    def do_POST(self):
        request_path = urlparse(self.path).path
        if request_path == "/api/assign-shop":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length) or b"{}")
                name, shop = Path(str(payload.get("file_name", ""))).name, str(payload.get("shop_name", "")).strip()
                source = INBOX / name
                if not shop: raise ValueError("请填写店铺名称")
                if not source.is_file() or source.parent.resolve() != INBOX.resolve(): raise ValueError("待处理文件不存在")
                with connect(DB_PATH) as conn: result = Importer(conn).import_path(source, shop)
                archive_import(source, shop)
                retry_pending_files(name)
                self.send_json({"ok": True, "results": result})
            except Exception as exc: self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        if request_path != "/api/import": self.send_error(404); return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length > 300 * 1024 * 1024: raise ValueError("单次文件不能超过 300MB")
            name = unquote(self.headers.get("X-Filename", "upload.csv"))
            shop = unquote(self.headers.get("X-Shop-Name", "")).strip() or None
            data = self.rfile.read(length)
            with connect(DB_PATH) as conn:
                result = Importer(conn).import_bytes(name, data, shop) if Path(name).suffix.lower() != ".zip" else self._import_zip(conn, name, data, shop)
            self.send_json({"ok": True, "results": result})
        except ShopRequiredError as exc:
            self.send_json({"ok": False, "needs_shop": True, "file_name": exc.file_name, "report_type": exc.report_type, "error": str(exc)}, 409)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)

    def _import_zip(self, conn, name, data, shop=None):
        path = Path(tempfile.mkdtemp()) / name
        path.write_bytes(data)
        try: return Importer(conn).import_path(path, shop)
        finally: shutil.rmtree(path.parent, ignore_errors=True)

    def log_message(self, fmt, *args):
        print("[本地服务]", fmt % args)


if __name__ == "__main__":
    INBOX.mkdir(parents=True, exist_ok=True)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    auto_import_first_run()
    url = f"http://{HOST}:{PORT}"
    try:
        server = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as exc:
        # A previous double-click may have left this client's local service running.
        # In that case the import above has still completed; reopen the existing page
        # instead of making the launcher appear to do nothing.
        if getattr(exc, "errno", None) in {48, 98, 10048}:
            print(f"分析程序已在运行，正在打开：{url}")
            open_page(url)
            raise SystemExit(0)
        raise
    print(f"拼多多订单经营分析已启动：{url}")
    print("关闭此窗口即可停止程序。")
    open_page(url)
    server.serve_forever()
