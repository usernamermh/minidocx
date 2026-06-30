from __future__ import annotations

import base64
import errno
import json
import os
import sys
import tempfile
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote

from docx_io import docx_bytes_to_document, document_to_docx_bytes
from resource_tools import empty_all_working_sets, get_resource_stats


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
HOST = "127.0.0.1"
PORT = 8765


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return ROOT


DEBUG_LOG_PATH = _runtime_root() / "debug_runtime.log"
SAFE_SAVE_DIR = Path.home() / "MiniDocxSafeSaves"
MAX_STAGED_SAVES_PER_FILE = 10


def _safe_filename(filename: str) -> str:
    raw = str(filename or "").strip() or "mini-docx.docx"
    cleaned = "".join(ch if ch not in '<>:"/\\|?*' and ord(ch) >= 32 else "_" for ch in raw).strip(" .")
    if not cleaned:
        cleaned = "mini-docx.docx"
    if not cleaned.lower().endswith(".docx"):
        cleaned += ".docx"
    return cleaned


def _prune_staged_saves(filename: str) -> None:
    stem = Path(filename).stem
    suffix = Path(filename).suffix or ".docx"
    pattern = f"{stem}-*{suffix}"
    candidates = []
    for path in SAFE_SAVE_DIR.glob(pattern):
        if not path.is_file():
            continue
        candidates.append(path)
    candidates.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    for stale in candidates[MAX_STAGED_SAVES_PER_FILE:]:
        try:
            stale.unlink(missing_ok=True)
        except Exception:
            pass


class EditorHandler(BaseHTTPRequestHandler):
    @staticmethod
    def _content_disposition(filename: str) -> str:
        safe_ascii = "".join(ch if ord(ch) < 128 and ch not in {'"', '\\'} else "_" for ch in filename) or "download.docx"
        encoded = quote(filename, safe="")
        return f'attachment; filename="{safe_ascii}"; filename*=UTF-8\'\'{encoded}'

    def do_GET(self) -> None:
        route = self.path.split("?", 1)[0]
        if route == "/":
            self._serve_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if route == "/api/resource-stats":
            self._resource_stats()
            return
        if route.startswith("/static/"):
            target = (STATIC_DIR / route.removeprefix("/static/")).resolve()
            if STATIC_DIR not in target.parents and target != STATIC_DIR:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            mime = {
                ".js": "application/javascript; charset=utf-8",
                ".css": "text/css; charset=utf-8",
                ".html": "text/html; charset=utf-8",
                ".svg": "image/svg+xml",
            }.get(target.suffix.lower(), "application/octet-stream")
            self._serve_file(target, mime)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        route = self.path.split("?", 1)[0]
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length else b""

        if route == "/api/stage-save":
            self._stage_save(raw)
            return

        try:
            payload = json.loads(raw.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            self._send_json({"error": "Invalid JSON payload."}, HTTPStatus.BAD_REQUEST)
            return

        if route == "/api/import-docx":
            self._import_docx(payload)
            return
        if route == "/api/export-docx":
            self._export_docx(payload)
            return
        if route == "/api/delete-staged-save":
            self._delete_staged_save(payload)
            return
        if route == "/api/debug-log":
            self._debug_log(payload)
            return
        if route == "/api/clean-resources":
            self._clean_resources()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def _import_docx(self, payload: dict) -> None:
        data_b64 = payload.get("data")
        if not data_b64:
            self._send_json({"error": "Missing DOCX data."}, HTTPStatus.BAD_REQUEST)
            return
        try:
            binary = base64.b64decode(data_b64)
            document = docx_bytes_to_document(binary)
        except Exception as exc:
            self._send_json({"error": f"Failed to open DOCX: {exc}"}, HTTPStatus.BAD_REQUEST)
            return
        self._send_json({"document": document})

    def _export_docx(self, payload: dict) -> None:
        document = payload.get("document") or {}
        filename = payload.get("filename") or "mini-docx.docx"
        if not filename.lower().endswith(".docx"):
            filename += ".docx"
        try:
            binary = document_to_docx_bytes(document)
        except Exception as exc:
            self._send_json({"error": f"Failed to save DOCX: {exc}"}, HTTPStatus.BAD_REQUEST)
            return
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        self.send_header("Content-Disposition", self._content_disposition(filename))
        self.send_header("Content-Length", str(len(binary)))
        self.end_headers()
        self.wfile.write(binary)

    def _stage_save(self, raw: bytes) -> None:
        filename = _safe_filename(self.headers.get("X-Filename", "mini-docx.docx"))
        SAFE_SAVE_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        prefix = f"{Path(filename).stem}-{stamp}-"
        suffix = Path(filename).suffix or ".docx"
        try:
            with tempfile.NamedTemporaryFile(prefix=prefix, suffix=suffix, dir=SAFE_SAVE_DIR, delete=False) as fh:
                fh.write(raw)
                staged_path = Path(fh.name)
        except Exception as exc:
            self._send_json(
                {"ok": False, "error": str(exc), "directory": str(SAFE_SAVE_DIR)},
                HTTPStatus.INTERNAL_SERVER_ERROR,
            )
            return
        _prune_staged_saves(filename)
        self._send_json({"ok": True, "path": str(staged_path), "directory": str(SAFE_SAVE_DIR)})

    def _delete_staged_save(self, payload: dict) -> None:
        raw_path = payload.get("path")
        if not raw_path:
            self._send_json({"ok": False, "error": "Missing staged save path."}, HTTPStatus.BAD_REQUEST)
            return
        try:
            target = Path(str(raw_path)).resolve()
            safe_root = SAFE_SAVE_DIR.resolve()
            if safe_root not in target.parents:
                self._send_json({"ok": False, "error": "Invalid staged save path."}, HTTPStatus.BAD_REQUEST)
                return
            target.unlink(missing_ok=True)
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        self._send_json({"ok": True})

    def _debug_log(self, payload: dict) -> None:
        try:
            event = str(payload.get("event") or "").strip() or "unknown"
            data = payload.get("data")
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            line = json.dumps({"ts": stamp, "event": event, "data": data}, ensure_ascii=False)
            with DEBUG_LOG_PATH.open("a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc), "path": str(DEBUG_LOG_PATH)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
        self._send_json({"ok": True, "path": str(DEBUG_LOG_PATH)})

    def _resource_stats(self) -> None:
        try:
            self._send_json(get_resource_stats())
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def _clean_resources(self) -> None:
        try:
            self._send_json(empty_all_working_sets())
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def _serve_file(self, path: Path, content_type: str) -> None:
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args: object) -> None:
        return


def create_server(preferred_port: int) -> tuple[ThreadingHTTPServer, int]:
    candidates = [preferred_port + offset for offset in range(10)]
    candidates.append(0)
    last_error: OSError | None = None

    for candidate in candidates:
        try:
            server = ThreadingHTTPServer((HOST, candidate), EditorHandler)
            actual_port = server.server_address[1]
            return server, actual_port
        except OSError as exc:
            last_error = exc
            if exc.errno not in {errno.EACCES, errno.EADDRINUSE}:
                raise

    if last_error is not None:
        raise last_error
    raise RuntimeError("Failed to create HTTP server.")


def main() -> None:
    preferred_port = int(os.environ.get("MINI_DOCX_PORT", PORT))
    try:
        DEBUG_LOG_PATH.write_text("", encoding="utf-8")
    except Exception:
        pass
    server, port = create_server(preferred_port)
    url = f"http://{HOST}:{port}"
    if port != preferred_port:
        print(f"Preferred port {preferred_port} is unavailable, switched to {port}.")
    print(f"Mini DOCX web editor is running at {url}")
    print("Press Ctrl+C to stop the server.")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
