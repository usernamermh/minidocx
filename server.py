from __future__ import annotations

import base64
import errno
import json
import os
import socket
import sys
from datetime import datetime
from urllib.parse import quote
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from docx_io import docx_bytes_to_document, document_to_docx_bytes


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
HOST = "127.0.0.1"
PORT = 8765


def _runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return ROOT


DEBUG_LOG_PATH = _runtime_root() / "debug_runtime.log"


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
        raw = self.rfile.read(length) if length else b"{}"
        try:
            payload = json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError:
            self._send_json({"error": "Invalid JSON payload."}, HTTPStatus.BAD_REQUEST)
            return

        if route == "/api/import-docx":
            self._import_docx(payload)
            return
        if route == "/api/export-docx":
            self._export_docx(payload)
            return
        if route == "/api/debug-log":
            self._debug_log(payload)
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
        body = json.dumps(payload).encode("utf-8")
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
