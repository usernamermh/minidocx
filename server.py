from __future__ import annotations

import base64
import ctypes
import errno
import json
import os
import sys
import tempfile
import threading
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote
from ctypes import wintypes

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
FILE_DIALOG_LOCK = threading.Lock()

OFN_OVERWRITEPROMPT = 0x00000002
OFN_NOCHANGEDIR = 0x00000008
OFN_PATHMUSTEXIST = 0x00000800
OFN_FILEMUSTEXIST = 0x00001000
OFN_EXPLORER = 0x00080000


class OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hInstance", wintypes.HINSTANCE),
        ("lpstrFilter", wintypes.LPCWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter", wintypes.DWORD),
        ("nFilterIndex", wintypes.DWORD),
        ("lpstrFile", wintypes.LPWSTR),
        ("nMaxFile", wintypes.DWORD),
        ("lpstrFileTitle", wintypes.LPWSTR),
        ("nMaxFileTitle", wintypes.DWORD),
        ("lpstrInitialDir", wintypes.LPCWSTR),
        ("lpstrTitle", wintypes.LPCWSTR),
        ("Flags", wintypes.DWORD),
        ("nFileOffset", wintypes.WORD),
        ("nFileExtension", wintypes.WORD),
        ("lpstrDefExt", wintypes.LPCWSTR),
        ("lCustData", wintypes.LPARAM),
        ("lpfnHook", wintypes.LPVOID),
        ("lpTemplateName", wintypes.LPCWSTR),
        ("pvReserved", wintypes.LPVOID),
        ("dwReserved", wintypes.DWORD),
        ("FlagsEx", wintypes.DWORD),
    ]


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
    candidates = [path for path in SAFE_SAVE_DIR.glob(f"{stem}-*{suffix}") if path.is_file()]
    candidates.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    for stale in candidates[MAX_STAGED_SAVES_PER_FILE:]:
        try:
            stale.unlink(missing_ok=True)
        except Exception:
            pass


def _create_staged_save(filename: str, binary: bytes) -> Path:
    safe_name = _safe_filename(filename)
    SAFE_SAVE_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    prefix = f"{Path(safe_name).stem}-{stamp}-"
    suffix = Path(safe_name).suffix or ".docx"
    with tempfile.NamedTemporaryFile(prefix=prefix, suffix=suffix, dir=SAFE_SAVE_DIR, delete=False) as fh:
        fh.write(binary)
        staged_path = Path(fh.name)
    _prune_staged_saves(safe_name)
    return staged_path


def _show_windows_file_dialog(*, save: bool, title: str, suggested_name: str = "", initial_dir: str = "") -> str:
    if os.name != "nt":
        raise RuntimeError("Native file dialogs are only available on Windows.")
    with FILE_DIALOG_LOCK:
        buffer = ctypes.create_unicode_buffer(32768)
        if suggested_name:
            buffer.value = suggested_name
        filters = "Word 文档 (*.docx)\0*.docx\0所有文件 (*.*)\0*.*\0\0"
        owner = ctypes.windll.user32.GetForegroundWindow()
        dialog = OPENFILENAMEW()
        dialog.lStructSize = ctypes.sizeof(OPENFILENAMEW)
        dialog.hwndOwner = owner
        dialog.lpstrFilter = filters
        dialog.nFilterIndex = 1
        dialog.lpstrFile = ctypes.cast(buffer, wintypes.LPWSTR)
        dialog.nMaxFile = len(buffer)
        dialog.lpstrInitialDir = initial_dir or None
        dialog.lpstrTitle = title
        dialog.lpstrDefExt = "docx"
        dialog.Flags = OFN_EXPLORER | OFN_NOCHANGEDIR | OFN_PATHMUSTEXIST
        if save:
            dialog.Flags |= OFN_OVERWRITEPROMPT
            succeeded = ctypes.windll.comdlg32.GetSaveFileNameW(ctypes.byref(dialog))
        else:
            dialog.Flags |= OFN_FILEMUSTEXIST
            succeeded = ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(dialog))
        if succeeded:
            return buffer.value
        error_code = ctypes.windll.comdlg32.CommDlgExtendedError()
        if error_code:
            raise OSError(f"Windows file dialog failed with code {error_code}.")
        return ""


def _absolute_windows_path(raw_path: str) -> Path:
    path = Path(str(raw_path or "").strip()).expanduser().resolve()
    if path.suffix.lower() != ".docx":
        path = path.with_suffix(".docx")
    return path


def _atomic_write(path: Path, binary: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f".{path.stem}-", suffix=".tmp", dir=path.parent, delete=False) as fh:
            fh.write(binary)
            fh.flush()
            os.fsync(fh.fileno())
            temporary_path = Path(fh.name)
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


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
        if route == "/api/pick-open-docx":
            self._pick_open_docx()
            return
        if route == "/api/pick-save-path":
            self._pick_save_path(payload)
            return
        if route == "/api/save-docx-path":
            self._save_docx_path(payload)
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

    def _pick_open_docx(self) -> None:
        try:
            selected = _show_windows_file_dialog(save=False, title="打开 DOCX 文件")
            if not selected:
                self._send_json({"ok": False, "cancelled": True})
                return
            path = Path(selected).resolve()
            document = docx_bytes_to_document(path.read_bytes())
        except Exception as exc:
            self._send_json({"ok": False, "error": f"Failed to open DOCX: {exc}"}, HTTPStatus.BAD_REQUEST)
            return
        self._send_json({"ok": True, "path": str(path), "name": path.name, "document": document})

    def _pick_save_path(self, payload: dict) -> None:
        suggested_name = _safe_filename(payload.get("suggested_name") or "mini-docx.docx")
        current_path = str(payload.get("current_path") or "").strip()
        initial_dir = str(Path(current_path).expanduser().parent) if current_path else str(Path.home() / "Documents")
        try:
            selected = _show_windows_file_dialog(save=True, title="保存 DOCX 文件", suggested_name=suggested_name, initial_dir=initial_dir)
            if not selected:
                self._send_json({"ok": False, "cancelled": True})
                return
            path = _absolute_windows_path(selected)
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        self._send_json({"ok": True, "path": str(path), "name": path.name})

    def _save_docx_path(self, payload: dict) -> None:
        raw_path = payload.get("path")
        if not raw_path:
            self._send_json({"ok": False, "error": "Missing target path."}, HTTPStatus.BAD_REQUEST)
            return
        try:
            path = _absolute_windows_path(str(raw_path))
            binary = document_to_docx_bytes(payload.get("document") or {})
            staged_path = _create_staged_save(path.name, binary)
            _atomic_write(path, binary)
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        self._send_json({"ok": True, "path": str(path), "name": path.name, "backup_path": str(staged_path)})

    def _stage_save(self, raw: bytes) -> None:
        filename = _safe_filename(self.headers.get("X-Filename", "mini-docx.docx"))
        try:
            staged_path = _create_staged_save(filename, raw)
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc), "directory": str(SAFE_SAVE_DIR)}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return
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
