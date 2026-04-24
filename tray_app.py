from __future__ import annotations

import os
import sys
import atexit
import ctypes
import json
import threading
import webbrowser
from ctypes import wintypes
from pathlib import Path

import pystray
from PIL import Image, ImageDraw, ImageFont

import server


ERROR_ALREADY_EXISTS = 183
MUTEX_NAME = "Local\\MiniDocxTraySingleInstance"
STATE_FILE_NAME = "mini_docx_tray_state.json"


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def runtime_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def state_path() -> Path:
    return runtime_root() / STATE_FILE_NAME


def acquire_single_instance_mutex():
    if os.name != "nt":
        return None
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(handle)
        return None
    return handle


def release_single_instance_mutex(handle) -> None:
    if os.name != "nt" or not handle:
        return
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.ReleaseMutex.argtypes = [wintypes.HANDLE]
    kernel32.ReleaseMutex.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    kernel32.ReleaseMutex(handle)
    kernel32.CloseHandle(handle)


def open_existing_instance() -> None:
    url = f"http://{server.HOST}:{server.PORT}"
    try:
        with state_path().open("r", encoding="utf-8") as fh:
            state = json.load(fh)
        port = int(state.get("port") or server.PORT)
        url = str(state.get("url") or f"http://{server.HOST}:{port}")
    except Exception:
        pass
    try:
        webbrowser.open(url)
    except Exception:
        pass


def load_icon_image() -> Image.Image:
    icon_path = resource_path("tray_icon.png")
    if icon_path.exists():
        return Image.open(icon_path)
    size = 64
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.ellipse((4, 4, size - 4, size - 4), fill=(155, 91, 52, 255))
    draw.ellipse((10, 10, size - 10, size - 10), fill=(255, 249, 242, 255))
    text = "DOC"
    try:
        font = ImageFont.truetype("arialbd.ttf", 18)
    except Exception:
        font = ImageFont.load_default()
    text_w, text_h = draw.textsize(text, font=font)
    draw.text(((size - text_w) / 2, (size - text_h) / 2), text, fill=(36, 48, 40, 255), font=font)
    return image


class TrayApp:
    def __init__(self) -> None:
        self.server = None
        self.thread: threading.Thread | None = None
        self.port = int(os.environ.get("MINI_DOCX_PORT", "8765"))
        self.common_ports = [8765, 8000, 9000, 10000, 0]
        self.requested_port = self.port
        self.icon = pystray.Icon(
            "MiniDocx",
            load_icon_image(),
            "Mini DOCX (stopped)",
            self._build_menu(),
        )
        atexit.register(self.stop_server)

    def write_state(self) -> None:
        payload = {
            "pid": os.getpid(),
            "host": server.HOST,
            "port": self.port,
            "url": f"http://{server.HOST}:{self.port}",
        }
        try:
            state_path().write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    def clear_state(self) -> None:
        try:
            state_path().unlink(missing_ok=True)
        except Exception:
            pass

    def _build_menu(self) -> pystray.Menu:
        return pystray.Menu(
            pystray.MenuItem("启动", self.start_action),
            pystray.MenuItem("停止", self.stop_action),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("打开网页", self.open_action),
            pystray.MenuItem(
                "选择端口",
                pystray.Menu(
                    *[
                        pystray.MenuItem(
                            f"{port}" if port else "随机端口",
                            self._make_port_action(port),
                            checked=self._make_port_checked(port),
                        )
                        for port in self.common_ports
                    ]
                ),
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("退出", self.exit_action),
        )

    def update_tooltip(self) -> None:
        if self.server:
            self.icon.title = f"Mini DOCX (运行中: {self.port})"
        else:
            self.icon.title = "Mini DOCX (stopped)"

    def _make_port_action(self, port: int):
        def action(icon: pystray.Icon, item: pystray.MenuItem) -> None:
            restart = self.server is not None
            self.port = port
            if restart:
                self.stop_server()
                self.start_server()
        return action

    def _make_port_checked(self, port: int):
        def checked(item: pystray.MenuItem) -> bool:
            return self.port == port
        return checked

    def start_server(self) -> None:
        if self.server:
            return
        try:
            self.requested_port = self.port
            self.server, actual_port = server.create_server(self.port)
        except Exception as exc:
            self.icon.title = f"启动失败: {exc}"
            return
        self.port = actual_port
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.write_state()
        self.update_tooltip()
        if self.requested_port != self.port and self.requested_port != 0:
            self.icon.title = f"端口被占用，已切换至 {self.port}"
        webbrowser.open(f"http://{server.HOST}:{self.port}")

    def stop_server(self) -> None:
        if not self.server:
            return
        self.server.shutdown()
        self.server.server_close()
        self.server = None
        if self.thread:
            self.thread.join(timeout=2)
        self.thread = None
        self.clear_state()
        self.update_tooltip()

    def start_action(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self.server:
            return
        self.start_server()

    def stop_action(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        self.stop_server()

    def open_action(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if not self.server:
            self.start_action(icon, item)
            return
        webbrowser.open(f"http://{server.HOST}:{self.port}")

    def exit_action(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        self.stop_server()
        self.icon.stop()

    def run(self) -> None:
        try:
            self.icon.run()
        finally:
            self.stop_server()


if __name__ == "__main__":
    mutex_handle = acquire_single_instance_mutex()
    if mutex_handle is None:
        open_existing_instance()
        sys.exit(0)
    try:
        TrayApp().run()
    finally:
        release_single_instance_mutex(mutex_handle)
