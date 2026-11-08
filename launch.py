"""Start the app locally and open it in the default browser.

Used by the portable Windows build (started by ``BuyBox.exe``) but works anywhere:
    python launch.py

- picks a free port starting at 8642 (or BUYBOX_PORT),
- if the app is already running, just opens the browser (no second copy),
- opens the browser when the server is ready and, under BuyBox.exe, minimizes the console.
"""

from __future__ import annotations

import json
import os
import socket
import urllib.request
import sys
import threading
import time
import webbrowser
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
INSTANCE_FILE = APP_DIR / "data" / ".instance.json"
MARKER_URL = "http://127.0.0.1:{port}/app/static/buybox_id.txt"
MARKER = "buybox-analyzer"


#: Dedicated default port — deliberately not Streamlit's 8501, which other internal
#: Streamlit tools (e.g. DK ATP) use. Override with the BUYBOX_PORT environment variable.
DEFAULT_PORT = 8642


def port_is_free(port: int) -> bool:
    """True when nothing listens on ``port`` and it can be bound.

    Checks a connection first: on Windows a bind to 127.0.0.1 can succeed even while another
    program listens on 0.0.0.0 with the same port, which would make the two apps collide.
    """
    for host in ("127.0.0.1", "::1"):
        fam = socket.AF_INET6 if ":" in host else socket.AF_INET
        try:
            with socket.socket(fam, socket.SOCK_STREAM) as s:
                s.settimeout(0.3)
                if s.connect_ex((host, port)) == 0:
                    return False
        except OSError:
            pass
    for host in ("0.0.0.0", "127.0.0.1"):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):  # Windows
                s.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
            try:
                s.bind((host, port))
            except OSError:
                return False
    return True


def free_port(start: int = DEFAULT_PORT, tries: int = 50) -> int:
    """First free TCP port starting at ``start``."""
    for port in range(start, start + tries):
        if port_is_free(port):
            return port
    raise SystemExit("No free port found between %d and %d." % (start, start + tries))


def _local_get(url: str, timeout: float = 1.5) -> str:
    """GET a localhost URL, ignoring any system HTTP proxy (common on corporate PCs)."""
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(url, timeout=timeout) as r:
        return r.read().decode("utf-8", "replace")


def is_our_app(port: int) -> bool:
    """True when this Buy Box Analyzer (not some other program) answers on ``port``."""
    try:
        return _local_get(MARKER_URL.format(port=port)).strip() == MARKER
    except (OSError, ValueError):
        return False


def running_instance() -> int | None:
    """Port of an already running copy of the app, if any."""
    try:
        port = int(json.loads(INSTANCE_FILE.read_text(encoding="utf-8"))["port"])
    except (OSError, ValueError, KeyError, TypeError):
        return None
    return port if is_our_app(port) else None


def remember_instance(port: int) -> None:
    try:
        INSTANCE_FILE.parent.mkdir(parents=True, exist_ok=True)
        INSTANCE_FILE.write_text(json.dumps({"port": port, "pid": os.getpid()}), encoding="utf-8")
    except OSError:
        pass


def minimize_console() -> None:
    """Under BuyBox.exe on Windows: tuck the console window into the taskbar."""
    if os.name != "nt" or os.environ.get("BUYBOX_LAUNCHER") != "exe" or os.environ.get("BUYBOX_NO_MINIMIZE"):
        return
    try:
        import ctypes

        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE
    except Exception:  # noqa: BLE001 - cosmetic only
        pass


def open_browser(url: str) -> None:
    if os.environ.get("BUYBOX_NO_BROWSER") != "1":
        webbrowser.open(url)


def open_when_ready(port: int, timeout: float = 90.0) -> None:
    """Open the browser once the server accepts connections."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                break
        except OSError:
            time.sleep(0.4)
    url = f"http://localhost:{port}"
    print(f"\n  Buy Box Analyzer is running at {url}\n"
          "  To stop it: close this window, or use «بستن برنامه» in the app's sidebar.\n", flush=True)
    open_browser(url)
    time.sleep(1.5)
    minimize_console()


def main() -> int:
    os.chdir(APP_DIR)  # so .streamlit/config.toml and static/ are found
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    existing = running_instance()
    if existing is not None:
        url = f"http://localhost:{existing}"
        print(f"Buy Box Analyzer is already running at {url} — opening it.", flush=True)
        open_browser(url)
        return 0
    port = int(os.environ.get("BUYBOX_PORT") or free_port())
    remember_instance(port)
    os.environ.setdefault("BUYBOX_LAUNCHER", "script")  # lets the app show its «بستن برنامه» button
    threading.Thread(target=open_when_ready, args=(port,), daemon=True).start()
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit", "run", str(APP_DIR / "app.py"),
        f"--server.port={port}",
        "--server.address=localhost",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--server.maxUploadSize=500",
        "--server.maxMessageSize=500",
    ]
    return stcli.main()


if __name__ == "__main__":
    sys.exit(main())
