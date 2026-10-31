"""Start the app locally and open it in the default browser.

Used by the portable Windows build (``BuyBox.bat``) but works anywhere:
    python launch.py
"""

from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent


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
    print(f"\n  Buy Box Analyzer is running at {url}\n  Close this window to stop it.\n", flush=True)
    if os.environ.get("BUYBOX_NO_BROWSER") != "1":
        webbrowser.open(url)


def main() -> int:
    os.chdir(APP_DIR)  # so .streamlit/config.toml and static/ are found
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    port = int(os.environ.get("BUYBOX_PORT") or free_port())
    threading.Thread(target=open_when_ready, args=(port,), daemon=True).start()
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit", "run", str(APP_DIR / "app.py"),
        f"--server.port={port}",
        "--server.address=localhost",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ]
    return stcli.main()


if __name__ == "__main__":
    sys.exit(main())
