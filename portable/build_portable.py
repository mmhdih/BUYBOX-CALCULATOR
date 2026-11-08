"""Build the portable Windows (x64) package of the Buy Box Analyzer.

The result is a folder (and a .zip of it) that runs on any 64-bit Windows 10/11 PC with
no installation and no admin rights: it contains the official *embeddable* Python, all
dependencies, the app and a native ``BuyBox.exe`` launcher (with icon).

Runs on Windows, Linux or macOS (dependencies are fetched as win_amd64 wheels):

    python portable/build_portable.py                 # -> dist/BuyBoxAnalyzer + dist/BuyBoxAnalyzer_portable_win64.zip
    python portable/build_portable.py --embed-zip python-3.12.10-embed-amd64.zip   # offline Python

Needs internet access to python.org (unless --embed-zip is given) and to PyPI.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP_NAME = "BuyBoxAnalyzer"
DEFAULT_PYTHON = "3.12.10"
EMBED_URL = "https://www.python.org/ftp/python/{v}/python-{v}-embed-amd64.zip"

#: Files/folders of the repository copied into <package>/app
APP_ITEMS = ("app.py", "launch.py", "buybox", "static", ".streamlit", "README.md")

README_PORTABLE = """تحلیل بای باکس شمش و سکه — نسخه پرتابل ویندوز
==================================================

اجرا:
  1. کل فایل zip را در یک پوشه Extract کنید (مثلاً روی دسکتاپ یا فلش).
  2. روی BuyBox.exe دوبار کلیک کنید. مرورگر خودکار باز می‌شود و پنجره سیاه کوچک می‌شود (Minimize).
  3. برای بستن برنامه: دکمه «بستن برنامه» در نوار کناری اپ، یا بستن پنجره Buy Box Analyzer در نوار وظیفه.
  اگر برنامه از قبل باز باشد، اجرای دوباره BuyBox.exe فقط همان را در مرورگر باز می‌کند.

نکات:
  - نیازی به نصب پایتون یا دسترسی ادمین نیست.
  - پورت پیش‌فرض 8642 است (نه 8501) تا با اپ‌های دیگر مثل DK ATP تداخل نداشته باشد؛
    اگر اشغال باشد، پورت آزاد بعدی انتخاب می‌شود.
  - داده‌ها فقط روی همین کامپیوتر پردازش می‌شوند و به اینترنت ارسال نمی‌شوند.
  - فایل‌های xlsx و csv تا ۵۰۰ مگابایت پشتیبانی می‌شوند.
  - تاریخچه اجراها، تنظیمات و آخرین لیست KAM در پوشه app\\data ذخیره می‌شوند.
    برای انتقال به نسخه جدید، همین پوشه data را کپی کنید.
  - BuyBox.exe امضای دیجیتال ندارد؛ اگر ویندوز هشدار SmartScreen داد: More info ← Run anyway.

Portable Windows build. Double-click BuyBox.exe (port 8642). Data stays on this PC (app\\data).
"""


def log(msg: str) -> None:
    print(f"[build] {msg}", flush=True)


def download(url: str, dest: Path, attempts: int = 4) -> None:
    for i in range(attempts):
        try:
            log(f"downloading {url}")
            with urllib.request.urlopen(url, timeout=120) as r, open(dest, "wb") as f:
                shutil.copyfileobj(r, f)
            return
        except OSError as exc:
            if i == attempts - 1:
                raise
            wait = 2 ** (i + 1)
            log(f"download failed ({exc}); retrying in {wait}s")
            time.sleep(wait)


def patch_pth(py_dir: Path) -> None:
    """Enable site-packages in the embeddable distribution's ._pth file."""
    pths = list(py_dir.glob("python3*._pth"))
    if not pths:
        raise SystemExit("No python3*._pth file found in the embeddable Python.")
    pth = pths[0]
    zip_name = next((ln.strip() for ln in pth.read_text().splitlines() if ln.strip().endswith(".zip")),
                    f"{pth.stem}.zip")
    pth.write_text("\n".join([zip_name, ".", "Lib\\site-packages", "..\\app", "import site", ""]))
    log(f"patched {pth.name}")


def pip_install(site: Path, requirements: Path, py_ver: str) -> None:
    short = ".".join(py_ver.split(".")[:2])
    cmd = [
        sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--no-compile",
        "--target", str(site), "--platform", "win_amd64", "--python-version", short,
        "--implementation", "cp", "--only-binary=:all:", "--upgrade", "-r", str(requirements),
    ]
    log("installing dependencies (win_amd64 wheels)")
    subprocess.run(cmd, check=True)


def slim(site: Path) -> None:
    """Remove test suites and caches that are not needed at runtime."""
    removed = 0
    for pkg in ("pandas", "numpy", "pyarrow", "plotly", "altair", "pydeck"):
        base = site / pkg
        if not base.exists():
            continue
        for d in list(base.rglob("tests")):
            if d.is_dir():
                shutil.rmtree(d, ignore_errors=True)
                removed += 1
    for d in list(site.rglob("__pycache__")):
        shutil.rmtree(d, ignore_errors=True)
    bin_dir = site / "bin"
    if bin_dir.exists():
        shutil.rmtree(bin_dir, ignore_errors=True)
    log(f"slimmed: removed {removed} test folders")


LAUNCHER_DIR = ROOT / "portable" / "launcher"


def build_launcher(dest: Path) -> None:
    """Compile BuyBox.exe (C + icon/version resources) with zig's cross-compiling C compiler.

    Works on Windows, Linux and macOS: ``pip install ziglang`` provides the compiler.
    """
    try:
        import ziglang  # noqa: F401
    except ImportError:
        log("installing the ziglang C compiler (pip install ziglang)")
        subprocess.run([sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "ziglang"],
                       check=True)
    cache = dest.parent.parent / ".zig-cache"
    env = {**os.environ, "ZIG_LOCAL_CACHE_DIR": str(cache), "ZIG_GLOBAL_CACHE_DIR": str(cache)}
    cmd = [sys.executable, "-m", "ziglang", "cc", "-target", "x86_64-windows-gnu", "-O2", "-municode", "-s",
           "launcher.c", "buybox.rc", "-o", str(dest)]
    log("compiling BuyBox.exe")
    subprocess.run(cmd, check=True, cwd=LAUNCHER_DIR, env=env)
    if not dest.exists() or dest.read_bytes()[:2] != b"MZ":
        raise SystemExit("BuyBox.exe was not produced.")


def copy_app(app_dir: Path) -> None:
    app_dir.mkdir(parents=True, exist_ok=True)
    for item in APP_ITEMS:
        src = ROOT / item
        dst = app_dir / item
        if src.is_dir():
            shutil.copytree(src, dst, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
        elif src.exists():
            shutil.copy2(src, dst)
    (app_dir / "data" / "runs").mkdir(parents=True, exist_ok=True)
    log("copied app files")


def precompile(pkg: Path) -> None:
    """Byte-compile with the embedded interpreter (only possible on Windows)."""
    py = pkg / "python" / "python.exe"
    if sys.platform != "win32" or not py.exists():
        log("skipping byte-compilation (not on Windows)")
        return
    log("byte-compiling site-packages")
    subprocess.run([str(py), "-m", "compileall", "-q", "-j", "0", str(pkg / "python" / "Lib" / "site-packages"),
                    str(pkg / "app")], check=False)


def make_zip(pkg: Path, zip_path: Path) -> None:
    log(f"writing {zip_path.name}")
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for f in sorted(pkg.rglob("*")):
            if f.is_file():
                zf.write(f, Path(pkg.name) / f.relative_to(pkg))


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--python", default=DEFAULT_PYTHON, help="embeddable Python version (default %(default)s)")
    ap.add_argument("--embed-zip", type=Path, help="use a local python-X.Y.Z-embed-amd64.zip instead of downloading")
    ap.add_argument("--requirements", type=Path, default=ROOT / "portable" / "requirements-portable.txt")
    ap.add_argument("--out", type=Path, default=ROOT / "dist")
    ap.add_argument("--no-zip", action="store_true", help="only build the folder")
    ap.add_argument("--no-slim", action="store_true", help="keep test folders of big packages")
    args = ap.parse_args()

    out = args.out.resolve()
    pkg = out / APP_NAME
    if pkg.exists():
        log(f"removing previous {pkg}")
        shutil.rmtree(pkg)
    py_dir = pkg / "python"
    py_dir.mkdir(parents=True)

    embed = args.embed_zip
    if embed is None:
        embed = out / f"python-{args.python}-embed-amd64.zip"
        if not embed.exists():
            download(EMBED_URL.format(v=args.python), embed)
    with zipfile.ZipFile(embed) as zf:
        zf.extractall(py_dir)
    log(f"extracted embeddable Python {args.python}")
    patch_pth(py_dir)

    site = py_dir / "Lib" / "site-packages"
    site.mkdir(parents=True, exist_ok=True)
    pip_install(site, args.requirements, args.python)
    if not args.no_slim:
        slim(site)

    copy_app(pkg / "app")
    build_launcher(pkg / "BuyBox.exe")
    (pkg / "README_PORTABLE.txt").write_bytes(README_PORTABLE.replace("\n", "\r\n").encode("utf-8-sig"))
    precompile(pkg)

    if not args.no_zip:
        make_zip(pkg, out / f"{APP_NAME}_portable_win64.zip")
    size = sum(f.stat().st_size for f in pkg.rglob("*") if f.is_file()) / 1e6
    log(f"done: {pkg} ({size:.0f} MB unpacked)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
