#!/usr/bin/env python3
"""使用隔离环境启动主程序。"""

from __future__ import annotations

import os
from pathlib import Path
import subprocess
import sys


APP_NAME = "invoice-mail-downloader"


def data_dir() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_NAME
    if os.name == "nt":
        return Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / APP_NAME
    raise SystemExit("首版仅支持 macOS 和 Windows。")


def main() -> int:
    env_dir = data_dir() / "venv"
    python = env_dir / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    if not python.exists():
        bootstrap = Path(__file__).resolve().parent / "bootstrap.py"
        launcher = "py -3" if os.name == "nt" else "python3"
        raise SystemExit(f'尚未初始化。请先运行：{launcher} "{bootstrap}"')
    target = Path(__file__).resolve().parent / "invoice_mail.py"
    return subprocess.run([str(python), str(target), *sys.argv[1:]]).returncode


if __name__ == "__main__":
    raise SystemExit(main())
