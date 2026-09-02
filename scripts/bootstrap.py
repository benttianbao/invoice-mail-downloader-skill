#!/usr/bin/env python3
"""为技能创建隔离的 Python 环境并安装依赖。"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import subprocess
import sys
import venv


APP_NAME = "invoice-mail-downloader"


def data_dir() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_NAME
    if os.name == "nt":
        return Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / APP_NAME
    raise SystemExit("首版仅支持 macOS 和 Windows。")


def venv_python(root: Path) -> Path:
    return root / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def main() -> int:
    parser = argparse.ArgumentParser(description="初始化邮箱发票下载技能")
    parser.parse_args()

    root = data_dir()
    env_dir = root / "venv"
    root.mkdir(parents=True, exist_ok=True)
    if not venv_python(env_dir).exists():
        print(f"正在创建隔离环境：{env_dir}")
        venv.EnvBuilder(with_pip=True).create(env_dir)

    python = venv_python(env_dir)
    requirements = Path(__file__).resolve().parent.parent / "requirements.txt"
    subprocess.run(
        [str(python), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(requirements)],
        check=True,
    )
    print("初始化完成。后续请使用 scripts/run_skill.py 调用技能。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
