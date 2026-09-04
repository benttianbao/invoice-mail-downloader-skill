#!/usr/bin/env python3
"""为技能创建隔离的 Python 环境并安装依赖。"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import platform
import shutil
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


def install_xparse_cli(root: Path) -> Path:
    """把 TextIn XParse CLI 安装在技能应用数据目录，避免依赖全局 npm 包。"""
    npm = os.environ.get("INVOICE_NPM_PATH", "").strip() or shutil.which("npm") or shutil.which("npm.cmd")
    if not npm:
        raise SystemExit("未找到 npm；TextIn XParse CLI 需要 Node.js 18 或更高版本及 npm。")
    install_root = root / "xparse-cli"
    subprocess.run([npm, "install", "--prefix", str(install_root), "xparse-cli"], check=True)

    package_json = install_root / "node_modules" / "xparse-cli" / "package.json"
    metadata = json.loads(package_json.read_text(encoding="utf-8"))
    version = str(metadata["version"])
    system = "darwin" if sys.platform == "darwin" else "windows" if os.name == "nt" else ""
    architecture = {"arm64": "arm64", "aarch64": "arm64", "x86_64": "amd64", "amd64": "amd64"}.get(
        platform.machine().lower()
    )
    if not system or not architecture:
        raise SystemExit(f"TextIn XParse CLI 暂不支持当前平台：{sys.platform}/{platform.machine()}")
    binary_package = f"xparse-cli-{system}-{architecture}@{version}"
    subprocess.run([npm, "install", "--prefix", str(install_root), binary_package], check=True)

    command = install_root / "node_modules" / ".bin" / ("xparse-cli.cmd" if os.name == "nt" else "xparse-cli")
    if not command.is_file():
        raise SystemExit("TextIn XParse CLI 安装完成但未找到启动文件。")
    return command


def main() -> int:
    parser = argparse.ArgumentParser(description="初始化邮箱发票下载技能")
    parser.add_argument("--with-receipts", action="store_true", help="同时安装 TextIn XParse CLI，用于报销凭证识别")
    args = parser.parse_args()

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
    if args.with_receipts:
        xparse = install_xparse_cli(root)
        print(f"TextIn XParse CLI 已安装：{xparse}")
    else:
        print(f'已跳过 TextIn XParse CLI。识别报销凭证时请重新运行：python3 "{Path(__file__).resolve()}" --with-receipts')
    print("初始化完成。后续请使用脚本的绝对路径调用 run_skill.py。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
