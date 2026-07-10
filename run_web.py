#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run_web.py — pywebview 入口（兼容新旧两种模式）。

用法：
    # 生产模式（加载 Vite 构建产物 dist/）
    python run_web.py

    # 开发模式（连接 vite dev server，热更新）
    python run_web.py --dev

旧的 PySide6 入口仍是 main.py：
    python main.py
"""
from __future__ import annotations
import sys


def main() -> None:
    if "--dev" in sys.argv:
        _run_dev()
    else:
        from webview_api import main as _web_main  # noqa: PLC0415
        _web_main()


def _run_dev() -> None:
    """开发模式：启动 vite dev server 后打开 pywebview。"""
    import subprocess
    import time
    import webview
    from pathlib import Path

    from webview_api import WebAPI  # noqa: PLC0415
    from core import app_config  # noqa: PLC0415

    web_dir = Path(__file__).resolve().parent / "web"

    print("🚀 启动 Vite 开发服务器...", file=sys.stderr)
    proc = subprocess.Popen(
        [sys.executable or "npx", "vite"],
        cwd=str(web_dir),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )

    # 等待 server 就绪
    time.sleep(3)

    api = WebAPI()
    url = "http://localhost:5173"
    window = webview.create_window(
        title=f"{app_config.APP_NAME} (Web Dev)",
        url=url,
        js_api=api,
        width=1200,
        height=780,
        min_size=(960, 620),
    )
    api.window = window
    try:
        webview.start()
    finally:
        proc.terminate()


if __name__ == "__main__":
    raise SystemExit(main())
