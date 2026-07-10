#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""webview_main.py — pywebview + React 桌面入口。

启动方式：
    .\.venv\Scripts\python.exe webview_main.py

流程：
    1. 先检查 web/dist/ 是否存在（Vite 构建产物）
    2. 若不存在则提示先 npm run build
    3. 启动 pywebview 加载 dist/index.html
    4. 通过 js_api 暴露 WebAPI 给 React 前端调用

旧的 PySide6 入口仍是 main.py，两者并存互不干扰。
"""
from __future__ import annotations

from pathlib import Path
import sys

import webview

from webview_api import WebAPI
from core import app_config


def main() -> None:
    app_config.ensure_dirs()

    # 检查 Vite 构建产物是否存在
    dist_dir = Path(__file__).resolve().parent / "web" / "dist"
    index_html = dist_dir / "index.html"

    if not index_html.exists():
        print("=" * 56, file=sys.stderr)
        print("  ⚠️  前端未构建！请先执行：", file=sys.stderr)
        print("", file=sys.stderr)
        print("    cd web && npm install && npm run build", file=sys.stderr)
        print("", file=sys.stderr)
        print("  或使用开发模式（需要启动 vite dev server）：", file=sys.stderr)
        print("    python run_web_dev.py", file=sys.stderr)
        print("=" * 56, file=sys.stderr)
        sys.exit(1)

    api = WebAPI()

    url = index_html.as_uri()

    window = webview.create_window(
        title=f"{app_config.APP_NAME} (Web版)",
        url=url,
        js_api=api,
        width=1200,
        height=780,
        min_size=(960, 620),
    )

    api.window = window
    webview.start()


if __name__ == "__main__":
    raise SystemExit(main())
