#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""core/app_config.py — 应用级配置与路径常量。

所有跨模块共用的路径、应用元信息集中在这里，避免散落硬编码。
按图片目录结构：data/ 在项目根，resources/ 在项目根（待建）。
"""
from __future__ import annotations

from pathlib import Path

# 项目根目录：core/app_config.py 的上上级即项目根
BASE_DIR: Path = Path(__file__).resolve().parent.parent

# 本地数据 / 缓存目录（与 features/ core/ 平级）
DATA_DIR: Path = BASE_DIR / "data"

# 资源目录：图标、QSS 样式表等（待建，先预留）
RESOURCES_DIR: Path = BASE_DIR / "resources"
ICONS_DIR: Path = RESOURCES_DIR / "icons"
STYLES_FILE: Path = RESOURCES_DIR / "styles.qss"

# bank_classify 示例数据（已迁入 data/）
BANKS_SAMPLE_FILE: Path = DATA_DIR / "banks.txt"

# 应用元信息
APP_NAME: str = "我的财务助手"
APP_VERSION: str = "0.1.0"
ORG_NAME: str = "MyFinAssistant"

# 默认输出目录：生成的 Excel 等放在 data/output/ 下
DEFAULT_OUTPUT_DIR: Path = DATA_DIR / "output"


def ensure_dirs() -> None:
    """启动时调用，保证关键目录存在。"""
    for d in (DATA_DIR, DEFAULT_OUTPUT_DIR, RESOURCES_DIR, ICONS_DIR):
        d.mkdir(parents=True, exist_ok=True)
