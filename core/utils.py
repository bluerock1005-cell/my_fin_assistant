#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""core/utils.py — 通用工具函数。

集中放跨模块复用的小工具，避免各 feature 各自重复造轮子。
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import QMessageBox, QWidget

from . import app_config


def load_stylesheet(path: Optional[Path] = None) -> str:
    """读取 QSS 样式表。未传 path 则读 app_config.STYLES_FILE。

    文件不存在时返回空串（不报错，便于 resources/ 尚未建好时也能启动）。
    """
    target = path or app_config.STYLES_FILE
    if not target.exists():
        return ""
    try:
        return target.read_text(encoding="utf-8")
    except OSError:
        return ""


def info(title: str, text: str, parent: Optional[QWidget] = None) -> None:
    """信息提示框。"""
    QMessageBox.information(parent, title, text)


def warn(title: str, text: str, parent: Optional[QWidget] = None) -> None:
    """警告提示框。"""
    QMessageBox.warning(parent, title, text)


def error(title: str, text: str, parent: Optional[QWidget] = None) -> None:
    """错误提示框。"""
    QMessageBox.critical(parent, title, text)


def confirm(title: str, text: str, parent: Optional[QWidget] = None) -> bool:
    """是/否确认框，返回 True 表示用户选了"是"。"""
    return QMessageBox.question(parent, title, text, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) == QMessageBox.Yes
