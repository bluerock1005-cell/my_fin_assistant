#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""core/feature_base.py — FeatureModule 统一接口。

每个功能模块（features/<name>/）对外暴露一个继承 FeatureModule 的类，
主窗口靠 get_widget() 统一加载，不用互相 import 细节。

约定：
    - name: 唯一标识（建议与目录名一致）
    - icon: qtawesome 图标名（如 "fa5s.home"、"file-excel"），主窗口据此渲染导航图标
    - get_widget(self, parent=None): 返回该功能的主界面 QWidget
"""
from __future__ import annotations

from abc import ABC, abstractmethod

from PySide6.QtWidgets import QWidget


class FeatureModule(ABC):
    """功能模块统一接口。

    子类需覆盖：
        - name: str — 模块唯一标识
        - icon: str — qtawesome 图标名（可选，如 "fa5s.home"）
        - get_widget(parent=None) -> QWidget — 返回主界面控件
    """

    # 子类覆盖以下类属性
    name: str = ""
    icon: str = ""

    def get_widget(self, parent: QWidget | None = None) -> QWidget:
        """返回该功能的主界面 Widget。

        Args:
            parent: 父容器，由主窗口传入 QStackedWidget 或 None。

        Returns:
            该功能的根 QWidget。
        """
        raise NotImplementedError(f"{type(self).__name__}.get_widget() 未实现")
