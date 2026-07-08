#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""core/main_window.py — 主窗口壳（PySide6 + Fluent）。

主窗口入口只做"壳"，不写任何业务逻辑：
    - 左侧可折叠侧边栏（qtawesome 图标 + 文本，160ms OutCubic 折叠动画）
    - 右侧 QStackedWidget 内容区，按 FeatureModule 注册表懒加载
    - 视觉完全由 core.theme 的设计系统驱动（apply_theme）

结构：
    QMainWindow
      central QWidget
        horizontal shell layout
          QFrame#sideBar
            brand row（图标 + 标题 + 折叠按钮）
            QListWidget#sideNav
            bottom action（关于）
          QStackedWidget#contentStack
"""
from __future__ import annotations

from typing import List

from PySide6.QtCore import QEasingCurve, QPropertyAnimation, Qt, QSize
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QListView,
    QMainWindow,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from . import app_config, theme
from .feature_base import FeatureModule

_EXPANDED = theme.SIDEBAR["expanded"]
_COLLAPSED = theme.SIDEBAR["collapsed"]
_ANIM_MS = theme.SIDEBAR["anim_ms"]


class MainWindow(QMainWindow):
    """应用主窗口：可折叠侧边栏导航 + 内容区（纯壳，无业务逻辑）。"""

    def __init__(self, features: List[FeatureModule] | None = None) -> None:
        super().__init__()
        self._features: List[FeatureModule] = []
        # feature 索引 -> (文本, qtawesome 图标名)
        self._nav_meta: List[tuple[str, str]] = []
        self._widgets: dict[int, QWidget] = {}  # feature 索引 -> 已创建的 widget
        self._name_to_index: dict[str, int] = {}  # feature.name -> 索引（首页卡片跳转用）
        self._connected: set[int] = set()  # 已连接 feature_requested 的 widget id
        self._collapsed = False

        self._setup_ui()
        self.setWindowTitle(app_config.APP_NAME)
        self.resize(1080, 720)
        theme.apply_theme(self)

        for feat in features or []:
            self.register_feature(feat)
        # 默认选中第一个
        if self._nav.count() > 0:
            self._nav.setCurrentRow(0)

    # ----- UI 构建 ---------------------------------------------------------

    def _setup_ui(self) -> None:
        central = QWidget(self)
        central.setObjectName("central")
        shell = QHBoxLayout(central)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)

        # ---- 侧边栏 ----
        self._side = QFrame(self)
        self._side.setObjectName("sideBar")
        self._side.setProperty("collapsed", "false")
        self._side.setMinimumWidth(_EXPANDED)
        self._side.setMaximumWidth(_EXPANDED)
        side_lay = QVBoxLayout(self._side)
        side_lay.setContentsMargins(12, 12, 12, 12)
        side_lay.setSpacing(4)

        # 品牌行
        self._brand_row = QHBoxLayout()
        self._brand_icon = QLabel(self._side)
        self._brand_icon.setPixmap(
            theme.nav_icon("fa5s.wallet").pixmap(QSize(theme.ICON["brand"], theme.ICON["brand"]))
        )
        self._brand_title = QLabel(app_config.APP_NAME, self._side)
        self._brand_title.setObjectName("brandLabel")
        self._collapse_btn = QPushButton(self._side)
        self._collapse_btn.setObjectName("collapseBtn")
        self._collapse_btn.setIcon(theme.nav_icon("fa5s.angle-left"))
        self._collapse_btn.setIconSize(QSize(16, 16))
        self._collapse_btn.setFixedSize(28, 28)
        self._collapse_btn.setToolTip("折叠/展开侧边栏")
        self._collapse_btn.clicked.connect(self.toggle_sidebar)
        self._brand_row.addWidget(self._brand_icon)
        self._brand_row.addWidget(self._brand_title, stretch=1)
        self._brand_row.addWidget(self._collapse_btn)
        side_lay.addLayout(self._brand_row)

        # 导航列表
        self._nav = QListWidget(self._side)
        self._nav.setObjectName("sideNav")
        self._nav.setIconSize(QSize(theme.ICON["nav_expanded"], theme.ICON["nav_expanded"]))
        self._nav.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._nav.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._nav.setMovement(QListView.Static)
        self._nav.setFlow(QListView.TopToBottom)
        self._nav.currentRowChanged.connect(self._on_nav_changed)
        side_lay.addWidget(self._nav, stretch=1)

        # 底部操作
        self._about_btn = QPushButton(self._side)
        self._about_btn.setObjectName("collapseBtn")
        self._about_btn.setIcon(theme.nav_icon("fa5s.info-circle"))
        self._about_btn.setIconSize(QSize(16, 16))
        self._about_btn.setText("关于")
        self._about_btn.setFixedHeight(34)
        self._about_btn.clicked.connect(self._on_about)
        side_lay.addWidget(self._about_btn)

        shell.addWidget(self._side)

        # ---- 内容区 ----
        self._stack = QStackedWidget(self)
        self._stack.setObjectName("contentStack")
        self._stack.addWidget(self._placeholder_widget())
        shell.addWidget(self._stack, stretch=1)

        self.setCentralWidget(central)

    @staticmethod
    def _placeholder_widget() -> QWidget:
        w = QWidget()
        w.setObjectName("page")
        lay = QHBoxLayout(w)
        lay.setAlignment(Qt.AlignCenter)
        lbl = QLabel("暂无功能，请在 main.py 的 FEATURES 注册表中添加模块。", w)
        lbl.setStyleSheet("color: #9CA3AF; font-size: 14px;")
        lay.addWidget(lbl)
        return w

    # ----- Feature 注册（统一装载）-----------------------------------------

    def register_feature(self, feature: FeatureModule) -> int:
        """注册一个功能模块到侧边栏和内容区。

        Args:
            feature: 继承 FeatureModule 的实例（其 icon 为 qtawesome 名）。

        Returns:
            该功能在注册列表中的索引。
        """
        idx = len(self._features)
        self._features.append(feature)
        self._name_to_index[feature.name] = idx
        text = feature.name or f"功能{idx}"
        icon_name = feature.icon or "fa5s.circle"
        self._nav_meta.append((text, icon_name))

        item = QListWidgetItem(theme.nav_icon(icon_name), text)
        item.setData(Qt.UserRole, (text, icon_name))
        item.setSizeHint(QSize(40, theme.SIDEBAR["item_h"]))
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self._nav.addItem(item)
        return idx

    # ----- 切换 ------------------------------------------------------------

    def _on_nav_changed(self, row: int) -> None:
        if row < 0 or row >= len(self._features):
            return
        feat = self._features[row]
        if row not in self._widgets:
            w = feat.get_widget(parent=self._stack)
            w.setObjectName("page")
            self._widgets[row] = w
            self._stack.addWidget(w)
            # 首页卡片点击跳转：连接一次
            if hasattr(w, "feature_requested") and id(w) not in self._connected:
                w.feature_requested.connect(self._go_to_feature)
                self._connected.add(id(w))
        self._stack.setCurrentWidget(self._widgets[row])

    def _go_to_feature(self, name: str) -> None:
        """首页卡片点击时跳转到对应功能（按 FeatureModule.name 查找）。"""
        idx = self._name_to_index.get(name)
        if idx is not None:
            self._nav.setCurrentRow(idx)

    # ----- 折叠 ------------------------------------------------------------

    def toggle_sidebar(self) -> None:
        self._collapsed = not self._collapsed
        collapsed = self._collapsed

        # 导航项：文本显隐 + 视图模式 + 图标尺寸
        self._nav.setViewMode(QListView.IconMode if collapsed else QListView.ListMode)
        self._nav.setGridSize(QSize(48, theme.SIDEBAR["item_h"]) if collapsed else QSize())
        self._nav.setIconSize(
            QSize(theme.ICON["nav_collapsed"], theme.ICON["nav_collapsed"]) if collapsed
            else QSize(theme.ICON["nav_expanded"], theme.ICON["nav_expanded"])
        )
        for i in range(self._nav.count()):
            item = self._nav.item(i)
            text, icon_name = item.data(Qt.UserRole)
            item.setText("" if collapsed else text)
            item.setTextAlignment(Qt.AlignCenter if collapsed else Qt.AlignLeft | Qt.AlignVCenter)
            item.setSizeHint(
                QSize(40, theme.SIDEBAR["item_h"]) if not collapsed
                else QSize(48, theme.SIDEBAR["item_h"])
            )

        # 品牌区：标题显隐
        self._brand_title.setVisible(not collapsed)
        self._about_btn.setText("" if collapsed else "关于")
        self._about_btn.setToolTip("关于" if collapsed else "")
        # 折叠按钮图标翻转
        self._collapse_btn.setIcon(
            theme.nav_icon("fa5s.angle-right" if collapsed else "fa5s.angle-left")
        )
        self._side.setProperty("collapsed", "true" if collapsed else "false")
        self._side.style().unpolish(self._side)
        self._side.style().polish(self._side)

        # 宽度动画
        anim = QPropertyAnimation(self._side, b"minimumWidth")
        anim.setDuration(_ANIM_MS)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        anim.setStartValue(self._side.width())
        anim.setEndValue(_COLLAPSED if collapsed else _EXPANDED)
        anim.valueChanged.connect(lambda v: self._side.setMaximumWidth(int(v)))
        anim.finished.connect(
            lambda: (self._side.setMinimumWidth(_COLLAPSED if collapsed else _EXPANDED),
                     self._side.setMaximumWidth(_COLLAPSED if collapsed else _EXPANDED))
        )
        anim.start()

    # ----- 关于 ------------------------------------------------------------

    def _on_about(self) -> None:
        from . import utils
        utils.info(
            app_config.APP_NAME,
            f"{app_config.APP_NAME}  v{app_config.APP_VERSION}\n\n"
            "个人财务桌面助手。\n"
            "基于 PySide6 + qtawesome 实现的 Fluent 风格界面。",
            parent=self,
        )
