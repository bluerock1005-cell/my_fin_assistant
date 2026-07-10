#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""core/theme.py — Fluent 设计系统与 QSS 单一真相来源。

按 pyside6-fluent-gui-design skill 的 token 规范把设计系统翻译成 QSS 与
widget 常量（而非 web CSS 变量）。任何界面都从这里取颜色/间距/圆角，不要在
各处手写散落的样式字符串，组件级样式（如主按钮色）才允许局部覆盖。

设计方向：浅色、冷静的 Windows Fluent 工作台
    - 白底侧边栏 + 浅灰内容区
    - 主色克制蓝，关键操作用 primary；状态色只在高对比处出现
    - 圆角小（控件 6、卡片 8、选中项 10）
    - 字体紧凑：标题 22 / 段落 16 / 正文 13-14 / 辅助 12
"""
from __future__ import annotations

# ---------------------------------------------------------------------------
# 1. 颜色 token（calm light Fluent palette）
# ---------------------------------------------------------------------------
COLOR = {
    #  surfaces
    "bg": "#F4F5F7",            # 应用画布（内容区背景）
    "surface": "#FFFFFF",       # 卡片 / 侧边栏 / 浮层
    "surface_muted": "#F8FAFC", # 次级表面（输入框底、禁用以太浅处）
    "border": "#E5E7EB",        # 描边
    "border_strong": "#D1D5DB",
    #  text
    "text": "#1F2937",          # 主文本
    "muted": "#6B7280",         # 次要文本 / 提示
    "faint": "#9CA3AF",         # 占位 / 极弱
    #  primary（克制蓝）
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_pressed": "#1E40AF",
    "primary_soft": "#EAF2FF",  # 选中项底色
    #  state
    "danger": "#DC2626",
    "danger_hover": "#B91C1C",
    "danger_soft": "#FEF2F2",
    "success": "#16A34A",
    "success_soft": "#F0FDF4",
    "warning": "#D97706",
    "warning_soft": "#FFFBEB",
    #  sidebar
    "nav_bg": "#FFFFFF",
    "nav_text": "#374151",
    "nav_text_active": "#2563EB",
    #  icon
    "icon": "#64748B",
    "icon_active": "#2563EB",
}

# ---------------------------------------------------------------------------
# 2. 间距 / 圆角 / 字号 / 图标尺寸
# ---------------------------------------------------------------------------
SPACING = {4: 4, 8: 8, 12: 12, 16: 16, 20: 20, 24: 24, 32: 32}
RADIUS = {"control": 6, "card": 8, "nav_selected": 10}
TYPE = {
    "title": 22,        # 视图标题
    "section": 16,      # 段落标题
    "body": 14,         # 正文 / 控件
    "helper": 12,       # 辅助 / 日志
}
ICON = {
    "nav_expanded": 18,
    "nav_collapsed": 21,
    "action": 16,
    "brand": 22,
}
# 侧边栏尺寸
SIDEBAR = {"expanded": 236, "collapsed": 72, "item_h": 46, "anim_ms": 160}

# ---------------------------------------------------------------------------
# 2.1 页面级统一间距（各功能模块共用，保证「卡片之间」「卡片与页边」距离一致）
# ---------------------------------------------------------------------------
PAGE_PAD = SPACING[24]   # 内容区四周留白（页边到第一张/最后一张卡片）
CARD_GAP = SPACING[16]   # 卡片与卡片之间的垂直间距

# ---------------------------------------------------------------------------
# 3. qtawesome 图标助手
# ---------------------------------------------------------------------------
_NAV_FAMILY = "fa5s"  # 统一用一个图标族，避免视觉权重不一致


def icon_name(name: str) -> str:
    """把简写或完整名规整成 qtawesome 可用的名。

    - 已带点号（如 "fa5s.home"）原样返回
    - 纯名（如 "home"）自动补默认族
    """
    if "." in name:
        return name
    return f"{_NAV_FAMILY}.{name}"


def nav_icon(name: str, active: bool = False):
    """返回 qtawesome 图标（延迟导入，避免未装 qtawesome 时导入失败）。"""
    import qtawesome as qta
    color = COLOR["icon_active"] if active else COLOR["icon"]
    return qta.icon(icon_name(name), color=color)


# ---------------------------------------------------------------------------
# 4. 完整 Fluent QSS（由上述 token 渲染，单一真相来源）
# ---------------------------------------------------------------------------
def _qss() -> str:
    c = COLOR
    r = RADIUS
    return f"""
/* ===== 全局 ===== */
QWidget {{
    background-color: {c['bg']};
    color: {c['text']};
    font-family: "Microsoft YaHei", "Segoe UI", system-ui, sans-serif;
    font-size: {TYPE['body']}px;
}}
QMainWindow, QStackedWidget#contentStack {{
    background-color: {c['bg']};
}}

/* ===== 侧边栏 ===== */
QFrame#sideBar {{
    background-color: {c['nav_bg']};
    border-right: 1px solid {c['border']};
}}
QLabel#brandLabel {{
    color: {c['text']};
    font-size: 15px;
    font-weight: 600;
    background: transparent;
}}
QPushButton#collapseBtn {{
    background: transparent;
    border: none;
    color: {c['muted']};
    padding: 4px;
    border-radius: {r['control']}px;
}}
QPushButton#collapseBtn:hover {{
    background-color: {c['surface_muted']};
    color: {c['text']};
}}

QListWidget#sideNav {{
    background: transparent;
    border: none;
    outline: none;
}}
QListWidget#sideNav::item {{
    min-height: {SIDEBAR['item_h']}px;
    padding: 8px 12px 8px 12px;
    margin: 2px 8px;
    border-radius: {r['nav_selected']}px;
    color: {c['nav_text']};
}}
QListWidget#sideNav::item:hover {{
    background-color: {c['surface_muted']};
}}
QListWidget#sideNav::item:selected {{
    background-color: {c['primary_soft']};
    color: {c['nav_text_active']};
    font-weight: 600;
}}
/* 折叠态：图标居中、做成方形命中区 */
QFrame#sideBar[collapsed="true"] QListWidget#sideNav::item {{
    min-width: 48px;
    max-width: 48px;
    min-height: 46px;
    padding: 0;
    margin: 2px 0;
    border-radius: {r['nav_selected']}px;
}}

/* ===== 内容区页面 ===== */
QWidget#page {{
    background-color: {c['bg']};
}}
QLabel#pageTitle {{
    font-size: {TYPE['title']}px;
    font-weight: 600;
    color: {c['text']};
    background: transparent;
}}
QLabel#pageDesc {{
    font-size: {TYPE['helper']}px;
    color: {c['muted']};
    background: transparent;
}}

/* ===== 卡片 ===== */
QFrame#card {{
    background-color: {c['surface']};
    border: 1px solid {c['border']};
    border-radius: {r['card']}px;
}}
QLabel#cardTitle {{
    font-size: {TYPE['section']}px;
    font-weight: 600;
    color: {c['text']};
    background: transparent;
}}

/* ===== 通用控件 ===== */
QLabel {{
    background: transparent;
}}
QLineEdit, QTextEdit, QPlainTextEdit {{
    background-color: {c['surface']};
    border: 1px solid {c['border']};
    border-radius: {r['control']}px;
    padding: 6px 10px;
    color: {c['text']};
    selection-background-color: {c['primary_soft']};
}}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
    border: 1px solid {c['primary']};
}}
QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled {{
    background-color: {c['surface_muted']};
    color: {c['faint']};
}}
QTextEdit#logView, QPlainTextEdit#logView {{
    background-color: #0F172A;
    color: #E2E8F0;
    border: 1px solid {c['border']};
    font-family: "Cascadia Code", "Consolas", "Courier New", monospace;
    font-size: {TYPE['helper']}px;
}}

QPushButton {{
    background-color: {c['surface']};
    border: 1px solid {c['border_strong']};
    border-radius: {r['control']}px;
    padding: 6px 14px;
    color: {c['text']};
}}
QPushButton:hover {{ border-color: {c['muted']}; }}
QPushButton:pressed {{ background-color: {c['surface_muted']}; }}
QPushButton:disabled {{
    background-color: {c['surface_muted']};
    color: {c['faint']};
    border-color: {c['border']};
}}
QPushButton#primary {{
    background-color: {c['primary']};
    border: 1px solid {c['primary']};
    color: #FFFFFF;
    font-weight: 600;
}}
QPushButton#primary:hover {{ background-color: {c['primary_hover']}; border-color: {c['primary_hover']}; }}
QPushButton#primary:pressed {{ background-color: {c['primary_pressed']}; }}
QPushButton#primary:disabled {{
    background-color: {c['primary']};
    color: rgba(255,255,255,0.6);
    border-color: {c['primary']};
}}
QPushButton#danger {{
    background-color: {c['danger']};
    border: 1px solid {c['danger']};
    color: #FFFFFF;
}}
QPushButton#danger:hover {{ background-color: {c['danger_hover']}; border-color: {c['danger_hover']}; }}

QProgressBar {{
    border: none;
    background: {c['surface_muted']};
    border-radius: {r['control']}px;
    text-align: center;
    height: 8px;
    color: transparent;
}}
QProgressBar::chunk {{
    background-color: {c['primary']};
    border-radius: {r['control']}px;
}}

QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar::handle:vertical {{
    background: {c['border_strong']};
    border-radius: 5px;
}}
QScrollBar::handle:vertical:hover {{ background: {c['muted']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
"""


# 渲染后的静态 QSS（模块加载时生成一次）
BUILD_QSS = _qss()


def apply_theme(window) -> None:
    """把设计系统应用到主窗口（QSS 作用于整个窗口树）。

    注：QSS 由 theme.BUILD_QSS 提供，单一真相来源；不使用外部 .qss 文件，
    避免 token 与样式表两处不同步。
    """
    window.setStyleSheet(BUILD_QSS)
