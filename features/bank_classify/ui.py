#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""features/bank_classify/ui.py — 银行承兑汇票分类界面（PySide6 + Fluent）。

UI 和逻辑严格分离：本文件只管界面与交互，
纯计算/数据处理放在 classify_logic.py（可单独测试或替换，不用启动界面）。

长耗时操作（生成 Excel）走 core.worker 后台线程，避免界面卡死。
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import concurrent.futures

from PySide6.QtCore import QPoint, QRect, QSize, Qt, QTimer
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLayout,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core import app_config, theme, utils
from core.feature_base import FeatureModule
from core.worker import run_in_thread

# 纯逻辑层：不导入任何 UI 依赖，可独立测试
from . import classify_logic as _logic


class FlowLayout(QLayout):
    """轻量流式布局：子项从左到右排，到边界自动换行（用于白名单标签云）。

    仅依赖 Qt 标准 API，无需额外组件；作为某个 QWidget 的顶层布局时
    会正确参与 heightForWidth。
    """

    def __init__(self, parent: QWidget | None = None, hspacing: int = 8, vspacing: int = 8) -> None:
        super().__init__(parent)
        self._hspacing = hspacing
        self._vspacing = vspacing
        self._items: list = []
        self.setContentsMargins(0, 0, 0, 0)

    def addItem(self, item) -> None:  # noqa: N802
        self._items.append(item)

    def count(self) -> int:  # noqa: N802
        return len(self._items)

    def itemAt(self, index: int):  # noqa: N802
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def takeAt(self, index: int):  # noqa: N802
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        return self._do_layout(QRect(0, 0, width, 0), True)

    def setGeometry(self, rect: QRect) -> None:  # noqa: N802
        super().setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self):  # noqa: N802
        return self.minimumSize()

    def minimumSize(self):  # noqa: N802
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        left, top, right, bottom = self.getContentsMargins()
        size += QSize(left + right, top + bottom)
        return size

    def _do_layout(self, rect: QRect, test_only: bool) -> int:
        left, top, right, bottom = self.getContentsMargins()
        effective = rect.adjusted(+left, +top, -right, -bottom)
        x = effective.x()
        y = effective.y()
        line_height = 0
        for item in self._items:
            w = item.sizeHint().width()
            h = item.sizeHint().height()
            next_x = x + w + self._hspacing
            if next_x - self._hspacing > effective.right() and line_height > 0:
                x = effective.x()
                y = y + line_height + self._vspacing
                next_x = x + w + self._hspacing
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), QSize(w, h)))
            x = next_x
            line_height = max(line_height, h)
        return y + line_height - rect.y() + bottom


class BankClassifyWidget(QWidget):
    """银行分类功能的具体 UI 控件（与 FeatureModule 解耦，方便单测）。"""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("page")
        self._out_path: Path = app_config.DEFAULT_OUTPUT_DIR / "银行承兑汇票分类.xlsx"
        self._thread = None  # 持有后台线程引用，避免被 GC
        self._executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        self._future = None
        self._poll_timer = None
        self.setAcceptDrops(True)  # 启用拖拽
        self._setup_ui()

    # ----- 布局 -----------------------------------------------------------

    def _setup_ui(self) -> None:
        # 页面根布局直接挂在 self 上（不使用滚动容器，对齐 notes_receivable_import）
        root = QVBoxLayout(self)
        root.setContentsMargins(theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD)
        root.setSpacing(theme.CARD_GAP)

        # 页头
        title = QLabel("票据银行分类", self)
        title.setObjectName("pageTitle")
        desc = QLabel(
            "按 21 家银行白名单，把每行一个银行全称分类为「21银行承兑汇票」与"
            "「非21银行承兑汇票」，并导出 Excel。",
            self,
        )
        desc.setObjectName("pageDesc")
        desc.setWordWrap(True)
        root.addWidget(title)
        root.addWidget(desc)
        root.addSpacing(theme.SPACING[4])

        # 21家银行白名单（可折叠展示）
        root.addWidget(self._build_whitelist_card())

        # 输入卡片
        input_card = QFrame(self)
        input_card.setObjectName("card")
        ic_lay = QVBoxLayout(input_card)
        ic_lay.setContentsMargins(theme.SPACING[16], theme.SPACING[16], theme.SPACING[16], theme.SPACING[16])
        ic_lay.setSpacing(theme.SPACING[12])

        # 标题行：左侧「银行全称输入」+ 右侧操作按钮（并排）
        head_row = QHBoxLayout()
        head_row.setSpacing(theme.SPACING[12])
        head_lbl = QLabel("银行全称输入", input_card)
        head_lbl.setObjectName("cardTitle")
        head_row.addWidget(head_lbl)
        head_row.addStretch(1)
        ic_lay.addLayout(head_row)

        self._txt_input = QPlainTextEdit(input_card)
        self._txt_input.setObjectName("logView")
        self._txt_input.setPlaceholderText(
            "在此粘贴银行名称，或点击上方按钮从文件加载...\n每行一个银行全称。"
        )
        self._txt_input.setMinimumHeight(140)
        self._txt_input.setMaximumHeight(140)
        self._txt_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        ic_lay.addWidget(self._txt_input)

        # 统一操作按钮：从文件加载 / 从剪贴板粘贴 / 清空 / 导出 Excel（置于标题右侧）
        btn_load = QPushButton("从文件加载", input_card)
        btn_load.setObjectName("primary")
        btn_load.setFixedHeight(34)
        btn_load.setMinimumWidth(120)
        btn_load.clicked.connect(self._load_file)
        head_row.addWidget(btn_load)

        btn_paste = QPushButton("从剪贴板粘贴", input_card)
        btn_paste.setObjectName("primary")
        btn_paste.setFixedHeight(34)
        btn_paste.setMinimumWidth(120)
        btn_paste.clicked.connect(self._paste_clipboard)
        head_row.addWidget(btn_paste)

        btn_clear = QPushButton("清空", input_card)
        btn_clear.setObjectName("primary")
        btn_clear.setFixedHeight(34)
        btn_clear.setMinimumWidth(120)
        btn_clear.clicked.connect(self._clear)
        head_row.addWidget(btn_clear)

        self._btn_run = QPushButton("导出 Excel", input_card)
        self._btn_run.setObjectName("primary")
        self._btn_run.setFixedHeight(34)
        self._btn_run.setMinimumWidth(120)
        self._btn_run.clicked.connect(self._process)
        head_row.addWidget(self._btn_run)

        root.addWidget(input_card)

        # 进度
        self._progress = QLabel("", self)
        self._progress.setObjectName("pageDesc")
        root.addWidget(self._progress)

        # 日志面板
        log_card = QFrame(self)
        log_card.setObjectName("card")
        lc_lay = QVBoxLayout(log_card)
        lc_lay.setContentsMargins(theme.SPACING[12], theme.SPACING[12], theme.SPACING[12], theme.SPACING[12])
        lc_lay.setSpacing(6)
        log_title = QLabel("运行日志（分类详情）", log_card)
        log_title.setObjectName("cardTitle")
        lc_lay.addWidget(log_title)
        self._log = QPlainTextEdit(log_card)
        self._log.setObjectName("logView")
        self._log.setReadOnly(True)
        self._log.setMinimumHeight(50)
        lc_lay.addWidget(self._log, stretch=1)
        root.addWidget(log_card, stretch=1)

    # ----- 21家银行白名单展示 -------------------------------------------

    def _make_chip(self, text: str, parent: QWidget) -> QLabel:
        """生成单个银行标签（Fluent 风格胶囊）。"""
        chip = QLabel(text, parent)
        chip.setObjectName("wlChip")
        chip.setAlignment(Qt.AlignCenter)
        chip.setStyleSheet(
            "QLabel#wlChip{"
            f"background-color:{theme.COLOR['primary_soft']};"
            f"color:{theme.COLOR['primary']};"
            f"border:1px solid {theme.COLOR['border']};"
            f"border-radius:{theme.RADIUS['control']}px;"
            f"padding:4px 12px;"
            f"font-size:{theme.TYPE['body']}px;"
            "}"
        )
        chip.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        return chip

    def _build_whitelist_card(self) -> QFrame:
        """构建「21家银行白名单」卡片（可折叠）。"""
        card = QFrame(self)
        card.setObjectName("card")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(theme.SPACING[16], theme.SPACING[12], theme.SPACING[16], theme.SPACING[12])
        lay.setSpacing(theme.SPACING[8])

        # 可点击表头（标题 + 数量 + 箭头）
        header_row = QHBoxLayout()
        header = QLabel("21家银行白名单", card)
        header.setObjectName("cardTitle")
        header.setCursor(Qt.PointingHandCursor)
        count = QLabel(f"（共 {len(_logic.get_whitelist_banks())} 家）", card)
        count.setObjectName("pageDesc")
        self._wl_chevron = QLabel("▾", card)  # ▾
        self._wl_chevron.setObjectName("pageDesc")
        header_row.addWidget(header)
        header_row.addWidget(count)
        header_row.addStretch(1)
        header_row.addWidget(self._wl_chevron)
        lay.addLayout(header_row)

        # 折叠主体：按类型分组 + 流式标签
        body = QWidget(card)
        body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(0, theme.SPACING[4], 0, 0)
        body_lay.setSpacing(theme.SPACING[8])
        for cat, banks in _logic.get_whitelist_groups():
            cat_label = QLabel(cat, body)
            cat_label.setStyleSheet(
                "background:transparent;"
                f"color:{theme.COLOR['muted']};"
                f"font-weight:600;"
                f"font-size:{theme.TYPE['helper']}px;"
            )
            body_lay.addWidget(cat_label)
            row = QWidget(body)
            flow = FlowLayout(row, hspacing=theme.SPACING[8], vspacing=theme.SPACING[8])
            for b in banks:
                flow.addWidget(self._make_chip(b, row))
            body_lay.addWidget(row)
        lay.addWidget(body)

        self._wl_body = body
        self._wl_collapsed = False

        def _toggle(_event=None) -> None:
            self._wl_collapsed = not self._wl_collapsed
            body.setVisible(not self._wl_collapsed)
            self._wl_chevron.setText("▸" if self._wl_collapsed else "▾")  # ▸ / ▾

        header.mousePressEvent = lambda e: _toggle()
        self._wl_chevron.mousePressEvent = lambda e: _toggle()
        return card

    # ----- 拖拽支持 -------------------------------------------------------

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        """处理拖拽进入事件。"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        """处理放下事件。"""
        mime_data = event.mimeData()
        if mime_data.hasUrls():
            urls = mime_data.urls()
            if urls:
                file_path = urls[0].toLocalFile()
                if file_path:
                    self._load_file_by_path(file_path)
                    event.acceptProposedAction()

    # ----- 交互逻辑 -------------------------------------------------------

    def _ts(self) -> str:
        return datetime.now().strftime("%H:%M:%S")

    def _log_line(self, msg: str) -> None:
        self._log.appendPlainText(f"[{self._ts()}] {msg}")

    def _paste_clipboard(self) -> None:
        cb = QApplication.instance().clipboard()
        text = ""
        if cb is not None:
            try:
                text = cb.text()
            except Exception:
                text = ""

        if not text:
            # Qt 剪贴板可能在某些环境（如 offscreen）下为空或不可用，尝试后备方案
            try:
                text = _logic._read_clipboard()
            except Exception as e:  # noqa: BLE001
                self._log_line(f"❌ 剪贴板读取失败：{e}")
                self._log_line("⚠ 剪贴板为空或不可用。")
                return

        self._txt_input.appendPlainText(text)
        self._log_line("ℹ 已从剪贴板粘贴内容。")

    def _load_file_by_path(self, file_path: str) -> None:
        """通用的文件加载方法（文件对话框和拖拽都用此方法）。"""
        try:
            banks = _logic.load_banks(file_path, None, False)
        except Exception as e:  # noqa: BLE001
            utils.error("加载失败", f"读取文件失败：{e}", parent=self)
            self._log_line(f"❌ 加载失败：{e}")
            return
        self._txt_input.clear()
        self._txt_input.appendPlainText("\n".join(banks))
        self._log_line(f"ℹ 已从文件加载 {len(banks)} 条银行名称。")

    def _load_file(self) -> None:
        p, _ = QFileDialog.getOpenFileName(
            self,
            "选择输入文件",
            str(app_config.DATA_DIR),
            "Excel/CSV/TXT (*.xlsx *.xlsm *.csv *.txt);;All files (*)",
        )
        if not p:
            return
        self._load_file_by_path(p)

    def _clear(self) -> None:
        self._txt_input.clear()
        self._log_line("ℹ 已清空输入。")

    def _process(self) -> None:
        text = self._txt_input.toPlainText().strip()
        banks = [line.strip() for line in text.splitlines() if line.strip()]
        if not banks:
            self._log_line("⚠ 未检测到任何银行名称，请先粘贴或加载。")
            return

        # 点击后让用户自己选择导出位置（类似 notes_receivable_import 的导出）
        out_path, _ = QFileDialog.getSaveFileName(
            self, "导出 Excel",
            str(app_config.DEFAULT_OUTPUT_DIR / "银行承兑汇票分类.xlsx"),
            "Excel 文件 (*.xlsx)",
        )
        if not out_path:
            self._log_line("ℹ 已取消导出。")
            return
        self._out_path = Path(out_path)

        self._btn_run.setEnabled(False)
        self._progress.setText("处理中…")
        self._log.clear()
        self._log_line(f"ℹ 开始处理 {len(banks)} 条银行名称 → {self._out_path}")

        out_path = self._out_path
        # 使用 ThreadPoolExecutor 代替 QThread/moveToThread 模式，避免跨线程 QObject 生命周期问题
        self._future = self._executor.submit(_logic.build_workbook, banks, out_path)

        def _check_future():
            if self._future.done():
                try:
                    res = self._future.result()
                except Exception as e:  # noqa: BLE001
                    self._on_fail(str(e))
                else:
                    self._on_done(banks, res)
                if self._poll_timer is not None:
                    self._poll_timer.stop()
                    self._poll_timer.deleteLater()
                    self._poll_timer = None

        # 轮询 future 状态，回调在主线程中执行
        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(100)
        self._poll_timer.timeout.connect(_check_future)
        self._poll_timer.start()

    # ----- worker 回调 ----------------------------------------------------

    def _on_done(self, banks: list[str], res) -> None:
        n_yes, n_no = res
        self._btn_run.setEnabled(True)
        self._progress.setText("")
        rows = _logic.classify_banks(banks)
        yes_rows = [r for r in rows if r[3] == "21银行承兑汇票"]
        no_rows = [r for r in rows if r[3] == "非21银行承兑汇票"]
        self._log_line(
            f"✅ 完成：总计 {len(banks)} 条 | "
            f"21银行承兑汇票 {n_yes} 条 | 非21银行承兑汇票 {n_no} 条"
        )
        self._log_line("ℹ 分类明细：")
        self._log_line(
            f"  · 21银行承兑汇票（{len(yes_rows)} 家）："
            + ("、".join(r[1] for r in yes_rows) or "（无）")
        )
        self._log_line(
            f"  · 非21银行承兑汇票（{len(no_rows)} 家）："
            + ("、".join(r[1] for r in no_rows) or "（无）")
        )
        self._log_line(f"✅ 已生成：{self._out_path}")
        from PySide6.QtCore import QTimer
        QTimer.singleShot(0, lambda: utils.info(
            "完成",
            f"已生成: {self._out_path}\n"
            f"总计 {len(banks)} 条 | "
            f"21银行承兑汇票 {n_yes} 条 | 非21银行承兑汇票 {n_no} 条",
            parent=self,
        ))

    def _on_fail(self, err: str) -> None:
        self._btn_run.setEnabled(True)
        self._progress.setText("")
        self._log_line(f"❌ 生成失败：{err}")
        from PySide6.QtCore import QTimer
        QTimer.singleShot(0, lambda: utils.error("生成失败", f"生成 Excel 失败：{err}", parent=self))


class BankClassifyFeature(FeatureModule):
    """银行承兑汇票白名单分类 — 功能模块入口。

    主窗口通过 get_widget() 加载此模块的界面，不关心内部实现细节。
    """

    name = "票据分类"
    icon = "fa5s.file-excel"

    def get_widget(self, parent: QWidget | None = None) -> QWidget:
        return BankClassifyWidget(parent)
