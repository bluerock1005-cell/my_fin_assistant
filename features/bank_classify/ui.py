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

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core import app_config, theme, utils
from core.feature_base import FeatureModule
from core.worker import run_in_thread

# 纯逻辑层：不导入任何 UI 依赖，可独立测试
from . import classify_logic as _logic


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
        root = QVBoxLayout(self)
        root.setContentsMargins(theme.SPACING[24], theme.SPACING[24], theme.SPACING[24], theme.SPACING[24])
        root.setSpacing(theme.SPACING[16])

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

        # 输入卡片
        input_card = QFrame(self)
        input_card.setObjectName("card")
        ic_lay = QVBoxLayout(input_card)
        ic_lay.setContentsMargins(theme.SPACING[16], theme.SPACING[16], theme.SPACING[16], theme.SPACING[16])
        ic_lay.setSpacing(theme.SPACING[12])

        head = QHBoxLayout()
        head_lbl = QLabel("银行全称输入", input_card)
        head_lbl.setObjectName("cardTitle")
        head.addWidget(head_lbl)
        head.addStretch(1)
        btn_load = QPushButton("从文件加载", input_card)
        btn_load.setFixedHeight(32)
        btn_load.clicked.connect(self._load_file)
        btn_paste = QPushButton("从剪贴板粘贴", input_card)
        btn_paste.setFixedHeight(32)
        btn_paste.clicked.connect(self._paste_clipboard)
        head.addWidget(btn_load)
        head.addWidget(btn_paste)
        ic_lay.addLayout(head)

        self._txt_input = QPlainTextEdit(input_card)
        self._txt_input.setObjectName("logView")
        self._txt_input.setPlaceholderText(
            "在此粘贴银行名称，或点击上方按钮从文件加载...\n每行一个银行全称。"
        )
        self._txt_input.setMinimumHeight(140)
        ic_lay.addWidget(self._txt_input)

        root.addWidget(input_card)

        # 输出路径行
        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("输出文件：", self))
        self._lbl_out = QLineEdit(self)
        self._lbl_out.setReadOnly(True)
        self._lbl_out.setText(str(self._out_path))
        out_row.addWidget(self._lbl_out, stretch=1)
        btn_browse = QPushButton("浏览", self)
        btn_browse.setFixedHeight(32)
        btn_browse.clicked.connect(self._browse_out)
        out_row.addWidget(btn_browse)
        root.addLayout(out_row)

        # 主操作
        actions = QHBoxLayout()
        actions.addStretch(1)
        btn_clear = QPushButton("清空", self)
        btn_clear.setFixedHeight(34)
        btn_clear.clicked.connect(self._clear)
        actions.addWidget(btn_clear)
        self._btn_run = QPushButton("生成 Excel", self)
        self._btn_run.setObjectName("primary")
        self._btn_run.setFixedHeight(34)
        self._btn_run.clicked.connect(self._process)
        actions.addWidget(self._btn_run)
        root.addLayout(actions)

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
        log_title = QLabel("运行日志", log_card)
        log_title.setObjectName("cardTitle")
        lc_lay.addWidget(log_title)
        self._log = QPlainTextEdit(log_card)
        self._log.setObjectName("logView")
        self._log.setReadOnly(True)
        self._log.setMinimumHeight(120)
        lc_lay.addWidget(self._log)
        root.addWidget(log_card, stretch=1)

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
                self._log_line(f"剪贴板读取失败：{e}")
                self._log_line("剪贴板为空或不可用。")
                return

        self._txt_input.appendPlainText(text)
        self._log_line("已从剪贴板粘贴内容。")

    def _load_file_by_path(self, file_path: str) -> None:
        """通用的文件加载方法（文件对话框和拖拽都用此方法）。"""
        try:
            banks = _logic.load_banks(file_path, None, False)
        except Exception as e:  # noqa: BLE001
            utils.error("加载失败", f"读取文件失败：{e}", parent=self)
            self._log_line(f"加载失败：{e}")
            return
        self._txt_input.clear()
        self._txt_input.appendPlainText("\n".join(banks))
        self._log_line(f"已从文件加载 {len(banks)} 条银行名称。")

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

    def _browse_out(self) -> None:
        p, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出文件",
            str(self._out_path),
            "Excel 文件 (*.xlsx)",
        )
        if p:
            self._out_path = Path(p)
            self._lbl_out.setText(str(self._out_path))

    def _clear(self) -> None:
        self._txt_input.clear()
        self._log_line("已清空输入。")

    def _process(self) -> None:
        text = self._txt_input.toPlainText().strip()
        banks = [line.strip() for line in text.splitlines() if line.strip()]
        if not banks:
            self._log_line("未检测到任何银行名称，请先粘贴或加载。")
            return

        self._btn_run.setEnabled(False)
        self._progress.setText("处理中…")
        self._log_line(f"开始处理 {len(banks)} 条银行名称。")

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
        self._log_line(
            f"完成：总计 {len(banks)} 条 | "
            f"21银行承兑汇票 {n_yes} 条 | 非21银行承兑汇票 {n_no} 条"
        )
        self._log_line(f"已生成：{self._out_path}")
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
        self._log_line(f"生成失败：{err}")
        from PySide6.QtCore import QTimer
        QTimer.singleShot(0, lambda: utils.error("生成失败", f"生成 Excel 失败：{err}", parent=self))


class BankClassifyFeature(FeatureModule):
    """银行承兑汇票白名单分类 — 功能模块入口。

    主窗口通过 get_widget() 加载此模块的界面，不关心内部实现细节。
    """

    name = "银行分类"
    icon = "fa5s.file-excel"

    def get_widget(self, parent: QWidget | None = None) -> QWidget:
        return BankClassifyWidget(parent)
