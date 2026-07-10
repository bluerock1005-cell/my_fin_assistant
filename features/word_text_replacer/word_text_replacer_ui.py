#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""features/word_text_replacer/word_text_replacer_ui.py — Word 批量文本替换界面。

UI 与逻辑严格分离：本文件只管界面与交互，文件扫描/替换/保存放在
word_text_replacer_logic.py 中。
"""
from __future__ import annotations

import concurrent.futures
from datetime import datetime
from pathlib import Path
from queue import Queue

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core import app_config, theme, utils
from core.feature_base import FeatureModule

from . import word_text_replacer_logic as _logic


class WordTextReplacerWidget(QWidget):
    """Word 批量文本替换功能的具体 UI 控件。

    工作流：
      1. 选择输入文件夹、输出文件夹
      2. 填写替换规则（左侧查找文字，右侧替换文字）
      3. 点击「开始替换」→ 后台扫描并处理所有 .docx
      4. 在日志面板查看进度与汇总结果
    """

    # 默认规则行数
    _DEFAULT_RULE_ROWS = 5

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("page")

        # 状态
        self._input_folder: Path | None = None
        self._output_folder: Path | None = None

        # 后台任务
        self._executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        self._future = None
        self._poll_timer: QTimer | None = None
        self._log_queue: Queue[str] = Queue()

        self._setup_ui()

    # ====== 布局 ======

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD)
        root.setSpacing(theme.CARD_GAP)

        # --- 页头 ---
        title = QLabel("Word 批量文本替换", self)
        title.setObjectName("pageTitle")
        desc = QLabel(
            "选择输入/输出文件夹，填写批量替换规则，自动处理所有 .docx 文件。"
            "原文件不会被修改，结果写入输出文件夹。",
            self,
        )
        desc.setObjectName("pageDesc")
        desc.setWordWrap(True)
        root.addWidget(title)
        root.addWidget(desc)
        root.addSpacing(theme.SPACING[4])

        # === 卡片1：文件夹选择 ===
        folder_card = QFrame(self)
        folder_card.setObjectName("card")
        fc = QVBoxLayout(folder_card)
        fc.setContentsMargins(theme.SPACING[16], theme.SPACING[16],
                              theme.SPACING[16], theme.SPACING[16])
        fc.setSpacing(theme.SPACING[12])

        # 标题行：文件夹选择 + 包含子文件夹（右侧）
        fc_head_row = QHBoxLayout()
        fc_head = QLabel("文件夹选择", folder_card)
        fc_head.setObjectName("cardTitle")
        fc_head_row.addWidget(fc_head)
        fc_head_row.addStretch(1)
        self._chk_subfolders = QCheckBox("包含子文件夹", folder_card)
        self._chk_subfolders.setChecked(False)
        fc_head_row.addWidget(self._chk_subfolders)
        fc.addLayout(fc_head_row)

        grid = QGridLayout()
        grid.setHorizontalSpacing(theme.SPACING[12])
        grid.setVerticalSpacing(theme.SPACING[12])
        grid.setColumnStretch(1, 1)

        # 输入文件夹
        lbl_in = QLabel("输入文件夹:", folder_card)
        lbl_in.setFixedWidth(80)
        grid.addWidget(lbl_in, 0, 0)
        self._edit_input = QLineEdit(folder_card)
        self._edit_input.setReadOnly(True)
        self._edit_input.setPlaceholderText("请选择包含 .docx 的文件夹")
        self._edit_input.setFixedHeight(30)
        grid.addWidget(self._edit_input, 0, 1)
        btn_in = QPushButton("浏览…", folder_card)
        btn_in.setFixedHeight(30)
        btn_in.setMinimumWidth(72)
        btn_in.clicked.connect(self._choose_input_folder)
        grid.addWidget(btn_in, 0, 2)

        # 输出文件夹
        lbl_out = QLabel("输出文件夹:", folder_card)
        lbl_out.setFixedWidth(80)
        grid.addWidget(lbl_out, 1, 0)
        self._edit_output = QLineEdit(folder_card)
        self._edit_output.setReadOnly(True)
        self._edit_output.setPlaceholderText("请选择保存结果的文件夹")
        self._edit_output.setFixedHeight(30)
        grid.addWidget(self._edit_output, 1, 1)
        btn_out = QPushButton("浏览…", folder_card)
        btn_out.setFixedHeight(30)
        btn_out.setMinimumWidth(72)
        btn_out.clicked.connect(self._choose_output_folder)
        grid.addWidget(btn_out, 1, 2)

        fc.addLayout(grid)

        root.addWidget(folder_card)

        # === 卡片2：替换规则 ===
        rule_card = QFrame(self)
        rule_card.setObjectName("card")
        rc = QVBoxLayout(rule_card)
        rc.setContentsMargins(theme.SPACING[16], theme.SPACING[16],
                              theme.SPACING[16], theme.SPACING[16])
        rc.setSpacing(theme.SPACING[12])

        # 标题行：替换规则 + 说明 + 添加一行 + 开始替换（最右）
        title_row = QHBoxLayout()
        title_row.setSpacing(theme.SPACING[12])

        rc_head = QLabel("替换规则", rule_card)
        rc_head.setObjectName("cardTitle")
        title_row.addWidget(rc_head)

        rc_desc = QLabel(
            "左侧填写要被替换的文字，右侧填写替换后的文字。空行会被忽略。",
            rule_card)
        rc_desc.setObjectName("pageDesc")
        rc_desc.setWordWrap(True)
        title_row.addWidget(rc_desc, stretch=1)

        btn_add = QPushButton("+ 添加一行", rule_card)
        btn_add.setFixedHeight(32)
        btn_add.setMinimumWidth(120)
        btn_add.clicked.connect(self._add_rule_row)
        title_row.addWidget(btn_add)

        self._btn_run = QPushButton("开始替换", rule_card)
        self._btn_run.setObjectName("primary")
        self._btn_run.setFixedHeight(38)
        self._btn_run.setMinimumWidth(140)
        self._btn_run.clicked.connect(self._start_replace)
        title_row.addWidget(self._btn_run)

        rc.addLayout(title_row)

        # 滚动区：规则行
        scroll = QScrollArea(rule_card)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._rules_container = QWidget(scroll)
        self._rules_layout = QVBoxLayout(self._rules_container)
        self._rules_layout.setContentsMargins(0, 0, 0, 0)
        self._rules_layout.setSpacing(theme.SPACING[12])
        self._rules_layout.addStretch(1)
        scroll.setWidget(self._rules_container)
        rc.addWidget(scroll, stretch=1)

        # 预置默认行数
        for _ in range(self._DEFAULT_RULE_ROWS):
            self._add_rule_row()

        root.addWidget(rule_card, stretch=10)

        # === 卡片3：运行日志 ===
        log_card = QFrame(self)
        log_card.setObjectName("card")
        lc = QVBoxLayout(log_card)
        lc.setContentsMargins(theme.SPACING[16], theme.SPACING[16],
                              theme.SPACING[16], theme.SPACING[16])
        lc.setSpacing(theme.SPACING[12])

        log_title = QLabel("运行日志", log_card)
        log_title.setObjectName("cardTitle")
        lc.addWidget(log_title)

        self._log = QPlainTextEdit(log_card)
        self._log.setObjectName("logView")
        self._log.setReadOnly(True)
        self._log.setMinimumHeight(80)
        lc.addWidget(self._log, stretch=1)

        root.addWidget(log_card, stretch=1)

    # ====== 规则行管理 ======

    def _add_rule_row(self) -> None:
        """在规则列表末尾添加一行「查找/替换」输入框。"""
        row = QWidget(self._rules_container)
        hbox = QHBoxLayout(row)
        hbox.setContentsMargins(0, 0, 0, 0)
        hbox.setSpacing(theme.SPACING[12])

        edit_find = QLineEdit(row)
        edit_find.setPlaceholderText("被替换的文字")
        edit_find.setFixedHeight(30)
        edit_replace = QLineEdit(row)
        edit_replace.setPlaceholderText("替换后的文字")
        edit_replace.setFixedHeight(30)

        btn_del = QPushButton("删除", row)
        btn_del.setFixedHeight(30)
        btn_del.setMinimumWidth(60)
        btn_del.clicked.connect(lambda _checked, r=row: self._remove_rule_row(r))

        hbox.addWidget(edit_find, stretch=1)
        hbox.addWidget(edit_replace, stretch=1)
        hbox.addWidget(btn_del)

        # 插入到 stretch 之前
        self._rules_layout.insertWidget(self._rules_layout.count() - 1, row)

    def _remove_rule_row(self, row: QWidget) -> None:
        """删除指定规则行。"""
        row.deleteLater()
        self._rules_layout.removeWidget(row)

    def _collect_rules(self) -> list[tuple[str, str]]:
        """收集所有非空的替换规则。"""
        rules: list[tuple[str, str]] = []
        for i in range(self._rules_layout.count()):
            item = self._rules_layout.itemAt(i)
            if item is None:
                continue
            row = item.widget()
            if not isinstance(row, QWidget):
                continue
            edits = row.findChildren(QLineEdit)
            if len(edits) < 2:
                continue
            old = edits[0].text()
            new = edits[1].text()
            if old.strip():
                rules.append((old, new))
        return rules

    # ====== 文件夹选择 ======

    def _choose_input_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "选择输入文件夹", str(app_config.DATA_DIR))
        if path:
            self._input_folder = Path(path)
            self._edit_input.setText(path)
            self._log_line(f"已选择输入文件夹：{path}")
            # 若输出文件夹未选，默认建议为 input 同级目录下的 output
            if self._output_folder is None:
                suggested = self._input_folder.parent / f"{self._input_folder.name}_replaced"
                self._output_folder = suggested
                self._edit_output.setText(str(suggested))
                self._log_line(f"默认输出文件夹：{suggested}")

    def _choose_output_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "选择输出文件夹", str(app_config.DATA_DIR))
        if path:
            self._output_folder = Path(path)
            self._edit_output.setText(path)
            self._log_line(f"已选择输出文件夹：{path}")

    # ====== 日志 ======

    def _ts(self) -> str:
        return datetime.now().strftime("%H:%M:%S")

    def _log_line(self, msg: str) -> None:
        self._log.appendPlainText(f"[{self._ts()}] {msg}")

    # ====== 开始替换 ======

    def _start_replace(self) -> None:
        # 校验
        if self._input_folder is None or self._output_folder is None:
            utils.warning("提示", "请先选择输入文件夹和输出文件夹。", parent=self)
            return
        if self._input_folder == self._output_folder:
            utils.warning("提示", "输入文件夹和输出文件夹不能相同，避免覆盖原文件。", parent=self)
            return

        rules = self._collect_rules()
        if not rules:
            utils.warning("提示", "请至少填写一条替换规则。", parent=self)
            return

        include_subfolders = self._chk_subfolders.isChecked()

        self._btn_run.setEnabled(False)
        self._log.clear()
        self._log_line(f"开始批量替换，共 {len(rules)} 条规则…")
        for i, (old, new) in enumerate(rules, start=1):
            self._log_line(f"  规则 {i}：「{old}」 → 「{new}」")

        self._log_queue = Queue()

        self._future = self._executor.submit(
            _logic.process_folder,
            self._input_folder,
            self._output_folder,
            rules,
            include_subfolders,
            self._log_queue.put,
        )

        def _check() -> None:
            # 先刷新日志队列
            while not self._log_queue.empty():
                try:
                    self._log_line(self._log_queue.get_nowait())
                except Exception:
                    break

            if self._future.done():
                try:
                    result = self._future.result()
                    self._on_done(result)
                except Exception as e:  # noqa: BLE001
                    self._on_fail(str(e))
                if self._poll_timer is not None:
                    self._poll_timer.stop()
                    self._poll_timer.deleteLater()
                    self._poll_timer = None

        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(100)
        self._poll_timer.timeout.connect(_check)
        self._poll_timer.start()

    def _on_done(self, result: _logic.BatchResult) -> None:
        # 清空剩余日志
        while not self._log_queue.empty():
            try:
                self._log_line(self._log_queue.get_nowait())
            except Exception:
                break

        self._btn_run.setEnabled(True)
        self._log_line("-" * 40)
        self._log_line(
            f"✅ 处理完成：{result.success_count}/{result.total_files} 个文件成功，"
            f"共替换 {result.total_replacements} 处文本。"
        )
        if result.fail_count:
            self._log_line(f"⚠ {result.fail_count} 个文件处理失败，请检查上方日志。")
            utils.warning(
                "处理完成（部分失败）",
                f"成功 {result.success_count} 个，失败 {result.fail_count} 个。\n"
                f"共替换 {result.total_replacements} 处文本。",
                parent=self,
            )
        else:
            utils.info(
                "处理完成",
                f"成功处理 {result.success_count} 个文件，\n"
                f"共替换 {result.total_replacements} 处文本。",
                parent=self,
            )

    def _on_fail(self, err: str) -> None:
        while not self._log_queue.empty():
            try:
                self._log_line(self._log_queue.get_nowait())
            except Exception:
                break
        self._btn_run.setEnabled(True)
        self._log_line(f"❌ 处理失败：{err}")
        utils.error("处理失败", f"批量替换时出错：{err}", parent=self)


class WordTextReplacerFeature(FeatureModule):
    """Word 批量文本替换 — 功能模块入口。"""

    name = "Word文本替换"
    icon = "fa5s.file-word"

    def get_widget(self, parent: QWidget | None = None) -> QWidget:
        return WordTextReplacerWidget(parent)
