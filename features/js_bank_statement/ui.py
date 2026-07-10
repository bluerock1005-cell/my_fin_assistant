#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""features/js_bank_statement/ui.py — 江苏银行对账单复制界面。

遵循项目约定：UI 与逻辑分离。UI 只负责交互、展示和剪贴板复制，
读取 Excel 的纯逻辑在同目录下的 `logic.py`。
"""
from __future__ import annotations

from pathlib import Path
from datetime import datetime, date

from PySide6.QtCore import Qt
from PySide6.QtGui import (
    QKeyEvent,
    QKeySequence,
    QStandardItem,
    QStandardItemModel,
)
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core.feature_base import FeatureModule
from core import app_config, theme, utils

# 纯逻辑：不依赖 UI
from . import logic as _logic


class _TableView(QTableView):
    """自定义 QTableView，重写 Ctrl+C 行为以支持多单元格选区复制。"""

    def __init__(self, copy_callback, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._copy_callback = copy_callback

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event == QKeySequence.Copy:
            self._copy_callback()
            return
        super().keyPressEvent(event)


class JsBankStmtWidget(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("page")
        self._path: Path | None = None
        self.setAcceptDrops(True)  # 启用拖拽
        self._setup_ui()

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD)
        root.setSpacing(theme.CARD_GAP)

        title = QLabel("江苏银行对账单复制", self)
        title.setObjectName("pageTitle")
        desc = QLabel("导入江苏银行对账单（Excel），在表格中查看、排序并复制到剪贴板。", self)
        desc.setObjectName("pageDesc")
        desc.setWordWrap(True)
        root.addWidget(title)
        root.addWidget(desc)

        card = QFrame(self)
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(theme.SPACING[12], theme.SPACING[12], theme.SPACING[12], theme.SPACING[12])
        cl.setSpacing(theme.SPACING[8])

        head = QHBoxLayout()
        head.addStretch(1)
        btn_load = QPushButton("加载 Excel", self)
        btn_load.setFixedHeight(32)
        btn_load.clicked.connect(self._load_file)
        head.addWidget(btn_load)
        btn_copy_all = QPushButton("复制全部", self)
        btn_copy_all.setFixedHeight(32)
        btn_copy_all.clicked.connect(self._copy_all)
        head.addWidget(btn_copy_all)
        btn_copy_sel = QPushButton("复制选中", self)
        btn_copy_sel.setFixedHeight(32)
        btn_copy_sel.clicked.connect(self._copy_selection)
        head.addWidget(btn_copy_sel)
        cl.addLayout(head)

        self._table = _TableView(self._copy_selection, card)
        self._table.setSortingEnabled(True)
        cl.addWidget(self._table)

        root.addWidget(card)

        self._model = QStandardItemModel(self)
        self._table.setModel(self._model)

    def _load_file(self) -> None:
        p, _ = QFileDialog.getOpenFileName(
            self,
            "选择江苏银行对账单（Excel）",
            str(app_config.DATA_DIR),
            "Excel 文件 (*.xlsx *.xlsm);;All files (*)",
        )
        if not p:
            return
        try:
            headers, rows = _logic.load_statement(p)
        except Exception as e:  # noqa: BLE001
            utils.error("加载失败", f"读取文件失败：{e}", parent=self)
            return
        self._path = Path(p)
        self._populate_table(headers, rows)
        utils.info("加载完成", f"已加载: {self._path}\n共 {len(rows)} 行", parent=self)

    def _populate_table(self, headers: list[str], rows: list[list]) -> None:
        self._model.clear()
        self._model.setColumnCount(len(headers))
        self._model.setRowCount(len(rows))
        self._model.setHorizontalHeaderLabels(headers)
        # 在每一列存入表头文本到 headerData (QStandardItem header 已由 setHorizontalHeaderLabels 设置)
        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                display_str = _logic._cell_to_display_str(v)
                item = QStandardItem(display_str)
                item.setEditable(False)
                # 将原始值存在 UserRole 上，排序后不丢失
                item.setData(v, Qt.ItemDataRole.UserRole)
                self._model.setItem(r, c, item)
        self._table.resizeColumnsToContents()
        # 第一列宽度固定为 50
        if len(headers) > 0:
            self._table.setColumnWidth(0, 50)

    def _copy_all(self) -> None:
        if self._model.rowCount() == 0:
            return
        text_lines = []
        cols = self._model.columnCount()
        headers = [self._model.horizontalHeaderItem(i).text() or "" for i in range(cols)]
        text_lines.append("\t".join(headers))

        for r in range(self._model.rowCount()):
            cells = []
            for c in range(cols):
                item = self._model.item(r, c)
                raw = item.data(Qt.ItemDataRole.UserRole) if item else ""
                # 日期值格式化
                if isinstance(raw, (datetime, date)):
                    cells.append(raw.strftime("%Y-%m-%d"))
                else:
                    cells.append(_logic._cell_to_display_str(raw))
            text_lines.append("\t".join(cells))

        cb = QApplication.instance().clipboard()
        cb.setText("\r\n".join(text_lines))
        utils.info("已复制", f"已复制 {self._model.rowCount()} 行到剪贴板。", parent=self)

    def _copy_selection(self) -> None:
        sel = self._table.selectionModel().selectedIndexes()
        if not sel:
            return

        # 找出选区的矩形范围
        rows = [idx.row() for idx in sel]
        cols = [idx.column() for idx in sel]
        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(cols), max(cols)

        # 将选中的单元格索引转换为集合，便于查询
        selected_set = {(idx.row(), idx.column()) for idx in sel}

        # 生成矩形区域的所有单元格（从模型中读取，排序后依然正确）
        lines = []
        for r in range(min_row, max_row + 1):
            row_cells = []
            for c in range(min_col, max_col + 1):
                if (r, c) in selected_set:
                    item = self._model.item(r, c)
                    raw = item.data(Qt.ItemDataRole.UserRole) if item else ""
                    if isinstance(raw, (datetime, date)):
                        row_cells.append(raw.strftime("%Y-%m-%d"))
                    else:
                        row_cells.append(_logic._cell_to_display_str(raw))
                else:
                    row_cells.append("")
            lines.append("\t".join(row_cells))

        cb = QApplication.instance().clipboard()
        cb.setText("\r\n".join(lines))
        rows_count = max_row - min_row + 1
        cols_count = max_col - min_col + 1
        utils.info("已复制", f"已复制 {rows_count} 行 × {cols_count} 列到剪贴板。", parent=self)

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
                    try:
                        headers, rows = _logic.load_statement(file_path)
                    except Exception as e:  # noqa: BLE001
                        utils.error("加载失败", f"读取文件失败：{e}", parent=self)
                        return
                    self._path = Path(file_path)
                    self._populate_table(headers, rows)
                    utils.info("加载完成", f"已加载: {self._path}\n共 {len(rows)} 行", parent=self)
                    event.acceptProposedAction()


class JsBankStmtFeature(FeatureModule):
    name = "江苏银行对账单复制"
    icon = "fa5s.file-excel"

    def get_widget(self, parent: QWidget | None = None) -> QWidget:
        return JsBankStmtWidget(parent)
