#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""notes_receivable_import 数据预览弹窗（headless 验证）。

覆盖：
  1) NotesReceivableImportWidget 构建：预览按钮存在、初始禁用、model 存在
  2) PreviewDialog 构造：共享同一 model、保存/关闭按钮存在且文案正确
  3) 预览「保存」写回：修改 model 后 _on_preview_save 收集到 _table_data

运行：cd 项目根 && QT_QPA_PLATFORM=offscreen ./.venv/Scripts/python.exe tests/test_notes_preview.py
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
# 保证项目根在 sys.path（无论从哪运行本脚本）
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtCore import Qt
from PySide6.QtGui import QStandardItemModel, QStandardItem
from PySide6.QtWidgets import QApplication

from features.notes_receivable_import.notes_receivable_import_ui import (
    NotesReceivableImportWidget,
    PreviewDialog,
)

# QApplication 单例：整个测试进程只创建一个
APP = QApplication.instance() or QApplication([])


def _fill(model, rows):
    """往 model 填几行数据（含 UserRole 原始值，模拟 _populate_table 行为）。"""
    headers = ["*导入序号", "*票据类型", "*付款单位", "*票面金额"]
    model.setColumnCount(len(headers))
    model.setHorizontalHeaderLabels(headers)
    model.setRowCount(len(rows))
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            item = QStandardItem(str(val))
            item.setData(val, Qt.ItemDataRole.UserRole)
            model.setItem(r, c, item)


def test_widget_builds():
    w = NotesReceivableImportWidget()
    assert w._btn_preview is not None
    assert w._model is not None
    assert not w._btn_preview.isEnabled()   # 未导入前禁用
    w.deleteLater()


def test_preview_dialog_shares_model():
    model = QStandardItemModel()
    _fill(model, [[1, "银行承兑汇票", "客户", 100.0],
                  [2, "银行承兑汇票", "客户", 200.0]])
    dlg = PreviewDialog(None, model, 2)
    # 弹窗与主界面共享同一个 model 实例
    assert dlg._view.model() is model
    # 底部按钮存在且文案正确
    assert dlg._btn_save is not None
    assert dlg._btn_close is not None
    assert dlg._btn_save.text() == "保存"
    assert dlg._btn_close.text() == "关闭"
    dlg.deleteLater()


def test_preview_save_writes_back():
    w = NotesReceivableImportWidget()
    _fill(w._model, [[1, "银行承兑汇票", "客户", 100.0]])
    assert w._model.rowCount() == 1
    # 模拟用户在弹窗里编辑（弹窗共享同一 model，直接改 model 即可）
    w._model.item(0, 3).setText("888.0")
    w._model.item(0, 3).setData(888.0, Qt.ItemDataRole.UserRole)
    w._on_preview_save()
    assert len(w._table_data) == 1
    # 编辑后的值（888.0）应被写回；按字符串化比较以兼容 float/str 两种回写形态
    flat = [str(v) for rec in w._table_data for v in rec.values()]
    assert "888.0" in flat
    w.deleteLater()


if __name__ == "__main__":
    test_widget_builds()
    test_preview_dialog_shares_model()
    test_preview_save_writes_back()
    print("ALL PREVIEW TESTS PASSED")
