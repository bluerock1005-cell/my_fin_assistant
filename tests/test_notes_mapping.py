#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""notes_receivable_import 映射确认环节（headless 验证）。

覆盖：
  1) import_logic：手工覆盖(manual_maps) + 显式不匹配(excluded_maps)
  2) MappingDialog：重复来源列冲突提示 / 必录未匹配提示 / 恢复自动匹配 / 结果结构
  3) 持久化：手工映射保存到 data/mapping_overrides.json 并能回载

运行：cd 项目根 && QT_QPA_PLATFORM=offscreen ./.venv/Scripts/python.exe tests/test_notes_mapping.py
"""
import os
import sys
import json
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
# 保证项目根在 sys.path（无论从哪运行本脚本）
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from pathlib import Path
from PySide6.QtWidgets import QApplication

from features.notes_receivable_import import import_logic as L
from features.notes_receivable_import.ui import (
    NotesReceivableImportWidget,
    MappingDialog,
    _OVERRIDES_FILE,
)

# QApplication 单例：整个测试进程只创建一个
APP = QApplication.instance() or QApplication([])

_HEADERS = ["序号", "票据（包）号码", "出票日", "到期日期", "面额",
            "出票人全称", "承兑银行名称", "收票日期", "备注"]


def _make_src(tmp: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(_HEADERS)
    ws.append([1, "PB2026001", "2026-04-01", "2026-10-01", "12345.6",
               "甲方有限公司", "宁波银行", "2026-07-08", "x"])
    ws.append([2, "PB2026002", "2026-05-01", "2026-11-01", "678.9",
               "乙方有限公司", "工商银行", "2026-07-09", "y"])
    wb.save(tmp)
    return tmp


def test_logic_manual_and_excluded() -> None:
    tmp = Path(tempfile.gettempdir()) / "notes_synth.xlsx"
    _make_src(tmp)
    try:
        manual = {tmp.name: {
            "序号": "*导入序号", "票据（包）号码": "*票据号",
            "出票日": "*签发日", "到期日期": "*到期日",
            "面额": "*票面金额", "出票人全称": "*出票人",
            "承兑银行名称": "*承兑人", "收票日期": "*收票日"}}
        excluded = {tmp.name: ["*币别", "*汇率"]}
        defaults = {"*币别": "人民币", "*汇率": 1.0,
                    "*票据类型": "银行承兑汇票", "*付款单位": "客户"}
        r = L.read_and_transform([tmp], defaults, manual, excluded)
        rec = r.table_data[0]
        assert rec["*票据号"] == "PB2026001", rec
        assert str(rec["*签发日"]).startswith("2026-04-01"), rec
        assert rec["*票面金额"] == 12345.6, rec
        assert rec["*出票人"] == "甲方有限公司", rec
        assert rec["*承兑人"] == "宁波银行", rec
        # excluded 字段不应被自动抢到来源列，而是用默认值
        auto_targets = {m.target_field for m in r.files[0].matched_columns}
        assert "*币别" not in auto_targets
        assert "*汇率" not in auto_targets
        assert rec["*币别"] == "人民币"
        assert rec["*汇率"] == 1.0
        assert len(r.table_data) == 2
    finally:
        tmp.unlink(missing_ok=True)
    print("[OK] test_logic_manual_and_excluded")


def test_dialog_conflict_and_restore() -> None:
    tmp = Path(tempfile.gettempdir()) / "notes_synth2.xlsx"
    _make_src(tmp)
    try:
        pure_auto, _u, _m = L.match_columns(_HEADERS, None)
        auto_map = {m.target_field: m.source_header for m in pure_auto}
        files_info = [{
            "name": tmp.name,
            "source_headers": _HEADERS,
            "auto_map": auto_map,
            "current_map": {},
        }]
        dlg = MappingDialog(None, files_info, L.TEMPLATE_FIELDS)
        # 固定字段（*导入序号/*票据类型/*币别/*汇率/*背书明细）不在对话框中
        assert len(dlg._row_widgets) == len(L.TEMPLATE_FIELDS) - len(L.FIXED_FIELD_HEADERS)

        # 把两个不同字段都选成同一来源列 → 重复冲突（用非固定字段）
        src = "票据（包）号码"
        for tgt in ("*票据号", "*出票人"):
            for tf, _l, c in dlg._row_widgets:
                if tf.header == tgt:
                    c.setCurrentIndex(c.findData(src))
        dlg._validate()
        assert dlg._warn.isVisible() or dlg._warn.text(), "冲突提示应出现"
        assert "被重复选择" in dlg._warn.text(), dlg._warn.text()

        # 显式不匹配一个必录字段 → 必录提示（*付款单位 非固定，可触发）
        for tf, _l, c in dlg._row_widgets:
            if tf.header == "*付款单位":
                c.setCurrentIndex(c.findData(""))
        dlg._validate()
        assert "必录字段" in dlg._warn.text(), dlg._warn.text()

        # 恢复自动匹配后，冲突应消失
        dlg._restore_auto()
        dlg._validate()
        assert "被重复选择" not in dlg._warn.text(), dlg._warn.text()

        # 结果结构：每个文件返回 {tgt: src_or_None}（固定字段不出现）
        dlg.accept()
        rm = dlg.result_maps
        assert tmp.name in rm
        assert isinstance(rm[tmp.name], dict)
        assert "*票据类型" not in rm[tmp.name]   # 固定字段不应在结果中
        assert rm[tmp.name].get("*付款单位") is None  # 被恢复成自动（未匹配）
    finally:
        tmp.unlink(missing_ok=True)
    print("[OK] test_dialog_conflict_and_restore")


def test_persistence_roundtrip() -> None:
    _OVERRIDES_FILE.unlink(missing_ok=True)
    try:
        w = NotesReceivableImportWidget()
        w._manual_full = {
            "应收票据测试1.xlsx": {
                "*票据号": "票据（包）号码",
                "*收票日": None,           # 显式不匹配（非固定字段）
                "*付款单位": "客户",
            }
        }
        w._save_overrides()
        assert _OVERRIDES_FILE.exists()

        # 重新载入（模拟下次启动）
        w2 = NotesReceivableImportWidget()
        loaded = w2._manual_full.get("应收票据测试1.xlsx", {})
        assert loaded.get("*票据号") == "票据（包）号码"
        assert loaded.get("*收票日") is None
        assert loaded.get("*付款单位") == "客户"
        # 固定字段不应出现在保存/回载的映射中
        assert "*币别" not in loaded and "*票据类型" not in loaded
    finally:
        _OVERRIDES_FILE.unlink(missing_ok=True)
    print("[OK] test_persistence_roundtrip")


if __name__ == "__main__":
    test_logic_manual_and_excluded()
    test_dialog_conflict_and_restore()
    test_persistence_roundtrip()
    print("\nALL TESTS PASSED")
