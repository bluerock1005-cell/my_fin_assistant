#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""features/js_bank_statement/logic.py — 江苏银行对账单读取逻辑。

仅包含纯数据读取函数，不依赖任何 UI，方便单测。
支持 .xlsx/.xlsm（openpyxl）和旧 .xls（xlrd 回退）。

保留原始数据类型（日期、数字等），便于后续处理显示和复制。
"""
from __future__ import annotations

from pathlib import Path
from typing import Tuple, List, Any
from datetime import datetime, date

import openpyxl


def _cell_to_display_str(v: Any) -> str:
    """将单元格值转换为显示用的字符串，保留原始类型检测。"""
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        # 日期格式：YYYY-MM-DD（可选时间部分）
        if isinstance(v, datetime) and v.time() != datetime.min.time():
            return v.strftime("%Y-%m-%d %H:%M:%S")
        return v.strftime("%Y-%m-%d")
    return str(v)


def _read_xls_with_xlrd(path: Path) -> List[List[Any]]:
    try:
        import xlrd
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(
            "文件是旧的 .xls 格式，openpyxl 无法读取。请安装 xlrd (pip install xlrd) 或将文件另存为 .xlsx 后重试。"
        ) from e

    book = xlrd.open_workbook(path, formatting_info=False)
    sheet = book.sheet_by_index(0)
    rows: List[List[Any]] = []
    for r in range(sheet.nrows):
        row: List[Any] = []
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            # xlrd 返回的日期值通常是浮点数，需要判断列类型
            # 但基于 xlrd 的限制，先简单保存原值
            row.append(val)
        rows.append(row)
    return rows


def load_statement(path: str) -> Tuple[List[str], List[List[Any]]]:
    """读取 Excel 文件，返回 (headers, rows)。

    - path: 文件路径（推荐 .xlsx/.xlsm；若为 .xls 会尝试使用 xlrd 回退）
    - headers: 第 3 行作为列标题列表
    - rows: 第 4 行开始的数据行（每行保留原始值：datetime 对象、数字等）
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"输入文件不存在: {p}")

    suffix = p.suffix.lower()
    raw_rows: List[List[Any]] = []

    def _process_openpyxl(pathp: Path) -> List[List[Any]]:
        wb = openpyxl.load_workbook(pathp, read_only=True, data_only=True)
        ws = wb.active
        # 保留原始数据类型（datetime、float 等），不转换为字符串
        return [list(r) for r in ws.iter_rows(values_only=True)]

    # 根据扩展名选择解析器，未知扩展名时优先尝试 openpyxl，再回退 xlrd
    if suffix in (".xlsx", ".xlsm"):
        raw_rows = _process_openpyxl(p)
    elif suffix == ".xls":
        raw_rows = _read_xls_with_xlrd(p)
    else:
        # 未知扩展：优先尝试 openpyxl（多数情况为 .xlsx），若报旧 .xls 错误再尝试 xlrd
        try:
            raw_rows = _process_openpyxl(p)
        except Exception as e:  # noqa: BLE001
            msg = str(e).lower()
            if "old .xls" in msg or "xlrd" in msg or "unsupported" in msg:
                # 可能是旧 .xls，尝试 xlrd
                raw_rows = _read_xls_with_xlrd(p)
            else:
                raise

    if len(raw_rows) < 3:
        # 文件行数不足 3 行，无法读取第 3 行表头
        return [], []

    # 以第 3 行（索引 2）为表头，跳过前两行
    headers = [_cell_to_display_str(h) if h is not None else "" for h in raw_rows[2]]
    data = raw_rows[3:]

    # 规范化：确保每行长度与 headers 一致
    max_cols = max(len(headers), max((len(r) for r in data), default=0))
    if len(headers) < max_cols:
        headers += [f"列{idx+1}" for idx in range(len(headers), max_cols)]
    norm_rows: List[List[Any]] = []
    for r in data:
        row = [("" if i >= len(r) or r[i] is None else r[i]) for i in range(max_cols)]
        norm_rows.append(row)
    return headers, norm_rows
