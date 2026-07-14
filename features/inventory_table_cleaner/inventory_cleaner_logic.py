#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""features/inventory_table_cleaner/inventory_cleaner_logic.py — 存货收发存汇总表清洗逻辑层。

实现方式：通过 win32com 驱动本机 **Microsoft Excel** 完成表头清洗（COM 自动化）。

处理的原始表头为 3 层合并结构（测试文档 `存货收发存汇总表测试.xlsx` 的 `zy` 表）：
  - 第 3 行：分组名（期初结存 / 本期收入 / 本期发出 / 本期结存），左侧 A~K 为竖向合并的单值列
  - 第 4 行：中间层（期初结存/本期结存的 数量/单价/金额；本期收入/本期发出的子类如采购类/生产类…）
  - 第 5 行：本期收入/本期发出的 数量/单价/金额

清洗步骤（与需求一致）：
  1) 删除第 1、2 行（标题/本位币说明）
  2) 取消第 3~5 行合并单元格（先按合并区域把标签填充到每个单元格，再 UnMerge，保留层级信息）
  3) 把分组名与 数量/单价/金额 等逐层结合成完整字段名（如期初结存-数量、本期收入-采购类-数量）
  4) 删除多余的中间/底行，得到单层（或双层）干净表头

为什么用 COM 而非 openpyxl：COM 由 Excel 自身执行删除/取消合并/保存，**完整保留原有
数字格式、列宽、字体等样式**，且不触发 openpyxl 重写带来的样式丢失。

UI 与逻辑严格分离：本文件不依赖任何 Qt 类型；`win32com` 为惰性导入，
缺少 Excel / COM 时模块仍可 import（仅运行时会抛出清晰错误）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

# ===== COM 常量（用字面量，避免依赖 gencache 生成的枚举常量） =====
_XL_FILE_FORMAT_XLSX = 51  # xlOpenXMLWorkbook：.xlsx


def _com_excel_available() -> bool:
    """本机是否具备 COM Excel 自动化所需的 pywin32。"""
    try:
        import win32com  # noqa: F401
        return True
    except Exception:
        return False


def _create_excel_app(visible: bool = False):
    """创建并初始化 Excel.Application 实例。

    必须在调用线程内调用（内部负责 CoInitialize）；返回的应用实例
    由调用方在用完后通过 :func:`_quit_excel_app` 释放。

    Raises:
        RuntimeError: 无法启动 Excel（未安装 / pywin32 缺失 / 非 Windows）。
    """
    import pythoncom
    import win32com.client

    # 同一线程重复 CoInitialize 会抛错，忽略即可
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except pythoncom.error:
        pass

    try:
        # DispatchEx 强制创建独立的新实例，避免复用用户已打开的 Excel
        # 或上一轮刚退出、尚在过渡态的残留实例（复用会导致后续属性设置失败）。
        app = win32com.client.DispatchEx("Excel.Application")
    except Exception as exc:  # noqa: BLE001
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        raise RuntimeError(
            "无法启动 Microsoft Excel（COM 自动化失败）。请确认：\n"
            "  1) 当前为 Windows 且已安装 Microsoft Excel；\n"
            "  2) 已安装 pywin32（pip install pywin32）；\n"
            "  3) Excel 未被其他进程独占锁定。"
        ) from exc

    # 以下属性设置个别环境可能抛 COM 异常（如共享实例状态异常），忽略即可，
    # 不影响后台静默处理。
    try:
        app.Visible = visible
    except Exception:
        pass
    try:
        app.DisplayAlerts = False   # 抑制保存/覆盖等弹窗
    except Exception:
        pass
    try:
        app.ScreenUpdating = False  # 后台静默处理，提速
    except Exception:
        pass
    return app


def _quit_excel_app(app) -> None:
    """退出 Excel 并释放当前线程的 COM 初始化。对 None 安全。"""
    if app is None:
        return
    try:
        app.Quit()
    except Exception:
        pass
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


def _as_text(value) -> str:
    """把 COM 单元格值规整为字符串（None / 数字 / 字符串统一处理）。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


@dataclass
class CleanResult:
    """单次清洗结果。"""
    input_path: Path
    output_path: Path
    header_mode: str
    header_count: int
    headers: list[str]
    data_rows: int
    ok: bool
    error: str | None = None
    log_lines: list[str] = field(default_factory=list)


def process_inventory(
    input_path: str | Path,
    output_path: str | Path,
    header_mode: str = "single",
    sep: str = "-",
    progress_callback: Callable[[str], None] | None = None,
) -> CleanResult:
    """清洗存货收发存汇总表，生成单层（或双层）表头的新 Excel。

    Args:
        input_path: 源 .xlsx 路径。
        output_path: 目标 .xlsx 路径（父目录会被创建；会另存，不覆盖源文件）。
        header_mode: "single" 单层表头（合并成一行，删除原第 4、5 行）；
                     "double" 双层表头（保留分组行 + 合并后的明细行，仅删原第 5 行）。
        sep: 层级结合的分隔符，默认 "-" 。
        progress_callback: 进度日志回调，用于向 UI 回传消息。

    Returns:
        CleanResult 汇总对象。

    Note:
        调用线程需已 CoInitialize（由 :func:`_create_excel_app` 处理）。
    """
    def _log(msg: str) -> None:
        if progress_callback is not None:
            progress_callback(msg)

    if not _com_excel_available():
        raise RuntimeError(
            "本机缺少 pywin32，无法进行 Excel COM 处理。请执行：pip install pywin32"
        )

    in_p = Path(input_path).resolve()
    out_p = Path(output_path).resolve()
    if not in_p.exists():
        raise FileNotFoundError(f"源文件不存在：{in_p}")
    if in_p == out_p:
        raise ValueError("输出文件不能与源文件相同（避免覆盖原文件）")
    out_p.parent.mkdir(parents=True, exist_ok=True)
    if out_p.suffix.lower() != ".xlsx":
        out_p = out_p.with_suffix(".xlsx")

    header_mode = "single" if header_mode != "double" else "double"

    app = _create_excel_app(False)
    try:
        _log(f"打开源文件：{in_p.name}")
        wb = app.Workbooks.Open(str(in_p))
        try:
            # 处理第一个工作表（测试文档为 zy 表）
            ws = wb.Worksheets(1)

            # 最大列（合并/数据都不会超过此范围）
            maxcol = int(ws.UsedRange.Columns.Count)
            if maxcol <= 0:
                maxcol = 1
            maxcol = min(maxcol, 1000)

            # 步骤 1：删除第 1、2 行（上移）
            _log("步骤1：删除第 1、2 行（标题/说明）")
            ws.Rows("1:2").Delete()

            # 此时原表头 3~5 行已上移到 1~3 行
            # 步骤 2：先按合并区域把标签填充到每个单元格（保留层级），再取消合并
            _log("步骤2：读取合并区域并填充层级标签，随后取消第 3~5 行合并")
            filled: dict[tuple[int, int], str] = {}
            for r in (1, 2, 3):
                for c in range(1, maxcol + 1):
                    cell = ws.Cells(r, c)
                    if cell.MergeCells:
                        ma = cell.MergeArea
                        # 只在合并区左上角处理一次，避免重复
                        if ma.Row == r and ma.Column == c:
                            val = _as_text(cell.Value)
                            for rr in range(ma.Row, ma.Row + ma.Rows.Count):
                                for cc in range(ma.Column,
                                                ma.Column + ma.Columns.Count):
                                    if 1 <= rr <= 3 and 1 <= cc <= maxcol:
                                        filled[(rr, cc)] = val
                    else:
                        filled[(r, c)] = _as_text(cell.Value)

            # 取消合并（仅作用于当前 1~3 行区域）
            ws.Range(ws.Cells(1, 1), ws.Cells(3, maxcol)).UnMerge()

            # 把填充后的标签写回每个单元格（取消合并后这些原本空白的单元格需显式赋值）
            for (r, c), v in filled.items():
                ws.Cells(r, c).Value = v

            # 步骤 3：逐列把 1~3 行层级结合成完整字段名
            _log("步骤3：结合分组名与 数量/单价/金额 等，生成完整字段名")
            combined: list[str] = []
            for c in range(1, maxcol + 1):
                parts: list[str] = []
                for r in (1, 2, 3):
                    v = filled.get((r, c), "")
                    if v:
                        parts.append(v)
                # 去除层级冗余：
                #  - 连续完全相同则去重；
                #  - 子级名已包含父级名（如 本期收入 / 本期收入合计）则只保留更具体的，
                #    避免 本期收入-本期收入合计-数量 这类冗余，得到 本期收入合计-数量。
                deduped: list[str] = []
                for p in parts:
                    if not deduped:
                        deduped.append(p)
                    else:
                        last = deduped[-1]
                        if p == last:
                            continue
                        elif p.startswith(last):
                            deduped[-1] = p
                        elif last.startswith(p):
                            pass
                        else:
                            deduped.append(p)
                combined.append(sep.join(deduped) if deduped else "")

            # 步骤 4：写回表头并删除多余行
            if header_mode == "single":
                _log("步骤4：写回单层表头，并删除原第 4、5 行")
                for c in range(1, maxcol + 1):
                    ws.Cells(1, c).Value = combined[c - 1]
                ws.Rows("2:3").Delete()
                header_rows = 1
            else:
                _log("步骤4：写回明细表头（第 2 行），保留分组行（第 1 行），删除原第 5 行")
                for c in range(1, maxcol + 1):
                    ws.Cells(2, c).Value = combined[c - 1]
                ws.Rows(3).Delete()
                header_rows = 2

            total_rows = int(ws.UsedRange.Rows.Count)
            data_rows = max(0, total_rows - header_rows)

            # 保存（另存为 .xlsx，不改动源文件）
            _log(f"保存结果：{out_p.name}")
            wb.SaveAs(str(out_p), FileFormat=_XL_FILE_FORMAT_XLSX)

            headers = [h for h in combined]
            _log(
                f"✅ 完成：共 {len(headers)} 列表头，{data_rows} 行数据。"
            )
            return CleanResult(
                input_path=in_p,
                output_path=out_p,
                header_mode=header_mode,
                header_count=len(headers),
                headers=headers,
                data_rows=data_rows,
                ok=True,
                log_lines=[],
            )
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
    finally:
        _quit_excel_app(app)
