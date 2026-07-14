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
  5) 删除数据区的「合计/总计」汇总行（通常为数据区最后一行）
  6) 若指定「保留期初/期末会计期间」，清空其余期间对应的期初结存/期末结存数据单元格
     （数量/单价/金额），**保留列不动**——只清数据、不删列

为什么用 COM 而非 openpyxl：COM 由 Excel 自身执行删除/取消合并/保存，**完整保留原有
数字格式、列宽、字体等样式**，且不触发 openpyxl 重写带来的样式丢失。

UI 与逻辑严格分离：本文件不依赖任何 Qt 类型；`win32com` 为惰性导入，
缺少 Excel / COM 时模块仍可 import（仅运行时会抛出清晰错误）。
"""
from __future__ import annotations

import re
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


def _detect_period_range(title_text: str) -> tuple[int, int]:
    """从标题行解析会计期间范围（如 '会计期间：2026年01期 - 2026年06期' → (1, 6)）。

    仅匹配紧跟「期」字的数字（如 '01期'），自动跳过年份 '2026' 之类。
    """
    matches = re.findall(r"(\d+)\s*期", title_text or "")
    nums: list[int] = []
    for m in matches:
        try:
            nums.append(int(m))
        except ValueError:
            pass
    if nums:
        return min(nums), max(nums)
    return 1, 1


def _classify_header(hdr: str) -> tuple[str, int | None]:
    """把合并后的表头归类，并返回其会计期间（若能识别）。

    返回 (group, period)：
      - ("open",  p) 期初结存-*   （累计表无期号时 period=None，视作起始期）
      - ("close", p) 期末结存-*
      - ("income",p) 本期收入-*
      - ("issue", p) 本期发出-*
      - ("fixed",None) 其余固定列（会计年度/物料编号…）
    """
    h = hdr or ""
    # 期号识别：优先「X期」；其次列名中首个数字串（多期间明细文件如「期初结存-02-数量」
    # 不含「期」字，但仍需识别期间）。仅对期初/期末余额组生效，income/issue/fixed 组
    # 不参与按期间清空，是否提取期号无副作用。
    m = re.search(r"(\d+)\s*期", h)
    if not m:
        m = re.search(r"(\d{1,2})", h)
    period = int(m.group(1)) if m else None
    if h.startswith("期初结存"):
        return "open", period
    # 期末结存 / 本期结存 都视作「期末（本期）结存」余额组
    if h.startswith("期末结存") or h.startswith("本期结存"):
        return "close", period
    if h.startswith("本期收入"):
        return "income", period
    if h.startswith("本期发出"):
        return "issue", period
    return "fixed", None


def _clear_non_kept_period_cells(ws, combined, title_text, keep_open, keep_close,
                                 header_rows, log) -> int:
    """按保留的期初/期末会计期间，清空其余期间余额数据单元格（列保留不动）。

    与早期「删列」实现不同，本函数**不删除任何列**，仅清空数据行中属于非保留期间的
    期初结存 / 期末结存（数量/单价/金额）单元格，避免破坏表格列结构。

    规则（与需求一致）：
      - 保留「保留期初会计期间」对应的期初结存 3 列数据，清空其他期间的期初结存数据；
      - 保留「保留期末会计期间」对应的期末结存 3 列数据，清空其他期间的期末结存数据；
      - 期初/期末列中无期号者，视作标题会计期间范围的起始期（期初）/期末期（期末）；
      - 本期收入 / 本期发出 及固定列不参与，整列保留（列不动）。
    返回被清空数据的列数。
    """
    p_start, p_end = _detect_period_range(title_text)
    log(f"检测到会计期间范围：{p_start}期 ~ {p_end}期")

    total_cols = len(combined)
    last_row = int(ws.UsedRange.Rows.Count)
    first_data_row = header_rows + 1
    if last_row < first_data_row:
        return 0

    to_clear: list[int] = []
    for c in range(1, total_cols + 1):
        hdr = combined[c - 1]
        if not hdr:
            continue
        gtype, period = _classify_header(hdr)
        if gtype == "open":
            target = keep_open
            ref = p_start
        elif gtype == "close":
            target = keep_close
            ref = p_end
        else:
            # 本期收入 / 本期发出 / 固定列：不参与按期间清空，列保留
            continue
        if target is None:
            continue
        # 无期号的余额列视作起始期（期初）/期末期（期末）
        eff = period if period is not None else ref
        if eff != target:
            to_clear.append(c)

    if not to_clear:
        return 0
    # 清空这些列数据行的单元格内容（列本身保留不动）
    for c in to_clear:
        for r in range(first_data_row, last_row + 1):
            ws.Cells(r, c).ClearContents()
    log(f"🧹 已清空 {len(to_clear)} 列的非保留期间余额数据（列保留不动）："
        + "、".join(combined[c - 1] for c in to_clear[:6])
        + ("…" if len(to_clear) > 6 else ""))
    return len(to_clear)


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
    deleted_columns: int = 0
    deleted_rows: int = 0
    cleared_columns: int = 0


def process_inventory(
    input_path: str | Path,
    output_path: str | Path,
    header_mode: str = "single",
    sep: str = "-",
    progress_callback: Callable[[str], None] | None = None,
    keep_open_period: int | None = None,
    keep_close_period: int | None = None,
) -> CleanResult:
    """清洗存货收发存汇总表，生成单层（或双层）表头的新 Excel。

    Args:
        input_path: 源 .xlsx 路径。
        output_path: 目标 .xlsx 路径（父目录会被创建；会另存，不覆盖源文件）。
        header_mode: "single" 单层表头（合并成一行，删除原第 4、5 行）；
                     "double" 双层表头（保留分组行 + 合并后的明细行，仅删原第 5 行）。
        sep: 层级结合的分隔符，默认 "-" 。
        progress_callback: 进度日志回调，用于向 UI 回传消息。
        keep_open_period: 保留的「期初会计期间」（如 1）。非 None 时清空其余期间的
            期初结存（数量/单价/金额）数据单元格，**列保留不动**；仅保留该期余额数据。
        keep_close_period: 保留的「期末会计期间」（如 6）。非 None 时清空其余期间的
            期末结存（列名通常为「期末结存-*」或「本期结存-*」）数据单元格，**列保留不动**。
            注意：本功能只清空数据、不删除任何列；本期收入/本期发出等流转列也整列保留。

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

            # 捕获标题行（删除前），稍后用于解析会计期间范围
            title_text = _as_text(ws.Cells(1, 1).Value)

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

            # 步骤 4b：删除数据区的「合计/总计」汇总行（通常为数据区最后一行）
            # 识别方式：扫描数据区第 1 列，凡含「合计」/「总计」的行即删除（从下往上删避免偏移）。
            _log("步骤4b：删除数据区合计/总计行")
            deleted_rows = 0
            data_start = header_rows + 1
            last_data_row = int(ws.UsedRange.Rows.Count)
            for r in range(last_data_row, data_start - 1, -1):
                a_val = _as_text(ws.Cells(r, 1).Value)
                if a_val and ("合计" in a_val or "总计" in a_val):
                    ws.Rows(r).Delete()
                    deleted_rows += 1
                    _log(f"  已删除合计行（原第 {r} 行）")
            if deleted_rows:
                _log(f"✂ 共删除 {deleted_rows} 行合计/总计行")

            # === 步骤5：按保留的会计期间，清空非保留期间余额数据（列保留不动）===
            cleared_count = 0
            if keep_open_period is not None or keep_close_period is not None:
                cleared_count = _clear_non_kept_period_cells(
                    ws, combined, title_text,
                    keep_open_period, keep_close_period, header_rows, _log)

            # 重新统计最终表头与数据行数（清空数据不改变列/行结构，仍重新读取以确保准确）
            final_maxcol = int(ws.UsedRange.Columns.Count)
            if header_mode == "single":
                headers = [_as_text(ws.Cells(1, c).Value)
                           for c in range(1, final_maxcol + 1)]
            else:
                # 双层：以明细行（第 2 行）作为预览表头
                headers = [_as_text(ws.Cells(2, c).Value)
                           for c in range(1, final_maxcol + 1)]
            total_rows = int(ws.UsedRange.Rows.Count)
            data_rows = max(0, total_rows - header_rows)

            # 保存（另存为 .xlsx，不改动源文件）
            _log(f"保存结果：{out_p.name}")
            wb.SaveAs(str(out_p), FileFormat=_XL_FILE_FORMAT_XLSX)

            summary = f"✅ 完成：共 {final_maxcol} 列表头，{data_rows} 行数据"
            extra = []
            if deleted_rows:
                extra.append(f"删除 {deleted_rows} 行合计/总计")
            if cleared_count:
                extra.append(f"清空 {cleared_count} 列非保留期间余额数据（列保留）")
            summary += ("；" + "、".join(extra) + "。") if extra else "。"
            _log(summary)
            return CleanResult(
                input_path=in_p,
                output_path=out_p,
                header_mode=header_mode,
                header_count=final_maxcol,
                headers=headers,
                data_rows=data_rows,
                ok=True,
                deleted_columns=0,
                deleted_rows=deleted_rows,
                cleared_columns=cleared_count,
                log_lines=[],
            )
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
    finally:
        _quit_excel_app(app)
