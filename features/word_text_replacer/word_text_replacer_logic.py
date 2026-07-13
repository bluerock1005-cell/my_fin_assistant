#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""features/word_text_replacer/word_text_replacer_logic.py — Word 批量文本替换逻辑层。

实现方式：通过 win32com 驱动本机 **Microsoft Word** 完成查找/替换（COM 自动化）。

相比纯 python-docx 方案的优势：
- 覆盖**正文、表格单元格、页眉、页脚、文本框**（遍历 Word 的 StoryRanges 全故事范围）；
- 由 Word 自身执行替换，**最大程度保留原有字体 / 段落格式**；
- 支持跨 run、跨段落、跨表格的连续文本匹配。

代价（使用前提）：
- 仅限 Windows，且**本机必须安装 Microsoft Word**；
- 依赖 `pywin32`（提供 win32com）；
- COM 必须在已 `CoInitialize` 的线程中使用（UI 层用后台线程处理，本模块负责初始化）。

UI 与逻辑严格分离：本文件不依赖任何 Qt 类型；`win32com` 为惰性导入，
缺少 Word / COM 时模块仍可 import（仅运行时会抛出清晰错误）。
"""
from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

# ===== COM 常量（用字面量，避免依赖 gencache 生成的枚举常量） =====
_WD_REPLACE_NONE = 0  # wdReplaceNone：仅查找、不替换（用于计数）
_WD_REPLACE_ALL = 2   # wdReplaceAll：替换范围内全部
_WD_FIND_STOP = 0     # wdFindStop：到范围末尾即停止（不回绕）
_WD_FMT_DOCX = 16     # wdFormatDocumentDefault：.docx


def _is_temp_docx(path: Path) -> bool:
    """Word 临时文件以 ~$ 开头，打开会报错，需要跳过。"""
    return path.name.startswith("~$")


def scan_docx_files(input_folder: Path, include_subfolders: bool = False) -> list[Path]:
    """扫描输入文件夹下所有 .docx 文件（排除 Word 临时文件）。

    Args:
        input_folder: 输入文件夹路径。
        include_subfolders: 是否递归包含子文件夹。

    Returns:
        按相对路径排序后的 .docx 文件路径列表。
    """
    if not input_folder.exists():
        return []

    pattern = "**/*.docx" if include_subfolders else "*.docx"
    files = [p for p in input_folder.glob(pattern)
             if p.is_file() and not _is_temp_docx(p)]
    files.sort(key=lambda p: p.relative_to(input_folder).as_posix())
    return files


# ===== COM / Word 生命周期 =====

def _com_word_available() -> bool:
    """本机是否具备 COM Word 自动化所需的 pywin32。"""
    try:
        import win32com  # noqa: F401
        return True
    except Exception:
        return False


def _create_word_app(visible: bool = False):
    """创建并初始化 Word.Application 实例。

    必须在调用线程内调用（内部负责 CoInitialize）；返回的应用实例
    由调用方在用完后通过 :func:`_quit_word_app` 释放。

    Raises:
        RuntimeError: 无法启动 Word（未安装 / pywin32 缺失 / 非 Windows）。
    """
    import pythoncom
    import win32com.client

    # 同一线程重复 CoInitialize 会抛错，忽略即可
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except pythoncom.error:
        pass

    try:
        app = win32com.client.Dispatch("Word.Application")
    except Exception as exc:  # noqa: BLE001
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        raise RuntimeError(
            "无法启动 Microsoft Word（COM 自动化失败）。请确认：\n"
            "  1) 当前为 Windows 且已安装 Microsoft Word；\n"
            "  2) 已安装 pywin32（pip install pywin32）；\n"
            "  3) Word 未被其他进程独占锁定。"
        ) from exc

    app.Visible = visible
    app.DisplayAlerts = False   # 抑制保存/格式等弹窗
    app.ScreenUpdating = False  # 后台静默处理，提速
    return app


def _quit_word_app(app) -> None:
    """退出 Word 并释放当前线程的 COM 初始化。对 None 安全。"""
    if app is None:
        return
    try:
        app.Quit(SaveChanges=False)
    except Exception:
        pass
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


# ===== 替换核心 =====

def _count_in_range(rng, text: str) -> int:
    """统计某 story range 内 text 的出现次数（Replace=None，不改变内容）。

    ⚠️ 必须用 wdReplaceNone(0) 计数：用 wdReplaceOne(1) 把文本替换为自身会让
    Word 在 Range.Find 上崩溃（已验证）。
    """
    if not text:
        return 0
    count = 0
    f = rng.Find
    f.ClearFormatting()
    while f.Execute(text, False, False, False, False, False,
                    True, _WD_FIND_STOP, False, text, _WD_REPLACE_NONE):
        count += 1
    return count


def _replace_all_rules(doc, rules: list[tuple[str, str]]) -> int:
    """在整篇文档（全部 story range）上，按规则顺序逐条执行「全量替换」。

    关键设计：
    - **每条规则都重新遍历一次 StoryRanges**，避免上一条规则替换后 range 被折叠，
      导致后续规则的查找范围残缺；
    - 计数在 `rng.Duplicate` 副本上进行，实际替换作用在原始 story range 上，
      以保证正文 / 表格 / 文本框等所有位置都被正确覆盖。

    Returns:
        全文档累计替换次数。
    """
    total = 0
    for old, new in rules:
        if not old:
            continue
        for story in doc.StoryRanges:
            rng = story
            while rng is not None:
                n = _count_in_range(rng.Duplicate, old)
                if n > 0:
                    f = rng.Find
                    f.ClearFormatting()
                    f.Replacement.ClearFormatting()
                    f.Execute(old, False, False, False, False, False,
                              True, _WD_FIND_STOP, False, new, _WD_REPLACE_ALL)
                    total += n
                rng = rng.NextStoryRange
    return total


def _replace_in_document_via_com(
    app, input_path: Path, output_path: Path, rules: list[tuple[str, str]]
) -> int:
    """用已创建的 Word 实例替换单篇文档并另存为 docx。"""
    doc = app.Documents.Open(
        str(input_path.resolve()),
        AddToRecentFiles=False,
        ConfirmConversions=False,
    )
    try:
        total = _replace_all_rules(doc, rules)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs(str(output_path.resolve()), FileFormat=_WD_FMT_DOCX)
        return total
    finally:
        try:
            doc.Close(SaveChanges=False)  # 源文件绝不回写
        except Exception:
            pass


def replace_in_document(
    input_path: Path,
    output_path: Path,
    rules: list[tuple[str, str]],
    word_app=None,
) -> int:
    """替换单篇 Word 文档中的文本。

    Args:
        input_path: 源 docx 路径。
        output_path: 目标 docx 路径（父目录会被创建）。
        rules: 替换规则列表，每项为 (old_text, new_text)。
        word_app: 可选，复用的 Word.Application 实例；
            为 None 时本函数自行创建并在结束时退出（仅用于单次调用）。

    Returns:
        该文档的替换次数。

    Note:
        调用线程需已 CoInitialize（由 :func:`_create_word_app` 处理）。
    """
    if not _com_word_available():
        raise RuntimeError(
            "本机缺少 pywin32，无法进行 Word COM 替换。请执行：pip install pywin32"
        )

    owned = word_app is None
    app = word_app if word_app is not None else _create_word_app(False)
    try:
        return _replace_in_document_via_com(app, input_path, output_path, rules)
    finally:
        if owned:
            _quit_word_app(app)


# ===== 数据结果 =====

@dataclass
class ReplaceResult:
    """单次文件替换结果。"""
    input_path: Path
    output_path: Path
    replace_count: int
    ok: bool
    error: str | None = None


@dataclass
class BatchResult:
    """批量替换汇总结果。"""
    input_folder: Path
    output_folder: Path
    files: list[ReplaceResult]
    total_files: int
    success_count: int
    fail_count: int
    total_replacements: int
    log_lines: list[str]

    @property
    def ok(self) -> bool:
        return self.fail_count == 0


# ===== 批量入口 =====

def process_folder(
    input_folder: Path,
    output_folder: Path,
    rules: list[tuple[str, str]],
    include_subfolders: bool = False,
    progress_callback: Callable[[str], None] | None = None,
) -> BatchResult:
    """批量替换入口（COM / Word 自动化）。

    Args:
        input_folder: 输入文件夹。
        output_folder: 输出文件夹（必须不同于输入文件夹）。
        rules: 替换规则列表。
        include_subfolders: 是否递归处理子文件夹。
        progress_callback: 进度日志回调，用于向 UI 回传消息。

    Returns:
        BatchResult 汇总对象。
    """
    def _log(msg: str) -> None:
        if progress_callback is not None:
            progress_callback(msg)

    log_lines: list[str] = []

    # ---- 校验（在启动 Word 之前完成，便于即时报错） ----
    if not input_folder or not output_folder:
        raise ValueError("输入文件夹和输出文件夹都必须选择")
    if input_folder == output_folder:
        raise ValueError("输入文件夹和输出文件夹不能相同（避免覆盖原文件）")
    if not input_folder.exists():
        raise FileNotFoundError(f"输入文件夹不存在：{input_folder}")

    files = scan_docx_files(input_folder, include_subfolders)
    _log(f"扫描到 {len(files)} 个 .docx 文件（{'含子文件夹' if include_subfolders else '仅当前文件夹'}）")

    if not files:
        return BatchResult(
            input_folder=input_folder,
            output_folder=output_folder,
            files=[],
            total_files=0,
            success_count=0,
            fail_count=0,
            total_replacements=0,
            log_lines=log_lines,
        )

    output_folder.mkdir(parents=True, exist_ok=True)

    results: list[ReplaceResult] = []
    success_count = 0
    fail_count = 0
    total_replacements = 0

    # 整个批量过程复用同一个 Word 实例（仅启动/退出一次）
    app = _create_word_app(False)
    try:
        for src in files:
            rel = src.relative_to(input_folder)
            dst = output_folder / rel
            _log(f"处理：{rel.as_posix()}")
            try:
                count = replace_in_document(src, dst, rules, word_app=app)
                results.append(ReplaceResult(src, dst, count, ok=True))
                success_count += 1
                total_replacements += count
                _log(f"  → 替换 {count} 处，已保存到 {dst.relative_to(output_folder).as_posix()}")
            except Exception as e:  # noqa: BLE001
                results.append(ReplaceResult(src, dst, 0, ok=False, error=str(e)))
                fail_count += 1
                _log(f"  → 失败：{e}")
    finally:
        _quit_word_app(app)

    return BatchResult(
        input_folder=input_folder,
        output_folder=output_folder,
        files=results,
        total_files=len(files),
        success_count=success_count,
        fail_count=fail_count,
        total_replacements=total_replacements,
        log_lines=log_lines,
    )
