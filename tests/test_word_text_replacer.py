#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""tests/test_word_text_replacer.py — word_text_replacer 逻辑层单元测试。"""
from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

from docx import Document

from features.word_text_replacer import word_text_replacer_logic as logic


def _make_docx(path: Path, paragraphs: list[str]) -> None:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(str(path))


def test_scan_docx_files() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / "a.docx").write_text("")
        (root / "b.docx").write_text("")
        (root / "c.txt").write_text("")
        (root / "~$temp.docx").write_text("")
        sub = root / "sub"
        sub.mkdir()
        (sub / "d.docx").write_text("")

        found = logic.scan_docx_files(root, include_subfolders=False)
        assert len(found) == 2
        assert all(p.suffix == ".docx" for p in found)
        assert not any(p.name.startswith("~$") for p in found)

        found_recursive = logic.scan_docx_files(root, include_subfolders=True)
        assert len(found_recursive) == 3


def test_replace_in_document() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "src.docx"
        dst = Path(tmp) / "dst.docx"
        _make_docx(src, ["Hello world", "world is great"])

        count = logic.replace_in_document(src, dst, [("world", "Python")])
        assert count == 2

        doc = Document(str(dst))
        texts = [p.text for p in doc.paragraphs]
        assert texts == ["Hello Python", "Python is great"]


def test_process_folder() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        input_folder = root / "input"
        output_folder = root / "output"
        input_folder.mkdir()
        output_folder.mkdir()

        _make_docx(input_folder / "a.docx", ["foo bar"])
        _make_docx(input_folder / "b.docx", ["bar baz"])

        logs: list[str] = []
        result = logic.process_folder(
            input_folder,
            output_folder,
            [("bar", "BAR")],
            include_subfolders=False,
            progress_callback=logs.append,
        )

        assert result.total_files == 2
        assert result.success_count == 2
        assert result.fail_count == 0
        assert result.total_replacements == 2
        assert len(logs) > 0

        # 输出文件存在且内容已替换
        assert (output_folder / "a.docx").exists()
        doc = Document(str(output_folder / "a.docx"))
        assert doc.paragraphs[0].text == "foo BAR"


def test_input_equals_output_raises() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        folder = Path(tmp)
        try:
            logic.process_folder(folder, folder, [("a", "b")])
            assert False, "应该抛出 ValueError"
        except ValueError as e:
            assert "不能相同" in str(e)


if __name__ == "__main__":
    test_scan_docx_files()
    test_replace_in_document()
    test_process_folder()
    test_input_equals_output_raises()
    print("word_text_replacer 逻辑层测试全部通过")
