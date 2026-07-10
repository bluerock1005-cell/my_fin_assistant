#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""webview_api.py — pywebview 桥接层（Python → React 前端）。

把三个功能的纯 Python 逻辑通过 pywebview 的 js_api 暴露给前端调用。
前端 (React) 通过 window.pywebview.api.<method>(...) 调用，返回 Promise。

设计原则：
    - 所有公开方法（不以 _ 开头）暴露给 JS
    - 重活（生成 Excel / 读大文件）写成 async def + asyncio.to_thread
    - 方法返回 JSON 可序列化的 dict；出错返回 {"error": "..."} 而不是抛异常
    - 文件选择/保存用原生对话框，在主线程弹出
"""
from __future__ import annotations

import asyncio
import base64
import json
from datetime import datetime
from pathlib import Path
from typing import Any

import webview

from core import app_config

# 纯逻辑层（不导入任何 ui，避免拉起 PySide6）
from features.bank_classify import classify_logic as bank_logic
from features.js_bank_statement import logic as js_logic
from features.notes_receivable_import import import_logic as notes_logic

# 文件对话框过滤器
_FT_EXCEL_OPEN = ("Excel 文件 (*.xlsx;*.xlsm;*.xls)",)
_FT_EXCEL_SAVE = ("Excel 文件 (*.xlsx)",)
_FT_BANK_OPEN = ("Excel/CSV/TXT (*.xlsx;*.xlsm;*.csv;*.txt)",)


def _p(path: Any) -> str:
    """Path / str -> str，方便 JSON 序列化。"""
    return str(path) if path is not None else ""


class WebAPI:
    """暴露给前端的 API 对象（js_api）。"""

    def __init__(self) -> None:
        self.window = None
        app_config.ensure_dirs()

        # ---- 应收票据导入的服务端状态 ----
        self._upload_dir = app_config.DATA_DIR / "_uploads"
        self._upload_dir.mkdir(parents=True, exist_ok=True)
        self.notes_files: list[Path] = []
        self.notes_template: Path = notes_logic.get_template_path(app_config.DATA_DIR)
        self.notes_manual: dict[str, dict[str, str | None]] = self._load_notes_overrides()
        self.notes_last_recv: str = ""

    # ===================================================================
    # 元信息与导航
    # ===================================================================

    def get_meta(self) -> dict:
        features = [
            {"name": "bank_classify", "icon": "landmark", "title": "票据银行分类"},
            {"name": "js_statement", "icon": "receipt", "title": "江苏银行对账单复制"},
            {"name": "notes_import", "icon": "table", "title": "应收票据批量导入"},
        ]
        return {
            "app_name": app_config.APP_NAME,
            "version": app_config.APP_VERSION,
            "features": features,
        }

    # ===================================================================
    # 通用：文件选择 / 保存 / 剪贴板
    # ===================================================================

    def pick_files(self, kind: str = "excel") -> list[str]:
        win = self.window or webview.windows[0]
        ft = {"excel": _FT_EXCEL_OPEN, "bank": _FT_BANK_OPEN}.get(kind, _FT_EXCEL_OPEN)
        try:
            paths = win.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=True, file_types=ft)
        except Exception:
            paths = None
        return list(paths or [])

    def save_file_dialog(self, default_name: str, kind: str = "excel") -> str | None:
        win = self.window or webview.windows[0]
        ft = _FT_EXCEL_SAVE if kind == "excel" else _FT_EXCEL_OPEN
        try:
            path = win.create_file_dialog(
                webview.SAVE_DIALOG, save_filename=default_name, file_types=ft
            )
        except Exception:
            path = None
        return path or None

    def read_clipboard(self) -> str:
        try:
            return bank_logic._read_clipboard() or ""
        except Exception as e:  # noqa: BLE001
            return f"__ERR__{e}"

    # ===================================================================
    # 票据银行分类
    # ===================================================================

    @staticmethod
    def _classify_preview(banks: list[str]) -> dict:
        rows = []
        for i, name in enumerate(banks, 1):
            short, cat = bank_logic.classify(name)
            rows.append({"seq": i, "name": name, "short": short, "cat": cat})
        n_yes = sum(1 for r in rows if r["cat"] == "21银行承兑汇票")
        n_no = len(rows) - n_yes
        return {"rows": rows, "n_yes": n_yes, "n_no": n_no, "total": len(banks)}

    def bank_whitelist(self) -> dict:
        groups = [
            ("国有大型银行", ["工商银行", "农业银行", "中国银行", "建设银行",
                             "交通银行", "邮储银行"]),
            ("股份制商业银行", ["中信银行", "光大银行", "华夏银行", "民生银行", "招商银行",
                               "兴业银行", "浦发银行", "浙商银行", "广发银行", "平安银行"]),
            ("城市商业银行", ["宁波银行", "江苏银行", "南京银行", "北京银行", "上海银行"]),
        ]
        return {"groups": groups, "total": sum(len(b) for _, b in groups)}

    def bank_load_file(self, path: str) -> dict:
        try:
            banks = bank_logic.load_banks(path, None, False)
        except Exception as e:  # noqa: BLE001
            return {"error": f"读取文件失败：{e}"}
        return {"banks": banks, **self._classify_preview(banks)}

    def bank_preview(self, text: str) -> dict:
        banks = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
        if not banks:
            return {"rows": [], "n_yes": 0, "n_no": 0, "total": 0}
        return self._classify_preview(banks)

    async def bank_run(self, text: str, out_path: str) -> dict:
        banks = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
        if not banks:
            return {"error": "未检测到任何银行名称"}
        if not out_path:
            return {"error": "未选择导出路径"}
        try:
            n_yes, n_no = await asyncio.to_thread(
                bank_logic.build_workbook, banks, Path(out_path)
            )
        except Exception as e:  # noqa: BLE001
            return {"error": f"生成 Excel 失败：{e}"}
        return {"n_yes": n_yes, "n_no": n_no, "total": len(banks), "path": out_path}

    # ===================================================================
    # 江苏银行对账单复制
    # ===================================================================

    async def js_load(self, path: str) -> dict:
        if not path:
            return {"error": "未选择文件"}
        try:
            headers, raw_rows = await asyncio.to_thread(js_logic.load_statement, path)
        except Exception as e:  # noqa: BLE001
            return {"error": f"读取文件失败：{e}"}
        rows = [[js_logic._cell_to_display_str(v) for v in r] for r in raw_rows]
        return {"headers": headers, "rows": rows, "name": Path(path).name}

    # ===================================================================
    # 应收票据批量导入
    # ===================================================================

    def notes_get_template_fields(self) -> dict:
        fields = [{"header": tf.header, "required": tf.required} for tf in notes_logic.TEMPLATE_FIELDS]
        readonly = sorted(notes_logic.FIXED_FIELD_HEADERS | notes_logic.PAGE_INPUT_HEADERS)
        return {
            "fields": fields,
            "readonly": readonly,
            "template_name": self.notes_template.name,
            "template_path": _p(self.notes_template),
        }

    def notes_add_files(self, paths: list[str]) -> dict:
        for p in paths or []:
            pp = Path(p)
            if pp.exists() and pp not in self.notes_files:
                self.notes_files.append(pp)
        return self._notes_file_list()

    def notes_add_dropped(self, name: str, b64: str) -> dict:
        try:
            data = base64.b64decode(b64)
        except Exception:
            return {"error": "文件解码失败"}
        target = self._upload_dir / name
        if target.exists():
            stem, suffix = target.stem, target.suffix
            target = self._upload_dir / f"{stem}_{datetime.now():%H%M%S}{suffix}"
        target.write_bytes(data)
        if target not in self.notes_files:
            self.notes_files.append(target)
        return self._notes_file_list()

    def notes_clear(self) -> dict:
        self.notes_files = []
        return self._notes_file_list()

    def notes_remove_file(self, name: str) -> dict:
        self.notes_files = [p for p in self.notes_files if p.name != name]
        return self._notes_file_list()

    def notes_list_files(self) -> dict:
        return self._notes_file_list()

    def _notes_file_list(self) -> dict:
        return {"files": [{"name": p.name, "path": _p(p)} for p in self.notes_files]}

    def _build_manual_maps(self) -> tuple[dict | None, dict | None]:
        positive: dict[str, dict[str, str]] = {}
        excluded: dict[str, list[str]] = {}
        fixed = notes_logic.FIXED_FIELD_HEADERS | notes_logic.PAGE_INPUT_HEADERS
        for name, mp in self.notes_manual.items():
            pos = {src: tgt for tgt, src in mp.items() if src not in (None, "") and tgt not in fixed}
            excl = [tgt for tgt, src in mp.items() if src in (None, "") and tgt not in fixed]
            if pos:
                positive[name] = pos
            if excl:
                excluded[name] = excl
        return (positive or None), (excluded or None)

    def _import_core(self, recv_org: str) -> dict:
        self.notes_last_recv = recv_org or ""
        if not self.notes_files:
            return {"error": "请先添加至少一个来源 Excel 文件"}
        if not recv_org or not recv_org.strip():
            return {"error": "收款组织不能为空"}

        defaults = {"*收款组织": recv_org.strip() or None}
        manual_maps, excluded_maps = self._build_manual_maps()

        try:
            result = notes_logic.read_and_transform(
                list(self.notes_files), defaults, manual_maps, excluded_maps
            )
        except Exception as e:  # noqa: BLE001
            return {"error": f"处理过程中出错：{e}"}

        files_info = []
        for fr in result.files:
            try:
                pure_auto, _u, _m = notes_logic.match_columns(list(fr.source_headers), None)
            except Exception:
                pure_auto = []
            auto_map = {m.target_field: m.source_header for m in pure_auto}
            eff_map = {m.target_field: m.source_header for m in fr.matched_columns}
            files_info.append({
                "name": fr.file_name,
                "source_headers": list(fr.source_headers),
                "auto_map": auto_map,
                "current_map": eff_map,
                "matched": [
                    {"target": m.target_field, "source": m.source_header, "conf": round(m.confidence, 2)}
                    for m in fr.matched_columns
                ],
                "unmatched": list(fr.unmatched_source),
                "missing": list(fr.missing_required),
                "errors": list(fr.errors),
            })

        table_data = [
            {tf.header: notes_logic.value_to_display(rec.get(tf.header)) for tf in notes_logic.TEMPLATE_FIELDS}
            for rec in result.table_data
        ]

        has_missing = any(f["missing"] for f in files_info)
        return {
            "files": files_info,
            "table_data": table_data,
            "headers": list(notes_logic.TEMPLATE_HEADERS),
            "log_lines": list(result.log_lines),
            "row_count": len(result.table_data),
            "has_missing": has_missing,
            "recv_org": recv_org,
            "applied_manual": {
                n: self.notes_manual.get(n)
                for n in (f["name"] for f in files_info)
                if self.notes_manual.get(n)
            },
        }

    async def notes_import(self, recv_org: str) -> dict:
        return await asyncio.to_thread(self._import_core, recv_org or "")

    def notes_save_mapping(self, maps: dict) -> dict:
        fixed = notes_logic.FIXED_FIELD_HEADERS | notes_logic.PAGE_INPUT_HEADERS
        clean: dict[str, dict[str, str | None]] = {}
        for name, mp in (maps or {}).items():
            if not isinstance(mp, dict):
                continue
            row = {
                tgt: (src if src not in (None, "") else None)
                for tgt, src in mp.items()
                if tgt not in fixed
            }
            clean[name] = row
        self.notes_manual = clean
        self._save_notes_overrides()
        return self._import_core(self.notes_last_recv or "")

    async def notes_export(self, rows: list[dict], out_path: str) -> dict:
        if not rows:
            return {"error": "没有可导出的数据"}
        if not out_path:
            return {"error": "未选择导出路径"}

        def _reclean(rec: dict) -> dict:
            return {tf.header: notes_logic.clean_value(rec.get(tf.header), tf.header) for tf in notes_logic.TEMPLATE_FIELDS}

        cleaned = [_reclean(r) for r in rows]
        try:
            written = await asyncio.to_thread(
                notes_logic.write_to_template, cleaned, Path(out_path),
                self.notes_template, app_config.DATA_DIR,
            )
        except Exception as e:  # noqa: BLE001
            return {"error": f"写入 Excel 失败：{e}"}
        return {"written": written, "path": out_path}

    # ===================================================================
    # 映射持久化
    # ===================================================================

    _OVERRIDES_FILE = app_config.DATA_DIR / "mapping_overrides.json"

    def _load_notes_overrides(self) -> dict[str, dict[str, str | None]]:
        if not self._OVERRIDES_FILE.exists():
            return {}
        try:
            with open(self._OVERRIDES_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (json.JSONDecodeError, OSError):
            return {}
        if not isinstance(data, dict):
            return {}
        fixed = notes_logic.FIXED_FIELD_HEADERS | notes_logic.PAGE_INPUT_HEADERS
        clean: dict[str, dict[str, str | None]] = {}
        for name, mp in data.items():
            if isinstance(mp, dict):
                clean[name] = {
                    tgt: (src if src not in (None, "") else None)
                    for tgt, src in mp.items()
                    if tgt not in fixed
                }
        return clean

    def _save_notes_overrides(self) -> None:
        try:
            app_config.DATA_DIR.mkdir(parents=True, exist_ok=True)
            with open(self._OVERRIDES_FILE, "w", encoding="utf-8") as fh:
                json.dump(self.notes_manual, fh, ensure_ascii=False, indent=2)
        except OSError:
            pass
