#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""features/notes_receivable_import/ui.py — 应收票据批量导入界面（PySide6 + Fluent）。

产品规范（截图）：
  - 输入方式：拖拽 .xls/.xlsx 文件到窗口
  - 导入解析：按关键字映射规则自动转换为标准字段
  - 新文件导入行为：清空表格，只显示本次导入数据（不追加历史）
  - 表格展示：可编辑表格，展示转换后数据，必填缺字段标红提示
  - 手动编辑：支持双击单元格直接修改
  - 模板：默认固定模板，可点击按钮切换为其他模板文件
  - 模板作用：仅提供表头结构/样式参考，不被程序改写
  - 导出：点击按钮触发导出，弹窗选择保存路径和文件名
  - 目标文件已存在：直接覆盖（不追加、不合并）

映射确认环节（新增）：
  - 导入来源文件后自动做关键字匹配；若不满意，可点「配置映射」打开对话框
  - 对话框左右两列：左=模板字段（固定，必录加*标红），右=来源列下拉
  - 下拉默认=自动匹配结果；用户只改不满意的几行，确定后按最终映射重算表格
  - 映射关系会保存（data/mapping_overrides.json），并在对话框内做冲突提示

UI 和逻辑严格分离：本文件只管界面与交互，
纯计算/数据处理放在 import_logic.py。
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
import concurrent.futures

from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import (
    QDragEnterEvent,
    QDropEvent,
    QColor,
    QStandardItem,
    QStandardItemModel,
)
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core import app_config, theme, utils
from core.feature_base import FeatureModule

# 纯逻辑层
from . import import_logic as _logic


# 映射覆盖持久化文件名
_OVERRIDES_FILE = app_config.DATA_DIR / "mapping_overrides.json"
# 下拉中「不匹配/固定值」哨兵（combobox userData 用空串表示）
_SENTINEL = ""


class _EditableTableView(QTableView):
    """可编辑表格视图，双击单元格可直接修改内容。"""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        # 启用编辑
        self.setEditTriggers(
            QTableView.EditTrigger.DoubleClicked | QTableView.EditTrigger.EditKeyPressed
        )
        self.setSelectionBehavior(QTableView.SelectionBehavior.SelectItems)
        # 交替行色（浅灰区分）
        self.setAlternatingRowColors(True)
        # 水平表头可拉伸
        self.horizontalHeader().setStretchLastSection(True)


class MappingDialog(QDialog):
    """字段映射配置对话框。

    左列：模板字段（固定，必录加 * 标红）。
    右列：来源列下拉，选项 = 本次导入文件实际列名 + 〈不匹配/固定值〉。
    默认值 = 自动匹配结果（或上次保存的手动映射）。
    实时冲突提示：同一来源列被多个字段重复选择 → 标红；必录字段未匹配 → 警告。
    支持多文件：顶部文件下拉切换，每个文件独立维护一份映射。

    result_maps 属性返回 {file_name: {tgt_header: src_header_or_None}}。
    """

    def __init__(self, parent: QWidget | None,
                 files_info: list[dict],
                 template_fields: list) -> None:
        super().__init__(parent)
        self.setWindowTitle("字段映射配置")
        self.setMinimumSize(640, 600)

        # files_info: list of dict(name, source_headers, auto_map, current_map)
        self._files_info = files_info
        self._names = [f["name"] for f in files_info]
        self._source_headers = {f["name"]: f["source_headers"] for f in files_info}
        # 自动匹配结果（对话框"恢复自动匹配"用）
        self._auto = {f["name"]: dict(f["auto_map"]) for f in files_info}
        # 当前编辑态（用户每次改动实时写回这里）
        self._editing: dict[str, dict[str, str | None]] = {}
        for f in files_info:
            # 固定字段由系统自动填充，不进入对话框编辑态
            base = {
                tf.header: f["current_map"].get(tf.header)
                for tf in template_fields
                if tf.header not in _logic.FIXED_FIELD_HEADERS
            }
            for tf in template_fields:
                if tf.header not in _logic.FIXED_FIELD_HEADERS:
                    base.setdefault(tf.header, None)
            self._editing[f["name"]] = base
        self._template_fields = template_fields
        # 对话框列表顺序：必录字段全部提前到最前，非必录字段排在后
        self._display_fields = sorted(
            template_fields,
            key=lambda tf: (0 if tf.required else 1,)
        )
        self._current = self._names[0] if self._names else ""
        self._result_maps: dict[str, dict[str, str | None]] = {}

        # 当前文件的行控件: list of (TemplateField, QLabel, QComboBox)
        self._row_widgets: list[tuple] = []

        self._setup_ui()
        self._load_file(self._current)
        self._validate()

    # ====== 布局 ======

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        title = QLabel("字段映射配置", self)
        title.setObjectName("pageTitle")
        root.addWidget(title)

        desc = QLabel(
            "必录字段排在最前（带 * 标红）；非必录字段默认留空，如需映射请在右侧下拉选择来源列。"
            "下拉默认已是自动匹配结果，只需调整不满意的几行。"
            "（导入序号、票据类型、币别、汇率、背书明细由系统自动填充，此处不可更改）",
            self,
        )
        desc.setObjectName("pageDesc")
        desc.setWordWrap(True)
        root.addWidget(desc)

        # 多文件选择器
        if len(self._names) > 1:
            fsel = QHBoxLayout()
            fsel.addWidget(QLabel("来源文件：", self))
            self._file_combo = QComboBox(self)
            self._file_combo.addItems(self._names)
            self._file_combo.setFixedHeight(30)
            self._file_combo.currentTextChanged.connect(self._on_file_switch)
            fsel.addWidget(self._file_combo, stretch=1)
            root.addLayout(fsel)

        # 表头行（与下方网格列对齐）
        head = QHBoxLayout()
        head.setContentsMargins(4, 0, 4, 0)
        hl = QLabel("导出列（模板字段）", self)
        hl.setObjectName("cardTitle")
        hr = QLabel("来源列（下拉选择）", self)
        hr.setObjectName("cardTitle")
        head.addWidget(hl, stretch=1)
        head.addWidget(hr, stretch=1)
        root.addLayout(head)

        divider = QFrame(self)
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Plain)
        divider.setObjectName("cardDivider")
        root.addWidget(divider)

        # 滚动区：内容外框加轻量卡片样式，避免行内容与背景混在一起
        self._scroll = QScrollArea(self)
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._content = QWidget(self)
        self._grid = QGridLayout(self._content)
        self._grid.setContentsMargins(4, 8, 4, 8)
        self._grid.setHorizontalSpacing(16)
        self._grid.setVerticalSpacing(10)
        self._grid.setColumnStretch(0, 1)
        self._grid.setColumnStretch(1, 1)
        self._scroll.setWidget(self._content)
        root.addWidget(self._scroll, stretch=1)

        # 冲突提示条
        self._warn = QLabel(self)
        self._warn.setObjectName("pageDesc")
        self._warn.setWordWrap(True)
        self._warn.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self._warn.hide()
        root.addWidget(self._warn)

        # 按钮行
        btns = QHBoxLayout()
        self._btn_restore = QPushButton("恢复自动匹配", self)
        self._btn_restore.setFixedHeight(32)
        self._btn_restore.clicked.connect(self._restore_auto)
        btns.addWidget(self._btn_restore)
        btns.addStretch(1)
        self._btn_ok = QPushButton("确定", self)
        self._btn_ok.setObjectName("primary")
        self._btn_ok.setFixedHeight(32)
        self._btn_ok.setDefault(True)
        self._btn_ok.clicked.connect(self.accept)
        self._btn_cancel = QPushButton("取消", self)
        self._btn_cancel.setFixedHeight(32)
        self._btn_cancel.clicked.connect(self.reject)
        btns.addWidget(self._btn_ok)
        btns.addWidget(self._btn_cancel)
        root.addLayout(btns)

    # ====== 文件切换 / 行构建 ======

    def _on_file_switch(self, name: str) -> None:
        if name and name in self._names:
            self._current = name
            self._load_file(name)
            self._validate()

    def _clear_grid(self) -> None:
        while self._grid.count():
            item = self._grid.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._row_widgets = []

    def _load_file(self, name: str) -> None:
        self._clear_grid()
        headers = self._source_headers.get(name, [])
        row_map = self._editing.get(name, {})
        r = 0
        for tf in self._display_fields:
            # 固定字段不显示在对话框中（由系统自动填充）
            if tf.header in _logic.FIXED_FIELD_HEADERS:
                continue
            # 左列：模板字段名（必录标红）
            lbl = QLabel(tf.header, self._content)
            lbl.setObjectName("cardTitle" if not tf.required else "reqField")
            if tf.required:
                lbl.setStyleSheet("color: #DC2626; font-weight: 600;")
            else:
                lbl.setStyleSheet("color: #1F2937;")
            lbl.setFixedHeight(30)
            lbl.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)

            # 右列：下拉
            combo = QComboBox(self._content)
            combo.setFixedHeight(30)
            # 哨兵项（不匹配/固定值）
            combo.addItem("〈不匹配/固定值〉", _SENTINEL)
            # 来源列
            for h in headers:
                if h and h.strip():
                    combo.addItem(h, h)
            combo.setCurrentIndex(self._index_for(combo, row_map.get(tf.header)))
            combo.currentIndexChanged.connect(
                lambda _idx, t=tf.header: self._on_combo_changed(t))

            self._grid.addWidget(lbl, r, 0)
            self._grid.addWidget(combo, r, 1)
            self._row_widgets.append((tf, lbl, combo))
            r += 1

    @staticmethod
    def _index_for(combo: QComboBox, src: str | None) -> int:
        """根据来源列名找到下拉索引（None/空 → 哨兵项）。"""
        target = src if src else _SENTINEL
        for i in range(combo.count()):
            if combo.itemData(i) == target:
                return i
        # 回退到哨兵项
        for i in range(combo.count()):
            if combo.itemData(i) == _SENTINEL:
                return i
        return 0

    # ====== 交互 ======

    def _on_combo_changed(self, tgt: str) -> None:
        combo = self._combo_for(tgt)
        if combo is None:
            return
        data = combo.currentData()
        src = data if data not in (None, _SENTINEL) else None
        self._editing.setdefault(self._current, {})[tgt] = src
        self._validate()

    def _combo_for(self, tgt: str):
        for tf, _lbl, combo in self._row_widgets:
            if tf.header == tgt:
                return combo
        return None

    def _restore_auto(self) -> None:
        """把当前文件的映射重置为自动匹配结果（仅非固定字段）。"""
        auto = self._auto.get(self._current, {})
        self._editing[self._current] = {
            tf.header: auto.get(tf.header)
            for tf in self._template_fields
            if tf.header not in _logic.FIXED_FIELD_HEADERS
        }
        # 重新载入下拉
        self._load_file(self._current)
        self._validate()

    def _validate(self) -> None:
        """冲突检测：重复来源列标红 + 必录未匹配警告。"""
        row_map = self._editing.get(self._current, {})
        # 统计来源列被哪些字段选用（排除哨兵）
        usage: dict[str, list[str]] = {}
        for tf, _lbl, _combo in self._row_widgets:
            src = row_map.get(tf.header)
            if src:  # 非 None/空
                usage.setdefault(src, []).append(tf.header)
        duplicates = {src: ts for src, ts in usage.items() if len(ts) > 1}

        warn_lines: list[str] = []
        for tf, lbl, combo in self._row_widgets:
            src = row_map.get(tf.header)
            is_dup = src is not None and src in duplicates
            is_req_unmatched = tf.required and (src is None or src == "")
            if is_dup:
                combo.setStyleSheet("border: 1px solid #DC2626; "
                                    "background-color: #FEF2F2;")
                lbl.setStyleSheet("color: #DC2626; font-weight: 700;")
            elif is_req_unmatched:
                combo.setStyleSheet("border: 1px solid #D97706; "
                                    "background-color: #FFFBEB;")
                lbl.setStyleSheet("color: #D97706; font-weight: 600;")
            else:
                combo.setStyleSheet("")
                if tf.required:
                    lbl.setStyleSheet("color: #DC2626; font-weight: 600;")
                else:
                    lbl.setStyleSheet("color: #1F2937;")

        for src, ts in duplicates.items():
            warn_lines.append(
                f"⚠ 来源列「{src}」被重复选择：{'、'.join(ts)}（同一列映射到多个字段）")
        for tf, _lbl, _combo in self._row_widgets:
            src = row_map.get(tf.header)
            if tf.required and (src is None or src == ""):
                warn_lines.append(
                    f"⚠ 必录字段「{tf.header}」未匹配来源列，将留空或取默认值")

        if warn_lines:
            self._warn.setText("\n".join(warn_lines))
            self._warn.setStyleSheet(
                "color: #DC2626; background-color: #FEF2F2; "
                "border: 1px solid #FECACA; border-radius: 6px; padding: 8px;")
            self._warn.show()
        else:
            self._warn.hide()

    # ====== 结果 ======

    def accept(self) -> None:
        # 把每个文件的编辑态整理为最终结果（"" 归一为 None），排除固定字段
        for name in self._names:
            mp = self._editing.get(name, {})
            self._result_maps[name] = {
                tgt: (src if src not in (None, _SENTINEL) else None)
                for tgt, src in mp.items()
                if tgt not in _logic.FIXED_FIELD_HEADERS
            }
        super().accept()

    @property
    def result_maps(self) -> dict[str, dict[str, str | None]]:
        return self._result_maps


class NotesReceivableImportWidget(QWidget):
    """应收票据批量导入功能的具体 UI 控件。

    工作流：
      1. 拖拽/选择来源文件 → 点击「导入」→ 数据清洗后展示到可编辑表格
      2. 用户在表格中检查/修正数据 → 点击「导出」→ 弹窗选路径 → 写入模板
      3. 可随时切换模板文件
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("page")

        # 状态
        self._files: list[Path] = []
        self._template_path: Path = _logic.get_template_path(app_config.DATA_DIR)
        self._table_data: list[dict[str, object]] = []   # 当前表格中的数据（用户可能已编辑）
        self._result: _logic.ImportResult | None = None    # 最近一次导入的完整结果
        # 手工映射（已保存）：{file_name: {tgt_header: src_header_or_None}}
        self._manual_full: dict[str, dict[str, str | None]] = self._load_overrides()

        # 后台线程
        self._executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        self._future = None
        self._poll_timer = None

        self.setAcceptDrops(True)
        self._setup_ui()

    # ====== 映射覆盖持久化 ======

    def _load_overrides(self) -> dict[str, dict[str, str | None]]:
        """从 data/mapping_overrides.json 载入已保存的映射配置。

        注意：本方法在 __init__ 中、日志控件创建之前调用，
        出错时只打印不依赖 UI。
        """
        if not _OVERRIDES_FILE.exists():
            return {}
        try:
            with open(_OVERRIDES_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if not isinstance(data, dict):
                return {}
            # 归一化：把空串视为 None
            clean: dict[str, dict[str, str | None]] = {}
            for name, mp in data.items():
                if isinstance(mp, dict):
                    clean[name] = {
                        tgt: (src if src not in (None, "") else None)
                        for tgt, src in mp.items()
                        if tgt not in _logic.FIXED_FIELD_HEADERS
                    }
            return clean
        except (json.JSONDecodeError, OSError) as e:
            print(f"[warn] 读取映射配置失败：{e}", file=__import__("sys").stderr)
            return {}

    def _save_overrides(self) -> None:
        """把当前 self._manual_full 写回 data/mapping_overrides.json。"""
        try:
            app_config.DATA_DIR.mkdir(parents=True, exist_ok=True)
            with open(_OVERRIDES_FILE, "w", encoding="utf-8") as fh:
                json.dump(self._manual_full, fh, ensure_ascii=False, indent=2)
        except OSError as e:
            self._log_line(f"⚠ 保存映射配置失败：{e}")

    # ====== 布局 ======

    def _setup_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(theme.SPACING[24], theme.SPACING[24],
                                 theme.SPACING[24], theme.SPACING[24])
        root.setSpacing(theme.SPACING[16])

        # --- 页头 ---
        title = QLabel("应收票据批量导入", self)
        title.setObjectName("pageTitle")
        desc = QLabel(
            "拖拽或选择来源 Excel 文件 → 自动匹配列名并清洗数据 → "
            "在表格中检查修正 → 导出到固定模板。"
            "新文件导入会清空旧数据。",
            self,
        )
        desc.setObjectName("pageDesc")
        desc.setWordWrap(True)
        root.addWidget(title)
        root.addWidget(desc)
        root.addSpacing(theme.SPACING[4])

        # === 卡片1：文件 + 导入 ===
        card1 = QFrame(self)
        card1.setObjectName("card")
        c1 = QVBoxLayout(card1)
        c1.setContentsMargins(theme.SPACING[16], theme.SPACING[16],
                               theme.SPACING[16], theme.SPACING[16])
        c1.setSpacing(theme.SPACING[12])

        h1 = QHBoxLayout()
        h1.setSpacing(theme.SPACING[12])
        h1_lbl = QLabel("来源文件（拖拽 / 选择）", card1)
        h1_lbl.setObjectName("cardTitle")
        h1.addWidget(h1_lbl)
        h1.addStretch(1)
        btn_add = QPushButton("+ 添加文件", card1)
        btn_add.setFixedHeight(32)
        btn_add.setMinimumWidth(100)
        btn_add.clicked.connect(self._add_files)
        h1.addWidget(btn_add)
        # 映射配置入口（导入后启用）
        self._btn_map = QPushButton("配置映射", card1)
        self._btn_map.setFixedHeight(32)
        self._btn_map.setMinimumWidth(100)
        self._btn_map.setEnabled(False)
        self._btn_map.setToolTip("导入后可手工调整列名映射关系")
        self._btn_map.clicked.connect(self._open_mapping_dialog)
        h1.addWidget(self._btn_map)
        c1.addLayout(h1)

        # 细分隔线，把「文件选择」与「主操作」区隔开
        divider1 = QFrame(card1)
        divider1.setFrameShape(QFrame.Shape.HLine)
        divider1.setFrameShadow(QFrame.Shadow.Plain)
        divider1.setObjectName("cardDivider")
        c1.addWidget(divider1)

        # 导入 / 导出：主操作行，等宽突出，右侧附带进度提示
        h2 = QHBoxLayout()
        h2.setSpacing(theme.SPACING[12])
        self._btn_import = QPushButton("导入 / 刷新", card1)
        self._btn_import.setObjectName("primary")
        self._btn_import.setFixedHeight(34)
        self._btn_import.setMinimumWidth(120)
        self._btn_import.clicked.connect(self._do_import)
        h2.addWidget(self._btn_import)
        self._btn_export = QPushButton("导出 Excel", card1)
        self._btn_export.setObjectName("primary")
        self._btn_export.setFixedHeight(34)
        self._btn_export.setMinimumWidth(120)
        self._btn_export.setEnabled(False)
        self._btn_export.clicked.connect(self._do_export)
        h2.addWidget(self._btn_export)
        h2.addStretch(1)
        c1.addLayout(h2)

        # 映射状态行
        self._lbl_map_status = QLabel("尚未导入文件。", card1)
        self._lbl_map_status.setObjectName("pageDesc")
        self._lbl_map_status.setWordWrap(True)
        c1.addWidget(self._lbl_map_status)

        self._lbl_files = QLabel("未选择文件", card1)
        self._lbl_files.setObjectName("pageDesc")
        self._lbl_files.setWordWrap(True)
        c1.addWidget(self._lbl_files)

        root.addWidget(card1)

        # === 卡片2：默认值 + 模板 ===
        card2 = QFrame(self)
        card2.setObjectName("card")
        c2 = QVBoxLayout(card2)
        c2.setContentsMargins(theme.SPACING[16], theme.SPACING[16],
                               theme.SPACING[16], theme.SPACING[16])
        c2.setSpacing(theme.SPACING[12])

        dc_head = QLabel("默认值 / 模板", card2)
        dc_head.setObjectName("cardTitle")
        c2.addWidget(dc_head)

        # 用网格对齐「标签 + 输入框」，两行整齐排布，视觉上比多个 HBox 更规整
        dg = QGridLayout()
        dg.setHorizontalSpacing(theme.SPACING[12])
        dg.setVerticalSpacing(theme.SPACING[12])

        lbl_tpl = QLabel("模板:", card2)
        lbl_tpl.setFixedWidth(70)
        dg.addWidget(lbl_tpl, 0, 0)
        self._lbl_template = QLineEdit(card2)
        self._lbl_template.setReadOnly(True)
        self._lbl_template.setText(self._template_path.name)
        self._lbl_template.setFixedHeight(30)
        self._lbl_template.setToolTip(str(self._template_path))
        dg.addWidget(self._lbl_template, 0, 1)
        btn_switch_tpl = QPushButton("切换", card2)
        btn_switch_tpl.setFixedHeight(30)
        btn_switch_tpl.setMinimumWidth(72)
        btn_switch_tpl.clicked.connect(self._switch_template)
        dg.addWidget(btn_switch_tpl, 0, 2)

        lbl_recv = QLabel("收款组织:", card2)
        lbl_recv.setFixedWidth(70)
        dg.addWidget(lbl_recv, 1, 0)
        self._edit_recv_org = QLineEdit("", card2)
        self._edit_recv_org.setFixedHeight(30)
        self._edit_recv_org.setPlaceholderText("请填写本单位/我公司名称")
        dg.addWidget(self._edit_recv_org, 1, 1)

        dg.setColumnStretch(1, 1)
        dg.setColumnMinimumWidth(2, 72)
        c2.addLayout(dg)

        note = QLabel(
            "币别 / 票据类型 / 汇率 / 导入序号 / 背书明细 由系统自动填充，无需配置。",
            card2)
        note.setObjectName("pageDesc")
        note.setWordWrap(True)
        c2.addWidget(note)

        root.addWidget(card2)

        # === 卡片3：可编辑数据表格 ===
        card3 = QFrame(self)
        card3.setObjectName("card")
        c3 = QVBoxLayout(card3)
        c3.setContentsMargins(theme.SPACING[12], theme.SPACING[12],
                               theme.SPACING[12], theme.SPACING[12])
        c3.setSpacing(6)

        t_head = QHBoxLayout()
        t_lbl = QLabel("数据预览（可编辑，红色 = 缺少必录值）", card3)
        t_lbl.setObjectName("cardTitle")
        t_head.addWidget(t_lbl)
        t_head.addStretch(1)
        self._progress = QLabel("", card3)
        self._progress.setObjectName("pageDesc")
        t_head.addWidget(self._progress)
        self._lbl_row_count = QLabel("", card3)
        self._lbl_row_count.setObjectName("pageDesc")
        t_head.addWidget(self._lbl_row_count)
        c3.addLayout(t_head)

        self._table = _EditableTableView(card3)
        self._table.setMinimumHeight(200)
        self._model = QStandardItemModel(self)
        self._table.setModel(self._model)
        c3.addWidget(self._table, stretch=1)

        # 日志面板
        log_card = QFrame(self)
        log_card.setObjectName("card")
        lc_lay = QVBoxLayout(log_card)
        lc_lay.setContentsMargins(theme.SPACING[12], theme.SPACING[12],
                                    theme.SPACING[12], theme.SPACING[12])
        lc_lay.setSpacing(6)
        log_title = QLabel("运行日志（列匹配详情）", log_card)
        log_title.setObjectName("cardTitle")
        lc_lay.addWidget(log_title)
        self._log = QPlainTextEdit(log_card)
        self._log.setObjectName("logView")
        self._log.setReadOnly(True)
        self._log.setMinimumHeight(50)
        lc_lay.addWidget(self._log, stretch=1)

        # 表格区与日志区放入可拖动分栏：默认表格占多数空间，
        # 用户可按需拖大日志区查看详细匹配信息，比固定高度更灵活。
        splitter = QSplitter(Qt.Orientation.Vertical, self)
        splitter.setObjectName("previewSplitter")
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(card3)
        splitter.addWidget(log_card)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([420, 120])
        root.addWidget(splitter, stretch=1)

    # ====== 拖拽 ======

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        added = 0
        for url in event.mimeData().urls():
            fp = url.toLocalFile()
            if fp and fp.lower().endswith((".xlsx", ".xlsm", ".xls")):
                p = Path(fp)
                if p not in self._files:
                    self._files.append(p)
                    added += 1
        if added:
            self._update_file_label()
            self._log_line(f"拖拽添加 {added} 个文件。")
        event.acceptProposedAction()

    # ====== 交互 ======

    def _ts(self) -> str:
        return datetime.now().strftime("%H:%M:%S")

    def _log_line(self, msg: str) -> None:
        self._log.appendPlainText(f"[{self._ts()}] {msg}")

    def _update_file_label(self) -> None:
        n = len(self._files)
        if n == 0:
            self._lbl_files.setText("未选择文件")
        elif n == 1:
            self._lbl_files.setText(f"已导入 1 个文件：\n• {self._files[0].name}")
        else:
            lines = [f"已导入 {n} 个文件："]
            for i, p in enumerate(self._files, start=1):
                lines.append(f"{i}. {p.name}")
            self._lbl_files.setText("\n".join(lines))

    def _add_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择来源 Excel 文件",
            str(app_config.DATA_DIR),
            "Excel 文件 (*.xlsx *.xlsm *.xls);;所有文件 (*)",
        )
        for p in paths:
            pp = Path(p)
            if pp not in self._files:
                self._files.append(pp)
        if paths:
            self._update_file_label()
            self._log_line(f"已添加 {len(paths)} 个文件（当前共 {len(self._files)} 个）。")

    def _switch_template(self) -> None:
        p, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件",
            str(app_config.DATA_DIR),
            "Excel 模板 (*.xlsx *.xlsm);;所有文件 (*)",
        )
        if p:
            self._template_path = Path(p)
            self._lbl_template.setText(Path(p).name)
            self._lbl_template.setToolTip(p)
            self._log_line(f"已切换模板: {Path(p).name}")

    @staticmethod
    def _parse_float(text: str, default: float) -> float:
        try:
            return float(text.strip())
        except (ValueError, AttributeError):
            return default

    def _get_defaults(self) -> dict[str, object]:
        # 注：*币别、*票据类型、*汇率、*导入序号、*背书明细/导入序号 为固定字段，
        # 由 import_logic.FIXED_FIELDS 统一填充（人民币 / 银行承兑汇票 /
        # 1 / 自动序号 / 留空），此处不再提供默认值入口。
        # *付款单位 原在「默认值/模板」卡片有输入框，现按用户要求移除该行，
        # 统一默认「客户」（来源文件匹配到「付款单位」列时仍以来源值为准）。
        return {
            "收款组织": self._edit_recv_org.text().strip() or None,
            "*付款单位": "客户",
        }

    # ====== 导入（读取+匹配+清洗 → 展示到表格）======

    def _do_import(self) -> None:
        if not self._files:
            self._log_line("请先添加至少一个来源 Excel 文件。")
            return

        self._btn_import.setEnabled(False)
        self._btn_export.setEnabled(False)
        self._progress.setText("正在导入…")
        self._log.clear()
        self._log_line(f"开始导入 {len(self._files)} 个源文件…")

        defaults = self._get_defaults()

        # 由已保存的手工映射推导 positive / excluded 映射
        positive: dict[str, dict[str, str]] = {}
        excluded: dict[str, list[str]] = {}
        fixed = _logic.FIXED_FIELD_HEADERS
        for name, mp in self._manual_full.items():
            pos = {src: tgt for tgt, src in mp.items()
                   if src not in (None, "") and tgt not in fixed}
            excl = [tgt for tgt, src in mp.items()
                    if src in (None, "") and tgt not in fixed]
            if pos:
                positive[name] = pos
            if excl:
                excluded[name] = excl

        self._future = self._executor.submit(
            _logic.read_and_transform,
            list(self._files),
            defaults,
            positive if positive else None,
            excluded if excluded else None,
        )

        def _check():
            if self._future.done():
                try:
                    res = self._future.result()
                except Exception as e:
                    self._on_import_fail(str(e))
                else:
                    self._on_import_done(res)
                if self._poll_timer is not None:
                    self._poll_timer.stop()
                    self._poll_timer.deleteLater()
                    self._poll_timer = None

        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(100)
        self._poll_timer.timeout.connect(_check)
        self._poll_timer.start()

    def _on_import_done(self, res: _logic.ImportResult) -> None:
        self._btn_import.setEnabled(True)
        self._progress.setText("")
        self._result = res
        self._table_data = list(res.table_data)  # 拷贝一份供 UI 展示/编辑

        # 填充日志
        for line in res.log_lines:
            self._log.appendPlainText(line)

        # 若本次导入应用了已保存的手工映射，给出透明提示
        for fr in res.files:
            saved = self._manual_full.get(fr.file_name)
            if saved:
                pos_n = sum(1 for s in saved.values() if s not in (None, ""))
                unmatched_n = sum(1 for s in saved.values() if s in (None, ""))
                self._log_line(
                    f"ℹ 已载入「{fr.file_name}」的已保存映射"
                    f"（手工映射 {pos_n} 项，显式不匹配 {unmatched_n} 项）；"
                    f"可点「配置映射」继续调整。")

        # 填充表格
        self._populate_table(res.table_data)

        # 状态
        n = len(res.table_data)
        has_missing = any(fr.missing_required for fr in res.files)
        if n > 0:
            self._btn_export.setEnabled(True)
            extra = "；⚠ 部分行缺少必录值（红色标记）" if has_missing else ""
            self._log_line(f"✅ 导入完成：{n} 条记录{extra}")
        else:
            self._log_line("⚠ 无有效数据可展示。")

        # 启用映射配置并刷新状态
        self._btn_map.setEnabled(bool(res.files))
        self._update_map_status()

    def _on_import_fail(self, err: str) -> None:
        self._btn_import.setEnabled(True)
        self._progress.setText("")
        self._log_line(f"❌ 导入失败：{err}")
        utils.error("导入失败", f"处理过程中出错：{err}", parent=self)

    # ====== 映射状态 / 配置对话框 ======

    def _update_map_status(self) -> None:
        """刷新映射状态标签。"""
        if not self._result or not self._result.files:
            self._lbl_map_status.setText("尚未导入文件。")
            return
        names = [f.file_name for f in self._result.files]
        edited = [n for n in names if self._manual_full.get(n)]
        if edited:
            parts = []
            for n in edited:
                mp = self._manual_full[n]
                pos = sum(1 for s in mp.values() if s not in (None, ""))
                un = sum(1 for s in mp.values() if s in (None, ""))
                bits = []
                if pos:
                    bits.append(f"手工映射 {pos} 项")
                if un:
                    bits.append(f"显式不匹配 {un} 项")
                parts.append(f"{n}（{'、'.join(bits) or '已配置'}）")
            self._lbl_map_status.setText("已应用手工映射：" + "；".join(parts)
                                         + "。确定后按最终映射重新生成表格。")
            self._lbl_map_status.setStyleSheet("color: #16A34A;")
        else:
            self._lbl_map_status.setText(
                "当前使用自动匹配结果；如不满意可点「配置映射」手工调整。")
            self._lbl_map_status.setStyleSheet("color: #6B7280;")

    def _open_mapping_dialog(self) -> None:
        """打开字段映射配置对话框（基于本次导入结果）。"""
        if not self._result or not self._result.files:
            self._log_line("请先导入文件再配置映射。")
            return

        files_info: list[dict] = []
        for fr in self._result.files:
            # 纯自动匹配（不带任何手工覆盖），用于「恢复自动匹配」
            try:
                pure_auto, _u, _m = _logic.match_columns(
                    list(fr.source_headers), None)
            except Exception:
                pure_auto = []
            pure_auto_map = {m.target_field: m.source_header
                             for m in pure_auto}
            # 初始映射：已保存的手工映射优先，否则回退到纯自动匹配
            cur = self._manual_full.get(fr.file_name)
            if not cur:
                cur = pure_auto_map
            # 固定字段由系统处理，不进入对话框
            cur_filtered = {
                k: v for k, v in cur.items()
                if k not in _logic.FIXED_FIELD_HEADERS
            }
            files_info.append({
                "name": fr.file_name,
                "source_headers": list(fr.source_headers),
                "auto_map": pure_auto_map,
                "current_map": cur_filtered,
            })

        dlg = MappingDialog(self, files_info, _logic.TEMPLATE_FIELDS)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            result_maps = dlg.result_maps  # {name: {tgt: src_or_None}}
            for name, mp in result_maps.items():
                self._manual_full[name] = dict(mp)
            self._save_overrides()
            self._log_line("✓ 已保存映射配置，按最终映射重新生成表格…")
            # 重新生成（复用导入流程，应用新映射）
            self._do_import()

    # ====== 填充可编辑表格 ======

    def _populate_table(self, data: list[dict[str, object]]) -> None:
        """将清洗后的数据填入 QTableModel。

        必录字段值为空的单元格用红色背景标记。
        """
        self._model.clear()
        headers = _logic.TEMPLATE_HEADERS_DISPLAY  # 去掉 * 号的可读表头
        self._model.setColumnCount(len(headers))
        self._model.setHorizontalHeaderLabels(headers)
        self._model.setRowCount(len(data))

        # 固定字段（如 *背书明细/导入序号 留空）不参与必录红标判断
        required_set = _logic.REQUIRED_FOR_CHECK

        for r_idx, record in enumerate(data):
            for c_idx, tf_header in enumerate(_logic.TEMPLATE_HEADERS):
                val = record.get(tf_header)
                display = _logic.value_to_display(val)
                item = QStandardItem(display)
                item.setEditable(True)

                # 存原始值到 UserRole（排序/回写时使用）
                item.setData(val, Qt.ItemDataRole.UserRole)

                # 必录字段为空 → 标红
                is_required = tf_header in required_set
                is_empty = (val is None
                            or val == ""
                            or (isinstance(val, float) and val == 0
                                and tf_header not in {"*汇率", "*票面金额", "*票面利率(%%)"}))
                if is_required and is_empty:
                    item.setBackground(QColor("#FEF2F2"))  # 浅红底
                    item.setForeground(QColor("#DC2626"))     # 红字
                    item.setToolTip("⚠ 此必录字段缺少数据，请补充")

                self._model.setItem(r_idx, c_idx, item)

        self._table.resizeColumnsToContents()
        self._lbl_row_count.setText(f"{len(data)} 行")
        # 第一列（序号）宽度固定
        if len(headers) > 0:
            self._table.setColumnWidth(0, 50)

    # ====== 从表格收集当前数据（含用户编辑）======

    def _collect_table_data(self) -> list[dict[str, object]]:
        """从 QStandardItemModel 收集当前数据（包含用户手动编辑的内容）。"""
        rows = self._model.rowCount()
        cols = self._model.columnCount()
        result = []
        for r in range(rows):
            record: dict[str, object] = {}
            for c in range(cols):
                item = self._model.item(r, c)
                if item is not None:
                    raw = item.data(Qt.ItemDataRole.UserRole)
                    text = item.text().strip()
                    # 如果用户编辑过（text 与原始 display 不一致），以 text 为准
                    expected_display = _logic.value_to_display(raw) if raw is not None else ""
                    if text and text != expected_display:
                        # 用户编辑了，尝试反向解析
                        record[_logic.TEMPLATE_HEADERS[c]] = self._parse_cell_text(
                            text, _logic.TEMPLATE_HEADERS[c]
                        )
                    else:
                        record[_logic.TEMPLATE_HEADERS[c]] = raw
            result.append(record)
        return result

    @staticmethod
    def _parse_cell_text(text: str, field_name: str) -> object:
        """将用户在表格中输入的文字解析回对应 Python 类型。"""
        name = field_name.lstrip("*")
        if not text:
            return None
        # 浮点数
        if name in ("票面金额", "汇率", "票面利率(%)"):
            try:
                return float(text)
            except ValueError:
                return text
        # 布尔
        if name in ("期初票据", "带追索权", "可撤销", "电子票据"):
            if text.lower() in ("是", "true", "yes", "1", "√"):
                return True
            if text.lower() in ("否", "false", "no", "0", "×", "x"):
                return False
            return text
        # 整数序号
        if name == "导入序号":
            try:
                return int(float(text))
            except ValueError:
                return text
        # 日期 — 保持字符串让 clean_value 处理
        if name in ("签发日", "到期日", "收票日", "承兑日期",
                     "应收票据明细/背书日期"):
            return text
        return text

    # ====== 导出（写入模板）======

    def _do_export(self) -> None:
        if not self._table_data:
            self._log_line("没有可导出的数据。")
            return

        # 先从表格收集最新数据（含用户编辑）
        current_data = self._collect_table_data()

        out_path, _ = QFileDialog.getSaveFileName(
            self, "保存导入结果",
            str(app_config.DEFAULT_OUTPUT_DIR / f"应收票据导入_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"),
            "Excel 文件 (*.xlsx)",
        )
        if not out_path:
            return

        self._btn_export.setEnabled(False)
        self._progress.setText("正在导出…")
        self._log_line(f"正在导出到: {out_path} …")

        tpl = self._template_path

        self._future = self._executor.submit(
            _logic.write_to_template,
            current_data,
            Path(out_path),
            tpl,
            app_config.DATA_DIR,
        )

        def _check():
            if self._future.done():
                try:
                    n = self._future.result()
                    self._on_export_done(n, Path(out_path))
                except Exception as e:
                    self._on_export_fail(str(e))
                finally:
                    if self._poll_timer is not None:
                        self._poll_timer.stop()
                        self._poll_timer.deleteLater()
                        self._poll_timer = None

        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(100)
        self._poll_timer.timeout.connect(_check)
        self._poll_timer.start()

    def _on_export_done(self, n: int, path: Path) -> None:
        self._btn_export.setEnabled(True)
        self._progress.setText("")
        self._log_line(f"✅ 已导出 {n} 条记录 → {path}")
        utils.info(
            "导出完成", f"已生成: {path}\n共 {n} 条记录", parent=self,
        )

    def _on_export_fail(self, err: str) -> None:
        self._btn_export.setEnabled(True)
        self._progress.setText("")
        self._log_line(f"❌ 导出失败：{err}")
        utils.error("导出失败", f"写入 Excel 时出错：{err}", parent=self)


class NotesReceivableImportFeature(FeatureModule):
    """应收票据批量导入 — 功能模块入口。"""

    name = "应收票据导入"
    icon = "fa5s.file-import"

    def get_widget(self, parent: QWidget | None = None) -> QWidget:
        return NotesReceivableImportWidget(parent)