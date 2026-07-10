# AGENTS.md

本文件供 AI 编码代理（agent）阅读，描述 `my_fin_assistant` 项目结构与开发约定。
人类用户看 [README.md](./README.md)。

## 项目概述

个人财务桌面客户端。架构核心原则（务必遵守）：

1. **主窗口做"壳"，不写业务逻辑。** `core/main_window.py` 只负责导航（侧边栏 + 内容区 `QStackedWidget`）与模块装载，绝不放具体功能代码。
2. **统一接口 `FeatureModule`。** 每个功能模块暴露一个继承 `core.feature_base.FeatureModule` 的类，主窗口只通过 `get_widget(parent)` 装载，**不 import 模块内部细节**。
3. **`FEATURES` 注册表统一装载。** 加新功能只在 `main.py` 追加一行实例化（用 `try/except` 包裹），主窗口逻辑无需改动。
4. **UI 与逻辑严格分离。** `features/<模块>/ui.py` 管界面交互；`features/<模块>/*_logic.py` 放纯计算/数据处理，**不依赖 UI、可单独测试**。

## 技术栈

- **PySide6 6.11.1**（GUI，非 PyQt5）
- **qtawesome 1.x**（侧边栏/按钮图标，用图标名字符串，不引入 SVG 文件）
- **openpyxl 3.1.5**（Excel 读写）
- Python 3.10+

> ⚠️ **不要引入 `QFluentWidgets`**。其许可为 GPLv3，与本项目不兼容。本项目用自定义 `QMainWindow` + qtawesome 实现 Fluent 风格（参考 `pyside6_fluent_gui_design` skill 的 `references/`）。

## 目录结构

```
my_fin_assistant/
├── main.py                     # 入口：QApplication + apply_theme + FEATURES 注册表
├── core/                       # 应用骨架（纯壳）
│   ├── app_config.py           # 路径常量(BASE_DIR/DATA_DIR/...) + APP_NAME + ensure_dirs()
│   ├── feature_base.py         # FeatureModule(ABC)：name / icon / get_widget
│   ├── main_window.py          # MainWindow：可折叠侧边栏 + 内容区
│   ├── theme.py                # Fluent 设计系统（COLOR/SPACING/... + apply_theme + nav_icon）
│   ├── utils.py                # info/warn/error/confirm 消息框封装
│   └── worker.py               # run_in_thread / run_blocking 后台任务
├── data/                       # 本地数据（banks.txt）
├── tests/                      # 测试文件（无头烟测 / 剪贴板）
└── features/
    # 首页模块已移除；功能由侧栏直接导航到各模块
    ├── bank_classify/
    │   ├── ui.py               # 界面（BankClassifyFeature + BankClassifyWidget）
    │   └── classify_logic.py   # 纯业务逻辑（白名单匹配/加载/写出 Excel）
    ├── js_bank_statement/
    │   ├── ui.py               # 界面（JsBankStmtFeature + JsBankStmtWidget）
    │   └── logic.py            # 纯业务逻辑（Excel 读取，保留日期/数字原始类型）
    └── notes_receivable_import/
        ├── ui.py               # 界面（NotesReceivableImportFeature + NotesReceivableImportWidget）
        │                       #   + MappingDialog（映射配置对话框，下拉+冲突提示+持久化）
        └── import_logic.py     # 纯业务逻辑（多来源读取/列名匹配/清洗/写入模板）
```

## 核心 API 速查

**`core.feature_base.FeatureModule`（抽象基类）** — 所有功能模块的实现契约：
- `name: str` — 英文唯一标识（用于注册表与首页卡片跳转）
- `icon: str` — qtawesome 图标名（如 `'fa5s.chart-pie'`，见 `theme.icon_name` 规整规则）
- `get_widget(self, parent=None) -> QWidget` — 返回该功能的根 widget

**`core.theme`** — 设计系统单一真相来源：
- `apply_theme(window)` — 把 QSS 应用到整个窗口树（在 `main.py` 调用）
- `nav_icon(name, active=False)` — 返回 qtawesome 图标对象
- `COLOR / SPACING / RADIUS / TYPE / ICON / SIDEBAR` — 设计 token 字典
- ⚠️ 组件需要局部覆盖样式时才写内联 QSS；全局颜色/间距一律从这里取，**不要散落硬编码颜色字符串**

**`core.worker`** — 长任务后台执行：
- `run_in_thread(fn, *args, on_progress=, on_finished=, on_failed=, parent=) -> QThread`
  调用方必须持有返回的 `QThread` 引用（挂到 `self` 防止被回收）。
- `run_blocking(fn, timeout_ms=)` — **仅供测试**，UI 代码禁止使用。

**`core.app_config`** — 路径常量（`BASE_DIR` / `DATA_DIR` / `DEFAULT_OUTPUT_DIR` 等），跨模块路径从这里取，不要硬编码绝对路径。

## 开发约定（DO / DON'T）

**DO**
- 功能模块用相对导入：`from . import classify_logic as logic`
- 跨级导入：`from core import app_config, theme`
- 功能模块的 `icon` 前缀固定用 `fa5s.`（Solid 风格），如 `fa5s.file-excel`、`fa5s.chart-pie`
- 长任务（读写大文件、生成 Excel）走 `run_in_thread`，开始时禁用主操作按钮、结束/失败恢复并给出摘要
- 每页布局遵循「页头(title+desc) / 输入卡片 / 操作」工作流
- 处理空态 / 加载态 / 错误态 / 成功态
- Ctrl+C 等快捷键行为由自定义 `QTableView` 子类 `_TableView` 重写 `keyPressEvent` 实现，避免默认单格复制的限制
- 复制/排序场景中，原始值存 `QStandardItem.UserRole`，不从 `QModelIndex.row()` 反查原始数组，防止排序后错位

**DON'T**
- ❌ 在 `core/main_window.py` 写任何业务逻辑
- ❌ 在 `ui.py` 里直接做重计算或文件 IO（放到 `*_logic.py`）
- ❌ 硬编码颜色、字体、尺寸（用 `theme.COLOR` 等）
- ❌ 硬编码路径（用 `app_config.DATA_DIR` 等）
- ❌ 引入 QFluentWidgets 或其他 GPL 依赖
- ❌ 在功能模块里 `import main`（会造成循环依赖）

## 如何新增一个功能模块

1. 建 `features/<模块>/ui.py`（界面，定义 `XxxFeature(FeatureModule)` + `XxxWidget(QWidget)`）与 `features/<模块>/<模块>_logic.py`（纯函数）。
2. `XxxFeature` 实现 `name` / `icon` / `get_widget`。
3. 在 `main.py` 的 `FEATURES` 注册表追加：
   ```python
   try:
       from features.<模块>.ui import XxxFeature
       FEATURES.append(XxxFeature())
   except Exception as e:
       print(f"[warn] 跳过 features.<模块>: {e}", file=sys.stderr)
   ```
4. 模块即插即用：在 `main.py` 注册表追加后侧边栏自动出现，无需改主窗口代码。
   > 首页卡片入口（`features/home/feature.py`）已移除，加新功能直接走侧边栏导航。

## 功能模块说明

### `features/notes_receivable_import/`（应收票据批量导入）

**主流程**：拖拽来源 Excel → 自动列名匹配（关键字规则）→「配置映射」(可选) → 表格展示 → 手动编辑 → 导出模板。

**接口契约**（`import_logic.py`，纯 Python、零 UI 依赖）：
- `TEMPLATE_FIELDS` — 模板字段定义列表（`TemplateField`：`header` / `required` / `aliases`），从模板表头（`TEMPLATE_HEADERS`）解析而来。
- `FIXED_FIELDS` — 固定字段注册表，**不出现在映射对话框**，始终按规则自动填充：
  | 字段 | 规则 |
  |------|------|
  | `*导入序号` | 自动顺序编号 (1, 2, 3…) |
  | `*票据类型` | 固定常量「银行承兑汇票」 |
  | `*币别` | 固定常量「人民币」 |
  | `*背书明细/导入序号` | 留空 |
  | `*汇率` | 固定常量 `1.0` |
  | `*付款单位多资料值` | 固定常量「客户」 |
  > 固定字段优先级最高：即便来源文件有同名列（如「币别=美元」「汇率=7.2」），也会被覆盖。
- `match_columns(source_headers, manual_map=None, excluded_targets=None)` → `(matched, unmatched_source, missing_required)`
  - **自动匹配仅针对必录字段**；非必录字段只有显式写入 `manual_map` 才会匹配，否则默认留空（避免「打开对话框点确定反而清空可选字段」）。
  - 跳过 `FIXED_FIELDS` 与 `excluded_targets`。
  - `missing_required` 已排除固定字段。
- `read_and_transform(file_paths, defaults=None, manual_maps=None, excluded_maps=None)` → `ImportResult`
  - `manual_maps`: `{文件名: {来源列: 模板字段}}`（用户显式映射/覆盖）。
  - `excluded_maps`: `{文件名: [模板字段…]}`（用户显式「不匹配」的字段，不参与自动匹配）。
  - `FileResult.source_headers` 供对话框下拉使用。
- `process_files(file_paths, output_path, template_path=None, data_dir=None, manual_maps=None, defaults=None, excluded_maps=None)` → 完整流程并写盘。

**UI 约定**（`ui.py`）：
- `NotesReceivableImportWidget` 继承 `FeatureModule` + `QWidget`；`_do_import` 由持久化映射 `self._manual_full` 推导 `positive`/`excluded` 传入 `read_and_transform`。
- `MappingDialog(QDialog)`：左列模板字段（必录 `*` 标红）、右列来源列下拉（来源实际列名 + 哨兵 `〈不匹配/固定值〉`）；默认值 = 自动匹配结果；**必录行排在最前、非必录在后**；多文件顶部下拉切换；实时冲突检测（重复来源列红框 + 必录未匹配橙框 + 底部警告条）；含「恢复自动匹配 / 确定 / 取消」。
- **映射持久化**：确认后写入 `data/mapping_overrides.json`（`{文件名: {模板字段: 来源列 或 null}}`）；启动时 `NotesReceivableImportWidget.__init__` 自动召回，日志透明提示「已载入已保存映射」。
- 固定字段全程从对话框行、`_editing`、`result_maps`、`_load_overrides` 中剔除，也不会出现在默认值卡片。

**测试**：`tests/test_notes_mapping.py`（offscreen 无头）覆盖手工覆盖+显式不匹配、对话框冲突/必录提示/恢复自动/结果结构、持久化回载。

## 运行与测试

```bash
# 启动客户端（venv 已含依赖）
cd C:\Users\chong elaine\Desktop\ai-coding-test\my_fin_assistant
.\.venv\Scripts\python.exe main.py

# 无头烟测（无显示服务器时验证 import / 构建 / 切换，不验证真实渲染）
QT_QPA_PLATFORM=offscreen .\.venv\Scripts\python.exe -c "import main; print([f.name for f in main.FEATURES])"
```

> ⚠️ **已知限制**：`run_in_thread` 的 `QThread` 在无显示服务器（offscreen 无头）环境会段错误——这是环境限制，**真实桌面 `python main.py` 有显示服务时正常**。无头环境只验证 import 链、主窗口构建、模块切换，不要在无头环境跑真实 QThread 任务。

## Web 版架构 (React + pywebview)

项目同时提供 **PySide6 原生版** 和 **Web 版** 两种前端，共享同一套纯 Python 业务逻辑层。

### 目录结构

```
web/                          # Vite + React 前端
├── index.html               # Vite 入口 HTML
├── package.json             # React 18 + Vite 6
├── vite.config.js           # 构建配置 → 输出到 dist/
└── src/
    ├── main.jsx             # React 挂载点
    ├── App.jsx              # 主组件（侧边栏 + 路由 + 3 个功能视图）
    └── styles.css           # 设计系统（light/dark token + 组件样式）

webview_api.py               # Python API 层（暴露给 JS 的所有方法）
webview_main.py              # 生产模式启动器（加载 web/dist/）
run_web.py                   # 统一入口（支持 --dev 开发模式）
```

### 架构分层

```
┌─────────────────────────────────────┐
│  React 18 (Vite 构建)              │  ← 界面、交互、状态展示
│  App.jsx / styles.css              │     内联 SVG 图标，无外部依赖
├─────────────────────────────────────┤
│  window.pywebview.api.*            │  ← pywebview 桥接（自动序列化）
├─────────────────────────────────────┤
│  webview_api.py (WebAPI)           │  ← 文件对话框 / 异步线程 / 映射持久化
│  → features/*_logic.py             │  ← 纯业务逻辑（零 UI 依赖）
└─────────────────────────────────────┘
```

### 启动方式

```bash
# PySide6 原生版（不变）
python main.py

# Web 版 — 生产模式（需先 npm run build）
python webview_main.py

# Web 版 — 开发模式（连接 vite dev server，热更新）
python run_web.py --dev

# 构建前端
cd web && npm install && npm run build
```

### 开发约定

- **React 视图只管 UI**：调用 `window.pywebview.api.*` 获取数据/触发操作，不直接 import Python
- **API 方法返回 dict 或 list**：JSON 可序列化；出错返回 `{"error": "..."}` 不抛异常
- **重活 async 化**：Excel 读写走 `asyncio.to_thread()`，UI 不卡顿
- **图标用内联 SVG**：`Icons = { home: <svg>...</svg>, ... }`，不用 emoji、不用外部库
- **设计系统 CSS 变量**：`--primary`, `--bg`, `--surface-*`, `--text-*` 等；支持 `[data-theme="dark"]` 切换

## 当前状态

- 已实现：`bank_classify`（银行承兑汇票白名单分类）、`js_bank_statement`（江苏银行对账单复制）、`notes_receivable_import`（应收票据批量导入，含「配置映射」对话框 + 固定字段自动填充 + 映射持久化）

## UI 设计规范（Fluent 风格）

所有功能模块的桌面端 UI 必须遵循统一设计语言，以 `features/notes_receivable_import/ui.py`（应收票据导入）为参考实现。**新增或修改界面前，请先读取 [`docs/fluent_ui_standard.md`](./docs/fluent_ui_standard.md)**——其中包含页面骨架模板、objectName 样式清单、间距常量、按钮 / 卡片 / 日志面板规范、反例与交付前验证清单，可直接照着写。
- 预留未实现：`features/invoice/`（发票模块，按 `bank_classify` 写法套用即可）
- 首页模块（home）已移除，功能通过侧边栏直接导航到各模块
