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
    └── js_bank_statement/
        ├── ui.py               # 界面（JsBankStmtFeature + JsBankStmtWidget）
        └── logic.py            # 纯业务逻辑（Excel 读取，保留日期/数字原始类型）
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

## 运行与测试

```bash
# 启动客户端（venv 已含依赖）
cd C:\Users\chong elaine\Desktop\ai-coding-test\my_fin_assistant
.\.venv\Scripts\python.exe main.py

# 无头烟测（无显示服务器时验证 import / 构建 / 切换，不验证真实渲染）
QT_QPA_PLATFORM=offscreen .\.venv\Scripts\python.exe -c "import main; print([f.name for f in main.FEATURES])"
```

> ⚠️ **已知限制**：`run_in_thread` 的 `QThread` 在无显示服务器（offscreen 无头）环境会段错误——这是环境限制，**真实桌面 `python main.py` 有显示服务时正常**。无头环境只验证 import 链、主窗口构建、模块切换，不要在无头环境跑真实 QThread 任务。

## 当前状态

- 已实现：`bank_classify`（银行承兑汇票白名单分类）、`js_bank_statement`（江苏银行对账单复制）
- 预留未实现：`features/invoice/`（发票模块，按 `bank_classify` 写法套用即可）
- 首页模块（home）已移除，功能通过侧边栏直接导航到各模块
