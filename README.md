# 我的财务助手 (my_fin_assistant)

个人财务桌面客户端。基于 **PySide6 + qtawesome** 的 Fluent 风格界面，采用「主窗口壳 + 功能模块注册表」的分层架构：主窗口只做导航与装载，业务逻辑与界面严格分离。

## ✨ 功能特性

| 功能 | 说明 |
|---|---|
| **银行承兑汇票分类** | 加载 Excel，按白名单银行规则自动分类/标记，导出分类结果 |
| **江苏银行对账单复制** | 加载对账单 Excel，表格查看/排序，支持拖拽导入、框选复制到剪贴板（可直接粘贴到 Excel） |
| **应收票据批量导入** | 多来源 Excel 读取 → 列名模糊匹配（33 字段模板）→ 可视化「配置映射」确认 → 数据清洗 → 写入固定模板导出，映射规则本地持久化 |
| **Word 批量文本替换** | 选择输入/输出文件夹，配置多行「查找→替换」规则，可含子文件夹，批量处理 .docx（原文件不受影响） |

## 📁 目录结构

```
my_fin_assistant/
├── main.py                     # 应用入口：QApplication + 应用主题 + FEATURES 注册表
├── core/                       # 应用骨架（纯壳，不写业务逻辑）
│   ├── app_config.py           # 路径常量 / 应用元信息
│   ├── feature_base.py         # FeatureModule 抽象基类（统一接口 get_widget）
│   ├── main_window.py          # 主窗口壳：侧边栏 + 内容区 QStackedWidget
│   ├── theme.py                # Fluent 设计系统（token + QSS + 图标助手）
│   ├── utils.py                # 消息框封装
│   └── worker.py               # 后台线程任务封装
├── data/                       # 本地数据 / 缓存
│   └── banks.txt               # 银行白名单样本数据
├── tests/                      # 测试文件
│   ├── test_clipboard.py
│   ├── test_direct_thread.py
│   ├── test_notes_mapping.py
│   ├── test_run_generate.py
│   └── test_word_text_replacer.py
└── features/                   # 功能模块（UI 与逻辑分离）
    ├── bank_classify/          # 银行承兑汇票分类
    │   ├── bank_classify_ui.py            # 界面（BankClassifyFeature + BankClassifyWidget）
    │   └── classify_logic.py              # 纯业务逻辑（白名单匹配 / 加载 / 写出 Excel）
    ├── js_bank_statement/      # 江苏银行对账单复制
    │   ├── js_bank_statement_ui.py        # 界面（JsBankStmtFeature + JsBankStmtWidget）
    │   └── logic.py                       # 纯业务逻辑（Excel 读取，保留日期/数字原始类型）
    ├── notes_receivable_import/ # 应收票据批量导入
    │   ├── notes_receivable_import_ui.py  # 界面（含 MappingDialog 映射配置对话框）
    │   └── import_logic.py                # 纯业务逻辑（多来源读取 / 列名匹配 / 清洗 / 写模板）
    └── word_text_replacer/     # Word 批量文本替换
        ├── word_text_replacer_ui.py       # 界面（WordTextReplacerFeature + Widget）
        └── word_text_replacer_logic.py    # 纯业务逻辑（扫描 / 段落替换 / 批量处理）
```

## 🔧 环境要求

- Python 3.10+
- Windows / macOS / Linux（GUI 需有显示服务）

## 📦 安装与运行

依赖已装在项目自带的虚拟环境 `.venv` 中（PySide6、qtawesome、openpyxl、python-docx、pyperclip、xlrd 等，详见 `requirements.txt`）。

```bash
# 1. 进入项目目录
cd "C:\Users\chong elaine\Desktop\ai-coding-test\my_fin_assistant"

# 2. 启动客户端（推荐，直接走 venv 的 python）
.\.venv\Scripts\python.exe main.py

# 或先激活 venv 再启动
.\.venv\Scripts\activate
python main.py
```

若需要重新安装依赖：

```bash
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 🧩 新增一个功能模块

遵循「UI 与逻辑分离」约定：

1. 在 `features/<你的模块>/` 下建 `<模块>_ui.py`（界面）与 `*_logic.py`（业务逻辑，纯函数、不依赖 UI）。
2. `<模块>_ui.py` 中定义一个继承 `core.feature_base.FeatureModule` 的类，实现：
   - `name`（英文标识，用于跳转与注册）
   - `icon`（qtawesome 图标名，如 `'fa5s.chart-pie'`）
   - `get_widget(self, parent=None) -> QWidget`
3. 在 `main.py` 的 `FEATURES` 注册表里追加一行（用 `try/except` 包裹，模块挂了不影响整体启动）：

```python
try:
    from features.invoice.invoice_ui import InvoiceFeature
    FEATURES.append(InvoiceFeature())
except Exception as e:
    print(f"[warn] 跳过 features.invoice: {e}", file=sys.stderr)
```

> 加新功能只改 `main.py` 一行，主窗口逻辑无需改动。

## 🎨 设计系统

所有视觉 token（颜色、间距、圆角、字体、图标尺寸）与 QSS 样式集中定义在 `core/theme.py`，是界面外观的单一真相来源。修改全局风格只需改这一处。

长任务通过 `core/worker.py` 的 `run_in_thread(fn, on_finished=..., on_failed=..., parent=...)` 在后台线程执行，避免界面假死。

## 📝 备注

- 本项目由 `21票据` 独立脚本项目（tkinter + openpyxl）按分层架构迁入并重写成 PySide6 桌面客户端。
- 首页模块（home）已移除，功能通过侧边栏直接导航到各模块。
