# 项目长期记忆 — my_fin_assistant

## 项目定位
个人财务助手桌面应用（**PySide6 + qtawesome，Fluent 风格**）。目录分层架构（来自用户提供的两张规范截图）：
- `main.py` — 入口，只做启动（`theme.apply_theme(app)` + FEATURES 注册表 + QApplication + MainWindow）
- `core/` — 应用骨架（纯壳，不写业务逻辑）：
  - `feature_base.py` → **`FeatureModule(ABC)`**：统一接口 `name` / `icon`(qtawesome 名) / `get_widget(parent) -> QWidget`
  - `main_window.py` → 可折叠侧边栏（QFrame#sideBar + QListWidget#sideNav）+ 内容区 QStackedWidget，**FEATURES 列表装载**，首页卡片 `feature_requested` 信号按 name 跳转
  - `theme.py` → Fluent 设计系统单一真相来源（颜色/间距/圆角/字号 token + 完整 QSS `BUILD_QSS` + `nav_icon()` + `apply_theme()`）
  - `worker.py` → QThread+Signal 后台任务封装 `run_in_thread`/`Worker`（started/progress/finished/failed）
  - `app_config.py` / `utils.py`
- `features/<模块>/ui.py` — 继承 FeatureModule，暴露 `get_widget()`；**UI 和逻辑严格分离**
- `features/<模块>/logic.py` 或 `*_logic.py` — 纯计算/数据处理，**不依赖 UI，可单独测试**
- `data/` — 本地数据/缓存

### 关键设计规则（来自图片规范）
1. 主窗口入口做"壳"，不写任何业务逻辑
2. 每个功能模块暴露统一接口（FeatureModule），主窗口靠 get_widget 加载，**不用互相 import 细节**
3. 主窗口用 **FEATURES 注册表** 统一装载，加新功能只需加一行
4. UI 和逻辑严格分离：ui.py 管界面交互，logic.py 放纯计算
5. 视觉由 `core/theme.py` 单一真相来源驱动，**不**用外部 `resources/styles.qss`、不散落样式字符串

## 已有功能模块
- `features/home/` — 首页（Fluent 欢迎页 + 功能卡片网格），已接入主窗口，HomeFeature(FeatureModule) → HomeWidget，含 `feature_requested` 跳转信号
- `features/bank_classify/` — 银行承兑汇票白名单分类（21家银行）。源自 `ai-coding-test/21票据`。**已 PySide6 Fluent 化**：BankClassifyFeature(FeatureModule) → BankClassifyWidget(生成 Excel 走 `run_in_thread` 后台线程)；逻辑层 `classify_logic.py` 纯 Python 零依赖 UI
- `features/invoice/` — 目录存在但空，待实现（PySide6 FeatureModule）

## 运行环境
- venv：`my_fin_assistant/.venv/Scripts/python.exe`
- 已装包：openpyxl 3.1.5、PySide6 6.11.1、qtawesome
- 启动：`python main.py`；无头验证用 `QT_QPA_PLATFORM=offscreen`（注意：offscreen 下 QThread 会段错误，仅主线程逻辑可在无头验证）

## 约定
- 功能模块用相对导入（`from . import classify_logic as _logic`）；跨级走 `from core import app_config, utils, theme`
- **Feature 注册**：在 `main.py` 的 **FEATURES 列表**里追加一行实例化，try/except 容错，注释掉即不加载互不影响
- 数据文件放 `data/`，路径走 `app_config.DATA_DIR` 等常量，不硬编码
- **icon 是 qtawesome 图标名**（如 "fa5s.home"、"fa5s.file-excel"），由 `theme.nav_icon()` 渲染，不用 resources/icons/ 文件
- 长耗时操作必须走 `core.worker.run_in_thread`，不阻塞 UI 线程
- 中文环境；股票涨用红色、跌用绿色（中国习惯）

## 待办
- features/invoice 待实现（PySide6 FeatureModule，参考 bank_classify 写法）
- 真实桌面运行 `python main.py` 验证 Fluent 视觉效果（折叠动画、卡片、日志面板）
