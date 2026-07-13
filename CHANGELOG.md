# 更新日志

本文件记录「我的财务助手」各版本的重要改动。

---

## v1.1.0 — 2026-07-13

### 新增 / 改进
- **打包为独立 .exe**：提供 `build.spec`（PyInstaller onefile），可一键产出 `dist/我的财务助手.exe`（约 59 MB，双击即用）。
  - 首次运行会在 exe 同级目录自动生成 `data/`、`output/`，并释放内置 `应收票据导入模版.xlsx` 与 `banks.txt`；用户已有的 `mapping_overrides.json`（已保存的映射）不会被覆盖。

### 修复
- **【应收票据导入】修复「导入后无数据 / 无法配置映射」**
  - 现象：导入部分银行/票交所导出的 xlsx（如 `银票客户持票结果查询01202607130076000149710726.xlsx`）后，表格空白、映射对话框里没有来源列可选。
  - 根因：这类文件 sheet 的 `<dimension ref="A1"/>` 属性被导出系统写错，而应用原先用 `openpyxl(read_only=True)` 读取、该模式直接信任 `dimension` 属性 → 误判为 1 行 1 列。
  - 修复：`read_source_file` 增加回退机制——当 `read_only` 模式探测到尺寸退化（≤1 行且 ≤1 列）或解析异常时，自动改用完整加载（`_read_full`）按真实单元格重算尺寸。修复后该文件正确读出 24 列 / 13 条数据，并自动匹配 6 列（票据号、签发日、到期日、票面金额、出票人、承兑人）。
  - 新增回归测试 `tests/test_read_broken_dimension.py`（构造 `dimension='A1'` 损坏 xlsx，断言能恢复）。

- **【Word 批量文本替换】修复校验分支崩溃**
  - 原 UI 中 `utils.warning` 应改为 `utils.warn`（`core.utils` 实际只暴露 `warn`/`info`/`error`/`confirm`），否则校验失败时会 `AttributeError` 崩溃。

### 变更
- **【Word 批量文本替换】替换引擎由 python-docx 改为 Word COM 自动化（win32com）**
  - 遍历 `doc.StoryRanges` 全故事范围（正文 / 表格 / 页眉 / 页脚 / 文本框），覆盖更全、更保格式。
  - `win32com` 惰性导入；`process_folder` 复用同一 Word 实例，在后台线程内 `CoInitialize`。
  - 仅在 Windows + 已安装 Microsoft Word 时可用（无 Word 环境下替换类用例自动跳过）。
  - `requirements.txt` 增加 `pywin32` 依赖。

### 构建 / 依赖
- `build.spec` 补充 `hiddenimports`（含 `win32com` / `pythoncom` / `pywintypes` 等），避免 PyInstaller 静态扫描漏掉惰性导入导致打包后 Word 替换报 `No module named win32com`。

---

## v1.0.0 — 早期版本（基线）

- 个人财务桌面客户端（PySide6 + qtawesome，Fluent 风格）。
- 功能模块：首页、银行承兑汇票白名单分类、江苏银行对账单复制、应收票据批量导入、Word 批量文本替换。
- 统一 `FeatureModule` 接口 + `FEATURES` 注册表即插即用；UI 与逻辑严格分离（`<模块>_ui.py` / `*_logic.py`）。
