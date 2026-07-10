# 桌面端 UI 设计规范（Fluent 风格）

> 适用：本项目所有 `features/*/ui.py` 的桌面端（PySide6）界面。
> 参考实现（基准）：`features/notes_receivable_import/ui.py`（应收票据批量导入）。
> 已对齐实例：`features/bank_classify/ui.py`（票据分类，已按本规范改造）。

---

## 0. 一句话原则

- **单一真相来源**：颜色 / 字体 / 圆角 / 间距一律从 `core/theme.py` 取（`theme.COLOR` / `RADIUS` / `TYPE` / `SPACING` / `PAGE_PAD` / `CARD_GAP`），组件用 `objectName` 挂样式，**绝不散落硬编码颜色或尺寸字符串**。
- **UI / 逻辑严格分离**：`ui.py` 只管界面与交互；重计算、文件 IO、Excel 读写放 `*_logic.py`（零 UI 依赖、可单独测试）。
- **统一间距**：页面与卡片之间的距离由 `theme.PAGE_PAD` / `theme.CARD_GAP` 决定，不要写裸数字 `24` / `16`。

---

## 1. 页面骨架（`_setup_ui` 模板）

每个 `FeatureModule` 的 `get_widget()` 返回的 widget，在 `_setup_ui` 里这样搭：

```python
def _setup_ui(self) -> None:
    root = QVBoxLayout(self)
    root.setContentsMargins(theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD, theme.PAGE_PAD)
    root.setSpacing(theme.CARD_GAP)

    # 1) 页头
    title = QLabel("模块名称", self); title.setObjectName("pageTitle")
    desc = QLabel("一句话说明本页用途……", self)
    desc.setObjectName("pageDesc"); desc.setWordWrap(True)
    root.addWidget(title); root.addWidget(desc)
    root.addSpacing(theme.SPACING[4])

    # 2) 输入卡片 ……
    # 3) 操作按钮行 ……
    # 4) 日志面板（stretch=1 占底） ……
```

**关于滚动容器：**
- 页面根布局直接挂 `self`（`QVBoxLayout(self)`），**不要整页包 `QScrollArea`**。小窗不出现整页滚动条（这是 notes 的既定行为）。
- 若某区域内容很长需滚动（如表格、映射网格），**只在该区域内单独包 `QScrollArea`**（参考 notes 的预览表格 `_scroll`、映射网格），并 `stretch=1` 占该区域剩余空间。
- ⚠️ 不能「整页滚动容器」与「日志卡片 `stretch=1`」同时用，二者会互相冲突。

---

## 2. objectName 样式清单

以下 `objectName` 均由 `core/theme.py` 的 `BUILD_QSS` 统一渲染，**按名字设置即可，不要自己写 QSS**：

| objectName | 适用组件 | 含义 / 样式 |
|---|---|---|
| `pageTitle` | `QLabel` | 页标题（大号、加粗） |
| `pageDesc` | `QLabel` | 页描述 / 提示性小字（次要色、可 `setWordWrap`） |
| `card` | `QFrame` | 卡片容器（浅色底、圆角、细边框、轻阴影） |
| `cardTitle` | `QLabel` | 卡片标题（中号、加粗） |
| `primary` | `QPushButton` | 主操作按钮（主题蓝填充、白字、加粗） |
| `logView` | `QPlainTextEdit` | 只读日志 / 文本区（等宽底色、圆角） |

> 次级小按钮（如「切换」「浏览」）**不加** `primary`，保持默认白边样式即可。

---

## 3. 间距常量（来自 `theme`）

| 用途 | 写法 | 值 |
|---|---|---|
| 页边到卡片留白 | `theme.PAGE_PAD` | `SPACING[24]` |
| 卡片之间垂直间距 | `theme.CARD_GAP` | `SPACING[16]` |
| 卡片**内部**边距 | `setContentsMargins(SPACING[16]…)` | 16 |
| 卡片**内部**子项间距 | `setSpacing(SPACING[12])` | 12 |
| 日志卡片**内部**边距 | `setContentsMargins(SPACING[12]…)` | 12 |
| 日志卡片**内部**子项间距 | `setSpacing(6)` | 6 |
| 页头到首个卡片 | `root.addSpacing(SPACING[4])` | 4 |

---

## 4. 按钮规范

| 类型 | objectName | 高度 | 最小宽 | 备注 |
|---|---|---|---|---|
| 主操作（导入 / 导出 / 预览 / 生成…） | `primary` | 34 | 120 | 同一行并排，行尾 `addStretch(1)` **左对齐** |
| 次级小按钮（切换 / 浏览…） | 无 | 30 | 72 | 默认白边；同类并排时同样 `addStretch(1)` |
| 对话框确认 / 取消 | 无（确认可加 `primary`） | 32 | — | 用在 `QDialog` 底部 |
| 页脚大操作（如预览） | `primary` | 38 | — | 独占一行时可用更高 |

**按钮行范式**（主按钮并排、左对齐）：

```python
h_btn = QHBoxLayout(); h_btn.setSpacing(theme.SPACING[12])

btn_add = QPushButton("添加文件", card); btn_add.setObjectName("primary")
btn_add.setFixedHeight(34); btn_add.setMinimumWidth(120)
btn_add.clicked.connect(self._add_files)
h_btn.addWidget(btn_add)

# ……其余主按钮（导出 / 清空 等）同样 objectName("primary") + h34 + minW120

h_btn.addStretch(1)            # 关键：左对齐，右侧留白
card_lay.addLayout(h_btn)
```

---

## 5. 卡片规范

```python
card = QFrame(self); card.setObjectName("card")
lay = QVBoxLayout(card)
lay.setContentsMargins(theme.SPACING[16], theme.SPACING[16],
                       theme.SPACING[16], theme.SPACING[16])
lay.setSpacing(theme.SPACING[12])

head = QLabel("卡片标题", card); head.setObjectName("cardTitle")
lay.addWidget(head)
# ……卡片内容（输入控件 / 表格 / 提示）
root.addWidget(card)            # 普通卡片：不加 stretch
```

- **双列并排卡片**：用 `QHBoxLayout` 包两个 `card`，再 `root.addLayout(row)`。
- ⚠️ **固定高度输入区（常见坑）**：`QPlainTextEdit` 默认垂直 `sizePolicy` 是 `Expanding`，如果不加限制会被卡片剩余空间撑开、产生大片空白。固定高度输入区必须：
  ```python
  self._txt.setMinimumHeight(140)
  self._txt.setMaximumHeight(140)
  self._txt.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
  ```
  （参考 bank_classify 的银行全称输入区。）

---

## 6. 日志面板（照搬 notes）

```python
log_card = QFrame(self); log_card.setObjectName("card")
lc_lay = QVBoxLayout(log_card)
lc_lay.setContentsMargins(theme.SPACING[12], theme.SPACING[12],
                          theme.SPACING[12], theme.SPACING[12])
lc_lay.setSpacing(6)

log_title = QLabel("运行日志（操作详情）", log_card)   # 标题带说明
log_title.setObjectName("cardTitle")
lc_lay.addWidget(log_title)

self._log = QPlainTextEdit(log_card)
self._log.setObjectName("logView")
self._log.setReadOnly(True)
self._log.setMinimumHeight(50)
lc_lay.addWidget(self._log, stretch=1)     # 日志区填满卡片剩余高度

root.addWidget(log_card, stretch=1)         # 卡片占页面剩余纵向空间（贴底）
```

**格式约定（状态图标前缀）**：

| 图标 | 含义 | 用法示例 |
|---|---|---|
| `✅` | 成功 | 完成、已生成、已导出 |
| `⚠` | 提醒 | 未检测到输入、剪贴板为空、已取消 |
| `❌` | 失败 | 加载失败、生成失败 |
| `ℹ` | 信息 | 开始处理、明细、已粘贴 / 已加载 / 已清空 |

**行为**：
- 开始处理前先 `self._log.clear()`，再写「ℹ 开始处理……」（对齐 notes 导入前清空日志的行为）。
- 完成后追加「ℹ 分类 / 匹配明细：」段落，逐条列出关键结果（对齐 notes 的「列匹配详情」）。
- 写日志统一走一个 `_log_line(msg)` 方法（必要时再包一层状态前缀）。

---

## 7. 工作流页通用结构（以 notes 导入页为例）

1. **页头**：`pageTitle` + `pageDesc`（一句话说明用途）
2. **输入卡片**：拖拽区 / 添加文件 / 参数填写
3. **操作按钮行**：主按钮并排、左对齐（`addStretch(1)`）
4. **状态 / 计数提示行**：`pageDesc` 小字（如「已匹配 N 条」）
5. **预览区**（如有表格）：内部 `QScrollArea`，`stretch=1` 占剩余空间
6. **日志面板**：`stretch=1` 贴底

---

## 8. 反例（不要做）

- ❌ 散落硬编码颜色 / 字体 / 尺寸（如 `setStyleSheet("background:#1f6feb")`、`font-size:14px`、裸数字 `setSpacing(12)`）。
- ❌ 整页包 `QScrollArea` 又让日志卡片 `stretch=1`——二者冲突。
- ❌ 主按钮不统一：有的加 `primary`、有的不加，高度 / 宽度各异。
- ❌ 卡片间距写裸数字而非 `theme.CARD_GAP`。
- ❌ 在 `ui.py` 里做重计算或文件 IO（应放 `*_logic.py`）。
- ❌ 日志区不设最小高度 / 不设 `stretch`，导致内容被压缩看不清。
- ❌ 固定高度文本输入区只设 `setMinimumHeight` 不设 `setMaximumHeight` + `Fixed`，被默认 `Expanding` 撑开留白。

---

## 9. 交付前验证清单

- [ ] `import main` 正常加载全部 `FEATURES`
- [ ] 各页面根布局统一用 `theme.PAGE_PAD` / `theme.CARD_GAP`
- [ ] 主按钮均为 `primary` + `setFixedHeight(34)` + `setMinimumWidth(120)` 并排
- [ ] 日志面板 `stretch=1`、标题带说明、状态图标前缀齐全
- [ ] 无硬编码颜色 / 尺寸（grep `setStyleSheet` 与裸数字间距应为空或仅引用 `theme`）
- [ ] 固定高度输入区已加 `setMaximumHeight` + `QSizePolicy.Fixed`
