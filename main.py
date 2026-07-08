#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""my_fin_assistant — 应用入口（PySide6）。

主窗口入口做"壳"，不写任何业务逻辑。
主窗口用一个注册表统一装载，加新功能只需要在这里加一行，不用改主窗口逻辑。
"""
from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from core import app_config, theme
from core.main_window import MainWindow

# ===== 功能模块注册表 =====
# 加新功能只需在此追加一行；注释掉即不加载，互不影响。
# 每个功能模块通过 ui.py 暴露一个继承 FeatureModule 的类。
FEATURES: list = []
# Home feature removed per user request. To restore, uncomment the import
# and append lines below.
# try:
#     from features.home.feature import HomeFeature
#     FEATURES.append(HomeFeature())
# except Exception as e:  # noqa: BLE001
#     print(f"[warn] 跳过 features.home: {e}", file=sys.stderr)

try:
    from features.bank_classify.ui import BankClassifyFeature
    FEATURES.append(BankClassifyFeature())
except Exception as e:  # noqa: BLE001
    print(f"[warn] 跳过 features.bank_classify: {e}", file=sys.stderr)

try:
    from features.js_bank_statement.ui import JsBankStmtFeature
    FEATURES.append(JsBankStmtFeature())
except Exception as e:  # noqa: BLE001
    print(f"[warn] 跳过 features.js_bank_statement: {e}", file=sys.stderr)

# try:
#     from features.invoice.ui import InvoiceFeature
#     FEATURES.append(InvoiceFeature())
# except Exception as e:
#     print(f"[warn] 跳过 features.invoice: {e}", file=sys.stderr)


def main() -> int:
    app_config.ensure_dirs()

    app = QApplication(sys.argv)
    app.setApplicationName(app_config.APP_NAME)
    app.setApplicationVersion(app_config.APP_VERSION)
    app.setOrganizationName(app_config.ORG_NAME)

    # 设计系统：给整个应用应用 Fluent 主题（单一真相来源见 core/theme.py）
    theme.apply_theme(app)

    win = MainWindow(features=FEATURES)
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
