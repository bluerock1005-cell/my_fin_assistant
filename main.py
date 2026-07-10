#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""my_fin_assistant — 应用入口（PySide6）。

主窗口入口做"壳"，不写任何业务逻辑。
主窗口用一个注册表统一装载，加新功能只需要在这里加一行，不用改主窗口逻辑。
"""
from __future__ import annotations

import sys
import traceback
from datetime import datetime
from pathlib import Path


def _install_crash_logger() -> None:
    """窗口化程序无控制台，未捕获异常会静默丢失。

    把异常堆栈写到 exe 同目录的 crash.log，便于发给他人使用时排查问题。
    """
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    log_path = base / "crash.log"

    def _hook(exc_type, exc, tb) -> None:
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("\n=== crash %s ===\n" % datetime.now().isoformat())
                f.write("".join(traceback.format_exception(exc_type, exc, tb)))
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = _hook


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
    from features.bank_classify.bank_classify_ui import BankClassifyFeature
    FEATURES.append(BankClassifyFeature())
except Exception as e:  # noqa: BLE001
    print(f"[warn] 跳过 features.bank_classify: {e}", file=sys.stderr)

try:
    from features.js_bank_statement.js_bank_statement_ui import JsBankStmtFeature
    FEATURES.append(JsBankStmtFeature())
except Exception as e:  # noqa: BLE001
    print(f"[warn] 跳过 features.js_bank_statement: {e}", file=sys.stderr)

# try:
#     from features.invoice.invoice_ui import InvoiceFeature
#     FEATURES.append(InvoiceFeature())
# except Exception as e:
#     print(f"[warn] 跳过 features.invoice: {e}", file=sys.stderr)

try:
    from features.notes_receivable_import.notes_receivable_import_ui import NotesReceivableImportFeature
    FEATURES.append(NotesReceivableImportFeature())
except Exception as e:  # noqa: BLE001
    print(f"[warn] 跳过 features.notes_receivable_import: {e}", file=sys.stderr)


def main() -> int:
    _install_crash_logger()
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
