# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— 我的财务助手（onefile 单文件模式）。

打包命令：
    .venv/Scripts/python.exe -m PyInstaller --noconfirm --clean build.spec

产物：dist/我的财务助手.exe  （单个可执行文件，发给别人双击即用）

说明：
- onefile 模式：所有依赖与数据打包进单个 exe，运行时自解压到临时目录。
- 内置数据（模板 / banks.txt）通过 datas 打进包；app_config 在首次运行时
  把它们从解压目录复制到 exe 同级的 data/（可读写、可替换）。
- 已排除 PyQt5/PyQt6/PySide2，仅保留 PySide6，避免多 Qt 绑定冲突。
"""
from PyInstaller.utils.hooks import collect_all

APP_NAME = "我的财务助手"

datas = []
binaries = []
hiddenimports = []

# Word COM 自动化依赖（word_text_replacer）：win32com / pythoncom 为
# 惰性导入（写在函数体内），PyInstaller 静态扫描会漏掉，必须显式声明。
hiddenimports += [
    "win32com",
    "win32com.client",
    "win32com.shell",
    "pythoncom",
    "pywintypes",
]

# 收集第三方包的数据文件与隐藏依赖（qtawesome 的字体、openpyxl 等）
for _pkg in ("qtawesome", "openpyxl"):
    _d, _b, _h = collect_all(_pkg)
    datas += _d
    binaries += _b
    hiddenimports += _h

# 内置只读资源：模板 + 银行白名单示例，打进包（运行时复制到 data/）
datas += [
    ("data/应收票据导入模版.xlsx", "data"),
    ("data/banks.txt", "data"),
]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 精简体积、排除明确用不到的大模块
        "tkinter",
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebEngineWidgets",
        "PySide6.Qt3DCore",
        "PySide6.Qt3DRender",
        "PySide6.QtMultimedia",
        "PySide6.QtQuick",
        "PySide6.QtQml",
        "PySide6.QtCharts",
        "PySide6.QtDataVisualization",
        # 多 Qt 绑定冲突：仅保留 PySide6
        "PyQt5",
        "PyQt6",
        "PySide2",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
