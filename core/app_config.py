#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""core/app_config.py — 应用级配置与路径常量。

所有跨模块共用的路径、应用元信息集中在这里，避免散落硬编码。
按图片目录结构：data/ 在项目根，resources/ 在项目根（待建）。
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path


def _is_frozen() -> bool:
    """是否由 PyInstaller 打包运行（sys.frozen）。"""
    return bool(getattr(sys, "frozen", False))


FROZEN: bool = _is_frozen()


def _resolve_base_dir() -> Path:
    """确定应用根目录（data/ 等所在位置）。

    - 源码运行：项目根 = core/app_config.py 的上上级。
    - PyInstaller 打包后（sys.frozen）：exe 所在目录，使 data/ 与 exe 平级、
      可读写、用户可见（便于替换模板、备份 output 与映射配置）。
    """
    if FROZEN:
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


# 项目根目录（源码运行）或 exe 所在目录（打包后）
BASE_DIR: Path = _resolve_base_dir()

# 打包时 PyInstaller 会把 datas 解压到 sys._MEIPASS；onedir 下与 exe 同目录。
# 内置只读资源（模板 / banks.txt）优先从此处取，缺失时再拷到可写的 DATA_DIR。
RESOURCE_DIR: Path = Path(getattr(sys, "_MEIPASS", str(BASE_DIR)))

# 本地数据 / 缓存目录（与 features/ core/ 平级）
DATA_DIR: Path = BASE_DIR / "data"

# 资源目录：图标、QSS 样式表等（待建，先预留）
RESOURCES_DIR: Path = BASE_DIR / "resources"
ICONS_DIR: Path = RESOURCES_DIR / "icons"
STYLES_FILE: Path = RESOURCES_DIR / "styles.qss"

# bank_classify 示例数据（已迁入 data/）
BANKS_SAMPLE_FILE: Path = DATA_DIR / "banks.txt"

# 应用元信息
APP_NAME: str = "我的财务助手"
APP_VERSION: str = "0.1.0"
ORG_NAME: str = "MyFinAssistant"

# 默认输出目录：生成的 Excel 等放在 data/output/ 下
DEFAULT_OUTPUT_DIR: Path = DATA_DIR / "output"


def ensure_dirs() -> None:
    """启动时调用，保证关键目录存在。

    打包后（FROZEN）额外把内置只读资源（模板 / banks.txt）从解压目录
    (_MEIPASS) 复制到可写的 DATA_DIR，保证首次运行即可用、且用户可替换。
    """
    for d in (DATA_DIR, DEFAULT_OUTPUT_DIR, RESOURCES_DIR, ICONS_DIR):
        d.mkdir(parents=True, exist_ok=True)
    if FROZEN:
        _copy_bundled_assets()


# 需要随包分发、并在运行时落到 DATA_DIR 的内置文件
_BUNDLED_ASSETS: tuple[str, ...] = (
    "应收票据导入模版.xlsx",
    "banks.txt",
)


def _copy_bundled_assets() -> None:
    """把打包资源中的只读文件复制到可写 data/（仅在目标缺失时复制）。"""
    src_data = RESOURCE_DIR / "data"
    if not src_data.is_dir():
        return
    for name in _BUNDLED_ASSETS:
        s = src_data / name
        d = DATA_DIR / name
        if s.is_file() and not d.exists():
            try:
                shutil.copy2(s, d)
            except OSError:
                # 复制失败不阻断启动（仅功能缺失，后续会报错提示）
                pass


# ====== 选文件夹操作的『上次使用目录』持久化 ======
# 规则：初次使用默认用户「下载」文件夹；之后默认上次使用的文件夹。
# 按 key 分别记忆（如 word 输入/输出各自独立），存于 data/last_dirs.json。

_LAST_DIRS_FILE: Path = DATA_DIR / "last_dirs.json"


def downloads_dir() -> Path:
    """用户『下载』文件夹（Windows 标准位置：~\\Downloads）。

    注：绝大多数 Windows 用户下载目录即此路径；若被重定向到其他盘，
    可在对应功能首次 선택后由 set_last_dir 记忆，不影响使用。
    """
    return Path.home() / "Downloads"


def _load_last_dirs() -> dict[str, str]:
    try:
        if _LAST_DIRS_FILE.is_file():
            return json.loads(_LAST_DIRS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        pass
    return {}


def get_last_dir(key: str) -> Path:
    """返回某选文件夹/选文件操作的『上次使用目录』。

    - 有记录且其本身是目录 → 返回它；
    - 有记录但它是文件（选文件/导出时存的是完整路径）→ 返回其所在目录；
    - 否则 → 返回用户下载文件夹（首次使用默认）。
    """
    d = _load_last_dirs().get(key)
    if d:
        p = Path(d)
        if p.is_dir():
            return p
        if p.parent.is_dir():
            return p.parent
    return downloads_dir()


def set_last_dir(key: str, path) -> None:
    """记录某选文件夹操作的『上次使用文件夹』，持久化到 data/last_dirs.json。"""
    data = _load_last_dirs()
    data[key] = str(Path(path))
    try:
        _LAST_DIRS_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        # 持久化失败不阻断主流程（仅下次启动不记忆）
        pass
