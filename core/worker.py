#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""core/worker.py — 后台任务封装。

按 pyside6-fluent-gui-design skill 的要求：长耗时操作必须离开 UI 线程，
用 QThread + 信号上报进度，避免界面卡死。UI 行为约定：
    - 开始时禁用主操作按钮
    - 进度/日志信号到达时追加消息
    - 结束/失败时恢复按钮，并给出结果摘要（InfoBar / 提示框）

用法：
    self._worker = run_in_thread(self._do_work, on_finished=self._on_done,
                                 on_failed=self._on_fail, on_progress=self._on_log)
"""
from __future__ import annotations

from typing import Callable, Optional

from PySide6.QtCore import QEventLoop, QObject, QThread, QTimer, Signal


class Worker(QObject):
    """在独立线程里执行一个 callable 的工作对象。

    Signals:
        started:  任务开始
        progress: 进度/日志文本（可多次）
        finished: 任务成功，附带 callable 的返回值
        failed:   任务抛异常，附带错误文本
    """

    started = Signal()
    progress = Signal(str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, fn: Callable, *args, **kwargs) -> None:
        super().__init__()
        self._fn = fn
        self._args = args
        self._kwargs = kwargs

    def run(self) -> None:
        # emit started signal and run the callable; add file logging to help debug
        self.started.emit()
        try:
            with open("worker_debug.log", "a", encoding="utf-8") as _f:
                _f.write("Worker.run: before fn call\n")
            result = self._fn(*self._args, **self._kwargs)
            with open("worker_debug.log", "a", encoding="utf-8") as _f:
                _f.write("Worker.run: after fn call\n")
        except Exception as e:  # noqa: BLE001
            with open("worker_debug.log", "a", encoding="utf-8") as _f:
                _f.write(f"Worker.run: exception: {type(e).__name__}: {e}\n")
            self.failed.emit(f"{type(e).__name__}: {e}")
            return
        self.finished.emit(result)


def run_in_thread(
    fn: Callable,
    *args,
    on_progress: Optional[Callable[[str], None]] = None,
    on_finished: Optional[Callable[[object], None]] = None,
    on_failed: Optional[Callable[[str], None]] = None,
    parent: Optional[QObject] = None,
) -> QThread:
    """把 fn 丢到后台线程执行，返回 QThread（调用方持有以便管理生命周期）。

    回调都在主线程执行（信号跨线程自动排队）。
    """
    thread = QThread(parent)
    worker = Worker(fn, *args)
    worker.moveToThread(thread)

    thread.started.connect(worker.run)
    worker.started.connect(lambda: on_progress and on_progress("[开始] 任务启动"))
    if on_progress is not None:
        worker.progress.connect(on_progress)
    if on_finished is not None:
        worker.finished.connect(on_finished)
    if on_failed is not None:
        worker.failed.connect(on_failed)
    worker.finished.connect(thread.quit)
    worker.failed.connect(thread.quit)
    worker.finished.connect(worker.deleteLater)
    worker.failed.connect(worker.deleteLater)
    thread.finished.connect(thread.deleteLater)

    thread.start()
    return thread


def run_blocking(fn: Callable, timeout_ms: int = 30_000) -> object:
    """极少数需要在测试中同步等待后台结果时用的辅助（UI 代码请勿使用）。"""
    result_box: dict[str, object] = {}

    loop = QEventLoop()
    timer = QTimer()
    timer.setSingleShot(True)
    timer.timeout.connect(loop.quit)

    thread = QThread()
    worker = Worker(fn)
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.finished.connect(lambda r: (result_box.__setitem__("r", r), loop.quit()))
    worker.failed.connect(lambda e: (result_box.__setitem__("e", e), loop.quit()))
    worker.finished.connect(thread.quit)
    worker.failed.connect(thread.quit)
    worker.finished.connect(worker.deleteLater)
    worker.failed.connect(worker.deleteLater)
    thread.finished.connect(thread.deleteLater)

    thread.start()
    timer.start(timeout_ms)
    loop.exec()
    if "e" in result_box:
        raise RuntimeError(result_box["e"])
    return result_box.get("r")
