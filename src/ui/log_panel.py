"""Панель отображения сообщений проектного логгера."""

from __future__ import annotations

import queue

from PyQt6.QtCore import QTimer
from PyQt6.QtGui import QFontDatabase
from PyQt6.QtWidgets import QPlainTextEdit, QVBoxLayout, QWidget

from src.utils.logger import Logger as log


class LogPanel(QWidget):
    """Показывает сообщения `Logger` в последней вкладке приложения."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._queue: queue.Queue[str] = queue.Queue()
        self._text = QPlainTextEdit(self)
        self._text.setReadOnly(True)
        self._text.setFont(QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont))
        self._text.setStyleSheet("QPlainTextEdit { font-size: 9pt; }")

        layout = QVBoxLayout(self)
        layout.addWidget(self._text)

        log.attach_gui_queue(self._queue)
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._read_queue)
        self._timer.start(150)

    def _read_queue(self) -> None:
        """Переносит накопленные сообщения из очереди в текстовое поле."""

        while True:
            try:
                message = self._queue.get_nowait()
            except queue.Empty:
                break
            self._text.appendPlainText(message)

