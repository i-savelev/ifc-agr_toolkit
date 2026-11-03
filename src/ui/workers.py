"""Рабочие потоки для длительных операций графического интерфейса."""

from __future__ import annotations

import traceback
from typing import Any, Callable

from PyQt6.QtCore import QThread, pyqtSignal

from src.utils.logger import Logger as log


class BaseWorker(QThread):
    """Выполняет переданную функцию в отдельном потоке Qt.

    :param operation: Функция без аргументов, которую нужно выполнить.
    :param operation_name: Название операции для журнала.
    """

    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, operation: Callable[[], Any], operation_name: str, parent=None) -> None:
        super().__init__(parent)
        self._operation = operation
        self._operation_name = operation_name

    def run(self) -> None:
        """Запускает операцию и передаёт результат в основной поток."""

        log.info(f"Запуск фоновой операции: {self._operation_name}", name="ui.worker")
        try:
            result = self._operation()
        except Exception as error:  # noqa: BLE001 - ошибка передаётся в интерфейс
            details = traceback.format_exc()
            log.error(f"Ошибка операции {self._operation_name}: {error}\n{details}", name="ui.worker")
            self.failed.emit(str(error))
            return
        log.info(f"Фоновая операция завершена: {self._operation_name}", name="ui.worker")
        self.succeeded.emit(result)

