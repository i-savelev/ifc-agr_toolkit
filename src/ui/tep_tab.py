"""Вкладка формирования ТЭП из IFC."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtWidgets import QLabel, QMessageBox, QPushButton, QVBoxLayout, QWidget

from src.tep_generator import generate_tep
from src.ui.file_selector import FileSelector
from src.ui.workers import BaseWorker
from src.utils.logger import Logger as log


class TepTab(QWidget):
    """Форма запуска расчёта ТЭП и экспорта результата в Excel."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._worker: BaseWorker | None = None
        self._ifc_selector = FileSelector("IFC-модель:")
        self._output_selector = FileSelector("Каталог результата:", directory=True)
        self._run_button = QPushButton("Сформировать Excel ТЭП")
        self._status = QLabel("Выберите IFC-модель и каталог результата.")
        self._run_button.clicked.connect(self._run)

        layout = QVBoxLayout(self)
        layout.addWidget(self._ifc_selector)
        layout.addWidget(self._output_selector)
        layout.addWidget(self._run_button)
        layout.addWidget(self._status)
        layout.addStretch(1)

    def _run(self) -> None:
        """Проверяет пути и запускает расчёт ТЭП в рабочем потоке."""

        ifc_path = self._ifc_selector.path()
        output_dir = self._output_selector.path()
        if not ifc_path.is_file():
            QMessageBox.warning(self, "ТЭП", "Укажите существующий IFC-файл.")
            return
        if not output_dir:
            QMessageBox.warning(self, "ТЭП", "Укажите каталог результата.")
            return

        self._set_running(True, "Формирование ТЭП выполняется…")
        self._worker = BaseWorker(
            lambda: generate_tep(ifc_path, output_dir),
            "формирование ТЭП",
            self,
        )
        self._worker.succeeded.connect(self._completed)
        self._worker.failed.connect(self._failed)
        self._worker.finished.connect(lambda: self._set_running(False, self._status.text()))
        self._worker.start()

    def _completed(self, result: dict[str, Path]) -> None:
        """Показывает путь к сформированному Excel-файлу."""

        excel_path = result["excel"]
        self._status.setText(f"Готово: {excel_path}")
        log.info(f"Интерфейс ТЭП получил результат: {excel_path}", name="ui.tep")

    def _failed(self, message: str) -> None:
        """Показывает ошибку фоновой операции."""

        self._status.setText(f"Ошибка: {message}")
        QMessageBox.critical(self, "ТЭП", message)

    def _set_running(self, running: bool, message: str) -> None:
        """Обновляет доступность формы во время расчёта."""

        self._run_button.setEnabled(not running)
        self._status.setText(message)

