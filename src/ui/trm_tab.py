"""Вкладка совмещения TRM и IFC с просмотром PDF."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QPointF
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPdfWidgets import QPdfView
from PyQt6.QtWidgets import QHBoxLayout, QLabel, QMessageBox, QPushButton, QVBoxLayout, QWidget

from src.trm_overlay import render_trm_album
from src.ui.file_selector import FileSelector
from src.ui.workers import BaseWorker
from src.utils.logger import Logger as log


class TrmTab(QWidget):
    """Форма запуска TRM-совмещения и встроенного просмотра PDF-альбома."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._worker: BaseWorker | None = None
        self._trm_selector = FileSelector("TRM-контейнер:")
        self._ifc_selector = FileSelector("IFC-модель:")
        self._output_selector = FileSelector("Каталог результата:", directory=True)
        self._run_button = QPushButton("Сформировать PDF-альбом")
        self._previous_button = QPushButton("‹")
        self._next_button = QPushButton("›")
        self._page_label = QLabel("Страница 0/0")
        self._status = QLabel("Выберите TRM-контейнер и IFC-модель.")
        self._document = QPdfDocument(self)
        self._viewer = QPdfView(self)
        self._viewer.setDocument(self._document)
        self._run_button.clicked.connect(self._run)
        self._previous_button.clicked.connect(self._previous_page)
        self._next_button.clicked.connect(self._next_page)

        controls = QHBoxLayout()
        controls.addWidget(self._previous_button)
        controls.addWidget(self._next_button)
        controls.addWidget(self._page_label)
        controls.addStretch(1)

        layout = QVBoxLayout(self)
        layout.addWidget(self._trm_selector)
        layout.addWidget(self._ifc_selector)
        layout.addWidget(self._output_selector)
        layout.addWidget(self._run_button)
        layout.addWidget(self._status)
        layout.addLayout(controls)
        layout.addWidget(self._viewer, 1)
        self._update_navigation()

    def _run(self) -> None:
        """Проверяет пути и запускает формирование PDF в рабочем потоке."""

        trm_path = self._trm_selector.path()
        ifc_path = self._ifc_selector.path()
        output_dir = self._output_selector.path()
        if not trm_path.is_file() or trm_path.suffix.lower() != ".zip":
            QMessageBox.warning(self, "TRM", "Укажите существующий TRM ZIP-контейнер.")
            return
        if not ifc_path.is_file():
            QMessageBox.warning(self, "TRM", "Укажите существующий IFC-файл.")
            return
        if not output_dir:
            QMessageBox.warning(self, "TRM", "Укажите каталог результата.")
            return

        output_path = output_dir / f"{trm_path.stem}_overlay.pdf"
        self._set_running(True, "Формирование PDF-альбома выполняется…")
        self._worker = BaseWorker(
            lambda: render_trm_album(trm_path, ifc_path, output_path),
            "формирование TRM PDF-альбома",
            self,
        )
        self._worker.succeeded.connect(self._completed)
        self._worker.failed.connect(self._failed)
        self._worker.finished.connect(lambda: self._set_running(False, self._status.text()))
        self._worker.start()

    def _completed(self, output_path: Path) -> None:
        """Загружает созданный PDF во встроенный просмотрщик."""

        error = self._document.load(str(output_path))
        if error != QPdfDocument.Error.None_:
            self._failed(f"Не удалось открыть PDF: {error}")
            return
        self._viewer.pageNavigator().jump(0, QPointF(), 0)
        self._status.setText(f"Готово: {output_path}")
        self._update_navigation()
        log.info(f"Интерфейс TRM открыл PDF: {output_path}", name="ui.trm")

    def _failed(self, message: str) -> None:
        """Показывает ошибку фоновой операции."""

        self._status.setText(f"Ошибка: {message}")
        QMessageBox.critical(self, "TRM", message)

    def _previous_page(self) -> None:
        """Переходит на предыдущую страницу PDF."""

        navigator = self._viewer.pageNavigator()
        navigator.jump(max(navigator.currentPage() - 1, 0), QPointF(), 0)
        self._update_navigation()

    def _next_page(self) -> None:
        """Переходит на следующую страницу PDF."""

        navigator = self._viewer.pageNavigator()
        navigator.jump(min(navigator.currentPage() + 1, self._document.pageCount() - 1), QPointF(), 0)
        self._update_navigation()

    def _update_navigation(self) -> None:
        """Обновляет состояние кнопок навигации по страницам."""

        page_count = self._document.pageCount()
        current_page = self._viewer.pageNavigator().currentPage() if page_count else 0
        self._page_label.setText(f"Страница {current_page + 1}/{page_count}" if page_count else "Страница 0/0")
        self._previous_button.setEnabled(current_page > 0)
        self._next_button.setEnabled(page_count > 0 and current_page < page_count - 1)

    def _set_running(self, running: bool, message: str) -> None:
        """Обновляет доступность формы во время расчёта."""

        self._run_button.setEnabled(not running)
        self._status.setText(message)

