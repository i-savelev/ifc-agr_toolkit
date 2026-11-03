"""Главное окно IFC AGR Toolkit."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtWidgets import QApplication, QMainWindow, QTabWidget
from qt_material import apply_stylesheet

from src.ui.log_panel import LogPanel
from src.ui.tep_tab import TepTab
from src.ui.trm_tab import TrmTab
from src.utils.logger import Logger as log


class MainWindow(QMainWindow):
    """Контейнер модульных вкладок ТЭП, TRM и журнала."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("IFC AGR Toolkit")
        self.resize(1400, 900)
        self._tabs = QTabWidget(self)
        self.setCentralWidget(self._tabs)
        self._tabs.addTab(TepTab(self), "ТЭП")
        self._tabs.addTab(TrmTab(self), "TRM")
        self._tabs.addTab(LogPanel(self), "Логи")
        self._tabs.setCurrentIndex(0)


def create_main_window() -> MainWindow:
    """Создаёт настроенное главное окно приложения."""

    log.configure(str(Path(".output") / "application.log"))
    log.init("IFC AGR Toolkit")
    window = MainWindow()
    application = QApplication.instance()
    if application is not None:
        apply_stylesheet(application, theme="dark_blue.xml")
    log.info("Главное окно приложения создано", name="ui.main")
    return window

