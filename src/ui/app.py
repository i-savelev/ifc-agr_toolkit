"""Точка запуска графического интерфейса IFC AGR Toolkit."""

from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication

from src.ui.main_window import create_main_window


def run_app() -> int:
    """Запускает приложение и возвращает код завершения Qt."""

    application = QApplication(sys.argv)
    window = create_main_window()
    window.show()
    return application.exec()


if __name__ == "__main__":
    run_app()

