"""Переиспользуемый виджет выбора файла или каталога."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QLineEdit, QPushButton, QWidget


class FileSelector(QWidget):
    """Строка с подписью, полем пути и кнопкой выбора."""

    def __init__(self, label: str, directory: bool = False, parent=None) -> None:
        super().__init__(parent)
        self._directory = directory
        self._edit = QLineEdit(self)
        self._button = QPushButton("Обзор…", self)
        self._button.clicked.connect(self._choose)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(QLabel(label, self))
        layout.addWidget(self._edit, 1)
        layout.addWidget(self._button)

    def path(self) -> Path:
        """Возвращает путь, указанный пользователем."""

        return Path(self._edit.text().strip())

    def set_path(self, path: Path) -> None:
        """Устанавливает путь в поле выбора."""

        self._edit.setText(str(path))

    def _choose(self) -> None:
        """Открывает диалог выбора файла или каталога."""

        if self._directory:
            selected = QFileDialog.getExistingDirectory(self, "Выберите каталог")
        else:
            selected, _ = QFileDialog.getOpenFileName(self, "Выберите файл")
        if selected:
            self._edit.setText(selected)

