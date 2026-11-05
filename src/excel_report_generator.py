"""Генерация единого Excel-отчёта из нескольких DataFrame.

Модуль принимает DataFrame'ы из разных модулей и формирует единый
Excel-файл, где каждый DataFrame записывается на отдельный лист.
Каждый лист оформляется как Excel-таблица (AutoFilter, заголовки).
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import ifcopenshell
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

# При запуске файла напрямую Python добавляет в sys.path только каталог src.
# Добавляем корень проекта, чтобы были доступны logger.py и пакет src.
if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.utils.logger import Logger as log


def create_excel_report(
    sheets: dict[str, pd.DataFrame],
    output_path: Path,
) -> Path:
    """Формирует единый Excel-отчёт из нескольких DataFrame.

    Каждый DataFrame записывается на отдельный лист с именем листа,
    указанным в ключе словаря.

    :param sheets: Словарь {имя_листа: DataFrame}.
    :param output_path: Путь к выходному XLSX-файлу.
    :returns: Путь к созданному файлу.
    """

    if not sheets:
        log.warning("Нет данных для экспорта в Excel", name="excel_report")
        return output_path

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            # Не пропускаю пустые листы — создаю всегда

            # Обрезаю длинные имена листов (лимит Excel — 31 символ)
            # Заменяю недопустимые символы
            invalid_chars = r'\/:*?[]'
            for ch in invalid_chars:
                sheet_name = sheet_name.replace(ch, "")
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]

            df.to_excel(writer, sheet_name=sheet_name, index=False)

            worksheet = writer.sheets[sheet_name]

            # Заголовки жирным
            for cell in worksheet[1]:
                cell.font = Font(bold=True)

            # Фильтры
            worksheet.auto_filter.ref = worksheet.dimensions

            # Заморозка первой строки
            worksheet.freeze_panes = "A2"

            # Автоширина столбцов
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        cell_length = len(str(cell.value or ""))
                        max_length = max(max_length, cell_length)
                    except Exception:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = max(adjusted_width, 10)

    log.info(f"Excel-отчёт сохранён: {output_path} (листов: {len(sheets)})", name="excel_report")
    return output_path


def generate_full_report(
    ifc_path: Path,
    output_dir: Path,
    spaces_df: pd.DataFrame | None = None,
    tep_df: pd.DataFrame | None = None,
    problems_df: pd.DataFrame | None = None,
) -> Path:
    """Формирует полный отчёт со всеми листами.

    :param ifc_path: Путь к IFC-файлу (для имени файла отчёта).
    :param output_dir: Каталог для результатов.
    :param spaces_df: Общий DataFrame помещений/зон.
    :param tep_df: DataFrame ТЭП.
    :param problems_df: DataFrame проблемных помещений.
    :returns: Путь к созданному файлу.
    """

    output_dir.mkdir(parents=True, exist_ok=True)
    log.configure(str(output_dir / "excel_report.log"))
    log.init("Генерация полного Excel-отчёта")

    sheets: dict[str, pd.DataFrame] = {}

    if spaces_df is not None:
        sheets["Помещения/зоны"] = spaces_df
        log.info(f"Добавлен лист 'Помещения/зоны': {len(spaces_df)} строк", name="excel_report")

    if tep_df is not None:
        sheets["ТЭП"] = tep_df
        log.info(f"Добавлен лист 'ТЭП': {len(tep_df)} строк", name="excel_report")

    if problems_df is not None:
        sheets["Проблемные зоны"] = problems_df
        log.info(f"Добавлен лист 'Проблемные зоны': {len(problems_df)} строк", name="excel_report")

    if not sheets:
        log.warning("Нет данных для формирования отчёта", name="excel_report")
        return output_dir / "report.xlsx"

    stem = ifc_path.stem
    output_path = output_dir / f"{stem}_report.xlsx"
    result_path = create_excel_report(sheets, output_path)

    log.info(
        f"Полный отчёт сформирован: {result_path} ({len(sheets)} листов)",
        name="excel_report",
    )
    return result_path


if __name__ == "__main__":
    # Пример использования
    IFC_PATH = Path(r"samples\НН_К00_С00_АР_АГР.ifc")
    OUTPUT_DIR = Path(r".output")

    # Импортирую для тестирования
    from src.ifc_space_checker import get_spaces_dataframe, get_problem_spaces_dataframe  # noqa: E402
    from src.tep_generator import calculate_tep  # noqa: E402

    model = ifcopenshell.open(str(IFC_PATH))

    spaces_df = get_spaces_dataframe(model, ifc_path=IFC_PATH, tolerance=0.1)
    problems_df = get_problem_spaces_dataframe(spaces_df)
    tep_df = calculate_tep(model, spaces_df)

    generate_full_report(IFC_PATH, OUTPUT_DIR, spaces_df, tep_df, problems_df)
