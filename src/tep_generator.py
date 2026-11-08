"""Формирование ТЭП из IFC-модели ЦИМ АГР.

Модуль рассчитывает ТЭП только на основе DataFrame помещений/зон.
Каждое поле ТЭП — отдельная функция, которая принимает DataFrame,
фильтрует по критериям и суммирует RUS_Area (м²).
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

# При запуске файла напрямую Python добавляет в sys.path только каталог src.
# Добавляем корень проекта, чтобы были доступны logger.py и пакет src.
if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.utils.logger import Logger as log
from src.utils.logger_decorator import log_dataframe


@dataclass(frozen=True)
class TepValue:
    """Значение ТЭП и сведения о его происхождении.

    :param value: Значение показателя или None, если оно не найдено.
    :param unit: Единица измерения.
    :param source: Источник значения в IFC.
    :param status: Статус извлечения значения.
    :param comment: Дополнительный комментарий.
    """

    value: Any
    unit: str
    source: str
    status: str = "calculated"
    comment: str = ""


FORM_FIELDS: tuple[tuple[str, str, str, bool, str | None], ...] = (
    ("totalOuterArea", "Суммарная поэтажная площадь в габаритах наружных стен (всего)", "м²", False, "totalOuterArea"),
    ("residentialOuterArea", "Суммарная поэтажная площадь в габаритах наружных стен жилых объектов", "м²", False, "residentialOuterArea"),
    ("residentialLivingOuterArea", "Суммарная поэтажная площадь жилых объектов (жилая часть)", "м²", False, "residentialLivingOuterArea"),
    ("residentialNonLivingOuterArea", "Суммарная поэтажная площадь жилых объектов (нежилая часть)", "м²", False, "residentialNonLivingOuterArea"),
    ("totalObjectArea", "Общая площадь объекта", "м²", False, "totalObjectArea"),
    ("aboveGroundArea", "Наземная площадь", "м²", False, "aboveGroundArea"),
    ("undergroundArea", "Подземная площадь", "м²", False, "undergroundArea"),
    ("nonResidentialAreaInResidential", "Нежилая наземная площадь, в том числе нежилая часть жилых зданий", "м²", False, "nonResidentialAreaInResidential"),
    ("apartmentArea", "Площадь квартир", "м²", False, "apartmentArea"),
    ("apartmentAreaWithoutSummer", "Площадь квартир (без летних помещений)", "м²", False, "apartmentAreaWithoutSummer"),
    ("apartmentCount", "Количество квартир", "шт.", False, "apartmentCount"),
    ("floorCount", "Количество этажей", "шт.", False, "floorCount"),
    ("totalApartmentsArea", "Общая площадь квартир", "м²", False, "totalApartmentsArea"),
)


def _sum_by_mask(spaces_df: pd.DataFrame, mask: pd.Series) -> Decimal | None:
    """Суммирует RUS_Area (м²) по условию (маске).

    :param spaces_df: Общий DataFrame помещений/зон.
    :param mask: Булева серия, определяющая отобранные строки.
    :returns: Сумма или None.
    """

    filtered = spaces_df[mask]
    if filtered.empty:
        return None
    area_values = filtered["RUS_Area (м²)"].dropna()
    if area_values.empty:
        return None
    return sum((v for v in area_values if v is not None), Decimal("0"))


# ============================================================
# Поля ТЭП — каждая функция принимает spaces_df и считает сумму
# ============================================================


def calc_total_outer_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Суммарная поэтажная площадь в габаритах наружных стен (всего)."""

    mask = spaces_df["RUS_Name"] == "СПП в ГНС"
    return _sum_by_mask(spaces_df, mask), "RUS_Name='СПП в ГНС'"


def calc_residential_outer_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Суммарная поэтажная площадь в габаритах наружных стен жилых объектов."""

    mask = spaces_df["RUS_Name"] == "СПП в ГНС"
    return _sum_by_mask(spaces_df, mask), "RUS_Name='СПП в ГНС'"


def calc_residential_living_outer_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Суммарная поэтажная площадь жилых объектов (жилая часть)."""

    mask = (spaces_df["RUS_Name"] == "СПП в ГНС") & (spaces_df["RUS_ResidentalRoom"] == True)
    return _sum_by_mask(spaces_df, mask), "RUS_Name='СПП в ГНС' + RUS_ResidentalRoom=True"


def calc_residential_non_living_outer_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Суммарная поэтажная площадь жилых объектов (нежилая часть)."""

    mask = (spaces_df["RUS_Name"] == "СПП в ГНС") & (spaces_df["RUS_ResidentalRoom"] == False)
    return _sum_by_mask(spaces_df, mask), "RUS_Name='СПП в ГНС' + RUS_ResidentalRoom=False"


def calc_total_object_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Общая площадь объекта."""

    mask = spaces_df["RUS_Name"] == "Общая площадь"
    return _sum_by_mask(spaces_df, mask), "RUS_Name='Общая площадь'"


def calc_above_ground_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Наземная площадь."""

    mask = (spaces_df["RUS_Name"] == "Общая площадь") & (spaces_df["RUS_BasementFloor"] == False)
    return _sum_by_mask(spaces_df, mask), "RUS_Name='Общая площадь' + RUS_BasementFloor=False"


def calc_underground_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Подземная площадь."""

    mask = (spaces_df["RUS_Name"] == "Общая площадь") & (spaces_df["RUS_BasementFloor"] == True)
    return _sum_by_mask(spaces_df, mask), "RUS_Name='Общая площадь' + RUS_BasementFloor=True"


def calc_non_residential_area_in_residential(spaces_df: pd.DataFrame) -> Decimal | None:
    """Нежилая наземная площадь, в том числе нежилая часть жилых зданий."""

    mask = (
        (spaces_df["RUS_Name"] == "Общая площадь")
        & (spaces_df["RUS_BasementFloor"] == False)
        & (spaces_df["RUS_ResidentalRoom"] == False)
    )
    return _sum_by_mask(spaces_df, mask), "RUS_Name='Общая площадь' + RUS_BasementFloor=False + RUS_ResidentalRoom=False"


def calc_apartment_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Площадь квартир."""

    mask = spaces_df["RUS_Group_Type"] == "Квартира"
    return _sum_by_mask(spaces_df, mask), "RUS_Group_Type='Квартира'"


def calc_apartment_area_without_summer(spaces_df: pd.DataFrame) -> Decimal | None:
    """Площадь квартир (без летних помещений)."""

    summer_zone_names = {"Балкон", "Веранда"}
    mask = (spaces_df["RUS_Group_Type"] == "Квартира") & (
        ~spaces_df["RUS_MSSK_SpaceAndZone_Name"].isin(summer_zone_names)
    )
    return _sum_by_mask(spaces_df, mask), "RUS_Group_Type='Квартира' + RUS_MSSK_SpaceAndZone_Name not in ('Балкон', 'Веранда')"


def calc_apartment_count(spaces_df: pd.DataFrame) -> int | None:
    """Количество квартир (уникальные RUS_Group_Number, содержащие цифры)."""

    apartment_spaces = spaces_df[spaces_df["RUS_Group_Type"] == "Квартира"]
    if apartment_spaces.empty:
        return None, "RUS_Group_Type='Квартира'"
    # Фильтрую только те Group_Number, которые содержат цифры
    group_nums = apartment_spaces["RUS_Group_Number"].dropna().astype(str)
    group_nums_with_digits = group_nums[group_nums.str.contains(r'\d', na=False)]
    count = group_nums_with_digits.nunique()
    return int(count) if count > 0 else None, "RUS_Group_Type='Квартира' + RUS_Group_Number содержит цифры"


def calc_floor_count(spaces_df: pd.DataFrame) -> int | None:
    """Количество этажей (уникальные Level, содержащие "основной")."""

    levels = spaces_df["Level"].dropna().astype(str)
    main_levels = levels[levels.str.contains("основной", na=False)]
    count = main_levels.nunique()
    return int(count) if count > 0 else None, "Level содержит 'основной'"


def calc_total_apartments_area(spaces_df: pd.DataFrame) -> Decimal | None:
    """Общая площадь квартир."""

    mask = spaces_df["RUS_Group_Type"] == "Квартира"
    return _sum_by_mask(spaces_df, mask), "RUS_Group_Type='Квартира'"


# ============================================================
# Таблицы расчётов
# ============================================================


CALCULATION_FUNCTIONS: tuple[tuple[str, str, str, str | None], ...] = (
    ("totalOuterArea", "Суммарная поэтажная площадь в габаритах наружных стен (всего)", "м²", "totalOuterArea"),
    ("residentialOuterArea", "Суммарная поэтажная площадь в габаритах наружных стен жилых объектов", "м²", "residentialOuterArea"),
    ("residentialLivingOuterArea", "Суммарная поэтажная площадь жилых объектов (жилая часть)", "м²", "residentialLivingOuterArea"),
    ("residentialNonLivingOuterArea", "Суммарная поэтажная площадь жилых объектов (нежилая часть)", "м²", "residentialNonLivingOuterArea"),
    ("totalObjectArea", "Общая площадь объекта", "м²", "totalObjectArea"),
    ("aboveGroundArea", "Наземная площадь", "м²", "aboveGroundArea"),
    ("undergroundArea", "Подземная площадь", "м²", "undergroundArea"),
    ("nonResidentialAreaInResidential", "Нежилая наземная площадь, в том числе нежилая часть жилых зданий", "м²", "nonResidentialAreaInResidential"),
    ("apartmentArea", "Площадь квартир", "м²", "apartmentArea"),
    ("apartmentAreaWithoutSummer", "Площадь квартир (без летних помещений)", "м²", "apartmentAreaWithoutSummer"),
    ("apartmentCount", "Количество квартир", "шт.", "apartmentCount"),
    ("floorCount", "Количество этажей", "шт.", "floorCount"),
    ("totalApartmentsArea", "Общая площадь квартир", "м²", "totalApartmentsArea"),
)


# ============================================================
# Публичные API
# ============================================================


@log_dataframe
def calculate_tep(spaces_df: pd.DataFrame) -> pd.DataFrame:
    """Рассчитывает ТЭП из DataFrame помещений/зон.

    :param spaces_df: DataFrame помещений/зон, возвращённый get_spaces_dataframe.
    :returns: DataFrame со всеми полями формы сайта.
    """

    func_map = {
        "totalOuterArea": calc_total_outer_area,
        "residentialOuterArea": calc_residential_outer_area,
        "residentialLivingOuterArea": calc_residential_living_outer_area,
        "residentialNonLivingOuterArea": calc_residential_non_living_outer_area,
        "totalObjectArea": calc_total_object_area,
        "aboveGroundArea": calc_above_ground_area,
        "undergroundArea": calc_underground_area,
        "nonResidentialAreaInResidential": calc_non_residential_area_in_residential,
        "apartmentArea": calc_apartment_area,
        "apartmentAreaWithoutSummer": calc_apartment_area_without_summer,
        "apartmentCount": calc_apartment_count,
        "floorCount": calc_floor_count,
        "totalApartmentsArea": calc_total_apartments_area,
    }

    values: dict[str, TepValue] = {}
    for key, title, unit, _ in CALCULATION_FUNCTIONS:
        func = func_map[key]
        result = func(spaces_df)
        # Функции теперь возвращают (значение, комментарий)
        value = result[0] if isinstance(result, tuple) else result
        comment = result[1] if isinstance(result, tuple) else f"Рассчитано через {func.__name__}"
        unit_str = "" if isinstance(value, int) else unit
        values[key] = TepValue(
            value=value,
            unit=unit_str,
            source="",
            status="",
            comment=comment,
        )

    rows = []
    for form_field, title, unit, optional, source_field in FORM_FIELDS:
        item = values.get(source_field, TepValue(None, "", "", "", "Поле не рассчитывается"))
        rows.append({
            "field": form_field,
            "indicator": title,
            "value": item.value,
            "unit": item.unit or unit,
            "comment": item.comment,
            "optional": optional,
        })
    return pd.DataFrame(rows)


def export_tep_excel(tep: pd.DataFrame, output_path: Path) -> Path:
    """Экспортирует ТЭП и источники значений в Excel.

    :param tep: DataFrame, возвращённый :func:`calculate_tep`.
    :param output_path: Путь к XLSX-файлу.
    :returns: Путь к созданному XLSX-файлу.
    """

    output_path.parent.mkdir(parents=True, exist_ok=True)
    export_frame = tep[["field", "indicator", "value", "unit", "comment"]].copy()
    export_frame.columns = ["Поле формы", "Показатель", "Значение", "Ед. изм.", "Способ расчёта"]
    export_frame.to_excel(output_path, index=False, sheet_name="ТЭП")
    workbook = load_workbook(output_path)
    worksheet = workbook["ТЭП"]
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column in worksheet.columns:
        worksheet.column_dimensions[column[0].column_letter].width = min(max(len(str(cell.value or "")) for cell in column) + 2, 60)
    workbook.save(output_path)
    log.info(f"Excel ТЭП сохранён: {output_path}", name="tep")
    return output_path


def generate_tep(ifc_path: Path, spaces_df: pd.DataFrame, output_dir: Path) -> dict[str, Path]:
    """Формирует Excel с ТЭП для IFC-файла.

    :param ifc_path: Путь к IFC-модели (для имени файла).
    :param spaces_df: DataFrame помещений/зон, возвращённый get_spaces_dataframe.
    :param output_dir: Каталог для результатов.
    :returns: Словарь с путём к XLSX-файлу.
    """

    output_dir.mkdir(parents=True, exist_ok=True)
    log.configure(str(output_dir / "tep_generation.log"))
    log.init("Генерация ТЭП из IFC")
    log.info(f"Открытие IFC: {ifc_path}", name="tep")
    tep = calculate_tep(spaces_df)
    stem = ifc_path.stem
    excel_path = export_tep_excel(tep, output_dir / f"{stem}_tep.xlsx")
    log.info(f"Генерация ТЭП завершена: {len(tep)} полей", name="tep")
    return {"excel": excel_path}


if __name__ == "__main__":
    import ifcopenshell

    IFC_PATH = Path(r"samples\НН_К00_С00_АР_АГР.ifc")
    OUTPUT_DIR = Path(r".output")

    # Импортирую для тестирования
    from src.ifc_space_checker import get_spaces_dataframe  # noqa: E402

    model = ifcopenshell.open(str(IFC_PATH))
    spaces_df = get_spaces_dataframe(model, ifc_path=IFC_PATH, tolerance=0.1)
    generate_tep(IFC_PATH, spaces_df, OUTPUT_DIR)
