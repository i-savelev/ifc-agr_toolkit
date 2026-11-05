"""Формирование ТЭП из IFC-модели ЦИМ АГР.

Модуль рассчитывает ТЭП только на основе DataFrame помещений/зон.
Данные извлекаются через публичные функции ifc_space_checker.
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
    ("uniqueObjectId", "Уникальный идентификационный номер (УИН)", "", False, "UniqueObjectId"),
    ("objectName", "Наименование объекта", "", False, "ObjectName"),
    ("registrationNumber", "Регистрационный номер", "", False, "RegistrationNumber"),
    ("statementDate", "Дата", "", False, "Date"),
    ("address", "Адрес", "", False, "Address"),
    ("cadastralNumbers", "Кадастровые номера земельных участков", "", False, "CadastralNumber"),
    ("gpzuNumbers", "Номера ГПЗУ", "", False, None),
    ("pptNumbers", "Номера ППТ", "", False, None),
    ("gzkPermitNumbers", "Номера решений ГЗК", "", False, None),
    ("krtNumbers", "Номера КРТ", "", False, None),
    ("ppmNumbers", "Номера ППМ", "", False, None),
    ("projectOrganization", "Проектная организация", "", False, "ProjectOrganization"),
    ("leadArchitect", "Руководитель авторского коллектива", "", False, "ProjectTeamLead"),
    ("authorTeam", "Авторский коллектив", "", False, "Member"),
    ("workTypes", "Виды работ", "", False, "WorkTypes"),
    ("functionalPurposes", "Функциональные назначения объекта", "", False, "FunctionalPurpose"),
    ("developmentHeight", "Высота застройки", "м", False, "BuildingHeight"),
    ("landArea", "Площадь участка", "га", False, "Area"),
    ("totalOuterArea", "Суммарная поэтажная площадь в габаритах наружных стен (всего)", "м²", False, "TotalFloorArea"),
    ("residentialOuterArea", "Суммарная поэтажная площадь в габаритах наружных стен жилых объектов", "м²", False, "ResidentialPartArea"),
    ("residentialLivingOuterArea", "Суммарная поэтажная площадь жилых объектов (жилая часть)", "м²", False, "ResidentialPartOfResidentialArea"),
    ("residentialNonLivingOuterArea", "Суммарная поэтажная площадь жилых объектов (нежилая часть)", "м²", False, "NonResidentialPartOfResidentialArea"),
    ("nonResidentialOuterArea", "Суммарная поэтажная площадь нежилых объектов", "м²", False, "NonResidentialObjectsArea"),
    ("totalObjectArea", "Общая площадь объекта", "м²", False, "TotalObjectArea"),
    ("aboveGroundArea", "Наземная площадь", "м²", False, "AboveGroundArea"),
    ("undergroundArea", "Подземная площадь", "м²", False, "UndergroundArea"),
    ("nonResidentialAreaInResidential", "Нежилая наземная площадь, в том числе нежилая часть жилых зданий", "м²", False, "NonResidentialAboveGroundAreaNonResidentialPart"),
    ("nonResidentialStandaloneArea", "Нежилая наземная площадь, в том числе нежилые объекты", "м²", False, "NonResidentialAboveGroundAreaNonResidentialObjects"),
    ("apartmentAreaWithoutSummer", "Площадь квартир (без летних помещений)", "м²", False, "ApartmentsAreaExSummerPremises"),
    ("apartmentCount", "Количество квартир", "шт.", False, "ApartmentsCount"),
    ("nonResidentialObject", "Нежилой объект", "", False, "NonResidentialObject"),
    ("landscapingArea", "Площадь озеленения", "м²", False, "GreenArea"),
    ("playgroundArea", "Площадь детских площадок", "м²", False, "ChildrenPlaygroundArea"),
    ("adultRecreationArea", "Площадь площадок для отдыха взрослых", "м²", False, "RecreationsArea"),
    ("objectHeight", "Высота объекта", "м", False, "ObjectHeight"),
    ("floorCount", "Количество этажей", "шт.", False, "FloorsCount"),
    ("apartmentArea", "Площадь квартир", "м²", False, "ApartmentsArea"),
    ("totalApartmentsArea", "Общая площадь квартир", "м²", False, "TotalApartmentsArea"),
    ("signageCount", "Количество информационных конструкций (вывесок)", "шт.", False, "InformationStructuresCount"),
    ("signageDetails", "Параметры и содержание информационных конструкций (вывесок)", "", False, "InformationStructuresParameters"),
)


def _sum_area_column(df: pd.DataFrame, column: str = "RUS_Area (м²)") -> Decimal | None:
    """Суммирует поле площади в наборе строк DataFrame."""

    if df.empty:
        return None
    values = df[column].dropna()
    if values.empty:
        return None
    return sum((value for value in values if value is not None), Decimal("0"))


def _filter_space(spaces_df: pd.DataFrame, name: str) -> pd.DataFrame:
    """Выбирает помещения/зоны по названию."""

    return spaces_df[spaces_df["Name"] == name]


@log_dataframe
def calculate_tep(spaces_df: pd.DataFrame) -> pd.DataFrame:
    """Рассчитывает ТЭП из DataFrame помещений/зон.

    :param spaces_df: DataFrame помещений/зон, возвращённый get_spaces_dataframe.
    :returns: DataFrame со всеми полями формы сайта.
    """

    spp = _filter_space(spaces_df, "СПП в ГНС")
    total = _filter_space(spaces_df, "Общая площадь")
    residential_spp = spp[spp["RUS_ResidentalRoom"] == True]
    non_residential_spp = spp[spp["RUS_ResidentalRoom"] == False]
    above_total = total[total["RUS_BasementFloor"] == False]
    underground_total = total[total["RUS_BasementFloor"] == True]
    non_residential_total_above = total[(total["RUS_BasementFloor"] == False) & (total["RUS_ResidentalRoom"] == False)]
    apartment_spaces = spaces_df[spaces_df["RUS_Group_Type"] == "Квартира"]
    summer_names = {"Балкон", "Лоджия", "Терраса", "Веранда", "Холодная кладовая", "Тамбур"}
    apartment_without_summer = apartment_spaces[~apartment_spaces["Name"].isin(summer_names)]

    def calculated(value: Any, unit: str, source: str, comment: str = "") -> TepValue:
        return TepValue(value, unit, source, "calculated" if value is not None else "missing", comment)

    def unavailable(comment: str) -> TepValue:
        return TepValue(None, "", "", "unavailable", comment)

    apartment_count = len(set(apartment_spaces["RUS_Group_Number"].dropna().astype(str)))

    values: dict[str, TepValue] = {
        "ObjectName": unavailable("Требуется метаданные проекта из ifc_project"),
        "Address": unavailable("Требуется метаданные проекта из ifc_project"),
        "FunctionalPurpose": unavailable("Требуется метаданные проекта из ifc_project"),
        "UniqueObjectId": unavailable("Требуется метаданные проекта из ifc_project"),
        "RegistrationNumber": unavailable("В IFC не найден регистрационный номер"),
        "Date": unavailable("В IFC не найдена дата формирования сведений"),
        "CadastralNumber": unavailable("Требуется метаданные проекта из ifc_project"),
        "ProjectOrganization": unavailable("В IFC нет сведений о проектной организации"),
        "ProjectTeamLead": unavailable("В IFC нет руководителя авторского коллектива"),
        "Member": unavailable("В IFC нет состава авторского коллектива"),
        "WorkTypes": unavailable("В IFC нет вида работ"),
        "BuildingHeight": unavailable("Требуется высота здания из ifc_project"),
        "Area": unavailable("Требуется площадь участка из ifc_project"),
        "TotalFloorArea": calculated(_sum_area_column(spp), "м²", "IfcSpace.Name=СПП в ГНС / RUS_Area (м²)"),
        "ResidentialPartArea": calculated(_sum_area_column(spp), "м²", "IfcSpace.Name=СПП в ГНС / RUS_Area (м²)", "Жилой объект определяется по RUS_FNO"),
        "ResidentialPartOfResidentialArea": calculated(_sum_area_column(residential_spp), "м²", "IfcSpace.RUS_ResidentalRoom=true / RUS_Area (м²)"),
        "NonResidentialPartOfResidentialArea": calculated(_sum_area_column(non_residential_spp), "м²", "IfcSpace.RUS_ResidentalRoom=false / RUS_Area (м²)"),
        "NonResidentialObjectsArea": unavailable("Для отдельного нежилого объекта нужна соответствующая IFC-модель"),
        "TotalObjectArea": calculated(_sum_area_column(total), "м²", "IfcSpace.Name=Общая площадь / RUS_Area (м²)"),
        "AboveGroundArea": calculated(_sum_area_column(above_total), "м²", "IfcSpace.Name=Общая площадь + RUS_BasementFloor=false / RUS_Area (м²)"),
        "UndergroundArea": calculated(_sum_area_column(underground_total), "м²", "IfcSpace.Name=Общая площадь + RUS_BasementFloor=true / RUS_Area (м²)"),
        "NonResidentialAboveGroundAreaNonResidentialPart": calculated(_sum_area_column(non_residential_total_above), "м²", "IfcSpace.Name=Общая площадь + RUS_ResidentalRoom=false / RUS_Area (м²)"),
        "NonResidentialAboveGroundAreaNonResidentialObjects": unavailable("Для отдельного нежилого объекта нужна соответствующая IFC-модель"),
        "ApartmentsAreaExSummerPremises": calculated(_sum_area_column(apartment_without_summer), "м²", "IfcSpace.RUS_Group_Number + исключение летних помещений"),
        "ApartmentsCount": calculated(apartment_count or None, "шт.", "уникальные IfcSpace.RUS_Group_Number"),
        "GreenArea": unavailable("Площадь озеленения находится в ЦИМ БиО"),
        "ChildrenPlaygroundArea": unavailable("Детские площадки находятся в ЦИМ БиО"),
        "RecreationsArea": unavailable("Площадки отдыха находятся в ЦИМ БиО"),
        "ObjectHeight": unavailable("Требуется высота здания из ifc_project"),
        "FloorsCount": calculated(apartment_count or None, "шт.", "IfcBuildingStorey.Elevation >= 0"),
        "ApartmentsArea": calculated(_sum_area_column(apartment_spaces), "м²", "IfcSpace.RUS_Group_Number / RUS_Area (м²)"),
        "TotalApartmentsArea": calculated(_sum_area_column(apartment_spaces), "м²", "IfcSpace.RUS_Group_Number / RUS_Area (м²)"),
        "InformationStructuresCount": unavailable("В IFC не найдены информационные конструкции"),
        "InformationStructuresParameters": unavailable("В IFC нет параметров информационных конструкций"),
        "NonResidentialObject": unavailable("Требуется функциональное назначение из ifc_project"),
    }
    rows = []
    for form_field, title, unit, optional, source_field in FORM_FIELDS:
        item = values.get(source_field, unavailable("Поле не извлекается из текущей IFC-модели")) if source_field else unavailable("В IFC не найдено соответствующее поле")
        rows.append({"field": form_field, "indicator": title, "value": item.value, "unit": item.unit or unit, "source": item.source, "status": item.status, "comment": item.comment, "optional": optional})
    return pd.DataFrame(rows)


def export_tep_excel(tep: pd.DataFrame, output_path: Path) -> Path:
    """Экспортирует ТЭП и источники значений в Excel.

    :param tep: DataFrame, возвращённый :func:`calculate_tep`.
    :param output_path: Путь к XLSX-файлу.
    :returns: Путь к созданному XLSX-файлу.
    """

    output_path.parent.mkdir(parents=True, exist_ok=True)
    export_frame = tep[["field", "indicator", "value", "unit", "source", "status", "comment"]].copy()
    export_frame.columns = ["Поле формы", "Показатель", "Значение", "Ед. изм.", "Источник IFC", "Статус", "Комментарий"]
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
