"""Формирование ТЭП из IFC-модели ЦИМ АГР.

Модуль использует атрибутивные данные IFC, предусмотренные локальными
требованиями к ЦИМ АГР. Геометрия применяется только как резервный источник
для высоты объекта, когда соответствующее атрибутивное значение отсутствует.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import ifcopenshell
import ifcopenshell.geom
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


SOURCE_FIELDS: tuple[tuple[str, str, str, bool], ...] = (
    ("ObjectName", "Наименование объекта", "", False),
    ("RegistrationNumber", "Регистрационный номер", "", True),
    ("Date", "Дата", "", True),
    ("Address", "Адрес", "", False),
    ("CadastralNumber", "Кадастровый номер", "", True),
    ("ProjectOrganization", "Проектная организация", "", False),
    ("ProjectTeamLead", "Руководитель авторского коллектива", "", False),
    ("Member", "Участник авторского коллектива", "", False),
    ("WorkTypes", "Виды работ", "", False),
    ("FunctionalPurpose", "Функциональное назначение", "", False),
    ("BuildingHeight", "Высота застройки", "м", False),
    ("Area", "Площадь участка", "га", True),
    ("TotalFloorArea", "Суммарная поэтажная площадь", "м²", False),
    ("ResidentialPartArea", "Площадь жилых объектов", "м²", True),
    ("ResidentialPartOfResidentialArea", "Жилая часть жилых объектов", "м²", True),
    ("NonResidentialPartOfResidentialArea", "Нежилая часть жилых объектов", "м²", True),
    ("NonResidentialObjectsArea", "Площадь нежилых объектов", "м²", True),
    ("TotalObjectArea", "Общая площадь объекта", "м²", True),
    ("AboveGroundArea", "Наземная площадь", "м²", True),
    ("UndergroundArea", "Подземная площадь", "м²", True),
    ("NonResidentialAboveGroundAreaNonResidentialPart", "Нежилая наземная площадь жилой части", "м²", True),
    ("NonResidentialAboveGroundAreaNonResidentialObjects", "Нежилая наземная площадь нежилых объектов", "м²", True),
    ("ApartmentsAreaExSummerPremises", "Площадь квартир без летних помещений", "м²", True),
    ("ApartmentsCount", "Количество квартир", "шт.", False),
    ("GreenArea", "Площадь озеленения", "м²", True),
    ("ChildrenPlaygroundArea", "Площадь детских площадок", "м²", True),
    ("RecreationsArea", "Площадь площадок отдыха взрослых", "м²", True),
    ("ObjectHeight", "Высота объекта", "м", False),
    ("FloorsCount", "Количество этажей", "шт.", False),
    ("ApartmentsArea", "Площадь квартир", "м²", True),
    ("TotalApartmentsArea", "Общая площадь квартир", "м²", True),
    ("InformationStructuresCount", "Количество информационных конструкций", "шт.", True),
    ("InformationStructuresParameters", "Параметры информационных конструкций", "", True),
)


# Поля формы сайта. Последний элемент строки связывает поле формы с уже
# рассчитанным значением из IFC; None означает, что поле пока не извлекается.
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


def _unwrap_value(value: Any) -> Any:
    """Извлекает Python-значение из IFC measure-типа."""

    return getattr(value, "wrappedValue", value)


def _entity_properties(entity: Any) -> dict[str, Any]:
    """Возвращает свойства экземпляра и назначенного ему типа IFC."""

    values: dict[str, Any] = {}
    for relation in getattr(entity, "IsDefinedBy", ()) or ():
        if not relation.is_a("IfcRelDefinesByProperties"):
            continue
        definition = relation.RelatingPropertyDefinition
        if not definition.is_a("IfcPropertySet"):
            continue
        for prop in definition.HasProperties:
            values[prop.Name] = _unwrap_value(getattr(prop, "NominalValue", None))

    for relation in getattr(entity, "IsTypedBy", ()) or ():
        type_entity = relation.RelatingType
        for definition in getattr(type_entity, "HasPropertySets", ()) or ():
            if not definition.is_a("IfcPropertySet"):
                continue
            for prop in definition.HasProperties:
                values.setdefault(prop.Name, _unwrap_value(getattr(prop, "NominalValue", None)))
    return values


def _to_decimal(value: Any) -> Decimal | None:
    """Преобразует числовое значение IFC в Decimal."""

    if value is None or value == "":
        return None
    try:
        return Decimal(str(value).replace(",", "."))
    except (InvalidOperation, ValueError):
        return None


def _to_bool(value: Any) -> bool | None:
    """Преобразует IFC Boolean или текстовое значение в bool."""

    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"true", "1", "yes", "да"}


@log_dataframe
def extract_spaces(model: ifcopenshell.file) -> pd.DataFrame:
    """Извлекает зоны и помещения IFC в табличную форму.

    :param model: Открытая IFC-модель.
    :returns: DataFrame с атрибутами помещений и зон из `RusSet_*`.
    """

    rows: list[dict[str, Any]] = []
    for entity in model.by_type("IfcSpace"):
        props = _entity_properties(entity)
        rows.append(
            {
                "ifc_id": entity.id(),
                "global_id": entity.GlobalId,
                "name": entity.Name,
                "rus_name": props.get("RUS_Name"),
                "area": _to_decimal(props.get("RUS_Area")),
                "volume": _to_decimal(props.get("RUS_Volume")),
                "basement": _to_bool(props.get("RUS_BasementFloor")),
                "ground_floor": _to_bool(props.get("RUS_GroundFloor")),
                "residential": _to_bool(props.get("RUS_ResidentalRoom")),
                "group_type": props.get("RUS_Group_Type"),
                "group_number": props.get("RUS_Group_Number"),
                "zone": props.get("RUS_Zone"),
                "object": props.get("RUS_Object"),
                "object_sub": props.get("RUS_Object_Sub"),
            }
        )
    result = pd.DataFrame(rows)
    log.info(f"Извлечено помещений/зон: {len(result)}", name="tep")
    return result


def _sum_rows(rows: pd.DataFrame) -> Decimal | None:
    """Суммирует поле area в наборе строк."""

    if rows.empty or rows["area"].dropna().empty:
        return None
    return sum((value for value in rows["area"] if value is not None), Decimal("0"))


def _filter_zone(spaces: pd.DataFrame, name: str) -> pd.DataFrame:
    """Выбирает зоны по нормативному имени или категории зоны."""

    return spaces[(spaces["rus_name"] == name) | (spaces["zone"] == name)]


def _geometry_height(model: ifcopenshell.file) -> Decimal | None:
    """Рассчитывает высоту модели по мировому геометрическому bounding box."""

    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)
    minimum: float | None = None
    maximum: float | None = None
    elements = model.by_type("IfcElement")
    iterator = ifcopenshell.geom.iterator(settings, model, include=elements)
    if not iterator.initialize():
        return None
    while True:
        shape = iterator.get()
        if shape is not None and shape.geometry is not None:
            vertices = shape.geometry.verts
            if vertices:
                z_values = vertices[2::3]
                minimum = min(minimum, min(z_values)) if minimum is not None else min(z_values)
                maximum = max(maximum, max(z_values)) if maximum is not None else max(z_values)
        if not iterator.next():
            break
    if minimum is None or maximum is None:
        return None
    return Decimal(str((maximum - minimum) / 1000))


def _first_property(model: ifcopenshell.file, entity_type: str, property_name: str) -> TepValue:
    """Получает первое значение свойства среди сущностей указанного IFC-класса."""

    for entity in model.by_type(entity_type):
        value = _entity_properties(entity).get(property_name)
        if value not in (None, ""):
            return TepValue(value, "", f"{entity_type}.{property_name}")
    return TepValue(None, "", "", "missing", f"Свойство {property_name} не найдено")


def _required_metadata(model: ifcopenshell.file) -> dict[str, TepValue]:
    """Собирает общие сведения объекта из атрибутов IFC."""

    building = model.by_type("IfcBuilding")
    ifc_building = building[0] if building else None
    props = _entity_properties(ifc_building) if ifc_building else {}
    name = props.get("RUS_ProjectName") or getattr(ifc_building, "Name", None)
    address = props.get("RUS_BuildingAddress")
    function = props.get("RUS_FNO")
    return {
        "ObjectName": TepValue(name, "", "IfcBuilding.RUS_ProjectName" if name else "", "calculated" if name else "missing"),
        "Address": TepValue(address, "", "IfcBuilding.RUS_BuildingAddress" if address else "", "calculated" if address else "missing"),
        "FunctionalPurpose": TepValue(function, "", "IfcBuilding.RUS_FNO" if function else "", "calculated" if function else "missing"),
    }


@log_dataframe
def calculate_tep(model: ifcopenshell.file) -> pd.DataFrame:
    """Рассчитывает ТЭП из атрибутов выбранной IFC-модели.

    :param model: Открытая IFC-модель ЦИМ АГР.
    :returns: DataFrame со всеми полями формы сайта, значениями,
        единицами измерения, источниками и статусами.
    """

    spaces = extract_spaces(model)
    spp = _filter_zone(spaces, "СПП в ГНС")
    total = _filter_zone(spaces, "Общая площадь")
    residential_spp = spp[spp["residential"] == True]
    non_residential_spp = spp[spp["residential"] == False]
    above_total = total[total["basement"] == False]
    underground_total = total[total["basement"] == True]
    non_residential_total_above = total[(total["basement"] == False) & (total["residential"] == False)]
    apartment_spaces = spaces[spaces["group_type"] == "Квартира"]
    summer_names = {"Балкон", "Лоджия", "Терраса", "Веранда", "Холодная кладовая", "Тамбур"}
    apartment_without_summer = apartment_spaces[~apartment_spaces["rus_name"].isin(summer_names)]
    building_props = _required_metadata(model)

    def calculated(value: Any, unit: str, source: str, comment: str = "") -> TepValue:
        return TepValue(value, unit, source, "calculated" if value is not None else "missing", comment)

    def unavailable(comment: str) -> TepValue:
        return TepValue(None, "", "", "unavailable", comment)

    height = _first_property(model, "IfcBuilding", "RUS_Height")
    if height.value is not None:
        height = calculated(_to_decimal(height.value) / Decimal("1000"), "м", height.source)
    else:
        log.warning(
            "Атрибут высоты IfcBuilding.RUS_Height не найден; геометрический расчёт отключён "
            "для базового запуска из-за высокой стоимости обработки модели",
            name="tep",
        )
        height = unavailable("В IFC не найден RUS_Height; геометрический fallback пока не включён")

    storeys = model.by_type("IfcBuildingStorey")
    floor_count = sum(1 for storey in storeys if not any(
        row["basement"] is True for row in spaces[spaces["object"].notna()].to_dict("records")
    )) if False else len(storeys)
    floor_count = len([storey for storey in storeys if getattr(storey, "Elevation", 0) is not None and float(storey.Elevation) >= 0])
    apartment_count = len(set(apartment_spaces["group_number"].dropna().astype(str)))
    site_area = _first_property(model, "IfcSite", "RUS_AreaSquare")
    area_value = _to_decimal(site_area.value) / Decimal("10000") if site_area.value is not None else None

    cadastral = _first_property(model, "IfcSite", "RUS_LandID")
    building_entities = model.by_type("IfcBuilding")
    unique_object_id = calculated(
        building_entities[0].GlobalId if building_entities else None,
        "",
        "IfcBuilding.GlobalId" if building_entities else "",
        "GlobalId IFC используется как доступный идентификатор модели; официальный УИН может задаваться отдельно",
    )
    building_function = building_props["FunctionalPurpose"]
    non_residential_object = calculated(
        str(building_function.value).strip().lower() == "нежилое здание"
        if building_function.value is not None
        else None,
        "",
        building_function.source,
        "Признак определён по IfcBuilding.RUS_FNO",
    )

    values: dict[str, TepValue] = {
        **building_props,
        "UniqueObjectId": unique_object_id,
        "RegistrationNumber": unavailable("В IFC не найден регистрационный номер"),
        "Date": unavailable("В IFC не найдена дата формирования сведений"),
        "CadastralNumber": cadastral,
        "ProjectOrganization": unavailable("В IFC нет сведений о проектной организации"),
        "ProjectTeamLead": unavailable("В IFC нет руководителя авторского коллектива"),
        "Member": unavailable("В IFC нет состава авторского коллектива"),
        "WorkTypes": unavailable("В IFC нет вида работ"),
        "BuildingHeight": height,
        "Area": calculated(area_value, "га", site_area.source, "RUS_AreaSquare преобразован из м² в га") if area_value is not None else unavailable("В IFC не найден RUS_AreaSquare"),
        "TotalFloorArea": calculated(_sum_rows(spp), "м²", "IfcSpace.RUS_Name=СПП в ГНС / RUS_Area"),
        "ResidentialPartArea": calculated(_sum_rows(spp), "м²", "IfcSpace.RUS_Name=СПП в ГНС / RUS_Area", "Жилой объект определяется по RUS_FNO"),
        "ResidentialPartOfResidentialArea": calculated(_sum_rows(residential_spp), "м²", "IfcSpace.RUS_ResidentalRoom=true / RUS_Area"),
        "NonResidentialPartOfResidentialArea": calculated(_sum_rows(non_residential_spp), "м²", "IfcSpace.RUS_ResidentalRoom=false / RUS_Area"),
        "NonResidentialObjectsArea": unavailable("Для отдельного нежилого объекта нужна соответствующая IFC-модель"),
        "TotalObjectArea": calculated(_sum_rows(total), "м²", "IfcSpace.RUS_Name=Общая площадь / RUS_Area"),
        "AboveGroundArea": calculated(_sum_rows(above_total), "м²", "IfcSpace.RUS_Name=Общая площадь + RUS_BasementFloor=false / RUS_Area"),
        "UndergroundArea": calculated(_sum_rows(underground_total), "м²", "IfcSpace.RUS_Name=Общая площадь + RUS_BasementFloor=true / RUS_Area"),
        "NonResidentialAboveGroundAreaNonResidentialPart": calculated(_sum_rows(non_residential_total_above), "м²", "IfcSpace.RUS_Name=Общая площадь + RUS_ResidentalRoom=false / RUS_Area"),
        "NonResidentialAboveGroundAreaNonResidentialObjects": unavailable("Для отдельного нежилого объекта нужна соответствующая IFC-модель"),
        "ApartmentsAreaExSummerPremises": calculated(_sum_rows(apartment_without_summer), "м²", "IfcSpace.RUS_Group_Number + исключение летних помещений"),
        "ApartmentsCount": calculated(apartment_count or None, "шт.", "уникальные IfcSpace.RUS_Group_Number"),
        "GreenArea": unavailable("Площадь озеленения находится в ЦИМ БиО"),
        "ChildrenPlaygroundArea": unavailable("Детские площадки находятся в ЦИМ БиО"),
        "RecreationsArea": unavailable("Площадки отдыха находятся в ЦИМ БиО"),
        "ObjectHeight": height,
        "FloorsCount": calculated(floor_count or None, "шт.", "IfcBuildingStorey.Elevation >= 0"),
        "ApartmentsArea": calculated(_sum_rows(apartment_spaces), "м²", "IfcSpace.RUS_Group_Number / RUS_Area"),
        "TotalApartmentsArea": calculated(_sum_rows(apartment_spaces), "м²", "IfcSpace.RUS_Group_Number / RUS_Area"),
        "InformationStructuresCount": unavailable("В IFC не найдены информационные конструкции"),
        "InformationStructuresParameters": unavailable("В IFC нет параметров информационных конструкций"),
        "NonResidentialObject": non_residential_object,
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


def generate_tep(ifc_path: Path, output_dir: Path) -> dict[str, Path]:
    """Формирует Excel с ТЭП для IFC-файла.

    :param ifc_path: Путь к IFC-модели.
    :param output_dir: Каталог для результатов.
    :returns: Словарь с путём к XLSX-файлу.
    """

    output_dir.mkdir(parents=True, exist_ok=True)
    log.configure(str(output_dir / "tep_generation.log"))
    log.init("Генерация ТЭП из IFC")
    log.info(f"Открытие IFC: {ifc_path}", name="tep")
    model = ifcopenshell.open(str(ifc_path))
    tep = calculate_tep(model)
    stem = ifc_path.stem
    excel_path = export_tep_excel(tep, output_dir / f"{stem}_tep.xlsx")
    log.info(f"Генерация ТЭП завершена: {len(tep)} полей", name="tep")
    return {"excel": excel_path}


if __name__ == "__main__":
    IFC_PATH = Path(r"samples\НН_ОБ01_П00_АР_АГР.ifc")
    OUTPUT_DIR = Path(r".output")

    generate_tep(
        ifc_path=IFC_PATH,
        output_dir=OUTPUT_DIR,
    )
