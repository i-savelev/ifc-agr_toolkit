"""Проверка геометрии помещений IfcSpace и сравнение площадей.

Модуль проверяет наличие геометрического представления у элементов
IfcSpace, вычисляет площадь проекции на горизонтальную плоскость XY
и сравнивает вычисленную площадь с атрибутом RUS_Area.
"""

from __future__ import annotations

from pathlib import Path
import sys
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
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
class SpaceGeometryCheck:
    """Результат проверки геометрии одного помещения IfcSpace.

    :param guid: Глобальный идентификатор сущности IFC (GlobalId).
    :param name: Имя помещения (LongName или RUS_Name).
    :param revit_id: Идентификатор Revit (часть Type Name после двоеточия).
    :param level_name: Имя строительного уровня (BuildingStorey).
    :param model_name: Название IFC-файла (без расширения).
    :param has_geometry: Есть ли геометрическое представление.
    :param geometry_area: Вычисленная геометрическая площадь проекции (м²).
    :param rus_area: Значение RUS_Area из атрибутов (м²).
    :param area_difference: Разница между геометрической и RUS_Area (м²).
    :param relative_error: Относительная ошибка (доля, округлена до 4 знаков).
    :param area_match: Совпадают ли площади в пределах допуска.
    :param comment: Дополнительный комментарий.
    :param rus_...: Все RUS-параметры помещения.
    """

    guid: str
    name: str
    number: str | None
    revit_id: str | None
    level_name: str | None
    model_name: str | None
    has_geometry: bool
    geometry_area: Decimal | None
    rus_area: Decimal | None
    area_difference: Decimal | None
    relative_error: Decimal | None
    area_match: bool | None
    comment: str = ""

    # RUS параметры
    rus_studio_apart: bool | None = None
    rus_basement_floor: bool | None = None
    rus_ground_floor: bool | None = None
    rus_residental_room: bool | None = None
    rus_area_square: Decimal | None = None
    rus_land_id: str | None = None
    rus_mssk_space_zone_code: str | None = None
    rus_name: str | None = None
    rus_mssk_space_zone_name: str | None = None
    rus_object: str | None = None
    rus_object_sub: str | None = None
    rus_zone: str | None = None
    rus_group_type: str | None = None
    rus_group_number: str | None = None
    rus_number: str | None = None


def _has_geometry(
    model: ifcopenshell.file,
    entity: Any,
) -> bool:
    """Проверяет, есть ли у сущности IfcSpace геометрическое представление.

    :param model: Открытая IFC-модель.
    :param entity: Экземпляр IfcSpace.
    :returns: True, если геометрия есть, False — если геометрия отсутствует.
    """

    representations = getattr(entity, "Representation", None)
    if representations is None:
        return False

    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)
    try:
        shape = ifcopenshell.geom.create_shape(settings=settings, inst=entity)
        return bool(hasattr(shape, "geometry") and shape.geometry.verts)
    except (ValueError, Exception) as error:
        log.warning(f"{entity.id()} '{entity.Name}': ошибка проверки геометрии: {error}", name="space_check")
        return False


def _compute_space_projected_area(
    model: ifcopenshell.file,
    entity: Any,
) -> Decimal | None:
    """Вычисляет площадь проекции геометрии IfcSpace на горизонтальную плоскость XY.

    Площадь считается как сумма площадей горизонтальных граней (пол/потолок).
    Вертикальные стены игнорируются.

    :param model: Открытая IFC-модель.
    :param entity: Экземпляр IfcSpace.
    :returns: Площадь проекции в м² или None.
    """

    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)
    try:
        shape = ifcopenshell.geom.create_shape(settings=settings, inst=entity)
    except Exception as error:
        log.warning(f"{entity.id()} '{entity.Name}': ошибка извлечения геометрии: {error}", name="space_check")
        return None

    geometry = getattr(shape, "geometry", None)
    if geometry is None:
        log.warning(f"{entity.id()} '{entity.Name}': геометрия отсутствует", name="space_check")
        return None

    vertices = geometry.verts
    faces = geometry.faces
    if not vertices or not faces:
        log.warning(
            f"{entity.id()} '{entity.Name}': нет вершин или граней "
            f"(verts={len(vertices) if vertices else 0}, faces={len(faces) if faces else 0})",
            name="space_check",
        )
        return None

    points = [
        (vertices[i], vertices[i + 1], vertices[i + 2])
        for i in range(0, len(vertices), 3)
    ]

    total_area = Decimal("0")
    horizontal_count = 0

    for face_start in range(0, len(faces), 3):
        indices = [faces[face_start + offset] for offset in range(3)]
        try:
            p1, p2, p3 = points[indices[0]], points[indices[1]], points[indices[2]]
        except (IndexError, KeyError):
            continue

        # Вычисляем нормаль треугольника
        dx1 = p2[0] - p1[0]
        dy1 = p2[1] - p1[1]
        dz1 = p2[2] - p1[2]
        dx2 = p3[0] - p1[0]
        dy2 = p3[1] - p1[1]
        dz2 = p3[2] - p1[2]

        nx = dy1 * dz2 - dz1 * dy2
        ny = dz1 * dx2 - dx1 * dz2
        nz = dx1 * dy2 - dy1 * dx2

        length = (nx * nx + ny * ny + nz * nz) ** 0.5
        if length == 0:
            continue

        nz_normalized = abs(nz / length)

        # Если нормаль близка к вертикали (nz > 0.9), грань горизонтальна
        if nz_normalized > 0.9:
            x1, y1 = p1[0], p1[1]
            x2, y2 = p2[0], p2[1]
            x3, y3 = p3[0], p3[1]

            area = abs(x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2)) / 2.0
            total_area += Decimal(str(area))
            horizontal_count += 1

    if horizontal_count == 0:
        log.warning(
            f"{entity.id()} '{entity.Name}': нет горизонтальных граней "
            f"(всего граней: {len(faces)}, вершин: {len(vertices)})"
            f" — все треугольники наклонные (nz_normalized <= 0.9)",
            name="space_check",
        )
        return None

    # Пол и потолок проецируются одинаково — делим пополам
    total_area = total_area / 2

    log.info(
        f"{entity.id()} '{entity.Name}': проекция из {horizontal_count} горизонт. граней → {total_area:.3f} м²",
        name="space_check",
    )
    return total_area


def _to_decimal(value: Any) -> Decimal | None:
    """Преобразует числовое значение IFC в Decimal."""

    if value is None or value == "":
        return None
    try:
        return Decimal(str(value).replace(",", "."))
    except (ValueError, InvalidOperation):
        return None


def _to_entity_properties(entity: Any) -> dict[str, Any]:
    """Возвращает свойства экземпляра и назначенного ему типа IFC."""

    values: dict[str, Any] = {}
    for relation in getattr(entity, "IsDefinedBy", ()) or ():
        if not relation.is_a("IfcRelDefinesByProperties"):
            continue
        definition = relation.RelatingPropertyDefinition
        if not definition.is_a("IfcPropertySet"):
            continue
        for prop in definition.HasProperties:
            unwrap = getattr(prop, "NominalValue", None)
            values[prop.Name] = getattr(unwrap, "wrappedValue", unwrap)

    for relation in getattr(entity, "IsTypedBy", ()) or ():
        type_entity = relation.RelatingType
        for definition in getattr(type_entity, "HasPropertySets", ()) or ():
            if not definition.is_a("IfcPropertySet"):
                continue
            for prop in definition.HasProperties:
                values.setdefault(prop.Name, getattr(prop, "NominalValue", None))
    return values


def _extract_rus_properties(props: dict[str, Any]) -> dict[str, Any]:
    """Извлекает все RUS-параметры из словаря свойств.

    :param props: Словарь свойств IFC-сущности.
    :returns: Словарь с RUS-параметрами.
    """

    def _to_bool(value: Any) -> bool | None:
        """Преобразует IFC Boolean или текстовое значение в bool."""

        if value is None or value == "":
            return None
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in {"true", "1", "yes", "да"}

    return {
        "RUS_StudioApart": _to_bool(props.get("RUS_StudioApart")),
        "RUS_BasementFloor": _to_bool(props.get("RUS_BasementFloor")),
        "RUS_GroundFloor": _to_bool(props.get("RUS_GroundFloor")),
        "RUS_ResidentalRoom": _to_bool(props.get("RUS_ResidentalRoom")),
        "RUS_AreaSquare": _to_decimal(props.get("RUS_AreaSquare")),
        "RUS_LandID": props.get("RUS_LandID"),
        "RUS_MSSK_SpaceAndZone_Code": props.get("RUS_MSSK_SpaceAndZone_Code"),
        "RUS_Name": props.get("RUS_Name"),
        "RUS_MSSK_SpaceAndZone_Name": props.get("RUS_MSSK_SpaceAndZone_Name"),
        "RUS_Area": _to_decimal(props.get("RUS_Area")),
        "RUS_Object": props.get("RUS_Object"),
        "RUS_Object_Sub": props.get("RUS_Object_Sub"),
        "RUS_Zone": props.get("RUS_Zone"),
        "RUS_Group_Type": props.get("RUS_Group_Type"),
        "RUS_Group_Number": props.get("RUS_Group_Number"),
        "RUS_Number": props.get("RUS_Number"),
    }


def _get_storey_name(entity: Any, model: ifcopenshell.file) -> str | None:
    """Возвращает имя строительного уровня (BuildingStorey) для помещения.

    :param entity: Экземпляр IfcSpace.
    :param model: Открытая IFC-модель.
    :returns: Имя уровня или None.
    """

    for rel in model.by_type("IfcRelAggregates"):
        related = getattr(rel, "RelatedObjects", None)
        if related and entity in related:
            relating = getattr(rel, "RelatingObject", None)
            if relating and relating.is_a("IfcBuildingStorey"):
                return getattr(relating, "Name", None)

    return None


def _get_model_name(ifc_path: Path) -> str:
    """Возвращает имя файла IFC без расширения.

    :param ifc_path: Путь к IFC-файлу.
    :returns: Имя файла без расширения.
    """

    return ifc_path.stem


def _parse_revit_id(entity: Any) -> str | None:
    """Извлекает Revit ID из имени типа помещения.

    Формат имени типа: 'ПОН 4 (Магазин) 2-4.1:1694085'.
    Revit ID — это часть после последнего двоеточия.

    :param entity: Экземпляр IfcSpace.
    :returns: Revit ID или None.
    """

    # Получаю тип через IsTypedBy
    for rel in getattr(entity, "IsTypedBy", ()) or []:
        if rel.is_a("IfcRelDefinesByType"):
            type_entity = rel.RelatingType
            type_name = getattr(type_entity, "Name", None)
            if type_name and ":" in str(type_name):
                parts = str(type_name).split(":")
                revit_id = parts[-1].strip()
                if revit_id:
                    return revit_id

    return None


def _check_space_geometry(
    model: ifcopenshell.file,
    entity: Any,
    tolerance: float = 0.01,
    ifc_path: Path | None = None,
) -> SpaceGeometryCheck:
    """Проверяет геометрию помещения и сравнивает площадь с RUS_Area.

    :param model: Открытая IFC-модель.
    :param entity: Экземпляр IfcSpace.
    :param tolerance: Допустимая относительная разница в долях (по умолчанию 1%).
    :param ifc_path: Путь к IFC-файлу для получения имени модели.
    :returns: Структура с результатами проверки.
    """

    props = _to_entity_properties(entity)
    rus_props = _extract_rus_properties(props)

    # Имя помещения: LongName
    name = getattr(entity, "LongName", None)
    number = getattr(entity, "Name", None)
    guid = getattr(entity, "GlobalId", None)
    has_geom = _has_geometry(model, entity)
    geom_area = None
    difference = None
    rel_error = None
    comment = ""

    revit_id = _parse_revit_id(entity)
    level_name = _get_storey_name(entity, model)

    # Имя модели - имя файла без расширения
    model_name = _get_model_name(ifc_path) if ifc_path else None

    if not has_geom:
        comment = "Геометрия отсутствует"
        log.info(
            f"{entity.id()} '{entity.Name}': {comment}, RUS_Area={rus_props.get('RUS_Area')}",
            name="space_check",
        )
        return SpaceGeometryCheck(
            guid=guid or "",
            name=name,
            number=number,
            revit_id=revit_id,
            level_name=level_name,
            model_name=model_name,
            has_geometry=False,
            geometry_area=None,
            rus_area=rus_props.get("RUS_Area"),
            area_difference=None,
            relative_error=None,
            area_match=None,
            comment=comment,
            rus_studio_apart=rus_props.get("RUS_StudioApart"),
            rus_basement_floor=rus_props.get("RUS_BasementFloor"),
            rus_ground_floor=rus_props.get("RUS_GroundFloor"),
            rus_residental_room=rus_props.get("RUS_ResidentalRoom"),
            rus_area_square=rus_props.get("RUS_AreaSquare"),
            rus_land_id=rus_props.get("RUS_LandID"),
            rus_mssk_space_zone_code=rus_props.get("RUS_MSSK_SpaceAndZone_Code"),
            rus_name=rus_props.get("RUS_Name"),
            rus_mssk_space_zone_name=rus_props.get("RUS_MSSK_SpaceAndZone_Name"),
            rus_object=rus_props.get("RUS_Object"),
            rus_object_sub=rus_props.get("RUS_Object_Sub"),
            rus_zone=rus_props.get("RUS_Zone"),
            rus_group_type=rus_props.get("RUS_Group_Type"),
            rus_group_number=rus_props.get("RUS_Group_Number"),
            rus_number=rus_props.get("RUS_Number"),
        )

    geom_area = _compute_space_projected_area(model, entity)

    if geom_area is None:
        comment = "Не удалось вычислить площадь — нет горизонтальных граней"
        log.info(
            f"{entity.id()} '{entity.Name}': {comment}, RUS_Area={rus_props.get('RUS_Area')}",
            name="space_check",
        )
        return SpaceGeometryCheck(
            guid=guid or "",
            name=name,
            number=rus_props.get("RUS_Number"),
            revit_id=revit_id,
            level_name=level_name,
            model_name=model_name,
            has_geometry=False,
            geometry_area=None,
            rus_area=rus_props.get("RUS_Area"),
            area_difference=None,
            relative_error=None,
            area_match=None,
            comment=comment,
            rus_studio_apart=rus_props.get("RUS_StudioApart"),
            rus_basement_floor=rus_props.get("RUS_BasementFloor"),
            rus_ground_floor=rus_props.get("RUS_GroundFloor"),
            rus_residental_room=rus_props.get("RUS_ResidentalRoom"),
            rus_area_square=rus_props.get("RUS_AreaSquare"),
            rus_land_id=rus_props.get("RUS_LandID"),
            rus_mssk_space_zone_code=rus_props.get("RUS_MSSK_SpaceAndZone_Code"),
            rus_name=rus_props.get("RUS_Name"),
            rus_mssk_space_zone_name=rus_props.get("RUS_MSSK_SpaceAndZone_Name"),
            rus_object=rus_props.get("RUS_Object"),
            rus_object_sub=rus_props.get("RUS_Object_Sub"),
            rus_zone=rus_props.get("RUS_Zone"),
            rus_group_type=rus_props.get("RUS_Group_Type"),
            rus_group_number=rus_props.get("RUS_Group_Number"),
            rus_number=rus_props.get("RUS_Number"),
        )

    rus_area = rus_props.get("RUS_Area")
    if rus_area is not None:
        difference = geom_area - rus_area
        if rus_area != 0:
            rel_error = abs(difference) / abs(rus_area)
        else:
            rel_error = Decimal("0") if difference == 0 else None

        if rel_error is not None and rel_error <= Decimal(str(tolerance)):
            comment = "Геометрическая площадь совпадает с RUS_Area в пределах допуска"
        else:
            comment = (
                f"Расхождение {difference:+.3f} м² ({float(rel_error)*100:.1f}%)"
                if rel_error is not None
                else f"Расхождение {difference:+.3f} м² — RUS_Area=0"
            )
    else:
        comment = "RUS_Area отсутствует — геометрическая площадь вычислена"

    # Округление значений
    geom_area_rounded = round(float(geom_area), 4) if geom_area is not None else None
    area_diff_rounded = round(float(difference), 4) if difference is not None else None
    rel_error_rounded = round(float(rel_error), 4) if rel_error is not None else None

    diff_str = '{:+.3f}'.format(float(difference)) if difference is not None else 'N/A'
    rel_str = '{:.1f}'.format(float(rel_error)*100) if rel_error is not None else 'N/A'
    geom_str = '{:.3f}'.format(float(geom_area)) if geom_area is not None else 'N/A'
    rus_str = '{:.3f}'.format(float(rus_area)) if rus_area is not None else 'N/A'

    log.info(
        f"{entity.id()} '{entity.Name}': geom={geom_str} м², "
        f"RUS_Area={rus_str}, diff={diff_str} м², "
        f"rel={rel_str}% → match={rel_error is not None and rel_error <= Decimal(str(tolerance))}",
        name="space_check",
    )
    return SpaceGeometryCheck(
        guid=guid or "",
        name=name,
        number=number,
        revit_id=revit_id,
        level_name=level_name,
        model_name=model_name,
        has_geometry=True,
        geometry_area=Decimal(str(geom_area_rounded)) if geom_area_rounded is not None else None,
        rus_area=rus_area,
        area_difference=Decimal(str(area_diff_rounded)) if area_diff_rounded is not None else None,
        relative_error=Decimal(str(rel_error_rounded)) if rel_error_rounded is not None else None,
        area_match=rel_error is not None and rel_error <= Decimal(str(tolerance)),
        comment=comment,
        rus_studio_apart=rus_props.get("RUS_StudioApart"),
        rus_basement_floor=rus_props.get("RUS_BasementFloor"),
        rus_ground_floor=rus_props.get("RUS_GroundFloor"),
        rus_residental_room=rus_props.get("RUS_ResidentalRoom"),
        rus_area_square=rus_props.get("RUS_AreaSquare"),
        rus_land_id=rus_props.get("RUS_LandID"),
        rus_mssk_space_zone_code=rus_props.get("RUS_MSSK_SpaceAndZone_Code"),
        rus_name=rus_props.get("RUS_Name"),
        rus_mssk_space_zone_name=rus_props.get("RUS_MSSK_SpaceAndZone_Name"),
        rus_object=rus_props.get("RUS_Object"),
        rus_object_sub=rus_props.get("RUS_Object_Sub"),
        rus_zone=rus_props.get("RUS_Zone"),
        rus_group_type=rus_props.get("RUS_Group_Type"),
        rus_group_number=rus_props.get("RUS_Group_Number"),
        rus_number=rus_props.get("RUS_Number"),
    )


@log_dataframe
def check_space_geometries(
    model: ifcopenshell.file,
    ifc_path: Path | None = None,
    tolerance: float = 0.1,
    output_path: Path | None = None,
) -> pd.DataFrame:
    """Проверяет все помещения IfcSpace на наличие геометрии и согласованность площадей.

    Для каждого IfcSpace:
    1. Проверяет наличие геометрического представления.
    2. Вычисляет площадь проекции на горизонтальную плоскость XY.
    3. Сравнивает вычисленную площадь с атрибутом RUS_Area.

    :param model: Открытая IFC-модель.
    :param ifc_path: Путь к IFC-файлу (для имени модели).
    :param tolerance: Допустимая относительная разница (доля, по умолчанию 0.1 = 10%).
    :param output_path: Путь к файлу Excel для сохранения результатов.
    :returns: DataFrame с результатами проверки всех помещений.
    """

    spaces = model.by_type("IfcSpace")
    if not spaces:
        log.warning("В модели не найдено ни одного IfcSpace", name="space_check")
        return pd.DataFrame()

    results: list[dict[str, Any]] = []
    for entity in spaces:
        check = _check_space_geometry(model, entity, tolerance=tolerance, ifc_path=ifc_path)

        # Формируем строку DataFrame
        row: dict[str, Any] = {
            "GUID": check.guid,
            "Наименование": check.name,
            "Номер": check.number,
            "Revit ID": check.revit_id,
            "Уровень": check.level_name,
            "Модель": check.model_name,
            "Есть геометрия": "ДА" if check.has_geometry else "НЕТ",
            "Геом. площадь (м²)": check.geometry_area,
            "RUS_Area (м²)": check.rus_area,
            "Разница (м²)": check.area_difference,
            "Расхождение (%)": float(check.relative_error) * 100 if check.relative_error is not None else None,
            "Совпадение площадей": "ДА" if check.area_match is True else ("НЕТ" if check.area_match is False else ""),
            "Комментарий": check.comment,
            # RUS параметры
            "RUS_StudioApart": check.rus_studio_apart,
            "RUS_BasementFloor": check.rus_basement_floor,
            "RUS_GroundFloor": check.rus_ground_floor,
            "RUS_ResidentalRoom": check.rus_residental_room,
            "RUS_AreaSquare": check.rus_area_square,
            "RUS_LandID": check.rus_land_id,
            "RUS_MSSK_SpaceAndZone_Code": check.rus_mssk_space_zone_code,
            "RUS_Name": check.rus_name,
            "RUS_MSSK_SpaceAndZone_Name": check.rus_mssk_space_zone_name,
            "RUS_Object": check.rus_object,
            "RUS_Object_Sub": check.rus_object_sub,
            "RUS_Zone": check.rus_zone,
            "RUS_Group_Type": check.rus_group_type,
            "RUS_Group_Number": check.rus_group_number,
            "RUS_Number": check.rus_number,
        }
        results.append(row)

    df = pd.DataFrame(results)

    # Статистика
    match_count = df[df["Совпадение площадей"] == "ДА"].shape[0]
    mismatch_count = df[df["Совпадение площадей"] == "НЕТ"].shape[0]
    missing_geom = df[df["Есть геометрия"] == "НЕТ"].shape[0]
    log.info(
        f"Проверено {len(df)} IfcSpace: совпало={match_count}, не совпало={mismatch_count}, без геометрии={missing_geom}",
        name="space_check",
    )

    if output_path is not None:
        _export_check_excel(df, output_path)

    return df


def _export_check_excel(df: pd.DataFrame, output_path: Path) -> Path:
    """Экспортирует результаты проверки геометрии в Excel.

    :param df: DataFrame с результатами проверки.
    :param output_path: Путь к выходному файлу.
    :returns: Путь к созданному файлу.
    """

    output_path.parent.mkdir(parents=True, exist_ok=True)

    export_df = df.copy()
    export_df.to_excel(output_path, index=False, sheet_name="Проверка геометрии")

    workbook = load_workbook(output_path)
    worksheet = workbook["Проверка геометрии"]

    # Заголовки жирным
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    # Фильтры
    worksheet.auto_filter.ref = worksheet.dimensions

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

    workbook.save(output_path)
    log.info(f"Результаты сохранены в Excel: {output_path}", name="space_check")
    return output_path


if __name__ == "__main__":
    # Измените этот путь перед запуском файла напрямую.
    IFC_PATH = Path(r"samples\НН_К00_С00_АР_АГР_new.ifc")
    OUTPUT_DIR = Path(r".output")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log.configure(str(OUTPUT_DIR / "space_check.log"))
    log.init("Проверка геометрии IfcSpace")

    model = ifcopenshell.open(str(IFC_PATH))
    excel_path = OUTPUT_DIR / "space_geometry_check.xlsx"
    check = check_space_geometries(model, ifc_path=IFC_PATH, tolerance=0.1, output_path=excel_path)

    if not check.empty:
        match = check[check["Совпадение площадей"] == "ДА"]
        mismatch = check[check["Совпадение площадей"] == "НЕТ"]
        missing_geom = check[check["Есть геометрия"] == "НЕТ"]

        log.info(f"Помещений с совпадением площадей: {len(match)}", name="space_check")
        log.info(f"Помещений без совпадения площадей: {len(mismatch)}", name="space_check")
        log.info(f"Помещений без геометрии: {len(missing_geom)}", name="space_check")

        if not missing_geom.empty:
            log.info("Список помещений без геометрии:", name="space_check")
            for _, row in missing_geom.iterrows():
                log.data({
                    "guid": str(row["GUID"]),
                    "name": str(row["Наименование"]),
                    "level": str(row["Уровень"]),
                }, name="space_check", label=f"{row['GUID']}")

        if not mismatch.empty:
            log.info("Список помещений без совпадения площадей (первые 5):", name="space_check")
            for _, row in mismatch.head(5).iterrows():
                log.data({
                    "guid": str(row["GUID"]),
                    "name": str(row["Наименование"]),
                    "geom_area": float(row["Геом. площадь (м²)"]) if row["Геом. площадь (м²)"] is not None else None,
                    "rus_area": float(row["RUS_Area (м²)"]) if row["RUS_Area (м²)"] is not None else None,
                    "diff": float(row["Разница (м²)"]) if row["Разница (м²)"] is not None else None,
                    "rel_error_pct": float(row["Расхождение (%)"]) if row["Расхождение (%)"] is not None else None,
                }, name="space_check", label=f"{row['GUID']}")

    log.info(f"Проверка геометрии завершена. Excel: {excel_path}", name="space_check")
