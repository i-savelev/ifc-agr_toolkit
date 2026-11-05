"""Проверка геометрии помещений IfcSpace и сравнение площадей.

Модуль собирает единый DataFrame помещений/зон, последовательно обогащая
его данными из базовых полей, RUS-свойств и геометрических расчётов.
"""

from __future__ import annotations

import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

import ifcopenshell
import ifcopenshell.geom
import pandas as pd

# При запуске файла напрямую Python добавляет в sys.path только каталог src.
# Добавляем корень проекта, чтобы были доступны logger.py и пакет src.
if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.utils.logger import Logger as log
from src.utils.logger_decorator import log_dataframe


# ============================================================
# Универсальные функции извлечения атрибутов
# ============================================================


def get_base_fields(
    entity: Any,
    model: ifcopenshell.file,
    ifc_path: Path | None = None,
) -> dict[str, Any]:
    """Возвращает базовые поля для одного IFC-элемента.

    :param entity: IFC-объект.
    :param model: Открытая IFC-модель.
    :param ifc_path: Путь к IFC-файлу.
    :returns: Словарь с базовыми полями.
    """

    ifc_class = entity.is_a()
    name = getattr(entity, "Name", None)
    long_name = getattr(entity, "LongName", None)
    level_name = _get_storey_name(entity, model)
    model_name = _get_model_name(ifc_path) if ifc_path else None
    guid = getattr(entity, "GlobalId", None)

    # Revit ID: из Type.Name после ":" для IfcSpace, из Name для остальных
    revit_id = None
    for rel in getattr(entity, "IsTypedBy", ()) or []:
        if rel.is_a("IfcRelDefinesByType"):
            type_entity = rel.RelatingType
            type_name = getattr(type_entity, "Name", None)
            if type_name and ":" in str(type_name):
                parts = str(type_name).split(":")
                revit_id = parts[-1].strip() or None
                break

    if revit_id is None and ":" in str(name) if name else False:
        parts = str(name).split(":")
        revit_id = parts[-1].strip() or None

    return {
        "IfcClass": ifc_class,
        "Name": name,
        "LongName": long_name,
        "Level": level_name,
        "Model": model_name,
        "Revit ID": revit_id,
        "GUID": guid,
    }


def add_rus_properties(entity: Any) -> dict[str, Any]:
    """Извлекает все параметры, в названии которых есть "RUS".

    Тип данных определяется по IFC-типу NominalValue:
    - IfcBoolean → bool
    - IfcReal, IfcInteger → Decimal
    - IfcText, IfcIdentifier → str

    :param entity: IFC-объект.
    :returns: Словарь с RUS-параметрами.
    """

    props = _to_entity_properties(entity)
    rus_props: dict[str, Any] = {}

    for relation in getattr(entity, "IsDefinedBy", ()) or []:
        if not relation.is_a("IfcRelDefinesByProperties"):
            continue
        definition = relation.RelatingPropertyDefinition
        if not definition.is_a("IfcPropertySet"):
            continue
        for prop in definition.HasProperties:
            if "RUS" not in prop.Name:
                continue
            nv = getattr(prop, "NominalValue", None)
            if nv is None:
                rus_props[prop.Name] = None
                continue
            ifc_type = nv.is_a()
            wrapped = getattr(nv, "wrappedValue", None)
            if ifc_type == "IfcBoolean":
                rus_props[prop.Name] = bool(wrapped) if wrapped is not None else None
            elif ifc_type in ("IfcReal", "IfcInteger"):
                rus_props[prop.Name] = _to_decimal(wrapped) if wrapped is not None else None
            else:
                rus_props[prop.Name] = str(wrapped) if wrapped is not None else None

    return rus_props


# ============================================================
# Специфичные функции для IfcSpace
# ============================================================


def _has_geometry(
    model: ifcopenshell.file,
    entity: Any,
) -> bool:
    """Проверяет, есть ли у элемента геометрическое представление.

    :param model: Открытая IFC-модель.
    :param entity: Экземпляр IfcSpace.
    :returns: True, если геометрия есть, False — если отсутствует.
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

    for face_start in range(0, len(faces), 3):
        indices = [faces[face_start + offset] for offset in range(3)]
        try:
            p1, p2, p3 = points[indices[0]], points[indices[1]], points[indices[2]]
        except (IndexError, KeyError):
            continue

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

        if nz_normalized > 0.9:
            x1, y1 = p1[0], p1[1]
            x2, y2 = p2[0], p2[1]
            x3, y3 = p3[0], p3[1]

            area = abs(x1 * (y2 - y3) + x2 * (y3 - y1) + x3 * (y1 - y2)) / 2.0
            total_area += Decimal(str(area))

    if total_area == 0:
        log.warning(
            f"{entity.id()} '{entity.Name}': нет горизонтальных граней — все треугольники наклонные (nz_normalized <= 0.9)",
            name="space_check",
        )
        return None

    # Пол и потолок проецируются одинаково — делим пополам
    total_area = total_area / 2

    log.info(
        f"{entity.id()} '{entity.Name}': проекция из горизонт. граней → {total_area:.3f} м²",
        name="space_check",
    )
    return total_area


def add_geometry_and_area(
    entity: Any,
    model: ifcopenshell.file,
    ifc_path: Path | None = None,
    tolerance: float = 0.1,
) -> dict[str, Any]:
    """Добавляет поля геометрии и сравнения с RUS_Area.

    :param entity: Экземпляр IfcSpace.
    :param model: Открытая IFC-модель.
    :param ifc_path: Путь к IFC-файлу.
    :param tolerance: Допустимая относительная разница.
    :returns: Словарь с полями геометрии и сравнения.
    """

    has_geom = _has_geometry(model, entity)
    geom_area = _compute_space_projected_area(model, entity) if has_geom else None

    # RUS_Area
    rus_area = None
    for key, value in _to_entity_properties(entity).items():
        if key == "RUS_Area":
            rus_area = _to_decimal(value)
            break

    difference = None
    rel_error = None
    area_match = None
    has_geom_flag = "ДА" if has_geom else "НЕТ"
    area_match_flag = ""
    comment = ""

    if has_geom and geom_area is None:
        comment = "Не удалось вычислить площадь — нет горизонтальных граней"
        has_geom_flag = "НЕТ"
    elif has_geom and geom_area is not None and rus_area is not None:
        difference = geom_area - rus_area
        if rus_area != 0:
            rel_error = abs(difference) / abs(rus_area)

        if rel_error is not None and rel_error <= Decimal(str(tolerance)):
            area_match_flag = "ДА"
            area_match = True
            comment = "Геометрическая площадь совпадает с RUS_Area в пределах допуска"
        else:
            area_match_flag = "НЕТ"
            area_match = False
            comment = (
                f"Расхождение {difference:+.3f} м² ({float(rel_error)*100:.1f}%)"
                if rel_error is not None
                else f"Расхождение {difference:+.3f} м² — RUS_Area=0"
            )
    elif not has_geom:
        area_match_flag = ""
        area_match = None
        comment = "Геометрия отсутствует"
    elif rus_area is None:
        area_match_flag = ""
        area_match = None
        comment = "RUS_Area отсутствует — геометрическая площадь вычислена"

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
        f"rel={rel_str}% → match={area_match_flag}",
        name="space_check",
    )

    return {
        "Есть геометрия": has_geom_flag,
        "Геом. площадь (м²)": Decimal(str(geom_area_rounded)) if geom_area_rounded is not None else None,
        "RUS_Area (м²)": rus_area,
        "Разница (м²)": Decimal(str(area_diff_rounded)) if area_diff_rounded is not None else None,
        "Расхождение (%)": float(rel_error_rounded) * 100 if rel_error_rounded is not None else None,
        "Совпадение площадей": area_match_flag,
        "Комментарий": comment,
    }


# ============================================================
# Сборка общего DataFrame
# ============================================================


def _get_storey_name(entity: Any, model: ifcopenshell.file) -> str | None:
    """Возвращает имя строительного уровня (BuildingStorey) для помещения."""

    for rel in model.by_type("IfcRelAggregates"):
        related = getattr(rel, "RelatedObjects", None)
        if related and entity in related:
            relating = getattr(rel, "RelatingObject", None)
            if relating and relating.is_a("IfcBuildingStorey"):
                return getattr(relating, "Name", None)
    return None


def _get_model_name(ifc_path: Path) -> str:
    """Возвращает имя файла IFC без расширения."""

    return ifc_path.stem


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


@log_dataframe
def get_spaces_dataframe(
    model: ifcopenshell.file,
    ifc_path: Path | None = None,
    tolerance: float = 0.1,
) -> pd.DataFrame:
    """Возвращает общий DataFrame со всеми помещениями/зонами IfcSpace.

    Для каждого IfcSpace последовательно вызывается:
    1. get_base_fields — базовые поля (IfcClass, Name, LongName, Level, Model, Revit ID, GUID)
    2. add_rus_properties — все RUS-свойства
    3. add_geometry_and_area — геометрия и сравнение с RUS_Area

    :param model: Открытая IFC-модель.
    :param ifc_path: Путь к IFC-файлу (для имени модели).
    :param tolerance: Допустимая относительная разница (доля, по умолчанию 0.1 = 10%).
    :returns: DataFrame с результатами проверки всех помещений.
    """

    spaces = model.by_type("IfcSpace")
    if not spaces:
        log.warning("В модели не найдено ни одного IfcSpace", name="space_check")
        return pd.DataFrame()

    rows: list[dict[str, Any]] = []
    for entity in spaces:
        row: dict[str, Any] = {}
        row.update(get_base_fields(entity, model, ifc_path))
        row.update(add_rus_properties(entity))
        row.update(add_geometry_and_area(entity, model, ifc_path, tolerance))
        rows.append(row)

    df = pd.DataFrame(rows)

    # Статистика
    match_count = df[df["Совпадение площадей"] == "ДА"].shape[0]
    mismatch_count = df[df["Совпадение площадей"] == "НЕТ"].shape[0]
    missing_geom = df[df["Есть геометрия"] == "НЕТ"].shape[0]
    log.info(
        f"Проверено {len(df)} IfcSpace: совпало={match_count}, не совпало={mismatch_count}, без геометрии={missing_geom}",
        name="space_check",
    )

    return df


def get_problem_spaces_dataframe(
    spaces_df: pd.DataFrame,
) -> pd.DataFrame:
    """Возвращает DataFrame с помещениями/зонами, имеющими проблемы.

    :param spaces_df: Общий DataFrame помещений, возвращённый get_spaces_dataframe.
    :returns: DataFrame с проблемными помещениями.
    """

    if spaces_df.empty:
        return pd.DataFrame()

    problem_mask = (
        (spaces_df["Есть геометрия"] == "НЕТ")
        | (spaces_df["Совпадение площадей"] == "НЕТ")
    )
    problems = spaces_df[problem_mask]

    columns = [
        "GUID",
        "Name",
        "LongName",
        "Revit ID",
        "Level",
        "Model",
        "Расхождение (%)",
        "Есть геометрия",
        "Комментарий",
    ]
    available_columns = [c for c in columns if c in problems.columns]
    result = problems[available_columns].copy()

    log.info(
        f"Проблемных помещений/зон: {len(result)} (из {len(spaces_df)} total)",
        name="space_check",
    )
    return result


if __name__ == "__main__":
    IFC_PATH = Path(r"samples\НН_К00_С00_АР_АГР.ifc")
    OUTPUT_DIR = Path(r".output")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    log.configure(str(OUTPUT_DIR / "space_check.log"))
    log.init("Проверка геометрии IfcSpace")

    model = ifcopenshell.open(str(IFC_PATH))
    spaces_df = get_spaces_dataframe(model, ifc_path=IFC_PATH, tolerance=0.1)
    problems_df = get_problem_spaces_dataframe(spaces_df)

    log.info(f"Общий DataFrame: {len(spaces_df)} помещений/зон", name="space_check")
    log.info(f"Проблемный DataFrame: {len(problems_df)} записей", name="space_check")

    log.info("Проверка геометрии завершена.", name="space_check")
