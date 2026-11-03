"""Совмещение геометрии IFC с видами из TRM-контейнера.

Модуль извлекает из TRM архива PDF и координатные параметры видов из
`export.log`. Геометрия IFC пересекается с плоскостью выбранного вида и
накладывается на соответствующую страницу PDF.
"""

from __future__ import annotations

import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

import ifcopenshell
import ifcopenshell.geom
import pymupdf
from ifcopenshell.util.placement import get_local_placement

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.utils.logger import Logger as log


@dataclass(frozen=True)
class TrmView:
    """Параметры проекции одного вида из журнала экспорта TRM."""

    name: str
    view_type: str
    scale_denominator: float
    origin: tuple[float, float, float]
    level_elevation: float | None
    right: tuple[float, float, float]
    up: tuple[float, float, float]
    view_direction: tuple[float, float, float]
    projection_center: tuple[float, float, float]
    page_number: int


@dataclass(frozen=True)
class IfcGeometry:
    """Подготовленная геометрия IFC без внутренних рёбер триангуляции."""

    triangles: list[tuple[tuple[float, float, float], ...]]
    feature_edges: list[tuple[tuple[float, float, float], tuple[float, float, float]]]


def _parse_vector(value: str) -> tuple[float, float, float]:
    """Разбирает тройку координат с точкой или запятой в качестве разделителя."""

    numbers = re.findall(r"[-+]?\d+(?:[,.]\d+)?", value)
    if len(numbers) != 3:
        raise ValueError(f"Ожидалась тройка координат, получено: {value}")
    return tuple(float(number.replace(",", ".")) for number in numbers)  # type: ignore[return-value]


def _feet_to_meters(vector: tuple[float, float, float]) -> tuple[float, float, float]:
    """Переводит координаты из внутренних футов Revit в метры IFC."""

    return tuple(value * 0.3048 for value in vector)  # type: ignore[return-value]


def _parse_scale(value: str) -> float:
    """Возвращает знаменатель масштаба из строки вида `1:100`."""

    match = re.search(r"1\s*:\s*([\d,.]+)", value)
    if not match:
        raise ValueError(f"Не удалось разобрать масштаб: {value}")
    return float(match.group(1).replace(",", "."))


def _extract_pdf_member(archive: zipfile.ZipFile) -> str:
    """Находит PDF-файл в TRM-контейнере по сигнатуре `%PDF`."""

    for member in archive.infolist():
        with archive.open(member) as stream:
            if stream.read(5) == b"%PDF-":
                return member.filename
    raise ValueError("В TRM-контейнере не найден PDF-файл")


def parse_trm_views(trm_path: Path) -> tuple[str, list[TrmView]]:
    """Извлекает имя PDF и параметры видов из TRM-архива.

    :param trm_path: Путь к ZIP-контейнеру TRM.
    :returns: Кортеж из имени PDF внутри архива и списка видов в порядке страниц.
    :raises ValueError: Если в журнале отсутствуют координаты или PDF.
    """

    with zipfile.ZipFile(trm_path) as archive:
        pdf_member = _extract_pdf_member(archive)
        log_member = "export.log"
        try:
            log_text = archive.read(log_member).decode("utf-8-sig")
        except KeyError as error:
            raise ValueError("В TRM-контейнере отсутствует export.log") from error

    page_names: list[str] = []
    for name in re.findall(r"EXPORT_2D «([^»]+)»:", log_text):
        if name not in page_names:
            page_names.append(name)
    views: list[TrmView] = []
    axis_pattern = re.compile(
        r"\[COORD\] «(?P<name>[^»]+)» sheet→model axes: "
        r"Origin=(?P<origin>\([^)]*\)), Right=(?P<right>\([^)]*\)), "
        r"Up=(?P<up>\([^)]*\)), ViewDir=(?P<direction>\([^)]*\))"
    )
    type_pattern = re.compile(
        r"\[COORD\] «(?P<name>[^»]+)» sheet→model: type=(?P<type>[^,]+), "
        r"scale=(?P<scale>[^,\s]+)"
    )
    center_pattern = re.compile(
        r"\[COORD\] «(?P<name>[^»]+)» sheet→model points: .*?afterView=(?P<center>\([^)]*\))"
    )
    level_pattern = re.compile(
        r"\[COORD\] «(?P<name>[^»]+)» Level «[^»]+»: Elevation=(?P<elevation>[-+]?\d+(?:[,.]\d+)?) ft"
    )
    view_types = {match.group("name"): (match.group("type").strip(), _parse_scale(match.group("scale"))) for match in type_pattern.finditer(log_text)}
    projection_centers = {
        match.group("name"): _feet_to_meters(_parse_vector(match.group("center")))
        for match in center_pattern.finditer(log_text)
    }
    level_elevations = {
        match.group("name"): float(match.group("elevation").replace(",", ".")) * 0.3048
        for match in level_pattern.finditer(log_text)
    }

    for match in axis_pattern.finditer(log_text):
        name = match.group("name")
        view_type, scale = view_types.get(name, ("Unknown", 100.0))
        if name not in page_names:
            log.warning(f"Вид {name} найден в координатах, но отсутствует среди страниц PDF", name="trm")
            continue
        views.append(
            TrmView(
                name=name,
                view_type=view_type,
                scale_denominator=scale,
                origin=_feet_to_meters(_parse_vector(match.group("origin"))),
                level_elevation=level_elevations.get(name),
                right=_parse_vector(match.group("right")),
                up=_parse_vector(match.group("up")),
                view_direction=_parse_vector(match.group("direction")),
                projection_center=projection_centers.get(name, _feet_to_meters(_parse_vector(match.group("origin")))),
                page_number=page_names.index(name),
            )
        )

    unique_views = {view.name: view for view in views}
    result = sorted(unique_views.values(), key=lambda view: view.page_number)
    log.info(f"Из TRM извлечено видов: {len(result)}, PDF: {pdf_member}", name="trm")
    return pdf_member, result


def _dot(left: tuple[float, float, float], right: tuple[float, float, float]) -> float:
    """Вычисляет скалярное произведение трёхмерных векторов."""

    return sum(a * b for a, b in zip(left, right))


def _subtract(left: tuple[float, float, float], right: tuple[float, float, float]) -> tuple[float, float, float]:
    """Вычитает один трёхмерный вектор из другого."""

    return tuple(a - b for a, b in zip(left, right))  # type: ignore[return-value]


def _cross(left: tuple[float, float, float], right: tuple[float, float, float]) -> tuple[float, float, float]:
    """Вычисляет векторное произведение трёхмерных векторов."""

    return (
        left[1] * right[2] - left[2] * right[1],
        left[2] * right[0] - left[0] * right[2],
        left[0] * right[1] - left[1] * right[0],
    )


def _normal(triangle: tuple[tuple[float, float, float], ...]) -> tuple[float, float, float]:
    """Возвращает нормализованную нормаль треугольника."""

    vector = _cross(_subtract(triangle[1], triangle[0]), _subtract(triangle[2], triangle[0]))
    length = sum(value * value for value in vector) ** 0.5
    return tuple(value / length for value in vector) if length else (0.0, 0.0, 0.0)  # type: ignore[return-value]


def _to_site_local(point: tuple[float, float, float], site_matrix: object) -> tuple[float, float, float]:
    """Переводит глобальную точку IFC в локальную систему `IfcSite`.

    IFC-файл примера содержит координаты размещения в миллиметрах, а
    геометрический API IfcOpenShell возвращает мировые координаты в метрах.
    Поэтому перенос из матрицы размещения дополнительно масштабируется.
    """

    matrix = site_matrix
    origin_x = float(matrix[0][3]) / 1000.0  # type: ignore[index]
    origin_y = float(matrix[1][3]) / 1000.0  # type: ignore[index]
    origin_z = float(matrix[2][3]) / 1000.0  # type: ignore[index]
    delta_x = point[0] - origin_x
    delta_y = point[1] - origin_y
    return (
        float(matrix[0][0]) * delta_x + float(matrix[1][0]) * delta_y,
        float(matrix[0][1]) * delta_x + float(matrix[1][1]) * delta_y,
        point[2] - origin_z,
    )


def _intersection_points(
    triangle: tuple[tuple[float, float, float], ...],
    view: TrmView,
    tolerance: float = 0.001,
) -> list[tuple[float, float, float]]:
    """Возвращает точки пересечения треугольника с плоскостью вида."""

    distances = [_dot(_subtract(point, view.origin), view.view_direction) for point in triangle]
    points: list[tuple[float, float, float]] = []
    for index in range(3):
        first = triangle[index]
        second = triangle[(index + 1) % 3]
        first_distance = distances[index]
        second_distance = distances[(index + 1) % 3]
        if abs(first_distance) <= tolerance:
            points.append(first)
        if first_distance * second_distance < -tolerance**2:
            ratio = first_distance / (first_distance - second_distance)
            points.append(tuple(a + ratio * (b - a) for a, b in zip(first, second)))  # type: ignore[return-value]
    unique: dict[tuple[float, float, float], tuple[float, float, float]] = {}
    for point in points:
        unique[tuple(round(value, 5) for value in point)] = point
    return list(unique.values())


def extract_ifc_triangles(model: ifcopenshell.file) -> IfcGeometry:
    """Извлекает геометрию IFC в локальной системе координат TRM.

    :param model: Открытая IFC-модель.
    :returns: Треугольники и реальные рёбра поверхностей в метрах.
    """

    settings = ifcopenshell.geom.settings()
    settings.set(settings.USE_WORLD_COORDS, True)
    sites = model.by_type("IfcSite")
    if not sites or sites[0].ObjectPlacement is None:
        raise ValueError("В IFC отсутствует IfcSite.ObjectPlacement для привязки координат TRM")
    site_matrix = get_local_placement(sites[0].ObjectPlacement)
    elements = [element for element in model.by_type("IfcElement") if not element.is_a("IfcOpeningElement")]
    log.info(f"Геометрия IFC для TRM: элементов после исключения IfcOpeningElement: {len(elements)}", name="trm")
    iterator = ifcopenshell.geom.iterator(settings, model, include=elements)
    if not iterator.initialize():
        return []

    triangles: list[tuple[tuple[float, float, float], ...]] = []
    feature_edges: list[tuple[tuple[float, float, float], tuple[float, float, float]]] = []
    while True:
        shape = iterator.get()
        geometry = shape.geometry if shape is not None else None
        if geometry is not None:
            vertices = geometry.verts
            points = [_to_site_local(tuple(vertices[index : index + 3]), site_matrix) for index in range(0, len(vertices), 3)]
            faces = geometry.faces
            element_triangles: list[tuple[tuple[float, float, float], ...]] = []
            for index in range(0, len(faces), 3):
                element_triangles.append(tuple(points[faces[index + offset]] for offset in range(3)))
            triangles.extend(element_triangles)

            edge_faces: dict[tuple[int, int], list[int]] = {}
            for triangle_index, face_start in enumerate(range(0, len(faces), 3)):
                face_indices = [faces[face_start + offset] for offset in range(3)]
                for edge_start, edge_end in zip(face_indices, face_indices[1:] + face_indices[:1]):
                    edge_faces.setdefault(tuple(sorted((edge_start, edge_end))), []).append(triangle_index)
            normals = [_normal(triangle) for triangle in element_triangles]
            for (edge_start, edge_end), adjacent_faces in edge_faces.items():
                is_feature = len(adjacent_faces) == 1
                if len(adjacent_faces) > 1:
                    first_normal = normals[adjacent_faces[0]]
                    second_normal = normals[adjacent_faces[1]]
                    is_feature = abs(_dot(first_normal, second_normal)) < 0.995
                if is_feature:
                    feature_edges.append((points[edge_start], points[edge_end]))
        if not iterator.next():
            break
    log.info(
        f"Из IFC извлечено треугольников: {len(triangles)}, "
        f"граничных/ломаных рёбер: {len(feature_edges)}",
        name="trm",
    )
    return IfcGeometry(triangles=triangles, feature_edges=feature_edges)


def extract_ifc_section_segments(
    model: ifcopenshell.file,
    view: TrmView,
    geometry: IfcGeometry | None = None,
) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    """Строит двумерные отрезки пересечения геометрии IFC с плоскостью вида.

    :param model: Открытая IFC-модель.
    :param view: Параметры вида TRM.
    :param geometry: Подготовленная геометрия; если не передана, строится из IFC.
    :returns: Список отрезков в координатах `Right`/`Up` вида.
    """

    geometry = geometry if geometry is not None else extract_ifc_triangles(model)
    segments: list[tuple[tuple[float, float], tuple[float, float]]] = []
    for triangle in geometry.triangles:
        intersections = _intersection_points(triangle, view)
        if len(intersections) == 2:
            projected = [
                (_dot(_subtract(point, view.origin), view.right), _dot(_subtract(point, view.origin), view.up))
                for point in intersections
            ]
            segments.append((projected[0], projected[1]))
    segments = _unique_segments(segments)
    if segments:
        coordinates = [coordinate for segment in segments for point in segment for coordinate in point]
        log.info(
            f"Для вида {view.name} построено отрезков IFC: {len(segments)}, "
            f"bbox проекции: x=({min(coordinates[0::2]):.3f},{max(coordinates[0::2]):.3f}), "
            f"y=({min(coordinates[1::2]):.3f},{max(coordinates[1::2]):.3f}) м",
            name="trm",
        )
    else:
        log.info(f"Для вида {view.name} построено отрезков IFC: 0", name="trm")
    return segments


def _projected_triangle_edges(
    triangle: tuple[tuple[float, float, float], ...],
    view: TrmView,
    depth_min: float,
    depth_max: float,
) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    """Проецирует рёбра треугольника на плоскость вида в заданном диапазоне глубины."""

    distances = [_dot(_subtract(point, view.origin), view.view_direction) for point in triangle]
    if max(distances) < depth_min or min(distances) > depth_max:
        return []
    projected = [
        (_dot(_subtract(point, view.origin), view.right), _dot(_subtract(point, view.origin), view.up))
        for point in triangle
    ]
    return [(projected[index], projected[(index + 1) % 3]) for index in range(3)]


def _unique_segments(
    segments: list[tuple[tuple[float, float], tuple[float, float]]],
) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    """Удаляет повторяющиеся отрезки после проецирования треугольников."""

    unique: dict[tuple[tuple[float, float], tuple[float, float]], tuple[tuple[float, float], tuple[float, float]]] = {}
    for first, second in segments:
        first_key = (round(first[0], 4), round(first[1], 4))
        second_key = (round(second[0], 4), round(second[1], 4))
        key = tuple(sorted((first_key, second_key)))  # type: ignore[assignment]
        unique[key] = (first, second)
    return list(unique.values())


def _floor_plan_depth_range(view: TrmView) -> tuple[float, float]:
    """Возвращает диапазон глубины проекции планового вида относительно `view.origin`.

    Для планов TRM верхняя граница принимается на 1,2 м выше отметки уровня,
    а нижняя граница — на 0,5 м ниже отметки уровня. Если отметка уровня
    отсутствует в `export.log`, используется `Origin.Z` как резервная основа.

    :param view: Параметры планового вида TRM.
    :returns: Минимальная и максимальная глубина вдоль направления взгляда в метрах.
    """

    level_elevation = view.level_elevation if view.level_elevation is not None else view.origin[2]
    level_offset = _dot(
        (0.0, 0.0, level_elevation - view.origin[2]),
        view.view_direction,
    )
    depth_min = level_offset - 0.5
    depth_max = level_offset + 1.2
    log.debug(
        f"Диапазон планового вида {view.name}: уровень={level_elevation:.3f} м, "
        f"относительно Origin=({depth_min:.3f},{depth_max:.3f}) м",
        name="trm",
    )
    return depth_min, depth_max


def extract_ifc_view_segments(
    model: ifcopenshell.file,
    view: TrmView,
    geometry: IfcGeometry | None = None,
) -> list[tuple[tuple[float, float], tuple[float, float]]]:
    """Формирует линии IFC с учётом типа TRM-вида.

    Для планов используется диапазон от 0,5 м ниже отметки уровня до 1,2 м
    выше неё. Для фасадов используется ортогональная проекция всей модели.
    Для разрезов сохраняется точное пересечение плоскостей и проекция на 0,5 м
    в глубину вида.

    :param model: Открытая IFC-модель.
    :param view: Параметры вида TRM.
    :param geometry: Подготовленная геометрия IFC.
    :returns: Линии проекции в координатах `Right`/`Up`.
    """

    geometry = geometry if geometry is not None else extract_ifc_triangles(model)
    if view.view_type == "Section":
        segments = extract_ifc_section_segments(model, view, geometry)
        depth_min, depth_max = -0.5, 0.0
    elif view.view_type == "Elevation":
        segments = []
        depth_min, depth_max = -float("inf"), float("inf")
    elif view.view_type == "FloorPlan":
        segments = []
        depth_min, depth_max = _floor_plan_depth_range(view)
    else:
        segments = []
        depth_min, depth_max = -float("inf"), 0.0

    for edge_start, edge_end in geometry.feature_edges:
        distances = [
            _dot(_subtract(edge_start, view.origin), view.view_direction),
            _dot(_subtract(edge_end, view.origin), view.view_direction),
        ]
        if max(distances) < depth_min or min(distances) > depth_max:
            continue
        segments.append(
            (
                (_dot(_subtract(edge_start, view.origin), view.right), _dot(_subtract(edge_start, view.origin), view.up)),
                (_dot(_subtract(edge_end, view.origin), view.right), _dot(_subtract(edge_end, view.origin), view.up)),
            )
        )
    segments = _unique_segments(segments)
    log.info(
        f"Для вида {view.name} построено проекционных отрезков IFC: {len(segments)}, "
        f"диапазон глубины=({depth_min},{depth_max}) м",
        name="trm",
    )
    return segments


def _draw_segments_on_page(page: pymupdf.Page, view: TrmView, segments: list[tuple[tuple[float, float], tuple[float, float]]]) -> None:
    """Накладывает отрезки проекции IFC на страницу PDF."""

    points_per_meter = 1000.0 / view.scale_denominator * 72.0 / 25.4
    center = pymupdf.Point(page.rect.width / 2, page.rect.height / 2)
    projection_center = (
        _dot(_subtract(view.projection_center, view.origin), view.right),
        _dot(_subtract(view.projection_center, view.origin), view.up),
    )
    for offset in range(0, len(segments), 200):
        shape = page.new_shape()
        for first, second in segments[offset : offset + 200]:
            start = pymupdf.Point(
                center.x + (first[0] - projection_center[0]) * points_per_meter,
                center.y - (first[1] - projection_center[1]) * points_per_meter,
            )
            end = pymupdf.Point(
                center.x + (second[0] - projection_center[0]) * points_per_meter,
                center.y - (second[1] - projection_center[1]) * points_per_meter,
            )
            shape.draw_line(start, end)
        shape.finish(color=(1, 0, 0), width=0.8, closePath=False)
        shape.commit(overlay=True)


def _add_view_title(page: pymupdf.Page, view: TrmView) -> None:
    """Добавляет название вида в верхнюю область листа."""

    page.draw_rect(
        pymupdf.Rect(18, 12, page.rect.width - 18, 36),
        color=(1, 1, 1),
        fill=(1, 1, 1),
        overlay=True,
    )
    page.insert_text(
        pymupdf.Point(28, 29),
        f"{view.name} | {view.view_type} | М 1:{view.scale_denominator:g}",
        fontsize=11,
        fontname="Arial",
        fontfile=r"C:\Windows\Fonts\arial.ttf",
        color=(0.05, 0.15, 0.45),
        overlay=True,
    )


def render_trm_album(trm_path: Path, ifc_path: Path, output_path: Path) -> Path:
    """Формирует PDF-альбом совмещённых видов TRM и IFC.

    :param trm_path: Путь к TRM ZIP-контейнеру.
    :param ifc_path: Путь к IFC-модели.
    :param output_path: Путь к итоговому PDF-альбому.
    :returns: Путь к созданному PDF-альбому.
    """

    pdf_member, views = parse_trm_views(trm_path)
    if not views:
        raise ValueError("В TRM не найдено ни одного вида с координатами")
    model = ifcopenshell.open(str(ifc_path))
    geometry = extract_ifc_triangles(model)
    with zipfile.ZipFile(trm_path) as archive:
        pdf_data = archive.read(pdf_member)
    document = pymupdf.open(stream=pdf_data, filetype="pdf")
    for view in views:
        segments = extract_ifc_view_segments(model, view, geometry=geometry)
        page = document[view.page_number]
        _draw_segments_on_page(page, view, segments)
        _add_view_title(page, view)
        log.info(f"Лист альбома сформирован: {view.page_number + 1}/{len(views)} — {view.name}", name="trm")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path), garbage=4, deflate=True)
    document.close()
    log.info(f"TRM-альбом сохранён: {output_path}", name="trm")
    return output_path


def render_trm_overlay(trm_path: Path, ifc_path: Path, view_name: str, output_path: Path) -> Path:
    """Накладывает сечение IFC на страницу выбранного TRM-вида.

    :param trm_path: Путь к TRM ZIP-контейнеру.
    :param ifc_path: Путь к IFC-модели.
    :param view_name: Точное название вида из `export.log`.
    :param output_path: Путь к итоговому PNG-файлу.
    :returns: Путь к созданному изображению.
    """

    pdf_member, views = parse_trm_views(trm_path)
    view = next((item for item in views if item.name == view_name), None)
    if view is None:
        raise ValueError(f"Вид не найден в TRM: {view_name}")

    model = ifcopenshell.open(str(ifc_path))
    segments = extract_ifc_view_segments(model, view)
    with zipfile.ZipFile(trm_path) as archive:
        pdf_data = archive.read(pdf_member)
    document = pymupdf.open(stream=pdf_data, filetype="pdf")
    page = document[view.page_number]
    _draw_segments_on_page(page, view, segments)
    _add_view_title(page, view)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pixmap = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
    pixmap.save(str(output_path))
    document.close()
    log.info(f"TRM-совмещение сохранено: {output_path}", name="trm")
    return output_path


if __name__ == "__main__":
    TRM_PATH = Path(r"samples\НН_ОБ01_П00_АР_АГР.zip")
    IFC_PATH = Path(r"samples\НН_ОБ01_П00_АР_АГР.ifc")
    OUTPUT_PATH = Path(r".output\trm_album_overlay.pdf")
    log.configure(str(Path(".output") / "trm_overlay.log"))
    log.init("Альбом совмещения TRM и IFC")
    render_trm_album(
        trm_path=TRM_PATH,
        ifc_path=IFC_PATH,
        output_path=OUTPUT_PATH,
    )
