import tempfile
from typing import List, Tuple

from presentation import Presentation


def split_polygon(polygon: List[dict], line: Tuple[float, float, float]) -> Tuple[List[dict], List[dict]]:
    polygon1, polygon2 = [], []
    a, b, c = line

    for i, point in enumerate(polygon):
        point2 = polygon[(i + 1) % len(polygon)]

        sign1 = a * point["x"] + b * point["y"] + c
        sign2 = a * point2["x"] + b * point2["y"] + c

        if sign1 <= 0:
            polygon1.append(point)

        if sign1 >= 0:
            polygon2.append(point)

        if sign1 * sign2 >= 0:
            continue

        t = sign1 / (sign1 - sign2)
        p = {
            "x": point["x"] + t * (point2["x"] - point["x"]),
            "y": point["y"] + t * (point2["y"] - point["y"])
        }

        polygon1.append(p)
        polygon2.append(p)

    return polygon1, polygon2


def mix_colors(color1: str, color2: str) -> str:
    r1, g1, b1 = color1[1:3], color1[3:5], color1[5:]
    r2, g2, b2 = color2[1:3], color2[3:5], color2[5:]

    r = (int(r1, 16) + int(r2, 16)) // 2
    g = (int(g1, 16) + int(g2, 16)) // 2
    b = (int(b1, 16) + int(b2, 16)) // 2

    return f"#{hex(r)[2:]}{hex(g)[2:]}{hex(b)[2:]}"


def split_polygons(polygons: List[dict], line: Tuple[float, float, float]) -> List[dict]:
    new_polygons = []

    for polygon in polygons:
        polygon1, polygon2 = split_polygon(polygon=polygon["points"], line=line)

        if len(polygon1) > 2:
            new_polygons.append({"points": polygon1, "fill": mix_colors(polygon["fill"], "#dd7373")})

        if len(polygon2) > 2:
            new_polygons.append({"points": polygon2, "fill": mix_colors(polygon["fill"], "#7699d4")})

    return new_polygons


def view_points(points: List[dict], limits: dict, x0: float, y0: float, width: float, height: float) -> List[dict]:
    mapped_points = []

    for point in points:
        x = x0 + (point["x"] - limits["x_min"]) / (limits["x_max"] - limits["x_min"]) * width
        y = y0 + (limits["y_max"] - point["y"]) / (limits["y_max"] - limits["y_min"]) * height
        mapped_points.append({"x": x, "y": y})

    return mapped_points


def main() -> None:
    lines = [
        (-0.04116849135320658, -0.3206942150255752, 0.008537946165246545),
        (-0.7584591414936934, 0.10385981035299775, -0.9776024852609245),
        (-0.14724284735742352, 0.9107555466460698, 0.9672584243608895),
        (1.1442906237746802, 0.18843945985349314, -1.0550500227489596),
        (1.2710556674273104, -0.06936417996306213, 0.04558304687187343),
        (-0.20456406729434004, 0.2423156584615096, -0.1523538420922868),
        (0.35682639927686643, 1.3392776721814947, -0.9683512977783645),
        (0.2645942236152999, -0.9065969173006434, -0.5403082382599146)
    ]

    limits = {"x_min": -1.7, "y_min": -1.7, "x_max": 1.7, "y_max": 1.7}
    polygons = [
        {"points": [{"x": -1.7, "y": -1.7}, {"x": -1.7, "y": 1.7}, {"x": 1.7, "y": 1.7}, {"x": 1.7, "y": -1.7}], "fill": "#ffffff"}
    ]

    x0 = 1
    y0 = 1
    width = 7.5
    height = 7.5
    gap = 0.5

    columns = 4

    with tempfile.TemporaryDirectory() as temp_path:
        presentation = Presentation(presentation_path="../empty.pptx", work_path=temp_path)

        for i, line in enumerate(lines):
            polygons = split_polygons(polygons=polygons, line=line)
            x = x0 + (width + gap) * (i % columns)
            y = y0 + (height + gap) * (i // columns)

            draw_polygons = []
            for polygon in polygons:
                points = view_points(points=polygon["points"], limits=limits, x0=x, y0=y, width=width, height=height)
                draw_polygons.append({"points": points, "fill": polygon["fill"], "stroke": "#222", "thickness": 0.5})

            presentation.add_polygons(draw_polygons)

        presentation.save("polygons.pptx")


if __name__ == "__main__":
    main()
