import math
import random
import tempfile
from typing import List, Tuple

from presentation import Presentation


def gaussian2d(mean_x: float, mean_y: float, cov_x: float, cov_y: float, cov_xy: float) -> Tuple[float, float]:
    std_x = math.sqrt(cov_x)
    std_y = math.sqrt(cov_y)

    rho = cov_xy / (std_x * std_y)

    z1 = random.gauss(0, 1)
    z2 = random.gauss(0, 1)

    x = mean_x + std_x * z1
    y = mean_y + std_y * (rho * z1 + math.sqrt(1 - rho ** 2) * z2)
    return x, y


def spiral(label: int, h: float = 1.25, delta: float = 0.1) -> Tuple[float, float]:
    angle = random.random()
    r = angle - delta + random.random() * delta * 2
    t = h * angle * 2 * math.pi

    if label == 1:
        t += math.pi

    return r * math.sin(t), r * math.cos(t)


def view_points(points: List[Tuple[float, float]], limits: dict, x0: float, y0: float, width: float, height: float) -> List[Tuple[float, float]]:
    mapped_points = []

    for x, y in points:
        x = x0 + (x - limits["x_min"]) / (limits["x_max"] - limits["x_min"]) * width
        y = y0 + (limits["y_max"] - y) / (limits["y_max"] - limits["y_min"]) * height
        mapped_points.append((x, y))

    return mapped_points


def main() -> None:
    x0 = 2
    y0 = 2
    width = 14
    height = 14

    limits1 = {"x_min": -1.7, "y_min": -1.9, "x_max": 1.6, "y_max": 1.4}
    gaussians1 = [gaussian2d(mean_x=-0.5, mean_y=-0.4, cov_x=0.1, cov_y=0.1, cov_xy=0.05) for _ in range(400)]
    gaussians2 = [gaussian2d(mean_x=0.4, mean_y=0, cov_x=0.08, cov_y=0.15, cov_xy=-0.07) for _ in range(400)]
    gaussian_points1 = view_points(points=gaussians1, limits=limits1, x0=x0, y0=y0, width=width, height=height)
    gaussian_points2 = view_points(points=gaussians2, limits=limits1, x0=x0, y0=y0, width=width, height=height)

    limits2 = {"x_min": -1.2, "y_min": -1.2, "x_max": 1.2, "y_max": 1.2}
    spiral1 = view_points(points=[spiral(1) for _ in range(400)], limits=limits2, x0=x0 + width + 2, y0=y0, width=width, height=height)
    spiral2 = view_points(points=[spiral(-1) for _ in range(400)], limits=limits2, x0=x0 + width + 2, y0=y0, width=width, height=height)

    with tempfile.TemporaryDirectory() as temp_path:
        presentation = Presentation(presentation_path="../empty.pptx", work_path=temp_path)

        presentation.add_ellipses(ellipses=[{"x": x, "y": y, "d": 0.2, "fill": "#dd7373", "stroke": "#fff", "thickness": 0.5} for x, y in gaussian_points1])
        presentation.add_ellipses(ellipses=[{"x": x, "y": y, "d": 0.2, "fill": "#7699d4", "stroke": "#fff", "thickness": 0.5} for x, y in gaussian_points2])
        presentation.add_rectangle({"x": x0, "y": y0, "w": width, "h": height, "stroke": "#222", "radius": 0.1})

        presentation.add_ellipses(ellipses=[{"x": x, "y": y, "d": 0.2, "fill": "#dd7373", "stroke": "#fff", "thickness": 0.5} for x, y in spiral1])
        presentation.add_ellipses(ellipses=[{"x": x, "y": y, "d": 0.2, "fill": "#7699d4", "stroke": "#fff", "thickness": 0.5} for x, y in spiral2])
        presentation.add_rectangle({"x": x0 + width + 2, "y": y0, "w": width, "h": height, "stroke": "#222", "radius": 0.1})

        presentation.save("scatter.pptx")


if __name__ == "__main__":
    main()
