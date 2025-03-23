import tempfile

from presentation import Presentation


def main() -> None:
    w, h = 33.867, 19.05

    with tempfile.TemporaryDirectory() as temp_path:
        presentation = Presentation(presentation_path="../empty.pptx", work_path=temp_path)

        presentation.add_shapes([
            {"shape": "line", "x1": w / 2, "y1": 7, "x2": 0, "y2": 0, "stroke": "#222", "thickness": 2},
            {"shape": "line", "x1": w / 2, "y1": 7, "x2": w, "y2": 0, "stroke": "#222", "thickness": 2},
            {"shape": "ellipse", "x": w / 2 - 2, "y": 5, "d": 4, "fill": "#7699d4"},
            {"shape": "rectangle", "x": w / 2 - 2, "y": 10, "w": 4, "h": 6, "radius": 0.2, "fill": "#7699d4"},
            {"shape": "polygon", "points": [{"x": 0, "y": 2}, {"x": 3, "y": 5}, {"x": 3, "y": h - 5}, {"x": 0, "y": h - 2}], "fill": "dd7373"},
            {"shape": "polygon", "points": [{"x": w, "y": 2}, {"x": w - 3, "y": 5}, {"x": w - 3, "y": h - 5}, {"x": w, "y": h - 2}], "fill": "dd7373"}
        ])

        presentation.save("shapes.pptx")


if __name__ == "__main__":
    main()
