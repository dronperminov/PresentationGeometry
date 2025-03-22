import tempfile

from presentation import Presentation


def main() -> None:
    with tempfile.TemporaryDirectory() as temp_path:
        presentation = Presentation(presentation_path="empty.pptx", work_path=temp_path)

        presentation.add_ellipse({"x": 20, "y": 2, "d": 4, "fill": "#7699d4"})

        presentation.add_rectangle({"x": 18, "y": 8, "w": 4, "h": 8.5, "radius": 0.25, "fill": "#dd7373", "stroke": "#222", "thickness": 3, "rotate": 30})
        presentation.add_rectangle({"x": 24, "y": 12, "w": 3, "h": 3, "radius": 0, "fill": "#dd7373", "stroke": "#222", "thickness": 1})

        presentation.add_lines([
            {"x1": 1, "y1": 1, "x2": 10, "y2": 1, "thickness": 2, "stroke": "#7699d4"},
            {"x1": 1, "y1": 1, "x2": 1, "y2": 5, "thickness": 2, "stroke": "#dd7373"},
            {"x1": 10, "y1": 1, "x2": 1, "y2": 5, "thickness": 2, "stroke": "#000"}
        ])

        presentation.add_ellipses([
            {"x": 4.5, "y": 5, "dx": 2, "dy": 3.5, "fill": "#dd7373", "stroke": "#000000", "thickness": 2, "fill-opacity": 0.5, "stroke-opacity": 0.75},
            {"x": 3, "y": 7.5, "dx": 3.5, "dy": 2, "fill": "#dd7373", "stroke": "#000000", "opacity": 0.25, "rotate": -45},
            {"x": 5, "y": 7.5, "dx": 3.5, "dy": 2, "fill": "#dd7373", "stroke": "#000000", "opacity": 0.85, "rotate": 45}
        ])

        presentation.add_polygons([
            {
                "points": [{"x": 10, "y": 10}, {"x": 12, "y": 12}, {"x": 10, "y": 14}, {"x": 8, "y": 12}, {"x": 10, "y": 10}],
                "stroke": "FF00FF",
                "fill": "888888",
                "thickness": 2.5
            },
            {
                "points": [{"x": 13, "y": 5}, {"x": 14, "y": 6}, {"x": 13, "y": 7}, {"x": 10, "y": 7}, {"x": 9, "y": 6}, {"x": 10, "y": 5}],
                "stroke": "FFFFFF",
                "fill": "88FF88",
                "thickness": 0.5,
                "rotate": 21
            },
        ])

        presentation.add_rectangles([
            {"x": 1, "y": 16, "w": 1.2, "h": 2.7, "radius": 0.2, "fill": "#7699d4", "stroke": "#fff"},
            {"x": 2.2, "y": 16.4, "w": 1.2, "h": 2.3, "radius": 0.2, "fill": "#7699d4", "stroke": "#fff"},
            {"x": 3.4, "y": 17, "w": 1.2, "h": 1.7, "radius": 0.2, "fill": "#7699d4", "stroke": "#fff"},
            {"x": 4.6, "y": 16.1, "w": 1.2, "h": 2.6, "radius": 0.2, "fill": "#7699d4", "stroke": "#fff"}
        ])

        presentation.save("examples/example.pptx")


if __name__ == "__main__":
    main()
