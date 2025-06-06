import tempfile

from presentation import Presentation


def main() -> None:
    with tempfile.TemporaryDirectory() as temp_path:
        presentation = Presentation(presentation_path="../empty.pptx", work_path=temp_path)

        presentation.add_shapes(shapes=[
            {"shape": "line", "x1": 2.85, "y1": 7.27, "x2": 6.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.63},
            {"shape": "line", "x1": 2.85, "y1": 7.27, "x2": 8.1, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.48},
            {"shape": "line", "x1": 2.85, "y1": 7.27, "x2": 9.6, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.51},
            {"shape": "line", "x1": 4.35, "y1": 7.27, "x2": 6.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.7},
            {"shape": "line", "x1": 4.35, "y1": 7.27, "x2": 8.1, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.31},
            {"shape": "line", "x1": 4.35, "y1": 7.27, "x2": 9.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.55},
            {"shape": "line", "x1": 5.85, "y1": 7.27, "x2": 6.6, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.07},
            {"shape": "line", "x1": 5.85, "y1": 7.27, "x2": 8.1, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.68},
            {"shape": "line", "x1": 5.85, "y1": 7.27, "x2": 9.6, "y2": 2.6, "stroke": "#7699d4", "thickness": 2.1},
            {"shape": "line", "x1": 7.35, "y1": 7.27, "x2": 6.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.52},
            {"shape": "line", "x1": 7.35, "y1": 7.27, "x2": 8.1, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.12},
            {"shape": "line", "x1": 7.35, "y1": 7.27, "x2": 9.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.55},
            {"shape": "line", "x1": 8.85, "y1": 7.27, "x2": 6.6, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.04},
            {"shape": "line", "x1": 8.85, "y1": 7.27, "x2": 8.1, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.66},
            {"shape": "line", "x1": 8.85, "y1": 7.27, "x2": 9.6, "y2": 2.6, "stroke": "#7699d4", "thickness": 1.08},
            {"shape": "line", "x1": 10.35, "y1": 7.27, "x2": 6.6, "y2": 2.6, "stroke": "#000000", "thickness": 0.03},
            {"shape": "line", "x1": 10.35, "y1": 7.27, "x2": 8.1, "y2": 2.6, "stroke": "#7699d4", "thickness": 0.07},
            {"shape": "line", "x1": 10.35, "y1": 7.27, "x2": 9.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.23},
            {"shape": "line", "x1": 11.85, "y1": 7.27, "x2": 6.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.6},
            {"shape": "line", "x1": 11.85, "y1": 7.27, "x2": 8.1, "y2": 2.6, "stroke": "#dd7373", "thickness": 0.64},
            {"shape": "line", "x1": 11.85, "y1": 7.27, "x2": 9.6, "y2": 2.6, "stroke": "#dd7373", "thickness": 3.31},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 2.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.15},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 4.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.3},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 5.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.12},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 7.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.06},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 8.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.11},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 10.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.18},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 11.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.28},
            {"shape": "line", "x1": 2.85, "y1": 11.93, "x2": 13.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.89},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 2.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.11},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 4.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.2},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 5.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.09},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 7.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.28},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 8.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.16},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 10.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.09},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 11.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.24},
            {"shape": "line", "x1": 4.35, "y1": 11.93, "x2": 13.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 1.13},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 2.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.32},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 4.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.09},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 5.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.16},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 7.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.32},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 8.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.09},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 10.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.14},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 11.85, "y2": 7.27, "stroke": "#000000", "thickness": 0.03},
            {"shape": "line", "x1": 5.85, "y1": 11.93, "x2": 13.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.68},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 2.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.22},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 4.35, "y2": 7.27, "stroke": "#000000", "thickness": 0.03},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 5.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.09},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 7.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.23},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 8.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.04},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 10.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.03},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 11.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.19},
            {"shape": "line", "x1": 7.35, "y1": 11.93, "x2": 13.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 1.58},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 2.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.09},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 4.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.19},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 5.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.26},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 7.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.37},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 8.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.3},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 10.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.22},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 11.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.3},
            {"shape": "line", "x1": 8.85, "y1": 11.93, "x2": 13.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.41},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 2.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.26},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 4.35, "y2": 7.27, "stroke": "#000000", "thickness": 0.03},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 5.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.13},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 7.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.34},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 8.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.25},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 10.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.38},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 11.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.07},
            {"shape": "line", "x1": 10.35, "y1": 11.93, "x2": 13.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 1.58},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 2.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.3},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 4.35, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.33},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 5.85, "y2": 7.27, "stroke": "#7699d4", "thickness": 0.29},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 7.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.15},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 8.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.03},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 10.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.07},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 11.85, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.3},
            {"shape": "line", "x1": 11.85, "y1": 11.93, "x2": 13.35, "y2": 7.27, "stroke": "#dd7373", "thickness": 0.42},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 2.85, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.09},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 4.35, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.04},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 5.85, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.11},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 7.35, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.3},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 8.85, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.18},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 10.35, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.37},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 11.85, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.12},
            {"shape": "line", "x1": 6.6, "y1": 16.6, "x2": 13.35, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.78},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 2.85, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.26},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 4.35, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.06},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 5.85, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.18},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 7.35, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.22},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 8.85, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.2},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 10.35, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.26},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 11.85, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.06},
            {"shape": "line", "x1": 8.1, "y1": 16.6, "x2": 13.35, "y2": 11.93, "stroke": "#dd7373", "thickness": 1.6},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 2.85, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.25},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 4.35, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.13},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 5.85, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.05},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 7.35, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.2},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 8.85, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.05},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 10.35, "y2": 11.93, "stroke": "#7699d4", "thickness": 0.13},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 11.85, "y2": 11.93, "stroke": "#dd7373", "thickness": 0.23},
            {"shape": "line", "x1": 9.6, "y1": 16.6, "x2": 13.35, "y2": 11.93, "stroke": "#dd7373", "thickness": 1.84},
            {"shape": "ellipse", "x": 6.1, "y": 2.1, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 6.1, "y": 2.1, "w": 1.0, "h": 1.0, "text": "x₁", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 7.6, "y": 2.1, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 7.6, "y": 2.1, "w": 1.0, "h": 1.0, "text": "x₂", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 9.1, "y": 2.1, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 9.1, "y": 2.1, "w": 1.0, "h": 1.0, "text": "1", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 2.35, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 2.35, "y": 6.77, "w": 1.0, "h": 1.0, "text": "A₁", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 3.85, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 3.85, "y": 6.77, "w": 1.0, "h": 1.0, "text": "A₂", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 5.35, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 5.35, "y": 6.77, "w": 1.0, "h": 1.0, "text": "A₃", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 6.85, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 6.85, "y": 6.77, "w": 1.0, "h": 1.0, "text": "A₄", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 8.35, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 8.35, "y": 6.77, "w": 1.0, "h": 1.0, "text": "A₅", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 9.85, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 9.85, "y": 6.77, "w": 1.0, "h": 1.0, "text": "A₆", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 11.35, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 11.35, "y": 6.77, "w": 1.0, "h": 1.0, "text": "A₇", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 12.85, "y": 6.77, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 12.85, "y": 6.77, "w": 1.0, "h": 1.0, "text": "1", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 2.35, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 2.35, "y": 11.43, "w": 1.0, "h": 1.0, "text": "B₁", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 3.85, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 3.85, "y": 11.43, "w": 1.0, "h": 1.0, "text": "B₂", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 5.35, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 5.35, "y": 11.43, "w": 1.0, "h": 1.0, "text": "B₃", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 6.85, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 6.85, "y": 11.43, "w": 1.0, "h": 1.0, "text": "B₄", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 8.35, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 8.35, "y": 11.43, "w": 1.0, "h": 1.0, "text": "B₅", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 9.85, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 9.85, "y": 11.43, "w": 1.0, "h": 1.0, "text": "B₆", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 11.35, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 11.35, "y": 11.43, "w": 1.0, "h": 1.0, "text": "B₇", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 12.85, "y": 11.43, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 12.85, "y": 11.43, "w": 1.0, "h": 1.0, "text": "1", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 6.1, "y": 16.1, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 6.1, "y": 16.1, "w": 1.0, "h": 1.0, "text": "C₁", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 7.6, "y": 16.1, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 7.6, "y": 16.1, "w": 1.0, "h": 1.0, "text": "C₂", "size": 14, "align": "center"},
            {"shape": "ellipse", "x": 9.1, "y": 16.1, "d": 1.0, "fill": "#ffffff", "stroke": "#000000", "thickness": 0.25},
            {"shape": "textbox", "x": 9.1, "y": 16.1, "w": 1.0, "h": 1.0, "text": "C₃", "size": 14, "align": "center"}
        ])

        presentation.add_textbox({
            "x": 16.5, "y": 2,
            "w": 16, "h": 2,
            "text": "Многослойный персептрон",
            "size": 36,
            "align": "center",
            "bold": True
        })

        presentation.add_textbox({
            "x": 16.5, "y": 4.5,
            "w": 16, "h": 4,
            "text": "- 2 входа\n- 2 скрытых слоя по 7 нейронов\n- выходной слой на 3 нейрона",
            "size": 24,
            "margin": {},
            "align": "left",
            "vertical-align": "top"
        })
        presentation.save("neural_network.pptx")


if __name__ == "__main__":
    main()
