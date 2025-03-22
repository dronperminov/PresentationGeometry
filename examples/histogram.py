import tempfile

from presentation import Presentation


def main() -> None:
    histograms = [
        {
            "values": [0, 2, 4, 9, 16, 25, 15, 10, 3, 1, 1],
            "fill": "#7699d4",
            "stroke": "#fff"
        },
        {
            "values": [1, 2, 1, 2, 1, 3, 1, 4, 5, 2, 3, 1],
            "fill": "#dd7373",
            "stroke": "#fff"
        },
        {
            "values": [200, 100, 50, 25, 12, 6, 3],
            "fill": "#89dd73",
            "stroke": "#fff"
        }
    ]

    x0 = 2
    y0 = 1.25
    bar_width = 1.8
    height = 5

    with tempfile.TemporaryDirectory() as temp_path:
        presentation = Presentation(presentation_path="../empty.pptx", work_path=temp_path)

        for histogram in histograms:
            values = histogram["values"]
            max_value = max(values)
            rectangles = []

            for i, value in enumerate(values):
                bar_height = value / max_value * height
                x = x0 + i * bar_width
                y = y0 + height - bar_height
                rectangles.append({"x": x, "y": y, "w": bar_width, "h": bar_height, "fill": histogram["fill"], "stroke": histogram["stroke"]})

            presentation.add_rectangles(rectangles=rectangles)
            presentation.add_rectangle({"x": x0 - 0.5, "y": y0 - 0.5, "w": bar_width * len(values) + 1, "h": height + 0.5, "stroke": "#222", "radius": 0.07})
            y0 += height + 1

        presentation.save("histogram.pptx")


if __name__ == "__main__":
    main()
