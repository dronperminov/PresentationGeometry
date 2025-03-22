import math
import os
import re
import zipfile
from typing import List, Tuple
from xml.etree import ElementTree


class Presentation:
    def __init__(self, presentation_path: str, work_path: str) -> None:
        self.presentation_path = presentation_path
        self.work_path = work_path
        self.slide_path = os.path.join(self.work_path, "ppt", "slides", "slide1.xml")

        with zipfile.ZipFile(presentation_path, "r") as f:
            f.extractall(self.work_path)

        self.tree = ElementTree.parse(self.slide_path)
        self.sp_tree = self.__get_tree(root=self.tree.getroot())
        self.shape_id = 2

    def add_line(self, line: dict) -> None:
        self.sp_tree.append(self.make_line(line=line))

    def add_ellipse(self, ellipse: dict) -> None:
        self.sp_tree.append(self.make_ellipse(ellipse=ellipse))

    def add_rectangle(self, rectangle: dict) -> None:
        self.sp_tree.append(self.make_rectangle(rectangle=rectangle))

    def add_polygon(self, polygon: dict) -> None:
        self.sp_tree.append(self.make_polygon(polygon=polygon))

    def add_lines(self, lines: List[dict]) -> None:
        if not lines:
            return

        x, y, width, height = self.__get_lines_bbox(lines=lines)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for line in lines:
            group.append(self.make_line(line=line))

        self.sp_tree.append(group)

    def add_ellipses(self, ellipses: List[dict]) -> None:
        if not ellipses:
            return

        x, y, width, height = self.__get_ellipses_bbox(ellipses=ellipses)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for ellipse in ellipses:
            group.append(self.make_ellipse(ellipse=ellipse))

        self.sp_tree.append(group)

    def add_rectangles(self, rectangles: List[dict]) -> None:
        if not rectangles:
            return

        x, y, width, height = self.__get_rectangles_bbox(rectangles=rectangles)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for rectangle in rectangles:
            group.append(self.make_rectangle(rectangle=rectangle))

        self.sp_tree.append(group)

    def add_polygons(self, polygons: List[dict]) -> None:
        if not polygons:
            return

        x, y, width, height = self.__get_polygons_bbox(polygons=polygons)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for polygon in polygons:
            group.append(self.make_polygon(polygon=polygon))

        self.sp_tree.append(group)

    def make_line(self, line: dict) -> ElementTree.Element:
        line_node = ElementTree.Element("p:cxnSp")

        nvcxnsppr = ElementTree.SubElement(line_node, "p:nvCxnSpPr")
        ElementTree.SubElement(nvcxnsppr, "p:cNvPr", {"id": str(self.shape_id), "name": f"Line {self.shape_id}"})
        ElementTree.SubElement(nvcxnsppr, "p:cNvCxnSpPr")
        ElementTree.SubElement(nvcxnsppr, "p:nvPr")

        x1, y1, x2, y2 = line["x1"], line["y1"], line["x2"], line["y2"]

        sppr = ElementTree.SubElement(line_node, "p:spPr")
        xfrm = ElementTree.SubElement(sppr, "a:xfrm", {"flipH": "0" if x1 <= x2 else "1", "flipV": "0" if y1 <= y2 else "1"})
        ElementTree.SubElement(xfrm, "a:off", {"x": self.__get_coordinate(min(x1, x2)), "y": self.__get_coordinate(min(y1, y2))})
        ElementTree.SubElement(xfrm, "a:ext", {"cx": self.__get_coordinate(abs(x2 - x1)), "cy": self.__get_coordinate(abs(y2 - y1))})

        ElementTree.SubElement(ElementTree.SubElement(sppr, "a:prstGeom", {"prst": "line"}), "a:avLst")

        self.__set_stroke(sppr, config=line)
        self.shape_id += 1

        return line_node

    def make_ellipse(self, ellipse: dict) -> ElementTree.Element:
        ellipse_node = ElementTree.Element("p:sp")

        nvsppr = ElementTree.SubElement(ellipse_node, "p:nvSpPr")
        ElementTree.SubElement(nvsppr, "p:cNvPr", {"id": str(self.shape_id), "name": f"Ellipse {self.shape_id}"})
        ElementTree.SubElement(nvsppr, "p:cNvSpPr")
        ElementTree.SubElement(nvsppr, "p:nvPr")

        dx, dy = self.__get_diameter(ellipse, key="dx"), self.__get_diameter(ellipse, key="dy")

        sppr = ElementTree.SubElement(ellipse_node, "p:spPr")
        xfrm = ElementTree.SubElement(sppr, "a:xfrm", {"rot": self.__get_angle(ellipse.get("rotate", 0))})
        ElementTree.SubElement(xfrm, "a:off", {"x": self.__get_coordinate(ellipse["x"]), "y": self.__get_coordinate(ellipse["y"])})
        ElementTree.SubElement(xfrm, "a:ext", {"cx": self.__get_coordinate(dx), "cy": self.__get_coordinate(dy)})

        ElementTree.SubElement(ElementTree.SubElement(sppr, "a:prstGeom", {"prst": "ellipse"}), "a:avLst")

        self.__set_fill(sppr, config=ellipse)
        self.__set_stroke(sppr, config=ellipse)
        self.shape_id += 1

        return ellipse_node

    def make_rectangle(self, rectangle: dict) -> ElementTree.Element:
        rectangle_node = ElementTree.Element("p:sp")

        nvsppr = ElementTree.SubElement(rectangle_node, "p:nvSpPr")
        ElementTree.SubElement(nvsppr, "p:cNvPr", {"id": str(self.shape_id), "name": f"Ellipse {self.shape_id}"})
        ElementTree.SubElement(nvsppr, "p:cNvSpPr")
        ElementTree.SubElement(nvsppr, "p:nvPr")

        sppr = ElementTree.SubElement(rectangle_node, "p:spPr")
        xfrm = ElementTree.SubElement(sppr, "a:xfrm", {"rot": self.__get_angle(rectangle.get("rotate", 0))})
        ElementTree.SubElement(xfrm, "a:off", {"x": self.__get_coordinate(rectangle["x"]), "y": self.__get_coordinate(rectangle["y"])})
        ElementTree.SubElement(xfrm, "a:ext", {"cx": self.__get_coordinate(rectangle["w"]), "cy": self.__get_coordinate(rectangle["h"])})

        geom = ElementTree.SubElement(sppr, "a:prstGeom", {"prst": "roundRect"})
        avlst = ElementTree.SubElement(geom, "a:avLst")
        ElementTree.SubElement(avlst, "a:gd", {"name": "adj", "fmla": f'val {self.__get_fraction(rectangle.get("radius", 0) / 2)}'})

        self.__set_fill(sppr, config=rectangle)
        self.__set_stroke(sppr, config=rectangle)
        self.shape_id += 1

        return rectangle_node

    def make_polygon(self, polygon: dict) -> ElementTree.Element:
        x, y, width, height = self.__get_polygons_bbox(polygons=[polygon])

        polygon_node = ElementTree.Element("p:sp")

        nvsppr = ElementTree.SubElement(polygon_node, "p:nvSpPr")
        ElementTree.SubElement(nvsppr, "p:cNvPr", {"id": str(self.shape_id), "name": f"Polygon {self.shape_id}"})
        ElementTree.SubElement(nvsppr, "p:cNvSpPr")
        ElementTree.SubElement(nvsppr, "p:nvPr")

        sppr = ElementTree.SubElement(polygon_node, "p:spPr")
        xfrm = ElementTree.SubElement(sppr, "a:xfrm", {"rot": self.__get_angle(polygon.get("rotate", 0))})
        ElementTree.SubElement(xfrm, "a:off", {"x": self.__get_coordinate(x), "y": self.__get_coordinate(y)})
        ElementTree.SubElement(xfrm, "a:ext", {"cx": self.__get_coordinate(width), "cy": self.__get_coordinate(height)})

        cust_geom = ElementTree.SubElement(sppr, "a:custGeom")
        ElementTree.SubElement(cust_geom, "a:avLst")
        ElementTree.SubElement(cust_geom, "a:ahLst")
        ElementTree.SubElement(cust_geom, "a:rect", {"b": "b", "l": "l", "r": "r", "t": "t"})

        path = ElementTree.SubElement(ElementTree.SubElement(cust_geom, "a:pathLst"), "a:path", {"w": self.__get_coordinate(width), "h": self.__get_coordinate(height)})
        for i, point in enumerate(polygon["points"]):
            to = ElementTree.SubElement(path, "a:moveTo" if i == 0 else "a:lnTo")
            ElementTree.SubElement(to, "a:pt", {"x": self.__get_coordinate(point["x"] - x), "y": self.__get_coordinate(point["y"] - y)})
        ElementTree.SubElement(path, "a:close")

        self.__set_fill(sppr, config=polygon)
        self.__set_stroke(sppr, config=polygon)
        self.shape_id += 1

        return polygon_node

    def make_group(self, x: float, y: float, width: float, height: float) -> ElementTree.Element:
        group_node = ElementTree.Element("p:grpSp")

        nvgrpsppr = ElementTree.SubElement(group_node, "p:nvGrpSpPr")
        ElementTree.SubElement(nvgrpsppr, "p:cNvPr", {"id": str(self.shape_id), "name": f"Group {self.shape_id}"})
        ElementTree.SubElement(nvgrpsppr, "p:cNvGrpSpPr")
        ElementTree.SubElement(nvgrpsppr, "p:nvPr")

        grpsppr = ElementTree.SubElement(group_node, "p:grpSpPr")
        xfrm = ElementTree.SubElement(grpsppr, "a:xfrm")
        ElementTree.SubElement(xfrm, "a:off", {"x": self.__get_coordinate(x), "y": self.__get_coordinate(y)})
        ElementTree.SubElement(xfrm, "a:ext", {"cx": self.__get_coordinate(width), "cy": self.__get_coordinate(height)})
        ElementTree.SubElement(xfrm, "a:chOff", {"x": self.__get_coordinate(x), "y": self.__get_coordinate(y)})
        ElementTree.SubElement(xfrm, "a:chExt", {"cx": self.__get_coordinate(width), "cy": self.__get_coordinate(height)})
        self.shape_id += 1

        return group_node

    def save(self, path: str) -> None:
        self.tree.write(self.slide_path)

        with zipfile.ZipFile(path, "w") as f:
            for root, _, files in os.walk(self.work_path):
                for file in files:
                    f.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), self.work_path))

    def __set_fill(self, node: ElementTree.Element, config: dict) -> None:
        if "fill" not in config:
            return

        fill = ElementTree.SubElement(ElementTree.SubElement(node, "a:solidFill"), "a:srgbClr", {"val": self.__get_color(config["fill"])})
        opacity = config.get("fill-opacity", config.get("opacity", 1))

        if opacity < 1:
            ElementTree.SubElement(fill, "a:alpha", {"val": self.__get_fraction(opacity)})

    def __set_stroke(self, node: ElementTree.Element, config: dict) -> None:
        if not config.get("stroke"):
            return

        ln = ElementTree.SubElement(node, "a:ln", {"w": self.__get_size(config.get("thickness", 1))})
        fill = ElementTree.SubElement(ElementTree.SubElement(ln, "a:solidFill"), "a:srgbClr", {"val": self.__get_color(config["stroke"])})

        opacity = config.get("stroke-opacity", config.get("opacity", 1))
        if opacity < 1:
            ElementTree.SubElement(fill, "a:alpha", {"val": self.__get_fraction(opacity)})

    def __get_tree(self, root: ElementTree.Element) -> ElementTree.Element:
        namespaces = {
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
        }

        for key, namespace in namespaces.items():
            ElementTree.register_namespace(key, namespace)

        return root.find("p:cSld", namespaces).find("p:spTree", namespaces)

    def __get_lines_bbox(self, lines: List[dict]) -> Tuple[float, float, float, float]:
        x1, y1, x2, y2 = lines[0]["x1"], lines[0]["y1"], lines[0]["x2"], lines[0]["y2"]
        xmin, xmax = min(x1, x2), max(x1, x2)
        ymin, ymax = min(y1, y2), max(y1, y2)

        for line in lines:
            x1, y1, x2, y2 = line["x1"], line["y1"], line["x2"], line["y2"]
            xmin, ymin = min(xmin, x1, x2), min(ymin, y1, y2)
            xmax, ymax = max(xmax, x1, x2), max(ymax, y1, y2)

        return xmin, ymin, xmax - xmin, ymax - ymin

    def __get_ellipses_bbox(self, ellipses: List[dict]) -> Tuple[float, float, float, float]:
        xmin = xmax = ellipses[0]["x"]
        ymin = ymax = ellipses[0]["y"]

        for ellipse in ellipses:
            dx, dy = self.__get_diameter(ellipse, key="dx"), self.__get_diameter(ellipse, key="dy")
            angle = ellipse.get("rotate", 0)

            if angle == 0:
                x1, y1, x2, y2 = ellipse["x"], ellipse["y"], ellipse["x"] + dx, ellipse["y"] + dy
            else:
                cx, cy = ellipse["x"] + dx / 2, ellipse["y"] + dy / 2
                r = angle / 180 * math.pi
                t1 = math.atan(-dy / dx * math.tan(r))
                t2 = math.atan(dy / dx / math.tan(r))
                w = dx * math.cos(t1) * math.cos(r) - dy * math.sin(t1) * math.sin(r)
                h = -dy * math.sin(t2) * math.cos(r) - dx * math.cos(t2) * math.sin(r)
                x1, y1, x2, y2 = cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2

            xmin, ymin = min(xmin, x1), min(ymin, y1)
            xmax, ymax = max(xmax, x2), max(ymax, y2)

        return xmin, ymin, xmax - xmin, ymax - ymin

    def __get_rectangles_bbox(self, rectangles: List[dict]) -> Tuple[float, float, float, float]:
        polygons = []

        for rectangle in rectangles:
            x, y, w, h = rectangle["x"], rectangle["y"], rectangle["w"], rectangle["h"]

            polygons.append({
                "points": [{"x": x, "y": y}, {"x": x, "y": y + h}, {"x": x + w, "y": y}, {"x": x + w, "y": y + h}],
                "rotate": rectangle.get("rotate", 0)
            })

        return self.__get_polygons_bbox(polygons=polygons)

    def __get_polygons_bbox(self, polygons: List[dict]) -> Tuple[float, float, float, float]:
        xmin = xmax = polygons[0]["points"][0]["x"]
        ymin = ymax = polygons[0]["points"][0]["y"]

        for polygon in polygons:
            angle = polygon.get("rotate", 0) / 180 * math.pi
            cx = sum(point["x"] for point in polygon["points"]) / len(polygon["points"])
            cy = sum(point["y"] for point in polygon["points"]) / len(polygon["points"])

            for point in polygon["points"]:
                dx, dy = point["x"] - cx, point["y"] - cy
                x = cx + dx * math.cos(angle) - dy * math.sin(angle)
                y = cy + dx * math.sin(angle) + dy * math.cos(angle)

                xmin, ymin = min(xmin, x), min(ymin, y)
                xmax, ymax = max(xmax, x), max(ymax, y)

        return xmin, ymin, xmax - xmin, ymax - ymin

    def __get_coordinate(self, coordinate: float) -> str:
        return str(round(coordinate * 360000))

    def __get_size(self, size: float) -> str:
        return str(round(size * 12700))

    def __get_color(self, color: str) -> str:
        if color.startswith("#"):
            color = color[1:]

        color = color.upper()

        if re.fullmatch(r"[\dA-F]{3}", color):
            r, g, b = color
            return f"{r}{r}{g}{g}{b}{b}"

        return color

    def __get_fraction(self, fraction: float) -> str:
        return str(round(max(0.0, min(1.0, fraction)) * 100000))

    def __get_angle(self, angle: float) -> str:
        return str(round(angle * 60000))

    def __get_diameter(self, ellipse: dict, key: str) -> float:
        return ellipse[key] if key in ellipse else ellipse["d"]
