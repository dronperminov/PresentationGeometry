import math
import os
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List, Tuple
from xml.etree import ElementTree


@dataclass
class Slide:
    filename: str
    tree: ElementTree
    sp_tree: ElementTree


class Presentation:
    def __init__(self, presentation_path: str, work_path: str) -> None:
        self.presentation_path = presentation_path
        self.work_path = work_path

        with zipfile.ZipFile(presentation_path, "r") as f:
            f.extractall(self.work_path)

        self.slides = self.__get_slides()
        self.shape_id = 2

    def add_line(self, line: dict, slide: str = "slide1") -> None:
        self.slides[slide].sp_tree.append(self.make_line(line=line))

    def add_ellipse(self, ellipse: dict, slide: str = "slide1") -> None:
        self.slides[slide].sp_tree.append(self.make_ellipse(ellipse=ellipse))

    def add_rectangle(self, rectangle: dict, slide: str = "slide1") -> None:
        self.slides[slide].sp_tree.append(self.make_rectangle(rectangle=rectangle))

    def add_polygon(self, polygon: dict, slide: str = "slide1") -> None:
        self.slides[slide].sp_tree.append(self.make_polygon(polygon=polygon))

    def add_lines(self, lines: List[dict], slide: str = "slide1") -> None:
        if not lines:
            return

        x, y, width, height = self.__get_lines_bbox(lines=lines)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for line in lines:
            group.append(self.make_line(line=line))

        self.slides[slide].sp_tree.append(group)

    def add_ellipses(self, ellipses: List[dict], slide: str = "slide1") -> None:
        if not ellipses:
            return

        x, y, width, height = self.__get_ellipses_bbox(ellipses=ellipses)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for ellipse in ellipses:
            group.append(self.make_ellipse(ellipse=ellipse))

        self.slides[slide].sp_tree.append(group)

    def add_rectangles(self, rectangles: List[dict], slide: str = "slide1") -> None:
        if not rectangles:
            return

        x, y, width, height = self.__get_rectangles_bbox(rectangles=rectangles)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for rectangle in rectangles:
            group.append(self.make_rectangle(rectangle=rectangle))

        self.slides[slide].sp_tree.append(group)

    def add_polygons(self, polygons: List[dict], slide: str = "slide1") -> None:
        if not polygons:
            return

        x, y, width, height = self.__get_polygons_bbox(polygons=polygons)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for polygon in polygons:
            group.append(self.make_polygon(polygon=polygon))

        self.slides[slide].sp_tree.append(group)

    def add_shapes(self, shapes: List[dict], slide: str = "slide1") -> None:
        if not shapes:
            return

        x, y, width, height = self.__get_shapes_bbox(shapes=shapes)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for shape in shapes:
            group.append(self.make_shape(shape=shape))

        self.slides[slide].sp_tree.append(group)

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

    def make_shape(self, shape: dict) -> ElementTree.Element:
        if shape["shape"] == "line":
            return self.make_line(line=shape)

        if shape["shape"] == "ellipse":
            return self.make_ellipse(ellipse=shape)

        if shape["shape"] == "rectangle":
            return self.make_rectangle(rectangle=shape)

        if shape["shape"] == "polygon":
            return self.make_polygon(polygon=shape)

        raise ValueError(f'Unknown shape type "{shape["shape"]}"')

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
        for slide in self.slides.values():
            slide.tree.write(os.path.join(self.work_path, "ppt", "slides", slide.filename))

        with zipfile.ZipFile(path, "w") as f:
            for root, _, files in os.walk(self.work_path):
                for file in files:
                    f.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), self.work_path))

    def __get_slides(self) -> Dict[str, Slide]:
        slide2tree = {}

        for filename in os.listdir(os.path.join(self.work_path, "ppt", "slides")):
            if not filename.endswith(".xml"):
                continue

            name = filename.replace(".xml", "")
            tree = ElementTree.parse(os.path.join(self.work_path, "ppt", "slides", filename))
            sp_tree = self.__get_sp_tree(root=tree.getroot())
            slide2tree[name] = Slide(filename=filename, tree=tree, sp_tree=sp_tree)

        return slide2tree

    def __get_sp_tree(self, root: ElementTree.Element) -> ElementTree.Element:
        namespaces = {
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
        }

        for key, namespace in namespaces.items():
            ElementTree.register_namespace(key, namespace)

        return root.find("p:cSld", namespaces).find("p:spTree", namespaces)

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

    def __get_lines_bbox(self, lines: List[dict]) -> Tuple[float, float, float, float]:
        x1, y1, x2, y2 = lines[0]["x1"], lines[0]["y1"], lines[0]["x2"], lines[0]["y2"]
        x_min, x_max = min(x1, x2), max(x1, x2)
        y_min, y_max = min(y1, y2), max(y1, y2)

        for line in lines:
            x1, y1, x2, y2 = line["x1"], line["y1"], line["x2"], line["y2"]
            x_min, y_min = min(x_min, x1, x2), min(y_min, y1, y2)
            x_max, y_max = max(x_max, x1, x2), max(y_max, y1, y2)

        return x_min, y_min, x_max - x_min, y_max - y_min

    def __get_ellipses_bbox(self, ellipses: List[dict]) -> Tuple[float, float, float, float]:
        x_min = x_max = ellipses[0]["x"]
        y_min = y_max = ellipses[0]["y"]

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

            x_min, y_min = min(x_min, x1), min(y_min, y1)
            x_max, y_max = max(x_max, x2), max(y_max, y2)

        return x_min, y_min, x_max - x_min, y_max - y_min

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
        x_min = x_max = polygons[0]["points"][0]["x"]
        y_min = y_max = polygons[0]["points"][0]["y"]

        for polygon in polygons:
            angle = polygon.get("rotate", 0) / 180 * math.pi
            cx = sum(point["x"] for point in polygon["points"]) / len(polygon["points"])
            cy = sum(point["y"] for point in polygon["points"]) / len(polygon["points"])

            for point in polygon["points"]:
                dx, dy = point["x"] - cx, point["y"] - cy
                x = cx + dx * math.cos(angle) - dy * math.sin(angle)
                y = cy + dx * math.sin(angle) + dy * math.cos(angle)

                x_min, y_min = min(x_min, x), min(y_min, y)
                x_max, y_max = max(x_max, x), max(y_max, y)

        return x_min, y_min, x_max - x_min, y_max - y_min

    def __get_shapes_bbox(self, shapes: List[dict]) -> Tuple[float, float, float, float]:
        x_min, y_min, x_max, y_max = None, None, None, None
        shape2shapes = defaultdict(list)

        for shape in shapes:
            shape2shapes[shape["shape"]].append(shape)

        for shape_type, elements in shape2shapes.items():
            if shape_type == "line":
                x, y, w, h = self.__get_lines_bbox(lines=elements)
            elif shape_type == "ellipse":
                x, y, w, h = self.__get_ellipses_bbox(ellipses=elements)
            elif shape_type == "rectangle":
                x, y, w, h = self.__get_rectangles_bbox(rectangles=elements)
            elif shape_type == "polygon":
                x, y, w, h = self.__get_polygons_bbox(polygons=elements)
            else:
                raise ValueError(f'Unknown shape type "{shape_type}"')

            if x_min is None or x < x_min:
                x_min = x

            if y_min is None or y < y_min:
                y_min = y

            if x_max is None or x + w > x_max:
                x_max = x + w

            if y_max is None or y + h > y_max:
                y_max = y + h

        return x_min, y_min, x_max - x_min, y_max - y_min

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

        if re.fullmatch(r"[\dA-F]{6}", color):
            return color

        if match := re.fullmatch(r"RGB\((?P<r>\d{1,3}),\s*(?P<g>\d{1,3}),\s*(?P<b>\d{1,3})\)", color):
            r, g, b = min(int(match.group("r")), 255), min(int(match.group("g")), 255), min(int(match.group("b")), 255)
            digits = "0123456789ABCDEF"
            return f"{digits[r >> 4]}{digits[r & 15]}{digits[g >> 4]}{digits[g & 15]}{digits[b >> 4]}{digits[b & 15]}"

        raise ValueError(f'Invalid color "{color}"')

    def __get_fraction(self, fraction: float) -> str:
        return str(round(max(0.0, min(1.0, fraction)) * 100000))

    def __get_angle(self, angle: float) -> str:
        return str(round(angle * 60000))

    def __get_diameter(self, ellipse: dict, key: str) -> float:
        return ellipse[key] if key in ellipse else ellipse["d"]
