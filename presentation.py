import math
import os
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from typing import List, Optional, Tuple

from lxml import etree


@dataclass
class Slide:
    tree: etree.ElementTree
    sp_tree: etree.ElementTree


class Presentation:
    def __init__(self, presentation_path: str, work_path: str) -> None:
        self.presentation_path = presentation_path
        self.work_path = work_path

        with zipfile.ZipFile(presentation_path, "r") as f:
            f.extractall(self.work_path)

        self.namespaces = {
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }

        self.slides = {}
        self.shape_id = 1

    def add_line(self, line: dict, slide: str = "slide1") -> None:
        self.__add_to_slide(self.make_line(line=line), slide=slide)

    def add_ellipse(self, ellipse: dict, slide: str = "slide1") -> None:
        self.__add_to_slide(self.make_ellipse(ellipse=ellipse), slide=slide)

    def add_rectangle(self, rectangle: dict, slide: str = "slide1") -> None:
        self.__add_to_slide(self.make_rectangle(rectangle=rectangle), slide=slide)

    def add_polygon(self, polygon: dict, slide: str = "slide1") -> None:
        self.__add_to_slide(self.make_polygon(polygon=polygon), slide=slide)

    def add_textbox(self, textbox: dict, slide: str = "slide1") -> None:
        self.__add_to_slide(self.make_textbox(textbox=textbox), slide=slide)

    def add_lines(self, lines: List[dict], slide: str = "slide1") -> None:
        if not lines:
            return

        x, y, width, height = self.__get_lines_bbox(lines=lines)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for line in lines:
            group.append(self.make_line(line=line))

        self.__add_to_slide(group, slide=slide)

    def add_ellipses(self, ellipses: List[dict], slide: str = "slide1") -> None:
        if not ellipses:
            return

        x, y, width, height = self.__get_ellipses_bbox(ellipses=ellipses)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for ellipse in ellipses:
            group.append(self.make_ellipse(ellipse=ellipse))

        self.__add_to_slide(group, slide=slide)

    def add_rectangles(self, rectangles: List[dict], slide: str = "slide1") -> None:
        if not rectangles:
            return

        x, y, width, height = self.__get_rectangles_bbox(rectangles=rectangles)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for rectangle in rectangles:
            group.append(self.make_rectangle(rectangle=rectangle))

        self.__add_to_slide(group, slide=slide)

    def add_polygons(self, polygons: List[dict], slide: str = "slide1") -> None:
        if not polygons:
            return

        x, y, width, height = self.__get_polygons_bbox(polygons=polygons)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for polygon in polygons:
            group.append(self.make_polygon(polygon=polygon))

        self.__add_to_slide(group, slide=slide)

    def add_textboxes(self, textboxes: List[dict], slide: str = "slide1") -> None:
        if not textboxes:
            return

        x, y, width, height = self.__get_textboxes_bbox(textboxes=textboxes)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for textbox in textboxes:
            group.append(self.make_textbox(textbox=textbox))

        self.__add_to_slide(group, slide=slide)

    def add_shapes(self, shapes: List[dict], slide: str = "slide1") -> None:
        if not shapes:
            return

        x, y, width, height = self.__get_shapes_bbox(shapes=shapes)
        group = self.make_group(x=x, y=y, width=width, height=height)

        for shape in shapes:
            group.append(self.make_shape(shape=shape))

        self.__add_to_slide(group, slide=slide)

    def make_line(self, line: dict) -> etree.Element:
        line_node = self.__element("p:cxnSp")

        nvcxnsppr = self.__element("p:nvCxnSpPr", parent=line_node)
        self.__element("p:cNvPr", {"id": str(self.shape_id), "name": f"Line {self.shape_id}"}, parent=nvcxnsppr)
        self.__element("p:cNvCxnSpPr", parent=nvcxnsppr)
        self.__element("p:nvPr", parent=nvcxnsppr)

        x1, y1, x2, y2 = line["x1"], line["y1"], line["x2"], line["y2"]
        x, y, w, h = min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1)

        sppr = self.__element("p:spPr", parent=line_node)
        self.__set_xfrm(sppr, {"flipH": "0" if x1 <= x2 else "1", "flipV": "0" if y1 <= y2 else "1"}, x=x, y=y, cx=w, cy=h)
        self.__element("a:avLst", parent=self.__element("a:prstGeom", {"prst": "line"}, parent=sppr))

        self.__set_stroke(sppr, config=line)
        self.shape_id += 1

        return line_node

    def make_ellipse(self, ellipse: dict) -> etree.Element:
        ellipse_node = self.__element("p:sp")

        nvsppr = self.__element("p:nvSpPr", parent=ellipse_node)
        self.__element("p:cNvPr", {"id": str(self.shape_id), "name": f"Ellipse {self.shape_id}"}, parent=nvsppr)
        self.__element("p:cNvSpPr", parent=nvsppr)
        self.__element("p:nvPr", parent=nvsppr)

        dx, dy = self.__get_diameter(ellipse, key="dx"), self.__get_diameter(ellipse, key="dy")

        sppr = self.__element("p:spPr", parent=ellipse_node)
        self.__set_xfrm(sppr, {"rot": self.__get_angle(ellipse.get("rotate", 0))}, x=ellipse["x"], y=ellipse["y"], cx=dx, cy=dy)
        self.__element("a:avLst", parent=self.__element("a:prstGeom", {"prst": "ellipse"}, parent=sppr))

        self.__set_fill(sppr, config=ellipse)
        self.__set_stroke(sppr, config=ellipse)
        self.shape_id += 1

        return ellipse_node

    def make_rectangle(self, rectangle: dict) -> etree.Element:
        rectangle_node = self.__element("p:sp")

        nvsppr = self.__element("p:nvSpPr", parent=rectangle_node)
        self.__element("p:cNvPr", {"id": str(self.shape_id), "name": f"Rectangle {self.shape_id}"}, parent=nvsppr)
        self.__element("p:cNvSpPr", parent=nvsppr)
        self.__element("p:nvPr", parent=nvsppr)

        sppr = self.__element("p:spPr", parent=rectangle_node)
        self.__set_xfrm(sppr, {"rot": self.__get_angle(rectangle.get("rotate", 0))}, x=rectangle["x"], y=rectangle["y"], cx=rectangle["w"], cy=rectangle["h"])

        geom = self.__element("a:prstGeom", {"prst": "roundRect"}, parent=sppr)
        avlst = self.__element("a:avLst", parent=geom)
        self.__element("a:gd", {"name": "adj", "fmla": f'val {self.__get_fraction(rectangle.get("radius", 0) / 2)}'}, parent=avlst)

        self.__set_fill(sppr, config=rectangle)
        self.__set_stroke(sppr, config=rectangle)
        self.shape_id += 1

        return rectangle_node

    def make_polygon(self, polygon: dict) -> etree.Element:
        x, y, width, height = self.__get_polygons_bbox(polygons=[polygon])

        polygon_node = self.__element("p:sp")

        nvsppr = self.__element("p:nvSpPr", parent=polygon_node)
        self.__element("p:cNvPr", {"id": str(self.shape_id), "name": f"Polygon {self.shape_id}"}, parent=nvsppr)
        self.__element("p:cNvSpPr", parent=nvsppr)
        self.__element("p:nvPr", parent=nvsppr)

        sppr = self.__element("p:spPr", parent=polygon_node)
        self.__set_xfrm(sppr, {"rot": self.__get_angle(polygon.get("rotate", 0))}, x=x, y=y, cx=width, cy=height)

        cust_geom = self.__element("a:custGeom", parent=sppr)
        self.__element("a:avLst", parent=cust_geom)
        self.__element("a:ahLst", parent=cust_geom)
        self.__element("a:rect", {"b": "b", "l": "l", "r": "r", "t": "t"}, parent=cust_geom)

        path = self.__element("a:path", {"w": self.__get_coordinate(width), "h": self.__get_coordinate(height)}, parent=self.__element("a:pathLst", parent=cust_geom))
        for i, point in enumerate(polygon["points"]):
            to = self.__element("a:moveTo" if i == 0 else "a:lnTo", parent=path)
            self.__element("a:pt", {"x": self.__get_coordinate(point["x"] - x), "y": self.__get_coordinate(point["y"] - y)}, parent=to)
        self.__element("a:close", parent=path)

        self.__set_fill(sppr, config=polygon)
        self.__set_stroke(sppr, config=polygon)
        self.shape_id += 1

        return polygon_node

    def make_textbox(self, textbox: dict) -> etree.Element:
        textbox_node = self.__element("p:sp")

        nvsppr = self.__element("p:nvSpPr", parent=textbox_node)
        self.__element("p:cNvPr", {"id": str(self.shape_id), "name": f"TextBox {self.shape_id}"}, parent=nvsppr)
        self.__element("p:cNvSpPr", {"txBox": "1"}, parent=nvsppr)
        self.__element("p:nvPr", parent=nvsppr)

        sppr = self.__element("p:spPr", parent=textbox_node)
        self.__set_xfrm(sppr, {"rot": self.__get_angle(textbox.get("rotate", 0))}, x=textbox["x"], y=textbox["y"], cx=textbox["w"], cy=textbox["h"])
        self.__element("a:avLst", parent=self.__element("a:prstGeom", {"prst": "rect"}, parent=sppr))

        tx_body = self.__element("p:txBody", parent=textbox_node)
        margin = textbox.get("margin", {"left": 0, "right": 0, "top": 0, "bottom": 0})
        body_attributes = {
            "anchor": self.__get_alignment(textbox.get("vertical-align", "center")),
            "anchorCtr": "0",
            "rtlCol": "0",
            "bIns": self.__get_coordinate(margin.get("bottom", 0.1)),
            "lIns": self.__get_coordinate(margin.get("left", 0.25)),
            "rIns": self.__get_coordinate(margin.get("right", 0.25)),
            "tIns": self.__get_coordinate(margin.get("top", 0.1)),
            "wrap": "square"
        }

        body_pr = self.__element("a:bodyPr", body_attributes, parent=tx_body)
        self.__element("a:spAutoFit" if textbox.get("auto-fit", False) else "a:noAutoFit", parent=body_pr)
        self.__element("a:lstStyle", parent=tx_body)

        text_attributes = self.__get_text_formatting(formatting=textbox)

        for line in textbox["text"].split("\n"):
            p = self.__element("a:p", parent=tx_body)
            self.__element("a:pPr", {"algn": self.__get_alignment(textbox["align"])}, parent=p)
            r = self.__element("a:r", parent=p)
            rpr = self.__element("a:rPr", {"smtClean": "0", **text_attributes}, parent=r)

            if "color" in textbox:
                self.__element("a:srgbClr", {"val": self.__get_color(textbox["color"])}, parent=self.__element("a:solidFill", parent=rpr))

            self.__element("a:t", parent=r).text = line
            self.__element("a:endParaRPr", text_attributes, parent=p)

        self.__set_fill(sppr, config=textbox)
        self.__set_stroke(sppr, config=textbox)
        self.shape_id += 1

        return textbox_node

    def make_shape(self, shape: dict) -> etree.Element:
        if shape["shape"] == "line":
            return self.make_line(line=shape)

        if shape["shape"] == "ellipse":
            return self.make_ellipse(ellipse=shape)

        if shape["shape"] == "rectangle":
            return self.make_rectangle(rectangle=shape)

        if shape["shape"] == "polygon":
            return self.make_polygon(polygon=shape)

        if shape["shape"] == "textbox":
            return self.make_textbox(textbox=shape)

        raise ValueError(f'Unknown shape type "{shape["shape"]}"')

    def make_group(self, x: float, y: float, width: float, height: float) -> etree.Element:
        group_node = self.__element("p:grpSp")

        nvgrpsppr = self.__element("p:nvGrpSpPr", parent=group_node)
        self.__element("p:cNvPr", {"id": str(self.shape_id), "name": f"Group {self.shape_id}"}, parent=nvgrpsppr)
        self.__element("p:cNvGrpSpPr", parent=nvgrpsppr)
        self.__element("p:nvPr", parent=nvgrpsppr)

        grpsppr = self.__element("p:grpSpPr", parent=group_node)
        xfrm = self.__element("a:xfrm", parent=grpsppr)
        self.__element("a:off", {"x": self.__get_coordinate(x), "y": self.__get_coordinate(y)}, parent=xfrm)
        self.__element("a:ext", {"cx": self.__get_coordinate(width), "cy": self.__get_coordinate(height)}, parent=xfrm)
        self.__element("a:chOff", {"x": self.__get_coordinate(x), "y": self.__get_coordinate(y)}, parent=xfrm)
        self.__element("a:chExt", {"cx": self.__get_coordinate(width), "cy": self.__get_coordinate(height)}, parent=xfrm)
        self.shape_id += 1

        return group_node

    def save(self, path: str) -> None:
        for name, slide in self.slides.items():
            slide.tree.write(os.path.join(self.work_path, "ppt", "slides", f"{name}.xml"))

        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as f:
            for root, _, files in os.walk(self.work_path):
                for file in files:
                    f.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), self.work_path))

    def __add_to_slide(self, node: etree.ElementTree, slide: str = "slide1") -> None:
        if slide not in self.slides:
            tree = etree.parse(os.path.join(self.work_path, "ppt", "slides", f"{slide}.xml"))
            sp_tree = tree.getroot().find("p:cSld", self.namespaces).find("p:spTree", self.namespaces)
            self.slides[slide] = Slide(tree=tree, sp_tree=sp_tree)

        self.slides[slide].sp_tree.append(node)

    def __element(self, tag: str, attrib: Optional[dict] = None, parent: Optional[etree.Element] = None) -> etree.Element:
        namespace, tag_name = tag.split(":")

        if attrib is None:
            attrib = {}

        if parent is not None:
            return etree.SubElement(parent, etree.QName(self.namespaces[namespace], tag_name), attrib=attrib)

        return etree.Element(etree.QName(self.namespaces[namespace], tag_name), attrib=attrib)

    def __set_fill(self, node: etree.Element, config: dict) -> None:
        if "fill" not in config:
            return

        fill = self.__element("a:srgbClr", {"val": self.__get_color(config["fill"])}, parent=self.__element("a:solidFill", parent=node))
        opacity = config.get("fill-opacity", config.get("opacity", 1))

        if opacity < 1:
            self.__element("a:alpha", {"val": self.__get_fraction(opacity)}, parent=fill)

    def __set_stroke(self, node: etree.Element, config: dict) -> None:
        if not config.get("stroke"):
            return

        ln = self.__element("a:ln", {"w": self.__get_size(config.get("thickness", 1))}, parent=node)
        fill = self.__element("a:srgbClr", {"val": self.__get_color(config["stroke"])}, parent=self.__element("a:solidFill", parent=ln))

        if config.get("stroke-dash", "solid") != "solid":
            self.__element("a:prstDash", {"val": self.__get_stroke_dash(config["stroke-dash"])}, parent=ln)

        opacity = config.get("stroke-opacity", config.get("opacity", 1))
        if opacity < 1:
            self.__element("a:alpha", {"val": self.__get_fraction(opacity)}, parent=fill)

    def __set_xfrm(self, sppr: etree.Element, attrib: dict, x: float, y: float, cx: float, cy: float) -> etree.Element:
        xfrm = self.__element("a:xfrm", attrib, parent=sppr)
        self.__element("a:off", {"x": self.__get_coordinate(x), "y": self.__get_coordinate(y)}, parent=xfrm)
        self.__element("a:ext", {"cx": self.__get_coordinate(cx), "cy": self.__get_coordinate(cy)}, parent=xfrm)
        return xfrm

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

    def __get_textboxes_bbox(self, textboxes: List[dict]) -> Tuple[float, float, float, float]:
        return self.__get_rectangles_bbox(rectangles=textboxes)

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
            elif shape_type == "textbox":
                x, y, w, h = self.__get_textboxes_bbox(textboxes=elements)
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

    def __get_text_formatting(self, formatting: dict) -> dict:
        text_attributes = {"dirty": "0", "sz": self.__get_font_size(formatting["size"])}

        if formatting.get("bold", False):
            text_attributes["b"] = "1"

        if formatting.get("italic", False):
            text_attributes["i"] = "1"

        if formatting.get("underline", False):
            text_attributes["u"] = "sng"

        if formatting.get("strike", False):
            text_attributes["strike"] = "sngStrike"

        return text_attributes

    def __get_coordinate(self, coordinate: float) -> str:
        return str(round(coordinate * 360000))

    def __get_size(self, size: float) -> str:
        return str(round(size * 12700))

    def __get_font_size(self, size: float) -> str:
        return str(round(size * 100))

    def __get_alignment(self, alignment: str) -> str:
        alignment2pptx = {"center": "ctr", "left": "l", "right": "r", "top": "t", "bottom": "b"}
        return alignment2pptx[alignment]

    def __get_color(self, color: str) -> str:
        if color.startswith("#"):
            color = color[1:]

        color = color.upper()

        if re.fullmatch(r"[\dA-F]{3}", color):
            r, g, b = color
            return f"{r}{r}{g}{g}{b}{b}"

        if re.fullmatch(r"[\dA-F]{6}", color):
            return color

        if match := re.fullmatch(r"RGB\((?P<r>\d{1,3}),\s*(?P<g>\d{1,3}),\s*(?P<b>\d{1,3})\)", color, re.I):
            r, g, b = min(int(match.group("r")), 255), min(int(match.group("g")), 255), min(int(match.group("b")), 255)
            digits = "0123456789ABCDEF"
            return f"{digits[r >> 4]}{digits[r & 15]}{digits[g >> 4]}{digits[g & 15]}{digits[b >> 4]}{digits[b & 15]}"

        raise ValueError(f'Invalid color "{color}"')

    def __get_stroke_dash(self, stroke_dash: str) -> str:
        dash2pptx = {
            "solid": "solid",
            "dashed": "dash",
            "dotted": "sysDot",
            "short-dashed": "sysDash",
            "dash-dotted": "dashDot",
            "long-dash": "lgDash",
            "long-dash-dotted": "lgDashDot",
            "long-dash-dot-dotted": "lgDashDotDot"
        }
        return dash2pptx[stroke_dash]

    def __get_fraction(self, fraction: float) -> str:
        return str(round(max(0.0, min(1.0, fraction)) * 100000))

    def __get_angle(self, angle: float) -> str:
        return str(round(angle * 60000))

    def __get_diameter(self, ellipse: dict, key: str) -> float:
        return ellipse[key] if key in ellipse else ellipse["d"]
