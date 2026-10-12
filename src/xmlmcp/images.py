from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO

from PIL import Image

from .errors import require, XMLMCPError
from .models import AddImage, ReplaceImage, Operation
from .namespaces import A, R, XDR, S
from .package import absolute_file
from .workbook import Workbook, Rectangle
from .xmlbytes import XMLDocument, Edit, append_children, set_attribute, xml_escape, edit_attributes

EMU_PER_POINT = 12700


@dataclass
class ImageChange:
    operation: AddImage | ReplaceImage
    data: bytes
    width: int
    height: int
    extension: str
    sheet: str
    rectangle: Rectangle | None = None
    image: dict | None = None

    @property
    def geometry_key(self) -> str:
        return self.operation.image_id if isinstance(self.operation, ReplaceImage) else f"{self.sheet}!{self.rectangle.address}"


def preflight_images(book: Workbook, operations: list[Operation]) -> list[ImageChange]:
    inventory = book.images()
    images = {image["image_id"]: image for image in inventory}
    replacements = set()
    changes = []
    for op in operations:
        if not isinstance(op, (AddImage, ReplaceImage)):
            continue
        path = absolute_file(op.image_path)
        data = path.read_bytes()
        try:
            with Image.open(BytesIO(data)) as image:
                require(image.format in ("PNG", "JPEG"), "IMAGE_TYPE", "图片仅支持 PNG 或 JPEG")
                width, height = image.size
                require(image.getexif().get(274, 1) == 1, "IMAGE_ORIENTATION_UNSUPPORTED", "带 EXIF 旋转的图片尚未完成双 Office 验证，请使用方向已固定的图片")
                extension = "png" if image.format == "PNG" else "jpeg"
            # PNG EXIF inspection can consume its stream; verify a fresh decoder.
            with Image.open(BytesIO(data)) as image:
                image.verify()
        except XMLMCPError:
            raise
        except Exception as exc:
            raise XMLMCPError("INVALID_IMAGE", "无法读取图片", error=str(exc)) from exc
        require(width > 0 and height > 0, "INVALID_IMAGE", "图片尺寸无效")
        if isinstance(op, AddImage):
            sheet, rect = book.resolve(op.target)
            require(len(book.package.related(sheet["part"], "drawing")) <= 1, "MULTIPLE_DRAWINGS", "工作表存在多个 drawing，不能可靠插入")
            changes.append(ImageChange(op, data, width, height, extension, sheet["name"], rect))
        else:
            require(op.image_id in images, "IMAGE_NOT_FOUND", "图片 ID 不存在；请先检查当前模板")
            require(op.image_id not in replacements, "TARGET_CONFLICT", "同一图片不能重复替换")
            replacements.add(op.image_id)
            original = images[op.image_id]
            require(sum(i["image_id"] == op.image_id for i in inventory) == 1,
                    "IMAGE_MAPPING", "模板包含重复图片对象 ID，不能唯一替换；可以保留这些图片并向其他区域新增图片")
            require(sum(i["sheet"] == original["sheet"] and i["name"] == original["name"] for i in inventory) == 1,
                    "IMAGE_MAPPING", "模板包含重复图片名称，无法向 Excel 唯一查询几何尺寸")
            require(original["replaceable"] and original["media_part"], "IMAGE_UNSUPPORTED", "不支持分组或外链图片替换")
            pic = original["_pic"]
            for node in pic.descendants():
                if node.ns == A:
                    supported_metadata = node.local == "extLst" and all(
                        child.local == "ext" and all((item.local, item.ns) in {
                            ("creationId", "http://schemas.microsoft.com/office/drawing/2014/main"),
                            ("useLocalDpi", "http://schemas.microsoft.com/office/drawing/2010/main")
                        } for item in child.children) for child in node.children)
                    require(node.local not in ("custGeom", "effectDag", "tile") and (node.local != "extLst" or supported_metadata)
                            and not (node.local == "effectLst" and node.children),
                            "IMAGE_UNSUPPORTED", "图片包含未适配的效果或扩展", element=node.local)
                    if node.local == "xfrm":
                        require(node.attrs.get("rot", "0") == "0" and node.attrs.get("flipH", "0") in ("0", "false")
                                and node.attrs.get("flipV", "0") in ("0", "false"), "IMAGE_UNSUPPORTED", "不支持旋转或翻转图片的替换")
            fill = pic.child("blipFill", XDR)
            require(fill and next(fill.descendants("blip", A), None), "IMAGE_UNSUPPORTED", "图片填充结构不支持")
            blip = next(fill.descendants("blip", A))
            require(all(n.name == A + "}extLst" for n in blip.children), "IMAGE_UNSUPPORTED", "不支持包含图片颜色或透明度变换的替换")
            require(fill.child("stretch", A), "IMAGE_UNSUPPORTED", "图片没有 stretch 填充，不能保持对象结构并适配")
            changes.append(ImageChange(op, data, width, height, extension, original["sheet"], image=original))
    return changes


def geometry_requests(changes: list[ImageChange]) -> list[dict]:
    result = []
    for change in changes:
        if change.rectangle:
            result.append({"key": change.geometry_key, "sheet": change.sheet, "range": change.rectangle.address})
        else:
            result.append({"key": change.geometry_key, "sheet": change.sheet, "shape_name": change.image["name"]})
    return result


def fill_insets(width: float, height: float, pixel_width: int, pixel_height: int) -> dict[str, str]:
    require(width > 0 and height > 0, "IMAGE_GEOMETRY", "图片目标区域的宽高必须大于零")
    scale = min(width / pixel_width, height / pixel_height)
    horizontal = round((width - scale * pixel_width) / (2 * width) * 100000)
    vertical = round((height - scale * pixel_height) / (2 * height) * 100000)
    return {"l": str(horizontal), "r": str(horizontal), "t": str(vertical), "b": str(vertical)}


def _fill_rect(attributes: dict[str, str]) -> bytes:
    return b'<a:fillRect xmlns:a="' + A.encode() + b'" ' + b" ".join(k.encode() + b'="' + v.encode() + b'"' for k, v in attributes.items()) + b"/>"


def _new_anchor(rect: Rectangle, shape_id: int, rid: str, width: float, height: float, insets: dict[str, str], left: float = 0, top: float = 0) -> bytes:
    def marker(tag: str, row: int, col: int):
        return (f"<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{tag}>")
    body = (f'<xdr:twoCellAnchor xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{R}">'
            + marker("from", rect.top - 1, rect.left - 1) + marker("to", rect.bottom, rect.right)
            + f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{shape_id}" name="XMLMCP Image {shape_id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
            + f'<xdr:blipFill><a:blip r:embed="{rid}"/><a:srcRect l="0" t="0" r="0" b="0"/><a:stretch>')
    ending = (f'</a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="{round(left*EMU_PER_POINT)}" y="{round(top*EMU_PER_POINT)}"/><a:ext cx="{round(width*EMU_PER_POINT)}" cy="{round(height*EMU_PER_POINT)}"/></a:xfrm>'
              '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>')
    return body.encode() + _fill_rect(insets) + ending.encode()


def apply_image_changes(book: Workbook, changes: list[ImageChange], geometry: dict[str, dict]) -> None:
    for change in changes:
        require(change.geometry_key in geometry, "IMAGE_GEOMETRY", "Excel 未返回目标几何尺寸")
        dimensions = geometry[change.geometry_key]
        width, height = dimensions["width"], dimensions["height"]
        insets = fill_insets(width, height, change.width, change.height)
        package = book.package
        media = package.fresh_part("xl/media", "xmlmcpImage", change.extension)
        package.add(media, change.data, "image:media")
        package.add_content_type(media, "image/png" if change.extension == "png" else "image/jpeg", change.extension)
        if change.rectangle:
            sheet = book.sheet(change.sheet)
            relations = package.related(sheet["part"], "drawing")
            if relations:
                drawing_part = relations[0]["part"]
            else:
                drawing_part = package.fresh_part("xl/drawings", "xmlmcpDrawing", "xml")
                package.add(drawing_part, f'<xdr:wsDr xmlns:xdr="{XDR}"/>'.encode(), "image:drawing")
                package.add_content_type(drawing_part, "application/vnd.openxmlformats-officedocument.drawing+xml")
                rid = package.add_relationship(sheet["part"], "drawing", drawing_part)
                doc = book.sheet_doc(sheet)
                tag = doc.root.qualify("drawing")
                fragment = b"<" + tag + b' xmlns:r="' + R.encode() + b'" r:id="' + rid.encode() + b'"/>'
                later = {"legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"}
                node = next((n for n in doc.root.children if n.ns == S and n.local in later), None)
                offset = node.start if node else doc.root.close_start
                package.patch(sheet["part"], [Edit(offset, offset, fragment, "image:sheet_drawing", change.sheet)])
            rid = package.add_relationship(drawing_part, "image", media)
            doc = package.doc(drawing_part)
            ids = [int(n.attrs["id"]) for n in doc.root.descendants("cNvPr", XDR)]
            shape_id = max(ids, default=0) + 1
            fragment = _new_anchor(change.rectangle, shape_id, rid, width, height, insets, dimensions.get("left", 0), dimensions.get("top", 0))
            package.patch(drawing_part, [append_children(doc.data, doc.root, fragment, "image:new_anchor", change.geometry_key)])
        else:
            original = next((i for i in book.images() if i["image_id"] == change.operation.image_id), None)
            require(original, "IMAGE_MAPPING", "图片在修改期间无法定位")
            drawing_part = original["drawing_part"]
            rid = package.add_relationship(drawing_part, "image", media)
            doc = package.doc(drawing_part)
            pic = next(p for p in doc.root.descendants("pic", XDR)
                       if next(p.descendants("cNvPr", XDR)).attrs["id"] == original["shape_id"])
            fill = pic.child("blipFill", XDR)
            blip = next(fill.descendants("blip", A))
            # Keep its namespace bindings local, including templates using a different r prefix.
            opening = doc.data[blip.start:blip.open_end]
            import re
            prefix = next((m[1] for m in re.finditer(br"([^\s=<>]+):embed\s*=", opening)), None)
            require(prefix, "IMAGE_MAPPING", "图片媒体关系属性无法定位")
            new_tag = set_attribute(doc.data, blip, prefix + b":embed", rid)
            edits = [Edit(blip.start, blip.open_end, new_tag, "image:media_ref", change.geometry_key)]
            src = fill.child("srcRect", A)
            if src:
                raw = doc.raw(src)
                from .cells import _set_open_attribute
                tag = doc.data[src.start:src.open_end]
                for attr in ("l", "t", "r", "b"):
                    tag = _set_open_attribute(tag, attr.encode(), "0")
                edits.append(Edit(src.start, src.open_end, tag, "image:crop_reset", change.geometry_key))
            stretch = fill.child("stretch", A)
            require(stretch, "IMAGE_UNSUPPORTED", "图片缺少 stretch 填充，不能保留原对象结构并适配")
            rect = stretch.child("fillRect", A)
            if rect:
                from .cells import _set_open_attribute
                tag = doc.data[rect.start:rect.open_end]
                for attr, value in insets.items():
                    tag = _set_open_attribute(tag, attr.encode(), value)
                edits.append(Edit(rect.start, rect.open_end, tag, "image:fit", change.geometry_key))
            else:
                edits.append(append_children(doc.data, stretch, _fill_rect(insets), "image:fit", change.geometry_key))
            package.patch(drawing_part, edits)

