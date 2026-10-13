"""Synchronize reviewed checkbox selections in VML and ctrlProps only."""
from .errors import require
from .namespaces import R
from .xmlbytes import Edit, append_children, set_attribute
import re


def checkbox_text_pattern(label):
    return re.compile(r'([☐☑])\s*' + re.escape(label) + r'(?=$|\s)')


def apply_checkbox_text_values(book, values):
    """Change only each checkbox character; isolate shared strings and retain all runs."""
    from .models import Target
    from .namespaces import S
    from .xmlbytes import XMLDocument, standalone, apply_edits, xml_escape
    from .strings import encode_xstring
    seen = set()
    for item in values:
        sheet, rect = book.resolve(Target.model_validate(item['target']))
        rect = book.merged_anchor(sheet, rect, True)
        key = (sheet['name'], rect.address)
        require(rect.size == (1,1) and key not in seen, 'TARGET_CONFLICT', '方框文字目标重复或不是单格')
        seen.add(key)
        doc = book.sheet_doc(sheet)
        cell = book.cell_map(sheet).get(rect.address)
        source = book.string_node(cell)
        require(source is not None and not cell.child('f',S), 'CHECKBOX_MAPPING', '方框目标不是文字单元格')
        copied = XMLDocument(standalone(source.document, source))
        nodes = [n for n in copied.root.descendants('t', S)]
        original = ''.join(n.text for n in nodes)
        replacements = {}
        for selection in item['selections']:
            matches = list(checkbox_text_pattern(selection['label']).finditer(original))
            require(len(matches) == 1, 'CHECKBOX_MAPPING', '单元格方框标签不存在或不唯一', label=selection['label'])
            index = matches[0].start(1)
            require(index not in replacements, 'TARGET_CONFLICT', '方框标签重复指定')
            replacements[index] = '☑' if selection['checked'] else '☐'
        edits, offset = [], 0
        for node in nodes:
            text = ''.join(replacements.get(offset+i, char) for i,char in enumerate(node.text))
            offset += len(node.text)
            if text != node.text:
                opening = set_attribute(copied.data, node, b'xml:space', 'preserve')
                fragment = opening + xml_escape(encode_xstring(text)) + b'</' + node.raw_name + b'>'
                edits.append(Edit(node.start,node.end,fragment,'control:text_checkbox',rect.address))
        if not edits:
            continue
        payload = apply_edits(copied.data, edits)[0]
        root = XMLDocument(payload).root
        new_name = root.qualify('is')
        payload = payload[:root.start+1] + new_name + payload[root.start+1+len(root.raw_name):root.close_start] + b'</' + new_name + b'>' + payload[root.end:]
        old_value = cell.child('is',S) or cell.child('v',S)
        require(old_value is not None, 'CHECKBOX_MAPPING', '方框文字存储无效')
        book.package.patch(sheet['part'], [Edit(old_value.start,old_value.end,payload,'control:text_checkbox',rect.address),
                                           Edit(cell.start,cell.open_end,set_attribute(doc.data,cell,b't','inlineStr'),'control:text_checkbox',rect.address)])

V = "urn:schemas-microsoft-com:vml"
X = "urn:schemas-microsoft-com:office:excel"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"


def checkboxes(book, sheet_name):
    sheet = book.sheet(sheet_name)
    result = []
    for rel in book.package.related(sheet["part"], "vmlDrawing"):
        doc = book.package.doc(rel["part"])
        for shape in doc.root.descendants("shape", V):
            client = shape.child("ClientData", X)
            if client is None or client.attrs.get("ObjectType") != "Checkbox":
                continue
            labels = [n.text for n in shape.descendants("font")]
            if not labels:
                labels = [n.text for n in shape.descendants("div") if n.text.strip()]
            label = "".join(labels).strip()
            checked = client.child("Checked", X)
            require(not checked or checked.text.strip() in ("0", "1"), "CHECKBOX_MIXED", "不能复制混合状态复选框", label=label)
            require(not client.child("FmlaLink", X), "CHECKBOX_LINKED", "离线复制不支持关联单元格的复选框", label=label)
            shape_id = shape.attrs["id"].rsplit("s", 1)[-1]
            controls = [n for n in book.sheet_doc(sheet).root.descendants("control") if n.attrs.get("shapeId") == shape_id and R + "}id" in n.attrs]
            props = []
            for control in controls:
                rid = control.attrs[R + "}id"]
                props += [r["part"] for r in book.package.relationships(sheet["part"]) if r["Id"] == rid]
            require(len(set(props)) <= 1, "CHECKBOX_MAPPING", "复选框属性关系不唯一", label=label)
            if props:
                root = book.package.doc(props[0]).root
                require(root.ns == X14 and root.attrs.get("objectType") == "CheckBox" and not root.attrs.get("fmlaLink"),
                        "CHECKBOX_MAPPING", "复选框属性或关联不支持", label=label)
                state = root.attrs.get("checked", "Unchecked")
                require(state in ("Checked", "Unchecked") and (state == "Checked") == bool(checked and checked.text.strip() == "1"),
                        "CHECKBOX_INCONSISTENT", "VML 与控件属性状态不一致", label=label)
            result.append({"label": label, "checked": bool(checked and checked.text.strip() == "1"), "vml_part": rel["part"],
                           "shape_id": shape.attrs["id"], "props_part": props[0] if props else None})
    return result


def apply_checkbox_values(book, values):
    seen = set()
    for value in values:
        matches = [c for c in checkboxes(book, value["sheet"]) if c["label"] == value["label"]]
        require(len(matches) == 1, "CHECKBOX_MAPPING", "复选框标签不唯一或不存在", label=value["label"])
        item = matches[0]
        key = (item["vml_part"], item["shape_id"])
        require(key not in seen, "TARGET_CONFLICT", "复选框重复修改")
        seen.add(key)
        if item["checked"] == value["checked"]:
            continue
        doc = book.package.doc(item["vml_part"])
        shape = next(n for n in doc.root.descendants("shape", V) if n.attrs["id"] == item["shape_id"])
        client = shape.child("ClientData", X)
        checked = client.child("Checked", X)
        if value["checked"]:
            tag = client.qualify("Checked")
            fragment = b"<" + tag + b">1</" + tag + b">"
            edit = Edit(checked.start, checked.end, fragment, "control:checked", item["shape_id"]) if checked else append_children(doc.data, client, fragment, "control:checked", item["shape_id"])
        else:
            edit = Edit(checked.start, checked.end, b"", "control:checked", item["shape_id"])
        book.package.patch(item["vml_part"], [edit])
        if item["props_part"]:
            props = book.package.doc(item["props_part"])
            tag = set_attribute(props.data, props.root, b"checked", "Checked" if value["checked"] else "Unchecked")
            book.package.patch(item["props_part"], [Edit(props.root.start, props.root.open_end, tag, "control:checked", item["shape_id"])])
