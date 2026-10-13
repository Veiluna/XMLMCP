"""Synchronize reviewed checkbox selections in VML and ctrlProps only."""
from .errors import require
from .namespaces import R
from .xmlbytes import Edit, append_children, set_attribute

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
