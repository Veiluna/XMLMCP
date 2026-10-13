from .errors import require
from .models import Target
from .namespaces import S
from .strings import encode_xstring
from .xmlbytes import XMLDocument, Edit, apply_edits, standalone, set_attribute, xml_escape


def apply_text_run_values(book, values):
    seen = set()
    for item in values:
        sheet, rect = book.resolve(Target.model_validate(item["target"]))
        rect = book.merged_anchor(sheet, rect, True)
        require(rect.size == (1, 1), "REPORT_RICH_RUN", "字符片段目标必须是单格")
        key = (sheet["name"], rect.address)
        require(key not in seen, "TARGET_CONFLICT", "同一富文本格只允许一个已确认字符片段")
        seen.add(key)
        doc = book.sheet_doc(sheet)
        cell = book.cell_map(sheet).get(rect.address)
        source = book.string_node(cell)
        require(source and not cell.child("f", S) and not any(n.local in ("rPh", "phoneticPr") for n in source.children),
                "REPORT_RICH_RUN", "目标不是可编辑的普通富文本")
        # Local inline copy isolates this cell from every other shared-string user.
        copied = XMLDocument(standalone(source.document, source))
        runs = [n for n in copied.root.children if n.name == S + "}r"]
        require(item["run_index"] < len(runs), "REPORT_RICH_RUN", "已确认的字符片段索引不存在")
        node = runs[item["run_index"]].child("t", S)
        require(node, "REPORT_RICH_RUN", "字符片段缺少 t 节点")
        value = item["value"]
        require(isinstance(value, str) and len(value.encode("utf-16-le")) // 2 <= 32767,
                "REPORT_RICH_RUN", "字符片段值不是可写文字")
        opening = set_attribute(copied.data, node, b"xml:space", "preserve")
        fragment = opening[:-2] + b">" if node.self_closing else opening
        fragment += xml_escape(encode_xstring(value)) + b"</" + node.raw_name + b">"
        payload = apply_edits(copied.data, [Edit(node.start, node.end, fragment, "cell:rich_run", rect.address)])[0]
        root = XMLDocument(payload).root
        new_name = root.qualify("is")
        payload = payload[:root.start + 1] + new_name + payload[root.start + 1 + len(root.raw_name):root.close_start] + b"</" + new_name + b">" + payload[root.end:]
        old_value = cell.child("is", S) or cell.child("v", S)
        require(old_value, "REPORT_RICH_RUN", "目标文字存储结构无效")
        edits = [Edit(old_value.start, old_value.end, payload, "cell:rich_run", rect.address),
                 Edit(cell.start, cell.open_end, set_attribute(doc.data, cell, b"t", "inlineStr"), "cell:rich_run_type", rect.address)]
        book.package.patch(sheet["part"], edits)
