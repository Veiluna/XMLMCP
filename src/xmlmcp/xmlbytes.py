"""Namespace-aware Expat nodes retaining byte offsets; only new fragments serialize."""
from __future__ import annotations

import base64
import hashlib
import re
from dataclasses import dataclass, field
from xml.parsers import expat
from xml.sax.saxutils import escape

from .errors import XMLMCPError, require


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def xml_escape(value: object, attribute: bool = False) -> bytes:
    text = str(value)
    require(not any(ord(c) < 32 and c not in "\t\n\r" for c in text),
            "INVALID_XML_TEXT", "文本包含 XML 不允许的控制字符")
    require(not any(0xD800 <= ord(c) <= 0xDFFF or ord(c) in (0xFFFE, 0xFFFF) for c in text),
            "INVALID_XML_TEXT", "文本包含 XML 不允许的字符")
    entities = {'"': "&quot;", "'": "&apos;"} if attribute else {}
    return escape(text, entities).encode("utf-8")


def tag_end(data: bytes, start: int) -> int:
    quote = 0
    for i in range(start, len(data)):
        char = data[i]
        if quote:
            if char == quote:
                quote = 0
        elif char in (34, 39):
            quote = char
        elif char == 62:
            return i + 1
    raise XMLMCPError("INVALID_XML", "XML 标签未闭合")


@dataclass
class Node:
    name: str
    raw_name: bytes
    attrs: dict[str, str]
    start: int
    open_end: int
    self_closing: bool
    parent: Node | None = field(default=None, repr=False)
    children: list[Node] = field(default_factory=list)
    chunks: list[str] = field(default_factory=list, repr=False)
    close_start: int = 0
    end: int = 0
    document: XMLDocument | None = field(default=None, repr=False)

    @property
    def local(self) -> str:
        return self.name.rsplit("}", 1)[-1]

    @property
    def ns(self) -> str:
        return self.name.rsplit("}", 1)[0] if "}" in self.name else ""

    @property
    def text(self) -> str:
        return "".join(self.chunks)

    def child(self, local: str, ns: str | None = None) -> Node | None:
        return next((n for n in self.children if n.local == local and (ns is None or n.ns == ns)), None)

    def descendants(self, local: str | None = None, ns: str | None = None):
        for n in self.children:
            if (local is None or n.local == local) and (ns is None or n.ns == ns):
                yield n
            yield from n.descendants(local, ns)

    def qualify(self, local: str) -> bytes:
        prefix = self.raw_name.rsplit(b":", 1)[0] + b":" if b":" in self.raw_name else b""
        return prefix + local.encode("ascii")


class XMLDocument:
    def __init__(self, data: bytes):
        self.data = data
        require(not data.startswith((b"\xff\xfe", b"\xfe\xff")), "XML_ENCODING", "修改部件仅支持 UTF-8 XML")
        declaration = re.match(br"\s*<\?xml[^?]*encoding\s*=\s*['\"]([^'\"]+)", data[:200])
        require(not declaration or declaration[1].lower() in (b"utf-8", b"utf8", b"us-ascii"),
                "XML_ENCODING", "修改部件仅支持 UTF-8 XML")
        self.root: Node | None = None
        stack: list[Node] = []
        parser = expat.ParserCreate(namespace_separator="}")

        def start(name, attrs):
            offset = parser.CurrentByteIndex
            end = tag_end(data, offset)
            raw = re.match(br"<([^\s/>]+)", data[offset:end])[1]
            node = Node(name, raw, dict(attrs), offset, end, data[offset:end].rstrip().endswith(b"/>"))
            node.document = self
            if stack:
                node.parent = stack[-1]
                stack[-1].children.append(node)
            else:
                self.root = node
            stack.append(node)

        def finish(name):
            node = stack.pop()
            node.close_start = node.open_end if node.self_closing else parser.CurrentByteIndex
            node.end = node.open_end if node.self_closing else tag_end(data, node.close_start)

        def doctype(*args):
            raise XMLMCPError("XML_DTD", "不支持带 DTD 的部件")

        parser.StartElementHandler = start
        parser.EndElementHandler = finish
        parser.CharacterDataHandler = lambda text: stack[-1].chunks.append(text) if stack else None
        parser.StartDoctypeDeclHandler = doctype
        parser.ExternalEntityRefHandler = lambda *args: 0
        try:
            parser.Parse(data, True)
        except expat.ExpatError as exc:
            raise XMLMCPError("INVALID_XML", "XML 解析失败", error=str(exc)) from exc
        require(self.root is not None, "INVALID_XML", "空 XML 部件")

    def raw(self, node: Node) -> bytes:
        return self.data[node.start:node.end]


@dataclass
class Edit:
    start: int
    end: int
    replacement: bytes
    reason: str
    selector: str


def apply_edits(data: bytes, edits: list[Edit]) -> tuple[bytes, list[dict]]:
    edits = sorted(edits, key=lambda e: (e.start, e.end))
    cursor = 0
    output: list[bytes] = []
    manifest: list[dict] = []
    last_insert = -1
    for edit in edits:
        require(0 <= edit.start <= edit.end <= len(data) and edit.start >= cursor,
                "OVERLAPPING_PATCH", "XML 修改片段重叠或越界", selector=edit.selector)
        require(not (edit.start == edit.end == last_insert), "OVERLAPPING_PATCH", "同位置插入必须先合并")
        old = data[edit.start:edit.end]
        if old == edit.replacement:
            continue
        output.extend((data[cursor:edit.start], edit.replacement))
        manifest.append({"start": edit.start, "end": edit.end, "old_sha256": digest(old),
                         "replacement_base64": base64.b64encode(edit.replacement).decode("ascii"),
                         "reason": edit.reason, "selector": edit.selector})
        cursor = edit.end
        last_insert = edit.start if edit.start == edit.end else -1
    output.append(data[cursor:])
    return b"".join(output), manifest


def set_attribute(data: bytes, node: Node, name: bytes, value: str | None) -> bytes:
    """Return the original opening tag with exactly one attribute edited."""
    tag = data[node.start:node.open_end]
    regex = re.compile(br"(?P<space>\s+)(?P<name>[^\s=/>]+)(?P<eq>\s*=\s*)(?P<quote>['\"])(?P<value>.*?)(?P=quote)", re.S)
    found = next((m for m in regex.finditer(tag) if m["name"] == name), None)
    if found:
        if value is None:
            return tag[:found.start()] + tag[found.end():]
        val = xml_escape(value, True)
        return tag[:found.start("value")] + val + tag[found.end("value"):]
    if value is None:
        return tag
    offset = len(tag) - (2 if node.self_closing else 1)
    return tag[:offset] + b" " + name + b'="' + xml_escape(value, True) + b'"' + tag[offset:]


def edit_attributes(raw: bytes, values: dict[bytes, str | None]) -> bytes:
    for name, value in values.items():
        doc = XMLDocument(raw)
        tag = set_attribute(raw, doc.root, name, value)
        raw = tag + raw[doc.root.open_end:]
    return raw


def append_children(data: bytes, parent: Node, fragments: bytes, reason: str, selector: str) -> Edit:
    if parent.self_closing:
        opening = data[parent.start:parent.open_end]
        new = opening[:-2] + b">" + fragments + b"</" + parent.raw_name + b">"
        return Edit(parent.start, parent.end, new, reason, selector)
    return Edit(parent.close_start, parent.close_start, fragments, reason, selector)


def standalone(doc: XMLDocument, node: Node) -> bytes:
    """Carry inherited xmlns declarations when harvesting a fragment from Excel."""
    declarations: dict[bytes, bytes] = {}
    chain: list[Node] = []
    current = node
    while current:
        chain.append(current)
        current = current.parent
    for ancestor in reversed(chain):
        opening = doc.data[ancestor.start:ancestor.open_end]
        for match in re.finditer(br"\s+(xmlns(?::[^\s=]+)?)\s*=\s*(['\"])(.*?)\2", opening):
            declarations[match[1]] = match[3]
    raw = doc.raw(node)
    root_open = raw[:node.open_end-node.start]
    own = {m[1] for m in re.finditer(br"\s+(xmlns(?::[^\s=]+)?)\s*=", root_open)}
    additions = b"".join(b" " + k + b'="' + v + b'"' for k, v in declarations.items() if k not in own)
    cut = len(root_open) - (2 if node.self_closing else 1)
    return raw[:cut] + additions + raw[cut:]

