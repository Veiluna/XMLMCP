from __future__ import annotations

import copy
import json
import posixpath
import zipfile
from xml.etree import ElementTree
from pathlib import Path

from .errors import XMLMCPError, require
from .namespaces import P, CT, REL_BASE
from .xmlbytes import XMLDocument, Edit, apply_edits, append_children, digest, xml_escape


def read_xml(data: bytes):
    # Read-only validation supports encodings in opaque XML parts without rewriting them.
    require(b"<!DOCTYPE" not in data.replace(b"\x00", b""), "XML_DTD", "不支持带 DTD 的部件")
    try:
        return ElementTree.fromstring(data)
    except ElementTree.ParseError as exc:
        raise XMLMCPError("INVALID_XML", "XML 解析失败", error=str(exc)) from exc


def absolute_file(path: str | Path, suffix: str | None = None) -> Path:
    value = Path(path)
    require(value.is_absolute(), "ABSOLUTE_PATH_REQUIRED", "文件必须使用绝对路径", path=str(value))
    value = value.resolve()
    require(value.is_file(), "FILE_NOT_FOUND", "文件不存在", path=str(value))
    if suffix:
        require(value.suffix.lower() == suffix, "FILE_TYPE", f"仅支持 {suffix} 文件")
    return value


def rels_path(part: str) -> str:
    directory, filename = posixpath.split(part)
    return posixpath.join(directory, "_rels", filename + ".rels") if part else "_rels/.rels"


def relationship_source(path: str) -> str:
    if path == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(path)
    require(posixpath.basename(directory) == "_rels" and filename.endswith(".rels"),
            "INVALID_RELATIONSHIP", "关系部件路径无效", part=path)
    return posixpath.join(posixpath.dirname(directory), filename[:-5])


def resolve_part(source: str, target: str) -> str:
    from urllib.parse import unquote
    target = unquote(target.split("#", 1)[0])
    require("\\" not in target, "INVALID_RELATIONSHIP", "关系目标不得包含反斜线")
    path = posixpath.normpath(target.lstrip("/")) if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    require(path not in ("", ".", "..") and not path.startswith("../"), "INVALID_RELATIONSHIP", "关系目标越出包根目录")
    return path


class Package:
    def __init__(self, path: str | Path):
        self.path = absolute_file(path, ".xlsx")
        self.fingerprint = digest(self.path.read_bytes())
        try:
            with zipfile.ZipFile(self.path) as archive:
                self.infos = {info.filename: copy.copy(info) for info in archive.infolist()}
                require(len(self.infos) == len(archive.infolist()), "DUPLICATE_PART", "ZIP 包含重名部件")
                self.comment = archive.comment
                self.parts = {name: archive.read(name) for name in self.infos}
        except (zipfile.BadZipFile, RuntimeError) as exc:
            raise XMLMCPError("INVALID_XLSX", "无法读取 XLSX ZIP 包；不支持加密文件", error=str(exc)) from exc
        require("[Content_Types].xml" in self.parts, "INVALID_XLSX", "缺少内容类型部件")
        require(not any(name.lower().startswith("_xmlsignatures/") for name in self.parts),
                "SIGNED_WORKBOOK", "不支持修改数字签名工作簿")
        for name in self.parts:
            require(not name.startswith("/") and "\\" not in name and ".." not in name.split("/"),
                    "INVALID_PART_NAME", "包部件路径无效", part=name)
        self.original = dict(self.parts)
        self.stages: list[dict] = []
        self.additions: dict[str, dict] = {}
        self.removals: dict[str, dict] = {}

    def doc(self, part: str) -> XMLDocument:
        require(part in self.parts, "MISSING_PART", "缺少 XML 部件", part=part)
        return XMLDocument(self.parts[part])

    def patch(self, part: str, edits: list[Edit]) -> None:
        old = self.parts[part]
        new, records = apply_edits(old, edits)
        if not records:
            return
        XMLDocument(new)
        self.parts[part] = new
        if part in self.original:
            self.stages.append({"part": part, "before_sha256": digest(old), "after_sha256": digest(new), "patches": records})
        else:
            self.additions[part]["sha256"] = digest(new)

    def add(self, part: str, data: bytes, reason: str) -> None:
        require(part not in self.parts, "PART_COLLISION", "新增部件名称冲突", part=part)
        self.parts[part] = data
        self.additions[part] = {"sha256": digest(data), "reason": reason}

    def remove(self, part: str, reason: str) -> None:
        if part in self.parts:
            require(part in self.original and not any(s["part"] == part for s in self.stages),
                    "PATCH_REMOVE_CONFLICT", "不能删除已修改部件", part=part)
            self.removals[part] = {"sha256": digest(self.parts.pop(part)), "reason": reason}

    def relationships(self, source: str) -> list[dict]:
        path = rels_path(source)
        if path not in self.parts:
            return []
        root = self.doc(path).root
        return [dict(n.attrs) | {"part": resolve_part(source, n.attrs["Target"]) if n.attrs.get("TargetMode") != "External" else None}
                for n in root.children if n.name == P + "}Relationship"]

    def related(self, source: str, kind: str) -> list[dict]:
        return [r for r in self.relationships(source) if r.get("Type", "").rsplit("/", 1)[-1] == kind]

    def add_relationship(self, source: str, kind: str, target_part: str) -> str:
        path = rels_path(source)
        target = posixpath.relpath(target_part, posixpath.dirname(source) or ".")
        ids = {r["Id"] for r in self.relationships(source)}
        index = 1
        while f"rId{index}" in ids:
            index += 1
        rid = f"rId{index}"
        if path not in self.parts:
            content = (f'<Relationships xmlns="{P}"><Relationship Id="{rid}" Type="{REL_BASE}{kind}" Target="{target}"/></Relationships>').encode()
            self.add(path, content, f"image:{kind}:relationship")
        else:
            doc = self.doc(path)
            name = doc.root.qualify("Relationship")
            content = b"<" + name + b' Id="' + rid.encode() + b'" Type="' + xml_escape(REL_BASE + kind, True) + b'" Target="' + xml_escape(target, True) + b'"/>'
            self.patch(path, [append_children(doc.data, doc.root, content, f"image:{kind}:relationship", rid)])
        return rid

    def add_content_type(self, part: str, content_type: str, extension: str | None = None) -> None:
        doc = self.doc("[Content_Types].xml")
        root = doc.root
        if extension:
            match = next((n for n in root.children if n.local == "Default" and n.attrs.get("Extension", "").lower() == extension), None)
            if match:
                require(match.attrs["ContentType"] == content_type, "CONTENT_TYPE_CONFLICT", "文件扩展名的内容类型冲突")
                return
            fragment = b"<" + root.qualify("Default") + b' Extension="' + xml_escape(extension, True) + b'" ContentType="' + xml_escape(content_type, True) + b'"/>'
        else:
            if any(n.local == "Override" and n.attrs.get("PartName") == "/" + part for n in root.children):
                return
            fragment = b"<" + root.qualify("Override") + b' PartName="/' + xml_escape(part, True) + b'" ContentType="' + xml_escape(content_type, True) + b'"/>'
        self.patch("[Content_Types].xml", [append_children(doc.data, root, fragment, "image:content_type", part)])

    def fresh_part(self, directory: str, prefix: str, extension: str) -> str:
        i = 1
        while f"{directory}/{prefix}{i}.{extension}" in self.parts:
            i += 1
        return f"{directory}/{prefix}{i}.{extension}"

    def write(self, path: str | Path) -> None:
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.comment = self.comment
            for name, data in self.parts.items():
                info = copy.copy(self.infos[name]) if name in self.infos else name
                archive.writestr(info, data)

    def validate(self) -> dict:
        from collections import Counter
        for part, data in self.parts.items():
            if part.endswith((".xml", ".rels")):
                read_xml(data)
        relation_count = 0
        for path in self.parts:
            if not path.endswith(".rels"):
                continue
            source = relationship_source(path)
            rels = self.relationships(source)
            require(len({r["Id"] for r in rels}) == len(rels), "DUPLICATE_RELATIONSHIP", "关系 ID 重复", part=path)
            for r in rels:
                if r["part"] is not None:
                    require(r["part"] in self.parts, "BROKEN_RELATIONSHIP", "关系引用的部件不存在", part=path, target=r["part"])
                relation_count += 1
            if source:
                require(source in self.parts, "ORPHAN_RELATIONSHIP", "关系的源部件不存在", part=path)
        types = self.doc("[Content_Types].xml").root
        defaults = {n.attrs["Extension"].lower(): n.attrs["ContentType"] for n in types.children if n.local == "Default"}
        overrides = {n.attrs["PartName"].lstrip("/"): n.attrs["ContentType"] for n in types.children if n.local == "Override"}
        require(all(name in self.parts for name in overrides), "BROKEN_CONTENT_TYPE", "内容类型指向不存在的部件")
        for name in self.parts:
            if name.endswith("/") or name == "[Content_Types].xml":
                continue
            require(name in overrides or name.rsplit(".", 1)[-1].lower() in defaults,
                    "MISSING_CONTENT_TYPE", "部件没有对应内容类型", part=name)
        return {"xml_valid": True, "relationships_valid": True, "content_types_valid": True, "relationship_count": relation_count}

    def manifest(self) -> dict:
        return {"schema_version": 1, "input_fingerprint": self.fingerprint,
                "original_parts": {n: digest(d) for n, d in self.original.items()},
                "output_parts": {n: digest(d) for n, d in self.parts.items()},
                "stages": self.stages, "added_parts": self.additions, "removed_parts": self.removals}

