"""Construct a no-calculation example from the already verified native template."""
from pathlib import Path

from xmlmcp.namespaces import C, S, P, CT
from xmlmcp.package import Package, rels_path
from xmlmcp.workbook import Workbook
from xmlmcp.xmlbytes import Edit, digest


def main():
    root = Path(__file__).resolve().parents[1]
    path = root / "examples" / "template.xlsx"
    package = Package(path)
    book = Workbook(package)
    removed = set()
    for sheet in book.sheets:
        doc = book.sheet_doc(sheet)
        formulas = list(doc.root.descendants("f", S))
        package.patch(sheet["part"], [Edit(n.start, n.end, b"", "cell:example_constant", n.parent.attrs["r"]) for n in formulas])
    for sheet, part, doc, relations in book.drawings():
        anchors = [n for n in doc.root.children if any(n.descendants("chart", C))]
        package.patch(part, [Edit(n.start, n.end, b"", "chart:example_remove", sheet["name"]) for n in anchors])
        chart_relations = [r for r in relations.values() if r["Type"].rsplit("/", 1)[-1] == "chart"]
        if chart_relations:
            rel_doc = package.doc(rels_path(part))
            ids = {r["Id"] for r in chart_relations}
            package.patch(rels_path(part), [Edit(n.start, n.end, b"", "chart:example_remove", n.attrs["Id"])
                                          for n in rel_doc.root.children if n.attrs.get("Id") in ids])
            for relation in chart_relations:
                removed.add(relation["part"])
                if rels_path(relation["part"]) in package.parts:
                    removed.add(rels_path(relation["part"]))
    chains = package.related(book.part, "calcChain")
    if chains:
        doc = package.doc(rels_path(book.part))
        ids = {c["Id"] for c in chains}
        package.patch(rels_path(book.part), [Edit(n.start, n.end, b"", "calculation:example_remove", n.attrs["Id"])
                                           for n in doc.root.children if n.attrs.get("Id") in ids])
        removed.update(c["part"] for c in chains)
    for part in removed:
        package.parts.pop(part)
    doc = package.doc("[Content_Types].xml")
    package.patch("[Content_Types].xml", [Edit(n.start, n.end, b"", "chart:example_remove", n.attrs["PartName"])
                                         for n in doc.root.children if n.name == CT + "}Override" and n.attrs["PartName"].lstrip("/") in removed])
    package.validate()
    assert not Workbook(package).charts()
    assert all(not cell.child("f", S) for sheet in book.sheets for cell in Workbook(package).cell_map(sheet).values())
    output = root / "examples" / "core-template.xlsx"
    if output.exists():
        raise FileExistsError("Will not overwrite an existing core example")
    package.write(output)
    print(f"Created {output}; source SHA-256: {digest(path.read_bytes())}")


if __name__ == "__main__":
    main()
