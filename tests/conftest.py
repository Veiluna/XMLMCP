from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from PIL import Image

from xmlmcp.namespaces import S, R, P, CT, XDR, A, C, REL_BASE


def write_template(path: Path, *, formula: bool = False, chart: bool = False, image: bool = False,
                   mixed: bool = False, prefix: bool = False, merged: bool = False):
    root_ns = f'xmlns:s="{S}"' if prefix else f'xmlns="{S}"'
    pre = "s:" if prefix else ""
    formula_cell = f'<{pre}c r="D5"><{pre}f>SUM(C5:C6)</{pre}f><{pre}v>30</{pre}v></{pre}c>' if formula else ""
    merge = f'<{pre}mergeCells count="1"><{pre}mergeCell ref="F2:G3"/></{pre}mergeCells>' if merged else ""
    drawing = f'<{pre}drawing r:id="rIdDrawing"/>' if image or chart else ""
    sheet = (f'<?xml version="1.0" encoding="UTF-8"?><{pre}worksheet {root_ns} xmlns:r="{R}">'
             f'<{pre}dimension ref="A1:H20"/><{pre}sheetViews><{pre}sheetView workbookViewId="0"/></{pre}sheetViews>'
             f'<{pre}sheetFormatPr defaultRowHeight="15"/><{pre}cols><{pre}col min="1" max="8" width="12" customWidth="1"/></{pre}cols>'
             f'<{pre}sheetData><!-- untouched whitespace -->\n'
             f'<{pre}row r="1"><{pre}c r="A1" s="1" t="s"><{pre}v>0</{pre}v></{pre}c></{pre}row>\n'
             f'<{pre}row r="2"><{pre}c r="A2" t="s"><{pre}v>1</{pre}v></{pre}c></{pre}row>\n'
             f'<{pre}row r="3"><{pre}c r="A3" t="s"><{pre}v>2</{pre}v></{pre}c></{pre}row>\n'
             f'<{pre}row r="5"><{pre}c r="B5" t="inlineStr"><{pre}is><{pre}t>甲</{pre}t></{pre}is></{pre}c><{pre}c r="C5" s="1"><{pre}v>10</{pre}v></{pre}c>{formula_cell}</{pre}row>\n'
             f'<{pre}row r="6"><{pre}c r="B6" t="inlineStr"><{pre}is><{pre}t>乙</{pre}t></{pre}is></{pre}c><{pre}c r="C6" s="1"><{pre}v>20</{pre}v></{pre}c></{pre}row>\n'
             f'<{pre}row r="20"/></{pre}sheetData>{merge}{drawing}'
             f'<{pre}extLst><{pre}ext uri="preserve-test"><opaque xmlns="urn:xmlmcp:test" a="002"/></{pre}ext></{pre}extLst></{pre}worksheet>')
    rich2 = '<r><rPr><sz val="8"/><rFont val="Arial"/></rPr><t>size</t></r>' if mixed else ''
    parts = {
        "_rels/.rels": f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{REL_BASE}officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView/></bookViews><sheets><sheet name="报告" sheetId="7" r:id="rIdSheet"/></sheets><definedNames><definedName name="Field">\'报告\'!$A$1</definedName><definedName name="Field" localSheetId="0">\'报告\'!$A$3</definedName><definedName name="SmallFont">\'报告\'!$A$2</definedName><definedName name="Grid">\'报告\'!$B$5:$C$8</definedName><definedName name="Photo">\'报告\'!$E$5:$G$9</definedName><definedName name="Dynamic">OFFSET(\'报告\'!$C$5,0,0,2,1)</definedName></definedNames><calcPr calcId="124519"/></workbook>',
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{P}"><Relationship Id="rIdSheet" Type="{REL_BASE}worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rIdStyles" Type="{REL_BASE}styles" Target="styles.xml"/><Relationship Id="rIdStrings" Type="{REL_BASE}sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": sheet,
        "xl/styles.xml": f'<styleSheet xmlns="{S}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="9"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',
        "xl/sharedStrings.xml": f'<sst xmlns="{S}" count="3" uniqueCount="3"><si><t>Old title</t></si><si><r><rPr><sz val="7"/><rFont val="Arial"/></rPr><t>Small English</t></r>{rich2}</si><si><t>Local title</t></si></sst>',
        "customXml/unknown.xml": '<?xml version="1.0" encoding="UTF-8"?><x:opaque xmlns:x="urn:unknown" b="2" a="1">\r\n  KEEP &amp; KEEP\r\n</x:opaque>',
        "xl/ctrlProps/ctrlProp1.xml": '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="Checked"/>',
        "xl/drawings/vmlDrawing1.vml": '<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="keep"><v:textbox>KEEP QR CAPTION</v:textbox></v:shape></xml>',
        "xl/comments/comment1.xml": f'<comments xmlns="{S}"><authors><author>Tester</author></authors><commentList><comment ref="H2" authorId="0"><text><t>KEEP COMMENT</t></text></comment></commentList></comments>',
        "customXml/opaque.bin": b"\x00\x80\xffUnknown binary preserved\x00",
    }
    drawing_rels = []
    anchors = []
    if image:
        media = path.with_name("original.png")
        Image.new("RGB", (80, 40), "navy").save(media)
        parts["xl/media/image1.png"] = media.read_bytes()
        drawing_rels.append(f'<Relationship Id="rIdImage" Type="{REL_BASE}image" Target="../media/image1.png"/>')
        for shape_id, col in ((1, 4), (2, 6)):
            anchors.append(f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{col+1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{shape_id}" name="Picture {shape_id}" descr="preserve this"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/><a:srcRect l="5000" r="5000"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="800000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>')
    if chart:
        drawing_rels.append(f'<Relationship Id="rIdChart" Type="{REL_BASE}chart" Target="../charts/chart1.xml"/>')
        anchors.append(f'<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></xdr:xfrm><a:graphic><a:graphicData uri="{C}"><c:chart xmlns:c="{C}" r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')
        parts["xl/charts/chart1.xml"] = f'<c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>\'报告\'!$B$5:$B$6</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>甲</c:v></c:pt><c:pt idx="1"><c:v>乙</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>\'报告\'!$C$5:$C$6</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="111"/><c:axId val="222"/></c:barChart><c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="111"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
    if anchors:
        parts["xl/worksheets/_rels/sheet1.xml.rels"] = f'<Relationships xmlns="{P}"><Relationship Id="rIdDrawing" Type="{REL_BASE}drawing" Target="../drawings/drawing1.xml"/></Relationships>'
        parts["xl/drawings/drawing1.xml"] = f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{R}">' + ''.join(anchors) + '</xdr:wsDr>'
        parts["xl/drawings/_rels/drawing1.xml.rels"] = f'<Relationships xmlns="{P}">' + ''.join(drawing_rels) + '</Relationships>'
    content_types = {"xl/workbook.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                     "xl/worksheets/sheet1.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                     "xl/styles.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                     "xl/sharedStrings.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                     "xl/comments/comment1.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"}
    if anchors:
        content_types["xl/drawings/drawing1.xml"] = "application/vnd.openxmlformats-officedocument.drawing+xml"
    if chart:
        content_types["xl/charts/chart1.xml"] = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    parts["[Content_Types].xml"] = f'<Types xmlns="{CT}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="png" ContentType="image/png"/>' + ''.join(f'<Override PartName="/{p}" ContentType="{t}"/>' for p, t in content_types.items()) + '</Types>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data.encode("utf-8") if isinstance(data, str) else data)
    return path


@pytest.fixture
def template(tmp_path):
    return write_template(tmp_path / "template.xlsx")


class GeometryCalculator:
    def run(self, book, directory, calculate, images):
        assert not calculate
        return {"calculated": False, "external_refresh": False, "engine": "test geometry",
                "geometry": {image.geometry_key: {"width": 100, "height": 100} for image in images}}, None

