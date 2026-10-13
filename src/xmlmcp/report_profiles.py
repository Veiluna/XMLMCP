from __future__ import annotations
import json
from .errors import require
from .models import Target
from .report_models import ReportProfile, OrderLayout, TableSlot, CheckboxBinding, TextRunBinding
from .package import absolute_file
from .workbook import cell_address, cell_position, Rectangle

FIELD_CATALOG = {
    "client_company": "委托单位", "client_contact": "委托人", "client_address": "委托方地址",
    "client_phone": "联系电话", "client_email": "邮箱", "background": "测试背景描述",
    "sample_order_no": "样品单号", "received_date": "收样日期", "test_date": "检测日期",
    "report_date": "报告日期", "operator": "操作人员", "instrument_model": "设备型号",
    "temperature_range": "实测温度范围（℃）", "humidity_range": "实测湿度范围（%RH）",
    "test_parameters": "实际测试参数", "single_analysis": "单点描述性分析", "pressure_analysis": "变压描述性分析",
}
SAMPLE_CATALOG = {"sample_name": "样品名称", "sample_number": "样品编号", "material": "材料种类",
                  "thickness_um": "极片厚度", "areal_density": "极片面密度", "foil_thickness_um": "箔材厚度",
                  "foil_areal_density": "箔材面密度"}
SOURCE_LABELS = {"client_company": ["送检单位名称", "委托单位"], "client_contact": ["联系人", "委托人"],
                 "client_address": ["送检单位地址", "委托方地址"], "client_phone": ["联系电话"],
                 "client_email": ["邮箱"], "background": ["测试背景描述"]}
SAMPLE_LABELS = {"sample_name": "样品名称", "sample_number": "样品编号", "material": "材料种类",
                 "thickness_um": "极片厚度", "areal_density": "极片面密度", "foil_thickness_um": "箔材厚度",
                 "foil_areal_density": "箔材面密度"}


def label(value):
    return str(value).strip().lstrip("*").rstrip("：:").split("（")[0].split("(")[0].strip()


def discover_order_layout(book) -> OrderLayout:
    matches = []
    for sheet in book.sheets:
        if sheet["kind"] != "worksheet":
            continue
        cells = book.cell_map(sheet)
        labels = {}
        for address, cell in cells.items():
            if cell_position(address)[1] == 1 and isinstance(book.value(cell), str):
                key = label(book.value(cell))
                labels.setdefault(key, []).append(address)
        if "样品名称" not in labels or "材料种类" not in labels:
            continue
        def address_of(aliases):
            addresses = [a for alias in aliases for a in labels.get(alias, [])]
            require(len(addresses) == 1, "ORDER_LAYOUT", "委托单字段缺失或不唯一，请配置 order 映射", aliases=aliases)
            return addresses[0]
        fields = {key: Target(sheet=sheet["name"], range="B" + str(cell_position(address_of(aliases))[0]))
                  for key, aliases in SOURCE_LABELS.items()}
        name_row = cell_position(address_of(["样品名称"]))[0]
        # Bound columns by the sample-number header, rather than source instruction text on the right.
        number_row = cell_position(address_of(["样品编号"]))[0]
        boundary = next((col for col in range(2, 100) if label(book.value(cells.get(cell_address(number_row, col)))) in ("注意事项", "备注")), None)
        require(boundary is not None and boundary > 2, "ORDER_LAYOUT", "无法确定样品列边界，请配置 sample_rows")
        rows = {key: Target(sheet=sheet["name"], range=f"B{cell_position(address_of([alias]))[0]}:{cell_address(cell_position(address_of([alias]))[0], boundary - 1)}")
                for key, alias in SAMPLE_LABELS.items()}
        matches.append(OrderLayout(fields=fields, sample_rows=rows, sample_numbers=rows["sample_number"], checkbox_sheet=sheet["name"]))
    require(len(matches) == 1, "ORDER_LAYOUT", "未找到唯一的极片委托单样品表，请提供 order 映射")
    return matches[0]


def known_report_profile(template, order_layout) -> ReportProfile | None:
    required = {"委托单位", "委托人", "委托方地址", "型号规格", "样品单号", "收样日期", "检测日期", "基本信息委托人",
                "联系电话", "测试背景描述", "样品信息表", "单点结果区", "单点分析", "变压结果区", "变压分析"}
    if not required <= {n["name"] for n in template.names if n["scope"] == "workbook"}:
        return None
    def named(name):
        return Target(name=name, scope="workbook")
    sheet, grid = template.resolve(named("样品信息表"))
    single_sheet, single = template.resolve(named("单点结果区"))
    pressure_sheet, pressure = template.resolve(named("变压结果区"))
    if sheet["name"] != single_sheet["name"] or sheet["name"] != pressure_sheet["name"] or grid.size != (10, 11) or single.size != (52, 11) or pressure.size != (35, 11):
        return None
    def at(row, col=1):
        return Target(sheet=sheet["name"], range=cell_address(row, col))
    def area(row1, col1, row2, col2):
        return Target(sheet=sheet["name"], range=Rectangle(row1, col1, row2, col2).address)
    fields = {
        "client_company": [named("委托单位")], "client_contact": [named("委托人"), named("基本信息委托人")],
        "client_address": [named("委托方地址")], "client_phone": [named("联系电话")], "background": [named("测试背景描述")],
        "instrument_model": [named("型号规格"), at(grid.bottom + 3)], "sample_order_no": [named("样品单号")],
        "received_date": [named("收样日期"), at(grid.bottom + 3, 8)], "test_date": [named("检测日期"), at(grid.bottom + 3, 9)],
        "report_date": [at(grid.bottom + 3, 10)], "operator": [at(grid.bottom + 3, 11)],
        "temperature_range": [at(grid.bottom + 3, 6)], "humidity_range": [at(grid.bottom + 3, 7)],
        "single_analysis": [named("单点分析")], "pressure_analysis": [named("变压分析")],
    }
    samples = {key: area(grid.top + i, 2, grid.top + i, 9) for i, key in enumerate(SAMPLE_CATALOG)}
    cells = template.cell_map(sheet)
    if [label(template.value(cells.get(cell_address(grid.top + i, 1)))) for i in range(7)] != ["样品编号", "样品名称", "材料种类", "极片厚度", "极片面密度", "箔材厚度", "箔材面密度"]:
        return None
    # SAMPLE_CATALOG has name first; map the actual numbered template rows explicitly.
    row_order = ["sample_number", "sample_name", "material", "thickness_um", "areal_density", "foil_thickness_um", "foil_areal_density"]
    samples = {key: area(grid.top + i, 2, grid.top + i, 9) for i, key in enumerate(row_order)}
    model = template.value(cells.get(cell_address(grid.bottom + 3, 1)))
    profile = ReportProfile(profile_id="electrode-marked-template-v1", template_fingerprint=template.package.fingerprint,
                         order=order_layout, fields=fields, sample_rows=samples,
                         tables={"single_summary": TableSlot(target=area(single.top + 22, 1, single.top + 28, 8)),
                                 "pressure_resistivity": TableSlot(target=area(pressure.top + 20, 1, pressure.bottom, 6))},
                         figures={"single_point": area(single.top + 2, 1, single.top + 20, 11), "pressure": area(pressure.top + 2, 1, pressure.top + 17, 11)},
                         checkboxes=[CheckboxBinding(source_label=l, target_label=l, target_sheet=sheet["name"]) for l in
                                     ("包装完好", "全部退回", "无需退回&报损", "其他要求", "腐蚀性", "毒性", "健康危害", "无健康危害", "需要前处理", "易吸潮", "寄回", "无需寄回")],
                         text_runs=[TextRunBinding(field="test_parameters", target=at(grid.bottom + 3, 2), run_index=3, prefix="\n")],
                         defaults={"instrument_model": model} if model else {})
    from .report_plots import REFERENCE_PLOTS
    replacements = {i['name'].removeprefix('XMLMCP Reference '):i['image_id'] for i in template.images()
                    if i['name'].startswith('XMLMCP Reference ')}
    if set(replacements) == set(REFERENCE_PLOTS):
        profile.profile_id = 'electrode-editable-template-v2'
        profile.figures = {}
        profile.figure_replacements = replacements
        profile.figure_style = 'template-statistics'
        profile.tables = {'single_measurements':TableSlot(target=area(single.top+32,1,single.top+45,11)),
                          'pressure_measurements':TableSlot(target=area(pressure.top+17,1,pressure.top+30,11))}
    return profile


def load_profile(path: str) -> ReportProfile:
    try:
        return ReportProfile.model_validate_json(absolute_file(path, ".json").read_text(encoding="utf-8-sig"))
    except (ValueError, UnicodeError) as exc:
        from .errors import XMLMCPError
        raise XMLMCPError("REPORT_PROFILE", "报告映射 JSON 无效", error=str(exc)) from exc
