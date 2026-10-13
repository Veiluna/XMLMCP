"""Order + instrument ZIP -> mapped, byte-preserving report, independent of Office."""
from __future__ import annotations

import json
import tempfile
from pathlib import Path

from .errors import require
from .models import Target, Scalar
from .report_models import ReportProfile, ReportInspection, ReportResult
from .report_profiles import FIELD_CATALOG, SAMPLE_CATALOG, discover_order_layout, known_report_profile, load_profile
from .report_data import read_raw_zip, summarize
from .package import Package, absolute_file
from .workbook import Workbook, cell_address, Rectangle
from .controls import checkboxes
from .xmlbytes import digest
from .offline import preflight_offline, image_geometry
from .images import ImageChange
from .models import AddImage, ReplaceImage
from .engine import apply_changes, OPERATIONS
from .cells import preflight_cells
from .report_plots import make_plot, plot_mode, REFERENCE_PLOTS
from .publication import publish_bundle
from .namespaces import S
from .report_objects import drawing_objects, apply_example_removals


def read_target(book, target, row=False):
    sheet, rect = book.resolve(target)
    if not row:
        rect = book.merged_anchor(sheet, rect, True)
        require(rect.size == (1, 1), "ORDER_LAYOUT", "委托单标量来源必须为单格或合并单格")
    else:
        require(rect.size[0] == 1, "ORDER_LAYOUT", "样品来源必须为一行")
    cells = book.cell_map(sheet)
    require(all(cells.get(a) is None or cells[a].child("f", S) is None for a in rect.cells()),
            "ORDER_FORMULA_INPUT", "委托单来源含公式，不能在离线报告中依赖未经核验的缓存值")
    values = [book.value(cells.get(a)) for a in rect.cells()]
    return values if row else values[0]


def _range_text(values):
    low, high = min(values), max(values)
    return f"{low:g}" if low == high else f"{low:g}～{high:g}"


def _prepare(template_path, order_path, raw_zip_path, profile_path=None, metadata=None, inline_profile=None):
    template = Workbook(Package(template_path))
    template.package.validate()
    order = Workbook(Package(order_path))
    order.package.validate()
    raw = read_raw_zip(raw_zip_path)
    require(not (profile_path and inline_profile is not None), "REPORT_PROFILE", "profile_path 和 profile 不能同时指定")
    profile = load_profile(profile_path) if profile_path else ReportProfile.model_validate(inline_profile) if inline_profile is not None else None
    layout = profile.order if profile else discover_order_layout(order)
    if not profile:
        profile = known_report_profile(template, layout)
    if profile:
        require(profile.template_fingerprint == template.package.fingerprint, "REPORT_TEMPLATE_CHANGED",
                "模板与已确认映射的指纹不同，请重新检查和确认映射")
        require(set(profile.fields) <= set(FIELD_CATALOG) and set(profile.sample_rows) <= set(SAMPLE_CATALOG),
                "REPORT_FIELD", "映射包含不支持的字段名称")
        require(set(profile.figures) | set(profile.figure_replacements) <= {'single_point', 'pressure', *REFERENCE_PLOTS}
                and set(profile.tables) <= {'single_summary', 'pressure_resistivity', 'single_measurements', 'pressure_measurements'},
                "REPORT_SLOT", "映射包含不支持的图表或表格类型")
        require(not set(profile.figures) & set(profile.figure_replacements), 'REPORT_SLOT', '同一图表不能同时新增和替换')
        require(set(profile.defaults) <= set(FIELD_CATALOG), "REPORT_FIELD", "默认值字段名称不支持")
        require(all(binding.field in FIELD_CATALOG for binding in profile.text_runs), "REPORT_FIELD", "富文本片段字段名称不支持")
    require(set(layout.fields) <= set(FIELD_CATALOG) and set(layout.sample_rows) <= set(SAMPLE_CATALOG), "ORDER_LAYOUT", "委托单映射字段名称不支持")
    require("sample_name" in layout.sample_rows, "ORDER_LAYOUT", "样品名称来源未配置")
    fields = {key: read_target(order, target) for key, target in layout.fields.items()}
    if isinstance(fields.get("client_phone"), (int, float)):
        phone = fields["client_phone"]
        require(float(phone).is_integer(), "ORDER_PHONE", "联系电话包含非整数数值")
        fields["client_phone"] = str(int(phone))
    rows = {key: read_target(order, target, True) for key, target in layout.sample_rows.items()}
    require(all(len(v) == len(rows["sample_name"]) for v in rows.values()), "ORDER_LAYOUT", "样品各行长度不一致")
    examples = read_target(order, layout.sample_numbers, True) if layout.sample_numbers else [None] * len(rows["sample_name"])
    require(len(examples) == len(rows["sample_name"]), "ORDER_LAYOUT", "样品编号行长度不一致")
    samples, skipped = [], []
    for i, name in enumerate(rows["sample_name"]):
        if isinstance(examples[i], str) and layout.example_marker and layout.example_marker in examples[i]:
            skipped.append(i)
            continue
        if name is None or name == "":
            require(all(v[i] in (None, "", "/") for v in rows.values()), "ORDER_SAMPLE", "无样品名的列仍有属性，不自动丢弃")
            continue
        require(isinstance(name, str) and name.strip() == name, "ORDER_SAMPLE", "样品名需为不含首尾空格的文字标识")
        require(not any(s["sample_name"] == name for s in samples), "ORDER_SAMPLE", "样品名称重复", sample=name)
        samples.append({key: values[i] for key, values in rows.items()})
    require(samples, "ORDER_SAMPLE", "委托单没有实际样品")
    source_names = {s["sample_name"] for s in samples}
    raw_names = {f["sample"] for f in raw["files"]}
    require(source_names == raw_names, "SAMPLE_MISMATCH", "委托单与 ZIP 样品名不能一一对应；不使用文件顺序猜测", order_only=sorted(source_names - raw_names), raw_only=sorted(raw_names - source_names))
    files = {(f["sample"], f["kind"]): f for f in raw["files"]}
    stats = [summarize(f) for f in raw["files"]]
    warnings = list(raw["warnings"])
    if skipped:
        warnings.append({"code": "ORDER_EXAMPLE_COLUMNS_SKIPPED", "column_offsets": skipped})
    for item, summary in zip(raw["files"], stats):
        exported = item["instrument_summary"].get("COV%", {}).get("resistivity_ohm_cm")
        computed = summary["sample_cov_pct"]["resistivity_ohm_cm"]
        if exported is not None and (computed is None or abs(exported - computed) > .02):
            warnings.append({"code": "INSTRUMENT_COV_DIFFERENCE", "sample": item["sample"], "file": item["filename"],
                             "instrument_cov_pct": exported, "computed_sample_cov_pct": computed, "method": "sample standard deviation / abs(mean) * 100; n-1"})
    dates = sorted({f["metadata"][key].split(" ")[0] for f in raw["files"] for key in ("开始时间", "结束时间")})
    fields["test_date"] = dates[0] if len(dates) == 1 else dates[0] + "～" + dates[-1]
    fields["temperature_range"] = _range_text([float(f["metadata"]["温度(℃)"]) for f in raw["files"]])
    fields["humidity_range"] = _range_text([float(f["metadata"]["湿度(%RH)"]) for f in raw["files"]])
    methods = []
    singles = [f for f in raw["files"] if f["kind"] == "single_point"]
    pressure = [f for f in raw["files"] if f["kind"] == "pressure"]
    if singles:
        settings = sorted({(f["metadata"].get("压强(MPa)"), len(f["records"]), f["metadata"].get("保压时间(s)")) for f in singles})
        require(all(p is not None and hold is not None for p, _, hold in settings), "RAW_METADATA", "单点 CSV 缺少压强或保压时间")
        methods.append("单点：" + "；".join(f"设定{p}MPa，{n}点/样品，保压{hold}s" for p, n, hold in settings))
        summaries = [s for s in stats if s["kind"] == "single_point"]
        means = [s["mean"]["resistivity_ohm_cm"] for s in summaries]
        covs = [s["sample_cov_pct"]["resistivity_ohm_cm"] for s in summaries if s["sample_cov_pct"]["resistivity_ohm_cm"] is not None]
        fields["single_analysis"] = f"分析：{len(singles)}个样品的单点电阻率均值为{min(means):.4f}～{max(means):.4f} Ω·cm。"
        if covs:
            fields["single_analysis"] += f"样本COV为{min(covs):.2f}%～{max(covs):.2f}%（按逐点数据、n−1重算）。"
    if pressure:
        settings = sorted({tuple(f["settings"][k] for k in ("压强下限(MPa)", "压强上限(MPa)", "压强步进(MPa)", "保压时间")) for f in pressure})
        methods.append("变压：" + "；".join(f"{lo:g}～{hi:g}MPa，步进{step:g}MPa，保压{hold:g}s" for lo, hi, step, hold in settings))
        first = [f["records"][0]["resistivity_ohm_cm"] for f in pressure]
        last = [f["records"][-1]["resistivity_ohm_cm"] for f in pressure]
        monotone = sum(all(b["resistivity_ohm_cm"] <= a["resistivity_ohm_cm"] for a, b in zip(f["records"], f["records"][1:])) for f in pressure)
        fields["pressure_analysis"] = f"分析：{monotone}/{len(pressure)}个样品的电阻率随本次实测压强逐点下降。首工步电阻率为{min(first):.4f}～{max(first):.4f}，末工步为{min(last):.4f}～{max(last):.4f} Ω·cm。"
    methods.append("实验面积：" + _range_text([float(f["metadata"]["实验面积(mm2)"]) for f in raw["files"]]) + "mm²")
    fields["test_parameters"] = "\n".join(methods)
    for key, val in (profile.defaults if profile else {}).items():
        fields.setdefault(key, val)
    require(set(metadata or {}) <= set(FIELD_CATALOG), "REPORT_FIELD", "补充信息包含未知字段")
    fields.update(metadata or {})
    requested = list(dict.fromkeys([*profile.fields, *(b.field for b in profile.text_runs)])) if profile else FIELD_CATALOG
    missing = [FIELD_CATALOG[k] for k in requested if fields.get(k) is None or fields.get(k) == ""]
    return template, order, raw, profile, fields, samples, files, stats, missing, warnings


def inspect_report_inputs(template_path: str, order_path: str, raw_zip_path: str, profile_path: str | None = None,
                          metadata: dict[str, Scalar] | None = None) -> ReportInspection:
    template, order, raw, profile, fields, samples, files, stats, missing, warnings = _prepare(template_path, order_path, raw_zip_path, profile_path, metadata)
    return ReportInspection(template_path=str(template.package.path), template_fingerprint=template.package.fingerprint,
                            fields=fields, samples=samples, raw_files=[{k: v for k, v in f.items() if k not in ("records", "instrument_summary")} | {"record_count": len(f["records"])} for f in raw["files"]],
                            statistics=stats, missing_fields=missing, warnings=warnings,
                            suggested_profile=profile.model_dump(mode="json", exclude_none=True) if profile else None,
                            mapping_status="profile_loaded" if profile_path else "known_template_suggestion" if profile else "mapping_required",
                            available_objects=[{k: v for k, v in o.items() if not k.startswith('_')} for o in drawing_objects(template)],
                            available_targets=template.names, field_catalog=FIELD_CATALOG)


def _operations(template, profile, fields, samples, files, stats):
    ops = []
    for key, targets in profile.fields.items():
        for target in targets:
            ops.append({"op": "write", "target": target.model_dump(exclude_none=True), "value": fields.get(key, "") if fields.get(key) is not None else ""})
    for key, target in profile.sample_rows.items():
        _, rect = template.resolve(target)
        require(rect.size[0] == 1 and len(samples) <= rect.size[1], "REPORT_CAPACITY", "样品数超出模板预留列数", field=key, capacity=rect.size[1], samples=len(samples))
        values = [s.get(key) for s in samples] + [None] * (rect.size[1] - len(samples))
        ops.append({"op": "write", "target": target.model_dump(exclude_none=True), "value": [values]})
    summaries = {(s["sample"], s["kind"]): s for s in stats}
    for kind, slot in profile.tables.items():
        if kind in ('single_measurements', 'pressure_measurements'):
            ops.extend(_measurement_table(template, kind, slot, samples, files, summaries))
            continue
        if kind == "single_summary":
            rows = [["样品", "点数", "平均压强", "平均厚度", "平均电阻", "电阻率均值", "样本标准差", "样本COV"],
                    ["", "", "MPa", "μm", "Ω", "Ω·cm", "Ω·cm", "%"]]
            for sample in samples:
                summary = summaries.get((sample["sample_name"], "single_point"))
                require(summary, "RAW_MODE_MISSING", "样品缺少单点记录", sample=sample["sample_name"])
                rows.append([sample["sample_name"], summary["n"], round(summary["mean"]["pressure_mpa"], 2),
                             round(summary["mean"]["thickness_um"], 2), round(summary["mean"]["resistance_ohm"], 7),
                             round(summary["mean"]["resistivity_ohm_cm"], 4),
                             round(summary["sample_std"]["resistivity_ohm_cm"], 4) if summary["sample_std"]["resistivity_ohm_cm"] is not None else None,
                             round(summary["sample_cov_pct"]["resistivity_ohm_cm"], 2) if summary["sample_cov_pct"]["resistivity_ohm_cm"] is not None else None])
        else:
            series = [files.get((s["sample_name"], "pressure")) for s in samples]
            require(all(series), "RAW_MODE_MISSING", "部分样品缺少变压记录")
            steps = [r["nominal_pressure_mpa"] for r in series[0]["records"]]
            require(all([r["nominal_pressure_mpa"] for r in s["records"]] == steps for s in series),
                    "REPORT_PRESSURE_GRID", "各样品压强工步不同；不能按行合并，需配置独立表格")
            rows = [["设定压强"] + [s["sample_name"] for s in samples], ["MPa"] + ["Ω·cm"] * len(samples)]
            rows += [[p] + [round(s["records"][i]["resistivity_ohm_cm"], 4) for s in series] for i, p in enumerate(steps)]
        sheet, rect = template.resolve(slot.target)
        require(len(rows) <= rect.size[0] and len(rows[0]) <= rect.size[1], "REPORT_CAPACITY", "结果表超出预留区域", table=kind, required=[len(rows), len(rows[0])], capacity=list(rect.size))
        padded = [r + [None] * (rect.size[1] - len(r)) for r in rows]
        padded += [[None] * rect.size[1] for _ in range(rect.size[0] - len(rows))]
        ops.append({"op": "write", "target": slot.target.model_dump(exclude_none=True), "value": padded})
    # Reject ambiguous overlap and rich-text/type issues before any plotting.
    preflight_cells(template, OPERATIONS.validate_python(ops))
    return ops


def _measurement_table(template, kind, slot, samples, files, summaries):
    sheet, rect = template.resolve(slot.target)
    capacity = (rect.size[1] - 1) // 2
    require(rect.size[1] == 1 + 2 * capacity and len(samples) <= capacity,
            'REPORT_CAPACITY', '测量表必须预留编号列及两组样品列', table=kind)
    single = kind == 'single_measurements'
    mode = 'single_point' if single else 'pressure'
    records = [files.get((s['sample_name'], mode)) for s in samples]
    require(all(records), 'RAW_MODE_MISSING', '测量表所需原始记录缺失')
    key = 'point' if single else 'nominal_pressure_mpa'
    points = [r[key] for r in records[0]['records']]
    require(all([r[key] for r in item['records']] == points for item in records),
            'REPORT_PRESSURE_GRID', '测量表各样品的测量点或工步不一致，不能按行拼接')
    data_rows = rect.size[0] - (4 if single else 2)
    require(0 < len(points) <= data_rows, 'REPORT_CAPACITY', '逐点记录超出测量表预留行数', table=kind, records=len(points), capacity=data_rows)
    headers = [s['sample_name'] for s in samples] + [None] * (capacity - len(samples))
    rows = [['编号' if single else '设定压强（MPa）', *headers, *headers]]
    metrics = ('resistivity_ohm_cm', 'thickness_um')
    for i, point in enumerate(points):
        rows.append([point, *[round(item['records'][i][metrics[0]], 4) for item in records], *[None] * (capacity - len(samples)),
                     *[round(item['records'][i][metrics[1]], 2) for item in records], *[None] * (capacity - len(samples))])
    rows += [[None] * rect.size[1] for _ in range(data_rows - len(points))]
    if single:
        for label, field in (('平均值', 'mean'), ('COV（%）', 'sample_cov_pct')):
            values = [label]
            for metric in metrics:
                values += [round(summaries[(s['sample_name'], mode)][field][metric], 4)
                           if summaries[(s['sample_name'], mode)][field][metric] is not None else None for s in samples]
                values += [None] * (capacity - len(samples))
            rows.append(values)
    def target(address):return {'sheet':sheet['name'], 'range':address}
    return [{'op':'write', 'target':target(cell_address(rect.top, rect.left + 1)), 'value':'电阻率（Ω·cm）'},
            {'op':'write', 'target':target(cell_address(rect.top, rect.left + 1 + capacity)), 'value':'极片总厚度（μm）'},
            {'op':'write', 'target':target(Rectangle(rect.top + 1, rect.left, rect.bottom, rect.right).address), 'value':rows}]


def generate_report(template_path: str, order_path: str, raw_zip_path: str, output_directory: str,
                    profile_path: str | None = None, metadata: dict[str, Scalar] | None = None,
                    profile: ReportProfile | None = None) -> ReportResult:
    destination = Path(output_directory)
    require(destination.is_absolute(), "ABSOLUTE_PATH_REQUIRED", "输出目录必须使用绝对路径")
    destination = destination.resolve()
    require(destination.parent.is_dir() and not destination.exists(), "OUTPUT_EXISTS", "输出目录必须不存在，其父目录必须存在")
    require(profile_path or profile is not None, "REPORT_PROFILE", "必须指定已确认的 profile 或 profile_path")
    template, order, raw, profile, fields, samples, files, stats, missing, warnings = _prepare(template_path, order_path, raw_zip_path, profile_path, metadata, profile)
    require(profile is not None, "REPORT_PROFILE", "请先确认并保存模板映射")
    # Preflight remaining objects after removing the reviewed examples. This in-memory
    # package is discarded; the final output still starts from the original input.
    apply_example_removals(template, profile.example_objects)
    preflight_offline(template)
    operations = _operations(template, profile, fields, samples, files, stats)
    text_runs = [{"target": binding.target.model_dump(exclude_none=True), "run_index": binding.run_index,
                  "value": binding.prefix + (str(fields[binding.field]) if fields.get(binding.field) is not None else "")} for binding in profile.text_runs]
    run_keys = set()
    changes = preflight_cells(template, OPERATIONS.validate_python(operations))
    for binding in profile.text_runs:
        sheet, rect = template.resolve(binding.target)
        rect = template.merged_anchor(sheet, rect, True)
        key = (sheet["name"], rect.address)
        require(key not in run_keys and rect.address not in changes.get(sheet["name"], {}), "TARGET_CONFLICT", "富文本片段与其他填充目标冲突")
        run_keys.add(key)
    selected = []
    if profile.checkboxes:
        require(profile.order.checkbox_sheet, "CHECKBOX_MAPPING", "未配置来源复选框工作表")
        source_controls = checkboxes(order, profile.order.checkbox_sheet)
        for binding in profile.checkboxes:
            matches = [c for c in source_controls if c["label"] == binding.source_label]
            require(len(matches) == 1, "CHECKBOX_MAPPING", "来源复选框不唯一或不存在", label=binding.source_label)
            target_matches = [c for c in checkboxes(template, binding.target_sheet) if c["label"] == binding.target_label]
            require(len(target_matches) == 1, "CHECKBOX_MAPPING", "目标复选框不唯一或不存在", label=binding.target_label)
            selected.append({"sheet": binding.target_sheet, "label": binding.target_label, "checked": matches[0]["checked"]})
    sources = {str(template.package.path): template.package.fingerprint, str(order.package.path): order.package.fingerprint,
               raw["path"]: raw["sha256"]}
    if profile_path:
        profile_file = absolute_file(profile_path, ".json")
        sources[str(profile_file)] = digest(profile_file.read_bytes())
    images = []
    plot_records = []
    with tempfile.TemporaryDirectory(prefix=".xmlmcp-report-", dir=destination.parent) as temporary:
        bundle = Path(temporary) / "bundle"
        bundle.mkdir()
        for kind in [*profile.figures, *profile.figure_replacements]:
            require(all((s['sample_name'], plot_mode(kind)) in files for s in samples), 'RAW_MODE_MISSING', '图片所需模式缺失', kind=kind)
            filename = {'single_point':'single-point.png', 'pressure':'pressure-curve.png'}.get(kind, kind.replace('_','-') + '.png')
            image = bundle / filename
            if kind in profile.figure_replacements:
                op = ReplaceImage(op='replace_image', image_id=profile.figure_replacements[kind], image_path=str(image))
                matches = [i for i in template.images() if i['image_id'] == op.image_id]
                require(len(matches) == 1, 'REPORT_OBJECT_MAPPING', '待替换统计图不存在或不唯一')
                change = ImageChange(op, b'', 1, 1, 'png', matches[0]['sheet'], image=matches[0])
            else:
                target = profile.figures[kind]
                sheet, rect = template.resolve(target)
                cells = template.cell_map(sheet)
                require(not any(template.value(cells.get(a)) not in (None, '', ' ') for a in rect.cells()),
                        'REPORT_FIGURE_OVERLAP', '图片预留区包含文字或数据，不覆盖已有内容', kind=kind)
                op = AddImage(op='add_image', target=target, image_path=str(image))
                change = ImageChange(op, b'', 1, 1, 'png', sheet['name'], rect)
            dimensions = image_geometry(template, [change], profile.maximum_digit_width_px)[change.geometry_key]
            record = make_plot(kind, samples, files, image, dimensions['width'] / dimensions['height'], profile.figure_style)
            if record:plot_records.append(record)
            images.append(image.name)
            operations.append(op.model_dump(exclude_none=True))
        require(operations, "NO_OPERATIONS", "映射未配置任何输出内容")
        provenance = {"adapter": profile.data_adapter, "profile_id": profile.profile_id, "sources": sources,
                      "statistics_method": "mean, sample standard deviation (n-1), sample COV = std / abs(mean) * 100; no interpolation",
                      "profile_sha256": digest(profile.model_dump_json().encode("utf-8")),
                      "external_refresh": False, "office_required": False}
        applied = apply_changes(str(template.package.path), str(bundle / "report.xlsx"), operations,
                                offline=True, digit_width=profile.maximum_digit_width_px, checkbox_values=selected, text_run_values=text_runs,
                                preserve_styles=True, example_objects=profile.example_objects,
                                discard_replaced_media=bool(profile.figure_replacements), provenance=provenance)
        data_report = {"provenance": provenance, "fields": fields, "samples": samples, "statistics": stats, "raw_files": raw["files"],
                       "missing_fields": missing, "warnings": warnings, "checkbox_values": selected,
                       "example_objects": [o.model_dump() for o in profile.example_objects],
                       'plots': plot_records,
                       "computation": {"engine": "python", "office_required": False, "workbook_recalculated": False, "plots": "static PNG; regenerate to update"}}
        (bundle / "report-data.json").write_text(json.dumps(data_report, ensure_ascii=False, indent=2), encoding="utf-8")
        (bundle / "template-profile.json").write_text(profile.model_dump_json(indent=2), encoding="utf-8")
        integrity_path = bundle / "report.integrity.json"
        integrity = json.loads(integrity_path.read_text(encoding="utf-8"))
        integrity["output_path"] = str(destination / "report.xlsx")
        for op in integrity["operations"]:
            if "image_path" in op:
                op["image_path"] = str(destination / Path(op["image_path"]).name)
        integrity_path.write_text(json.dumps(integrity, ensure_ascii=False, indent=2), encoding="utf-8")
        require(all(digest(Path(path).read_bytes()) == fingerprint for path, fingerprint in sources.items()),
                "INPUT_CHANGED", "生成过程中输入文件或映射改变，取消输出")
        # Directory rename publishes the entire bundle together. Neither caller outputs nor inputs are replaced.
        require(not destination.exists(), "OUTPUT_EXISTS", "输出目录已被其他调用创建")
        publish_bundle(bundle, destination)
    return ReportResult(output_directory=str(destination), output_path=str(destination / "report.xlsx"),
                        integrity_report_path=str(destination / "report.integrity.json"), data_report_path=str(destination / "report-data.json"),
                        image_paths=[str(destination / p) for p in images], missing_fields=missing, warnings=warnings,
                        integrity=applied.integrity, computation=data_report["computation"])
