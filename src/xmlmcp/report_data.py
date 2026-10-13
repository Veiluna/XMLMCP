"""Read instrument CSV records directly from ZIP; no extraction or Office engine."""
from __future__ import annotations

import csv
import io
import math
import re
import statistics
import zipfile
from datetime import datetime
from pathlib import PurePosixPath

from .errors import XMLMCPError, require
from .package import absolute_file
from .xmlbytes import digest

HEADERS = {"序号": "point", "压力(kg)": "force_kg", "压强(MPa)": "pressure_mpa",
           "极片总厚度(μm)": "thickness_um", "电阻(Ω)": "resistance_ohm",
           "电导率(S/cm)": "conductivity_s_cm", "电阻率(Ω*cm)": "resistivity_ohm_cm"}


def normal(value: str) -> str:
    return value.strip().replace("（", "(").replace("）", ")").replace("µ", "μ").replace("²", "2")


def number(value: str, file: str, field: str) -> float:
    try:
        result = float(value)
    except ValueError as exc:
        raise XMLMCPError("RAW_NUMBER", "测量值缺失或不是数值", file=file, field=field, value=value) from exc
    require(math.isfinite(result), "RAW_NUMBER", "测量值必须是有限数值", file=file, field=field)
    return result


def parse_csv(data: bytes, filename: str) -> dict:
    text = None
    for encoding in ("utf-8-sig", "gb18030", "utf-16"):
        try:
            candidate = data.decode(encoding)
            if "实验报告" in candidate:
                text = candidate
                break
        except UnicodeError:
            continue
    require(text is not None, "RAW_FORMAT", "不支持的仪器 CSV 标题或编码", file=filename)
    # Some instrument exports put one NUL after each newline; do not remove embedded NULs.
    nul_count = text.count("\x00")
    require(all(not line.lstrip("\x00").count("\x00") for line in text.splitlines()),
            "RAW_CONTROL_CHARACTER", "CSV 含行首之外的 NUL 字符", file=filename)
    rows = list(csv.reader(io.StringIO("\n".join(line.lstrip("\x00") for line in text.splitlines()))))
    title = rows[0][0].strip() if rows and rows[0] else ""
    require(title in ("单点实验报告", "变压实验报告"), "RAW_FORMAT", "仅支持单点/变压仪器 CSV", file=filename)
    kind = "single_point" if title == "单点实验报告" else "pressure"
    header_index = next((i for i, row in enumerate(rows) if row and normal(row[0]) == "序号"), None)
    require(header_index is not None, "RAW_HEADER", "CSV 缺少测量列标题", file=filename)
    keys = [HEADERS.get(normal(v)) for v in rows[header_index]]
    require(all(keys.count(key) == 1 for key in HEADERS.values()), "RAW_HEADER", "CSV 测量列缺失或重复", file=filename)
    metadata = {}
    for row in rows[:header_index]:
        for item in row:
            item = item.strip()
            if ":" in item or "：" in item:
                key, val = re.split("[:：]", item, maxsplit=1)
                require(key not in metadata, "RAW_METADATA", "CSV 参数重复", file=filename, field=key)
                metadata[normal(key)] = val.strip()
    records, exported = [], {}
    for row in rows[header_index + 1:]:
        if not row or not row[0].strip():
            continue
        if row[0].strip() in ("平均值", "COV%"):
            require(row[0].strip() not in exported, "RAW_SUMMARY", "CSV 汇总行重复", file=filename)
            exported[row[0].strip()] = {key: number(row[i], filename, key) for i, key in enumerate(keys) if key and key != "point" and i < len(row)}
            continue
        require(row[0].strip().isdigit(), "RAW_RECORD", "测量区包含未知记录", file=filename, row=row[0])
        require(len(row) >= len(keys) - (1 if keys[-1] is None else 0), "RAW_RECORD", "测量行缺列", file=filename)
        record = {key: number(row[i], filename, key) for i, key in enumerate(keys) if key}
        require(record["point"] == len(records) + 1, "RAW_POINT_SEQUENCE", "测量序号不连续或重复", file=filename)
        require(record["thickness_um"] > 0 and record["pressure_mpa"] >= 0,
                "RAW_RECORD", "厚度应大于零，压强不能小于零", file=filename)
        records.append(record)
    require(records, "RAW_EMPTY", "CSV 没有测量记录", file=filename)
    if kind == "single_point":
        require(metadata.get("采样数目", "").isdigit() and int(metadata["采样数目"]) == len(records),
                "RAW_SAMPLE_COUNT", "采样数目与逐点记录不符", file=filename)
    for key in ("开始时间", "结束时间"):
        require(key in metadata, "RAW_METADATA", "CSV 缺少时间", file=filename, field=key)
        try:
            datetime.strptime(metadata[key], "%Y-%m-%d %H:%M:%S")
        except ValueError as exc:
            raise XMLMCPError("RAW_DATE", "仪器日期格式无效", file=filename, field=key) from exc
    require(metadata["结束时间"] >= metadata["开始时间"], "RAW_DATE", "结束时间早于开始时间", file=filename)
    for key in ("温度(℃)", "湿度(%RH)", "实验面积(mm2)"):
        require(key in metadata, "RAW_METADATA", "CSV 缺少环境/面积参数", file=filename, field=key)
        number(metadata[key], filename, key)
    settings = {}
    if kind == "pressure":
        for key in ("压强下限(MPa)", "压强上限(MPa)", "压强步进(MPa)", "保压时间"):
            require(key in metadata, "RAW_METADATA", "变压 CSV 缺少工步参数", file=filename, field=key)
            match = re.fullmatch(r"([+-]?[0-9.]+)(?:\((?:MPa|s)\))?", metadata[key])
            require(match, "RAW_METADATA", "变压参数格式无效", file=filename, field=key)
            settings[key] = number(match[1], filename, key)
        lower, upper, step = (settings[k] for k in ("压强下限(MPa)", "压强上限(MPa)", "压强步进(MPa)"))
        require(step > 0 and upper >= lower, "RAW_PRESSURE_STEPS", "压强工步无效", file=filename)
        expected = (upper - lower) / step + 1
        require(math.isclose(expected, len(records), abs_tol=1e-8), "RAW_PRESSURE_STEPS", "工步数与测量点数不符；不补点或插值", file=filename)
        for i, record in enumerate(records):
            nominal = lower + step * i
            require(abs(record["pressure_mpa"] - nominal) <= max(0.5, step * 0.1),
                    "RAW_PRESSURE_STEPS", "实测压强无法映射至设定工步", file=filename, point=i + 1)
            record["nominal_pressure_mpa"] = nominal
    return {"sample": PurePosixPath(filename).stem, "kind": kind, "filename": filename,
            "sha256": digest(data), "encoding": encoding, "removed_line_prefix_nuls": nul_count,
            "metadata": metadata, "settings": settings, "records": records, "instrument_summary": exported}


def read_raw_zip(path: str) -> dict:
    source = absolute_file(path, ".zip")
    files, warnings = [], []
    payload = source.read_bytes()
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            entries = archive.infolist()
            require(len(entries) <= 1000 and sum(e.file_size for e in entries) <= 100_000_000,
                    "RAW_ZIP_SIZE", "ZIP 超出 1000 项或 100MB 解压数据上限")
            seen = set()
            for entry in entries:
                name = entry.filename.replace("\\", "/")
                p = PurePosixPath(name)
                require(not p.is_absolute() and ".." not in p.parts and ":" not in name,
                        "RAW_ZIP_PATH", "ZIP 含越界路径", file=name)
                require(name not in seen, "RAW_DUPLICATE", "ZIP 含重名项", file=name)
                seen.add(name)
                if entry.is_dir() or name.startswith("__MACOSX/") or p.name.startswith("._") or p.name == ".DS_Store":
                    continue
                require(p.suffix.lower() == ".csv", "RAW_FORMAT", "首版原始数据 ZIP 只支持仪器 CSV；未忽略未知数据", file=name)
                require(entry.file_size <= 10_000_000 and not entry.flag_bits & 1, "RAW_ZIP_SIZE", "CSV 过大或已加密", file=name)
                item = parse_csv(archive.read(entry), name)
                require(not any(f["sample"] == item["sample"] and f["kind"] == item["kind"] for f in files),
                        "RAW_DUPLICATE_SAMPLE", "同一样品/模式出现多个文件", sample=item["sample"], kind=item["kind"])
                files.append(item)
                if item["removed_line_prefix_nuls"]:
                    warnings.append({"code": "INSTRUMENT_NUL_PREFIX", "file": name, "count": item["removed_line_prefix_nuls"]})
    except (zipfile.BadZipFile, RuntimeError) as exc:
        raise XMLMCPError("RAW_ZIP", "无法读取原始数据 ZIP", error=str(exc)) from exc
    require(files, "RAW_EMPTY", "ZIP 没有仪器 CSV")
    return {"path": str(source), "sha256": digest(payload), "files": files, "warnings": warnings}


def summarize(item: dict) -> dict:
    records = item["records"]
    result = {"sample": item["sample"], "kind": item["kind"], "n": len(records), "mean": {}, "sample_std": {}, "sample_cov_pct": {}}
    for key in HEADERS.values():
        if key == "point":
            continue
        values = [r[key] for r in records]
        mean = statistics.mean(values)
        std = statistics.stdev(values) if len(values) > 1 else None
        result["mean"][key] = mean
        result["sample_std"][key] = std
        result["sample_cov_pct"][key] = std / abs(mean) * 100 if std is not None and mean != 0 else None
    return result
