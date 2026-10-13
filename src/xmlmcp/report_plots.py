"""Deterministic static plots, rendered without a GUI or spreadsheet application."""
from pathlib import Path
from threading import Lock
from .errors import require
from .report_data import summarize

PLOT_LOCK = Lock()
SYSTEM_FONTS = (Path("C:/Windows/Fonts/msyh.ttc"), Path("/System/Library/Fonts/PingFang.ttc"),
                Path("/System/Library/Fonts/STHeiti Light.ttc"))
SYSTEM_BOLD_FONTS = (Path('C:/Windows/Fonts/msyhbd.ttc'), Path('/System/Library/Fonts/PingFang.ttc'))
LABEL_FONT = Path(__file__).parent / "fonts" / "XMLMCPChartLabels.ttf"
LABEL_BOLD_FONT = Path(__file__).parent / 'fonts' / 'XMLMCPChartLabelsBold.ttf'
METRICS = {
    'resistance': {'key': 'resistance_ohm', 'scale': 1000, 'title': '电阻', 'y': '电阻（mΩ）'},
    'resistivity': {'key': 'resistivity_ohm_cm', 'scale': 1, 'title': '电阻率', 'y': '电阻率（Ω·cm）'},
    'conductivity': {'key': 'conductivity_s_cm', 'scale': 1, 'title': '电导率', 'y': '电导率（S/cm）'},
    'thickness': {'key': 'thickness_um', 'scale': 1, 'title': '极片总厚度', 'y': '极片总厚度（μm）'},
}
REFERENCE_PLOTS = {**{f'single_{m}': ('single_point', m) for m in METRICS},
                   'pressure_resistivity': ('pressure', 'resistivity'),
                   'pressure_thickness': ('pressure', 'thickness')}
PLOT_LABELS = {
    "single_point": {"title": "单点电阻率散点图", "x": "测量点序号", "y": "电阻率（Ω·cm）"},
    "pressure": {"title": "变压电阻率曲线", "x": "实测压强（MPa）", "y": "电阻率（Ω·cm）"},
}


def _font_properties(texts, bold=False):
    from matplotlib.font_manager import FontProperties
    from matplotlib.ft2font import FT2Font
    characters = {ord(c) for text in texts for c in text}
    paths = (*SYSTEM_BOLD_FONTS, LABEL_BOLD_FONT) if bold else (*SYSTEM_FONTS, LABEL_FONT)
    for path in paths:
        if path.is_file() and characters <= FT2Font(str(path)).get_charmap().keys():
            return FontProperties(fname=str(path))
    require(False, "PLOT_FONT", "中文标题或样品名含当前字体不能显示的字符，请安装相应中文字体")


def plot_mode(kind):
    return REFERENCE_PLOTS[kind][0] if kind in REFERENCE_PLOTS else kind


def reference_plot_data(kind, samples, files):
    mode, metric = REFERENCE_PLOTS[kind]
    spec = METRICS[metric]
    result = {'kind': kind, 'mode': mode, 'metric': spec['key'], 'scale': spec['scale'], 'series': []}
    for sample in samples:
        item = files[(sample['sample_name'], mode)]
        if mode == 'single_point':
            stats = summarize(item)
            values = {'mean': stats['mean'][spec['key']] * spec['scale'], 'cov_pct': stats['sample_cov_pct'][spec['key']]}
        else:
            values = {'x': [r['pressure_mpa'] for r in item['records']],
                      'y': [r[spec['key']] * spec['scale'] for r in item['records']]}
        result['series'].append({'sample': sample['sample_name'], **values})
    return result


def _reference_plot(kind, samples, files, path):
    import json
    import math
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.ticker import FormatStrFormatter
    payload = reference_plot_data(kind, samples, files)
    mode, metric = REFERENCE_PLOTS[kind]
    spec = METRICS[metric]
    title = spec['title'] if mode == 'single_point' else '压强与' + spec['title']
    texts = [title, spec['y'], '样品', '变异系数（%）', '压强（MPa）', *(s['sample_name'] for s in samples)]
    prop = _font_properties(texts, bold=True)
    style = {'id': 'template-statistics-v1', 'background': 'transparent', 'foreground': '#333333',
             'grid_color': '#cccccc', 'mean_color': '#ff0000', 'cov_color': '#008000',
             'series_colors': ['#dc143c', '#008000', '#0000ff', '#ffa500', '#800080'],
             'canvas_px': [2000, 1600], 'axes_box': [.15, .12, .70, .78],
             'font': Path(prop.get_file()).name, 'style_origin': 'reference PNG reconstruction'}
    fig = Figure(figsize=(10, 8), dpi=200)
    FigureCanvasAgg(fig)
    ax = fig.add_axes(style['axes_box'])
    ax.set_facecolor('none')
    for spine in ax.spines.values():
        spine.set_color(style['foreground']); spine.set_linewidth(1.5)
    ax.tick_params(axis='both', colors=style['foreground'], labelsize=10, length=0, pad=6)
    ax.set_ylabel(spec['y'], fontproperties=prop, fontsize=14, color=style['foreground'], labelpad=12)
    ax.set_title(title, fontproperties=prop, fontsize=18, color=style['foreground'], pad=20)
    ax.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
    ax.set_axisbelow(True)
    ax.grid(axis='y', color=style['grid_color'], linewidth=1, linestyle=(0, (5, 5)))
    if mode == 'single_point':
        means = [v['mean'] for v in payload['series']]
        covs = [v['cov_pct'] for v in payload['series']]
        positions = list(range(len(samples)))
        ax.plot(positions, means, linestyle='none', marker='o', markersize=8,
                color=style['mean_color'], markeredgewidth=0)
        for x, y in zip(positions, means):
            ax.annotate(f'{y:.2f}', (x, y), xytext=(-8, 0), textcoords='offset points',
                        ha='right', va='center', color=style['mean_color'], fontproperties=prop, fontsize=10)
        ax.set_xticks(positions, [s['sample_name'] for s in samples], fontproperties=prop)
        ax.set_xlim(-.5, len(samples) - .5)
        if metric != 'thickness' and min(means) >= 0:
            lower, upper = 0, max(means) * 1.16 if max(means) else 1
        else:
            spread = max(means) - min(means) or abs(max(means)) * .1 or 1
            lower, upper = min(means) - spread * .1, max(means) + spread * .1
        ax.set_ylim(lower, upper)
        ax.set_yticks([lower + (upper - lower) * i / 5 for i in range(6)])
        right = ax.twinx()
        for side, spine in right.spines.items():
            spine.set_visible(side == 'right')
            spine.set_color(style['cov_color']); spine.set_linewidth(1.5)
        right.tick_params(axis='y', colors=style['cov_color'], labelsize=10, length=0, pad=6)
        right.set_ylabel('变异系数（%）', color=style['cov_color'], fontproperties=prop, fontsize=14, labelpad=10)
        right.yaxis.set_major_formatter(FormatStrFormatter('%.2f%%'))
        valid = [(x, c) for x, c in zip(positions, covs) if c is not None]
        if valid:
            span = max(c for _, c in valid) - min(c for _, c in valid) or max(c for _, c in valid) * .1 or 1
            low = max(0, min(c for _, c in valid) - span * .1)
            high = max(c for _, c in valid) + span * .1
            right.set_ylim(low, high)
            right.set_yticks([low + (high - low) * i / 5 for i in range(6)])
            right.plot([x for x, c in valid], [c for x, c in valid], linestyle='none', marker='^', markersize=11,
                       color=style['cov_color'], markeredgewidth=0)
            for x, c in valid:
                right.annotate(f'{c:.2f}%', (x, c), xytext=(8, 0), textcoords='offset points',
                               ha='left', va='center', color=style['cov_color'], fontproperties=prop, fontsize=10)
        else:
            right.set_ylim(0, 1); right.set_yticks([])
        for label in right.get_yticklabels():
            label.set_fontproperties(prop); label.set_fontsize(10)
    else:
        for i, series in enumerate(payload['series']):
            ax.plot(series['x'], series['y'], color=style['series_colors'][i % len(style['series_colors'])],
                    marker='o', markersize=5, linewidth=1, label=series['sample'])
        ax.set_xlabel('压强（MPa）', fontproperties=prop, fontsize=14, color=style['foreground'], labelpad=10)
        ax.legend(loc='upper center', prop=prop, ncols=min(3, len(samples)), frameon=False,
                  labelcolor=style['foreground'], handlelength=1.5, columnspacing=1.2)
        all_y = [y for s in payload['series'] for y in s['y']]
        spread = max(all_y) - min(all_y) or abs(max(all_y)) * .1 or 1
        low, high = min(all_y) - spread * .08, max(all_y) + spread * .16
        if min(all_y) >= 0: low = max(0, low)
        ax.set_ylim(low, high)
        ax.set_yticks([low + (high - low) * i / 5 for i in range(6)])
        all_x = [x for s in payload['series'] for x in s['x']]
        ax.set_xlim(min(0, min(all_x)), math.ceil(max(all_x) / 10) * 10 or 1)
    for label in (*ax.get_xticklabels(), *ax.get_yticklabels()):
        label.set_fontproperties(prop); label.set_fontsize(10)
    fig.savefig(path, format='png', transparent=True,
                metadata={'Software': 'XMLMCP Python report', 'XMLMCPStyle': json.dumps(style, ensure_ascii=False)})
    return {'style': style, 'data': payload}


def make_plot(kind: str, samples: list[dict], files: dict, path: Path, aspect: float = 2.8, style='standard'):
    # Matplotlib is lazy-loaded so inspection and stdio startup remain lightweight.
    with PLOT_LOCK:
        if kind in REFERENCE_PLOTS:
            require(style == 'template-statistics', 'PLOT_STYLE', '统计图需使用已确认的 template-statistics 样式')
            return _reference_plot(kind, samples, files, path)
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        labels = PLOT_LABELS[kind]
        prop = _font_properties([*labels.values(), *(s["sample_name"] for s in samples)])
        fig = Figure(figsize=(11, 11 / max(1.8, min(3.2, aspect))), dpi=180)
        FigureCanvasAgg(fig)
        ax = fig.add_subplot()
        colors = ("#17689c", "#bf5b1b", "#21815b", "#8a569c", "#ba3e60", "#687a18", "#385974", "#6b4d31")
        for i, sample in enumerate(samples):
            item = files[(sample["sample_name"], kind)]
            y = [r["resistivity_ohm_cm"] for r in item["records"]]
            color = colors[i % len(colors)]
            if kind == "single_point":
                x = [r["point"] for r in item["records"]]
                ax.scatter(x, y, color=color, label=sample["sample_name"], s=45, alpha=.85, edgecolors="white", linewidths=.5)
            else:
                ax.plot([r["pressure_mpa"] for r in item["records"]], y, color=color, marker="o", markersize=4, linewidth=1.7, label=sample["sample_name"])
        ax.set_xlabel(labels["x"], fontproperties=prop, fontsize=14)
        ax.set_ylabel(labels["y"], fontproperties=prop, fontsize=14)
        ax.set_title(labels["title"], fontproperties=prop, fontsize=16, loc="left", pad=12)
        ax.tick_params(labelsize=12)
        ax.grid(True, color="#d5dde3", linewidth=.6, alpha=.8)
        ax.spines[["top", "right"]].set_visible(False)
        ax.set_axisbelow(True)
        ax.legend(loc="upper right", prop=prop, ncols=min(5, len(samples)), frameon=False)
        fig.subplots_adjust(left=.095, right=.98, bottom=.18, top=.86)
        fig.savefig(path, format="png", facecolor="white", metadata={"Software": "XMLMCP Python report"})
