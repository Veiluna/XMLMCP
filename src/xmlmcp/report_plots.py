"""Deterministic static plots, rendered without a GUI or spreadsheet application."""
from pathlib import Path
from threading import Lock
from .errors import require

PLOT_LOCK = Lock()
SYSTEM_FONTS = (Path("C:/Windows/Fonts/msyh.ttc"), Path("/System/Library/Fonts/PingFang.ttc"),
                Path("/System/Library/Fonts/STHeiti Light.ttc"))
LABEL_FONT = Path(__file__).parent / "fonts" / "XMLMCPChartLabels.ttf"
PLOT_LABELS = {
    "single_point": {"title": "单点电阻率散点图", "x": "测量点序号", "y": "电阻率（Ω·cm）"},
    "pressure": {"title": "变压电阻率曲线", "x": "实测压强（MPa）", "y": "电阻率（Ω·cm）"},
}


def _font_properties(texts):
    from matplotlib.font_manager import FontProperties
    from matplotlib.ft2font import FT2Font
    characters = {ord(c) for text in texts for c in text}
    for path in (*SYSTEM_FONTS, LABEL_FONT):
        if path.is_file() and characters <= FT2Font(str(path)).get_charmap().keys():
            return FontProperties(fname=str(path))
    require(False, "PLOT_FONT", "中文标题或样品名含当前字体不能显示的字符，请安装相应中文字体")


def make_plot(kind: str, samples: list[dict], files: dict, path: Path, aspect: float = 2.8):
    # Matplotlib is lazy-loaded so inspection and stdio startup remain lightweight.
    with PLOT_LOCK:
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
