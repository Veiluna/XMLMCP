"""Deterministic static plots, rendered without a GUI or spreadsheet application."""
from pathlib import Path
from threading import Lock
from .errors import require

PLOT_LOCK = Lock()


def make_plot(kind: str, samples: list[dict], files: dict, path: Path, aspect: float = 2.8):
    # Matplotlib is lazy-loaded so inspection and stdio startup remain lightweight.
    with PLOT_LOCK:
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        from matplotlib.font_manager import FontProperties
        from matplotlib.ft2font import FT2Font
        fonts = [Path("C:/Windows/Fonts/msyh.ttc"), Path("/System/Library/Fonts/PingFang.ttc"),
                 Path("/System/Library/Fonts/STHeiti Light.ttc")]
        font = next((p for p in fonts if p.is_file()), None)
        prop = FontProperties(fname=str(font)) if font else FontProperties(family="DejaVu Sans")
        from matplotlib.font_manager import findfont
        charmap = FT2Font(findfont(prop)).get_charmap()
        require(all(all(ord(c) in charmap for c in sample["sample_name"]) for sample in samples),
                "PLOT_FONT", "样品名含当前字体不能显示的字符，请安装相应字体或使用可显示的样品标识")
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
        ax.set_xlabel("Measurement point" if kind == "single_point" else "Measured pressure (MPa)", fontsize=14)
        ax.set_ylabel("Resistivity (Ω·cm)", fontsize=14)
        ax.set_title("Single point resistivity" if kind == "single_point" else "Pressure and resistivity", fontsize=16, loc="left", pad=12)
        ax.tick_params(labelsize=12)
        ax.grid(True, color="#d5dde3", linewidth=.6, alpha=.8)
        ax.spines[["top", "right"]].set_visible(False)
        ax.set_axisbelow(True)
        ax.legend(loc="upper right", prop=prop, ncols=min(5, len(samples)), frameon=False)
        fig.subplots_adjust(left=.095, right=.98, bottom=.18, top=.86)
        fig.savefig(path, format="png", facecolor="white", metadata={"Software": "XMLMCP Python report"})
