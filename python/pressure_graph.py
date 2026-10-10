from __future__ import annotations
import base64
import io
import math

DEFAULT_REINFORCED = [
    (6.0, 0.0), (7.0, 0.002209), (8.0, 0.05521), (9.0, 0.1156),
    (10.0, 0.186), (11.0, 0.2731), (12.0, 0.3772), (13.0, 0.5281),
    (14.0, 0.7744), (15.0, 1.174), (16.0, 1.979), (17.0, 4.41),
    (17.529192, 20.06),
]

DEFAULT_UNREINFORCED = [
    (6.0, 0.3915), (7.0, 0.6452), (8.0, 1.087), (9.0, 1.667),
    (10.0, 2.39), (11.0, 3.355), (12.0, 4.682), (13.0, 6.607),
    (14.0, 9.842), (15.0, 18.29), (16.0, 20.22),
]


def pairs_to_text(pairs):
    return "\n".join(f"{x:g}\t{y:g}" for x, y in pairs)


def _parse_pairs(text: str):
    rows = []
    for raw in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        raw = raw.strip()
        if not raw:
            continue
        if "\t" in raw:
            parts = [p.strip() for p in raw.split("\t") if p.strip()]
        elif "," in raw:
            parts = [p.strip() for p in raw.split(",") if p.strip()]
        else:
            parts = raw.split()
        if len(parts) < 2:
            continue
        try:
            x, y = float(parts[0]), float(parts[1])
        except ValueError:
            continue
        if not (math.isfinite(x) and math.isfinite(y)):
            raise ValueError("Graph data must contain finite numeric values.")
        rows.append((x, y))
    if len(rows) < 2:
        raise ValueError("Enter at least two numeric pressure/PEEQ rows.")
    rows.sort(key=lambda p: p[0])
    return rows


def _add_tangent_arrows(ax, x_data, y_data):
    import numpy as np
    n = len(x_data)
    if n < 3:
        return
    base_length = 1.6

    def slope(i):
        if i == 0:
            return (y_data[1] - y_data[0]) / (x_data[1] - x_data[0])
        if i == n - 1:
            return (y_data[-1] - y_data[-2]) / (x_data[-1] - x_data[-2])
        return (y_data[i + 1] - y_data[i - 1]) / (x_data[i + 1] - x_data[i - 1])

    for i in sorted(set([1, n // 2, n - 2])):
        arrow_length = base_length * (1.5 if i == n - 2 else 1.0)
        m = slope(i)
        dx, dy = 1.0, m
        mag = np.sqrt(dx * dx + dy * dy)
        if mag == 0:
            continue
        dx, dy = dx / mag * arrow_length, dy / mag * arrow_length
        ax.annotate(
            "", xy=(x_data[i] + dx, y_data[i] + dy),
            xytext=(x_data[i] - dx, y_data[i] - dy),
            arrowprops=dict(arrowstyle="->", linewidth=1.2, color="black", alpha=0.45),
            zorder=10,
        )


def render_graph(reinforced_text: str, unreinforced_text: str, design_pressure: float = 9.5) -> str:
    import matplotlib
    matplotlib.use("Agg", force=True)
    import matplotlib.pyplot as plt
    import numpy as np
    from scipy.interpolate import PchipInterpolator
    from matplotlib.lines import Line2D

    reinforced = _parse_pairs(reinforced_text)
    unreinforced = _parse_pairs(unreinforced_text)
    xr = np.array([p[0] for p in reinforced], dtype=float)
    yr = np.array([p[1] for p in reinforced], dtype=float)
    xu = np.array([p[0] for p in unreinforced], dtype=float)
    yu = np.array([p[1] for p in unreinforced], dtype=float)

    fig, ax = plt.subplots(figsize=(12.5, 5.8))
    blue = "#2d6a8a"
    teal = "#2a9d8f"

    x1 = np.linspace(xr.min(), xr.max(), 400)
    x2 = np.linspace(xu.min(), xu.max(), 400)
    rline, = ax.plot(x1, PchipInterpolator(xr, yr)(x1), linewidth=2.5, label="Reinforced", color=blue)
    uline, = ax.plot(x2, PchipInterpolator(xu, yu)(x2), linewidth=2.5, label="Unreinforced", color=teal)
    ax.plot(xr, yr, "o", markersize=5, color=blue)
    ax.plot(xu, yu, "o", markersize=5, color=teal)
    _add_tangent_arrows(ax, xr, yr)
    _add_tangent_arrows(ax, xu, yu)

    design_pressure = float(design_pressure)
    ax.axvline(design_pressure, color="#667788", linestyle="--", linewidth=1.5)
    ymax = max(float(yr.max()), float(yu.max()))
    ax.text(
        design_pressure - 0.45, ymax * 0.96, "Design Pressure",
        fontsize=10, rotation=90, va="top", ha="right",
        bbox=dict(boxstyle="round,pad=0.25", facecolor="white", edgecolor="none", alpha=0.9),
    )
    arrow_handle = Line2D([0], [0], color="black", linewidth=1.2, marker=r"$\rightarrow$", markersize=12, label="Tangent Direction")
    ax.legend(handles=[rline, uline, arrow_handle], loc="upper left", frameon=False)
    ax.set_xlabel("Applied Pressure (MPa)", fontsize=11)
    ax.set_ylabel("PEEQ (%)", fontsize=11)
    ax.set_title("Failure Pressure Prediction", fontsize=14)
    ax.grid(True, color="#d9e2e8", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlim(left=0)
    ax.set_ylim(bottom=0)
    fig.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("ascii")
