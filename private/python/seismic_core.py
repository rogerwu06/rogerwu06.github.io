from __future__ import annotations
import math
import re
from decimal import Decimal, ROUND_HALF_UP
from typing import List, Sequence

DEFAULTS = {
    'latitude': "",
    'longitude': "",
    'sa02': "",
    'sa05': "",
    'sa10': "",
    'sa20': "",
    'sa50': "",
    'ie': "",
    'site_class': "",
    'snow': "",
    'rd': "",
    'ro': "",
    'height_adjustment': "",
}

def _as_number(value: object, *, field: str = "", integer: bool = False) -> float | int:
    text = str(value).strip().replace(",", "")
    if text == "":
        raise ValueError(f"{field or 'Value'} is blank.")
    try:
        num = float(text)
    except ValueError:
        raise ValueError(f"{field or 'Value'} must be numeric, got {value!r}.")
    if not math.isfinite(num):
        raise ValueError(f"{field or 'Value'} must be finite.")
    if integer:
        rounded = int(round(num))
        if abs(num - rounded) > 1e-9:
            raise ValueError(f"{field or 'Value'} must be an integer.")
        return rounded
    return num

def _excel_round_3(value: float) -> str:
    """Excel-like ROUND(value,3), rendered the way text concatenation normally appears."""
    d = Decimal(str(value)).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s or "0"

def _parse_clipboard_table(text: str, expected_cols: int) -> List[List[str]]:
    rows: List[List[str]] = []
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not raw_line.strip():
            continue

        if "\t" in raw_line:
            parts = [p.strip().strip('"') for p in raw_line.split("\t")]
        elif "," in raw_line:
            parts = [p.strip().strip('"') for p in raw_line.split(",")]
        else:
            parts = re.split(r"\s+", raw_line.strip())

        try:
            float(parts[0])
        except (ValueError, IndexError):
            continue

        if len(parts) < expected_cols:
            parts.extend([""] * (expected_cols - len(parts)))
        elif len(parts) > expected_cols:
            if expected_cols == 8 and len(parts) >= 9:
                parts = [parts[0], " ".join(parts[1:-6]), *parts[-6:]]
            else:
                parts = parts[:expected_cols]
        rows.append(parts[:expected_cols])

    if not rows:
        raise ValueError("No numeric data rows were found in the clipboard.")
    return rows

def _normalize_nodes(rows: Sequence[Sequence[object]]) -> dict[int, tuple[float, float, float]]:
    nodes: dict[int, tuple[float, float, float]] = {}
    for idx, row in enumerate(rows, start=1):
        if not row or all(str(v).strip() == "" for v in row):
            continue
        if len(row) < 4:
            raise ValueError(f"Node row {idx} has fewer than 4 columns.")
        node = _as_number(row[0], field=f"Node row {idx}: Node", integer=True)
        x = _as_number(row[1], field=f"Node {node}: X")
        y = _as_number(row[2], field=f"Node {node}: Y")
        z = _as_number(row[3], field=f"Node {node}: Z")
        if node in nodes:
            raise ValueError(f"Duplicate node {node} in the geometry table.")
        nodes[node] = (x, y, z)
    if not nodes:
        raise ValueError("No node geometry has been entered.")
    return nodes

def _normalize_reactions(rows: Sequence[Sequence[object]]) -> list[tuple[int, str, float, float, float, float, float, float]]:
    out = []
    seen: set[int] = set()
    for idx, row in enumerate(rows, start=1):
        if not row or all(str(v).strip() == "" for v in row):
            continue
        if len(row) < 8:
            raise ValueError(f"Reaction row {idx} has fewer than 8 columns.")
        node = _as_number(row[0], field=f"Reaction row {idx}: Node", integer=True)
        if node in seen:
            raise ValueError(
                f"Duplicate reaction node {node}. Paste only the single 1.0D + 0.25S reaction set."
            )
        seen.add(node)
        lc = str(row[1]).strip()
        nums = [
            _as_number(row[2], field=f"Reaction node {node}: Fx"),
            _as_number(row[3], field=f"Reaction node {node}: Fy"),
            _as_number(row[4], field=f"Reaction node {node}: Fz"),
            _as_number(row[5], field=f"Reaction node {node}: Mx"),
            _as_number(row[6], field=f"Reaction node {node}: My"),
            _as_number(row[7], field=f"Reaction node {node}: Mz"),
        ]
        out.append((node, lc, *nums))
    if not out:
        raise ValueError("No reactions have been entered.")
    return out

def _wrap_staads(items: Sequence[str], max_len: int = 79) -> list[str]:
    """Match the workbook's intended semicolon grouping, but include every node."""
    lines: list[str] = []
    current = ""
    for item in items:
        if not current:
            current = item
        elif len(current) + 2 + len(item) <= max_len:
            current += "; " + item
        else:
            lines.append(current)
            current = item
    if current:
        lines.append(current)
    return lines

def calculate_seismic(
    node_rows: Sequence[Sequence[object]],
    reaction_rows: Sequence[Sequence[object]],
    *,
    sa02: float,
    ie: float,
    rd: float,
    ro: float,
    height_adjustment: float,
    with_pad: bool,
) -> dict:
    """
    Recreate the workbook's intended calculation with corrected node/reaction matching.

    Vmax coefficient:
        (2/3) * Sa(0.2) * IE / (Rd * Ro)

    Nodal distribution:
        Fi = V * Wi * hi / sum(Wj * hj)

    WITH PAD:
        nodes at adjusted elevation <= 0 are treated as pad/mat reactions and excluded.
    """
    if rd == 0 or ro == 0:
        raise ValueError("Rd and Ro must be non-zero.")

    nodes = _normalize_nodes(node_rows)
    reactions = _normalize_reactions(reaction_rows)

    joined = []
    for order, reaction in enumerate(reactions):
        node, lc, fx, fy, fz, mx, my, mz = reaction
        if node not in nodes:
            raise ValueError(f"Reaction node {node} is missing from the geometry table.")
        x, y, z = nodes[node]
        hi = y + height_adjustment
        joined.append({
            "order": order,
            "node": node,
            "lc": lc,
            "x": x,
            "y": y,
            "z": z,
            "hi": hi,
            "fx": fx,
            "fy": fy,
            "fz": fz,
            "mx": mx,
            "my": my,
            "mz": mz,
        })

    joined.sort(key=lambda r: (r["hi"], r["order"]))

    total_vertical = sum(r["fy"] for r in joined)
    if with_pad:
        excluded = [r for r in joined if r["hi"] <= 0]
        active = [r for r in joined if r["hi"] > 0]
    else:
        excluded = []
        active = list(joined)

    if not active:
        raise ValueError("No active seismic nodes remain after applying the pad/base-elevation rule.")

    excluded_weight = sum(r["fy"] for r in excluded)
    effective_weight = sum(r["fy"] for r in active)
    coefficient = (2.0 / 3.0) * sa02 * ie / (rd * ro)
    vmax = coefficient * effective_weight

    wihi_sum = sum(r["fy"] * r["hi"] for r in active)
    if abs(wihi_sum) < 1e-12:
        raise ValueError("Sum(Wi × hi) is zero; check reaction weights and height adjustment.")

    loads = []
    for r in active:
        force = vmax * (r["fy"] * r["hi"]) / wihi_sum
        loads.append((r["node"], force, r["hi"], r["fy"]))

    fz_items = [f"{node} FZ {_excel_round_3(force)}" for node, force, _, _ in loads]
    fx_items = [f"{node} FX {_excel_round_3(force)}" for node, force, _, _ in loads]

    fz_lines = ["JOINT LOAD", *_wrap_staads(fz_items)]
    fx_lines = ["JOINT LOAD", *_wrap_staads(fx_items)]

    return {
        "coefficient": coefficient,
        "total_vertical": total_vertical,
        "excluded_weight": excluded_weight,
        "effective_weight": effective_weight,
        "vmax": vmax,
        "wihi_sum": wihi_sum,
        "distributed_sum": sum(force for _, force, _, _ in loads),
        "count": len(loads),
        "fz_lines": fz_lines,
        "fx_lines": fx_lines,
        "loads": loads,
    }
