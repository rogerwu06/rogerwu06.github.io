from __future__ import annotations
import math
import re
from collections import OrderedDict
from typing import List, Sequence

ORIGINAL_COMPAT_MEMBER_TRIM = 5

def _as_number(value: object, *, integer: bool = False, field: str = ""):
    text = str(value).strip()
    if text == "":
        raise ValueError(f"{field or 'Value'} is blank.")
    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(f"{field or 'Value'} must be numeric: {text!r}") from exc
    if not math.isfinite(number):
        raise ValueError(f"{field or 'Value'} must be finite.")
    if integer:
        rounded = int(round(number))
        if abs(number - rounded) > 1e-9:
            raise ValueError(f"{field or 'Value'} must be an integer: {text!r}")
        return rounded
    return number

def _general_2dp(value: float) -> str:
    return f"{value:.2f}".rstrip("0").rstrip(".")

def _roundup_2(value: float) -> float:
    return math.ceil(value * 100.0 - 1e-12) / 100.0

def _k_factor(l_over_halpha: float) -> float:
    x = l_over_halpha
    if x <= 5:
        return 0.60
    if x <= 10:
        return 0.60 + (x - 5) * (0.65 - 0.60) / (10 - 5)
    if x <= 20:
        return 0.65 + (x - 10) * (0.75 - 0.65) / (20 - 10)
    if x <= 35:
        return 0.75 + (x - 20) * (0.85 - 0.75) / (35 - 20)
    if x <= 50:
        return 0.85 + (x - 35) * (0.90 - 0.85) / (50 - 35)
    if x <= 100:
        return 0.90 + (x - 50) * (0.95 - 0.90) / (100 - 50)
    return 0.95

def calculate_wind_outputs(
    node_rows: Sequence[Sequence[object]],
    member_rows: Sequence[Sequence[object]],
    section_rows: Sequence[Sequence[object]],
    height_adjustment: float,
    q: float,
    cg: float,
    *,
    exact_workbook_compatibility: bool = True,
):
    """Return (max_height, formatted_gz_lines, formatted_gx_lines)."""
    nodes = {}
    for row_index, row in enumerate(node_rows, start=1):
        if not row or all(str(v).strip() == "" for v in row):
            continue
        if len(row) < 4:
            raise ValueError(f"Node row {row_index} has fewer than 4 columns.")
        node = _as_number(row[0], integer=True, field=f"Node row {row_index}: NODE")
        x = _as_number(row[1], field=f"Node {node}: X")
        y = _as_number(row[2], field=f"Node {node}: Y")
        z = _as_number(row[3], field=f"Node {node}: Z")
        if node not in nodes:
            nodes[node] = (x, y, z)

    if not nodes:
        raise ValueError("No node data is present.")

    sections = {}
    for row_index, row in enumerate(section_rows, start=1):
        if not row or all(str(v).strip() == "" for v in row):
            continue
        if len(row) < 5:
            raise ValueError(f"Section-property row {row_index} has fewer than 5 columns.")
        prop = _as_number(row[0], integer=True, field=f"Section row {row_index}: prop")
        d = _as_number(row[3], field=f"Section prop {prop}: d")
        bf = _as_number(row[4], field=f"Section prop {prop}: bf")
        if d <= 0 or bf <= 0:
            raise ValueError(f"Section prop {prop}: d and bf must both be greater than zero.")
        if prop not in sections:
            sections[prop] = (d, bf)

    if not sections:
        raise ValueError("No section-property data is present.")

    members = []
    for row_index, row in enumerate(member_rows, start=1):
        if not row or all(str(v).strip() == "" for v in row):
            continue
        if len(row) < 7:
            raise ValueError(f"Member row {row_index} has fewer than 7 columns.")
        beam = _as_number(row[0], integer=True, field=f"Member row {row_index}: BEAM")
        node_a = _as_number(row[1], integer=True, field=f"Beam {beam}: NODEA")
        node_b = _as_number(row[2], integer=True, field=f"Beam {beam}: NODE B")
        prop = _as_number(row[3], integer=True, field=f"Beam {beam}: PROPERTY DEFINE")
        material = str(row[4]).strip()
        beta = _as_number(row[5], field=f"Beam {beam}: BETA")
        length = _as_number(row[6], field=f"Beam {beam}: LENGTH")
        if node_a not in nodes:
            raise ValueError(f"Beam {beam} references missing node {node_a}.")
        if node_b not in nodes:
            raise ValueError(f"Beam {beam} references missing node {node_b}.")
        if prop not in sections:
            raise ValueError(f"Beam {beam} references missing section property {prop}.")
        members.append((beam, node_a, node_b, prop, material, beta, length))

    if not members:
        raise ValueError("No member data is present.")

    heights = []
    for beam, node_a, node_b, *_ in members:
        lower_y = min(nodes[node_a][1], nodes[node_b][1])
        height = lower_y + height_adjustment
        if height < 0:
            raise ValueError(
                f"Beam {beam} has calculated height {height:g} m. "
                "The spreadsheet's Ce equation is not valid for a negative height; "
                "adjust the base elevation/height adjustment."
            )
        heights.append(height)

    max_height = max(heights)

    filtered_members = members
    if exact_workbook_compatibility and len(filtered_members) > ORIGINAL_COMPAT_MEMBER_TRIM:
        filtered_members = filtered_members[:-ORIGINAL_COMPAT_MEMBER_TRIM]

    def direction_output(direction: str) -> List[str]:
        code = "GZ" if direction == "Z" else "GX"
        individual = []

        for beam, node_a, node_b, prop, _material, _beta, _length in filtered_members:
            xa, ya, za = nodes[node_a]
            xb, yb, zb = nodes[node_b]

            if direction == "Z":
                keep = (xa != xb) or (ya != yb)
            else:
                keep = (za != zb) or (ya != yb)
            if not keep:
                continue

            d, bf = sections[prop]
            h_alpha = math.sqrt(d * d + bf * bf)
            l_over_halpha = max_height * 1000.0 / h_alpha
            k = _k_factor(l_over_halpha)
            height = min(ya, yb) + height_adjustment
            ce = max((height / 10.0) ** 0.2, 0.90)

            fn = ce * cg * k * 2.1 * q
            ft = fn
            kn_per_m_fn = fn * d / 1000.0
            kn_per_m_ft = ft * bf / 1000.0
            load = _roundup_2(max(kn_per_m_fn, kn_per_m_ft))
            individual.append((beam, f"UNI {code} {_general_2dp(load)}"))

        groups = OrderedDict()
        for beam, suffix in individual:
            groups.setdefault(suffix, []).append(str(beam))
        return [" ".join(beams) + " " + suffix for suffix, beams in groups.items()]

    return max_height, direction_output("Z"), direction_output("X")

def _parse_clipboard_table(text: str, expected_cols: int) -> List[List[str]]:
    rows: List[List[str]] = []
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        if "\t" in raw_line:
            parts = [p.strip() for p in raw_line.split("\t")]
        elif "," in raw_line:
            parts = [p.strip().strip('"') for p in raw_line.split(",")]
        else:
            parts = re.split(r"\s+", line)

        try:
            float(parts[0])
        except (ValueError, IndexError):
            continue

        if len(parts) < expected_cols:
            parts.extend([""] * (expected_cols - len(parts)))
        rows.append(parts[:expected_cols])

    if not rows:
        raise ValueError(
            "No data rows were found. Copy the table rows from STAAD/Excel "
            "and paste them again."
        )
    return rows
