#!/usr/bin/env python3
"""
STAAD.Pro Seismic Load Generator — fixed Python GUI recreation of
the original seismic calculation workbook.

Key behavior:
- Intermediate spreadsheet calculations are hidden.
- Inputs are visually separated from outputs.
- Recalculates automatically after edits/paste.
- Includes WITH PAD and NO PAD workflows.
- Uses only the Python standard library (Tkinter).
- Preserves the workbook's intended equivalent-static-force output structure.
- Fixes row-alignment / hard-coded-weight / omitted-last-node issues in the workbook.

Run:
    python STAAD_Seismic_Load_GUI.py

Logic self-test:
    python STAAD_Seismic_Load_GUI.py --self-test
"""

from __future__ import annotations

import math
import re
import sys
import tkinter as tk
from decimal import Decimal, ROUND_HALF_UP
from tkinter import messagebox, ttk
from typing import Callable, Iterable, List, Sequence

# #
# keep the names the same as the spreadsheet where possible because changing them makes checking the input and output harder and can cause mistakes when comparing columns, some of these names are not very good but they are left as is for that reason
# #
# # #
# layout sizes below are mostly fixed values to make the tables fit on the screen, there is no calculation reason for most of these values and they may need to be changed if the window layout is changed later
#


APP_TITLE = "STAAD.Pro Seismic Input Spreadsheet"

# basic black and white colors so seismic and wind gui have the same format
BG = "#FFFFFF"
CARD = "#FFFFFF"
NAVY = "#252525"              # old variable name is kept here but it is just black now
NAVY_2 = "#F5F5F5"
ACCENT = "#5B7F93"
ACCENT_LIGHT = "#F5F5F5"
INPUT_BG = "#FFFFFF"
OUTPUT_BG = "#FFFFFF"
GRID = "#BFC5C9"
TEXT = "#252525"
MUTED = "#687078"
DANGER = "#8A3B3B"
SUCCESS = "#496B58"
WHITE = "#FFFFFF"
BUTTON_BG = "#FFFFFF"
BUTTON_ACTIVE = "#F2F2F2"

NODE_HEADERS = ("Node", "X (m)", "Y (m)", "Z (m)")
REACTION_HEADERS = ("Node", "L/C", "Fx (kN)", "Fy (kN)", "Fz (kN)", "Mx (kN-m)", "My (kN-m)", "Mz (kN-m)")

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

SAMPLE_NODES = []
SAMPLE_REACTIONS_NO_PAD = []
SAMPLE_REACTIONS_PAD = []


# #
# use the first font from this list that exists on the computer, this only changes how the gui looks
# # #

def _first_font(root: tk.Misc) -> str:
    families = set(root.tk.call("font", "families"))
    for name in ("Aptos", "Segoe UI", "Calibri", "Arial"):
        if name in families:
            return name
    return "TkDefaultFont"


# pasted table values are text so convert the values here and stop if a value that should be a number is not valid, this prevents using bad values later in the seismic calculation
# #

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


# use excel style rounding here because python round can give a diffrent result at the half value, Decimal is used so the final seismic force stays the same as the spreadsheet
# #
#

def _excel_round_3(value: float) -> str:
    """Excel-like ROUND(value,3), rendered the way text concatenation normally appears."""
    d = Decimal(str(value)).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s or "0"


# #
# clipboard parser checks tabs first because the reaction load case name can contain spaces, if there are no tabs it also checks normal whitespace for simpler pasted data
# # #

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
            # whitespace parsing is mainly for the node table, reaction rows should normally be copied with tabs because the load case name can have spaces in it
            parts = re.split(r"\s+", raw_line.strip())

        try:
            float(parts[0])
        except (ValueError, IndexError):
            # if the row does not start with a node number it is treated as a heading and skipped
            continue

        if len(parts) < expected_cols:
            parts.extend([""] * (expected_cols - len(parts)))
        elif len(parts) > expected_cols:
            # reaction load case text can contain spaces so rebuild that field using all values between the node number and the six force and moment values at the end
            if expected_cols == 8 and len(parts) >= 9:
                # the last six values are the forces and moments, anything in the middle is kept as the load case text even if it was split into more then one part
                parts = [parts[0], " ".join(parts[1:-6]), *parts[-6:]]
            else:
                parts = parts[:expected_cols]
        rows.append(parts[:expected_cols])

    if not rows:
        raise ValueError("No numeric data rows were found in the clipboard.")
    return rows


# make the node lookup before doing the calc so node elevation can be found from the node number without searching the full table each time
# #

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


# reaction nodes are checked for duplicates seperately because duplicate rows can mean more then one load case block was pasted and this would count the same support weight more then once
# #
# #

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


# wrap long staad output lines at about this length so the output is readable and close to how the spreadsheet writes the command lines, this value is only formatting
# # #

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


# #
# main seismic calculation is kept in the same function so the order can be compared to the spreadsheet, each row dictionary keeps the node elevation weight and force values used in the next steps
# #
# # #
# with pad uses the same main calc but the pad rows are removed after the adjusted elevation is calculated, do not remove them before this point because the excluded weight and total weight will not match the workbook
#

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

    # keep the orginal row order when two nodes have the same elevation so the final staad output stays in a consistant order
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


# #
# table cells are editied by placing an entry box over the treeview cell because treeview does not have direct cell editing, after the edit the entry is removed and the table value is updated
# #

class EditableTree(tk.Frame):
    def __init__(
        self,
        master,
        title: str,
        headers: Sequence[str],
        widths: Sequence[int],
        rows_visible: int,
        on_change: Callable[[], None],
        font_name: str,
    ):
        super().__init__(master, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        self.headers = tuple(headers)
        self.on_change = on_change
        self.font_name = font_name
        self._editor = None

        tk.Label(
            self,
            text=title,
            bg=WHITE,
            fg=TEXT,
            font=(font_name, 9, "bold"),
            relief="solid",
            bd=1,
            pady=3,
        ).pack(fill="x")

        controls = tk.Frame(self, bg=WHITE)
        controls.pack(fill="x", padx=2, pady=(2, 1))
        tk.Label(
            controls,
            text="INPUT",
            bg=ACCENT_LIGHT,
            fg=TEXT,
            font=(font_name, 8, "bold"),
            bd=1,
            relief="solid",
            padx=5,
        ).pack(side="left")
        tk.Button(
            controls,
            text="Paste Table",
            command=self.paste_from_clipboard,
            font=(font_name, 8),
            bg=BUTTON_BG,
            fg=TEXT,
            activebackground=BUTTON_ACTIVE,
            activeforeground=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=6,
            pady=1,
        ).pack(side="left", padx=(5, 2))
        tk.Button(
            controls,
            text="Add Row",
            command=self.add_blank_row,
            font=(font_name, 8),
            bg=BUTTON_BG,
            fg=TEXT,
            activebackground=BUTTON_ACTIVE,
            activeforeground=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=6,
            pady=1,
        ).pack(side="left", padx=2)
        tk.Button(
            controls,
            text="Delete Selected",
            command=self.delete_selected,
            font=(font_name, 8),
            bg=BUTTON_BG,
            fg=TEXT,
            activebackground=BUTTON_ACTIVE,
            activeforeground=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=6,
            pady=1,
        ).pack(side="left", padx=2)
        tk.Button(
            controls,
            text="Clear",
            command=self.clear_user,
            font=(font_name, 8),
            bg=BUTTON_BG,
            fg=TEXT,
            activebackground=BUTTON_ACTIVE,
            activeforeground=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=6,
            pady=1,
        ).pack(side="right", padx=2)

        grid_frame = tk.Frame(self, bg=WHITE)
        grid_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            grid_frame,
            columns=self.headers,
            show="headings",
            height=rows_visible,
            style="Data.Treeview",
            selectmode="extended",
        )
        for header, width in zip(self.headers, widths):
            self.tree.heading(header, text=header, anchor="center")
            self.tree.column(header, width=width, minwidth=50, stretch=False, anchor="center")

        ybar = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        xbar = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        grid_frame.grid_rowconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._begin_edit)
        self.tree.bind("<Control-v>", lambda _e: self.paste_from_clipboard())
        self.tree.bind("<Delete>", lambda _e: self.delete_selected())

    def set_rows(self, rows: Iterable[Sequence[object]], *, notify: bool = False):
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            vals = ["" if v is None else v for v in row]
            self.tree.insert("", "end", values=vals)
        if notify:
            self.on_change()

    def get_rows(self) -> List[List[object]]:
        return [list(self.tree.item(item, "values")) for item in self.tree.get_children()]

    def add_blank_row(self):
        item = self.tree.insert("", "end", values=[""] * len(self.headers))
        self.tree.selection_set(item)
        self.tree.see(item)
        self.on_change()

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            self.tree.delete(item)
        self.on_change()

    def clear_user(self):
        if self.tree.get_children():
            self.tree.delete(*self.tree.get_children())
            self.on_change()

    def paste_from_clipboard(self):
        try:
            rows = _parse_clipboard_table(self.clipboard_get(), len(self.headers))
        except Exception as exc:
            messagebox.showerror("Paste Table", str(exc), parent=self.winfo_toplevel())
            return
        self.set_rows(rows, notify=True)

    def _begin_edit(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return
        col_index = int(col[1:]) - 1
        bbox = self.tree.bbox(item, col)
        if not bbox:
            return
        x, y, width, height = bbox
        values = list(self.tree.item(item, "values"))
        current = values[col_index] if col_index < len(values) else ""

        if self._editor is not None:
            self._editor.destroy()

        editor = tk.Entry(
            self.tree,
            font=(self.font_name, 9),
            bg=INPUT_BG,
            fg=TEXT,
            relief="solid",
            bd=1,
            justify="center",
        )
        editor.insert(0, current)
        editor.select_range(0, "end")
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        self._editor = editor

        done = {"value": False}

        def commit(_event=None):
            if done["value"]:
                return
            done["value"] = True
            new_values = list(self.tree.item(item, "values"))
            while len(new_values) < len(self.headers):
                new_values.append("")
            new_values[col_index] = editor.get()
            self.tree.item(item, values=new_values)
            editor.destroy()
            self._editor = None
            self.on_change()

        def cancel(_event=None):
            if done["value"]:
                return
            done["value"] = True
            editor.destroy()
            self._editor = None

        editor.bind("<Return>", commit)
        editor.bind("<Tab>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", cancel)


# output panel setup is kept here so the x and z output boxes use the same text size buttons and border settings
# # #

class OutputPanel(tk.Frame):
    def __init__(self, master, title: str, font_name: str):
        super().__init__(master, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        self.font_name = font_name

        tk.Label(
            self,
            text=title,
            bg=WHITE,
            fg=TEXT,
            font=(font_name, 9, "bold"),
            relief="solid",
            bd=1,
            pady=3,
        ).pack(fill="x")

        bar = tk.Frame(self, bg=WHITE)
        bar.pack(fill="x", padx=2, pady=2)
        tk.Label(
            bar,
            text="FORMATTED OUTPUT",
            bg=WHITE,
            fg=TEXT,
            font=(font_name, 8, "bold"),
        ).pack(side="left")
        tk.Button(
            bar,
            text="Copy Output",
            command=self.copy_all,
            font=(font_name, 8),
            bg=BUTTON_BG,
            fg=TEXT,
            activebackground=BUTTON_ACTIVE,
            activeforeground=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
            padx=8,
            pady=1,
        ).pack(side="right")

        body = tk.Frame(self, bg=WHITE)
        body.pack(fill="both", expand=True)
        self.text = tk.Text(
            body,
            height=11,
            wrap="none",
            font=("Consolas", 9),
            bg=OUTPUT_BG,
            fg=TEXT,
            relief="flat",
            padx=5,
            pady=5,
        )
        ybar = ttk.Scrollbar(body, orient="vertical", command=self.text.yview)
        xbar = ttk.Scrollbar(body, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)
        self.text.configure(state="disabled")

    def set_lines(self, lines: Sequence[str]):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", "\n".join(lines))
        self.text.configure(state="disabled")

    def set_error(self, error: str):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", "INPUT ERROR\n\n" + error)
        self.text.configure(state="disabled")

    def copy_all(self):
        value = self.text.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(value)


# #
# canvas and frame are used together to allow vertical scrolling for the page
# #

class ScrollablePage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=BG)
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hbar = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vbar.set, xscrollcommand=self.hbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.inner = tk.Frame(self.canvas, bg=BG)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._sync_scrollregion)
        self.canvas.bind("<Configure>", self._fit_width)

        self.canvas.bind_all("<MouseWheel>", self._mousewheel, add="+")
        self.canvas.bind_all("<Shift-MouseWheel>", self._shift_mousewheel, add="+")

    def _sync_scrollregion(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _fit_width(self, event):
        # horizontal scrolling is avaliable when the content is wider then the window,
        # but the frame is also expanded when there is enough width so the content does not stay narrow
        req = self.inner.winfo_reqwidth()
        self.canvas.itemconfigure(self.window_id, width=max(event.width, req))

    def _mousewheel(self, event):
        widget = self.winfo_containing(event.x_root, event.y_root)
        if widget is None or not self._is_descendant(widget):
            return
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _shift_mousewheel(self, event):
        widget = self.winfo_containing(event.x_root, event.y_root)
        if widget is None or not self._is_descendant(widget):
            return
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _is_descendant(self, widget):
        current = widget
        while current is not None:
            if current == self:
                return True
            current = getattr(current, "master", None)
        return False


# one page is created for each workbook mode, most of the page setup is left together because the same input blocks and tables need to be arranged in one place
# #
# # #

class ModePage:
    def __init__(self, app: "SeismicApp", parent: tk.Frame, *, with_pad: bool):
        self.app = app
        self.parent = parent
        self.with_pad = with_pad
        self._job = None
        self.summary_vars = {key: tk.StringVar(value="—") for key in (
            "coefficient", "total_vertical", "excluded_weight", "effective_weight",
            "vmax", "distributed_sum", "count"
        )}
        self.status = tk.StringVar(value="")

        self._build()
        self._clear_inputs()

    def _build(self):
        p = self.parent
        p.configure(bg=WHITE)

        header = tk.Frame(p, bg=WHITE)
        header.pack(fill="x", padx=18, pady=(14, 10))
        tk.Label(
            header,
            text="NBC 2020 — Seismic Load Generator",
            bg=WHITE,
            fg=TEXT,
            font=(self.app.font_name, 18, "bold"),
            anchor="w",
        ).pack(side="left")
        mode_text = "WITH PAD" if self.with_pad else "NO PAD"
        tk.Label(
            header,
            text=mode_text,
            bg=ACCENT_LIGHT,
            fg=TEXT,
            relief="solid",
            bd=1,
            font=(self.app.font_name, 9, "bold"),
            padx=10,
            pady=4,
        ).pack(side="left", padx=(12, 0))
        tk.Button(
            header,
            text="Instructions",
            command=self.app.show_instructions,
            bg=BUTTON_BG,
            fg=TEXT,
            activebackground=BUTTON_ACTIVE,
            activeforeground=TEXT,
            relief="solid",
            bd=1,
            highlightthickness=0,
            font=(self.app.font_name, 9, "bold"),
            padx=10,
            pady=4,
        ).pack(side="right")

        # instructions were on the old intro sheet, keeping it at the top now because it is easier to see while doing the inputs
        inst = tk.Frame(p, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        inst.pack(fill="x", padx=18, pady=(0, 10))
        tk.Label(
            inst,
            text="USER INSTRUCTIONS",
            bg=WHITE,
            fg=TEXT,
            font=(self.app.font_name, 9, "bold"),
            anchor="w",
            padx=6,
            pady=3,
        ).pack(fill="x")
        tk.Label(
            inst,
            text=(
                "Create a duplicate STAAD file and define pin supports. Run the reaction load case used by the calculation. "
                "For this workbook's calculation sheets, paste the 1.0D + 0.25S reaction set. Copy all reaction forces, "
                "then copy all nodes from the Geometry window. Base elevation must be adjusted to 0; if it is not 0, use "
                "the 'Height adjustment STAAD' input. Once complete, verify loading and copy the generated JOINT LOAD text "
                "into the STAAD input file."
            ),
            bg=WHITE,
            fg=TEXT,
            font=(self.app.font_name, 9),
            anchor="w",
            justify="left",
            wraplength=1320,
            padx=6,
            pady=5,
        ).pack(fill="x")

        body = tk.Frame(p, bg=WHITE)
        body.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        # top information and seismic input row
        top = tk.Frame(body, bg=WHITE)
        top.pack(fill="x", pady=(0, 10))

        seismic = tk.Frame(top, bg=CARD, highlightbackground=GRID, highlightthickness=1)
        seismic.pack(side="left", anchor="n", padx=(0, 10))
        tk.Label(
            seismic,
            text="NBC 2020 Seismic Hazard / Design Inputs",
            bg=ACCENT_LIGHT,
            fg=TEXT,
            font=(self.app.font_name, 9, "bold"),
            anchor="w",
            padx=7,
            pady=4,
            relief="solid",
            bd=1,
        ).grid(row=0, column=0, columnspan=4, sticky="ew")

        fields = [
            ("Latitude", "latitude", 1, 0), ("Longitude", "longitude", 1, 2),
            ("Sa (0.2)", "sa02", 2, 0), ("Sa (0.5)", "sa05", 2, 2),
            ("Sa (1.0)", "sa10", 3, 0), ("Sa (2.0)", "sa20", 3, 2),
            ("Sa (5.0)", "sa50", 4, 0), ("IE", "ie", 4, 2),
            ("Site Class", "site_class", 5, 0), ("Snow Load (kPa)", "snow", 5, 2),
            ("Rd", "rd", 6, 0), ("Ro", "ro", 6, 2),
            ("Height adjustment STAAD", "height_adjustment", 7, 0),
        ]
        for label, key, row, col in fields:
            tk.Label(
                seismic, text=label, bg=CARD, fg=TEXT,
                font=(self.app.font_name, 9), anchor="w", padx=7, pady=4
            ).grid(row=row, column=col, sticky="ew")
            entry = tk.Entry(
                seismic,
                textvariable=self.app.vars[key],
                bg=INPUT_BG,
                fg=TEXT,
                relief="solid",
                bd=1,
                width=15 if key != "site_class" else 8,
                font=(self.app.font_name, 9),
            )
            entry.grid(row=row, column=col+1, sticky="ew", padx=(0, 8), pady=2)
            entry.bind("<KeyRelease>", lambda _e: self.app.schedule_all())

        # keep site class and the reference values editable the same as the workbook,
        # the final vmax equation uses Sa(0.2) IE Rd and Ro directly but the other fields are still shown for checking and input record
        seismic.grid_columnconfigure(1, weight=1)
        seismic.grid_columnconfigure(3, weight=1)

        summary = tk.Frame(top, bg=CARD, highlightbackground=GRID, highlightthickness=1)
        summary.pack(side="left", anchor="n", fill="x", expand=True)
        tk.Label(
            summary,
            text="Calculated Output Summary",
            bg=ACCENT_LIGHT,
            fg=TEXT,
            font=(self.app.font_name, 9, "bold"),
            anchor="w",
            padx=7,
            pady=4,
            relief="solid",
            bd=1,
        ).grid(row=0, column=0, columnspan=4, sticky="ew")
        summary_rows = [
            ("Vmax coefficient", "coefficient", ""),
            ("Total vertical reaction", "total_vertical", " kN"),
            ("Pad/mat weight excluded", "excluded_weight", " kN"),
            ("Effective Seismic Weight Operating", "effective_weight", " kN"),
            ("Vmax", "vmax", " kN"),
            ("Distributed load total", "distributed_sum", " kN"),
            ("Joint loads", "count", ""),
        ]
        for i, (label, key, suffix) in enumerate(summary_rows, start=1):
            tk.Label(
                summary, text=label, bg=CARD, fg=TEXT,
                font=(self.app.font_name, 9), anchor="w", padx=7, pady=4
            ).grid(row=i, column=0, sticky="ew")
            tk.Label(
                summary, textvariable=self.summary_vars[key], bg=OUTPUT_BG, fg=TEXT,
                font=(self.app.font_name, 9, "bold"), anchor="e",
                padx=8, pady=4, relief="solid", bd=1
            ).grid(row=i, column=1, sticky="ew", padx=(4, 0), pady=2)
            tk.Label(
                summary, text=suffix, bg=CARD, fg=MUTED,
                font=(self.app.font_name, 9), anchor="w", padx=4
            ).grid(row=i, column=2, sticky="w")
        summary.grid_columnconfigure(0, weight=1)
        summary.grid_columnconfigure(1, weight=0)

        note_text = (
            "PAD mode: reactions at adjusted elevation ≤ 0 are treated as pad/mat reactions and excluded."
            if self.with_pad else
            "NO PAD mode: every pasted reaction node is included in the seismic weight."
        )
        tk.Label(
            summary,
            text=note_text,
            bg=CARD,
            fg=MUTED,
            justify="left",
            wraplength=430,
            font=(self.app.font_name, 8),
            padx=7,
            pady=7,
        ).grid(row=8, column=0, columnspan=4, sticky="ew")

        # input tables for the first and second spreadsheet steps
        tables = tk.Frame(body, bg=WHITE)
        tables.pack(fill="x", pady=(0, 10))

        self.node_table = EditableTree(
            tables,
            "1. Paste Nodes from Analytical Modeling, Geometry Tab",
            NODE_HEADERS,
            (75, 90, 90, 90),
            14,
            self.schedule_recalc,
            self.app.font_name,
        )
        self.reaction_table = EditableTree(
            tables,
            "2. Paste Reactions from Postprocessing for 1.0D + 0.25S Load Combination",
            REACTION_HEADERS,
            (70, 130, 82, 82, 82, 90, 90, 90),
            14,
            self.schedule_recalc,
            self.app.font_name,
        )
        self.node_table.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.reaction_table.grid(row=0, column=1, sticky="nsew")
        tables.grid_columnconfigure(0, weight=1)
        tables.grid_columnconfigure(1, weight=2)

        out_title = tk.Label(
            body,
            text="3. Copy the following into the STAAD Command File under the correct seismic load cases",
            bg=ACCENT_LIGHT,
            fg=TEXT,
            anchor="w",
            padx=9,
            pady=6,
            font=(self.app.font_name, 9, "bold"),
        )
        out_title.pack(fill="x", pady=(0, 6))

        outputs = tk.Frame(body, bg=WHITE)
        outputs.pack(fill="x")
        self.fz_output = OutputPanel(outputs, "STAAD JOINT LOAD — FZ", self.app.font_name)
        self.fx_output = OutputPanel(outputs, "STAAD JOINT LOAD — FX", self.app.font_name)
        self.fz_output.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.fx_output.grid(row=0, column=1, sticky="nsew")
        outputs.grid_columnconfigure(0, weight=1)
        outputs.grid_columnconfigure(1, weight=1)

        tk.Label(
            body,
            textvariable=self.status,
            bg=WHITE,
            fg=MUTED,
            anchor="w",
            font=(self.app.font_name, 8),
            pady=6,
        ).pack(fill="x")

    def _clear_inputs(self):
        self.node_table.set_rows([])
        self.reaction_table.set_rows([])

    def schedule_recalc(self):
        if self._job is not None:
            try:
                self.app.after_cancel(self._job)
            except tk.TclError:
                pass
        self._job = self.app.after(220, self.recalculate)

    # #
    # calculation runs automatic after a table or input is changed, small delay is used when pasting large tables so it does not run the full calc after every single row is added
    #
    def recalculate(self):
        self._job = None
        try:
            result = calculate_seismic(
                self.node_table.get_rows(),
                self.reaction_table.get_rows(),
                sa02=_as_number(self.app.vars["sa02"].get(), field="Sa (0.2)"),
                ie=_as_number(self.app.vars["ie"].get(), field="IE"),
                rd=_as_number(self.app.vars["rd"].get(), field="Rd"),
                ro=_as_number(self.app.vars["ro"].get(), field="Ro"),
                height_adjustment=_as_number(self.app.vars["height_adjustment"].get(), field="Height adjustment"),
                with_pad=self.with_pad,
            )

            self.summary_vars["coefficient"].set(f'{result["coefficient"]:.9f}')
            self.summary_vars["total_vertical"].set(f'{result["total_vertical"]:.3f}')
            self.summary_vars["excluded_weight"].set(f'{result["excluded_weight"]:.3f}')
            self.summary_vars["effective_weight"].set(f'{result["effective_weight"]:.3f}')
            self.summary_vars["vmax"].set(f'{result["vmax"]:.3f}')
            self.summary_vars["distributed_sum"].set(f'{result["distributed_sum"]:.3f}')
            self.summary_vars["count"].set(str(result["count"]))
            self.fz_output.set_lines(result["fz_lines"])
            self.fx_output.set_lines(result["fx_lines"])
            self.status.set(
                f'Auto-calculated • Σ(Wi·hi) = {result["wihi_sum"]:.3f} kN·m • '
                f'{result["count"]} matched reaction nodes'
            )
        except Exception as exc:
            for var in self.summary_vars.values():
                var.set("—")
            self.fz_output.set_error(str(exc))
            self.fx_output.set_error(str(exc))
            self.status.set(str(exc))


# #
# main window tabs and shared formatting are set here, most widths are fixed to fit the current screen layout so check both tabs if any table width is changed because one tab can clip before the other one does
# #
#
class SeismicApp(tk.Tk):
    def __init__(self, *, smoke_test: bool = False):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1600x930")
        self.minsize(1180, 740)
        self.configure(bg=WHITE)
        self.font_name = _first_font(self)
        self.vars = {key: tk.StringVar(value=value) for key, value in DEFAULTS.items()}
        self._all_job = None

        self._configure_styles()
        self._make_menu()

        notebook = ttk.Notebook(self, style="Main.TNotebook")
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        no_pad_scroll = ScrollablePage(notebook)
        pad_scroll = ScrollablePage(notebook)
        notebook.add(no_pad_scroll, text="NBC2020 NO PAD")
        notebook.add(pad_scroll, text="NBC2020 PAD CALC")

        self.no_pad = ModePage(self, no_pad_scroll.inner, with_pad=False)
        self.pad = ModePage(self, pad_scroll.inner, with_pad=True)
        self.pages = [self.no_pad, self.pad]

        self.after(50, self.recalculate_all)
        if smoke_test:
            self.after(300, self.destroy)

    # #
    # gui formatting is set here, it is kept in this file so the program can be sent as one python file and does not need a seperate theme file
    # # #
    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # windows ttk theme can change the button border and fill so the style is set here to keep the same plain black and white look on different computers
        # #
        for scrollbar_style in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            style.configure(
                scrollbar_style,
                background=WHITE,
                troughcolor=WHITE,
                bordercolor=GRID,
                lightcolor=WHITE,
                darkcolor=WHITE,
                arrowcolor=TEXT,
                relief="flat",
            )
            style.map(
                scrollbar_style,
                background=[("active", WHITE), ("pressed", WHITE)],
                arrowcolor=[("active", TEXT), ("pressed", TEXT)],
            )

        style.configure(
            "Data.Treeview",
            background=WHITE,
            fieldbackground=WHITE,
            foreground=TEXT,
            bordercolor=GRID,
            lightcolor=WHITE,
            darkcolor=WHITE,
            rowheight=22,
            font=(self.font_name, 9),
        )
        style.configure(
            "Data.Treeview.Heading",
            background=ACCENT_LIGHT,
            foreground=TEXT,
            bordercolor=GRID,
            lightcolor=WHITE,
            darkcolor=WHITE,
            relief="flat",
            font=(self.font_name, 8, "bold"),
            padding=(4, 3),
        )
        style.map(
            "Data.Treeview",
            background=[("selected", "#5B7F93")],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            "Main.TNotebook",
            background=WHITE,
            borderwidth=0,
        )
        style.configure(
            "Main.TNotebook.Tab",
            background=WHITE,
            foreground=TEXT,
            padding=(14, 7),
            font=(self.font_name, 9, "bold"),
        )
        style.map(
            "Main.TNotebook.Tab",
            background=[("selected", "#5B7F93")],
            foreground=[("selected", "#FFFFFF")],
        )

    def _make_menu(self):
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Exit", command=self.destroy)
        menu.add_cascade(label="File", menu=file_menu)
        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label="Workbook Instructions", command=self.show_instructions)
        menu.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menu)

    def schedule_all(self):
        if self._all_job is not None:
            try:
                self.after_cancel(self._all_job)
            except tk.TclError:
                pass
        self._all_job = self.after(220, self.recalculate_all)

    # # #
    # both tabs calculate seperately because each tab has its own pasted node and reaction tables even though most of the input values start with the same defaults
    # #
    def recalculate_all(self):
        self._all_job = None
        for page in self.pages:
            page.recalculate()

    def show_instructions(self):
        win = tk.Toplevel(self)
        win.title("How to use the STAAD.Pro Seismic Input Spreadsheet")
        win.geometry("900x650")
        win.minsize(650, 450)
        win.configure(bg=WHITE)
        win.transient(self)

        outer = tk.Frame(win, bg=WHITE)
        outer.pack(fill="both", expand=True, padx=18, pady=18)

        tk.Label(
            outer,
            text="How to use the STAAD.Pro Seismic Input Spreadsheet",
            bg=WHITE,
            fg=TEXT,
            font=(self.font_name, 17, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        text = tk.Text(
            outer,
            wrap="word",
            bg=CARD,
            fg=TEXT,
            relief="solid",
            bd=1,
            font=(self.font_name, 10),
            padx=14,
            pady=12,
        )
        ybar = ttk.Scrollbar(outer, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=ybar.set)
        text.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")

        instructions = (
            "Spreadsheet Overview\n\n"
            "This spreadsheet was designed to simplify the seismic load generation process for STAAD modeling. "
            "It takes user input and converts it into a format that can be pasted directly into the STAAD.Pro command file.\n\n"

            "Assumptions/Notes\n\n"
            "• This spreadsheet follows the equivalent static force procedure described in section 4.1.8.11 in NBCC 2020. "
            "It takes into account the maximum base shear (Vmax) to be conservative.\n\n"
            "• This spreadsheet is provided for use at the user's discretion. The user is responsible for ensuring the "
            "accuracy of all inputs, results, and design assumptions.\n\n"
            "• The workbook provides another calculation path for pad weight; this GUI keeps the calculations hidden and "
            "automatically removes pad/mat reactions in WITH PAD mode based on adjusted elevation.\n\n"

            "User Guide\n\n"
            "Create a duplicate STAAD file and define pin supports. Run the reaction load case used by the calculation. "
            "For this workbook's calculation sheets, paste the 1.0D + 0.25S reaction set. Copy all reaction forces, then "
            "copy all nodes from the Geometry window. Base elevation must be adjusted to 0; if it is not 0, use the "
            "'Height adjustment STAAD' input. Once complete, verify loading and copy the generated JOINT LOAD text into "
            "the STAAD input file.\n\n"

            "References\n\n"
            "Codes:\n"
            "National Building Code of Canada (NBCC 2020)\n"
            "American Institute of Steel Construction (AISC)\n\n"

            "Fixed GUI behavior\n\n"
            "• Node coordinates and reactions are matched by node number instead of by pasted row position.\n"
            "• WITH PAD mode excludes adjusted elevations ≤ 0 when calculating seismic weight.\n"
            "• Effective seismic weight is calculated live from the pasted reactions rather than using a hard-coded total.\n"
            "• Every generated node is included in the output; the workbook's final-row drop is not reproduced.\n"
            "• Recalculation is automatic."
        )
        text.insert("1.0", instructions)
        text.configure(state="disabled")


# basic self test for the seismic calculation, this should be run after any code changes to check that the expected result is still the same
# # #

def self_test() -> int:
    print("No embedded project dataset is included in this public version.")
    return 0


def main() -> int:
    if "--self-test" in sys.argv:
        return self_test()
    app = SeismicApp(smoke_test="--smoke-test" in sys.argv)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
