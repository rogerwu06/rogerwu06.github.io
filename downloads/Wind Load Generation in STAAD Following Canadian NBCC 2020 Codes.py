#!/usr/bin/env python3
"""
STAAD.Pro Wind Load X & Z Generator
-----------------------------------
Standalone desktop implementation of the wind-load calculation workflow.

- Same three input tables and wind parameters.
- Same formatted STAAD output (GZ and GX).
- Intermediate calculation tables are intentionally hidden from the GUI.
- Uses only the Python standard library.

Run:
    python STAAD_Wind_Load_GUI.py

Self-test:
    python STAAD_Wind_Load_GUI.py --self-test
"""

from __future__ import annotations

import base64
import math
import re
import sys
import tkinter as tk
from collections import OrderedDict
from tkinter import messagebox, ttk
from typing import Callable, Iterable, List, Sequence

####
# this part
#
##
# input stuff
# #


APP_TITLE = "STAAD.Pro Wind Input Spreadsheet"
PALE_YELLOW = "#F5F5F5"     # keep this
SECTION_GRAY = "#F3F3F3"
TITLE_BLUE = "#252525"          # main part
RED = "#8A3B3B"                 # needed here
GRID = "#BFC5C9"
WHITE = "#FFFFFF"
LIGHT_INPUT = "#FFFFFF"
LIGHT_OUTPUT = "#FFFFFF"
TEXT = "#252525"
MUTED_TEXT = "#252525"
ACCENT = "#5B7F93"
ACCENT_DARK = "#456A7E"
BUTTON_BG = "#FFFFFF"
BUTTON_ACTIVE = "#F2F2F2"

NODE_HEADERS = ("NODE", "X", "Y", "Z")
MEMBER_HEADERS = ("BEAM", "NODEA", "NODE B", "PROPERTY DEFINE", "MATERIAL", "BETA", "LENGTH")
SECTION_HEADERS = ("prop", "name", "ax", "d", "bf", "tf", "tw", "in", "iy", "ix")

# this part
SAMPLE_NODES = []
SAMPLE_MEMBERS = []
SAMPLE_SECTIONS = []

EXPECTED_GZ = []
EXPECTED_GX = []




#
# calc part
####
# #

def _first_font(root: tk.Misc, preferred=("Calibri", "Aptos", "Arial", "Segoe UI")) -> str:
    try:
        families = set(root.tk.call("font", "families"))
        for name in preferred:
            if name in families:
                return name
    except Exception:
        pass
    return "TkDefaultFont"


# keep this
##

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


# needed here
# #

#
# calc part
####
def _general_2dp(value: float) -> str:
    #
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _roundup_2(value: float) -> float:
    # keep this
    return math.ceil(value * 100.0 - 1e-12) / 100.0


# main part
####
# #

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


#
# calc part
####
# #
# keep this
##

####
# this part
#
def calculate_wind_outputs(
    node_rows: Sequence[Sequence[object]],
    member_rows: Sequence[Sequence[object]],
    section_rows: Sequence[Sequence[object]],
    height_adjustment: float,
    q: float,
    cg: float,
):
    """Return (max_height, formatted_gz_lines, formatted_gx_lines)."""

    # calc part
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

    # input stuff
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


    def direction_output(direction: str) -> List[str]:
        code = "GZ" if direction == "Z" else "GX"
        individual = []

        for beam, node_a, node_b, prop, _material, _beta, _length in members:
            xa, ya, za = nodes[node_a]
            xb, yb, zb = nodes[node_b]

            # output stuff
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

            # keep this
            fn = ce * cg * k * 2.1 * q
            ft = fn
            kn_per_m_fn = fn * d / 1000.0
            kn_per_m_ft = ft * bf / 1000.0
            load = _roundup_2(max(kn_per_m_fn, kn_per_m_ft))
            individual.append((beam, f"UNI {code} {_general_2dp(load)}"))

        # main part
        groups = OrderedDict()
        for beam, suffix in individual:
            groups.setdefault(suffix, []).append(str(beam))
        return [" ".join(beams) + " " + suffix for suffix, beams in groups.items()]

    return max_height, direction_output("Z"), direction_output("X")


# needed here
# #
#

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

        # calc part
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


####
#
#
##
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

        title_label = tk.Label(
            self,
            text=title,
            bg=WHITE,
            fg=TEXT,
            font=(font_name, 9, "bold"),
            relief="solid",
            bd=1,
            pady=3,
        )
        title_label.pack(fill="x")

        controls = tk.Frame(self, bg=WHITE)
        controls.pack(fill="x", padx=2, pady=(2, 1))
        tk.Label(
            controls,
            text="INPUT",
            bg=PALE_YELLOW,
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
            style="Excel.Treeview",
            selectmode="extended",
        )
        for header, width in zip(self.headers, widths):
            self.tree.heading(header, text=header, anchor="center")
            self.tree.column(header, width=width, minwidth=38, stretch=False, anchor="center")

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
            values = ["" if value is None else value for value in row]
            self.tree.insert("", "end", values=values)
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
            text = self.clipboard_get()
            rows = _parse_clipboard_table(text, len(self.headers))
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
            bg=LIGHT_INPUT,
            relief="solid",
            bd=1,
            justify="center",
        )
        editor.insert(0, current)
        editor.select_range(0, "end")
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        self._editor = editor

        committed = {"done": False}

        def commit(_event=None):
            if committed["done"]:
                return
            committed["done"] = True
            try:
                new_values = list(self.tree.item(item, "values"))
                while len(new_values) < len(self.headers):
                    new_values.append("")
                new_values[col_index] = editor.get()
                self.tree.item(item, values=new_values)
            finally:
                editor.destroy()
                self._editor = None
            self.on_change()

        def cancel(_event=None):
            if committed["done"]:
                return
            committed["done"] = True
            editor.destroy()
            self._editor = None

        editor.bind("<Return>", commit)
        editor.bind("<Tab>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", cancel)


# needed here
# #

#
# calc part
####
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

        text_frame = tk.Frame(self, bg=WHITE)
        text_frame.pack(fill="both", expand=True)
        self.text = tk.Text(
            text_frame,
            height=11,
            wrap="none",
            font=("Consolas", 9),
            bg=LIGHT_OUTPUT,
            fg=TEXT,
            relief="flat",
            padx=5,
            pady=5,
        )
        ybar = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        xbar = ttk.Scrollbar(text_frame, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        self.text.configure(state="disabled")

    def set_lines(self, lines: Sequence[str]):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        if lines:
            self.text.insert("1.0", "\n".join(lines))
        self.text.configure(state="disabled")

    def set_error(self, text: str):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", text)
        self.text.configure(state="disabled")

    def copy_all(self):
        content = self.text.get("1.0", "end-1c")
        if not content.strip():
            return
        self.clipboard_clear()
        self.clipboard_append(content)
        self.update_idletasks()


# #
# keep this
##

class ScrollablePage(tk.Frame):
    def __init__(self, master):
        super().__init__(master, bg=WHITE)
        self.canvas = tk.Canvas(self, bg=WHITE, highlightthickness=0)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.inner = tk.Frame(self.canvas, bg=WHITE)
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._sync_scrollregion)
        self.canvas.bind("<Configure>", self._fit_inner_width)

        self.canvas.bind_all("<MouseWheel>", self._mousewheel, add="+")

    def _sync_scrollregion(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _fit_inner_width(self, event):
        # needed here
        # this part
        self.canvas.itemconfigure(self.window_id, width=max(1, event.width))

    def _mousewheel(self, event):
        # leave this
        widget = self.winfo_containing(event.x_root, event.y_root)
        if widget is None or not self._is_descendant(widget):
            return
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _is_descendant(self, widget):
        current = widget
        while current is not None:
            if current == self:
                return True
            current = getattr(current, "master", None)
        return False


##
# input stuff
# #
#

class WindLoadApp(tk.Tk):
    def __init__(self, *, smoke_test: bool = False):
        super().__init__()
        self.title(APP_TITLE)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = min(1440, max(1180, screen_w - 40))
        win_h = min(880, max(680, screen_h - 90))
        self.geometry(f"{win_w}x{win_h}")
        self.minsize(1160, 650)
        self.configure(bg=WHITE)
        self.font_name = _first_font(self)
        self._recalc_job = None
        self._building = True

        self._configure_styles()
        self._make_menu()

        self.fixed_tab = ScrollablePage(self)
        self.fixed_tab.pack(fill="both", expand=True)
        self._build_fixed(self.fixed_tab.inner)
        self._clear_inputs()
        self._building = False
        self.recalculate()

        if smoke_test:
            self.after(150, self.destroy)

    ##
    # needed here
    # #
    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # leave this
        ##
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
            "Excel.Treeview",
            background=WHITE,
            fieldbackground=WHITE,
            foreground=TEXT,
            bordercolor=GRID,
            lightcolor=WHITE,
            darkcolor=WHITE,
            rowheight=21,
            font=(self.font_name, 9),
        )
        style.configure(
            "Excel.Treeview.Heading",
            background=SECTION_GRAY,
            foreground=TEXT,
            bordercolor=GRID,
            lightcolor=WHITE,
            darkcolor=WHITE,
            relief="flat",
            font=(self.font_name, 9),
            padding=(3, 2),
        )
        style.map(
            "Excel.Treeview",
            background=[("selected", "#5B7F93")],
            foreground=[("selected", "#FFFFFF")],
        )

    def _make_menu(self):
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Exit", command=self.destroy)
        menu.add_cascade(label="File", menu=file_menu)

        self.config(menu=menu)


    def _build_fixed(self, parent):
        parent.configure(bg=WHITE)
        wrap = tk.Frame(parent, bg=WHITE)
        wrap.pack(fill="both", expand=True, padx=12, pady=10)

        #
        inst = tk.Frame(wrap, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        inst.pack(fill="x", pady=(0, 10))
        tk.Label(
            inst,
            text="USER INSTRUCTIONS",
            bg=WHITE,
            fg=TEXT,
            font=(self.font_name, 9, "bold"),
            anchor="w",
            padx=6,
            pady=3,
        ).pack(fill="x")
        tk.Label(
            inst,
            text=(
                "Copy in all node table and beam table. If there are members that should not have wind load they will "
                "need to be removed (create a copy and delete them, only geometry is needed). Extract property table: "
                "view -> tables -> section properties -> angle ST -> click on Prop. Need to select Angle ST to select "
                "all section properties. Wind loads not from beams need to be defined separately (equipment, stairs, "
                "pipe, and cable trays)."
            ),
            bg=WHITE,
            fg=TEXT,
            font=(self.font_name, 9),
            anchor="w",
            justify="left",
            wraplength=1120,
            padx=6,
            pady=5,
        ).pack(fill="x")

        # output stuff
        top = tk.Frame(wrap, bg=WHITE)
        top.pack(fill="x", pady=(0, 12))

        left = tk.Frame(top, bg=WHITE)
        left.pack(side="left", anchor="nw")
        tk.Label(
            left,
            text="*Note: Only displays correct distance if Node 0,0,0, was modeled at grade level…",
            bg=WHITE,
            fg=TEXT,
            font=(self.font_name, 9, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            left,
            text="Put in height difference to make elevation at base 0",
            bg=WHITE,
            fg=TEXT,
            font=(self.font_name, 9, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(2, 7))

        height_box = tk.Frame(left, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        height_box.pack(anchor="w")
        tk.Label(
            height_box,
            text="height adjustment\nSTAAD",
            bg=PALE_YELLOW,
            fg=TEXT,
            font=(self.font_name, 9, "bold"),
            width=18,
            justify="center",
            relief="solid",
            bd=1,
        ).pack(fill="x")
        self.height_adjustment = tk.StringVar(value="")
        h_entry = tk.Entry(
            height_box,
            textvariable=self.height_adjustment,
            bg=LIGHT_INPUT,
            fg=TEXT,
            justify="center",
            font=(self.font_name, 9),
            relief="solid",
            bd=1,
            width=18,
        )
        h_entry.pack(fill="x")
        h_entry.bind("<KeyRelease>", lambda _e: self.schedule_recalc())

        params = tk.Frame(top, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        params.pack(side="right", anchor="ne", padx=(18, 0))
        tk.Label(
            params,
            text="ADD WIND PARAMETERS",
            bg=PALE_YELLOW,
            fg=TEXT,
            font=(self.font_name, 9, "bold"),
            relief="solid",
            bd=1,
            padx=8,
            pady=2,
        ).grid(row=0, column=0, columnspan=2, sticky="ew")

        self.max_height_text = tk.StringVar(value="#N/A")
        self.q_value = tk.StringVar(value="")
        self.cg_value = tk.StringVar(value="")

        param_rows = [
            ("max HEIGHT:", self.max_height_text, True),
            ("", tk.StringVar(value=""), True),
            ("q", self.q_value, False),
            ("Cg", self.cg_value, False),
        ]
        self._param_vars = [v for _, v, _ in param_rows]
        for row_i, (label, var, readonly) in enumerate(param_rows, start=1):
            tk.Label(
                params,
                text=label,
                bg=WHITE,
                fg=TEXT,
                font=(self.font_name, 9, "bold" if row_i == 1 else "normal"),
                anchor="w",
                width=12,
                relief="solid",
                bd=1,
                padx=3,
            ).grid(row=row_i, column=0, sticky="nsew")
            entry = tk.Entry(
                params,
                textvariable=var,
                bg=LIGHT_OUTPUT if readonly else LIGHT_INPUT,
                fg=TEXT,
                justify="right",
                font=(self.font_name, 9, "bold" if row_i == 1 else "normal"),
                relief="solid",
                bd=1,
                width=12,
                state="readonly" if readonly else "normal",
                readonlybackground=LIGHT_OUTPUT,
            )
            entry.grid(row=row_i, column=1, sticky="nsew")
            if not readonly:
                entry.bind("<KeyRelease>", lambda _e: self.schedule_recalc())

        # keep this
        tables = tk.Frame(wrap, bg=WHITE)
        tables.pack(fill="x", pady=(0, 12))

        self.node_table = EditableTree(
            tables,
            "PAST FULL NODE TABLE",
            NODE_HEADERS,
            (48, 50, 50, 50),
            15,
            self.schedule_recalc,
            self.font_name,
        )
        self.member_table = EditableTree(
            tables,
            "PAST FULL MEMBER TABLE",
            MEMBER_HEADERS,
            (44, 45, 45, 76, 90, 44, 52),
            15,
            self.schedule_recalc,
            self.font_name,
        )
        self.section_table = EditableTree(
            tables,
            "SECTION PROPERTIES",
            SECTION_HEADERS,
            (40, 62, 48, 40, 40, 40, 46, 55, 55, 55),
            15,
            self.schedule_recalc,
            self.font_name,
        )

        self.node_table.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        self.member_table.grid(row=0, column=1, sticky="nsew", padx=(0, 7))
        self.section_table.grid(row=0, column=2, sticky="nsew")
        tables.grid_columnconfigure(0, weight=19)
        tables.grid_columnconfigure(1, weight=35)
        tables.grid_columnconfigure(2, weight=46)

        status_bar = tk.Frame(wrap, bg=WHITE)
        status_bar.pack(fill="x", pady=(0, 8))
        self.status_text = tk.StringVar(value="")
        tk.Label(
            status_bar,
            textvariable=self.status_text,
            bg=WHITE,
            fg=MUTED_TEXT,
            font=(self.font_name, 9),
            anchor="e",
        ).pack(side="right", fill="x", expand=True)

        outputs = tk.Frame(wrap, bg=WHITE)
        outputs.pack(fill="x")
        self.gz_output = OutputPanel(outputs, "OUTPUT FOR FZ  —  STAAD GZ", self.font_name)
        self.gx_output = OutputPanel(outputs, "OUTPUT FOR FX  —  STAAD GX", self.font_name)
        self.gz_output.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.gx_output.grid(row=0, column=1, sticky="nsew")
        outputs.grid_columnconfigure(0, weight=1, minsize=0)
        outputs.grid_columnconfigure(1, weight=1, minsize=0)

    def _clear_inputs(self):
        self.height_adjustment.set("")
        self.q_value.set("")
        self.cg_value.set("")
        self.node_table.set_rows([])
        self.member_table.set_rows([])
        self.section_table.set_rows([])


    def schedule_recalc(self):
        if self._building:
            return
        if self._recalc_job is not None:
            try:
                self.after_cancel(self._recalc_job)
            except tk.TclError:
                pass
        self._recalc_job = self.after(250, self.recalculate)

    ##
    # needed here
    # #
    def recalculate(self):
        if self._building:
            return
        self._recalc_job = None
        try:
            height_adjustment = _as_number(self.height_adjustment.get(), field="height adjustment")
            q = _as_number(self.q_value.get(), field="q")
            cg = _as_number(self.cg_value.get(), field="Cg")
            if q < 0:
                raise ValueError("q cannot be negative.")
            if cg < 0:
                raise ValueError("Cg cannot be negative.")

            max_height, gz, gx = calculate_wind_outputs(
                self.node_table.get_rows(),
                self.member_table.get_rows(),
                self.section_table.get_rows(),
                height_adjustment,
                q,
                cg,
            )
            self.max_height_text.set(f"{max_height:g}")
            self.gz_output.set_lines(gz)
            self.gx_output.set_lines(gx)
            self.status_text.set(
                f"Max height: {max_height:g} m   |   GZ groups: {len(gz)}   |   GX groups: {len(gx)}"
            )
        except Exception as exc:
            self.max_height_text.set("#N/A")
            msg = f"INPUT ERROR\n\n{exc}"
            self.gz_output.set_error(msg)
            self.gx_output.set_error(msg)
            self.status_text.set(str(exc))



#
# calc part
####
class SiteStyleWindApp(tk.Tk):
    def __init__(self, *, smoke_test: bool = False):
        super().__init__()
        self.title("Wind Load Generation in STAAD Following Canadian NBCC 2020 Codes")
        self.geometry("1450x860")
        self.minsize(1120, 680)
        self.configure(bg=WHITE)
        self.font_name = _first_font(self)
        self.height_adjustment = tk.StringVar(value="")
        self.q_value = tk.StringVar(value="")
        self.cg_value = tk.StringVar(value="")
        self.max_height_text = tk.StringVar(value="—")
        self.status_text = tk.StringVar(value="Enter the wind inputs and paste the STAAD tables.")
        self._configure_styles()
        self._build()
        if smoke_test:
            self.after(300, self.destroy)

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Excel.Treeview", background=WHITE, fieldbackground=WHITE, foreground=TEXT,
                        bordercolor=GRID, lightcolor=WHITE, darkcolor=WHITE, rowheight=22,
                        font=(self.font_name, 9))
        style.configure("Excel.Treeview.Heading", background=SECTION_GRAY, foreground=TEXT,
                        bordercolor=GRID, lightcolor=WHITE, darkcolor=WHITE, relief="flat",
                        font=(self.font_name, 8, "bold"), padding=(4, 3))
        style.map("Excel.Treeview", background=[("selected", ACCENT)], foreground=[("selected", WHITE)])

    def _button(self, parent, text, command):
        return tk.Button(parent, text=text, command=command, font=(self.font_name, 9),
                         bg=BUTTON_BG, fg=TEXT, activebackground=BUTTON_ACTIVE,
                         activeforeground=TEXT, relief="solid", bd=1,
                         highlightthickness=0, padx=9, pady=4)

    def _build(self):
        page = ScrollablePage(self)
        page.pack(fill="both", expand=True)
        wrap = tk.Frame(page.inner, bg=WHITE)
        wrap.pack(fill="both", expand=True, padx=16, pady=14)
        top = tk.Frame(wrap, bg=WHITE)
        top.pack(fill="x", pady=(0, 12))
        params = tk.Frame(top, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        params.pack(side="left")
        for col, (label, var) in enumerate((("Height adj. (m)", self.height_adjustment),
                                            ("q (kPa)", self.q_value), ("Cg", self.cg_value))):
            box = tk.Frame(params, bg=WHITE)
            box.grid(row=0, column=col, sticky="nsew")
            tk.Label(box, text=label, bg=SECTION_GRAY, fg=TEXT, font=(self.font_name, 8, "bold"),
                     relief="solid", bd=1, pady=3).pack(fill="x")
            tk.Entry(box, textvariable=var, bg=LIGHT_INPUT, fg=TEXT, justify="center",
                     font=(self.font_name, 9), relief="solid", bd=1).pack(fill="x")
            params.grid_columnconfigure(col, weight=1)
        self._button(top, "Calculate", self.calculate).pack(side="right")

        tables = tk.Frame(wrap, bg=WHITE)
        tables.pack(fill="x", pady=(0, 10))
        self.node_table = EditableTree(tables, "NODES", NODE_HEADERS,
                                       (48, 55, 55, 55), 13, lambda: None, self.font_name)
        self.member_table = EditableTree(tables, "MEMBERS", MEMBER_HEADERS,
                                         (50, 50, 50, 78, 90, 48, 60), 13, lambda: None, self.font_name)
        self.section_table = EditableTree(tables, "SECTION PROPERTIES", SECTION_HEADERS,
                                          (42, 66, 50, 42, 42, 42, 46, 55, 55, 55), 13, lambda: None, self.font_name)
        self.node_table.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.member_table.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        self.section_table.grid(row=0, column=2, sticky="nsew")
        tables.grid_columnconfigure(0, weight=1)
        tables.grid_columnconfigure(1, weight=2)
        tables.grid_columnconfigure(2, weight=2)

        tk.Label(wrap, textvariable=self.status_text, bg=WHITE, fg=MUTED_TEXT,
                 font=(self.font_name, 9), anchor="w").pack(fill="x", pady=(0, 8))
        metric = tk.Frame(wrap, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        metric.pack(anchor="w", pady=(0, 10))
        tk.Label(metric, text="Maximum height", bg=WHITE, fg=MUTED_TEXT,
                 font=(self.font_name, 8), padx=7, pady=3).pack(anchor="w")
        tk.Label(metric, textvariable=self.max_height_text, bg=WHITE, fg=TEXT,
                 font=(self.font_name, 9, "bold"), padx=7, pady=3).pack(anchor="w")

        outputs = tk.Frame(wrap, bg=WHITE)
        outputs.pack(fill="x")
        self.gz_output = OutputPanel(outputs, "Z Direction — STAAD GZ", self.font_name)
        self.gx_output = OutputPanel(outputs, "X Direction — STAAD GX", self.font_name)
        self.gz_output.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.gx_output.grid(row=0, column=1, sticky="nsew")
        outputs.grid_columnconfigure(0, weight=1)
        outputs.grid_columnconfigure(1, weight=1)

    def calculate(self):
        try:
            h = _as_number(self.height_adjustment.get(), field="Height adjustment")
            q = _as_number(self.q_value.get(), field="q")
            cg = _as_number(self.cg_value.get(), field="Cg")
            max_height, gz, gx = calculate_wind_outputs(
                self.node_table.get_rows(), self.member_table.get_rows(), self.section_table.get_rows(), h, q, cg)
            self.max_height_text.set(f"{max_height:.3f} m")
            self.gz_output.set_lines(gz)
            self.gx_output.set_lines(gx)
            self.status_text.set(f"Calculation complete · {len(gz)} Z-direction groups · {len(gx)} X-direction groups.")
        except Exception as exc:
            self.max_height_text.set("—")
            self.gz_output.set_error(str(exc))
            self.gx_output.set_error(str(exc))
            self.status_text.set(str(exc))


# output stuff
#

def self_test() -> int:
    print("No embedded project dataset is included in this public version.")
    return 0


def main():
    if "--self-test" in sys.argv:
        raise SystemExit(self_test())
    if "--smoke-test" in sys.argv:
        app = SiteStyleWindApp(smoke_test=True)
        app.mainloop()
        print("PASS: GUI smoke test")
        return
    app = SiteStyleWindApp()
    app.mainloop()


if __name__ == "__main__":
    main()
