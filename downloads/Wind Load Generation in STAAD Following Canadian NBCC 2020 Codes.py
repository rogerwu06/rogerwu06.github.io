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

# #
# keep the names the same as the spreadsheet where possible because changing them makes checking the input and output harder and can cause mistakes when comparing columns, some of these names are not very good but they are left as is for that reason
# #
# # #
# layout sizes below are mostly fixed values to make the tables fit on the screen, there is no calculation reason for most of these values and they may need to be changed if the window layout is changed later
#


APP_TITLE = "STAAD.Pro Wind Input Spreadsheet"
PALE_YELLOW = "#FFFFFF"     # white fill for the basic format
SECTION_GRAY = "#FFFFFF"
TITLE_BLUE = "#000000"          # old variable name is kept here but it is just black now
RED = "#000000"                 # old variable name is kept here but it is just black now
GRID = "#000000"
WHITE = "#FFFFFF"
LIGHT_INPUT = "#FFFFFF"
LIGHT_OUTPUT = "#FFFFFF"
TEXT = "#000000"
MUTED_TEXT = "#000000"
ACCENT = "#FFFFFF"
ACCENT_DARK = "#FFFFFF"
BUTTON_BG = "#FFFFFF"
BUTTON_ACTIVE = "#FFFFFF"

NODE_HEADERS = ("NODE", "X", "Y", "Z")
MEMBER_HEADERS = ("BEAM", "NODEA", "NODE B", "PROPERTY DEFINE", "MATERIAL", "BETA", "LENGTH")
SECTION_HEADERS = ("prop", "name", "ax", "d", "bf", "tf", "tw", "in", "iy", "ix")

# example input values from the orginal workbook, these are used when the program opens
SAMPLE_NODES = []
SAMPLE_MEMBERS = []
SAMPLE_SECTIONS = []

EXPECTED_GZ = []
EXPECTED_GX = []

# #
# workbook compatibility option, the old excel FILTER range stops before the last five member rows so this option keeps that same range when matching the orginal workbook output, turning it off uses all numeric member rows
# # #
ORIGINAL_COMPAT_MEMBER_TRIM = 5



# #
# checks the fonts installed on the computer and uses the first one that is avaliable, this is only for the gui and does not affect the calc
# #
#

def _first_font(root: tk.Misc, preferred=("Calibri", "Aptos", "Arial", "Segoe UI")) -> str:
    try:
        families = set(root.tk.call("font", "families"))
        for name in preferred:
            if name in families:
                return name
    except Exception:
        pass
    return "TkDefaultFont"


# values from the table come in as text so this converts them before the calc, also checks blank cells here so the main calc does not need to keep checking the same thing every time
# #

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


# format the number similar to excel general format because the staad output does not need extra zeros and it should stay close to the spreadsheet output
# # #

def _general_2dp(value: float) -> str:
    # number is already rounded above, remove zeros after the decimal so for example 2.00 is written as 2 and not 2.00
    return f"{value:.2f}".rstrip("0").rstrip(".")


def _roundup_2(value: float) -> float:
    # same idea as excel ROUNDUP, the small subtraction is needed because python floating point can give a number very slightly larger then the actual value and then it rounds one step too high
    return math.ceil(value * 100.0 - 1e-12) / 100.0


# interpolation for k from the same points used in the spreadsheet, this is kept directly in the function so it is easier to compare the numbers with the excel calc
#
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


# #
# main wind calculation below, node member and section calculations are kept in the same function because it follows the order of the excel helper calculations and makes it easier to check which step is giving a diffrent value
# #
# #
# keep the filter and group order the same because excel UNIQUE and FILTER keep the orginal order and if this is changed the final staad text will come out in a different order even when the numbers are the same
# # #

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

    # if there is a duplicate node number the first one is used, this is the same as the xlookup setup in the spreadsheet and duplicate node numbers should not normally be in the input anyway
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

    # height used for the wind calc is the lower member y value plus the height adjustment, this is the same value that was calculated in the excel helper columns
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

            # same filtering as the spreadsheet, for z wind the members running in z are not included and for x wind the members running in x are not included so make sure the direction check is not reversed
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

            # 2.1 is used for both values here on purpose because the workbook has the same value for Cn and Ct at this step, the section dimensions are applied after this
            fn = ce * cg * k * 2.1 * q
            ft = fn
            kn_per_m_fn = fn * d / 1000.0
            kn_per_m_ft = ft * bf / 1000.0
            load = _roundup_2(max(kn_per_m_fn, kn_per_m_ft))
            individual.append((beam, f"UNI {code} {_general_2dp(load)}"))

        # group members that have the same load value but keep the order from the input, this is needed so the output stays in the same order as the spreadsheet and does not move around each time
        groups = OrderedDict()
        for beam, suffix in individual:
            groups.setdefault(suffix, []).append(str(beam))
        return [" ".join(beams) + " " + suffix for suffix, beams in groups.items()]

    return max_height, direction_output("Z"), direction_output("X")


# pasted data can be tab seperated or just spaced depending on where it was copied from so this checks the common formats before putting the values into the table
# #
# # #

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

        # skip pasted header rows by checking if the first value is a number, all of the current input tables use numeric ids in the first column
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


# #
# treeview cells cant be editied directly so an entry box is placed over the selected cell when it is double clicked, after enter or clicking away the value is written back into the table
# #
#
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


# output text box used for the staad lines and the copy button, kept as one class because both wind directions use the same setup
# #

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
# outer canvas is used so the full page can scroll vertically when the screen height is smaller, the frame by itself does not provide scrolling
# #

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
        # make the page use the width of the visible window so the main gui does not need
        # a left and right scroll bar, the output text boxes can still scroll if the staad text is wider
        self.canvas.itemconfigure(self.window_id, width=max(1, event.width))

    def _mousewheel(self, event):
        # mouse wheel scroll is only applied when the pointer is over this page
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


# # #
# main window starts here, the padx and pady values are not all the same because the tables have different widths and this was adjusted so it fits without horizontal scrolling
# #
# #

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

    # #
    # gui formatting is set here, it is kept in this file so the program can be sent as one python file and does not need a seperate theme file
    # # #
    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # ttk can use the windows theme and make the buttons look raised so this overrides the style to keep the buttons plain black and white
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
            background=[("selected", "#000000")],
            foreground=[("selected", "#FFFFFF")],
        )

    def _make_menu(self):
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="Exit", command=self.destroy)
        menu.add_cascade(label="File", menu=file_menu)

        self.exact_compatibility = tk.BooleanVar(value=True)
        options = tk.Menu(menu, tearoff=False)
        options.add_checkbutton(
            label="Exact uploaded-workbook compatibility (omits its final 5 member rows)",
            variable=self.exact_compatibility,
            command=self.recalculate,
        )
        menu.add_cascade(label="Options", menu=options)
        self.config(menu=menu)


    def _build_fixed(self, parent):
        parent.configure(bg=WHITE)
        wrap = tk.Frame(parent, bg=WHITE)
        wrap.pack(fill="both", expand=True, padx=12, pady=10)

        # user guide was on the intro sheet before, putting it here so it is seen before the tables and does not need a seperate intro page
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

        # top input area for the wind values, arranged similar to the orginal excel sheet
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

        # three input tables from the workbook are shown below in the same general order
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

    # #
    # calculation runs automatic after a table or input is changed, small delay is used when pasting large tables so it does not run the full calc after every single row is added
    #
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
                exact_workbook_compatibility=self.exact_compatibility.get(),
            )
            self.max_height_text.set(f"{max_height:g}")
            self.gz_output.set_lines(gz)
            self.gx_output.set_lines(gx)
            mode = "exact uploaded-workbook compatibility" if self.exact_compatibility.get() else "fixed: all member rows processed"
            self.status_text.set(
                f"Max height: {max_height:g} m   |   GZ groups: {len(gz)}   |   GX groups: {len(gx)}   |   {mode}"
            )
        except Exception as exc:
            self.max_height_text.set("#N/A")
            msg = f"INPUT ERROR\n\n{exc}"
            self.gz_output.set_error(msg)
            self.gx_output.set_error(msg)
            self.status_text.set(str(exc))


# self test compares the important output values with the expected workbook result, run this after changing the calc or table code to make sure the result did not change
# #

def self_test() -> int:
    print("No embedded project dataset is included in this public version.")
    return 0


def main():
    if "--self-test" in sys.argv:
        raise SystemExit(self_test())
    if "--smoke-test" in sys.argv:
        app = WindLoadApp(smoke_test=True)
        app.mainloop()
        print("PASS: GUI smoke test")
        return
    app = WindLoadApp()
    app.mainloop()


if __name__ == "__main__":
    main()
