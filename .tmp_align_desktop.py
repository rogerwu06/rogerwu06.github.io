from pathlib import Path

seismic_path = Path('downloads/Seismic Load Generation in STAAD Following Canadian NBCC 2020 Codes.py')
wind_path = Path('downloads/Wind Load Generation in STAAD Following Canadian NBCC 2020 Codes.py')
seismic = seismic_path.read_text(encoding='utf-8')
wind = wind_path.read_text(encoding='utf-8')

seismic_class = r'''
class SiteStyleSeismicApp(tk.Tk):
    def __init__(self, *, smoke_test: bool = False):
        super().__init__()
        self.title("Seismic Load Generation in STAAD Following Canadian NBCC 2020 Codes")
        self.geometry("1380x860")
        self.minsize(1050, 680)
        self.configure(bg=WHITE)
        self.font_name = _first_font(self)
        self.with_pad = False
        self.inputs = {key: tk.StringVar(value="") for key in ("sa02", "ie", "rd", "ro", "height")}
        self.summary = {key: tk.StringVar(value="—") for key in (
            "coefficient", "total_vertical", "excluded_weight", "effective_weight", "vmax", "count"
        )}
        self.status = tk.StringVar(value="Enter the design inputs and paste the STAAD tables.")
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
        style.configure("Data.Treeview", background=WHITE, fieldbackground=WHITE, foreground=TEXT,
                        bordercolor=GRID, lightcolor=WHITE, darkcolor=WHITE, rowheight=22,
                        font=(self.font_name, 9))
        style.configure("Data.Treeview.Heading", background=ACCENT_LIGHT, foreground=TEXT,
                        bordercolor=GRID, lightcolor=WHITE, darkcolor=WHITE, relief="flat",
                        font=(self.font_name, 8, "bold"), padding=(4, 3))
        style.map("Data.Treeview", background=[("selected", ACCENT)], foreground=[("selected", WHITE)])

    def _button(self, parent, text, command, *, active=False):
        return tk.Button(parent, text=text, command=command, font=(self.font_name, 9),
                         bg="#E9E9E9" if active else BUTTON_BG, fg=TEXT,
                         activebackground=BUTTON_ACTIVE, activeforeground=TEXT,
                         relief="solid", bd=1, highlightthickness=0, padx=9, pady=4)

    def _build(self):
        page = ScrollablePage(self)
        page.pack(fill="both", expand=True)
        wrap = tk.Frame(page.inner, bg=WHITE)
        wrap.pack(fill="both", expand=True, padx=16, pady=14)

        toolbar = tk.Frame(wrap, bg=WHITE)
        toolbar.pack(fill="x", pady=(0, 10))
        self.no_pad_btn = self._button(toolbar, "NO PAD", lambda: self._set_mode(False), active=True)
        self.no_pad_btn.pack(side="left")
        self.pad_btn = self._button(toolbar, "WITH PAD", lambda: self._set_mode(True))
        self.pad_btn.pack(side="left", padx=(5, 0))
        self._button(toolbar, "Calculate", self.calculate).pack(side="right")

        params = tk.Frame(wrap, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
        params.pack(fill="x", pady=(0, 12))
        for col, (label, key) in enumerate((("Sa(0.2)", "sa02"), ("IE", "ie"), ("Rd", "rd"),
                                            ("Ro", "ro"), ("Height adj. (m)", "height"))):
            box = tk.Frame(params, bg=WHITE)
            box.grid(row=0, column=col, sticky="nsew")
            tk.Label(box, text=label, bg=ACCENT_LIGHT, fg=TEXT, font=(self.font_name, 8, "bold"),
                     relief="solid", bd=1, pady=3).pack(fill="x")
            tk.Entry(box, textvariable=self.inputs[key], bg=INPUT_BG, fg=TEXT, justify="center",
                     font=(self.font_name, 9), relief="solid", bd=1).pack(fill="x")
            params.grid_columnconfigure(col, weight=1)

        tables = tk.Frame(wrap, bg=WHITE)
        tables.pack(fill="x", pady=(0, 10))
        self.node_table = EditableTree(tables, "NODE GEOMETRY", NODE_HEADERS,
                                       (75, 90, 90, 90), 13, lambda: None, self.font_name)
        self.reaction_table = EditableTree(tables, "REACTIONS", REACTION_HEADERS,
                                           (70, 100, 72, 72, 72, 78, 78, 78), 13, lambda: None, self.font_name)
        self.node_table.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.reaction_table.grid(row=0, column=1, sticky="nsew")
        tables.grid_columnconfigure(0, weight=1)
        tables.grid_columnconfigure(1, weight=2)

        tk.Label(wrap, textvariable=self.status, bg=WHITE, fg=MUTED,
                 font=(self.font_name, 9), anchor="w").pack(fill="x", pady=(0, 8))

        metrics = tk.Frame(wrap, bg=WHITE)
        metrics.pack(fill="x", pady=(0, 10))
        defs = [("Coefficient", "coefficient"), ("Total vertical", "total_vertical"),
                ("Excluded weight", "excluded_weight"), ("Effective weight", "effective_weight"),
                ("Vmax", "vmax"), ("Matched nodes", "count")]
        for col, (label, key) in enumerate(defs):
            card = tk.Frame(metrics, bg=WHITE, highlightbackground=GRID, highlightthickness=1)
            card.grid(row=0, column=col, sticky="nsew", padx=(0, 6 if col < len(defs)-1 else 0))
            tk.Label(card, text=label, bg=WHITE, fg=MUTED, font=(self.font_name, 8),
                     anchor="w", padx=7, pady=3).pack(fill="x")
            tk.Label(card, textvariable=self.summary[key], bg=WHITE, fg=TEXT,
                     font=(self.font_name, 9, "bold"), anchor="w", padx=7, pady=3).pack(fill="x")
            metrics.grid_columnconfigure(col, weight=1)

        outputs = tk.Frame(wrap, bg=WHITE)
        outputs.pack(fill="x")
        self.fz_output = OutputPanel(outputs, "Z Direction — STAAD JOINT LOAD (FZ)", self.font_name)
        self.fx_output = OutputPanel(outputs, "X Direction — STAAD JOINT LOAD (FX)", self.font_name)
        self.fz_output.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.fx_output.grid(row=0, column=1, sticky="nsew")
        outputs.grid_columnconfigure(0, weight=1)
        outputs.grid_columnconfigure(1, weight=1)

    def _set_mode(self, with_pad: bool):
        self.with_pad = with_pad
        self.no_pad_btn.configure(bg=BUTTON_BG if with_pad else "#E9E9E9")
        self.pad_btn.configure(bg="#E9E9E9" if with_pad else BUTTON_BG)
        self.status.set("Mode: WITH PAD" if with_pad else "Mode: NO PAD")

    def calculate(self):
        try:
            result = calculate_seismic(
                self.node_table.get_rows(), self.reaction_table.get_rows(),
                sa02=_as_number(self.inputs["sa02"].get(), field="Sa(0.2)"),
                ie=_as_number(self.inputs["ie"].get(), field="IE"),
                rd=_as_number(self.inputs["rd"].get(), field="Rd"),
                ro=_as_number(self.inputs["ro"].get(), field="Ro"),
                height_adjustment=_as_number(self.inputs["height"].get(), field="Height adjustment"),
                with_pad=self.with_pad,
            )
            self.summary["coefficient"].set(f'{result["coefficient"]:.9f}')
            self.summary["total_vertical"].set(f'{result["total_vertical"]:.3f} kN')
            self.summary["excluded_weight"].set(f'{result["excluded_weight"]:.3f} kN')
            self.summary["effective_weight"].set(f'{result["effective_weight"]:.3f} kN')
            self.summary["vmax"].set(f'{result["vmax"]:.3f} kN')
            self.summary["count"].set(str(result["count"]))
            self.fz_output.set_lines(result["fz_lines"])
            self.fx_output.set_lines(result["fx_lines"])
            self.status.set(f'Calculation complete · {result["count"]} matched reaction nodes.')
        except Exception as exc:
            for value in self.summary.values(): value.set("—")
            self.fz_output.set_error(str(exc))
            self.fx_output.set_error(str(exc))
            self.status.set(str(exc))
'''

wind_class = r'''
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
'''

if 'class SiteStyleSeismicApp' not in seismic:
    seismic = seismic.replace('\n# basic self test for the seismic calculation', '\n' + seismic_class + '\n\n# basic self test for the seismic calculation')
seismic = seismic.replace('app = SeismicApp(smoke_test="--smoke-test" in sys.argv)',
                          'app = SiteStyleSeismicApp(smoke_test="--smoke-test" in sys.argv)')
seismic_path.write_text(seismic, encoding='utf-8')

if 'class SiteStyleWindApp' not in wind:
    wind = wind.replace('\n# self test compares the important output values', '\n' + wind_class + '\n\n# self test compares the important output values')
wind = wind.replace('app = WindLoadApp(smoke_test=True)', 'app = SiteStyleWindApp(smoke_test=True)')
wind = wind.replace('app = WindLoadApp()\n    app.mainloop()', 'app = SiteStyleWindApp()\n    app.mainloop()')
wind_path.write_text(wind, encoding='utf-8')
