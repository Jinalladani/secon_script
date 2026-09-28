"""
DST Synthetic Data Generator - Multi Depth GUI v2
Uses the existing dst_generator.py and soil_models.py modules.

New in this version:
- Enter multiple sample depths in a table.
- Each depth can have its own soil type, density, sample weight, c and phi.
- Common PR constant and three normal stresses.
- Generate all depths in one operation.
- Results table shows target c/phi and calculated c/phi/R² for every depth.
- Select a depth to view its three DST curves.
- Export all generated depths to one Excel workbook.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import math
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib.lines import Line2D
from openpyxl import load_workbook

from soil_models import get_soil_names, get_soil_profile, validate_target_strength
from dst_generator import classify_density, generate_three_trials, calculate_summary
from excel_export_multi import export_multiple_dst_to_excel


class DSTGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Synthetic Direct Shear Test Data Generator - Multi Depth v2")
        self.root.geometry("1220x900")
        self.root.minsize(1100, 820)

        self.pr_constant = tk.StringVar(value="0.325")
        self.normal_1 = tk.StringVar(value="0.50")
        self.normal_2 = tk.StringVar(value="1.00")
        self.normal_3 = tk.StringVar(value="1.50")
        self.random_seed = tk.StringVar(value="AUTO")
        self.realistic_variation = tk.BooleanVar(value=True)
        self.instrument_variation = tk.BooleanVar(value=True)
        self.natural_variation = tk.BooleanVar(value=True)

        # Direct shear specimen volume: 6 cm × 6 cm × 2.5 cm = 90 cm³
        self.sample_volume_cm3 = 90.0

        self.rows = []
        self.results = []
        self.selected_result = None
        self.last_excel_output_dir = None

        self.create_gui()
        self.add_depth_row()

    # ----------------------------------------------------------
    # GUI
    # ----------------------------------------------------------
    def create_gui(self):
        # ------------------------------------------------------
        # Scrollable main GUI container
        # ------------------------------------------------------
        # The existing GUI layout is kept unchanged.  The complete
        # form is placed inside a vertically scrollable canvas so that
        # adding many depth rows does not make the lower controls/results
        # inaccessible on smaller screens.
        scroll_container = ttk.Frame(self.root)
        scroll_container.pack(fill="both", expand=True)

        scroll_canvas = tk.Canvas(scroll_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            scroll_container,
            orient="vertical",
            command=scroll_canvas.yview
        )
        scroll_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        scroll_canvas.pack(side="left", fill="both", expand=True)

        outer = ttk.Frame(scroll_canvas, padding=12)
        window_id = scroll_canvas.create_window(
            (0, 0), window=outer, anchor="nw"
        )

        def update_scrollregion(event=None):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

        def fit_inner_width(event):
            scroll_canvas.itemconfigure(window_id, width=event.width)
            update_scrollregion()

        outer.bind("<Configure>", update_scrollregion)
        scroll_canvas.bind("<Configure>", fit_inner_width)

        # Windows / macOS mouse-wheel support.
        def on_mousewheel(event):
            if event.delta:
                scroll_canvas.yview_scroll(
                    int(-event.delta / 120), "units"
                )

        scroll_canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Linux mouse-wheel support.
        scroll_canvas.bind_all(
            "<Button-4>", lambda e: scroll_canvas.yview_scroll(-1, "units")
        )
        scroll_canvas.bind_all(
            "<Button-5>", lambda e: scroll_canvas.yview_scroll(1, "units")
        )

        title = ttk.Label(
            outer,
            text="SYNTHETIC DIRECT SHEAR TEST DATA GENERATOR",
            font=("Segoe UI", 16, "bold"),
        )
        title.pack(pady=(0, 10))

        # Test-wide inputs
        common = ttk.LabelFrame(outer, text="TEST / COMMON INPUTS", padding=8)
        common.pack(fill="x", pady=5)

        self._entry(common, "BH ID", 0, 0, "")
        self.bh_id_var = self._last_var

        self._entry(common, "PR Constant", 0, 2, self.pr_constant)
        self._entry(common, "Normal Stress 1", 1, 0, self.normal_1)
        self._entry(common, "Normal Stress 2", 1, 2, self.normal_2)
        self._entry(common, "Normal Stress 3", 1, 4, self.normal_3)
        self._entry(common, "Random Seed", 2, 0, self.random_seed)

        # Generation options
        opt = ttk.Frame(common)
        opt.grid(row=2, column=2, columnspan=4, sticky="w", padx=5)
        ttk.Checkbutton(opt, text="Realistic Soil Behaviour",
                        variable=self.realistic_variation).pack(side="left", padx=5)
        ttk.Checkbutton(opt, text="Instrument Variation",
                        variable=self.instrument_variation).pack(side="left", padx=5)
        ttk.Checkbutton(opt, text="Natural Test Variation",
                        variable=self.natural_variation).pack(side="left", padx=5)

        # Multiple-depth table
        table_box = ttk.LabelFrame(
            outer,
            text="MULTIPLE DEPTH INPUT  (one row = one DST sample)",
            padding=8,
        )
        table_box.pack(fill="x", pady=5)

        headers = [
            ("No.", 5),
            ("Sample Depth (m)", 17),
            ("Soil Type", 23),
            ("Sample Type", 20),
            ("Density (g/cc)", 16),
            ("Sample Weight (g)", 19),
            ("Target c (kg/cm²)", 18),
            ("Target φ (°)", 14),
            ("Behaviour Class", 17),
            ("Remove", 9),
        ]
        for col, (text, width) in enumerate(headers):
            ttk.Label(
                table_box, text=text, width=width, anchor="center",
                font=("Segoe UI", 9, "bold")
            ).grid(row=0, column=col, padx=2, pady=3)

        self.table_frame = table_box

        ttk.Label(
            table_box,
            text="AUTOMATIC: Sample Weight (g) = Density (g/cc) × 90 cm³  |  Mould volume = 6 × 6 × 2.5 cm = 90 cm³",
            font=("Segoe UI", 8, "italic")
        ).grid(row=999, column=0, columnspan=9, sticky="w", padx=4, pady=(5, 2))

        button_bar = ttk.Frame(outer)
        button_bar.pack(fill="x", pady=6)
        ttk.Button(button_bar, text="+ ADD DEPTH", command=self.add_depth_row).pack(side="left", padx=4)
        ttk.Button(button_bar, text="VALIDATE INPUT", command=self.validate_inputs).pack(side="left", padx=4)
        ttk.Button(button_bar, text="GENERATE ALL DEPTHS", command=self.generate_all).pack(side="left", padx=4)
        ttk.Button(button_bar, text="EXPORT ALL TO EXCEL", command=self.export_all).pack(side="left", padx=4)
        ttk.Button(button_bar, text="EXPORT FAILURE ENVELOPE PLOTS", command=self.export_failure_envelope_plots).pack(side="left", padx=4)
        ttk.Button(button_bar, text="RESET", command=self.reset_form).pack(side="right", padx=4)
        ttk.Button(button_bar, text="IMPORT FROM EXCEL", command=self.import_from_excel).pack(side="right", padx=4)

        # Results
        result_box = ttk.LabelFrame(outer, text="CALCULATED STRENGTH ENVELOPE - ALL DEPTHS", padding=6)
        result_box.pack(fill="x", pady=5)

        cols = ("depth", "soil", "sample_type", "target_c", "target_phi", "calc_c", "calc_phi", "r2", "trials")
        self.result_tree = ttk.Treeview(result_box, columns=cols, show="headings", height=6)
        headings = {
            "depth": "Depth (m)", "soil": "Soil Type", "sample_type": "Sample Type", "target_c": "Target c",
            "target_phi": "Target φ", "calc_c": "Calculated c",
            "calc_phi": "Calculated φ", "r2": "R²", "trials": "Trials"
        }
        widths = {"depth": 90, "soil": 190, "sample_type": 150, "target_c": 95, "target_phi": 95,
                  "calc_c": 105, "calc_phi": 105, "r2": 80, "trials": 70}
        for c in cols:
            self.result_tree.heading(c, text=headings[c])
            self.result_tree.column(c, width=widths[c], anchor="center")
        self.result_tree.pack(fill="x", expand=True)
        self.result_tree.bind("<<TreeviewSelect>>", self.on_result_select)

        # Status
        self.status_var = tk.StringVar(value="Enter one or more depths, then generate.")
        ttk.Label(outer, textvariable=self.status_var, font=("Segoe UI", 9, "italic")).pack(
            anchor="w", pady=(3, 5)
        )

        # Graph
        graph_box = ttk.LabelFrame(outer, text="SELECTED DEPTH - SYNTHETIC DIRECT SHEAR CURVES", padding=6)
        # Keep the report area fully inside the outer scrollbar.  The
        # failure-envelope report is taller than the old 350 px canvas,
        # so a fixed report-height canvas is required; the outer GUI
        # scrollbar will then scroll through the complete report.
        graph_box.pack(fill="x", pady=5)
        self.canvas = tk.Canvas(graph_box, background="white", height=820)
        self.canvas.pack(fill="x", expand=False)
        self.canvas.bind("<Configure>", lambda e: self.redraw())

    def _entry(self, parent, label, row, col, value):
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=col, padx=7, pady=3, sticky="w")
        ttk.Label(frame, text=label, width=18).pack(side="left")
        if isinstance(value, tk.Variable):
            var = value
        else:
            var = tk.StringVar(value=value)
        ttk.Entry(frame, textvariable=var, width=16).pack(side="left")
        self._last_var = var
        return var

    # ----------------------------------------------------------
    # Depth rows
    # ----------------------------------------------------------
    def add_depth_row(self):
        idx = len(self.rows) + 1
        depth = tk.StringVar(value="")
        soil = tk.StringVar(value="Medium Dense Sand")
        sample_type = tk.StringVar(value="Undisturbed Sample")
        density = tk.StringVar(value="")
        weight = tk.StringVar(value="")
        c = tk.StringVar(value="")
        phi = tk.StringVar(value="")
        behaviour = tk.StringVar(value="Enter density")

        row = {
            "depth": depth, "soil": soil, "sample_type": sample_type,
            "density": density, "weight": weight, "c": c, "phi": phi,
            "behaviour": behaviour, "widgets": []
        }
        self.rows.append(row)
        self._draw_depth_row(idx, row)

    def _draw_depth_row(self, idx, row):
        r = idx

        ttk.Label(self.table_frame, text=str(idx), width=5, anchor="center").grid(
            row=r, column=0, padx=2, pady=2
        )

        # Depth and editable engineering inputs
        for col, var, width in [
            (1, row["depth"], 17),
            (4, row["density"], 16),
            (6, row["c"], 18),
            (7, row["phi"], 14),
        ]:
            e = ttk.Entry(self.table_frame, textvariable=var, width=width)
            e.grid(row=r, column=col, padx=2, pady=2)
            row["widgets"].append(e)

        soil_combo = ttk.Combobox(
            self.table_frame, textvariable=row["soil"],
            values=get_soil_names(), state="readonly", width=21
        )
        soil_combo.grid(row=r, column=2, padx=2, pady=2)
        soil_combo.bind("<<ComboboxSelected>>", lambda e, rr=row: self.update_row_behaviour(rr))
        row["widgets"].append(soil_combo)

        sample_combo = ttk.Combobox(
            self.table_frame, textvariable=row["sample_type"],
            values=("Undisturbed Sample", "Remolded"),
            state="readonly", width=18
        )
        sample_combo.grid(row=r, column=3, padx=2, pady=2)
        row["widgets"].append(sample_combo)

        # Sample weight is calculated automatically from density × 90 cm³.
        weight_entry = ttk.Entry(
            self.table_frame, textvariable=row["weight"], width=19, state="readonly"
        )
        weight_entry.grid(row=r, column=5, padx=2, pady=2)
        # Prevent the earlier generic loop from creating a duplicate weight widget.
        row["widgets"].append(weight_entry)

        ttk.Label(
            self.table_frame, textvariable=row["behaviour"], width=17, anchor="center"
        ).grid(row=r, column=8, padx=2, pady=2)

        ttk.Button(
            self.table_frame, text="X", width=5,
            command=lambda rr=row: self.remove_depth_row(rr)
        ).grid(row=r, column=9, padx=2, pady=2)

        row["density"].trace_add("write", lambda *args, rr=row: self.update_row_from_density(rr))

    def update_row_from_density(self, row):
        """Update behaviour and automatically calculate sample weight.

        Mould/specimen volume = 6 × 6 × 2.5 = 90 cm³.
        With density in g/cc, weight is directly density × 90 g.
        """
        try:
            d = float(row["density"].get())
            if d > 0:
                row["weight"].set(f"{d * self.sample_volume_cm3:.2f}")
                row["behaviour"].set(classify_density(row["soil"].get(), d))
            else:
                row["weight"].set("")
                row["behaviour"].set("Enter density")
        except (ValueError, TypeError):
            row["weight"].set("")
            row["behaviour"].set("Enter density")

    def update_row_behaviour(self, row):
        try:
            d = float(row["density"].get())
            row["behaviour"].set(classify_density(row["soil"].get(), d))
        except (ValueError, TypeError):
            row["behaviour"].set("Enter density")

    def remove_depth_row(self, row):
        if len(self.rows) == 1:
            messagebox.showwarning("Depth Input", "At least one depth row is required.")
            return
        self.rows.remove(row)
        for r in self.rows:
            for w in r["widgets"]:
                w.destroy()
        # Rebuild row widgets while keeping variables
        for r in self.rows:
            r["widgets"] = []
        # destroy all table data rows except header
        for child in list(self.table_frame.winfo_children()):
            info = child.grid_info()
            if info and int(info.get("row", 0)) > 0:
                child.destroy()
        for i, r in enumerate(self.rows, 1):
            self._draw_depth_row(i, r)

    # ----------------------------------------------------------
    # Validation / generation
    # ----------------------------------------------------------
    def _common_values(self):
        try:
            stresses = [float(self.normal_1.get()), float(self.normal_2.get()), float(self.normal_3.get())]
            pr = float(self.pr_constant.get())
        except ValueError:
            raise ValueError("PR Constant and all three normal stresses must be numeric.")

        if pr <= 0:
            raise ValueError("PR Constant must be greater than zero.")
        if not (stresses[0] > 0 and stresses[1] > 0 and stresses[2] > 0):
            raise ValueError("Normal stresses must be greater than zero.")
        if not (stresses[0] < stresses[1] < stresses[2]):
            raise ValueError("Normal stresses must satisfy Trial 1 < Trial 2 < Trial 3.")
        return pr, stresses

    def _read_rows(self, show_warnings=True):
        if not self.bh_id_var.get().strip():
            raise ValueError("BH ID is required.")

        self._common_values()
        parsed = []

        for i, row in enumerate(self.rows, 1):
            try:
                depth = float(row["depth"].get())
                density = float(row["density"].get())
                weight = float(row["weight"].get())
                c = float(row["c"].get())
                phi = float(row["phi"].get())
            except ValueError:
                raise ValueError(f"Row {i}: Depth, density, sample weight, c and phi must be numeric.")

            if row["sample_type"].get() not in ("Undisturbed Sample", "Remolded"):
                raise ValueError(f"Row {i}: Select a valid sample type.")
            if depth < 0:
                raise ValueError(f"Row {i}: Sample depth cannot be negative.")
            if density <= 0:
                raise ValueError(f"Row {i}: Density must be greater than zero.")
            expected_weight = density * self.sample_volume_cm3
            if weight <= 0:
                raise ValueError(f"Row {i}: Sample weight must be greater than zero.")
            if abs(weight - expected_weight) > 0.01:
                row["weight"].set(f"{expected_weight:.2f}")
                weight = expected_weight
            if c < 0:
                raise ValueError(f"Row {i}: Cohesion cannot be negative.")
            if phi < 0 or phi >= 90:
                raise ValueError(f"Row {i}: Friction angle must be between 0 and 90 degrees.")

            inside, msg = validate_target_strength(row["soil"].get(), c, phi)
            if not inside and show_warnings:
                if not messagebox.askyesno(
                    "Target Strength Warning",
                    f"Depth {depth:g} m:\n{msg}\n\nContinue with this target?"
                ):
                    raise ValueError("Generation cancelled by user.")

            parsed.append({
                "depth": depth,
                "soil": row["soil"].get(),
                "sample_type": row["sample_type"].get(),
                "density": density,
                "weight": weight,
                "c": c,
                "phi": phi,
            })
        return parsed

    def validate_inputs(self):
        try:
            parsed = self._read_rows(show_warnings=True)
            self.status_var.set(f"Valid input: {len(parsed)} depth(s).")
            messagebox.showinfo("Validation", f"All inputs are valid for {len(parsed)} depth(s).")
            return True
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
            return False

    def generate_all(self):
        try:
            parsed = self._read_rows(show_warnings=True)
            _, stresses = self._common_values()

            seed_text = self.random_seed.get().strip()
            base_seed = None if not seed_text or seed_text.upper() == "AUTO" else int(seed_text)

            self.results = []

            for i, item in enumerate(parsed):
                seed = None if base_seed is None else base_seed + i
                trials = generate_three_trials(
                    soil_type=item["soil"],
                    density=item["density"],
                    c=item["c"],
                    phi=item["phi"],
                    normal_stresses=stresses,
                    seed=seed,
                    n_points=61,
                )
                summary = calculate_summary(trials)
                self.results.append({
                    **item,
                    "trials": trials,
                    "summary": summary,
                })

            self.refresh_results()
            if self.results:
                self.selected_result = self.results[0]
                self.result_tree.selection_set(self.result_tree.get_children()[0])
                self.redraw()

            self.status_var.set(
                f"Generated {len(self.results)} depth(s) × 3 trials × 61 observations."
            )
        except (ValueError, TypeError) as e:
            messagebox.showerror("Generation Error", str(e))
        except Exception as e:
            messagebox.showerror("Generation Error", str(e))

    # ----------------------------------------------------------
    # Results
    # ----------------------------------------------------------
    def refresh_results(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        for result in self.results:
            s = result["summary"]
            self.result_tree.insert(
                "",
                "end",
                values=(
                    f"{result['depth']:.2f}",
                    result["soil"],
                    result["sample_type"],
                    f"{result['c']:.3f}",
                    f"{result['phi']:.2f}",
                    f"{s['calculated_c']:.3f}",
                    f"{s['calculated_phi']:.2f}",
                    f"{s['r_squared']:.3f}",
                    "3",
                ),
            )

    def on_result_select(self, event=None):
        selected = self.result_tree.selection()
        if not selected:
            return
        index = self.result_tree.index(selected[0])
        if 0 <= index < len(self.results):
            self.selected_result = self.results[index]
            self.redraw()

    # ----------------------------------------------------------
    # Report-style Failure Envelope Plot
    # ----------------------------------------------------------
    def redraw(self):
        """Draw a report-style DST failure-envelope sheet for the selected depth."""
        self.canvas.delete("all")
        result = self.selected_result
        width = max(self.canvas.winfo_width(), 1100)
        height = max(self.canvas.winfo_height(), 720)

        if not result or not result.get("trials"):
            self.canvas.create_text(
                width / 2, height / 2,
                text="Generate data to display the DST Failure Envelope.",
                font=("Segoe UI", 14, "bold")
            )
            return

        trials = result["trials"]
        summary = result["summary"]
        c_calc = float(summary["calculated_c"])
        phi_calc = float(summary["calculated_phi"])
        r2 = float(summary["r_squared"])
        slope = math.tan(math.radians(phi_calc))

        # ------------------------------------------------------
        # Report geometry
        # ------------------------------------------------------
        margin = 8
        x0, x1 = margin, width - margin
        y0 = 6
        title_h = 58
        info_h = 70
        graph_title_h = 35
        graph_h = 410
        result_title_h = 35
        table_header_h = 50
        table_row_h = 38

        # If the window is short, keep the report visible without clipping.
        required = title_h + info_h + graph_title_h + graph_h + result_title_h + table_header_h + 3 * table_row_h + 20
        if height < required:
            graph_h = max(300, height - (title_h + info_h + graph_title_h + result_title_h + table_header_h + 3 * table_row_h + 25))

        y = y0

        # ------------------------------------------------------
        # Outer border and title
        # ------------------------------------------------------
        self.canvas.create_rectangle(x0, y, x1, y + title_h, outline="#222222", width=2)
        self.canvas.create_text(
            (x0 + x1) / 2, y + 22,
            text="DIRECT SHEAR TEST – FAILURE ENVELOPE",
            font=("Segoe UI", 22, "bold"), fill="#172554"
        )
        self.canvas.create_text(
            (x0 + x1) / 2, y + 47,
            text="IS 2720 (Part 13) – 1986",
            font=("Segoe UI", 11, "bold"), fill="#172554"
        )
        y += title_h

        # ------------------------------------------------------
        # Top information table
        # ------------------------------------------------------
        self.canvas.create_rectangle(x0, y, x1, y + info_h, outline="#222222", width=1)
        c1 = x0 + (x1 - x0) * 0.28
        c2 = x0 + (x1 - x0) * 0.52
        c3 = x0 + (x1 - x0) * 0.78
        for xx in (c1, c2, c3):
            self.canvas.create_line(xx, y, xx, y + info_h, fill="#222222")

        sample_display = "UDS" if result["sample_type"] == "Undisturbed Sample" else "Remoulded"
        depth_text = f"{result['depth']:.2f} m"
        density_text = f"{result['density']:.2f} gm/cc"
        weight_text = f"{result['weight']:.2f} g"

        self._report_pair(x0 + 18, y + 22, "Borehole ID", self.bh_id_var.get().strip())
        self._report_pair(x0 + 18, y + 49, "Depth", depth_text)

        self._report_pair(c1 + 18, y + 22, "Sample Type", sample_display)
        self._report_pair(c1 + 18, y + 49, "Density", density_text)

        self._report_pair(c2 + 18, y + 22, "Shear Box Size", "60 mm × 60 mm")
        self._report_pair(c2 + 18, y + 49, "Sample Weight", weight_text)

        self._report_pair(c3 + 18, y + 22, "Condition", sample_display)
        self._report_pair(c3 + 18, y + 49, "Normal Stress", "/".join(f"{t.normal_stress:.2f}" for t in trials) + " kg/cm²")
        y += info_h

        # ------------------------------------------------------
        # Graph heading
        # ------------------------------------------------------
        self.canvas.create_rectangle(x0, y, x1, y + graph_title_h, outline="#222222")
        self.canvas.create_text(
            (x0 + x1) / 2, y + graph_title_h / 2,
            text="NORMAL STRESS vs PEAK SHEAR STRESS (FAILURE ENVELOPE)",
            font=("Segoe UI", 12, "bold")
        )
        y += graph_title_h

        # ------------------------------------------------------
        # Graph area
        # ------------------------------------------------------
        graph_top = y
        graph_bottom = y + graph_h
        self.canvas.create_rectangle(x0, graph_top, x1, graph_bottom, outline="#222222")

        left = x0 + 90
        right = x1 - 40
        top = graph_top + 40
        bottom = graph_bottom - 60
        pw = right - left
        ph = bottom - top

        normal = [float(t.normal_stress) for t in trials]
        peak = [float(t.peak_shear_stress) for t in trials]
        max_sigma = max(normal)
        max_tau = max(peak)
        xlim = max(1.0, math.ceil((max_sigma * 1.18) / 0.1) * 0.1)
        # Keep a report-like 0–1.2 range when the data fits.
        if max_sigma <= 1.0:
            xlim = 1.2
        ylim = max(0.4, math.ceil((max_tau * 1.20) / 0.1) * 0.1)
        if max_tau <= 1.0:
            ylim = 1.2

        def xp(v):
            return left + (v / xlim) * pw

        def yp(v):
            return bottom - (v / ylim) * ph

        # Grid and tick labels.
        x_step = 0.2 if xlim <= 1.2 else 0.5
        y_step = 0.2 if ylim <= 1.2 else 0.5
        xv = 0.0
        while xv <= xlim + 1e-9:
            xx = xp(xv)
            self.canvas.create_line(xx, top, xx, bottom, fill="#d7d7d7", dash=(3, 3))
            self.canvas.create_text(xx, bottom + 18, text=f"{xv:.2f}", font=("Segoe UI", 9))
            xv += x_step
        yv = 0.0
        while yv <= ylim + 1e-9:
            yy = yp(yv)
            self.canvas.create_line(left, yy, right, yy, fill="#d7d7d7", dash=(3, 3))
            self.canvas.create_text(left - 12, yy, text=f"{yv:.2f}", anchor="e", font=("Segoe UI", 9))
            yv += y_step

        self.canvas.create_line(left, bottom, right, bottom, width=2)
        self.canvas.create_line(left, top, left, bottom, width=2)

        self.canvas.create_text(
            (left + right) / 2, graph_bottom - 20,
            text="Normal Stress, σₙ (kg/cm²)", font=("Segoe UI", 11, "bold")
        )
        self.canvas.create_text(
            x0 + 25, (top + bottom) / 2,
            text="Peak Shear Stress, τ (kg/cm²)", angle=90,
            font=("Segoe UI", 11, "bold")
        )

        # ------------------------------------------------------
        # Failure envelope – dashed extrapolation + solid fitted part
        # ------------------------------------------------------
        tau_at_zero = c_calc
        x_intercept = -c_calc / slope if slope > 0 else 0.0
        env_x_start = max(0.0, x_intercept)
        env_x_end = xlim
        # Dashed extrapolation from sigma=0 to first test point.
        first_sigma = min(normal)
        if c_calc >= 0 and first_sigma > 0:
            self.canvas.create_line(
                xp(0), yp(c_calc), xp(first_sigma), yp(c_calc + slope * first_sigma),
                fill="#e00000", width=2, dash=(7, 5)
            )
        # Solid fitted envelope across the tested range.
        self.canvas.create_line(
            xp(first_sigma), yp(c_calc + slope * first_sigma),
            xp(env_x_end), yp(c_calc + slope * env_x_end),
            fill="#e00000", width=2
        )

        # Peak points + coordinate labels.
        for t in trials:
            xx, yy = xp(t.normal_stress), yp(t.peak_shear_stress)
            self.canvas.create_oval(xx - 7, yy - 7, xx + 7, yy + 7, fill="white", outline="#111111", width=2)
            label = f"({t.normal_stress:.2f}, {t.peak_shear_stress:.2f})"
            self.canvas.create_text(xx, yy - 18, text=label, fill="#172bb5", font=("Segoe UI", 9, "bold"))

        # Legend.
        lx, ly = left + 15, top + 20
        self.canvas.create_rectangle(lx, ly, lx + 205, ly + 70, outline="#222222", fill="white")
        self.canvas.create_oval(lx + 28 - 7, ly + 20 - 7, lx + 28 + 7, ly + 20 + 7, fill="white", outline="#111111", width=2)
        self.canvas.create_text(lx + 55, ly + 20, text="Test Results (Peak)", anchor="w", font=("Segoe UI", 9))
        self.canvas.create_line(lx + 15, ly + 49, lx + 45, ly + 49, fill="#e00000", width=2)
        self.canvas.create_text(lx + 55, ly + 49, text="Failure Envelope", anchor="w", font=("Segoe UI", 9))

        # Clean equation / strength box.
        bx1, by1 = right - 280, bottom - 165
        bx2, by2 = right - 8, bottom - 8
        self.canvas.create_rectangle(bx1, by1, bx2, by2, outline="#222222", fill="white")
        self.canvas.create_text(bx1 + 16, by1 + 22, text="τ = c + σₙ tan φ", anchor="w", font=("Segoe UI", 10, "bold"))
        self.canvas.create_text(bx1 + 16, by1 + 48, text=f"= {c_calc:.2f} + σₙ tan {phi_calc:.1f}°", anchor="w", font=("Segoe UI", 9, "bold"))
        self.canvas.create_text(bx1 + 16, by1 + 83, text=f"c = {c_calc:.2f} kg/cm²", anchor="w", font=("Segoe UI", 10, "bold"), fill="#e00000")
        self.canvas.create_text(bx1 + 16, by1 + 111, text=f"φ = {phi_calc:.1f}°", anchor="w", font=("Segoe UI", 10, "bold"), fill="#e00000")
        self.canvas.create_text(bx1 + 16, by1 + 139, text=f"R² = {r2:.3f}", anchor="w", font=("Segoe UI", 10, "bold"))

        y = graph_bottom

        # ------------------------------------------------------
        # Bottom results table
        # ------------------------------------------------------
        self.canvas.create_rectangle(x0, y, x1, y + result_title_h, outline="#222222")
        self.canvas.create_text(
            (x0 + x1) / 2, y + result_title_h / 2,
            text="DIRECT SHEAR TEST RESULTS",
            font=("Segoe UI", 12, "bold"), fill="#172bb5"
        )
        y += result_title_h

        # Table column widths proportional to the attached report.
        rel = [0.07, 0.13, 0.12, 0.11, 0.15, 0.16, 0.11, 0.15]
        colx = [x0]
        for frac in rel:
            colx.append(colx[-1] + (x1 - x0) * frac)

        headers = [
            "Sr. No.", "Sample Depth\n(m)", "Sample Type", "Density\n(gm/cc)",
            "Normal Stress,\nσₙ (kg/cm²)", "Shear Stress at Failure\n(Peak), τ (kg/cm²)",
            "Cohesion,\nc (kg/cm²)", "Angle of Internal Friction,\nφ (°)"
        ]
        self.canvas.create_rectangle(x0, y, x1, y + table_header_h, outline="#222222", fill="#f1f4fb")
        for i in range(1, len(colx) - 1):
            self.canvas.create_line(colx[i], y, colx[i], y + table_header_h + 3 * table_row_h, fill="#222222")
        for i, h in enumerate(headers):
            self.canvas.create_text((colx[i] + colx[i+1]) / 2, y + table_header_h / 2, text=h,
                                    justify="center", font=("Segoe UI", 8, "bold"))
        y += table_header_h

        # Draw 3 trial rows. Sample Depth, Sample Type, Density, Cohesion and φ
        # are visually merged vertically across all three trials.
        table_y0 = y
        for i, t in enumerate(trials, 1):
            self.canvas.create_rectangle(x0, y, x1, y + table_row_h, outline="#222222", fill="white")
            vals = [
                str(i),
                depth_text if i == 2 else "",
                sample_display if i == 2 else "",
                f"{result['density']:.2f}" if i == 2 else "",
                f"{t.normal_stress:.2f}",
                f"{t.peak_shear_stress:.2f}",
                f"{c_calc:.2f}" if i == 2 else "",
                f"{phi_calc:.1f}" if i == 2 else ""
            ]
            for j, val in enumerate(vals):
                self.canvas.create_text((colx[j] + colx[j+1]) / 2, y + table_row_h / 2,
                                        text=val, font=("Segoe UI", 9), justify="center")
            y += table_row_h

        # Remove internal horizontal lines for vertically merged columns.
        merged_cols = [1, 2, 3, 6, 7]
        for j in merged_cols:
            self.canvas.create_rectangle(colx[j] + 1, table_y0 + 1, colx[j+1] - 1,
                                         table_y0 + 3 * table_row_h - 1,
                                         outline="", fill="white")
            value = {1: depth_text, 2: sample_display, 3: f"{result['density']:.2f}",
                     6: f"{c_calc:.2f}", 7: f"{phi_calc:.1f}"}[j]
            self.canvas.create_text((colx[j] + colx[j+1]) / 2,
                                    table_y0 + 1.5 * table_row_h,
                                    text=value, font=("Segoe UI", 9), justify="center")
            # Restore only the outer vertical boundaries of the merged cell.
            self.canvas.create_line(colx[j], table_y0, colx[j], table_y0 + 3 * table_row_h, fill="#222222")
            self.canvas.create_line(colx[j+1], table_y0, colx[j+1], table_y0 + 3 * table_row_h, fill="#222222")

    def _report_pair(self, x, y, label, value):
        """Draw one bold label/value pair in the report header."""
        self.canvas.create_text(x, y, text=label, anchor="w", font=("Segoe UI", 9, "bold"))
        self.canvas.create_text(x + 120, y, text=":", anchor="w", font=("Segoe UI", 9, "bold"))
        self.canvas.create_text(x + 135, y, text=str(value), anchor="w", font=("Segoe UI", 9))

    # ----------------------------------------------------------
    # Excel
    # ----------------------------------------------------------
    def export_all(self):
        if not self.results:
            messagebox.showwarning("Export", "Please generate all depth data first.")
            return

        try:
            pr, stresses = self._common_values()
            default_template = Path(__file__).resolve().parent / "DST Template(2).xlsx"

            if not default_template.exists():
                selected = filedialog.askopenfilename(
                    title="Select DST Excel Template",
                    filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
                )
                if not selected:
                    return
                template = Path(selected)
            else:
                template = default_template

            bh = self.bh_id_var.get().strip() or "DST_Result"
            output = filedialog.asksaveasfilename(
                title="Save Multi-Depth DST Excel Result",
                defaultextension=".xlsx",
                initialfile=f"{bh}_DST_MultiDepth.xlsx",
                filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
            )
            if not output:
                return

            result = export_multiple_dst_to_excel(
                template_path=template,
                output_path=output,
                bh_id=bh,
                pr_constant=pr,
                normal_stresses=stresses,
                results=self.results,
            )

            self.last_excel_output_dir = Path(output).resolve().parent

            messagebox.showinfo(
                "Excel Export Complete",
                f"Exported {len(self.results)} depth(s) successfully.\n\n"
                f"File: {output}\n\n"
                f"A Summary sheet contains calculated c, φ and R² for every depth."
            )
        except Exception as e:
            messagebox.showerror("Excel Export Error", str(e))

    # ----------------------------------------------------------
    # Failure-envelope image export
    # ----------------------------------------------------------
    def export_failure_envelope_plots(self):
        """Export one clean report-style PNG failure-envelope sheet for every depth."""
        if not self.results:
            messagebox.showwarning("Plot Export", "Please generate all depth test results first.")
            return

        initial_dir = str(self.last_excel_output_dir) if self.last_excel_output_dir else None
        output_dir = filedialog.askdirectory(
            title="Select Folder to Save Failure Envelope Plots",
            initialdir=initial_dir
        )
        if not output_dir:
            return

        output_dir = Path(output_dir)
        bh = self.bh_id_var.get().strip() or "DST_Result"

        def safe_name(text):
            text = str(text).strip()
            text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
            return text.strip("._") or "Result"

        created = []
        try:
            for result in self.results:
                trials = result["trials"]
                summary = result["summary"]
                sigma = [float(t.normal_stress) for t in trials]
                tau = [float(t.peak_shear_stress) for t in trials]
                calc_c = float(summary["calculated_c"])
                calc_phi = float(summary["calculated_phi"])
                r2 = float(summary["r_squared"])
                slope = math.tan(math.radians(calc_phi))

                xmax = max(1.2, math.ceil(max(sigma) * 1.20 * 10) / 10)
                ymax = max(1.2, math.ceil(max(tau) * 1.20 * 10) / 10)
                # Nice tick spacing for common 0.25/0.50/1.00/1.50 stress inputs.
                x_step = 0.2 if xmax <= 1.2 else 0.5
                y_step = 0.2 if ymax <= 1.2 else 0.5

                fig = plt.figure(figsize=(15.36, 10.24), dpi=150, facecolor="white")
                fig.subplots_adjust(0, 0, 1, 1)
                fig.patches.append(Rectangle(
                    (0.006, 0.006), 0.988, 0.988,
                    transform=fig.transFigure, fill=False,
                    linewidth=1.4, edgecolor="black"
                ))

                # Header
                fig.text(0.5, 0.965, "DIRECT SHEAR TEST – FAILURE ENVELOPE",
                         ha="center", va="center", fontsize=22,
                         fontweight="bold", color="#14255c")
                fig.text(0.5, 0.936, "IS 2720 (Part 13) – 1986",
                         ha="center", va="center", fontsize=12,
                         fontweight="bold", color="#14255c")

                # Top information table
                y_top, y_bottom = 0.842, 0.925
                x_edges = [0.008, 0.258, 0.508, 0.758, 0.992]
                for x in x_edges:
                    fig.add_artist(Line2D([x, x], [y_top, y_bottom],
                                          transform=fig.transFigure, color="black", linewidth=0.9))
                fig.add_artist(Line2D([0.008, 0.992], [y_top, y_top],
                                      transform=fig.transFigure, color="black", linewidth=0.9))
                fig.add_artist(Line2D([0.008, 0.992], [y_bottom, y_bottom],
                                      transform=fig.transFigure, color="black", linewidth=0.9))

                sample_label = result["sample_type"]
                condition = "Remoulded" if sample_label == "Remolded" else "Undisturbed"
                info = [
                    [("Borehole ID", bh), ("Depth", f"{result['depth']:.2f} m")],
                    [("Sample Type", sample_label)],
                    [("Shear Box Size", "60 mm × 60 mm")],
                    [("Condition", condition)],
                ]
                for ci, entries in enumerate(info):
                    xa, xb = x_edges[ci], x_edges[ci + 1]
                    if ci == 0:
                        for k, (lab, val) in enumerate(entries):
                            yy = 0.892 - k * 0.027
                            fig.text(xa + 0.014, yy, lab, ha="left", va="center", fontsize=10, fontweight="bold")
                            fig.text(xa + 0.115, yy, ":", ha="left", va="center", fontsize=10)
                            fig.text(xa + 0.135, yy, val, ha="left", va="center", fontsize=10)
                    else:
                        lab, val = entries[0]
                        fig.text(xa + 0.014, 0.883, lab, ha="left", va="center", fontsize=10, fontweight="bold")
                        fig.text(xa + 0.115, 0.883, ":", ha="left", va="center", fontsize=10)
                        fig.text(xa + 0.135, 0.883, val, ha="left", va="center", fontsize=10)

                # Graph title BELOW the top information table, never overlapping it.
                fig.text(0.5, 0.828,
                         "NORMAL STRESS vs PEAK SHEAR STRESS (FAILURE ENVELOPE)",
                         ha="center", va="center", fontsize=14, fontweight="bold")

                # Plot axes; leave a clean band above for the title.
                ax = fig.add_axes([0.09, 0.335, 0.87, 0.475])
                ax.set_xlim(0, xmax)
                ax.set_ylim(0, ymax)
                ax.set_xlabel("Normal Stress, σₙ (kg/cm²)", fontsize=13, fontweight="bold", labelpad=8)
                ax.xaxis.set_label_coords(0.5, -0.035)
                ax.set_ylabel("Peak Shear Stress, τ (kg/cm²)", fontsize=13, fontweight="bold")
                ax.grid(True, linestyle="--", linewidth=0.7, alpha=0.45)
                ax.tick_params(labelsize=10)
                ax.set_xticks([round(i * x_step, 10) for i in range(int(round(xmax / x_step)) + 1)])
                ax.set_yticks([round(i * y_step, 10) for i in range(int(round(ymax / y_step)) + 1)])

                first_x = min(sigma)
                first_y = calc_c + slope * first_x
                ax.plot([0, first_x], [calc_c, first_y], color="red", linewidth=2.0, linestyle="--")
                ax.plot([first_x, xmax], [first_y, calc_c + slope * xmax],
                        color="red", linewidth=2.0, label="Failure Envelope")
                ax.scatter(sigma, tau, s=105, facecolors="white", edgecolors="black",
                           linewidths=2.0, zorder=5, label="Test Results (Peak)")

                for x, yy in zip(sigma, tau):
                    ax.annotate(f"({x:.2f}, {yy:.2f})", (x, yy),
                                xytext=(0, 14), textcoords="offset points", ha="center",
                                fontsize=10, fontweight="bold", color="#1238a5")

                # Keep legend order exactly as in the requested report.
                handles, labels = ax.get_legend_handles_labels()
                if len(handles) == 2:
                    ax.legend([handles[1], handles[0]], ["Test Results (Peak)", "Failure Envelope"],
                              loc="upper left", fontsize=10, frameon=True,
                              fancybox=False, edgecolor="black")

                # Clean, non-overlapping equation box.
                eq = ("τ = c + σₙ tan φ\n"
                      f"    = {calc_c:.2f} + σₙ tan {calc_phi:.1f}°\n\n"
                      f"c = {calc_c:.2f} kg/cm²\n"
                      f"φ = {calc_phi:.1f}°\n\n"
                      f"R² = {r2:.3f}")
                ax.text(0.985, 0.035, eq, transform=ax.transAxes,
                        ha="right", va="bottom", fontsize=10.5, fontweight="bold",
                        color="black", linespacing=1.45,
                        bbox=dict(boxstyle="square,pad=0.65", facecolor="white",
                                  edgecolor="black", linewidth=1.0))

                # Bottom section title
                fig.text(0.5, 0.275, "DIRECT SHEAR TEST RESULTS",
                         ha="center", va="center", fontsize=14,
                         fontweight="bold", color="#1425a0")

                # Bottom table. Vertically merged cells are created visually by
                # hiding the internal horizontal lines and placing one value at the center.
                bottom_ax = fig.add_axes([0.008, 0.065, 0.984, 0.205])
                bottom_ax.axis("off")
                headers = [
                    "Sr. No.", "Sample Depth\n(m)", "Sample Type", "Density\n(gm/cc)",
                    "Normal Stress,\nσₙ (kg/cm²)", "Shear Stress at Failure\n(Peak), τ (kg/cm²)",
                    "Cohesion,\nc (kg/cm²)", "Angle of Internal Friction,\nφ (°)"
                ]
                body = []
                for i, trial in enumerate(trials, 1):
                    body.append([
                        str(i), f"{result['depth']:.2f}" if i == 2 else "",
                        result["sample_type"] if i == 2 else "",
                        f"{result['density']:.2f}" if i == 2 else "",
                        f"{trial.normal_stress:.2f}", f"{trial.peak_shear_stress:.2f}",
                        f"{calc_c:.2f}" if i == 2 else "",
                        f"{calc_phi:.1f}" if i == 2 else ""
                    ])

                table = bottom_ax.table(
                    cellText=body, colLabels=headers, cellLoc="center", colLoc="center",
                    loc="center", colWidths=[0.07, 0.13, 0.13, 0.11, 0.14, 0.17, 0.105, 0.145]
                )
                table.auto_set_font_size(False)
                table.set_fontsize(9.2)
                table.scale(1, 2.25)
                merged_cols = {1, 2, 3, 6, 7}
                for (r, c), cell in table.get_celld().items():
                    cell.set_edgecolor("black")
                    cell.set_linewidth(0.8)
                    cell.set_facecolor("white")
                    cell.PAD = 0.02
                    if r == 0:
                        cell.set_facecolor("#eef3fb")
                        cell.get_text().set_fontweight("bold")
                    elif c in merged_cols:
                        # Keep text only in the middle row; suppress internal horizontal borders.
                        if r in (1, 3):
                            cell.get_text().set_text("")
                        if r == 1:
                            cell.visible_edges = "TLR"
                        elif r == 2:
                            cell.visible_edges = "LR"
                        elif r == 3:
                            cell.visible_edges = "BLR"

                filename = f"{safe_name(bh)}_Depth_{result['depth']:.2f}m_Failure_Envelope.png"
                out_file = output_dir / filename
                fig.savefig(out_file, dpi=200, bbox_inches="tight", pad_inches=0.04,
                            facecolor="white", edgecolor="none")
                plt.close(fig)
                created.append(out_file)

            messagebox.showinfo(
                "Plot Export Complete",
                f"Generated {len(created)} separate failure-envelope plot(s).\n\nSaved in:\n{output_dir}"
            )
            self.status_var.set(f"Exported {len(created)} separate failure-envelope plot image(s).")
        except Exception as e:
            messagebox.showerror("Plot Export Error", str(e))

    def import_from_excel(self):
        """Import common settings and multiple depth rows from an Excel input template."""
        input_path = filedialog.askopenfilename(
            title="Select DST Excel Input File",
            filetypes=[
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro-Enabled Workbook", "*.xlsm"),
                ("All Files", "*.*"),
            ],
        )
        if not input_path:
            return

        try:
            wb = load_workbook(input_path, data_only=True)
            if "Input" not in wb.sheetnames:
                raise ValueError("Worksheet 'Input' was not found in the selected Excel file.")

            ws = wb["Input"]

            def cell_value(address):
                return ws[address].value

            # Common settings
            bh = cell_value("B2")
            pr = cell_value("B3")
            n1 = cell_value("B4")
            n2 = cell_value("B5")
            n3 = cell_value("B6")
            seed = cell_value("B7")

            if bh is None or str(bh).strip() == "":
                raise ValueError("BH ID is required in cell B2.")
            for name, value in (
                ("PR Constant", pr),
                ("Normal Stress 1", n1),
                ("Normal Stress 2", n2),
                ("Normal Stress 3", n3),
            ):
                if value is None or str(value).strip() == "":
                    raise ValueError(f"{name} is required in the Excel input.")

            self.bh_id_var.set(str(bh).strip())
            self.pr_constant.set(str(pr))
            self.normal_1.set(str(n1))
            self.normal_2.set(str(n2))
            self.normal_3.set(str(n3))
            self.random_seed.set("AUTO" if seed is None or str(seed).strip() == "" else str(seed))

            # Read depth rows. Header row = 10, data starts at row 11.
            headers = {}
            for col in range(1, ws.max_column + 1):
                value = ws.cell(10, col).value
                if value is not None:
                    headers[str(value).strip().lower()] = col

            required_headers = [
                "depth (m)",
                "soil type",
                "sample type",
                "density (g/cc)",
                "target c (kg/cm²)",
                "target φ (°)",
            ]
            missing = [h for h in required_headers if h.lower() not in headers]
            if missing:
                raise ValueError(
                    "Missing required Excel column(s): " + ", ".join(missing)
                )

            imported_rows = []
            for r in range(11, ws.max_row + 1):
                values = {
                    key: ws.cell(r, col).value
                    for key, col in headers.items()
                }

                # Ignore completely blank rows.
                if all(v is None or str(v).strip() == "" for v in values.values()):
                    continue

                depth = values.get("depth (m)")
                soil = values.get("soil type")
                sample_type = values.get("sample type")
                density = values.get("density (g/cc)")
                c = values.get("target c (kg/cm²)")
                phi = values.get("target φ (°)")

                if any(v is None or str(v).strip() == "" for v in (depth, soil, sample_type, density, c, phi)):
                    raise ValueError(
                        f"Excel row {r}: Depth, Soil Type, Sample Type, Density, Target c and Target φ are required."
                    )

                imported_rows.append({
                    "depth": str(depth),
                    "soil": str(soil),
                    "sample_type": str(sample_type),
                    "density": str(density),
                    "weight": "",
                    "c": str(c),
                    "phi": str(phi),
                    "behaviour": "Enter density",
                })

            if not imported_rows:
                raise ValueError("No depth data was found. Enter data from row 11 onward.")

            # Clear existing GUI depth rows and rebuild them from Excel.
            for row in self.rows:
                for widget in row["widgets"]:
                    widget.destroy()

            for child in list(self.table_frame.winfo_children()):
                info = child.grid_info()
                if info and int(info.get("row", 0)) > 0:
                    child.destroy()

            self.rows = []
            for item in imported_rows:
                row = {
                    "depth": tk.StringVar(value=item["depth"]),
                    "soil": tk.StringVar(value=item["soil"]),
                    "sample_type": tk.StringVar(value=item["sample_type"]),
                    "density": tk.StringVar(value=item["density"]),
                    "weight": tk.StringVar(value=item["weight"]),
                    "c": tk.StringVar(value=item["c"]),
                    "phi": tk.StringVar(value=item["phi"]),
                    "behaviour": tk.StringVar(value="Enter density"),
                    "widgets": [],
                }
                self.rows.append(row)

            for i, row in enumerate(self.rows, 1):
                self._draw_depth_row(i, row)
                self.update_row_from_density(row)
                self.update_row_behaviour(row)

            self.results = []
            self.selected_result = None
            self.refresh_results()
            self.canvas.delete("all")
            self.status_var.set(
                f"Imported {len(self.rows)} depth(s) from Excel. Review/validate, then generate."
            )
            messagebox.showinfo(
                "Excel Import Complete",
                f"Imported {len(self.rows)} depth(s) successfully.\n\n"
                "Review the imported values and click GENERATE ALL DEPTHS."
            )

        except Exception as e:
            messagebox.showerror("Excel Import Error", str(e))

    def reset_form(self):
        self.bh_id_var.set("")
        self.pr_constant.set("0.325")
        self.normal_1.set("0.50")
        self.normal_2.set("1.00")
        self.normal_3.set("1.50")
        self.random_seed.set("AUTO")
        self.results = []
        self.selected_result = None
        self.last_excel_output_dir = None
        self.refresh_results()
        self.canvas.delete("all")

        for r in self.rows:
            for w in r["widgets"]:
                w.destroy()
        for child in list(self.table_frame.winfo_children()):
            info = child.grid_info()
            if info and int(info.get("row", 0)) > 0:
                child.destroy()
        self.rows = []
        self.add_depth_row()
        self.status_var.set("Enter one or more depths, then generate.")


if __name__ == "__main__":
    root = tk.Tk()
    app = DSTGeneratorApp(root)
    root.mainloop()
