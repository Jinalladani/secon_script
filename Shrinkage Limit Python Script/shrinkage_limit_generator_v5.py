import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import random
import math

APP_TITLE = "Shrinkage Limit Test – Observation Sheet Generator V5.2"

CONTAINERS = [f"SL_{i:02d}" for i in range(1, 11)]
SOIL_CLASSES = ["CL", "CI", "CH"]
SAMPLE_TYPES = ["UDS", "SPT", "DS", "Block Sample", "Other"]

# User-provided shrinkage dish/container details.
# Weight: g; Diameter/Height: mm.
DISHES = {
    "SL_01": {"weight": 37.10, "diameter": 44.34, "height": 14.20},
    "SL_02": {"weight": 36.38, "diameter": 44.36, "height": 14.18},
    "SL_03": {"weight": 37.29, "diameter": 44.34, "height": 15.08},
    "SL_04": {"weight": 41.98, "diameter": 45.62, "height": 16.34},
    "SL_05": {"weight": 47.97, "diameter": 45.50, "height": 16.40},
    "SL_06": {"weight": 51.18, "diameter": 45.21, "height": 16.24},
    "SL_07": {"weight": 52.38, "diameter": 45.22, "height": 16.53},
    "SL_08": {"weight": 47.07, "diameter": 44.41, "height": 16.08},
    "SL_09": {"weight": 57.57, "diameter": 44.89, "height": 16.36},
    "SL_10": {"weight": 33.27, "diameter": 44.07, "height": 14.98},
}

MERCURY_DENSITY = 13.6  # g/cm³, approximate

# ----------------------------------------------------------------------
# Soil-classification logic
# ----------------------------------------------------------------------
def classify_from_ll(ll):
    """
    IS-style consistency classification used by this generator:
      LL < 35       -> CL (low plasticity)
      35 <= LL <=50 -> CI (intermediate plasticity)
      LL > 50       -> CH (high plasticity)

    The exact 35 and 50 boundaries are assigned to the intermediate
    class so there is no gap in classification.
    """
    if ll < 35.0:
        return "CL"
    elif ll <= 50.0:
        return "CI"
    else:
        return "CH"


def plasticity_description(ll):
    if ll < 35.0:
        return "Low plasticity"
    elif ll <= 50.0:
        return "Intermediate plasticity"
    return "High plasticity"


def validate_soil_class(ll, soil_class):
    expected = classify_from_ll(ll)
    if soil_class != expected:
        return (
            False,
            f"LL = {ll:g}% corresponds to {expected} "
            f"({plasticity_description(ll)}), not {soil_class}."
        )
    return True, ""


def dish_volume_cm3(diameter_mm, height_mm):
    """Cylindrical volume from supplied dish dimensions."""
    d_cm = diameter_mm / 10.0
    h_cm = height_mm / 10.0
    return math.pi * (d_cm ** 2) / 4.0 * h_cm


def safe_float(value, name):
    try:
        x = float(value)
        if x < 0:
            raise ValueError
        return x
    except Exception:
        raise ValueError(f"Enter a valid positive number for {name}.")


# ----------------------------------------------------------------------
# Synthetic observation model – Version 5
# ----------------------------------------------------------------------
def generate_observation(ll, pl, soil_class, dish):
    """
    Generate a synthetic shrinkage-limit observation for template/testing.

    Important:
      This is NOT a prediction equation prescribed by IS code and it is
      NOT a substitute for measured laboratory observations.

    Model sequence:
      LL + PL -> PI -> class-consistent expected SL band
      -> dish geometry -> wet/dry pat volume
      -> dry soil mass -> water loss -> recorded balance readings
      -> mercury displacement -> final calculated SL.

    Version 5 deliberately introduces small independent measurement
    effects before rounding, so the final SL is not simply forced to a
    preselected number.
    """

    pi = ll - pl

    # Broad synthetic starting bands for SL below PL.
    # These are modelling assumptions for generating realistic-looking
    # test data, NOT code correlations.
    if soil_class == "CL":
        gap_low, gap_high = 2.5, 6.0
        m2_low, m2_high = 27.0, 40.0
        shrink_ratio_low, shrink_ratio_high = 0.035, 0.12
    elif soil_class == "CI":
        gap_low, gap_high = 4.0, 9.0
        m2_low, m2_high = 29.0, 44.0
        shrink_ratio_low, shrink_ratio_high = 0.045, 0.16
    else:
        gap_low, gap_high = 7.0, 15.0
        m2_low, m2_high = 31.0, 48.0
        shrink_ratio_low, shrink_ratio_high = 0.055, 0.20

    # PI effect: higher PI generally permits a somewhat larger
    # separation between PL and SL in the synthetic population.
    # Keep the influence deliberately modest.
    pi_effect = max(-1.0, min(4.0, (pi - 15.0) * 0.08))

    # Depth/sample-to-sample natural scatter represented by random effect.
    natural_scatter = random.gauss(0.0, 0.65)

    gap = (
        random.uniform(gap_low, gap_high)
        + pi_effect
        + natural_scatter
    )

    # Keep SL below PL with a practical lower bound.
    sl_expected = max(3.5, min(pl - 1.0, pl - gap))

    # Dish volume from actual supplied dimensions.
    v1_true = dish_volume_cm3(
        dish["diameter"],
        dish["height"]
    )

    # Small geometry/readout effect. V1 is still anchored to the actual
    # dish dimensions, but a real observation may differ slightly from
    # ideal geometry/reading.
    v1_observed = v1_true + random.gauss(0.0, 0.03)
    v1_observed = max(0.80 * v1_true, v1_observed)

    # Generate dry-pat volume reduction. Higher-plasticity soils are
    # allowed somewhat greater volume change, but not deterministically.
    shrink_ratio = random.uniform(
        shrink_ratio_low,
        shrink_ratio_high
    )
    shrink_ratio += random.gauss(0.0, 0.008)
    shrink_ratio = max(0.025, min(0.22, shrink_ratio))

    v2_true = v1_observed * (1.0 - shrink_ratio)

    # Small mercury/displacement reading uncertainty.
    mercury_mass_true = v2_true * MERCURY_DENSITY
    mercury_mass_observed = (
        mercury_mass_true
        + random.gauss(0.0, 0.35)
    )
    mercury_mass_observed = max(
        1.0,
        mercury_mass_observed
    )

    v2_observed = (
        mercury_mass_observed
        / MERCURY_DENSITY
    )

    # Dry soil mass varies with soil class and trial preparation.
    m2_true = random.uniform(
        m2_low,
        m2_high
    )

    # Slight preparation/mass scatter.
    m2_true *= random.uniform(0.97, 1.03)

    # Use the expected SL only as a centre point. Generate water loss
    # around that centre rather than algebraically forcing final SL.
    #
    # Rearranged shrinkage-limit equation:
    #
    # SL = [((M1-M2) - (V1-V2)) / M2] * 100
    #
    # => M1-M2 = (SL/100)*M2 + (V1-V2)
    #
    target_water_loss = (
        (sl_expected / 100.0) * m2_true
        + (v1_observed - v2_observed)
    )

    # Independent preparation/moisture variation.
    water_loss = target_water_loss + random.gauss(0.0, 0.12)
    water_loss = max(
        0.10,
        water_loss
    )

    m1_true = m2_true + water_loss

    # ------------------------------------------------------------------
    # Simulated balance readings
    # ------------------------------------------------------------------
    # W1 remains the actual dish weight supplied by the user.
    w1 = dish["weight"]

    # W2 and W3 have small balance/readability effects.
    w2 = (
        w1
        + m1_true
        + random.gauss(0.0, 0.015)
    )
    w3 = (
        w1
        + m2_true
        + random.gauss(0.0, 0.015)
    )

    # 0.01 g readability.
    w1 = round(w1, 2)
    w2 = round(w2, 2)
    w3 = round(w3, 2)

    # ------------------------------------------------------------------
    # Calculated quantities from recorded observations
    # ------------------------------------------------------------------
    calc_m1 = round(w2 - w1, 2)
    calc_m2 = round(w3 - w1, 2)

    v1 = round(v1_observed, 2)
    mercury_mass = round(mercury_mass_observed, 2)

    # V2 must be calculated from the recorded mercury mass.
    calc_v2 = round(
        mercury_mass / MERCURY_DENSITY,
        2
    )

    # Final SL is calculated only after the observations have been
    # rounded, so the result naturally differs slightly from the model
    # centre.
    denominator = calc_m2

    if denominator <= 0:
        denominator = 0.01

    sl = (
        (
            (calc_m1 - calc_m2)
            - (v1 - calc_v2)
        )
        / denominator
        * 100.0
    )

    # Prevent an occasional extreme value caused by measurement scatter.
    sl = max(
        0.5,
        min(
            max(0.5, pl - 0.5),
            sl
        )
    )

    return {
        "W1": w1,
        "W2": w2,
        "W3": w3,
        "M1": calc_m1,
        "M2": calc_m2,
        "V1": v1,
        "Mercury Mass": mercury_mass,
        "V2": calc_v2,
        "SL": round(sl, 2),
    }


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1500x880")
        self.minsize(1280, 740)

        self.bh_var = tk.StringVar()

        self.depth_var = tk.StringVar()
        self.sample_type_var = tk.StringVar(value="UDS")
        self.soil_class_var = tk.StringVar(value="CL")
        self.container_var = tk.StringVar(value="SL_01")
        self.ll_var = tk.StringVar()
        self.pl_var = tk.StringVar()

        self.build_ui()

    def build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # ==============================================================
        # BH INFORMATION
        # ==============================================================
        info = ttk.LabelFrame(
            self,
            text="Test / Project Information",
            padding=10
        )
        info.pack(fill="x", padx=12, pady=(12, 6))

        ttk.Label(
            info,
            text="BH ID"
        ).grid(row=0, column=0, sticky="w")

        ttk.Entry(
            info,
            textvariable=self.bh_var,
            width=24
        ).grid(row=0, column=1, padx=(6, 30))

        ttk.Label(
            info,
            text="BH ID is common for all depth entries."
        ).grid(row=0, column=2, sticky="w")

        # ==============================================================
        # REQUIRED INPUT – MULTIPLE DEPTHS
        # ==============================================================
        input_frame = ttk.LabelFrame(
            self,
            text="Required Input – Multiple Depths",
            padding=10
        )
        input_frame.pack(fill="x", padx=12, pady=6)

        fields = [
            ("Depth (m)", 0),
            ("Sample Type", 2),
            ("Soil Class", 4),
            ("Dish No.", 6),
            ("LL (%)", 8),
            ("PL (%)", 10),
        ]

        for text, col in fields:
            ttk.Label(
                input_frame,
                text=text
            ).grid(row=0, column=col, sticky="w")

        ttk.Entry(
            input_frame,
            textvariable=self.depth_var,
            width=14
        ).grid(row=0, column=1, padx=5)

        ttk.Combobox(
            input_frame,
            textvariable=self.sample_type_var,
            values=SAMPLE_TYPES,
            state="readonly",
            width=14
        ).grid(row=0, column=3, padx=5)

        ttk.Combobox(
            input_frame,
            textvariable=self.soil_class_var,
            values=SOIL_CLASSES,
            state="readonly",
            width=9
        ).grid(row=0, column=5, padx=5)

        self.container_box = ttk.Combobox(
            input_frame,
            textvariable=self.container_var,
            values=CONTAINERS,
            state="readonly",
            width=11
        )
        self.container_box.grid(row=0, column=7, padx=5)
        self.container_box.bind(
            "<<ComboboxSelected>>",
            self.show_dish_details
        )

        ttk.Entry(
            input_frame,
            textvariable=self.ll_var,
            width=10
        ).grid(row=0, column=9, padx=5)

        ttk.Entry(
            input_frame,
            textvariable=self.pl_var,
            width=10
        ).grid(row=0, column=11, padx=5)

        ttk.Button(
            input_frame,
            text="Add Depth",
            command=self.add_depth
        ).grid(row=0, column=12, padx=(15, 5))

        ttk.Button(
            input_frame,
            text="Remove Selected",
            command=self.remove_selected
        ).grid(row=0, column=13, padx=5)

        ttk.Button(
            input_frame,
            text="Clear Inputs",
            command=self.clear_inputs
        ).grid(row=0, column=14, padx=5)

        self.dish_info_var = tk.StringVar()
        ttk.Label(
            input_frame,
            textvariable=self.dish_info_var,
            foreground="#555555"
        ).grid(
            row=1,
            column=0,
            columnspan=15,
            sticky="w",
            pady=(8, 0)
        )

        ttk.Label(
            input_frame,
            text="Classification rule: LL < 35 = CL | 35 ≤ LL ≤ 50 = CI | LL > 50 = CH",
            foreground="#555555"
        ).grid(
            row=2,
            column=0,
            columnspan=15,
            sticky="w",
            pady=(4, 0)
        )

        self.show_dish_details()

        self.input_tree = ttk.Treeview(
            input_frame,
            columns=(
                "depth",
                "sample",
                "soil",
                "container",
                "ll",
                "pl",
                "pi"
            ),
            show="headings",
            height=5
        )

        headings = {
            "depth": "Depth (m)",
            "sample": "Sample Type",
            "soil": "Soil Class",
            "container": "Dish No.",
            "ll": "LL (%)",
            "pl": "PL (%)",
            "pi": "PI (%)",
        }

        widths = {
            "depth": 120,
            "sample": 145,
            "soil": 95,
            "container": 120,
            "ll": 90,
            "pl": 90,
            "pi": 90,
        }

        for col in self.input_tree["columns"]:
            self.input_tree.heading(col, text=headings[col])
            self.input_tree.column(
                col,
                width=widths[col],
                anchor="center"
            )

        self.input_tree.grid(
            row=3,
            column=0,
            columnspan=15,
            sticky="ew",
            pady=(10, 0)
        )

        # ==============================================================
        # DISH DETAILS
        # ==============================================================
        dish_frame = ttk.LabelFrame(
            self,
            text="Shrinkage Limit Dish / Container Details",
            padding=8
        )
        dish_frame.pack(fill="x", padx=12, pady=6)

        dish_cols = (
            "sr",
            "dish",
            "weight",
            "diameter",
            "height",
            "volume"
        )

        self.dish_tree = ttk.Treeview(
            dish_frame,
            columns=dish_cols,
            show="headings",
            height=4
        )

        dish_headings = {
            "sr": "Sr. No.",
            "dish": "Dish No.",
            "weight": "Weight (g)",
            "diameter": "Diameter (mm)",
            "height": "Height (mm)",
            "volume": "Calculated Dish Volume (cm³)"
        }

        for col in dish_cols:
            self.dish_tree.heading(
                col,
                text=dish_headings[col]
            )
            self.dish_tree.column(
                col,
                width=185 if col == "volume" else 125,
                anchor="center"
            )

        for i, container in enumerate(CONTAINERS, 1):
            d = DISHES[container]
            volume = dish_volume_cm3(
                d["diameter"],
                d["height"]
            )

            self.dish_tree.insert(
                "",
                "end",
                values=(
                    i,
                    container,
                    f'{d["weight"]:.2f}',
                    f'{d["diameter"]:.2f}',
                    f'{d["height"]:.2f}',
                    f'{volume:.2f}'
                )
            )

        self.dish_tree.pack(fill="x")

        # ==============================================================
        # ACTIONS
        # ==============================================================
        actions = ttk.Frame(
            self,
            padding=(12, 8)
        )
        actions.pack(fill="x")

        ttk.Button(
            actions,
            text="Generate Observation Data",
            command=self.generate
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            actions,
            text="Export to Excel",
            command=self.export_excel
        ).pack(side="left", padx=8)

        ttk.Button(
            actions,
            text="Clear Results",
            command=self.clear_results
        ).pack(side="left", padx=8)

        ttk.Label(
            actions,
            text=(
                "Synthetic data for testing/template development only; "
                "not a substitute for measured laboratory observations."
            ),
            foreground="#555555"
        ).pack(side="right")

        # ==============================================================
        # OBSERVATION TABLE
        # ==============================================================
        result_frame = ttk.LabelFrame(
            self,
            text="Shrinkage Limit – Observation Sheet",
            padding=8
        )
        result_frame.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(0, 12)
        )

        columns = [
            "BH ID",
            "Depth (m)",
            "Sample Type",
            "Soil Class",
            "Dish No.",
            "Empty Container Weight (W1), g",
            "Container + Wet Soil Weight (W2), g",
            "Container + Dry Soil Weight (W3), g",
            "Wet Soil Mass (M1), g",
            "Dry Soil Mass (M2), g",
            "Volume of Wet Soil Pat (V1), cm³",
            "Mass of Displaced Mercury (Mv), g",
            "Volume of Dry Soil Pat (V2), cm³",
            "Shrinkage Limit (SL), %"
        ]

        self.result_tree = ttk.Treeview(
            result_frame,
            columns=columns,
            show="headings"
        )

        for c in columns:
            self.result_tree.heading(c, text=c)
            self.result_tree.column(
                c,
                width=155,
                minwidth=115,
                anchor="center"
            )

        for c in columns[5:]:
            self.result_tree.column(c, width=205)

        ybar = ttk.Scrollbar(
            result_frame,
            orient="vertical",
            command=self.result_tree.yview
        )
        xbar = ttk.Scrollbar(
            result_frame,
            orient="horizontal",
            command=self.result_tree.xview
        )

        self.result_tree.configure(
            yscrollcommand=ybar.set,
            xscrollcommand=xbar.set
        )

        self.result_tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")

        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

    def show_dish_details(self, event=None):
        container = self.container_var.get()

        if container in DISHES:
            d = DISHES[container]
            volume = dish_volume_cm3(
                d["diameter"],
                d["height"]
            )

            self.dish_info_var.set(
                f'{container}: Weight = {d["weight"]:.2f} g | '
                f'Diameter = {d["diameter"]:.2f} mm | '
                f'Height = {d["height"]:.2f} mm | '
                f'Calculated cylindrical volume = {volume:.2f} cm³'
            )

    def add_depth(self):
        depth = self.depth_var.get().strip()
        sample = self.sample_type_var.get().strip()
        soil = self.soil_class_var.get().strip()
        container = self.container_var.get().strip()
        ll_text = self.ll_var.get().strip()
        pl_text = self.pl_var.get().strip()

        if not depth:
            messagebox.showerror(
                "Missing Depth",
                "Enter the depth."
            )
            return

        try:
            safe_float(depth, "Depth")
        except ValueError as e:
            messagebox.showerror("Invalid Depth", str(e))
            return

        if not sample:
            messagebox.showerror(
                "Missing Sample Type",
                "Select Sample Type."
            )
            return

        if soil not in SOIL_CLASSES:
            messagebox.showerror(
                "Missing Soil Class",
                "Select CL, CI or CH."
            )
            return

        if container not in CONTAINERS:
            messagebox.showerror(
                "Missing Dish",
                "Select SL_01 to SL_10."
            )
            return

        try:
            ll = safe_float(ll_text, "LL")
            pl = safe_float(pl_text, "PL")
        except ValueError as e:
            messagebox.showerror("Invalid LL/PL", str(e))
            return

        if ll <= 0 or pl <= 0:
            messagebox.showerror(
                "Invalid LL/PL",
                "LL and PL must be greater than zero."
            )
            return

        if pl >= ll:
            messagebox.showerror(
                "Invalid LL/PL",
                "PL must be less than LL."
            )
            return

        # IMPORTANT: Soil class must agree with LL.
        valid, error = validate_soil_class(ll, soil)
        if not valid:
            messagebox.showerror(
                "Soil Classification Mismatch",
                error
            )
            return

        # Prevent duplicate dish assignment.
        for item in self.input_tree.get_children():
            values = self.input_tree.item(item, "values")
            if values[3] == container:
                messagebox.showerror(
                    "Duplicate Dish",
                    f"{container} is already assigned to another depth."
                )
                return

        pi = ll - pl

        self.input_tree.insert(
            "",
            "end",
            values=(
                depth,
                sample,
                soil,
                container,
                round(ll, 2),
                round(pl, 2),
                round(pi, 2)
            )
        )

        self.depth_var.set("")
        self.ll_var.set("")
        self.pl_var.set("")

    def remove_selected(self):
        for item in self.input_tree.selection():
            self.input_tree.delete(item)

    def clear_inputs(self):
        for item in self.input_tree.get_children():
            self.input_tree.delete(item)

    def clear_results(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

    def generate(self):
        bh = self.bh_var.get().strip()

        if not bh:
            messagebox.showerror(
                "Missing BH ID",
                "Enter BH ID."
            )
            return

        entries = self.input_tree.get_children()

        if not entries:
            messagebox.showerror(
                "No Depths",
                "Add at least one depth."
            )
            return

        self.clear_results()

        for item in entries:
            depth, sample, soil, container, ll, pl, pi = (
                self.input_tree.item(item, "values")
            )

            ll = float(ll)
            pl = float(pl)

            # Safety check again before generating.
            valid, error = validate_soil_class(ll, soil)
            if not valid:
                messagebox.showerror(
                    "Soil Classification Mismatch",
                    error
                )
                return

            dish = DISHES[container]

            obs = generate_observation(
                ll,
                pl,
                soil,
                dish
            )

            values = (
                bh,
                depth,
                sample,
                soil,
                container,
                obs["W1"],
                obs["W2"],
                obs["W3"],
                obs["M1"],
                obs["M2"],
                obs["V1"],
                obs["Mercury Mass"],
                obs["V2"],
                obs["SL"]
            )

            self.result_tree.insert(
                "",
                "end",
                values=values
            )

        messagebox.showinfo(
            "Completed",
            f"Generated observations for {len(entries)} depth(s)."
        )

    def export_excel(self):
        rows = self.result_tree.get_children()

        if not rows:
            messagebox.showwarning(
                "No Data",
                "Generate observation data first."
            )
            return

        path = filedialog.asksaveasfilename(
            title="Save Shrinkage Limit Observation Sheet",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Workbook", "*.xlsx")
            ]
        )

        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Shrinkage Limit"

        bh = self.bh_var.get().strip()

        # ==============================================================
        # EXCEL STYLES
        # No yellow/highlight fill is used.
        # ==============================================================
        thin = Side(style="thin")
        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        title_font = Font(
            bold=True,
            size=15
        )
        header_font = Font(bold=True)
        normal_font = Font(size=10)

        # ==============================================================
        # TITLE
        # ==============================================================
        ws["A1"] = "SHRINKAGE LIMIT TEST – OBSERVATION SHEET"
        ws["A1"].font = title_font
        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=14
        )
        ws["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws["A3"] = "BH ID"
        ws["B3"] = bh
        ws["A3"].font = header_font

        # ==============================================================
        # REQUIRED INPUT – NO HIGHLIGHT
        # ==============================================================
        ws["A4"] = "Required Input – Multiple Depths"
        ws["A4"].font = header_font

        input_headers = [
            "Depth (m)",
            "Sample Type",
            "Soil Class",
            "Dish No.",
            "LL (%)",
            "PL (%)",
            "PI (%)"
        ]

        input_header_row = 5

        for c, h in enumerate(input_headers, 1):
            cell = ws.cell(
                input_header_row,
                c,
                h
            )
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        input_rows = self.input_tree.get_children()

        for r, item in enumerate(
            input_rows,
            input_header_row + 1
        ):
            depth, sample, soil, container, ll, pl, pi = (
                self.input_tree.item(item, "values")
            )

            values = [
                float(depth) if str(depth).replace(".", "", 1).isdigit() else depth,
                sample,
                soil,
                container,
                float(ll),
                float(pl),
                float(pi)
            ]

            for c, value in enumerate(values, 1):
                cell = ws.cell(r, c, value)
                cell.font = normal_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                cell.border = border

        # Numeric formats for input section.
        for r in range(
            input_header_row + 1,
            input_header_row + 1 + len(input_rows)
        ):
            ws.cell(r, 1).number_format = "0.00"
            ws.cell(r, 5).number_format = "0.00"
            ws.cell(r, 6).number_format = "0.00"
            ws.cell(r, 7).number_format = "0.00"

        # ==============================================================
        # OBSERVATION TABLE
        # ==============================================================
        observation_start = (
            input_header_row + len(input_rows) + 2
        )

        ws.cell(
            observation_start,
            1,
            "Shrinkage Limit – Observation"
        ).font = header_font

        headers = [
            "BH ID",
            "Depth (m)",
            "Sample Type",
            "Soil Class",
            "Dish No.",
            "Empty Container Weight (W1), g",
            "Container + Wet Soil Weight (W2), g",
            "Container + Dry Soil Weight (W3), g",
            "Wet Soil Mass (M1), g",
            "Dry Soil Mass (M2), g",
            "Volume of Wet Soil Pat (V1), cm³",
            "Mass of Displaced Mercury (Mv), g",
            "Volume of Dry Soil Pat (V2), cm³",
            "Shrinkage Limit (SL), %"
        ]

        header_row = observation_start + 1

        for col, header in enumerate(headers, 1):
            cell = ws.cell(
                header_row,
                col,
                header
            )
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        for r, item in enumerate(
            rows,
            header_row + 1
        ):
            values = self.result_tree.item(
                item,
                "values"
            )

            for c, value in enumerate(values, 1):
                cell = ws.cell(r, c, value)
                cell.font = normal_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                cell.border = border

        # All numerical observation values are stored as numeric cells,
        # not strings.
        for r in range(
            header_row + 1,
            header_row + 1 + len(rows)
        ):
            ws.cell(r, 2).number_format = "0.00"
            for c in range(6, 11):
                ws.cell(r, c).number_format = "0.00"
            ws.cell(r, 11).number_format = "0.00"
            ws.cell(r, 12).number_format = "0.00"
            ws.cell(r, 13).number_format = "0.00"
            ws.cell(r, 14).number_format = "0.00"

        # ==============================================================
        # DISH DETAILS
        # ==============================================================
        dish_start = (
            header_row + len(rows) + 3
        )

        ws.cell(
            dish_start,
            1,
            "Shrinkage Limit Dish / Container Details"
        ).font = header_font

        dish_headers = [
            "Sr. No.",
            "Dish No.",
            "Weight (g)",
            "Diameter (mm)",
            "Height (mm)",
            "Calculated Dish Volume (cm³)"
        ]

        dish_header_row = dish_start + 1

        for c, h in enumerate(dish_headers, 1):
            cell = ws.cell(
                dish_header_row,
                c,
                h
            )
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

        for i, container in enumerate(CONTAINERS, 1):
            d = DISHES[container]
            row = dish_header_row + i

            vals = [
                i,
                container,
                d["weight"],
                d["diameter"],
                d["height"],
                round(
                    dish_volume_cm3(
                        d["diameter"],
                        d["height"]
                    ),
                    2
                )
            ]

            for c, value in enumerate(vals, 1):
                cell = ws.cell(row, c, value)
                cell.font = normal_font
                cell.alignment = Alignment(
                    horizontal="center"
                )
                cell.border = border

            ws.cell(row, 3).number_format = "0.00"
            ws.cell(row, 4).number_format = "0.00"
            ws.cell(row, 5).number_format = "0.00"
            ws.cell(row, 6).number_format = "0.00"

        # ==============================================================
        # FORMATTING
        # ==============================================================
        widths = {
            "A": 15,
            "B": 16,
            "C": 16,
            "D": 14,
            "E": 13,
            "F": 29,
            "G": 33,
            "H": 33,
            "I": 23,
            "J": 23,
            "K": 29,
            "L": 30,
            "M": 29,
            "N": 24,
        }

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        ws.row_dimensions[1].height = 24
        ws.row_dimensions[header_row].height = 58

        # Explicitly remove fills from the entire used range.
        for row in ws.iter_rows():
            for cell in row:
                cell.fill = cell.fill.copy(fill_type=None)

        ws.freeze_panes = f"A{header_row + 1}"
        ws.sheet_view.showGridLines = False

        try:
            wb.save(path)
        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"Could not save Excel file:\n{e}"
            )
            return

        messagebox.showinfo(
            "Export Complete",
            f"Observation sheet saved to:\n{path}"
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
