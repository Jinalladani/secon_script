import sys
import subprocess
import importlib.util
import math
import random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def ensure_openpyxl():
    if importlib.util.find_spec("openpyxl") is not None:
        return True

    answer = messagebox.askyesno(
        "Missing package",
        "Excel support requires 'openpyxl'.\n\n"
        "Click YES to install it automatically."
    )
    if not answer:
        return False

    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "openpyxl"]
        )
        return importlib.util.find_spec("openpyxl") is not None
    except Exception as exc:
        messagebox.showerror(
            "Installation failed",
            "Automatic installation failed.\n\n"
            "Run this in Command Prompt:\n\n"
            "python -m pip install openpyxl\n\n"
            f"Details: {exc}"
        )
        return False


if not ensure_openpyxl():
    raise SystemExit

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


KPA_PER_KGCM2 = 98.0665


def num(value, default=0.0):
    if value is None or str(value).strip() == "":
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def clamp(value, low, high):
    return max(low, min(high, value))


def piecewise_k_pi(PI):
    """
    PI-dependent SPT/Su coefficient.
    This is deliberately a smooth engineering approximation based
    on commonly reported PI-dependent SPT correlations, not a
    universal laboratory equation.
    """
    if PI < 15:
        return 6.5
    if PI < 20:
        return 6.5 + (5.5 - 6.5) * (PI - 15) / 5
    if PI < 25:
        return 5.5 + (4.8 - 5.5) * (PI - 20) / 5
    if PI < 30:
        return 4.8 + (4.5 - 4.8) * (PI - 25) / 5
    if PI < 40:
        return 4.5 + (4.4 - 4.5) * (PI - 30) / 10
    return 4.3


def classify_soil(fines, sand, gravel, clay, silt):
    if fines < 5:
        return "Clean Granular Soil"
    if fines < 35:
        if sand >= gravel:
            return "Silty/Clayey Sand"
        return "Gravelly Soil with Fines"
    if clay > silt:
        return "Clayey Soil"
    return "Silty Soil"


def effective_overburden_kgcm2(depth_m):
    """
    Effective overburden pressure at SPT depth + 0.30 m.

    User-specified average bulk density:
      1.85 gm/cc up to 10 m
      2.00 gm/cc from 10 m onward

    Groundwater level is at ground level, so the pressure below the
    groundwater table is treated as effective overburden using the
    supplied average density model.

    NOTE:
    The 0.30 m increment is included because IS 2131 section 5.6.3
    specifies effective overburden pressure at (SPT depth + 0.30 m).
    """
    depth_m = max(0.0, depth_m)
    z = depth_m + 0.30

    if z <= 10.0:
        return 1.85 * z * 0.0980665

    first_10m = 1.85 * 10.0 * 0.0980665
    remaining = 2.00 * (z - 10.0) * 0.0980665
    return first_10m + remaining


def calculate_corrected_spt(field_n, depth_m, soil_is_fine_sand_or_silt):
    """
    IS 2131 correction sequence based on the supplied standard extract.

    1. Effective overburden pressure sigma'_0 at SPT depth + 0.30 m.
    2. CN = 0.77 log10(pa / sigma'_0)
    3. Limit CN to 0.40 <= CN <= 2.00.
    4. N1 = CN * Field N.
    5. Dilatancy correction only when the stratum is fine sand or silt,
       below the groundwater table, and N1 > 15:
           N1' = 15 + 0.5*(N1 - 15)
       Otherwise N1' = N1.

    pa and sigma'_0 are both represented in kg/cm2.
    """
    field_n = max(0.0, field_n)
    depth_m = max(0.0, depth_m)

    sigma0 = effective_overburden_kgcm2(depth_m)

    # Atmospheric pressure in kg/cm2.
    pa = 1.0332

    # Prevent log10(0) at an impossible zero-pressure location.
    sigma0_for_log = max(sigma0, 1.0e-9)

    cn_raw = 0.77 * math.log10((20.0 * pa) / sigma0_for_log)
    cn = clamp(cn_raw, 0.40, 2.00)

    N1 = field_n * cn

    # Groundwater is at ground level, therefore the sample is below
    # the water table for positive depth. The material criterion is
    # still applied: fine sand or silt only.
    below_water_table = depth_m > 0.0

    if soil_is_fine_sand_or_silt and below_water_table and N1 > 15.0:
        N1_prime = 15.0 + 0.5 * (N1 - 15.0)
        dilatancy_applied = True
    else:
        N1_prime = N1
        dilatancy_applied = False

    density_gmcc = 1.85 if depth_m <= 10.0 else 2.00

    return {
        "Density_gmcc": density_gmcc,
        "Effective_Overburden_kgcm2": sigma0,
        "Pa_kgcm2": pa,
        "CN_Raw": cn_raw,
        "CN": cn,
        "N1": N1,
        "N1_prime": N1_prime,
        "Dilatancy_Applied": "Yes" if dilatancy_applied else "No",
    }



def base_calculation(
    corrected_N, gravel, sand, silt, clay, LL, PL, moisture
):
    corrected_N = max(0.0, corrected_N)
    gravel = max(0.0, gravel)
    sand = max(0.0, sand)
    silt = max(0.0, silt)
    clay = max(0.0, clay)
    LL = max(0.0, LL)
    PL = max(0.0, PL)

    PI = max(0.0, LL - PL)
    fines = silt + clay
    CF = clay / fines if fines > 0 else 0.0
    LI = (moisture - PL) / PI if PI > 0 else 0.0
    LI = clamp(LI, -1.0, 1.5)

    soil = classify_soil(fines, sand, gravel, clay, silt)

    # -----------------------------
    # UNDRAINED BASELINE
    # -----------------------------
    K_PI = piecewise_k_pi(PI)

    # LI should have strong influence on cohesive soils,
    # moderate influence on mixed soils, and essentially no
    # influence on clean granular soils.
    li_factor_map = {
        "Clean Granular Soil": 1.00,
        "Gravelly Soil with Fines": 1.00 - 0.07 * LI,
        "Silty/Clayey Sand": 1.00 - 0.12 * LI,
        "Silty Soil": 1.00 - 0.20 * LI,
        "Clayey Soil": 1.00 - 0.25 * LI,
    }
    F_LI = clamp(li_factor_map[soil], 0.75, 1.25)

    # Clay fraction is only a secondary modifier so PI and LI
    # are not double-counted.
    if soil == "Clean Granular Soil":
        F_C = 1.00
    else:
        F_C = clamp(0.92 + 0.16 * CF, 0.88, 1.08)

    # Slight soil-behaviour modifier around the PI correlation.
    soil_factor = {
        "Clean Granular Soil": 1.00,
        "Gravelly Soil with Fines": 0.95,
        "Silty/Clayey Sand": 0.98,
        "Silty Soil": 0.95,
        "Clayey Soil": 0.92,
    }[soil]

    Su_kPa = corrected_N * K_PI * F_LI * F_C * soil_factor
    Su_kPa = clamp(Su_kPa, 0.0, 450.0)

    # -----------------------------
    # DRAINED FRICTION ANGLE
    # -----------------------------
    # corrected_N controls granular behaviour most strongly.
    # Fine-grained soils receive progressively weaker corrected_N influence.
    granular_phi = 27.0 + 5.2 * math.log10(corrected_N + 1.0)
    granular_phi = clamp(granular_phi, 27.0, 40.0)

    if soil == "Clean Granular Soil":
        phi_dr = granular_phi + 0.025 * gravel
    elif soil == "Gravelly Soil with Fines":
        phi_dr = granular_phi - 0.045 * fines + 0.015 * gravel
    elif soil == "Silty/Clayey Sand":
        phi_dr = granular_phi - 0.055 * fines - 0.02 * PI
    elif soil == "Silty Soil":
        phi_dr = 28.0 + 0.065 * corrected_N - 0.035 * fines - 0.03 * PI
    else:
        phi_dr = 21.0 + 0.055 * corrected_N - 0.025 * fines - 0.045 * PI

    phi_dr = clamp(phi_dr, 20.0, 42.0)

    # -----------------------------
    # DRAINED COHESION
    # -----------------------------
    # Avoid a strong universal c' proportional-to-corrected_N relationship.
    # Small intercepts are used, with structured fine-grained soils
    # allowed somewhat higher values.
    if soil == "Clean Granular Soil":
        c_dr_kPa = 0.0
    elif soil == "Gravelly Soil with Fines":
        c_dr_kPa = 2.0 + 0.05 * fines + 0.02 * PI
    elif soil == "Silty/Clayey Sand":
        c_dr_kPa = 2.5 + 0.06 * fines + 0.025 * PI
    elif soil == "Silty Soil":
        c_dr_kPa = 3.0 + 0.07 * fines + 0.025 * PI
    else:
        c_dr_kPa = 8.0 + 0.08 * PI + 0.025 * fines

    c_dr_kPa = clamp(c_dr_kPa, 0.0, 20.0)

    return {
        "PI": PI,
        "LI": LI,
        "CF": CF,
        "Fines": fines,
        "Soil": soil,
        "K_PI": K_PI,
        "F_LI": F_LI,
        "F_C": F_C,
        "Su_kPa": Su_kPa,
        "phi_dr": phi_dr,
        "c_dr_kPa": c_dr_kPa,
    }


def synthetic_calculation(
    corrected_N, gravel, sand, silt, clay, LL, PL, moisture,
    rng, variability_pct
):
    b = base_calculation(
        corrected_N, gravel, sand, silt, clay, LL, PL, moisture
    )

    v = clamp(variability_pct, 0.0, 30.0) / 100.0

    # Soil-dependent scatter.
    soil_spread = {
        "Clean Granular Soil": 0.70,
        "Gravelly Soil with Fines": 0.85,
        "Silty/Clayey Sand": 0.95,
        "Silty Soil": 1.05,
        "Clayey Soil": 1.15,
    }[b["Soil"]]

    strength_spread = v * soil_spread

    # Strongly correlated undrained family:
    # Su -> Cu -> UCS.
    common_undrained = rng.triangular(
        1.0 - strength_spread,
        1.0,
        1.0 + strength_spread
    )

    su_factor = common_undrained * rng.triangular(0.99, 1.0, 1.01)
    cu_factor = common_undrained * rng.triangular(0.985, 1.0, 1.015)

    Su_kPa = clamp(
        b["Su_kPa"] * su_factor,
        0.0,
        450.0
    )

    Cu_UU_kPa = clamp(
        b["Su_kPa"] * cu_factor,
        0.0,
        450.0
    )

    # UCS/Su ratio varies modestly around 2.
    if b["Soil"] in ("Clayey Soil", "Silty Soil"):
        ucs_ratio = rng.triangular(1.80, 2.00, 2.20)
    else:
        ucs_ratio = rng.triangular(1.85, 2.00, 2.15)

    UCS_kPa = clamp(Su_kPa * ucs_ratio, 0.0, 900.0)

    # -----------------------------
    # REPORTED UU PHI
    # -----------------------------
    # Important distinction:
    # For fully saturated soil, UU theory normally uses phi=0.
    # In actual multi-pressure laboratory data, however, a fitted
    # total-stress envelope can show a small apparent phi because
    # of specimen variability, partial saturation, disturbance,
    # test scatter and regression of several Mohr circles.
    #
    # Therefore this generator reports a SMALL APPARENT/FITTED UU phi,
    # rather than pretending every real report will display exactly 0.
    # User-specified ranges for the reported/fitted UU triaxial phi.
    # These represent the range to reproduce in the synthetic laboratory
    # dataset. A triangular distribution is used so values near the middle
    # of the range occur more frequently than values at the limits.
    uu_phi_ranges = {
        "Clayey Soil": (0.0, 11.0),
        "Silty Soil": (5.0, 15.0),
        "Silty/Clayey Sand": (5.0, 17.0),
        "Gravelly Soil with Fines": (12.0, 22.0),
        "Clean Granular Soil": (13.0, 24.0),
    }

    phi_uu_min, phi_uu_max = uu_phi_ranges[b["Soil"]]
    phi_uu_mode = (phi_uu_min + phi_uu_max) / 2.0

    # Add a small soil-test scatter around the central tendency while
    # strictly retaining the requested minimum and maximum limits.
    scatter = 0.10 * (phi_uu_max - phi_uu_min) * max(0.0, min(v / 0.10, 1.5))
    phi_uu_mode = clamp(
        phi_uu_mode + rng.uniform(-scatter, scatter),
        phi_uu_min,
        phi_uu_max
    )

    # Python's triangular() signature is triangular(low, high, mode).
    # Keep the requested minimum/maximum as the hard limits.
    phi_uu = rng.triangular(
        phi_uu_min,
        phi_uu_max,
        phi_uu_mode
    )

    # -----------------------------
    # DRAINED FAMILY
    # -----------------------------
    common_drained = rng.triangular(
        1.0 - 0.70 * v,
        1.0,
        1.0 + 0.70 * v
    )

    phi_noise = rng.triangular(
        -2.0 * (0.5 + 3.0 * v),
        0.0,
        2.0 * (0.5 + 3.0 * v)
    )

    c_noise_factor = rng.triangular(
        1.0 - 0.35 * v,
        1.0,
        1.0 + 0.35 * v
    )

    phi_dr = clamp(
        b["phi_dr"] + phi_noise,
        20.0,
        42.0
    )

    c_dr_kPa = clamp(
        b["c_dr_kPa"] * common_drained * c_noise_factor,
        0.0,
        20.0
    )

    return [
        b["PI"],
        b["LI"],
        b["CF"],
        b["Soil"],
        b["K_PI"],
        b["F_LI"],
        b["F_C"],
        Su_kPa / KPA_PER_KGCM2,
        UCS_kPa / KPA_PER_KGCM2,
        c_dr_kPa / KPA_PER_KGCM2,
        phi_dr,
        Cu_UU_kPa / KPA_PER_KGCM2,
        phi_uu,
    ]


def process_excel(
    input_file, output_file, sheet_name, variability_pct, seed
):
    wb = load_workbook(input_file)
    ws = wb[sheet_name]

    input_cols = {
        "Depth": 3,
        "Field_N": 4,
        "Gravel": 5,
        "Sand": 6,
        "Silt": 7,
        "Clay": 8,
        "LL": 9,
        "PL": 10,
        "Moisture": 11,
    }

    headers = [
        "Density (gm/cc)",
        "Effective OBD (kg/cm2)",
        "pa (kg/cm2)",
        "CN Raw",
        "CN (IS 2131)",
        "N1 (OB Corrected)",
        "N1' (Final Corrected N)",
        "PI",
        "LI",
        "Clay Fraction",
        "Soil Behaviour",
        "K_PI",
        "Liquidity Factor",
        "Clay Factor",
        "Su (kg/cm2)",
        "UCS (kg/cm2)",
        "Drained Cohesion c' (kg/cm2)",
        "Drained Phi' (deg)",
        "UU Cohesion Cu (kg/cm2)",
        "UU Phi (reported deg)",
    ]

    for i, header in enumerate(headers, start=12):
        ws.cell(1, i).value = header
        ws.cell(1, i).font = Font(bold=True)
        ws.cell(1, i).fill = PatternFill("solid", fgColor="D9EAF7")
        ws.cell(1, i).alignment = Alignment(horizontal="center")

    rng = random.Random(seed if seed is not None else None)
    processed = 0

    for r in range(2, ws.max_row + 1):
        raw_depth = ws.cell(r, input_cols["Depth"]).value
        raw_field_n = ws.cell(r, input_cols["Field_N"]).value

        if raw_depth is None or str(raw_depth).strip() == "":
            continue
        if raw_field_n is None or str(raw_field_n).strip() == "":
            continue

        try:
            depth_m = float(raw_depth)
            field_n = float(raw_field_n)
        except (ValueError, TypeError):
            continue

        gravel = num(ws.cell(r, input_cols["Gravel"]).value)
        sand = num(ws.cell(r, input_cols["Sand"]).value)
        silt = num(ws.cell(r, input_cols["Silt"]).value)
        clay = num(ws.cell(r, input_cols["Clay"]).value)

        # IS 2131 dilatancy correction applies when the stratum consists
        # of fine sand and silt. Use the input fractions to identify this
        # condition. Groundwater is fixed at ground level by the user.
        fines = silt + clay
        fine_sand_or_silt = (
            fines > 0 and
            clay <= silt and
            sand >= clay
        )

        spt = calculate_corrected_spt(
            field_n,
            depth_m,
            fine_sand_or_silt
        )

        result = synthetic_calculation(
            spt["N1_prime"],
            gravel,
            sand,
            silt,
            clay,
            num(ws.cell(r, input_cols["LL"]).value),
            num(ws.cell(r, input_cols["PL"]).value),
            num(ws.cell(r, input_cols["Moisture"]).value),
            rng,
            variability_pct,
        )

        # Duplicated/derived SPT calculations are placed first in the
        # output section, followed by the strength correlations.
        derived = [
            spt["Density_gmcc"],
            spt["Effective_Overburden_kgcm2"],
            spt["Pa_kgcm2"],
            spt["CN_Raw"],
            spt["CN"],
            spt["N1"],
            spt["N1_prime"],
        ]

        for i, value in enumerate(derived, start=12):
            ws.cell(r, i).value = value

        for i, value in enumerate(result, start=19):
            ws.cell(r, i).value = value

        processed += 1

    for r in range(2, ws.max_row + 1):
        for c in [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28, 29, 30, 31, 32]:
            ws.cell(r, c).number_format = "0.000"
        ws.cell(r, 26).number_format = "0.00"
        ws.cell(r, 33).number_format = "0.00"

    for c in range(12, 34):
        ws.column_dimensions[get_column_letter(c)].width = 23

    wb.save(output_file)
    return processed


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Engineering Strength Parameter Generator v8")
        self.root.geometry("960x690")
        self.root.minsize(880, 630)

        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.sheet = tk.StringVar()
        self.variability = tk.StringVar(value="10")
        self.seed = tk.StringVar(value="")

        self.build_ui()

    def build_ui(self):
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Engineering Strength Parameter Generator",
            font=("Segoe UI", 18, "bold")
        ).pack(anchor="w", pady=(0, 15))

        box = ttk.LabelFrame(frame, text="Excel File", padding=12)
        box.pack(fill="x")

        ttk.Label(box, text="Input Excel:").grid(
            row=0, column=0, sticky="w", pady=5
        )
        ttk.Entry(box, textvariable=self.input_file).grid(
            row=0, column=1, sticky="ew", padx=8
        )
        ttk.Button(box, text="Browse", command=self.browse_input).grid(
            row=0, column=2
        )

        ttk.Label(box, text="Sheet:").grid(
            row=1, column=0, sticky="w", pady=5
        )
        self.combo = ttk.Combobox(
            box, textvariable=self.sheet, state="readonly", width=35
        )
        self.combo.grid(row=1, column=1, sticky="w", padx=8)

        ttk.Label(box, text="Output Excel:").grid(
            row=2, column=0, sticky="w", pady=5
        )
        ttk.Entry(box, textvariable=self.output_file).grid(
            row=2, column=1, sticky="ew", padx=8
        )
        ttk.Button(box, text="Save As", command=self.browse_output).grid(
            row=2, column=2
        )

        box.columnconfigure(1, weight=1)

        settings = ttk.LabelFrame(frame, text="Synthetic Variability", padding=12)
        settings.pack(fill="x", pady=12)

        ttk.Label(settings, text="Variability (%):").grid(
            row=0, column=0, sticky="w", padx=5, pady=5
        )
        ttk.Entry(
            settings, textvariable=self.variability, width=12
        ).grid(row=0, column=1, sticky="w", padx=5)

        ttk.Label(
            settings,
            text="Recommended starting value: 10"
        ).grid(row=0, column=2, sticky="w", padx=10)

        ttk.Label(settings, text="Random Seed (optional):").grid(
            row=1, column=0, sticky="w", padx=5, pady=5
        )
        ttk.Entry(
            settings, textvariable=self.seed, width=12
        ).grid(row=1, column=1, sticky="w", padx=5)

        ttk.Label(
            settings,
            text="Use an integer to reproduce the same generated dataset."
        ).grid(row=1, column=2, sticky="w", padx=10)

        info = ttk.LabelFrame(frame, text="Important modelling changes in v4", padding=12)
        info.pack(fill="x")

        ttk.Label(
            info,
            text=(
                "• Field SPT N is corrected using the supplied IS 2131 overburden and dilatancy procedure.\n"
                "• N1 and final N1' are shown separately; final N1' is used for strength correlations.\n"
                "• Effective overburden is evaluated at SPT depth + 0.30 m; density is 1.85 gm/cc to 10 m and 2.00 gm/cc below 10 m; groundwater is at ground level.\n"
                "• Drained φ' is soil-specific; N60 has strongest influence for granular soils.\n"
                "• Drained c' is no longer strongly proportional to N60.\n"
                "• Su, Cu and UCS form one strongly correlated undrained family.\n"
                "• UU Phi now follows the specified reported laboratory ranges by soil behaviour.\n"
                "  For a fully saturated UU interpretation, the theoretical total-stress assumption remains φ = 0°.\n\n"
                "The generator is for synthetic/testing datasets and does not replace laboratory measurements."
            ),
            justify="left",
            wraplength=900
        ).pack(anchor="w")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=10)

        ttk.Button(
            buttons, text="Generate & Export", command=self.run
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            buttons, text="Clear", command=self.clear
        ).pack(side="left")

        status_box = ttk.LabelFrame(frame, text="Status", padding=8)
        status_box.pack(fill="both", expand=True)

        self.status = tk.Text(
            status_box, state="disabled", wrap="word"
        )
        self.status.pack(fill="both", expand=True)

        self.log("Ready. Select your Excel file. Column C = Depth; Column D = Field SPT N. The program calculates CN, N1 and N1' automatically.")

    def log(self, text):
        self.status.config(state="normal")
        self.status.insert("end", text + "\n")
        self.status.see("end")
        self.status.config(state="disabled")

    def browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*")
            ]
        )
        if not path:
            return

        self.input_file.set(path)

        try:
            wb = load_workbook(path, read_only=True)
            sheets = wb.sheetnames
            wb.close()

            self.combo["values"] = sheets
            if sheets:
                self.sheet.set(sheets[0])

            base = path.rsplit(".", 1)[0]
            self.output_file.set(
                base + "_Synthetic_Strength_Parameters_v4.xlsx"
            )

            self.log("Input loaded: " + path)
            self.log("Sheets found: " + ", ".join(sheets))

        except Exception as exc:
            messagebox.showerror("Excel Error", str(exc))

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Output Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.output_file.set(path)

    def run(self):
        if not self.input_file.get():
            messagebox.showwarning(
                "Input Required",
                "Please select an input Excel file."
            )
            return

        if not self.sheet.get():
            messagebox.showwarning(
                "Sheet Required",
                "Please select a sheet."
            )
            return

        if not self.output_file.get():
            messagebox.showwarning(
                "Output Required",
                "Please select an output Excel file."
            )
            return

        try:
            variability = float(self.variability.get())
        except ValueError:
            messagebox.showwarning(
                "Invalid Variability",
                "Variability must be a number, for example 10."
            )
            return

        if not 0 <= variability <= 30:
            messagebox.showwarning(
                "Invalid Variability",
                "Please enter a value between 0 and 30%."
            )
            return

        seed_text = self.seed.get().strip()
        if seed_text:
            try:
                seed = int(seed_text)
            except ValueError:
                messagebox.showwarning(
                    "Invalid Seed",
                    "Random Seed must be an integer or left blank."
                )
                return
        else:
            seed = None

        try:
            self.log("Generating v4 synthetic strength parameters...")
            self.log(f"Variability: {variability:.1f}%")
            self.log(
                "Random seed: automatic"
                if seed is None
                else f"Random seed: {seed}"
            )

            count = process_excel(
                self.input_file.get(),
                self.output_file.get(),
                self.sheet.get(),
                variability,
                seed
            )

            self.log(f"Rows processed: {count}")
            self.log("Output saved: " + self.output_file.get())

            messagebox.showinfo(
                "Completed",
                f"v4 synthetic generation completed.\n\n"
                f"Rows processed: {count}"
            )

        except PermissionError:
            messagebox.showerror(
                "File is Open",
                "Please close the output Excel file and try again."
            )

        except Exception as exc:
            self.log("ERROR: " + str(exc))
            messagebox.showerror("Generation Error", str(exc))

    def clear(self):
        self.input_file.set("")
        self.output_file.set("")
        self.sheet.set("")
        self.variability.set("10")
        self.seed.set("")
        self.combo["values"] = ()

        self.status.config(state="normal")
        self.status.delete("1.0", "end")
        self.status.config(state="disabled")

        self.log("Ready. Select your Excel file. Column C = Depth; Column D = Field SPT N. The program calculates CN, N1 and N1' automatically.")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
