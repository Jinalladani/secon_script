
import math
import os
import random
import csv
from pathlib import Path
import tkinter as tk

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import Rectangle
import numpy as np
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook


class DecimalLogFormatter(mticker.Formatter):
    """Format major ticks of a log x-axis as ordinary decimal numbers."""
    def __call__(self, x, pos=None):
        if x <= 0:
            return ""
        exponent = round(math.log10(x))
        if abs(x - (10 ** exponent)) > max(1e-12, abs(x) * 1e-10):
            return ""
        if exponent >= 0:
            return f"{10 ** exponent:.0f}"
        return f"{10 ** exponent:.{-exponent}f}"

# ================================================================
# HYDROMETER SYNTHETIC DATA GENERATOR v1.8
# ================================================================

OBS_TIMES = [0.5, 1.0, 2.0, 4.0, 8.0, 15.0, 30.0, 60.0,
             120.0, 240.0, 480.0, 960.0, 1440.0]
SPECIAL_TIMES = {0.5, 1.0, 2.0, 4.0}

# hydrometer -> cylinder -> (slope, special intercept, other intercept)
CALIBRATION = {
    1: {1: (-312.0, 329.33, 330.51),
        2: (-312.0, 329.16, 330.51)},
    2: {1: (-307.4, 325.68, 326.82),
        2: (-307.4, 325.52, 326.82),
        3: (-307.4, 325.69, 326.82)},
    3: {1: (-248.2, 263.81, 264.75),
        2: (-248.2, 263.67, 264.75),
        3: (-248.2, 263.82, 264.75)},
    4: {1: (-223.4, 237.72, 238.69),
        2: (-223.4, 237.58, 238.69),
        3: (-223.4, 237.73, 238.69)},
}

# The first value controls curvature. The other values are reserved
# for future preset refinement and are not used as hidden calibration.
PRESETS = {
    "1 - Coarse Silt Dominant": 0.80,
    "2 - Normal Silt": 1.15,
    "3 - Fine Silt Dominant": 1.55,
    "4 - Clayey Silt": 2.00,
    "5 - Silty Clay": 2.80,
    "6 - Broad Natural": 1.30,
}


def to_float(value, label, row=None):
    if value is None or str(value).strip() == "":
        suffix = f" at GSD row {row}" if row else ""
        raise ValueError(f"{label}{suffix} is blank.")
    try:
        return float(value)
    except Exception:
        suffix = f" at GSD row {row}" if row else ""
        raise ValueError(f"{label}{suffix} is not numeric: {value!r}")


def viscosity_from_temperature(T):
    return (
        0.000000002615 * T**4
        - 0.00000030607 * T**3
        + 0.000016427499 * T**2
        - 0.000618546309 * T
        + 0.017919182437
    )


def get_calibration(hydrometer, cylinder, t):
    try:
        slope, special_i, other_i = CALIBRATION[int(hydrometer)][int(cylinder)]
    except KeyError:
        raise ValueError(
            f"Hydrometer {hydrometer} with Cylinder {cylinder} has no supplied calibration."
        )
    intercept = special_i if t in SPECIAL_TIMES else other_i
    return slope, intercept


def effective_depth(reading, cm, hydrometer, cylinder, t):
    slope, intercept = get_calibration(hydrometer, cylinder, t)
    return slope * (reading + cm) + intercept


def particle_diameter(reading, cm, hydrometer, cylinder, t, mu, gs):
    hr = effective_depth(reading, cm, hydrometer, cylinder, t)
    if hr <= 0 or gs <= 1 or t <= 0:
        return None
    return math.sqrt((30.0 * mu * hr) / (980.0 * (gs - 1.0) * t))


def inverse_reading(percent_finer, wb, gs, cm, mt, x):
    """
    Exact inverse of the supplied % finer equation:

    Rh = (Reading - 1)*1000 + Cm
    W  = (100*Gs)/(Wb*(Gs-1)) * (Rh + Mt - X)
    """
    rh = percent_finer * wb * (gs - 1.0) / (100.0 * gs) - mt + x
    return 1.0 + (rh - cm) / 1000.0


def build_psd_function(fines, clay, preset):
    exponent = PRESETS[preset]
    if clay < 0 or clay > fines:
        raise ValueError("Target Clay % must be between 0 and GSD Fines %.")

    def percent_finer(d):
        if d >= 0.075:
            return fines
        if d <= 0.002:
            return clay

        # u=0 at 0.075 mm and u=1 at 0.002 mm.
        u = (math.log10(d) - math.log10(0.075)) / (
            math.log10(0.002) - math.log10(0.075)
        )
        u = max(0.0, min(1.0, u))
        return fines - (fines - clay) * (u ** exponent)

    return percent_finer


def solve_reading(t, pfun, wb, gs, cm, mt, x, hydro, cylinder, mu, rng):
    """
    Solve the self-consistent system:
        P = PSD(D)
        Reading = inverse_percent_finer(P)
        D = Stokes(Reading, t)

    The final instrument value is rounded to 0.001.
    """
    # Search the actual instrument resolution, not an ideal continuous
    # reading. This ensures the output reflects the 0.001 hydrometer scale.
    candidates = [round(0.995 + 0.001 * i, 3) for i in range(36)]
    scored = []

    for reading in candidates:
        d = particle_diameter(reading, cm, hydro, cylinder, t, mu, gs)
        if d is None or not (0.0005 <= d <= 0.10):
            continue

        p = pfun(d)
        ideal_reading = inverse_reading(p, wb, gs, cm, mt, x)

        # Distance from the self-consistent inverse solution.
        score = abs(reading - ideal_reading)

        # Very small controlled perturbation avoids every generated
        # dataset having identical tie-breaking.
        score += rng.uniform(0.0, 0.00015)

        scored.append((score, reading, p, d, ideal_reading))

    if not scored:
        raise ValueError(
            f"No valid hydrometer reading in 0.995–1.030 at {t} min. "
            "Check Temperature, SPGR, corrections and target fractions."
        )

    return min(scored, key=lambda z: z[0])


def generate_sample(row, wb_weight, tolerance, rng):
    # Fixed laboratory settings requested by user
    T = 27.0
    cm = 0.0005
    mt = 0.0
    x = 0.3
    gsd_row = row["gsd_row"]

    fines = to_float(row["fines"], "GSD Fines %", gsd_row)
    silt = to_float(row["silt"].get(), "Target Silt %", gsd_row)

    # Clay is derived automatically from total GSD fines:
    # Clay % = Fines % - Silt %
    clay = fines - silt

    if not 0 <= fines <= 100:
        raise ValueError(f"GSD row {gsd_row}: Fines must be 0–100%.")
    if silt < 0:
        raise ValueError(f"GSD row {gsd_row}: Target Silt % cannot be negative.")

    if silt > fines + tolerance:
        raise ValueError(
            f"GSD row {gsd_row}: Target Silt = {silt:.2f}% exceeds "
            f"GSD Fines = {fines:.2f}%."
        )

    if clay < 0:
        raise ValueError(
            f"GSD row {gsd_row}: Calculated Clay = Fines - Silt = "
            f"{fines:.2f} - {silt:.2f} = {clay:.2f}%, which is negative."
        )

    gs = to_float(row["spgr"].get(), "SPGR", gsd_row)
    hydro = int(round(to_float(row["hydrometer"].get(), "Hydrometer No.", gsd_row)))
    cylinder = int(round(to_float(row["cylinder"].get(), "Cylinder Type", gsd_row)))

    if gs <= 1:
        raise ValueError(f"GSD row {gsd_row}: SPGR must be > 1.")
    if hydro not in CALIBRATION or cylinder not in CALIBRATION[hydro]:
        raise ValueError(
            f"GSD row {gsd_row}: Hydrometer {hydro} + Cylinder {cylinder} "
            "has no supplied calibration."
        )

    mu = viscosity_from_temperature(T)
    pfun = build_psd_function(fines, clay, row["preset"].get())

    readings = []
    calculated_points = []

    for t in OBS_TIMES:
        score, reading, p, d, ideal_reading = solve_reading(
            t, pfun, wb_weight, gs, cm, mt, x,
            hydro, cylinder, mu, rng
        )

        # Reading is already at 0.001 resolution.
        readings.append(reading)
        calculated_points.append((t, reading, d, p, ideal_reading))

    # Hydrometer response should normally decrease with time.
    # Repeated values are allowed because the instrument resolution is 0.001.
    # We only prevent an accidental upward jump larger than one least count.
    for i in range(1, len(readings)):
        if readings[i] > readings[i - 1] + 0.001:
            readings[i] = readings[i - 1]

    return {
        "temperature": T,
        "mu": mu,
        "gs": gs,
        "cm": cm,
        "hydrometer": hydro,
        "cylinder": cylinder,
        "mt": mt,
        "x": x,
        "readings": readings,
        "fines": fines,
        "silt": silt,
        "clay": clay,
        "preset": row["preset"].get(),
        "points": calculated_points,
    }


def _safe_float(value):
    try:
        if value is None or str(value).strip() == "":
            return None
        return float(value)
    except Exception:
        return None


def _interp_log_d(particle_sizes, percent_finer, target_percent):
    """Interpolate particle size at a target % finer on a log particle-size axis."""
    pts = sorted((d, p) for d, p in zip(particle_sizes, percent_finer)
                 if d is not None and p is not None and d > 0)
    if len(pts) < 2:
        return None
    # Remove duplicate particle sizes.
    clean = []
    for d, p in pts:
        if not clean or abs(d - clean[-1][0]) > 1e-12:
            clean.append((d, p))
        else:
            clean[-1] = (d, p)
    # Interpolate P as a function of log10(D).
    clean.sort(key=lambda z: z[1])
    ps = [z[1] for z in clean]
    ds = [z[0] for z in clean]
    if target_percent < min(ps) or target_percent > max(ps):
        return None
    for i in range(len(ps) - 1):
        p1, p2 = ps[i], ps[i + 1]
        if p1 <= target_percent <= p2:
            if abs(p2 - p1) < 1e-12:
                return 10 ** ((np.log10(ds[i]) + np.log10(ds[i + 1])) / 2.0)
            f = (target_percent - p1) / (p2 - p1)
            logd = np.log10(ds[i]) + f * (np.log10(ds[i + 1]) - np.log10(ds[i]))
            return float(10 ** logd)
    return None


# ---------------------------------------------------------------------
# REPORT-STYLE PSD PLOT DRAWING (matches PSD_plot_generator_v2_9)
# ---------------------------------------------------------------------
FIG_W = 16.0
FIG_H = 9.4
DPI = 180
X_MIN = 1e-4
X_MAX = 100.0
Y_MIN = 0.0
Y_MAX = 100.0
LEFT = 0.07
RIGHT = 0.985
GRAPH_BOTTOM = 0.205
GRAPH_TOP = 0.765
INFO_BOTTOM = 0.845
INFO_TOP = 0.915
TITLE_BOTTOM = 0.915
TABLE_BOTTOM = 0.045
TABLE_TOP = 0.135


def _plot_clean_text(value):
    return "" if value is None else str(value).strip()


def draw_header(fig, bh_id, depth):
    fig.text(0.5, 0.958, "PARTICLE SIZE DISTRIBUTION CURVE",
             ha="center", va="center", fontsize=21, fontweight="bold")
    fig.text(0.5, 0.928, "(Sieve Analysis + Hydrometer Analysis)",
             ha="center", va="center", fontsize=12)

    fig.patches.append(Rectangle(
        (0.008, 0.008), 0.984, 0.984,
        transform=fig.transFigure, fill=False, linewidth=1.3
    ))

    fig.lines.append(plt.Line2D(
        [0.008, 0.992], [TITLE_BOTTOM, TITLE_BOTTOM],
        transform=fig.transFigure, linewidth=1.0
    ))
    fig.lines.append(plt.Line2D(
        [0.008, 0.992], [INFO_BOTTOM, INFO_BOTTOM],
        transform=fig.transFigure, linewidth=1.0
    ))

    # Same two-field information row as PSD generator: BH ID | Depth.
    x_mid = 0.675
    fig.lines.append(plt.Line2D(
        [x_mid, x_mid], [INFO_BOTTOM, INFO_TOP],
        transform=fig.transFigure, linewidth=1.0
    ))

    fields = [
        (0.030, "BH ID", bh_id, 0.080, 0.115),
        (0.695, "Depth", _plot_clean_text(depth), 0.080, 0.115),
    ]
    for x, label, value, colon_offset, value_offset in fields:
        fig.text(x, 0.878, label, ha="left", va="center",
                 fontsize=11.5, fontweight="bold")
        fig.text(x + colon_offset, 0.878, ":", ha="center", va="center",
                 fontsize=11.5, fontweight="bold")
        fig.text(x + value_offset, 0.878, _plot_clean_text(value),
                 ha="left", va="center", fontsize=11.5)


def draw_classification_band(fig):
    def xfig(mm):
        lx = math.log10(mm)
        lo = math.log10(X_MIN)
        hi = math.log10(X_MAX)
        frac = (hi - lx) / (hi - lo)
        return LEFT + frac * (RIGHT - LEFT)

    band_bottom = GRAPH_TOP + 0.004
    band_top = INFO_BOTTOM - 0.008
    band_h = band_top - band_bottom
    ymid = band_bottom + 0.52 * band_h

    fig.patches.append(Rectangle(
        (LEFT, band_bottom), RIGHT - LEFT, band_h,
        transform=fig.transFigure, facecolor="white",
        edgecolor="black", linewidth=0.8, zorder=8
    ))

    major = [
        ("GRAVEL", 75.0, 4.75),
        ("SAND", 4.75, 0.075),
        ("SILT", 0.075, 0.002),
        ("CLAY", 0.002, X_MIN),
    ]
    for name, high, low in major:
        xa = max(LEFT, min(RIGHT, xfig(high)))
        xb = max(LEFT, min(RIGHT, xfig(low)))
        if xb < xa:
            xa, xb = xb, xa
        fig.patches.append(Rectangle(
            (xa, band_bottom), xb - xa, band_h,
            transform=fig.transFigure, fill=False,
            linewidth=0.75, zorder=9
        ))
        fig.text((xa + xb) / 2, band_bottom + 0.72 * band_h,
                 name, ha="center", va="center", fontsize=9.2,
                 fontweight="bold", zorder=10)

    fig.lines.append(plt.Line2D(
        [LEFT, RIGHT], [ymid, ymid], transform=fig.transFigure,
        linewidth=0.7, zorder=9
    ))

    for mm in (20.0, 2.0, 0.425):
        xx = xfig(mm)
        if LEFT <= xx <= RIGHT:
            fig.lines.append(plt.Line2D(
                [xx, xx], [band_bottom, ymid], transform=fig.transFigure,
                linewidth=0.7, zorder=9
            ))

    lower = [
        ("Coarse", 75.0, 20.0),
        ("Fine", 20.0, 4.75),
        ("Coarse", 4.75, 2.0),
        ("Medium", 2.0, 0.425),
        ("Fine", 0.425, 0.075),
    ]
    for label, high, low in lower:
        xa = max(LEFT, xfig(high))
        xb = min(RIGHT, xfig(low))
        if xb > xa:
            fig.text((xa + xb) / 2, band_bottom + 0.20 * band_h,
                     label, ha="center", va="center", fontsize=8.2,
                     fontweight="bold", zorder=10)


def draw_graph(fig, sieve, hydro):
    ax = fig.add_axes([
        LEFT, GRAPH_BOTTOM, RIGHT - LEFT, GRAPH_TOP - GRAPH_BOTTOM
    ])
    ax.set_xscale("log")
    ax.set_xlim(X_MAX, X_MIN)
    ax.set_ylim(Y_MIN, Y_MAX)

    ax.xaxis.set_major_locator(
        mticker.LogLocator(base=10.0, subs=(1.0,), numticks=20)
    )
    ax.xaxis.set_major_formatter(DecimalLogFormatter())
    ax.xaxis.set_minor_locator(
        mticker.LogLocator(base=10.0, subs=tuple(range(2, 10)), numticks=100)
    )
    ax.xaxis.set_minor_formatter(mticker.NullFormatter())

    ax.set_xlabel("PARTICLE SIZE / GRAIN SIZE (mm)",
                  fontsize=11.5, fontweight="bold", labelpad=8)
    ax.set_ylabel("PERCENTAGE FINER (%)",
                  fontsize=11.5, fontweight="bold", labelpad=8)
    ax.set_yticks(range(0, 101, 10))
    ax.grid(True, which="major", linewidth=0.75, alpha=0.42)
    ax.grid(True, which="minor", linewidth=0.35, alpha=0.18)

    if sieve:
        sx, sy = zip(*sieve)
        ax.plot(sx, sy, color="blue", marker="o", markersize=4.8,
                linewidth=1.2, label="Sieve Analysis", zorder=3)

    if hydro:
        hx, hy = zip(*hydro)
        ax.plot(hx, hy, color="red", marker="o", markersize=4.8,
                linewidth=1.2, label="Hydrometer Analysis", zorder=3)

    # Connect the true transition: smallest sieve size -> largest HYD size.
    if sieve and hydro:
        gsd_end = min(sieve, key=lambda p: p[0])
        hyd_start = max(hydro, key=lambda p: p[0])
        if abs(gsd_end[0] - hyd_start[0]) > 1e-10:
            ax.plot([gsd_end[0], hyd_start[0]],
                    [gsd_end[1], hyd_start[1]],
                    color="red", linewidth=1.2, zorder=2)

    handles, labels = ax.get_legend_handles_labels()
    if handles:
        ax.legend(loc="lower left", fontsize=8.5, frameon=True)
    return ax


def draw_percentage_table(fig, percentages, hyd_present=False):
    """Draw the percentage table using the same convention as the PSD plot generator.

    HYD present: Gravel | Sand | Silt | Clay
    GSD only:    Gravel | Sand | Fines
    """
    x0, x1 = 0.305, 0.715
    y0, y1 = TABLE_BOTTOM, TABLE_TOP

    fig.patches.append(Rectangle(
        (x0, y0), x1 - x0, y1 - y0,
        transform=fig.transFigure, fill=False, linewidth=0.9
    ))
    header_y = y1 - 0.035
    fig.lines.append(plt.Line2D(
        [x0, x1], [header_y, header_y], transform=fig.transFigure,
        linewidth=0.8
    ))
    fig.text((x0 + x1) / 2, y1 - 0.017, "PERCENTAGE OF PARTICLES",
             ha="center", va="center", fontsize=11, fontweight="bold")

    if hyd_present:
        labels = ["Gravel (%)", "Sand (%)", "Silt (%)", "Clay (%)"]
        vals = percentages[:4]
    else:
        labels = ["Gravel (%)", "Sand (%)", "Fines (%)"]
        vals = [percentages[0], percentages[1], percentages[2]]

    col_w = (x1 - x0) / len(labels)

    for i in range(1, len(labels)):
        xx = x0 + i * col_w
        fig.lines.append(plt.Line2D(
            [xx, xx], [y0, header_y], transform=fig.transFigure,
            linewidth=0.7
        ))

    for i, (label, value) in enumerate(zip(labels, vals)):
        xc = x0 + (i + 0.5) * col_w
        fig.text(xc, header_y - 0.020, label,
                 ha="center", va="center", fontsize=9.2, fontweight="bold")
        display_value = "—" if value is None else f"{value:.2f}"
        fig.text(xc, y0 + 0.018, display_value,
                 ha="center", va="center", fontsize=10)


def generate_psd_plot(ws_gsd, gsd_row, sample_type, depth, bh_id,
                      template_type, result, output_dir):
    """Create an individual report-style PSD plot using the PSD v2.9 design.

    Data generation remains unchanged: GSD observations are read from the
    GSD sheet and HYD observations come from the synthetic readings.
    Only the report plot layout/styling is updated to match the PSD generator.
    """
    if template_type == "Type 1":
        sieve_cols = list(range(19, 27))  # S:Z
        sieve_sizes = [4.75, 2.00, 1.18, 0.600, 0.425, 0.300, 0.150, 0.075]
    else:
        sieve_cols = list(range(17, 23))  # Q:V
        sieve_sizes = [4.75, 2.00, 0.425, 0.212, 0.150, 0.075]

    sieve = []
    for col, size in zip(sieve_cols, sieve_sizes):
        passing = _safe_float(ws_gsd.cell(gsd_row, col).value)
        if passing is not None:
            sieve.append((float(size), max(0.0, min(100.0, passing))))
    sieve.sort(key=lambda z: z[0], reverse=True)

    hydro = []
    # result is None for GSD-only samples. In that case, generate the
    # PSD plot from the actual GSD sieve observations only.
    if result is not None:
        pfun = build_psd_function(
            result["fines"], result["clay"],
            result.get("preset", "2 - Normal Silt")
        )
        for t, reading in zip(OBS_TIMES, result["readings"]):
            d = particle_diameter(
                reading, result["cm"], result["hydrometer"], result["cylinder"],
                t, result["mu"], result["gs"]
            )
            if d is not None and 0.0001 <= d <= 0.10:
                hydro.append((float(d), max(0.0, min(100.0, float(pfun(d))))))
        hydro.sort(key=lambda z: z[0], reverse=True)

    # Match PSD generator's percentage table calculation from the curve.
    combined = list(sieve) + list(hydro)
    combined.sort(key=lambda z: z[0], reverse=True)

    def interp(points, target):
        pts = sorted(points, key=lambda p: p[0], reverse=True)
        for x, y in pts:
            if abs(x - target) <= 1e-12:
                return y
        for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
            if x1 >= target >= x2 and x1 != x2:
                lx1, lx2, lt = math.log10(x1), math.log10(x2), math.log10(target)
                frac = (lt - lx1) / (lx2 - lx1)
                return y1 + frac * (y2 - y1)
        return None

    p475 = interp(combined, 4.75)
    p0075 = interp(combined, 0.075)
    p0002 = interp(combined, 0.002)

    hyd_present = bool(hydro)

    if p475 is None:
        percentages = [None, None, None, None]
    elif not hyd_present:
        # GSD-only plot: the material passing 0.075 mm is reported as FINES.
        # Do not split the fines into silt/clay because no hydrometer result exists.
        if p0075 is None:
            percentages = [max(0.0, 100.0 - p475), None, None, None]
        else:
            percentages = [
                max(0.0, 100.0 - p475),
                max(0.0, p475 - p0075),
                max(0.0, p0075),
                None,
            ]
    elif p0075 is None:
        percentages = [max(0.0, 100.0 - p475), None, None, None]
    elif p0002 is None:
        percentages = [
            max(0.0, 100.0 - p475),
            max(0.0, p475 - p0075), None, None
        ]
    else:
        percentages = [
            max(0.0, 100.0 - p475),
            max(0.0, p475 - p0075),
            max(0.0, p0075 - p0002),
            max(0.0, p0002),
        ]
        total = sum(percentages)
        if total > 0:
            percentages = [v * 100.0 / total for v in percentages]

    fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=DPI, facecolor="white")
    draw_header(fig, bh_id, depth)
    draw_classification_band(fig)
    draw_graph(fig, sieve, hydro)
    draw_percentage_table(fig, percentages, hyd_present=hyd_present)

    safe_bh = str(bh_id).replace("/", "_").replace("\\", "_")
    safe_sample = str(sample_type).replace("/", "_").replace("\\", "_")
    safe_depth = str(depth).replace("/", "_").replace("\\", "_")
    output_path = Path(output_dir) / f"PSD_{safe_bh}_{safe_sample}_{safe_depth}.png"

    fig.savefig(output_path, dpi=DPI, facecolor="white", edgecolor="white",
                bbox_inches=None, pad_inches=0)
    plt.close(fig)
    # Return both the plot path and the exact four percentage values
    # displayed in the plot table, so PSD_Summary.csv uses identical values.
    return output_path, percentages


def export_psd_summary(summary_rows, output_dir):
    """Export the same PSD_Summary.csv structure used by the PSD plot generator."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    summary_path = output_dir / "PSD_Summary.csv"

    headers = [
        "BH ID",
        "Depth",
        "Sample Type",
        "Sieve Set",
        "% Gravel",
        "% Sand",
        "% Silt",
        "% Clay",
    ]

    def fmt(value):
        if value is None or str(value).strip() == "":
            return ""
        try:
            return f"{float(value):.2f}"
        except (TypeError, ValueError):
            return str(value)

    with open(summary_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for item in summary_rows:
            writer.writerow([
                item.get("bh_id", ""),
                item.get("depth", ""),
                item.get("sample_type", ""),
                item.get("sieve_set", ""),
                fmt(item.get("gravel")),
                fmt(item.get("sand")),
                fmt(item.get("silt")),
                fmt(item.get("clay")),
            ])

    return summary_path


class HydrometerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hydrometer Synthetic Data Generator v1.9")
        self.root.geometry("1750x850")
        self.root.minsize(1250, 650)

        self.file_path = None
        self.rows = []
        self.ws_gsd_values = None
        self.bh_id = ""

        self.template_type = tk.StringVar(value="Type 1")
        self.wb_var = tk.StringVar(value="50.0")
        self.tolerance_var = tk.StringVar(value="0.75")
        self.seed_var = tk.StringVar(value="20260819")

        self.build_ui()

    def build_ui(self):
        general = ttk.LabelFrame(self.root, text="General Settings", padding=8)
        general.pack(fill="x", padx=10, pady=8)

        ttk.Label(general, text="Template Type:").grid(row=0, column=0, padx=5)
        self.type_box = ttk.Combobox(
            general, textvariable=self.template_type,
            values=["Type 1", "Type 2"], state="readonly", width=10
        )
        self.type_box.grid(row=0, column=1, padx=5)
        self.type_box.bind("<<ComboboxSelected>>", lambda e: self.reload_rows())

        ttk.Label(general, text="Wb (g):").grid(row=0, column=2, padx=5)
        ttk.Entry(general, textvariable=self.wb_var, width=9).grid(row=0, column=3)

        ttk.Label(general, text="Silt/Clay allowed error ±%:").grid(row=0, column=4, padx=5)
        ttk.Entry(general, textvariable=self.tolerance_var, width=9).grid(row=0, column=5)

        ttk.Label(general, text="Random seed:").grid(row=0, column=6, padx=5)
        ttk.Entry(general, textvariable=self.seed_var, width=12).grid(row=0, column=7)

        ttk.Button(
            general, text="Open GSD & HYD Template",
            command=self.open_template
        ).grid(row=0, column=8, padx=15)

        self.file_label = ttk.Label(general, text="No template selected")
        self.file_label.grid(row=0, column=9, sticky="w")

        note = ttk.Label(
            self.root,
            text=(
                "Type 1: GSD AA/AB/AC = Gravel/Sand/Fines    |    "
                "Type 2: GSD W/X/Y = Gravel/Sand/Fines    |    "
                "Fractions displayed to 2 decimal places. Fixed HYD settings: T=27°C, Cm=0.0005, Mt=0, X=0.3."
            ),
            padding=(12, 0)
        )
        note.pack(anchor="w")

        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True, padx=10, pady=6)

        self.canvas = tk.Canvas(outer, highlightthickness=0)
        self.vbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.hbar = ttk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)
        self.table = ttk.Frame(self.canvas)
        self.window = self.canvas.create_window((0, 0), window=self.table, anchor="nw")

        self.canvas.configure(
            yscrollcommand=self.vbar.set,
            xscrollcommand=self.hbar.set
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        self.table.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.headers = [
            "Depth (m)", "Sample Type", "% Gravel", "% Sand", "% Fines",
            "Target Silt %", "Target Clay %", "Soil Behaviour",
            "Generate HYD?", "SPGR", "Hydrometer No.", "Cylinder Type"
        ]
        self.build_headers()

        bottom = ttk.Frame(self.root, padding=8)
        bottom.pack(fill="x")
        self.status = ttk.Label(bottom, text="Open a GSD & HYD template. Fixed settings: T=27°C, Cm=0.0005, Mt=0, X=0.3.")
        self.status.pack(side="left")
        ttk.Button(
            bottom, text="Generate HYD Data",
            command=self.generate
        ).pack(side="right")

    def build_headers(self):
        for c, text in enumerate(self.headers):
            ttk.Label(
                self.table, text=text, relief="ridge",
                anchor="center", justify="center", padding=5
            ).grid(row=0, column=c, sticky="nsew")

    def clear_table(self):
        for child in self.table.winfo_children():
            child.destroy()
        self.rows = []
        self.build_headers()

    @staticmethod
    def display_fraction(v):
        if v is None or str(v).strip() == "":
            return ""
        try:
            return f"{float(v):.2f}"
        except Exception:
            return str(v)

    def open_template(self):
        path = filedialog.askopenfilename(
            title="Select GSD & HYD Template",
            filetypes=[
                ("Excel Macro-Enabled Workbook", "*.xlsm"),
                ("Excel Workbook", "*.xlsx")
            ]
        )
        if not path:
            return

        try:
            wb = load_workbook(path, data_only=True)
            if "GSD" not in wb.sheetnames:
                raise ValueError("GSD sheet not found.")
            if "HYD" not in wb.sheetnames:
                raise ValueError("HYD sheet not found.")

            self.file_path = path
            self.ws_gsd_values = wb["GSD"]
            self.bh_id = self.ws_gsd_values["B5"].value
            self.file_label.config(text=os.path.basename(path))
            self.reload_rows()
            wb.close()

        except Exception as e:
            messagebox.showerror("Template Error", str(e))

    def reload_rows(self):
        if self.ws_gsd_values is None:
            return

        self.clear_table()

        if self.template_type.get() == "Type 1":
            gravel_col, sand_col, fines_col = 27, 28, 29  # AA AB AC
        else:
            gravel_col, sand_col, fines_col = 23, 24, 25  # W X Y

        for r in range(5, self.ws_gsd_values.max_row + 1):
            depth = self.ws_gsd_values.cell(r, 6).value       # F
            sample_type = self.ws_gsd_values.cell(r, 5).value # E
            gravel = self.ws_gsd_values.cell(r, gravel_col).value
            sand = self.ws_gsd_values.cell(r, sand_col).value
            fines = self.ws_gsd_values.cell(r, fines_col).value

            if all(v is None or str(v).strip() == ""
                   for v in (depth, sample_type, gravel, sand, fines)):
                break

            self.add_row(r, depth, sample_type, gravel, sand, fines)

        self.status.config(text=f"{len(self.rows)} sample row(s) loaded.")

    def add_row(self, gsd_row, depth, sample_type, gravel, sand, fines):
        r = len(self.rows) + 1

        silt = tk.StringVar()
        clay = tk.StringVar()
        preset = tk.StringVar(value="2 - Normal Silt")
        generate = tk.BooleanVar(value=False)

        spgr = tk.StringVar()
        hydrometer = tk.StringVar()
        cylinder = tk.StringVar()

        display = [
            depth, sample_type,
            self.display_fraction(gravel),
            self.display_fraction(sand),
            self.display_fraction(fines)
        ]

        for c, value in enumerate(display):
            ttk.Label(
                self.table, text=str(value), relief="ridge",
                anchor="center", padding=4
            ).grid(row=r, column=c, sticky="nsew")

        silt_entry = ttk.Entry(self.table, textvariable=silt, width=12)
        silt_entry.grid(row=r, column=5, sticky="ew", padx=2)

        # Clay is automatically calculated as:
        # Clay % = Total Fines % - Silt %
        clay_entry = ttk.Entry(
            self.table, textvariable=clay, width=12, state="readonly"
        )
        clay_entry.grid(row=r, column=6, sticky="ew", padx=2)

        def update_clay(*_):
            try:
                fines_value = float(fines)
                silt_text = silt.get().strip()
                if not silt_text:
                    clay.set("")
                    return

                silt_value = float(silt_text)
                clay_value = fines_value - silt_value

                # Do not allow an invalid negative clay value.
                if clay_value < 0:
                    clay.set("")
                else:
                    clay.set(f"{clay_value:.2f}")
            except (ValueError, TypeError):
                clay.set("")

        silt.trace_add("write", update_clay)

        ttk.Combobox(
            self.table, textvariable=preset,
            values=list(PRESETS.keys()), state="readonly", width=24
        ).grid(row=r, column=7, sticky="ew", padx=2)

        ttk.Checkbutton(
            self.table, variable=generate,
            command=lambda idx=len(self.rows): self.toggle_hyd(idx)
        ).grid(row=r, column=8)

        entries = []
        for c, var in enumerate(
            [spgr, hydrometer, cylinder], start=9
        ):
            e = ttk.Entry(
                self.table, textvariable=var, width=15, state="disabled"
            )
            e.grid(row=r, column=c, sticky="ew", padx=2)
            entries.append(e)

        self.rows.append({
            "gsd_row": gsd_row,
            "depth": depth,
            "sample_type": sample_type,
            "gravel": gravel,
            "sand": sand,
            "fines": fines,
            "silt": silt,
            "clay": clay,
            "preset": preset,
            "generate": generate,
            "spgr": spgr,
            "hydrometer": hydrometer,
            "cylinder": cylinder,
            "entries": entries,
        })

    def toggle_hyd(self, idx):
        row = self.rows[idx]
        state = "normal" if row["generate"].get() else "disabled"
        for e in row["entries"]:
            e.configure(state=state)

    def generate(self):
        if not self.file_path:
            messagebox.showerror("Generation Error", "Open a GSD & HYD template first.")
            return

        # 'Generate HYD?' controls only synthetic HYD generation.
        # PSD/GSD plots are generated for EVERY loaded GSD sample, whether
        # HYD is selected or not.
        selected = [r for r in self.rows if r["generate"].get()]
        if not self.rows:
            messagebox.showinfo(
                "Generation",
                "No GSD sample rows are loaded."
            )
            return

        try:
            wb_weight = to_float(self.wb_var.get(), "Wb")
            tolerance = to_float(self.tolerance_var.get(), "Silt/Clay tolerance")
            seed = int(float(self.seed_var.get()))
            rng = random.Random(seed)

            # Preserve VBA project when input is XLSM.
            keep_vba = self.file_path.lower().endswith(".xlsm")
            wb = load_workbook(self.file_path, keep_vba=keep_vba)
            hyd = wb["HYD"]

            base, ext = os.path.splitext(self.file_path)
            output = base + "_HYD_Generated" + ext
            plot_dir = os.path.join(os.path.dirname(output), "PSD_Plots")
            os.makedirs(plot_dir, exist_ok=True)

            # Generate HYD only for rows where 'Generate HYD?' is checked,
            # but generate a PSD plot for ALL GSD rows. For non-HYD rows the
            # plot contains the GSD/sieve curve only.
            selected_by_row = {row["gsd_row"]: row for row in selected}
            summary_rows = []

            for row in self.rows:
                hr = row["gsd_row"]
                result = None

                if hr in selected_by_row:
                    result = generate_sample(
                        row, wb_weight, tolerance, rng
                    )

                    # HYD E:L
                    # Fixed laboratory settings:
                    # E = Temperature = 27°C
                    # F = Coefficient of viscosity calculated from E
                    # G = Specific Gravity
                    # H = Meniscus correction = 0.0005
                    # I = Hydrometer No.
                    # J = Cylinder Type
                    # K = Temperature correction (Mt) = 0
                    # L = Dispersion correction (X) = 0.3
                    hyd.cell(hr, 5).value = 27.0
                    hyd.cell(hr, 6).value = result["mu"]
                    hyd.cell(hr, 7).value = result["gs"]
                    hyd.cell(hr, 8).value = 0.0005
                    hyd.cell(hr, 9).value = result["hydrometer"]
                    hyd.cell(hr, 10).value = result["cylinder"]
                    hyd.cell(hr, 11).value = 0.0
                    hyd.cell(hr, 12).value = 0.3

                    # HYD M:Y
                    for i, reading in enumerate(result["readings"], start=13):
                        hyd.cell(hr, i).value = reading
                        hyd.cell(hr, i).number_format = "0.000"

                # ALWAYS create the PSD plot. If result is None, this is a
                # GSD-only plot; if result exists, it contains GSD + HYD.
                _, plot_percentages = generate_psd_plot(
                    self.ws_gsd_values, hr, row["sample_type"], row["depth"],
                    self.bh_id, self.template_type.get(), result,
                    Path(plot_dir)
                )

                summary_rows.append({
                    "bh_id": self.bh_id,
                    "depth": row["depth"],
                    "sample_type": row["sample_type"],
                    "sieve_set": self.template_type.get(),
                    "gravel": plot_percentages[0],
                    "sand": plot_percentages[1],
                    "silt": plot_percentages[2],
                    "clay": plot_percentages[3],
                })

            summary_path = export_psd_summary(summary_rows, plot_dir)

            wb.save(output)
            wb.close()

            self.status.config(
                text=f"Generation completed for {len(selected)} sample(s)."
            )
            messagebox.showinfo(
                "Generation Complete",
                f"HYD data generated successfully.\n\n{output}\n\n"
                f"PSD summary:\n{summary_path}"
            )

        except Exception as e:
            messagebox.showerror("Generation Error", str(e))
            self.status.config(text="Generation stopped because of an error.")


if __name__ == "__main__":
    root = tk.Tk()
    try:
        ttk.Style(root).theme_use("vista")
    except Exception:
        pass
    HydrometerApp(root)
    root.mainloop()
