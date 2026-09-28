"""
PSD Plot Generator
------------------
Reads the PLOT sheet from the generated GSD/HYD workbook and creates:

1. Individual PSD plots for every depth/sample row.
2. Combined PSD plots for groups of up to 5 depths.

PLOT sheet mapping:

TYPE 1
  D:K  -> particle size (mm), 8 sieve points
  Y:AF -> GSD % passing, 8 sieve points
  L:X  -> HYD particle size (mm), 13 points
  AG:AS -> HYD % passing, 13 points

TYPE 2
  D:I  -> particle size (mm), 6 sieve points
  W:AB -> GSD % passing, 6 sieve points
  J:V  -> HYD particle size (mm), 13 points
  AC:AO -> HYD % passing, 13 points

Sample Type is taken from GSD column E by matching BH ID + Depth for internal data grouping only.
BH ID is taken from PLOT column B.
Depth is taken from PLOT column C.
Sample Type is NOT printed on the final plot.

No Particle Size Characteristics table is included.
The individual plot uses the final report-style header with only BH ID and Depth,
classification band, continuous PSD graph, and Percentage of Particles table.
Percentages are calculated from the PSD curve rather than copied from AT:AW.
"""

import os
import re
import math
import csv
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

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

from matplotlib.patches import Rectangle


# ---------------------------------------------------------------------
# USER-ADJUSTABLE SETTINGS
# ---------------------------------------------------------------------
FIG_W = 16.0
FIG_H = 9.4
DPI = 180

X_MIN = 1e-4
X_MAX = 100.0
Y_MIN = 0.0
Y_MAX = 100.0

# Reference design dimensions (figure coordinates)
LEFT = 0.07
RIGHT = 0.985
GRAPH_BOTTOM = 0.205
GRAPH_TOP = 0.765
BAND_BOTTOM = GRAPH_TOP
BAND_TOP = 0.845
INFO_BOTTOM = 0.845
INFO_TOP = 0.915
TITLE_BOTTOM = 0.915
TITLE_TOP = 0.985

# Percentage table
TABLE_BOTTOM = 0.045
TABLE_TOP = 0.135


# ---------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------
def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def norm_text(value):
    return clean_text(value).casefold()


def depth_sort_key(value):
    s = clean_text(value)
    nums = re.findall(r"[-+]?\d*\.?\d+", s)
    if nums:
        try:
            return float(nums[0])
        except Exception:
            pass
    return float("inf")


def safe_filename(text):
    text = clean_text(text)
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    text = re.sub(r"\s+", "_", text)
    return text.strip("_") or "PSD"


def as_float(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        v = float(value)
        if math.isfinite(v):
            return v
    except Exception:
        pass
    return None


def get_row_values(ws, row, start_col, end_col):
    vals = []
    for col in range(start_col, end_col + 1):
        vals.append(as_float(ws.cell(row=row, column=col).value))
    return vals


def paired_points(xvals, yvals):
    points = []
    for x, y in zip(xvals, yvals):
        if x is None or y is None:
            continue
        if x <= 0:
            continue
        y = max(0.0, min(100.0, y))
        points.append((x, y))

    # Sort by particle size descending, as is conventional for PSD plots.
    points.sort(key=lambda p: p[0], reverse=True)

    # Remove duplicate x values, retaining the first occurrence.
    out = []
    seen = set()
    for x, y in points:
        key = round(x, 12)
        if key in seen:
            continue
        seen.add(key)
        out.append((x, y))
    return out


def interpolate_percent_finer(points, target_mm):
    """Interpolate % finer at a particle size on the logarithmic X-axis."""
    if not points or target_mm <= 0:
        return None

    pts = sorted(points, key=lambda p: p[0], reverse=True)

    for x, y in pts:
        if abs(x - target_mm) <= 1e-12:
            return y

    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        if x1 >= target_mm >= x2 and x1 != x2:
            lx1 = math.log10(x1)
            lx2 = math.log10(x2)
            lt = math.log10(target_mm)
            frac = (lt - lx1) / (lx2 - lx1)
            return y1 + frac * (y2 - y1)

    return None


def calculate_percentages(sieve_points, hyd_points):
    """Calculate particle fractions only where the available PSD supports them."""
    all_points = list(sieve_points or []) + list(hyd_points or [])
    all_points.sort(key=lambda p: p[0], reverse=True)

    merged = []
    sieve_x = {round(x, 10) for x, _ in (sieve_points or [])}

    for x, y in all_points:
        key = round(x, 10)
        if merged and key == round(merged[-1][0], 10):
            if key in sieve_x:
                merged[-1] = (x, y)
        else:
            merged.append((x, y))

    p475 = interpolate_percent_finer(merged, 4.75)
    p0075 = interpolate_percent_finer(merged, 0.075)
    p0002 = interpolate_percent_finer(merged, 0.002)

    if p475 is None:
        return [None, None, None, None]

    if p0075 is None:
        return [max(0.0, 100.0 - p475), None, None, None]

    # GSD-only case: sieve analysis defines gravel, sand and total fines
    # directly. Fines are the material passing the 0.075 mm sieve.
    # Do not attempt to split fines into silt and clay without HYD data.
    if not hyd_points:
        gravel = max(0.0, 100.0 - p475)
        sand = max(0.0, p475 - p0075)
        fines = max(0.0, p0075)
        values = [gravel, sand, fines]
        total = sum(values)
        if total > 0:
            values = [v * 100.0 / total for v in values]
        return values

    if p0002 is None:
        gravel = max(0.0, 100.0 - p475)
        sand = max(0.0, p475 - p0075)
        return [gravel, sand, None, None]

    values = [
        100.0 - p475,
        p475 - p0075,
        p0075 - p0002,
        p0002,
    ]
    values = [max(0.0, v) for v in values]

    total = sum(values)
    if total > 0:
        values = [v * 100.0 / total for v in values]

    return values


def get_percentages(ws, row):
    # Retained for workbook compatibility. The plot generator does not use
    # AT:AW for reported percentages; it calculates them from the PSD curve.
    vals = [as_float(ws.cell(row=row, column=c).value) for c in range(46, 50)]
    return [v if v is not None else 0.0 for v in vals]


# ---------------------------------------------------------------------
# SAMPLE TYPE + BOREHOLE SIEVE-SET TYPE
# ---------------------------------------------------------------------
def _depth_key(value):
    """Normalize depth so 18, 18.0 and '18 m' match."""
    if value is None or value == "":
        return None
    s = clean_text(value).replace(",", "")
    nums = re.findall(r"[-+]?\d*\.?\d+", s)
    if not nums:
        return None
    try:
        return round(float(nums[0]), 6)
    except Exception:
        return None


def build_gsd_lookup(gsd_ws):
    """
    Read the actual Sample Type from GSD column E.

    BH ID is column B, Sample Type is column E, and Depth is column F.
    GSD data begins at row 5. BH ID is carried downward when B is blank.
    """
    lookup = {}
    current_bh = clean_text(gsd_ws["B5"].value)

    # GSD data starts from row 5.
    # In the GSD sheet:
    #   B = BH ID
    #   E = Sample Type
    #   F = Depth
    # Therefore depth is read from column F.
    depth_cols = [6]

    for r in range(5, gsd_ws.max_row + 1):
        row_bh = clean_text(gsd_ws.cell(r, 2).value)
        if row_bh:
            current_bh = row_bh

        sample_type = clean_text(gsd_ws.cell(r, 5).value)
        if not current_bh or not sample_type:
            continue

        dk = None
        for c in depth_cols:
            if c == 5:
                continue
            candidate = _depth_key(gsd_ws.cell(r, c).value)
            if candidate is not None:
                dk = candidate
                break

        if dk is not None:
            lookup[(norm_text(current_bh), dk)] = sample_type

    return lookup


def find_sample_type(gsd_lookup, bh_id, depth):
    """Return the actual Sample Type stored in GSD column E."""
    st = gsd_lookup.get((norm_text(bh_id), _depth_key(depth)))

    if st:
        return st

    raise ValueError(
        f"Could not find Sample Type in GSD Column E for "
        f"BH ID='{bh_id}', Depth='{depth}'."
    )


def detect_borehole_sieve_type(plot_ws, bh_id):
    """
    Determine Type 1 / Type 2 ONCE per borehole from the fixed sieve
    particle-size sequence in PLOT.

    Type 1:
      D:K = 4.75, 2.00, 1.18, 0.600, 0.425, 0.300, 0.150, 0.075

    Type 2:
      D:I = 4.75, 2.00, 0.425, 0.212, 0.150, 0.075

    Sample Type (SPT/DS/UDS) is not used here.
    """
    expected1 = [4.75, 2.00, 1.18, 0.600, 0.425, 0.300, 0.150, 0.075]
    expected2 = [4.75, 2.00, 0.425, 0.212, 0.150, 0.075]
    key = norm_text(bh_id)
    t1 = 0
    t2 = 0

    for r in range(3, plot_ws.max_row + 1):
        if norm_text(clean_text(plot_ws.cell(r, 2).value)) != key:
            continue

        vals1 = [as_float(plot_ws.cell(r, c).value) for c in range(4, 12)]
        vals2 = [as_float(plot_ws.cell(r, c).value) for c in range(4, 10)]

        m1 = sum(v is not None and abs(v-e) <= 1e-6 for v, e in zip(vals1, expected1))
        m2 = sum(v is not None and abs(v-e) <= 1e-6 for v, e in zip(vals2, expected2))
        t1 = max(t1, m1)
        t2 = max(t2, m2)

    if t1 == 8 and t2 < 6:
        return "Type 1"
    if t2 == 6 and t1 < 8:
        return "Type 2"
    if t1 > t2:
        return "Type 1"
    if t2 > t1:
        return "Type 2"

    raise ValueError(
        f"Could not determine Type 1/Type 2 for borehole '{bh_id}'. "
        "Check the fixed sieve sizes in PLOT D:K / D:I."
    )




# ---------------------------------------------------------------------
# PLOT DATA EXTRACTION
# ---------------------------------------------------------------------
def extract_plot_data(plot_ws, row, sieve_type):
    st = norm_text(sieve_type)

    if "type 1" in st or st in ("1", "type1"):
        sieve_x_cols = range(4, 12)       # D:K
        sieve_y_cols = range(25, 33)      # Y:AF
        hyd_x_cols = range(12, 25)        # L:X
        hyd_y_cols = range(33, 46)        # AG:AS
        type_label = "Type 1"
    elif "type 2" in st or st in ("2", "type2"):
        sieve_x_cols = range(4, 10)       # D:I
        sieve_y_cols = range(23, 29)      # W:AB
        hyd_x_cols = range(10, 23)        # J:V
        hyd_y_cols = range(29, 42)        # AC:AO
        type_label = "Type 2"
    else:
        raise ValueError(
            f"Unknown sieve-set type '{sieve_type}' in PLOT row {row}. "
            f"Expected Type 1 or Type 2."
        )

    sieve_x = [as_float(plot_ws.cell(row=row, column=c).value) for c in sieve_x_cols]
    sieve_y = [as_float(plot_ws.cell(row=row, column=c).value) for c in sieve_y_cols]

    hyd_x = [as_float(plot_ws.cell(row=row, column=c).value) for c in hyd_x_cols]
    hyd_y = [as_float(plot_ws.cell(row=row, column=c).value) for c in hyd_y_cols]

    sieve_points = paired_points(sieve_x, sieve_y)
    hyd_points = paired_points(hyd_x, hyd_y)

    return {
        "type": type_label,
        "sieve": sieve_points,
        "hyd": hyd_points,
        "percentages": calculate_percentages(sieve_points, hyd_points),
    }


# ---------------------------------------------------------------------
# DRAWING
# ---------------------------------------------------------------------
def draw_header(fig, bh_id, sample_type, depth, combined=False):
    """Draw the report-style title and two-field information header.

    Final reference layout:
      - Title centered at top.
      - Subtitle directly below title.
      - Information row contains ONLY BH ID and Depth.
      - No Sample Type is displayed in the plot header.
      - The classification band sits below the information row and above
        the graph without overlapping either section.
    """
    fig.text(
        0.5, 0.958,
        "PARTICLE SIZE DISTRIBUTION CURVE",
        ha="center", va="center",
        fontsize=21, fontweight="bold"
    )
    fig.text(
        0.5, 0.928,
        "(Sieve Analysis + Hydrometer Analysis)",
        ha="center", va="center",
        fontsize=12
    )

    # Main outer border.
    fig.patches.append(
        Rectangle(
            (0.008, 0.008), 0.984, 0.984,
            transform=fig.transFigure,
            fill=False, linewidth=1.3
        )
    )

    # Title / information-row separator.
    fig.lines.append(
        plt.Line2D(
            [0.008, 0.992], [TITLE_BOTTOM, TITLE_BOTTOM],
            transform=fig.transFigure,
            linewidth=1.0
        )
    )

    # Information row / classification-band separator.
    fig.lines.append(
        plt.Line2D(
            [0.008, 0.992], [INFO_BOTTOM, INFO_BOTTOM],
            transform=fig.transFigure,
            linewidth=1.0
        )
    )

    # Only one vertical separator: BH ID | Depth.
    x_mid = 0.675
    fig.lines.append(
        plt.Line2D(
            [x_mid, x_mid], [INFO_BOTTOM, INFO_TOP],
            transform=fig.transFigure,
            linewidth=1.0
        )
    )

    depth_text = "5 Depths" if combined else clean_text(depth)

    # Final reference positions.
    fields = [
        (0.030, "BH ID", bh_id, 0.080, 0.115),
        (0.695, "Depth", depth_text, 0.080, 0.115),
    ]

    for x, label, value, colon_offset, value_offset in fields:
        fig.text(
            x, 0.878, label,
            ha="left", va="center",
            fontsize=11.5, fontweight="bold"
        )
        fig.text(
            x + colon_offset, 0.878, ":",
            ha="center", va="center",
            fontsize=11.5, fontweight="bold"
        )
        fig.text(
            x + value_offset, 0.878, clean_text(value),
            ha="left", va="center",
            fontsize=11.5
        )


def draw_classification_band(fig):
    """Draw the soil-size classification band in the header/plot gap.

    Cobble classification is omitted.

    Major classes:
      GRAVEL : 75 to 4.75 mm
      SAND   : 4.75 to 0.075 mm
      SILT   : 0.075 to 0.002 mm
      CLAY   : 0.002 to 0.0001 mm

    Subdivisions:
      Gravel: Coarse 75-20 mm, Fine 20-4.75 mm
      Sand:   Coarse 4.75-2.0 mm, Medium 2.0-0.425 mm,
              Fine 0.425-0.075 mm
    """

    def xfig(mm):
        lx = math.log10(mm)
        lo = math.log10(X_MIN)
        hi = math.log10(X_MAX)
        frac = (hi - lx) / (hi - lo)
        return LEFT + frac * (RIGHT - LEFT)

    # GREEN BOX LOCATION FROM THE USER'S MARKUP:
    # classification band occupies the complete space between the
    # information row and the graph. It must NOT overlap the graph.
    band_bottom = GRAPH_TOP + 0.004
    band_top = INFO_BOTTOM - 0.008
    band_h = band_top - band_bottom
    ymid = band_bottom + 0.52 * band_h

    fig.patches.append(Rectangle(
        (LEFT, band_bottom),
        RIGHT - LEFT,
        band_h,
        transform=fig.transFigure,
        facecolor="white",
        edgecolor="black",
        linewidth=0.8,
        zorder=8
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
            (xa, band_bottom),
            xb - xa,
            band_h,
            transform=fig.transFigure,
            fill=False,
            linewidth=0.75,
            zorder=9
        ))

        fig.text(
            (xa + xb) / 2,
            band_bottom + 0.72 * band_h,
            name,
            ha="center",
            va="center",
            fontsize=9.2,
            fontweight="bold",
            zorder=10
        )

    # Horizontal division between class-name and subdivision rows.
    fig.lines.append(plt.Line2D(
        [LEFT, RIGHT],
        [ymid, ymid],
        transform=fig.transFigure,
        linewidth=0.7,
        zorder=9
    ))

    # Subdivision boundaries.
    for mm in (20.0, 2.0, 0.425):
        xx = xfig(mm)
        if LEFT <= xx <= RIGHT:
            fig.lines.append(plt.Line2D(
                [xx, xx],
                [band_bottom, ymid],
                transform=fig.transFigure,
                linewidth=0.7,
                zorder=9
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
            fig.text(
                (xa + xb) / 2,
                band_bottom + 0.20 * band_h,
                label,
                ha="center",
                va="center",
                fontsize=8.2,
                fontweight="bold",
                zorder=10
            )


def merge_curve_points(sieve_points, hyd_points):
    """Combine available GSD and HYD points for combined-depth plotting."""
    points = []
    for x, y in (sieve_points or []):
        points.append((x, y, "sieve"))
    for x, y in (hyd_points or []):
        points.append((x, y, "hyd"))

    points.sort(key=lambda p: p[0], reverse=True)

    merged = []
    for x, y, source in points:
        if merged and abs(x - merged[-1][0]) <= 1e-10:
            if source == "sieve":
                merged[-1] = (x, y, source)
        else:
            merged.append((x, y, source))
    return merged


def draw_graph(fig, series, combined=False):
    ax = fig.add_axes([
        LEFT, GRAPH_BOTTOM,
        RIGHT - LEFT, GRAPH_TOP - GRAPH_BOTTOM
    ])

    ax.set_xscale("log")
    ax.set_xlim(X_MAX, X_MIN)
    ax.set_ylim(Y_MIN, Y_MAX)

    # Keep the X-axis logarithmic, but display major tick labels in
    # ordinary decimal notation instead of 10^n / scientific notation.
    ax.xaxis.set_major_locator(
        mticker.LogLocator(base=10.0, subs=(1.0,), numticks=20)
    )
    ax.xaxis.set_major_formatter(DecimalLogFormatter())
    ax.xaxis.set_minor_locator(
        mticker.LogLocator(base=10.0, subs=tuple(range(2, 10)), numticks=100)
    )
    ax.xaxis.set_minor_formatter(mticker.NullFormatter())

    ax.set_xlabel(
        "PARTICLE SIZE / GRAIN SIZE (mm)",
        fontsize=11.5, fontweight="bold", labelpad=8
    )
    ax.set_ylabel(
        "PERCENTAGE FINER (%)",
        fontsize=11.5, fontweight="bold", labelpad=8
    )

    ax.set_yticks(range(0, 101, 10))
    ax.grid(True, which="major", linewidth=0.75, alpha=0.42)
    ax.grid(True, which="minor", linewidth=0.35, alpha=0.18)

    if combined:
        # One curve per depth.
        # If HYD exists, connect the actual GSD endpoint to the actual
        # first HYD point. If HYD is absent, plot GSD only.
        for item in series:
            sieve = item.get("sieve", [])
            hyd = item.get("hyd", [])

            if sieve and hyd:
                # GSD points are ordered from coarse -> fine.
                # Therefore the transition point is the SMALLEST GSD size.
                gsd_end = min(sieve, key=lambda p: p[0])

                # HYD starts at its LARGEST particle size.
                hyd_start = max(hyd, key=lambda p: p[0])

                # Plot GSD portion.
                sx, sy = zip(*sieve)
                ax.plot(
                    sx, sy,
                    marker="o", markersize=3.8,
                    linewidth=1.15,
                    label=str(item["depth"])
                )

                # Connect GSD -> HYD using the HYD curve colour.
                # This is the only transition segment and is NOT black.
                if abs(gsd_end[0] - hyd_start[0]) > 1e-10:
                    ax.plot(
                        [gsd_end[0], hyd_start[0]],
                        [gsd_end[1], hyd_start[1]],
                        color="red",
                        linewidth=1.15,
                        zorder=2
                    )

                # Plot HYD portion.
                hx, hy = zip(*hyd)
                ax.plot(
                    hx, hy,
                    color="red",
                    marker="o", markersize=3.8,
                    linewidth=1.15,
                    zorder=3
                )

            elif sieve:
                sx, sy = zip(*sieve)
                ax.plot(
                    sx, sy,
                    marker="o", markersize=3.8,
                    linewidth=1.15,
                    label=str(item["depth"])
                )

            elif hyd:
                hx, hy = zip(*hyd)
                ax.plot(
                    hx, hy,
                    marker="o", markersize=3.8,
                    linewidth=1.15,
                    label=str(item["depth"])
                )

        handles, labels = ax.get_legend_handles_labels()
        if handles:
            ax.legend(
                loc="lower left",
                fontsize=8.5,
                frameon=True,
                ncol=2,
                title="Depth",
                title_fontsize=8.5
            )

    else:
        item = series[0]
        sieve = item.get("sieve", [])
        hyd = item.get("hyd", [])

        # Sieve analysis: blue.
        if sieve:
            sx, sy = zip(*sieve)
            ax.plot(
                sx, sy,
                color="blue",
                marker="o",
                markersize=4.8,
                linewidth=1.2,
                label="Sieve Analysis",
                zorder=3
            )

        # Hydrometer analysis: red.
        if hyd:
            hx, hy = zip(*hyd)
            ax.plot(
                hx, hy,
                color="red",
                marker="o",
                markersize=4.8,
                linewidth=1.2,
                label="Hydrometer Analysis",
                zorder=3
            )

        # If both analyses are available, connect the TRUE transition:
        # smallest GSD particle size -> largest HYD particle size.
        #
        # IMPORTANT:
        # Do NOT use max(sieve). That would connect 4.75 mm to the
        # hydrometer region and create the long diagonal line seen earlier.
        if sieve and hyd:
            gsd_end = min(sieve, key=lambda p: p[0])
            hyd_start = max(hyd, key=lambda p: p[0])

            if abs(gsd_end[0] - hyd_start[0]) > 1e-10:
                ax.plot(
                    [gsd_end[0], hyd_start[0]],
                    [gsd_end[1], hyd_start[1]],
                    color="red",
                    linewidth=1.2,
                    zorder=2
                )

        handles, labels = ax.get_legend_handles_labels()
        if handles:
            ax.legend(
                loc="lower left",
                fontsize=8.5,
                frameon=True
            )

    return ax


def draw_percentage_table(fig, percentages, combined=False, hyd_present=False):
    """Draw the Percentage of Particles table.

    If HYD is present:
        Gravel | Sand | Silt | Clay

    If HYD is NOT present:
        Gravel | Sand | Fines

    For GSD-only samples, Fines means material passing the 0.075 mm sieve.
    """
    x0, x1 = 0.305, 0.715
    y0, y1 = TABLE_BOTTOM, TABLE_TOP

    fig.patches.append(
        Rectangle(
            (x0, y0),
            x1 - x0,
            y1 - y0,
            transform=fig.transFigure,
            fill=False,
            linewidth=0.9
        )
    )

    header_y = y1 - 0.035

    fig.lines.append(
        plt.Line2D(
            [x0, x1], [header_y, header_y],
            transform=fig.transFigure,
            linewidth=0.8
        )
    )

    fig.text(
        (x0 + x1) / 2,
        y1 - 0.017,
        "PERCENTAGE OF PARTICLES",
        ha="center",
        va="center",
        fontsize=11,
        fontweight="bold"
    )

    if hyd_present:
        labels = ["Gravel (%)", "Sand (%)", "Silt (%)", "Clay (%)"]
        vals = percentages[:4]
    else:
        labels = ["Gravel (%)", "Sand (%)", "Fines (%)"]
        vals = percentages[:3]

    col_w = (x1 - x0) / len(labels)

    for i in range(1, len(labels)):
        xx = x0 + i * col_w
        fig.lines.append(
            plt.Line2D(
                [xx, xx],
                [y0, header_y],
                transform=fig.transFigure,
                linewidth=0.7
            )
        )

    for i, (label, value) in enumerate(zip(labels, vals)):
        xc = x0 + (i + 0.5) * col_w

        fig.text(
            xc,
            header_y - 0.020,
            label,
            ha="center",
            va="center",
            fontsize=9.2,
            fontweight="bold"
        )

        display_value = "—" if value is None else f"{value:.2f}"

        fig.text(
            xc,
            y0 + 0.018,
            display_value,
            ha="center",
            va="center",
            fontsize=10
        )


def create_individual_plot(item, output_path):
    fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=DPI)

    draw_header(
        fig,
        item["bh_id"],
        item["sample_type"],
        item["depth"],
        combined=False
    )
    draw_classification_band(fig)
    draw_graph(fig, [item], combined=False)
    draw_percentage_table(fig, item["percentages"], combined=False, hyd_present=bool(item["hyd"]))

    fig.savefig(
        output_path,
        dpi=DPI,
        facecolor="white",
        edgecolor="white",
        bbox_inches=None,
        pad_inches=0
    )
    plt.close(fig)


def create_combined_plot(items, output_path):
    if not items:
        return

    bh_id = items[0]["bh_id"]
    sample_type = items[0]["sample_type"]

    fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=DPI)

    draw_header(
        fig,
        bh_id,
        sample_type,
        "5 Depths",
        combined=True
    )
    draw_classification_band(fig)
    draw_graph(fig, items, combined=True)

    # No particle size characteristics table.
    # No percentage table on combined graph because percentages vary by depth.
    fig.savefig(
        output_path,
        dpi=DPI,
        facecolor="white",
        edgecolor="white",
        bbox_inches=None,
        pad_inches=0
    )
    plt.close(fig)


# ---------------------------------------------------------------------
# MAIN WORKFLOW
# ---------------------------------------------------------------------
def choose_workbook():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    path = filedialog.askopenfilename(
        title="Select GSD & HYD Generated Workbook",
        filetypes=[
            ("Excel Macro-Enabled Workbook", "*.xlsm"),
            ("Excel Workbook", "*.xlsx"),
            ("All files", "*.*"),
        ]
    )

    root.destroy()
    return path


def load_workbook(path):
    keep_vba = str(path).lower().endswith(".xlsm")
    return openpyxl.load_workbook(path, data_only=True, keep_vba=keep_vba)


def collect_items(wb):
    if "PLOT" not in wb.sheetnames:
        raise ValueError("The selected workbook does not contain a 'PLOT' sheet.")
    if "GSD" not in wb.sheetnames:
        raise ValueError("The selected workbook does not contain a 'GSD' sheet.")

    plot_ws = wb["PLOT"]
    gsd_ws = wb["GSD"]
    gsd_lookup = build_gsd_lookup(gsd_ws)

    # PLOT rows 1–2 are headers. Actual sample data start at row 3.
    # Detect sieve-set type once per BH and reuse it for every depth.
    bh_ids = []
    seen = set()
    for row in range(3, plot_ws.max_row + 1):
        bh_id = clean_text(plot_ws.cell(row=row, column=2).value)
        depth = plot_ws.cell(row=row, column=3).value
        if not bh_id or depth in (None, ""):
            continue
        if norm_text(bh_id) not in seen:
            seen.add(norm_text(bh_id))
            bh_ids.append(bh_id)

    borehole_types = {
        bh: detect_borehole_sieve_type(plot_ws, bh)
        for bh in bh_ids
    }

    items = []
    for row in range(3, plot_ws.max_row + 1):
        bh_id = clean_text(plot_ws.cell(row=row, column=2).value)
        depth = clean_text(plot_ws.cell(row=row, column=3).value)
        if not bh_id or depth == "":
            continue

        sample_type = find_sample_type(gsd_lookup, bh_id, depth)
        sieve_type = borehole_types[bh_id]
        data = extract_plot_data(plot_ws, row, sieve_type)

        if not data["sieve"] and not data["hyd"]:
            continue

        items.append({
            "row": row,
            "bh_id": bh_id,
            "depth": depth,
            "sample_type": sample_type,
            "sieve_type": sieve_type,
            **data,
        })

    items.sort(key=lambda x: (norm_text(x["bh_id"]), depth_sort_key(x["depth"])))
    return items

def export_psd_summary(items, output_dir):
    """Export Gravel/Sand/Silt/Clay percentages for every PSD item to CSV."""
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

    with open(summary_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)

        for item in items:
            vals = item.get("percentages", [None, None, None, None])

            # For GSD-only data, the existing calculation returns:
            # [Gravel, Sand, Fines]. There is no defensible split of
            # fines into silt and clay without hydrometer data.
            if item.get("hyd"):
                gravel, sand, silt, clay = vals[:4]
            else:
                gravel = vals[0] if len(vals) > 0 else None
                sand = vals[1] if len(vals) > 1 else None
                silt = None
                clay = None

            def fmt(v):
                return "" if v is None else f"{v:.2f}"

            writer.writerow([
                item["bh_id"],
                item["depth"],
                item["sample_type"],
                item["sieve_type"],
                fmt(gravel),
                fmt(sand),
                fmt(silt),
                fmt(clay),
            ])

    return summary_path


def generate_all(path):
    wb = load_workbook(path)
    items = collect_items(wb)

    if not items:
        raise ValueError("No usable PSD data were found in the PLOT sheet.")

    source = Path(path)
    out_dir = source.parent / "PSD_Plots"
    out_dir.mkdir(parents=True, exist_ok=True)

    individual_dir = out_dir / "Individual_Depths"
    combined_dir = out_dir / "Combined_5_Depths"
    individual_dir.mkdir(exist_ok=True)
    combined_dir.mkdir(exist_ok=True)

    # Export PSD percentage summary for all depths.
    summary_path = export_psd_summary(items, out_dir)

    individual_count = 0
    combined_count = 0

    # Individual plots
    for idx, item in enumerate(items, start=1):
        filename = (
            f"{safe_filename(item['bh_id'])}_"
            f"{safe_filename(item['sample_type'])}_"
            f"{safe_filename(item['depth'])}.png"
        )
        create_individual_plot(item, individual_dir / filename)
        individual_count += 1

    # Combined plots: groups of 5 depths within the same borehole.
    # Sample Type is independent of sieve-set type, so a combined group may
    # contain SPT/DS/UDS depths. The header will show "Multiple" in that case.
    groups = {}
    for item in items:
        key = norm_text(item["bh_id"])
        groups.setdefault(key, []).append(item)

    for bh_key, group in groups.items():
        for start in range(0, len(group), 5):
            chunk = group[start:start + 5]

            # User requested 5-depth combined plots. Do not make a combined
            # plot for a group containing fewer than 5 depths.
            if len(chunk) < 5:
                continue

            first = chunk[0]
            last = chunk[-1]
            sample_types = {norm_text(x["sample_type"]): x["sample_type"] for x in chunk}
            combined_sample_type = next(iter(sample_types.values())) if len(sample_types) == 1 else "Multiple"
            filename = (
                f"{safe_filename(first['bh_id'])}_"
                f"{safe_filename(combined_sample_type)}_"
                f"Depths_{safe_filename(first['depth'])}_to_"
                f"{safe_filename(last['depth'])}.png"
            )

            create_combined_plot(chunk, combined_dir / filename)
            combined_count += 1

    return out_dir, individual_count, combined_count, summary_path


def main():
    path = choose_workbook()
    if not path:
        return

    try:
        out_dir, individual_count, combined_count, summary_path = generate_all(path)

        messagebox.showinfo(
            "PSD Plot Generation Complete",
            f"Individual plots generated: {individual_count}\n"
            f"Combined 5-depth plots generated: {combined_count}\n\n"
            f"PSD summary exported:\n{summary_path}\n\n"
            f"Output folder:\n{out_dir}"
        )

    except Exception as exc:
        messagebox.showerror(
            "PSD Plot Generation Error",
            str(exc)
        )
        raise


if __name__ == "__main__":
    main()
