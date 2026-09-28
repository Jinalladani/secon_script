import math
import random
import csv
import re
from pathlib import Path
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.patches import Arc
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
from copy import copy

try:
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
except ImportError:
    load_workbook = None
    MergedCell = ()

RATE_MM_MIN = 1.25
LVDT_RESOLUTION_MM = 0.001

# Soil-specific tendencies. These are not fixed curves; they control the
# statistical shape of the continuous resistance-development model.
SOIL = {
    "Soft Clay": {
        "peak": (9.0, 14.5),
        "rise": (0.34, 0.44),
        "hard": (0.72, 0.84),
        "peak_width": (1.10, 1.30),
        "soft_end": (1.48, 1.75),
        "drop": (0.08, 0.18),
        "res": (0.72, 0.84),
        "noise": (0.010, 0.025),
        "seat": (2, 5),
    },
    "Medium Clay": {
        "peak": (7.0, 11.5),
        "rise": (0.30, 0.41),
        "hard": (0.70, 0.82),
        "peak_width": (1.06, 1.22),
        "soft_end": (1.42, 1.68),
        "drop": (0.08, 0.17),
        "res": (0.74, 0.86),
        "noise": (0.009, 0.022),
        "seat": (2, 4),
    },
    "Stiff Clay": {
        "peak": (5.0, 8.5),
        "rise": (0.26, 0.37),
        "hard": (0.68, 0.80),
        "peak_width": (1.04, 1.16),
        "soft_end": (1.35, 1.58),
        "drop": (0.10, 0.21),
        "res": (0.72, 0.84),
        "noise": (0.008, 0.019),
        "seat": (1, 3),
    },
    "Very Stiff Clay": {
        "peak": (3.5, 6.5),
        "rise": (0.22, 0.34),
        "hard": (0.66, 0.78),
        "peak_width": (1.02, 1.12),
        "soft_end": (1.30, 1.52),
        "drop": (0.14, 0.27),
        "res": (0.66, 0.81),
        "noise": (0.007, 0.017),
        "seat": (1, 3),
    },
    "Silty Clayey Soil": {
        "peak": (7.0, 12.0),
        "rise": (0.31, 0.43),
        "hard": (0.70, 0.84),
        "peak_width": (1.08, 1.28),
        "soft_end": (1.43, 1.72),
        "drop": (0.10, 0.23),
        "res": (0.67, 0.83),
        "noise": (0.013, 0.030),
        "seat": (2, 5),
    },
    "Silty Clayey Sandy Soil": {
        "peak": (5.5, 10.0),
        "rise": (0.27, 0.39),
        "hard": (0.68, 0.82),
        "peak_width": (1.06, 1.23),
        "soft_end": (1.38, 1.65),
        "drop": (0.13, 0.29),
        "res": (0.61, 0.80),
        "noise": (0.015, 0.035),
        "seat": (2, 5),
    },
}


def mc_peak(c, phi, sigma3):
    p = math.radians(phi)
    s = math.sin(p)
    co = math.cos(p)
    sigma1 = sigma3 * (1 + s) / (1 - s) + 2 * c * co / (1 - s)
    return max(0.0, sigma1 - sigma3), sigma1


def smoothstep(x):
    x = max(0.0, min(1.0, x))
    return x * x * (3.0 - 2.0 * x)


def smootherstep(x):
    x = max(0.0, min(1.0, x))
    return x * x * x * (x * (x * 6.0 - 15.0) + 10.0)


def choose_peak_strain(soil, sigma3, c, phi, rng, pressures):
    """
    Peak is an interval, not a fixed strain.
    Cell pressure and strength influence the centre of the interval,
    while the soil type controls its broad statistical range.
    """
    prof = SOIL[soil]
    lo, hi = prof["peak"]

    pmin, pmax = min(pressures), max(pressures)
    pressure_position = (
        (sigma3 - pmin) / (pmax - pmin) if pmax > pmin else 0.5
    )

    # Pressure effect is deliberately small; soil behaviour remains dominant.
    pressure_effect = (pressure_position - 0.5) * rng.uniform(-1.0, 1.0)

    # Higher strength can shift peak strain slightly, but not deterministically.
    strength_index = c + 0.08 * phi
    strength_effect = max(-0.8, min(0.8, (strength_index - 1.0))) * rng.uniform(-0.8, 0.8)

    peak = rng.uniform(lo, hi) + pressure_effect + strength_effect
    return max(lo, min(hi, peak))


def build_curve_parameters(soil, sigma3, c, phi, rng, pressures, peak_pct):
    """
    Construct a continuous set of anchor points.

    The anchors represent behavioural zones:
        seating -> gradual rise -> progressive hardening ->
        broad peak -> gradual softening -> residual.

    Smooth transitions between anchors eliminate artificial threshold jumps.
    """
    prof = SOIL[soil]

    rise_end = rng.uniform(*prof["rise"])
    hard_end = rng.uniform(*prof["hard"])

    # Peak zone is deliberately broad. peak_width is the end of the
    # near-maximum resistance region, expressed relative to peak-centre strain.
    peak_zone_end = rng.uniform(*prof["peak_width"])
    peak_zone_end = max(hard_end + 0.10, peak_zone_end)

    soft_end = rng.uniform(*prof["soft_end"])
    soft_end = max(peak_zone_end + 0.12, soft_end)

    residual = rng.uniform(*prof["res"])
    drop = rng.uniform(*prof["drop"])

    # Resistance fractions at the anchors.
    # Hardening is progressive: increments become larger as the peak is approached.
    y_rise = rng.uniform(0.34, 0.46)
    y_hard = rng.uniform(0.72, 0.84)
    y_peak_start = rng.uniform(0.965, 0.985)

    # The broad peak is slightly asymmetric; resistance may stay very close
    # to maximum after the actual maximum point.
    y_peak_end = rng.uniform(0.985, 0.997)

    # End of softening is above residual; the final transition is gradual.
    y_soft_end = residual + rng.uniform(0.07, 0.14)
    y_soft_end = min(0.94, max(residual + 0.04, y_soft_end))

    return {
        "rise_end": rise_end,
        "hard_end": hard_end,
        "peak_start": rng.uniform(0.90, 0.96),
        "peak_end": peak_zone_end,
        "soft_end": soft_end,
        "residual": residual,
        "drop": drop,
        "y_rise": y_rise,
        "y_hard": y_hard,
        "y_peak_start": y_peak_start,
        "y_peak_end": y_peak_end,
        "y_soft_end": y_soft_end,
        "noise": rng.uniform(*prof["noise"]),
    }


def continuous_resistance(x, p):
    """
    Continuous normalized resistance function.

    x = current strain / actual peak-centre strain.

    The function is continuous and has smooth derivatives at the zone
    boundaries. There are no hard 'if strain > threshold -> multiply load'
    jumps in the physical curve.
    """
    if x <= 0:
        return 0.0

    # Small seating region.
    seating_end = 0.045
    if x <= seating_end:
        z = smootherstep(x / seating_end)
        return 0.015 * z

    # Gradual rise.
    if x <= p["rise_end"]:
        z = smootherstep(
            (x - seating_end) / (p["rise_end"] - seating_end)
        )
        return 0.015 + (p["y_rise"] - 0.015) * z

    # Progressive hardening.
    if x <= p["hard_end"]:
        z = smootherstep(
            (x - p["rise_end"]) / (p["hard_end"] - p["rise_end"])
        )
        # Slightly accelerating hardening. The smooth transition is retained.
        z2 = z * (0.78 + 0.22 * z)
        return p["y_rise"] + (p["y_hard"] - p["y_rise"]) * z2

    # Approach to broad peak.
    if x <= p["peak_start"]:
        z = smootherstep(
            (x - p["hard_end"]) / (p["peak_start"] - p["hard_end"])
        )
        return p["y_hard"] + (p["y_peak_start"] - p["y_hard"]) * z

    # Actual peak centre: smoothly rise to 1.0.
    if x <= 1.0:
        z = smootherstep(
            (x - p["peak_start"]) / (1.0 - p["peak_start"])
        )
        return p["y_peak_start"] + (1.0 - p["y_peak_start"]) * z

    # Broad peak after the maximum. This is intentionally very flat.
    if x <= p["peak_end"]:
        z = smootherstep((x - 1.0) / (p["peak_end"] - 1.0))
        return 1.0 + (p["y_peak_end"] - 1.0) * z

    # Gradual softening.
    if x <= p["soft_end"]:
        z = smootherstep(
            (x - p["peak_end"]) / (p["soft_end"] - p["peak_end"])
        )
        return p["y_peak_end"] + (p["y_soft_end"] - p["y_peak_end"]) * z

    # Irregular residual baseline.
    return p["residual"]


def next_observation(t, deformation, rng, interval_mm):
    """
    Generate logger observations around the requested nominal deformation
    interval. The interval is approximate, not a forced exact grid.
    """
    mean_dt = interval_mm / RATE_MM_MIN * 60.0

    # Normal machine/data-logger timing scatter.
    dt = rng.triangular(mean_dt * 0.72, mean_dt * 1.28, mean_dt)

    # Occasional shorter/longer observation gaps.
    r = rng.random()
    if r < 0.06:
        dt *= rng.uniform(0.55, 0.80)
    elif r < 0.10:
        dt *= rng.uniform(1.20, 1.45)

    dt = max(1.0, round(dt))

    # Rate is centred around 1.25 mm/min. A very small observation noise
    # prevents a perfectly mechanical deformation sequence.
    increment = (
        RATE_MM_MIN * dt / 60.0
        + rng.gauss(0.0, max(0.001, interval_mm * 0.020))
    )
    increment = max(0.001, increment)

    return (
        t + timedelta(seconds=dt),
        round(deformation + increment, 3),
    )


def correlated_noise(previous, rng, amplitude):
    """
    Low-frequency measurement noise.

    This is deliberately correlated between observations so that the curve
    looks like a real logger trace rather than independent white noise.
    """
    return 0.78 * previous + rng.gauss(0.0, amplitude * 0.42)


def generate_test(
    test_no, sigma3, c, phi, diameter, height, soil,
    start, rng, pressures, interval_mm
):
    prof = SOIL[soil]
    h_mm = height * 10.0
    a0 = math.pi * diameter ** 2 / 4.0

    # Ideal target deviator stress from c and phi.
    target_q, target_s1 = mc_peak(c, phi, sigma3)

    # Real laboratory specimens rarely give exactly the theoretical peak
    # strength.  Introduce a small specimen-to-specimen strength variation.
    # This is deliberately restrained so the results remain close to the
    # requested c/phi relationship while preventing all three circles from
    # appearing perfectly tangent to one common failure envelope.
    strength_factor = rng.triangular(0.90, 1.10, 1.00)
    actual_target_q = target_q * strength_factor
    actual_target_s1 = sigma3 + actual_target_q

    # Actual peak centre is random within a soil/pressure-dependent range.
    peak_pct = choose_peak_strain(
        soil, sigma3, c, phi, rng, pressures
    )

    # Final deformation is independently selected between 15 and 20%.
    final_pct = rng.uniform(15.0, 20.0)

    # For very large requested intervals, still keep final LVDT resolution at 0.001 mm.
    final_mm = round(h_mm * final_pct / 100.0, 3)

    curve = build_curve_parameters(
        soil, sigma3, c, phi, rng, pressures, peak_pct
    )

    # Start of each test always has zero deformation and zero load.
    rows = [{
        "Seq#": None,
        "DateTime": start,
        "STRAIN": 0.000,
        "LOAD": 0.000,
        "_test": test_no,
    }]

    t = start
    deformation = 0.0
    noise_state = 0.0
    seating_points = rng.randint(*prof["seat"])
    i = 1

    while deformation < final_mm:
        # Existing XLSM template supports observation rows 14:169.
        # Keep the safety limit rather than silently exceeding the template.
        if len(rows) >= 155:
            t += timedelta(seconds=rng.uniform(3.5, 7.0))
            deformation = final_mm
        else:
            t, deformation = next_observation(
                t, deformation, rng, interval_mm
            )
            if deformation > final_mm:
                deformation = final_mm

        eps = deformation / h_mm
        x = eps / (peak_pct / 100.0)

        baseline = continuous_resistance(x, curve)

        # Before the residual zone, noise is relatively small. In the
        # residual zone, fluctuations increase slightly.
        if x <= curve["peak_end"]:
            noise_amp = curve["noise"]
        else:
            noise_amp = curve["noise"] * 1.25

        noise_state = correlated_noise(
            noise_state, rng, noise_amp
        )

        # Noise is damped around the peak so that random measurement noise
        # cannot create a false secondary peak.
        peak_damping = 0.45 if 0.90 <= x <= curve["peak_end"] else 1.0
        qnorm = baseline * (
            1.0 + peak_damping * noise_state
            + rng.gauss(0.0, noise_amp * 0.20)
        )

        # Initial LVDT movement before the load cell responds.
        if i <= seating_points:
            if rng.random() < 0.72:
                qnorm = 0.0
            else:
                qnorm *= rng.uniform(0.02, 0.10)

        qnorm = max(0.0, qnorm)

        # Do not allow random noise to exceed the physical peak.
        # The physical peak is now specimen-specific through strength_factor.
        if x < 1.0:
            qnorm = min(qnorm, 0.997)
        elif x <= curve["peak_end"]:
            qnorm = min(qnorm, 1.000)

        # During residual behaviour, keep the baseline physically meaningful.
        if x > curve["soft_end"]:
            qnorm = max(curve["residual"] * 0.94, qnorm)

        q = actual_target_q * qnorm

        # Corrected area as deformation increases.
        area = a0 / (1.0 - eps)
        load = q * area

        # Small probability of an unchanged load-cell reading, representing
        # logger quantisation/adjustment effects. This does NOT create spikes.
        if rows and rng.random() < 0.045 and eps > 0.025:
            previous_load = rows[-1]["LOAD"]
            if abs(load - previous_load) < max(
                0.015, 0.010 * max(load, 1.0)
            ):
                load = previous_load

        rows.append({
            "Seq#": None,
            "DateTime": t,
            "STRAIN": deformation,
            "LOAD": round(max(0.0, load), 3),
            "_test": test_no,
        })
        i += 1

    # Force the physical maximum onto the recorded observation closest to
    # the chosen peak-centre strain. This preserves the entered c/phi target.
    peak_i = min(
        range(1, len(rows)),
        key=lambda j: abs(
            rows[j]["STRAIN"] / h_mm * 100.0 - peak_pct
        ),
    )

    peak_eps = rows[peak_i]["STRAIN"] / h_mm
    peak_area = a0 / (1.0 - peak_eps)
    rows[peak_i]["LOAD"] = round(actual_target_q * peak_area, 3)

    # Smooth the local region around the peak so the forced target point
    # does not produce a sudden artificial jump. The surrounding values
    # are blended toward the continuous curve.
    for j in range(max(1, peak_i - 2), min(len(rows), peak_i + 3)):
        if j == peak_i:
            continue
        eps_j = rows[j]["STRAIN"] / h_mm
        xj = eps_j / (peak_pct / 100.0)
        qj_norm = continuous_resistance(xj, curve)
        if xj < 1.0:
            qj_norm = min(qj_norm, 0.997)
        area_j = a0 / (1.0 - eps_j)
        rows[j]["LOAD"] = round(
            actual_target_q * qj_norm * area_j, 3
        )

    # Recalculate a conservative, dynamic quality check. The check is based
    # on the neighbouring slope trend, not on a fixed load jump such as
    # "5 kg is too high". If an isolated numerical anomaly is found, rebuild
    # the test with a new random stream.
    slopes = []
    for j in range(1, len(rows)):
        dd = rows[j]["STRAIN"] - rows[j - 1]["STRAIN"]
        if dd > 0:
            slopes.append(
                (rows[j]["LOAD"] - rows[j - 1]["LOAD"]) / dd
            )

    summary = {
        "Test": test_no,
        "Cell Pressure": sigma3,
        "Peak Strain %": rows[peak_i]["STRAIN"] / h_mm * 100.0,
        "Peak Zone Start %": peak_pct * curve["peak_start"],
        "Peak Zone End %": peak_pct * curve["peak_end"],
        "End Strain %": rows[-1]["STRAIN"] / h_mm * 100.0,
        "End Time": rows[-1]["DateTime"],
        "Target q": target_q,
        "Target sigma1": target_s1,
        "Actual Peak q": actual_target_q,
        "Actual Peak sigma1": actual_target_s1,
        "Strength Factor": strength_factor,
        "Peak Load kg": rows[peak_i]["LOAD"],
        "Max Load Step kg": max(
            abs(rows[j]["LOAD"] - rows[j - 1]["LOAD"])
            for j in range(1, len(rows))
        ),
    }

    return rows, summary


def generate_all(
    c, phi, diameter, height, soil, start_dt,
    pressures, seed, interval_mm
):
    master = random.Random(seed)
    all_rows = []
    summaries = []
    current = start_dt

    for n, pressure in enumerate(pressures, 1):
        # Retry the complete physical curve if an unusual random realization
        # fails the continuity quality check. This is a safety mechanism only;
        # the curve itself is generated continuously.
        accepted = False

        for attempt in range(12):
            rng = random.Random(
                master.randint(1, 2_000_000_000)
            )
            rows, summary = generate_test(
                n, pressure, c, phi, diameter, height, soil,
                current, rng, pressures, interval_mm
            )

            # Dynamic continuity check:
            # compare each local load slope with its neighbouring slopes.
            bad = False
            if len(rows) >= 5:
                local_slopes = []
                for j in range(1, len(rows)):
                    dd = rows[j]["STRAIN"] - rows[j - 1]["STRAIN"]
                    if dd <= 0:
                        continue
                    local_slopes.append(
                        (rows[j]["LOAD"] - rows[j - 1]["LOAD"]) / dd
                    )

                # The check uses a robust percentile-like scale.
                ordered = sorted(
                    abs(v) for v in local_slopes if math.isfinite(v)
                )
                if ordered:
                    reference = ordered[
                        min(len(ordered) - 1, int(0.75 * len(ordered)))
                    ]
                    reference = max(reference, 0.001)

                    for j in range(1, len(local_slopes)):
                        if (
                            abs(local_slopes[j])
                            > reference * 4.0
                            and abs(local_slopes[j]) > 2.0
                        ):
                            bad = True
                            break

            if not bad:
                accepted = True
                break

        if not accepted:
            # Extremely unlikely; keep the final generated realization rather
            # than blocking the user.
            pass

        all_rows.extend(rows)
        summaries.append(summary)

        # Each subsequent trial begins 5-10 minutes after the previous test.
        current = summary["End Time"] + timedelta(
            minutes=master.uniform(5.0, 10.0)
        )

    for seq, row in enumerate(all_rows, 1):
        row["Seq#"] = seq

    return all_rows, summaries


def save_csv(rows,path,c=None,phi=None,soil=None):
    with open(path,"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f)
        w.writerow(["Seq#", "DateTime", "STRAIN (mm)", "LOAD (kg)"])
        for r in rows:
            w.writerow([
                r["Seq#"],
                r["DateTime"].strftime("%d-%m-%Y %H:%M:%S"),
                f'{r["STRAIN"]:.3f}',
                f'{r["LOAD"]:.3f}'
            ])

def copy_style(ws,source_row,target_row,cols):
    for col in cols:
        src,dst=ws.cell(source_row,col),ws.cell(target_row,col)
        if isinstance(dst,MergedCell): continue
        if src.has_style: dst._style=copy(src._style)
        dst.number_format=src.number_format
        dst.alignment=copy(src.alignment)
        dst.protection=copy(src.protection)

def populate_xlsm(template,output,rows,diameter,height,density,pressures,borehole=None,depth=None):
    if load_workbook is None:
        raise RuntimeError("Install openpyxl: pip install openpyxl")
    if Path(template).resolve()==Path(output).resolve():
        raise ValueError("Save the populated workbook to a new XLSM file.")

    wb=load_workbook(template,keep_vba=True,data_only=False)
    if "Input" not in wb.sheetnames: raise ValueError("Input sheet not found.")
    ws=wb["Input"]

    for cell,val in {
        "D1": borehole if borehole is not None else ws["D1"].value,
        "D2": depth if depth is not None else ws["D2"].value,
        "D5":diameter,"D7":height,"D8":density,
        "L5":diameter,"L7":height,
        "T5":diameter,"T7":height,"D10":pressures[0],
        "L10":pressures[1],"T10":pressures[2]
    }.items():
        ws[cell]=val

    blocks={1:(2,3,4,5,6,7),2:(10,11,12,13,14,15),3:(18,19,20,21,22,23)}
    data={n:[r for r in rows if r["_test"]==n] for n in (1,2,3)}

    # Clear old observation values/formulas.
    for cols in blocks.values():
        for rr in range(14,170):
            for cc in cols:
                ws.cell(rr,cc).value=None

    for n,rrs in data.items():
        if len(rrs)>156:
            raise ValueError(f"Test {n} has {len(rrs)} rows; template supports 156 rows.")
        dcol,lcol,*_=blocks[n]
        for rr,r in enumerate(rrs,14):
            ws.cell(rr,dcol).value=r["STRAIN"]/10.0   # mm -> cm
            ws.cell(rr,lcol).value=r["LOAD"]

            if n==1:
                formulas=[(4,f"=(B{rr}/$D$7)*100"),(5,f"=$D$6/(1-(D{rr}/100))"),
                          (6,f"=C{rr}/E{rr}"),(7,f"=F{rr}+$D$10")]
            elif n==2:
                formulas=[(12,f"=(J{rr}/$L$7)*100"),(13,f"=$L$6/(1-(L{rr}/100))"),
                          (14,f"=K{rr}/M{rr}"),(15,f"=N{rr}+$L$10")]
            else:
                formulas=[(20,f"=(R{rr}/$T$7)*100"),(21,f"=$T$6/(1-(T{rr}/100))"),
                          (22,f"=S{rr}/U{rr}"),(23,f"=V{rr}+$T$10")]
            for cc,formula in formulas: ws.cell(rr,cc).value=formula
            copy_style(ws,14,rr,list(blocks[n]))

    try:
        wb.calculation.fullCalcOnLoad=True
        wb.calculation.forceFullCalc=True
        wb.calculation.calcMode="auto"
    except Exception:
        pass
    wb.save(output)

def calculate_final_mohr_coulomb_envelope(sigma3_values, sigma1_values):
    """
    Fit the final Mohr-Coulomb envelope from the generated UU test results.

    For each test:
        sigma_n = (sigma1 + sigma3) / 2
        tau_max  = (sigma1 - sigma3) / 2

    A least-squares straight-line fit:
        tau = Cu + sigma_n * tan(phi)

    is then calculated from the actual generated failure points.
    """
    if len(sigma3_values) != len(sigma1_values) or len(sigma3_values) < 2:
        raise ValueError("At least two generated tests are required.")

    sigma_n = [
        (s1 + s3) / 2.0
        for s1, s3 in zip(sigma1_values, sigma3_values)
    ]
    tau = [
        (s1 - s3) / 2.0
        for s1, s3 in zip(sigma1_values, sigma3_values)
    ]

    xbar = sum(sigma_n) / len(sigma_n)
    ybar = sum(tau) / len(tau)

    den = sum((x - xbar) ** 2 for x in sigma_n)
    slope = (
        sum((x - xbar) * (y - ybar) for x, y in zip(sigma_n, tau)) / den
        if den > 1e-12 else 0.0
    )

    intercept = ybar - slope * xbar

    # Physical reporting limits.
    intercept = max(0.0, intercept)
    slope = max(0.0, slope)

    phi_deg = math.degrees(math.atan(slope))

    # R² is retained internally for fit assessment, but the report UI
    # should not display it, as requested earlier.
    yhat = [intercept + slope * x for x in sigma_n]
    ss_res = sum((y - yh) ** 2 for y, yh in zip(tau, yhat))
    ss_tot = sum((y - ybar) ** 2 for y in tau)
    r2 = 1.0 - ss_res / ss_tot if ss_tot > 1e-12 else 1.0

    return {
        "Cu": intercept,
        "phi": phi_deg,
        "sigma_n": sigma_n,
        "tau": tau,
        "r2": r2,
    }


def build_final_mohr_plot(ax, sigma3_values, sigma1_values, test_labels=None):
    """
    Draw the final Mohr plot using ACTUAL generated peak strengths.

    Only the upper semicircles are shown. The circles are based on generated results, while the
    displayed failure envelope is the best-fit line through the generated
    failure points. This deliberately differs from the preliminary
    theoretical envelope when realistic test-to-test variation is present.
    """
    if test_labels is None:
        test_labels = [f"Test {i+1}" for i in range(len(sigma3_values))]

    fit = calculate_final_mohr_coulomb_envelope(
        sigma3_values, sigma1_values
    )

    # Three distinct plot colours are intentionally obtained from the
    # matplotlib default cycle rather than hard-coded colours.
    theta = [i * math.pi / 360.0 for i in range(721)]

    all_x = []
    all_y = []

    for i, (s3, s1) in enumerate(zip(sigma3_values, sigma1_values)):
        center = (s1 + s3) / 2.0
        radius = (s1 - s3) / 2.0

        xs = [
            center + radius * math.cos(t)
            for t in theta
        ]
        ys = [
            radius * math.sin(t)
            for t in theta
        ]

        all_x.extend(xs)
        all_y.extend(ys)

        ax.plot(xs, ys, linewidth=2.0, label=test_labels[i])
        ax.plot(
            center + radius,
            0.0,
            marker="o",
            markersize=6
        )

        # Failure point in Mohr coordinates.
        ax.plot(
            center,
            radius,
            marker="o",
            markersize=6,
            color="black"
        )

        ax.plot(
            [center, center],
            [0, radius],
            linestyle="--",
            linewidth=0.9,
            color="black"
        )

    # Final fitted envelope from actual generated failure points.
    xmax = max(all_x) if all_x else 1.0
    xline = [0.0, xmax * 1.03]
    yline = [
        fit["Cu"] + x * math.tan(math.radians(fit["phi"]))
        for x in xline
    ]

    ax.plot(
        xline,
        yline,
        linewidth=2.2,
        label="Mohr–Coulomb Failure Envelope"
    )

    ax.set_xlabel("Normal Stress, σ (kg/cm²)")
    ax.set_ylabel("Shear Stress, τ (kg/cm²)")

    # Mohr's circles are shown as upper semicircles only.
    # Keep the horizontal axis at τ = 0 and force the vertical axis to
    # coincide with σ = 0 (no negative x-axis portion).
    xmin = 0.0
    xmax_plot = max(all_x) * 1.55 if all_x else 1.0
    ymax_plot = max(all_y) * 1.90 if all_y else 1.0
    ymax_plot = max(ymax_plot, 0.10)

    ax.set_xlim(xmin, xmax_plot)
    ax.set_ylim(0.0, ymax_plot)

    # Explicitly place the axes at σ = 0 and τ = 0.
    ax.spines["left"].set_position(("data", 0.0))
    ax.spines["bottom"].set_position(("data", 0.0))
    ax.spines["right"].set_visible(False)
    ax.spines["top"].set_visible(False)

    # Keep ticks on the actual axes.
    ax.xaxis.set_ticks_position("bottom")
    ax.yaxis.set_ticks_position("left")

    ax.grid(True, linestyle="--", alpha=0.35)

    return fit



def build_report_mohr_figure(
    borehole, depth_item, summaries, sigma3_values, sigma1_values
):
    """
    Build the final report-style Mohr's circle sheet matching the approved
    reference layout:
      - report title
      - Borehole/Sample/Condition/Depth box
      - Test Summary box
      - Mohr plot
      - Test Results table
      - Mohr-Coulomb Parameters table
      - Observations box

    The plotted circles and fitted envelope use the ACTUAL generated peak
    stresses, so the reported Cu and phi are the same values used by the
    final generated data.
    """
    fit = calculate_final_mohr_coulomb_envelope(
        sigma3_values, sigma1_values
    )

    fig = Figure(figsize=(15.36, 10.24), dpi=100, facecolor="white")

    # ---------------- Overall title ----------------
    fig.text(
        0.5, 0.975,
        "UU TRIAXIAL TEST – MOHR'S CIRCLE PLOT",
        ha="center", va="top",
        fontsize=21, fontweight="bold"
    )

    # ---------------- Top information boxes ----------------
    top = fig.add_axes([0.02, 0.765, 0.96, 0.155])
    top.set_xlim(0, 1)
    top.set_ylim(0, 1)
    top.axis("off")

    # Outer box and vertical divider.
    top.plot([0, 1, 1, 0, 0], [0, 0, 1, 1, 0],
             color="black", linewidth=1.0)
    top.plot([0.54, 0.54], [0, 1], color="black", linewidth=1.0)

    sample_type = str(depth_item.get("Sample Type", ""))
    condition_map = {
        "UDS": "Undisturbed",
        "SPT": "Disturbed",
        "DS": "Remoulded",
    }
    sample_condition = condition_map.get(sample_type, "")

    left_lines = [
        ("Borehole ID", str(borehole)),
        ("Sample Type", sample_type),
        ("Sample Condition", sample_condition),
        ("Depth", f"{depth_item['Depth']:.2f} m"),
    ]
    for j, (lab, val) in enumerate(left_lines):
        y = 0.78 - j * 0.24
        top.text(0.012, y, lab, fontsize=10.5, va="center")
        top.text(0.135, y, ":", fontsize=10.5, va="center")
        top.text(0.155, y, val, fontsize=10.5, va="center")

    top.text(0.557, 0.79, "Test Summary",
             fontsize=11, fontweight="bold", va="center")
    top.plot([0.557, 0.655], [0.755, 0.755],
             color="black", linewidth=0.8)

    top.text(0.557, 0.50, "No. of Tests", fontsize=10.5, va="center")
    top.text(0.655, 0.50, ":", fontsize=10.5, va="center")
    top.text(0.675, 0.50, "3", fontsize=10.5, va="center")

    top.text(0.557, 0.27, "Test Type", fontsize=10.5, va="center")
    top.text(0.655, 0.27, ":", fontsize=10.5, va="center")
    top.text(0.675, 0.27, "UU Triaxial Compression Test",
             fontsize=10.5, va="center")

    # ---------------- Main plot ----------------
    ax = fig.add_axes([0.065, 0.125, 0.59, 0.605])
    test_labels = [
        f"Test {i+1} (σ₃ = {p:.2f} kg/cm²)"
        for i, p in enumerate(sigma3_values)
    ]
    build_final_mohr_plot(ax, sigma3_values, sigma1_values, test_labels)

    # Match approved report style: concise legend.
    handles, labels = ax.get_legend_handles_labels()
    wanted = [
        i for i, label in enumerate(labels)
        if label.startswith("Test ") or "Failure Envelope" in label
    ]
    ax.legend(
        [handles[i] for i in wanted],
        [labels[i] for i in wanted],
        fontsize=8.5,
        loc="upper left",
        frameon=True
    )

    # Cu intercept annotation.
    xmax = ax.get_xlim()[1]
    ymax = ax.get_ylim()[1]
    cu = fit["Cu"]
    phi = fit["phi"]

    if 0 < cu < ymax * 0.85:
        ax.annotate(
            f"Cu = {cu:.2f} kg/cm²",
            xy=(0, cu),
            xytext=(0.055 * xmax, cu),
            fontsize=8.5,
            color="black",
            va="center",
            arrowprops=dict(
                arrowstyle="-",
                linewidth=0.8
            )
        )

    # Small phi angle construction at the failure envelope.
    x_phi = xmax * 0.73
    y_phi = cu + x_phi * math.tan(math.radians(phi))
    dx = xmax * 0.085
    dy = dx * math.tan(math.radians(phi))
    if y_phi < ymax * 0.90:
        ax.plot(
            [x_phi, x_phi + dx],
            [y_phi, y_phi],
            color="black", linewidth=0.8
        )
        ax.plot(
            [x_phi, x_phi + dx],
            [y_phi, y_phi + dy],
            color="black", linewidth=0.8
        )
        arc_radius = dx * 0.60
        ax.add_patch(
            Arc(
                (x_phi, y_phi),
                2 * arc_radius, 2 * arc_radius,
                angle=0, theta1=0, theta2=phi,
                linewidth=0.8, color="black"
            )
        )
        ax.text(
            x_phi + dx * 0.62,
            y_phi + max(dy * 0.38, ymax * 0.018),
            f"φ = {phi:.2f}°",
            fontsize=8.5,
            va="center"
        )

    # ---------------- Right-hand report tables ----------------
    # Test Results
    tr = fig.add_axes([0.675, 0.515, 0.305, 0.215])
    tr.axis("off")
    tr.text(0.5, 0.965, "Test Results",
            ha="center", va="top", fontsize=12, fontweight="bold")
    tr.add_patch(
        __import__("matplotlib").patches.Rectangle(
            (0, 0.86), 1, 0.14, fill=False, linewidth=0.8
        )
    )

    tr_cols = [
        "Test\nNo.",
        "Cell Pressure\nσ₃ (kg/cm²)",
        "Deviator Stress\nat Failure\n(σ₁−σ₃) (kg/cm²)",
        "Major Principal\nStress at Failure\nσ₁ (kg/cm²)",
        "Shear Stress\nat Failure\nτf (kg/cm²)"
    ]
    tr_data = []
    for i, s in enumerate(summaries):
        s3 = s["Cell Pressure"]
        s1 = s["Actual Peak sigma1"]
        q = s1 - s3
        tau = q / 2.0
        tr_data.append([
            str(i + 1),
            f"{s3:.2f}",
            f"{q:.2f}",
            f"{s1:.2f}",
            f"{tau:.2f}",
        ])

    table = tr.table(
        cellText=tr_data,
        colLabels=tr_cols,
        cellLoc="center",
        colLoc="center",
        bbox=[0, 0, 1, 0.86],
        colWidths=[0.10, 0.19, 0.25, 0.25, 0.21]
    )
    table.auto_set_font_size(False)
    table.set_fontsize(7.6)
    for cell in table.get_celld().values():
        cell.set_edgecolor("black")
        cell.set_linewidth(0.7)
    for c in range(len(tr_cols)):
        table[(0, c)].set_text_props(weight="bold")

    # Mohr-Coulomb parameters
    mp = fig.add_axes([0.675, 0.335, 0.305, 0.145])
    mp.axis("off")
    mp.text(0.5, 0.965, "Mohr-Coulomb Parameters",
            ha="center", va="top", fontsize=12, fontweight="bold")
    mp.add_patch(
        __import__("matplotlib").patches.Rectangle(
            (0, 0.84), 1, 0.16, fill=False, linewidth=0.8
        )
    )

    mp_data = [
        ["Undrain cohesion", "Cᵤ", f"{cu:.2f}", "kg/cm²"],
        ["Angle of Internal Friction", "φ", f"{phi:.2f}", "°"],
    ]
    mp_table = mp.table(
        cellText=mp_data,
        colLabels=["Parameter", "Symbol", "Value", "Unit"],
        cellLoc="center",
        colLoc="center",
        bbox=[0, 0, 1, 0.84],
        colWidths=[0.49, 0.15, 0.18, 0.18]
    )
    mp_table.auto_set_font_size(False)
    mp_table.set_fontsize(8.0)
    for cell in mp_table.get_celld().values():
        cell.set_edgecolor("black")
        cell.set_linewidth(0.7)
    for c in range(4):
        mp_table[(0, c)].set_text_props(weight="bold")

    # Observations box
    ob = fig.add_axes([0.675, 0.125, 0.305, 0.185])
    ob.set_xlim(0, 1)
    ob.set_ylim(0, 1)
    ob.axis("off")
    ob.plot([0, 1, 1, 0, 0], [0, 0, 1, 1, 0],
            color="black", linewidth=0.8)
    ob.text(0.045, 0.90, "Observations",
            fontsize=11, fontweight="bold", va="center")
    ob.plot([0, 1], [0.80, 0.80], color="black", linewidth=0.7)
    ob.text(
        0.045, 0.64,
        "• Failure envelope drawn as per Mohr–Coulomb criteria.",
        fontsize=8.5, va="center"
    )

    return fig, fit


def preliminary_mohr_parameters(cu, phi, sigma3_values):
    """
    Return theoretical σ1 values for the preliminary envelope preview.
    """
    tan_phi = math.tan(math.radians(phi))
    sigma1_values = []

    for s3 in sigma3_values:
        sigma1 = (
            s3 * (1.0 + tan_phi)
            + 2.0 * cu * math.sqrt(1.0 + tan_phi ** 2)
        )
        sigma1_values.append(sigma1)

    return sigma1_values


class App:
    """Multi-depth UU triaxial GUI with realistic Mohr-circle preview."""

    def __init__(self, root):
        self.root = root
        self.root.title("AI-Assisted UU Triaxial Data Generator - Multi Depth")
        self.root.geometry("1380x900")
        self.root.minsize(1120, 760)
        self.rows = None
        self.summaries = None
        self.preview_item = None
        self.preview_canvas = None

        main = ttk.Frame(root, padding=10)
        main.pack(fill="both", expand=True)

        ttk.Label(
            main,
            text="AI-Assisted UU Triaxial Synthetic Data Generator",
            font=("Segoe UI", 16, "bold")
        ).pack(anchor="w", pady=(0, 6))

        # ---------------- Common inputs ----------------
        common = ttk.LabelFrame(main, text="Common Test Settings", padding=8)
        common.pack(fill="x", pady=(0, 7))
        self.common = {}
        common_fields = [
            ("Borehole ID", "BH-01"),
            ("Initial Date/Time (dd-mm-yyyy HH:MM:SS)", "05-08-2026 15:47:00"),
            ("Nominal Data Recording Interval (mm)", "0.10"),
            ("Random Seed", "20260816"),
        ]
        for i, (label, value) in enumerate(common_fields):
            col = (i % 2) * 2
            row = i // 2
            ttk.Label(common, text=label).grid(
                row=row, column=col, sticky="w", padx=5, pady=3
            )
            var = tk.StringVar(value=value)
            self.common[label] = var
            ttk.Entry(common, textvariable=var, width=28).grid(
                row=row, column=col + 1, sticky="w", padx=5, pady=3
            )

        # ---------------- Multiple-depth input ----------------
        depth_box = ttk.LabelFrame(
            main,
            text="Multiple Depth Input  |  Diameter & Height are entered separately for each depth",
            padding=7
        )
        depth_box.pack(fill="x", pady=(0, 7))

        ttk.Label(
            depth_box,
            text=(
                "Each depth is an independent specimen set. Enter diameter, height, "
                "sample type, soil, Cu, φ and three cell pressures."
            )
        ).pack(anchor="w", pady=(0, 5))

        table_wrap = ttk.Frame(depth_box)
        table_wrap.pack(fill="x")
        self.depth_canvas = tk.Canvas(table_wrap, height=245, highlightthickness=0)
        self.depth_canvas.pack(side="left", fill="x", expand=True)
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.depth_canvas.yview)
        ysb.pack(side="right", fill="y")
        self.depth_canvas.configure(yscrollcommand=ysb.set)
        self.depth_frame = ttk.Frame(self.depth_canvas)
        self.depth_window = self.depth_canvas.create_window(
            (0, 0), window=self.depth_frame, anchor="nw"
        )
        self.depth_frame.bind(
            "<Configure>",
            lambda e: self.depth_canvas.configure(scrollregion=self.depth_canvas.bbox("all"))
        )
        self.depth_canvas.bind(
            "<Configure>",
            lambda e: self.depth_canvas.itemconfigure(self.depth_window, width=e.width)
        )

        headers = [
            "Depth (m)", "Diameter (cm)", "Height (cm)", "Density (gm/cc)",
            "Sample Type", "Soil Type", "Cu (kg/cm²)", "φ (°)",
            "Cell Pressure 1", "Cell Pressure 2", "Cell Pressure 3",
            "Select", "Preview"
        ]
        for c, h in enumerate(headers):
            ttk.Label(
                self.depth_frame, text=h, font=("Segoe UI", 9, "bold")
            ).grid(row=0, column=c, padx=2, pady=3, sticky="ew")
            self.depth_frame.columnconfigure(c, weight=1)
        self.depth_frame.columnconfigure(11, weight=0)
        self.depth_frame.columnconfigure(12, weight=0)

        self.depth_rows = []
        self.add_depth_row(
            5.00, "3.80", "7.64", "1.80", "UDS", "Medium Clay", "0.50", "10.00",
            "0.50", "1.00", "1.50"
        )
        self.add_depth_row(
            12.00, "3.80", "7.64", "1.85", "UDS", "Stiff Clay", "0.80", "8.00",
            "1.00", "2.00", "3.00"
        )

        buttons = ttk.Frame(depth_box)
        buttons.pack(fill="x", pady=(6, 0))
        ttk.Button(buttons, text="+ Add Depth", command=self.add_depth_row).pack(side="left", padx=3)
        ttk.Button(buttons, text="Remove Selected Depth", command=self.remove_depth_row).pack(side="left", padx=3)
        ttk.Button(buttons, text="Clear All", command=self.clear_depth_rows).pack(side="left", padx=3)
        ttk.Button(buttons, text="Add 5 Depth Rows", command=lambda: self.add_n_rows(5)).pack(side="left", padx=3)
        ttk.Button(buttons, text="Preview Selected", command=self.preview_selected).pack(side="left", padx=15)
        ttk.Button(buttons, text="Import Excel Input", command=self.import_excel_input).pack(side="left", padx=15)

        # ---------------- Main actions (kept above the large preview area) ----------------
        bar = ttk.Frame(main)
        bar.pack(fill="x", pady=(7, 4))
        ttk.Button(bar, text="Generate All Data", command=self.generate).pack(side="left", padx=3)
        ttk.Button(bar, text="Final Mohr Plot", command=self.final_mohr_plot).pack(side="left", padx=3)
        ttk.Button(bar, text="Export All Mohr Plots", command=self.export_all_mohr_plots).pack(side="left", padx=3)
        ttk.Button(bar, text="Save CSV", command=self.csv).pack(side="left", padx=3)
        ttk.Button(bar, text="Populate XLSM Template", command=self.xlsm).pack(side="left", padx=3)
        ttk.Button(bar, text="All Export", command=self.export_all_package).pack(side="left", padx=3)

        # ---------------- Preview + data area ----------------
        lower = ttk.Frame(main)
        lower.pack(fill="both", expand=True)
        lower.columnconfigure(0, weight=1)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)

        plot_fr = ttk.LabelFrame(
            lower,
            text="Mohr Failure Envelope Preview  |  Realistic generated-test preview",
            padding=5
        )
        plot_fr.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        if MATPLOTLIB_AVAILABLE:
            self.fig = Figure(figsize=(6.6, 5.0), dpi=100)
            self.ax = self.fig.add_subplot(111)
            self.preview_canvas = FigureCanvasTkAgg(self.fig, master=plot_fr)
            self.preview_canvas.get_tk_widget().pack(fill="both", expand=True)
        else:
            self.fig = None
            self.ax = None
            ttk.Label(
                plot_fr,
                text="Matplotlib is required for the Mohr-circle preview.\n"
                     "Install it with: py -m pip install matplotlib"
            ).pack(expand=True)

        data_fr = ttk.LabelFrame(lower, text="Preview - First 200 Observations", padding=5)
        data_fr.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        tree_wrap = ttk.Frame(data_fr)
        tree_wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_wrap, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)
        ytree = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        ytree.pack(side="right", fill="y")
        xtree = ttk.Scrollbar(data_fr, orient="horizontal", command=self.tree.xview)
        xtree.pack(side="bottom", fill="x")
        self.tree.configure(yscrollcommand=ytree.set, xscrollcommand=xtree.set)

        # ---------------- Status ----------------
        self.status = tk.StringVar(value="Ready. Edit depths and use Preview Selected to inspect a realistic generated-test envelope.")
        ttk.Label(main, textvariable=self.status).pack(anchor="w")

        # Show the first depth immediately.
        if self.depth_rows:
            self.preview_item = self.depth_rows[0]
            self.draw_preliminary_preview(self.preview_item)

    # ---------------- Depth-row management ----------------
    def add_depth_row(
        self, depth="", diameter="3.80", height="7.64", density="1.80",
        sample_type="UDS", soil_type="Medium Clay", c="0.50", phi="10.00",
        p1="0.50", p2="1.00", p3="1.50"
    ):
        r = len(self.depth_rows) + 1
        values = [
            str(depth), str(diameter), str(height), str(density),
            str(sample_type), str(soil_type), str(c), str(phi),
            str(p1), str(p2), str(p3)
        ]
        vars_ = [tk.StringVar(value=v) for v in values]
        widgets = []

        for col, var in enumerate(vars_):
            if col == 4:
                w = ttk.Combobox(
                    self.depth_frame, textvariable=var,
                    values=("UDS", "SPT"), state="readonly", width=10
                )
            elif col == 5:
                w = ttk.Combobox(
                    self.depth_frame, textvariable=var,
                    values=list(SOIL), state="readonly", width=19
                )
            else:
                w = ttk.Entry(self.depth_frame, textvariable=var, width=13)
            w.grid(row=r, column=col, padx=2, pady=2, sticky="ew")
            widgets.append(w)

        selected = tk.BooleanVar(value=False)
        check = ttk.Checkbutton(self.depth_frame, variable=selected)
        check.grid(row=r, column=11, padx=2, pady=2)
        widgets.append(check)

        item = {"vars": vars_, "widgets": widgets, "selected": selected}
        preview_button = ttk.Button(
            self.depth_frame, text="Preview", width=9,
            command=lambda x=item: self.preview_item_and_draw(x)
        )
        preview_button.grid(row=r, column=12, padx=2, pady=2)
        item["preview_button"] = preview_button
        self.depth_rows.append(item)

        self.depth_frame.update_idletasks()
        self.depth_canvas.configure(scrollregion=self.depth_canvas.bbox("all"))
        self.depth_canvas.yview_moveto(1.0)

    def preview_item_and_draw(self, item):
        self.preview_item = item
        self.draw_preliminary_preview(item)

    def preview_selected(self):
        selected = [x for x in self.depth_rows if x["selected"].get()]
        if not selected:
            if not self.depth_rows:
                return
            messagebox.showwarning("Preview", "Tick the checkbox of a depth, then click Preview Selected.")
            return
        self.preview_item_and_draw(selected[0])

    def add_n_rows(self, n):
        for _ in range(n):
            self.add_depth_row()

    def remove_depth_row(self):
        selected = [x for x in self.depth_rows if x["selected"].get()]
        if not selected:
            messagebox.showwarning("Remove Depth", "Tick the checkbox at the right of the depth row(s) to remove.")
            return
        for item in selected:
            for w in item["widgets"]:
                w.destroy()
            item["preview_button"].destroy()
        self.depth_rows = [x for x in self.depth_rows if x not in selected]
        self.regrid_depth_rows()
        if self.depth_rows:
            self.preview_item_and_draw(self.depth_rows[0])
        else:
            self.clear_plot()

    def regrid_depth_rows(self):
        for r, item in enumerate(self.depth_rows, start=1):
            for col, w in enumerate(item["widgets"][:-1]):
                w.grid_configure(row=r, column=col)
            item["widgets"][-1].grid_configure(row=r, column=11)
            item["preview_button"].grid_configure(row=r, column=12)
            item["selected"].set(False)
        self.depth_frame.update_idletasks()
        self.depth_canvas.configure(scrollregion=self.depth_canvas.bbox("all"))

    def _clear_depth_rows_without_prompt(self):
        """Clear current depth rows without confirmation."""
        for item in self.depth_rows:
            for w in item["widgets"]:
                w.destroy()
            item["preview_button"].destroy()
        self.depth_rows = []
        self.depth_frame.update_idletasks()
        self.depth_canvas.configure(scrollregion=self.depth_canvas.bbox("all"))

    def import_excel_input(self):
        """Import common settings and all depth rows from an Excel input workbook."""
        if load_workbook is None:
            messagebox.showerror(
                "Excel Import",
                "openpyxl is required for Excel import.\n"
                "Install it with: py -m pip install openpyxl"
            )
            return

        path = filedialog.askopenfilename(
            title="Select Excel Input File",
            filetypes=[("Excel Input Workbook", "*.xlsx"),
                       ("Excel Workbook", "*.xlsm")]
        )
        if not path:
            return

        wb = None
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
            if "Input" not in wb.sheetnames:
                raise ValueError("The Excel input file must contain an 'Input' sheet.")
            ws = wb["Input"]

            borehole = ws["B4"].value
            start_dt = ws["B5"].value
            interval = ws["B6"].value
            seed = ws["B7"].value

            if borehole is None or str(borehole).strip() == "":
                raise ValueError("Borehole ID cannot be blank in the Excel input.")
            if start_dt is None:
                raise ValueError("Initial Date/Time cannot be blank in the Excel input.")

            if isinstance(start_dt, datetime):
                start_text = start_dt.strftime("%d-%m-%Y %H:%M:%S")
            else:
                start_text = str(start_dt).strip()
                parsed_dt = None
                for fmt in ("%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
                            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
                    try:
                        parsed_dt = datetime.strptime(start_text, fmt)
                        break
                    except ValueError:
                        pass
                if parsed_dt is None:
                    raise ValueError(
                        "Initial Date/Time must be in dd-mm-yyyy HH:MM:SS format."
                    )
                start_text = parsed_dt.strftime("%d-%m-%Y %H:%M:%S")

            self.common["Borehole ID"].set(str(borehole).strip())
            self.common["Initial Date/Time (dd-mm-yyyy HH:MM:SS)"].set(start_text)
            self.common["Nominal Data Recording Interval (mm)"].set(str(interval))
            self.common["Random Seed"].set(str(int(seed)))

            imported = []
            for row_num in range(11, ws.max_row + 1):
                vals = [ws.cell(row_num, c).value for c in range(1, 12)]
                if all(v is None or str(v).strip() == "" for v in vals):
                    continue
                if vals[0] is None or str(vals[0]).strip() == "":
                    continue
                imported.append(tuple(
                    str(v) if v is not None else "" for v in vals
                ))

            if not imported:
                raise ValueError("No depth rows were found in the Excel input.")

            self._clear_depth_rows_without_prompt()
            for row in imported:
                self.add_depth_row(*row)

            self.preview_item = self.depth_rows[0]
            self.draw_preliminary_preview(self.preview_item)
            self.rows = None
            self.summaries = None
            self.status.set(
                f"Imported {len(imported)} depth(s) from Excel input. "
                "Click Generate All Data."
            )
            messagebox.showinfo(
                "Excel Import Complete",
                f"{len(imported)} depth row(s) imported successfully.\n\n"
                "The imported values are now loaded into the generator."
            )
        except Exception as e:
            messagebox.showerror("Excel Import Error", str(e))
        finally:
            if wb is not None:
                wb.close()

    def clear_depth_rows(self):
        if not self.depth_rows:
            return
        if not messagebox.askyesno("Clear All", "Remove all depth rows?"):
            return
        for item in self.depth_rows:
            for w in item["widgets"]:
                w.destroy()
            item["preview_button"].destroy()
        self.depth_rows = []
        self.clear_plot()

    # ---------------- Preliminary Mohr preview ----------------
    def clear_plot(self):
        if self.ax is not None:
            self.ax.clear()
            self.ax.text(0.5, 0.5, "No depth selected", ha="center", va="center", transform=self.ax.transAxes)
            self.ax.set_axis_off()
            self.preview_canvas.draw_idle()

    def draw_preliminary_preview(self, item):
        """
        Preview the same kind of realistic Mohr plot that will result after
        generation. It is not a purely theoretical/tangent plot.

        The three preview tests are generated with the same physical generator
        and specimen-strength variation used by generate_test(), using a
        deterministic seed derived from the current inputs.
        """
        if not MATPLOTLIB_AVAILABLE or self.ax is None:
            return

        try:
            v = item["vars"]
            depth = float(v[0].get())
            diameter = float(v[1].get())
            height = float(v[2].get())
            density = float(v[3].get())
            sample_type = v[4].get()
            soil = v[5].get()
            c = float(v[6].get())
            phi = float(v[7].get())
            pressures = [
                float(v[8].get()),
                float(v[9].get()),
                float(v[10].get())
            ]
            interval_mm = float(
                self.common["Nominal Data Recording Interval (mm)"].get()
            )
            base_seed = int(self.common["Random Seed"].get())

            if (
                depth < 0 or diameter <= 0 or height <= 0 or density <= 0 or c < 0
                or any(p < 0 for p in pressures)
                or phi < 0 or phi >= 45 or interval_mm <= 0
                or sample_type not in ("UDS", "SPT") or soil not in SOIL
            ):
                raise ValueError

        except Exception:
            self.ax.clear()
            self.ax.text(
                0.5, 0.5,
                "Enter valid values to preview",
                ha="center", va="center",
                transform=self.ax.transAxes
            )
            self.ax.set_axis_off()
            self.preview_canvas.draw_idle()
            return

        # Deterministic preview: unchanged inputs give the same realistic
        # specimen-to-specimen variation each time.
        preview_seed = (
            base_seed
            + int(round(depth * 1000))
            + int(round(diameter * 1000))
            + int(round(height * 1000))
            + int(round(density * 1000))
            + int(round(c * 10000))
            + int(round(phi * 100))
        )
        master = random.Random(preview_seed)

        sigma3_values = []
        sigma1_values = []
        strength_factors = []

        start = datetime.strptime(
            self.common[
                "Initial Date/Time (dd-mm-yyyy HH:MM:SS)"
            ].get().strip(),
            "%d-%m-%Y %H:%M:%S"
        )

        for test_no, pressure in enumerate(pressures, start=1):
            rng = random.Random(master.randint(1, 2_000_000_000))

            # Use the exact same generator used for final data.
            _, summary = generate_test(
                test_no,
                pressure,
                c,
                phi,
                diameter,
                height,
                soil,
                start,
                rng,
                pressures,
                interval_mm
            )

            sigma3_values.append(summary["Cell Pressure"])
            sigma1_values.append(summary["Actual Peak sigma1"])
            strength_factors.append(summary["Strength Factor"])

        self.ax.clear()

        # Actual generated peak stresses are used to draw the circles and
        # fit the envelope. This is intentionally different from the old
        # theoretical/tangent preview.
        fit = build_final_mohr_plot(
            self.ax,
            sigma3_values,
            sigma1_values,
            [
                f"Test {i+1} (σ₃ = {p:.2f})"
                for i, p in enumerate(sigma3_values)
            ]
        )

        self.ax.set_title(
            f"Realistic Mohr Failure Envelope Preview | "
            f"Depth = {depth:.2f} m | {soil}",
            fontsize=11,
            fontweight="bold"
        )

        # Only the three test circles and the fitted failure envelope.
        handles, labels = self.ax.get_legend_handles_labels()
        wanted = [
            i for i, label in enumerate(labels)
            if label.startswith("Test ")
            or "Failure Envelope" in label
        ]
        self.ax.legend(
            [handles[i] for i in wanted],
            [labels[i] for i in wanted],
            fontsize=8,
            loc="upper left"
        )

        self.ax.text(
            0.98, 0.96,
            f"Preview fitted Cu = {fit['Cu']:.2f} kg/cm²\n"
            f"Preview fitted φ = {fit['phi']:.2f}°\n"
            f"Input Cu = {c:.2f} kg/cm²\n"
            f"Input φ = {phi:.2f}°\n"
            f"D = {diameter:.2f} cm\n"
            f"H = {height:.2f} cm\n"
            f"Density = {density:.2f} gm/cc",
            transform=self.ax.transAxes,
            ha="right",
            va="top",
            fontsize=8,
            bbox=dict(
                boxstyle="round,pad=0.3",
                facecolor="white",
                alpha=0.88
            )
        )

        self.ax.set_xlabel("Normal Stress, σ (kg/cm²)")
        self.ax.set_ylabel("Shear Stress, τ (kg/cm²)")
        self.ax.grid(True, linestyle="--", alpha=0.35)

        self.preview_canvas.draw_idle()

        spread = max(strength_factors) - min(strength_factors)
        self.status.set(
            f"Realistic preview: Depth {depth:.2f} m | "
            f"specimen strength variation up to ±20% | "
            f"fitted Cu={fit['Cu']:.2f} kg/cm², "
            f"φ={fit['phi']:.2f}° | "
            f"strength spread={spread*100:.1f}%"
        )

    # ---------------- Input collection ----------------
    def inputs(self):
        def f(value, label):
            try:
                return float(value.strip())
            except Exception:
                raise ValueError(f"Invalid numeric value for {label}.")

        borehole = self.common["Borehole ID"].get().strip()
        if not borehole:
            raise ValueError("Borehole ID cannot be blank.")
        interval_mm = f(self.common["Nominal Data Recording Interval (mm)"].get(), "Nominal Data Recording Interval")
        if interval_mm < 0.02 or interval_mm > 1.0:
            raise ValueError("Nominal data recording interval should be between 0.02 and 1.00 mm.")
        try:
            seed = int(self.common["Random Seed"].get().strip())
        except Exception:
            raise ValueError("Random Seed must be an integer.")
        try:
            dt = datetime.strptime(
                self.common["Initial Date/Time (dd-mm-yyyy HH:MM:SS)"].get().strip(),
                "%d-%m-%Y %H:%M:%S"
            )
        except ValueError:
            raise ValueError("Date/Time must be: 05-08-2026 15:47:00")
        if not self.depth_rows:
            raise ValueError("Add at least one depth.")

        depths = []
        for i, item in enumerate(self.depth_rows, start=1):
            v = item["vars"]
            try:
                depth = float(v[0].get())
                diameter = float(v[1].get())
                height = float(v[2].get())
                density = float(v[3].get())
                c = float(v[6].get())
                phi = float(v[7].get())
                pressures = [float(v[8].get()), float(v[9].get()), float(v[10].get())]
            except Exception:
                raise ValueError(f"Invalid numeric input in depth row {i}.")
            sample_type = v[4].get().strip()
            soil = v[5].get().strip()
            if depth < 0:
                raise ValueError(f"Depth in row {i} cannot be negative.")
            if diameter <= 0 or height <= 0:
                raise ValueError(f"Diameter and height in depth row {i} must be > 0.")
            if density <= 0:
                raise ValueError(f"Density in depth row {i} must be > 0 gm/cc.")
            if c < 0:
                raise ValueError(f"Cu in depth row {i} must be >= 0.")
            if phi < 0 or phi >= 45:
                raise ValueError(f"φ in depth row {i} must be between 0 and 45 degrees.")
            if any(p < 0 for p in pressures):
                raise ValueError(f"Cell pressure in depth row {i} cannot be negative.")
            if soil not in SOIL:
                raise ValueError(f"Invalid soil type in depth row {i}.")
            if sample_type not in ("UDS", "SPT"):
                raise ValueError(f"Sample Type in depth row {i} must be UDS or SPT.")
            depths.append({
                "Depth": depth, "Diameter": diameter, "Height": height,
                "Density": density, "Sample Type": sample_type, "Soil": soil,
                "c": c, "phi": phi, "pressures": pressures,
            })
        return {
            "borehole": borehole, "start_dt": dt,
            "interval_mm": interval_mm, "seed": seed, "depths": depths,
        }

    # ---------------- Generation ----------------
    def generate(self):
        try:
            inp = self.inputs()
            master = random.Random(inp["seed"])
            all_rows, all_summaries = [], []
            current = inp["start_dt"]
            test_counter = 0

            for depth_no, item in enumerate(inp["depths"], start=1):
                pressures = item["pressures"]

                # IMPORTANT: use the exact same deterministic seed sequence as
                # draw_preliminary_preview(). This guarantees that the preview
                # failure points, fitted Cu and fitted phi are identical to the
                # subsequently generated/exported final Mohr plot.
                preview_seed = (
                    inp["seed"]
                    + int(round(item["Depth"] * 1000))
                    + int(round(item["Diameter"] * 1000))
                    + int(round(item["Height"] * 1000))
                    + int(round(item["Density"] * 1000))
                    + int(round(item["c"] * 10000))
                    + int(round(item["phi"] * 100))
                )
                local_master = random.Random(preview_seed)

                for local_n, pressure in enumerate(pressures, start=1):
                    test_counter += 1
                    rng = random.Random(local_master.randint(1, 2_000_000_000))
                    rows, summary = generate_test(
                        local_n, pressure, item["c"], item["phi"],
                        item["Diameter"], item["Height"], item["Soil"],
                        current, rng, pressures, inp["interval_mm"]
                    )

                    for r in rows:
                        r["_depth_no"] = depth_no
                        r["_depth"] = item["Depth"]
                        r["_sample_type"] = item["Sample Type"]
                        r["_soil"] = item["Soil"]
                        r["_diameter"] = item["Diameter"]
                        r["_height"] = item["Height"]
                        r["_density"] = item["Density"]
                        r["_c"] = item["c"]
                        r["_phi"] = item["phi"]
                        r["_test_global"] = test_counter
                    summary.update({
                        "_depth_no": depth_no, "_depth": item["Depth"],
                        "_sample_type": item["Sample Type"], "_soil": item["Soil"],
                        "_diameter": item["Diameter"], "_height": item["Height"],
                        "_density": item["Density"],
                        "_c": item["c"], "_phi": item["phi"],
                        "_test_global": test_counter,
                    })
                    all_rows.extend(rows)
                    all_summaries.append(summary)
                    current = summary["End Time"] + timedelta(minutes=master.uniform(5.0, 10.0))

            for seq, row in enumerate(all_rows, start=1):
                row["Seq#"] = seq
            self.rows, self.summaries = all_rows, all_summaries

            for x in self.tree.get_children():
                self.tree.delete(x)
            cols = ["Seq#", "Depth (m)", "Test", "Sample Type", "DateTime", "STRAIN (mm)", "LOAD (kg)"]
            self.tree["columns"] = cols
            for c in cols:
                self.tree.heading(c, text=c)
                self.tree.column(c, width=120, anchor="center")
            for r in self.rows[:200]:
                self.tree.insert("", "end", values=(
                    r["Seq#"], f'{r["_depth"]:.2f}', f'Test {r["_test_global"]}',
                    r["_sample_type"], r["DateTime"].strftime("%d-%m-%Y %H:%M:%S"),
                    f'{r["STRAIN"]:.3f}', f'{r["LOAD"]:.3f}'
                ))
            self.status.set(
                f"Generated {len(self.rows)} observations | {len(inp['depths'])} depth(s) | "
                f"{len(inp['depths'])*3} UU tests."
            )
        except Exception as e:
            messagebox.showerror("Generation Error", str(e))

    def ensure(self):
        if not self.rows:
            self.generate()
        return bool(self.rows)

    # ---------------- Final Mohr plot from generated results ----------------
    def final_mohr_plot(self):
        if not self.ensure():
            return
        if not MATPLOTLIB_AVAILABLE:
            messagebox.showerror(
                "Matplotlib Required",
                "Matplotlib is required for the final Mohr plot."
            )
            return

        inp = self.inputs()

        # Select the depth whose three ACTUAL generated tests will be plotted.
        choices = [
            f'{i+1}: Depth {d["Depth"]:.2f} m | {d["Sample Type"]} | {d["Soil"]}'
            for i, d in enumerate(inp["depths"])
        ]

        dialog = tk.Toplevel(self.root)
        dialog.title("Select Depth - Final Mohr Plot")
        dialog.geometry("560x255")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(
            dialog,
            text=(
                "Select a generated depth. The final plot uses the actual "
                "generated peak strengths, not the preliminary theoretical values."
            ),
            wraplength=520
        ).pack(anchor="w", padx=15, pady=(15, 8))

        choice = tk.StringVar(value=choices[0])
        ttk.Combobox(
            dialog, textvariable=choice, values=choices,
            state="readonly", width=68
        ).pack(padx=15, pady=5)

        result = {"index": None}

        def accept():
            result["index"] = choices.index(choice.get())
            dialog.destroy()

        bb = ttk.Frame(dialog)
        bb.pack(pady=15)
        ttk.Button(bb, text="Plot", command=accept).pack(
            side="left", padx=5
        )
        ttk.Button(bb, text="Cancel", command=dialog.destroy).pack(
            side="left", padx=5
        )

        self.root.wait_window(dialog)
        if result["index"] is None:
            return

        idx = result["index"]
        depth_item = inp["depths"][idx]
        summaries = [
            s for s in self.summaries
            if s.get("_depth_no") == idx + 1
        ]

        if len(summaries) < 3:
            messagebox.showerror(
                "Final Mohr Plot",
                "Three generated test summaries are required for this plot."
            )
            return

        sigma3_values = [s["Cell Pressure"] for s in summaries]
        sigma1_values = [s["Actual Peak sigma1"] for s in summaries]

        fit = calculate_final_mohr_coulomb_envelope(
            sigma3_values, sigma1_values
        )

        win = tk.Toplevel(self.root)
        win.title(
            f"Final UU Mohr's Circle Plot - Depth {depth_item['Depth']:.2f} m"
        )
        win.geometry("1120x780")
        win.minsize(900, 650)

        frame = ttk.Frame(win, padding=8)
        frame.pack(fill="both", expand=True)

        fig, fit = build_report_mohr_figure(
            inp["borehole"],
            depth_item,
            summaries,
            sigma3_values,
            sigma1_values
        )

        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        bottom = ttk.Frame(frame)
        bottom.pack(fill="x", pady=(6, 0))

        ttk.Label(
            bottom,
            text=(
                "Final envelope is fitted from the three actual generated "
                "failure points. Small specimen-to-specimen strength variation "
                "is intentionally retained."
            )
        ).pack(side="left", padx=5)

        def save_plot():
            path = filedialog.asksaveasfilename(
                title="Save Final Mohr Plot",
                defaultextension=".png",
                filetypes=[
                    ("PNG image", "*.png"),
                    ("JPEG image", "*.jpg"),
                    ("PDF", "*.pdf")
                ]
            )
            if not path:
                return
            fig.savefig(path, dpi=300, bbox_inches="tight")
            messagebox.showinfo(
                "Plot Saved",
                f"Final Mohr plot saved successfully:\n{path}"
            )

        ttk.Button(
            bottom, text="Save Plot Image", command=save_plot
        ).pack(side="right", padx=5)

        self.status.set(
            f"Final Mohr plot: Depth {depth_item['Depth']:.2f} m | "
            f"Cu={fit['Cu']:.2f} kg/cm² | φ={fit['phi']:.2f}°"
        )

    # ---------------- Export all final Mohr plots ----------------
    def export_all_mohr_plots(self):
        if not self.ensure():
            return
        if not MATPLOTLIB_AVAILABLE:
            messagebox.showerror("Matplotlib Required", "Matplotlib is required for plot export.")
            return

        inp = self.inputs()
        folder = filedialog.askdirectory(title="Select Folder for Mohr Plot Export")
        if not folder:
            return

        try:
            exported = []
            for idx, depth_item in enumerate(inp["depths"], start=1):
                summaries = [s for s in self.summaries if s.get("_depth_no") == idx]
                if len(summaries) < 3:
                    continue

                sigma3_values = [s["Cell Pressure"] for s in summaries]
                sigma1_values = [s["Actual Peak sigma1"] for s in summaries]

                fig, fit = build_report_mohr_figure(
                    inp["borehole"],
                    depth_item,
                    summaries,
                    sigma3_values,
                    sigma1_values
                )

                safe_depth = f"{depth_item['Depth']:.2f}".replace(".", "_")
                path = Path(folder) / f"Mohr_Plot_Depth_{safe_depth}m.png"
                fig.savefig(path, dpi=300, bbox_inches="tight")
                exported.append(str(path))
                fig.clear()

            if not exported:
                raise ValueError("No depth has three generated UU test results to plot.")

            self.status.set(f"Exported {len(exported)} final Mohr plot(s).")
            messagebox.showinfo(
                "Plot Export Complete",
                f"{len(exported)} final Mohr plot(s) exported successfully."
            )
        except Exception as e:
            messagebox.showerror("Plot Export Error", str(e))

    # ---------------- CSV ----------------
    def csv(self):
        if not self.ensure():
            return

        path = filedialog.asksaveasfilename(
            title="Export Observation CSV",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path:
            return

        try:
            # Build a completely new 4-column export table.
            # No other fields from the generated result dictionaries are
            # written to the CSV.
            export_rows = [
                [
                    row.get("Seq#", ""),
                    row["DateTime"].strftime("%d-%m-%Y %H:%M:%S")
                    if hasattr(row.get("DateTime"), "strftime")
                    else str(row.get("DateTime", "")),
                    f'{float(row.get("STRAIN", 0.0)):.3f}',
                    f'{float(row.get("LOAD", 0.0)):.3f}'
                ]
                for row in self.rows
            ]

            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow([
                    "Seq#",
                    "DateTime",
                    "STRAIN (mm)",
                    "LOAD (kg)"
                ])
                writer.writerows(export_rows)

            self.status.set(
                f"CSV exported: {len(export_rows)} observations, 4 columns."
            )
            messagebox.showinfo(
                "CSV Export Complete",
                "CSV exported successfully.\\n\\n"
                "Only these columns are included:\\n"
                "Seq# | DateTime | STRAIN (mm) | LOAD (kg)"
            )

        except Exception as e:
            messagebox.showerror("CSV Export Error", str(e))

    # ---------------- XLSM ----------------
    def xlsm(self):
        """
        Populate the 3-test-block XLSM template for ALL generated depths in one
        operation. The template is selected once and one XLSM file is created
        for each depth; no depth-selection dialog is used.
        """
        if not self.ensure():
            return

        inp = self.inputs()
        template = filedialog.askopenfilename(
            title="Select TRIAXIAL_TEMPLATE.xlsm",
            filetypes=[("Excel Macro Workbook", "*.xlsm")]
        )
        if not template:
            return

        folder = filedialog.askdirectory(
            title="Select Output Folder for All Depth XLSM Files"
        )
        if not folder:
            return

        try:
            output_folder = Path(folder)
            exported = []
            safe_bh = re.sub(r"[^A-Za-z0-9_-]+", "_", inp["borehole"]).strip("_")
            safe_bh = safe_bh or "BH"

            for idx, depth_item in enumerate(inp["depths"], start=1):
                depth_rows = [
                    r for r in self.rows
                    if r.get("_depth_no") == idx
                ]
                if not depth_rows:
                    raise ValueError(
                        f"No generated observation data found for depth "
                        f"{depth_item['Depth']:.2f} m."
                    )

                safe_depth = f"{depth_item['Depth']:.2f}".replace(".", "_")
                out = output_folder / (
                    f"{safe_bh}_Depth_{safe_depth}m_TRIAXIAL.xlsm"
                )

                populate_xlsm(
                    template, out, depth_rows,
                    depth_item["Diameter"], depth_item["Height"],
                    depth_item["Density"], depth_item["pressures"],
                    borehole=inp["borehole"],
                    depth=depth_item["Depth"]
                )
                exported.append(str(out))

            self.status.set(
                f"Populated XLSM template for all {len(exported)} depth(s)."
            )
            messagebox.showinfo(
                "All Depths Completed",
                f"{len(exported)} XLSM file(s) created successfully.\n\n"
                "One file was created for each depth; no depth selection was required."
            )
        except Exception as e:
            messagebox.showerror("XLSM Error", str(e))

    # ---------------- All Export (Plots + CSV + XLSM in a single ZIP) ----------------
    def export_all_package(self):
        if not self.ensure():
            return

        import zipfile
        import tempfile

        inp = self.inputs()
        safe_bh = re.sub(r"[^A-Za-z0-9_-]+", "_", inp["borehole"]).strip("_") or "BH"

        # Check default template location or prompt
        default_template = Path(__file__).resolve().parent / "TRIAXIAL_TEMPLATE.xlsm"
        if default_template.exists():
            template_file = str(default_template)
        else:
            template_file = filedialog.askopenfilename(
                title="Select TRIAXIAL_TEMPLATE.xlsm",
                filetypes=[("Excel Macro Workbook", "*.xlsm")]
            )
            if not template_file:
                return

        zip_save_path = filedialog.asksaveasfilename(
            title="Save All Export ZIP Archive",
            defaultextension=".zip",
            initialfile=f"{safe_bh}_Triaxial_All_Export.zip",
            filetypes=[("ZIP Archive", "*.zip")]
        )
        if not zip_save_path:
            return

        try:
            self.status.set("Creating all export package (Plots, CSV, XLSM)...")
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                plot_count = 0
                xlsm_count = 0

                # 1. Plots
                if MATPLOTLIB_AVAILABLE:
                    for idx, depth_item in enumerate(inp["depths"], start=1):
                        summaries = [s for s in self.summaries if s.get("_depth_no") == idx]
                        if len(summaries) < 3:
                            continue
                        sigma3_values = [s["Cell Pressure"] for s in summaries]
                        sigma1_values = [s["Actual Peak sigma1"] for s in summaries]

                        fig, fit = build_report_mohr_figure(
                            inp["borehole"],
                            depth_item,
                            summaries,
                            sigma3_values,
                            sigma1_values
                        )
                        safe_depth = f"{depth_item['Depth']:.2f}".replace(".", "_")
                        sample_type = depth_item.get("Sample Type", "UDS")
                        p_file = temp_path / f"{safe_bh}_Depth_{safe_depth}m_{sample_type}_Mohr_Plot.png"
                        fig.savefig(p_file, dpi=300, bbox_inches="tight")
                        fig.clear()
                        plot_count += 1

                # 2. CSV
                csv_file = temp_path / f"{safe_bh}_Triaxial_Observations.csv"
                with open(csv_file, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Seq#", "DateTime", "STRAIN (mm)", "LOAD (kg)"])
                    for row in self.rows:
                        dt_str = row["DateTime"].strftime("%d-%m-%Y %H:%M:%S") if hasattr(row.get("DateTime"), "strftime") else str(row.get("DateTime", ""))
                        writer.writerow([
                            row.get("Seq#", ""),
                            dt_str,
                            f'{float(row.get("STRAIN", 0.0)):.3f}',
                            f'{float(row.get("LOAD", 0.0)):.3f}'
                        ])

                # 3. XLSM
                for idx, depth_item in enumerate(inp["depths"], start=1):
                    depth_rows = [r for r in self.rows if r.get("_depth_no") == idx]
                    if not depth_rows:
                        continue
                    safe_depth = f"{depth_item['Depth']:.2f}".replace(".", "_")
                    sample_type = depth_item.get("Sample Type", "UDS")
                    xlsm_file = temp_path / f"{safe_bh}_Depth_{safe_depth}m_{sample_type}.xlsm"
                    populate_xlsm(
                        template_file, xlsm_file, depth_rows,
                        depth_item["Diameter"], depth_item["Height"],
                        depth_item["Density"], depth_item["pressures"],
                        borehole=inp["borehole"],
                        depth=depth_item["Depth"]
                    )
                    xlsm_count += 1

                # 4. Pack into ZIP
                with zipfile.ZipFile(zip_save_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for f in temp_path.glob("*"):
                        if f.is_file():
                            zipf.write(f, arcname=f.name)

            self.status.set(f"All Export Complete: {Path(zip_save_path).name}")
            messagebox.showinfo(
                "All Export Complete",
                f"Successfully exported and packaged into ZIP:\n\n"
                f"• {plot_count} Mohr Circle Plot(s) (.png)\n"
                f"• 1 Observations Data File (.csv)\n"
                f"• {xlsm_count} Populated Excel Template(s) (.xlsm)\n\n"
                f"Saved to: {zip_save_path}"
            )
        except Exception as e:
            messagebox.showerror("All Export Error", str(e))



if __name__=="__main__":
    root=tk.Tk()
    App(root)
    root.mainloop()


