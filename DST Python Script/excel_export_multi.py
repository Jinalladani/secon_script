"""
Multi-depth Excel export for the DST Synthetic Data Generator.

Creates one workbook:
- Summary sheet: one row per depth with target and calculated c, phi and R².
- One DST sheet per depth, based on the supplied DST template.

The original template formatting/charts are retained on the first populated
template sheet; additional sheets are copied from the template structure.
"""

from pathlib import Path
import math
from copy import copy
from openpyxl import load_workbook


SHEET_NAME = "Template"
START_ROW = 8
N_OBSERVATIONS = 61


def _safe_sheet_name(name):
    bad = '[]:*?/\\'
    text = "".join("_" if c in bad else c for c in str(name))
    return text[:31] or "DST"


def _excel_strength_envelope(trials, pr_constant):
    sigma, tau = [], []

    for trial in trials:
        peak_tau = -float("inf")
        for displacement, shear_stress in zip(trial.displacement, trial.shear_stress):
            dial_reading = round(float(displacement) * 100.0)
            dial_cm = dial_reading / 1000.0
            corrected_area = 36.0 * (1.0 - (dial_cm / 3.0))
            if corrected_area <= 0:
                raise ValueError("Corrected area became zero or negative.")

            pr_reading = round(
                float(shear_stress) * corrected_area / pr_constant, 1
            )
            exported_tau = (pr_reading * pr_constant) / corrected_area
            peak_tau = max(peak_tau, exported_tau)

        sigma.append(float(trial.normal_stress))
        tau.append(peak_tau)

    xm = sum(sigma) / len(sigma)
    ym = sum(tau) / len(tau)
    sxx = sum((x-xm)**2 for x in sigma)
    sxy = sum((x-xm)*(y-ym) for x, y in zip(sigma, tau))
    syy = sum((y-ym)**2 for y in tau)

    if sxx <= 0 or syy <= 0:
        return 0.0, 0.0, 0.0

    slope = sxy / sxx
    c = ym - slope*xm
    phi = math.degrees(math.atan(slope))
    r = sxy / math.sqrt(sxx*syy)
    return c, phi, r*r


def _copy_template_sheet_structure(src, dst):
    # Copy cell values, formulas, styles and basic layout.
    for row in src.iter_rows():
        for cell in row:
            new = dst[cell.coordinate]
            new.value = cell.value
            if cell.has_style:
                new._style = copy(cell._style)
            if cell.number_format:
                new.number_format = cell.number_format
            if cell.alignment:
                new.alignment = copy(cell.alignment)
            if cell.protection:
                new.protection = copy(cell.protection)
            if cell.font:
                new.font = copy(cell.font)
            if cell.fill:
                new.fill = copy(cell.fill)
            if cell.border:
                new.border = copy(cell.border)

    for key, dim in src.column_dimensions.items():
        dst.column_dimensions[key].width = dim.width
        dst.column_dimensions[key].hidden = dim.hidden

    for key, dim in src.row_dimensions.items():
        dst.row_dimensions[key].height = dim.height
        dst.row_dimensions[key].hidden = dim.hidden

    for merged in src.merged_cells.ranges:
        dst.merge_cells(str(merged))

    dst.sheet_view.showGridLines = src.sheet_view.showGridLines


def _populate_sheet(ws, bh_id, item, pr_constant, normal_stresses):
    trials = item["trials"]

    ws["B1"] = bh_id
    ws["B2"] = float(item["depth"])
    ws["B3"] = float(item["density"])
    ws["B4"] = float(item["weight"])
    ws["B5"] = float(pr_constant)

    # Keep these blank as in the current template.
    for c in ("H2", "I2", "J2", "H3", "I3", "J3"):
        ws[c] = None

    ws["H4"] = float(normal_stresses[0])
    ws["I4"] = float(normal_stresses[1])
    ws["J4"] = float(normal_stresses[2])

    for i in range(N_OBSERVATIONS):
        row = START_ROW + i
        displacement = float(trials[0].displacement[i])
        dial_reading = round(displacement * 100.0)

        ws.cell(row, 1).value = i + 1
        ws.cell(row, 2).value = dial_reading
        ws.cell(row, 3).value = f"=B{row}/1000"
        ws.cell(row, 7).value = f"=36*(1-((C{row})/(3)))"

        dial_cm = dial_reading / 1000.0
        corrected_area = 36.0 * (1.0 - (dial_cm / 3.0))
        if corrected_area <= 0:
            raise ValueError(f"Corrected area became non-positive at observation {i+1}.")

        for trial_index, trial in enumerate(trials):
            col = 4 + trial_index
            pr_reading = round(
                float(trial.shear_stress[i]) * corrected_area / pr_constant, 1
            )
            ws.cell(row, col).value = pr_reading

        ws.cell(row, 8).value = f"=(D{row}*$B$5)/G{row}"
        ws.cell(row, 9).value = f"=(E{row}*$B$5)/G{row}"
        ws.cell(row, 10).value = f"=(F{row}*$B$5)/G{row}"

    calc_c, calc_phi, calc_r2 = _excel_strength_envelope(trials, pr_constant)

    # Put the calculated values into the same result cells used by the current template.
    ws["H2"] = calc_c
    ws["H3"] = calc_phi

    return calc_c, calc_phi, calc_r2


def export_multiple_dst_to_excel(
    template_path,
    output_path,
    bh_id,
    pr_constant,
    normal_stresses,
    results,
):
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"DST Excel template not found:\n{template_path}")
    if not results:
        raise ValueError("No generated depth results are available.")
    if len(normal_stresses) != 3:
        raise ValueError("Exactly three normal stresses are required.")

    wb = load_workbook(template_path)

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Worksheet '{SHEET_NAME}' was not found in the DST template.")

    template_ws = wb[SHEET_NAME]
    template_ws.title = _safe_sheet_name(f"Depth_{results[0]['depth']:.2f}m")

    summary_rows = []

    for index, item in enumerate(results):
        if index == 0:
            ws = template_ws
        else:
            ws = wb.create_sheet(_safe_sheet_name(f"Depth_{item['depth']:.2f}m"))
            _copy_template_sheet_structure(template_ws, ws)

        calc_c, calc_phi, calc_r2 = _populate_sheet(
            ws, bh_id, item, pr_constant, normal_stresses
        )

        summary_rows.append([
            bh_id,
            float(item["depth"]),
            item["soil"],
            float(item["density"]),
            float(item["weight"]),
            float(item["c"]),
            float(item["phi"]),
            float(calc_c),
            float(calc_phi),
            float(calc_r2),
        ])

    # Summary first
    summary = wb.create_sheet("Summary", 0)
    headers = [
        "BH ID", "Depth (m)", "Soil Type", "Density (g/cc)",
        "Sample Weight (g)", "Target c (kg/cm²)", "Target φ (°)",
        "Calculated c (kg/cm²)", "Calculated φ (°)", "R²"
    ]
    summary.append(headers)
    for row in summary_rows:
        summary.append(row)

    for cell in summary[1]:
        cell.font = copy(template_ws["A1"].font)
        cell.alignment = copy(template_ws["A1"].alignment)

    widths = [14, 12, 24, 16, 19, 20, 14, 24, 19, 10]
    for i, width in enumerate(widths, 1):
        summary.column_dimensions[chr(64+i)].width = width

    for row in summary.iter_rows(min_row=2):
        for col in (2, 4, 5, 6, 7, 8, 9, 10):
            row[col-1].number_format = "0.00"
        row[5].number_format = "0.000"
        row[7].number_format = "0.000"
        row[9].number_format = "0.000"

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return {
        "output_path": str(output_path),
        "depth_count": len(results),
        "summary_rows": summary_rows,
    }
