#!/usr/bin/env python3
"""
Populate Excel workbook with add-in function examples.
Examples are consistent with the test suite in ActuarialAddIn.Tests.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import os

# Paths
EXCEL_PATH = os.path.join(os.path.dirname(__file__), '..', 'excel', 'actuarial_add_in_v0.0.xlsm')

# Common styles
TITLE_FONT = Font(bold=True, size=14)
HEADER_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT_WHITE = Font(bold=True, size=11, color="FFFFFF")
NOTE_FONT = Font(italic=True, size=10, color="666666")
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)


def set_column_widths(ws, widths):
    """Set column widths from a dictionary."""
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def add_title(ws, title, row=1, col=1):
    """Add a title cell."""
    cell = ws.cell(row=row, column=col, value=title)
    cell.font = TITLE_FONT
    return row + 2


def add_note(ws, note, row, col=1):
    """Add a note/description."""
    cell = ws.cell(row=row, column=col, value=note)
    cell.font = NOTE_FONT
    return row + 1


def add_table_header(ws, headers, row, start_col=1):
    """Add a table header row."""
    for i, header in enumerate(headers):
        cell = ws.cell(row=row, column=start_col + i, value=header)
        cell.font = HEADER_FONT_WHITE
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal='center')
    return row + 1


def add_data_row(ws, values, row, start_col=1, formulas=None):
    """Add a data row with optional formulas."""
    for i, value in enumerate(values):
        cell = ws.cell(row=row, column=start_col + i)
        if formulas and i < len(formulas) and formulas[i]:
            cell.value = formulas[i]
        else:
            cell.value = value
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal='center')
    return row + 1


def create_distributions_sheet(wb):
    """Create the Distributions examples sheet."""
    ws = wb.create_sheet("Distributions")
    set_column_widths(ws, {'A': 15, 'B': 20, 'C': 20, 'D': 20, 'E': 25})

    row = add_title(ws, "Statistical Distributions")
    row = add_note(ws, "PDF = Probability Density/Mass Function, CDF = Cumulative Distribution Function, INV = Inverse CDF (Quantile)", row)
    row += 1

    # Poisson
    row = add_note(ws, "POISSON DISTRIBUTION (lambda=5)", row)
    row = add_table_header(ws, ["k", "PDF", "CDF", "Notes"], row)
    for k in range(11):
        formulas = [None, f'=ACT_POISSON_PDF(A{row}, 5)', f'=ACT_POISSON_CDF(A{row}, 5)', None]
        notes = "Mode of distribution" if k == 5 else ""
        row = add_data_row(ws, [k, "", "", notes], row, formulas=formulas)

    row = add_note(ws, f"Inverse CDF example: =ACT_POISSON_INV(0.5, 5) returns the median", row)
    ws.cell(row=row-1, column=5, value="=ACT_POISSON_INV(0.5, 5)")
    row += 1

    # Negative Binomial
    row = add_note(ws, "NEGATIVE BINOMIAL DISTRIBUTION (r=5, p=0.3)", row)
    row = add_table_header(ws, ["k", "PDF", "CDF", "Notes"], row)
    for k in range(11):
        formulas = [None, f'=ACT_NEGBIN_PDF(A{row}, 5, 0.3)', f'=ACT_NEGBIN_CDF(A{row}, 5, 0.3)', None]
        row = add_data_row(ws, [k, "", "", ""], row, formulas=formulas)
    row += 1

    # Lognormal
    row = add_note(ws, "LOGNORMAL DISTRIBUTION (mu=0, sigma=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    for x in [0.5, 1.0, 1.5, 2.0, 3.0, 5.0]:
        formulas = [None, f'=ACT_LOGNORM_PDF(A{row}, 0, 1)', f'=ACT_LOGNORM_CDF(A{row}, 0, 1)', None]
        notes = "Median (exp(mu))" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    row += 1

    # Gamma
    row = add_note(ws, "GAMMA DISTRIBUTION (alpha=2, beta=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    for x in [0.5, 1.0, 2.0, 3.0, 5.0]:
        formulas = [None, f'=ACT_GAMMA_PDF(A{row}, 2, 1)', f'=ACT_GAMMA_CDF(A{row}, 2, 1)', None]
        notes = "Mean (alpha/beta)" if x == 2.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    row += 1

    # Pareto
    row = add_note(ws, "PARETO DISTRIBUTION (alpha=2, xm=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    for x in [1.0, 1.5, 2.0, 3.0, 5.0, 10.0]:
        formulas = [None, f'=ACT_PARETO_PDF(A{row}, 2, 1)', f'=ACT_PARETO_CDF(A{row}, 2, 1)', None]
        notes = "Minimum value (xm)" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)

    return ws


def create_exposure_curves_sheet(wb):
    """Create the Exposure Curves examples sheet."""
    ws = wb.create_sheet("Exposure Curves")
    set_column_widths(ws, {'A': 15, 'B': 20, 'C': 20, 'D': 20, 'E': 30})

    row = add_title(ws, "Exposure Curves")
    row = add_note(ws, "Exposure curves relate damage ratio (d) to expected loss ratio G(d)", row)
    row += 1

    # MBBEFD
    row = add_note(ws, "MBBEFD CURVES (b=2, g=3) - Swiss Re parameterization", row)
    row = add_table_header(ws, ["d", "G(d)", "Notes"], row)
    for d in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        formulas = [None, f'=ACT_MBBEFD(A{row}, 2, 3)', None]
        row = add_data_row(ws, [d, "", ""], row, formulas=formulas)
    row += 1

    # Swiss Re curves comparison
    row = add_note(ws, "SWISS RE CURVES COMPARISON at d=0.5", row)
    row = add_table_header(ws, ["Curve #", "G(0.5)", "Description"], row)
    descriptions = ["Light", "Medium-Light", "Medium", "Medium-Heavy", "Heavy"]
    for c in range(1, 6):
        formulas = [None, f'=ACT_SWISSRE_CURVE(0.5, A{row})', None]
        row = add_data_row(ws, [c, "", descriptions[c-1]], row, formulas=formulas)
    row += 1

    # Lloyd's curves
    row = add_note(ws, "LLOYD'S CURVES COMPARISON at d=0.5", row)
    row = add_table_header(ws, ["Curve", "G(0.5)", "Description"], row)
    lloyds = [("Y1", "Light industrial/commercial"), ("Y2", "Medium industrial"),
              ("Y3", "Heavy industrial"), ("Y4", "Petrochemical/refinery")]
    for curve, desc in lloyds:
        formulas = [None, f'=ACT_LLOYDS_CURVE(0.5, "{curve}")', None]
        row = add_data_row(ws, [curve, "", desc], row, formulas=formulas)
    row += 1

    # Power and Pareto
    row = add_note(ws, "OTHER EXPOSURE CURVES at d=0.5", row)
    row = add_table_header(ws, ["Function", "Formula", "Result", "Notes"], row)
    curves = [
        ("Power (n=2)", '=ACT_POWER_CURVE(0.5, 2)', "G(d) = d^n"),
        ("Inverse Power (n=2)", '=ACT_INVERSE_POWER_CURVE(0.5, 2)', "G(d) = 1 - (1-d)^n"),
        ("Pareto (alpha=2)", '=ACT_PARETO_EXPOSURE(0.5, 2)', "Based on Pareto severity"),
    ]
    for name, formula, notes in curves:
        ws.cell(row=row, column=1, value=name).border = THIN_BORDER
        ws.cell(row=row, column=2, value=formula.replace('=', '')).border = THIN_BORDER
        ws.cell(row=row, column=3, value=formula).border = THIN_BORDER
        ws.cell(row=row, column=4, value=notes).border = THIN_BORDER
        row += 1

    return ws


def create_reinsurance_sheet(wb):
    """Create the Reinsurance examples sheet."""
    ws = wb.create_sheet("Reinsurance")
    set_column_widths(ws, {'A': 20, 'B': 20, 'C': 20, 'D': 30})

    row = add_title(ws, "Reinsurance Functions")
    row = add_note(ws, "Excess of loss, quota share, and aggregate layer calculations", row)
    row += 1

    # XOL Layer
    row = add_note(ws, "XOL LAYER LOSS (Attachment=1,000,000, Limit=5,000,000)", row)
    row = add_table_header(ws, ["Ground-Up Loss", "Layer Loss", "Notes"], row)
    losses = [500000, 1000000, 2000000, 4000000, 6000000, 10000000]
    notes_list = ["Below attachment", "At attachment", "Partial layer", "Partial layer", "Full limit", "Capped at limit"]
    for loss, note in zip(losses, notes_list):
        formulas = [None, f'=ACT_XOL_LAYER_LOSS(A{row}, 1000000, 5000000)', None]
        row = add_data_row(ws, [loss, "", note], row, formulas=formulas)
    row += 1

    # Quota Share
    row = add_note(ws, "QUOTA SHARE (50% cession)", row)
    row = add_table_header(ws, ["Ground-Up Loss", "Ceded Loss", "Retained"], row)
    ws.cell(row=row, column=1, value=1000000).border = THIN_BORDER
    ws.cell(row=row, column=2, value="=ACT_QS_CEDED(A" + str(row) + ", 0.5)").border = THIN_BORDER
    ws.cell(row=row, column=3, value="=A" + str(row) + "-B" + str(row)).border = THIN_BORDER
    row += 2

    # Aggregate Layer
    row = add_note(ws, "AGGREGATE LAYER (Deductible=2,000,000, Limit=10,000,000)", row)
    row = add_table_header(ws, ["Aggregate Loss", "Layer Recovery", "Notes"], row)
    agg_losses = [1000000, 2000000, 5000000, 12000000, 15000000]
    agg_notes = ["Below deductible", "At deductible", "Partial recovery", "Full limit", "Capped at limit"]
    for loss, note in zip(agg_losses, agg_notes):
        formulas = [None, f'=ACT_AGGREGATE_LAYER(A{row}, 2000000, 10000000)', None]
        row = add_data_row(ws, [loss, "", note], row, formulas=formulas)

    return ws


def create_interpolation_sheet(wb):
    """Create the Interpolation examples sheet."""
    ws = wb.create_sheet("Interpolation")
    set_column_widths(ws, {'A': 15, 'B': 15, 'C': 20, 'D': 20, 'E': 30})

    row = add_title(ws, "Linear Interpolation")
    row = add_note(ws, "Interpolate between known points with optional extrapolation", row)
    row += 1

    # Known points
    row = add_note(ws, "KNOWN DATA POINTS", row)
    row = add_table_header(ws, ["X", "Y"], row)
    points = [(1, 10), (2, 20), (3, 35), (4, 55), (5, 80)]
    data_start_row = row
    for x, y in points:
        row = add_data_row(ws, [x, y], row)
    data_end_row = row - 1
    row += 1

    # Interpolation examples
    row = add_note(ws, "INTERPOLATION EXAMPLES", row)
    row = add_table_header(ws, ["X", "Y (FLAT extrap)", "Y (GRADIENT extrap)", "Notes"], row)
    x_range = f"$A${data_start_row}:$A${data_end_row}"
    y_range = f"$B${data_start_row}:$B${data_end_row}"

    test_x = [0.5, 1.5, 2.5, 3.5, 4.5, 5.5, 6.0]
    notes = ["Below range", "Interpolated", "Interpolated", "Interpolated", "Interpolated", "Above range", "Above range"]
    for x, note in zip(test_x, notes):
        flat_formula = f'=ACT_INTERP({x_range}, {y_range}, A{row}, "FLAT")'
        grad_formula = f'=ACT_INTERP({x_range}, {y_range}, A{row}, "GRADIENT")'
        ws.cell(row=row, column=1, value=x).border = THIN_BORDER
        ws.cell(row=row, column=2, value=flat_formula).border = THIN_BORDER
        ws.cell(row=row, column=3, value=grad_formula).border = THIN_BORDER
        ws.cell(row=row, column=4, value=note).border = THIN_BORDER
        row += 1

    row += 1
    row = add_note(ws, "FLAT = extrapolate using last known value; GRADIENT = extrapolate using slope from last two points", row)

    return ws


def create_chainladder_sheet(wb):
    """Create the Chain Ladder examples sheet."""
    ws = wb.create_sheet("Chain Ladder")
    set_column_widths(ws, {'A': 12, 'B': 12, 'C': 12, 'D': 12, 'E': 12, 'F': 12, 'G': 20, 'H': 15})

    row = add_title(ws, "Chain Ladder Reserving")
    row = add_note(ws, "Development factors, ultimates, IBNR, and Mack standard errors", row)
    row += 1

    # Input triangle
    row = add_note(ws, "INPUT TRIANGLE (Cumulative Paid Losses)", row)
    row = add_table_header(ws, ["AY \\ Dev", "1", "2", "3", "4", "5"], row)
    triangle = [
        [100, 150, 170, 180, 185],
        [110, 165, 190, 200, ""],
        [120, 180, 210, "", ""],
        [130, 195, "", "", ""],
        [140, "", "", "", ""]
    ]
    triangle_start_row = row
    for i, tri_row in enumerate(triangle):
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        for j, val in enumerate(tri_row):
            ws.cell(row=row, column=j+2, value=val if val != "" else None).border = THIN_BORDER
        row += 1
    triangle_end_row = row - 1
    row += 1

    # Triangle range for formulas
    tri_range = f"B{triangle_start_row}:F{triangle_end_row}"

    # Development factors
    row = add_note(ws, "DEVELOPMENT FACTORS (calculated from triangle)", row)
    ws.cell(row=row, column=1, value="Factors:").font = HEADER_FONT
    ws.cell(row=row, column=2, value=f"=ACT_CL_FACTORS({tri_range})")
    row += 2

    # Ultimates and IBNR
    row = add_note(ws, "PROJECTED ULTIMATES AND IBNR", row)
    row = add_table_header(ws, ["AY", "Latest", "Ultimate", "IBNR"], row)
    latest_values = [185, 200, 210, 195, 140]
    for i in range(5):
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        ws.cell(row=row, column=2, value=latest_values[i]).border = THIN_BORDER
        ws.cell(row=row, column=3, value=f"=INDEX(ACT_CL_ULTIMATE({tri_range}), {i+1})").border = THIN_BORDER
        ws.cell(row=row, column=4, value=f"=INDEX(ACT_CL_IBNR({tri_range}), {i+1})").border = THIN_BORDER
        row += 1

    row += 1
    row = add_note(ws, "MACK STANDARD ERRORS (reserve uncertainty)", row)
    ws.cell(row=row, column=1, value="Reserve SE:").font = HEADER_FONT
    ws.cell(row=row, column=2, value=f"=ACT_MACK_RESERVE_SE({tri_range})")

    return ws


def create_copulas_sheet(wb):
    """Create the Copulas examples sheet."""
    ws = wb.create_sheet("Copulas")
    set_column_widths(ws, {'A': 12, 'B': 12, 'C': 12, 'D': 12, 'E': 25})

    row = add_title(ws, "Student-t Copula")
    row = add_note(ws, "Generate correlated uniform random numbers for Monte Carlo simulation", row)
    row += 1

    # Correlation matrix
    row = add_note(ws, "CORRELATION MATRIX (3x3)", row)
    row = add_table_header(ws, ["", "X1", "X2", "X3"], row)
    corr_matrix = [
        ["X1", 1.0, 0.5, 0.3],
        ["X2", 0.5, 1.0, 0.4],
        ["X3", 0.3, 0.4, 1.0]
    ]
    corr_start_row = row
    for corr_row in corr_matrix:
        for j, val in enumerate(corr_row):
            cell = ws.cell(row=row, column=j+1, value=val)
            cell.border = THIN_BORDER
            if j == 0:
                cell.font = HEADER_FONT
        row += 1
    corr_end_row = row - 1
    row += 1

    corr_range = f"B{corr_start_row}:D{corr_end_row}"

    # Parameters
    row = add_note(ws, "PARAMETERS", row)
    ws.cell(row=row, column=1, value="Degrees of freedom:").border = THIN_BORDER
    ws.cell(row=row, column=2, value=5).border = THIN_BORDER
    df_cell = f"B{row}"
    row += 1
    ws.cell(row=row, column=1, value="Number of samples:").border = THIN_BORDER
    ws.cell(row=row, column=2, value=5).border = THIN_BORDER
    n_cell = f"B{row}"
    row += 1
    ws.cell(row=row, column=1, value="Random seed:").border = THIN_BORDER
    ws.cell(row=row, column=2, value=42).border = THIN_BORDER
    seed_cell = f"B{row}"
    row += 2

    # Output
    row = add_note(ws, "GENERATED SAMPLES (uniform marginals, correlated via t-copula)", row)
    row = add_note(ws, f"Formula: =ACT_STUDENT_T_COPULA({corr_range}, {df_cell}, {n_cell}, {seed_cell})", row)
    row = add_table_header(ws, ["Sample", "U1", "U2", "U3"], row)

    # The formula will spill results
    ws.cell(row=row, column=1, value=1).border = THIN_BORDER
    ws.cell(row=row, column=2, value=f"=ACT_STUDENT_T_COPULA({corr_range}, {df_cell}, {n_cell}, {seed_cell})").border = THIN_BORDER

    row += 6
    row = add_note(ws, "Note: Results are uniform[0,1] values. Apply inverse CDF to get desired marginal distributions.", row)

    return ws


def create_return_period_sheet(wb):
    """Create the Return Period examples sheet."""
    ws = wb.create_sheet("Return Periods")
    set_column_widths(ws, {'A': 18, 'B': 18, 'C': 20, 'D': 30})

    row = add_title(ws, "Return Period Functions")
    row = add_note(ws, "Generate losses from return period tables (OEP/AEP curves)", row)
    row += 1

    # Sample EP curve
    row = add_note(ws, "SAMPLE EXCEEDANCE PROBABILITY CURVE", row)
    row = add_table_header(ws, ["Return Period", "OEP Loss"], row)
    ep_data = [(10, 100000), (25, 250000), (50, 500000), (100, 1000000), (250, 2500000)]
    data_start = row
    for rp, loss in ep_data:
        row = add_data_row(ws, [rp, loss], row)
    data_end = row - 1
    row += 1

    rp_range = f"A{data_start}:A{data_end}"
    loss_range = f"B{data_start}:B{data_end}"

    # Interpolation examples
    row = add_note(ws, "LOSS INTERPOLATION", row)
    row = add_table_header(ws, ["Target RP", "Loss (LOG)", "Loss (LINEAR)", "Notes"], row)
    targets = [20, 75, 150, 200]
    for target in targets:
        log_formula = f'=ACT_RETURN_PERIOD_LOSS({rp_range}, {loss_range}, A{row}, "LOG")'
        lin_formula = f'=ACT_RETURN_PERIOD_LOSS({rp_range}, {loss_range}, A{row}, "LINEAR")'
        ws.cell(row=row, column=1, value=target).border = THIN_BORDER
        ws.cell(row=row, column=2, value=log_formula).border = THIN_BORDER
        ws.cell(row=row, column=3, value=lin_formula).border = THIN_BORDER
        ws.cell(row=row, column=4, value="Log interpolation typical for cat curves").border = THIN_BORDER
        row += 1

    row += 1

    # AAL calculation
    row = add_note(ws, "AVERAGE ANNUAL LOSS (AAL) FROM OEP CURVE", row)
    ws.cell(row=row, column=1, value="AAL:").font = HEADER_FONT
    ws.cell(row=row, column=2, value=f"=ACT_AAL_FROM_OEP({rp_range}, {loss_range})")

    return ws


def main():
    """Main function to populate the workbook."""
    print(f"Opening workbook: {EXCEL_PATH}")

    # Open existing workbook (preserve macros)
    wb = openpyxl.load_workbook(EXCEL_PATH, keep_vba=True)

    # Remove default sheet if it's empty
    if "Sheet1" in wb.sheetnames and wb["Sheet1"].max_row == 1:
        del wb["Sheet1"]

    # Create example sheets
    print("Creating Distributions sheet...")
    create_distributions_sheet(wb)

    print("Creating Exposure Curves sheet...")
    create_exposure_curves_sheet(wb)

    print("Creating Reinsurance sheet...")
    create_reinsurance_sheet(wb)

    print("Creating Interpolation sheet...")
    create_interpolation_sheet(wb)

    print("Creating Chain Ladder sheet...")
    create_chainladder_sheet(wb)

    print("Creating Copulas sheet...")
    create_copulas_sheet(wb)

    print("Creating Return Periods sheet...")
    create_return_period_sheet(wb)

    # Save
    print(f"Saving workbook...")
    wb.save(EXCEL_PATH)
    print("Done!")


if __name__ == "__main__":
    main()
