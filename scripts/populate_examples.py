#!/usr/bin/env python3
"""
Populate Excel workbook with add-in function examples.
Examples are consistent with the test suite in ActuarialAddIn.Tests.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, Color
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, ScatterChart, BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.title import Title
from openpyxl.chart.data_source import StrRef
import os

# Paths
EXCEL_PATH = os.path.join(os.path.dirname(__file__), '..', 'excel', 'actuarial_add_in_v0.2.xlsm')

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

# Chart colors
CHART_COLORS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]


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


def create_scatter_chart_with_axes(ws, title, x_col, y_col, min_row, max_row, anchor,
                                   x_title=None, y_title=None, width=10, height=7,
                                   x_min=None, y_min=None, x_max=None, y_max=None):
    """Create a simple line chart with a single series."""
    chart = ScatterChart()
    chart.title = title
    chart.width = width
    chart.height = height
    chart.roundedCorners = False  # No rounded corners

    # Axis titles and settings
    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title

    # Axis scaling
    if x_min is not None:
        chart.x_axis.scaling.min = x_min
    if x_max is not None:
        chart.x_axis.scaling.max = x_max
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    if y_max is not None:
        chart.y_axis.scaling.max = y_max

    # Ensure axes are visible with tick marks and numbers
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "low"
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    xvalues = Reference(ws, min_col=x_col, min_row=min_row+1, max_row=max_row)
    yvalues = Reference(ws, min_col=y_col, min_row=min_row+1, max_row=max_row)

    from openpyxl.chart import Series
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker = Marker(symbol='none')  # No markers, just line
    series.graphicalProperties.line.solidFill = CHART_COLORS[0]
    chart.series.append(series)
    chart.legend = None  # No legend for single series

    ws.add_chart(chart, anchor)
    return chart


def create_bar_chart_with_axes(ws, title, cat_col, data_col, min_row, max_row, anchor,
                                x_title=None, y_title=None, width=12, height=8):
    """Create a bar chart with proper axes."""
    chart = BarChart()
    chart.title = title
    chart.type = "col"
    chart.grouping = "standard"
    chart.width = width
    chart.height = height
    chart.roundedCorners = False  # No rounded corners
    
    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title
    
    chart.y_axis.scaling.min = 0
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "low"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    
    data = Reference(ws, min_col=data_col, min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=cat_col, min_row=min_row+1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    
    ws.add_chart(chart, anchor)
    return chart


def add_distribution_section(ws, row, title, x_values, pdf_formula, cdf_formula, x_label="x", notes_dict=None):
    """Add a distribution section with table and return data range info."""
    section_start = row
    row = add_note(ws, title, row)
    row = add_table_header(ws, [x_label, "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in x_values:
        formulas = [None, pdf_formula.format(row=row), cdf_formula.format(row=row), None]
        notes = notes_dict.get(x, "") if notes_dict else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    data_end = row - 1
    return row, section_start, data_start, data_end


def create_distributions_sheet(wb):
    """Create the Distributions examples sheet with ALL distributions."""
    ws = wb.create_sheet("Distributions")
    set_column_widths(ws, {'A': 12, 'B': 15, 'C': 15, 'D': 20, 'E': 5, 'F': 15})

    # Chart settings - all charts same size
    CHART_WIDTH = 8
    CHART_HEIGHT = 6
    CHART_COL = "F"  # Column for charts
    SECTION_PADDING = 15  # Minimum rows per section (for chart height)

    row = add_title(ws, "Statistical Distributions")
    row = add_note(ws, "PDF = Probability Density Function, CDF = Cumulative Distribution Function", row)
    row += 1

    distributions = []

    # 1. Normal Distribution
    section_start = row
    row = add_note(ws, "NORMAL DISTRIBUTION (mu=0, sigma=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [-3, -2, -1, -0.5, 0, 0.5, 1, 2, 3]:
        formulas = [None, f'=ACT_DIST_NORMAL_PDF(A{row}, 0, 1)', f'=ACT_DIST_NORMAL_CDF(A{row}, 0, 1)', None]
        notes = "Mean" if x == 0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Normal PDF (μ=0, σ=1)", section_start, data_start, row - 1, "x", "f(x)", None, None))
    row = max(row, section_start + SECTION_PADDING)

    # 2. Poisson Distribution
    section_start = row
    row = add_note(ws, "POISSON DISTRIBUTION (lambda=5)", row)
    row = add_table_header(ws, ["k", "PDF", "CDF", "Notes"], row)
    data_start = row
    for k in range(11):
        formulas = [None, f'=ACT_DIST_POISSON_PDF(A{row}, 5)', f'=ACT_DIST_POISSON_CDF(A{row}, 5)', None]
        notes = "Mode" if k == 5 else ""
        row = add_data_row(ws, [k, "", "", notes], row, formulas=formulas)
    distributions.append(("Poisson PMF (λ=5)", section_start, data_start, row - 1, "k", "P(X=k)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 3. Negative Binomial
    section_start = row
    row = add_note(ws, "NEGATIVE BINOMIAL DISTRIBUTION (r=5, p=0.3)", row)
    row = add_table_header(ws, ["k", "PDF", "CDF", "Notes"], row)
    data_start = row
    for k in range(12):
        formulas = [None, f'=ACT_DIST_NEGBIN_PDF(A{row}, 5, 0.3)', f'=ACT_DIST_NEGBIN_CDF(A{row}, 5, 0.3)', None]
        row = add_data_row(ws, [k, "", "", ""], row, formulas=formulas)
    distributions.append(("Negative Binomial PMF", section_start, data_start, row - 1, "k", "P(X=k)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 4. Lognormal
    section_start = row
    row = add_note(ws, "LOGNORMAL DISTRIBUTION (mu=0, sigma=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.1, 0.25, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]:
        formulas = [None, f'=ACT_DIST_LOGNORM_PDF(A{row}, 0, 1)', f'=ACT_DIST_LOGNORM_CDF(A{row}, 0, 1)', None]
        notes = "Median" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Lognormal PDF (μ=0, σ=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 5. Gamma
    section_start = row
    row = add_note(ws, "GAMMA DISTRIBUTION (alpha=2, beta=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0, 6.0]:
        formulas = [None, f'=ACT_DIST_GAMMA_PDF(A{row}, 2, 1)', f'=ACT_DIST_GAMMA_CDF(A{row}, 2, 1)', None]
        notes = "Mean" if x == 2.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Gamma PDF (α=2, β=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 6. Exponential
    section_start = row
    row = add_note(ws, "EXPONENTIAL DISTRIBUTION (lambda=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0]:
        formulas = [None, f'=ACT_DIST_EXP_PDF(A{row}, 1)', f'=ACT_DIST_EXP_CDF(A{row}, 1)', None]
        notes = "Mean" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Exponential PDF (λ=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 7. Weibull
    section_start = row
    row = add_note(ws, "WEIBULL DISTRIBUTION (k=2, lambda=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.25, 0.5, 0.75, 1.0, 1.5, 2.0, 2.5, 3.0]:
        formulas = [None, f'=ACT_DIST_WEIBULL_PDF(A{row}, 2, 1)', f'=ACT_DIST_WEIBULL_CDF(A{row}, 2, 1)', None]
        row = add_data_row(ws, [x, "", "", ""], row, formulas=formulas)
    distributions.append(("Weibull PDF (k=2, λ=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 8. Beta
    section_start = row
    row = add_note(ws, "BETA DISTRIBUTION (alpha=2, beta=5)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        formulas = [None, f'=ACT_DIST_BETA_PDF(A{row}, 2, 5)', f'=ACT_DIST_BETA_CDF(A{row}, 2, 5)', None]
        row = add_data_row(ws, [x, "", "", ""], row, formulas=formulas)
    distributions.append(("Beta PDF (α=2, β=5)", section_start, data_start, row - 1, "x", "f(x)", 0, 1))
    row = max(row, section_start + SECTION_PADDING)

    # 9. Pareto Type I
    section_start = row
    row = add_note(ws, "PARETO TYPE I DISTRIBUTION (alpha=2, xm=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [1.0, 1.5, 2.0, 3.0, 4.0, 5.0, 7.0, 10.0]:
        formulas = [None, f'=ACT_DIST_PARETO_PDF(A{row}, 2, 1)', f'=ACT_DIST_PARETO_CDF(A{row}, 2, 1)', None]
        notes = "Minimum (xm)" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Pareto I PDF (α=2, xm=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 10. Lomax (Pareto Type II)
    section_start = row
    row = add_note(ws, "LOMAX (PARETO TYPE II) DISTRIBUTION (alpha=2, lambda=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.5, 1.0, 2.0, 3.0, 4.0, 5.0, 7.0, 10.0]:
        formulas = [None, f'=ACT_DIST_LOMAX_PDF(A{row}, 2, 1)', f'=ACT_DIST_LOMAX_CDF(A{row}, 2, 1)', None]
        row = add_data_row(ws, [x, "", "", ""], row, formulas=formulas)
    distributions.append(("Lomax PDF (α=2, λ=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 11. GPD (Generalized Pareto)
    section_start = row
    row = add_note(ws, "GENERALIZED PARETO (GPD) DISTRIBUTION (xi=0.5, sigma=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.25, 0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]:
        formulas = [None, f'=ACT_DIST_GPD_PDF(A{row}, 0.5, 1)', f'=ACT_DIST_GPD_CDF(A{row}, 0.5, 1)', None]
        row = add_data_row(ws, [x, "", "", ""], row, formulas=formulas)
    distributions.append(("GPD PDF (ξ=0.5, σ=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 12. Burr Type XII
    section_start = row
    row = add_note(ws, "BURR TYPE XII DISTRIBUTION (c=2, k=1, lambda=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.25, 0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]:
        formulas = [None, f'=ACT_DIST_BURR_PDF(A{row}, 2, 1, 1)', f'=ACT_DIST_BURR_CDF(A{row}, 2, 1, 1)', None]
        row = add_data_row(ws, [x, "", "", ""], row, formulas=formulas)
    distributions.append(("Burr XII PDF (c=2, k=1, λ=1)", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 13. Lognormal-Pareto (Composite)
    section_start = row
    row = add_note(ws, "LOGNORMAL-PARETO COMPOSITE (mu=0, sigma=1, theta=2)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.1, 0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0, 7.0, 10.0]:
        formulas = [None, f'=ACT_DIST_LNPARETO_PDF(A{row}, 0, 1, 2)', f'=ACT_DIST_LNPARETO_CDF(A{row}, 0, 1, 2)', None]
        notes = "Threshold" if x == 2.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Lognormal-Pareto PDF", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 14. Exponential-Pareto (Composite)
    section_start = row
    row = add_note(ws, "EXPONENTIAL-PARETO COMPOSITE (rate=0.5, theta=2)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.5, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0, 7.0, 10.0]:
        formulas = [None, f'=ACT_DIST_EXPPARETO_PDF(A{row}, 0.5, 2)', f'=ACT_DIST_EXPPARETO_CDF(A{row}, 0.5, 2)', None]
        notes = "Threshold" if x == 2.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Exponential-Pareto PDF", section_start, data_start, row - 1, "x", "f(x)", 0, None))
    row = max(row, section_start + SECTION_PADDING)

    # 15. Power-Pareto (Composite)
    section_start = row
    row = add_note(ws, "POWER-PARETO COMPOSITE (alpha=2, beta=3, theta=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    data_start = row
    for x in [0.0, 0.25, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]:
        formulas = [None, f'=ACT_DIST_POWPARETO_PDF(A{row}, 2, 3, 1)', f'=ACT_DIST_POWPARETO_CDF(A{row}, 2, 3, 1)', None]
        notes = "Threshold" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    distributions.append(("Power-Pareto PDF", section_start, data_start, row - 1, "x", "f(x)", 0, None))

    # Create all charts aligned with their tables
    for title, section_start, data_start, data_end, x_title, y_title, x_min, x_max in distributions:
        anchor = f"{CHART_COL}{section_start}"
        create_scatter_chart_with_axes(ws, title, x_col=1, y_col=2,
            min_row=data_start-1, max_row=data_end, anchor=anchor,
            x_title=x_title, y_title=y_title,
            width=CHART_WIDTH, height=CHART_HEIGHT,
            x_min=x_min, y_min=0, x_max=x_max)

    return ws


def create_exposure_curves_sheet(wb):
    """Create the Exposure Curves examples sheet."""
    ws = wb.create_sheet("Exposure Curves")
    set_column_widths(ws, {'A': 15, 'B': 20, 'C': 20, 'D': 20, 'E': 30, 'F': 5})

    row = add_title(ws, "Exposure Curves")
    row = add_note(ws, "Exposure curves relate damage ratio (d) to expected loss ratio G(d)", row)
    row += 1

    # MBBEFD
    row = add_note(ws, "MBBEFD CURVES (b=2, g=3) - Swiss Re parameterization", row)
    row = add_table_header(ws, ["d", "G(d)", "Notes"], row)
    mbbefd_data_start = row
    for d in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        formulas = [None, f'=ACT_EXPOSURE_MBBEFD(A{row}, 2, 3)', None]
        row = add_data_row(ws, [d, "", ""], row, formulas=formulas)
    mbbefd_data_end = row - 1
    row += 1

    create_scatter_chart_with_axes(ws, "MBBEFD Exposure Curve (b=2, g=3)", x_col=1, y_col=2,
        min_row=mbbefd_data_start-1, max_row=mbbefd_data_end, anchor="F4",
        x_title="Damage Ratio (d)", y_title="Loss Ratio G(d)",
        x_min=0, x_max=1, y_min=0, y_max=1)

    # Swiss Re curves comparison
    row = add_note(ws, "SWISS RE CURVES COMPARISON (all curves)", row)
    row = add_table_header(ws, ["d", "Curve 1", "Curve 2", "Curve 3", "Curve 4", "Curve 5"], row)
    swissre_data_start = row
    for d in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        ws.cell(row=row, column=1, value=d).border = THIN_BORDER
        for c in range(1, 6):
            ws.cell(row=row, column=c+1, value=f'=ACT_EXPOSURE_SWISSRE(A{row}, {c})').border = THIN_BORDER
        row += 1
    swissre_data_end = row - 1
    row += 1

    # Lloyd's curves
    row = add_note(ws, "LLOYD'S CURVES COMPARISON", row)
    row = add_table_header(ws, ["d", "Y1", "Y2", "Y3", "Y4"], row)
    lloyds_data_start = row
    for d in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        ws.cell(row=row, column=1, value=d).border = THIN_BORDER
        for i, curve in enumerate(["Y1", "Y2", "Y3", "Y4"]):
            ws.cell(row=row, column=i+2, value=f'=ACT_EXPOSURE_LLOYDS(A{row}, "{curve}")').border = THIN_BORDER
        row += 1
    lloyds_data_end = row - 1
    row += 1

    # Power and Pareto
    row = add_note(ws, "OTHER EXPOSURE CURVES at d=0.5", row)
    row = add_table_header(ws, ["Function", "Formula", "Result", "Notes"], row)
    curves = [
        ("Power (n=2)", '=ACT_EXPOSURE_POWER(0.5, 2)', "G(d) = d^n"),
        ("Inverse Power (n=2)", '=ACT_EXPOSURE_INVERSE_POWER(0.5, 2)', "G(d) = 1 - (1-d)^n"),
        ("Pareto (alpha=2)", '=ACT_EXPOSURE_PARETO(0.5, 2)', "Based on Pareto severity"),
        ("Riebesell (n=0.5)", '=ACT_EXPOSURE_RIEBESELL(0.5, 0.5)', "Alternative single-parameter curve"),
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
    row = add_note(ws, "Excess of loss layer calculations", row)
    row += 1

    # XOL Layer
    row = add_note(ws, "XOL LAYER LOSS (Attachment=1,000,000, Limit=5,000,000)", row)
    row = add_table_header(ws, ["Ground-Up Loss", "Layer Loss", "Notes"], row)
    losses = [500000, 1000000, 2000000, 4000000, 6000000, 10000000]
    notes_list = ["Below attachment", "At attachment", "Partial layer", "Partial layer", "Full limit", "Capped at limit"]
    for loss, note in zip(losses, notes_list):
        cell_a = ws.cell(row=row, column=1, value=loss)
        cell_a.border = THIN_BORDER
        cell_a.number_format = '#,##0'
        cell_b = ws.cell(row=row, column=2, value=f'=ACT_XOL_LAYER_LOSS(A{row}, 1000000, 5000000)')
        cell_b.border = THIN_BORDER
        cell_b.number_format = '#,##0'
        cell_c = ws.cell(row=row, column=3, value=note)
        cell_c.border = THIN_BORDER
        row += 1

    return ws


def create_interpolation_sheet(wb):
    """Create the Interpolation examples sheet."""
    ws = wb.create_sheet("Interpolation")
    set_column_widths(ws, {'A': 15, 'B': 15, 'C': 20, 'D': 20, 'E': 30, 'F': 5})

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

    interp_start_row = row
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
    interp_end_row = row - 1

    row += 1
    row = add_note(ws, "FLAT = extrapolate using last known value; GRADIENT = extrapolate using slope from last two points", row)

    # Simple line chart showing known data points
    chart = ScatterChart()
    chart.title = "Known Data Points"
    chart.width = 12
    chart.height = 8
    chart.roundedCorners = False
    chart.x_axis.title = "X"
    chart.y_axis.title = "Y"
    chart.x_axis.scaling.min = 0
    chart.y_axis.scaling.min = 0
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "low"
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    from openpyxl.chart import Series
    xvalues = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)
    yvalues = Reference(ws, min_col=2, min_row=data_start_row, max_row=data_end_row)
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker = Marker(symbol='none')
    series.graphicalProperties.line.solidFill = CHART_COLORS[0]
    chart.series.append(series)
    chart.legend = None

    ws.add_chart(chart, "F4")

    return ws


def create_chainladder_sheet(wb):
    """Create the Chain Ladder examples sheet."""
    ws = wb.create_sheet("Chain Ladder")
    set_column_widths(ws, {'A': 10, 'B': 12, 'C': 12, 'D': 12, 'E': 12, 'F': 12,
                           'G': 12, 'H': 12, 'I': 12, 'J': 12, 'K': 12, 'L': 5, 'M': 15})

    row = add_title(ws, "Chain Ladder Reserving - Taylor-Ashe Dataset")
    row = add_note(ws, "Using the Taylor-Ashe triangle from Peter England's bootstrapping presentation (England & Verrall, 1999)", row)
    row += 1

    # Input triangle
    row = add_note(ws, "INPUT TRIANGLE (Cumulative Paid Losses - Taylor-Ashe)", row)
    row = add_table_header(ws, ["AY", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], row)

    triangle = [
        [357848, 1124788, 1735330, 2218270, 2745596, 3319994, 3466336, 3606286, 3833515, 3901463],
        [352118, 1236139, 2170033, 3353322, 3799067, 4120063, 4647867, 4914039, 5339085, None],
        [290507, 1292306, 2218525, 3235179, 3985995, 4132918, 4628910, 4909315, None, None],
        [310608, 1418858, 2195047, 3757447, 4029929, 4381982, 4588268, None, None, None],
        [443160, 1136350, 2128333, 2897821, 3402672, 3873311, None, None, None, None],
        [396132, 1333217, 2180715, 2985752, 3691712, None, None, None, None, None],
        [440832, 1288463, 2419861, 3483130, None, None, None, None, None, None],
        [359480, 1421128, 2864498, None, None, None, None, None, None, None],
        [376686, 1363294, None, None, None, None, None, None, None, None],
        [344014, None, None, None, None, None, None, None, None, None]
    ]

    triangle_start_row = row
    for i, tri_row in enumerate(triangle):
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        for j, val in enumerate(tri_row):
            cell = ws.cell(row=row, column=j+2, value=val)
            cell.border = THIN_BORDER
            if val is not None:
                cell.number_format = '#,##0'
        row += 1
    triangle_end_row = row - 1
    row += 1

    tri_range = f"B{triangle_start_row}:K{triangle_end_row}"

    # Development factors - use INDEX to extract single values
    row = add_note(ws, "DEVELOPMENT FACTORS", row)
    row = add_note(ws, "Expected: 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018", row)
    row = add_table_header(ws, ["Period", "1-2", "2-3", "3-4", "4-5", "5-6", "6-7", "7-8", "8-9", "9-10"], row)
    factors_row = row
    ws.cell(row=row, column=1, value="Factors").border = THIN_BORDER
    for i in range(9):
        ws.cell(row=row, column=i+2, value=f"=INDEX(ACT_CL_FACTORS({tri_range}), {i+1})").border = THIN_BORDER
    row += 2

    # Ultimates and IBNR
    row = add_note(ws, "PROJECTED ULTIMATES AND IBNR", row)
    row = add_note(ws, "Expected Total IBNR: ~18,680,856", row)
    row = add_table_header(ws, ["AY", "Latest Cumulative", "Ultimate", "IBNR"], row)

    latest_values = [3901463, 5339085, 4909315, 4588268, 3873311, 3691712, 3483130, 2864498, 1363294, 344014]
    ibnr_start = row
    for i in range(10):
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        cell_latest = ws.cell(row=row, column=2, value=latest_values[i])
        cell_latest.border = THIN_BORDER
        cell_latest.number_format = '#,##0'
        cell_ult = ws.cell(row=row, column=3, value=f"=INDEX(ACT_CL_ULTIMATE({tri_range}), {i+1})")
        cell_ult.border = THIN_BORDER
        cell_ult.number_format = '#,##0'
        cell_ibnr = ws.cell(row=row, column=4, value=f"=INDEX(ACT_CL_IBNR({tri_range}), {i+1})")
        cell_ibnr.border = THIN_BORDER
        cell_ibnr.number_format = '#,##0'
        row += 1
    ibnr_end = row - 1

    # Total row
    ws.cell(row=row, column=1, value="Total").font = HEADER_FONT
    ws.cell(row=row, column=1).border = THIN_BORDER
    cell_sum_latest = ws.cell(row=row, column=2, value=f"=SUM(B{ibnr_start}:B{ibnr_end})")
    cell_sum_latest.border = THIN_BORDER
    cell_sum_latest.number_format = '#,##0'
    cell_sum_latest.font = HEADER_FONT
    cell_sum_ult = ws.cell(row=row, column=3, value=f"=SUM(C{ibnr_start}:C{ibnr_end})")
    cell_sum_ult.border = THIN_BORDER
    cell_sum_ult.number_format = '#,##0'
    cell_sum_ult.font = HEADER_FONT
    cell_sum_ibnr = ws.cell(row=row, column=4, value=f"=SUM(D{ibnr_start}:D{ibnr_end})")
    cell_sum_ibnr.border = THIN_BORDER
    cell_sum_ibnr.number_format = '#,##0'
    cell_sum_ibnr.font = HEADER_FONT
    row += 2

    # Mack Standard Errors with England & Verrall reference
    row = add_note(ws, "MACK STANDARD ERRORS (reserve uncertainty)", row)
    row = add_table_header(ws, ["AY", "Reserve SE", "E&V Reference"], row)
    
    # England & Verrall (2002) Mack SE reference values
    mack_se_reference = [0, 75535, 121699, 136554, 212054, 282372, 444283, 640396, 1091089, 2029818]
    
    se_start = row
    for i in range(10):
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        cell_se = ws.cell(row=row, column=2, value=f"=INDEX(ACT_MACK_RESERVE_SE({tri_range}), {i+1})")
        cell_se.border = THIN_BORDER
        cell_se.number_format = '#,##0'
        cell_ref = ws.cell(row=row, column=3, value=mack_se_reference[i])
        cell_ref.border = THIN_BORDER
        cell_ref.number_format = '#,##0'
        row += 1
    se_end = row - 1
    
    # Total Mack SE
    ws.cell(row=row, column=1, value="Total").font = HEADER_FONT
    ws.cell(row=row, column=1).border = THIN_BORDER
    ws.cell(row=row, column=3, value=2447318).border = THIN_BORDER  # E&V total SE
    ws.cell(row=row, column=3).number_format = '#,##0'
    row += 2

    # Bootstrap Chain Ladder
    row = add_note(ws, "BOOTSTRAP CHAIN LADDER (ODP) - England & Verrall (2002)", row)
    ws.cell(row=row, column=1, value="Iterations").border = THIN_BORDER
    ws.cell(row=row, column=2, value=1000).border = THIN_BORDER
    iter_cell = f"B{row}"
    row += 1
    ws.cell(row=row, column=1, value="Seed").border = THIN_BORDER
    ws.cell(row=row, column=2, value=42).border = THIN_BORDER
    seed_cell = f"B{row}"
    row += 1

    row = add_note(ws, "BOOTSTRAP RESULTS (array output - use INDEX to extract)", row)
    row = add_table_header(ws, ["Statistic", "Value", "E&V Reference"], row)
    
    # England & Verrall reference values for bootstrap (Table 6, constant scale)
    bootstrap_reference = {
        "Mean": 18680856,  # Same as deterministic
        "StdDev": 3087570,  # Total prediction error
        "P1": None,
        "P5": None,
        "P10": None,
        "P25": None,
        "P50": 18186033,  # Median
        "P75": 20376564,  # 75th percentile
        "P90": None,
        "P95": 24221956,  # 95th percentile
        "P99": None
    }
    
    bootstrap_stats = ["Mean", "StdDev", "P1", "P5", "P10", "P25", "P50", "P75", "P90", "P95", "P99"]
    for i, stat in enumerate(bootstrap_stats):
        ws.cell(row=row, column=1, value=stat).border = THIN_BORDER
        cell_val = ws.cell(row=row, column=2, value=f"=INDEX(ACT_CL_BOOTSTRAP({tri_range}, {iter_cell}, {seed_cell}), {i+1}, 2)")
        cell_val.border = THIN_BORDER
        cell_val.number_format = '#,##0'
        ref_val = bootstrap_reference.get(stat)
        if ref_val:
            cell_ref = ws.cell(row=row, column=3, value=ref_val)
            cell_ref.border = THIN_BORDER
            cell_ref.number_format = '#,##0'
        row += 1
    row += 1

    row = add_note(ws, "BOOTSTRAP BY ORIGIN YEAR - Compare with England & Verrall (2002) Table 5", row)
    row = add_table_header(ws, ["AY", "Mean", "StdDev", "E&V SE Ref", "P50", "P75", "P90"], row)
    
    # England & Verrall Table 5 - Prediction Error by Origin Year (constant scale)
    ev_se_by_origin = [0, 112379, 178443, 209399, 286636, 440310, 571035, 820842, 1258296, 2046223]
    
    for i in range(10):
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        # Mean
        cell_mean = ws.cell(row=row, column=2, value=f"=INDEX(ACT_CL_BOOTSTRAP_ORIGIN({tri_range}, {iter_cell}, {seed_cell}), {i+2}, 2)")
        cell_mean.border = THIN_BORDER
        cell_mean.number_format = '#,##0'
        # StdDev
        cell_sd = ws.cell(row=row, column=3, value=f"=INDEX(ACT_CL_BOOTSTRAP_ORIGIN({tri_range}, {iter_cell}, {seed_cell}), {i+2}, 3)")
        cell_sd.border = THIN_BORDER
        cell_sd.number_format = '#,##0'
        # E&V Reference SE
        cell_ref = ws.cell(row=row, column=4, value=ev_se_by_origin[i])
        cell_ref.border = THIN_BORDER
        cell_ref.number_format = '#,##0'
        # P50
        cell_p50 = ws.cell(row=row, column=5, value=f"=INDEX(ACT_CL_BOOTSTRAP_ORIGIN({tri_range}, {iter_cell}, {seed_cell}), {i+2}, 4)")
        cell_p50.border = THIN_BORDER
        cell_p50.number_format = '#,##0'
        # P75
        cell_p75 = ws.cell(row=row, column=6, value=f"=INDEX(ACT_CL_BOOTSTRAP_ORIGIN({tri_range}, {iter_cell}, {seed_cell}), {i+2}, 5)")
        cell_p75.border = THIN_BORDER
        cell_p75.number_format = '#,##0'
        # P90
        cell_p90 = ws.cell(row=row, column=7, value=f"=INDEX(ACT_CL_BOOTSTRAP_ORIGIN({tri_range}, {iter_cell}, {seed_cell}), {i+2}, 6)")
        cell_p90.border = THIN_BORDER
        cell_p90.number_format = '#,##0'
        row += 1
    
    # Total row with E&V reference
    ws.cell(row=row, column=1, value="Total").font = HEADER_FONT
    ws.cell(row=row, column=1).border = THIN_BORDER
    ws.cell(row=row, column=4, value=3087570).border = THIN_BORDER  # E&V total SE
    ws.cell(row=row, column=4).number_format = '#,##0'
    row += 1

    # Charts - IBNR bar chart
    chart1 = BarChart()
    chart1.title = "IBNR by Accident Year"
    chart1.type = "col"
    chart1.width = 12
    chart1.height = 8
    chart1.roundedCorners = False
    chart1.x_axis.title = "Accident Year"
    chart1.y_axis.title = "IBNR"
    chart1.y_axis.scaling.min = 0
    chart1.x_axis.tickLblPos = "low"
    chart1.y_axis.tickLblPos = "low"
    chart1.x_axis.delete = False
    chart1.y_axis.delete = False
    chart1.varyColors = False  # All bars same color
    data1 = Reference(ws, min_col=4, min_row=ibnr_start-1, max_row=ibnr_end)
    cats1 = Reference(ws, min_col=1, min_row=ibnr_start, max_row=ibnr_end)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.legend = None
    ws.add_chart(chart1, "L4")

    # Ultimates bar chart
    chart2 = BarChart()
    chart2.title = "Ultimate vs Latest by Accident Year"
    chart2.type = "col"
    chart2.grouping = "clustered"
    chart2.width = 12
    chart2.height = 8
    chart2.roundedCorners = False
    chart2.x_axis.title = "Accident Year"
    chart2.y_axis.title = "Amount"
    chart2.y_axis.scaling.min = 0
    chart2.x_axis.tickLblPos = "low"
    chart2.y_axis.tickLblPos = "low"
    chart2.x_axis.delete = False
    chart2.y_axis.delete = False
    data_latest = Reference(ws, min_col=2, min_row=ibnr_start-1, max_row=ibnr_end)
    data_ultimate = Reference(ws, min_col=3, min_row=ibnr_start-1, max_row=ibnr_end)
    chart2.add_data(data_latest, titles_from_data=True)
    chart2.add_data(data_ultimate, titles_from_data=True)
    cats2 = Reference(ws, min_col=1, min_row=ibnr_start, max_row=ibnr_end)
    chart2.set_categories(cats2)
    ws.add_chart(chart2, "L20")

    # Standard Error bar chart
    chart3 = BarChart()
    chart3.title = "Mack Standard Error by Accident Year"
    chart3.type = "col"
    chart3.width = 12
    chart3.height = 8
    chart3.roundedCorners = False
    chart3.x_axis.title = "Accident Year"
    chart3.y_axis.title = "Standard Error"
    chart3.y_axis.scaling.min = 0
    chart3.x_axis.tickLblPos = "low"
    chart3.y_axis.tickLblPos = "low"
    chart3.x_axis.delete = False
    chart3.y_axis.delete = False
    chart3.varyColors = False  # All bars same color
    data3 = Reference(ws, min_col=2, min_row=se_start-1, max_row=se_end)
    cats3 = Reference(ws, min_col=1, min_row=se_start, max_row=se_end)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.legend = None
    ws.add_chart(chart3, "L36")

    return ws


def create_copulas_sheet(wb):
    """Create the Copulas examples sheet."""
    ws = wb.create_sheet("Copulas")
    set_column_widths(ws, {'A': 18, 'B': 12, 'C': 12, 'D': 12, 'E': 25, 'F': 5})

    row = add_title(ws, "Student-t Copula")
    row = add_note(ws, "Generate correlated uniform random numbers for Monte Carlo simulation", row)
    row += 1

    # Correlation matrix - lower triangle has values, upper triangle links to lower, diagonal = 1
    row = add_note(ws, "CORRELATION MATRIX (3x3)", row)
    row = add_table_header(ws, ["", "X1", "X2", "X3"], row)
    corr_start_row = row

    # Row 1: X1 - diagonal=1, others are formulas linking to lower triangle
    ws.cell(row=row, column=1, value="X1").font = HEADER_FONT
    ws.cell(row=row, column=1).border = THIN_BORDER
    ws.cell(row=row, column=2, value=1.0).border = THIN_BORDER  # Diagonal
    ws.cell(row=row, column=3, value=f"=B{row+1}").border = THIN_BORDER  # Link to C6 -> B7
    ws.cell(row=row, column=4, value=f"=B{row+2}").border = THIN_BORDER  # Link to D6 -> B8
    row += 1

    # Row 2: X2 - lower triangle value, diagonal=1, upper links
    ws.cell(row=row, column=1, value="X2").font = HEADER_FONT
    ws.cell(row=row, column=1).border = THIN_BORDER
    ws.cell(row=row, column=2, value=0.9).border = THIN_BORDER  # Editable: corr(X2,X1) = 0.9
    corr_x2_x1_cell = f"B{row}"  # Store reference for chart title
    ws.cell(row=row, column=3, value=1.0).border = THIN_BORDER  # Diagonal
    ws.cell(row=row, column=4, value=f"=C{row+1}").border = THIN_BORDER  # Link to D7 -> C8
    row += 1

    # Row 3: X3 - lower triangle values, diagonal=1
    ws.cell(row=row, column=1, value="X3").font = HEADER_FONT
    ws.cell(row=row, column=1).border = THIN_BORDER
    ws.cell(row=row, column=2, value=0.3).border = THIN_BORDER  # Editable: corr(X3,X1)
    ws.cell(row=row, column=3, value=0.4).border = THIN_BORDER  # Editable: corr(X3,X2)
    ws.cell(row=row, column=4, value=1.0).border = THIN_BORDER  # Diagonal
    row += 1

    corr_end_row = row - 1

    # Add color scale conditional formatting (white to red) for correlation matrix
    corr_range_cf = f"B{corr_start_row}:D{corr_end_row}"
    ws.conditional_formatting.add(corr_range_cf,
        ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF',
                       end_type='num', end_value=1, end_color='FF6B6B'))

    row += 1

    corr_range = f"B{corr_start_row}:D{corr_end_row}"

    # Parameters
    row = add_note(ws, "PARAMETERS", row)
    ws.cell(row=row, column=1, value="Degrees of freedom:").border = THIN_BORDER
    ws.cell(row=row, column=2, value=5).border = THIN_BORDER
    df_cell = f"B{row}"
    row += 1
    ws.cell(row=row, column=1, value="Number of samples:").border = THIN_BORDER
    ws.cell(row=row, column=2, value=100).border = THIN_BORDER
    n_cell = f"B{row}"
    row += 1
    ws.cell(row=row, column=1, value="Random seed:").border = THIN_BORDER
    ws.cell(row=row, column=2, value=42).border = THIN_BORDER
    seed_cell = f"B{row}"
    row += 2

    # Output - use INDEX to extract individual values (all 100 samples for chart)
    row = add_note(ws, "GENERATED SAMPLES (use INDEX to extract from array)", row)
    row = add_note(ws, f"Formula: =INDEX(ACT_COPULA_STUDENT_T({corr_range}, {df_cell}, {n_cell}, {seed_cell}), row, col)", row)
    row = add_table_header(ws, ["Sample", "U1", "U2", "U3"], row)
    samples_start_row = row

    for i in range(100):  # All 100 samples for plotting
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        for j in range(3):
            ws.cell(row=row, column=j+2, value=f"=INDEX(ACT_COPULA_STUDENT_T({corr_range}, {df_cell}, {n_cell}, {seed_cell}), {i+1}, {j+1})").border = THIN_BORDER
        row += 1
    sample_end_row = row - 1

    # Create chart title cell with formula linking to correlation coefficient
    title_cell_row = 2  # Put title formula in row 2
    ws.cell(row=title_cell_row, column=5, value=f'="Copula Samples: U1 vs U2 (correlation = "&TEXT({corr_x2_x1_cell},"0.0")&")"')

    # Scatter chart with axes
    chart = ScatterChart()
    chart.roundedCorners = False
    # Link chart title to cell with dynamic formula
    chart.title = Title()
    chart.title.strRef = StrRef(f="'Copulas'!$E$2")
    chart.width = 10
    chart.height = 10
    chart.x_axis.title = "U1"
    chart.y_axis.title = "U2"
    chart.x_axis.scaling.min = 0
    chart.x_axis.scaling.max = 1
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "low"
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    from openpyxl.chart import Series
    xvalues = Reference(ws, min_col=2, min_row=samples_start_row, max_row=sample_end_row)
    yvalues = Reference(ws, min_col=3, min_row=samples_start_row, max_row=sample_end_row)
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker = Marker(symbol='circle', size=5)
    series.marker.graphicalProperties.solidFill = CHART_COLORS[0]  # All points same color
    series.graphicalProperties.line.noFill = True
    chart.series.append(series)
    chart.legend = None

    ws.add_chart(chart, "F4")

    row += 1
    row = add_note(ws, "Note: Results are uniform[0,1] values. Apply inverse CDF to get desired marginal distributions.", row)

    return ws


def create_return_period_sheet(wb):
    """Create the Return Period examples sheet."""
    ws = wb.create_sheet("Return Periods")
    set_column_widths(ws, {'A': 18, 'B': 18, 'C': 20, 'D': 30, 'E': 5})

    row = add_title(ws, "Return Period Functions")
    row = add_note(ws, "Generate losses from return period tables (OEP/AEP curves)", row)
    row += 1

    # Sample EP curve
    row = add_note(ws, "SAMPLE EXCEEDANCE PROBABILITY CURVE", row)
    row = add_table_header(ws, ["Return Period", "OEP Loss"], row)
    ep_data = [(10, 100000), (25, 250000), (50, 500000), (100, 1000000), (250, 2500000)]
    data_start = row
    for rp, loss in ep_data:
        ws.cell(row=row, column=1, value=rp).border = THIN_BORDER
        cell_loss = ws.cell(row=row, column=2, value=loss)
        cell_loss.border = THIN_BORDER
        cell_loss.number_format = '#,##0'
        row += 1
    data_end = row - 1
    row += 1

    rp_range = f"A{data_start}:A{data_end}"
    loss_range = f"B{data_start}:B{data_end}"

    # EP Curve chart with axes
    chart = ScatterChart()
    chart.title = "Exceedance Probability Curve"
    chart.width = 12
    chart.height = 8
    chart.roundedCorners = False
    chart.x_axis.title = "Return Period (years)"
    chart.y_axis.title = "OEP Loss"
    chart.y_axis.scaling.min = 0
    chart.x_axis.scaling.logBase = 10
    chart.x_axis.tickLblPos = "low"
    chart.y_axis.tickLblPos = "low"
    chart.x_axis.delete = False
    chart.y_axis.delete = False

    from openpyxl.chart import Series
    xvalues = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
    yvalues = Reference(ws, min_col=2, min_row=data_start, max_row=data_end)
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker = Marker(symbol='none')
    series.graphicalProperties.line.solidFill = CHART_COLORS[0]
    chart.series.append(series)
    chart.legend = None
    ws.add_chart(chart, "E4")

    # Interpolation examples
    row = add_note(ws, "LOSS INTERPOLATION", row)
    row = add_table_header(ws, ["Target RP", "Loss (LOG)", "Loss (LINEAR)", "Notes"], row)
    targets = [20, 75, 150, 200]
    for target in targets:
        log_formula = f'=ACT_RETURN_PERIOD_LOSS({rp_range}, {loss_range}, A{row}, "LOG")'
        lin_formula = f'=ACT_RETURN_PERIOD_LOSS({rp_range}, {loss_range}, A{row}, "LINEAR")'
        ws.cell(row=row, column=1, value=target).border = THIN_BORDER
        cell_log = ws.cell(row=row, column=2, value=log_formula)
        cell_log.border = THIN_BORDER
        cell_log.number_format = '#,##0'
        cell_lin = ws.cell(row=row, column=3, value=lin_formula)
        cell_lin.border = THIN_BORDER
        cell_lin.number_format = '#,##0'
        ws.cell(row=row, column=4, value="Log interpolation typical for cat curves").border = THIN_BORDER
        row += 1

    row += 1

    # AAL calculation
    row = add_note(ws, "AVERAGE ANNUAL LOSS (AAL) FROM OEP CURVE", row)
    ws.cell(row=row, column=1, value="AAL:").font = HEADER_FONT
    cell_aal = ws.cell(row=row, column=2, value=f"=ACT_AAL_FROM_OEP({rp_range}, {loss_range})")
    cell_aal.number_format = '#,##0'

    return ws


def strip_all_implicit_intersection(wb):
    """Robustly remove ALL @ symbols from formulas in the entire workbook."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    # Remove @ anywhere in the formula
                    original = cell.value
                    # Replace @ACT_ with ACT_
                    new_value = original.replace('@ACT_', 'ACT_')
                    # Replace =@ with =
                    new_value = new_value.replace('=@', '=')
                    # Replace any remaining @ followed by a function-like pattern
                    import re
                    new_value = re.sub(r'@([A-Za-z_][A-Za-z0-9_]*)\(', r'\1(', new_value)
                    if new_value != original:
                        cell.value = new_value


def update_all_functions_test_tab(wb):
    """Update the All Functions Test tab with correct function names."""
    if "All Functions Test" not in wb.sheetnames:
        return
    
    ws = wb["All Functions Test"]
    
    # Mapping of old function names to new ones
    replacements = {
        'ACT_POISSON_': 'ACT_DIST_POISSON_',
        'ACT_NEGBIN_': 'ACT_DIST_NEGBIN_',
        'ACT_LOGNORM_': 'ACT_DIST_LOGNORM_',
        'ACT_GAMMA_': 'ACT_DIST_GAMMA_',
        'ACT_PARETO_PDF': 'ACT_DIST_PARETO_PDF',
        'ACT_PARETO_CDF': 'ACT_DIST_PARETO_CDF',
        'ACT_PARETO_INV': 'ACT_DIST_PARETO_INV',
        'ACT_GPD_': 'ACT_DIST_GPD_',
        'ACT_WEIBULL_': 'ACT_DIST_WEIBULL_',
        'ACT_BETA_': 'ACT_DIST_BETA_',
        'ACT_EXP_': 'ACT_DIST_EXP_',
        'ACT_BURR_': 'ACT_DIST_BURR_',
        'ACT_MBBEFD': 'ACT_EXPOSURE_MBBEFD',
        'ACT_SWISSRE_CURVE': 'ACT_EXPOSURE_SWISSRE',
        'ACT_LLOYDS_CURVE': 'ACT_EXPOSURE_LLOYDS',
        'ACT_POWER_CURVE': 'ACT_EXPOSURE_POWER',
        'ACT_INVERSE_POWER_CURVE': 'ACT_EXPOSURE_INVERSE_POWER',
        'ACT_PARETO_EXPOSURE': 'ACT_EXPOSURE_PARETO',
        'ACT_RIEBESELL_CURVE': 'ACT_EXPOSURE_RIEBESELL',
        'ACT_LAYER_RATE_ON_LINE': 'ACT_EXPOSURE_LAYER_RATE',
        'ACT_BOOTSTRAP_CL_ORIGIN': 'ACT_CL_BOOTSTRAP_ORIGIN',
        'ACT_BOOTSTRAP_CL': 'ACT_CL_BOOTSTRAP',
        'ACT_AGGREGATE_LAYER': '',  # Remove
        'ACT_QS_CEDED': '',  # Remove
    }
    
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                new_value = cell.value
                for old, new in replacements.items():
                    if old in new_value:
                        if new == '':
                            # Remove rows with these functions
                            continue
                        new_value = new_value.replace(old, new)
                # Also strip @
                new_value = new_value.replace('@ACT_', 'ACT_').replace('=@', '=')
                if new_value != cell.value:
                    cell.value = new_value


def main():
    """Main function to populate the workbook."""
    print(f"Opening workbook: {EXCEL_PATH}")

    # Open existing workbook (preserve macros)
    wb = openpyxl.load_workbook(EXCEL_PATH, keep_vba=True)

    # First, strip @ from ALL existing sheets
    print("Stripping @ symbols from existing sheets...")
    strip_all_implicit_intersection(wb)

    # Update All Functions Test tab
    print("Updating All Functions Test tab...")
    update_all_functions_test_tab(wb)

    # Remove existing sheets we're going to recreate
    sheets_to_create = ["Distributions", "Exposure Curves", "Reinsurance", "Interpolation",
                        "Chain Ladder", "Copulas", "Return Periods"]
    for sheet in sheets_to_create:
        if sheet in wb.sheetnames:
            del wb[sheet]

    # Remove default sheet if it's empty
    if "Sheet1" in wb.sheetnames and wb["Sheet1"].max_row == 1:
        del wb["Sheet1"]

    # Create example sheets
    print("Creating Distributions sheet with ALL distributions and charts...")
    create_distributions_sheet(wb)

    print("Creating Exposure Curves sheet with charts...")
    create_exposure_curves_sheet(wb)

    print("Creating Reinsurance sheet...")
    create_reinsurance_sheet(wb)

    print("Creating Interpolation sheet with chart...")
    create_interpolation_sheet(wb)

    print("Creating Chain Ladder sheet with bar charts...")
    create_chainladder_sheet(wb)

    print("Creating Copulas sheet with scatter plot...")
    create_copulas_sheet(wb)

    print("Creating Return Periods sheet with EP curve...")
    create_return_period_sheet(wb)

    # Final strip of @ symbols
    print("Final cleanup of @ symbols...")
    strip_all_implicit_intersection(wb)

    # Save
    print(f"Saving workbook...")
    wb.save(EXCEL_PATH)
    print("Done!")


if __name__ == "__main__":
    main()
