#!/usr/bin/env python3
"""
Populate Excel workbook with add-in function examples.
Examples are consistent with the test suite in ActuarialAddIn.Tests.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, ScatterChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
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


def create_line_chart(ws, title, x_col, y_col, min_row, max_row, anchor,
                      x_title=None, y_title=None, width=10, height=7):
    """Create a line chart."""
    chart = LineChart()
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height

    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title

    # Data reference
    data = Reference(ws, min_col=y_col, min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=x_col, min_row=min_row+1, max_row=max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None

    ws.add_chart(chart, anchor)
    return chart


def create_multi_line_chart(ws, title, x_col, y_cols, min_row, max_row, anchor,
                            series_names=None, x_title=None, y_title=None, width=12, height=8):
    """Create a line chart with multiple series."""
    chart = LineChart()
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height

    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title

    cats = Reference(ws, min_col=x_col, min_row=min_row+1, max_row=max_row)

    for i, y_col in enumerate(y_cols):
        data = Reference(ws, min_col=y_col, min_row=min_row, max_row=max_row)
        chart.add_data(data, titles_from_data=True)

    chart.set_categories(cats)

    ws.add_chart(chart, anchor)
    return chart


def create_scatter_chart(ws, title, x_col, y_col, min_row, max_row, anchor,
                         x_title=None, y_title=None, width=10, height=7):
    """Create a scatter chart."""
    chart = ScatterChart()
    chart.title = title
    chart.style = 10
    chart.width = width
    chart.height = height

    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title

    xvalues = Reference(ws, min_col=x_col, min_row=min_row+1, max_row=max_row)
    yvalues = Reference(ws, min_col=y_col, min_row=min_row+1, max_row=max_row)

    from openpyxl.chart import Series
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker = Marker(symbol='circle', size=7)
    series.graphicalProperties.line.noFill = True
    chart.series.append(series)

    ws.add_chart(chart, anchor)
    return chart


def create_distributions_sheet(wb):
    """Create the Distributions examples sheet."""
    ws = wb.create_sheet("Distributions")
    set_column_widths(ws, {'A': 15, 'B': 20, 'C': 20, 'D': 20, 'E': 25, 'F': 5, 'G': 15})

    row = add_title(ws, "Statistical Distributions")
    row = add_note(ws, "PDF = Probability Density/Mass Function, CDF = Cumulative Distribution Function, INV = Inverse CDF (Quantile)", row)
    row += 1

    # Poisson
    poisson_start = row
    row = add_note(ws, "POISSON DISTRIBUTION (lambda=5)", row)
    row = add_table_header(ws, ["k", "PDF", "CDF", "Notes"], row)
    poisson_data_start = row
    for k in range(11):
        formulas = [None, f'=ACT_POISSON_PDF(A{row}, 5)', f'=ACT_POISSON_CDF(A{row}, 5)', None]
        notes = "Mode of distribution" if k == 5 else ""
        row = add_data_row(ws, [k, "", "", notes], row, formulas=formulas)
    poisson_data_end = row - 1

    row = add_note(ws, f"Inverse CDF example: =ACT_POISSON_INV(0.5, 5) returns the median", row)
    ws.cell(row=row-1, column=5, value="=ACT_POISSON_INV(0.5, 5)")
    row += 1

    # Poisson chart
    chart = LineChart()
    chart.title = "Poisson PMF (λ=5)"
    chart.style = 10
    chart.width = 10
    chart.height = 6
    chart.x_axis.title = "k"
    chart.y_axis.title = "P(X=k)"
    data = Reference(ws, min_col=2, min_row=poisson_data_start-1, max_row=poisson_data_end)
    cats = Reference(ws, min_col=1, min_row=poisson_data_start, max_row=poisson_data_end)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    ws.add_chart(chart, "F4")

    # Negative Binomial
    row = add_note(ws, "NEGATIVE BINOMIAL DISTRIBUTION (r=5, p=0.3)", row)
    row = add_table_header(ws, ["k", "PDF", "CDF", "Notes"], row)
    negbin_data_start = row
    for k in range(11):
        formulas = [None, f'=ACT_NEGBIN_PDF(A{row}, 5, 0.3)', f'=ACT_NEGBIN_CDF(A{row}, 5, 0.3)', None]
        row = add_data_row(ws, [k, "", "", ""], row, formulas=formulas)
    negbin_data_end = row - 1
    row += 1

    # Negative Binomial chart
    chart2 = LineChart()
    chart2.title = "Negative Binomial PMF (r=5, p=0.3)"
    chart2.style = 10
    chart2.width = 10
    chart2.height = 6
    chart2.x_axis.title = "k"
    chart2.y_axis.title = "P(X=k)"
    data2 = Reference(ws, min_col=2, min_row=negbin_data_start-1, max_row=negbin_data_end)
    cats2 = Reference(ws, min_col=1, min_row=negbin_data_start, max_row=negbin_data_end)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.legend = None
    ws.add_chart(chart2, "F18")

    # Lognormal
    row = add_note(ws, "LOGNORMAL DISTRIBUTION (mu=0, sigma=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    lognorm_data_start = row
    for x in [0.5, 1.0, 1.5, 2.0, 3.0, 5.0]:
        formulas = [None, f'=ACT_LOGNORM_PDF(A{row}, 0, 1)', f'=ACT_LOGNORM_CDF(A{row}, 0, 1)', None]
        notes = "Median (exp(mu))" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    lognorm_data_end = row - 1
    row += 1

    # Lognormal chart
    chart3 = LineChart()
    chart3.title = "Lognormal PDF (μ=0, σ=1)"
    chart3.style = 10
    chart3.width = 10
    chart3.height = 6
    chart3.x_axis.title = "x"
    chart3.y_axis.title = "f(x)"
    data3 = Reference(ws, min_col=2, min_row=lognorm_data_start-1, max_row=lognorm_data_end)
    cats3 = Reference(ws, min_col=1, min_row=lognorm_data_start, max_row=lognorm_data_end)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.legend = None
    ws.add_chart(chart3, "F32")

    # Gamma
    row = add_note(ws, "GAMMA DISTRIBUTION (alpha=2, beta=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    gamma_data_start = row
    for x in [0.5, 1.0, 2.0, 3.0, 5.0]:
        formulas = [None, f'=ACT_GAMMA_PDF(A{row}, 2, 1)', f'=ACT_GAMMA_CDF(A{row}, 2, 1)', None]
        notes = "Mean (alpha/beta)" if x == 2.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    gamma_data_end = row - 1
    row += 1

    # Pareto
    row = add_note(ws, "PARETO DISTRIBUTION (alpha=2, xm=1)", row)
    row = add_table_header(ws, ["x", "PDF", "CDF", "Notes"], row)
    pareto_data_start = row
    for x in [1.0, 1.5, 2.0, 3.0, 5.0, 10.0]:
        formulas = [None, f'=ACT_PARETO_PDF(A{row}, 2, 1)', f'=ACT_PARETO_CDF(A{row}, 2, 1)', None]
        notes = "Minimum value (xm)" if x == 1.0 else ""
        row = add_data_row(ws, [x, "", "", notes], row, formulas=formulas)
    pareto_data_end = row - 1

    # Pareto chart
    chart4 = LineChart()
    chart4.title = "Pareto PDF (α=2, xm=1)"
    chart4.style = 10
    chart4.width = 10
    chart4.height = 6
    chart4.x_axis.title = "x"
    chart4.y_axis.title = "f(x)"
    data4 = Reference(ws, min_col=2, min_row=pareto_data_start-1, max_row=pareto_data_end)
    cats4 = Reference(ws, min_col=1, min_row=pareto_data_start, max_row=pareto_data_end)
    chart4.add_data(data4, titles_from_data=True)
    chart4.set_categories(cats4)
    chart4.legend = None
    ws.add_chart(chart4, "F46")

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
        formulas = [None, f'=ACT_MBBEFD(A{row}, 2, 3)', None]
        row = add_data_row(ws, [d, "", ""], row, formulas=formulas)
    mbbefd_data_end = row - 1
    row += 1

    # MBBEFD chart
    chart = LineChart()
    chart.title = "MBBEFD Exposure Curve (b=2, g=3)"
    chart.style = 10
    chart.width = 12
    chart.height = 8
    chart.x_axis.title = "Damage Ratio (d)"
    chart.y_axis.title = "Loss Ratio G(d)"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1
    data = Reference(ws, min_col=2, min_row=mbbefd_data_start-1, max_row=mbbefd_data_end)
    cats = Reference(ws, min_col=1, min_row=mbbefd_data_start, max_row=mbbefd_data_end)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    ws.add_chart(chart, "F4")

    # Swiss Re curves comparison - build data for chart
    row = add_note(ws, "SWISS RE CURVES COMPARISON (all curves)", row)
    row = add_table_header(ws, ["d", "Curve 1", "Curve 2", "Curve 3", "Curve 4", "Curve 5"], row)
    swissre_data_start = row
    for d in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        ws.cell(row=row, column=1, value=d).border = THIN_BORDER
        for c in range(1, 6):
            ws.cell(row=row, column=c+1, value=f'=ACT_SWISSRE_CURVE(A{row}, {c})').border = THIN_BORDER
        row += 1
    swissre_data_end = row - 1
    row += 1

    # Swiss Re comparison chart
    chart2 = LineChart()
    chart2.title = "Swiss Re Curves Comparison"
    chart2.style = 10
    chart2.width = 12
    chart2.height = 8
    chart2.x_axis.title = "Damage Ratio (d)"
    chart2.y_axis.title = "Loss Ratio G(d)"
    chart2.y_axis.scaling.min = 0
    chart2.y_axis.scaling.max = 1
    data2 = Reference(ws, min_col=2, max_col=6, min_row=swissre_data_start-1, max_row=swissre_data_end)
    cats2 = Reference(ws, min_col=1, min_row=swissre_data_start, max_row=swissre_data_end)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    ws.add_chart(chart2, "F20")

    # Lloyd's curves
    row = add_note(ws, "LLOYD'S CURVES COMPARISON", row)
    row = add_table_header(ws, ["d", "Y1", "Y2", "Y3", "Y4"], row)
    lloyds_data_start = row
    for d in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]:
        ws.cell(row=row, column=1, value=d).border = THIN_BORDER
        for i, curve in enumerate(["Y1", "Y2", "Y3", "Y4"]):
            ws.cell(row=row, column=i+2, value=f'=ACT_LLOYDS_CURVE(A{row}, "{curve}")').border = THIN_BORDER
        row += 1
    lloyds_data_end = row - 1
    row += 1

    # Lloyd's chart
    chart3 = LineChart()
    chart3.title = "Lloyd's Curves Comparison"
    chart3.style = 10
    chart3.width = 12
    chart3.height = 8
    chart3.x_axis.title = "Damage Ratio (d)"
    chart3.y_axis.title = "Loss Ratio G(d)"
    chart3.y_axis.scaling.min = 0
    chart3.y_axis.scaling.max = 1
    data3 = Reference(ws, min_col=2, max_col=5, min_row=lloyds_data_start-1, max_row=lloyds_data_end)
    cats3 = Reference(ws, min_col=1, min_row=lloyds_data_start, max_row=lloyds_data_end)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    ws.add_chart(chart3, "F36")

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
    xol_start = row
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
    row += 1

    # Quota Share
    row = add_note(ws, "QUOTA SHARE (50% cession)", row)
    row = add_table_header(ws, ["Ground-Up Loss", "Ceded Loss", "Retained"], row)
    cell_a = ws.cell(row=row, column=1, value=1000000)
    cell_a.border = THIN_BORDER
    cell_a.number_format = '#,##0'
    cell_b = ws.cell(row=row, column=2, value="=ACT_QS_CEDED(A" + str(row) + ", 0.5)")
    cell_b.border = THIN_BORDER
    cell_b.number_format = '#,##0'
    cell_c = ws.cell(row=row, column=3, value="=A" + str(row) + "-B" + str(row))
    cell_c.border = THIN_BORDER
    cell_c.number_format = '#,##0'
    row += 2

    # Aggregate Layer
    row = add_note(ws, "AGGREGATE LAYER (Deductible=2,000,000, Limit=10,000,000)", row)
    row = add_table_header(ws, ["Aggregate Loss", "Layer Recovery", "Notes"], row)
    agg_losses = [1000000, 2000000, 5000000, 12000000, 15000000]
    agg_notes = ["Below deductible", "At deductible", "Partial recovery", "Full limit", "Capped at limit"]
    for loss, note in zip(agg_losses, agg_notes):
        cell_a = ws.cell(row=row, column=1, value=loss)
        cell_a.border = THIN_BORDER
        cell_a.number_format = '#,##0'
        cell_b = ws.cell(row=row, column=2, value=f'=ACT_AGGREGATE_LAYER(A{row}, 2000000, 10000000)')
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

    # Chart showing known points and interpolated values
    chart = ScatterChart()
    chart.title = "Interpolation with Extrapolation"
    chart.style = 10
    chart.width = 12
    chart.height = 8
    chart.x_axis.title = "X"
    chart.y_axis.title = "Y"

    # Known data points
    from openpyxl.chart import Series
    xvalues1 = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)
    yvalues1 = Reference(ws, min_col=2, min_row=data_start_row, max_row=data_end_row)
    series1 = Series(yvalues1, xvalues1, title="Known Points")
    series1.marker = Marker(symbol='circle', size=10)
    series1.graphicalProperties.line.noFill = True
    chart.series.append(series1)

    # FLAT extrapolation
    xvalues2 = Reference(ws, min_col=1, min_row=interp_start_row, max_row=interp_end_row)
    yvalues2 = Reference(ws, min_col=2, min_row=interp_start_row, max_row=interp_end_row)
    series2 = Series(yvalues2, xvalues2, title="FLAT Extrapolation")
    series2.marker = Marker(symbol='triangle', size=7)
    series2.graphicalProperties.line.noFill = True
    chart.series.append(series2)

    # GRADIENT extrapolation
    yvalues3 = Reference(ws, min_col=3, min_row=interp_start_row, max_row=interp_end_row)
    series3 = Series(yvalues3, xvalues2, title="GRADIENT Extrapolation")
    series3.marker = Marker(symbol='square', size=7)
    series3.graphicalProperties.line.noFill = True
    chart.series.append(series3)

    ws.add_chart(chart, "F4")

    return ws


def create_chainladder_sheet(wb):
    """Create the Chain Ladder examples sheet.

    Uses the Taylor-Ashe dataset from Peter England's bootstrapping presentation.
    This is the standard benchmark dataset for validating chain ladder implementations.
    Source: England & Verrall (1999), Taylor & Ashe (1983)
    """
    ws = wb.create_sheet("Chain Ladder")
    set_column_widths(ws, {'A': 10, 'B': 12, 'C': 12, 'D': 12, 'E': 12, 'F': 12,
                           'G': 12, 'H': 12, 'I': 12, 'J': 12, 'K': 12, 'L': 5, 'M': 15})

    row = add_title(ws, "Chain Ladder Reserving - Taylor-Ashe Dataset")
    row = add_note(ws, "Using the Taylor-Ashe triangle from Peter England's bootstrapping presentation (England & Verrall, 1999)", row)
    row += 1

    # Input triangle - Taylor-Ashe cumulative paid losses
    row = add_note(ws, "INPUT TRIANGLE (Cumulative Paid Losses - Taylor-Ashe)", row)
    row = add_table_header(ws, ["AY", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"], row)

    # Taylor-Ashe triangle data
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

    # Triangle range for formulas
    tri_range = f"B{triangle_start_row}:K{triangle_end_row}"

    # Development factors
    row = add_note(ws, "DEVELOPMENT FACTORS", row)
    row = add_note(ws, "Expected: 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018", row)
    row = add_table_header(ws, ["Period", "1-2", "2-3", "3-4", "4-5", "5-6", "6-7", "7-8", "8-9", "9-10"], row)
    factors_row = row
    ws.cell(row=row, column=1, value="Factors").border = THIN_BORDER
    ws.cell(row=row, column=2, value=f"=ACT_CL_FACTORS({tri_range})").border = THIN_BORDER
    row += 2

    # Ultimates and IBNR
    row = add_note(ws, "PROJECTED ULTIMATES AND IBNR", row)
    row = add_note(ws, "Expected Total IBNR: ~18,680,856", row)
    row = add_table_header(ws, ["AY", "Latest Cumulative", "Ultimate", "IBNR"], row)

    # Latest diagonal values
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

    # Mack Standard Errors
    row = add_note(ws, "MACK STANDARD ERRORS (reserve uncertainty)", row)
    row = add_table_header(ws, ["AY", "Reserve SE"], row)
    se_start = row
    for i in range(10):
        ws.cell(row=row, column=1, value=i+1).border = THIN_BORDER
        cell_se = ws.cell(row=row, column=2, value=f"=INDEX(ACT_MACK_RESERVE_SE({tri_range}), {i+1})")
        cell_se.border = THIN_BORDER
        cell_se.number_format = '#,##0'
        row += 1
    se_end = row - 1
    row += 1

    # Charts
    # IBNR by Accident Year
    chart1 = LineChart()
    chart1.title = "IBNR by Accident Year"
    chart1.style = 10
    chart1.width = 12
    chart1.height = 8
    chart1.x_axis.title = "Accident Year"
    chart1.y_axis.title = "IBNR"
    data1 = Reference(ws, min_col=4, min_row=ibnr_start-1, max_row=ibnr_end)
    cats1 = Reference(ws, min_col=1, min_row=ibnr_start, max_row=ibnr_end)
    chart1.add_data(data1, titles_from_data=True)
    chart1.set_categories(cats1)
    chart1.legend = None
    ws.add_chart(chart1, "L4")

    # Development pattern (cumulative to ultimate for each AY)
    chart2 = LineChart()
    chart2.title = "Development Pattern - Latest Diagonal to Ultimate"
    chart2.style = 10
    chart2.width = 12
    chart2.height = 8
    chart2.x_axis.title = "Accident Year"
    chart2.y_axis.title = "Amount"

    # Add latest and ultimate series
    data_latest = Reference(ws, min_col=2, min_row=ibnr_start-1, max_row=ibnr_end)
    data_ultimate = Reference(ws, min_col=3, min_row=ibnr_start-1, max_row=ibnr_end)
    chart2.add_data(data_latest, titles_from_data=True)
    chart2.add_data(data_ultimate, titles_from_data=True)
    cats2 = Reference(ws, min_col=1, min_row=ibnr_start, max_row=ibnr_end)
    chart2.set_categories(cats2)
    ws.add_chart(chart2, "L20")

    # Standard Error chart
    chart3 = LineChart()
    chart3.title = "Mack Standard Error by Accident Year"
    chart3.style = 10
    chart3.width = 12
    chart3.height = 8
    chart3.x_axis.title = "Accident Year"
    chart3.y_axis.title = "Standard Error"
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
    ws.cell(row=row, column=2, value=100).border = THIN_BORDER  # More samples for better chart
    n_cell = f"B{row}"
    row += 1
    ws.cell(row=row, column=1, value="Random seed:").border = THIN_BORDER
    ws.cell(row=row, column=2, value=42).border = THIN_BORDER
    seed_cell = f"B{row}"
    row += 2

    # Output
    row = add_note(ws, "GENERATED SAMPLES (uniform marginals, correlated via t-copula)", row)
    row = add_note(ws, f"Formula: =ACT_STUDENT_T_COPULA({corr_range}, {df_cell}, {n_cell}, {seed_cell})", row)
    row = add_table_header(ws, ["U1", "U2", "U3"], row)
    samples_start_row = row

    # The formula will spill results
    ws.cell(row=row, column=1, value=f"=ACT_STUDENT_T_COPULA({corr_range}, {df_cell}, {n_cell}, {seed_cell})").border = THIN_BORDER

    # Scatter chart showing correlation between U1 and U2
    chart = ScatterChart()
    chart.title = "Copula Samples: U1 vs U2 (correlation = 0.5)"
    chart.style = 10
    chart.width = 10
    chart.height = 10
    chart.x_axis.title = "U1"
    chart.y_axis.title = "U2"
    chart.x_axis.scaling.min = 0
    chart.x_axis.scaling.max = 1
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 1

    # Reference for 100 samples
    from openpyxl.chart import Series
    xvalues = Reference(ws, min_col=1, min_row=samples_start_row, max_row=samples_start_row+99)
    yvalues = Reference(ws, min_col=2, min_row=samples_start_row, max_row=samples_start_row+99)
    series = Series(yvalues, xvalues, title_from_data=False)
    series.marker = Marker(symbol='circle', size=5)
    series.graphicalProperties.line.noFill = True
    chart.series.append(series)
    chart.legend = None

    ws.add_chart(chart, "F4")

    row += 102
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

    # EP Curve chart
    chart = LineChart()
    chart.title = "Exceedance Probability Curve"
    chart.style = 10
    chart.width = 12
    chart.height = 8
    chart.x_axis.title = "Return Period (years)"
    chart.y_axis.title = "OEP Loss"
    chart.x_axis.scaling.logBase = 10  # Log scale for return period
    data = Reference(ws, min_col=2, min_row=data_start-1, max_row=data_end)
    cats = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
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


def main():
    """Main function to populate the workbook."""
    print(f"Opening workbook: {EXCEL_PATH}")

    # Open existing workbook (preserve macros)
    wb = openpyxl.load_workbook(EXCEL_PATH, keep_vba=True)

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
    print("Creating Distributions sheet with charts...")
    create_distributions_sheet(wb)

    print("Creating Exposure Curves sheet with charts...")
    create_exposure_curves_sheet(wb)

    print("Creating Reinsurance sheet...")
    create_reinsurance_sheet(wb)

    print("Creating Interpolation sheet with chart...")
    create_interpolation_sheet(wb)

    print("Creating Chain Ladder sheet...")
    create_chainladder_sheet(wb)

    print("Creating Copulas sheet with scatter plot...")
    create_copulas_sheet(wb)

    print("Creating Return Periods sheet with EP curve...")
    create_return_period_sheet(wb)

    # Save
    print(f"Saving workbook...")
    wb.save(EXCEL_PATH)
    print("Done!")


if __name__ == "__main__":
    main()
