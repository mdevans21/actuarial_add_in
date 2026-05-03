#!/usr/bin/env python3
"""Fix all formulas using xlwings Formula2 (dynamic arrays).

Run from Windows Python (not WSL) after populate_examples.py to remove the
@ implicit intersection prefix. openpyxl doesn't write Excel 365 dynamic
array metadata, so all formulas get @ when opened. This script re-sets
every formula using Formula2 to upgrade them.

Array-returning formulas in the All Functions Test tab are handled specially:
- FIT formulas (multi-param): TRANSPOSE + parameter labels + spacing
- FIT formulas (single-param): plain formula2 + parameter label
- Other array formulas (copulas, tables): wrapped in INDEX() to show first value

Requires: pip install xlwings
Requires: Excel must be closed before running.
"""
import os
import re
import xlwings as xw

EXCEL_PATH = os.path.join(os.path.dirname(__file__), '..', 'excel', 'actuarial_add_in.xlsx')
EXCEL_PATH = os.path.abspath(EXCEL_PATH)

# FIT function parameter names (from C# source)
FIT_PARAMS = {
    'EXP_FIT':     ['lambda'],
    'POISSON_FIT': ['lambda'],
    'LOGNORM_FIT': ['mu', 'sigma'],
    'GAMMA_FIT':   ['alpha (shape)', 'beta (rate)'],
    'PARETO_FIT':  ['alpha', 'xm'],
    'WEIBULL_FIT': ['k (shape)', 'lambda (scale)'],
    'GPD_FIT':     ['xi (shape)', 'sigma (scale)'],
    'BETA_FIT':    ['alpha', 'beta'],
    'NEGBIN_FIT':  ['r (successes)', 'p (probability)'],
    'BURR_FIT':    ['c', 'k', 'lambda'],
}

# Non-FIT formulas that return arrays
ARRAY_FUNCS = [
    'COPULA_GAUSSIAN(', 'COPULA_GAUSSIAN_SINGLE(',
    'COPULA_STUDENT_T(', 'COPULA_STUDENT_T_SINGLE(',
    'COPULA_CLAYTON(', 'COPULA_CLAYTON_SINGLE(',
    'COPULA_FRANK(', 'COPULA_FRANK_SINGLE(',
    'COPULA_GUMBEL(', 'COPULA_GUMBEL_SINGLE(',
    'RETURN_PERIOD_TABLE(', 'COMMIT_HISTORY(',
    'DISCRETIZE_EXPONENTIAL(', 'DISCRETIZE_GAMMA(', 'DISCRETIZE_LOGNORMAL(',
    'PANJER_POISSON(', 'PANJER_NEGBIN(', 'PANJER_BINOMIAL(',
    'CAT_ELT_TO_YLT(', 'CAT_YLT_OEP_CURVE(', 'CAT_YLT_AEP_CURVE(',
    'CAT_OEP_CURVE_RP(', 'CAT_AEP_CURVE_RP(',
]


def get_fit_key(formula):
    """Extract the FIT function key from a formula, e.g. 'NEGBIN_FIT'."""
    m = re.search(r'ACT_DIST_(\w+_FIT)\(', formula)
    return m.group(1) if m else None


def strip_wrappers(formula):
    """Strip existing TRANSPOSE() and INDEX() wrappers to get the raw formula.

    Handles nested wrappers from previous runs, e.g.:
    =TRANSPOSE(TRANSPOSE(ACT_DIST_FIT(...))) -> =ACT_DIST_FIT(...)
    =INDEX(ACT_COPULA(...), 1, 1) -> =ACT_COPULA(...)
    """
    f = formula
    changed = True
    while changed:
        changed = False
        # Strip =TRANSPOSE(...) or =@TRANSPOSE(...)
        m = re.match(r'^=@?TRANSPOSE\((.+)\)$', f)
        if m:
            f = '=' + m.group(1)
            changed = True
        # Strip =INDEX(..., 1, 1) or =@INDEX(..., 1, 1)
        m = re.match(r'^=@?INDEX\((.+),\s*1,\s*1\)$', f)
        if m:
            f = '=' + m.group(1)
            changed = True
    return f


def is_array_formula(formula):
    """Check if formula is a known array-returning function (non-FIT)."""
    return any(af in formula for af in ARRAY_FUNCS)


def fix_all_functions_test(ws):
    """Handle array-returning formulas in the All Functions Test tab.

    For FIT formulas:
    - Single-param: plain formula2, add parameter name in column C
    - Multi-param: TRANSPOSE() wrapper, clear column C for spill,
      add parameter names in next column after spill, add formula label in A,
      ensure blank row after last parameter
    For other array formulas:
    - Wrap in INDEX(..., 1, 1) to show first value without spill
    """
    used = ws.used_range
    if used is None:
        return 0

    # First pass: collect all formulas and identify FIT cells
    fit_cells = []  # (row, formula, fit_key, params)
    other_cells = []  # (row, col_letter, formula)
    regular_cells = []  # (row, col_letter, formula)

    for cell in used:
        f = cell.formula
        if not isinstance(f, str) or not f.startswith('='):
            continue

        clean = f.replace('=@', '=', 1) if f.startswith('=@') else f
        clean = strip_wrappers(clean)  # remove TRANSPOSE/INDEX from previous runs
        # Get column letter from address like $B$9
        parts = cell.address.replace('$', '')
        col_letter = re.match(r'[A-Z]+', parts).group()

        fit_key = get_fit_key(clean)
        if fit_key and fit_key in FIT_PARAMS:
            params = FIT_PARAMS[fit_key]
            fit_cells.append((cell.row, clean, fit_key, params))
        elif is_array_formula(clean):
            other_cells.append((cell.row, col_letter, clean))
        else:
            regular_cells.append((cell.row, col_letter, clean))

    fixed = 0

    # Handle regular formulas
    for row, col, formula in regular_cells:
        ws.range(f'{col}{row}').formula2 = formula
        fixed += 1

    # Handle non-FIT array formulas: wrap in INDEX
    for row, col, formula in other_cells:
        inner = formula[1:]  # strip =
        ws.range(f'{col}{row}').formula2 = f'=INDEX({inner}, 1, 1)'
        fixed += 1

    # Handle FIT formulas - process bottom-up so row inserts don't shift earlier ones
    fit_cells.sort(key=lambda x: x[0], reverse=True)
    for row, formula, fit_key, params in fit_cells:
        n_params = len(params)

        # Put formula text as label in column A
        formula_text = formula[1:]  # strip leading =
        ws.range(f'A{row}').value = formula_text

        if n_params == 1:
            # Single param: just formula2, no TRANSPOSE needed
            ws.range(f'B{row}').formula2 = formula
            # Clear C if it has content, then add param name
            ws.range(f'C{row}').value = params[0]
            fixed += 1
        else:
            # Multi-param: TRANSPOSE so params display vertically
            ws.range(f'B{row}').formula2 = f'=TRANSPOSE({formula_text})'
            # Clear column C for the spill range
            for i in range(n_params):
                ws.range(f'C{row + i}').clear_contents()

            # Add parameter names in column after spill (column C, offset by spill)
            # Since TRANSPOSE makes it vertical in column B, C is free for labels
            for i, param_name in enumerate(params):
                ws.range(f'C{row + i}').value = param_name

            # Ensure blank row after last parameter
            last_param_row = row + n_params - 1
            next_row = last_param_row + 1
            next_cell = ws.range(f'A{next_row}')
            if next_cell.value is not None or (isinstance(next_cell.formula, str) and next_cell.formula.startswith('=')):
                # Need to insert a row to create spacing
                ws.range(f'{next_row}:{next_row}').insert('down')

            fixed += 1

    return fixed


def fix_regular_sheet(ws):
    """Convert all formulas to formula2 on a regular sheet."""
    used = ws.used_range
    if used is None:
        return 0

    fixed = 0
    for cell in used:
        f = cell.formula
        if isinstance(f, str) and f.startswith('='):
            clean = f.replace('=@', '=', 1) if f.startswith('=@') else f
            cell.formula2 = clean
            fixed += 1

    return fixed


print(f"Opening: {EXCEL_PATH}")
app = xw.App(visible=False)
try:
    wb = app.books.open(EXCEL_PATH)

    total_fixed = 0
    for ws in wb.sheets:
        if ws.name == 'All Functions Test':
            fixed = fix_all_functions_test(ws)
        else:
            fixed = fix_regular_sheet(ws)

        if fixed:
            print(f"  {ws.name}: {fixed} formulas fixed")
            total_fixed += fixed

    wb.save()
    wb.close()
    print(f"Done! {total_fixed} formulas upgraded to Formula2 (no @ prefix).")
finally:
    app.quit()
