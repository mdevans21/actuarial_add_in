# Changelog

All notable changes to the Actuarial Add-In are documented here. Dates are
ISO-8601. Versions follow [SemVer](https://semver.org/).

## [Unreleased]

## [0.7.2] — 2026-05-04

Workbook-fix release: every issue surfaced by the first complete v0.7.1
dump test plus a number of latent issues the dump exposed (Mack SE
reference column, Cat Modeling spill, smooth-line charts, hard-coded
"Latest Cumulative" column, missing Gaussian / Clayton / Frank / Gumbel
copula sections). No numerical changes to any add-in function — just
workbook content + chart layout + Excel-side metadata.

### Fixed
- **Mack SE reference column was wrong.** The Chain Ladder sheet had a
  hard-coded "E&V Reference" column with values
  `[0, 75535, 121699, 136554, 212054, 282372, 444283, 640396, 1091089, 2029818]`
  totalling `2,447,318`. Re-implementing Mack 1993 §3 from scratch in
  Python on the Taylor-Ashe triangle yields
  `[0, 75535, 121699, 133549, 261406, 411010, 558317, 875328, 971258, 1363155]`
  totalling `2,447,095`, which our C# `ACT_MACK_RESERVE_SE` returns
  bit-exact. Replaced the column values + relabelled "E&V Reference"
  → "Mack 1993 reference".
- **Cat Modeling: `=ACT_CAT_ELT_TO_YLT(...)` not spilling.** openpyxl
  writes formulas via `.Formula`, which Excel `@`-prefixes for legacy
  scalar behaviour on save — the YLT collapsed to a single cell.
  Added Formula2 promotion in `dump_workbook.ps1 -SaveWorkbook`:
  every array-returning ACT_* call (583 cells in v0.7.1, 970 in v0.7.2
  after the new Copula sections) is rewritten via `.Formula2` so
  Excel re-detects the dynamic-array spill. Same fix unblocks Chain
  Ladder bootstrap percentiles and Cat Modeling YLT.
- **Cat Modeling: "EP Curves at Return Periods" table empty.** The
  RP column was populated but the OEP / AEP cells were blank.
  Added `ACT_CAT_OEP_CURVE_RP` / `ACT_CAT_AEP_CURVE_RP` formulas.
- **Chain Ladder: "Latest Cumulative" column hard-coded.** Used to be
  10 constant values; restored as
  `=INDEX(ACT_CL_LATEST(B7:K16), N)` so triangle edits propagate.
- **Charts: smooth lines → straight lines.** Smooth scatter rendering
  imposes unnatural curves over discrete data (exposure curves, EP
  curves, etc). `dump_workbook.ps1 -SaveWorkbook` now sets
  `Series.Smooth = False` on every series of every embedded chart
  (38 series across the workbook).
- **Dump CVErr decoding silently swallowed every Excel error.** Excel
  COM `Range.Value2` returns the seven CVErr codes typed as `[double]`
  in PowerShell; the previous code only attempted lookup when the
  declared type was `[int]`, so all 373 v0.7.0-era `#VALUE!` cells
  were logged as numeric `-2146826273` and the post-dump scan saw
  zero errors. Format-CellValue now decodes whenever the value is a
  finite integer-valued number, regardless of declared type, and
  uses `[int32]` consistently with PowerShell hashtable keys.

### Added
- **Copulas sheet: Gaussian / Clayton / Frank / Gumbel sections.**
  Previously only Student-t had its own visualization block; now each
  Archimedean copula plus Gaussian gets a parameter input area, a
  50-row sample table, and a U1×U2 scatter chart. Per the alignment
  principle every package function should be exercised in the
  spreadsheet.
- **Interpolation chart: interpolated overlay.** The chart used to
  show only the input data; now overlays the GRADIENT-extrapolated
  series in a contrasting colour with diamond markers so the
  interpolation visibly tracks the input curve.
- **`SeedUtil` + Formula2 / chart-smoothness post-processing**
  (continuation from v0.7.1): driven by the new `-SaveWorkbook` flag
  on `dump_workbook.ps1`, which now also handles array-formula
  promotion and chart line-style cleanup.
- **`run_dump.sh LOCAL_WORKBOOK=1`**: workbook iteration without
  shipping a fresh release tag for each tweak — uses the in-repo
  `excel/actuarial_add_in.xlsx` against the released XLL.

### Changed (Excel function browser)
- **Cat Modeling**, **Aggregate Claims**, and **Return Period**
  functions retagged from `Actuarial.Reinsurance` /
  `Actuarial.Aggregate` to **`Actuarial.Experimental`** — matching
  the existing treatment of Copulas and ODP Bootstrap.
  - `ACT_CAT_ELT_TO_YLT`, `ACT_CAT_*_CURVE`, `ACT_CAT_*_CURVE_RP`,
    `ACT_VAR_FROM_SAMPLES`, `ACT_TVAR_FROM_SAMPLES` (7 functions).
  - `ACT_AGGREGATE_*`, `ACT_DISCRETIZE_*`, `ACT_PANJER_*` (12).
  - `ACT_RETURN_PERIOD_LOSS`, `ACT_RETURN_PERIOD_TABLE`,
    `ACT_AAL_FROM_OEP` (3).
  README §Experimental updated to cover the five experimental
  families (was two).

## [0.7.1] — 2026-05-04

### Fixed
- **`int? seed` parameters returned `#VALUE!` from Excel.** Every
  stochastic function declared the optional seed as `int? seed = null`
  (10 copula functions + 2 bootstrap variants + 1 cat ELT_TO_YLT, 13
  in total). The C# harness — which calls these directly via MathNet
  / dotnet — was happy with that signature, but ExcelDna 1.9.0
  mis-marshalled `int?` from Excel-side numbers and *every* cell
  calling these functions returned `#VALUE!`. The first complete
  workbook dump under v0.7.0 surfaced 373 such cells (300 of them
  Student-t copula INDEX cells on the Copulas sheet, 61 bootstrap
  cells on Chain Ladder, plus the 12 demo cells on the All Functions
  Test sheet).
  - Replaced the registered type with `object seed = null` and
    added `Functions/SeedUtil.cs` with a `ResolveSeed(object)` helper
    that handles `null`, `ExcelMissing`, `ExcelEmpty`, `double`,
    `int`, and `long` uniformly. Function bodies use
    `SeedUtil.ResolveSeed(seed) is { } _seed ? new Random(_seed) : new Random();`.
  - Linux harness output is bit-exact identical against v0.7.0
    (zero diffs across all 547 records) — the marshalling fix is
    Excel-side only.

### Tooling
- **Dump script now decodes Excel CVErr codes properly.** PowerShell
  COM hands `Range.Value2` back as `[double]` for some cell types,
  including the seven Excel error codes (e.g. `#VALUE!` arrives as
  `-2146826273` typed as double, not int). The previous code only
  caught int-typed CVErr lookups, so the 373 `#VALUE!` cells in the
  v0.7.0 dump were silently logged as numeric `-2146826273` —
  invisible to the post-dump error scan. `Format-CellValue` now does
  the lookup whenever the value is a finite integer-valued number,
  and decodes to the human-readable name (`#VALUE!`, `#NUM!`, etc).

## [0.7.0] — 2026-05-04

Coverage-alignment release. The four test surfaces — C# definitions,
Excel workbook formulas, harness records, notebook reconciliations —
now contain **exactly the same 174 functions**. A coverage assertion
in the notebook fails the build if any function is missing from any
surface.

### Removed (breaking)
- **Cape Cod functions:** `ACT_CAPECOD_ULTIMATE`, `ACT_CAPECOD_ELR`.
- **Triangle utilities:** `ACT_TRIANGLE_TO_INCREMENTAL`,
  `ACT_INCREMENTAL_TO_CUMULATIVE`, `ACT_TRIANGLE_DIAGONAL`,
  `ACT_TRIANGLE_LINK_RATIOS`.
- **Other CL helpers:** `ACT_CL_WEIGHTED_AVERAGE`,
  `ACT_CL_CALENDAR_ADJUST`, `ACT_CL_CALENDAR_TOTALS`.
  (`ACT_CL_LATEST` retained.)
- File `ChainLadderCapeCod.cs` deleted.

### Added
- **Spreadsheet:** new sections on the *Chain Ladder* sheet exercising
  `ACT_CL_LATEST` (latest diagonal), `ACT_MACK_FACTOR_SE` (factor SEs),
  and `ACT_BF_ULTIMATE` (Bornhuetter-Ferguson). The Windows dump test
  now hits these on round-trip.
- **Harness:** 17 previously-untested functions emit records — the 6
  ZT/ZM `_INV` quantile functions, 5 single-row copula variants
  (`*_SINGLE`), the three scalar copula helpers (`TAU_TO_THETA`,
  `TAIL_LOWER`, `TAIL_UPPER` — now using their real names instead of
  per-copula synthetic split keys), `ACT_CL_BOOTSTRAP` /
  `ACT_CL_BOOTSTRAP_ORIGIN` (the underlying matrix-returning
  variants), `ACT_INTERP` (scalar), and the version-metadata family
  (`ACT_VERSION`, `ACT_BUILD_DATE`, `ACT_GITHUB_URL`,
  `ACT_COMMIT_*`).
- **Notebook reconciliations** added against external references:
  - **Aggregate:** `ACT_AGGREGATE_VAR/TVAR/CDF` against numpy cumsum
    on the Panjer PMF; `ACT_DISCRETIZE_EXP/GAMMA/LOGNORMAL` against
    scipy CDF differences (Klugman 9.6 mass-dispersal); `ACT_AAL_FROM_OEP`
    against the trapezoidal-EP integral.
  - **Cat modelling:** `ACT_CAT_YLT_OEP_CURVE`, `_AEP_CURVE`,
    `_OEP_CURVE_RP`, `_AEP_CURVE_RP` checked for monotonic, finite
    non-negative loss vs return-period.
  - **Chain ladder:** `ACT_CL_LATEST`, `ACT_CL_ULTIMATE`,
    `ACT_BF_ULTIMATE`, `ACT_MACK_FACTOR_SE` against numpy on the
    Taylor-Ashe triangle (E&V 2002 / Mack 1993).
  - **Copulas (experimental):** `ACT_COPULA_FRANK_CDF` against the
    analytic Frank generator; `ACT_COPULA_TAU_TO_THETA` (Clayton /
    Gumbel closed-form, Frank via Debye D₁ round-trip);
    `ACT_COPULA_TAIL_LOWER/UPPER` against closed-form λ formulae for
    each copula type; `_SINGLE` variants checked for uniform marginal
    range; bulk samples checked for column-mean ≈ 0.5.
  - **Exposure curves:** `ACT_EXPOSURE_MBBEFD`, `_SWISSRE`,
    `_RIEBESELL`, `_RIEBESELL_INV`, `_LAYER_RATE` against Bernegger
    (1997) formulae; SwissRe c=5 now reconciles (was NaN before
    v0.5.4).
  - **Reinsurance:** `ACT_ILF_PARETO`, `ACT_XOL_EXPECTED_LOSS`,
    `ACT_RETURN_PERIOD_TABLE` against analytic Pareto LEV / row-wise
    interp.
  - **Interpolation:** `ACT_INTERP` (FLAT + GRADIENT extrapolation),
    `ACT_INTERP_LOG` (log-linear in x), `ACT_INTERP2D` against
    `scipy.interpolate.RegularGridInterpolator`.
  - **Distributions:** ZT/ZM `_INV` quantile functions reconciled via
    inverse-CDF rescaling; `ACT_DIST_LNPARETO_ALPHA` against the
    Scollnik (2007) hazard ratio.

### Changed
- **Notebook coverage gate.** A new final cell asserts every harness
  function name (less six trivial metadata exemptions) appears in
  `ALL_RESULTS`. Adding a new C# function without wiring it through
  the notebook now fails the build.
- **Error-string standardisation, round 2.** 30 more sites in
  `Fitting.cs` and `ChainLadder.cs` that returned literal
  `new object[] { "Error: …" }` now return
  `new object[] { ExcelError.ExcelErrorValue }`. Excel renders
  `#VALUE!` instead of a text cell on bad input.
- **Harness call fixes:** `ACT_INTERP_FLAT` / `_GRADIENT` synthetic
  records collapsed to `ACT_INTERP` (real function name; selector in
  args). Same for `ACT_CL_BOOTSTRAP_EV/CLP`,
  `ACT_CL_BOOTSTRAP_ORIGIN_EV/CLP`,
  `ACT_PANJER_POISSON_EXP`, and the per-copula
  `ACT_COPULA_TAU_TO_THETA_*` / `ACT_COPULA_TAIL_LOWER_*` /
  `ACT_COPULA_TAIL_UPPER_*` keys.

## [0.6.0] — 2026-05-04

Internal cleanups deferred from v0.5.0; no numerical changes to any
add-in function. The validated v0.5.4 outputs remain bit-identical
under v0.6.0 (cross-checked via the C# harness on 515 records).

### Changed
- **Split `Distributions.cs` (1686 lines) into four focused
  partial-class files**: `DistributionsDiscrete.cs` (Poisson, NegBin,
  ZT/ZM variants), `DistributionsContinuous.cs` (Lognormal, Gamma,
  Pareto I-IV, GPD, Weibull, Beta, Exponential, Burr, Normal, Lomax,
  Inverse Gaussian, Loglogistic), `DistributionsComposite.cs`
  (Lognormal-Pareto, Exponential-Pareto, Power-Pareto), and
  `DistributionsLEV.cs` (limited-expected-value functions). All
  remain in `static partial class Distributions` so the `ACT_DIST_*`
  Excel surface is unchanged.
- **Split `ChainLadder.cs` (1510 lines) into four files**:
  `ChainLadder.cs` (basic chain ladder, triangle utilities, calendar
  adjustments), `ChainLadderMack.cs` (~700 lines of Mack standard-
  error machinery), `ChainLadderBootstrap.cs` (ODP bootstrap), and
  `ChainLadderCapeCod.cs`. Same partial-class pattern.
- **Standardised inline error returns to `ExcelError.ExcelErrorValue`.**
  24 sites in `ChainLadder*.cs`, `Copulas.cs`, and `Fitting.cs` that
  returned literal `"Error: ..."` strings now return the proper
  `ExcelError.ExcelErrorValue` enum, which Excel renders as
  `#VALUE!` rather than as a text cell. Functions returning `double`
  continue to return `double.NaN` for math-domain failures (Excel
  renders this as `#NUM!`).

### Fixed
- **Test harness now exits non-zero on assertion failures.** Previous
  behaviour was to write `FALSE` rows into the markdown summary but
  always exit `0`, hiding regressions from CI. `Program.Main` now
  returns an `int` and tracks pass/fail counts via `FormatMatch`,
  emitting a summary line and exiting `1` if any assertion failed.
- **Cape Cod harness call had argument order swapped.**
  `AddinOutputsEmitter` invoked `ACT_CAPECOD_ULTIMATE` /
  `ACT_CAPECOD_ELR` as `(triangle, factors, premium)` but the
  signatures are `(triangle, premium, factors)`. Both fixture cells
  previously emitted error strings instead of real Cape Cod outputs;
  now they emit the proper ELR (~0.889 on the genins triangle) and
  per-origin ultimates.

## [0.5.4] — 2026-05-04

### Fixed
- **`ACT_EXPOSURE_MBBEFD` numerical stability.** The Bernegger formula
  `G(d) = ln((a + b^d)/(a + 1)) / ln((a + b)/(a + 1))` with
  `a = (g - b^g)/(b^g - 1)` failed in two regimes when computed
  literally:
  - **`b < 1`, large `g` (Swiss Re curve 5: `b≈0.247`, `g≈992`):**
    `b^g` underflows to 0, `a` collapses to `-g`, and the log arguments
    go negative — function returned `NaN` instead of `0.668...`.
  - **`b > 1`, large `g` (Swiss Re curve 3):** catastrophic
    cancellation in `a + 1 ≈ -1 + 1` cost ~3 decimals of precision
    (returned `0.97110` instead of the true `0.97133`).
  - Rewrote as `(a + x)/(a + 1) = 1 + (x - 1)(b^g - 1)/(g - 1)`, which
    never forms `a` explicitly and stays well-conditioned. Added a
    log-space fallback for the rare case of `b^g` overflow. Cross-
    checked the new formula against `mpmath` at 80 decimal digits
    across `c ∈ {1.5, 2, 3, 4, 5}` and `d ∈ [0.001, 0.999]` — worst-
    case relative error is `7.5e-11` (vs the prior `NaN` for c=5).
- **`ACT_EXPOSURE_SWISSRE` curve 5** now returns finite values across
  the full domain (was `NaN` everywhere except `d ∈ {0, 1}`).

## [0.5.3] — 2026-05-04

### Fixed
- **Revert target framework `net8.0-windows` → `net6.0-windows`.**
  Three different `ExcelDna.AddIn` versions (1.9.0, 1.10-preview4, plus
  the prior 1.7.0) all failed to host .NET 8 in Excel on a target
  machine that had `Microsoft.WindowsDesktop.App 8.0.26` installed —
  Excel kept reporting *"initialization failed; could not find runtime
  8.0.0.0 ... file not found"* on add-in load. The published
  `.runtimeconfig.json` was correct; removing the legacy
  `RuntimeVersion="v4.0"` from the .dna manifest was necessary but
  insufficient. Falling back to .NET 6 (which v0.4.0 used and which
  worked end-to-end) restores a known-working state.
- `ExcelDna.AddIn` stays at **1.9.0** so the `int?` nullable seed
  parameter (introduced in v0.5.0) keeps working.
- Documentation, install scripts, and CI workflows reverted to .NET 6
  references throughout.

  Future revisit: bump back to .NET 8 once a stable `ExcelDna.AddIn`
  release ships with a host that reliably finds the .NET 8 Desktop
  Runtime.

## [0.5.2] — 2026-05-04

### Fixed
- Bump `ExcelDna.AddIn` from 1.9.0 to **1.10.0-preview4**. The 1.9.0
  bundled host doesn't reliably locate the .NET 8 Desktop Runtime even
  when it's installed — Excel reports "initialization failed; could
  not find runtime 8.0.0.0 ... file not found" on add-in load. v0.5.1
  removed the legacy `RuntimeVersion="v4.0"` attribute from the .dna
  manifest (which was necessary but not sufficient); 1.10-preview4
  ships an updated host that actually finds .NET 8 via the SDK-style
  `runtimeconfig.json`.

  This is a preview package; fold back to a stable 1.10 release when
  one ships.

## [0.5.1] — 2026-05-03

### Fixed
- Drop the legacy `RuntimeVersion="v4.0"` attribute from
  `ActuarialAddIn-AddIn.dna`. With that attribute present, Excel-DNA
  treats the add-in as a .NET Framework 4.x assembly and tries to host
  it under the .NET Framework runtime — even though the csproj targets
  `net8.0-windows` and the published `.runtimeconfig.json` requests
  Microsoft.WindowsDesktop.App 8.0. On Windows machines with the .NET 8
  Desktop Runtime installed, this surfaced as the Excel-DNA load error
  *"could not find runtime 8.0.0.0 ... file not found"* — the v0.5.0
  `.xll` was unloadable. Removing the attribute lets the host read the
  runtimeconfig.json and load .NET 8 correctly.

  Note: `RuntimeVersion` is for .NET Framework only. For SDK-style
  projects targeting .NET Core / .NET 5+, the framework is determined
  from the runtimeconfig — the attribute should be absent.

## [0.5.0] — 2026-05-03

### Fixed (numerical correctness)
- `ACT_EXPOSURE_LAYER_RATE` no longer divides by the layer's physical
  width. The exposure curve `G(d)` is already the cumulative *fraction*
  of expected loss; the layer's share is `G(exhaust) − G(attach)`
  directly. Old output was wrong by 1–2 orders of magnitude.
- `ACT_ILF_PARETO` rewritten as the standard Pareto-I LEV ratio
  `LEV(target) / LEV(base)`. The previous implementation had a
  hard-coded `1000×` magic number and a non-derivable additive term.
  The signature gains a fourth argument `xm` (Pareto scale, default
  1.0) — both `target` and `base` must be `>= xm`.
- `ACT_EXPOSURE_SWISSRE` now uses the published Bernegger (1997)
  parameterisation: `b = exp(3.1 − 0.15·c·(1+c))`,
  `g = exp(c·(0.78 + 0.12·c))` with curves 1–5 mapping to c = 1.5,
  2, 3, 4, 5. Curves 1–4 are Swiss Re Y1–Y4; curve 5 coincides with
  Lloyd's industrial-risks curve. The previous round-number `(b=g)`
  parameter pairs were not Bernegger curves at all.
- `ACT_AGGREGATE_TVAR` divides by the empirical tail mass
  (`Σ pmf[i] for i ≥ varIndex`) rather than the requested `1 − p`.
  Fixes Klugman 3.5.1 CTE under discretisation (the old denominator
  was biased by up to one PMF cell).

### Changed
- `ExcelDna.AddIn` bumped from 1.7.0 to 1.9.0 — adds native support for
  `int?` / nullable parameters and proper hosting of .NET 8 assemblies.
- Target framework retargeted to `net8.0-windows` (back from .NET 6
  EOL — possible now that ExcelDna 1.9 hosts net8 cleanly).
- All `seed = 0 ⇒ random` sentinels in `ACT_COPULA_*`,
  `ACT_CL_BOOTSTRAP*`, `ACT_CAT_ELT_TO_YLT` switched to nullable
  `int? seed = null`. Leave the cell blank for a system-time seed; pass
  any integer (including 0) for a reproducible stream. **Breaking
  change**: callers that previously passed `0` for "random" will now
  get the deterministic seed-zero stream.
- `ACT_CL_WEIGHTED_AVERAGE` now renormalises weights to sum to 1
  rather than erroring on small rounding mismatches.
- Copula error handling: dropped the locale-fragile
  `ex.Message.Contains("positive definite")` string-match in
  `ACT_COPULA_GAUSSIAN` / `_STUDENT_T`; non-positive-definite
  correlation matrices now surface as `#VALUE!` cleanly. The other
  copula samplers no longer leak exception messages into cells.
- Marginal-totals fallback hat-matrix in the bootstrap (only used when
  the GLM-QR fails on a degenerate triangle) clamps `h_ii` to
  `[0, 0.999]` to prevent residual shrinkage / blow-up.

### Added
- Bernegger (1997) and the published Swiss Re Y1–Y4 parameterisation
  documented in `ExposureCurves.cs` and the README references.
- README quickstart now lists the **.NET 8 Windows Desktop Runtime**
  install (via `winget install Microsoft.DotNet.DesktopRuntime.8`) as
  the one-time prerequisite, plus the `Unblock-File` step Windows
  requires for downloaded `.xll` files.

## [0.4.0] — 2026-05-02

### Changed
- ODP bootstrap (`ACT_CL_BOOTSTRAP`, `ACT_CL_BOOTSTRAP_ORIGIN`) now uses
  the **proper GLM hat matrix** (thin QR of √W·Z over the ODP design
  matrix) instead of the previous `m·(1/R + 1/C − 1/T)` marginal-totals
  approximation. Both formulations have trace = p, but per-cell `h_ii`
  values differ by O(few %). The GLM version reconciles per-origin to
  England (2010) slide 35 within Monte Carlo noise; the old version did
  not. Bootstrap and copulas remain in the `Actuarial.Experimental`
  category — the Experimental tag is an additional caution layer on top
  of the project-wide "no production-readiness claim" position, not the
  only risk indicator.
- Bootstrap reconciliation tightened: total PE within 5 %, per-origin
  PE within 10 %, mean within 2 % of deterministic IBNR, all pinned to
  England (2010) slide 35 reference values. Tests stay in the
  `experimental_bootstrap` section (do not fail CI on miss). Reference
  algorithm walkthrough (M0 → M5) and Shapland (2016) cell-by-cell
  reconciliation live in the companion repo
  [`mdevans21/bootstrapping_exposition`](https://github.com/mdevans21/bootstrapping_exposition).
- README disclaimers rewritten: clarified that AI-generated tests can
  contain mistakes, no function is offered as production-ready (with or
  without an Experimental tag), and validation against original sources
  is the user's responsibility. Closes user feedback that the prior
  language ("do not use Experimental in production") implicitly endorsed
  the rest for production.
- Target framework stays on `net6.0-windows` for this release.
  (Retargeted to `net8.0-windows` in v0.5.0 alongside an ExcelDna
  bump that supports it.)
- GitHub Actions workflows opt in to the Node.js 24 runtime via
  `FORCE_JAVASCRIPT_ACTIONS_TO_NODE24=true`. GitHub will default to
  Node 24 in June 2026; opting in now silences the deprecation warnings.
- `softprops/action-gh-release` is now pinned to a commit SHA
  (`3bb12739c2…`) instead of the floating `@v1` tag.
- Explicit `permissions:` blocks on every workflow job (`contents: read`
  for build/reconcile, `contents: write` for release only). Closes
  Should-fix #14.

### Added
- `tests/reconciliation.ipynb` — CI-gated Jupyter notebook that reconciles
  every Basic function against `scipy.stats`, `chainladder-python`, or an
  analytic reference and fails the build on tolerance miss.
- `src/ActuarialAddIn.Tests/AddinOutputsEmitter.cs` — structured JSON
  emitter (`tests/fixtures/addin_outputs.json`) consumed by the notebook.
- `tests/build_reconciliation_notebook.py` — generator that keeps the
  notebook reviewable as plain Python.
- `LICENSE` file (MIT).
- `docs/DEVELOPING.md` — developer-facing build / release / workbook
  regeneration workflow.
- CI job `reconcile` in `.github/workflows/build.yml` that runs
  papermill on Linux against the Windows-built JSON artefact.

### Changed
- Copulas (`ACT_COPULA_*`) and bootstrap (`ACT_CL_BOOTSTRAP`,
  `ACT_CL_BOOTSTRAP_ORIGIN`) are now tagged
  `Category = "Actuarial.Experimental"` in the Excel function browser
  and carry an explicit `[EXPERIMENTAL]` prefix in their tooltips.
- `ACT_TRIANGLE_LINK_RATIOS` now rejects triangles with fewer than two
  development periods instead of silently returning an empty array.
- README rewritten as a "shop window": badges, 60-second install,
  feature matrix, worked Taylor-Ashe IBNR example, Experimental callout,
  validation section, and 10 previously undocumented `_FIT` functions.

### Removed
- `archive/credibility/` and `tests/fixtures/credibility.json` — the
  credibility functions were removed in v0.3.0; the fixture and archived
  source are no longer relevant.
- `tests/eltr_worked_example.py`, `tests/elt_to_ylt_oep.py` —
  paste-the-data stubs that never materialised.
- `NEXT_STEPS.md` — content moved here.

## [0.3.1] — 2026-02-21

### Added
- Aggregate Claims tab in `excel/actuarial_add_in.xlsx` with Panjer
  worked examples (`ACT_DISCRETIZE_*`, `ACT_PANJER_*`, `ACT_AGGREGATE_*`).

### Changed
- Spreadsheet rebuilt from scratch as `.xlsx` (dropping the `.xlsm`
  macro container); full coverage of every Add-In function.
- Documented the two-step spreadsheet generation workflow
  (openpyxl → xlwings `Formula2` upgrade).

## [0.3.0] — 2026-02-21

### Added
- **Enhanced ODP Bootstrap** (`ACT_CL_BOOTSTRAP`,
  `ACT_CL_BOOTSTRAP_ORIGIN`) with two methods:
  - `"EV"` (default): full England & Verrall (2002) non-constant-scale
    ODP with hat-matrix adjusted residuals, corner exclusion,
    per-period φⱼ, scaled global pool, Bessel correction, pseudo-
    diagonal convention, Mack tail extrapolation. Reconciles to within
    ~0.5% of E&V Table 3 on Taylor-Ashe.
  - `"CHAINLADDER-PYTHON"`: constant φ, no hat-matrix, original
    diagonal — matches `chainladder-python` with `hat_adj=False`.
- `ACT_CL_CAPECOD_ULTIMATE`, `ACT_CL_CAPECOD_ELR` — Cape Cod method.
- Versions tab auto-generated by `scripts/populate_examples.py` from
  `src/ActuarialAddIn/VersionInfo.cs`.

### Fixed
- Chart rounded-corners (force `roundedCorners=False` on every openpyxl
  chart).
- `ACT_CAT_ELT_TO_YLT` `#NAME` error when invoked with default
  parameters from Excel 365.
- Bootstrap default method renamed from `"BASIC"` to
  `"CHAINLADDER-PYTHON"` for clarity.

### Removed
- Credibility functions (limited-fluctuation and Bühlmann) — they had
  low actuarial value compared with the cost of maintaining them; moved
  to `archive/credibility/` for reference.

## [0.2.0] — 2026-02-04

### Added
- **Panjer recursion** for aggregate claims (`ACT_PANJER_POISSON`,
  `_NEGBIN`, `_BINOMIAL`) with severity discretisation
  (`ACT_DISCRETIZE_*`) and risk-measure functions (`ACT_AGGREGATE_*`).
- **Zero-truncated** distributions: Poisson, NegBin, Binomial, Geometric
  (`ACT_DIST_ZT*`).
- **Zero-modified (zero-inflated)** distributions: Poisson (ZIP) and
  NegBin (ZINB) (`ACT_DIST_ZM*`).
- **Extended Pareto** family: Pareto III and Pareto IV
  (`ACT_DIST_PARETO3_*`, `ACT_DIST_PARETO4_*`).
- **Inverse Gaussian** distribution (`ACT_DIST_INVGAUSS_*`).
- **Loglogistic** distribution (`ACT_DIST_LOGLOGISTIC_*`).
- **Cat modelling** functions: `ACT_CAT_ELT_TO_YLT`,
  `ACT_CAT_YLT_OEP_CURVE`, `ACT_CAT_YLT_AEP_CURVE`,
  `ACT_CAT_OEP_CURVE_RP`, `ACT_CAT_AEP_CURVE_RP`,
  `ACT_VAR_FROM_SAMPLES`, `ACT_TVAR_FROM_SAMPLES`.
- First fresh-eyes code review against six design principles.

### Changed
- Cat functions renamed to `ACT_CAT_*` prefix for consistency.

## [0.1.0] — 2026-01-31

Initial tagged release.

### Added
- Limited Expected Value functions for 13 distributions (`ACT_DIST_*_LEV`).
- GitHub Actions workflow for auto-publishing releases.

### Fixed
- `ACT_DIST_GPD_FIT` probability-weighted-moments formula bug.
- `ACT_ILF_LAYER` array-parameter handling.
- `ActuarialAddIn.Tests` test runner for WSL/Windows path quirks.

### Removed
- `ACT_QS_CEDED`, `ACT_AGGREGATE_LAYER`, Berquist-Sherman method —
  replaced or out of scope.

---

For the full commit log, see
[`git log`](https://github.com/mdevans21/actuarial_add_in/commits/main)
or the **Versions** tab inside `excel/actuarial_add_in.xlsx`.
