# Changelog

All notable changes to the Actuarial Add-In are documented here. Dates are
ISO-8601. Versions follow [SemVer](https://semver.org/).

## [Unreleased]

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
