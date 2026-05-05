# Actuarial Add-In for Excel

[![Build](https://github.com/mdevans21/actuarial_add_in/actions/workflows/build.yml/badge.svg)](https://github.com/mdevans21/actuarial_add_in/actions/workflows/build.yml)
[![Release](https://img.shields.io/github/v/release/mdevans21/actuarial_add_in?display_name=tag&sort=semver)](https://github.com/mdevans21/actuarial_add_in/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform: Excel (Windows)](https://img.shields.io/badge/Excel-2016%2B%20%7C%20365-217346)](https://www.microsoft.com/microsoft-365/excel)
[![.NET 6.0](https://img.shields.io/badge/.NET-6.0-512BD4)](https://dotnet.microsoft.com/)

> A native Excel add-in that brings 174 general-insurance actuarial
> functions — **chain ladder**, **Mack standard errors**, **ODP bootstrap**,
> **MBBEFD / Swiss Re exposure curves**, **Panjer recursion**, **copulas**,
> **cat modelling (ELT → YLT → OEP)**, **17 loss distributions** with
> LEV — directly to any worksheet. Type `ACT_` in a cell and Excel's
> autocomplete takes over.

Every function is prefixed `ACT_`, has an Insert-Function description, and
returns spill-friendly arrays for Excel 365. Numerical results are
reconciled against `scipy`, `chainladder-python`, and published
actuarial literature in a CI-gated Jupyter notebook (see
[Validation](#validation--tests)).

---

## Install in 60 seconds

**One-time prerequisite** — install the **.NET 6 Windows Desktop
Runtime** if you don't already have it. Excel-DNA's host needs it to
load the add-in; without it Excel logs `Could not load file or assembly
'System.Runtime, Version=6.0.0.0'` on startup. Pick whichever:

```powershell
winget install Microsoft.DotNet.DesktopRuntime.6
```

or download from
[dotnet.microsoft.com/download/dotnet/6.0](https://dotnet.microsoft.com/download/dotnet/6.0)
under "Run desktop apps" → ".NET Desktop Runtime 6.0.x x64". Verify
with `dotnet --list-runtimes` — you should see a
`Microsoft.WindowsDesktop.App 6.0.x` line.

Then:

1. Grab the latest `.xll` from the
   [Releases page](https://github.com/mdevans21/actuarial_add_in/releases) —
   `ActuarialAddIn-AddIn64-packed.xll` for 64-bit Excel (most installs) or
   `ActuarialAddIn-AddIn-packed.xll` for 32-bit.
2. **Unblock the file** (Windows tags downloaded `.xll` files as
   untrusted): right-click the `.xll` → **Properties** → tick the
   **Unblock** checkbox at the bottom → **OK**. Or in PowerShell:
   `Unblock-File <path>`.
3. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go… →
   Browse** and pick the `.xll` you just unblocked.
4. When prompted, tick "Trust access" to the add-in (Excel remembers this
   for future sessions).
5. In any cell, type `=ACT_` — every add-in function shows up in
   autocomplete, grouped by category.

The `.xll` is a single signed assembly with no Python or R runtime
required and no telemetry.

Worked examples for every function family live in
[`excel/actuarial_add_in.xlsx`](excel/actuarial_add_in.xlsx).

---

## Worked example — Chain ladder IBNR in five lines

Paste the Taylor-Ashe triangle (10 × 10 cumulative paid losses from
Taylor & Ashe 1983) into `B2:K11`, then:

```excel
B13  =ACT_CL_FACTORS(B2:K11)          ' 9 LDFs  → 3.491, 1.747, 1.457, ...
B14  =ACT_CL_ULTIMATE(B2:K11)         ' projected ultimate by accident year
B15  =ACT_CL_IBNR(B2:K11)             ' reserve = ultimate - paid
B16  =SUM(B15#)                       ' Total IBNR → 18,680,856
B17  =ACT_MACK_RESERVE_SE(B2:K11)     ' Mack SE per AY; total SE ≈ 2,447,095
```

Need a stochastic reserve? `=ACT_CL_BOOTSTRAP(B2:K11, 10000, 123)` returns
the full percentile table (mean, stddev, P1…P99) from an England & Verrall
(2002) ODP bootstrap with 10 000 iterations and a fixed seed.

---

## What's inside

| Category | `[ExcelFunction]` count | Highlights | Primary reference |
|---|---|---|---|
| **Distributions** (continuous, discrete, ZT, ZM, composite, LEV) | 96 | Lognormal, Gamma, Pareto I/III/IV, GPD, Weibull, Burr, Beta, Inverse Gaussian, Loglogistic, Poisson, NegBin, plus zero-truncated & zero-modified variants and 12 LEV functions | Klugman, Panjer & Willmot, *Loss Models* |
| **Parameter fitting** (MLE) | 10 | `ACT_DIST_*_FIT` for Exp, Poisson, Lognormal, Gamma, Pareto, Weibull, GPD, Beta, NegBin, Burr | scipy.stats |
| **Chain ladder** | 9 | Volume-weighted LDFs, Mack standard errors, IBNR, Bornhuetter-Ferguson | Mack (1993); England & Verrall (2002) |
| **Aggregate claims** ⚠️ experimental | 12 | Panjer recursion (Poisson / NegBin / Binomial), severity discretisation, VaR, TVaR | Klugman ch. 6 |
| **Exposure curves / ILF** | 6 | MBBEFD, Swiss Re curves 1–5, Lloyd's Y1–Y4, Riebesell, power, Pareto ILF | Bernegger (1997) |
| **Cat modelling** ⚠️ experimental | 7 | ELT → YLT simulation, OEP/AEP empirical curves, VaR/TVaR from samples | McNeil / Frey / Embrechts |
| **Reinsurance layers** | 3 | XOL layer loss & expected loss, Pareto ILF (XOL pricing inputs only) | standard texts |
| **Return periods** ⚠️ experimental | 3 | RP-loss interpolation, RP table builder, AAL from OEP | standard texts |
| **Copulas** ⚠️ experimental | 16 | Gaussian, Student-t, Clayton, Frank, Gumbel — samplers, CDFs, τ↔θ, tail dependence | McNeil et al. (2015) |
| **ODP bootstrap** ⚠️ experimental | 2 | `ACT_CL_BOOTSTRAP` + `_ORIGIN`, full E&V 2002 with GLM hat matrix, per-period φⱼ, Bessel correction; reconciles per-origin to England (2010) slide 35 | England & Verrall (2002), England (2010) |
| **Interpolation & version info** | 9 | 1D / log / 2D bilinear, version badges inside the sheet | — |

**Total:** 174 worksheet functions. See the
[full function reference](#function-reference) below or use Excel's
**Insert Function** dialog and pick the `Actuarial.*` category.

---

## ⚠️ No production-readiness claim

Every function in this add-in is tested against an independent reference
(`scipy.stats`, `chainladder-python`, an analytic formula, or a published
worked example) in the CI-gated reconciliation notebook. **That is not the
same thing as being production-ready.**

- The reconciliation harness, fixture data, reference values, and the
  Python notebook itself are AI-generated. AI can — and does — make
  mistakes: a wrong reference value, a mismatched parameter convention, a
  miscoded formula in the Python check, an edge case never tested.
- "All tests pass" means the C# output agrees with the Python reference
  the harness was told to compare against. It does **not** mean the
  function is correct in the broader actuarial sense, robust on your
  data, or fit for any particular regulatory or pricing use.
- No warranty is offered on any function, even those not flagged below.
  Anyone using these functions for reserving, pricing, capital, solvency
  reporting, or any other professional purpose is responsible for their
  own validation against original sources.

### `Actuarial.Experimental` — additional caution

Five function families carry an additional warning. They are tagged
`Actuarial.Experimental` in Excel's function browser and prefixed
`[EXPERIMENTAL]` in their tooltips:

- **Copulas** (`ACT_COPULA_*`) — Gaussian, Student-t, Clayton, Frank,
  Gumbel samplers and bivariate CDFs. Analytical tail-dependence
  coefficients are stable; the Gumbel Marshall-Olkin sampler and the
  Kendall's τ → θ inversion for Frank are under active validation.
- **ODP bootstrap** (`ACT_CL_BOOTSTRAP`, `ACT_CL_BOOTSTRAP_ORIGIN`) —
  the full England & Verrall 2002 implementation reconciles per-origin
  to England (2010) slide 35 within Monte Carlo noise on Taylor-Ashe
  (see [Validation](#validation--tests)). Bootstrap correctness depends
  on a handful of subtle conventions (hat-matrix construction,
  per-period scale, pseudo-diagonal IBNR); pathological triangles may
  expose edge cases the tests don't cover.
- **Aggregate claims** (`ACT_AGGREGATE_*`, `ACT_DISCRETIZE_*`,
  `ACT_PANJER_*`) — Panjer recursion + severity discretisation. The
  individual building blocks reconcile to scipy CDF differences and
  closed-form moments, but real-world workflows compose several steps
  (discretisation grid choice, max-S truncation, frequency parameters)
  whose interaction can be sensitive on heavy-tailed severities.
- **Cat modelling** (`ACT_CAT_*`, `ACT_VAR_FROM_SAMPLES`,
  `ACT_TVAR_FROM_SAMPLES`) — ELT → YLT simulation and empirical
  EP-curve construction. Outputs are simulation-based: monotonic and
  bounded by sanity, but no published reference exists for direct
  per-cell comparison.
- **Return periods** (`ACT_RETURN_PERIOD_LOSS`,
  `ACT_RETURN_PERIOD_TABLE`, `ACT_AAL_FROM_OEP`) — log-linear
  interpolation across an EP curve plus trapezoidal AAL integration;
  thinly tested on the Taylor-Ashe demo curve only.

The Experimental tag exists so that we don't claim more than we know on
top of the base "no production-readiness" position above. It is an
*additional* caution layer, not the only risk indicator. Reconciliation
failures in these sections are reported but do **not** fail the CI
build.

---

## Validation & tests

Every Basic function is reconciled on each push against an independent
Python implementation. The canonical validation artefact is
[`tests/reconciliation.ipynb`](tests/reconciliation.ipynb) — a
papermill-executable notebook that:

1. Loads the C# harness's output from
   [`tests/fixtures/addin_outputs.json`](tests/fixtures/addin_outputs.json)
   (regenerated by `src/ActuarialAddIn.Tests` on every CI run).
2. Recomputes each value with `scipy.stats`, `chainladder-python`, or
   an analytical formula from Klugman / Mack / Bernegger.
3. Tallies pass/fail per family with explicit tolerances and raises
   `AssertionError` if any Basic reconciliation falls outside tolerance —
   which fails CI.

The ODP bootstrap is additionally pinned to England (2010) "Stochastic
Claims Reserving Made Simple" slide 35 per-origin prediction errors on
the Taylor & Ashe (1983) `genins` triangle. Reference values, methodology
walkthrough (M0 → M5: chainladder-python built-in → pseudo diagonal →
corner exclusion → hat-adjusted residuals → per-period φⱼ → Bessel
correction), and reconciliation against the Shapland (2016) companion
Excel and R `ChainLadder::BootChainLadder` are documented in the
companion repository
[`mdevans21/bootstrapping_exposition`](https://github.com/mdevans21/bootstrapping_exposition).
Both repos use the same algorithm for the EV mode; the C# implementation
matches England's published total CV (11.9 %) and per-origin PE pattern
to within Monte Carlo noise at 10 000 simulations.

Secondary artefacts:

- **[`test_results.md`](test_results.md)** — markdown-formatted
  value tables from the C# harness, uploaded as a build artefact. Useful
  for eyeballing current Add-In output without running scipy.
- Open issues, must-fix bugs, and roadmap items are tracked on the
  [GitHub Issues](https://github.com/mdevans21/actuarial_add_in/issues)
  page. The "no production-readiness" caveat above applies regardless
  of issue tags.

Run the notebook locally:

```bash
cd tests
pip install -r requirements.txt
python build_reconciliation_notebook.py        # regenerate from source
papermill reconciliation.ipynb executed.ipynb  # runs and fails on any Basic miss
```

---

## Function reference

### Distributions — PDF / CDF / Inverse

Shared test parameters (match the `excel/actuarial_add_in.xlsx` workbook and the C# test harness):

| Distribution | Parameters | Mean | Variance |
|---|---|---|---|
| Poisson | λ = 5 | 5 | 5 |
| Negative Binomial | r = 5, p = 0.3 | 11.67 | 38.89 |
| Lognormal | μ = 0, σ = 1 | 1.649 | 4.671 |
| Gamma | α = 2, β = 1 | 2 | 2 |
| Pareto I | α = 2, x_m = 1 | 2 | ∞ |

| Distribution | PDF | CDF | Inverse |
|---|---|---|---|
| Normal | `ACT_DIST_NORMAL_PDF(x, μ, σ)` | `ACT_DIST_NORMAL_CDF(x, μ, σ)` | `ACT_DIST_NORMAL_INV(p, μ, σ)` |
| Poisson | `ACT_DIST_POISSON_PDF(k, λ)` | `ACT_DIST_POISSON_CDF(k, λ)` | `ACT_DIST_POISSON_INV(p, λ)` |
| Negative Binomial | `ACT_DIST_NEGBIN_PDF(k, r, p)` | `ACT_DIST_NEGBIN_CDF(k, r, p)` | `ACT_DIST_NEGBIN_INV(prob, r, p)` |
| Lognormal | `ACT_DIST_LOGNORM_PDF(x, μ, σ)` | `ACT_DIST_LOGNORM_CDF(x, μ, σ)` | `ACT_DIST_LOGNORM_INV(p, μ, σ)` |
| Gamma | `ACT_DIST_GAMMA_PDF(x, α, β)` | `ACT_DIST_GAMMA_CDF(x, α, β)` | `ACT_DIST_GAMMA_INV(p, α, β)` |
| Pareto I | `ACT_DIST_PARETO_PDF(x, α, xm)` | `ACT_DIST_PARETO_CDF(x, α, xm)` | `ACT_DIST_PARETO_INV(p, α, xm)` |
| Lomax (Pareto II) | `ACT_DIST_LOMAX_PDF(x, α, λ)` | `ACT_DIST_LOMAX_CDF(x, α, λ)` | `ACT_DIST_LOMAX_INV(p, α, λ)` |
| GPD | `ACT_DIST_GPD_PDF(x, ξ, σ)` | `ACT_DIST_GPD_CDF(x, ξ, σ)` | `ACT_DIST_GPD_INV(p, ξ, σ)` |
| Weibull | `ACT_DIST_WEIBULL_PDF(x, k, λ)` | `ACT_DIST_WEIBULL_CDF(x, k, λ)` | `ACT_DIST_WEIBULL_INV(p, k, λ)` |
| Beta | `ACT_DIST_BETA_PDF(x, α, β)` | `ACT_DIST_BETA_CDF(x, α, β)` | `ACT_DIST_BETA_INV(p, α, β)` |
| Exponential | `ACT_DIST_EXP_PDF(x, λ)` | `ACT_DIST_EXP_CDF(x, λ)` | `ACT_DIST_EXP_INV(p, λ)` |
| Burr XII | `ACT_DIST_BURR_PDF(x, c, k, λ)` | `ACT_DIST_BURR_CDF(x, c, k, λ)` | `ACT_DIST_BURR_INV(p, c, k, λ)` |
| Inverse Gaussian | `ACT_DIST_INVGAUSS_PDF(x, μ, λ)` | `ACT_DIST_INVGAUSS_CDF(x, μ, λ)` | `ACT_DIST_INVGAUSS_INV(p, μ, λ)` |
| Loglogistic | `ACT_DIST_LOGLOGISTIC_PDF(x, α, β)` | `ACT_DIST_LOGLOGISTIC_CDF(x, α, β)` | `ACT_DIST_LOGLOGISTIC_INV(p, α, β)` |
| Pareto III | `ACT_DIST_PARETO3_PDF(x, μ, σ, γ)` | `ACT_DIST_PARETO3_CDF(x, μ, σ, γ)` | `ACT_DIST_PARETO3_INV(p, μ, σ, γ)` |
| Pareto IV | `ACT_DIST_PARETO4_PDF(x, μ, σ, γ, α)` | `ACT_DIST_PARETO4_CDF(x, μ, σ, γ, α)` | `ACT_DIST_PARETO4_INV(p, μ, σ, γ, α)` |

Also: `ACT_DIST_PARETO3_MEAN(μ, σ, γ)` (requires γ > 1).

**Composite distributions** — smooth Scollnik (2007) composites for heavy-tail modelling:

| Distribution | PDF / CDF | α derivation |
|---|---|---|
| Lognormal-Pareto | `ACT_DIST_LNPARETO_PDF/CDF(x, μ, σ, θ)` | `ACT_DIST_LNPARETO_ALPHA(μ, σ, θ)` |
| Exponential-Pareto | `ACT_DIST_EXPPARETO_PDF/CDF(x, λ, θ)` | `ACT_DIST_EXPPARETO_ALPHA(λ, θ)` |
| Power-Pareto | `ACT_DIST_POWPARETO_PDF/CDF(x, α, β, θ)` | — |

### Zero-truncated (X \| X > 0)

| Distribution | PMF | CDF | Inverse | Mean |
|---|---|---|---|---|
| ZT Poisson | `ACT_DIST_ZTPOISSON_PDF(k, λ)` | `ACT_DIST_ZTPOISSON_CDF(k, λ)` | `ACT_DIST_ZTPOISSON_INV(p, λ)` | `ACT_DIST_ZTPOISSON_MEAN(λ)` |
| ZT NegBin | `ACT_DIST_ZTNEGBIN_PDF(k, r, p)` | `ACT_DIST_ZTNEGBIN_CDF(k, r, p)` | `ACT_DIST_ZTNEGBIN_INV(prob, r, p)` | `ACT_DIST_ZTNEGBIN_MEAN(r, p)` |
| ZT Binomial | `ACT_DIST_ZTBINOM_PDF(k, n, p)` | `ACT_DIST_ZTBINOM_CDF(k, n, p)` | `ACT_DIST_ZTBINOM_INV(prob, n, p)` | `ACT_DIST_ZTBINOM_MEAN(n, p)` |
| ZT Geometric | `ACT_DIST_ZTGEOM_PDF(k, p)` | `ACT_DIST_ZTGEOM_CDF(k, p)` | `ACT_DIST_ZTGEOM_INV(prob, p)` | `ACT_DIST_ZTGEOM_MEAN(p)` |

### Zero-modified (zero-inflated)

`P(X=0) = p₀ + (1−p₀)·g(0)`, `P(X=k) = (1−p₀)·g(k)` for k ≥ 1.

| Distribution | PMF | CDF | Inverse | Mean / Variance |
|---|---|---|---|---|
| ZM Poisson | `ACT_DIST_ZMPOISSON_PDF(k, λ, p0)` | `ACT_DIST_ZMPOISSON_CDF(k, λ, p0)` | `ACT_DIST_ZMPOISSON_INV(p, λ, p0)` | `ACT_DIST_ZMPOISSON_MEAN(λ, p0)`, `_VAR(λ, p0)` |
| ZM NegBin | `ACT_DIST_ZMNEGBIN_PDF(k, r, p, p0)` | `ACT_DIST_ZMNEGBIN_CDF(k, r, p, p0)` | `ACT_DIST_ZMNEGBIN_INV(prob, r, p, p0)` | `ACT_DIST_ZMNEGBIN_MEAN(r, p, p0)`, `_VAR(r, p, p0)` |

### Limited Expected Value — E[min(X, d)]

| Distribution | LEV | Notes |
|---|---|---|
| Exponential | `ACT_DIST_EXP_LEV(d, λ)` | closed form |
| Pareto I | `ACT_DIST_PARETO_LEV(d, α, xm)` | α > 1 |
| Lomax (Pareto II) | `ACT_DIST_LOMAX_LEV(d, α, λ)` | α > 1 |
| GPD | `ACT_DIST_GPD_LEV(d, ξ, σ)` | ξ < 1 |
| Gamma | `ACT_DIST_GAMMA_LEV(d, α, β)` | incomplete gamma |
| Lognormal | `ACT_DIST_LOGNORM_LEV(d, μ, σ)` | uses Φ |
| Weibull | `ACT_DIST_WEIBULL_LEV(d, k, λ)` | incomplete gamma |
| Beta | `ACT_DIST_BETA_LEV(d, α, β)` | on [0,1], incomplete beta |
| Burr XII | `ACT_DIST_BURR_LEV(d, c, k, λ)` | ck > 1 |
| Inverse Gaussian | `ACT_DIST_INVGAUSS_LEV(d, μ, λ)` | numerical integration |
| Loglogistic | `ACT_DIST_LOGLOGISTIC_LEV(d, α, β)` | numerical integration |
| Poisson | `ACT_DIST_POISSON_LEV(d, λ)` | discrete summation |
| Negative Binomial | `ACT_DIST_NEGBIN_LEV(d, r, p)` | discrete summation |

### Parameter fitting (MLE)

Pass a single-column range of observations; the function returns a
scalar (one-parameter fits) or a horizontal array of parameters.

| Function | Returns | Method |
|---|---|---|
| `ACT_DIST_EXP_FIT(data)` | `λ` | closed-form: `1/mean(data)` |
| `ACT_DIST_POISSON_FIT(data)` | `λ` | closed-form: `mean(data)` |
| `ACT_DIST_LOGNORM_FIT(data)` | `{μ, σ}` | closed-form on `log(data)` |
| `ACT_DIST_GAMMA_FIT(data)` | `{α, β}` | MLE (Newton iteration) |
| `ACT_DIST_PARETO_FIT(data)` | `{α, xm}` | MLE (x_m = min, closed-form α) |
| `ACT_DIST_WEIBULL_FIT(data)` | `{k, λ}` | MLE (Newton iteration) |
| `ACT_DIST_GPD_FIT(data)` | `{ξ, σ}` | MLE |
| `ACT_DIST_BETA_FIT(data)` | `{α, β}` | MLE on [0,1] |
| `ACT_DIST_NEGBIN_FIT(data)` | `{r, p}` | method-of-moments |
| `ACT_DIST_BURR_FIT(data)` | `{c, k, λ}` | MLE (numeric optimisation) |

All fits return `#VALUE!` on degenerate or invalid input (the C# code
emits `ExcelError.ExcelErrorValue` so Excel renders a real cell error
rather than a text string).

### Chain ladder & reserving

Reference data: **Taylor-Ashe** 10 × 10 cumulative paid-loss triangle
(Taylor & Ashe 1983). The same triangle drives the C# test harness and
the reconciliation notebook.

| Function | Returns | Reference |
|---|---|---|
| `ACT_CL_FACTORS(tri, [topN], [exclRecent], [exclHiLo], [vertical])` | volume-weighted LDFs | Mack (1993) |
| `ACT_CL_ULTIMATE(tri, [vertical])` | projected ultimates | |
| `ACT_CL_IBNR(tri, [vertical])` | reserves = ultimate − paid | |
| `ACT_CL_LATEST(tri, [vertical])` | latest diagonal | |
| `ACT_BF_ULTIMATE(tri, factors, apriori)` | Bornhuetter-Ferguson | BF (1972) |
| `ACT_MACK_FACTOR_SE(tri, [vertical])` | standard error of LDFs | Mack (1993) |
| `ACT_MACK_RESERVE_SE(tri, [vertical])` | reserve SE by AY | Mack (1993, 1999) |
| `ACT_CL_BOOTSTRAP(tri, iter, [seed], [method])` ⚠️ | total reserve distribution | England & Verrall (2002) |
| `ACT_CL_BOOTSTRAP_ORIGIN(tri, iter, [seed], [method])` ⚠️ | reserve distribution by AY | |

**Bootstrap modes:** pass `method = "EV"` (default) for full
England & Verrall 2002 non-constant-scale ODP with hat-matrix adjustment
and per-period φⱼ, or `method = "CHAINLADDER-PYTHON"` for a stripped
variant that matches `chainladder-python`'s `hat_adj=False` output.
See [`CONTRIBUTING.md`](CONTRIBUTING.md) for the full reconciliation matrix.

### Aggregate claims (Panjer recursion)

Compute `S = X₁ + X₂ + … + X_N` where N ~ frequency, Xᵢ ~ severity.

```excel
' 1. Discretise severity (Exponential rate = 1, grid h = 0.5, 40 points)
B2: =ACT_DISCRETIZE_EXPONENTIAL(1, 0.5, 40)

' 2. Compound with Poisson(λ = 2) frequency up to S = 100
C2: =ACT_PANJER_POISSON(2, B2#, 100)

' 3. Derive risk measures
D1: =ACT_AGGREGATE_MEAN(C2#, 0.5)
D2: =ACT_AGGREGATE_VAR(0.95, C2#, 0.5)
D3: =ACT_AGGREGATE_TVAR(0.95, C2#, 0.5)
```

| Purpose | Function |
|---|---|
| Discretise severity | `ACT_DISCRETIZE_EXPONENTIAL`, `_GAMMA`, `_LOGNORMAL` |
| Panjer recursion | `ACT_PANJER_POISSON`, `_NEGBIN`, `_BINOMIAL` |
| Aggregate statistics | `ACT_AGGREGATE_MEAN`, `_STDEV`, `_VAR_STAT`, `_CDF`, `_VAR`, `_TVAR` |

Reference: Klugman, Panjer & Willmot, *Loss Models* ch. 6 (Panjer
recursion for the (a, b, 0) class).

### Exposure curves, ILF & return periods

| Function | Description |
|---|---|
| `ACT_EXPOSURE_MBBEFD(d, b, g)` | MBBEFD G(d) — Bernegger (1997) |
| `ACT_EXPOSURE_SWISSRE(d, curve)` | Swiss Re curves 1–5 |
| `ACT_EXPOSURE_LLOYDS(d, curve)` | Lloyd's Y1–Y4 |
| `ACT_EXPOSURE_POWER(d, n)` | `G(d) = d^n` |
| `ACT_EXPOSURE_INVERSE_POWER(d, n)` | `G(d) = 1 − (1−d)^n` |
| `ACT_EXPOSURE_PARETO(d, α)` | Pareto exposure curve |
| `ACT_EXPOSURE_RIEBESELL(d, n)` | Riebesell curve (Riegel 2008) |
| `ACT_EXPOSURE_RIEBESELL_INV(g, n, [tol], [maxIter])` | numerical inverse |
| `ACT_EXPOSURE_LAYER_RATE(attach%, exhaust%, burningCost, b, g)` | layer rate on line |
| `ACT_ILF_PARETO(target, base, α)` | Pareto increased-limit factor |
| `ACT_RETURN_PERIOD_LOSS(rps, losses, targetRP, [method])` | log-/linear interpolation |
| `ACT_RETURN_PERIOD_TABLE(rps, losses, targetRPs, [method])` | bulk interpolation |
| `ACT_AAL_FROM_OEP(rps, oepLosses)` | AAL via trapezoidal integration |

### Cat modelling

| Function | Description |
|---|---|
| `ACT_CAT_ELT_TO_YLT(rates, losses, years, [seed], [header])` | simulate YLT (aggregate, max, count) |
| `ACT_CAT_YLT_OEP_CURVE(maxLosses, [method], [header])` | OEP via plotting positions |
| `ACT_CAT_YLT_AEP_CURVE(aggLosses, [method], [header])` | AEP via plotting positions |
| `ACT_CAT_OEP_CURVE_RP(maxLosses, returnPeriods, [header])` | OEP at specific RPs |
| `ACT_CAT_AEP_CURVE_RP(aggLosses, returnPeriods, [header])` | AEP at specific RPs |
| `ACT_VAR_FROM_SAMPLES(samples, α)` | empirical VaR |
| `ACT_TVAR_FROM_SAMPLES(samples, α)` | empirical TVaR |

### Reinsurance layers

| Function | Description |
|---|---|
| `ACT_XOL_LAYER_LOSS(loss, attach, limit)` | `min(max(0, loss − attach), limit)` |
| `ACT_XOL_EXPECTED_LOSS(freq, attach, limit, α, xm)` | Pareto LEV-based layer mean |

### Copulas ⚠️ experimental

| Function | Description |
|---|---|
| `ACT_COPULA_GAUSSIAN(Σ, n, [seed])` | correlated uniforms via Cholesky |
| `ACT_COPULA_STUDENT_T(Σ, df, n, [seed])` | t-copula with tail dependence |
| `ACT_COPULA_CLAYTON(θ, n, [seed])` | lower-tail dependence |
| `ACT_COPULA_FRANK(θ, n, [seed])` | symmetric, no tail dependence |
| `ACT_COPULA_GUMBEL(θ, n, [seed])` | upper-tail dependence (Marshall-Olkin) |
| `ACT_COPULA_*_SINGLE(...)` | single-sample variants |
| `ACT_COPULA_CLAYTON_CDF`, `_FRANK_CDF`, `_GUMBEL_CDF` | analytic CDFs |
| `ACT_COPULA_TAU_TO_THETA(τ, type)` | Kendall's τ → θ |
| `ACT_COPULA_TAIL_LOWER / _UPPER(type, θ, [df])` | tail dependence coefficients |

### Interpolation

| Function | Description |
|---|---|
| `ACT_INTERP(xs, ys, x, [extrap])` | linear with FLAT/GRADIENT/ERROR extrapolation |
| `ACT_INTERP_LOG(xs, ys, x, [extrap])` | log-linear (common for yield curves, return periods) |
| `ACT_INTERP2D(xs, ys, zs, x, y)` | bilinear 2D |

### Version / build info

`ACT_VERSION()`, `ACT_BUILD_DATE()`, `ACT_GITHUB_URL()`,
`ACT_COMMIT_INFO(index, [field])`, `ACT_COMMIT_COUNT()`,
`ACT_COMMIT_HISTORY()` — useful for pinning a workbook's calculations
to a specific add-in release.

---

## Building from source

Building requires the .NET 6 SDK and targets `net6.0-windows` — building
on Linux requires `-p:EnableWindowsTargeting=true`, and the resulting
add-in only runs on Windows Excel.

```bash
dotnet restore ActuarialAddIn.sln
dotnet build ActuarialAddIn.sln --configuration Release

# 64-bit XLL lands here:
# src/ActuarialAddIn/bin/Release/net6.0-windows/publish/ActuarialAddIn-AddIn64-packed.xll
```

The C# test harness and JSON-emitter are in
`src/ActuarialAddIn.Tests`. Running `dotnet run --project
src/ActuarialAddIn.Tests -- test_results.md` produces both
`test_results.md` (human-readable) and
`tests/fixtures/addin_outputs.json` (machine-readable; consumed by the
reconciliation notebook).

See [`docs/DEVELOPING.md`](docs/DEVELOPING.md) for the full build /
release / spreadsheet-regeneration workflow, including the two-step
`.xlsx` generation process and WSL / Windows filesystem notes.

---

## Contributing

- Bug reports and feature requests: open an issue on GitHub with a
  minimal reproducer or the reference you want implemented.
- Pull requests: please add a reconciliation entry in
  `src/ActuarialAddIn.Tests/AddinOutputsEmitter.cs` and a matching
  reference value in `tests/build_reconciliation_notebook.py`. The CI
  `reconcile` job must pass.
- Numerical issues: if a function's output differs from a reference,
  the preferred channel is a PR that adds the test case — the bisection
  of C# vs. reference is usually more productive than a bug report.

---

## References

### Academic

- Bornhuetter, R. L. and Ferguson, R. E. (1972) "The actuary and IBNR." *Proc. CAS* LIX: 181–195.
- Taylor, G. and Ashe, F. R. (1983) "Second moments of estimates of outstanding claims." *J. Econometrics* 23: 37–61.
- Mack, T. (1993) "Distribution-free calculation of the standard error of chain ladder reserve estimates." *ASTIN Bulletin* 23(2): 213–225.
- Mack, T. (1999) "The standard error of chain ladder reserve estimates: recursive calculation and inclusion of a tail factor." *ASTIN Bulletin* 29(2): 361–366.
- Bernegger, S. (1997) "The Swiss Re exposure curves and the MBBEFD distribution class." *ASTIN Bulletin* 27(1): 99–111.
- England, P. D. and Verrall, R. J. (2002) "Stochastic claims reserving in general insurance." *British Actuarial Journal* 8(3): 443–518.
- England, P. D. (2010) "Addendum to analytic and bootstrap estimates of prediction errors in claims reserving." *Actuarial Science*.
- Klugman, S. A., Panjer, H. H. and Willmot, G. E. (2019) *Loss Models: From Data to Decisions*, 5th ed., Wiley.
- McNeil, A. J., Frey, R. and Embrechts, P. (2015) *Quantitative Risk Management*, Princeton University Press.
- Riegel, U. (2008) "Generalizations of common ILF models." *Blätter der DGVFM* 29: 45–71.
- Scollnik, D. P. M. (2007) "On composite lognormal-Pareto models." *Scandinavian Actuarial Journal* 1: 20–33.

### Software

- [Excel-DNA](https://excel-dna.net/) — Excel add-in framework
- [MathNet.Numerics](https://numerics.mathdotnet.com/) — numerical library
- [scipy.stats](https://docs.scipy.org/doc/scipy/reference/stats.html) — reconciliation reference
- [chainladder-python](https://chainladder-python.readthedocs.io/) — reserving reference

---

## Changelog

Release history lives in
[`VersionInfo.GetCommitHistory()`](src/ActuarialAddIn/VersionInfo.cs)
and surfaces in the workbook's *Versions* tab as a live spill from
`=ACT_COMMIT_HISTORY()`. Per-tag release notes are on the
[GitHub Releases page](https://github.com/mdevans21/actuarial_add_in/releases).

## License

[MIT](LICENSE)
