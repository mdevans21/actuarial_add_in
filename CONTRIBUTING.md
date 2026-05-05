# Agent Instructions

Conventions for AI agents working on this project. Full developer
documentation lives in [`docs/DEVELOPING.md`](docs/DEVELOPING.md); this
file keeps only what an agent needs to know to avoid breaking the
repo's invariants.

## Golden rules

1. **Three test sources must stay consistent:**
   1. `src/ActuarialAddIn.Tests/Program.cs` — primary value tables and
      markdown report.
   2. `src/ActuarialAddIn.Tests/AddinOutputsEmitter.cs` — JSON snapshot
      consumed by the reconciliation notebook.
   3. `excel/actuarial_add_in.xlsx` — worked examples (regenerated via
      `scripts/populate_examples.py`).
   If you add a function, add it to all three.

2. **Ask before pushing a release.** Pushing to `main` is safe; pushing
   a version tag (`v0.X.Y`) triggers `.github/workflows/release.yml`
   and publishes public artefacts. Always confirm the user's intent
   before `git push origin main --tags`.

3. **Rebuild before committing C# changes.** The `.xll` artefact in
   the repo must stay in sync with source; stale DLLs break Excel
   examples. See `docs/DEVELOPING.md` for the build command (and for
   WSL/Windows filesystem sync mitigations when builds silently use
   stale code).

4. **Reconciliation before release.** Every new function needs a row
   in `AddinOutputsEmitter.cs` and a matching scipy/analytic reference
   in `tests/build_reconciliation_notebook.py`. The CI `reconcile`
   job must pass before tagging a release.

5. **Preserve the two-step spreadsheet workflow.** Do not attempt ZIP
   postprocessing of `.xlsx` files — it corrupts dynamic-array
   metadata. Use `scripts/populate_examples.py` then
   `scripts/fix_array_formulas.py` (xlwings, Windows-only).

6. **Experimental tagging is load-bearing.** Functions tagged
   `Category = "Actuarial.Experimental"` appear that way in Excel's
   function browser and are allowed to fail in the reconciliation
   notebook without breaking CI. Don't retag them to a stable category
   without a round of validation.

## Canonical test parameters

Every function family has a fixed parameter set shared across the three
test sources. **Do not drift these.**

| Family | Parameters |
|---|---|
| Poisson | λ = 5 |
| Negative Binomial | r = 5, p = 0.3 |
| Lognormal | μ = 0, σ = 1 |
| Gamma | α = 2, β = 1 |
| Pareto I | α = 2, xm = 1 |
| Inverse Gaussian | μ = 2, λ = 3 |
| Loglogistic | α = 2, β = 3 |
| Pareto III | μ = 1, σ = 2, γ = 3 |
| Pareto IV | μ = 0, σ = 2, γ = 0.5, α = 3 |
| MBBEFD | b = 2, g = 3 |
| Swiss Re curves | 1–5 |
| Lloyd's curves | Y1–Y4 |
| XOL layer | attach = 1M, limit = 5M |
| Student-t copula | df = 5, 7×7 corr ρ = 0.6 decay, seed = 42 |
| Panjer Poisson | λ = 2, max_s = 40 |
| Panjer NegBin | r = 2, p = 0.5, max_s = 40 |
| Panjer Binomial | n = 5, p = 0.3, max_s = 40 |
| Aggregate grid | h = 0.5 |
| Chain ladder triangle | Taylor-Ashe 10×10 cumulative paid |
| Expected LDFs | 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018 |
| Expected total IBNR | 18,680,856 |
| Mack total reserve SE | 2,447,095 (analytic), ≈ 2,454,616 (simulated) |

## ODP Bootstrap reference values

Target: England (2010) "Stochastic Claims Reserving Made Simple"
slide 35 (Taylor & Ashe 1983 `genins`, non-constant scale ODP). Per-origin
prediction errors and the algorithm walkthrough (M0 → M5) live in the
companion repo
[`mdevans21/bootstrapping_exposition`](https://github.com/mdevans21/bootstrapping_exposition).

| AY | IBNR | Constant SE | Non-Const SE |
|---|---:|---:|---:|
| 1 | 0 | 0 | 0 |
| 2 | 94 634 | 112 552 | 43 882 |
| 3 | 469 511 | 217 547 | 109 449 |
| 4 | 709 638 | 262 934 | 141 509 |
| 5 | 984 889 | 306 595 | 256 031 |
| 6 | 1 419 459 | 375 745 | 398 377 |
| 7 | 2 177 641 | 500 332 | 529 898 |
| 8 | 3 920 301 | 791 481 | 735 245 |
| 9 | 4 278 972 | 1 060 473 | 809 457 |
| 10 | 4 625 811 | 2 025 898 | 1 285 560 |
| **Total** | **18 680 856** | **2 992 296** | **2 228 677** |

### Bootstrap `method` parameter

`ACT_CL_BOOTSTRAP` and `ACT_CL_BOOTSTRAP_ORIGIN` take a `method`
argument:

- **`"EV"`** (default) — full England & Verrall methodology:
  hat-matrix adjusted residuals `r* = r / √(1 − h_ii)`, corner
  exclusion at `(0, n-1)` and `(n-1, 0)`, per-period heteroscedastic
  φⱼ, scaled global pool (standardise by √φⱼ, pool, un-standardise),
  Bessel variance correction (inflate by √(N/(N-1))), pseudo-diagonal
  convention (`IBNR = ultimate − pseudo_diagonal`), Mack tail
  extrapolation for the last φ.
- **`"CHAINLADDER-PYTHON"`** — stripped variant for parity with
  chainladder-python `hat_adj=False`: no hat matrix, no corner
  exclusion, single global φ, global pooled residuals, no Bessel
  correction, `IBNR = ultimate − original_diagonal`.

### Implementation notes (non-obvious invariants)

- **φ is computed from UNSCALED Pearson residuals**, not hat-adjusted
  — a common bug in ad-hoc implementations.
- **Hat matrix is the GLM hat,** computed via thin QR of `√W·Z` where Z
  is the `(2n-1)`-column ODP design matrix (n origin + n-1 dev params,
  β_0 = 0 constraint). Older versions used the marginal-totals
  approximation `m·(1/R + 1/C - 1/T)` — change preserves trace = p but
  individual h_ii values shift. Per-origin PE pattern matches England
  slide 35 only with the GLM hat.
- **Per-period df allocation:** `df_j = n_j · (n − p) / n` (proportional
  degrees of freedom, England 2010).
- **Mack tail extrapolation:**
  `φ_last = min(φ_{n-1}² / φ_{n-2}, min(φ_{n-1}, φ_{n-2}))`.
- **Process variance is INDEPENDENT per cell** — deterministic projection
  first, then independent Gamma noise per projected incremental. Not
  propagated through the chain ladder recursion.
- **Projection always from the pseudo-diagonal** — both EV and
  CHAINLADDER-PYTHON project from the pseudo-triangle's diagonal; only
  the subtracted diagonal differs between the two methods.
- **Backward projection for fitted values:**
  `fitted_cum[i,j] = diagonal[i] / Π_{k=j}^{lastCol-1} f_k`.

## Version info maintenance (during a release)

Update `src/ActuarialAddIn/VersionInfo.cs` before committing the
release:

1. `CurrentVersion` — the new tag (e.g. `"0.4.0"`).
2. `BuildDate` — today's ISO date.
3. `GetCommitHistory()` — replace with the latest 20 commits from:
   ```bash
   git log --format='new CommitInfo("%h", "%cs", "%s"),' -20
   ```

This file feeds `ACT_VERSION()`, `ACT_BUILD_DATE()`,
`ACT_COMMIT_HISTORY()`, and the **Versions** tab in
`excel/actuarial_add_in.xlsx` (regenerated by
`scripts/populate_examples.py`).

## See also

- [`README.md`](README.md) — user-facing shop window.
- Release history lives in
  [`VersionInfo.GetCommitHistory()`](src/ActuarialAddIn/VersionInfo.cs)
  and surfaces in the workbook's Versions tab via `=ACT_COMMIT_HISTORY()`.
- [`docs/DEVELOPING.md`](docs/DEVELOPING.md) — build, release, workbook
  regeneration, WSL troubleshooting, dynamic-array patterns.
- [GitHub Issues](https://github.com/mdevans21/actuarial_add_in/issues)
  — open bugs, must-fix items, and roadmap.
