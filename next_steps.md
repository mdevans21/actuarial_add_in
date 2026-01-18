# Actuarial Add-In: Review Findings and Next Steps

## 1. Function Naming Consistency

### Current Issues

**Copula Functions** - Current naming puts the copula type first:
- `ACT_STUDENT_T_COPULA` → Should be `ACT_COPULA_STUDENT_T`
- `ACT_STUDENT_T_COPULA_SINGLE` → Should be `ACT_COPULA_STUDENT_T_SINGLE`

This matters because Excel's function autocomplete sorts alphabetically. When more copula types are added (Gaussian, Clayton, Frank, Gumbel), having `ACT_COPULA_` as the prefix groups them together.

**Recommended Copula Naming Convention:**
```
ACT_COPULA_GAUSSIAN       (to be added)
ACT_COPULA_STUDENT_T      (rename from ACT_STUDENT_T_COPULA)
ACT_COPULA_CLAYTON        (to be added)
ACT_COPULA_FRANK          (to be added)
ACT_COPULA_GUMBEL         (to be added)
```

### Other Naming Observations

The rest of the naming is consistent and well-structured:
- Distributions: `ACT_{DIST}_PDF/CDF/INV` ✓
- Chain Ladder: `ACT_CL_*`, `ACT_BF_*`, `ACT_MACK_*`, `ACT_BOOTSTRAP_CL_*` ✓
- Reinsurance: `ACT_XOL_*`, `ACT_QS_*`, `ACT_ILF_*` ✓
- Exposure Curves: `ACT_MBBEFD`, `ACT_SWISSRE_CURVE`, `ACT_LLOYDS_CURVE`, `ACT_*_CURVE` ✓
- Interpolation: `ACT_INTERP*` ✓

---

## 2. Formula Help Text Completeness

### Missing/Incomplete Help Text

**Distributions.cs:**
- Missing: Mathematical formula in description
- Missing: Parameter constraints reminder (e.g., "Returns NaN if parameters invalid")
- Example: ACT_PARETO_PDF should mention "Uses Type I Pareto with PDF: α·xm^α / x^(α+1)"

**ExposureCurves.cs:**
- ACT_MBBEFD: Good description but could mention this is also known as "Maxwell-Boltzmann-Bose-Einstein-Fermi-Dirac" curve
- ACT_SWISSRE_CURVE: Should list what curves 1-5 represent (light/medium/heavy exposure)
- ACT_LLOYDS_CURVE: Missing description of what Y1-Y4 represent

**ChainLadder.cs:**
- ACT_BF_ULTIMATE: Could explain when to use B-F vs standard chain ladder
- ACT_MACK_* functions: Should reference Mack (1993, 1999) papers
- ACT_BERQUIST_SHERMAN: Should reference the original 1978 paper

**Reinsurance.cs:**
- ACT_XOL_EXPECTED_LOSS: Missing explanation of the Pareto severity assumption
- ACT_ILF_PARETO: Formula shown in code comments doesn't match standard ILF formula

**Copulas.cs:**
- Missing: Explanation of what copula degrees of freedom means for tail dependence
- Missing: Warning about correlation matrix needing to be positive definite

---

## 3. References to Add

### Academic References

**Chain Ladder Methods:**
- Mack, T. (1993). "Distribution-free calculation of the standard error of chain ladder reserve estimates." ASTIN Bulletin 23(2): 213-225.
- Mack, T. (1999). "The standard error of chain ladder reserve estimates: Recursive calculation and inclusion of a tail factor." ASTIN Bulletin 29(2): 361-366.
- England, P.D. & Verrall, R.J. (2002). "Stochastic claims reserving in general insurance." British Actuarial Journal 8(3): 443-518.

**Bornhuetter-Ferguson:**
- Bornhuetter, R.L. & Ferguson, R.E. (1972). "The actuary and IBNR." Proceedings of the CAS, LIX: 181-195.

**Berquist-Sherman:**
- Berquist, J.R. & Sherman, R.E. (1977). "Loss reserve adequacy testing: A comprehensive, systematic approach." Proceedings of the CAS, LXIV: 123-184.

**Exposure Curves:**
- Bernegger, S. (1997). "The Swiss Re Exposure Curves and the MBBEFD Distribution Class." ASTIN Bulletin 27(1): 99-111.
- Riegel, U. (2008). "Generalizations of common ILF models." Blätter der DGVFM 29: 45-71.

**Copulas:**
- McNeil, A.J., Frey, R., & Embrechts, P. (2015). "Quantitative Risk Management: Concepts, Techniques and Tools." Princeton University Press.
- Joe, H. (2014). "Dependence Modeling with Copulas." CRC Press.

**Power Curves:**
- Salzmann, R. & Wüthrich, M.V. (2012). "Modeling accounting year dependence in runoff triangles." European Actuarial Journal 2(2): 227-242.

---

## 4. Missing Popular Actuarial Functions (General Insurance)

### High Priority - Distributions

1. **Generalized Pareto Distribution (GPD)** - Essential for extreme value theory
   - `ACT_GPD_PDF`, `ACT_GPD_CDF`, `ACT_GPD_INV`

2. **Weibull Distribution** - Common for claim severity
   - `ACT_WEIBULL_PDF`, `ACT_WEIBULL_CDF`, `ACT_WEIBULL_INV`

3. **Burr Distribution (Type XII)** - Heavy-tailed severity modeling
   - `ACT_BURR_PDF`, `ACT_BURR_CDF`, `ACT_BURR_INV`

4. **Beta Distribution** - Loss ratios, proportions
   - `ACT_BETA_PDF`, `ACT_BETA_CDF`, `ACT_BETA_INV`

5. **Exponential Distribution** - Waiting times, simple severity
   - `ACT_EXP_PDF`, `ACT_EXP_CDF`, `ACT_EXP_INV`

### High Priority - Parameter Estimation

6. **Maximum Likelihood Estimators (MLE):**
   - `ACT_PARETO_FIT` - Fit Pareto alpha from data
   - `ACT_LOGNORM_FIT` - Fit lognormal mu, sigma
   - `ACT_GAMMA_FIT` - Fit gamma shape, rate
   - `ACT_GPD_FIT` - Fit GPD parameters
   - `ACT_WEIBULL_FIT` - Fit Weibull parameters

### High Priority - Copulas

7. **Additional Copula Types:**
   - `ACT_COPULA_GAUSSIAN` - Gaussian/Normal copula
   - `ACT_COPULA_CLAYTON` - Lower tail dependence (common for claims)
   - `ACT_COPULA_FRANK` - Symmetric, no tail dependence
   - `ACT_COPULA_GUMBEL` - Upper tail dependence

### Medium Priority - Reserving

8. **Cape Cod Method:**
   - `ACT_CAPECOD_ULTIMATE` - Alternative to B-F
   - `ACT_CAPECOD_ELR` - Estimate expected loss ratio

9. **Frequency-Severity Separation:**
   - `ACT_PAID_TO_INCURRED_RATIO` - P/I ratio analysis
   - `ACT_DISPOSAL_RATE` - Claims closure analysis

10. **Calendar Year Effects:**
    - `ACT_CL_CALENDAR_ADJUSTMENT` - Adjust for calendar year inflation

### Medium Priority - Exposure Curves

11. **Riebesell Curve** (from next_steps.md):
    - `ACT_RIEBESELL_CURVE` - Alternative exposure curve

12. **Customizable Curves:**
    - `ACT_EXPOSURE_RATE_CHANGE` - Rate change impact using exposure curves

### Medium Priority - Pricing

13. **Experience Rating:**
    - `ACT_CREDIBILITY_BUHLMANN` - Bühlmann credibility factor
    - `ACT_CREDIBILITY_BUHLMANN_STRAUB` - Bühlmann-Straub model
    - `ACT_EXPERIENCE_MOD` - Experience modification factor

14. **Large Loss Loading:**
    - `ACT_EXCESS_RATIO` - Excess loss ratio
    - `ACT_ILFS_TABLE` - Generate full ILF table

### Lower Priority - Utilities

15. **Triangle Utilities:**
    - `ACT_TRIANGLE_TO_INCREMENTAL` - Convert cumulative to incremental
    - `ACT_INCREMENTAL_TO_CUMULATIVE` - Convert incremental to cumulative
    - `ACT_CL_WEIGHTED_AVERAGE` - Weighted average of methods

16. **Statistical Tests:**
    - `ACT_KOLMOGOROV_SMIRNOV` - K-S test for distribution fit
    - `ACT_ANDERSON_DARLING` - A-D test for distribution fit

---

## 5. Implementation Plan

### Phase 1: Quick Wins (Breaking Changes + Core Additions)
1. **Rename copula functions** for consistency
2. **Add GPD distribution** (per original next_steps.md)
3. **Add Weibull distribution**
4. **Add Beta distribution**
5. **Enhance help text** with formulas and constraints

### Phase 2: Parameter Estimation
1. Add MLE fitting functions for existing distributions
2. Add MLE fitting for new distributions
3. Implement method-of-moments as alternative

### Phase 3: Copulas Expansion
1. Add Gaussian copula
2. Add Clayton copula
3. Add Frank copula
4. Add Gumbel copula
5. Add copula selection/comparison utilities

### Phase 4: Exposure Curves
1. Add Riebesell curve (per original next_steps.md)
2. Add exposure curve inversion utilities

### Phase 5: Reserving Enhancements
1. Add Cape Cod method
2. Add triangle conversion utilities
3. Add calendar year adjustment functions

### Phase 6: Pricing & Credibility
1. Add Bühlmann credibility
2. Add experience modification
3. Add ILF table generation

---

## 6. Code Quality Items

### Add XML Documentation Comments
Consider adding `<summary>`, `<param>`, `<returns>`, and `<example>` XML comments for IntelliSense support.

### Add References Section
Create a REFERENCES.md file or add a references section to code comments citing the academic sources.

### Unit Test Coverage
Ensure test coverage for:
- Edge cases (zero, negative, boundary values)
- Known analytical results (compare to published examples)
- Numerical stability (very large/small values)

---

## Summary of Immediate Actions

| Priority | Action | Files Affected |
|----------|--------|----------------|
| 1 | Rename `ACT_STUDENT_T_COPULA*` → `ACT_COPULA_STUDENT_T*` | Copulas.cs |
| 2 | Add GPD distribution | Distributions.cs |
| 3 | Add parameter estimation functions | New: Fitting.cs |
| 4 | Enhance help text with formulas | All function files |
| 5 | Add Riebesell exposure curve | ExposureCurves.cs |
| 6 | Add Gaussian copula | Copulas.cs |
| 7 | Add Weibull, Beta, Exponential distributions | Distributions.cs |
