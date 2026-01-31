# Actuarial Add-In Validation Report

Generated: 2026-01-31 16:15:59

## Summary

- **Total Tests**: 80
- **Passed**: 80 (100.0%)
- **Failed**: 0

## Distributions

### Beta

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x03 | 2.160900 | 2.160900 | 0.00e+00 | scipy.stats.beta | PASS |
| pdf_x05 | 0.937500 | 0.937500 | 0.00e+00 | scipy.stats.beta | PASS |
| cdf_x03 | 0.579825 | 0.579825 | 0.00e+00 | scipy.stats.beta | PASS |
| inv_p50 | 0.264450 | 0.264450 | 0.00e+00 | scipy.stats.beta | PASS |

### Burr

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x1 | 0.375000 | 0.375000 | 0.00e+00 | scipy.stats.burr12 | PASS |
| cdf_x1 | 0.875000 | 0.875000 | 0.00e+00 | scipy.stats.burr12 | PASS |
| inv_p50 | 0.509825 | 0.509825 | 0.00e+00 | scipy.stats.burr12 | PASS |

### Exponential

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x1 | 0.270671 | 0.270671 | 0.00e+00 | scipy.stats.expon | PASS |
| cdf_x1 | 0.864665 | 0.864665 | 0.00e+00 | scipy.stats.expon | PASS |
| inv_p50 | 0.346574 | 0.346574 | 0.00e+00 | scipy.stats.expon | PASS |

### Gamma

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x1 | 0.367879 | 0.367879 | 0.00e+00 | scipy.stats.gamma | PASS |
| pdf_x2 | 0.270671 | 0.270671 | 0.00e+00 | scipy.stats.gamma | PASS |
| cdf_x2 | 0.593994 | 0.593994 | 0.00e+00 | scipy.stats.gamma | PASS |
| inv_p50 | 1.678347 | 1.678347 | 0.00e+00 | scipy.stats.gamma | PASS |

### Gpd

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x1 | 0.296296 | 0.296296 | 0.00e+00 | scipy.stats.genpareto | PASS |
| cdf_x1 | 0.555556 | 0.555556 | 0.00e+00 | scipy.stats.genpareto | PASS |
| inv_p50 | 0.828427 | 0.828427 | 0.00e+00 | scipy.stats.genpareto | PASS |

### Lognormal

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x1 | 0.398942 | 0.398942 | 0.00e+00 | scipy.stats.lognorm | PASS |
| pdf_x2 | 0.156874 | 0.156874 | 0.00e+00 | scipy.stats.lognorm | PASS |
| cdf_x1 | 0.500000 | 0.500000 | 0.00e+00 | scipy.stats.lognorm | PASS |
| cdf_x2 | 0.755891 | 0.755891 | 0.00e+00 | scipy.stats.lognorm | PASS |
| inv_p50 | 1.000000 | 1.000000 | 0.00e+00 | scipy.stats.lognorm | PASS |
| inv_p95 | 5.180252 | 5.180252 | 0.00e+00 | scipy.stats.lognorm | PASS |

### Negative Binomial

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_k5 | 0.051460 | 0.051460 | 0.00e+00 | scipy.stats.nbinom | PASS |
| pdf_k10 | 0.068710 | 0.068710 | 0.00e+00 | scipy.stats.nbinom | PASS |
| cdf_k10 | 0.484509 | 0.484509 | 0.00e+00 | scipy.stats.nbinom | PASS |
| inv_p50 | 11.000000 | 11.000000 | 0.00e+00 | scipy.stats.nbinom | PASS |

### Pareto

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x1 | 2.000000 | 2.000000 | 0.00e+00 | scipy.stats.pareto | PASS |
| pdf_x2 | 0.250000 | 0.250000 | 0.00e+00 | scipy.stats.pareto | PASS |
| cdf_x2 | 0.750000 | 0.750000 | 0.00e+00 | scipy.stats.pareto | PASS |
| inv_p50 | 1.414214 | 1.414214 | 0.00e+00 | scipy.stats.pareto | PASS |

### Poisson

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_k5 | 0.175467 | 0.175467 | 0.00e+00 | scipy.stats.poisson | PASS |
| pdf_k0 | 0.006738 | 0.006738 | 0.00e+00 | scipy.stats.poisson | PASS |
| pdf_k10 | 0.018133 | 0.018133 | 0.00e+00 | scipy.stats.poisson | PASS |
| cdf_k5 | 0.615961 | 0.615961 | 0.00e+00 | scipy.stats.poisson | PASS |
| cdf_k3 | 0.265026 | 0.265026 | 0.00e+00 | scipy.stats.poisson | PASS |
| inv_p50 | 5.000000 | 5.000000 | 0.00e+00 | scipy.stats.poisson | PASS |
| inv_p95 | 9.000000 | 9.000000 | 0.00e+00 | scipy.stats.poisson | PASS |

### Weibull

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| pdf_x05 | 0.778801 | 0.778801 | 0.00e+00 | scipy.stats.weibull_min | PASS |
| pdf_x1 | 0.735759 | 0.735759 | 0.00e+00 | scipy.stats.weibull_min | PASS |
| cdf_x1 | 0.632121 | 0.632121 | 0.00e+00 | scipy.stats.weibull_min | PASS |
| inv_p50 | 0.832555 | 0.832555 | 0.00e+00 | scipy.stats.weibull_min | PASS |

## Chain Ladder

### Development Factors

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| factor_1_to_2 | 3.4906 | 3.4906 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_2_to_3 | 1.7473 | 1.7473 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_3_to_4 | 1.4574 | 1.4574 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_4_to_5 | 1.1739 | 1.1739 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_5_to_6 | 1.1038 | 1.1038 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_6_to_7 | 1.0863 | 1.0863 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_7_to_8 | 1.0539 | 1.0539 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_8_to_9 | 1.0766 | 1.0766 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |
| factor_9_to_10 | 1.0177 | 1.0177 | 0.0000 | E&V (2002) / Taylor-Ashe | PASS |

### IBNR and Bootstrap

| Test | Expected | Actual | Diff | Source | Status | Notes |
|------|----------|--------|------|--------|--------|-------|
| total_ibnr | 18,680,856 | 18,680,856 | 0 | E&V (2002) | PASS |  |
| total_se_vs_ev_non_constant | 2,228,677 | 2,161,659 | 67,018 | E&V (2002) Table 3 | PASS | Match: 97.0% of E&V non-constant scale |

## Copulas

| Test | Expected | Actual | Diff | Source | Status |
|------|----------|--------|------|--------|--------|
| ACT_COPULA_CLAYTON_CDF cdf_0.5_0.5 | 0.3780 | 0.3780 | 0.0000 | Analytical formula | PASS |
| ACT_COPULA_CLAYTON_CDF cdf_0.3_0.7 | 0.2869 | 0.2869 | 0.0000 | Analytical formula | PASS |
| ACT_COPULA_TAIL_LOWER clayton_lower_tail | 0.7071 | 0.7071 | 0.0000 | lambda_L = 2^(-1/theta) | PASS |
| ACT_COPULA_GUMBEL_CDF cdf_0.5_0.5 | 0.3752 | 0.3752 | 0.0000 | Analytical formula | PASS |
| ACT_COPULA_TAIL_UPPER gumbel_upper_tail | 0.5858 | 0.5858 | 0.0000 | lambda_U = 2 - 2^(1/theta) | PASS |
| ACT_COPULA_FRANK_CDF cdf_0.5_0.5 | 0.3888 | 0.3888 | 0.0000 | Analytical formula | PASS |
| ACT_COPULA_TAIL_LOWER student_t_tail | 0.2666 | 0.2666 | 0.0000 | lambda = 2*t_{df+1}(-sqrt((df+1)(1-rho)/(1+rho))) | PASS |

## Exposure Curves

| Function | Test | Expected | Actual | Diff | Source | Status |
|----------|------|----------|--------|------|--------|--------|
| ACT_EXPOSURE_POWER | n=2_d=0.5 | 0.2500 | 0.2500 | 0.0000 | G(d) = d^n | PASS |
| ACT_EXPOSURE_POWER | n=0.5_d=0.5 | 0.7071 | 0.7071 | 0.0000 | G(d) = d^n | PASS |
| ACT_EXPOSURE_INVERSE_POWER | n=2_d=0.5 | 0.7500 | 0.7500 | 0.0000 | G(d) = 1 - (1-d)^n | PASS |
| ACT_EXPOSURE_INVERSE_POWER | n=0.5_d=0.5 | 0.2929 | 0.2929 | 0.0000 | G(d) = 1 - (1-d)^n | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.0 | 0.0000 | 0.0000 | 0.0000 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.1 | 0.1490 | 0.1490 | 0.0000 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.2 | 0.2784 | 0.2786 | 0.0002 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.3 | 0.3940 | 0.3941 | 0.0001 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.4 | 0.4993 | 0.4990 | 0.0003 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.5 | 0.5960 | 0.5957 | 0.0003 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.6 | 0.6860 | 0.6857 | 0.0003 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.7 | 0.7705 | 0.7704 | 0.0001 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.8 | 0.8507 | 0.8505 | 0.0002 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=0.9 | 0.9273 | 0.9269 | 0.0004 | Bernegger (1997) | PASS |
| ACT_EXPOSURE_MBBEFD | d=1.0 | 1.0000 | 1.0000 | 0.0000 | Bernegger (1997) | PASS |

## Credibility

| Function | Test | Expected | Actual | Diff | Source | Status |
|----------|------|----------|--------|------|--------|--------|
| ACT_CREDIBILITY_BUHLMANN | n=5_k=10 | 0.3330 | 0.3333 | 0.0003 | Z = n/(n+k) | PASS |
| ACT_CREDIBILITY_BUHLMANN | n=20_k=10 | 0.6670 | 0.6667 | 0.0003 | Z = n/(n+k) | PASS |
| ACT_CREDIBILITY_BUHLMANN | asymptotic_n=100 | 0.9090 | 0.9091 | 0.0001 | Z = n/(n+k) -> 1 | PASS |
| ACT_CREDIBILITY_BUHLMANN | asymptotic_n=1000 | 0.9900 | 0.9901 | 0.0001 | Z = n/(n+k) -> 1 | PASS |
| ACT_FULL_CREDIBILITY_STANDARD | p=0.9_r=0.05 | 1082 | 1082 | 0.2174 | k = (z_p/r)^2 * CV^2 | PASS |

## References

1. England, P.D. and Verrall, R.J. (2002). Stochastic claims reserving in general insurance. British Actuarial Journal 8(3): 443-518.
2. Mack, T. (1993). Distribution-free calculation of the standard error of chain ladder reserve estimates. ASTIN Bulletin 23(2): 213-225.
3. Bernegger, S. (1997). The Swiss Re Exposure Curves and the MBBEFD Distribution Class. ASTIN Bulletin 27(1): 99-111.
4. McNeil, A.J., Frey, R., and Embrechts, P. (2015). Quantitative Risk Management. Princeton University Press.
5. scipy.stats documentation: https://docs.scipy.org/doc/scipy/reference/stats.html
6. chainladder-python documentation: https://chainladder-python.readthedocs.io/
