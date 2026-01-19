#!/usr/bin/env python3
"""
ODP Bootstrap Chain Ladder - Compare with England & Verrall (2002)

CORRECTED: Use raw Pearson residuals with sqrt(fitted), not sqrt(fitted * phi)
"""

import numpy as np

print("=" * 70)
print("ODP BOOTSTRAP - CORRECTED VERSION")
print("=" * 70)

# Taylor-Ashe triangle data (cumulative)
triangle = np.array([
    [357848, 1124788, 1735330, 2218270, 2745596, 3319994, 3466336, 3606286, 3833515, 3901463],
    [352118, 1236139, 2170033, 3353322, 3799067, 4120063, 4647867, 4914039, 5339085, np.nan],
    [290507, 1292306, 2218525, 3235179, 3985995, 4132918, 4628910, 4909315, np.nan, np.nan],
    [310608, 1418858, 2195047, 3757447, 4029929, 4381982, 4588268, np.nan, np.nan, np.nan],
    [443160, 1136350, 2128333, 2897821, 3402672, 3873311, np.nan, np.nan, np.nan, np.nan],
    [396132, 1333217, 2180715, 2985752, 3691712, np.nan, np.nan, np.nan, np.nan, np.nan],
    [440832, 1288463, 2419861, 3483130, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
    [359480, 1421128, 2864498, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
    [376686, 1363294, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
    [344014, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan]
], dtype=float)

n = 10

# England & Verrall Reference Values
EV_SE = [0, 112379, 178443, 209399, 286636, 440310, 571035, 820842, 1258296, 2046223]
EV_TOTAL_SE = 3087570
EV_TOTAL_IBNR = 18680856

# Calculate factors
factors = []
for j in range(n - 1):
    sum_current = sum_next = 0
    for i in range(n - j - 1):
        if not np.isnan(triangle[i, j]) and not np.isnan(triangle[i, j + 1]):
            sum_current += triangle[i, j]
            sum_next += triangle[i, j + 1]
    factors.append(sum_next / sum_current if sum_current > 0 else 1.0)

print("\nDevelopment Factors:")
for j, f in enumerate(factors):
    print(f"  {j+1}-{j+2}: {f:.4f}")

# Calculate incremental and fitted
incremental = np.full_like(triangle, np.nan)
fitted = np.full_like(triangle, np.nan)
for i in range(n):
    for j in range(n - i):
        if j == 0:
            incremental[i, j] = triangle[i, j]
            fitted[i, j] = triangle[i, j]
        else:
            incremental[i, j] = triangle[i, j] - triangle[i, j - 1]
            fitted[i, j] = triangle[i, j - 1] * (factors[j - 1] - 1.0)

# Calculate Pearson residuals (j > 0 only)
# r_ij = (actual - fitted) / sqrt(fitted)
residuals = []
for i in range(n):
    for j in range(1, n - i):
        if fitted[i, j] > 0:
            resid = (incremental[i, j] - fitted[i, j]) / np.sqrt(fitted[i, j])
            residuals.append(resid)

residuals = np.array(residuals)
n_resid = len(residuals)
n_params = n - 1  # number of development factors
df = n_resid - n_params

# Scale parameter phi = sum(r^2) / df
phi = np.sum(residuals ** 2) / df

print(f"\nPhi (scale parameter): {phi:.2f}")
print(f"Number of residuals: {n_resid}")
print(f"Degrees of freedom: {df}")
print(f"Raw residuals: mean={np.mean(residuals):.2f}, std={np.std(residuals):.2f}")

# Adjust residuals for degrees of freedom (E&V formula)
# r_adj = r * sqrt(n / (n - p))
adj_factor = np.sqrt(n_resid / df)
adj_residuals = residuals * adj_factor

print(f"Adjusted residuals: mean={np.mean(adj_residuals):.2f}, std={np.std(adj_residuals):.2f}")

# Bootstrap
print("\n" + "=" * 70)
print("BOOTSTRAP (10,000 iterations)")
print("=" * 70)

np.random.seed(42)
n_sims = 10000
reserves = np.zeros((n_sims, n))

for sim in range(n_sims):
    # Resample adjusted residuals
    resampled = np.random.choice(adj_residuals, size=n_resid, replace=True)
    
    # Create pseudo-incremental triangle
    # m*_ij = m_ij + r_adj * sqrt(m_ij)  [NOT sqrt(m_ij * phi)!]
    pseudo_incr = np.full_like(triangle, np.nan)
    resid_idx = 0
    for i in range(n):
        for j in range(n - i):
            if j == 0:
                pseudo_incr[i, j] = fitted[i, j]  # First column unchanged
            else:
                mean = fitted[i, j]
                if mean > 0:
                    r = resampled[resid_idx]
                    resid_idx += 1
                    # CORRECTED: sqrt(mean), not sqrt(mean * phi)
                    pseudo_incr[i, j] = mean + r * np.sqrt(mean)
                else:
                    pseudo_incr[i, j] = mean
    
    # Convert to cumulative
    pseudo_cum = np.full_like(triangle, np.nan)
    for i in range(n):
        cumsum = 0
        for j in range(n - i):
            cumsum += pseudo_incr[i, j]
            pseudo_cum[i, j] = cumsum
    
    # Re-estimate factors from pseudo-triangle
    pseudo_factors = []
    for j in range(n - 1):
        sum_c = sum_n = 0
        for i in range(n - j - 1):
            if pseudo_cum[i, j] > 0:
                sum_c += pseudo_cum[i, j]
                sum_n += pseudo_cum[i, j + 1]
        pseudo_factors.append(sum_n / sum_c if sum_c > 0 else 1.0)
    
    # Project ORIGINAL triangle using pseudo-factors + process variance
    for i in range(n):
        last_col = n - 1 - i
        if last_col == n - 1:
            reserves[sim, i] = 0
        else:
            current = triangle[i, last_col]
            for j in range(last_col, n - 1):
                mean_next = current * pseudo_factors[j]
                mean_incr = mean_next - current
                
                if mean_incr > 0:
                    # Process variance using Gamma
                    # E[incr] = mean_incr, Var[incr] = mean_incr * phi
                    # Gamma: shape = mean^2/var = mean_incr/phi, scale = var/mean = phi
                    shape = mean_incr / phi
                    scale = phi
                    if shape > 0:
                        sim_incr = np.random.gamma(shape, scale)
                        current = current + sim_incr
                    else:
                        current = mean_next
                else:
                    current = mean_next
            
            reserves[sim, i] = current - triangle[i, last_col]

print(f"\n{'AY':<5} {'Mean':>12} {'StdDev':>12} {'E&V SE':>12} {'Ratio':>8}")
print("-" * 50)
for i in range(n):
    mean = np.mean(reserves[:, i])
    std = np.std(reserves[:, i])
    ev_se = EV_SE[i]
    ratio = std / ev_se if ev_se > 0 else 0
    print(f"{i+1:<5} {mean:>12,.0f} {std:>12,.0f} {ev_se:>12,.0f} {ratio:>8.2f}")

total = reserves.sum(axis=1)
print("-" * 50)
print(f"{'Total':<5} {np.mean(total):>12,.0f} {np.std(total):>12,.0f} {EV_TOTAL_SE:>12,.0f} {np.std(total)/EV_TOTAL_SE:>8.2f}")

print(f"\nPercentiles:")
print(f"{'Stat':<8} {'Value':>15} {'E&V Ref':>15}")
print(f"{'Mean':<8} {np.mean(total):>15,.0f} {EV_TOTAL_IBNR:>15,.0f}")
print(f"{'StdDev':<8} {np.std(total):>15,.0f} {EV_TOTAL_SE:>15,.0f}")
print(f"{'P50':<8} {np.percentile(total, 50):>15,.0f} {18186033:>15,.0f}")
print(f"{'P75':<8} {np.percentile(total, 75):>15,.0f} {20376564:>15,.0f}")
print(f"{'P95':<8} {np.percentile(total, 95):>15,.0f} {24221956:>15,.0f}")
