"""
Generate LEV (Limited Expected Value) test fixtures using scipy.
LEV = E[min(X, d)] = integral of survival function from 0 to d
"""
import json
from scipy import stats
from scipy.integrate import quad
import numpy as np

def lev_from_frozen(frozen_dist, limit):
    """Calculate LEV via numerical integration of survival function from a frozen distribution."""
    # LEV = integral from 0 to limit of S(x) dx where S(x) = 1 - F(x)
    result, _ = quad(lambda x: frozen_dist.sf(x), 0, limit)
    return result

def lev_discrete(cdf_func, limit):
    """Calculate LEV for discrete distribution."""
    n = int(np.floor(limit))
    lev = sum(1 - cdf_func(k) for k in range(n))
    # Add fractional part
    frac = limit - n
    if frac > 0:
        lev += frac * (1 - cdf_func(n))
    return lev

# Generate LEV values for each distribution with consistent test parameters

lev_fixtures = {
    "_metadata": {
        "source": "scipy.stats numerical integration",
        "description": "LEV = E[min(X, d)] calculated via integral of survival function",
        "generated": "2026-01-31"
    }
}

# Exponential: lambda=2
print("Exponential LEV...")
exp_dist = stats.expon(scale=1/2)  # scipy uses scale=1/lambda
lev_fixtures["exponential"] = {
    "parameters": {"lambda": 2},
    "lev_limit_0.5": {"limit": 0.5, "expected": round(lev_from_frozen(exp_dist, 0.5), 10)},
    "lev_limit_1": {"limit": 1.0, "expected": round(lev_from_frozen(exp_dist, 1.0), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(exp_dist, 2.0), 10)},
}

# Pareto: alpha=2, xm=1
print("Pareto LEV...")
# scipy.stats.pareto with b=alpha, scale=xm
pareto_alpha, pareto_xm = 2, 1
pareto_dist = stats.pareto(b=pareto_alpha, scale=pareto_xm)
lev_fixtures["pareto"] = {
    "parameters": {"alpha": 2, "xm": 1},
    "lev_limit_1.5": {"limit": 1.5, "expected": round(lev_from_frozen(pareto_dist, 1.5), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(pareto_dist, 2.0), 10)},
    "lev_limit_5": {"limit": 5.0, "expected": round(lev_from_frozen(pareto_dist, 5.0), 10)},
}

# Lomax (Pareto Type II): alpha=2, lambda=1
print("Lomax LEV...")
lomax_alpha, lomax_lambda = 2, 1
lomax_dist = stats.lomax(c=lomax_alpha, scale=lomax_lambda)
lev_fixtures["lomax"] = {
    "parameters": {"alpha": 2, "lambda": 1},
    "lev_limit_1": {"limit": 1.0, "expected": round(lev_from_frozen(lomax_dist, 1.0), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(lomax_dist, 2.0), 10)},
    "lev_limit_5": {"limit": 5.0, "expected": round(lev_from_frozen(lomax_dist, 5.0), 10)},
}

# GPD: xi=0.5, sigma=1
print("GPD LEV...")
gpd_xi, gpd_sigma = 0.5, 1
gpd_dist = stats.genpareto(c=gpd_xi, scale=gpd_sigma)
lev_fixtures["gpd"] = {
    "parameters": {"xi": 0.5, "sigma": 1},
    "lev_limit_1": {"limit": 1.0, "expected": round(lev_from_frozen(gpd_dist, 1.0), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(gpd_dist, 2.0), 10)},
    "lev_limit_5": {"limit": 5.0, "expected": round(lev_from_frozen(gpd_dist, 5.0), 10)},
}

# Gamma: alpha=2, beta=1
print("Gamma LEV...")
gamma_alpha, gamma_beta = 2, 1
gamma_dist = stats.gamma(a=gamma_alpha, scale=1/gamma_beta)
lev_fixtures["gamma"] = {
    "parameters": {"alpha": 2, "beta": 1},
    "lev_limit_1": {"limit": 1.0, "expected": round(lev_from_frozen(gamma_dist, 1.0), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(gamma_dist, 2.0), 10)},
    "lev_limit_5": {"limit": 5.0, "expected": round(lev_from_frozen(gamma_dist, 5.0), 10)},
}

# Lognormal: mu=0, sigma=1
print("Lognormal LEV...")
lognorm_mu, lognorm_sigma = 0, 1
lognorm_dist = stats.lognorm(s=lognorm_sigma, scale=np.exp(lognorm_mu))
lev_fixtures["lognormal"] = {
    "parameters": {"mu": 0, "sigma": 1},
    "lev_limit_1": {"limit": 1.0, "expected": round(lev_from_frozen(lognorm_dist, 1.0), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(lognorm_dist, 2.0), 10)},
    "lev_limit_5": {"limit": 5.0, "expected": round(lev_from_frozen(lognorm_dist, 5.0), 10)},
}

# Weibull: k=2, lambda=1
print("Weibull LEV...")
weibull_k, weibull_lambda = 2, 1
weibull_dist = stats.weibull_min(c=weibull_k, scale=weibull_lambda)
lev_fixtures["weibull"] = {
    "parameters": {"k": 2, "lambda": 1},
    "lev_limit_0.5": {"limit": 0.5, "expected": round(lev_from_frozen(weibull_dist, 0.5), 10)},
    "lev_limit_1": {"limit": 1.0, "expected": round(lev_from_frozen(weibull_dist, 1.0), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(weibull_dist, 2.0), 10)},
}

# Beta: alpha=2, beta=5 (on [0,1])
print("Beta LEV...")
beta_a, beta_b = 2, 5
beta_dist = stats.beta(a=beta_a, b=beta_b)
lev_fixtures["beta"] = {
    "parameters": {"alpha": 2, "beta": 5},
    "lev_limit_0.25": {"limit": 0.25, "expected": round(lev_from_frozen(beta_dist, 0.25), 10)},
    "lev_limit_0.5": {"limit": 0.5, "expected": round(lev_from_frozen(beta_dist, 0.5), 10)},
    "lev_limit_1": {"limit": 1.0, "expected": round(beta_a / (beta_a + beta_b), 10)},  # Full mean
}

# Burr XII: c=2, k=3, lambda=1
print("Burr XII LEV...")
burr_c, burr_k, burr_lambda = 2, 3, 1
burr_dist = stats.burr12(c=burr_c, d=burr_k, scale=burr_lambda)
lev_fixtures["burr"] = {
    "parameters": {"c": 2, "k": 3, "lambda": 1},
    "lev_limit_0.5": {"limit": 0.5, "expected": round(lev_from_frozen(burr_dist, 0.5), 10)},
    "lev_limit_1": {"limit": 1.0, "expected": round(lev_from_frozen(burr_dist, 1.0), 10)},
    "lev_limit_2": {"limit": 2.0, "expected": round(lev_from_frozen(burr_dist, 2.0), 10)},
}

# Poisson: lambda=5 (discrete)
print("Poisson LEV...")
poisson_lambda = 5
poisson_cdf = lambda k: stats.poisson.cdf(k, mu=poisson_lambda)
lev_fixtures["poisson"] = {
    "parameters": {"lambda": 5},
    "lev_limit_3": {"limit": 3.0, "expected": round(lev_discrete(poisson_cdf, 3.0), 10)},
    "lev_limit_5": {"limit": 5.0, "expected": round(lev_discrete(poisson_cdf, 5.0), 10)},
    "lev_limit_10": {"limit": 10.0, "expected": round(lev_discrete(poisson_cdf, 10.0), 10)},
}

# Negative Binomial: r=5, p=0.3 (discrete)
print("Negative Binomial LEV...")
nbinom_r, nbinom_p = 5, 0.3
nbinom_cdf = lambda k: stats.nbinom.cdf(k, n=nbinom_r, p=nbinom_p)
lev_fixtures["negative_binomial"] = {
    "parameters": {"r": 5, "p": 0.3},
    "lev_limit_5": {"limit": 5.0, "expected": round(lev_discrete(nbinom_cdf, 5.0), 10)},
    "lev_limit_10": {"limit": 10.0, "expected": round(lev_discrete(nbinom_cdf, 10.0), 10)},
    "lev_limit_20": {"limit": 20.0, "expected": round(lev_discrete(nbinom_cdf, 20.0), 10)},
}

# Print results
print("\n=== LEV Fixtures ===\n")
print(json.dumps(lev_fixtures, indent=2))

# Save to file
with open('tests/fixtures/lev.json', 'w') as f:
    json.dump(lev_fixtures, f, indent=2)

print("\n\nSaved to tests/fixtures/lev.json")
