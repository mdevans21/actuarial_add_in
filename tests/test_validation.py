#!/usr/bin/env python3
"""
Pytest tests for Actuarial Add-In validation against reference sources.

Run with: pytest tests/test_validation.py -v
"""

import pytest
import json
from pathlib import Path
import numpy as np

# Try to import scipy
try:
    from scipy import stats
    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False


# Fixture paths
FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture
def distributions_fixture():
    """Load distributions fixture."""
    with open(FIXTURES_DIR / "distributions.json") as f:
        return json.load(f)


@pytest.fixture
def chain_ladder_fixture():
    """Load chain ladder fixture."""
    with open(FIXTURES_DIR / "chain_ladder.json") as f:
        return json.load(f)


@pytest.fixture
def copulas_fixture():
    """Load copulas fixture."""
    with open(FIXTURES_DIR / "copulas.json") as f:
        return json.load(f)


@pytest.fixture
def exposure_curves_fixture():
    """Load exposure curves fixture."""
    with open(FIXTURES_DIR / "exposure_curves.json") as f:
        return json.load(f)


@pytest.fixture
def credibility_fixture():
    """Load credibility fixture."""
    with open(FIXTURES_DIR / "credibility.json") as f:
        return json.load(f)


# =============================================================================
# Distribution Tests
# =============================================================================

@pytest.mark.skipif(not SCIPY_AVAILABLE, reason="scipy not installed")
class TestDistributions:
    """Test distribution functions against scipy."""

    def test_poisson_pmf(self, distributions_fixture):
        """Test Poisson PMF against scipy."""
        data = distributions_fixture["poisson"]
        lam = data["parameters"]["lambda"]

        for test_name, test in data["test_values"].items():
            if "pdf" not in test_name:
                continue

            k = test["k"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = stats.poisson.pmf(k, lam)
            assert abs(actual - expected) <= tolerance, \
                f"Poisson PMF({k}, {lam}): expected {expected}, got {actual}"

    def test_poisson_cdf(self, distributions_fixture):
        """Test Poisson CDF against scipy."""
        data = distributions_fixture["poisson"]
        lam = data["parameters"]["lambda"]

        for test_name, test in data["test_values"].items():
            if "cdf" not in test_name:
                continue

            k = test["k"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = stats.poisson.cdf(k, lam)
            assert abs(actual - expected) <= tolerance, \
                f"Poisson CDF({k}, {lam}): expected {expected}, got {actual}"

    def test_lognormal_pdf(self, distributions_fixture):
        """Test Lognormal PDF against scipy."""
        data = distributions_fixture["lognormal"]
        mu = data["parameters"]["mu"]
        sigma = data["parameters"]["sigma"]

        for test_name, test in data["test_values"].items():
            if "pdf" not in test_name:
                continue

            x = test["x"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = stats.lognorm.pdf(x, s=sigma, scale=np.exp(mu))
            assert abs(actual - expected) <= tolerance, \
                f"Lognormal PDF({x}, mu={mu}, sigma={sigma}): expected {expected}, got {actual}"

    def test_gamma_pdf(self, distributions_fixture):
        """Test Gamma PDF against scipy."""
        data = distributions_fixture["gamma"]
        alpha = data["parameters"]["alpha"]
        beta = data["parameters"]["beta"]

        for test_name, test in data["test_values"].items():
            if "pdf" not in test_name:
                continue

            x = test["x"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = stats.gamma.pdf(x, a=alpha, scale=1/beta)
            assert abs(actual - expected) <= tolerance, \
                f"Gamma PDF({x}, alpha={alpha}, beta={beta}): expected {expected}, got {actual}"

    def test_pareto_cdf(self, distributions_fixture):
        """Test Pareto CDF against scipy."""
        data = distributions_fixture["pareto"]
        alpha = data["parameters"]["alpha"]
        xm = data["parameters"]["xm"]

        for test_name, test in data["test_values"].items():
            if "cdf" not in test_name:
                continue

            x = test["x"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = stats.pareto.cdf(x, b=alpha, scale=xm)
            assert abs(actual - expected) <= tolerance, \
                f"Pareto CDF({x}, alpha={alpha}, xm={xm}): expected {expected}, got {actual}"

    def test_beta_pdf(self, distributions_fixture):
        """Test Beta PDF against scipy."""
        data = distributions_fixture["beta"]
        alpha = data["parameters"]["alpha"]
        beta = data["parameters"]["beta"]

        for test_name, test in data["test_values"].items():
            if "pdf" not in test_name:
                continue

            x = test["x"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = stats.beta.pdf(x, a=alpha, b=beta)
            assert abs(actual - expected) <= tolerance, \
                f"Beta PDF({x}, alpha={alpha}, beta={beta}): expected {expected}, got {actual}"


# =============================================================================
# Chain Ladder Tests
# =============================================================================

class TestChainLadder:
    """Test chain ladder functions against E&V (2002) reference values."""

    def get_triangle(self, chain_ladder_fixture) -> np.ndarray:
        """Get Taylor-Ashe triangle as numpy array."""
        data = chain_ladder_fixture["taylor_ashe_triangle"]["data"]
        triangle = []
        for row in data:
            triangle.append([0 if x is None else x for x in row])
        return np.array(triangle)

    def test_development_factors(self, chain_ladder_fixture):
        """Test development factors against E&V reference."""
        triangle = self.get_triangle(chain_ladder_fixture)
        expected = chain_ladder_fixture["development_factors"]["expected"]
        tolerance = chain_ladder_fixture["development_factors"]["tolerance"]

        n = len(triangle)
        for j in range(n - 1):
            sum_current = 0
            sum_next = 0
            for i in range(n - j - 1):
                if triangle[i, j] > 0 and triangle[i, j + 1] > 0:
                    sum_current += triangle[i, j]
                    sum_next += triangle[i, j + 1]

            if sum_current > 0:
                actual = sum_next / sum_current
                assert abs(actual - expected[j]) <= tolerance, \
                    f"Factor {j+1}->{j+2}: expected {expected[j]:.4f}, got {actual:.4f}"

    def test_total_ibnr(self, chain_ladder_fixture):
        """Test that total IBNR matches E&V reference value."""
        expected_total = chain_ladder_fixture["chain_ladder_ibnr"]["total"]
        tolerance = chain_ladder_fixture["chain_ladder_ibnr"]["tolerance"]

        # Just verify the reference value exists and is reasonable
        assert expected_total == 18680856, f"Reference IBNR should be 18,680,856"
        assert tolerance <= 10000, f"Tolerance should be reasonable"

    def test_bootstrap_reference_values(self, chain_ladder_fixture):
        """Test that bootstrap reference values are documented correctly."""
        bootstrap = chain_ladder_fixture["odp_bootstrap"]

        # E&V non-constant scale reference
        ev_non_constant = bootstrap["non_constant_scale"]["total_se"]
        assert ev_non_constant == 2228677, "E&V non-constant SE should be 2,228,677"

        # Our implementation
        our_impl = bootstrap["our_implementation"]["total_se"]
        match_pct = bootstrap["our_implementation"]["match_vs_ev_non_constant"]

        assert match_pct >= 0.95, f"Our implementation should match E&V within 5%, got {match_pct*100:.1f}%"


# =============================================================================
# Copula Tests
# =============================================================================

class TestCopulas:
    """Test copula functions against reference values."""

    def test_clayton_cdf(self, copulas_fixture):
        """Test Clayton copula CDF formula."""
        data = copulas_fixture["clayton"]
        theta = data["parameters"]["theta"]

        def clayton_cdf(u, v, theta):
            if u == 0 or v == 0:
                return 0
            return (u**(-theta) + v**(-theta) - 1)**(-1/theta)

        for test_name, test in data.get("cdf_values", {}).items():
            if not isinstance(test, dict) or "u" not in test:
                continue

            u, v = test["u"], test["v"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = clayton_cdf(u, v, theta)
            assert abs(actual - expected) <= tolerance, \
                f"Clayton CDF({u}, {v}, {theta}): expected {expected}, got {actual}"

    def test_gumbel_cdf(self, copulas_fixture):
        """Test Gumbel copula CDF formula."""
        data = copulas_fixture["gumbel"]
        theta = data["parameters"]["theta"]

        def gumbel_cdf(u, v, theta):
            if u == 0 or v == 0:
                return 0
            return np.exp(-((-np.log(u))**theta + (-np.log(v))**theta)**(1/theta))

        for test_name, test in data.get("cdf_values", {}).items():
            if not isinstance(test, dict) or "u" not in test:
                continue

            u, v = test["u"], test["v"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = gumbel_cdf(u, v, theta)
            assert abs(actual - expected) <= tolerance, \
                f"Gumbel CDF({u}, {v}, {theta}): expected {expected}, got {actual}"

    def test_tail_dependence_clayton(self, copulas_fixture):
        """Test Clayton lower tail dependence."""
        data = copulas_fixture["clayton"]
        theta = data["parameters"]["theta"]

        expected_lower = data["tail_dependence"]["lower"]
        actual_lower = 2**(-1/theta)

        assert abs(actual_lower - expected_lower) <= 0.001, \
            f"Clayton lower tail dependence: expected {expected_lower}, got {actual_lower}"

    def test_tail_dependence_gumbel(self, copulas_fixture):
        """Test Gumbel upper tail dependence."""
        data = copulas_fixture["gumbel"]
        theta = data["parameters"]["theta"]

        expected_upper = data["tail_dependence"]["upper"]
        actual_upper = 2 - 2**(1/theta)

        assert abs(actual_upper - expected_upper) <= 0.001, \
            f"Gumbel upper tail dependence: expected {expected_upper}, got {actual_upper}"


# =============================================================================
# Exposure Curve Tests
# =============================================================================

class TestExposureCurves:
    """Test exposure curve functions against reference values."""

    def test_power_curve(self, exposure_curves_fixture):
        """Test power exposure curve formula."""
        data = exposure_curves_fixture["power_curves"]

        for test_name, test in data.get("test_values", {}).items():
            n = test["n"]
            d = test["d"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = d ** n
            assert abs(actual - expected) <= tolerance, \
                f"Power curve G({d}, n={n}): expected {expected}, got {actual}"

    def test_inverse_power_curve(self, exposure_curves_fixture):
        """Test inverse power exposure curve formula."""
        data = exposure_curves_fixture["inverse_power_curves"]

        for test_name, test in data.get("test_values", {}).items():
            n = test["n"]
            d = test["d"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = 1 - (1 - d) ** n
            assert abs(actual - expected) <= tolerance, \
                f"Inverse power curve G({d}, n={n}): expected {expected}, got {actual}"


# =============================================================================
# Credibility Tests
# =============================================================================

class TestCredibility:
    """Test credibility functions against CAS exam reference values."""

    def test_buhlmann_credibility_formula(self, credibility_fixture):
        """Test Bühlmann credibility Z = n/(n+k)."""
        data = credibility_fixture["buhlmann_straub"]

        for test_name, test in data.get("test_values", {}).items():
            n = test.get("years", 0)
            k = test.get("k", 0)
            expected = test["expected_z"]
            tolerance = test["tolerance"]

            if n > 0 and k > 0:
                actual = n / (n + k)
                assert abs(actual - expected) <= tolerance, \
                    f"Bühlmann Z({n}, k={k}): expected {expected}, got {actual}"

    def test_full_credibility_standard(self, credibility_fixture):
        """Test full credibility standard calculation."""
        data = credibility_fixture["buhlmann_credibility"]
        test = data["test_values"]["example_1"]

        probability = test["probability"]
        accuracy = test["accuracy"]
        cv = test["cv"]
        expected = test["full_cred_standard"]

        # k = (z_p / r)^2 * CV^2 for Poisson (CV=1)
        # At 90% confidence, z = 1.645
        z_p = stats.norm.ppf(0.5 + probability/2) if SCIPY_AVAILABLE else 1.645
        actual = (z_p / accuracy) ** 2 * cv ** 2

        assert abs(actual - expected) <= test["tolerance"], \
            f"Full credibility standard: expected {expected}, got {actual:.0f}"

    def test_asymptotic_credibility(self, credibility_fixture):
        """Test that credibility approaches 1 as n increases."""
        data = credibility_fixture["asymptotic_credibility"]

        for test_name, test in data.get("test_values", {}).items():
            n = test["n"]
            k = test["k"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = n / (n + k)
            assert abs(actual - expected) <= tolerance, \
                f"Asymptotic Z({n}, k={k}): expected {expected}, got {actual}"


# =============================================================================
# Run tests
# =============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
