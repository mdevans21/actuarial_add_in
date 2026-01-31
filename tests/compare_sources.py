#!/usr/bin/env python3
"""
Comparison Framework for Actuarial Add-In Validation

This module provides tools for comparing function outputs across:
- Excel formulas (via add-in)
- Python implementations (scipy, chainladder)
- Online reference values (E&V 2002, academic papers)

Usage:
    python compare_sources.py --generate-report
    python compare_sources.py --validate-distributions
    python compare_sources.py --validate-chain-ladder
"""

import json
import os
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime
import numpy as np

# Try to import optional dependencies
try:
    from scipy import stats
    SCIPY_AVAILABLE = True
except ImportError:
    SCIPY_AVAILABLE = False

try:
    import chainladder as cl
    CHAINLADDER_AVAILABLE = True
except ImportError:
    CHAINLADDER_AVAILABLE = False


@dataclass
class ValidationResult:
    """Result of validating a single function/value."""
    function_name: str
    test_name: str
    expected: float
    actual: Optional[float]
    source: str
    tolerance: float
    passed: bool
    difference: Optional[float] = None
    notes: str = ""

    @property
    def status_icon(self) -> str:
        if self.actual is None:
            return "?"
        return "PASS" if self.passed else "FAIL"


@dataclass
class FunctionValidation:
    """Validation results for a single function."""
    function_name: str
    category: str
    results: List[ValidationResult] = field(default_factory=list)

    @property
    def all_passed(self) -> bool:
        return all(r.passed for r in self.results if r.actual is not None)

    @property
    def pass_count(self) -> int:
        return sum(1 for r in self.results if r.passed and r.actual is not None)

    @property
    def total_count(self) -> int:
        return len([r for r in self.results if r.actual is not None])


class FixtureLoader:
    """Load and parse fixture files."""

    def __init__(self, fixtures_dir: str = None):
        if fixtures_dir is None:
            fixtures_dir = Path(__file__).parent / "fixtures"
        self.fixtures_dir = Path(fixtures_dir)

    def load(self, filename: str) -> Dict:
        """Load a fixture JSON file."""
        filepath = self.fixtures_dir / filename
        if not filepath.exists():
            raise FileNotFoundError(f"Fixture not found: {filepath}")
        with open(filepath, 'r') as f:
            return json.load(f)

    def load_distributions(self) -> Dict:
        return self.load("distributions.json")

    def load_chain_ladder(self) -> Dict:
        return self.load("chain_ladder.json")

    def load_copulas(self) -> Dict:
        return self.load("copulas.json")

    def load_exposure_curves(self) -> Dict:
        return self.load("exposure_curves.json")

    def load_credibility(self) -> Dict:
        return self.load("credibility.json")

    def load_sources(self) -> Dict:
        return self.load("sources.json")


class ScipyValidator:
    """Validate distribution functions against scipy."""

    def __init__(self, fixtures: Dict):
        self.fixtures = fixtures

    def _validate_distribution(self, dist_name: str, scipy_dist, param_map: callable) -> List[ValidationResult]:
        """Generic distribution validation."""
        results = []
        if not SCIPY_AVAILABLE:
            return results

        data = self.fixtures.get(dist_name, {})
        if not data:
            return results

        params = data.get("parameters", {})

        for test_name, test in data.get("test_values", {}).items():
            expected = test.get("expected")
            tolerance = test.get("tolerance", 1e-10)
            actual = None

            try:
                if "pdf" in test_name:
                    x = test.get("x") if "x" in test else test.get("k")
                    actual = scipy_dist.pdf(x, **param_map(params)) if hasattr(scipy_dist, 'pdf') else scipy_dist.pmf(x, **param_map(params))
                elif "cdf" in test_name:
                    x = test.get("x") if "x" in test else test.get("k")
                    actual = scipy_dist.cdf(x, **param_map(params))
                elif "inv" in test_name:
                    p = test.get("p")
                    actual = scipy_dist.ppf(p, **param_map(params))
                else:
                    continue
            except Exception as e:
                actual = None

            if actual is not None and expected is not None:
                diff = abs(actual - expected)
                passed = diff <= tolerance
            else:
                diff = None
                passed = False

            results.append(ValidationResult(
                function_name=f"ACT_DIST_{dist_name.upper()}_*",
                test_name=test_name,
                expected=expected,
                actual=actual,
                source=f"scipy.stats.{scipy_dist.name if hasattr(scipy_dist, 'name') else dist_name}",
                tolerance=tolerance,
                passed=passed,
                difference=diff
            ))

        return results

    def validate_all_distributions(self) -> Dict[str, List[ValidationResult]]:
        """Validate all distributions against scipy."""
        results = {}

        if not SCIPY_AVAILABLE:
            return results

        # Poisson
        results["poisson"] = self._validate_distribution(
            "poisson", stats.poisson,
            lambda p: {"mu": p.get("lambda", 5)}
        )

        # Negative Binomial
        results["negative_binomial"] = self._validate_distribution(
            "negative_binomial", stats.nbinom,
            lambda p: {"n": p.get("r", 5), "p": p.get("p", 0.3)}
        )

        # Lognormal
        results["lognormal"] = self._validate_distribution(
            "lognormal", stats.lognorm,
            lambda p: {"s": p.get("sigma", 1), "scale": np.exp(p.get("mu", 0))}
        )

        # Gamma
        results["gamma"] = self._validate_distribution(
            "gamma", stats.gamma,
            lambda p: {"a": p.get("alpha", 2), "scale": 1/p.get("beta", 1)}
        )

        # Pareto
        results["pareto"] = self._validate_distribution(
            "pareto", stats.pareto,
            lambda p: {"b": p.get("alpha", 2), "scale": p.get("xm", 1)}
        )

        # Weibull
        results["weibull"] = self._validate_distribution(
            "weibull", stats.weibull_min,
            lambda p: {"c": p.get("k", 2), "scale": p.get("lambda", 1)}
        )

        # Beta
        results["beta"] = self._validate_distribution(
            "beta", stats.beta,
            lambda p: {"a": p.get("alpha", 2), "b": p.get("beta", 5)}
        )

        # Exponential
        results["exponential"] = self._validate_distribution(
            "exponential", stats.expon,
            lambda p: {"scale": 1/p.get("lambda", 2)}
        )

        # GPD
        results["gpd"] = self._validate_distribution(
            "gpd", stats.genpareto,
            lambda p: {"c": p.get("xi", 0.5), "scale": p.get("sigma", 1)}
        )

        # Burr
        results["burr"] = self._validate_distribution(
            "burr", stats.burr12,
            lambda p: {"c": p.get("c", 2), "d": p.get("k", 3), "scale": p.get("lambda", 1)}
        )

        # Filter out empty results
        results = {k: v for k, v in results.items() if v}

        return results


class CopulaValidator:
    """Validate copula functions against analytical formulas."""

    def __init__(self, fixtures: Dict):
        self.fixtures = fixtures

    def validate_all(self) -> List[ValidationResult]:
        """Validate all copula functions."""
        results = []

        # Clayton CDF
        clayton = self.fixtures.get("clayton", {})
        theta = clayton.get("parameters", {}).get("theta", 2.0)

        for test_name, test in clayton.get("cdf_values", {}).items():
            if not isinstance(test, dict) or "u" not in test:
                continue

            u, v = test["u"], test["v"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            # Clayton CDF formula
            if u == 0 or v == 0:
                actual = 0
            else:
                actual = (u**(-theta) + v**(-theta) - 1)**(-1/theta)

            diff = abs(actual - expected)
            passed = diff <= tolerance

            results.append(ValidationResult(
                function_name="ACT_COPULA_CLAYTON_CDF",
                test_name=f"cdf_{u}_{v}",
                expected=expected,
                actual=actual,
                source="Analytical formula",
                tolerance=tolerance,
                passed=passed,
                difference=diff
            ))

        # Clayton tail dependence
        tail = clayton.get("tail_dependence", {})
        expected_lower = tail.get("lower", 0)
        actual_lower = 2**(-1/theta)
        diff = abs(actual_lower - expected_lower)
        results.append(ValidationResult(
            function_name="ACT_COPULA_TAIL_LOWER",
            test_name="clayton_lower_tail",
            expected=expected_lower,
            actual=actual_lower,
            source="lambda_L = 2^(-1/theta)",
            tolerance=0.001,
            passed=diff <= 0.001,
            difference=diff
        ))

        # Gumbel CDF
        gumbel = self.fixtures.get("gumbel", {})
        theta_g = gumbel.get("parameters", {}).get("theta", 2.0)

        for test_name, test in gumbel.get("cdf_values", {}).items():
            if not isinstance(test, dict) or "u" not in test:
                continue

            u, v = test["u"], test["v"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            # Gumbel CDF formula
            if u == 0 or v == 0:
                actual = 0
            else:
                actual = np.exp(-((-np.log(u))**theta_g + (-np.log(v))**theta_g)**(1/theta_g))

            diff = abs(actual - expected)
            passed = diff <= tolerance

            results.append(ValidationResult(
                function_name="ACT_COPULA_GUMBEL_CDF",
                test_name=f"cdf_{u}_{v}",
                expected=expected,
                actual=actual,
                source="Analytical formula",
                tolerance=tolerance,
                passed=passed,
                difference=diff
            ))

        # Gumbel tail dependence
        tail_g = gumbel.get("tail_dependence", {})
        expected_upper = tail_g.get("upper", 0)
        actual_upper = 2 - 2**(1/theta_g)
        diff = abs(actual_upper - expected_upper)
        results.append(ValidationResult(
            function_name="ACT_COPULA_TAIL_UPPER",
            test_name="gumbel_upper_tail",
            expected=expected_upper,
            actual=actual_upper,
            source="lambda_U = 2 - 2^(1/theta)",
            tolerance=0.001,
            passed=diff <= 0.001,
            difference=diff
        ))

        # Frank CDF
        frank = self.fixtures.get("frank", {})
        theta_f = frank.get("parameters", {}).get("theta", 5.736)

        for test_name, test in frank.get("cdf_values", {}).items():
            if not isinstance(test, dict) or "u" not in test:
                continue

            u, v = test["u"], test["v"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            # Frank CDF formula
            if abs(theta_f) < 1e-10:
                actual = u * v
            else:
                actual = -np.log(1 + (np.exp(-theta_f*u)-1)*(np.exp(-theta_f*v)-1)/(np.exp(-theta_f)-1)) / theta_f

            diff = abs(actual - expected)
            passed = diff <= tolerance

            results.append(ValidationResult(
                function_name="ACT_COPULA_FRANK_CDF",
                test_name=f"cdf_{u}_{v}",
                expected=expected,
                actual=actual,
                source="Analytical formula",
                tolerance=tolerance,
                passed=passed,
                difference=diff
            ))

        # Student-t tail dependence
        student_t = self.fixtures.get("student_t", {})
        if SCIPY_AVAILABLE and student_t:
            params = student_t.get("parameters", {})
            rho = params.get("rho", 0.6)
            df = params.get("df", 5)

            tail_t = student_t.get("tail_dependence", {})
            expected_tail = tail_t.get("lower", 0)
            tolerance = tail_t.get("tolerance", 0.01)

            arg = -np.sqrt((df + 1) * (1 - rho) / (1 + rho))
            actual_tail = 2 * stats.t.cdf(arg, df + 1)

            diff = abs(actual_tail - expected_tail)
            results.append(ValidationResult(
                function_name="ACT_COPULA_TAIL_LOWER",
                test_name="student_t_tail",
                expected=expected_tail,
                actual=actual_tail,
                source="lambda = 2*t_{df+1}(-sqrt((df+1)(1-rho)/(1+rho)))",
                tolerance=tolerance,
                passed=diff <= tolerance,
                difference=diff
            ))

        return results


class ExposureCurveValidator:
    """Validate exposure curve functions."""

    def __init__(self, fixtures: Dict):
        self.fixtures = fixtures

    def validate_all(self) -> List[ValidationResult]:
        """Validate all exposure curve functions."""
        results = []

        # Power curves
        power = self.fixtures.get("power_curves", {})
        for test_name, test in power.get("test_values", {}).items():
            n = test["n"]
            d = test["d"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = d ** n
            diff = abs(actual - expected)

            results.append(ValidationResult(
                function_name="ACT_EXPOSURE_POWER",
                test_name=f"n={n}_d={d}",
                expected=expected,
                actual=actual,
                source="G(d) = d^n",
                tolerance=tolerance,
                passed=diff <= tolerance,
                difference=diff
            ))

        # Inverse power curves
        inv_power = self.fixtures.get("inverse_power_curves", {})
        for test_name, test in inv_power.get("test_values", {}).items():
            n = test["n"]
            d = test["d"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = 1 - (1 - d) ** n
            diff = abs(actual - expected)

            results.append(ValidationResult(
                function_name="ACT_EXPOSURE_INVERSE_POWER",
                test_name=f"n={n}_d={d}",
                expected=expected,
                actual=actual,
                source="G(d) = 1 - (1-d)^n",
                tolerance=tolerance,
                passed=diff <= tolerance,
                difference=diff
            ))

        # MBBEFD curves
        mbbefd = self.fixtures.get("mbbefd", {})
        params = mbbefd.get("parameters", {})
        b = params.get("b", 2)
        g = params.get("g", 3)

        for test_name, test in mbbefd.get("test_values", {}).items():
            d = test["d"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            # MBBEFD formula (Bernegger 1997, equation 3)
            # G(d) = [ln(a + b^d) - ln(a+1)] / [ln(a+b) - ln(a+1)]
            # where a = (g - b^g) / (b^g - 1)
            if d == 0:
                actual = 0
            elif d == 1:
                actual = 1
            else:
                bg = b ** g
                a = (g - bg) / (bg - 1)
                numerator = np.log(a + b ** d) - np.log(a + 1)
                denominator = np.log(a + b) - np.log(a + 1)
                actual = numerator / denominator

            diff = abs(actual - expected)

            results.append(ValidationResult(
                function_name="ACT_EXPOSURE_MBBEFD",
                test_name=f"d={d}",
                expected=expected,
                actual=actual,
                source="Bernegger (1997)",
                tolerance=tolerance,
                passed=diff <= tolerance,
                difference=diff
            ))

        return results


class CredibilityValidator:
    """Validate credibility functions."""

    def __init__(self, fixtures: Dict):
        self.fixtures = fixtures

    def validate_all(self) -> List[ValidationResult]:
        """Validate all credibility functions."""
        results = []

        # Bühlmann-Straub credibility
        bs = self.fixtures.get("buhlmann_straub", {})
        for test_name, test in bs.get("test_values", {}).items():
            n = test.get("years", 0)
            k = test.get("k", 0)
            expected = test["expected_z"]
            tolerance = test["tolerance"]

            if n > 0 and k > 0:
                actual = n / (n + k)
                diff = abs(actual - expected)

                results.append(ValidationResult(
                    function_name="ACT_CREDIBILITY_BUHLMANN",
                    test_name=f"n={n}_k={k}",
                    expected=expected,
                    actual=actual,
                    source="Z = n/(n+k)",
                    tolerance=tolerance,
                    passed=diff <= tolerance,
                    difference=diff
                ))

        # Asymptotic credibility
        asym = self.fixtures.get("asymptotic_credibility", {})
        for test_name, test in asym.get("test_values", {}).items():
            n = test["n"]
            k = test["k"]
            expected = test["expected"]
            tolerance = test["tolerance"]

            actual = n / (n + k)
            diff = abs(actual - expected)

            results.append(ValidationResult(
                function_name="ACT_CREDIBILITY_BUHLMANN",
                test_name=f"asymptotic_n={n}",
                expected=expected,
                actual=actual,
                source="Z = n/(n+k) -> 1",
                tolerance=tolerance,
                passed=diff <= tolerance,
                difference=diff
            ))

        # Full credibility standard
        if SCIPY_AVAILABLE:
            buhl = self.fixtures.get("buhlmann_credibility", {})
            test = buhl.get("test_values", {}).get("example_1", {})
            if test:
                probability = test.get("probability", 0.90)
                accuracy = test.get("accuracy", 0.05)
                cv = test.get("cv", 1.0)
                expected = test.get("full_cred_standard", 1082)
                tolerance = test.get("tolerance", 10)

                z_p = stats.norm.ppf(0.5 + probability/2)
                actual = (z_p / accuracy) ** 2 * cv ** 2
                diff = abs(actual - expected)

                results.append(ValidationResult(
                    function_name="ACT_FULL_CREDIBILITY_STANDARD",
                    test_name=f"p={probability}_r={accuracy}",
                    expected=expected,
                    actual=actual,
                    source="k = (z_p/r)^2 * CV^2",
                    tolerance=tolerance,
                    passed=diff <= tolerance,
                    difference=diff
                ))

        return results


class ChainLadderValidator:
    """Validate chain ladder functions against reference values."""

    def __init__(self, fixtures: Dict):
        self.fixtures = fixtures

    def get_taylor_ashe_triangle(self) -> np.ndarray:
        """Get Taylor-Ashe triangle as numpy array."""
        data = self.fixtures.get("taylor_ashe_triangle", {}).get("data", [])
        triangle = []
        for row in data:
            triangle.append([0 if x is None else x for x in row])
        return np.array(triangle)

    def validate_factors(self) -> List[ValidationResult]:
        """Validate development factors."""
        results = []

        expected_factors = self.fixtures.get("development_factors", {}).get("expected", [])
        tolerance = self.fixtures.get("development_factors", {}).get("tolerance", 0.001)

        triangle = self.get_taylor_ashe_triangle()
        n = len(triangle)
        calculated_factors = []

        for j in range(n - 1):
            sum_current = 0
            sum_next = 0
            for i in range(n - j - 1):
                if triangle[i, j] > 0 and triangle[i, j + 1] > 0:
                    sum_current += triangle[i, j]
                    sum_next += triangle[i, j + 1]
            if sum_current > 0:
                calculated_factors.append(sum_next / sum_current)
            else:
                calculated_factors.append(1.0)

        for i, (expected, actual) in enumerate(zip(expected_factors, calculated_factors)):
            diff = abs(actual - expected)
            passed = diff <= tolerance

            results.append(ValidationResult(
                function_name="ACT_CL_FACTORS",
                test_name=f"factor_{i+1}_to_{i+2}",
                expected=expected,
                actual=actual,
                source="E&V (2002) / Taylor-Ashe",
                tolerance=tolerance,
                passed=passed,
                difference=diff
            ))

        return results

    def validate_ibnr(self) -> List[ValidationResult]:
        """Validate IBNR calculations."""
        results = []

        expected_ibnr = self.fixtures.get("chain_ladder_ibnr", {})
        total = expected_ibnr.get("total", 0)
        tolerance = expected_ibnr.get("tolerance", 1000)

        # Calculate IBNR from triangle
        triangle = self.get_taylor_ashe_triangle()
        n = len(triangle)

        # Get factors
        factors = []
        for j in range(n - 1):
            sum_current = 0
            sum_next = 0
            for i in range(n - j - 1):
                if triangle[i, j] > 0 and triangle[i, j + 1] > 0:
                    sum_current += triangle[i, j]
                    sum_next += triangle[i, j + 1]
            factors.append(sum_next / sum_current if sum_current > 0 else 1.0)

        # Calculate CDF to ultimate
        cdf = [1.0] * n
        for j in range(n - 2, -1, -1):
            cdf[j] = cdf[j + 1] * factors[j]

        # Calculate ultimates and IBNR
        total_ibnr = 0
        for i in range(n):
            last_col = n - 1 - i
            latest = triangle[i, last_col]
            ultimate = latest * cdf[last_col]
            ibnr = ultimate - latest
            total_ibnr += ibnr

        diff = abs(total_ibnr - total)
        results.append(ValidationResult(
            function_name="ACT_CL_IBNR",
            test_name="total_ibnr",
            expected=total,
            actual=total_ibnr,
            source="E&V (2002)",
            tolerance=tolerance,
            passed=diff <= tolerance,
            difference=diff
        ))

        return results

    def validate_bootstrap(self) -> List[ValidationResult]:
        """Validate ODP Bootstrap results."""
        results = []

        bootstrap_data = self.fixtures.get("odp_bootstrap", {})
        non_constant = bootstrap_data.get("non_constant_scale", {})
        our_impl = bootstrap_data.get("our_implementation", {})

        ev_se = non_constant.get("total_se", 0)
        our_se = our_impl.get("total_se", 0)
        match_pct = our_impl.get("match_vs_ev_non_constant", 0)

        results.append(ValidationResult(
            function_name="ACT_CL_BOOTSTRAP",
            test_name="total_se_vs_ev_non_constant",
            expected=ev_se,
            actual=our_se,
            source="E&V (2002) Table 3",
            tolerance=ev_se * 0.10,
            passed=match_pct >= 0.90,
            difference=abs(our_se - ev_se),
            notes=f"Match: {match_pct*100:.1f}% of E&V non-constant scale"
        ))

        return results


class MarkdownReporter:
    """Generate markdown validation reports."""

    def __init__(self, output_dir: str = None):
        if output_dir is None:
            output_dir = Path(__file__).parent / "reports"
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)

    def generate_report(
        self,
        distribution_results: Dict[str, List[ValidationResult]],
        chain_ladder_results: List[ValidationResult],
        copula_results: List[ValidationResult] = None,
        exposure_results: List[ValidationResult] = None,
        credibility_results: List[ValidationResult] = None
    ) -> str:
        """Generate full validation report."""

        lines = [
            "# Actuarial Add-In Validation Report",
            "",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "## Summary",
            "",
        ]

        # Calculate totals
        total_tests = 0
        passed_tests = 0

        all_results = []
        for results in distribution_results.values():
            all_results.extend(results)
        all_results.extend(chain_ladder_results)
        if copula_results:
            all_results.extend(copula_results)
        if exposure_results:
            all_results.extend(exposure_results)
        if credibility_results:
            all_results.extend(credibility_results)

        for r in all_results:
            if r.actual is not None:
                total_tests += 1
                if r.passed:
                    passed_tests += 1

        pass_rate = (passed_tests / total_tests * 100) if total_tests > 0 else 0

        lines.extend([
            f"- **Total Tests**: {total_tests}",
            f"- **Passed**: {passed_tests} ({pass_rate:.1f}%)",
            f"- **Failed**: {total_tests - passed_tests}",
            "",
        ])

        # Distribution results
        lines.extend([
            "## Distributions",
            "",
        ])

        for dist_name, results in sorted(distribution_results.items()):
            if not results:
                continue
            lines.append(f"### {dist_name.replace('_', ' ').title()}")
            lines.append("")
            lines.append("| Test | Expected | Actual | Diff | Source | Status |")
            lines.append("|------|----------|--------|------|--------|--------|")

            for r in results:
                if r.actual is None:
                    continue
                actual_str = f"{r.actual:.6f}"
                expected_str = f"{r.expected:.6f}" if r.expected is not None else "N/A"
                diff_str = f"{r.difference:.2e}" if r.difference is not None else "N/A"
                lines.append(
                    f"| {r.test_name} | {expected_str} | {actual_str} | "
                    f"{diff_str} | {r.source} | {r.status_icon} |"
                )
            lines.append("")

        # Chain Ladder results
        lines.extend([
            "## Chain Ladder",
            "",
            "### Development Factors",
            "",
            "| Test | Expected | Actual | Diff | Source | Status |",
            "|------|----------|--------|------|--------|--------|",
        ])

        for r in chain_ladder_results:
            if "factor" in r.test_name and r.actual is not None:
                lines.append(
                    f"| {r.test_name} | {r.expected:.4f} | {r.actual:.4f} | "
                    f"{r.difference:.4f} | {r.source} | {r.status_icon} |"
                )

        lines.append("")
        lines.extend([
            "### IBNR and Bootstrap",
            "",
            "| Test | Expected | Actual | Diff | Source | Status | Notes |",
            "|------|----------|--------|------|--------|--------|-------|",
        ])

        for r in chain_ladder_results:
            if "factor" not in r.test_name and r.actual is not None:
                expected_str = f"{r.expected:,.0f}"
                actual_str = f"{r.actual:,.0f}"
                diff_str = f"{r.difference:,.0f}" if r.difference else "N/A"
                lines.append(
                    f"| {r.test_name} | {expected_str} | {actual_str} | "
                    f"{diff_str} | {r.source} | {r.status_icon} | {r.notes} |"
                )

        lines.append("")

        # Copula results
        if copula_results:
            lines.extend([
                "## Copulas",
                "",
                "| Test | Expected | Actual | Diff | Source | Status |",
                "|------|----------|--------|------|--------|--------|",
            ])

            for r in copula_results:
                if r.actual is None:
                    continue
                lines.append(
                    f"| {r.function_name} {r.test_name} | {r.expected:.4f} | {r.actual:.4f} | "
                    f"{r.difference:.4f} | {r.source} | {r.status_icon} |"
                )
            lines.append("")

        # Exposure Curve results
        if exposure_results:
            lines.extend([
                "## Exposure Curves",
                "",
                "| Function | Test | Expected | Actual | Diff | Source | Status |",
                "|----------|------|----------|--------|------|--------|--------|",
            ])

            for r in exposure_results:
                if r.actual is None:
                    continue
                lines.append(
                    f"| {r.function_name} | {r.test_name} | {r.expected:.4f} | {r.actual:.4f} | "
                    f"{r.difference:.4f} | {r.source} | {r.status_icon} |"
                )
            lines.append("")

        # Credibility results
        if credibility_results:
            lines.extend([
                "## Credibility",
                "",
                "| Function | Test | Expected | Actual | Diff | Source | Status |",
                "|----------|------|----------|--------|------|--------|--------|",
            ])

            for r in credibility_results:
                if r.actual is None:
                    continue
                expected_str = f"{r.expected:.4f}" if r.expected < 100 else f"{r.expected:.0f}"
                actual_str = f"{r.actual:.4f}" if r.actual < 100 else f"{r.actual:.0f}"
                diff_str = f"{r.difference:.4f}" if r.difference and r.difference < 100 else f"{r.difference:.0f}" if r.difference else "N/A"
                lines.append(
                    f"| {r.function_name} | {r.test_name} | {expected_str} | {actual_str} | "
                    f"{diff_str} | {r.source} | {r.status_icon} |"
                )
            lines.append("")

        # References
        lines.extend([
            "## References",
            "",
            "1. England, P.D. and Verrall, R.J. (2002). Stochastic claims reserving in general insurance. British Actuarial Journal 8(3): 443-518.",
            "2. Mack, T. (1993). Distribution-free calculation of the standard error of chain ladder reserve estimates. ASTIN Bulletin 23(2): 213-225.",
            "3. Bernegger, S. (1997). The Swiss Re Exposure Curves and the MBBEFD Distribution Class. ASTIN Bulletin 27(1): 99-111.",
            "4. McNeil, A.J., Frey, R., and Embrechts, P. (2015). Quantitative Risk Management. Princeton University Press.",
            "5. scipy.stats documentation: https://docs.scipy.org/doc/scipy/reference/stats.html",
            "6. chainladder-python documentation: https://chainladder-python.readthedocs.io/",
            "",
        ])

        report = "\n".join(lines)

        # Save report
        report_path = self.output_dir / "VALIDATION_REPORT.md"
        with open(report_path, 'w') as f:
            f.write(report)

        return report


def run_validation():
    """Run full validation and generate report."""
    loader = FixtureLoader()

    # Load fixtures
    dist_fixtures = loader.load_distributions()
    cl_fixtures = loader.load_chain_ladder()
    copula_fixtures = loader.load_copulas()
    exposure_fixtures = loader.load_exposure_curves()
    credibility_fixtures = loader.load_credibility()

    # Validate distributions
    scipy_validator = ScipyValidator(dist_fixtures)
    dist_results = scipy_validator.validate_all_distributions()

    # Validate chain ladder
    cl_validator = ChainLadderValidator(cl_fixtures)
    cl_results = []
    cl_results.extend(cl_validator.validate_factors())
    cl_results.extend(cl_validator.validate_ibnr())
    cl_results.extend(cl_validator.validate_bootstrap())

    # Validate copulas
    copula_validator = CopulaValidator(copula_fixtures)
    copula_results = copula_validator.validate_all()

    # Validate exposure curves
    exposure_validator = ExposureCurveValidator(exposure_fixtures)
    exposure_results = exposure_validator.validate_all()

    # Validate credibility
    credibility_validator = CredibilityValidator(credibility_fixtures)
    credibility_results = credibility_validator.validate_all()

    # Generate report
    reporter = MarkdownReporter()
    report = reporter.generate_report(
        dist_results, cl_results,
        copula_results, exposure_results, credibility_results
    )

    print("Validation complete!")
    print(f"Report saved to: {reporter.output_dir / 'VALIDATION_REPORT.md'}")

    # Print summary
    total = 0
    passed = 0
    for results in dist_results.values():
        for r in results:
            if r.actual is not None:
                total += 1
                if r.passed:
                    passed += 1
    for r in cl_results + copula_results + exposure_results + credibility_results:
        if r.actual is not None:
            total += 1
            if r.passed:
                passed += 1

    print(f"\nSummary: {passed}/{total} tests passed ({passed/total*100:.1f}%)")

    return dist_results, cl_results, copula_results, exposure_results, credibility_results


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Actuarial Add-In Validation")
    parser.add_argument("--generate-report", action="store_true", help="Generate validation report")
    parser.add_argument("--validate-distributions", action="store_true", help="Validate distributions only")
    parser.add_argument("--validate-chain-ladder", action="store_true", help="Validate chain ladder only")

    args = parser.parse_args()

    if args.generate_report or not any(vars(args).values()):
        run_validation()
