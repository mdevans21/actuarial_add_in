using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics;

namespace ActuarialAddIn.Functions;

public static partial class Distributions
{
    #region Limited Expected Value (LEV) Functions

    /// <summary>
    /// Limited Expected Value (LEV) is E[min(X, d)] - the expected value of X capped at limit d.
    /// LEV = ∫₀ᵈ S(x) dx where S(x) = 1 - F(x) is the survival function.
    /// Used for calculating expected layer losses in reinsurance pricing.
    /// </summary>

    [ExcelFunction(Description = "Exponential LEV: E[min(X,d)] = (1/λ)[1 - e^(-λd)]. Mean capped at limit d.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_EXP_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;
        // LEV = (1/λ)[1 - e^(-λd)]
        return (1 - Math.Exp(-lambda * limit)) / lambda;
    }

    [ExcelFunction(Description = "Pareto Type I LEV: E[min(X,d)] for Pareto distribution. Requires α > 1 for finite mean.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO_LEV(
        [ExcelArgument(Description = "Limit d (>= xm)")] double limit,
        [ExcelArgument(Description = "Alpha - shape parameter (> 1 for finite mean)")] double alpha,
        [ExcelArgument(Description = "Xm - scale/minimum value (> 0)")] double xm)
    {
        if (alpha <= 0 || xm <= 0 || limit < 0) return double.NaN;
        if (limit <= xm) return limit;
        if (alpha <= 1) return double.NaN; // Mean doesn't exist

        // LEV = xm * α/(α-1) * [1 - (xm/d)^(α-1)] + d * (xm/d)^α
        // Klugman Loss Models Appendix A.2 / Klugman 3.10.
        double term1 = alpha * xm / (alpha - 1);
        double term2 = alpha * xm * Math.Pow(xm / limit, alpha - 1) / (alpha - 1);
        double term3 = limit * Math.Pow(xm / limit, alpha);
        return term1 - term2 + term3;
    }

    [ExcelFunction(Description = "Lomax (Pareto Type II) LEV: E[min(X,d)] for Lomax distribution starting at 0. Requires α > 1.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOMAX_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "Alpha - shape parameter (> 1 for finite mean)")] double alpha,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (alpha <= 1 || lambda <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // LEV for Lomax: λ/(α-1) * [1 - (λ/(λ+d))^(α-1)]
        double ratio = lambda / (lambda + limit);
        return lambda / (alpha - 1) * (1 - Math.Pow(ratio, alpha - 1));
    }

    [ExcelFunction(Description = "GPD LEV: E[min(X,d)] for Generalized Pareto Distribution. Used in EVT.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GPD_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "Xi - shape parameter (< 1 for finite mean)")] double xi,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma)
    {
        if (sigma <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;
        if (xi >= 1) return double.NaN; // Mean doesn't exist

        // Handle xi = 0 (exponential case)
        if (Math.Abs(xi) < 1e-10)
        {
            return sigma * (1 - Math.Exp(-limit / sigma));
        }

        // Check upper bound for xi < 0
        double upperBound = xi < 0 ? -sigma / xi : double.PositiveInfinity;
        double d = Math.Min(limit, upperBound);

        // LEV = σ/(1-ξ) * [1 - (1 + ξd/σ)^(1-1/ξ)]
        double t = 1 + xi * d / sigma;
        if (t <= 0) t = 1e-10;

        return sigma / (1 - xi) * (1 - Math.Pow(t, 1 - 1 / xi));
    }

    [ExcelFunction(Description = "Gamma LEV: E[min(X,d)] using incomplete gamma function.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GAMMA_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - rate parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // Mean of Gamma = α/β
        // LEV = (α/β) * P(α+1, β*d) + d * [1 - P(α, β*d)]
        // where P is the regularized lower incomplete gamma function
        double x = beta * limit;
        double gammaLower = SpecialFunctions.GammaLowerRegularized(alpha, x);
        double gammaLowerPlus1 = SpecialFunctions.GammaLowerRegularized(alpha + 1, x);

        return (alpha / beta) * gammaLowerPlus1 + limit * (1 - gammaLower);
    }

    [ExcelFunction(Description = "Lognormal LEV: E[min(X,d)] using normal CDF.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGNORM_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "Mu - mean of log(X)")] double mu,
        [ExcelArgument(Description = "Sigma - std dev of log(X) (> 0)")] double sigma)
    {
        if (sigma <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // Mean of Lognormal = exp(μ + σ²/2)
        // LEV = exp(μ + σ²/2) * Φ((ln(d) - μ - σ²)/σ) + d * [1 - Φ((ln(d) - μ)/σ)]
        double lnD = Math.Log(limit);
        double z1 = (lnD - mu - sigma * sigma) / sigma;
        double z2 = (lnD - mu) / sigma;

        double meanX = Math.Exp(mu + sigma * sigma / 2);
        double phi1 = Normal.CDF(0, 1, z1);
        double phi2 = Normal.CDF(0, 1, z2);

        return meanX * phi1 + limit * (1 - phi2);
    }

    [ExcelFunction(Description = "Weibull LEV: E[min(X,d)] using incomplete gamma function.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_WEIBULL_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "k - shape parameter (> 0)")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (k <= 0 || lambda <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // Mean of Weibull = λ * Γ(1 + 1/k)
        // LEV = λ * Γ(1 + 1/k) * P(1 + 1/k, (d/λ)^k) + d * exp(-(d/λ)^k)
        // where P is the regularized lower incomplete gamma function
        double a = 1 + 1 / k;
        double x = Math.Pow(limit / lambda, k);

        double gammaComplete = SpecialFunctions.Gamma(a);
        double gammaLower = SpecialFunctions.GammaLowerRegularized(a, x);
        double survivalAtD = Math.Exp(-x);

        return lambda * gammaComplete * gammaLower + limit * survivalAtD;
    }

    [ExcelFunction(Description = "Beta LEV: E[min(X,d)] for Beta distribution on [0,1].", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BETA_LEV(
        [ExcelArgument(Description = "Limit d (0 to 1)")] double limit,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0)")] double betaParam)
    {
        if (alpha <= 0 || betaParam <= 0) return double.NaN;
        if (limit < 0) return double.NaN;
        if (limit == 0) return 0;
        if (limit >= 1) return alpha / (alpha + betaParam); // Full mean

        // Mean of Beta = α/(α+β)
        // LEV = (α/(α+β)) * I_d(α+1, β) + d * [1 - I_d(α, β)]
        // where I is the regularized incomplete beta function
        double cdfAtD = SpecialFunctions.BetaRegularized(alpha, betaParam, limit);
        double cdfPlus1AtD = SpecialFunctions.BetaRegularized(alpha + 1, betaParam, limit);

        return (alpha / (alpha + betaParam)) * cdfPlus1AtD + limit * (1 - cdfAtD);
    }

    [ExcelFunction(Description = "Burr Type XII LEV: E[min(X,d)] using incomplete beta function.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BURR_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "c - first shape parameter (> 0)")] double c,
        [ExcelArgument(Description = "k - second shape parameter (> 0). Need ck > 1 for finite mean.")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (c <= 0 || k <= 0 || lambda <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;
        if (c * k <= 1) return double.NaN; // Mean doesn't exist

        // For Burr XII, LEV uses incomplete beta
        // Mean = λ * k * B(k - 1/c, 1 + 1/c) = λ * Γ(k - 1/c) * Γ(1 + 1/c) / Γ(k)
        // LEV = Mean * I_u(1 + 1/c, k - 1/c) + d * S(d)
        // where u = (d/λ)^c / (1 + (d/λ)^c)
        double yc = Math.Pow(limit / lambda, c);
        double u = yc / (1 + yc);

        double a1 = 1 + 1 / c;
        double a2 = k - 1 / c;

        // Mean of Burr XII = λ * Γ(k - 1/c) * Γ(1 + 1/c) / Γ(k) = λ * k * B(k - 1/c, 1 + 1/c)
        double betaComplete = SpecialFunctions.Beta(a2, a1);
        double mean = lambda * k * betaComplete;

        // Regularized incomplete beta I_u(a1, a2)
        double betaIncomplete = SpecialFunctions.BetaRegularized(a1, a2, u);

        // Survival function at d
        double survivalAtD = Math.Pow(1 + yc, -k);

        return mean * betaIncomplete + limit * survivalAtD;
    }

    [ExcelFunction(Description = "Poisson LEV: E[min(X,d)] for discrete Poisson. Sum of survival probabilities.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_POISSON_LEV(
        [ExcelArgument(Description = "Limit d (>= 0)")] double limit,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // For discrete: LEV = Σ_{k=0}^{floor(d)-1} P(X > k) = Σ_{k=0}^{floor(d)-1} [1 - F(k)]
        // Or equivalently: LEV = Σ_{k=0}^{floor(d)} k*P(k) + d*P(X > floor(d))
        int n = (int)Math.Floor(limit);

        double lev = 0;
        for (int i = 0; i < n; i++)
        {
            lev += 1 - Poisson.CDF(lambda, i);
        }

        // Add fractional part if limit is not integer
        double frac = limit - n;
        if (frac > 0)
        {
            lev += frac * (1 - Poisson.CDF(lambda, n));
        }

        return lev;
    }

    [ExcelFunction(Description = "Negative Binomial LEV: E[min(X,d)] for discrete Negative Binomial.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_NEGBIN_LEV(
        [ExcelArgument(Description = "Limit d (>= 0)")] double limit,
        [ExcelArgument(Description = "r - number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // For discrete: LEV = Σ_{k=0}^{floor(d)-1} [1 - F(k)]
        int n = (int)Math.Floor(limit);

        double lev = 0;
        for (int i = 0; i < n; i++)
        {
            lev += 1 - NegativeBinomial.CDF(r, p, i);
        }

        // Add fractional part
        double frac = limit - n;
        if (frac > 0)
        {
            lev += frac * (1 - NegativeBinomial.CDF(r, p, n));
        }

        return lev;
    }

    [ExcelFunction(Description = "Inverse Gaussian LEV: E[min(X,d)] via numerical integration of survival function.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_INVGAUSS_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "Mu - mean parameter (> 0)")] double mu,
        [ExcelArgument(Description = "Lambda - shape parameter (> 0)")] double lambda)
    {
        if (mu <= 0 || lambda <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // LEV = ∫₀ᵈ S(x) dx = ∫₀ᵈ [1 - F(x)] dx
        // Use numerical integration (Simpson's rule)
        const int n = 1000;
        double h = limit / n;
        double sum = 0;

        // Simpson's rule: (h/3) * [f(0) + 4f(1) + 2f(2) + 4f(3) + ... + f(n)]
        for (int i = 0; i <= n; i++)
        {
            double x = i * h;
            double survival = 1 - ACT_DIST_INVGAUSS_CDF(x, mu, lambda);

            if (i == 0 || i == n)
                sum += survival;
            else if (i % 2 == 1)
                sum += 4 * survival;
            else
                sum += 2 * survival;
        }

        return h * sum / 3;
    }

    [ExcelFunction(Description = "Loglogistic LEV: E[min(X,d)] via numerical integration of survival function.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGLOGISTIC_LEV(
        [ExcelArgument(Description = "Limit d (> 0)")] double limit,
        [ExcelArgument(Description = "Alpha - scale parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0). Mean finite only if β > 1.")] double beta)
    {
        if (alpha <= 0 || beta <= 0 || limit < 0) return double.NaN;
        if (limit == 0) return 0;

        // LEV = ∫₀ᵈ S(x) dx = ∫₀ᵈ [1 - F(x)] dx
        // Use numerical integration (Simpson's rule)
        const int n = 1000;
        double h = limit / n;
        double sum = 0;

        // Simpson's rule: (h/3) * [f(0) + 4f(1) + 2f(2) + 4f(3) + ... + f(n)]
        for (int i = 0; i <= n; i++)
        {
            double x = i * h;
            double survival = 1 - ACT_DIST_LOGLOGISTIC_CDF(x, alpha, beta);

            if (i == 0 || i == n)
                sum += survival;
            else if (i % 2 == 1)
                sum += 4 * survival;
            else
                sum += 2 * survival;
        }

        return h * sum / 3;
    }

    #endregion
}
