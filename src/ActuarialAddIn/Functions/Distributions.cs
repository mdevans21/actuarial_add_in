using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Probability distribution functions commonly used in actuarial science.
/// Includes frequency distributions (Poisson, Negative Binomial) and severity distributions
/// (Lognormal, Gamma, Pareto, GPD, Weibull, Burr, Beta, Exponential).
/// </summary>
public static class Distributions
{
    #region Poisson Distribution

    [ExcelFunction(Description = "Poisson PMF: P(X=k) = λ^k * e^(-λ) / k!. Used for modeling claim counts.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_POISSON_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (non-negative integer)")] int k,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0). Returns NaN if invalid.")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (k < 0) return double.NaN;
        return Poisson.PMF(lambda, k);
    }

    [ExcelFunction(Description = "Poisson cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_POISSON_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (x < 0) return 0;
        return Poisson.CDF(lambda, x);
    }

    [ExcelFunction(Description = "Poisson inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_POISSON_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;

        // Binary search for inverse CDF
        int low = 0;
        int high = (int)Math.Max(1000, lambda * 10);
        while (low < high)
        {
            int mid = (low + high) / 2;
            double cdf = Poisson.CDF(lambda, mid);
            if (cdf < p)
                low = mid + 1;
            else
                high = mid;
        }
        return low;
    }

    #endregion

    #region Negative Binomial Distribution

    [ExcelFunction(Description = "Negative Binomial probability mass function (PDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_NEGBIN_PDF(
        [ExcelArgument(Description = "Number of failures (non-negative integer)")] int k,
        [ExcelArgument(Description = "Number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "Probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (k < 0) return double.NaN;
        return NegativeBinomial.PMF(r, p, k);
    }

    [ExcelFunction(Description = "Negative Binomial cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_NEGBIN_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "Probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (x < 0) return 0;
        return NegativeBinomial.CDF(r, p, x);
    }

    [ExcelFunction(Description = "Negative Binomial inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_NEGBIN_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double prob,
        [ExcelArgument(Description = "Number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "Probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (prob < 0 || prob > 1) return double.NaN;

        // Binary search for inverse since MathNet doesn't provide direct inverse
        int low = 0;
        int high = 1000000;
        while (low < high)
        {
            int mid = (low + high) / 2;
            double cdf = NegativeBinomial.CDF(r, p, mid);
            if (cdf < prob)
                low = mid + 1;
            else
                high = mid;
        }
        return low;
    }

    #endregion

    #region Zero-Truncated Poisson Distribution

    [ExcelFunction(Description = "Zero-Truncated Poisson PMF: P(X=k|X>0). Used when zero counts aren't observable.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTPOISSON_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (positive integer, k >= 1)")] int k,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (k < 1) return 0;

        // f(k) = P(k) / (1 - P(0)) = P(k) / (1 - e^(-λ))
        double p0 = Math.Exp(-lambda);
        double pk = Poisson.PMF(lambda, k);
        return pk / (1 - p0);
    }

    [ExcelFunction(Description = "Zero-Truncated Poisson CDF: P(X<=k|X>0)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTPOISSON_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (x < 1) return 0;

        // F_ZT(k) = (F(k) - F(0)) / (1 - F(0)) = (F(k) - P(0)) / (1 - P(0))
        double p0 = Math.Exp(-lambda);
        double cdfK = Poisson.CDF(lambda, x);
        return (cdfK - p0) / (1 - p0);
    }

    [ExcelFunction(Description = "Zero-Truncated Poisson inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTPOISSON_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        if (p == 0) return 1;

        // Binary search for inverse
        int low = 1;
        int high = (int)Math.Max(1000, lambda * 10);
        while (low < high)
        {
            int mid = (low + high) / 2;
            double cdf = ACT_DIST_ZTPOISSON_CDF(mid, lambda);
            if (cdf < p)
                low = mid + 1;
            else
                high = mid;
        }
        return low;
    }

    [ExcelFunction(Description = "Zero-Truncated Poisson mean: E[X|X>0] = λ / (1 - e^(-λ))", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTPOISSON_MEAN(
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        double p0 = Math.Exp(-lambda);
        return lambda / (1 - p0);
    }

    #endregion

    #region Zero-Truncated Negative Binomial Distribution

    [ExcelFunction(Description = "Zero-Truncated Negative Binomial PMF: P(X=k|X>0)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTNEGBIN_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (positive integer, k >= 1)")] int k,
        [ExcelArgument(Description = "r - number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (k < 1) return 0;

        // f(k) = NB(k) / (1 - NB(0)) where NB(0) = p^r
        double p0 = Math.Pow(p, r);
        double pk = NegativeBinomial.PMF(r, p, k);
        return pk / (1 - p0);
    }

    [ExcelFunction(Description = "Zero-Truncated Negative Binomial CDF: P(X<=k|X>0)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTNEGBIN_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "r - number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (x < 1) return 0;

        double p0 = Math.Pow(p, r);
        double cdfK = NegativeBinomial.CDF(r, p, x);
        return (cdfK - p0) / (1 - p0);
    }

    [ExcelFunction(Description = "Zero-Truncated Negative Binomial inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTNEGBIN_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double prob,
        [ExcelArgument(Description = "r - number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (prob < 0 || prob > 1) return double.NaN;
        if (prob == 0) return 1;

        // Binary search for inverse
        int low = 1;
        int high = 1000000;
        while (low < high)
        {
            int mid = (low + high) / 2;
            double cdf = ACT_DIST_ZTNEGBIN_CDF(mid, r, p);
            if (cdf < prob)
                low = mid + 1;
            else
                high = mid;
        }
        return low;
    }

    [ExcelFunction(Description = "Zero-Truncated Negative Binomial mean: E[X|X>0] = r(1-p) / (p(1 - p^r))", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTNEGBIN_MEAN(
        [ExcelArgument(Description = "r - number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        double p0 = Math.Pow(p, r);
        // Mean of NegBin = r(1-p)/p, so ZT mean = r(1-p)/p / (1-p^r)
        return r * (1 - p) / (p * (1 - p0));
    }

    #endregion

    #region Zero-Truncated Binomial Distribution

    [ExcelFunction(Description = "Zero-Truncated Binomial PMF: P(X=k|X>0)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTBINOM_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (positive integer, 1 <= k <= n)")] int k,
        [ExcelArgument(Description = "n - number of trials (positive integer)")] int n,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (n <= 0 || p < 0 || p > 1) return double.NaN;
        if (k < 1 || k > n) return 0;

        // f(k) = B(k) / (1 - B(0)) where B(0) = (1-p)^n
        double p0 = Math.Pow(1 - p, n);
        double pk = Binomial.PMF(p, n, k);
        return pk / (1 - p0);
    }

    [ExcelFunction(Description = "Zero-Truncated Binomial CDF: P(X<=k|X>0)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTBINOM_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "n - number of trials (positive integer)")] int n,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (n <= 0 || p < 0 || p > 1) return double.NaN;
        if (x < 1) return 0;
        if (x >= n) return 1;

        double p0 = Math.Pow(1 - p, n);
        double cdfK = Binomial.CDF(p, n, x);
        return (cdfK - p0) / (1 - p0);
    }

    [ExcelFunction(Description = "Zero-Truncated Binomial inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTBINOM_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double prob,
        [ExcelArgument(Description = "n - number of trials (positive integer)")] int n,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (n <= 0 || p < 0 || p > 1) return double.NaN;
        if (prob < 0 || prob > 1) return double.NaN;
        if (prob == 0) return 1;
        if (prob == 1) return n;

        // Binary search for inverse
        int low = 1;
        int high = n;
        while (low < high)
        {
            int mid = (low + high) / 2;
            double cdf = ACT_DIST_ZTBINOM_CDF(mid, n, p);
            if (cdf < prob)
                low = mid + 1;
            else
                high = mid;
        }
        return low;
    }

    [ExcelFunction(Description = "Zero-Truncated Binomial mean: E[X|X>0] = np / (1 - (1-p)^n)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTBINOM_MEAN(
        [ExcelArgument(Description = "n - number of trials (positive integer)")] int n,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (n <= 0 || p < 0 || p > 1) return double.NaN;
        double p0 = Math.Pow(1 - p, n);
        return n * p / (1 - p0);
    }

    #endregion

    #region Zero-Truncated Geometric Distribution

    [ExcelFunction(Description = "Zero-Truncated Geometric PMF: P(X=k|X>0). Note: Geometric counts failures before first success.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTGEOM_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (positive integer, k >= 1)")] int k,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (p <= 0 || p > 1) return double.NaN;
        if (k < 1) return 0;

        // Geometric PMF: P(k) = p(1-p)^k for k=0,1,2,... (counting failures)
        // P(0) = p, so ZT PMF = p(1-p)^k / (1-p) = p(1-p)^(k-1)
        // This simplifies to geometric starting at k=1
        return p * Math.Pow(1 - p, k - 1);
    }

    [ExcelFunction(Description = "Zero-Truncated Geometric CDF: P(X<=k|X>0)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTGEOM_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (p <= 0 || p > 1) return double.NaN;
        if (x < 1) return 0;

        int k = (int)Math.Floor(x);
        // CDF = sum_{j=1}^{k} p(1-p)^(j-1) = 1 - (1-p)^k
        return 1 - Math.Pow(1 - p, k);
    }

    [ExcelFunction(Description = "Zero-Truncated Geometric inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTGEOM_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double prob,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (p <= 0 || p > 1) return double.NaN;
        if (prob < 0 || prob > 1) return double.NaN;
        if (prob == 0) return 1;
        if (prob == 1) return double.PositiveInfinity;

        // CDF = 1 - (1-p)^k, so k = ceil(log(1-prob) / log(1-p))
        return Math.Ceiling(Math.Log(1 - prob) / Math.Log(1 - p));
    }

    [ExcelFunction(Description = "Zero-Truncated Geometric mean: E[X|X>0] = 1/p", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZTGEOM_MEAN(
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p)
    {
        if (p <= 0 || p > 1) return double.NaN;
        return 1.0 / p;
    }

    #endregion

    #region Zero-Modified (Zero-Inflated) Poisson Distribution

    /// <summary>
    /// Zero-Modified Poisson (ZIP) distribution.
    /// Adds extra probability mass at zero: P(X=0) = p0 + (1-p0)*e^(-lambda)
    /// For k >= 1: P(X=k) = (1-p0) * Poisson(k; lambda)
    /// </summary>

    [ExcelFunction(Description = "Zero-Modified Poisson PMF: P(X=k) with extra mass at zero", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMPOISSON_PDF(
        [ExcelArgument(Description = "k - number of events (>= 0)")] int k,
        [ExcelArgument(Description = "lambda - Poisson rate parameter (> 0)")] double lambda,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (lambda <= 0 || p0 < 0 || p0 > 1) return double.NaN;
        if (k < 0) return 0;
        if (k == 0) return p0 + (1 - p0) * Math.Exp(-lambda);
        return (1 - p0) * Poisson.PMF(lambda, k);
    }

    [ExcelFunction(Description = "Zero-Modified Poisson CDF: P(X<=k)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMPOISSON_CDF(
        [ExcelArgument(Description = "k - number of events")] int k,
        [ExcelArgument(Description = "lambda - Poisson rate parameter (> 0)")] double lambda,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (lambda <= 0 || p0 < 0 || p0 > 1) return double.NaN;
        if (k < 0) return 0;
        // CDF = p0 + (1-p0) * Poisson_CDF(k)
        return p0 + (1 - p0) * Poisson.CDF(lambda, k);
    }

    [ExcelFunction(Description = "Zero-Modified Poisson inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static int ACT_DIST_ZMPOISSON_INV(
        [ExcelArgument(Description = "p - probability (0 to 1)")] double p,
        [ExcelArgument(Description = "lambda - Poisson rate parameter (> 0)")] double lambda,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (lambda <= 0 || p0 < 0 || p0 > 1 || p < 0 || p > 1) return -1;

        // If p <= P(X=0), return 0
        double prob0 = p0 + (1 - p0) * Math.Exp(-lambda);
        if (p <= prob0) return 0;

        // Otherwise find smallest k where CDF >= p
        for (int k = 1; k < 1000; k++)
        {
            if (ACT_DIST_ZMPOISSON_CDF(k, lambda, p0) >= p) return k;
        }
        return 999;
    }

    [ExcelFunction(Description = "Zero-Modified Poisson mean: E[X] = (1-p0)*lambda", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMPOISSON_MEAN(
        [ExcelArgument(Description = "lambda - Poisson rate parameter (> 0)")] double lambda,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (lambda <= 0 || p0 < 0 || p0 > 1) return double.NaN;
        return (1 - p0) * lambda;
    }

    [ExcelFunction(Description = "Zero-Modified Poisson variance", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMPOISSON_VAR(
        [ExcelArgument(Description = "lambda - Poisson rate parameter (> 0)")] double lambda,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (lambda <= 0 || p0 < 0 || p0 > 1) return double.NaN;
        // Var(X) = (1-p0) * lambda * (1 + p0*lambda)
        return (1 - p0) * lambda * (1 + p0 * lambda);
    }

    #endregion

    #region Zero-Modified (Zero-Inflated) Negative Binomial Distribution

    /// <summary>
    /// Zero-Modified Negative Binomial (ZINB) distribution.
    /// Adds extra probability mass at zero: P(X=0) = p0 + (1-p0)*p^r
    /// For k >= 1: P(X=k) = (1-p0) * NegBin(k; r, p)
    /// </summary>

    [ExcelFunction(Description = "Zero-Modified Negative Binomial PMF: P(X=k) with extra mass at zero", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMNEGBIN_PDF(
        [ExcelArgument(Description = "k - number of failures (>= 0)")] int k,
        [ExcelArgument(Description = "r - number of successes (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (r <= 0 || p <= 0 || p > 1 || p0 < 0 || p0 > 1) return double.NaN;
        if (k < 0) return 0;
        if (k == 0) return p0 + (1 - p0) * Math.Pow(p, r);
        return (1 - p0) * NegativeBinomial.PMF(r, p, k);
    }

    [ExcelFunction(Description = "Zero-Modified Negative Binomial CDF: P(X<=k)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMNEGBIN_CDF(
        [ExcelArgument(Description = "k - number of failures")] int k,
        [ExcelArgument(Description = "r - number of successes (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (r <= 0 || p <= 0 || p > 1 || p0 < 0 || p0 > 1) return double.NaN;
        if (k < 0) return 0;
        // CDF = p0 + (1-p0) * NegBin_CDF(k)
        return p0 + (1 - p0) * NegativeBinomial.CDF(r, p, k);
    }

    [ExcelFunction(Description = "Zero-Modified Negative Binomial inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static int ACT_DIST_ZMNEGBIN_INV(
        [ExcelArgument(Description = "prob - probability (0 to 1)")] double prob,
        [ExcelArgument(Description = "r - number of successes (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (r <= 0 || p <= 0 || p > 1 || p0 < 0 || p0 > 1 || prob < 0 || prob > 1) return -1;

        // If prob <= P(X=0), return 0
        double prob0 = p0 + (1 - p0) * Math.Pow(p, r);
        if (prob <= prob0) return 0;

        // Otherwise find smallest k where CDF >= prob
        for (int k = 1; k < 10000; k++)
        {
            if (ACT_DIST_ZMNEGBIN_CDF(k, r, p, p0) >= prob) return k;
        }
        return 9999;
    }

    [ExcelFunction(Description = "Zero-Modified Negative Binomial mean: E[X] = (1-p0)*r*(1-p)/p", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMNEGBIN_MEAN(
        [ExcelArgument(Description = "r - number of successes (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (r <= 0 || p <= 0 || p > 1 || p0 < 0 || p0 > 1) return double.NaN;
        // Mean of NegBin is r*(1-p)/p
        return (1 - p0) * r * (1 - p) / p;
    }

    [ExcelFunction(Description = "Zero-Modified Negative Binomial variance", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_ZMNEGBIN_VAR(
        [ExcelArgument(Description = "r - number of successes (> 0)")] double r,
        [ExcelArgument(Description = "p - probability of success (0 to 1)")] double p,
        [ExcelArgument(Description = "p0 - extra probability at zero (0 to 1)")] double p0)
    {
        if (r <= 0 || p <= 0 || p > 1 || p0 < 0 || p0 > 1) return double.NaN;
        // E[X] for base NegBin
        double mu = r * (1 - p) / p;
        // Var(X) for base NegBin
        double sigma2 = r * (1 - p) / (p * p);
        // Var for ZM: (1-p0) * (sigma2 + p0*mu^2)
        return (1 - p0) * (sigma2 + p0 * mu * mu);
    }

    #endregion

    #region Lognormal Distribution

    [ExcelFunction(Description = "Lognormal probability density function (PDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGNORM_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (> 0)")] double x,
        [ExcelArgument(Description = "Mu - mean of log(X)")] double mu,
        [ExcelArgument(Description = "Sigma - standard deviation of log(X) (> 0)")] double sigma)
    {
        if (sigma <= 0) return double.NaN;
        if (x < 0) return double.NaN;
        if (x == 0) return 0;
        return LogNormal.PDF(mu, sigma, x);
    }

    [ExcelFunction(Description = "Lognormal cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGNORM_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - mean of log(X)")] double mu,
        [ExcelArgument(Description = "Sigma - standard deviation of log(X) (> 0)")] double sigma)
    {
        if (sigma <= 0) return double.NaN;
        if (x <= 0) return 0;
        return LogNormal.CDF(mu, sigma, x);
    }

    [ExcelFunction(Description = "Lognormal inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGNORM_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Mu - mean of log(X)")] double mu,
        [ExcelArgument(Description = "Sigma - standard deviation of log(X) (> 0)")] double sigma)
    {
        if (sigma <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        return LogNormal.InvCDF(mu, sigma, p);
    }

    #endregion

    #region Gamma Distribution

    [ExcelFunction(Description = "Gamma probability density function (PDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GAMMA_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - rate parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (x < 0) return double.NaN;
        return Gamma.PDF(alpha, beta, x);
    }

    [ExcelFunction(Description = "Gamma cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GAMMA_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - rate parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (x < 0) return 0;
        return Gamma.CDF(alpha, beta, x);
    }

    [ExcelFunction(Description = "Gamma inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GAMMA_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - rate parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        return Gamma.InvCDF(alpha, beta, p);
    }

    #endregion

    #region Pareto Distribution

    [ExcelFunction(Description = "Pareto Type I PDF: f(x) = α*xm^α / x^(α+1) for x >= xm. Heavy-tailed severity distribution.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= xm)")] double x,
        [ExcelArgument(Description = "Alpha - shape/tail parameter (> 0). Lower alpha = heavier tail. Mean exists only if alpha > 1.")] double alpha,
        [ExcelArgument(Description = "Xm - scale/minimum value (> 0)")] double xm)
    {
        if (alpha <= 0 || xm <= 0) return double.NaN;
        if (x < xm) return 0;
        return Pareto.PDF(xm, alpha, x);
    }

    [ExcelFunction(Description = "Pareto cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Xm - scale/minimum value (> 0)")] double xm)
    {
        if (alpha <= 0 || xm <= 0) return double.NaN;
        if (x < xm) return 0;
        return Pareto.CDF(xm, alpha, x);
    }

    [ExcelFunction(Description = "Pareto Type I inverse CDF (quantile function): x = xm / (1-p)^(1/α)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Alpha - shape/tail parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Xm - scale/minimum value (> 0)")] double xm)
    {
        if (alpha <= 0 || xm <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        return Pareto.InvCDF(xm, alpha, p);
    }

    #endregion

    #region Generalized Pareto Distribution (GPD)

    [ExcelFunction(Description = "GPD PDF: Used in Extreme Value Theory for modeling exceedances over thresholds. Shape xi controls tail: xi>0 heavy tail, xi=0 exponential, xi<0 bounded.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GPD_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0 if xi >= 0, or 0 <= x <= -sigma/xi if xi < 0)")] double x,
        [ExcelArgument(Description = "Xi - shape parameter. xi > 0: Pareto-like heavy tail; xi = 0: exponential; xi < 0: bounded upper tail")] double xi,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma)
    {
        if (sigma <= 0) return double.NaN;
        if (x < 0) return 0;

        // Handle xi = 0 (exponential case)
        if (Math.Abs(xi) < 1e-10)
        {
            return Math.Exp(-x / sigma) / sigma;
        }

        // Check upper bound for xi < 0
        if (xi < 0 && x > -sigma / xi)
        {
            return 0;
        }

        double t = 1 + xi * x / sigma;
        if (t <= 0) return 0;

        return Math.Pow(t, -1 / xi - 1) / sigma;
    }

    [ExcelFunction(Description = "GPD CDF: F(x) = 1 - (1 + xi*x/sigma)^(-1/xi). Essential for peaks-over-threshold analysis.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GPD_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Xi - shape parameter")] double xi,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma)
    {
        if (sigma <= 0) return double.NaN;
        if (x < 0) return 0;

        // Handle xi = 0 (exponential case)
        if (Math.Abs(xi) < 1e-10)
        {
            return 1 - Math.Exp(-x / sigma);
        }

        // Check upper bound for xi < 0
        if (xi < 0 && x >= -sigma / xi)
        {
            return 1;
        }

        double t = 1 + xi * x / sigma;
        if (t <= 0) return xi < 0 ? 1 : 0;

        return 1 - Math.Pow(t, -1 / xi);
    }

    [ExcelFunction(Description = "GPD inverse CDF: x = sigma * ((1-p)^(-xi) - 1) / xi. Used for VaR calculations.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_GPD_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Xi - shape parameter")] double xi,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma)
    {
        if (sigma <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        if (p == 0) return 0;
        if (p == 1) return xi >= 0 ? double.PositiveInfinity : -sigma / xi;

        // Handle xi = 0 (exponential case)
        if (Math.Abs(xi) < 1e-10)
        {
            return -sigma * Math.Log(1 - p);
        }

        return sigma * (Math.Pow(1 - p, -xi) - 1) / xi;
    }

    #endregion

    #region Weibull Distribution

    [ExcelFunction(Description = "Weibull PDF: f(x) = (k/λ)*(x/λ)^(k-1)*exp(-(x/λ)^k). Versatile severity distribution: k<1 decreasing hazard, k=1 exponential, k>1 increasing hazard.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_WEIBULL_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "k - shape parameter (> 0). Controls hazard rate shape.")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (k <= 0 || lambda <= 0) return double.NaN;
        if (x < 0) return double.NaN;
        if (x == 0) return k < 1 ? double.PositiveInfinity : (k == 1 ? 1 / lambda : 0);
        return Weibull.PDF(k, lambda, x);
    }

    [ExcelFunction(Description = "Weibull CDF: F(x) = 1 - exp(-(x/λ)^k)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_WEIBULL_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "k - shape parameter (> 0)")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (k <= 0 || lambda <= 0) return double.NaN;
        if (x < 0) return 0;
        return Weibull.CDF(k, lambda, x);
    }

    [ExcelFunction(Description = "Weibull inverse CDF: x = λ * (-ln(1-p))^(1/k)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_WEIBULL_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "k - shape parameter (> 0)")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (k <= 0 || lambda <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        if (p == 0) return 0;
        if (p == 1) return double.PositiveInfinity;
        // Weibull inverse CDF: x = λ * (-ln(1-p))^(1/k)
        return lambda * Math.Pow(-Math.Log(1 - p), 1.0 / k);
    }

    #endregion

    #region Beta Distribution

    [ExcelFunction(Description = "Beta PDF on [0,1]: f(x) = x^(α-1)*(1-x)^(β-1)/B(α,β). Used for loss ratios, proportions, and credibility weights.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BETA_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (0 to 1)")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (x < 0 || x > 1) return double.NaN;
        return Beta.PDF(alpha, beta, x);
    }

    [ExcelFunction(Description = "Beta CDF: F(x) = I_x(α,β) where I is the regularized incomplete beta function", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BETA_CDF(
        [ExcelArgument(Description = "Value at which to evaluate (0 to 1)")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (x <= 0) return 0;
        if (x >= 1) return 1;
        return Beta.CDF(alpha, beta, x);
    }

    [ExcelFunction(Description = "Beta inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BETA_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        return Beta.InvCDF(alpha, beta, p);
    }

    #endregion

    #region Exponential Distribution

    [ExcelFunction(Description = "Exponential PDF: f(x) = λ*exp(-λx). Simplest continuous severity distribution with constant hazard rate.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_EXP_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0). Mean = 1/λ")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (x < 0) return double.NaN;
        return Exponential.PDF(lambda, x);
    }

    [ExcelFunction(Description = "Exponential CDF: F(x) = 1 - exp(-λx)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_EXP_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (x < 0) return 0;
        return Exponential.CDF(lambda, x);
    }

    [ExcelFunction(Description = "Exponential inverse CDF: x = -ln(1-p)/λ", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_EXP_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        return Exponential.InvCDF(lambda, p);
    }

    #endregion

    #region Burr Type XII Distribution

    [ExcelFunction(Description = "Burr Type XII PDF: f(x) = ck/λ * (x/λ)^(c-1) / (1+(x/λ)^c)^(k+1). Very flexible heavy-tailed distribution, generalizes Pareto and Loglogistic.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BURR_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "c - first shape parameter (> 0)")] double c,
        [ExcelArgument(Description = "k - second shape parameter (> 0). Controls tail heaviness.")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (c <= 0 || k <= 0 || lambda <= 0) return double.NaN;
        if (x < 0) return double.NaN;
        if (x == 0) return c < 1 ? double.PositiveInfinity : (c == 1 ? k / lambda : 0);

        double y = x / lambda;
        double yc = Math.Pow(y, c);
        return c * k / lambda * Math.Pow(y, c - 1) / Math.Pow(1 + yc, k + 1);
    }

    [ExcelFunction(Description = "Burr Type XII CDF: F(x) = 1 - (1+(x/λ)^c)^(-k)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BURR_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "c - first shape parameter (> 0)")] double c,
        [ExcelArgument(Description = "k - second shape parameter (> 0)")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (c <= 0 || k <= 0 || lambda <= 0) return double.NaN;
        if (x < 0) return 0;
        if (x == 0) return 0;

        double y = x / lambda;
        double yc = Math.Pow(y, c);
        return 1 - Math.Pow(1 + yc, -k);
    }

    [ExcelFunction(Description = "Burr Type XII inverse CDF: x = λ * ((1-p)^(-1/k) - 1)^(1/c)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_BURR_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "c - first shape parameter (> 0)")] double c,
        [ExcelArgument(Description = "k - second shape parameter (> 0)")] double k,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (c <= 0 || k <= 0 || lambda <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        if (p == 0) return 0;
        if (p == 1) return double.PositiveInfinity;

        return lambda * Math.Pow(Math.Pow(1 - p, -1 / k) - 1, 1 / c);
    }

    #endregion

    #region Normal Distribution

    [ExcelFunction(Description = "Standard Normal PDF: φ(z) = exp(-z²/2) / √(2π)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_NORMAL_PDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mean (default 0)")] double mu = 0,
        [ExcelArgument(Description = "Standard deviation (> 0, default 1)")] double sigma = 1)
    {
        if (sigma <= 0) return double.NaN;
        return Normal.PDF(mu, sigma, x);
    }

    [ExcelFunction(Description = "Standard Normal CDF: Φ(z)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_NORMAL_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mean (default 0)")] double mu = 0,
        [ExcelArgument(Description = "Standard deviation (> 0, default 1)")] double sigma = 1)
    {
        if (sigma <= 0) return double.NaN;
        return Normal.CDF(mu, sigma, x);
    }

    [ExcelFunction(Description = "Standard Normal inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_NORMAL_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Mean (default 0)")] double mu = 0,
        [ExcelArgument(Description = "Standard deviation (> 0, default 1)")] double sigma = 1)
    {
        if (sigma <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        return Normal.InvCDF(mu, sigma, p);
    }

    #endregion

    #region Lomax Distribution (Pareto Type II)

    [ExcelFunction(Description = "Lomax (Pareto Type II) PDF: f(x) = (α/λ) * (1 + x/λ)^(-(α+1)). Shifted Pareto starting at 0.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOMAX_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "Alpha - shape/tail parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (alpha <= 0 || lambda <= 0) return double.NaN;
        if (x < 0) return 0;
        return (alpha / lambda) * Math.Pow(1 + x / lambda, -(alpha + 1));
    }

    [ExcelFunction(Description = "Lomax (Pareto Type II) CDF: F(x) = 1 - (1 + x/λ)^(-α)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOMAX_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Alpha - shape/tail parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (alpha <= 0 || lambda <= 0) return double.NaN;
        if (x < 0) return 0;
        return 1 - Math.Pow(1 + x / lambda, -alpha);
    }

    [ExcelFunction(Description = "Lomax (Pareto Type II) inverse CDF: x = λ * ((1-p)^(-1/α) - 1)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOMAX_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Alpha - shape/tail parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Lambda - scale parameter (> 0)")] double lambda)
    {
        if (alpha <= 0 || lambda <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        if (p == 0) return 0;
        if (p == 1) return double.PositiveInfinity;
        return lambda * (Math.Pow(1 - p, -1 / alpha) - 1);
    }

    #endregion

    #region Pareto III Distribution (Lomax with location)

    /// <summary>
    /// Pareto Type III distribution (Lomax with location parameter).
    /// Parameters: μ (location), σ (scale), γ (shape)
    /// CDF: F(x) = 1 - (1 + (x-μ)/σ)^(-γ) for x >= μ
    /// Special case of Burr XII with c=1.
    /// </summary>

    [ExcelFunction(Description = "Pareto III PDF: f(x) = (γ/σ) * (1 + (x-μ)/σ)^(-(γ+1)) for x >= μ", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO3_PDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - location parameter")] double mu,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma,
        [ExcelArgument(Description = "Gamma - shape parameter (> 0)")] double gamma)
    {
        if (sigma <= 0 || gamma <= 0) return double.NaN;
        if (x < mu) return 0;
        if (x == mu) return gamma / sigma;
        double z = (x - mu) / sigma;
        return (gamma / sigma) * Math.Pow(1 + z, -(gamma + 1));
    }

    [ExcelFunction(Description = "Pareto III CDF: F(x) = 1 - (1 + (x-μ)/σ)^(-γ) for x >= μ", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO3_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - location parameter")] double mu,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma,
        [ExcelArgument(Description = "Gamma - shape parameter (> 0)")] double gamma)
    {
        if (sigma <= 0 || gamma <= 0) return double.NaN;
        if (x <= mu) return 0;
        double z = (x - mu) / sigma;
        return 1 - Math.Pow(1 + z, -gamma);
    }

    [ExcelFunction(Description = "Pareto III inverse CDF: x = μ + σ * ((1-p)^(-1/γ) - 1)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO3_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Mu - location parameter")] double mu,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma,
        [ExcelArgument(Description = "Gamma - shape parameter (> 0)")] double gamma)
    {
        if (sigma <= 0 || gamma <= 0 || p < 0 || p > 1) return double.NaN;
        if (p == 0) return mu;
        if (p == 1) return double.PositiveInfinity;
        return mu + sigma * (Math.Pow(1 - p, -1 / gamma) - 1);
    }

    [ExcelFunction(Description = "Pareto III mean: E[X] = μ + σ/(γ-1) for γ > 1", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO3_MEAN(
        [ExcelArgument(Description = "Mu - location parameter")] double mu,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma,
        [ExcelArgument(Description = "Gamma - shape parameter (> 1 for finite mean)")] double gamma)
    {
        if (sigma <= 0 || gamma <= 1) return double.NaN;
        return mu + sigma / (gamma - 1);
    }

    #endregion

    #region Pareto IV Distribution (Generalized Pareto)

    /// <summary>
    /// Pareto Type IV distribution (Generalized Pareto with 4 parameters).
    /// Parameters: μ (location), σ (scale), γ (inequality/shape1), α (shape2)
    /// CDF: F(x) = 1 - (1 + ((x-μ)/σ)^(1/γ))^(-α) for x >= μ
    /// Encompasses Pareto I (γ=1, α=shape), II (μ=0, γ=1), and III (α=γ).
    /// </summary>

    [ExcelFunction(Description = "Pareto IV PDF: Generalized 4-parameter Pareto distribution", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO4_PDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - location parameter")] double mu,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma,
        [ExcelArgument(Description = "Gamma - inequality parameter (> 0)")] double gamma,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha)
    {
        if (sigma <= 0 || gamma <= 0 || alpha <= 0) return double.NaN;
        if (x < mu) return 0;
        if (x == mu) return (gamma == 1) ? alpha / sigma : 0;

        double z = (x - mu) / sigma;
        double zPowGammaInv = Math.Pow(z, 1 / gamma);
        double baseTerm = 1 + zPowGammaInv;

        // PDF = (α / (σ * γ)) * z^(1/γ - 1) * (1 + z^(1/γ))^(-(α+1))
        return (alpha / (sigma * gamma)) * Math.Pow(z, 1 / gamma - 1) * Math.Pow(baseTerm, -(alpha + 1));
    }

    [ExcelFunction(Description = "Pareto IV CDF: F(x) = 1 - (1 + ((x-μ)/σ)^(1/γ))^(-α)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO4_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - location parameter")] double mu,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma,
        [ExcelArgument(Description = "Gamma - inequality parameter (> 0)")] double gamma,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha)
    {
        if (sigma <= 0 || gamma <= 0 || alpha <= 0) return double.NaN;
        if (x <= mu) return 0;

        double z = (x - mu) / sigma;
        double zPowGammaInv = Math.Pow(z, 1 / gamma);
        return 1 - Math.Pow(1 + zPowGammaInv, -alpha);
    }

    [ExcelFunction(Description = "Pareto IV inverse CDF: x = μ + σ * ((1-p)^(-1/α) - 1)^γ", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_PARETO4_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Mu - location parameter")] double mu,
        [ExcelArgument(Description = "Sigma - scale parameter (> 0)")] double sigma,
        [ExcelArgument(Description = "Gamma - inequality parameter (> 0)")] double gamma,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha)
    {
        if (sigma <= 0 || gamma <= 0 || alpha <= 0 || p < 0 || p > 1) return double.NaN;
        if (p == 0) return mu;
        if (p == 1) return double.PositiveInfinity;

        double innerTerm = Math.Pow(1 - p, -1 / alpha) - 1;
        return mu + sigma * Math.Pow(innerTerm, gamma);
    }

    #endregion

    #region Inverse Gaussian (Wald) Distribution

    [ExcelFunction(Description = "Inverse Gaussian (Wald) PDF: f(x) = √(λ/(2πx³)) * exp(-λ(x-μ)²/(2μ²x)). Used in operational risk and as GLM family.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_INVGAUSS_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (> 0)")] double x,
        [ExcelArgument(Description = "Mu - mean parameter (> 0)")] double mu,
        [ExcelArgument(Description = "Lambda - shape parameter (> 0)")] double lambda)
    {
        if (x <= 0 || mu <= 0 || lambda <= 0) return double.NaN;
        double z = (x - mu) / mu;
        return Math.Sqrt(lambda / (2 * Math.PI * x * x * x))
               * Math.Exp(-lambda * z * z / (2 * x));
    }

    [ExcelFunction(Description = "Inverse Gaussian CDF: Uses Φ(√(λ/x)((x/μ)-1)) + exp(2λ/μ) * Φ(-√(λ/x)((x/μ)+1))", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_INVGAUSS_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - mean parameter (> 0)")] double mu,
        [ExcelArgument(Description = "Lambda - shape parameter (> 0)")] double lambda)
    {
        if (mu <= 0 || lambda <= 0) return double.NaN;
        if (x <= 0) return 0;

        double sqrtLambdaOverX = Math.Sqrt(lambda / x);
        double xOverMu = x / mu;

        double z1 = sqrtLambdaOverX * (xOverMu - 1);
        double z2 = -sqrtLambdaOverX * (xOverMu + 1);

        double phi1 = Normal.CDF(0, 1, z1);
        double phi2 = Normal.CDF(0, 1, z2);

        return phi1 + Math.Exp(2 * lambda / mu) * phi2;
    }

    [ExcelFunction(Description = "Inverse Gaussian inverse CDF (quantile function). Uses Newton-Raphson iteration.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_INVGAUSS_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Mu - mean parameter (> 0)")] double mu,
        [ExcelArgument(Description = "Lambda - shape parameter (> 0)")] double lambda)
    {
        if (mu <= 0 || lambda <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        if (p == 0) return 0;
        if (p == 1) return double.PositiveInfinity;

        // Newton-Raphson to find x such that CDF(x) = p
        // Initial guess: use mean
        double x = mu;
        const int maxIter = 100;
        const double tol = 1e-10;

        for (int i = 0; i < maxIter; i++)
        {
            double cdf = ACT_DIST_INVGAUSS_CDF(x, mu, lambda);
            double pdf = ACT_DIST_INVGAUSS_PDF(x, mu, lambda);

            if (pdf < 1e-15) break;

            double delta = (cdf - p) / pdf;
            x = x - delta;

            // Ensure x stays positive
            if (x <= 0) x = tol;

            if (Math.Abs(delta) < tol * x) break;
        }

        return x;
    }

    #endregion

    #region Loglogistic Distribution

    [ExcelFunction(Description = "Loglogistic PDF: f(x) = (β/α)(x/α)^(β-1) / (1 + (x/α)^β)². Popular for liability claims; heavier tail than lognormal, lighter than Pareto.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGLOGISTIC_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "Alpha - scale parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0)")] double beta)
    {
        if (x < 0 || alpha <= 0 || beta <= 0) return double.NaN;
        if (x == 0) return beta == 1 ? 1.0 / alpha : (beta < 1 ? double.PositiveInfinity : 0.0);

        double z = x / alpha;
        double zb = Math.Pow(z, beta);
        double denom = (1 + zb) * (1 + zb);
        return (beta / alpha) * Math.Pow(z, beta - 1) / denom;
    }

    [ExcelFunction(Description = "Loglogistic CDF: F(x) = (x/α)^β / (1 + (x/α)^β)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGLOGISTIC_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Alpha - scale parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (x < 0) return 0;
        if (x == 0) return 0;

        double zb = Math.Pow(x / alpha, beta);
        return zb / (1 + zb);
    }

    [ExcelFunction(Description = "Loglogistic inverse CDF: Q(p) = α * (p/(1-p))^(1/β)", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LOGLOGISTIC_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Alpha - scale parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - shape parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        if (p == 0) return 0;
        if (p == 1) return double.PositiveInfinity;

        return alpha * Math.Pow(p / (1 - p), 1.0 / beta);
    }

    #endregion

    #region Lognormal-Pareto (Composite) Distribution

    [ExcelFunction(Description = "Lognormal-Pareto composite PDF. Lognormal below threshold θ, Pareto tail above.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LNPARETO_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (> 0)")] double x,
        [ExcelArgument(Description = "Mu - mean of underlying normal")] double mu,
        [ExcelArgument(Description = "Sigma - std dev of underlying normal (> 0)")] double sigma,
        [ExcelArgument(Description = "Theta - threshold where distribution switches to Pareto (> 0)")] double theta)
    {
        if (sigma <= 0 || theta <= 0) return double.NaN;
        if (x <= 0) return 0;

        double zTheta = (Math.Log(theta) - mu) / sigma;
        double phiZTheta = Normal.PDF(0, 1, zTheta);
        double PhiZTheta = Normal.CDF(0, 1, zTheta);
        double alpha = phiZTheta / (sigma * (1 - PhiZTheta));

        if (x < theta)
        {
            double z = (Math.Log(x) - mu) / sigma;
            return Normal.PDF(0, 1, z) / (x * sigma);
        }
        else
        {
            return (1 - PhiZTheta) * alpha * Math.Pow(theta, alpha) / Math.Pow(x, alpha + 1);
        }
    }

    [ExcelFunction(Description = "Lognormal-Pareto composite CDF. Lognormal below threshold θ, Pareto tail above.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LNPARETO_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - mean of underlying normal")] double mu,
        [ExcelArgument(Description = "Sigma - std dev of underlying normal (> 0)")] double sigma,
        [ExcelArgument(Description = "Theta - threshold where distribution switches to Pareto (> 0)")] double theta)
    {
        if (sigma <= 0 || theta <= 0) return double.NaN;
        if (x <= 0) return 0;

        double zTheta = (Math.Log(theta) - mu) / sigma;
        double PhiZTheta = Normal.CDF(0, 1, zTheta);

        if (x < theta)
        {
            double z = (Math.Log(x) - mu) / sigma;
            return Normal.CDF(0, 1, z);
        }
        else
        {
            double phiZTheta = Normal.PDF(0, 1, zTheta);
            double alpha = phiZTheta / (sigma * (1 - PhiZTheta));
            return 1 - (1 - PhiZTheta) * Math.Pow(theta / x, alpha);
        }
    }

    [ExcelFunction(Description = "Lognormal-Pareto derived tail index α", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_LNPARETO_ALPHA(
        [ExcelArgument(Description = "Mu - mean of underlying normal")] double mu,
        [ExcelArgument(Description = "Sigma - std dev of underlying normal (> 0)")] double sigma,
        [ExcelArgument(Description = "Theta - threshold (> 0)")] double theta)
    {
        if (sigma <= 0 || theta <= 0) return double.NaN;
        double zTheta = (Math.Log(theta) - mu) / sigma;
        double phiZTheta = Normal.PDF(0, 1, zTheta);
        double PhiZTheta = Normal.CDF(0, 1, zTheta);
        return phiZTheta / (sigma * (1 - PhiZTheta));
    }

    #endregion

    #region Exponential-Pareto (Composite) Distribution

    [ExcelFunction(Description = "Exponential-Pareto composite PDF. Exponential below threshold θ, Pareto tail above.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_EXPPARETO_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "Rate - exponential rate parameter (> 0)")] double rate,
        [ExcelArgument(Description = "Theta - threshold where distribution switches to Pareto (> 0)")] double theta)
    {
        if (rate <= 0 || theta <= 0) return double.NaN;
        if (x < 0) return 0;

        double alpha = theta * rate;
        double r = 1 - Math.Exp(-rate * theta);

        if (x < theta)
        {
            return rate * Math.Exp(-rate * x);
        }
        else
        {
            return (1 - r) * alpha * Math.Pow(theta, alpha) / Math.Pow(x, alpha + 1);
        }
    }

    [ExcelFunction(Description = "Exponential-Pareto composite CDF. Exponential below threshold θ, Pareto tail above.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_EXPPARETO_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Rate - exponential rate parameter (> 0)")] double rate,
        [ExcelArgument(Description = "Theta - threshold where distribution switches to Pareto (> 0)")] double theta)
    {
        if (rate <= 0 || theta <= 0) return double.NaN;
        if (x < 0) return 0;

        double alpha = theta * rate;
        double r = 1 - Math.Exp(-rate * theta);

        if (x < theta)
        {
            return 1 - Math.Exp(-rate * x);
        }
        else
        {
            return 1 - (1 - r) * Math.Pow(theta / x, alpha);
        }
    }

    [ExcelFunction(Description = "Exponential-Pareto derived tail index α = θ * rate", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_EXPPARETO_ALPHA(
        [ExcelArgument(Description = "Rate - exponential rate parameter (> 0)")] double rate,
        [ExcelArgument(Description = "Theta - threshold (> 0)")] double theta)
    {
        if (rate <= 0 || theta <= 0) return double.NaN;
        return theta * rate;
    }

    #endregion

    #region Power-Pareto (Composite) Distribution

    [ExcelFunction(Description = "Power-Pareto composite PDF. Power law below threshold θ, Pareto tail above.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_POWPARETO_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "Alpha - Pareto tail parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - Power law exponent (> 0)")] double beta,
        [ExcelArgument(Description = "Theta - threshold where distribution switches to Pareto (> 0)")] double theta)
    {
        if (alpha <= 0 || beta <= 0 || theta <= 0) return double.NaN;
        if (x < 0) return 0;

        double r = alpha / (alpha + beta);

        if (x < theta)
        {
            return (r * beta / theta) * Math.Pow(x / theta, beta - 1);
        }
        else
        {
            return (1 - r) * alpha * Math.Pow(theta, alpha) / Math.Pow(x, alpha + 1);
        }
    }

    [ExcelFunction(Description = "Power-Pareto composite CDF. Power law below threshold θ, Pareto tail above.", Category = "Actuarial.Distributions")]
    public static double ACT_DIST_POWPARETO_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Alpha - Pareto tail parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - Power law exponent (> 0)")] double beta,
        [ExcelArgument(Description = "Theta - threshold where distribution switches to Pareto (> 0)")] double theta)
    {
        if (alpha <= 0 || beta <= 0 || theta <= 0) return double.NaN;
        if (x < 0) return 0;

        double r = alpha / (alpha + beta);

        if (x < theta)
        {
            return r * Math.Pow(x / theta, beta);
        }
        else
        {
            return 1 - (1 - r) * Math.Pow(theta / x, alpha);
        }
    }

    #endregion

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
