using ExcelDna.Integration;
using MathNet.Numerics.Distributions;

namespace ActuarialAddIn.Functions;

public static partial class Distributions
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

}
