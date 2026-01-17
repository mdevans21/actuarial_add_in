using ExcelDna.Integration;
using MathNet.Numerics.Distributions;

namespace ActuarialAddIn.Functions;

public static class Distributions
{
    #region Poisson Distribution

    [ExcelFunction(Description = "Poisson probability mass function (PDF)", Category = "Actuarial.Distributions")]
    public static double ACT_POISSON_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (non-negative integer)")] int k,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (k < 0) return double.NaN;
        return Poisson.PMF(lambda, k);
    }

    [ExcelFunction(Description = "Poisson cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_POISSON_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Lambda - rate parameter (> 0)")] double lambda)
    {
        if (lambda <= 0) return double.NaN;
        if (x < 0) return 0;
        return Poisson.CDF(lambda, x);
    }

    [ExcelFunction(Description = "Poisson inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_POISSON_INV(
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
    public static double ACT_NEGBIN_PDF(
        [ExcelArgument(Description = "Number of failures (non-negative integer)")] int k,
        [ExcelArgument(Description = "Number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "Probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (k < 0) return double.NaN;
        return NegativeBinomial.PMF(r, p, k);
    }

    [ExcelFunction(Description = "Negative Binomial cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_NEGBIN_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Number of successes required (> 0)")] double r,
        [ExcelArgument(Description = "Probability of success (0 to 1)")] double p)
    {
        if (r <= 0 || p <= 0 || p > 1) return double.NaN;
        if (x < 0) return 0;
        return NegativeBinomial.CDF(r, p, x);
    }

    [ExcelFunction(Description = "Negative Binomial inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_NEGBIN_INV(
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

    #region Lognormal Distribution

    [ExcelFunction(Description = "Lognormal probability density function (PDF)", Category = "Actuarial.Distributions")]
    public static double ACT_LOGNORM_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (> 0)")] double x,
        [ExcelArgument(Description = "Mu - mean of log(X)")] double mu,
        [ExcelArgument(Description = "Sigma - standard deviation of log(X) (> 0)")] double sigma)
    {
        if (sigma <= 0 || x <= 0) return double.NaN;
        return LogNormal.PDF(mu, sigma, x);
    }

    [ExcelFunction(Description = "Lognormal cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_LOGNORM_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Mu - mean of log(X)")] double mu,
        [ExcelArgument(Description = "Sigma - standard deviation of log(X) (> 0)")] double sigma)
    {
        if (sigma <= 0) return double.NaN;
        if (x <= 0) return 0;
        return LogNormal.CDF(mu, sigma, x);
    }

    [ExcelFunction(Description = "Lognormal inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_LOGNORM_INV(
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
    public static double ACT_GAMMA_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= 0)")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - rate parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (x < 0) return double.NaN;
        return Gamma.PDF(alpha, beta, x);
    }

    [ExcelFunction(Description = "Gamma cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_GAMMA_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Beta - rate parameter (> 0)")] double beta)
    {
        if (alpha <= 0 || beta <= 0) return double.NaN;
        if (x < 0) return 0;
        return Gamma.CDF(alpha, beta, x);
    }

    [ExcelFunction(Description = "Gamma inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_GAMMA_INV(
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

    [ExcelFunction(Description = "Pareto probability density function (PDF)", Category = "Actuarial.Distributions")]
    public static double ACT_PARETO_PDF(
        [ExcelArgument(Description = "Value at which to evaluate (>= xm)")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Xm - scale/minimum value (> 0)")] double xm)
    {
        if (alpha <= 0 || xm <= 0) return double.NaN;
        if (x < xm) return 0;
        return Pareto.PDF(xm, alpha, x);
    }

    [ExcelFunction(Description = "Pareto cumulative distribution function (CDF)", Category = "Actuarial.Distributions")]
    public static double ACT_PARETO_CDF(
        [ExcelArgument(Description = "Value at which to evaluate")] double x,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Xm - scale/minimum value (> 0)")] double xm)
    {
        if (alpha <= 0 || xm <= 0) return double.NaN;
        if (x < xm) return 0;
        return Pareto.CDF(xm, alpha, x);
    }

    [ExcelFunction(Description = "Pareto inverse CDF (quantile function)", Category = "Actuarial.Distributions")]
    public static double ACT_PARETO_INV(
        [ExcelArgument(Description = "Probability (0 to 1)")] double p,
        [ExcelArgument(Description = "Alpha - shape parameter (> 0)")] double alpha,
        [ExcelArgument(Description = "Xm - scale/minimum value (> 0)")] double xm)
    {
        if (alpha <= 0 || xm <= 0) return double.NaN;
        if (p < 0 || p > 1) return double.NaN;
        return Pareto.InvCDF(xm, alpha, p);
    }

    #endregion
}
