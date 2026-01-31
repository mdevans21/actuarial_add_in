using ExcelDna.Integration;
using MathNet.Numerics.Distributions;

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
}
