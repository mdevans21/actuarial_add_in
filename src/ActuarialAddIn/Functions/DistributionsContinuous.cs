using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics;

namespace ActuarialAddIn.Functions;

public static partial class Distributions
{
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

}
