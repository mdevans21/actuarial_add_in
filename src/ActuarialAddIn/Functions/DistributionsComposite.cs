using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics;

namespace ActuarialAddIn.Functions;

public static partial class Distributions
{
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
