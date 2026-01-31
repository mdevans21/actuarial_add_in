using ExcelDna.Integration;
using MathNet.Numerics.Statistics;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Parameter estimation functions for fitting distributions to data.
/// Includes Maximum Likelihood Estimation (MLE) and Method of Moments (MoM) approaches.
/// </summary>
public static class Fitting
{
    #region Exponential Distribution

    [ExcelFunction(Description = "Fit exponential distribution to data. MLE estimate: λ = 1/mean. Returns rate parameter lambda.", Category = "Actuarial.Distributions")]
    public static object ACT_DIST_EXP_FIT(
        [ExcelArgument(Description = "Sample data (positive values)")] double[] data)
    {
        if (data == null || data.Length == 0)
            return "Error: Data array is empty";

        var validData = data.Where(x => x > 0).ToArray();
        if (validData.Length == 0)
            return "Error: No positive values in data";

        double mean = validData.Average();
        if (mean <= 0)
            return "Error: Mean must be positive";

        return 1.0 / mean;  // MLE estimate for lambda
    }

    #endregion

    #region Poisson Distribution

    [ExcelFunction(Description = "Fit Poisson distribution to count data. MLE estimate: λ = mean. Returns lambda parameter.", Category = "Actuarial.Distributions")]
    public static object ACT_DIST_POISSON_FIT(
        [ExcelArgument(Description = "Sample count data (non-negative integers)")] double[] data)
    {
        if (data == null || data.Length == 0)
            return "Error: Data array is empty";

        var validData = data.Where(x => x >= 0).ToArray();
        if (validData.Length == 0)
            return "Error: No non-negative values in data";

        return validData.Average();  // MLE estimate for lambda
    }

    #endregion

    #region Lognormal Distribution

    [ExcelFunction(Description = "Fit lognormal distribution to data. MLE: μ = mean(ln(x)), σ = std(ln(x)). Returns array [mu, sigma].", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_LOGNORM_FIT(
        [ExcelArgument(Description = "Sample data (positive values)")] double[] data)
    {
        if (data == null || data.Length < 2)
            return new object[] { "Error: Need at least 2 data points" };

        var validData = data.Where(x => x > 0).ToArray();
        if (validData.Length < 2)
            return new object[] { "Error: Need at least 2 positive values" };

        var logData = validData.Select(x => Math.Log(x)).ToArray();
        double mu = logData.Average();
        double sigma = Math.Sqrt(logData.Select(x => Math.Pow(x - mu, 2)).Sum() / (logData.Length - 1));

        return new object[] { mu, sigma };
    }

    #endregion

    #region Gamma Distribution

    [ExcelFunction(Description = "Fit gamma distribution using method of moments. Returns array [alpha (shape), beta (rate)]. Mean = α/β, Var = α/β².", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_GAMMA_FIT(
        [ExcelArgument(Description = "Sample data (positive values)")] double[] data)
    {
        if (data == null || data.Length < 2)
            return new object[] { "Error: Need at least 2 data points" };

        var validData = data.Where(x => x > 0).ToArray();
        if (validData.Length < 2)
            return new object[] { "Error: Need at least 2 positive values" };

        double mean = validData.Average();
        double variance = validData.Select(x => Math.Pow(x - mean, 2)).Sum() / (validData.Length - 1);

        if (variance <= 0)
            return new object[] { "Error: Variance must be positive" };

        // Method of moments: mean = alpha/beta, var = alpha/beta^2
        // So: beta = mean/var, alpha = mean * beta = mean^2/var
        double beta = mean / variance;
        double alpha = mean * beta;

        return new object[] { alpha, beta };
    }

    #endregion

    #region Pareto Distribution

    [ExcelFunction(Description = "Fit Pareto Type I distribution. MLE: α = n / Σln(x/xm). Returns array [alpha, xm] where xm = min(data).", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_PARETO_FIT(
        [ExcelArgument(Description = "Sample data (positive values)")] double[] data,
        [ExcelArgument(Description = "Known minimum xm (optional, uses min(data) if 0 or omitted)")] double knownXm = 0)
    {
        if (data == null || data.Length < 2)
            return new object[] { "Error: Need at least 2 data points" };

        var validData = data.Where(x => x > 0).ToArray();
        if (validData.Length < 2)
            return new object[] { "Error: Need at least 2 positive values" };

        double xm = knownXm > 0 ? knownXm : validData.Min();

        // Filter data >= xm
        var paretoData = validData.Where(x => x >= xm).ToArray();
        if (paretoData.Length < 2)
            return new object[] { "Error: Need at least 2 values >= xm" };

        // MLE for alpha: α = n / Σln(xi/xm)
        double sumLogRatios = paretoData.Select(x => Math.Log(x / xm)).Sum();
        if (sumLogRatios <= 0)
            return new object[] { "Error: Invalid data for Pareto fit" };

        double alpha = paretoData.Length / sumLogRatios;

        return new object[] { alpha, xm };
    }

    #endregion

    #region Weibull Distribution

    [ExcelFunction(Description = "Fit Weibull distribution using method of moments approximation. Returns array [k (shape), lambda (scale)].", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_WEIBULL_FIT(
        [ExcelArgument(Description = "Sample data (positive values)")] double[] data)
    {
        if (data == null || data.Length < 2)
            return new object[] { "Error: Need at least 2 data points" };

        var validData = data.Where(x => x > 0).ToArray();
        if (validData.Length < 2)
            return new object[] { "Error: Need at least 2 positive values" };

        double mean = validData.Average();
        double variance = validData.Select(x => Math.Pow(x - mean, 2)).Sum() / (validData.Length - 1);
        double cv = Math.Sqrt(variance) / mean;  // Coefficient of variation

        if (cv <= 0)
            return new object[] { "Error: Coefficient of variation must be positive" };

        // Approximate k from CV using empirical relationship
        // CV ≈ sqrt(Γ(1+2/k)/Γ(1+1/k)² - 1)
        // For CV < 1.2, k ≈ 1.2/CV works reasonably well
        // Use Newton-Raphson for better estimate
        double k = EstimateWeibullShape(cv);

        // lambda = mean / Γ(1 + 1/k)
        double gamma1pk = MathNet.Numerics.SpecialFunctions.Gamma(1 + 1 / k);
        double lambda = mean / gamma1pk;

        return new object[] { k, lambda };
    }

    private static double EstimateWeibullShape(double cv)
    {
        // Newton-Raphson to solve for k given CV
        // Start with approximation
        double k = cv < 0.5 ? 2.0 : (cv < 1.0 ? 1.2 / cv : 1.0);

        for (int i = 0; i < 50; i++)
        {
            double g1 = MathNet.Numerics.SpecialFunctions.Gamma(1 + 1 / k);
            double g2 = MathNet.Numerics.SpecialFunctions.Gamma(1 + 2 / k);
            double cvCalc = Math.Sqrt(g2 / (g1 * g1) - 1);

            if (Math.Abs(cvCalc - cv) < 1e-8)
                break;

            // Numerical derivative
            double h = 0.0001;
            double g1h = MathNet.Numerics.SpecialFunctions.Gamma(1 + 1 / (k + h));
            double g2h = MathNet.Numerics.SpecialFunctions.Gamma(1 + 2 / (k + h));
            double cvCalcH = Math.Sqrt(g2h / (g1h * g1h) - 1);
            double deriv = (cvCalcH - cvCalc) / h;

            if (Math.Abs(deriv) < 1e-10)
                break;

            k = k - (cvCalc - cv) / deriv;
            k = Math.Max(0.1, Math.Min(10, k));  // Keep in reasonable bounds
        }

        return k;
    }

    #endregion

    #region Generalized Pareto Distribution (GPD)

    [ExcelFunction(Description = "Fit GPD to exceedance data using probability-weighted moments. Returns array [xi (shape), sigma (scale)]. Data should be exceedances over threshold.", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_GPD_FIT(
        [ExcelArgument(Description = "Exceedance data (positive values representing amounts over threshold)")] double[] data)
    {
        if (data == null || data.Length < 10)
            return new object[] { "Error: Need at least 10 data points for reliable GPD fit" };

        var validData = data.Where(x => x > 0).OrderBy(x => x).ToArray();
        if (validData.Length < 10)
            return new object[] { "Error: Need at least 10 positive exceedances" };

        int n = validData.Length;

        // Probability-weighted moments (PWM) method
        // b0 = mean, b1 = (1/n) * Σ((i-0.35)/n * x[i])
        double b0 = validData.Average();
        double b1 = 0;
        for (int i = 0; i < n; i++)
        {
            double pj = (i + 1 - 0.35) / n;
            b1 += pj * validData[i];
        }
        b1 /= n;

        // PWM estimators for GPD:
        // xi = 2 - b0/(b0 - 2*b1)
        // sigma = 2*b0*b1/(b0 - 2*b1)
        double denom = b0 - 2 * b1;
        if (Math.Abs(denom) < 1e-10)
            return new object[] { "Error: Cannot estimate parameters (b0 ≈ 2*b1)" };

        double xi = 2 - b0 / denom;
        double sigma = 2 * b0 * b1 / denom;

        if (sigma <= 0)
            return new object[] { "Error: Estimated sigma is not positive" };

        return new object[] { xi, sigma };
    }

    #endregion

    #region Beta Distribution

    [ExcelFunction(Description = "Fit beta distribution using method of moments. Returns array [alpha, beta]. Data should be in (0,1).", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_BETA_FIT(
        [ExcelArgument(Description = "Sample data (values between 0 and 1)")] double[] data)
    {
        if (data == null || data.Length < 2)
            return new object[] { "Error: Need at least 2 data points" };

        var validData = data.Where(x => x > 0 && x < 1).ToArray();
        if (validData.Length < 2)
            return new object[] { "Error: Need at least 2 values in (0,1)" };

        double mean = validData.Average();
        double variance = validData.Select(x => Math.Pow(x - mean, 2)).Sum() / (validData.Length - 1);

        if (variance <= 0)
            return new object[] { "Error: Variance must be positive" };

        // Method of moments for Beta:
        // mean = α/(α+β)
        // var = αβ/((α+β)²(α+β+1))
        // Solving: α = mean * ((mean*(1-mean)/var) - 1)
        //          β = (1-mean) * ((mean*(1-mean)/var) - 1)
        double common = mean * (1 - mean) / variance - 1;

        if (common <= 0)
            return new object[] { "Error: Variance too large for beta distribution" };

        double alpha = mean * common;
        double beta = (1 - mean) * common;

        if (alpha <= 0 || beta <= 0)
            return new object[] { "Error: Estimated parameters are not positive" };

        return new object[] { alpha, beta };
    }

    #endregion

    #region Negative Binomial Distribution

    [ExcelFunction(Description = "Fit negative binomial distribution using method of moments. Returns array [r (successes), p (probability)]. Variance must exceed mean.", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_NEGBIN_FIT(
        [ExcelArgument(Description = "Sample count data (non-negative integers)")] double[] data)
    {
        if (data == null || data.Length < 2)
            return new object[] { "Error: Need at least 2 data points" };

        var validData = data.Where(x => x >= 0).ToArray();
        if (validData.Length < 2)
            return new object[] { "Error: Need at least 2 non-negative values" };

        double mean = validData.Average();
        double variance = validData.Select(x => Math.Pow(x - mean, 2)).Sum() / (validData.Length - 1);

        if (variance <= mean)
            return new object[] { "Error: Variance must exceed mean for negative binomial (overdispersion)" };

        // Method of moments:
        // mean = r(1-p)/p
        // var = r(1-p)/p²
        // So: p = mean/var, r = mean²/(var - mean)
        double p = mean / variance;
        double r = mean * mean / (variance - mean);

        if (p <= 0 || p >= 1 || r <= 0)
            return new object[] { "Error: Invalid parameter estimates" };

        return new object[] { r, p };
    }

    #endregion

    #region Burr Type XII Distribution

    [ExcelFunction(Description = "Fit Burr Type XII distribution using method of moments. Returns array [c, k, lambda]. Approximate fit - verify with goodness-of-fit tests.", Category = "Actuarial.Distributions")]
    public static object[] ACT_DIST_BURR_FIT(
        [ExcelArgument(Description = "Sample data (positive values)")] double[] data)
    {
        if (data == null || data.Length < 3)
            return new object[] { "Error: Need at least 3 data points" };

        var validData = data.Where(x => x > 0).ToArray();
        if (validData.Length < 3)
            return new object[] { "Error: Need at least 3 positive values" };

        double mean = validData.Average();
        double variance = validData.Select(x => Math.Pow(x - mean, 2)).Sum() / (validData.Length - 1);
        double cv = Math.Sqrt(variance) / mean;

        // Compute skewness
        double m3 = validData.Select(x => Math.Pow(x - mean, 3)).Sum() / validData.Length;
        double skewness = m3 / Math.Pow(Math.Sqrt(variance * (validData.Length - 1) / validData.Length), 3);

        // Use heuristics based on CV and skewness
        // For Burr XII: start with c=2, k based on CV
        double c = 2.0;
        double k = 2.0;

        // Iterative adjustment based on moments
        // This is approximate - Burr fitting typically requires numerical MLE
        if (cv > 1.5)
        {
            c = 1.5;
            k = 1.5;
        }
        else if (cv > 1.0)
        {
            c = 2.0;
            k = 2.0;
        }
        else
        {
            c = 3.0;
            k = 3.0;
        }

        // Adjust based on skewness
        if (skewness > 2)
        {
            k = Math.Max(1.0, k - 0.5);
        }

        // Estimate lambda from mean
        // Mean of Burr = lambda * B(k-1/c, 1+1/c) where B is beta function
        // Approximate: lambda ≈ mean for initial estimate, refine numerically
        double lambda = mean;

        // Simple iteration to match mean
        for (int i = 0; i < 20; i++)
        {
            double theoreticalMean = ComputeBurrMean(c, k, lambda);
            if (Math.Abs(theoreticalMean - mean) < mean * 0.001)
                break;
            lambda = lambda * mean / theoreticalMean;
        }

        return new object[] { c, k, lambda };
    }

    private static double ComputeBurrMean(double c, double k, double lambda)
    {
        // Mean = lambda * k * B(k - 1/c, 1 + 1/c) if k > 1/c
        if (k <= 1 / c)
            return double.PositiveInfinity;

        try
        {
            double b = MathNet.Numerics.SpecialFunctions.Beta(k - 1 / c, 1 + 1 / c);
            return lambda * k * b;
        }
        catch
        {
            return lambda;  // Fallback
        }
    }

    #endregion
}
