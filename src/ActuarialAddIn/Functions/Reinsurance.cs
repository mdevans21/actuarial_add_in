using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Reinsurance functions for excess of loss, quota share, and increased limit factor calculations.
/// Includes return period analysis for catastrophe modeling.
/// </summary>
public static class Reinsurance
{
    #region Excess of Loss Layer Calculations

    [ExcelFunction(Description = "Calculate loss to an XOL layer: min(max(0, loss - attachment), limit). Standard formula for layer pricing and allocation.", Category = "Actuarial.Reinsurance")]
    public static double ACT_XOL_LAYER_LOSS(
        [ExcelArgument(Description = "Ground-up loss amount")] double groundUpLoss,
        [ExcelArgument(Description = "Layer attachment point (deductible)")] double attachment,
        [ExcelArgument(Description = "Layer limit")] double limit)
    {
        if (attachment < 0 || limit <= 0) return double.NaN;

        double excessLoss = Math.Max(0, groundUpLoss - attachment);
        return Math.Min(excessLoss, limit);
    }

    [ExcelFunction(Description = "Calculate expected layer loss assuming Pareto Type I severity. Uses LEV (Limited Expected Value) method: E[layer] = freq * P(X>attach) * (LEV(exhaust) - LEV(attach)). Requires alpha > 1 for finite mean.", Category = "Actuarial.Reinsurance")]
    public static double ACT_XOL_EXPECTED_LOSS(
        [ExcelArgument(Description = "Expected frequency (number of claims) at ground-up")] double frequency,
        [ExcelArgument(Description = "Layer attachment point")] double attachment,
        [ExcelArgument(Description = "Layer limit")] double limit,
        [ExcelArgument(Description = "Pareto alpha parameter (> 1 for finite mean). Typical values: 1.5-3 for liability.")] double alpha,
        [ExcelArgument(Description = "Pareto minimum (xm) - scale parameter")] double xm)
    {
        if (frequency < 0 || attachment < 0 || limit <= 0) return double.NaN;
        if (alpha <= 1 || xm <= 0) return double.NaN;

        // For Pareto distribution, expected limited loss in layer
        double exhaust = attachment + limit;

        // LEV (Limited Expected Value) for Pareto
        Func<double, double> lev = (d) =>
        {
            if (d <= xm) return d;
            double term1 = xm * alpha / (alpha - 1);
            double term2 = 1 - Math.Pow(xm / d, alpha - 1);
            return term1 * term2 + d * Math.Pow(xm / d, alpha);
        };

        // Probability of exceeding attachment
        double probExceed = attachment <= xm ? 1.0 : Math.Pow(xm / attachment, alpha);

        double layerLEV = lev(exhaust) - lev(attachment);
        return frequency * probExceed * layerLEV / (alpha / (alpha - 1) * xm);
    }

    #endregion

    #region Increased Limit Factors

    [ExcelFunction(Description = "Pareto increased-limit factor: ILF = LEV(target) / LEV(base) using the Pareto I limited expected value. Both limits must be >= xm.", Category = "Actuarial.Reinsurance")]
    public static double ACT_ILF_PARETO(
        [ExcelArgument(Description = "Target limit (the higher limit you're pricing)")] double targetLimit,
        [ExcelArgument(Description = "Base limit (the reference limit you're pricing relative to)")] double baseLimit,
        [ExcelArgument(Description = "Pareto alpha (α > 1). Lower α = heavier tail = higher ILFs. Typical 1.5–2.5 for liability.")] double alpha,
        [ExcelArgument(Description = "Pareto scale x_m > 0 (the minimum loss). Defaults to 1.0. Both target and base limits must be >= xm.")] double xm = 1.0)
    {
        if (targetLimit <= 0 || baseLimit <= 0 || alpha <= 1 || xm <= 0) return double.NaN;
        if (baseLimit < xm || targetLimit < xm) return double.NaN;

        // Pareto I LEV (Klugman A.2.3.1): for d >= xm,
        //   LEV(d) = α·xm/(α-1) · (1 - (xm/d)^(α-1)) + d · (xm/d)^α
        double Lev(double d)
        {
            double r = xm / d;
            return alpha * xm / (alpha - 1) * (1 - Math.Pow(r, alpha - 1))
                 + d * Math.Pow(r, alpha);
        }

        double levBase = Lev(baseLimit);
        if (levBase <= 0) return double.NaN;
        return Lev(targetLimit) / levBase;
    }

    #endregion

    #region Return Period Functions

    [ExcelFunction(Description = "Interpolate loss for a target return period from an EP curve. Uses log-linear interpolation by default (standard for cat modeling).", Category = "Actuarial.Reinsurance")]
    public static double ACT_RETURN_PERIOD_LOSS(
        [ExcelArgument(Description = "Return periods (column)")] double[] returnPeriods,
        [ExcelArgument(Description = "Corresponding losses (column)")] double[] losses,
        [ExcelArgument(Description = "Target return period")] double targetReturnPeriod,
        [ExcelArgument(Description = "Interpolation method: 'LOG' (default) or 'LINEAR'")] string method = "LOG")
    {
        if (returnPeriods.Length != losses.Length || returnPeriods.Length < 2)
            return double.NaN;

        if (targetReturnPeriod <= 0) return double.NaN;

        // Sort by return period
        var pairs = returnPeriods.Zip(losses, (rp, l) => new { RP = rp, Loss = l })
                                  .OrderBy(p => p.RP)
                                  .ToArray();

        // Find bracketing points
        for (int i = 0; i < pairs.Length - 1; i++)
        {
            if (targetReturnPeriod >= pairs[i].RP && targetReturnPeriod <= pairs[i + 1].RP)
            {
                if (method.ToUpper() == "LOG")
                {
                    // Log-linear interpolation (common for return periods)
                    double logRP1 = Math.Log(pairs[i].RP);
                    double logRP2 = Math.Log(pairs[i + 1].RP);
                    double logTarget = Math.Log(targetReturnPeriod);
                    double t = (logTarget - logRP1) / (logRP2 - logRP1);
                    return pairs[i].Loss + t * (pairs[i + 1].Loss - pairs[i].Loss);
                }
                else
                {
                    // Linear interpolation
                    double t = (targetReturnPeriod - pairs[i].RP) / (pairs[i + 1].RP - pairs[i].RP);
                    return pairs[i].Loss + t * (pairs[i + 1].Loss - pairs[i].Loss);
                }
            }
        }

        // Extrapolation
        if (targetReturnPeriod < pairs[0].RP)
            return pairs[0].Loss;
        else
            return pairs[pairs.Length - 1].Loss;
    }

    [ExcelFunction(Description = "Generate a return period loss table for target return periods", Category = "Actuarial.Reinsurance")]
    public static object[,] ACT_RETURN_PERIOD_TABLE(
        [ExcelArgument(Description = "Return periods (column)")] double[] returnPeriods,
        [ExcelArgument(Description = "Corresponding losses (column)")] double[] losses,
        [ExcelArgument(Description = "Target return periods (column)")] double[] targetReturnPeriods,
        [ExcelArgument(Description = "Interpolation method: 'LOG' (default) or 'LINEAR'")] string method = "LOG")
    {
        if (returnPeriods.Length != losses.Length || returnPeriods.Length < 2)
            return new object[,] { { "Error: Return periods and losses must be same length" } };

        if (targetReturnPeriods.Length < 1)
            return new object[,] { { "Error: Target return periods required" } };

        var result = new object[targetReturnPeriods.Length, 2];
        for (int i = 0; i < targetReturnPeriods.Length; i++)
        {
            double rp = targetReturnPeriods[i];
            result[i, 0] = rp;
            result[i, 1] = ACT_RETURN_PERIOD_LOSS(returnPeriods, losses, rp, method);
        }

        return result;
    }

    [ExcelFunction(Description = "Calculate AAL (Average Annual Loss) from OEP curve via numerical integration. AAL = integral of loss over exceedance probability. Standard cat modeling metric.", Category = "Actuarial.Reinsurance")]
    public static double ACT_AAL_FROM_OEP(
        [ExcelArgument(Description = "Return periods (column)")] double[] returnPeriods,
        [ExcelArgument(Description = "Corresponding OEP losses (column)")] double[] oepLosses)
    {
        if (returnPeriods.Length != oepLosses.Length || returnPeriods.Length < 2)
            return double.NaN;

        // Sort by return period descending
        var pairs = returnPeriods.Zip(oepLosses, (rp, l) => new { RP = rp, Loss = l, EP = 1.0 / rp })
                                  .OrderByDescending(p => p.RP)
                                  .ToArray();

        // Numerical integration using trapezoidal rule
        double aal = 0;
        for (int i = 0; i < pairs.Length - 1; i++)
        {
            double deltaEP = pairs[i + 1].EP - pairs[i].EP;
            double avgLoss = (pairs[i].Loss + pairs[i + 1].Loss) / 2;
            aal += deltaEP * avgLoss;
        }

        // Add tail (from highest RP to infinity - assume zero loss beyond)
        // Add head (from EP=1 to lowest RP - assume constant loss)
        aal += pairs[pairs.Length - 1].EP * pairs[pairs.Length - 1].Loss;

        return aal;
    }

    #endregion
}
