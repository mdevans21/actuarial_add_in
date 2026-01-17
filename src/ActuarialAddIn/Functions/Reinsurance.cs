using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

public static class Reinsurance
{
    #region Excess of Loss Layer Calculations

    [ExcelFunction(Description = "Calculate loss to an excess of loss layer", Category = "Actuarial.Reinsurance")]
    public static double ACT_XOL_LAYER_LOSS(
        [ExcelArgument(Description = "Ground-up loss amount")] double groundUpLoss,
        [ExcelArgument(Description = "Layer attachment point (deductible)")] double attachment,
        [ExcelArgument(Description = "Layer limit")] double limit)
    {
        if (attachment < 0 || limit <= 0) return double.NaN;

        double excessLoss = Math.Max(0, groundUpLoss - attachment);
        return Math.Min(excessLoss, limit);
    }

    [ExcelFunction(Description = "Calculate ceded loss for quota share", Category = "Actuarial.Reinsurance")]
    public static double ACT_QS_CEDED(
        [ExcelArgument(Description = "Ground-up loss amount")] double groundUpLoss,
        [ExcelArgument(Description = "Cession percentage (0 to 1)")] double cessionPct)
    {
        if (cessionPct < 0 || cessionPct > 1) return double.NaN;
        return groundUpLoss * cessionPct;
    }

    [ExcelFunction(Description = "Calculate loss after applying aggregate deductible and limit", Category = "Actuarial.Reinsurance")]
    public static double ACT_AGGREGATE_LAYER(
        [ExcelArgument(Description = "Total aggregate losses")] double aggregateLoss,
        [ExcelArgument(Description = "Aggregate deductible")] double aggDeductible,
        [ExcelArgument(Description = "Aggregate limit")] double aggLimit)
    {
        if (aggDeductible < 0 || aggLimit <= 0) return double.NaN;

        double excessLoss = Math.Max(0, aggregateLoss - aggDeductible);
        return Math.Min(excessLoss, aggLimit);
    }

    [ExcelFunction(Description = "Calculate expected layer loss using Pareto severity", Category = "Actuarial.Reinsurance")]
    public static double ACT_XOL_EXPECTED_LOSS(
        [ExcelArgument(Description = "Expected frequency (number of claims)")] double frequency,
        [ExcelArgument(Description = "Layer attachment point")] double attachment,
        [ExcelArgument(Description = "Layer limit")] double limit,
        [ExcelArgument(Description = "Pareto alpha parameter")] double alpha,
        [ExcelArgument(Description = "Pareto minimum (xm)")] double xm)
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

    [ExcelFunction(Description = "Calculate increased limit factor using Pareto distribution", Category = "Actuarial.Reinsurance")]
    public static double ACT_ILF_PARETO(
        [ExcelArgument(Description = "Target limit")] double targetLimit,
        [ExcelArgument(Description = "Base limit")] double baseLimit,
        [ExcelArgument(Description = "Pareto alpha parameter (> 1)")] double alpha)
    {
        if (targetLimit <= 0 || baseLimit <= 0 || alpha <= 1) return double.NaN;
        if (targetLimit <= baseLimit) return 1.0;

        // ILF based on Pareto LEV ratio
        // ILF = LEV(target) / LEV(base)
        // For Pareto: LEV(L) = L * [1 - (1/(1+L))^(alpha-1)] * alpha/(alpha-1)
        // Simplified: ILF = 1 - (base/target)^(alpha-1)  / [1 - approaches 1]

        // Using standard actuarial ILF formula for Pareto
        double ilf = (1 - Math.Pow(baseLimit / targetLimit, alpha - 1)) /
                     (1 - Math.Pow(baseLimit / (baseLimit * 1000), alpha - 1));

        // Normalize to base limit
        return 1 + (ilf * (alpha - 1) / alpha) * (targetLimit / baseLimit - 1);
    }

    #endregion

    #region Return Period Functions

    [ExcelFunction(Description = "Calculate loss from return period using exceedance probability curve", Category = "Actuarial.Reinsurance")]
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

    [ExcelFunction(Description = "Calculate AAL (Average Annual Loss) from OEP curve", Category = "Actuarial.Reinsurance")]
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
