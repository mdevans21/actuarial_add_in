using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Credibility theory and experience rating functions for actuarial pricing.
/// References:
/// - Bühlmann, H. (1967). "Experience rating and credibility." ASTIN Bulletin 4(3): 199-207.
/// - Bühlmann, H. and Straub, E. (1970). "Glaubwürdigkeit für Schadensätze." Bulletin of Swiss Actuaries 70: 111-133.
/// </summary>
public static class Credibility
{
    #region Bühlmann Credibility

    [ExcelFunction(Description = "Bühlmann credibility factor: Z = n / (n + k) where k = σ²/τ². Classic limited fluctuation credibility.", Category = "Actuarial.Credibility")]
    public static object ACT_CREDIBILITY_BUHLMANN(
        [ExcelArgument(Description = "Number of exposure units or claims (n)")] double n,
        [ExcelArgument(Description = "k parameter = Expected Value of Process Variance / Variance of Hypothetical Means = E[σ²]/Var[μ]")] double k)
    {
        if (n < 0)
            return "Error: n must be non-negative";
        if (k <= 0)
            return "Error: k must be positive";

        return n / (n + k);
    }

    [ExcelFunction(Description = "Bühlmann credibility-weighted estimate: Z * X̄ + (1-Z) * μ₀. Blends experience with prior.", Category = "Actuarial.Credibility")]
    public static object ACT_CREDIBILITY_ESTIMATE(
        [ExcelArgument(Description = "Observed experience mean (X̄)")] double observedMean,
        [ExcelArgument(Description = "Prior/collateral mean (μ₀)")] double priorMean,
        [ExcelArgument(Description = "Credibility factor Z (0 to 1)")] double credibility)
    {
        if (credibility < 0 || credibility > 1)
            return "Error: Credibility must be between 0 and 1";

        return credibility * observedMean + (1 - credibility) * priorMean;
    }

    [ExcelFunction(Description = "Calculate Bühlmann k parameter from data. k = E[σ²]/Var[μ] estimated from group data.", Category = "Actuarial.Credibility")]
    public static object ACT_CREDIBILITY_K(
        [ExcelArgument(Description = "Array of observed means by group")] double[] groupMeans,
        [ExcelArgument(Description = "Array of exposure/weight by group")] double[] groupWeights,
        [ExcelArgument(Description = "Array of within-group variances (optional, estimated if omitted)")] double[]? groupVariances = null)
    {
        if (groupMeans == null || groupWeights == null)
            return "Error: Group means and weights required";

        int m = groupMeans.Length;
        if (groupWeights.Length != m)
            return "Error: Arrays must be same length";

        if (m < 2)
            return "Error: Need at least 2 groups";

        // Calculate grand mean (weighted)
        double totalWeight = groupWeights.Sum();
        double grandMean = groupMeans.Zip(groupWeights, (mean, w) => mean * w).Sum() / totalWeight;

        // Estimate variance of hypothetical means (between-group variance)
        double betweenGroupVar = 0;
        for (int i = 0; i < m; i++)
        {
            betweenGroupVar += groupWeights[i] * Math.Pow(groupMeans[i] - grandMean, 2);
        }
        betweenGroupVar /= (totalWeight - totalWeight * totalWeight / groupWeights.Sum());

        // Estimate expected process variance (within-group variance)
        double withinGroupVar;
        if (groupVariances != null && groupVariances.Length == m)
        {
            // Use provided variances
            withinGroupVar = groupVariances.Zip(groupWeights, (v, w) => v * w).Sum() / totalWeight;
        }
        else
        {
            // Estimate from the data - simplified assumption
            withinGroupVar = betweenGroupVar;  // Rough approximation
        }

        // Adjust between-group variance for within-group contribution
        double avgWeight = totalWeight / m;
        double tau2 = betweenGroupVar - withinGroupVar / avgWeight;
        if (tau2 <= 0)
            tau2 = betweenGroupVar * 0.1;  // Floor at 10% of between-group

        double k = withinGroupVar / tau2;
        return k;
    }

    #endregion

    #region Bühlmann-Straub Model

    [ExcelFunction(Description = "Bühlmann-Straub credibility for heterogeneous exposures. Z = w / (w + k) where w is exposure weight.", Category = "Actuarial.Credibility")]
    public static object ACT_CREDIBILITY_BUHLMANN_STRAUB(
        [ExcelArgument(Description = "Exposure weight for this risk (w)")] double weight,
        [ExcelArgument(Description = "k parameter = a/s² where a = process variance parameter, s² = variance of hypothetical means")] double k)
    {
        if (weight < 0)
            return "Error: Weight must be non-negative";
        if (k <= 0)
            return "Error: k must be positive";

        return weight / (weight + k);
    }

    [ExcelFunction(Description = "Estimate Bühlmann-Straub parameters from grouped data. Returns array [k, grandMean, a (process var), s² (structural var)].", Category = "Actuarial.Credibility")]
    public static object[] ACT_BUHLMANN_STRAUB_PARAMS(
        [ExcelArgument(Description = "Array of loss ratios or pure premiums by group")] double[] groupMeans,
        [ExcelArgument(Description = "Array of exposure weights by group")] double[] groupWeights,
        [ExcelArgument(Description = "Array of within-group sum of squared deviations (optional)")] double[]? groupSSE = null)
    {
        if (groupMeans == null || groupWeights == null)
            return new object[] { "Error: Group means and weights required" };

        int m = groupMeans.Length;
        if (groupWeights.Length != m)
            return new object[] { "Error: Arrays must be same length" };

        if (m < 2)
            return new object[] { "Error: Need at least 2 groups" };

        double totalWeight = groupWeights.Sum();
        double sumW2 = groupWeights.Sum(w => w * w);

        // Grand mean (weighted)
        double grandMean = groupMeans.Zip(groupWeights, (x, w) => x * w).Sum() / totalWeight;

        // Between-group sum of squares
        double SSB = 0;
        for (int i = 0; i < m; i++)
        {
            SSB += groupWeights[i] * Math.Pow(groupMeans[i] - grandMean, 2);
        }

        // Within-group variance estimate (a)
        double a;
        if (groupSSE != null && groupSSE.Length == m)
        {
            // Use provided SSE
            double totalSSE = groupSSE.Sum();
            double dfWithin = totalWeight - m;  // Approximate
            a = dfWithin > 0 ? totalSSE / dfWithin : 0;
        }
        else
        {
            // Estimate a from between-group variance (rough)
            a = SSB / (m - 1);
        }

        // Structural variance estimate (s²)
        double c = totalWeight - sumW2 / totalWeight;
        double s2 = (SSB - (m - 1) * a) / c;
        if (s2 <= 0)
            s2 = SSB / (c * 10);  // Floor

        // k parameter
        double k = a / s2;

        return new object[] { k, grandMean, a, s2 };
    }

    #endregion

    #region Experience Rating

    [ExcelFunction(Description = "Experience modification factor: EMod = Z * (A/E) + (1-Z). Returns multiplier to apply to manual premium.", Category = "Actuarial.Credibility")]
    public static object ACT_EXPERIENCE_MOD(
        [ExcelArgument(Description = "Actual losses")] double actualLosses,
        [ExcelArgument(Description = "Expected losses (at manual rates)")] double expectedLosses,
        [ExcelArgument(Description = "Credibility factor Z (0 to 1)")] double credibility)
    {
        if (expectedLosses <= 0)
            return "Error: Expected losses must be positive";
        if (credibility < 0 || credibility > 1)
            return "Error: Credibility must be between 0 and 1";

        double lossRatio = actualLosses / expectedLosses;
        return credibility * lossRatio + (1 - credibility);
    }

    [ExcelFunction(Description = "Experience mod with primary/excess split (workers comp style). Limits large losses to reduce volatility.", Category = "Actuarial.Credibility")]
    public static object ACT_EXPERIENCE_MOD_SPLIT(
        [ExcelArgument(Description = "Array of individual loss amounts")] double[] losses,
        [ExcelArgument(Description = "Expected primary losses")] double expectedPrimary,
        [ExcelArgument(Description = "Expected excess losses")] double expectedExcess,
        [ExcelArgument(Description = "Primary threshold (losses capped at this amount)")] double primaryThreshold,
        [ExcelArgument(Description = "Primary credibility (Zp)")] double credibilityPrimary,
        [ExcelArgument(Description = "Excess credibility (Ze)")] double credibilityExcess)
    {
        if (losses == null || losses.Length == 0)
            return "Error: Loss array required";
        if (expectedPrimary <= 0 || expectedExcess < 0)
            return "Error: Expected values must be positive";
        if (credibilityPrimary < 0 || credibilityPrimary > 1 || credibilityExcess < 0 || credibilityExcess > 1)
            return "Error: Credibilities must be between 0 and 1";

        double actualPrimary = 0;
        double actualExcess = 0;

        foreach (double loss in losses)
        {
            if (loss <= primaryThreshold)
            {
                actualPrimary += loss;
            }
            else
            {
                actualPrimary += primaryThreshold;
                actualExcess += loss - primaryThreshold;
            }
        }

        double expectedTotal = expectedPrimary + expectedExcess;

        // Experience mod formula
        double primaryComponent = credibilityPrimary * (actualPrimary - expectedPrimary);
        double excessComponent = credibilityExcess * (actualExcess - expectedExcess);

        double mod = (expectedTotal + primaryComponent + excessComponent) / expectedTotal;

        // Typically floored and capped
        return Math.Max(0.5, Math.Min(2.0, mod));
    }

    [ExcelFunction(Description = "Full credibility standard: n₀ = (z/r)² * CV² where z is normal quantile, r is error tolerance, CV is coefficient of variation.", Category = "Actuarial.Credibility")]
    public static object ACT_FULL_CREDIBILITY_STANDARD(
        [ExcelArgument(Description = "Confidence level (e.g., 0.90 for 90%)")] double confidence,
        [ExcelArgument(Description = "Error tolerance (e.g., 0.05 for ±5%)")] double errorTolerance,
        [ExcelArgument(Description = "Coefficient of variation of losses")] double cv)
    {
        if (confidence <= 0 || confidence >= 1)
            return "Error: Confidence must be between 0 and 1";
        if (errorTolerance <= 0)
            return "Error: Error tolerance must be positive";
        if (cv <= 0)
            return "Error: CV must be positive";

        // Two-tailed z-score
        double z = MathNet.Numerics.Distributions.Normal.InvCDF(0, 1, (1 + confidence) / 2);

        double n0 = Math.Pow(z / errorTolerance, 2) * cv * cv;

        return n0;
    }

    [ExcelFunction(Description = "Partial credibility using square root rule: Z = sqrt(n/n₀) capped at 1.", Category = "Actuarial.Credibility")]
    public static object ACT_PARTIAL_CREDIBILITY(
        [ExcelArgument(Description = "Actual number of claims or exposure units")] double n,
        [ExcelArgument(Description = "Full credibility standard (n₀)")] double n0)
    {
        if (n < 0)
            return "Error: n must be non-negative";
        if (n0 <= 0)
            return "Error: n₀ must be positive";

        double z = Math.Sqrt(n / n0);
        return Math.Min(1.0, z);
    }

    #endregion

    #region Increased Limit Factor Tables

    [ExcelFunction(Description = "Generate ILF table using Pareto severity. Returns limits and corresponding ILFs.", Category = "Actuarial.Credibility")]
    public static object[,] ACT_ILF_TABLE(
        [ExcelArgument(Description = "Base limit (denominator for ILFs)")] double baseLimit,
        [ExcelArgument(Description = "Pareto alpha parameter (> 1)")] double alpha,
        [ExcelArgument(Description = "Target limits (array)")] double[] targetLimits)
    {
        if (baseLimit <= 0)
            return new object[,] { { "Error: Base limit must be positive" } };
        if (alpha <= 1)
            return new object[,] { { "Error: Alpha must be greater than 1 for finite mean" } };
        if (targetLimits == null || targetLimits.Length == 0)
            return new object[,] { { "Error: Target limits required" } };

        int n = targetLimits.Length;
        var result = new object[n, 2];

        // LEV for Pareto: LEV(L) = (alpha * xm / (alpha - 1)) * (1 - (xm/L)^(alpha-1)) for L >= xm
        // Simplified ILF assuming xm << limits: ILF(L) ≈ (1 - (b/L)^(a-1)) / (1 - (b/B)^(a-1))
        // where b is a scale parameter and B is base limit

        // Using the simplified formula for large limits relative to minimum:
        // ILF(L) = (1 - k^(alpha-1)) where k = baseLimit/L
        Func<double, double> lev = (limit) =>
        {
            // Normalized LEV formula
            return 1 - Math.Pow(baseLimit / Math.Max(limit, baseLimit), alpha - 1);
        };

        double baseLEV = lev(baseLimit);

        for (int i = 0; i < n; i++)
        {
            double limit = targetLimits[i];
            result[i, 0] = limit;

            if (limit <= 0)
            {
                result[i, 1] = "Error";
            }
            else if (limit <= baseLimit)
            {
                // ILF for limits below base
                result[i, 1] = lev(limit) / baseLEV;
            }
            else
            {
                // ILF for limits above base
                result[i, 1] = lev(limit) / baseLEV;
            }
        }

        return result;
    }

    [ExcelFunction(Description = "Generate standard ILF table with common limits. Returns limits from base to max.", Category = "Actuarial.Credibility")]
    public static object[,] ACT_ILF_TABLE_STANDARD(
        [ExcelArgument(Description = "Base limit (e.g., 100000)")] double baseLimit,
        [ExcelArgument(Description = "Maximum limit (e.g., 10000000)")] double maxLimit,
        [ExcelArgument(Description = "Pareto alpha parameter (> 1)")] double alpha)
    {
        if (baseLimit <= 0 || maxLimit <= 0)
            return new object[,] { { "Error: Limits must be positive" } };
        if (maxLimit <= baseLimit)
            return new object[,] { { "Error: Max limit must exceed base limit" } };
        if (alpha <= 1)
            return new object[,] { { "Error: Alpha must be greater than 1" } };

        // Standard limits progression
        var limits = new List<double> { baseLimit };
        double current = baseLimit;

        // Add common progression: 2x, 3x, 5x, 10x, etc.
        double[] multipliers = { 2, 3, 5, 10, 20, 30, 50, 100 };
        foreach (double mult in multipliers)
        {
            double limit = baseLimit * mult;
            if (limit <= maxLimit && limit > current)
            {
                limits.Add(limit);
                current = limit;
            }
        }

        if (limits[limits.Count - 1] < maxLimit)
            limits.Add(maxLimit);

        return ACT_ILF_TABLE(baseLimit, alpha, limits.ToArray());
    }

    [ExcelFunction(Description = "Calculate layer ILF: (ILF at exhaustion - ILF at attachment) / (exhaustion - attachment) * base. For pricing excess layers.", Category = "Actuarial.Credibility")]
    public static object ACT_ILF_LAYER(
        [ExcelArgument(Description = "Layer attachment point")] double attachment,
        [ExcelArgument(Description = "Layer limit")] double layerLimit,
        [ExcelArgument(Description = "Base limit for ILF curve")] double baseLimit,
        [ExcelArgument(Description = "Pareto alpha parameter (> 1)")] double alpha)
    {
        if (attachment < 0 || layerLimit <= 0 || baseLimit <= 0)
            return "Error: Limits must be positive";
        if (alpha <= 1)
            return "Error: Alpha must be greater than 1";

        double exhaustion = attachment + layerLimit;

        // LEV formula
        Func<double, double> lev = (limit) =>
        {
            if (limit <= 0) return 0;
            return 1 - Math.Pow(baseLimit / Math.Max(limit, baseLimit), alpha - 1);
        };

        double baseLEV = lev(baseLimit);
        if (baseLEV <= 0) return "Error: Invalid base LEV";

        double ilfAttach = attachment <= baseLimit ? lev(attachment) / baseLEV : lev(attachment) / baseLEV;
        double ilfExhaust = lev(exhaustion) / baseLEV;

        // Layer ILF = difference in ILFs
        double layerILF = ilfExhaust - ilfAttach;

        return layerILF;
    }

    #endregion

    #region Loss Rating

    [ExcelFunction(Description = "Calculate burning cost (experience rate) from historical losses. BC = Trended Losses / On-Level Premium.", Category = "Actuarial.Credibility")]
    public static object ACT_BURNING_COST(
        [ExcelArgument(Description = "Historical losses by year")] double[] losses,
        [ExcelArgument(Description = "Earned premium by year")] double[] premium,
        [ExcelArgument(Description = "Loss trend factors by year (to bring to current level)")] double[]? lossTrend = null,
        [ExcelArgument(Description = "Premium on-level factors by year")] double[]? premiumOnLevel = null)
    {
        if (losses == null || premium == null)
            return "Error: Losses and premium arrays required";

        int n = losses.Length;
        if (premium.Length != n)
            return "Error: Arrays must be same length";

        // Default trends to 1.0 if not provided
        double[] lTrend = lossTrend ?? Enumerable.Repeat(1.0, n).ToArray();
        double[] pOnLevel = premiumOnLevel ?? Enumerable.Repeat(1.0, n).ToArray();

        if (lTrend.Length != n || pOnLevel.Length != n)
            return "Error: Trend arrays must match data length";

        double trendedLosses = 0;
        double onLevelPremium = 0;

        for (int i = 0; i < n; i++)
        {
            trendedLosses += losses[i] * lTrend[i];
            onLevelPremium += premium[i] * pOnLevel[i];
        }

        if (onLevelPremium <= 0)
            return "Error: On-level premium must be positive";

        return trendedLosses / onLevelPremium;
    }

    [ExcelFunction(Description = "Credibility-weighted rate: Blends burning cost with manual rate using credibility.", Category = "Actuarial.Credibility")]
    public static object ACT_CREDIBILITY_RATE(
        [ExcelArgument(Description = "Burning cost (experience rate)")] double burningCost,
        [ExcelArgument(Description = "Manual rate (from class rating)")] double manualRate,
        [ExcelArgument(Description = "Credibility factor Z")] double credibility,
        [ExcelArgument(Description = "Loss development factor (to ultimate)")] double ldf = 1.0,
        [ExcelArgument(Description = "Trend factor (to prospective period)")] double trend = 1.0)
    {
        if (credibility < 0 || credibility > 1)
            return "Error: Credibility must be between 0 and 1";
        if (ldf <= 0 || trend <= 0)
            return "Error: LDF and trend must be positive";

        // Develop and trend the burning cost
        double developedBC = burningCost * ldf * trend;

        // Credibility-weighted rate
        return credibility * developedBC + (1 - credibility) * manualRate;
    }

    #endregion

    #region Loss Elimination and Deductible Pricing

    [ExcelFunction(Description = "Loss Elimination Ratio: proportion of expected loss eliminated by deductible. LER = LEV(d) / E[X] using Pareto severity.", Category = "Actuarial.Credibility")]
    public static object ACT_LOSS_ELIMINATION_RATIO(
        [ExcelArgument(Description = "Deductible amount")] double deductible,
        [ExcelArgument(Description = "Pareto alpha parameter (> 1)")] double alpha,
        [ExcelArgument(Description = "Pareto minimum (xm) - scale parameter")] double xm)
    {
        if (deductible < 0)
            return "Error: Deductible must be non-negative";
        if (alpha <= 1)
            return "Error: Alpha must be greater than 1 for finite mean";
        if (xm <= 0)
            return "Error: xm must be positive";

        if (deductible <= 0)
            return 0.0;

        // Expected value of Pareto: E[X] = alpha * xm / (alpha - 1)
        double expectedValue = alpha * xm / (alpha - 1);

        // LEV (Limited Expected Value) for Pareto at deductible
        double lev;
        if (deductible <= xm)
        {
            lev = deductible;
        }
        else
        {
            // LEV(d) = xm * alpha / (alpha - 1) * (1 - (xm/d)^(alpha-1)) + d * (xm/d)^alpha
            double ratio = xm / deductible;
            lev = xm * alpha / (alpha - 1) * (1 - Math.Pow(ratio, alpha - 1)) + deductible * Math.Pow(ratio, alpha);
        }

        // LER = LEV(d) / E[X]
        return lev / expectedValue;
    }

    [ExcelFunction(Description = "Loss Elimination Ratio using Lognormal severity. LER = E[min(X,d)] / E[X].", Category = "Actuarial.Credibility")]
    public static object ACT_LOSS_ELIMINATION_RATIO_LOGNORMAL(
        [ExcelArgument(Description = "Deductible amount")] double deductible,
        [ExcelArgument(Description = "Lognormal mu (log-mean)")] double mu,
        [ExcelArgument(Description = "Lognormal sigma (log-std dev, > 0)")] double sigma)
    {
        if (deductible < 0)
            return "Error: Deductible must be non-negative";
        if (sigma <= 0)
            return "Error: Sigma must be positive";

        if (deductible <= 0)
            return 0.0;

        // E[X] for lognormal
        double expectedValue = Math.Exp(mu + sigma * sigma / 2);

        // LEV for lognormal: E[min(X,d)] = E[X] * Φ((ln(d) - mu - σ²) / σ) + d * (1 - Φ((ln(d) - mu) / σ))
        double logD = Math.Log(deductible);
        double z1 = (logD - mu - sigma * sigma) / sigma;
        double z2 = (logD - mu) / sigma;

        double phi1 = MathNet.Numerics.Distributions.Normal.CDF(0, 1, z1);
        double phi2 = MathNet.Numerics.Distributions.Normal.CDF(0, 1, z2);

        double lev = expectedValue * phi1 + deductible * (1 - phi2);

        return lev / expectedValue;
    }

    [ExcelFunction(Description = "Deductible credit factor: 1 - LER. Multiply by base premium to get premium with deductible.", Category = "Actuarial.Credibility")]
    public static object ACT_DEDUCTIBLE_CREDIT(
        [ExcelArgument(Description = "Loss Elimination Ratio (from ACT_LOSS_ELIMINATION_RATIO)")] double ler,
        [ExcelArgument(Description = "Expense adjustment factor (optional, default 1.0). Use < 1 if expenses are fixed and don't reduce with deductible.")] double expenseAdjust = 1.0)
    {
        if (ler < 0 || ler > 1)
            return "Error: LER must be between 0 and 1";
        if (expenseAdjust <= 0)
            return "Error: Expense adjustment must be positive";

        // Deductible credit = 1 - LER, adjusted for expense portion
        return 1 - ler * expenseAdjust;
    }

    [ExcelFunction(Description = "Calculate premium with deductible from ground-up premium using LER.", Category = "Actuarial.Credibility")]
    public static object ACT_PREMIUM_WITH_DEDUCTIBLE(
        [ExcelArgument(Description = "Ground-up premium (first dollar coverage)")] double groundUpPremium,
        [ExcelArgument(Description = "Loss Elimination Ratio")] double ler,
        [ExcelArgument(Description = "Loss ratio portion of premium (e.g., 0.65 if 65% expected loss ratio)")] double lossRatioPortion)
    {
        if (groundUpPremium < 0)
            return "Error: Premium must be non-negative";
        if (ler < 0 || ler > 1)
            return "Error: LER must be between 0 and 1";
        if (lossRatioPortion <= 0 || lossRatioPortion > 1)
            return "Error: Loss ratio portion must be between 0 and 1";

        // Premium reduction only applies to loss portion, not expense portion
        double lossReduction = groundUpPremium * lossRatioPortion * ler;
        return groundUpPremium - lossReduction;
    }

    #endregion

    #region Large Loss Loading

    [ExcelFunction(Description = "Large loss loading factor based on excess frequency. Load = excess freq * avg excess severity / base premium.", Category = "Actuarial.Credibility")]
    public static object ACT_LARGE_LOSS_LOAD(
        [ExcelArgument(Description = "Expected frequency of claims exceeding threshold")] double excessFrequency,
        [ExcelArgument(Description = "Average severity of excess claims (above threshold)")] double avgExcessSeverity,
        [ExcelArgument(Description = "Base expected losses (losses up to threshold)")] double baseExpectedLoss,
        [ExcelArgument(Description = "Apply as multiplicative factor (true) or additive (false)?")] bool multiplicative = true)
    {
        if (excessFrequency < 0)
            return "Error: Frequency must be non-negative";
        if (avgExcessSeverity < 0)
            return "Error: Severity must be non-negative";
        if (baseExpectedLoss <= 0)
            return "Error: Base expected loss must be positive";

        double excessLoad = excessFrequency * avgExcessSeverity;

        if (multiplicative)
        {
            return 1 + excessLoad / baseExpectedLoss;
        }
        else
        {
            return excessLoad;
        }
    }

    [ExcelFunction(Description = "Shock loss load: adds explicit charge for low-frequency high-severity events. Uses Poisson frequency assumption.", Category = "Actuarial.Credibility")]
    public static object ACT_SHOCK_LOSS_LOAD(
        [ExcelArgument(Description = "Annual frequency of shock events (e.g., 0.02 for 1-in-50 year)")] double annualFrequency,
        [ExcelArgument(Description = "Expected severity if shock occurs")] double shockSeverity,
        [ExcelArgument(Description = "Risk load multiplier (e.g., 1.5 for 50% risk margin)")] double riskMultiplier = 1.0)
    {
        if (annualFrequency < 0)
            return "Error: Frequency must be non-negative";
        if (shockSeverity < 0)
            return "Error: Severity must be non-negative";
        if (riskMultiplier < 1)
            return "Error: Risk multiplier should be >= 1";

        // Expected value plus risk margin
        return annualFrequency * shockSeverity * riskMultiplier;
    }

    #endregion

    #region Schedule Rating

    [ExcelFunction(Description = "Calculate schedule-rated premium. Applies debits/credits to base premium.", Category = "Actuarial.Credibility")]
    public static object ACT_SCHEDULE_RATING(
        [ExcelArgument(Description = "Base (manual) premium")] double basePremium,
        [ExcelArgument(Description = "Array of schedule credits (negative) and debits (positive) as decimals, e.g., -0.10 for 10% credit")] double[] adjustments)
    {
        if (basePremium < 0)
            return "Error: Premium must be non-negative";
        if (adjustments == null || adjustments.Length == 0)
            return basePremium;

        // Sum all adjustments
        double totalAdjustment = adjustments.Sum();

        // Apply adjustment (additive approach)
        double factor = 1 + totalAdjustment;

        // Typically capped
        factor = Math.Max(0.5, Math.Min(1.5, factor));

        return basePremium * factor;
    }

    [ExcelFunction(Description = "Schedule rating with category limits. Each category (e.g., premises, equipment) has its own max credit/debit.", Category = "Actuarial.Credibility")]
    public static object ACT_SCHEDULE_RATING_CAPPED(
        [ExcelArgument(Description = "Base (manual) premium")] double basePremium,
        [ExcelArgument(Description = "Array of adjustments by category")] double[] adjustments,
        [ExcelArgument(Description = "Array of maximum allowed (absolute value) for each category")] double[] categoryMaximums)
    {
        if (basePremium < 0)
            return "Error: Premium must be non-negative";
        if (adjustments == null || adjustments.Length == 0)
            return basePremium;
        if (categoryMaximums == null || categoryMaximums.Length != adjustments.Length)
            return "Error: Category maximums must match adjustments length";

        double totalAdjustment = 0;
        for (int i = 0; i < adjustments.Length; i++)
        {
            double adj = adjustments[i];
            double maxAdj = Math.Abs(categoryMaximums[i]);
            double cappedAdj = Math.Max(-maxAdj, Math.Min(maxAdj, adj));
            totalAdjustment += cappedAdj;
        }

        double factor = 1 + totalAdjustment;
        return basePremium * factor;
    }

    #endregion

    #region Retrospective Rating

    [ExcelFunction(Description = "Retrospective premium: Premium = (Basic + Converted Losses) × Tax, subject to min/max. Standard retro formula.", Category = "Actuarial.Credibility")]
    public static object ACT_RETRO_PREMIUM(
        [ExcelArgument(Description = "Standard premium")] double standardPremium,
        [ExcelArgument(Description = "Actual incurred losses")] double actualLosses,
        [ExcelArgument(Description = "Loss conversion factor (LCF, typically 1.0-1.2)")] double lcf,
        [ExcelArgument(Description = "Basic premium factor (as proportion of standard)")] double basicFactor,
        [ExcelArgument(Description = "Tax multiplier (typically 1.03-1.05)")] double taxMultiplier,
        [ExcelArgument(Description = "Minimum premium factor")] double minFactor,
        [ExcelArgument(Description = "Maximum premium factor")] double maxFactor)
    {
        if (standardPremium <= 0)
            return "Error: Standard premium must be positive";
        if (lcf < 1)
            return "Error: LCF should be >= 1";
        if (minFactor < 0 || maxFactor < minFactor)
            return "Error: Min/max factors invalid";

        // Basic premium
        double basic = standardPremium * basicFactor;

        // Converted losses
        double convertedLosses = actualLosses * lcf;

        // Retro premium before min/max
        double retroPremium = (basic + convertedLosses) * taxMultiplier;

        // Apply min/max
        double minPremium = standardPremium * minFactor;
        double maxPremium = standardPremium * maxFactor;

        retroPremium = Math.Max(minPremium, Math.Min(maxPremium, retroPremium));

        return retroPremium;
    }

    [ExcelFunction(Description = "Calculate retrospective rating parameters. Returns [basic factor, LCF, ELG] given min/max and insurance charge.", Category = "Actuarial.Credibility")]
    public static object[] ACT_RETRO_PARAMETERS(
        [ExcelArgument(Description = "Expected loss ratio")] double expectedLossRatio,
        [ExcelArgument(Description = "Minimum premium factor")] double minFactor,
        [ExcelArgument(Description = "Maximum premium factor")] double maxFactor,
        [ExcelArgument(Description = "Tax multiplier")] double taxMultiplier,
        [ExcelArgument(Description = "Insurance charge (for unlimited max)")] double insuranceCharge)
    {
        if (expectedLossRatio <= 0 || expectedLossRatio > 1)
            return new object[] { "Error: Expected loss ratio should be between 0 and 1" };
        if (minFactor < 0 || maxFactor <= minFactor)
            return new object[] { "Error: Invalid min/max factors" };
        if (taxMultiplier < 1)
            return new object[] { "Error: Tax multiplier should be >= 1" };

        // Expected loss group (ELG) determines table lookup
        // Simplified calculation - in practice this uses published tables
        double elg = expectedLossRatio * 100;  // Convert to percentage points

        // Loss conversion factor - accounts for ALAE as percentage of loss
        double lcf = 1.0 + 0.20 * expectedLossRatio;  // Simplified

        // Basic premium factor
        // Basic = (Min/Tax - Insurance Savings at Min) / (1 - expected LR / LCF)
        double basicFactor = minFactor / taxMultiplier + insuranceCharge * expectedLossRatio;

        return new object[] { basicFactor, lcf, elg, insuranceCharge };
    }

    [ExcelFunction(Description = "Loss limitation premium credit. Reduces retro premium for capping individual losses.", Category = "Actuarial.Credibility")]
    public static object ACT_RETRO_LOSS_LIMIT_CREDIT(
        [ExcelArgument(Description = "Per-occurrence loss limit")] double lossLimit,
        [ExcelArgument(Description = "Expected loss ratio")] double expectedLossRatio,
        [ExcelArgument(Description = "Pareto alpha (tail parameter, > 1)")] double alpha,
        [ExcelArgument(Description = "Average claim severity")] double avgSeverity)
    {
        if (lossLimit <= 0)
            return "Error: Loss limit must be positive";
        if (expectedLossRatio <= 0 || expectedLossRatio > 1)
            return "Error: ELR should be between 0 and 1";
        if (alpha <= 1)
            return "Error: Alpha must be > 1";
        if (avgSeverity <= 0)
            return "Error: Average severity must be positive";

        // Calculate excess ratio using Pareto assumption
        double excessRatio;
        if (lossLimit > avgSeverity)
        {
            excessRatio = Math.Pow(avgSeverity / lossLimit, alpha - 1);
        }
        else
        {
            excessRatio = 1.0;
        }

        // Credit is the expected excess portion
        double credit = expectedLossRatio * excessRatio * 0.5;  // 50% of excess becomes credit (simplified)

        return credit;
    }

    #endregion

    #region Composite and Multi-Year Rating

    [ExcelFunction(Description = "Composite rate: blends experience from multiple years with declining weights.", Category = "Actuarial.Credibility")]
    public static object ACT_COMPOSITE_RATE(
        [ExcelArgument(Description = "Array of loss ratios by year (most recent first)")] double[] lossRatios,
        [ExcelArgument(Description = "Array of weights by year (most recent first). If omitted, uses declining weights.")] double[]? weights = null,
        [ExcelArgument(Description = "Decay factor for automatic weights (e.g., 0.8 means each year is 80% of prior)")] double decay = 0.8)
    {
        if (lossRatios == null || lossRatios.Length == 0)
            return "Error: Loss ratios required";

        int n = lossRatios.Length;
        double[] w;

        if (weights != null && weights.Length == n)
        {
            w = weights;
        }
        else
        {
            // Generate declining weights
            w = new double[n];
            double weight = 1.0;
            for (int i = 0; i < n; i++)
            {
                w[i] = weight;
                weight *= decay;
            }
        }

        // Normalize weights
        double totalWeight = w.Sum();
        if (totalWeight <= 0)
            return "Error: Total weight must be positive";

        // Weighted average
        double compositeRate = 0;
        for (int i = 0; i < n; i++)
        {
            compositeRate += lossRatios[i] * w[i];
        }

        return compositeRate / totalWeight;
    }

    [ExcelFunction(Description = "Multi-year credibility: accounts for correlation between years when combining experience.", Category = "Actuarial.Credibility")]
    public static object ACT_MULTIYEAR_CREDIBILITY(
        [ExcelArgument(Description = "Single-year credibility factor")] double singleYearZ,
        [ExcelArgument(Description = "Number of years of experience")] int years,
        [ExcelArgument(Description = "Year-over-year correlation (0 to 1)")] double correlation = 0)
    {
        if (singleYearZ < 0 || singleYearZ > 1)
            return "Error: Single-year credibility must be between 0 and 1";
        if (years < 1)
            return "Error: Years must be at least 1";
        if (correlation < 0 || correlation > 1)
            return "Error: Correlation must be between 0 and 1";

        // With perfect correlation (rho=1), multiple years don't add credibility
        // With no correlation (rho=0), credibility increases with sqrt(n)
        // General formula: Z_n = n * Z / (1 + (n-1) * rho)

        if (correlation >= 1)
            return singleYearZ;

        double effectiveN = years / (1 + (years - 1) * correlation);
        double multiYearZ = 1 - Math.Pow(1 - singleYearZ, effectiveN);

        return Math.Min(1.0, multiYearZ);
    }

    #endregion

    #region Alternative Credibility Formulas

    [ExcelFunction(Description = "Linear credibility: Z = min(1, n/n₀). Alternative to square root rule.", Category = "Actuarial.Credibility")]
    public static object ACT_LINEAR_CREDIBILITY(
        [ExcelArgument(Description = "Actual exposure or claims")] double n,
        [ExcelArgument(Description = "Full credibility standard")] double n0)
    {
        if (n < 0)
            return "Error: n must be non-negative";
        if (n0 <= 0)
            return "Error: n₀ must be positive";

        return Math.Min(1.0, n / n0);
    }

    [ExcelFunction(Description = "Asymptotic credibility: Z = n^p / (n^p + k) where p allows different approach rates to full credibility.", Category = "Actuarial.Credibility")]
    public static object ACT_ASYMPTOTIC_CREDIBILITY(
        [ExcelArgument(Description = "Actual exposure or claims")] double n,
        [ExcelArgument(Description = "k parameter (from Bühlmann or calibrated)")] double k,
        [ExcelArgument(Description = "Power parameter p (0.5 = sqrt rule, 1.0 = linear, 2.0 = squared)")] double p = 1.0)
    {
        if (n < 0)
            return "Error: n must be non-negative";
        if (k <= 0)
            return "Error: k must be positive";
        if (p <= 0)
            return "Error: p must be positive";

        double nPower = Math.Pow(n, p);
        return nPower / (nPower + k);
    }

    [ExcelFunction(Description = "Bayesian credibility update: combines prior with new observation using conjugate prior approach.", Category = "Actuarial.Credibility")]
    public static object ACT_BAYESIAN_UPDATE(
        [ExcelArgument(Description = "Prior mean estimate")] double priorMean,
        [ExcelArgument(Description = "Prior variance (uncertainty in prior)")] double priorVariance,
        [ExcelArgument(Description = "Observed value (new data point)")] double observed,
        [ExcelArgument(Description = "Observation variance (process variance)")] double observationVariance)
    {
        if (priorVariance <= 0)
            return "Error: Prior variance must be positive";
        if (observationVariance <= 0)
            return "Error: Observation variance must be positive";

        // Bayesian update for Normal-Normal conjugate
        // Posterior mean = weighted average of prior and observation
        // Weight determined by relative precision (1/variance)

        double priorPrecision = 1.0 / priorVariance;
        double obsPrecision = 1.0 / observationVariance;

        double posteriorPrecision = priorPrecision + obsPrecision;
        double posteriorMean = (priorPrecision * priorMean + obsPrecision * observed) / posteriorPrecision;

        return posteriorMean;
    }

    [ExcelFunction(Description = "Returns array [posterior mean, posterior variance] after Bayesian update.", Category = "Actuarial.Credibility")]
    public static object[] ACT_BAYESIAN_UPDATE_FULL(
        [ExcelArgument(Description = "Prior mean estimate")] double priorMean,
        [ExcelArgument(Description = "Prior variance")] double priorVariance,
        [ExcelArgument(Description = "Observed value")] double observed,
        [ExcelArgument(Description = "Observation variance")] double observationVariance)
    {
        if (priorVariance <= 0)
            return new object[] { "Error: Prior variance must be positive" };
        if (observationVariance <= 0)
            return new object[] { "Error: Observation variance must be positive" };

        double priorPrecision = 1.0 / priorVariance;
        double obsPrecision = 1.0 / observationVariance;

        double posteriorPrecision = priorPrecision + obsPrecision;
        double posteriorMean = (priorPrecision * priorMean + obsPrecision * observed) / posteriorPrecision;
        double posteriorVariance = 1.0 / posteriorPrecision;

        // Implied credibility given to observation
        double impliedZ = obsPrecision / posteriorPrecision;

        return new object[] { posteriorMean, posteriorVariance, impliedZ };
    }

    #endregion

    #region Credibility for Frequencies

    [ExcelFunction(Description = "Credibility for Poisson frequencies: Z = n / (n + E[N]/Var[λ]) where n is observed claims.", Category = "Actuarial.Credibility")]
    public static object ACT_CREDIBILITY_POISSON(
        [ExcelArgument(Description = "Observed number of claims")] double observedClaims,
        [ExcelArgument(Description = "Prior expected frequency")] double priorFrequency,
        [ExcelArgument(Description = "Prior variance of frequency (across risks)")] double priorVariance)
    {
        if (observedClaims < 0)
            return "Error: Observed claims must be non-negative";
        if (priorFrequency <= 0)
            return "Error: Prior frequency must be positive";
        if (priorVariance <= 0)
            return "Error: Prior variance must be positive";

        // For Poisson-Gamma model, k = E[λ] / Var[λ]
        double k = priorFrequency / priorVariance;

        return observedClaims / (observedClaims + k);
    }

    [ExcelFunction(Description = "Credibility-weighted frequency estimate using Poisson-Gamma model.", Category = "Actuarial.Credibility")]
    public static object ACT_FREQUENCY_ESTIMATE(
        [ExcelArgument(Description = "Observed number of claims")] double observedClaims,
        [ExcelArgument(Description = "Exposure units (e.g., policy years)")] double exposure,
        [ExcelArgument(Description = "Prior expected frequency per exposure")] double priorFrequency,
        [ExcelArgument(Description = "Prior variance of frequency")] double priorVariance)
    {
        if (observedClaims < 0)
            return "Error: Observed claims must be non-negative";
        if (exposure <= 0)
            return "Error: Exposure must be positive";
        if (priorFrequency <= 0)
            return "Error: Prior frequency must be positive";
        if (priorVariance <= 0)
            return "Error: Prior variance must be positive";

        // Observed frequency
        double obsFreq = observedClaims / exposure;

        // Bühlmann credibility for this setup
        double k = priorFrequency / priorVariance;
        double z = observedClaims / (observedClaims + k);

        // Credibility-weighted estimate
        return z * obsFreq + (1 - z) * priorFrequency;
    }

    #endregion
}
