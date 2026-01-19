using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using System.Linq;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Chain ladder and reserving methods for general insurance loss reserving.
/// References:
/// - Mack, T. (1993). "Distribution-free calculation of the standard error of chain ladder reserve estimates." ASTIN Bulletin 23(2): 213-225.
/// - Mack, T. (1999). "The standard error of chain ladder reserve estimates: Recursive calculation and inclusion of a tail factor." ASTIN Bulletin 29(2): 361-366.
/// - England, P.D. and Verrall, R.J. (2002). "Stochastic claims reserving in general insurance." British Actuarial Journal 8(3): 443-518.
/// - Bornhuetter, R.L. and Ferguson, R.E. (1972). "The actuary and IBNR." Proceedings of the CAS, LIX: 181-195.
/// - Berquist, J.R. and Sherman, R.E. (1977). "Loss reserve adequacy testing: A comprehensive, systematic approach." Proceedings of the CAS, LXIV: 123-184.
/// </summary>
public static class ChainLadder
{
    #region Basic Chain Ladder

    [ExcelFunction(Description = "Calculate volume-weighted chain ladder development factors from a cumulative triangle. Standard reserving method per Mack (1993).", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_FACTORS(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Use only top N (oldest) years; default is all")] double topNYears = double.PositiveInfinity,
        [ExcelArgument(Description = "Exclude most recent year")] bool excludeMostRecentYear = false,
        [ExcelArgument(Description = "Exclude highest and lowest ratios")] bool excludeHighLowRatios = false)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        int maxYears = int.MaxValue;
        if (!double.IsInfinity(topNYears) && topNYears > 0)
            maxYears = (int)Math.Floor(topNYears);

        var factors = new object[n - 1];

        for (int j = 0; j < n - 1; j++)
        {
            int count = n - j - 1;
            var indices = Enumerable.Range(0, count).ToList();
            if (excludeMostRecentYear && indices.Count > 0)
                indices.RemoveAt(indices.Count - 1);

            if (maxYears != int.MaxValue && indices.Count > maxYears)
                indices = indices.Take(maxYears).ToList();

            var ratios = new List<(double ratio, double current, double next)>();

            foreach (int i in indices)
            {
                if (triangle[i, j] > 0 && triangle[i, j + 1] > 0)
                {
                    double current = triangle[i, j];
                    double next = triangle[i, j + 1];
                    ratios.Add((next / current, current, next));
                }
            }

            if (excludeHighLowRatios && ratios.Count > 2)
                ratios = ratios.OrderBy(r => r.ratio).Skip(1).Take(ratios.Count - 2).ToList();

            double sumCurrent = ratios.Sum(r => r.current);
            double sumNext = ratios.Sum(r => r.next);
            factors[j] = sumCurrent > 0 ? sumNext / sumCurrent : 1.0;
        }

        return factors;
    }

    [ExcelFunction(Description = "Return the latest diagonal of a cumulative triangle", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_LATEST(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        var latest = new object[n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            latest[i] = triangle[i, lastCol];
        }

        return latest;
    }

    [ExcelFunction(Description = "Project ultimate losses using chain ladder", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_ULTIMATE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        // Get development factors
        var factorsObj = ACT_CL_FACTORS(triangle);
        if (factorsObj[0] is string)
            return factorsObj;

        var factors = factorsObj.Cast<double>().ToArray();

        // Calculate cumulative factors to ultimate
        var cdfToUltimate = new double[n];
        cdfToUltimate[n - 1] = 1.0;
        for (int j = n - 2; j >= 0; j--)
        {
            cdfToUltimate[j] = cdfToUltimate[j + 1] * factors[j];
        }

        // Project ultimates
        var ultimates = new object[n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double latestValue = triangle[i, lastCol];
            ultimates[i] = latestValue * cdfToUltimate[lastCol];
        }

        return ultimates;
    }

    [ExcelFunction(Description = "Calculate IBNR reserves using chain ladder", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_IBNR(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        var ultimatesObj = ACT_CL_ULTIMATE(triangle);
        if (ultimatesObj[0] is string)
            return ultimatesObj;

        var ibnr = new object[n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double latestValue = triangle[i, lastCol];
            ibnr[i] = (double)ultimatesObj[i] - latestValue;
        }

        return ibnr;
    }

    [ExcelFunction(Description = "Bornhuetter-Ferguson ultimate: Ult = Paid + Unreported%. Balances actual experience with a priori expectation. Ref: Bornhuetter & Ferguson (1972). Use when data is immature or volatile.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_BF_ULTIMATE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Development factors (n-1)")] double[] developmentFactors,
        [ExcelArgument(Description = "A priori ultimate (n or single value)")] double[] aPrioriUltimate)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        if (developmentFactors.Length < n - 1)
            return new object[] { "Error: Development factors length must be n-1" };

        var apriori = new double[n];
        if (aPrioriUltimate.Length == 1)
        {
            for (int i = 0; i < n; i++)
                apriori[i] = aPrioriUltimate[0];
        }
        else if (aPrioriUltimate.Length == n)
        {
            for (int i = 0; i < n; i++)
                apriori[i] = aPrioriUltimate[i];
        }
        else
        {
            return new object[] { "Error: A priori ultimate must be length 1 or n" };
        }

        // Calculate cumulative factors to ultimate
        var cdfToUltimate = new double[n];
        cdfToUltimate[n - 1] = 1.0;
        for (int j = n - 2; j >= 0; j--)
        {
            cdfToUltimate[j] = cdfToUltimate[j + 1] * developmentFactors[j];
        }

        var ultimates = new object[n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double latestValue = triangle[i, lastCol];
            double cdf = cdfToUltimate[lastCol];
            if (cdf <= 0)
                ultimates[i] = double.NaN;
            else
                ultimates[i] = latestValue + apriori[i] * (1.0 - 1.0 / cdf);
        }

        return ultimates;
    }

    #endregion

    #region Mack Chain Ladder

    [ExcelFunction(Description = "Mack chain ladder standard errors for development factors. Distribution-free method per Mack (1993). Measures uncertainty in LDF estimates.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_MACK_FACTOR_SE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        var factorsObj = ACT_CL_FACTORS(triangle);
        if (factorsObj[0] is string)
            return factorsObj;

        var factors = factorsObj.Cast<double>().ToArray();
        var sigmaSquared = new double[n - 1];
        var sumWeights = new double[n - 1];
        var factorSE = new object[n - 1];

        // Calculate sigma^2 (variance of individual factors)
        for (int j = 0; j < n - 1; j++)
        {
            double sumWeightedVar = 0;
            int count = n - j - 1;

            for (int i = 0; i < count; i++)
            {
                if (triangle[i, j] > 0 && triangle[i, j + 1] > 0)
                {
                    double individualFactor = triangle[i, j + 1] / triangle[i, j];
                    sumWeightedVar += triangle[i, j] * Math.Pow(individualFactor - factors[j], 2);
                    sumWeights[j] += triangle[i, j];
                }
            }

            sigmaSquared[j] = count > 1 ? sumWeightedVar / (count - 1) : 0;
        }

        AdjustSigmaSquaredTail(sigmaSquared);

        for (int j = 0; j < n - 1; j++)
        {
            factorSE[j] = sumWeights[j] > 0 ? Math.Sqrt(sigmaSquared[j] / sumWeights[j]) : 0;
        }

        return factorSE;
    }

    [ExcelFunction(Description = "Mack chain ladder reserve standard errors by origin year. Per Mack (1993, 1999). Use for reserve range and risk margin calculations.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_MACK_RESERVE_SE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        var factorsObj = ACT_CL_FACTORS(triangle);
        var ultimatesObj = ACT_CL_ULTIMATE(triangle);
        var factorSEObj = ACT_MACK_FACTOR_SE(triangle);

        if (factorsObj[0] is string || ultimatesObj[0] is string || factorSEObj[0] is string)
            return new object[] { "Error in calculation" };

        var factors = factorsObj.Cast<double>().ToArray();
        var ultimates = ultimatesObj.Cast<double>().ToArray();
        var factorSE = factorSEObj.Cast<double>().ToArray();

        // Calculate sigma^2 for each development period
        var sigmaSquared = new double[n - 1];
        var sumC = new double[n - 1];
        for (int j = 0; j < n - 1; j++)
        {
            sumC[j] = GetSumC(triangle, j);
            sigmaSquared[j] = Math.Pow(factorSE[j], 2) * sumC[j];
        }
        AdjustSigmaSquaredTail(sigmaSquared);

        // Calculate reserve standard errors using Mack formula
        var reserveSE = new object[n];
        reserveSE[0] = 0.0; // First period has no uncertainty

        for (int i = 1; i < n; i++)
        {
            double variance = 0;
            int lastCol = n - 1 - i;
            for (int j = lastCol; j < n - 1; j++)
            {
                double Cij = GetProjectedC(triangle, factors, i, j);
                if (Cij <= 0 || sumC[j] <= 0)
                    continue;
                double term = sigmaSquared[j] / (factors[j] * factors[j]);
                variance += term * (1.0 / Cij + 1.0 / sumC[j]);
            }

            reserveSE[i] = Math.Sqrt(variance) * ultimates[i];
        }

        return reserveSE;
    }

    private static double GetSumC(double[,] triangle, int col)
    {
        int n = triangle.GetLength(0);
        double sum = 0;
        for (int i = 0; i < n - col - 1; i++)
        {
            if (triangle[i, col] > 0)
                sum += triangle[i, col];
        }
        return sum;
    }

    private static double GetProjectedC(double[,] triangle, double[] factors, int row, int col)
    {
        int n = triangle.GetLength(0);
        int lastKnownCol = n - 1 - row;

        if (col <= lastKnownCol)
            return triangle[row, col];

        double value = triangle[row, lastKnownCol];
        for (int j = lastKnownCol; j < col; j++)
        {
            value *= factors[j];
        }
        return value;
    }

    private static void AdjustSigmaSquaredTail(double[] sigmaSquared)
    {
        if (sigmaSquared.Length < 3)
            return;

        int last = sigmaSquared.Length - 1;
        if (sigmaSquared[last] <= 0)
        {
            double prev1 = sigmaSquared[last - 1];
            double prev2 = sigmaSquared[last - 2];
            if (prev1 > 0 && prev2 > 0)
                sigmaSquared[last] = Math.Min(prev1, prev2);
        }
    }

    private static double[] CalculateFactors(double[,] triangle)
    {
        int n = triangle.GetLength(0);
        var factors = new double[n - 1];
        for (int j = 0; j < n - 1; j++)
        {
            double sumCurrent = 0;
            double sumNext = 0;
            for (int i = 0; i < n - j - 1; i++)
            {
                if (triangle[i, j] > 0 && triangle[i, j + 1] > 0)
                {
                    sumCurrent += triangle[i, j];
                    sumNext += triangle[i, j + 1];
                }
            }
            factors[j] = sumCurrent > 0 ? sumNext / sumCurrent : 1.0;
        }
        return factors;
    }

    private static bool TryGetBootstrapInputs(
        double[,] triangle,
        double[] factors,
        out double[,] fittedIncremental,
        out double phi,
        out double[] standardizedResiduals,
        out string error)
    {
        int n = triangle.GetLength(0);
        fittedIncremental = new double[n, n];
        phi = 1.0;
        standardizedResiduals = Array.Empty<double>();
        error = "";

        var residuals = new List<double>();
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n - i; j++)
            {
                double current = triangle[i, j];
                double incremental = j == 0 ? current : current - triangle[i, j - 1];
                
                // For j=0, fitted value is the actual value (no development)
                // For j>0, fitted value is previous cumulative * (factor - 1)
                double mean = j == 0 ? current : triangle[i, j - 1] * (factors[j - 1] - 1.0);
                fittedIncremental[i, j] = mean;

                // Only calculate residuals for j > 0 (development periods)
                // j=0 has no development factor applied, so no residual
                if (j > 0 && mean > 0)
                {
                    double residual = (incremental - mean) / Math.Sqrt(mean);
                    residuals.Add(residual);
                }
            }
        }

        if (residuals.Count <= n - 1)
        {
            error = "Error: Not enough residuals to bootstrap";
            return false;
        }

        // Degrees of freedom = number of residuals - number of parameters (n-1 factors)
        int df = residuals.Count - (n - 1);
        phi = residuals.Sum(r => r * r) / df;
        if (phi <= 0)
            phi = 1.0;

        // Adjust residuals for degrees of freedom bias per England & Verrall (2002)
        // r_adj = r * sqrt(n / (n-p)) where n = number of residuals, p = number of parameters
        // NOTE: We keep raw Pearson residuals (NOT divided by sqrt(phi))
        // because the bootstrap formula is: pseudo = fitted + r_adj * sqrt(fitted)
        double adjFactor = Math.Sqrt((double)residuals.Count / df);
        standardizedResiduals = residuals.Select(r => r * adjFactor).ToArray();
        return true;
    }

    private static double[,] BuildBootstrapTriangle(
        double[,] triangle,
        double[,] fittedIncremental,
        double phi,
        double[] standardizedResiduals,
        Random random)
    {
        int n = triangle.GetLength(0);
        var bootIncremental = new double[n, n];
        
        // Build bootstrap incremental triangle (pseudo-data)
        // Per England & Verrall (2002): pseudo = fitted + residual * sqrt(fitted)
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n - i; j++)
            {
                if (j == 0)
                {
                    // First column: use original values (no development to perturb)
                    bootIncremental[i, j] = fittedIncremental[i, j];
                }
                else
                {
                    // Development periods: perturb with resampled residuals
                    double mean = fittedIncremental[i, j];
                    if (mean > 0)
                    {
                        double residual = standardizedResiduals[random.Next(standardizedResiduals.Length)];
                        double pseudoValue = mean + residual * Math.Sqrt(mean);
                        // Ensure non-negative (E&V methodology allows small positive floor)
                        bootIncremental[i, j] = Math.Max(1.0, pseudoValue);
                    }
                    else
                    {
                        bootIncremental[i, j] = Math.Max(1.0, mean);
                    }
                }
            }
        }

        // Convert pseudo-incremental to pseudo-cumulative
        var pseudoTriangle = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            double cumulative = 0;
            for (int j = 0; j < n - i; j++)
            {
                cumulative += bootIncremental[i, j];
                pseudoTriangle[i, j] = cumulative;
            }
        }

        // Re-estimate factors from pseudo-triangle
        var bootFactors = CalculateFactors(pseudoTriangle);

        // Project ORIGINAL triangle's latest diagonal using re-estimated factors + process variance
        // This is the key E&V insight: pseudo-triangle is only used to estimate parameter uncertainty
        var bootTriangle = new double[n, n];
        
        // Copy original observed values
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j <= n - 1 - i; j++)
            {
                bootTriangle[i, j] = triangle[i, j];
            }
        }
        
        // Project future values starting from ORIGINAL latest diagonal
        for (int i = 1; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double current = triangle[i, lastCol];  // Start from ORIGINAL latest
            
            for (int j = lastCol; j < n - 1; j++)
            {
                double meanNext = current * bootFactors[j];
                double meanIncremental = meanNext - current;
                
                if (meanIncremental > 0)
                {
                    // Gamma distribution for process variance
                    // E[X] = mean, Var[X] = mean * phi
                    // MathNet Gamma.Sample uses (shape, RATE) where rate = 1/scale
                    double shape = meanIncremental / phi;
                    double rate = 1.0 / phi;  // Convert scale to rate
                    double simulatedIncremental = Gamma.Sample(random, shape, rate);
                    current = current + simulatedIncremental;
                }
                else
                {
                    current = meanNext;
                }
                bootTriangle[i, j + 1] = current;
            }
        }

        return bootTriangle;
    }

    #endregion

    #region Bootstrap Methods

    [ExcelFunction(Description = "Bootstrap chain ladder reserves - returns percentiles. Stochastic reserving method per England & Verrall (2002). Generates full distribution of reserve outcomes.", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_CL_BOOTSTRAP(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations")] int iterations,
        [ExcelArgument(Description = "Random seed (0 for random)")] int seed = 0)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        if (iterations <= 0)
            return new object[,] { { "Error: Iterations must be positive" } };

        var random = seed == 0 ? new Random() : new Random(seed);

        // Get original factors
        var factorsObj = ACT_CL_FACTORS(triangle);
        if (factorsObj[0] is string)
            return new object[,] { { factorsObj[0] } };

        var factors = factorsObj.Cast<double>().ToArray();

        if (!TryGetBootstrapInputs(triangle, factors, out var fittedIncremental, out var phi, out var standardizedResiduals, out var error))
            return new object[,] { { error } };

        // Bootstrap iterations
        var bootstrapReserves = new double[iterations];
        for (int iter = 0; iter < iterations; iter++)
        {
            var bootTriangle = BuildBootstrapTriangle(triangle, fittedIncremental, phi, standardizedResiduals, random);

            // Calculate total reserve for this iteration
            // Note: i=0 (year 1) is fully developed, contributes 0 reserve
            double totalReserve = 0;
            for (int i = 1; i < n; i++)
            {
                int lastCol = n - 1 - i;
                double latest = triangle[i, lastCol];
                totalReserve += bootTriangle[i, n - 1] - latest;
            }
            bootstrapReserves[iter] = totalReserve;
        }

        // Sort and return percentiles
        Array.Sort(bootstrapReserves);

        var percentiles = new double[] { 0.01, 0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99 };
        var result = new object[percentiles.Length + 2, 2];

        result[0, 0] = "Mean";
        double mean = bootstrapReserves.Average();
        result[0, 1] = mean;
        result[1, 0] = "StdDev";
        result[1, 1] = Math.Sqrt(bootstrapReserves.Select(x => Math.Pow(x - mean, 2)).Average());

        for (int i = 0; i < percentiles.Length; i++)
        {
            result[i + 2, 0] = $"P{percentiles[i] * 100:0}";
            int idx = Math.Min((int)(percentiles[i] * iterations), iterations - 1);
            result[i + 2, 1] = bootstrapReserves[idx];
        }

        return result;
    }

    [ExcelFunction(Description = "Bootstrap chain ladder reserves by origin year - returns statistics", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_CL_BOOTSTRAP_ORIGIN(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations")] int iterations,
        [ExcelArgument(Description = "Random seed (0 for random)")] int seed = 0)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        if (iterations <= 0)
            return new object[,] { { "Error: Iterations must be positive" } };

        var random = seed == 0 ? new Random() : new Random(seed);

        var factorsObj = ACT_CL_FACTORS(triangle);
        if (factorsObj[0] is string)
            return new object[,] { { factorsObj[0] } };

        var factors = factorsObj.Cast<double>().ToArray();
        if (!TryGetBootstrapInputs(triangle, factors, out var fittedIncremental, out var phi, out var standardizedResiduals, out var error))
            return new object[,] { { error } };

        var reservesByOrigin = new double[n][];
        for (int i = 0; i < n; i++)
            reservesByOrigin[i] = new double[iterations];

        for (int iter = 0; iter < iterations; iter++)
        {
            var bootTriangle = BuildBootstrapTriangle(triangle, fittedIncremental, phi, standardizedResiduals, random);

            for (int i = 0; i < n; i++)
            {
                int lastCol = n - 1 - i;
                if (lastCol == n - 1)
                {
                    // Fully developed row - no future projection, reserve is exactly 0
                    reservesByOrigin[i][iter] = 0;
                }
                else
                {
                    double latest = triangle[i, lastCol];
                    reservesByOrigin[i][iter] = bootTriangle[i, n - 1] - latest;
                }
            }
        }

        var percentiles = new double[] { 0.50, 0.75, 0.90, 0.95, 0.99 };
        var result = new object[n + 1, percentiles.Length + 3];

        result[0, 0] = "AY";
        result[0, 1] = "Mean";
        result[0, 2] = "StdDev";
        for (int p = 0; p < percentiles.Length; p++)
            result[0, p + 3] = $"P{percentiles[p] * 100:0}";

        for (int i = 0; i < n; i++)
        {
            var data = reservesByOrigin[i];
            Array.Sort(data);
            double mean = data.Average();
            double stdDev = Math.Sqrt(data.Select(x => Math.Pow(x - mean, 2)).Average());

            result[i + 1, 0] = i + 1;
            result[i + 1, 1] = mean;
            result[i + 1, 2] = stdDev;

            for (int p = 0; p < percentiles.Length; p++)
            {
                int idx = Math.Min((int)(percentiles[p] * iterations), iterations - 1);
                result[i + 1, p + 3] = data[idx];
            }
        }

        return result;
    }

    #endregion

    #region Berquist-Sherman Adjustment

    [ExcelFunction(Description = "Berquist-Sherman paid loss adjustment for case reserve adequacy changes. Ref: Berquist & Sherman (1977). Restates historical paids to current adequacy level.", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_BERQUIST_SHERMAN(
        [ExcelArgument(Description = "Paid loss triangle (n x n cumulative values)")] double[,] paidTriangle,
        [ExcelArgument(Description = "Case reserve triangle (n x n values)")] double[,] caseTriangle,
        [ExcelArgument(Description = "Trend rate for case reserve adequacy (e.g., 0.05 for 5%)")] double trendRate)
    {
        int n = paidTriangle.GetLength(0);
        if (paidTriangle.GetLength(1) != n || caseTriangle.GetLength(0) != n || caseTriangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangles must be square and same size" } };

        // Calculate incurred triangle
        var incurredTriangle = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j <= n - 1 - i; j++)
            {
                incurredTriangle[i, j] = paidTriangle[i, j] + caseTriangle[i, j];
            }
        }

        // Adjust case reserves to common adequacy level (latest diagonal)
        var adjustedCaseTriangle = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j <= n - 1 - i; j++)
            {
                int periodsToAdjust = (n - 1 - i) - j;
                double adjustmentFactor = Math.Pow(1 + trendRate, periodsToAdjust);
                adjustedCaseTriangle[i, j] = caseTriangle[i, j] * adjustmentFactor;
            }
        }

        // Calculate adjusted paid triangle (maintain incurred, adjust paid)
        var adjustedPaidTriangle = new object[n, n];
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j <= n - 1 - i; j++)
            {
                double adjustedIncurred = paidTriangle[i, j] + adjustedCaseTriangle[i, j];
                double originalIncurred = incurredTriangle[i, j];

                // Adjusted paid = original paid * (adjusted incurred / original incurred)
                if (originalIncurred > 0)
                {
                    adjustedPaidTriangle[i, j] = paidTriangle[i, j] * (adjustedIncurred / originalIncurred);
                }
                else
                {
                    adjustedPaidTriangle[i, j] = paidTriangle[i, j];
                }
            }
            // Fill remaining with empty
            for (int j = n - i; j < n; j++)
            {
                adjustedPaidTriangle[i, j] = "";
            }
        }

        return adjustedPaidTriangle;
    }

    #endregion

    #region Cape Cod Method

    [ExcelFunction(Description = "Cape Cod ultimate: Alternative to B-F that estimates ELR from the data itself. Ult = Paid + ELR * OnLevelPremium * UnreportedPct. Good when a priori ELR is uncertain.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CAPECOD_ULTIMATE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "On-level earned premium by origin year (n values)")] double[] premium,
        [ExcelArgument(Description = "Development factors (n-1). If omitted, calculated from triangle.")] double[] developmentFactors = null)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        if (premium == null || premium.Length != n)
            return new object[] { "Error: Premium array must have n values" };

        // Get development factors if not provided
        double[] factors;
        if (developmentFactors == null || developmentFactors.Length == 0)
        {
            var factorsObj = ACT_CL_FACTORS(triangle);
            if (factorsObj[0] is string)
                return factorsObj;
            factors = factorsObj.Cast<double>().ToArray();
        }
        else
        {
            if (developmentFactors.Length < n - 1)
                return new object[] { "Error: Development factors must have n-1 values" };
            factors = developmentFactors;
        }

        // Calculate cumulative factors to ultimate (% reported)
        var pctReported = new double[n];
        double cumulativeFactor = 1.0;
        for (int j = n - 2; j >= 0; j--)
        {
            cumulativeFactor *= factors[j];
        }
        pctReported[0] = 1.0 / cumulativeFactor;

        cumulativeFactor = 1.0;
        for (int j = n - 2; j >= 0; j--)
        {
            cumulativeFactor *= factors[j];
            int col = n - 1 - j - 1;
            if (col >= 0 && col < n)
                pctReported[n - 1 - j] = 1.0 / cumulativeFactor;
        }
        pctReported[n - 1] = 1.0;

        // Calculate Cape Cod ELR
        double sumUsedUp = 0;
        double sumLatest = 0;
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double latest = triangle[i, lastCol];
            double pct = pctReported[lastCol];
            sumLatest += latest;
            sumUsedUp += premium[i] * pct;
        }

        if (sumUsedUp <= 0)
            return new object[] { "Error: Used-up premium is zero" };

        double elr = sumLatest / sumUsedUp;

        // Calculate ultimates
        var ultimates = new object[n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double latest = triangle[i, lastCol];
            double pct = pctReported[lastCol];
            double unreported = 1.0 - pct;
            ultimates[i] = latest + elr * premium[i] * unreported;
        }

        return ultimates;
    }

    [ExcelFunction(Description = "Cape Cod Expected Loss Ratio: Estimates ELR from triangle and premium data. ELR = ΣLatest / Σ(Premium × PctReported).", Category = "Actuarial.ChainLadder")]
    public static object ACT_CAPECOD_ELR(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "On-level earned premium by origin year (n values)")] double[] premium,
        [ExcelArgument(Description = "Development factors (n-1). If omitted, calculated from triangle.")] double[] developmentFactors = null)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return "Error: Triangle must be square";

        if (premium == null || premium.Length != n)
            return "Error: Premium array must have n values";

        // Get development factors if not provided
        double[] factors;
        if (developmentFactors == null || developmentFactors.Length == 0)
        {
            var factorsObj = ACT_CL_FACTORS(triangle);
            if (factorsObj[0] is string)
                return factorsObj[0];
            factors = factorsObj.Cast<double>().ToArray();
        }
        else
        {
            if (developmentFactors.Length < n - 1)
                return "Error: Development factors must have n-1 values";
            factors = developmentFactors;
        }

        // Calculate cumulative factors to ultimate
        var cdfToUltimate = new double[n];
        cdfToUltimate[n - 1] = 1.0;
        for (int j = n - 2; j >= 0; j--)
        {
            cdfToUltimate[j] = cdfToUltimate[j + 1] * factors[j];
        }

        // Calculate Cape Cod ELR
        double sumUsedUp = 0;
        double sumLatest = 0;
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double latest = triangle[i, lastCol];
            double pctReported = 1.0 / cdfToUltimate[lastCol];
            sumLatest += latest;
            sumUsedUp += premium[i] * pctReported;
        }

        if (sumUsedUp <= 0)
            return "Error: Used-up premium is zero";

        return sumLatest / sumUsedUp;
    }

    #endregion

    #region Triangle Utilities

    [ExcelFunction(Description = "Convert cumulative triangle to incremental triangle. Incremental[i,j] = Cumulative[i,j] - Cumulative[i,j-1].", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_TRIANGLE_TO_INCREMENTAL(
        [ExcelArgument(Description = "Cumulative triangle data (n x n)")] double[,] cumulativeTriangle)
    {
        int n = cumulativeTriangle.GetLength(0);
        if (cumulativeTriangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        var incremental = new object[n, n];

        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j <= n - 1 - i; j++)
            {
                if (j == 0)
                    incremental[i, j] = cumulativeTriangle[i, j];
                else
                    incremental[i, j] = cumulativeTriangle[i, j] - cumulativeTriangle[i, j - 1];
            }
            // Fill remaining with empty
            for (int j = n - i; j < n; j++)
            {
                incremental[i, j] = "";
            }
        }

        return incremental;
    }

    [ExcelFunction(Description = "Convert incremental triangle to cumulative triangle. Cumulative[i,j] = Σ Incremental[i,0..j].", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_INCREMENTAL_TO_CUMULATIVE(
        [ExcelArgument(Description = "Incremental triangle data (n x n)")] double[,] incrementalTriangle)
    {
        int n = incrementalTriangle.GetLength(0);
        if (incrementalTriangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        var cumulative = new object[n, n];

        for (int i = 0; i < n; i++)
        {
            double sum = 0;
            for (int j = 0; j <= n - 1 - i; j++)
            {
                sum += incrementalTriangle[i, j];
                cumulative[i, j] = sum;
            }
            // Fill remaining with empty
            for (int j = n - i; j < n; j++)
            {
                cumulative[i, j] = "";
            }
        }

        return cumulative;
    }

    [ExcelFunction(Description = "Extract the latest diagonal from a triangle (most recent evaluation for each origin year).", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_TRIANGLE_DIAGONAL(
        [ExcelArgument(Description = "Triangle data (n x n)")] double[,] triangle,
        [ExcelArgument(Description = "Diagonal offset (0=latest, 1=one period prior, etc.)")] int offset = 0)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        if (offset < 0 || offset >= n)
            return new object[] { "Error: Offset must be between 0 and n-1" };

        var diagonal = new object[n];
        for (int i = 0; i < n; i++)
        {
            int col = n - 1 - i - offset;
            if (col >= 0)
                diagonal[i] = triangle[i, col];
            else
                diagonal[i] = "";
        }

        return diagonal;
    }

    [ExcelFunction(Description = "Calculate age-to-age factors (link ratios) for each cell in the triangle.", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_TRIANGLE_LINK_RATIOS(
        [ExcelArgument(Description = "Cumulative triangle data (n x n)")] double[,] triangle)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        var linkRatios = new object[n, n - 1];

        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n - 1 - i; j++)
            {
                if (triangle[i, j] > 0)
                    linkRatios[i, j] = triangle[i, j + 1] / triangle[i, j];
                else
                    linkRatios[i, j] = "";
            }
            // Fill remaining with empty
            for (int j = n - 1 - i; j < n - 1; j++)
            {
                linkRatios[i, j] = "";
            }
        }

        return linkRatios;
    }

    #endregion

    #region Calendar Year Adjustments

    [ExcelFunction(Description = "Adjust triangle for calendar year inflation/trend. Restates all values to latest calendar year dollars.", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_CL_CALENDAR_ADJUST(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Annual trend rate (e.g., 0.03 for 3%)")] double trendRate)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        var adjusted = new object[n, n];

        // Calendar year = origin year + development period
        // Latest calendar year = (n-1) + 0 = n-1 (for origin year n-1, dev period 0)
        // Actually latest calendar year is n-1 (0-indexed)
        int latestCalendarYear = n - 1;

        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j <= n - 1 - i; j++)
            {
                int calendarYear = i + j;
                int yearsToAdjust = latestCalendarYear - calendarYear;
                double adjustmentFactor = Math.Pow(1 + trendRate, yearsToAdjust);
                adjusted[i, j] = triangle[i, j] * adjustmentFactor;
            }
            // Fill remaining with empty
            for (int j = n - i; j < n; j++)
            {
                adjusted[i, j] = "";
            }
        }

        return adjusted;
    }

    [ExcelFunction(Description = "Calculate calendar year totals from a triangle (sum along diagonals).", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_CALENDAR_TOTALS(
        [ExcelArgument(Description = "Incremental triangle data (n x n)")] double[,] incrementalTriangle)
    {
        int n = incrementalTriangle.GetLength(0);
        if (incrementalTriangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        // Calendar years go from 0 to 2*(n-1), but only 0 to n-1 have data
        var calendarTotals = new object[n];

        for (int cy = 0; cy < n; cy++)
        {
            double sum = 0;
            // Calendar year cy = origin year i + development period j
            // So for each cy, iterate over valid (i, j) pairs where i + j = cy
            for (int i = 0; i <= cy; i++)
            {
                int j = cy - i;
                if (i < n && j < n && j <= n - 1 - i)
                {
                    sum += incrementalTriangle[i, j];
                }
            }
            calendarTotals[cy] = sum;
        }

        return calendarTotals;
    }

    [ExcelFunction(Description = "Weighted average of multiple ultimate estimates. Weights should sum to 1.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_WEIGHTED_AVERAGE(
        [ExcelArgument(Description = "First set of ultimates (n values)")] double[] ultimates1,
        [ExcelArgument(Description = "Weight for first set (0 to 1)")] double weight1,
        [ExcelArgument(Description = "Second set of ultimates (n values)")] double[] ultimates2,
        [ExcelArgument(Description = "Weight for second set (0 to 1)")] double weight2,
        [ExcelArgument(Description = "Third set of ultimates (optional)")] double[] ultimates3 = null,
        [ExcelArgument(Description = "Weight for third set (optional)")] double weight3 = 0)
    {
        if (ultimates1 == null || ultimates2 == null)
            return new object[] { "Error: Ultimates arrays required" };

        int n = ultimates1.Length;
        if (ultimates2.Length != n)
            return new object[] { "Error: All ultimate arrays must be same length" };

        if (ultimates3 != null && ultimates3.Length != n)
            return new object[] { "Error: All ultimate arrays must be same length" };

        double totalWeight = weight1 + weight2 + weight3;
        if (Math.Abs(totalWeight - 1.0) > 0.01)
            return new object[] { $"Warning: Weights sum to {totalWeight:F2}, not 1.0" };

        var result = new object[n];
        for (int i = 0; i < n; i++)
        {
            double weighted = ultimates1[i] * weight1 + ultimates2[i] * weight2;
            if (ultimates3 != null)
                weighted += ultimates3[i] * weight3;
            result[i] = weighted;
        }

        return result;
    }

    #endregion
}
