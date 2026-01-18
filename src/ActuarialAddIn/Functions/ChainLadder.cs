using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using System.Linq;

namespace ActuarialAddIn.Functions;

public static class ChainLadder
{
    #region Basic Chain Ladder

    [ExcelFunction(Description = "Calculate chain ladder development factors", Category = "Actuarial.ChainLadder")]
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

    [ExcelFunction(Description = "Bornhuetter-Ferguson ultimate using development factors and a priori ultimates", Category = "Actuarial.ChainLadder")]
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

    [ExcelFunction(Description = "Mack chain ladder standard errors for development factors", Category = "Actuarial.ChainLadder")]
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

    [ExcelFunction(Description = "Mack chain ladder reserve standard errors", Category = "Actuarial.ChainLadder")]
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
                double mean = j == 0 ? current : triangle[i, j - 1] * (factors[j - 1] - 1.0);
                fittedIncremental[i, j] = mean;

                if (mean > 0)
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

        phi = residuals.Sum(r => r * r) / (residuals.Count - (n - 1));
        if (phi <= 0)
            phi = 1.0;

        double phiSqrt = Math.Sqrt(phi);
        standardizedResiduals = residuals.Select(r => r / phiSqrt).ToArray();
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
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n - i; j++)
            {
                double mean = fittedIncremental[i, j];
                if (mean > 0)
                {
                    double residual = standardizedResiduals[random.Next(standardizedResiduals.Length)];
                    bootIncremental[i, j] = mean + residual * Math.Sqrt(mean * phi);
                }
            }
        }

        var bootTriangle = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            double cumulative = 0;
            for (int j = 0; j < n - i; j++)
            {
                cumulative += bootIncremental[i, j];
                bootTriangle[i, j] = cumulative;
            }
        }

        var bootFactors = CalculateFactors(bootTriangle);

        for (int i = 1; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = lastCol; j < n - 1; j++)
            {
                double current = bootTriangle[i, j];
                double expected = current * bootFactors[j];
                double meanIncremental = expected - current;
                if (meanIncremental > 0)
                {
                    double shape = meanIncremental / phi;
                    double scale = phi;
                    double simulatedIncremental = Gamma.Sample(random, shape, scale);
                    bootTriangle[i, j + 1] = current + simulatedIncremental;
                }
                else
                {
                    bootTriangle[i, j + 1] = current;
                }
            }
        }

        return bootTriangle;
    }

    #endregion

    #region Bootstrap Methods

    [ExcelFunction(Description = "Bootstrap chain ladder reserves - returns percentiles", Category = "Actuarial.ChainLadder")]
    public static object[,] ACT_BOOTSTRAP_CL(
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
            double totalReserve = 0;
            for (int i = 1; i < n; i++)
            {
                totalReserve += bootTriangle[i, n - 1] - triangle[i, n - 1 - i];
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
    public static object[,] ACT_BOOTSTRAP_CL_ORIGIN(
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
                double latest = triangle[i, lastCol];
                reservesByOrigin[i][iter] = bootTriangle[i, n - 1] - latest;
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

    [ExcelFunction(Description = "Berquist-Sherman paid loss adjustment for case reserve adequacy changes", Category = "Actuarial.ChainLadder")]
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
}
