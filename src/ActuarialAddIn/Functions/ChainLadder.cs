using ExcelDna.Integration;
using MathNet.Numerics.Distributions;

namespace ActuarialAddIn.Functions;

public static class ChainLadder
{
    #region Basic Chain Ladder

    [ExcelFunction(Description = "Calculate chain ladder development factors", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_FACTORS(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { "Error: Triangle must be square" };

        var factors = new object[n - 1];

        for (int j = 0; j < n - 1; j++)
        {
            double sumCurrent = 0;
            double sumNext = 0;
            int count = 0;

            for (int i = 0; i < n - j - 1; i++)
            {
                if (triangle[i, j] > 0 && triangle[i, j + 1] > 0)
                {
                    sumCurrent += triangle[i, j];
                    sumNext += triangle[i, j + 1];
                    count++;
                }
            }

            factors[j] = count > 0 ? sumNext / sumCurrent : 1.0;
        }

        return factors;
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
        var factorSE = new object[n - 1];

        // Calculate sigma^2 (variance of individual factors)
        for (int j = 0; j < n - 1; j++)
        {
            double sumWeightedVar = 0;
            double sumWeights = 0;
            int count = n - j - 1;

            for (int i = 0; i < count; i++)
            {
                if (triangle[i, j] > 0 && triangle[i, j + 1] > 0)
                {
                    double individualFactor = triangle[i, j + 1] / triangle[i, j];
                    sumWeightedVar += triangle[i, j] * Math.Pow(individualFactor - factors[j], 2);
                    sumWeights += triangle[i, j];
                }
            }

            sigmaSquared[j] = count > 1 ? sumWeightedVar / (count - 1) : 0;
            factorSE[j] = sumWeights > 0 ? Math.Sqrt(sigmaSquared[j] / sumWeights) : 0;
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
        for (int j = 0; j < n - 1; j++)
        {
            sigmaSquared[j] = Math.Pow(factorSE[j], 2) * GetSumC(triangle, j);
        }

        // Calculate reserve standard errors using Mack formula
        var reserveSE = new object[n];
        reserveSE[0] = 0.0; // First period has no uncertainty

        for (int i = 1; i < n; i++)
        {
            double variance = 0;
            int lastCol = n - 1 - i;
            double projected = ultimates[i];

            for (int j = lastCol; j < n - 1; j++)
            {
                double Cij = GetProjectedC(triangle, factors, i, j);
                double term1 = sigmaSquared[j] / Cij;
                double term2 = sigmaSquared[j] / GetSumC(triangle, j);
                variance += Math.Pow(projected, 2) * (term1 + term2);
            }

            reserveSE[i] = Math.Sqrt(variance);
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

        // Calculate residuals
        var residuals = new List<double>();
        for (int j = 0; j < n - 1; j++)
        {
            for (int i = 0; i < n - j - 1; i++)
            {
                if (triangle[i, j] > 0 && triangle[i, j + 1] > 0)
                {
                    double expected = triangle[i, j] * factors[j];
                    double actual = triangle[i, j + 1];
                    double residual = (actual - expected) / Math.Sqrt(triangle[i, j]);
                    residuals.Add(residual);
                }
            }
        }

        if (residuals.Count == 0)
            return new object[,] { { "Error: No valid residuals" } };

        var residualArray = residuals.ToArray();

        // Bootstrap iterations
        var bootstrapReserves = new double[iterations];
        for (int iter = 0; iter < iterations; iter++)
        {
            // Create bootstrap triangle
            var bootTriangle = (double[,])triangle.Clone();

            // Simulate future development
            for (int i = 1; i < n; i++)
            {
                int lastCol = n - 1 - i;
                for (int j = lastCol; j < n - 1; j++)
                {
                    double current = bootTriangle[i, j];
                    double expected = current * factors[j];

                    // Resample residual
                    double residual = residualArray[random.Next(residualArray.Length)];
                    double simulated = expected + residual * Math.Sqrt(current);

                    bootTriangle[i, j + 1] = Math.Max(simulated, current); // Ensure non-decreasing
                }
            }

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
        result[0, 1] = bootstrapReserves.Average();
        result[1, 0] = "StdDev";
        result[1, 1] = Math.Sqrt(bootstrapReserves.Select(x => Math.Pow(x - bootstrapReserves.Average(), 2)).Average());

        for (int i = 0; i < percentiles.Length; i++)
        {
            result[i + 2, 0] = $"P{percentiles[i] * 100:0}";
            int idx = Math.Min((int)(percentiles[i] * iterations), iterations - 1);
            result[i + 2, 1] = bootstrapReserves[idx];
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
