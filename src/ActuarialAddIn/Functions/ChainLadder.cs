using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Factorization;
using System.Linq;

namespace ActuarialAddIn.Functions;

public static partial class ChainLadder
{
    #region Basic Chain Ladder

    [ExcelFunction(Description = "Calculate volume-weighted chain ladder development factors from a cumulative triangle. Standard reserving method per Mack (1993).", Category = "Actuarial.ChainLadder")]
    public static object ACT_CL_FACTORS(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Use only top N (oldest) years; default is all")] double topNYears = double.PositiveInfinity,
        [ExcelArgument(Description = "Exclude most recent year")] bool excludeMostRecentYear = false,
        [ExcelArgument(Description = "Exclude highest and lowest ratios")] bool excludeHighLowRatios = false,
        [ExcelArgument(Description = "Output as column (TRUE) or row (FALSE). Default TRUE.")] bool vertical = true)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return ExcelError.ExcelErrorValue;

        int maxYears = int.MaxValue;
        if (!double.IsInfinity(topNYears) && topNYears > 0)
            maxYears = (int)Math.Floor(topNYears);

        var factors = new double[n - 1];

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

        if (vertical)
        {
            var result = new object[n - 1, 1];
            for (int i = 0; i < n - 1; i++)
                result[i, 0] = factors[i];
            return result;
        }
        else
        {
            var result = new object[1, n - 1];
            for (int i = 0; i < n - 1; i++)
                result[0, i] = factors[i];
            return result;
        }
    }

    // Internal helper that returns double[] for use by other functions
    private static double[] GetFactorsArray(double[,] triangle)
    {
        int n = triangle.GetLength(0);
        var factors = new double[n - 1];

        for (int j = 0; j < n - 1; j++)
        {
            int count = n - j - 1;
            double sumCurrent = 0;
            double sumNext = 0;

            for (int i = 0; i < count; i++)
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

    [ExcelFunction(Description = "Return the latest diagonal of a cumulative triangle", Category = "Actuarial.ChainLadder")]
    public static object ACT_CL_LATEST(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Output as column (TRUE) or row (FALSE). Default TRUE.")] bool vertical = true)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return ExcelError.ExcelErrorValue;

        if (vertical)
        {
            var result = new object[n, 1];
            for (int i = 0; i < n; i++)
            {
                int lastCol = n - 1 - i;
                result[i, 0] = triangle[i, lastCol];
            }
            return result;
        }
        else
        {
            var result = new object[1, n];
            for (int i = 0; i < n; i++)
            {
                int lastCol = n - 1 - i;
                result[0, i] = triangle[i, lastCol];
            }
            return result;
        }
    }

    // Internal helper that returns double[] for use by other functions
    private static double[] GetLatestArray(double[,] triangle)
    {
        int n = triangle.GetLength(0);
        var latest = new double[n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            latest[i] = triangle[i, lastCol];
        }
        return latest;
    }

    [ExcelFunction(Description = "Project ultimate losses using chain ladder", Category = "Actuarial.ChainLadder")]
    public static object ACT_CL_ULTIMATE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Output as column (TRUE) or row (FALSE). Default TRUE.")] bool vertical = true)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return ExcelError.ExcelErrorValue;

        var factors = GetFactorsArray(triangle);
        var ultimates = GetUltimatesArray(triangle, factors);

        if (vertical)
        {
            var result = new object[n, 1];
            for (int i = 0; i < n; i++)
                result[i, 0] = ultimates[i];
            return result;
        }
        else
        {
            var result = new object[1, n];
            for (int i = 0; i < n; i++)
                result[0, i] = ultimates[i];
            return result;
        }
    }

    // Internal helper that returns double[] for use by other functions
    private static double[] GetUltimatesArray(double[,] triangle, double[] factors)
    {
        int n = triangle.GetLength(0);

        // Calculate cumulative factors to ultimate
        var cdfToUltimate = new double[n];
        cdfToUltimate[n - 1] = 1.0;
        for (int j = n - 2; j >= 0; j--)
        {
            cdfToUltimate[j] = cdfToUltimate[j + 1] * factors[j];
        }

        // Project ultimates
        var ultimates = new double[n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            double latestValue = triangle[i, lastCol];
            ultimates[i] = latestValue * cdfToUltimate[lastCol];
        }

        return ultimates;
    }

    [ExcelFunction(Description = "Calculate IBNR reserves using chain ladder", Category = "Actuarial.ChainLadder")]
    public static object ACT_CL_IBNR(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Output as column (TRUE) or row (FALSE). Default TRUE.")] bool vertical = true)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return ExcelError.ExcelErrorValue;

        var factors = GetFactorsArray(triangle);
        var ultimates = GetUltimatesArray(triangle, factors);

        if (vertical)
        {
            var result = new object[n, 1];
            for (int i = 0; i < n; i++)
            {
                int lastCol = n - 1 - i;
                double latestValue = triangle[i, lastCol];
                result[i, 0] = ultimates[i] - latestValue;
            }
            return result;
        }
        else
        {
            var result = new object[1, n];
            for (int i = 0; i < n; i++)
            {
                int lastCol = n - 1 - i;
                double latestValue = triangle[i, lastCol];
                result[0, i] = ultimates[i] - latestValue;
            }
            return result;
        }
    }

    [ExcelFunction(Description = "Bornhuetter-Ferguson ultimate: Ult = Paid + Unreported%. Balances actual experience with a priori expectation. Ref: Bornhuetter & Ferguson (1972). Use when data is immature or volatile.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_BF_ULTIMATE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Development factors (n-1)")] double[] developmentFactors,
        [ExcelArgument(Description = "A priori ultimate (n or single value)")] double[] aPrioriUltimate)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[] { ExcelError.ExcelErrorValue };

        if (developmentFactors.Length < n - 1)
            return new object[] { ExcelError.ExcelErrorValue };

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
            return new object[] { ExcelError.ExcelErrorValue };
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

}
