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
        if (n < 2)
            return new object[,] { { "Error: Triangle must have at least 2 development periods" } };

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

    [ExcelFunction(Description = "Weighted average of multiple ultimate estimates. Weights are renormalised so their sum is 1.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CL_WEIGHTED_AVERAGE(
        [ExcelArgument(Description = "First set of ultimates (n values)")] double[] ultimates1,
        [ExcelArgument(Description = "Weight for first set (>= 0)")] double weight1,
        [ExcelArgument(Description = "Second set of ultimates (n values)")] double[] ultimates2,
        [ExcelArgument(Description = "Weight for second set (>= 0)")] double weight2,
        [ExcelArgument(Description = "Third set of ultimates (optional)")] double[]? ultimates3 = null,
        [ExcelArgument(Description = "Weight for third set (>= 0; optional)")] double weight3 = 0)
    {
        if (ultimates1 == null || ultimates2 == null)
            return new object[] { ExcelError.ExcelErrorValue };

        int n = ultimates1.Length;
        if (ultimates2.Length != n)
            return new object[] { ExcelError.ExcelErrorValue };

        if (ultimates3 != null && ultimates3.Length != n)
            return new object[] { ExcelError.ExcelErrorValue };

        // Renormalise instead of erroring on weights that don't sum exactly to 1.
        double totalWeight = weight1 + weight2 + weight3;
        if (totalWeight <= 0)
            return new object[] { ExcelError.ExcelErrorValue };
        weight1 /= totalWeight;
        weight2 /= totalWeight;
        weight3 /= totalWeight;

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
