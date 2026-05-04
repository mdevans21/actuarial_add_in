using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Factorization;
using System.Linq;

namespace ActuarialAddIn.Functions;

public static partial class ChainLadder
{
    #region Cape Cod Method

    [ExcelFunction(Description = "Cape Cod ultimate: Alternative to B-F that estimates ELR from the data itself. Ult = Paid + ELR * OnLevelPremium * UnreportedPct. Good when a priori ELR is uncertain.", Category = "Actuarial.ChainLadder")]
    public static object[] ACT_CAPECOD_ULTIMATE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "On-level earned premium by origin year (n values)")] double[] premium,
        [ExcelArgument(Description = "Development factors (n-1). If omitted, calculated from triangle.")] double[]? developmentFactors = null)
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
            factors = GetFactorsArray(triangle);
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
        [ExcelArgument(Description = "Development factors (n-1). If omitted, calculated from triangle.")] double[]? developmentFactors = null)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return ExcelError.ExcelErrorValue;

        if (premium == null || premium.Length != n)
            return ExcelError.ExcelErrorValue;

        // Get development factors if not provided
        double[] factors;
        if (developmentFactors == null || developmentFactors.Length == 0)
        {
            factors = GetFactorsArray(triangle);
        }
        else
        {
            if (developmentFactors.Length < n - 1)
                return ExcelError.ExcelErrorValue;
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
            return ExcelError.ExcelErrorValue;

        return sumLatest / sumUsedUp;
    }

    #endregion

}
