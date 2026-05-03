using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Factorization;
using System.Linq;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Chain ladder and reserving methods for general insurance loss reserving.
/// References:
/// - Mack, T. (1993). "Distribution-free calculation of the standard error of chain ladder reserve estimates." ASTIN Bulletin 23(2): 213-225.
/// - Mack, T. (1999). "The standard error of chain ladder reserve estimates: Recursive calculation and inclusion of a tail factor." ASTIN Bulletin 29(2): 361-366.
/// - England, P.D. and Verrall, R.J. (2002). "Stochastic claims reserving in general insurance." British Actuarial Journal 8(3): 443-518.
/// - Bornhuetter, R.L. and Ferguson, R.E. (1972). "The actuary and IBNR." Proceedings of the CAS, LIX: 181-195.
/// </summary>
public static class ChainLadder
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
            return "Error: Triangle must be square";

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
            return "Error: Triangle must be square";

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
            return "Error: Triangle must be square";

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
            return "Error: Triangle must be square";

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

    #region Mack Chain Ladder

    [ExcelFunction(Description = "Mack chain ladder standard errors for development factors. Distribution-free method per Mack (1993). Measures uncertainty in LDF estimates.", Category = "Actuarial.ChainLadder")]
    public static object ACT_MACK_FACTOR_SE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Output as column (TRUE) or row (FALSE). Default TRUE.")] bool vertical = true)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return "Error: Triangle must be square";

        var factors = GetFactorsArray(triangle);
        var factorSE = GetFactorSEArray(triangle, factors);

        if (vertical)
        {
            var result = new object[n - 1, 1];
            for (int i = 0; i < n - 1; i++)
                result[i, 0] = factorSE[i];
            return result;
        }
        else
        {
            var result = new object[1, n - 1];
            for (int i = 0; i < n - 1; i++)
                result[0, i] = factorSE[i];
            return result;
        }
    }

    // Internal helper that returns double[] for use by other functions
    private static double[] GetFactorSEArray(double[,] triangle, double[] factors)
    {
        int n = triangle.GetLength(0);
        var sigmaSquared = new double[n - 1];
        var sumWeights = new double[n - 1];
        var factorSE = new double[n - 1];

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
    public static object ACT_MACK_RESERVE_SE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Output as column (TRUE) or row (FALSE). Default TRUE.")] bool vertical = true)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return "Error: Triangle must be square";

        var factors = GetFactorsArray(triangle);
        var ultimates = GetUltimatesArray(triangle, factors);
        var factorSE = GetFactorSEArray(triangle, factors);

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
        var reserveSE = new double[n];
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

        if (vertical)
        {
            var result = new object[n, 1];
            for (int i = 0; i < n; i++)
                result[i, 0] = reserveSE[i];
            return result;
        }
        else
        {
            var result = new object[1, n];
            for (int i = 0; i < n; i++)
                result[0, i] = reserveSE[i];
            return result;
        }
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

    // ===================================================================
    // Bootstrap Internal Helpers
    // Implements the full England & Verrall (2002) / England (2010) ODP
    // bootstrap with: GLM hat-matrix adjusted residuals, corner exclusion,
    // per-period φⱼ via proportional df, Mack tail extrapolation, scaled
    // global pool with Bessel correction, pseudo-diagonal IBNR.
    //
    // Algorithm walkthrough, step-by-step reconciliation against England
    // (2010) slide 35, Shapland (2016)'s companion Excel and R
    // `ChainLadder::BootChainLadder` lives in:
    //   https://github.com/mdevans21/bootstrapping_exposition
    // ===================================================================

    /// <summary>
    /// Compute fitted cumulative and incremental values for all cells in the upper triangle.
    /// Uses backward projection from the diagonal: fitted_cum[i,j] = diagonal[i] / CDF(j→lastCol_i).
    /// This matches the ODP/multiplicative model fit used by chainladder-python's full_expectation_.
    /// </summary>
    private static void ComputeFittedIncrementals(double[,] triangle, double[] factors, int n,
        out double[,] fittedIncr, out double[,] actualIncr, out int nObserved)
    {
        fittedIncr = new double[n, n];
        actualIncr = new double[n, n];
        nObserved = 0;

        // Compute fitted cumulative via backward projection from diagonal
        var fittedCum = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            fittedCum[i, lastCol] = triangle[i, lastCol]; // exact at diagonal
            for (int j = lastCol - 1; j >= 0; j--)
                fittedCum[i, j] = fittedCum[i, j + 1] / factors[j];
        }

        // Convert to incremental and compute actuals
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            fittedIncr[i, 0] = fittedCum[i, 0];
            actualIncr[i, 0] = triangle[i, 0];
            nObserved++;
            for (int j = 1; j <= lastCol; j++)
            {
                fittedIncr[i, j] = fittedCum[i, j] - fittedCum[i, j - 1];
                actualIncr[i, j] = triangle[i, j] - triangle[i, j - 1];
                nObserved++;
            }
        }
    }

    /// <summary>
    /// Compute unscaled Pearson residuals: r_ij = (actual - fitted) / sqrt(|fitted|).
    /// These are used for phi computation (always unscaled, never hat-adjusted).
    /// </summary>
    private static double[,] ComputeUnscaledResiduals(double[,] actualIncr, double[,] fittedIncr, int n)
    {
        var resids = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                double m = Math.Abs(fittedIncr[i, j]);
                if (m > 0)
                    resids[i, j] = (actualIncr[i, j] - fittedIncr[i, j]) / Math.Sqrt(m);
            }
        }
        return resids;
    }

    /// <summary>
    /// Compute hat matrix diagonal for the ODP/multiplicative model with log link
    /// (England &amp; Verrall 2002, eq. 2.5):
    ///   log m_ij = a_i + b_j     (β_0 = 0 identifiability constraint)
    /// Working weight at IRLS convergence is W = diag(m_ij). The hat matrix is
    ///   H = W^{1/2} Z (Z' W Z)^{-1} Z' W^{1/2}
    /// computed numerically via thin QR of A = √W · Z, giving h_ii = ||q_i||².
    ///
    /// This matches chainladder-python's `BootstrapODPSample.hat_` (which exposes
    /// hat_factor = 1/√(1−h_ii)) and the Shapland (2016) Excel companion within
    /// floating-point precision; verified in the bootstrapping_exposition repo.
    /// </summary>
    private static double[,] ComputeHatDiag(double[,] fittedIncr, int n)
    {
        var hii = new double[n, n];

        int p = 2 * n - 1; // n origin params + (n-1) dev params (β_0 = 0)
        int nObs = n * (n + 1) / 2;

        var A = Matrix<double>.Build.Dense(nObs, p);
        var rowIndex = new (int i, int j)[nObs];

        int row = 0;
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                double w = Math.Sqrt(Math.Abs(fittedIncr[i, j]));
                A[row, i] = w; // origin parameter column
                if (j > 0)
                    A[row, n + j - 1] = w; // dev parameter column (β_0 absorbed)
                rowIndex[row] = (i, j);
                row++;
            }
        }

        QR<double> qr;
        try
        {
            qr = A.QR(QRMethod.Thin);
        }
        catch
        {
            // Fallback to the marginal-totals approximation if QR fails on a
            // degenerate triangle. This branch should never fire on the standard
            // square ODP setup.
            return ComputeHatDiagMarginalTotals(fittedIncr, n);
        }

        var Q = qr.Q;
        for (int r = 0; r < nObs; r++)
        {
            double sumSq = 0;
            for (int k = 0; k < p; k++)
                sumSq += Q[r, k] * Q[r, k];

            var (i, j) = rowIndex[r];
            // Numerical guard: full-leverage cells should be exactly 1 but
            // can come out as 1 + ε; clip to [0, 1].
            if (sumSq > 1.0) sumSq = 1.0;
            if (sumSq < 0.0) sumSq = 0.0;
            hii[i, j] = sumSq;
        }

        return hii;
    }

    /// <summary>
    /// Closed-form marginal-totals approximation to the ODP hat matrix.
    /// Kept as a fallback for degenerate triangles where the QR of √W·Z fails.
    /// </summary>
    private static double[,] ComputeHatDiagMarginalTotals(double[,] fittedIncr, int n)
    {
        var hii = new double[n, n];
        var rowSums = new double[n];
        var colSums = new double[n];
        double grandTotal = 0;

        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                double m = fittedIncr[i, j];
                if (m > 0)
                {
                    rowSums[i] += m;
                    colSums[j] += m;
                    grandTotal += m;
                }
            }
        }

        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                double m = fittedIncr[i, j];
                if (m > 0 && rowSums[i] > 0 && colSums[j] > 0 && grandTotal > 0)
                {
                    double h = m * (1.0 / rowSums[i] + 1.0 / colSums[j] - 1.0 / grandTotal);
                    // Clamp to [0, 0.999] — the marginal-totals approximation
                    // can run negative or exceed 1 on pathological triangles,
                    // which would shrink (h<0) or blow up (h≈1) the residual
                    // standardisation r/√(1−h).
                    if (h < 0) h = 0;
                    else if (h > 0.999) h = 0.999;
                    hii[i, j] = h;
                }
            }
        }

        return hii;
    }

    /// <summary>
    /// Prepare bootstrap inputs for the EV (England & Verrall) method.
    /// Implements: hat-adjusted residuals, corner exclusion, per-period phi_j with
    /// proportional df, Mack extrapolation, scaled global pool with Bessel correction.
    /// </summary>
    private static bool PrepareEVBootstrap(double[,] triangle, double[] factors, int n,
        out double[,] fittedIncr, out double[] scaledPool,
        out double[] phiByDev, out double phiGlobal, out string error)
    {
        scaledPool = Array.Empty<double>();
        phiByDev = new double[n];
        phiGlobal = 1.0;
        error = "";

        ComputeFittedIncrementals(triangle, factors, n, out fittedIncr, out var actualIncr, out int nObs);
        var unscaled = ComputeUnscaledResiduals(actualIncr, fittedIncr, n);
        var hatDiag = ComputeHatDiag(fittedIncr, n);

        // Hat-adjusted residuals: r_adj = r / sqrt(1 - h_ii)
        var adjResids = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                double h = hatDiag[i, j];
                if (h < 1.0 && unscaled[i, j] != 0)
                    adjResids[i, j] = unscaled[i, j] / Math.Sqrt(1.0 - h);
            }
        }

        // Exclude corners: top-right (0, n-1) and bottom-left (n-1, 0)
        adjResids[0, n - 1] = 0.0;
        adjResids[n - 1, 0] = 0.0;

        // Degrees of freedom
        int nParams = n + (n - 1); // n origin + (n-1) development parameters
        int degFree = nObs - nParams;
        if (degFree <= 0)
        {
            error = "Error: Not enough degrees of freedom for bootstrap";
            return false;
        }

        // Per-period phi_j from UNSCALED residuals with proportional df
        // df_j = n_j × (n_obs - p) / n_obs (England 2010)
        for (int j = 0; j < n; j++)
        {
            double ssrJ = 0;
            int nJ = 0;
            for (int i = 0; i < n; i++)
            {
                int lastCol = n - 1 - i;
                if (j <= lastCol && unscaled[i, j] != 0)
                {
                    ssrJ += unscaled[i, j] * unscaled[i, j];
                    nJ++;
                }
            }
            double propDf = (double)nJ * degFree / nObs;
            if (nJ >= 2 && ssrJ > 0 && propDf > 0)
                phiByDev[j] = ssrJ / propDf;
        }

        // Global phi for reference and fallback
        double totalSSR = 0;
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
                totalSSR += unscaled[i, j] * unscaled[i, j];
        }
        phiGlobal = totalSSR / degFree;

        // Mack (1993) extrapolation for periods where phi can't be estimated
        int lastValid = -1;
        for (int j = n - 1; j >= 0; j--)
        {
            if (phiByDev[j] > 1.0)
            {
                lastValid = j;
                break;
            }
        }
        for (int j = lastValid + 1; j < n; j++)
        {
            if (lastValid >= 1)
            {
                double p1 = phiByDev[lastValid], p2 = phiByDev[lastValid - 1];
                phiByDev[j] = p2 > 0 ? Math.Min(p1 * p1 / p2, Math.Min(p1, p2)) : p1;
            }
            else
            {
                phiByDev[j] = phiGlobal;
            }
        }

        // Build scaled global pool (England 2010):
        // 1. Standardize hat-adjusted residuals by sqrt(phi_j)
        // 2. Pool globally, centre, apply Bessel correction
        var poolList = new List<double>();
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                double r = adjResids[i, j];
                if (r != 0 && phiByDev[j] > 0)
                    poolList.Add(r / Math.Sqrt(phiByDev[j]));
            }
        }

        if (poolList.Count < 5)
        {
            error = "Error: Not enough residuals for bootstrap";
            return false;
        }

        // Centre
        double poolMean = poolList.Average();
        for (int i = 0; i < poolList.Count; i++)
            poolList[i] -= poolMean;

        // Bessel correction: sqrt(N/(N-1))
        int poolN = poolList.Count;
        double bessel = Math.Sqrt((double)poolN / (poolN - 1));
        for (int i = 0; i < poolList.Count; i++)
            poolList[i] *= bessel;

        scaledPool = poolList.ToArray();
        return true;
    }

    /// <summary>
    /// Prepare bootstrap inputs for the BASIC method (chainladder-python compatible).
    /// No hat adjustment, no corner exclusion, constant global phi, simple centred pool.
    /// </summary>
    private static bool PrepareBasicBootstrap(double[,] triangle, double[] factors, int n,
        out double[,] fittedIncr, out double[] residPool, out double phiGlobal, out string error)
    {
        residPool = Array.Empty<double>();
        phiGlobal = 1.0;
        error = "";

        ComputeFittedIncrementals(triangle, factors, n, out fittedIncr, out var actualIncr, out int nObs);
        var unscaled = ComputeUnscaledResiduals(actualIncr, fittedIncr, n);

        int nParams = n + (n - 1);
        int degFree = nObs - nParams;
        if (degFree <= 0)
        {
            error = "Error: Not enough degrees of freedom for bootstrap";
            return false;
        }

        // Global phi
        double totalSSR = 0;
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
                totalSSR += unscaled[i, j] * unscaled[i, j];
        }
        phiGlobal = totalSSR / degFree;

        // Pool all non-zero residuals, centre
        var poolList = new List<double>();
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                if (unscaled[i, j] != 0)
                    poolList.Add(unscaled[i, j]);
            }
        }

        if (poolList.Count < 5)
        {
            error = "Error: Not enough residuals for bootstrap";
            return false;
        }

        double poolMean = poolList.Average();
        residPool = poolList.Select(r => r - poolMean).ToArray();
        return true;
    }

    /// <summary>
    /// Run a single bootstrap iteration. Returns IBNR by origin year.
    /// For EV: uses scaled global pool, un-standardizes by sqrt(phi_j), pseudo-diagonal convention.
    /// For BASIC: uses raw residual pool, constant phi, original diagonal convention.
    ///
    /// Process variance follows England (2010): each projected incremental cell gets
    /// independent Gamma noise. The expected incrementals come from the deterministic
    /// chain ladder projection of the pseudo-triangle (not sequential random walk).
    /// </summary>
    private static double[] BootstrapOneIteration(
        double[,] triangle, double[,] fittedIncr,
        double[] pool, double[]? phiByDev, double phiGlobal,
        bool isScaledPool, bool usePseudoDiag, int n, Random rng)
    {
        // Step 1: Create pseudo-incremental triangle from resampled residuals
        var bootIncr = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            int lastCol = n - 1 - i;
            for (int j = 0; j <= lastCol; j++)
            {
                double fitted = fittedIncr[i, j];
                if (Math.Abs(fitted) > 0)
                {
                    double rSample = pool[rng.Next(pool.Length)];
                    if (isScaledPool && phiByDev != null)
                    {
                        // EV: un-standardize by sqrt(phi_j) for target column
                        double sqrtPhiJ = phiByDev[j] > 0 ? Math.Sqrt(phiByDev[j]) : 1.0;
                        bootIncr[i, j] = rSample * sqrtPhiJ * Math.Sqrt(Math.Abs(fitted)) + fitted;
                    }
                    else
                    {
                        // BASIC: direct application
                        bootIncr[i, j] = rSample * Math.Sqrt(Math.Abs(fitted)) + fitted;
                    }
                }
                else
                {
                    bootIncr[i, j] = fitted;
                }
            }
        }

        // Step 2: Convert pseudo-incremental to pseudo-cumulative
        var pseudoCum = new double[n, n];
        for (int i = 0; i < n; i++)
        {
            double cum = 0;
            for (int j = 0; j < n - i; j++)
            {
                cum += bootIncr[i, j];
                pseudoCum[i, j] = cum;
            }
        }

        // Step 3: Re-estimate factors from pseudo-triangle
        var bootFactors = CalculateFactors(pseudoCum);

        // Step 4: Get pseudo-diagonal (ALWAYS used for projection starting point)
        // and the IBNR subtraction diagonal (pseudo for EV, original for BASIC).
        // This follows chainladder-python: the lower triangle is always projected
        // from the pseudo-data, but IBNR convention chooses what to subtract.
        var pseudoDiag = new double[n];
        for (int i = 0; i < n; i++)
            pseudoDiag[i] = pseudoCum[i, n - 1 - i];

        // Step 5: Project forward from pseudo-diagonal with independent process variance.
        // Deterministic projection first, then independent Gamma noise per cell (England 2010).
        var ibnr = new double[n];
        ibnr[0] = 0.0; // Fully developed origin
        for (int i = 1; i < n; i++)
        {
            int lastCol = n - 1 - i;

            // Deterministic chain ladder projection from pseudo-diagonal
            var projCum = new double[n];
            projCum[lastCol] = pseudoDiag[i];
            for (int j = lastCol; j < n - 1; j++)
                projCum[j + 1] = projCum[j] * bootFactors[j];

            // Apply independent process variance to each projected incremental
            double ultimate = pseudoDiag[i];
            for (int j = lastCol + 1; j < n; j++)
            {
                double expIncr = projCum[j] - projCum[j - 1];
                double phiJ = (phiByDev != null && j < phiByDev.Length)
                    ? phiByDev[j] : phiGlobal;
                if (phiJ <= 0) phiJ = phiGlobal;

                double noisyIncr;
                if (expIncr > 0 && phiJ > 0)
                {
                    double shape = expIncr / phiJ;
                    noisyIncr = shape > 1e-10
                        ? Gamma.Sample(rng, shape, 1.0 / phiJ)
                        : expIncr;
                }
                else if (expIncr < 0 && phiJ > 0)
                {
                    double shape = Math.Abs(expIncr) / phiJ;
                    noisyIncr = shape > 1e-10
                        ? -Gamma.Sample(rng, shape, 1.0 / phiJ)
                        : expIncr;
                }
                else
                {
                    noisyIncr = expIncr;
                }
                ultimate += noisyIncr;
            }

            // IBNR = projected ultimate minus the appropriate diagonal
            double subtractDiag = usePseudoDiag ? pseudoDiag[i] : triangle[i, lastCol];
            ibnr[i] = ultimate - subtractDiag;
        }

        return ibnr;
    }

    #endregion

    #region Bootstrap Methods

    [ExcelFunction(Description = "[EXPERIMENTAL] ODP Bootstrap for total reserve distribution. Default method='EV' implements full England & Verrall (2002) methodology with hat-adjusted residuals, corner exclusion, per-period scale parameters, and pseudo-diagonal convention. Use method='CHAINLADDER-PYTHON' for chainladder-python compatible results (constant phi, no hat matrix, original diagonal). Returns statistics: Mean, StdDev, P1, P5, P10, P25, P50, P75, P90, P95, P99.", Category = "Actuarial.Experimental")]
    public static object[,] ACT_CL_BOOTSTRAP(
        [ExcelArgument(Description = "Cumulative triangle (n x n, zeros for future cells)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations (recommend 10000+)")] int iterations,
        [ExcelArgument(Description = "Random seed for reproducibility (leave blank for system seed)")] int? seed = null,
        [ExcelArgument(Description = "Method: 'EV' (default) = full England & Verrall (2002) with hat matrix, per-period phi, pseudo-diagonal. 'CHAINLADDER-PYTHON' = chainladder-python compatible with constant phi and original diagonal.")] string method = "EV")
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        if (iterations <= 0)
            return new object[,] { { "Error: Iterations must be positive" } };

        var rng = seed is null ? new Random() : new Random(seed.Value);
        var factors = GetFactorsArray(triangle);

        double[,] fittedIncr;
        double[] pool;
        double[]? phiByDev;
        double phiGlobal;
        bool isScaledPool, usePseudoDiag;
        string error;

        string m = string.IsNullOrWhiteSpace(method) ? "EV" : method.Trim().ToUpperInvariant();
        if (m == "EV")
        {
            if (!PrepareEVBootstrap(triangle, factors, n,
                out fittedIncr, out pool, out phiByDev, out phiGlobal, out error))
                return new object[,] { { error } };
            isScaledPool = true;
            usePseudoDiag = true;
        }
        else if (m == "CHAINLADDER-PYTHON")
        {
            if (!PrepareBasicBootstrap(triangle, factors, n,
                out fittedIncr, out pool, out phiGlobal, out error))
                return new object[,] { { error } };
            phiByDev = null;
            isScaledPool = false;
            usePseudoDiag = false;
        }
        else
        {
            return new object[,] { { "Error: method must be 'EV' or 'CHAINLADDER-PYTHON'" } };
        }

        var bootstrapReserves = new double[iterations];
        for (int iter = 0; iter < iterations; iter++)
        {
            var ibnr = BootstrapOneIteration(triangle, fittedIncr, pool,
                phiByDev, phiGlobal, isScaledPool, usePseudoDiag, n, rng);
            bootstrapReserves[iter] = ibnr.Sum();
        }

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

    [ExcelFunction(Description = "[EXPERIMENTAL] ODP Bootstrap reserve distribution by origin year. Default method='EV' implements full England & Verrall (2002) methodology. Use method='CHAINLADDER-PYTHON' for chainladder-python compatible results (constant phi, no hat matrix, original diagonal). Returns header row plus data for each accident year with columns: AY, Mean, StdDev, P50, P75, P90, P95, P99.", Category = "Actuarial.Experimental")]
    public static object[,] ACT_CL_BOOTSTRAP_ORIGIN(
        [ExcelArgument(Description = "Cumulative triangle (n x n, zeros for future cells)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations (recommend 10000+)")] int iterations,
        [ExcelArgument(Description = "Random seed for reproducibility (leave blank for system seed)")] int? seed = null,
        [ExcelArgument(Description = "Method: 'EV' (default) = full England & Verrall (2002). 'CHAINLADDER-PYTHON' = chainladder-python compatible with constant phi and original diagonal.")] string method = "EV")
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        if (iterations <= 0)
            return new object[,] { { "Error: Iterations must be positive" } };

        var rng = seed is null ? new Random() : new Random(seed.Value);
        var factors = GetFactorsArray(triangle);

        double[,] fittedIncr;
        double[] pool;
        double[]? phiByDev;
        double phiGlobal;
        bool isScaledPool, usePseudoDiag;
        string error;

        string m = string.IsNullOrWhiteSpace(method) ? "EV" : method.Trim().ToUpperInvariant();
        if (m == "EV")
        {
            if (!PrepareEVBootstrap(triangle, factors, n,
                out fittedIncr, out pool, out phiByDev, out phiGlobal, out error))
                return new object[,] { { error } };
            isScaledPool = true;
            usePseudoDiag = true;
        }
        else if (m == "CHAINLADDER-PYTHON")
        {
            if (!PrepareBasicBootstrap(triangle, factors, n,
                out fittedIncr, out pool, out phiGlobal, out error))
                return new object[,] { { error } };
            phiByDev = null;
            isScaledPool = false;
            usePseudoDiag = false;
        }
        else
        {
            return new object[,] { { "Error: method must be 'EV' or 'CHAINLADDER-PYTHON'" } };
        }

        var reservesByOrigin = new double[n][];
        for (int i = 0; i < n; i++)
            reservesByOrigin[i] = new double[iterations];

        for (int iter = 0; iter < iterations; iter++)
        {
            var ibnr = BootstrapOneIteration(triangle, fittedIncr, pool,
                phiByDev, phiGlobal, isScaledPool, usePseudoDiag, n, rng);
            for (int i = 0; i < n; i++)
                reservesByOrigin[i][iter] = ibnr[i];
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
            return "Error: Triangle must be square";

        if (premium == null || premium.Length != n)
            return "Error: Premium array must have n values";

        // Get development factors if not provided
        double[] factors;
        if (developmentFactors == null || developmentFactors.Length == 0)
        {
            factors = GetFactorsArray(triangle);
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
