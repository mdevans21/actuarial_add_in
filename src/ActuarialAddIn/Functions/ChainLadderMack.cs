using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using System.Linq;

namespace ActuarialAddIn.Functions;

public static partial class ChainLadder
{
    #region Mack Chain Ladder

    [ExcelFunction(Description = "Mack chain ladder standard errors for development factors. Distribution-free method per Mack (1993). Measures uncertainty in LDF estimates.", Category = "Actuarial.ChainLadder")]
    public static object ACT_MACK_FACTOR_SE(
        [ExcelArgument(Description = "Triangle data (n x n cumulative values)")] double[,] triangle,
        [ExcelArgument(Description = "Output as column (TRUE) or row (FALSE). Default TRUE.")] bool vertical = true)
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return ExcelError.ExcelErrorValue;

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
            return ExcelError.ExcelErrorValue;

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
    // Legacy basic bootstrap helpers. The EV reference implementation lives in
    // StochasticReservingBootstrap.cs and follows Peter England's Python code.
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

}
