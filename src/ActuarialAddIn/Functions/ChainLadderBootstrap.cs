using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Factorization;
using System.Linq;

namespace ActuarialAddIn.Functions;

public static partial class ChainLadder
{
    #region Bootstrap Methods

    [ExcelFunction(Description = "[EXPERIMENTAL] ODP Bootstrap for total reserve distribution. Default method='EV' implements full England & Verrall (2002) methodology with hat-adjusted residuals, corner exclusion, per-period scale parameters, and pseudo-diagonal convention. Use method='CHAINLADDER-PYTHON' for chainladder-python compatible results (constant phi, no hat matrix, original diagonal). Returns statistics: Mean, StdDev, P1, P5, P10, P25, P50, P75, P90, P95, P99.", Category = "Actuarial.Experimental")]
    public static object[,] ACT_CL_BOOTSTRAP(
        [ExcelArgument(Description = "Cumulative triangle (n x n, zeros for future cells)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations (recommend 10000+)")] int iterations,
        [ExcelArgument(Description = "Random seed for reproducibility (leave blank for system seed)")] object? seed = null,
        [ExcelArgument(Description = "Method: 'EV' (default) = full England & Verrall (2002) with hat matrix, per-period phi, pseudo-diagonal. 'CHAINLADDER-PYTHON' = chainladder-python compatible with constant phi and original diagonal.")] string method = "EV")
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        if (iterations <= 0)
            return new object[,] { { "Error: Iterations must be positive" } };

        var rng = SeedUtil.ResolveSeed(seed) is { } _seed ? new Random(_seed) : new Random();
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
        [ExcelArgument(Description = "Random seed for reproducibility (leave blank for system seed)")] object? seed = null,
        [ExcelArgument(Description = "Method: 'EV' (default) = full England & Verrall (2002). 'CHAINLADDER-PYTHON' = chainladder-python compatible with constant phi and original diagonal.")] string method = "EV")
    {
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
            return new object[,] { { "Error: Triangle must be square" } };

        if (iterations <= 0)
            return new object[,] { { "Error: Iterations must be positive" } };

        var rng = SeedUtil.ResolveSeed(seed) is { } _seed ? new Random(_seed) : new Random();
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

}
