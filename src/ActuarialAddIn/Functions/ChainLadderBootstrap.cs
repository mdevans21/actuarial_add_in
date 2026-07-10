using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

public static partial class ChainLadder
{
    private const string BootstrapCategory = "Actuarial.Experimental";
    private const int MaxBootstrapIterations = 1_000_000;
    private const long MaxStoredSampleValues = 25_000_000;
    private const long MaxRawOutputCells = 5_000_000;

    [ExcelFunction(
        Description = "[EXPERIMENTAL] ODP bootstrap total reserve distribution. EV follows StochasticReserving Main_ODP_Bstrap pathwise; CHAINLADDER-PYTHON retains the legacy basic mode.",
        Category = BootstrapCategory)]
    public static object[,] ACT_CL_BOOTSTRAP(
        [ExcelArgument(Description = "Cumulative triangle (n x n, zeros for future cells)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations")] int iterations,
        [ExcelArgument(Description = "Random seed; leave blank for system seed")] object? seed = null,
        [ExcelArgument(Description = "EV (default) or CHAINLADDER-PYTHON")] string method = "EV",
        [ExcelArgument(Description = "EV scale: NONCONSTANT (default) or CONSTANT")] string scale = "NONCONSTANT",
        [ExcelArgument(Description = "EV bootstrap distribution: NONPARAMETRIC, GAMMA, or LOGNORMAL")] string bootstrapDistribution = "NONPARAMETRIC",
        [ExcelArgument(Description = "EV forecast distribution: GAMMA, LOGNORMAL, or NONPARAMETRIC")] string forecastDistribution = "GAMMA",
        [ExcelArgument(Description = "Optional (n-1) x (n-1) EV link-ratio inclusion mask")] object? mask = null,
        [ExcelArgument(Description = "Optional n-value EV forecast sqrt-scale vector")] object? userSqrtScale = null)
    {
        if (!TryGenerateReserveSamples(
                triangle, iterations, seed, method, scale, bootstrapDistribution,
                forecastDistribution, mask, userSqrtScale,
                out var reserves, out var totals, out string error))
            return Error(error);

        var sorted = (double[])totals.Clone();
        Array.Sort(sorted);
        double mean = totals.Average();
        double standardDeviation = PopulationStandardDeviation(totals, mean);
        double[] percentiles = { 0.01, 0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95, 0.99 };
        var result = new object[percentiles.Length + 2, 2];
        result[0, 0] = "Mean";
        result[0, 1] = mean;
        result[1, 0] = "StdDev";
        result[1, 1] = standardDeviation;
        for (int i = 0; i < percentiles.Length; i++)
        {
            result[i + 2, 0] = $"P{percentiles[i] * 100:0}";
            result[i + 2, 1] = ReferencePercentile(sorted, percentiles[i]);
        }
        return result;
    }

    [ExcelFunction(
        Description = "[EXPERIMENTAL] ODP bootstrap reserve distribution by origin year. EV follows StochasticReserving Main_ODP_Bstrap pathwise.",
        Category = BootstrapCategory)]
    public static object[,] ACT_CL_BOOTSTRAP_ORIGIN(
        [ExcelArgument(Description = "Cumulative triangle (n x n, zeros for future cells)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations")] int iterations,
        [ExcelArgument(Description = "Random seed; leave blank for system seed")] object? seed = null,
        [ExcelArgument(Description = "EV (default) or CHAINLADDER-PYTHON")] string method = "EV",
        [ExcelArgument(Description = "EV scale: NONCONSTANT (default) or CONSTANT")] string scale = "NONCONSTANT",
        [ExcelArgument(Description = "EV bootstrap distribution: NONPARAMETRIC, GAMMA, or LOGNORMAL")] string bootstrapDistribution = "NONPARAMETRIC",
        [ExcelArgument(Description = "EV forecast distribution: GAMMA, LOGNORMAL, or NONPARAMETRIC")] string forecastDistribution = "GAMMA",
        [ExcelArgument(Description = "Optional (n-1) x (n-1) EV link-ratio inclusion mask")] object? mask = null,
        [ExcelArgument(Description = "Optional n-value EV forecast sqrt-scale vector")] object? userSqrtScale = null)
    {
        if (!TryGenerateReserveSamples(
                triangle, iterations, seed, method, scale, bootstrapDistribution,
                forecastDistribution, mask, userSqrtScale,
                out var reserves, out _, out string error))
            return Error(error);

        int n = reserves.GetLength(1);
        double[] percentiles = { 0.50, 0.75, 0.90, 0.95, 0.99 };
        var result = new object[n + 1, percentiles.Length + 3];
        result[0, 0] = "AY";
        result[0, 1] = "Mean";
        result[0, 2] = "StdDev";
        for (int p = 0; p < percentiles.Length; p++)
            result[0, p + 3] = $"P{percentiles[p] * 100:0}";

        var samples = new double[reserves.GetLength(0)];
        for (int origin = 0; origin < n; origin++)
        {
            for (int iteration = 0; iteration < samples.Length; iteration++)
                samples[iteration] = reserves[iteration, origin];
            double mean = samples.Average();
            result[origin + 1, 0] = origin + 1;
            result[origin + 1, 1] = mean;
            result[origin + 1, 2] = PopulationStandardDeviation(samples, mean);
            Array.Sort(samples);
            for (int p = 0; p < percentiles.Length; p++)
                result[origin + 1, p + 3] = ReferencePercentile(samples, percentiles[p]);
        }
        return result;
    }

    [ExcelFunction(
        Description = "[EXPERIMENTAL] Raw ODP bootstrap paths. EV follows StochasticReserving Main_ODP_Bstrap pathwise. Select RESERVES, ULTIMATES, PSEUDO-LRS, CUMULATIVES, or COMPLETE-CUMULATIVES.",
        Category = BootstrapCategory)]
    public static object[,] ACT_CL_BOOTSTRAP_SAMPLES(
        [ExcelArgument(Description = "Cumulative triangle (n x n, zeros for future cells)")] double[,] triangle,
        [ExcelArgument(Description = "Number of bootstrap iterations")] int iterations,
        [ExcelArgument(Description = "Random seed; leave blank for system seed")] object? seed = null,
        [ExcelArgument(Description = "EV (default) or CHAINLADDER-PYTHON")] string method = "EV",
        [ExcelArgument(Description = "EV scale: NONCONSTANT (default) or CONSTANT")] string scale = "NONCONSTANT",
        [ExcelArgument(Description = "EV bootstrap distribution: NONPARAMETRIC, GAMMA, or LOGNORMAL")] string bootstrapDistribution = "NONPARAMETRIC",
        [ExcelArgument(Description = "EV forecast distribution: GAMMA, LOGNORMAL, or NONPARAMETRIC")] string forecastDistribution = "GAMMA",
        [ExcelArgument(Description = "Output: RESERVES, ULTIMATES, PSEUDO-LRS, CUMULATIVES, or COMPLETE-CUMULATIVES")] string output = "RESERVES",
        [ExcelArgument(Description = "Optional (n-1) x (n-1) EV link-ratio inclusion mask")] object? mask = null,
        [ExcelArgument(Description = "Optional n-value EV forecast sqrt-scale vector")] object? userSqrtScale = null)
    {
        string normalizedMethod = Normalize(method, "EV");
        string normalizedOutput = Normalize(output, "RESERVES");
        if (normalizedOutput is not ("RESERVES" or "ULTIMATES" or "PSEUDO-LRS" or "CUMULATIVES" or "COMPLETE-CUMULATIVES"))
            return Error("Error: output must be RESERVES, ULTIMATES, PSEUDO-LRS, CUMULATIVES, or COMPLETE-CUMULATIVES");
        if (normalizedMethod == "CHAINLADDER-PYTHON" && normalizedOutput != "RESERVES")
            return Error("Error: CHAINLADDER-PYTHON supports RESERVES output only");

        if (!ValidateCommonInputs(triangle, iterations, normalizedMethod, out string error))
            return Error(error);
        if (!ValidateRawOutputSize(triangle.GetLength(0), iterations, normalizedOutput, out error))
            return Error(error);
        if (!TryResolveSeed(seed, out uint resolvedSeed, out error))
            return Error(error);

        if (normalizedMethod == "CHAINLADDER-PYTHON")
        {
            if (!TryGenerateLegacySamples(triangle, iterations, (int)resolvedSeed, out var legacyReserves, out var legacyTotals, out error))
                return Error(error);
            return FormatReserveSamples(legacyReserves, legacyTotals);
        }

        if (!TryParseReferenceOptions(
                triangle.GetLength(0), scale, bootstrapDistribution, forecastDistribution,
                mask, userSqrtScale,
                out string normalizedScale, out string normalizedBootstrap,
                out string normalizedForecast, out var parsedMask, out var parsedUserScale,
                out error))
            return Error(error);

        StochasticReservingOutputs requestedOutputs = normalizedOutput switch
        {
            "RESERVES" => StochasticReservingOutputs.Reserves | StochasticReservingOutputs.TotalReserves,
            "ULTIMATES" => StochasticReservingOutputs.Ultimates,
            "PSEUDO-LRS" => StochasticReservingOutputs.PseudoLinkRatios,
            "CUMULATIVES" => StochasticReservingOutputs.Cumulatives,
            "COMPLETE-CUMULATIVES" => StochasticReservingOutputs.CompleteCumulatives,
            _ => StochasticReservingOutputs.None
        };

        StochasticReservingBootstrapResult reference;
        try
        {
            reference = StochasticReservingBootstrap.Run(
                triangle, iterations, resolvedSeed, normalizedScale,
                normalizedBootstrap, normalizedForecast, parsedMask, parsedUserScale,
                requestedOutputs);
        }
        catch (ArgumentException exception)
        {
            return Error($"Error: {exception.Message}");
        }

        return normalizedOutput switch
        {
            "RESERVES" => FormatReserveSamples(reference.Reserves, reference.TotalReserves),
            "ULTIMATES" => FormatMatrixSamples(reference.Ultimates, "AY"),
            "PSEUDO-LRS" => FormatMatrixSamples(reference.PseudoLinkRatios, "Dev"),
            "CUMULATIVES" => FormatCubeSamples(reference.Cumulatives),
            "COMPLETE-CUMULATIVES" => FormatCubeSamples(reference.CompleteCumulatives),
            _ => throw new InvalidOperationException("Validated bootstrap output was not handled.")
        };
    }

    private static bool TryGenerateReserveSamples(
        double[,] triangle,
        int iterations,
        object? seed,
        string method,
        string scale,
        string bootstrapDistribution,
        string forecastDistribution,
        object? mask,
        object? userSqrtScale,
        out double[,] reserves,
        out double[] totals,
        out string error)
    {
        reserves = new double[0, 0];
        totals = Array.Empty<double>();
        string normalizedMethod = Normalize(method, "EV");
        if (!ValidateCommonInputs(triangle, iterations, normalizedMethod, out error))
            return false;
        if (!TryResolveSeed(seed, out uint resolvedSeed, out error))
            return false;

        if (normalizedMethod == "CHAINLADDER-PYTHON")
            return TryGenerateLegacySamples(triangle, iterations, (int)resolvedSeed, out reserves, out totals, out error);

        if (!TryParseReferenceOptions(
                triangle.GetLength(0), scale, bootstrapDistribution, forecastDistribution,
                mask, userSqrtScale,
                out string normalizedScale, out string normalizedBootstrap,
                out string normalizedForecast, out var parsedMask, out var parsedUserScale,
                out error))
            return false;

        try
        {
            var result = StochasticReservingBootstrap.Run(
                triangle, iterations, resolvedSeed, normalizedScale,
                normalizedBootstrap, normalizedForecast, parsedMask, parsedUserScale,
                StochasticReservingOutputs.Reserves | StochasticReservingOutputs.TotalReserves);
            reserves = result.Reserves;
            totals = result.TotalReserves;
            return true;
        }
        catch (ArgumentException exception)
        {
            error = $"Error: {exception.Message}";
            return false;
        }
    }

    private static bool TryGenerateLegacySamples(
        double[,] triangle,
        int iterations,
        int seed,
        out double[,] reserves,
        out double[] totals,
        out string error)
    {
        int n = triangle.GetLength(0);
        reserves = new double[iterations, n];
        totals = new double[iterations];
        var factors = GetFactorsArray(triangle);
        if (!PrepareBasicBootstrap(triangle, factors, n,
                out var fittedIncremental, out var pool, out double phiGlobal, out error))
            return false;

        var rng = new Random(seed);
        for (int iteration = 0; iteration < iterations; iteration++)
        {
            var sample = BootstrapOneIteration(
                triangle, fittedIncremental, pool, null, phiGlobal,
                isScaledPool: false, usePseudoDiag: false, n, rng);
            double total = 0.0;
            for (int origin = 0; origin < n; origin++)
            {
                reserves[iteration, origin] = sample[origin];
                total += sample[origin];
            }
            totals[iteration] = total;
        }
        return true;
    }

    private static bool TryParseReferenceOptions(
        int n,
        string scale,
        string bootstrapDistribution,
        string forecastDistribution,
        object? mask,
        object? userSqrtScale,
        out string normalizedScale,
        out string normalizedBootstrap,
        out string normalizedForecast,
        out double[,]? parsedMask,
        out double[]? parsedUserScale,
        out string error)
    {
        normalizedScale = Normalize(scale, "NONCONSTANT");
        normalizedBootstrap = Normalize(bootstrapDistribution, "NONPARAMETRIC");
        normalizedForecast = Normalize(forecastDistribution, "GAMMA");
        parsedMask = null;
        parsedUserScale = null;
        error = "";

        if (normalizedScale is not ("NONCONSTANT" or "CONSTANT"))
        {
            error = "Error: scale must be NONCONSTANT or CONSTANT";
            return false;
        }
        if (normalizedBootstrap is not ("NONPARAMETRIC" or "GAMMA" or "LOGNORMAL"))
        {
            error = "Error: bootstrap distribution must be NONPARAMETRIC, GAMMA, or LOGNORMAL";
            return false;
        }
        if (normalizedForecast is not ("NONPARAMETRIC" or "GAMMA" or "LOGNORMAL"))
        {
            error = "Error: forecast distribution must be NONPARAMETRIC, GAMMA, or LOGNORMAL";
            return false;
        }
        if (!TryParseMatrix(mask, n - 1, n - 1, "mask", out parsedMask, out error))
            return false;
        if (!TryParseVector(userSqrtScale, n, "user sqrt scale", out parsedUserScale, out error))
            return false;
        if (parsedMask is not null)
        {
            foreach (double value in parsedMask)
            {
                if ((value != 0.0 && value != 1.0) || double.IsNaN(value) || double.IsInfinity(value))
                {
                    error = "Error: mask values must be 0 or 1";
                    return false;
                }
            }
        }
        if (parsedUserScale is not null)
        {
            foreach (double value in parsedUserScale)
            {
                if (value < 0.0 || double.IsNaN(value) || double.IsInfinity(value))
                {
                    error = "Error: user sqrt scale values must be finite and non-negative";
                    return false;
                }
            }
        }
        return true;
    }

    private static bool ValidateCommonInputs(double[,] triangle, int iterations, string method, out string error)
    {
        error = "";
        int n = triangle.GetLength(0);
        if (triangle.GetLength(1) != n)
        {
            error = "Error: Triangle must be square";
            return false;
        }
        if (n < 3)
        {
            error = "Error: Triangle must have at least three rows";
            return false;
        }
        if (iterations <= 0)
        {
            error = "Error: Iterations must be positive";
            return false;
        }
        if (iterations > MaxBootstrapIterations)
        {
            error = $"Error: Iterations cannot exceed {MaxBootstrapIterations:N0}";
            return false;
        }
        if ((long)iterations * (n + 1L) > MaxStoredSampleValues)
        {
            error = "Error: Requested simulation is too large for the in-memory sample limit";
            return false;
        }
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n - i; j++)
        {
            double value = triangle[i, j];
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                error = "Error: Observed triangle cells must be finite";
                return false;
            }
        }
        if (method is not ("EV" or "CHAINLADDER-PYTHON"))
        {
            error = "Error: method must be EV or CHAINLADDER-PYTHON";
            return false;
        }
        return true;
    }

    private static bool TryResolveSeed(object? seed, out uint resolvedSeed, out string error)
    {
        if (!SeedUtil.TryResolveSeed(seed, out int? parsedSeed, out error))
        {
            resolvedSeed = 0;
            return false;
        }
        int value = parsedSeed ?? new Random().Next(0, int.MaxValue);
        resolvedSeed = (uint)value;
        error = "";
        return true;
    }

    private static bool ValidateRawOutputSize(
        int n, int iterations, string output, out string error)
    {
        long columns = output switch
        {
            "RESERVES" => n + 2L,
            "ULTIMATES" => n + 1L,
            "PSEUDO-LRS" => n,
            "CUMULATIVES" or "COMPLETE-CUMULATIVES" => (long)n * n + 1L,
            _ => 1L
        };
        long rows = iterations + 1L;
        if (rows > 1_048_576L || columns > 16_384L)
        {
            error = "Error: Requested output exceeds Excel worksheet dimensions";
            return false;
        }
        if (rows * columns > MaxRawOutputCells)
        {
            error = $"Error: Raw bootstrap output cannot exceed {MaxRawOutputCells:N0} cells";
            return false;
        }
        error = "";
        return true;
    }

    private static bool TryParseMatrix(
        object? value, int rows, int columns, string name,
        out double[,]? parsed, out string error)
    {
        parsed = null;
        error = "";
        if (IsMissing(value))
            return true;
        if (value is double[,] doubles)
        {
            if (doubles.GetLength(0) != rows || doubles.GetLength(1) != columns)
            {
                error = $"Error: {name} must be {rows} x {columns}";
                return false;
            }
            parsed = (double[,])doubles.Clone();
            return true;
        }
        if (value is object[,] objects)
        {
            if (objects.GetLength(0) != rows || objects.GetLength(1) != columns)
            {
                error = $"Error: {name} must be {rows} x {columns}";
                return false;
            }
            parsed = new double[rows, columns];
            try
            {
                for (int i = 0; i < rows; i++)
                for (int j = 0; j < columns; j++)
                    parsed[i, j] = Convert.ToDouble(objects[i, j]);
            }
            catch
            {
                error = $"Error: {name} must contain numeric values";
                return false;
            }
            return true;
        }
        error = $"Error: {name} must be a range";
        return false;
    }

    private static bool TryParseVector(
        object? value, int length, string name,
        out double[]? parsed, out string error)
    {
        parsed = null;
        error = "";
        if (IsMissing(value))
            return true;
        if (value is double[] doubles)
        {
            if (doubles.Length != length)
            {
                error = $"Error: {name} must contain {length} values";
                return false;
            }
            parsed = (double[])doubles.Clone();
            return true;
        }
        if (value is double[,] matrix)
        {
            if (matrix.Length != length || (matrix.GetLength(0) != 1 && matrix.GetLength(1) != 1))
            {
                error = $"Error: {name} must be a row or column of {length} values";
                return false;
            }
            parsed = new double[length];
            for (int i = 0; i < length; i++)
                parsed[i] = matrix.GetLength(0) == 1 ? matrix[0, i] : matrix[i, 0];
            return true;
        }
        if (value is object[,] objects)
        {
            if (objects.Length != length || (objects.GetLength(0) != 1 && objects.GetLength(1) != 1))
            {
                error = $"Error: {name} must be a row or column of {length} values";
                return false;
            }
            parsed = new double[length];
            try
            {
                for (int i = 0; i < length; i++)
                    parsed[i] = Convert.ToDouble(objects.GetLength(0) == 1 ? objects[0, i] : objects[i, 0]);
            }
            catch
            {
                error = $"Error: {name} must contain numeric values";
                return false;
            }
            return true;
        }
        error = $"Error: {name} must be a range";
        return false;
    }

    private static bool IsMissing(object? value)
        => value is null or ExcelMissing or ExcelEmpty;

    private static object[,] FormatReserveSamples(double[,] reserves, double[] totals)
    {
        int iterations = reserves.GetLength(0);
        int n = reserves.GetLength(1);
        var result = new object[iterations + 1, n + 2];
        result[0, 0] = "Iteration";
        result[0, 1] = "TotalReserve";
        for (int i = 0; i < n; i++)
            result[0, i + 2] = $"AY{i + 1}";
        for (int iteration = 0; iteration < iterations; iteration++)
        {
            result[iteration + 1, 0] = iteration + 1;
            result[iteration + 1, 1] = totals[iteration];
            for (int i = 0; i < n; i++)
                result[iteration + 1, i + 2] = reserves[iteration, i];
        }
        return result;
    }

    private static object[,] FormatMatrixSamples(double[,] values, string prefix)
    {
        int iterations = values.GetLength(0);
        int columns = values.GetLength(1);
        var result = new object[iterations + 1, columns + 1];
        result[0, 0] = "Iteration";
        for (int column = 0; column < columns; column++)
            result[0, column + 1] = $"{prefix}{column + 1}";
        for (int iteration = 0; iteration < iterations; iteration++)
        {
            result[iteration + 1, 0] = iteration + 1;
            for (int column = 0; column < columns; column++)
                result[iteration + 1, column + 1] = values[iteration, column];
        }
        return result;
    }

    private static object[,] FormatCubeSamples(double[,,] values)
    {
        int iterations = values.GetLength(0);
        int n = values.GetLength(1);
        var result = new object[iterations + 1, n * n + 1];
        result[0, 0] = "Iteration";
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n; j++)
            result[0, 1 + i * n + j] = $"AY{i + 1}-Dev{j + 1}";
        for (int iteration = 0; iteration < iterations; iteration++)
        {
            result[iteration + 1, 0] = iteration + 1;
            for (int i = 0; i < n; i++)
            for (int j = 0; j < n; j++)
                result[iteration + 1, 1 + i * n + j] = values[iteration, i, j];
        }
        return result;
    }

    private static double PopulationStandardDeviation(IEnumerable<double> values, double mean)
        => Math.Sqrt(values.Select(value => (value - mean) * (value - mean)).Average());

    private static double ReferencePercentile(double[] sortedValues, double percentile)
        => sortedValues[Math.Min((int)(percentile * sortedValues.Length), sortedValues.Length - 1)];

    private static string Normalize(string? value, string fallback)
        => string.IsNullOrWhiteSpace(value) ? fallback : value!.Trim().ToUpperInvariant();

    private static object[,] Error(string message)
        => new object[,] { { message } };
}
