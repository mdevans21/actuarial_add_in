namespace ActuarialAddIn.Functions;

internal sealed class StochasticReservingBootstrapResult
{
    public double[,] PseudoLinkRatios { get; init; } = null!;
    public double[,] Reserves { get; init; } = null!;
    public double[,] Ultimates { get; init; } = null!;
    public double[] TotalReserves { get; init; } = null!;
    public double[,,] Cumulatives { get; init; } = null!;
    public double[,,] CompleteCumulatives { get; init; } = null!;
}

/// <summary>
/// Pathwise port of Python_Examples/StochResFunctions.py Main_ODP_Bstrap.
/// Array traversal and RNG consumption intentionally follow the Python code.
/// </summary>
internal static class StochasticReservingBootstrap
{
    private const double DistributionTolerance = 1e-12;

    public static StochasticReservingBootstrapResult Run(
        double[,] inputTriangle,
        int iterations,
        uint seed,
        string scale,
        string bootstrapDistribution,
        string forecastDistribution,
        double[,]? inputMask = null,
        double[]? userSqrtScale = null)
    {
        int n = inputTriangle.GetLength(0);
        if (n < 3 || inputTriangle.GetLength(1) != n)
            throw new ArgumentException("Triangle must be square with at least three rows.", nameof(inputTriangle));
        if (iterations <= 0)
            throw new ArgumentOutOfRangeException(nameof(iterations));

        var triangle = ToReferenceTriangle(inputTriangle);
        var mask = inputMask is null ? CreateDefaultMask(n) : ValidateAndCopyMask(inputMask, n);
        var factors = LinkRatioFactors(triangle, mask);
        var residuals = CalculateResiduals(triangle, factors, mask, scale);
        var residualPool = FlattenAndSort(residuals.ZeroAverageAdjustedScaled);
        if (residualPool.Length == 0)
            throw new ArgumentException("Not enough residuals for bootstrap.", nameof(inputTriangle));

        double[] forecastSqrtScale;
        if (userSqrtScale is null)
        {
            forecastSqrtScale = residuals.SqrtScale;
        }
        else
        {
            if (userSqrtScale.Length != n)
                throw new ArgumentException($"User sqrt scale must contain {n} values.", nameof(userSqrtScale));
            forecastSqrtScale = (double[])userSqrtScale.Clone();
        }

        var rng = new NumpyRandomState(seed);
        var pseudoLinkRatios = new double[iterations, n - 1];
        var reserves = new double[iterations, n];
        var ultimates = new double[iterations, n];
        var totalReserves = new double[iterations];
        var cumulatives = new double[iterations, n, n];
        var completeCumulatives = new double[iterations, n, n];

        for (int iteration = 0; iteration < iterations; iteration++)
        {
            var resampledResiduals = ResampleResiduals(residualPool, n, rng);
            var pseudoIncrementals = bootstrapDistribution switch
            {
                "GAMMA" => CreateParametricPseudoData(residuals.FittedIncremental, residuals.SqrtScale, rng, lognormal: false),
                "LOGNORMAL" => CreateParametricPseudoData(residuals.FittedIncremental, residuals.SqrtScale, rng, lognormal: true),
                _ => CreateNonParametricPseudoData(resampledResiduals, residuals.SqrtScale, residuals.FittedIncremental)
            };
            var pseudoCumulatives = CumulativeSum(pseudoIncrementals);
            var iterationFactors = LinkRatioFactors(pseudoCumulatives, mask);
            for (int j = 0; j < n - 1; j++)
                pseudoLinkRatios[iteration, j] = iterationFactors[j];

            var forecast = Forecast(
                triangle,
                iterationFactors,
                resampledResiduals,
                pseudoCumulatives,
                forecastSqrtScale,
                forecastDistribution,
                rng);

            double total = 0.0;
            for (int i = 0; i < n; i++)
            {
                reserves[iteration, i] = forecast.Reserves[i];
                ultimates[iteration, i] = forecast.Ultimates[i];
                if (!double.IsNaN(forecast.Reserves[i]))
                    total += forecast.Reserves[i];
                for (int j = 0; j < n; j++)
                {
                    cumulatives[iteration, i, j] = forecast.Cumulatives[i, j];
                    completeCumulatives[iteration, i, j] = forecast.Cumulatives[i, j];
                }
            }
            totalReserves[iteration] = total;
        }

        return new StochasticReservingBootstrapResult
        {
            PseudoLinkRatios = pseudoLinkRatios,
            Reserves = reserves,
            Ultimates = ultimates,
            TotalReserves = totalReserves,
            Cumulatives = cumulatives,
            CompleteCumulatives = completeCumulatives
        };
    }

    private static double[,] ToReferenceTriangle(double[,] input)
    {
        int n = input.GetLength(0);
        var result = CreateNaNMatrix(n, n);
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n - i; j++)
                result[i, j] = input[i, j];
        }
        return result;
    }

    private static double[,] CreateDefaultMask(int n)
    {
        var mask = new double[n - 1, n - 1];
        for (int i = 0; i < n - 1; i++)
        for (int j = 0; j < n - 1; j++)
            mask[i, j] = 1.0;
        return mask;
    }

    private static double[,] ValidateAndCopyMask(double[,] input, int n)
    {
        if (input.GetLength(0) != n - 1 || input.GetLength(1) != n - 1)
            throw new ArgumentException($"Exclusion mask must be {n - 1} x {n - 1}.", nameof(input));
        return (double[,])input.Clone();
    }

    private static double[] LinkRatioFactors(double[,] triangle, double[,] mask)
    {
        int n = triangle.GetLength(0);
        var factors = new double[n - 1];
        for (int j = 0; j < n - 1; j++)
        {
            double weightedRatios = 0.0;
            double weights = 0.0;
            for (int i = 0; i < n - 1; i++)
            {
                double numerator = triangle[i, j + 1];
                if (double.IsNaN(numerator))
                    continue;

                double denominator = triangle[i, j];
                double ratio = denominator == 0.0 ? 0.0 : numerator / denominator;
                if (double.IsInfinity(ratio) || double.IsNaN(ratio))
                    continue;

                double maskedWeight = denominator * mask[i, j];
                if (double.IsNaN(maskedWeight))
                    continue;
                weightedRatios += ratio * maskedWeight;
                weights += maskedWeight;
            }
            factors[j] = weights == 0.0 ? 1.0 : weightedRatios / weights;
            if (double.IsNaN(factors[j]))
                factors[j] = 1.0;
        }
        return factors;
    }

    private static ResidualResult CalculateResiduals(
        double[,] triangle,
        double[] factors,
        double[,] mask,
        string scale)
    {
        int n = triangle.GetLength(0);
        var fittedCumulative = CreateNaNMatrix(n, n);
        for (int i = 0; i < n - 1; i++)
        {
            int diagonal = n - i - 1;
            fittedCumulative[i, diagonal] = triangle[i, diagonal];
            for (int j = diagonal - 1; j >= 0; j--)
                fittedCumulative[i, j] = fittedCumulative[i, j + 1] / factors[j];
        }
        fittedCumulative[n - 1, 0] = triangle[n - 1, 0];

        var fittedIncremental = Incrementals(fittedCumulative);
        var observedIncremental = Incrementals(triangle);
        var indicators = new double[n, n];
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n; j++)
            indicators[i, j] = 1.0;

        var maskForCounts = (double[,])mask.Clone();
        for (int i = 0; i < n - 1; i++)
        for (int j = 0; j < n - 1; j++)
        {
            double numerator = triangle[i, j + 1];
            double denominator = triangle[i, j];
            double ratio = denominator == 0.0 ? 0.0 : numerator / denominator;
            if (!double.IsNaN(numerator) && ratio == 0.0)
                maskForCounts[i, j] = 0.0;
        }

        for (int i = 0; i < n - 1; i++)
        for (int j = 0; j < n - i; j++)
            indicators[i, j] = j <= 1 ? maskForCounts[i, 0] : maskForCounts[i, j - 1];

        var observationsByDevelopment = new double[n];
        double observations = 0.0;
        for (int j = 0; j < n; j++)
        {
            for (int i = 0; i < n - j; i++)
                observationsByDevelopment[j] += indicators[i, j];
            observations += observationsByDevelopment[j];
        }

        int parameters = 2 * n - 1;
        if (observations <= parameters)
            throw new ArgumentException("Not enough degrees of freedom for bootstrap.", nameof(triangle));
        double bias = Math.Sqrt(observations / (observations - parameters));

        var adjustedUnscaled = CreateNaNMatrix(n, n);
        var scaleSquares = CreateNaNMatrix(n, n);
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n - i; j++)
        {
            double fitted = fittedIncremental[i, j];
            double unscaled = fitted == 0.0
                ? 0.0
                : indicators[i, j] * (observedIncremental[i, j] - fitted) / Math.Sqrt(Math.Abs(fitted));
            if (unscaled == 0.0)
                continue;
            adjustedUnscaled[i, j] = bias * unscaled;
            scaleSquares[i, j] = adjustedUnscaled[i, j] * adjustedUnscaled[i, j];
        }

        var sqrtScale = new double[n];
        if (scale == "CONSTANT")
        {
            double sum = SumIgnoringNaN(scaleSquares);
            double value = Math.Sqrt(sum / observations);
            for (int j = 0; j < n; j++)
                sqrtScale[j] = value;
        }
        else
        {
            for (int j = 0; j < n - 1; j++)
            {
                if (observationsByDevelopment[j] <= 1.0)
                {
                    sqrtScale[j] = j == 0 ? 0.0 : sqrtScale[j - 1];
                }
                else
                {
                    double sum = 0.0;
                    for (int i = 0; i < n - j; i++)
                    {
                        if (!double.IsNaN(scaleSquares[i, j]))
                            sum += scaleSquares[i, j];
                    }
                    sqrtScale[j] = Math.Sqrt(sum / observationsByDevelopment[j]);
                }
            }
            sqrtScale[n - 1] = Math.Min(sqrtScale[n - 2], sqrtScale[n - 3]);
        }

        var reverseProducts = new double[n - 1];
        double reverseProduct = 1.0;
        for (int j = n - 2; j >= 0; j--)
        {
            reverseProduct *= factors[j];
            reverseProducts[j] = reverseProduct;
        }
        if (reverseProducts[0] == 1.0)
            sqrtScale[0] = 0.0;
        for (int j = 1; j < n; j++)
        {
            if (reverseProducts[j - 1] == 1.0)
                sqrtScale[j] = 0.0;
        }

        var adjustedScaled = CreateNaNMatrix(n, n);
        double adjustedScaledSum = 0.0;
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n - i; j++)
        {
            if (!double.IsNaN(adjustedUnscaled[i, j]) && sqrtScale[j] != 0.0)
            {
                adjustedScaled[i, j] = adjustedUnscaled[i, j] / sqrtScale[j];
                adjustedScaledSum += adjustedScaled[i, j];
            }
        }

        double averageResidual = adjustedScaledSum / (observations - 2.0);
        var zeroAverage = CreateNaNMatrix(n, n);
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n - i; j++)
        {
            if (!double.IsNaN(adjustedScaled[i, j]))
                zeroAverage[i, j] = adjustedScaled[i, j] - averageResidual;
        }
        zeroAverage[0, n - 1] = double.NaN;
        zeroAverage[n - 1, 0] = double.NaN;

        return new ResidualResult(fittedIncremental, sqrtScale, zeroAverage);
    }

    private static double[,] Incrementals(double[,] cumulatives)
    {
        int n = cumulatives.GetLength(0);
        var incrementals = CreateNaNMatrix(n, n);
        for (int i = 0; i < n; i++)
        {
            incrementals[i, 0] = cumulatives[i, 0];
            for (int j = 1; j < n; j++)
            {
                if (!double.IsNaN(cumulatives[i, j]))
                {
                    double previous = double.IsNaN(cumulatives[i, j - 1]) ? 0.0 : cumulatives[i, j - 1];
                    incrementals[i, j] = cumulatives[i, j] - previous;
                }
            }
        }
        return incrementals;
    }

    private static double[] FlattenAndSort(double[,] matrix)
    {
        var values = new List<double>();
        for (int i = 0; i < matrix.GetLength(0); i++)
        for (int j = 0; j < matrix.GetLength(1); j++)
        {
            if (!double.IsNaN(matrix[i, j]))
                values.Add(matrix[i, j]);
        }
        values.Sort();
        return values.ToArray();
    }

    private static double[,] ResampleResiduals(double[] pool, int n, NumpyRandomState rng)
    {
        var result = new double[n, n];
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n; j++)
            result[i, j] = pool[rng.NextIndex(pool.Length)];
        return result;
    }

    private static double[,] CreateNonParametricPseudoData(
        double[,] residuals,
        double[] sqrtScale,
        double[,] fittedIncremental)
    {
        int n = residuals.GetLength(0);
        var result = CreateNaNMatrix(n, n);
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n - i; j++)
        {
            double fitted = fittedIncremental[i, j];
            result[i, j] = residuals[i, j] * sqrtScale[j] * Math.Sqrt(Math.Abs(fitted)) + fitted;
        }
        return result;
    }

    private static double[,] CreateParametricPseudoData(
        double[,] fittedIncremental,
        double[] sqrtScale,
        NumpyRandomState rng,
        bool lognormal)
    {
        int n = fittedIncremental.GetLength(0);
        var result = CreateNaNMatrix(n, n);
        for (int i = 0; i < n; i++)
        for (int j = 0; j < n - i; j++)
        {
            double mean = fittedIncremental[i, j];
            double standardDeviation = sqrtScale[j] * Math.Sqrt(Math.Abs(mean));
            result[i, j] = DrawWithMoments(mean, standardDeviation, rng, lognormal);
        }
        return result;
    }

    private static double[,] CumulativeSum(double[,] incrementals)
    {
        int n = incrementals.GetLength(0);
        var result = CreateNaNMatrix(n, n);
        for (int i = 0; i < n; i++)
        {
            double sum = 0.0;
            bool missing = false;
            for (int j = 0; j < n; j++)
            {
                if (missing || double.IsNaN(incrementals[i, j]))
                {
                    missing = true;
                    continue;
                }
                sum += incrementals[i, j];
                result[i, j] = sum;
            }
        }
        return result;
    }

    private static ForecastResult Forecast(
        double[,] triangle,
        double[] factors,
        double[,] residuals,
        double[,] pseudoCumulatives,
        double[] sqrtScale,
        string distribution,
        NumpyRandomState rng)
    {
        int n = triangle.GetLength(0);
        var cumulatives = (double[,])triangle.Clone();
        for (int i = 1; i < n; i++)
        {
            for (int j = n - i; j < n; j++)
            {
                pseudoCumulatives[i, j] = pseudoCumulatives[i, j - 1] * factors[j - 1];
                double mean = pseudoCumulatives[i, j] - pseudoCumulatives[i, j - 1];
                double incremental;
                if (distribution == "GAMMA")
                {
                    double sd = sqrtScale[j] * Math.Sqrt(Math.Abs(mean));
                    incremental = DrawWithMoments(mean, sd, rng, lognormal: false);
                }
                else if (distribution == "LOGNORMAL")
                {
                    double sd = sqrtScale[j] * Math.Sqrt(Math.Abs(mean));
                    incremental = DrawWithMoments(mean, sd, rng, lognormal: true);
                }
                else
                {
                    incremental = residuals[i, j] * sqrtScale[j] * Math.Sqrt(Math.Abs(mean)) + mean;
                }
                cumulatives[i, j] = cumulatives[i, j - 1] + incremental;
            }
        }

        var reserves = new double[n];
        var ultimates = new double[n];
        for (int i = 0; i < n; i++)
        {
            ultimates[i] = cumulatives[i, n - 1];
            reserves[i] = ultimates[i] - triangle[i, n - i - 1];
        }
        return new ForecastResult(cumulatives, reserves, ultimates);
    }

    private static double DrawWithMoments(
        double mean,
        double standardDeviation,
        NumpyRandomState rng,
        bool lognormal)
    {
        if (mean > DistributionTolerance)
        {
            if (!lognormal && standardDeviation < DistributionTolerance)
                return mean;
            if (lognormal)
            {
                double sigmaNormal = Math.Sqrt(Math.Log(1.0 + Math.Pow(standardDeviation / mean, 2.0)));
                double meanNormal = Math.Log(mean) - 0.5 * sigmaNormal * sigmaNormal;
                return rng.Lognormal(meanNormal, sigmaNormal);
            }
            double scale = standardDeviation * standardDeviation / mean;
            double shape = mean / scale;
            return rng.Gamma(shape, scale);
        }
        return rng.Normal(mean, standardDeviation);
    }

    private static double SumIgnoringNaN(double[,] matrix)
    {
        double sum = 0.0;
        foreach (double value in matrix)
        {
            if (!double.IsNaN(value))
                sum += value;
        }
        return sum;
    }

    private static double[,] CreateNaNMatrix(int rows, int columns)
    {
        var result = new double[rows, columns];
        for (int i = 0; i < rows; i++)
        for (int j = 0; j < columns; j++)
            result[i, j] = double.NaN;
        return result;
    }

    private sealed record ResidualResult(
        double[,] FittedIncremental,
        double[] SqrtScale,
        double[,] ZeroAverageAdjustedScaled);

    private sealed record ForecastResult(
        double[,] Cumulatives,
        double[] Reserves,
        double[] Ultimates);
}
