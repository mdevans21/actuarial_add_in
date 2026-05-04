using System;
using System.Linq;
using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Cat modeling utilities for transforming an Event Loss Table (ELT) into a Year Loss Table (YLT)
/// and generating OEP/AEP exceedance probability curves.
/// </summary>
public static class CatModeling
{
    [ExcelFunction(Description = "Simulate a Year Loss Table (YLT) from an Event Loss Table (ELT). Returns columns: Year, Aggregate Loss, Max Loss, Event Count.", Category = "Actuarial.Reinsurance")]
    public static object[,] ACT_CAT_ELT_TO_YLT(
        [ExcelArgument(Description = "Event annual rates (column)")] double[] eventRates,
        [ExcelArgument(Description = "Event losses (column)")] double[] eventLosses,
        [ExcelArgument(Description = "Number of simulation years")] int years,
        [ExcelArgument(Description = "Random seed; leave blank for system seed")] object? seed = null,
        [ExcelArgument(Description = "Include header row")] bool includeHeader = true)
    {
        if (eventRates.Length != eventLosses.Length || eventRates.Length == 0)
            return new object[,] { { "Error: Rates and losses must be same length" } };

        if (years <= 0)
            return new object[,] { { "Error: Years must be positive" } };

        for (int i = 0; i < eventRates.Length; i++)
        {
            if (eventRates[i] < 0 || eventLosses[i] < 0)
                return new object[,] { { "Error: Rates and losses must be non-negative" } };
        }

        var rng = SeedUtil.ResolveSeed(seed) is { } _seed ? new Random(_seed) : new Random();
        int rows = years + (includeHeader ? 1 : 0);
        var result = new object[rows, 4];
        int row = 0;

        if (includeHeader)
        {
            result[row, 0] = "Year";
            result[row, 1] = "Aggregate Loss";
            result[row, 2] = "Max Loss";
            result[row, 3] = "Event Count";
            row++;
        }

        for (int year = 0; year < years; year++)
        {
            double aggregateLoss = 0.0;
            double maxLoss = 0.0;
            int eventCount = 0;

            for (int i = 0; i < eventRates.Length; i++)
            {
                double rate = eventRates[i];
                if (rate <= 0) continue;

                int occurrences = SamplePoisson(rng, rate);
                if (occurrences <= 0) continue;

                double loss = eventLosses[i];
                aggregateLoss += occurrences * loss;
                eventCount += occurrences;
                if (loss > maxLoss) maxLoss = loss;
            }

            result[row, 0] = year + 1;
            result[row, 1] = aggregateLoss;
            result[row, 2] = maxLoss;
            result[row, 3] = eventCount;
            row++;
        }

        return result;
    }

    [ExcelFunction(Description = "Generate an Occurrence Exceedance Probability (OEP) curve from annual maximum losses.", Category = "Actuarial.Reinsurance")]
    public static object[,] ACT_CAT_YLT_OEP_CURVE(
        [ExcelArgument(Description = "Annual maximum losses (column)")] double[] annualMaxLosses,
        [ExcelArgument(Description = "Plotting position method: WEIBULL, HAZEN, GRINGORTEN, MEDIAN")] string plottingPosition = "WEIBULL",
        [ExcelArgument(Description = "Include header row")] bool includeHeader = true)
    {
        return BuildEpCurve(annualMaxLosses, plottingPosition, includeHeader, "OEP Loss");
    }

    [ExcelFunction(Description = "Generate an Aggregate Exceedance Probability (AEP) curve from annual aggregate losses.", Category = "Actuarial.Reinsurance")]
    public static object[,] ACT_CAT_YLT_AEP_CURVE(
        [ExcelArgument(Description = "Annual aggregate losses (column)")] double[] annualAggregateLosses,
        [ExcelArgument(Description = "Plotting position method: WEIBULL, HAZEN, GRINGORTEN, MEDIAN")] string plottingPosition = "WEIBULL",
        [ExcelArgument(Description = "Include header row")] bool includeHeader = true)
    {
        return BuildEpCurve(annualAggregateLosses, plottingPosition, includeHeader, "AEP Loss");
    }

    [ExcelFunction(Description = "Generate an OEP curve at specified return periods using empirical quantiles.", Category = "Actuarial.Reinsurance")]
    public static object[,] ACT_CAT_OEP_CURVE_RP(
        [ExcelArgument(Description = "Annual maximum losses (column)")] double[] annualMaxLosses,
        [ExcelArgument(Description = "Return periods (column)")] double[] returnPeriods,
        [ExcelArgument(Description = "Include header row")] bool includeHeader = true)
    {
        return BuildEpCurveAtReturnPeriods(annualMaxLosses, returnPeriods, includeHeader, "OEP Loss");
    }

    [ExcelFunction(Description = "Generate an AEP curve at specified return periods using empirical quantiles.", Category = "Actuarial.Reinsurance")]
    public static object[,] ACT_CAT_AEP_CURVE_RP(
        [ExcelArgument(Description = "Annual aggregate losses (column)")] double[] annualAggregateLosses,
        [ExcelArgument(Description = "Return periods (column)")] double[] returnPeriods,
        [ExcelArgument(Description = "Include header row")] bool includeHeader = true)
    {
        return BuildEpCurveAtReturnPeriods(annualAggregateLosses, returnPeriods, includeHeader, "AEP Loss");
    }

    [ExcelFunction(Description = "Calculate Value-at-Risk (VaR) from samples at confidence level alpha.", Category = "Actuarial.Reinsurance")]
    public static double ACT_VAR_FROM_SAMPLES(
        [ExcelArgument(Description = "Sample losses (column)")] double[] samples,
        [ExcelArgument(Description = "Confidence level (e.g., 0.99)")] double alpha)
    {
        if (samples.Length == 0 || alpha <= 0 || alpha >= 1)
            return double.NaN;

        var sorted = samples.OrderBy(x => x).ToArray();
        int index = (int)Math.Ceiling(alpha * sorted.Length) - 1;
        index = Math.Max(0, Math.Min(index, sorted.Length - 1));
        return sorted[index];
    }

    [ExcelFunction(Description = "Calculate Tail Value-at-Risk (TVaR) from samples at confidence level alpha.", Category = "Actuarial.Reinsurance")]
    public static double ACT_TVAR_FROM_SAMPLES(
        [ExcelArgument(Description = "Sample losses (column)")] double[] samples,
        [ExcelArgument(Description = "Confidence level (e.g., 0.99)")] double alpha)
    {
        if (samples.Length == 0 || alpha <= 0 || alpha >= 1)
            return double.NaN;

        var sorted = samples.OrderBy(x => x).ToArray();
        int index = (int)Math.Ceiling(alpha * sorted.Length) - 1;
        index = Math.Max(0, Math.Min(index, sorted.Length - 1));
        double varValue = sorted[index];
        var tail = sorted.Where(x => x >= varValue).ToArray();
        if (tail.Length == 0) return varValue;
        return tail.Average();
    }

    private static object[,] BuildEpCurve(double[] losses, string plottingPosition, bool includeHeader, string lossHeader)
    {
        var cleaned = losses.Where(loss => !double.IsNaN(loss) && !double.IsInfinity(loss)).ToArray();
        if (cleaned.Length == 0)
            return new object[,] { { "Error: Losses required" } };

        var sorted = cleaned.OrderByDescending(x => x).ToArray();
        int n = sorted.Length;
        int rows = n + (includeHeader ? 1 : 0);
        var result = new object[rows, 2];
        int row = 0;

        if (includeHeader)
        {
            result[row, 0] = "Return Period";
            result[row, 1] = lossHeader;
            row++;
        }

        for (int i = 0; i < n; i++)
        {
            int rank = i + 1;
            double ep = PlottingPosition(rank, n, plottingPosition);
            double returnPeriod = ep > 0 ? 1.0 / ep : double.PositiveInfinity;
            result[row, 0] = returnPeriod;
            result[row, 1] = sorted[i];
            row++;
        }

        return result;
    }

    private static object[,] BuildEpCurveAtReturnPeriods(double[] losses, double[] returnPeriods, bool includeHeader, string lossHeader)
    {
        var cleaned = losses.Where(loss => !double.IsNaN(loss) && !double.IsInfinity(loss)).ToArray();
        if (cleaned.Length == 0)
            return new object[,] { { "Error: Losses required" } };
        if (returnPeriods.Length == 0)
            return new object[,] { { "Error: Return periods required" } };

        var sorted = cleaned.OrderBy(x => x).ToArray();
        int rows = returnPeriods.Length + (includeHeader ? 1 : 0);
        var result = new object[rows, 2];
        int row = 0;

        if (includeHeader)
        {
            result[row, 0] = "Return Period";
            result[row, 1] = lossHeader;
            row++;
        }

        for (int i = 0; i < returnPeriods.Length; i++)
        {
            double rp = returnPeriods[i];
            if (rp <= 0)
            {
                result[row, 0] = rp;
                result[row, 1] = double.NaN;
                row++;
                continue;
            }
            double alpha = 1.0 - 1.0 / rp;
            result[row, 0] = rp;
            result[row, 1] = Quantile(sorted, alpha);
            row++;
        }

        return result;
    }

    private static double Quantile(double[] sorted, double alpha)
    {
        if (alpha <= 0) return sorted[0];
        if (alpha >= 1) return sorted[sorted.Length - 1];

        double index = alpha * (sorted.Length - 1);
        int lower = (int)Math.Floor(index);
        int upper = (int)Math.Ceiling(index);
        if (lower == upper) return sorted[lower];
        double weight = index - lower;
        return sorted[lower] * (1 - weight) + sorted[upper] * weight;
    }

    private static double PlottingPosition(int rank, int n, string method)
    {
        string normalized = (method ?? "WEIBULL").Trim().ToUpperInvariant();
        return normalized switch
        {
            "HAZEN" => (rank - 0.5) / n,
            "GRINGORTEN" => (rank - 0.44) / (n + 0.12),
            "MEDIAN" => (rank - 0.3) / (n + 0.4),
            _ => rank / (double)(n + 1)
        };
    }

    private static int SamplePoisson(Random rng, double lambda)
    {
        if (lambda <= 0) return 0;
        if (lambda < 30.0)
        {
            double limit = Math.Exp(-lambda);
            int k = 0;
            double p = 1.0;
            do
            {
                k++;
                p *= rng.NextDouble();
            } while (p > limit);
            return k - 1;
        }

        // Normal approximation for large lambda
        double normalSample = SampleStandardNormal(rng) * Math.Sqrt(lambda) + lambda;
        int result = (int)Math.Round(normalSample);
        return Math.Max(0, result);
    }

    private static double SampleStandardNormal(Random rng)
    {
        double u1 = rng.NextDouble();
        double u2 = rng.NextDouble();
        return Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Cos(2.0 * Math.PI * u2);
    }
}
