using System.Text.Json;
using System.Text.Json.Serialization;
using ActuarialAddIn.Functions;

namespace ActuarialAddIn.Tests;

// Emits a JSON snapshot of Add-In outputs over a systematic input grid. Consumed by
// tests/reconciliation.ipynb, which compares these values to independent references
// (scipy, chainladder, analytical formulae). Deterministic functions use fixed inputs;
// stochastic functions use a fixed seed.
public static class AddinOutputsEmitter
{
    private class Record
    {
        public string Group { get; set; } = "";
        public string Function { get; set; } = "";
        public object?[] Args { get; set; } = Array.Empty<object?>();
        public object? Result { get; set; }
    }

    private static readonly List<Record> _records = new();

    public static void Emit(string outputPath)
    {
        _records.Clear();

        EmitDistributions();
        EmitZeroTruncated();
        EmitZeroModified();
        EmitParetoExtended();
        EmitInverseGaussian();
        EmitLoglogistic();
        EmitCompositeDistributions();
        EmitLEV();
        EmitFitting();
        EmitAggregate();
        EmitExposureCurves();
        EmitReinsurance();
        EmitCatModeling();
        EmitInterpolation();
        EmitChainLadder();
        EmitBootstrap();
        EmitCopulas();

        var payload = new
        {
            generated_at = DateTime.UtcNow.ToString("o"),
            version = ActuarialAddIn.Functions.Version.ACT_VERSION(),
            record_count = _records.Count,
            records = _records
        };

        var options = new JsonSerializerOptions
        {
            WriteIndented = false,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.Never,
            NumberHandling = JsonNumberHandling.AllowNamedFloatingPointLiterals
        };

        File.WriteAllText(outputPath, JsonSerializer.Serialize(payload, options));
        Console.WriteLine($"Emitted {_records.Count} records to: {outputPath}");
    }

    // === Recording helpers ===

    private static void R(string group, string function, object?[] args, object? result)
    {
        _records.Add(new Record
        {
            Group = group,
            Function = function,
            Args = args.Select(NormalizeArg).ToArray(),
            Result = NormalizeResult(result)
        });
    }

    private static object? NormalizeArg(object? a) => a switch
    {
        double[,] m => To2DList(m),
        double[] v => v.Select(x => (object)x).ToList(),
        _ => a
    };

    private static object? NormalizeResult(object? r) => r switch
    {
        null => null,
        double d => JsonSafeDouble(d),
        int i => i,
        string s => s,
        bool b => b,
        double[] v => v.Select(JsonSafeDouble).ToList(),
        object[] arr => arr.Select(NormalizeResult).ToList(),
        object[,] m => To2DListNormalized(m),
        double[,] m2 => To2DList(m2),
        _ => r.ToString()
    };

    private static object JsonSafeDouble(double d)
    {
        if (double.IsNaN(d)) return "NaN";
        if (double.IsPositiveInfinity(d)) return "Infinity";
        if (double.IsNegativeInfinity(d)) return "-Infinity";
        return d;
    }

    private static List<List<double>> To2DList(double[,] m)
    {
        int rows = m.GetLength(0), cols = m.GetLength(1);
        var outer = new List<List<double>>(rows);
        for (int i = 0; i < rows; i++)
        {
            var row = new List<double>(cols);
            for (int j = 0; j < cols; j++) row.Add(m[i, j]);
            outer.Add(row);
        }
        return outer;
    }

    private static List<List<object?>> To2DListNormalized(object[,] m)
    {
        int rows = m.GetLength(0), cols = m.GetLength(1);
        var outer = new List<List<object?>>(rows);
        for (int i = 0; i < rows; i++)
        {
            var row = new List<object?>(cols);
            for (int j = 0; j < cols; j++) row.Add(NormalizeResult(m[i, j]));
            outer.Add(row);
        }
        return outer;
    }

    // === Distributions (continuous + basic discrete) ===

    private static void EmitDistributions()
    {
        const string g = "distributions";

        // Poisson (lambda=5) and (lambda=0.5)
        foreach (var lam in new[] { 0.5, 5.0 })
        {
            foreach (int k in new[] { 0, 1, 2, 5, 10 })
            {
                R(g, "ACT_DIST_POISSON_PDF", new object?[] { k, lam }, Distributions.ACT_DIST_POISSON_PDF(k, lam));
                R(g, "ACT_DIST_POISSON_CDF", new object?[] { (double)k, lam }, Distributions.ACT_DIST_POISSON_CDF(k, lam));
            }
            foreach (double p in new[] { 0.1, 0.5, 0.9 })
                R(g, "ACT_DIST_POISSON_INV", new object?[] { p, lam }, Distributions.ACT_DIST_POISSON_INV(p, lam));
        }

        // Negative Binomial (r=5, p=0.3)
        double nbR = 5, nbP = 0.3;
        foreach (int k in new[] { 0, 1, 5, 10, 20 })
        {
            R(g, "ACT_DIST_NEGBIN_PDF", new object?[] { k, nbR, nbP }, Distributions.ACT_DIST_NEGBIN_PDF(k, nbR, nbP));
            R(g, "ACT_DIST_NEGBIN_CDF", new object?[] { (double)k, nbR, nbP }, Distributions.ACT_DIST_NEGBIN_CDF(k, nbR, nbP));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_NEGBIN_INV", new object?[] { p, nbR, nbP }, Distributions.ACT_DIST_NEGBIN_INV(p, nbR, nbP));

        // Normal(0,1)
        foreach (double x in new[] { -2.0, -1.0, 0.0, 1.0, 2.0 })
        {
            R(g, "ACT_DIST_NORMAL_PDF", new object?[] { x, 0.0, 1.0 }, Distributions.ACT_DIST_NORMAL_PDF(x, 0, 1));
            R(g, "ACT_DIST_NORMAL_CDF", new object?[] { x, 0.0, 1.0 }, Distributions.ACT_DIST_NORMAL_CDF(x, 0, 1));
        }
        foreach (double p in new[] { 0.025, 0.5, 0.975 })
            R(g, "ACT_DIST_NORMAL_INV", new object?[] { p, 0.0, 1.0 }, Distributions.ACT_DIST_NORMAL_INV(p, 0, 1));

        // Lognormal(mu=0, sigma=1)
        foreach (double x in new[] { 0.5, 1.0, 2.0, 5.0 })
        {
            R(g, "ACT_DIST_LOGNORM_PDF", new object?[] { x, 0.0, 1.0 }, Distributions.ACT_DIST_LOGNORM_PDF(x, 0, 1));
            R(g, "ACT_DIST_LOGNORM_CDF", new object?[] { x, 0.0, 1.0 }, Distributions.ACT_DIST_LOGNORM_CDF(x, 0, 1));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_LOGNORM_INV", new object?[] { p, 0.0, 1.0 }, Distributions.ACT_DIST_LOGNORM_INV(p, 0, 1));

        // Gamma (alpha=2, beta=1)
        foreach (double x in new[] { 0.5, 1.0, 2.0, 5.0 })
        {
            R(g, "ACT_DIST_GAMMA_PDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_GAMMA_PDF(x, 2, 1));
            R(g, "ACT_DIST_GAMMA_CDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_GAMMA_CDF(x, 2, 1));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_GAMMA_INV", new object?[] { p, 2.0, 1.0 }, Distributions.ACT_DIST_GAMMA_INV(p, 2, 1));

        // Pareto I (alpha=2, xm=1)
        foreach (double x in new[] { 1.0, 1.5, 2.0, 5.0 })
        {
            R(g, "ACT_DIST_PARETO_PDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_PARETO_PDF(x, 2, 1));
            R(g, "ACT_DIST_PARETO_CDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_PARETO_CDF(x, 2, 1));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_PARETO_INV", new object?[] { p, 2.0, 1.0 }, Distributions.ACT_DIST_PARETO_INV(p, 2, 1));

        // GPD (xi=0.5, sigma=1)
        foreach (double x in new[] { 0.5, 1.0, 2.0, 5.0 })
        {
            R(g, "ACT_DIST_GPD_PDF", new object?[] { x, 0.5, 1.0 }, Distributions.ACT_DIST_GPD_PDF(x, 0.5, 1));
            R(g, "ACT_DIST_GPD_CDF", new object?[] { x, 0.5, 1.0 }, Distributions.ACT_DIST_GPD_CDF(x, 0.5, 1));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_GPD_INV", new object?[] { p, 0.5, 1.0 }, Distributions.ACT_DIST_GPD_INV(p, 0.5, 1));

        // Weibull (k=2, lambda=1)
        foreach (double x in new[] { 0.5, 1.0, 2.0 })
        {
            R(g, "ACT_DIST_WEIBULL_PDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_WEIBULL_PDF(x, 2, 1));
            R(g, "ACT_DIST_WEIBULL_CDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_WEIBULL_CDF(x, 2, 1));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_WEIBULL_INV", new object?[] { p, 2.0, 1.0 }, Distributions.ACT_DIST_WEIBULL_INV(p, 2, 1));

        // Beta (alpha=2, beta=5)
        foreach (double x in new[] { 0.1, 0.25, 0.5, 0.9 })
        {
            R(g, "ACT_DIST_BETA_PDF", new object?[] { x, 2.0, 5.0 }, Distributions.ACT_DIST_BETA_PDF(x, 2, 5));
            R(g, "ACT_DIST_BETA_CDF", new object?[] { x, 2.0, 5.0 }, Distributions.ACT_DIST_BETA_CDF(x, 2, 5));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_BETA_INV", new object?[] { p, 2.0, 5.0 }, Distributions.ACT_DIST_BETA_INV(p, 2, 5));

        // Exponential (lambda=2)
        foreach (double x in new[] { 0.1, 0.5, 1.0, 2.0 })
        {
            R(g, "ACT_DIST_EXP_PDF", new object?[] { x, 2.0 }, Distributions.ACT_DIST_EXP_PDF(x, 2));
            R(g, "ACT_DIST_EXP_CDF", new object?[] { x, 2.0 }, Distributions.ACT_DIST_EXP_CDF(x, 2));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_EXP_INV", new object?[] { p, 2.0 }, Distributions.ACT_DIST_EXP_INV(p, 2));

        // Burr XII (c=2, k=3, lambda=1)
        foreach (double x in new[] { 0.5, 1.0, 2.0 })
        {
            R(g, "ACT_DIST_BURR_PDF", new object?[] { x, 2.0, 3.0, 1.0 }, Distributions.ACT_DIST_BURR_PDF(x, 2, 3, 1));
            R(g, "ACT_DIST_BURR_CDF", new object?[] { x, 2.0, 3.0, 1.0 }, Distributions.ACT_DIST_BURR_CDF(x, 2, 3, 1));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_BURR_INV", new object?[] { p, 2.0, 3.0, 1.0 }, Distributions.ACT_DIST_BURR_INV(p, 2, 3, 1));

        // Lomax (alpha=2, lambda=1)
        foreach (double x in new[] { 0.5, 1.0, 2.0, 5.0 })
        {
            R(g, "ACT_DIST_LOMAX_PDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_LOMAX_PDF(x, 2, 1));
            R(g, "ACT_DIST_LOMAX_CDF", new object?[] { x, 2.0, 1.0 }, Distributions.ACT_DIST_LOMAX_CDF(x, 2, 1));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_LOMAX_INV", new object?[] { p, 2.0, 1.0 }, Distributions.ACT_DIST_LOMAX_INV(p, 2, 1));
    }

    private static void EmitZeroTruncated()
    {
        const string g = "zero_truncated";
        foreach (int k in new[] { 1, 2, 3, 5, 10 })
        {
            R(g, "ACT_DIST_ZTPOISSON_PDF", new object?[] { k, 5.0 }, Distributions.ACT_DIST_ZTPOISSON_PDF(k, 5));
            R(g, "ACT_DIST_ZTPOISSON_CDF", new object?[] { k, 5.0 }, Distributions.ACT_DIST_ZTPOISSON_CDF(k, 5));
            R(g, "ACT_DIST_ZTNEGBIN_PDF", new object?[] { k, 5.0, 0.3 }, Distributions.ACT_DIST_ZTNEGBIN_PDF(k, 5, 0.3));
            R(g, "ACT_DIST_ZTNEGBIN_CDF", new object?[] { k, 5.0, 0.3 }, Distributions.ACT_DIST_ZTNEGBIN_CDF(k, 5, 0.3));
            R(g, "ACT_DIST_ZTGEOM_PDF", new object?[] { k, 0.3 }, Distributions.ACT_DIST_ZTGEOM_PDF(k, 0.3));
            R(g, "ACT_DIST_ZTGEOM_CDF", new object?[] { k, 0.3 }, Distributions.ACT_DIST_ZTGEOM_CDF(k, 0.3));
        }
        foreach (int k in new[] { 1, 2, 3, 5, 10 })
        {
            R(g, "ACT_DIST_ZTBINOM_PDF", new object?[] { k, 10, 0.3 }, Distributions.ACT_DIST_ZTBINOM_PDF(k, 10, 0.3));
            R(g, "ACT_DIST_ZTBINOM_CDF", new object?[] { k, 10, 0.3 }, Distributions.ACT_DIST_ZTBINOM_CDF(k, 10, 0.3));
        }
        R(g, "ACT_DIST_ZTPOISSON_MEAN", new object?[] { 5.0 }, Distributions.ACT_DIST_ZTPOISSON_MEAN(5));
        R(g, "ACT_DIST_ZTNEGBIN_MEAN", new object?[] { 5.0, 0.3 }, Distributions.ACT_DIST_ZTNEGBIN_MEAN(5, 0.3));
        R(g, "ACT_DIST_ZTBINOM_MEAN", new object?[] { 10, 0.3 }, Distributions.ACT_DIST_ZTBINOM_MEAN(10, 0.3));
        R(g, "ACT_DIST_ZTGEOM_MEAN", new object?[] { 0.3 }, Distributions.ACT_DIST_ZTGEOM_MEAN(0.3));
    }

    private static void EmitZeroModified()
    {
        const string g = "zero_modified";
        foreach (int k in new[] { 0, 1, 2, 5, 10 })
        {
            R(g, "ACT_DIST_ZMPOISSON_PDF", new object?[] { k, 5.0, 0.2 }, Distributions.ACT_DIST_ZMPOISSON_PDF(k, 5, 0.2));
            R(g, "ACT_DIST_ZMPOISSON_CDF", new object?[] { k, 5.0, 0.2 }, Distributions.ACT_DIST_ZMPOISSON_CDF(k, 5, 0.2));
            R(g, "ACT_DIST_ZMNEGBIN_PDF", new object?[] { k, 5.0, 0.3, 0.2 }, Distributions.ACT_DIST_ZMNEGBIN_PDF(k, 5, 0.3, 0.2));
            R(g, "ACT_DIST_ZMNEGBIN_CDF", new object?[] { k, 5.0, 0.3, 0.2 }, Distributions.ACT_DIST_ZMNEGBIN_CDF(k, 5, 0.3, 0.2));
        }
        R(g, "ACT_DIST_ZMPOISSON_MEAN", new object?[] { 5.0, 0.2 }, Distributions.ACT_DIST_ZMPOISSON_MEAN(5, 0.2));
        R(g, "ACT_DIST_ZMPOISSON_VAR", new object?[] { 5.0, 0.2 }, Distributions.ACT_DIST_ZMPOISSON_VAR(5, 0.2));
        R(g, "ACT_DIST_ZMNEGBIN_MEAN", new object?[] { 5.0, 0.3, 0.2 }, Distributions.ACT_DIST_ZMNEGBIN_MEAN(5, 0.3, 0.2));
        R(g, "ACT_DIST_ZMNEGBIN_VAR", new object?[] { 5.0, 0.3, 0.2 }, Distributions.ACT_DIST_ZMNEGBIN_VAR(5, 0.3, 0.2));
    }

    private static void EmitParetoExtended()
    {
        const string g = "pareto_extended";
        // Pareto III (mu=1, sigma=2, gamma=3)
        foreach (double x in new[] { 1.0, 2.0, 3.0, 5.0 })
        {
            R(g, "ACT_DIST_PARETO3_PDF", new object?[] { x, 1.0, 2.0, 3.0 }, Distributions.ACT_DIST_PARETO3_PDF(x, 1, 2, 3));
            R(g, "ACT_DIST_PARETO3_CDF", new object?[] { x, 1.0, 2.0, 3.0 }, Distributions.ACT_DIST_PARETO3_CDF(x, 1, 2, 3));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_PARETO3_INV", new object?[] { p, 1.0, 2.0, 3.0 }, Distributions.ACT_DIST_PARETO3_INV(p, 1, 2, 3));
        R(g, "ACT_DIST_PARETO3_MEAN", new object?[] { 1.0, 2.0, 3.0 }, Distributions.ACT_DIST_PARETO3_MEAN(1, 2, 3));

        // Pareto IV (mu=0, sigma=2, gamma=0.5, alpha=3)
        foreach (double x in new[] { 0.5, 1.0, 2.0, 5.0 })
        {
            R(g, "ACT_DIST_PARETO4_PDF", new object?[] { x, 0.0, 2.0, 0.5, 3.0 }, Distributions.ACT_DIST_PARETO4_PDF(x, 0, 2, 0.5, 3));
            R(g, "ACT_DIST_PARETO4_CDF", new object?[] { x, 0.0, 2.0, 0.5, 3.0 }, Distributions.ACT_DIST_PARETO4_CDF(x, 0, 2, 0.5, 3));
        }
        foreach (double p in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_DIST_PARETO4_INV", new object?[] { p, 0.0, 2.0, 0.5, 3.0 }, Distributions.ACT_DIST_PARETO4_INV(p, 0, 2, 0.5, 3));
    }

    private static void EmitInverseGaussian()
    {
        const string g = "inverse_gaussian";
        double mu = 2.0, lam = 3.0;
        foreach (double x in new[] { 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            R(g, "ACT_DIST_INVGAUSS_PDF", new object?[] { x, mu, lam }, Distributions.ACT_DIST_INVGAUSS_PDF(x, mu, lam));
            R(g, "ACT_DIST_INVGAUSS_CDF", new object?[] { x, mu, lam }, Distributions.ACT_DIST_INVGAUSS_CDF(x, mu, lam));
            R(g, "ACT_DIST_INVGAUSS_LEV", new object?[] { x, mu, lam }, Distributions.ACT_DIST_INVGAUSS_LEV(x, mu, lam));
        }
        foreach (double p in new[] { 0.25, 0.5, 0.75, 0.95 })
            R(g, "ACT_DIST_INVGAUSS_INV", new object?[] { p, mu, lam }, Distributions.ACT_DIST_INVGAUSS_INV(p, mu, lam));
    }

    private static void EmitLoglogistic()
    {
        const string g = "loglogistic";
        double alpha = 2.0, beta = 3.0;
        foreach (double x in new[] { 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            R(g, "ACT_DIST_LOGLOGISTIC_PDF", new object?[] { x, alpha, beta }, Distributions.ACT_DIST_LOGLOGISTIC_PDF(x, alpha, beta));
            R(g, "ACT_DIST_LOGLOGISTIC_CDF", new object?[] { x, alpha, beta }, Distributions.ACT_DIST_LOGLOGISTIC_CDF(x, alpha, beta));
            R(g, "ACT_DIST_LOGLOGISTIC_LEV", new object?[] { x, alpha, beta }, Distributions.ACT_DIST_LOGLOGISTIC_LEV(x, alpha, beta));
        }
        foreach (double p in new[] { 0.25, 0.5, 0.75, 0.95 })
            R(g, "ACT_DIST_LOGLOGISTIC_INV", new object?[] { p, alpha, beta }, Distributions.ACT_DIST_LOGLOGISTIC_INV(p, alpha, beta));
    }

    private static void EmitCompositeDistributions()
    {
        const string g = "composite";
        // Lognormal-Pareto: mu=0, sigma=1, theta=2, alpha consistent
        foreach (double x in new[] { 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            R(g, "ACT_DIST_LNPARETO_PDF", new object?[] { x, 0.0, 1.0, 2.0 }, Distributions.ACT_DIST_LNPARETO_PDF(x, 0, 1, 2));
            R(g, "ACT_DIST_LNPARETO_CDF", new object?[] { x, 0.0, 1.0, 2.0 }, Distributions.ACT_DIST_LNPARETO_CDF(x, 0, 1, 2));
        }
        R(g, "ACT_DIST_LNPARETO_ALPHA", new object?[] { 0.0, 1.0, 2.0 }, Distributions.ACT_DIST_LNPARETO_ALPHA(0, 1, 2));

        // Exp-Pareto: lambda=1, theta=2
        foreach (double x in new[] { 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            R(g, "ACT_DIST_EXPPARETO_PDF", new object?[] { x, 1.0, 2.0 }, Distributions.ACT_DIST_EXPPARETO_PDF(x, 1, 2));
            R(g, "ACT_DIST_EXPPARETO_CDF", new object?[] { x, 1.0, 2.0 }, Distributions.ACT_DIST_EXPPARETO_CDF(x, 1, 2));
        }
        R(g, "ACT_DIST_EXPPARETO_ALPHA", new object?[] { 1.0, 2.0 }, Distributions.ACT_DIST_EXPPARETO_ALPHA(1, 2));

        // Power-Pareto: alpha=2, beta=1.5, theta=2
        foreach (double x in new[] { 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            R(g, "ACT_DIST_POWPARETO_PDF", new object?[] { x, 2.0, 1.5, 2.0 }, Distributions.ACT_DIST_POWPARETO_PDF(x, 2, 1.5, 2));
            R(g, "ACT_DIST_POWPARETO_CDF", new object?[] { x, 2.0, 1.5, 2.0 }, Distributions.ACT_DIST_POWPARETO_CDF(x, 2, 1.5, 2));
        }
    }

    private static void EmitLEV()
    {
        const string g = "lev";
        foreach (double d in new[] { 0.5, 1.0, 2.0, 5.0 })
        {
            R(g, "ACT_DIST_EXP_LEV", new object?[] { d, 2.0 }, Distributions.ACT_DIST_EXP_LEV(d, 2));
            R(g, "ACT_DIST_PARETO_LEV", new object?[] { d, 2.0, 1.0 }, Distributions.ACT_DIST_PARETO_LEV(d, 2, 1));
            R(g, "ACT_DIST_LOMAX_LEV", new object?[] { d, 2.0, 1.0 }, Distributions.ACT_DIST_LOMAX_LEV(d, 2, 1));
            R(g, "ACT_DIST_GPD_LEV", new object?[] { d, 0.5, 1.0 }, Distributions.ACT_DIST_GPD_LEV(d, 0.5, 1));
            R(g, "ACT_DIST_GAMMA_LEV", new object?[] { d, 2.0, 1.0 }, Distributions.ACT_DIST_GAMMA_LEV(d, 2, 1));
            R(g, "ACT_DIST_LOGNORM_LEV", new object?[] { d, 0.0, 1.0 }, Distributions.ACT_DIST_LOGNORM_LEV(d, 0, 1));
            R(g, "ACT_DIST_WEIBULL_LEV", new object?[] { d, 2.0, 1.0 }, Distributions.ACT_DIST_WEIBULL_LEV(d, 2, 1));
            R(g, "ACT_DIST_BURR_LEV", new object?[] { d, 2.0, 3.0, 1.0 }, Distributions.ACT_DIST_BURR_LEV(d, 2, 3, 1));
        }
        foreach (double d in new[] { 0.25, 0.5, 1.0 })
            R(g, "ACT_DIST_BETA_LEV", new object?[] { d, 2.0, 5.0 }, Distributions.ACT_DIST_BETA_LEV(d, 2, 5));

        foreach (double d in new[] { 3.0, 5.0, 10.0 })
        {
            R(g, "ACT_DIST_POISSON_LEV", new object?[] { d, 5.0 }, Distributions.ACT_DIST_POISSON_LEV(d, 5));
            R(g, "ACT_DIST_NEGBIN_LEV", new object?[] { d, 5.0, 0.3 }, Distributions.ACT_DIST_NEGBIN_LEV(d, 5, 0.3));
        }
    }

    private static void EmitFitting()
    {
        const string g = "fitting";

        // Deterministic datasets (drawn once, hardcoded here).
        var expSample = new double[,] { { 0.32, 1.20, 0.85, 2.10, 0.45, 1.80, 3.00, 0.55, 1.10, 2.40 } };
        R(g, "ACT_DIST_EXP_FIT", new object?[] { expSample }, Fitting.ACT_DIST_EXP_FIT(expSample));

        var poiSample = new double[,] { { 3, 5, 4, 6, 5, 3, 7, 4, 5, 2 } };
        R(g, "ACT_DIST_POISSON_FIT", new object?[] { poiSample }, Fitting.ACT_DIST_POISSON_FIT(poiSample));

        var lnSample = new double[,] { { 1.2, 2.3, 4.1, 3.0, 1.7, 5.2, 2.8, 3.6, 4.5, 2.0 } };
        R(g, "ACT_DIST_LOGNORM_FIT", new object?[] { lnSample }, Fitting.ACT_DIST_LOGNORM_FIT(lnSample));

        var gammaSample = new double[,] { { 2.1, 3.5, 1.8, 4.2, 2.9, 3.1, 2.5, 4.7, 3.8, 2.3 } };
        R(g, "ACT_DIST_GAMMA_FIT", new object?[] { gammaSample }, Fitting.ACT_DIST_GAMMA_FIT(gammaSample));

        var paretoSample = new double[,] { { 1.5, 2.1, 3.0, 1.8, 5.0, 2.3, 1.2, 4.5, 6.0, 1.7 } };
        R(g, "ACT_DIST_PARETO_FIT", new object?[] { paretoSample }, Fitting.ACT_DIST_PARETO_FIT(paretoSample));

        var weibullSample = new double[,] { { 0.8, 1.2, 1.5, 0.9, 1.8, 2.1, 1.3, 0.6, 1.7, 1.1 } };
        R(g, "ACT_DIST_WEIBULL_FIT", new object?[] { weibullSample }, Fitting.ACT_DIST_WEIBULL_FIT(weibullSample));

        var gpdSample = new double[,] { { 0.5, 1.0, 2.0, 3.5, 0.8, 1.5, 2.8, 0.3, 4.0, 1.2 } };
        R(g, "ACT_DIST_GPD_FIT", new object?[] { gpdSample }, Fitting.ACT_DIST_GPD_FIT(gpdSample));

        var betaSample = new double[,] { { 0.2, 0.3, 0.1, 0.4, 0.25, 0.15, 0.35, 0.45, 0.2, 0.3 } };
        R(g, "ACT_DIST_BETA_FIT", new object?[] { betaSample }, Fitting.ACT_DIST_BETA_FIT(betaSample));

        var negbinSample = new double[,] { { 2, 4, 5, 3, 6, 7, 4, 3, 8, 5, 4, 6, 7, 3, 5 } };
        R(g, "ACT_DIST_NEGBIN_FIT", new object?[] { negbinSample }, Fitting.ACT_DIST_NEGBIN_FIT(negbinSample));

        var burrSample = new double[,] { { 0.5, 1.0, 1.5, 2.0, 0.8, 1.2, 1.8, 0.6, 1.3, 1.7 } };
        R(g, "ACT_DIST_BURR_FIT", new object?[] { burrSample }, Fitting.ACT_DIST_BURR_FIT(burrSample));
    }

    private static void EmitAggregate()
    {
        const string g = "aggregate";
        var fDeg = new double[] { 0.0, 1.0 };

        var gPoi = (double[])Aggregate.ACT_PANJER_POISSON(2.0, fDeg, 10);
        R(g, "ACT_PANJER_POISSON", new object?[] { 2.0, fDeg, 10 }, gPoi);

        var gNB = (double[])Aggregate.ACT_PANJER_NEGBIN(2.0, 0.5, fDeg, 10);
        R(g, "ACT_PANJER_NEGBIN", new object?[] { 2.0, 0.5, fDeg, 10 }, gNB);

        var gBin = (double[])Aggregate.ACT_PANJER_BINOMIAL(10, 0.3, fDeg, 10);
        R(g, "ACT_PANJER_BINOMIAL", new object?[] { 10, 0.3, fDeg, 10 }, gBin);

        var fExp = (double[])Aggregate.ACT_DISCRETIZE_EXPONENTIAL(1.0, 0.5, 40);
        R(g, "ACT_DISCRETIZE_EXPONENTIAL", new object?[] { 1.0, 0.5, 40 }, fExp);

        var fGamma = (double[])Aggregate.ACT_DISCRETIZE_GAMMA(2.0, 1.0, 0.5, 40);
        R(g, "ACT_DISCRETIZE_GAMMA", new object?[] { 2.0, 1.0, 0.5, 40 }, fGamma);

        var fLN = (double[])Aggregate.ACT_DISCRETIZE_LOGNORMAL(0.0, 1.0, 0.5, 40);
        R(g, "ACT_DISCRETIZE_LOGNORMAL", new object?[] { 0.0, 1.0, 0.5, 40 }, fLN);

        var gExp = (double[])Aggregate.ACT_PANJER_POISSON(2.0, fExp, 100);
        R(g, "ACT_PANJER_POISSON_EXP", new object?[] { 2.0, "DISCRETIZE_EXPONENTIAL(1,0.5,40)", 100 }, gExp);
        R(g, "ACT_AGGREGATE_MEAN", new object?[] { "gExp", 0.5 }, Aggregate.ACT_AGGREGATE_MEAN(gExp, 0.5));
        R(g, "ACT_AGGREGATE_STDEV", new object?[] { "gExp", 0.5 }, Aggregate.ACT_AGGREGATE_STDEV(gExp, 0.5));
        R(g, "ACT_AGGREGATE_VAR_STAT", new object?[] { "gExp", 0.5 }, Aggregate.ACT_AGGREGATE_VAR_STAT(gExp, 0.5));
        foreach (double a in new[] { 0.5, 0.75, 0.95, 0.99 })
        {
            R(g, "ACT_AGGREGATE_VAR", new object?[] { a, "gExp", 0.5 }, Aggregate.ACT_AGGREGATE_VAR(a, gExp, 0.5));
            R(g, "ACT_AGGREGATE_TVAR", new object?[] { a, "gExp", 0.5 }, Aggregate.ACT_AGGREGATE_TVAR(a, gExp, 0.5));
            R(g, "ACT_AGGREGATE_CDF", new object?[] { a * 5, "gExp", 0.5 }, Aggregate.ACT_AGGREGATE_CDF(a * 5, gExp, 0.5));
        }
    }

    private static void EmitExposureCurves()
    {
        const string g = "exposure_curves";
        foreach (double d in new[] { 0.1, 0.25, 0.5, 0.75, 0.9 })
        {
            R(g, "ACT_EXPOSURE_MBBEFD", new object?[] { d, 2.0, 3.0 }, ExposureCurves.ACT_EXPOSURE_MBBEFD(d, 2, 3));
            R(g, "ACT_EXPOSURE_POWER", new object?[] { d, 2.0 }, ExposureCurves.ACT_EXPOSURE_POWER(d, 2));
            R(g, "ACT_EXPOSURE_INVERSE_POWER", new object?[] { d, 2.0 }, ExposureCurves.ACT_EXPOSURE_INVERSE_POWER(d, 2));
            R(g, "ACT_EXPOSURE_PARETO", new object?[] { d, 2.0 }, ExposureCurves.ACT_EXPOSURE_PARETO(d, 2));
            R(g, "ACT_EXPOSURE_RIEBESELL", new object?[] { d, 0.5 }, ExposureCurves.ACT_EXPOSURE_RIEBESELL(d, 0.5));
        }
        for (int c = 1; c <= 5; c++)
            R(g, "ACT_EXPOSURE_SWISSRE", new object?[] { 0.5, c }, ExposureCurves.ACT_EXPOSURE_SWISSRE(0.5, c));
        foreach (var code in new[] { "Y1", "Y2", "Y3", "Y4" })
            R(g, "ACT_EXPOSURE_LLOYDS", new object?[] { 0.5, code }, ExposureCurves.ACT_EXPOSURE_LLOYDS(0.5, code));

        foreach (double g0 in new[] { 0.1, 0.5, 0.9 })
            R(g, "ACT_EXPOSURE_RIEBESELL_INV", new object?[] { g0, 0.5, 1e-8, 100 }, ExposureCurves.ACT_EXPOSURE_RIEBESELL_INV(g0, 0.5));

        R(g, "ACT_EXPOSURE_LAYER_RATE", new object?[] { 0.2, 0.8, 1.0, 2.0, 3.0 },
            ExposureCurves.ACT_EXPOSURE_LAYER_RATE(0.2, 0.8, 1.0, 2, 3));
    }

    private static void EmitReinsurance()
    {
        const string g = "reinsurance";
        foreach (var (loss, att, lim) in new[] { (500_000.0, 1_000_000.0, 5_000_000.0), (2_000_000, 1_000_000, 5_000_000), (10_000_000, 1_000_000, 5_000_000) })
            R(g, "ACT_XOL_LAYER_LOSS", new object?[] { loss, att, lim }, Reinsurance.ACT_XOL_LAYER_LOSS(loss, att, lim));

        R(g, "ACT_XOL_EXPECTED_LOSS", new object?[] { 10.0, 1_000_000, 5_000_000, 2.0, 100_000 },
            Reinsurance.ACT_XOL_EXPECTED_LOSS(10, 1_000_000, 5_000_000, 2, 100_000));

        foreach (double tgt in new[] { 100_000.0, 500_000, 1_000_000 })
            R(g, "ACT_ILF_PARETO", new object?[] { tgt, 100_000.0, 2.0 }, Reinsurance.ACT_ILF_PARETO(tgt, 100_000, 2));

        var rps = new double[] { 10, 25, 50, 100, 250 };
        var losses = new double[] { 100_000, 250_000, 500_000, 1_000_000, 2_500_000 };
        foreach (double target in new double[] { 20, 75, 150, 200 })
            R(g, "ACT_RETURN_PERIOD_LOSS", new object?[] { rps, losses, target, "LOG" }, Reinsurance.ACT_RETURN_PERIOD_LOSS(rps, losses, target, "LOG"));

        var table = Reinsurance.ACT_RETURN_PERIOD_TABLE(rps, losses, new double[] { 20, 75, 150, 200 }, "LOG");
        R(g, "ACT_RETURN_PERIOD_TABLE", new object?[] { rps, losses, new double[] { 20, 75, 150, 200 }, "LOG" }, table);

        R(g, "ACT_AAL_FROM_OEP", new object?[] { rps, losses }, Reinsurance.ACT_AAL_FROM_OEP(rps, losses));
    }

    private static void EmitCatModeling()
    {
        const string g = "cat_modeling";
        var rates = new double[] { 0.2, 0.1, 0.05 };
        var lossEvents = new double[] { 1_000_000, 2_500_000, 10_000_000 };

        var ylt = CatModeling.ACT_CAT_ELT_TO_YLT(rates, lossEvents, 1000, 42, true);
        R(g, "ACT_CAT_ELT_TO_YLT", new object?[] { rates, lossEvents, 1000, 42, true }, ylt);

        // Derived columns
        int n = ylt.GetLength(0) - 1;
        var aggLosses = new double[n];
        var maxLosses = new double[n];
        for (int i = 0; i < n; i++)
        {
            aggLosses[i] = Convert.ToDouble(ylt[i + 1, 1]);
            maxLosses[i] = Convert.ToDouble(ylt[i + 1, 2]);
        }

        var oep = CatModeling.ACT_CAT_YLT_OEP_CURVE(maxLosses, "WEIBULL", true);
        R(g, "ACT_CAT_YLT_OEP_CURVE", new object?[] { "maxLosses_n1000_seed42", "WEIBULL", true }, oep);

        var aep = CatModeling.ACT_CAT_YLT_AEP_CURVE(aggLosses, "WEIBULL", true);
        R(g, "ACT_CAT_YLT_AEP_CURVE", new object?[] { "aggLosses_n1000_seed42", "WEIBULL", true }, aep);

        var targetRps = new double[] { 10, 50, 100, 250 };
        var oepRp = CatModeling.ACT_CAT_OEP_CURVE_RP(maxLosses, targetRps, true);
        R(g, "ACT_CAT_OEP_CURVE_RP", new object?[] { "maxLosses", targetRps, true }, oepRp);

        var aepRp = CatModeling.ACT_CAT_AEP_CURVE_RP(aggLosses, targetRps, true);
        R(g, "ACT_CAT_AEP_CURVE_RP", new object?[] { "aggLosses", targetRps, true }, aepRp);

        foreach (double a in new[] { 0.95, 0.99, 0.995 })
        {
            R(g, "ACT_VAR_FROM_SAMPLES", new object?[] { "aggLosses", a }, CatModeling.ACT_VAR_FROM_SAMPLES(aggLosses, a));
            R(g, "ACT_TVAR_FROM_SAMPLES", new object?[] { "aggLosses", a }, CatModeling.ACT_TVAR_FROM_SAMPLES(aggLosses, a));
        }
    }

    private static void EmitInterpolation()
    {
        const string g = "interpolation";
        var x = new double[] { 1, 2, 3, 4, 5 };
        var y = new double[] { 10, 20, 35, 55, 80 };
        foreach (double xq in new[] { 0.5, 1.5, 2.5, 3.5, 4.5, 5.5, 6.0 })
        {
            R(g, "ACT_INTERP_FLAT", new object?[] { x, y, xq, "FLAT" }, Interpolation.ACT_INTERP(x, y, xq, "FLAT"));
            R(g, "ACT_INTERP_GRADIENT", new object?[] { x, y, xq, "GRADIENT" }, Interpolation.ACT_INTERP(x, y, xq, "GRADIENT"));
            R(g, "ACT_INTERP_LOG", new object?[] { x, y, xq, "FLAT" }, Interpolation.ACT_INTERP_LOG(x, y, xq, "FLAT"));
        }

        var xv = new double[] { 1, 2, 3 };
        var yv = new double[] { 10, 20, 30 };
        var zv = new double[,] { { 1, 2, 3 }, { 4, 5, 6 }, { 7, 8, 9 } };
        foreach (var (xq, yq) in new[] { (1.5, 15.0), (2.0, 20.0), (2.5, 25.0) })
            R(g, "ACT_INTERP2D", new object?[] { xv, yv, zv, xq, yq }, Interpolation.ACT_INTERP2D(xv, yv, zv, xq, yq));
    }

    private static double[,] TaylorAshe()
    {
        return new double[,]
        {
            { 357848, 1124788, 1735330, 2218270, 2745596, 3319994, 3466336, 3606286, 3833515, 3901463 },
            { 352118, 1236139, 2170033, 3353322, 3799067, 4120063, 4647867, 4914039, 5339085, 0 },
            { 290507, 1292306, 2218525, 3235179, 3985995, 4132918, 4628910, 4909315, 0, 0 },
            { 310608, 1418858, 2195047, 3757447, 4029929, 4381982, 4588268, 0, 0, 0 },
            { 443160, 1136350, 2128333, 2897821, 3402672, 3873311, 0, 0, 0, 0 },
            { 396132, 1333217, 2180715, 2985752, 3691712, 0, 0, 0, 0, 0 },
            { 440832, 1288463, 2419861, 3483130, 0, 0, 0, 0, 0, 0 },
            { 359480, 1421128, 2864498, 0, 0, 0, 0, 0, 0, 0 },
            { 376686, 1363294, 0, 0, 0, 0, 0, 0, 0, 0 },
            { 344014, 0, 0, 0, 0, 0, 0, 0, 0, 0 }
        };
    }

    // Unbox a column-vector object[,] (or object[]) into a double[]
    private static double[] UnboxDoubleColumn(object excelResult)
    {
        if (excelResult is object[,] arr2d)
        {
            int len = arr2d.GetLength(0);
            var result = new double[len];
            for (int i = 0; i < len; i++) result[i] = Convert.ToDouble(arr2d[i, 0]);
            return result;
        }
        if (excelResult is object[] arr1d)
        {
            var result = new double[arr1d.Length];
            for (int i = 0; i < arr1d.Length; i++) result[i] = Convert.ToDouble(arr1d[i]);
            return result;
        }
        throw new InvalidOperationException($"Expected column-shaped output, got {excelResult?.GetType()}");
    }

    private static double[,] UnboxDoubleMatrix(object[,] src)
    {
        int rows = src.GetLength(0), cols = src.GetLength(1);
        var dst = new double[rows, cols];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
            {
                var v = src[i, j];
                // ACT_TRIANGLE_TO_INCREMENTAL fills future cells with empty strings.
                // Treat those as zeros for the downstream cum-incremental round-trip.
                if (v is string s && string.IsNullOrEmpty(s))
                    dst[i, j] = 0.0;
                else
                    dst[i, j] = Convert.ToDouble(v);
            }
        return dst;
    }

    private static void EmitChainLadder()
    {
        const string g = "chain_ladder";
        var tri = TaylorAshe();

        R(g, "ACT_CL_FACTORS", new object?[] { tri }, ChainLadder.ACT_CL_FACTORS(tri));
        R(g, "ACT_CL_LATEST", new object?[] { tri }, ChainLadder.ACT_CL_LATEST(tri));
        R(g, "ACT_CL_ULTIMATE", new object?[] { tri }, ChainLadder.ACT_CL_ULTIMATE(tri));
        R(g, "ACT_CL_IBNR", new object?[] { tri }, ChainLadder.ACT_CL_IBNR(tri));

        var factorsD = UnboxDoubleColumn(ChainLadder.ACT_CL_FACTORS(tri));
        var latestD = UnboxDoubleColumn(ChainLadder.ACT_CL_LATEST(tri));
        var ultD = UnboxDoubleColumn(ChainLadder.ACT_CL_ULTIMATE(tri));
        var apriori = latestD.Select(v => v * 1.10).ToArray();
        R(g, "ACT_BF_ULTIMATE", new object?[] { tri, factorsD, apriori }, ChainLadder.ACT_BF_ULTIMATE(tri, factorsD, apriori));

        R(g, "ACT_MACK_FACTOR_SE", new object?[] { tri }, ChainLadder.ACT_MACK_FACTOR_SE(tri));
        R(g, "ACT_MACK_RESERVE_SE", new object?[] { tri }, ChainLadder.ACT_MACK_RESERVE_SE(tri));

        var premium = latestD.Select(v => v / 0.7).ToArray();
        R(g, "ACT_CAPECOD_ULTIMATE", new object?[] { tri, factorsD, premium }, ChainLadder.ACT_CAPECOD_ULTIMATE(tri, factorsD, premium));
        R(g, "ACT_CAPECOD_ELR", new object?[] { tri, factorsD, premium }, ChainLadder.ACT_CAPECOD_ELR(tri, factorsD, premium));

        var incr = ChainLadder.ACT_TRIANGLE_TO_INCREMENTAL(tri);
        R(g, "ACT_TRIANGLE_TO_INCREMENTAL", new object?[] { tri }, incr);

        var incrD = UnboxDoubleMatrix(incr);
        R(g, "ACT_INCREMENTAL_TO_CUMULATIVE", new object?[] { incrD }, ChainLadder.ACT_INCREMENTAL_TO_CUMULATIVE(incrD));

        R(g, "ACT_TRIANGLE_DIAGONAL", new object?[] { tri }, ChainLadder.ACT_TRIANGLE_DIAGONAL(tri));
        R(g, "ACT_TRIANGLE_LINK_RATIOS", new object?[] { tri }, ChainLadder.ACT_TRIANGLE_LINK_RATIOS(tri));

        R(g, "ACT_CL_CALENDAR_ADJUST", new object?[] { tri, 0.03 }, ChainLadder.ACT_CL_CALENDAR_ADJUST(tri, 0.03));
        R(g, "ACT_CL_CALENDAR_TOTALS", new object?[] { tri }, ChainLadder.ACT_CL_CALENDAR_TOTALS(tri));

        // 50/50 weighted average of CL ultimate and BF apriori
        R(g, "ACT_CL_WEIGHTED_AVERAGE", new object?[] { ultD, 0.5, apriori, 0.5 },
            ChainLadder.ACT_CL_WEIGHTED_AVERAGE(ultD, 0.5, apriori, 0.5));
    }

    private static void EmitBootstrap()
    {
        const string g = "bootstrap";
        var tri = TaylorAshe();

        // Seed 42 matches the bootstrapping_exposition repo so per-origin PE
        // values can be cross-read against England (2010) slide 35 directly
        // (modulo numpy vs System.Random sequence differences — comparison is
        // statistical, not bit-exact).
        var evTotal = ChainLadder.ACT_CL_BOOTSTRAP(tri, 10000, 42, "EV");
        R(g, "ACT_CL_BOOTSTRAP_EV", new object?[] { "TaylorAshe", 10000, 42, "EV" }, evTotal);

        var evOrigin = ChainLadder.ACT_CL_BOOTSTRAP_ORIGIN(tri, 10000, 42, "EV");
        R(g, "ACT_CL_BOOTSTRAP_ORIGIN_EV", new object?[] { "TaylorAshe", 10000, 42, "EV" }, evOrigin);

        var clpTotal = ChainLadder.ACT_CL_BOOTSTRAP(tri, 10000, 42, "CHAINLADDER-PYTHON");
        R(g, "ACT_CL_BOOTSTRAP_CLP", new object?[] { "TaylorAshe", 10000, 42, "CHAINLADDER-PYTHON" }, clpTotal);

        var clpOrigin = ChainLadder.ACT_CL_BOOTSTRAP_ORIGIN(tri, 10000, 42, "CHAINLADDER-PYTHON");
        R(g, "ACT_CL_BOOTSTRAP_ORIGIN_CLP", new object?[] { "TaylorAshe", 10000, 42, "CHAINLADDER-PYTHON" }, clpOrigin);
    }

    private static void EmitCopulas()
    {
        const string g = "copulas";
        // Small sample draws with fixed seed, plus analytical quantities.
        int dim = 3;
        var corr = new double[dim, dim];
        for (int i = 0; i < dim; i++)
            for (int j = 0; j < dim; j++)
                corr[i, j] = Math.Pow(0.6, Math.Abs(i - j));

        var gauss = Copulas.ACT_COPULA_GAUSSIAN(corr, 500, 42);
        R(g, "ACT_COPULA_GAUSSIAN", new object?[] { corr, 500, 42 }, gauss);

        var studt = Copulas.ACT_COPULA_STUDENT_T(corr, 5.0, 500, 42);
        R(g, "ACT_COPULA_STUDENT_T", new object?[] { corr, 5.0, 500, 42 }, studt);

        var clayton = Copulas.ACT_COPULA_CLAYTON(2.0, 500, 42);
        R(g, "ACT_COPULA_CLAYTON", new object?[] { 2.0, 500, 42 }, clayton);

        var frank = Copulas.ACT_COPULA_FRANK(5.0, 500, 42);
        R(g, "ACT_COPULA_FRANK", new object?[] { 5.0, 500, 42 }, frank);

        var gumbel = Copulas.ACT_COPULA_GUMBEL(2.0, 500, 42);
        R(g, "ACT_COPULA_GUMBEL", new object?[] { 2.0, 500, 42 }, gumbel);

        // Analytical CDFs
        foreach (var (u, v) in new[] { (0.3, 0.7), (0.5, 0.5), (0.1, 0.9) })
        {
            R(g, "ACT_COPULA_CLAYTON_CDF", new object?[] { u, v, 2.0 }, Copulas.ACT_COPULA_CLAYTON_CDF(u, v, 2));
            R(g, "ACT_COPULA_FRANK_CDF", new object?[] { u, v, 5.0 }, Copulas.ACT_COPULA_FRANK_CDF(u, v, 5));
            R(g, "ACT_COPULA_GUMBEL_CDF", new object?[] { u, v, 2.0 }, Copulas.ACT_COPULA_GUMBEL_CDF(u, v, 2));
        }

        // Tau → theta conversions
        foreach (double tau in new[] { 0.2, 0.5, 0.8 })
        {
            R(g, "ACT_COPULA_TAU_TO_THETA_CLAYTON", new object?[] { tau, "CLAYTON" }, Copulas.ACT_COPULA_TAU_TO_THETA(tau, "CLAYTON"));
            R(g, "ACT_COPULA_TAU_TO_THETA_FRANK", new object?[] { tau, "FRANK" }, Copulas.ACT_COPULA_TAU_TO_THETA(tau, "FRANK"));
            R(g, "ACT_COPULA_TAU_TO_THETA_GUMBEL", new object?[] { tau, "GUMBEL" }, Copulas.ACT_COPULA_TAU_TO_THETA(tau, "GUMBEL"));
        }

        // Tail dependence
        R(g, "ACT_COPULA_TAIL_LOWER_GAUSSIAN", new object?[] { "GAUSSIAN", 0.5, 0.0 }, Copulas.ACT_COPULA_TAIL_LOWER("GAUSSIAN", 0.5));
        R(g, "ACT_COPULA_TAIL_LOWER_T", new object?[] { "STUDENT_T", 0.5, 5.0 }, Copulas.ACT_COPULA_TAIL_LOWER("STUDENT_T", 0.5, 5));
        R(g, "ACT_COPULA_TAIL_LOWER_CLAYTON", new object?[] { "CLAYTON", 2.0, 0.0 }, Copulas.ACT_COPULA_TAIL_LOWER("CLAYTON", 2));
        R(g, "ACT_COPULA_TAIL_UPPER_GUMBEL", new object?[] { "GUMBEL", 2.0, 0.0 }, Copulas.ACT_COPULA_TAIL_UPPER("GUMBEL", 2));
        R(g, "ACT_COPULA_TAIL_UPPER_T", new object?[] { "STUDENT_T", 0.5, 5.0 }, Copulas.ACT_COPULA_TAIL_UPPER("STUDENT_T", 0.5, 5));
    }
}
