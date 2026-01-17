using System.Text;
using ActuarialAddIn.Functions;

namespace ActuarialAddIn.Tests;

class Program
{
    static StringBuilder _output = new();

    static void Main(string[] args)
    {
        _output.AppendLine("# Actuarial Add-In Test Results");
        _output.AppendLine($"\nTest run: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");

        TestDistributions();
        TestExposureCurves();
        TestReinsurance();
        TestInterpolation();
        TestChainLadder();
        TestCopulas();

        var outputPath = args.Length > 0 ? args[0] : "test_results.md";
        File.WriteAllText(outputPath, _output.ToString());
        Console.WriteLine(_output.ToString());
        Console.WriteLine($"\nResults written to: {outputPath}");
    }

    static void Log(string message) => _output.AppendLine(message);

    static void TestDistributions()
    {
        Log("## 1. Statistical Distributions\n");

        Log("### Poisson Distribution (λ=5)");
        Log("| k | PDF | CDF |");
        Log("|---|-----|-----|");
        for (int k = 0; k <= 10; k++)
        {
            var pdf = Distributions.ACT_POISSON_PDF(k, 5);
            var cdf = Distributions.ACT_POISSON_CDF(k, 5);
            Log($"| {k} | {pdf:F6} | {cdf:F6} |");
        }
        Log($"\nInverse CDF: P(0.5) = {Distributions.ACT_POISSON_INV(0.5, 5)}\n");

        Log("### Negative Binomial Distribution (r=5, p=0.3)");
        Log("| k | PDF | CDF |");
        Log("|---|-----|-----|");
        for (int k = 0; k <= 10; k++)
        {
            var pdf = Distributions.ACT_NEGBIN_PDF(k, 5, 0.3);
            var cdf = Distributions.ACT_NEGBIN_CDF(k, 5, 0.3);
            Log($"| {k} | {pdf:F6} | {cdf:F6} |");
        }
        Log($"\nInverse CDF: P(0.5) = {Distributions.ACT_NEGBIN_INV(0.5, 5, 0.3)}\n");

        Log("### Lognormal Distribution (μ=0, σ=1)");
        Log("| x | PDF | CDF |");
        Log("|---|-----|-----|");
        foreach (var x in new[] { 0.5, 1.0, 1.5, 2.0, 3.0, 5.0 })
        {
            var pdf = Distributions.ACT_LOGNORM_PDF(x, 0, 1);
            var cdf = Distributions.ACT_LOGNORM_CDF(x, 0, 1);
            Log($"| {x} | {pdf:F6} | {cdf:F6} |");
        }
        Log($"\nInverse CDF: P(0.5) = {Distributions.ACT_LOGNORM_INV(0.5, 0, 1):F6}\n");

        Log("### Gamma Distribution (α=2, β=1)");
        Log("| x | PDF | CDF |");
        Log("|---|-----|-----|");
        foreach (var x in new[] { 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            var pdf = Distributions.ACT_GAMMA_PDF(x, 2, 1);
            var cdf = Distributions.ACT_GAMMA_CDF(x, 2, 1);
            Log($"| {x} | {pdf:F6} | {cdf:F6} |");
        }
        Log($"\nInverse CDF: P(0.5) = {Distributions.ACT_GAMMA_INV(0.5, 2, 1):F6}\n");

        Log("### Pareto Distribution (α=2, xm=1)");
        Log("| x | PDF | CDF |");
        Log("|---|-----|-----|");
        foreach (var x in new[] { 1.0, 1.5, 2.0, 3.0, 5.0, 10.0 })
        {
            var pdf = Distributions.ACT_PARETO_PDF(x, 2, 1);
            var cdf = Distributions.ACT_PARETO_CDF(x, 2, 1);
            Log($"| {x} | {pdf:F6} | {cdf:F6} |");
        }
        Log($"\nInverse CDF: P(0.5) = {Distributions.ACT_PARETO_INV(0.5, 2, 1):F6}\n");
    }

    static void TestExposureCurves()
    {
        Log("## 2. Exposure Curves\n");

        Log("### MBBEFD Curves (b=2, g=3)");
        Log("| d | G(d) |");
        Log("|---|------|");
        foreach (var d in new[] { 0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0 })
        {
            var g = ExposureCurves.ACT_MBBEFD(d, 2, 3);
            Log($"| {d:F1} | {g:F6} |");
        }

        Log("\n### Swiss Re Curves Comparison at d=0.5");
        Log("| Curve | G(0.5) |");
        Log("|-------|--------|");
        for (int c = 1; c <= 5; c++)
        {
            var g = ExposureCurves.ACT_SWISSRE_CURVE(0.5, c);
            Log($"| {c} | {g:F6} |");
        }

        Log("\n### Lloyd's Curves Comparison at d=0.5");
        Log("| Curve | G(0.5) |");
        Log("|-------|--------|");
        foreach (var curve in new[] { "Y1", "Y2", "Y3", "Y4" })
        {
            var g = ExposureCurves.ACT_LLOYDS_CURVE(0.5, curve);
            Log($"| {curve} | {g:F6} |");
        }

        Log("\n### Power and Pareto Curves at d=0.5");
        Log($"- Power curve (n=2): {ExposureCurves.ACT_POWER_CURVE(0.5, 2):F6}");
        Log($"- Inverse power (n=2): {ExposureCurves.ACT_INVERSE_POWER_CURVE(0.5, 2):F6}");
        Log($"- Pareto exposure (α=2): {ExposureCurves.ACT_PARETO_EXPOSURE(0.5, 2):F6}\n");
    }

    static void TestReinsurance()
    {
        Log("## 3. Reinsurance Functions\n");

        Log("### XOL Layer Loss (Attachment=1M, Limit=5M)");
        Log("| Ground-Up Loss | Layer Loss |");
        Log("|----------------|------------|");
        foreach (var loss in new[] { 500_000.0, 1_000_000, 2_000_000, 4_000_000, 6_000_000, 10_000_000 })
        {
            var layerLoss = Reinsurance.ACT_XOL_LAYER_LOSS(loss, 1_000_000, 5_000_000);
            Log($"| {loss:N0} | {layerLoss:N0} |");
        }

        Log("\n### Quota Share (50% cession)");
        Log($"- Ground-up: 1,000,000 → Ceded: {Reinsurance.ACT_QS_CEDED(1_000_000, 0.5):N0}");

        Log("\n### Aggregate Layer (Deductible=2M, Limit=10M)");
        Log($"- Aggregate loss 15M → Layer: {Reinsurance.ACT_AGGREGATE_LAYER(15_000_000, 2_000_000, 10_000_000):N0}\n");
    }

    static void TestInterpolation()
    {
        Log("## 4. Interpolation\n");

        var xVals = new double[] { 1, 2, 3, 4, 5 };
        var yVals = new double[] { 10, 20, 35, 55, 80 };

        Log("### Linear Interpolation");
        Log("Known points: (1,10), (2,20), (3,35), (4,55), (5,80)");
        Log("| x | y (FLAT) | y (GRADIENT) |");
        Log("|---|----------|--------------|");
        foreach (var x in new[] { 0.5, 1.5, 2.5, 3.5, 4.5, 5.5, 6.0 })
        {
            var yFlat = Interpolation.ACT_INTERP(xVals, yVals, x, "FLAT");
            var yGradient = Interpolation.ACT_INTERP(xVals, yVals, x, "GRADIENT");
            Log($"| {x} | {yFlat:F2} | {yGradient:F2} |");
        }
        Log("");
    }

    static void TestChainLadder()
    {
        Log("## 5. Chain Ladder\n");

        // Sample triangle (cumulative paid losses)
        var triangle = new double[,]
        {
            { 100, 150, 170, 180, 185 },
            { 110, 165, 190, 200, 0 },
            { 120, 180, 210, 0, 0 },
            { 130, 195, 0, 0, 0 },
            { 140, 0, 0, 0, 0 }
        };

        Log("### Input Triangle (Cumulative Paid Losses)");
        Log("| AY\\Dev | 1 | 2 | 3 | 4 | 5 |");
        Log("|--------|---|---|---|---|---|");
        for (int i = 0; i < 5; i++)
        {
            var row = $"| {i + 1} |";
            for (int j = 0; j < 5; j++)
            {
                row += triangle[i, j] > 0 ? $" {triangle[i, j]} |" : " - |";
            }
            Log(row);
        }

        Log("\n### Development Factors");
        var factors = ChainLadder.ACT_CL_FACTORS(triangle);
        Log("| Period | Factor |");
        Log("|--------|--------|");
        for (int i = 0; i < factors.Length; i++)
        {
            Log($"| {i + 1}-{i + 2} | {factors[i]:F4} |");
        }

        Log("\n### Projected Ultimates");
        var ultimates = ChainLadder.ACT_CL_ULTIMATE(triangle);
        Log("| AY | Ultimate |");
        Log("|----|----------|");
        for (int i = 0; i < ultimates.Length; i++)
        {
            Log($"| {i + 1} | {ultimates[i]:F2} |");
        }

        Log("\n### IBNR Reserves");
        var ibnr = ChainLadder.ACT_CL_IBNR(triangle);
        Log("| AY | IBNR |");
        Log("|----|------|");
        double totalIBNR = 0;
        for (int i = 0; i < ibnr.Length; i++)
        {
            Log($"| {i + 1} | {ibnr[i]:F2} |");
            totalIBNR += (double)ibnr[i];
        }
        Log($"| **Total** | **{totalIBNR:F2}** |");

        Log("\n### Mack Standard Errors");
        var reserveSE = ChainLadder.ACT_MACK_RESERVE_SE(triangle);
        Log("| AY | Reserve SE |");
        Log("|----|------------|");
        for (int i = 0; i < reserveSE.Length; i++)
        {
            Log($"| {i + 1} | {reserveSE[i]:F2} |");
        }
        Log("");
    }

    static void TestCopulas()
    {
        Log("## 6. Student-t Copula\n");

        var corrMatrix = new double[,]
        {
            { 1.0, 0.5, 0.3 },
            { 0.5, 1.0, 0.4 },
            { 0.3, 0.4, 1.0 }
        };

        Log("### Correlation Matrix");
        Log("| | X1 | X2 | X3 |");
        Log("|--|----|----|---|");
        Log($"| X1 | 1.00 | 0.50 | 0.30 |");
        Log($"| X2 | 0.50 | 1.00 | 0.40 |");
        Log($"| X3 | 0.30 | 0.40 | 1.00 |");

        Log("\n### Generated Samples (df=5, 5 samples, seed=42)");
        var samples = Copulas.ACT_STUDENT_T_COPULA(corrMatrix, 5, 5, 42);

        if (samples[0, 0] is string errorMsg)
        {
            Log($"Error: {errorMsg}");
        }
        else
        {
            Log("| Sample | U1 | U2 | U3 |");
            Log("|--------|----|----|---|");
            for (int i = 0; i < 5; i++)
            {
                Log($"| {i + 1} | {samples[i, 0]:F4} | {samples[i, 1]:F4} | {samples[i, 2]:F4} |");
            }
        }
        Log("");
    }
}
