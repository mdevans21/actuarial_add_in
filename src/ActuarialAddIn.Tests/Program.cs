using System.Text;
using System.Linq;
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
        TestReturnPeriods();
        TestInterpolation();
        TestChainLadder();
        TestBootstrapChainLadder();
        TestBerquistSherman();
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
        foreach (var x in new[] { 0.0, 0.5, 1.0, 1.5, 2.0, 3.0, 5.0 })
        {
            var pdf = Distributions.ACT_LOGNORM_PDF(x, 0, 1);
            var cdf = Distributions.ACT_LOGNORM_CDF(x, 0, 1);
            Log($"| {x} | {pdf:F6} | {cdf:F6} |");
        }
        Log($"\nInverse CDF: P(0.5) = {Distributions.ACT_LOGNORM_INV(0.5, 0, 1):F6}\n");

        Log("### Gamma Distribution (α=2, β=1)");
        Log("| x | PDF | CDF |");
        Log("|---|-----|-----|");
        foreach (var x in new[] { 0.0, 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            var pdf = Distributions.ACT_GAMMA_PDF(x, 2, 1);
            var cdf = Distributions.ACT_GAMMA_CDF(x, 2, 1);
            Log($"| {x} | {pdf:F6} | {cdf:F6} |");
        }
        Log($"\nInverse CDF: P(0.5) = {Distributions.ACT_GAMMA_INV(0.5, 2, 1):F6}\n");

        Log("### Pareto Distribution (α=2, xm=1)");
        Log("| x | PDF | CDF |");
        Log("|---|-----|-----|");
        foreach (var x in new[] { 0.0, 1.0, 1.5, 2.0, 3.0, 5.0, 10.0 })
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

    static void TestReturnPeriods()
    {
        Log("## 4. Return Period Tables\n");

        var returnPeriods = new double[] { 10, 25, 50, 100, 250 };
        var losses = new double[] { 100_000, 250_000, 500_000, 1_000_000, 2_500_000 };
        var targets = new double[] { 20, 75, 150, 200 };

        Log("### Return Period Table (LOG interpolation)");
        var table = Reinsurance.ACT_RETURN_PERIOD_TABLE(returnPeriods, losses, targets, "LOG");
        Log("| Target RP | Loss |");
        Log("|-----------|------|");
        for (int i = 0; i < table.GetLength(0); i++)
        {
            Log($"| {table[i, 0]} | {table[i, 1]:N0} |");
        }
        Log("");
    }

    static void TestInterpolation()
    {
        Log("## 5. Interpolation\n");

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
        Log("## 6. Chain Ladder (Taylor-Ashe Dataset)\n");
        Log("Using the Taylor-Ashe triangle from Peter England's bootstrapping presentation.\n");

        // Taylor-Ashe triangle (cumulative paid losses)
        // Source: England & Verrall (1999), Taylor & Ashe (1983)
        var triangle = new double[,]
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

        Log("### Input Triangle (Cumulative Paid Losses - Taylor-Ashe)");
        Log("| AY | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 |");
        Log("|--|--|--|--|--|--|--|--|--|--|--|");
        for (int i = 0; i < 10; i++)
        {
            var row = $"| {i + 1} |";
            for (int j = 0; j < 10; j++)
            {
                row += triangle[i, j] > 0 ? $" {triangle[i, j]:N0} |" : " - |";
            }
            Log(row);
        }

        Log("\n### Development Factors");
        Log("Expected: 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018");
        var factors = ChainLadder.ACT_CL_FACTORS(triangle);
        Log("| Period | Factor |");
        Log("|--------|--------|");
        for (int i = 0; i < factors.Length; i++)
        {
            Log($"| {i + 1}-{i + 2} | {factors[i]:F4} |");
        }

        Log("\n### Latest Diagonal");
        var latestDiag = ChainLadder.ACT_CL_LATEST(triangle);
        Log("| AY | Latest |");
        Log("|----|--------|");
        for (int i = 0; i < latestDiag.Length; i++)
        {
            Log($"| {i + 1} | {latestDiag[i]:N0} |");
        }

        Log("\n### Development Factors (Variations)");
        var factorsTop5 = ChainLadder.ACT_CL_FACTORS(triangle, 5, false, false);
        var factorsExcludeRecent = ChainLadder.ACT_CL_FACTORS(triangle, double.PositiveInfinity, true, false);
        var factorsExcludeHighLow = ChainLadder.ACT_CL_FACTORS(triangle, double.PositiveInfinity, false, true);
        Log("| Period | Base | Top5 | ExclRecent | ExclHighLow |");
        Log("|--------|------|------|------------|-------------|");
        for (int i = 0; i < factors.Length; i++)
        {
            Log($"| {i + 1}-{i + 2} | {factors[i]:F4} | {factorsTop5[i]:F4} | {factorsExcludeRecent[i]:F4} | {factorsExcludeHighLow[i]:F4} |");
        }

        Log("\n### Projected Ultimates and IBNR");
        Log("Expected Total IBNR: ~18,680,856");
        var ultimates = ChainLadder.ACT_CL_ULTIMATE(triangle);
        var ibnr = ChainLadder.ACT_CL_IBNR(triangle);
        Log("| AY | Ultimate | IBNR |");
        Log("|----|----------|------|");
        double totalIBNR = 0;
        for (int i = 0; i < ultimates.Length; i++)
        {
            Log($"| {i + 1} | {ultimates[i]:N0} | {ibnr[i]:N0} |");
            totalIBNR += (double)ibnr[i];
        }
        Log($"| **Total** | | **{totalIBNR:N0}** |");

        Log("\n### Mack Standard Errors");
        Log("Reference values from england.pdf (Bootstrapping Mack's Model, analytic column):");
        Log("Expected Total Reserve SE: 2,447,095 (analytic) vs 2,454,616 (simulated)");
        var reserveSE = ChainLadder.ACT_MACK_RESERVE_SE(triangle);
        Log("| AY | Reserve SE | Expected SE |");
        Log("|----|------------|-------------|");
        var expectedSE = new double[] { 0, 75535, 121699, 133549, 261406, 411010, 558317, 875328, 971258, 1363155 };
        double totalSE = CalculateMackTotalSe(triangle);
        for (int i = 0; i < reserveSE.Length; i++)
        {
            double se = (double)reserveSE[i];
            Log($"| {i + 1} | {se:N0} | {expectedSE[i]:N0} |");
        }
        Log($"| **Total** | **{totalSE:N0}** | **2,447,095** |");
        Log("");

        Log("### Bornhuetter-Ferguson Ultimates (A priori = Latest * 1.10)");
        var apriori = latestDiag.Cast<double>().Select(x => x * 1.10).ToArray();
        var bfUltimates = ChainLadder.ACT_BF_ULTIMATE(triangle, factors.Cast<double>().ToArray(), apriori);
        Log("| AY | A Priori | BF Ultimate |");
        Log("|----|----------|-------------|");
        for (int i = 0; i < bfUltimates.Length; i++)
        {
            Log($"| {i + 1} | {apriori[i]:N0} | {bfUltimates[i]:N0} |");
        }
        Log("");
    }

    static double[,] GetTaylorAsheTriangle()
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

    static void TestBootstrapChainLadder()
    {
        Log("## 7. Bootstrap Chain Ladder\n");
        Log("Using Taylor-Ashe data with ODP bootstrap (England & Verrall, 2002 method).\n");
        Log("Reference: england.pdf slide 35 (ODP constant vs non-constant scale parameters)\n");

        var triangle = GetTaylorAsheTriangle();

        Log("### Bootstrap Results (10,000 iterations, seed=42)");
        Log("Expected mean reserve: ~18,680,856 (same as deterministic chain ladder)");
        Log("Expected total prediction error: 2,992,296 (constant scale), 2,228,677 (non-constant scale)");
        Log("Expected bootstrap SE: ~3,000,000 - 4,000,000 (includes process variance)\n");

        var bootstrap = ChainLadder.ACT_BOOTSTRAP_CL(triangle, 10000, 42);

        Log("| Statistic | Value |");
        Log("|-----------|-------|");
        for (int i = 0; i < bootstrap.GetLength(0); i++)
        {
            var label = bootstrap[i, 0]?.ToString() ?? "";
            var value = bootstrap[i, 1];
            if (value is double d)
                Log($"| {label} | {d:N0} |");
            else
                Log($"| {label} | {value} |");
        }

        Log("\n### Reconciliation to england.pdf");
        Log("Taylor-Ashe ODP totals (constant vs non-constant scale): 2,992,296 vs 2,228,677");
        Log("Typical percentiles for the total reserve (ODP bootstrap):");
        Log("- 75th percentile: ~20-21M");
        Log("- 95th percentile: ~24-26M");
        Log("- 99th percentile: ~27-30M");
        Log("\n### Bootstrap by Origin Year (1,000 iterations, seed=42)");
        var originStats = ChainLadder.ACT_BOOTSTRAP_CL_ORIGIN(triangle, 1000, 42);
        Log("| AY | Mean | StdDev | P50 | P75 | P90 | P95 | P99 |");
        Log("|----|------|--------|-----|-----|-----|-----|-----|");
        for (int i = 1; i < originStats.GetLength(0); i++)
        {
            Log($"| {originStats[i, 0]} | {originStats[i, 1]:N0} | {originStats[i, 2]:N0} | {originStats[i, 3]:N0} | {originStats[i, 4]:N0} | {originStats[i, 5]:N0} | {originStats[i, 6]:N0} | {originStats[i, 7]:N0} |");
        }
        Log("");
    }

    static void TestBerquistSherman()
    {
        Log("## 8. Berquist-Sherman Adjustment\n");
        Log("Demonstrates case reserve adequacy adjustment for changing reserve practices.\n");

        // Simple example triangle for Berquist-Sherman
        var paidTriangle = new double[,]
        {
            { 1000, 1800, 2200, 2400, 2500 },
            { 1100, 2000, 2500, 2700, 0 },
            { 1200, 2200, 2700, 0, 0 },
            { 1300, 2400, 0, 0, 0 },
            { 1400, 0, 0, 0, 0 }
        };

        var caseTriangle = new double[,]
        {
            { 500, 300, 150, 80, 50 },
            { 600, 350, 180, 100, 0 },
            { 700, 400, 200, 0, 0 },
            { 800, 450, 0, 0, 0 },
            { 900, 0, 0, 0, 0 }
        };

        Log("### Input Paid Triangle");
        Log("| AY | 1 | 2 | 3 | 4 | 5 |");
        Log("|--|--|--|--|--|--|");
        for (int i = 0; i < 5; i++)
        {
            var row = $"| {i + 1} |";
            for (int j = 0; j < 5; j++)
                row += paidTriangle[i, j] > 0 ? $" {paidTriangle[i, j]:N0} |" : " - |";
            Log(row);
        }

        Log("\n### Input Case Reserve Triangle");
        Log("| AY | 1 | 2 | 3 | 4 | 5 |");
        Log("|--|--|--|--|--|--|");
        for (int i = 0; i < 5; i++)
        {
            var row = $"| {i + 1} |";
            for (int j = 0; j < 5; j++)
                row += (j <= 4 - i && caseTriangle[i, j] > 0) ? $" {caseTriangle[i, j]:N0} |" : " - |";
            Log(row);
        }

        Log("\n### Adjusted Paid Triangle (5% annual case reserve strengthening)");
        var adjusted = ChainLadder.ACT_BERQUIST_SHERMAN(paidTriangle, caseTriangle, 0.05);
        Log("| AY | 1 | 2 | 3 | 4 | 5 |");
        Log("|--|--|--|--|--|--|");
        for (int i = 0; i < 5; i++)
        {
            var row = $"| {i + 1} |";
            for (int j = 0; j < 5; j++)
            {
                var val = adjusted[i, j];
                if (val is double d && d > 0)
                    row += $" {d:N0} |";
                else
                    row += " - |";
            }
            Log(row);
        }
        Log("");
    }

    static void TestCopulas()
    {
        Log("## 9. Student-t Copula\n");

        int dim = 7;
        double rho = 0.6;
        var corrMatrix = new double[dim, dim];
        for (int i = 0; i < dim; i++)
        {
            for (int j = 0; j < dim; j++)
                corrMatrix[i, j] = Math.Pow(rho, Math.Abs(i - j));
        }

        Log($"### {dim}x{dim} Correlation Matrix (rho={rho:F1})");
        var labels = Enumerable.Range(1, dim).Select(i => $"X{i}").ToArray();
        var header = "| | " + string.Join(" | ", labels) + " |";
        Log(header);
        Log("|--|" + string.Join("|", labels.Select(_ => "---")) + "|");
        for (int i = 0; i < dim; i++)
        {
            var row = $"| {labels[i]} |";
            for (int j = 0; j < dim; j++)
                row += $" {corrMatrix[i, j]:F2} |";
            Log(row);
        }

        Log("\n### Generated Samples (df=5, 10 samples, seed=42)");
        var samples = Copulas.ACT_STUDENT_T_COPULA(corrMatrix, 5, 10, 42);

        if (samples[0, 0] is string errorMsg)
        {
            Log($"Error: {errorMsg}");
        }
        else
        {
            Log("| Sample | " + string.Join(" | ", labels.Select((_, idx) => $"U{idx + 1}")) + " |");
            Log("|--------|" + string.Join("|", labels.Select(_ => "---")) + "|");
            for (int i = 0; i < 10; i++)
            {
                var row = $"| {i + 1} |";
                for (int j = 0; j < dim; j++)
                    row += $" {samples[i, j]:F4} |";
                Log(row);
            }
        }

        Log("\n### Correlation Validation");
        Log("Generating 10,000 samples to verify correlation structure is preserved...");
        var largeSample = Copulas.ACT_STUDENT_T_COPULA(corrMatrix, 5, 10000, 123);

        if (largeSample[0, 0] is not string)
        {
            // Calculate empirical correlations
            var data = new double[10000, dim];
            for (int i = 0; i < 10000; i++)
                for (int j = 0; j < dim; j++)
                    data[i, j] = (double)largeSample[i, j];

            Log("| Pair | Target | Empirical |");
            Log("|------|--------|-----------|");
            var pairs = new (int, int)[] { (0, 1), (0, 2), (1, 3), (2, 4), (4, 6) };
            foreach (var (a, b) in pairs)
            {
                double corr = CalculateCorrelation(data, a, b);
                Log($"| X{a + 1}-X{b + 1} | {corrMatrix[a, b]:F2} | {corr:F2} |");
            }
        }
        Log("");
    }

    static double CalculateMackTotalSe(double[,] triangle)
    {
        int n = triangle.GetLength(0);
        var factors = ChainLadder.ACT_CL_FACTORS(triangle).Cast<double>().ToArray();
        var ultimates = ChainLadder.ACT_CL_ULTIMATE(triangle).Cast<double>().ToArray();

        var sigma2 = new double[n - 1];
        var sumC = new double[n - 1];
        for (int j = 0; j < n - 1; j++)
        {
            double sumWeightedVar = 0;
            int count = n - j - 1;
            for (int i = 0; i < count; i++)
            {
                if (triangle[i, j] > 0 && triangle[i, j + 1] > 0)
                {
                    double f_ij = triangle[i, j + 1] / triangle[i, j];
                    sumWeightedVar += triangle[i, j] * Math.Pow(f_ij - factors[j], 2);
                }
            }

            sigma2[j] = count > 1 ? sumWeightedVar / (count - 1) : 0;
            sumC[j] = GetSumC(triangle, j);
        }

        if (sigma2.Length >= 3 && sigma2[^1] <= 0)
            sigma2[^1] = Math.Min(sigma2[^2], sigma2[^3]);

        var term = sigma2.Select((s, idx) => s / (factors[idx] * factors[idx])).ToArray();
        var lastCols = Enumerable.Range(0, n).Select(i => n - 1 - i).ToArray();

        double totalVar = 0;
        for (int i = 1; i < n; i++)
        {
            double proc = 0;
            double est = 0;
            for (int j = lastCols[i]; j < n - 1; j++)
            {
                double Cij = GetProjected(triangle, factors, i, j);
                proc += term[j] / Cij;
                est += term[j] / sumC[j];
            }
            totalVar += ultimates[i] * ultimates[i] * (proc + est);
        }

        for (int i = 1; i < n; i++)
        {
            for (int k = i + 1; k < n; k++)
            {
                int maxCol = Math.Max(lastCols[i], lastCols[k]);
                double cov = 0;
                for (int j = maxCol; j < n - 1; j++)
                    cov += term[j] / sumC[j];
                totalVar += 2 * ultimates[i] * ultimates[k] * cov;
            }
        }

        return Math.Sqrt(totalVar);
    }

    static double GetSumC(double[,] triangle, int col)
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

    static double GetProjected(double[,] triangle, double[] factors, int row, int col)
    {
        int n = triangle.GetLength(0);
        int lastKnownCol = n - 1 - row;
        if (col <= lastKnownCol)
            return triangle[row, col];

        double value = triangle[row, lastKnownCol];
        for (int j = lastKnownCol; j < col; j++)
            value *= factors[j];
        return value;
    }

    static double CalculateCorrelation(double[,] data, int col1, int col2)
    {
        int n = data.GetLength(0);
        double sum1 = 0, sum2 = 0, sum1Sq = 0, sum2Sq = 0, sumProd = 0;
        for (int i = 0; i < n; i++)
        {
            sum1 += data[i, col1];
            sum2 += data[i, col2];
            sum1Sq += data[i, col1] * data[i, col1];
            sum2Sq += data[i, col2] * data[i, col2];
            sumProd += data[i, col1] * data[i, col2];
        }
        double mean1 = sum1 / n, mean2 = sum2 / n;
        double var1 = sum1Sq / n - mean1 * mean1;
        double var2 = sum2Sq / n - mean2 * mean2;
        double cov = sumProd / n - mean1 * mean2;
        return cov / Math.Sqrt(var1 * var2);
    }
}
