using System.Text;
using System.Text.Json;
using System.Linq;
using ActuarialAddIn.Functions;

namespace ActuarialAddIn.Tests;

class Program
{
    static StringBuilder _output = new();
    static string _fixturesPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "fixtures");
    static JsonDocument? _distributionsFixture;
    static JsonDocument? _exposureCurvesFixture;
    static JsonDocument? _chainLadderFixture;

    static void Main(string[] args)
    {
        // Load fixtures
        LoadFixtures();

        _output.AppendLine("# Actuarial Add-In Test Results");
        _output.AppendLine($"\nTest run: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");

        WriteRunInstructions();

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

    static void LoadFixtures()
    {
        try
        {
            var distPath = Path.Combine(_fixturesPath, "distributions.json");
            if (File.Exists(distPath))
                _distributionsFixture = JsonDocument.Parse(File.ReadAllText(distPath));

            var expPath = Path.Combine(_fixturesPath, "exposure_curves.json");
            if (File.Exists(expPath))
                _exposureCurvesFixture = JsonDocument.Parse(File.ReadAllText(expPath));

            var clPath = Path.Combine(_fixturesPath, "chain_ladder.json");
            if (File.Exists(clPath))
                _chainLadderFixture = JsonDocument.Parse(File.ReadAllText(clPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not load fixtures: {ex.Message}");
        }
    }

    static void WriteRunInstructions()
    {
        Log("## How to Run These Tests\n");
        Log("### C# Test Suite (generates this file)");
        Log("```bash");
        Log("cd src/ActuarialAddIn.Tests");
        Log("dotnet run");
        Log("```\n");
        Log("### Python Validation (pytest)");
        Log("```bash");
        Log("cd tests");
        Log("pip install -r requirements.txt");
        Log("pytest test_validation.py -v");
        Log("```\n");
        Log("---\n");
    }

    static bool WithinTolerance(double actual, double expected, double tolerance)
    {
        if (double.IsNaN(actual) && double.IsNaN(expected)) return true;
        if (double.IsInfinity(actual) && double.IsInfinity(expected)) return true;
        return Math.Abs(actual - expected) <= tolerance;
    }

    static string FormatMatch(bool match) => match ? "TRUE" : "**FALSE**";

    // Helper to extract 1D array from Excel object[,] return (assumes vertical column)
    static double[] ToDoubleArray(object excelResult)
    {
        if (excelResult is object[,] arr2d)
        {
            int len = arr2d.GetLength(0);
            var result = new double[len];
            for (int i = 0; i < len; i++)
                result[i] = Convert.ToDouble(arr2d[i, 0]);
            return result;
        }
        throw new InvalidOperationException("Expected object[,] from Excel function");
    }

    static void TestDistributions()
    {
        Log("## 1. Statistical Distributions\n");

        // Poisson Distribution
        Log("### Poisson Distribution (λ=5)");
        Log("| k  | PDF      | CDF      |");
        Log("|----|----------|----------|");
        for (int k = 0; k <= 10; k++)
        {
            var pdf = Distributions.ACT_DIST_POISSON_PDF(k, 5);
            var cdf = Distributions.ACT_DIST_POISSON_CDF(k, 5);
            Log($"| {k,-2} | {pdf:F6} | {cdf:F6} |");
        }
        var poissonInv = Distributions.ACT_DIST_POISSON_INV(0.5, 5);
        Log($"\nInverse CDF: P(0.5) = {poissonInv}");

        // Poisson Reconciliation
        Log("\n#### Reconciliation: Poisson");
        Log("Reference: `scipy.stats.poisson` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.poisson.html))\n");
        Log("| Test | C# | Python | Function | Tol | Match |");
        Log("|------|-----|--------|----------|-----|-------|");
        var pdfK5 = Distributions.ACT_DIST_POISSON_PDF(5, 5);
        var pdfK0 = Distributions.ACT_DIST_POISSON_PDF(0, 5);
        var cdfK5 = Distributions.ACT_DIST_POISSON_CDF(5, 5);
        Log($"| PDF(k=5) | {pdfK5:F10} | 0.1754673698 | `poisson.pmf(5, 5)` | 1e-6 | {FormatMatch(WithinTolerance(pdfK5, 0.17546736976785068, 1e-6))} |");
        Log($"| PDF(k=0) | {pdfK0:F10} | 0.0067379470 | `poisson.pmf(0, 5)` | 1e-6 | {FormatMatch(WithinTolerance(pdfK0, 0.006737946999085467, 1e-6))} |");
        Log($"| CDF(k=5) | {cdfK5:F10} | 0.6159606548 | `poisson.cdf(5, 5)` | 1e-6 | {FormatMatch(WithinTolerance(cdfK5, 0.615960654833063, 1e-6))} |");
        Log($"| INV(p=0.5) | {poissonInv} | 5 | `poisson.ppf(0.5, 5)` | 0 | {FormatMatch(poissonInv == 5)} |");
        Log("");

        // Negative Binomial Distribution
        Log("### Negative Binomial Distribution (r=5, p=0.3)");
        Log("| k  | PDF      | CDF      |");
        Log("|----|----------|----------|");
        for (int k = 0; k <= 10; k++)
        {
            var pdf = Distributions.ACT_DIST_NEGBIN_PDF(k, 5, 0.3);
            var cdf = Distributions.ACT_DIST_NEGBIN_CDF(k, 5, 0.3);
            Log($"| {k,-2} | {pdf:F6} | {cdf:F6} |");
        }
        var nbInv = Distributions.ACT_DIST_NEGBIN_INV(0.5, 5, 0.3);
        Log($"\nInverse CDF: P(0.5) = {nbInv}");

        // Negative Binomial Reconciliation
        Log("\n#### Reconciliation: Negative Binomial");
        Log("Reference: `scipy.stats.nbinom` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.nbinom.html))");
        Log("Note: scipy uses (n, p) where n=r successes, p=success probability\n");
        Log("| Test | C# | Python | Function | Tol | Match |");
        Log("|------|-----|--------|----------|-----|-------|");
        var nbPdfK5 = Distributions.ACT_DIST_NEGBIN_PDF(5, 5, 0.3);
        var nbPdfK10 = Distributions.ACT_DIST_NEGBIN_PDF(10, 5, 0.3);
        var nbCdfK10 = Distributions.ACT_DIST_NEGBIN_CDF(10, 5, 0.3);
        Log($"| PDF(k=5) | {nbPdfK5:F10} | 0.0514596726 | `nbinom.pmf(5, 5, 0.3)` | 1e-6 | {FormatMatch(WithinTolerance(nbPdfK5, 0.0514596726, 1e-6))} |");
        Log($"| PDF(k=10) | {nbPdfK10:F10} | 0.0687101270 | `nbinom.pmf(10, 5, 0.3)` | 1e-6 | {FormatMatch(WithinTolerance(nbPdfK10, 0.06871012699250698, 1e-6))} |");
        Log($"| CDF(k=10) | {nbCdfK10:F10} | 0.4845089408 | `nbinom.cdf(10, 5, 0.3)` | 1e-6 | {FormatMatch(WithinTolerance(nbCdfK10, 0.48450894077315665, 1e-6))} |");
        Log($"| INV(p=0.5) | {nbInv} | 11 | `nbinom.ppf(0.5, 5, 0.3)` | 0 | {FormatMatch(nbInv == 11)} |");
        Log("");

        // Lognormal Distribution
        Log("### Lognormal Distribution (μ=0, σ=1)");
        Log("| x   | PDF      | CDF      |");
        Log("|-----|----------|----------|");
        foreach (var x in new[] { 0.0, 0.5, 1.0, 1.5, 2.0, 3.0, 5.0 })
        {
            var pdf = Distributions.ACT_DIST_LOGNORM_PDF(x, 0, 1);
            var cdf = Distributions.ACT_DIST_LOGNORM_CDF(x, 0, 1);
            Log($"| {x,-3} | {pdf:F6} | {cdf:F6} |");
        }
        var lnInv = Distributions.ACT_DIST_LOGNORM_INV(0.5, 0, 1);
        Log($"\nInverse CDF: P(0.5) = {lnInv:F6}");

        // Lognormal Reconciliation
        Log("\n#### Reconciliation: Lognormal");
        Log("Reference: `scipy.stats.lognorm` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.lognorm.html))");
        Log("Note: scipy lognorm uses s=sigma, scale=exp(mu)\n");
        Log("| Test | C# | Python | Function | Tol | Match |");
        Log("|------|-----|--------|----------|-----|-------|");
        var lnPdfX1 = Distributions.ACT_DIST_LOGNORM_PDF(1.0, 0, 1);
        var lnPdfX2 = Distributions.ACT_DIST_LOGNORM_PDF(2.0, 0, 1);
        var lnCdfX1 = Distributions.ACT_DIST_LOGNORM_CDF(1.0, 0, 1);
        var lnCdfX2 = Distributions.ACT_DIST_LOGNORM_CDF(2.0, 0, 1);
        Log($"| PDF(x=1) | {lnPdfX1:F10} | 0.3989422804 | `lognorm.pdf(1, s=1, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(lnPdfX1, 0.3989422804014327, 1e-6))} |");
        Log($"| PDF(x=2) | {lnPdfX2:F10} | 0.1568740193 | `lognorm.pdf(2, s=1, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(lnPdfX2, 0.15687401927898112, 1e-6))} |");
        Log($"| CDF(x=1) | {lnCdfX1:F10} | 0.5000000000 | `lognorm.cdf(1, s=1, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(lnCdfX1, 0.5, 1e-6))} |");
        Log($"| CDF(x=2) | {lnCdfX2:F10} | 0.7558914042 | `lognorm.cdf(2, s=1, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(lnCdfX2, 0.7558914042144173, 1e-6))} |");
        Log($"| INV(p=0.5) | {lnInv:F10} | 1.0000000000 | `lognorm.ppf(0.5, s=1, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(lnInv, 1.0, 1e-6))} |");
        Log("");

        // Gamma Distribution
        Log("### Gamma Distribution (α=2, β=1)");
        Log("| x   | PDF      | CDF      |");
        Log("|-----|----------|----------|");
        foreach (var x in new[] { 0.0, 0.5, 1.0, 2.0, 3.0, 5.0 })
        {
            var pdf = Distributions.ACT_DIST_GAMMA_PDF(x, 2, 1);
            var cdf = Distributions.ACT_DIST_GAMMA_CDF(x, 2, 1);
            Log($"| {x,-3} | {pdf:F6} | {cdf:F6} |");
        }
        var gammaInv = Distributions.ACT_DIST_GAMMA_INV(0.5, 2, 1);
        Log($"\nInverse CDF: P(0.5) = {gammaInv:F6}");

        // Gamma Reconciliation
        Log("\n#### Reconciliation: Gamma");
        Log("Reference: `scipy.stats.gamma` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.gamma.html))");
        Log("Note: scipy gamma uses a=alpha (shape), scale=1/beta\n");
        Log("| Test | C# | Python | Function | Tol | Match |");
        Log("|------|-----|--------|----------|-----|-------|");
        var gammaPdfX1 = Distributions.ACT_DIST_GAMMA_PDF(1.0, 2, 1);
        var gammaPdfX2 = Distributions.ACT_DIST_GAMMA_PDF(2.0, 2, 1);
        var gammaCdfX2 = Distributions.ACT_DIST_GAMMA_CDF(2.0, 2, 1);
        Log($"| PDF(x=1) | {gammaPdfX1:F10} | 0.3678794412 | `gamma.pdf(1, a=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(gammaPdfX1, 0.36787944117144233, 1e-6))} |");
        Log($"| PDF(x=2) | {gammaPdfX2:F10} | 0.2706705665 | `gamma.pdf(2, a=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(gammaPdfX2, 0.2706705664732254, 1e-6))} |");
        Log($"| CDF(x=2) | {gammaCdfX2:F10} | 0.5939941503 | `gamma.cdf(2, a=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(gammaCdfX2, 0.5939941502901616, 1e-6))} |");
        Log($"| INV(p=0.5) | {gammaInv:F10} | 1.6783469900 | `gamma.ppf(0.5, a=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(gammaInv, 1.6783469900166612, 1e-6))} |");
        Log("");

        // Pareto Distribution
        Log("### Pareto Distribution (α=2, xm=1)");
        Log("| x   | PDF      | CDF      |");
        Log("|-----|----------|----------|");
        foreach (var x in new[] { 0.0, 1.0, 1.5, 2.0, 3.0, 5.0, 10.0 })
        {
            var pdf = Distributions.ACT_DIST_PARETO_PDF(x, 2, 1);
            var cdf = Distributions.ACT_DIST_PARETO_CDF(x, 2, 1);
            Log($"| {x,-3} | {pdf:F6} | {cdf:F6} |");
        }
        var paretoInv = Distributions.ACT_DIST_PARETO_INV(0.5, 2, 1);
        Log($"\nInverse CDF: P(0.5) = {paretoInv:F6}");

        // Pareto Reconciliation
        Log("\n#### Reconciliation: Pareto");
        Log("Reference: `scipy.stats.pareto` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.pareto.html))");
        Log("Note: scipy pareto uses b=alpha, scale=xm\n");
        Log("| Test | C# | Python | Function | Tol | Match |");
        Log("|------|-----|--------|----------|-----|-------|");
        var paretoPdfX1 = Distributions.ACT_DIST_PARETO_PDF(1.0, 2, 1);
        var paretoPdfX2 = Distributions.ACT_DIST_PARETO_PDF(2.0, 2, 1);
        var paretoCdfX2 = Distributions.ACT_DIST_PARETO_CDF(2.0, 2, 1);
        Log($"| PDF(x=1) | {paretoPdfX1:F10} | 2.0000000000 | `pareto.pdf(1, b=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(paretoPdfX1, 2.0, 1e-6))} |");
        Log($"| PDF(x=2) | {paretoPdfX2:F10} | 0.2500000000 | `pareto.pdf(2, b=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(paretoPdfX2, 0.25, 1e-6))} |");
        Log($"| CDF(x=2) | {paretoCdfX2:F10} | 0.7500000000 | `pareto.cdf(2, b=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(paretoCdfX2, 0.75, 1e-6))} |");
        Log($"| INV(p=0.5) | {paretoInv:F10} | 1.4142135624 | `pareto.ppf(0.5, b=2, scale=1)` | 1e-6 | {FormatMatch(WithinTolerance(paretoInv, 1.4142135623730951, 1e-6))} |");
        Log("");
    }

    static void TestExposureCurves()
    {
        Log("## 2. Exposure Curves\n");

        Log("### MBBEFD Curves (b=2, g=3)");
        Log("| d   | G(d)     |");
        Log("|-----|----------|");
        foreach (var d in new[] { 0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0 })
        {
            var g = ExposureCurves.ACT_EXPOSURE_MBBEFD(d, 2, 3);
            Log($"| {d:F1} | {g:F6} |");
        }

        Log("\n### Swiss Re Curves Comparison at d=0.5");
        Log("| Curve | G(0.5)   |");
        Log("|-------|----------|");
        for (int c = 1; c <= 5; c++)
        {
            var g = ExposureCurves.ACT_EXPOSURE_SWISSRE(0.5, c);
            Log($"| {c}     | {g:F6} |");
        }

        Log("\n### Lloyd's Curves Comparison at d=0.5");
        Log("| Curve | G(0.5)   |");
        Log("|-------|----------|");
        foreach (var curve in new[] { "Y1", "Y2", "Y3", "Y4" })
        {
            var g = ExposureCurves.ACT_EXPOSURE_LLOYDS(0.5, curve);
            Log($"| {curve}    | {g:F6} |");
        }

        Log("\n### Power and Pareto Curves at d=0.5");
        var powerN2 = ExposureCurves.ACT_EXPOSURE_POWER(0.5, 2);
        var invPowerN2 = ExposureCurves.ACT_EXPOSURE_INVERSE_POWER(0.5, 2);
        var paretoExp = ExposureCurves.ACT_EXPOSURE_PARETO(0.5, 2);
        Log($"- Power curve (n=2): {powerN2:F6}");
        Log($"- Inverse power (n=2): {invPowerN2:F6}");
        Log($"- Pareto exposure (α=2): {paretoExp:F6}");

        // Exposure Curves Reconciliation
        Log("\n### Reconciliation: Exposure Curves");
        Log("Reference: Bernegger (1997) \"The Swiss Re Exposure Curves and the MBBEFD Distribution Class\"");
        Log("([ASTIN Bulletin](https://doi.org/10.2143/AST.27.1.563208))\n");
        Log("| Test | C# | Reference | Source | Tol | Match |");
        Log("|------|-----|-----------|--------|-----|-------|");

        // MBBEFD at d=0.5 with b=2, g=3
        var mbbefd05 = ExposureCurves.ACT_EXPOSURE_MBBEFD(0.5, 2, 3);
        Log($"| MBBEFD(d=0.5, b=2, g=3) | {mbbefd05:F6} | 0.5960 | Bernegger formula | 0.01 | {FormatMatch(WithinTolerance(mbbefd05, 0.596, 0.01))} |");

        // Power curve: G(d) = d^n
        Log($"| Power(d=0.5, n=2) | {powerN2:F6} | 0.2500 | G(d)=d^n | 0.0001 | {FormatMatch(WithinTolerance(powerN2, 0.25, 0.0001))} |");

        // Inverse power: G(d) = 1 - (1-d)^n
        Log($"| InvPower(d=0.5, n=2) | {invPowerN2:F6} | 0.7500 | G(d)=1-(1-d)^n | 0.0001 | {FormatMatch(WithinTolerance(invPowerN2, 0.75, 0.0001))} |");

        // Pareto exposure: G(d) = 1 - (1-d)^((α-1)/α)
        Log($"| Pareto(d=0.5, α=2) | {paretoExp:F6} | 0.2929 | G(d)=1-(1-d)^((α-1)/α) | 0.001 | {FormatMatch(WithinTolerance(paretoExp, 0.2929, 0.001))} |");
        Log("");
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
        var factors = ToDoubleArray(ChainLadder.ACT_CL_FACTORS(triangle));
        Log("| Period | Factor |");
        Log("|--------|--------|");
        for (int i = 0; i < factors.Length; i++)
        {
            Log($"| {i + 1}-{i + 2} | {factors[i]:F4} |");
        }

        Log("\n### Latest Diagonal");
        var latestDiag = ToDoubleArray(ChainLadder.ACT_CL_LATEST(triangle));
        Log("| AY | Latest |");
        Log("|----|--------|");
        for (int i = 0; i < latestDiag.Length; i++)
        {
            Log($"| {i + 1} | {latestDiag[i]:N0} |");
        }

        Log("\n### Development Factors (Variations)");
        var factorsTop5 = ToDoubleArray(ChainLadder.ACT_CL_FACTORS(triangle, 5, false, false));
        var factorsExcludeRecent = ToDoubleArray(ChainLadder.ACT_CL_FACTORS(triangle, double.PositiveInfinity, true, false));
        var factorsExcludeHighLow = ToDoubleArray(ChainLadder.ACT_CL_FACTORS(triangle, double.PositiveInfinity, false, true));
        Log("| Period | Base | Top5 | ExclRecent | ExclHighLow |");
        Log("|--------|------|------|------------|-------------|");
        for (int i = 0; i < factors.Length; i++)
        {
            Log($"| {i + 1}-{i + 2} | {factors[i]:F4} | {factorsTop5[i]:F4} | {factorsExcludeRecent[i]:F4} | {factorsExcludeHighLow[i]:F4} |");
        }

        Log("\n### Projected Ultimates and IBNR");
        Log("Expected Total IBNR: ~18,680,856");
        var ultimates = ToDoubleArray(ChainLadder.ACT_CL_ULTIMATE(triangle));
        var ibnr = ToDoubleArray(ChainLadder.ACT_CL_IBNR(triangle));
        Log("| AY | Ultimate | IBNR |");
        Log("|----|----------|------|");
        double totalIBNR = 0;
        for (int i = 0; i < ultimates.Length; i++)
        {
            Log($"| {i + 1} | {ultimates[i]:N0} | {ibnr[i]:N0} |");
            totalIBNR += ibnr[i];
        }
        Log($"| **Total** | | **{totalIBNR:N0}** |");

        Log("\n### Mack Standard Errors");
        Log("Reference values from england.pdf (Bootstrapping Mack's Model, analytic column):");
        Log("Expected Total Reserve SE: 2,447,095 (analytic) vs 2,454,616 (simulated)");
        var reserveSE = ToDoubleArray(ChainLadder.ACT_MACK_RESERVE_SE(triangle));
        Log("| AY | Reserve SE | Expected SE |");
        Log("|----|------------|-------------|");
        var expectedSE = new double[] { 0, 75535, 121699, 133549, 261406, 411010, 558317, 875328, 971258, 1363155 };
        double totalSE = CalculateMackTotalSe(triangle);
        for (int i = 0; i < reserveSE.Length; i++)
        {
            Log($"| {i + 1} | {reserveSE[i]:N0} | {expectedSE[i]:N0} |");
        }
        Log($"| **Total** | **{totalSE:N0}** | **2,447,095** |");
        Log("");

        Log("### Bornhuetter-Ferguson Ultimates (A priori = Latest * 1.10)");
        var apriori = latestDiag.Select(x => x * 1.10).ToArray();
        var bfUltimates = ToDoubleArray(ChainLadder.ACT_BF_ULTIMATE(triangle, factors, apriori));
        Log("| AY | A Priori | BF Ultimate |");
        Log("|----|----------|-------------|");
        for (int i = 0; i < bfUltimates.Length; i++)
        {
            Log($"| {i + 1} | {apriori[i]:N0} | {bfUltimates[i]:N0} |");
        }

        // Chain Ladder Reconciliation
        Log("\n### Reconciliation: Chain Ladder");
        Log("References:");
        Log("- England & Verrall (2002) \"Stochastic Claims Reserving\" ([British Actuarial Journal](https://doi.org/10.1017/S1357321700003809))");
        Log("- Mack (1993) \"Distribution-free calculation of the standard error\" ([ASTIN Bulletin](https://doi.org/10.2143/AST.23.2.2005092))");
        Log("- `chainladder-python` package ([docs](https://chainladder-python.readthedocs.io/))\n");

        // Expected development factors from E&V
        double[] expectedFactors = { 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018 };

        Log("| Test | C# | Reference | Source | Tol | Match |");
        Log("|------|-----|-----------|--------|-----|-------|");

        // Development factors
        for (int i = 0; i < Math.Min(3, factors.Length); i++)
        {
            Log($"| Factor {i+1}-{i+2} | {factors[i]:F4} | {expectedFactors[i]:F3} | E&V 2002 / chainladder-python | 0.01 | {FormatMatch(WithinTolerance(factors[i], expectedFactors[i], 0.01))} |");
        }

        // Total IBNR
        Log($"| Total IBNR | {totalIBNR:N0} | 18,680,856 | E&V 2002 Table 3 | 1000 | {FormatMatch(WithinTolerance(totalIBNR, 18680856, 1000))} |");

        // Total Mack SE
        Log($"| Total Mack SE | {totalSE:N0} | 2,447,095 | E&V 2002 (analytic) | 1000 | {FormatMatch(WithinTolerance(totalSE, 2447095, 1000))} |");

        // Individual reserve SEs
        Log($"| Mack SE (AY2) | {reserveSE[1]:N0} | 75,535 | E&V 2002 | 100 | {FormatMatch(WithinTolerance(reserveSE[1], 75535, 100))} |");
        Log($"| Mack SE (AY5) | {reserveSE[4]:N0} | 261,406 | E&V 2002 | 100 | {FormatMatch(WithinTolerance(reserveSE[4], 261406, 100))} |");
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

        var triangle = GetTaylorAsheTriangle();

        Log("### Deterministic Chain Ladder Results");
        var factors = ToDoubleArray(ChainLadder.ACT_CL_FACTORS(triangle));
        Log("Development Factors:");
        for (int j = 0; j < factors.Length; j++)
            Log($"  {j+1}-{j+2}: {factors[j]:F4}");

        var ibnr = ToDoubleArray(ChainLadder.ACT_CL_IBNR(triangle));
        Log("\nIBNR by Origin:");
        double totalIBNR = 0;
        for (int i = 0; i < ibnr.Length; i++)
        {
            totalIBNR += ibnr[i];
            Log($"  AY{i+1}: {ibnr[i]:N0}");
        }
        Log($"  Total: {totalIBNR:N0}");
        Log($"Expected Total IBNR: 18,680,856\n");

        Log("### Bootstrap Results (10,000 iterations, seed=123)");
        Log("Reference: E&V 2002 ODP Bootstrap (non-constant scale target)\n");

        var originStats = (object[,])ChainLadder.ACT_CL_BOOTSTRAP_ORIGIN(triangle, 10000, 123);

        Log("| AY | Mean | StdDev | E&V Non-Const | Ratio |");
        Log("|----|------|--------|---------------|-------|");
        // E&V 2002 Non-Constant Scale ODP Bootstrap results (Taylor-Ashe)
        double[] evNonConst = {0, 43882, 109449, 141509, 256031, 398377, 529898, 735245, 809457, 1285560};
        for (int i = 1; i < originStats.GetLength(0); i++)
        {
            double ourSE = Convert.ToDouble(originStats[i, 2]);
            double evSE = evNonConst[i-1];
            string ratio = evSE > 0 ? $"{ourSE / evSE:P0}" : "-";
            Log($"| {originStats[i, 0]} | {originStats[i, 1]:N0} | {originStats[i, 2]:N0} | {evSE:N0} | {ratio} |");
        }

        var bootstrap = (object[,])ChainLadder.ACT_CL_BOOTSTRAP(triangle, 10000, 123);
        double totalSE = Convert.ToDouble(bootstrap[1, 1]);
        Log($"\nTotal Bootstrap: Mean={bootstrap[0, 1]:N0}, StdDev={totalSE:N0}");
        Log($"E&V Non-Constant Total SE: 2,228,677 (Ratio: {totalSE / 2228677:P0})");
        Log("\nImplementation: E&V 2002 ODP Bootstrap with non-constant scale parameters.");
        Log("Period-specific phi values and stratified residual sampling.");

        // Bootstrap Reconciliation
        Log("\n### Reconciliation: Bootstrap");
        Log("Reference: England & Verrall (2002) Table 3 - ODP Bootstrap with non-constant scale");
        Log("Note: `chainladder-python` uses constant scale (different methodology)\n");
        Log("| Test | C# | E&V 2002 | Ratio | Tol | Match |");
        Log("|------|-----|----------|-------|-----|-------|");
        Log($"| Total SE | {totalSE:N0} | 2,228,677 | {totalSE / 2228677:P0} | 10% | {FormatMatch(Math.Abs(totalSE / 2228677 - 1.0) <= 0.10)} |");

        // Check a few individual origin years
        double se6 = Convert.ToDouble(originStats[6, 2]);
        double se10 = Convert.ToDouble(originStats[10, 2]);
        Log($"| AY6 SE | {se6:N0} | 398,377 | {se6 / 398377:P0} | 15% | {FormatMatch(Math.Abs(se6 / 398377 - 1.0) <= 0.15)} |");
        Log($"| AY10 SE | {se10:N0} | 1,285,560 | {se10 / 1285560:P0} | 15% | {FormatMatch(Math.Abs(se10 / 1285560 - 1.0) <= 0.15)} |");
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
        var samples = Copulas.ACT_COPULA_STUDENT_T(corrMatrix, 5, 10, 42);

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
        var largeSample = Copulas.ACT_COPULA_STUDENT_T(corrMatrix, 5, 10000, 123);

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

            // Copula Reconciliation
            Log("\n### Reconciliation: Student-t Copula");
            Log("Reference: `scipy.stats.multivariate_t` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.multivariate_t.html))");
            Log("Note: Copula generates uniform marginals with specified correlation structure\n");
            Log("| Test | Target | Empirical | Tol | Match |");
            Log("|------|--------|-----------|-----|-------|");

            // Check that empirical correlations match targets (with sampling error tolerance)
            foreach (var (a, b) in new (int, int)[] { (0, 1), (1, 3) })
            {
                double target = corrMatrix[a, b];
                double empirical = CalculateCorrelation(data, a, b);
                bool match = WithinTolerance(empirical, target, 0.05);
                Log($"| Corr(X{a+1}, X{b+1}) | {target:F2} | {empirical:F2} | 0.05 | {FormatMatch(match)} |");
            }

            // Check that values are in [0,1] (uniform marginals)
            bool allInRange = true;
            for (int i = 0; i < 1000 && allInRange; i++)
                for (int j = 0; j < dim && allInRange; j++)
                    if (data[i, j] < 0 || data[i, j] > 1) allInRange = false;
            Log($"| Uniform [0,1] range | TRUE | {allInRange} | - | {FormatMatch(allInRange)} |");
        }
        Log("");
    }

    static double CalculateMackTotalSe(double[,] triangle)
    {
        int n = triangle.GetLength(0);
        var factors = ToDoubleArray(ChainLadder.ACT_CL_FACTORS(triangle));
        var ultimates = ToDoubleArray(ChainLadder.ACT_CL_ULTIMATE(triangle));

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
