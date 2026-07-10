using System.Text;
using System.Text.Json;
using System.Linq;
using ActuarialAddIn.Functions;

namespace ActuarialAddIn.Tests;

class Program
{
    static StringBuilder _output = new();
    static string _fixturesPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "fixtures");
    static int _passCount = 0;
    static int _failCount = 0;

    static int Main(string[] args)
    {
        _output.AppendLine("# Actuarial Add-In Test Results");
        _output.AppendLine($"\nTest run: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");

        WriteRunInstructions();

        TestDistributions();
        TestInverseGaussian();
        TestLoglogistic();
        TestZeroTruncated();
        TestZeroModified();
        TestParetoExtended();
        TestPanjer();
        TestLEV();
        TestExposureCurves();
        TestReinsurance();
        TestReturnPeriods();
        TestCatModeling();
        TestInterpolation();
        TestChainLadder();
        TestBootstrapChainLadder();
        TestCopulas();

        var summary = $"\n## Summary\n\n{_passCount} passed, {_failCount} failed (of {_passCount + _failCount} assertions).\n";
        _output.Append(summary);

        var outputPath = args.Length > 0 ? args[0] : "test_results.md";
        File.WriteAllText(outputPath, _output.ToString());
        Console.WriteLine(_output.ToString());
        Console.WriteLine($"\nResults written to: {outputPath}");

        // Emit the structured JSON snapshot consumed by tests/reconciliation.ipynb.
        var jsonPath = Path.Combine(_fixturesPath, "addin_outputs.json");
        Directory.CreateDirectory(Path.GetDirectoryName(jsonPath)!);
        AddinOutputsEmitter.Emit(jsonPath);

        Console.WriteLine($"Assertions: {_passCount} passed, {_failCount} failed.");
        return _failCount == 0 ? 0 : 1;
    }

    static void Log(string message) => _output.AppendLine(message);

    static void WriteRunInstructions()
    {
        Log("## How to Run These Tests\n");
        Log("### C# Test Suite (generates this file plus the JSON fixture)");
        Log("```bash");
        Log("cd src/ActuarialAddIn.Tests");
        Log("dotnet run");
        Log("```\n");
        Log("### Python Reconciliation (papermill)");
        Log("```bash");
        Log("cd tests");
        Log("pip install -r requirements.txt");
        Log("python build_reconciliation_notebook.py");
        Log("papermill reconciliation.ipynb executed.ipynb");
        Log("```\n");
        Log("---\n");
    }

    static bool WithinTolerance(double actual, double expected, double tolerance)
    {
        if (double.IsNaN(actual) && double.IsNaN(expected)) return true;
        if (double.IsInfinity(actual) && double.IsInfinity(expected)) return true;
        return Math.Abs(actual - expected) <= tolerance;
    }

    static string FormatMatch(bool match)
    {
        if (match) _passCount++; else _failCount++;
        return match ? "TRUE " : "FALSE";
    }

    // Helper for aligned table rows
    static string TableRow(params (string val, int width)[] cols)
    {
        var sb = new StringBuilder("|");
        foreach (var (val, width) in cols)
            sb.Append($" {val.PadRight(width)} |");
        return sb.ToString();
    }

    static string TableSep(params int[] widths)
    {
        var sb = new StringBuilder("|");
        foreach (var w in widths)
            sb.Append(new string('-', w + 2) + "|");
        return sb.ToString();
    }

    // Helper to extract 1D array from Excel return (handles object[,] or object[])
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
        if (excelResult is object[] arr1d)
        {
            var result = new double[arr1d.Length];
            for (int i = 0; i < arr1d.Length; i++)
                result[i] = Convert.ToDouble(arr1d[i]);
            return result;
        }
        if (excelResult is double[] dblArr)
        {
            return dblArr;
        }
        throw new InvalidOperationException($"Expected object[,] or object[] from Excel function, got {excelResult?.GetType()}");
    }

    static double[] ExtractColumn(object[,] table, int columnIndex, bool hasHeader)
    {
        int startRow = hasHeader ? 1 : 0;
        int len = table.GetLength(0) - startRow;
        var result = new double[len];
        for (int i = 0; i < len; i++)
        {
            result[i] = Convert.ToDouble(table[i + startRow, columnIndex]);
        }
        return result;
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
        Log(TableRow(("Test", 12), ("C#", 14), ("Python", 14), ("Function", 22), ("Tol", 6), ("Match", 5)));
        Log(TableSep(12, 14, 14, 22, 6, 5));
        var pdfK5 = Distributions.ACT_DIST_POISSON_PDF(5, 5);
        var pdfK0 = Distributions.ACT_DIST_POISSON_PDF(0, 5);
        var cdfK5 = Distributions.ACT_DIST_POISSON_CDF(5, 5);
        Log(TableRow(("PDF(k=5)", 12), ($"{pdfK5:F10}", 14), ("0.1754673698", 14), ("`poisson.pmf(5, 5)`", 22), ("1e-6", 6), (FormatMatch(WithinTolerance(pdfK5, 0.17546736976785068, 1e-6)), 5)));
        Log(TableRow(("PDF(k=0)", 12), ($"{pdfK0:F10}", 14), ("0.0067379470", 14), ("`poisson.pmf(0, 5)`", 22), ("1e-6", 6), (FormatMatch(WithinTolerance(pdfK0, 0.006737946999085467, 1e-6)), 5)));
        Log(TableRow(("CDF(k=5)", 12), ($"{cdfK5:F10}", 14), ("0.6159606548", 14), ("`poisson.cdf(5, 5)`", 22), ("1e-6", 6), (FormatMatch(WithinTolerance(cdfK5, 0.615960654833063, 1e-6)), 5)));
        Log(TableRow(("INV(p=0.5)", 12), ($"{poissonInv}", 14), ("5", 14), ("`poisson.ppf(0.5, 5)`", 22), ("0", 6), (FormatMatch(poissonInv == 5), 5)));
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
        Log(TableRow(("Test", 12), ("C#", 14), ("Python", 14), ("Function", 26), ("Tol", 6), ("Match", 5)));
        Log(TableSep(12, 14, 14, 26, 6, 5));
        var nbPdfK5 = Distributions.ACT_DIST_NEGBIN_PDF(5, 5, 0.3);
        var nbPdfK10 = Distributions.ACT_DIST_NEGBIN_PDF(10, 5, 0.3);
        var nbCdfK10 = Distributions.ACT_DIST_NEGBIN_CDF(10, 5, 0.3);
        Log(TableRow(("PDF(k=5)", 12), ($"{nbPdfK5:F10}", 14), ("0.0514596726", 14), ("`nbinom.pmf(5, 5, 0.3)`", 26), ("1e-6", 6), (FormatMatch(WithinTolerance(nbPdfK5, 0.0514596726, 1e-6)), 5)));
        Log(TableRow(("PDF(k=10)", 12), ($"{nbPdfK10:F10}", 14), ("0.0687101270", 14), ("`nbinom.pmf(10, 5, 0.3)`", 26), ("1e-6", 6), (FormatMatch(WithinTolerance(nbPdfK10, 0.06871012699250698, 1e-6)), 5)));
        Log(TableRow(("CDF(k=10)", 12), ($"{nbCdfK10:F10}", 14), ("0.4845089408", 14), ("`nbinom.cdf(10, 5, 0.3)`", 26), ("1e-6", 6), (FormatMatch(WithinTolerance(nbCdfK10, 0.48450894077315665, 1e-6)), 5)));
        Log(TableRow(("INV(p=0.5)", 12), ($"{nbInv}", 14), ("11", 14), ("`nbinom.ppf(0.5, 5, 0.3)`", 26), ("0", 6), (FormatMatch(nbInv == 11), 5)));
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
        Log(TableRow(("Test", 12), ("C#", 14), ("Python", 14), ("Function", 34), ("Tol", 6), ("Match", 5)));
        Log(TableSep(12, 14, 14, 34, 6, 5));
        var lnPdfX1 = Distributions.ACT_DIST_LOGNORM_PDF(1.0, 0, 1);
        var lnPdfX2 = Distributions.ACT_DIST_LOGNORM_PDF(2.0, 0, 1);
        var lnCdfX1 = Distributions.ACT_DIST_LOGNORM_CDF(1.0, 0, 1);
        var lnCdfX2 = Distributions.ACT_DIST_LOGNORM_CDF(2.0, 0, 1);
        Log(TableRow(("PDF(x=1)", 12), ($"{lnPdfX1:F10}", 14), ("0.3989422804", 14), ("`lognorm.pdf(1, s=1, scale=1)`", 34), ("1e-6", 6), (FormatMatch(WithinTolerance(lnPdfX1, 0.3989422804014327, 1e-6)), 5)));
        Log(TableRow(("PDF(x=2)", 12), ($"{lnPdfX2:F10}", 14), ("0.1568740193", 14), ("`lognorm.pdf(2, s=1, scale=1)`", 34), ("1e-6", 6), (FormatMatch(WithinTolerance(lnPdfX2, 0.15687401927898112, 1e-6)), 5)));
        Log(TableRow(("CDF(x=1)", 12), ($"{lnCdfX1:F10}", 14), ("0.5000000000", 14), ("`lognorm.cdf(1, s=1, scale=1)`", 34), ("1e-6", 6), (FormatMatch(WithinTolerance(lnCdfX1, 0.5, 1e-6)), 5)));
        Log(TableRow(("CDF(x=2)", 12), ($"{lnCdfX2:F10}", 14), ("0.7558914042", 14), ("`lognorm.cdf(2, s=1, scale=1)`", 34), ("1e-6", 6), (FormatMatch(WithinTolerance(lnCdfX2, 0.7558914042144173, 1e-6)), 5)));
        Log(TableRow(("INV(p=0.5)", 12), ($"{lnInv:F10}", 14), ("1.0000000000", 14), ("`lognorm.ppf(0.5, s=1, scale=1)`", 34), ("1e-6", 6), (FormatMatch(WithinTolerance(lnInv, 1.0, 1e-6)), 5)));
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
        Log(TableRow(("Test", 12), ("C#", 14), ("Python", 14), ("Function", 30), ("Tol", 6), ("Match", 5)));
        Log(TableSep(12, 14, 14, 30, 6, 5));
        var gammaPdfX1 = Distributions.ACT_DIST_GAMMA_PDF(1.0, 2, 1);
        var gammaPdfX2 = Distributions.ACT_DIST_GAMMA_PDF(2.0, 2, 1);
        var gammaCdfX2 = Distributions.ACT_DIST_GAMMA_CDF(2.0, 2, 1);
        Log(TableRow(("PDF(x=1)", 12), ($"{gammaPdfX1:F10}", 14), ("0.3678794412", 14), ("`gamma.pdf(1, a=2, scale=1)`", 30), ("1e-6", 6), (FormatMatch(WithinTolerance(gammaPdfX1, 0.36787944117144233, 1e-6)), 5)));
        Log(TableRow(("PDF(x=2)", 12), ($"{gammaPdfX2:F10}", 14), ("0.2706705665", 14), ("`gamma.pdf(2, a=2, scale=1)`", 30), ("1e-6", 6), (FormatMatch(WithinTolerance(gammaPdfX2, 0.2706705664732254, 1e-6)), 5)));
        Log(TableRow(("CDF(x=2)", 12), ($"{gammaCdfX2:F10}", 14), ("0.5939941503", 14), ("`gamma.cdf(2, a=2, scale=1)`", 30), ("1e-6", 6), (FormatMatch(WithinTolerance(gammaCdfX2, 0.5939941502901616, 1e-6)), 5)));
        Log(TableRow(("INV(p=0.5)", 12), ($"{gammaInv:F10}", 14), ("1.6783469900", 14), ("`gamma.ppf(0.5, a=2, scale=1)`", 30), ("1e-6", 6), (FormatMatch(WithinTolerance(gammaInv, 1.6783469900166612, 1e-6)), 5)));
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
        Log(TableRow(("Test", 12), ("C#", 14), ("Python", 14), ("Function", 32), ("Tol", 6), ("Match", 5)));
        Log(TableSep(12, 14, 14, 32, 6, 5));
        var paretoPdfX1 = Distributions.ACT_DIST_PARETO_PDF(1.0, 2, 1);
        var paretoPdfX2 = Distributions.ACT_DIST_PARETO_PDF(2.0, 2, 1);
        var paretoCdfX2 = Distributions.ACT_DIST_PARETO_CDF(2.0, 2, 1);
        Log(TableRow(("PDF(x=1)", 12), ($"{paretoPdfX1:F10}", 14), ("2.0000000000", 14), ("`pareto.pdf(1, b=2, scale=1)`", 32), ("1e-6", 6), (FormatMatch(WithinTolerance(paretoPdfX1, 2.0, 1e-6)), 5)));
        Log(TableRow(("PDF(x=2)", 12), ($"{paretoPdfX2:F10}", 14), ("0.2500000000", 14), ("`pareto.pdf(2, b=2, scale=1)`", 32), ("1e-6", 6), (FormatMatch(WithinTolerance(paretoPdfX2, 0.25, 1e-6)), 5)));
        Log(TableRow(("CDF(x=2)", 12), ($"{paretoCdfX2:F10}", 14), ("0.7500000000", 14), ("`pareto.cdf(2, b=2, scale=1)`", 32), ("1e-6", 6), (FormatMatch(WithinTolerance(paretoCdfX2, 0.75, 1e-6)), 5)));
        Log(TableRow(("INV(p=0.5)", 12), ($"{paretoInv:F10}", 14), ("1.4142135624", 14), ("`pareto.ppf(0.5, b=2, scale=1)`", 32), ("1e-6", 6), (FormatMatch(WithinTolerance(paretoInv, 1.4142135623730951, 1e-6)), 5)));
        Log("");
    }

    static void TestInverseGaussian()
    {
        Log("## Inverse Gaussian (Wald) Distribution\n");
        Log("Parameters: μ=2.0 (mean), λ=3.0 (shape)");
        Log("Used in operational risk modeling and as a GLM family.\n");

        double mu = 2.0;
        double lambda = 3.0;

        // PDF tests
        Log("### PDF Values");
        Log("| x   | PDF      | Expected (scipy) |");
        Log("|-----|----------|------------------|");
        var pdfTests = new (double x, double expected)[]
        {
            (1.0, 0.4749088496333091),
            (2.0, 0.24430125595146002),
            (3.0, 0.11735510892143318)
        };
        foreach (var (x, expected) in pdfTests)
        {
            var pdf = Distributions.ACT_DIST_INVGAUSS_PDF(x, mu, lambda);
            Log($"| {x:F1} | {pdf:F10} | {expected:F10} |");
        }

        // CDF tests
        Log("\n### CDF Values");
        Log("| x   | CDF      | Expected (scipy) |");
        Log("|-----|----------|------------------|");
        var cdfTests = new (double x, double expected)[]
        {
            (1.0, 0.28738674440477363),
            (2.0, 0.6436706247667282),
            (3.0, 0.8161869234555278)
        };
        foreach (var (x, expected) in cdfTests)
        {
            var cdf = Distributions.ACT_DIST_INVGAUSS_CDF(x, mu, lambda);
            Log($"| {x:F1} | {cdf:F10} | {expected:F10} |");
        }

        // Inverse CDF tests
        Log("\n### Inverse CDF (Quantile) Values");
        Log("| p   | INV      | Expected (scipy) |");
        Log("|-----|----------|------------------|");
        var invTests = new (double p, double expected)[]
        {
            (0.25, 0.9222895777740259),
            (0.50, 1.512250663605367),
            (0.75, 2.5263919895700826),
            (0.90, 3.984738795776087),
            (0.95, 5.177395529298311),
            (0.99, 8.187156145754324)
        };
        foreach (var (p, expected) in invTests)
        {
            var inv = Distributions.ACT_DIST_INVGAUSS_INV(p, mu, lambda);
            Log($"| {p:F2} | {inv:F10} | {expected:F10} |");
        }

        // LEV tests
        Log("\n### LEV Values");
        Log("| Limit | LEV      | Expected (scipy) |");
        Log("|-------|----------|------------------|");
        var levTests = new (double limit, double expected)[]
        {
            (1.0, 0.9107922283280907),
            (2.0, 1.4253175009330339),
            (5.0, 1.8964701630288672)
        };
        foreach (var (limit, expected) in levTests)
        {
            var lev = Distributions.ACT_DIST_INVGAUSS_LEV(limit, mu, lambda);
            Log($"| {limit:F1} | {lev:F10} | {expected:F10} |");
        }

        // Reconciliation table
        Log("\n### Reconciliation: Inverse Gaussian");
        Log("Reference: `scipy.stats.invgauss` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.invgauss.html))");
        Log("Note: scipy parameterization uses mu_scipy = mu/lambda, scale = lambda\n");
        Log(TableRow(("Test", 12), ("C#", 14), ("scipy", 14), ("Tol", 8), ("Match", 5)));
        Log(TableSep(12, 14, 14, 8, 5));

        var pdf1 = Distributions.ACT_DIST_INVGAUSS_PDF(1.0, mu, lambda);
        var cdf2 = Distributions.ACT_DIST_INVGAUSS_CDF(2.0, mu, lambda);
        var inv50 = Distributions.ACT_DIST_INVGAUSS_INV(0.5, mu, lambda);
        var lev2 = Distributions.ACT_DIST_INVGAUSS_LEV(2.0, mu, lambda);

        Log(TableRow(("PDF(x=1)", 12), ($"{pdf1:F10}", 14), ("0.4749088496", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(pdf1, 0.4749088496333091, 1e-6)), 5)));
        Log(TableRow(("CDF(x=2)", 12), ($"{cdf2:F10}", 14), ("0.6436706248", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(cdf2, 0.6436706247667282, 1e-6)), 5)));
        Log(TableRow(("INV(p=0.5)", 12), ($"{inv50:F10}", 14), ("1.5122506636", 14), ("1e-4", 8), (FormatMatch(WithinTolerance(inv50, 1.512250663605367, 1e-4)), 5)));
        Log(TableRow(("LEV(d=2)", 12), ($"{lev2:F10}", 14), ("1.4253175009", 14), ("1e-3", 8), (FormatMatch(WithinTolerance(lev2, 1.4253175009330339, 1e-3)), 5)));
        Log("");
    }

    static void TestLoglogistic()
    {
        Log("## Loglogistic Distribution\n");
        Log("Parameters: α=2.0 (scale), β=3.0 (shape)");
        Log("Popular for liability claims; heavier tail than lognormal, lighter than Pareto.\n");

        double alpha = 2.0;
        double beta = 3.0;

        // PDF tests
        Log("### PDF Values");
        Log("| x   | PDF      | Expected (scipy.fisk) |");
        Log("|-----|----------|------------------------|");
        var pdfTests = new (double x, double expected)[]
        {
            (1.0, 0.2962962962962963),
            (2.0, 0.375),
            (3.0, 0.1763265306122449),
            (5.0, 0.033919384928486625)
        };
        foreach (var (x, expected) in pdfTests)
        {
            var pdf = Distributions.ACT_DIST_LOGLOGISTIC_PDF(x, alpha, beta);
            Log($"| {x:F1} | {pdf:F10} | {expected:F10} |");
        }

        // CDF tests
        Log("\n### CDF Values");
        Log("| x   | CDF      | Expected (scipy.fisk) |");
        Log("|-----|----------|------------------------|");
        var cdfTests = new (double x, double expected)[]
        {
            (1.0, 0.1111111111111111),
            (2.0, 0.5),
            (3.0, 0.7714285714285715),
            (5.0, 0.9398496240601504)
        };
        foreach (var (x, expected) in cdfTests)
        {
            var cdf = Distributions.ACT_DIST_LOGLOGISTIC_CDF(x, alpha, beta);
            Log($"| {x:F1} | {cdf:F10} | {expected:F10} |");
        }

        // Inverse CDF tests
        Log("\n### Inverse CDF (Quantile) Values");
        Log("| p   | INV      | Expected (scipy.fisk) |");
        Log("|-----|----------|------------------------|");
        var invTests = new (double p, double expected)[]
        {
            (0.25, 1.3867225487012695),
            (0.50, 2.0),
            (0.75, 2.884499140614817),
            (0.90, 4.160167646103807),
            (0.95, 5.3368032974438915),
            (0.99, 9.252130018365463)
        };
        foreach (var (p, expected) in invTests)
        {
            var inv = Distributions.ACT_DIST_LOGLOGISTIC_INV(p, alpha, beta);
            Log($"| {p:F2} | {inv:F10} | {expected:F10} |");
        }

        // LEV tests
        Log("\n### LEV Values");
        Log("| Limit | LEV      | Expected (scipy) |");
        Log("|-------|----------|------------------|");
        var levTests = new (double limit, double expected)[]
        {
            (1.0, 0.9708038843007759),
            (2.0, 1.671297696529442),
            (5.0, 2.2623385740103537)
        };
        foreach (var (limit, expected) in levTests)
        {
            var lev = Distributions.ACT_DIST_LOGLOGISTIC_LEV(limit, alpha, beta);
            Log($"| {limit:F1} | {lev:F10} | {expected:F10} |");
        }

        // Reconciliation table
        Log("\n### Reconciliation: Loglogistic");
        Log("Reference: `scipy.stats.fisk` ([docs](https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.fisk.html))");
        Log("Note: scipy.stats.fisk uses c=beta (shape), scale=alpha\n");
        Log(TableRow(("Test", 12), ("C#", 14), ("scipy.fisk", 14), ("Tol", 8), ("Match", 5)));
        Log(TableSep(12, 14, 14, 8, 5));

        var pdf2 = Distributions.ACT_DIST_LOGLOGISTIC_PDF(2.0, alpha, beta);
        var cdf2 = Distributions.ACT_DIST_LOGLOGISTIC_CDF(2.0, alpha, beta);
        var inv50 = Distributions.ACT_DIST_LOGLOGISTIC_INV(0.5, alpha, beta);
        var levVal = Distributions.ACT_DIST_LOGLOGISTIC_LEV(2.0, alpha, beta);

        Log(TableRow(("PDF(x=2)", 12), ($"{pdf2:F10}", 14), ("0.3750000000", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(pdf2, 0.375, 1e-6)), 5)));
        Log(TableRow(("CDF(x=2)", 12), ($"{cdf2:F10}", 14), ("0.5000000000", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(cdf2, 0.5, 1e-6)), 5)));
        Log(TableRow(("INV(p=0.5)", 12), ($"{inv50:F10}", 14), ("2.0000000000", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(inv50, 2.0, 1e-6)), 5)));
        Log(TableRow(("LEV(d=2)", 12), ($"{levVal:F10}", 14), ("1.6712976965", 14), ("1e-3", 8), (FormatMatch(WithinTolerance(levVal, 1.671297696529442, 1e-3)), 5)));
        Log("");
    }

    static void TestZeroTruncated()
    {
        Log("## Zero-Truncated Distributions\n");
        Log("Zero-truncated distributions condition on X > 0, used in claim count modeling.\n");

        // Zero-Truncated Poisson (lambda=5)
        Log("### Zero-Truncated Poisson (λ=5)");
        Log("| k | PMF | CDF | scipy Expected PMF |");
        Log("|---|-----|-----|-------------------|");
        var ztpTests = new (int k, double expPmf, double expCdf)[]
        {
            (1, 0.03391827453152116, 0.033918274531521145),
            (2, 0.08479568632880287, 0.11871396086032399),
            (3, 0.14132614388133807, 0.26004010474166206),
            (5, 0.17665767985167266, 0.6133554644450074),
            (10, 0.01825579528890465, 0.9862118274255965)
        };
        foreach (var (k, expPmf, expCdf) in ztpTests)
        {
            var pmf = Distributions.ACT_DIST_ZTPOISSON_PDF(k, 5);
            var cdf = Distributions.ACT_DIST_ZTPOISSON_CDF(k, 5);
            Log($"| {k} | {pmf:F10} | {cdf:F10} | {expPmf:F10} |");
        }
        var ztpMean = Distributions.ACT_DIST_ZTPOISSON_MEAN(5);
        Log($"\nMean: {ztpMean:F10} (expected: 5.033918274531521)");
        var ztpInv50 = Distributions.ACT_DIST_ZTPOISSON_INV(0.5, 5);
        Log($"Quantile(0.5): {ztpInv50} (expected: 5)");

        // Zero-Truncated Negative Binomial (r=5, p=0.3)
        Log("\n### Zero-Truncated Negative Binomial (r=5, p=0.3)");
        Log("| k | PMF | CDF | scipy Expected PMF |");
        Log("|---|-----|-----|-------------------|");
        var ztnbTests = new (int k, double expPmf, double expCdf)[]
        {
            (1, 0.008525717493509223, 0.008525717493509223),
            (2, 0.017904006736369375, 0.026429724229878597),
            (3, 0.029243211002736646, 0.05567293523261523),
            (5, 0.05158502420882745, 0.14819845484527394),
            (10, 0.06887749931584448, 0.4832532461613287)
        };
        foreach (var (k, expPmf, expCdf) in ztnbTests)
        {
            var pmf = Distributions.ACT_DIST_ZTNEGBIN_PDF(k, 5, 0.3);
            var cdf = Distributions.ACT_DIST_ZTNEGBIN_CDF(k, 5, 0.3);
            Log($"| {k} | {pmf:F10} | {cdf:F10} | {expPmf:F10} |");
        }
        var ztnbMean = Distributions.ACT_DIST_ZTNEGBIN_MEAN(5, 0.3);
        Log($"\nMean: {ztnbMean:F10} (expected: 11.695085724978366)");

        // Zero-Truncated Binomial (n=10, p=0.3)
        Log("\n### Zero-Truncated Binomial (n=10, p=0.3)");
        Log("| k | PMF | CDF | scipy Expected PMF |");
        Log("|---|-----|-----|-------------------|");
        var ztbTests = new (int k, double expPmf, double expCdf)[]
        {
            (1, 0.12457989467692573, 0.12457989467692576),
            (2, 0.2402612254483571, 0.36484112012528264),
            (3, 0.27458425765526495, 0.6394253777805481),
            (5, 0.10591107080988786, 0.9512746418318848),
            (10, 6.076547424684813e-06, 1.0)
        };
        foreach (var (k, expPmf, expCdf) in ztbTests)
        {
            var pmf = Distributions.ACT_DIST_ZTBINOM_PDF(k, 10, 0.3);
            var cdf = Distributions.ACT_DIST_ZTBINOM_CDF(k, 10, 0.3);
            Log($"| {k} | {pmf:E10} | {cdf:F10} | {expPmf:E10} |");
        }
        var ztbMean = Distributions.ACT_DIST_ZTBINOM_MEAN(10, 0.3);
        Log($"\nMean: {ztbMean:F10} (expected: 3.0872059262738483)");

        // Zero-Truncated Geometric (p=0.3)
        Log("\n### Zero-Truncated Geometric (p=0.3)");
        Log("| k | PMF | CDF | Expected PMF |");
        Log("|---|-----|-----|--------------|");
        var ztgTests = new (int k, double expPmf, double expCdf)[]
        {
            (1, 0.3, 0.3),
            (2, 0.21, 0.51),
            (3, 0.147, 0.657),
            (5, 0.07203, 0.83193),
            (10, 0.0121060821, 0.9717524751)
        };
        foreach (var (k, expPmf, expCdf) in ztgTests)
        {
            var pmf = Distributions.ACT_DIST_ZTGEOM_PDF(k, 0.3);
            var cdf = Distributions.ACT_DIST_ZTGEOM_CDF(k, 0.3);
            Log($"| {k} | {pmf:F10} | {cdf:F10} | {expPmf} |");
        }
        var ztgMean = Distributions.ACT_DIST_ZTGEOM_MEAN(0.3);
        Log($"\nMean: {ztgMean:F10} (expected: 3.3333333333333335)");

        // Reconciliation table
        Log("\n### Reconciliation: Zero-Truncated Distributions");
        Log("Reference: scipy.stats with zero-truncation formula f_ZT(k) = P(k) / (1 - P(0))\n");
        Log(TableRow(("Distribution", 15), ("Test", 12), ("C#", 14), ("scipy", 14), ("Tol", 8), ("Match", 5)));
        Log(TableSep(15, 12, 14, 14, 8, 5));

        // ZT Poisson
        var ztpPdf5 = Distributions.ACT_DIST_ZTPOISSON_PDF(5, 5);
        Log(TableRow(("ZT Poisson", 15), ("PMF(k=5)", 12), ($"{ztpPdf5:F10}", 14), ("0.1766576799", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztpPdf5, 0.17665767985167266, 1e-10)), 5)));

        var ztpCdf5 = Distributions.ACT_DIST_ZTPOISSON_CDF(5, 5);
        Log(TableRow(("ZT Poisson", 15), ("CDF(k=5)", 12), ($"{ztpCdf5:F10}", 14), ("0.6133554644", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztpCdf5, 0.6133554644450074, 1e-10)), 5)));

        Log(TableRow(("ZT Poisson", 15), ("Mean", 12), ($"{ztpMean:F10}", 14), ("5.0339182745", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztpMean, 5.033918274531521, 1e-10)), 5)));

        // ZT Neg Binomial
        var ztnbPdf5 = Distributions.ACT_DIST_ZTNEGBIN_PDF(5, 5, 0.3);
        Log(TableRow(("ZT NegBin", 15), ("PMF(k=5)", 12), ($"{ztnbPdf5:F10}", 14), ("0.0515850242", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztnbPdf5, 0.05158502420882745, 1e-10)), 5)));

        Log(TableRow(("ZT NegBin", 15), ("Mean", 12), ($"{ztnbMean:F10}", 14), ("11.695085725", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztnbMean, 11.695085724978366, 1e-10)), 5)));

        // ZT Binomial
        var ztbPdf3 = Distributions.ACT_DIST_ZTBINOM_PDF(3, 10, 0.3);
        Log(TableRow(("ZT Binom", 15), ("PMF(k=3)", 12), ($"{ztbPdf3:F10}", 14), ("0.2745842577", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztbPdf3, 0.27458425765526495, 1e-10)), 5)));

        Log(TableRow(("ZT Binom", 15), ("Mean", 12), ($"{ztbMean:F10}", 14), ("3.0872059263", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztbMean, 3.0872059262738483, 1e-10)), 5)));

        // ZT Geometric
        var ztgPdf2 = Distributions.ACT_DIST_ZTGEOM_PDF(2, 0.3);
        Log(TableRow(("ZT Geom", 15), ("PMF(k=2)", 12), ($"{ztgPdf2:F10}", 14), ("0.2100000000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztgPdf2, 0.21, 1e-10)), 5)));

        Log(TableRow(("ZT Geom", 15), ("Mean", 12), ($"{ztgMean:F10}", 14), ("3.3333333333", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(ztgMean, 3.3333333333333335, 1e-10)), 5)));
        Log("");
    }

    static void TestZeroModified()
    {
        Log("## Zero-Modified (Zero-Inflated) Distributions\n");
        Log("Zero-modified distributions add extra probability mass at zero.\n");
        Log("PMF(0) = p0 + (1-p0)*g(0), PMF(k) = (1-p0)*g(k) for k >= 1\n");

        // Zero-Modified Poisson (lambda=5, p0=0.2)
        Log("### Zero-Modified Poisson (λ=5, p0=0.2)");
        Log("| k | PMF | CDF | scipy Expected PMF |");
        Log("|---|-----|-----|-------------------|");
        var zmpTests = new (int k, double expPmf, double expCdf)[]
        {
            (0, 0.20539035759926838, 0.20539035759926838),
            (1, 0.02695178799634187, 0.23234214559561026),
            (2, 0.06737946999085466, 0.2997216155864649),
            (3, 0.1122991166514244, 0.41202073223788926),
            (5, 0.14037389581428056, 0.6927685238664504),
            (10, 0.014506230966257484, 0.9890437851212937)
        };
        foreach (var (k, expPmf, expCdf) in zmpTests)
        {
            var pmf = Distributions.ACT_DIST_ZMPOISSON_PDF(k, 5, 0.2);
            var cdf = Distributions.ACT_DIST_ZMPOISSON_CDF(k, 5, 0.2);
            Log($"| {k} | {pmf:F10} | {cdf:F10} | {expPmf:F10} |");
        }
        var zmpMean = Distributions.ACT_DIST_ZMPOISSON_MEAN(5, 0.2);
        var zmpVar = Distributions.ACT_DIST_ZMPOISSON_VAR(5, 0.2);
        Log($"\nMean: {zmpMean:F10} (expected: 4.0)");
        Log($"Variance: {zmpVar:F10} (expected: 8.0)");
        var zmpInv50 = Distributions.ACT_DIST_ZMPOISSON_INV(0.5, 5, 0.2);
        Log($"Quantile(0.5): {zmpInv50} (expected: 4)");

        // Zero-Modified Negative Binomial (r=5, p=0.3, p0=0.2)
        Log("\n### Zero-Modified Negative Binomial (r=5, p=0.3, p0=0.2)");
        Log("| k | PMF | CDF | scipy Expected PMF |");
        Log("|---|-----|-----|-------------------|");
        var zmnbTests = new (int k, double expPmf, double expCdf)[]
        {
            (0, 0.201944, 0.201944),
            (1, 0.006803999999999997, 0.20874800000000002),
            (2, 0.0142884, 0.2230364),
            (3, 0.023337719999999996, 0.24637412),
            (5, 0.04116773808, 0.32021466607999993),
            (10, 0.054968101594005586, 0.5876071526185254)
        };
        foreach (var (k, expPmf, expCdf) in zmnbTests)
        {
            var pmf = Distributions.ACT_DIST_ZMNEGBIN_PDF(k, 5, 0.3, 0.2);
            var cdf = Distributions.ACT_DIST_ZMNEGBIN_CDF(k, 5, 0.3, 0.2);
            Log($"| {k} | {pmf:F10} | {cdf:F10} | {expPmf:F10} |");
        }
        var zmnbMean = Distributions.ACT_DIST_ZMNEGBIN_MEAN(5, 0.3, 0.2);
        var zmnbVar = Distributions.ACT_DIST_ZMNEGBIN_VAR(5, 0.3, 0.2);
        Log($"\nMean: {zmnbMean:F10} (expected: 9.333333333)");
        Log($"Variance: {zmnbVar:F10} (expected: 52.888888889)");

        // Reconciliation table
        Log("\n### Reconciliation: Zero-Modified Distributions");
        Log("Reference: scipy.stats with zero-modification formula\n");
        Log(TableRow(("Distribution", 15), ("Test", 12), ("C#", 14), ("scipy", 14), ("Tol", 8), ("Match", 5)));
        Log(TableSep(15, 12, 14, 14, 8, 5));

        // ZM Poisson
        var zmpPdf0 = Distributions.ACT_DIST_ZMPOISSON_PDF(0, 5, 0.2);
        Log(TableRow(("ZM Poisson", 15), ("PMF(k=0)", 12), ($"{zmpPdf0:F10}", 14), ("0.2053903576", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(zmpPdf0, 0.20539035759926838, 1e-10)), 5)));

        var zmpPdf5 = Distributions.ACT_DIST_ZMPOISSON_PDF(5, 5, 0.2);
        Log(TableRow(("ZM Poisson", 15), ("PMF(k=5)", 12), ($"{zmpPdf5:F10}", 14), ("0.1403738958", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(zmpPdf5, 0.14037389581428056, 1e-10)), 5)));

        Log(TableRow(("ZM Poisson", 15), ("Mean", 12), ($"{zmpMean:F10}", 14), ("4.0000000000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(zmpMean, 4.0, 1e-10)), 5)));

        Log(TableRow(("ZM Poisson", 15), ("Variance", 12), ($"{zmpVar:F10}", 14), ("8.0000000000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(zmpVar, 8.0, 1e-10)), 5)));

        // ZM Neg Binomial
        var zmnbPdf0 = Distributions.ACT_DIST_ZMNEGBIN_PDF(0, 5, 0.3, 0.2);
        Log(TableRow(("ZM NegBin", 15), ("PMF(k=0)", 12), ($"{zmnbPdf0:F10}", 14), ("0.2019440000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(zmnbPdf0, 0.201944, 1e-10)), 5)));

        Log(TableRow(("ZM NegBin", 15), ("Mean", 12), ($"{zmnbMean:F10}", 14), ("9.3333333333", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(zmnbMean, 9.333333333333334, 1e-10)), 5)));

        Log(TableRow(("ZM NegBin", 15), ("Variance", 12), ($"{zmnbVar:F10}", 14), ("52.8888888889", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(zmnbVar, 52.88888888888889, 1e-6)), 5)));
        Log("");
    }

    static void TestParetoExtended()
    {
        Log("## Extended Pareto Family (III/IV)\n");
        Log("Generalized Pareto distributions for heavy-tailed loss modeling.\n");

        // Pareto III: μ=1 (location), σ=2 (scale), γ=3 (shape)
        Log("### Pareto III (μ=1, σ=2, γ=3)");
        Log("Lomax with location parameter. CDF: F(x) = 1 - (1 + (x-μ)/σ)^(-γ)\n");
        Log("| x | PDF | CDF | Expected PDF |");
        Log("|---|-----|-----|--------------|");
        var p3Tests = new (double x, double expPdf, double expCdf)[]
        {
            (1.0, 1.5, 0.0),
            (2.0, 0.2962962962962963, 0.7037037037037037),
            (3.0, 0.09375, 0.875),
            (5.0, 0.018518518518518517, 0.962962962962963)
        };
        foreach (var (x, expPdf, expCdf) in p3Tests)
        {
            var pdf = Distributions.ACT_DIST_PARETO3_PDF(x, 1, 2, 3);
            var cdf = Distributions.ACT_DIST_PARETO3_CDF(x, 1, 2, 3);
            Log($"| {x} | {pdf:F10} | {cdf:F10} | {expPdf:F10} |");
        }
        var p3Mean = Distributions.ACT_DIST_PARETO3_MEAN(1, 2, 3);
        Log($"\nMean: {p3Mean:F10} (expected: 2.0)");
        var p3Inv50 = Distributions.ACT_DIST_PARETO3_INV(0.5, 1, 2, 3);
        Log($"Median (p=0.5): {p3Inv50:F10} (expected: 1.5198420998)");

        // Pareto IV: μ=0, σ=2, γ=0.5, α=3
        Log("\n### Pareto IV (μ=0, σ=2, γ=0.5, α=3)");
        Log("4-parameter generalized Pareto. CDF: F(x) = 1 - (1 + ((x-μ)/σ)^(1/γ))^(-α)\n");
        Log("| x | PDF | CDF | Expected PDF |");
        Log("|---|-----|-----|--------------|");
        var p4Tests = new (double x, double expPdf, double expCdf)[]
        {
            (0.5, 0.5884987009255157, 0.16629350702218604),
            (1.0, 0.6144000000000001, 0.488),
            (2.0, 0.1875, 0.875),
            (5.0, 0.002714621204302109, 0.9973758661691746)
        };
        foreach (var (x, expPdf, expCdf) in p4Tests)
        {
            var pdf = Distributions.ACT_DIST_PARETO4_PDF(x, 0, 2, 0.5, 3);
            var cdf = Distributions.ACT_DIST_PARETO4_CDF(x, 0, 2, 0.5, 3);
            Log($"| {x} | {pdf:F10} | {cdf:F10} | {expPdf:F10} |");
        }
        var p4Inv50 = Distributions.ACT_DIST_PARETO4_INV(0.5, 0, 2, 0.5, 3);
        Log($"\nMedian (p=0.5): {p4Inv50:F10} (expected: 1.0196490571)");

        // Reconciliation table
        Log("\n### Reconciliation: Extended Pareto Family");
        Log("Reference: Analytical formulas\n");
        Log(TableRow(("Distribution", 12), ("Test", 12), ("C#", 14), ("Expected", 14), ("Tol", 8), ("Match", 5)));
        Log(TableSep(12, 12, 14, 14, 8, 5));

        // Pareto III
        var p3Pdf2 = Distributions.ACT_DIST_PARETO3_PDF(2, 1, 2, 3);
        Log(TableRow(("Pareto III", 12), ("PDF(x=2)", 12), ($"{p3Pdf2:F10}", 14), ("0.2962962963", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(p3Pdf2, 0.2962962962962963, 1e-10)), 5)));

        var p3Cdf3 = Distributions.ACT_DIST_PARETO3_CDF(3, 1, 2, 3);
        Log(TableRow(("Pareto III", 12), ("CDF(x=3)", 12), ($"{p3Cdf3:F10}", 14), ("0.8750000000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(p3Cdf3, 0.875, 1e-10)), 5)));

        Log(TableRow(("Pareto III", 12), ("Mean", 12), ($"{p3Mean:F10}", 14), ("2.0000000000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(p3Mean, 2.0, 1e-10)), 5)));

        // Pareto IV
        var p4Pdf1 = Distributions.ACT_DIST_PARETO4_PDF(1, 0, 2, 0.5, 3);
        Log(TableRow(("Pareto IV", 12), ("PDF(x=1)", 12), ($"{p4Pdf1:F10}", 14), ("0.6144000000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(p4Pdf1, 0.6144, 1e-10)), 5)));

        var p4Cdf2 = Distributions.ACT_DIST_PARETO4_CDF(2, 0, 2, 0.5, 3);
        Log(TableRow(("Pareto IV", 12), ("CDF(x=2)", 12), ($"{p4Cdf2:F10}", 14), ("0.8750000000", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(p4Cdf2, 0.875, 1e-10)), 5)));

        Log(TableRow(("Pareto IV", 12), ("INV(0.5)", 12), ($"{p4Inv50:F10}", 14), ("1.0196490571", 14), ("1e-10", 8), (FormatMatch(WithinTolerance(p4Inv50, 1.0196490570679173, 1e-10)), 5)));
        Log("");
    }

    static void TestPanjer()
    {
        Log("## Panjer Recursion (Aggregate Claims)\n");
        Log("S = X1 + X2 + ... + XN where N ~ frequency and Xi ~ severity (iid)\n");

        // Test with degenerate severity (X=1) so S = N
        Log("### Validation: Degenerate Severity (X=1 always, so S=N)");
        Log("When all claims are exactly 1, aggregate = claim count.\n");

        // Poisson(2) with degenerate severity
        var f_deg = new double[] { 0.0, 1.0 };  // f[0]=0, f[1]=1
        var g_poisson = (double[])Aggregate.ACT_PANJER_POISSON(2.0, f_deg, 10);

        Log("#### Poisson(λ=2) frequency");
        Log("| s | g_s (Panjer) | Poisson PMF | Match |");
        Log("|---|--------------|-------------|-------|");
        var poissonExpected = new double[] { 0.1353352832, 0.2706705665, 0.2706705665, 0.1804470443, 0.0902235222, 0.0360894089 };
        for (int s = 0; s <= 5; s++)
        {
            var match = WithinTolerance(g_poisson[s], poissonExpected[s], 1e-6);
            Log($"| {s} | {g_poisson[s]:F10} | {poissonExpected[s]:F10} | {FormatMatch(match)} |");
        }

        // NegBin(2, 0.5) with degenerate severity
        var g_negbin = (double[])Aggregate.ACT_PANJER_NEGBIN(2.0, 0.5, f_deg, 10);

        Log("\n#### Negative Binomial(r=2, p=0.5) frequency");
        Log("| s | g_s (Panjer) | NegBin PMF | Match |");
        Log("|---|--------------|------------|-------|");
        var negbinExpected = new double[] { 0.25, 0.25, 0.1875, 0.125, 0.078125, 0.046875 };
        for (int s = 0; s <= 5; s++)
        {
            var match = WithinTolerance(g_negbin[s], negbinExpected[s], 1e-6);
            Log($"| {s} | {g_negbin[s]:F10} | {negbinExpected[s]:F10} | {FormatMatch(match)} |");
        }

        // Test with discretized exponential severity
        Log("\n### Realistic Example: Poisson(2) + Exponential(1) Severity");
        var f_exp = (double[])Aggregate.ACT_DISCRETIZE_EXPONENTIAL(1.0, 0.5, 40);
        var g_exp = (double[])Aggregate.ACT_PANJER_POISSON(2.0, f_exp, 100);

        Log("Discretization: h=0.5, m=40 grid points\n");
        Log("| s | Aggregate x | g_s (PMF) |");
        Log("|---|-------------|-----------|");
        for (int s = 0; s <= 10; s++)
        {
            Log($"| {s} | {s * 0.5:F1} | {g_exp[s]:F10} |");
        }

        // Compute aggregate statistics
        var mean = Aggregate.ACT_AGGREGATE_MEAN(g_exp, 0.5);
        var stdev = Aggregate.ACT_AGGREGATE_STDEV(g_exp, 0.5);
        var var95 = Aggregate.ACT_AGGREGATE_VAR(0.95, g_exp, 0.5);

        Log($"\nAggregate Statistics:");
        Log($"- Mean: {mean:F4} (theoretical: 2.0)");
        Log($"- Std Dev: {stdev:F4} (theoretical: 2.0)");
        Log($"- VaR(95%): {var95:F4}");

        // Reconciliation table
        Log("\n### Reconciliation: Panjer Recursion");
        Log("Reference: scipy.stats frequency distributions\n");
        Log(TableRow(("Test", 20), ("Panjer", 14), ("Expected", 14), ("Tol", 8), ("Match", 5)));
        Log(TableSep(20, 14, 14, 8, 5));

        Log(TableRow(("Poisson g_0", 20), ($"{g_poisson[0]:F10}", 14), ("0.1353352832", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(g_poisson[0], 0.1353352832, 1e-6)), 5)));
        Log(TableRow(("Poisson g_2", 20), ($"{g_poisson[2]:F10}", 14), ("0.2706705665", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(g_poisson[2], 0.2706705665, 1e-6)), 5)));
        Log(TableRow(("NegBin g_0", 20), ($"{g_negbin[0]:F10}", 14), ("0.2500000000", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(g_negbin[0], 0.25, 1e-6)), 5)));
        Log(TableRow(("NegBin g_3", 20), ($"{g_negbin[3]:F10}", 14), ("0.1250000000", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(g_negbin[3], 0.125, 1e-6)), 5)));
        Log("");
    }

    static void TestLEV()
    {
        Log("## Limited Expected Value (LEV) Functions\n");
        Log("LEV = E[min(X, d)] - expected value capped at limit d.\n");

        // Exponential LEV: lambda=2
        Log("### Exponential LEV (λ=2)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var expLev05 = Distributions.ACT_DIST_EXP_LEV(0.5, 2);
        var expLev1 = Distributions.ACT_DIST_EXP_LEV(1.0, 2);
        var expLev2 = Distributions.ACT_DIST_EXP_LEV(2.0, 2);
        Log($"| 0.5 | {expLev05:F10} | 0.3160602794 |");
        Log($"| 1.0 | {expLev1:F10} | 0.4323323584 |");
        Log($"| 2.0 | {expLev2:F10} | 0.4908421806 |");

        // Lomax LEV: alpha=2, lambda=1
        Log("\n### Lomax (Pareto II) LEV (α=2, λ=1)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var lomaxLev1 = Distributions.ACT_DIST_LOMAX_LEV(1.0, 2, 1);
        var lomaxLev2 = Distributions.ACT_DIST_LOMAX_LEV(2.0, 2, 1);
        var lomaxLev5 = Distributions.ACT_DIST_LOMAX_LEV(5.0, 2, 1);
        Log($"| 1.0 | {lomaxLev1:F10} | 0.5000000000 |");
        Log($"| 2.0 | {lomaxLev2:F10} | 0.6666666667 |");
        Log($"| 5.0 | {lomaxLev5:F10} | 0.8333333333 |");

        // GPD LEV: xi=0.5, sigma=1
        Log("\n### GPD LEV (ξ=0.5, σ=1)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var gpdLev1 = Distributions.ACT_DIST_GPD_LEV(1.0, 0.5, 1);
        var gpdLev2 = Distributions.ACT_DIST_GPD_LEV(2.0, 0.5, 1);
        var gpdLev5 = Distributions.ACT_DIST_GPD_LEV(5.0, 0.5, 1);
        Log($"| 1.0 | {gpdLev1:F10} | 0.6666666667 |");
        Log($"| 2.0 | {gpdLev2:F10} | 1.0000000000 |");
        Log($"| 5.0 | {gpdLev5:F10} | 1.4285714286 |");

        // Gamma LEV: alpha=2, beta=1
        Log("\n### Gamma LEV (α=2, β=1)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var gammaLev1 = Distributions.ACT_DIST_GAMMA_LEV(1.0, 2, 1);
        var gammaLev2 = Distributions.ACT_DIST_GAMMA_LEV(2.0, 2, 1);
        var gammaLev5 = Distributions.ACT_DIST_GAMMA_LEV(5.0, 2, 1);
        Log($"| 1.0 | {gammaLev1:F10} | 0.8963616765 |");
        Log($"| 2.0 | {gammaLev2:F10} | 1.4586588671 |");
        Log($"| 5.0 | {gammaLev5:F10} | 1.9528343710 |");

        // Lognormal LEV: mu=0, sigma=1
        Log("\n### Lognormal LEV (μ=0, σ=1)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var lnLev1 = Distributions.ACT_DIST_LOGNORM_LEV(1.0, 0, 1);
        var lnLev2 = Distributions.ACT_DIST_LOGNORM_LEV(2.0, 0, 1);
        var lnLev5 = Distributions.ACT_DIST_LOGNORM_LEV(5.0, 0, 1);
        Log($"| 1.0 | {lnLev1:F10} | 0.7615782919 |");
        Log($"| 2.0 | {lnLev2:F10} | 1.1138701492 |");
        Log($"| 5.0 | {lnLev5:F10} | 1.4705262812 |");

        // Weibull LEV: k=2, lambda=1
        Log("\n### Weibull LEV (k=2, λ=1)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var weibullLev05 = Distributions.ACT_DIST_WEIBULL_LEV(0.5, 2, 1);
        var weibullLev1 = Distributions.ACT_DIST_WEIBULL_LEV(1.0, 2, 1);
        var weibullLev2 = Distributions.ACT_DIST_WEIBULL_LEV(2.0, 2, 1);
        Log($"| 0.5 | {weibullLev05:F10} | 0.4612810064 |");
        Log($"| 1.0 | {weibullLev1:F10} | 0.7468241328 |");
        Log($"| 2.0 | {weibullLev2:F10} | 0.8820813908 |");

        // Beta LEV: alpha=2, beta=5
        Log("\n### Beta LEV (α=2, β=5)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var betaLev025 = Distributions.ACT_DIST_BETA_LEV(0.25, 2, 5);
        var betaLev05 = Distributions.ACT_DIST_BETA_LEV(0.5, 2, 5);
        var betaLev1 = Distributions.ACT_DIST_BETA_LEV(1.0, 2, 5);
        Log($"| 0.25 | {betaLev025:F10} | 0.2030814035 |");
        Log($"| 0.50 | {betaLev05:F10} | 0.2756696429 |");
        Log($"| 1.00 | {betaLev1:F10} | 0.2857142857 |");

        // Burr XII LEV: c=2, k=3, lambda=1
        Log("\n### Burr XII LEV (c=2, k=3, λ=1)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var burrLev05 = Distributions.ACT_DIST_BURR_LEV(0.5, 2, 3, 1);
        var burrLev1 = Distributions.ACT_DIST_BURR_LEV(1.0, 2, 3, 1);
        var burrLev2 = Distributions.ACT_DIST_BURR_LEV(2.0, 2, 3, 1);
        Log($"| 0.5 | {burrLev05:F10} | 0.4038678534 |");
        Log($"| 1.0 | {burrLev1:F10} | 0.5445243113 |");
        Log($"| 2.0 | {burrLev2:F10} | 0.5851807692 |");

        // Poisson LEV: lambda=5
        Log("\n### Poisson LEV (λ=5)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var poissonLev3 = Distributions.ACT_DIST_POISSON_LEV(3.0, 5);
        var poissonLev5 = Distributions.ACT_DIST_POISSON_LEV(5.0, 5);
        var poissonLev10 = Distributions.ACT_DIST_POISSON_LEV(10.0, 5);
        Log($"| 3.0 | {poissonLev3:F10} | 2.8281823515 |");
        Log($"| 5.0 | {poissonLev5:F10} | 4.1226631512 |");
        Log($"| 10.0 | {poissonLev10:F10} | 4.9778123995 |");

        // Negative Binomial LEV: r=5, p=0.3
        Log("\n### Negative Binomial LEV (r=5, p=0.3)");
        Log("| Limit | LEV | scipy Expected |");
        Log("|-------|-----|----------------|");
        var nbLev5 = Distributions.ACT_DIST_NEGBIN_LEV(5.0, 5, 0.3);
        var nbLev10 = Distributions.ACT_DIST_NEGBIN_LEV(10.0, 5, 0.3);
        var nbLev20 = Distributions.ACT_DIST_NEGBIN_LEV(20.0, 5, 0.3);
        Log($"| 5.0 | {nbLev5:F10} | 4.8010631900 |");
        Log($"| 10.0 | {nbLev10:F10} | 8.4026604566 |");
        Log($"| 20.0 | {nbLev20:F10} | 11.2187398801 |");

        // Reconciliation table
        Log("\n### Reconciliation: LEV Functions");
        Log("Reference: scipy.stats numerical integration of survival function\n");
        Log(TableRow(("Distribution", 15), ("Limit", 6), ("C#", 14), ("scipy", 14), ("Tol", 8), ("Match", 5)));
        Log(TableSep(15, 6, 14, 14, 8, 5));

        Log(TableRow(("Exponential", 15), ("1.0", 6), ($"{expLev1:F10}", 14), ("0.4323323584", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(expLev1, 0.4323323584, 1e-6)), 5)));
        Log(TableRow(("Lomax", 15), ("2.0", 6), ($"{lomaxLev2:F10}", 14), ("0.6666666667", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(lomaxLev2, 0.6666666667, 1e-6)), 5)));
        Log(TableRow(("GPD", 15), ("2.0", 6), ($"{gpdLev2:F10}", 14), ("1.0000000000", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(gpdLev2, 1.0, 1e-6)), 5)));
        Log(TableRow(("Gamma", 15), ("2.0", 6), ($"{gammaLev2:F10}", 14), ("1.4586588671", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(gammaLev2, 1.4586588671, 1e-6)), 5)));
        Log(TableRow(("Lognormal", 15), ("2.0", 6), ($"{lnLev2:F10}", 14), ("1.1138701492", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(lnLev2, 1.1138701492, 1e-6)), 5)));
        Log(TableRow(("Weibull", 15), ("1.0", 6), ($"{weibullLev1:F10}", 14), ("0.7468241328", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(weibullLev1, 0.7468241328, 1e-6)), 5)));
        Log(TableRow(("Beta", 15), ("0.5", 6), ($"{betaLev05:F10}", 14), ("0.2756696429", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(betaLev05, 0.2756696429, 1e-6)), 5)));
        Log(TableRow(("Burr XII", 15), ("1.0", 6), ($"{burrLev1:F10}", 14), ("0.5445243113", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(burrLev1, 0.5445243113, 1e-6)), 5)));
        Log(TableRow(("Poisson", 15), ("5.0", 6), ($"{poissonLev5:F10}", 14), ("4.1226631512", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(poissonLev5, 4.1226631512, 1e-6)), 5)));
        Log(TableRow(("Neg Binomial", 15), ("10.0", 6), ($"{nbLev10:F10}", 14), ("8.4026604566", 14), ("1e-6", 8), (FormatMatch(WithinTolerance(nbLev10, 8.4026604566, 1e-6)), 5)));
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
        Log(TableRow(("Test", 25), ("C#", 10), ("Reference", 10), ("Source", 24), ("Tol", 8), ("Match", 5)));
        Log(TableSep(25, 10, 10, 24, 8, 5));

        // MBBEFD at d=0.5 with b=2, g=3
        var mbbefd05 = ExposureCurves.ACT_EXPOSURE_MBBEFD(0.5, 2, 3);
        Log(TableRow(("MBBEFD(d=0.5, b=2, g=3)", 25), ($"{mbbefd05:F6}", 10), ("0.5960", 10), ("Bernegger formula", 24), ("0.01", 8), (FormatMatch(WithinTolerance(mbbefd05, 0.596, 0.01)), 5)));

        // Power curve: G(d) = d^n
        Log(TableRow(("Power(d=0.5, n=2)", 25), ($"{powerN2:F6}", 10), ("0.2500", 10), ("G(d)=d^n", 24), ("0.0001", 8), (FormatMatch(WithinTolerance(powerN2, 0.25, 0.0001)), 5)));

        // Inverse power: G(d) = 1 - (1-d)^n
        Log(TableRow(("InvPower(d=0.5, n=2)", 25), ($"{invPowerN2:F6}", 10), ("0.7500", 10), ("G(d)=1-(1-d)^n", 24), ("0.0001", 8), (FormatMatch(WithinTolerance(invPowerN2, 0.75, 0.0001)), 5)));

        // Pareto exposure: G(d) = 1 - (1-d)^((α-1)/α)
        Log(TableRow(("Pareto(d=0.5, α=2)", 25), ($"{paretoExp:F6}", 10), ("0.2929", 10), ("G(d)=1-(1-d)^((α-1)/α)", 24), ("0.001", 8), (FormatMatch(WithinTolerance(paretoExp, 0.2929, 0.001)), 5)));
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

    static void TestCatModeling()
    {
        Log("## 5. Cat Modeling (ELT → YLT → EP Curves)\n");

        var eventRates = new double[] { 0.2, 0.1, 0.05 };
        var eventLosses = new double[] { 1_000_000, 2_500_000, 10_000_000 };
        int years = 1000;
        int seed = 42;

        var ylt = CatModeling.ACT_CAT_ELT_TO_YLT(eventRates, eventLosses, years, seed, true);

        Log("### Event Loss Table (ELT)");
        Log("| Event | Rate | Loss |");
        Log("|-------|------|------|");
        for (int i = 0; i < eventRates.Length; i++)
        {
            Log($"| {i + 1} | {eventRates[i]:0.00} | {eventLosses[i]:N0} |");
        }
        Log("");

        Log("### Year Loss Table (YLT) - First 5 Years");
        Log("| Year | Aggregate Loss | Max Loss | Event Count |");
        Log("|------|----------------|----------|-------------|");
        for (int i = 1; i <= 5; i++)
        {
            Log($"| {ylt[i, 0]} | {Convert.ToDouble(ylt[i, 1]):N0} | {Convert.ToDouble(ylt[i, 2]):N0} | {ylt[i, 3]} |");
        }
        Log("");

        var aggregateLosses = ExtractColumn(ylt, 1, true);
        var maxLosses = ExtractColumn(ylt, 2, true);

        var oep = CatModeling.ACT_CAT_YLT_OEP_CURVE(maxLosses, "WEIBULL", true);
        var aep = CatModeling.ACT_CAT_YLT_AEP_CURVE(aggregateLosses, "WEIBULL", true);

        Log("### OEP Curve (First 5 Return Periods)");
        Log("| Return Period | OEP Loss |");
        Log("|---------------|----------|");
        for (int i = 1; i <= 5; i++)
        {
            Log($"| {Convert.ToDouble(oep[i, 0]):0.00} | {Convert.ToDouble(oep[i, 1]):N0} |");
        }
        Log("");

        Log("### AEP Curve (First 5 Return Periods)");
        Log("| Return Period | AEP Loss |");
        Log("|---------------|----------|");
        for (int i = 1; i <= 5; i++)
        {
            Log($"| {Convert.ToDouble(aep[i, 0]):0.00} | {Convert.ToDouble(aep[i, 1]):N0} |");
        }
        Log("");
    }

    static void TestInterpolation()
    {
        Log("## 6. Interpolation\n");

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
        Log("## 7. Chain Ladder (Taylor-Ashe Dataset)\n");
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

        Log(TableRow(("Test", 15), ("C#", 12), ("Reference", 12), ("Source", 32), ("Tol", 6), ("Match", 5)));
        Log(TableSep(15, 12, 12, 32, 6, 5));

        // Development factors
        for (int i = 0; i < Math.Min(3, factors.Length); i++)
        {
            Log(TableRow(($"Factor {i+1}-{i+2}", 15), ($"{factors[i]:F4}", 12), ($"{expectedFactors[i]:F3}", 12), ("E&V 2002 / chainladder-python", 32), ("0.01", 6), (FormatMatch(WithinTolerance(factors[i], expectedFactors[i], 0.01)), 5)));
        }

        // Total IBNR
        Log(TableRow(("Total IBNR", 15), ($"{totalIBNR:N0}", 12), ("18,680,856", 12), ("E&V 2002 Table 3", 32), ("1000", 6), (FormatMatch(WithinTolerance(totalIBNR, 18680856, 1000)), 5)));

        // Total Mack SE
        Log(TableRow(("Total Mack SE", 15), ($"{totalSE:N0}", 12), ("2,447,095", 12), ("E&V 2002 (analytic)", 32), ("1000", 6), (FormatMatch(WithinTolerance(totalSE, 2447095, 1000)), 5)));

        // Individual reserve SEs
        Log(TableRow(("Mack SE (AY2)", 15), ($"{reserveSE[1]:N0}", 12), ("75,535", 12), ("E&V 2002", 32), ("100", 6), (FormatMatch(WithinTolerance(reserveSE[1], 75535, 100)), 5)));
        Log(TableRow(("Mack SE (AY5)", 15), ($"{reserveSE[4]:N0}", 12), ("261,406", 12), ("E&V 2002", 32), ("100", 6), (FormatMatch(WithinTolerance(reserveSE[4], 261406, 100)), 5)));
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
        Log("## 8. Bootstrap Chain Ladder\n");
        Log("Using Taylor-Ashe data with ODP bootstrap (England & Verrall, 2002 method).\n");

        var triangle = GetTaylorAsheTriangle();

        TestStochasticReservingGoldenPaths(triangle);

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

        // --- EV Method (default) ---
        Log("### EV Method Bootstrap (10,000 iterations, seed=123)");
        Log("Reference: StochasticReserving Main_ODP_Bstrap with non-constant scale,");
        Log("non-parametric pseudo data, Gamma forecast, and NumPy-compatible RNG.\n");

        var originStats = (object[,])ChainLadder.ACT_CL_BOOTSTRAP_ORIGIN(triangle, 10000, 123, "EV");

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

        var bootstrap = (object[,])ChainLadder.ACT_CL_BOOTSTRAP(triangle, 10000, 123, "EV");
        double totalSE = Convert.ToDouble(bootstrap[1, 1]);
        Log($"\nTotal EV Bootstrap: Mean={bootstrap[0, 1]:N0}, StdDev={totalSE:N0}");
        Log($"E&V Non-Constant Total SE: 2,228,677 (Ratio: {totalSE / 2228677:P0})");

        // --- BASIC Method ---
        Log("\n### CHAINLADDER-PYTHON Method Bootstrap (10,000 iterations, seed=123)");
        Log("chainladder-python compatible: constant phi, no hat, original diagonal.\n");

        var basicBootstrap = (object[,])ChainLadder.ACT_CL_BOOTSTRAP(triangle, 10000, 123, "CHAINLADDER-PYTHON");
        double basicSE = Convert.ToDouble(basicBootstrap[1, 1]);
        double basicMean = Convert.ToDouble(basicBootstrap[0, 1]);
        double basicCV = basicSE / basicMean;
        Log($"Total BASIC Bootstrap: Mean={basicMean:N0}, StdDev={basicSE:N0}, CV={basicCV:P1}");

        // Bootstrap Reconciliation
        Log("\n### Reconciliation: Bootstrap");
        Log("Reference: England & Verrall (2002) Table 3 - ODP Bootstrap with non-constant scale\n");
        Log(TableRow(("Test", 16), ("C#", 12), ("Reference", 12), ("Ratio", 8), ("Tol", 6), ("Match", 5)));
        Log(TableSep(16, 12, 12, 8, 6, 5));
        Log(TableRow(("EV Total SE", 16), ($"{totalSE:N0}", 12), ("2,228,677", 12), ($"{totalSE / 2228677:P0}", 8), ("5%", 6), (FormatMatch(Math.Abs(totalSE / 2228677 - 1.0) <= 0.05), 5)));

        // Check a few individual origin years
        double se6 = Convert.ToDouble(originStats[6, 2]);
        double se10 = Convert.ToDouble(originStats[10, 2]);
        Log(TableRow(("EV AY6 SE", 16), ($"{se6:N0}", 12), ("398,377", 12), ($"{se6 / 398377:P0}", 8), ("10%", 6), (FormatMatch(Math.Abs(se6 / 398377 - 1.0) <= 0.10), 5)));
        Log(TableRow(("EV AY10 SE", 16), ($"{se10:N0}", 12), ("1,285,560", 12), ($"{se10 / 1285560:P0}", 8), ("10%", 6), (FormatMatch(Math.Abs(se10 / 1285560 - 1.0) <= 0.10), 5)));
        Log(TableRow(("BASIC Total CV", 16), ($"{basicCV:P1}", 12), ("13-17%", 12), ("", 8), ("", 6), (FormatMatch(basicCV >= 0.10 && basicCV <= 0.20), 5)));
        Log("");
    }

    static void TestStochasticReservingGoldenPaths(double[,] triangle)
    {
        Log("### StochasticReserving pathwise golden reconciliation");
        string fixturePath = Path.Combine(_fixturesPath, "stochastic_reserving_golden.json");
        using var document = JsonDocument.Parse(File.ReadAllText(fixturePath));
        JsonElement root = document.RootElement;
        int iterations = root.GetProperty("iterations").GetInt32();
        int seed = root.GetProperty("seed").GetInt32();
        string commit = root.GetProperty("golden_commit").GetString()!;
        Log($"Golden commit: `{commit}`; iterations={iterations}; seed={seed}\n");

        foreach (JsonElement testCase in root.GetProperty("cases").EnumerateArray())
        {
            string scale = testCase.GetProperty("scale").GetString()!;
            string bootstrapDistribution = testCase.GetProperty("bootstrap_distribution").GetString()!;
            string forecastDistribution = testCase.GetProperty("forecast_distribution").GetString()!;
            object? mask = ParseOptionalMatrix(testCase.GetProperty("mask"));
            object? userSqrtScale = ParseOptionalVector(testCase.GetProperty("user_sqrt_scale"));
            string optionSuffix = mask is not null ? "/MASK" : userSqrtScale is not null ? "/USER-SCALE" : "";
            string label = $"{scale}/{bootstrapDistribution}/{forecastDistribution}{optionSuffix}";

            var reserveOutput = ChainLadder.ACT_CL_BOOTSTRAP_SAMPLES(
                triangle, iterations, seed, "EV", scale,
                bootstrapDistribution, forecastDistribution, "RESERVES", mask, userSqrtScale);
            RecordGoldenDiff($"{label} TotalReserve", MaxVectorDiff(
                reserveOutput, 1, testCase.GetProperty("total_reserves")));
            RecordGoldenDiff($"{label} Reserves", MaxMatrixDiff(
                reserveOutput, 2, testCase.GetProperty("reserves")));

            var ultimateOutput = ChainLadder.ACT_CL_BOOTSTRAP_SAMPLES(
                triangle, iterations, seed, "EV", scale,
                bootstrapDistribution, forecastDistribution, "ULTIMATES", mask, userSqrtScale);
            RecordGoldenDiff($"{label} Ultimates", MaxMatrixDiff(
                ultimateOutput, 1, testCase.GetProperty("ultimates")));

            var linkRatioOutput = ChainLadder.ACT_CL_BOOTSTRAP_SAMPLES(
                triangle, iterations, seed, "EV", scale,
                bootstrapDistribution, forecastDistribution, "PSEUDO-LRS", mask, userSqrtScale);
            RecordGoldenDiff($"{label} Pseudo_LRs", MaxMatrixDiff(
                linkRatioOutput, 1, testCase.GetProperty("pseudo_link_ratios")));

            var cumulativeOutput = ChainLadder.ACT_CL_BOOTSTRAP_SAMPLES(
                triangle, iterations, seed, "EV", scale,
                bootstrapDistribution, forecastDistribution, "CUMULATIVES", mask, userSqrtScale);
            RecordGoldenDiff($"{label} Cumulatives", MaxCubeDiff(
                cumulativeOutput, testCase.GetProperty("cumulatives")));

            var completeOutput = ChainLadder.ACT_CL_BOOTSTRAP_SAMPLES(
                triangle, iterations, seed, "EV", scale,
                bootstrapDistribution, forecastDistribution, "COMPLETE-CUMULATIVES", mask, userSqrtScale);
            RecordGoldenDiff($"{label} Complete_Cumulatives", MaxCubeDiff(
                completeOutput, testCase.GetProperty("complete_cumulatives")));
        }

        const int reconciliationIterations = 64;
        var samples = ChainLadder.ACT_CL_BOOTSTRAP_SAMPLES(
            triangle, reconciliationIterations, 123, "EV");
        var totals = ExtractColumn(samples, 1, hasHeader: true);
        var totalSummary = ChainLadder.ACT_CL_BOOTSTRAP(
            triangle, reconciliationIterations, 123, "EV");
        double totalMean = totals.Average();
        double totalStdDev = Math.Sqrt(totals.Select(x => Math.Pow(x - totalMean, 2)).Average());
        RecordGoldenDiff("Samples reconcile: total mean",
            Math.Abs(totalMean - Convert.ToDouble(totalSummary[0, 1])));
        RecordGoldenDiff("Samples reconcile: total stddev",
            Math.Abs(totalStdDev - Convert.ToDouble(totalSummary[1, 1])));

        var originSummary = ChainLadder.ACT_CL_BOOTSTRAP_ORIGIN(
            triangle, reconciliationIterations, 123, "EV");
        double maxOriginDiff = 0.0;
        for (int origin = 0; origin < triangle.GetLength(0); origin++)
        {
            var originSamples = ExtractColumn(samples, origin + 2, hasHeader: true);
            double mean = originSamples.Average();
            double stdDev = Math.Sqrt(originSamples.Select(x => Math.Pow(x - mean, 2)).Average());
            maxOriginDiff = Math.Max(maxOriginDiff, Math.Abs(mean - Convert.ToDouble(originSummary[origin + 1, 1])));
            maxOriginDiff = Math.Max(maxOriginDiff, Math.Abs(stdDev - Convert.ToDouble(originSummary[origin + 1, 2])));
        }
        RecordGoldenDiff("Samples reconcile: origin summaries", maxOriginDiff);

        var invalidOutput = ChainLadder.ACT_CL_BOOTSTRAP_SAMPLES(
            triangle, 1, 42, "EV", "NONCONSTANT", "INVALID", "GAMMA");
        bool invalidHandled = invalidOutput.GetLength(0) == 1
            && Convert.ToString(invalidOutput[0, 0])!.StartsWith("Error:", StringComparison.Ordinal);
        Log($"Invalid option contract: {FormatMatch(invalidHandled)}");
        Log("");
    }

    static void RecordGoldenDiff(string label, double maxDifference)
    {
        Log($"{label}: max abs diff={maxDifference:G6}, match={FormatMatch(maxDifference == 0.0)}");
    }

    static double MaxVectorDiff(object[,] actual, int actualColumn, JsonElement expected)
    {
        double maximum = 0.0;
        int iteration = 0;
        foreach (JsonElement value in expected.EnumerateArray())
        {
            maximum = Math.Max(maximum,
                Math.Abs(Convert.ToDouble(actual[iteration + 1, actualColumn]) - value.GetDouble()));
            iteration++;
        }
        return maximum;
    }

    static double[,]? ParseOptionalMatrix(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Null)
            return null;
        int rows = element.GetArrayLength();
        int columns = element[0].GetArrayLength();
        var result = new double[rows, columns];
        int i = 0;
        foreach (JsonElement row in element.EnumerateArray())
        {
            int j = 0;
            foreach (JsonElement value in row.EnumerateArray())
                result[i, j++] = value.GetDouble();
            i++;
        }
        return result;
    }

    static double[]? ParseOptionalVector(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Null)
            return null;
        var result = new double[element.GetArrayLength()];
        int i = 0;
        foreach (JsonElement value in element.EnumerateArray())
            result[i++] = value.GetDouble();
        return result;
    }

    static double MaxMatrixDiff(object[,] actual, int firstActualColumn, JsonElement expected)
    {
        double maximum = 0.0;
        int iteration = 0;
        foreach (JsonElement row in expected.EnumerateArray())
        {
            int column = 0;
            foreach (JsonElement value in row.EnumerateArray())
            {
                maximum = Math.Max(maximum,
                    Math.Abs(Convert.ToDouble(actual[iteration + 1, firstActualColumn + column]) - value.GetDouble()));
                column++;
            }
            iteration++;
        }
        return maximum;
    }

    static double MaxCubeDiff(object[,] actual, JsonElement expected)
    {
        double maximum = 0.0;
        int iteration = 0;
        foreach (JsonElement matrix in expected.EnumerateArray())
        {
            int flattenedColumn = 1;
            foreach (JsonElement row in matrix.EnumerateArray())
            foreach (JsonElement value in row.EnumerateArray())
            {
                maximum = Math.Max(maximum,
                    Math.Abs(Convert.ToDouble(actual[iteration + 1, flattenedColumn]) - value.GetDouble()));
                flattenedColumn++;
            }
            iteration++;
        }
        return maximum;
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
            Log(TableRow(("Test", 22), ("Target", 8), ("Empirical", 10), ("Tol", 6), ("Match", 5)));
            Log(TableSep(22, 8, 10, 6, 5));

            // Check that empirical correlations match targets (with sampling error tolerance)
            foreach (var (a, b) in new (int, int)[] { (0, 1), (1, 3) })
            {
                double target = corrMatrix[a, b];
                double empirical = CalculateCorrelation(data, a, b);
                bool match = WithinTolerance(empirical, target, 0.05);
                Log(TableRow(($"Corr(X{a+1}, X{b+1})", 22), ($"{target:F2}", 8), ($"{empirical:F2}", 10), ("0.05", 6), (FormatMatch(match), 5)));
            }

            // Check that values are in [0,1] (uniform marginals)
            bool allInRange = true;
            for (int i = 0; i < 1000 && allInRange; i++)
                for (int j = 0; j < dim && allInRange; j++)
                    if (data[i, j] < 0 || data[i, j] > 1) allInRange = false;
            Log(TableRow(("Uniform [0,1] range", 22), ("TRUE", 8), ($"{allInRange}", 10), ("-", 6), (FormatMatch(allInRange), 5)));
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
