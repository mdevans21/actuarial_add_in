using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Exposure curve functions for property and casualty insurance pricing.
/// An exposure curve G(d) gives the proportion of expected losses below d proportion of sum insured.
/// References:
/// - Bernegger, S. (1997). "The Swiss Re Exposure Curves and the MBBEFD Distribution Class." ASTIN Bulletin 27(1): 99-111.
/// - Riegel, U. (2008). "Generalizations of common ILF models." Blätter der DGVFM 29: 45-71.
/// </summary>
public static class ExposureCurves
{
    #region MBBEFD / Swiss Re Curves

    [ExcelFunction(Description = "MBBEFD exposure curve (Maxwell-Boltzmann-Bose-Einstein-Fermi-Dirac): G(d) gives proportion of losses below d% of sum insured. See Bernegger (1997).", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_MBBEFD(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1). d=0.5 means 50% of TSI.")] double d,
        [ExcelArgument(Description = "b parameter (b > 0). Higher b = more concentrated exposure.")] double b,
        [ExcelArgument(Description = "g parameter (g > 0). Higher g = heavier distribution tail.")] double g)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (b <= 0 || g <= 0) return double.NaN;

        if (d == 0) return 0;
        if (d == 1) return 1;

        // Handle special cases
        if (Math.Abs(b - 1) < 1e-10)
        {
            // When b approaches 1, use limit formula
            return (Math.Log(1 + (g - 1) * d)) / Math.Log(g);
        }

        if (Math.Abs(g - 1) < 1e-10)
        {
            // When g approaches 1, linear case
            return d;
        }

        // General MBBEFD formula: G(d) = ln(a + b^d) - ln(a + 1) / ln(a + b) - ln(a + 1)
        // where a = (g - b^g) / (b^g - 1) when g != 1 and b != 1
        double bg = Math.Pow(b, g);
        double a = (g - bg) / (bg - 1);

        double numerator = Math.Log(a + Math.Pow(b, d)) - Math.Log(a + 1);
        double denominator = Math.Log(a + b) - Math.Log(a + 1);

        return numerator / denominator;
    }

    [ExcelFunction(Description = "Swiss Re standard exposure curves Y1-Y4 (curve 1-4) and the Lloyd's industrial-risks curve (curve 5). Implements Bernegger (1997) Table 1: c = 1.5, 2, 3, 4, 5 with b = exp(3.1 - 0.15·c·(1+c)), g = exp(c·(0.78 + 0.12·c)).", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_SWISSRE(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Curve number 1-5: 1=Y1 (light/office), 2=Y2, 3=Y3, 4=Y4 (petrochemical), 5=Lloyd's industrial. Maps to Bernegger c=1.5,2,3,4,5.")] int curveNumber)
    {
        // Bernegger (1997) ASTIN Bulletin 27(1): the Swiss Re Y1-Y4 curves are
        // MBBEFD curves at c = 1.5, 2, 3, 4 respectively, with parameters
        //     b = exp(3.1 - 0.15·c·(1+c))
        //     g = exp(c·(0.78 + 0.12·c))
        // c = 5 coincides with the Lloyd's curve for industrial risks.
        double c = curveNumber switch
        {
            1 => 1.5,
            2 => 2.0,
            3 => 3.0,
            4 => 4.0,
            5 => 5.0,
            _ => double.NaN,
        };
        if (double.IsNaN(c)) return double.NaN;

        double b = Math.Exp(3.1 - 0.15 * c * (1 + c));
        double g = Math.Exp(c * (0.78 + 0.12 * c));
        return ACT_EXPOSURE_MBBEFD(d, b, g);
    }

    #endregion

    #region Lloyd's Curves

    [ExcelFunction(Description = "Lloyd's Y exposure curves: G(d) = 1 - (1-d)^c. Standard market curves for property insurance by occupancy type.", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_LLOYDS(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Curve code: 'Y1'=Light commercial (c=1.5), 'Y2'=Medium industrial (c=2), 'Y3'=Heavy industrial (c=3), 'Y4'=Petrochemical (c=4). Also accepts 1-4.")] string curveCode)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (d == 0) return 0;
        if (d == 1) return 1;

        // Lloyd's Y curves (standard market curves)
        // Y1: Light industrial/commercial - lower spread of loss, office buildings
        // Y2: Medium industrial - warehouses, light manufacturing
        // Y3: Heavy industrial - factories, processing plants
        // Y4: Petrochemical/refinery - highly concentrated loss potential

        double c;
        switch (curveCode.ToUpper())
        {
            case "Y1":
            case "1":
                c = 1.5;  // Light - office, retail, low fire load
                break;
            case "Y2":
            case "2":
                c = 2.0;  // Medium - warehouses, light manufacturing
                break;
            case "Y3":
            case "3":
                c = 3.0;  // Heavy - factories, heavier manufacturing
                break;
            case "Y4":
            case "4":
                c = 4.0;  // Very heavy - petrochemical, chemical plants
                break;
            default:
                return double.NaN;
        }

        // Lloyd's curve formula: G(d) = 1 - (1-d)^c
        return 1 - Math.Pow(1 - d, c);
    }

    #endregion

    #region Power Curves

    [ExcelFunction(Description = "Power exposure curve: G(d) = d^n", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_POWER(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Power exponent n (> 0)")] double n)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (n <= 0) return double.NaN;

        return Math.Pow(d, n);
    }

    [ExcelFunction(Description = "Inverse power exposure curve: G(d) = 1 - (1-d)^n", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_INVERSE_POWER(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Power exponent n (> 0)")] double n)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (n <= 0) return double.NaN;

        return 1 - Math.Pow(1 - d, n);
    }

    #endregion

    #region Pareto Exposure Curves

    [ExcelFunction(Description = "Pareto exposure curve for excess of loss pricing", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_PARETO(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Pareto alpha parameter (> 0)")] double alpha)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (alpha <= 0) return double.NaN;
        if (d == 0) return 0;
        if (d == 1) return 1;

        // Pareto exposure curve: G(d) = 1 - (1-d)^(1-1/alpha) for alpha > 1
        // For alpha <= 1, mean is infinite, curve approaches linear
        if (alpha <= 1)
        {
            return d;
        }

        double exponent = 1 - 1 / alpha;
        return 1 - Math.Pow(1 - d, exponent);
    }

    #endregion

    #region Riebesell Curve

    [ExcelFunction(Description = "Riebesell exposure curve: G(d) = d^n + (1-n)*d*(1-d^n)/(1-d) for d<1. Alternative to MBBEFD with single shape parameter.", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_RIEBESELL(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "n - shape parameter (0 < n <= 1). n=1 gives linear curve; smaller n = heavier concentration of losses.")] double n)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (n <= 0 || n > 1) return double.NaN;
        if (d == 0) return 0;
        if (d == 1) return 1;

        // Handle n = 1 (linear case)
        if (Math.Abs(n - 1) < 1e-10)
        {
            return d;
        }

        // Riebesell formula: G(d) = d^n + (1-n)*d*(1-d^n)/(1-d)
        double dn = Math.Pow(d, n);
        return dn + (1 - n) * d * (1 - dn) / (1 - d);
    }

    [ExcelFunction(Description = "Inverse Riebesell curve - find d given G(d) value using Newton-Raphson iteration.", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_RIEBESELL_INV(
        [ExcelArgument(Description = "Target G(d) value (0 to 1)")] double g,
        [ExcelArgument(Description = "n - shape parameter (0 < n <= 1)")] double n,
        [ExcelArgument(Description = "Tolerance for convergence (default 1e-8)")] double tolerance = 1e-8,
        [ExcelArgument(Description = "Maximum iterations (default 100)")] int maxIterations = 100)
    {
        if (g < 0 || g > 1) return double.NaN;
        if (n <= 0 || n > 1) return double.NaN;
        if (g == 0) return 0;
        if (g == 1) return 1;

        // Handle n = 1 (linear case)
        if (Math.Abs(n - 1) < 1e-10)
        {
            return g;
        }

        // Newton-Raphson to solve G(d) = g
        double d = g;  // Initial guess
        for (int i = 0; i < maxIterations; i++)
        {
            double currentG = ACT_EXPOSURE_RIEBESELL(d, n);
            double error = currentG - g;
            if (Math.Abs(error) < tolerance)
                return d;

            // Numerical derivative
            double h = 1e-8;
            double derivative = (ACT_EXPOSURE_RIEBESELL(d + h, n) - currentG) / h;
            if (Math.Abs(derivative) < 1e-15)
                break;

            d = d - error / derivative;
            d = Math.Max(0, Math.Min(1, d));  // Keep in bounds
        }

        return d;
    }

    #endregion

    #region Layer Rating using Exposure Curves

    [ExcelFunction(Description = "Layer rate on line: ground-up burning cost × (G(exhaust) − G(attach)). The exposure curve G(d) is already the cumulative fraction of expected loss below d, so the layer's share of expected loss is the difference of curve values directly. Bernegger (1997).", Category = "Actuarial.ExposureCurves")]
    public static double ACT_EXPOSURE_LAYER_RATE(
        [ExcelArgument(Description = "Attachment point as proportion of sum insured")] double attachmentPct,
        [ExcelArgument(Description = "Exhaustion point as proportion of sum insured")] double exhaustionPct,
        [ExcelArgument(Description = "Ground-up burning cost (rate on line)")] double burningCost,
        [ExcelArgument(Description = "Exposure curve b parameter")] double b,
        [ExcelArgument(Description = "Exposure curve g parameter")] double g)
    {
        if (attachmentPct < 0 || exhaustionPct > 1 || attachmentPct >= exhaustionPct)
            return double.NaN;

        double gAttach = ACT_EXPOSURE_MBBEFD(attachmentPct, b, g);
        double gExhaust = ACT_EXPOSURE_MBBEFD(exhaustionPct, b, g);

        // Layer share of expected loss = G(exhaust) − G(attach).
        // (The previous division by the physical layer width was incorrect.)
        return burningCost * (gExhaust - gAttach);
    }

    #endregion
}
