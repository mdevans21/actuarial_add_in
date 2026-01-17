using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

public static class ExposureCurves
{
    #region MBBEFD / Swiss Re Curves

    [ExcelFunction(Description = "MBBEFD exposure curve - returns proportion of losses captured at given proportion of sum insured", Category = "Actuarial.ExposureCurves")]
    public static double ACT_MBBEFD(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "b parameter (b > 0, b != 1)")] double b,
        [ExcelArgument(Description = "g parameter (g > 0, g != 1)")] double g)
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

    [ExcelFunction(Description = "Swiss Re exposure curve by curve number (1-5)", Category = "Actuarial.ExposureCurves")]
    public static double ACT_SWISSRE_CURVE(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Curve number (1 to 5)")] int curveNumber)
    {
        // Swiss Re standard curves parameterized as MBBEFD
        // These are commonly used b,g parameter combinations
        double b, g;
        switch (curveNumber)
        {
            case 1: b = 1.5; g = 1.5; break;   // Light
            case 2: b = 2.0; g = 2.0; break;   // Medium-Light
            case 3: b = 3.0; g = 3.0; break;   // Medium
            case 4: b = 4.0; g = 4.0; break;   // Medium-Heavy
            case 5: b = 5.0; g = 5.0; break;   // Heavy
            default: return double.NaN;
        }
        return ACT_MBBEFD(d, b, g);
    }

    #endregion

    #region Lloyd's Curves

    [ExcelFunction(Description = "Lloyd's exposure curve", Category = "Actuarial.ExposureCurves")]
    public static double ACT_LLOYDS_CURVE(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Curve code: 'Y1', 'Y2', 'Y3', 'Y4', or curve number 1-4")] string curveCode)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (d == 0) return 0;
        if (d == 1) return 1;

        // Lloyd's Y curves (standard market curves)
        // Y1: Light industrial/commercial
        // Y2: Medium industrial
        // Y3: Heavy industrial
        // Y4: Petrochemical/refinery

        double c;
        switch (curveCode.ToUpper())
        {
            case "Y1":
            case "1":
                c = 1.5;
                break;
            case "Y2":
            case "2":
                c = 2.0;
                break;
            case "Y3":
            case "3":
                c = 3.0;
                break;
            case "Y4":
            case "4":
                c = 4.0;
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
    public static double ACT_POWER_CURVE(
        [ExcelArgument(Description = "Proportion of sum insured (0 to 1)")] double d,
        [ExcelArgument(Description = "Power exponent n (> 0)")] double n)
    {
        if (d < 0 || d > 1) return double.NaN;
        if (n <= 0) return double.NaN;

        return Math.Pow(d, n);
    }

    [ExcelFunction(Description = "Inverse power exposure curve: G(d) = 1 - (1-d)^n", Category = "Actuarial.ExposureCurves")]
    public static double ACT_INVERSE_POWER_CURVE(
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
    public static double ACT_PARETO_EXPOSURE(
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

    #region Layer Rating using Exposure Curves

    [ExcelFunction(Description = "Calculate layer rate on line using exposure curve", Category = "Actuarial.ExposureCurves")]
    public static double ACT_LAYER_RATE_ON_LINE(
        [ExcelArgument(Description = "Attachment point as proportion of sum insured")] double attachmentPct,
        [ExcelArgument(Description = "Exhaustion point as proportion of sum insured")] double exhaustionPct,
        [ExcelArgument(Description = "Ground-up burning cost (rate on line)")] double burningCost,
        [ExcelArgument(Description = "Exposure curve b parameter")] double b,
        [ExcelArgument(Description = "Exposure curve g parameter")] double g)
    {
        if (attachmentPct < 0 || exhaustionPct > 1 || attachmentPct >= exhaustionPct)
            return double.NaN;

        double gAttach = ACT_MBBEFD(attachmentPct, b, g);
        double gExhaust = ACT_MBBEFD(exhaustionPct, b, g);

        // Layer rate = Ground-up rate * (G(exhaust) - G(attach)) / (exhaust - attach)
        double exposureFactor = (gExhaust - gAttach) / (exhaustionPct - attachmentPct);
        return burningCost * exposureFactor;
    }

    #endregion
}
