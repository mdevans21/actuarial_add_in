using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

public static class Interpolation
{
    [ExcelFunction(Description = "Linear interpolation with optional extrapolation", Category = "Actuarial.Utilities")]
    public static double ACT_INTERP(
        [ExcelArgument(Description = "X values (known points)")] double[] xValues,
        [ExcelArgument(Description = "Y values (known points)")] double[] yValues,
        [ExcelArgument(Description = "X value to interpolate")] double x,
        [ExcelArgument(Description = "Extrapolation mode: 'FLAT' (default), 'GRADIENT', or 'ERROR'")] string extrapolation = "FLAT")
    {
        if (xValues.Length != yValues.Length || xValues.Length < 2)
            return double.NaN;

        // Create sorted pairs
        var pairs = xValues.Zip(yValues, (xv, yv) => new { X = xv, Y = yv })
                           .OrderBy(p => p.X)
                           .ToArray();

        // Check if x is within range
        double xMin = pairs[0].X;
        double xMax = pairs[pairs.Length - 1].X;

        if (x < xMin)
        {
            // Extrapolation below range
            switch (extrapolation.ToUpper())
            {
                case "FLAT":
                    return pairs[0].Y;
                case "GRADIENT":
                    // Use gradient from first two points
                    double gradient = (pairs[1].Y - pairs[0].Y) / (pairs[1].X - pairs[0].X);
                    return pairs[0].Y - gradient * (pairs[0].X - x);
                case "ERROR":
                    return double.NaN;
                default:
                    return pairs[0].Y;
            }
        }

        if (x > xMax)
        {
            // Extrapolation above range
            int n = pairs.Length;
            switch (extrapolation.ToUpper())
            {
                case "FLAT":
                    return pairs[n - 1].Y;
                case "GRADIENT":
                    // Use gradient from last two points
                    double gradient = (pairs[n - 1].Y - pairs[n - 2].Y) / (pairs[n - 1].X - pairs[n - 2].X);
                    return pairs[n - 1].Y + gradient * (x - pairs[n - 1].X);
                case "ERROR":
                    return double.NaN;
                default:
                    return pairs[n - 1].Y;
            }
        }

        // Interpolation within range
        for (int i = 0; i < pairs.Length - 1; i++)
        {
            if (x >= pairs[i].X && x <= pairs[i + 1].X)
            {
                double t = (x - pairs[i].X) / (pairs[i + 1].X - pairs[i].X);
                return pairs[i].Y + t * (pairs[i + 1].Y - pairs[i].Y);
            }
        }

        return double.NaN;
    }

    [ExcelFunction(Description = "Bilinear interpolation for 2D data", Category = "Actuarial.Utilities")]
    public static double ACT_INTERP2D(
        [ExcelArgument(Description = "X values (row headers)")] double[] xValues,
        [ExcelArgument(Description = "Y values (column headers)")] double[] yValues,
        [ExcelArgument(Description = "Z values (2D grid)")] double[,] zValues,
        [ExcelArgument(Description = "X value to interpolate")] double x,
        [ExcelArgument(Description = "Y value to interpolate")] double y)
    {
        if (xValues.Length != zValues.GetLength(0) || yValues.Length != zValues.GetLength(1))
            return double.NaN;

        // Find bracketing indices for x
        int ix1 = 0, ix2 = 0;
        for (int i = 0; i < xValues.Length - 1; i++)
        {
            if (x >= xValues[i] && x <= xValues[i + 1])
            {
                ix1 = i;
                ix2 = i + 1;
                break;
            }
        }
        if (x < xValues[0]) { ix1 = 0; ix2 = 1; x = xValues[0]; }
        if (x > xValues[xValues.Length - 1]) { ix1 = xValues.Length - 2; ix2 = xValues.Length - 1; x = xValues[xValues.Length - 1]; }

        // Find bracketing indices for y
        int iy1 = 0, iy2 = 0;
        for (int j = 0; j < yValues.Length - 1; j++)
        {
            if (y >= yValues[j] && y <= yValues[j + 1])
            {
                iy1 = j;
                iy2 = j + 1;
                break;
            }
        }
        if (y < yValues[0]) { iy1 = 0; iy2 = 1; y = yValues[0]; }
        if (y > yValues[yValues.Length - 1]) { iy1 = yValues.Length - 2; iy2 = yValues.Length - 1; y = yValues[yValues.Length - 1]; }

        // Bilinear interpolation
        double tx = (x - xValues[ix1]) / (xValues[ix2] - xValues[ix1]);
        double ty = (y - yValues[iy1]) / (yValues[iy2] - yValues[iy1]);

        double z11 = zValues[ix1, iy1];
        double z12 = zValues[ix1, iy2];
        double z21 = zValues[ix2, iy1];
        double z22 = zValues[ix2, iy2];

        double z1 = z11 + ty * (z12 - z11);
        double z2 = z21 + ty * (z22 - z21);

        return z1 + tx * (z2 - z1);
    }

    [ExcelFunction(Description = "Log-linear interpolation (common for yield curves)", Category = "Actuarial.Utilities")]
    public static double ACT_INTERP_LOG(
        [ExcelArgument(Description = "X values (known points)")] double[] xValues,
        [ExcelArgument(Description = "Y values (known points)")] double[] yValues,
        [ExcelArgument(Description = "X value to interpolate")] double x,
        [ExcelArgument(Description = "Extrapolation mode: 'FLAT' (default), 'GRADIENT', or 'ERROR'")] string extrapolation = "FLAT")
    {
        // Convert to log space and interpolate
        var logX = xValues.Select(v => v > 0 ? Math.Log(v) : double.NaN).ToArray();
        if (logX.Any(v => double.IsNaN(v))) return double.NaN;
        if (x <= 0) return double.NaN;

        return ACT_INTERP(logX, yValues, Math.Log(x), extrapolation);
    }
}
