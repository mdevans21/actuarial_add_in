using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Helpers for the optional `seed` parameter on stochastic functions.
///
/// Excel-DNA 1.9.0 nominally supports nullable parameter types like `int?`,
/// but in practice every function in this add-in that exposed `int? seed`
/// returned #VALUE! when called from Excel even though the same call worked
/// from C# directly. Switching the registered type to `object` and parsing
/// internally is the standard Excel-DNA workaround — Excel passes
/// `ExcelMissing.Value` for omitted optional args, `ExcelEmpty.Value` for
/// blank cell refs, and `double` for any number cell, all of which need
/// distinct handling.
/// </summary>
internal static class SeedUtil
{
    public static bool TryResolveSeed(object? seed, out int? resolvedSeed, out string error)
    {
        resolvedSeed = null;
        error = "";

        if (seed is null or ExcelMissing or ExcelEmpty)
            return true;

        long value;
        switch (seed)
        {
            case int i:
                value = i;
                break;
            case long l:
                value = l;
                break;
            case double d when !double.IsNaN(d) && !double.IsInfinity(d) && d == Math.Truncate(d):
                if (d < 0 || d > int.MaxValue)
                    return Invalid(out error);
                value = (long)d;
                break;
            default:
                return Invalid(out error);
        }

        if (value < 0 || value > int.MaxValue)
            return Invalid(out error);

        resolvedSeed = (int)value;
        return true;
    }

    private static bool Invalid(out string error)
    {
        error = "Error: seed must be a whole number between 0 and 2147483647";
        return false;
    }
}
