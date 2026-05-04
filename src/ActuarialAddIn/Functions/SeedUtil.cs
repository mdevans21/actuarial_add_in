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
    public static int? ResolveSeed(object? seed)
    {
        if (seed is null) return null;
        if (seed is ExcelMissing) return null;
        if (seed is ExcelEmpty) return null;
        if (seed is double d)
        {
            if (double.IsNaN(d)) return null;
            return (int)d;
        }
        if (seed is int i) return i;
        if (seed is long l) return (int)l;
        return null;
    }
}
