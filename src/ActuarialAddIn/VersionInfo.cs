namespace ActuarialAddIn;

/// <summary>
/// Contains version information and commit history for the Actuarial Add-In.
/// This file is updated during the build process or manually when releasing new versions.
/// </summary>
public static class VersionInfo
{
    public const string CurrentVersion = "0.7.0";
    public const string BuildDate = "2026-05-04";
    public const string GitHubUrl = "https://github.com/mdevans21/actuarial_add_in";

    public record CommitInfo(string ShortHash, string Date, string Message);

    /// <summary>
    /// Returns the commit history for the add-in.
    /// </summary>
    public static IEnumerable<CommitInfo> GetCommitHistory()
    {
        // Release history — see GitHub Releases for full notes per tag.
        return new[]
        {
            new CommitInfo("v0.7.0", "2026-05-04", "Breaking: dropped Cape Cod, triangle-utility, calendar-adjust, weighted-average functions (9 in total). Added Mack factor SE / BF / CL_LATEST formulas to spreadsheet; harness now exercises every C# function (174); notebook reconciles every non-trivial function against scipy/Klugman/Bernegger references."),
            new CommitInfo("v0.6.0", "2026-05-04", "Internal cleanups: split Distributions.cs (1686 lines) and ChainLadder.cs (1510 lines) into focused partial-class files; standardised inline string-error returns to ExcelError.ExcelErrorValue (24 sites); test harness now exits non-zero on assertion failures."),
            new CommitInfo("v0.5.4", "2026-05-04", "Patch: ACT_EXPOSURE_MBBEFD numerical stability — rewrote without explicit `a` to fix NaN at SwissRe c=5 (b<1, b^g underflow) and a 4th-decimal precision loss at c=3."),
            new CommitInfo("v0.5.3", "2026-05-04", "Patch: revert to net6.0-windows — neither ExcelDna 1.9.0 nor 1.10-preview4 reliably hosted .NET 8 in Excel. Keeps int? from 1.9; net6 path is known-good."),
            new CommitInfo("v0.5.2", "2026-05-04", "Patch: tried ExcelDna.AddIn 1.10-preview4 to fix .NET 8 hosting (didn't work)."),
            new CommitInfo("v0.5.1", "2026-05-03", "Patch: drop legacy RuntimeVersion=\"v4.0\" from .dna manifest (necessary but not sufficient — see v0.5.2)."),
            new CommitInfo("v0.5.0", "2026-05-03", "Numerical fixes (LAYER_RATE, ILF_PARETO with xm, SWISSRE→Bernegger, AGGREGATE_TVAR); ExcelDna 1.9 + net8.0-windows; nullable seed; copula error handling."),
            new CommitInfo("v0.4.0", "2026-05-02", "GLM hat matrix bootstrap; Aggregate Claims formula fix; broader disclaimers."),
            new CommitInfo("v0.3.1", "2026-02-21", "Aggregate Claims tab; spreadsheet rebuilt as .xlsx with full coverage."),
            new CommitInfo("v0.3.0", "2026-02-21", "Enhanced ODP Bootstrap (E&V 2002); Cape Cod method; credibility removed."),
            new CommitInfo("v0.2.0", "2026-02-04", "Panjer recursion; ZT/ZM distributions; Pareto III/IV; Inverse Gaussian; Loglogistic; cat modelling."),
            new CommitInfo("v0.1.0", "2026-01-31", "Initial tagged release: distributions, LEV, chain ladder, Mack, exposure curves.")
        };
    }
}
