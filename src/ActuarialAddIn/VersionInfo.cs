namespace ActuarialAddIn;

/// <summary>
/// Contains version information and commit history for the Actuarial Add-In.
/// This file is updated during the build process or manually when releasing new versions.
/// </summary>
public static class VersionInfo
{
    public const string CurrentVersion = "0.1.0";
    public const string BuildDate = "2026-01-21";
    public const string GitHubUrl = "https://github.com/mdevans21/actuarial_add_in";

    public record CommitInfo(string ShortHash, string Date, string Message);

    /// <summary>
    /// Returns the commit history for the add-in.
    /// </summary>
    public static IEnumerable<CommitInfo> GetCommitHistory()
    {
        // Commit history - updated during build or release
        return new[]
        {
            new CommitInfo("b1b839f", "2026-01-19", "Repository cleanup: remove obsolete files and fix code quality issues"),
            new CommitInfo("09491bc", "2026-01-19", "Improve consistency across outputs and documentation"),
            new CommitInfo("2b8cb04", "2026-01-18", "Fix ODP bootstrap to match England & Verrall (2002)"),
            new CommitInfo("fd0205a", "2026-01-18", "Add examples for all 10 distributions to Distributions tab"),
            new CommitInfo("fb7f919", "2026-01-18", "Add tests, utility scripts, and example spreadsheets"),
            new CommitInfo("97704d9", "2026-01-18", "Add ribbon UI and version tracking system"),
            new CommitInfo("b2a74d4", "2026-01-18", "Add Phase 6: Credibility and experience rating functions"),
            new CommitInfo("8e5109b", "2026-01-18", "Add Phase 5: Reserving enhancements - Cape Cod, triangle utilities, calendar year"),
            new CommitInfo("6a2d830", "2026-01-18", "Add Phase 3: Archimedean copulas and utility functions"),
            new CommitInfo("2f71e98", "2026-01-18", "Add Phase 2: Parameter estimation functions for distribution fitting"),
            new CommitInfo("79ef7bc", "2026-01-18", "Implement Phase 1 from next_steps.md: distributions, copulas, and help text"),
            new CommitInfo("bc30112", "2026-01-18", "Add comprehensive review and roadmap for actuarial add-in"),
            new CommitInfo("fb0a7fe", "2026-01-17", "Merge branch 'main' of https://github.com/mdevans21/actuarial_add_in"),
            new CommitInfo("83a9008", "2026-01-17", "Add chain ladder enhancements and update examples"),
            new CommitInfo("b16e651", "2026-01-17", "Update README.md"),
            new CommitInfo("7bb88e5", "2026-01-17", "Update README.md"),
            new CommitInfo("7458388", "2026-01-17", "Add charts to Excel sheets and update Chain Ladder with Taylor-Ashe data"),
            new CommitInfo("dc52de8", "2026-01-17", "Add documentation, examples, and agent instructions"),
            new CommitInfo("1fbd1ef", "2026-01-17", "Rename PRD file to fix double extension"),
            new CommitInfo("d2323be", "2026-01-17", "Fix Poisson inverse CDF implementation")
        };
    }
}
