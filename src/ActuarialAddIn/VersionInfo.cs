namespace ActuarialAddIn;

/// <summary>
/// Contains version information and commit history for the Actuarial Add-In.
/// This file is updated during the build process or manually when releasing new versions.
/// </summary>
public static class VersionInfo
{
    public const string CurrentVersion = "0.1.0";
    public const string BuildDate = "2026-01-31";
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
            new CommitInfo("eb380cf", "2026-01-31", "Update validation report"),
            new CommitInfo("7682e43", "2026-01-31", "Fix ACT_DIST_GPD_FIT PWM formula bug"),
            new CommitInfo("d0b688f", "2026-01-31", "Fix ACT_ILF_LAYER bug and array parameter handling"),
            new CommitInfo("dab1263", "2026-01-31", "Regenerate spreadsheet with populate_examples.py"),
            new CommitInfo("108fea1", "2026-01-31", "Fix remaining formula errors in All Functions Test"),
            new CommitInfo("1c852e4", "2026-01-31", "Fix All Functions Test formulas and rename Excel file"),
            new CommitInfo("46d7c9b", "2026-01-31", "Update validation report"),
            new CommitInfo("080ece8", "2026-01-31", "Remove ACT_QS_CEDED, ACT_AGGREGATE_LAYER, and Berquist-Sherman from Excel"),
            new CommitInfo("b1a62e1", "2026-01-31", "Remove ACT_QS_CEDED and ACT_AGGREGATE_LAYER functions"),
            new CommitInfo("c98fd3e", "2026-01-31", "Update validation report timestamp"),
            new CommitInfo("6991e7e", "2026-01-31", "Reorganize README: move ILF/EP to Exposure section, rename Layer Functions"),
            new CommitInfo("7b2cd4e", "2026-01-31", "Fix test results path in run_tests.py"),
            new CommitInfo("2ac4970", "2026-01-31", "Fix run_tests.py same-file error when run from Windows path"),
            new CommitInfo("116ac37", "2026-01-31", "Add unified test runner for WSL"),
            new CommitInfo("bd64697", "2026-01-31", "Remove Berquist-Sherman method and format reconciliation tables"),
            new CommitInfo("152bf66", "2026-01-31", "Fix ToDoubleArray to handle object[] returns"),
            new CommitInfo("59665b0", "2026-01-31", "Add reconciliation checks to test results"),
            new CommitInfo("817b184", "2026-01-31", "Add chainladder-python and Bernegger references to README"),
            new CommitInfo("acd021e", "2026-01-31", "Add missing distributions to test workbook and expand README testing docs"),
            new CommitInfo("f5242ac", "2026-01-31", "Add WSL/Windows filesystem sync troubleshooting to agents.md")
        };
    }
}
