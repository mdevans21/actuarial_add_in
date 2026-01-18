using ExcelDna.Integration;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Version information functions for the Actuarial Add-In.
/// </summary>
public static class Version
{
    /// <summary>
    /// Returns the current version of the Actuarial Add-In.
    /// </summary>
    [ExcelFunction(
        Name = "ACT_VERSION",
        Description = "Returns the current version of the Actuarial Add-In.",
        Category = "Actuarial - Info")]
    public static string ACT_VERSION()
    {
        return VersionInfo.CurrentVersion;
    }

    /// <summary>
    /// Returns the build date of the Actuarial Add-In.
    /// </summary>
    [ExcelFunction(
        Name = "ACT_BUILD_DATE",
        Description = "Returns the build date of the Actuarial Add-In.",
        Category = "Actuarial - Info")]
    public static string ACT_BUILD_DATE()
    {
        return VersionInfo.BuildDate;
    }

    /// <summary>
    /// Returns the GitHub repository URL for the Actuarial Add-In.
    /// </summary>
    [ExcelFunction(
        Name = "ACT_GITHUB_URL",
        Description = "Returns the GitHub repository URL for the Actuarial Add-In.",
        Category = "Actuarial - Info")]
    public static string ACT_GITHUB_URL()
    {
        return VersionInfo.GitHubUrl;
    }

    /// <summary>
    /// Returns commit information from the version history.
    /// </summary>
    /// <param name="index">The 1-based index of the commit (1 = most recent).</param>
    /// <param name="field">The field to return: "hash", "date", or "message". Defaults to "message".</param>
    [ExcelFunction(
        Name = "ACT_COMMIT_INFO",
        Description = "Returns commit information from the version history.",
        Category = "Actuarial - Info")]
    public static object ACT_COMMIT_INFO(
        [ExcelArgument(Description = "The 1-based index of the commit (1 = most recent)")] int index,
        [ExcelArgument(Description = "The field to return: 'hash', 'date', or 'message'. Defaults to 'message'.")] string field = "message")
    {
        var commits = VersionInfo.GetCommitHistory().ToList();

        if (index < 1 || index > commits.Count)
        {
            return ExcelError.ExcelErrorValue;
        }

        var commit = commits[index - 1];

        return field.ToLower() switch
        {
            "hash" => commit.ShortHash,
            "date" => commit.Date,
            "message" => commit.Message,
            _ => commit.Message
        };
    }

    /// <summary>
    /// Returns the total number of commits in the version history.
    /// </summary>
    [ExcelFunction(
        Name = "ACT_COMMIT_COUNT",
        Description = "Returns the total number of commits in the version history.",
        Category = "Actuarial - Info")]
    public static int ACT_COMMIT_COUNT()
    {
        return VersionInfo.GetCommitHistory().Count();
    }

    /// <summary>
    /// Returns an array of all commits in the version history.
    /// </summary>
    [ExcelFunction(
        Name = "ACT_COMMIT_HISTORY",
        Description = "Returns an array of all commits in the version history (hash, date, message).",
        Category = "Actuarial - Info")]
    public static object[,] ACT_COMMIT_HISTORY()
    {
        var commits = VersionInfo.GetCommitHistory().ToList();
        var result = new object[commits.Count + 1, 3];

        // Header row
        result[0, 0] = "Commit";
        result[0, 1] = "Date";
        result[0, 2] = "Message";

        // Data rows
        for (int i = 0; i < commits.Count; i++)
        {
            result[i + 1, 0] = commits[i].ShortHash;
            result[i + 1, 1] = commits[i].Date;
            result[i + 1, 2] = commits[i].Message;
        }

        return result;
    }
}
