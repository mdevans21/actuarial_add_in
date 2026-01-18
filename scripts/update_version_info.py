#!/usr/bin/env python3
"""
Update VersionInfo.cs with the latest git commit history.
Run this script before building the add-in to ensure version info is current.
"""

import subprocess
import datetime
from pathlib import Path

def get_git_commits(count=20):
    """Get recent git commits."""
    result = subprocess.run(
        ['git', 'log', f'--pretty=format:%h|%ad|%s', '--date=short', f'-{count}'],
        capture_output=True,
        text=True,
        cwd=Path(__file__).parent.parent
    )
    commits = []
    for line in result.stdout.strip().split('\n'):
        if line:
            parts = line.split('|', 2)
            if len(parts) == 3:
                commits.append({
                    'hash': parts[0],
                    'date': parts[1],
                    'message': parts[2].replace('"', '\\"')
                })
    return commits

def generate_version_info():
    """Generate the VersionInfo.cs content."""
    commits = get_git_commits()
    build_date = datetime.date.today().isoformat()

    commit_entries = ',\n            '.join([
        f'new CommitInfo("{c["hash"]}", "{c["date"]}", "{c["message"]}")'
        for c in commits
    ])

    content = f'''namespace ActuarialAddIn;

/// <summary>
/// Contains version information and commit history for the Actuarial Add-In.
/// This file is updated during the build process or manually when releasing new versions.
/// </summary>
public static class VersionInfo
{{
    public const string CurrentVersion = "0.1.0";
    public const string BuildDate = "{build_date}";
    public const string GitHubUrl = "https://github.com/mdevans21/actuarial_add_in";

    public record CommitInfo(string ShortHash, string Date, string Message);

    /// <summary>
    /// Returns the commit history for the add-in.
    /// </summary>
    public static IEnumerable<CommitInfo> GetCommitHistory()
    {{
        // Commit history - updated during build or release
        return new[]
        {{
            {commit_entries},
        }};
    }}
}}
'''
    return content

def main():
    version_info_path = Path(__file__).parent.parent / 'src' / 'ActuarialAddIn' / 'VersionInfo.cs'
    content = generate_version_info()

    version_info_path.write_text(content)
    print(f"Updated {version_info_path}")
    print(f"Build date: {datetime.date.today().isoformat()}")

    commits = get_git_commits()
    print(f"Included {len(commits)} commits")

if __name__ == '__main__':
    main()
