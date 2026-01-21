#!/usr/bin/env python3
"""
Update VersionInfo.cs with current git commit history.

This script reads the git log and updates the VersionInfo.cs file with:
- Current build date
- Recent commit history (short hash, date, message)

Usage:
    python scripts/update_version_info.py [--max-commits N]
"""

import subprocess
import os
import re
from datetime import date
import argparse


# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VERSION_INFO_PATH = os.path.join(SCRIPT_DIR, '..', 'src', 'ActuarialAddIn', 'VersionInfo.cs')


def get_git_commits(max_commits=20):
    """Get recent git commits as list of (hash, date, message) tuples."""
    try:
        result = subprocess.run(
            ['git', 'log', f'--max-count={max_commits}', '--format=%h|%ad|%s', '--date=short'],
            capture_output=True,
            text=True,
            check=True,
            cwd=os.path.dirname(SCRIPT_DIR)
        )
        commits = []
        for line in result.stdout.strip().split('\n'):
            if line:
                parts = line.split('|', 2)
                if len(parts) == 3:
                    short_hash, commit_date, message = parts
                    # Escape quotes in message for C# string
                    message = message.replace('"', '\\"')
                    commits.append((short_hash, commit_date, message))
        return commits
    except subprocess.CalledProcessError as e:
        print(f"Error running git log: {e}")
        return []


def generate_commit_array(commits):
    """Generate C# code for the commit history array."""
    lines = []
    for i, (short_hash, commit_date, message) in enumerate(commits):
        comma = "," if i < len(commits) - 1 else ""
        lines.append(f'            new CommitInfo("{short_hash}", "{commit_date}", "{message}"){comma}')
    return '\n'.join(lines)


def update_version_info(max_commits=20):
    """Update VersionInfo.cs with current git history."""
    # Read current file
    with open(VERSION_INFO_PATH, 'r') as f:
        content = f.read()

    # Get git commits
    commits = get_git_commits(max_commits)
    if not commits:
        print("No commits found or git error. Aborting.")
        return False

    # Update BuildDate
    today = date.today().isoformat()
    content = re.sub(
        r'public const string BuildDate = "[^"]*";',
        f'public const string BuildDate = "{today}";',
        content
    )

    # Generate new commit array
    commit_array = generate_commit_array(commits)

    # Replace the commit history array
    # Match from "return new[]" to the closing "};"
    pattern = r'(return new\[\]\s*\{)\s*\n.*?\n(\s*\};)'
    replacement = f'\\1\n{commit_array}\n\\2'
    content = re.sub(pattern, replacement, content, flags=re.DOTALL)

    # Write updated file
    with open(VERSION_INFO_PATH, 'w') as f:
        f.write(content)

    print(f"Updated {VERSION_INFO_PATH}")
    print(f"  - BuildDate: {today}")
    print(f"  - Commits: {len(commits)}")
    return True


def main():
    parser = argparse.ArgumentParser(description='Update VersionInfo.cs with git commit history')
    parser.add_argument('--max-commits', type=int, default=20,
                        help='Maximum number of commits to include (default: 20)')
    args = parser.parse_args()

    print("Updating VersionInfo.cs from git history...")
    success = update_version_info(args.max_commits)
    if success:
        print("Done!")
    else:
        print("Failed to update VersionInfo.cs")
        exit(1)


if __name__ == "__main__":
    main()
