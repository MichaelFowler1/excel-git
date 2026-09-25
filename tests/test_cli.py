"""The command line: setup, everyday commands, and failing gently.

Every test gets its own HOME and global git config, so nothing here touches
the real ~/.gitconfig.
"""
import os
import subprocess
import sys
from pathlib import Path

import pytest

from test_cell_types import book

XLGIT = Path(__file__).resolve().parents[1] / "xlgit.py"


@pytest.fixture
def env(tmp_path, monkeypatch):
    home = tmp_path / "home"
    home.mkdir()
    e = dict(os.environ, HOME=str(home), XDG_CONFIG_HOME=str(home / ".config"),
             GIT_CONFIG_GLOBAL=str(home / ".gitconfig"), GIT_CONFIG_NOSYSTEM="1",
             GIT_AUTHOR_NAME="t", GIT_AUTHOR_EMAIL="t@t", GIT_COMMITTER_NAME="t", GIT_COMMITTER_EMAIL="t@t")
    e.pop("XLGIT_DEBUG", None)
    return e


def run(env, cwd, *args):
    return subprocess.run([sys.executable, str(XLGIT), *args], cwd=cwd, env=env, capture_output=True, text=True)


def git(env, cwd, *args):
    return subprocess.run(["git", *args], cwd=cwd, env=env, check=True, capture_output=True, text=True).stdout


def new_repo(env, path):
    path.mkdir()
    git(env, path, "init", "-q", "-b", "main")
    return path


def test_help_shows_setup_status(env, tmp_path):
    r = run(env, tmp_path, )
    assert r.returncode == 0
    assert "xlgit install" in r.stdout and "Not set up yet" in r.stdout


def test_install_once_works_in_every_repo(env, tmp_path):
    r = run(env, tmp_path, "install")
    assert r.returncode == 0, r.stderr
    assert "every git repository" in r.stdout
    repo = new_repo(env, tmp_path / "later")  # a repo made after setup, with no .gitattributes
    assert "merge: xlsx" in git(env, repo, "check-attr", "merge", "--", "book.xlsx")
    assert not (repo / ".gitattributes").exists()
    assert "[ok] Set up for every repository" in run(env, repo).stdout
    run(env, tmp_path, "install")  # twice is harmless
    attrs = (tmp_path / "home" / ".config" / "git" / "attributes").read_text()
    assert attrs.count("*.xlsx diff=xlsx merge=xlsx") == 1


def test_uninstall(env, tmp_path):
    run(env, tmp_path, "install")
    assert run(env, tmp_path, "uninstall").returncode == 0
    assert "Not set up yet" in run(env, tmp_path).stdout


def test_install_repo_needs_a_repo(env, tmp_path):
    r = run(env, tmp_path, "install", "--repo")
    assert r.returncode == 2
    assert "Traceback" not in r.stderr and "isn't inside a git repository" in r.stderr


def test_install_github_writes_workflow(env, tmp_path):
    repo = new_repo(env, tmp_path / "r")
    r = run(env, repo, "install", "--github")
    assert r.returncode == 0, r.stderr
    wf = (repo / ".github" / "workflows" / "excel-diff.yml").read_text()
    assert "uses: MichaelFowler1/excel-git@v" in wf and "pull-requests: write" in wf
    assert "git push" in r.stdout


def test_diff_without_arguments_shows_changes_since_last_commit(env, tmp_path):
    run(env, tmp_path, "install")
    repo = new_repo(env, tmp_path / "r")
    book(repo / "budget.xlsx", {"A1": "Rent", "B1": 1000})
    git(env, repo, "add", "-A")
    git(env, repo, "commit", "-qm", "first")
    assert "No workbook changes" in run(env, repo, "diff").stdout
    book(repo / "budget.xlsx", {"A1": "Rent", "B1": 1100})
    book(repo / "new.xlsx", {"A1": "hi"})
    r = run(env, repo, "diff")
    assert "Budget" not in r.stderr and "Traceback" not in r.stderr
    assert "Data!B1  1000 -> 1100" in r.stdout
    assert "new.xlsx" in r.stdout and "Data!A1" in r.stdout
    one = run(env, repo, "diff", "budget.xlsx").stdout
    assert "1000 -> 1100" in one and "new.xlsx" not in one
    # and plain git diff shows cells too, once set up
    assert "changed        Data!B1  1000 -> 1100" in git(env, repo, "diff")
    # git log -p goes through the text form
    git(env, repo, "commit", "-qam", "raise rent")
    assert "+Data!B1\t1100" in git(env, repo, "log", "-p", "-1")


def test_friendly_errors(env, tmp_path):
    r = run(env, tmp_path, "diff", "nope.xlsx", "also-nope.xlsx")
    assert r.returncode == 2 and "no such file" in r.stderr and "Traceback" not in r.stderr
    assert "unknown command" in run(env, tmp_path, "frobnicate").stderr


def test_a_broken_file_never_breaks_git_diff(env, tmp_path):
    bad = tmp_path / "old.xlsx"
    bad.write_bytes(b"this is really an old .xls file")
    r = run(env, tmp_path, "textconv", str(bad))
    assert r.returncode == 0 and "save as .xlsx" in r.stdout


def test_a_merge_that_fails_keeps_your_file_and_says_what_to_do(env, tmp_path):
    base, theirs = book(tmp_path / "base.xlsx", {"A1": 1}), book(tmp_path / "theirs.xlsx", {"A1": 2})
    ours = tmp_path / "ours.xlsx"
    ours.write_bytes(b"not a workbook")
    r = run(env, tmp_path, "merge", base, str(ours), theirs, "budget.xlsx")
    assert r.returncode == 1
    assert ours.read_bytes() == b"not a workbook"
    assert "Your version is kept" in r.stderr and "git checkout --theirs" in r.stderr
    assert "Traceback" not in r.stderr


def test_conflicts_say_how_to_finish(env, tmp_path):
    base = book(tmp_path / "base.xlsx", {"A1": 1})
    ours = book(tmp_path / "ours.xlsx", {"A1": 2})
    theirs = book(tmp_path / "theirs.xlsx", {"A1": 3})
    r = run(env, tmp_path, "merge", base, ours, theirs, "budget.xlsx")
    assert r.returncode == 1
    assert "Data A1: yours 2, theirs 3 (was 1)" in r.stderr
    assert "_merge_conflicts" in r.stderr and 'git add "budget.xlsx"' in r.stderr


def test_pr_comment(env, tmp_path):
    repo = new_repo(env, tmp_path / "r")
    book(repo / "budget.xlsx", {"B1": 1000})
    git(env, repo, "add", "-A")
    git(env, repo, "commit", "-qm", "base")
    git(env, repo, "checkout", "-qb", "change")
    book(repo / "budget.xlsx", {"B1": 1100})
    git(env, repo, "commit", "-qam", "raise")
    r = run(env, repo, "pr-comment", "main", "change")
    assert r.returncode == 0, r.stderr
    assert r.stdout.startswith("<!-- xlgit-pr-comment -->")
    assert "| Data | B1 | changed | 1000 | 1100 |" in r.stdout
