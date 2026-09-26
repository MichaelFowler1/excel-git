"""The command line: setup, everyday commands, and failing gently.

Every test gets its own HOME and global git config, so nothing here touches
the real ~/.gitconfig.
"""
import os
import subprocess
import sys
from pathlib import Path

import pytest

import xlgit
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


def test_scrub_keeps_structure_and_equal_values(env, tmp_path):
    import zipfile
    import xlgit
    path = book(tmp_path / "secret.xlsx", {"A1": "Acme Corp", "A2": "Acme Corp", "B1": 1234.5, "B2": 1234.5, "C1": 0})
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    xml = parts["xl/worksheets/sheet1.xml"].decode()
    parts["xl/worksheets/sheet1.xml"] = xml.replace("</row>", '<c r="D1"><f>IF(B1&gt;0,"Acme Corp","x")</f><v>9</v></c></row>', 1).encode()
    parts["docProps/core.xml"] = parts["docProps/core.xml"].replace(b"</cp:coreProperties>",
        b"<dc:creator>Jane Secret</dc:creator></cp:coreProperties>")
    with zipfile.ZipFile(path, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)
    r = run(env, tmp_path, "scrub", path)
    assert r.returncode == 0, r.stderr
    out = tmp_path / "secret.scrubbed.xlsx"
    cells = xlgit.read_cells(str(out))["Data"]
    assert cells["A1"] == cells["A2"] != "Acme Corp" and len(cells["A1"]) == len("Acme Corp")
    assert cells["B1"] == cells["B2"] != 1234.5 and cells["C1"] == 0
    assert cells["D1"].startswith('=IF(B1>0,"') and "Acme" not in cells["D1"]
    blob = b"".join(zipfile.ZipFile(out).read(n) for n in zipfile.ZipFile(out).namelist())
    assert b"Acme" not in blob and b"Jane Secret" not in blob and b"1234.5" not in blob


def test_scrub_merge_report(env, tmp_path):
    run(env, tmp_path, "install")
    repo = new_repo(env, tmp_path / "r")
    book(repo / "b.xlsx", {"A1": "Private", "B1": 1})
    git(env, repo, "add", "-A")
    git(env, repo, "commit", "-qm", "base")
    git(env, repo, "checkout", "-qb", "other")
    book(repo / "b.xlsx", {"A1": "Private", "B1": 2})
    git(env, repo, "commit", "-qam", "two")
    git(env, repo, "checkout", "-q", "main")
    book(repo / "b.xlsx", {"A1": "Private", "B1": 3})
    git(env, repo, "commit", "-qam", "three")
    subprocess.run(["git", "merge", "other"], cwd=repo, env=env, capture_output=True)
    r = run(env, repo, "scrub", "--merge", "b.xlsx")
    assert r.returncode == 0, r.stderr
    import zipfile
    z = zipfile.ZipFile(repo / "b-merge-report.zip")
    assert {"base.xlsx", "ours.xlsx", "theirs.xlsx"} <= set(z.namelist())
    assert all(b"Private" not in z.read(n) for n in z.namelist() if n.endswith(".xlsx"))


def test_any_characters_reach_git_intact(env, tmp_path):
    """Windows consoles and pipes default to a code page without most
    symbols; output must still reach git as UTF-8, never crash."""
    book(tmp_path / "old.xlsx", {"A1": "plain"})
    book(tmp_path / "new.xlsx", {"A1": "done \u2713 caf\u00e9 \u6771\u4eac"})
    r = subprocess.run([sys.executable, str(XLGIT), "diff", str(tmp_path / "old.xlsx"), str(tmp_path / "new.xlsx")],
                       cwd=tmp_path, env=dict(env, PYTHONIOENCODING="ascii"), capture_output=True)
    assert r.returncode in (0, 1), r.stderr.decode("utf-8", "replace")
    assert "done \u2713 caf\u00e9 \u6771\u4eac" in r.stdout.decode("utf-8")


def test_demo_merges_both_estimators(env, tmp_path):
    folder = tmp_path / "demo"
    r = run(env, tmp_path, "demo", str(folder), "--no-open")
    assert r.returncode == 0, r.stdout + r.stderr
    assert "took 2 cell(s) from the other branch, no conflicts" in r.stdout
    assert "$140,340" in r.stdout and "$93,900" in r.stdout
    cells = xlgit.read_cells(str(folder / "estimate.xlsx"))["Estimate"]
    assert cells["A4"] == "Roofing"                      # Anna's inserted line
    assert cells["A5"] == "Electrical" and cells["D5"] == 4.6  # Ben's rate, moved down a row
    assert cells["A6"] == "Plumbing" and cells["C6"] == 22     # Ben's quantity
    assert cells["E8"] == "=SUM(E2:E7)"
    assert any(o.startswith("chart:") for o in xlgit.describe_objects(
        xlgit.Package.open(str(folder / "estimate.xlsx")))["Estimate"])
    assert (folder / "changes.html").exists()
    # The demo stays inside its folder: the real global git config is untouched.
    assert not os.path.exists(env["GIT_CONFIG_GLOBAL"]) or "xlsx" not in Path(env["GIT_CONFIG_GLOBAL"]).read_text()


def test_demo_refuses_a_folder_with_files(env, tmp_path):
    (tmp_path / "busy").mkdir()
    (tmp_path / "busy" / "keep.txt").write_text("mine")
    r = run(env, tmp_path, "demo", str(tmp_path / "busy"), "--no-open")
    assert r.returncode == 2 and "isn't empty" in r.stderr
    assert (tmp_path / "busy" / "keep.txt").read_text() == "mine"


def test_demo_ignores_the_users_own_git_settings(env, tmp_path):
    """Fast-forward-only merges, signed commits and hooks set globally are
    the user's business; the demo's throwaway repository shouldn't trip on them."""
    hooks = tmp_path / "hooks"
    hooks.mkdir()
    (hooks / "pre-commit").write_text("#!/bin/sh\nexit 1\n")
    (hooks / "pre-commit").chmod(0o755)
    settings = {"merge.ff": "only", "commit.gpgsign": "true", "core.hooksPath": str(hooks)}
    for k, v in settings.items():  # in the (test's own) global ~/.gitconfig, where users keep them
        git(env, tmp_path, "config", "--global", k, v)
    r = run(env, tmp_path, "demo", str(tmp_path / "demo"), "--no-open")
    assert r.returncode == 0, r.stdout + r.stderr
    assert "no conflicts" in r.stdout


def test_diff_html_out_lands_where_you_are(env, tmp_path):
    repo = new_repo(env, tmp_path / "repo")
    sub = repo / "sub"
    sub.mkdir()
    book(sub / "b.xlsx", {"A1": "before"})
    git(env, repo, "add", "-A")
    git(env, repo, "commit", "-q", "-m", "base")
    book(sub / "b.xlsx", {"A1": "after"})
    r = run(env, sub, "diff", "--html", "--out=report.html", "--no-open")
    assert r.returncode == 0, r.stdout + r.stderr
    assert (sub / "report.html").exists() and not (repo / "report.html").exists()
    assert "after" in (sub / "report.html").read_text(encoding="utf-8")
    r = run(env, sub, "diff")  # plain diff from a subfolder still finds the change
    assert "'before' -> 'after'" in r.stdout
