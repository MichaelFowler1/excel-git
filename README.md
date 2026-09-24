# excel-git

Git and GitHub treat `.xlsx` as an opaque binary blob. You can commit, fork and branch it, but a diff just says "binary file changed" and any merge where both sides touched the file is a conflict. This fixes that at the cell level.

## What you get

| | Without xlgit | With xlgit |
|---|---|---|
| `git diff` | `Binary files differ` | `-Budget!B2 1000` / `+Budget!B2 1100` |
| Merge, different cells edited | conflict, pick one whole file | merges cleanly |
| Merge, same cell edited | conflict | conflict, but only on that cell; both values in an Excel comment and a `_merge_conflicts` sheet |
| Pull request on GitHub | "binary file not shown" | bot comment with a table of every changed cell |

Formulas are compared as formulas (`=B2+C2`), not their cached results.

## Setup

In any repo holding Excel files:

```bash
pip install openpyxl
python path/to/xlgit.py install
```

That writes a `.gitattributes` entry and registers the diff and merge drivers in `.git/config`. The config part is per clone, so each collaborator runs `install` once. The `.gitattributes` part gets committed.

For the GitHub side, copy `.github/workflows/excel-diff.yml` and `xlgit.py` into the repo root. Every PR that touches a workbook gets a cell-diff comment.

## Commands

```
python xlgit.py textconv book.xlsx          # dump as text
python xlgit.py diff old.xlsx new.xlsx      # list cell changes
python xlgit.py diff a.xlsx b.xlsx --markdown
python xlgit.py merge base.xlsx ours.xlsx theirs.xlsx
```

## Limits

- Cell values and formulas only. Formatting, column widths, charts and pivot tables don't show up in the diff.
- Merging rewrites the file through openpyxl, which drops charts, images and pivot tables from the merged workbook. Fine for data and formula sheets, don't rely on it for dashboard-style workbooks.
- Row inserts show up as many changed cells, because every cell below the insert moves.
- `.xls` (the old pre-2007 format) isn't supported. Save as `.xlsx`.
