# excel-git

Git and GitHub treat `.xlsx` as an opaque binary blob. You can commit, fork and branch it, but a diff just says "binary file changed" and any merge where both sides touched the file is a conflict. This fixes that.

## What you get

| | Without xlgit | With xlgit |
|---|---|---|
| `git diff` | `Binary files differ` | `-Budget!B2 1000` / `+Budget!B2 1100`, plus chart, image and comment changes |
| Merge, different cells edited | conflict, pick one whole file | merges cleanly |
| Merge, same cell edited | conflict | conflict on just that cell, listed in a `_merge_conflicts` sheet with a link to it |
| Charts, images, comments, formatting, macros | n/a | kept, and their edits to them carried over |
| Pull request on GitHub | "binary file not shown" | bot comment with a table of every changed cell and chart |

Formulas are compared as formulas (`=B2+C2`), not their cached results.

## How the merge works

A workbook is a zip of XML files: one per sheet, one per chart, one per image and so on. Instead of re-saving the whole thing through a spreadsheet library (which is how charts used to get lost), xlgit starts from your copy's zip and only rewrites the XML that has to change:

- **Cells**: 3-way merge cell by cell. A cell only one side changed takes that side's value.
- **Charts, images, comments, macros**: 3-way merge object by object. If only their branch changed a chart, you get their version. If both did, yours is kept and it's flagged as a conflict. Chart edits caused by cell changes (Excel caches plotted values inside the chart) don't count as edits.
- **Sheets**: new sheets on their branch come over whole, charts included. Renames and deletes merge too.
- **Named ranges**: merged by name.

Excel recalculates every formula when it opens the merged file.

## Setup

In any repo holding Excel files:

```bash
pip install openpyxl lxml
python path/to/xlgit.py install
```

That writes a `.gitattributes` entry and registers the diff and merge drivers in `.git/config`. The config part is per clone, so each collaborator runs `install` once. The `.gitattributes` part gets committed.

For the GitHub side, copy `.github/workflows/excel-diff.yml`, `requirements.txt` and `xlgit.py` into the repo root. Every PR that touches a workbook gets a cell-diff comment.

## Commands

```
python xlgit.py textconv book.xlsx          # dump as text
python xlgit.py diff old.xlsx new.xlsx      # list cell and object changes
python xlgit.py diff a.xlsx b.xlsx --markdown
python xlgit.py merge base.xlsx ours.xlsx theirs.xlsx
```

## Tests

```bash
pip install -r requirements-dev.txt
python -m pytest tests
```

The tests build workbook versions with charts, comments and named ranges, run real `git merge` through the driver, and check every object survived. Set `XLGIT_KEEP=some/dir` to keep the merged files and open them in Excel yourself.

## Limits

- Formatting changes their branch made to existing cells come over only when both branches have the same set of styles. Otherwise your formatting is kept.
- Column widths, merged cells and conditional formatting on existing sheets aren't merged. Yours are kept.
- Tables and pivot tables added on their branch aren't merged yet. They're reported as conflicts so nothing disappears silently.
- Inserting a row shows up as many changed cells, because every cell below it moves.
- `.xls` (the old pre-2007 format) isn't supported. Save as `.xlsx`.
