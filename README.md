# xlgit

> Beta (0.1.0). It's tested end to end, but it hasn't met many real-world workbooks yet. Git keeps every version, so a bad merge can always be undone. Please [open an issue](https://github.com/MichaelFowler1/excel-git/issues) when something looks wrong.

Git and GitHub treat `.xlsx` as an opaque binary blob. You can commit, fork and branch it, but a diff just says "binary file changed" and any merge where both sides touched the file is a conflict. This fixes that.

## What you get

| | Without xlgit | With xlgit |
|---|---|---|
| `git diff` | `Binary files differ` | `-Budget!B2 1000` / `+Budget!B2 1100`, plus chart, image and comment changes |
| Merge, different cells edited | conflict, pick one whole file | merges cleanly |
| Merge, same cell edited | conflict | conflict on just that cell, listed in a `_merge_conflicts` sheet with a link to it |
| Charts, images, comments, formatting, macros | n/a | kept, and their edits to them carried over |
| Tables and pivot tables | n/a | merged: you add rows, they add a column, you get both |
| Pull request on GitHub | "binary file not shown" | bot comment with a table of every changed cell, chart, table and pivot |

Formulas are compared as formulas (`=B2+C2`), not their cached results.

## How the merge works

A workbook is a zip of XML files: one per sheet, one per chart, one per image and so on. Instead of re-saving the whole thing through a spreadsheet library (which is how charts used to get lost), xlgit starts from your copy's zip and only rewrites the XML that has to change:

- **Cells**: 3-way merge cell by cell. A cell only one side changed takes that side's value.
- **Charts, images, comments, macros**: 3-way merge object by object. If only their branch changed a chart, you get their version. If both did, yours is kept and it's flagged as a conflict. Chart edits caused by cell changes (Excel caches plotted values inside the chart) don't count as edits.
- **Tables**: merged field by field. Their new column plus your new rows gives a table with both. Tables added on their branch come over, and if both branches added a `Table2`, theirs is renamed `Table3` and their formulas are updated to match.
- **Pivot tables**: a pivot, its data cache and the cached records merge as one bundle. A branch that only refreshed a pivot (same layout, new data) doesn't count as an edit, so your data change and their "switch to Average" merge cleanly. New pivots from their branch come over, sharing an existing cache when they used one.
- **Sheets**: new sheets on their branch come over whole, with their charts, tables and pivots. Renames and deletes merge too.
- **Named ranges**: merged by name.

Writers renumber a workbook's internal files on every save (add a chart to an early sheet and every later `chart1.xml` becomes `chart2.xml`). xlgit matches objects by what they are, like "the chart called Chart 1 on sheet Notes" or "table id 3", not by file name, so renumbering doesn't cause false conflicts.

Excel recalculates every formula and refreshes affected pivot tables when it opens the merged file.

## Setup

Install it (Python 3.9 or newer), then run `install` inside any repo holding Excel files:

```bash
pip install xlgit
xlgit install
```

That writes a `.gitattributes` entry and registers the diff and merge drivers in `.git/config`. The config part is per clone, so each collaborator runs `xlgit install` once. The `.gitattributes` part gets committed.

Don't want a package? `xlgit.py` is a single file. Copy it in, `pip install openpyxl lxml`, and run `python xlgit.py install`.

For the GitHub side, copy `.github/workflows/excel-diff.yml`, `requirements.txt` and `xlgit.py` into the repo root. Every PR that touches a workbook gets a cell-diff comment.

## Commands

```
xlgit textconv book.xlsx          # dump as text
xlgit diff old.xlsx new.xlsx      # list cell and object changes
xlgit diff a.xlsx b.xlsx --markdown
xlgit merge base.xlsx ours.xlsx theirs.xlsx
xlgit --version
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
- Until Excel refreshes a merged pivot, the numbers in its cells are the old ones. Excel does this on open, but tools that read the file without Excel (pandas, openpyxl) see the stale values.
- Slicers, timelines and tables linked to external data connections aren't merged. They're reported as conflicts so nothing disappears silently.
- If both branches added a chart to a sheet that had none, only yours is kept (flagged).
- Inserting a row shows up as many changed cells, because every cell below it moves. Fixing this is next on the [roadmap](ROADMAP.md).
- `.xls` (the old pre-2007 format) isn't supported. Save as `.xlsx`.

## Roadmap

Next up: testing against thousands of real workbooks, understanding inserted and moved rows, then code review for spreadsheets: showing what a change does to the numbers, tests that run on every pull request, and a linter. See [ROADMAP.md](ROADMAP.md).

## License

Apache License 2.0. See [LICENSE](LICENSE).

You can use, change and ship xlgit, including commercially. If you pass it on, modified or not, keep the [NOTICE](NOTICE) file and the copyright line at the top of `xlgit.py` with it. The license doesn't grant use of the xlgit name for your own version.

Created by Michael Fowler.
