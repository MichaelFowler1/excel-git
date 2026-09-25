# Roadmap

Version control is the start. The goal is code review for spreadsheets.

A bad change to a spreadsheet can cost a company real money, and usually nobody catches it. Software teams catch bad code changes all the time, with tests and reviews that run on every change. xlgit aims to give Excel the same thing, for free, on GitHub.

The items are in rough order. Nothing here has a date. If you want to help with one, [open an issue](https://github.com/MichaelFowler1/excel-git/issues) and say which.

## Stage 1: make the core solid

### 1. Test against thousands of real workbooks

Done for the first round: `fuzz/merge_fuzz.py` gives each workbook random edits on two branches, merges them, and checks the result opens, lost nothing and has every edit from both sides. It has run over about 3,000 real workbooks so far (a sample of the Enron corpus plus the test files of other spreadsheet libraries), and found and fixed a dozen bugs before launch. Next: the rest of the Enron corpus (about 16,000 files), FUSE, and opening merged files in real Excel.

### 2. Understand inserted, deleted and moved rows

Diffs now report inserted, deleted and moved rows as rows (on real Enron revisions, a 6,849-line diff became 18 lines). Still to do: merges that follow rows too, so a row one branch inserted doesn't conflict with the other branch's edits below it; inserted columns; and matching table rows by an ID column, so two people can add rows to the same table without a conflict, the way a database handles it.

## Stage 2: code review for spreadsheets

### 3. Show what a change does to the numbers

A pull request comment today says "B7 changed from 5% to 6%". It should also say "this changes Net Profit by -$41,200 (-3.1%) and breaks the balance sheet check". To do that, xlgit would recalculate both versions of the workbook without Excel, using an open-source formula engine, and compare the results, not just the inputs.

The formula engine won't support every Excel function. Cells it can't calculate will be listed as such, not guessed.

### 4. Tests for spreadsheets

Write checks in a small config file, like "the balance sheet balances" or "Total equals the sum of its parts". A GitHub Action runs them on every pull request and can block a merge that breaks one. This reuses the recalculation from item 3.

### 5. A spreadsheet linter

Flag risky patterns on every pull request:

- numbers typed into formulas (`=B2*1.2`)
- a formula that differs from its neighbours in the same row or column
- broken references (`#REF!`) and circular references
- a range that stops one row short, such as a `SUM` that misses the last row

### 6. Diff and merge Power Query and VBA

Power Query and VBA are the code behind a lot of modern Excel work, and in git both are unreadable blobs today. xlgit should show them as text in diffs and merge them like code.

## Stage 3: for people who'll never type `git`

### 7. An Excel add-in

Buttons inside Excel to save a version, branch, compare and merge, with GitHub behind the scenes. It should work in Excel for the web too.

### 8. A visual conflict resolver

A web page that shows the sheet with the conflicting cells highlighted. Click a cell to take yours or theirs.

### 9. Plain-English change summaries

A short summary at the top of the pull request comment, like "Bob raised the Q2 rent assumption 10% and added a gas line; net cost up $4.2k", so reviewers don't have to read a table of 300 cells.
