"""Diffs report inserted, deleted and moved rows as rows, not as every cell
below them changing."""
import xlsxwriter

import xlgit

BASE = [("Rent", 1000, 1000), ("Food", 300, 350), ("Power", 80, 90), ("Phone", 40, 40), ("Car", 200, 210)]


def budget(path, rows, blank_after=None):
    """A header, one row per item with a =B+C total, and a SUM row, all
    written the way Excel would after the edit (formulas already shifted)."""
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Budget")
    ws.write_row(0, 0, ["Item", "Q1", "Q2", "Total"])
    r = 1
    for i, (name, a, b) in enumerate(rows):
        ws.write(r, 0, name)
        ws.write(r, 1, a)
        ws.write(r, 2, b)
        ws.write_formula(r, 3, f"=B{r + 1}+C{r + 1}")
        r += 1
        if blank_after == i:
            r += 1
    ws.write(r, 0, "Total")
    ws.write_formula(r, 3, f"=SUM(D2:D{r})")
    wb.close()
    return str(path)


def changes(tmp_path, new_rows, **kw):
    a = budget(tmp_path / "a.xlsx", BASE)
    b = budget(tmp_path / "b.xlsx", new_rows, **kw)
    return [(k, c, o, n) for k, s, c, o, n in xlgit.diff_cells(xlgit.read_cells(a), xlgit.read_cells(b))]


def test_inserted_row(tmp_path):
    got = changes(tmp_path, BASE[:2] + [("Gas", 60, 70)] + BASE[2:])
    assert got == [("row inserted", "row 4", None, "A: 'Gas', B: 60, C: 70, D: =B4+C4"),
                   ("changed", "D8", "=SUM(D2:D6)", "=SUM(D2:D7)")]


def test_deleted_row(tmp_path):
    got = changes(tmp_path, BASE[:1] + BASE[2:])
    assert got[0] == ("row deleted", "row 3", "A: 'Food', B: 300, C: 350, D: =B3+C3", None)
    assert got[1:] == [("changed", "D6", "=SUM(D2:D6)", "=SUM(D2:D5)")]


def test_moved_row(tmp_path):
    got = changes(tmp_path, BASE[1:4] + BASE[:1] + BASE[4:])
    assert got == [("row moved", "row 2 -> 5", "A: 'Rent', B: 1000, C: 1000, D: =B2+C2", None)]


def test_inserted_row_and_an_edit_below_it(tmp_path):
    food = ("Food", 300, 999)
    got = changes(tmp_path, BASE[:1] + [food, ("Gas", 60, 70)] + BASE[2:])
    assert ("changed", "C3", 350, 999) in got
    assert got[1][0] == "row inserted" and len(got) == 3


def test_empty_row_inserted(tmp_path):
    got = changes(tmp_path, BASE, blank_after=1)
    assert ("rows inserted", "above row 5", None, "1 empty row(s)") in got
    assert len(got) == 2  # plus the SUM row's range growing


def test_plain_edits_stay_cell_by_cell(tmp_path):
    got = changes(tmp_path, [("Rent", 1100, 1000)] + BASE[1:])
    assert got == [("changed", "B2", 1000, 1100)]


def page(tmp_path, new_rows, **kw):
    a = budget(tmp_path / "a.xlsx", BASE)
    b = budget(tmp_path / "b.xlsx", new_rows, **kw)
    return xlgit.html_report([("budget.xlsx", xlgit.Package.open(a), xlgit.Package.open(b))])


def test_visual_diff_marks_rows_and_cells(tmp_path):
    html = page(tmp_path, BASE[:1] + [("Food", 300, 999), ("Gas", 60, 70)] + BASE[2:])
    assert '<tr class="ins"><td class="rn" title="inserted">4</td>' in html
    assert '<span class="was">350</span>999' in html
    assert "=SUM(D2:D7)" in html and "3 change(s)" in html


def test_visual_diff_shows_deleted_and_moved_rows(tmp_path):
    assert '<tr class="dele"><td class="rn" title="deleted">−3</td>' in page(tmp_path, BASE[:1] + BASE[2:])
    assert 'title="moved from row 2">5</td>' in page(tmp_path, BASE[1:4] + BASE[:1] + BASE[4:])


def test_visual_diff_never_runs_cell_text(tmp_path):
    html = page(tmp_path, [("<script>alert(1)</script>", 1000, 1000)] + BASE[1:])
    assert "<script>" not in html and "&lt;script&gt;alert(1)&lt;/script&gt;" in html
