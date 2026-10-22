"""Diffs report inserted, deleted and moved columns as columns, not as every
cell to the right of them changing."""
import xlsxwriter
from openpyxl.utils.cell import get_column_letter

import xlgit

BASE = [("Rent", 1000, 1000), ("Food", 300, 350), ("Power", 80, 90), ("Phone", 40, 40), ("Car", 200, 210)]


def ledger(path, cols, blank_after=None, sheet="Ledger", extra=None):
    """The budget laid out across columns: one column per item with Q1 and Q2
    below it, a =B2+B3 total under each, and a SUM column at the right, all
    written the way Excel would after the edit (formulas already shifted).
    extra: {cell: formula or value} written on top."""
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet(sheet)
    ws.write_column(0, 0, ["Item", "Q1", "Q2", "Total"])
    c = 1
    for i, (name, a, b) in enumerate(cols):
        ws.write(0, c, name)
        ws.write(1, c, a)
        ws.write(2, c, b)
        L = get_column_letter(c + 1)
        ws.write_formula(3, c, f"={L}2+{L}3")
        c += 1
        if blank_after == i:
            c += 1
    ws.write(0, c, "Total")
    last = get_column_letter(c)
    for r in (1, 2, 3):
        ws.write_formula(r, c, f"=SUM(B{r + 1}:{last}{r + 1})")
    for cell, v in (extra or {}).items():
        (ws.write_formula if isinstance(v, str) and v.startswith("=") else ws.write)(cell, v)
    wb.close()
    return str(path)


def changes(tmp_path, new_cols, **kw):
    a = ledger(tmp_path / "a.xlsx", BASE)
    b = ledger(tmp_path / "b.xlsx", new_cols, **kw)
    return [(k, c, o, n) for k, s, c, o, n in xlgit.diff_cells(xlgit.read_cells(a), xlgit.read_cells(b))]


def test_inserted_column(tmp_path):
    got = changes(tmp_path, BASE[:2] + [("Gas", 60, 70)] + BASE[2:])
    assert got[0] == ("column inserted", "column D", None, "D1: 'Gas', D2: 60, D3: 70, D4: =D2+D3")
    # the SUM column's ranges grew: three changed cells, one per row
    assert got[1:] == [("changed", "H2", "=SUM(B2:F2)", "=SUM(B2:G2)"),
                       ("changed", "H3", "=SUM(B3:F3)", "=SUM(B3:G3)"),
                       ("changed", "H4", "=SUM(B4:F4)", "=SUM(B4:G4)")]


def test_deleted_column(tmp_path):
    got = changes(tmp_path, BASE[:1] + BASE[2:])
    assert got[0] == ("column deleted", "column C", "C1: 'Food', C2: 300, C3: 350, C4: =C2+C3", None)
    assert got[1:] == [("changed", "F2", "=SUM(B2:F2)", "=SUM(B2:E2)"),
                       ("changed", "F3", "=SUM(B3:F3)", "=SUM(B3:E3)"),
                       ("changed", "F4", "=SUM(B4:F4)", "=SUM(B4:E4)")]


def test_moved_column(tmp_path):
    got = changes(tmp_path, BASE[1:4] + BASE[:1] + BASE[4:])
    assert got == [("column moved", "column B -> E", "B1: 'Rent', B2: 1000, B3: 1000, B4: =B2+B3", None)]


def test_inserted_column_and_an_edit_to_its_right(tmp_path):
    food = ("Food", 300, 999)
    got = changes(tmp_path, BASE[:1] + [food, ("Gas", 60, 70)] + BASE[2:])
    assert got[0][0] == "column inserted"
    assert ("changed", "C3", 350, 999) in got
    assert len(got) == 5  # plus the three SUM ranges


def test_empty_column_inserted(tmp_path):
    got = changes(tmp_path, BASE, blank_after=1)
    assert got[0] == ("columns inserted", "left of column E", None, "1 empty column(s)")
    assert len(got) == 4  # plus the three SUM ranges


def test_plain_edits_stay_cell_by_cell(tmp_path):
    got = changes(tmp_path, [("Rent", 1100, 1000)] + BASE[1:])
    assert got == [("changed", "B2", 1000, 1100)]


def test_text_diff_lines_up_with_row_lines(tmp_path, capsys):
    a = ledger(tmp_path / "a.xlsx", BASE)
    b = ledger(tmp_path / "b.xlsx", BASE[:1] + BASE[2:])
    xlgit.diff(a, b)
    out = capsys.readouterr().out
    assert "column deleted  Ledger column C  C1: 'Food', C2: 300, C3: 350, C4: =C2+C3" in out
    assert "changed         Ledger!F2  =SUM(B2:F2) -> =SUM(B2:E2)" in out


def test_markdown_diff_shows_the_column(tmp_path, capsys):
    a = ledger(tmp_path / "a.xlsx", BASE)
    b = ledger(tmp_path / "b.xlsx", BASE[:2] + [("Gas", 60, 70)] + BASE[2:])
    xlgit.diff(a, b, markdown=True)
    assert "| Ledger | column D | column inserted |  | D1: 'Gas', D2: 60, D3: 70, D4: =D2+D3 |" in capsys.readouterr().out


def page(tmp_path, new_cols, **kw):
    a = ledger(tmp_path / "a.xlsx", BASE)
    b = ledger(tmp_path / "b.xlsx", new_cols, **kw)
    return xlgit.html_report([("ledger.xlsx", xlgit.Package.open(a), xlgit.Package.open(b))])


def test_visual_diff_marks_columns(tmp_path):
    html = page(tmp_path, BASE[:2] + [("Gas", 60, 70)] + BASE[2:])
    assert '<th class="ins" title="inserted">D</th>' in html
    assert '<td class="cins">Gas</td>' in html and '<td class="f cins">=D2+D3</td>' in html
    assert "=SUM(B2:G2)" in html and "4 change(s)" in html
    html = page(tmp_path, BASE[:1] + BASE[2:])
    assert "column C deleted: C1: &#x27;Food&#x27;, C2: 300" in html
    html = page(tmp_path, BASE[1:4] + BASE[:1] + BASE[4:])
    assert '<th class="mov" title="moved from column B">E</th>' in html
