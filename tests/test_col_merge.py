"""Merges follow inserted and deleted columns: when one branch inserts or
deletes columns and the other edits cells, the edits land in the right
columns and their formulas point at the right cells."""
import pytest
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter

import xlgit
from test_cols import BASE, ledger
from test_rows import BASE as ROWS_BASE, budget

GAS = ("Gas", 60, 70)
INSERTED = BASE[:2] + [GAS] + BASE[2:]  # Gas becomes column D; Power D -> E, Phone E -> F, Car F -> G
DELETED = BASE[:1] + BASE[2:]  # Food (column C) gone; Power D -> C, Phone E -> D, Car F -> E


def merge(tmp_path, ours, theirs, base=BASE, **kw):
    paths = [ledger(tmp_path / f"{n}.xlsx", cols, **kw) for n, cols in (("base", base), ("ours", ours), ("theirs", theirs))]
    rc = xlgit.merge(*paths)
    return rc, xlgit.read_cells(paths[1])["Ledger"]


def with_car(cols, q1):
    return [(n, q1 if n == "Car" else a, b) for n, a, b in cols]


@pytest.mark.parametrize("who_inserts", ["ours", "theirs"])
def test_edit_lands_in_the_moved_column(tmp_path, who_inserts):
    edited = with_car(BASE, 250)
    ours, theirs = (INSERTED, edited) if who_inserts == "ours" else (edited, INSERTED)
    rc, cells = merge(tmp_path, ours, theirs)
    assert rc == 0
    assert cells["D1"] == "Gas" and cells["G1"] == "Car"
    assert cells["G2"] == 250  # Car was column F in base, G after the insert
    assert cells["G4"] == "=G2+G3" and cells["H2"] == "=SUM(B2:G2)"
    assert cells["F2"] == 40  # nothing written where Car used to be


@pytest.mark.parametrize("who_inserts", ["ours", "theirs"])
def test_their_formula_is_renumbered(tmp_path, who_inserts):
    ours, theirs = (INSERTED, BASE) if who_inserts == "ours" else (BASE, INSERTED)
    extra = {"F6": "=F4*2"}  # the branch that didn't insert points at Car's total
    paths = [ledger(tmp_path / f"{n}.xlsx", c, extra=x) for n, c, x in
             (("base", BASE, None), ("ours", ours, extra if who_inserts == "theirs" else None),
              ("theirs", theirs, extra if who_inserts == "ours" else None))]
    assert xlgit.merge(*paths) == 0
    cells = xlgit.read_cells(paths[1])["Ledger"]
    assert cells["G6"] == "=G4*2" and "F6" not in cells


def test_edit_in_another_column_merges_with_a_deletion(tmp_path):
    rc, cells = merge(tmp_path, with_car(BASE, 250), DELETED)
    assert rc == 0
    assert "Food" not in cells.values()
    assert cells["E1"] == "Car" and cells["E2"] == 250 and cells["F2"] == "=SUM(B2:E2)"


def test_edit_to_a_column_they_deleted_conflicts(tmp_path, capsys):
    rc, cells = merge(tmp_path, [(n, 999 if n == "Food" else a, b) for n, a, b in BASE], DELETED)
    assert rc == 1
    assert "Food" not in cells.values()
    assert "C2 (column deleted)" in capsys.readouterr().err


def test_clearing_a_cell_of_a_column_they_deleted_agrees(tmp_path):
    base = ledger(tmp_path / "base.xlsx", BASE)
    ours = ledger(tmp_path / "ours.xlsx", BASE)
    theirs = ledger(tmp_path / "theirs.xlsx", DELETED)
    wb = xlsxwriter.Workbook(str(ours))  # rewrite ours with C3 blank
    ws = wb.add_worksheet("Ledger")
    for coord, v in xlgit.read_cells(base)["Ledger"].items():
        if coord != "C3":
            (ws.write_formula if isinstance(v, str) and v.startswith("=") else ws.write)(coord, v)
    wb.close()
    assert xlgit.merge(base, ours, theirs) == 0


@pytest.mark.parametrize("who_inserts", ["ours", "theirs"])
def test_clash_keeps_your_value_in_the_right_column(tmp_path, who_inserts, capsys):
    if who_inserts == "ours":
        ours, theirs = with_car(INSERTED, 111), with_car(BASE, 222)
    else:
        ours, theirs = with_car(BASE, 111), with_car(INSERTED, 222)
    rc, cells = merge(tmp_path, ours, theirs)
    assert rc == 1
    assert cells["D1"] == "Gas" and cells["G2"] == 111
    assert "Ledger G2: yours 111, theirs 222" in capsys.readouterr().err


def test_both_branches_changing_columns_merges_cell_by_cell(tmp_path, capsys):
    rc, cells = merge(tmp_path, INSERTED, DELETED)
    err = capsys.readouterr().err
    assert "both branches inserted or deleted columns on 'Ledger'; merged cell by cell" in err
    assert cells["D1"] == "Gas"  # ours kept where the two readings clash


def test_rows_on_one_branch_and_columns_on_the_other_merges_cell_by_cell(tmp_path, capsys):
    base = ledger(tmp_path / "base.xlsx", BASE, extra={"A7": "note"})
    ours = ledger(tmp_path / "ours.xlsx", BASE, extra={"A8": "note"})  # a row inserted above the note
    theirs = ledger(tmp_path / "theirs.xlsx", INSERTED, extra={"A7": "note"})
    xlgit.merge(base, ours, theirs)
    assert "both branches inserted or deleted columns and rows on 'Ledger'; merged cell by cell" \
        in capsys.readouterr().err


def test_rows_and_columns_on_one_branch_merges_cell_by_cell(tmp_path, capsys):
    base = ledger(tmp_path / "base.xlsx", BASE, extra={"A7": "note"})
    ours = ledger(tmp_path / "ours.xlsx", with_car(BASE, 250), extra={"A7": "note"})
    theirs = ledger(tmp_path / "theirs.xlsx", INSERTED, extra={"A8": "note"})
    xlgit.merge(base, ours, theirs)
    assert "their branch inserted or deleted both rows and columns on 'Ledger'; merged cell by cell" \
        in capsys.readouterr().err


def test_row_merge_still_works(tmp_path):
    """The row path shares the column code: the budget merge of test_row_merge
    still lands edits on the moved row."""
    paths = [budget(tmp_path / f"{n}.xlsx", r) for n, r in
             (("base", ROWS_BASE), ("ours", [(n, 250 if n == "Car" else a, b) for n, a, b in ROWS_BASE]),
              ("theirs", ROWS_BASE[:2] + [("Gas", 60, 70)] + ROWS_BASE[2:]))]
    assert xlgit.merge(*paths) == 0
    cells = xlgit.read_cells(paths[1])["Budget"]
    assert cells["B7"] == 250 and cells["D8"] == "=SUM(D2:D7)"


# --- every reference form, on the shifted sheet and from another sheet ---

FORMS = {  # written on the Ledger sheet by the branch that didn't change columns
    "A10": "=C2", "A11": "=$C$2", "A12": "=C$2", "A13": "=$C2",
    "A14": "=SUM(B2:D2)", "A15": "=SUM(C:D)", "A16": "=SUM(C:C)", "A17": "=Rent+1",
    "A18": "='Other Sheet'!C3", "A19": "=SUM(Other!B1:D1)",
}
AFTER_INSERT = {  # Gas inserted at D (C and everything right of it moves on)... C stays, D -> E
    "A10": "=C2", "A11": "=$C$2", "A12": "=C$2", "A13": "=$C2",
    "A14": "=SUM(B2:E2)", "A15": "=SUM(C:E)", "A16": "=SUM(C:C)", "A17": "=Rent+1",
    "A18": "='Other Sheet'!C3", "A19": "=SUM(Other!B1:D1)",
}
AFTER_DELETE = {  # Food (column C) deleted
    "A10": "=#REF!", "A11": "=#REF!", "A12": "=#REF!", "A13": "=#REF!",
    "A14": "=SUM(B2:C2)", "A15": "=SUM(C:C)", "A16": "=SUM(#REF!)", "A17": "=Rent+1",
    "A18": "='Other Sheet'!C3", "A19": "=SUM(Other!B1:D1)",
}


def workbook(path, cols, forms=None, other_forms=None, name="=Ledger!$C$2"):
    """A Ledger laid out by `cols`, an 'Other Sheet' and 'Other' the Ledger
    points at, a formula on Other pointing back at Ledger, and a defined
    name Rent pointing at Ledger."""
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Ledger")
    ws.write_column(0, 0, ["Item", "Q1", "Q2", "Total"])
    for c, (item, a, b) in enumerate(cols, start=1):
        ws.write(0, c, item)
        ws.write(1, c, a)
        ws.write(2, c, b)
        L = get_column_letter(c + 1)
        ws.write_formula(3, c, f"={L}2+{L}3")
    for cell, f in (forms or {}).items():
        ws.write_formula(cell, f)
    other = wb.add_worksheet("Other Sheet")
    other.write("C3", 7)
    o2 = wb.add_worksheet("Other")
    o2.write_row(0, 0, [1, 2, 3, 4])
    for cell, f in (other_forms or {}).items():
        o2.write_formula(cell, f)
    wb.define_name("Rent", name)
    wb.close()
    return str(path)



@pytest.mark.parametrize("who_changes", ["ours", "theirs"])
@pytest.mark.parametrize("change,expected,other_expected,name_after", [
    (INSERTED, AFTER_INSERT, "=SUM(Ledger!C2:E2)", "=Ledger!$C$2"),
    (DELETED, AFTER_DELETE, "=SUM(Ledger!C2:C2)", "=Ledger!#REF!"),  # Excel writes #REF! into the name on that branch
])
def test_every_reference_form(tmp_path, who_changes, change, expected, other_expected, name_after):
    other = {"A3": "=SUM(Ledger!C2:D2)"}  # from another sheet into the shifted one
    base = workbook(tmp_path / "base.xlsx", BASE)
    edited, moved = dict(forms=FORMS, other_forms=other), dict(name=name_after)
    ours = workbook(tmp_path / "ours.xlsx", BASE if who_changes == "theirs" else change,
                    **(edited if who_changes == "theirs" else moved))
    theirs = workbook(tmp_path / "theirs.xlsx", change if who_changes == "theirs" else BASE,
                      **(edited if who_changes == "ours" else moved))
    rc = xlgit.merge(base, ours, theirs)
    cells = xlgit.read_cells(ours)
    got = {k: cells["Ledger"].get(k) for k in FORMS}
    assert got == expected
    assert cells["Other"]["A3"] == other_expected
    assert rc == 0
    wb = load_workbook(ours)
    assert wb.defined_names["Rent"].attr_text == name_after.lstrip("=")


def test_deleted_column_breaks_the_name_that_pointed_at_it(tmp_path):
    """Their branch deleted the column a name points at; Excel writes
    Ledger!#REF! into the name and so does the merge."""
    base = workbook(tmp_path / "base.xlsx", BASE)
    ours = workbook(tmp_path / "ours.xlsx", BASE, forms={"A10": "=Rent*2"})
    theirs = workbook(tmp_path / "theirs.xlsx", DELETED, name="=Ledger!#REF!")
    assert xlgit.merge(base, ours, theirs) == 0
    assert load_workbook(ours).defined_names["Rent"].attr_text == "Ledger!#REF!"
    assert xlgit.read_cells(ours)["Ledger"]["A10"] == "=Rent*2"
