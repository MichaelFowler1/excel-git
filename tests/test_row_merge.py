"""Merges follow inserted and deleted rows: when one branch inserts or
deletes rows and the other edits cells, the edits land on the right rows and
their formulas point at the right cells."""
import pytest

import xlgit
from test_rows import BASE, budget

GAS = ("Gas", 60, 70)


def merge(tmp_path, ours_rows, theirs_rows):
    paths = [budget(tmp_path / f"{n}.xlsx", rows) for n, rows in
             (("base", BASE), ("ours", ours_rows), ("theirs", theirs_rows))]
    rc = xlgit.merge(*paths)
    return rc, xlgit.read_cells(paths[1])["Budget"]


def with_car(rows, q1):
    return [(n, q1 if n == "Car" else a, b) for n, a, b in rows]


@pytest.mark.parametrize("who_inserts", ["ours", "theirs"])
def test_edit_lands_on_the_moved_row(tmp_path, who_inserts):
    inserted = BASE[:2] + [GAS] + BASE[2:]
    edited = with_car(BASE, 250)
    ours, theirs = (inserted, edited) if who_inserts == "ours" else (edited, inserted)
    rc, cells = merge(tmp_path, ours, theirs)
    assert rc == 0
    assert cells["A4"] == "Gas" and cells["A7"] == "Car"
    assert cells["B7"] == 250  # Car was row 6 in base, row 7 after the insert
    assert cells["D7"] == "=B7+C7" and cells["D8"] == "=SUM(D2:D7)"
    assert "B8" not in cells  # nothing written where Car used to be


@pytest.mark.parametrize("who_inserts", ["ours", "theirs"])
def test_their_formula_is_renumbered(tmp_path, who_inserts):
    inserted = BASE[:2] + [GAS] + BASE[2:]
    ours, theirs = (inserted, BASE) if who_inserts == "ours" else (BASE, inserted)
    paths = [budget(tmp_path / f"{n}.xlsx", r) for n, r in (("base", BASE), ("ours", ours), ("theirs", theirs))]
    # The branch that didn't insert adds a formula pointing at Car's total.
    other = paths[2] if who_inserts == "ours" else paths[1]
    add_cell(other, "F6", "=D6*2")
    assert xlgit.merge(*paths) == 0
    cells = xlgit.read_cells(paths[1])["Budget"]
    assert cells["F7"] == "=D7*2"


def test_edit_to_a_row_you_deleted_conflicts(tmp_path):
    rc, cells = merge(tmp_path, BASE[:1] + BASE[2:], [(n, 999 if n == "Food" else a, b) for n, a, b in BASE])
    assert rc == 1
    assert "Food" not in cells.values()


@pytest.mark.parametrize("who_inserts", ["ours", "theirs"])
def test_clash_keeps_your_value_on_the_right_row(tmp_path, who_inserts, capsys):
    inserted = BASE[:2] + [GAS] + BASE[2:]
    if who_inserts == "ours":
        ours, theirs = with_car(inserted, 111), with_car(BASE, 222)
    else:
        ours, theirs = with_car(BASE, 111), with_car(inserted, 222)
    rc, cells = merge(tmp_path, ours, theirs)
    assert rc == 1
    assert cells["A4"] == "Gas" and cells["B7"] == 111
    assert "Budget B7: yours 111, theirs 222" in capsys.readouterr().err


def add_cell(path, ref, formula):
    import re
    import zipfile
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    xml = parts["xl/worksheets/sheet1.xml"].decode()
    row = re.search(rf'<row r="{ref[1:]}"[^>]*>.*?</row>', xml).group(0)
    parts["xl/worksheets/sheet1.xml"] = xml.replace(
        row, row.replace("</row>", f'<c r="{ref}"><f>{formula[1:]}</f></c></row>')).encode()
    with zipfile.ZipFile(path, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)
