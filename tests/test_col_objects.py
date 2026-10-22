"""Charts, tables and named ranges survive a merge where one branch inserted
a column, through real `git merge`; a table both sides changed gets the
right range; a chart the other branch edited is flagged, not fixed."""
import pytest
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter

import xlgit
from test_cols import BASE
from test_merge import Repo, assert_opens_clean, charts, keep

GAS = ("Gas", 60, 70)
INSERTED = BASE[:2] + [GAS] + BASE[2:]  # Gas at D; Power, Phone, Car move one column right
PEOPLE = [["Ann", "rent", 1, "x"], ["Ben", "food", 2, "y"], ["Cy", "car", 3, "z"]]


def office(path, cols=BASE, *, car_q1=200, title="Spend", people=PEOPLE, table_cols=("Who", "What", "N", "Note"),
           chart_items=None):
    """A Ledger across columns with a chart over its Q1 row (all items, or
    the first `chart_items`), a table at A6 (its columns span the insert
    point D) and a name on it."""
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Ledger")
    ws.write_column(0, 0, ["Item", "Q1", "Q2", "Total"])
    for c, (item, a, b) in enumerate(cols, start=1):
        ws.write(0, c, item)
        ws.write(1, c, car_q1 if item == "Car" else a)
        ws.write(2, c, b)
        L = get_column_letter(c + 1)
        ws.write_formula(3, c, f"={L}2+{L}3")
    last = get_column_letter((chart_items or len(cols)) + 1)
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"categories": f"=Ledger!$B$1:${last}$1", "values": f"=Ledger!$B$2:${last}$2"})
    ch.set_title({"name": title})
    ws.insert_chart("B12", ch)
    ws.add_table(5, 0, 5 + len(people), len(table_cols) - 1,
                 {"name": "People", "columns": [{"header": h} for h in table_cols], "data": people})
    wb.define_name("CarQ1", f"=Ledger!${get_column_letter([c[0] for c in cols].index('Car') + 2)}$2")
    wb.close()


class Office(Repo):
    def commit(self, branch, from_branch, **variant):
        self.git("checkout", "-q", from_branch, check=branch != from_branch)
        if branch != from_branch:
            self.git("checkout", "-q", "-b", branch)
        office(self.book, **variant)
        self.git("add", "-A")
        self.git("commit", "-q", "-m", branch)


def table(path):
    return load_workbook(path)["Ledger"].tables["People"]


@pytest.mark.parametrize("ours,theirs", [("alice", "bob"), ("bob", "alice")])
def test_column_insert_keeps_chart_table_and_name(tmp_path, ours, theirs):
    r = Office(tmp_path)
    r.commit("main", "main")
    r.commit("alice", "main", car_q1=250)
    r.commit("bob", "main", cols=INSERTED)
    result = r.merge(ours, theirs)
    assert result.returncode == 0, result.stdout + result.stderr
    keep(r.book, f"merged_cols_{ours}_{theirs}.xlsx")
    assert_opens_clean(r.book)
    wb = load_workbook(r.book)
    ws = wb["Ledger"]
    assert ws["D1"].value == "Gas" and ws["G1"].value == "Car"
    assert ws["G2"].value == 250 and ws["F2"].value == 40  # alice's edit followed Car
    assert ws["G4"].value == "=G2+G3"
    assert charts(r.book, "Ledger") == ["Spend"]
    with open(r.book, "rb") as f:
        pkg = xlgit.Package(f.read())
    chart_xml = b"".join(d for n, d in pkg.parts.items() if n.startswith("xl/charts/chart"))
    assert b"Ledger!$B$2:$G$2" in chart_xml  # bob's Excel widened the series
    assert wb.defined_names["CarQ1"].attr_text == "Ledger!$G$2"
    assert table(r.book).ref == "A6:D9"  # the table sits left of the insert... until Excel widens it:


@pytest.mark.parametrize("ours,theirs", [("alice", "bob"), ("bob", "alice")])
def test_table_extended_on_one_branch_and_widened_on_the_other(tmp_path, ours, theirs):
    """Bob inserts a column inside the table (Excel adds a table column);
    Alice adds a row to it. The merged table has both: five columns, four
    rows, with Alice's row landing in the shifted columns."""
    r = Office(tmp_path)
    r.commit("main", "main")
    r.commit("alice", "main", people=PEOPLE + [["Di", "gas", 4, "w"]])
    r.commit("bob", "main", cols=INSERTED, table_cols=("Who", "What", "N", "Column1", "Note"),
             people=[p[:3] + [None] + p[3:] for p in PEOPLE])
    result = r.merge(ours, theirs)
    assert result.returncode == 0, result.stdout + result.stderr
    keep(r.book, f"merged_table_{ours}_{theirs}.xlsx")
    assert_opens_clean(r.book)
    t = table(r.book)
    assert t.ref == "A6:E10" and [c.name for c in t.tableColumns] == ["Who", "What", "N", "Column1", "Note"]
    ws = load_workbook(r.book)["Ledger"]
    assert [ws[f"{c}10"].value for c in "ABCDE"] == ["Di", "gas", 4, None, "w"]


def test_chart_edited_on_the_other_branch_is_flagged(tmp_path):
    """Alice retitled the chart; Bob inserted a column to the right of its
    ranges, so his copy of the chart is unchanged. Her chart comes over, but
    its ranges were written before the column moved, so the merge says so
    instead of staying quiet. (A chart whose ranges the insert did shift is
    changed on both branches, and conflicts the usual way.)"""
    r = Office(tmp_path)
    r.commit("main", "main", chart_items=2)
    r.commit("alice", "main", chart_items=2, title="Spend 2026")
    r.commit("bob", "main", chart_items=2, cols=INSERTED)
    result = r.merge("bob", "alice")
    assert result.returncode != 0
    assert_opens_clean(r.book)
    assert charts(r.book, "Ledger") == ["Spend 2026"]
    rows = [[c.value for c in row] for row in load_workbook(r.book)["_merge_conflicts"].iter_rows()]
    assert any(row[0] == "(chart)" and "columns on 'Ledger' moved" in str(row[4]) for row in rows[1:]), rows
