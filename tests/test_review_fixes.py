"""Regression tests for bugs found in the September 2026 code review."""
import xlgit


def test_table_rename_leaves_quoted_sheet_names_and_text_alone():
    f = "=SUM(Table2[Y])+'Table2 notes'!A1&\"Table2 total\"+COUNT(Table2)"
    assert xlgit.rename_refs(f, {"Table2": "Table3"}) == \
        "=SUM(Table3[Y])+'Table2 notes'!A1&\"Table2 total\"+COUNT(Table3)"


def test_table_rename_handles_escaped_quotes():
    f = "='It''s Table2'!A1&\"say \"\"Table2\"\"\"&Table2[X]"
    assert xlgit.rename_refs(f, {"Table2": "Table3"}) == \
        "='It''s Table2'!A1&\"say \"\"Table2\"\"\"&Table3[X]"


def test_array_formulas_follow_a_renamed_table(tmp_path):
    from test_tables_pivots import repo_with
    ours_t = dict(sheet="Sales", ref="E1:F3", header=["K", "V"], data=[["a", 1], ["b", 2]])
    theirs_t = dict(sheet="Summary", ref="H1:I3", header=["X", "Y"], data=[["c", 3], ["d", 4]])
    r = repo_with(tmp_path, dict(pivots=()), dict(pivots=(), tables=[ours_t]),
                  dict(pivots=(), tables=[theirs_t], formulas={"Summary": {"K2": "{=SUM(Table2[Y]*2)}"}}))
    result = r.merge("ours", "theirs")
    assert result.returncode == 0, result.stderr
    k2 = xlgit.read_cells(str(r.book))["Summary"]["K2"]
    assert isinstance(k2, xlgit.ArrayF) and k2.text == "=SUM(Table3[Y]*2)"


def test_long_whole_numbers_are_compared_exactly():
    assert not xlgit.same(1234567890123456, 1234567890123457)
    assert xlgit.same(1234567890123456, 1234567890123456)
    # Excel keeps 15 significant digits, so its copy of the same number still matches.
    assert xlgit.same(1234567890123456, 1.23456789012346e15)
    assert xlgit.same(0.1 + 0.2, 0.3)  # float noise between writers still isn't a change


def test_a_changed_long_id_survives_the_merge(tmp_path):
    from test_cell_types import book
    paths = [str(tmp_path / f"{n}.xlsx") for n in ("base", "ours", "theirs")]
    for p, ident in zip(paths, (1234567890123456, 1234567890123456, 1234567890123457)):
        book(p, {"A1": "Account", "B1": ident, "A2": "Owner", "B2": "Ann"})
    assert xlgit.merge(*paths) == 0
    assert xlgit.read_cells(paths[1])["Data"]["B1"] == 1234567890123457


def test_clearing_a_cell_in_a_row_the_other_side_deleted_is_no_clash(tmp_path):
    from test_rows import BASE, budget
    for who_deletes in ("ours", "theirs"):
        d = tmp_path / who_deletes
        d.mkdir()
        deleted = BASE[:1] + BASE[2:]                                     # Food row gone
        cleared = [(n, None if n == "Food" else a, b) for n, a, b in BASE]  # Food's Q1 emptied
        ours, theirs = (deleted, cleared) if who_deletes == "ours" else (cleared, deleted)
        paths = [budget(d / f"{n}.xlsx", rows) for n, rows in (("base", BASE), ("ours", ours), ("theirs", theirs))]
        assert xlgit.merge(*paths) == 0, who_deletes
        cells = xlgit.read_cells(paths[1])["Budget"]
        assert "Food" not in cells.values() and cells["A3"] == "Power"
