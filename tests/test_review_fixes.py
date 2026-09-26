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


def _shared_group_book(path, extra=None, b1="A1*2"):
    """B1:B3 as one shared formula group (=A1*2 filled down), the way Excel
    stores it. extra: {cell: number} added in column C."""
    import re
    import zipfile
    import xlsxwriter
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Data")
    for r in range(3):
        ws.write_number(r, 0, r + 1)
        ws.write_formula(r, 1, f"=A{r + 1}*2")
    for ref, v in (extra or {}).items():
        ws.write_number(ref, v)
    wb.close()
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    xml = parts["xl/worksheets/sheet1.xml"].decode()
    xml = re.sub(r'<c r="B1"([^>]*)><f>[^<]*</f>', rf'<c r="B1"\1><f t="shared" ref="B1:B3" si="0">{b1}</f>', xml)
    for ref in ("B2", "B3"):
        xml = re.sub(rf'<c r="{ref}"([^>]*)><f>[^<]*</f>', rf'<c r="{ref}"\1><f t="shared" si="0"/>', xml)
    parts["xl/worksheets/sheet1.xml"] = xml.encode()
    with zipfile.ZipFile(path, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)
    return str(path)


def _group_intact(path):
    """Every cell pointing at a shared formula group still has the group's formula cell."""
    pkg = xlgit.Package.open(path)
    root = pkg.xml("xl/worksheets/sheet1.xml")
    masters, children = set(), set()
    for f in root.iter(xlgit.m("f")):
        if f.get("t") == "shared":
            (masters if (f.text or "").strip() else children).add(f.get("si"))
    return children <= masters


def test_a_shared_formula_group_is_never_left_half_spelled_out(tmp_path, monkeypatch):
    class Stubborn(xlgit.Translator):  # this one cell's formula can't be worked out
        def translate_formula(self, dest=None, row_delta=0, col_delta=0):
            if dest == "B3":
                raise ValueError("can't translate")
            return super().translate_formula(dest, row_delta, col_delta)
    monkeypatch.setattr(xlgit, "Translator", Stubborn)

    # Their edit elsewhere on the sheet: the group has to stay whole.
    paths = [_shared_group_book(tmp_path / f"{n}.xlsx", extra) for n, extra in
             (("base", None), ("ours", None), ("theirs", {"C1": 7}))]
    assert xlgit.merge(*paths) == 0
    assert _group_intact(paths[1])
    assert xlgit.read_cells(paths[1])["Data"]["C1"] == 7

    # Their edit to the group's formula cell itself: reported, not written half-way.
    d = tmp_path / "master"
    d.mkdir()
    paths = [_shared_group_book(d / f"{n}.xlsx", b1=f) for n, f in
             (("base", "A1*2"), ("ours", "A1*2"), ("theirs", "A1*3"))]
    assert xlgit.merge(*paths) == 1
    assert _group_intact(paths[1])


def test_a_result_that_cant_be_saved_says_so_plainly(tmp_path, capsys):
    import os
    import stat
    from test_cell_types import book
    paths = [book(tmp_path / f"{n}.xlsx", {"A1": v}) or str(tmp_path / f"{n}.xlsx")
             for n, v in (("base", "a"), ("ours", "a"), ("theirs", "b"))]
    os.chmod(paths[1], stat.S_IREAD)  # stands in for a file something else holds open
    try:
        assert xlgit.merge(*paths) == 1
    finally:
        os.chmod(paths[1], stat.S_IREAD | stat.S_IWRITE)
    err = capsys.readouterr().err
    assert "couldn't save the result" in err and "close it and run the merge again" in err
