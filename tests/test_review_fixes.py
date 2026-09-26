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
