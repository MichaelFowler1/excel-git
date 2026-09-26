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
