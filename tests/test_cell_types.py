"""Cell values survive a merge with their exact type and value.

Each case here was found by fuzzing the merge against real workbooks
(fuzz/merge_fuzz.py). The merge runs in-process on three files; no git.
"""
import re
import zipfile

import xlsxwriter

import xlgit

SHEET = "xl/worksheets/sheet1.xml"


def book(path, cells, fmt=None):
    """cells: {ref: value}; strings are written as text, never as formulas."""
    wb = xlsxwriter.Workbook(str(path), {"strings_to_formulas": False, "strings_to_numbers": False})
    ws = wb.add_worksheet("Data")
    date = wb.add_format({"num_format": "yyyy-mm-dd hh:mm"})
    for ref, v in cells.items():
        if isinstance(v, bool):
            ws.write_boolean(ref, v)
        elif isinstance(v, (int, float)):
            ws.write_number(ref, v, date if fmt and ref in fmt else None)
        else:
            ws.write_string(ref, v)
    wb.close()
    return str(path)


def merge3(tmp_path, base, ours, theirs, **kw):
    paths = [book(tmp_path / f"{n}.xlsx", cells, **kw) for n, cells in
             (("base", base), ("ours", ours), ("theirs", theirs))]
    return paths, xlgit.merge(*paths)


def raw_cell(path, ref):
    xml = zipfile.ZipFile(path).read(SHEET).decode()
    return re.search(rf'<c r="{ref}"[^>]*?(/>|>.*?</c>)', xml).group(0)


def test_text_that_looks_like_a_formula_stays_text(tmp_path):
    (b, o, t), rc = merge3(tmp_path, {"A1": "x", "B1": 1}, {"A1": "x", "B1": 2},
                           {"A1": "=== Total ===", "B1": 1})
    assert rc == 0
    assert "<f>" not in raw_cell(o, "A1")
    cells = xlgit.read_cells(o)["Data"]
    assert cells["A1"] == xlgit.Text("=== Total ===") and cells["B1"] == 2


def test_text_that_looks_like_an_error_stays_text(tmp_path):
    (b, o, t), rc = merge3(tmp_path, {"A1": "x", "B1": 1}, {"A1": "x", "B1": 2}, {"A1": "#N/A", "B1": 1})
    assert rc == 0
    assert 't="e"' not in raw_cell(o, "A1")
    assert xlgit.read_cells(o)["Data"]["A1"] == xlgit.Text("#N/A")


def test_text_and_formula_that_read_alike_differ(tmp_path):
    assert xlgit.Text("=A1") != "=A1" and "=A1" != xlgit.Text("=A1")


def test_true_is_not_one(tmp_path):
    (b, o, t), rc = merge3(tmp_path, {"A1": 1, "B1": 1}, {"A1": 1, "B1": 2}, {"A1": True, "B1": 1})
    assert rc == 0
    assert xlgit.read_cells(o)["Data"]["A1"] is True


def test_numbers_in_date_cells_keep_every_digit(tmp_path):
    # Going through a datetime rounds to the microsecond, and a number past
    # year 9999 has no datetime at all; the merge must write the number back.
    dates = {"A1", "A2", "A3"}
    base = {"A1": 45000.5, "A2": 45000.5, "A3": 45000.5, "B1": 1}
    theirs = {"A1": 1e-07, "A2": 123456789.123, "A3": 45000.123456789123, "B1": 1}
    (b, o, t), rc = merge3(tmp_path, base, dict(base, B1=2), theirs, fmt=dates)
    assert rc == 0
    for ref in ("A1", "A2", "A3"):
        v = re.search(r"<v>(.*?)</v>", raw_cell(o, ref)).group(1)
        assert float(v) == theirs[ref], ref


def test_rows_and_cells_without_references(tmp_path):
    # r="..." is optional on rows and cells; some writers leave it out.
    (b, o, t), _ = merge3(tmp_path, {"A1": 1, "B1": 1, "A2": 1}, {"A1": 1, "B1": 1, "A2": 1},
                          {"A1": 1, "B1": 1, "A2": 1})
    for p, value in ((b, None), (o, "2"), (t, "3")):
        strip_refs(p, value)
    assert xlgit.read_cells(t)["Data"] == {"A1": 1, "B1": 3, "A2": 1}
    assert xlgit.merge(b, o, t) == 0
    assert xlgit.read_cells(o)["Data"] == {"A1": 2, "B1": 3, "A2": 1}


def strip_refs(path, value):
    """Drop every r= from rows and cells; set A1 (ours) or B1 (theirs) to value."""
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    xml = re.sub(r'(<(?:row|c)) r="[A-Z]*\d+"', r"\1", parts[SHEET].decode())
    target = {"2": 0, "3": 1}.get(value)
    n = iter(range(99))
    xml = re.sub(r"<c[ >].*?</c>", lambda mt: f"<c><v>{value}</v></c>" if next(n) == target else mt.group(0), xml)
    parts[SHEET] = xml.encode()
    with zipfile.ZipFile(path, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)


def test_sheet_with_no_part(tmp_path):
    # Excel lists old macro sheets with an empty relationship id.
    paths, _ = merge3(tmp_path, {"A1": 1}, {"A1": 1}, {"A1": 2})
    for p in paths:
        with zipfile.ZipFile(p) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
            b"</sheets>", b'<sheet name="Macro1" sheetId="9" state="veryHidden" r:id=""/></sheets>')
        with zipfile.ZipFile(p, "w") as z:
            for n, d in parts.items():
                z.writestr(n, d)
    assert xlgit.merge(*paths) == 0
    assert xlgit.read_cells(paths[1])["Data"]["A1"] == 2


def test_external_entities_are_never_read(tmp_path):
    # A workbook in a pull request must not be able to pull runner files
    # into the diff comment.
    secret = tmp_path / "secret.txt"
    secret.write_text("TOPSECRET")
    path = book(tmp_path / "evil.xlsx", {"A1": "x"})
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    doctype = f'<!DOCTYPE sst [<!ENTITY e SYSTEM "{secret.as_uri()}">]>'.encode()
    sst = parts["xl/sharedStrings.xml"]
    sst = sst.replace(b"?>", b"?>" + doctype, 1).replace(b">x<", b">&e;<")
    parts["xl/sharedStrings.xml"] = sst
    with zipfile.ZipFile(path, "w") as z:
        for n, d in parts.items():
            z.writestr(n, d)
    assert "TOPSECRET" not in repr(xlgit.read_cells(path))
