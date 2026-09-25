"""A branch that rewrites most cells of a small sheet in place hasn't moved
any rows, even if the diff finds it easier to show as rows deleted and
re-inserted. The merge must not follow that reading: the other branch's
edits to those rows would be dropped. Found by the fuzzer on a real
workbook (a totals block where every row holds the same formula)."""
import xlsxwriter

import xlgit

F = "=SUMIFS(Q:Q,P:P,$B{r},O:O,{c}$4)"
BASE = {"A2": "Totals by region", "C4": "East", "D4": "West",
        **{f"B{r}": name for r, name in zip(range(5, 9), ["Cookies", "Bars", "Snacks", "Crackers"])},
        **{f"{c}{r}": F.format(r=r, c=c) for r in range(5, 9) for c in "CD"},
        "C9": "=SUM(C5:C8)", "D9": "=SUM(D5:D8)"}
OURS = {"A2": True, "D2": "  padded  ", "C4": "#NULL!", "D4": False,
        "B5": "Cookies", "C5": 1e-07, "D5": 0.30000000000000004,
        "B6": "0123", "C6": F.format(r=6, c="C"), "D6": "=PI()*2",
        "B7": "#NAME?", "C7": 1e-07, "D7": -6690.585,
        "B8": "Crackers", "C8": "a < b & c > d", "D8": F.format(r=8, c="D"),
        "C9": "=SUM(C5:C8)", "D9": -2661.1, "D14": "#N/A"}
THEIRS = {"A2": True, "D2": "  padded  ", "C4": -891838, "D4": False,
          "B5": 1e-07, "C5": -9283.414893, "D5": F.format(r=5, c="D"),
          "B6": "Bars", "C6": '=IF(A1>0,"pos","neg")', "D6": 123456789.12345679,
          "B7": "#NAME?", "C7": 1e-07, "D7": F.format(r=7, c="D"),
          "B8": 9220.67, "C8": F.format(r=8, c="C"), "D8": 5699.3, "C9": "0123"}


def write(path, cells):
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Totals")
    for ref, v in cells.items():
        if isinstance(v, bool):
            ws.write_boolean(ref, v)
        elif isinstance(v, (int, float)):
            ws.write_number(ref, v)
        elif v.startswith("="):
            ws.write_formula(ref, v)
        else:
            ws.write_string(ref, v)
    wb.close()
    return str(path)


def test_rewritten_rows_are_not_treated_as_moved(tmp_path):
    paths = [write(tmp_path / f"{n}.xlsx", c) for n, c in (("base", BASE), ("ours", OURS), ("theirs", THEIRS))]
    b, o, t = (xlgit.read_cells(p)["Totals"] for p in paths)
    expected, clashes = {}, set()
    for ref in set(b) | set(o) | set(t):
        bv, ov, tv = b.get(ref), o.get(ref), t.get(ref)
        if tv == bv or tv == ov:
            v = ov
        elif ov == bv:
            v = tv
        else:
            v = ov
            clashes.add(ref)
        if v is not None:
            expected[ref] = v
    rc = xlgit.merge(*paths)
    merged = xlgit.read_cells(paths[1])["Totals"]
    assert merged == expected
    assert rc == (1 if clashes else 0)
    assert merged["B5"] == 1e-07  # their edit, on a row ours only rewrote
