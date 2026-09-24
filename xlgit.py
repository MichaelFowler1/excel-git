"""xlgit: make Excel workbooks behave like code in git and GitHub.

Commands:
  textconv FILE                    print a workbook as diffable text (git diff driver)
  diff OLD NEW [--markdown]        cell-level diff between two workbooks
  merge BASE OURS THEIRS [PATH]    cell-level 3-way merge (git merge driver)
  install                          wire the drivers into the current git repo
"""
import io
import os
import subprocess
import sys

from openpyxl import load_workbook
from openpyxl.comments import Comment

EXTS = ("*.xlsx", "*.xlsm")


def _load(path, keep_vba=False):
    """openpyxl refuses files without an Excel extension, but git hands its
    drivers temp files like .merge_file_a1b2c3, so read through a byte stream."""
    with open(path, "rb") as f:
        return load_workbook(io.BytesIO(f.read()), keep_vba=keep_vba)


def read_cells(path):
    """Return {sheet: {coord: value}} where formulas stay as '=...' strings."""
    if not path or not os.path.exists(path) or os.path.getsize(path) == 0:
        return {}
    wb = _load(path)
    out = {}
    for ws in wb.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    cells[c.coordinate] = c.value
        out[ws.title] = cells
    return out


def fmt(v):
    if v is None:
        return "(empty)"
    return repr(v) if isinstance(v, str) and not v.startswith("=") else str(v)


# ---------- textconv ----------

def textconv(path):
    for sheet, cells in read_cells(path).items():
        print(f"=== sheet: {sheet} ===")
        for coord, v in cells.items():
            print(f"{sheet}!{coord}\t{fmt(v)}")


# ---------- diff ----------

def diff_cells(old, new):
    """Yield (kind, sheet, coord, old_value, new_value)."""
    for sheet in sorted(set(old) | set(new)):
        if sheet not in old:
            yield ("sheet added", sheet, "", None, None)
        elif sheet not in new:
            yield ("sheet removed", sheet, "", None, None)
            continue
        o, n = old.get(sheet, {}), new.get(sheet, {})
        for coord in sorted(set(o) | set(n), key=_cell_sort_key):
            ov, nv = o.get(coord), n.get(coord)
            if ov != nv:
                kind = "added" if ov is None else "removed" if nv is None else "changed"
                yield (kind, sheet, coord, ov, nv)


def _cell_sort_key(coord):
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    col, row = coordinate_from_string(coord)
    return (row, column_index_from_string(col))


def diff(old_path, new_path, markdown=False, title=None):
    changes = list(diff_cells(read_cells(old_path), read_cells(new_path)))
    if markdown:
        print(f"### {title or new_path}\n")
        if not changes:
            print("No cell changes (formatting/charts may still differ).\n")
            return 0
        print(f"{len(changes)} change(s)\n")
        print("| Sheet | Cell | Change | Before | After |")
        print("|---|---|---|---|---|")
        for kind, sheet, coord, ov, nv in changes[:500]:
            esc = lambda v: fmt(v).replace("|", "\\|") if v is not None else ""
            print(f"| {sheet} | {coord} | {kind} | {esc(ov)} | {esc(nv)} |")
        if len(changes) > 500:
            print(f"\n...and {len(changes) - 500} more.")
        print()
    else:
        for kind, sheet, coord, ov, nv in changes:
            print(f"{kind:13} {sheet}!{coord}  {fmt(ov)} -> {fmt(nv)}")
    return 1 if changes else 0


# ---------- merge ----------

def merge(base_path, ours_path, theirs_path, display_path=None):
    """3-way merge at cell level. Result is written over OURS (git convention).

    Exit 0 = clean merge, 1 = conflicts (conflicting cells keep our value
    and get an Excel comment showing both sides)."""
    base, ours, theirs = read_cells(base_path), read_cells(ours_path), read_cells(theirs_path)
    keep_vba = (display_path or ours_path).lower().endswith(".xlsm")
    wb = _load(ours_path, keep_vba=keep_vba)
    conflicts = []
    applied = 0

    for sheet in theirs:
        if sheet not in wb.sheetnames:
            if sheet in base:
                continue  # we deleted it, they edited it: keep our deletion, flag below
            wb.create_sheet(sheet)  # new sheet on their side
        ws = wb[sheet]
        b, o, t = base.get(sheet, {}), ours.get(sheet, {}), theirs[sheet]
        for coord in set(b) | set(t):
            bv, ov, tv = b.get(coord), o.get(coord), t.get(coord)
            if tv == bv or tv == ov:
                continue  # they didn't touch it, or we both made the same edit
            if ov == bv:
                ws[coord].value = tv  # only they changed it
                applied += 1
            else:
                conflicts.append((sheet, coord, bv, ov, tv))
                ws[coord].comment = Comment(
                    f"MERGE CONFLICT\nbase: {fmt(bv)}\nours (kept): {fmt(ov)}\ntheirs: {fmt(tv)}",
                    "xlgit")

    for sheet in base:
        if sheet in theirs or sheet not in wb.sheetnames:
            continue
        if ours.get(sheet) == base[sheet]:
            del wb[sheet]  # they deleted a sheet we didn't touch
        else:
            conflicts.append((sheet, "(sheet)", "exists", "edited", "deleted"))

    if conflicts:
        cs = wb["_merge_conflicts"] if "_merge_conflicts" in wb.sheetnames else wb.create_sheet("_merge_conflicts")
        cs.append(["sheet", "cell", "base", "ours (kept)", "theirs"])
        for row in conflicts:
            cs.append([str(x) if x is not None else "" for x in row])

    wb.save(ours_path)
    name = display_path or ours_path
    print(f"xlgit merge {name}: {applied} cell(s) taken from theirs, {len(conflicts)} conflict(s)",
          file=sys.stderr)
    for sheet, coord, bv, ov, tv in conflicts:
        print(f"  CONFLICT {sheet}!{coord}: base={fmt(bv)} ours={fmt(ov)} theirs={fmt(tv)}", file=sys.stderr)
    return 1 if conflicts else 0


# ---------- install ----------

def install():
    here = os.path.abspath(__file__).replace("\\", "/")
    py = sys.executable.replace("\\", "/")
    cmd = f'"{py}" "{here}"'
    subprocess.check_call(["git", "config", "diff.xlsx.textconv", f"{cmd} textconv"])
    subprocess.check_call(["git", "config", "diff.xlsx.binary", "true"])
    subprocess.check_call(["git", "config", "merge.xlsx.name", "xlgit cell-level merge"])
    subprocess.check_call(["git", "config", "merge.xlsx.driver", f"{cmd} merge %O %A %B %P"])
    top = subprocess.check_output(["git", "rev-parse", "--show-toplevel"], text=True).strip()
    attrs = os.path.join(top, ".gitattributes")
    existing = open(attrs).read() if os.path.exists(attrs) else ""
    with open(attrs, "a") as f:
        for ext in EXTS:
            line = f"{ext} diff=xlsx merge=xlsx"
            if line not in existing:
                f.write(line + "\n")
    print("xlgit installed: git diff / git merge now understand .xlsx/.xlsm in this repo.")


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    cmd, args = argv[0], argv[1:]
    if cmd == "textconv":
        textconv(args[0])
        return 0
    if cmd == "diff":
        md = "--markdown" in args
        paths = [a for a in args if not a.startswith("--")]
        title = next((a.split("=", 1)[1] for a in args if a.startswith("--title=")), None)
        rc = diff(paths[0], paths[1], markdown=md, title=title)
        return 0 if md else rc
    if cmd == "merge":
        return merge(*args[:4])
    if cmd == "install":
        install()
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
