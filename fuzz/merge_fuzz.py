"""Merge-fuzz xlgit against a folder of real workbooks.

For every workbook it makes two branches ("ours" and "theirs") by editing the
workbook's XML directly, the way a spreadsheet app's save would, runs the
xlgit merge, and checks the result:

- the zip opens, every XML part parses, every part has a content type and
  every relationship points at a part that exists
- sheet XML is in the order Excel insists on (rows and cells ascending, no
  duplicates, a cell's row number matching its row), so Excel won't offer
  to "repair" the file
- openpyxl can read it
- every edit from both branches is there with the exact type and value
  written (a number stays that number, text that starts with "=" stays text)
- every cell nobody edited is unchanged, and nothing else was lost: every
  chart, image, comment, table and pivot in ours is still there
- the conflicts reported are exactly the cells both branches edited
  differently

The edits are written by this file, not by xlgit, so a bug in xlgit's cell
writer can't hide itself.

    python fuzz/merge_fuzz.py CORPUS_DIR [--jobs 4] [--limit N] [--seed 1]
        [--rounds 1] [--out results.jsonl] [--keep failed/]

Prints a summary grouped by failure; each JSONL line holds one file's result.
"""
import argparse
import collections
import io
import json
import math
import multiprocessing
import os
import random
import re
import signal
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import xlgit  # noqa: E402

MAIN = xlgit.MAIN
CT_NS = xlgit.CTYPES
m = xlgit.m
TIMEOUT = 120
MAX_BYTES = 20_000_000

# Text chosen to trip up a writer: leading "=" and "#" look like formulas and
# errors, whitespace must be preserved, XML-special and non-ASCII characters.
TEXTS = ["hello", "  padded  ", "=== TOTAL ===", "#N/A", "a < b & c > d", "naïve café ✓",
         "line1\nline2", "TRUE", "0123", "1e5", "'quoted", "Budget!A1", "x" * 300]
FORMULAS = ["SUM(A1:A3)", "1+1", "IF(A1>0,\"pos\",\"neg\")", "A1&\"x\"", "PI()*2", "TODAY()"]


# ---------- an independent reader/writer for sheet XML ----------

def col_row(ref):
    col, row = coordinate_from_string(ref)
    return column_index_from_string(col), row


def iter_rows(sd):
    """(row_number, row_el) with implicit numbers filled in (r is optional)."""
    last = 0
    for row in sd.findall(m("row")):
        last = int(row.get("r")) if row.get("r") else last + 1
        yield last, row


def iter_cells(row, rnum):
    last = 0
    for c in row.findall(m("c")):
        last = col_row(c.get("r"))[0] if c.get("r") else last + 1
        yield f"{get_column_letter(last)}{rnum}", c


def shared_strings(pkg):
    part = pkg.target(pkg.workbook, "sharedStrings")
    if part not in pkg.parts:
        return []
    return ["".join(t.text or "" for t in si.iter(m("t")) if _not_phonetic(t))
            for si in pkg.xml(part).iter(m("si"))]


def _not_phonetic(t):
    p = t.getparent()
    return p is None or etree.QName(p).localname != "rPh"


def raw_value(c, sst):
    """A cell's value exactly as stored: ('f', text) | ('s', text) | ('n', float)
    | ('b', bool) | ('e', text) | ('arr',) | None."""
    f = c.find(m("f"))
    if f is not None:
        if f.get("t") in ("shared", "array") and not (f.text or "").strip():
            return ("sharedref",)
        return ("f", f.text or "")
    t = c.get("t")
    v = c.find(m("v"))
    if t == "inlineStr":
        return ("s", "".join(x.text or "" for x in c.iter(m("t")) if _not_phonetic(x)))
    if v is None or v.text is None:
        return None
    if t == "s":
        i = int(v.text)
        return ("s", sst[i] if i < len(sst) else None)
    if t == "str":
        return ("s", v.text)
    if t == "b":
        return ("b", v.text.strip() in ("1", "true"))
    if t == "e":
        return ("e", v.text)
    try:
        return ("n", float(v.text))
    except ValueError:
        return ("bad", v.text)


def same(a, b):
    if a and b and a[0] == b[0] == "n":
        return a[1] == b[1] or math.isclose(a[1], b[1], rel_tol=1e-15, abs_tol=0)
    return a == b


class Editor:
    """Applies cell edits to one package's sheet XML."""

    def __init__(self, data):
        self.pkg = xlgit.Package(data)
        self.data = data
        self.trees = {}
        self.sst_part = self.pkg.target(self.pkg.workbook, "sharedStrings")
        self.sst = shared_strings(self.pkg)
        self.sst_root = self.pkg.xml(self.sst_part) if self.sst_part in self.pkg.parts else None

    def tree(self, part):
        if part not in self.trees:
            self.trees[part] = etree.fromstring(self.pkg.parts[part])
        return self.trees[part]

    def set(self, part, coord, value, style=None, rng=random):
        sd = self.tree(part).find(m("sheetData"))
        col, rnum = col_row(coord)
        row = next((r for n, r in iter_rows(sd) if n == rnum), None)
        if row is None:
            row = etree.Element(m("row"))
            row.set("r", str(rnum))
            later = next((r for n, r in iter_rows(sd) if n > rnum), None)
            later.addprevious(row) if later is not None else sd.append(row)
            if any(r.get("r") is None for r in sd.findall(m("row"))):
                # Keep implicit numbering intact: spell every row number out.
                for n, r in list(iter_rows(sd)):
                    r.set("r", str(n))
        c = next((x for ref, x in iter_cells(row, rnum) if ref == coord), None)
        if c is None:
            for ref, x in list(iter_cells(row, rnum)):
                x.set("r", ref)
            c = etree.Element(m("c"))
            c.set("r", coord)
            later = next((x for ref, x in iter_cells(row, rnum) if col_row(ref)[0] > col), None)
            later.addprevious(c) if later is not None else row.append(c)
        for child in list(c):
            if etree.QName(child).localname in ("f", "v", "is"):
                c.remove(child)
        for a in ("t", "cm", "vm"):
            c.attrib.pop(a, None)
        if style is not None:
            c.set("s", style)
        if value is None:
            return
        kind, v = value
        new = []
        if kind == "n":
            new.append(_el("v", repr(v)))
        elif kind == "b":
            c.set("t", "b")
            new.append(_el("v", "1" if v else "0"))
        elif kind == "f":
            new.append(_el("f", v))
        elif kind == "e":
            c.set("t", "e")
            new.append(_el("v", v))
        elif kind == "s":
            if self.sst_root is not None and rng.random() < 0.6:
                si = etree.SubElement(self.sst_root, m("si"))
                t = etree.SubElement(si, m("t"))
                t.text = v
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                self.sst.append(v)
                c.set("t", "s")
                new.append(_el("v", str(len(self.sst) - 1)))
            else:
                c.set("t", "inlineStr")
                is_ = etree.Element(m("is"))
                t = etree.SubElement(is_, m("t"))
                t.text = v
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                new.append(is_)
        for i, el in enumerate(new):
            c.insert(i, el)

    def save(self):
        trees = dict(self.trees)
        if self.sst_root is not None and self.sst_root.find(m("si")) is not None:
            self.sst_root.set("uniqueCount", str(len(self.sst)))
            self.sst_root.attrib.pop("count", None)
            trees[self.sst_part] = self.sst_root
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(self.data)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for info in zin.infolist():
                if info.filename in trees:
                    z.writestr(info.filename, etree.tostring(trees[info.filename], xml_declaration=True,
                                                             encoding="UTF-8", standalone=True))
                else:
                    z.writestr(info, zin.read(info))
        return out.getvalue()


def _el(tag, text):
    el = etree.Element(m(tag))
    el.text = text
    return el


# ---------- choosing edits ----------

def protected_areas(pkg, part):
    """Ranges not to edit: table headers (renaming a column rewrites the
    table), pivot output, array formulas, and shared-formula ranges."""
    refs = xlgit.pivot_locations(pkg, part)
    for _, typ, t, ext in pkg.rels(part):
        if not ext and xlgit.kind_of(typ) == "table" and t in pkg.parts:
            ref = pkg.xml(t).get("ref")
            if ref:
                c1, r1, c2, _ = xlgit.parse_ref(ref)
                refs.append(xlgit.format_ref(c1, r1, c2, r1))
                if pkg.xml(t).get("totalsRowCount") not in (None, "0"):
                    refs.append(ref)
    root = pkg.xml(part)
    for f in root.iter(m("f")):
        if f.get("ref") and f.get("t") in ("shared", "array"):
            refs.append(f.get("ref"))
    for mc in root.iter(m("mergeCell")):
        refs.append(mc.get("ref"))
    return [r for r in refs if r]


def plan(pkg, rng):
    """Pick cells to edit on each branch: {part: [(coord, ours, theirs)]}
    where ours/theirs is a value, None to clear, or "keep"."""
    sst = shared_strings(pkg)
    sheets = [(n, p) for n, _, p in pkg.sheets() if p in pkg.parts and pkg.xml(p).find(m("sheetData")) is not None]
    out = {}
    for name, part in rng.sample(sheets, min(len(sheets), 3)):
        root = pkg.xml(part)
        sd = root.find(m("sheetData"))
        existing = {}
        for rnum, row in iter_rows(sd):
            for ref, c in iter_cells(row, rnum):
                existing[ref] = c
        prot = protected_areas(pkg, part)
        ok = lambda ref: not any(xlgit.overlaps(ref, p) for p in prot)
        editable = [ref for ref, c in existing.items() if ok(ref) and raw_value(c, sst) not in (("sharedref",),)]
        maxc = max([col_row(r)[0] for r in existing] or [1])
        maxr = max([col_row(r)[1] for r in existing] or [1])
        fresh = set()
        for _ in range(rng.randint(2, 8)):
            ref = f"{get_column_letter(rng.randint(1, maxc + 3))}{rng.randint(1, maxr + 5)}"
            if ref not in existing and ok(ref):
                fresh.add(ref)
        picks = rng.sample(editable, min(len(editable), rng.randint(3, 25))) + sorted(fresh)
        rng.shuffle(picks)
        edits = []
        for i, ref in enumerate(picks):
            roll = rng.random()
            if roll < 0.4:
                edits.append((ref, random_value(rng), "keep"))
            elif roll < 0.8:
                edits.append((ref, "keep", random_value(rng, ref in existing)))
            elif roll < 0.9:
                edits.append((ref, random_value(rng), random_value(rng)))  # conflict (usually)
            else:
                v = random_value(rng)
                edits.append((ref, v, v))  # both made the same edit
        if edits:
            out[part] = (name, edits)
    return out


def random_value(rng, can_clear=True):
    roll = rng.random()
    if roll < 0.08 and can_clear:
        return None
    if roll < 0.4:
        return ("n", rng.choice([rng.randint(-10**6, 10**6), round(rng.uniform(-1e4, 1e4), rng.randint(0, 6)),
                                 0.1 + 0.2, 1e-7, 123456789.123456789, 45000.5]))
    if roll < 0.75:
        return ("s", rng.choice(TEXTS))
    if roll < 0.9:
        return ("f", rng.choice(FORMULAS))
    if roll < 0.95:
        return ("b", rng.random() < 0.5)
    return ("e", rng.choice(sorted(xlgit.ERRORS)))


# ---------- checking a result ----------

class Check(Exception):
    def __init__(self, kind, detail=""):
        super().__init__(f"{kind}: {detail}")
        self.kind, self.detail = kind, detail


def structural(data, label):
    """Problems any reader might trip on, as a Counter of "kind: where".
    Raises Check only when the file can't be read at all."""
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        bad = z.testzip()
    except Exception as e:
        raise Check(f"{label}: zip unreadable", repr(e))
    if bad:
        raise Check(f"{label}: zip crc", bad)
    out = collections.Counter()
    names = set(z.namelist())
    if len(names) != len(z.namelist()):
        out["duplicate zip entries"] += 1
    pkg = xlgit.Package(data)
    for name in names:
        if name.endswith((".xml", ".rels")):
            try:
                etree.fromstring(pkg.parts[name])
            except etree.XMLSyntaxError:
                out[f"xml does not parse: {name}"] += 1
    try:
        ct = pkg.xml("[Content_Types].xml")
        wb = pkg.xml(pkg.workbook)
        sheets = pkg.sheets()
    except Exception as e:
        raise Check(f"{label}: unreadable package", repr(e)[:200])
    if ct is None or wb is None:
        raise Check(f"{label}: unreadable package", "no content types or workbook")
    lower = {n.lower() for n in names}
    overrides = {o.get("PartName").lstrip("/").lower() for o in ct.iter(f"{{{CT_NS}}}Override")}
    defaults = {d.get("Extension").lower() for d in ct.iter(f"{{{CT_NS}}}Default")}
    for name in names:
        if name != "[Content_Types].xml" and not name.endswith("/") and name.lower() not in overrides \
                and name.rsplit(".", 1)[-1].lower() not in defaults:
            out[f"part without content type: {name}"] += 1
    for o in overrides - lower:
        out[f"content type for missing part: {o}"] += 1
    for name in names:
        if not name.endswith(".rels"):
            continue
        d, b = name.rsplit("_rels/", 1) if "_rels/" in name else ("", name)
        owner = (d + b[:-5]) if b != ".rels" else ""
        ids = set()
        try:
            rels = etree.fromstring(pkg.parts[name])
        except etree.XMLSyntaxError:
            continue
        for r in rels:
            if r.get("Id") in ids:
                out[f"duplicate relationship id: {name}"] += 1
            ids.add(r.get("Id"))
            if r.get("TargetMode") == "External" or not r.get("Target"):
                continue
            t = xlgit.resolve(owner, r.get("Target"))
            if t.lower() not in lower and not t.endswith("/"):
                out[f"relationship to missing part: {name} -> {t}"] += 1
    seen = collections.Counter(n.lower() for n, _, _ in sheets)
    for n, k in seen.items():
        if k > 1:
            out[f"duplicate sheet name: {n}"] += 1
    for _, _, part in sheets:
        if part in pkg.parts:
            try:
                for kind, k in sheet_order_problems(pkg.xml(part)).items():
                    out[f"{kind}: {part}"] += k
            except etree.XMLSyntaxError:
                pass
    return out


def new_problems(before, after):
    """Problems in the merged file that weren't in the input. Part names can
    move in a merge, so compare by kind."""
    kind = lambda k: k.split(":", 1)[0]
    b = collections.Counter()
    for k, n in before.items():
        b[kind(k)] += n
    a = collections.Counter()
    examples = {}
    for k, n in after.items():
        a[kind(k)] += n
        examples.setdefault(kind(k), k)
    return [examples[k] for k in a if a[k] > b[k]]


def sheet_order_problems(root):
    out = collections.Counter()
    order = {t: i for i, t in enumerate(xlgit.WS_ORDER)}
    idx = [order.get(etree.QName(c).localname, -1) for c in root if isinstance(c.tag, str)]
    if idx != sorted(idx):
        out["worksheet children out of order"] += 1
    sd = root.find(m("sheetData"))
    if sd is None:
        return out
    prev = 0
    for rnum, row in iter_rows(sd):
        if rnum <= prev:
            out["rows out of order or duplicated"] += 1
        prev = rnum
        pc = 0
        for ref, c in iter_cells(row, rnum):
            col, r = col_row(ref)
            if r != rnum:
                out["cell in wrong row"] += 1
            if col <= pc:
                out["cells out of order or duplicated"] += 1
            pc = col
            kids = [etree.QName(k).localname for k in c if isinstance(k.tag, str)]
            ko = {"f": 0, "v": 1, "is": 2, "extLst": 3}
            if [ko.get(k, 9) for k in kids] != sorted(ko.get(k, 9) for k in kids):
                out["cell children out of order"] += 1
            if c.get("t") == "inlineStr" and c.find(m("is")) is None and c.find(m("f")) is None:
                out["inlineStr without <is>"] += 1
            if c.get("t") == "s" and c.find(m("v")) is None:
                out["shared string cell without <v>"] += 1
    return out


def sheet_values(pkg):
    """{sheet_name: {coord: raw_value}} straight from the XML."""
    sst = shared_strings(pkg)
    out = {}
    for name, _, part in pkg.sheets():
        if part not in pkg.parts:
            continue
        sd = pkg.xml(part).find(m("sheetData"))
        vals = {}
        for rnum, row in (iter_rows(sd) if sd is not None else []):
            for ref, c in iter_cells(row, rnum):
                v = raw_value(c, sst)
                if v is not None:
                    vals[ref] = v
        out[name] = vals
    return out


def styles(pkg, part):
    sd = pkg.xml(part).find(m("sheetData"))
    return {ref: c.get("s") for rnum, row in iter_rows(sd) for ref, c in iter_cells(row, rnum)}


# ---------- one file ----------

def run_one(path, seed, keep=None):
    rng = random.Random(f"{seed}:{os.path.basename(path)}")
    res = {"file": str(path), "seed": seed}
    data = Path(path).read_bytes()
    try:
        base_problems = structural(data, "input")
    except Check as e:
        res.update(status="skip", reason=e.kind, detail=e.detail[:300])
        return res
    except Exception as e:
        res.update(status="skip", reason="input unreadable", detail=repr(e)[:300])
        return res
    pkg = xlgit.Package(data)
    try:
        xlgit.read_package_cells(pkg)
    except Exception as e:
        res.update(status="fail", reason="xlgit cannot read input", detail=_where(e))
        return res
    try:
        load_workbook(io.BytesIO(data))
        openpyxl_ok = True
    except Exception:
        openpyxl_ok = False

    edits = plan(pkg, rng)
    if not edits:
        res.update(status="skip", reason="nothing to edit")
        return res
    ours, theirs = Editor(data), Editor(data)
    # Occasionally restyle a cell theirs edits (a style that exists in styles.xml).
    styles_root = pkg.xml(pkg.target(pkg.workbook, "styles")) if pkg.target(pkg.workbook, "styles") in pkg.parts else None
    nxf = len(styles_root.find(m("cellXfs"))) if styles_root is not None and styles_root.find(m("cellXfs")) is not None else 0
    expect = {}  # (sheet, coord) -> (value, conflict, their_style)
    base_vals = sheet_values(pkg)
    for part, (name, cell_edits) in edits.items():
        for coord, ov, tv in cell_edits:
            tstyle = str(rng.randrange(nxf)) if nxf and rng.random() < 0.2 and tv != "keep" else None
            if ov != "keep":
                ours.set(part, coord, ov, rng=rng)
            if tv != "keep":
                theirs.set(part, coord, tv, style=tstyle, rng=rng)
            bv = base_vals.get(name, {}).get(coord)
            ov = bv if ov == "keep" else ov
            tv = bv if tv == "keep" else tv
            if same(tv, bv) or same(ov, tv):
                expect[(name, coord)] = (ov, False, None)
            elif same(ov, bv):
                expect[(name, coord)] = (tv, False, tstyle)
            else:
                expect[(name, coord)] = (ov, True, None)
    obytes, tbytes = ours.save(), theirs.save()
    res["edits"] = len(expect)

    old_alarm = signal.signal(signal.SIGALRM, _timeout)
    limit = TIMEOUT + len(data) // 50_000  # a 14 MB workbook gets ~7 minutes
    signal.alarm(limit)
    try:
        with tempfile.TemporaryDirectory() as td:
            b, o, t = (os.path.join(td, n) for n in ("base.xlsx", "ours.xlsx", "theirs.xlsx"))
            Path(b).write_bytes(data)
            Path(o).write_bytes(obytes)
            Path(t).write_bytes(tbytes)
            try:
                mg = xlgit.Merger(b, o, t)
                merged = mg.run()
            except Timeout:
                raise Check("merge timed out", f">{limit}s")
            except Exception as e:
                raise Check("merge crashed", _where(e))
            Path(o).write_bytes(merged)
            verify(pkg, data, obytes, merged, mg, expect, base_problems, o, openpyxl_ok)
            # The diff path should cope with anything the merge handles.
            try:
                with open(os.devnull, "w") as dn:
                    old = sys.stdout
                    sys.stdout = dn
                    try:
                        xlgit.diff(b, o, markdown=True)
                        xlgit.textconv(b)
                    finally:
                        sys.stdout = old
            except Exception as e:
                raise Check("diff crashed", _where(e))
        res.update(status="ok", conflicts=len(mg.conflicts))
    except Timeout:
        res.update(status="fail", reason="check timed out", detail=f">{limit}s")
    except Check as e:
        res.update(status="fail", reason=e.kind, detail=str(e.detail)[:600])
        if keep:
            d = Path(keep) / Path(path).stem
            d.mkdir(parents=True, exist_ok=True)
            for n, blob in (("base", data), ("ours", obytes), ("theirs", tbytes)):
                (d / f"{n}.xlsx").write_bytes(blob)
            if "merged" in locals():
                (d / "merged.xlsx").write_bytes(merged)
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, old_alarm)
    return res


def verify(pkg, data, obytes, merged, mg, expect, base_problems, merged_path, openpyxl_ok):
    worse = new_problems(base_problems, structural(merged, "merged"))
    if worse:
        raise Check(f"merged: {worse[0].split(':', 1)[0]}", "; ".join(worse)[:300])
    try:
        merged_cells = xlgit.read_cells(merged_path)
    except Exception as e:
        raise Check("xlgit cannot read merged", _where(e))
    if openpyxl_ok:
        try:
            load_workbook(io.BytesIO(merged))
        except Exception as e:
            raise Check("openpyxl cannot read merged", _where(e))
    mpkg = xlgit.Package(merged)
    mvals = sheet_values(mpkg)
    ovals = sheet_values(xlgit.Package(obytes))
    conflicts = {(s, c) for s, c, *_ in mg.conflicts}
    for (sheet, coord), (want, conflict, tstyle) in expect.items():
        got = mvals.get(sheet, {}).get(coord)
        if conflict:
            if (sheet, coord) not in conflicts:
                raise Check("conflict not reported", f"{sheet}!{coord}")
        elif (sheet, coord) in conflicts:
            raise Check("false conflict", f"{sheet}!{coord} want={want} ours={ovals.get(sheet, {}).get(coord)}")
        if not same(got, want):
            kind = "edit lost" if got == ovals.get(sheet, {}).get(coord) else "edit written wrong"
            if want and got and want[0] == "n" and got[0] == "n":
                kind = "number changed"
            raise Check(kind, f"{sheet}!{coord} want={want!r} got={got!r}")
        if tstyle is not None and mg.same_styles:
            part = dict((n, p) for n, _, p in mpkg.sheets())[sheet]
            if styles(mpkg, part).get(coord) != tstyle:
                raise Check("their cell style not carried over", f"{sheet}!{coord}")
    extra = conflicts - set(expect) - {(s, c) for s, c in conflicts if s == "(file)" or c == "(sheet)"}
    if extra:
        raise Check("false conflict", repr(sorted(extra)[:3]))
    if any(s == "(file)" or c == "(sheet)" for s, c in conflicts):
        raise Check("false object/sheet conflict", repr([c for c in mg.conflicts if c[0] == "(file)" or c[1] == "(sheet)"][:2]))
    # Cells nobody touched: merged == ours (openpyxl's view, formulas expanded).
    with tempfile.TemporaryDirectory() as td:
        op = os.path.join(td, "o.xlsx")
        Path(op).write_bytes(obytes)
        ours_cells = xlgit.read_cells(op)
    for sheet, cells in ours_cells.items():
        mc = merged_cells.get(sheet)
        if mc is None:
            raise Check("sheet lost", sheet)
        for coord in set(cells) | set(mc):
            if (sheet, coord) in expect:
                continue
            if not _eq(cells.get(coord), mc.get(coord)):
                raise Check("untouched cell changed", f"{sheet}!{coord} {cells.get(coord)!r} -> {mc.get(coord)!r}")
    if set(merged_cells) - set(ours_cells) - {xlgit.CONFLICT_SHEET}:
        raise Check("sheet appeared", repr(set(merged_cells) - set(ours_cells)))
    # Objects: nothing lost.
    oobj = xlgit.describe_objects(xlgit.Package(obytes))
    mobj = xlgit.describe_objects(mpkg)
    for sheet, items in oobj.items():
        missing = collections.Counter(items) - collections.Counter(mobj.get(sheet, []))
        if missing:
            raise Check("object lost or changed", f"{sheet}: {list(missing)[:2]}")
    lost = [p for p in xlgit.reachable(xlgit.Package(obytes).parts) - set(mpkg.parts)
            if not re.search(r"calcChain", p)]
    if lost:
        raise Check("part lost", repr(sorted(lost)[:3]))


def _eq(a, b):
    if isinstance(a, float) and isinstance(b, float):
        return a == b or (math.isnan(a) and math.isnan(b))
    return a == b


class Timeout(BaseException):
    """Not an Exception, so code under test can't swallow it."""


def _timeout(*_):
    raise Timeout()


def _where(e):
    tb = traceback.extract_tb(e.__traceback__)
    frames = [f for f in tb if f.filename.endswith("xlgit.py")] or tb[-1:]
    f = frames[-1]
    return f"{type(e).__name__}: {str(e)[:200]} @ xlgit.py:{f.lineno} {f.name}"


def _job(args):
    path, seed, keep = args
    try:
        return run_one(path, seed, keep)
    except Exception as e:
        return {"file": str(path), "seed": seed, "status": "fail", "reason": "harness error", "detail": _where(e)}


def signature(r):
    d = r.get("detail", "")
    if r["reason"] in ("merge crashed", "diff crashed", "openpyxl cannot read merged", "xlgit cannot read merged",
                       "xlgit cannot read input", "harness error"):
        return f"{r['reason']}: {re.sub(r': .*? @', ' @', d)}"
    return r["reason"]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("corpus")
    ap.add_argument("--jobs", type=int, default=os.cpu_count())
    ap.add_argument("--limit", type=int)
    ap.add_argument("--seed", type=int, default=1)
    ap.add_argument("--rounds", type=int, default=1, help="seeds per file")
    ap.add_argument("--sample", type=int, help="random sample of this many files")
    ap.add_argument("--out", default="fuzz-results.jsonl")
    ap.add_argument("--keep", help="copy base/ours/theirs/merged of failures here")
    ap.add_argument("--only", help="rerun only files listed (one per line) in this file")
    a = ap.parse_args(argv)
    files = sorted(p for p in Path(a.corpus).rglob("*") if p.suffix.lower() in (".xlsx", ".xlsm")
                   and p.stat().st_size <= MAX_BYTES)
    if a.only:
        wanted = set(Path(a.only).read_text().split("\n"))
        files = [f for f in files if str(f) in wanted]
    if a.sample:
        files = random.Random(a.seed).sample(files, min(a.sample, len(files)))
    files = files[:a.limit]
    jobs = [(str(f), a.seed + i, a.keep) for f in files for i in range(a.rounds)]
    stats, sigs = collections.Counter(), collections.defaultdict(list)
    with open(a.out, "w") as out, multiprocessing.Pool(a.jobs, maxtasksperchild=50) as pool:
        for i, r in enumerate(pool.imap_unordered(_job, jobs, chunksize=4), 1):
            out.write(json.dumps(r) + "\n")
            stats[r["status"]] += 1
            if r["status"] == "fail":
                sigs[signature(r)].append(r)
            elif r["status"] == "skip":
                stats["skip: " + r["reason"]] += 1
            if i % 200 == 0:
                print(f"{i}/{len(jobs)} {dict((k, v) for k, v in stats.items() if ':' not in k)}", file=sys.stderr)
    print(f"\n{len(jobs)} runs: " + ", ".join(f"{k} {v}" for k, v in sorted(stats.items())))
    for sig, rs in sorted(sigs.items(), key=lambda kv: -len(kv[1])):
        print(f"\n[{len(rs)}] {sig}")
        for r in rs[:3]:
            print(f"    {os.path.basename(r['file'])} (seed {r['seed']}): {r.get('detail', '')[:200]}")
    return 1 if sigs else 0


if __name__ == "__main__":
    sys.exit(main())
