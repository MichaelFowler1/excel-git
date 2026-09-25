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
from openpyxl.formula.translate import Translator
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


def raw_value(c, sst, ref=None, shared=None):
    """A cell's value exactly as stored: ('f', text) | ('s', text) | ('n', float)
    | ('b', bool) | ('e', text) | None. A cell sharing a formula gets it
    translated from the cell that holds it (pass ref and a dict for shared)."""
    f = c.find(m("f"))
    if f is not None:
        if f.get("t") == "shared" and shared is not None:
            if (f.text or "").strip():
                shared[f.get("si")] = Translator("=" + f.text, ref)
            elif f.get("si") in shared:
                return ("f", shared[f.get("si")].translate_formula(ref)[1:])
        if not (f.text or "").strip():
            return ("sharedref",)
        return ("f", f.text or "")
    t = c.get("t")
    v = c.find(m("v"))
    if t == "inlineStr":
        if c.find(m("is")) is None:
            return None  # <c t="inlineStr"/> with no text is an empty cell
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
    """Equal as Excel sees it: numbers to 15 significant digits."""
    if a and b and a[0] == b[0] == "n":
        return a[1] == b[1] or f"{a[1]:.15g}" == f"{b[1]:.15g}"
    return a == b


class Editor:
    """Applies cell edits to one package's sheet XML."""

    def __init__(self, data):
        self.pkg = xlgit.Package(data)
        self.data = data
        self.trees = {}
        self.rows = {}  # part -> {row number: row element}
        self.removed = set()
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
        if part not in self.rows:
            self.rows[part] = dict(iter_rows(sd))
        row = self.rows[part].get(rnum)
        if row is None:
            row = etree.Element(m("row"))
            row.set("r", str(rnum))
            later = next((r for n, r in iter_rows(sd) if n > rnum), None)
            later.addprevious(row) if later is not None else sd.append(row)
            if any(r.get("r") is None for r in sd.findall(m("row"))):
                # Keep implicit numbering intact: spell every row number out.
                for n, r in list(iter_rows(sd)):
                    r.set("r", str(n))
            self.rows[part] = dict(iter_rows(sd))
        c = next((x for ref, x in iter_cells(row, rnum) if ref == coord), None)
        if c is None:
            for ref, x in list(iter_cells(row, rnum)):
                x.set("r", ref)
            c = etree.Element(m("c"))
            c.set("r", coord)
            later = next((x for ref, x in iter_cells(row, rnum) if col_row(ref)[0] > col), None)
            later.addprevious(c) if later is not None else row.append(c)
        f = c.find(m("f"))
        if f is not None and f.get("t") == "shared" and (f.text or "").strip():
            self.unshare(sd, f.get("si"))
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

    @staticmethod
    def unshare(sd, si):
        """Editing the cell that holds a shared formula: spell the formula out
        in every cell that shared it, as Excel does."""
        master = None
        for rnum, row in iter_rows(sd):
            for ref, c in iter_cells(row, rnum):
                f = c.find(m("f"))
                if f is None or f.get("t") != "shared" or f.get("si") != si:
                    continue
                if master is None and (f.text or "").strip():
                    master = Translator("=" + f.text, ref)
                elif master is not None:
                    f.text = master.translate_formula(ref)[1:]
                for a in ("t", "si", "ref"):
                    f.attrib.pop(a, None)

    # --- whole sheets, the way Excel does them ---

    def rename_sheet(self, old, new):
        wb = self.tree(self.pkg.workbook)
        for sh in wb.find(m("sheets")):
            if sh.get("name") == old:
                sh.set("name", new)
        quoted = lambda n: "'" + n.replace("'", "''") + "'"
        for d in wb.iter(m("definedName")):
            if d.text:
                d.text = d.text.replace(quoted(old) + "!", quoted(new) + "!")
                d.text = re.sub(rf"(?<![\w'.]){re.escape(old)}!", quoted(new) + "!", d.text)

    def add_sheet(self, name, cells):
        wbpart = self.pkg.workbook
        n = 1
        while f"xl/worksheets/sheet{n}.xml" in self.pkg.parts or f"xl/worksheets/sheet{n}.xml" in self.trees:
            n += 1
        part = f"xl/worksheets/sheet{n}.xml"
        root = etree.Element(m("worksheet"), nsmap={None: MAIN, "r": xlgit.REL})
        etree.SubElement(root, m("sheetData"))
        self.trees[part] = root
        for coord, v in sorted(cells.items(), key=lambda kv: col_row(kv[0])[::-1]):
            self.set(part, coord, v, rng=random.Random(0))
        rels = self.rels_tree(wbpart)
        ids = {r.get("Id") for r in rels}
        k = 1
        while f"rId{k}" in ids:
            k += 1
        r = etree.SubElement(rels, f"{{{xlgit.PKG_REL}}}Relationship")
        r.set("Id", f"rId{k}")
        r.set("Type", xlgit.WORKSHEET_REL)
        r.set("Target", xlgit.relative(part, wbpart))
        wb = self.tree(wbpart)
        sheets = wb.find(m("sheets"))
        sid = max([int(sh.get("sheetId") or 0) for sh in sheets] + [0]) + 1
        sh = etree.SubElement(sheets, m("sheet"))
        sh.set("name", name)
        sh.set("sheetId", str(sid))
        sh.set(xlgit.RID, f"rId{k}")
        ct = self.tree("[Content_Types].xml")
        o = etree.SubElement(ct, f"{{{CT_NS}}}Override")
        o.set("PartName", "/" + part)
        o.set("ContentType", xlgit.WORKSHEET_CT)

    def delete_sheet(self, name):
        wbpart = self.pkg.workbook
        wb = self.tree(wbpart)
        sheets = wb.find(m("sheets"))
        els = list(sheets)
        el = next(sh for sh in els if sh.get("name") == name)
        idx = els.index(el)
        rid = el.get(xlgit.RID)
        sheets.remove(el)
        rels = self.rels_tree(wbpart)
        part = None
        for r in list(rels):
            if r.get("Id") == rid:
                part = xlgit.resolve(wbpart, r.get("Target"))
                rels.remove(r)
        dn = wb.find(m("definedNames"))
        for d in list(dn if dn is not None else []):
            lid = d.get("localSheetId")
            if lid is not None and int(lid) == idx:
                dn.remove(d)
            elif lid is not None and int(lid) > idx:
                d.set("localSheetId", str(int(lid) - 1))
            elif d.text and re.search(rf"(^|[^\w.]){re.escape(name)}!|'{re.escape(name)}'!", d.text):
                d.text = "#REF!"
        for view in wb.iter(m("workbookView")):
            for a in ("activeTab", "firstSheet"):
                if view.get(a) and int(view.get(a)) >= len(sheets):
                    view.set(a, "0")
        if part:
            self.removed |= {part, xlgit.rels_name(part)}
            ct = self.tree("[Content_Types].xml")
            for o in list(ct):
                if (o.get("PartName") or "").lstrip("/") == part:
                    ct.remove(o)

    def rels_tree(self, owner):
        name = xlgit.rels_name(owner)
        if name not in self.trees:
            self.trees[name] = etree.fromstring(self.pkg.parts[name])
        return self.trees[name]

    def save(self):
        trees = dict(self.trees)
        if self.sst_root is not None and self.sst_root.find(m("si")) is not None:
            self.sst_root.set("uniqueCount", str(len(self.sst)))
            self.sst_root.attrib.pop("count", None)
            trees[self.sst_part] = self.sst_root
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(self.data)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            written = set()
            for info in zin.infolist():
                if info.filename in self.removed:
                    continue
                if info.filename in trees:
                    z.writestr(info.filename, etree.tostring(trees[info.filename], xml_declaration=True,
                                                             encoding="UTF-8", standalone=True))
                else:
                    z.writestr(info, zin.read(info))
                written.add(info.filename)
            for name, root in trees.items():
                if name not in written and name not in self.removed:
                    z.writestr(name, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))
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
        if f.get("ref") and f.get("t") == "array":
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
        editable = [ref for ref, c in existing.items() if ok(ref)]
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
        vals, shared = {}, {}
        for rnum, row in (iter_rows(sd) if sd is not None else []):
            for ref, c in iter_cells(row, rnum):
                v = raw_value(c, sst, ref, shared)
                if v is not None:
                    vals[ref] = v
        out[name] = vals
    return out


def styles(pkg, part):
    sd = pkg.xml(part).find(m("sheetData"))
    return {ref: c.get("s") for rnum, row in iter_rows(sd) for ref, c in iter_cells(row, rnum)}


# ---------- one file ----------

def run_one(path, seed, keep=None, writer=None):
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
    # Cross-check with openpyxl (what pandas uses), unless it can't read the
    # input either or is very slow on it.
    signal.signal(signal.SIGALRM, _timeout)
    signal.alarm(10)
    try:
        load_workbook(io.BytesIO(data))
        openpyxl_ok = True
    except (Exception, Timeout):
        openpyxl_ok = False
    finally:
        signal.alarm(0)

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
    ops = sheet_ops(pkg, edits, rng) if writer is None else {}
    if "rename" in ops:
        theirs.rename_sheet(*ops["rename"])
    if "add" in ops:
        theirs.add_sheet(*ops["add"])
    if "delete" in ops:
        theirs.delete_sheet(ops["delete"])
    obytes, tbytes = ours.save(), theirs.save()
    res["edits"] = len(expect)
    res["ops"] = sorted(ops)
    if writer == "libreoffice":
        try:
            obytes = libreoffice_resave(obytes)
            ours_problems = structural(obytes, "ours")
        except Exception as e:
            res.update(status="skip", reason="libreoffice could not re-save", detail=repr(e)[:200])
            return res
        base_problems = base_problems + ours_problems

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
            if writer:
                verify_three_way(data, obytes, tbytes, merged, mg, base_problems)
            else:
                verify(pkg, data, obytes, tbytes, merged, mg, expect, base_problems, o, openpyxl_ok, ops)
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


def verify(pkg, data, obytes, tbytes, merged, mg, expect, base_problems, merged_path, openpyxl_ok, ops):
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
    renamed = dict([ops["rename"]]) if "rename" in ops else {}
    to_merged = lambda sheet: renamed.get(sheet, sheet)
    if "delete" in ops and ops["delete"] in mvals:
        raise Check("deleted sheet came back", ops["delete"])
    if "add" in ops:
        name = ops["add"][0]
        if name not in mvals:
            raise Check("added sheet lost", name)
        if mvals[name] != sheet_values(xlgit.Package(tbytes))[name]:
            raise Check("added sheet changed", name)
    for old, new in renamed.items():
        if new not in mvals:
            raise Check("sheet rename lost", f"{old} -> {new}")
    for (sheet, coord), (want, conflict, tstyle) in expect.items():
        got = mvals.get(to_merged(sheet), {}).get(coord)
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
            part = dict((n, p) for n, _, p in mpkg.sheets())[to_merged(sheet)]
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
    gone = ops.get("delete")
    for sheet, cells in ours_cells.items():
        if sheet == gone:
            continue
        mc = merged_cells.get(to_merged(sheet))
        if mc is None:
            raise Check("sheet lost", sheet)
        for coord in set(cells) | set(mc):
            if (sheet, coord) in expect:
                continue
            if not _eq(cells.get(coord), mc.get(coord)):
                raise Check("untouched cell changed", f"{sheet}!{coord} {cells.get(coord)!r} -> {mc.get(coord)!r}")
    appeared = set(merged_cells) - {to_merged(n) for n in ours_cells} - {xlgit.CONFLICT_SHEET}
    if appeared - {ops.get("add", ("",))[0]}:
        raise Check("sheet appeared", repr(appeared))
    # Objects: nothing lost.
    oobj = xlgit.describe_objects(xlgit.Package(obytes))
    mobj = xlgit.describe_objects(mpkg)
    for sheet, items in oobj.items():
        if sheet == gone:
            continue
        missing = collections.Counter(items) - collections.Counter(mobj.get(to_merged(sheet), []))
        if missing:
            raise Check("object lost or changed", f"{sheet}: {list(missing)[:2]}")
    opkg = xlgit.Package(obytes)
    # The deleted sheet's part, and everything only it used, may go.
    dropped = {p for n, _, p in opkg.sheets() if n == gone}
    lost = [p for p in xlgit.reachable(opkg.parts) - set(mpkg.parts)
            if not re.search(r"calcChain", p) and not (gone and _only_used_by(opkg, p, dropped))]
    if lost:
        raise Check("part lost", repr(sorted(lost)[:3]))


def sheet_ops(pkg, edits, rng):
    """Whole-sheet changes on their branch: rename a sheet, add one, delete
    one nobody edited."""
    names = [n for n, _, p in pkg.sheets()]
    if any(n.lower() == xlgit.CONFLICT_SHEET for n in names):
        return {}
    edited = {name for name, _ in edits.values()}
    ops = {}
    if rng.random() < 0.3:
        old = rng.choice(names)
        new = ("Renamed " + old)[:31]
        if new.lower() not in {n.lower() for n in names}:
            ops["rename"] = (old, new)
    if rng.random() < 0.3:
        new = "Added by them"
        if new.lower() not in {n.lower() for n in names}:
            ops["add"] = (new, {"A1": ("s", "note"), "B2": ("n", 42), "C3": ("f", "B2*2")})
    candidates = [n for n, _, p in pkg.sheets() if n not in edited and n != ops.get("rename", ("",))[0] and p]
    if rng.random() < 0.4 and len(names) > 1 and candidates:
        ops["delete"] = rng.choice(candidates)
    return ops


# ---------- rows inserted or deleted on one branch ----------
#
# Written independently of xlgit's row handling: a regex finds references,
# not openpyxl's tokenizer, and the expected result is worked out here.

_REF = re.compile(r"(?<![A-Za-z0-9_.!\]$])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
                  r"(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?\d+:\$?\d+)(?![A-Za-z0-9_(\[])")
_PART = re.compile(r"^(\$?[A-Za-z]{0,3}\$?)(\d+)$")


def shift_formula(text, host, sheet, fn, fn_start, fn_end):
    """Renumber rows in references to `sheet` in a formula written on `host`."""
    if '"' in text:
        pieces = text.split('"')  # leave string literals alone
        return '"'.join(shift_formula(p, host, sheet, fn, fn_start, fn_end) if i % 2 == 0 else p
                        for i, p in enumerate(pieces))

    def one(mt):
        prefix, ref = mt.group(1) or "", mt.group(2)
        target = prefix[:-1] if prefix else host
        if target.startswith("'"):
            target = target[1:-1].replace("''", "'")
        if target != sheet:
            return mt.group(0)
        ends = ref.split(":")
        if len(ends) == 1:
            p = _PART.match(ends[0])
            r = fn(int(p.group(2)))
            return prefix + (f"{p.group(1)}{r}" if r is not None else "#REF!") if r is not None else "#REF!"
        a, b = _PART.match(ends[0]), _PART.match(ends[1])
        ra, rb = fn_start(int(a.group(2))), fn_end(int(b.group(2)))
        if ra is None or rb is None or ra > rb:
            return "#REF!"
        return f"{prefix}{a.group(1)}{ra}:{b.group(1)}{rb}"
    return _REF.sub(one, text)


class RowOp:
    """Insert `count` rows before row `at`, or delete rows at..at+count-1."""

    def __init__(self, kind, at, count):
        self.kind, self.at, self.count = kind, at, count

    def __call__(self, r):
        if self.kind == "insert":
            return r if r < self.at else r + self.count
        if self.at <= r < self.at + self.count:
            return None
        return r if r < self.at else r - self.count

    def start(self, r):
        return self(r) if self(r) is not None else self.at

    def end(self, r):
        return self(r) if self(r) is not None else (self.at - 1 if self.at > 1 else None)


def row_op_eligible(pkg, name, part):
    """Sheets simple enough for this file's own row insert: nothing on them
    that also needs moving (merged cells, tables, drawings, pivots...)."""
    root = pkg.xml(part)
    if root is None or etree.QName(root).localname != "worksheet":
        return False
    blockers = {"mergeCells", "conditionalFormatting", "dataValidations", "hyperlinks", "tableParts",
                "drawing", "legacyDrawing", "autoFilter", "rowBreaks", "extLst", "protectedRanges"}
    if any(etree.QName(c).localname in blockers for c in root if isinstance(c.tag, str)):
        return False
    if any(f.get("t") in ("array", "dataTable") for f in root.iter(m("f"))):
        return False
    wb = pkg.xml(pkg.workbook)
    if wb.find(m("pivotCaches")) is not None or wb.find(m("externalReferences")) is not None:
        return False
    dn = wb.find(m("definedNames"))
    if dn is not None and any(name in (d.text or "") for d in dn):
        return False
    if any(p.startswith("xl/charts/") and name.encode() in d for p, d in pkg.parts.items()):
        return False
    return bool(root.find(m("sheetData")) is not None and len(root.find(m("sheetData"))))


def unshare_all(sd):
    """Spell out every shared formula in one pass (Excel does this when
    rows move under them)."""
    masters = {}
    for rnum, row in iter_rows(sd):
        for ref, c in iter_cells(row, rnum):
            f = c.find(m("f"))
            if f is None or f.get("t") != "shared":
                continue
            si = f.get("si")
            if (f.text or "").strip() and si not in masters:
                masters[si] = Translator("=" + f.text, ref)
            elif si in masters:
                f.text = masters[si].translate_formula(ref)[1:]
            for a in ("t", "si", "ref"):
                f.attrib.pop(a, None)


def apply_row_op(ed, sheet, part, op):
    """Do to ed's copy what Excel does: move the cells, and renumber every
    formula that points at this sheet, on every sheet."""
    for pname, root_part in [(n, p) for n, _, p in ed.pkg.sheets() if p in ed.pkg.parts]:
        root = ed.tree(root_part)
        sd = root.find(m("sheetData"))
        if sd is None:
            continue
        unshare_all(sd)
        for f in root.iter(m("f")):
            if f.text:
                f.text = shift_formula(f.text, pname, sheet, op, op.start, op.end)
    root = ed.tree(part)
    sd = root.find(m("sheetData"))
    for n, row in list(iter_rows(sd)):
        new = op(n)
        if new is None:
            sd.remove(row)
            continue
        row.set("r", str(new))
        for ref, c in list(iter_cells(row, n)):
            c.set("r", f"{coordinate_from_string(ref)[0]}{new}")
    dim = root.find(m("dimension"))
    if dim is not None:
        root.remove(dim)
    ed.rows.pop(part, None)


def run_rows(path, seed, keep=None):
    """One branch inserts or deletes rows on a sheet; both edit cells."""
    rng = random.Random(f"rows:{seed}:{os.path.basename(path)}")
    res = {"file": str(path), "seed": seed, "mode": "rows"}
    signal.signal(signal.SIGALRM, _timeout)
    signal.alarm(TIMEOUT * 3)
    try:
        return _run_rows(path, seed, keep, rng, res)
    except Timeout:
        res.update(status="fail", reason="check timed out", detail="while preparing the row edit")
        return res
    finally:
        signal.alarm(0)


def _run_rows(path, seed, keep, rng, res):
    data = Path(path).read_bytes()
    try:
        base_problems = structural(data, "input")
        pkg = xlgit.Package(data)
        base_vals = sheet_values(pkg)
    except Exception as e:
        res.update(status="skip", reason="input unreadable", detail=repr(e)[:200])
        return res
    sheets = [(n, p) for n, _, p in pkg.sheets() if p in pkg.parts and row_op_eligible(pkg, n, p)]
    if not sheets:
        res.update(status="skip", reason="no sheet simple enough for a row edit")
        return res
    sheet, part = rng.choice(sheets)
    rows_used = sorted({col_row(c)[1] for c in base_vals.get(sheet, {})})
    if len(rows_used) < 3:
        res.update(status="skip", reason="sheet too small")
        return res
    kind = rng.choice(["insert", "delete"])
    at = rng.choice(rows_used[1:])
    op = RowOp(kind, at, rng.randint(1, 3))
    mover = rng.choice(["ours", "theirs"])
    S, N = Editor(data), Editor(data)
    apply_row_op(S, sheet, part, op)
    if kind == "insert":
        for r in range(op.at, op.at + op.count):
            for c in rng.sample(range(1, 6), 2):
                S.set(part, f"{get_column_letter(c)}{r}", random_value(rng, False), rng=rng)
    # Both sides edit cells: N in base positions, S in its new positions.
    base_cells = list(base_vals.get(sheet, {}))
    n_edits = {c: random_value(rng) for c in rng.sample(base_cells, min(len(base_cells), rng.randint(2, 10)))}
    s_edits = {}
    for c in rng.sample(base_cells, min(len(base_cells), rng.randint(0, 5))):
        col, r = coordinate_from_string(c)
        if op(r) is not None:
            s_edits[f"{col}{op(r)}"] = random_value(rng)
    if n_edits and rng.random() < 0.3:  # sometimes both edit the same cell
        c = rng.choice(list(n_edits))
        col, r = coordinate_from_string(c)
        if op(r) is not None:
            s_edits[f"{col}{op(r)}"] = random_value(rng)
    for c, v in n_edits.items():
        N.set(part, c, v, rng=rng)
    for c, v in s_edits.items():
        S.set(part, c, v, rng=rng)
    sbytes, nbytes = S.save(), N.save()
    obytes, tbytes = (sbytes, nbytes) if mover == "ours" else (nbytes, sbytes)
    res.update(op=f"{mover} {kind} {op.count} at {sheet}!{op.at}", edits=len(n_edits) + len(s_edits))

    old_alarm = signal.signal(signal.SIGALRM, _timeout)
    limit = TIMEOUT + len(data) // 50_000
    signal.alarm(limit)
    try:
        with tempfile.TemporaryDirectory() as td:
            b, o, t = (os.path.join(td, n) for n in ("base.xlsx", "ours.xlsx", "theirs.xlsx"))
            for p_, blob in ((b, data), (o, obytes), (t, tbytes)):
                Path(p_).write_bytes(blob)
            try:
                mg = xlgit.Merger(b, o, t)
                merged = mg.run()
            except Timeout:
                raise Check("merge timed out", f">{limit}s")
            except Exception as e:
                raise Check("merge crashed", _where(e))
            res["detected"] = verify_rows(data, obytes, tbytes, merged, mg, sheet, op, base_problems)
        res.update(status="ok", conflicts=len(mg.conflicts))
    except Timeout:
        res.update(status="fail", reason="check timed out", detail=f">{limit}s")
    except Check as e:
        res.update(status="fail", reason=e.kind, detail=str(e.detail)[:600])
        if keep:
            d = Path(keep) / (Path(path).stem + "-rows")
            d.mkdir(parents=True, exist_ok=True)
            for n_, blob in (("base", data), ("ours", obytes), ("theirs", tbytes)):
                (d / f"{n_}.xlsx").write_bytes(blob)
            if "merged" in locals():
                (d / "merged.xlsx").write_bytes(merged)
    finally:
        signal.alarm(0)
        signal.signal(signal.SIGALRM, old_alarm)
    return res


class _Same:
    """No rows moved."""
    def __call__(self, r):
        return r
    start = end = __call__


class _Follow:
    """Where base rows went, worked out here from the rows xlgit matched
    (xlgit's own lookup isn't trusted): a matched row goes where it was
    matched, a deleted one nowhere, any other row moves with the nearest
    matched row above it."""

    def __init__(self, matched, deleted):
        self.matched, self.deleted = dict(matched), set(deleted)
        self.order = sorted(self.matched)

    def __call__(self, r):
        if r in self.deleted:
            return None
        if r in self.matched:
            return self.matched[r]
        above = [b for b in self.order if b < r]
        return r + (self.matched[above[-1]] - above[-1]) if above else r

    def start(self, r):
        while r in self.deleted:
            r += 1
        return self(r)

    def end(self, r):
        while r in self.deleted and r > 1:
            r -= 1
        return None if r in self.deleted else self(r)


def verify_rows(data, obytes, tbytes, merged, mg, sheet, op, known_problems):
    """Check the merged cells against a 3-way merge worked out here, following
    the rows xlgit recognised as moved (the true edit can be impossible to
    recognise, e.g. deleting one of many identical rows). Returns whether
    xlgit recognised the row change exactly."""
    worse = new_problems(known_problems, structural(merged, "merged"))
    if worse:
        raise Check(f"merged: {worse[0].split(':', 1)[0]}", "; ".join(worse)[:300])
    mp = mg.maps_b.get(sheet)
    fn = _Follow(mp.exact, mp.deleted) if mp else _Same()
    # xlgit builds on the branch whose rows moved; the other side's edits move onto it.
    startb, otherb = (tbytes, obytes) if mg.swapped else (obytes, tbytes)
    bv, sv, nv, mv = (sheet_values(xlgit.Package(x)) for x in (data, startb, otherb, merged))
    rw = lambda v, host: ("f", shift_formula(v[1], host, sheet, fn, fn.start, fn.end)) if mp and v and v[0] == "f" else v
    want, conflicts = {}, set()
    for name in sv:
        if name not in mv:
            raise Check("sheet lost", name)
        exp = dict(sv[name])
        b, n = bv.get(name, {}), nv.get(name, {})
        moves = name == sheet and mp is not None
        for c in set(b) | set(n):
            if same(n.get(c), b.get(c)):
                continue
            col, r = coordinate_from_string(c)
            r2 = fn(r) if moves else r
            if r2 is None:
                conflicts.add((name, f"{c} (row deleted)"))
                continue
            tgt = f"{col}{r2}"
            bx, nx, s_now = rw(b.get(c), name), rw(n.get(c), name), sv[name].get(tgt)
            if same(nx, s_now):
                continue
            if same(s_now, bx):
                exp[tgt] = nx
            else:
                conflicts.add((name, tgt))
                exp[tgt] = nx if mg.swapped else s_now  # your value wins a clash
            if exp.get(tgt) is None:
                exp.pop(tgt, None)
        want[name] = exp
    got_conflicts = {(s_, c_) for s_, c_, *_ in mg.conflicts if not str(c_).startswith("(")}
    if got_conflicts - conflicts:
        raise Check("false conflict", repr(sorted(got_conflicts - conflicts)[:3]))
    if conflicts - got_conflicts:
        raise Check("conflict not reported", repr(sorted(conflicts - got_conflicts)[:3]))
    for name, exp in want.items():
        got = mv[name]
        for c in set(exp) | set(got):
            if not same(exp.get(c), got.get(c)):
                raise Check("cell wrong after row change", f"{name}!{c} want={exp.get(c)} got={got.get(c)}")
    if set(mv) - set(sv) - {xlgit.CONFLICT_SHEET}:
        raise Check("sheet appeared", repr(set(mv) - set(sv)))
    base_rows = {coordinate_from_string(c)[1] for c in bv.get(sheet, {})}
    return mp is not None and all(fn(r) == op(r) for r in base_rows)


def libreoffice_resave(data):
    """Open and save the workbook in LibreOffice Calc: a different writer,
    which renumbers parts and rewrites strings, styles and formulas."""
    import subprocess
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "ours.xlsx")
        Path(src).write_bytes(data)
        out = os.path.join(td, "out")
        subprocess.run(["soffice", "--headless", f"-env:UserInstallation=file://{td}/profile",
                        "--convert-to", "xlsx:Calc Office Open XML", "--outdir", out, src],
                       check=True, capture_output=True, timeout=120)
        return Path(out, "ours.xlsx").read_bytes()


def verify_three_way(data, obytes, tbytes, merged, mg, known_problems):
    """Check every cell against a 3-way merge computed from the raw XML
    values of base, ours and theirs (for when ours changed in ways this file
    didn't plan, such as a re-save by another app)."""
    worse = new_problems(known_problems, structural(merged, "merged"))
    if worse:
        raise Check(f"merged: {worse[0].split(':', 1)[0]}", "; ".join(worse)[:300])
    bpkg, opkg, tpkg, mpkg = (xlgit.Package(x) for x in (data, obytes, tbytes, merged))
    bv, ov, tv, mv = (sheet_values(p) for p in (bpkg, opkg, tpkg, mpkg))
    pivots = collections.defaultdict(list)  # a pivot's cells never conflict in any version
    for pk in (bpkg, opkg, tpkg):
        for n, _, part in pk.sheets():
            if part:
                pivots[n] += xlgit.pivot_locations(pk, part)
    conflicts = {(s, c) for s, c, *_ in mg.conflicts}
    for sheet in ov:
        if sheet not in mv:
            raise Check("sheet lost", sheet)
        b, o, t, got = bv.get(sheet, {}), ov[sheet], tv.get(sheet, {}), mv[sheet]
        for coord in set(b) | set(o) | set(t) | set(got):
            x, y, z = b.get(coord), o.get(coord), t.get(coord)
            if ("sharedref",) in (x, y, z) or any(xlgit.overlaps(coord, r) for r in pivots.get(sheet, [])):
                continue
            conflict = not (same(z, x) or same(y, x) or same(y, z))
            want = t.get(coord) if same(y, x) and not same(z, x) else y
            if conflict != ((sheet, coord) in conflicts):
                raise Check("conflict not reported" if conflict else "false conflict",
                            f"{sheet}!{coord} base={x} ours={y} theirs={z}")
            if not same(got.get(coord), want):
                raise Check("edit lost" if same(got.get(coord), y) else "cell written wrong",
                            f"{sheet}!{coord} base={x} ours={y} theirs={z} got={got.get(coord)}")
    oobj, mobj = xlgit.describe_objects(opkg), xlgit.describe_objects(mpkg)
    for sheet, items in oobj.items():
        missing = collections.Counter(items) - collections.Counter(mobj.get(sheet, []))
        if missing:
            raise Check("object lost or changed", f"{sheet}: {list(missing)[:2]}")


def _only_used_by(pkg, part, sheets):
    """Is part reachable only through the given sheet parts?"""
    seen, stack = set(), [""]
    while stack:
        owner = stack.pop()
        for _, _, t, ext in pkg.rels(owner):
            if not ext and t in pkg.parts and t not in seen and t not in sheets:
                seen.add(t)
                stack.append(t)
    return part not in seen


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
    path, seed, keep, writer, mode = args
    try:
        if mode == "rows":
            return run_rows(path, seed, keep)
        return run_one(path, seed, keep, writer)
    except Exception as e:
        return {"file": str(path), "seed": seed, "status": "fail", "reason": "harness error", "detail": _where(e)}


def signature(r):
    d = r.get("detail", "")
    if r["reason"] in ("merge crashed", "diff crashed", "openpyxl cannot read merged", "xlgit cannot read merged",
                       "xlgit cannot read input", "harness error"):
        return f"{r['reason']}: {re.sub(r': .*? @', ' @', d)}"
    return r["reason"]


def _cap_memory(gigabytes):
    """A file that blows up memory should fail its own run with MemoryError,
    not get a worker killed by the system (the pool would wait for it forever)."""
    try:
        import resource
        limit = int(gigabytes * 1024 ** 3)
        resource.setrlimit(resource.RLIMIT_AS, (limit, limit))
    except (ImportError, ValueError, OSError):
        pass


def pool_context():
    """Workers are replaced every 50 files to cap memory. Forking a
    replacement from the pool's helper thread can copy a held queue lock
    into it and hang the run, so start workers from a clean server process."""
    methods = multiprocessing.get_all_start_methods()
    return multiprocessing.get_context("forkserver" if "forkserver" in methods else "spawn")


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
    ap.add_argument("--mode", choices=["cells", "rows"], default="cells",
                    help="cells: random cell and sheet edits; rows: one branch inserts or deletes rows")
    ap.add_argument("--writer", choices=["libreoffice"],
                    help="re-save ours in another app before merging (needs soffice)")
    ap.add_argument("--max-memory", type=float, default=4.0, help="GB per worker (default 4)")
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
    jobs = [(str(f), a.seed + i, a.keep, a.writer, a.mode) for f in files for i in range(a.rounds)]
    stats, sigs, done = collections.Counter(), collections.defaultdict(list), []
    with open(a.out, "w") as out, pool_context().Pool(a.jobs, _cap_memory, (a.max_memory,), maxtasksperchild=50) as pool:
        for i, r in enumerate(pool.imap_unordered(_job, jobs, chunksize=4), 1):
            out.write(json.dumps(r) + "\n")
            out.flush()
            done.append(r)
            stats[r["status"]] += 1
            if r["status"] == "fail":
                sigs[signature(r)].append(r)
            elif r["status"] == "skip":
                stats["skip: " + r["reason"]] += 1
            if i % 200 == 0:
                print(f"{i}/{len(jobs)} {dict((k, v) for k, v in stats.items() if ':' not in k)}", file=sys.stderr)
    print(f"\n{len(jobs)} runs: " + ", ".join(f"{k} {v}" for k, v in sorted(stats.items())))
    if a.mode == "rows":
        found = sum(1 for r in done if r.get("detected"))
        print(f"row change recognised exactly in {found} of {sum(1 for r in done if r['status'] == 'ok')} passing runs")
    for sig, rs in sorted(sigs.items(), key=lambda kv: -len(kv[1])):
        print(f"\n[{len(rs)}] {sig}")
        for r in rs[:3]:
            print(f"    {os.path.basename(r['file'])} (seed {r['seed']}): {r.get('detail', '')[:200]}")
    return 1 if sigs else 0


if __name__ == "__main__":
    sys.exit(main())
