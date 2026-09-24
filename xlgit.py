"""xlgit: make Excel workbooks behave like code in git and GitHub.

Commands:
  textconv FILE                    print a workbook as diffable text (git diff driver)
  diff OLD NEW [--markdown]        cell and object (chart/image/comment) diff
  merge BASE OURS THEIRS [PATH]    3-way merge (git merge driver)
  install                          wire the drivers into the current git repo

The merge edits the workbook's XML parts in place instead of re-saving it
through a spreadsheet library, so charts, images, formatting, comments and
macros in OUR copy survive untouched, and their edits to those objects are
carried over part by part.
"""
import copy
import datetime
import hashlib
import io
import os
import posixpath
import re
import subprocess
import sys
import zipfile
from collections import namedtuple

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, get_column_letter
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, to_excel
from openpyxl.worksheet.formula import ArrayFormula

EXTS = ("*.xlsx", "*.xlsm")

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CTYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAW = "http://schemas.openxmlformats.org/drawingml/2006/main"
RID = f"{{{REL}}}id"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
WORKSHEET_REL = REL + "/worksheet"
WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
CONFLICT_SHEET = "_merge_conflicts"

# Worksheet child order from the OOXML schema; new elements must slot in here.
WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
            "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
            "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
            "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
            "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
            "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
            "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
            "controls", "webPublishItems", "tableParts", "extLst"]
WB_ORDER = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews",
            "sheets", "functionGroups", "externalReferences", "definedNames", "calcPr",
            "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes",
            "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"]

# Parts merged some other way (cells, sheet list, names) or that every save
# rewrites; everything else is merged whole, part by part.
NOT_OBJECT_PART = re.compile(
    r"^(\[Content_Types\]\.xml|docProps/.*|customXml/.*|xl/workbook\.xml|xl/worksheets/.*"
    r"|xl/sharedStrings\.xml|xl/calcChain\.xml|xl/styles\.xml|xl/theme/.*|xl/metadata\.xml"
    r"|xl/persons/.*|xl/printerSettings/.*|.*\.rels)$")

ArrayF = namedtuple("ArrayF", "ref text")
Opaque = namedtuple("Opaque", "desc")


def m(tag):
    return f"{{{MAIN}}}{tag}"


def local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def digest(data):
    return hashlib.sha1(data or b"").hexdigest()[:8]


# ---------- package (zip of XML parts) helpers ----------

def rels_name(owner):
    if owner == "":
        return "_rels/.rels"
    d, b = posixpath.split(owner)
    return f"{d}/_rels/{b}.rels" if d else f"_rels/{b}.rels"


def resolve(owner, target):
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))


def relative(target, owner):
    d = posixpath.dirname(owner)
    return posixpath.relpath(target, d) if d else target


def reachable(parts):
    """Every part reachable from the package root through relationships."""
    seen, stack = set(), [""]
    while stack:
        owner = stack.pop()
        rn = rels_name(owner)
        if rn not in parts:
            continue
        for r in etree.fromstring(parts[rn]):
            if r.get("TargetMode") == "External":
                continue
            t = resolve(owner, r.get("Target"))
            if t in parts and t not in seen:
                seen.add(t)
                stack.append(t)
    return seen


class Package:
    """Read-only raw view of an .xlsx zip."""

    def __init__(self, data=b""):
        self.parts, self.order = {}, []
        if data:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for info in z.infolist():
                    if not info.is_dir():
                        self.parts[info.filename] = z.read(info)
                        self.order.append(info.filename)
        self.workbook = next((t for _, typ, t, _ in self.rels("") if typ.endswith("/officeDocument")),
                             "xl/workbook.xml")

    @classmethod
    def open(cls, path):
        if not path or not os.path.exists(path) or os.path.getsize(path) == 0:
            return cls()
        with open(path, "rb") as f:
            return cls(f.read())

    def xml(self, name):
        return etree.fromstring(self.parts[name]) if name in self.parts else None

    def rels(self, owner):
        """[(rId, type, target, external)]; internal targets are absolute part names."""
        root = self.xml(rels_name(owner))
        if root is None:
            return []
        out = []
        for r in root:
            ext = r.get("TargetMode") == "External"
            out.append((r.get("Id"), r.get("Type"), r.get("Target") if ext else resolve(owner, r.get("Target")), ext))
        return out

    def sheets(self):
        """[(name, sheetId, part)] in workbook order."""
        wb = self.xml(self.workbook)
        if wb is None:
            return []
        targets = {rid: t for rid, _, t, _ in self.rels(self.workbook)}
        return [(s.get("name"), s.get("sheetId"), targets.get(s.get(RID)))
                for s in wb.find(m("sheets"))]

    def content_type(self, name):
        ct = self.xml("[Content_Types].xml")
        for o in ct.iter(f"{{{CTYPES}}}Override"):
            if o.get("PartName").lstrip("/").lower() == name.lower():
                return ("override", o.get("ContentType"))
        ext = name.rsplit(".", 1)[-1].lower()
        for d in ct.iter(f"{{{CTYPES}}}Default"):
            if d.get("Extension").lower() == ext:
                return ("default", ext, d.get("ContentType"))
        return None

    def shared_strings(self):
        for _, typ, t, _ in self.rels(self.workbook):
            if typ.endswith("/sharedStrings") and t in self.parts:
                return list(self.xml(t).iter(m("si")))
        return []

    def styles_bytes(self):
        return next((self.parts.get(t) for _, typ, t, _ in self.rels(self.workbook) if typ.endswith("/styles")), None)

    def sheet_object(self, part, key):
        """Target part of a sheet's drawing / legacyDrawing (element-wired) or comments (rel-only)."""
        if not part or part not in self.parts:
            return None, None
        rels = {rid: (typ, t) for rid, typ, t, ext in self.rels(part) if not ext}
        if key == "comments":
            return next(((t, typ) for typ, t in rels.values() if typ.endswith("/comments")), (None, None))
        el = self.xml(part).find(m(key))
        if el is None:
            return None, None
        typ, t = rels.get(el.get(RID), (None, None))
        return t, typ


# ---------- reading cells ----------

def _load(path):
    """openpyxl refuses files without an Excel extension, but git hands its
    drivers temp files like .merge_file_a1b2c3, so read through a byte stream."""
    with open(path, "rb") as f:
        return load_workbook(io.BytesIO(f.read()))


def _norm(v):
    if isinstance(v, ArrayFormula):
        return ArrayF(v.ref, v.text)
    if v is not None and not isinstance(v, (str, int, float, bool, datetime.date, datetime.time, datetime.timedelta)):
        return Opaque(f"{type(v).__name__}")
    return v


def read_cells(path):
    """Return {sheet: {coord: value}} where formulas stay as '=...' strings."""
    if not path or not os.path.exists(path) or os.path.getsize(path) == 0:
        return {}
    wb = _load(path)
    out = {}
    for ws in wb.worksheets:
        out[ws.title] = {c.coordinate: _norm(c.value)
                         for row in ws.iter_rows() for c in row if c.value is not None}
    return out


def fmt(v):
    if v is None:
        return "(empty)"
    if isinstance(v, ArrayF):
        return "{" + v.text + "}"
    if isinstance(v, Opaque):
        return f"<{v.desc}>"
    return repr(v) if isinstance(v, str) and not v.startswith("=") else str(v)


def _cell_sort_key(coord):
    col, row = coordinate_from_string(coord)
    return (row, column_index_from_string(col))


# ---------- describing charts, images, comments ----------

def _strip_chart_caches(data):
    """Chart XML embeds cached copies of the cell values it plots; drop them so
    a cell edit doesn't look like a chart edit."""
    root = etree.fromstring(data)
    for el in list(root.iter(f"{{{CHART}}}numCache", f"{{{CHART}}}strCache")):
        el.getparent().remove(el)
    return etree.tostring(root)


def comparable(name, data):
    if data is not None and re.match(r"xl/charts/chart\d*\.xml$", name):
        return _strip_chart_caches(data)
    return data


def _chart_summary(pkg, name):
    root = pkg.xml(name)
    plot = root.find(f".//{{{CHART}}}plotArea")
    kinds = [local(e) for e in plot if local(e).endswith("Chart")] if plot is not None else []
    title_el = root.find(f"{{{CHART}}}chart/{{{CHART}}}title")
    title = "".join(t.text or "" for t in title_el.iter(f"{{{DRAW}}}t")) if title_el is not None else ""
    refs = [f.text for f in root.iter(f"{{{CHART}}}f") if f.text]
    return (f"chart: {'+'.join(kinds) or 'chart'} {title!r} data={','.join(refs)} "
            f"#{digest(_strip_chart_caches(pkg.parts[name]))}")


def describe_objects(pkg):
    """{sheet: [one line per chart / image / comment / table / pivot]}."""
    out = {}
    for sheet, _, part in pkg.sheets():
        items = []
        for _, typ, tgt, ext in pkg.rels(part):
            if ext or tgt not in pkg.parts:
                continue
            kind = typ.rsplit("/", 1)[-1]
            if kind == "drawing":
                for _, t2, tgt2, ext2 in pkg.rels(tgt):
                    if ext2 or tgt2 not in pkg.parts:
                        continue
                    if t2.endswith("/chart"):
                        items.append(_chart_summary(pkg, tgt2))
                    elif t2.endswith("/image"):
                        items.append(f"image: {posixpath.basename(tgt2)} #{digest(pkg.parts[tgt2])}")
            elif kind == "comments":
                for c in pkg.xml(tgt).iter(m("comment")):
                    text = "".join(t.text or "" for t in c.iter(m("t")))
                    items.append(f"comment {c.get('ref')}: {text!r}")
            elif kind == "table":
                t = pkg.xml(tgt)
                items.append(f"table: {t.get('displayName')} {t.get('ref')}")
            elif kind == "pivotTable":
                items.append(f"pivot table: {pkg.xml(tgt).get('name')} #{digest(pkg.parts[tgt])}")
        out[sheet] = items
    return out


# ---------- textconv ----------

def textconv(path):
    objects = describe_objects(Package.open(path))
    for sheet, cells in read_cells(path).items():
        print(f"=== sheet: {sheet} ===")
        for coord, v in cells.items():
            print(f"{sheet}!{coord}\t{fmt(v)}")
        for item in objects.get(sheet, []):
            print(f"{sheet} [{item}]")


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


def diff_objects(old, new):
    for sheet in sorted(set(old) | set(new)):
        o, n = list(old.get(sheet, [])), list(new.get(sheet, []))
        for item in o:
            if item in n:
                n.remove(item)
            else:
                yield ("object removed", sheet, "", item, None)
        for item in n:
            yield ("object added", sheet, "", None, item)


def diff(old_path, new_path, markdown=False, title=None):
    changes = list(diff_cells(read_cells(old_path), read_cells(new_path)))
    changes += list(diff_objects(describe_objects(Package.open(old_path)),
                                 describe_objects(Package.open(new_path))))
    changes = [(k, s, c, o, n, k.startswith("object")) for k, s, c, o, n in changes]
    show = lambda v, raw: "" if v is None else v if raw else fmt(v)
    if markdown:
        print(f"### {title or new_path}\n")
        if not changes:
            print("No cell or object changes (formatting may still differ).\n")
            return 0
        print(f"{len(changes)} change(s)\n")
        print("| Sheet | Cell | Change | Before | After |")
        print("|---|---|---|---|---|")
        for kind, sheet, coord, ov, nv, raw in changes[:500]:
            esc = lambda v: show(v, raw).replace("|", "\\|")
            print(f"| {sheet} | {coord} | {kind} | {esc(ov)} | {esc(nv)} |")
        if len(changes) > 500:
            print(f"\n...and {len(changes) - 500} more.")
        print()
    else:
        for kind, sheet, coord, ov, nv, raw in changes:
            where = f"{sheet}!{coord}" if coord else sheet
            print(f"{kind:14} {where}  {show(ov, raw)} -> {show(nv, raw)}")
    return 1 if changes else 0


# ---------- merge ----------

def _pair(a, b):
    """Map sheet names in a to names in b: same name first, then same sheetId
    (Excel keeps sheetId when a sheet is renamed)."""
    names_b = {n for n, _, _ in b}
    out = {n: n for n, _, _ in a if n in names_b}
    used = set(out.values())
    by_id = {sid: n for n, sid, _ in b if n not in used}
    for n, sid, _ in a:
        if n not in out and sid in by_id and by_id[sid] not in used:
            out[n] = by_id[sid]
            used.add(by_id[sid])
    return out


class Merger:
    """3-way merge of BASE/OURS/THEIRS; the result starts as a copy of OURS'
    zip and only the XML that has to change gets rewritten."""

    def __init__(self, base_path, ours_path, theirs_path):
        self.b, self.o, self.t = (Package.open(p) for p in (base_path, ours_path, theirs_path))
        self.bc, self.oc, self.tc = (read_cells(p) for p in (base_path, ours_path, theirs_path))
        self.res = dict(self.o.parts)
        self.order = list(self.o.order)
        self.trees = {}
        self.imported = {}  # their part name -> part name in the result
        self.conflicts = []
        self.cells_taken = 0
        self.objects_taken = []
        self.wb = self.o.workbook
        self.same_styles = self.o.styles_bytes() == self.t.styles_bytes()
        wbpr = self.o.xml(self.wb).find(m("workbookPr"))
        self.epoch = CALENDAR_MAC_1904 if wbpr is not None and wbpr.get("date1904") in ("1", "true") \
            else CALENDAR_WINDOWS_1900

    # --- result tree plumbing ---

    def tree(self, name, create=None):
        if name not in self.trees:
            if name in self.res:
                self.trees[name] = etree.fromstring(self.res[name])
            elif create is not None:
                self.trees[name] = create
                self.put(name, b"")
            else:
                return None
        return self.trees[name]

    def put(self, name, data):
        self.res[name] = data
        self.trees.pop(name, None) if data else None
        if name not in self.order:
            self.order.append(name)

    def drop(self, name):
        self.res.pop(name, None)
        self.trees.pop(name, None)
        if name in self.order:
            self.order.remove(name)
        ct = self.tree("[Content_Types].xml")
        for o in list(ct.iter(f"{{{CTYPES}}}Override")):
            if o.get("PartName").lstrip("/") == name:
                ct.remove(o)

    def rels_tree(self, owner):
        return self.tree(rels_name(owner), create=etree.Element(f"{{{PKG_REL}}}Relationships", nsmap={None: PKG_REL}))

    def add_rel(self, owner, typ, target, external=False):
        rt = self.rels_tree(owner)
        ids = {r.get("Id") for r in rt}
        n = 1
        while f"rId{n}" in ids:
            n += 1
        r = etree.SubElement(rt, f"{{{PKG_REL}}}Relationship")
        r.set("Id", f"rId{n}")
        r.set("Type", typ)
        r.set("Target", target if external else relative(target, owner))
        if external:
            r.set("TargetMode", "External")
        return f"rId{n}"

    def remove_rel(self, owner, rid):
        rt = self.rels_tree(owner)
        for r in list(rt):
            if r.get("Id") == rid:
                rt.remove(r)

    def copy_content_type(self, their_name, our_name):
        info = self.t.content_type(their_name)
        if not info:
            return
        ct = self.tree("[Content_Types].xml")
        if info[0] == "override":
            for o in ct.iter(f"{{{CTYPES}}}Override"):
                if o.get("PartName").lstrip("/") == our_name:
                    o.set("ContentType", info[1])
                    return
            o = etree.SubElement(ct, f"{{{CTYPES}}}Override")
            o.set("PartName", "/" + our_name)
            o.set("ContentType", info[1])
        elif not any(d.get("Extension").lower() == info[1] for d in ct.iter(f"{{{CTYPES}}}Default")):
            d = etree.Element(f"{{{CTYPES}}}Default")
            d.set("Extension", info[1])
            d.set("ContentType", info[2])
            ct.insert(0, d)

    def unique_name(self, name):
        d, b = posixpath.split(name)
        stem, num, ext = re.match(r"(.*?)(\d*)(\.[^.]*)?$", b).groups()
        n = int(num or 1)
        while True:
            n += 1
            cand = posixpath.join(d, f"{stem}{n}{ext or ''}")
            if cand not in self.res and cand not in self.t.parts and cand not in self.b.parts:
                return cand

    def import_part(self, their_name, keep_rel=None):
        """Copy one of their parts (and everything it links to) into the result,
        renaming anything that would collide with a part we already have."""
        if their_name in self.imported:
            return self.imported[their_name]
        ours = their_name if their_name not in self.res else self.unique_name(their_name)
        self.imported[their_name] = ours
        self.put(ours, self.t.parts[their_name])
        self.copy_content_type(their_name, ours)
        self.copy_rels(their_name, ours, keep_rel)
        return ours

    def copy_rels(self, their_owner, our_owner, keep_rel=None):
        """Their relationships for a part, retargeted into the result.
        Returns rIds that keep_rel rejected (the caller strips their use)."""
        rn = rels_name(our_owner)
        rels = self.t.rels(their_owner)
        self.trees.pop(rn, None)
        if not rels:
            if rn in self.res:
                self.drop(rn)
            return []
        rt = etree.Element(f"{{{PKG_REL}}}Relationships", nsmap={None: PKG_REL})
        dropped = []
        for rid, typ, tgt, ext in rels:
            if keep_rel and not ext and not keep_rel(typ):
                dropped.append((rid, typ))
                continue
            r = etree.SubElement(rt, f"{{{PKG_REL}}}Relationship")
            r.set("Id", rid)
            r.set("Type", typ)
            if ext:
                r.set("Target", tgt)
                r.set("TargetMode", "External")
                continue
            if tgt in self.b.parts and tgt in self.res and tgt not in self.imported.values():
                ours = tgt  # the same part on both branches
            elif tgt in self.t.parts:
                ours = self.import_part(tgt)
            else:
                ours = tgt
            r.set("Target", relative(ours, our_owner))
        self.trees[rn] = rt
        self.put(rn, b"")
        return dropped

    def sheet_map(self):
        """{name: (sheet element, part)} for the result as it stands now."""
        wb = self.tree(self.wb)
        targets = {r.get("Id"): resolve(self.wb, r.get("Target")) for r in self.rels_tree(self.wb)}
        return {s.get("name"): (s, targets.get(s.get(RID))) for s in wb.find(m("sheets"))}

    def conflict(self, sheet, cell, base, ours, theirs):
        self.conflicts.append((sheet, cell, base, ours, theirs))

    # --- the merge ---

    def run(self):
        bs, os_, ts = self.b.sheets(), self.o.sheets(), self.t.sheets()
        bo, bt = _pair(bs, os_), _pair(bs, ts)
        self.lineage_o = {v: k for k, v in bo.items()}
        self.lineage_t = {v: k for k, v in bt.items()}
        self.sheet_el = {n: el for n, (el, _) in self.sheet_map().items()}  # ours, before any change
        self.new_sheet_el = {}
        self.capture_names()

        done_t = set()
        for bname, _, bpart in bs:
            o, t = bo.get(bname), bt.get(bname)
            done_t.add(t)
            if o and t:
                self.merge_sheet(bname, o, t)
                self.merge_sheet_name(bname, o, t)
            elif o:
                if self.oc.get(o) == self.bc.get(bname):
                    self.delete_sheet(o)
                else:
                    self.conflict(o, "(sheet)", "exists", "edited", "deleted")
            elif t and self.tc.get(t) != self.bc.get(bname):
                self.conflict(t, "(sheet)", "exists", "deleted", "edited")
        ours_from_base = set(bo.values())
        for tname, _, tpart in ts:
            if tname in done_t:
                continue
            if tname in self.sheet_el and tname not in ours_from_base:
                self.merge_sheet(None, tname, tname)  # both branches added a sheet with this name
            elif tname in self.sheet_map():
                self.conflict(tname, "(sheet)", "", "exists", "added with the same name")
            else:
                self.import_sheet(tname, tpart)

        self.merge_names()
        self.merge_object_parts()
        self.report_unmerged_additions()
        self.write_conflict_sheet()
        return self.finish()

    # --- cells ---

    def merge_sheet(self, bname, oname, tname):
        b = self.bc.get(bname, {}) if bname else {}
        o, t = self.oc.get(oname, {}), self.tc.get(tname, {})
        edits = {}
        for coord in set(b) | set(t):
            bv, ov, tv = b.get(coord), o.get(coord), t.get(coord)
            if tv == bv or tv == ov:
                continue
            if ov == bv and not isinstance(tv, Opaque):
                edits[coord] = tv
            else:
                self.conflict(oname, coord, bv, ov, tv)
        if edits:
            self.edit_cells(oname, tname, edits)
        bpart = dict((n, p) for n, _, p in self.b.sheets()).get(bname)
        opart = dict((n, p) for n, _, p in self.o.sheets()).get(oname)
        tpart = dict((n, p) for n, _, p in self.t.sheets()).get(tname)
        for key in ("drawing", "legacyDrawing", "comments"):
            self.merge_sheet_object(oname, key, bpart, opart, tpart)

    def edit_cells(self, oname, tname, edits):
        part = self.sheet_map()[oname][1]
        root = self.tree(part)
        sd = root.find(m("sheetData"))
        self.unshare_formulas(root, self.oc.get(oname, {}))
        # Style numbers index into styles.xml, so theirs only carry over when
        # both branches have the same styles.xml, and only if we didn't restyle.
        their_styles, base_styles = {}, None
        if self.same_styles:
            tpart = dict((n, p) for n, _, p in self.t.sheets())[tname]
            their_styles = {c.get("r"): c.get("s") for c in self.t.xml(tpart).iter(m("c"))}
            bpart = dict((n, p) for n, _, p in self.b.sheets()).get(self.lineage_o.get(oname))
            if bpart and self.b.styles_bytes() == self.o.styles_bytes():
                base_styles = {c.get("r"): c.get("s") for c in self.b.xml(bpart).iter(m("c"))}
        rows = {}
        last = 0
        for row in sd.findall(m("row")):
            last = int(row.get("r") or last + 1)
            rows[last] = row
        for coord, value in sorted(edits.items(), key=lambda kv: _cell_sort_key(kv[0])):
            col, rnum = coordinate_from_string(coord)
            row = self.get_row(sd, rows, rnum)
            row.attrib.pop("spans", None)
            c = self.get_cell(row, coord, column_index_from_string(col))
            restyled = base_styles is not None and c.get("s") != base_styles.get(coord)
            self.write_value(c, value)
            if self.same_styles and not restyled:
                s = their_styles.get(coord)
                c.set("s", s) if s else c.attrib.pop("s", None)
            if value is None and c.get("s") is None and len(c) == 0:
                row.remove(c)
            self.cells_taken += 1
        for rnum, row in rows.items():
            if len(row) == 0 and set(row.attrib) <= {"r", "spans"}:
                sd.remove(row)
        self.update_dimension(root)

    def unshare_formulas(self, root, values):
        """Shared formulas are stored once and implied for a range; editing one
        cell of the range would break the rest, so spell each one out."""
        for c in root.iter(m("c")):
            f = c.find(m("f"))
            if f is not None and f.get("t") == "shared":
                v = values.get(c.get("r"))
                if isinstance(v, str) and v.startswith("="):
                    f.text = v[1:]
                    for a in ("t", "ref", "si"):
                        f.attrib.pop(a, None)

    @staticmethod
    def get_row(sd, rows, rnum):
        row = rows.get(rnum)
        if row is None:
            row = etree.SubElement(sd, m("row"))
            row.set("r", str(rnum))
            later = [rows[n] for n in sorted(rows) if n > rnum]
            if later:
                later[0].addprevious(row)
            rows[rnum] = row
        return row

    @staticmethod
    def get_cell(row, coord, colnum):
        for c in row.findall(m("c")):
            if c.get("r") == coord:
                return c
        c = etree.SubElement(row, m("c"))
        c.set("r", coord)
        for other in row.findall(m("c")):
            if other is not c and column_index_from_string(coordinate_from_string(other.get("r"))[0]) > colnum:
                other.addprevious(c)
                break
        return c

    def write_value(self, c, v, as_text=False):
        for child in list(c):
            if local(child) in ("f", "v", "is"):
                c.remove(child)
        for a in ("t", "cm", "vm"):
            c.attrib.pop(a, None)
        if v is None:
            return
        new = []
        if isinstance(v, ArrayF):
            f = etree.SubElement(c, m("f"))
            f.set("t", "array")
            f.set("ref", v.ref)
            f.text = v.text[1:] if v.text.startswith("=") else v.text
            new.append(f)
        elif isinstance(v, bool):
            c.set("t", "b")
            new.append(self._sub(c, "v", "1" if v else "0"))
        elif isinstance(v, (int, float)):
            new.append(self._sub(c, "v", repr(v)))
        elif isinstance(v, (datetime.date, datetime.time, datetime.timedelta)):
            new.append(self._sub(c, "v", repr(to_excel(v, self.epoch))))
        elif v.startswith("=") and len(v) > 1 and not as_text:
            new.append(self._sub(c, "f", v[1:]))
        elif v in ERRORS and not as_text:
            c.set("t", "e")
            new.append(self._sub(c, "v", v))
        else:
            c.set("t", "inlineStr")
            is_ = etree.SubElement(c, m("is"))
            t = etree.SubElement(is_, m("t"))
            t.text = v
            t.set(XML_SPACE, "preserve")
            new.append(is_)
        for i, el in enumerate(new):
            c.insert(i, el)

    @staticmethod
    def _sub(parent, tag, text):
        el = etree.SubElement(parent, m(tag))
        el.text = text
        return el

    @staticmethod
    def update_dimension(root):
        dim = root.find(m("dimension"))
        coords = [coordinate_from_string(c.get("r")) for c in root.iter(m("c")) if c.get("r")]
        if dim is None or not coords:
            return
        cols = [column_index_from_string(col) for col, _ in coords]
        rows = [r for _, r in coords]
        lo, hi = f"{get_column_letter(min(cols))}{min(rows)}", f"{get_column_letter(max(cols))}{max(rows)}"
        dim.set("ref", lo if lo == hi else f"{lo}:{hi}")

    # --- charts, images and comments hanging off a sheet ---

    def merge_sheet_object(self, oname, key, bpart, opart, tpart):
        b = self.b.sheet_object(bpart, key)[0]
        o = self.o.sheet_object(opart, key)[0]
        t, ttype = self.t.sheet_object(tpart, key)
        if t == b or (t is None and o is None):
            return
        if o != b:
            if t is not None:
                self.conflict(oname, f"({key})", "", "changed", "changed; theirs not merged")
            return
        part = self.sheet_map()[oname][1]
        root = self.tree(part)
        old_el = root.find(m(key)) if key != "comments" else None
        for r in list(self.rels_tree(part)):
            if (old_el is not None and r.get("Id") == old_el.get(RID)) or \
                    (key == "comments" and r.get("Type").endswith("/comments")):
                self.rels_tree(part).remove(r)
        if old_el is not None:
            root.remove(old_el)
        if t is None:
            return
        target = t if (t in self.b.parts and t in self.res) else self.import_part(t)
        rid = self.add_rel(part, ttype, target)
        if key != "comments":
            el = etree.SubElement(root, m(key))
            el.set(RID, rid)
            self.place(root, el, WS_ORDER)
        self.objects_taken.append(f"{oname}: {key} from their branch")

    @staticmethod
    def place(root, el, order):
        """Move el to its schema-mandated position among root's children.
        Extension elements (mc:AlternateContent etc.) are assumed to sit where
        Excel writes them: after workbookPr, or among a sheet's oleObjects."""
        idx = order.index(local(el))
        unknown = order.index("workbookPr" if "workbookPr" in order else "oleObjects")
        for child in root:
            if child is el:
                continue
            name = local(child)
            pos = order.index(name) if name in order else unknown
            if pos > idx:
                child.addprevious(el)
                return

    # --- whole sheets ---

    def merge_sheet_name(self, bname, oname, tname):
        if tname in (bname, oname):
            return
        if oname != bname:
            self.conflict(oname, "(sheet name)", bname, oname, tname)
        elif tname in self.sheet_map():
            self.conflict(oname, "(sheet name)", bname, oname, f"{tname} (name already taken)")
        else:
            self.sheet_el[oname].set("name", tname)

    def delete_sheet(self, oname):
        el = self.sheet_el[oname]
        self.remove_rel(self.wb, el.get(RID))
        el.getparent().remove(el)

    def add_sheet(self, name, part):
        sheets = self.tree(self.wb).find(m("sheets"))
        sid = max([int(s.get("sheetId")) for s in sheets] + [0]) + 1
        el = etree.SubElement(sheets, m("sheet"))
        el.set("name", name)
        el.set("sheetId", str(sid))
        el.set(RID, self.add_rel(self.wb, WORKSHEET_REL, part))
        return el

    def import_sheet(self, tname, tpart):
        """A sheet only they added: copy it whole, charts included. Shared
        strings become inline strings since the string table isn't merged."""
        root = self.t.xml(tpart)
        sst = self.t.shared_strings()
        for c in root.iter(m("c")):
            for a in ("cm", "vm"):
                c.attrib.pop(a, None)
            if c.get("t") == "s":
                v = c.find(m("v"))
                si = sst[int(v.text)]
                c.remove(v)
                c.set("t", "inlineStr")
                is_ = etree.SubElement(c, m("is"))
                for ch in si:
                    is_.append(copy.deepcopy(ch))
        if not self.same_styles:  # their style numbers mean nothing in our styles.xml
            for el in root.iter(m("c"), m("row"), m("col")):
                for a in ("s", "style", "customFormat"):
                    el.attrib.pop(a, None)
            for tag in ("conditionalFormatting", "extLst"):
                for el in root.findall(m(tag)):
                    root.remove(el)
        ours = tpart if tpart not in self.res else self.unique_name(tpart)
        self.imported[tpart] = ours
        self.put(ours, b"")
        self.trees[ours] = root
        self.copy_content_type(tpart, ours)
        keep = ("/drawing", "/vmlDrawing", "/comments", "/hyperlink", "/printerSettings")
        dropped = self.copy_rels(tpart, ours, keep_rel=lambda typ: typ.endswith(keep))
        for rid, typ in dropped:
            for el in [e for e in root.iter() if e.get(RID) == rid]:
                parent = el.getparent()
                parent.remove(el)
                if local(parent) == "tableParts":
                    parent.set("count", str(len(parent)))
                    if len(parent) == 0:
                        root.remove(parent)
            self.conflict(tname, f"({typ.rsplit('/', 1)[-1]})", "", "", "on their new sheet; not merged")
        self.new_sheet_el[tname] = self.add_sheet(tname, ours)
        self.objects_taken.append(f"{tname}: new sheet from their branch")

    # --- defined names ---

    def _names(self, pkg, lineage):
        wb = pkg.xml(pkg.workbook)
        if wb is None:
            return {}
        sheets = [s.get("name") for s in wb.find(m("sheets"))]
        dn = wb.find(m("definedNames"))
        out = {}
        for d in (dn if dn is not None else []):
            lid = d.get("localSheetId")
            scope = sheets[int(lid)] if lid is not None and int(lid) < len(sheets) else None
            scope = lineage.get(scope, scope) if scope else None
            attrs = tuple(sorted((k, v) for k, v in d.attrib.items() if k != "localSheetId"))
            out[(d.get("name"), scope)] = (d.text, attrs)
        return out

    def capture_names(self):
        wb = self.tree(self.wb)
        sheets = list(wb.find(m("sheets")))
        self.name_els = {}
        dn = wb.find(m("definedNames"))
        for d in (dn if dn is not None else []):
            lid = d.get("localSheetId")
            scope_el = sheets[int(lid)] if lid is not None and int(lid) < len(sheets) else None
            scope = self.lineage_o.get(scope_el.get("name"), scope_el.get("name")) if scope_el is not None else None
            self.name_els[(d.get("name"), scope)] = (d, scope_el)

    def scope_element(self, scope):
        if scope is None:
            return None
        ours = {v: k for k, v in self.lineage_o.items()}.get(scope, scope)
        return self.sheet_el.get(ours) or self.new_sheet_el.get(scope)

    def merge_names(self):
        b = self._names(self.b, {})
        o = self._names(self.o, self.lineage_o)
        t = self._names(self.t, self.lineage_t)
        wb = self.tree(self.wb)
        for key in set(b) | set(t):
            bv, ov, tv = b.get(key), o.get(key), t.get(key)
            if tv == bv or tv == ov:
                continue
            if ov != bv:
                self.conflict("(names)", key[0], bv and bv[0], ov and ov[0], tv and tv[0])
                continue
            if tv is None:
                d, _ = self.name_els.pop(key)
                d.getparent().remove(d)
                continue
            if key in self.name_els:
                d = self.name_els[key][0]
            elif key[1] is not None and self.scope_element(key[1]) is None:
                self.conflict("(names)", key[0], None, None, f"{tv[0]} (its sheet wasn't merged)")
                continue
            else:
                dn = wb.find(m("definedNames"))
                if dn is None:
                    dn = etree.SubElement(wb, m("definedNames"))
                    self.place(wb, dn, WB_ORDER)
                d = etree.SubElement(dn, m("definedName"))
                self.name_els[key] = (d, self.scope_element(key[1]))
            for k in list(d.attrib):
                if k != "localSheetId":
                    del d.attrib[k]
            for k, v in tv[1]:
                d.set(k, v)
            d.text = tv[0]

    # --- everything else: charts, drawings, images, macros, ... ---

    def merge_object_parts(self):
        for name in sorted(self.b.parts):
            if NOT_OBJECT_PART.match(name):
                continue
            unit = lambda pkg: (comparable(name, pkg.parts.get(name)), pkg.parts.get(rels_name(name)))
            b, o, t = unit(self.b), unit(self.o), unit(self.t)
            if t == b or t == o or t[0] is None:
                continue  # a deleted part vanishes below once nothing links to it
            if o != b:
                what = "deleted on ours" if o[0] is None else "changed on both branches; kept ours"
                self.conflict("(file)", name, "", "", what)
                continue
            self.put(name, self.t.parts[name])
            self.copy_rels(name, name)
            self.objects_taken.append(name)

    def report_unmerged_additions(self):
        their_reach = reachable(self.t.parts)
        for name in sorted(their_reach - set(self.b.parts) - set(self.imported)):
            if not NOT_OBJECT_PART.match(name):
                self.conflict("(file)", name, "", "", "added on their branch; not merged")

    def write_conflict_sheet(self):
        if not self.conflicts:
            return
        root = etree.Element(m("worksheet"), nsmap={None: MAIN, "r": REL})
        sd = etree.SubElement(root, m("sheetData"))
        header = ["sheet", "cell", "base", "ours (kept)", "theirs", "go to"]
        live = self.sheet_map()
        for i, row in enumerate([header] + [list(c) for c in self.conflicts], start=1):
            r = etree.SubElement(sd, m("row"))
            r.set("r", str(i))
            is_cell = i > 1 and re.match(r"^[A-Z]+\d+$", str(row[1]))
            values = [fmt(v) if is_cell and j in (2, 3, 4) else ("" if v is None else str(v))
                      for j, v in enumerate(row)]
            if is_cell and row[0] in live:
                sheet = row[0].replace("'", "''")
                values.append(f"=HYPERLINK(\"#'{sheet}'!{row[1]}\",\"{row[0]}!{row[1]}\")")
            for j, v in enumerate(values):
                c = etree.SubElement(r, m("c"))
                c.set("r", f"{get_column_letter(j + 1)}{i}")
                self.write_value(c, v, as_text=j < 5)
        if CONFLICT_SHEET in live:
            self.trees[live[CONFLICT_SHEET][1]] = root
            return
        part = self.unique_name("xl/worksheets/sheet1.xml")
        self.put(part, b"")
        self.trees[part] = root
        ct = self.tree("[Content_Types].xml")
        o = etree.SubElement(ct, f"{{{CTYPES}}}Override")
        o.set("PartName", "/" + part)
        o.set("ContentType", WORKSHEET_CT)
        self.add_sheet(CONFLICT_SHEET, part)

    # --- write out ---

    def finish(self):
        wb = self.tree(self.wb)
        sheets = list(wb.find(m("sheets")))
        for key, (d, scope_el) in list(self.name_els.items()):
            if scope_el is None:
                d.attrib.pop("localSheetId", None)
            elif scope_el.getparent() is None:  # its sheet was deleted
                d.getparent().remove(d)
            else:
                d.set("localSheetId", str(sheets.index(scope_el)))
        dn = wb.find(m("definedNames"))
        if dn is not None and len(dn) == 0:
            wb.remove(dn)
        for view in wb.iter(m("workbookView")):
            for a in ("activeTab", "firstSheet"):
                if view.get(a) and int(view.get(a)) >= len(sheets):
                    view.set(a, "0")
        # Cached formula results are stale now; have Excel recalculate on open
        # and rebuild its calculation chain.
        calc = wb.find(m("calcPr"))
        if calc is None:
            calc = etree.SubElement(wb, m("calcPr"))
            self.place(wb, calc, WB_ORDER)
        calc.set("fullCalcOnLoad", "1")
        for r in list(self.rels_tree(self.wb)):
            if r.get("Type").endswith("/calcChain"):
                self.rels_tree(self.wb).remove(r)

        for name, root in self.trees.items():
            self.res[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        # Drop parts nothing links to any more (deleted sheets, replaced charts).
        before = reachable(self.o.parts) | set(self.imported.values())
        after = reachable(self.res)
        for name in list(self.res):
            if name != "[Content_Types].xml" and not name.endswith(".rels") and name in before and name not in after:
                self.drop(name)
                self.drop(rels_name(name))
        ct = self.tree("[Content_Types].xml")
        self.res["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            names = ["[Content_Types].xml"] + [n for n in self.order if n != "[Content_Types].xml"]
            for name in names:
                if name in self.res:
                    z.writestr(name, self.res[name])
        return out.getvalue()


def merge(base_path, ours_path, theirs_path, display_path=None):
    """Write the merge of THEIRS into OURS over OURS (git's convention).
    Exit 0 = clean, 1 = conflicts (listed in a _merge_conflicts sheet)."""
    mg = Merger(base_path, ours_path, theirs_path)
    data = mg.run()
    with open(ours_path, "wb") as f:
        f.write(data)
    name = display_path or ours_path
    print(f"xlgit merge {name}: {mg.cells_taken} cell(s) and {len(mg.objects_taken)} object(s) "
          f"taken from theirs, {len(mg.conflicts)} conflict(s)", file=sys.stderr)
    for sheet, coord, bv, ov, tv in mg.conflicts:
        print(f"  CONFLICT {sheet} {coord}: base={fmt(bv)} ours={fmt(ov)} theirs={fmt(tv)}", file=sys.stderr)
    return 1 if mg.conflicts else 0


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
