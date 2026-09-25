"""xlgit: make Excel workbooks behave like code in git and GitHub.

Commands:
  textconv FILE                    print a workbook as diffable text (git diff driver)
  diff OLD NEW [--markdown]        cell and object (chart/table/pivot/...) diff
  merge BASE OURS THEIRS [PATH]    3-way merge (git merge driver)
  install                          wire the drivers into the current git repo

The merge edits the workbook's XML parts in place instead of re-saving it
through a spreadsheet library, so charts, images, formatting, comments,
tables, pivot tables and macros in OUR copy survive untouched, and their
edits to those objects are carried over.
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
REL_TABLE = REL + "/table"
REL_PIVOT = REL + "/pivotTable"
REL_CACHE = REL + "/pivotCacheDefinition"
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
TABLE_ORDER = ["autoFilter", "sortState", "tableColumns", "tableStyleInfo", "extLst"]

# Parts that are merged some other way (cells, sheet list, names) or that
# every save rewrites. Anything else their branch adds that doesn't get
# merged is reported instead of silently dropped.
NOT_OBJECT_PART = re.compile(
    r"^(\[Content_Types\]\.xml|docProps/.*|customXml/.*|xl/workbook\.xml|xl/worksheets/.*"
    r"|xl/sharedStrings\.xml|xl/calcChain\.xml|xl/styles\.xml|xl/theme/.*|xl/metadata\.xml"
    r"|xl/persons/.*|xl/printerSettings/.*|.*\.rels)$")
# Workbook-level parts left as ours.
WB_KEEP = {"sharedStrings", "styles", "theme", "calcChain", "metadata", "sheetMetadata", "person"}
# Pivot cache attributes that change on every refresh without meaning anything.
VOLATILE_CACHE_ATTRS = ("refreshedDate", "refreshedDateIso", "refreshedBy", "refreshedVersion",
                        "refreshOnLoad", "recordCount", "invalid")

ArrayF = namedtuple("ArrayF", "ref text")
Opaque = namedtuple("Opaque", "desc")


def m(tag):
    return f"{{{MAIN}}}{tag}"


def local(el):
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def kind_of(rel_type):
    return rel_type.rsplit("/", 1)[-1]


def digest(data):
    return hashlib.sha1(data or b"").hexdigest()[:8]


def parse_ref(ref):
    a, _, b = ref.replace("$", "").partition(":")
    (c1, r1), (c2, r2) = coordinate_from_string(a), coordinate_from_string(b or a)
    return column_index_from_string(c1), r1, column_index_from_string(c2), r2


def format_ref(c1, r1, c2, r2):
    lo, hi = f"{get_column_letter(c1)}{r1}", f"{get_column_letter(c2)}{r2}"
    return lo if lo == hi else f"{lo}:{hi}"


def overlaps(a, b):
    a1, ar1, a2, ar2 = parse_ref(a)
    b1, br1, b2, br2 = parse_ref(b)
    return not (a2 < b1 or b2 < a1 or ar2 < br1 or br2 < ar1)


def rename_refs(text, renames):
    """Point structured references (Table2[Amount], =SUM(Table2)) at renamed tables."""
    for old, new in renames.items():
        text = re.sub(rf"(?<![\w.]){re.escape(old)}(?=\[|(?![\w.(!]))", new, text, flags=re.I)
    return text


def strip_dxf(root):
    """Differential formats (dxfId) and custom table/pivot styles index into
    styles.xml; when the branches' styles differ those numbers are meaningless."""
    for el in root.iter():
        for a in list(el.attrib):
            if a.lower().endswith("dxfid"):
                del el.attrib[a]
    for el in root.iter(m("tableStyleInfo"), m("pivotTableStyleInfo")):
        name = el.get("name") or ""
        if not re.match(r"(Table|Pivot)Style(Light|Medium|Dark)\d+$", name):
            el.set("name", "TableStyleMedium2" if local(el) == "tableStyleInfo" else "PivotStyleLight16")


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

    def target(self, owner, kind):
        return next((t for _, typ, t, ext in self.rels(owner) if not ext and kind_of(typ) == kind), None)

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
        t = self.target(self.workbook, "sharedStrings")
        return list(self.xml(t).iter(m("si"))) if t in self.parts else []

    def styles_bytes(self):
        return self.parts.get(self.target(self.workbook, "styles"))

    def sheet_object(self, part, key):
        """Target part of a sheet's drawing / legacyDrawing (element-wired) or comments (rel-only)."""
        if not part or part not in self.parts:
            return None, None
        rels = {rid: (typ, t) for rid, typ, t, ext in self.rels(part) if not ext}
        if key == "comments":
            return next(((t, typ) for typ, t in rels.values() if kind_of(typ) == "comments"), (None, None))
        el = self.xml(part).find(m(key))
        if el is None:
            return None, None
        typ, t = rels.get(el.get(RID), (None, None))
        return t, typ


class Keys:
    """Structural identity for every part of a package.

    Excel and the Python writers renumber files on every save (add a chart
    to an early sheet and every later chart1.xml becomes chart2.xml), so
    part names can't be used to find "the same chart" in another version.
    Instead each part gets a key describing what it is: the drawing of sheet
    Budget, the shape called "Chart 1" in that drawing, table id 3, the
    pivot table called PivotTable1 on sheet Summary, and so on."""

    def __init__(self, pkg, sheet_keys, base_keys=()):
        self.pkg = pkg
        self.sheet_keys = sheet_keys
        self.key_of, self.part_of = {}, {}
        self._xml = {}
        self.pivots = []
        for _, typ, t, ext in pkg.rels(""):
            if not ext and t in pkg.parts:
                self._set(t, ("workbook",) if kind_of(typ) == "officeDocument" else ("pkg", kind_of(typ)))
        wb = pkg.workbook
        if wb in pkg.parts:
            sheet_rids = {s.get(RID): s.get("name") for s in pkg.xml(wb).find(m("sheets"))}
            for rid, typ, t, ext in pkg.rels(wb):
                if ext or t not in pkg.parts or kind_of(typ) == "pivotCacheDefinition":
                    continue
                if rid in sheet_rids:
                    self._set(t, ("sheet", sheet_keys.get(sheet_rids[rid], sheet_rids[rid])))
                else:
                    self._set(t, ("wb", kind_of(typ)))
            for part in list(self.key_of):
                self._walk(part)
            # A pivot cache is known by the data it summarizes; if several
            # caches read the same source, by the pivot tables that use them.
            users = {}
            for pkey, ppart in self.pivots:
                cache = pkg.target(ppart, "pivotCacheDefinition")
                if cache in pkg.parts:
                    users.setdefault(cache, []).append(pkey)
            # Caches whose pivots existed in base claim their keys first, so a
            # new cache on the same data doesn't steal an old one's identity.
            rank = lambda cache: (not any(k in base_keys for k in users[cache]), min(repr(k) for k in users[cache]))
            for cache in sorted(users, key=rank):
                self._set(cache, ("cache", self._cache_source(cache)))
                self._walk(cache)
            for _, typ, t, ext in pkg.rels(wb):
                if not ext and kind_of(typ) == "pivotCacheDefinition" and t in pkg.parts and t not in self.key_of:
                    self._set(t, ("cache", "unused", t))
                    self._walk(t)
        for name in pkg.parts:
            if name not in self.key_of and not name.endswith(".rels"):
                self._set(name, ("part", name))

    def _set(self, part, key):
        if part in self.key_of:
            return
        n = 2
        base = key
        while key in self.part_of:
            key = base + (n,)
            n += 1
        self.key_of[part] = key
        self.part_of[key] = part

    def _tree(self, part):
        if part not in self._xml:
            try:
                self._xml[part] = etree.fromstring(self.pkg.parts[part])
            except etree.XMLSyntaxError:
                self._xml[part] = None
        return self._xml[part]

    def _cache_source(self, cache):
        root = self._tree(cache)
        src = root.find(m("cacheSource")) if root is not None else None
        if src is None:
            return ("unknown",)
        ws = src.find(m("worksheetSource"))
        if ws is None:
            return (src.get("type"), src.get("connectionId"))
        if ws.get("name"):
            return ("name", ws.get("name"))
        return ("sheet", self.sheet_keys.get(ws.get("sheet"), ws.get("sheet")))

    def _shape_name(self, drawing, rid):
        """Name of the drawing shape (chart, picture) that uses relationship rid."""
        root = self._tree(drawing)
        for el in (root.iter() if root is not None else []):
            if not any(k.startswith(f"{{{REL}}}") and v == rid for k, v in el.attrib.items()):
                continue
            anc = el
            while anc is not None:
                for child in anc:
                    if local(child).startswith("nv"):
                        for c in child:
                            if local(c) == "cNvPr":
                                return c.get("name") or c.get("id")
                anc = anc.getparent()
        return rid

    def _walk(self, owner):
        okey = self.key_of[owner]
        for rid, typ, t, ext in self.pkg.rels(owner):
            if ext or t not in self.pkg.parts or t in self.key_of:
                continue
            kind = kind_of(typ)
            if kind == "pivotCacheDefinition":
                continue  # keyed through the pivot tables that use it
            if okey[0] == "sheet":
                sk = okey[1]
                if kind == "drawing":
                    key = ("drawing", sk)
                elif kind == "vmlDrawing":
                    root = self._tree(owner)
                    el = next((e for e in root.iter() if e.get(RID) == rid), None) if root is not None else None
                    key = ("vml", sk, local(el) if el is not None else rid)
                elif kind == "comments":
                    key = ("comments", sk)
                elif kind == "table":
                    key = ("table", (self._tree(t).get("id") if self._tree(t) is not None else t))
                elif kind == "pivotTable":
                    key = ("pivot", sk, self._tree(t).get("name") if self._tree(t) is not None else t)
                    self.pivots.append((key, t))
                else:
                    key = ("sub", okey, kind)
            elif okey[0] == "drawing":
                key = ("shape", okey, self._shape_name(owner, rid), kind)
            else:
                key = ("sub", okey, kind)
            self._set(t, key)
            if key[0] == "pivot":
                self.pivots[-1] = (self.key_of[t], t)
            self._walk(t)


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


# ---------- normalizing objects for comparison ----------

def _strip_chart_caches(root):
    """Charts embed cached copies of the cells they plot; drop them so a
    cell edit doesn't look like a chart edit."""
    for el in list(root.iter(f"{{{CHART}}}numCache", f"{{{CHART}}}strCache")):
        el.getparent().remove(el)
    # Axis ids only link a chart's axes to each other, and some writers derive
    # them from the chart's file number; compare them by order of appearance.
    ids = {}
    for el in root.iter(f"{{{CHART}}}axId", f"{{{CHART}}}crossAx"):
        el.set("val", str(ids.setdefault(el.get("val"), len(ids))))


def _norm_cache(root, derived):
    for a in VOLATILE_CACHE_ATTRS:
        root.attrib.pop(a, None)
    if derived:  # the shared items are just the distinct values of the source data
        for si in root.iter(m("sharedItems")):
            si.clear()


def _norm_pivot(root, derived):
    for a in ("cacheId", "updatedVersion"):
        root.attrib.pop(a, None)
    if derived:  # what a refresh rewrites: rendered rows/columns, size, item lists
        for tag in ("rowItems", "colItems"):
            for el in root.findall(m(tag)):
                root.remove(el)
        loc = root.find(m("location"))
        if loc is not None:
            loc.attrib.pop("ref", None)
        for items in root.iter(m("items")):
            hidden = sum(1 for i in items if i.get("h") in ("1", "true"))
            items.clear()
            items.set("hidden", str(hidden))


def normalized(keys, part, derived=False):
    """Comparable form of a part: relationship ids replaced by what they point
    at, and noise (chart caches, refresh stamps) removed. derived=True also
    drops everything a pivot refresh recomputes from the source data."""
    if part is None:
        return None
    pkg = keys.pkg
    rels = {rid: (typ, t, ext) for rid, typ, t, ext in pkg.rels(part)}
    target_key = lambda t, ext: t if ext else repr(keys.key_of.get(t, t))
    relset = tuple(sorted((kind_of(typ), target_key(t, ext)) for typ, t, ext in rels.values()
                          if not (derived and kind_of(typ) == "pivotCacheRecords")))
    try:
        root = etree.fromstring(pkg.parts[part])
    except etree.XMLSyntaxError:
        return pkg.parts[part], relset
    kind = keys.key_of.get(part, ("",))[0]
    if kind == "cache":
        _norm_cache(root, derived)
    elif kind == "pivot":
        _norm_pivot(root, derived)
    elif local(root) == "chartSpace":
        _strip_chart_caches(root)
    for el in root.iter():
        for k, v in list(el.attrib.items()):
            if k.startswith(f"{{{REL}}}") and v in rels:
                el.set(k, target_key(rels[v][1], rels[v][2]))
    return etree.tostring(root), relset


# ---------- describing charts, images, comments, tables, pivots ----------

def _chart_summary(pkg, name):
    root = pkg.xml(name)
    plot = root.find(f".//{{{CHART}}}plotArea")
    kinds = [local(e) for e in plot if local(e).endswith("Chart")] if plot is not None else []
    title_el = root.find(f"{{{CHART}}}chart/{{{CHART}}}title")
    title = "".join(t.text or "" for t in title_el.iter(f"{{{DRAW}}}t")) if title_el is not None else ""
    refs = [f.text for f in root.iter(f"{{{CHART}}}f") if f.text]
    _strip_chart_caches(root)
    return f"chart: {'+'.join(kinds) or 'chart'} {title!r} data={','.join(refs)} #{digest(etree.tostring(root))}"


def _pivot_summary(pkg, name):
    p = pkg.xml(name)
    cache = pkg.target(name, "pivotCacheDefinition")
    fields = [cf.get("name") for cf in pkg.xml(cache).iter(m("cacheField"))] if cache in pkg.parts else []
    field = lambda x: fields[int(x)] if x and x.isdigit() and int(x) < len(fields) else "Values"
    axis = lambda tag: [field(f.get("x")) for el in p.findall(m(tag)) for f in el]
    values = [d.get("name") for d in p.iter(m("dataField"))]
    loc = p.find(m("location"))
    where = loc.get("ref") if loc is not None else "?"
    _norm_pivot(p, derived=True)
    return (f"pivot table: {p.get('name')} at {where} "
            f"rows={axis('rowFields')} cols={axis('colFields')} values={values} #{digest(etree.tostring(p))}")


def describe_objects(pkg):
    """{sheet: [one line per chart / image / comment / table / pivot table]}."""
    out = {}
    for sheet, _, part in pkg.sheets():
        items = []
        for _, typ, tgt, ext in pkg.rels(part):
            if ext or tgt not in pkg.parts:
                continue
            kind = kind_of(typ)
            if kind == "drawing":
                for _, t2, tgt2, ext2 in pkg.rels(tgt):
                    if ext2 or tgt2 not in pkg.parts:
                        continue
                    if kind_of(t2) == "chart":
                        items.append(_chart_summary(pkg, tgt2))
                    elif kind_of(t2) == "image":
                        items.append(f"image: {posixpath.basename(tgt2)} #{digest(pkg.parts[tgt2])}")
            elif kind == "comments":
                for c in pkg.xml(tgt).iter(m("comment")):
                    text = "".join(t.text or "" for t in c.iter(m("t")))
                    items.append(f"comment {c.get('ref')}: {text!r}")
            elif kind == "table":
                t = pkg.xml(tgt)
                cols = [c.get("name") for c in t.iter(m("tableColumn"))]
                items.append(f"table: {t.get('displayName')} {t.get('ref')} columns={cols}")
            elif kind == "pivotTable":
                items.append(_pivot_summary(pkg, tgt))
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


def pivot_locations(pkg, sheet_part):
    refs = []
    for _, typ, t, ext in (pkg.rels(sheet_part) if sheet_part else []):
        if not ext and kind_of(typ) == "pivotTable" and t in pkg.parts:
            loc = pkg.xml(t).find(m("location"))
            if loc is not None and loc.get("ref"):
                refs.append(loc.get("ref"))
    return refs


def _three_way(b, o, t):
    """(value, ok): the merged value, or ok=False if both sides changed it differently."""
    if t == b or t == o:
        return o, True
    if o == b:
        return t, True
    return o, False


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
        self.notes = []
        self.cells_taken = 0
        self.objects_taken = []
        self.stale_pivots = False
        self.touched_tables = set()
        self.table_renames = {}
        self.wb = self.o.workbook
        self.same_styles = self.o.styles_bytes() == self.t.styles_bytes()
        self.sst = self.o.shared_strings()
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

    def rels_of(self, owner):
        rt = self.tree(rels_name(owner))
        if rt is None:
            return []
        out = []
        for r in rt:
            ext = r.get("TargetMode") == "External"
            out.append((r.get("Id"), r.get("Type"), r.get("Target") if ext else resolve(owner, r.get("Target")), ext))
        return out

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

    def ours_for(self, their_part):
        """Our part for the same object as one of theirs, if both branches
        inherited it from base and we still have it."""
        key = self.kt.key_of.get(their_part)
        if key is None or key not in self.kb.part_of:
            return None
        ours = self.ko.part_of.get(key)
        return ours if ours in self.res else None

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
        """Their relationships for a part, retargeted into the result: links to
        objects we share point at our copy, anything new gets imported.
        Returns (rId, type) pairs that keep_rel rejected."""
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
            ours = self.ours_for(tgt)
            if ours is None:
                ours = self.import_part(tgt) if tgt in self.t.parts else tgt
            r.set("Target", relative(ours, our_owner))
        self.trees[rn] = rt
        self.put(rn, b"")
        return dropped

    def take_theirs(self, our_part, their_part):
        self.put(our_part, self.t.parts[their_part])
        self.copy_rels(their_part, our_part)

    def changed_since_base(self, our_part, seen=None):
        """Did we edit this object, or anything hanging off it?"""
        seen = seen if seen is not None else set()
        if our_part in seen:
            return False
        seen.add(our_part)
        key = self.ko.key_of.get(our_part)
        base_part = self.kb.part_of.get(key)
        if base_part is None or normalized(self.ko, our_part) != normalized(self.kb, base_part):
            return True
        return any(self.changed_since_base(t, seen) for _, typ, t, ext in self.o.rels(our_part)
                   if not ext and t in self.o.parts and kind_of(typ) != "pivotCacheDefinition")

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
        self.kb = Keys(self.b, {n: n for n, _, _ in bs})
        self.ko = Keys(self.o, {n: self.lineage_o.get(n, "ours:" + n) for n, _, _ in os_}, self.kb.part_of)
        self.kt = Keys(self.t, {n: self.lineage_t.get(n, "theirs:" + n) for n, _, _ in ts}, self.kb.part_of)
        self.sheet_el = {n: el for n, (el, _) in self.sheet_map().items()}  # ours, before any change
        self.new_sheet_el = {}
        self.capture_names()
        self.plan_table_renames()

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
        self.merge_tables()
        self.merge_pivots()
        self.fix_tables()
        self.fix_pivots()
        self.report_unmerged_additions()
        self.write_conflict_sheet()
        return self.finish()

    # --- cells ---

    def merge_sheet(self, bname, oname, tname):
        bpart = dict((n, p) for n, _, p in self.b.sheets()).get(bname)
        opart = dict((n, p) for n, _, p in self.o.sheets()).get(oname)
        tpart = dict((n, p) for n, _, p in self.t.sheets()).get(tname)
        # Cells a pivot table displays are regenerated when it refreshes, so
        # they never conflict: we keep ours and Excel refreshes on open.
        pivot_areas = [ref for pkg, part in ((self.b, bpart), (self.o, opart), (self.t, tpart))
                       for ref in pivot_locations(pkg, part)]
        b = self.bc.get(bname, {}) if bname else {}
        o, t = self.oc.get(oname, {}), self.tc.get(tname, {})
        edits = {}
        for coord in set(b) | set(t):
            bv, ov, tv = b.get(coord), o.get(coord), t.get(coord)
            if tv == bv or tv == ov:
                continue
            if ov == bv and not isinstance(tv, Opaque):
                edits[coord] = tv
            elif any(overlaps(coord, ref) for ref in pivot_areas):
                self.stale_pivots = True
            else:
                self.conflict(oname, coord, bv, ov, tv)
        if edits:
            self.edit_cells(oname, tname, edits)
        for key in ("drawing", "legacyDrawing", "comments"):
            self.merge_sheet_object(oname, key, bpart, opart, tpart)
        for kind in ("table", "pivotTable"):
            self.merge_sheet_collection(oname, kind, bpart, opart, tpart)

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
        dim.set("ref", format_ref(min(cols), min(rows), max(cols), max(rows)))

    def cell_text(self, c):
        if c is None:
            return ""
        if c.get("t") == "s":
            v = c.find(m("v"))
            si = self.sst[int(v.text)] if v is not None and int(v.text) < len(self.sst) else None
            return "".join(t.text or "" for t in si.iter(m("t"))) if si is not None else ""
        if c.get("t") == "inlineStr":
            return "".join(t.text or "" for t in c.iter(m("t")))
        v = c.find(m("v"))
        return v.text or "" if v is not None else ""

    # --- charts, images and comments hanging off a sheet ---

    def merge_sheet_object(self, oname, key, bpart, opart, tpart):
        """Whether a sheet has a drawing / comments at all. What's inside
        them merges separately, object by object."""
        b = self.b.sheet_object(bpart, key)[0] is not None
        o_part = self.o.sheet_object(opart, key)[0]
        t_part, ttype = self.t.sheet_object(tpart, key)
        o, t = o_part is not None, t_part is not None
        if t == b:
            return
        if o != b:  # both branches added one
            self.conflict(oname, f"({key})", "", "added", "added; theirs not merged")
            return
        part = self.sheet_map()[oname][1]
        if not t and self.changed_since_base(o_part):
            self.conflict(oname, f"({key})", "", "changed", "removed; kept ours")
            return
        root = self.tree(part)
        old_el = root.find(m(key)) if key != "comments" else None
        for r in list(self.rels_tree(part)):
            if (old_el is not None and r.get("Id") == old_el.get(RID)) or \
                    (key == "comments" and kind_of(r.get("Type")) == "comments"):
                self.rels_tree(part).remove(r)
        if old_el is not None:
            root.remove(old_el)
        if not t:
            self.objects_taken.append(f"{oname}: {key} removed as on their branch")
            return
        rid = self.add_rel(part, ttype, self.import_part(t_part))
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

    # --- tables and pivot tables on a sheet ---

    def merge_sheet_collection(self, oname, kind, bpart, opart, tpart):
        """Tables / pivot tables added or removed on their branch."""
        def keys_on(keys, pkg, part):
            if not part:
                return {}
            return {keys.key_of[t]: t for _, typ, t, ext in pkg.rels(part)
                    if not ext and kind_of(typ) == kind and t in keys.key_of}
        b = keys_on(self.kb, self.b, bpart)
        o = keys_on(self.ko, self.o, opart)
        t = keys_on(self.kt, self.t, tpart)
        for key in sorted(set(t) - set(b), key=repr):
            if key in o:
                if normalized(self.ko, o[key]) != normalized(self.kt, t[key]):
                    self.conflict(oname, f"({kind})", "", "added", "added with the same identity; kept ours")
                continue
            self.wire_new(oname, kind, t[key])
        for key in sorted((set(b) - set(t)) & set(o), key=repr):
            if kind == "pivotTable":  # a refresh on our side isn't a reason to keep it
                changed = normalized(self.ko, o[key], True) != normalized(self.kb, b[key], True)
            else:
                changed = self.changed_since_base(o[key])
            if changed:
                self.conflict(oname, f"({kind} {self.o.xml(o[key]).get('name')})", "", "changed", "deleted; kept ours")
            else:
                self.unwire(oname, o[key])
                self.objects_taken.append(f"{oname}: {kind} removed as on their branch")

    def wire_new(self, oname, kind, their_part):
        part = self.sheet_map()[oname][1]
        theirs = self.t.xml(their_part)
        label = theirs.get("displayName") or theirs.get("name")
        new_ref = theirs.get("ref") if kind == "table" else theirs.find(m("location")).get("ref")
        names = set()
        for _, typ, tgt, ext in self.rels_of(part):
            if ext or tgt not in self.res or kind_of(typ) not in ("table", "pivotTable"):
                continue
            x = self.tree(tgt)
            ref = x.get("ref") if kind_of(typ) == "table" else x.find(m("location")).get("ref")
            if overlaps(ref, new_ref):
                self.conflict(oname, f"({kind} {label})", "", "", f"added at {new_ref}, overlapping ours; not merged")
                return
            names.add((x.get("name") or "").lower())
        if kind == "table" and self.t.rels(their_part):
            self.conflict(oname, f"(table {label})", "", "", "added with an external data connection; not merged")
            return
        target = self.import_part(their_part)
        typ = REL_TABLE if kind == "table" else REL_PIVOT
        rid = self.add_rel(part, typ, target)
        if kind == "table":
            root = self.tree(part)
            parts = root.find(m("tableParts"))
            if parts is None:
                parts = etree.SubElement(root, m("tableParts"))
                self.place(root, parts, WS_ORDER)
            etree.SubElement(parts, m("tablePart")).set(RID, rid)
            parts.set("count", str(len(parts)))
        elif label.lower() in names:
            n = 2
            while f"{label} ({n})".lower() in names:
                n += 1
            self.tree(target).set("name", f"{label} ({n})")
        self.objects_taken.append(f"{oname}: {kind} {label} from their branch")

    def unwire(self, oname, our_part):
        part = self.sheet_map()[oname][1]
        root = self.tree(part)
        for rid, typ, tgt, ext in self.rels_of(part):
            if tgt != our_part:
                continue
            self.remove_rel(part, rid)
            for el in [e for e in root.iter() if e.get(RID) == rid]:
                parent = el.getparent()
                parent.remove(el)
                if local(parent) == "tableParts":
                    parent.set("count", str(len(parent)))
                    if len(parent) == 0:
                        root.remove(parent)

    def plan_table_renames(self):
        """Table names are unique per workbook. If both branches added a table
        called Table2, theirs becomes Table3 and their formulas follow."""
        taken = {n.lower() for n, _ in self._names(self.o, {})}
        for key, part in self.ko.part_of.items():
            if key[0] == "table":
                taken.add((self.o.xml(part).get("displayName") or "").lower())
        new = []
        for key, part in self.kt.part_of.items():
            if key[0] == "table":
                name = self.t.xml(part).get("displayName") or ""
                if key in self.kb.part_of:
                    taken.add(name.lower())
                else:
                    new.append(name)
        for name in sorted(new):
            if name.lower() in taken:
                stem, num = re.match(r"(.*?)(\d*)$", name).groups()
                n = int(num or 1) + 1
                while f"{stem}{n}".lower() in taken:
                    n += 1
                self.table_renames[name] = f"{stem}{n}"
                self.notes.append(f"their new table {name} renamed to {stem}{n} (ours has a {name})")
                name = f"{stem}{n}"
            taken.add(name.lower())
        if self.table_renames:
            for cells in self.tc.values():
                for coord, v in cells.items():
                    if isinstance(v, str) and v.startswith("="):
                        cells[coord] = rename_refs(v, self.table_renames)

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
        """A sheet only they added: copy it whole, charts, tables and pivots
        included. Shared strings become inline strings since the string table
        isn't merged."""
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
        if self.table_renames:
            for f in root.iter(m("f")):
                if f.text:
                    f.text = rename_refs(f.text, self.table_renames)
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
        keep = ("drawing", "vmlDrawing", "comments", "hyperlink", "printerSettings", "table", "pivotTable")
        dropped = self.copy_rels(tpart, ours, keep_rel=lambda typ: kind_of(typ) in keep)
        for rid, typ in dropped:
            for el in [e for e in root.iter() if e.get(RID) == rid]:
                el.getparent().remove(el)
            self.conflict(tname, f"({kind_of(typ)})", "", "", "on their new sheet; not merged")
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
            d.text = rename_refs(tv[0], self.table_renames) if tv[0] else tv[0]

    # --- charts, drawings, images, comments, macros: object by object ---

    def family(self, key):
        while key and key[0] in ("sub", "shape"):
            key = key[1]
        return key[0] if key else ""

    def merge_object_parts(self):
        for key, bpart in sorted(self.kb.part_of.items(), key=lambda kv: repr(kv[0])):
            if key[0] in ("pkg", "workbook", "sheet", "part") or (key[0] == "wb" and key[1] in WB_KEEP):
                continue
            if self.family(key) in ("table", "pivot", "cache"):
                continue  # merged by merge_tables / merge_pivots
            opart, tpart = self.ko.part_of.get(key), self.kt.part_of.get(key)
            b = normalized(self.kb, bpart)
            o, t = normalized(self.ko, opart), normalized(self.kt, tpart)
            if t is None or t == b or t == o:
                continue  # a part they deleted vanishes once nothing links to it
            if o is None:
                self.conflict("(file)", tpart, "", "", "deleted on ours, changed on theirs; kept deletion")
            elif o != b:
                self.conflict("(file)", opart, "", "", "changed on both branches; kept ours")
            else:
                self.take_theirs(opart, tpart)
                self.objects_taken.append(opart)

    # --- tables ---

    def merge_tables(self):
        for key, bpart in sorted(self.kb.part_of.items(), key=lambda kv: repr(kv[0])):
            if key[0] != "table":
                continue
            opart, tpart = self.ko.part_of.get(key), self.kt.part_of.get(key)
            if opart is None or tpart is None:
                if opart is None and tpart and normalized(self.kt, tpart) != normalized(self.kb, bpart):
                    self.conflict("(table)", self.t.xml(tpart).get("displayName"), "", "deleted",
                                  "changed; kept deletion")
                continue  # removal on their side is handled per sheet
            b, o, t = (normalized(k, p) for k, p in ((self.kb, bpart), (self.ko, opart), (self.kt, tpart)))
            if t == b or t == o:
                continue
            name = self.o.xml(opart).get("displayName")
            if o == b:
                self.put(opart, self.t.parts[tpart])
                if not self.same_styles:
                    strip_dxf(self.tree(opart))
            else:
                merged = self.merge_table_xml(self.b.xml(bpart), self.o.xml(opart), self.t.xml(tpart))
                if merged is None:
                    self.conflict("(table)", name, "", "", "changed on both branches; kept ours")
                    continue
                self.trees[opart] = merged
                self.put(opart, b"")
            self.touched_tables.add(opart)
            self.objects_taken.append(f"table {name}")

    def merge_table_xml(self, b, o, t):
        """Field-by-field merge of a table definition: its range, each column,
        filter and style. None if the two sides can't be reconciled."""
        r = copy.deepcopy(o)
        their_el = lambda el: strip_dxf(el) or el if not self.same_styles else el
        for k in set(b.attrib) | set(t.attrib):
            if k in ("id", "ref") or (not self.same_styles and k.lower().endswith("dxfid")):
                continue
            v, ok = _three_way(b.get(k), o.get(k), t.get(k))
            if not ok:
                return None
            r.set(k, v) if v is not None else r.attrib.pop(k, None)
        ref, ok = _three_way(b.get("ref"), o.get("ref"), t.get("ref"))
        if not ok:  # e.g. we added rows, they added a column: merge each edge
            edges = [_three_way(*x) for x in zip(*(parse_ref(x.get("ref")) for x in (b, o, t)))]
            if not all(ok for _, ok in edges):
                return None
            ref = format_ref(*(v for v, _ in edges))
        r.set("ref", ref)

        cols = lambda x: x.find(m("tableColumns"))
        ids = lambda x: [c.get("id") for c in cols(x)]
        order, ok = _three_way(ids(b), ids(o), ids(t))
        if not ok:
            return None
        by_id = lambda x: {c.get("id"): c for c in cols(x)}
        bc, oc, tc = by_id(b), by_id(o), by_id(t)
        ser = lambda el: etree.tostring(el) if el is not None else None
        new_cols = []
        for cid in order:
            if cid not in oc:
                new_cols.append(their_el(copy.deepcopy(tc[cid])))
                continue
            if cid not in tc or cid not in bc:
                new_cols.append(copy.deepcopy(oc[cid]))
                continue
            _, ok = _three_way(ser(bc[cid]), ser(oc[cid]), ser(tc[cid]))
            if not ok:
                return None
            pick = oc[cid] if ser(tc[cid]) in (ser(bc[cid]), ser(oc[cid])) else their_el(copy.deepcopy(tc[cid]))
            new_cols.append(copy.deepcopy(pick))
        rc = cols(r)
        for c in list(rc):
            rc.remove(c)
        for c in new_cols:
            rc.append(c)
        rc.set("count", str(len(new_cols)))

        for tag in ("autoFilter", "sortState", "tableStyleInfo", "extLst"):
            def plain(x):
                el = x.find(m(tag))
                if el is None:
                    return None
                el = copy.deepcopy(el)
                el.attrib.pop("ref", None)
                return etree.tostring(el)
            _, ok = _three_way(plain(b), plain(o), plain(t))
            if not ok:
                return None
            if plain(t) not in (plain(b), plain(o)):
                old = r.find(m(tag))
                if old is not None:
                    r.remove(old)
                if t.find(m(tag)) is not None:
                    el = their_el(copy.deepcopy(t.find(m(tag))))
                    r.append(el)
                    self.place(r, el, TABLE_ORDER)
        c1, r1, c2, r2 = parse_ref(r.get("ref"))
        if c2 - c1 + 1 != len(new_cols):
            return None
        af = r.find(m("autoFilter"))
        if af is not None:
            totals = int(r.get("totalsRowCount") or 0)
            af.set("ref", format_ref(c1, r1, c2, r2 - totals))
        ss = r.find(m("sortState"))
        if ss is not None and ss.get("ref") and not overlaps(ss.get("ref"), r.get("ref")):
            r.remove(ss)
        return r

    def fix_tables(self):
        """Every table needs a workbook-unique id and name, and its column
        names must match its header cells or Excel 'repairs' the file."""
        live = []
        for name, (_, spart) in self.sheet_map().items():
            for _, typ, tgt, ext in self.rels_of(spart):
                if not ext and kind_of(typ) == "table" and tgt in self.res:
                    live.append((tgt, spart))
        imported = set(self.imported.values())
        live.sort(key=lambda x: x[0] in imported)  # ours keep their ids
        used = set()
        top = max([int(self.tree(t).get("id") or 0) for t, _ in live] + [0])
        for tpart, spart in live:
            root = self.tree(tpart)
            if root.get("id") in used:
                top += 1
                root.set("id", str(top))
            used.add(root.get("id"))
            if tpart in imported:
                self.touched_tables.add(tpart)
                old = root.get("displayName")
                if old in self.table_renames:
                    root.set("displayName", self.table_renames[old])
                    root.set("name", self.table_renames[old])
                for f in root.iter(m("calculatedColumnFormula")):
                    if f.text:
                        f.text = rename_refs(f.text, self.table_renames)
                if not self.same_styles:
                    strip_dxf(root)
            if (tpart in self.touched_tables or spart in self.trees) and root.get("headerRowCount") != "0":
                c1, r1, _, _ = parse_ref(root.get("ref"))
                cells = {c.get("r"): c for c in self.tree(spart).iter(m("c"))}
                cols = list(root.find(m("tableColumns")))
                names = [self.cell_text(cells.get(f"{get_column_letter(c1 + i)}{r1}")) or col.get("name")
                         for i, col in enumerate(cols)]
                if len({n.lower() for n in names}) == len(names):
                    for col, n in zip(cols, names):
                        col.set("name", n)

    # --- pivot tables ---

    def merge_pivots(self):
        """A pivot table, its cache and the cache's records only make sense
        together, so they merge as one bundle per cache. A side that only
        refreshed (same layout, new data) doesn't count as an edit: Excel is
        told to refresh on open, which rebuilds the pivot from merged data."""
        for ckey, bcache in sorted(self.kb.part_of.items(), key=lambda kv: repr(kv[0])):
            if ckey[0] != "cache":
                continue
            members = [ckey] + [k for k in self.kb.part_of if k[0] == "sub" and k[1] == ckey]
            members += [pk for pk, pp in self.kb.pivots if self.b.target(pp, "pivotCacheDefinition") == bcache]
            bundle = lambda keys, derived: tuple(normalized(keys, keys.part_of.get(k), derived) for k in members
                                                 if not (derived and k[0] == "sub"))  # records are pure data
            b, o, t = bundle(self.kb, False), bundle(self.ko, False), bundle(self.kt, False)
            if t == b or t == o:
                continue
            names = ", ".join(str(k[2]) for k in members if k[0] == "pivot")
            if o != b:
                if bundle(self.kt, True) == bundle(self.kb, True):
                    self.stale_pivots = True  # they only refreshed
                    continue
                if bundle(self.ko, True) != bundle(self.kb, True):
                    self.conflict("(pivot)", names, "", "", "changed on both branches; kept ours")
                    self.stale_pivots = True
                    continue
            for k in members:
                opart, tpart = self.ko.part_of.get(k), self.kt.part_of.get(k)
                if opart and tpart:
                    self.take_theirs(opart, tpart)
            self.stale_pivots = True
            self.objects_taken.append(f"pivot table(s) {names}")

    def fix_pivots(self):
        """Register every pivot's cache in workbook.xml with a matching cacheId,
        drop caches nothing uses any more, and have Excel refresh pivots whose
        source data changed in the merge."""
        wb = self.tree(self.wb)
        pcs = wb.find(m("pivotCaches"))
        wb_caches = {r.get("Id"): resolve(self.wb, r.get("Target")) for r in self.rels_tree(self.wb)
                     if kind_of(r.get("Type")) == "pivotCacheDefinition"}
        entry = {wb_caches.get(pc.get(RID)): pc for pc in (pcs if pcs is not None else [])}
        used = set()
        imported = set(self.imported.values())
        for name, (_, spart) in self.sheet_map().items():
            for _, typ, pv, ext in self.rels_of(spart):
                if ext or kind_of(typ) != "pivotTable" or pv not in self.res:
                    continue
                cache = next((t for _, ty, t, e in self.rels_of(pv) if kind_of(ty) == "pivotCacheDefinition"), None)
                if cache not in self.res:
                    continue
                if cache not in entry:
                    if pcs is None:
                        pcs = etree.SubElement(wb, m("pivotCaches"))
                        self.place(wb, pcs, WB_ORDER)
                    pc = etree.SubElement(pcs, m("pivotCache"))
                    pc.set("cacheId", str(max([int(p.get("cacheId")) for p in pcs if p.get("cacheId")] + [0]) + 1))
                    pc.set(RID, self.add_rel(self.wb, REL_CACHE, cache))
                    entry[cache] = pc
                root = self.tree(pv)
                if root.get("cacheId") != entry[cache].get("cacheId"):
                    root.set("cacheId", entry[cache].get("cacheId"))
                if pv in imported and not self.same_styles:
                    strip_dxf(root)
                if cache in imported and self.table_renames:
                    for src in self.tree(cache).iter(m("worksheetSource")):
                        if src.get("name") in self.table_renames:
                            src.set("name", self.table_renames[src.get("name")])
                used.add(cache)
        ours_used = {self.o.target(p, "pivotCacheDefinition") for _, p in self.ko.pivots}
        for cache, pc in list(entry.items()):
            if cache not in used and cache in ours_used:
                self.remove_rel(self.wb, pc.get(RID))
                pcs.remove(pc)
        if pcs is not None and len(pcs) == 0:
            wb.remove(pcs)
        if self.cells_taken or self.stale_pivots:
            for cache in used:
                root = self.tree(cache)
                src = root.find(m("cacheSource"))
                if src is not None and src.get("type") == "worksheet":
                    root.set("refreshOnLoad", "1")

    def report_unmerged_additions(self):
        for part in sorted(reachable(self.t.parts)):
            key = self.kt.key_of.get(part)
            if key in self.kb.part_of or part in self.imported or NOT_OBJECT_PART.match(part):
                continue
            self.conflict("(file)", part, "", "", "added on their branch; not merged")

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
            if kind_of(r.get("Type")) == "calcChain":
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
    for note in mg.notes:
        print(f"  note: {note}", file=sys.stderr)
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
