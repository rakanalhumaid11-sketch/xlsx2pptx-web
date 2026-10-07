"""دمج ملفات «حالة الاعتماد» في ملف واحد.

كل ملف خرج من أداة حالة الاعتماد فيه ورقتان: ورقة الملاحظات و«الداتا شيت».
المخرج ملف واحد:

* ورقة «الداتا شيت» أولًا: داتا شيت كل ملف كما هي، واحدة تحت الأخرى، وبين
  كل واحدة والتالية فاصل صفحة عند الطباعة. وتبقى حيّة: صيغها تقرأ ورقة
  مغذيها في الملف المدموج، فإذا غيّرت حالة صف تحدّث رقمه.
* ثم ورقة لكل ملف باسم مغذيه (AH08-8911)، منسوخة كما هي: الألوان والتنسيق
  الشرطي والقائمة المنسدلة والروابط وعروض الأعمدة.

العمل على XML مباشرة كما في xlsxedit: لا تمرّ الأوراق بمكتبة تعيد بناءها
فتُسقط ما لا تفهمه. ودمج أوراق من ملفات مختلفة يحتاج ثلاثة أشياء:
* جدول أنماط واحد: أنماط كل ملف تُضاف إليه، ويُعاد ترقيم كل خلية إليها.
* جدول نصوص مشتركة واحد بالطريقة نفسها.
* إزاحة صفوف كل داتا شيت إلى موضعها الجديد، مع صيغها ودمج خلاياها وتنسيقها
  الشرطي، وإعادة تسمية الورقة التي تقرأ منها الصيغ.
"""
from __future__ import annotations

import html
import os
import posixpath
import re
import zipfile
from typing import Any, Dict, List, Optional, Tuple

import builder
import sorter
from xlsxedit import _get_attr, _set_attr, col_index, col_letter, esc

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_SHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"

DATA_TITLE = "الداتا شيت"
GAP = 2                      # صفان فارغان بين داتا شيت مغذٍّ والذي بعده

# علاقات لا تُنقل: تعتمد على أجزاء على مستوى المصنّف (الجداول المحورية
# وذاكرتها، والتعليقات المترابطة وقائمة أشخاصها). الملاحظات القديمة تبقى.
SKIP_RELS = ("pivotTable", "threadedComment", "slicer", "timeline")


# ------------------------------------------------------------- أدوات عامة

def _norm_name(s: str) -> str:
    s = (s or "").strip().replace("أ", "ا").replace("إ", "ا").replace("ة", "ه")
    return re.sub(r"\s+", " ", s)


def is_data_sheet(name: str) -> bool:
    n = _norm_name(name)
    return n.startswith(_norm_name(DATA_TITLE)) or n.startswith("داتا شيت")


def _canon(raw: str) -> str:
    """صيغة موحّدة لعنصر XML تكفي لاكتشاف المكرَّر."""
    raw = re.sub(r">\s+<", "><", raw.strip())
    return re.sub(r"\s+/>", "/>", raw)


def _part_rels(part: str) -> str:
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _resolve(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(base_dir, target))


def _ns_decls(open_tag: str) -> Dict[str, str]:
    return dict(re.findall(r'\sxmlns:(\w+)="([^"]*)"', open_tag))


def _ignorable(open_tag: str) -> List[str]:
    v = _get_attr(open_tag, "mc:Ignorable") or ""
    return v.split()


def quote_sheet(name: str) -> str:
    return "'%s'" % name.replace("'", "''")


def safe_sheet_name(name: str) -> str:
    """اسم ورقة يقبله إكسل: بلا [ ] : * ? / \\ وبطول 31 حرفًا على الأكثر."""
    s = re.sub(r"[\[\]:*?/\\]", "-", (name or "").strip()).strip("'")
    return (s or "ورقة")[:31]


# ------------------------------------------------------------ تحريك الصيغ

_REF_RX = re.compile(
    r"(?P<sheet>'(?:[^']|'')+'|[^\W\d][\w.]*)!"
    r"(?P<qref>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?"
    r"|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)"
    r"|(?<![\w$.'!])(?P<ref>\$?[A-Z]{1,3}\$?\d+)(?![\w(!])")
_ONE_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _shift_ref(ref: str, dr: int) -> str:
    if not dr:
        return ref
    return _ONE_REF.sub(lambda m: "%s%s%s%d" % (m.group(1), m.group(2), m.group(3),
                                                int(m.group(4)) + dr), ref)


def shift_formula(f: str, dr: int, rename: Dict[str, str], self_name: str = "") -> str:
    """يزيح مراجع الورقة نفسها dr صفًا، ويعيد تسمية الأوراق الأخرى.

    الكتلة كلها تنتقل، فتتحرّك المراجع المطلقة ($B$6) أيضًا — كما يفعل إكسل
    حين تقصّ خلايا وتلصقها. والنصوص بين علامتي تنصيص لا تُمسّ."""
    parts = f.split('"')
    for i in range(0, len(parts), 2):             # خارج النصوص فقط
        def one(m: "re.Match") -> str:
            if m.group("ref"):
                return _shift_ref(m.group("ref"), dr)
            sheet = m.group("sheet")
            bare = sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet
            qref = m.group("qref")
            if self_name and bare == self_name:
                return quote_sheet(rename.get(bare, bare)) + "!" + _shift_ref(qref, dr)
            if bare in rename:
                return quote_sheet(rename[bare]) + "!" + qref
            return m.group(0)
        parts[i] = _REF_RX.sub(one, parts[i])
    return '"'.join(parts)


def shift_sqref(sqref: str, dr: int) -> str:
    return " ".join(_shift_ref(p, dr) for p in sqref.split())


# --------------------------------------------------------------- الأنماط

def _items(section: Optional[str], tag: str) -> List[str]:
    if not section:
        return []
    return re.findall(r"<%s\b(?:[^>]*?/>|[^>]*?>.*?</%s>)" % (tag, tag), section, re.S)


def _section(xml: str, tag: str) -> Optional[str]:
    m = re.search(r"<%s\b[^>]*?(?:/>|>(.*?)</%s>)" % (tag, tag), xml, re.S)
    return (m.group(1) or "") if m else None


class StyleMerger:
    """جدول أنماط واحد يجمع أنماط كل الملفات دون تكرار."""

    def __init__(self):
        self.numfmts: Dict[int, str] = {}
        self.fonts: List[str] = []
        self.fills: List[str] = ['<fill><patternFill patternType="none"/></fill>',
                                 '<fill><patternFill patternType="gray125"/></fill>']
        self.borders: List[str] = []
        self.csxfs: List[str] = []
        self.cxfs: List[str] = []
        self.cellstyles: List[str] = []
        self.style_names: set = set()
        self.dxfs: List[str] = []
        self._idx: Dict[str, Dict[str, int]] = {k: {} for k in
                                                  ("font", "fill", "border", "csxf", "cxf", "dxf")}
        for i, f in enumerate(self.fills):
            self._idx["fill"][_canon(f)] = i
        self.ns: Dict[str, str] = {}
        self.ign: List[str] = []
        self.colors = ""
        self.table_styles = ""

    def _put(self, kind: str, lst: List[str], raw: str) -> int:
        key = _canon(raw)
        hit = self._idx[kind].get(key)
        if hit is None:
            hit = len(lst)
            lst.append(raw)
            self._idx[kind][key] = hit
        return hit

    def add(self, styles_xml: str) -> Tuple[List[int], List[int]]:
        """يضيف أنماط ملف. يرجع (خريطة أنماط الخلايا، خريطة أنماط التنسيق الشرطي)."""
        root = re.search(r"<styleSheet\b[^>]*>", styles_xml)
        if root:
            self.ns.update(_ns_decls(root.group(0)))
            for p in _ignorable(root.group(0)):
                if p not in self.ign:
                    self.ign.append(p)
        if not self.colors:
            m = re.search(r"<colors\b.*?</colors>", styles_xml, re.S)
            self.colors = m.group(0) if m else ""
        if not self.table_styles:
            m = re.search(r"<tableStyles\b[^>]*?(?:/>|>.*?</tableStyles>)", styles_xml, re.S)
            self.table_styles = m.group(0) if m else ""

        # تنسيقات الأرقام: المبنية في إكسل (< 164) ثابتة، والمخصّصة بنصّها
        nf_map: Dict[int, int] = {}
        for raw in _items(_section(styles_xml, "numFmts"), "numFmt"):
            try:
                old = int(_get_attr(raw, "numFmtId") or "0")
            except ValueError:
                continue
            code = _get_attr(raw, "formatCode") or ""
            same = next((i for i, c in self.numfmts.items() if c == code), None)
            if same is None:
                same = max([163] + list(self.numfmts)) + 1
                self.numfmts[same] = code
            nf_map[old] = same

        def mapped(lst_kind: str, lst: List[str], tag: str, section: str) -> List[int]:
            return [self._put(lst_kind, lst, raw) for raw in _items(_section(styles_xml, section), tag)]

        font_map = mapped("font", self.fonts, "font", "fonts")
        fill_map = mapped("fill", self.fills, "fill", "fills")
        border_map = mapped("border", self.borders, "border", "borders")

        def pick(mp: List[int], i: int) -> int:
            return mp[i] if 0 <= i < len(mp) else 0

        def remap_xf(raw: str, xf_map: Optional[List[int]]) -> str:
            end = raw.index(">") + 1
            head, rest = raw[:end], raw[end:]
            for attr, mp in (("fontId", font_map), ("fillId", fill_map),
                             ("borderId", border_map), ("xfId", xf_map)):
                v = _get_attr(head, attr)
                if v is not None and mp is not None:
                    try:
                        head = _set_attr(head, attr, str(pick(mp, int(v))))
                    except ValueError:
                        pass
            v = _get_attr(head, "numFmtId")
            if v is not None and v.isdigit() and int(v) in nf_map:
                head = _set_attr(head, "numFmtId", str(nf_map[int(v)]))
            return head + rest

        csxf_map = [self._put("csxf", self.csxfs, remap_xf(raw, None))
                    for raw in _items(_section(styles_xml, "cellStyleXfs"), "xf")]
        cxf_map = [self._put("cxf", self.cxfs, remap_xf(raw, csxf_map))
                   for raw in _items(_section(styles_xml, "cellXfs"), "xf")]
        for raw in _items(_section(styles_xml, "cellStyles"), "cellStyle"):
            name = _get_attr(raw, "name") or ""
            builtin = _get_attr(raw, "builtinId")
            # اسم النمط ورقمه المبني فريدان في المصنّف، وإلا أصلحه إكسل عند الفتح
            if name in self.style_names or (builtin is not None and
                                            ("#" + builtin) in self.style_names):
                continue
            if builtin is not None:
                self.style_names.add("#" + builtin)
            v = _get_attr(raw, "xfId")
            if v is not None and v.isdigit():
                raw = _set_attr(raw, "xfId", str(pick(csxf_map, int(v))))
            raw = re.sub(r'\s[\w]+:uid="[^"]*"', "", raw)
            self.style_names.add(name)
            self.cellstyles.append(raw)
        dxf_map = mapped("dxf", self.dxfs, "dxf", "dxfs")
        if not self.fonts:
            self.fonts.append('<font><sz val="11"/><name val="Calibri"/></font>')
        if not self.borders:
            self.borders.append("<border><left/><right/><top/><bottom/><diagonal/></border>")
        if not self.csxfs:
            self.csxfs.append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>')
        if not self.cxfs:
            self.cxfs.append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
        return cxf_map or [0], dxf_map

    def render(self) -> str:
        ns = "".join(' xmlns:%s="%s"' % kv for kv in sorted(self.ns.items()))
        ign = [p for p in self.ign if p in self.ns]
        if ign and "mc" in self.ns:
            ns += ' mc:Ignorable="%s"' % " ".join(ign)
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
               '<styleSheet xmlns="%s"%s>' % (MAIN_NS, ns)]
        if self.numfmts:
            out.append('<numFmts count="%d">%s</numFmts>' % (
                len(self.numfmts), "".join('<numFmt numFmtId="%d" formatCode="%s"/>'
                                           % (i, html.escape(html.unescape(c), quote=True))
                                           for i, c in sorted(self.numfmts.items()))))
        for tag, lst in (("fonts", self.fonts), ("fills", self.fills), ("borders", self.borders),
                         ("cellStyleXfs", self.csxfs), ("cellXfs", self.cxfs)):
            out.append('<%s count="%d">%s</%s>' % (tag, len(lst), "".join(lst), tag))
        styles = self.cellstyles or ['<cellStyle name="Normal" xfId="0" builtinId="0"/>']
        out.append('<cellStyles count="%d">%s</cellStyles>' % (len(styles), "".join(styles)))
        out.append('<dxfs count="%d">%s</dxfs>' % (len(self.dxfs), "".join(self.dxfs)))
        if self.table_styles:
            out.append(self.table_styles)
        if self.colors:
            out.append(self.colors)
        out.append("</styleSheet>")
        return "".join(out)


# --------------------------------------------------------- النصوص المشتركة

class SharedStrings:
    def __init__(self):
        self.items: List[str] = []
        self._idx: Dict[str, int] = {}
        self.refs = 0

    def add(self, sst_xml: Optional[str]) -> List[int]:
        out = []
        for raw in re.findall(r"<si\b[^>]*?(?:/>|>.*?</si>)", sst_xml or "", re.S):
            key = _canon(raw)
            hit = self._idx.get(key)
            if hit is None:
                hit = len(self.items)
                self.items.append(raw)
                self._idx[key] = hit
            out.append(hit)
        return out

    def render(self) -> str:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>'
                % (MAIN_NS, max(self.refs, len(self.items)), len(self.items),
                   "".join(self.items)))


def _si_text(raw: str) -> str:
    return html.unescape("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", raw, re.S)))


# ---------------------------------------------------------- الملف المصدر

class Source:
    """ملف واحد من ملفات الدمج: أوراقه وأنماطه ونصوصه."""

    def __init__(self, path: str, display: str):
        self.path = path
        self.display = display
        self.z = zipfile.ZipFile(path)
        names = set(self.z.namelist())
        self.names = names
        wb = self.z.read("xl/workbook.xml").decode("utf-8")
        rels = self.z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rid_target = {}
        for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
            rid_target[_get_attr(m.group(0), "Id")] = _get_attr(m.group(0), "Target") or ""
        self.sheets: List[Tuple[str, str]] = []
        for m in re.finditer(r"<sheet\b[^>]*/>", wb):
            tag = m.group(0)
            name = html.unescape(_get_attr(tag, "name") or "")
            rid = _get_attr(tag, "r:id") or ""
            part = _resolve("xl", rid_target.get(rid, ""))
            if part in names:
                self.sheets.append((name, part))
        self.defined = re.findall(r"<definedName\b[^>]*>.*?</definedName>", wb, re.S)
        self.ct = self.z.read("[Content_Types].xml").decode("utf-8")
        self.styles = (self.z.read("xl/styles.xml").decode("utf-8")
                       if "xl/styles.xml" in names else "")
        sst_part = next((_resolve("xl", t) for t in rid_target.values()
                         if t.lower().endswith("sharedstrings.xml")), "xl/sharedStrings.xml")
        self.sst_xml = (self.z.read(sst_part).decode("utf-8") if sst_part in names else "")
        self.sst_text = [_si_text(r) for r in
                         re.findall(r"<si\b[^>]*?(?:/>|>.*?</si>)", self.sst_xml, re.S)]

        mains = [(n, p) for n, p in self.sheets if not is_data_sheet(n)]
        datas = [(n, p) for n, p in self.sheets if is_data_sheet(n)]
        if not mains:
            raise ValueError("الملف «%s» لا يحتوي على ورقة ملاحظات." % display)
        self.main_name, self.main_part = mains[0]
        self.data_name, self.data_part = datas[0] if datas else ("", "")
        self.main_xml = self.z.read(self.main_part).decode("utf-8")
        self.data_xml = self.z.read(self.data_part).decode("utf-8") if self.data_part else ""
        self.code, self.notes = self._feeder()
        self.has_cell_images = any(n.startswith("xl/richData/") for n in names)

    def _feeder(self) -> Tuple[str, int]:
        """رمز المغذي (AH08-8911) وعدد الملاحظات من ورقة الملاحظات نفسها."""
        rows = re.findall(r"<row\b[^>]*?(?:/>|>(.*?)</row>)", self.main_xml, re.S)

        def cells(body: str) -> Dict[int, str]:
            out = {}
            for a, inner in re.findall(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", body or "", re.S):
                ref = _get_attr(" " + a, "r") or ""
                t = _get_attr(" " + a, "t") or ""
                if t == "inlineStr":
                    v = html.unescape("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", inner or "", re.S)))
                else:
                    m = re.search(r"<v>(.*?)</v>", inner or "", re.S)
                    v = html.unescape(m.group(1)) if m else ""
                    if t == "s" and v.isdigit():
                        i = int(v)
                        v = self.sst_text[i] if i < len(self.sst_text) else ""
                if ref:
                    out[col_index(ref)] = v.strip()
            return out

        hdr_i, cols = None, {}
        for i, body in enumerate(rows[:20]):
            c = cells(body)
            names = {_norm_name(v): k for k, v in c.items() if v}
            if "رقم الملاحظه" in names or "المغذي" in names:
                hdr_i = i
                cols = {"feeder": names.get("المغذي", names.get("المغذى")),
                        "station": names.get("المحطه"),
                        "id": names.get("رقم الملاحظه")}
                break
        if hdr_i is None:
            return "", 0
        recs, n = [], 0
        for body in rows[hdr_i + 1:]:
            c = cells(body)
            if cols.get("id") is not None and not c.get(cols["id"]):
                continue
            n += 1
            if len(recs) < 5:
                recs.append({"feeder": c.get(cols["feeder"], "") if cols["feeder"] is not None else "",
                             "station": c.get(cols["station"], "") if cols["station"] is not None else ""})
        try:
            code = builder.feeder_code(recs)
        except Exception:  # noqa: BLE001
            code = ""
        return code, n

    def close(self):
        self.z.close()


# ------------------------------------------------------- تحويل الأوراق

_ROW_OPEN = re.compile(r"<row\b[^>]*?/?>")
_CELL = re.compile(r"<c\b[^>]*?(?:/>|>.*?</c>)", re.S)
_F = re.compile(r"(<f\b[^>]*?(?<!/)>)(.*?)(</f>)", re.S)


class SheetMap:
    """كل ما يلزم لنقل ورقة من ملفها إلى الملف المدموج."""

    def __init__(self, xf: List[int], dxf: List[int], sst: List[int],
                 rename: Dict[str, str], self_name: str = "", dr: int = 0):
        self.xf, self.dxf, self.sst = xf, dxf, sst
        self.rename, self.self_name, self.dr = rename, self_name, dr

    def s(self, v: Optional[str]) -> Optional[str]:
        if v is None or not v.isdigit():
            return v
        i = int(v)
        return str(self.xf[i] if i < len(self.xf) else 0)

    def formula(self, text: str) -> str:
        return esc(shift_formula(html.unescape(text), self.dr, self.rename, self.self_name))


def transform_sheet_data(data: str, m: SheetMap, sst: SharedStrings) -> Tuple[str, int, int]:
    """يحوّل محتوى sheetData. يرجع (المحتوى، أول صف، آخر صف)."""
    first, last = [10 ** 9], [0]

    def row(mo: "re.Match") -> str:
        tag = mo.group(0)
        r = _get_attr(tag, "r")
        if r and r.isdigit():
            nr = int(r) + m.dr
            tag = _set_attr(tag, "r", str(nr))
            first[0] = min(first[0], nr)
            last[0] = max(last[0], nr)
        if _get_attr(tag, "s") is not None:
            tag = _set_attr(tag, "s", m.s(_get_attr(tag, "s")))
        return tag

    def cell(mo: "re.Match") -> str:
        raw = mo.group(0)
        end = raw.index(">") + 1
        head, body = raw[:end], raw[end:]
        ref = _get_attr(head, "r")
        if ref and m.dr:
            head = _set_attr(head, "r", _shift_ref(ref, m.dr))
        if _get_attr(head, "s") is not None:
            head = _set_attr(head, "s", m.s(_get_attr(head, "s")))
        # بيانات الخلية الوصفية (صور داخل الخلايا) لا تنتقل مع الورقة
        head = re.sub(r'\s(?:vm|cm)="\d+"', "", head)
        if _get_attr(head, "t") == "s":
            def v(mm):
                i = int(mm.group(1))
                sst.refs += 1
                return "<v>%d</v>" % (m.sst[i] if i < len(m.sst) else 0)
            body = re.sub(r"<v>(\d+)</v>", v, body, count=1)
        if "<f" in body:
            def f(mm):
                ftag = mm.group(1)
                fref = _get_attr(ftag, "ref")
                if fref and m.dr:
                    ftag = _set_attr(ftag, "ref", shift_sqref(fref, m.dr))
                return ftag + m.formula(mm.group(2)) + mm.group(3)
            body = _F.sub(f, body)
            body = re.sub(r"<f\b[^>]*/>", lambda mm: _set_attr(mm.group(0), "ref", shift_sqref(
                _get_attr(mm.group(0), "ref"), m.dr)) if _get_attr(mm.group(0), "ref") and m.dr
                else mm.group(0), body)
        return head + body

    out = _ROW_OPEN.sub(row, data)
    out = _CELL.sub(cell, out)
    return out, (first[0] if last[0] else 0), last[0]


def _split_sheet(xml: str) -> Tuple[str, str, str]:
    """(ما قبل sheetData، محتواها، ما بعدها)."""
    m = re.search(r"<sheetData\b[^>]*?(/>|>)", xml)
    if not m:
        return xml, "", ""
    if m.group(1) == "/>":
        return xml[:m.start()] + "<sheetData>", "", "</sheetData>" + xml[m.end():]
    end = xml.index("</sheetData>", m.end())
    return xml[:m.end()], xml[m.end():end], xml[end:]


def _remap_cf(xml: str, m: SheetMap, prio: List[int]) -> str:
    """التنسيق الشرطي: أرقام أنماطه، ونطاقه وصيغه بعد الإزاحة، وأولوياته."""
    def block(mo: "re.Match") -> str:
        b = mo.group(0)
        head_end = b.index(">") + 1
        head = b[:head_end]
        sq = _get_attr(head, "sqref")
        if sq and m.dr:
            head = _set_attr(head, "sqref", shift_sqref(sq, m.dr))
        body = b[head_end:]

        def rule(rm: "re.Match") -> str:
            tag = rm.group(0)
            d = _get_attr(tag, "dxfId")
            if d is not None and d.isdigit():
                i = int(d)
                tag = _set_attr(tag, "dxfId", str(m.dxf[i] if i < len(m.dxf) else 0))
            if prio is not None:
                prio[0] += 1
                tag = _set_attr(tag, "priority", str(prio[0]))
            return tag
        body = re.sub(r"<cfRule\b[^>]*>", rule, body)
        body = re.sub(r"(<formula>)(.*?)(</formula>)",
                      lambda fm: fm.group(1) + m.formula(fm.group(2)) + fm.group(3), body, flags=re.S)
        return head + body
    return re.sub(r"<conditionalFormatting\b.*?</conditionalFormatting>", block, xml, flags=re.S)


# ------------------------------------------------------------- الداتا شيت

def _elem(xml: str, tag: str) -> str:
    m = re.search(r"<%s\b[^>]*?(?:/>|>.*?</%s>)" % (tag, tag), xml, re.S)
    return m.group(0) if m else ""


def build_data_sheet(blocks: List[Dict[str, Any]], sst: SharedStrings) -> Tuple[str, List[int]]:
    """داتا شيت كل ملف تحت الذي قبله في ورقة واحدة. يرجع (XML، بدايات الكتل)."""
    first_xml = blocks[0]["xml"]
    ns: Dict[str, str] = {}
    ign: List[str] = []
    for b in blocks:
        root = re.search(r"<worksheet\b[^>]*>", b["xml"]).group(0)
        ns.update(_ns_decls(root))
        for p in _ignorable(root):
            if p not in ign:
                ign.append(p)
    ns.setdefault("r", REL_NS)
    root = '<worksheet xmlns="%s"%s' % (MAIN_NS, "".join(' xmlns:%s="%s"' % kv for kv in sorted(ns.items())))
    ign = [p for p in ign if p in ns]
    if ign and "mc" in ns:
        root += ' mc:Ignorable="%s"' % " ".join(ign)
    root += ">"

    rows: List[str] = []
    merges: List[str] = []
    cfs: List[str] = []
    x14: List[str] = []
    starts: List[int] = []
    prio = [0]
    next_row = 1
    max_col = 0
    for b in blocks:
        _pre, data, post = _split_sheet(b["xml"])
        rnums = [int(x) for x in re.findall(r'<row\b[^>]*?\sr="(\d+)"', data)]
        top = min(rnums) if rnums else 1
        m: SheetMap = b["map"]
        m.dr = next_row - top
        starts.append(next_row)
        body, _f, last = transform_sheet_data(data, m, sst)
        rows.append(body)
        for c in re.findall(r'<c\b[^>]*?\sr="([A-Z]+)\d+"', body):
            max_col = max(max_col, col_index(c))
        for mc in re.findall(r'<mergeCell\b[^>]*/>', post):
            merges.append(_set_attr(mc, "ref", shift_sqref(_get_attr(mc, "ref") or "", m.dr)))
        cf_xml = "".join(re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", post, re.S))
        cfs.append(_remap_cf(cf_xml, m, prio))
        # أشرطة البيانات بصيغة إكسل 2010 تُحفظ في extLst بنطاقها: تُزاح مثلها
        for item in re.findall(r"<x14:conditionalFormatting\b.*?</x14:conditionalFormatting>", post, re.S):
            item = re.sub(r"(<xm:sqref>)(.*?)(</xm:sqref>)",
                          lambda mm: mm.group(1) + shift_sqref(mm.group(2), m.dr) + mm.group(3), item)
            item = re.sub(r"(<xm:f>)(.*?)(</xm:f>)",
                          lambda mm: mm.group(1) + m.formula(mm.group(2)) + mm.group(3), item, flags=re.S)
            x14.append(item)
        next_row = (last or next_row) + GAP + 1

    last_row = next_row - GAP - 1
    sheet_pr = _elem(first_xml, "sheetPr")
    views = _elem(first_xml, "sheetViews")
    views = re.sub(r"<selection\b[^>]*/>", "", views)
    views = re.sub(r"<pane\b[^>]*/>", "", views)
    if "tabSelected" in views:
        views = re.sub(r'tabSelected="\d"', 'tabSelected="1"', views)
    else:
        views = re.sub(r"<sheetView\b", '<sheetView tabSelected="1"', views, count=1)
    fmt = _elem(first_xml, "sheetFormatPr")
    cols = _elem(first_xml, "cols")
    first_map: SheetMap = blocks[0]["map"]
    cols = re.sub(r'(<col\b[^>]*?\sstyle=")(\d+)(")',
                  lambda mm: mm.group(1) + first_map.s(mm.group(2)) + mm.group(3), cols)
    print_opts = _elem(first_xml, "printOptions")
    margins = _elem(first_xml, "pageMargins")
    setup = re.sub(r'\sr:id="[^"]*"', "", _elem(first_xml, "pageSetup"))
    hf = _elem(first_xml, "headerFooter")
    # كل مغذٍّ في صفحة مطبوعة: فاصل صفحة قبل كل داتا شيت بعد الأولى
    brks = "".join('<brk id="%d" max="16383" man="1"/>' % (s - 1) for s in starts[1:])
    row_breaks = ('<rowBreaks count="%d" manualBreakCount="%d">%s</rowBreaks>'
                  % (len(starts) - 1, len(starts) - 1, brks)) if brks else ""
    ext = ""
    if x14:
        ext = ('<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" '
               'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
               "<x14:conditionalFormattings>%s</x14:conditionalFormattings></ext></extLst>"
               % "".join(x14))
    dim = '<dimension ref="A1:%s%d"/>' % (col_letter(max(max_col, 0)), max(last_row, 1))
    xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n', root,
           sheet_pr, dim, views, fmt, cols, "<sheetData>", "".join(rows), "</sheetData>"]
    if merges:
        xml.append('<mergeCells count="%d">%s</mergeCells>' % (len(merges), "".join(merges)))
    xml += cfs + [print_opts, margins, setup, hf, row_breaks, ext, "</worksheet>"]
    return "".join(xml), starts


# ------------------------------------------------------- أوراق الملاحظات

def build_main_sheet(src: Source, m: SheetMap, sst: SharedStrings) -> str:
    xml = src.main_xml
    pre, data, post = _split_sheet(xml)
    body, _f, _l = transform_sheet_data(data, m, sst)
    pre = re.sub(r'\stabSelected="\d"', "", pre)
    pre = re.sub(r'(<col\b[^>]*?\sstyle=")(\d+)(")',
                 lambda mm: mm.group(1) + m.s(mm.group(2)) + mm.group(3), pre)
    post = _remap_cf(post, m, None)
    post = re.sub(r"(<formula[12]>)(.*?)(</formula[12]>)",
                  lambda mm: mm.group(1) + m.formula(mm.group(2)) + mm.group(3), post, flags=re.S)
    # ما لا يُنقل من علاقات الورقة يُحذف وسمه، وإلا أشار إلى جزء غير موجود
    return pre + body + post


class PartCopier:
    """ينسخ أجزاء ورقة (روابط، إعدادات طابعة، رسوم، ملاحظات، جداول) مع علاقاتها."""

    def __init__(self):
        self.parts: Dict[str, bytes] = {}
        self.overrides: Dict[str, str] = {}
        self.defaults: Dict[str, str] = {}
        self.dropped: List[str] = []
        self._table_id = 0

    def _unique(self, path: str) -> str:
        if path not in self.parts:
            return path
        d, b = posixpath.split(path)
        stem, ext = posixpath.splitext(b)
        stem = re.sub(r"\d+$", "", stem)
        i = 1
        while True:
            cand = posixpath.join(d, "%s%d%s" % (stem, i, ext))
            if cand not in self.parts:
                return cand
            i += 1

    def copy_rels(self, src: Source, part: str, new_part: str, memo: Dict[str, str]) -> Tuple[Optional[str], set]:
        """ينسخ علاقات جزء. يرجع (XML علاقاته الجديدة، معرّفات حُذفت)."""
        rp = _part_rels(part)
        if rp not in src.names:
            return None, set()
        rels = src.z.read(rp).decode("utf-8")
        out, gone = [], set()
        for mo in re.finditer(r"<Relationship\b[^>]*/>", rels):
            tag = mo.group(0)
            rid = _get_attr(tag, "Id") or ""
            typ = _get_attr(tag, "Type") or ""
            if (_get_attr(tag, "TargetMode") or "") == "External":
                out.append(tag)
                continue
            if any(typ.endswith("/" + s) for s in SKIP_RELS):
                gone.add(rid)
                continue
            target = _resolve(posixpath.dirname(part), html.unescape(_get_attr(tag, "Target") or ""))
            if target not in src.names:
                gone.add(rid)
                continue
            if target not in memo:
                nt = self._unique(target)
                memo[target] = nt
                data = src.z.read(target)
                ct = self._content_type(src, target)
                if ct == CT_TABLE:
                    data = self._fix_table(data)
                self.parts[nt] = data
                if ct:
                    self.overrides["/" + nt] = ct
                sub, _g = self.copy_rels(src, target, nt, memo)
                if sub is not None:
                    self.parts[_part_rels(nt)] = sub.encode("utf-8")
            rel_target = posixpath.relpath(memo[target], posixpath.dirname(new_part))
            out.append(_set_attr(tag, "Target", esc(rel_target)))
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<Relationships xmlns="%s">%s</Relationships>' % (PKG_REL_NS, "".join(out)))
        return xml, gone

    def _content_type(self, src: Source, part: str) -> str:
        m = re.search(r'<Override\b[^>]*PartName="/%s"[^>]*/>' % re.escape(part), src.ct)
        if m:
            return _get_attr(m.group(0), "ContentType") or ""
        ext = posixpath.splitext(part)[1].lstrip(".").lower()
        m = re.search(r'<Default\b[^>]*Extension="%s"[^>]*/>' % re.escape(ext), src.ct, re.I)
        if m:
            self.defaults[ext] = _get_attr(m.group(0), "ContentType") or ""
        return ""

    def _fix_table(self, data: bytes) -> bytes:
        """أرقام الجداول وأسماؤها فريدة في المصنّف كله."""
        self._table_id += 1
        x = data.decode("utf-8")
        head = re.search(r"<table\b[^>]*>", x).group(0)
        new = _set_attr(head, "id", str(self._table_id))
        for a in ("name", "displayName"):
            v = _get_attr(head, a)
            if v:
                new = _set_attr(new, a, "%s_%d" % (v, self._table_id))
        return x.replace(head, new, 1).encode("utf-8")


def _drop_rel_elements(xml: str, gone: set) -> str:
    for rid in gone:
        xml = re.sub(r'<\w+\b[^>]*\sr:id="%s"[^>]*/>' % re.escape(rid), "", xml)
    return xml


# ------------------------------------------------------------------ الدمج

def merge(files: List[Tuple[str, str]], out_path: str) -> Dict[str, Any]:
    """files: [(المسار، الاسم المعروض)]. يكتب الملف المدموج ويرجع ملخّصه."""
    sources: List[Source] = []
    warnings: List[str] = []
    try:
        for p, name in files:
            try:
                sources.append(Source(p, name))
            except zipfile.BadZipFile:
                raise ValueError("الملف «%s» ليس ملف إكسل سليمًا." % name)
            except KeyError:
                raise ValueError("الملف «%s» ناقص الأجزاء — افتحه في إكسل واحفظه ثم أعد رفعه." % name)
        if not sources:
            raise ValueError("لم تُرفع ملفات.")

        sources.sort(key=lambda s: sorter.natural_key(s.code or s.display))

        # اسم كل ورقة: رمز المغذي، وإن تكرّر المغذي أُلحق به رقم
        used = {DATA_TITLE}
        for s in sources:
            base = safe_sheet_name(s.code or os.path.splitext(s.display)[0])
            name, k = base, 2
            while name in used:
                name = safe_sheet_name("%s (%d)" % (base[:26], k))
                k += 1
            if name != base:
                warnings.append("المغذي %s مكرّر — ورقة الملف «%s» باسم «%s»."
                                % (base, s.display, name))
            used.add(name)
            s.sheet_name = name
            if not s.data_part:
                warnings.append("الملف «%s» بلا داتا شيت — شغّله على أداة حالة الاعتماد أولًا ليظهر في ورقة الداتا شيت."
                                % s.display)
            if s.has_cell_images:
                warnings.append("في «%s» صور داخل الخلايا لا تُنقل — روابط الصور تبقى." % s.display)

        styles = StyleMerger()
        sst = SharedStrings()
        maps = []
        for s in sources:
            xf, dxf = styles.add(s.styles or "")
            sm = sst.add(s.sst_xml)
            rename = {s.main_name: s.sheet_name}
            if s.data_name:
                rename[s.data_name] = DATA_TITLE
            maps.append((xf, dxf, sm, rename))

        copier = PartCopier()
        sheets_out: List[Tuple[str, str]] = []        # (الاسم، XML)
        sheet_rels: Dict[int, str] = {}
        defined: List[str] = []

        # الداتا شيت المدموجة أولًا
        blocks = []
        for s, (xf, dxf, sm, rename) in zip(sources, maps):
            if s.data_xml:
                blocks.append({"xml": s.data_xml, "src": s,
                               "map": SheetMap(xf, dxf, sm, rename, self_name=s.data_name)})
        starts: List[int] = []
        if blocks:
            data_xml, starts = build_data_sheet(blocks, sst)
            sheets_out.append((DATA_TITLE, data_xml))

        for s, (xf, dxf, sm, rename) in zip(sources, maps):
            idx = len(sheets_out)
            new_part = "xl/worksheets/sheet%d.xml" % (idx + 1)
            m = SheetMap(xf, dxf, sm, rename, self_name=s.main_name)
            xml = build_main_sheet(s, m, sst)
            rels, gone = copier.copy_rels(s, s.main_part, new_part, {})
            if gone:
                xml = _drop_rel_elements(xml, gone)
            if rels is not None:
                sheet_rels[idx] = rels
            sheets_out.append((s.sheet_name, xml))
            # أسماء معرّفة خاصة بالورقة (نطاق التصفية، منطقة الطباعة)
            src_idx = [n for n, _p in s.sheets].index(s.main_name)
            for dn in s.defined:
                head = re.match(r"<definedName\b[^>]*>", dn).group(0)
                if _get_attr(head, "localSheetId") != str(src_idx):
                    continue
                text = re.sub(r"^<definedName\b[^>]*>|</definedName>$", "", dn)
                text = m.formula(text)
                head = _set_attr(head, "localSheetId", str(idx))
                defined.append(head + text + "</definedName>")

        # ---------------------------------------------------- كتابة الحزمة
        ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>']
        for ext, typ in sorted(copier.defaults.items()):
            if ext not in ("rels", "xml") and typ:
                ct.append('<Default Extension="%s" ContentType="%s"/>' % (ext, typ))
        ct.append('<Override PartName="/xl/workbook.xml" ContentType="application/'
                  'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
        ct.append('<Override PartName="/xl/styles.xml" ContentType="application/'
                  'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
        ct.append('<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                  'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
        theme = None
        for s in sources:
            if "xl/theme/theme1.xml" in s.names:
                theme = s.z.read("xl/theme/theme1.xml")
                break
        if theme:
            ct.append('<Override PartName="/xl/theme/theme1.xml" ContentType="application/'
                      'vnd.openxmlformats-officedocument.theme+xml"/>')
        for i in range(len(sheets_out)):
            ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%s"/>'
                      % (i + 1, CT_SHEET))
        for pn, typ in sorted(copier.overrides.items()):
            ct.append('<Override PartName="%s" ContentType="%s"/>' % (pn, typ))
        ct.append("</Types>")

        wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   '<Relationships xmlns="%s">' % PKG_REL_NS]
        sheets_tag = []
        for i, (name, _x) in enumerate(sheets_out):
            wb_rels.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>'
                           % (i + 1, REL_NS, i + 1))
            sheets_tag.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                              % (html.escape(name, quote=True), i + 1, i + 1))
        n = len(sheets_out)
        wb_rels.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>' % (n + 1, REL_NS))
        wb_rels.append('<Relationship Id="rId%d" Type="%s/sharedStrings" Target="sharedStrings.xml"/>'
                       % (n + 2, REL_NS))
        if theme:
            wb_rels.append('<Relationship Id="rId%d" Type="%s/theme" Target="theme/theme1.xml"/>'
                           % (n + 3, REL_NS))
        wb_rels.append("</Relationships>")

        workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<workbook xmlns="%s" xmlns:r="%s"><workbookPr/>'
                    '<bookViews><workbookView activeTab="0" tabRatio="750"/></bookViews>'
                    "<sheets>%s</sheets>%s"
                    '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'
                    % (MAIN_NS, REL_NS, "".join(sheets_tag),
                       "<definedNames>%s</definedNames>" % "".join(defined) if defined else ""))

        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            zout.writestr("[Content_Types].xml", "".join(ct))
            zout.writestr("_rels/.rels",
                          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                          '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/officeDocument" '
                          'Target="xl/workbook.xml"/></Relationships>' % (PKG_REL_NS, REL_NS))
            zout.writestr("xl/workbook.xml", workbook)
            zout.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
            zout.writestr("xl/styles.xml", styles.render())
            zout.writestr("xl/sharedStrings.xml", sst.render())
            if theme:
                zout.writestr("xl/theme/theme1.xml", theme)
            for i, (_name, xml) in enumerate(sheets_out):
                zout.writestr("xl/worksheets/sheet%d.xml" % (i + 1), xml)
                if i in sheet_rels:
                    zout.writestr("xl/worksheets/_rels/sheet%d.xml.rels" % (i + 1), sheet_rels[i])
            for p, data in copier.parts.items():
                zout.writestr(p, data)

        return {
            "sheets": [{"name": s.sheet_name, "file": s.display, "notes": s.notes,
                        "data": bool(s.data_part)} for s in sources],
            "data_blocks": len(blocks),
            "warnings": warnings,
            "codes": [s.code for s in sources if s.code],
        }
    finally:
        for s in sources:
            s.close()
