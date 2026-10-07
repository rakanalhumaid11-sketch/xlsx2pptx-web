# -*- coding: utf-8 -*-
"""
xlsxedit.py
===========
تعديل ملف إكسل **دون إعادة بنائه**.

نفتح الملف كأرشيف مضغوط، ونغيّر الأجزاء التي تحتاج تغييرًا فقط (ورقة واحدة
+ جدول الأنماط + فهرس الأوراق)، وننسخ كل جزء آخر كما هو بايتًا ببايت. لذلك
تبقى الصور المدمجة في الخلايا، وعروض الأعمدة، وألوان المستخدم، وترتيب
الأعمدة، والأوراق الأخرى، والصيغ — كما تركها صاحب الملف تمامًا.

(openpyxl كان يعيد كتابة الملف كله فيُسقط الصور والرسوم ويستبدل تنسيق
المستخدم بتنسيقه، ولهذا لا يُستعمل هنا.)
"""

import re
import zipfile
from typing import Any, Dict, List, Optional, Tuple

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WS_TYPE = ("application/vnd.openxmlformats-officedocument."
           "spreadsheetml.worksheet+xml")


# ------------------------------------------------------------------ أدوات عامة

def col_letter(i: int) -> str:
    """0 -> A، 26 -> AA"""
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(ref: str) -> int:
    """'AB12' -> 27 (صفري الأساس)"""
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def esc(v: Any) -> str:
    return (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _set_attr(tag: str, name: str, value: str) -> str:
    """يضبط قيمة خاصية داخل نص وسم مفرد، ويضيفها إن لم تكن موجودة."""
    pat = re.compile(r'\s%s="[^"]*"' % re.escape(name))
    if pat.search(tag):
        return pat.sub(' %s="%s"' % (name, value), tag, count=1)
    if tag.endswith("/>"):
        return tag[:-2] + ' %s="%s"/>' % (name, value)
    return tag[:-1] + ' %s="%s">' % (name, value)


def _get_attr(tag: str, name: str) -> Optional[str]:
    m = re.search(r'\s%s="([^"]*)"' % re.escape(name), tag)
    return m.group(1) if m else None


# ------------------------------------------------------------------ الأنماط

class Styles:
    """يضيف خطوطًا وتعبئات وحدودًا وأنماط خلايا إلى styles.xml الأصلي.

    لا نحذف ولا نعدّل أي نمط قائم — نضيف فقط في نهاية كل قائمة، فتبقى
    أرقام الأنماط التي تشير إليها خلايا المستخدم صحيحة كما هي."""

    CHILD_OF = {"numFmt": "numFmts", "font": "fonts", "fill": "fills",
                "border": "borders", "xf": "cellXfs", "dxf": "dxfs"}
    ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs",
             "cellXfs", "cellStyles", "dxfs", "tableStyles"]

    def __init__(self, xml: str):
        self.xml = xml
        self.added: Dict[str, List[str]] = {k: [] for k in self.CHILD_OF}
        self.base: Dict[str, int] = {}
        for child, cont in self.CHILD_OF.items():
            inner = self._inner(cont)
            self.base[child] = len(re.findall(r"<%s\b" % child, inner))
        self._xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>",
                               self._inner("cellXfs"), re.S)
        self._fill_cache: Dict[str, int] = {}
        self._tinted: Dict[Tuple[int, str], int] = {}
        self._next_numfmt = self._max_numfmt() + 1

    # ---- قراءة
    def _span(self, tag: str) -> Optional[Tuple[int, int, str, bool]]:
        m = re.search(r"<%s\b[^>]*?/>" % tag, self.xml)
        if m:
            return m.start(), m.end(), "", True
        m = re.search(r"<%s\b[^>]*?>" % tag, self.xml)
        if not m:
            return None
        close = "</%s>" % tag
        end = self.xml.index(close, m.end())
        return m.start(), end + len(close), self.xml[m.end():end], False

    def _inner(self, tag: str) -> str:
        sp = self._span(tag)
        return sp[2] if sp else ""

    def _max_numfmt(self) -> int:
        ids = [int(i) for i in re.findall(r'<numFmt\b[^>]*numFmtId="(\d+)"',
                                          self._inner("numFmts"))]
        return max(ids + [163])

    # ---- إضافة
    def fill(self, hex_rgb: str) -> int:
        """تعبئة صلبة بلون معيّن — تُعاد نفس التعبئة إن طُلبت مرتين."""
        key = hex_rgb.upper()
        if key in self._fill_cache:
            return self._fill_cache[key]
        self.added["fill"].append(
            '<fill><patternFill patternType="solid">'
            '<fgColor rgb="FF%s"/><bgColor indexed="64"/></patternFill></fill>' % key)
        idx = self.base["fill"] + len(self.added["fill"]) - 1
        self._fill_cache[key] = idx
        return idx

    def font(self, size: int = 11, bold: bool = False, color: str = "000000",
             name: str = "Calibri") -> int:
        self.added["font"].append(
            '<font><sz val="%d"/>%s<color rgb="FF%s"/><name val="%s"/>'
            '<family val="2"/></font>' % (size, "<b/>" if bold else "", color.upper(), name))
        return self.base["font"] + len(self.added["font"]) - 1

    def border(self, color: str = "BFCAD4", sides: str = "lrtb",
               thick_bottom: Optional[str] = None,
               thick_top: Optional[str] = None) -> int:
        def side(tag, on, col, style="thin"):
            if not on:
                return "<%s/>" % tag
            return '<%s style="%s"><color rgb="FF%s"/></%s>' % (tag, style, col.upper(), tag)
        self.added["border"].append(
            "<border>%s%s%s%s<diagonal/></border>" % (
                side("left", "l" in sides, color),
                side("right", "r" in sides, color),
                side("top", True, thick_top, "medium") if thick_top
                else side("top", "t" in sides, color),
                side("bottom", True, thick_bottom, "medium") if thick_bottom
                else side("bottom", "b" in sides, color)))
        return self.base["border"] + len(self.added["border"]) - 1

    def numfmt(self, code: str) -> int:
        nid = self._next_numfmt
        self._next_numfmt += 1
        self.added["numFmt"].append('<numFmt numFmtId="%d" formatCode="%s"/>'
                                    % (nid, esc(code)))
        return nid

    def xf(self, font: int = 0, fill: int = 0, border: int = 0, numfmt: int = 0,
           halign: str = "", valign: str = "center", wrap: bool = False,
           indent: int = 0, rtl: bool = False) -> int:
        """`rtl` يثبّت اتجاه القراءة من اليمين: نص عربي فيه «D1» أو رقم يقلبه
        بعض البرامج إذا تُرك الاتجاه لتخمين السياق."""
        align = ""
        if halign or valign or wrap or indent or rtl:
            align = ('<alignment%s%s%s%s%s/>' % (
                ' horizontal="%s"' % halign if halign else "",
                ' vertical="%s"' % valign if valign else "",
                ' wrapText="1"' if wrap else "",
                ' indent="%d"' % indent if indent else "",
                ' readingOrder="2"' if rtl else ""))
        self.added["xf"].append(
            '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"'
            ' applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"'
            ' applyAlignment="1">%s</xf>' % (numfmt, font, fill, border, align))
        return self.base["xf"] + len(self.added["xf"]) - 1

    def tinted(self, base_xf: int, hex_rgb: str) -> int:
        """نمط جديد = نمط الخلية الأصلي نفسه مع تغيير لون التعبئة وحده.

        هكذا يحتفظ الصف بخطه وحدوده وتنسيق أرقامه كما كان."""
        key = (base_xf, hex_rgb.upper())
        if key in self._tinted:
            return self._tinted[key]
        src = (self._xfs[base_xf] if 0 <= base_xf < len(self._xfs)
               else '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
        head = src if "</xf>" not in src else src[:src.index(">") + 1]
        new = _set_attr(head, "fillId", str(self.fill(hex_rgb)))
        new = _set_attr(new, "applyFill", "1")
        if "</xf>" in src:                       # نمط له عنصر محاذاة داخلي
            new += src[src.index(">") + 1:]
        self.added["xf"].append(new)
        idx = self.base["xf"] + len(self.added["xf"]) - 1
        self._tinted[key] = idx
        return idx

    def dxf_fill(self, hex_rgb: str) -> int:
        """تنسيق شرطي: تعبئة فقط. (في dxf يكون اللون في bgColor لا fgColor.)"""
        key = "dxf:" + hex_rgb.upper()
        if key in self._fill_cache:
            return self._fill_cache[key]
        self.added["dxf"].append(
            '<dxf><fill><patternFill><bgColor rgb="FF%s"/></patternFill>'
            "</fill></dxf>" % hex_rgb.upper())
        idx = self.base["dxf"] + len(self.added["dxf"]) - 1
        self._fill_cache[key] = idx
        return idx

    # ---- كتابة
    def render(self) -> str:
        xml = self.xml
        for child, cont in self.CHILD_OF.items():
            items = self.added[child]
            if not items:
                continue
            total = self.base[child] + len(items)
            # المدى يُحسب على النص الحالي لا الأصلي: إدراجات سابقة أزاحت المواضع
            sp = self._span_in(xml, cont)
            if sp is None:
                block = '<%s count="%d">%s</%s>' % (cont, total, "".join(items), cont)
                xml = self._insert_container(xml, cont, block)
            else:
                start, end, inner, _ = sp
                block = '<%s count="%d">%s%s</%s>' % (cont, total, inner,
                                                      "".join(items), cont)
                xml = xml[:start] + block + xml[end:]
        return xml

    @staticmethod
    def _span_in(xml: str, tag: str):
        m = re.search(r"<%s\b[^>]*?/>" % tag, xml)
        if m:
            return m.start(), m.end(), "", True
        m = re.search(r"<%s\b[^>]*?>" % tag, xml)
        if not m:
            return None
        close = "</%s>" % tag
        end = xml.index(close, m.end())
        return m.start(), end + len(close), xml[m.end():end], False

    def _insert_container(self, xml: str, cont: str, block: str) -> str:
        """يُدرج قائمة غائبة (مثل numFmts) في موضعها الصحيح من الترتيب."""
        after = self.ORDER[self.ORDER.index(cont) + 1:]
        for nxt in after:
            m = re.search(r"<%s\b" % nxt, xml)
            if m:
                return xml[:m.start()] + block + xml[m.start():]
        return xml.replace("</styleSheet>", block + "</styleSheet>")


# ------------------------------------------------------------------ بناء ورقة

class SheetBuilder:
    """يبني XML ورقة جديدة من خلايا بسيطة (نصوص وأرقام) بأنماط معطاة."""

    def __init__(self, rtl: bool = True, gridlines: bool = False,
                 tab_color: str = "", landscape: bool = False,
                 selected: bool = False, centered: bool = False):
        self.rtl = rtl
        self.gridlines = gridlines
        self.tab_color = tab_color
        self.landscape = landscape
        # الورقة المحدَّدة هي ما يصدّره إكسل عند «حفظ كـ PDF»، والتوسيط يجعل
        # الصفحة المطبوعة متّزنة بدل أن تلتصق بحافة
        self.selected = selected
        self.centered = centered
        self.cells: Dict[Tuple[int, int], Tuple[Any, Optional[int], str]] = {}
        self.merges: List[str] = []
        self.widths: Dict[int, float] = {}
        self.heights: Dict[int, float] = {}
        self.databars: List[Tuple[str, str]] = []
        self.drawing_rid = ""          # يُضبط عند إرفاق رسم بياني بالورقة

    def set(self, row: int, col: int, value: Any, style: Optional[int] = None,
            formula: str = ""):
        """`formula` يجعل الخلية حيّة: تُعاد الحسبة كلما عدّل المستخدم البيانات.
        نكتب معها القيمة المحسوبة الآن كي تظهر صحيحة فورًا قبل أي حساب."""
        self.cells[(row, col)] = (value, style, formula)

    def blank(self, row: int, col: int, style: Optional[int] = None):
        self.cells.setdefault((row, col), (None, style, ""))

    def merge(self, r1: int, c1: int, r2: int, c2: int):
        self.merges.append("%s%d:%s%d" % (col_letter(c1), r1, col_letter(c2), r2))

    def width(self, col: int, w: float):
        self.widths[col] = w

    def height(self, row: int, h: float):
        self.heights[row] = h

    def databar(self, sqref: str, color: str = "4CAF7D"):
        self.databars.append((sqref, color))

    def _cell_xml(self, row: int, col: int) -> str:
        value, style, formula = self.cells[(row, col)]
        ref = "%s%d" % (col_letter(col), row)
        s = ' s="%d"' % style if style is not None else ""
        if formula:
            cached = value if isinstance(value, (int, float)) and not isinstance(value, bool) else 0
            return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (ref, s, esc(formula), cached)
        if value is None or value == "":
            return '<c r="%s"%s/>' % (ref, s)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return '<c r="%s"%s><v>%s</v></c>' % (ref, s, value)
        return ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t>'
                '</is></c>' % (ref, s, esc(value)))

    def to_xml(self) -> str:
        rows = sorted({r for r, _ in self.cells})
        cols = sorted({c for _, c in self.cells})
        dim = "A1" if not rows else "%s%d:%s%d" % (
            col_letter(min(cols)), min(rows), col_letter(max(cols)), max(rows))

        cols_xml = ""
        if self.widths:
            cols_xml = "<cols>%s</cols>" % "".join(
                '<col min="%d" max="%d" width="%.2f" customWidth="1"/>'
                % (c + 1, c + 1, w) for c, w in sorted(self.widths.items()))

        body = []
        for r in rows:
            rc = sorted(c for rr, c in self.cells if rr == r)
            h = ' ht="%.2f" customHeight="1"' % self.heights[r] if r in self.heights else ""
            body.append('<row r="%d" spans="%d:%d"%s>%s</row>' % (
                r, min(rc) + 1, max(rc) + 1, h,
                "".join(self._cell_xml(r, c) for c in rc)))

        merges = ""
        if self.merges:
            merges = '<mergeCells count="%d">%s</mergeCells>' % (
                len(self.merges),
                "".join('<mergeCell ref="%s"/>' % m for m in self.merges))

        cf = "".join(
            '<conditionalFormatting sqref="%s"><cfRule type="dataBar" priority="%d">'
            '<dataBar><cfvo type="num" val="0"/><cfvo type="num" val="1"/>'
            '<color rgb="FF%s"/></dataBar></cfRule></conditionalFormatting>'
            % (sq, i + 1, color.upper()) for i, (sq, color) in enumerate(self.databars))

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s" xmlns:r="%s">'
            '%s<dimension ref="%s"/>'
            '<sheetViews><sheetView%s%s%s workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="16.5"/>'
            '%s<sheetData>%s</sheetData>%s%s'
            '%s<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5"'
            ' header="0.3" footer="0.3"/>'
            '<pageSetup orientation="%s" fitToWidth="1" fitToHeight="0"/>'
            '%s</worksheet>' % (
                MAIN_NS, REL_NS,
                # fitToWidth لا يعمل إلا مع pageSetUpPr، وبدونه تنقسم الورقة صفحتين
                "<sheetPr>%s<pageSetUpPr fitToPage=\"1\"/></sheetPr>" % (
                    '<tabColor rgb="FF%s"/>' % self.tab_color.upper()
                    if self.tab_color else ""), dim,
                ' rightToLeft="1"' if self.rtl else "",
                ' showGridLines="0"' if not self.gridlines else "",
                ' tabSelected="1"' if self.selected else "",
                cols_xml, "".join(body), merges, cf,
                '<printOptions horizontalCentered="1"/>' if self.centered else "",
                "landscape" if self.landscape else "portrait",
                '<drawing r:id="%s"/>' % self.drawing_rid if self.drawing_rid else ""))


# ------------------------------------------------------------------ تعديل ورقة

_ROW_RE = re.compile(r"<row\b[^>]*/>|<row\b[^>]*>.*?</row>", re.S)
_CELL_RE = re.compile(r"<c\b[^>]*/>|<c\b[^>]*>.*?</c>", re.S)


def _row_number(row_xml: str, fallback: int) -> int:
    r = _get_attr(row_xml.split(">", 1)[0], "r")
    return int(r) if r else fallback


def _split_row(row_xml: str) -> Tuple[str, str, str]:
    """(وسم الفتح، المحتوى، وسم الإغلاق)"""
    if row_xml.endswith("/>"):
        return row_xml[:-2] + ">", "", "</row>"
    head_end = row_xml.index(">") + 1
    return row_xml[:head_end], row_xml[head_end:-len("</row>")], "</row>"


def _cells_by_col(inner: str) -> Dict[int, str]:
    out: Dict[int, str] = {}
    auto = 0
    for m in _CELL_RE.finditer(inner):
        cell = m.group(0)
        ref = _get_attr(cell.split(">", 1)[0], "r")
        c = col_index(ref) if ref else auto
        auto = c + 1
        out[c] = cell
    return out


def _cell_style(cell_xml: str) -> int:
    s = _get_attr(cell_xml.split(">", 1)[0], "s")
    return int(s) if s else 0


def _text_cell(ref: str, style: Optional[int], value: Any) -> str:
    s = ' s="%d"' % style if style is not None else ""
    if value is None or value == "":
        return '<c r="%s"%s/>' % (ref, s)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, value)
    return ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, s, esc(value)))


#  العناصر التي تأتي — بحسب مواصفة الملف — بعد التنسيق الشرطي والتحقق
_AFTER_CF = ["<dataValidations", "<hyperlinks", "<printOptions", "<pageMargins",
             "<pageSetup", "<headerFooter", "<rowBreaks", "<colBreaks",
             "<customProperties", "<cellWatches", "<ignoredErrors", "<drawing",
             "<legacyDrawing", "<picture", "<oleObjects", "<controls",
             "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"]
_AFTER_DV = _AFTER_CF[1:]


def _insert_before_first(xml: str, tags: List[str], block: str) -> str:
    best = None
    for t in tags:
        i = xml.find(t)
        if i >= 0 and (best is None or i < best):
            best = i
    if best is None:
        return xml
    return xml[:best] + block + xml[best:]


def add_conditional_formats(sheet_xml: str, styles: Styles, sqref: str,
                            rules: List[Tuple[str, str]]) -> str:
    """قواعد تلوين حيّة: يتغيّر لون الصف فور تغيير المستخدم للقيمة بيده."""
    if not rules:
        return sheet_xml
    used = [int(p) for p in re.findall(r'<cfRule\b[^>]*priority="(\d+)"', sheet_xml)]
    pri = max(used + [0]) + 1
    body = []
    for formula, hex_rgb in rules:
        body.append(
            '<cfRule type="expression" dxfId="%d" priority="%d" stopIfTrue="1">'
            "<formula>%s</formula></cfRule>"
            % (styles.dxf_fill(hex_rgb), pri, esc(formula)))
        pri += 1
    block = '<conditionalFormatting sqref="%s">%s</conditionalFormatting>' % (
        sqref, "".join(body))
    return _insert_before_first(sheet_xml, _AFTER_CF, block)


def add_list_validation(sheet_xml: str, sqref: str, options: List[str]) -> str:
    """قائمة منسدلة على عمود الحالة — الإدخال اليدوي باختيار لا بكتابة."""
    dv = ('<dataValidation type="list" allowBlank="1" showInputMessage="1"'
          ' showErrorMessage="1" sqref="%s"><formula1>"%s"</formula1>'
          "</dataValidation>" % (sqref, esc(",".join(options))))
    m = re.search(r'<dataValidations\b[^>]*>', sheet_xml)
    if m:
        n = len(re.findall(r"<dataValidation\b", sheet_xml)) + 1
        head = _set_attr(m.group(0), "count", str(n))
        sheet_xml = sheet_xml[:m.start()] + head + sheet_xml[m.end():]
        return sheet_xml.replace("</dataValidations>", dv + "</dataValidations>", 1)
    block = '<dataValidations count="1">%s</dataValidations>' % dv
    return _insert_before_first(sheet_xml, _AFTER_DV, block)


def patch_sheet(sheet_xml: str, styles: Styles, *, header_row: int, n_cols: int,
                cell_values: Dict[Tuple[int, int], Any],
                new_columns: List[Tuple[str, Dict[int, Any]]],
                row_fills: Optional[Dict[int, str]] = None,
                cell_fills: Optional[Dict[Tuple[int, int], str]] = None) -> str:
    """يكتب قيمًا ويضيف أعمدة في آخر الورقة، ويلوّن صفوفًا أو خلايا بعينها.

    cell_fills يلوّن خلية واحدة لا الصف كله: الصف قد يحمل لونًا وضعه
    المستخدم بيده وهو عنده بيانات (أداة الفرز تقرأه)، فلا يجوز أن يُطمس.

    التلوين هنا ثابت لأنه يخصّ صفوفًا محدّدة بأعيانها (ما طابق ملفات التنفيذ)
    لا شرطًا في البيانات؛ أما التلوين حسب الحالة فتنسيق شرطي حيّ."""
    row_fills = row_fills or {}
    # فهرسة بالصف: مسح روابط الصور قد يمسّ آلاف الخلايا، والمرور على القاموس
    # كله مع كل صف يصير بطيئًا في ملف من ستمئة صف
    values_by_row: Dict[int, List[Tuple[int, Any]]] = {}
    for (vr, vc), value in cell_values.items():
        values_by_row.setdefault(vr, []).append((vc, value))
    fills_by_row: Dict[int, List[Tuple[int, str]]] = {}
    for (fr, fc), fhex in (cell_fills or {}).items():
        fills_by_row.setdefault(fr, []).append((fc, fhex))
    n_new = len(new_columns)
    total_cols = n_cols + n_new

    m = re.search(r"<sheetData\b[^>]*?/>|<sheetData\b[^>]*?>", sheet_xml)
    if not m:
        return sheet_xml
    if m.group(0).endswith("/>"):
        return sheet_xml
    data_start = m.end()
    data_end = sheet_xml.index("</sheetData>", data_start)
    head, data, tail = (sheet_xml[:data_start], sheet_xml[data_start:data_end],
                        sheet_xml[data_end:])

    # نمط خلايا الترويسة الجديدة يُنسخ من آخر خلية في ترويسة المستخدم نفسها
    header_style: Optional[int] = None
    for mm in _ROW_RE.finditer(data):
        if _row_number(mm.group(0), 0) == header_row:
            cells = _cells_by_col(_split_row(mm.group(0))[1])
            if cells:
                header_style = _cell_style(cells[max(cells)])
            break

    out: List[str] = []
    pos = 0
    auto_row = 0
    for mm in _ROW_RE.finditer(data):
        out.append(data[pos:mm.start()])
        pos = mm.end()
        row_xml = mm.group(0)
        auto_row += 1
        rn = _row_number(row_xml, auto_row)
        auto_row = rn

        open_tag, inner, close_tag = _split_row(row_xml)
        cells = _cells_by_col(inner)
        touched = False

        for vc, value in values_by_row.get(rn, ()):
            base = _cell_style(cells[vc]) if vc in cells else None
            cells[vc] = _text_cell("%s%d" % (col_letter(vc), rn), base, value)
            touched = True

        if rn == header_row and n_new:
            for j, (title, _vals) in enumerate(new_columns):
                c = n_cols + j
                cells[c] = _text_cell("%s%d" % (col_letter(c), rn), header_style, title)
            touched = True
        elif rn > header_row and n_new:
            for j, (_title, vals) in enumerate(new_columns):
                if rn not in vals:
                    continue
                c = n_cols + j
                base = _cell_style(cells[c]) if c in cells else None
                cells[c] = _text_cell("%s%d" % (col_letter(c), rn), base, vals[rn])
                touched = True

        fill = row_fills.get(rn)
        if fill and rn > header_row:
            for c in range(total_cols):
                base = _cell_style(cells[c]) if c in cells else 0
                new_s = styles.tinted(base, fill)
                if c in cells:
                    cut = cells[c].index(">") + 1      # يشمل علامة الإغلاق
                    cells[c] = (_set_attr(cells[c][:cut], "s", str(new_s))
                                + cells[c][cut:])
                else:
                    cells[c] = '<c r="%s%d" s="%d"/>' % (col_letter(c), rn, new_s)
            touched = True

        for fc, fhex in fills_by_row.get(rn, ()):
            if rn <= header_row:
                continue
            base = _cell_style(cells[fc]) if fc in cells else 0
            new_s = styles.tinted(base, fhex)
            if fc in cells:
                cut = cells[fc].index(">") + 1
                cells[fc] = (_set_attr(cells[fc][:cut], "s", str(new_s))
                             + cells[fc][cut:])
            else:
                cells[fc] = '<c r="%s%d" s="%d"/>' % (col_letter(fc), rn, new_s)
            touched = True

        if touched:
            open_tag = _set_attr(open_tag, "spans", "1:%d" % total_cols)
            row_xml = open_tag + "".join(cells[c] for c in sorted(cells)) + close_tag
        out.append(row_xml)
    out.append(data[pos:])
    data = "".join(out)

    last = col_letter(total_cols - 1)
    head = re.sub(r'(<dimension ref="[A-Z]+\d+:)[A-Z]+(\d+"/>)',
                  lambda x: x.group(1) + last + x.group(2), head, count=1)
    tail = re.sub(r'(<autoFilter ref="[A-Z]+\d+:)[A-Z]+(\d+")',
                  lambda x: x.group(1) + last + x.group(2), tail, count=1)
    head = re.sub(r'(<autoFilter ref="[A-Z]+\d+:)[A-Z]+(\d+")',
                  lambda x: x.group(1) + last + x.group(2), head, count=1)

    if n_new:
        widths = "".join('<col min="%d" max="%d" width="16" customWidth="1"/>'
                         % (n_cols + j + 1, n_cols + j + 1) for j in range(n_new))
        if "<cols>" in head:
            head = head.replace("</cols>", widths + "</cols>", 1)
        else:
            head = head[:head.rindex("<sheetData")] + "<cols>" + widths + "</cols>" \
                + head[head.rindex("<sheetData"):]

    return head + data + tail


def _ref_cells(ref: str) -> List[str]:
    """خلايا مرجع مثل «AD9» أو «AD9:AE10»."""
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", ref.replace("$", ""))
    if not m:
        return [ref]
    if not m.group(3):
        return [m.group(1) + m.group(2)]
    c1, c2 = sorted((col_index(m.group(1)), col_index(m.group(3))))
    r1, r2 = sorted((int(m.group(2)), int(m.group(4))))
    return ["%s%d" % (col_letter(c), r) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def drop_hyperlinks(sheet_xml: str, refs: set) -> Tuple[str, List[str]]:
    """يحذف روابط الخلايا المذكورة. يرجع (الورقة، معرّفات علاقاتها المحذوفة).

    الرابط في إكسل مربوط برقم الخلية لا بنصّها: لو مُسح النص وبقي الرابط
    لبقيت الخلية الفارغة تفتح الصورة عند الضغط عليها."""
    if not refs or "<hyperlink" not in sheet_xml:
        return sheet_xml, []
    removed: List[str] = []

    def one(m: "re.Match") -> str:
        tag = m.group(0)
        ref = _get_attr(tag, "ref") or ""
        if ref and all(c in refs for c in _ref_cells(ref)):
            rid = _get_attr(tag, "r:id")
            if rid:
                removed.append(rid)
            return ""
        return tag

    sheet_xml = re.sub(r"<hyperlink\b[^>]*?(?:/>|>.*?</hyperlink>)", one, sheet_xml, flags=re.S)
    # حاوية فارغة غير مسموحة في مواصفة الملف، فتُحذف معها
    sheet_xml = re.sub(r"<hyperlinks>\s*</hyperlinks>|<hyperlinks/>", "", sheet_xml)
    return sheet_xml, removed


def move_hyperlinks(sheet_xml: str, moves: Dict[str, str]) -> str:
    """ينقل روابط خلايا إلى خلايا أخرى ({"AD9": "AF9"}).

    حين تُبادَل قيمتا خليتين يجب أن يتبع كلَّ رابط نصُّه، وإلا صار نص
    «صورة قبل 1» صورةً وضغطه يفتح صورة أخرى."""
    if not moves or "<hyperlink" not in sheet_xml:
        return sheet_xml

    def one(m: "re.Match") -> str:
        tag = m.group(0)
        ref = _get_attr(tag, "ref") or ""
        return _set_attr(tag, "ref", moves[ref]) if ref in moves else tag

    return re.sub(r"<hyperlink\b[^>]*?/>", one, sheet_xml)


def _drop_rels(rels_xml: str, rids: List[str]) -> str:
    want = set(rids)
    return re.sub(r"<Relationship\b[^>]*/>",
                  lambda m: "" if _get_attr(m.group(0), "Id") in want else m.group(0),
                  rels_xml)


# ------------------------------------------------------------------ الأرشيف

def _rels_target(rels_xml: str, rid: str) -> Optional[str]:
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels_xml):
        tag = m.group(0)
        if _get_attr(tag, "Id") == rid:
            return _get_attr(tag, "Target")
    return None


def locate_sheet(zf: zipfile.ZipFile, wanted: str) -> Tuple[str, str]:
    """يرجّع (اسم الجزء داخل الأرشيف، الاسم المعروض) للورقة المطلوبة."""
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    first = None
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        tag = m.group(0)
        name = _get_attr(tag, "name") or ""
        rid = (_get_attr(tag, "r:id") or _get_attr(tag, "id") or "")
        target = _rels_target(rels, rid) or ""
        if target.startswith("/"):
            part = target[1:]
        else:
            part = "xl/" + target.lstrip("./")
        if first is None:
            first = (part, name)
        if name == wanted:
            return part, name
    if first is None:
        raise ValueError("لم يُعثر على أي ورقة في الملف.")
    return first


def _next_sheet_part(names: set) -> str:
    i = 1
    while "xl/worksheets/sheet%d.xml" % i in names:
        i += 1
    return "xl/worksheets/sheet%d.xml" % i


def _next_rid(rels_xml: str) -> str:
    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)}
    i = 1
    while i in used:
        i += 1
    return "rId%d" % i


def write_patched(src_path: str, dst_path: str, *, sheet_name: str,
                  header_row: int, n_cols: int,
                  cell_values: Dict[Tuple[int, int], Any],
                  new_columns: List[Tuple[str, Dict[int, Any]]],
                  row_fills: Optional[Dict[int, str]] = None,
                  cell_fills: Optional[Dict[Tuple[int, int], str]] = None,
                  cf: Optional[Tuple[str, List[Tuple[str, str]]]] = None,
                  validation: Optional[Tuple[str, List[str]]] = None,
                  sheets_factory=None,
                  drop_links: Optional[set] = None,
                  move_links: Optional[Dict[str, str]] = None) -> None:
    """ينسخ الملف كما هو مع تعديل ورقة واحدة وإضافة أوراق جديدة.

    drop_links: مراجع خلايا («AD9») تُحذف روابطها مع مسح قيمها.
    move_links: روابط تنتقل مع قيمها حين تُبادَل خليتان."""
    with zipfile.ZipFile(src_path) as zin:
        names = set(zin.namelist())
        part, _display = locate_sheet(zin, sheet_name)
        sheet_xml = zin.read(part).decode("utf-8")
        styles_xml = (zin.read("xl/styles.xml").decode("utf-8")
                      if "xl/styles.xml" in names else _MIN_STYLES)
        wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
        rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        ct_xml = zin.read("[Content_Types].xml").decode("utf-8")

        styles = Styles(styles_xml)
        # الأوراق الجديدة تُبنى أولًا لأنها تسجّل أنماطها في نفس جدول الأنماط
        extra_sheets = sheets_factory(styles) if sheets_factory else []

        sheet_xml = patch_sheet(
            sheet_xml, styles, header_row=header_row, n_cols=n_cols,
            cell_values=cell_values, new_columns=new_columns, row_fills=row_fills,
            cell_fills=cell_fills)
        sheet_rels_part = "%s/_rels/%s.rels" % tuple(part.rsplit("/", 1))
        sheet_rels = None
        if move_links:
            sheet_xml = move_hyperlinks(sheet_xml, move_links)
        if drop_links:
            sheet_xml, gone = drop_hyperlinks(sheet_xml, drop_links)
            if gone and sheet_rels_part in names:
                sheet_rels = _drop_rels(zin.read(sheet_rels_part).decode("utf-8"), gone)
        if validation:
            sheet_xml = add_list_validation(sheet_xml, validation[0], validation[1])
        if cf:
            sheet_xml = add_conditional_formats(sheet_xml, styles, cf[0], cf[1])

        # الأوراق الجديدة مليئة بالصيغ: نطلب من إكسل حسبة كاملة عند الفتح
        if "<calcPr" in wb_xml:
            m = re.search(r"<calcPr\b[^>]*/>", wb_xml)
            if m:
                wb_xml = wb_xml.replace(
                    m.group(0), _set_attr(m.group(0), "fullCalcOnLoad", "1"), 1)
        else:
            wb_xml = wb_xml.replace(
                "</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>', 1)

        # ورقة واحدة محدَّدة فقط: تعدّد التحديد يجمع الأوراق فيطبعها إكسل كلها
        if any(b.selected for _t, b in extra_sheets):
            sheet_xml = re.sub(r'\s+tabSelected="1"', "", sheet_xml, count=1)
            n_before = len(re.findall(r"<sheet\b[^>]*?/>", wb_xml))
            m = re.search(r"<workbookView\b[^>]*?/>", wb_xml)
            if m:
                wb_xml = wb_xml.replace(
                    m.group(0), _set_attr(m.group(0), "activeTab", str(n_before)), 1)

        new_parts: Dict[str, str] = {}
        taken = set(names)
        sheet_ids = [int(i) for i in re.findall(r'<sheet\b[^>]*sheetId="(\d+)"', wb_xml)]
        next_id = max(sheet_ids + [0]) + 1
        for title, builder in extra_sheets:
            p = _next_sheet_part(taken)
            taken.add(p)
            new_parts[p] = builder.to_xml()
            rid = _next_rid(rels_xml)
            rels_xml = rels_xml.replace(
                "</Relationships>",
                '<Relationship Id="%s" Type="%s/worksheet" Target="%s"/>'
                "</Relationships>" % (rid, REL_NS, p[len("xl/"):]))
            # نُعلن xmlns:r على العنصر نفسه: بعض الملفات لا تعلنه على <workbook>
            wb_xml = wb_xml.replace(
                "</sheets>",
                '<sheet xmlns:r="%s" name="%s" sheetId="%d" r:id="%s"/></sheets>'
                % (REL_NS, esc(title)[:31], next_id, rid))
            next_id += 1
            ct_xml = ct_xml.replace(
                "</Types>",
                '<Override PartName="/%s" ContentType="%s"/></Types>' % (p, WS_TYPE))

        replaced = {
            part: sheet_xml,
            "xl/styles.xml": styles.render(),
            "xl/workbook.xml": wb_xml,
            "xl/_rels/workbook.xml.rels": rels_xml,
            "[Content_Types].xml": ct_xml,
        }
        if sheet_rels is not None:
            replaced[sheet_rels_part] = sheet_rels

        with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename in replaced:
                    zout.writestr(info.filename, replaced[info.filename].encode("utf-8"))
                elif info.filename == "xl/calcChain.xml":
                    continue        # تعديل الخلايا يُبطله، وإكسل يعيد بناءه وحده
                else:
                    zout.writestr(info, zin.read(info.filename))
            if "xl/styles.xml" not in names:
                zout.writestr("xl/styles.xml", styles.render().encode("utf-8"))
            for p, xml in new_parts.items():
                zout.writestr(p, xml.encode("utf-8"))


_MIN_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="%s">'
    '<fonts count="1"><font><sz val="11"/><color theme="1"/>'
    '<name val="Calibri"/><family val="2"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>" % MAIN_NS)


# ------------------------------------------------------------ مصنّف جديد

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def _txpr(size: int = 1000, bold: bool = False, color: str = "404040") -> str:
    return ('<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="%d" b="%d">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:latin typeface="Arial"/>'
            '<a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="ar-SA"/></a:p>'
            '</c:txPr>' % (size, 1 if bold else 0, color))


def stacked_bar_chart(title: str, sheet: str, categories: List[str], cat_ref: str,
                      series: List[Dict[str, Any]]) -> str:
    """رسم أشرطة أفقية متراكبة — بالقيم المحسوبة مخزّنة فيه كي يظهر فورًا.

    series: [{"name", "name_ref", "ref", "values", "color"}]، والمراجع بصيغة
    $F$6:$F$10 دون اسم الورقة. اسم الورقة يُكتب بين علامتي اقتباس دائمًا
    لأن الاسم العربي قد يحوي مسافة أو شرطة."""
    q = "'%s'" % sheet.replace("'", "''")

    def str_cache(items: List[str]) -> str:
        return '<c:strCache><c:ptCount val="%d"/>%s</c:strCache>' % (
            len(items), "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, esc(v))
                                for i, v in enumerate(items)))

    def num_cache(vals: List[float]) -> str:
        return ('<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>%s'
                '</c:numCache>' % (len(vals), "".join(
                    '<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v) for i, v in enumerate(vals))))

    sers = []
    for i, se in enumerate(series):
        sers.append(
            '<c:ser><c:idx val="%d"/><c:order val="%d"/>'
            '<c:tx><c:strRef><c:f>%s!%s</c:f>%s</c:strRef></c:tx>'
            '<c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>'
            '<c:invertIfNegative val="0"/>'
            '<c:cat><c:strRef><c:f>%s!%s</c:f>%s</c:strRef></c:cat>'
            '<c:val><c:numRef><c:f>%s!%s</c:f>%s</c:numRef></c:val></c:ser>' % (
                i, i, q, se["name_ref"], str_cache([se["name"]]), se["color"],
                q, cat_ref, str_cache(categories), q, se["ref"], num_cache(se["values"])))
    grid = '<c:spPr><a:ln w="9360"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
    axis_ln = '<c:spPr><a:ln w="9360"><a:solidFill><a:srgbClr val="878787"/></a:solidFill></a:ln></c:spPr>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<c:chartSpace xmlns:c="%s" xmlns:a="%s" xmlns:r="%s">'
        '<c:date1904 val="0"/><c:lang val="ar-SA"/><c:roundedCorners val="0"/>'
        '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
        '<a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>'
        '<a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:defRPr></a:pPr>'
        '<a:r><a:rPr lang="ar-SA" sz="1600" b="1"><a:solidFill><a:srgbClr val="1F3864"/>'
        '</a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr>'
        '<a:t>%s</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>'
        '<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
        '%s<c:gapWidth val="80"/><c:overlap val="100"/>'
        '<c:axId val="31236366"/><c:axId val="63572213"/></c:barChart>'
        '<c:catAx><c:axId val="31236366"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="1"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '%s%s<c:crossAx val="63572213"/><c:crosses val="autoZero"/><c:auto val="1"/>'
        '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
        '<c:valAx><c:axId val="63572213"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/><c:majorGridlines>%s</c:majorGridlines>'
        '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/>'
        '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>%s%s'
        '<c:crossAx val="31236366"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>'
        '</c:valAx><c:spPr><a:noFill/></c:spPr></c:plotArea>'
        '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>%s</c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
        '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln w="9360"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
        '</c:chartSpace>' % (
            C_NS, A_NS, REL_NS, esc(title), "".join(sers),
            axis_ln, _txpr(1000, True, "1F3864"), grid, axis_ln, _txpr(900),
            _txpr(1000)))


def _drawing_xml(anchors: List[Tuple[int, int, int, int]]) -> str:
    """مرسى لكل رسم: (عمود البداية، صفها، عمود النهاية، صفها) بترقيم يبدأ من 0."""
    parts = []
    for i, (c1, r1, c2, r2) in enumerate(anchors):
        parts.append(
            '<xdr:twoCellAnchor editAs="oneCell">'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row>'
            '<xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row>'
            '<xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="Chart %d"/>'
            '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/>'
            '<a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>'
            '<a:graphicData uri="%s"><c:chart xmlns:c="%s" r:id="rId%d"/></a:graphicData>'
            '</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
            % (c1, r1, c2, r2, i + 2, i + 1, C_NS, C_NS, i + 1))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<xdr:wsDr xmlns:xdr="%s" xmlns:a="%s" xmlns:r="%s">%s</xdr:wsDr>'
            % (XDR_NS, A_NS, REL_NS, "".join(parts)))


def write_workbook(dst_path: str, sheets: List[Tuple[str, "SheetBuilder"]],
                   styles: "Styles",
                   charts: Optional[Dict[int, List[Tuple[str, Tuple[int, int, int, int]]]]] = None
                   ) -> None:
    """يكتب ملف إكسل جديدًا من أوراق مبنية بـ SheetBuilder.

    نكتب الحزمة بأيدينا بدل openpyxl لسبب واحد: الصيغ هنا تُكتب ومعها
    قيمتها المحسوبة. ملف openpyxl يصل بصيغ بلا قيم، فيفتحه عارض الجوال
    (معاينة واتساب مثلًا) بخلايا فارغة لأنه لا يحسب — والملخص يُرسل غالبًا
    بهذه الطريقة. و fullCalcOnLoad يجعل إكسل يعيد الحساب عند الفتح."""
    charts = charts or {}
    ct_sheets = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
        for i in range(len(sheets)))
    # ترقيم الرسومات عبر الأوراق: drawingN لكل ورقة فيها رسم، و chartN لكل رسم
    drawing_of: Dict[int, int] = {}
    chart_parts: List[Tuple[int, int, str]] = []        # (ورقة، رقم الرسم، XML)
    for si in sorted(charts):
        if not charts[si]:
            continue
        drawing_of[si] = len(drawing_of) + 1
        for xml, _anchor in charts[si]:
            chart_parts.append((si, len(chart_parts) + 1, xml))
    for si, dn in drawing_of.items():
        ct_sheets += ('<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/'
                      'vnd.openxmlformats-officedocument.drawing+xml"/>' % dn)
        sheets[si][1].drawing_rid = "rId1"
    for _si, cn, _x in chart_parts:
        ct_sheets += ('<Override PartName="/xl/charts/chart%d.xml" ContentType="application/'
                      'vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' % cn)
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '%s</Types>' % ct_sheets)
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb_sheets = "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (esc(n), i + 1, i + 1)
                        for i, (n, _sb) in enumerate(sheets))
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="%s" xmlns:r="%s"><bookViews><workbookView activeTab="0"/>'
        '</bookViews><sheets>%s</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/>'
        '</workbook>' % (MAIN_NS, REL_NS, wb_sheets))
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '%s<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships/styles" Target="styles.xml"/></Relationships>' % (
            "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/worksheet" Target="worksheets/'
                    'sheet%d.xml"/>' % (i + 1, i + 1) for i in range(len(sheets))),
            len(sheets) + 1))
    with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles.render())
        for i, (_n, sb) in enumerate(sheets):
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1), sb.to_xml())
        rel_t = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        rels_head = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        for si, dn in drawing_of.items():
            mine = [(cn, x) for s_i, cn, x in chart_parts if s_i == si]
            z.writestr("xl/worksheets/_rels/sheet%d.xml.rels" % (si + 1),
                       rels_head + '<Relationship Id="rId1" Type="%sdrawing" '
                       'Target="../drawings/drawing%d.xml"/></Relationships>' % (rel_t, dn))
            z.writestr("xl/drawings/drawing%d.xml" % dn,
                       _drawing_xml([a for _x, a in charts[si]]))
            z.writestr("xl/drawings/_rels/drawing%d.xml.rels" % dn, rels_head + "".join(
                '<Relationship Id="rId%d" Type="%schart" Target="../charts/chart%d.xml"/>'
                % (k + 1, rel_t, cn) for k, (cn, _x) in enumerate(mine)) + "</Relationships>")
            for cn, x in mine:
                z.writestr("xl/charts/chart%d.xml" % cn, x)
