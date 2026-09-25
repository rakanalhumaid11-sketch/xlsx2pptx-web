"""فرز ملف الملاحظات الكبير: تصفية وترتيب دون المساس بالبيانات.

المشكلة التي تحلّها هذه الوحدة: ملف الجودة يصدُر بآلاف الصفوف تغطي كل
المغذيات، والعمل اليومي لا يكون إلا على بعضها. والأهم أن المستخدم يلوّن
الصفوف بيده — أحمر وأزرق وأصفر وأخضر — فيصير اللون حالةً لا يحملها أي
عمود في الملف. فالفرز هنا يقرأ اللون بوصفه بيانات، ويرتّب عليه.

القاعدة الحاكمة: لا تتغيّر خلية واحدة. تُحذف صفوف وتتحرّك صفوف، وما عدا
ذلك — القيم والأنماط والألوان وعروض الأعمدة وتجميد الترويسة والتصفية
التلقائية وروابط الصور — ينتقل كما هو. ولذلك تعمل الوحدة على XML الورقة
مباشرة بدل إعادة بنائها بمكتبة، تمامًا كما يفعل xlsxedit.py.
"""
from __future__ import annotations

import colorsys
import html
import os
import re
import shutil
import zipfile
from typing import Any, Dict, List, Optional, Tuple

# ----------------------------------------------------------------- الألوان

# ترتيب ألوان القالب كما يفهمه إكسل في السمة theme="n": أول زوجين مقلوبان
# عن ترتيبهما في ملف السمة (lt1 قبل dk1)، وهذا مصدر أخطاء شائع.
THEME_ORDER = ["lt1", "dk1", "lt2", "dk2",
               "accent1", "accent2", "accent3", "accent4", "accent5",
               "accent6", "hlink", "folHlink"]

# الترتيب الثابت الذي يطلبه المستخدم. المفتاح الداخلي ثم الاسم المعروض.
COLOR_ORDER: List[Tuple[str, str]] = [
    ("red", "أحمر"),
    ("blue", "أزرق"),
    ("yellow", "أصفر"),
    ("green", "أخضر"),
    ("other", "ألوان أخرى"),
    ("none", "بلا لون"),
]
COLOR_RANK = {k: i for i, (k, _) in enumerate(COLOR_ORDER)}
COLOR_LABEL = dict(COLOR_ORDER)

# اللون المعروض في الصفحة لكل مفتاح حين لا نعرف درجته الحقيقية
COLOR_SWATCH = {"red": "FF0000", "blue": "8EA9DB", "yellow": "FFFF00",
                "green": "00B050", "other": "C9C9C9", "none": "FFFFFF"}


def _srgb(hexs: str) -> Tuple[float, float, float]:
    hexs = hexs[-6:]
    return tuple(int(hexs[i:i + 2], 16) / 255.0 for i in (0, 2, 4))  # type: ignore


def apply_tint(hexs: str, tint: float) -> str:
    """درجة التفتيح/التغميق التي يخزّنها إكسل مع لون السمة.

    بدونها يخرج «الأزرق الفاتح» أزرق داكنًا فيُصنَّف خطأً."""
    if not tint:
        return hexs[-6:].upper()
    r, g, b = _srgb(hexs)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def classify(hexs: str) -> str:
    """اسم اللون بالمعنى الذي يقصده المستخدم، لا بدرجته الدقيقة.

    نصنّف بالصبغة لا بالمطابقة التامة، حتى لا تنهار الأداة إذا لوّن
    المستخدم بدرجة أفتح أو أغمق قليلًا في الجولة القادمة."""
    if not hexs:
        return "none"
    r, g, b = _srgb(hexs)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    hue = h * 360.0
    if l >= 0.95 and s < 0.25:      # أبيض أو شبه أبيض = بلا لون فعليًا
        return "none"
    if s < 0.12:                    # رمادي
        return "other"
    if hue < 18 or hue >= 336:
        return "red"
    if hue < 48:                    # برتقالي يُقرأ أحمر في الميدان
        return "red"
    if hue < 72:
        return "yellow"
    if hue < 170:
        return "green"
    if hue < 265:
        return "blue"
    return "other"


# ------------------------------------------------------- قراءة XML الورقة

_ROW_RE = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row\s*>)", re.S)
_CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c\s*>)", re.S)
_ATTR_RE = re.compile(r'(\w[\w:]*)\s*=\s*"([^"]*)"')
_V_RE = re.compile(r"<v[^>]*>(.*?)</v\s*>", re.S)
_IS_T_RE = re.compile(r"<t[^>]*>(.*?)</t\s*>", re.S)


def _attrs(s: str) -> Dict[str, str]:
    return dict(_ATTR_RE.findall(s))


def col_letters(ref: str) -> str:
    return "".join(ch for ch in ref if ch.isalpha())


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _shared_strings(z: zipfile.ZipFile) -> List[str]:
    try:
        raw = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for si in re.findall(r"<si\b[^>]*>(.*?)</si\s*>", raw, re.S):
        out.append(html.unescape("".join(_IS_T_RE.findall(si))))
    return out


def _fill_of_xf(styles: str) -> Tuple[List[int], List[str]]:
    """(xf -> fillId، fillId -> لون سداسي) من styles.xml."""
    fills_xml = re.search(r"<fills\b[^>]*>(.*?)</fills\s*>", styles, re.S)
    fills: List[str] = []
    if fills_xml:
        for f in re.findall(r"<fill\b[^>]*?(?:/>|>(.*?)</fill\s*>)", fills_xml.group(1), re.S):
            fills.append(f or "")
    xf_fill: List[int] = []
    cx = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs\s*>", styles, re.S)
    if cx:
        for x in re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf\s*>)", cx.group(1), re.S):
            m = re.search(r'fillId="(\d+)"', x)
            xf_fill.append(int(m.group(1)) if m else 0)
    return xf_fill, fills


def _theme_colors(z: zipfile.ZipFile) -> Dict[str, str]:
    try:
        raw = z.read("xl/theme/theme1.xml").decode("utf-8")
    except KeyError:
        return {}
    cs = re.search(r"<a:clrScheme\b.*?</a:clrScheme\s*>", raw, re.S)
    if not cs:
        return {}
    out: Dict[str, str] = {}
    for name, body in re.findall(r"<a:(\w+)>(.*?)</a:\1>", cs.group(0), re.S):
        m = re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', body)
        if m:
            out[name] = m.group(1).upper()
            continue
        m = re.search(r'sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"', body)
        if m:
            out[name] = m.group(1).upper()
    return out


def _fill_color(fill_xml: str, theme: Dict[str, str]) -> str:
    """لون التعبئة السداسي، أو "" إذا كانت الخلية بلا تعبئة."""
    if 'patternType="solid"' not in fill_xml:
        return ""
    m = re.search(r"<fgColor\b([^>]*)/?>", fill_xml)
    if not m:
        return ""
    a = _attrs(m.group(1))
    tint = float(a.get("tint", "0") or 0)
    if "rgb" in a:
        return apply_tint(a["rgb"], tint)
    if "theme" in a:
        try:
            name = THEME_ORDER[int(a["theme"])]
        except (ValueError, IndexError):
            return ""
        base = theme.get(name, "")
        return apply_tint(base, tint) if base else ""
    return ""


# --------------------------------------------------------------- التحليل

FEEDER_ALIASES = ["المغذي", "المغذى", "اسم المغذي"]
TYPE_ALIASES = ["الملاحظة", "الملاحظه", "وصف الملاحظة"]
D1_ALIASES = ["إشعار D1", "اشعار D1", "إشعار d1", "اشعار d1", "D1"]

MAX_TYPES = 400          # سقف يمنع صفحة بآلاف المرابع لو كان العمود نصًّا حرًّا


def _norm(s: str) -> str:
    """توحيد بسيط لمقارنة عناوين الأعمدة: الهمزات والتاء المربوطة والمسافات."""
    s = (s or "").strip().lower()
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ة", "ه").replace("ى", "ي").replace("ـ", "")
    return re.sub(r"\s+", " ", s)


def _find_col(headers: Dict[str, str], aliases: List[str]) -> Optional[str]:
    """حرف العمود المطابق لأحد الأسماء؛ المطابقة التامة تغلب البادئة."""
    want = [_norm(a) for a in aliases]
    prefix = None
    for letter, text in headers.items():
        nh = _norm(text)
        if nh in want:
            return letter
        if prefix is None and any(w and nh.startswith(w) for w in want):
            prefix = letter
    return prefix


def natural_key(s: str) -> List[Tuple[int, Any]]:
    """ترتيب يفهم الأرقام داخل النص: 211 قبل 2597 لا بعده."""
    parts = re.findall(r"\d+|\D+", s or "")
    return [(0, int(p)) if p.isdigit() else (1, p) for p in parts]


def _sheet_path(z: zipfile.ZipFile) -> Tuple[str, str]:
    """(مسار ورقة العمل الأولى داخل الأرشيف، اسمها المعروض)."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    sheets = re.findall(r"<sheet\b([^>]*)/?>", wb)
    if not sheets:
        raise ValueError("الملف لا يحتوي على أوراق عمل.")
    a = _attrs(sheets[0])
    rid = a.get("r:id") or a.get("id") or ""
    name = html.unescape(a.get("name", "ورقة1"))
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = None
    for r in re.findall(r"<Relationship\b([^>]*)/?>", rels):
        ra = _attrs(r)
        if ra.get("Id") == rid:
            target = ra.get("Target", "")
            break
    if not target:
        target = "worksheets/sheet1.xml"
    target = target.lstrip("/")
    path = target if target.startswith("xl/") else "xl/" + target
    if path not in z.namelist():
        cand = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
        if not cand:
            raise ValueError("تعذّر العثور على ورقة العمل داخل الملف.")
        path = sorted(cand)[0]
    return path, name


def analyze(path: str) -> Dict[str, Any]:
    """كل ما تحتاجه صفحة الاختيار، بمرور واحد على الملف."""
    with zipfile.ZipFile(path) as z:
        sheet_path, sheet_name = _sheet_path(z)
        sst = _shared_strings(z)
        theme = _theme_colors(z)
        xf_fill, fills = _fill_of_xf(z.read("xl/styles.xml").decode("utf-8"))
        sheet = z.read(sheet_path).decode("utf-8")

    # لون كل نمط خلية، محسوبًا مرة واحدة
    xf_color: List[str] = []
    for fid in xf_fill:
        xf_color.append(_fill_color(fills[fid], theme) if fid < len(fills) else "")

    def value(cell_attrs: Dict[str, str], body: str) -> str:
        t = cell_attrs.get("t", "")
        if t == "inlineStr":
            return html.unescape("".join(_IS_T_RE.findall(body or "")))
        m = _V_RE.search(body or "")
        if not m:
            return ""
        v = html.unescape(m.group(1))
        if t == "s":
            try:
                return sst[int(v)]
            except (ValueError, IndexError):
                return ""
        return v

    rows_raw: List[Tuple[int, str]] = []
    for attrs_s, body in _ROW_RE.findall(sheet):
        a = _attrs(attrs_s)
        if "r" not in a:
            continue
        rows_raw.append((int(a["r"]), body or ""))
    rows_raw.sort(key=lambda x: x[0])
    if not rows_raw:
        raise ValueError("الورقة فارغة.")

    header_row = rows_raw[0][0]
    headers: Dict[str, str] = {}
    for ca, cb in _CELL_RE.findall(rows_raw[0][1]):
        a = _attrs(ca)
        if "r" in a:
            headers[col_letters(a["r"])] = value(a, cb)

    c_feeder = _find_col(headers, FEEDER_ALIASES)
    c_type = _find_col(headers, TYPE_ALIASES)
    c_d1 = _find_col(headers, D1_ALIASES)

    recs: List[Dict[str, Any]] = []
    for rn, body in rows_raw[1:]:
        cells: Dict[str, Tuple[Dict[str, str], str]] = {}
        colors: Dict[str, int] = {}
        any_value = False
        for ca, cb in _CELL_RE.findall(body):
            a = _attrs(ca)
            ref = a.get("r", "")
            if not ref:
                continue
            letter = col_letters(ref)
            cells[letter] = (a, cb)
            s = a.get("s")
            if s is not None:
                try:
                    c = xf_color[int(s)]
                except (ValueError, IndexError):
                    c = ""
                if c:
                    colors[c] = colors.get(c, 0) + 1
            if (cb or "").strip():
                any_value = True
        if not any_value:
            continue
        hexc = max(colors.items(), key=lambda kv: kv[1])[0] if colors else ""
        key = classify(hexc)
        recs.append({
            "row": rn,
            "feeder": value(*cells[c_feeder]) if c_feeder in cells else "",
            "type": value(*cells[c_type]) if c_type in cells else "",
            "d1": bool(value(*cells[c_d1]).strip()) if c_d1 in cells else False,
            "color": key,
            "hex": hexc,
        })

    def tally(field: str) -> List[Tuple[str, int]]:
        c: Dict[str, int] = {}
        for r in recs:
            c[r[field]] = c.get(r[field], 0) + 1
        return sorted(c.items(), key=lambda kv: natural_key(kv[0]))

    colors_seen: Dict[str, Dict[str, Any]] = {}
    for r in recs:
        e = colors_seen.setdefault(r["color"], {"key": r["color"], "count": 0,
                                                "hex": r["hex"] or COLOR_SWATCH[r["color"]]})
        e["count"] += 1
    color_list = []
    for k, _ in COLOR_ORDER:
        if k in colors_seen:
            e = dict(colors_seen[k])
            e["label"] = COLOR_LABEL[k]
            e["rank"] = COLOR_RANK[k]
            color_list.append(e)

    types = tally("type")
    # مكعّب عدّ صغير (مغذي × لون × نوع) ترسله الصفحة إلى المتصفح، فيحسب
    # المتبقي لحظيًا مع كل تأشيرة دون أن نرسل ستة آلاف صف
    f_idx = {f: i for i, (f, _) in enumerate(tally("feeder"))}
    t_idx = {t: i for i, (t, _) in enumerate(types)}
    c_idx = {k: i for i, (k, _) in enumerate(COLOR_ORDER)}
    cube: Dict[str, int] = {}
    for r in recs:
        k = "%d,%d,%d" % (f_idx.get(r["feeder"], -1), c_idx.get(r["color"], -1),
                          t_idx.get(r["type"], -1))
        cube[k] = cube.get(k, 0) + 1

    return {
        "sheet": sheet_name,
        "sheet_path": sheet_path,
        "header_row": header_row,
        "total": len(recs),
        "records": recs,
        "feeders": tally("feeder"),
        "colors": color_list,
        "types": types if len(types) <= MAX_TYPES else [],
        "types_too_many": len(types) > MAX_TYPES,
        "cube": cube,
        "n_d1": sum(1 for r in recs if r["d1"]),
        "has_feeder": c_feeder is not None,
        "has_type": c_type is not None,
        "has_d1": c_d1 is not None,
    }


def plan(an: Dict[str, Any], feeders: List[str], colors: List[str],
         types: List[str]) -> List[int]:
    """أرقام الصفوف الباقية بالترتيب المطلوب.

    الترتيب ثابت لا يُسأل عنه: المغذي باسمه، ثم اللون بتسلسله (أحمر ثم
    أزرق ثم أصفر ثم أخضر)، ثم التي لها إشعار D1 قبل التي بلا إشعار، وما
    تساوى بقي بترتيبه الأصلي في الملف."""
    fset, cset, tset = set(feeders), set(colors), set(types)
    # حين يتعذّر عرض الأنواع (عمود نصّي حرّ تجاوز السقف) نُسقط تصفيتها كلها
    # بدل أن نحذف الملف كله لأن الصفحة لم تعرض مرابعها
    skip_types = (not an["has_type"]) or an.get("types_too_many")
    keep = [r for r in an["records"]
            if (not an["has_feeder"] or r["feeder"] in fset)
            and r["color"] in cset
            and (skip_types or r["type"] in tset)]
    keep.sort(key=lambda r: (natural_key(r["feeder"]),
                             COLOR_RANK.get(r["color"], 99),
                             0 if r["d1"] else 1,
                             r["row"]))
    return [r["row"] for r in keep]


# ---------------------------------------------------------- كتابة الملف

_HL_RE = re.compile(r"<hyperlink\b[^>]*?/>|<hyperlink\b[^>]*?>.*?</hyperlink\s*>", re.S)


def _renumber(row_xml_attrs: str, body: str, new: int) -> str:
    """صفّ كامل بعد تغيير رقمه ورقم كل خلية فيه."""
    attrs = re.sub(r'(\br=")\d+(")', lambda m: m.group(1) + str(new) + m.group(2),
                   row_xml_attrs, count=1)
    # لا نلمس إلا خاصية r في وسم <c>؛ أي نص آخر يبقى حرفًا بحرف
    body = re.sub(r'(<c\b[^>]*?\br=")([A-Z]+)\d+(")',
                  lambda m: m.group(1) + m.group(2) + str(new) + m.group(3), body)
    return "<row%s>%s</row>" % (attrs, body)


def _last_col(dim_ref: str) -> str:
    part = dim_ref.split(":")[-1]
    return col_letters(part) or "A"


def write_sorted(src: str, dst: str, sheet_path: str, header_row: int,
                 order: List[int]) -> Dict[str, Any]:
    """ينسخ الملف كما هو، ولا يعيد بناء إلا ورقة البيانات بصفوفها الجديدة."""
    with zipfile.ZipFile(src) as z:
        sheet = z.read(sheet_path).decode("utf-8")
        names = z.namelist()

        rows: Dict[int, Tuple[str, str]] = {}
        for attrs_s, body in _ROW_RE.findall(sheet):
            a = _attrs(attrs_s)
            if "r" in a:
                rows[int(a["r"])] = (attrs_s, body or "")

        if header_row not in rows:
            raise ValueError("تعذّر العثور على صف الترويسة في الورقة.")

        # دمج خلايا داخل نطاق البيانات يستحيل معه الفرز: الدمج مربوط بأرقام
        # صفوف بعينها، فتحريك الصفوف يمزّقه. نرفض بدل أن نُخرج ملفًا تالفًا.
        merges = re.findall(r'<mergeCell\b[^>]*ref="([^"]+)"', sheet)
        for ref in merges:
            rs = [int(x) for x in re.findall(r"[A-Z]+(\d+)", ref)]
            if rs and max(rs) > header_row:
                raise ValueError(
                    "الورقة فيها خلايا مدمجة داخل البيانات، والفرز يفسدها. "
                    "أزل الدمج من الملف ثم أعد المحاولة.")

        new_of: Dict[int, int] = {}
        out_rows: List[str] = [_renumber(rows[header_row][0], rows[header_row][1], 1)]
        new_of[header_row] = 1
        for i, rn in enumerate(order, start=2):
            if rn not in rows:
                continue
            new_of[rn] = i
            out_rows.append(_renumber(rows[rn][0], rows[rn][1], i))
        last_row = len(out_rows)

        # نطاق الورقة وحدودها: dimension والتصفية التلقائية والنطاق المعرَّف
        dim_m = re.search(r'<dimension\b[^>]*ref="([^"]+)"', sheet)
        last_col = _last_col(dim_m.group(1)) if dim_m else "A"
        new_ref = "A1:%s%d" % (last_col, last_row)

        head, _, rest = sheet.partition("<sheetData")
        if not rest:
            raise ValueError("الورقة بلا بيانات.")
        _, _, tail = rest.partition(">")
        close = tail.find("</sheetData>")
        if close < 0:                      # <sheetData/> فارغة
            close = 0
            tail = tail.lstrip("/").lstrip(">")
            after = tail
        else:
            after = tail[close + len("</sheetData>"):]

        head = re.sub(r'(<dimension\b[^>]*ref=")[^"]+(")',
                      lambda m: m.group(1) + new_ref + m.group(2), head)
        after = re.sub(r'(<autoFilter\b[^>]*ref=")[^"]+(")',
                       lambda m: m.group(1) + new_ref + m.group(2), after)

        # الروابط مربوطة برقم الصف لا بمحتواه: بلا إعادة ربط تهاجر صورة كل
        # ملاحظة إلى ملاحظة غيرها — وهذا أخطر ما في الفرز
        kept_links = 0
        dropped_links = 0

        def fix_link(m: re.Match) -> str:
            nonlocal kept_links, dropped_links
            s = m.group(0)
            ref = re.search(r'ref="([^"]+)"', s)
            if not ref:
                return s
            old = ref.group(1).split(":")[0]
            num = re.search(r"\d+", old)
            if not num:
                return s
            new = new_of.get(int(num.group(0)))
            if new is None:
                dropped_links += 1
                return ""
            kept_links += 1
            return s.replace('ref="%s"' % ref.group(1),
                             'ref="%s%d"' % (col_letters(old), new))

        after = _HL_RE.sub(fix_link, after)
        after = re.sub(r"<hyperlinks\b[^>]*>\s*</hyperlinks\s*>", "", after)

        new_sheet = head + "<sheetData>" + "".join(out_rows) + "</sheetData>" + after

        os.makedirs(os.path.dirname(os.path.abspath(dst)) or ".", exist_ok=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as out:
            for n in names:
                if n == "xl/calcChain.xml":
                    continue           # يفسد بعد تغيّر أرقام الصفوف؛ يعيد إكسل بناءه
                if n == sheet_path:
                    out.writestr(n, new_sheet)
                elif n == "xl/workbook.xml":
                    # النطاق المعرَّف للتصفية يحفظ آخر صف؛ لو بقي على القديم
                    # فتح إكسل الملف بتحذير «محتوى غير قابل للقراءة»
                    wb = z.read(n).decode("utf-8")
                    wb = re.sub(r"(<definedName[^>]*_xlnm\._FilterDatabase[^>]*>)([^<]*)(</definedName>)",
                                lambda m: m.group(1) + re.sub(r"\$[A-Z]+\$\d+$",
                                                              "$%s$%d" % (last_col, last_row),
                                                              m.group(2)) + m.group(3), wb)
                    out.writestr(n, wb)
                elif n == "[Content_Types].xml":
                    ct = z.read(n).decode("utf-8")
                    ct = ct.replace(
                        '<Override PartName="/xl/calcChain.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.'
                        'spreadsheetml.calcChain+xml"/>', "")
                    out.writestr(n, ct)
                else:
                    out.writestr(n, z.read(n))

    return {"rows": last_row - 1, "links_kept": kept_links,
            "links_dropped": dropped_links, "last_ref": new_ref}


def suggested_name(src_name: str) -> str:
    base = os.path.splitext(os.path.basename(src_name))[0]
    return "%s - مفروز.xlsx" % base
