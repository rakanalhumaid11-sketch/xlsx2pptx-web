# -*- coding: utf-8 -*-
"""
approval.py
===========
متابعة حالة الاعتماد من ملف واحد — بلا ملفات ثانوية.

يقرأ عمود الحالة في ملف الملاحظات ويترجم رموزه (APPROVED / ASSIGNED / …)
إلى أربع خانات يفهمها المتابع: تم الاقفال، بانتظار الاقفال، مسترجع،
بانتظار المعالجة. ثم يضيف في آخر الملف عمود «حالة الاعتماد» بقائمة منسدلة
ليعدّله المستخدم بيده، ويلوّن الصفوف تنسيقًا شرطيًا، ويبني داتا شيت حيًّا
حسب المقاول والتصنيف والأولوية.

الملف الأصلي لا يُمسّ: لا تنسيق ولا ترتيب ولا خلية قائمة — فقط عمود جديد في
آخره وأوراق مضافة (انظر xlsxedit).
"""

import os
import re
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

import engine
import matcher
import xlsxedit
from matcher import _clean, find_col, read_table, make_styles, band, _q, COLOR_HEX
from xlsxedit import SheetBuilder, Styles

# ------------------------------------------------------------------ ثوابت

AP_APPROVED = "تم الاقفال"
AP_DONE = "بانتظار الاقفال"
AP_RETURNED = "مسترجع"
AP_PENDING = "بانتظار المعالجة"
AP_STATES = [AP_APPROVED, AP_DONE, AP_RETURNED, AP_PENDING]

# تسميات جولات سابقة — تُترجم للجديدة كي لا تضيع التعديلات اليدوية القديمة
LEGACY_STATES = {
    "موافق عليه": AP_APPROVED,
    "موافق عليها": AP_APPROVED,
    "موافق عليها (تم الاقفال)": AP_APPROVED,
    "منجز بانتظار الموافقة": AP_DONE,
    "تمت المعالجة": AP_DONE,
    "معاد": AP_RETURNED,
}

AP_HEADER = "حالة الاعتماد"
AP_ALIASES = [AP_HEADER, "حاله الاعتماد"]

DEFAULT_COLORS = {AP_APPROVED: "green", AP_DONE: "yellow",
                  AP_RETURNED: "red", AP_PENDING: "none"}

STATUS_ALIASES = ["الحالة", "الحاله", "حالة الملاحظة", "حالة الملاحظه",
                  "الحالة الحالية", "status"]
PRIORITY_ALIASES = ["الأولوية", "الاولوية", "الاولويه", "priority"]

# رموز النظام كما تصدر من التطبيق، ومقابلها بالعربية
CODE_MAP = {
    "APPROVED": AP_APPROVED, "ACCEPTED": AP_APPROVED, "CLOSED": AP_APPROVED,
    "DONE": AP_DONE, "COMPLETED": AP_DONE, "FINISHED": AP_DONE,
    "RETURNED": AP_RETURNED, "REJECTED": AP_RETURNED, "REOPENED": AP_RETURNED,
    "NEW": AP_PENDING, "ASSIGNED": AP_PENDING, "OPEN": AP_PENDING,
    "IN_PROGRESS": AP_PENDING, "WORK_NOT_POSSIBLE": AP_PENDING,
    "CANCELLED": AP_PENDING, "PENDING": AP_PENDING,
}

EXTRA_ROWS = matcher.EXTRA_ROWS

# لوحة الداتا شيت (نفس لوحة أداة المتابعة كي تبدو الأدوات عائلة واحدة)
INK, SOFT, MUTED, GREEN = matcher.INK, matcher.SOFT, matcher.MUTED, matcher.GREEN


def guess_bucket(raw: Any) -> str:
    """يخمّن الخانة من رمز الحالة، إنجليزيًا كان أو عربيًا."""
    t = _clean(raw)
    if not t:
        return AP_PENDING
    # المطابقة التامة أولًا: «بانتظار الاقفال» تحوي كلمة «اقفال» فلو فحصنا
    # الكلمات المفتاحية قبلها لصُنّفت خطأً على أنها مقفلة
    if t in AP_STATES:
        return t
    if t in LEGACY_STATES:
        return LEGACY_STATES[t]
    code = t.strip().upper().replace(" ", "_").replace("-", "_")
    if code in CODE_MAP:
        return CODE_MAP[code]
    ar = engine.normalize_ar(t)
    if any(w in ar for w in ("موافق", "معتمد", "مقبول", "مغلق", "اقفال", "مقفل")):
        return AP_APPROVED
    if any(w in ar for w in ("معاد", "مرفوض", "مرتجع", "مسترجع", "اعاده")):
        return AP_RETURNED
    if any(w in ar for w in ("منجز", "منفذ", "مكتمل", "تمت", "تم ")):
        return AP_DONE
    return AP_PENDING


# ------------------------------------------------------------------ التحليل

def analyze(path: str) -> Dict[str, Any]:
    """يقرأ الملف ويستخرج قيم عمود الحالة وأعمدة التقسيم المقترحة."""
    sheet, header_row, headers, rows, row_nums = read_table(path)
    if not rows:
        raise ValueError("الملف لا يحتوي على صفوف بيانات.")

    status_col = find_col(headers, STATUS_ALIASES)
    prev_col = find_col(headers, AP_ALIASES)

    # قيم الحالة الخام مرتّبة تنازليًا حسب التكرار، مع خانتها المقترحة
    raw_counts: Dict[str, int] = {}
    if status_col is not None:
        for r in rows:
            v = _clean(r[status_col]) if status_col < len(r) else ""
            raw_counts[v or "(فارغة)"] = raw_counts.get(v or "(فارغة)", 0) + 1
    statuses = [{"raw": k, "count": v, "bucket": guess_bucket("" if k == "(فارغة)" else k)}
                for k, v in sorted(raw_counts.items(), key=lambda t: -t[1])]

    # الأعمدة الصالحة للتقسيم: قليلة القيم المتكررة ومملوءة بما يكفي
    groupable = []
    for c, h in enumerate(headers):
        if not h or c == status_col:
            continue
        vals = {_clean(r[c]) for r in rows if c < len(r) and _clean(r[c])}
        if 1 < len(vals) <= 40:
            groupable.append({"i": c, "name": h, "n": len(vals)})

    return {
        "sheet": sheet,
        "header_row": header_row,
        "headers": headers,
        "rows": rows,
        "row_nums": row_nums,
        "status_col": status_col,
        "prev_col": prev_col,
        "statuses": statuses,
        "groupable": groupable,
        "contractor_col": find_col(headers, matcher.CONTRACTOR_ALIASES),
        "class_col": find_col(headers, matcher.CLASS_ALIASES),
        "priority_col": find_col(headers, PRIORITY_ALIASES),
        "feeder": _feeder(headers, rows),
        "total": len(rows),
    }


def _feeder(headers: List[str], rows: List[List[Any]]) -> str:
    """رمز المغذي والمحطة للعنوان، مثل «AH08 — 8911»."""
    f = find_col(headers, ["المغذي", "المغذى"])
    s = find_col(headers, ["المحطة", "المحطه"])
    fv = _clean(rows[0][f]) if f is not None and f < len(rows[0]) else ""
    sv = _clean(rows[0][s]) if s is not None and s < len(rows[0]) else ""
    return " — ".join(x for x in (fv, sv) if x)


def resolve(an: Dict[str, Any], mapping: Dict[str, str]) -> List[str]:
    """خانة كل صف: من عمود «حالة الاعتماد» السابق إن وُجد، وإلا من الخريطة."""
    out: List[str] = []
    for r in an["rows"]:
        prev = (_clean(r[an["prev_col"]])
                if an["prev_col"] is not None and an["prev_col"] < len(r) else "")
        prev = prev if prev in AP_STATES else LEGACY_STATES.get(prev, "")

        raw = (_clean(r[an["status_col"]])
               if an["status_col"] is not None and an["status_col"] < len(r) else "")
        now = mapping.get(raw or "(فارغة)") or guess_bucket(raw)

        # الإدخال اليدوي السابق يُحترم، إلا أن يكون النظام قد حسم الأمر بعده
        # (إقفال أو استرجاع) — وإلا بقي الملف يعرض حالة قديمة بعد التحديث
        out.append(now if (not prev or now in (AP_APPROVED, AP_RETURNED)) else prev)
    return out


# ------------------------------------------------------------------ الداتا شيت

HEAD = ["القيمة", "الإجمالي"] + AP_STATES + ["نسبة الإنجاز"]
WIDTHS = [(0, 2.5), (1, 26), (2, 12), (3, 15), (4, 15), (5, 12), (6, 17), (7, 14), (8, 2.5)]

# كل بطاقة مؤشر تقف فوق العمود الذي تلخّصه في الجداول، فيقرأ العين عموديًا.
# بطاقة الإجمالي وحدها تُمدّ على عمودين لأن أولها (عمود القيمة) لا مؤشر له.
KPI_SPAN = [(1, 2), (3, 3), (4, 4), (5, 5), (6, 6), (7, 7)]
# خلية كل مؤشر: الإجمالي ثم الخانات الأربع بترتيب أعمدة الجدول
KPI_CELL = ["B6", "D6", "E6", "F6", "G6"]


_EPOCH = datetime(1899, 12, 30)
_DATE_TEXT = re.compile(r"(\d{4}-\d{2}-\d{2})[ T]\d{2}:\d{2}:\d{2}(\.\d+)?")


def _excel_date(v: Any) -> Optional[float]:
    """تاريخ إكسل رقمًا — كي يُعرض «2026-09-03» لا «2026-09-03 00:00:00»."""
    if isinstance(v, datetime):
        return (v - _EPOCH).days + (v.hour * 3600 + v.minute * 60 + v.second) / 86400.0
    if isinstance(v, date):
        return (datetime(v.year, v.month, v.day) - _EPOCH).days
    return None


def _counts(raw: List[Any], buckets: List[str]) -> List[Tuple[str, Any, List[int]]]:
    """(النص، القيمة الأصلية، [الإجمالي، مقفل، معالَج، مسترجع، بانتظار])."""
    agg: Dict[str, List[Any]] = {}
    for v, b in zip(raw, buckets):
        label = _clean(v) or "(غير محدد)"
        key = engine.normalize_ar(label) or label
        slot = agg.setdefault(key, [label, v, [0, 0, 0, 0, 0]])
        slot[2][0] += 1
        slot[2][1 + AP_STATES.index(b)] += 1
    return sorted(((v[0], v[1], v[2]) for v in agg.values()),
                  key=lambda t: (t[0] == "(غير محدد)", -t[2][0]))


def _table(sb: SheetBuilder, s: Dict[str, int], row: int, title: str,
           data: List[Tuple[str, List[int]]], src: str, ap: str) -> int:
    """جدول حيّ: كل خلية صيغة COUNTIFS تقرأ عمود «حالة الاعتماد» مباشرة."""
    band(sb, row, 1, 7, title, s["sect"])
    sb.height(row, 22)
    row += 1
    for j, h in enumerate(HEAD):
        sb.set(row, 1 + j, h, s["th"])
    sb.height(row, 30)
    first = row + 1

    for k, (label, raw, n) in enumerate(data):
        row += 1
        alt = k % 2 == 1
        crit = "$B%d" % row
        serial = _excel_date(raw)
        stamp = _DATE_TEXT.fullmatch(label)
        if serial is not None:                  # تاريخ حقيقي: يُكتب قيمة ويُنسّق
            sb.set(row, 1, serial, s["date_b"] if alt else s["date"])
        elif stamp:
            # التاريخ مخزَّن نصًّا مع وقت («… 00:00:00»): نعرض اليوم وحده،
            # ونطابق ببادئة كي يبقى العدّ صحيحًا رغم اختلاف النص المعروض
            sb.set(row, 1, stamp.group(1), s["name_b"] if alt else s["name"])
            crit = '$B%d&"*"' % row
        else:
            sb.set(row, 1, label, s["name_b"] if alt else s["name"])
        # صف «(غير محدد)» يُحسب طرحًا لا بـ COUNTIFS بمعيار فارغ: المعيار
        # الفارغ يعدّ الصفوف الخالية أسفل البيانات أيضًا فيتضخّم الرقم.
        blank = label == "(غير محدد)"
        for j in range(5):            # الإجمالي + الحالات الأربع
            col = 2 + j
            kpi = KPI_CELL[j].replace("6", "$6")
            if blank:
                f = ("%s-SUM(%s%d:%s%d)" % (kpi, xlsxedit.col_letter(col), first,
                                            xlsxedit.col_letter(col), row - 1)
                     if row > first else kpi)
            elif j == 0:
                f = "COUNTIFS(%s,%s)" % (src, crit)
            else:
                f = "COUNTIFS(%s,%s,%s,%s)" % (src, crit, ap, _q(AP_STATES[j - 1]))
            sb.set(row, col, n[j], s["td_b"] if alt else s["td"], formula=f)
        sb.set(row, 7, (n[1] / n[0]) if n[0] else 0.0, s["pct_b"] if alt else s["pct"],
               formula="IFERROR(D%d/C%d,0)" % (row, row))
        sb.height(row, 19)
    if data:
        sb.databar("H%d:H%d" % (first, row))
    return row + 2


def kpi_styles(st: Styles, colors: Dict[str, str]) -> Tuple[List[int], List[int]]:
    """أنماط بطاقات المؤشرات: شريط علوي بلون الخانة، ورقمها بلون مناسب.

    هكذا يصير الشريط نفسه دليل ألوان: من يفتح الملف يربط لون الصف بخانته
    دون شرح."""
    f_lbl = st.font(10, False, MUTED)
    f_num = st.font(20, True, INK)
    f_ok = st.font(20, True, GREEN)
    f_bad = st.font(20, True, "B03A3A")
    fl = st.fill(SOFT)
    pct = st.numfmt("0%")
    num = st.numfmt("#,##0")

    labels, nums = [], []
    accents = [INK] + [COLOR_HEX.get(colors.get(x) or "", "C9D4DF") for x in AP_STATES] + [INK]
    fonts = [f_num, f_ok, f_num, f_bad, f_num, f_num]
    for i, accent in enumerate(accents):
        top = st.border("FFFFFF", sides="lr", thick_top=accent)
        bot = st.border("FFFFFF", sides="lrb")
        labels.append(st.xf(f_lbl, fl, top, halign="center"))
        nums.append(st.xf(fonts[i], fl, bot, pct if i == 5 else num, halign="center"))
    return labels, nums


def build_datasheet(s: Dict[str, int], an: Dict[str, Any], buckets: List[str],
                    groups: List[Tuple[str, int]], ap_range: str,
                    col_range, kpi: Tuple[List[int], List[int]]) -> SheetBuilder:
    sb = SheetBuilder(tab_color=INK, landscape=True, selected=True, centered=True)
    for c, w in WIDTHS:
        sb.width(c, w)

    total = len(buckets)
    n = [buckets.count(st) for st in AP_STATES]

    sb.height(1, 8)
    sb.set(1, 1, None, s["spacer"])
    title = "داتا شيت حالة الاعتماد"
    if an.get("feeder"):
        title += " — " + an["feeder"]
    band(sb, 2, 1, 7, title, s["title"])
    sb.height(2, 36)
    # كل رقم محفوف بكلمات عربية والتاريخ في آخر السطر مسبوقًا بكلمة: السطر
    # المختلط في اتجاه RTL يعيد ترتيب مقاطعه إذا بدأ برقم أو فصلت بينها نقاط
    band(sb, 3, 1, 7, "إجمالي الملاحظات %d، %s %d، نسبة الإنجاز %.0f٪، بتاريخ %s"
         % (total, AP_APPROVED, n[0], 100.0 * n[0] / total if total else 0,
            datetime.now().strftime("%Y-%m-%d")), s["sub"])
    sb.height(3, 20)
    sb.height(4, 10)
    sb.set(4, 1, None, s["spacer"])

    # شريط المؤشرات: كل بطاقة فوق عمودها، وشريطها العلوي بلون خانتها
    kpi_lbl, kpi_num = kpi
    labels = ["إجمالي الملاحظات"] + AP_STATES + ["نسبة الإنجاز"]
    values = [total] + n + [(n[0] / total) if total else 0.0]
    formulas = (["COUNTA(%s)" % ap_range]
                + ["COUNTIF(%s,%s)" % (ap_range, _q(st)) for st in AP_STATES]
                + ["IFERROR(%s/%s,0)" % (KPI_CELL[1], KPI_CELL[0])])
    for j, (c1, c2) in enumerate(KPI_SPAN):
        band(sb, 5, c1, c2, labels[j], kpi_lbl[j])
        band(sb, 6, c1, c2, None, kpi_num[j])
        sb.set(6, c1, values[j], kpi_num[j], formula=formulas[j])
    sb.height(5, 32)
    sb.height(6, 34)
    sb.height(7, 14)
    sb.set(7, 1, None, s["spacer"])

    row = 8
    for k, (label, col) in enumerate(groups):
        vals = [r[col] if col < len(r) else None for r in an["rows"]]
        data = _counts(vals, buckets)
        # في الجداول التالية للأول: صف «(غير محدد)» بلا إقفال ولا معالجة ولا
        # استرجاع لا يضيف شيئًا فيُحذف. ويبقى في الجدول الأول لاكتمال التوزيع.
        if k and data and data[-1][0] == "(غير محدد)" and sum(data[-1][2][1:4]) == 0:
            data = data[:-1]
        if data:
            row = _table(sb, s, row, label, data, col_range(col), ap_range)
    return sb


# ------------------------------------------------------------------ المخرج

def build_output(an: Dict[str, Any], mapping: Dict[str, str],
                 colors: Dict[str, str], src_path: str, out_path: str, *,
                 group_cols: Optional[List[int]] = None) -> Dict[str, Any]:
    """ينسخ الملف كما هو ويضيف عمود الاعتماد وتلوينه الحيّ وورقة الداتا شيت."""
    rows, nums, headers = an["rows"], an["row_nums"], an["headers"]
    n_cols = len(headers)
    buckets = resolve(an, mapping)

    cell_values: Dict[Tuple[int, int], Any] = {}
    new_columns: List[Tuple[str, Dict[int, Any]]] = []
    values = {nums[i]: buckets[i] for i in range(len(rows))}
    if an["prev_col"] is not None:
        ap_col = an["prev_col"]
        for rn, v in values.items():
            cell_values[(rn, ap_col)] = v
    else:
        ap_col = n_cols
        new_columns.append((AP_HEADER, values))

    first_row = an["header_row"] + 1
    last_row = (max(nums) if nums else first_row) + EXTRA_ROWS
    quoted = "'%s'!" % an["sheet"].replace("'", "''")

    def col_range(idx: int) -> str:
        letter = xlsxedit.col_letter(idx)
        return "%s$%s$%d:$%s$%d" % (quoted, letter, first_row, letter, last_row)

    ap_letter = xlsxedit.col_letter(ap_col)
    ap_range = col_range(ap_col)

    if group_cols is None:
        group_cols = [c for c in (an.get("class_col"),) if c is not None]
    # «حسب تمت المعالجة بواسطة» تقرأ ركيكة — نعطي الأعمدة المعروفة اسمًا لائقًا
    nice = {an.get("contractor_col"): "حسب المقاول",
            an.get("class_col"): "حسب التصنيف الرئيسي",
            an.get("priority_col"): "حسب الأولوية"}
    groups = [(nice.get(c) or "حسب %s" % headers[c], c) for c in group_cols
              if 0 <= c < n_cols and headers[c]]

    def sheets_factory(st: Styles):
        s = make_styles(st)
        return [("الداتا شيت",
                 build_datasheet(s, an, buckets, groups, ap_range, col_range,
                                 kpi_styles(st, colors)))]

    cf = [("$%s%d=%s" % (ap_letter, first_row, _q(state)), COLOR_HEX[colors[state]])
          for state in AP_STATES if colors.get(state) in COLOR_HEX]

    # تعبئة ثابتة تحت التنسيق الشرطي: عارضات الجوال البسيطة (واتساب) لا تعرض
    # التنسيق الشرطي، فيصل الملف بلا ألوان لمن ليس عنده إكسل. وفي إكسل يغطّيها
    # التنسيق الشرطي فيبقى اللون حيًّا مع التعديل اليدوي.
    fills = {nums[i]: COLOR_HEX[colors[b]]
             for i, b in enumerate(buckets) if colors.get(b) in COLOR_HEX}
    total_cols = n_cols + len(new_columns)
    sqref = "A%d:%s%d" % (first_row, xlsxedit.col_letter(total_cols - 1), last_row)

    xlsxedit.write_patched(
        src_path, out_path,
        sheet_name=an["sheet"], header_row=an["header_row"], n_cols=n_cols,
        cell_values=cell_values, new_columns=new_columns, row_fills=fills,
        cf=(sqref, cf),
        validation=("%s%d:%s%d" % (ap_letter, first_row, ap_letter, last_row),
                    AP_STATES),
        sheets_factory=sheets_factory)

    counts = {st: buckets.count(st) for st in AP_STATES}
    return {
        "total": len(rows),
        "counts": counts,
        "percent": round(100.0 * counts[AP_APPROVED] / len(rows), 1) if rows else 0.0,
        "groups": [g[0] for g in groups],
        "new_cols": [t for t, _ in new_columns],
        "stem": os.path.splitext(os.path.basename(src_path))[0],
        "feeder": an.get("feeder", ""),
    }
