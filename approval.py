# -*- coding: utf-8 -*-
"""
approval.py
===========
متابعة حالة الاعتماد من ملف واحد — بلا ملفات ثانوية.

يقرأ عمود الحالة في ملف الملاحظات ويترجم رموزه (APPROVED / ASSIGNED / …)
إلى أربع خانات يفهمها المتابع: موافق عليه، منجز بانتظار الموافقة، معاد،
بانتظار المعالجة. ثم يضيف في آخر الملف عمود «حالة الاعتماد» بقائمة منسدلة
ليعدّله المستخدم بيده، ويلوّن الصفوف تنسيقًا شرطيًا، ويبني داتا شيت حيًّا
حسب المقاول والتصنيف والأولوية.

الملف الأصلي لا يُمسّ: لا تنسيق ولا ترتيب ولا خلية قائمة — فقط عمود جديد في
آخره وأوراق مضافة (انظر xlsxedit).
"""

import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import engine
import matcher
import xlsxedit
from matcher import _clean, find_col, read_table, make_styles, band, _q, COLOR_HEX
from xlsxedit import SheetBuilder, Styles

# ------------------------------------------------------------------ ثوابت

AP_APPROVED = "موافق عليه"
AP_DONE = "منجز بانتظار الموافقة"
AP_RETURNED = "معاد"
AP_PENDING = "بانتظار المعالجة"
AP_STATES = [AP_APPROVED, AP_DONE, AP_RETURNED, AP_PENDING]

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
    code = t.strip().upper().replace(" ", "_").replace("-", "_")
    if code in CODE_MAP:
        return CODE_MAP[code]
    ar = engine.normalize_ar(t)
    if any(w in ar for w in ("موافق", "معتمد", "مقبول", "مغلق")):
        return AP_APPROVED
    if any(w in ar for w in ("معاد", "مرفوض", "مرتجع", "اعاده")):
        return AP_RETURNED
    if any(w in ar for w in ("منجز", "منفذ", "مكتمل", "تم ")):
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
        if prev in AP_STATES:
            out.append(prev)          # إدخال يدوي سابق يُحترم ولا يُداس
            continue
        raw = (_clean(r[an["status_col"]])
               if an["status_col"] is not None and an["status_col"] < len(r) else "")
        key = raw or "(فارغة)"
        out.append(mapping.get(key) or guess_bucket(raw))
    return out


# ------------------------------------------------------------------ الداتا شيت

HEAD = ["القيمة", "الإجمالي"] + AP_STATES + ["نسبة الاعتماد"]
WIDTHS = [(0, 2.5), (1, 26), (2, 12), (3, 13), (4, 17), (5, 11), (6, 16), (7, 14), (8, 2.5)]


def _counts(values: List[str], buckets: List[str]) -> List[Tuple[str, List[int]]]:
    """(القيمة، [الإجمالي، موافق، منجز، معاد، بانتظار]) تنازليًا حسب الإجمالي."""
    agg: Dict[str, List[Any]] = {}
    for v, b in zip(values, buckets):
        label = v or "(غير محدد)"
        key = engine.normalize_ar(label) or label
        slot = agg.setdefault(key, [label, [0, 0, 0, 0, 0]])
        slot[1][0] += 1
        slot[1][1 + AP_STATES.index(b)] += 1
    return sorted(((v[0], v[1]) for v in agg.values()),
                  key=lambda t: (t[0] == "(غير محدد)", -t[1][0]))


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

    for k, (label, n) in enumerate(data):
        row += 1
        alt = k % 2 == 1
        sb.set(row, 1, label, s["name_b"] if alt else s["name"])
        # صف «(غير محدد)» يُحسب طرحًا لا بـ COUNTIFS بمعيار فارغ: المعيار
        # الفارغ يعدّ الصفوف الخالية أسفل البيانات أيضًا فيتضخّم الرقم.
        blank = label == "(غير محدد)"
        for j in range(5):            # الإجمالي + الحالات الأربع
            col = 2 + j
            kpi = "%s$6" % xlsxedit.col_letter(1 + j)
            if blank:
                f = ("%s-SUM(%s%d:%s%d)" % (kpi, xlsxedit.col_letter(col), first,
                                            xlsxedit.col_letter(col), row - 1)
                     if row > first else kpi)
            elif j == 0:
                f = "COUNTIFS(%s,$B%d)" % (src, row)
            else:
                f = "COUNTIFS(%s,$B%d,%s,%s)" % (src, row, ap, _q(AP_STATES[j - 1]))
            sb.set(row, col, n[j], s["td_b"] if alt else s["td"], formula=f)
        sb.set(row, 7, (n[1] / n[0]) if n[0] else 0.0, s["pct_b"] if alt else s["pct"],
               formula="IFERROR(D%d/C%d,0)" % (row, row))
        sb.height(row, 19)
    if data:
        sb.databar("H%d:H%d" % (first, row))
    return row + 2


def build_datasheet(s: Dict[str, int], an: Dict[str, Any], buckets: List[str],
                    groups: List[Tuple[str, int]], ap_range: str,
                    col_range) -> SheetBuilder:
    sb = SheetBuilder(tab_color=INK, landscape=True)
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
    band(sb, 3, 1, 7, "%s · %d ملاحظة · %d موافق عليها (%.0f%%)" % (
        datetime.now().strftime("%Y-%m-%d"), total, n[0],
        100.0 * n[0] / total if total else 0), s["sub"])
    sb.height(3, 20)
    sb.height(4, 10)
    sb.set(4, 1, None, s["spacer"])

    # شريط المؤشرات: ستّ بطاقات، الأخيرة (النسبة) على عمودين
    labels = ["إجمالي الملاحظات"] + AP_STATES
    values = [total] + n
    styles = [s["kpi_num"], s["kpi_ok"], s["kpi_num"], s["kpi_num"], s["kpi_num"]]
    formulas = ["COUNTA(%s)" % ap_range] + [
        "COUNTIF(%s,%s)" % (ap_range, _q(st)) for st in AP_STATES]
    for j in range(5):
        sb.set(5, 1 + j, labels[j], s["kpi_lbl"])
        sb.set(6, 1 + j, values[j], styles[j], formula=formulas[j])
    band(sb, 5, 6, 7, "نسبة الاعتماد", s["kpi_lbl"])
    sb.set(6, 6, (n[0] / total) if total else 0.0, s["kpi_pct"],
           formula="IFERROR(C6/B6,0)")
    sb.set(6, 7, None, s["kpi_pct"])
    sb.merge(6, 6, 6, 7)
    sb.height(5, 26)
    sb.height(6, 34)
    sb.height(7, 14)
    sb.set(7, 1, None, s["spacer"])

    row = 8
    for label, col in groups:
        vals = [_clean(r[col]) if col < len(r) else "" for r in an["rows"]]
        data = _counts(vals, buckets)
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
                 build_datasheet(s, an, buckets, groups, ap_range, col_range))]

    cf = [("$%s%d=%s" % (ap_letter, first_row, _q(state)), COLOR_HEX[colors[state]])
          for state in AP_STATES if colors.get(state) in COLOR_HEX]
    total_cols = n_cols + len(new_columns)
    sqref = "A%d:%s%d" % (first_row, xlsxedit.col_letter(total_cols - 1), last_row)

    xlsxedit.write_patched(
        src_path, out_path,
        sheet_name=an["sheet"], header_row=an["header_row"], n_cols=n_cols,
        cell_values=cell_values, new_columns=new_columns,
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
