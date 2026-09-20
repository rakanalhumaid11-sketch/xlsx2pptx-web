# -*- coding: utf-8 -*-
"""
matcher.py
==========
مطابقة الأعمال المنفّذة: ملف رئيسي واحد + عدة ملفات ثانوية.

الفكرة: الملفات الثانوية ليست بيانات تُدمج، بل مجرد *مرجع* يقول أي أرقام
ملاحظات تم تنفيذها. فكل رقم ملاحظة موجود في ملف ثانوي يُعلَّم في الملف
الرئيسي (عمود «منفذ» + تلوين الصف)، ثم يُبنى داتا شيت من الملف الرئيسي
نفسه: كم أُنجز، وما توزيع التصنيفات، وكم بقي.

الملف الرئيسي لا يُعدَّل أبدًا؛ المخرج ملف متابعة جديد.
"""

import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import engine

# ------------------------------------------------------------------ ثوابت

ID_ALIASES = ["رقم الملاحظة", "رقم الملاحظه", "الملاحظة رقم", "رقم الملاحظات",
              "رقم البلاغ", "رقم"]
CLASS_ALIASES = ["التصنيف الرئيسي", "التصنيف الرئيسى", "التصنيف", "تصنيف الملاحظة",
                 "نوع الملاحظة", "النوع"]
CONTRACTOR_ALIASES = ["المقاول", "اسم المقاول", "الشركة", "الشركه"]
STATUS_ALIASES = ["حالة الملاحظة", "حالة الملاحظه", "الحالة", "الحاله"]

DONE_HEADER = "منفذ"
SOURCE_HEADER = "مصدر التنفيذ"
YES, NO = "نعم", "لا"
DONE_STATUS = "منجز"

MAX_RULES = 6
MAX_DISTINCT = 60          # أكثر من هذا يصير عمودًا حرًّا لا قائمة اختيار
SCAN_LIMIT = 100_000       # سقف أمان لعدد الصفوف المقروءة

# ألوان فاتحة متعمَّدة: النص الأسود يبقى مقروءًا فوقها عند الطباعة
COLORS: List[Tuple[str, str, str]] = [
    ("green",  "أخضر",    "C6EFCE"),
    ("yellow", "أصفر",    "FFEB9C"),
    ("red",    "أحمر",    "FFC7CE"),
    ("blue",   "أزرق",    "BDD7EE"),
    ("orange", "برتقالي", "FCE4D6"),
    ("purple", "بنفسجي",  "E4DFEC"),
    ("gray",   "رمادي",   "D9D9D9"),
]
COLOR_HEX = {c[0]: c[2] for c in COLORS}
COLOR_NAME = {c[0]: c[1] for c in COLORS}

HEAD_FILL = PatternFill("solid", fgColor="1F3A52")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F3A52")
SUB_FONT = Font(bold=True, size=11, color="1F3A52")
THIN = Side(style="thin", color="BFCAD4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


# ------------------------------------------------------------------ أدوات

def _clean(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.startswith("#") and s.upper() in ("#N/A", "#VALUE!", "#REF!", "#NAME?", "#DIV/0!"):
        return ""
    # إكسل يقرأ الأرقام الصحيحة أحيانًا كـ 1234.0 — نحذف الكسر الصفري
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


def norm_id(v: Any) -> str:
    """توحيد شكل رقم الملاحظة قبل المقارنة: أرقام عربية، فواصل، مسافات.

    الفرق تكتب الرقم بأشكال مختلفة (‎M13-8903‎ / ‎m13 8903‎ / ‎١٣٨٩٠٣‎)،
    وبدون التوحيد تفشل المطابقة بلا سبب ظاهر للمستخدم."""
    s = _clean(v).translate(_AR_DIGITS)
    if not s:
        return ""
    return re.sub(r"[\s\-_/\\.،,]+", "", s).upper()


def find_col(headers: List[str], aliases: List[str]) -> Optional[int]:
    """رقم العمود المطابق لأحد الأسماء — المطابقة التامة تغلب البادئة."""
    norm_aliases = [engine.normalize_ar(a) for a in aliases]
    exact = prefix = None
    for idx, h in enumerate(headers):
        if not h:
            continue
        nh = engine.normalize_ar(h)
        if nh in norm_aliases:
            if exact is None:
                exact = idx
        else:
            for na in norm_aliases:
                if na and nh.startswith(na) and prefix is None:
                    prefix = idx
    return exact if exact is not None else prefix


def read_table(path: str) -> Tuple[str, List[str], List[List[Any]]]:
    """يقرأ أنسب ورقة في الملف ويرجّع (اسم الورقة، الترويسة، الصفوف)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        best = None
        for name in wb.sheetnames:
            ws = wb[name]
            try:
                hdr = engine._detect_header_row(ws)
            except Exception:  # noqa: BLE001
                hdr = 0
            headers: List[str] = []
            n_rows = 0
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i > SCAN_LIMIT:
                    break
                if i == hdr:
                    headers = [_clean(c) for c in row]
                elif i > hdr and any(c is not None and str(c).strip() for c in row):
                    n_rows += 1
            has_id = find_col(headers, ID_ALIASES) is not None
            score = (1 if has_id else 0, n_rows)
            if best is None or score > best[0]:
                best = (score, name, hdr, headers)

        if best is None:
            raise ValueError("الملف لا يحتوي على أوراق عمل.")
        _, sheet_name, hdr, headers = best

        while headers and not headers[-1]:
            headers.pop()
        if not headers:
            raise ValueError("لم يُعثر على صف عناوين في الملف.")

        ws = wb[sheet_name]
        width = len(headers)
        rows: List[List[Any]] = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i <= hdr:
                continue
            if i > SCAN_LIMIT:
                break
            # نحتفظ بالقيم كما هي (أرقام تبقى أرقامًا وتواريخ تبقى تواريخ) كي
            # يخرج ملف المتابعة قابلًا للفرز والحساب لا نصًّا صامتًا
            vals = [row[c] if c < len(row) else None for c in range(width)]
            vals = [None if _clean(v) == "" else v for v in vals]
            if not any(v is not None for v in vals):
                continue
            rows.append(vals)
        return sheet_name, headers, rows
    finally:
        wb.close()


def read_ids(path: str) -> List[str]:
    """أرقام الملاحظات في ملف ثانوي — نقرأ عمود الرقم فقط لا غير."""
    _sheet, headers, rows = read_table(path)
    col = find_col(headers, ID_ALIASES)
    out: List[str] = []
    if col is None:
        # لا عمود معنون: نأخذ أول عمود تكون أغلب قيمه أرقام ملاحظات معقولة
        col = _guess_id_col(headers, rows)
    if col is None:
        return out
    for r in rows:
        nid = norm_id(r[col]) if col < len(r) else ""
        if nid:
            out.append(nid)
    return out


def _guess_id_col(headers: List[str], rows: List[List[Any]]) -> Optional[int]:
    best, best_score = None, 0
    sample = rows[:200]
    for c in range(len(headers)):
        vals = [norm_id(r[c]) for r in sample if c < len(r)]
        vals = [v for v in vals if v]
        if not vals:
            continue
        uniq = len(set(vals)) / len(vals)
        looks = sum(1 for v in vals if re.fullmatch(r"[A-Z]{0,4}\d{3,12}", v)) / len(vals)
        score = looks * uniq
        if score > best_score:
            best, best_score = c, score
    return best if best_score >= 0.6 else None


# ------------------------------------------------------------------ التحليل

def analyze(main_path: str, sec_paths: List[str]) -> Dict[str, Any]:
    """يطابق أرقام الملفات الثانوية على الملف الرئيسي دون تعديل أي منهما."""
    sheet, headers, rows = read_table(main_path)

    id_col = find_col(headers, ID_ALIASES)
    if id_col is None:
        id_col = _guess_id_col(headers, rows)
    if id_col is None:
        raise ValueError("لم يُعثر على عمود «رقم الملاحظة» في الملف الرئيسي.")

    # أرقام الملفات الثانوية: كل رقم -> اسم الملف الذي ورد فيه
    done_src: Dict[str, str] = {}
    sec_names: List[str] = []
    for p in sec_paths:
        name = os.path.basename(p)
        sec_names.append(name)
        for nid in read_ids(p):
            done_src.setdefault(nid, name)

    # عمود «منفذ» سابق في الملف الرئيسي: نحترمه كي تتراكم النتائج أسبوعًا بعد أسبوع
    prev_done_col = find_col(headers, [DONE_HEADER])

    matched: List[int] = []
    carried = 0
    sources: Dict[int, str] = {}
    seen: Dict[str, int] = {}
    dups: List[Tuple[str, int]] = []
    hit_ids = set()

    for i, r in enumerate(rows):
        nid = norm_id(r[id_col]) if id_col < len(r) else ""
        if nid:
            if nid in seen:
                dups.append((_clean(r[id_col]), i + 1))
            else:
                seen[nid] = i
        if nid and nid in done_src:
            matched.append(i)
            sources[i] = done_src[nid]
            hit_ids.add(nid)
        elif prev_done_col is not None and prev_done_col < len(r) \
                and engine.normalize_ar(_clean(r[prev_done_col])) == engine.normalize_ar(YES):
            matched.append(i)
            sources[i] = "جولة سابقة"
            carried += 1

    missing = [(nid, fn) for nid, fn in done_src.items() if nid not in hit_ids]

    # القيم المتكررة في كل عمود — تغذّي قوائم اختيار قواعد التلوين
    distinct: Dict[int, List[str]] = {}
    for c in range(len(headers)):
        vals: List[str] = []
        seen_v = set()
        for r in rows:
            v = _clean(r[c]) if c < len(r) else ""
            if v and v not in seen_v:
                seen_v.add(v)
                vals.append(v)
                if len(vals) > MAX_DISTINCT:
                    break
        if 0 < len(vals) <= MAX_DISTINCT:
            distinct[c] = sorted(vals)

    return {
        "sheet": sheet,
        "headers": headers,
        "rows": rows,
        "id_col": id_col,
        "class_col": find_col(headers, CLASS_ALIASES),
        "contractor_col": find_col(headers, CONTRACTOR_ALIASES),
        "status_col": find_col(headers, STATUS_ALIASES),
        "prev_done_col": prev_done_col,
        "matched": matched,
        "sources": sources,
        "carried": carried,
        "missing": missing,
        "dups": dups,
        "distinct": distinct,
        "sec_names": sec_names,
        "total": len(rows),
    }


# ------------------------------------------------------------------ المخرج

def _fill_for(row: List[Any], done: bool,
              rules: List[Dict[str, Any]]) -> Optional[PatternFill]:
    """أول قاعدة تنطبق هي التي تلوّن — الترتيب في الصفحة هو الأولوية."""
    for rule in rules:
        col = rule.get("col")
        color = rule.get("color")
        if not color or color not in COLOR_HEX:
            continue
        if col == -1:                       # قاعدة «المنفذ»
            if done:
                return PatternFill("solid", fgColor=COLOR_HEX[color])
            continue
        if col is None or col < 0 or col >= len(row):
            continue
        want = engine.normalize_ar(rule.get("value") or "")
        if not want:
            continue
        if engine.normalize_ar(_clean(row[col])) == want:
            return PatternFill("solid", fgColor=COLOR_HEX[color])
    return None


def _fit_width(ws):
    """تُطبع الورقة بعرض صفحة واحدة مهما زادت الأعمدة."""
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def _style_header(ws, n_cols: int, row: int = 1):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _autosize(ws, headers: List[str], rows: List[List[Any]], limit: int = 300):
    for c, h in enumerate(headers, start=1):
        width = len(str(h))
        for r in rows[:limit]:
            if c - 1 < len(r):
                width = max(width, len(_clean(r[c - 1])))
        ws.column_dimensions[get_column_letter(c)].width = min(max(width + 3, 10), 45)


def build_output(an: Dict[str, Any], rules: List[Dict[str, Any]], out_path: str,
                 set_status_done: bool = False, add_source_col: bool = True) -> Dict[str, Any]:
    headers = list(an["headers"])
    rows = an["rows"]
    matched = set(an["matched"])

    wb = Workbook()

    # ---------------------------------------------------- ورقة الملاحظات
    ws = wb.active
    ws.title = (an["sheet"] or "الملاحظات")[:31]
    ws.sheet_view.rightToLeft = True

    out_headers = list(headers)
    done_col = an.get("prev_done_col")
    if done_col is None:
        done_col = len(out_headers)
        out_headers.append(DONE_HEADER)
    src_col = None
    if add_source_col:
        src_col = find_col(out_headers, [SOURCE_HEADER])
        if src_col is None:
            src_col = len(out_headers)
            out_headers.append(SOURCE_HEADER)

    ws.append(out_headers)
    _style_header(ws, len(out_headers))

    status_col = an.get("status_col")
    for i, r in enumerate(rows):
        vals = list(r) + [None] * (len(out_headers) - len(r))
        done = i in matched
        vals[done_col] = YES if done else NO
        if src_col is not None:
            vals[src_col] = an["sources"].get(i, "")
        if set_status_done and done and status_col is not None:
            vals[status_col] = DONE_STATUS
        ws.append(vals)
        fill = _fill_for(r, done, rules)
        if fill is not None:
            for c in range(1, len(out_headers) + 1):
                ws.cell(row=i + 2, column=c).fill = fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(out_headers))}{len(rows) + 1}"
    _autosize(ws, out_headers, rows)

    # ---------------------------------------------------- الداتا شيت
    _build_datasheet(wb, an, matched, set_status_done)

    # ---------------------------------------------------- التنبيهات
    n_alerts = _build_alerts(wb, an)

    wb.save(out_path)
    wb.close()

    n_done = len(matched)
    return {
        "total": len(rows),
        "done": n_done,
        "remaining": len(rows) - n_done,
        "percent": round(100.0 * n_done / len(rows), 1) if rows else 0.0,
        "carried": an.get("carried", 0),
        "missing": len(an["missing"]),
        "dups": len(an["dups"]),
        "alerts": n_alerts,
    }


def _breakdown(rows: List[List[Any]], matched: set, col: Optional[int]) -> List[Tuple[str, int, int]]:
    """(القيمة، الإجمالي، المنفذ) مرتبة تنازليًا حسب الإجمالي.

    التجميع يتجاهل اختلاف حالة الأحرف وشكل الألف والهمزة، وإلا ظهر ‎mv‎ و‎MV‎
    صفّين منفصلين في الداتا شيت مع أنهما تصنيف واحد."""
    if col is None:
        return []
    agg: Dict[str, List[Any]] = {}
    for i, r in enumerate(rows):
        raw = (_clean(r[col]) if col < len(r) else "") or "(غير محدد)"
        key = engine.normalize_ar(raw) or raw
        slot = agg.setdefault(key, [raw, 0, 0])
        slot[1] += 1
        if i in matched:
            slot[2] += 1
    return sorted(((v[0], v[1], v[2]) for v in agg.values()), key=lambda t: -t[1])


def _table(ws, row: int, title: str, head: List[str],
           data: List[List[Any]], pct_col: Optional[int] = None) -> int:
    """يرسم جدولًا معنونًا ويرجّع رقم الصف التالي الفارغ."""
    ws.cell(row=row, column=1, value=title).font = SUB_FONT
    row += 1
    for c, h in enumerate(head, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    first = row + 1
    for d in data:
        row += 1
        for c, v in enumerate(d, start=1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")
            if pct_col is not None and c == pct_col:
                cell.number_format = "0%"
    if pct_col is not None and data:
        letter = get_column_letter(pct_col)
        ws.conditional_formatting.add(
            f"{letter}{first}:{letter}{row}",
            DataBarRule(start_type="num", start_value=0, end_type="num",
                        end_value=1, color="4CAF7D"))
    return row + 3


def _build_datasheet(wb: Workbook, an: Dict[str, Any], matched: set, status_changed: bool):
    rows = an["rows"]
    ws = wb.create_sheet("الداتا شيت")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    total = len(rows)
    done = len(matched)
    rest = total - done
    pct = (done / total) if total else 0.0

    ws.cell(row=1, column=1, value="داتا شيت متابعة الإنجاز").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value=f"{datetime.now().strftime('%Y-%m-%d')} · "
                  f"مبني على {len(an['sec_names'])} ملف تنفيذ · "
                  f"{done} منفذة من {total}").font = Font(size=10, color="5A6B7B")

    r = _table(ws, 4, "الإنجاز الإجمالي",
               ["البند", "العدد", "النسبة"],
               [["منفذ", done, pct],
                ["متبقي", rest, (rest / total) if total else 0.0],
                ["الإجمالي", total, 1.0 if total else 0.0]],
               pct_col=3)

    for title, col in (("حسب التصنيف الرئيسي", an.get("class_col")),
                       ("حسب المقاول", an.get("contractor_col"))):
        bd = _breakdown(rows, matched, col)
        if not bd:
            continue
        data = [[k, tot, dn, tot - dn, (dn / tot) if tot else 0.0] for k, tot, dn in bd]
        r = _table(ws, r, title,
                   ["القيمة", "الإجمالي", "منفذ", "متبقي", "نسبة الإنجاز"],
                   data, pct_col=5)

    for c, w in enumerate([30, 12, 12, 12, 15], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    _fit_width(ws)


def _build_alerts(wb: Workbook, an: Dict[str, Any]) -> int:
    missing, dups = an["missing"], an["dups"]
    if not missing and not dups:
        return 0
    ws = wb.create_sheet("تنبيهات")
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value="تنبيهات المطابقة").font = TITLE_FONT

    r = 3
    if missing:
        r = _table(ws, r, "أرقام في ملفات التنفيذ ولا وجود لها في الملف الرئيسي",
                   ["رقم الملاحظة", "الملف الثانوي"],
                   [[nid, fn] for nid, fn in missing[:2000]])
    if dups:
        r = _table(ws, r, "أرقام مكررة داخل الملف الرئيسي",
                   ["رقم الملاحظة", "رقم الصف"],
                   [[nid, ln] for nid, ln in dups[:2000]])

    for c, w in enumerate([26, 40], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    _fit_width(ws)
    return len(missing) + len(dups)
