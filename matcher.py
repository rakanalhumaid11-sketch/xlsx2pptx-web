# -*- coding: utf-8 -*-
"""
matcher.py
==========
مطابقة الأعمال المنفّذة: ملف رئيسي واحد + عدة ملفات ثانوية.

الملفات الثانوية ليست بيانات تُدمج، بل *مرجع* يقول أي أرقام ملاحظات نُفّذت.
كل رقم موجود في ملف ثانوي يُعلَّم في الملف الرئيسي، ثم يُبنى داتا شيت من
الملف الرئيسي نفسه: كم أُنجز، وما توزيع التصنيفات، وكم بقي.

الملف الرئيسي يُنسخ كما هو بالضبط (صوره وتنسيقاته وترتيب أعمدته وأوراقه
الأخرى) عبر xlsxedit؛ لا يُضاف إليه إلا تلوين الصفوف وأعمدة في آخره.
"""

import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook

import engine
import xlsxedit
from xlsxedit import SheetBuilder, Styles

# ------------------------------------------------------------------ ثوابت

ID_ALIASES = ["رقم الملاحظة", "رقم الملاحظه", "الملاحظة رقم", "رقم الملاحظات",
              "رقم البلاغ", "رقم"]
CLASS_ALIASES = ["التصنيف الرئيسي", "التصنيف الرئيسى", "التصنيف", "تصنيف الملاحظة",
                 "نوع الملاحظة", "النوع"]
CONTRACTOR_ALIASES = ["المقاول", "اسم المقاول", "الشركة", "الشركه", "المنفذ",
                      "تمت المعالجة بواسطة", "تمت المعالجه بواسطة"]
STATUS_ALIASES = ["حالة الملاحظة", "حالة الملاحظه", "الحالة", "الحاله"]

SEC_STATUS_ALIASES = ["حالة التنفيذ", "الحالة", "حالة الملاحظة", "الوضع",
                      "منفذ", "التنفيذ", "حالة العمل"]

EXEC_HEADER = "حالة التنفيذ"
EXEC_ALIASES = [EXEC_HEADER, "منفذ"]          # «منفذ» تسمية جولات سابقة
SOURCE_HEADER = "مصدر التنفيذ"
CONTRACTOR_HEADER = "المقاول"

ST_DONE, ST_WIP, ST_NONE = "منفذ", "جاري العمل", "لم ينفذ"
EXEC_STATES = [ST_DONE, ST_WIP, ST_NONE]
DONE_STATUS = "منجز"

MAX_RULES = 6
MAX_DISTINCT = 60
SCAN_LIMIT = 100_000

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

# لوحة الداتا شيت
INK = "1F3A52"       # كحلي العناوين
SOFT = "EEF3F8"      # خلفية فاتحة
BAND = "F6F9FC"      # تظليل الصفوف المتناوبة
MUTED = "5A6B7B"
GREEN = "2E7D57"
LINE = "D5DFE9"

_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


# ------------------------------------------------------------------ أدوات

def _clean(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.startswith("#") and s.upper() in ("#N/A", "#VALUE!", "#REF!", "#NAME?", "#DIV/0!"):
        return ""
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


def map_status(text: Any) -> Optional[str]:
    """يترجم ما كتبه الفريق في عمود الحالة إلى إحدى الحالات الثلاث.

    الفرق تكتبها بصيغ شتى (تم / منجز / مكتمل / نعم …)، فنقبلها كلها بدل أن
    نطالبهم بصياغة واحدة."""
    t = engine.normalize_ar(_clean(text))
    if not t:
        return None
    if any(w in t for w in ("جاري", "قيد", "تحت التنفيذ", "جزئ", "بدا", "مستمر")):
        return ST_WIP
    if t.startswith("لم") or t in ("لا", "غير منجز", "غير منفذ", "معلق", "متبقي"):
        return ST_NONE
    if any(w in t for w in ("منجز", "منفذ", "تم", "مكتمل", "انجز", "نعم", "مغلق")):
        return ST_DONE
    return None


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


def read_table(path: str) -> Tuple[str, int, List[str], List[List[Any]], List[int]]:
    """(اسم الورقة، رقم صف الترويسة في إكسل، الترويسة، الصفوف، أرقام صفوفها)."""
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
        row_nums: List[int] = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i <= hdr:
                continue
            if i > SCAN_LIMIT:
                break
            # القيم تبقى كما هي (أرقام أرقامًا وتواريخ تواريخ) — لا نلمس الملف
            vals = [row[c] if c < len(row) else None for c in range(width)]
            vals = [None if _clean(v) == "" else v for v in vals]
            if not any(v is not None for v in vals):
                continue
            rows.append(vals)
            row_nums.append(i + 1)          # إكسل يبدأ من 1
        return sheet_name, hdr + 1, headers, rows, row_nums
    finally:
        wb.close()


def _guess_id_col(headers: List[str], rows: List[List[Any]]) -> Optional[int]:
    best, best_score = None, 0.0
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


def read_pairs(path: str) -> Tuple[List[Tuple[str, str, Optional[str]]], bool]:
    """((رقم، مقاول، حالة)…، هل في الملف عمود حالة) من ملف تنفيذ ثانوي."""
    _sheet, _hdr, headers, rows, _nums = read_table(path)
    col = find_col(headers, ID_ALIASES)
    if col is None:
        col = _guess_id_col(headers, rows)
    if col is None:
        return [], False
    ccol = find_col(headers, CONTRACTOR_ALIASES)
    scol = find_col(headers, SEC_STATUS_ALIASES)
    if scol == col or scol == ccol:
        scol = None
    out = []
    for r in rows:
        nid = norm_id(r[col]) if col < len(r) else ""
        if not nid:
            continue
        name = _clean(r[ccol]) if ccol is not None and ccol < len(r) else ""
        state = map_status(r[scol]) if scol is not None and scol < len(r) else None
        out.append((nid, name, state))
    return out, scol is not None


# ------------------------------------------------------------------ التحليل

def analyze(main_path: str, sec_paths: List[str]) -> Dict[str, Any]:
    """يطابق أرقام الملفات الثانوية على الملف الرئيسي دون تعديل أي منهما."""
    sheet, header_row, headers, rows, row_nums = read_table(main_path)

    id_col = find_col(headers, ID_ALIASES)
    if id_col is None:
        id_col = _guess_id_col(headers, rows)
    if id_col is None:
        raise ValueError("لم يُعثر على عمود «رقم الملاحظة» في الملف الرئيسي.")

    # رقم -> (اسم الملف، المقاول، الحالة المقروءة من الملف الثانوي)
    found: Dict[str, Tuple[str, str, Optional[str]]] = {}
    sec_names: List[str] = []
    sec_has_contractor = False
    sec_has_status = False
    n_sec_states = {ST_DONE: 0, ST_WIP: 0, ST_NONE: 0}
    for p in sec_paths:
        name = os.path.basename(p)
        sec_names.append(name)
        pairs, has_status = read_pairs(p)
        sec_has_status = sec_has_status or has_status
        for nid, contractor, state in pairs:
            if contractor:
                sec_has_contractor = True
            if state:
                n_sec_states[state] += 1
            old = found.get(nid)
            if old is None:
                found[nid] = (name, contractor, state)
            else:
                # عند التكرار: الحالة الأقوى والاسم غير الفارغ هما ما يبقى
                rank = {None: 0, ST_NONE: 1, ST_WIP: 2, ST_DONE: 3}
                found[nid] = (old[0],
                              old[1] or contractor,
                              state if rank[state] > rank[old[2]] else old[2])

    # عمود حالة تنفيذ سابق: يُحترم كي تتراكم نتائج الجولات بدل أن تُستبدل
    prev_exec_col = find_col(headers, EXEC_ALIASES)

    states: List[str] = []
    found_rows: List[int] = []
    sources: Dict[int, str] = {}
    sec_contractor: Dict[int, str] = {}
    carried = 0
    seen: Dict[str, int] = {}
    dups: List[Tuple[str, int]] = []
    hit_ids = set()

    for i, r in enumerate(rows):
        nid = norm_id(r[id_col]) if id_col < len(r) else ""
        if nid:
            if nid in seen:
                dups.append((_clean(r[id_col]), row_nums[i]))
            else:
                seen[nid] = i

        prev = (map_status(r[prev_exec_col])
                if prev_exec_col is not None and prev_exec_col < len(r) else None)

        if nid and nid in found:
            fname, contractor, state = found[nid]
            sources[i] = fname
            if contractor:
                sec_contractor[i] = contractor
            hit_ids.add(nid)
            found_rows.append(i)
            states.append(state or "")          # "" = يقرّره اختيار المستخدم
        elif prev in (ST_DONE, ST_WIP):
            sources[i] = "جولة سابقة"
            carried += 1
            states.append(prev)
        else:
            states.append(ST_NONE)

    missing = [(nid, v[0]) for nid, v in found.items() if nid not in hit_ids]

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
        "header_row": header_row,
        "headers": headers,
        "rows": rows,
        "row_nums": row_nums,
        "id_col": id_col,
        "class_col": find_col(headers, CLASS_ALIASES),
        "contractor_col": find_col(headers, CONTRACTOR_ALIASES),
        "status_col": find_col(headers, STATUS_ALIASES),
        "prev_exec_col": prev_exec_col,
        "prev_source_col": find_col(headers, [SOURCE_HEADER]),
        "states": states,
        "found_rows": found_rows,
        "sources": sources,
        "sec_contractor": sec_contractor,
        "sec_has_contractor": sec_has_contractor,
        "sec_has_status": sec_has_status,
        "n_sec_states": n_sec_states,
        "n_found": len(hit_ids),
        "carried": carried,
        "missing": missing,
        "dups": dups,
        "distinct": distinct,
        "sec_names": sec_names,
        "total": len(rows),
    }




# ------------------------------------------------------------------ التلوين

EXTRA_ROWS = 500      # مدى إضافي أسفل البيانات كي يشمل ما يضيفه المستخدم يدويًا


def _q(text: str) -> str:
    """نص داخل صيغة إكسل: علامة الاقتباس تُضاعَف."""
    return '"%s"' % str(text).replace('"', '""')


def state_cf_rules(rules: List[Dict[str, Any]], first_row: int,
                   exec_letter: str) -> List[Tuple[str, str]]:
    """ألوان الحالة كصيغ تنسيق شرطي — يتغيّر اللون فور تعديل المستخدم للخلية.

    وهي تعلو على تلوين قواعد الأعمدة: الصف المنفَّذ يظهر بلون حالته."""
    out: List[Tuple[str, str]] = []
    for rule in rules:
        if rule.get("kind") != "state":
            continue
        color = COLOR_HEX.get(rule.get("color") or "")
        if color:
            out.append(("$%s%d=%s" % (exec_letter, first_row, _q(rule["key"])), color))
    return out


def column_fills(rules: List[Dict[str, Any]], rows: List[List[Any]],
                 row_nums: List[int], matched: set) -> Dict[int, str]:
    """قواعد الأعمدة تلوّن الصفوف التي وردت في ملفات التنفيذ **وحدها**.

    تلوين كل صفوف الملف التي تحمل القيمة يُفقد التقرير معناه: المقصود تمييز
    ما سُلّم للفرق فعلًا، لا كل ما في الملف."""
    out: Dict[int, str] = {}
    for i, r in enumerate(rows):
        if i not in matched:
            continue
        for rule in rules:
            if rule.get("kind") == "state":
                continue
            color = COLOR_HEX.get(rule.get("color") or "")
            col = rule.get("key")
            want = engine.normalize_ar(rule.get("value") or "")
            if not color or not want or not isinstance(col, int) or col >= len(r):
                continue
            if engine.normalize_ar(_clean(r[col])) == want:
                out[row_nums[i]] = color
                break              # أول قاعدة تنطبق هي التي تلوّن
    return out


def _breakdown(values: List[str], states: List[str]) -> List[Tuple[str, int, int, int]]:
    """(القيمة، الإجمالي، المنفّذ، الجاري) مرتبة تنازليًا حسب الإجمالي.

    التجميع يتجاهل اختلاف حالة الأحرف وشكل الألف والهمزة، وإلا ظهر ‎mv‎
    و‎MV‎ صفّين منفصلين مع أنهما تصنيف واحد."""
    agg: Dict[str, List[Any]] = {}
    for i, raw in enumerate(values):
        label = raw or "(غير محدد)"
        key = engine.normalize_ar(label) or label
        slot = agg.setdefault(key, [label, 0, 0, 0])
        slot[1] += 1
        if states[i] == ST_DONE:
            slot[2] += 1
        elif states[i] == ST_WIP:
            slot[3] += 1
    # «(غير محدد)» يُدفع لآخر الجدول مهما كبر عدده — ليس فئة حقيقية
    return sorted((tuple(v) for v in agg.values()),
                  key=lambda t: (t[0] == "(غير محدد)", -t[1]))


# ------------------------------------------------------------------ الداتا شيت

def make_styles(st: Styles) -> Dict[str, int]:
    """يسجّل أنماط الأوراق الجديدة داخل جدول أنماط الملف الأصلي."""
    f_title = st.font(16, True, "FFFFFF")
    f_sub = st.font(10, False, MUTED)
    f_sect = st.font(12, True, INK)
    f_th = st.font(11, True, "FFFFFF")
    f_td = st.font(11, False, "233140")
    f_kpi = st.font(20, True, INK)
    f_kpi_ok = st.font(20, True, GREEN)
    f_lbl = st.font(10, False, MUTED)

    fl_ink = st.fill(INK)
    fl_soft = st.fill(SOFT)
    fl_band = st.fill(BAND)

    b_all = st.border(LINE)
    b_none = st.border(LINE, sides="")
    b_sect = st.border(LINE, sides="", thick_bottom=INK)
    # حدود بيضاء بين بطاقات المؤشرات: تفصلها بصريًا دون خطوط ظاهرة
    b_top = st.border("FFFFFF", sides="lrt")
    b_bot = st.border("FFFFFF", sides="lrb")

    pct = st.numfmt("0%")
    num = st.numfmt("#,##0")

    return {
        "title": st.xf(f_title, fl_ink, b_none, halign="right", indent=1),
        "sub": st.xf(f_sub, fl_soft, b_none, halign="right", indent=1),
        "sect": st.xf(f_sect, 0, b_sect, halign="right"),
        "spacer": st.xf(0, 0, b_none),
        "th": st.xf(f_th, fl_ink, b_all, halign="center", wrap=True),
        "td": st.xf(f_td, 0, b_all, num, halign="center"),
        "td_b": st.xf(f_td, fl_band, b_all, num, halign="center"),
        "name": st.xf(f_td, 0, b_all, halign="right", indent=1),
        "name_b": st.xf(f_td, fl_band, b_all, halign="right", indent=1),
        "pct": st.xf(f_td, 0, b_all, pct, halign="center"),
        "pct_b": st.xf(f_td, fl_band, b_all, pct, halign="center"),
        "kpi_lbl": st.xf(f_lbl, fl_soft, b_top, halign="center"),
        "kpi_num": st.xf(f_kpi, fl_soft, b_bot, num, halign="center"),
        "kpi_ok": st.xf(f_kpi_ok, fl_soft, b_bot, num, halign="center"),
        "kpi_pct": st.xf(f_kpi, fl_soft, b_bot, pct, halign="center"),
    }


def band(sb: SheetBuilder, row: int, c1: int, c2: int, text: Any, style: int):
    """خلية ممتدة على عدة أعمدة — كل خلايا الامتداد تأخذ النمط ليكتمل اللون."""
    sb.set(row, c1, text, style)
    for c in range(c1 + 1, c2 + 1):
        sb.set(row, c, None, style)
    if c2 > c1:
        sb.merge(row, c1, row, c2)


HEAD = ["القيمة", "الإجمالي", "منفّذ", "جاري العمل", "متبقي", "نسبة الإنجاز"]


def _live_table(sb: SheetBuilder, s: Dict[str, int], row: int, title: str,
                data: List[Tuple[str, int, int, int]], src_range: str,
                exec_range: str) -> int:
    """جدول حيّ: الأرقام صيغ COUNTIF تتحدّث فور تعديل الورقة الرئيسية."""
    band(sb, row, 1, 6, title, s["sect"])
    sb.height(row, 22)
    row += 1
    for j, h in enumerate(HEAD):
        sb.set(row, 1 + j, h, s["th"])
    sb.height(row, 26)

    for k, (label, total, done, wip) in enumerate(data):
        row += 1
        alt = k % 2 == 1
        crit = '""' if label == "(غير محدد)" else "$B%d" % row
        sb.set(row, 1, label, s["name_b"] if alt else s["name"])
        sb.set(row, 2, total, s["td_b"] if alt else s["td"],
               formula="COUNTIFS(%s,%s)" % (src_range, crit))
        sb.set(row, 3, done, s["td_b"] if alt else s["td"],
               formula="COUNTIFS(%s,%s,%s,%s)" % (src_range, crit, exec_range, _q(ST_DONE)))
        sb.set(row, 4, wip, s["td_b"] if alt else s["td"],
               formula="COUNTIFS(%s,%s,%s,%s)" % (src_range, crit, exec_range, _q(ST_WIP)))
        sb.set(row, 5, total - done - wip, s["td_b"] if alt else s["td"],
               formula="C%d-D%d-E%d" % (row, row, row))
        sb.set(row, 6, (done / total) if total else 0.0,
               s["pct_b"] if alt else s["pct"],
               formula="IFERROR(D%d/C%d,0)" % (row, row))
        sb.height(row, 19)
    if data:
        sb.databar("G%d:G%d" % (row - len(data) + 1, row))
    return row + 2


def build_datasheet(s: Dict[str, int], an: Dict[str, Any], states: List[str],
                    contractors: Optional[List[str]], subtitle: str,
                    exec_range: str, class_range: str,
                    contractor_range: str) -> SheetBuilder:
    sb = SheetBuilder(tab_color=INK)
    for c, w in ((0, 2.5), (1, 24), (2, 13), (3, 13), (4, 14), (5, 13), (6, 17), (7, 2.5)):
        sb.width(c, w)

    total = len(states)
    done = states.count(ST_DONE)
    wip = states.count(ST_WIP)

    sb.height(1, 8)
    sb.set(1, 1, None, s["spacer"])
    band(sb, 2, 1, 6, "داتا شيت متابعة الإنجاز", s["title"])
    sb.height(2, 36)
    band(sb, 3, 1, 6, subtitle, s["sub"])
    sb.height(3, 20)
    sb.height(4, 10)
    sb.set(4, 1, None, s["spacer"])

    # شريط المؤشرات — كل رقم صيغة حيّة، فيتغيّر فور تعديل الورقة الرئيسية
    sb.set(5, 1, "إجمالي الملاحظات", s["kpi_lbl"])
    sb.set(5, 2, "منفّذة", s["kpi_lbl"])
    sb.set(5, 3, "جاري العمل", s["kpi_lbl"])
    sb.set(5, 4, "متبقية", s["kpi_lbl"])
    band(sb, 5, 5, 6, "نسبة الإنجاز", s["kpi_lbl"])

    sb.set(6, 1, total, s["kpi_num"], formula="COUNTA(%s)" % exec_range)
    sb.set(6, 2, done, s["kpi_ok"], formula="COUNTIF(%s,%s)" % (exec_range, _q(ST_DONE)))
    sb.set(6, 3, wip, s["kpi_num"], formula="COUNTIF(%s,%s)" % (exec_range, _q(ST_WIP)))
    sb.set(6, 4, total - done - wip, s["kpi_num"], formula="B6-C6-D6")
    sb.set(6, 5, (done / total) if total else 0.0, s["kpi_pct"],
           formula="IFERROR(C6/B6,0)")
    sb.set(6, 6, None, s["kpi_pct"])
    sb.merge(6, 5, 6, 6)
    sb.height(5, 20)
    sb.height(6, 34)
    sb.height(7, 14)
    sb.set(7, 1, None, s["spacer"])

    row = 8
    cls_col = an.get("class_col")
    if cls_col is not None and class_range:
        vals = [_clean(r[cls_col]) if cls_col < len(r) else "" for r in an["rows"]]
        bd = _breakdown(vals, states)
        if bd:
            row = _live_table(sb, s, row, "حسب التصنيف الرئيسي", bd,
                              class_range, exec_range)

    if contractors is not None and contractor_range:
        bd = _breakdown(contractors, states)
        if bd:
            _live_table(sb, s, row, "حسب المقاول", bd, contractor_range, exec_range)
    return sb


def build_alerts(s: Dict[str, int], missing: List[Tuple[str, str]],
                 dups: List[Tuple[str, int]]) -> SheetBuilder:
    sb = SheetBuilder()
    for c, w in ((0, 2.5), (1, 26), (2, 38), (3, 2.5)):
        sb.width(c, w)
    sb.height(1, 8)
    sb.set(1, 1, None, s["spacer"])
    band(sb, 2, 1, 2, "تنبيهات المطابقة", s["title"])
    sb.height(2, 36)
    sb.height(3, 12)
    sb.set(3, 1, None, s["spacer"])

    row = 4
    for title, head, data in (
            ("أرقام في ملفات التنفيذ ولا وجود لها في الملف الرئيسي",
             ("رقم الملاحظة", "الملف الثانوي"), missing[:2000]),
            ("أرقام مكررة داخل الملف الرئيسي",
             ("رقم الملاحظة", "رقم الصف"), dups[:2000])):
        if not data:
            continue
        band(sb, row, 1, 2, title, s["sect"])
        sb.height(row, 22)
        row += 1
        sb.set(row, 1, head[0], s["th"])
        sb.set(row, 2, head[1], s["th"])
        for k, (a, b) in enumerate(data):
            row += 1
            alt = k % 2 == 1
            sb.set(row, 1, a, s["name_b"] if alt else s["name"])
            sb.set(row, 2, b, s["name_b"] if alt else s["name"])
        row += 2
    return sb


# ------------------------------------------------------------------ المخرج

def resolve_states(an: Dict[str, Any], sec_status_mode: str) -> List[str]:
    """يحسم حالة كل صف: ما قرأناه من ملف التنفيذ أو ما اختاره المستخدم."""
    states = list(an["states"])
    for i in an["found_rows"]:
        if sec_status_mode == "done":
            states[i] = ST_DONE
        elif sec_status_mode == "wip":
            states[i] = ST_WIP
        else:                       # "column": ما لم تُذكر له حالة لم يُنفَّذ بعد
            states[i] = states[i] or ST_NONE
    return [s or ST_NONE for s in states]


def build_output(an: Dict[str, Any], rules: List[Dict[str, Any]],
                 src_path: str, out_path: str, *,
                 set_status_done: bool = False, add_source_col: bool = True,
                 move_contractor: bool = True, contractor_mode: str = "column",
                 sec_status_mode: str = "done") -> Dict[str, Any]:
    """ينسخ الملف الرئيسي كما هو ويضيف إليه الأعمدة والتلوين الحيّ والأوراق."""
    rows = an["rows"]
    nums = an["row_nums"]
    headers = an["headers"]
    n_cols = len(headers)
    states = resolve_states(an, sec_status_mode)

    cell_values: Dict[Tuple[int, int], Any] = {}
    new_columns: List[Tuple[str, Dict[int, Any]]] = []
    col_of: Dict[str, int] = {}

    def place(header: str, existing: Optional[int], values: Dict[int, Any]) -> int:
        """يكتب في عمود قائم إن وُجد، وإلا يفتح عمودًا جديدًا في آخر الملف."""
        if existing is not None:
            for rn, v in values.items():
                cell_values[(rn, existing)] = v
            return existing
        idx = n_cols + len(new_columns)
        new_columns.append((header, values))
        return idx

    col_of["exec"] = place(EXEC_HEADER, an.get("prev_exec_col"),
                           {nums[i]: states[i] for i in range(len(rows))})

    if add_source_col:
        src_vals = {nums[i]: an["sources"][i] for i in range(len(rows))
                    if an["sources"].get(i)}
        place(SOURCE_HEADER, an.get("prev_source_col"), src_vals)

    if set_status_done and an.get("status_col") is not None:
        for i, st in enumerate(states):
            if st == ST_DONE:
                cell_values[(nums[i], an["status_col"])] = DONE_STATUS

    # أسماء المقاولين تُنقل من ملفات التنفيذ إلى الملف الرئيسي
    ccol = an.get("contractor_col")
    effective: List[str] = []
    transfer: Dict[int, Any] = {}
    moved = 0
    for i, r in enumerate(rows):
        own = _clean(r[ccol]) if ccol is not None and ccol < len(r) else ""
        from_sec = an.get("sec_contractor", {}).get(i, "") if move_contractor else ""
        if not own and from_sec:
            transfer[nums[i]] = from_sec
            moved += 1
            own = from_sec
        effective.append(own)
    if ccol is not None:
        for rn, v in transfer.items():
            cell_values[(rn, ccol)] = v
        col_of["contractor"] = ccol
    elif transfer:
        col_of["contractor"] = place(CONTRACTOR_HEADER, None, transfer)

    # مدى الصيغ يمتد تحت آخر صف كي يشمل ما يضيفه المستخدم لاحقًا بيده
    first_row = an["header_row"] + 1
    last_row = (max(nums) if nums else first_row) + EXTRA_ROWS
    quoted = "'%s'!" % an["sheet"].replace("'", "''")

    def rng(idx: Optional[int]) -> str:
        if idx is None:
            return ""
        letter = xlsxedit.col_letter(idx)
        return "%s$%s$%d:$%s$%d" % (quoted, letter, first_row, letter, last_row)

    exec_letter = xlsxedit.col_letter(col_of["exec"])
    exec_range = rng(col_of["exec"])
    class_range = rng(an.get("class_col"))
    contractor_range = rng(col_of.get("contractor"))
    contractors = effective if contractor_mode != "none" else None

    subtitle = "%s · %d ملف تنفيذ · %d منفّذة و%d جاري العمل من %d" % (
        datetime.now().strftime("%Y-%m-%d"), len(an["sec_names"]),
        states.count(ST_DONE), states.count(ST_WIP), len(rows))

    def sheets_factory(st: Styles):
        s = make_styles(st)
        out = [("الداتا شيت",
                build_datasheet(s, an, states, contractors, subtitle,
                                exec_range, class_range, contractor_range))]
        if an["missing"] or an["dups"]:
            out.append(("تنبيهات", build_alerts(s, an["missing"], an["dups"])))
        return out

    total_cols = n_cols + len(new_columns)
    sqref = "A%d:%s%d" % (first_row, xlsxedit.col_letter(total_cols - 1), last_row)
    matched = set(an["sources"])          # ما ورد في ملفات التنفيذ (الآن أو سابقًا)
    fills = column_fills(rules, rows, nums, matched)

    xlsxedit.write_patched(
        src_path, out_path,
        sheet_name=an["sheet"], header_row=an["header_row"], n_cols=n_cols,
        cell_values=cell_values, new_columns=new_columns, row_fills=fills,
        cf=(sqref, state_cf_rules(rules, first_row, exec_letter)),
        validation=("%s%d:%s%d" % (exec_letter, first_row, exec_letter, last_row),
                    EXEC_STATES),
        sheets_factory=sheets_factory)

    colored_states = {r["key"] for r in rules if r.get("kind") == "state"
                      and COLOR_HEX.get(r.get("color") or "")}
    colored = len(set(fills) | {nums[i] for i, st in enumerate(states)
                                if st in colored_states})
    n_done = states.count(ST_DONE)
    n_wip = states.count(ST_WIP)
    return {
        "total": len(rows),
        "done": n_done,
        "wip": n_wip,
        "remaining": len(rows) - n_done - n_wip,
        "percent": round(100.0 * n_done / len(rows), 1) if rows else 0.0,
        "carried": an.get("carried", 0),
        "missing": len(an["missing"]),
        "dups": len(an["dups"]),
        "alerts": len(an["missing"]) + len(an["dups"]),
        "colored": colored,
        "new_cols": [t for t, _ in new_columns],
        "contractors_moved": moved,
        "stem": os.path.splitext(os.path.basename(src_path))[0],
    }
