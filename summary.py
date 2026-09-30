"""ملخص الأداء: ملفات مغذيات كثيرة ← ورقة «الملخص» واحدة.

يقرأ نوعين من الملفات لكل مغذٍّ:

* **الملف الأساسي** كما يصدّره النظام (أو كما تُخرجه أداة حالة الاعتماد):
  فيه كل الملاحظات بما فيها المقفلة، فمنه الجدولان الأول (حالة الإنجاز)
  والثاني (الإنجاز حسب إشعار D1).
* **الملف الملوّن** بعد الفرز اليدوي: حُذفت منه المقفلة ولُوّنت الصفوف
  بالألوان الأربعة، فمنه الجدول الثالث (تصنيف الملاحظات حسب اللون).

ولا يُشترط ترتيب ولا تسمية: كل ملف ملوّن يُربط بملفه الأساسي بأرقام
الملاحظات المشتركة بينهما.

والتصميم منقول عن الملخص الذي يعدّه المستخدم بيده، بفرق واحد مقصود: في
جدول D1 كان «بانتظار المعالجة» = الإجمالي − المقفل، فتُعدّ المسترجعة مرتين
(AH08: ‏200 + 11 + 215 = 426 والإجمالي 415). هنا = الإجمالي − المقفل −
المسترجع، مثل الجدول الأول.
"""
from __future__ import annotations

import datetime
import os
import re
import zipfile
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple

import builder
import engine
import matcher
import photoreview
import sorter
import xlsxedit
from xlsxedit import SheetBuilder, Styles

# ألوان الجدول الثالث بترتيب أعمدته ومعنى كل لون عند المستخدم
COLOR_COLS = [
    ("red", "تحتاج برنامج فصل", "FF0000", "FFFFFF"),
    ("green", "قابلة للعمل", "92D050", "000000"),
    ("yellow", "تحتاج مراجعة", "FFFF00", "000000"),
    ("blue", "تحويل إلى قسم آخر", "00B0F0", "000000"),
]

# ألوان تضعها أدوات الموقع نفسها (حالة الاعتماد ومتابعة المنفذ): وردي
# المسترجع وأخضر المقفل الفاتحان وغيرهما. ليست من ألوان المستخدم الأربعة،
# ولو عُدّت لصار كل مسترجع «تحتاج برنامج فصل» — قيس على ملفاته فطابق
# الأحمر في الخمسة حين استُبعدت، وخالف في ثلاثة حين عُدّت.
TOOL_PALETTE = {h.upper() for _k, _n, h in matcher.COLORS}

CLOSED_KINDS = {"ok"}
RETURNED_KINDS = {"returned"}

EXPORT_RE = re.compile(r"تاريخ\s*التصدير\s*[:：]?\s*(\d{4}-\d{2}-\d{2})")
PHONE_LIKE = re.compile(r"^05\d{8}$")

ID_ALIASES = ["رقم الملاحظة", "رقم الملاحظه", "الملاحظة رقم"]


def _clean(v: Any) -> str:
    return builder._clean(v)


def _find(headers: List[str], aliases: List[str]) -> Optional[int]:
    return photoreview._find(headers, aliases)


def short_name(contractor: str) -> str:
    """اسم مختصر مقترح للجدول: «صيانة عامة - مقاول المشعان» ← «المشعان».
    اقتراح فقط؛ المستخدم يكتب الاسم الذي يريده ويحفظه المتصفح."""
    s = (contractor or "").strip()
    s = re.sub(r"^\s*(صيانة\s+عامة|ملاحظات\s+أشجار)\s*[-–—]\s*", "", s)
    s = re.sub(r"^\s*مقاول\s+", "", s)
    return s.strip() or contractor


# -------------------------------------------------------- الملف الأساسي

def _read_rows(path: str) -> Tuple[List[str], List[Tuple[Any, ...]], str]:
    """(الترويسة، صفوف البيانات، تاريخ التصدير) من أول ورقة فيها ملاحظات."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        best = None
        for name in wb.sheetnames:
            ws = wb[name]
            hi = engine._detect_header_row(ws)
            rows = list(ws.iter_rows(values_only=True))
            if hi >= len(rows):
                continue
            headers = [_clean(c) for c in rows[hi]]
            if _find(headers, ID_ALIASES) is None:
                continue
            best = (headers, rows, hi)
            break
    finally:
        wb.close()
    if best is None:
        raise ValueError("لا توجد في الملف ورقة فيها عمود «رقم الملاحظة».")
    headers, rows, hi = best
    export = ""
    for r in rows[:hi]:
        for v in r or ():
            m = EXPORT_RE.search(str(v or ""))
            if m:
                export = m.group(1)
    idc = _find(headers, ID_ALIASES)
    data = [r for r in rows[hi + 1:] if r and idc < len(r) and _clean(r[idc])]
    return headers, data, export


def read_basic(path: str, name: str = "") -> Dict[str, Any]:
    headers, data, export = _read_rows(path)
    col = {k: _find(headers, a) for k, a in (
        ("id", ID_ALIASES),
        ("status", photoreview.STATUS_ALIASES),
        ("d1", ["إشعار D1", "اشعار D1", "إشعار d1", "اشعار d1"]),
        ("contractor", ["تمت المعالجة بواسطة", "تمت المعالجه بواسطة", "المقاول"]),
    )}
    if col["status"] is None:
        raise ValueError("الملف «%s» بلا عمود «الحالة»." % (name or os.path.basename(path)))

    def g(r, k):
        c = col[k]
        return _clean(r[c]) if c is not None and c < len(r) else ""

    closed = returned = 0
    ids = set()
    notices: Dict[str, Dict[str, Any]] = {}
    d1_by_contractor: Counter = Counter()
    for r in data:
        ids.add(matcher.norm_id(g(r, "id")))
        kind = photoreview.status_info(g(r, "status"))["kind"]
        is_closed, is_ret = kind in CLOSED_KINDS, kind in RETURNED_KINDS
        closed += is_closed
        returned += is_ret
        d1 = g(r, "d1")
        if d1:
            n = notices.setdefault(d1, {"num": d1, "total": 0, "closed": 0,
                                        "returned": 0, "contractors": Counter()})
            n["total"] += 1
            n["closed"] += is_closed
            n["returned"] += is_ret
            c = g(r, "contractor")
            n["contractors"][c or "(بلا مقاول)"] += 1
            if c:
                d1_by_contractor[c] += 1

    # رمز المغذي بالطريقة نفسها التي يسمّي بها مولّد التقرير (AH08-8911)
    try:
        recs, _m, _h = builder.read_excel(path)
        code = builder.feeder_code(recs)
    except Exception:  # noqa: BLE001
        code = ""
    if not code:
        code = os.path.splitext(name or os.path.basename(path))[0]

    main = d1_by_contractor.most_common(1)[0][0] if d1_by_contractor else ""
    nlist = sorted(notices.values(), key=lambda n: -n["total"])
    for n in nlist:
        n["phone_like"] = bool(PHONE_LIKE.match(n["num"]))
        n["contractors"] = n["contractors"].most_common()
    return {
        "name": name or os.path.basename(path),
        "code": code,
        "export": export,
        "total": len(data),
        "closed": closed,
        "returned": returned,
        "pending": len(data) - closed - returned,
        "ids": ids,
        "notices": nlist,
        "contractors": d1_by_contractor.most_common(),
        "main": main,
    }


def default_ticks(notices: List[Dict[str, Any]], contractor: str) -> List[str]:
    """الإشعارات المؤشَّرة ابتداءً: ما عليه مقاول الجدول، إلا ما يشبه رقم جوال."""
    out = []
    for n in notices:
        if n["phone_like"]:
            continue
        if any(c == contractor for c, _ in n["contractors"]):
            out.append(n["num"])
    return out


# --------------------------------------------------------- الملف الملوّن

def read_colored(path: str, name: str = "") -> Dict[str, Any]:
    """أرقام الملاحظات وعدّ الألوان الأربعة من تعبئة الصفوف."""
    an = sorter.analyze(path)
    counts = {k: 0 for k, *_ in COLOR_COLS}
    skipped = 0
    for r in an["records"]:
        hexc = (r.get("hex") or "").upper()
        key = sorter.classify(hexc) if hexc and hexc not in TOOL_PALETTE else "none"
        if key in counts:
            counts[key] += 1
        elif hexc:
            skipped += 1
    # أرقام الملاحظات للربط بالملف الأساسي
    headers, data, export = _read_rows(path)
    idc = _find(headers, ID_ALIASES)
    ids = {matcher.norm_id(_clean(r[idc])) for r in data if idc < len(r)}
    ids.discard("")
    return {
        "name": name or os.path.basename(path),
        "rows": an["total"],
        "counts": counts,
        "other": skipped,
        "uncolored": an["total"] - sum(counts.values()) - skipped,
        "ids": ids,
        "export": export,
    }


# --------------------------------------------------------------- الربط

def analyze(basics: List[Tuple[str, str]], coloreds: List[Tuple[str, str]]) -> Dict[str, Any]:
    """(مسار، اسم) لكل ملف ← مغذيات مرتّبة، كل منها بملفه الملوّن إن وُجد."""
    feeders = [read_basic(p, n) for p, n in basics]
    warnings: List[str] = []

    seen: Dict[str, str] = {}
    for f in feeders:
        if f["code"] in seen:
            raise ValueError("الملفان «%s» و«%s» للمغذي نفسه %s — ارفع ملفًا واحدًا لكل مغذٍّ."
                             % (seen[f["code"]], f["name"], f["code"]))
        seen[f["code"]] = f["name"]

    for f in feeders:
        f["colored"] = None
    for p, n in coloreds:
        c = read_colored(p, n)
        best, share = None, 0.0
        for f in feeders:
            if not c["ids"]:
                break
            s = len(c["ids"] & f["ids"]) / len(c["ids"])
            if s > share:
                best, share = f, s
        # نصف أرقام الملف الملوّن على الأقل يجب أن تكون في الأساسي، وإلا فهو
        # لمغذٍّ لم يُرفع ملفه الأساسي — ربطه بأقرب مغذٍّ يُفسد جدوله
        if best is None or share < 0.5:
            warnings.append("الملف الملوّن «%s» لا يطابق أي ملف أساسي مرفوع، فلم يُحسب."
                            % c["name"])
            continue
        if best["colored"] is not None:
            prev = best["colored"]
            if prev["share"] >= share:
                warnings.append("الملفان الملوّنان «%s» و«%s» كلاهما للمغذي %s؛ حُسب الأول."
                                % (prev["name"], c["name"], best["code"]))
                continue
            warnings.append("الملفان الملوّنان «%s» و«%s» كلاهما للمغذي %s؛ حُسب الثاني."
                            % (prev["name"], c["name"], best["code"]))
        c["share"] = share
        best["colored"] = c

    feeders.sort(key=lambda f: sorter.natural_key(f["code"]))
    dates = sorted({f["export"] for f in feeders if f["export"]})
    for f in feeders:
        f["ids"] = len(f["ids"])                 # لا حاجة للمجموعة بعد الربط
        if f["colored"]:
            f["colored"]["ids"] = len(f["colored"]["ids"])
        f["ticks"] = default_ticks(f["notices"], f["main"])
    return {"feeders": feeders, "warnings": warnings, "dates": dates}


def today_riyadh() -> str:
    """تاريخ اليوم بتوقيت السعودية (UTC+3 بلا توقيت صيفي). الخادم يعمل بتوقيت
    غرينتش، فبعد التاسعة مساءً بتوقيتنا كان سيكتب تاريخ الأمس."""
    tz = datetime.timezone(datetime.timedelta(hours=3))
    return datetime.datetime.now(tz).date().isoformat()


def date_label(dates: List[str]) -> str:
    if not dates:
        return datetime.date.today().isoformat()
    if len(dates) == 1:
        return dates[0]
    return "%s إلى %s" % (dates[0], dates[-1])


# --------------------------------------------------------------- الكتابة

def _styles(st: Styles) -> Dict[str, int]:
    b = st.border("BFBFBF")
    f = st.font

    def x(**kw):
        kw.setdefault("rtl", True)
        return st.xf(**kw)
    s: Dict[str, int] = {}
    s["title"] = x(font=f(16, True, "FFFFFF", "Arial"), fill=st.fill("1F3864"),
                   halign="center", wrap=True)
    s["sub"] = x(font=f(10, False, "595959", "Arial"), halign="center", wrap=True)
    s["sect"] = x(font=f(13, True, "1F3864", "Arial"), halign="right")
    s["th"] = x(font=f(11, True, "FFFFFF", "Arial"), fill=st.fill("1F3864"),
                border=b, halign="center", wrap=True)
    s["code"] = x(font=f(11, True, "000000", "Arial"), border=b, halign="center", wrap=True)
    s["name"] = x(font=f(11, False, "1F3864", "Arial"), fill=st.fill("FFF2CC"),
                  border=b, halign="center", wrap=True)
    s["tot"] = x(font=f(11, True, "000000", "Arial"), border=b, halign="center")
    s["ok"] = x(font=f(11, False, "006100", "Arial"), fill=st.fill("C6EFCE"),
                border=b, halign="center")
    s["ret"] = x(font=f(11, False, "9C0006", "Arial"), fill=st.fill("FFC7CE"),
                 border=b, halign="center")
    s["num"] = x(font=f(11, False, "000000", "Arial"), border=b, halign="center")
    s["pct"] = x(font=f(11, True, "000000", "Arial"), border=b, numfmt=9, halign="center")
    sum_fill = st.fill("D9E1F2")
    s["sum"] = x(font=f(11, True, "000000", "Arial"), fill=sum_fill, border=b, halign="center")
    s["sum_pct"] = x(font=f(11, True, "000000", "Arial"), fill=sum_fill, border=b,
                     numfmt=9, halign="center")
    s["note"] = x(font=f(10, False, "7F7F7F", "Arial"), border=b, halign="center", wrap=True)
    # تاريخ حقيقي لا نص: يُفرز ويُصفّى في إكسل، ويُعرض سنة-شهر-يوم في كل لغة
    s["date"] = x(font=f(11, False, "000000", "Arial"), border=b, halign="center",
                  numfmt=st.numfmt("yyyy-mm-dd"))
    for key, _label, bg, fg in COLOR_COLS:
        s["c_" + key] = x(font=f(10, True, fg, "Arial"), fill=st.fill(bg),
                          border=b, halign="center", wrap=True)
    return s


def _excel_date(text: str) -> Optional[int]:
    """«2026-02-10» ← رقم إكسل التسلسلي (أيام منذ 1899-12-30)، أو None."""
    try:
        d = datetime.datetime.strptime((text or "").strip()[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    return (d - datetime.date(1899, 12, 30)).days


def build_sheet(s: Dict[str, int], rows: List[Dict[str, Any]],
                date_text: str) -> Tuple[SheetBuilder, Dict[str, int]]:
    """الورقة كاملة. `rows` لكل مغذٍّ: الرمز، والاسم، وأرقام الجداول الثلاثة."""
    sb = SheetBuilder(rtl=True, gridlines=False, selected=True, centered=True)
    # A فاصل | B المغذي | C:D المقاول | E..I الأرقام | J تاريخ الانتهاء
    for c, w in ((0, 3), (1, 16), (2, 15), (3, 12), (4, 13), (5, 11), (6, 10),
                 (7, 13), (8, 11), (9, 15)):
        sb.width(c, w)
    last_col = 9
    sb.set(1, 1, "ملخص الملاحظات وحالة الإنجاز حسب المغذي", s["title"])
    sb.merge(1, 1, 1, last_col)
    sb.height(1, 34)
    # التاريخ آخر الجملة ومسبوق بكلمة: رقم في أول سطر عربي يقلبه محرّك الاتجاه
    sb.set(2, 1, "بيانات الإنجاز حسب ملفات المغذيات بتاريخ %s" % date_text, s["sub"])
    sb.merge(2, 1, 2, last_col)
    sb.height(2, 18)

    n = len(rows)

    def section(r0: int, title: str, heads: List[Tuple[int, str, str]]):
        sb.set(r0, 1, title, s["sect"])
        sb.merge(r0, 1, r0, last_col)
        sb.height(r0, 20)
        h = r0 + 1
        sb.height(h, 32)
        sb.set(h, 1, "المغذي", s["th"])
        sb.set(h, 2, "اسم المقاول", s["th"])
        sb.blank(h, 3, s["th"])
        sb.merge(h, 2, h, 3)
        for c, text, style in heads:
            sb.set(h, c, text, s[style])
        return h + 1

    def name_cells(r: int, row: Dict[str, Any]):
        sb.set(r, 1, row["code"], s["code"])
        sb.set(r, 2, row["name"], s["name"])
        sb.blank(r, 3, s["name"])
        sb.merge(r, 2, r, 3)

    def sum_label(r: int):
        sb.set(r, 1, "الإجمالي", s["sum"])
        sb.blank(r, 2, s["sum"])
        sb.blank(r, 3, s["sum"])
        sb.merge(r, 2, r, 3)

    L = xlsxedit.col_letter

    # ---------------------------------------- أولًا: حالة الإنجاز
    heads = [(4, "إجمالي الملاحظات", "th"), (5, "تم الإقفال", "th"), (6, "مسترجع", "th"),
             (7, "بانتظار المعالجة", "th"), (8, "نسبة الإنجاز", "th")]
    r = first1 = section(4, "أولاً: حالة الإنجاز حسب المغذي", heads)
    for row in rows:
        a = row["t1"]
        name_cells(r, row)
        sb.set(r, 4, a["total"], s["tot"], formula="SUM(F%d:H%d)" % (r, r))
        sb.set(r, 5, a["closed"], s["ok"])
        sb.set(r, 6, a["returned"], s["ret"])
        sb.set(r, 7, a["pending"], s["num"])
        sb.set(r, 8, a["closed"] / a["total"] if a["total"] else 0.0, s["pct"],
               formula="IFERROR(F%d/E%d,0)" % (r, r))
        r += 1
    last1 = r - 1
    sum_label(r)
    tot = {k: sum(x["t1"][k] for x in rows) for k in ("total", "closed", "returned", "pending")}
    for c, k in ((4, "total"), (5, "closed"), (6, "returned"), (7, "pending")):
        sb.set(r, c, tot[k], s["sum"], formula="SUM(%s%d:%s%d)" % (L(c), first1, L(c), last1))
    sb.set(r, 8, tot["closed"] / tot["total"] if tot["total"] else 0.0, s["sum_pct"],
           formula="IFERROR(F%d/E%d,0)" % (r, r))
    if n:
        sb.databar("I%d:I%d" % (first1, last1), "5B9BD5")

    # ---------------------------------------- ثانيًا: D1
    r0 = r + 2
    # بلا أقواس: القوس بعد «D1» اللاتيني يقلبه محرّك الاتجاه في بعض البرامج
    r = first2 = section(r0, "ثانياً: الإنجاز حسب إشعار D1 — الملاحظات التي عليها إشعار فقط", heads)
    for row in rows:
        a = row["t2"]
        name_cells(r, row)
        sb.set(r, 4, a["total"], s["tot"])
        sb.set(r, 5, a["closed"], s["ok"])
        sb.set(r, 6, a["returned"], s["ret"])
        sb.set(r, 7, a["total"] - a["closed"] - a["returned"], s["num"],
               formula="E%d-F%d-G%d" % (r, r, r))
        sb.set(r, 8, a["closed"] / a["total"] if a["total"] else 0.0, s["pct"],
               formula="IFERROR(F%d/E%d,0)" % (r, r))
        r += 1
    last2 = r - 1
    sum_label(r)
    t2 = {k: sum(x["t2"][k] for x in rows) for k in ("total", "closed", "returned")}
    t2["pending"] = t2["total"] - t2["closed"] - t2["returned"]
    for c, k in ((4, "total"), (5, "closed"), (6, "returned"), (7, "pending")):
        sb.set(r, c, t2[k], s["sum"], formula="SUM(%s%d:%s%d)" % (L(c), first2, L(c), last2))
    sb.set(r, 8, t2["closed"] / t2["total"] if t2["total"] else 0.0, s["sum_pct"],
           formula="IFERROR(F%d/E%d,0)" % (r, r))
    if n:
        sb.databar("I%d:I%d" % (first2, last2), "5B9BD5")

    # ---------------------------------------- ثالثًا: الألوان
    r0 = r + 2
    heads3 = [(4 + i, label, "c_" + key) for i, (key, label, _b, _f) in enumerate(COLOR_COLS)]
    heads3 += [(8, "الإجمالي", "th"), (9, "الانتهاء من الصيانة", "th")]
    r = first3 = section(r0, "ثالثاً: تصنيف الملاحظات حسب اللون", heads3)
    for row in rows:
        name_cells(r, row)
        cc = row["t3"]
        if cc is None:
            # لا نكتب أصفارًا: الصفر يقول «لا ملاحظات» والحقيقة «لا ملف»
            sb.set(r, 4, "لم يُرفع ملف ملوّن لهذا المغذي", s["note"])
            for c in (5, 6, 7, 8):
                sb.blank(r, c, s["note"])
            sb.merge(r, 4, r, 8)
        else:
            for i, (key, *_x) in enumerate(COLOR_COLS):
                sb.set(r, 4 + i, cc[key], s["num"])
            sb.set(r, 8, sum(cc.values()), s["tot"], formula="SUM(E%d:H%d)" % (r, r))
        serial = _excel_date(row["end"])
        if serial is not None:
            sb.set(r, 9, serial, s["date"])
        else:
            sb.set(r, 9, row["end"] or "لم يتم التحديد", s["num"])
        r += 1
    last3 = r - 1
    sum_label(r)
    for i, (key, *_x) in enumerate(COLOR_COLS + [("__all", "", "", "")]):
        c = 4 + i
        if key == "__all":
            v = sum(sum(x["t3"].values()) for x in rows if x["t3"])
        else:
            v = sum(x["t3"][key] for x in rows if x["t3"])
        sb.set(r, c, v, s["sum"], formula="SUM(%s%d:%s%d)" % (L(c), first3, L(c), last3))
    sb.blank(r, 9, s["sum"])
    return sb, {"first1": first1, "last1": last1, "after": r}


def compute_rows(an: Dict[str, Any], choices: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """أرقام الجداول الثلاثة لكل مغذٍّ حسب اختيارات المستخدم."""
    rows = []
    for f in an["feeders"]:
        ch = choices.get(f["code"]) or {}
        ticks = set(ch.get("ticks", f["ticks"]))
        sel = [n for n in f["notices"] if n["num"] in ticks]
        rows.append({
            "code": f["code"],
            "name": (ch.get("name") or "").strip() or short_name(ch.get("contractor") or f["main"]),
            "end": (ch.get("end") or "").strip(),
            "t1": {k: f[k] for k in ("total", "closed", "returned", "pending")},
            "t2": {"total": sum(n["total"] for n in sel),
                   "closed": sum(n["closed"] for n in sel),
                   "returned": sum(n["returned"] for n in sel),
                   "notices": [n["num"] for n in sel]},
            "t3": dict(f["colored"]["counts"]) if f.get("colored") else None,
        })
    return rows


def write_output(an: Dict[str, Any], choices: Dict[str, Dict[str, Any]],
                 out_path: str, today: str = "") -> Dict[str, Any]:
    rows = compute_rows(an, choices)
    # رأس الملخص بتاريخ اليوم الذي أُعدّ فيه، لا تاريخ تصدير الملفات
    date_text = today if re.fullmatch(r"\d{4}-\d{2}-\d{2}", today or "") else today_riyadh()
    st = Styles(xlsxedit._MIN_STYLES)
    s = _styles(st)
    sb, meta = build_sheet(s, rows, date_text)
    charts = {}
    if rows:
        # رسم «حالة الإنجاز حسب المغذي» كما في ملخص المستخدم: أشرطة متراكبة
        # من الجدول الأول نفسه، فيتحدّث إن عُدّلت أرقامه في إكسل
        f1, l1 = meta["first1"], meta["last1"]
        codes = [r["code"] for r in rows]
        series = [
            {"name": "تم الإقفال", "name_ref": "$F$%d" % (f1 - 1), "ref": "$F$%d:$F$%d" % (f1, l1),
             "values": [r["t1"]["closed"] for r in rows], "color": "70AD47"},
            {"name": "مسترجع", "name_ref": "$G$%d" % (f1 - 1), "ref": "$G$%d:$G$%d" % (f1, l1),
             "values": [r["t1"]["returned"] for r in rows], "color": "FF7C80"},
            {"name": "بانتظار المعالجة", "name_ref": "$H$%d" % (f1 - 1),
             "ref": "$H$%d:$H$%d" % (f1, l1),
             "values": [r["t1"]["pending"] for r in rows], "color": "BFBFBF"},
        ]
        xml = xlsxedit.stacked_bar_chart("حالة الإنجاز حسب المغذي", "الملخص", codes,
                                         "$B$%d:$B$%d" % (f1, l1), series)
        top = meta["after"] + 1                       # صف فارغ بعد آخر جدول (ترقيم من 0)
        height = max(14, 6 + 2 * len(rows))
        charts[0] = [(xml, (1, top, 9, top + height))]
    xlsxedit.write_workbook(out_path, [("الملخص", sb)], st, charts=charts)
    return {"rows": rows, "date": date_text, "exports": date_label(an["dates"])}
