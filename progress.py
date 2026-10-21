"""المنجز والمتبقي: جدول واحد لكل المغذيات، ووراءه ملف كل مغذٍّ كما هو.

المدخل ملفات المغذيات بعد أداة حالة الاعتماد وبعد تلوين المستخدم:
* الأخضر مقفل،
* الأحمر (والأزرق معه) متبقٍّ يحتاج فصل جهد متوسط MV،
* بلا لون متبقٍّ جهد منخفض يمكن العمل عليه.

المخرج ملف إكسل: الورقة الأولى جدول فيه لكل مغذٍّ الإجمالي والمنجز ونسبة
الإنجاز والمتبقي جهد متوسط والمتبقي جهد منخفض، واسم المغذي فيها رابط إلى
ورقته. وبعدها ورقة لكل مغذٍّ منسوخة كما هي (merge.py) — الملف نفسه محفوظ
داخل الملخص، والوصول إليه بضغطة.

اللون المعتمد هو الظاهر في إكسل، لأنه ما يعرّفه المستخدم ويعدّه بعينه:
أداة الاعتماد تلوّن المقفل أخضر بتنسيق شرطي فوق لون الصف الأصلي، فالصف
الأحمر الذي أُقفل يظهر أخضر ويُعدّ منجزًا. أما الوردي الذي تضعه للمسترجع
فليس من ألوان المستخدم: يُقرأ اللون الذي تحته، وهو تصنيفه هو للملاحظة.
"""
from __future__ import annotations

import datetime
import os
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook

import merge
import sorter
import summary
from xlsxedit import SheetBuilder, Styles

SHEET_TITLE = "المنجز والمتبقي"
TOOL_PALETTE = summary.TOOL_PALETTE
ID_ALIASES = ["رقم الملاحظة", "رقم الملاحظه", "الملاحظة رقم"]

# الأعمدة كما طلبها المستخدم، بلون رأس يطابق لون الصفوف في ملفاته
COLS = [
    ("total", "الإجمالي", "1F3864", "FFFFFF"),
    ("done", "المنجز", "C6EFCE", "006100"),
    ("pct", "نسبة الإنجاز", "1F3864", "FFFFFF"),
    ("mv", "المتبقي جهد متوسط (يحتاج فصل MV)", "FF0000", "FFFFFF"),
    ("lv", "المتبقي جهد منخفض (ممكن العمل عليه)", "F2F2F2", "000000"),
]
OTHER_COL = ("other", "ألوان أخرى", "FFF2CC", "7F6000")


def bucket(shown: str, static: str) -> str:
    """done / mv / lv / other لملاحظة واحدة من لونها الظاهر ولونها المخزَّن."""
    shown = (shown or "").upper()
    static = (static or "").upper()
    if sorter.classify(shown) == "green":
        return "done"
    # لون أداة الاعتماد (وردي المسترجع، أصفر بانتظار الإقفال) يغطي لون المستخدم:
    # التصنيف من اللون الذي تحته
    base = shown if shown and shown not in TOOL_PALETTE else static
    if base in TOOL_PALETTE:
        base = ""
    k = sorter.classify(base)
    if k in ("red", "blue"):
        return "mv"
    if k in ("none", "green"):
        # الأخضر تحت تنسيق آخر (مسترجع مثلًا) ليس مقفلًا: يُعدّ ممكن العمل عليه
        return "lv"
    return "other"


def read_file(path: str, name: str) -> Dict[str, Any]:
    """أعداد ملف واحد: الإجمالي والمنجز والمتبقي بنوعيه."""
    an = sorter.analyze(path)
    # الصفوف التي لها رقم ملاحظة فقط: سطر إجمالي أو ملاحظة حرّة أسفل الجدول لا يُعدّ
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[an["sheet"]]
        idc = None
        ids: Dict[int, str] = {}
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if i == an["header_row"]:
                header = [str(c or "").strip() for c in row]
                idc = next((header.index(a) for a in ID_ALIASES if a in header), None)
                if idc is None:
                    break
                continue
            if idc is None or i < an["header_row"]:
                continue
            v = row[idc] if idc < len(row) else None
            if v is not None and str(v).strip():
                ids[i] = str(v).strip()
    finally:
        wb.close()

    counts = {"total": 0, "done": 0, "mv": 0, "lv": 0, "other": 0}
    for r in an["records"]:
        if ids and r["row"] not in ids:
            continue
        counts["total"] += 1
        counts[bucket(r.get("shown", ""), r.get("hex", ""))] += 1
    return {"name": name, "path": path, **counts}


def _styles(st: Styles) -> Dict[str, int]:
    b = st.border("BFBFBF")
    f = st.font

    def x(**kw):
        kw.setdefault("rtl", True)
        return st.xf(**kw)
    s = {
        "title": x(font=f(16, True, "FFFFFF", "Arial"), fill=st.fill("1F3864"),
                   halign="center", wrap=True),
        "sub": x(font=f(10, False, "595959", "Arial"), halign="center", wrap=True),
        "code": x(font=f(11, True, "0563C1", "Arial"), border=b, halign="center"),
        "num": x(font=f(11, False, "000000", "Arial"), border=b, halign="center"),
        "pct": x(font=f(11, True, "000000", "Arial"), border=b, numfmt=9, halign="center"),
        "note": x(font=f(10, False, "7F7F7F", "Arial"), halign="right", wrap=True),
    }
    sum_fill = st.fill("D9E1F2")
    s["sum"] = x(font=f(11, True, "000000", "Arial"), fill=sum_fill, border=b, halign="center")
    s["sum_pct"] = x(font=f(11, True, "000000", "Arial"), fill=sum_fill, border=b,
                     numfmt=9, halign="center")
    s["th_code"] = x(font=f(11, True, "FFFFFF", "Arial"), fill=st.fill("1F3864"),
                     border=b, halign="center", wrap=True)
    for key, _label, bg, fg in COLS + [OTHER_COL]:
        s["th_" + key] = x(font=f(11, True, fg, "Arial"), fill=st.fill(bg),
                           border=b, halign="center", wrap=True)
    return s


def build_front(rows: List[Dict[str, Any]], date_text: str):
    """دالة تبني الورقة الأولى بعد أن يعرف الدمج أسماء أوراق المغذيات."""
    show_other = any(r["other"] for r in rows)
    cols = COLS + ([OTHER_COL] if show_other else [])

    def front(st: Styles, sources: List[Any]) -> str:
        s = _styles(st)
        sheet_of = {src.path: src.sheet_name for src in sources}
        sb = SheetBuilder(rtl=True, gridlines=False, selected=True, centered=True,
                          tab_color="1F3864")
        last = 1 + len(cols)                      # B المغذي، ثم الأعمدة
        sb.width(0, 3)
        sb.width(1, 18)
        for j, (key, *_r) in enumerate(cols):
            sb.width(2 + j, 15 if key in ("total", "done", "pct") else 22)
        sb.set(1, 1, "المنجز والمتبقي حسب المغذي", s["title"])
        sb.merge(1, 1, 1, last)
        sb.height(1, 34)
        # التاريخ آخر السطر ومسبوق بكلمة: رقم في أول سطر عربي يقلبه محرّك الاتجاه
        sb.set(2, 1, "بيانات ملفات المغذيات بتاريخ %s" % date_text, s["sub"])
        sb.merge(2, 1, 2, last)
        sb.height(2, 20)

        hr = 4
        sb.set(hr, 1, "المغذي", s["th_code"])
        for j, (key, label, *_c) in enumerate(cols):
            sb.set(hr, 2 + j, label, s["th_" + key])
        sb.height(hr, 36)

        first = hr + 1
        for i, r in enumerate(rows):
            row = first + i
            name = sheet_of.get(r["path"], r["name"])
            sb.set(row, 1, name, s["code"])
            # اسم المغذي رابط إلى ورقته في الملف نفسه
            sb.link(row, 1, "%s!A1" % merge.quote_sheet(name))
            for j, (key, *_c) in enumerate(cols):
                col = 2 + j
                L = merge.col_letter
                if key == "pct":
                    sb.set(row, col, (r["done"] / r["total"]) if r["total"] else 0.0, s["pct"],
                           formula="IFERROR(%s%d/%s%d,0)" % (L(3), row, L(2), row))
                else:
                    sb.set(row, col, r[key], s["num"])
            sb.height(row, 20)
        lastrow = first + len(rows) - 1
        tot = lastrow + 1
        sb.set(tot, 1, "الإجمالي", s["sum"])
        for j, (key, *_c) in enumerate(cols):
            col = 2 + j
            L = merge.col_letter(col)
            if key == "pct":
                t, d = sum(r["total"] for r in rows), sum(r["done"] for r in rows)
                sb.set(tot, col, (d / t) if t else 0.0, s["sum_pct"],
                       formula="IFERROR(%s%d/%s%d,0)" % (merge.col_letter(3), tot,
                                                          merge.col_letter(2), tot))
            else:
                sb.set(tot, col, sum(r[key] for r in rows), s["sum"],
                       formula="SUM(%s%d:%s%d)" % (L, first, L, lastrow))
        sb.height(tot, 22)
        pct_col = merge.col_letter(2 + [k for k, *_r in cols].index("pct"))
        sb.databar("%s%d:%s%d" % (pct_col, first, pct_col, lastrow), "5B9BD5")

        sb.set(tot + 2, 1, "الأخضر مقفل · الأحمر والأزرق يحتاج فصل جهد متوسط MV · "
               "بلا لون جهد منخفض ممكن العمل عليه. اضغط اسم المغذي لفتح ورقته.", s["note"])
        sb.merge(tot + 2, 1, tot + 2, last)
        sb.height(tot + 2, 30)
        return sb.to_xml()

    return front


def build(files: List[Tuple[str, str]], out_path: str, today: str = "") -> Dict[str, Any]:
    """files: [(المسار، الاسم)] ← ملف الملخص مع أوراق المغذيات، وأرقامه."""
    rows = [read_file(p, n) for p, n in files]
    date_text = today if _is_date(today) else summary.today_riyadh()
    # ترتيب الجدول نفسه ترتيب أوراق المغذيات في الدمج (بالرمز)
    ordered = _sorted_files(files)
    by_path = {r["path"]: r for r in rows}
    res = merge.merge(files, out_path, include_data=False,
                      front=(SHEET_TITLE, build_front([by_path[p] for p, _n in ordered],
                                                      date_text)))
    sheet_rows = []
    for (p, _n), sh in zip(ordered, res["sheets"]):
        r = dict(by_path[p])
        r["sheet"] = sh["name"]
        sheet_rows.append(r)
    warnings = list(res["warnings"])
    n_other = sum(r["other"] for r in rows)
    if n_other:
        warnings.append("%d ملاحظة بألوان غير الأخضر والأحمر والأزرق — في عمود «ألوان أخرى»."
                        % n_other)
    return {"rows": sheet_rows, "date": date_text, "warnings": warnings,
            "other": bool(n_other)}


# ترتيب المغذيات: merge.merge يرتّب الأوراق بالرمز، والجدول يتبعها
def _sorted_files(files: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    codes = {}
    for p, n in files:
        try:
            src = merge.Source(p, n)
            codes[p] = src.code or n
            src.close()
        except Exception:  # noqa: BLE001
            codes[p] = n
    return sorted(files, key=lambda f: sorter.natural_key(codes[f[0]]))


def _is_date(t: str) -> bool:
    try:
        datetime.datetime.strptime((t or "")[:10], "%Y-%m-%d")
        return len(t) == 10
    except ValueError:
        return False
