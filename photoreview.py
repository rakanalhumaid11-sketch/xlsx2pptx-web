"""مراجعة صور «قبل» و«بعد» بالعين، صفًّا بعد صف.

المشكلة: مولّد التقرير يأخذ «صورة قبل 1» و«صورة بعد 1» دائمًا، وهما كثيرًا
ما لا يقارَنان — «قبل» لقطة عامة للعمود من عشرين مترًا و«بعد» لقطة قريبة
لقاعدته. وفي الصف عادةً ثلاث صور «بعد» أو أكثر، فيها غالبًا واحدة من زاوية
تقارب «قبل» وتُظهر المعالجة.

الحل هنا عين المستخدم لا حكم آلي: تُعرض الملاحظات كلها في صفحة واحدة
للتمرير السريع، وقد اختيرت الصور مبدئيًا بأعلى تشابه بصري، فلا يتدخّل إلا
حيث يرى خطأ. والمخرج نسخة من الإكسل أُعيد فيها ترتيب أعمدة الصور، فتصير
«صورة قبل 1» و«صورة بعد 1» هما ما اختاره — ويستفيد منه مولّد التقرير وكل
أداة أخرى بلا تغيير.

التصغير شرط لا تحسين: المستضيف لا يعطي نسخة مصغّرة (جُرِّبت كل الصيغ)، فكل
صورة نحو 1.1 ميجا — ألف ومئتا صورة لستمئة ملاحظة تعني 1.3 جيجا في المتصفح.
بعد التصغير نحو أربعين كيلوبايت، أي خمسين ميجا للملف كله.
"""
from __future__ import annotations

import concurrent.futures as cf
import hashlib
import io
import os
import re
import urllib.request
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from PIL import Image, ImageOps

import engine
import xlsxedit

# مقاس المصغّرة: يكفي للحكم على الزاوية والمعالجة في التمرير، ويصغّر الحجم
# خمسًا وعشرين مرة. والعرض في الصفحة أصغر منه، فتبقى واضحة عند التكبير.
THUMB = int(os.environ.get("REVIEW_THUMB", "520"))
THUMB_Q = int(os.environ.get("REVIEW_THUMB_Q", "72"))
DL_WORKERS = int(os.environ.get("REVIEW_WORKERS", "12"))
DL_TIMEOUT = 30
UA = "Mozilla/5.0 (compatible; feeder-report/1.0)"

BEFORE_PAT = re.compile(r"^صور[ةه]\s*قبل\s*(\d+)$")
AFTER_PAT = re.compile(r"^صور[ةه]\s*بعد\s*(\d+)$")
PLAIN_PAT = re.compile(r"^(?:ال)?صور[ةه]\s*(\d+)$")

ID_ALIASES = ["رقم الملاحظة", "رقم الملاحظه", "الملاحظة رقم"]
TYPE_ALIASES = ["الملاحظة", "الملاحظه", "وصف الملاحظة"]
INSPECT_ALIASES = ["ملاحظات الفحص", "ملاحظات الفحص الإضافية"]
MAINT_ALIASES = ["ملاحظات الصيانة", "ملاحظات الصيانه"]
CONTRACTOR_ALIASES = ["تمت المعالجة بواسطة", "تمت المعالجه بواسطة", "المقاول"]


def _norm(s: Any) -> str:
    return engine.normalize_ar(str(s or "").strip())


def _find(headers: List[str], aliases: List[str]) -> Optional[int]:
    want = [_norm(a) for a in aliases]
    for i, h in enumerate(headers):
        if _norm(h) in want:
            return i
    for i, h in enumerate(headers):
        nh = _norm(h)
        if any(w and nh.startswith(w) for w in want):
            return i
    return None


def key_of(url: str) -> str:
    return hashlib.md5(url.strip().encode()).hexdigest()


# --------------------------------------------------------------- التحليل

def analyze(path: str) -> Dict[str, Any]:
    """الملاحظات وصورها من الإكسل، دون تنزيل شيء."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet_name = wb.sheetnames[0]
        ws = wb[sheet_name]
        hdr_i = engine._detect_header_row(ws)
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if hdr_i >= len(rows):
        raise ValueError("لم يُعثر على صف عناوين في الملف.")
    headers = [str(c).strip() if c is not None else "" for c in rows[hdr_i]]

    before_cols: List[Tuple[int, int]] = []
    after_cols: List[Tuple[int, int]] = []
    plain_cols: List[Tuple[int, int]] = []
    for i, h in enumerate(headers):
        t = re.sub(r"\s+", " ", str(h or "").strip())
        m = BEFORE_PAT.match(t)
        if m:
            before_cols.append((int(m.group(1)), i))
            continue
        m = AFTER_PAT.match(t)
        if m:
            after_cols.append((int(m.group(1)), i))
            continue
        m = PLAIN_PAT.match(t)
        if m:
            plain_cols.append((int(m.group(1)), i))
    # الملفات القديمة تسمّي أعمدة «قبل» بـ«صورة 1…» بلا كلمة قبل
    if not before_cols and plain_cols:
        before_cols = plain_cols
    before = [c for _, c in sorted(before_cols)]
    after = [c for _, c in sorted(after_cols)]
    if not before:
        raise ValueError("لم يُعثر على أعمدة صور في الملف.")

    id_c = _find(headers, ID_ALIASES)
    ty_c = _find(headers, TYPE_ALIASES)
    in_c = _find(headers, INSPECT_ALIASES)
    mt_c = _find(headers, MAINT_ALIASES)
    ct_c = _find(headers, CONTRACTOR_ALIASES)

    notes: List[Dict[str, Any]] = []
    for i, row in enumerate(rows):
        if i <= hdr_i:
            continue
        if not any(c is not None and str(c).strip() for c in row):
            continue
        def g(c):
            return str(row[c]).strip() if c is not None and c < len(row) and row[c] is not None else ""
        def urls(cols):
            out = []
            for c in cols:
                v = g(c)
                if v.lower().startswith("http") and v not in out:
                    out.append(v)
            return out
        b, a = urls(before), urls(after)
        if not b and not a:
            continue
        notes.append({
            "row": i + 1,                       # رقم الصف في إكسل
            "id": g(id_c), "type": g(ty_c),
            "inspect": g(in_c), "maint": g(mt_c), "contractor": g(ct_c),
            "before": b, "after": a,
        })
    return {
        "sheet": sheet_name,
        "header_row": hdr_i + 1,
        "headers": headers,
        "n_cols": len(headers),
        "before_cols": before,
        "after_cols": after,
        "notes": notes,
        "total": len(notes),
        "with_after": sum(1 for n in notes if n["after"]),
        "n_images": sum(len(n["before"]) + len(n["after"]) for n in notes),
    }


# ------------------------------------------------------------- التجهيز

def _fetch(url: str) -> Optional[bytes]:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        with urllib.request.urlopen(req, timeout=DL_TIMEOUT) as r:
            return r.read()
    except Exception:  # noqa: BLE001
        return None


def _thumb(raw: bytes, dst: str) -> bool:
    try:
        im = ImageOps.exif_transpose(Image.open(io.BytesIO(raw)))
        im = im.convert("RGB")
        im.thumbnail((THUMB, THUMB), Image.LANCZOS)
        im.save(dst, "JPEG", quality=THUMB_Q, optimize=True)
        return True
    except Exception:  # noqa: BLE001
        return False


def prepare(an: Dict[str, Any], thumb_dir: str, progress_cb=None,
            rank: bool = True) -> Dict[str, Any]:
    """ينزّل صور كل الملاحظات ويصغّرها، ثم يرشّح أفضل زوج لكل ملاحظة."""
    os.makedirs(thumb_dir, exist_ok=True)
    urls: List[str] = []
    seen = set()
    for n in an["notes"]:
        for u in n["before"] + n["after"]:
            if u not in seen:
                seen.add(u)
                urls.append(u)

    done = [0]
    total = len(urls)

    def one(u: str) -> Tuple[str, bool]:
        dst = os.path.join(thumb_dir, key_of(u) + ".jpg")
        ok = os.path.exists(dst)
        if not ok:
            raw = _fetch(u)
            ok = bool(raw) and _thumb(raw, dst)
        done[0] += 1
        if progress_cb and done[0] % 5 == 0:
            progress_cb(done[0], total, "download")
        return u, ok

    ok_map: Dict[str, bool] = {}
    with cf.ThreadPoolExecutor(DL_WORKERS) as ex:
        for u, ok in ex.map(one, urls):
            ok_map[u] = ok
    if progress_cb:
        progress_cb(total, total, "download")

    choices: Dict[int, Dict[str, Any]] = {}
    if rank:
        choices = _rank_all(an, thumb_dir, ok_map, progress_cb)
    else:
        for n in an["notes"]:
            b = next((u for u in n["before"] if ok_map.get(u)), "")
            a = next((u for u in n["after"] if ok_map.get(u)), "")
            choices[n["row"]] = {"before": b, "after": a, "score": 0}

    return {
        "images": total,
        "fetched": sum(1 for v in ok_map.values() if v),
        "failed": [u for u, v in ok_map.items() if not v],
        "choices": choices,
        "bytes": sum(os.path.getsize(os.path.join(thumb_dir, f))
                     for f in os.listdir(thumb_dir)),
    }


def _rank_all(an: Dict[str, Any], thumb_dir: str, ok_map: Dict[str, bool],
              progress_cb=None) -> Dict[int, Dict[str, Any]]:
    """يرشّح لكل ملاحظة الزوج الأقرب زاويةً بمطابقة المعالم البصرية.

    المطابقة المطلقة («هل الصورتان لموقع واحد») لا تعمل على هذه الصور —
    قِست فكان وسيط الأزواج الحقيقية 6 نقاط والمخلوطة 5. لكن الترشيح
    *داخل الصف* مسألة أسهل: ثلاثة مرشّحين نرتّبهم، والأعلى غالبًا هو
    الذي صُوّر من زاوية «قبل». والمستخدم يصحّح بضغطة."""
    try:
        import cv2
        import numpy as np
    except Exception:  # noqa: BLE001
        return {n["row"]: {"before": next((u for u in n["before"] if ok_map.get(u)), ""),
                           "after": next((u for u in n["after"] if ok_map.get(u)), ""),
                           "score": 0} for n in an["notes"]}

    sift = cv2.SIFT_create(nfeatures=800)
    bf = cv2.BFMatcher()
    feats: Dict[str, Any] = {}

    def f(u: str):
        if u in feats:
            return feats[u]
        p = os.path.join(thumb_dir, key_of(u) + ".jpg")
        im = cv2.imread(p, cv2.IMREAD_GRAYSCALE) if os.path.exists(p) else None
        feats[u] = sift.detectAndCompute(im, None) if im is not None else None
        return feats[u]

    def score(u1: str, u2: str) -> int:
        a, b = f(u1), f(u2)
        if not a or not b or a[1] is None or b[1] is None:
            return 0
        k1, d1 = a
        k2, d2 = b
        if len(k1) < 8 or len(k2) < 8:
            return 0
        good = [m for m, n in bf.knnMatch(d1, d2, k=2) if m.distance < 0.75 * n.distance]
        if len(good) < 8:
            return len(good)
        src = np.float32([k1[m.queryIdx].pt for m in good]).reshape(-1, 1, 2)
        dst = np.float32([k2[m.trainIdx].pt for m in good]).reshape(-1, 1, 2)
        _M, mask = cv2.findHomography(src, dst, cv2.RANSAC, 6.0)
        return int(mask.sum()) if mask is not None else 0

    out: Dict[int, Dict[str, Any]] = {}
    notes = an["notes"]
    for i, n in enumerate(notes):
        bs = [u for u in n["before"] if ok_map.get(u)]
        as_ = [u for u in n["after"] if ok_map.get(u)]
        if progress_cb and i % 5 == 0:
            progress_cb(i, len(notes), "rank")
        if not bs or not as_:
            out[n["row"]] = {"before": bs[0] if bs else "",
                             "after": as_[0] if as_ else "", "score": 0}
            continue
        best = (-1, bs[0], as_[0])
        for b in bs:
            for a in as_:
                s = score(b, a)
                if s > best[0]:
                    best = (s, b, a)
        out[n["row"]] = {"before": best[1], "after": best[2], "score": best[0]}
        # الصف الواحد قد يحمل عشر صور؛ إفراغ الذاكرة يمنع تضخّمها على خادم صغير
        for u in bs + as_:
            feats.pop(u, None)
    if progress_cb:
        progress_cb(len(notes), len(notes), "rank")
    return out


# --------------------------------------------------------------- المخرج

# عتبة «يصلح للمقارنة»: قِيست على ملف فعلي — دونها لم يصوّر الفريق صورة
# «بعد» من زاوية تقارب «قبل» أصلًا، فلا شيء يُقارَن مهما نظرت
COMPARABLE = int(os.environ.get("REVIEW_COMPARABLE", "12"))


def write_output(an: Dict[str, Any], choices: Dict[int, Dict[str, Any]],
                 src_path: str, out_path: str) -> Dict[str, Any]:
    """نسخة من الملف أُعيد فيها ترتيب أعمدة الصور حسب اختيار المستخدم.

    لا تُحذف صورة ولا تُضاف: تُبادَل مواضعها داخل أعمدتها فقط، فتصير
    المختارة في «صورة قبل 1» و«صورة بعد 1»."""
    cell_values: Dict[Tuple[int, int], Any] = {}
    moved = 0
    by_row = {n["row"]: n for n in an["notes"]}
    for row, ch in choices.items():
        n = by_row.get(row)
        if not n:
            continue
        for kind, cols in (("before", an["before_cols"]), ("after", an["after_cols"])):
            urls = list(n[kind])
            pick = ch.get(kind) or ""
            if not urls or not pick or pick not in urls:
                continue
            if urls[0] == pick:
                continue
            j = urls.index(pick)
            urls[0], urls[j] = urls[j], urls[0]
            moved += 1
            for k, c in enumerate(cols):
                cell_values[(row, c)] = urls[k] if k < len(urls) else None

    xlsxedit.write_patched(
        src_path, out_path,
        sheet_name=an["sheet"], header_row=an["header_row"],
        n_cols=an["n_cols"], cell_values=cell_values, new_columns=[])

    comparable = sum(1 for r, c in choices.items()
                     if by_row.get(r) and by_row[r]["after"]
                     and c.get("score", 0) >= COMPARABLE)
    with_after = sum(1 for r in choices if by_row.get(r) and by_row[r]["after"])
    return {
        "notes": len(by_row),
        "with_after": with_after,
        "no_after": len(by_row) - with_after,
        "reordered": moved,
        "comparable": comparable,
        "not_comparable": with_after - comparable,
        "stem": os.path.splitext(os.path.basename(src_path))[0],
    }
