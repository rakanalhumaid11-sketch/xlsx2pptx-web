"""إضافة صور «بعد المعالجة» إلى تقرير تم توليده سابقًا.

يفتح التقرير الجاهز، يتعرّف على شرائح الملاحظات ورقم كل ملاحظة، ثم يربط
الصور المرفوعة بها: تلقائيًا إذا كان اسم ملف الصورة يحوي رقم الملاحظة،
ويدويًا لما تبقّى عبر صفحة مراجعة. لا يُعاد توليد التقرير من الصفر، بل
تُضاف الصور في خانة «صورة بعد» الفارغة فقط.
"""
from __future__ import annotations

import io
import os
import re
import zipfile
from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation

import engine

NOTE_TITLE_RE = re.compile(r"معالجة\s+الملاحظة\s+رقم\s*([0-9]+)")
DIGITS_RE = re.compile(r"\d+")
IMAGE_EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif")

# الصور تُلصق داخل إطار «صورة بعد» بنفس طريقة صور «قبل» (احتواء بلا قص)
FRAME_AFTER_NAME = "PH_FRAME_AFTER"
FRAME_BEFORE_NAME = "PH_FRAME_BEFORE"


def _find_by_name(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    return None


def _note_title_shape(slide):
    """شكل عنوان الملاحظة: بالاسم أولًا، ثم بالنص كاحتياط للتقارير القديمة."""
    sh = _find_by_name(slide, "PH_NOTE_TITLE")
    if sh is not None:
        return sh
    for sh in slide.shapes:
        if sh.has_text_frame and NOTE_TITLE_RE.search(sh.text_frame.text or ""):
            return sh
    return None


def _after_box(slide, slide_width: int):
    """إطار خانة «صورة بعد». التقارير الجديدة تحمل اسمًا صريحًا؛ وإن غاب
    نبحث هندسيًا عن أكبر مستطيل فارغ في النصف الأيسر بمستوى الصور."""
    sh = _find_by_name(slide, FRAME_AFTER_NAME)
    if sh is not None:
        return sh.left, sh.top, sh.width, sh.height
    best = None
    for shp in slide.shapes:
        if shp.shape_type == 13:
            continue
        if shp.has_text_frame and shp.text_frame.text.strip():
            continue
        if shp.left is None or shp.width is None:
            continue
        if shp.left + shp.width / 2 >= slide_width / 2:
            continue
        if not (400000 <= (shp.top or 0) <= 1600000):
            continue
        area = (shp.width or 0) * (shp.height or 0)
        if best is None or area > best[0]:
            best = (area, shp.left, shp.top, shp.width, shp.height)
    if best:
        return best[1], best[2], best[3], best[4]
    return None


def _has_picture_in(slide, box) -> bool:
    if not box:
        return False
    left, top, width, height = box
    cx, cy = left + width / 2, top + height / 2
    for shp in slide.shapes:
        if shp.shape_type != 13:
            continue
        pcx = (shp.left or 0) + (shp.width or 0) / 2
        pcy = (shp.top or 0) + (shp.height or 0) / 2
        if abs(pcx - cx) < width / 2 and abs(pcy - cy) < height / 2:
            return True
    return False


def analyze_report(pptx_path: str) -> List[Dict[str, Any]]:
    """يرجع قائمة شرائح الملاحظات: رقم الشريحة ورقم الملاحظة ووصفها وهل
    خانة «بعد» مشغولة أصلًا."""
    prs = Presentation(pptx_path)
    width = prs.slide_width
    notes: List[Dict[str, Any]] = []
    for idx, slide in enumerate(prs.slides):
        title = _note_title_shape(slide)
        if title is None:
            continue
        m = NOTE_TITLE_RE.search(title.text_frame.text or "")
        note_id = m.group(1) if m else ""
        desc_shape = _find_by_name(slide, "PH_DESC")
        desc = desc_shape.text_frame.text.strip() if desc_shape is not None else ""
        box = _after_box(slide, width)
        notes.append({
            "slide": idx,
            "note_id": note_id,
            "desc": desc,
            "has_after": _has_picture_in(slide, box),
            "box_ok": box is not None,
        })
    return notes


def collect_images(paths: List[str], dest_dir: str) -> List[str]:
    """ينسخ الصور المرفوعة إلى مجلد العمل، ويفك أي ملف مضغوط فيه صور.
    يرجع أسماء الملفات المحفوظة."""
    os.makedirs(dest_dir, exist_ok=True)
    saved: List[str] = []

    def _unique(name: str) -> str:
        base = os.path.basename(name).replace("/", "_").replace("\\", "_")
        root, ext = os.path.splitext(base)
        out, i = base, 1
        while os.path.exists(os.path.join(dest_dir, out)):
            out = f"{root}_{i}{ext}"
            i += 1
        return out

    for src in paths:
        low = src.lower()
        if low.endswith(".zip"):
            try:
                with zipfile.ZipFile(src) as z:
                    for info in z.infolist():
                        if info.is_dir() or not info.filename.lower().endswith(IMAGE_EXT):
                            continue
                        if "__MACOSX" in info.filename:
                            continue
                        name = _unique(info.filename)
                        with z.open(info) as fsrc, open(os.path.join(dest_dir, name), "wb") as fdst:
                            fdst.write(fsrc.read())
                        saved.append(name)
            except zipfile.BadZipFile:
                continue
        elif low.endswith(IMAGE_EXT):
            name = _unique(src)
            with open(src, "rb") as fsrc, open(os.path.join(dest_dir, name), "wb") as fdst:
                fdst.write(fsrc.read())
            saved.append(name)
    return saved


def auto_match(image_names: List[str], notes: List[Dict[str, Any]]) -> Dict[str, str]:
    """يربط كل صورة بملاحظة إذا كان اسم ملفها يحوي رقم الملاحظة كاملًا."""
    by_id = {n["note_id"]: n for n in notes if n["note_id"]}
    used = set()
    out: Dict[str, str] = {}
    for name in image_names:
        stem = os.path.splitext(os.path.basename(name))[0]
        for group in DIGITS_RE.findall(stem):
            note = by_id.get(group)
            if note is not None and group not in used:
                out[name] = group
                used.add(group)
                break
    return out


def make_thumbnail(src_path: str, dst_path: str, width: int = 220):
    """يصنع صورة مصغّرة لعرضها في صفحة المراجعة (لتفادي تحميل صور كاملة)."""
    from PIL import Image, ImageOps
    with Image.open(src_path) as im:
        im = ImageOps.exif_transpose(im)
        ratio = width / float(im.width or 1)
        size = (width, max(1, int(im.height * ratio)))
        im = im.convert("RGB").resize(size, Image.LANCZOS)
        im.save(dst_path, format="JPEG", quality=72)


def apply_after_photos(pptx_path: str, out_path: str,
                       assignments: Dict[str, str], images_dir: str,
                       progress_cb=None) -> Dict[str, int]:
    """assignments: {اسم ملف الصورة -> رقم الملاحظة}. يضع كل صورة في خانة
    «بعد» في شريحة تلك الملاحظة ثم يحفظ نسخة جديدة من التقرير."""
    prs = Presentation(pptx_path)
    width = prs.slide_width

    slide_of: Dict[str, int] = {}
    for idx, slide in enumerate(prs.slides):
        title = _note_title_shape(slide)
        if title is None:
            continue
        m = NOTE_TITLE_RE.search(title.text_frame.text or "")
        if m:
            slide_of.setdefault(m.group(1), idx)

    added, skipped, replaced = 0, 0, 0
    items = [(img, nid) for img, nid in assignments.items() if nid]
    total = len(items) or 1
    for i, (img_name, note_id) in enumerate(items, 1):
        idx = slide_of.get(note_id)
        if idx is None:
            skipped += 1
            continue
        slide = prs.slides[idx]
        box = _after_box(slide, width)
        if not box:
            skipped += 1
            continue
        path = os.path.join(images_dir, img_name)
        try:
            with open(path, "rb") as f:
                data = f.read()
        except OSError:
            skipped += 1
            continue

        # إن كانت الخانة مشغولة بصورة سابقة نزيلها أولًا حتى لا تتراكم الصور
        left, top, w, h = box
        cx, cy = left + w / 2, top + h / 2
        for shp in list(slide.shapes):
            if shp.shape_type != 13:
                continue
            pcx = (shp.left or 0) + (shp.width or 0) / 2
            pcy = (shp.top or 0) + (shp.height or 0) / 2
            if abs(pcx - cx) < w / 2 and abs(pcy - cy) < h / 2:
                engine.remove_shape(shp)
                replaced += 1

        engine.insert_picture_in_box(slide, left, top, w, h, data)
        added += 1
        del data
        if progress_cb:
            progress_cb(i, total)

    prs.save(out_path)
    return {"added": added, "replaced": replaced, "skipped": skipped}
