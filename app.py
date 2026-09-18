# -*- coding: utf-8 -*-
"""
app.py
======
موقع ويب يغلّف أداة "تحويل الإكسل إلى بوربوينت" (engine.py) بواجهة متصفح
بسيطة بدلاً من الويزارد النصي في الطرفية. لا حاجة لتثبيت شيء على جهاز
المستخدم؛ كل المعالجة تتم على الخادم.
"""

import json
import os
import shutil
import threading
import time
import uuid
from typing import Any, Dict, List, Optional

from flask import (
    Flask, request, redirect, url_for, render_template,
    send_file, jsonify, abort,
)
from pptx import Presentation

import engine

APP_DIR = os.path.dirname(os.path.abspath(__file__))
JOBS_DIR = os.environ.get("JOBS_DIR", "/tmp/xlsx2pptx_web_jobs")
os.makedirs(JOBS_DIR, exist_ok=True)

MAX_CONTENT_LENGTH = 200 * 1024 * 1024  # 200MB (قوالب بوربوينت قد تكون كبيرة)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "xlsx2pptx-dev-secret")

# تسميات عربية لطيفة لكل حقل (لعرضها في القوائم المنسدلة)
FIELD_LABELS: Dict[str, str] = {
    "seq": "الرقم التسلسلي (م)",
    "note_id": "رقم الملاحظة",
    "station": "المحطة",
    "feeder": "المغذي",
    "note_text": "نص الملاحظة",
    "type_code": "كود التصنيف",
    "main_category": "التصنيف الرئيسي",
    "priority": "الأولوية",
    "inspection_date": "تاريخ الفحص",
    "isolation_point": "نقطة العزل",
    "nearest_isolation": "أقرب نقطة عزل",
    "longitude": "خط الطول",
    "latitude": "خط العرض",
    "inspection_notes": "ملاحظات الفحص",
    "city": "الإدارة / المدينة",
    "office": "المكتب",
    "photo_before": "صورة قبل",
    "photo_after": "صورة بعد",
    "note_classification": "تصنيف الملاحظة",
    "note_status": "حالة الملاحظة",
    "contractor": "المقاول",
    "notice_number": "رقم إشعار الصيانة",
}
FIELD_ORDER = list(engine.FIELD_ALIASES.keys())

KIND_LABELS = {
    "skip": "تخطي (لا تفعل شيئًا)",
    "field": "نص من حقل إكسل",
    "seq_title": "عنوان برقم متسلسل",
    "coordinates": "إحداثيات (مع رابط خرائط)",
    "picture_before": "صورة — قبل",
    "picture_after": "صورة — بعد",
}


# --------------------------------------------------------------------------
# أدوات مساعدة لإدارة "الأعمال" (jobs) على القرص
# --------------------------------------------------------------------------

def job_dir(job_id: str) -> str:
    d = os.path.join(JOBS_DIR, job_id)
    if not os.path.isdir(d):
        abort(404)
    return d


def new_job_dir() -> str:
    job_id = uuid.uuid4().hex[:12]
    d = os.path.join(JOBS_DIR, job_id)
    os.makedirs(d, exist_ok=True)
    return job_id, d


def state_path(d: str) -> str:
    return os.path.join(d, "state.json")


def read_state(d: str) -> Dict[str, Any]:
    p = state_path(d)
    if not os.path.exists(p):
        return {}
    # القراءة قد تتزامن مع كتابة من خيط توليد التقرير في الخلفية؛ الكتابة
    # الذرية أدناه تمنع قراءة ملف ناقص، لكن نعيد المحاولة بأمان تام.
    for attempt in range(5):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            time.sleep(0.05)
    return {}


def write_state(d: str, state: Dict[str, Any]):
    """كتابة ذرية: نكتب لملف مؤقت ثم نستبدل الملف الأصلي دفعة واحدة (os.replace)
    كي لا تقرأ طلبات /status ملفًا نصفه مكتوب أثناء توليد التقرير في الخلفية."""
    p = state_path(d)
    tmp = p + f".tmp{os.getpid()}"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


def cleanup_old_jobs(max_age_hours: float = 12.0):
    now = time.time()
    try:
        for name in os.listdir(JOBS_DIR):
            p = os.path.join(JOBS_DIR, name)
            try:
                if os.path.isdir(p) and (now - os.path.getmtime(p)) > max_age_hours * 3600:
                    shutil.rmtree(p, ignore_errors=True)
            except OSError:
                pass
    except OSError:
        pass


# --------------------------------------------------------------------------
# الصفحة الرئيسية
# --------------------------------------------------------------------------

@app.route("/")
def index():
    cleanup_old_jobs()
    return render_template("index.html")


# --------------------------------------------------------------------------
# مسار 1: توليد تقرير جديد
# --------------------------------------------------------------------------

@app.route("/new", methods=["GET"])
def new_form():
    return render_template("new.html")


@app.route("/new/start", methods=["POST"])
def new_start():
    excel_file = request.files.get("excel_file")
    pptx_file = request.files.get("pptx_file")
    start_1based = request.form.get("start_slide", "").strip()
    end_1based = request.form.get("end_slide", "").strip()

    errors = []
    if not excel_file or not excel_file.filename:
        errors.append("لم يتم رفع ملف الإكسل.")
    if not pptx_file or not pptx_file.filename:
        errors.append("لم يتم رفع ملف قالب البوربوينت.")
    if not start_1based.isdigit() or not end_1based.isdigit():
        errors.append("رقم أول/آخر سلايد يجب أن يكون رقمًا صحيحًا.")
    if errors:
        return render_template("new.html", errors=errors)

    job_id, d = new_job_dir()
    excel_path = os.path.join(d, "input.xlsx")
    pptx_path = os.path.join(d, "template.pptx")
    excel_file.save(excel_path)
    pptx_file.save(pptx_path)

    state = {
        "stage": "uploaded",
        "excel_path": excel_path,
        "pptx_path": pptx_path,
        "start_1based": int(start_1based),
        "end_1based": int(end_1based),
    }
    write_state(d, state)

    sheets = engine.list_sheets_with_headers(excel_path)
    if len(sheets) == 1:
        state["sheet_name"] = sheets[0].name
        state["stage"] = "sheet_selected"
        write_state(d, state)
        return redirect(url_for("review", job_id=job_id))

    return render_template(
        "sheet_select.html", job_id=job_id,
        sheets=[{"name": s.name, "n_rows": s.n_rows, "n_cols": len(s.headers)} for s in sheets],
    )


@app.route("/job/<job_id>/sheet", methods=["POST"])
def choose_sheet(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    sheet_name = request.form.get("sheet_name")
    if not sheet_name:
        abort(400)
    state["sheet_name"] = sheet_name
    state["stage"] = "sheet_selected"
    write_state(d, state)
    return redirect(url_for("review", job_id=job_id))


def _analyze_for_review(state: Dict[str, Any]):
    """يحلل الإكسل والقالب ويرجع (headers, col_mapping_guess, roles_with_guess,
    master_index, sample_indices)."""
    sheets = engine.list_sheets_with_headers(state["excel_path"])
    info = next(s for s in sheets if s.name == state["sheet_name"])
    headers = info.headers
    col_mapping_guess = engine.suggest_column_mapping(headers)

    prs = Presentation(state["pptx_path"])
    start0 = state["start_1based"] - 1
    end0 = state["end_1based"] - 1
    n_slides = len(prs.slides._sldIdLst)
    start0 = max(0, min(start0, n_slides - 1))
    end0 = max(start0, min(end0, n_slides - 1))
    sample_indices = list(range(start0, end0 + 1))

    text_roles = engine.analyze_template_slides(prs, sample_indices)
    variable_text_roles = [r for r in text_roles if r.is_variable]
    master_index = engine.pick_best_master_slide(prs, sample_indices)
    constant_labels = [r for r in text_roles if not r.is_variable
                        and len(set(s for s in r.samples if s)) == 1 and r.samples[0]]
    picture_roles = engine.get_master_slide_pictures(prs, master_index, constant_labels)

    has_note_id = col_mapping_guess.get("note_id") is not None
    default_title_source = "note_id" if has_note_id else "seq"

    roles = variable_text_roles + picture_roles
    roles_out = []
    for r in roles:
        guess_kind, guess_detail = engine.guess_field_for_role(r)
        sample = next((s for s in r.samples if s), "")
        entry = {
            "geo_key": list(r.geo_key),
            "shape_type": r.shape_type,
            "nearby_label": r.nearby_label,
            "sample": sample[:80],
            "default_kind": "skip",
            "default_field_key": "note_text",
            "default_title_source_field": default_title_source,
            "default_title_template": "{seq}",
            "default_photo_field": "photo_before" if r.shape_type == "picture" else "",
        }
        if guess_kind.startswith("field:"):
            entry["default_kind"] = "field"
            entry["default_field_key"] = guess_kind.split(":", 1)[1]
        elif guess_kind == "coordinates":
            entry["default_kind"] = "coordinates"
        elif guess_kind == "seq_title":
            entry["default_kind"] = "seq_title"
            entry["default_title_template"] = guess_detail or "{seq}"
        elif guess_kind in ("picture_before", "picture_after"):
            entry["default_kind"] = guess_kind
            entry["default_photo_field"] = "photo_before" if guess_kind == "picture_before" else "photo_after"
        roles_out.append(entry)
    return headers, col_mapping_guess, roles_out, master_index, (start0, end0)


@app.route("/job/<job_id>/review", methods=["GET"])
def review(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if "sheet_name" not in state:
        return redirect(url_for("index"))

    headers, col_mapping_guess, roles_out, master_index, (start0, end0) = _analyze_for_review(state)
    state["master_index"] = master_index
    state["start0"] = start0
    state["end0"] = end0
    write_state(d, state)

    field_options = [(k, FIELD_LABELS.get(k, k)) for k in FIELD_ORDER]

    return render_template(
        "review.html", job_id=job_id, headers=list(enumerate(headers)),
        field_options=field_options, col_mapping_guess=col_mapping_guess,
        roles=roles_out, kind_labels=KIND_LABELS,
        has_note_id="note_id" in [k for k, idx in col_mapping_guess.items() if idx is not None],
    )


def _parse_review_form(form, n_roles: int) -> Dict[str, Any]:
    col_mapping = {}
    for k in FIELD_ORDER:
        v = form.get(f"col_{k}", "")
        col_mapping[k] = int(v) if v.isdigit() or (v.startswith("-") and v[1:].isdigit()) else None
        if col_mapping[k] is not None and col_mapping[k] < 0:
            col_mapping[k] = None

    decisions = []
    for i in range(n_roles):
        prefix = f"role_{i}_"
        kind = form.get(prefix + "kind", "skip")
        geo_key = json.loads(form.get(prefix + "geo_key", "[]"))
        dec: Dict[str, Any] = {"kind": kind if kind != "skip" else "skip", "geo_key": geo_key}
        if kind == "field":
            dec["kind"] = "field"
            dec["field_key"] = form.get(prefix + "field_key", "note_text")
        elif kind == "seq_title":
            dec["kind"] = "seq_title"
            dec["title_source_field"] = form.get(prefix + "title_source_field", "note_id")
            dec["title_template"] = form.get(prefix + "title_template", "{seq}")
        elif kind == "coordinates":
            dec["kind"] = "coordinates"
            dec["lat_field"] = "latitude"
            dec["lon_field"] = "longitude"
            dec["add_maps_link"] = form.get(prefix + "add_maps_link") == "on"
        elif kind in ("picture_before", "picture_after"):
            dec["kind"] = kind
            dec["photo_field"] = form.get(prefix + "photo_field",
                                           "photo_before" if kind == "picture_before" else "photo_after")
        decisions.append(dec)
    return {"col_mapping": col_mapping, "decisions": decisions}


@app.route("/job/<job_id>/review", methods=["POST"])
def review_submit(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    n_roles = int(request.form.get("n_roles", "0"))
    parsed = _parse_review_form(request.form, n_roles)
    state["col_mapping"] = parsed["col_mapping"]
    state["decisions"] = parsed["decisions"]
    state["stage"] = "generating"
    state["progress"] = {"done": 0, "total": 1}
    write_state(d, state)

    thread = threading.Thread(target=_run_generation, args=(job_id,), daemon=True)
    thread.start()
    return redirect(url_for("progress_page", job_id=job_id))


def _run_generation(job_id: str):
    d = job_dir(job_id)
    state = read_state(d)
    try:
        rows = engine.read_rows(state["excel_path"], state["sheet_name"])
        records = engine.build_records(rows, state["col_mapping"])

        def progress_cb(done, total):
            st = read_state(d)
            st["progress"] = {"done": done, "total": total}
            write_state(d, st)

        prs = Presentation(state["pptx_path"])
        urls = []
        for rec in records:
            for key in ("photo_before", "photo_after"):
                v = rec.get(key)
                if v:
                    urls.append(str(v).strip())
        image_cache = engine.prefetch_images(urls, progress_cb=progress_cb)

        result = engine.generate_report(
            prs, state["master_index"], (state["start0"], state["end0"]),
            records, state["decisions"], progress_cb=progress_cb, image_cache=image_cache,
        )

        summary_idx = engine.find_summary_slide(prs)
        if summary_idx is not None:
            engine.update_summary_totals(prs, summary_idx, len(records))

        out_name = f"report_{job_id}.pptx"
        out_path = os.path.join(d, out_name)
        prs.save(out_path)

        state = read_state(d)
        state["stage"] = "done"
        state["output_path"] = out_path
        state["result"] = result
        write_state(d, state)
    except Exception as e:
        state = read_state(d)
        state["stage"] = "error"
        state["error_message"] = str(e)
        write_state(d, state)


@app.route("/job/<job_id>/progress")
def progress_page(job_id):
    job_dir(job_id)
    return render_template("progress.html", job_id=job_id)


@app.route("/job/<job_id>/status")
def job_status(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    out = {
        "stage": state.get("stage"),
        "progress": state.get("progress", {"done": 0, "total": 1}),
        "error_message": state.get("error_message"),
        "result": state.get("result"),
    }
    return jsonify(out)


@app.route("/job/<job_id>/download")
def job_download(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("stage") != "done":
        abort(404)
    return send_file(state["output_path"], as_attachment=True,
                      download_name="تقرير_الملاحظات.pptx")


# --------------------------------------------------------------------------
# مسار 2: إضافة صور "بعد المعالجة" لاحقًا
# --------------------------------------------------------------------------

@app.route("/after", methods=["GET"])
def after_form():
    return render_template("after_new.html")


@app.route("/after/start", methods=["POST"])
def after_start():
    pptx_file = request.files.get("pptx_file")
    start_1based = request.form.get("start_slide", "").strip()
    end_1based = request.form.get("end_slide", "").strip()

    errors = []
    if not pptx_file or not pptx_file.filename:
        errors.append("لم يتم رفع ملف البوربوينت.")
    if not start_1based.isdigit() or not end_1based.isdigit():
        errors.append("رقم أول/آخر سلايد يجب أن يكون رقمًا صحيحًا.")
    if errors:
        return render_template("after_new.html", errors=errors)

    job_id, d = new_job_dir()
    pptx_path = os.path.join(d, "existing.pptx")
    pptx_file.save(pptx_path)

    prs = Presentation(pptx_path)
    n_slides = len(prs.slides._sldIdLst)
    start0 = max(0, min(int(start_1based) - 1, n_slides - 1))
    end0 = max(start0, min(int(end_1based) - 1, n_slides - 1))
    indices = list(range(start0, end0 + 1))

    boxes = engine.find_picture_boxes(prs, indices)
    if not boxes:
        shutil.rmtree(d, ignore_errors=True)
        return render_template(
            "after_new.html",
            errors=["لم أستطع إيجاد خانتي صورة \"قبل\" و\"بعد\" تلقائيًا ضمن هذا النطاق. "
                    "تأكد من رقم أول وآخر سلايد ملاحظة في الملف."])

    missing = engine.slides_missing_photo(prs, indices, boxes["after"])
    previews = [{"slide_index": i, "text": engine.slide_title_preview(prs.slides[i])} for i in missing]

    state = {
        "stage": "missing_listed",
        "pptx_path": pptx_path,
        "start0": start0, "end0": end0,
        "after_box": list(boxes["after"]),
        "missing": missing,
    }
    write_state(d, state)

    if not missing:
        return render_template("after_review.html", job_id=job_id, previews=[], none_missing=True)

    return render_template("after_review.html", job_id=job_id, previews=previews, none_missing=False)


@app.route("/job/<job_id>/after/apply", methods=["POST"])
def after_apply(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    prs = Presentation(state["pptx_path"])
    after_box = tuple(state["after_box"])

    applied = 0
    for idx in state["missing"]:
        url_val = request.form.get(f"url_{idx}", "").strip()
        file_val = request.files.get(f"file_{idx}")
        img_bytes = None
        if file_val and file_val.filename:
            img_bytes = file_val.read()
        elif url_val:
            img_bytes = engine.load_image_bytes(url_val)
        if img_bytes:
            slide = prs.slides[idx]
            engine.insert_picture_in_box(slide, after_box[0], after_box[1], after_box[2], after_box[3], img_bytes)
            applied += 1

    indices = list(range(state["start0"], state["end0"] + 1))
    resolved_count = engine.count_pictures_near_box(prs, indices, after_box)
    summary_idx = engine.find_summary_slide(prs)
    if summary_idx is not None:
        for kw in ("الملاحظات المعالجة", "المعالجة", "منجز"):
            if engine.update_summary_totals(prs, summary_idx, resolved_count, total_caption_kw=kw):
                break

    out_name = f"after_{job_id}.pptx"
    out_path = os.path.join(d, out_name)
    prs.save(out_path)

    state["stage"] = "done"
    state["output_path"] = out_path
    state["applied"] = applied
    state["resolved_count"] = resolved_count
    write_state(d, state)

    return render_template("after_result.html", job_id=job_id, applied=applied, resolved_count=resolved_count)


@app.route("/job/<job_id>/after/download")
def after_download(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("stage") != "done":
        abort(404)
    return send_file(state["output_path"], as_attachment=True,
                      download_name="تقرير_محدث_بالصور.pptx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
