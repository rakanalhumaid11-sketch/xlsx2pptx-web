# -*- coding: utf-8 -*-
"""
app.py
======
موقع توليد تقرير صيانة المغذي: يرفع المستخدم ملف الإكسل فقط، والقالب مخزّن
داخل الموقع (template.pptx). لا اختيار قالب، ولا ربط أعمدة، ولا تحديد عدد
شرائح: كل ذلك يُستنتج من الإكسل تلقائيًا.
"""

import contextlib
import fcntl
import json
import os
import shutil
import threading
import time
import uuid
from typing import Any, Dict

from flask import (
    Flask, request, redirect, url_for, render_template,
    send_file, jsonify, abort,
)

import afterfill
import builder

APP_DIR = os.path.dirname(os.path.abspath(__file__))
JOBS_DIR = os.environ.get("JOBS_DIR", "/tmp/xlsx2pptx_web_jobs")
os.makedirs(JOBS_DIR, exist_ok=True)

MAX_CONTENT_LENGTH = 300 * 1024 * 1024  # 300MB: ملفات الإكسل تحوي الصور بداخلها

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "xlsx2pptx-dev-secret")


# --------------------------------------------------------------------------
# إدارة "الأعمال" (jobs) على القرص
# --------------------------------------------------------------------------

def job_dir(job_id: str) -> str:
    d = os.path.join(JOBS_DIR, job_id)
    if not os.path.isdir(d):
        abort(404)
    return d


def new_job_dir():
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
    # القراءة قد تتزامن مع كتابة من خيط التوليد في الخلفية؛ الكتابة الذرية
    # أدناه تمنع قراءة ملف ناقص، لكن نعيد المحاولة بأمان تام.
    for _ in range(5):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            time.sleep(0.05)
    return {}


def write_state(d: str, state: Dict[str, Any]):
    """كتابة ذرية: ملف مؤقت ثم استبدال دفعة واحدة، كي لا تقرأ طلبات /status
    ملفًا نصفه مكتوب أثناء التوليد."""
    p = state_path(d)
    tmp = p + f".tmp{os.getpid()}"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


HEAVY_LOCK_PATH = os.path.join(JOBS_DIR, ".heavy.lock")


@contextlib.contextmanager
def heavy_lock():
    """طابور: يسمح بعملية ثقيلة واحدة فقط (توليد تقرير أو إدراج صور) في وقت
    واحد على مستوى الخادم كله. ذروة الذاكرة للعملية الواحدة تقارب 300 ميجا،
    والخادم المجاني لا يملك إلا 512، فلو تصادفت عمليتان لانهار التشغيل. قفل
    على ملف يعمل عبر كل عمّال gunicorn لأنها عمليات منفصلة."""
    f = open(HEAVY_LOCK_PATH, "w")
    try:
        fcntl.flock(f, fcntl.LOCK_EX)
        yield
    finally:
        try:
            fcntl.flock(f, fcntl.LOCK_UN)
        finally:
            f.close()


def cleanup_old_jobs(max_age_hours: float = 3.0):
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
# الصفحات
# --------------------------------------------------------------------------

@app.route("/")
def index():
    cleanup_old_jobs()
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start():
    excel = request.files.get("excel_file")
    notice = (request.form.get("notice") or "").strip()

    if not excel or not excel.filename:
        return render_template("index.html", error="الرجاء اختيار ملف الإكسل أولًا."), 400
    if not excel.filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template("index.html", error="الملف يجب أن يكون بصيغة ‎.xlsx‎."), 400

    job_id, d = new_job_dir()
    excel_path = os.path.join(d, "input.xlsx")
    excel.save(excel_path)

    state = {
        "job_id": job_id,
        "excel_path": excel_path,
        "notice": notice,
        "output_path": os.path.join(d, "report.pptx"),
        "stage": "queued",
        "progress": {"done": 0, "total": 1},
        "created_at": time.time(),
    }
    write_state(d, state)

    threading.Thread(target=_run_generation, args=(job_id,), daemon=True).start()
    return redirect(url_for("progress_page", job_id=job_id))


def _run_generation(job_id: str):
    d = os.path.join(JOBS_DIR, job_id)
    state = read_state(d)
    try:
        state["stage"] = "waiting"
        write_state(d, state)

        with heavy_lock():
            st = read_state(d)
            st["stage"] = "generating"
            write_state(d, st)

            last = [0.0]

            def progress_cb(done, total, source=None):
                # نكتب الحالة مرة كل نصف ثانية على الأكثر حتى لا نُثقل القرص
                now = time.time()
                if source is None and now - last[0] < 0.5 and done < total:
                    return
                last[0] = now
                s2 = read_state(d)
                s2["progress"] = {"done": done, "total": total}
                s2["updated_at"] = now
                if source:
                    s2["photo_source"] = source
                write_state(d, s2)

            result = builder.build_report(
                state["excel_path"], state["output_path"],
                notice=state.get("notice", ""), progress_cb=progress_cb,
            )

        # ملف الإكسل ضخم (يحوي الصور بداخله) ولم نعد بحاجته بعد بناء التقرير،
        # فنحذفه فورًا حتى لا تمتلئ مساحة الخادم عند توليد عدة تقارير.
        try:
            os.remove(state["excel_path"])
        except OSError:
            pass

        st = read_state(d)
        st["stage"] = "done"
        st["result"] = result
        st["progress"] = {"done": result["records"], "total": result["records"]}
        write_state(d, st)
    except Exception as exc:  # noqa: BLE001
        st = read_state(d)
        st["stage"] = "error"
        st["error_message"] = f"{type(exc).__name__}: {exc}"
        write_state(d, st)


@app.route("/job/<job_id>/progress")
def progress_page(job_id):
    job_dir(job_id)
    return render_template("progress.html", job_id=job_id)


@app.route("/job/<job_id>/status")
def job_status(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    return jsonify({
        "stage": state.get("stage"),
        "progress": state.get("progress", {"done": 0, "total": 1}),
        "error_message": state.get("error_message"),
        "result": state.get("result"),
        "photo_source": state.get("photo_source"),
        # ثوانٍ منذ آخر تقدّم فعلي: تستخدمها الصفحة لتنبّه المستخدم إن توقف
        # التوليد بدل أن يظل ينتظر أمام شريط ساكن
        "stalled_for": int(time.time() - state["updated_at"]) if state.get("updated_at") else 0,
    })


@app.route("/job/<job_id>/download")
def job_download(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("stage") != "done":
        abort(404)
    feeder = (state.get("result") or {}).get("feeder") or "المغذي"
    return send_file(state["output_path"], as_attachment=True,
                     download_name=f"تقرير_{feeder}.pptx")


# --------------------------------------------------------------------------
# إضافة صور «بعد المعالجة» إلى تقرير جاهز
# --------------------------------------------------------------------------

@app.route("/after", methods=["GET"])
def after_form():
    return render_template("after_new.html")


@app.route("/after/start", methods=["POST"])
def after_start():
    report = request.files.get("report_file")
    photos = [f for f in request.files.getlist("photos") if f and f.filename]

    if not report or not report.filename:
        return render_template("after_new.html", error="الرجاء اختيار ملف التقرير."), 400
    if not report.filename.lower().endswith(".pptx"):
        return render_template("after_new.html", error="ملف التقرير يجب أن يكون ‎.pptx‎."), 400
    if not photos:
        return render_template("after_new.html", error="الرجاء اختيار صور «بعد» أولًا."), 400

    job_id, d = new_job_dir()
    report_path = os.path.join(d, "report_in.pptx")
    report.save(report_path)

    raw_dir = os.path.join(d, "raw")
    os.makedirs(raw_dir, exist_ok=True)
    raw_paths = []
    for f in photos:
        p = os.path.join(raw_dir, os.path.basename(f.filename))
        f.save(p)
        raw_paths.append(p)

    images_dir = os.path.join(d, "photos")
    names = afterfill.collect_images(raw_paths, images_dir)
    shutil.rmtree(raw_dir, ignore_errors=True)
    if not names:
        shutil.rmtree(d, ignore_errors=True)
        return render_template("after_new.html",
                               error="لم يُعثر على أي صور صالحة في ما رفعته."), 400

    try:
        notes = afterfill.analyze_report(report_path)
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(d, ignore_errors=True)
        return render_template("after_new.html",
                               error=f"تعذّرت قراءة ملف التقرير: {exc}"), 400
    if not notes:
        shutil.rmtree(d, ignore_errors=True)
        return render_template(
            "after_new.html",
            error="لم يُعثر على شرائح ملاحظات في هذا الملف. تأكد أنه التقرير المولَّد من هذا الموقع."), 400

    thumbs_dir = os.path.join(d, "thumbs")
    os.makedirs(thumbs_dir, exist_ok=True)
    valid = []
    for name in names:
        try:
            afterfill.make_thumbnail(os.path.join(images_dir, name),
                                     os.path.join(thumbs_dir, name + ".jpg"))
            valid.append(name)
        except Exception:  # noqa: BLE001
            continue

    matches = afterfill.auto_match(valid, notes)
    write_state(d, {
        "job_id": job_id,
        "kind": "after",
        "report_path": report_path,
        "images_dir": images_dir,
        "thumbs_dir": thumbs_dir,
        "output_path": os.path.join(d, "report_after.pptx"),
        "notes": notes,
        "images": valid,
        "matches": matches,
        "stage": "review",
        "created_at": time.time(),
    })
    return redirect(url_for("after_review", job_id=job_id))


@app.route("/job/<job_id>/after/review", methods=["GET"])
def after_review(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "after":
        abort(404)
    notes = state["notes"]
    matches = state.get("matches", {})
    return render_template("after_review.html", job_id=job_id, notes=notes,
                           images=state["images"], matches=matches,
                           auto_count=len(matches))


@app.route("/job/<job_id>/after/thumb/<path:name>")
def after_thumb(job_id, name):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "after" or name not in state.get("images", []):
        abort(404)
    return send_file(os.path.join(state["thumbs_dir"], name + ".jpg"),
                     mimetype="image/jpeg")


@app.route("/job/<job_id>/after/apply", methods=["POST"])
def after_apply(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "after":
        abort(404)

    assignments = {}
    for i, name in enumerate(state["images"]):
        note_id = (request.form.get(f"img_{i}") or "").strip()
        if note_id:
            assignments[name] = note_id

    if not assignments:
        return redirect(url_for("after_review", job_id=job_id))

    try:
        with heavy_lock():
            result = afterfill.apply_after_photos(
                state["report_path"], state["output_path"], assignments, state["images_dir"])
    except Exception as exc:  # noqa: BLE001
        state["stage"] = "error"
        state["error_message"] = f"{type(exc).__name__}: {exc}"
        write_state(d, state)
        return render_template("after_result.html", job_id=job_id, result=None,
                               error=state["error_message"])

    state["stage"] = "done"
    state["result"] = result
    write_state(d, state)
    return render_template("after_result.html", job_id=job_id, result=result, error=None)


@app.route("/job/<job_id>/after/download")
def after_download(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "after" or state.get("stage") != "done":
        abort(404)
    return send_file(state["output_path"], as_attachment=True,
                     download_name="تقرير_الملاحظات_مع_صور_بعد.pptx")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
