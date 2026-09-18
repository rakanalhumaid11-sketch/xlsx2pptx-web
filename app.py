# -*- coding: utf-8 -*-
"""
app.py
======
موقع توليد تقرير صيانة المغذي: يرفع المستخدم ملف الإكسل فقط، والقالب مخزّن
داخل الموقع (template.pptx). لا اختيار قالب، ولا ربط أعمدة، ولا تحديد عدد
شرائح: كل ذلك يُستنتج من الإكسل تلقائيًا.
"""

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
        state["stage"] = "generating"
        write_state(d, state)

        last = [0.0]

        def progress_cb(done, total):
            # نكتب الحالة مرة كل نصف ثانية على الأكثر حتى لا نُثقل القرص
            now = time.time()
            if now - last[0] < 0.5 and done < total:
                return
            last[0] = now
            st = read_state(d)
            st["progress"] = {"done": done, "total": total}
            write_state(d, st)

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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
