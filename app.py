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
import pickle
import shutil
import threading
import time
import uuid
from typing import Any, Dict

from flask import (
    Flask, request, redirect, url_for, render_template,
    send_file, jsonify, abort, Response,
)

import afterfill
import approval
import builder
import matcher
import sorter
import zones

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
    result = state.get("result") or {}
    feeder = result.get("feeder") or "المغذي"
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


# --------------------------------------------------------------------------
# تقسيم الملاحظات إلى زونات عمل
# --------------------------------------------------------------------------

@app.route("/zones", methods=["GET"])
def zones_form():
    return render_template("zones_new.html", max_zones=zones.MAX_ZONES)


@app.route("/zones/start", methods=["POST"])
def zones_start():
    excel = request.files.get("excel_file")
    k = max(1, min(_int_ar(request.form.get("zones"), 4), zones.MAX_ZONES))

    if not excel or not excel.filename:
        return render_template("zones_new.html", max_zones=zones.MAX_ZONES,
                               error="الرجاء اختيار ملف الإكسل أولًا."), 400
    if not excel.filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template("zones_new.html", max_zones=zones.MAX_ZONES,
                               error="الملف يجب أن يكون بصيغة ‎.xlsx‎."), 400

    job_id, d = new_job_dir()
    excel_path = os.path.join(d, "input.xlsx")
    excel.save(excel_path)

    try:
        with heavy_lock():
            records, _mapping, _headers = builder.read_excel(excel_path)
            points, skipped = zones.extract_points(records)
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(d, ignore_errors=True)
        return render_template("zones_new.html", max_zones=zones.MAX_ZONES,
                               error=f"تعذّرت قراءة الملف: {exc}"), 400
    finally:
        # لم نعد بحاجة للإكسل: النقاط وحدها تكفي لإعادة التقسيم بأي عدد
        try:
            os.remove(excel_path)
        except OSError:
            pass

    if not points:
        shutil.rmtree(d, ignore_errors=True)
        return render_template("zones_new.html", max_zones=zones.MAX_ZONES,
                               error="لا توجد إحداثيات صالحة في هذا الملف."), 400

    seen, contractors = set(), []
    for p in points:
        c = (p.get("contractor") or "").strip()
        if c and c not in seen:
            seen.add(c)
            contractors.append(c)

    write_state(d, {
        "job_id": job_id,
        "kind": "zones",
        "points": points,
        "skipped": skipped,
        "feeder": builder.feeder_code(records),
        "contractors": contractors[:20],
        "zones": k,
        "created_at": time.time(),
    })
    return redirect(url_for("zones_map", job_id=job_id, k=k))


def _int_ar(value, default=0):
    """يقرأ رقمًا مكتوبًا بالأرقام العربية (٥) أو اللاتينية (5) على السواء."""
    if value is None:
        return default
    s = str(value).strip()
    if not s:
        return default
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩" + "۰۱۲۳۴۵۶۷۸۹", "0123456789" * 2)
    try:
        return int(s.translate(trans))
    except ValueError:
        return default


def _zones_state(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "zones":
        abort(404)
    k = _int_ar(request.values.get("k"), state.get("zones") or 4)
    k = max(1, min(k, zones.MAX_ZONES, len(state["points"])))
    try:
        slack = float(request.values.get("slack", state.get("slack", 0)))
    except ValueError:
        slack = 0.0
    slack = max(0.0, min(1.0, slack))
    return state, k, slack


def _kml_response(kml_text, filename_stem):
    """يرسل ملف KML باسم عربي صحيح (ترميز RFC 5987 حتى لا يُشوَّه الاسم)."""
    from urllib.parse import quote
    name = quote(f"{filename_stem}.kml".replace("/", "-"), safe="")
    return Response(
        kml_text, mimetype="application/vnd.google-earth.kml+xml",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{name}"})


def _built_zones(state, k, slack):
    caps = {int(z): int(v) for z, v in (state.get("caps") or {}).items() if int(v) > 0}
    return zones.build_zones(state["points"], k, slack=slack,
                             overrides=state.get("overrides") or {},
                             caps=caps, names=state.get("names") or {})


@app.route("/job/<job_id>/zones")
def zones_map(job_id):
    state, k, slack = _zones_state(job_id)
    points = state["points"]
    n_isolated = zones.mark_isolation(points)
    built = _built_zones(state, k, slack)
    return render_template("zones_map.html", job_id=job_id, k=k, slack=slack,
                           zones=built, feeder=state.get("feeder", ""),
                           skipped=state.get("skipped", 0),
                           total=len(points), n_isolated=n_isolated,
                           iso_km=zones.ISOLATED_KM,
                           contractors=state.get("contractors") or [],
                           n_overrides=len(state.get("overrides") or {}),
                           max_zones=min(zones.MAX_ZONES, len(points)))


@app.route("/job/<job_id>/zones/settings", methods=["POST"])
def zones_settings(job_id):
    """حفظ اسم المقاول وعدد الملاحظات المطلوب لكل زون."""
    d = job_dir(job_id)
    state, k, slack = _zones_state(job_id)
    names, caps = {}, {}
    for i in range(1, k + 1):
        nm = (request.form.get(f"name_{i}") or "").strip()
        if nm:
            names[str(i)] = nm[:40]
        cap = _int_ar(request.form.get(f"cap_{i}"), 0)
        if cap > 0:
            caps[str(i)] = cap
    state["names"] = names
    state["caps"] = caps
    state["zones"] = k
    state["slack"] = slack
    write_state(d, state)
    return redirect(url_for("zones_map", job_id=job_id, k=k, slack=slack))


@app.route("/job/<job_id>/zones/move", methods=["POST"])
def zones_move(job_id):
    """نقل ملاحظة يدويًا إلى زون آخر — يُحفظ ويغلب على نتيجة الخوارزمية."""
    d = job_dir(job_id)
    state, k, slack = _zones_state(job_id)
    note_id = (request.form.get("note_id") or "").strip()
    reset = request.form.get("reset")

    overrides = state.get("overrides") or {}
    if reset:
        overrides = {}
    elif note_id:
        z = _int_ar(request.form.get("zone"), 0)
        if z <= 0:
            overrides.pop(note_id, None)      # إرجاعها لتقدير الخوارزمية
        else:
            overrides[note_id] = min(z, k)

    state["overrides"] = overrides
    state["zones"] = k
    state["slack"] = slack
    write_state(d, state)
    return redirect(url_for("zones_map", job_id=job_id, k=k, slack=slack))


@app.route("/job/<job_id>/zones/kml/<int:zone_no>")
def zones_kml_one(job_id, zone_no):
    """ملف KML لزون واحد بنفس لونه في الملف الشامل — يُرسل لفريقه وحده."""
    state, k, slack = _zones_state(job_id)
    zones.mark_isolation(state["points"])
    built = _built_zones(state, k, slack)
    one = [z for z in built if z["index"] == zone_no]
    if not one:
        abort(404)
    z = one[0]
    title = f'زون {z["index"]} — {z["color_name"]}'
    if z.get("name"):
        title += f' — {z["name"]}'
    return _kml_response(zones.build_kml(one, title), title)


@app.route("/job/<job_id>/zones/kml")
def zones_kml(job_id):
    state, k, slack = _zones_state(job_id)
    zones.mark_isolation(state["points"])
    built = _built_zones(state, k, slack)
    feeder = state.get("feeder") or "المغذي"
    title = f"زونات {feeder}"
    return _kml_response(zones.build_kml(built, title), title)


# --------------------------------------------------------------------------
# مطابقة الأعمال المنفّذة: ملف رئيسي + ملفات تنفيذ ثانوية
# --------------------------------------------------------------------------

def _match_state(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "match":
        abort(404)
    return d, state


def _load_analysis(d):
    with open(os.path.join(d, "analysis.pkl"), "rb") as f:
        return pickle.load(f)


@app.route("/match", methods=["GET"])
def match_form():
    cleanup_old_jobs()
    return render_template("match_new.html")


@app.route("/match/start", methods=["POST"])
def match_start():
    main = request.files.get("main_file")
    secs = [f for f in request.files.getlist("sec_files") if f and f.filename]

    def fail(msg):
        return render_template("match_new.html", error=msg), 400

    if not main or not main.filename:
        return fail("الرجاء اختيار الملف الرئيسي أولًا.")
    if not main.filename.lower().endswith((".xlsx", ".xlsm")):
        return fail("الملف الرئيسي يجب أن يكون بصيغة ‎.xlsx‎.")
    if not secs:
        return fail("الرجاء اختيار ملف تنفيذ واحد على الأقل.")
    bad = [f.filename for f in secs if not f.filename.lower().endswith((".xlsx", ".xlsm"))]
    if bad:
        return fail(f"ملفات التنفيذ يجب أن تكون ‎.xlsx‎ — تحقّق من: {bad[0]}")

    job_id, d = new_job_dir()
    main_path = os.path.join(d, "main.xlsx")
    main.save(main_path)

    sec_dir = os.path.join(d, "sec")
    os.makedirs(sec_dir, exist_ok=True)
    sec_paths = []
    for f in secs:
        p = os.path.join(sec_dir, os.path.basename(f.filename))
        f.save(p)
        sec_paths.append(p)

    try:
        with heavy_lock():
            an = matcher.analyze(main_path, sec_paths)
    except matcher.NoIdColumn as exc:
        # صمت الأداة عن ملف لا عمود رقم فيه كان يُنقص النتيجة دون أن يظهر
        shutil.rmtree(d, ignore_errors=True)
        return fail(
            f"ملف التنفيذ «{exc}» لا يحتوي على عمود رقم ملاحظة يمكن التعرّف "
            "عليه، فلن يُحتسب منه شيء. سمِّ العمود «رقم الملاحظة» ثم أعد رفعه.")
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(d, ignore_errors=True)
        return fail(f"تعذّرت قراءة الملفات: {exc}")

    if not an["rows"]:
        shutil.rmtree(d, ignore_errors=True)
        return fail("الملف الرئيسي لا يحتوي على صفوف بيانات.")

    with open(os.path.join(d, "analysis.pkl"), "wb") as f:
        pickle.dump(an, f, protocol=pickle.HIGHEST_PROTOCOL)

    stem = os.path.splitext(os.path.basename(main.filename))[0][:60] or "الملاحظات"
    write_state(d, {
        "job_id": job_id,
        "kind": "match",
        "stem": stem,
        "main_path": main_path,
        "sec_names": an["sec_names"],
        "created_at": time.time(),
    })
    # الملف الرئيسي يبقى: المخرج نسخة منه بالضبط لا ملف مبنيّ من جديد
    shutil.rmtree(sec_dir, ignore_errors=True)

    return redirect(url_for("match_rules", job_id=job_id))


@app.route("/job/<job_id>/match")
def match_rules(job_id):
    d, state = _match_state(job_id)
    an = _load_analysis(d)
    total = an["total"]
    mode = "column" if an.get("sec_has_status") else "done"
    # الأعداد تتغيّر بتغيّر اختيار الحالة، فنحسب الأوضاع الثلاثة كلها ونترك
    # الصفحة تبدّلها لحظيًا: كان المستخدم يرى عددًا لا يطابق ما سيخرج
    mode_counts = {}
    for m in ("column", "done", "wip"):
        st = matcher.resolve_states(an, m)
        dn, wp = st.count(matcher.ST_DONE), st.count(matcher.ST_WIP)
        mode_counts[m] = {
            "done": dn, "wip": wp, "remaining": total - dn - wp,
            "percent": round(100.0 * dn / total, 1) if total else 0.0,
        }
    states = matcher.resolve_states(an, mode)
    done = states.count(matcher.ST_DONE)
    wip = states.count(matcher.ST_WIP)
    columns = [{"i": i, "name": h} for i, h in enumerate(an["headers"]) if h]
    return render_template(
        "match_rules.html", job_id=job_id, state=state, an=an,
        total=total, done=done, wip=wip, remaining=total - done - wip,
        percent=round(100.0 * done / total, 1) if total else 0.0,
        columns=columns, sec_mode=mode, states=matcher.EXEC_STATES,
        mode_counts=mode_counts,
        state_color={matcher.ST_DONE: "green", matcher.ST_WIP: "yellow",
                     matcher.ST_NONE: "none"},
        n_found=an.get("n_found", 0),
        n_found_rows=an.get("n_found_rows", 0),
        ledger=an.get("ledger") or [],
        sec_rows=sum(f["total"] for f in (an.get("ledger") or [])),
        distinct={str(k): v for k, v in an["distinct"].items()},
        colors=matcher.COLORS, max_rules=matcher.MAX_RULES,
        n_missing=len(an["missing"]), n_dups=len(an["dups"]),
        n_dup_sec=len(an.get("dup_sec_rows") or []),
        n_main_blank=len(an.get("main_blank") or []),
        has_status=an.get("status_col") is not None,
        sec_has_contractor=an.get("sec_has_contractor", False),
        n_sec_contractor=len(an.get("sec_contractor") or {}),
    )


@app.route("/job/<job_id>/match/build", methods=["POST"])
def match_build(job_id):
    d, state = _match_state(job_id)
    an = _load_analysis(d)

    rules = []
    for exec_state in matcher.EXEC_STATES:
        color = request.form.get("state_color_" + exec_state) or ""
        if color in matcher.COLOR_HEX:
            rules.append({"kind": "state", "key": exec_state, "color": color})
    for i in range(1, matcher.MAX_RULES + 1):
        col = request.form.get(f"rule_col_{i}") or ""
        val = (request.form.get(f"rule_val_{i}") or "").strip()
        color = request.form.get(f"rule_color_{i}") or ""
        if not col or not val or color not in matcher.COLOR_HEX:
            continue
        rules.append({"kind": "column", "key": _int_ar(col, -1),
                      "value": val, "color": color})

    # المستخدم قد يصحّح عمود التصنيف إن أخطأ الاستنتاج التلقائي
    raw = request.form.get("class_col")
    if raw is not None:
        an["class_col"] = _int_ar(raw, -1) if raw != "" else -1
        if an["class_col"] < 0 or an["class_col"] >= len(an["headers"]):
            an["class_col"] = None

    contractor_mode = request.form.get("contractor_mode") or "column"

    src = state.get("main_path") or ""
    if not os.path.exists(src):
        return render_template("match_rules.html", job_id=job_id, state=state, an=an,
                               error="انتهت صلاحية الجلسة — أعد رفع الملفات."), 410

    out_path = os.path.join(d, "output.xlsx")
    try:
        with heavy_lock():
            result = matcher.build_output(
                an, rules, src, out_path,
                set_status_done=bool(request.form.get("status_done")),
                add_source_col=bool(request.form.get("source_col")),
                move_contractor=bool(request.form.get("move_contractor")),
                contractor_mode=contractor_mode,
                sec_status_mode=request.form.get("sec_status_mode") or "done",
            )
    except Exception as exc:  # noqa: BLE001
        return render_template("match_rules.html", job_id=job_id, state=state, an=an,
                               error=f"تعذّر بناء الملف: {exc}"), 500

    state["result"] = result
    state["output_path"] = out_path
    write_state(d, state)
    return redirect(url_for("match_result", job_id=job_id))


@app.route("/job/<job_id>/match/result")
def match_result(job_id):
    _d, state = _match_state(job_id)
    if not state.get("result"):
        return redirect(url_for("match_rules", job_id=job_id))
    return render_template("match_result.html", job_id=job_id,
                           state=state, r=state["result"])


@app.route("/job/<job_id>/match/download")
def match_download(job_id):
    _d, state = _match_state(job_id)
    if not state.get("output_path") or not os.path.exists(state["output_path"]):
        abort(404)
    return send_file(state["output_path"], as_attachment=True,
                     download_name=f'متابعة_{state.get("stem") or "الملاحظات"}.xlsx')


# --------------------------------------------------------------------------
# حالة الاعتماد: ملف واحد -> تلوين حسب الحالة + داتا شيت
# --------------------------------------------------------------------------

def _approve_state(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "approve":
        abort(404)
    return d, state


@app.route("/approve", methods=["GET"])
def approve_form():
    cleanup_old_jobs()
    return render_template("approve_new.html")


@app.route("/approve/start", methods=["POST"])
def approve_start():
    f = request.files.get("excel_file")
    if not f or not f.filename:
        return render_template("approve_new.html", error="الرجاء اختيار الملف أولًا."), 400
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template("approve_new.html",
                               error="الملف يجب أن يكون بصيغة ‎.xlsx‎."), 400

    job_id, d = new_job_dir()
    path = os.path.join(d, "input.xlsx")
    f.save(path)
    try:
        with heavy_lock():
            an = approval.analyze(path)
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(d, ignore_errors=True)
        return render_template("approve_new.html",
                               error=f"تعذّرت قراءة الملف: {exc}"), 400

    with open(os.path.join(d, "analysis.pkl"), "wb") as fh:
        pickle.dump(an, fh, protocol=pickle.HIGHEST_PROTOCOL)
    write_state(d, {
        "job_id": job_id,
        "kind": "approve",
        "stem": os.path.splitext(os.path.basename(f.filename))[0][:60] or "الملاحظات",
        "main_path": path,
        "created_at": time.time(),
    })
    return redirect(url_for("approve_rules", job_id=job_id))


@app.route("/job/<job_id>/approve")
def approve_rules(job_id):
    d, state = _approve_state(job_id)
    an = _load_analysis(d)
    mapping = {s["raw"]: s["bucket"] for s in an["statuses"]}
    buckets = approval.resolve(an, mapping)
    counts = {st: buckets.count(st) for st in approval.AP_STATES}
    total = an["total"]
    return render_template(
        "approve_rules.html", job_id=job_id, state=state, an=an,
        total=total, counts=counts,
        percent=round(100.0 * counts[approval.AP_APPROVED] / total, 1) if total else 0.0,
        states=approval.AP_STATES, colors=matcher.COLORS,
        default_colors=approval.DEFAULT_COLORS,
        groupable=an["groupable"],
        # الافتراضي جدول واحد فقط (التصنيف)؛ المقاول والأولوية يبقيان متاحين
        # في القائمتين الأخريين لمن أرادهما، لكنهما مغلقان دائمًا في البداية
        defaults=[c for c in (an.get("class_col"),) if c is not None],
    )


@app.route("/job/<job_id>/approve/build", methods=["POST"])
def approve_build(job_id):
    d, state = _approve_state(job_id)
    an = _load_analysis(d)

    mapping = {}
    for s in an["statuses"]:
        chosen = request.form.get("map_" + s["raw"])
        mapping[s["raw"]] = chosen if chosen in approval.AP_STATES else s["bucket"]

    colors = {}
    for st in approval.AP_STATES:
        c = request.form.get("color_" + st) or ""
        colors[st] = c if c in matcher.COLOR_HEX else "none"

    groups = []
    for i in range(1, 4):
        raw = request.form.get("group_%d" % i)
        if raw:
            c = _int_ar(raw, -1)
            if 0 <= c < len(an["headers"]) and c not in groups:
                groups.append(c)

    src = state.get("main_path") or ""
    if not os.path.exists(src):
        return render_template("approve_rules.html", job_id=job_id, state=state, an=an,
                               error="انتهت صلاحية الجلسة — أعد رفع الملف."), 410

    out_path = os.path.join(d, "output.xlsx")
    try:
        with heavy_lock():
            result = approval.build_output(an, mapping, colors, src, out_path,
                                           group_cols=groups)
    except Exception as exc:  # noqa: BLE001
        return render_template("approve_rules.html", job_id=job_id, state=state, an=an,
                               error=f"تعذّر بناء الملف: {exc}"), 500

    state["result"] = result
    state["output_path"] = out_path
    write_state(d, state)
    return redirect(url_for("approve_result", job_id=job_id))


@app.route("/job/<job_id>/approve/result")
def approve_result(job_id):
    _d, state = _approve_state(job_id)
    if not state.get("result"):
        return redirect(url_for("approve_rules", job_id=job_id))
    return render_template("approve_result.html", job_id=job_id, state=state,
                           r=state["result"], states=approval.AP_STATES)


@app.route("/job/<job_id>/approve/download")
def approve_download(job_id):
    _d, state = _approve_state(job_id)
    if not state.get("output_path") or not os.path.exists(state["output_path"]):
        abort(404)
    name = (state.get("result") or {}).get("feeder") or state.get("stem") or "الملاحظات"
    return send_file(state["output_path"], as_attachment=True,
                     download_name=f"اعتماد_{name}.xlsx".replace(" — ", "-"))


# --------------------------------------------------------------------------
# فرز الملاحظات: تصفية بالمغذي واللون والنوع، وترتيب ثابت
# --------------------------------------------------------------------------

def _sort_state(job_id):
    d = job_dir(job_id)
    state = read_state(d)
    if state.get("kind") != "sort":
        abort(404)
    return d, state


@app.route("/sort", methods=["GET"])
def sort_form():
    cleanup_old_jobs()
    return render_template("sort_new.html")


@app.route("/sort/start", methods=["POST"])
def sort_start():
    f = request.files.get("excel_file")
    if not f or not f.filename:
        return render_template("sort_new.html", error="الرجاء اختيار الملف أولًا."), 400
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template("sort_new.html",
                               error="الملف يجب أن يكون بصيغة ‎.xlsx‎."), 400

    job_id, d = new_job_dir()
    path = os.path.join(d, "input.xlsx")
    f.save(path)
    try:
        with heavy_lock():
            an = sorter.analyze(path)
    except Exception as exc:  # noqa: BLE001
        shutil.rmtree(d, ignore_errors=True)
        return render_template("sort_new.html",
                               error=f"تعذّرت قراءة الملف: {exc}"), 400

    if not an["total"]:
        shutil.rmtree(d, ignore_errors=True)
        return render_template("sort_new.html",
                               error="لم يُعثر على أي صف بيانات في الملف."), 400

    with open(os.path.join(d, "analysis.pkl"), "wb") as fh:
        pickle.dump(an, fh, protocol=pickle.HIGHEST_PROTOCOL)
    write_state(d, {
        "job_id": job_id,
        "kind": "sort",
        "stem": os.path.splitext(os.path.basename(f.filename))[0][:60] or "الملاحظات",
        "main_path": path,
        "created_at": time.time(),
    })
    return redirect(url_for("sort_rules", job_id=job_id))


@app.route("/job/<job_id>/sort")
def sort_rules(job_id):
    _d, state = _sort_state(job_id)
    an = _load_analysis(job_dir(job_id))
    return render_template("sort_rules.html", job_id=job_id, state=state, an=an,
                           color_order=sorter.COLOR_ORDER)


@app.route("/job/<job_id>/sort/build", methods=["POST"])
def sort_build(job_id):
    d, state = _sort_state(job_id)
    an = _load_analysis(d)

    feeders = request.form.getlist("feeder")
    colors = [c for c in request.form.getlist("color")
              if c in dict(sorter.COLOR_ORDER)]
    types = request.form.getlist("type")

    def again(msg, code=400):
        return render_template("sort_rules.html", job_id=job_id, state=state, an=an,
                               color_order=sorter.COLOR_ORDER, error=msg), code

    if an["has_feeder"] and not feeders:
        return again("اختر مغذيًا واحدًا على الأقل.")
    if not colors:
        return again("اختر لونًا واحدًا على الأقل.")
    if an["has_type"] and not an["types_too_many"] and not types:
        return again("اختر نوع ملاحظة واحدًا على الأقل.")

    order = sorter.plan(an, feeders, colors, types)
    if not order:
        return again("لا يوجد صف واحد يطابق اختيارك — وسّع الاختيار.")

    src = state.get("main_path") or ""
    if not os.path.exists(src):
        return again("انتهت صلاحية الجلسة — أعد رفع الملف.", 410)

    out_path = os.path.join(d, "output.xlsx")
    try:
        with heavy_lock():
            res = sorter.write_sorted(src, out_path, an["sheet_path"],
                                      an["header_row"], order)
    except Exception as exc:  # noqa: BLE001
        return again(f"تعذّر بناء الملف: {exc}", 500)

    kept = set(order)
    breakdown = {}
    feeders_out = []
    for r in an["records"]:
        if r["row"] in kept:
            breakdown[r["color"]] = breakdown.get(r["color"], 0) + 1
            if r["feeder"] not in feeders_out:
                feeders_out.append(r["feeder"])
    res.update({
        "total_in": an["total"],
        "removed": an["total"] - len(order),
        "breakdown": [(sorter.COLOR_LABEL[k], breakdown[k])
                      for k, _ in sorter.COLOR_ORDER if k in breakdown],
        "feeders": sorted(feeders_out, key=sorter.natural_key),
        "n_types": len(types) if types else 0,
        "all_types": bool(an["has_type"]) and len(types) == len(an["types"]),
    })
    state["result"] = res
    state["output_path"] = out_path
    write_state(d, state)
    return redirect(url_for("sort_result", job_id=job_id))


@app.route("/job/<job_id>/sort/result")
def sort_result(job_id):
    _d, state = _sort_state(job_id)
    if not state.get("result"):
        return redirect(url_for("sort_rules", job_id=job_id))
    return render_template("sort_result.html", job_id=job_id, state=state,
                           r=state["result"])


@app.route("/job/<job_id>/sort/download")
def sort_download(job_id):
    _d, state = _sort_state(job_id)
    if not state.get("output_path") or not os.path.exists(state["output_path"]):
        abort(404)
    return send_file(state["output_path"], as_attachment=True,
                     download_name=sorter.suggested_name(state.get("stem") or "الملاحظات"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
