"""เว็บแปลงข้อมูลฝึกงานให้เป็น Anonymous แบบหลายขั้นตอน
อัพโหลด -> เลือกตาราง -> ผล anonymize -> จัดกลุ่มอาชีพ -> ผลจัดกลุ่ม -> ผลรวม (ดาวน์โหลดได้ทั้ง CSV และ JSON)
ไม่มีฐานข้อมูล ไม่มีล็อกอิน ไม่เก็บไฟล์ใดๆ ไว้บนเซิร์ฟเวอร์ (ประมวลผลในหน่วยความจำแล้วส่งกลับทันที)
สถานะของขั้นตอนต่างๆ อยู่ในเบราว์เซอร์ของผู้ใช้เท่านั้น"""
import os, io, re, json, base64, threading
from urllib.parse import urlparse
from flask import Flask, Request, request, jsonify, render_template

import anonymizer as an

class MemRequest(Request):
    """ไฟล์ที่อัพโหลดเก็บในหน่วยความจำล้วนๆ (ค่าเริ่มต้นของ Flask จะพักไฟล์ใหญ่ลงดิสก์ชั่วคราว)"""
    def _get_file_stream(self, total_content_length, content_type, filename=None, content_length=None):
        return io.BytesIO()


app = Flask(__name__)
app.request_class = MemRequest
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_MB", "50")) * 1024 * 1024
app.config["MAX_FORM_MEMORY_SIZE"] = app.config["MAX_CONTENT_LENGTH"]     # mapping.json ของหลายปีอาจใหญ่เกินค่าเริ่มต้น (500 KB)
app.config["MAX_FORM_PARTS"] = 2000
_lock = threading.Lock()          # การแปลงใช้ stdout ชั่วคราว จึงทำทีละคำขอ

APP_VERSION = "2.0"
CSV, JSON = "text/csv;charset=utf-8", "application/json"


def fail(msg, code=400):
    return jsonify(error=msg), code


# ---- กันเว็บอื่นในเบราว์เซอร์เดียวกันแอบเรียกใช้ (DNS rebinding / CSRF) ----
_LOOPBACK = {"localhost", "127.0.0.1", "[::1]"}
_ALLOWED = {h.strip().lower() for h in os.environ.get("ALLOWED_HOSTS", "").split(",") if h.strip()}
_LOCAL_ONLY = os.environ.get("HOST", "127.0.0.1") in ("127.0.0.1", "localhost", "::1")


def _hostname(hostport):
    m = re.fullmatch(r"(\[[0-9a-f:]+\]|[^:]+)(:\d+)?", (hostport or "").lower())
    return m.group(1) if m else ""


@app.before_request
def guard():
    if _LOCAL_ONLY or _ALLOWED:                          # เปิดเฉพาะเครื่องนี้ (ค่าเริ่มต้น) หรือกำหนดชื่อเว็บที่อนุญาตไว้
        if _hostname(request.host) not in (_LOOPBACK | _ALLOWED):
            return fail("ไม่อนุญาตให้เข้าด้วยที่อยู่นี้ (ใช้ http://localhost:พอร์ต)", 403)
    if request.method == "POST":
        origin = request.headers.get("Origin")
        if origin:                                       # เบราว์เซอร์ใส่ Origin มาเสมอเมื่อเป็นการส่งข้ามเว็บ ("null" ก็ไม่รับ)
            o = urlparse(origin).netloc
            if o != request.host and _hostname(o) not in _ALLOWED:
                return fail("คำขอมาจากหน้าเว็บอื่น จึงไม่อนุญาต", 403)


@app.after_request
def secure_headers(resp):
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; "
        "img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'")
    return resp


@app.errorhandler(413)
def too_big(_):
    return fail(f"ไฟล์ใหญ่เกินไป (รวมกันไม่เกิน {app.config['MAX_CONTENT_LENGTH'] // 1024 // 1024} MB)", 413)


@app.errorhandler(an.UserError)
def user_error(e):
    return fail(str(e))


def _log_exception(exc_info):
    """log เฉพาะชนิดของ error: ข้อความ/traceback ของ pandas อาจมีค่าจากไฟล์จริง (ชื่อ เบอร์ ฯลฯ) ซึ่งไม่ควรโผล่ในหน้าต่างดำหรือ log"""
    app.logger.error("internal error: %s", getattr(exc_info[0], "__name__", "?"))


app.log_exception = _log_exception


@app.errorhandler(500)
def server_error(e):
    kind = type(getattr(e, "original_exception", None) or e).__name__      # ชื่อชนิด error เท่านั้น ไม่มีข้อมูลในไฟล์
    return fail(f"เกิดข้อผิดพลาดภายในระบบ ({kind}) ลองใหม่อีกครั้ง หรือส่งข้อความนี้ให้ผู้พัฒนา", 500)


# ---------------------------------------------------------------- ตัวช่วย
def b64(b):
    return base64.b64encode(b).decode("ascii")


def out_file(name, mime, data):
    return dict(filename=name, mime=mime, b64=b64(data))


def dataset_outputs(df, prefix):
    return [out_file(prefix + ".csv", CSV, an.csv_bytes(df)), out_file(prefix + ".json", JSON, an.json_bytes(df))]


def uploaded_files():
    files = []
    for f in request.files.getlist("files"):
        name = os.path.basename(f.filename or "")
        if name: files.append((name, f.read()))
    return files


def form_json(name, default=None):
    raw = request.form.get(name)
    if not raw: return default
    try:
        return json.loads(raw)
    except ValueError:
        raise an.UserError(f"ข้อมูล {name} ไม่ถูกต้อง")


def read_mapping(files):
    """mapping จากช่องฟอร์ม (สถานะปัจจุบันในเบราว์เซอร์) มาก่อน ไม่งั้นใช้ไฟล์ mapping.json ที่แนบมา"""
    obj = form_json("mapping")
    if obj is None:
        for n, d in files:
            if an.classify_file(n, d) == "mapping":
                try: obj = json.loads(d.decode("utf-8-sig"))
                except ValueError: raise an.UserError("อ่านไฟล์ mapping.json ไม่ได้")
    return an.validate_mapping(obj) if obj else None


def alias_param():
    a = form_json("alias", {})
    if not (isinstance(a, dict) and all(isinstance(k, str) and isinstance(v, str) for k, v in a.items())):
        raise an.UserError("ข้อมูลการรวมชื่อบริษัทไม่ถูกต้อง")
    return a


# ---------------------------------------------------------------- หน้าเว็บ
@app.get("/")
def index():
    return render_template("index.html", version=APP_VERSION)


@app.get("/health")
def health():
    return "ok " + APP_VERSION


# ---------------------------------------------------------------- ขั้น 1->2: ดูโครงสร้างไฟล์ (ยังไม่แปลง)
@app.post("/analyze")
def analyze():
    files = uploaded_files()
    if not files: return fail("ยังไม่ได้เลือกไฟล์")
    kinds = [dict(name=n, kind=an.classify_file(n, d)) for n, d in files]
    for k in kinds:
        if k["kind"] in ("unknown", "unsupported"):
            return fail(f"ไม่รองรับไฟล์ “{k['name']}” (รองรับ .xlsx, .csv และ mapping.json / ไฟล์รวมปีก่อน)")
    data = [(n, d) for (n, d), k in zip(files, kinds) if k["kind"] == "data"]
    if not data: return fail("ยังไม่ได้เลือกไฟล์ข้อมูล (Excel/CSV) ที่จะแปลง")
    overrides = form_json("header_overrides", {})
    if not (isinstance(overrides, dict) and all(isinstance(v, int) and 1 <= v <= 10 for v in overrides.values())):
        return fail("ข้อมูลจำนวนแถวหัวตารางไม่ถูกต้อง")
    mapping = read_mapping(files)
    with _lock:
        tables = an.analyze_tables(data, overrides)
    return jsonify(files=kinds, tables=tables, mapping=mapping, field_labels=an.FIELD_LABELS)


# ---------------------------------------------------------------- ขั้น 2->3: anonymize ตามตารางที่เลือก
@app.post("/anonymize")
def anonymize():
    files = uploaded_files()
    mapping = read_mapping(files)
    selection = form_json("selection")
    keep = request.form.get("keep_original") == "1"
    with _lock:
        r = an.run_pipeline(files, mapping, alias_param(), selection, keep, "anonymize")
    df = r["dataset"]
    return jsonify(
        tables=r["results"], rows=int(len(df)), columns=list(df.columns), counts=r["counts"],
        similar=r["similar"], titles=r["titles"], categories=r["categories"],
        mapping=r["mapping"], preview=an.preview_records(df),
        outputs=dataset_outputs(df, "anonymized"),
    )


# ---------------------------------------------------------------- ขั้น 4->5: บันทึกการจัดกลุ่มอาชีพ (ไม่ต้องใช้ไฟล์ข้อมูล)
@app.post("/positions")
def positions():
    body = request.get_json(silent=True)
    if not isinstance(body, dict): return fail("ข้อมูลไม่ถูกต้อง")
    res = an.build_position_result(body.get("mapping"), body.get("rows"), body.get("categories"))
    import pandas as pd
    df = pd.DataFrame(res["rows"], columns=["position", "category", "group", "count", "code", "confidence"])
    groups = (df.groupby("group")["count"].sum().sort_values(ascending=False)
              .rename_axis("group").reset_index().to_dict("records"))
    return jsonify(mapping=res["mapping"], rows=res["rows"], groups=groups, outputs=dataset_outputs(df, "position_groups"))


# ---------------------------------------------------------------- ขั้น 6: anonymize รวม (ข้อมูล + หมวด/กลุ่มตำแหน่ง + ต่อกับปีก่อน)
@app.post("/finalize")
def finalize():
    files = uploaded_files()
    mapping = read_mapping(files)
    selection = form_json("selection")
    keep = request.form.get("keep_original") == "1"
    with _lock:
        r = an.run_pipeline(files, mapping, alias_param(), selection, keep, "final")
    df = r["dataset"]
    by_group = (df["position_group"].fillna("ไม่มีข้อมูลตำแหน่ง").value_counts()
                .rename_axis("group").reset_index(name="count").to_dict("records"))
    return jsonify(
        rows=int(len(df)), columns=list(df.columns), counts=r["counts"], prior_rows=r["prior_rows"],
        by_group=by_group, mapping=r["mapping"], preview=an.preview_records(df),
        outputs=dataset_outputs(df, "all_anonymized"),
    )


if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")          # ค่าเริ่มต้น: เปิดเฉพาะเครื่องนี้
    port = int(os.environ.get("PORT", "5000"))
    try:
        from waitress import serve
        print(f"Anonymize v{APP_VERSION} (multi-step) running at http://{host}:{port}")
        serve(app, host=host, port=port, threads=4)
    except ImportError:
        app.run(host=host, port=port)
