"""แปลงข้อมูลฝึกงานให้เป็น anonymous + จัดกลุ่มตำแหน่งงาน (หมวด -> กลุ่ม)
ไม่เก็บอะไรลงดิสก์เลย: รับไฟล์เข้ามา -> ประมวลผลในหน่วยความจำ -> คืนผลลัพธ์"""
import os, re, io, csv, copy, json, difflib, contextlib, warnings, zipfile
import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", message="Could not infer format")
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =====================================================================
# 1) ตั้งค่า: คอลัมน์ที่ "เก็บ" (ที่ไม่อยู่ในนี้ถูกทิ้งหมด = ปลอดภัยโดยปริยาย)
#    pat = ชื่อหัวตารางที่ระบบรู้จักอัตโนมัติ  (ถ้าชื่อไม่ตรง ผู้ใช้จับคู่คอลัมน์เองได้ในหน้าเว็บ)
# =====================================================================
SCHEMA = {
    "advisor":    dict(pat=r"^(ลงชื่อ)?อาจารย์นิเทศ$", req=True),
    "company":    dict(pat=r"^บริษัท$",        req=True,  extra=True),
    "province":   dict(pat=r"^จังหวัด$",       req=True,  coalesce=True),
    "student_id": dict(pat=r"^ID$",            req=True),
    "position":   dict(pat=r"^ตำแหน่ง",        req=False, extra=True, exclude_group=r"ข้อมูลติดต่อ"),
    "course":     dict(pat=r"^รายวิชา$",       req=False),
    "mou":        dict(pat=r"^MOU",            req=False),
    "start":      dict(pat=r"^วันเริ่มฝึกงาน$", req=False),
    "end":        dict(pat=r"^สิ้นสุดฝึกงาน$",  req=False),
    "period":     dict(pat=r"^ช่วงเวลาฝึกงาน$", req=False, leaf=True),   # ปี 2024: ช่วงเวลาเป็นข้อความเดียว (ดูเฉพาะหัวแถวล่างสุด ไม่งั้นไปจับหัวกลุ่มที่คร่อมคอลัมน์วันเริ่ม/วันสิ้นสุด)
}
FIELD_LABELS = {   # ป้ายภาษาไทยสำหรับหน้าจับคู่คอลัมน์
    "student_id": "รหัสนักศึกษา", "company": "บริษัท / หน่วยงาน", "advisor": "อาจารย์นิเทศ", "province": "จังหวัด",
    "start": "วันเริ่มฝึกงาน", "end": "วันสิ้นสุดฝึกงาน", "period": "ช่วงเวลาฝึกงาน (ข้อความเดียว)",
    "position": "ตำแหน่งงาน", "course": "รายวิชา", "mou": "MOU", "date": "วันที่ฝึกงาน",
}
HEADER_ROWS = 3
NOT_REAL_COMPANY = ["กำลังหาที่ฝึกงาน"]
CODE_COLS = {"advisor": "อาจารย์", "company": "บริษัท", "student_id": "นักศึกษา"}


class UserError(ValueError):
    """ข้อความที่ปลอดภัยและเข้าใจง่ายพอจะแสดงให้ผู้ใช้เห็น"""


def col_letter(i):
    s, i = "", int(i) + 1
    while i:
        i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s


def detect_columns(raw, header_rows=HEADER_ROWS):
    """หาคอลัมน์จากชื่อหัวตาราง -> (found, missing) ไม่ raise"""
    head = raw.iloc[:header_rows]
    group = head.iloc[1].ffill() if len(head) > 1 else pd.Series(index=raw.columns, dtype=object)   # หัวกลุ่ม (merged cell)
    found = {}
    for name, r in SCHEMA.items():
        hits = []
        for c in raw.columns:
            cells = [str(x).strip() for x in head[c].dropna()]
            if r.get("leaf"): cells = cells[-1:]
            if not any(re.search(r["pat"], t, flags=re.I) for t in cells): continue
            g = str(group[c]) if pd.notna(group[c]) else ""
            if r.get("exclude_group") and re.search(r["exclude_group"], g): continue
            hits.append(c)
        if hits: found[name] = hits
    return found, missing_fields(found)


def missing_fields(found):
    miss = [n for n, r in SCHEMA.items() if r["req"] and n not in found]
    if not ("start" in found and "end" in found) and "period" not in found: miss.append("date")
    return miss


def find_columns(raw, header_rows=HEADER_ROWS):
    found, miss = detect_columns(raw, header_rows)
    if miss: raise UserError("หาคอลัมน์ไม่เจอ: " + ", ".join(FIELD_LABELS[m] for m in miss))
    return found


def col_label(raw, idx, header_rows):
    """ชื่อคอลัมน์สำหรับแสดงในหน้าจับคู่: ชื่อที่เจาะจงที่สุด (แถวล่างสุดที่อ่านออก) นำหน้า ตามด้วยหัวกลุ่มแบบย่อ"""
    cells = [re.sub(r"\s+", " ", str(x)).strip() for x in raw.iloc[:header_rows][idx].dropna()]
    cells = list(dict.fromkeys(c for c in cells if c and re.search(r"[A-Za-zก-๙0-9]", c)))     # ข้ามเซลล์ที่เป็นแค่อีโมจิ/สัญลักษณ์
    if not cells: return "(ไม่มีชื่อหัวตาราง)"
    main = cells[-1][:40]
    ctx = " › ".join(c[:18] + ("…" if len(c) > 18 else "") for c in cells[:-1])
    return f"{main}  ({ctx})" if ctx else main


# =====================================================================
# 2) วันที่ภาษาไทย
# =====================================================================
TH_MONTHS = {"มค":1,"มกราคม":1,"กพ":2,"กุมภาพันธ์":2,"มีค":3,"มีนาคม":3,"เมย":4,"เมษายน":4,
             "พค":5,"พฤษภาคม":5,"มิย":6,"มิถุนายน":6,"กค":7,"กรกฎาคม":7,"สค":8,"สิงหาคม":8,
             "กย":9,"กันยายน":9,"ตค":10,"ตุลาคม":10,"พย":11,"พฤศจิกายน":11,"ธค":12,"ธันวาคม":12,
             "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
_MONTH_KEYS = sorted(TH_MONTHS, key=len, reverse=True)
_NOISE = r"วันที่|วัน(จันทร์|อังคาร|พุธ|พฤหัสบดี|ศุกร์|เสาร์|อาทิตย์)|พ\.ศ\.|ต้นเดือน|กลางเดือน|ปลายเดือน|สิ้นเดือน|ค\.ศ\.|ที่"


def _part(t):
    t = re.sub(r"\(.*?\)", " ", t)                      # ตัดวงเล็บ เช่น (2 เดือน)
    t = re.sub(_NOISE, " ", t)
    y = re.search(r"\d{4}", t); y = int(y.group()) if y else None
    if y and y > 2400: y -= 543                         # พ.ศ. -> ค.ศ.
    t2 = re.sub(r"\d{3,4}", "", t)
    d = re.search(r"\d{1,2}", t2); d = int(d.group()) if d else None
    letters = re.sub(r"[^a-zก-๙]", "", t2.lower())      # เหลือเฉพาะตัวอักษร (ตัด ., -, ช่องว่าง)
    m = next((TH_MONTHS[k] for k in _MONTH_KEYS if letters.startswith(k)), None)
    return d, m, y


def parse_period(text, default_year=None):
    if pd.isna(text): return pd.NaT, pd.NaT
    parts = re.split(r"\s*[-–—~]\s*|\s+ถึง\s+", str(text).strip())
    if len(parts) != 2:                                 # ไม่มีขีดคั่น เช่น "1 มิ.ย. 30 พ.ย."
        parts = re.findall(r"\d{1,2}\s*[^\d\s]+", str(text))
        if len(parts) != 2: return pd.NaT, pd.NaT
    (sd, sm, sy), (ed, em, ey) = _part(parts[0]), _part(parts[1])
    if sm is None or em is None: return pd.NaT, pd.NaT
    if sy is None and ey is None:                       # ไม่มีปีเลย -> ใช้ปีจากชื่อไฟล์
        if default_year is None: return pd.NaT, pd.NaT
        sy, ey = default_year, default_year + (1 if em < sm else 0)
    if ey is None: ey = sy + (1 if em < sm else 0)
    if sy is None: sy = ey - (1 if sm > em else 0)
    try:
        s = pd.Timestamp(sy, sm, sd or 1)
        e = pd.Timestamp(ey, em, ed) if ed else pd.Timestamp(ey, em, 1) + pd.offsets.MonthEnd(0)
        return s, e
    except ValueError:
        return pd.NaT, pd.NaT


def fix_text_dates(raw_col, parsed, ref, default_year):
    """ช่องวันที่ที่เป็นข้อความ เช่น '28 Nov.' (to_datetime อ่านไม่ได้) -> แปลงเอง"""
    for i in parsed.index[parsed.isna() & raw_col.notna()]:
        d, m, y = _part(str(raw_col[i]))
        if m is None: continue
        if y is None:
            r = ref.get(i, pd.NaT)
            base = r.year if pd.notna(r) else default_year
            if base is None: continue
            y = base + (1 if pd.notna(r) and m < r.month else 0)
        try: parsed[i] = pd.Timestamp(y, m, d or 1)
        except ValueError: pass
    return parsed


# =====================================================================
# 3) ทำความสะอาด + ออกรหัส
# =====================================================================
def norm_company(s):
    s = re.sub(r"\s+", " ", str(s)).strip().lower()
    s = s.split(" / ")[0]
    s = re.sub(r"บริษัท|จำกัด|\(มหาชน\)|co\.?,? ?ltd\.?|public company limited", "", s)
    return re.sub(r"\s+", " ", s).strip()


def next_no(m):
    nums = [int(x.group()) for x in (re.search(r"\d+$", str(v)) for v in m.values()) if x]
    return max(nums, default=0) + 1


def norm_id(v):
    if isinstance(v, (int, float)) and float(v).is_integer(): return str(int(v))   # 6599.0 -> "6599"
    return str(v).strip().lower()


RESERVED_NAMES = {"source", "table", "year", "months", "start", "end", "student_id", "company", "advisor", "province",
                  "course", "mou", "position", "position_code", "position_category", "position_group", "position_confidence"}
RESERVED_NAMES |= {f"{n}_2" for n in list(RESERVED_NAMES)}          # ชื่อที่ระบบใช้เอง (รวมรูปแบบ xxx_2) ห้ามให้คอลัมน์เสริมใช้ซ้ำ


def _unique_name(base, used):
    base = re.sub(r"\s+", " ", str(base)).strip()[:60] or "column"
    name, i = base, 2
    while name in used or name.lower() in {u.lower() for u in used}:
        name = f"{base}_{i}"; i += 1
    return name


def process_sheet(raw, sheet, mapping, alias, default_year=None, cols=None, header_rows=HEADER_ROWS, extras=None, stats=None):
    if cols is None: cols = find_columns(raw, header_rows)               # โหมดอัตโนมัติ: ต้องหัวตารางครบตามที่รู้จัก
    elif "student_id" not in cols:                                        # ผู้ใช้เลือกเอง: ต้องมีอย่างน้อยคอลัมน์ที่บอกว่าแต่ละแถวคือใคร
        raise UserError("ต้องเลือกคอลัมน์รหัสนักศึกษา (ใช้บอกว่าแต่ละแถวคือใคร)")
    body = raw.iloc[header_rows:]
    out = pd.DataFrame(index=body.index)
    for name, hits in cols.items():
        if SCHEMA[name].get("coalesce"): out[name] = body[hits].bfill(axis=1).iloc[:, 0]
        else:
            out[name] = body[hits[0]]
            if SCHEMA[name].get("extra") and len(hits) > 1: out[name + "_2"] = body[hits[1]]
    taken = {i for v in cols.values() for i in v}
    extra_cols = []                                                       # คอลัมน์เสริมที่ผู้ใช้เลือกเอง
    for e in extras or []:
        if e["idx"] not in raw.columns or e["idx"] in taken: continue
        base = e.get("name") or col_label(raw, e["idx"], header_rows).split("  (")[0]
        nm = _unique_name(base, set(out.columns) | RESERVED_NAMES | {x[0] for x in extra_cols})
        out[nm] = body[e["idx"]]; extra_cols.append((nm, e["mode"], (e.get("prefix") or nm)[:40]))
    no_id = out["student_id"].isna() & body.loc[out.index].notna().any(axis=1)      # แถวที่มีข้อมูลแต่ไม่มีรหัสนักศึกษา (แถวว่างเปล่าไม่นับ)
    out = out.dropna(subset=["student_id"])
    st = stats.setdefault(sheet, {}) if stats is not None else {}
    st["no_id"] = int(no_id.sum())
    if "course" not in out:                              # ปี 2024: รายวิชาอยู่ในชื่อชีต/ชื่อไฟล์
        codes = re.findall(r"ITDS\d{3}", sheet)
        if codes: out["course"] = "+".join(codes)

    if "company" in out:
        drop = out["company"].astype(str).str.contains("|".join(NOT_REAL_COMPANY), na=False)
        st["not_company"] = int(drop.sum())
        if drop.sum(): print(f"    ตัดแถวที่ไม่ใช่บริษัทจริง: {drop.sum()}")
        out = out[~drop].copy()

    for c in [c for c in ("company", "company_2") if c in out]:
        out[c] = out[c].map(lambda v: pd.NA if pd.isna(v) else alias.get(norm_company(v), norm_company(v)))
    if "company_2" in out:
        out.loc[out["company_2"] == out["company"], "company_2"] = pd.NA
        if out["company_2"].isna().all(): out = out.drop(columns="company_2")
    if "position_2" in out:
        same = out["position_2"].astype(str).str.strip().str.lower() == out["position"].astype(str).str.strip().str.lower()
        out.loc[same, "position_2"] = pd.NA
        if out["position_2"].isna().all(): out = out.drop(columns="position_2")

    for col, label in CODE_COLS.items():
        m = mapping.setdefault(col, {})
        def enc(v, m=m, label=label, col=col):
            if pd.isna(v): return pd.NA
            k = v if col == "company" else (norm_id(v) if col == "student_id" else str(v).strip().lower())
            if k not in m:
                if col == "company":
                    near = difflib.get_close_matches(k, list(m), n=1, cutoff=0.85)
                    if near: print(f"    ? '{k}' คล้าย '{near[0]}' -> ถ้าเป็นที่เดียวกัน ใส่ alias")
                m[k] = f"{label} {next_no(m)}"
            return m[k]
        for c in [c for c in (col, col + "_2") if c in out]: out[c] = out[c].map(enc)

    for nm, mode, prefix in extra_cols:                                   # คอลัมน์เสริมแบบ "แปลงเป็นรหัส"
        if mode != "code": continue
        mm = mapping.setdefault("_codes", {}).setdefault(nm, {})
        def enc_x(v, mm=mm, prefix=prefix):
            if pd.isna(v): return pd.NA
            k = str(v).strip().lower()
            if k not in mm: mm[k] = f"{prefix} {next_no(mm)}"
            return mm[k]
        out[nm] = out[nm].map(enc_x)

    if "start" in out and "end" in out:
        s = pd.to_datetime(out["start"], errors="coerce"); e = pd.to_datetime(out["end"], errors="coerce")
        s = s.where(s.dt.year >= 2000); e = e.where(e.dt.year >= 2000)     # กันปีเพี้ยนจาก '28 Nov.'
        s = fix_text_dates(out["start"], s, pd.Series(dtype="datetime64[ns]"), default_year)
        e = fix_text_dates(out["end"], e, s, default_year)
        out = out.drop(columns=[c for c in ("start", "end", "period") if c in out])
    elif "period" in out:
        se = out["period"].map(lambda t: parse_period(t, default_year))
        s = pd.to_datetime(se.map(lambda x: x[0])); e = pd.to_datetime(se.map(lambda x: x[1]))
        out = out.drop(columns=["period"])
    else:
        s = e = None
        if "start" in out or "end" in out:
            print("    วันที่ฝึกงานเลือกไม่ครบคู่ (วันเริ่ม + วันสิ้นสุด) จึงไม่คำนวณจำนวนเดือน")
            out = out.drop(columns=[c for c in ("start", "end") if c in out])
    bad_n = 0
    if s is not None:
        out["months"] = ((e - s).dt.days / 30).round().astype("Int64")
        out["year"]   = s.dt.year.astype("Int64")
        out["start"]  = s.dt.to_period("M").astype(str)
        out["end"]    = e.dt.to_period("M").astype(str)
        bad_n = int((out["months"].isna() | (out["months"] < 1) | (out["months"] > 12)).sum())
    print(f"    {len(out)} แถว, วันที่ผิดปกติ {bad_n} แถว (ตรวจมือ)")
    out["table"] = sheet
    return out


def _read_csv(content):
    """CSV จาก Excel (UTF-8 หรือ Thai Windows) -> DataFrame แบบไม่มีหัวตาราง ทุกช่องเป็นข้อความ"""
    rows = None
    for enc in ("utf-8-sig", "cp874"):
        try:
            rows = list(csv.reader(io.StringIO(content.decode(enc), newline="")))
            break
        except UnicodeDecodeError: continue
    if rows is None: raise UserError("อ่านไฟล์ CSV ไม่ได้ (ลองบันทึกเป็น CSV UTF-8)")
    width = max((len(r) for r in rows), default=0)
    df = pd.DataFrame([r + [""] * (width - len(r)) for r in rows], dtype=object)
    return df.replace(r"^\s*$", pd.NA, regex=True)


MAX_UNZIPPED_MB = int(os.environ.get("MAX_UNZIPPED_MB", "300"))


def check_xlsx_safe(content):
    """กัน zip bomb: .xlsx คือไฟล์ zip ขนาดเล็กอาจแตกออกมาใหญ่มากจนเครื่องค้าง"""
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as z:
            infos = z.infolist()
            if len(infos) > 5000 or sum(i.file_size for i in infos) > MAX_UNZIPPED_MB * 1024 * 1024:
                raise UserError("ไฟล์ Excel นี้มีข้อมูลภายในใหญ่ผิดปกติ จึงไม่เปิดให้ (ลองบันทึกใหม่จาก Excel หรือแยกไฟล์)")
    except zipfile.BadZipFile:
        raise UserError("อ่านไฟล์ไม่ได้ (ไม่ใช่ไฟล์ Excel .xlsx ที่ถูกต้อง หรือไฟล์เสียหาย)")


def read_sheets(content, filename):
    """content = bytes ของไฟล์ (อ่านจากหน่วยความจำล้วนๆ ไม่มีไฟล์ชั่วคราวบนดิสก์)"""
    if filename.lower().endswith(".csv"):                 # CSV มีชีตเดียว ใช้ชื่อไฟล์แทนชื่อชีต (หารหัสวิชา ITDS...)
        return [(os.path.splitext(os.path.basename(filename))[0], _read_csv(content))]
    check_xlsx_safe(content)
    with pd.ExcelFile(io.BytesIO(content)) as xl:
        return [(sh, xl.parse(sh, header=None)) for sh in xl.sheet_names]


def parse_selection(sel):
    """รายการตารางที่ผู้ใช้เลือกจากหน้าเว็บ -> {(ไฟล์, ชีต): {enabled, header_rows, columns}}"""
    bad = UserError("ข้อมูลการเลือกตารางไม่ถูกต้อง")
    if not isinstance(sel, list): raise bad
    out = {}
    for s in sel:
        if not isinstance(s, dict) or not isinstance(s.get("file"), str) or not isinstance(s.get("sheet"), str): raise bad
        hr = s.get("header_rows", HEADER_ROWS)
        if not isinstance(hr, int) or not 1 <= hr <= 10: raise bad
        cols = s.get("columns")
        if cols is not None:
            if not isinstance(cols, dict): raise bad
            clean = {}
            for k, v in cols.items():
                v = v if isinstance(v, list) else [v]
                if k not in SCHEMA or not v or not all(isinstance(i, int) and 0 <= i < 2000 for i in v): raise bad
                clean[k] = v
            cols = clean
        extras = s.get("extras", [])
        if not isinstance(extras, list): raise bad
        clean_x = []
        for e in extras:
            if not (isinstance(e, dict) and isinstance(e.get("idx"), int) and 0 <= e["idx"] < 2000 and e.get("mode") in ("keep", "code")): raise bad
            clean_x.append(dict(idx=e["idx"], mode=e["mode"], name=re.sub(r"\s+", " ", str(e.get("name") or "")).strip()[:60],
                                prefix=re.sub(r"\s+", " ", str(e.get("prefix") or "")).strip()[:40]))
        out[(s["file"], s["sheet"])] = dict(enabled=bool(s.get("enabled")), header_rows=hr, columns=cols, extras=clean_x)
    return out


def process_file(content, mapping, alias, filename, selection=None, stats=None):
    """selection=None -> โหมดอัตโนมัติ (ใช้ทุกชีตที่หัวตารางตรง) / มีค่า -> ใช้เฉพาะตารางที่ผู้ใช้ติ๊กและจับคู่คอลัมน์ไว้"""
    ym = re.search(r"(20\d{2})", os.path.basename(filename))
    default_year = int(ym.group(1)) if ym else None       # ปีจากชื่อไฟล์ ใช้เมื่อช่วงเวลาไม่ระบุปี
    frames = []
    for sh, raw in read_sheets(content, filename):
        spec = None
        if selection is not None:
            spec = selection.get((os.path.basename(filename), sh))
            if not spec or not spec["enabled"]: continue
        try:
            df = process_sheet(raw, sh, mapping, alias, default_year,
                               cols=(spec or {}).get("columns"), header_rows=(spec or {}).get("header_rows") or HEADER_ROWS,
                               extras=(spec or {}).get("extras"), stats=stats)
        except ValueError as ex:
            if spec: raise UserError(f"ตาราง “{sh}”: {ex}") from ex
            print(f"  ข้ามชีต '{sh}': {ex}"); continue
        print(f"  ใช้ชีต '{sh}'"); frames.append(df)
    if not frames:
        if selection is not None: return None          # ไฟล์นี้ไม่มีตารางที่ถูกติ๊ก -> ข้ามไป (ตรวจรวมที่ run_pipeline)
        raise UserError("ไม่พบตารางที่รู้จัก (หัวตารางไม่ตรงกับที่ระบบรองรับ)")
    return pd.concat(frames, ignore_index=True)


# ---------- ระบบเดาคอลัมน์ (ใช้เมื่อหัวตารางไม่ตรงกับที่รู้จัก) ----------
FIELD_HINTS = {          # คำที่มักอยู่ในชื่อหัวตาราง (ตัวพิมพ์เล็ก)
    "student_id": ["รหัสนักศึกษา", "รหัส", "student id", "student code", "student no", "studentid", "std id", "id"],
    "company":    ["บริษัท", "หน่วยงาน", "สถานประกอบการ", "สถานที่ฝึก", "องค์กร", "company", "organization", "organisation", "employer", "workplace"],
    "advisor":    ["อาจารย์นิเทศ", "อาจารย์", "นิเทศ", "advisor", "adviser", "supervisor"],
    "province":   ["จังหวัด", "province"],
    "start":      ["วันเริ่ม", "เริ่มฝึก", "start", "begin", "from"],
    "end":        ["สิ้นสุด", "วันสิ้นสุด", "end", "finish", "to"],
    "period":     ["ช่วงเวลา", "ระยะเวลาฝึก", "period", "duration"],
    "position":   ["ตำแหน่ง", "position", "job title", "job", "role", "หน้าที่"],
    "course":     ["รายวิชา", "วิชา", "course", "subject"],
    "mou":        ["mou"],
}
_NOT_FIELD = re.compile(r"ประสานงาน|ผุ้ประสาน|ติดต่อ|contact|e-?mail|อีเมล|โทร|phone|\btel\b|ที่อยู่|address|link|ลิงก์|ไปรษณีย์|zip")
PROVINCES = set("""กรุงเทพมหานคร กรุงเทพ กทม bangkok กระบี่ กาญจนบุรี กาฬสินธุ์ กำแพงเพชร ขอนแก่น จันทบุรี ฉะเชิงเทรา ชลบุรี ชัยนาท ชัยภูมิ ชุมพร
เชียงราย เชียงใหม่ ตรัง ตราด ตาก นครนายก นครปฐม นครพนม นครราชสีมา นครศรีธรรมราช นครสวรรค์ นนทบุรี นราธิวาส น่าน บึงกาฬ บุรีรัมย์ ปทุมธานี
ประจวบคีรีขันธ์ ปราจีนบุรี ปัตตานี พระนครศรีอยุธยา พะเยา พังงา พัทลุง พิจิตร พิษณุโลก เพชรบุรี เพชรบูรณ์ แพร่ ภูเก็ต มหาสารคาม มุกดาหาร แม่ฮ่องสอน
ยโสธร ยะลา ร้อยเอ็ด ระนอง ระยอง ราชบุรี ลพบุรี ลำปาง ลำพูน เลย ศรีสะเกษ สกลนคร สงขลา สตูล สมุทรปราการ สมุทรสงคราม สมุทรสาคร สระแก้ว สระบุรี
สิงห์บุรี สุโขทัย สุพรรณบุรี สุราษฎร์ธานี สุรินทร์ หนองคาย หนองบัวลำภู อ่างทอง อำนาจเจริญ อุดรธานี อุตรดิตถ์ อุทัยธานี อุบลราชธานี""".split())


def _has(h, text):
    if h.isascii() and len(h) <= 3: return re.search(rf"(?<![a-z0-9]){re.escape(h)}(?![a-z0-9])", text) is not None
    return h in text


def _share(vals, test):
    return sum(1 for v in vals if test(v)) / len(vals) if vals else 0.0


def _is_date_like(v):
    if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year"): return True
    t = str(v).strip()
    if not re.search(r"\d", t): return False
    d = pd.to_datetime(t, errors="coerce", dayfirst=False)
    return bool(pd.notna(d) and 1990 <= d.year <= 2100)


def _col_cells(raw, c, header_rows):
    cells = [re.sub(r"\s+", " ", str(x)).strip().lower() for x in raw.iloc[:header_rows][c].dropna() if str(x).strip()]
    return [t for t in cells if re.search(r"[a-zก-๙0-9]", t)]          # ข้ามเซลล์ที่เป็นแค่อีโมจิ/สัญลักษณ์


def suggest_columns(raw, header_rows, found):
    """เดาคอลัมน์ของฟิลด์ที่ยังหาไม่เจอ จากชื่อหัวตาราง + หน้าตาของค่าในคอลัมน์ -> (suggested, guessed_fields)"""
    body = raw.iloc[header_rows:]
    used = {c for v in found.values() for c in v}
    cand = []
    for c in raw.columns:
        if c in used: continue
        cells = _col_cells(raw, c, header_rows); leaf = cells[-1] if cells else ""
        vals = [v for v in body[c].dropna().head(60).tolist() if str(v).strip()]
        if not vals: continue
        sv = [str(v).strip() for v in vals]
        for f in FIELD_HINTS:
            if f in found or (f in ("start", "end") and "period" in found) or (f == "period" and ("start" in found and "end" in found)): continue
            hint = 0
            if leaf and not _NOT_FIELD.search(leaf):
                for h in FIELD_HINTS[f]:
                    if leaf == h: hint = max(hint, 6)
                    elif _has(h, leaf): hint = max(hint, 4)
                    elif any(_has(h, a) for a in cells[:-1]): hint = max(hint, 1)
            val = 0
            if f == "province" and _share([x.replace("จ.", "").replace(" ", "").lower() for x in sv], lambda x: x in PROVINCES) >= 0.6: val = 6
            elif f == "student_id" and _share(sv, lambda x: re.fullmatch(r"\d{7,12}(\.0)?", x) is not None) >= 0.8 and len(set(sv)) / len(sv) >= 0.8: val = 6
            elif f == "advisor" and _share([x.lower() for x in sv], lambda x: re.match(r"(aj\.?|อ\.|อาจารย์|ดร\.|dr\.?|ผศ|รศ)", x) is not None) >= 0.6: val = 6
            elif f == "course" and _share(sv, lambda x: re.search(r"ITDS\d{3}", x) is not None) >= 0.5: val = 6
            elif f in ("start", "end") and _share(vals, _is_date_like) >= 0.6: val = 2
            elif f == "period" and _share(sv, lambda x: re.search(r"\d{1,2}\s*[ก-๙a-zA-Z\.]+.*[-–—].*\d", x) is not None) >= 0.5: val = 5
            score = hint + val
            if score >= 3 or (f in ("start", "end") and val and hint >= 0 and score >= 2): cand.append((score, f, c))
    out, taken = {k: list(v) for k, v in found.items()}, set(used)
    for score, f, c in sorted(cand, key=lambda x: (-x[0], x[2])):
        if f in out or c in taken: continue
        if f in ("start", "end") and score < 3 and "start" not in out and f == "end": continue
        out[f] = [c]; taken.add(c)
    if "start" in out and "end" in out and out["start"][0] > out["end"][0]:           # วันเริ่มควรอยู่ซ้ายของวันสิ้นสุด
        out["start"], out["end"] = out["end"], out["start"]
    guessed = [f for f in out if f not in found]
    return out, guessed


_SENS_HEAD = re.compile(r"ชื่อ|นามสกุล|\bname\b|surname|ผู้ประสาน|ผุ้ประสาน|ติดต่อ|contact|e-?mail|อีเมล|โทร|phone|\btel\b|ที่อยู่|address|link|ลิงก์|ไปรษณีย์|zip|ไลน์|บัตร|passport|เลขที่")
_SAFE_HEAD = re.compile(r"บริษัท|company|หน่วยงาน|ตำแหน่ง|position|job")


def sensitive_flag(raw, c, header_rows):
    """คอลัมน์ที่น่าจะมีข้อมูลส่วนตัว (ชื่อ เบอร์ อีเมล ที่อยู่ ลิงก์) -> ใช้เตือนผู้ใช้ ไม่ได้บังคับ"""
    cells = _col_cells(raw, c, header_rows); leaf = cells[-1] if cells else ""
    if leaf and not _SAFE_HEAD.search(leaf) and _SENS_HEAD.search(leaf): return True
    vals = [str(v) for v in raw.iloc[header_rows:][c].dropna().head(40)]
    if not vals: return False
    hit = sum(1 for v in vals if "@" in v or re.search(r"https?://|www\.", v) or re.fullmatch(r"\s*0\d{1,2}[- ]?\d{3,4}[- ]?\d{3,4}\s*", v))
    return hit / len(vals) >= 0.3


def col_samples(raw, c, header_rows, n=3):
    seen, out = set(), []
    for v in raw.iloc[header_rows:][c].dropna().head(80):
        t = re.sub(r"\s+", " ", str(v)).strip()
        if not t or t in seen: continue
        seen.add(t); out.append(t[:24] + ("…" if len(t) > 24 else ""))
        if len(out) >= n: break
    return out


def analyze_tables(files, header_overrides=None):
    """ดูโครงสร้างทุกตาราง (ชีต) ในไฟล์ โดยยังไม่แปลงอะไร -> ใช้แสดงหน้า 'เลือกตาราง'"""
    header_overrides = header_overrides or {}
    out = []
    for name, content in files:
        try: sheets = read_sheets(content, name)
        except Exception as ex: raise UserError(f"{name}: {_friendly(ex)}") from ex
        for sh, raw in sheets:
            key = f"{name}||{sh}"
            hr = header_overrides.get(key)
            if hr:
                found, miss = detect_columns(raw, hr)
            else:
                hr = HEADER_ROWS; found, miss = detect_columns(raw, hr)
                if miss:                                   # ลองจำนวนแถวหัวตารางแบบอื่น (เช่น หัวตารางแถวเดียว)
                    for t in (1, 2, 4, 5):
                        f2, m2 = detect_columns(raw, t)
                        if not m2: hr, found, miss = t, f2, m2; break
            if miss:   suggested, guessed = suggest_columns(raw, hr, found)          # หัวตารางไม่ตรง: ให้ระบบเดาให้
            else:      suggested, guessed = {k: list(v) for k, v in found.items()}, []
            out.append(dict(id=key, file=name, sheet=sh, rows=int(raw.iloc[hr:].notna().any(axis=1).sum()),
                            header_rows=hr, ok=not miss, missing=miss,
                            detected={k: [int(i) for i in v] for k, v in found.items()},
                            suggested={k: [int(i) for i in v] for k, v in suggested.items()}, guessed=guessed,
                            complete=not missing_fields(suggested),
                            columns=[dict(idx=int(i), letter=col_letter(i), label=col_label(raw, i, hr),
                                          samples=col_samples(raw, i, hr), sensitive=sensitive_flag(raw, i, hr)) for i in raw.columns]))
    return out


# =====================================================================
# 4) จัดกลุ่มตำแหน่งงาน:  ตำแหน่ง -> หมวด (category) -> กลุ่ม (group)
#    กติกาแนะนำตั้งต้นอยู่ใน taxonomy.json (แก้ได้โดยไม่ต้องแตะโค้ด) และผู้ใช้ปรับรายตำแหน่ง/รายหมวดได้ในหน้าเว็บ
#    ที่ผู้ใช้ปรับจะถูกจำไว้ใน mapping.json (_position = ตำแหน่ง->หมวด, _taxonomy = รายการหมวด/กลุ่ม)
# =====================================================================
_TAX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "taxonomy.json")
with open(_TAX_PATH, encoding="utf-8") as _f:
    DEFAULT_TAX = json.load(_f)
DEFAULT_CATEGORIES = [dict(code=int(c["code"]), name=str(c["name"]), group=str(c["group"])) for c in DEFAULT_TAX["categories"]]
_R = DEFAULT_TAX["rules"]
_HARD = re.compile("|".join(_R["hard_other"]))
_SPECIFIC = {int(c): re.compile("|".join(p)) for c, p in _R["specific"].items()}
_GENERIC = re.compile("|".join(_R["generic_dev"]))
_INFERRED = {int(c): re.compile("|".join(p)) for c, p in _R["inferred"].items()}
_OVERRIDES = {k: int(v) for k, v in _R["overrides"].items()}
FALLBACK = int(_R.get("fallback_category", 16))


def norm_pos(t): return re.sub(r"\s+", " ", str(t)).strip().lower()


def categorize(text, extra=None):
    """-> (รหัสหมวด, ความมั่นใจ High/Medium/Low/Manual)   extra = {ชื่อตำแหน่ง(พิมพ์เล็ก): รหัสหมวด} ที่ผู้ใช้เลือกเอง"""
    if pd.isna(text) or not str(text).strip(): return None, None
    t = norm_pos(text)
    if extra and t in extra: return int(extra[t]), "Manual"
    if t in _OVERRIDES: return _OVERRIDES[t], "High"
    if _HARD.search(t): return FALLBACK, "High"                       # IT Support / Help Desk -> หมวด Other
    hits = sorted((m.start(), c) for c, p in _SPECIFIC.items() if (m := p.search(t)))
    if hits:                                                          # ลูกผสม -> อาชีพหลักคำแรก
        return hits[0][1], ("High" if len({c for _, c in hits}) == 1 else "Medium")
    if _GENERIC.search(t): return 1, "High"
    for code, p in _INFERRED.items():
        if p.search(t): return code, "Medium"
    return FALLBACK, "Low"                                            # ไม่แน่ใจ ห้ามเดา -> หมวด Other


def norm_course(v):
    """'ITDS397 Intern. II + ITDS499 Coop. Edu.' / 'ITDS397 การฝึกงาน 2 + ITDS499 สหกิจศึกษา' -> 'ITDS397+ITDS499'"""
    if pd.isna(v): return pd.NA
    codes = sorted(set(re.findall(r"ITDS\d{3}", str(v))))
    return "+".join(codes) if codes else str(v).strip()


def clean_taxonomy(obj, msg="ข้อมูลหมวด/กลุ่มตำแหน่งไม่ถูกต้อง"):
    """ตรวจรายการหมวด -> คืน {'categories': [...]} ที่สะอาด (ต้องมีหมวดตั้งต้นครบ ห้ามลบ แต่เปลี่ยนชื่อ/กลุ่มและเพิ่มหมวดใหม่ได้)"""
    cats = obj.get("categories") if isinstance(obj, dict) else None
    if not isinstance(cats, list) or not cats: raise UserError(msg)
    out, seen_codes, seen_names = [], set(), set()
    for c in cats:
        if not isinstance(c, dict): raise UserError(msg)
        try: code = int(c.get("code"))
        except (TypeError, ValueError): raise UserError(msg)
        name, group = str(c.get("name", "")).strip(), str(c.get("group", "")).strip()
        if code < 1 or code in seen_codes or not name or not group or len(name) > 80 or len(group) > 80: raise UserError(msg)
        if name.lower() in seen_names: raise UserError(f"ชื่อหมวด “{name}” ซ้ำกัน")
        seen_codes.add(code); seen_names.add(name.lower()); out.append(dict(code=code, name=name, group=group))
    if not {c["code"] for c in DEFAULT_CATEGORIES} <= seen_codes: raise UserError("ห้ามลบหมวดตั้งต้น (เปลี่ยนชื่อหรือย้ายกลุ่มได้)")
    return dict(categories=sorted(out, key=lambda c: c["code"]))


def effective_categories(mapping):
    tax = (mapping or {}).get("_taxonomy")
    return [dict(c) for c in (tax["categories"] if tax else DEFAULT_CATEGORIES)]


def add_position_columns(df, extra=None, cats=None):
    by = {c["code"]: c for c in (cats or DEFAULT_CATEGORIES)}
    if "position" not in df:
        for c in ("position_code", "position_category", "position_group", "position_confidence"): df[c] = pd.NA
        return df
    r = df["position"].map(lambda t: categorize(t, extra))
    df["position_code"] = r.map(lambda x: x[0]).astype("Int64")
    df["position_category"] = df["position_code"].map(lambda c: by[int(c)]["name"] if pd.notna(c) and int(c) in by else pd.NA)
    df["position_group"] = df["position_code"].map(lambda c: by[int(c)]["group"] if pd.notna(c) and int(c) in by else pd.NA)
    df["position_confidence"] = r.map(lambda x: x[1])
    return df


def titles_table(frames, extra):
    """ตำแหน่งที่ไม่ซ้ำกัน (จากข้อมูลดิบ) พร้อมจำนวน หมวดที่ระบบแนะนำ และหมวดที่ใช้อยู่ตอนนี้"""
    agg = {}
    for df in frames:
        if "position" not in df: continue
        for t in df["position"].dropna():
            k = norm_pos(t)
            if not k: continue
            a = agg.setdefault(k, dict(key=k, title=str(t).strip(), count=0)); a["count"] += 1
    out = []
    for a in agg.values():
        sug, conf0 = categorize(a["title"]); code, conf = categorize(a["title"], extra)
        out.append(dict(a, code=int(code), suggested=int(sug), confidence=conf))
    return sorted(out, key=lambda a: (-a["count"], a["title"].lower()))


def build_position_result(mapping, rows, categories):
    """รับสิ่งที่ผู้ใช้จัดในหน้าเว็บ -> อัปเดต mapping (_position/_taxonomy) + ตาราง ตำแหน่ง -> หมวด -> กลุ่ม"""
    mapping = validate_mapping(mapping) if mapping else {}
    cats = clean_taxonomy(dict(categories=categories))["categories"]
    by = {c["code"]: c for c in cats}
    pos = {k: v for k, v in mapping.get("_position", {}).items() if v in by}
    if not isinstance(rows, list): raise UserError("ข้อมูลตำแหน่งไม่ถูกต้อง")
    out = []
    for r in rows:
        try: title, code, count = str(r["title"]).strip(), int(r["code"]), int(r.get("count", 0))
        except (KeyError, TypeError, ValueError): raise UserError("ข้อมูลตำแหน่งไม่ถูกต้อง")
        if code not in by or not title: raise UserError("ข้อมูลตำแหน่งไม่ถูกต้อง")
        sug, conf = categorize(title)
        if code != sug: pos[norm_pos(title)] = code; conf = "Manual"
        else: pos.pop(norm_pos(title), None)
        out.append(dict(position=title, category=by[code]["name"], group=by[code]["group"], count=count, code=code, confidence=conf))
    mapping["_position"] = pos
    if cats == DEFAULT_CATEGORIES: mapping.pop("_taxonomy", None)
    else: mapping["_taxonomy"] = dict(categories=cats)
    out.sort(key=lambda r: (r["group"], r["code"], -r["count"], r["position"].lower()))
    return dict(mapping=mapping, rows=out)


# =====================================================================
# 5) ตัวแปลงหลัก: ไฟล์เข้า -> ผลลัพธ์ออก (ไม่เก็บอะไรไว้)
# =====================================================================
_SIMILAR = re.compile(r"\?\s*'(.*)'\s*คล้าย\s*'(.*)'\s*->")
_MASTER_COLS = {"student_id", "company", "source"}
PREFERRED_ORDER = ["source", "table", "year", "student_id", "company", "company_2", "advisor", "province", "course",
                   "start", "end", "months", "mou", "position_code", "position_category", "position_group",
                   "position_confidence", "position", "position_2"]


def safe_stem(filename):
    stem = os.path.splitext(os.path.basename(filename))[0]
    stem = re.sub(r"\s*\(\d+\)$", "", stem)               # ตัด " (1)"
    stem = re.sub(r"[^\w\-]+", "_", stem).strip("_")
    return stem or "upload"


def validate_mapping(obj):
    """ตรวจ mapping.json ที่ผู้ใช้นำมาจากปีก่อน -> คืน dict ที่สะอาดแล้ว"""
    if not isinstance(obj, dict): raise UserError("ไฟล์ mapping.json ไม่ถูกต้อง")
    if not any(k in obj for k in list(CODE_COLS) + ["_alias", "_position", "_codes", "_taxonomy"]):
        raise UserError("ไฟล์ .json นี้ไม่ใช่ mapping.json ของระบบนี้ (ถ้าใช้ผิดไฟล์ รหัสจะเริ่มนับใหม่และไม่ตรงกับปีก่อน)")
    out = {}
    for key in list(CODE_COLS) + ["_alias"]:
        part = obj.get(key, {})
        if not isinstance(part, dict) or not all(isinstance(k, str) and isinstance(v, str) for k, v in part.items()):
            raise UserError(f"ไฟล์ mapping.json ไม่ถูกต้อง (ส่วน {key})")
        out[key] = dict(part)
    pos = obj.get("_position", {})
    if not isinstance(pos, dict): raise UserError("ไฟล์ mapping.json ไม่ถูกต้อง (ส่วน _position)")
    cleaned = {}
    for k, v in pos.items():
        try: cleaned[norm_pos(k)] = int(v)
        except (TypeError, ValueError): raise UserError("ไฟล์ mapping.json ไม่ถูกต้อง (ส่วน _position)")
    out["_position"] = cleaned
    codes = obj.get("_codes", {})
    if codes:
        if not (isinstance(codes, dict) and all(isinstance(k, str) and isinstance(v, dict) and all(isinstance(a, str) and isinstance(b, str) for a, b in v.items()) for k, v in codes.items())):
            raise UserError("ไฟล์ mapping.json ไม่ถูกต้อง (ส่วน _codes)")
        out["_codes"] = {k: dict(v) for k, v in codes.items()}
    if obj.get("_taxonomy"): out["_taxonomy"] = clean_taxonomy(obj["_taxonomy"], "ไฟล์ mapping.json ไม่ถูกต้อง (ส่วน _taxonomy)")
    return out


def _friendly(ex):
    return str(ex) if isinstance(ex, UserError) else "อ่านไฟล์ไม่ได้ (ไม่ใช่ไฟล์ Excel/CSV ที่ถูกต้อง หรือไฟล์เสียหาย)"


def classify_file(name, content):
    """-> 'data' (ข้อมูลดิบ) | 'master' (ไฟล์รวมที่ระบบเคยสร้าง) | 'mapping' | 'unknown' | 'unsupported'"""
    low = name.lower()
    if low.endswith(".xlsx"): check_xlsx_safe(content)               # ต้องเช็กนอก try ไม่งั้นข้อความถูกกลืนเป็น "ไม่รองรับไฟล์"
    try:
        if low.endswith(".json"):
            obj = json.loads(content.decode("utf-8-sig"))
            if isinstance(obj, dict): return "mapping"
            if isinstance(obj, list) and obj and isinstance(obj[0], dict) and _MASTER_COLS <= set(obj[0]): return "master"
            return "unknown"
        if low.endswith(".csv"):
            try: text = content.decode("utf-8-sig")
            except UnicodeDecodeError: text = content.decode("cp874", errors="replace")
            header = next(csv.reader(io.StringIO(text)), [])
            return "master" if _MASTER_COLS <= {h.strip() for h in header} else "data"
        if low.endswith(".xlsx"):
            cols = {str(c) for c in pd.read_excel(io.BytesIO(content), nrows=0).columns}
            return "master" if _MASTER_COLS <= cols else "data"
    except Exception:
        return "unknown"
    return "unsupported"


def _read_master(content, mapping, name):
    low = name.lower()
    if low.endswith(".xlsx"): check_xlsx_safe(content)               # ต้องเช็กนอก try ไม่งั้นข้อความถูกกลืนเป็น "ไม่รองรับไฟล์"
    try:
        if low.endswith(".json"): prior = pd.DataFrame(json.loads(content.decode("utf-8-sig")))
        elif low.endswith(".csv"): prior = pd.read_csv(io.BytesIO(content), encoding="utf-8-sig")
        else: prior = pd.read_excel(io.BytesIO(content))
    except Exception as ex:
        raise UserError(f"{name}: อ่านไฟล์รวมปีก่อนไม่ได้") from ex
    for c in ("year", "months", "position_code"):
        if c in prior: prior[c] = pd.to_numeric(prior[c], errors="coerce").astype("Int64")
    for col in CODE_COLS:                                    # รหัสในไฟล์รวมต้องมีอยู่ใน mapping.json (ไม่งั้นรหัสจะชนกัน)
        known = set(mapping.get(col, {}).values())
        used = set(prior[col].dropna().astype(str)) if col in prior else set()
        if used - known:
            raise UserError(f"{name}: ไฟล์รวมนี้กับ mapping.json ไม่ใช่ชุดเดียวกัน (พบรหัสที่ mapping.json ไม่รู้จัก) "
                            "ต้องใช้คู่ที่ดาวน์โหลดมาพร้อมกัน")
    return prior


def _order(df):
    first = [c for c in PREFERRED_ORDER if c in df.columns]
    return df[first + [c for c in df.columns if c not in first]]


def run_pipeline(files, mapping=None, alias_extra=None, selection=None, keep_original=False, stage="anonymize"):
    """files = [(ชื่อไฟล์, bytes)]  (ไฟล์ข้อมูลดิบ และ/หรือ ไฟล์รวมปีก่อน)
    stage='anonymize' -> แทนรหัส/ตัดข้อมูลส่วนตัว ยังไม่ใส่หมวดตำแหน่ง
    stage='final'     -> ใส่หมวด/กลุ่มตำแหน่ง และต่อกับไฟล์รวมปีก่อน
    selection = รายการตาราง/คอลัมน์ที่ผู้ใช้เลือก (None = อัตโนมัติ)
    ถ้ามีไฟล์ใดพัง จะ raise UserError และไม่คืนอะไรเลย"""
    mapping = validate_mapping(mapping) if mapping else {}
    alias = dict(mapping.get("_alias", {})); alias.update(alias_extra or {})
    cats = effective_categories(mapping)
    valid = {c["code"] for c in cats}
    pos_extra = {k: v for k, v in mapping.get("_position", {}).items() if v in valid}
    sel = parse_selection(selection) if selection is not None else None

    raw_files, masters = [], []
    for name, content in sorted(files, key=lambda x: x[0]):
        kind = classify_file(name, content)
        if kind == "master": masters.append((name, content))
        elif kind == "data": raw_files.append((name, content))
        elif kind == "mapping": continue
        else: raise UserError(f"ไม่รองรับไฟล์ “{name}” (รองรับ .xlsx, .csv และ mapping.json)")
    if not raw_files: raise UserError("ยังไม่ได้เลือกไฟล์ข้อมูลที่จะแปลง")
    if masters and not mapping and stage == "final":
        raise UserError("ไฟล์รวมของปีก่อนต้องใส่ mapping.json ที่คู่กันมาด้วย")

    results, frames, similar = [], [], []
    for name, content in raw_files:
        buf, stats = io.StringIO(), {}
        try:
            with contextlib.redirect_stdout(buf):
                df = process_file(content, mapping, alias, filename=name, selection=sel, stats=stats)
        except UserError as ex:
            raise UserError(f"{name}: {ex}") from ex
        except Exception as ex:
            raise UserError(f"{name}: {_friendly(ex)}") from ex
        if df is None: continue
        for c in ("start", "end"):
            if c in df: df[c] = df[c].replace("NaT", pd.NA)
        if "course" in df: df["course"] = df["course"].map(norm_course)
        df["source"] = safe_stem(name)
        for table, g in df.groupby("table", sort=False):
            bad = int((g["months"].isna() | (g["months"] < 1) | (g["months"] > 12)).sum()) if "months" in g else 0
            sk = stats.get(str(table), {})
            results.append(dict(file=name, table=str(table), rows=int(len(g)), bad_dates=bad,
                                no_id=sk.get("no_id", 0), not_company=sk.get("not_company", 0)))
        frames.append(df)
        log = buf.getvalue()
        similar += [dict(name=m.group(1), near=m.group(2), file=name) for m in map(_SIMILAR.search, log.splitlines()) if m]

    if not frames: raise UserError("ไม่ได้เลือกตารางที่จะแปลง")
    titles = titles_table(frames, pos_extra)
    prior_rows = 0
    if stage == "final":
        frames = [add_position_columns(f, pos_extra, cats) for f in frames]
    if not keep_original:
        frames = [f.drop(columns=[c for c in ("position", "position_2") if c in f]) for f in frames]
    dataset = pd.concat(frames, ignore_index=True)
    if stage == "final" and masters:
        by = {c["code"]: c for c in cats}
        new_sources = set(dataset["source"].dropna())
        kept = []
        for name, content in masters:
            p = _read_master(content, mapping, name)
            p = p[~p["source"].isin(new_sources)].copy()
            if "position_code" in p:                           # จัดหมวด/กลุ่มของแถวเก่าใหม่ตามรายการหมวดปัจจุบัน
                p["position_category"] = p["position_code"].map(lambda c: by[int(c)]["name"] if pd.notna(c) and int(c) in by else pd.NA)
                p["position_group"] = p["position_code"].map(lambda c: by[int(c)]["group"] if pd.notna(c) and int(c) in by else pd.NA)
            kept.append(p)
        prior_rows = int(sum(len(p) for p in kept))
        dataset = pd.concat(kept + [dataset], ignore_index=True)
    for c in ("year", "months", "position_code"):
        if c in dataset: dataset[c] = pd.to_numeric(dataset[c], errors="coerce").astype("Int64")
    dataset = _order(dataset)

    mapping["_alias"] = alias
    mapping["_position"] = pos_extra
    return dict(dataset=dataset, results=results, similar=similar, mapping=mapping, titles=titles, categories=cats, prior_rows=prior_rows,
                counts=_code_counts(dataset))      # จำนวนรหัสที่ใช้จริงในข้อมูลชุดนี้


def _code_counts(ds):
    out = {}
    for k in CODE_COLS:
        vals = set(ds[k].dropna()) if k in ds else set()
        if k == "company" and "company_2" in ds: vals |= set(ds["company_2"].dropna())
        out[k] = len(vals)
    return out


# ---------- ไฟล์ผลลัพธ์ ----------
_FORMULA_START = ("=", "+", "-", "@", "\t", "\r")


def _neutralize(v):
    """ข้อความที่ขึ้นต้นด้วย = + - @ จะถูก Excel ตีเป็นสูตร (รันคำสั่ง/ดึงข้อมูลได้) -> ใส่ ' นำหน้าให้เป็นข้อความธรรมดา"""
    return "'" + v if isinstance(v, str) and v.startswith(_FORMULA_START) else v


def csv_bytes(df):
    safe = df.copy()
    for c in safe.columns:
        if safe[c].dtype == object or str(safe[c].dtype) in ("string", "str"): safe[c] = safe[c].map(_neutralize)
    return safe.to_csv(index=False).encode("utf-8-sig")           # BOM ให้ Excel เปิดภาษาไทยถูก


def json_bytes(df):
    recs = json.loads(df.to_json(orient="records", force_ascii=False))
    return json.dumps(recs, ensure_ascii=False, indent=1).encode("utf-8")


def mapping_bytes(mapping):
    return json.dumps(mapping, ensure_ascii=False, indent=2).encode("utf-8")


def preview_records(df, n=8):
    return json.loads(df.head(n).to_json(orient="records", force_ascii=False))
