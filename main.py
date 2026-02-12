from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import re
import os
from typing import Optional, Tuple, Any, Dict, List

app = FastAPI(title="Library Locator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

EXCEL_PATH = "Data_Lib.xlsx"
SHEET_NAME = "api"

data_cache: Optional[List[Dict[str, Any]]] = None


# ---------- Utilities ----------

def safe_str(x) -> str:
    return "" if x is None else str(x)

def to_key(num_str: str) -> Optional[int]:
    """
    370.113 -> 370113
    370113 -> 370113
    """
    num_str = safe_str(num_str).strip()
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)", num_str)
    if not m:
        return None
    v = float(m.group(1))
    if v < 1000:
        return int(round(v * 1000))
    return int(round(v))

def parse_call_number(raw: str) -> Tuple[Optional[int], str]:
    """
    370.1ศ. -> (370100, ศ)
    370.1อ  -> (370100, อ)
    370.111ค -> (370111, ค)
    """
    raw = safe_str(raw).strip()
    if not raw:
        return None, ""

    # เอาเฉพาะรูป "เลข + อักษรไทยท้าย (ถ้ามี)"
    m = re.match(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*([ก-๙]{0,3})", raw)
    if not m:
        return to_key(raw), ""

    key = to_key(m.group(1))
    suf = m.group(2) or ""
    return key, suf

def parse_range(range_text: str):
    """
    "370.1ศ.-370.1อ" => ((370100,'ศ'), (370100,'อ'))
    """
    s = safe_str(range_text).strip()
    if not s:
        return None
    s = s.replace("–", "-").replace("—", "-")
    parts = [p.strip() for p in s.split("-") if p.strip()]
    if len(parts) != 2:
        return None

    a_key, a_suf = parse_call_number(parts[0])
    b_key, b_suf = parse_call_number(parts[1])
    if a_key is None or b_key is None:
        return None

    return (a_key, a_suf), (b_key, b_suf)

def find_header(headers: List[str], target: str) -> Optional[str]:
    """
    หา header ที่ใกล้เคียง target (กันปัญหาช่องว่าง/บรรทัดใหม่)
    """
    t = re.sub(r"\s+", "", target)
    for h in headers:
        if re.sub(r"\s+", "", h) == t:
            return h
    return None


# ---------- Load Excel ----------

def load_data():
    global data_cache
    if data_cache is not None:
        return data_cache

    if not os.path.exists(EXCEL_PATH):
        raise RuntimeError(f"Excel file not found: {EXCEL_PATH}")

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(f"Sheet '{SHEET_NAME}' not found. Available: {wb.sheetnames}")

    ws = wb[SHEET_NAME]

    headers_raw = [safe_str(c.value) for c in ws[1]]
    headers = [h.strip() for h in headers_raw]

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        rows.append(d)

    wb.close()
    data_cache = rows
    return rows


@app.get("/health")
def health():
    rows = load_data()
    return {"ok": True, "rows": len(rows), "file": EXCEL_PATH, "sheet": SHEET_NAME}


@app.get("/debug")
def debug():
    """
    เอาไว้ดู header จริง ๆ ในไฟล์ (แก้ 500 ได้ไวมาก)
    """
    if not os.path.exists(EXCEL_PATH):
        return {"ok": False, "error": f"Excel file not found: {EXCEL_PATH}"}

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    info = {"sheets": wb.sheetnames}
    if SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
        headers = [safe_str(c.value) for c in ws[1]]
        info["headers"] = headers
    wb.close()
    return {"ok": True, **info}


@app.get("/search")
def search(q: str = Query(...)):
    rows = load_data()

    # อ่าน header ที่ต้องใช้แบบ robust
    headers = list(rows[0].keys()) if rows else []
    col_range = find_header(headers, "ช่องหมวดหนังสือ")
    col_row = find_header(headers, "แถว")
    col_shelf = find_header(headers, "ชั้นที่")
    col_locker = find_header(headers, "ล็อคที่")
    col_floor = find_header(headers, "ในอาคารชั้นที่")
    col_side = find_header(headers, "ด้าน")
    col_group = find_header(headers, "กลุ่ม")
    col_map = find_header(headers, "urlแผนที่")

    missing = [name for name, col in [
        ("ช่องหมวดหนังสือ", col_range),
        ("แถว", col_row),
        ("ชั้นที่", col_shelf),
        ("ล็อคที่", col_locker),
        ("ในอาคารชั้นที่", col_floor),
        ("ด้าน", col_side),
        ("กลุ่ม", col_group),
        ("urlแผนที่", col_map),
    ] if col is None]

    if missing:
        return {
            "found": False,
            "error": "Missing columns in Excel",
            "missing": missing,
            "hint": "เปิด /debug เพื่อดูชื่อคอลัมน์จริง แล้วแก้ชื่อให้ตรง",
        }

    q_key, q_suf = parse_call_number(q)
    if q_key is None:
        return {"found": False, "query": q, "error": "Invalid query format"}

    # match
    matches = []
    for r in rows:
        range_text = r.get(col_range)
        parsed = parse_range(range_text)
        if not parsed:
            continue

        (a_key, a_suf), (b_key, b_suf) = parsed

        if a_key <= q_key <= b_key:
            matches.append(r)

    if not matches:
        return {"found": False, "query": q, "count": 0}

    r = matches[0]
    message = (
        f"พบหนังสือ แถว {r.get(col_row)} ชั้น {r.get(col_shelf)} ล็อค {r.get(col_locker)} "
        f"อาคารชั้น {r.get(col_floor)} ด้าน{r.get(col_side)} กลุ่ม {r.get(col_group)}"
    )

    return {
        "found": True,
        "query": q,
        "count": len(matches),
        "message": message,
        "location": {
            "row": r.get(col_row),
            "shelf": r.get(col_shelf),
            "locker": r.get(col_locker),
            "floor": r.get(col_floor),
            "side": r.get(col_side),
            "group": r.get(col_group),
        },
        "map_url": r.get(col_map),
    }
