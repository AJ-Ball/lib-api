from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import re

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

data_cache = None


# ---------- Utilities ----------

def to_key(num_str: str):
    """
    370.113 -> 370113
    370113 -> 370113
    """
    num_str = str(num_str)
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)", num_str)
    if not m:
        return None
    v = float(m.group(1))
    if v < 1000:
        return int(v * 1000)
    return int(v)


def parse_call_number(raw: str):
    """
    370.1ศ. -> (370100, ศ)
    """
    raw = raw.strip()
    m = re.search(r"([0-9.]+)\s*([ก-๙]*)", raw)
    if not m:
        return None, ""
    key = to_key(m.group(1))
    suffix = m.group(2)
    return key, suffix


def parse_range(range_text: str):
    """
    370.1ศ.-370.1อ
    """
    range_text = range_text.replace("–", "-").replace("—", "-")
    parts = [p.strip() for p in range_text.split("-")]
    if len(parts) != 2:
        return None, None

    start_key, start_suf = parse_call_number(parts[0])
    end_key, end_suf = parse_call_number(parts[1])

    return (start_key, start_suf), (end_key, end_suf)


# ---------- Load Excel ----------

def load_data():
    global data_cache
    if data_cache:
        return data_cache

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    headers = [c.value for c in ws[1]]
    rows = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        rows.append(d)

    wb.close()
    data_cache = rows
    return rows


# ---------- API ----------

@app.get("/health")
def health():
    rows = load_data()
    return {"ok": True, "rows": len(rows)}


@app.get("/search")
def search(q: str = Query(...)):
    rows = load_data()
    q_key, q_suffix = parse_call_number(q)

    if q_key is None:
        return {"found": False, "message": "รูปแบบเลขไม่ถูกต้อง"}

    matches = []

    for r in rows:
        range_text = r.get("ช่องหมวดหนังสือ", "")
        parsed = parse_range(range_text)
        if not parsed:
            continue

        (start_key, start_suf), (end_key, end_suf) = parsed

        if start_key is None or end_key is None:
            continue

        # เช็คว่าอยู่ในช่วง
        if start_key <= q_key <= end_key:
            matches.append(r)

    if not matches:
        return {"found": False, "query": q}

    # เลือกแถวแรก (ช่วงของคุณจัดเรียงอยู่แล้ว)
    r = matches[0]

    message = f"""
พบหนังสือ
แถว {r['แถว']}
ชั้น {r['ชั้นที่']}
ล็อค {r['ล็อคที่']}
อาคารชั้น {r['ในอาคารชั้นที่']}
ด้าน{r['ด้าน']}
กลุ่ม {r['กลุ่ม']}
""".strip()

    return {
        "found": True,
        "query": q,
        "message": message,
        "location": {
            "row": r["แถว"],
            "shelf": r["ชั้นที่"],
            "locker": r["ล็อคที่"],
            "floor": r["ในอาคารชั้นที่"],
            "side": r["ด้าน"],
            "group": r["กลุ่ม"],
        },
        "map_url": r["urlแผนที่"],
    }
