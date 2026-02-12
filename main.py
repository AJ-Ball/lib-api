from __future__ import annotations

import re
from typing import Optional, Dict, Any, List

from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook


EXCEL_PATH = "Data_Lib.xlsx"
SHEET_NAME = "api"

app = FastAPI(title="Library Locator API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

_rows: Optional[List[Dict[str, Any]]] = None


def to_float(x) -> Optional[float]:
    if x is None:
        return None
    try:
        return float(x)
    except Exception:
        return None


def to_int(x) -> Optional[int]:
    if x is None:
        return None
    try:
        return int(float(x))
    except Exception:
        return None


def load_data() -> List[Dict[str, Any]]:
    global _rows
    if _rows is not None:
        return _rows

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    # header row
    headers = []
    for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True)):
        headers.append(str(cell).strip() if cell is not None else "")

    rows: List[Dict[str, Any]] = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = {}
        for k, v in zip(headers, r):
            if not k:
                continue
            d[k] = v

        # normalize expected fields
        d["range_start_num"] = to_float(d.get("range_start_num"))
        d["range_end_num"] = to_float(d.get("range_end_num"))
        d["row"] = to_int(d.get("row"))
        d["shelf_level"] = to_int(d.get("shelf_level"))
        d["locker"] = to_int(d.get("locker"))
        d["building_floor"] = to_int(d.get("building_floor"))

        # strings
        for c in ["id", "side", "category", "call_range", "map_url", "range_start_raw", "range_end_raw", "range_start_suffix", "range_end_suffix"]:
            if c in d and d[c] is not None:
                d[c] = str(d[c]).strip()
            else:
                d[c] = ""

        # keep only rows with numeric range
        if d["range_start_num"] is None or d["range_end_num"] is None:
            continue

        rows.append(d)

    wb.close()
    _rows = rows
    return _rows


THAI_SUFFIX_RE = re.compile(r"(?P<num>[0-9.]+)\s*(?P<suffix>[ก-๙]{0,3})$", re.UNICODE)


def normalize_call_number(raw: str) -> tuple[Optional[float], str]:
    if not raw:
        return None, ""

    s = str(raw).strip().replace("–", "-").replace("—", "-")
    m = THAI_SUFFIX_RE.search(s)
    if not m:
        return None, ""

    num_part = m.group("num").strip()
    suffix = (m.group("suffix") or "").strip()

    if "." not in num_part:
        digits = re.sub(r"\D", "", num_part)
        if len(digits) <= 3:
            try:
                return round(float(digits), 3), suffix
            except Exception:
                return None, suffix
        num_part = digits[:3] + "." + digits[3:]

    try:
        num = round(float(num_part), 3)
    except Exception:
        return None, suffix

    return num, suffix


def match_row(d: Dict[str, Any], q_num: float, q_suffix: str, strict_suffix: bool) -> bool:
    start_n = d["range_start_num"]
    end_n = d["range_end_num"]

    if q_num < start_n or q_num > end_n:
        return False

    if not strict_suffix:
        return True

    if q_num > start_n and q_num < end_n:
        return True

    start_s = d.get("range_start_suffix", "") or ""
    end_s = d.get("range_end_suffix", "") or ""

    if q_num == start_n and start_s:
        if q_suffix < start_s:
            return False

    if q_num == end_n and end_s:
        if q_suffix > end_s:
            return False

    return True


def rank_candidates(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    def span(d):
        return abs((d["range_end_num"] or 0) - (d["range_start_num"] or 0))

    return sorted(items, key=lambda d: (span(d), d.get("row") or 10**9, d.get("shelf_level") or 10**9, d.get("locker") or 10**9))


@app.get("/health")
def health() -> Dict[str, Any]:
    data = load_data()
    return {"ok": True, "rows": len(data)}


@app.get("/search")
def search(
    q: str = Query(..., description="เลขหมวดหรือคำค้น เช่น 370.113พ หรือ 370113 หรือ สังคมศาสตร์"),
    limit: int = Query(5, ge=1, le=20),
) -> Dict[str, Any]:
    data = load_data()
    q = (q or "").strip()

    q_num, q_suffix = normalize_call_number(q)

    # 1) call number search
    if q_num is not None:
        strict_suffix = bool(q_suffix)
        hits = [d for d in data if match_row(d, q_num, q_suffix, strict_suffix)]

        if not hits:
            return {
                "found": False,
                "mode": "call_number",
                "query": q,
                "normalized": {"num": q_num, "suffix": q_suffix},
                "results": [],
            }

        hits = rank_candidates(hits)[:limit]
        results = []
        for d in hits:
            results.append({
                "id": d.get("id", ""),
                "call_range": d.get("call_range", ""),
                "category": d.get("category", ""),
                "location": {
                    "row": d.get("row"),
                    "shelf_level": d.get("shelf_level"),
                    "locker": d.get("locker"),
                    "building_floor": d.get("building_floor"),
                    "side": d.get("side", ""),
                },
                "range": {
                    "start_raw": d.get("range_start_raw", ""),
                    "end_raw": d.get("range_end_raw", ""),
                    "start_num": d.get("range_start_num"),
                    "end_num": d.get("range_end_num"),
                    "start_suffix": d.get("range_start_suffix", ""),
                    "end_suffix": d.get("range_end_suffix", ""),
                },
                "map_url": d.get("map_url", ""),
            })

        return {
            "found": True,
            "mode": "call_number",
            "query": q,
            "normalized": {"num": q_num, "suffix": q_suffix},
            "count": len(results),
            "results": results,
        }

    # 2) text search
    q_low = q.lower()
    def s(x): return (x or "").lower()

    hits = [d for d in data if (q_low in s(d.get("category")) or q_low in s(d.get("call_range")) or q_low in s(d.get("id")))]
    if not hits:
        return {"found": False, "mode": "text", "query": q, "results": []}

    hits = hits[:limit]
    results = []
    for d in hits:
        results.append({
            "id": d.get("id", ""),
            "call_range": d.get("call_range", ""),
            "category": d.get("category", ""),
            "location": {
                "row": d.get("row"),
                "shelf_level": d.get("shelf_level"),
                "locker": d.get("locker"),
                "building_floor": d.get("building_floor"),
                "side": d.get("side", ""),
            },
            "map_url": d.get("map_url", ""),
        })

    return {"found": True, "mode": "text", "query": q, "count": len(results), "results": results}
