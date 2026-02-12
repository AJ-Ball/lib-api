from __future__ import annotations

import re
from typing import Optional, List, Dict, Any

import pandas as pd
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware

EXCEL_PATH = "Data_Lib.xlsx"
SHEET_NAME = "api"

app = FastAPI(title="Library Locator API", version="1.0.0")

# เปิด CORS ไว้ก่อน (Jotform มักเรียกจาก client/webhook ได้หลายแบบ)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # ปรับให้เข้มงวดทีหลังได้
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

_df: Optional[pd.DataFrame] = None


def load_data() -> pd.DataFrame:
    """Load once and keep in memory."""
    global _df
    if _df is not None:
        return _df

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

    # ทำความสะอาดชนิดข้อมูลที่จำเป็น
    num_cols = ["range_start_num", "range_end_num", "row", "shelf_level", "locker", "building_floor"]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # เติม suffix ให้เป็น string ว่างแทน NaN
    for c in ["range_start_suffix", "range_end_suffix"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()

    # เติมข้อความอื่น ๆ
    for c in ["id", "side", "category", "call_range", "map_url", "range_start_raw", "range_end_raw"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()

    # ตัดแถวที่ตัวเลขช่วงไม่ครบ
    df = df.dropna(subset=["range_start_num", "range_end_num"]).copy()

    _df = df
    return _df


THAI_SUFFIX_RE = re.compile(r"(?P<num>[0-9.]+)\s*(?P<suffix>[ก-๙]{0,3})$", re.UNICODE)


def normalize_call_number(raw: str) -> tuple[Optional[float], str]:
    """
    รับอินพุตได้เช่น:
      - "370.113"
      - "370113"
      - "370.113พ"
      - "370113พ"
      - "370.1ศ."
    คืนค่า: (num_float, suffix_str)
    """
    if not raw:
        return None, ""

    s = str(raw).strip()
    # ลบช่องว่าง/ขีด/จุดท้าย ๆ ที่เป็น noise
    s = s.replace("–", "-").replace("—", "-")
    s = s.strip()

    m = THAI_SUFFIX_RE.search(s)
    if not m:
        return None, ""

    num_part = m.group("num").strip()
    suffix = (m.group("suffix") or "").strip()

    # ถ้า num_part ไม่มีจุด ให้พยายามใส่จุดหลัง 3 หลักแรก
    # ตัวอย่าง: 370113 -> 370.113, 3701 -> 370.1
    if "." not in num_part:
        digits = re.sub(r"\D", "", num_part)
        if len(digits) <= 3:
            # เช่น "370" -> 370.0
            try:
                return float(digits), suffix
            except Exception:
                return None, suffix
        else:
            # ใส่จุดหลัง 3 หลักแรก
            num_part = digits[:3] + "." + digits[3:]

    try:
        num = float(num_part)
    except Exception:
        return None, suffix

    # ปัด 3 ตำแหน่งให้ใกล้เคียงข้อมูลในไฟล์ (ที่มี 3 decimals)
    num = round(num, 3)
    return num, suffix


def suffix_ge(a: str, b: str) -> bool:
    """a >= b สำหรับ suffix ไทย (ใช้ unicode order)"""
    return a >= b


def suffix_le(a: str, b: str) -> bool:
    """a <= b สำหรับ suffix ไทย (ใช้ unicode order)"""
    return a <= b


def match_row(row: pd.Series, q_num: float, q_suffix: str, strict_suffix: bool) -> bool:
    """
    strict_suffix:
      - True  = ถ้าผู้ใช้พิมพ์ suffix มา ให้เช็คขอบเขต suffix ตอนเลขเท่ากับขอบช่วง
      - False = ถ้าผู้ใช้ไม่พิมพ์ suffix มา จะใช้เฉพาะเลข (inclusive) ไม่สน suffix
    """
    start_n = row["range_start_num"]
    end_n = row["range_end_num"]

    if q_num < start_n or q_num > end_n:
        return False

    if not strict_suffix:
        return True

    # ถ้าเลขอยู่กลางช่วง (ไม่ชนขอบ) ก็ผ่าน
    if q_num > start_n and q_num < end_n:
        return True

    # ถ้าเลขชนขอบ ให้พิจารณา suffix
    start_s = row.get("range_start_suffix", "") or ""
    end_s = row.get("range_end_suffix", "") or ""

    # ชนขอบต้น
    if q_num == start_n and start_s:
        if not suffix_ge(q_suffix, start_s):
            return False

    # ชนขอบท้าย
    if q_num == end_n and end_s:
        if not suffix_le(q_suffix, end_s):
            return False

    return True


def rank_candidates(df: pd.DataFrame, q_num: float) -> pd.DataFrame:
    """จัดอันดับให้ช่วงแคบกว่าอยู่ก่อน (ช่วยลดความกำกวม)"""
    d = df.copy()
    d["span"] = (d["range_end_num"] - d["range_start_num"]).abs()
    d = d.sort_values(["span", "row", "shelf_level", "locker"], ascending=[True, True, True, True])
    return d


@app.get("/health")
def health() -> Dict[str, Any]:
    df = load_data()
    return {"ok": True, "rows": int(df.shape[0])}


@app.get("/search")
def search(
    q: str = Query(..., description="เลขหมวดหรือคำค้น เช่น 370.113พ หรือ 370113 หรือ สังคมศาสตร์"),
    limit: int = Query(5, ge=1, le=20),
) -> Dict[str, Any]:
    df = load_data()

    q = (q or "").strip()

    # 1) พยายามตีความเป็น call number ก่อน
    q_num, q_suffix = normalize_call_number(q)

    if q_num is not None:
        strict_suffix = bool(q_suffix)  # ถ้ามี suffix ค่อย strict
        mask = df.apply(lambda r: match_row(r, q_num, q_suffix, strict_suffix), axis=1)
        hits = df[mask].copy()

        if hits.empty:
            return {
                "found": False,
                "mode": "call_number",
                "query": q,
                "normalized": {"num": q_num, "suffix": q_suffix},
                "results": [],
                "suggest": [
                    "ลองพิมพ์เฉพาะตัวเลข เช่น 370.113 หรือ 370113",
                    "ถ้ามีตัวอักษรท้ายเลข ลองใส่ด้วย เช่น 370.113พ",
                    "หรือค้นด้วยคำหมวด เช่น สังคมศาสตร์",
                ],
            }

        hits = rank_candidates(hits, q_num).head(limit)

        results = []
        for _, r in hits.iterrows():
            results.append(
                {
                    "id": r.get("id", ""),
                    "call_range": r.get("call_range", ""),
                    "category": r.get("category", ""),
                    "location": {
                        "row": int(r.get("row", 0)) if pd.notna(r.get("row", None)) else None,
                        "shelf_level": int(r.get("shelf_level", 0)) if pd.notna(r.get("shelf_level", None)) else None,
                        "locker": int(r.get("locker", 0)) if pd.notna(r.get("locker", None)) else None,
                        "building_floor": int(r.get("building_floor", 0)) if pd.notna(r.get("building_floor", None)) else None,
                        "side": r.get("side", ""),
                    },
                    "range": {
                        "start_raw": r.get("range_start_raw", ""),
                        "end_raw": r.get("range_end_raw", ""),
                        "start_num": float(r.get("range_start_num", 0.0)),
                        "end_num": float(r.get("range_end_num", 0.0)),
                        "start_suffix": r.get("range_start_suffix", ""),
                        "end_suffix": r.get("range_end_suffix", ""),
                    },
                    "map_url": r.get("map_url", ""),
                }
            )

        return {
            "found": True,
            "mode": "call_number",
            "query": q,
            "normalized": {"num": q_num, "suffix": q_suffix},
            "count": len(results),
            "results": results,
        }

    # 2) ถ้าไม่ใช่เลขหมวด ให้ค้นแบบข้อความ (category / call_range / id)
    q_low = q.lower()
    text_mask = (
        df["category"].str.lower().str.contains(q_low, na=False)
        | df["call_range"].str.lower().str.contains(q_low, na=False)
        | df["id"].str.lower().str.contains(q_low, na=False)
    )
    hits = df[text_mask].copy()

    if hits.empty:
        return {
            "found": False,
            "mode": "text",
            "query": q,
            "results": [],
            "suggest": [
                "ลองค้นด้วยเลขหมวด เช่น 370.113 หรือ 370113",
                "หรือค้นด้วยคำหมวด เช่น สังคมศาสตร์ / วิทยาศาสตร์",
            ],
        }

    hits = hits.head(limit)
    results = []
    for _, r in hits.iterrows():
        results.append(
            {
                "id": r.get("id", ""),
                "call_range": r.get("call_range", ""),
                "category": r.get("category", ""),
                "location": {
                    "row": int(r.get("row", 0)) if pd.notna(r.get("row", None)) else None,
                    "shelf_level": int(r.get("shelf_level", 0)) if pd.notna(r.get("shelf_level", None)) else None,
                    "locker": int(r.get("locker", 0)) if pd.notna(r.get("locker", None)) else None,
                    "building_floor": int(r.get("building_floor", 0)) if pd.notna(r.get("building_floor", None)) else None,
                    "side": r.get("side", ""),
                },
                "map_url": r.get("map_url", ""),
            }
        )

    return {
        "found": True,
        "mode": "text",
        "query": q,
        "count": len(results),
        "results": results,
    }
