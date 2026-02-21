# ============================================================
# SLA LOGIC FILE (Cloud Compatible - NO tkinter)
# ============================================================

import os
import re
import math
import calendar
from datetime import datetime

import pandas as pd
import numpy as np


# -----------------------------
# Helpers
# -----------------------------
def safe_float(x, default=np.nan):
    try:
        s = str(x).strip()
        if s == "":
            return default
        return float(s)
    except Exception:
        return default


def norm_route_name(s):
    if pd.isna(s):
        return ""
    t = str(s).strip().lower()
    t = t.replace("\u00a0", " ")
    t = t.replace("–", "-").replace("—", "-")
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"\s*-\s*", "-", t)
    return t.strip()


def sanitize_filename(s, max_len=60):
    s = "" if s is None else str(s)
    s = s.strip()
    s = re.sub(r"[\\/:*?\"<>|]", "_", s)
    s = re.sub(r"\s+", " ", s)
    s = s.strip(" .-_")
    if len(s) == 0:
        s = "Unknown"
    return s[:max_len]


def ensure_engine(path):
    ext = os.path.splitext(path.lower())[1]
    if ext == ".xls":
        return "xlrd"
    return None


def read_excel_any(path, header=0):
    engine = ensure_engine(path)
    if engine:
        return pd.read_excel(path, header=header, engine=engine)
    return pd.read_excel(path, header=header)


def parse_duration_to_hours(val):
    if pd.isna(val):
        return np.nan
    if isinstance(val, pd.Timedelta):
        return val.total_seconds() / 3600.0

    s = str(val).strip()
    if s == "":
        return np.nan

    if re.match(r"^\d{1,4}:\d{2}(:\d{2})?$", s):
        parts = s.split(":")
        h = int(parts[0])
        m = int(parts[1])
        sec = int(parts[2]) if len(parts) == 3 else 0
        return h + (m / 60.0) + (sec / 3600.0)

    try:
        return float(s)
    except Exception:
        return np.nan


def uptime_deduction_pct(uptime_pct):
    if pd.isna(uptime_pct):
        return 0
    if uptime_pct >= 99:
        return 0
    if uptime_pct >= 98:
        return 10
    if uptime_pct >= 97:
        return 25
    if uptime_pct >= 96:
        return 50
    if uptime_pct >= 95:
        return 75
    return 100


def mttr_penalty_non_cumulative(duration_hours):
    if pd.isna(duration_hours) or duration_hours <= 0:
        return 0, "Invalid/Blank"

    h = int(math.ceil(float(duration_hours)))

    if h <= 4:
        return 0, "≤4"
    elif h <= 6:
        return 500, ">4–6"
    elif h <= 24:
        return int(500 + 100 * int(math.ceil(h - 6))), ">6–24"
    elif h <= 48:
        return 5000, ">24–48"
    else:
        extra_days = int(math.ceil((h - 48) / 24))
        return int(5000 + 500 * extra_days), ">48"


def pan_4th_digit_to_tds_rate(pan4):
    if pan4 is None:
        return None
    s = str(pan4).strip().upper()
    if s == "":
        return None
    if s in ["P", "H"]:
        return 0.01
    return 0.02


def fmt_money(x):
    try:
        return f"{float(x):,.2f}"
    except Exception:
        return "0.00"


# ============================================================
# MAIN PROCESS FUNCTION (UNCHANGED LOGIC)
# ============================================================

def process_sla(
    annex_a_path,
    annex_c_path,
    rate_per_km,
    save_dir,
    vendor_basic_value=None,
    pan4=None,
    field_unit_penalty=0.0,
    vendor_deducted_penalty=0.0,
    other_recovery=0.0,
    splice_loss_amt=0.0,
    supervisor_abs_amt=0.0,
    frt_abs_amt=0.0,
    petroller_abs_amt=0.0,
    relaying_not_done_amt=0.0,
):

    # === FORMAT A ===
    a = read_excel_any(annex_a_path)
    a.columns = [str(c).strip() for c in a.columns]

    required = [
        "FORMAT", "BA", "OA", "Month", "Sr.No.",
        "Transnet Route ID", "Working Route Name as per Transnet",
        "RKM", "Name of Maintenance Agency"
    ]
    missing = [c for c in required if c not in a.columns]
    if missing:
        raise ValueError(f"Format A missing columns: {missing}")

    routes = a[required].copy()
    routes.rename(columns={
        "Sr.No.": "Sl_No",
        "Transnet Route ID": "Route_ID",
        "Working Route Name as per Transnet": "Route_Name",
        "RKM": "Route_KM",
        "Name of Maintenance Agency": "Vendor_Name"
    }, inplace=True)

    routes["Route_KM"] = pd.to_numeric(routes["Route_KM"], errors="coerce").fillna(0.0)
    routes["SLA_Charges_Rs"] = routes["Route_KM"] * float(rate_per_km)

    # === FORMAT C ===
    c = read_excel_any(annex_c_path)
    c.columns = [str(col).strip() for col in c.columns]

    duration_col = None
    for col in c.columns:
        if "fault" in col.lower() and "duration" in col.lower():
            duration_col = col
            break

    if duration_col is None:
        duration_col = c.columns[13]

    c["Duration_Hrs"] = c[duration_col].apply(parse_duration_to_hours)
    faults = c[c["Duration_Hrs"].notna() & (c["Duration_Hrs"] > 0)].copy()

    penalties = faults["Duration_Hrs"].apply(mttr_penalty_non_cumulative)
    faults["MTTR_Penalty"] = [p[0] for p in penalties]

    total_penalty = faults["MTTR_Penalty"].sum()

    # === OUTPUT FILES ===
    os.makedirs(save_dir, exist_ok=True)

    excel_path = os.path.join(save_dir, "SLA_Output.xlsx")
    txt1 = os.path.join(save_dir, "Accounts_Note.txt")
    txt2 = os.path.join(save_dir, "Clause_14_1.txt")

    faults.to_excel(excel_path, index=False)

    with open(txt1, "w") as f:
        f.write(f"Total MTTR Penalty: Rs. {fmt_money(total_penalty)}")

    with open(txt2, "w") as f:
        f.write(f"SLA Penalty Calculation\nTotal: Rs. {fmt_money(total_penalty)}")

    return excel_path, txt1, txt2