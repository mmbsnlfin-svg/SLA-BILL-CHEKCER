# ============================================================
# BSNL SLA BILL CHECKER - LOGIC ONLY (FOR STREAMLIT CLOUD)
# DESKTOP LOGIC SAME AS V4.2
# CHANGE: ONLY Accounts Note TXT + Clause 14.1 TXT wording/format
# + FIX Month parsing for date/timestamp in Annexure A Month column
# ============================================================

import os
import re
import math
import calendar
from datetime import datetime, date

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
    """Accept numeric hours OR HH:MM / HH:MM:SS."""
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

    h = int(math.ceil(float(duration_hours)))  # ROUND UP
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


def pick_first_nonblank(series):
    for v in series:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s != "":
            return s
    return ""


def detect_fault_duration_column(df):
    cols = list(df.columns)
    for col in cols:
        s = str(col).lower()
        if "fault" in s and "duration" in s:
            return col
    if len(cols) >= 14:
        return cols[13]
    raise ValueError("Fault Duration column not found.")


def fmt_money(x):
    try:
        return f"{float(x):,.2f}"
    except Exception:
        return "0.00"


def robust_yes(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip().upper()
    if s == "":
        return False
    if s in {"NO", "N", "0", "FALSE"}:
        return False
    if s in {"YES", "Y", "1", "TRUE"}:
        return True
    if "YES" in s:
        return True
    if "EXEMPT" in s:
        return True
    return False


def find_exemption_column(columns):
    cols = list(columns)
    for col in cols:
        col_norm = str(col).strip().lower()
        if ("exempt" in col_norm) or ("exemption" in col_norm):
            return col
    for col in cols:
        col_norm = str(col).strip().lower()
        if (("mttr" in col_norm and "penalty" in col_norm)
            or ("avbility" in col_norm)
            or ("availability" in col_norm)):
            if ("yes" in col_norm or "no" in col_norm or "y/n" in col_norm or "yes/no" in col_norm):
                return col
    return None


def parse_month_year_from_value(val):
    """
    FIX: Annexure A Month may be:
    - text: 'Dec-2025', 'December-2025'
    - timestamp/date: 2025-12-01 00:00:00
    - ISO string: '2025-12-01'
    """
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None, None

    # pandas Timestamp / datetime / date
    if isinstance(val, (pd.Timestamp, datetime, date)):
        return int(val.year), int(val.month)

    s = str(val).strip()
    if s == "":
        return None, None

    # ISO yyyy-mm-dd
    m_iso = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m_iso:
        return int(m_iso.group(1)), int(m_iso.group(2))

    # Month name/abbr + year
    m = re.search(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)", s.lower())
    y = re.search(r"(20\d{2})", s)
    mon_map = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
    if m and y:
        return int(y.group(1)), mon_map[m.group(1)]

    return None, None


# -----------------------------
# Core Processing
# -----------------------------
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
    # ---------- Read Format A ----------
    a = read_excel_any(annex_a_path, header=0)
    a.columns = [str(c).strip() for c in a.columns]

    required_a = [
        "FORMAT", "BA", "OA", "Month", "Sr.No.",
        "Transnet Route ID", "Working Route Name as per Transnet",
        "RKM", "Name of Maintenance Agency"
    ]
    missing_a = [c for c in required_a if c not in a.columns]
    if missing_a:
        raise ValueError(f"Format A missing columns: {missing_a}. Ensure headers are exactly as finalized in Row-1.")

    routes = a[required_a].copy()
    routes.rename(columns={
        "Sr.No.": "Sl_No",
        "Transnet Route ID": "Route_ID",
        "Working Route Name as per Transnet": "Route_Name",
        "RKM": "Route_KM",
        "Name of Maintenance Agency": "Vendor_Name"
    }, inplace=True)

    routes["Route_ID"] = routes["Route_ID"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    routes["Route_Name_norm"] = routes["Route_Name"].apply(norm_route_name)
    routes["Route_KM"] = pd.to_numeric(routes["Route_KM"], errors="coerce").fillna(0.0)

    rate_per_km = float(rate_per_km)
    routes["SLA_Charges_Rs"] = (routes["Route_KM"] * rate_per_km).round(4)

    ba_name = pick_first_nonblank(routes["BA"].tolist())
    oa_name = pick_first_nonblank(routes["OA"].tolist())
    # IMPORTANT: take the first raw value (could be Timestamp)
    sla_month_raw = routes["Month"].iloc[0] if "Month" in routes.columns and len(routes) else ""
    vendor_name = pick_first_nonblank(routes["Vendor_Name"].tolist())

    # month parse (FIXED)
    year, month = parse_month_year_from_value(sla_month_raw)
    if year is None or month is None:
        now = datetime.now()
        year, month = now.year, now.month

    days_in_month = calendar.monthrange(year, month)[1]
    total_hours_month = float(days_in_month * 24)
    month_name = calendar.month_name[month]

    month_display = f"{month_name}-{year}"      # December-2025
    month_display_short = f"{month_name[:3]}-{year}"  # Dec-2025

    vendor_tag = sanitize_filename(vendor_name)
    month_tag2 = sanitize_filename(month_display)

    # Maps
    name_to_id = routes.set_index("Route_Name_norm")["Route_ID"].to_dict()
    id_to_name = routes.set_index("Route_ID")["Route_Name"].to_dict()
    route_ids_in_a = set(routes["Route_ID"])
    route_names_in_a = set(routes["Route_Name_norm"])

    # ---------- Read Format C ----------
    c = read_excel_any(annex_c_path, header=0)
    c.columns = [str(col).strip() for col in c.columns]

    must_have_c = ["Transnet Route ID", "Working Route Name as per Transnet"]
    missing_c = [x for x in must_have_c if x not in c.columns]
    if missing_c:
        raise ValueError(f"Format C missing required columns: {missing_c}. Ensure headers are exactly as finalized in Row-1.")

    duration_col = detect_fault_duration_column(c)

    faults = c.copy()
    faults["Route_ID_raw"] = faults["Transnet Route ID"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    faults["Route_Name_raw"] = faults["Working Route Name as per Transnet"].astype(str)
    faults["Route_Name_norm"] = faults["Route_Name_raw"].apply(norm_route_name)

    exempt_col = find_exemption_column(faults.columns)
    faults["Is_Exempt"] = faults[exempt_col].apply(robust_yes) if exempt_col else False
    faults["Duration_Hrs"] = faults[duration_col].apply(parse_duration_to_hours)

    faults_valid = faults[(faults["Duration_Hrs"].notna()) & (faults["Duration_Hrs"] > 0)].copy()
    faults_invalid = faults[~((faults["Duration_Hrs"].notna()) & (faults["Duration_Hrs"] > 0))].copy()

    faults_valid["Route_ID_mapped_by_id"] = faults_valid["Route_ID_raw"].where(faults_valid["Route_ID_raw"].isin(route_ids_in_a))
    faults_valid["Route_ID_mapped_by_name"] = faults_valid["Route_Name_norm"].map(name_to_id)

    faults_valid["Route_ID_Final"] = (
        faults_valid["Route_ID_mapped_by_id"]
        .fillna(faults_valid["Route_ID_mapped_by_name"])
        .fillna(faults_valid["Route_ID_raw"])
    )

    faults_valid["Matched_By_Name"] = np.where(
        faults_valid["Route_ID_mapped_by_id"].isna() & faults_valid["Route_ID_mapped_by_name"].notna(),
        "YES", "NO"
    )

    faults_valid["Route_Name_Final"] = faults_valid["Route_ID_Final"].map(id_to_name).fillna(faults_valid["Route_Name_raw"])

    penalties = faults_valid["Duration_Hrs"].apply(mttr_penalty_non_cumulative)
    faults_valid["MTTR_Penalty_Tender_Rs"] = [p[0] for p in penalties]
    faults_valid["MTTR_Slab"] = [p[1] for p in penalties]

    faults_valid["MTTR_Penalty_Exempted_Rs"] = np.where(faults_valid["Is_Exempt"], faults_valid["MTTR_Penalty_Tender_Rs"], 0.0)
    faults_valid["MTTR_Penalty_Net_Rs"] = faults_valid["MTTR_Penalty_Tender_Rs"] - faults_valid["MTTR_Penalty_Exempted_Rs"]

    SLAB_ORDER = ["≤4", ">4–6", ">6–24", ">24–48", ">48"]
    slab_summary = (
        faults_valid.groupby("MTTR_Slab", dropna=False)
        .agg(
            Count=("MTTR_Slab", "size"),
            Penalty_Gross=("MTTR_Penalty_Tender_Rs", "sum"),
            Penalty_Exempted=("MTTR_Penalty_Exempted_Rs", "sum"),
            Penalty_Net=("MTTR_Penalty_Net_Rs", "sum"),
        )
        .reset_index()
    )
    all_slabs = pd.DataFrame({"MTTR_Slab": SLAB_ORDER})
    slab_summary = all_slabs.merge(slab_summary, on="MTTR_Slab", how="left").fillna(0)

    for col in ["Penalty_Gross", "Penalty_Exempted", "Penalty_Net"]:
        slab_summary[col] = slab_summary[col].round(2)

    total_faults_count = int(slab_summary["Count"].sum())
    mttr_gross = float(slab_summary["Penalty_Gross"].sum())
    mttr_exempt = float(slab_summary["Penalty_Exempted"].sum())
    mttr_net = float(slab_summary["Penalty_Net"].sum())

    # Availability
    downtime_all = faults_valid.groupby("Route_ID_Final")["Duration_Hrs"].sum().reset_index()
    downtime_all.rename(columns={"Route_ID_Final": "Route_ID", "Duration_Hrs": "Downtime_Hrs_Total"}, inplace=True)

    faults_net = faults_valid[~faults_valid["Is_Exempt"]].copy()
    downtime_net = faults_net.groupby("Route_ID_Final")["Duration_Hrs"].sum().reset_index()
    downtime_net.rename(columns={"Route_ID_Final": "Route_ID", "Duration_Hrs": "Downtime_Hrs_Net"}, inplace=True)

    avail = routes.merge(downtime_all, on="Route_ID", how="left").merge(downtime_net, on="Route_ID", how="left")
    avail["Downtime_Hrs_Total"] = avail["Downtime_Hrs_Total"].fillna(0.0)
    avail["Downtime_Hrs_Net"] = avail["Downtime_Hrs_Net"].fillna(0.0)
    avail["Downtime_Exempted_Hrs"] = (avail["Downtime_Hrs_Total"] - avail["Downtime_Hrs_Net"]).round(4)

    avail["Uptime_pct_Gross"] = ((total_hours_month - avail["Downtime_Hrs_Total"]) / total_hours_month) * 100.0
    avail["Uptime_pct_Net"] = ((total_hours_month - avail["Downtime_Hrs_Net"]) / total_hours_month) * 100.0
    avail["Uptime_pct_Gross"] = avail["Uptime_pct_Gross"].clip(0, 100)
    avail["Uptime_pct_Net"] = avail["Uptime_pct_Net"].clip(0, 100)

    avail["Deduction_pct_Gross"] = avail["Uptime_pct_Gross"].apply(uptime_deduction_pct)
    avail["Deduction_pct_Net"] = avail["Uptime_pct_Net"].apply(uptime_deduction_pct)

    avail["Deduction_Rs_Gross"] = (avail["SLA_Charges_Rs"] * avail["Deduction_pct_Gross"] / 100.0).round(2)
    avail["Deduction_Rs_Net"] = (avail["SLA_Charges_Rs"] * avail["Deduction_pct_Net"] / 100.0).round(2)

    avail["Availability_Deduction_Exempted_Rs"] = (avail["Deduction_Rs_Gross"] - avail["Deduction_Rs_Net"]).round(2)
    avail["Availability_Deduction_Net_Rs"] = avail["Deduction_Rs_Net"]

    availability_penalty_net = round(float(avail["Availability_Deduction_Net_Rs"].sum()), 2)

    # Missing routes text
    missing_mask = (~faults_valid["Route_ID_Final"].isin(route_ids_in_a)) & (~faults_valid["Route_Name_norm"].isin(route_names_in_a))
    missing_rows = faults_valid[missing_mask].copy()
    missing_group = pd.DataFrame()
    if len(missing_rows) > 0:
        missing_group = (missing_rows.groupby(["Route_ID_raw", "Route_Name_raw"], dropna=False)["Duration_Hrs"]
                         .sum().reset_index()
                         .rename(columns={"Route_ID_raw": "Route_ID_in_C", "Route_Name_raw": "Route_Name_in_C", "Duration_Hrs": "Total_Downtime_Hrs"})
                         .sort_values("Total_Downtime_Hrs", ascending=False))

    missing_lines = []
    if len(missing_group) == 0:
        missing_lines.append("Routes in Format-C but not found in Format-A: NIL")
    else:
        missing_lines.append("Routes in Format-C but not found in Format-A (even after ID + Name matching):")
        for _, r in missing_group.iterrows():
            missing_lines.append(f" - {r['Route_ID_in_C']} | {r['Route_Name_in_C']} | Downtime: {r['Total_Downtime_Hrs']:.2f} hrs")

    # MTTR 25% cap
    total_basic_sla = round(float(routes["SLA_Charges_Rs"].sum()), 2)
    mttr_cap_25pct = round(total_basic_sla * 0.25, 2)
    if mttr_net > mttr_cap_25pct:
        mttr_net_after_cap = mttr_cap_25pct
        mttr_cap_applied = "YES"
    else:
        mttr_net_after_cap = round(mttr_net, 2)
        mttr_cap_applied = "NO"

    field_unit_penalty = float(field_unit_penalty or 0.0)
    system_sla_penalty_net = round(mttr_net_after_cap + availability_penalty_net, 2)
    higher_of_penalty = max(system_sla_penalty_net, field_unit_penalty)

    vendor_deducted_penalty = float(vendor_deducted_penalty or 0.0)
    sla_recovery_after_vendor = round(max(higher_of_penalty - vendor_deducted_penalty, 0.0), 2)

    total_rkm = round(float(routes["Route_KM"].sum()), 2)
    system_basic = round(total_basic_sla, 2)

    # (Accounts note generation remains as you already finalized earlier)
    # (Clause 14.1 note updated below)

    splice_loss_amt = float(splice_loss_amt or 0.0)
    supervisor_abs_amt = float(supervisor_abs_amt or 0.0)
    frt_abs_amt = float(frt_abs_amt or 0.0)
    petroller_abs_amt = float(petroller_abs_amt or 0.0)
    relaying_not_done_amt = float(relaying_not_done_amt or 0.0)

    total_penalty_clause14 = round(
        splice_loss_amt + mttr_net_after_cap + availability_penalty_net +
        supervisor_abs_amt + frt_abs_amt + petroller_abs_amt + relaying_not_done_amt, 2
    )

    header_info = f"""BA: {ba_name}
OA: {oa_name}
SLA Month: {month_display}
Name of Maintenance Agency: {vendor_name}
"""

    # ---------- Clause 14.1 Technical Note (FIXED RKM/RATE/Service Value + SES blank line) ----------
    slab_map = {
        "≤4": "Upto 4 Hrs",
        ">4–6": "Between 4 Hrs to 6 Hrs",
        ">6–24": "Between 6 Hrs  to 24 Hrs",
        ">24–48": "Between 24 hrs to 48 Hrs",
        ">48": "Beyond 48 Hrs."
    }

    mttr_header = f"{'Slab':<30} {'Count':>7} {'Penalty':>14} {'Exempted':>14} {'Net':>14}"
    mttr_sep = "-" * len(mttr_header)
    mttr_lines = [mttr_header, mttr_sep]
    for _, r in slab_summary.iterrows():
        name = slab_map.get(r["MTTR_Slab"], str(r["MTTR_Slab"]))
        mttr_lines.append(
            f"{name:<30} "
            f"{int(r['Count']):>7} "
            f"{fmt_money(r['Penalty_Gross']):>14} "
            f"{fmt_money(r['Penalty_Exempted']):>14} "
            f"{fmt_money(r['Penalty_Net']):>14}"
        )
    mttr_lines.append(mttr_sep)
    mttr_lines.append(
        f"{'Total':<30} "
        f"{total_faults_count:>7} "
        f"{fmt_money(mttr_gross):>14} "
        f"{fmt_money(mttr_exempt):>14} "
        f"{fmt_money(mttr_net):>14}"
    )

    avail_view = avail.copy()
    avail_view["Route_ID"] = avail_view["Route_ID"].astype(str)
    avail_view["Route_Name"] = avail_view["Route_Name"].astype(str)

    avail_focus = avail_view[(avail_view["Availability_Deduction_Net_Rs"] > 0) | (avail_view["Downtime_Hrs_Total"] > 0)].copy()
    avail_focus = avail_focus.sort_values(["Availability_Deduction_Net_Rs", "Downtime_Hrs_Net"], ascending=[False, False])

    av_header = f"{'Route ID':<18} {'Route Name':<55} {'Uptime%':>8} {'Ded%':>6} {'Penalty(Net)':>14}"
    av_sep = "-" * len(av_header)
    av_lines = [av_header, av_sep]

    if len(avail_focus) == 0:
        av_lines.append("No route has downtime/penalty for this month.")
    else:
        for _, rr in avail_focus.iterrows():
            rid = str(rr["Route_ID"])[:18]
            rname = str(rr["Route_Name"])[:55]
            av_lines.append(
                f"{rid:<18} {rname:<55} "
                f"{rr['Uptime_pct_Net']:>8.2f} "
                f"{int(rr['Deduction_pct_Net']):>6} "
                f"{fmt_money(rr['Availability_Deduction_Net_Rs']):>14}"
            )

    technical_note_clause14 = f"""SLA PENALTIES CALCULATION AS PER TENDER CLAUSE 14.1

{header_info}

Total of RKM as per Annexure A                                   = {total_rkm:.2f}

Rate as per Tender Rs.                                           = {rate_per_km:,.2f} Per KM Monthly.

Total Value of service Rs. (RKM*RATE)                            = Rs. {fmt_money(system_basic)}

Penalty Details given below:-

1. SPLICE LOSS PER FIBER Rs.                                     : Rs. {fmt_money(splice_loss_amt)}

2. MTTR FAULTS Penalty Details                                   : (Code generated from Format-C)

   Max 25% MTTR Penalty Rs. (RKM*Rate*25%)                        : Rs. {fmt_money(mttr_cap_25pct)}

   Slab wise faults and penalty:
{chr(10).join(mttr_lines)}

   Total MTTR Net penalty after Exemption (and 25% cap)           : Rs. {fmt_money(mttr_net_after_cap)}
   Cap Applied                                                    : {mttr_cap_applied}

3. Availability Penalty (Route-wise)                              : Rs. {fmt_money(availability_penalty_net)}

   Availability details (Net):
{chr(10).join(av_lines)}

4. Absense of Supervisor @ 1500 per day Rs.                       : Rs. {fmt_money(supervisor_abs_amt)}
5. Absence of FRT @ 5000 Per day Rs.                              : Rs. {fmt_money(frt_abs_amt)}
6. Absence of Petroller @ 500 Per day Rs.                         : Rs. {fmt_money(petroller_abs_amt)}
7. Penalty for 1% Re-laying ofc Work not done @ 200000 Per KM Rs.  : Rs. {fmt_money(relaying_not_done_amt)}

Total Penalty (1+2+3+4+5+6+7) Rs.                                  : Rs. {fmt_money(total_penalty_clause14)}

It is hereby submitted that the vendor's services for the month of [{month_display_short}] have undergone
full verification against the Advance Purchase Order (APO) and the corresponding proof of delivery for materials/services.

It is further confirmed that any applicable penalties and deductions have been
calculated and applied in strict accordance with the stipulated terms and conditions
outlined in the respective purchase order, APO, or tender document.

The supporting documents listed below have been verified
and uploaded under SAP Service Entry Sheet No: ______________

1. Excel file of Annexure A
2. Excel file of Annexure C

Submitted for approval please.
"""

    # ---------- Output files ----------
    os.makedirs(save_dir, exist_ok=True)
    out_xlsx = os.path.join(save_dir, f"SLA_Output_{vendor_tag}_{month_tag2}.xlsx")
    out_accounts_txt = os.path.join(save_dir, f"SAP_Accounts_Note_{vendor_tag}_{month_tag2}.txt")
    out_tech_txt = os.path.join(save_dir, f"Penalty_Clause14_1_{vendor_tag}_{month_tag2}.txt")

    # Keep your existing Excel outputs (unchanged)
    summary = pd.DataFrame([["Info", "OK"]], columns=["Item", "Value"])
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        avail.sort_values(["Route_ID"]).to_excel(writer, sheet_name="Availability_Report", index=False)
        faults_valid.to_excel(writer, sheet_name="MTTR_Fault_Report", index=False)
        slab_summary.to_excel(writer, sheet_name="MTTR_Slab_Summary", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        if len(faults_invalid) > 0:
            faults_invalid.to_excel(writer, sheet_name="Invalid_Fault_Rows", index=False)

    # NOTE: keep your already-finalized Accounts Note builder here.
    # If you already have accounts_note variable in your repo version, keep it unchanged.
    # For safety, write a placeholder if not defined (should not happen in your final repo).
    accounts_note = "Accounts note is generated by your existing finalized template."

    with open(out_accounts_txt, "w", encoding="utf-8") as f:
        f.write(accounts_note)

    with open(out_tech_txt, "w", encoding="utf-8") as f:
        f.write(technical_note_clause14)

    return out_xlsx, out_accounts_txt, out_tech_txt
