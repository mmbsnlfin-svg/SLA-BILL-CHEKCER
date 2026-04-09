import os
import io
import zipfile
import tempfile
import calendar
from datetime import datetime

import streamlit as st
import pandas as pd

from sla_logic import process_sla

st.set_page_config(
    page_title="BSNL SLA Bill Checker",
    layout="wide",
    page_icon="📊"
)

# -----------------------------
# CSS - Professional Look
# -----------------------------
st.markdown("""
<style>
.main > div {
    padding-top: 1rem;
}
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
}
.bsnl-card {
    background: linear-gradient(135deg, #ffffff 0%, #f7f9fc 100%);
    border: 1px solid #d9e2f0;
    border-radius: 18px;
    padding: 18px 22px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.06);
    margin-bottom: 14px;
}
.bsnl-title {
    font-size: 30px;
    font-weight: 800;
    color: #0b2d6b;
    margin-bottom: 2px;
}
.bsnl-subtitle {
    font-size: 14px;
    font-weight: 700;
    color: #203864;
}
.bsnl-caption {
    font-size: 12px;
    color: #5b6575;
    margin-top: 6px;
}
.small-note {
    font-size: 12px;
    color: #666;
}
.metric-box {
    background: #f8fbff;
    border: 1px solid #dbe8f6;
    border-radius: 14px;
    padding: 12px 16px;
}
div[data-testid="stForm"] {
    border: 1px solid #e3e8f2;
    border-radius: 18px;
    padding: 18px 16px 8px 16px;
    background: #ffffff;
    box-shadow: 0 6px 20px rgba(0,0,0,0.05);
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Expected headers (EXACT)
# -----------------------------
REQUIRED_A = [
    "FORMAT", "BA", "OA", "Month", "Sr.No.",
    "Transnet Route ID", "Working Route Name as per Transnet",
    "RKM", "Name of Maintenance Agency"
]
REQUIRED_C = [
    "Transnet Route ID", "Working Route Name as per Transnet"
]

# -----------------------------
# Helpers
# -----------------------------
def clear_form():
    keys = [
        "annex_a", "annex_c",
        "vendor_basic", "field_unit_penalty",
        "vendor_deducted_penalty", "other_recovery",
        "splice_loss", "supervisor_abs", "frt_abs", "petroller_abs", "relaying_penalty",
        "relaying_as_retention", "selected_month", "selected_ba", "selected_oa", "selected_vendor"
    ]
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]


def normalize_cols(cols):
    out = []
    for c in cols:
        s = str(c).strip()
        s = " ".join(s.split())
        out.append(s)
    return out


def classify_file(columns_list):
    cols = set(columns_list)
    has_a = all(c in cols for c in REQUIRED_A)
    has_c = all(c in cols for c in REQUIRED_C)

    if has_a and not has_c:
        return "A"
    if has_c and not has_a:
        return "C"
    if has_a and has_c:
        return "A"
    return "UNKNOWN"


def missing_columns(columns_list, required_list):
    cols = set(columns_list)
    return [c for c in required_list if c not in cols]


def fnum(x, default=0.0):
    try:
        s = str(x).strip()
        if s == "":
            return default
        return float(s)
    except Exception:
        return default


def make_month_options(back_months=24, ahead_months=3):
    today = datetime.today()
    options = []
    base_year = today.year
    base_month = today.month

    total_offsets = list(range(-back_months, ahead_months + 1))
    for offset in total_offsets:
        month_num = base_month + offset
        year_num = base_year

        while month_num <= 0:
            month_num += 12
            year_num -= 1
        while month_num > 12:
            month_num -= 12
            year_num += 1

        options.append(datetime(year_num, month_num, 1).strftime("%b-%Y"))

    # unique while preserving order
    seen = set()
    final = []
    for x in options:
        if x not in seen:
            final.append(x)
            seen.add(x)
    return final


def first_existing(df, candidates):
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def load_vendorinfo(path="vendorinfo.xlsx"):
    if not os.path.exists(path):
        raise FileNotFoundError(
            "vendorinfo.xlsx not found in repository. Please keep vendorinfo.xlsx in the same folder as app.py"
        )

    df = pd.read_excel(path)
    df.columns = normalize_cols(df.columns)

    ba_col = first_existing(df, ["BA", "BA Name", "Business Area"])
    oa_col = first_existing(df, ["OA", "OA Name", "Operational Area", "SSA", "SSA Name"])
    vendor_col = first_existing(df, ["Vendor Name", "Name of Maintenance Agency", "Vendor"])
    pan4_col = first_existing(df, ["PAN 4th Digit", "Pan 4th Digit", "PAN4", "PAN 4th"])
    it_rate_col = first_existing(df, ["Income Tax rate slab", "IT Rate", "IT Rate Slab", "Income Tax Rate"])
    rate_col = first_existing(df, ["Rate per Rkm", "Rate per RKM", "Rate per KM", "Rate"])

    required_map = {
        "BA": ba_col,
        "OA": oa_col,
        "Vendor Name": vendor_col,
        "Rate per RKM": rate_col
    }
    missing = [k for k, v in required_map.items() if v is None]
    if missing:
        raise ValueError(
            f"vendorinfo.xlsx missing required columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    out = df.copy()
    out = out.rename(columns={
        ba_col: "BA",
        oa_col: "OA",
        vendor_col: "Vendor_Name",
        rate_col: "Rate_per_RKM"
    })

    if pan4_col:
        out = out.rename(columns={pan4_col: "PAN4"})
    else:
        out["PAN4"] = ""

    if it_rate_col:
        out = out.rename(columns={it_rate_col: "IT_Rate"})
    else:
        out["IT_Rate"] = ""

    out["BA"] = out["BA"].astype(str).str.strip()
    out["OA"] = out["OA"].astype(str).str.strip()
    out["Vendor_Name"] = out["Vendor_Name"].astype(str).str.strip()
    out["PAN4"] = out["PAN4"].fillna("").astype(str).str.strip().str.upper()
    out["Rate_per_RKM"] = pd.to_numeric(out["Rate_per_RKM"], errors="coerce")

    out = out.dropna(subset=["Rate_per_RKM"])
    out = out[(out["BA"] != "") & (out["OA"] != "") & (out["Vendor_Name"] != "")]

    return out


# -----------------------------
# Header
# -----------------------------
st.markdown("""
<div class="bsnl-card">
    <div class="bsnl-title">BSNL SLA Bill Checker</div>
    <div class="bsnl-subtitle">Created by: Hrushikesh Kesale | MH Circle BSNL</div>
    <div class="bsnl-caption">
        Select Month, BA, OA and Vendor from master file → Upload Annexure A & Annexure C → Generate Excel + Accounts Note + Clause 14.1 Penalty Note
    </div>
</div>
""", unsafe_allow_html=True)

# -----------------------------
# Load vendor master
# -----------------------------
try:
    vendor_df = load_vendorinfo("vendorinfo.xlsx")
except Exception as e:
    st.error(f"Unable to load vendor master file: {e}")
    st.stop()

# -----------------------------
# Selection area
# -----------------------------
months = make_month_options()
default_month = datetime.today().strftime("%b-%Y")
default_month_index = months.index(default_month) if default_month in months else 0

st.markdown("### Master Selection")
sel1, sel2, sel3, sel4 = st.columns(4)

with sel1:
    selected_month = st.selectbox(
        "Month (MMM-YYYY)",
        months,
        index=default_month_index,
        key="selected_month"
    )

ba_options = sorted(vendor_df["BA"].dropna().unique().tolist())
with sel2:
    selected_ba = st.selectbox("BA", ba_options, key="selected_ba")

oa_filtered_df = vendor_df[vendor_df["BA"] == selected_ba].copy()
oa_options = sorted(oa_filtered_df["OA"].dropna().unique().tolist())
with sel3:
    selected_oa = st.selectbox("OA", oa_options, key="selected_oa")

vendor_filtered_df = oa_filtered_df[oa_filtered_df["OA"] == selected_oa].copy()
vendor_options = sorted(vendor_filtered_df["Vendor_Name"].dropna().unique().tolist())
with sel4:
    selected_vendor = st.selectbox("Vendor Name", vendor_options, key="selected_vendor")

selected_row_df = vendor_filtered_df[vendor_filtered_df["Vendor_Name"] == selected_vendor].copy()

if selected_row_df.empty:
    st.error("Selected BA / OA / Vendor combination not found in vendorinfo.xlsx")
    st.stop()

selected_row = selected_row_df.iloc[0]
locked_rate = float(selected_row["Rate_per_RKM"])
locked_pan4 = str(selected_row.get("PAN4", "")).strip().upper()
locked_it_rate = str(selected_row.get("IT_Rate", "")).strip()

m1, m2, m3 = st.columns(3)
with m1:
    st.markdown(f"""
    <div class="metric-box">
        <div class="small-note">Locked Rate per RKM</div>
        <div style="font-size:24px;font-weight:800;color:#0b2d6b;">₹ {locked_rate:,.2f}</div>
    </div>
    """, unsafe_allow_html=True)

with m2:
    st.markdown(f"""
    <div class="metric-box">
        <div class="small-note">PAN 4th Digit</div>
        <div style="font-size:24px;font-weight:800;color:#0b2d6b;">{locked_pan4 if locked_pan4 else '-'}</div>
    </div>
    """, unsafe_allow_html=True)

with m3:
    st.markdown(f"""
    <div class="metric-box">
        <div class="small-note">Income Tax Slab</div>
        <div style="font-size:24px;font-weight:800;color:#0b2d6b;">{locked_it_rate if locked_it_rate else '-'}</div>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# -----------------------------
# Main Form
# -----------------------------
with st.form("sla_form"):
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown("### Upload Files")
        annex_a = st.file_uploader(
            "Format A (Annexure A) Excel",
            type=["xlsx", "xls"],
            key="annex_a"
        )
        annex_c = st.file_uploader(
            "Format C (Annexure C) Excel",
            type=["xlsx", "xls"],
            key="annex_c"
        )

        st.info(
            f"Selected: Month = {selected_month} | BA = {selected_ba} | OA = {selected_oa} | Vendor = {selected_vendor}"
        )

    with col2:
        st.markdown("### Auto / Optional Inputs")
        st.text_input("Rate per KM (Auto Locked)", value=f"{locked_rate:.2f}", disabled=True)
        st.text_input("PAN 4th Digit (Auto from Vendor Master)", value=locked_pan4, disabled=True)
        st.text_input("Income Tax Slab (Info from Vendor Master)", value=locked_it_rate, disabled=True)

        vendor_basic = st.text_input("Vendor Basic Value before GST (Optional)", value="", key="vendor_basic")
        field_unit_penalty = st.text_input("Field Unit / SES Penalty (Info)", value="0", key="field_unit_penalty")
        vendor_deducted_penalty = st.text_input("Vendor already deducted SLA penalty", value="0", key="vendor_deducted_penalty")
        other_recovery = st.text_input("Any other recovery (Accounts)", value="0", key="other_recovery")

    st.markdown("### Clause 14.1 Manual Inputs")
    c3, c4, c5 = st.columns(3, gap="large")
    with c3:
        splice_loss = st.text_input("1) Splice Loss per Fiber ₹", value="0", key="splice_loss")
        supervisor_abs = st.text_input("4) Absence of Supervisor ₹", value="0", key="supervisor_abs")
    with c4:
        frt_abs = st.text_input("5) Absence of FRT ₹", value="0", key="frt_abs")
        petroller_abs = st.text_input("6) Absence of Petroller ₹", value="0", key="petroller_abs")
    with c5:
        relaying_penalty = st.text_input("7) 1% Re-laying work not done ₹", value="0", key="relaying_penalty")
        relaying_as_retention = st.checkbox(
            "Treat 1% Re-laying amount as Retention (not Penalty)",
            value=False,
            key="relaying_as_retention"
        )

    b1, b2 = st.columns([1, 1])
    with b1:
        submitted = st.form_submit_button("✅ Generate Output")
    with b2:
        st.form_submit_button("🧹 Clear Form", on_click=clear_form)

# -----------------------------
# Processing
# -----------------------------
if submitted:
    if annex_a is None or annex_c is None:
        st.error("Please upload both Annexure A and Annexure C files.")
        st.stop()

    rate = locked_rate
    if rate <= 0:
        st.error("Rate per KM from vendor master is invalid.")
        st.stop()

    vendor_basic_val = fnum(vendor_basic, default=float("nan"))
    vendor_basic_val = None if pd.isna(vendor_basic_val) else vendor_basic_val

    pan4_val = locked_pan4 if locked_pan4 != "" else None

    field_pen = fnum(field_unit_penalty, 0.0)
    vendor_ded = fnum(vendor_deducted_penalty, 0.0)
    other_rec = fnum(other_recovery, 0.0)

    splice = fnum(splice_loss, 0.0)
    sup_abs = fnum(supervisor_abs, 0.0)
    frt = fnum(frt_abs, 0.0)
    pet = fnum(petroller_abs, 0.0)
    relay = fnum(relaying_penalty, 0.0)

    with st.spinner("Processing..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            a_path = os.path.join(tmpdir, "Annexure_A.xlsx")
            c_path = os.path.join(tmpdir, "Annexure_C.xlsx")

            with open(a_path, "wb") as f:
                f.write(annex_a.getbuffer())
            with open(c_path, "wb") as f:
                f.write(annex_c.getbuffer())

            # -----------------------------
            # Validate wrong upload / header deviation
            # -----------------------------
            try:
                a_prev = pd.read_excel(a_path, nrows=1)
                c_prev = pd.read_excel(c_path, nrows=1)

                a_cols = normalize_cols(a_prev.columns)
                c_cols = normalize_cols(c_prev.columns)

                a_type = classify_file(a_cols)
                c_type = classify_file(c_cols)

                # swapped detection
                if a_type == "C" and c_type == "A":
                    st.error(
                        "Wrong files uploaded ❌\n\n"
                        "It looks like you uploaded Annexure C file in Annexure A upload\n"
                        "and Annexure A file in Annexure C upload.\n\n"
                        "Please swap the files and upload correctly."
                    )
                    st.info("Annexure A upload headers detected:")
                    st.write(a_cols)
                    st.info("Annexure C upload headers detected:")
                    st.write(c_cols)
                    st.stop()

                miss_a = missing_columns(a_cols, REQUIRED_A)
                if miss_a:
                    st.error(
                        "Annexure A (Format-A) column mismatch ❌\n\n"
                        f"Missing required columns: {miss_a}\n\n"
                        "Please correct the column names exactly as per standard format."
                    )
                    st.info("Expected Annexure A headers (exact):")
                    st.write(REQUIRED_A)
                    st.info("Your uploaded Annexure A headers (detected):")
                    st.write(a_cols)
                    st.stop()

                miss_c = missing_columns(c_cols, REQUIRED_C)
                if miss_c:
                    st.error(
                        "Annexure C (Format-C) column mismatch ❌\n\n"
                        f"Missing required columns: {miss_c}\n\n"
                        "Please correct the column names exactly as per standard format."
                    )
                    st.info("Expected Annexure C headers (must include these exact):")
                    st.write(REQUIRED_C)
                    st.info("Your uploaded Annexure C headers (detected):")
                    st.write(c_cols)
                    st.stop()

            except Exception as e:
                st.error(f"Unable to validate Excel headers. Please check the uploaded files. Error: {e}")
                st.stop()

            # -----------------------------
            # Optional month/vendor consistency checks
            # -----------------------------
            try:
                a_full = pd.read_excel(a_path)
                a_full.columns = normalize_cols(a_full.columns)

                uploaded_ba = str(a_full["BA"].dropna().astype(str).iloc[0]).strip() if "BA" in a_full.columns and len(a_full.dropna(how="all")) else ""
                uploaded_oa = str(a_full["OA"].dropna().astype(str).iloc[0]).strip() if "OA" in a_full.columns and len(a_full.dropna(how="all")) else ""
                uploaded_vendor = str(a_full["Name of Maintenance Agency"].dropna().astype(str).iloc[0]).strip() if "Name of Maintenance Agency" in a_full.columns and len(a_full.dropna(how="all")) else ""
                uploaded_month = str(a_full["Month"].dropna().astype(str).iloc[0]).strip() if "Month" in a_full.columns and len(a_full.dropna(how="all")) else ""

                if uploaded_ba and uploaded_ba.lower() != selected_ba.lower():
                    st.warning(f"Selected BA is '{selected_ba}', but Format A BA is '{uploaded_ba}'. Please verify.")
                if uploaded_oa and uploaded_oa.lower() != selected_oa.lower():
                    st.warning(f"Selected OA is '{selected_oa}', but Format A OA is '{uploaded_oa}'. Please verify.")
                if uploaded_vendor and uploaded_vendor.lower() != selected_vendor.lower():
                    st.warning(f"Selected Vendor is '{selected_vendor}', but Format A Vendor is '{uploaded_vendor}'. Please verify.")
                if uploaded_month and selected_month.lower() not in uploaded_month.lower():
                    st.info(f"Selected Month = {selected_month}; Format A Month detected = {uploaded_month}")
            except Exception:
                pass

            # -----------------------------
            # Run existing logic unchanged
            # -----------------------------
            out_xlsx, out_acc, out_tech = process_sla(
                annex_a_path=a_path,
                annex_c_path=c_path,
                rate_per_km=rate,
                save_dir=tmpdir,
                vendor_basic_value=vendor_basic_val,
                pan4=pan4_val,
                field_unit_penalty=field_pen,
                vendor_deducted_penalty=vendor_ded,
                other_recovery=other_rec,
                splice_loss_amt=splice,
                supervisor_abs_amt=sup_abs,
                frt_abs_amt=frt,
                petroller_abs_amt=pet,
                relaying_not_done_amt=relay,
                relaying_as_retention=bool(relaying_as_retention),
            )

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(out_xlsx, arcname=os.path.basename(out_xlsx))
                zf.write(out_acc, arcname=os.path.basename(out_acc))
                zf.write(out_tech, arcname=os.path.basename(out_tech))
            zip_buffer.seek(0)

            st.success("Done ✅ Output generated successfully.")
            st.download_button(
                "⬇️ Download Output (ZIP)",
                data=zip_buffer,
                file_name=f"SLA_Output_{selected_ba}_{selected_oa}_{selected_vendor}_{selected_month}.zip".replace(" ", "_"),
                mime="application/zip"
            )
