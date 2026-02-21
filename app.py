import os
import io
import zipfile
import tempfile
from datetime import datetime

import streamlit as st
from sla_logic import process_sla

st.set_page_config(page_title="SLA Bill Checker", layout="wide", page_icon="📘")

# ---------- CSS ----------
st.markdown("""
<style>
.main {background-color:#f6f8fb;}
.big-title {font-size:40px;font-weight:900;color:#0B3D91;margin-bottom:4px;}
.creator {font-size:16px;font-weight:700;color:#2b2b2b;margin-bottom:10px;}
.section {font-size:20px;font-weight:800;color:#0B3D91;margin-top:18px;margin-bottom:8px;}
small.help {color:#666;}
.stButton > button {height:52px;font-size:18px;font-weight:800;border-radius:10px;background:#0B3D91;color:white;}
.stDownloadButton > button {height:52px;font-size:18px;font-weight:800;border-radius:10px;}
label {font-weight:800 !important; font-size:16px !important;}
</style>
""", unsafe_allow_html=True)

# ---------- HEADER ----------
st.markdown('<div class="big-title">SLA Bill Checker</div>', unsafe_allow_html=True)
st.markdown('<div class="creator">Created by: Hrushikesh Kesale | MH Circle </div>', unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ---------- Session reset helper ----------
def reset_form():
    keys = [
        "annex_a","annex_c","rate","vendor_basic","pan4","field_pen","vendor_ded","other_rec",
        "splice","sup_abs","frt_abs","pet_abs","relay"
    ]
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

# ---------- FILE UPLOAD ----------
st.markdown('<div class="section">📂 Upload Files</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    annex_a = st.file_uploader("Upload Format A (Annexure A) *", type=["xlsx","xls"], key="annex_a")
with c2:
    annex_c = st.file_uploader("Upload Format C (Annexure C) *", type=["xlsx","xls"], key="annex_c")

# ---------- INPUTS ----------
st.markdown('<div class="section">⚙️ Inputs</div>', unsafe_allow_html=True)

i1, i2, i3 = st.columns(3)
with i1:
    rate = st.number_input("Rate per KM (Required) *", min_value=0.0, step=1.0, format="%.2f", key="rate")
with i2:
    vendor_basic = st.text_input("Vendor Basic Value before GST (Optional)", placeholder="Leave blank = Σ(RKM×Rate)", key="vendor_basic")
with i3:
    pan4 = st.text_input("PAN 4th Digit (Optional)", placeholder="P/H=1% else 2%", max_chars=1, key="pan4")

j1, j2, j3 = st.columns(3)
with j1:
    field_pen = st.number_input("Field Unit / SES Penalty (Info)", value=0.0, step=1.0, format="%.2f", key="field_pen")
with j2:
    vendor_ded = st.number_input("Vendor already deducted SLA penalty", value=0.0, step=1.0, format="%.2f", key="vendor_ded")
with j3:
    other_rec = st.number_input("Any other recovery (Accounts)", value=0.0, step=1.0, format="%.2f", key="other_rec")

# ---------- CLAUSE 14.1 MANUAL ----------
st.markdown('<div class="section">📑 Clause 14.1 Manual Inputs</div>', unsafe_allow_html=True)

m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    splice = st.number_input("Splice Loss ₹", value=0.0, step=1.0, format="%.2f", key="splice")
with m2:
    sup_abs = st.number_input("Supervisor Abs ₹", value=0.0, step=1.0, format="%.2f", key="sup_abs")
with m3:
    frt_abs = st.number_input("FRT Abs ₹", value=0.0, step=1.0, format="%.2f", key="frt_abs")
with m4:
    pet_abs = st.number_input("Petroller Abs ₹", value=0.0, step=1.0, format="%.2f", key="pet_abs")
with m5:
    relay = st.number_input("1% Relaying Not Done ₹", value=0.0, step=1.0, format="%.2f", key="relay")

st.markdown("<br>", unsafe_allow_html=True)

# ---------- Helpers ----------
def to_float_or_none(s):
    s = (s or "").strip()
    if s == "":
        return None
    try:
        return float(s)
    except:
        return None

def pan_or_none(s):
    s = (s or "").strip().upper()
    return None if s == "" else s[0]

# ---------- ACTION BUTTONS ----------
a1, a2 = st.columns([2, 1])
with a1:
    generate = st.button("🚀 Generate Output", use_container_width=True)
with a2:
    st.button("🧹 Clear Page", use_container_width=True, on_click=reset_form)

# ---------- PROCESS ----------
if generate:
    if annex_a is None or annex_c is None:
        st.error("Please upload BOTH Annexure A and Annexure C.")
        st.stop()
    if rate <= 0:
        st.error("Rate per KM must be greater than 0.")
        st.stop()

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            a_path = os.path.join(tmpdir, f"AnnexureA_{annex_a.name}")
            c_path = os.path.join(tmpdir, f"AnnexureC_{annex_c.name}")

            with open(a_path, "wb") as f:
                f.write(annex_a.read())
            with open(c_path, "wb") as f:
                f.write(annex_c.read())

            out_xlsx, out_acc, out_tech = process_sla(
                annex_a_path=a_path,
                annex_c_path=c_path,
                rate_per_km=rate,
                save_dir=tmpdir,
                vendor_basic_value=to_float_or_none(vendor_basic),
                pan4=pan_or_none(pan4),
                field_unit_penalty=field_pen,
                vendor_deducted_penalty=vendor_ded,
                other_recovery=other_rec,
                splice_loss_amt=splice,
                supervisor_abs_amt=sup_abs,
                frt_abs_amt=frt_abs,
                petroller_abs_amt=pet_abs,
                relaying_not_done_amt=relay,
            )

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as z:
                z.write(out_xlsx, arcname=os.path.basename(out_xlsx))
                z.write(out_acc, arcname=os.path.basename(out_acc))
                z.write(out_tech, arcname=os.path.basename(out_tech))

            zip_buffer.seek(0)

            st.success("✅ Files generated successfully. Download ZIP below.")

            st.download_button(
                "⬇️ Download ZIP (Excel + 2 Notes)",
                data=zip_buffer,
                file_name=f"SLA_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )

            st.info("After downloading, click **Clear Page** to start fresh.")

    except Exception as e:
        st.exception(e)