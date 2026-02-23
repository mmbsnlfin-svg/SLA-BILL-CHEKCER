import os
import io
import zipfile
import tempfile

import streamlit as st
import pandas as pd

from sla_logic import process_sla

st.set_page_config(page_title="SLA Bill Checker", layout="wide")


def clear_form():
    keys = [
        "annex_a", "annex_c",
        "rate_per_km", "vendor_basic", "pan4", "field_unit_penalty",
        "vendor_deducted_penalty", "other_recovery",
        "splice_loss", "supervisor_abs", "frt_abs", "petroller_abs", "relaying_penalty",
        "relaying_as_retention"
    ]
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]


st.markdown(
    """
    <div style="padding:8px 0;">
      <div style="font-size:28px; font-weight:800; color:#0b2d6b;">SLA Bill Checker</div>
      <div style="font-size:14px; font-weight:700; margin-top:2px;">
        Created by: Hrushikesh Kesale | MH Circle 
      </div>
      <div style="font-size:12px; color:#666; margin-top:6px;">
        Upload Annexure A & Annexure C → Generate Excel + Accounts Note + Clause 14.1 Penalty Note
      </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.divider()

with st.form("sla_form"):
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown("### **Upload Files**")
        annex_a = st.file_uploader("**Format A (Annexure A) Excel**", type=["xlsx", "xls"], key="annex_a")
        annex_c = st.file_uploader("**Format C (Annexure C) Excel**", type=["xlsx", "xls"], key="annex_c")

    with col2:
        st.markdown("### **Inputs**")
        rate_per_km = st.text_input("**Rate per KM (Required)**", value="", key="rate_per_km")
        vendor_basic = st.text_input("**Vendor Basic Value before GST (Optional)**", value="", key="vendor_basic")
        pan4 = st.text_input("**PAN 4th Digit (Optional)**", value="", key="pan4")
        field_unit_penalty = st.text_input("**Field Unit / SES Penalty (Info)**", value="0", key="field_unit_penalty")
        vendor_deducted_penalty = st.text_input("**Vendor already deducted SLA penalty**", value="0", key="vendor_deducted_penalty")
        other_recovery = st.text_input("**Any other recovery (Accounts)**", value="0", key="other_recovery")

    st.markdown("### **Clause 14.1 Manual Inputs**")
    c3, c4, c5 = st.columns(3, gap="large")
    with c3:
        splice_loss = st.text_input("**1) Splice Loss per Fiber ₹**", value="0", key="splice_loss")
        supervisor_abs = st.text_input("**4) Absence of Supervisor ₹**", value="0", key="supervisor_abs")
    with c4:
        frt_abs = st.text_input("**5) Absence of FRT ₹**", value="0", key="frt_abs")
        petroller_abs = st.text_input("**6) Absence of Petroller ₹**", value="0", key="petroller_abs")
    with c5:
        relaying_penalty = st.text_input("**7) 1% Re-laying work not done ₹**", value="0", key="relaying_penalty")
        relaying_as_retention = st.checkbox(
            "**Treat 1% Re-laying amount as Retention (not Penalty)**",
            value=False,
            key="relaying_as_retention"
        )

    b1, b2 = st.columns([1, 1])
    with b1:
        submitted = st.form_submit_button("✅ Generate Output")
    with b2:
        st.form_submit_button("🧹 Clear Form", on_click=clear_form)


if submitted:
    if annex_a is None or annex_c is None:
        st.error("Please upload both Annexure A and Annexure C files.")
        st.stop()

    # Validate rate
    try:
        rate = float(str(rate_per_km).strip())
        if rate <= 0:
            raise ValueError
    except Exception:
        st.error("Rate per KM is required and must be a number > 0.")
        st.stop()

    def fnum(x, default=0.0):
        try:
            s = str(x).strip()
            if s == "":
                return default
            return float(s)
        except Exception:
            return default

    vendor_basic_val = fnum(vendor_basic, default=float("nan"))
    vendor_basic_val = None if pd.isna(vendor_basic_val) else vendor_basic_val

    pan4_val = str(pan4).strip().upper()
    pan4_val = None if pan4_val == "" else pan4_val[0]

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

            # zip outputs
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
                file_name="SLA_Output_Files.zip",
                mime="application/zip"
            )

