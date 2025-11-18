import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

from recon_engine import generate_reconciliation_file  # your backend logic


# -------------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------------
st.set_page_config(
    page_title="Recon File Generator",
    layout="wide"
)

# -------------------------------------------------------
# PATHS
# -------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

logo_path = STATIC_DIR / "logo.png"   # your actual logo


# -------------------------------------------------------
# HEADER (bigger logo + tighter spacing + inline layout)
# -------------------------------------------------------
col1, col2 = st.columns([1, 10])

with col1:
    if logo_path.exists():
        st.image(str(logo_path), width=110)   # <-- 50% bigger (previously 70)
    else:
        st.warning(f"⚠ Logo not found at: {logo_path}")

with col2:
    st.markdown(
        """
        <div style="display:flex; flex-direction:column; justify-content:center; margin-top:10px;">
            <h1 style="margin-bottom:0px;">Recon File Generator</h1>
            <p style="font-size:16px; margin-top:4px; margin-bottom:0px;">
                Upload the required files below and generate a standardized reconciliation workbook.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )



# -------------------------------------------------------
# STEP 1 — FILE UPLOADS
# -------------------------------------------------------
st.header("Step 1 — Upload Inputs")

trial_balance_file = st.file_uploader(
    "Upload Trial Balance file",
    type=["xlsx"],
    key="trial_balance_upload"
)

entries_file = st.file_uploader(
    "Upload All Entries file",
    type=["xlsx"],
    key="entries_upload"
)

icp_code = st.text_input("Enter ICP Code", placeholder="Example: SKPVAB")


# -------------------------------------------------------
# STEP 2 — GENERATE BUTTON
# -------------------------------------------------------
st.write("---")
st.header("Step 2 — Generate Recon File")

generate_button = st.button("Generate Recon File", type="primary")


# -------------------------------------------------------
# PROCESS FILES
# -------------------------------------------------------
if generate_button:

    if not trial_balance_file or not entries_file or not icp_code.strip():
        st.error("❌ Please upload both files and enter an ICP code.")
        st.stop()

    with st.spinner("⏳ Generating reconciliation file..."):

        output_bytes = generate_reconciliation_file(
            trial_balance_file,
            entries_file,
            icp_code.strip().upper()
        )

    st.success("✅ Reconciliation file generated successfully!")

    st.download_button(
        label="📥 Download Reconciliation Workbook",
        data=output_bytes,
        file_name="Reconciliation_Mapped.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# -------------------------------------------------------
# FOOTER
# -------------------------------------------------------
st.write("---")
st.caption("EE Internal Tool — Powered by Streamlit")

