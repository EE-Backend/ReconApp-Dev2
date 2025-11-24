import streamlit as st
import pandas as pd
from pathlib import Path
from recon_engine import generate_reconciliation_file

# ============================================
# PAGE CONFIG
# ============================================
st.set_page_config(page_title="Recon File Generator", layout="wide")

# Base directory of this file
BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
LOGO_PATH = STATIC_DIR / "logo.png"

# ============================================
# HEADER WITH LOGO + TITLE
# ============================================
col_logo, col_title = st.columns([1, 15])

with col_logo:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=120)
    else:
        st.warning(f"Logo not found at {LOGO_PATH}")

with col_title:
    st.markdown(
        """
        <h1 style='margin-top: 10px; margin-bottom: 0px;'>
            Recon File Generator
        </h1>
        """,
        unsafe_allow_html=True
    )

st.markdown("<div style='height:20px;'></div>", unsafe_allow_html=True)  # small spacing

# ============================================
# STEP 1 — INPUTS
# ============================================
st.markdown("## Step 1 — Upload Inputs")

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

icp_code = st.text_input("Enter ICP Code", placeholder="Example: SKPVAB").strip()

st.markdown("<div style='height:30px;'></div>", unsafe_allow_html=True)

quarter_year = st.text_input(
    "Quarter & Year",
    placeholder="Q4 2025"
).strip()


# ============================================
# STEP 2 — GENERATE
# ============================================
st.markdown("## Step 2 — Generate Recon File")

if st.button("Generate Recon File", type="primary"):

    if not trial_balance_file or not entries_file or not icp_code or not quarter_year:
        st.error("Please upload both files and enter the ICP Code.")
        st.stop()

    with st.spinner("Generating file, please wait..."):
        try:
            output_bytes = generate_reconciliation_file(
                trial_balance_file,
                entries_file,
                icp_code.upper(),
                quarter_year
            )
        except Exception as e:
            st.error(f"❌ An error occurred:\n\n{e}")
            st.stop()

    st.success("✅ Reconciliation file generated successfully!")

    st.download_button(
        label="📥 Download Reconciliation Workbook",
        data=output_bytes,
        file_name=f"Recon File {icp_code.upper()} {quarter_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("European Energy — Internal Tool")






