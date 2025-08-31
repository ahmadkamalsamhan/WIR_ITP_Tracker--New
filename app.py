import streamlit as st
import pandas as pd
import re
import tempfile
from io import BytesIO
import time

st.set_page_config(page_title="ITP-WIR Matching App", layout="wide")
st.title("📊 ITP-WIR Activity Status Matching (Fast & Memory-Safe)")

# -------------------------------
# Preprocessing functions
# -------------------------------
def preprocess_text(text):
    """Normalize, remove special chars, split into tokens set"""
    if pd.isna(text):
        return set()
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return set(text.split())

def assign_status(code):
    """Convert PM Web Code to 0/1/2"""
    if pd.isna(code):
        return 0
    code = str(code).strip().upper()
    if code in ['A','B']:
        return 1
    elif code in ['C','D']:
        return 2
    return 0

# -------------------------------
# Upload files
# -------------------------------
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
wir_file = st.file_uploader("Upload WIR Log (Document Control Log)", type=["xlsx"])

if itp_file and activity_file and wir_file:
    itp_log = pd.read_excel(itp_file)
    activity_log = pd.read_excel(activity_file)
    wir_log = pd.read_excel(wir_file)

    # Clean columns
    itp_log.columns = itp_log.columns.str.strip()
    activity_log.columns = activity_log.columns.str.strip()
    wir_log.columns = wir_log.columns.str.strip()

    st.success("✅ Files uploaded successfully!")

    # -------------------------------
    # Column Selection
    # -------------------------------
    st.subheader("Select Columns")
    itp_no_col = st.selectbox("ITP No.", itp_log.columns)
    itp_title_col = st.selectbox("ITP Title", itp_log.columns)
    activity_desc_col = st.selectbox("Activity Description", activity_log.columns)
    itp_ref_col = st.selectbox("ITP Reference", activity_log.columns)
    activity_no_col = st.selectbox("Activity No.", activity_log.columns)
    wir_title_col = st.selectbox("WIR Title (Title / Description2)", wir_log.columns)
    wir_pm_col = st.selectbox("PM Web Code", wir_log.columns)

    # -------------------------------
    # Matching Button
    # -------------------------------
    if st.button("Generate WIR Status (Fast)"):
        st.info("⏳ Processing...")

        # Preprocess WIR Titles and ITP Titles
        wir_log['WIR_Tokens'] = wir_log[wir_title_col].apply(preprocess_text)
        itp_log['ITP_Tokens'] = itp_log[itp_title_col].apply(preprocess_text)

        # Create mapping from ITP No -> ITP Title tokens
        itp_tokens_map = dict(zip(itp_log[itp_no_col], itp_log['ITP_Tokens']))

        # Preprocess Activity Descriptions
        activity_log['Activity_Tokens'] = activity_log[activity_desc_col].apply(preprocess_text)

        # Prepare results lists
        status_list = []
        score_list = []

        total_rows = len(activity_log)
        progress_bar = st.progress(0)
        status_text = st.empty()
        start_time = time.time()

        # Batch-wise processing
        for i, row in activity_log.iterrows():
            itp_no = row[itp_ref_col]
            act_tokens = row['Activity_Tokens']

            itp_tokens = itp_tokens_map.get(itp_no, set())

            # Match WIRs: token overlap
            best_score = 0
            best_pm_code = None

            for _, wir_row in wir_log.iterrows():
                wir_tokens = wir_row['WIR_Tokens']
                # token overlap score
                common = itp_tokens & wir_tokens
                score = len(common) / max(len(itp_tokens), 1)
                if score > best_score:
                    best_score = score
                    best_pm_code = wir_row[wir_pm_col]

            # Assign WIR Status Code
            status_list.append(assign_status(best_pm_code))
            score_list.append(round(best_score*100, 1))

            # Update progress
            if (i % 10 == 0) or (i == total_rows-1):
                progress_bar.progress((i+1)/total_rows)
                status_text.text(f"Processing row {i+1}/{total_rows}")

        # Append results
        activity_log['WIR Status Code'] = status_list
        activity_log['Match Score (%)'] = score_list

        end_time = time.time()
        st.success(f"✅ Completed in {end_time - start_time:.2f} seconds")
        st.dataframe(activity_log)

        # Download Excel
        output = BytesIO()
        activity_log.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        st.download_button(
            "📥 Download Updated Activity Log",
            data=output,
            file_name="ITP_Activities_With_WIR_Status.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
