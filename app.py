import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 ITP-WIR Activity Status Updater")

# -------------------------------
# Preprocessing
# -------------------------------
def preprocess_text(text):
    if pd.isna(text):
        return ""
    return str(text).lower().strip()

# -------------------------------
# Assign Status
# -------------------------------
def assign_status(code):
    if pd.isna(code):
        return 0
    code = str(code).strip().upper()
    if code in ['A','B']:
        return 1
    elif code in ['C','D']:
        return 2
    return 0

# -------------------------------
# Upload Files
# -------------------------------
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
wir_file = st.file_uploader("Upload WIR Log (Document Control Log)", type=["xlsx"])

if activity_file and wir_file:
    activity_log = pd.read_excel(activity_file)
    wir_log = pd.read_excel(wir_file)

    # Clean column names
    activity_log.columns = activity_log.columns.str.strip().str.replace('\n','').str.replace('\r','')
    wir_log.columns = wir_log.columns.str.strip().str.replace('\n','').str.replace('\r','')

    st.success("✅ Files uploaded successfully!")

    # -------------------------------
    # Column Selection
    # -------------------------------
    st.subheader("Activity Log Columns")
    activity_desc_col = st.selectbox("Select Activity Description column", options=activity_log.columns.tolist())
    itp_ref_col = st.selectbox("Select ITP Reference column", options=activity_log.columns.tolist())

    st.subheader("WIR Log Columns")
    wir_title_col = st.selectbox("Select WIR Title column (Title / Description2)", options=wir_log.columns.tolist())
    wir_pm_col = st.selectbox("Select PM Web Code column", options=wir_log.columns.tolist())

    if st.button("Add Status Column to Activity Log"):
        st.info("Processing...")

        # Build WIR lookup dictionary (title -> PM Web Code)
        wir_lookup = {}
        for idx, row in wir_log.iterrows():
            title = preprocess_text(row[wir_title_col])
            wir_lookup[title] = row[wir_pm_col]

        # Create a new column for status
        status_list = []

        for _, activity_row in activity_log.iterrows():
            activity_desc = preprocess_text(activity_row[activity_desc_col])
            status_code = 0

            # Match activity with WIR title
            for wir_title, pm_code in wir_lookup.items():
                if activity_desc in wir_title:
                    status_code = assign_status(pm_code)
                    break

            status_list.append(status_code)

        # Add new column
        activity_log['WIR Status Code'] = status_list

        st.success("✅ Status column added!")
        st.dataframe(activity_log)

        # Download updated file
        output = BytesIO()
        activity_log.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="📥 Download Updated Activity Log",
            data=output,
            file_name="ITP_Activities_With_Status.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
