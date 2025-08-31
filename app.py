import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 ITP-WIR Activity Status Updater (ITP Title Verified)")

# -------------------------------
# Preprocessing
# -------------------------------
def preprocess_text(text):
    if pd.isna(text):
        return ""
    return str(text).lower().strip()

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
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
wir_file = st.file_uploader("Upload WIR Log (Document Control Log)", type=["xlsx"])

if itp_file and activity_file and wir_file:
    itp_log = pd.read_excel(itp_file)
    activity_log = pd.read_excel(activity_file)
    wir_log = pd.read_excel(wir_file)

    # Clean column names
    itp_log.columns = itp_log.columns.str.strip().str.replace('\n','').str.replace('\r','')
    activity_log.columns = activity_log.columns.str.strip().str.replace('\n','').str.replace('\r','')
    wir_log.columns = wir_log.columns.str.strip().str.replace('\n','').str.replace('\r','')

    st.success("✅ Files uploaded successfully!")

    # -------------------------------
    # Column Selection
    # -------------------------------
    st.subheader("ITP Log Columns")
    itp_no_col = st.selectbox("Select ITP No. column", options=itp_log.columns.tolist())
    itp_title_col = st.selectbox("Select ITP Title column", options=itp_log.columns.tolist())

    st.subheader("Activity Log Columns")
    activity_desc_col = st.selectbox("Select Activity Description column", options=activity_log.columns.tolist())
    itp_ref_col = st.selectbox("Select ITP Reference column", options=activity_log.columns.tolist())

    st.subheader("WIR Log Columns")
    wir_title_col = st.selectbox("Select WIR Title column (Title / Description2)", options=wir_log.columns.tolist())
    wir_pm_col = st.selectbox("Select PM Web Code column", options=wir_log.columns.tolist())

    if st.button("Add Status Column to Activity Log"):
        st.info("Processing...")

        # Preprocess WIR lookup: {ITP Title -> list of WIR rows}
        wir_lookup = {}
        for idx, row in wir_log.iterrows():
            title = preprocess_text(row[wir_title_col])
            wir_lookup[title] = row

        # Prepare status column
        status_list = []

        for _, activity_row in activity_log.iterrows():
            itp_no = activity_row[itp_ref_col]
            activity_desc = preprocess_text(activity_row[activity_desc_col])

            # Get ITP Title from ITP Log
            itp_row = itp_log[itp_log[itp_no_col] == itp_no]
            if itp_row.empty:
                status_list.append(0)
                continue
            itp_title = preprocess_text(itp_row.iloc[0][itp_title_col])

            # Match ITP Title with WIR Title
            matched_wir = None
            for wir_title, wir_row in wir_lookup.items():
                if itp_title in wir_title:
                    # Title matches, check activity in WIR Title
                    if activity_desc in wir_title:
                        matched_wir = wir_row
                        break

            if matched_wir is not None:
                status_code = assign_status(matched_wir[wir_pm_col])
            else:
                status_code = 0

            status_list.append(status_code)

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
