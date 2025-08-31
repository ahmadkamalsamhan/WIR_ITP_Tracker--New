import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 ITP-WIR Activity Status with Match Score")

# -------------------------------
# Preprocessing
# -------------------------------
def preprocess_text(text):
    """Lowercase, remove special chars, split into tokens"""
    if pd.isna(text):
        return []
    text = re.sub(r'[^a-zA-Z0-9\s]', '', str(text).lower())
    tokens = text.split()
    return tokens

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

def token_match_score(title_tokens, target_tokens):
    """Compute token overlap ratio"""
    common = set(title_tokens) & set(target_tokens)
    score = len(common) / max(len(title_tokens), 1)
    return score

# -------------------------------
# File Upload
# -------------------------------
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
wir_file = st.file_uploader("Upload WIR Log (Document Control Log)", type=["xlsx"])

if itp_file and activity_file and wir_file:
    itp_log = pd.read_excel(itp_file)
    activity_log = pd.read_excel(activity_file)
    wir_log = pd.read_excel(wir_file)

    # Strip spaces/newlines from columns
    itp_log.columns = itp_log.columns.str.strip()
    activity_log.columns = activity_log.columns.str.strip()
    wir_log.columns = wir_log.columns.str.strip()

    st.success("✅ Files uploaded successfully!")

    # -------------------------------
    # Column Selection
    # -------------------------------
    st.subheader("ITP Log Columns")
    itp_no_col = st.selectbox("Select ITP No.", itp_log.columns)
    itp_title_col = st.selectbox("Select ITP Title", itp_log.columns)

    st.subheader("Activity Log Columns")
    activity_desc_col = st.selectbox("Select Activity Description", activity_log.columns)
    itp_ref_col = st.selectbox("Select ITP Reference", activity_log.columns)
    activity_no_col = st.selectbox("Select Activity No.", activity_log.columns)

    st.subheader("WIR Log Columns")
    wir_title_col = st.selectbox("Select WIR Title (Title / Description2)", wir_log.columns)
    wir_pm_col = st.selectbox("Select PM Web Code", wir_log.columns)

    # -------------------------------
    # Generate Status Column
    # -------------------------------
    if st.button("Generate WIR Status for Activities"):
        st.info("Processing...")

        # Preprocess WIR titles
        wir_log['Title_Tokens'] = wir_log[wir_title_col].apply(preprocess_text)

        status_list = []
        match_score_list = []

        for _, act_row in activity_log.iterrows():
            itp_no = act_row[itp_ref_col]
            activity_desc = act_row[activity_desc_col]
            activity_tokens = preprocess_text(activity_desc)

            # Get ITP Title
            itp_row = itp_log[itp_log[itp_no_col]==itp_no]
            if itp_row.empty:
                status_list.append(0)
                match_score_list.append(0)
                continue
            itp_title = itp_row.iloc[0][itp_title_col]
            itp_tokens = preprocess_text(itp_title)

            # Match WIR titles by token overlap
            best_score = 0
            best_wir = None
            for idx, wir_row in wir_log.iterrows():
                score = token_match_score(itp_tokens, wir_row['Title_Tokens'])
                if score > best_score:
                    best_score = score
                    best_wir = wir_row

            if best_wir is not None and best_score > 0:
                # Optional: activity closeness (can improve later)
                final_score = round(best_score*100, 1)
                status_code = assign_status(best_wir[wir_pm_col])
            else:
                final_score = 0
                status_code = 0

            status_list.append(status_code)
            match_score_list.append(final_score)

        activity_log['WIR Status Code'] = status_list
        activity_log['Match Score (%)'] = match_score_list

        st.success("✅ Status and Match Score added!")
        st.dataframe(activity_log)

        # Download updated Excel
        output = BytesIO()
        activity_log.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            "📥 Download Updated Activity Log",
            data=output,
            file_name="ITP_Activities_With_WIR_Status.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
