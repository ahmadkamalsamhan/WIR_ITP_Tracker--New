import streamlit as st
import pandas as pd
import re
from io import BytesIO
from collections import defaultdict

st.title("📊 ITP-WIR Activity Status (Optimized)")

def preprocess_text(text):
    if pd.isna(text):
        return []
    text = re.sub(r'[^a-zA-Z0-9\s]', '', str(text).lower())
    return text.split()

def assign_status(code):
    if pd.isna(code):
        return 0
    code = str(code).strip().upper()
    if code in ['A','B']:
        return 1
    elif code in ['C','D']:
        return 2
    return 0

def token_match_score(tokens1, tokens2):
    common = set(tokens1) & set(tokens2)
    return len(common)/max(len(tokens1), 1)

# Upload files
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
wir_file = st.file_uploader("Upload WIR Log", type=["xlsx"])

if itp_file and activity_file and wir_file:
    itp_log = pd.read_excel(itp_file)
    activity_log = pd.read_excel(activity_file)
    wir_log = pd.read_excel(wir_file)

    itp_log.columns = itp_log.columns.str.strip()
    activity_log.columns = activity_log.columns.str.strip()
    wir_log.columns = wir_log.columns.str.strip()

    st.success("✅ Files uploaded successfully!")

    # Column selection
    itp_no_col = st.selectbox("ITP No.", itp_log.columns)
    itp_title_col = st.selectbox("ITP Title", itp_log.columns)
    activity_desc_col = st.selectbox("Activity Description", activity_log.columns)
    itp_ref_col = st.selectbox("ITP Reference", activity_log.columns)
    activity_no_col = st.selectbox("Activity No.", activity_log.columns)
    wir_title_col = st.selectbox("WIR Title (Title / Description2)", wir_log.columns)
    wir_pm_col = st.selectbox("PM Web Code", wir_log.columns)

    if st.button("Generate WIR Status (Optimized)"):
        st.info("Processing...")

        # Preprocess WIR tokens
        wir_log['Title_Tokens'] = wir_log[wir_title_col].apply(preprocess_text)

        # Build token -> WIR index mapping
        token_dict = defaultdict(list)
        for idx, row in wir_log.iterrows():
            for token in row['Title_Tokens']:
                token_dict[token].append(idx)

        # Prepare output columns
        status_list = []
        match_score_list = []

        total_rows = len(activity_log)
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, act_row in activity_log.iterrows():
            itp_no = act_row[itp_ref_col]
            activity_desc = act_row[activity_desc_col]
            activity_tokens = preprocess_text(activity_desc)

            # Get ITP Title
            itp_row = itp_log[itp_log[itp_no_col]==itp_no]
            if itp_row.empty:
                status_list.append(0)
                match_score_list.append(0)
                progress_bar.progress((i+1)/total_rows)
                status_text.text(f"Processing row {i+1} of {total_rows}")
                continue
            itp_title = itp_row.iloc[0][itp_title_col]
            itp_tokens = preprocess_text(itp_title)

            # Candidate WIRs via any token overlap
            candidate_indices = set()
            for token in itp_tokens:
                candidate_indices.update(token_dict.get(token, []))

            # Find best WIR
            best_score = 0
            best_wir = None
            for idx in candidate_indices:
                wir_row = wir_log.iloc[idx]
                score = token_match_score(itp_tokens, wir_row['Title_Tokens'])
                if score > best_score:
                    best_score = score
                    best_wir = wir_row

            if best_wir is not None and best_score > 0:
                final_score = round(best_score*100, 1)
                status_code = assign_status(best_wir[wir_pm_col])
            else:
                final_score = 0
                status_code = 0

            status_list.append(status_code)
            match_score_list.append(final_score)

            # Update progress
            progress_bar.progress((i+1)/total_rows)
            status_text.text(f"Processing row {i+1} of {total_rows}")

        progress_bar.empty()
        status_text.empty()

        activity_log['WIR Status Code'] = status_list
        activity_log['Match Score (%)'] = match_score_list

        st.success("✅ Status and Match Score added!")
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
