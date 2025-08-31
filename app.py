import streamlit as st
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO
import time
from sentence_transformers import SentenceTransformer, util

st.set_page_config(page_title="ITP-WIR Semantic Matching", layout="wide")
st.title("📊 ITP-WIR Matching with Semantic Accuracy")

# -------------------------------
# Text preprocessing
# -------------------------------
def preprocess_text(text):
    if pd.isna(text):
        return set()
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return set(text.split())

def assign_status(code):
    if pd.isna(code):
        return 0
    code = str(code).upper().strip()
    if code in ['A','B']:
        return 1
    elif code in ['C','D']:
        return 2
    return 0

# Load small, fast sentence transformer model
@st.cache_resource
def load_model():
    return SentenceTransformer('all-MiniLM-L6-v2')

model = load_model()

# -------------------------------
# Tabs for Part 1 and Part 2
# -------------------------------
tab1, tab2 = st.tabs(["Part 1: Title Matching", "Part 2: Activity Matching"])

# ===============================
# Part 1: Title Matching
# ===============================
with tab1:
    st.header("🔹 Part 1: WIR ↔ ITP Title Matching (Semantic + Token)")

    wir_file = st.file_uploader("Upload WIR Log (Document Control Log)", type=["xlsx"], key="wir1")
    itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"], key="itp1")

    if wir_file and itp_file:
        wir_log = pd.read_excel(wir_file)
        itp_log = pd.read_excel(itp_file)

        wir_log.columns = wir_log.columns.str.strip()
        itp_log.columns = itp_log.columns.str.strip()

        # Column selection
        wir_doc_col = st.selectbox("WIR Document No.", wir_log.columns, key="wir_doc")
        wir_title_col = st.selectbox("WIR Title (Title / Description2)", wir_log.columns, key="wir_title")
        wir_pm_col = st.selectbox("WIR PM Web Code", wir_log.columns, key="wir_pm")

        itp_no_col = st.selectbox("ITP No.", itp_log.columns, key="itp_no")
        itp_title_col = st.selectbox("ITP Title (Title / Description)", itp_log.columns, key="itp_title")

        if st.button("Start Title Matching"):
            st.info("⏳ Matching WIR titles with ITP titles...")

            start_time = time.time()

            # Preprocess tokens
            wir_log['WIR_Tokens'] = wir_log[wir_title_col].apply(preprocess_text)
            itp_log['ITP_Tokens'] = itp_log[itp_title_col].apply(preprocess_text)

            # Build token → ITP indices lookup
            token_to_itp = defaultdict(set)
            for idx, tokens in enumerate(itp_log['ITP_Tokens']):
                for token in tokens:
                    token_to_itp[token].add(idx)

            # Precompute embeddings for ITP titles
            itp_log['ITP_Embedding'] = list(model.encode(itp_log[itp_title_col].astype(str).tolist(), convert_to_tensor=True))

            matched_rows = []
            total_rows = len(wir_log)
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, row in wir_log.iterrows():
                wir_tokens = row['WIR_Tokens']
                # Candidate filtering via token overlap
                candidate_indices = set()
                for token in wir_tokens:
                    candidate_indices.update(token_to_itp.get(token, set()))

                best_score = 0
                best_idx = None

                if candidate_indices:
                    wir_emb = model.encode(str(row[wir_title_col]), convert_to_tensor=True)
                    for idx in candidate_indices:
                        itp_emb = itp_log.at[idx, 'ITP_Embedding']
                        semantic_score = util.cos_sim(wir_emb, itp_emb).item()  # 0-1
                        token_score = len(wir_tokens & itp_log.at[idx,'ITP_Tokens']) / max(len(wir_tokens),1)
                        combined_score = 0.5*token_score + 0.5*semantic_score
                        if combined_score > best_score:
                            best_score = combined_score
                            best_idx = idx

                if best_idx is not None:
                    itp_row = itp_log.loc[best_idx]
                    matched_rows.append({
                        "WIR Document No": row[wir_doc_col],
                        "WIR Title": row[wir_title_col],
                        "ITP No": itp_row[itp_no_col],
                        "ITP Title": itp_row[itp_title_col],
                        "PM Web Code": row[wir_pm_col],
                        "Match Score (%)": round(best_score*100,1)
                    })

                if i % 10 == 0 or i == total_rows -1:
                    progress_bar.progress((i+1)/total_rows)
                    status_text.text(f"Processing {i+1}/{total_rows}")

            result_df = pd.DataFrame(matched_rows)
            st.success(f"✅ Completed in {time.time()-start_time:.2f} seconds")
            st.dataframe(result_df)

            # Download
            output = BytesIO()
            result_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            st.download_button("📥 Download Part 1 Result", data=output,
                               file_name="Part1_WIR_ITP_SemanticMatch.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ===============================
# Part 2: Activity Matching
# ===============================
with tab2:
    st.header("🔹 Part 2: Activity Matching")

    part1_file = st.file_uploader("Upload Part 1 Result Excel", type=["xlsx"], key="part1")
    activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"], key="activity2")

    if part1_file and activity_file:
        part1_df = pd.read_excel(part1_file)
        activity_log = pd.read_excel(activity_file)

        part1_df.columns = part1_df.columns.str.strip()
        activity_log.columns = activity_log.columns.str.strip()

        # Column selection
        activity_desc_col = st.selectbox("Activity Description", activity_log.columns, key="act_desc")
        itp_ref_col = st.selectbox("ITP Reference in Activities", activity_log.columns, key="act_itp_ref")
        activity_no_col = st.selectbox("Activity No.", activity_log.columns, key="act_no")

        if st.button("Start Activity Matching"):
            st.info("⏳ Matching Activities with WIRs...")

            start_time = time.time()

            activity_log['Activity_Tokens'] = activity_log[activity_desc_col].apply(preprocess_text)
            part1_df['ITP_Tokens'] = part1_df['ITP Title'].apply(preprocess_text)

            status_list = []
            score_list = []

            total_rows = len(activity_log)
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, row in activity_log.iterrows():
                act_tokens = row['Activity_Tokens']
                itp_no = row[itp_ref_col]
                matched_itp_row = part1_df[part1_df['ITP No']==itp_no]

                if not matched_itp_row.empty:
                    itp_tokens = matched_itp_row.iloc[0]['ITP_Tokens']
                    token_score = len(act_tokens & itp_tokens)/max(len(itp_tokens),1)
                    # Optionally semantic score for activity vs ITP title
                    activity_emb = model.encode(str(row[activity_desc_col]), convert_to_tensor=True)
                    itp_emb = model.encode(str(matched_itp_row.iloc[0]['ITP Title']), convert_to_tensor=True)
                    semantic_score = util.cos_sim(activity_emb, itp_emb).item()
                    combined_score = 0.5*token_score + 0.5*semantic_score

                    score_list.append(round(combined_score*100,1))
                    pm_code = matched_itp_row.iloc[0]['PM Web Code']
                    status_list.append(assign_status(pm_code))
                else:
                    score_list.append(0)
                    status_list.append(0)

                if i % 20 == 0 or i == total_rows-1:
                    progress_bar.progress((i+1)/total_rows)
                    status_text.text(f"Processing {i+1}/{total_rows}")

            activity_log['WIR Status Code'] = status_list
            activity_log['Match Score (%)'] = score_list

            st.success(f"✅ Completed in {time.time()-start_time:.2f} seconds")
            st.dataframe(activity_log)

            # Download
            output = BytesIO()
            activity_log.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            st.download_button("📥 Download Activity Matched Result", data=output,
                               file_name="Part2_Activity_SemanticMatch.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
