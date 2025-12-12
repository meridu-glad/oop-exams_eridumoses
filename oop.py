# import libraries
import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import re
from collections import defaultdict, Counter
from datetime import datetime
import math
import difflib

# --- Helpers (pure Python & Streamlit session management) ---

def now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_action(msg):
    """Logs actions to a persistent audit log in session state."""
    if st.session_state.get("audit_log") is None:
        st.session_state.audit_log = []
    st.session_state.audit_log.insert(0, f"{now_ts()} — {msg}")

def df_to_bytes(df):
    """Converts a DataFrame to an in-memory Excel buffer for download."""
    try:
        import openpyxl  # noqa: F401 (ensure the library is available)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="cleaned")
        buf.seek(0)
        return buf.getvalue(), "excel"
    except ImportError:
        # Fallback to CSV if openpyxl isn't installed
        return df.to_csv(index=False).encode("utf-8"), "csv"

def basic_clean(df):
    """Applies basic cleaning: title case, strip whitespace, handle NaNs."""
    df = df.copy()
    for c in df.select_dtypes(include=["object", "string"]).columns:
        df[c] = df[c].astype(str).replace({"nan": np.nan, "None": np.nan})
        df[c] = df[c].where(df[c].notnull(), np.nan)
        # Use .apply(str) safely within the lambda for non-null items
        df[c] = df[c].apply(lambda x: " ".join(str(x).strip().title().split()) if pd.notna(x) else x)
    return df

def detect_exact_groups(df, key_cols):
    """Detects groups of exact duplicates based on key columns."""
    if not key_cols:
        return []
    # Use reset_index(drop=False) to get actual dataframe indices
    grouped = df.groupby(key_cols, dropna=False).groups
    # Filter for groups that have more than one entry
    groups = [list(indices) for indices in grouped.values() if len(indices) > 1]
    return groups

# --- Fuzzy Matching Helpers (using built-in difflib) ---

def normalize_tokens(s: str) -> list:
    tokens = re.findall(r"\w+", s.lower())
    return tokens

def token_sort_key(s: str) -> str:
    toks = normalize_tokens(s)
    toks.sort()
    return " ".join(toks)

def token_sort_score(a: str, b: str) -> int:
    """Token-sort similarity using difflib.SequenceMatcher (0..100)."""
    na = token_sort_key(a)
    nb = token_sort_key(b)
    # difflib.SequenceMatcher works well for token-sorted strings
    ratio = difflib.SequenceMatcher(None, na, nb).ratio()
    return int(round(ratio * 100))

def soundex_code(s: str) -> str:
    """Basic Soundex implementation: returns 4-character code."""
    s = re.sub(r"[^A-Za-z]", "", str(s).upper())
    if not s: return ""
    first_letter = s[0]
    mappings = {c: "1" for c in "BFPV"}
    mappings.update({c: "2" for c in "CGJKQSXZ"})
    mappings.update({c: "3" for c in "DT"})
    mappings.update({c: "4" for c in "L"})
    mappings.update({c: "5" for c in "MN"})
    mappings.update({c: "6" for c in "R"})

    digits = []
    prev = mappings.get(first_letter, "0")
    for ch in s[1:]:
        code = mappings.get(ch, "0")
        if code != prev:
            digits.append(code)
        prev = code

    cleaned = [d for d in digits if d != "0"]
    code = first_letter + ("".join(cleaned) + "000")[:3]
    return code

def fuzzy_pairs_hybrid(df, cols, threshold=85, max_pairs=2000000, block_by=None):
    """
    Detect likely duplicate pairs using token-sort similarity and blocking.
    Note: For large datasets, this pure Python implementation can be slow.
    """
    start = time.time()
    n = len(df)
    if n <= 1:
        return [], 0.0

    # Prepare comparison texts (concatenated fields)
    texts = df[cols].fillna("").astype(str).agg(" | ".join, axis=1).tolist()
    norm_texts = [token_sort_key(t) for t in texts]

    # Build blocks using normalized text keys
    blocks = defaultdict(list)
    for i, key in enumerate(norm_texts):
        if key: # Don't block on empty strings
            blocks[key].append(i)
    
    pairs_list = []
    # We use a set to avoid processing the same pair twice (i, j)
    seen_pairs = set() 
    
    status_placeholder = st.empty()
    status_placeholder.markdown(f'<p class="progress-label">Processing {len(blocks)} blocks...</p>', unsafe_allow_html=True)

    # Iterate through blocks and find pairs within
    for block_key, indices in blocks.items():
        if len(indices) < 2:
            continue
            
        for i_idx in range(len(indices)):
            for j_idx in range(i_idx + 1, len(indices)):
                idx1 = indices[i_idx]
                idx2 = indices[j_idx]

                if (idx1, idx2) in seen_pairs:
                    continue
                
                # Compare only if they haven't been seen via another block
                score = token_sort_score(texts[idx1], texts[idx2])
                
                if score >= threshold:
                    pairs_list.append((idx1, idx2, score))
                    seen_pairs.add((idx1, idx2))
                    
                    if len(pairs_list) >= max_pairs:
                        st.warning(f"Stopped processing: Maximum pairs limit reached ({max_pairs}).")
                        end = time.time()
                        return pairs_list, end - start

    end = time.time()
    status_placeholder.empty()
    return pairs_list, end - start


# --- Streamlit App UI and Logic ---

# Page configuration & styling
st.set_page_config(
    page_title="PLAYMATTERS DATABASE APP",
    layout="wide",
    initial_sidebar_state="auto"
)

# Inject custom CSS for styling
st.markdown("""
<style>
/* ... (CSS provided by user, slightly cleaned up) ... */
.title { text-align: center; font-size: 36px; font-weight: 800; margin-bottom: 6px; color: #1E3A8A; font-family: 'Segoe UI', Tahoma, sans-serif; }
.subtitle { text-align: center; font-size: 16px; margin-top: 0px; color: #475569; font-family: 'Segoe UI', Tahoma, sans-serif; }
.developer { position: fixed; right: 14px; bottom: 10px; font-style: italic; color: #1E3A8A; font-size: 14px; }
.stButton>button, .stDownloadButton>button { background-color: #2563EB; color: white; font-weight: 700; border-radius: 8px; padding: 8px 16px; font-size: 14px; transition: background-color 0.3s ease; }
.stButton>button:hover, .stDownloadButton>button:hover { background-color: #1E40AF; }
.progress-label { font-weight: 700; color: #1E293B; }
table.data { border-collapse: collapse; width: 100%; }
table.data td, th { border: 1px solid #ddd; padding: 8px; }
</style>
""", unsafe_allow_html=True)

# Render title and subtitle
st.markdown('<div class="title">PLAYMATTERS DATABASE APP</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Exact and Fuzzy Deduplication for Attendance Records</div>', unsafe_allow_html=True)


# --- Main Application Flow ---

# Initialize session state variables
if 'df_cleaned' not in st.session_state:
    st.session_state.df_cleaned = None
if 'df_original' not in st.session_state:
    st.session_state.df_original = None
if 'fuzzy_pairs' not in st.session_state:
    st.session_state.fuzzy_pairs = None
if 'audit_log' not in st.session_state:
    st.session_state.audit_log = []


st.sidebar.header("1. Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload your Excel or CSV file", type=['csv', 'xlsx', 'xls'])

if uploaded_file is not None:
    # Load the file into a DataFrame
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        st.session_state.df_original = df
        log_action(f"File uploaded: {uploaded_file.name} with {len(df)} rows.")

    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

# --- SIDEBAR: Cleaning Controls ---
if st.session_state.df_original is not None:
    st.sidebar.header("2. Cleaning & Prep")
    
    if st.sidebar.button("Run Basic Cleaning & Standardize"):
        st.session_state.df_cleaned = basic_clean(st.session_state.df_original)
        log_action("Basic cleaning applied (Title case, NaN handling).")
        st.sidebar.success("Cleaning applied! Check main view.")


    st.sidebar.header("3. Deduplication (Exact Match)")
    
    current_df = st.session_state.df_cleaned if st.session_state.df_cleaned is not None else st.session_state.df_original
    cols = current_df.columns.tolist()
    exact_cols_to_check = st.sidebar.multiselect(
        "Select columns for EXACT match check:", 
        options=cols, 
        default=[cols[0]] if cols else []
    )

    if st.sidebar.button("Detect EXACT Duplicates"):
        if exact_cols_to_check:
            groups = detect_exact_groups(current_df, exact_cols_to_check)
            if groups:
                total_duplicates = sum(len(g) for g in groups) - len(groups) # count extra rows, not groups
                st.session_state.exact_groups = groups
                st.session_state.exact_duplicates_count = total_duplicates
                log_action(f"Found {total_duplicates} exact duplicates across {len(groups)} groups.")
                st.sidebar.info(f"Found {total_duplicates} exact duplicate rows.")
            else:
                st.sidebar.success("No exact duplicates found with selected columns.")
                st.session_state.exact_groups = []
                st.session_state.exact_duplicates_count = 0
        else:
            st.sidebar.warning("Please select columns for exact match detection.")
    
    
    st.sidebar.header("4. Deduplication (Fuzzy Match)")

    fuzzy_cols_to_check = st.sidebar.multiselect(
        "Select columns for FUZZY match (concatenated):", 
        options=cols, 
        default=[]
    )
    fuzzy_threshold = st.sidebar.slider("Fuzzy Match Threshold (%)", min_value=75, max_value=100, value=90)

    if st.sidebar.button("Run FUZZY Deduplication"):
        if fuzzy_cols_to_check:
            with st.spinner("Running complex fuzzy matching... this might take time for large data."):
                pairs, elapsed_time = fuzzy_pairs_hybrid(
                    df=current_df, 
                    cols=fuzzy_cols_to_check, 
                    threshold=fuzzy_threshold
                )
                st.session_state.fuzzy_pairs = pairs
                log_action(f"Fuzzy match complete in {elapsed_time:.2f}s. Found {len(pairs)} potential fuzzy pairs.")
                st.sidebar.success(f"Found {len(pairs)} potential fuzzy pairs.")
        else:
            st.sidebar.warning("Please select columns for fuzzy matching.")


# --- MAIN CONTENT AREA ---

# Displaying the data
if st.session_state.df_original is None:
    st.info("Upload a file in the sidebar to begin data cleaning and deduplication.")
else:
    st.header("Data View")
    
    view_option = st.radio(
        "Select Data View:",
        ["Original Data", "Cleaned Data", "Fuzzy Match Results", "Audit Log"],
        horizontal=True
    )

    if view_option == "Original Data":
        st.dataframe(st.session_state.df_original)
        st.markdown(f"**Total Rows:** {len(st.session_state.df_original)}")
        
    elif view_option == "Cleaned Data" and st.session_state.df_cleaned is not None:
        st.dataframe(st.session_state.df_cleaned)
        st.markdown(f"**Total Rows:** {len(st.session_state.df_cleaned)}")
        
        # Add download button for cleaned data
        excel_bytes, file_type = df_to_bytes(st.session_state.df_cleaned)
        st.download_button(
            label=f"Download Cleaned Data as {'Excel' if file_type == 'excel' else 'CSV'}",
            data=excel_bytes,
            file_name=f"cleaned_data_{now_ts().replace(' ', '_').replace(':', '-')}.{'xlsx' if file_type == 'excel' else 'csv'}",
            mime=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if file_type == 'excel' else 'csv'}"
        )
        
    elif view_option == "Fuzzy Match Results" and st.session_state.fuzzy_pairs is not None:
        st.subheader(f"Potential Fuzzy Matches ({len(st.session_state.fuzzy_pairs)} pairs)")
        
        current_df_view = st.session_state.df_cleaned if st.session_state.df_cleaned is not None else st.session_state.df_original
        
        # Display fuzzy pairs in a digestible format (showing concatenated column values)
        pairs_data = []
        
        # Prepare the concatenated texts again for viewing ease
        view_cols = fuzzy_cols_to_check if fuzzy_cols_to_check else current_df_view.columns[:3] # Default to first 3 if none selected
        view_texts = current_df_view[view_cols].fillna("").astype(str).agg(" | ".join, axis=1).tolist()
        
        for idx1, idx2, score in st.session_state.fuzzy_pairs:
            pairs_data.append({
                'Record A Index': idx1,
                'Record B Index': idx2,
                'Similarity Score (%)': score,
                'Record A Snippet': view_texts[idx1][:100] + ('...' if len(view_texts[idx1]) > 100 else ''),
                'Record B Snippet': view_texts[idx2][:100] + ('...' if len(view_texts[idx2]) > 100 else ''),
            })
        
        st.dataframe(pd.DataFrame(pairs_data), height=500)
        st.info("These are potential matches that require manual review.")

    elif view_option == "Audit Log":
        st.subheader("Application Audit Log")
        st.code("\n".join(st.session_state.audit_log))
        
    else:
        st.warning("Please run the previous steps to view this data.")


# Sticky footer for developer credit
st.markdown('<div class="developer">Developed by Eridu Moses </div>', unsafe_allow_html=True)
