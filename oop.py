# import libraries
import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import math
import plotly.express as px 
from thefuzz import process, fuzz 

# --- Helpers (pure Python & Streamlit session management) ---

def now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_action(msg):
    """Logs actions to a persistent audit log in session state."""
    if st.session_state.get("audit_log") is None:
        st.session_state.audit_log = []
    st.session_state.audit_log.insert(0, f"{now_ts()} — {msg}")

def df_to_bytes(df, sheet_name="data"):
    """Converts a DataFrame to an in-memory Excel buffer for download."""
    try:
        import openpyxl 
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        buf.seek(0)
        return buf.getvalue(), "excel"
    except ImportError:
        return df.to_csv(index=False).encode("utf-8"), "csv"

def basic_clean(df):
    """Applies basic cleaning: title case, strip whitespace, handle NaNs."""
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == 'object' or pd.api.types.is_string_dtype(df[c]):
            df[c] = df[c].astype(str).replace({"nan": np.nan, "None": np.nan})
            df[c] = df[c].where(df[c].notnull(), np.nan)
            df[c] = df[c].str.strip().str.title().str.split().str.join(" ")
    return df

def detect_exact_groups(df, key_cols):
    """Detects groups of exact duplicates based on key columns, returns indices of duplicates."""
    if not key_cols:
        return {}
    grouped = df.groupby(key_cols, dropna=False).groups
    duplicate_groups = {k: list(indices) for k, indices in grouped.items() if len(indices) > 1}
    return duplicate_groups

def perform_fuzzy_detection(df, col1, col2, threshold=85):
    """Detects fuzzy matches between two columns and stores indices of matches."""
    st.info(f"Starting fuzzy matching between '{col1}' and '{col2}' with threshold {threshold}%...")
    
    # Use thefuzz to find matches between items in col1 and col2
    unique_col1 = df[col1].dropna().unique().tolist()
    unique_col2 = df[col2].dropna().unique().tolist()
    
    matches_list = []
    progress_bar = st.progress(0)
    
    # This process compares every unique item in Col 1 against every unique item in Col 2 (can be slow)
    for i, item1 in enumerate(unique_col1):
        # Find the best match in Col 2 that meets the threshold
        best_match = process.extractOne(item1, unique_col2, scorer=fuzz.token_sort_ratio, score_cutoff=threshold)
        
        if best_match:
            item2, score = best_match
            # Log the original indices where these items appear
            indices1 = df[df[col1] == item1].index.tolist()
            indices2 = df[df[col2] == item2].index.tolist()
            matches_list.append({'col1_value': item1, 'col2_value': item2, 'score': score, 'indices_col1': indices1, 'indices_col2': indices2})
        
        progress_bar.progress((i + 1) / len(unique_col1))
    
    progress_bar.empty()
    
    if matches_list:
        st.session_state.fuzzy_duplicates = matches_list
        st.success(f"Fuzzy matching detected {len(matches_list)} potential match groups.")
        log_action(f"Fuzzy detection found {len(matches_list)} matches.")
    else:
        st.info("No fuzzy matches detected with the current threshold.")

    st.text("Thank you for using this APP_Eridu.") # Appreciation note


# --- Streamlit App UI and Logic ---
st.set_page_config(page_title="PLAYMATTERS DATABASE APP", layout="wide", initial_sidebar_state="auto")

st.markdown("""
<style>
.title { text-align: center; font-size: 36px; font-weight: 800; margin-bottom: 6px; color: #1E3A8A; font-family: 'Segoe UI', Tahoma, sans-serif; }
.subtitle { text-align: center; font-size: 16px; margin-top: 0px; color: #475569; font-family: 'Segoe UI', Tahoma, sans-serif; }
.developer { position: fixed; right: 14px; bottom: 10px; font-style: italic; color: #1E3A8A; font-size: 14px; }
.stButton>button, .stDownloadButton>button { background-color: #2563EB; color: white; font-weight: 700; border-radius: 8px; padding: 8px 16px; font-size: 14px; transition: background-color 0.3s ease; }
.stButton>button:hover, .stDownloadButton>button:hover { background-color: #1E40AF; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="title">PLAYMATTERS DATABASE APP</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Deduplication, Cleaning, and Data Summarization for Attendance Records</div>', unsafe_allow_html=True)

# Initialize session state variables
if 'df_cleaned' not in st.session_state: st.session_state.df_cleaned = None
if 'df_original' not in st.session_state: st.session_state.df_original = None
if 'exact_groups' not in st.session_state: st.session_state.exact_groups = {}
if 'audit_log' not in st.session_state: st.session_state.audit_log = []
if 'confirm_refresh' not in st.session_state: st.session_state.confirm_refresh = False
if 'fuzzy_duplicates' not in st.session_state: st.session_state.fuzzy_duplicates = []
if 'fuzzy_action' not in st.session_state: st.session_state.fuzzy_action = None


def refresh_app_confirmed():
    st.session_state.clear()
    st.rerun()

def show_refresh_confirmation():
    st.session_state.confirm_refresh = True

# --- SIDEBAR: Upload and Processing Controls ---
with st.sidebar:
    st.header("1. Upload Data")

    if st.button("Refresh Application (Show Options)", type="secondary"): 
        show_refresh_confirmation()
    
    if st.session_state.confirm_refresh:
        st.warning("Are you sure you want to proceed?")
        col_clear, col_cancel = st.columns(2)
        with col_clear:
            if st.button("Clear ALL Data & Refresh", use_container_width=True):
                refresh_app_confirmed()
        with col_cancel:
            if st.button("Cancel", use_container_width=True):
                st.session_state.confirm_refresh = False
                st.rerun()
        st.stop()

    uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=['csv', 'xlsx', 'xls'])

    if uploaded_file is not None and st.session_state.df_original is None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            st.session_state.df_original = df
            st.session_state.df_cleaned = df.copy() 
            log_action(f"File uploaded: {uploaded_file.name} ({len(df)} rows).")
            st.rerun() 
        except Exception as e:
            st.error(f"Error reading file: {e}")
    
    if st.session_state.df_original is not None:
        st.header("2. Data Preparation")
        
        st.subheader("A. Missing Data Check & Handling")
        missing_data = st.session_state.df_cleaned.isnull().sum()
        missing_data = missing_data[missing_data > 0]
        if not missing_data.empty:
            # ... (missing data logic remains the same) ...
            if st.button("Apply Missing Data Action"):
                # ... (action logic remains the same) ...
                st.text("Thank you for using this APP_Eridu.") 
                st.rerun()
        else:
            st.success("No missing data detected in current view.")
        
        st.subheader("B. Clean & Format Data")
        if st.button("Apply Basic Cleaning (Title Case, Strip)"):
            st.session_state.df_cleaned = basic_clean(st.session_state.df_cleaned)
            log_action("Data cleaning applied.")
            st.success("Data cleaning complete.")
            st.text("Thank you for using this APP_Eridu.") 
            st.rerun()

        st.header("3. Deduplication Controls")
        current_df = st.session_state.df_cleaned
        cols = current_df.columns.tolist() 

        # Exact Match Section
        st.subheader("Exact Match (Rule-Based)")
        # ... (Exact match logic remains the same) ...

        # Fuzzy Match Section (Comparing two columns)
        st.subheader("Interactive Fuzzy Match (Two Columns)")
        col_fuzzy1, col_fuzzy2 = st.columns(2)
        with col_fuzzy1:
            fuzzy_col_a = st.selectbox("Column A:", options=['None'] + cols, key='fuzzy_col_a')
        with col_fuzzy2:
            fuzzy_col_b = st.selectbox("Column B:", options=['None'] + cols, key='fuzzy_col_b')
        
        fuzzy_threshold = st.slider("Fuzzy match threshold (%)", min_value=70, max_value=100, value=85, step=1, key='fuzzy_thresh')

        if st.button("Detect Fuzzy Duplicates"):
            if fuzzy_col_a != 'None' and fuzzy_col_b != 'None':
                # Reset previous fuzzy state
                st.session_state.fuzzy_duplicates = []
                st.session_state.fuzzy_action = None
                perform_fuzzy_detection(st.session_state.df_cleaned, fuzzy_col_a, fuzzy_col_b, fuzzy_threshold)
                st.rerun()
            else:
                st.warning("Please select two columns for fuzzy matching.")

        # --- Interactive Fuzzy Management Area (Pop-up functionally) ---
        if st.session_state.fuzzy_duplicates:
            st.markdown("---")
            st.subheader("Manage Fuzzy Results")

            # Create a dataframe summary of duplicates for viewing/download
            # We filter the full dataset to show all rows involved in a match
            all_matched_indices = set(idx for match in st.session_state.fuzzy_duplicates for indices in [match['indices_col1'], match['indices_col2']] for idx in indices)
            df_fuzzy_duplicates_full = current_df.loc[list(all_matched_indices)].copy()
            # Sort for clarity
            sort_col = fuzzy_col_a if fuzzy_col_a != 'None' else cols[0]
            df_fuzzy_duplicates_full = df_fuzzy_duplicates_full.sort_values(by=sort_col)

            st.dataframe(df_fuzzy_duplicates_full)
            st.info(f"Showing {len(df_fuzzy_duplicates_full)} rows involved in the fuzzy match analysis.")

            col_dup_dl, col_dup_del, col_dup_ignore = st.columns(3)

            with col_dup_dl:
                dup_bytes, dup_type = df_to_bytes(df_fuzzy_duplicates_full, sheet_name="FuzzyDuplicates")
                st.download_button(
                    label="Download Duplicates",
                    data=dup_bytes,
                    file_name=f"detected_fuzzy_duplicates_{now_ts().replace(' ', '_').replace(':', '-')}.{'xlsx' if dup_type == 'excel' else 'csv'}",
                    mime=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if dup_type == 'excel' else 'csv'}"
                )
            
            with col_dup_ignore:
                if st.button("Ignore / Clear Results"):
                    st.session_state.fuzzy_duplicates = []
                    st.session_state.fuzzy_action = None
                    st.success("Fuzzy results cleared from view.")
                    st.rerun()

            with col_dup_del:
                # Need confirmation for deletion, setting session state flag
                if st.button("Delete Fuzzy Duplicates", type='primary'):
                   st.session_state.fuzzy_action = 'confirm_delete'
                   st.rerun()
            
            # --- Confirmation Pop-up Logic for Delete ---
            if st.session_state.fuzzy_action == 'confirm_delete':
                st.warning("Are you sure you want to delete rows involved in fuzzy matches? This keeps only the first unique instance based on Column A values.")
                col_c_del, col_c_cancel = st.columns(2)

                with col_c_del:
                    if st.button("Confirm Delete", use_container_width=True):
                        # Logic to perform actual deletion: keep only first match based on Column A value
                        cols_to_use = [fuzzy_col_a, fuzzy_col_b]
                        st.session_state.df_cleaned = current_df.drop_duplicates(subset=cols_to_use, keep='first').reset_index(drop=True)
                        deleted_count = len(current_df) - len(st.session_state.df_cleaned)
                        
                        log_action(f"Deleted {deleted_count} fuzzy duplicates based on {fuzzy_col_a} and {fuzzy_col_b}.")
                        st.success(f"Deleted {deleted_count} fuzzy duplicates. Total rows remaining: {len(st.session_state.df_cleaned)}")
                        st.text("Thank you for using this APP_Eridu.")
                        st.session_state.fuzzy_duplicates = []
                        st.session_state.fuzzy_action = None
                        st.rerun()
                with col_c_cancel:
                    if st.button("Cancel Delete", use_container_width=True):
                        st.session_state.fuzzy_action = None
                        st.rerun()
                st.stop() # Stop further execution while waiting for confirmation

# --- MAIN CONTENT AREA ---
# ... (Main content area logic for tabs 1, 2, and 3 remains the same) ...

if st.session_state.df_cleaned is None:
    st.info("Upload a file in the sidebar to begin data cleaning and analysis.")
else:
    tab1, tab2, tab3 = st.tabs(["📊 Data Summaries & Visualization", "📄 View & Download Data", "🕒 Audit Log"])

    with tab1:
        st.header("Data Summary & Custom Visuals")
        # ... (Tab 1 content remains the same) ...

    with tab2:
        st.header("Current Cleaned Dataset")
        st.dataframe(st.session_state.df_cleaned, use_container_width=True)
        st.markdown(f"**Total Rows in current view:** {len(st.session_state.df_cleaned)}")
        
        st.subheader("Download Final Cleaned Data")
        # ... (Download logic remains the same) ...

    with tab3:
        st.header("Activity Audit Log")
        st.json(st.session_state.audit_log)

# Sticky footer for developer credit
st.markdown('<div class="developer">Built by Eridu Moses</div>', unsafe_allow_html=True)
