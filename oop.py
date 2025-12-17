# import libraries
import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import math
import plotly.express as px 

# Import fuzzy matching library (requires 'pip install thefuzz python-Levenshtein')
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
        import openpyxl  # noqa: F401 (ensure the library is available)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        buf.seek(0)
        return buf.getvalue(), "excel"
    except ImportError:
        # Fallback to CSV if openpyxl isn't installed
        return df.to_csv(index=False).encode("utf-8"), "csv"

def basic_clean(df):
    """Applies basic cleaning: title case, strip whitespace, handle NaNs."""
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == 'object' or pd.api.types.is_string_dtype(df[c]):
            # Coerce to string safely before applying string methods
            df[c] = df[c].astype(str).replace({"nan": np.nan, "None": np.nan})
            df[c] = df[c].where(df[c].notnull(), np.nan)
            # Use vectorized operations where possible for speed
            df[c] = df[c].str.strip().str.title().str.split().str.join(" ")
    return df

def detect_exact_groups(df, key_cols):
    """Detects groups of exact duplicates based on key columns, returns indices of duplicates."""
    if not key_cols:
        return {}
    grouped = df.groupby(key_cols, dropna=False).groups
    duplicate_groups = {k: list(indices) for k, indices in grouped.items() if len(indices) > 1}
    return duplicate_groups

def perform_fuzzy_dedupe(df, col_name, threshold=85):
    """
    Performs fuzzy deduplication on a single column. 
    Keeps the first instance of a matched group.
    """
    st.info(f"Starting fuzzy matching on '{col_name}' with threshold {threshold}%...")
    
    # Get unique items to match against
    unique_items = df[col_name].dropna().unique().tolist()
    
    # We will map similar strings to a canonical (first occurring) version
    canonical_map = {}
    
    # Use the process.extract function which is optimized for matching one list against itself/another
    progress_bar = st.progress(0)
    for i, item in enumerate(unique_items):
        # Find matches above the threshold
        matches = process.extract(item, unique_items, scorer=fuzz.token_sort_ratio, limit=None)
        
        # Filter matches above the threshold (excluding itself)
        similar_items = [match[0] for match in matches if match[1] >= threshold and match[0] != item]
        
        # Assign all similar items to the same canonical name (the first one encountered in the loop)
        current_canonical = canonical_map.get(item, item)
        for similar in similar_items:
            if similar not in canonical_map:
                canonical_map[similar] = current_canonical
        
        canonical_map[item] = current_canonical
        progress_bar.progress((i + 1) / len(unique_items))
        
    progress_bar.empty()
    st.success("Fuzzy matching complete.")

    # Apply the mapping back to a new column for review/grouping
    df['Fuzzy_Group'] = df[col_name].map(canonical_map)
    
    # Now use this new group ID to drop duplicates
    df_deduped = df.drop_duplicates(subset=['Fuzzy_Group'], keep='first').drop(columns=['Fuzzy_Group'])
    
    deleted_count = len(df) - len(df_deduped)
    log_action(f"Fuzzy deduplication deleted {deleted_count} rows from column '{col_name}'.")
    return df_deduped


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
.title { text-align: center; font-size: 36px; font-weight: 800; margin-bottom: 6px; color: #1E3A8A; font-family: 'Segoe UI', Tahoma, sans-serif; }
.subtitle { text-align: center; font-size: 16px; margin-top: 0px; color: #475569; font-family: 'Segoe UI', Tahoma, sans-serif; }
.developer { position: fixed; right: 14px; bottom: 10px; font-style: italic; color: #1E3A8A; font-size: 14px; }
.stButton>button, .stDownloadButton>button { background-color: #2563EB; color: white; font-weight: 700; border-radius: 8px; padding: 8px 16px; font-size: 14px; transition: background-color 0.3s ease; }
.stButton>button:hover, .stDownloadButton>button:hover { background-color: #1E40AF; }
</style>
""", unsafe_allow_html=True)

# Render title and subtitle
st.markdown('<div class="title">PLAYMATTERS DATABASE APP</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Deduplication, Cleaning, and Data Summarization for Attendance Records</div>', unsafe_allow_html=True)


# Initialize session state variables
if 'df_cleaned' not in st.session_state: st.session_state.df_cleaned = None
if 'df_original' not in st.session_state: st.session_state.df_original = None
if 'exact_groups' not in st.session_state: st.session_state.exact_groups = {}
if 'audit_log' not in st.session_state: st.session_state.audit_log = []


# --- SIDEBAR: Upload and Pre-Cleaning Check ---
with st.sidebar:
    st.header("1. Upload Data & Check Quality")
    uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=['csv', 'xlsx', 'xls'])

    if uploaded_file is not None and st.session_state.df_original is None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            st.session_state.df_original = df
            st.session_state.df_cleaned = basic_clean(df.copy()) # Auto-clean on upload
            log_action(f"File uploaded & auto-cleaned: {uploaded_file.name} ({len(df)} rows).")
            st.rerun() 

        except Exception as e:
            st.error(f"Error reading file: {e}")
    
    if st.session_state.df_original is not None:
        st.subheader("Missing Data Report (Pre-Clean)")
        # Check for missing data before any manual cleaning steps
        missing_data = st.session_state.df_original.isnull().sum()
        missing_data = missing_data[missing_data > 0]
        if not missing_data.empty:
            st.warning("Missing values detected in original data:")
            st.dataframe(missing_data.rename("Count"))
        else:
            st.success("No missing data detected in original file.")


# --- SIDEBAR: Cleaning Controls ---
if st.session_state.df_cleaned is not None:
    with st.sidebar:
        st.header("2. Deduplication Controls")
        current_df = st.session_state.df_cleaned
        cols = current_df.columns.tolist()
        
        # --- Exact Match Section ---
        st.subheader("A. Exact Match (Rule-Based)")
        exact_cols_to_check = st.multiselect(
            "Select columns that *must* match exactly:", 
            options=cols, 
            default=[]
        )

        if st.button("Detect EXACT Duplicates"):
            if exact_cols_to_check:
                with st.spinner("Detecting exact matches..."):
                    st.session_state.exact_groups = detect_exact_groups(current_df, exact_cols_to_check)
                    count = sum(len(indices) - 1 for indices in st.session_state.exact_groups.values())
                    log_action(f"Found {count} exact duplicates.")
                    st.info(f"Found {count} exact duplicate rows.")
            else:
                st.warning("Please select columns.")

        if st.session_state.exact_groups:
            # Action: Delete Duplicates (keeping only one record per group)
            if st.button("Delete ALL EXACT Duplicates (Keep 1st Instance)"):
                indices_to_keep_mask = ~current_df.index.isin([idx for indices in st.session_state.exact_groups.values() for idx in indices[1:]])
                st.session_state.df_cleaned = current_df[indices_to_keep_mask].reset_index(drop=True)
                log_action(f"Deleted exact duplicates. New row count: {len(st.session_state.df_cleaned)}")
                st.success(f"Duplicates removed. Total rows remaining: {len(st.session_state.df_cleaned)}")
                st.session_state.exact_groups = {} # Clear the duplicate list after action
                st.rerun() 

        # --- Fuzzy Match Section (Optional) ---
        st.subheader("B. Optional Fuzzy Match")
        fuzzy_col_to_check = st.selectbox("Select a single column for fuzzy deduplication:", options=['None'] + cols)
        fuzzy_threshold = st.slider("Fuzzy match threshold (%)", min_value=70, max_value=100, value=85, step=1)

        if st.button("Perform Fuzzy Deduplication"):
            if fuzzy_col_to_check != 'None':
                st.session_state.df_cleaned = perform_fuzzy_dedupe(st.session_state.df_cleaned, fuzzy_col_to_check, fuzzy_threshold)
                st.rerun()
            else:
                st.warning("Please select a column for fuzzy matching.")


# --- MAIN CONTENT AREA ---

if st.session_state.df_cleaned is None:
    st.info("Upload a file in the sidebar to begin data cleaning and analysis.")
else:
    tab1, tab2, tab3 = st.tabs(["📊 Data Summaries & Visualization", "📄 View & Download Data", "🕒 Audit Log"])

    with tab1:
        st.header("Data Summary & Custom Visuals")
        
        col_summary_1, col_summary_2 = st.columns(2)
        col_summary_1.metric("Original Rows", len(st.session_state.df_original))
        col_summary_2.metric("Cleaned Rows (Current)", len(st.session_state.df_cleaned))

        st.subheader("Conditional Visualizations")
        # The visualization options should be dynamic based on the *currently* cleaned data
        cols_for_viz = [c for c in st.session_state.df_cleaned.columns if pd.api.types.is_string_dtype(st.session_state.df_cleaned[c]) and st.session_state.df_cleaned[c].nunique() < 50]
        
        if cols_for_viz:
            selected_col_x = st.selectbox("Select a column for the X-axis (Categorical):", ['None'] + cols_for_viz, key='viz_x')
            
            if selected_col_x != 'None':
                st.markdown(f"##### Frequency of {selected_col_x}")
                counts_df = st.session_state.df_cleaned[selected_col_x].value_counts().reset_index()
                counts_df.columns = [selected_col_x, 'Count']
                
                fig_bar = px.bar(counts_df, x=selected_col_x, y='Count', title=f"Distribution of {selected_col_x}", color='Count')
                st.plotly_chart(fig_bar, use_container_width=True)

        else:
            st.info("No suitable categorical columns found (fewer than 50 unique string values) for visualization.")


    with tab2:
        st.header("View Current Cleaned Data")
        st.dataframe(st.session_state.df_cleaned, use_container_width=True)
        st.markdown(f"**Total Rows in current view:** {len(st.session_state.df_cleaned)}")
        
        st.subheader("Download Final Cleaned Data")
        clean_bytes, clean_type = df_to_bytes(st.session_state.df_cleaned, sheet_name="CleanedData")
        st.download_button(
            label=f"Download Final Cleaned Data as {'Excel' if clean_type == 'excel' else 'CSV'}",
            data=clean_bytes,
            file_name=f"final_cleaned_data_{now_ts().replace(' ', '_').replace(':', '-')}.{'xlsx' if clean_type == 'excel' else 'csv'}",
            mime=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if clean_type == 'excel' else 'csv'}"
        )

    with tab3:
        st.header("Activity Audit Log")
        st.json(st.session_state.audit_log)

# Sticky footer for developer credit
st.markdown('<div class="developer">Built by Eridu Moses</div>', unsafe_allow_html=True)
