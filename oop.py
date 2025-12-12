# import libraries
import streamlit as st
import pandas as pd
import numpy as np
import io
import time
# Removed rapidfuzz/fuzz/difflib imports as we focus on exact matching now
from collections import defaultdict, Counter
from datetime import datetime
import math
import plotly.express as px # Added Plotly import

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
    for c in df.select_dtypes(include=["object", "string"]).columns:
        df[c] = df[c].astype(str).replace({"nan": np.nan, "None": np.nan})
        df[c] = df[c].where(df[c].notnull(), np.nan)
        df[c] = df[c].apply(lambda x: " ".join(str(x).strip().title().split()) if pd.notna(x) else x)
    return df

def detect_exact_groups(df, key_cols):
    """Detects groups of exact duplicates based on key columns, returns indices of duplicates."""
    if not key_cols:
        return {}
    
    # Grouped is a dictionary where keys are the column values combo, and values are the indices
    grouped = df.groupby(key_cols, dropna=False).groups
    
    # Filter for groups that have more than one entry
    duplicate_groups = {k: list(indices) for k, indices in grouped.items() if len(indices) > 1}
    return duplicate_groups


# --- Streamlit App UI and Logic ---

# Page configuration & styling (kept as provided by user)
st.set_page_config(
    page_title="PLAYMATTERS DATABASE APP",
    layout="wide",
    initial_sidebar_state="auto"
)

# Inject custom CSS for styling (assuming this is copied from your original code)
st.markdown("""
<style>
/* ... your CSS here ... */
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
st.markdown('<div class="subtitle">Exact Deduplication, Cleaning, and Data Summarization for Attendance Records</div>', unsafe_allow_html=True)


# Initialize session state variables
if 'df_cleaned' not in st.session_state: st.session_state.df_cleaned = None
if 'df_original' not in st.session_state: st.session_state.df_original = None
if 'exact_groups' not in st.session_state: st.session_state.exact_groups = {}
if 'audit_log' not in st.session_state: st.session_state.audit_log = []


st.sidebar.header("1. Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload your Excel or CSV file", type=['csv', 'xlsx', 'xls'])

if uploaded_file is not None and st.session_state.df_original is None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        st.session_state.df_original = df
        st.session_state.df_cleaned = basic_clean(df.copy()) # Auto-clean on upload
        log_action(f"File uploaded & auto-cleaned: {uploaded_file.name} ({len(df)} rows).")
        st.rerun() # Rerun to show new sidebar options

    except Exception as e:
        st.error(f"Error reading file: {e}")

# --- SIDEBAR: Cleaning Controls ---
if st.session_state.df_cleaned is not None:
    st.sidebar.header("2. Exact Deduplication (Rule-Based)")
    
    current_df = st.session_state.df_cleaned
    cols = current_df.columns.tolist()
    exact_cols_to_check = st.sidebar.multiselect(
        "Select columns that *must* match exactly:", 
        options=cols, 
        default=[]
    )

    if st.sidebar.button("Detect EXACT Duplicates"):
        if exact_cols_to_check:
            with st.spinner("Detecting exact matches..."):
                st.session_state.exact_groups = detect_exact_groups(current_df, exact_cols_to_check)
                count = sum(len(indices) - 1 for indices in st.session_state.exact_groups.values())
                log_action(f"Found {count} exact duplicates across {len(st.session_state.exact_groups)} groups.")
                st.sidebar.info(f"Found {count} exact duplicate rows.")
        else:
            st.sidebar.warning("Please select columns for exact match detection.")

    if st.session_state.exact_groups:
        st.sidebar.subheader("Manage Duplicates")
        
        # Action 1: Download Duplicates for Manual Review
        all_duplicate_indices = [idx for indices in st.session_state.exact_groups.values() for idx in indices]
        duplicates_df = st.session_state.df_cleaned.loc[all_duplicate_indices].sort_index()
        
        dl_bytes, dl_type = df_to_bytes(duplicates_df, sheet_name="Duplicates")
        st.sidebar.download_button(
            label=f"Download {len(duplicates_df)} Duplicates",
            data=dl_bytes,
            file_name=f"duplicates_for_review_{now_ts().replace(' ', '_')}.{'xlsx' if dl_type == 'excel' else 'csv'}",
            mime=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if dl_type == 'excel' else 'csv'}"
        )
        
        # Action 2: Delete Duplicates (keeping only one record per group)
        if st.sidebar.button("Delete ALL Duplicates (Keep 1st Instance Only)"):
            
            # --- Corrected Logic for Indexing (Fixes the TypeError) ---
            # Extract only the first index of each group to keep
            indices_to_keep_list = [indices[0] for indices in st.session_state.exact_groups.values()]
            
            # Filter the main dataframe using a list indexer
            st.session_state.df_cleaned = st.session_state.df_cleaned.loc[indices_to_keep_list].copy()
            st.session_state.df_cleaned = st.session_state.df_cleaned.reset_index(drop=True)
            
            log_action(f"Deleted duplicates. New row count: {len(st.session_state.df_cleaned)}")
            st.sidebar.success(f"Duplicates removed. Total rows remaining: {len(st.session_state.df_cleaned)}")
            st.session_state.exact_groups = {} # Clear the duplicate list after action
            st.rerun()


# --- MAIN CONTENT AREA: Data View & Summaries ---

if st.session_state.df_cleaned is None:
    st.info("Upload a file in the sidebar to begin data cleaning and analysis.")
else:
    st.header("Cleaned Data Overview")
    
    df_display = st.session_state.df_cleaned
    
    # --- Data Summarization & Visuals ---
    # We check if Sex and District columns are present for the visuals requested by the user
    if all(col in df_display.columns for col in ['Sex', 'District']):
        st.subheader("Data Summarization & Visuals")
        
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("##### Sex Distribution (Pie Chart)")
            sex_counts = df_display['Sex'].value_counts().reset_index()
            sex_counts.columns = ['Sex', 'Count']
            
            fig_pie = px.pie(sex_counts, values='Count', names='Sex', title='Gender Distribution')
            st.plotly_chart(fig_pie, use_container_width=True)

        with col2:
            st.markdown("##### Sex count per District (Table)")
            sex_district_table = pd.crosstab(df_display['District'], df_display['Sex'])
            st.dataframe(sex_district_table)
            
    else:
        st.warning("Cannot generate data summaries. Please ensure your data has 'Sex' and 'District' columns (case sensitive) after basic cleaning.")

    # --- Main Dataframe Display and Download Option ---
    st.subheader("Current Cleaned Dataset")
    st.dataframe(df_display, use_container_width=True)
    st.markdown(f"**Total Rows in current view:** {len(df_display)}")
    
    # Final option for downloading clean data
    clean_bytes, clean_type = df_to_bytes(df_display, sheet_name="CleanedData")
    st.download_button(
        label=f"Download Final Cleaned Data as {'Excel' if clean_type == 'excel' else 'CSV'}",
        data=clean_bytes,
        file_name=f"final_cleaned_data_{now_ts().replace(' ', '_').replace(':', '-')}.{'xlsx' if clean_type == 'excel' else 'csv'}",
        mime=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if clean_type == 'excel' else 'csv'}"
    )


# Sticky footer for developer credit
st.markdown('<div class="developer">Built by Eridu Moses</div>', unsafe_allow_html=True)
