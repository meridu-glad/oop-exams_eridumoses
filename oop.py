# import libraries
import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import re
from collections import defaultdict
from datetime import datetime
import plotly.express as px # Added Plotly import

# --- Helpers (pure Python & Streamlit session management) ---

def now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_action(msg):
    if st.session_state.get("audit_log") is None:
        st.session_state.audit_log = []
    st.session_state.audit_log.insert(0, f"{now_ts()} — {msg}")

def df_to_bytes(df):
    """Converts a DataFrame to an in-memory Excel buffer for download."""
    try:
        import openpyxl  # noqa: F401
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="data")
        buf.seek(0)
        return buf.getvalue(), "excel"
    except ImportError:
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
    """Detects groups of exact duplicates based on key columns."""
    if not key_cols:
        return []
    grouped = df.groupby(key_cols, dropna=False).groups
    groups = [list(indices) for indices in grouped.values() if len(indices) > 1]
    return groups

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


def detect_phonetic_duplicates(df, name_col, district_col=None):
    """
    Groups potential duplicates using Soundex codes.
    Returns a dictionary where keys are a composite key (e.g., soundex + district)
    and values are lists of DataFrame indices.
    """
    if not name_col:
        return {}

    # Calculate soundex codes for the name column
    df['soundex_temp'] = df[name_col].apply(soundex_code)
    
    # Create a composite key for grouping
    if district_col and district_col in df.columns:
        df['group_key_temp'] = df['soundex_temp'] + "_" + df[district_col].astype(str).fillna("NA")
    else:
        df['group_key_temp'] = df['soundex_temp']
    
    # Group by the composite key and find indices of groups > 1
    grouped_indices = defaultdict(list)
    for idx, key in enumerate(df['group_key_temp']):
        if key and key != "NA":
            grouped_indices[key].append(idx)
    
    # Filter to keep only actual duplicate groups
    duplicate_groups = {k: v for k, v in grouped_indices.items() if len(v) > 1}
    
    # Clean up temp columns (important for not cluttering the main DF)
    df.drop(columns=['soundex_temp', 'group_key_temp'], inplace=True)

    return duplicate_groups


# --- Streamlit App UI and Logic ---

# Page configuration & styling (kept as provided by user)
st.set_page_config(page_title="PLAYMATTERS DATABASE APP", layout="wide", initial_sidebar_state="auto")
st.markdown("""... your CSS styles ...""", unsafe_allow_html=True)
st.markdown('<div class="title">PLAYMATTERS DATABASE APP</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Deduplication, Cleaning, and Data Summarization</div>', unsafe_allow_html=True)


# Initialize session state variables
if 'df_cleaned' not in st.session_state: st.session_state.df_cleaned = None
if 'df_original' not in st.session_state: st.session_state.df_original = None
if 'duplicate_groups' not in st.session_state: st.session_state.duplicate_groups = {}
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
    st.sidebar.header("2. Phonetic Deduplication")
    
    current_df = st.session_state.df_cleaned
    cols = current_df.columns.tolist()
    
    name_col = st.sidebar.selectbox("Select 'Name' column for phonetic matching:", options=cols)
    district_col = st.sidebar.selectbox("Select 'District' (optional blocker):", options=[None] + cols)

    if st.sidebar.button("Detect Phonetic Duplicates"):
        if name_col:
            with st.spinner(f"Detecting phonetic matches using {name_col}..."):
                st.session_state.duplicate_groups = detect_phonetic_duplicates(current_df, name_col, district_col)
                count = sum(len(indices) - 1 for indices in st.session_state.duplicate_groups.values())
                log_action(f"Found {count} potential phonetic duplicates across {len(st.session_state.duplicate_groups)} groups.")
                st.sidebar.info(f"Found {count} potential phonetic duplicate rows.")
        else:
            st.sidebar.warning("Please select a name column.")

    if st.session_state.duplicate_groups:
        st.sidebar.subheader("Manage Duplicates")
        
        # Action 1: Download Duplicates
        all_duplicate_indices = [idx for indices in st.session_state.duplicate_groups.values() for idx in indices]
        duplicates_df = st.session_state.df_cleaned.loc[all_duplicate_indices].sort_index()
        
        dl_bytes, dl_type = df_to_bytes(duplicates_df)
        st.sidebar.download_button(
            label=f"Download {len(duplicates_df)} Duplicates for Review",
            data=dl_bytes,
            file_name=f"duplicates_for_review_{now_ts().replace(' ', '_')}.{'xlsx' if dl_type == 'excel' else 'csv'}",
            mime=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if dl_type == 'excel' else 'csv'}"
        )
        
        # Action 2: Delete Duplicates (keeping only one record per group)
        if st.sidebar.button("Delete ALL Duplicates (Keep 1st Instance Only)", help="This action removes all but the first record in each identified group."):
            # Logic: Identify which indices to KEEP (the first one of each group)
            indices_to_keep = [indices[0] for indices in st.session_state.duplicate_groups.values()]
            # Keep only unique indices if some records were in multiple groups (unlikely here)
            indices_to_keep_set = set(indices_to_keep) 
            
            # Filter the main dataframe
            st.session_state.df_cleaned = st.session_state.df_cleaned.loc[indices_to_keep_set].copy()
            st.session_state.df_cleaned = st.session_state.df_cleaned.reset_index(drop=True) # Reset index after dropping
            
            log_action(f"Deleted duplicates. New row count: {len(st.session_state.df_cleaned)}")
            st.sidebar.success(f"Duplicates removed. Total rows remaining: {len(st.session_state.df_cleaned)}")
            st.session_state.duplicate_groups = {} # Clear the duplicate list after action
            st.rerun() # Rerun to update the main view instantly


# --- MAIN CONTENT AREA: Data View & Summaries ---

if st.session_state.df_cleaned is None:
    st.info("Upload a file in the sidebar to begin data cleaning and analysis.")
else:
    st.header("Cleaned Data Overview")
    
    # Ensure standard column names exist for summary (e.g., 'Sex' and 'District')
    # If your data uses different names, you'll need a way to map them.
    df_display = st.session_state.df_cleaned
    
    if all(col in df_display.columns for col in ['Sex', 'District']):
        st.subheader("Data Summarization & Visuals")
        
        col1, col2 = st.columns([1, 2])

        with col1:
            st.markdown("##### Sex Distribution (Pie Chart)")
            # Calculate counts for the pie chart
            sex_counts = df_display['Sex'].value_counts().reset_index()
            sex_counts.columns = ['Sex', 'Count']
            
            fig_pie = px.pie(sex_counts, values='Count', names='Sex', title='Gender Distribution')
            st.plotly_chart(fig_pie, use_container_width=True)

        with col2:
            st.markdown("##### Sex count per District (Table)")
            # Pivot table for Sex per District
            sex_district_table = pd.crosstab(df_display['District'], df_display['Sex'])
            st.dataframe(sex_district_table)
            
    else:
        st.warning("Cannot generate data summaries. Please ensure your data has 'Sex' and 'District' columns (case sensitive) after basic cleaning.")

    # Always show the main cleaned dataframe below the charts
    st.subheader("Current Cleaned Dataset")
    st.dataframe(df_display, use_container_width=True)
    st.markdown(f"**Total Rows in current view:** {len(df_display)}")

    # Audit Log View (as a separate view option)
    with st.expander("View Audit Log"):
        st.code("\n".join(st.session_state.audit_log))


# Sticky footer for developer credit
st.markdown('<div class="developer">Developed by ERIDU MOSES</div>', unsafe_allow_html=True)
