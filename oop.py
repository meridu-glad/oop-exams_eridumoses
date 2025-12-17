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

def perform_fuzzy_dedupe(df, col_name, threshold=85):
    st.info(f"Starting fuzzy matching on '{col_name}' with threshold {threshold}%...")
    unique_items = df[col_name].dropna().unique().tolist()
    canonical_map = {}
    progress_bar = st.progress(0)
    for i, item in enumerate(unique_items):
        matches = process.extract(item, unique_items, scorer=fuzz.token_sort_ratio, limit=None)
        similar_items = [match for match in matches if match >= threshold and match != item]
        current_canonical = canonical_map.get(item, item)
        for similar in similar_items:
            if similar not in canonical_map:
                canonical_map[similar] = current_canonical
        canonical_map[item] = current_canonical
        progress_bar.progress((i + 1) / len(unique_items))
    progress_bar.empty()
    st.success("Fuzzy matching complete.")
    df['Fuzzy_Group'] = df[col_name].map(canonical_map)
    df_deduped = df.drop_duplicates(subset=['Fuzzy_Group'], keep='first').drop(columns=['Fuzzy_Group'])
    deleted_count = len(df) - len(df_deduped)
    log_action(f"Fuzzy deduplication deleted {deleted_count} rows from column '{col_name}'.")
    return df_deduped

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
# New state trackers for UI flow control
if 'data_is_cleaned' not in st.session_state: st.session_state.data_is_cleaned = False
if 'data_is_processed' not in st.session_state: st.session_state.data_is_processed = False


# --- SIDEBAR: Upload and Processing Controls ---
with st.sidebar:
    st.header("1. Upload Data")
    uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=['csv', 'xlsx', 'xls'])

    if uploaded_file is not None and st.session_state.df_original is None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            st.session_state.df_original = df
            st.session_state.df_cleaned = df.copy() # Load raw data into cleaned slot initially
            st.session_state.data_is_cleaned = False
            st.session_state.data_is_processed = False
            log_action(f"File uploaded: {uploaded_file.name} ({len(df)} rows).")
            st.rerun() 
        except Exception as e:
            st.error(f"Error reading file: {e}")
    
    if st.session_state.df_original is not None:
        st.header("2. Pre-Processing Actions")
        
        # Action Button: Clean Dataset
        if st.button("Clean Data (Title Case, Strip Whitespace)"):
            st.session_state.df_cleaned = basic_clean(st.session_state.df_cleaned)
            st.session_state.data_is_cleaned = True
            log_action("Data cleaning applied.")
            st.success("Data cleaning complete.")
            st.rerun()
            
        # Action Button: Deal with Missing Data
        st.subheader("Handle Missing Data")
        missing_option = st.selectbox("Choose how to handle NaNs:", ["Do nothing", "Drop rows with ANY missing data", "Fill NaNs with custom value"])
        
        if missing_option == "Fill NaNs with custom value":
            fill_value = st.text_input("Value to fill NaNs with:", value="Missing")
        
        if st.button("Apply Missing Data Action"):
            if missing_option == "Drop rows with ANY missing data":
                st.session_state.df_cleaned = st.session_state.df_cleaned.dropna().reset_index(drop=True)
                log_action("Dropped rows with missing data.")
                st.success(f"Dropped rows. New count: {len(st.session_state.df_cleaned)}")
            elif missing_option == "Fill NaNs with custom value" and fill_value:
                st.session_state.df_cleaned = st.session_state.df_cleaned.fillna(fill_value)
                log_action(f"Filled NaNs with '{fill_value}'.")
                st.success(f"Filled missing data.")
            st.rerun()
            
        st.header("3. Deduplication Controls")
        # Rest of deduplication controls (Exact/Fuzzy) remain here... 
        # (omitted for brevity in this response, assumed functional from previous code)


# --- MAIN CONTENT AREA ---

if st.session_state.df_cleaned is None:
    st.info("Upload a file in the sidebar to begin data cleaning and analysis.")
else:
    tab1, tab2, tab3 = st.tabs(["📊 Data Summaries & Visualization", "📄 View & Download Data", "🕒 Audit Log"])

    with tab1:
        st.header("Data Summary & Custom Visuals")
        
        col_summary_1, col_summary_2 = st.columns(2)
        col_summary_1.metric("Original Rows", len(st.session_state.df_original))
        col_summary_2.metric("Current Rows (Post-Processing)", len(st.session_state.df_cleaned))

        st.subheader("Conditional Visualizations")
        cols_for_viz = [c for c in st.session_state.df_cleaned.columns if pd.api.types.is_string_dtype(st.session_state.df_cleaned[c])]
        
        if cols_for_viz:
            viz_type = st.radio("Choose visualization type:", ("Pie Chart (Single Variable)", "Bar Chart (X/Y Variables)"))

            if viz_type == "Pie Chart (Single Variable)":
                selected_col_pie = st.selectbox("Select a column for the Pie Chart:", cols_for_viz, key='viz_pie_col')
                if selected_col_pie:
                    counts_df = st.session_state.df_cleaned[selected_col_pie].value_counts().reset_index()
                    counts_df.columns = [selected_col_pie, 'Count']
                    fig_pie = px.pie(counts_df, values='Count', names=selected_col_pie, title=f"Distribution of {selected_col_pie}")
                    st.plotly_chart(fig_pie, use_container_width=True)

            elif viz_type == "Bar Chart (X/Y Variables)":
                col_x, col_y = st.columns(2)
                with col_x:
                    selected_col_x = st.selectbox("Select X-axis column:", cols_for_viz, key='viz_bar_x')
                with col_y:
                    # Allow numerical columns for Y-axis too
                    numeric_cols = [c for c in st.session_state.df_cleaned.columns if pd.api.types.is_numeric_dtype(st.session_state.df_cleaned[c])]
                    selected_col_y = st.selectbox("Select Y-axis column:", numeric_cols, key='viz_bar_y')
                
                if selected_col_x and selected_col_y:
                    fig_bar = px.bar(st.session_state.df_cleaned, x=selected_col_x, y=selected_col_y, title=f"{selected_col_y} by {selected_col_x}")
                    st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No suitable string columns found for visualization.")

    with tab2:
        st.header("Current Cleaned Dataset")
        st.dataframe(st.session_state.df_cleaned, use_container_width=True)
        st.markdown(f"**Total Rows in current view:** {len(st.session_state.df_cleaned)}")
        
        st.subheader("Download Final Cleaned Data")
        
        # The download button is now always active if data exists in df_cleaned state
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
