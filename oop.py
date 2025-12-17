# import libraries
import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import math
import plotly.express as px 
# Removed thefuzz library import as it wasn't used in the provided code snippet
# from thefuzz import process, fuzz 

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
        # Ensured openpyxl is imported locally if not in global scope
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
    # Store the actual duplicate rows data in session state for downloading later
    duplicate_indices = [idx for indices in grouped.values() if len(indices) > 1 for idx in indices]
    
    # Store only the indices of the duplicate rows, not the whole dataframe in memory here
    duplicate_groups = {k: list(indices) for k, indices in grouped.items() if len(indices) > 1}
    
    return duplicate_groups

# Note: The fuzzy dedupe function was not used in the final code provided, but kept here for completeness
def perform_fuzzy_dedupe(df, col_name, threshold=85):
    from thefuzz import process, fuzz # Import locally if needed
    st.info(f"Starting fuzzy matching on '{col_name}' with threshold {threshold}%...")
    unique_items = df[col_name].dropna().unique().tolist()
    canonical_map = {}
    progress_bar = st.progress(0)
    for i, item in enumerate(unique_items):
        matches = process.extract(item, unique_items, scorer=fuzz.token_sort_ratio, limit=None)
        similar_items = [match for match in matches if match[1] >= threshold and match[0] != item] # Check match score element
        current_canonical = canonical_map.get(item, item)
        for similar_text, score in similar_items:
            if similar_text not in canonical_map:
                canonical_map[similar_text] = current_canonical
        canonical_map[item] = current_canonical
        progress_bar.progress((i + 1) / len(unique_items))
    progress_bar.empty()
    st.success("Fuzzy matching complete.")
    df['Fuzzy_Group'] = df[col_name].map(canonical_map)
    df_deduped = df.drop_duplicates(subset=['Fuzzy_Group'], keep='first').drop(columns=['Fuzzy_Group'])
    deleted_count = len(df) - len(df_deduped)
    log_action(f"Fuzzy deduplication deleted {deleted_count} rows from column '{col_name}'.")
    st.text("Thank you for using this APP_Eridu.") # Appreciation note
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

# Function to clear session state and rerun (Refresh Button Logic)
def refresh_app():
    st.session_state.clear()
    st.rerun()

# --- SIDEBAR: Upload and Processing Controls ---
with st.sidebar:
    st.header("1. Upload Data")
    uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=['csv', 'xlsx', 'xls'])

    if st.button("Refresh Application (Clear All Data)", type="secondary"): # Refresh button
        refresh_app()

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
        # Display missing data report first
        missing_data = st.session_state.df_cleaned.isnull().sum()
        missing_data = missing_data[missing_data > 0]
        if not missing_data.empty:
            st.warning("Missing values detected in current data:")
            st.dataframe(missing_data.rename("Count"))
            
            # Then provide options to handle it
            missing_option = st.selectbox("Choose how to handle NaNs:", ["Do nothing", "Drop rows with ANY missing data", "Fill NaNs with custom value"])
            
            fill_value = None
            if missing_option == "Fill NaNs with custom value":
                fill_value = st.text_input("Value to fill NaNs with:", value="Missing")
            
            if st.button("Apply Missing Data Action"):
                if missing_option == "Drop rows with ANY missing data":
                    st.session_state.df_cleaned = st.session_state.df_cleaned.dropna().reset_index(drop=True)
                    log_action("Dropped rows with missing data.")
                    st.success(f"Dropped rows. New count: {len(st.session_state.df_cleaned)}")
                elif missing_option == "Fill NaNs with custom value" and fill_value is not None:
                    st.session_state.df_cleaned = st.session_state.df_cleaned.fillna(fill_value)
                    log_action(f"Filled NaNs with '{fill_value}'.")
                    st.success(f"Filled missing data.")
                st.text("Thank you for using this APP_Eridu.") # Appreciation note
                st.rerun()
        else:
            st.success("No missing data detected in current view.")
        
        st.subheader("B. Clean & Format Data")
        if st.button("Apply Basic Cleaning (Title Case, Strip)"):
            st.session_state.df_cleaned = basic_clean(st.session_state.df_cleaned)
            log_action("Data cleaning applied.")
            st.success("Data cleaning complete.")
            st.text("Thank you for using this APP_Eridu.") # Appreciation note
            st.rerun()

        # --- Deduplication Controls (Now correctly nested and active) ---
        st.header("3. Deduplication Controls")
        current_df = st.session_state.df_cleaned
        cols = current_df.columns.tolist() 

        # Exact Match Section
        st.subheader("Exact Match (Rule-Based)")
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
                    if count > 0:
                        st.info(f"Found {count} exact duplicate rows. Options to delete or download below.")
                    else:
                        st.success("No exact duplicates found.")
            else:
                st.warning("Please select columns.")

        if st.session_state.exact_groups:
            # Prepare duplicate data for download
            duplicate_indices_list = [idx for indices in st.session_state.exact_groups.values() for idx in indices]
            # Filter the original DF to show *all* instances of the duplicates grouped together
            df_duplicates = current_df.loc[duplicate_indices_list].sort_values(by=exact_cols_to_check)
            
            dup_bytes, dup_type = df_to_bytes(df_duplicates, sheet_name="ExactDuplicates")
            
            st.markdown("---")
            st.subheader("Manage Detected Duplicates")
            
            # Download Button for Duplicates
            st.download_button(
                label=f"Download {len(df_duplicates)} Duplicate Rows ({'Excel' if dup_type == 'excel' else 'CSV'})",
                data=dup_bytes,
                file_name=f"detected_duplicates_{now_ts().replace(' ', '_').replace(':', '-')}.{'xlsx' if dup_type == 'excel' else 'csv'}",
                mime=f"application/{'vnd.openxmlformats-officedocument.spreadsheetml.sheet' if dup_type == 'excel' else 'csv'}"
            )

            # Delete Button for Duplicates
            if st.button("Delete ALL EXACT Duplicates (Keep 1st Instance)"):
                # Logic refined to drop subsequent duplicates
                st.session_state.df_cleaned = current_df.drop_duplicates(subset=exact_cols_to_check, keep='first').reset_index(drop=True)
                
                deleted_count = len(current_df) - len(st.session_state.df_cleaned)
                log_action(f"Deleted {deleted_count} exact duplicates. New row count: {len(st.session_state.df_cleaned)}")
                st.success(f"Duplicates removed. Total rows remaining: {len(st.session_state.df_cleaned)}")
                st.text("Thank you for using this APP_Eridu.") # Appreciation note
                st.session_state.exact_groups = {} 
                st.rerun() 

        # Fuzzy Match Section (Optional - kept for original functionality)
        st.subheader("Optional Fuzzy Match")
        fuzzy_col_to_check = st.selectbox("Select a single column for fuzzy deduplication:", options=['None'] + cols)
        fuzzy_threshold = st.slider("Fuzzy match threshold (%)", min_value=70, max_value=100, value=85, step=1)

        if st.button("Perform Fuzzy Deduplication"):
            if fuzzy_col_to_check != 'None':
                # Note: this function requires thefuzz library which needs installation
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
        col_summary_2.metric("Current Rows (Post-Processing)", len(st.session_state.df_cleaned))

        st.subheader("Conditional Visualizations")
        
        # Determine categorical/string columns dynamically
        cols_for_viz_cat = [c for c in st.session_state.df_cleaned.columns if pd.api.types.is_string_dtype(st.session_state.df_cleaned[c])]
        # Determine numerical columns dynamically
        cols_for_viz_num = [c for c in st.session_state.df_cleaned.columns if pd.api.types.is_numeric_dtype(st.session_state.df_cleaned[c])]

        if cols_for_viz_cat:
            viz_type = st.radio("Choose visualization type:", ("Pie Chart (Single Variable)", "Bar Chart (X/Y Variables)"))

            if viz_type == "Pie Chart (Single Variable)":
                selected_col_pie = st.selectbox("Select a column for the Pie Chart:", cols_for_viz_cat, key='viz_pie_col')
                if selected_col_pie:
                    counts_df = st.session_state.df_cleaned[selected_col_pie].value_counts().reset_index()
                    counts_df.columns = [selected_col_pie, 'Count']
                    fig_pie = px.pie(counts_df, values='Count', names=selected_col_pie, title=f"Distribution of {selected_col_pie}")
                    st.plotly_chart(fig_pie, use_container_width=True)

            elif viz_type == "Bar Chart (X/Y Variables)":
                col_x, col_y = st.columns(2)
                with col_x:
                    selected_col_x = st.selectbox("Select X-axis column (Categorical):", cols_for_viz_cat, key='viz_bar_x')
                with col_y:
                    # FIX APPLIED HERE: The selectbox options are dynamically generated based on numeric columns only
                    # and the default value is handled if the list is empty.
                    if cols_for_viz_num:
                        selected_col_y = st.selectbox(
                            "Select Y-axis column (Numeric):", 
                            cols_for_viz_num, 
                            key='viz_bar_y'
                        )
                    else:
                        selected_col_y = None
                        st.info("No numeric columns available for Y-axis.")
                
                if selected_col_x and selected_col_y:
                    fig_bar = px.bar(st.session_state.df_cleaned, x=selected_col_x, y=selected_col_y, title=f"{selected_col_y} by {selected_col_x}")
                    st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No suitable categorical columns found for visualization.")

    with tab2:
        st.header("Current Cleaned Dataset")
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
