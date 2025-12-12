# import libraries
import streamlit as st
import pandas as pd
import numpy as np
import io
import time
from rapidfuzz import fuzz
from collections import defaultdict, Counter
from datetime import datetime
import math

# Page config & styling


# Page configuration
st.set_page_config(
    page_title="PLAYMATTERS DATABASE APP",
    layout="wide",
    initial_sidebar_state="auto"
)

# Inject custom CSS for styling
st.markdown("""
<style>
/* Background */
body {
    background-color: #F8F9FA; /* Light gray for clean look */
}

/* Title */
.title {
    text-align: center;
    font-size: 36px;
    font-weight: 800;
    margin-bottom: 6px;
    color: #1E3A8A; /* Deep blue */
    font-family: 'Segoe UI', Tahoma, sans-serif;
}

/* Subtitle */
.subtitle {
    text-align: center;
    font-size: 16px;
    margin-top: 0px;
    color: #475569; /* Slate gray */
    font-family: 'Segoe UI', Tahoma, sans-serif;
}

/* Developer credit (sticky footer) */
.developer {
    position: fixed;
    right: 14px;
    bottom: 10px;
    font-style: italic;
    color: #1E3A8A;
    font-size: 14px;
}

/* Buttons */
.stButton>button, .stDownloadButton>button {
    background-color: #2563EB; /* Blue */
    color: white;
    font-weight: 700;
    border-radius: 8px;
    padding: 8px 16px;
    font-size: 14px;
    transition: background-color 0.3s ease;
}
.stButton>button:hover, .stDownloadButton>button:hover {
    background-color: #1E40AF; /* Darker blue on hover */
}

/* Progress label */
.progress-label {
    font-weight: 700;
    color: #1E293B; /* Dark slate */
}

/* Table styling */
table.data {
    border-collapse: collapse;
    width: 100%;
}
table.data td, th {
    border: 1px solid #ddd;
    padding: 8px;
}
</style>
""", unsafe_allow_html=True)

# Render title and subtitle
st.markdown('<div class="title">PLAYMATTERS DATABASE APP</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Exact Deduplication and Data Cleaning for Attendance Records</div>', unsafe_allow_html=True)


# Helpers

def now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_action(msg):
    if st.session_state.get("audit_log") is None:
        st.session_state.audit_log = []
    st.session_state.audit_log.insert(0, f"{now_ts()} — {msg}")

def df_to_bytes(df):
    try:
        import openpyxl  # noqa
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="cleaned")
        buf.seek(0)
        return buf.getvalue(), "excel"
    except Exception:
        return df.to_csv(index=False).encode("utf-8"), "csv"

def basic_clean(df):
    df = df.copy()
    for c in df.select_dtypes(include=["object"]).columns:
        df[c] = df[c].astype(str).replace({"nan": np.nan, "None": np.nan})
        df[c] = df[c].where(df[c].notnull(), np.nan)
        df[c] = df[c].apply(lambda x: " ".join(x.strip().title().split()) if isinstance(x, str) else x)
    return df

def detect_exact_groups(df, key_cols):
    if not key_cols:
        return []
    grouped = df.groupby(key_cols, dropna=False).indices
    groups = [list(v) for v in grouped.values() if len(v) > 1]
    return groups

def soundex(word: str) -> str:
    if not isinstance(word, str) or word.strip() == "":
        return ""
    w = word.upper()
    mapping = {
        "B":"1","F":"1","P":"1","V":"1",
        "C":"2","G":"2","J":"2","K":"2","Q":"2","S":"2","X":"2","Z":"2",
        "D":"3","T":"3",
        "L":"4",
        "M":"5","N":"5",
        "R":"6"
    }
    first = w[0]
    tail = w[1:]
    prev = mapping.get(first, "")
    digits = []
    for ch in tail:
        code = mapping.get(ch, "0")
        if code != prev:
            digits.append(code)
            prev = code
    digits = [d for d in digits if d != "0"]
    code = first + "".join(digits)
    code = (code + "000")[:4]
    return code

def sorted_token_key(text: str) -> str:
    if not isinstance(text, str):
        return ""
    toks = [t.strip().lower() for t in text.split() if t.strip()]
    toks.sort()
    return " ".join(toks)



# Hybrid blocking pair generation (


import re
from collections import defaultdict
import difflib
import streamlit as st  # assuming you are inside a Streamlit app

# -----------------------------
# Helpers (pure Python)
# -----------------------------
def normalize_tokens(s: str) -> list:
    """
    Lowercase, keep word characters, split into tokens.
    """
    tokens = re.findall(r"\w+", s.lower())
    return tokens

def token_sort_key(s: str) -> str:
    """
    Token-sort normalized representation (used for blocking and comparison).
    """
    toks = normalize_tokens(s)
    toks.sort()
    return " ".join(toks)

def token_sort_score(a: str, b: str) -> int:
    """
    Token-sort similarity using difflib.SequenceMatcher.
    Returns a percentage 0..100 (int).
    """
    na = token_sort_key(a)
    nb = token_sort_key(b)
    ratio = difflib.SequenceMatcher(None, na, nb).ratio()
    return int(round(ratio * 100))

def soundex(s: str) -> str:
    """
    Basic Soundex implementation: returns 4-character code.
    Good enough for blocking by phonetics without external libs.
    """
    s = re.sub(r"[^A-Za-z]", "", s.upper())
    if not s:
        return ""
    first_letter = s[0]

    # Soundex mappings
    mappings = {
        **{c: "1" for c in "BFPV"},
        **{c: "2" for c in "CGJKQSXZ"},
        **{c: "3" for c in "DT"},
        **{c: "4" for c in "L"},
        **{c: "5" for c in "MN"},
        **{c: "6" for c in "R"},
    }

    # Map letters to digits; vowels/H/W/Y become "0"
    digits = []
    for ch in s[1:]:
        digits.append(mappings.get(ch, "0"))

    # Remove consecutive duplicates
    cleaned = []
    prev = None
    for d in digits:
        if d != prev:
            cleaned.append(d)
        prev = d

    # Remove zeros (vowels/H/W/Y)
    cleaned = [d for d in cleaned if d != "0"]

    # Construct code: first letter + 3 digits (pad with zeros)
    code = first_letter + ("".join(cleaned) + "000")[:3]
    return code

def sorted_token_key(s: str) -> str:
    """
    Your original helper name kept for compatibility.
    """
    return token_sort_key(s)

# -----------------------------
# Pair detection without RapidFuzz
# -----------------------------
def fuzzy_pairs_hybrid(df, cols, threshold=85, max_pairs=2_000_000, block_by=None, soundex_cols=None):
    """
    Detect likely duplicate pairs using:
      - token-sort similarity via difflib (no RapidFuzz/fuzzywuzzy)
      - optional Soundex-based blocking
      - token-sort bucket blocking

    Parameters
    ----------
    df : pandas.DataFrame
    cols : list[str]
        Columns to concatenate for text comparison.
    threshold : int (0..100)
        Minimum token-sort similarity (difflib) to accept a pair.
    max_pairs : int
        Safety cap on number of pairs to collect.
    block_by : str or None
        Optional column name to block by exact value (stringified).
    soundex_cols : list[str] or None
        Optional columns to compute Soundex code for phonetic blocking.

    Returns
    -------
    pairs_list : list[tuple[int,int,int]]
        (i, j, score) with i < j
    elapsed_sec : float
    """
    start = time.time()
    n = len(df)
    if n <= 1:
        return [], 0.0

    # Prepare comparison texts (concatenated fields)
    texts = df[cols].fillna("").astype(str).agg(" | ".join, axis=1).tolist()

    # Precompute token-sort normalized strings once (perf)
    norm_texts = [token_sort_key(t) for t in texts]

    # --- Soundex map for blocking ---
    soundex_map = {}
    if soundex_cols:
        for idx, row in df[soundex_cols].fillna("").astype(str).iterrows():
            parts = [soundex(str(row[c])) if str(row[c]).strip() else "" for c in soundex_cols]
            soundex_map[idx] = "|".join(parts)
    else:
        soundex_map = {i: "" for i in range(n)}

    # --- Token-sort key (for additional blocking) ---
    sorted_keys = [sorted_token_key(t) for t in texts]

    # --- Build blocks ---
    blocks = defaultdict(list)

    # 1) Soundex blocks
    if soundex_cols:
        for i, k in soundex_map.items():
            if k.strip():
                blocks[f"sx|{k}"].append(i)

    # 2) Token-sort bucket blocks (first token + bucket off hash)
    for i, sk in enumerate(sorted_keys):
        if not sk:
            continue
        toks = sk.split()
        first = toks[0] if toks else ""
        # Note: Python's hash is randomized per process but it's fine for bucketing.
        bucket = abs(hash(sk)) % 10000
        blocks[f"st|{first}|{bucket}"].append(i)

    # 3) Optional exact value block_by
    if block_by:
        series = df[block_by].fillna("__MISSING__").astype(str)
        for val, idxs in df.groupby(series).indices.items():
            blocks[f"blk|{val}"].extend(list(idxs))

    # Clean blocks: unique indices, keep blocks with >1 element
    cleaned_blocks = []
    for _, idxs in blocks.items():
        uniq = list(dict.fromkeys(idxs))
        if len(uniq) > 1:
            cleaned_blocks.append(uniq)
    if not cleaned_blocks:
        cleaned_blocks = [list(range(n))]

    # --- Pair enumeration with progress ---
    pairs = set()
    total_blocks = len(cleaned_blocks)
    pb = st.progress(0)
    pct_text = st.empty()

    for bi, idxs in enumerate(cleaned_blocks):
        m = len(idxs)
        chunk_size = 2000
        chunks = [idxs[i:i + chunk_size] for i in range(0, m, chunk_size)] if m > chunk_size else [idxs]

        for chunk in chunks:
            L = len(chunk)
            for a in range(L):
                i = chunk[a]
                ni = norm_texts[i]
                for b in range(a + 1, L):
                    j = chunk[b]
                    score = int(round(difflib.SequenceMatcher(None, ni, norm_texts[j]).ratio() * 100))
                    if score >= threshold:
                        if i < j:
                            pairs.add((i, j, score))
                        else:
                            pairs.add((j, i, score))
                    if len(pairs) >= max_pairs:
                        elapsed = round(time.time() - start, 2)
                        pb.progress(100)
                        pct_text.text("100%")
                        return list(pairs), elapsed

        pb.progress(int((bi + 1) / total_blocks * 100))
        pct_text.text(f"Detecting... {int((bi + 1) / total_blocks * 100)}%")

    pb.progress(100)
    pct_text.text("100%")
    elapsed = round(time.time() - start, 2)
    return list(pairs), elapsed

# -----------------------------
# Clustering (unchanged logic)
# -----------------------------
def cluster_from_pairs_list(pairs):
    """
    Union-Find clustering of indices from pair list.
    """
    parent = {}

    def find(x):
        parent.setdefault(x, x)
        if parent[x] != x:
            parent[x] = find(parent[x])
        return parent[x]

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for i, j, _ in pairs:
        union(i, j)

    clusters = {}
    for node in parent:
        root = find(node)
        clusters.setdefault(root, []).append(node)

    clusters_list = list(clusters.values())
    clusters_list.sort(key=lambda x: -len(x))
    return clusters_list



import re
import time
import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------
# Utilities for smart merging
# -----------------------------
def _normalize_phone(x: str) -> str:
    """Digits-only phone normalization, tolerates local/international formats."""
    if pd.isna(x):
        return ""
    digits = re.sub(r"\D+", "", str(x))
    # Keep last up to 12 digits (adjust if you prefer strict validation)
    return digits[-12:] if len(digits) >= 9 else digits

def _pick_best_phone(vals):
    """Pick the longest normalized phone; tie-break by earliest occurrence."""
    cands = [_normalize_phone(v) for v in vals if str(v).strip() != ""]
    cands = [c for c in cands if c]
    if not cands:
        return np.nan
    # sort by length desc, keep first occurrence for ties
    return sorted(range(len(cands)), key=lambda i: (-len(cands[i]), i))[0] and cands[sorted(range(len(cands)), key=lambda i: (-len(cands[i]), i))[0]]

def _clean_text(s):
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    return s

def _pick_longest_text(vals):
    vals = [_clean_text(v) for v in vals if str(v).strip() != ""]
    if not vals:
        return np.nan
    # length desc, earliest occurrence tie-break
    return sorted(range(len(vals)), key=lambda i: (-len(vals[i]), i))[0] and vals[sorted(range(len(vals)), key=lambda i: (-len(vals[i]), i))[0]]

def _pick_mode(vals):
    ser = pd.Series([v for v in vals if pd.notna(v)])
    if ser.empty:
        return np.nan
    counts = ser.value_counts()
    return counts.index[0]

def _pick_earliest_date(vals):
    dt = pd.to_datetime(pd.Series(vals), errors="coerce").dropna()
    return dt.min() if not dt.empty else np.nan

def _pick_latest_date(vals):
    dt = pd.to_datetime(pd.Series(vals), errors="coerce").dropna()
    return dt.max() if not dt.empty else np.nan


# -----------------------------
# Heuristics: classify columns
# -----------------------------
def infer_column_roles(df: pd.DataFrame):
    """
    Infer roles for columns based on names and dtypes:
      - ids: columns likely to be identifiers
      - phones: phone-like columns
      - sum_cols: numeric totals to sum
      - earliest_dates: date-like columns to keep earliest
      - latest_dates: date-like columns to keep latest (updated/modified)
      - mode_cols: categorical columns to keep most frequent
    Returns a dict with sets/lists for each role.
    """
    cols = list(df.columns)
    lower_map = {c: c.lower() for c in cols}

    # name-based hints
    id_like = {"id", "reg", "registration", "student_id", "learner_id", "uuid"}
    phone_like = {"phone", "mobile", "contact", "guardian_phone", "parent_phone", "tel"}
    early_date_like = {"enrol", "enroll", "admission", "start", "created", "dob", "birth"}
    late_date_like  = {"update", "updated", "modified", "last", "end"}
    sum_like = {"total", "sum", "count", "present", "absent", "days", "hours"}
    mode_like = {"gender", "class", "stream", "school", "status", "category"}

    roles = {
        "ids": set(),
        "phones": set(),
        "sum_cols": set(),
        "earliest_dates": set(),
        "latest_dates": set(),
        "mode_cols": set(),
        "name_cols": set()
    }

    for c in cols:
        lc = lower_map[c]
        # IDs
        if any(k in lc for k in id_like):
            roles["ids"].add(c)
        # Phones
        if any(k in lc for k in phone_like):
            roles["phones"].add(c)
        # Names (heuristic)
        if any(k in lc for k in ["name", "fullname", "learner_name", "student_name"]):
            roles["name_cols"].add(c)
        # Dates
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            # classify by name hints
            if any(k in lc for k in early_date_like):
                roles["earliest_dates"].add(c)
            elif any(k in lc for k in late_date_like):
                roles["latest_dates"].add(c)
            else:
                # default: earliest for arbitrary date columns
                roles["earliest_dates"].add(c)
        else:
            # strings that look like dates by name
            if any(k in lc for k in early_date_like):
                roles["earliest_dates"].add(c)
            elif any(k in lc for k in late_date_like):
                roles["latest_dates"].add(c)

        # Sum columns (numeric with name hints)
        if pd.api.types.is_numeric_dtype(df[c]) and any(k in lc for k in sum_like):
            roles["sum_cols"].add(c)

        # Mode columns (categorical-like by name hints)
        if any(k in lc for k in mode_like):
            roles["mode_cols"].add(c)

    return roles


# -----------------------------
# Merge one group → one row
# -----------------------------
def merge_group_generic(
    df: pd.DataFrame,
    indices: list,
    overrides: dict | None = None
):
    """
    Merge a single group of duplicate rows into a consolidated row.
    Smart defaults + optional overrides.

    overrides structure (all optional):
    {
        "ids":        {"columns": [...], "strategy": "first"},
        "phones":     {"columns": [...], "strategy": "best"},
        "name_cols":  {"columns": [...], "strategy": "longest"},
        "sum_cols":   {"columns": [...], "strategy": "sum"},
        "earliest_dates": {"columns": [...], "strategy": "min"},
        "latest_dates":   {"columns": [...], "strategy": "max"},
        "mode_cols":  {"columns": [...], "strategy": "mode"},
        "numeric_default": "first" | "sum",
        "text_default": "first"
    }
    """
    sub = df.loc[indices]
    roles = infer_column_roles(sub)  # infer on subset (ok) or use df

    # Apply overrides to roles if provided
    def apply_override(role_key, default_set):
        if overrides and role_key in overrides and "columns" in overrides[role_key]:
            return set(overrides[role_key]["columns"])
        return default_set

    ids           = apply_override("ids", roles["ids"])
    phones        = apply_override("phones", roles["phones"])
    name_cols     = apply_override("name_cols", roles["name_cols"])
    sum_cols      = apply_override("sum_cols", roles["sum_cols"])
    earliest_dates= apply_override("earliest_dates", roles["earliest_dates"])
    latest_dates  = apply_override("latest_dates", roles["latest_dates"])
    mode_cols     = apply_override("mode_cols", roles["mode_cols"])

    numeric_default = (overrides.get("numeric_default") if overrides else "first")
    text_default    = (overrides.get("text_default") if overrides else "first")

    out = {}
    for col in df.columns:
        s = sub[col]
        vals = s.dropna().tolist()

        # IDs: keep first non-null
        if col in ids:
            out[col] = vals[0] if vals else np.nan
            continue

        # Phones: best normalized
        if col in phones:
            out[col] = _pick_best_phone(vals)
            continue

        # Names: longest cleaned text
        if col in name_cols:
            out[col] = _pick_longest_text(vals)
            continue

        # Dates
        if col in earliest_dates:
            out[col] = _pick_earliest_date(vals)
            continue
        if col in latest_dates:
            out[col] = _pick_latest_date(vals)
            continue

        # Sum columns: sum numerics
        if col in sum_cols:
            out[col] = float(pd.to_numeric(s, errors="coerce").fillna(0).sum())
            continue

        # Numeric columns (default behavior)
        if pd.api.types.is_numeric_dtype(s):
            if numeric_default == "sum":
                out[col] = float(pd.to_numeric(s, errors="coerce").fillna(0).sum())
            else:
                out[col] = vals[0] if vals else np.nan
            continue

        # Datetime dtype not listed above → earliest
        if pd.api.types.is_datetime64_any_dtype(s):
            out[col] = _pick_earliest_date(vals)
            continue

        # Mode columns → most frequent
        if col in mode_cols:
            out[col] = _pick_mode(vals)
            continue

        # Text/object fallback
       
  

# The code block you provided starts here (assuming it is inside a function or script flow)
st.subheader("Build exact groups and merge")

# Check if 'df' exists in session state before trying to access it anywhere
if 'df' not in st.session_state or st.session_state.df is None:
    st.info("Upload a file first.")
else:
    # If the file IS uploaded, access the DataFrame safely
    df = st.session_state.df
    
    # Add your main logic here to use the 'df' DataFrame
    # For now, we keep the original display logic
    st.write(f"Rows: {df.shape[0]} | Columns: {df.shape[1]}")
    
    # Example of where you might add more processing steps:
    # st.write("Ready for further processing of the dataframe.")


     

df = st.session_state.get("df")
if df is None:
    st.warning("Upload a dataset first.")
else:
    key_cols = st.multiselect(
        "Select columns to define exact duplicates (group by these keys)",
        options=list(df.columns),
        default=[c for c in df.columns if "id" in c.lower()][:1] or list(df.columns[:1])
    )

    # Optional overrides (example)
    use_sum_for_numeric = st.checkbox("Sum numeric columns by default", value=False)
    text_strategy = st.selectbox("Text fallback strategy", ["first", "longest", "mode"], index=0)

    if st.button("Detect groups & Merge"):
        # Build groups by exact key
        key_series = df[key_cols].fillna("").astype(str).agg("|".join, axis=1)
        groups_map = {}
        for idx, k in enumerate(key_series):
            groups_map.setdefault(k, []).append(idx)
        groups = [idxs for idxs in groups_map.values() if len(idxs) > 1]

        st.write(f"Found {len(groups)} duplicate groups based on {key_cols}.")

        overrides = {
            "numeric_default": "sum" if use_sum_for_numeric else "first",
            "text_default": text_strategy
        }

        df_merged, removed_count, secs = merge_groups_with_progress_generic(df, groups, overrides=overrides)
        st.session_state.df = df_merged
        st.success(f"Merged {len(groups)} groups. Removed {removed_count} original rows in {secs}s.")
        st.dataframe(df_merged.head(30))




# Session state init

import streamlit as st
import pandas as pd # You will likely use this later in your app

# --- Initialize Session State for Exact-Only Workflow ---
# This block ensures all required keys exist before the rest of the script runs.

for key in [
    "step",                # current step in the app
    "df_raw",              # original uploaded DataFrame
    "df",                  # working DataFrame after cleaning
    "exact_groups",        # detected duplicate groups
    "exact_detect_time",   # time taken for detection
    "exact_merge_time",    # time taken for merging
    "exact_removed",       # number of rows removed during merge
    "last_snapshot",       # backup before last merge (for undo)
    "audit_log",           # log of actions performed
    "show_exact_confirm"   # flag for merge confirmation UI
]:
    # Check if the key is missing from the session state
    if key not in st.session_state:
        # If missing, initialize it with a default value (in this case, None)
        st.session_state[key] = None

# You can optionally set more specific default values after the loop, if needed:
# if st.session_state.step is None:
#     st.session_state.step = 1
# if st.session_state.audit_log is None:
#     st.session_state.audit_log = []








# Set initial step if not defined
if st.session_state.step is None:
    st.session_state.step = 1


# Sidebar

ith st.sidebar:
    st.markdown("### 🔍 Audit Log")
    if st.session_state.audit_log:
        # Show the most recent 100 actions (top-first or reverse as you prefer)
        for a in st.session_state.audit_log[:100]:
            st.write(f"- {a}")
    else:
        st.write("No actions yet.")

    st.write("---")

    # Restart wizard button
    if st.button("Restart Wizard"):
        for k in [
            "step", "df_raw", "df",
            "exact_groups", "exact_detect_time", "exact_merge_time", "exact_removed",
            "last_snapshot", "audit_log", "show_exact_confirm"
        ]:
            st.session_state[k] = None

        # Start at step 1
        st.session_state.step = 1

        # Rerun app (Streamlit >= 1.27 has st.rerun; older versions use st.experimental_rerun)
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()

    # Undo last merge/delete if we have a snapshot
    if st.session_state.last_snapshot is not None:
        if st.button("Undo Last Merge/Delete"):
            st.session_state.df = st.session_state.last_snapshot.copy()

            # Optional helper to record actions in audit log
            def log_action(msg: str):
                if st.session_state.audit_log is None:
                    st.session_state.audit_log = []
                st.session_state.audit_log.append(msg)

            log_action("Undo last merge/delete")

            # Clear snapshot after undo
            st.session_state.last_snapshot = None

            # Rerun to refresh UI/data
            try:
                st.rerun()
            except AttributeError:
                st.experimental_rerun()


# Navigation buttons later at bottom
st.markdown("---")

# Upload your data 

if st.session_state.step == 1:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.header("Step 1 — Upload dataset (CSV / Excel)")
    uploaded = st.file_uploader("Upload file", type=["csv","xlsx","xls"], key="upload1")
    if uploaded:
        try:
            if uploaded.name.lower().endswith(".csv"):
                st.session_state.df_raw = pd.read_csv(uploaded)
            else:
                st.session_state.df_raw = pd.read_excel(uploaded, engine="openpyxl")
            st.session_state.df = st.session_state.df_raw.copy()
            st.success(f"Loaded {st.session_state.df.shape[0]} rows × {st.session_state.df.shape[1]} columns")
            log_action(f"Uploaded file: {uploaded.name}")
            st.dataframe(st.session_state.df.head(5))
        except Exception as e:
            st.error(f"Failed to load file: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

Next Clean & summary 


if st.session_state.step == 2:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.header("Next — Cleaning & Summary")

    if st.session_state.df is None:
        st.info("Upload a file in the previous section first.")
    else:
        df = st.session_state.df

        # --- Summary cards ---
        st.write(f"Rows: **{df.shape[0]}** | Columns: **{df.shape[1]}**")
        st.write("Top missing counts (by column):")
        st.dataframe(
            df.isna().sum().sort_values(ascending=False).to_frame("missing_count")
        )
        st.write(f"Exact duplicate row count: **{df.duplicated().sum()}**")

        # Optional: controls for cleaning behavior
        with st.expander("Cleaning options", expanded=True):
            remove_dupes = st.checkbox("Remove exact duplicate rows", value=True)
            coerce_numeric = st.checkbox("Coerce numeric-like columns to numbers", value=True)
            coerce_dates = st.checkbox("Coerce date-like columns to datetime", value=True)
            title_case_names = st.checkbox("Title-case name-like columns", value=True)
            normalize_phones = st.checkbox("Normalize phone columns", value=True)

        if st.button("Run basic clean"):
            start = time.time()
            with st.spinner("Cleaning..."):
                pb = st.progress(0)

                cleaned = basic_clean(
                    df,
                    remove_duplicates=remove_dupes,
                    coerce_numeric=coerce_numeric,
                    coerce_dates=coerce_dates,
                    title_case_names=title_case_names,
                    normalize_phones=normalize_phones,
                )

                pb.progress(100)

            elapsed = round(time.time() - start, 2)

            # Reset exact groups after cleaning (fresh detection will be needed)
            st.session_state.exact_groups = []

            # REMOVE fuzzy references entirely
            # st.session_state.fuzzy_runs_results = None  # <-- deleted

            log_action(f"Basic clean applied ({elapsed}s)")
            st.success(f"Cleaning finished in {elapsed} seconds.")
            st.dataframe(cleaned.head(10))

            # Save back to session
            st.session_state.df = cleaned

    st.markdown("</div>", unsafe_allow_html=True)



 Exact detect & optional merge


elif st.session_state.step == 3:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.header("Next — Exact Duplicate Detection")

    if st.session_state.df is None:
        st.info("Upload & clean first.")
    else:
        df = st.session_state.df
        col_opts = list(df.columns)

        # Smart default: try Name/Sex/Gender/School; else first 2 columns
        default_cols = [c for c in col_opts if any(k in c.lower() for k in ["name", "sex", "gender", "school"])][:3]
        if len(default_cols) < 2:
            default_cols = col_opts[:2]

        exact_cols = st.multiselect(
            "Select 2–3 columns to define duplicates (rows with the SAME values across all selected columns are duplicates)",
            options=col_opts,
            default=default_cols,
            help="Examples: Name + Sex + School; or Student_ID + Class; or Name + School."
        )

        # Optional: ID columns that should never be summed during merge
        id_cols_input = st.text_input(
            "Comma-separated ID columns (these will NOT be summed during merge)",
            value=""
        )
        id_cols = [c.strip() for c in id_cols_input.split(",")] if id_cols_input.strip() else []

        # Detect exact duplicates
        if st.button("Detect exact duplicates"):
            # Enforce 2–3 columns
            if not exact_cols or len(exact_cols) < 2 or len(exact_cols) > 3:
                st.warning("Please select **2 or 3** columns to define duplicates.")
            else:
                start = time.time()
                with st.spinner("Detecting exact duplicates..."):
                    pb = st.progress(0)

                    # If you already have detect_exact_groups(df, exact_cols), use it.
                    # Otherwise, uncomment the simple implementation provided below.
                    groups = detect_exact_groups(df, exact_cols)

                    pb.progress(100)
                elapsed = round(time.time() - start, 2)

                st.session_state.exact_groups = groups
                st.session_state.exact_detect_time = elapsed
                st.success(f"Found **{len(groups)}** exact duplicate groups in **{elapsed}** seconds.")
                log_action(f"Exact detection finished ({len(groups)} groups in {elapsed}s)")

                # Preview
                if groups:
                    st.write("Preview of first 5 exact duplicate groups:")
                    for g in groups[:5]:
                        st.write(df.loc[g])

                    # Download duplicates
                    try:
                        dupdf = pd.concat([df.loc[g] for g in groups], ignore_index=True)
                        data_bytes, ftype = df_to_bytes(dupdf)
                        if ftype == "excel":
                            st.download_button(
                                "Download Exact Duplicates (Excel)",
                                data=data_bytes,
                                file_name="exact_duplicates.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.download_button(
                                "Download Exact Duplicates (CSV)",
                                data=data_bytes,
                                file_name="exact_duplicates.csv",
                                mime="text/csv"
                            )
                    except Exception as e:
                        st.error(f"Prepare download failed: {e}")

        # Merge & delete exact duplicates (confirmation gate)
        if st.session_state.exact_groups:
            st.markdown("<div class='danger'>", unsafe_allow_html=True)
            st.write("⚠️ **WARNING:** You are about to **MERGE & DELETE** exact duplicates.")
            if st.button("Proceed"):
                st.session_state.show_exact_confirm = True
            if st.button("Cancel"):
                st.session_state.show_exact_confirm = False
            st.markdown("</div>", unsafe_allow_html=True)

            if st.session_state.get("show_exact_confirm", False):
                st.warning("Confirm merge & delete exact duplicates.")
                yes_col, no_col = st.columns([1, 1])

                with yes_col:
                    if st.button("Yes"):
                        st.session_state.last_snapshot = st.session_state.df.copy()

                        # Heuristics to assist merge strategy
                        name_cols = [c for c in df.columns if "name" in c.lower()]
                        phone_candidates = [c for c in df.columns if any(p in c.lower() for p in ["phone", "mobile", "contact", "tel"])]

                        with st.spinner("Merging exact duplicates..."):
                            # Use the merge helper you added earlier:
                            # If you adopted the generic merge:
                            # df_new, removed, elapsed_merge = merge_groups_with_progress_generic(
                            #     st.session_state.df, st.session_state.exact_groups,
                            #     overrides={"ids": {"columns": id_cols}}
                            # )
                            # Otherwise, use your original keep-first merge:
                            df_new, removed, elapsed_merge = merge_groups_with_progress_keep_first(
                                st.session_state.df,
                                st.session_state.exact_groups,
                                id_cols=id_cols,
                                phone_cols=phone_candidates,
                                name_cols=name_cols
                            )

                        st.session_state.df = df_new
                        st.session_state.exact_removed = removed
                        st.session_state.exact_merge_time = round(elapsed_merge, 2)
                        st.session_state.exact_groups = []
                        st.session_state.show_exact_confirm = False

                        log_action(f"Exact merged & deleted ({removed} rows in {st.session_state.exact_merge_time}s)")
                        st.success(f"Merged & deleted **{removed}** exact duplicate rows in **{st.session_state.exact_merge_time}** seconds.")

                with no_col:
                    if st.button("No — Cancel exact merge"):
                        st.session_state.show_exact_confirm = False

    st.markdown("</div>", unsafe_allow_html=True)


ef detect_exact_groups(df: pd.DataFrame, by_cols: list[str]) -> list[list[int]]:
    """
    Return a list of index lists; each index list is one exact-duplicate group
    (same values across all `by_cols`), size >= 2.
    """
    # Build a key per row
    key_series = df[by_cols].fillna("__NA__").astype(str).agg("|".join, axis=1)
    groups_map = {}
    for idx, k in enumerate(key_series):
        groups_map.setdefault(k, []).append(idx)
    # Keep only groups with >1 row
    return [idxs for idxs in groups_map.values() if len(idxs) > 1]




# Step 6: Final download & summary

elif st.session_state.step == 6:
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.header("Final — Download & Summary")

    if st.session_state.df is None:
        st.info("No dataset available.")
    else:
        df = st.session_state.df

        # --- Download final cleaned dataset ---
        try:
            bytes_out, ftype = df_to_bytes(df)
            if ftype == "excel":
                st.download_button(
                    "Download Final Cleaned Excel",
                    data=bytes_out,
                    file_name="cleaned_dataset.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.download_button(
                    "Download Final Cleaned CSV",
                    data=bytes_out,
                    file_name="cleaned_dataset.csv",
                    mime="text/csv"
                )
        except Exception as e:
            st.error(f"Download preparation failed: {e}")

        st.markdown("### Summary of actions")
        st.write(f"- Exact detection time: **{st.session_state.exact_detect_time if st.session_state.exact_detect_time else 0}** s")
        st.write(f"- Exact rows removed (merged): **{st.session_state.exact_removed if st.session_state.exact_removed else 0}**")
        st.write(f"- Exact merge time: **{st.session_state.exact_merge_time if st.session_state.exact_merge_time else 0}** s")
        st.write(f"- Final rows: **{df.shape[0]}** | Columns: **{df.shape[1]}**")

        # --- District & Sex summary (auto-detect columns) ---
        st.markdown("### Clean Data Summary — by District and Sex")

        # Try to detect likely district and sex columns
        cols_lower = {c.lower(): c for c in df.columns}
        # District detection
        district_candidates = [c for c in df.columns if any(k in c.lower() for k in ["district", "location_district", "dist"])]
        # Sex/Gender detection
        sex_candidates = [c for c in df.columns if any(k in c.lower() for k in ["sex", "gender"])]

        # Allow user to confirm/override detected columns
        col1, col2 = st.columns(2)
        with col1:
            district_col = st.selectbox(
                "Select District column",
                options=["(None)"] + df.columns.tolist(),
                index=(df.columns.tolist().index(district_candidates[0]) + 1) if district_candidates else 0
            )
        with col2:
            sex_col = st.selectbox(
                "Select Sex/Gender column",
                options=["(None)"] + df.columns.tolist(),
                index((df.columns.tolist().index(sex_candidates[0]) + 1) if sex_candidates else 0)
            )

        # Safe handling if columns are missing or user picks (None)
        chosen_district = None if district_col == "(None)" else district_col
        chosen_sex = None if sex_col == "(None)" else sex_col

        if chosen_district and chosen_sex:
            # Build a count pivot table: rows = district, columns = sex/gender
            try:
                summary = (
                    df
                    .assign(_one_=1)
                    .pivot_table(
                        index=chosen_district,
                        columns=chosen_sex,
                        values="_one_",
                        aggfunc="count",
                        fill_value=0,
                        observed=False
                    )
                    .sort_index()
                )

                # Add row totals and grand total
                summary["Total"] = summary.sum(axis=1)
                summary.loc["Total"] = summary.sum(axis=0)

                st.dataframe(summary)
            except Exception as e:
                st.warning(f"Could not build district/sex summary: {e}")
        else:
            # Fallback summaries if columns not available
            st.info("Select both a District column and a Sex/Gender column to view the breakdown.")
            if district_candidates:
                st.write("Detected possible district columns:", ", ".join(district_candidates))
            if sex_candidates:
                st.write("Detected possible sex/gender columns:", ", ".join(sex_candidates))

        st.success("Process complete — use Undo in the sidebar to revert the most recent merge if needed.")

    st.markdown("</div>", unsafe_allow_html=True)



# Bottom navigation

st.markdown("<div class='footer-space'></div>", unsafe_allow_html=True)

nav_left, nav_right = st.columns([1, 1])

with nav_left:
    if st.button("◀ Previous"):
        st.session_state.step = max(1, st.session_state.step - 1)
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()

with nav_right:
    if st.button("Next ▶"):
        st.session_state.step = min(6, st.session_state.step + 1)
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()


# --- CSS: sticky footer ---
st.markdown("""
<style>
/* Root variables to adjust colors for light/dark themes */
:root {
    --footer-bg: rgba(248, 250, 252, 0.92); /* light: slate-50 */
    --footer-border: #e2e8f0;               /* slate-200 */
    --footer-text: #334155;                 /* slate-700 */
    --footer-link: #0ea5e9;                 /* sky-500 */
}
@media (prefers-color-scheme: dark) {
    :root {
        --footer-bg: rgba(15, 23, 42, 0.85); /* dark: slate-900-ish */
        --footer-border: #334155;            /* slate-700 */
        --footer-text: #cbd5e1;              /* slate-300 */
        --footer-link: #38bdf8;              /* sky-400 */
    }
}

/* Sticky footer container */
.footer-credit {
    position: fixed;
    left: 0;
    right: 0;
    bottom: 0;
    z-index: 9999;
    display: flex;
    justify-content: center;
    align-items: center;

    padding: 10px 16px;
    background: var(--footer-bg);
    color: var(--footer-text);
    font-size: 0.95rem;
    border-top: 1px solid var(--footer-border);

    /* Subtle glass effect */
    backdrop-filter: blur(6px);
    -webkit-backdrop-filter: blur(6px);
}

/* Inner content for better layout control */
.footer-credit .content {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    max-width: 1200px;
    width: 100%;
    justify-content: center;
    flex-wrap: wrap;
}

/* Optional dot accent */
.footer-credit .dot {
    width: 8px;
    height: 8px;
    background: #22c55e; /* green-500 */
    border-radius: 50%;
    display: inline-block;
}

/* Emphasis for labels and name */
.footer-credit .label { 
    font-weight: 600; 
}
.footer-credit .name { 
    font-weight: 600; 
}

/* Optional link styling */
.footer-credit a {
    color: var(--footer-link);
    text-decoration: none;
}
.footer-credit a:hover {
    text-decoration: underline;
}

/* Make sure Streamlit main content doesn't hide behind the footer */
.main-spacer {
    height: 56px; /* matches approx footer height */
}

/* Small screens: tighter spacing */
@media (max-width: 480px) {
    .footer-credit { font-size: 0.9rem; padding: 8px 12px; }
    .main-spacer { height: 48px; }
}
</style>
""", unsafe_allow_html=True)

# --- Optional content / layout above ---
st.title("My Streamlit App")
st.write("Your app content goes here...")

# --- spacer to prevent overlap with bottom content ---
st.markdown('<div class="main-spacer"></div>', unsafe_allow_html=True)

# --- Sticky footer HTML ---
st.markdown(
    '''
    <div class="footer-credit">
        <div class="content">
            <span class="dot"></span>
            <span class="label">Built by</span>
            <span class="name">Moses Eridu</span>
            <!-- Optional: add a link -->
            <!-- <span>&middot;</span> https://example.comPortfolio</a> -->
        </div>
    </div>
    ''',
    unsafe_allow_html=True
)








