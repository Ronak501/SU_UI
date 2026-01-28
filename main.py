import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# -------------------------------------------------
# Page config
# -------------------------------------------------
st.set_page_config(
    page_title="Admission Source Analysis",
    layout="wide"
)

st.title("🎓 Admission Source Analysis Dashboard")

# -------------------------------------------------
# File Upload
# -------------------------------------------------
st.sidebar.header("📂 Upload Excel File")
uploaded_file = st.sidebar.file_uploader(
    "Upload Admission Data (.xls / .xlsx)",
    type=["xls", "xlsx"]
)

if uploaded_file is None:
    st.warning("⚠️ Please upload an Excel file to continue.")
    st.stop()

# -------------------------------------------------
# Load & Clean Data
# -------------------------------------------------
def load_data(file):
    try:
        # Case 1: CSV file
        if file.name.lower().endswith(".csv"):
            df = pd.read_csv(file)

        # Case 2: Old Excel (.xls)
        elif file.name.lower().endswith(".xls"):
            df = pd.read_excel(file, engine="xlrd")

        # Case 3: Modern Excel (.xlsx)
        elif file.name.lower().endswith(".xlsx"):
            df = pd.read_excel(file, engine="openpyxl")

        else:
            st.error("❌ Unsupported file format. Please upload CSV or Excel.")
            st.stop()

    except Exception as e:
        st.error("❌ Unable to read the uploaded file.")
        st.error("Possible reasons:")
        st.markdown("""
        - File is corrupted  
        - File is not a real Excel file  
        - CSV renamed as .xlsx  
        - Google Sheet exported incorrectly  
        """)
        st.code(str(e))
        st.stop()

    df = df.drop_duplicates()
    df.columns = df.columns.str.strip()
    return df


df = load_data(uploaded_file)

st.success(f"✅ File loaded successfully ({len(df):,} records)")

# -------------------------------------------------
# Column Selection
# -------------------------------------------------
st.sidebar.header("⚙️ Column Settings")

SOURCE_COL = st.sidebar.selectbox(
    "Select Admission Source Column",
    df.columns
)

CONFIRM_COL = st.sidebar.selectbox(
    "Select Confirmation Column",
    df.columns
)

# Clean text values
df[SOURCE_COL] = df[SOURCE_COL].astype(str).str.strip().str.title()
df[CONFIRM_COL] = df[CONFIRM_COL].astype(str).str.strip().str.title()

# -------------------------------------------------
# Top-N Control
# -------------------------------------------------
st.sidebar.header("📊 Display Settings")

TOP_N = st.sidebar.slider(
    "Show Top N Admission Sources",
    min_value=5,
    max_value=30,
    value=15
)

# -------------------------------------------------
# Aggregation (CORRECT & SAFE)
# -------------------------------------------------
cross_tab = pd.crosstab(
    df[SOURCE_COL],
    df[CONFIRM_COL]
)

# Sort by total count
cross_tab["Total"] = cross_tab.sum(axis=1)
cross_tab = (
    cross_tab
    .sort_values("Total", ascending=False)
    .drop(columns="Total")
    .head(TOP_N)
)

# -------------------------------------------------
# Crosstab Preview
# -------------------------------------------------
st.subheader("📊 Admission Source vs Confirmation Status (Top Sources)")
st.dataframe(cross_tab, use_container_width=True)

# -------------------------------------------------
# Matplotlib Bar Chart (EXACT LOGIC LIKE YOUR CODE)
# -------------------------------------------------
st.subheader("📊 Matplotlib Bar Chart")

fig, ax = plt.subplots(figsize=(14, 7))

cross_tab.plot(
    kind="bar",
    ax=ax
)

ax.set_title("Admission Source vs Confirmation Status")
ax.set_xlabel("Admission Source")
ax.set_ylabel("Number of Students")
ax.legend(title="Confirmed")
plt.xticks(rotation=60, ha="right")
plt.tight_layout()

st.pyplot(fig)

# -------------------------------------------------
# Seaborn Bar Chart (AGGREGATED)
# -------------------------------------------------
st.subheader("📈 Seaborn Visualization")

plot_df = cross_tab.reset_index().melt(
    id_vars=SOURCE_COL,
    var_name=CONFIRM_COL,
    value_name="Count"
)

fig2, ax2 = plt.subplots(figsize=(14, 7))

sns.barplot(
    data=plot_df,
    x=SOURCE_COL,
    y="Count",
    hue=CONFIRM_COL,
    ax=ax2
)

ax2.set_title("Admission Source vs Confirmation Status")
ax2.set_xlabel("Admission Source")
ax2.set_ylabel("Count")
plt.xticks(rotation=60, ha="right")
plt.tight_layout()

st.pyplot(fig2)

# -------------------------------------------------
# Raw Data Preview
# -------------------------------------------------
with st.expander("🔍 Preview Raw Data (Sample)"):
    st.dataframe(df.sample(min(500, len(df))), use_container_width=True)
