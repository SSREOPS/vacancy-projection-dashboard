import streamlit as st #type: ignore
import pandas as pd #type: ignore
from datetime import datetime, timedelta
from io import BytesIO


st.set_page_config(page_title="Vacancy Projection & Analytics Dashboard", layout="wide")

app_desc = """This interactive dashboard enables users to upload raw vacancy reports, transform the data into a structured format, and generate vacancy projections based on a selected date.\n
Users can dynamically filter by unit status, analyze projected vacancies, and create customized pivot summaries for actionable insights across properties and management teams."""

st.title("Vacancy Projection & Analytics Dashboard", help=app_desc, text_alignment="center")
st.caption("Analyze, project, and visualize vacancy trends with dynamic filtering and pivot analysis.", text_alignment="center")

# -----------------------------
# Helper: Next Monday
# -----------------------------
def get_next_monday():
    today = datetime.today()
    days_ahead = 0 - today.weekday() + 7  # next Monday
    return today + timedelta(days=days_ahead)

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# -----------------------------
# GRID STRUCTURE
# -----------------------------
top_left,    top_middle,    top_right    = st.columns([1.95, 1.05, 5])
middle_left, middle_middle, middle_right = st.columns([1.95, 3.05, 3])
bottom_left, bottom_right                = st.columns([1.95, 6.05])

# -----------------------------
# Upload file
# -----------------------------
with top_left.container(border=True, height="stretch", horizontal_alignment="center", vertical_alignment="center"):
    uploaded = st.file_uploader("Upload YSR 1001 Vacancy Report as downloaded", type=["xlsx"], max_upload_size=1)
    #st.stop() if not uploaded else ""
    
    # ✅ Show sample download if no file uploaded
    if not uploaded:
        st.caption("Download a sample file for testing.")

        with open("VacancyReport_Dummy.xlsx", "rb") as file:
            st.download_button(
                label="📥 Download Sample Vacancy Report",
                data=file,
                file_name="VacancyReport_Dummy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.stop()  # ✅ Stop app AFTER showing download

# -----------------------------
# Projection date (Next Monday default)
# -----------------------------
with top_middle.container(border=True, height="stretch", horizontal_alignment="center", vertical_alignment="center"):
    projection_date = st.date_input("Select Projection Date", value=get_next_monday(), format="MM/DD/YYYY")

# -----------------------------
# PROCESS DATA
# -----------------------------
df = pd.read_excel(uploaded)
df = df.rename(columns=df.iloc[0]).iloc[1:]
df = df[['#', 'Property Code', 'Unit #', 'Unit Type', 'Manager Name', 'Regional Manager', 'Move in', 'Move out']]
df = df[~df['Property Code'].isin(['Rent Mtx vs Rent', 'Variance Total', 'Variance Average', 'Property Code'])]
df.loc[:, 'Status'] = df.iloc[:, 0].where(df.iloc[:, 0].astype(str).str.contains('[A-Za-z]', na=False) & ~df.iloc[:, 0].astype(str).str.contains(r'^\d+$', na=False)).ffill()
df = df[~df['#'].isin(['READY UNITS', 'NOT READY UNITS', 'NOTICE UNITS', 'EVICTION UNITS', 'NON-RENT UNITS'])]

# -----------------------------
# status Filter (NEW)
# -----------------------------
with top_right.container(border=True, height="stretch", horizontal_alignment="center", vertical_alignment="center"):
    st.subheader("Filter Vacancy by Status", text_alignment="center")

    status = df['Status'].dropna().unique().tolist()

    cols_status = st.columns(5)

    selected_status = []

    for i, s in enumerate(status):
        with cols_status[i % 5]:
            if st.checkbox(s, value=True):
                selected_status.append(s)

if not selected_status:
    st.warning("Please select at least one status.")
    st.stop()

df = df[df['Status'].isin(selected_status)]


# -----------------------------
# Vacancy Calculation
# -----------------------------
d = pd.to_datetime(projection_date)

mi = pd.to_datetime(df['Move in'], errors='coerce')
mo = pd.to_datetime(df['Move out'], errors='coerce')

vacancy_col = f'Vacant As Of {projection_date.strftime("%m.%d.%Y")}'

df.loc[:, vacancy_col] = (mo.isna() | (mo <= d)) & (mi.isna() | (mi > d))


# -----------------------------
# PIVOT BUILDER
# -----------------------------
desc = "Customize your view by selecting grouping fields and analyze projected vacancy distribution across properties, managers, and unit types."

# -----------------------------
# (2,1) TITLE
# -----------------------------
with middle_left.container(border=True, height="stretch", horizontal_alignment="center", vertical_alignment="center"):
    st.subheader("Vacancy Summary & Analysis", text_alignment="center", help = desc)


# -----------------------------
# (2,2) COLUMNS SELECTOR
# -----------------------------
with middle_middle.container(border=True, height="stretch", horizontal_alignment="center", vertical_alignment="center"):
    st.markdown("### Columns", text_alignment="center")

    # ✅ IMPORTANT → use session state (pre-read rows)
    columns = ['Regional Manager', 'Manager Name', 'Property Code', 'Unit Type', 'Status']

    index_cols = [
        col for col in columns
        if st.session_state.get(f"row_{col}", col == "Regional Manager")
    ]

    available_columns = [col for col in columns if col not in index_cols]

    column_col = st.pills(" ", available_columns, key="column_selector")


# -----------------------------
# (2,3) METRICS
# -----------------------------
total_units = len(df)
vacant_units = df[vacancy_col].sum()

st.markdown("""
<style>

/* Center entire metric block */
div[data-testid="stMetric"] {
    display: flex;
    flex-direction: column;
    align-items: center !important;
}

/* Center label container */
div[data-testid="stMetricLabel"] {
    width: 100%;
    display: flex;
    justify-content: center !important;
}

/* Center label text INNER wrapper */
div[data-testid="stMetricLabel"] > div {
    text-align: center !important;
    width: 100%;
}

/* Center value */
div[data-testid="stMetricValue"] {
    width: 100%;
    display: flex;
    justify-content: center !important;
}

</style>
""", unsafe_allow_html=True)


with middle_right.container(border=True, height="stretch", horizontal_alignment="center", vertical_alignment="center"):
    st.subheader("For Selected Status Only", text_alignment="center")

    st.metric("Total Projected Vacant Units / Total Units on Vacancy Report", str(int(vacant_units))+"/"+str(total_units))


# -----------------------------
# (3,1) ROWS CHECKBOXES
# -----------------------------
with bottom_left.container(border=True, height="stretch", horizontal_alignment="right"):
    st.markdown("### Rows", text_alignment="right")

    selected_index_cols = []

    for col in columns:
        default_checked = (col == "Regional Manager")

        if st.checkbox(col, value=default_checked, key=f"row_{col}"):
            selected_index_cols.append(col)

    index_cols = selected_index_cols


# -----------------------------
# (3,2) PIVOT TABLE
# -----------------------------
with bottom_right.container(border=True, height="stretch"):
    st.markdown("### Projected Vacancy Pivot")

    pivot_df = df.copy()
    pivot_df['Projected Vacant Units'] = pivot_df[vacancy_col].astype(int)

    if index_cols:
        pivot = pd.pivot_table(
            pivot_df,
            index=index_cols,
            columns=column_col,
            values='Projected Vacant Units',
            aggfunc='sum',
            fill_value=0
        )

        # ✅ Status ordering
        desired_order = ['READY UNITS', 'NOT READY UNITS', 'NOTICE UNITS', 'EVICTION UNITS', 'NON-RENT UNITS']

        if column_col == "Status":
            pivot = pivot.reindex(columns=desired_order, fill_value=0)

        # ✅ Total column
        pivot['Projected Vacant Units'] = pivot_df.groupby(index_cols)['Projected Vacant Units'].sum()

        pivot_display = pivot.reset_index()

        st.dataframe(pivot_display, width="content", hide_index=True)

        # ✅ TOTAL ROW
        total_row = pivot.sum(numeric_only=True).to_frame().T

        for col in index_cols:
            total_row.insert(len(total_row.columns) - len(pivot.columns), col, "")

        total_row[index_cols[0]] = "Total"
        total_row = total_row[pivot_display.columns]

        st.dataframe(total_row, width="content", hide_index=True)

        st.download_button(
            "Download Projected Vacancy Pivot (Excel)",
            to_excel_bytes(pivot_display),
            file_name=f"Projected_Vacant_Summary_{projection_date.strftime('%m.%d.%Y')}.xlsx"
        )
    else:
        st.warning("Please select at least one Row Category.")

# -----------------------------
# PREVIEW
# -----------------------------
with st.expander("See Processed Vacancy Data"):

    st.subheader("Processed Vacancy Data")

    df_display = df.copy()

    df_display['Move in'] = pd.to_datetime(df_display['Move in'], errors='coerce').dt.strftime('%m/%d/%Y').fillna('')
    df_display['Move out'] = pd.to_datetime(df_display['Move out'], errors='coerce').dt.strftime('%m/%d/%Y').fillna('')

    st.dataframe(df_display, width="content", hide_index=True)

    # -----------------------------
    # DOWNLOAD DATA
    # -----------------------------
    st.download_button(
        "Download Processed Vacancy Data (Excel)",
        to_excel_bytes(df),
        file_name="processed_vacancy_data.xlsx"
    )

st.markdown("---")
st.caption(
    "Note: Vacancy projections are based on Move-In and Move-Out dates available in the uploaded report. "
    "Units with missing dates are treated as vacant as of the selected projection date."
)
