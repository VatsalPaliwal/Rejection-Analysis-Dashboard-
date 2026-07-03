import streamlit as st
import pandas as pd
import re
from io import BytesIO
import plotly.express as px
import traceback

DEFECT_COLS = [
    "FPPA", "U/O SIZE", "SHRINKAGE", "B.SPOT", "FLASH", "COLOR",
    "CRACK", "SHORT", "BURN MARK", "EJECT. MARK", "WELD LINE",
    "BEND", "S.STREAK", "BOP ISSUE", "SCRATCH", "OTHER"
]

MONTH_ORDER = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

MONTH_ABBR = {
    "JANUARY": "JAN", "FEBRUARY": "FEB", "MARCH": "MAR",
    "APRIL": "APR", "MAY": "MAY", "JUNE": "JUN",
    "JULY": "JUL", "AUGUST": "AUG", "SEPTEMBER": "SEP",
    "OCTOBER": "OCT", "NOVEMBER": "NOV", "DECEMBER": "DEC"
}

st.set_page_config(page_title="Rejection Analysis", layout="wide")

if "month_data" not in st.session_state:
    st.session_state.month_data = {}
if "file_id" not in st.session_state:
    st.session_state.file_id = None


def month_sort_key(month_key: str):
    """
    Chronological sort key for keys shaped like 'JAN-25'.
    """
    abbr, year = month_key.split("-")
    month_full = next(m for m, a in MONTH_ABBR.items() if a == abbr).capitalize()
    return (int(year), MONTH_ORDER.index(month_full))


def parse_available_months(xl: pd.ExcelFile) -> dict:
    """
    Only PROCESS sheets are expected in the uploaded workbook now —
    no FINAL, no COMPILED. Returns {month_key: sheet_name}.
    """

    months = {}

    month_lookup = {
        "JAN": "JAN", "JANUARY": "JAN",
        "FEB": "FEB", "FEBRUARY": "FEB",
        "MAR": "MAR", "MARCH": "MAR",
        "APR": "APR", "APRIL": "APR",
        "MAY": "MAY",
        "JUN": "JUN", "JUNE": "JUN",
        "JUL": "JUL", "JULY": "JUL",
        "AUG": "AUG", "AUGUST": "AUG",
        "SEP": "SEP", "SEPT": "SEP", "SEPTEMBER": "SEP",
        "OCT": "OCT", "OCTOBER": "OCT",
        "NOV": "NOV", "NOVEMBER": "NOV",
        "DEC": "DEC", "DECEMBER": "DEC",
    }

    for sheet in xl.sheet_names:

        name = sheet.upper().strip()

        if "PROCESS" not in name:
            continue

        words = re.split(r"[\s_\-/]+", name)

        # Find the month token's position so we can look at its
        # immediate neighbors for the year, instead of scanning the
        # whole string and risking a false match on an unrelated number.
        month = None
        month_idx = None
        for i, word in enumerate(words):
            if word in month_lookup:
                month = month_lookup[word]
                month_idx = i
                break

        if month is None:
            continue

        neighbors = []
        if month_idx > 0:
            neighbors.append(words[month_idx - 1])
        if month_idx < len(words) - 1:
            neighbors.append(words[month_idx + 1])

        year = None
        for tok in neighbors:
            m4 = re.fullmatch(r"20\d{2}", tok)
            m2 = re.fullmatch(r"\d{2}", tok)
            if m4:
                year = tok[2:]
                break
            elif m2 and year is None:
                year = tok

        # Fall back to a whole-string scan only if nothing adjacent
        # to the month worked.
        if year is None:
            four = re.search(r"(20\d{2})", name)
            two = re.search(r"(?<!\d)(\d{2})(?!\d)", name)
            if four:
                year = four.group(1)[2:]
            elif two:
                year = two.group(1)

        if year is None:
            continue

        month_key = f"{month}-{year}"
        months[month_key] = sheet

    return dict(sorted(months.items(), key=lambda x: month_sort_key(x[0])))


def find_header_row(df_raw):
    for i, row in df_raw.iterrows():
        vals = [str(v).strip().upper() for v in row.values]
        if "DATE" in vals and "COMPONENT" in vals:
            return i
    return 0


def read_sheet(xl: pd.ExcelFile, sheet_name: str):
    try:
        raw = xl.parse(sheet_name, header=None)

        header_row = find_header_row(raw)

        df = xl.parse(sheet_name, header=header_row)

        df.columns = [str(c).strip().upper() for c in df.columns]

        df = df.loc[:, ~df.columns.str.startswith("UNNAMED")]

        df = df.rename(columns={
            "TOTAL REJ": "TOTAL REJECTION",
            "TOTAL REJ.": "TOTAL REJECTION",
            "TOTAL    REJECTION": "TOTAL REJECTION",
            "TOTAL  REJECTION": "TOTAL REJECTION",

            "MACHINE NO": "M/C. NO.",
            "MACHINE NO.": "M/C. NO.",
            "MC NO": "M/C. NO.",
            "MC NO.": "M/C. NO."
        })

        df = df[df["COMPONENT"].notna()]
        df = df[df["COMPONENT"].astype(str).str.strip() != ""]
        df = df[df["COMPONENT"].astype(str).str.strip().str.upper() != "NAN"]

        df["DATE"] = pd.to_datetime(
            df["DATE"],
            dayfirst=True,
            errors="coerce"
        )

        df = df[df["DATE"].notna()]

        df["COMPONENT"] = (
            df["COMPONENT"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        if "CUSTOMER" not in df.columns:
            df["CUSTOMER"] = ""

        df["CUSTOMER"] = (
            df["CUSTOMER"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        if "M/C. NO." not in df.columns:
            df["M/C. NO."] = ""

        df["M/C. NO."] = (
            df["M/C. NO."]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        if "TOTAL PRODUCTION" not in df.columns:
            df["TOTAL PRODUCTION"] = 0

        df["TOTAL PRODUCTION"] = (
            pd.to_numeric(
                df["TOTAL PRODUCTION"],
                errors="coerce"
            ).fillna(0)
        )

        for col in DEFECT_COLS:
            if col in df.columns:
                df[col] = (
                    pd.to_numeric(df[col], errors="coerce")
                    .fillna(0)
                )
            else:
                df[col] = 0

        df["TOTAL REJECTION"] = df[DEFECT_COLS].sum(axis=1)

        df_raw = df.copy()

        group_cols = ["DATE", "COMPONENT", "CUSTOMER"]

        sum_cols = ["TOTAL PRODUCTION", "TOTAL REJECTION"] + DEFECT_COLS

        df = (
            df.groupby(group_cols, as_index=False)[sum_cols]
              .sum()
        )

        df["M/C. NO."] = ""

        total_prod = df["TOTAL PRODUCTION"]
        total_rej = df["TOTAL REJECTION"]

        df["REJ. %"] = (
            (total_rej / total_prod * 100)
            .round(2)
            .where(total_prod > 0, 0)
        )

        df["REJ. PPM"] = (
            (total_rej / total_prod * 1_000_000)
            .round(0)
            .where(total_prod > 0, 0)
        )

        df.reset_index(drop=True, inplace=True)

        return df, df_raw

    except Exception as e:
        st.error(sheet_name)
        st.exception(e)
        st.code(traceback.format_exc())
        return None, None


def get_month_data(xl: pd.ExcelFile, month_key: str, sheet_name: str):
    """
    Reads a PROCESS sheet for a month and caches it in session_state so
    repeated reruns within the same session don't re-parse the sheet.
    Returns (grouped_df, raw_df) same as read_sheet.
    """
    if month_key in st.session_state.month_data:
        return st.session_state.month_data[month_key]

    grouped, raw = read_sheet(xl, sheet_name)

    if grouped is not None:
        st.session_state.month_data[month_key] = (grouped, raw)

    return grouped, raw


def combine_months(xl: pd.ExcelFile, selected_keys: list, months: dict):
    """
    Concatenates the grouped and raw PROCESS data for each selected month.
    """
    grouped_list = []
    raw_list = []

    for key in selected_keys:
        sheet_name = months[key]
        grouped, raw = get_month_data(xl, key, sheet_name)

        if grouped is not None:
            g = grouped.copy()
            g["MONTH"] = key
            grouped_list.append(g)

        if raw is not None:
            r = raw.copy()
            r["MONTH"] = key
            raw_list.append(r)

    combined_grouped = (
        pd.concat(grouped_list, ignore_index=True) if grouped_list else None
    )
    combined_raw = (
        pd.concat(raw_list, ignore_index=True) if raw_list else None
    )

    return combined_grouped, combined_raw


def show_dashboard(data, process_raw, title):

    total_prod = data["TOTAL PRODUCTION"].sum()
    total_rej = data["TOTAL REJECTION"].sum()
    overall_rej_pct = round((total_rej / total_prod) * 100, 2) if total_prod > 0 else 0
    overall_ppm = round((total_rej / total_prod) * 1_000_000, 0) if total_prod > 0 else 0

    st.header(f"{title} — Dashboard")

    k1, k2, k3, k4 = st.columns(4)

    k1.metric("Total Production", f"{int(total_prod):,}")
    k2.metric("Total Rejection", f"{int(total_rej):,}")
    k3.metric("Overall REJ. %", f"{overall_rej_pct}%")
    k4.metric("Overall REJ. PPM", f"{int(overall_ppm):,}")

    st.divider()

    # ---------------- Customer ----------------

    st.subheader("Customer-wise Rejection")

    cust = data.groupby("CUSTOMER").agg(
        Total_Production=("TOTAL PRODUCTION", "sum"),
        Total_Rejection=("TOTAL REJECTION", "sum")
    ).reset_index()

    cust["REJ. %"] = (
        cust["Total_Rejection"] /
        cust["Total_Production"] * 100
    ).round(2).where(cust["Total_Production"] > 0, 0)

    cust["REJ. PPM"] = (
        cust["Total_Rejection"] /
        cust["Total_Production"] * 1_000_000
    ).round(0).where(cust["Total_Production"] > 0, 0)

    cust = cust.sort_values("REJ. %", ascending=False)

    st.dataframe(cust, use_container_width=True)

    st.divider()

    # ---------------- Defects ----------------

    st.subheader("Defect-wise Breakdown")

    defect_totals = data[DEFECT_COLS].sum().sort_values(ascending=False)

    defect_df = defect_totals.reset_index()
    defect_df.columns = ["Defect", "Count"]
    defect_df = defect_df[defect_df["Count"] > 0]

    fig = px.bar(defect_df, x="Defect", y="Count", text="Count")
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ---------------- Components ----------------

    st.subheader("Top 10 Components — Rejection")

    top_components = (
        data.groupby("COMPONENT")["TOTAL REJECTION"]
        .sum()
        .nlargest(10)
        .reset_index()
    )
    top_components.columns = ["Component", "Total Rejection"]

    fig = px.bar(
        top_components, x="Total Rejection", y="Component",
        orientation="h", text="Total Rejection"
    )
    fig.update_layout(yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ---------------- Machine ----------------

    st.subheader("Top 10 Machines — Rejection %")

    if process_raw is not None and not process_raw.empty:

        machine_data = process_raw[
            process_raw["M/C. NO."].notna() &
            (process_raw["M/C. NO."] != "")
        ].copy()

        if not machine_data.empty:

            machine = machine_data.groupby("M/C. NO.").agg(
                Total_Production=("TOTAL PRODUCTION", "sum"),
                Total_Rejection=("TOTAL REJECTION", "sum")
            ).reset_index()

            machine["REJ. %"] = (
                machine["Total_Rejection"] /
                machine["Total_Production"] * 100
            ).round(2).where(machine["Total_Production"] > 0, 0)

            machine = machine.sort_values("REJ. %", ascending=False).head(10)

            fig = px.bar(machine, x="M/C. NO.", y="REJ. %", text="REJ. %")
            st.plotly_chart(fig, use_container_width=True)

        else:
            st.info("No machine data available.")

    else:
        st.info("No machine data available.")

    st.divider()

    # ---------------- Weekly ----------------

    st.subheader("Weekly Rejection Summary")

    weekly_data = data[data["DATE"].notna()].copy()

    weekly_data["WEEK"] = (
        weekly_data["DATE"]
        .dt.to_period("W")
        .astype(str)
    )

    weekly = weekly_data.groupby("WEEK").agg(
        Total_Production=("TOTAL PRODUCTION", "sum"),
        Total_Rejection=("TOTAL REJECTION", "sum")
    ).reset_index()

    weekly["REJ. %"] = (
        weekly["Total_Rejection"] /
        weekly["Total_Production"] * 100
    ).round(2).where(weekly["Total_Production"] > 0, 0)

    fig = px.bar(weekly, x="WEEK", y="REJ. %", text="REJ. %")
    st.plotly_chart(fig, use_container_width=True)


def main():

    st.title("Rejection Analysis")

    with st.sidebar:
        st.header("Workbook")
        uploaded_file = st.file_uploader(
            "Upload workbook (PROCESS sheets)",
            type=["xlsx", "xlsm"]
        )

    if uploaded_file is None:
        st.info("Upload a workbook from the sidebar to get started.")
        return

    # Reset cached reads when a new/different file is uploaded, since
    # month_data is keyed only by month string and would otherwise
    # silently mix data across different uploads.
    file_id = f"{uploaded_file.name}-{uploaded_file.size}"
    if st.session_state.file_id != file_id:
        st.session_state.file_id = file_id
        st.session_state.month_data = {}

    file_bytes = uploaded_file.getvalue()
    xl = pd.ExcelFile(BytesIO(file_bytes))

    months = parse_available_months(xl)

    if not months:
        st.error("No PROCESS sheets detected in this workbook.")
        return

    with st.sidebar:
        st.divider()
        st.subheader("Detected Months")
        for month in months:
            st.write(f"**{month}**")

    st.divider()

    selected_keys = st.multiselect(
        "Select Month(s)",
        list(months.keys()),
        default=[list(months.keys())[-1]]
    )

    if not selected_keys:
        st.info("Select at least one month to view the dashboard.")
        return

    selected_keys = sorted(selected_keys, key=month_sort_key)

    if len(selected_keys) == 1:
        key = selected_keys[0]
        data, process_raw = get_month_data(xl, key, months[key])

        if data is not None:
            show_dashboard(data, process_raw, key)

    else:
        combined_data, combined_raw = combine_months(xl, selected_keys, months)

        if combined_data is None:
            st.error("Could not build combined data for the selected months.")
            return

        title = f"{selected_keys[0]} – {selected_keys[-1]} (Combined)"
        show_dashboard(combined_data, combined_raw, title)


main()