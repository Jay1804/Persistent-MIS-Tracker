from pathlib import Path

import pandas as pd
import streamlit as st

from app import generate_report

st.set_page_config(page_title="Persistent MIS Tracker Report", page_icon="📊", layout="wide")

st.title("Persistent MIS Tracker Report")
st.caption("Pulls the latest Persistent Systems Limited tracker data from the AuthBridge MIS Query Browser and builds the report workbook.")

if st.button("Generate Report", type="primary"):
    with st.spinner("Logging in and pulling data from mis.authbridge.com..."):
        try:
            report_path = generate_report(headless=True)
        except Exception as exc:
            st.error("Report generation failed:")
            st.exception(exc)
        else:
            st.session_state["report_path"] = str(report_path)
            st.success(f"Report generated: {Path(report_path).name}")

if "report_path" in st.session_state:
    report_path = Path(st.session_state["report_path"])

    if report_path.exists():
        with open(report_path, "rb") as f:
            st.download_button(
                label="Download Report",
                data=f.read(),
                file_name=report_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        data_df = pd.read_excel(report_path, sheet_name="Data Tab")
        status_summary = pd.read_excel(report_path, sheet_name="Summary", usecols="A:B", skiprows=2)
        severity_summary = pd.read_excel(report_path, sheet_name="Summary", usecols="E:F", skiprows=2)

        st.subheader("Summary")
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Overall Case wise Status**")
            st.dataframe(status_summary, hide_index=True)
        with col2:
            st.write("**CBC Severity**")
            st.dataframe(severity_summary, hide_index=True)

        st.subheader("Data Tab preview")
        st.dataframe(data_df.head(100), hide_index=True)
    else:
        st.warning("Previously generated report file no longer exists on disk.")
