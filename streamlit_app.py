"""Streamlit front-end for the Wipro BGV report pipeline in app.py.

Runs the same export -> process -> pivot steps as `python app.py`, with
progress shown in the browser and a download button for the resulting
workbook. Credentials are not entered here - app.py already resolves them
from MIS_USERNAME/MIS_PASSWORD env vars (falling back to its hardcoded
defaults), and this app reuses that as-is.
"""

import datetime
import os

import streamlit as st

import app

st.set_page_config(page_title="Wipro BGV Report", page_icon="📊", layout="centered")
st.title("Wipro BGV Report")
st.caption("Exports the Client Wise All Cases - Bridge report and builds the processed workbook + Summary Sheet pivots.")

existing_path = os.path.join(app.ROOT_DIR, "Client Wise All Cases - Bridge (Processed).xlsx")

if os.path.exists(existing_path):
    modified = datetime.datetime.fromtimestamp(os.path.getmtime(existing_path))
    st.info(f"Existing report last generated: {modified:%Y-%m-%d %H:%M:%S}")
    with open(existing_path, "rb") as f:
        st.download_button(
            "Download existing report",
            data=f.read(),
            file_name=os.path.basename(existing_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("No report has been generated yet.")

st.divider()

if st.button("Run report now", type="primary"):
    try:
        with st.status("Running pipeline...", expanded=True) as status:
            st.write("Logging in and exporting CSV from the MIS Query Browser...")
            csv_path = app.export_csv()
            st.write(f"Saved: `{csv_path}`")

            st.write("Filtering to the current month and building the processed workbook...")
            out_path = app.process_csv_to_xlsx(csv_path)
            st.write(f"Saved: `{out_path}`")

            st.write("Building Summary Sheet pivot tables (drives Excel via COM)...")
            app.build_summary_pivots(out_path)
            st.write("Summary Sheet rebuilt.")

            status.update(label="Done", state="complete")

        st.success("Report generated successfully.")
        with open(out_path, "rb") as f:
            st.download_button(
                "Download report",
                data=f.read(),
                file_name=os.path.basename(out_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as exc:
        st.error(f"Pipeline failed: {exc}")
        st.exception(exc)
