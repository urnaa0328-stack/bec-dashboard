import streamlit as st
import pandas as pd
from pathlib import Path
from attendance_dashboard import render_attendance_dashboard

BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_FILE = BASE_DIR / "allcall_bi_data.xlsx"

def render_operation_dashboard():
    st.markdown("## Operation Dashboard")

    tab1, tab2, tab3, tab4 = st.tabs([
        "AllCall Operation",
        "AllMed Operation",
        "Нэгдсэн Operation",
        "Attendance"
    ])

    with tab1:
        st.markdown("### AllCall Operation")
        # энд одоо байгаа allcall operation code-оо үлдээнэ

    with tab2:
        st.markdown("### AllMed Operation")
        # энд одоо байгаа allmed operation code-оо үлдээнэ

    with tab3:
        st.markdown("### Нэгдсэн Operation")
        # энд одоо байгаа нэгдсэн code-оо үлдээнэ

    with tab4:
        render_attendance_dashboard(str(EXCEL_FILE))
