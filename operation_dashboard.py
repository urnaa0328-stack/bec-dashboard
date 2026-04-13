import streamlit as st
import pandas as pd
from attendance_dashboard import render_attendance_dashboard

def render_operation_dashboard(excel_path: str):
    st.markdown("## Operation Dashboard")

    tab1, tab2, tab3, tab4 = st.tabs([
        "AllCall Operation",
        "AllMed Operation",
        "Нэгдсэн Operation",
        "Attendance"
    ])

    with tab1:
        st.markdown("### AllCall Operation")
        st.info("Энд AllCall operation code орно.")

    with tab2:
        st.markdown("### AllMed Operation")
        st.info("Энд AllMed operation code орно.")

    with tab3:
        st.markdown("### Нэгдсэн Operation")
        st.info("Энд нэгдсэн operation code орно.")

    with tab4:
        render_attendance_dashboard(excel_path)
