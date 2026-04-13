import streamlit as st
import pandas as pd
from pathlib import Path
from attendance_dashboard import render_attendance_dashboard

BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_FILE = BASE_DIR / "allcall_bi_data.xlsx"
SHEET_NAME = "Operation"

WHITE = "#C9CED6"
CARD_BG = "rgba(255,255,255,0.10)"
CARD_BORDER = "rgba(255,255,255,0.16)"


def metric_card(title, value):
    st.markdown(
        f"""
        <div style="
            background:{CARD_BG};
            border:1px solid {CARD_BORDER};
            border-radius:18px;
            padding:16px 18px;
            min-height:100px;
        ">
            <div style="font-size:14px;color:{WHITE};opacity:.85;">{title}</div>
            <div style="font-size:34px;font-weight:700;color:white;margin-top:10px;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


@st.cache_data(show_spinner=False)
def load_operation_data():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel файл олдсонгүй: {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]

    for col in ["Эхлэх огноо", "Дуусах огноо"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if "Хугацаа" in df.columns:
        df["Хугацаа"] = pd.to_numeric(df["Хугацаа"], errors="coerce")

    if {"Эхлэх огноо", "Дуусах огноо", "Хугацаа"}.issubset(df.columns):
        duration_missing = (
            df["Хугацаа"].isna() &
            df["Эхлэх огноо"].notna() &
            df["Дуусах огноо"].notna()
        )
        df.loc[duration_missing, "Хугацаа"] = (
            df.loc[duration_missing, "Дуусах огноо"] -
            df.loc[duration_missing, "Эхлэх огноо"]
        ).dt.days

    return df


def render_operation_dashboard():
    st.markdown("## Operation Dashboard")

    tab1, tab2, tab3, tab4 = st.tabs([
        "Operation",
        "Нэгдсэн summary",
        "Жагсаалт",
        "Attendance"
    ])

    try:
        df = load_operation_data().copy()
    except Exception as e:
        st.error(f"Operation sheet уншихад алдаа гарлаа: {e}")
        return

    with tab1:
        f1, f2, f3, f4 = st.columns(4)

        with f1:
            if "Ажлын төрөл" in df.columns:
                work_types = ["Бүгд"] + sorted(df["Ажлын төрөл"].dropna().astype(str).unique().tolist())
            else:
                work_types = ["Бүгд"]
            selected_type = st.selectbox("Ажлын төрөл", work_types)

        with f2:
            if "Хариуцагч" in df.columns:
                owners = ["Бүгд"] + sorted(df["Хариуцагч"].dropna().astype(str).unique().tolist())
            else:
                owners = ["Бүгд"]
            selected_owner = st.selectbox("Хариуцагч", owners)

        with f3:
            status_col = "Явц" if "Явц" in df.columns else None
            if status_col:
                statuses = ["Бүгд"] + sorted(df[status_col].dropna().astype(str).unique().tolist())
            else:
                statuses = ["Бүгд"]
            selected_status = st.selectbox("Явц", statuses)

        with f4:
            search_text = st.text_input("Төслийн нэр хайх")

        if selected_type != "Бүгд" and "Ажлын төрөл" in df.columns:
            df = df[df["Ажлын төрөл"].astype(str) == selected_type]

        if selected_owner != "Бүгд" and "Хариуцагч" in df.columns:
            df = df[df["Хариуцагч"].astype(str) == selected_owner]

        if selected_status != "Бүгд" and status_col:
            df = df[df[status_col].astype(str) == selected_status]

        if search_text and "Төслийн нэр" in df.columns:
            df = df[df["Төслийн нэр"].astype(str).str.contains(search_text, case=False, na=False)]

        total = len(df)
        done = len(df[df[status_col].astype(str).str.contains("хийгдсэн|done|complete", case=False, na=False)]) if status_col else 0
        active = len(df[df[status_col].astype(str).str.contains("явж|progress|ongoing", case=False, na=False)]) if status_col else 0
        extended = len(df[df[status_col].astype(str).str.contains("сунгасан|extend", case=False, na=False)]) if status_col else 0

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            metric_card("Нийт ажил", total)
        with c2:
            metric_card("Хийгдсэн", done)
        with c3:
            metric_card("Хийгдэж буй", active)
        with c4:
            metric_card("Хугацаа сунгасан", extended)

    with tab2:
        st.dataframe(df.describe(include="all").transpose(), use_container_width=True)

    with tab3:
        preferred_cols = [
            "Ажлын төрөл", "Төслийн нэр", "Эхлэх огноо", "Дуусах огноо",
            "Хугацаа", "Хариуцагч", "Дэмжлэг", "Явцын тайлбар", "Явц"
        ]
        show_cols = [c for c in preferred_cols if c in df.columns]
        st.dataframe(df[show_cols] if show_cols else df, use_container_width=True, hide_index=True)

    with tab4:
        render_attendance_dashboard()
