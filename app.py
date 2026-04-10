import streamlit as st
from pathlib import Path

from ticket_dashboard import render_ticket_dashboard
from social_media_dashboard import render_social_media_dashboard
from sales_dashboard import render_sales_dashboard
from operation_dashboard import render_operation_dashboard
from attendance_dashboard import render_attendance_dashboard


st.set_page_config(
    page_title="AllCall BI Dashboard",
    page_icon="📊",
    layout="wide"
)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "allcall_bi_data.xlsx"

NAVY = "#02013B"
NAVY_2 = "#060658"
ACCENT = "#0ACAF9"
WHITE = "#FFFFFF"
SOFT_WHITE = "#EAF2FF"


def inject_css():
    st.markdown(
        f"""
        <style>
        .stApp {{
            background: linear-gradient(180deg, {NAVY} 0%, {NAVY_2} 100%);
            color: {WHITE};
        }}
        header[data-testid="stHeader"] {{
            background: transparent !important;
            height: 0rem !important;
        }}
        [data-testid="stDecoration"] {{
            display: none !important;
        }}
        .block-container {{
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }}
        .main-title {{
            font-size: 32px;
            font-weight: 800;
            color: {WHITE};
            margin-bottom: 0.25rem;
        }}
        .sub-title {{
            font-size: 15px;
            color: {SOFT_WHITE};
            margin-bottom: 1.1rem;
        }}
        section[data-testid="stSidebar"] {{
            background: rgba(255,255,255,0.04) !important;
        }}
        section[data-testid="stSidebar"] * {{
            color: {WHITE} !important;
        }}
        .stTabs [data-baseweb="tab"] {{
            background: rgba(255,255,255,0.08);
            border-radius: 10px;
            color: {WHITE} !important;
        }}
        .stTabs [aria-selected="true"] {{
            background: {ACCENT} !important;
            color: #001b2a !important;
            font-weight: 700 !important;
        }}
        .stTabs [aria-selected="true"] * {{
            color: #001b2a !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def main():
    inject_css()

    st.markdown('<div class="main-title">AllCall BI Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">Excel-д суурилсан Ticket, Sales, Social media, Operation dashboard</div>',
        unsafe_allow_html=True
    )

    if not EXCEL_PATH.exists():
        st.error(f"Excel файл олдсонгүй: {EXCEL_PATH}")
        return

    menu = st.sidebar.radio(
        "Цэс",
        ["Ticket", "Sales", "Social media", "Operation"]
    )

    if menu == "Ticket":
        render_ticket_dashboard(str(EXCEL_PATH))
    elif menu == "Sales":
        render_sales_dashboard(str(EXCEL_PATH))
    elif menu == "Social media":
        render_social_media_dashboard(str(EXCEL_PATH))
    elif menu == "Operation":
        render_operation_dashboard(str(EXCEL_PATH))


if __name__ == "__main__":
    main()
