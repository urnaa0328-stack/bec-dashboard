import streamlit as st
import pandas as pd

from modules.helpers import (
    try_load_sheet,
    find_column,
    infer_date_column,
    safe_to_datetime,
    safe_to_numeric,
    metric_card,
    highlight_title,
    render_bar_chart,
    render_line_chart,
)


def _load_overview_data():
    data = {}

    # Ticket
    try:
        df_ticket, ticket_sheet = try_load_sheet(["Ticket AllMed", "Ticket AllCall", "Ticket"])
        date_col = find_column(df_ticket, ["date", "ognoo", "огноо", "created_date", "created at"])
        if date_col is None:
            date_col = infer_date_column(df_ticket)

        count_col = find_column(df_ticket, ["count", "too", "тоо", "ticket_count"])
        if count_col is None:
            df_ticket["count"] = 1
            count_col = "count"

        status_col = find_column(df_ticket, ["status", "tuluv", "төлөв", "state"])
        if status_col is None:
            df_ticket["status"] = "Тодорхойгүй"
            status_col = "status"

        data["ticket"] = pd.DataFrame({
            "date": safe_to_datetime(df_ticket[date_col]) if date_col else pd.NaT,
            "count": safe_to_numeric(df_ticket[count_col], 0),
            "status": df_ticket[status_col].astype(str)
        })
        data["ticket_sheet"] = ticket_sheet
    except Exception:
        data["ticket"] = pd.DataFrame(columns=["date", "count", "status"])
        data["ticket_sheet"] = "-"

    # Sales
    try:
        df_sales, sales_sheet = try_load_sheet(["AllCall Sales", "Sales"])
        date_col = find_column(df_sales, ["date", "ognoo", "огноо", "sale_date", "created_date"])
        if date_col is None:
            date_col = infer_date_column(df_sales)

        amount_col = find_column(df_sales, ["amount", "dun", "дүн", "total", "revenue", "borluulalt"])
        if amount_col is None:
            df_sales["amount"] = 0
            amount_col = "amount"

        data["sales"] = pd.DataFrame({
            "date": safe_to_datetime(df_sales[date_col]) if date_col else pd.NaT,
            "amount": safe_to_numeric(df_sales[amount_col], 0)
        })
        data["sales_sheet"] = sales_sheet
    except Exception:
        data["sales"] = pd.DataFrame(columns=["date", "amount"])
        data["sales_sheet"] = "-"

    # Social
    try:
        df_social, social_sheet = try_load_sheet(["Social media", "Social Media", "Social"])
        channel_col = find_column(df_social, ["channel", "source", "platform", "suvag"])
        if channel_col is None:
            df_social["channel"] = "Тодорхойгүй"
            channel_col = "channel"

        engagement_col = find_column(df_social, ["engagement", "reach", "interaction", "like", "result"])
        if engagement_col is None:
            df_social["engagement"] = 0
            engagement_col = "engagement"

        data["social"] = pd.DataFrame({
            "channel": df_social[channel_col].astype(str),
            "engagement": safe_to_numeric(df_social[engagement_col], 0)
        })
        data["social_sheet"] = social_sheet
    except Exception:
        data["social"] = pd.DataFrame(columns=["channel", "engagement"])
        data["social_sheet"] = "-"

    return data


def render_overview_dashboard():
    st.title("Overview Dashboard")
    st.caption("Нэгдсэн тойм үзүүлэлт")

    data = _load_overview_data()

    ticket_total = int(data["ticket"]["count"].sum()) if not data["ticket"].empty else 0
    sales_total = float(data["sales"]["amount"].sum()) if not data["sales"].empty else 0
    social_total = float(data["social"]["engagement"].sum()) if not data["social"].empty else 0

    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("Нийт Ticket", f"{ticket_total:,}", f"Sheet: {data['ticket_sheet']}")
    with c2:
        metric_card("Нийт Борлуулалт", f"{sales_total:,.0f}", f"Sheet: {data['sales_sheet']}")
    with c3:
        metric_card("Нийт Social Engagement", f"{social_total:,.0f}", f"Sheet: {data['social_sheet']}")

    col1, col2 = st.columns(2)

    with col1:
        highlight_title("Ticket төлөвийн хураангуй")
        if not data["ticket"].empty:
            by_status = data["ticket"].groupby("status", as_index=False)["count"].sum()
            render_bar_chart(by_status, "status", "count")
        else:
            st.info("Ticket data хоосон байна.")

    with col2:
        highlight_title("Social сувгийн хураангуй")
        if not data["social"].empty:
            by_channel = data["social"].groupby("channel", as_index=False)["engagement"].sum()
            render_bar_chart(by_channel, "channel", "engagement")
        else:
            st.info("Social data хоосон байна.")

    highlight_title("Борлуулалтын trend")
    if not data["sales"].empty:
        trend = data["sales"].dropna(subset=["date"]).copy()
        if not trend.empty:
            trend["day"] = trend["date"].dt.floor("D")
            trend = trend.groupby("day", as_index=False)["amount"].sum()
            render_line_chart(trend, "day", "amount")
        else:
            st.info("Sales date data хоосон байна.")
    else:
        st.info("Sales data хоосон байна.")