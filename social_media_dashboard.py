import streamlit as st
import pandas as pd
import altair as alt


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def _safe_datetime(series: pd.Series) -> pd.Series:
    s1 = pd.to_datetime(series, errors="coerce")
    if s1.notna().sum() > 0:
        return s1
    return pd.to_datetime(series.astype(str).str.replace(".", "-", regex=False), errors="coerce")


def _fmt_money(x) -> str:
    try:
        return f"₮{int(round(float(x))):,}"
    except Exception:
        return "₮0"


def render_social_media_dashboard(excel_path: str):
    try:
        df = pd.read_excel(excel_path, sheet_name="Social media")
    except Exception as e:
        st.error(f"Social media sheet уншихад алдаа гарлаа: {e}")
        return

    df = _clean_columns(df)

    for c in [
        "Boost-н өдөр", "Пост үзсэн тоо", "Үзэгчид", "Чат эхлүүлсэн тоо",
        "Танилцуулга авсан", "Постын төсөв($ өдөрт)", "Нийт зарцуулсан ($)",
        "Нийт зарцуулсан (₮)", "Хоолой (₮)", "Adobe (₮)", "Hera (₮)"
    ]:
        if c not in df.columns:
            df[c] = 0
        df[c] = _to_num(df[c])

    for c in ["Постын агуулга"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()

    if "Эхэлсэн огноо" in df.columns:
        df["Эхэлсэн_dt"] = _safe_datetime(df["Эхэлсэн огноо"])
    else:
        df["Эхэлсэн_dt"] = pd.NaT

    if "Дууссан огноо" in df.columns:
        df["Дууссан_dt"] = _safe_datetime(df["Дууссан огноо"])
    else:
        df["Дууссан_dt"] = pd.NaT

    st.markdown("## Social Media Dashboard")

    search_text = st.text_input("Постын агуулгаар хайх", key="social_search").strip().lower()
    fdf = df.copy()

    if search_text:
        fdf = fdf[fdf["Постын агуулга"].str.lower().str.contains(search_text, na=False)]

    total_posts = len(fdf)
    total_views = int(fdf["Пост үзсэн тоо"].sum())
    total_chat = int(fdf["Чат эхлүүлсэн тоо"].sum())
    total_spent_mnt = float(fdf["Нийт зарцуулсан (₮)"].sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Нийт пост", f"{total_posts:,}")
    m2.metric("Пост үзсэн тоо", f"{total_views:,}")
    m3.metric("Чат эхлүүлсэн", f"{total_chat:,}")
    m4.metric("Нийт зарцуулсан", _fmt_money(total_spent_mnt))

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### Пост бүрийн үзэлт")
        post_views = (
            fdf[["Постын агуулга", "Пост үзсэн тоо"]]
            .sort_values("Пост үзсэн тоо", ascending=False)
            .copy()
        )
        if len(post_views) > 0:
            chart = alt.Chart(post_views).mark_bar().encode(
                x=alt.X("Пост үзсэн тоо:Q", title="Үзсэн тоо"),
                y=alt.Y("Постын агуулга:N", sort="-x", title="Пост")
            ).properties(height=380)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Өгөгдөл алга.")

    with c2:
        st.markdown("### Пост бүрийн чат эхлүүлэлт")
        post_chat = (
            fdf[["Постын агуулга", "Чат эхлүүлсэн тоо"]]
            .sort_values("Чат эхлүүлсэн тоо", ascending=False)
            .copy()
        )
        if len(post_chat) > 0:
            chart = alt.Chart(post_chat).mark_bar().encode(
                x=alt.X("Чат эхлүүлсэн тоо:Q", title="Чат эхлүүлсэн"),
                y=alt.Y("Постын агуулга:N", sort="-x", title="Пост")
            ).properties(height=380)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Өгөгдөл алга.")

    st.markdown("### Огнооны дагуух зарцуулалт")
    spent_df = (
        fdf.dropna(subset=["Эхэлсэн_dt"])
        .groupby("Эхэлсэн_dt", dropna=False)["Нийт зарцуулсан (₮)"]
        .sum()
        .reset_index()
        .sort_values("Эхэлсэн_dt")
    )

    if len(spent_df) > 0:
        line = alt.Chart(spent_df).mark_line(point=True).encode(
            x=alt.X("Эхэлсэн_dt:T", title="Эхэлсэн огноо"),
            y=alt.Y("Нийт зарцуулсан (₮):Q", title="Зарцуулалт")
        ).properties(height=320)
        st.altair_chart(line, use_container_width=True)
    else:
        st.info("Огнооны өгөгдөл алга.")

    st.markdown("### Дэлгэрэнгүй хүснэгт")
    show_df = fdf.copy()
    for c in ["Нийт зарцуулсан (₮)", "Хоолой (₮)", "Adobe (₮)", "Hera (₮)"]:
        if c in show_df.columns:
            show_df[c] = show_df[c].apply(_fmt_money)

    st.dataframe(show_df, use_container_width=True, hide_index=True)
    