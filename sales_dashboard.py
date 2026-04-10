import streamlit as st
import pandas as pd
import altair as alt


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def _fmt_money(x) -> str:
    try:
        return f"₮{int(round(float(x))):,}"
    except Exception:
        return "₮0"


def _prepare_sales(df: pd.DataFrame, system_name: str) -> pd.DataFrame:
    df = _clean_columns(df)

    expected_cols = [
        "Холбогдсон байгууллага",
        "Суваг",
        "Харилцагч хариу өгөөгүй",
        "Системийн танилцуулга өгсөн",
        "Харилцаж байгаа магадлал өндөр",
        "Шийдсэн",
    ]

    for col in expected_cols:
        if col not in df.columns:
            df[col] = "" if col in ["Холбогдсон байгууллага", "Суваг"] else 0

    df["Холбогдсон байгууллага"] = df["Холбогдсон байгууллага"].fillna("").astype(str).str.strip()
    df["Суваг"] = df["Суваг"].fillna("").astype(str).str.strip()

    money_cols = [
        "Харилцагч хариу өгөөгүй",
        "Системийн танилцуулга өгсөн",
        "Харилцаж байгаа магадлал өндөр",
        "Шийдсэн",
    ]

    for col in money_cols:
        df[col] = _to_num(df[col])

    df["Нийт дүн"] = df[money_cols].sum(axis=1)
    df["Систем"] = system_name

    df = df[
        (df["Холбогдсон байгууллага"] != "") |
        (df["Суваг"] != "") |
        (df["Нийт дүн"] != 0)
    ].copy()

    return df


def _sales_summary(df: pd.DataFrame) -> dict:
    return {
        "lead_count": len(df),
        "not_answered": float(df["Харилцагч хариу өгөөгүй"].sum()),
        "intro": float(df["Системийн танилцуулга өгсөн"].sum()),
        "high_prob": float(df["Харилцаж байгаа магадлал өндөр"].sum()),
        "closed": float(df["Шийдсэн"].sum()),
        "total": float(df["Нийт дүн"].sum()),
    }


def _render_metrics(summary: dict):
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Нийт мөр", f"{summary['lead_count']:,}")
    c2.metric("Хариу өгөөгүй", _fmt_money(summary["not_answered"]))
    c3.metric("Танилцуулга өгсөн", _fmt_money(summary["intro"]))
    c4.metric("Магадлал өндөр", _fmt_money(summary["high_prob"]))
    c5.metric("Шийдсэн", _fmt_money(summary["closed"]))

    st.markdown(
        f"""
        <div class="card">
            <div class="card-title">Нийт боломжит дүн</div>
            <div class="card-value">{_fmt_money(summary["total"])}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def _render_sales_table(df: pd.DataFrame, key_prefix: str):
    c1, c2 = st.columns(2)

    with c1:
        channels = ["Бүгд"] + sorted([x for x in df["Суваг"].dropna().astype(str).unique().tolist() if x.strip()])
        selected_channel = st.selectbox("Суваг", channels, key=f"{key_prefix}_channel")

    with c2:
        search_text = st.text_input("Байгууллага хайх", key=f"{key_prefix}_search").strip().lower()

    fdf = df.copy()

    if selected_channel != "Бүгд":
        fdf = fdf[fdf["Суваг"] == selected_channel]

    if search_text:
        fdf = fdf[fdf["Холбогдсон байгууллага"].str.lower().str.contains(search_text, na=False)]

    show_df = fdf.copy()
    money_cols = [
        "Харилцагч хариу өгөөгүй",
        "Системийн танилцуулга өгсөн",
        "Харилцаж байгаа магадлал өндөр",
        "Шийдсэн",
        "Нийт дүн",
    ]
    for col in money_cols:
        show_df[col] = show_df[col].apply(_fmt_money)

    cols = [
        "Холбогдсон байгууллага",
        "Суваг",
        "Харилцагч хариу өгөөгүй",
        "Системийн танилцуулга өгсөн",
        "Харилцаж байгаа магадлал өндөр",
        "Шийдсэн",
        "Нийт дүн",
    ]

    st.dataframe(show_df[cols], use_container_width=True, hide_index=True)

    return fdf


def _render_sales_charts(df: pd.DataFrame):
    st.markdown("### Sales шатлалын дүн")

    stage_df = pd.DataFrame({
        "Шат": [
            "Харилцагч хариу өгөөгүй",
            "Системийн танилцуулга өгсөн",
            "Харилцаж байгаа магадлал өндөр",
            "Шийдсэн",
        ],
        "Дүн": [
            df["Харилцагч хариу өгөөгүй"].sum(),
            df["Системийн танилцуулга өгсөн"].sum(),
            df["Харилцаж байгаа магадлал өндөр"].sum(),
            df["Шийдсэн"].sum(),
        ]
    })

    chart = alt.Chart(stage_df).mark_bar().encode(
        x=alt.X("Шат:N", title="Sales үе шат"),
        y=alt.Y("Дүн:Q", title="Дүн"),
        tooltip=["Шат", "Дүн"]
    ).properties(height=350)

    st.altair_chart(chart, use_container_width=True)

    st.markdown("### Суваг тус бүрийн нийт дүн")
    channel_df = (
        df.groupby("Суваг", dropna=False)["Нийт дүн"]
        .sum()
        .reset_index()
        .sort_values("Нийт дүн", ascending=False)
    )
    if len(channel_df) > 0:
        chart2 = alt.Chart(channel_df).mark_bar().encode(
            x=alt.X("Нийт дүн:Q", title="Нийт дүн"),
            y=alt.Y("Суваг:N", sort="-x", title="Суваг"),
            tooltip=["Суваг", "Нийт дүн"]
        ).properties(height=320)
        st.altair_chart(chart2, use_container_width=True)
    else:
        st.info("Сувгийн өгөгдөл алга.")


def render_sales_dashboard(excel_path: str):
    try:
        allcall_raw = pd.read_excel(excel_path, sheet_name="AllCall Sales")
        allmed_raw = pd.read_excel(excel_path, sheet_name="AllMed Sales")
    except Exception as e:
        st.error(f"Sales sheet уншихад алдаа гарлаа: {e}")
        return

    allcall_df = _prepare_sales(allcall_raw, "AllCall")
    allmed_df = _prepare_sales(allmed_raw, "AllMed")
    combined_df = pd.concat([allcall_df, allmed_df], ignore_index=True)

    st.markdown("## Sales Dashboard")

    tab1, tab2, tab3 = st.tabs(["AllCall Sales", "AllMed Sales", "Нэгдсэн Sales"])

    with tab1:
        summary = _sales_summary(allcall_df)
        _render_metrics(summary)
        filtered_df = _render_sales_table(allcall_df, "allcall_sales")
        _render_sales_charts(filtered_df)

    with tab2:
        summary = _sales_summary(allmed_df)
        _render_metrics(summary)
        filtered_df = _render_sales_table(allmed_df, "allmed_sales")
        _render_sales_charts(filtered_df)

    with tab3:
        summary = _sales_summary(combined_df)
        _render_metrics(summary)

        c1, c2, c3 = st.columns(3)
        with c1:
            system = st.selectbox("Систем", ["Бүгд", "AllCall", "AllMed"], key="sales_system")
        with c2:
            channels = ["Бүгд"] + sorted([x for x in combined_df["Суваг"].dropna().astype(str).unique().tolist() if x.strip()])
            channel = st.selectbox("Суваг", channels, key="sales_channel_all")
        with c3:
            search_text = st.text_input("Байгууллага хайх", key="sales_search_all").strip().lower()

        fdf = combined_df.copy()

        if system != "Бүгд":
            fdf = fdf[fdf["Систем"] == system]
        if channel != "Бүгд":
            fdf = fdf[fdf["Суваг"] == channel]
        if search_text:
            fdf = fdf[fdf["Холбогдсон байгууллага"].str.lower().str.contains(search_text, na=False)]

        show_df = fdf.copy()
        for col in [
            "Харилцагч хариу өгөөгүй",
            "Системийн танилцуулга өгсөн",
            "Харилцаж байгаа магадлал өндөр",
            "Шийдсэн",
            "Нийт дүн",
        ]:
            show_df[col] = show_df[col].apply(_fmt_money)

        cols = [
            "Систем",
            "Холбогдсон байгууллага",
            "Суваг",
            "Харилцагч хариу өгөөгүй",
            "Системийн танилцуулга өгсөн",
            "Харилцаж байгаа магадлал өндөр",
            "Шийдсэн",
            "Нийт дүн",
        ]
        st.dataframe(show_df[cols], use_container_width=True, hide_index=True)

        _render_sales_charts(fdf)