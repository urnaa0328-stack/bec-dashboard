import streamlit as st
import pandas as pd
import altair as alt


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _safe_datetime(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype="datetime64[ns]")

    s = pd.to_datetime(series, errors="coerce")
    if s.notna().sum() > 0:
        return s

    return pd.to_datetime(
        series.astype(str)
        .str.strip()
        .str.replace(".", "-", regex=False)
        .str.replace("/", "-", regex=False),
        errors="coerce"
    )


def _ensure_str_cols(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
    return df


def _normalize_sheet_name(name: str) -> str:
    return (
        str(name)
        .strip()
        .lower()
        .replace("_", "")
        .replace("-", "")
        .replace(" ", "")
    )


def _find_sheet_name(excel_path: str, candidates: list[str]) -> tuple[str | None, list[str]]:
    try:
        xls = pd.ExcelFile(excel_path)
        actual_sheets = xls.sheet_names
    except Exception as e:
        st.error(f"Excel файл нээхэд алдаа гарлаа: {e}")
        return None, []

    normalized_map = {_normalize_sheet_name(s): s for s in actual_sheets}

    for candidate in candidates:
        key = _normalize_sheet_name(candidate)
        if key in normalized_map:
            return normalized_map[key], actual_sheets

    return None, actual_sheets


def _read_excel_sheet_auto(excel_path: str, candidates: list[str], label: str) -> pd.DataFrame | None:
    matched_sheet, actual_sheets = _find_sheet_name(excel_path, candidates)

    if matched_sheet is None:
        st.error(
            f"{label} sheet олдсонгүй. Боломжит sheet нэрс: {', '.join(actual_sheets) if actual_sheets else 'sheet алга'}"
        )
        return None

    try:
        df = pd.read_excel(excel_path, sheet_name=matched_sheet)
        df = _clean_columns(df)
        st.caption(f"{label}: `{matched_sheet}` sheet уншигдлаа")
        return df
    except Exception as e:
        st.error(f"{label} sheet уншихад алдаа гарлаа: {e}")
        return None


def _render_empty(msg="Өгөгдөл алга."):
    st.info(msg)


def _contains_any(series: pd.Series, keywords: list[str]) -> pd.Series:
    pattern = "|".join(keywords)
    return series.astype(str).str.contains(pattern, case=False, na=False)


# =========================================================
# ALLCALL
# =========================================================
def render_allcall_ticket_dashboard(excel_path: str):
    df = _read_excel_sheet_auto(
        excel_path,
        candidates=[
            "AllCall Ticket",
            "Ticket AllCall",
            "AllCallTicket",
            "AllCall",
            "Ticket_AllCall",
            "allcall ticket",
            "ticket allcall",
        ],
        label="AllCall Ticket",
    )
    if df is None:
        return

    if "Unnamed: 3" in df.columns and "Дэд төрөл" not in df.columns:
        df = df.rename(columns={"Unnamed: 3": "Дэд төрөл"})

    if "Огноо" in df.columns:
        df["Огноо_dt"] = _safe_datetime(df["Огноо"])
        df["Огноо_өдөр"] = df["Огноо_dt"].dt.date
    else:
        df["Огноо_dt"] = pd.NaT
        df["Огноо_өдөр"] = pd.NaT

    df = _ensure_str_cols(
        df,
        ["Суваг", "Төрөл", "Дэд төрөл", "Төлөв", "Оператор", "Нэр", "ААН", "Санал, гомдол"]
    )

    st.markdown("## AllCall Ticket Dashboard")

    min_date = df["Огноо_dt"].min()
    max_date = df["Огноо_dt"].max()

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        selected_channel = st.selectbox(
            "Суваг",
            ["Бүгд"] + sorted([x for x in df["Суваг"].unique().tolist() if x]),
            key="allcall_ticket_channel",
        )

    with c2:
        selected_status = st.selectbox(
            "Төлөв",
            ["Бүгд"] + sorted([x for x in df["Төлөв"].unique().tolist() if x]),
            key="allcall_ticket_status",
        )

    with c3:
        selected_type = st.selectbox(
            "Төрөл",
            ["Бүгд"] + sorted([x for x in df["Төрөл"].unique().tolist() if x]),
            key="allcall_ticket_type",
        )

    with c4:
        search_text = st.text_input(
            "Хайх (Нэр / ААН / Санал)",
            key="allcall_ticket_search",
        ).strip().lower()

    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.date_input(
            "Огнооны шүүлтүүр",
            value=(min_date.date(), max_date.date()),
            key="allcall_ticket_date_range",
        )
    else:
        date_range = None

    fdf = df.copy()

    if selected_channel != "Бүгд":
        fdf = fdf[fdf["Суваг"] == selected_channel]

    if selected_status != "Бүгд":
        fdf = fdf[fdf["Төлөв"] == selected_status]

    if selected_type != "Бүгд":
        fdf = fdf[fdf["Төрөл"] == selected_type]

    if search_text:
        mask = (
            fdf["Нэр"].str.lower().str.contains(search_text, na=False)
            | fdf["ААН"].str.lower().str.contains(search_text, na=False)
            | fdf["Санал, гомдол"].str.lower().str.contains(search_text, na=False)
        )
        fdf = fdf[mask]

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        fdf = fdf[
            (fdf["Огноо_dt"].dt.date >= start_date) &
            (fdf["Огноо_dt"].dt.date <= end_date)
        ]

    total_count = len(fdf)
    solved_count = len(fdf[_contains_any(fdf["Төлөв"], ["шийдвэр", "дуус", "хаагд", "closed", "resolved"])])
    channel_count = fdf["Суваг"].replace("", pd.NA).dropna().nunique()
    operator_count = fdf["Оператор"].replace("", pd.NA).dropna().nunique()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Нийт ticket", f"{total_count:,}")
    m2.metric("Шийдвэрлэсэн", f"{solved_count:,}")
    m3.metric("Сувгийн тоо", f"{channel_count:,}")
    m4.metric("Операторын тоо", f"{operator_count:,}")

    c5, c6 = st.columns(2)

    with c5:
        st.markdown("### Суваг тус бүр")
        ch_df = (
            fdf.groupby("Суваг", dropna=False)
            .size()
            .reset_index(name="Тоо")
            .sort_values("Тоо", ascending=False)
        )
        ch_df = ch_df[ch_df["Суваг"].astype(str).str.strip() != ""]
        if len(ch_df) > 0:
            chart = alt.Chart(ch_df).mark_bar().encode(
                x=alt.X("Тоо:Q"),
                y=alt.Y("Суваг:N", sort="-x"),
                tooltip=["Суваг", "Тоо"]
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            _render_empty()

    with c6:
        st.markdown("### Төлөв тус бүр")
        st_df = (
            fdf.groupby("Төлөв", dropna=False)
            .size()
            .reset_index(name="Тоо")
            .sort_values("Тоо", ascending=False)
        )
        st_df = st_df[st_df["Төлөв"].astype(str).str.strip() != ""]
        if len(st_df) > 0:
            chart = alt.Chart(st_df).mark_bar().encode(
                x=alt.X("Тоо:Q"),
                y=alt.Y("Төлөв:N", sort="-x"),
                tooltip=["Төлөв", "Тоо"]
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            _render_empty()

    st.markdown("### Өдөр тутмын ticket урсгал")
    daily = (
        fdf.dropna(subset=["Огноо_dt"])
        .groupby("Огноо_өдөр")
        .size()
        .reset_index(name="Тоо")
        .sort_values("Огноо_өдөр")
    )
    if len(daily) > 0:
        chart = alt.Chart(daily).mark_line(point=True).encode(
            x=alt.X("Огноо_өдөр:T", title="Огноо"),
            y=alt.Y("Тоо:Q", title="Тоо"),
            tooltip=["Огноо_өдөр", "Тоо"]
        ).properties(height=320)
        st.altair_chart(chart, use_container_width=True)
    else:
        _render_empty("Огнооны өгөгдөл алга.")

    st.markdown("### Дэлгэрэнгүй хүснэгт")
    show_cols = [c for c in [
        "Суваг", "Огноо", "Төрөл", "Дэд төрөл", "Утас", "И-мэйл",
        "ААН", "Нэр", "Регистр", "Санал, гомдол", "Төлөв",
        "Огноо.1", "Хугацаа", "Оператор"
    ] if c in fdf.columns]

    if show_cols:
        st.dataframe(fdf[show_cols], use_container_width=True, hide_index=True)
    else:
        _render_empty("Хүснэгтэнд харуулах багана олдсонгүй.")


# =========================================================
# ALLMED
# =========================================================
def render_allmed_ticket_dashboard(excel_path: str):
    df = _read_excel_sheet_auto(
        excel_path,
        candidates=[
            "AllMed Ticket",
            "Ticket AllMed",
            "AllMedTicket",
            "AllMed",
            "Ticket_AllMed",
            "allmed ticket",
            "ticket allmed",
        ],
        label="AllMed Ticket",
    )
    if df is None:
        return

    if "Огноо" in df.columns:
        df["Огноо_dt"] = _safe_datetime(df["Огноо"])
        df["Огноо_өдөр"] = df["Огноо_dt"].dt.date
    else:
        df["Огноо_dt"] = pd.NaT
        df["Огноо_өдөр"] = pd.NaT

    if "Дууссан огноо" in df.columns:
        df["Дууссан_огноо_dt"] = _safe_datetime(df["Дууссан огноо"])
    else:
        df["Дууссан_огноо_dt"] = pd.NaT

    df = _ensure_str_cols(df, ["Явц", "Ажлын нэр", "Хариуцагчийн тайлбар", "Хариуцагч"])

    st.markdown("## AllMed Ticket Dashboard")

    min_date = df["Огноо_dt"].min()
    max_date = df["Огноо_dt"].max()

    c1, c2, c3 = st.columns(3)

    with c1:
        selected_progress = st.selectbox(
            "Явц",
            ["Бүгд"] + sorted([x for x in df["Явц"].unique().tolist() if x]),
            key="allmed_ticket_progress",
        )

    with c2:
        selected_owner = st.selectbox(
            "Хариуцагч",
            ["Бүгд"] + sorted([x for x in df["Хариуцагч"].unique().tolist() if x and x != "-"]),
            key="allmed_ticket_owner",
        )

    with c3:
        search_text = st.text_input(
            "Хайх (Ажлын нэр / Тайлбар)",
            key="allmed_ticket_search",
        ).strip().lower()

    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.date_input(
            "Огнооны шүүлтүүр",
            value=(min_date.date(), max_date.date()),
            key="allmed_ticket_date_range",
        )
    else:
        date_range = None

    fdf = df.copy()

    if selected_progress != "Бүгд":
        fdf = fdf[fdf["Явц"] == selected_progress]

    if selected_owner != "Бүгд":
        fdf = fdf[fdf["Хариуцагч"] == selected_owner]

    if search_text:
        mask = (
            fdf["Ажлын нэр"].str.lower().str.contains(search_text, na=False)
            | fdf["Хариуцагчийн тайлбар"].str.lower().str.contains(search_text, na=False)
        )
        fdf = fdf[mask]

    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        fdf = fdf[
            (fdf["Огноо_dt"].dt.date >= start_date) &
            (fdf["Огноо_dt"].dt.date <= end_date)
        ]

    total_count = len(fdf)
    done_count = len(fdf[_contains_any(fdf["Явц"], ["дуус", "done", "completed", "closed"])])
    owner_count = fdf["Хариуцагч"].replace(["", "-"], pd.NA).dropna().nunique()

    m1, m2, m3 = st.columns(3)
    m1.metric("Нийт ажил", f"{total_count:,}")
    m2.metric("Дууссан", f"{done_count:,}")
    m3.metric("Хариуцагчийн тоо", f"{owner_count:,}")

    c4, c5 = st.columns(2)

    with c4:
        st.markdown("### Явцын ангилал")
        prog_df = (
            fdf.groupby("Явц", dropna=False)
            .size()
            .reset_index(name="Тоо")
            .sort_values("Тоо", ascending=False)
        )
        prog_df = prog_df[prog_df["Явц"].astype(str).str.strip() != ""]
        if len(prog_df) > 0:
            chart = alt.Chart(prog_df).mark_bar().encode(
                x=alt.X("Тоо:Q"),
                y=alt.Y("Явц:N", sort="-x"),
                tooltip=["Явц", "Тоо"]
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            _render_empty()

    with c5:
        st.markdown("### Хариуцагч тус бүр")
        owner_df = (
            fdf.groupby("Хариуцагч", dropna=False)
            .size()
            .reset_index(name="Тоо")
            .sort_values("Тоо", ascending=False)
        )
        owner_df = owner_df[
            (owner_df["Хариуцагч"].astype(str).str.strip() != "") &
            (owner_df["Хариуцагч"].astype(str).str.strip() != "-")
        ]
        if len(owner_df) > 0:
            chart = alt.Chart(owner_df).mark_bar().encode(
                x=alt.X("Тоо:Q"),
                y=alt.Y("Хариуцагч:N", sort="-x"),
                tooltip=["Хариуцагч", "Тоо"]
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            _render_empty()

    st.markdown("### Дэлгэрэнгүй хүснэгт")
    show_cols = [c for c in [
        "Огноо", "Дууссан огноо", "Явц", "Ажлын нэр",
        "Хариуцагчийн тайлбар", "Хариуцагч"
    ] if c in fdf.columns]

    if show_cols:
        st.dataframe(fdf[show_cols], use_container_width=True, hide_index=True)
    else:
        _render_empty("Хүснэгтэнд харуулах багана олдсонгүй.")


def render_ticket_dashboard(excel_path: str):
    tab1, tab2 = st.tabs(["AllCall Ticket", "AllMed Ticket"])

    with tab1:
        render_allcall_ticket_dashboard(excel_path)

    with tab2:
        render_allmed_ticket_dashboard(excel_path)