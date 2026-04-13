import streamlit as st
import pandas as pd
import altair as alt

from modules.attendance_dashboard import render_attendance_dashboard


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _normalize_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


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
            f"{label} sheet олдсонгүй. Боломжит sheet нэрс: "
            f"{', '.join(actual_sheets) if actual_sheets else 'sheet алга'}"
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


def _prepare_operation(df: pd.DataFrame, system_name: str) -> pd.DataFrame:
    df = _clean_columns(df)

    expected_cols = [
        "Ажлын төрөл",
        "Төслийн нэр",
        "Эхлэх огноо",
        "Дуусах огноо",
        "Хугацаа",
        "Хариуцагч",
        "Дэмжлэг",
        "Явц",
        "Явцын тайлбар",
    ]

    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    df["Ажлын төрөл"] = _normalize_text(df["Ажлын төрөл"])
    df["Төслийн нэр"] = _normalize_text(df["Төслийн нэр"])
    df["Хариуцагч"] = _normalize_text(df["Хариуцагч"])
    df["Дэмжлэг"] = _normalize_text(df["Дэмжлэг"])
    df["Явц"] = _normalize_text(df["Явц"])
    df["Явцын тайлбар"] = _normalize_text(df["Явцын тайлбар"])

    df["Эхлэх огноо"] = pd.to_datetime(df["Эхлэх огноо"], errors="coerce")
    df["Дуусах огноо"] = pd.to_datetime(df["Дуусах огноо"], errors="coerce")
    df["Хугацаа"] = pd.to_numeric(df["Хугацаа"], errors="coerce").fillna(0).astype(int)
    df["Систем"] = system_name

    df = df[
        (df["Ажлын төрөл"] != "") |
        (df["Төслийн нэр"] != "") |
        (df["Хариуцагч"] != "")
    ].copy()

    today = pd.Timestamp.today().normalize()

    def calc_state(row):
        progress = str(row.get("Явц", "")).strip().lower()
        due = row.get("Дуусах огноо", pd.NaT)

        if any(x in progress for x in ["хийгдсэн", "дууссан", "шийдсэн", "done", "closed", "completed"]):
            return "Дууссан"

        if pd.notna(due):
            if due < today:
                return "Хугацаа хэтэрсэн"
            if due == today:
                return "Өнөөдөр дуусна"

        if "хийгдэж" in progress:
            return "Хийгдэж байна"
        if "төлөвлөсөн" in progress:
            return "Төлөвлөсөн"
        if "хүлээгдэж" in progress:
            return "Хүлээгдэж байна"
        if "тест" in progress:
            return "Тест хийж байна"
        if "хэлэлц" in progress:
            return "Хэлэлцэж байна"

        return "Тодорхойгүй"

    df["Төлөв_тооцоолсон"] = df.apply(calc_state, axis=1)
    return df


def _render_metrics(df: pd.DataFrame):
    total = len(df)
    done = len(df[df["Төлөв_тооцоолсон"] == "Дууссан"])
    overdue = len(df[df["Төлөв_тооцоолсон"] == "Хугацаа хэтэрсэн"])
    planning = len(df[df["Төлөв_тооцоолсон"] == "Төлөвлөсөн"])
    in_progress = len(df[df["Төлөв_тооцоолсон"] == "Хийгдэж байна"])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Нийт ажил", f"{total:,}")
    c2.metric("Дууссан", f"{done:,}")
    c3.metric("Хугацаа хэтэрсэн", f"{overdue:,}")
    c4.metric("Төлөвлөсөн", f"{planning:,}")
    c5.metric("Хийгдэж байна", f"{in_progress:,}")


def _render_operation_table(df: pd.DataFrame, key_prefix: str) -> pd.DataFrame:
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        types = ["Бүгд"] + sorted(
            [x for x in df["Ажлын төрөл"].dropna().astype(str).unique().tolist() if x.strip()]
        )
        selected_type = st.selectbox("Ажлын төрөл", types, key=f"{key_prefix}_type")

    with c2:
        owners = ["Бүгд"] + sorted(
            [x for x in df["Хариуцагч"].dropna().astype(str).unique().tolist() if x.strip()]
        )
        selected_owner = st.selectbox("Хариуцагч", owners, key=f"{key_prefix}_owner")

    with c3:
        states = ["Бүгд"] + sorted(
            [x for x in df["Төлөв_тооцоолсон"].dropna().astype(str).unique().tolist() if x.strip()]
        )
        selected_state = st.selectbox("Төлөв", states, key=f"{key_prefix}_state")

    with c4:
        search_text = st.text_input("Төслийн нэр хайх", key=f"{key_prefix}_search").strip().lower()

    fdf = df.copy()

    if selected_type != "Бүгд":
        fdf = fdf[fdf["Ажлын төрөл"] == selected_type]
    if selected_owner != "Бүгд":
        fdf = fdf[fdf["Хариуцагч"] == selected_owner]
    if selected_state != "Бүгд":
        fdf = fdf[fdf["Төлөв_тооцоолсон"] == selected_state]
    if search_text:
        fdf = fdf[fdf["Төслийн нэр"].str.lower().str.contains(search_text, na=False)]

    show_df = fdf.copy()
    for col in ["Эхлэх огноо", "Дуусах огноо"]:
        if col in show_df.columns:
            show_df[col] = pd.to_datetime(show_df[col], errors="coerce").dt.strftime("%Y-%m-%d")

    cols = [
        "Ажлын төрөл",
        "Төслийн нэр",
        "Эхлэх огноо",
        "Дуусах огноо",
        "Хугацаа",
        "Хариуцагч",
        "Дэмжлэг",
        "Явц",
        "Төлөв_тооцоолсон",
        "Явцын тайлбар",
    ]
    cols = [c for c in cols if c in show_df.columns]

    st.dataframe(show_df[cols], use_container_width=True, hide_index=True)
    return fdf


def _render_operation_charts(df: pd.DataFrame):
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### Төлөвийн тархалт")
        state_df = (
            df.groupby("Төлөв_тооцоолсон", dropna=False)
            .size()
            .reset_index(name="Тоо")
            .sort_values("Тоо", ascending=False)
        )
        state_df = state_df[state_df["Төлөв_тооцоолсон"].astype(str).str.strip() != ""]
        if len(state_df) > 0:
            chart = alt.Chart(state_df).mark_bar().encode(
                x=alt.X("Тоо:Q", title="Тоо"),
                y=alt.Y("Төлөв_тооцоолсон:N", sort="-x", title="Төлөв"),
                tooltip=["Төлөв_тооцоолсон", "Тоо"]
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Өгөгдөл алга.")

    with c2:
        st.markdown("### Хариуцагч тус бүр")
        owner_df = (
            df.groupby("Хариуцагч", dropna=False)
            .size()
            .reset_index(name="Тоо")
            .sort_values("Тоо", ascending=False)
        )
        owner_df = owner_df[owner_df["Хариуцагч"].astype(str).str.strip() != ""]
        if len(owner_df) > 0:
            chart = alt.Chart(owner_df).mark_bar().encode(
                x=alt.X("Тоо:Q", title="Тоо"),
                y=alt.Y("Хариуцагч:N", sort="-x", title="Хариуцагч"),
                tooltip=["Хариуцагч", "Тоо"]
            ).properties(height=320)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Өгөгдөл алга.")


def render_operation_dashboard(excel_path: str):
    allcall_raw = _read_excel_sheet_auto(
        excel_path,
        candidates=[
            "AllCall operation",
            "AllCall Operation",
            "Operation AllCall",
            "AllCalloperation",
            "operation allcall",
        ],
        label="AllCall Operation",
    )

    allmed_raw = _read_excel_sheet_auto(
        excel_path,
        candidates=[
            "AllMed operation",
            "AllMed Operation",
            "Operation AllMed",
            "AllMedoperation",
            "operation allmed",
        ],
        label="AllMed Operation",
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "AllCall Operation",
        "AllMed Operation",
        "Нэгдсэн Operation",
        "Attendance",
    ])

    with tab1:
        st.markdown("## AllCall Operation Dashboard")
        if allcall_raw is None:
            st.info("AllCall operation өгөгдөл олдсонгүй.")
        else:
            allcall_df = _prepare_operation(allcall_raw, "AllCall")
            _render_metrics(allcall_df)
            filtered_df = _render_operation_table(allcall_df, "allcall_op")
            _render_operation_charts(filtered_df)

    with tab2:
        st.markdown("## AllMed Operation Dashboard")
        if allmed_raw is None:
            st.info("AllMed operation өгөгдөл олдсонгүй.")
        else:
            allmed_df = _prepare_operation(allmed_raw, "AllMed")
            _render_metrics(allmed_df)
            filtered_df = _render_operation_table(allmed_df, "allmed_op")
            _render_operation_charts(filtered_df)

    with tab3:
        st.markdown("## Нэгдсэн Operation Dashboard")
        if allcall_raw is None and allmed_raw is None:
            st.info("Operation өгөгдөл олдсонгүй.")
        else:
            frames = []
            if allcall_raw is not None:
                frames.append(_prepare_operation(allcall_raw, "AllCall"))
            if allmed_raw is not None:
                frames.append(_prepare_operation(allmed_raw, "AllMed"))

            combined_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

            if combined_df.empty:
                st.info("Нэгдсэн operation өгөгдөл хоосон байна.")
            else:
                _render_metrics(combined_df)

                c1, c2, c3, c4 = st.columns(4)

                with c1:
                    selected_system = st.selectbox(
                        "Систем",
                        ["Бүгд", "AllCall", "AllMed"],
                        key="op_system"
                    )
                with c2:
                    types = ["Бүгд"] + sorted(
                        [x for x in combined_df["Ажлын төрөл"].dropna().astype(str).unique().tolist() if x.strip()]
                    )
                    selected_type = st.selectbox("Ажлын төрөл", types, key="op_type_all")

                with c3:
                    states = ["Бүгд"] + sorted(
                        [x for x in combined_df["Төлөв_тооцоолсон"].dropna().astype(str).unique().tolist() if x.strip()]
                    )
                    selected_state = st.selectbox("Төлөв", states, key="op_state_all")

                with c4:
                    search_text = st.text_input("Төслийн нэр хайх", key="op_search_all").strip().lower()

                fdf = combined_df.copy()

                if selected_system != "Бүгд":
                    fdf = fdf[fdf["Систем"] == selected_system]
                if selected_type != "Бүгд":
                    fdf = fdf[fdf["Ажлын төрөл"] == selected_type]
                if selected_state != "Бүгд":
                    fdf = fdf[fdf["Төлөв_тооцоолсон"] == selected_state]
                if search_text:
                    fdf = fdf[fdf["Төслийн нэр"].str.lower().str.contains(search_text, na=False)]

                show_df = fdf.copy()
                for col in ["Эхлэх огноо", "Дуусах огноо"]:
                    if col in show_df.columns:
                        show_df[col] = pd.to_datetime(show_df[col], errors="coerce").dt.strftime("%Y-%m-%d")

                cols = [
                    "Систем",
                    "Ажлын төрөл",
                    "Төслийн нэр",
                    "Эхлэх огноо",
                    "Дуусах огноо",
                    "Хугацаа",
                    "Хариуцагч",
                    "Дэмжлэг",
                    "Явц",
                    "Төлөв_тооцоолсон",
                    "Явцын тайлбар",
                ]
                cols = [c for c in cols if c in show_df.columns]

                st.dataframe(show_df[cols], use_container_width=True, hide_index=True)
                _render_operation_charts(fdf)

    with tab4:
        st.markdown("## Attendance Dashboard")
        render_attendance_dashboard(excel_path)
