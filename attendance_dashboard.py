import streamlit as st
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import load_workbook

# =========================
# TIMEZONE
# =========================
APP_TZ = ZoneInfo("Asia/Ulaanbaatar")

def _now_local():
    return datetime.now(APP_TZ)


# =========================
# HELPERS
# =========================
def _normalize_sheet_name(name: str) -> str:
    return str(name).strip().lower().replace("_", "").replace("-", "").replace(" ", "")


def _find_sheet_name(excel_path: str, candidates: list[str]):
    xls = pd.ExcelFile(excel_path)
    actual = xls.sheet_names
    mapping = {_normalize_sheet_name(s): s for s in actual}

    for c in candidates:
        key = _normalize_sheet_name(c)
        if key in mapping:
            return mapping[key]
    return None


def _ensure_text(df, cols):
    for c in cols:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
    return df


def _name_match(a, b):
    a = str(a).lower()
    b = str(b).lower()
    return b in a


def _is_done(x):
    x = str(x).lower()
    return any(w in x for w in ["дууссан", "done", "closed"])


# =========================
# LOAD TASKS
# =========================
def _load_all_tasks(excel_path):

    frames = []

    # ===== OPERATION =====
    for name, cands in [
        ("AllCall Operation", ["AllCall operation"]),
        ("AllMed Operation", ["AllMed operation"]),
    ]:
        sheet = _find_sheet_name(excel_path, cands)
        if not sheet:
            continue

        df = pd.read_excel(excel_path, sheet_name=sheet)
        df = _ensure_text(df, ["Төслийн нэр", "Ажлын төрөл", "Явц", "Хариуцагч"])

        df = df[~df["Явц"].apply(_is_done)]

        df["_source"] = "operation"
        df["_sheet"] = sheet
        df["_system"] = name
        df["_name"] = df["Төслийн нэр"]
        df["_type"] = df["Ажлын төрөл"]
        df["_owner"] = df["Хариуцагч"]
        df["_status"] = df["Явц"]

        df["task_label"] = df["_system"] + " | " + df["_name"] + " | " + df["_type"]

        frames.append(df)

    # ===== ALLMED TICKET =====
    sheet = _find_sheet_name(excel_path, ["AllMed Ticket"])
    if sheet:
        df = pd.read_excel(excel_path, sheet_name=sheet)
        df = _ensure_text(df, ["Ажлын нэр", "Явц", "Хариуцагч"])

        df = df[~df["Явц"].apply(_is_done)]

        df["_source"] = "ticket"
        df["_sheet"] = sheet
        df["_system"] = "AllMed Ticket"
        df["_name"] = df["Ажлын нэр"]
        df["_type"] = "Ticket"
        df["_owner"] = df["Хариуцагч"]
        df["_status"] = df["Явц"]

        df["task_label"] = df["_system"] + " | " + df["_name"]

        frames.append(df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def _get_employees(df):
    return sorted(set(df["_owner"].dropna()))


def _filter_tasks(df, emp):
    return df[df["_owner"].apply(lambda x: _name_match(x, emp))]


# =========================
# LOG
# =========================
def _ensure_log(excel_path):
    wb = load_workbook(excel_path)

    if "Attendance Log" not in wb.sheetnames:
        ws = wb.create_sheet("Attendance Log")
        ws.append([
            "ID","Employee","Source","Sheet","Task","Type",
            "StartDate","StartTime","EndDate","EndTime",
            "Minutes","Status"
        ])
        wb.save(excel_path)

    wb.close()


def _append_log(excel_path, emp, row):
    _ensure_log(excel_path)
    wb = load_workbook(excel_path)
    ws = wb["Attendance Log"]

    now = _now_local()

    ws.append([
        ws.max_row,
        emp,
        row["_source"],
        row["_sheet"],
        row["_name"],
        row["_type"],
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M:%S"),
        "",
        "",
        "",
        "ACTIVE"
    ])

    wb.save(excel_path)
    wb.close()


def _finish_log(excel_path, emp):
    _ensure_log(excel_path)
    wb = load_workbook(excel_path)
    ws = wb["Attendance Log"]

    now = _now_local()

    for r in range(ws.max_row, 1, -1):
        if ws.cell(r,2).value == emp and ws.cell(r,12).value == "ACTIVE":

            start = ws.cell(r,7).value + " " + ws.cell(r,8).value
            start_dt = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")

            mins = int((now - start_dt).total_seconds()/60)

            ws.cell(r,9).value = now.strftime("%Y-%m-%d")
            ws.cell(r,10).value = now.strftime("%H:%M:%S")
            ws.cell(r,11).value = mins
            ws.cell(r,12).value = "DONE"

            wb.save(excel_path)
            wb.close()

            return True

    wb.close()
    return False


# =========================
# MAIN
# =========================
def render_attendance_dashboard(excel_path):

    df = _load_all_tasks(excel_path)
    if df.empty:
        st.warning("Task алга")
        return

    st.markdown("### Clock In / Clock Out")

    col1, col2 = st.columns(2)

    employees = _get_employees(df)

    with col1:
        emp = st.selectbox("Ажилтан", [""] + employees)

    tasks = _filter_tasks(df, emp)

    if emp and not tasks.empty:

        task_label = st.selectbox("Task", tasks["task_label"])

        row = tasks[tasks["task_label"] == task_label].iloc[0]

        c1, c2 = st.columns(2)

        with c1:
            if st.button("⏱️ Clock In"):
                _append_log(excel_path, emp, row)
                st.success("Эхэллээ")
                st.rerun()

        with c2:
            if st.button("✅ Clock Out"):
                ok = _finish_log(excel_path, emp)
                if ok:
                    st.success("Дууслаа")
                else:
                    st.error("Идэвхтэй task алга")
                st.rerun()
