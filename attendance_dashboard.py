import streamlit as st
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import load_workbook

APP_TZ = ZoneInfo("Asia/Ulaanbaatar")


def _now_local():
    return datetime.now(APP_TZ)


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
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("").astype(str).str.strip()
    return df


def _name_match(a, b):
    a = str(a).strip().lower()
    b = str(b).strip().lower()

    if not a or not b:
        return False

    if a == b:
        return True

    normalized = a.replace(";", ",").replace("/", ",").replace("\n", ",")
    parts = [x.strip() for x in normalized.split(",") if x.strip()]
    if b in parts:
        return True

    return b in a


def _is_done(x):
    x = str(x).strip().lower()
    return any(w in x for w in ["дууссан", "done", "closed", "completed", "resolved", "хийгдсэн"])


def _load_operation_sheet(excel_path: str, system_name: str, candidates: list[str]) -> pd.DataFrame:
    sheet = _find_sheet_name(excel_path, candidates)
    if not sheet:
        return pd.DataFrame()

    df = pd.read_excel(excel_path, sheet_name=sheet)
    df = _ensure_text(df, ["Төслийн нэр", "Ажлын төрөл", "Явц", "Хариуцагч"])
    df = df[
        (df["Төслийн нэр"] != "") |
        (df["Ажлын төрөл"] != "") |
        (df["Хариуцагч"] != "")
    ].copy()

    if df.empty:
        return pd.DataFrame()

    df = df[~df["Явц"].apply(_is_done)].copy()

    df["_source"] = "operation"
    df["_sheet"] = sheet
    df["_system"] = system_name
    df["_name"] = df["Төслийн нэр"]
    df["_type"] = df["Ажлын төрөл"]
    df["_owner"] = df["Хариуцагч"]
    df["_status"] = df["Явц"]
    df["_match_key1"] = df["Төслийн нэр"]
    df["_match_key2"] = df["Ажлын төрөл"]

    df["task_label"] = (
        df["_system"].astype(str) + " | " +
        df["_name"].astype(str) + " | " +
        df["_type"].astype(str)
    )

    return df


def _load_allmed_ticket_sheet(excel_path: str) -> pd.DataFrame:
    sheet = _find_sheet_name(
        excel_path,
        ["AllMed Ticket", "Ticket AllMed", "AllMedTicket", "AllMed", "ticket allmed"]
    )
    if not sheet:
        return pd.DataFrame()

    df = pd.read_excel(excel_path, sheet_name=sheet)
    df = _ensure_text(df, ["Ажлын нэр", "Явц", "Хариуцагч"])
    df = df[
        (df["Ажлын нэр"] != "") |
        (df["Хариуцагч"] != "")
    ].copy()

    if df.empty:
        return pd.DataFrame()

    df = df[~df["Явц"].apply(_is_done)].copy()

    df["_source"] = "allmed_ticket"
    df["_sheet"] = sheet
    df["_system"] = "AllMed Ticket"
    df["_name"] = df["Ажлын нэр"]
    df["_type"] = "Ticket"
    df["_owner"] = df["Хариуцагч"]
    df["_status"] = df["Явц"]
    df["_match_key1"] = df["Ажлын нэр"]
    df["_match_key2"] = df["Хариуцагч"]

    df["task_label"] = (
        df["_system"].astype(str) + " | " +
        df["_name"].astype(str)
    )

    return df


def _load_all_tasks(excel_path: str):
    frames = []

    allcall_op = _load_operation_sheet(
        excel_path,
        "AllCall Operation",
        ["AllCall operation", "AllCall Operation", "Operation AllCall", "operation allcall"]
    )
    if not allcall_op.empty:
        frames.append(allcall_op)

    allmed_op = _load_operation_sheet(
        excel_path,
        "AllMed Operation",
        ["AllMed operation", "AllMed Operation", "Operation AllMed", "operation allmed"]
    )
    if not allmed_op.empty:
        frames.append(allmed_op)

    allmed_ticket = _load_allmed_ticket_sheet(excel_path)
    if not allmed_ticket.empty:
        frames.append(allmed_ticket)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


def _get_employees(df):
    if df.empty:
        return []
    return sorted(set([str(x).strip() for x in df["_owner"].dropna() if str(x).strip()]))


def _filter_tasks(df, emp):
    if df.empty or not emp:
        return pd.DataFrame()
    return df[df["_owner"].apply(lambda x: _name_match(x, emp))].copy()


def _ensure_log(excel_path):
    wb = load_workbook(excel_path)

    if "Attendance Log" not in wb.sheetnames:
        ws = wb.create_sheet("Attendance Log")
        ws.append([
            "ID", "Employee", "Source", "Sheet", "Task", "Type",
            "MatchKey1", "MatchKey2",
            "StartDate", "StartTime", "EndDate", "EndTime",
            "Minutes", "Status"
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
        row["_match_key1"],
        row["_match_key2"],
        now.strftime("%Y-%m-%d"),
        now.strftime("%H:%M:%S"),
        "",
        "",
        "",
        "ACTIVE",
    ])

    wb.save(excel_path)
    wb.close()


def _finish_log(excel_path, emp):
    _ensure_log(excel_path)
    wb = load_workbook(excel_path)
    ws = wb["Attendance Log"]

    now = _now_local()

    for r in range(ws.max_row, 1, -1):
        row_emp = str(ws.cell(r, 2).value).strip()
        row_status = str(ws.cell(r, 14).value).strip()

        if row_emp == str(emp).strip() and row_status == "ACTIVE":
            start_date = str(ws.cell(r, 9).value).strip()
            start_time = str(ws.cell(r, 10).value).strip()
            start = f"{start_date} {start_time}"

            try:
                start_dt = datetime.strptime(start, "%Y-%m-%d %H:%M:%S").replace(tzinfo=APP_TZ)
                mins = int((now - start_dt).total_seconds() / 60)
            except Exception:
                mins = ""

            ws.cell(r, 11).value = now.strftime("%Y-%m-%d")
            ws.cell(r, 12).value = now.strftime("%H:%M:%S")
            ws.cell(r, 13).value = mins
            ws.cell(r, 14).value = "DONE"

            result = {
                "source": ws.cell(r, 3).value,
                "sheet": ws.cell(r, 4).value,
                "task": ws.cell(r, 5).value,
                "type": ws.cell(r, 6).value,
                "match_key1": ws.cell(r, 7).value,
                "match_key2": ws.cell(r, 8).value,
                "minutes": mins,
            }

            wb.save(excel_path)
            wb.close()
            return True, result

    wb.close()
    return False, None


def _update_operation_status(excel_path, sheet_name, project_name, task_type, new_status):
    wb = load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return False, f"{sheet_name} sheet олдсонгүй."

    ws = wb[sheet_name]

    headers = {}
    for col in range(1, ws.max_column + 1):
        headers[str(ws.cell(1, col).value).strip()] = col

    project_col = headers.get("Төслийн нэр")
    type_col = headers.get("Ажлын төрөл")
    status_col = headers.get("Явц")

    if not project_col or not type_col or not status_col:
        wb.close()
        return False, "Operation sheet-д шаардлагатай багана олдсонгүй."

    target_row = None
    for row in range(2, ws.max_row + 1):
        p = str(ws.cell(row, project_col).value).strip()
        t = str(ws.cell(row, type_col).value).strip()
        if p == str(project_name).strip() and t == str(task_type).strip():
            target_row = row
            break

    if target_row is None:
        wb.close()
        return False, "Operation task мөр олдсонгүй."

    ws.cell(target_row, status_col).value = new_status
    wb.save(excel_path)
    wb.close()
    return True, "Амжилттай шинэчлэгдлээ."


def _update_allmed_ticket_status(excel_path, sheet_name, task_name, owner_name, new_status):
    wb = load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return False, f"{sheet_name} sheet олдсонгүй."

    ws = wb[sheet_name]

    headers = {}
    for col in range(1, ws.max_column + 1):
        headers[str(ws.cell(1, col).value).strip()] = col

    task_col = headers.get("Ажлын нэр")
    owner_col = headers.get("Хариуцагч")
    status_col = headers.get("Явц")
    done_date_col = headers.get("Дууссан огноо")

    if not task_col or not owner_col or not status_col:
        wb.close()
        return False, "AllMed Ticket sheet-д шаардлагатай багана олдсонгүй."

    target_row = None
    for row in range(2, ws.max_row + 1):
        t = str(ws.cell(row, task_col).value).strip()
        o = str(ws.cell(row, owner_col).value).strip()
        if t == str(task_name).strip() and _name_match(o, str(owner_name).strip()):
            target_row = row
            break

    if target_row is None:
        wb.close()
        return False, "AllMed Ticket мөр олдсонгүй."

    ws.cell(target_row, status_col).value = new_status

    if new_status == "Дууссан" and done_date_col:
        ws.cell(target_row, done_date_col).value = _now_local().strftime("%Y-%m-%d %H:%M:%S")

    wb.save(excel_path)
    wb.close()
    return True, "Амжилттай шинэчлэгдлээ."


def _update_source_status(excel_path, source, sheet_name, match_key1, match_key2, new_status):
    if source == "operation":
        return _update_operation_status(
            excel_path,
            sheet_name,
            match_key1,
            match_key2,
            new_status
        )

    if source == "allmed_ticket":
        return _update_allmed_ticket_status(
            excel_path,
            sheet_name,
            match_key1,
            match_key2,
            new_status
        )

    return False, f"Танигдаагүй source: {source}"


def _get_active_task(log_df, emp):
    if log_df.empty or not emp:
        return pd.DataFrame()

    fdf = log_df.copy()
    if "Employee" not in fdf.columns or "Status" not in fdf.columns:
        return pd.DataFrame()

    fdf["Employee"] = fdf["Employee"].fillna("").astype(str).str.strip()
    fdf["Status"] = fdf["Status"].fillna("").astype(str).str.strip()

    return fdf[(fdf["Employee"] == str(emp).strip()) & (fdf["Status"] == "ACTIVE")].tail(1)


def render_attendance_dashboard(excel_path: str):
    df = _load_all_tasks(excel_path)
    if df.empty:
        st.warning("Task алга.")
        return

    log_df = pd.read_excel(excel_path, sheet_name="Attendance Log") if "Attendance Log" in pd.ExcelFile(excel_path).sheet_names else pd.DataFrame()
    if not log_df.empty:
        log_df.columns = [str(c).strip() for c in log_df.columns]

    st.markdown("### Clock In / Clock Out")

    col1, col2 = st.columns(2)

    employees = _get_employees(df)

    with col1:
        emp = st.selectbox("Ажилтан", [""] + employees)

    with col2:
        pass

    tasks = _filter_tasks(df, emp)

    if emp and not tasks.empty:
        task_label = st.selectbox("Task", tasks["task_label"])
        row = tasks[tasks["task_label"] == task_label].iloc[0]

        c1, c2 = st.columns(2)

        with c1:
            if st.button("⏱️ Clock In", use_container_width=True):
                active_df = _get_active_task(log_df, emp) if not log_df.empty else pd.DataFrame()
                if not active_df.empty:
                    st.error("Энэ ажилтан аль хэдийн идэвхтэй task дээр ажиллаж байна.")
                else:
                    _append_log(excel_path, emp, row)

                    ok, msg = _update_source_status(
                        excel_path,
                        row["_source"],
                        row["_sheet"],
                        row["_match_key1"],
                        row["_match_key2"],
                        "Хийгдэж байна"
                    )

                    if ok:
                        st.success("Clock In амжилттай.")
                    else:
                        st.warning(f"Clock In бүртгэгдсэн ч source update дутуу: {msg}")
                    st.rerun()

        with c2:
            if st.button("✅ Clock Out", use_container_width=True):
                ok, result = _finish_log(excel_path, emp)
                if not ok:
                    st.error("Идэвхтэй task алга.")
                else:
                    upd_ok, upd_msg = _update_source_status(
                        excel_path,
                        result["source"],
                        result["sheet"],
                        result["match_key1"],
                        result["match_key2"],
                        "Дууссан"
                    )
                    if upd_ok:
                        st.success(f"Clock Out амжилттай. Ажилласан хугацаа: {result['minutes']} минут.")
                    else:
                        st.warning(f"Clock Out бүртгэгдсэн ч source update дутуу: {upd_msg}")
                    st.rerun()

    elif emp:
        st.info("Энэ ажилтанд task алга.")

    if not log_df.empty:
        st.markdown("---")
        st.markdown("### Attendance Log")
        st.dataframe(log_df.sort_index(ascending=False), use_container_width=True, hide_index=True)
