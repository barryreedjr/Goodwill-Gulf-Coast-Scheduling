import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
import io

st.set_page_config(page_title="Goodwill Scheduler", layout="wide")
st.title("Goodwill Gulf Coast — Weekly Scheduler (weekly caps, lunch, night rotation, Paycom export)")
st.caption(
    "Upload the input Excel, choose an optional start date, and download a 4-week schedule. "
    "Weekly max hours are enforced per week, long shifts add lunch automatically (>5 paid hours), "
    "night shifts are spread (>=1 per employee per week, when possible), and a Paycom import template is generated."
)

DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
KNOWN_ROLES = ["Cashier", "Donation", "Pricer", "Hanger", "Manager"]


# ------------------ Helpers ------------------
def fmt_time(hhmm: str) -> str:
    try:
        t = datetime.strptime(hhmm, "%H:%M")
        s = t.strftime("%I:%M %p")
        return s.lstrip("0")
    except Exception:
        return hhmm


def hhmm_to_4dig(hhmm: str) -> str:
    # "HH:MM" -> "HHMM" (always 4 digits)
    try:
        t = datetime.strptime(hhmm, "%H:%M")
        return t.strftime("%H%M")
    except Exception:
        # best effort fallback
        return hhmm.replace(":", "")[:4].rjust(4, "0")


def get_numeric_series(df: pd.DataFrame, col: str, default_val, length: int) -> pd.Series:
    if col in df.columns:
        s = pd.to_numeric(df[col], errors="coerce").fillna(default_val)
    else:
        s = pd.Series([default_val] * length)
    return s


# ------------------ Loaders ------------------
def load_inputs(file):
    xls = pd.ExcelFile(file)
    employees = pd.read_excel(xls, "employees").fillna("")

    # availability_simple optional
    availability_simple = (
        pd.read_excel(xls, "availability_simple").fillna("")
        if "availability_simple" in xls.sheet_names
        else None
    )
    # availability optional (detailed windows)
    availability = (
        pd.read_excel(xls, "availability").fillna("")
        if "availability" in xls.sheet_names
        else None
    )

    # coverage can be "coverage" (new) or "coverage_template" (older)
    coverage_sheet = "coverage" if "coverage" in xls.sheet_names else "coverage_template"
    coverage = pd.read_excel(xls, coverage_sheet).fillna("")

    anchors = pd.read_excel(xls, "shift_anchors").fillna("")
    rules = pd.read_excel(xls, "rules").fillna("")

    if "Day" in coverage.columns:
        coverage["Day"] = coverage["Day"].astype(str).str.strip().str.title()
    if "ShiftType" in coverage.columns:
        coverage["ShiftType"] = coverage["ShiftType"].astype(str).str.strip().str.title()

    return employees, availability, availability_simple, coverage, anchors, rules


def build_availability_map(employees, availability, availability_simple):
    """
    Return dict[employee][day] -> list of (start,end) windows in HH:MM.
    If no availability provided, assume full 08:00-20:30 daily.
    Note: availability_simple uses Yes = NOT available (blocked).
    """
    avail = {}

    # Simple Yes/No per day: Yes = NOT available
    if availability_simple is not None and not availability_simple.empty:
        df = availability_simple.copy()
        df.columns = [("Employee" if c == "Employee" else str(c).strip().title()) for c in df.columns]
        for _, row in df.iterrows():
            emp = row["Employee"]
            for d in DAYS:
                resp = str(row.get(d, "")).strip().lower()
                if resp in ("yes", "y", "true", "1", "x"):
                    avail.setdefault(emp, {}).setdefault(d, [])
                else:
                    avail.setdefault(emp, {}).setdefault(d, []).append(("08:00", "20:30"))
        return avail

    # Detailed availability windows
    if availability is not None and not availability.empty:
        df = availability.copy()
        df["Day"] = df["Day"].astype(str).str.strip().str.title()
        for _, row in df.iterrows():
            emp = row["Employee"]
            d = row["Day"]
            start = str(row.get("Start", "")).strip()
            end = str(row.get("End", "")).strip()
            if start and end:
                avail.setdefault(emp, {}).setdefault(d, []).append((start, end))
        return avail

    # Default: everyone fully available
    for emp in employees["Employee"]:
        for d in DAYS:
            avail.setdefault(emp, {}).setdefault(d, []).append(("08:00", "20:30"))
    return avail


def employee_functions_map(employees):
    # Prefer skills matrix Yes/No columns; fallback to CSV 'Functions' if present.
    role_cols_present = [r for r in KNOWN_ROLES if r in employees.columns]
    m = {}
    if role_cols_present:
        for _, r in employees.iterrows():
            emp = r["Employee"]
            skills = set()
            for role in role_cols_present:
                val = str(r.get(role, "")).strip().lower()
                if val in ("yes", "y", "true", "1", "x"):
                    skills.add(role)
            if not skills and "Functions" in employees.columns:
                raw = str(r.get("Functions", ""))
                for token in raw.replace(";", ",").replace("/", ",").split(","):
                    t = token.strip().title()
                    if t in KNOWN_ROLES:
                        skills.add(t)
            m[emp] = skills
        return m

    # CSV fallback only
    for _, r in employees.iterrows():
        emp = r["Employee"]
        skills = set()
        raw = str(r.get("Functions", ""))
        for token in raw.replace(";", ",").replace("/", ",").split(","):
            t = token.strip().title()
            if t in KNOWN_ROLES:
                skills.add(t)
        m[emp] = skills
    return m


def weekend_rotation(employees):
    emps = list(employees["Employee"])
    groups = {i: set() for i in range(4)}
    for i, e in enumerate(emps):
        groups[i % 4].add(e)
    return groups


def generate_week_dates(start_date: date):
    # Align to Sunday
    start = start_date - timedelta(days=start_date.weekday() + 1 if start_date.weekday() != 6 else 0)
    return [[start + timedelta(days=7 * w + i) for i in range(7)] for w in range(4)]


# ------------------ Lunch & Night Rules ------------------
def is_night_shift(stype: str, end_time_str: str, night_types: set, night_end_threshold: str | None):
    if stype in night_types:
        return True
    if night_end_threshold:
        try:
            end_dt = datetime.strptime(end_time_str, "%H:%M").time()
            thr_dt = datetime.strptime(night_end_threshold, "%H:%M").time()
            return end_dt >= thr_dt
        except Exception:
            return False
    return False


def compute_lunch_pad(paid_hours: float, trigger_hours: float, lunch_minutes: int) -> float:
    """Return lunch padding hours to add to the scheduled block (not paid hours)."""
    if paid_hours > trigger_hours:
        return lunch_minutes / 60.0
    return 0.0


def anchored_shift_window(shift_type, paid_hours, anchors_map, rules) -> tuple[str, str, float]:
    """
    Return (start_str, end_str, lunch_pad_hours).
    If lunch is required, expand the scheduled block by lunch_pad_hours while keeping paid_hours as-is.
    """
    a = anchors_map[shift_type]
    trigger = float(rules.get("LunchTriggerHours", 5))
    lunch_minutes = int(rules.get("LunchMinutes", 30))
    lunch_pad = compute_lunch_pad(float(paid_hours), trigger, lunch_minutes)

    if a["Anchor"] == "Start":
        stt = datetime.strptime(a["Time"], "%H:%M")
        en = stt + timedelta(hours=float(paid_hours) + lunch_pad)
        return (a["Time"], en.strftime("%H:%M"), lunch_pad)
    else:
        en = datetime.strptime(a["Time"], "%H:%M")
        stt = en - timedelta(hours=float(paid_hours) + lunch_pad)
        return (stt.strftime("%H:%M"), a["Time"], lunch_pad)


def scheduled_hours_between(start_str, end_str):
    return (datetime.strptime(end_str, "%H:%M") - datetime.strptime(start_str, "%H:%M")).seconds / 3600


# ------------------ Constraints ------------------
def can_assign(
    emp,
    role,
    day_name,
    start_str,
    end_str,
    avail_map,
    funcs_map,
    assigned_list_for_week,
    rules,
    assigned_paid_hours_week,
    days_worked_week,
    min_hours_map,
    weekly_max_paid_hours_map,
):
    if role not in funcs_map.get(emp, set()):
        return False

    windows = avail_map.get(emp, {}).get(day_name, [])
    if not any(a_start <= start_str and a_end >= end_str for a_start, a_end in windows):
        return False

    # overlap + gap
    min_gap = float(rules.get("MinGapBetweenShiftsHours", 1))

    def to_dt_local(t):
        return datetime.combine(datetime.today(), datetime.strptime(t, "%H:%M").time())

    slot_s, slot_e = to_dt_local(start_str), to_dt_local(end_str)
    for a in assigned_list_for_week:
        if a["Employee"] != emp or a["Day"] != day_name:
            continue
        s, e = to_dt_local(a["Start"]), to_dt_local(a["End"])
        if not (slot_e <= s or slot_s >= e):
            return False
        if 0 <= (s - slot_e).total_seconds() / 3600 < min_gap or 0 <= (slot_s - e).total_seconds() / 3600 < min_gap:
            return False

    # Weekly max *paid* hours constraint
    paid_hours = float(rules.get("_candidate_paid_hours", 0))
    if assigned_paid_hours_week.get(emp, 0) + paid_hours > weekly_max_paid_hours_map.get(emp, float("inf")):
        return False

    if len(days_worked_week.get(emp, set())) >= int(rules.get("MaxDaysPerWeek", 6)) and day_name not in days_worked_week.get(emp, set()):
        return False

    return True


# ------------------ Scheduler ------------------
def run_schedule_with_summary(file, startdate_str=None, role_colors=None):
    # Load
    employees, availability, availability_simple, coverage, anchors, rules_df = load_inputs(file)
    rules = dict(zip(rules_df["Rule"], rules_df["Value"]))

    # Defaults / new knobs
    rules["MinShiftHours"] = float(rules.get("MinShiftHours", 3))
    rules["MaxShiftHours"] = float(rules.get("MaxShiftHours", 10))
    rules["MinGapBetweenShiftsHours"] = float(rules.get("MinGapBetweenShiftsHours", 1))
    rules["MaxDaysPerWeek"] = int(rules.get("MaxDaysPerWeek", 6))
    # Lunch config
    rules["LunchTriggerHours"] = float(rules.get("LunchTriggerHours", 5))
    rules["LunchMinutes"] = int(rules.get("LunchMinutes", 30))
    # Night shift config
    night_types = set([s.strip() for s in str(rules.get("NightShiftTypes", "Close")).split(",") if s.strip()])
    night_end_threshold = str(rules.get("NightShiftEndAtOrAfter", "")).strip() or None

    anchors_map = {r["ShiftType"]: {"Anchor": r["Anchor"], "Time": r["Time"]} for _, r in anchors.iterrows()}

    # Weekly max (paid) hours column
    weekly_max_series = get_numeric_series(
        employees,
        "MaxHoursPerWeek" if "MaxHoursPerWeek" in employees.columns else "MaxHours",
        999,
        len(employees),
    )
    weekly_max_paid_hours_map = dict(zip(employees["Employee"], weekly_max_series))

    # Employee ID map (supports 'EmployeeID' or 'Employee ID')
    emp_id_col = None
    for cand in ["EmployeeID", "Employee ID"]:
        if cand in employees.columns:
            emp_id_col = cand
            break
    if emp_id_col:
        emp_id_map = dict(zip(
            employees["Employee"],
            employees[emp_id_col].astype(str).str.strip()
        ))
    else:
        emp_id_map = {e: "" for e in employees["Employee"]}

    pref_hours = dict(
        zip(
            employees["Employee"],
            get_numeric_series(employees, "PreferredShiftHours", rules["MaxShiftHours"], len(employees)),
        )
    )
    funcs_map = employee_functions_map(employees)
    avail_map = build_availability_map(employees, availability, availability_simple)
    min_hours_map = dict(
        zip(employees["Employee"], get_numeric_series(employees, "MinHours", 0, len(employees)))
    )

    # start date
    if startdate_str:
        startdate = datetime.strptime(startdate_str, "%Y-%m-%d").date()
    else:
        today = datetime.today().date()
        delta = (6 - today.weekday()) % 7
        startdate = today + timedelta(days=delta)

    weeks = generate_week_dates(startdate)
    weekend_groups = weekend_rotation(employees)

    assignments_all = []

    for w_idx, week_dates in enumerate(weeks):
        # tracking per week
        assigned_paid_hours_week = {e: 0.0 for e in employees["Employee"]}
        days_worked_week = {e: set() for e in employees["Employee"]}
        nights_worked_week = {e: 0 for e in employees["Employee"]}
        assignments_this_week = []

        # rotating Sat/Sun off-group
        off_group = weekend_groups[w_idx]
        local_avail = {e: {d: list(w) for d, w in days.items()} for e, days in avail_map.items()}
        for e in off_group:
            for d in ["Saturday", "Sunday"]:
                if e in local_avail and d in local_avail[e]:
                    local_avail[e][d] = []

        for d_idx, dt in enumerate(week_dates):
            day_name = DAYS[d_idx]
            day_cov = coverage[coverage["Day"] == day_name].copy()

            # expand demand
            demands = []
            for _, r in day_cov.iterrows():
                count = int(pd.to_numeric(r.get("Count", 0), errors="coerce") or 0)
                role = r.get("Role", "")
                stype = r.get("ShiftType", "")
                demands += [{"Role": role, "ShiftType": stype} for _ in range(count)]

            # fill rare roles first
            rarity = {}
            roles_unique = [x for x in day_cov["Role"].unique() if isinstance(x, str)]
            for role in roles_unique:
                count = sum(1 for e in employees["Employee"] if role in funcs_map.get(e, set()))
                rarity[role] = 1.0 / (count if count > 0 else 0.5)
            demands.sort(key=lambda d: rarity.get(d["Role"], 1.0), reverse=True)

            for dem in demands:
                role, stype = dem["Role"], dem["ShiftType"]
                candidates = []

                for e in employees["Employee"]:
                    if role not in funcs_map.get(e, set()):
                        continue

                    paid_hours = float(pref_hours.get(e, rules["MaxShiftHours"]))
                    paid_hours = min(max(paid_hours, rules["MinShiftHours"]), rules["MaxShiftHours"])

                    # Build window with lunch pad
                    s_str, e_str, lunch_pad = anchored_shift_window(stype, paid_hours, anchors_map, rules)

                    # Mark candidate paid hours for cap check inside can_assign()
                    rules["_candidate_paid_hours"] = paid_hours

                    if not can_assign(
                        e,
                        role,
                        day_name,
                        s_str,
                        e_str,
                        local_avail,
                        funcs_map,
                        assignments_this_week,
                        rules,
                        assigned_paid_hours_week,
                        days_worked_week,
                        min_hours_map,
                        weekly_max_paid_hours_map,
                    ):
                        continue

                    night_flag = is_night_shift(stype, e_str, night_types, night_end_threshold)

                    # Priority: ensure each employee gets >=1 night per week when possible
                    needs_night = 1 if (night_flag and nights_worked_week[e] == 0) else 0
                    under_min = 1 if assigned_paid_hours_week[e] < min_hours_map.get(e, 0) else 0

                    candidates.append(
                        (needs_night, under_min, -assigned_paid_hours_week[e], -len(days_worked_week[e]),
                         e, s_str, e_str, paid_hours, lunch_pad, role, night_flag, stype)
                    )

                if not candidates:
                    assignments_this_week.append(
                        {
                            "Week": w_idx + 1,
                            "Date": dt.strftime("%Y-%m-%d"),
                            "Day": day_name,
                            "Employee": "UNFILLED",
                            "EmployeeID": "",
                            "Start": "",
                            "End": "",
                            "Role": role,
                            "ShiftType": stype,
                            "PaidHours": 0.0,
                            "ScheduledHours": 0.0,
                            "Night": False,
                        }
                    )
                    continue

                candidates.sort(reverse=True)
                _, _, _, _, e, s_str, e_str, paid_hours, lunch_pad, role, night_flag, stype = candidates[0]

                assignments_this_week.append(
                    {
                        "Week": w_idx + 1,
                        "Date": dt.strftime("%Y-%m-%d"),
                        "Day": day_name,
                        "Employee": e,
                        "EmployeeID": emp_id_map.get(e, ""),
                        "Start": s_str,
                        "End": e_str,
                        "Role": role,
                        "ShiftType": stype,
                        "PaidHours": paid_hours,
                        "ScheduledHours": paid_hours + lunch_pad,
                        "Night": bool(night_flag),
                    }
                )
                assigned_paid_hours_week[e] += paid_hours
                days_worked_week[e].add(day_name)
                if night_flag:
                    nights_worked_week[e] += 1

        assignments_all.extend(assignments_this_week)

    out = pd.DataFrame(assignments_all)

    # ------------ PREVIEW + CSV -------------
    st.subheader("Week 1 Preview (on-screen)")
    wk1 = out[out["Week"] == 1].copy()
    if wk1.empty:
        st.info("No assignments produced for Week 1. Check coverage/skills/availability.")
    else:
        emps = sorted([e for e in wk1["Employee"].dropna().unique() if e != "UNFILLED"])
        mat = pd.DataFrame({"Employee": emps})
        for d in DAYS:
            mat[d] = ""
        for _, r in wk1.iterrows():
            e, d = r["Employee"], r["Day"]
            if e == "UNFILLED":
                continue
            text = f"{fmt_time(r['Start'])} - {fmt_time(r['End'])} {r['Role']}"
            ridx = mat.index[mat["Employee"] == e][0]
            mat.at[ridx, d] = (mat.at[ridx, d] + "\n" if mat.at[ridx, d] else "") + text
        st.dataframe(mat, use_container_width=True)

    csv_bytes = out.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download assignments_detailed.csv",
        data=csv_bytes,
        file_name="assignments_detailed.csv",
        mime="text/csv",
    )

    # ------------ Build Hours Summary (per week + total) -------------
    hours_paid = out.groupby(["Week", "Employee"], dropna=False)["PaidHours"].sum().reset_index()
    hours_sched = out.groupby(["Week", "Employee"], dropna=False)["ScheduledHours"].sum().reset_index().rename(
        columns={"ScheduledHours": "ScheduledHoursSum"}
    )
    hours_summary = hours_paid.merge(hours_sched, on=["Week", "Employee"], how="left")

    caps_df = pd.DataFrame(
        {"Employee": list(weekly_max_paid_hours_map.keys()), "MaxHoursPerWeek_Configured": list(weekly_max_paid_hours_map.values())}
    )
    hours_summary = hours_summary.merge(caps_df, on="Employee", how="left")

    pivot_paid = hours_summary.pivot(index="Employee", columns="Week", values="PaidHours").fillna(0)
    pivot_sched = hours_summary.pivot(index="Employee", columns="Week", values="ScheduledHoursSum").fillna(0)
    for wk in range(1, 5):
        if wk not in pivot_paid.columns:
            pivot_paid[wk] = 0.0
        if wk not in pivot_sched.columns:
            pivot_sched[wk] = 0.0

    hours_pivot = pd.DataFrame({"Employee": pivot_paid.index})
    for wk in [1, 2, 3, 4]:
        hours_pivot[f"Week{wk}_PaidHours"] = pivot_paid[wk].values
    for wk in [1, 2, 3, 4]:
        hours_pivot[f"Week{wk}_ScheduledHours"] = pivot_sched[wk].values

    hours_pivot = hours_pivot.merge(caps_df, on="Employee", how="left")
    hours_pivot["PaidHours_All_4_Weeks"] = hours_pivot[[f"Week{wk}_PaidHours" for wk in [1, 2, 3, 4]]].sum(axis=1)
    hours_pivot["ScheduledHours_All_4_Weeks"] = hours_pivot[[f"Week{wk}_ScheduledHours" for wk in [1, 2, 3, 4]]].sum(axis=1)

    # ------------ Build Paycom Import Template -------------
    # Two lines per shift: In (ID) and Out (OD); Paycom date MM/DD/YYYY
    import_rows = []
    for _, r in out.iterrows():
        if r["Employee"] in ("", "UNFILLED"):
            continue

        empid = str(r.get("EmployeeID", "")).strip()
        date_dt = pd.to_datetime(r["Date"]).to_pydatetime()
        date_str = date_dt.strftime("%m/%d/%Y")
        start_4 = hhmm_to_4dig(str(r["Start"]))
        end_4 = hhmm_to_4dig(str(r["End"]))
        func = r["Role"]

        # In punch
        import_rows.append({
            "A_EmployeeID": empid, "B_Blank": "", "C_Date": date_str, "D_Time4": start_4,
            "E_PunchType": "ID", "F_Blank": "", "G_Blank": "", "H_Function": func
        })
        # Out punch
        import_rows.append({
            "A_EmployeeID": empid, "B_Blank": "", "C_Date": date_str, "D_Time4": end_4,
            "E_PunchType": "OD", "F_Blank": "", "G_Blank": "", "H_Function": func
        })

    import_df = pd.DataFrame(import_rows, columns=[
        "A_EmployeeID", "B_Blank", "C_Date", "D_Time4", "E_PunchType", "F_Blank", "G_Blank", "H_Function"
    ])

    # ------------ Excel Export -------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 1})
        name_fmt = workbook.add_format({"bold": True, "align": "left", "valign": "vcenter", "border": 1})
        cell_fmt = workbook.add_format({"text_wrap": True, "align": "left", "valign": "top", "border": 1})
        off_fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "font_color": "#666666"})
        default_role_colors = {"Cashier": "#E3F2FD", "Donation": "#E8F5E9", "Pricer": "#FFF3E0", "Hanger": "#F3E5F5", "Manager": "#FFEBEE"}
        role_colors = default_role_colors
        role_formats = {
            role: workbook.add_format({"text_wrap": True, "align": "left", "valign": "top", "border": 1, "bg_color": color})
            for role, color in role_colors.items()
        }

        # Detailed tab first (include EmployeeID)
        out.to_excel(writer, index=False, sheet_name="assignments_detailed")

        for wk in range(1, 5):
            wkdf = out[out["Week"] == wk].copy()
            sheetname = f"Week {wk}"
            if wkdf.empty:
                empty = pd.DataFrame({"Employee": []})
                empty.to_excel(writer, index=False, sheet_name=sheetname, startrow=1)
                ws = writer.sheets[sheetname]
                ws.merge_range(0, 0, 0, 8, f"Week {wk} (no assignments)", header_fmt)
                continue

            start_date_str = wkdf["Date"].min()
            end_date_str = wkdf["Date"].max()
            title = f"Week {wk} ({start_date_str} to {end_date_str})"

            emps = sorted([e for e in wkdf["Employee"].dropna().unique() if e != "UNFILLED"])
            frame = pd.DataFrame({"Employee": emps}, dtype=object)
            for d in DAYS:
                frame[d] = ""

            role_map = {}

            for _, rr in wkdf.iterrows():
                e = rr["Employee"]
                if e == "UNFILLED" or pd.isna(e):
                    continue
                d = rr["Day"]
                text = f"{fmt_time(rr['Start'])}-{fmt_time(rr['End'])} {rr['Role']}"
                ridx = frame.index[frame["Employee"] == e][0]
                existing = frame.at[ridx, d]
                frame.at[ridx, d] = (existing + "\n" if existing else "") + text
                if (ridx, d) not in role_map:
                    role_map[(ridx, d)] = rr["Role"]

            frame.to_excel(writer, index=False, sheet_name=sheetname, startrow=1)
            ws = writer.sheets[sheetname]
            ws.merge_range(0, 0, 0, 8, title, header_fmt)
            ws.set_column(0, 0, 24)
            ws.set_column(1, 7, 22)

            # Headers with dates
            for col_idx, d in enumerate(DAYS, start=1):
                drows = wkdf[wkdf["Day"] == d]
                if not drows.empty:
                    dt = pd.to_datetime(drows["Date"].iloc[0]).date().strftime("%m/%d")
                    label = f"{d} {dt}"
                else:
                    label = d
                ws.write(1, col_idx, label, header_fmt)
            ws.write(1, 0, "Employee", header_fmt)

            ws.set_landscape()
            ws.fit_to_pages(1, 0)
            ws.repeat_rows(1, 1)
            ws.center_horizontally()

            # Safe writes
            for r in range(2, 2 + len(frame)):
                emp_val = frame.iloc[r - 2]["Employee"]
                if pd.isna(emp_val):
                    emp_val = ""
                ws.write_string(r, 0, str(emp_val), name_fmt)
                for c, d in enumerate(DAYS, start=1):
                    val = frame.iloc[r - 2][d]
                    if pd.isna(val) or val == "":
                        ws.write_string(r, c, "OFF", off_fmt)
                    else:
                        fmt = role_formats.get(role_map.get((r - 2, d), None), cell_fmt)
                        ws.write_string(r, c, str(val), fmt)

        # Hours Summary tab
        hours_pivot.to_excel(writer, index=False, sheet_name="hours_summary")
        ws_sum = writer.sheets["hours_summary"]
        ws_sum.set_column(0, 0, 26)
        ws_sum.set_column(1, 12, 18)
        for ci, col in enumerate(hours_pivot.columns):
            ws_sum.write(0, ci, str(col), header_fmt)

        # Paycom Import Template tab (A..H)
        import_df.to_excel(writer, index=False, sheet_name="import_template")
        ws_imp = writer.sheets["import_template"]
        ws_imp.set_column(0, 7, 18)

    output.seek(0)
    return output, out, hours_pivot


# ------------------ Template Generator (LOCKED coverage; EmployeeID next to Employee) ------------------
def generate_template_bytes():
    # Employees: headers only (IDs blank for managers to fill)
    employees_df = pd.DataFrame(
        {
            "Employee": [],
            "EmployeeID": [],
            "MaxHoursPerWeek": [],
            "MinHours": [],
            "PreferredShiftHours": [],
            "Cashier": [],
            "Donation": [],
            "Pricer": [],
            "Hanger": [],
            "Manager": [],
        }
    )

    availability_simple_df = pd.DataFrame(
        {
            "Employee": [],
            "Sunday": [],
            "Monday": [],
            "Tuesday": [],
            "Wednesday": [],
            "Thursday": [],
            "Friday": [],
            "Saturday": [],
        }
    )

    # LOCKED weekly coverage (Role, Day, ShiftType, Count)
    coverage_rows = [
        # Cashier
        {"Role":"Cashier","Day":"Friday","ShiftType":"Close","Count":1},
        {"Role":"Cashier","Day":"Friday","ShiftType":"Mid","Count":1},
        {"Role":"Cashier","Day":"Friday","ShiftType":"Open","Count":1},
        {"Role":"Cashier","Day":"Monday","ShiftType":"Close","Count":1},
        {"Role":"Cashier","Day":"Monday","ShiftType":"Mid","Count":1},
        {"Role":"Cashier","Day":"Monday","ShiftType":"Open","Count":1},
        {"Role":"Cashier","Day":"Saturday","ShiftType":"Close","Count":1},
        {"Role":"Cashier","Day":"Saturday","ShiftType":"Mid","Count":2},
        {"Role":"Cashier","Day":"Saturday","ShiftType":"Open","Count":1},
        {"Role":"Cashier","Day":"Sunday","ShiftType":"Close","Count":1},
        {"Role":"Cashier","Day":"Sunday","ShiftType":"Open","Count":1},
        {"Role":"Cashier","Day":"Thursday","ShiftType":"Close","Count":1},
        {"Role":"Cashier","Day":"Thursday","ShiftType":"Mid","Count":1},
        {"Role":"Cashier","Day":"Thursday","ShiftType":"Open","Count":1},
        {"Role":"Cashier","Day":"Tuesday","ShiftType":"Close","Count":1},
        {"Role":"Cashier","Day":"Tuesday","ShiftType":"Mid","Count":1},
        {"Role":"Cashier","Day":"Tuesday","ShiftType":"Open","Count":1},
        {"Role":"Cashier","Day":"Wednesday","ShiftType":"Close","Count":1},
        {"Role":"Cashier","Day":"Wednesday","ShiftType":"Mid","Count":1},
        {"Role":"Cashier","Day":"Wednesday","ShiftType":"Open","Count":1},
        # Donation
        {"Role":"Donation","Day":"Friday","ShiftType":"Close","Count":1},
        {"Role":"Donation","Day":"Friday","ShiftType":"Mid","Count":1},
        {"Role":"Donation","Day":"Friday","ShiftType":"Open","Count":1},
        {"Role":"Donation","Day":"Monday","ShiftType":"Close","Count":1},
        {"Role":"Donation","Day":"Monday","ShiftType":"Mid","Count":1},
        {"Role":"Donation","Day":"Monday","ShiftType":"Open","Count":1},
        {"Role":"Donation","Day":"Saturday","ShiftType":"Close","Count":1},
        {"Role":"Donation","Day":"Saturday","ShiftType":"Mid","Count":2},
        {"Role":"Donation","Day":"Saturday","ShiftType":"Open","Count":1},
        {"Role":"Donation","Day":"Sunday","ShiftType":"Close","Count":1},
        {"Role":"Donation","Day":"Sunday","ShiftType":"Mid","Count":1},
        {"Role":"Donation","Day":"Sunday","ShiftType":"Open","Count":1},
        {"Role":"Donation","Day":"Thursday","ShiftType":"Close","Count":1},
        {"Role":"Donation","Day":"Thursday","ShiftType":"Mid","Count":1},
        {"Role":"Donation","Day":"Thursday","ShiftType":"Open","Count":1},
        {"Role":"Donation","Day":"Tuesday","ShiftType":"Close","Count":1},
        {"Role":"Donation","Day":"Tuesday","ShiftType":"Mid","Count":1},
        {"Role":"Donation","Day":"Tuesday","ShiftType":"Open","Count":1},
        {"Role":"Donation","Day":"Wednesday","ShiftType":"Close","Count":1},
        {"Role":"Donation","Day":"Wednesday","ShiftType":"Mid","Count":1},
        {"Role":"Donation","Day":"Wednesday","ShiftType":"Open","Count":1},
        # Hanger
        {"Role":"Hanger","Day":"Friday","ShiftType":"Close","Count":1},
        {"Role":"Hanger","Day":"Friday","ShiftType":"Mid","Count":1},
        {"Role":"Hanger","Day":"Friday","ShiftType":"Open","Count":2},
        {"Role":"Hanger","Day":"Monday","ShiftType":"Close","Count":1},
        {"Role":"Hanger","Day":"Monday","ShiftType":"Mid","Count":1},
        {"Role":"Hanger","Day":"Monday","ShiftType":"Open","Count":2},
        {"Role":"Hanger","Day":"Saturday","ShiftType":"Close","Count":1},
        {"Role":"Hanger","Day":"Saturday","ShiftType":"Mid","Count":1},
        {"Role":"Hanger","Day":"Saturday","ShiftType":"Open","Count":2},
        {"Role":"Hanger","Day":"Sunday","ShiftType":"Close","Count":1},
        {"Role":"Hanger","Day":"Sunday","ShiftType":"Mid","Count":1},
        {"Role":"Hanger","Day":"Sunday","ShiftType":"Open","Count":2},
        {"Role":"Hanger","Day":"Thursday","ShiftType":"Close","Count":1},
        {"Role":"Hanger","Day":"Thursday","ShiftType":"Mid","Count":1},
        {"Role":"Hanger","Day":"Thursday","ShiftType":"Open","Count":2},
        {"Role":"Hanger","Day":"Tuesday","ShiftType":"Close","Count":1},
        {"Role":"Hanger","Day":"Tuesday","ShiftType":"Mid","Count":1},
        {"Role":"Hanger","Day":"Tuesday","ShiftType":"Open","Count":2},
        {"Role":"Hanger","Day":"Wednesday","ShiftType":"Close","Count":1},
        {"Role":"Hanger","Day":"Wednesday","ShiftType":"Mid","Count":1},
        {"Role":"Hanger","Day":"Wednesday","ShiftType":"Open","Count":2},
        # Manager
        {"Role":"Manager","Day":"Friday","ShiftType":"Close","Count":1},
        {"Role":"Manager","Day":"Friday","ShiftType":"Open","Count":1},
        {"Role":"Manager","Day":"Monday","ShiftType":"Close","Count":1},
        {"Role":"Manager","Day":"Monday","ShiftType":"Open","Count":1},
        {"Role":"Manager","Day":"Saturday","ShiftType":"Close","Count":1},
        {"Role":"Manager","Day":"Saturday","ShiftType":"Open","Count":1},
        {"Role":"Manager","Day":"Sunday","ShiftType":"Close","Count":1},
        {"Role":"Manager","Day":"Sunday","ShiftType":"Open","Count":1},
        {"Role":"Manager","Day":"Thursday","ShiftType":"Close","Count":1},
        {"Role":"Manager","Day":"Thursday","ShiftType":"Open","Count":1},
        {"Role":"Manager","Day":"Tuesday","ShiftType":"Close","Count":1},
        {"Role":"Manager","Day":"Tuesday","ShiftType":"Open","Count":1},
        {"Role":"Manager","Day":"Wednesday","ShiftType":"Close","Count":1},
        {"Role":"Manager","Day":"Wednesday","ShiftType":"Open","Count":1},
        # Pricer
        {"Role":"Pricer","Day":"Friday","ShiftType":"Close","Count":2},
        {"Role":"Pricer","Day":"Friday","ShiftType":"Open","Count":2},
        {"Role":"Pricer","Day":"Monday","ShiftType":"Close","Count":2},
        {"Role":"Pricer","Day":"Monday","ShiftType":"Open","Count":2},
        {"Role":"Pricer","Day":"Saturday","ShiftType":"Close","Count":2},
        {"Role":"Pricer","Day":"Saturday","ShiftType":"Open","Count":2},
        {"Role":"Pricer","Day":"Sunday","ShiftType":"Close","Count":1},
        {"Role":"Pricer","Day":"Sunday","ShiftType":"Open","Count":2},
        {"Role":"Pricer","Day":"Thursday","ShiftType":"Close","Count":2},
        {"Role":"Pricer","Day":"Thursday","ShiftType":"Open","Count":2},
        {"Role":"Pricer","Day":"Tuesday","ShiftType":"Close","Count":2},
        {"Role":"Pricer","Day":"Tuesday","ShiftType":"Open","Count":2},
        {"Role":"Pricer","Day":"Wednesday","ShiftType":"Close","Count":2},
        {"Role":"Pricer","Day":"Wednesday","ShiftType":"Open","Count":2},
    ]
    coverage_df = pd.DataFrame(coverage_rows)[["Role", "Day", "ShiftType", "Count"]]

    anchors_df = pd.DataFrame({
        "ShiftType": ["Open", "Mid", "Close"],
        "Anchor": ["Start", "Start", "End"],
        "Time": ["08:00", "10:00", "20:30"]
    })

    rules_df = pd.DataFrame({
        "Rule": [
            "MinShiftHours", "MaxShiftHours", "MinGapBetweenShiftsHours", "MaxDaysPerWeek",
            "AllowSplitShifts", "LunchTriggerHours", "LunchMinutes",
            "NightShiftTypes", "NightShiftEndAtOrAfter"
        ],
        "Value": [3, 10, 1, 6, "No", 5, 30, "Close", ""]
    })

    # Empty import template tab with headers only (Paycom)
    import_template_df = pd.DataFrame(columns=[
        "A_EmployeeID", "B_Blank", "C_Date", "D_Time4", "E_PunchType", "F_Blank", "G_Blank", "H_Function"
    ])

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        employees_df.to_excel(writer, index=False, sheet_name="employees")
        availability_simple_df.to_excel(writer, index=False, sheet_name="availability_simple")
        # LOCKED coverage on a sheet named 'coverage'
        coverage_df.to_excel(writer, index=False, sheet_name="coverage")
        anchors_df.to_excel(writer, index=False, sheet_name="shift_anchors")
        rules_df.to_excel(writer, index=False, sheet_name="rules")
        import_template_df.to_excel(writer, index=False, sheet_name="import_template")

        wb = writer.book
        ws_emp = writer.sheets["employees"]
        ws_avs = writer.sheets["availability_simple"]
        ws_cov = writer.sheets["coverage"]
        ws_imp = writer.sheets["import_template"]
        yesno = ["Yes", "No"]

        # Add data validation dropdowns to role columns (if managers add rows)
        start_row = 1
        end_row = start_row + 300
        for role in ["Cashier", "Donation", "Pricer", "Hanger", "Manager"]:
            if role in employees_df.columns:
                idx = list(employees_df.columns).index(role)
                ws_emp.data_validation(first_row=start_row, first_col=idx, last_row=end_row, last_col=idx,
                                       options={"validate": "list", "source": yesno})

        # PreferredShiftHours numeric min/max if present
        if "PreferredShiftHours" in employees_df.columns:
            pref_idx = list(employees_df.columns).index("PreferredShiftHours")
            ws_emp.data_validation(first_row=start_row, first_col=pref_idx, last_row=end_row, last_col=pref_idx,
                                   options={"validate":"integer","criteria":">=","value":3})
            ws_emp.data_validation(first_row=start_row, first_col=pref_idx, last_row=end_row, last_col=pref_idx,
                                   options={"validate":"integer","criteria":"<=","value":10})

        # availability_simple dropdowns
        start_row_av = 1
        end_row_av = start_row_av + 600
        for d in DAYS:
            if d in availability_simple_df.columns:
                col_idx = list(availability_simple_df.columns).index(d)
                ws_avs.data_validation(first_row=start_row_av, first_col=col_idx, last_row=end_row_av, last_col=col_idx,
                                       options={"validate":"list","source":yesno})

        # widen columns for readability
        ws_emp.set_column(0, len(employees_df.columns)-1, 18)
        ws_cov.set_column(0, 3, 16)
        ws_imp.set_column(0, 7, 18)

    bio.seek(0)
    return bio.read()


# ------------------ UI ------------------
with st.sidebar:
    st.header("Template")
    st.download_button(
        label="Download blank template (LOCKED coverage)",
        data=generate_template_bytes(),
        file_name="schedule_input_template_v_locked_coverage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(
        "Template includes: employees (Employee + EmployeeID), availability_simple (optional), "
        "coverage (pre-filled Role/Day/ShiftType/Count you provided), shift_anchors, rules, and an empty import_template."
    )

uploaded = st.file_uploader("Upload your filled template (xlsx)", type=["xlsx"])
startdate = st.text_input("Start date (YYYY-MM-DD, optional; defaults to upcoming Sunday)", "")

if st.button("Generate 4-week schedule", type="primary"):
    if not uploaded:
        st.error("Please upload the input Excel first.")
    else:
        try:
            schedule_file, detailed, hours_pivot = run_schedule_with_summary(uploaded, startdate.strip() or None)
            st.success("Schedule generated! Download below.")
            st.download_button(
                "Download schedule_output_calendar_weeks.xlsx",
                data=schedule_file,
                file_name="schedule_output_calendar_weeks.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.subheader("Hours Summary (Paid vs Scheduled per week)")
            st.dataframe(hours_pivot, use_container_width=True)
        except Exception as e:
            st.exception(e)
  
