import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
import io

st.set_page_config(page_title="Goodwill Scheduler", layout="wide")
st.title("Goodwill Gulf Coast — Weekly Scheduler (weekly hour caps)")
st.caption("Upload the input Excel, choose an optional start date, and download a 4-week schedule. This build enforces MaxHours **per week** and adds an Hours Summary tab to verify.")

DAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
KNOWN_ROLES = ["Cashier","Donation","Pricer","Hanger","Manager"]

# ------------------ Loaders ------------------
def load_inputs(file):
    xls = pd.ExcelFile(file)
    employees = pd.read_excel(xls, "employees").fillna("")
    availability_simple = pd.read_excel(xls, "availability_simple").fillna("") if "availability_simple" in xls.sheet_names else None
    availability = pd.read_excel(xls, "availability").fillna("") if "availability" in xls.sheet_names else None
    coverage = pd.read_excel(xls, "coverage_template").fillna("")
    anchors = pd.read_excel(xls, "shift_anchors").fillna("")
    rules = pd.read_excel(xls, "rules").fillna("")
    # Normalize day names
    if "Day" in coverage.columns:
        coverage["Day"] = coverage["Day"].astype(str).str.strip().str.title()
    return employees, availability, availability_simple, coverage, anchors, rules

def build_availability_map(employees, availability, availability_simple):
    # Return dict[employee][day] -> list of (start,end) windows in HH:MM.
    # If no availability provided, assume full 08:00-20:30 daily.
    avail = {}
    # Simple Yes/No per day: Yes = NOT available (per the template's caption)
    if availability_simple is not None and not availability_simple.empty:
        df = availability_simple.copy()
        df.columns = [("Employee" if c=="Employee" else str(c).strip().title()) for c in df.columns]
        for _, row in df.iterrows():
            emp = row["Employee"]
            for d in DAYS:
                resp = str(row.get(d, "")).strip().lower()
                if resp in ("yes","y","true","1","x"):
                    avail.setdefault(emp, {}).setdefault(d, [])
                else:
                    avail.setdefault(emp, {}).setdefault(d, []).append(("08:00","20:30"))
        return avail
    # Detailed availability windows
    if availability is not None and not availability.empty:
        df = availability.copy()
        df["Day"] = df["Day"].astype(str).str.strip().str.title()
        for _, row in df.iterrows():
            emp = row["Employee"]; d = row["Day"]
            start = str(row.get("Start",""))).strip(); end = str(row.get("End",""))).strip()
            if start and end:
                avail.setdefault(emp, {}).setdefault(d, []).append((start, end))
        return avail
    # Default: everyone fully available
    for emp in employees["Employee"]:
        for d in DAYS:
            avail.setdefault(emp, {}).setdefault(d, []).append(("08:00","20:30"))
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
                if val in ("yes","y","true","1","x"):
                    skills.add(role)
            if not skills and "Functions" in employees.columns:
                raw = str(r.get("Functions",""))
                for token in raw.replace(";",",").replace("/",",").split(","):
                    t = token.strip().title()
                    if t in KNOWN_ROLES: skills.add(t)
            m[emp] = skills
        return m
    # CSV fallback only
    for _, r in employees.iterrows():
        emp = r["Employee"]
        skills = set()
        raw = str(r.get("Functions",""))
        for token in raw.replace(";",",").replace("/",",").split(","):
            t = token.strip().title()
            if t in KNOWN_ROLES: skills.add(t)
        m[emp] = skills
    return m

def weekend_rotation(employees):
    emps = list(employees["Employee"])
    groups = {i:set() for i in range(4)}
    for i, e in enumerate(emps):
        groups[i % 4].add(e)
    return groups

def generate_week_dates(start_date: date):
    # Align to Sunday
    start = start_date - timedelta(days=start_date.weekday()+1 if start_date.weekday()!=6 else 0)
    return [[start + timedelta(days=7*w + i) for i in range(7)] for w in range(4)]

def anchored_shift_window(shift_type, hours, anchors_map):
    a = anchors_map[shift_type]
    if a["Anchor"] == "Start":
        stt = datetime.strptime(a["Time"], "%H:%M")
        en = stt + timedelta(hours=float(hours))
        return (a["Time"], en.strftime("%H:%M"))
    else:
        en = datetime.strptime(a["Time"], "%H:%M")
        stt = en - timedelta(hours=float(hours))
        return (stt.strftime("%H:%M"), a["Time"])

def hours_between(start_str, end_str):
    return (datetime.strptime(end_str, "%H:%M") - datetime.strptime(start_str, "%H:%M")).seconds/3600

def can_assign(emp, role, day_name, start_str, end_str, avail_map, funcs_map, assigned_list_for_week, rules, assigned_hours_week, days_worked_week, min_hours_map, weekly_max_hours_map):
    if role not in funcs_map.get(emp, set()): return False
    windows = avail_map.get(emp, {}).get(day_name, [])
    if not any(a_start <= start_str and a_end >= end_str for a_start, a_end in windows): return False
    # overlap + gap
    min_gap = float(rules.get("MinGapBetweenShiftsHours", 1))
    def to_dt_local(t): return datetime.combine(datetime.today(), datetime.strptime(t, "%H:%M").time())
    slot_s, slot_e = to_dt_local(start_str), to_dt_local(end_str)
    for a in assigned_list_for_week:
        if a["Employee"]!=emp or a["Day"]!=day_name: continue
        s, e = to_dt_local(a["Start"]), to_dt_local(a["End"])
        if not (slot_e <= s or slot_s >= e): return False
        if 0 <= (s - slot_e).total_seconds()/3600 < min_gap or 0 <= (slot_s - e).total_seconds()/3600 < min_gap: return False
    # **Weekly** max hours constraint
    sh = hours_between(start_str, end_str)
    if assigned_hours_week.get(emp, 0) + sh > weekly_max_hours_map.get(emp, float('inf')):
        return False
    if len(days_worked_week.get(emp, set())) >= int(rules.get("MaxDaysPerWeek", 6)) and day_name not in days_worked_week.get(emp,set()):
        return False
    return True

# ------------------ Scheduler ------------------
def run_schedule(file, startdate_str=None, role_colors=None):
    employees, availability, availability_simple, coverage, anchors, rules_df = load_inputs(file)
    rules = dict(zip(rules_df["Rule"], rules_df["Value"]))
    # defaults
    rules["MinShiftHours"] = float(rules.get("MinShiftHours", 3))
    rules["MaxShiftHours"] = float(rules.get("MaxShiftHours", 10))
    rules["MinGapBetweenShiftsHours"] = float(rules.get("MinGapBetweenShiftsHours", 1))
    rules["MaxDaysPerWeek"] = int(rules.get("MaxDaysPerWeek", 6))
    anchors_map = {r["ShiftType"]: {"Anchor": r["Anchor"], "Time": r["Time"]} for _, r in anchors.iterrows()}

    # If a MaxHoursPerWeek column exists, prefer it; else fall back to MaxHours
    if "MaxHoursPerWeek" in employees.columns:
        weekly_max_series = pd.to_numeric(employees["MaxHoursPerWeek"], errors="coerce")
    else:
        weekly_max_series = pd.to_numeric(employees.get("MaxHours", pd.Series([999]*len(employees))), errors="coerce")
    weekly_max_series = weekly_max_series.fillna(999)

    pref_hours = dict(zip(
        employees["Employee"],
        pd.to_numeric(employees.get("PreferredShiftHours", pd.Series([rules["MaxShiftHours"]]*len(employees))), errors="coerce").fillna(rules["MaxShiftHours"]) 
    ))
    funcs_map = employee_functions_map(employees)
    avail_map = build_availability_map(employees, availability, availability_simple)

    # weekly min/max hours maps (per employee)
    min_hours_map = dict(zip(employees["Employee"], pd.to_numeric(employees.get("MinHours", 0), errors="coerce").fillna(0)))
    weekly_max_hours_map = dict(zip(employees["Employee"], weekly_max_series))

    # start date
    if startdate_str:
        startdate = datetime.strptime(startdate_str, "%Y-%m-%d").date()
    else:
        today = datetime.today().date()
        delta = (6 - today.weekday()) % 7
        startdate = today + timedelta(days=delta)

    weeks = generate_week_dates(startdate)
    weekend_groups = weekend_rotation(employees)

    # Keep assignments for all weeks, but we also keep a per-week view when evaluating caps
    assignments_all = []

    for w_idx, week_dates in enumerate(weeks):
        # reset **weekly** tracking
        assigned_hours_week = {e:0 for e in employees["Employee"]}
        days_worked_week = {e:set() for e in employees["Employee"]}
        assignments_this_week = []

        # rotating Sat/Sun off-group
        off_group = weekend_groups[w_idx]
        local_avail = {e:{d:list(w) for d,w in days.items()} for e,days in avail_map.items()}
        for e in off_group:
            for d in ["Saturday","Sunday"]:
                if e in local_avail and d in local_avail[e]:
                    local_avail[e][d] = []

        for d_idx, dt in enumerate(week_dates):
            day_name = DAYS[d_idx]
            day_cov = coverage[coverage["Day"]==day_name].copy()

            # expand demand
            demands = []
            for _, r in day_cov.iterrows():
                count = int(pd.to_numeric(r.get("Count", 0), errors="coerce") or 0)
                role = r.get("Role", "")
                stype = r.get("ShiftType", "")
                demands += [{"Role": role, "ShiftType": stype} for _ in range(count)]

            # fill rare roles first
            rarity = {}
            for role in day_cov["Role"].unique():
                count = sum(1 for e in employees["Employee"] if role in funcs_map.get(e,set()))
                rarity[role] = 1.0/(count if count>0 else 0.5)
            demands.sort(key=lambda d: rarity.get(d["Role"], 1.0), reverse=True)

            for dem in demands:
                role, stype = dem["Role"], dem["ShiftType"]
                candidates = []
                for e in employees["Employee"]:
                    if role not in funcs_map.get(e,set()):
                        continue
                    hours = float(pref_hours.get(e, rules["MaxShiftHours"]))
                    hours = min(max(hours, rules["MinShiftHours"]), rules["MaxShiftHours"])
                    s_str, e_str = anchored_shift_window(stype, hours, anchors_map)
                    if not can_assign(e, role, day_name, s_str, e_str, local_avail, funcs_map, assignments_this_week, rules, assigned_hours_week, days_worked_week, min_hours_map, weekly_max_hours_map):
                        continue
                    # prefer under weekly MinHours, then fewer hours, then fewer days
                    under_min = 1 if assigned_hours_week[e] < min_hours_map.get(e,0) else 0
                    candidates.append((under_min, -assigned_hours_week[e], -len(days_worked_week[e]), e, s_str, e_str, role))
                if not candidates:
                    assignments_this_week.append({"Week": w_idx+1, "Date": dt.strftime("%Y-%m-%d"), "Day": day_name, "Employee": "UNFILLED", "Start": "", "End": "", "Role": role})
                    continue
                candidates.sort(reverse=True)
                _,_,_, e, s_str, e_str, role = candidates[0]
                assignments_this_week.append({"Week": w_idx+1, "Date": dt.strftime("%Y-%m-%d"), "Day": day_name, "Employee": e, "Start": s_str, "End": e_str, "Role": role})
                sh = hours_between(s_str, e_str)
                assigned_hours_week[e] += sh
                days_worked_week[e].add(day_name)
        # merge week into all
        assignments_all.extend(assignments_this_week)

    out = pd.DataFrame(assignments_all)

    # ------------ PREVIEW + CSV -------------
    st.subheader("Week 1 Preview (on-screen)")
    wk1 = out[out["Week"]==1].copy()
    if wk1.empty:
        st.info("No assignments produced for Week 1. Check coverage/skills/availability.")
    else:
        emps = sorted([e for e in wk1["Employee"].dropna().unique() if e != "UNFILLED"])
        mat = pd.DataFrame({"Employee": emps})
        for d in DAYS: mat[d] = ""
        for _, r in wk1.iterrows():
            e, d = r["Employee"], r["Day"]
            if e == "UNFILLED": continue
            text = f"{datetime.strptime(r['Start'],'%H:%M').strftime('%-I:%M %p')} - {datetime.strptime(r['End'],'%H:%M').strftime('%-I:%M %p')} {r['Role']}"
            ridx = mat.index[mat["Employee"]==e][0]
            mat.at[ridx, d] = (mat.at[ridx, d] + "\n" if mat.at[ridx, d] else "") + text
        st.dataframe(mat, use_container_width=True)

    csv_bytes = out.to_csv(index=False).encode("utf-8")
    st.download_button("Download assignments_detailed.csv", data=csv_bytes, file_name="assignments_detailed.csv", mime="text/csv")

    # ------------ Build Hours Summary (per week + total) -------------
    def _hours_col(r):
        if not r["Start"] or not r["End"] or r["Employee"] in ("", "UNFILLED"):
            return 0.0
        return hours_between(r["Start"], r["End"]) 

    out["Hours"] = out.apply(_hours_col, axis=1)
    hours_summary = out.groupby(["Week","Employee"], dropna=False)["Hours"].sum().reset_index()
    total_summary = out.groupby(["Employee"], dropna=False)["Hours"].sum().reset_index().rename(columns={"Hours":"Hours_All_4_Weeks"})
    # Add the configured weekly cap to the summary for easy comparison
    weekly_cap_df = pd.DataFrame({"Employee": list(weekly_max_hours_map.keys()), "MaxHoursPerWeek_Configured": list(weekly_max_hours_map.values())})
    hours_summary = hours_summary.merge(weekly_cap_df, on="Employee", how="left")
    hours_pivot = hours_summary.pivot(index="Employee", columns="Week", values="Hours").fillna(0)
    for wk in range(1,5):
        if wk not in hours_pivot.columns:
            hours_pivot[wk] = 0.0
    hours_pivot = hours_pivot[[1,2,3,4]].reset_index().rename(columns={1:"Week1_Hours",2:"Week2_Hours",3:"Week3_Hours",4:"Week4_Hours"})
    hours_pivot = hours_pivot.merge(weekly_cap_df, on="Employee", how="left").merge(total_summary, on="Employee", how="left")

    # ------------ EXCEL EXPORT -------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "align":"center", "valign":"vcenter", "border":1})
        name_fmt = workbook.add_format({"bold": True, "align":"left", "valign":"vcenter", "border":1})
        cell_fmt = workbook.add_format({"text_wrap": True, "align":"left", "valign":"top", "border":1})
        off_fmt = workbook.add_format({"align":"center", "valign":"vcenter", "border":1, "font_color":"#666666"})
        default_role_colors = {"Cashier":"#E3F2FD","Donation":"#E8F5E9","Pricer":"#FFF3E0","Hanger":"#F3E5F5","Manager":"#FFEBEE"}
        role_colors = role_colors or default_role_colors
        role_formats = {role: workbook.add_format({"text_wrap": True, "align":"left", "valign":"top", "border":1, "bg_color": color}) for role, color in role_colors.items()}

        # Detailed tab first
        out.drop(columns=["Hours"], errors="ignore").to_excel(writer, index=False, sheet_name="assignments_detailed")

        for wk in range(1,5):
            wkdf = out[out["Week"]==wk].copy()
            sheetname = f"Week {wk}"
            if wkdf.empty:
                # create labeled empty sheet
                empty = pd.DataFrame({"Employee": []})
                empty.to_excel(writer, index=False, sheet_name=sheetname, startrow=1)
                ws = writer.sheets[sheetname]
                ws.merge_range(0,0,0,8, f"Week {wk} (no assignments)", header_fmt)
                continue

            start_date = wkdf["Date"].min(); end_date = wkdf["Date"].max()
            title = f"Week {wk} ({start_date} to {end_date})"

            # Build matrix
            emps = sorted([e for e in wkdf["Employee"].dropna().unique() if e != "UNFILLED"])
            frame = pd.DataFrame({"Employee": emps}, dtype=object)
            for d in DAYS: frame[d] = ""

            # Track first role in each cell for color
            role_map = {}

            for _, r in wkdf.iterrows():
                e = r["Employee"]
                if e == "UNFILLED" or pd.isna(e): continue
                d = r["Day"]
                stime = datetime.strptime(r["Start"], "%H:%M").strftime("%-I:%M %p")
                etime = datetime.strptime(r["End"], "%H:%M").strftime("%-I:%M %p")
                text = f"{stime}-{etime} {r['Role']}"
                ridx = frame.index[frame["Employee"]==e][0]
                existing = frame.at[ridx, d]
                frame.at[ridx, d] = (existing + "\n" if existing else "") + text
                if (ridx, d) not in role_map:
                    role_map[(ridx, d)] = r["Role"]

            # Write to sheet
            frame.to_excel(writer, index=False, sheet_name=sheetname, startrow=1)
            ws = writer.sheets[sheetname]
            ws.merge_range(0,0,0,8, title, header_fmt)
            ws.set_column(0,0,24); ws.set_column(1,7,22)
            # Headers with dates
            for col_idx, d in enumerate(DAYS, start=1):
                drows = wkdf[wkdf["Day"]==d]
                if not drows.empty:
                    dt = pd.to_datetime(drows["Date"].iloc[0]).date().strftime("%m/%d")
                    label = f"{d} {dt}"
                else:
                    label = d
                ws.write(1, col_idx, label, header_fmt)
            ws.write(1, 0, "Employee", header_fmt)

            # Print setup
            ws.set_landscape(); ws.fit_to_pages(1, 0); ws.repeat_rows(1, 1); ws.center_horizontally()

            # Apply colored cells safely (write_string avoids NaN/INF issues)
            for r in range(2, 2 + len(frame)):
                emp_val = frame.iloc[r-2]["Employee"]
                if pd.isna(emp_val): emp_val = ""
                ws.write_string(r, 0, str(emp_val), name_fmt)
                for c, d in enumerate(DAYS, start=1):
                    val = frame.iloc[r-2][d]
                    if pd.isna(val) or val == "":
                        ws.write_string(r, c, "OFF", off_fmt)
                    else:
                        fmt = role_formats.get(role_map.get((r-2, d), None), cell_fmt)
                        ws.write_string(r, c, str(val), fmt)

        # Hours Summary tab
        hours_pivot.to_excel(writer, index=False, sheet_name="hours_summary")
        ws_sum = writer.sheets["hours_summary"]
        ws_sum.set_column(0, 0, 26)
        ws_sum.set_column(1, 6, 18)
        ws_sum.write(0, 0, "Employee", header_fmt)
        # Add a little header row formatting for clarity
        for ci, col in enumerate(hours_pivot.columns):
            ws_sum.write(0, ci, str(col), header_fmt)

    output.seek(0)
    return output, out

# ------------------ Template Generator ------------------
def generate_template_bytes():
    # Starter template with dropdowns and default rules (Min=3, Max=10). Managers can edit counts/roles.
    employees_df = pd.DataFrame({
        "Employee": ["Example Person 1", "Example Person 2"],
        "MaxHoursPerWeek": [35, 30],
        "MinHours": [15, 12],
        "PreferredShiftHours": [8, 6],
        "Cashier": ["Yes","No"],
        "Donation": ["No","Yes"],
        "Pricer": ["No","Yes"],
        "Hanger": ["No","No"],
        "Manager": ["Yes","No"],
    })
    availability_simple_df = pd.DataFrame({
        "Employee": employees_df["Employee"],
        "Sunday": ["No","No"], "Monday": ["No","No"], "Tuesday": ["No","No"],
        "Wednesday": ["No","No"], "Thursday": ["No","No"], "Friday": ["No","No"], "Saturday": ["No","No"],
    })
    # Put some sample coverage rows; managers should overwrite with their real counts
    sample_cov = [
        {"Role":"Cashier","Day":"Sunday","ShiftType":"Open","Count":1},
        {"Role":"Donation","Day":"Sunday","ShiftType":"Open","Count":1},
        {"Role":"Cashier","Day":"Saturday","ShiftType":"Mid","Count":2},
        {"Role":"Donation","Day":"Saturday","ShiftType":"Mid","Count":2},
    ]
    coverage_df = pd.DataFrame(sample_cov)
    anchors_df = pd.DataFrame({"ShiftType":["Open","Mid","Close"],"Anchor":["Start","Start","End"],"Time":["08:00","10:00","20:30"]})
    rules_df = pd.DataFrame({"Rule":["MinShiftHours","MaxShiftHours","MinGapBetweenShiftsHours","MaxDaysPerWeek","AllowSplitShifts"],
                             "Value":[3,10,1,6,"No"]})

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        employees_df.to_excel(writer, index=False, sheet_name="employees")
        availability_simple_df.to_excel(writer, index=False, sheet_name="availability_simple")
        coverage_df.to_excel(writer, index=False, sheet_name="coverage_template")
        anchors_df.to_excel(writer, index=False, sheet_name="shift_anchors")
        rules_df.to_excel(writer, index=False, sheet_name="rules")

        wb = writer.book
        ws_emp = writer.sheets["employees"]
        ws_avs = writer.sheets["availability_simple"]
        yesno = ["Yes","No"]
        # Add data validation dropdowns to role columns
        start_row = 1; end_row = start_row + 300
        for role in ["Cashier","Donation","Pricer","Hanger","Manager"]:
            idx = employees_df.columns.get_loc(role)
            ws_emp.data_validation(first_row=start_row, first_col=idx, last_row=end_row, last_col=idx,
                                   options={"validate":"list","source":yesno})
        # PreferredShiftHours numeric 3..10
        pref_idx = employees_df.columns.get_loc("PreferredShiftHours")
        ws_emp.data_validation(first_row=start_row, first_col=pref_idx, last_row=end_row, last_col=pref_idx,
                               options={"validate":"integer","criteria":">=","value":3})
        ws_emp.data_validation(first_row=start_row, first_col=pref_idx, last_row=end_row, last_col=pref_idx,
                               options={"validate":"integer","criteria":"<=","value":10})
        # availability_simple dropdowns
        start_row_av = 1; end_row_av = start_row_av + 600
        for d in DAYS:
            col_idx = availability_simple_df.columns.get_loc(d)
            ws_avs.data_validation(first_row=start_row_av, first_col=col_idx, last_row=end_row_av, last_col=col_idx,
                                   options={"validate":"list","source":yesno})
    bio.seek(0)
    return bio.read()

# ------------------ UI ------------------
with st.sidebar:
    st.header("Template")
    st.download_button(
        label="Download blank template (v4)",
        data=generate_template_bytes(),
        file_name="schedule_input_template_v4.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.caption("Template includes skills matrix (Yes/No), availability_simple (Yes = NOT available), rules Min=3 / Max=10, and anchors (Open 08:00, Mid 10:00, Close 20:30). New: MaxHoursPerWeek column.")

uploaded = st.file_uploader("Upload your filled template (xlsx)", type=["xlsx"])
startdate = st.text_input("Start date (YYYY-MM-DD, optional; defaults to upcoming Sunday)", "")
if st.button("Generate 4-week schedule", type="primary"):
    if not uploaded:
        st.error("Please upload the input Excel first.")
    else:
        try:
            schedule_file, detailed = run_schedule(uploaded, startdate.strip() or None)
            st.success("Schedule generated! Download below.")
            st.download_button("Download schedule_output_calendar_weeks.xlsx", data=schedule_file, file_name="schedule_output_calendar_weeks.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.exception(e)
