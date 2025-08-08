
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="Goodwill Scheduler", layout="wide")
st.title("Goodwill Gulf Coast — Weekly Scheduler")
st.caption("Upload the input Excel, choose an optional start date, and download a 4-week schedule.")

DAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
KNOWN_ROLES = ["Cashier","Donation","Pricer","Hanger","Manager"]

# Your baked-in coverage counts (from schedule_input_template_v3_rules_and_coverage.xlsx)
COVERAGE_DEFAULT = [{"Role": "Cashier", "Day": "Friday", "ShiftType": "Close", "Count": 1}, {"Role": "Cashier", "Day": "Friday", "ShiftType": "Mid", "Count": 1}, {"Role": "Cashier", "Day": "Friday", "ShiftType": "Open", "Count": 1}, {"Role": "Cashier", "Day": "Monday", "ShiftType": "Close", "Count": 1}, {"Role": "Cashier", "Day": "Monday", "ShiftType": "Mid", "Count": 1}, {"Role": "Cashier", "Day": "Monday", "ShiftType": "Open", "Count": 1}, {"Role": "Cashier", "Day": "Saturday", "ShiftType": "Close", "Count": 1}, {"Role": "Cashier", "Day": "Saturday", "ShiftType": "Mid", "Count": 2}, {"Role": "Cashier", "Day": "Saturday", "ShiftType": "Open", "Count": 1}, {"Role": "Cashier", "Day": "Sunday", "ShiftType": "Close", "Count": 1}, {"Role": "Cashier", "Day": "Sunday", "ShiftType": "Open", "Count": 1}, {"Role": "Cashier", "Day": "Thursday", "ShiftType": "Close", "Count": 1}, {"Role": "Cashier", "Day": "Thursday", "ShiftType": "Mid", "Count": 1}, {"Role": "Cashier", "Day": "Thursday", "ShiftType": "Open", "Count": 1}, {"Role": "Cashier", "Day": "Tuesday", "ShiftType": "Close", "Count": 1}, {"Role": "Cashier", "Day": "Tuesday", "ShiftType": "Mid", "Count": 1}, {"Role": "Cashier", "Day": "Tuesday", "ShiftType": "Open", "Count": 1}, {"Role": "Cashier", "Day": "Wednesday", "ShiftType": "Close", "Count": 1}, {"Role": "Cashier", "Day": "Wednesday", "ShiftType": "Mid", "Count": 1}, {"Role": "Cashier", "Day": "Wednesday", "ShiftType": "Open", "Count": 1}, {"Role": "Donation", "Day": "Friday", "ShiftType": "Close", "Count": 1}, {"Role": "Donation", "Day": "Friday", "ShiftType": "Mid", "Count": 1}, {"Role": "Donation", "Day": "Friday", "ShiftType": "Open", "Count": 1}, {"Role": "Donation", "Day": "Monday", "ShiftType": "Close", "Count": 1}, {"Role": "Donation", "Day": "Monday", "ShiftType": "Mid", "Count": 1}, {"Role": "Donation", "Day": "Monday", "ShiftType": "Open", "Count": 1}, {"Role": "Donation", "Day": "Saturday", "ShiftType": "Close", "Count": 1}, {"Role": "Donation", "Day": "Saturday", "ShiftType": "Mid", "Count": 2}, {"Role": "Donation", "Day": "Saturday", "ShiftType": "Open", "Count": 1}, {"Role": "Donation", "Day": "Sunday", "ShiftType": "Close", "Count": 1}, {"Role": "Donation", "Day": "Sunday", "ShiftType": "Mid", "Count": 1}, {"Role": "Donation", "Day": "Sunday", "ShiftType": "Open", "Count": 1}, {"Role": "Donation", "Day": "Thursday", "ShiftType": "Close", "Count": 1}, {"Role": "Donation", "Day": "Thursday", "ShiftType": "Mid", "Count": 1}, {"Role": "Donation", "Day": "Thursday", "ShiftType": "Open", "Count": 1}, {"Role": "Donation", "Day": "Tuesday", "ShiftType": "Close", "Count": 1}, {"Role": "Donation", "Day": "Tuesday", "ShiftType": "Mid", "Count": 1}, {"Role": "Donation", "Day": "Tuesday", "ShiftType": "Open", "Count": 1}, {"Role": "Donation", "Day": "Wednesday", "ShiftType": "Close", "Count": 1}, {"Role": "Donation", "Day": "Wednesday", "ShiftType": "Mid", "Count": 1}, {"Role": "Donation", "Day": "Wednesday", "ShiftType": "Open", "Count": 1}, {"Role": "Hanger", "Day": "Friday", "ShiftType": "Close", "Count": 1}, {"Role": "Hanger", "Day": "Friday", "ShiftType": "Mid", "Count": 1}, {"Role": "Hanger", "Day": "Friday", "ShiftType": "Open", "Count": 2}, {"Role": "Hanger", "Day": "Monday", "ShiftType": "Close", "Count": 1}, {"Role": "Hanger", "Day": "Monday", "ShiftType": "Mid", "Count": 1}, {"Role": "Hanger", "Day": "Monday", "ShiftType": "Open", "Count": 2}, {"Role": "Hanger", "Day": "Saturday", "ShiftType": "Close", "Count": 1}, {"Role": "Hanger", "Day": "Saturday", "ShiftType": "Mid", "Count": 1}, {"Role": "Hanger", "Day": "Saturday", "ShiftType": "Open", "Count": 2}, {"Role": "Hanger", "Day": "Sunday", "ShiftType": "Close", "Count": 1}, {"Role": "Hanger", "Day": "Sunday", "ShiftType": "Mid", "Count": 1}, {"Role": "Hanger", "Day": "Sunday", "ShiftType": "Open", "Count": 2}, {"Role": "Hanger", "Day": "Thursday", "ShiftType": "Close", "Count": 1}, {"Role": "Hanger", "Day": "Thursday", "ShiftType": "Mid", "Count": 1}, {"Role": "Hanger", "Day": "Thursday", "ShiftType": "Open", "Count": 2}, {"Role": "Hanger", "Day": "Tuesday", "ShiftType": "Close", "Count": 1}, {"Role": "Hanger", "Day": "Tuesday", "ShiftType": "Mid", "Count": 1}, {"Role": "Hanger", "Day": "Tuesday", "ShiftType": "Open", "Count": 2}, {"Role": "Hanger", "Day": "Wednesday", "ShiftType": "Close", "Count": 1}, {"Role": "Hanger", "Day": "Wednesday", "ShiftType": "Mid", "Count": 1}, {"Role": "Hanger", "Day": "Wednesday", "ShiftType": "Open", "Count": 2}, {"Role": "Manager", "Day": "Friday", "ShiftType": "Close", "Count": 1}, {"Role": "Manager", "Day": "Friday", "ShiftType": "Open", "Count": 1}, {"Role": "Manager", "Day": "Monday", "ShiftType": "Close", "Count": 1}, {"Role": "Manager", "Day": "Monday", "ShiftType": "Open", "Count": 1}, {"Role": "Manager", "Day": "Saturday", "ShiftType": "Close", "Count": 1}, {"Role": "Manager", "Day": "Saturday", "ShiftType": "Open", "Count": 1}, {"Role": "Manager", "Day": "Sunday", "ShiftType": "Close", "Count": 1}, {"Role": "Manager", "Day": "Sunday", "ShiftType": "Open", "Count": 1}, {"Role": "Manager", "Day": "Thursday", "ShiftType": "Close", "Count": 1}, {"Role": "Manager", "Day": "Thursday", "ShiftType": "Open", "Count": 1}, {"Role": "Manager", "Day": "Tuesday", "ShiftType": "Close", "Count": 1}, {"Role": "Manager", "Day": "Tuesday", "ShiftType": "Open", "Count": 1}, {"Role": "Manager", "Day": "Wednesday", "ShiftType": "Close", "Count": 1}, {"Role": "Manager", "Day": "Wednesday", "ShiftType": "Open", "Count": 1}, {"Role": "Pricer", "Day": "Friday", "ShiftType": "Close", "Count": 2}, {"Role": "Pricer", "Day": "Friday", "ShiftType": "Open", "Count": 2}, {"Role": "Pricer", "Day": "Monday", "ShiftType": "Close", "Count": 2}, {"Role": "Pricer", "Day": "Monday", "ShiftType": "Open", "Count": 2}, {"Role": "Pricer", "Day": "Saturday", "ShiftType": "Close", "Count": 2}, {"Role": "Pricer", "Day": "Saturday", "ShiftType": "Open", "Count": 2}, {"Role": "Pricer", "Day": "Sunday", "ShiftType": "Close", "Count": 1}, {"Role": "Pricer", "Day": "Sunday", "ShiftType": "Open", "Count": 2}, {"Role": "Pricer", "Day": "Thursday", "ShiftType": "Close", "Count": 2}, {"Role": "Pricer", "Day": "Thursday", "ShiftType": "Open", "Count": 2}, {"Role": "Pricer", "Day": "Tuesday", "ShiftType": "Close", "Count": 2}, {"Role": "Pricer", "Day": "Tuesday", "ShiftType": "Open", "Count": 2}, {"Role": "Pricer", "Day": "Wednesday", "ShiftType": "Close", "Count": 2}, {"Role": "Pricer", "Day": "Wednesday", "ShiftType": "Open", "Count": 2}]

def load_inputs(file):
    xls = pd.ExcelFile(file)
    employees = pd.read_excel(xls, "employees").fillna("")
    availability_simple = pd.read_excel(xls, "availability_simple").fillna("") if "availability_simple" in xls.sheet_names else None
    availability = pd.read_excel(xls, "availability").fillna("") if "availability" in xls.sheet_names else None
    coverage = pd.read_excel(xls, "coverage_template").fillna("")
    anchors = pd.read_excel(xls, "shift_anchors").fillna("")
    rules = pd.read_excel(xls, "rules")
    for df in [coverage]:
        df["Day"] = df["Day"].astype(str).str.strip().str.title()
    return employees, availability, availability_simple, coverage, anchors, rules

def build_availability_map(employees, availability, availability_simple):
    avail = {}
    if availability_simple is not None and not availability_simple.empty:
        df = availability_simple.copy()
        df.columns = [str(c).strip().title() if c != "Employee" else "Employee" for c in df.columns]
        for _, row in df.iterrows():
            emp = row["Employee"]
            for d in DAYS:
                resp = str(row.get(d, "")).strip().lower()
                if resp in ("yes","y","true","1"):
                    avail.setdefault(emp, {}).setdefault(d, [])
                else:
                    avail.setdefault(emp, {}).setdefault(d, []).append(("08:00","20:30"))
        return avail
    if availability is not None and not availability.empty:
        df = availability.copy()
        df["Day"] = df["Day"].astype(str).str.strip().str.title()
        for _, row in df.iterrows():
            emp = row["Employee"]; d = row["Day"]
            start = str(row["Start"]).strip(); end = str(row["End"]).strip()
            if not start or not end: continue
            avail.setdefault(emp, {}).setdefault(d, []).append((start, end))
        return avail
    for emp in employees["Employee"]:
        for d in DAYS:
            avail.setdefault(emp, {}).setdefault(d, []).append(("08:00","20:30"))
    return avail

def employee_functions_map(employees):
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
                    t = token.strip()
                    if not t: continue
                    t = t[:1].upper() + t[1:].lower()
                    if t in KNOWN_ROLES: skills.add(t)
            m[emp] = skills
        return m
    for _, r in employees.iterrows():
        emp = r["Employee"]
        skills = set()
        raw = str(r.get("Functions",""))
        for token in raw.replace(";",",").replace("/",",").split(","):
            t = token.strip()
            if not t: continue
            t = t[:1].upper() + t[1:].lower()
            if t in KNOWN_ROLES: skills.add(t)
        m[emp] = skills
    return m

def weekend_rotation(employees):
    emps = list(employees["Employee"])
    groups = {i:set() for i in range(4)}
    for i, e in enumerate(emps):
        groups[i % 4].add(e)
    return groups

def generate_week_dates(start_date):
    start = start_date - timedelta(days=start_date.weekday()+1 if start_date.weekday()!=6 else 0)
    return [[start + timedelta(days=7*w + i) for i in range(7)] for w in range(4)]

def anchored_shift_window(shift_type, hours, anchors_map):
    a = anchors_map[shift_type]
    if a["Anchor"] == "Start":
        st = datetime.strptime(a["Time"], "%H:%M")
        en = st + timedelta(hours=float(hours))
        return (a["Time"], en.strftime("%H:%M"))
    else:
        en = datetime.strptime(a["Time"], "%H:%M")
        st = en - timedelta(hours=float(hours))
        return (st.strftime("%H:%M"), a["Time"])

def can_assign(emp, role, day_name, start_str, end_str, avail_map, funcs_map, assigned, rules, assigned_hours, days_worked, min_hours_map):
    if role not in funcs_map.get(emp, set()): return False
    windows = avail_map.get(emp, {}).get(day_name, [])
    if not any(a_start <= start_str and a_end >= end_str for a_start, a_end in windows): return False
    min_gap = float(rules.get("MinGapBetweenShiftsHours", 1))
    def to_dt_local(t): return datetime.combine(datetime.today(), datetime.strptime(t, "%H:%M").time())
    slot_s, slot_e = to_dt_local(start_str), to_dt_local(end_str)
    for a in assigned:
        if a["Employee"]!=emp or a["Day"]!=day_name: continue
        s, e = to_dt_local(a["Start"]), to_dt_local(a["End"])
        if not (slot_e <= s or slot_s >= e): return False
        if 0 <= (s - slot_e).total_seconds()/3600 < min_gap or 0 <= (slot_s - e).total_seconds()/3600 < min_gap: return False
    if len(days_worked.get(emp, set())) >= int(rules.get("MaxDaysPerWeek", 6)) and day_name not in days_worked.get(emp,set()): return False
    return True

def run_schedule(file, startdate_str=None, role_colors=None):
    employees, availability, availability_simple, coverage, anchors, rules_df = load_inputs(file)
    rules = dict(zip(rules_df["Rule"], rules_df["Value"]))
    # default to 3..10 if missing
    rules["MinShiftHours"] = float(rules.get("MinShiftHours", 3))
    rules["MaxShiftHours"] = float(rules.get("MaxShiftHours", 10))
    rules["MinGapBetweenShiftsHours"] = float(rules.get("MinGapBetweenShiftsHours", 1))
    rules["MaxDaysPerWeek"] = int(rules.get("MaxDaysPerWeek", 6))
    anchors_map = {r["ShiftType"]: {"Anchor": r["Anchor"], "Time": r["Time"]} for _, r in anchors.iterrows()}
    pref_hours = dict(zip(employees["Employee"], employees.get("PreferredShiftHours", pd.Series([rules["MaxShiftHours"]]*len(employees))).fillna(rules["MaxShiftHours"])))
    funcs_map = employee_functions_map(employees)
    avail_map = build_availability_map(employees, availability, availability_simple)
    min_hours_map = dict(zip(employees["Employee"], employees["MinHours"]))
    if startdate_str:
        startdate = datetime.strptime(startdate_str, "%Y-%m-%d").date()
    else:
        today = datetime.today().date()
        delta = (6 - today.weekday()) % 7
        startdate = today + timedelta(days=delta)
    weeks = generate_week_dates(startdate)
    weekend_groups = weekend_rotation(employees)
    assignments = []
    for w_idx, week_dates in enumerate(weeks):
        off_group = weekend_groups[w_idx]
        local_avail = {e:{d:list(w) for d,w in days.items()} for e,days in avail_map.items()}
        for e in off_group:
            for d in ["Saturday","Sunday"]:
                if e in local_avail and d in local_avail[e]:
                    local_avail[e][d] = []
        assigned_hours = {e:0 for e in employees["Employee"]}
        days_worked = {e:set() for e in employees["Employee"]}
        for d_idx, date in enumerate(week_dates):
            day_name = DAYS[d_idx]
            day_cov = coverage[coverage["Day"]==day_name].copy()
            rarity = {}
            for role in day_cov["Role"].unique():
                count = sum(1 for e in employees["Employee"] if role in funcs_map.get(e,set()))
                rarity[role] = 1.0/(count if count>0 else 0.5)
            day_cov["rarity"] = day_cov["Role"].map(lambda r: rarity.get(r,1.0))
            day_cov = day_cov.sort_values("rarity", ascending=False)
            demands = []
            for _, r in day_cov.iterrows():
                demands += [{"Role": r["Role"], "ShiftType": r["ShiftType"]} for _ in range(int(r["Count"]))]
            for dem in demands:
                role, stype = dem["Role"], dem["ShiftType"]
                candidates = []
                for e in employees["Employee"]:
                    if role not in funcs_map.get(e,set()): continue
                    hours = float(pref_hours.get(e, rules["MaxShiftHours"]))
                    hours = min(max(hours, rules["MinShiftHours"]), rules["MaxShiftHours"])
                    s_str, e_str = anchored_shift_window(stype, hours, anchors_map)
                    if not can_assign(e, role, day_name, s_str, e_str, local_avail, funcs_map, assignments, rules, assigned_hours, days_worked, min_hours_map):
                        continue
                    under_min = 1 if assigned_hours[e] < min_hours_map.get(e,0) else 0
                    candidates.append((under_min, -assigned_hours[e], -len(days_worked[e]), e, s_str, e_str, role))
                if not candidates:
                    assignments.append({"Week": w_idx+1, "Date": date.strftime("%Y-%m-%d"), "Day": day_name, "Employee": "UNFILLED", "Start": "", "End": "", "Role": role})
                    continue
                candidates.sort(reverse=True)
                _,_,_, e, s_str, e_str, role = candidates[0]
                assignments.append({"Week": w_idx+1, "Date": date.strftime("%Y-%m-%d"), "Day": day_name, "Employee": e, "Start": s_str, "End": e_str, "Role": role})
                sh = (datetime.strptime(e_str, "%H:%M") - datetime.strptime(s_str, "%H:%M")).seconds/3600
                assigned_hours[e] += sh
                days_worked[e].add(day_name)
    out = pd.DataFrame(assignments)

    # Build separate sheets with color-by-role and print setup
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "align":"center", "valign":"vcenter", "border":1})
        name_fmt = workbook.add_format({"bold": True, "align":"left", "valign":"vcenter", "border":1})
        cell_fmt = workbook.add_format({"text_wrap": True, "align":"left", "valign":"top", "border":1})
        off_fmt = workbook.add_format({"align":"center", "valign":"vcenter", "border":1, "font_color":"#666666"})
        default_role_colors = {"Cashier":"#E3F2FD","Donation":"#E8F5E9","Pricer":"#FFF3E0","Hanger":"#F3E5F5","Manager":"#FFEBEE"}
        role_colors = role_colors or default_role_colors
        role_formats = {role: workbook.add_format({"text_wrap": True, "align":"left", "valign":"top", "border":1, "bg_color": color})
                        for role, color in role_colors.items()}
        out.to_excel(writer, index=False, sheet_name="assignments_detailed")
        for wk in range(1,5):
            wkdf = out[out["Week"]==wk].copy()
            if wkdf.empty: continue
            start_date = wkdf["Date"].min(); end_date = wkdf["Date"].max()
            title = f"Week {wk} ({start_date} to {end_date})"
            emps = sorted([e for e in wkdf["Employee"].dropna().unique() if e != "UNFILLED"])
            frame = pd.DataFrame({"Employee": emps}, dtype=object)
            for d in DAYS: frame[d] = ""
            role_map = {}  # (row_idx, day_name) -> role for formatting
            for e in emps:
                for d in DAYS:
                    rows = wkdf[(wkdf["Employee"]==e) & (wkdf["Day"]==d)]
                    if len(rows)==0:
                        row_idx = frame.index[frame["Employee"]==e][0]
                    frame.at[row_idx, d] = "OFF"
                    role_map[(row_idx, d)] = None
                    continue
                    parts = []; first_role = None
                    for _, r in rows.iterrows():
                        st = datetime.strptime(r["Start"], "%H:%M").strftime("%-I:%M %p")
                        en = datetime.strptime(r["End"], "%H:%M").strftime("%-I:%M %p")
                        parts.append(f"{st}-{en} {r['Role']}")
                        if first_role is None: first_role = r["Role"]
                    row_idx = frame.index[frame["Employee"]==e][0]
                    frame.at[row_idx, d] = "\n".join(parts)
                    role_map[(row_idx, d)] = first_role
            sheetname = f"Week {wk}"
            frame_out = frame.copy()
            for d in DAYS:
                frame_out[d] = frame_out[d].apply(lambda x: x[0] if isinstance(x, tuple) else x)
            frame_out.to_excel(writer, index=False, sheet_name=sheetname, startrow=1)
            ws = writer.sheets[sheetname]
            ws.merge_range(0,0,0,8, title, header_fmt)
            ws.set_column(0,0,24); ws.set_column(1,7,22)
            for col_idx, d in enumerate(DAYS, start=1):
                drows = wkdf[wkdf["Day"]==d]
                if not drows.empty:
                    dt = pd.to_datetime(drows["Date"].iloc[0]).date().strftime("%m/%d")
                    label = f"{d} {dt}"
                else:
                    label = d
                ws.write(1, col_idx, label, header_fmt)
            ws.write(1, 0, "Employee", header_fmt)
            ws.set_landscape(); ws.fit_to_pages(1, 0); ws.repeat_rows(1, 1); ws.center_horizontally()
            for r in range(2, 2 + len(frame)):
                emp_val = frame.iloc[r-2]["Employee"]
                try:
                    import pandas as _pd
                    if _pd.isna(emp_val): emp_val = ""
                except Exception:
                    pass
                ws.write_string(r, 0, str(emp_val), name_fmt)
                for c, d in enumerate(DAYS, start=1):
                    val = frame.iloc[r-2][d]
                    role = role_map.get((r-2, d), None)
                    import pandas as _pd
                    if (isinstance(val, float) and (val != val)) or (hasattr(_pd, "isna") and _pd.isna(val)):
                        val = ""
                    if val == "" or val == "OFF":
                        ws.write_string(r, c, "OFF", off_fmt)
                    else:
                        fmt = role_formats.get(role, cell_fmt)
                        ws.write_string(r, c, str(val).replace("\n","\n"), fmt)
    output.seek(0)
    return output

def generate_template_bytes():
    # Build employees + availability_simple skeleton (no store data), with baked-in coverage and rules
    employees_df = pd.DataFrame({
        "Employee": ["Example Person 1", "Example Person 2"],
        "MaxHours": [35, 30],
        "MinHours": [15, 15],
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
    coverage_df = pd.DataFrame(COVERAGE_DEFAULT)
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
        # Add data validation dropdowns
        wb = writer.book
        ws_emp = writer.sheets["employees"]
        ws_avs = writer.sheets["availability_simple"]
        yesno = ["Yes","No"]
        start_row = 1; end_row = start_row + 200
        for role in ["Cashier","Donation","Pricer","Hanger","Manager"]:
            idx = employees_df.columns.get_loc(role)
            ws_emp.data_validation(first_row=start_row, first_col=idx, last_row=end_row, last_col=idx,
                                   options={"validate":"list","source":yesno})
        pref_idx = employees_df.columns.get_loc("PreferredShiftHours")
        ws_emp.data_validation(first_row=start_row, first_col=pref_idx, last_row=end_row, last_col=pref_idx,
                               options={"validate":"integer","criteria":">=","value":3})
        ws_emp.data_validation(first_row=start_row, first_col=pref_idx, last_row=end_row, last_col=pref_idx,
                               options={"validate":"integer","criteria":"<=","value":10})
        start_row_av = 1; end_row_av = start_row_av + 400
        for d in DAYS:
            col_idx = availability_simple_df.columns.get_loc(d)
            ws_avs.data_validation(first_row=start_row_av, first_col=col_idx, last_row=end_row_av, last_col=col_idx,
                                   options={"validate":"list","source":yesno})
    bio.seek(0)
    return bio.read()

# Sidebar: v3 template download with baked-in coverage
st.sidebar.header("Template")
st.sidebar.download_button(
    label="Download blank template",
    data=generate_template_bytes(),
    file_name="schedule_input_template_v3.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Upload + run
uploaded = st.file_uploader("Upload your filled template (xlsx)", type=["xlsx"])
startdate = st.text_input("Start date (YYYY-MM-DD, optional; defaults to upcoming Sunday)", "")
if st.button("Generate 4-week schedule", type="primary"):
    if not uploaded:
        st.error("Please upload the input Excel first.")
    else:
        try:
            schedule_file = run_schedule(uploaded, startdate.strip() or None)
            st.success("Schedule generated!")
            st.download_button("Download schedule_output_calendar_weeks.xlsx", data=schedule_file, file_name="schedule_output_calendar_weeks.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.exception(e)
