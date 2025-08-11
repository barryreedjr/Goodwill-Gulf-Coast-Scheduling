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
    try:
        t = datetime.strptime(hhmm, "%H:%M")
        return t.strftime("%H%M")
    except Exception:
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
    availability_simple = (
        pd.read_excel(xls, "availability_simple").fillna("")
        if "availability_simple" in xls.sheet_names
        else None
    )
    availability = (
        pd.read_excel(xls, "availability").fillna("")
        if "availability" in xls.sheet_names
        else None
    )
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
    avail = {}
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
    for emp in employees["Employee"]:
        for d in DAYS:
            avail.setdefault(emp, {}).setdefault(d, []).append(("08:00", "20:30"))
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
    if paid_hours > trigger_hours:
        return lunch_minutes / 60.0
    return 0.0

def anchored_shift_window(shift_type, paid_hours, anchors_map, rules) -> tuple[str, str, float]:
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
    emp, role, day_name, start_str, end_str,
    avail_map, funcs_map, assigned_list_for_week, rules,
    assigned_paid_hours_week, days_worked_week, min_hours_map,
    weekly_max_paid_hours_map
):
    if role not in funcs_map.get(emp, set()):
        return False
    windows = avail_map.get(emp, {}).get(day_name, [])
    if not any(a_start <= start_str and a_end >= end_str for a_start, a_end in windows):
        return False
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
    paid_hours = float(rules.get("_candidate_paid_hours", 0))
    if assigned_paid_hours_week.get(emp, 0) + paid_hours > weekly_max_paid_hours_map.get(emp, float("inf")):
        return False
    if len(days_worked_week.get(emp, set())) >= int(rules.get("MaxDaysPerWeek", 6)) and day_name not in days_worked_week.get(emp, set()):
        return False
    return True

# ------------------ Scheduler ------------------
def run_schedule_with_summary(file, startdate_str=None, role_colors=None):
    # Load inputs
    employees, availability, availability_simple, coverage, anchors, rules_df = load_inputs(file)
    rules = dict(zip(rules_df["Rule"], rules_df["Value"]))

    rules["MinShiftHours"] = float(rules.get("MinShiftHours", 3))
    rules["MaxShiftHours"] = float(rules.get("MaxShiftHours", 10))
    rules["MinGapBetweenShiftsHours"] = float(rules.get("MinGapBetweenShiftsHours", 1))
    rules["MaxDaysPerWeek"] = int(rules.get("MaxDaysPerWeek", 6))
    rules["LunchTriggerHours"] = float(rules.get("LunchTriggerHours", 5))
    rules["LunchMinutes"] = int(rules.get("LunchMinutes", 30))

    night_types = set([s.strip() for s in str(rules.get("NightShiftTypes", "Close")).split(",") if s.strip()])
    night_end_threshold = str(rules.get("NightShiftEndAtOrAfter", "")).strip() or None

    anchors_map = {r["ShiftType"]: {"Anchor": r["Anchor"], "Time": r["Time"]} for _, r in anchors.iterrows()}

    weekly_max_series = get_numeric_series(
        employees,
        "MaxHoursPerWeek" if "MaxHoursPerWeek" in employees.columns else "MaxHours",
        999,
        len(employees),
    )
    weekly_max_paid_hours_map = dict(zip(employees["Employee"], weekly_max_series))

    emp_id_col = None
    for cand in ["EmployeeID", "Employee ID"]:
        if cand in employees.columns:
            emp_id_col = cand
            break
    if emp_id_col:
        emp_id_map = dict(zip(employees["Employee"], employees[emp_id_col].astype(str).str.strip()))
    else:
        emp_id_map = {e: "" for e in employees["Employee"]}

    pref_hours = dict(zip(employees["Employee"], get_numeric_series(employees, "PreferredShiftHours", rules["MaxShiftHours"], len(employees))))
    funcs_map = employee_functions_map(employees)
    avail_map = build_availability_map(employees, availability, availability_simple)
    min_hours_map = dict(zip(employees["Employee"], get_numeric_series(employees, "MinHours", 0, len(employees))))

    if startdate_str:
        startdate = datetime.strptime(startdate_str, "%Y-%m-%d").date()
    else:
        today = datetime.today().date()
        delta = (6 - today.weekday()) % 7
        startdate = today + timedelta(days=delta)

    weeks = generate_week_dates(startdate)
    weekend_groups = weekend_rotation(employees)

    # Priority mapping
    SHIFT_PRIORITY = {"Open": 1, "Close": 2, "Mid": 3}
    def get_shift_priority(stype):
        return SHIFT_PRIORITY.get(str(stype).strip().title(), 99)

    assignments_all = []
    for w_idx, week_dates in enumerate(weeks):
        assigned_paid_hours_week = {e: 0.0 for e in employees["Employee"]}
        days_worked_week = {e: set() for e in employees["Employee"]}
        nights_worked_week = {e: 0 for e in employees["Employee"]}
        assignments_this_week = []
        off_group = weekend_groups[w_idx]
        local_avail = {e: {d: list(w) for d, w in days.items()} for e, days in avail_map.items()}
        for e in off_group:
            for d in ["Saturday", "Sunday"]:
                if e in local_avail and d in local_avail[e]:
                    local_avail[e][d] = []

        for d_idx, dt in enumerate(week_dates):
            day_name = DAYS[d_idx]
            day_cov = coverage[coverage["Day"] == day_name].copy()

            demands = []
            for _, r in day_cov.iterrows():
                count = int(pd.to_numeric(r.get("Count", 0), errors="coerce") or 0)
                role = r.get("Role", "")
                stype = r.get("ShiftType", "")
                demands += [{"Role": role, "ShiftType": stype} for _ in range(count)]

            rarity = {}
            roles_unique = [x for x in day_cov["Role"].unique() if isinstance(x, str)]
            for role in roles_unique:
                count = sum(1 for e in employees["Employee"] if role in funcs_map.get(e, set()))
                rarity[role] = 1.0 / (count if count > 0 else 0.5)

            demands.sort(key=lambda d: (get_shift_priority(d["ShiftType"]), -rarity.get(d["Role"], 1.0)))

            for dem in demands:
                role, stype = dem["Role"], dem["ShiftType"]
                candidates = []
                for e in employees["Employee"]:
                    if role not in funcs_map.get(e, set()):
                        continue
                    paid_hours = float(pref_hours.get(e, rules["MaxShiftHours"]))
                    paid_hours = min(max(paid_hours, rules["MinShiftHours"]), rules["MaxShiftHours"])
                    s_str, e_str, lunch_pad = anchored_shift_window(stype, paid_hours, anchors_map, rules)
                    rules["_candidate_paid_hours"] = paid_hours
                    if not can_assign(e, role, day_name, s_str, e_str, local_avail, funcs_map, assignments_this_week, rules, assigned_paid_hours_week, days_worked_week, min_hours_map, weekly_max_paid_hours_map):
                        continue
                    night_flag = is_night_shift(stype, e_str, night_types, night_end_threshold)
                    needs_night = 1 if (night_flag and nights_worked_week[e] == 0) else 0
                    under_min = 1 if assigned_paid_hours_week[e] < min_hours_map.get(e, 0) else 0
                    remaining_cap = weekly_max_paid_hours_map.get(e, float("inf")) - assigned_paid_hours_week.get(e, 0)
                    candidates.append((needs_night, under_min, remaining_cap, -assigned_paid_hours_week[e], -len(days_worked_week[e]), e, s_str, e_str, paid_hours, lunch_pad, role, night_flag, stype))
                if not candidates:
                    assignments_this_week.append({"Week": w_idx + 1, "Date": dt.strftime("%Y-%m-%d"), "Day": day_name, "Employee": "UNFILLED", "EmployeeID": "", "Start": "", "End": "", "Role": role, "ShiftType": stype, "PaidHours": 0.0, "ScheduledHours": 0.0, "Night": False})
                    continue
                candidates.sort(reverse=True)
                _, _, _, _, _, e, s_str, e_str, paid_hours, lunch_pad, role, night_flag, stype = candidates[0]
                assignments_this_week.append({"Week": w_idx + 1, "Date": dt.strftime("%Y-%m-%d"), "Day": day_name, "Employee": e, "EmployeeID": emp_id_map.get(e, ""), "Start": s_str, "End": e_str, "Role": role, "ShiftType": stype, "PaidHours": paid_hours, "ScheduledHours": paid_hours + lunch_pad, "Night": bool(night_flag)})
                assigned_paid_hours_week[e] += paid_hours
                days_worked_week[e].add(day_name)
                if night_flag:
                    nights_worked_week[e] += 1
        assignments_all.extend(assignments_this_week)

    out = pd.DataFrame(assignments_all)

    # --- rest of your preview, export, Paycom import code remains unchanged from your original ---
    # (paste your existing "preview", "csv download", "hours summary", and "Excel export" code here)
