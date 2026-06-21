import streamlit as st
import pandas as pd
import calendar
import json
import os
import holidays
from collections import Counter
import datetime
from ortools.sat.python import cp_model
import xlsxwriter
import io
from github import Github

# ==========================================
# 0. GITHUB CLOUD SYNC ENGINE
# ==========================================
def push_to_github(file_path, commit_message):
    """Silently pushes local json updates to GitHub if running on Streamlit Cloud."""
    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
        try:
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(st.secrets["GITHUB_REPO"])
            
            with open(file_path, 'r') as file:
                content = file.read()
                
            try:
                contents = repo.get_contents(file_path)
                repo.update_file(contents.path, commit_message, content, contents.sha)
            except:
                repo.create_file(file_path, commit_message, content)
        except Exception as e:
            st.toast(f"⚠️ Could not sync {file_path} to cloud: {e}")

# ==========================================
# 1. PERSISTENT DATABASES
# ==========================================
SETTINGS_FILE = "persistent_settings.json"
COUNTER_FILE = "historical_counters.json"

COLOR_PALETTE = {
    "White": "#FFFFFF", "Light Blue": "#CCEBFF", "Light Green": "#CCFFCC", 
    "Light Red": "#FFCCCC", "Light Yellow": "#FFFFCC", "Peach": "#FFE5CC", 
    "Lavender": "#E5CCFF", "Light Pink": "#FFCCFF", "Mint": "#CCFFEA", 
    "Light Grey": "#E0E0E0", "Light Orange": "#FFD699"
}

def load_persistent_settings():
    default_doctors = ["BURGAZZI", "CACCAMO", "CARDINALI", "CICCARELLI", "CICCHIRILLO", "FLORI", "FINIZIO", "NUCCI", "ROBERTI", "GAUDENZI"]
    default_clinics = ['PACEMAKER', 'DIMESSI', 'SCOMPENSO']
    
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                data = json.load(f)
            if "docs_base" in data and "capabilities" in data:
                return data
            elif "doctors" in data:
                docs_base = []
                caps = []
                for doc in data["doctors"]:
                    docs_base.append({
                        "Doctor": doc.get("Doctor", ""),
                        "Private Practice (Afternoon)": doc.get("Private Practice (Afternoon)", ""),
                        "Color": doc.get("Color", "White")
                    })
                    cap_row = {"Doctor": doc.get("Doctor", ""), "Ward Preferred": doc.get("Ward Preferred", False), "OR Capable": doc.get("OR Capable", False)}
                    for c in data.get("clinics", default_clinics): cap_row[c] = doc.get(c, False)
                    caps.append(cap_row)
                return {"clinics": data.get("clinics", default_clinics), "docs_base": docs_base, "capabilities": caps}
        except Exception:
            pass
            
    docs_base = [{"Doctor": d, "Private Practice (Afternoon)": "", "Color": "White"} for d in default_doctors]
    caps = [{"Doctor": d, "Ward Preferred": d in ["BURGAZZI", "ROBERTI"], "OR Capable": d in ["CACCAMO", "NUCCI"]} for d in default_doctors]
    for c in caps:
        for clin in default_clinics: c[clin] = False
    return {"clinics": default_clinics, "docs_base": docs_base, "capabilities": caps}

def save_persistent_settings(clinics_list, docs_df, cap_df):
    data = {"clinics": clinics_list, "docs_base": docs_df.fillna("").to_dict('records'), "capabilities": cap_df.fillna(False).to_dict('records')}
    with open(SETTINGS_FILE, 'w') as f: json.dump(data, f, indent=4)

def load_historical_counters():
    if not os.path.exists(COUNTER_FILE):
        return {"legacy_baseline": {}}
    try:
        with open(COUNTER_FILE, 'r') as f:
            data = json.load(f)
        if "legacy_baseline" in data:
            return data
        else:
            new_data = {"legacy_baseline": {}}
            for doc, stats in data.items():
                new_data["legacy_baseline"][doc] = stats
            return new_data
    except:
        return {"legacy_baseline": {}}

def get_lifetime_stats(ledger, doc):
    stats = {'nights': 0.0, 'doubles': 0.0, 'saturdays': 0.0, 'sundays': 0.0, 'holidays': 0.0, 'super_holidays': 0.0, 'golden_weekends': 0.0, 'reps': 0.0}
    for month_key, month_data in ledger.items():
        if doc in month_data:
            for k in stats.keys():
                val = month_data[doc].get(k, 0.0)
                stats[k] += float(val) if not isinstance(val, dict) else 0.0
    return stats

# ==========================================
# 2. CALENDAR & UTILS
# ==========================================
def get_roster_dates(year, month):
    cal = calendar.Calendar(firstweekday=0) 
    mondays = []
    for week in cal.monthdatescalendar(year, month):
        if week[0].month == month:
            if week[0] not in mondays: mondays.append(week[0])
    start_date = mondays[0]
    end_date = mondays[-1] + datetime.timedelta(days=6)
    roster_dates = []
    curr = start_date
    while curr <= end_date:
        roster_dates.append(curr)
        curr += datetime.timedelta(days=1)
    return roster_dates

def get_dates_for_weekdays(year, month, weekdays):
    day_map = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}
    target_days = [day_map[w] for w in weekdays]
    num_days = calendar.monthrange(year, month)[1]
    dates = []
    for d in range(1, num_days + 1):
        dt = datetime.date(year, month, d)
        if dt.weekday() in target_days: dates.append(dt)
    return dates

def get_master_name(s):
    master_display = {'WARD_AM': 'REPARTO', 'URG_AM': 'URGENZE', 'OR_AM': 'SALA', 'WARD_PM': 'REPARTO', 'URG_PM': 'URGENZE', 'NIGHT': 'NOTTE', 'REP_NIGHT': 'REPERIBILE NOTTE', 'REP_DAY': 'REPERIBILE GIORNO MATTINA E POMERIGGIO'}
    if s in master_display: return master_display[s]
    if s.startswith('OUT_'): return s.replace('OUT_', 'AMBULATORIO ').replace('_AM', '').replace('_', ' ')
    return s

def get_indiv_name(s):
    indiv_display = {'WARD_AM': 'REPARTO MATTINA', 'URG_AM': 'URGENZE MATTINA', 'OR_AM': 'SALA', 'WARD_PM': 'REPARTO POMERIGGIO', 'URG_PM': 'URGENZE POMERIGGIO', 'NIGHT': 'NOTTE', 'REP_NIGHT': 'REPERIBILE NOTTE', 'REP_DAY': 'REPERIBILE GIORNO (MATTINA E POMERIGGIO)'}
    if s in indiv_display: return indiv_display[s]
    if s.startswith('OUT_'): return s.replace('OUT_', 'AMB. ').replace('_AM', '').replace('_', ' ')
    return s

# ==========================================
# 3. DRAFT GENERATOR (Math Engine Only)
# ==========================================
def generate_draft_schedule(year, month, conditional_or_days, manual_festivities, manual_super_holidays, manual_assignments,
                            doctor_capabilities, outpatient_configs, ferie, leave_weeks, desiderate, 
                            private_practice_afternoons, doctors_list, debug_mode=False):
    
    if manual_super_holidays is None: manual_super_holidays = []
    
    doctors = doctors_list
    roster_dates = get_roster_dates(year, month)
    num_days = len(roster_dates)
    date_to_idx = {d.strftime("%Y-%m-%d"): idx for idx, d in enumerate(roster_dates)}
    it_holidays = holidays.IT(years=[year-1, year, year+1])
    ledger = load_historical_counters()
    
    lifetime = {d: get_lifetime_stats(ledger, d) for d in doctors}
    
    sunday_equivalent_days = []
    for day_idx, current_date in enumerate(roster_dates):
        is_sunday = current_date.weekday() == 6
        is_holiday = current_date in it_holidays or current_date.strftime("%Y-%m-%d") in manual_festivities or current_date.strftime("%Y-%m-%d") in manual_super_holidays
        if is_sunday or is_holiday: sunday_equivalent_days.append(day_idx)
            
    manual_keys = set()
    for (d, date_str, s) in manual_assignments:
        if date_str in date_to_idx and d in doctors: manual_keys.add((d, date_to_idx[date_str], s))

    # 🟢 SAFELY RE-ADDED OVERLAP WARNING SCANNER
    all_leave_dates = []
    for d, dates in ferie.items(): all_leave_dates.extend([(d,
