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
    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
        try:
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(st.secrets["GITHUB_REPO"])
            with open(file_path, 'r') as file: content = file.read()
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
    default_recurrence = {
        "OR": {"weekdays": ["Monday", "Friday"], "weeks": ["All"]},
        "PACEMAKER": {"weekdays": [], "weeks": ["All"]},
        "DIMESSI": {"weekdays": [], "weeks": ["All"]},
        "SCOMPENSO": {"weekdays": [], "weeks": ["All"]}
    }
    
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f: data = json.load(f)
            if "docs_base" in data and "capabilities" in data:
                if "recurrence" not in data: data["recurrence"] = default_recurrence
                return data
            elif "doctors" in data:
                docs_base, caps = [], []
                for doc in data["doctors"]:
                    docs_base.append({"Doctor": doc.get("Doctor", ""), "Private Practice (Afternoon)": doc.get("Private Practice (Afternoon)", ""), "Color": doc.get("Color", "White")})
                    cap_row = {"Doctor": doc.get("Doctor", ""), "Ward Preferred": doc.get("Ward Preferred", False), "OR Capable": doc.get("OR Capable", False)}
                    for c in data.get("clinics", default_clinics): cap_row[c] = doc.get(c, False)
                    caps.append(cap_row)
                return {"clinics": data.get("clinics", default_clinics), "docs_base": docs_base, "capabilities": caps, "recurrence": default_recurrence}
        except Exception: pass
            
    docs_base = [{"Doctor": d, "Private Practice (Afternoon)": "", "Color": "White"} for d in default_doctors]
    caps = [{"Doctor": d, "Ward Preferred": d in ["BURGAZZI", "ROBERTI"], "OR Capable": d in ["CACCAMO", "NUCCI"]} for d in default_doctors]
    for c in caps:
        for clin in default_clinics: c[clin] = False
    return {"clinics": default_clinics, "docs_base": docs_base, "capabilities": caps, "recurrence": default_recurrence}

def save_persistent_settings(clinics_list, docs_df, cap_df, recurrence_dict):
    data = {"clinics": clinics_list, "docs_base": docs_df.fillna("").to_dict('records'), "capabilities": cap_df.fillna(False).to_dict('records'), "recurrence": recurrence_dict}
    with open(SETTINGS_FILE, 'w') as f: json.dump(data, f, indent=4)

def load_historical_counters():
    if not os.path.exists(COUNTER_FILE): return {"legacy_baseline": {}}
    try:
        with open(COUNTER_FILE, 'r') as f: data = json.load(f)
        if "legacy_baseline" in data: return data
        else:
            new_data = {"legacy_baseline": {}}
            for doc, stats in data.items(): new_data["legacy_baseline"][doc] = stats
            return new_data
    except: return {"legacy_baseline": {}}

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
    if not mondays: return []
    start_date = mondays[0]
    end_date = mondays[-1] + datetime.timedelta(days=6)
    roster_dates = []
    curr = start_date
    while curr <= end_date:
        roster_dates.append(curr)
        curr += datetime.timedelta(days=1)
    return roster_dates

def calculate_recurring_days(year, month, weekdays, weeks_list):
    if not weekdays or not weeks_list: return []
    roster_dates = get_roster_dates(year, month)
    weeks = [roster_dates[i:i+7] for i in range(0, len(roster_dates), 7)]
    
    day_map = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}
    target_wds = [day_map[w] for w in weekdays if w in day_map]
    
    valid_dates = []
    for week_idx, week in enumerate(weeks):
        w_num = week_idx + 1
        is_valid = False
        if "All" in weeks_list: is_valid = True
        elif "1st" in weeks_list and w_num == 1: is_valid = True
        elif "2nd" in weeks_list and w_num == 2: is_valid = True
        elif "3rd" in weeks_list and w_num == 3: is_valid = True
        elif "4th" in weeks_list and w_num == 4: is_valid = True
        elif "5th" in weeks_list and w_num == 5: is_valid = True
        
        if is_valid:
            for wd in target_wds:
                valid_dates.append(week[wd].strftime("%Y-%m-%d"))
    return valid_dates

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

def compute_monthly_stats(doc_schedule, num_days, roster_dates, day_active, it_holidays, manual_festivities, manual_super_holidays, doctors):
    monthly_stats = {doc: {'nights': 0.0, 'doubles': 0.0, 'saturdays': 0.0, 'sundays': 0.0, 'holidays': 0.0, 'super_holidays': 0.0, 'golden_weekends': 0.0, 'reps': 0.0, 'total_shifts': 0.0, 'total_hours': 0.0} for doc in doctors}
    for doc in doctors:
        for day_idx in range(num_days):
            curr_date = roster_dates[day_idx]
            worked_today = doc_schedule[doc].get(day_idx, [])
            active_shifts = [s for s in worked_today if s in day_active]
            rep_shifts = [s for s in worked_today if s in ['REP_DAY', 'REP_NIGHT']]
            
            for s in worked_today:
                if s == "[UNCOVERED]": continue
                monthly_stats[doc]['total_shifts'] += 1
                if s == 'NIGHT': monthly_stats[doc]['total_hours'] += 12 
                elif 'REP' in s: monthly_stats[doc]['total_hours'] += 0 
                else: monthly_stats[doc]['total_hours'] += 6
            
            if 'NIGHT' in worked_today: monthly_stats[doc]['nights'] += 1
            if len(rep_shifts) > 0: monthly_stats[doc]['reps'] += len(rep_shifts)
            if len(active_shifts) >= 2: monthly_stats[doc]['doubles'] += 1
                
            day_pts = 0.0
            if 'NIGHT' in worked_today: day_pts = 1.0
            else: day_pts = len(active_shifts) * 0.5
            
            if curr_date.weekday() == 5: monthly_stats[doc]['saturdays'] += day_pts
            elif curr_date.weekday() == 6: monthly_stats[doc]['sundays'] += day_pts
                
            if worked_today and (curr_date in it_holidays or curr_date.strftime("%Y-%m-%d") in manual_festivities or curr_date.strftime("%Y-%m-%d") in manual_super_holidays):
                monthly_stats[doc]['holidays'] += 1

            is_easter = (it_holidays.get(curr_date) == "Pasqua di Resurrezione")
            is_proper_holiday = ((curr_date.month == 12 and curr_date.day == 25) or (curr_date.month == 1 and curr_date.day == 1) or (curr_date.month == 4 and curr_date.day == 25) or (curr_date.month == 6 and curr_date.day in [1, 2]) or (curr_date.month == 8 and curr_date.day == 15) or is_easter or curr_date.strftime("%Y-%m-%d") in manual_super_holidays)
            is_eve = (curr_date.month == 12 and curr_date.day in [24, 31])
            
            sh_pts = 0.0
            if (is_eve and 'NIGHT' in worked_today) or (is_proper_holiday and len(active_shifts) > 0): sh_pts = 1.0
            elif (is_proper_holiday and len(rep_shifts) > 0) or (is_eve and 'REP_NIGHT' in worked_today): sh_pts = 0.5
            if sh_pts > 0: monthly_stats[doc]['super_holidays'] += sh_pts

        for week_start_idx in range(0, num_days, 7):
            ruined_fri = 'NIGHT' in doc_schedule[doc].get(week_start_idx + 4, []) or 'REP_NIGHT' in doc_schedule[doc].get(week_start_idx + 4, [])
            if not ruined_fri and not doc_schedule[doc].get(week_start_idx + 5, []) and not doc_schedule[doc].get(week_start_idx + 6, []):
                monthly_stats[doc]['golden_weekends'] += 1
    return monthly_stats

# ==========================================
# 3. DRAFT GENERATOR (Math Engine Only)
# ==========================================
def generate_draft_schedule(year, month, conditional_or_days, manual_festivities, manual_super_holidays, manual_assignments,
                            doctor_capabilities, outpatient_configs, daily_absences, private_practice_afternoons, 
                            doctors_list, debug_mode=False, resolution_toggles=None):
    
    if manual_super_holidays is None: manual_super_holidays = []
    if resolution_toggles is None: resolution_toggles = {}
    
    allow_understaffing = resolution_toggles.get("allow_understaffing", False)
    ignore_34h = resolution_toggles.get("ignore_34h", False)
    
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

    fully_off_dates = []
    for (d, date_str), val in daily_absences.items():
        if val in ['F', 'X']: fully_off_dates.append(date_str)
                    
    day_counts = Counter(fully_off_dates)
    overlaps = [date for date, count in day_counts.items() if count > 1]
    overlap_warning = f"⚠️ Notice: Multiple doctors requested completely off (F/X) on: {', '.join(overlaps)}" if overlaps else ""

    doc_unavailable_set = {d: set() for d in doctors}
    doc_contract_off_set = {d: set() for d in doctors}

    for (d, date_str), val in daily_absences.items():
        if val in ['F', 'X']: doc_unavailable_set[d].add(date_str)
        if val == 'F': doc_contract_off_set[d].add(date_str)
                            
    doc_active_days = {d: max(1, num_days - len(doc_unavailable_set[d])) for d in doctors}
    doc_contract_days = {d: max(1, num_days - len(doc_contract_off_set[d])) for d in doctors}

    shifts = ['WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']
    for out_type in outpatient_configs.keys(): shifts.append(f'OUT_{out_type}_AM')
    day_active = [s for s in shifts if s not in ['NIGHT', 'REP_NIGHT', 'REP_DAY']]
    
    total_available_hours = 0
    diagnostic_errors = []
    
    for day_idx in range(num_days):
        current_date_str = roster_dates[day_idx].strftime("%Y-%m-%d")
        
        req_shifts = 0
        if day_idx in sunday_equivalent_days:
            req_shifts = 3 
            total_available_hours += 24
        else:
            req_shifts = 5 
            total_available_hours += 36
            if current_date_str in conditional_or_days:
                req_shifts += 1
                total_available_hours += 6
                or_capable_docs = [d for d in doctors if d != 'GAUDENZI' and doctor_capabilities.get(d, {}).get("OR Capable", False) and current_date_str not in doc_unavailable_set[d]]
                if not or_capable_docs and not any((d == 'GAUDENZI' and day == day_idx and s == 'OR_AM') for (d, day, s) in manual_keys):
                    diagnostic_errors.append(f"**{current_date_str}:** Operating Room is scheduled, but NO OR-capable doctors are available.")
                
            for out_type, config in outpatient_configs.items():
                if current_date_str in config['days']:
                    req_shifts += 1
                    total_available_hours += 6
                    clin_capable = [d for d in doctors if d != 'GAUDENZI' and d in config['capable'] and current_date_str not in doc_unavailable_set[d]]
                    if not clin_capable and not any((d == 'GAUDENZI' and day == day_idx and s == f'OUT_{out_type}_AM') for (d, day, s) in manual_keys):
                        diagnostic_errors.append(f"**{current_date_str}:** Clinic {out_type} is scheduled, but NO capable doctors are available.")

        unavailable = sum(1 for d in doctors if d != 'GAUDENZI' and current_date_str in doc_unavailable_set[d])
        available_bodies = len([d for d in doctors if d != 'GAUDENZI']) - unavailable
        if any((d == 'GAUDENZI' and day == day_idx) for (d, day, s) in manual_keys): available_bodies += 1
        
        if available_bodies < req_shifts:
            diagnostic_errors.append(f"**{current_date_str}:** You scheduled {req_shifts} shifts, but only {available_bodies} doctors are available to work.")

    if diagnostic_errors and not allow_understaffing and not debug_mode:
        return False, None, "### 🚨 Deep Diagnostic Pre-Flight Failed\nWe caught several mathematical impossibilities before generating:\n\n" + "\n".join(f"- {err}" for err in diagnostic_errors) + "\n\n*Use the **Resolution Options** below to override these rules.*"

    gaudenzi_manual_active_hrs = sum(12 if s == 'NIGHT' else (6 if 'REP' not in s else 0) for (d, day, s) in manual_keys if d == 'GAUDENZI')
    total_available_hours -= gaudenzi_manual_active_hrs
    
    total_required_hours_min = sum(max(0, int((doc_contract_days[d] / 7.0) * 34) - 12) for d in doctors if d != 'GAUDENZI')
    if total_required_hours_min > total_available_hours and not debug_mode and not ignore_34h:
        return False, None, f"### 🚨 Contractual Hour Deficit\nThe 34h clinical minimum requires at least {total_required_hours_min}h total from active doctors, but department only has {total_available_hours}h scheduled.\n\n*Solution: Add clinics, or use the Resolution Options below to ignore the 34h rule.*"

    pass_levels = [3] if debug_mode else [0, 1, 2, 3]
    
    for pass_level in pass_levels:
        model = cp_model.CpModel()
        work = {}
        for d in doctors:
            for day_idx in range(num_days):
                for s in shifts: work[(d, day_idx, s)] = model.NewBoolVar(f'work_{d}_{day_idx}_{s}')

        objective_terms = []

        for (d, date_str, s) in manual_assignments:
            if date_str in date_to_idx and d in doctors:
                key = (d, date_to_idx[date_str], s)
                if key in work: model.Add(work[key] == 1)

        if 'GAUDENZI' in doctors:
            for day_idx in range(num_days):
                for s in shifts:
                    if ('GAUDENZI', day_idx, s) not in manual_keys: model.Add(work[('GAUDENZI', day_idx, s)] == 0)

        for d in doctors:
            if d == 'GAUDENZI': continue
            for day_idx in range(num_days):
                date_str = roster_dates[day_idx].strftime("%Y-%m-%d")
                val = daily_absences.get((d, date_str), "")
                
                if val in ['F', 'X']:
                    for s in shifts: model.Add(work[(d, day_idx, s)] == 0)
                elif val == 'P':
                    for s in ['WARD_PM', 'URG_PM', 'NIGHT', 'REP_NIGHT', 'REP_DAY']:
                        if s in shifts: model.Add(work[(d, day_idx, s)] == 0)
                elif val == 'N':
                    for s in ['NIGHT', 'REP_NIGHT']:
                        if s in shifts: model.Add(work[(d, day_idx, s)] == 0)
                    
        for day_idx in range(num_days):
            current_date_str = roster_dates[day_idx].strftime("%Y-%m-%d")
            model.AddExactlyOne(work[(d, day_idx, 'NIGHT')] for d in doctors)
            model.AddExactlyOne(work[(d, day_idx, 'REP_NIGHT')] for d in doctors)
            
            def add_shift_constraint(s_name, allowed_docs):
                if allow_understaffing and s_name not in ['NIGHT', 'REP_NIGHT']:
                    covered = model.NewBoolVar(f'cov_{day_idx}_{s_name}')
                    model.Add(sum(work[(d, day_idx, s_name)] for d in allowed_docs) == covered)
                    objective_terms.append(100000 * covered) 
                    for d in doctors:
                        if d not in allowed_docs: model.Add(work[(d, day_idx, s_name)] == 0)
                else:
                    model.AddExactlyOne(work[(d, day_idx, s_name)] for d in allowed_docs)
                    for d in doctors:
                        if d not in allowed_docs: model.Add(work[(d, day_idx, s_name)] == 0)

            if day_idx not in sunday_equivalent_days: 
                add_shift_constraint('WARD_AM', doctors)
                add_shift_constraint('URG_AM', doctors)
                add_shift_constraint('WARD_PM', doctors)
                add_shift_constraint('URG_PM', doctors)
                for d in doctors: model.Add(work[(d, day_idx, 'REP_DAY')] == 0)
                
                if current_date_str in conditional_or_days:
                    capable_docs = [d for d in doctors if doctor_capabilities.get(d, {}).get("OR Capable", False) or (d == 'GAUDENZI')]
                    add_shift_constraint('OR_AM', capable_docs)
                else:
                    for d in doctors: model.Add(work[(d, day_idx, 'OR_AM')] == 0)
                    
                for out_type, config in outpatient_configs.items():
                    s_name = f'OUT_{out_type}_AM'
                    if current_date_str in config['days']:
                        capable_docs = [d for d in doctors if d in config['capable'] or (d == 'GAUDENZI')]
                        add_shift_constraint(s_name, capable_docs)
                    else:
                        for d in doctors: model.Add(work[(d, day_idx, s_name)] == 0)
            else: 
                add_shift_constraint('WARD_AM', doctors)
                add_shift_constraint('WARD_PM', doctors)
                add_shift_constraint('REP_DAY', doctors)
                for d in doctors:
                    for s in shifts:
                        if s not in ['WARD_AM', 'WARD_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']: model.Add(work[(d, day_idx, s)] == 0)

        day_name_to_num = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}

        for d in doctors:
            if d == 'GAUDENZI': continue
            
            pp_day_num = day_name_to_num.get(private_practice_afternoons.get(d, "").strip().capitalize(), -1)
            for day_idx in range(num_days):
                weekday = roster_dates[day_idx].weekday()
                
                # 🟢 FINIZIO OVERRIDE BLOCK (HARD RULES)
                if d == 'FINIZIO':
                    model.Add(work[(d, day_idx, 'NIGHT')] == 0)
                    model.Add(work[(d, day_idx, 'REP_NIGHT')] == 0)
                    model.Add(sum(work[(d, day_idx, s)] for s in shifts) <= 1)
                else:
                    if weekday == pp_day_num:
                        for s in shifts:
                            if s.endswith('_PM') or s == 'REP_DAY': model.Add(work[(d, day_idx, s)] == 0)
                        objective_terms.append(500 * work[(d, day_idx, 'NIGHT')])
                        objective_terms.append(300 * work[(d, day_idx, 'REP_NIGHT')])
                            
                    next_day_date = roster_dates[day_idx] + datetime.timedelta(days=1)
                    if next_day_date.weekday() == pp_day_num:
                        model.Add(work[(d, day_idx, 'NIGHT')] == 0)
                        model.Add(work[(d, day_idx, 'REP_NIGHT')] == 0)
                    
                    model.Add(work[(d, day_idx, 'NIGHT')] == 0).OnlyEnforceIf(work[(d, day_idx, 'REP_NIGHT')])
                    model.Add(sum(work[(d, day_idx, s)] for s in day_active) == 0).OnlyEnforceIf(work[(d, day_idx, 'REP_DAY')])
                    
                    if day_idx in sunday_equivalent_days:
                        rep_and_night = model.NewBoolVar('')
                        model.AddBoolAnd([work[(d, day_idx, 'NIGHT')], work[(d, day_idx, 'REP_DAY')]]).OnlyEnforceIf(rep_and_night)
                        objective_terms.append(1000 * rep_and_night) 
                        
                    model.Add(work[(d, day_idx, 'REP_NIGHT')] == 0).OnlyEnforceIf(work[(d, day_idx, 'REP_DAY')])
                    
                    if day_idx + 1 < num_days:
                        model.Add(sum(work[(d, day_idx + 1, s)] for s in day_active) == 0).OnlyEnforceIf(work[(d, day_idx, 'REP_NIGHT')])
                        model.Add(work[(d, day_idx + 1, 'REP_DAY')] == 0).OnlyEnforceIf(work[(d, day_idx, 'REP_NIGHT')])
                    
                    if day_idx < num_days - 1:
                        is_pre_night = work[(d, day_idx + 1, 'NIGHT')]
                        
                        if pass_level < 3: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 2)
                        
                        if day_idx not in sunday_equivalent_days:
                            if pass_level < 3: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 1).OnlyEnforceIf(is_pre_night.Not())
                            
                        pre_night_double = model.NewBoolVar('')
                        model.Add(sum(work[(d, day_idx, s)] for s in day_active) == 2).OnlyEnforceIf(pre_night_double)
                        model.Add(sum(work[(d, day_idx, s)] for s in day_active) != 2).OnlyEnforceIf(pre_night_double.Not())
                        
                        ideal_pre_night = model.NewBoolVar('')
                        model.AddBoolAnd([is_pre_night, pre_night_double, work[(d, day_idx, 'REP_NIGHT')]]).OnlyEnforceIf(ideal_pre_night)
                        
                        if pass_level == 0:
                            model.Add(ideal_pre_night == 1).OnlyEnforceIf(is_pre_night)
                        else:
                            objective_terms.append(2000 * ideal_pre_night)
                    else:
                        if pass_level < 3:
                            if day_idx not in sunday_equivalent_days: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 1)
                            else: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 2)
                    
                    other_than_night = [s for s in shifts if s != 'NIGHT']
                    model.Add(sum(work[(d, day_idx, s)] for s in other_than_night) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
                    if day_idx + 1 < num_days: model.Add(sum(work[(d, day_idx + 1, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
                    if day_idx + 2 < num_days: model.Add(sum(work[(d, day_idx + 2, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])

            # 🟢 BOSS LOGIC: THE WEEKEND PENDULUM & ROBIN HOOD EQUITY
            doctor_gws = []
            for week_start_idx in range(0, num_days, 7):
                fri_idx, sat_idx, sun_idx = week_start_idx + 4, week_start_idx + 5, week_start_idx + 6
                gw_var = model.NewBoolVar(f'gw_{d}_{week_start_idx}')
                
                ruining_shifts = [work[(d, fri_idx, 'NIGHT')], work[(d, fri_idx, 'REP_NIGHT')]]
                for s in shifts: ruining_shifts.extend([work[(d, sat_idx, s)], work[(d, sun_idx, s)]])
                model.Add(sum(ruining_shifts) == 0).OnlyEnforceIf(gw_var)
                model.Add(sum(ruining_shifts) > 0).OnlyEnforceIf(gw_var.Not())
                doctor_gws.append(gw_var)
                
                gw_pts = int(lifetime[d]['golden_weekends'] * 50)
                objective_terms.append((1000 - gw_pts) * gw_var)
                
                if week_start_idx >= 7:
                    prev_gw_var = doctor_gws[-2]
                    two_gws = model.NewBoolVar('')
                    model.AddBoolAnd([gw_var, prev_gw_var]).OnlyEnforceIf(two_gws)
                    two_works = model.NewBoolVar('')
                    model.AddBoolAnd([gw_var.Not(), prev_gw_var.Not()]).OnlyEnforceIf(two_works)
                    
                    if pass_level == 0:
                        model.Add(two_works == 0)
                        model.Add(two_gws == 0)
                    else:
                        objective_terms.append(-5000 * two_works)
                        objective_terms.append(-3000 * two_gws)
                
            if pass_level < 2: model.Add(sum(doctor_gws) >= 1)

        # 🟢 PRE-VACATION NIGHT (Absolute Priority)
        for d in doctors:
            if d == 'GAUDENZI' or d == 'FINIZIO': continue
            for week_start_idx in range(0, num_days, 7):
                week_days = range(week_start_idx, min(week_start_idx + 7, num_days))
                if len(week_days) == 7:
                    is_off_week = all(daily_absences.get((d, roster_dates[day_idx].strftime("%Y-%m-%d")), "") in ['F', 'X'] for day_idx in week_days)
                    if is_off_week:
                        if 0 <= week_start_idx - 1 < num_days:
                            for s in shifts: model.Add(work[(d, week_start_idx - 1, s)] == 0)
                        if 0 <= week_start_idx - 4 < num_days:
                            objective_terms.append(5000 * work[(d, week_start_idx - 4, 'NIGHT')])

        weeks = [range(i, i + 7) for i in range(0, num_days, 7)]

        # 🟢 MIN-MAX FAIRNESS FOR HOUR DISTRIBUTION
        diff_p_vars = {}
        diff_m_vars = {}
        
        for d in doctors:
            if d == 'GAUDENZI': continue
            tgt_hours = int((doc_contract_days[d] / 7.0) * 34)
            active_expr = sum(work[(d, day_idx, s)] * 6 for day_idx in range(num_days) for s in day_active) + \
                          sum(work[(d, day_idx, 'NIGHT')] * 12 for day_idx in range(num_days))
                          
            active_var = model.NewIntVar(0, 1000, f'act_hrs_{d}')
            model.Add(active_var == active_expr)
            
            dp = model.NewIntVar(0, 1000, f'dp_{d}')
            dm = model.NewIntVar(0, 1000, f'dm_{d}')
            model.Add(active_var - tgt_hours == dp - dm)
            
            diff_p_vars[d] = dp
            diff_m_vars[d] = dm
            
            if pass_level == 0 and not ignore_34h:
                model.Add(dm <= 6)
            elif pass_level == 1 and not ignore_34h:
                model.Add(dm <= 18)

            objective_terms.append(-100 * dp)
            objective_terms.append(-500 * dm) 
            
        max_dm = model.NewIntVar(0, 1000, 'max_dm')
        max_dp = model.NewIntVar(0, 1000, 'max_dp')
        
        for d in doctors:
            if d == 'GAUDENZI': continue
            model.Add(diff_m_vars[d] <= max_dm)
            model.Add(diff_p_vars[d] <= max_dp)
            
        objective_terms.append(-2000 * max_dm)
        objective_terms.append(-1000 * max_dp)

        # 🟢 WARD CONTINUITY
        if pass_level < 2:
            for d in doctors:
                if d == 'GAUDENZI': continue
                if doctor_capabilities.get(d, {}).get("Ward Preferred", False):
                    for day_idx in range(num_days):
                        objective_terms.extend([20 * work[(d, day_idx, 'WARD_AM')], 20 * work[(d, day_idx, 'WARD_PM')]])

            for d in doctors:
                if d == 'GAUDENZI': continue
                for day_idx in range(num_days - 1):
                    am_pm = model.NewBoolVar('')
                    model.AddBoolAnd([work[(d, day_idx, 'WARD_AM')], work[(d, day_idx+1, 'WARD_PM')]]).OnlyEnforceIf(am_pm)
                    objective_terms.append(150 * am_pm) 
                    pm_am = model.NewBoolVar('')
                    model.AddBoolAnd([work[(d, day_idx, 'WARD_PM')], work[(d, day_idx+1, 'WARD_AM')]]).OnlyEnforceIf(pm_am)
                    objective_terms.append(150 * pm_am)
                    
                for w_idx in range(len(weeks)):
                    ward_active = model.NewBoolVar('')
                    model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(ward_active)
                    model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(ward_active.Not())
                    objective_terms.append(-400 * ward_active)
                    
                    if w_idx < len(weeks) - 1:
                        ward_next_active = model.NewBoolVar('')
                        model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(ward_next_active)
                        model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(ward_next_active.Not())
                        consecutive_ward = model.NewBoolVar('')
                        model.AddBoolAnd([ward_active, ward_next_active]).OnlyEnforceIf(consecutive_ward)
                        
                        if pass_level == 0:
                            model.Add(consecutive_ward == 0)
                        else:
                            objective_terms.append(-2000 * consecutive_ward)

        for day_idx in sunday_equivalent_days:
            for d in doctors:
                if d == 'GAUDENZI': continue
                ward_double = model.NewBoolVar('')
                model.AddBoolAnd([work[(d, day_idx, 'WARD_AM')], work[(d, day_idx, 'WARD_PM')]]).OnlyEnforceIf(ward_double)
                objective_terms.append(10 * ward_double) 
                urg_double = model.NewBoolVar('')
                model.AddBoolAnd([work[(d, day_idx, 'URG_AM')], work[(d, day_idx, 'URG_PM')]]).OnlyEnforceIf(urg_double)
                objective_terms.append(10 * urg_double) 

        # 🟢 HEAVY LIFETIME EQUITY GRAVITY
        for d in doctors:
            if d == 'GAUDENZI': continue
            n_pts = int(lifetime[d]['nights'] * 50)           
            r_pts = int(lifetime[d]['reps'] * 30)             
            sat_pts = int(lifetime[d]['saturdays'] * 30)      
            sun_pts = int(lifetime[d]['sundays'] * 30)        
            sh_pts = int(lifetime[d]['super_holidays'] * 50)  
            
            for day_idx in range(num_days):
                curr_date = roster_dates[day_idx]
                
                objective_terms.append(-1 * n_pts * work[(d, day_idx, 'NIGHT')])
                objective_terms.append(-1 * int(r_pts/2) * work[(d, day_idx, 'REP_NIGHT')])
                if day_idx in sunday_equivalent_days:
                    objective_terms.append(-1 * int(r_pts/2) * work[(d, day_idx, 'REP_DAY')])
                
                if curr_date.weekday() == 5:
                    objective_terms.append(-1 * sat_pts * work[(d, day_idx, 'NIGHT')])
                    for s in day_active: objective_terms.append(-1 * int(sat_pts/2) * work[(d, day_idx, s)])
                if curr_date.weekday() == 6:
                    objective_terms.append(-1 * sun_pts * work[(d, day_idx, 'NIGHT')])
                    for s in day_active: objective_terms.append(-1 * int(sun_pts/2) * work[(d, day_idx, s)])
                
                if curr_date in it_holidays or curr_date.strftime("%Y-%m-%d") in manual_festivities or curr_date.strftime("%Y-%m-%d") in manual_super_holidays: 
                    worked_any = model.NewBoolVar('')
                    model.Add(sum(work[(d, day_idx, s)] for s in shifts) > 0).OnlyEnforceIf(worked_any)
                    model.Add(sum(work[(d, day_idx, s)] for s in shifts) == 0).OnlyEnforceIf(worked_any.Not())
                    objective_terms.append(-4 * int(lifetime[d]['holidays']) * worked_any)
                    
                is_easter = (it_holidays.get(curr_date) == "Pasqua di Resurrezione")
                is_proper_holiday = ((curr_date.month == 12 and curr_date.day == 25) or 
                                     (curr_date.month == 1 and curr_date.day == 1) or
                                     (curr_date.month == 4 and curr_date.day == 25) or
                                     (curr_date.month == 6 and curr_date.day in [1, 2]) or
                                     (curr_date.month == 8 and curr_date.day == 15) or is_easter or
                                     curr_date.strftime("%Y-%m-%d") in manual_super_holidays)
                is_eve = (curr_date.month == 12 and curr_date.day in [24, 31])
                    
                if is_eve:
                    objective_terms.append(-1 * sh_pts * work[(d, day_idx, 'NIGHT')])
                    objective_terms.append(-1 * int(sh_pts/2) * work[(d, day_idx, 'REP_NIGHT')])
                if is_proper_holiday:
                    for s in day_active:
                        objective_terms.append(-1 * sh_pts * work[(d, day_idx, s)])
                    objective_terms.append(-1 * int(sh_pts/2) * work[(d, day_idx, 'REP_DAY')])
                    objective_terms.append(-1 * int(sh_pts/2) * work[(d, day_idx, 'REP_NIGHT')])

        model.Maximize(sum(objective_terms))
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 15.0 
        status = solver.Solve(model)
        
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            draft_grids = {}
            for w_idx, week_start_idx in enumerate(range(0, num_days, 7)):
                week_dates = roster_dates[week_start_idx : week_start_idx + 7]
                col_names = [d.strftime("%Y-%m-%d") for d in week_dates]
                df_dict = {"Shift": shifts.copy()}
                for c in col_names: df_dict[c] = [""] * len(shifts)
                df = pd.DataFrame(df_dict)
                
                for i, d_date in enumerate(week_dates):
                    day_idx = week_start_idx + i
                    d_str = d_date.strftime("%Y-%m-%d")
                    if day_idx >= num_days: continue
                    for s in shifts:
                        assigned_doc = ""
                        for d in doctors:
                            if solver.Value(work[(d, day_idx, s)]) == 1:
                                assigned_doc = d
                        if not assigned_doc and allow_understaffing and s not in ['NIGHT', 'REP_NIGHT']:
                            is_active = True
                            if s == 'OR_AM' and d_str not in conditional_or_days: is_active = False
                            elif s == 'REP_DAY' and day_idx not in sunday_equivalent_days: is_active = False
                            elif s.startswith('OUT_'):
                                c_name = s.replace('OUT_', '').replace('_AM', '')
                                if d_str not in outpatient_configs.get(c_name, {}).get('days', []): is_active = False
                            if is_active: assigned_doc = "[UNCOVERED]"
                        if assigned_doc: df.loc[df["Shift"] == s, d_str] = assigned_doc
                draft_grids[w_idx] = df.copy()
            
            warning_msg = ""
            if pass_level == 1: warning_msg = "⚠️ GEAR 2: Hard consecutive ward/weekend limits were relaxed to find a schedule."
            elif pass_level == 2: warning_msg = "⚠️ GEAR 3: Monthly hour constraints were loosened completely. Equity is handled via Min-Max fairness."
            elif pass_level == 3: warning_msg = "⚠️ SURVIVAL GEAR: Maximum shift caps and Monto packages removed. Verify manually."
            elif pass_level == 4: warning_msg = "🚨 DEBUG MODE: All equity limits, max deficits, and floors disabled."
            
            if allow_understaffing or ignore_34h:
                warning_msg += "\n\n🛠️ **RESOLUTION MODE ACTIVE:** Overrides were used to force completion."
            if overlap_warning: warning_msg += f"\n\n{overlap_warning}"
            return True, draft_grids, warning_msg

    return False, None, "🛑 Constraints are too tight. Even after shifting to Survival Gear, the algorithm cannot find a schedule. Try using the Resolution Options below."

# ==========================================
# 4. EXPORT ENGINE (Reads Edited UI Draft)
# ==========================================
def process_and_export_schedule(edited_weekly_grids, year, month, conditional_or_days, manual_festivities, manual_super_holidays,
                                outpatient_configs, daily_absences, doctor_colors, doctors_list):
    
    if manual_super_holidays is None: manual_super_holidays = []
    
    doctors = doctors_list
    roster_dates = get_roster_dates(year, month)
    num_days = len(roster_dates)
    date_to_idx = {d.strftime("%Y-%m-%d"): idx for idx, d in enumerate(roster_dates)}
    it_holidays = holidays.IT(years=[year-1, year, year+1])
    
    ledger = load_historical_counters()

    sunday_equivalent_days = []
    for day_idx, current_date in enumerate(roster_dates):
        is_sunday = current_date.weekday() == 6
        is_holiday = current_date in it_holidays or current_date.strftime("%Y-%m-%d") in manual_festivities or current_date.strftime("%Y-%m-%d") in manual_super_holidays
        if is_sunday or is_holiday: sunday_equivalent_days.append(day_idx)
            
    doc_contract_off_set = {d: set() for d in doctors}
    for (d, date_str), val in daily_absences.items():
        if val == 'F' and date_str in date_to_idx: 
            doc_contract_off_set[d].add(date_str)
                        
    doc_contract_days = {d: max(1, num_days - len(doc_contract_off_set[d])) for d in doctors}

    doc_schedule = {doc: {day_idx: [] for day_idx in range(num_days)} for doc in doctors}
    for w_idx, df_w in edited_weekly_grids.items():
        week_dates = roster_dates[w_idx*7 : (w_idx*7)+7]
        for idx, row in df_w.iterrows():
            s = row["Shift"]
            for i, d_date in enumerate(week_dates):
                day_idx = (w_idx * 7) + i
                if day_idx >= num_days: continue
                d_str = d_date.strftime("%Y-%m-%d")
                assigned_doc = str(row.get(d_str, "")).strip().upper()
                if assigned_doc and assigned_doc in doctors:
                    doc_schedule[assigned_doc][day_idx].append(s)

    shifts = ['WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']
    for out_type in outpatient_configs.keys(): shifts.append(f'OUT_{out_type}_AM')
    day_active = [s for s in shifts if s not in ['NIGHT', 'REP_NIGHT', 'REP_DAY']]

    monthly_stats = compute_monthly_stats(doc_schedule, num_days, roster_dates, day_active, it_holidays, [], manual_super_holidays, doctors)

    month_key = f"{year}-{month:02d}"
    ledger[month_key] = monthly_stats
    with open(COUNTER_FILE, 'w') as f: json.dump(ledger, f, indent=4)
    push_to_github(COUNTER_FILE, f"Auto-sync: Published Roster for {month_key}")
    
    lifetime = {d: get_lifetime_stats(ledger, d) for d in doctors}

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    worksheet = workbook.add_worksheet('Master Schedule')
    header_format = workbook.add_format({'bold': True, 'bg_color': '#E8E8E8', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    section_format = workbook.add_format({'bold': True, 'bg_color': '#4F4F4F', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    shift_format = workbook.add_format({'bold': True, 'bg_color': '#F5F5F5', 'border': 1, 'align': 'left'})
    empty_border = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    gray_dash_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': '#D3D3D3'})
    uncovered_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': 'red', 'bold': True})
    
    doc_formats = {}
    for d in doctors:
        hex_color = doctor_colors.get(d, '#FFFFFF')
        doc_formats[d] = workbook.add_format({'bg_color': hex_color, 'border': 1, 'align': 'center', 'bold': True, 'valign': 'vcenter'})

    row_cursor = 0
    for w_idx, week_start_idx in enumerate(range(0, num_days, 7)):
        worksheet.write(row_cursor, 0, f"WEEK {w_idx + 1}", header_format)
        for i in range(7):
            day_date = roster_dates[week_start_idx + i]
            day_str = f"{calendar.day_abbr[day_date.weekday()]} {day_date.strftime('%b %d')}"
            if week_start_idx + i in sunday_equivalent_days: day_str += " (*)"
            worksheet.write(row_cursor, i + 1, day_str, header_format)
        
        row_cursor += 1
        
        week_shifts_am = ['WARD_AM', 'URG_AM']
        week_shifts_pm = ['WARD_PM', 'URG_PM']
        
        or_open = False
        outpatient_open = {out_type: False for out_type in outpatient_configs.keys()}
        rep_day_open = False
        
        for i in range(7):
            day_idx = week_start_idx + i
            if day_idx >= num_days: continue
            d_str = roster_dates[day_idx].strftime("%Y-%m-%d")
            if day_idx in sunday_equivalent_days: rep_day_open = True
            if d_str in conditional_or_days: or_open = True
            for out_type, config in outpatient_configs.items():
                if d_str in config['days']: outpatient_open[out_type] = True
                
        if or_open: week_shifts_am.append('OR_AM')
        for out_type, is_open in outpatient_open.items():
            if is_open: week_shifts_am.append(f'OUT_{out_type}_AM')
        
        rep_shifts = ['REP_NIGHT']
        if rep_day_open: rep_shifts.append('REP_DAY')

        for section_name, section_shifts in [("MATTINA", week_shifts_am), ("POMERIGGIO", week_shifts_pm), ("NOTTE E REPERIBILITÀ", ['NIGHT'] + rep_shifts)]:
            worksheet.merge_range(row_cursor, 0, row_cursor, 7, section_name, section_format)
            row_cursor += 1
            for s in section_shifts:
                worksheet.write(row_cursor, 0, get_master_name(s), shift_format)
                for i in range(7):
                    day_idx = week_start_idx + i
                    if day_idx >= num_days:
                        worksheet.write(row_cursor, i + 1, "", empty_border)
                        continue
                    
                    assigned_doc = ""
                    ui_val = str(edited_weekly_grids[w_idx].loc[edited_weekly_grids[w_idx]["Shift"] == s, roster_dates[day_idx].strftime("%Y-%m-%d")].values[0]).strip().upper()
                    
                    if ui_val == "[UNCOVERED]":
                        worksheet.write(row_cursor, i + 1, "UNCOVERED", uncovered_format)
                    elif ui_val in doctors:
                        worksheet.write(row_cursor, i + 1, ui_val, doc_formats[ui_val])
                    else:
                        is_active_today = True
                        d_str = roster_dates[day_idx].strftime("%Y-%m-%d")
                        if s == 'OR_AM' and d_str not in conditional_or_days: is_active_today = False
                        if s == 'REP_DAY' and day_idx not in sunday_equivalent_days: is_active_today = False
                        if s.startswith('OUT_'):
                            c_name = s.replace('OUT_', '').replace('_AM', '')
                            if d_str not in outpatient_configs.get(c_name, {}).get('days', []): is_active_today = False
                        
                        if is_active_today: worksheet.write(row_cursor, i + 1, "", empty_border)
                        else: worksheet.write(row_cursor, i + 1, "-", gray_dash_format)
                row_cursor += 1
        row_cursor += 1

    dashboard_header = ['Doctor', 'Proportional Target Hours', 'Actual Active Hours', 'Difference (+/-)',
                        'Monthly Nights', 'Monthly On-Call (Rep)', 'Monthly Doubles', 'Monthly Saturdays', 'Monthly Sundays', 'Monthly Holidays', 'Monthly Super Hols', 'Monthly Golden Wknds',
                        'LIFETIME Nights', 'LIFETIME On-Call', 'LIFETIME Doubles', 'LIFETIME Saturdays', 'LIFETIME Sundays', 'LIFETIME Holidays', 'LIFETIME Super Hols', 'LIFETIME Golden Wknds']
    
    worksheet.write(row_cursor, 0, "DOCTOR STATISTICS DASHBOARD", header_format)
    row_cursor += 1
    
    for col_idx, h in enumerate(dashboard_header): worksheet.write(row_cursor, col_idx, h, header_format)
    row_cursor += 1
    
    for doc in doctors:
        if doc == 'GAUDENZI':
            doc_target_display = "MANUAL"
            actual_hours = monthly_stats[doc]['total_hours']
            difference_display = "N/A"
        else:
            doc_target = int((doc_contract_days[doc] / 7) * 34)
            doc_target_display = doc_target
            actual_hours = monthly_stats[doc]['total_hours']
            difference = actual_hours - doc_target
            difference_display = f"+{difference}" if difference > 0 else str(difference)
        
        data_row = [
            doc, doc_target_display, actual_hours, difference_display,
            monthly_stats[doc]['nights'], monthly_stats[doc]['reps'], monthly_stats[doc]['doubles'], f"{monthly_stats[doc]['saturdays']:.1f}", f"{monthly_stats[doc]['sundays']:.1f}", monthly_stats[doc]['holidays'], f"{monthly_stats[doc]['super_holidays']:.1f}", monthly_stats[doc]['golden_weekends'],
            lifetime[doc]['nights'], lifetime[doc]['reps'], lifetime[doc]['doubles'], f"{lifetime[doc]['saturdays']:.1f}", f"{lifetime[doc]['sundays']:.1f}", lifetime[doc]['holidays'], f"{lifetime[doc]['super_holidays']:.1f}", lifetime[doc]['golden_weekends']
        ]
        for col_idx, val in enumerate(data_row): worksheet.write(row_cursor, col_idx, val, doc_formats[doc] if col_idx == 0 else empty_border)
        row_cursor += 1

    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 7, 18)

    for doc in doctors:
        doc_sheet = workbook.add_worksheet(doc)
        doc_header_fmt = workbook.add_format({'bold': True, 'bg_color': doctor_colors.get(doc, '#E8E8E8'), 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        off_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'font_color': '#A0A0A0', 'italic': True})
        holiday_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

        days_header = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        for col_idx, day_name in enumerate(days_header): doc_sheet.write(0, col_idx, day_name, doc_header_fmt)
        doc_sheet.set_column(0, 6, 16) 

        row_c = 1
        for w_idx, week_start_idx in enumerate(range(0, num_days, 7)):
            doc_sheet.set_row(row_c, 55) 
            for i in range(7):
                day_idx = week_start_idx + i
                if day_idx >= num_days: continue
                curr_date = roster_dates[day_idx]
                
                my_shifts = []
                for s in shifts:
                    if s in doc_schedule[doc][day_idx]:
                        if 'REP' in s: my_shifts.append("📞 " + get_indiv_name(s))
                        else: my_shifts.append(get_indiv_name(s))
                        
                is_holiday_or_weekend = (curr_date.weekday() >= 5 or day_idx in sunday_equivalent_days)
                if my_shifts:
                    shift_text = f"{curr_date.strftime('%b %d')}\n\n" + "\n".join(my_shifts)
                    fmt = holiday_fmt if is_holiday_or_weekend else cell_fmt
                else:
                    shift_text = f"{curr_date.strftime('%b %d')}\n\nOFF"
                    fmt = holiday_fmt if is_holiday_or_weekend else off_fmt

                doc_sheet.write(row_c, i, shift_text, fmt)
            row_c += 1

    rep_sheet = workbook.add_worksheet('Reperibilità Ufficiale')
    rep_sheet.set_column(0, 0, 20)
    rep_sheet.set_column(1, 7, 18)
    title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left', 'valign': 'vcenter'})
    rep_sheet.write(0, 0, "CALENDARIO REPERIBILITA' UFFICIALE", title_fmt)
    
    row_c = 2
    for w_idx, week_start_idx in enumerate(range(0, num_days, 7)):
        rep_sheet.write(row_c, 0, f"WEEK {w_idx + 1}", header_format)
        for i in range(7):
            day_date = roster_dates[week_start_idx + i]
            day_str = f"{calendar.day_abbr[day_date.weekday()]} {day_date.strftime('%b %d')}"
            if week_start_idx + i in sunday_equivalent_days: day_str += " (*)"
            rep_sheet.write(row_c, i + 1, day_str, header_format)
        row_c += 1

        for s in ['REP_DAY', 'REP_NIGHT']:
            rep_sheet.write(row_c, 0, get_indiv_name(s), shift_format)
            for i in range(7):
                day_idx = week_start_idx + i
                if day_idx >= num_days: continue
                assigned_doc = ""
                for d in doctors:
                    if s in doc_schedule[d][day_idx]: assigned_doc = d
                
                if assigned_doc: rep_sheet.write(row_c, i + 1, assigned_doc, doc_formats[assigned_doc])
                else:
                    if s == 'REP_DAY' and day_idx not in sunday_equivalent_days: rep_sheet.write(row_c, i + 1, "-", gray_dash_format)
                    else: rep_sheet.write(row_c, i + 1, "", empty_border)
            row_c += 1
        row_c += 2

    workbook.close()
    return output.getvalue()

# ==========================================
# 5. THE GRAPHICAL USER INTERFACE (GUI)
# ==========================================
st.set_page_config(page_title="Cardiology Scheduler", page_icon="🩺", layout="wide")

if "app_init_v8" not in st.session_state:
    settings = load_persistent_settings()
    st.session_state.clinics_base_df = pd.DataFrame({"Clinic Name": settings["clinics"]})
    st.session_state.docs_base_df = pd.DataFrame(settings["docs_base"]) 
    st.session_state.cap_base_df = pd.DataFrame(settings["capabilities"]) 
    st.session_state.recurrence_rules = settings.get("recurrence", {})
    st.session_state.absences_memory = {}
    st.session_state.active_month_key = ""
    st.session_state.app_init_v8 = True

current_clinics = [str(c).strip().upper().replace(" ", "_") for c in st.session_state.clinics_base_df["Clinic Name"].dropna().unique() if str(c).strip()]

st.title("🩺 Cardiology Shift Scheduler")
st.markdown("Automated constraints-based roster generation.")

with st.sidebar:
    st.header("1. Time Period")
    selected_year = st.number_input("Year", min_value=2024, max_value=2050, value=datetime.date.today().year)
    selected_month = st.number_input("Month", min_value=1, max_value=12, value=datetime.date.today().month)
    
    month_key = f"{selected_year}-{selected_month:02d}"
    if st.session_state.active_month_key != month_key:
        st.session_state.active_month_key = month_key
        # Wipe session grids to force rebuild for new month
        keys_to_clear = [k for k in st.session_state.keys() if k.startswith("src_") or k.startswith("editor_") or k.startswith("df_")]
        for k in keys_to_clear: del st.session_state[k]
    
    st.markdown("---")
    st.header("👨‍⚕️ Persistent Doctor Settings")
    
    doc_col_config = {
        "Doctor": st.column_config.TextColumn("Doctor Name", required=True),
        "Private Practice (Afternoon)": st.column_config.SelectboxColumn(
            "Private Practice", options=["", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        ),
        "Color": st.column_config.SelectboxColumn("Color", options=list(COLOR_PALETTE.keys()))
    }
    
    edited_docs_df = st.data_editor(
        st.session_state.docs_base_df,
        num_rows="dynamic",
        column_config=doc_col_config,
        hide_index=True, use_container_width=True,
        key="docs_editor_persistent"
    )
    
    current_doctors = [str(d).strip().upper() for d in edited_docs_df["Doctor"].dropna().unique() if str(d).strip()]
    
    st.markdown("---")
    st.header("👨‍⚕️ Doctor Capabilities")
    
    cap_cols = ["Doctor", "Ward Preferred", "OR Capable"] + current_clinics
    
    if list(st.session_state.cap_base_df.columns) != cap_cols or list(st.session_state.cap_base_df["Doctor"]) != current_doctors:
        new_df = pd.DataFrame({"Doctor": current_doctors})
        for c in cap_cols[1:]: new_df[c] = False
        
        last_edited_caps = st.session_state.get('cap_editor_persistent', st.session_state.cap_base_df)
        for idx, row in last_edited_caps.iterrows():
            d = row["Doctor"]
            if d in current_doctors:
                for c in cap_cols[1:]:
                    if c in last_edited_caps.columns:
                        new_df.loc[new_df["Doctor"] == d, c] = row[c]
        st.session_state.cap_base_df = new_df
        
    edited_cap_df = st.data_editor(st.session_state.cap_base_df, hide_index=True, use_container_width=True, key="cap_editor_persistent")
    save_persistent_settings(current_clinics, edited_docs_df, edited_cap_df, st.session_state.recurrence_rules)
    
    st.markdown("---")
    with st.expander("🗄️ Real-World Baseline & Reset"):
        st.markdown("Type the real-world historical stats your doctors have already accumulated. The algorithm uses this as a starting point for fairness.")
        ledger = load_historical_counters()
        baseline_stats = ledger.get("legacy_baseline", {})
        
        stat_cols = ['nights', 'reps', 'doubles', 'saturdays', 'sundays', 'holidays', 'super_holidays', 'golden_weekends']
        baseline_rows = []
        for doc in current_doctors:
            row = {"Doctor": doc}
            for c in stat_cols: row[c] = baseline_stats.get(doc, {}).get(c, 0.0)
            baseline_rows.append(row)
            
        baseline_df = pd.DataFrame(baseline_rows)
        edited_baseline = st.data_editor(baseline_df, hide_index=True, use_container_width=True, key="baseline_editor")
        
        for idx, row in edited_baseline.iterrows():
            doc = row["Doctor"]
            if doc not in baseline_stats: baseline_stats[doc] = {}
            for c in stat_cols: baseline_stats[doc][c] = float(row[c])
        ledger["legacy_baseline"] = baseline_stats
        with open(COUNTER_FILE, 'w') as f: json.dump(ledger, f, indent=4)
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ WIPE ALL COUNTERS (Factory Reset)", type="primary"):
            with open(COUNTER_FILE, 'w') as f: json.dump({"legacy_baseline": {}}, f, indent=4)
            st.rerun()

    st.markdown("---")
    st.header("🛠️ Troubleshooter")
    debug_toggle = st.checkbox("🚨 Enable Emergency Debug Mode", help="Ignores all equity rules to force a schedule.")

ui_roster_dates = get_roster_dates(selected_year, selected_month)
num_days_in_month = len(ui_roster_dates)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Absences & Desiderate", "🔒 Forced Shifts", "🏥 OR & Clinics", "⚙️ Generate Draft", "🚀 Edit & Publish"])

with tab1:
    st.subheader("Daily Absences & Desiderate")
    st.markdown("""
    💡 **Pro Tip:** Type 'F' in a box, select the cell, grab the little blue square in the bottom right, and drag it across the week to bulk-fill!
    - **F**: Ferie (Full day off, reduces monthly 34h target)
    - **X**: Desiderata Off (Full day off, does *not* reduce target)
    - **P**: Desiderata Mattina (No afternoon/night shifts)
    - **N**: No Notti (No night shifts)
    """)
    
    for w_idx, week_start_idx in enumerate(range(0, len(ui_roster_dates), 7)):
        st.markdown(f"#### Absences Week {w_idx + 1}")
        week_dates = ui_roster_dates[week_start_idx : week_start_idx + 7]
        col_dates = [d.strftime("%Y-%m-%d") for d in week_dates]
        col_displays = []
        for d in week_dates:
            day_str = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
            col_displays.append(day_str)
        
        src_key = f"src_abs_w{w_idx}_{month_key}"
        
        if src_key not in st.session_state:
            df_dict = {"Doctor": current_doctors}
            for i, d_str in enumerate(col_dates):
                df_dict[col_displays[i]] = [st.session_state.absences_memory.get(d_str, {}).get(doc, "") for doc in current_doctors]
            st.session_state[src_key] = pd.DataFrame(df_dict)
            
        if list(st.session_state[src_key]["Doctor"]) != current_doctors:
            df_dict = {"Doctor": current_doctors}
            for i, d_str in enumerate(col_dates):
                df_dict[col_displays[i]] = [st.session_state.absences_memory.get(d_str, {}).get(doc, "") for doc in current_doctors]
            st.session_state[src_key] = pd.DataFrame(df_dict)

        col_config = {"Doctor": st.column_config.TextColumn("Doctor Name", disabled=True, pinned=True)}
        for i, d_str in enumerate(col_dates):
            col_config[col_displays[i]] = st.column_config.SelectboxColumn(col_displays[i], options=["", "F", "X", "P", "N"], width="small")
            
        edited_df = st.data_editor(st.session_state[src_key], column_config=col_config, hide_index=True, use_container_width=True, key=f"editor_abs_w{w_idx}_{month_key}")
        
        # Save edits back to memory dictionary
        for idx, row in edited_df.iterrows():
            doc = row["Doctor"]
            for i, d_str in enumerate(col_dates):
                val = row[col_displays[i]]
                if d_str not in st.session_state.absences_memory: st.session_state.absences_memory[d_str] = {}
                st.session_state.absences_memory[d_str][doc] = val
        st.markdown("---")

with tab2:
    st.subheader("Visual Override Grid")
    st.markdown("Click on any empty cell to forcefully lock a specific doctor into that shift. Use the **🌟 SUPER HOLIDAY** row to manually flag a day (e.g. Patron Saint) as a holiday.")
    
    dynamic_shift_options = ['🌟 SUPER HOLIDAY', 'WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT'] + [f'OUT_{c}_AM' for c in current_clinics]

    edited_manual_grids = {}
    for w_idx, week_start_idx in enumerate(range(0, len(ui_roster_dates), 7)):
        st.markdown(f"#### Forced Week {w_idx + 1}")
        week_dates = ui_roster_dates[week_start_idx : week_start_idx + 7]
        col_names = []
        for d in week_dates:
            day_str = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
            col_names.append(day_str)
        
        src_key = f"src_manual_w{w_idx}_{month_key}"
        
        if src_key not in st.session_state:
            init_dict = {"Shift": dynamic_shift_options}
            for c in col_names: init_dict[c] = [""] * len(dynamic_shift_options)
            st.session_state[src_key] = pd.DataFrame(init_dict)
            
        if list(st.session_state[src_key]["Shift"]) != dynamic_shift_options or list(st.session_state[src_key].columns)[1:] != col_names:
            init_dict = {"Shift": dynamic_shift_options}
            for c in col_names: init_dict[c] = [""] * len(dynamic_shift_options)
            st.session_state[src_key] = pd.DataFrame(init_dict)

        col_config_manual = {"Shift": st.column_config.TextColumn("Shift", disabled=True, pinned=True)}
        for i, d in enumerate(week_dates):
            display_name = col_names[i]
            col_config_manual[display_name] = st.column_config.SelectboxColumn(display_name, options=["", "YES"] + current_doctors)
            
        edited_grid = st.data_editor(
            st.session_state[src_key], 
            column_config=col_config_manual, 
            hide_index=True, 
            use_container_width=True,
            key=f"editor_manual_w{w_idx}_{month_key}"
        )
        edited_manual_grids[w_idx] = edited_grid
        st.markdown("---")

with tab3:
    st.subheader("🏥 Operating Room & Outpatient Clinics Setup")
    
    edited_clinics_df = st.data_editor(
        st.session_state.clinics_base_df, 
        num_rows="dynamic",
        column_config={"Clinic Name": st.column_config.TextColumn("Define New Clinics Here", required=True)},
        use_container_width=True,
        hide_index=True,
        key="clinics_editor_persistent_t3"
    )
    st.session_state.clinics_base_df = edited_clinics_df
    current_clinics = [str(c).strip().upper().replace(" ", "_") for c in edited_clinics_df["Clinic Name"].dropna().unique() if str(c).strip()]
    
    with st.expander("⚙️ Define Usual Recurring Schedule (Global)", expanded=False):
        st.markdown("Set up the standard cadence. These rules save permanently across all months.")
        activities = ["OR"] + current_clinics
        for act in activities:
            if act not in st.session_state.recurrence_rules:
                st.session_state.recurrence_rules[act] = {"weekdays": [], "weeks": ["All"]}
                
            cols_r = st.columns(2)
            with cols_r[0]:
                st.session_state.recurrence_rules[act]["weekdays"] = st.multiselect(
                    f"{act} Days", 
                    ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
                    default=st.session_state.recurrence_rules[act].get("weekdays", []),
                    key=f"rec_wd_{act}"
                )
            with cols_r[1]:
                st.session_state.recurrence_rules[act]["weeks"] = st.multiselect(
                    f"{act} Weeks of Month", 
                    ["All", "1st", "2nd", "3rd", "4th", "5th"],
                    default=st.session_state.recurrence_rules[act].get("weeks", ["All"]),
                    key=f"rec_wk_{act}"
                )
        save_persistent_settings(current_clinics, edited_docs_df, edited_cap_df, st.session_state.recurrence_rules)

    st.markdown("#### 📅 Schedule Specific Dates for This Month")
    st.markdown("Tick the exact dates where an OR or Clinic session is active.")
    
    activities = ["OR"] + current_clinics
    day_cols_act = [d.strftime('%b %d') + (" Sa" if d.weekday()==5 else " Su" if d.weekday()==6 else "") for d in ui_roster_dates]
    
    src_key = f"src_activities_{month_key}"
    
    if st.button("🔄 Apply Recurring Rules to Grid"):
        init_df = pd.DataFrame({"Activity": activities})
        for c in day_cols_act: init_df[c] = False
        for act in activities:
            rules = st.session_state.recurrence_rules.get(act, {})
            valid_dates = calculate_recurring_days(selected_year, selected_month, rules.get("weekdays", []), rules.get("weeks", ["All"]))
            for vd in valid_dates:
                d_obj = datetime.datetime.strptime(vd, "%Y-%m-%d").date()
                if d_obj in ui_roster_dates:
                    col_idx = ui_roster_dates.index(d_obj)
                    col_name = day_cols_act[col_idx]
                    init_df.loc[init_df["Activity"] == act, col_name] = True
        st.session_state[src_key] = init_df
        st.rerun()

    if src_key not in st.session_state:
        init_df = pd.DataFrame({"Activity": activities})
        for c in day_cols_act: init_df[c] = False
        for act in activities:
            rules = st.session_state.recurrence_rules.get(act, {})
            valid_dates = calculate_recurring_days(selected_year, selected_month, rules.get("weekdays", []), rules.get("weeks", ["All"]))
            for vd in valid_dates:
                d_obj = datetime.datetime.strptime(vd, "%Y-%m-%d").date()
                if d_obj in ui_roster_dates:
                    col_idx = ui_roster_dates.index(d_obj)
                    col_name = day_cols_act[col_idx]
                    init_df.loc[init_df["Activity"] == act, col_name] = True
        st.session_state[src_key] = init_df

    if list(st.session_state[src_key]["Activity"]) != activities or list(st.session_state[src_key].columns)[1:] != day_cols_act:
        new_df = pd.DataFrame({"Activity": activities})
        for c in day_cols_act: new_df[c] = False
        st.session_state[src_key] = new_df
        
    act_col_config = {"Activity": st.column_config.TextColumn("Activity", disabled=True, pinned=True)}
    for d_str in day_cols_act:
        act_col_config[d_str] = st.column_config.CheckboxColumn(d_str, default=False, width="small")

    edited_act_df = st.data_editor(
        st.session_state[src_key],
        column_config=act_col_config,
        hide_index=True, use_container_width=True,
        key=f"editor_activities_{month_key}"
    )
    
    or_days_formatted = []
    outpatient_setup = {c: {'days': [], 'capable': []} for c in current_clinics}
    
    for idx, row in edited_act_df.iterrows():
        act = row["Activity"]
        for i, d in enumerate(ui_roster_dates):
            d_str = d.strftime("%Y-%m-%d")
            col_name = day_cols_act[i]
            if row[col_name]:
                if act == "OR": or_days_formatted.append(d_str)
                else: outpatient_setup[act]['days'].append(d_str)
                    
    for c in current_clinics:
        if c in edited_cap_df.columns:
            outpatient_setup[c]['capable'] = edited_cap_df[edited_cap_df[c] == True]["Doctor"].dropna().tolist()

with tab4:
    st.subheader("⚙️ Generate Draft Schedule")
    st.markdown("The algorithm will mathematically balance hours, nights, and holidays, outputting an editable draft in **Tab 5**.")
    
    if "generation_error" in st.session_state:
        st.error(st.session_state.generation_error)
        with st.expander("🛠️ Resolution Center (Lesser Evils)", expanded=True):
            st.markdown("The department is mathematically understaffed based on your current settings. Select an override to continue:")
            allow_understaffing = st.checkbox("☑️ Allow Shift Understaffing (Leave some non-critical Clinic/Ward PM shifts empty)")
            ignore_34h = st.checkbox("☑️ Ignore 34h Legal Minimum (Allow some doctors to run an hour deficit)")
            st.session_state.resolution_toggles = {"allow_understaffing": allow_understaffing, "ignore_34h": ignore_34h}
    
    if st.button("🛠️ GENERATE DRAFT", use_container_width=True):
        if len(current_doctors) == 0:
            st.error("❌ Please add at least one doctor to the Persistent Settings.")
        else:
            if "resolution_toggles" not in st.session_state: st.session_state.resolution_toggles = {}
            with st.spinner("Calculating constraints (this may cascade through multiple fallback levels and take up to 60 seconds)..."):
                pp_dict = {}
                doc_capabilities_dict = {}
                for idx, row in edited_docs_df.iterrows():
                    doc = str(row["Doctor"]).strip().upper()
                    pp = row.get("Private Practice (Afternoon)", "")
                    if pd.notna(pp) and pp: pp_dict[doc] = pp
                for idx, row in edited_cap_df.iterrows():
                    doc = row["Doctor"]
                    doc_capabilities_dict[doc] = {col: bool(row[col]) for col in cap_cols if col != "Doctor"}

                daily_absences = {}
                for d_str, doc_dict in st.session_state.absences_memory.items():
                    for doc, val in doc_dict.items():
                        if val in ["F", "X", "P", "N"]:
                            daily_absences[(doc, d_str)] = val

                manual_shifts = []
                manual_super_holidays = []
                for w_idx, df_w in edited_manual_grids.items():
                    for idx, row in df_w.iterrows():
                        shift_val = str(row["Shift"]).strip().upper()
                        week_dates = ui_roster_dates[w_idx*7 : (w_idx*7)+7]
                        for i, d in enumerate(week_dates):
                            d_str = d.strftime("%Y-%m-%d")
                            display_name = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
                            if display_name in row:
                                val = str(row.get(display_name, "")).strip().upper()
                                if shift_val == '🌟 SUPER HOLIDAY':
                                    if val == 'YES': manual_super_holidays.append(d_str)
                                elif val and val in current_doctors:
                                    manual_shifts.append((val, d_str, shift_val))

                success, draft_grids, warning = generate_draft_schedule(
                    year=selected_year, month=selected_month, conditional_or_days=or_days_formatted, 
                    manual_festivities=[], manual_super_holidays=manual_super_holidays, manual_assignments=manual_shifts,
                    doctor_capabilities=doc_capabilities_dict, outpatient_configs=outpatient_setup, 
                    daily_absences=daily_absences, private_practice_afternoons=pp_dict, doctors_list=current_doctors, 
                    debug_mode=debug_toggle, resolution_toggles=st.session_state.get("resolution_toggles", {})
                )
                
                if success:
                    st.success("✅ Draft generated successfully! Go to **Tab 5 (🚀 Edit & Publish)** to review, modify, and publish the final Excel.")
                    if warning: st.warning(warning)
                    st.session_state.generated_draft_grids = draft_grids
                    st.session_state.draft_month_key = month_key
                    if "generation_error" in st.session_state: del st.session_state["generation_error"]
                    for k in list(st.session_state.keys()):
                        if k.startswith("src_draft_w"): del st.session_state[k]
                else:
                    st.session_state.generation_error = warning
                    st.rerun()

with tab5:
    st.subheader("🚀 Review, Edit & Publish")
    
    if "generated_draft_grids" not in st.session_state:
        st.info("No draft generated yet. Go to **Tab 4** to generate the base schedule first.")
    elif st.session_state.get("draft_month_key") != month_key:
        st.warning(f"⚠️ You are currently viewing a different month, but the saved draft is for an old layout. Please go to **Tab 4** and click 'GENERATE DRAFT' again.")
    else:
        st.markdown("Make any last-minute human adjustments here. The Live Fairness Dashboard below will update instantly as you edit.")
        
        edited_final_drafts = {}
        for w_idx, draft_df in st.session_state.generated_draft_grids.items():
                
            st.markdown(f"#### Draft Week {w_idx + 1}")
            week_dates = ui_roster_dates[w_idx*7 : (w_idx*7)+7]
            col_config = {"Shift": st.column_config.TextColumn("Shift", disabled=True)}
            for i, d in enumerate(week_dates):
                d_str = d.strftime("%Y-%m-%d")
                display_name = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
                col_config[d_str] = st.column_config.SelectboxColumn(display_name, options=["", "[UNCOVERED]"] + current_doctors)
                
            src_key = f"src_draft_w{w_idx}_{month_key}"
            if src_key not in st.session_state:
                st.session_state[src_key] = draft_df.copy()
                
            edited_final_drafts[w_idx] = st.data_editor(
                st.session_state[src_key], 
                column_config=col_config, 
                hide_index=True, 
                use_container_width=True,
                key=f"editor_draft_w{w_idx}_{month_key}"
            )
            st.markdown("---")

        st.markdown("### 📊 Live Fairness Dashboard")
        
        date_to_idx = {d.strftime("%Y-%m-%d"): idx for idx, d in enumerate(ui_roster_dates)}
        it_holidays = holidays.IT(years=[selected_year-1, selected_year, selected_year+1])
        
        manual_super_holidays = []
        for w_idx, df_w in edited_manual_grids.items():
            for idx, row in df_w.iterrows():
                shift_val = str(row["Shift"]).strip().upper()
                week_dates = ui_roster_dates[w_idx*7 : (w_idx*7)+7]
                for i, d in enumerate(week_dates):
                    d_str = d.strftime("%Y-%m-%d")
                    display_name = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
                    if display_name in row:
                        val = str(row.get(display_name, "")).strip().upper()
                        if shift_val == '🌟 SUPER HOLIDAY' and val == 'YES':
                            manual_super_holidays.append(d_str)

        daily_absences = {}
        for d_str, doc_dict in st.session_state.absences_memory.items():
            for doc, val in doc_dict.items():
                if val in ["F", "X", "P", "N"]:
                    daily_absences[(doc, d_str)] = val

        doc_contract_off_set = {d: set() for d in current_doctors}
        for (d, date_str), val in daily_absences.items():
            if val == 'F' and date_str in date_to_idx:
                doc_contract_off_set[d].add(date_str)
                
        doc_contract_days = {d: max(1, num_days_in_month - len(doc_contract_off_set[d])) for d in current_doctors}

        doc_schedule = {doc: {day_idx: [] for day_idx in range(num_days_in_month)} for doc in current_doctors}
        for w_idx, df_w in edited_final_drafts.items():
            week_dates = ui_roster_dates[w_idx*7 : (w_idx*7)+7]
            for idx, row in df_w.iterrows():
                s = row["Shift"]
                for i, d_date in enumerate(week_dates):
                    day_idx = (w_idx * 7) + i
                    if day_idx >= num_days_in_month: continue
                    d_str = d_date.strftime("%Y-%m-%d")
                    assigned_doc = str(row.get(d_str, "")).strip().upper()
                    if assigned_doc and assigned_doc in current_doctors:
                        doc_schedule[assigned_doc][day_idx].append(s)

        shifts = ['WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']
        for out_type in outpatient_setup.keys(): shifts.append(f'OUT_{out_type}_AM')
        day_active = [s for s in shifts if s not in ['NIGHT', 'REP_NIGHT', 'REP_DAY']]

        monthly_stats = compute_monthly_stats(doc_schedule, num_days_in_month, ui_roster_dates, day_active, it_holidays, [], manual_super_holidays, current_doctors)
        
        dashboard_data = []
        for doc in current_doctors:
            if doc == 'GAUDENZI':
                doc_target_display = "MANUAL"
                actual_hours = monthly_stats[doc]['total_hours']
                difference_display = "N/A"
            else:
                doc_target = int((doc_contract_days[doc] / 7) * 34)
                doc_target_display = doc_target
                actual_hours = monthly_stats[doc]['total_hours']
                difference = actual_hours - doc_target
                difference_display = f"+{difference}" if difference > 0 else str(difference)
            
            dashboard_data.append({
                "Doctor": doc,
                "Target Hours": doc_target_display,
                "Actual Active Hours": actual_hours,
                "Difference (+/-)": difference_display,
                "Nights": monthly_stats[doc]['nights'],
                "Reps": monthly_stats[doc]['reps'],
                "Doubles": monthly_stats[doc]['doubles'],
                "Saturdays": f"{monthly_stats[doc]['saturdays']:.1f}",
                "Sundays": f"{monthly_stats[doc]['sundays']:.1f}",
                "Super Holidays": f"{monthly_stats[doc]['super_holidays']:.1f}",
            })
            
        st.dataframe(pd.DataFrame(dashboard_data), use_container_width=True)

        if st.button("🚀 APPROVE & PUBLISH (Save Month Stats)", use_container_width=True, type="primary"):
            with st.spinner("Compiling Excel and updating ledgers..."):
                doc_colors_ui = {}
                for idx, row in edited_docs_df.iterrows():
                    doc = str(row["Doctor"]).strip().upper()
                    col = row.get("Color", "White")
                    doc_colors_ui[doc] = COLOR_PALETTE.get(col, "#FFFFFF")

                excel_data = process_and_export_schedule(
                    edited_weekly_grids=edited_final_drafts, year=selected_year, month=selected_month, 
                    conditional_or_days=or_days_formatted, manual_festivities=[], manual_super_holidays=manual_super_holidays,
                    outpatient_configs=outpatient_setup, daily_absences=daily_absences,
                    doctor_colors=doc_colors_ui, doctors_list=current_doctors
                )
                
                st.success("✅ Schedule finalized! The Monthly Ledger has been updated.")
                timestamp = datetime.datetime.now().strftime("%H%M%S")
                dl_filename = f'OFFICIAL_cardiology_schedule_{selected_year}_{selected_month}_{timestamp}.xlsx'
                
                st.download_button(
                    label="📥 Download Official Excel Schedule",
                    data=excel_data,
                    file_name=dl_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
