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
    cal = calendar.monthcalendar(year, month)
    day_map = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}
    target_wds = [day_map[w] for w in weekdays if w in day_map]
    
    valid_dates = []
    for wd in target_wds:
        occurrences = [week[wd] for week in cal if week[wd] != 0]
        for i, day in enumerate(occurrences):
            week_idx = i + 1
            is_valid = False
            if "All" in weeks_list: is_valid = True
            elif "1st" in weeks_list and week_idx == 1: is_valid = True
            elif "2nd" in weeks_list and week_idx == 2: is_valid = True
            elif "3rd" in weeks_list and week_idx == 3: is_valid = True
            elif "4th" in weeks_list and week_idx == 4: is_valid = True
            elif "5th" in weeks_list and week_idx == 5: is_valid = True
            
            if is_valid: valid_dates.append(datetime.date(year, month, day).strftime("%Y-%m-%d"))
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
                            doctor_capabilities, outpatient_configs, daily_absences, private_practice_afternoons, doctors_list, debug_mode=False):
    
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
    for day_idx in range(num_days):
        current_date_str = roster_dates[day_idx].strftime("%Y-%m-%d")
        if day_idx in sunday_equivalent_days: total_available_hours += 24
        else:
            total_available_hours += 36
            if current_date_str in conditional_or_days: total_available_hours += 6
            for out_type, config in outpatient_configs.items():
                if current_date_str in config['days']: total_available_hours += 6

    gaudenzi_manual_active_hrs = sum(12 if s == 'NIGHT' else (6 if 'REP' not in s else 0) for (d, day, s) in manual_keys if d == 'GAUDENZI')
    total_available_hours -= gaudenzi_manual_active_hrs

    # 🟢 THE CASCADING ENGINE
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
            
            if day_idx not in sunday_equivalent_days: 
                model.AddExactlyOne(work[(d, day_idx, 'WARD_AM')] for d in doctors)
                model.AddExactlyOne(work[(d, day_idx, 'URG_AM')] for d in doctors)
                model.AddExactlyOne(work[(d, day_idx, 'WARD_PM')] for d in doctors)
                model.AddExactlyOne(work[(d, day_idx, 'URG_PM')] for d in doctors)
                for d in doctors: model.Add(work[(d, day_idx, 'REP_DAY')] == 0)
                
                if current_date_str in conditional_or_days:
                    model.AddExactlyOne(work[(d, day_idx, 'OR_AM')] for d in doctors)
                    for d in doctors:
                        if not doctor_capabilities.get(d, {}).get("OR Capable", False): model.Add(work[(d, day_idx, 'OR_AM')] == 0)
                else:
                    for d in doctors: model.Add(work[(d, day_idx, 'OR_AM')] == 0)
                    
                for out_type, config in outpatient_configs.items():
                    s_name = f'OUT_{out_type}_AM'
                    if current_date_str in config['days']:
                        model.AddExactlyOne(work[(d, day_idx, s_name)] for d in doctors)
                        for d in doctors:
                            if d not in config['capable']: model.Add(work[(d, day_idx, s_name)] == 0)
                    else:
                        for d in doctors: model.Add(work[(d, day_idx, s_name)] == 0)
            else: 
                model.AddExactlyOne(work[(d, day_idx, 'WARD_AM')] for d in doctors) 
                model.AddExactlyOne(work[(d, day_idx, 'WARD_PM')] for d in doctors) 
                model.AddExactlyOne(work[(d, day_idx, 'REP_DAY')] for d in doctors)
                for d in doctors:
                    for s in shifts:
                        if s not in ['WARD_AM', 'WARD_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']: model.Add(work[(d, day_idx, s)] == 0)

        day_name_to_num = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}

        for d in doctors:
            if d == 'GAUDENZI': continue
            
            pp_day_num = day_name_to_num.get(private_practice_afternoons.get(d, "").strip().capitalize(), -1)
            for day_idx in range(num_days):
                weekday = roster_dates[day_idx].weekday()
                
                if weekday == pp_day_num:
                    for s in shifts:
                        if s.endswith('_PM') or s == 'REP_DAY': model.Add(work[(d, day_idx, s)] == 0)
                        
                    # Incentivize Night & Rep_Night on PP day
                    objective_terms.append(500 * work[(d, day_idx, 'NIGHT')])
                    objective_terms.append(300 * work[(d, day_idx, 'REP_NIGHT')])
                        
                next_day_date = roster_dates[day_idx] + datetime.timedelta(days=1)
                if next_day_date.weekday() == pp_day_num:
                    # Banned the night before PP
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
                    objective_terms.append(1000 * ideal_pre_night)
                else:
                    if pass_level < 3:
                        if day_idx not in sunday_equivalent_days: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 1)
                        else: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 2)
                
                other_than_night = [s for s in shifts if s != 'NIGHT']
                model.Add(sum(work[(d, day_idx, s)] for s in other_than_night) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
                if day_idx + 1 < num_days: model.Add(sum(work[(d, day_idx + 1, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
                if day_idx + 2 < num_days: model.Add(sum(work[(d, day_idx + 2, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])

            doctor_gws = []
            for week_start_idx in range(0, num_days, 7):
                fri_idx, sat_idx, sun_idx = week_start_idx + 4, week_start_idx + 5, week_start_idx + 6
                gw_var = model.NewBoolVar(f'gw_{d}_{week_start_idx}')
                ruining_shifts = [work[(d, fri_idx, 'NIGHT')], work[(d, fri_idx, 'REP_NIGHT')]]
                for s in shifts: ruining_shifts.extend([work[(d, sat_idx, s)], work[(d, sun_idx, s)]])
                model.Add(sum(ruining_shifts) == 0).OnlyEnforceIf(gw_var)
                model.Add(sum(ruining_shifts) > 0).OnlyEnforceIf(gw_var.Not())
                doctor_gws.append(gw_var)
                objective_terms.append(-5 * int(lifetime[d]['golden_weekends']) * gw_var)
                
            if pass_level < 2: model.Add(sum(doctor_gws) >= 1)

        # 🟢 PRE-VACATION NIGHT (Absolute Priority)
        for d in doctors:
            if d == 'GAUDENZI': continue
            for week_start_idx in range(0, num_days, 7):
                week_days = range(week_start_idx, min(week_start_idx + 7, num_days))
                if len(week_days) == 7:
                    is_off_week = all(daily_absences.get((d, roster_dates[day_idx].strftime("%Y-%m-%d")), "") in ['F', 'X'] for day_idx in week_days)
                    if is_off_week:
                        if 0 <= week_start_idx - 1 < num_days:
                            for s in shifts: model.Add(work[(d, week_start_idx - 1, s)] == 0)
                        if 0 <= week_start_idx - 4 < num_days:
                            # 5000 point absolute mathematical priority
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
            
            if pass_level == 0:
                model.Add(dm <= 6)
            elif pass_level == 1:
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
                    objective_terms.append(150 * am_pm)  # BOOSTED
                    pm_am = model.NewBoolVar('')
                    model.AddBoolAnd([work[(d, day_idx, 'WARD_PM')], work[(d, day_idx+1, 'WARD_AM')]]).OnlyEnforceIf(pm_am)
                    objective_terms.append(150 * pm_am)  # BOOSTED
                    
                for w_idx in range(len(weeks)):
                    ward_active = model.NewBoolVar('')
                    model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(ward_active)
                    model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(ward_active.Not())
                    objective_terms.append(-400 * ward_active) # BOOSTED
                    
                    if w_idx < len(weeks) - 1:
                        ward_next_active = model.NewBoolVar('')
                        model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(ward_next_active)
                        model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(ward_next_active.Not())
                        consecutive_ward = model.NewBoolVar('')
                        model.AddBoolAnd([ward_active, ward_next_active]).OnlyEnforceIf(consecutive_ward)
                        
                        if pass_level == 0:
                            model.Add(consecutive_ward == 0)
                        else:
                            objective_terms.append(-2000 * consecutive_ward) # BOOSTED

        for day_idx in sunday_equivalent_days:
            for d in doctors:
                if d == 'GAUDENZI': continue
                ward_double = model.NewBoolVar('')
                model.AddBoolAnd([work[(d, day_idx, 'WARD_AM')], work[(d, day_idx, 'WARD_PM')]]).OnlyEnforceIf(ward_double)
                objective_terms.append(10 * ward_double) 
                urg_double = model.NewBoolVar('')
                model.AddBoolAnd([work[(d, day_idx, 'URG_AM')], work[(d, day_idx, 'URG_PM')]]).OnlyEnforceIf(urg_double)
                objective_terms.append(10 * urg_double) 

        # 🟢 LIFETIME EQUITY GRAVITY
        for d in doctors:
            if d == 'GAUDENZI': continue
            n_pts = int(lifetime[d]['nights'] * 10)
            r_pts = int(lifetime[d]['reps'] * 10)
            sat_pts = int(lifetime[d]['saturdays'] * 10) 
            sun_pts = int(lifetime[d]['sundays'] * 10)
            sh_pts = int(lifetime[d]['super_holidays'] * 10)
            
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
                df_dict = {"Shift": shifts}
                for c in col_names: df_dict[c] = [""] * len(shifts)
                df = pd.DataFrame(df_dict)
                
                for i, d_date in enumerate(week_dates):
                    day_idx = week_start_idx + i
                    d_str = d_date.strftime("%Y-%m-%d")
                    if day_idx >= num_days: continue
                    for s in shifts:
                        for d in doctors:
                            if solver.Value(work[(d, day_idx, s)]) == 1:
                                df.loc[df["Shift"] == s, d_str] = d
            draft_grids[w_idx] = df
            
            warning_msg = ""
            if pass_level == 1: warning_msg = "⚠️ GEAR 2: Hard consecutive ward limits were relaxed to prevent an hour deficit."
            elif pass_level == 2: warning_msg = "⚠️ GEAR 3: Monthly hour constraints were loosened completely. Equity is handled via Min-Max fairness."
            elif pass_level == 3: warning_msg = "⚠️ SURVIVAL GEAR: Maximum shift caps removed. Verify fairness manually."
            elif pass_level == 4: warning_msg = "🚨 DEBUG MODE: All equity limits, max deficits, and 34h floors disabled."
            
            if overlap_warning: warning_msg += f"\n\n{overlap_warning}"
            return True, draft_grids, warning_msg

    return False, None, "🛑 Constraints are too tight. Even after shifting to Survival Gear, the algorithm cannot find a legally compliant schedule. Use Emergency Debug Mode."

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
    lifetime = {d: get_lifetime_stats(ledger, d) for d in doctors}

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
                    for d in doctors:
                        if s in doc_schedule[d][day_idx]: assigned_doc = d
                    
                    if assigned_doc:
                        worksheet.write(row_cursor, i + 1, assigned_doc, doc_formats[assigned_doc])
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
