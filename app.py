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
                    objective_terms.append(1000 * ideal_pre_night)
                else:
                    if pass_level < 3:
                        if day_idx not in sunday_equivalent_days: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 1)
                        else: model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 2)
                
                other_than_night = [s for s in shifts if s != 'NIGHT']
                model.Add(sum(work[(d, day_idx, s)] for s in other_than_night) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
                if day_idx + 1 < num_days: model.Add(sum(work[(d, day_idx + 1, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
                if day_idx + 2 < num_days: model.Add(sum(work[(d, day_idx + 2, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])

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
            # Pass 2 and 3 have NO hard hour boundaries; min-max fairness takes over entirely

            objective_terms.append(-100 * dp)
            objective_terms.append(-500 * dm) # Stronger penalty for deficits to encourage hitting 34h
            
        max_dm = model.NewIntVar(0, 1000, 'max_dm')
        max_dp = model.NewIntVar(0, 1000, 'max_dp')
        
        for d in doctors:
            if d == 'GAUDENZI': continue
            model.Add(diff_m_vars[d] <= max_dm)
            model.Add(diff_p_vars[d] <= max_dp)
            
        # Squeeze the extremes to equalize everyone
        objective_terms.append(-2000 * max_dm)
        objective_terms.append(-1000 * max_dp)

        # 🟢 WARD CONTINUITY (Drops off in Pass 2)
        for d in doctors:
            if d == 'GAUDENZI': continue
            if doctor_capabilities.get(d, {}).get("Ward Preferred", False):
                for day_idx in range(num_days):
                    objective_terms.extend([20 * work[(d, day_idx, 'WARD_AM')], 20 * work[(d, day_idx, 'WARD_PM')]])

            for day_idx in range(num_days - 1):
                am_pm = model.NewBoolVar('')
                model.AddBoolAnd([work[(d, day_idx, 'WARD_AM')], work[(d, day_idx+1, 'WARD_PM')]]).OnlyEnforceIf(am_pm)
                objective_terms.append(30 * am_pm) 
                pm_am = model.NewBoolVar('')
                model.AddBoolAnd([work[(d, day_idx, 'WARD_PM')], work[(d, day_idx+1, 'WARD_AM')]]).OnlyEnforceIf(pm_am)
                objective_terms.append(30 * pm_am)
                
            for w_idx in range(len(weeks)):
                ward_active = model.NewBoolVar('')
                model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(ward_active)
                model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(ward_active.Not())
                objective_terms.append(-200 * ward_active)
                
                if w_idx < len(weeks) - 1:
                    ward_next_active = model.NewBoolVar('')
                    model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(ward_next_active)
                    model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(ward_next_active.Not())
                    consecutive_ward = model.NewBoolVar('')
                    model.AddBoolAnd([ward_active, ward_next_active]).OnlyEnforceIf(consecutive_ward)
                    
                    if pass_level == 0:
                        model.Add(consecutive_ward == 0) # Hard rule
                    else:
                        objective_terms.append(-800 * consecutive_ward) # Soft rule

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
st.set_page_config(page_title="Cardiology Scheduler", page_icon="🩺", layout="wide")

if "app_initialized_v5" not in st.session_state:
    settings = load_persistent_settings()
    st.session_state.clinics_base_df = pd.DataFrame({"Clinic Name": settings["clinics"]})
    st.session_state.docs_base_df = pd.DataFrame(settings["docs_base"]) 
    st.session_state.cap_base_df = pd.DataFrame(settings["capabilities"]) 
    st.session_state.recurrence_rules = settings.get("recurrence", {})
    st.session_state.app_initialized_v5 = True

current_clinics = [str(c).strip().upper().replace(" ", "_") for c in st.session_state.clinics_base_df["Clinic Name"].dropna().unique() if str(c).strip()]

st.title("🩺 Cardiology Shift Scheduler")
st.markdown("Automated constraints-based roster generation.")

with st.sidebar:
    st.header("1. Time Period")
    selected_year = st.number_input("Year", min_value=2024, max_value=2050, value=datetime.date.today().year)
    selected_month = st.number_input("Month", min_value=1, max_value=12, value=datetime.date.today().month)
    
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
    st.header("🛠️ Troubleshooter")
    debug_toggle = st.checkbox("🚨 Enable Emergency Debug Mode", help="Ignores all equity rules to force a schedule.")

month_key = f"{selected_year}-{selected_month:02d}"
ui_roster_dates = get_roster_dates(selected_year, selected_month)
num_days_in_month = calendar.monthrange(selected_year, selected_month)[1]
month_dates = [datetime.date(selected_year, selected_month, d) for d in range(1, num_days_in_month + 1)]

day_cols = []
for d in month_dates:
    day_str = str(d.day)
    if d.weekday() == 5: day_str += " Sa"
    elif d.weekday() == 6: day_str += " Su"
    day_cols.append(day_str)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Absences & Desiderate", "🔒 Forced Shifts", "🏥 OR & Clinics", "⚙️ Generate Draft", "🚀 Edit & Publish"])

with tab1:
    st.subheader("Daily Absences & Desiderate")
    st.markdown("""
    Select specific flags for each doctor per day:
    - **F**: Ferie (Full day off, reduces monthly 34h target)
    - **X**: Desiderata Off (Full day off, does *not* reduce target)
    - **P**: Desiderata Mattina (No afternoon/night shifts)
    - **N**: No Notti (No night shifts)
    """)
    
    abs_grid_key = f"absences_base_{month_key}"
    
    col_config = {"Doctor": st.column_config.TextColumn("Doctor Name", disabled=True, pinned=True)}
    for d_str in day_cols:
        col_config[d_str] = st.column_config.SelectboxColumn(d_str, options=["", "F", "X", "P", "N"], width="small")
        
    if abs_grid_key not in st.session_state:
        init_df = pd.DataFrame({"Doctor": current_doctors})
        for c in day_cols: init_df[c] = ""
        st.session_state[abs_grid_key] = init_df
        
    if list(st.session_state[abs_grid_key]["Doctor"]) != current_doctors:
        new_df = pd.DataFrame({"Doctor": current_doctors})
        for c in day_cols: new_df[c] = ""
        
        last_edited_abs = st.session_state.get(f"editor_{abs_grid_key}", st.session_state[abs_grid_key])
        for idx, row in last_edited_abs.iterrows():
            if row["Doctor"] in current_doctors:
                for c in day_cols:
                    if c in last_edited_abs.columns: new_df.loc[new_df["Doctor"] == row["Doctor"], c] = row[c]
        st.session_state[abs_grid_key] = new_df
        
    edited_absences_df = st.data_editor(
        st.session_state[abs_grid_key],
        column_config=col_config,
        hide_index=True, use_container_width=True,
        key=f"editor_{abs_grid_key}"
    )

with tab2:
    st.subheader("Visual Override Grid")
    st.markdown("Click on any empty cell to forcefully lock a specific doctor into that shift. Use the **🌟 SUPER HOLIDAY** row to manually flag a day (e.g. Patron Saint) as a holiday.")
    
    dynamic_shift_options = ['🌟 SUPER HOLIDAY', 'WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT'] + [f'OUT_{c}_AM' for c in current_clinics]

    if st.session_state.get("current_ym_manual") != f"{selected_year}-{selected_month}":
        st.session_state["current_ym_manual"] = f"{selected_year}-{selected_month}"
        for k in list(st.session_state.keys()):
            if k.startswith("manual_grid_base_") or k.startswith("last_edited_manual_grid_"):
                del st.session_state[k]

    edited_manual_grids = {}
    for w_idx, week_start_idx in enumerate(range(0, len(ui_roster_dates), 7)):
        st.markdown(f"#### Forced Week {w_idx + 1}")
        week_dates = ui_roster_dates[week_start_idx : week_start_idx + 7]
        col_names = [d.strftime("%Y-%m-%d") for d in week_dates]
        
        base_key = f"manual_grid_base_{w_idx}"
        last_edited_key = f"last_edited_manual_grid_{w_idx}"
        
        if base_key not in st.session_state:
            init_dict = {"Shift": dynamic_shift_options}
            for c in col_names: init_dict[c] = [""] * len(dynamic_shift_options)
            st.session_state[base_key] = pd.DataFrame(init_dict)
            
        current_grid_state = st.session_state.get(last_edited_key, st.session_state[base_key])
        
        if list(current_grid_state["Shift"]) != dynamic_shift_options:
            init_dict = {"Shift": dynamic_shift_options}
            for c in col_names: init_dict[c] = [""] * len(dynamic_shift_options)
            new_df = pd.DataFrame(init_dict)
            for idx, row in current_grid_state.iterrows():
                s = row["Shift"]
                if s in dynamic_shift_options:
                    for c in col_names:
                        if c in current_grid_state.columns:
                            new_df.loc[new_df["Shift"] == s, c] = row[c]
            st.session_state[base_key] = new_df

        col_config_manual = {"Shift": st.column_config.TextColumn("Shift", disabled=True, pinned=True)}
        for d in week_dates:
            d_str = d.strftime("%Y-%m-%d")
            display_name = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
            col_config_manual[d_str] = st.column_config.SelectboxColumn(display_name, options=["", "YES"] + current_doctors)
            
        edited_grid = st.data_editor(
            st.session_state[base_key], 
            column_config=col_config_manual, 
            hide_index=True, 
            use_container_width=True,
            key=f"editor_manual_grid_week_{w_idx}"
        )
        st.session_state[last_edited_key] = edited_grid
        edited_manual_grids[w_idx] = edited_grid
        st.markdown("---")

with tab3:
    st.subheader("🏥 Operating Room & Outpatient Clinics Setup")
    
    act_grid_key = f"activities_base_{month_key}"
    
    edited_clinics_df = st.data_editor(
        st.session_state.clinics_base_df, 
        num_rows="dynamic",
        column_config={"Clinic Name": st.column_config.TextColumn("Define New Clinics Here", required=True)},
        use_container_width=True,
        hide_index=True,
        key="clinics_editor_persistent"
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
    st.markdown("Tick the exact dates where an OR or Clinic session is active. It is pre-filled based on your Recurring Schedule.")
    
    activities = ["OR"] + current_clinics
    act_col_config = {"Activity": st.column_config.TextColumn("Activity", disabled=True, pinned=True)}
    for d_str in day_cols:
        act_col_config[d_str] = st.column_config.CheckboxColumn(d_str, default=False, width="small")
        
    if act_grid_key not in st.session_state:
        init_df = pd.DataFrame({"Activity": activities})
        for c in day_cols: init_df[c] = False
        
        # Pre-fill based on recurrence rules!
        for act in activities:
            rules = st.session_state.recurrence_rules.get(act, {})
            valid_dates = calculate_recurring_days(selected_year, selected_month, rules.get("weekdays", []), rules.get("weeks", ["All"]))
            for vd in valid_dates:
                d_obj = datetime.datetime.strptime(vd, "%Y-%m-%d")
                col_idx = d_obj.day - 1
                col_name = day_cols[col_idx]
                init_df.loc[init_df["Activity"] == act, col_name] = True
                
        st.session_state[act_grid_key] = init_df
        
    curr_act_grid = st.session_state[act_grid_key]
    if list(curr_act_grid["Activity"]) != activities:
        new_df = pd.DataFrame({"Activity": activities})
        for c in day_cols: new_df[c] = False
        for act in activities:
            if act in list(curr_act_grid["Activity"]):
                for c in day_cols:
                    if c in curr_act_grid.columns: 
                        new_df.loc[new_df["Activity"] == act, c] = curr_act_grid.loc[curr_act_grid["Activity"] == act, c].values[0]
            else:
                rules = st.session_state.recurrence_rules.get(act, {})
                valid_dates = calculate_recurring_days(selected_year, selected_month, rules.get("weekdays", []), rules.get("weeks", ["All"]))
                for vd in valid_dates:
                    d_obj = datetime.datetime.strptime(vd, "%Y-%m-%d")
                    col_idx = d_obj.day - 1
                    new_df.loc[new_df["Activity"] == act, day_cols[col_idx]] = True
                    
        st.session_state[act_grid_key] = new_df
        
    edited_act_df = st.data_editor(
        st.session_state[act_grid_key],
        column_config=act_col_config,
        hide_index=True, use_container_width=True,
        key=f"editor_{act_grid_key}"
    )
    st.session_state[act_grid_key] = edited_act_df
    
    or_days_formatted = []
    outpatient_setup = {c: {'days': [], 'capable': []} for c in current_clinics}
    
    for idx, row in edited_act_df.iterrows():
        act = row["Activity"]
        for i, d in enumerate(month_dates):
            d_str = d.strftime("%Y-%m-%d")
            col_name = day_cols[i]
            if row[col_name]:
                if act == "OR": or_days_formatted.append(d_str)
                else: outpatient_setup[act]['days'].append(d_str)
                    
    for c in current_clinics:
        if c in edited_cap_df.columns:
            outpatient_setup[c]['capable'] = edited_cap_df[edited_cap_df[c] == True]["Doctor"].dropna().tolist()

with tab4:
    st.subheader("⚙️ Generate Draft Schedule")
    st.markdown("The algorithm will mathematically balance hours, nights, and holidays, outputting an editable draft in **Tab 5**.")
    
    if st.button("🛠️ GENERATE DRAFT", use_container_width=True):
        if len(current_doctors) == 0:
            st.error("❌ Please add at least one doctor to the Persistent Settings.")
        else:
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
                for idx, row in edited_absences_df.iterrows():
                    doc = row["Doctor"]
                    for i, d in enumerate(month_dates):
                        d_str = d.strftime("%Y-%m-%d")
                        col_name = day_cols[i]
                        val = row.get(col_name, "")
                        if val in ["F", "X", "P", "N"]:
                            daily_absences[(doc, d_str)] = val

                manual_shifts = []
                manual_super_holidays = []
                for w_idx, df_w in edited_manual_grids.items():
                    for idx, row in df_w.iterrows():
                        shift_val = str(row["Shift"]).strip().upper()
                        for d in ui_roster_dates[w_idx*7 : (w_idx*7)+7]:
                            d_str = d.strftime("%Y-%m-%d")
                            val = str(row.get(d_str, "")).strip().upper()
                            if shift_val == '🌟 SUPER HOLIDAY':
                                if val == 'YES': manual_super_holidays.append(d_str)
                            elif val and val in current_doctors:
                                manual_shifts.append((val, d_str, shift_val))

                success, draft_grids, warning = generate_draft_schedule(
                    year=selected_year, month=selected_month, conditional_or_days=or_days_formatted, 
                    manual_festivities=[], manual_super_holidays=manual_super_holidays, manual_assignments=manual_shifts,
                    doctor_capabilities=doc_capabilities_dict, outpatient_configs=outpatient_setup, 
                    daily_absences=daily_absences, private_practice_afternoons=pp_dict, doctors_list=current_doctors, debug_mode=debug_toggle
                )
                
                if success:
                    st.success("✅ Draft generated successfully! Go to **Tab 5 (🚀 Edit & Publish)** to review, modify, and publish the final Excel.")
                    if warning: st.warning(warning)
                    st.session_state.generated_draft_grids = draft_grids
                    for k in list(st.session_state.keys()):
                        if k.startswith("draft_editor_"): del st.session_state[k]
                else:
                    st.error(f"❌ {warning}")

with tab5:
    st.subheader("🚀 Review, Edit & Publish")
    
    if "generated_draft_grids" not in st.session_state:
        st.info("No draft generated yet. Go to **Tab 4** to generate the base schedule first.")
    else:
        st.markdown("Make any last-minute human adjustments here. The Live Fairness Dashboard below will update instantly as you edit.")
        
        edited_final_drafts = {}
        for w_idx in range(0, len(ui_roster_dates) // 7):
            st.markdown(f"#### Draft Week {w_idx + 1}")
            week_dates = ui_roster_dates[w_idx*7 : (w_idx*7)+7]
            col_config = {"Shift": st.column_config.TextColumn("Shift", disabled=True)}
            for d in week_dates:
                d_str = d.strftime("%Y-%m-%d")
                display_name = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
                col_config[d_str] = st.column_config.SelectboxColumn(display_name, options=[""] + current_doctors)
                
            edited_final_drafts[w_idx] = st.data_editor(
                st.session_state.generated_draft_grids[w_idx], 
                column_config=col_config, 
                hide_index=True, 
                use_container_width=True,
                key=f"draft_editor_week_{w_idx}_{month_key}"
            )
            st.markdown("---")

        st.markdown("### 📊 Live Fairness Dashboard")
        
        num_days = len(ui_roster_dates)
        date_to_idx = {d.strftime("%Y-%m-%d"): idx for idx, d in enumerate(ui_roster_dates)}
        it_holidays = holidays.IT(years=[selected_year-1, selected_year, selected_year+1])
        
        manual_super_holidays = []
        for w_idx, df_w in edited_manual_grids.items():
            for idx, row in df_w.iterrows():
                shift_val = str(row["Shift"]).strip().upper()
                for d in ui_roster_dates[w_idx*7 : (w_idx*7)+7]:
                    d_str = d.strftime("%Y-%m-%d")
                    val = str(row.get(d_str, "")).strip().upper()
                    if shift_val == '🌟 SUPER HOLIDAY' and val == 'YES':
                        manual_super_holidays.append(d_str)

        daily_absences = {}
        for idx, row in edited_absences_df.iterrows():
            doc = row["Doctor"]
            for i, d in enumerate(month_dates):
                d_str = d.strftime("%Y-%m-%d")
                col_name = day_cols[i]
                val = row.get(col_name, "")
                if val in ["F", "X", "P", "N"]:
                    daily_absences[(doc, d_str)] = val

        doc_contract_off_set = {d: set() for d in current_doctors}
        for (d, date_str), val in daily_absences.items():
            if val == 'F' and date_str in date_to_idx:
                doc_contract_off_set[d].add(date_str)
                
        doc_contract_days = {d: max(1, num_days - len(doc_contract_off_set[d])) for d in current_doctors}

        doc_schedule = {doc: {day_idx: [] for day_idx in range(num_days)} for doc in current_doctors}
        for w_idx, df_w in edited_final_drafts.items():
            week_dates = ui_roster_dates[w_idx*7 : (w_idx*7)+7]
            for idx, row in df_w.iterrows():
                s = row["Shift"]
                for i, d_date in enumerate(week_dates):
                    day_idx = (w_idx * 7) + i
                    if day_idx >= num_days: continue
                    d_str = d_date.strftime("%Y-%m-%d")
                    assigned_doc = str(row.get(d_str, "")).strip().upper()
                    if assigned_doc and assigned_doc in current_doctors:
                        doc_schedule[assigned_doc][day_idx].append(s)

        shifts = ['WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']
        for out_type in outpatient_setup.keys(): shifts.append(f'OUT_{out_type}_AM')
        day_active = [s for s in shifts if s not in ['NIGHT', 'REP_NIGHT', 'REP_DAY']]

        monthly_stats = compute_monthly_stats(doc_schedule, num_days, ui_roster_dates, day_active, it_holidays, [], manual_super_holidays, current_doctors)
        
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
