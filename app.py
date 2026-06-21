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
# 1. PERSISTENT SETTINGS DATABASE
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
                return json.load(f)
        except Exception:
            pass
            
    docs_base = [{"Doctor": d, "Private Practice (Afternoon)": "", "Color": "White"} for d in default_doctors]
    caps = [{"Doctor": d, "Ward Preferred": d in ["BURGAZZI", "ROBERTI"], "OR Capable": d in ["CACCAMO", "NUCCI"]} for d in default_doctors]
    for c in caps:
        for clin in default_clinics: c[clin] = False
        
    return {"clinics": default_clinics, "docs_base": docs_base, "capabilities": caps}

def save_persistent_settings(clinics_list, docs_df, cap_df):
    data = {
        "clinics": clinics_list,
        "docs_base": docs_df.fillna("").to_dict('records'),
        "capabilities": cap_df.fillna(False).to_dict('records')
    }
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(data, f, indent=4)


# ==========================================
# 2. THE MATH ENGINE (Core Logic)
# ==========================================
def get_roster_dates(year, month):
    cal = calendar.Calendar(firstweekday=0) 
    mondays = []
    for week in cal.monthdatescalendar(year, month):
        if week[0].month == month:
            if week[0] not in mondays:
                mondays.append(week[0])
                
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
        if dt.weekday() in target_days:
            dates.append(dt)
    return dates

def generate_cardiology_schedule(year, month, conditional_or_days, manual_festivities, manual_assignments,
                                 doctor_capabilities, outpatient_configs,
                                 ferie, leave_weeks, desiderate, private_practice_afternoons, doctor_colors, doctors_list, 
                                 commit_to_history=False, debug_mode=False):
    
    doctors = doctors_list
    roster_dates = get_roster_dates(year, month)
    num_days = len(roster_dates)
    date_to_idx = {d.strftime("%Y-%m-%d"): idx for idx, d in enumerate(roster_dates)}
    it_holidays = holidays.IT(years=[year-1, year, year+1])
    
    if os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE, 'r') as f:
            lifetime = json.load(f)
    else:
        lifetime = {}
        
    for doc in doctors:
        if doc not in lifetime:
            lifetime[doc] = {'nights': 0.0, 'doubles': 0.0, 'saturdays': 0.0, 'sundays': 0.0, 'holidays': 0.0, 'super_holidays': 0.0, 'golden_weekends': 0.0, 'reps': 0.0}
        elif isinstance(lifetime[doc], int): 
            lifetime[doc] = {'nights': 0.0, 'doubles': 0.0, 'saturdays': 0.0, 'sundays': 0.0, 'holidays': float(lifetime[doc]), 'super_holidays': 0.0, 'golden_weekends': 0.0, 'reps': 0.0}
        else:
            for key in ['nights', 'doubles', 'saturdays', 'sundays', 'holidays', 'super_holidays', 'golden_weekends', 'reps']:
                if key not in lifetime[doc]: 
                    lifetime[doc][key] = 0.0
                else:
                    lifetime[doc][key] = float(lifetime[doc][key])

    sunday_equivalent_days = []
    for day_idx, current_date in enumerate(roster_dates):
        is_sunday = current_date.weekday() == 6
        is_holiday = current_date in it_holidays or current_date.strftime("%Y-%m-%d") in manual_festivities
        if is_sunday or is_holiday:
            sunday_equivalent_days.append(day_idx)
            
    all_leave_dates = []
    for d, dates in ferie.items(): all_leave_dates.extend([(d, date_str) for date_str in dates])
    for d, dates in desiderate.items(): all_leave_dates.extend([(d, date_str) for date_str in dates])
    for d, start_mon_str in leave_weeks:
        if start_mon_str in date_to_idx:
            start_idx = date_to_idx[start_mon_str]
            for i in range(7):
                if start_idx + i < num_days:
                    all_leave_dates.extend([(d, roster_dates[start_idx + i].strftime("%Y-%m-%d"))])
        
    day_counts = Counter([date for doc, date in all_leave_dates])
    overlaps = [date for date, count in day_counts.items() if count > 1]
    overlap_warning = f"⚠️ Notice: Multiple doctors are off on: {overlaps}" if overlaps else ""

    doc_unavailable_days = {d: 0 for d in doctors}
    for d in doctors:
        off_set = set()
        if d in ferie:
            for date_str in ferie[d]:
                if date_str in date_to_idx: off_set.add(date_str)
        if d in desiderate:
            for date_str in desiderate[d]:
                if date_str in date_to_idx: off_set.add(date_str)
        for doc_l, start_mon in leave_weeks:
            if doc_l == d:
                start_idx_l = date_to_idx.get(start_mon, -1)
                if start_idx_l != -1:
                    for i in range(7):
                        if start_idx_l + i < num_days:
                            off_set.add(roster_dates[start_idx_l + i].strftime("%Y-%m-%d"))
        doc_unavailable_days[d] = len(off_set)
        
    doc_active_days = {d: max(1, num_days - doc_unavailable_days[d]) for d in doctors}

    for day_idx in range(num_days):
        current_date_str = roster_dates[day_idx].strftime("%Y-%m-%d")
        if day_idx in sunday_equivalent_days: min_bodies_needed = 3
        else: min_bodies_needed = 4
        unavailable = 0
        for d in doctors:
            is_off = False
            if d in ferie and current_date_str in ferie[d]: is_off = True
            if d in desiderate and current_date_str in desiderate[d]: is_off = True
            for doc_l, start_mon in leave_weeks:
                if doc_l == d:
                    start_idx_l = date_to_idx.get(start_mon, -1)
                    if start_idx_l != -1 and start_idx_l - 1 <= day_idx <= start_idx_l + 6:
                        is_off = True
            if is_off: unavailable += 1
            
        available_bodies = len(doctors) - unavailable
        if available_bodies < min_bodies_needed:
            return False, None, "", f"🛑 CRITICAL STAFFING SHORTAGE ON {current_date_str}: You only have {available_bodies} doctors available, but you need at least {min_bodies_needed} to legally run the department."

    shifts = ['WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']
    for out_type in outpatient_configs.keys(): shifts.append(f'OUT_{out_type}_AM')
    
    day_active = [s for s in shifts if s not in ['NIGHT', 'REP_NIGHT', 'REP_DAY']]
    
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

    for d, dates in ferie.items():
        if d in doctors:
            for date_str in dates:
                if date_str in date_to_idx:
                    for s in shifts: model.Add(work[(d, date_to_idx[date_str], s)] == 0)
                
    for d, dates in desiderate.items():
        if d in doctors:
            for date_str in dates:
                if date_str in date_to_idx:
                    for s in shifts: model.Add(work[(d, date_to_idx[date_str], s)] == 0)

    for d, start_monday_str in leave_weeks:
        if d in doctors and start_monday_str in date_to_idx:
            start_idx = date_to_idx[start_monday_str]
            for i in range(7):
                if 0 <= start_idx + i < num_days:
                    for s in shifts: model.Add(work[(d, start_idx + i, s)] == 0)
            if 0 <= start_idx - 1 < num_days:
                for s in shifts: model.Add(work[(d, start_idx - 1, s)] == 0)
            if 0 <= start_idx - 4 < num_days:
                objective_terms.append(50 * work[(d, start_idx - 4, 'NIGHT')])
                
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
                    if not doctor_capabilities.get(d, {}).get("OR Capable", False): 
                        model.Add(work[(d, day_idx, 'OR_AM')] == 0)
            else:
                for d in doctors: model.Add(work[(d, day_idx, 'OR_AM')] == 0)
                
            for out_type, config in outpatient_configs.items():
                s_name = f'OUT_{out_type}_AM'
                if current_date_str in config['days']:
                    model.AddExactlyOne(work[(d, day_idx, s_name)] for d in doctors)
                    for d in doctors:
                        if d not in config['capable']: 
                            model.Add(work[(d, day_idx, s_name)] == 0)
                else:
                    for d in doctors: model.Add(work[(d, day_idx, s_name)] == 0)
        else: 
            model.AddExactlyOne(work[(d, day_idx, 'WARD_AM')] for d in doctors) 
            model.AddExactlyOne(work[(d, day_idx, 'WARD_PM')] for d in doctors) 
            model.AddExactlyOne(work[(d, day_idx, 'REP_DAY')] for d in doctors)
            
            for d in doctors:
                for s in shifts:
                    if s not in ['WARD_AM', 'WARD_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT']: 
                        model.Add(work[(d, day_idx, s)] == 0)

    day_name_to_num = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}

    for d in doctors:
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
            
            if d == 'FINIZIO':
                model.Add(work[(d, day_idx, 'NIGHT')] == 0)
                model.Add(work[(d, day_idx, 'REP_NIGHT')] == 0)
                model.Add(sum(work[(d, day_idx, s)] for s in shifts) <= 1)
                
            elif d == 'GAUDENZI':
                model.Add(work[(d, day_idx, 'NIGHT')] == 0)
                model.Add(work[(d, day_idx, 'REP_NIGHT')] == 0)
                model.Add(work[(d, day_idx, 'REP_DAY')] == 0)
                if weekday == 4: 
                    for s in shifts:
                        if s not in ['WARD_PM', 'URG_PM']: model.Add(work[(d, day_idx, s)] == 0)
                elif weekday == 5: pass 
                else: 
                    for s in shifts: model.Add(work[(d, day_idx, s)] == 0)
            else:
                if day_idx < num_days - 1:
                    is_pre_night = work[(d, day_idx + 1, 'NIGHT')]
                    model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 2)
                    if day_idx not in sunday_equivalent_days:
                        model.Add(sum(work[(d, day_idx, s)] for s in day_active) <= 1).OnlyEnforceIf(is_pre_night.Not())
                    
                    pre_night_double = model.NewBoolVar('')
                    model.Add(sum(work[(d, day_idx, s)] for s in day_active) == 2).OnlyEnforceIf(pre_night_double)
                    model.Add(sum(work[(d, day_idx, s)] for s in day_active) != 2).OnlyEnforceIf(pre_night_double.Not())
                    
                    ideal_pre_night = model.NewBoolVar('')
                    model.AddBoolAnd([is_pre_night, pre_night_double, work[(d, day_idx, 'REP_NIGHT')]]).OnlyEnforceIf(ideal_pre_night)
                    objective_terms.append(1000 * ideal_pre_night)
                else:
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
            for s in shifts:
                ruining_shifts.extend([work[(d, sat_idx, s)], work[(d, sun_idx, s)]])
                
            model.Add(sum(ruining_shifts) == 0).OnlyEnforceIf(gw_var)
            model.Add(sum(ruining_shifts) > 0).OnlyEnforceIf(gw_var.Not())
            doctor_gws.append(gw_var)
            objective_terms.append(-5 * int(lifetime[d]['golden_weekends']) * gw_var)
            
        if not debug_mode:
            model.Add(sum(doctor_gws) >= 1)

    weeks = [range(i, i + 7) for i in range(0, num_days, 7)]
    for w in weeks:
        if 'GAUDENZI' in doctors:
            model.Add(sum((12 if s in ['NIGHT', 'REP_NIGHT', 'REP_DAY'] else 6) * work[('GAUDENZI', day_idx, s)] for day_idx in w for s in shifts) <= 20)

    if not debug_mode:
        night_doctors = [d for d in doctors if d not in ['FINIZIO', 'GAUDENZI']]
        total_active_night_days = sum(doc_active_days[d] for d in night_doctors)
        
        holiday_doctors = [d for d in doctors if d != 'GAUDENZI']
        total_active_hol_days = sum(doc_active_days[d] for d in holiday_doctors)

        standard_doctors = [d for d in doctors if d != 'GAUDENZI']
        total_active_std_days = sum(doc_active_days[d] for d in standard_doctors)
        
        total_day_shifts = 0
        for day_idx in range(num_days):
            if day_idx not in sunday_equivalent_days:
                total_day_shifts += 4
                current_date_str = roster_dates[day_idx].strftime("%Y-%m-%d")
                if current_date_str in conditional_or_days: total_day_shifts += 1
                for out_type, config in outpatient_configs.items():
                    if current_date_str in config['days']: total_day_shifts += 1
            else:
                total_day_shifts += 2

        for d in night_doctors:
            expected_nights = (doc_active_days[d] / total_active_night_days) * num_days if total_active_night_days else 0
            min_n = max(0, int(expected_nights) - 1)
            max_n = int(expected_nights) + 2
            model.Add(sum(work[(d, day_idx, 'NIGHT')] for day_idx in range(num_days)) >= min_n)
            model.Add(sum(work[(d, day_idx, 'NIGHT')] for day_idx in range(num_days)) <= max_n)
            
            expected_rn = (doc_active_days[d] / total_active_night_days) * num_days if total_active_night_days else 0
            min_rn = max(0, int(expected_rn) - 2)
            max_rn = int(expected_rn) + 3
            model.Add(sum(work[(d, day_idx, 'REP_NIGHT')] for day_idx in range(num_days)) >= min_rn)
            model.Add(sum(work[(d, day_idx, 'REP_NIGHT')] for day_idx in range(num_days)) <= max_rn)

        for d in holiday_doctors:
            total_hols = len(sunday_equivalent_days) * 2
            expected_hols = (doc_active_days[d] / total_active_hol_days) * total_hols if total_active_hol_days else 0
            min_h = max(0, int(expected_hols) - 2)
            max_h = int(expected_hols) + 2
            model.Add(sum(work[(d, day_idx, s)] for day_idx in sunday_equivalent_days for s in ['WARD_AM', 'WARD_PM']) >= min_h)
            model.Add(sum(work[(d, day_idx, s)] for day_idx in sunday_equivalent_days for s in ['WARD_AM', 'WARD_PM']) <= max_h)
            
            expected_rd = (doc_active_days[d] / total_active_hol_days) * len(sunday_equivalent_days) if total_active_hol_days else 0
            min_rd = max(0, int(expected_rd) - 1)
            max_rd = int(expected_rd) + 2
            model.Add(sum(work[(d, day_idx, 'REP_DAY')] for day_idx in sunday_equivalent_days) >= min_rd)
            model.Add(sum(work[(d, day_idx, 'REP_DAY')] for day_idx in sunday_equivalent_days) <= max_rd)

        for d in standard_doctors:
            expected_days = (doc_active_days[d] / total_active_std_days) * total_day_shifts if total_active_std_days else 0
            min_d = max(0, int(expected_days) - 3)
            max_d = int(expected_days) + 5
            model.Add(sum(work[(d, day_idx, s)] for day_idx in range(num_days) for s in day_active) >= min_d)
            model.Add(sum(work[(d, day_idx, s)] for day_idx in range(num_days) for s in day_active) <= max_d)

    for d in doctors:
        if doctor_capabilities.get(d, {}).get("Ward Preferred", False):
            for day_idx in range(num_days):
                objective_terms.extend([20 * work[(d, day_idx, 'WARD_AM')], 20 * work[(d, day_idx, 'WARD_PM')]])

    for d in doctors:
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
                objective_terms.append(-800 * consecutive_ward)

        for day_idx in sunday_equivalent_days:
            ward_double = model.NewBoolVar('')
            model.AddBoolAnd([work[(d, day_idx, 'WARD_AM')], work[(d, day_idx, 'WARD_PM')]]).OnlyEnforceIf(ward_double)
            objective_terms.append(10 * ward_double) 
            urg_double = model.NewBoolVar('')
            model.AddBoolAnd([work[(d, day_idx, 'URG_AM')], work[(d, day_idx, 'URG_PM')]]).OnlyEnforceIf(urg_double)
            objective_terms.append(10 * urg_double) 

        for day_idx in range(num_days):
            curr_date = roster_dates[day_idx]
            is_easter = (it_holidays.get(curr_date) == "Pasqua di Resurrezione")
            is_proper_holiday = ((curr_date.month == 12 and curr_date.day == 25) or 
                                 (curr_date.month == 1 and curr_date.day == 1) or
                                 (curr_date.month == 4 and curr_date.day == 25) or
                                 (curr_date.month == 6 and curr_date.day in [1, 2]) or
                                 (curr_date.month == 8 and curr_date.day == 15) or is_easter)
            is_eve = (curr_date.month == 12 and curr_date.day in [24, 31])
            
            sat_pts = int(lifetime[d]['saturdays'] * 10) 
            sun_pts = int(lifetime[d]['sundays'] * 10)
            sh_pts = int(lifetime[d]['super_holidays'] * 10)
            
            if curr_date.weekday() == 5:
                objective_terms.append(-1 * sat_pts * work[(d, day_idx, 'NIGHT')])
                for s in day_active: objective_terms.append(-1 * int(sat_pts/2) * work[(d, day_idx, s)])
                
            if curr_date.weekday() == 6:
                objective_terms.append(-1 * sun_pts * work[(d, day_idx, 'NIGHT')])
                for s in day_active: objective_terms.append(-1 * int(sun_pts/2) * work[(d, day_idx, s)])
            
            if curr_date in it_holidays or curr_date.strftime("%Y-%m-%d") in manual_festivities: 
                worked_any = model.NewBoolVar('')
                model.Add(sum(work[(d, day_idx, s)] for s in shifts) > 0).OnlyEnforceIf(worked_any)
                model.Add(sum(work[(d, day_idx, s)] for s in shifts) == 0).OnlyEnforceIf(worked_any.Not())
                objective_terms.append(-4 * int(lifetime[d]['holidays']) * worked_any)
                
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
    solver.parameters.max_time_in_seconds = 45.0 
    status = solver.Solve(model)
    
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        monthly = {doc: {'nights': 0.0, 'doubles': 0.0, 'saturdays': 0.0, 'sundays': 0.0, 'holidays': 0.0, 'super_holidays': 0.0, 'golden_weekends': 0.0, 'reps': 0.0, 'total_shifts': 0.0, 'total_hours': 0.0} for doc in doctors}
        doc_schedule = {doc: {day_idx: [] for day_idx in range(num_days)} for doc in doctors}
        
        for day_idx in range(num_days):
            for s in shifts:
                for doc in doctors:
                    if solver.Value(work[(doc, day_idx, s)]) == 1: doc_schedule[doc][day_idx].append(s)

        for doc in doctors:
            for day_idx in range(num_days):
                curr_date = roster_dates[day_idx]
                worked_today = doc_schedule[doc][day_idx]
                active_shifts = [s for s in worked_today if s in day_active]
                rep_shifts = [s for s in worked_today if s in ['REP_DAY', 'REP_NIGHT']]
                
                for s in worked_today:
                    monthly[doc]['total_shifts'] += 1
                    monthly[doc]['total_hours'] += 12 if s in ['NIGHT', 'REP_NIGHT', 'REP_DAY'] else 6
                
                if 'NIGHT' in worked_today:
                    monthly[doc]['nights'] += 1
                    if commit_to_history: lifetime[doc]['nights'] += 1
                
                if len(rep_shifts) > 0:
                    monthly[doc]['reps'] += len(rep_shifts)
                    if commit_to_history: lifetime[doc]['reps'] += len(rep_shifts)
                    
                if len(active_shifts) >= 2:
                    monthly[doc]['doubles'] += 1
                    if commit_to_history: lifetime[doc]['doubles'] += 1
                    
                day_pts = 0.0
                if 'NIGHT' in worked_today: day_pts = 1.0
                else: day_pts = len(active_shifts) * 0.5
                
                if curr_date.weekday() == 5:
                    monthly[doc]['saturdays'] += day_pts
                    if commit_to_history: lifetime[doc]['saturdays'] += day_pts
                elif curr_date.weekday() == 6:
                    monthly[doc]['sundays'] += day_pts
                    if commit_to_history: lifetime[doc]['sundays'] += day_pts
                    
                if worked_today and (curr_date in it_holidays or curr_date.strftime("%Y-%m-%d") in manual_festivities):
                    monthly[doc]['holidays'] += 1
                    if commit_to_history: lifetime[doc]['holidays'] += 1

                is_easter = (it_holidays.get(curr_date) == "Pasqua di Resurrezione")
                is_proper_holiday = ((curr_date.month == 12 and curr_date.day == 25) or 
                                     (curr_date.month == 1 and curr_date.day == 1) or
                                     (curr_date.month == 4 and curr_date.day == 25) or
                                     (curr_date.month == 6 and curr_date.day in [1, 2]) or
                                     (curr_date.month == 8 and curr_date.day == 15) or is_easter)
                is_eve = (curr_date.month == 12 and curr_date.day in [24, 31])
                
                sh_pts = 0.0
                if (is_eve and 'NIGHT' in worked_today) or (is_proper_holiday and len(active_shifts) > 0):
                    sh_pts = 1.0
                elif (is_proper_holiday and len(rep_shifts) > 0) or (is_eve and 'REP_NIGHT' in worked_today):
                    sh_pts = 0.5
                    
                if sh_pts > 0:
                    monthly[doc]['super_holidays'] += sh_pts
                    if commit_to_history: lifetime[doc]['super_holidays'] += sh_pts

            for week_start_idx in range(0, num_days, 7):
                ruined_fri = 'NIGHT' in doc_schedule[doc][week_start_idx + 4] or 'REP_NIGHT' in doc_schedule[doc][week_start_idx + 4]
                if not ruined_fri and not doc_schedule[doc][week_start_idx + 5] and not doc_schedule[doc][week_start_idx + 6]:
                    monthly[doc]['golden_weekends'] += 1
                    if commit_to_history: lifetime[doc]['golden_weekends'] += 1

        # ==========================================
        # 🟢 EXCEL (.xlsx) EXPORT ENGINE
        # ==========================================
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        
        worksheet = workbook.add_worksheet('Master Schedule')
        
        header_format = workbook.add_format({'bold': True, 'bg_color': '#E8E8E8', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
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
            week_shifts_night = ['REP_NIGHT', 'NIGHT']
            
            or_open = False
            outpatient_open = {out_type: False for out_type in outpatient_configs.keys()}
            rep_day_open = False
            
            for i in range(7):
                day_idx = week_start_idx + i
                d_str = roster_dates[day_idx].strftime("%Y-%m-%d")
                if day_idx in sunday_equivalent_days: rep_day_open = True
                if d_str in conditional_or_days: or_open = True
                for out_type, config in outpatient_configs.items():
                    if d_str in config['days']: outpatient_open[out_type] = True
                    
            if or_open: week_shifts_am.append('OR_AM')
            for out_type, is_open in outpatient_open.items():
                if is_open: week_shifts_am.append(f'OUT_{out_type}_AM')

            ordered_shifts_this_week = week_shifts_am + week_shifts_pm
            if rep_day_open: ordered_shifts_this_week.append('REP_DAY')
            ordered_shifts_this_week.extend(week_shifts_night)
            
            for s in ordered_shifts_this_week:
                worksheet.write(row_cursor, 0, s.replace('_', ' '), shift_format)
                for i in range(7):
                    day_idx = week_start_idx + i
                    assigned_doc = ""
                    for d in doctors:
                        if solver.Value(work[(d, day_idx, s)]) == 1:
                            assigned_doc = d
                    
                    if assigned_doc:
                        worksheet.write(row_cursor, i + 1, assigned_doc, doc_formats[assigned_doc])
                    else:
                        if s == 'REP_DAY' and day_idx not in sunday_equivalent_days:
                            worksheet.write(row_cursor, i + 1, "-", gray_dash_format)
                        else:
                            worksheet.write(row_cursor, i + 1, "", empty_border)
                row_cursor += 1
                
            row_cursor += 2

        # --- THE DASHBOARD ---
        dashboard_header = ['Doctor', 'Proportional Target Hours', 'Actual Monthly Hours', 'Difference (+/-)',
                            'Monthly Nights', 'Monthly On-Call (Rep)', 'Monthly Doubles', 'Monthly Saturdays', 'Monthly Sundays', 'Monthly Holidays', 'Monthly Super Hols', 'Monthly Golden Wknds',
                            'LIFETIME Nights', 'LIFETIME On-Call', 'LIFETIME Doubles', 'LIFETIME Saturdays', 'LIFETIME Sundays', 'LIFETIME Holidays', 'LIFETIME Super Hols', 'LIFETIME Golden Wknds']
        
        worksheet.write(row_cursor, 0, "DOCTOR STATISTICS DASHBOARD", header_format)
        row_cursor += 1
        
        for col_idx, h in enumerate(dashboard_header):
            worksheet.write(row_cursor, col_idx, h, header_format)
        row_cursor += 1
        
        for doc in doctors:
            doc_target = int((doc_active_days[doc] / 7) * 20) if doc == 'GAUDENZI' else int((doc_active_days[doc] / 7) * 38)
            actual_hours = monthly[doc]['total_hours']
            difference = actual_hours - doc_target
            
            data_row = [
                doc, doc_target, actual_hours, f"+{difference}" if difference > 0 else str(difference),
                monthly[doc]['nights'], monthly[doc]['reps'], monthly[doc]['doubles'], f"{monthly[doc]['saturdays']:.1f}", f"{monthly[doc]['sundays']:.1f}", monthly[doc]['holidays'], f"{monthly[doc]['super_holidays']:.1f}", monthly[doc]['golden_weekends'],
                lifetime[doc]['nights'], lifetime[doc]['reps'], lifetime[doc]['doubles'], f"{lifetime[doc]['saturdays']:.1f}", f"{lifetime[doc]['sundays']:.1f}", lifetime[doc]['holidays'], f"{lifetime[doc]['super_holidays']:.1f}", lifetime[doc]['golden_weekends']
            ]
            for col_idx, val in enumerate(data_row):
                worksheet.write(row_cursor, col_idx, val, doc_formats[doc] if col_idx == 0 else empty_border)
            row_cursor += 1

        worksheet.set_column(0, 0, 20)
        worksheet.set_column(1, 7, 18)

        # --- 2. THE INDIVIDUAL DOCTOR TABS ---
        for doc in doctors:
            doc_sheet = workbook.add_worksheet(doc)
            
            doc_header_fmt = workbook.add_format({'bold': True, 'bg_color': doctor_colors.get(doc, '#E8E8E8'), 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            cell_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
            off_fmt = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'font_color': '#A0A0A0', 'italic': True})
            holiday_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFF3E0', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

            days_header = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            for col_idx, day_name in enumerate(days_header):
                doc_sheet.write(0, col_idx, day_name, doc_header_fmt)
                
            doc_sheet.set_column(0, 6, 16) 

            row_c = 1
            for w_idx, week_start_idx in enumerate(range(0, num_days, 7)):
                doc_sheet.set_row(row_c, 55) 
                for i in range(7):
                    day_idx = week_start_idx + i
                    curr_date = roster_dates[day_idx]
                    
                    my_shifts = []
                    for s in shifts:
                        if solver.Value(work[(doc, day_idx, s)]) == 1:
                            if 'REP' in s:
                                my_shifts.append("📞 " + s.replace('_', ' '))
                            else:
                                my_shifts.append(s.replace('_', ' '))
                            
                    is_holiday_or_weekend = (curr_date.weekday() >= 5 or day_idx in sunday_equivalent_days)
                    
                    if my_shifts:
                        shift_text = f"{curr_date.strftime('%b %d')}\n\n" + "\n".join(my_shifts)
                        fmt = holiday_fmt if is_holiday_or_weekend else cell_fmt
                    else:
                        shift_text = f"{curr_date.strftime('%b %d')}\n\nOFF"
                        fmt = holiday_fmt if is_holiday_or_weekend else off_fmt

                    doc_sheet.write(row_c, i, shift_text, fmt)
                row_c += 1

        # --- 3. REPERIBILITÀ UFFICIALE ---
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
                rep_sheet.write(row_c, 0, s.replace('_', ' '), shift_format)
                for i in range(7):
                    day_idx = week_start_idx + i
                    assigned_doc = ""
                    for d in doctors:
                        if solver.Value(work[(d, day_idx, s)]) == 1:
                            assigned_doc = d
                    
                    if assigned_doc:
                        rep_sheet.write(row_c, i + 1, assigned_doc, doc_formats[assigned_doc])
                    else:
                        if s == 'REP_DAY' and day_idx not in sunday_equivalent_days:
                            rep_sheet.write(row_c, i + 1, "-", gray_dash_format)
                        else:
                            rep_sheet.write(row_c, i + 1, "", empty_border)
                row_c += 1
            row_c += 2

        workbook.close()
        
        debug_msg = " [DEBUG MODE ENABLED: Fairness metrics ignored]" if debug_mode else ""
        if commit_to_history and not debug_mode: 
            with open(COUNTER_FILE, 'w') as f: json.dump(lifetime, f, indent=4)
            push_to_github(COUNTER_FILE, "Auto-sync: Official Roster Published (Stats Saved)")
            return True, output.getvalue(), overlap_warning, f"Schedule solved! Block: {roster_dates[0].strftime('%b %d')} to {roster_dates[-1].strftime('%b %d')}. (OFFICIAL: Stats Saved){debug_msg}"
        else:
            return True, output.getvalue(), overlap_warning, f"Schedule solved! Block: {roster_dates[0].strftime('%b %d')} to {roster_dates[-1].strftime('%b %d')}. (SIMULATION: Stats NOT Saved){debug_msg}"
            
    else:
        return False, None, overlap_warning, "Constraints are too tight. The algorithm cannot find a mathematically legal schedule. Try using Emergency Debug Mode."

# ==========================================
# 3. THE GRAPHICAL USER INTERFACE (GUI)
# ==========================================
st.set_page_config(page_title="Cardiology Scheduler", page_icon="🩺", layout="wide")

if "app_initialized" not in st.session_state:
    settings = load_persistent_settings()
    st.session_state.clinics_df = pd.DataFrame({"Clinic Name": settings["clinics"]})
    st.session_state.docs_df = pd.DataFrame(settings["docs_base"]) 
    st.session_state.capabilities_df = pd.DataFrame(settings["capabilities"]) 
    st.session_state.absences_df = pd.DataFrame(columns=["Doctor", "Ferie (YYYY-MM-DD)", "Desiderate (YYYY-MM-DD)", "Leave Weeks (Type the Monday)"])
    st.session_state.app_initialized = True

current_clinics = [str(c).strip().upper().replace(" ", "_") for c in st.session_state.clinics_df["Clinic Name"].dropna().unique() if str(c).strip()]

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
    
    # Render without Key/Copy to avoid Double-Click bugs
    st.session_state.docs_df = st.data_editor(
        st.session_state.docs_df,
        num_rows="dynamic",
        column_config=doc_col_config,
        hide_index=True, use_container_width=True
    )
    
    current_doctors = [str(d).strip().upper() for d in st.session_state.docs_df["Doctor"].dropna().unique() if str(d).strip()]
    
    st.markdown("---")
    st.header("👨‍⚕️ Doctor Capabilities")
    
    cap_cols = ["Doctor", "Ward Preferred", "OR Capable"] + current_clinics
    
    # Safely synchronize capabilities without destroying memory references
    needs_cap_update = False
    if list(st.session_state.capabilities_df.columns) != cap_cols: needs_cap_update = True
    if set(st.session_state.capabilities_df["Doctor"].tolist()) != set(current_doctors): needs_cap_update = True
    
    if needs_cap_update:
        existing_docs = st.session_state.capabilities_df["Doctor"].tolist()
        for d in current_doctors:
            if d not in existing_docs:
                new_row = {c: False for c in cap_cols}
                new_row["Doctor"] = d
                st.session_state.capabilities_df.loc[len(st.session_state.capabilities_df)] = new_row
        st.session_state.capabilities_df = st.session_state.capabilities_df[st.session_state.capabilities_df["Doctor"].isin(current_doctors)]
        for c in cap_cols:
            if c not in st.session_state.capabilities_df.columns: st.session_state.capabilities_df[c] = False
        st.session_state.capabilities_df = st.session_state.capabilities_df[cap_cols].reset_index(drop=True)

    st.session_state.capabilities_df = st.data_editor(st.session_state.capabilities_df, hide_index=True, use_container_width=True)
    
    # Save the states to persistent memory
    save_persistent_settings(current_clinics, st.session_state.docs_df, st.session_state.capabilities_df)
    
    st.markdown("---")
    st.header("🛠️ Troubleshooter")
    debug_toggle = st.checkbox("🚨 Enable Emergency Debug Mode", help="Ignores all equity rules to force a schedule.")

ui_roster_dates = get_roster_dates(selected_year, selected_month)

tab1, tab2, tab3, tab4 = st.tabs(["📝 Absences (Ferie/Leave)", "🔒 Forced Shifts", "🏥 OR & Outpatient Clinics", "🚀 Generate Schedule"])

with tab1:
    st.subheader("Monthly Absences (Ferie, Desiderate, Leave Weeks)")
    st.markdown("Type the dates. Multiple dates must be separated by commas (e.g. `2026-07-04, 2026-07-05`).")
    
    abs_docs = st.session_state.absences_df["Doctor"].tolist()
    if set(abs_docs) != set(current_doctors):
        for d in current_doctors:
            if d not in abs_docs:
                st.session_state.absences_df.loc[len(st.session_state.absences_df)] = [d, "", "", ""]
        st.session_state.absences_df = st.session_state.absences_df[st.session_state.absences_df["Doctor"].isin(current_doctors)].reset_index(drop=True)
    
    st.session_state.absences_df = st.data_editor(
        st.session_state.absences_df,
        column_config={"Doctor": st.column_config.TextColumn("Doctor Name", disabled=True)},
        hide_index=True, use_container_width=True
    )

with tab2:
    st.subheader("Visual Override Grid")
    st.markdown("Click on any empty cell to forcefully lock a specific doctor into that shift.")
    
    if st.session_state.get("current_ym") != f"{selected_year}-{selected_month}":
        st.session_state["current_ym"] = f"{selected_year}-{selected_month}"
        for k in list(st.session_state.keys()):
            if k.startswith("grid_week_"):
                del st.session_state[k]

    dynamic_shift_options = ['WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT', 'REP_DAY', 'REP_NIGHT'] + [f'OUT_{c}_AM' for c in current_clinics]

    for w_idx, week_start_idx in enumerate(range(0, len(ui_roster_dates), 7)):
        st.markdown(f"#### Week {w_idx + 1}")
        week_dates = ui_roster_dates[week_start_idx : week_start_idx + 7]
        col_names = [d.strftime("%Y-%m-%d") for d in week_dates]
        state_key = f"grid_week_{w_idx}"
        
        if state_key not in st.session_state:
            init_dict = {"Shift": dynamic_shift_options}
            for c in col_names: init_dict[c] = [""] * len(dynamic_shift_options)
            st.session_state[state_key] = pd.DataFrame(init_dict)
            
        # Re-sync shifts if new clinics added
        if list(st.session_state[state_key]["Shift"]) != dynamic_shift_options:
            init_dict = {"Shift": dynamic_shift_options}
            for c in col_names: init_dict[c] = [""] * len(dynamic_shift_options)
            st.session_state[state_key] = pd.DataFrame(init_dict)
            
        col_config = {"Shift": st.column_config.TextColumn("Shift", disabled=True)}
        for d in week_dates:
            d_str = d.strftime("%Y-%m-%d")
            display_name = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
            col_config[d_str] = st.column_config.SelectboxColumn(display_name, options=[""] + current_doctors)
            
        st.session_state[state_key] = st.data_editor(
            st.session_state[state_key], 
            column_config=col_config, 
            hide_index=True, 
            use_container_width=True
        )
        st.markdown("---")

with tab3:
    all_dates_in_month = [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1)]
    
    st.subheader("🏥 Operating Room (OR)")
    or_weekdays = st.multiselect("Standard OR Weekdays", 
                                 ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"], 
                                 default=["Monday", "Friday"])
    
    auto_or_dates = get_dates_for_weekdays(selected_year, selected_month, or_weekdays)
    or_days_input = st.multiselect("Final OR Dates (Add/Remove specific days)", 
                                   all_dates_in_month, 
                                   default=auto_or_dates,
                                   key=f"or_dates_{selected_year}_{selected_month}")
    or_days_formatted = [d.strftime("%Y-%m-%d") for d in or_days_input]
    
    or_docs = st.session_state.capabilities_df[st.session_state.capabilities_df["OR Capable"] == True]["Doctor"].dropna().tolist() if "OR Capable" in st.session_state.capabilities_df.columns else []
    st.caption(f"**OR Capable Doctors (from Sidebar):** {', '.join(or_docs) if or_docs else 'None'}")

    st.markdown("---")
    st.subheader("Manage Outpatient Clinics")
    st.markdown("##### 1. Define Clinics")
    
    st.session_state.clinics_df = st.data_editor(
        st.session_state.clinics_df, 
        num_rows="dynamic",
        column_config={"Clinic Name": st.column_config.TextColumn("Clinic Name", required=True)},
        use_container_width=True,
        hide_index=True
    )
    current_clinics = [str(c).strip().upper().replace(" ", "_") for c in st.session_state.clinics_df["Clinic Name"].dropna().unique() if str(c).strip()]
    
    st.markdown("##### 2. Schedule Clinics")
    
    outpatient_setup = {}
    if current_clinics:
        for clinic in current_clinics:
            with st.expander(f"🩺 {clinic.replace('_', ' ')} Schedule", expanded=True):
                if clinic in st.session_state.capabilities_df.columns:
                    capable = st.session_state.capabilities_df[st.session_state.capabilities_df[clinic] == True]["Doctor"].dropna().tolist()
                else:
                    capable = []
                    
                st.caption(f"**Assigned Doctors (from Sidebar):** {', '.join(capable) if capable else 'None'}")
                
                clinic_weekdays = st.multiselect("Standard Weekly Days", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"], key=f"wd_{clinic}")
                auto_clinic_dates = get_dates_for_weekdays(selected_year, selected_month, clinic_weekdays)
                final_clinic_dates = st.multiselect("Final Specific Dates", all_dates_in_month, default=auto_clinic_dates, key=f"dates_{clinic}_{selected_year}_{selected_month}")
                
                outpatient_setup[clinic] = {
                    'days': [d.strftime("%Y-%m-%d") for d in final_clinic_dates],
                    'capable': capable
                }
    else:
        st.info("No clinics defined. Add a clinic name above to schedule it.")

with tab4:
    st.subheader("Generate File")
    
    col_sim, col_pub = st.columns(2)
    
    with col_sim:
        simulate_clicked = st.button("🛠️ SIMULATE SCHEDULE (No Save)", use_container_width=True)
    with col_pub:
        publish_clicked = st.button("🚀 APPROVE & PUBLISH (Save Stats)", use_container_width=True, type="primary")
    
    if simulate_clicked or publish_clicked:
        is_commit = publish_clicked
        
        if len(current_doctors) == 0:
            st.error("❌ Please add at least one doctor to the Persistent Settings in the Sidebar.")
        else:
            with st.spinner("Calculating mathematical constraints... This may take up to 45 seconds."):
                
                pp_dict = {}
                doc_colors_ui = {}
                doc_capabilities_dict = {}

                for idx, row in st.session_state.docs_df.iterrows():
                    doc = str(row["Doctor"]).strip().upper()
                    if not doc: continue
                    pp = row.get("Private Practice (Afternoon)", "")
                    if pd.notna(pp) and pp: pp_dict[doc] = pp
                    col = row.get("Color", "White")
                    doc_colors_ui[doc] = COLOR_PALETTE.get(col, "#FFFFFF")

                for idx, row in st.session_state.capabilities_df.iterrows():
                    doc = row["Doctor"]
                    doc_capabilities_dict[doc] = {col: bool(row[col]) for col in cap_cols if col != "Doctor"}

                ferie_dict = {}
                desiderate_dict = {}
                leave_list = []

                for idx, row in st.session_state.absences_df.iterrows():
                    doc = str(row["Doctor"]).strip().upper()
                    if not doc: continue
                    
                    ferie = row.get("Ferie (YYYY-MM-DD)", "")
                    if pd.notna(ferie) and ferie:
                        dates = [d.strip() for d in str(ferie).split(",") if d.strip()]
                        if dates: ferie_dict[doc] = dates
                        
                    des = row.get("Desiderate (YYYY-MM-DD)", "")
                    if pd.notna(des) and des:
                        dates = [d.strip() for d in str(des).split(",") if d.strip()]
                        if dates: desiderate_dict[doc] = dates
                        
                    l_w = row.get("Leave Weeks (Type the Monday)", "")
                    if pd.notna(l_w) and l_w:
                        dates = [d.strip() for d in str(l_w).split(",") if d.strip()]
                        for d_str in dates: leave_list.append((doc, d_str))

                manual_shifts = []
                for w_idx in range(0, len(ui_roster_dates) // 7):
                    state_key = f"grid_week_{w_idx}"
                    if state_key in st.session_state:
                        df_w = st.session_state[state_key]
                        for idx, row in df_w.iterrows():
                            shift_val = str(row["Shift"]).strip().upper()
                            for d in ui_roster_dates[w_idx*7 : (w_idx*7)+7]:
                                d_str = d.strftime("%Y-%m-%d")
                                doc = row.get(d_str, "")
                                if pd.notna(doc) and str(doc).strip():
                                    manual_shifts.append((str(doc).strip().upper(), d_str, shift_val))

                manual_festivities = []

                success, excel_data, warnings, message = generate_cardiology_schedule(
                    year=selected_year, month=selected_month, conditional_or_days=or_days_formatted, 
                    manual_festivities=manual_festivities, manual_assignments=manual_shifts,
                    doctor_capabilities=doc_capabilities_dict, outpatient_configs=outpatient_setup, 
                    ferie=ferie_dict, leave_weeks=leave_list, desiderate=desiderate_dict,
                    private_practice_afternoons=pp_dict, doctor_colors=doc_colors_ui, doctors_list=current_doctors,
                    commit_to_history=is_commit, debug_mode=debug_toggle
                )
                
                if success:
                    st.success(f"✅ {message}")
                    if warnings: st.warning(warnings)
                    
                    timestamp = datetime.datetime.now().strftime("%H%M%S")
                    prefix = "OFFICIAL_" if is_commit else "DRAFT_"
                    dl_filename = f'{prefix}cardiology_schedule_{selected_year}_{selected_month}_{timestamp}.xlsx'
                    
                    st.download_button(
                        label="📥 Download Excel Schedule",
                        data=excel_data,
                        file_name=dl_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    st.error(f"❌ {message}")
                    if warnings: st.warning(warnings)
