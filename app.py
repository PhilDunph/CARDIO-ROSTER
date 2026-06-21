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

# ==========================================
# 1. THE MATH ENGINE (Core Logic)
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

def generate_cardiology_schedule(year, month, conditional_or_days, manual_festivities, manual_assignments,
                                 ward_preferred_doctors, or_capable_doctors, outpatient_configs,
                                 ferie, leave_weeks, desiderate, private_practice_afternoons, doctor_colors, doctors_list, 
                                 commit_to_history=False, debug_mode=False):
    
    doctors = doctors_list
    
    roster_dates = get_roster_dates(year, month)
    num_days = len(roster_dates)
    date_to_idx = {d.strftime("%Y-%m-%d"): idx for idx, d in enumerate(roster_dates)}
    it_holidays = holidays.IT(years=[year-1, year, year+1])
    
    counter_file = 'historical_counters.json'
    if os.path.exists(counter_file):
        with open(counter_file, 'r') as f:
            lifetime = json.load(f)
    else:
        lifetime = {}
        
    for doc in doctors:
        if doc not in lifetime:
            lifetime[doc] = {'nights': 0, 'doubles': 0, 'saturdays': 0, 'sundays': 0, 'holidays': 0, 'super_holidays': 0, 'golden_weekends': 0, 'reps': 0}
        elif isinstance(lifetime[doc], int): 
            lifetime[doc] = {'nights': 0, 'doubles': 0, 'saturdays': 0, 'sundays': 0, 'holidays': lifetime[doc], 'super_holidays': 0, 'golden_weekends': 0, 'reps': 0}
        else:
            for key in ['nights', 'doubles', 'saturdays', 'sundays', 'holidays', 'super_holidays', 'golden_weekends', 'reps']:
                if key not in lifetime[doc]: lifetime[doc][key] = 0

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

    # PRE-FLIGHT SCANNER
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
            return False, None, "", f"🛑 CRITICAL STAFFING SHORTAGE ON {current_date_str}: You only have {available_bodies} doctors available, but you need at least {min_bodies_needed} to legally run the department. Remove some Ferie/Desiderate on this day."

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
            if key in work:  
                model.Add(work[key] == 1)

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
                    if d not in or_capable_doctors: model.Add(work[(d, day_idx, 'OR_AM')] == 0)
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
            objective_terms.append(-5 * lifetime[d]['golden_weekends'] * gw_var)
            
        if not debug_mode:
            model.Add(sum(doctor_gws) >= 1)

    weeks = [range(i, i + 7) for i in range(0, num_days, 7)]
    for w in weeks:
        if 'GAUDENZI' in doctors:
            model.Add(sum((12 if s in ['NIGHT', 'REP_NIGHT', 'REP_DAY'] else 6) * work[('GAUDENZI', day_idx, s)] for day_idx in w for s in shifts) <= 20)

    if not debug_mode:
        night_doctors = [d for d in doctors if d not in ['FINIZIO', 'GAUDENZI']]
        base_nights = num_days // len(night_doctors) if night_doctors else 0
        
        holiday_doctors = [d for d in doctors if d != 'GAUDENZI']
        base_hols = (len(sunday_equivalent_days) * 2) // len(holiday_doctors) if holiday_doctors else 0

        standard_doctors = [d for d in doctors if d != 'GAUDENZI']
        base_day_shifts = (num_days * 4 + len(sunday_equivalent_days)) // len(standard_doctors) if standard_doctors else 0

        for d in night_doctors:
            model.Add(sum(work[(d, day_idx, 'NIGHT')] for day_idx in range(num_days)) >= max(0, base_nights - 2))
            model.Add(sum(work[(d, day_idx, 'NIGHT')] for day_idx in range(num_days)) <= base_nights + 2)
        for d in holiday_doctors:
            model.Add(sum(work[(d, day_idx, s)] for day_idx in sunday_equivalent_days for s in ['WARD_AM', 'WARD_PM']) >= 0)
            model.Add(sum(work[(d, day_idx, s)] for day_idx in sunday_equivalent_days for s in ['WARD_AM', 'WARD_PM']) <= base_hols + 3)
        for d in standard_doctors:
            model.Add(sum(work[(d, day_idx, s)] for day_idx in range(num_days) for s in day_active) >= max(0, base_day_shifts - 4))
            model.Add(sum(work[(d, day_idx, s)] for day_idx in range(num_days) for s in day_active) <= base_day_shifts + 8)

    for d in ward_preferred_doctors:
        if d in doctors:
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
            worked_any = model.NewBoolVar('')
            model.Add(sum(work[(d, day_idx, s)] for s in shifts) > 0).OnlyEnforceIf(worked_any)
            model.Add(sum(work[(d, day_idx, s)] for s in shifts) == 0).OnlyEnforceIf(worked_any.Not())
            
            worked_day_shifts = model.NewBoolVar('')
            model.Add(sum(work[(d, day_idx, s)] for s in day_active) > 0).OnlyEnforceIf(worked_day_shifts)
            model.Add(sum(work[(d, day_idx, s)] for s in day_active) == 0).OnlyEnforceIf(worked_day_shifts.Not())
            
            if curr_date.weekday() == 5: objective_terms.append(-3 * lifetime[d]['saturdays'] * worked_any)
            if curr_date.weekday() == 6: objective_terms.append(-4 * lifetime[d]['sundays'] * worked_any)
            if curr_date in it_holidays or curr_date.strftime("%Y-%m-%d") in manual_festivities: 
                objective_terms.append(-4 * lifetime[d]['holidays'] * worked_any)
                
            is_easter = (it_holidays.get(curr_date) == "Pasqua di Resurrezione")
            if (curr_date.month == 12 and curr_date.day in [24, 31]):
                objective_terms.append(-10 * lifetime[d]['super_holidays'] * work[(d, day_idx, 'NIGHT')])
            if (curr_date.month == 12 and curr_date.day == 25) or (curr_date.month == 1 and curr_date.day == 1) or \
               (curr_date.month == 4 and curr_date.day == 25) or (curr_date.month == 6 and curr_date.day in [1, 2]) or \
               (curr_date.month == 8 and curr_date.day == 15) or is_easter:
                objective_terms.append(-10 * lifetime[d]['super_holidays'] * worked_day_shifts)

    model.Maximize(sum(objective_terms))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 45.0 
    status = solver.Solve(model)
    
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        monthly = {doc: {'nights': 0, 'doubles': 0, 'saturdays': 0, 'sundays': 0, 'holidays': 0, 'super_holidays': 0, 'golden_weekends': 0, 'reps': 0, 'total_shifts': 0, 'total_hours': 0} for doc in doctors}
        doc_schedule = {doc: {day_idx: [] for day_idx in range(num_days)} for doc in doctors}
        
        for day_idx in range(num_days):
            for s in shifts:
                for doc in doctors:
                    if solver.Value(work[(doc, day_idx, s)]) == 1: doc_schedule[doc][day_idx].append(s)

        for doc in doctors:
            for day_idx in range(num_days):
                curr_date = roster_dates[day_idx]
                worked_today = doc_schedule[doc][day_idx]
                
                for s in worked_today:
                    monthly[doc]['total_shifts'] += 1
                    monthly[doc]['total_hours'] += 12 if s in ['NIGHT', 'REP_NIGHT', 'REP_DAY'] else 6
                
                if 'NIGHT' in worked_today:
                    monthly[doc]['nights'] += 1
                    if commit_to_history: lifetime[doc]['nights'] += 1
                
                if 'REP_NIGHT' in worked_today or 'REP_DAY' in worked_today:
                    monthly[doc]['reps'] += 1
                    if commit_to_history: lifetime[doc]['reps'] += 1
                    
                day_shifts_worked = [s for s in worked_today if s not in ['NIGHT', 'REP_NIGHT']]
                if len(day_shifts_worked) >= 2:
                    monthly[doc]['doubles'] += 1
                    if commit_to_history: lifetime[doc]['doubles'] += 1
                    
                if worked_today:
                    if curr_date.weekday() == 5:
                        monthly[doc]['saturdays'] += 1
                        if commit_to_history: lifetime[doc]['saturdays'] += 1
                    elif curr_date.weekday() == 6:
                        monthly[doc]['sundays'] += 1
                        if commit_to_history: lifetime[doc]['sundays'] += 1
                    if curr_date in it_holidays or curr_date.strftime("%Y-%m-%d") in manual_festivities:
                        monthly[doc]['holidays'] += 1
                        if commit_to_history: lifetime[doc]['holidays'] += 1

                is_easter = (it_holidays.get(curr_date) == "Pasqua di Resurrezione")
                if 'NIGHT' in worked_today and (curr_date.month == 12 and curr_date.day in [24, 31]):
                    monthly[doc]['super_holidays'] += 1
                    if commit_to_history: lifetime[doc]['super_holidays'] += 1
                if len(day_shifts_worked) > 0 and ((curr_date.month == 12 and curr_date.day == 25) or \
                   (curr_date.month == 1 and curr_date.day == 1) or (curr_date.month == 4 and curr_date.day == 25) or \
                   (curr_date.month == 6 and curr_date.day in [1, 2]) or (curr_date.month == 8 and curr_date.day == 15) or is_easter):
                    monthly[doc]['super_holidays'] += 1
                    if commit_to_history: lifetime[doc]['super_holidays'] += 1

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
        dashboard_header = ['Doctor', 'Total Monthly Shifts', 'Target Hours', 'Actual Monthly Hours', 'Difference (+/-)',
                            'Monthly Nights', 'Monthly On-Call (Rep)', 'Monthly Doubles', 'Monthly Saturdays', 'Monthly Sundays', 'Monthly Holidays', 'Monthly Super Hols', 'Monthly Golden Wknds',
                            'LIFETIME Nights', 'LIFETIME On-Call', 'LIFETIME Doubles', 'LIFETIME Saturdays', 'LIFETIME Sundays', 'LIFETIME Holidays', 'LIFETIME Super Hols', 'LIFETIME Golden Wknds']
        
        target_monthly_hours = int((num_days / 7) * 38)
        worksheet.write(row_cursor, 0, "DOCTOR STATISTICS DASHBOARD", header_format)
        row_cursor += 1
        
        for col_idx, h in enumerate(dashboard_header):
            worksheet.write(row_cursor, col_idx, h, header_format)
        row_cursor += 1
        
        for doc in doctors:
            doc_target = int((num_days / 7) * 20) if doc == 'GAUDENZI' else target_monthly_hours
            actual_hours = monthly[doc]['total_hours']
            difference = actual_hours - doc_target
            
            data_row = [
                doc, monthly[doc]['total_shifts'], doc_target, actual_hours, f"+{difference}" if difference > 0 else str(difference),
                monthly[doc]['nights'], monthly[doc]['reps'], monthly[doc]['doubles'], monthly[doc]['saturdays'], monthly[doc]['sundays'], monthly[doc]['holidays'], monthly[doc]['super_holidays'], monthly[doc]['golden_weekends'],
                lifetime[doc]['nights'], lifetime[doc]['reps'], lifetime[doc]['doubles'], lifetime[doc]['saturdays'], lifetime[doc]['sundays'], lifetime[doc]['holidays'], lifetime[doc]['super_holidays'], lifetime[doc]['golden_weekends']
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
            with open(counter_file, 'w') as f: json.dump(lifetime, f, indent=4)
            return True, output.getvalue(), overlap_warning, f"Schedule solved! Block: {roster_dates[0].strftime('%b %d')} to {roster_dates[-1].strftime('%b %d')}. (OFFICIAL: Stats Saved){debug_msg}"
        else:
            return True, output.getvalue(), overlap_warning, f"Schedule solved! Block: {roster_dates[0].strftime('%b %d')} to {roster_dates[-1].strftime('%b %d')}. (SIMULATION: Stats NOT Saved){debug_msg}"
            
    else:
        return False, None, overlap_warning, "Constraints are too tight. The algorithm cannot find a mathematically legal schedule. Try using Emergency Debug Mode."


# ==========================================
# 2. THE GRAPHICAL USER INTERFACE (GUI)
# ==========================================
st.set_page_config(page_title="Cardiology Scheduler", page_icon="🩺", layout="wide")
st.title("🩺 Cardiology Shift Scheduler")
st.markdown("Automated constraints-based roster generation.")

default_doctors = ["BURGAZZI", "CACCAMO", "CARDINALI", "CICCARELLI", "CICCHIRILLO", "FLORI", "FINIZIO", "NUCCI", "ROBERTI", "GAUDENZI"]

if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame({
        "Doctor": default_doctors,
        "OR Capable (Checkbox)": [True if d in ["CACCAMO", "NUCCI"] else False for d in default_doctors],
        "Ward Preferred (Checkbox)": [True if d in ["BURGAZZI", "ROBERTI"] else False for d in default_doctors],
        "Private Practice (Afternoon)": [""] * len(default_doctors),
        "Ferie (YYYY-MM-DD)": [""] * len(default_doctors),
        "Desiderate (YYYY-MM-DD)": [""] * len(default_doctors),
        "Leave Weeks (Type the Monday)": [""] * len(default_doctors)
    })

COLOR_PALETTE = {
    "White": "#FFFFFF", "Light Blue": "#CCEBFF", "Light Green": "#CCFFCC", 
    "Light Red": "#FFCCCC", "Light Yellow": "#FFFFCC", "Peach": "#FFE5CC", 
    "Lavender": "#E5CCFF", "Light Pink": "#FFCCFF", "Mint": "#CCFFEA", 
    "Light Grey": "#E0E0E0", "Light Orange": "#FFD699"
}

with st.sidebar:
    st.header("1. Time Period")
    selected_year = st.number_input("Year", min_value=2024, max_value=2050, value=datetime.date.today().year)
    selected_month = st.number_input("Month", min_value=1, max_value=12, value=datetime.date.today().month)
    
    st.header("2. Base Settings")
    or_days_input = st.multiselect("Operating Room (OR) Days", 
                                   [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1)])
    or_days_formatted = [d.strftime("%Y-%m-%d") for d in or_days_input]
    
    st.markdown("---")
    st.header("🛠️ Troubleshooter")
    debug_toggle = st.checkbox("🚨 Enable Emergency Debug Mode", help="Ignores all equity rules to force a schedule.")

# Calculate global roster dates for dynamic UI elements
ui_roster_dates = get_roster_dates(selected_year, selected_month)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Master Spreadsheet", "🔒 Forced Shifts (Grid)", "🏥 Outpatient Clinics", "🎨 Doctor Colors", "🚀 Generate Schedule"])

with tab1:
    st.subheader("Fast Input Grid (Ferie, Desiderate, Leave, Private Practice)")
    st.markdown("Add new rows by scrolling to the bottom. Type the new doctor's name, tick their skills, and set their days. **Names must be unique.**")
    
    edited_df = st.data_editor(
        st.session_state.df,
        num_rows="dynamic",
        column_config={
            "Doctor": st.column_config.TextColumn("Doctor Name", required=True),
            "Private Practice (Afternoon)": st.column_config.SelectboxColumn(
                "Private Practice", options=["", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
            )
        },
        hide_index=True, use_container_width=True
    )
    st.session_state.df = edited_df

current_doctors = [str(d).strip().upper() for d in edited_df["Doctor"].dropna().unique() if str(d).strip()]

with tab2:
    st.subheader("Visual Override Grid")
    st.markdown("Click on any empty cell to forcefully lock a specific doctor into that shift. The algorithm will build the rest of the schedule around your choices.")
    st.warning("⚠️ **Warning:** If you force a shift that breaks a hard rule (e.g. 12-hour rest, or overlapping with Ferie), the algorithm will fail.")
    
    static_shifts = ['WARD_AM', 'URG_AM', 'OR_AM', 'OUT_PACEMAKER_AM', 'OUT_DIMESSI_AM', 'OUT_SCOMPENSO_AM', 'WARD_PM', 'URG_PM', 'REP_DAY', 'REP_NIGHT', 'NIGHT']
    
    # Check if the year/month changed so we can wipe the manual grid memory
    grid_memory_key = f"manual_grid_memory_{selected_year}_{selected_month}"
    if st.session_state.get("current_ym") != f"{selected_year}-{selected_month}":
        st.session_state["current_ym"] = f"{selected_year}-{selected_month}"
        # Wipe old grids
        for k in list(st.session_state.keys()):
            if k.startswith("grid_week_"):
                del st.session_state[k]

    # Generate the weekly visual tables
    for w_idx, week_start_idx in enumerate(range(0, len(ui_roster_dates), 7)):
        st.markdown(f"#### Week {w_idx + 1}")
        week_dates = ui_roster_dates[week_start_idx : week_start_idx + 7]
        col_names = [d.strftime("%Y-%m-%d") for d in week_dates]
        
        state_key = f"grid_week_{w_idx}"
        
        # Initialize an empty grid for this week if it doesn't exist
        if state_key not in st.session_state:
            init_dict = {"Shift": static_shifts}
            for c in col_names:
                init_dict[c] = [""] * len(static_shifts)
            st.session_state[state_key] = pd.DataFrame(init_dict)
            
        # Configure the columns so dates display nicely but output the correct YYYY-MM-DD format
        col_config = {"Shift": st.column_config.TextColumn("Shift", disabled=True)}
        for d in week_dates:
            d_str = d.strftime("%Y-%m-%d")
            display_name = f"{d.strftime('%b %d')} ({calendar.day_abbr[d.weekday()]})"
            col_config[d_str] = st.column_config.SelectboxColumn(
                display_name, options=[""] + current_doctors
            )
            
        edited_week_df = st.data_editor(
            st.session_state[state_key], 
            column_config=col_config, 
            hide_index=True, 
            key=f"editor_{state_key}",
            use_container_width=True
        )
        st.session_state[state_key] = edited_week_df
        st.markdown("---")

with tab3:
    st.subheader("Outpatient Setup")
    outpatient_setup = {}
    clinics = ['PACEMAKER', 'DIMESSI', 'SCOMPENSO']
    cols = st.columns(3)
    
    for i, clinic in enumerate(clinics):
        with cols[i]:
            st.markdown(f"#### {clinic}")
            dates = st.multiselect("Dates", [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1)], key=f"out_date_{clinic}")
            capable = st.multiselect("Capable Doctors", current_doctors, key=f"out_doc_{clinic}")
            outpatient_setup[clinic] = {'days': [d.strftime("%Y-%m-%d") for d in dates], 'capable': capable}

with tab4:
    st.subheader("Doctor Color Assignments")
    st.markdown("Choose a background color for each doctor's cells in the Excel output.")
    doc_colors_ui = {}
    
    default_palette_keys = list(COLOR_PALETTE.keys())
    
    col1, col2 = st.columns(2)
    for idx, doc in enumerate(current_doctors):
        with col1 if idx % 2 == 0 else col2:
            default_c = default_palette_keys[(idx % (len(default_palette_keys)-1)) + 1] 
            choice = st.selectbox(doc, list(COLOR_PALETTE.keys()), index=list(COLOR_PALETTE.keys()).index(default_c))
            doc_colors_ui[doc] = COLOR_PALETTE[choice]

with tab5:
    st.subheader("Generate File")
    st.info("💡 **Best Practice:** Run simulations until you are happy with the result. Only click 'Approve & Publish' when the schedule is final to avoid inflating the historical data.")
    
    col_sim, col_pub = st.columns(2)
    
    with col_sim:
        simulate_clicked = st.button("🛠️ SIMULATE SCHEDULE (No Save)", use_container_width=True)
    with col_pub:
        publish_clicked = st.button("🚀 APPROVE & PUBLISH (Save Stats)", use_container_width=True, type="primary")
    
    if simulate_clicked or publish_clicked:
        is_commit = publish_clicked
        
        if len(current_doctors) == 0:
            st.error("❌ Please add at least one doctor to the Master Spreadsheet.")
        else:
            with st.spinner("Calculating mathematical constraints... This may take up to 45 seconds."):
                
                ward_preferred = edited_df[edited_df["Ward Preferred (Checkbox)"] == True]["Doctor"].dropna().str.strip().str.upper().tolist()
                or_capable = edited_df[edited_df["OR Capable (Checkbox)"] == True]["Doctor"].dropna().str.strip().str.upper().tolist()
                
                # --- PARSE THE VISUAL WEEKLY GRIDS INTO FORCED SHIFTS ---
                manual_shifts = []
                for w_idx in range(0, len(ui_roster_dates) // 7):
                    state_key = f"grid_week_{w_idx}"
                    if state_key in st.session_state:
                        df_w = st.session_state[state_key]
                        for idx, row in df_w.iterrows():
                            shift_val = str(row["Shift"]).strip().upper()
                            # Loop over the actual date columns
                            for d in ui_roster_dates[w_idx*7 : (w_idx*7)+7]:
                                d_str = d.strftime("%Y-%m-%d")
                                doc = row.get(d_str, "")
                                if pd.notna(doc) and str(doc).strip():
                                    manual_shifts.append((str(doc).strip().upper(), d_str, shift_val))

                manual_festivities = []
                pp_dict = {}
                ferie_dict = {}
                desiderate_dict = {}
                leave_list = []

                for index, row in edited_df.iterrows():
                    doc = str(row["Doctor"]).strip().upper()
                    if not doc: continue
                    
                    if pd.notna(row["Private Practice (Afternoon)"]) and row["Private Practice (Afternoon)"]:
                        pp_dict[doc] = row["Private Practice (Afternoon)"]
                    if pd.notna(row["Ferie (YYYY-MM-DD)"]) and row["Ferie (YYYY-MM-DD)"]:
                        dates = [d.strip() for d in str(row["Ferie (YYYY-MM-DD)"]).split(",") if d.strip()]
                        if dates: ferie_dict[doc] = dates
                    if pd.notna(row["Desiderate (YYYY-MM-DD)"]) and row["Desiderate (YYYY-MM-DD)"]:
                        dates = [d.strip() for d in str(row["Desiderate (YYYY-MM-DD)"]).split(",") if d.strip()]
                        if dates: desiderate_dict[doc] = dates
                    if pd.notna(row["Leave Weeks (Type the Monday)"]) and row["Leave Weeks (Type the Monday)"]:
                        dates = [d.strip() for d in str(row["Leave Weeks (Type the Monday)"]).split(",") if d.strip()]
                        for d_str in dates: leave_list.append((doc, d_str))

                success, excel_data, warnings, message = generate_cardiology_schedule(
                    year=selected_year, month=selected_month, conditional_or_days=or_days_formatted, 
                    manual_festivities=manual_festivities, manual_assignments=manual_shifts,
                    ward_preferred_doctors=ward_preferred, or_capable_doctors=or_capable, outpatient_configs=outpatient_setup, 
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
                    
                    if not debug_toggle:
                        st.info("💡 **Tip:** Try checking 'Enable Emergency Debug Mode' in the left sidebar and running the simulation again. If it works, it means your vacations are too heavily grouped to allow for fair shift distribution.")
                    else:
                        st.info("💡 **Tip:** Even with Emergency Debug Mode enabled, the schedule failed. You likely have conflicting Forced Shifts, or too many doctors on vacation to cover the basic required shifts.")
