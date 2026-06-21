import streamlit as st
import csv
import calendar
import json
import os
import holidays
from collections import Counter
import datetime
from ortools.sat.python import cp_model

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
                                 ferie, leave_weeks, desiderate, private_practice_afternoons):
    doctors = ["BURGAZZI", "CACCAMO", "CARDINALI", "CICCARELLI", "CICCHIRILLO", 
               "FLORI", "FINIZIO", "NUCCI", "ROBERTI", "GAUDENZI"]
    
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
            lifetime[doc] = {'nights': 0, 'doubles': 0, 'saturdays': 0, 'sundays': 0, 'holidays': 0, 'super_holidays': 0, 'golden_weekends': 0}
        elif isinstance(lifetime[doc], int): 
            lifetime[doc] = {'nights': 0, 'doubles': 0, 'saturdays': 0, 'sundays': 0, 'holidays': lifetime[doc], 'super_holidays': 0, 'golden_weekends': 0}
        else:
            for key in ['nights', 'doubles', 'saturdays', 'sundays', 'holidays', 'super_holidays', 'golden_weekends']:
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
    overlap_warning = f"⚠️ Multiple doctors are off on: {overlaps}" if overlaps else ""

    shifts = ['WARD_AM', 'URG_AM', 'OR_AM', 'WARD_PM', 'URG_PM', 'NIGHT']
    for out_type in outpatient_configs.keys(): shifts.append(f'OUT_{out_type}_AM')
    
    model = cp_model.CpModel()
    work = {}
    for d in doctors:
        for day_idx in range(num_days):
            for s in shifts: work[(d, day_idx, s)] = model.NewBoolVar(f'work_{d}_{day_idx}_{s}')

    objective_terms = []

    for (d, date_str, s) in manual_assignments:
        if date_str in date_to_idx:
            model.Add(work[(d, date_to_idx[date_str], s)] == 1)

    for d, dates in ferie.items():
        for date_str in dates:
            if date_str in date_to_idx:
                for s in shifts: model.Add(work[(d, date_to_idx[date_str], s)] == 0)
                
    for d, dates in desiderate.items():
        for date_str in dates:
            if date_str in date_to_idx:
                for s in shifts: model.Add(work[(d, date_to_idx[date_str], s)] == 0)

    for d, start_monday_str in leave_weeks:
        if start_monday_str in date_to_idx:
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
        
        if day_idx not in sunday_equivalent_days: 
            model.AddExactlyOne(work[(d, day_idx, 'WARD_AM')] for d in doctors)
            model.AddExactlyOne(work[(d, day_idx, 'URG_AM')] for d in doctors)
            model.AddExactlyOne(work[(d, day_idx, 'WARD_PM')] for d in doctors)
            model.AddExactlyOne(work[(d, day_idx, 'URG_PM')] for d in doctors)
            
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
            for d in doctors:
                for s in shifts:
                    if s not in ['WARD_AM', 'WARD_PM', 'NIGHT']: model.Add(work[(d, day_idx, s)] == 0)

    day_name_to_num = {'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6}

    for d in doctors:
        pp_day_num = day_name_to_num.get(private_practice_afternoons.get(d, "").strip().capitalize(), -1)
        for day_idx in range(num_days):
            weekday = roster_dates[day_idx].weekday()
            day_shifts_only = [s for s in shifts if s != 'NIGHT']
            
            if weekday == pp_day_num:
                for s in shifts:
                    if s.endswith('_PM'): model.Add(work[(d, day_idx, s)] == 0)
            
            if d == 'FINIZIO':
                model.Add(work[(d, day_idx, 'NIGHT')] == 0)
                model.Add(sum(work[(d, day_idx, s)] for s in shifts) <= 1)
            elif d == 'GAUDENZI':
                model.Add(work[(d, day_idx, 'NIGHT')] == 0)
                if weekday == 4: 
                    for s in shifts:
                        if s not in ['WARD_PM', 'URG_PM']: model.Add(work[(d, day_idx, s)] == 0)
                elif weekday == 5: pass 
                else: 
                    for s in shifts: model.Add(work[(d, day_idx, s)] == 0)
            else:
                if day_idx < num_days - 1:
                    is_pre_night = work[(d, day_idx + 1, 'NIGHT')]
                    model.Add(sum(work[(d, day_idx, s)] for s in day_shifts_only) == 2).OnlyEnforceIf(is_pre_night)
                    if day_idx not in sunday_equivalent_days:
                        model.Add(sum(work[(d, day_idx, s)] for s in shifts) <= 1).OnlyEnforceIf(is_pre_night.Not())
                    else:
                        model.Add(sum(work[(d, day_idx, s)] for s in shifts) <= 2).OnlyEnforceIf(is_pre_night.Not())
                else:
                    if day_idx not in sunday_equivalent_days: model.Add(sum(work[(d, day_idx, s)] for s in shifts) <= 1)
                    else: model.Add(sum(work[(d, day_idx, s)] for s in shifts) <= 2)
            
            model.Add(sum(work[(d, day_idx, s)] for s in day_shifts_only) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
            if day_idx + 1 < num_days: model.Add(sum(work[(d, day_idx + 1, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])
            if day_idx + 2 < num_days: model.Add(sum(work[(d, day_idx + 2, s)] for s in shifts) == 0).OnlyEnforceIf(work[(d, day_idx, 'NIGHT')])

        doctor_gws = []
        for week_start_idx in range(0, num_days, 7):
            fri_idx, sat_idx, sun_idx = week_start_idx + 4, week_start_idx + 5, week_start_idx + 6
            gw_var = model.NewBoolVar(f'gw_{d}_{week_start_idx}')
            ruining_shifts = [work[(d, fri_idx, 'NIGHT')]]
            for s in shifts:
                ruining_shifts.extend([work[(d, sat_idx, s)], work[(d, sun_idx, s)]])
            model.Add(sum(ruining_shifts) == 0).OnlyEnforceIf(gw_var)
            model.Add(sum(ruining_shifts) > 0).OnlyEnforceIf(gw_var.Not())
            doctor_gws.append(gw_var)
            objective_terms.append(-5 * lifetime[d]['golden_weekends'] * gw_var)
        model.Add(sum(doctor_gws) >= 1)

    weeks = [range(i, i + 7) for i in range(0, num_days, 7)]
    for w in weeks:
        model.Add(sum((12 if s == 'NIGHT' else 6) * work[('GAUDENZI', day_idx, s)] for day_idx in w for s in shifts) <= 20)

    night_doctors = [d for d in doctors if d not in ['FINIZIO', 'GAUDENZI']]
    base_nights = num_days // len(night_doctors)
    
    holiday_doctors = [d for d in doctors if d != 'GAUDENZI']
    base_hols = (len(sunday_equivalent_days) * 2) // len(holiday_doctors)

    standard_doctors = [d for d in doctors if d != 'GAUDENZI']
    base_day_shifts = (num_days * 4) // len(standard_doctors)

    for d in night_doctors:
        model.Add(sum(work[(d, day_idx, 'NIGHT')] for day_idx in range(num_days)) >= max(0, base_nights - 2))
        model.Add(sum(work[(d, day_idx, 'NIGHT')] for day_idx in range(num_days)) <= base_nights + 2)
    for d in holiday_doctors:
        model.Add(sum(work[(d, day_idx, s)] for day_idx in sunday_equivalent_days for s in ['WARD_AM', 'WARD_PM']) >= 0)
        model.Add(sum(work[(d, day_idx, s)] for day_idx in sunday_equivalent_days for s in ['WARD_AM', 'WARD_PM']) <= base_hols + 3)
    for d in standard_doctors:
        day_shifts_only = [s for s in shifts if s != 'NIGHT']
        model.Add(sum(work[(d, day_idx, s)] for day_idx in range(num_days) for s in day_shifts_only) >= max(0, base_day_shifts - 4))
        model.Add(sum(work[(d, day_idx, s)] for day_idx in range(num_days) for s in day_shifts_only) <= base_day_shifts + 8)

    for d in ward_preferred_doctors:
        for day_idx in range(num_days):
            objective_terms.extend([5 * work[(d, day_idx, 'WARD_AM')], 5 * work[(d, day_idx, 'WARD_PM')]])

    for d in doctors:
        for day_idx in range(num_days - 1):
            am_pm = model.NewBoolVar('')
            model.AddBoolAnd([work[(d, day_idx, 'WARD_AM')], work[(d, day_idx+1, 'WARD_PM')]]).OnlyEnforceIf(am_pm)
            objective_terms.append(5 * am_pm)
            pm_am = model.NewBoolVar('')
            model.AddBoolAnd([work[(d, day_idx, 'WARD_PM')], work[(d, day_idx+1, 'WARD_AM')]]).OnlyEnforceIf(pm_am)
            objective_terms.append(5 * pm_am)
            
        for w_idx in range(len(weeks) - 1):
            w_curr = model.NewBoolVar('')
            w_next = model.NewBoolVar('')
            model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(w_curr)
            model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(w_curr.Not())
            model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) > 0).OnlyEnforceIf(w_next)
            model.Add(sum(work[(d, day_idx, s)] for day_idx in weeks[w_idx+1] for s in ['WARD_AM', 'WARD_PM']) == 0).OnlyEnforceIf(w_next.Not())
            cons = model.NewBoolVar('')
            model.AddBoolAnd([w_curr, w_next]).OnlyEnforceIf(cons)
            objective_terms.append(-20 * cons)

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
            day_shifts_only = [s for s in shifts if s != 'NIGHT']
            model.Add(sum(work[(d, day_idx, s)] for s in day_shifts_only) > 0).OnlyEnforceIf(worked_day_shifts)
            model.Add(sum(work[(d, day_idx, s)] for s in day_shifts_only) == 0).OnlyEnforceIf(worked_day_shifts.Not())
            
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
        monthly = {doc: {'nights': 0, 'doubles': 0, 'saturdays': 0, 'sundays': 0, 'holidays': 0, 'super_holidays': 0, 'golden_weekends': 0, 'total_shifts': 0, 'total_hours': 0} for doc in doctors}
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
                    monthly[doc]['total_hours'] += 12 if s == 'NIGHT' else 6
                
                if 'NIGHT' in worked_today:
                    monthly[doc]['nights'] += 1
                    lifetime[doc]['nights'] += 1
                    
                day_shifts_worked = [s for s in worked_today if s != 'NIGHT']
                if len(day_shifts_worked) >= 2:
                    monthly[doc]['doubles'] += 1
                    lifetime[doc]['doubles'] += 1
                    
                if worked_today:
                    if curr_date.weekday() == 5:
                        lifetime[doc]['saturdays'] += 1; monthly[doc]['saturdays'] += 1
                    elif curr_date.weekday() == 6:
                        lifetime[doc]['sundays'] += 1; monthly[doc]['sundays'] += 1
                    if curr_date in it_holidays or curr_date.strftime("%Y-%m-%d") in manual_festivities:
                        lifetime[doc]['holidays'] += 1; monthly[doc]['holidays'] += 1

                is_easter = (it_holidays.get(curr_date) == "Pasqua di Resurrezione")
                if 'NIGHT' in worked_today and (curr_date.month == 12 and curr_date.day in [24, 31]):
                    lifetime[doc]['super_holidays'] += 1; monthly[doc]['super_holidays'] += 1
                if len(day_shifts_worked) > 0 and ((curr_date.month == 12 and curr_date.day == 25) or \
                   (curr_date.month == 1 and curr_date.day == 1) or (curr_date.month == 4 and curr_date.day == 25) or \
                   (curr_date.month == 6 and curr_date.day in [1, 2]) or (curr_date.month == 8 and curr_date.day == 15) or is_easter):
                    lifetime[doc]['super_holidays'] += 1; monthly[doc]['super_holidays'] += 1

            for week_start_idx in range(0, num_days, 7):
                if not ('NIGHT' in doc_schedule[doc][week_start_idx + 4]) and not doc_schedule[doc][week_start_idx + 5] and not doc_schedule[doc][week_start_idx + 6]:
                    lifetime[doc]['golden_weekends'] += 1; monthly[doc]['golden_weekends'] += 1

        timestamp = datetime.datetime.now().strftime("%H%M%S")
        filename = f'cardiology_schedule_{year}_{month}_{timestamp}.csv'
        target_monthly_hours = int((num_days / 7) * 38)
        
        with open(filename, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Shift'] + [f"{roster_dates[day_idx].strftime('%b %d')} ({calendar.day_abbr[roster_dates[day_idx].weekday()]}){'*' if day_idx in sunday_equivalent_days else ''}" for day_idx in range(num_days)])
            for s in shifts:
                row = [s]
                for day_idx in range(num_days):
                    assigned = ""
                    for doc in doctors:
                        if solver.Value(work[(doc, day_idx, s)]) == 1: assigned = doc
                    row.append(assigned)
                writer.writerow(row)
                
            writer.writerows([[], [], ['DOCTOR STATISTICS DASHBOARD']])
            writer.writerow(['Doctor', 'Total Monthly Shifts', f'Target Hours ({target_monthly_hours}h max)', 'Actual Monthly Hours', 'Difference (+/-)',
                             'Monthly Nights', 'Monthly Doubles', 'Monthly Saturdays', 'Monthly Sundays', 'Monthly Holidays', 'Monthly Super Hols', 'Monthly Golden Wknds',
                             'LIFETIME Nights', 'LIFETIME Doubles', 'LIFETIME Saturdays', 'LIFETIME Sundays', 'LIFETIME Holidays', 'LIFETIME Super Hols', 'LIFETIME Golden Wknds'])
            for doc in doctors:
                doc_target = int((num_days / 7) * 20) if doc == 'GAUDENZI' else target_monthly_hours
                actual_hours = monthly[doc]['total_hours']
                difference = actual_hours - doc_target
                writer.writerow([doc, monthly[doc]['total_shifts'], doc_target, actual_hours, f"+{difference}" if difference > 0 else str(difference),
                                 monthly[doc]['nights'], monthly[doc]['doubles'], monthly[doc]['saturdays'], monthly[doc]['sundays'], monthly[doc]['holidays'], monthly[doc]['super_holidays'], monthly[doc]['golden_weekends'],
                                 lifetime[doc]['nights'], lifetime[doc]['doubles'], lifetime[doc]['saturdays'], lifetime[doc]['sundays'], lifetime[doc]['holidays'], lifetime[doc]['super_holidays'], lifetime[doc]['golden_weekends']])

        with open(counter_file, 'w') as f: json.dump(lifetime, f, indent=4)
        return True, filename, overlap_warning, f"Schedule solved! Block: {roster_dates[0].strftime('%b %d')} to {roster_dates[-1].strftime('%b %d')}."
    else:
        return False, "", overlap_warning, "Constraints are too tight. Check if too many doctors are on leave."


# ==========================================
# 2. THE GRAPHICAL USER INTERFACE (GUI)
# ==========================================
st.set_page_config(page_title="Cardiology Scheduler", page_icon="🩺", layout="wide")
st.title("🩺 Cardiology Shift Scheduler")
st.markdown("Automated constraints-based roaster generation. Select your parameters and click **Generate**.")

doctors_list = ["BURGAZZI", "CACCAMO", "CARDINALI", "CICCARELLI", "CICCHIRILLO", "FLORI", "FINIZIO", "NUCCI", "ROBERTI", "GAUDENZI"]

# --- SIDEBAR SETTINGS ---
with st.sidebar:
    st.header("1. Time Period")
    selected_year = st.number_input("Year", min_value=2024, max_value=2050, value=datetime.date.today().year)
    selected_month = st.number_input("Month", min_value=1, max_value=12, value=datetime.date.today().month)
    
    st.header("2. Base Settings")
    or_days_input = st.multiselect("Operating Room (OR) Days", 
                                   [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1)])
    or_days_formatted = [d.strftime("%Y-%m-%d") for d in or_days_input]
    
    st.markdown("---")
    st.markdown("### Private Practice Afternoons")
    pp_dict = {}
    for d in doctors_list:
        if d not in ['FINIZIO', 'GAUDENZI']: # Exclude restricted doctors
            day = st.selectbox(f"{d}", ["", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"], key=f"pp_{d}")
            if day: pp_dict[d] = day

# --- MAIN TABS ---
tab1, tab2, tab3 = st.tabs(["🗓️ Time Off (Ferie & Desiderate)", "🏥 Outpatient Clinics", "🚀 Generate Schedule"])

with tab1:
    st.subheader("Manage Vacations and Requested Days Off")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Ferie & Desiderate (Specific Days)")
        ferie_dict = {}
        desiderate_dict = {}
        for doc in doctors_list:
            with st.expander(f"🌴 {doc} Days Off"):
                f_dates = st.multiselect("Ferie (Official)", [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1)], key=f"ferie_{doc}")
                d_dates = st.multiselect("Desiderate (Requested)", [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1)], key=f"des_{doc}")
                if f_dates: ferie_dict[doc] = [d.strftime("%Y-%m-%d") for d in f_dates]
                if d_dates: desiderate_dict[doc] = [d.strftime("%Y-%m-%d") for d in d_dates]

    with col2:
        st.markdown("### Full Leave Weeks")
        st.info("Select the **Monday** that starts the week of leave.")
        leave_list = []
        for doc in doctors_list:
            mondays = [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1) if datetime.date(selected_year, selected_month, d).weekday() == 0]
            leave_mons = st.multiselect(f"{doc} Weeks", mondays, key=f"leave_{doc}")
            for mon in leave_mons:
                leave_list.append((doc, mon.strftime("%Y-%m-%d")))

with tab2:
    st.subheader("Outpatient Setup")
    st.markdown("Assign dates and qualified doctors for each clinic.")
    
    outpatient_setup = {}
    clinics = ['PACEMAKER', 'DIMESSI', 'SCOMPENSO']
    cols = st.columns(3)
    
    for i, clinic in enumerate(clinics):
        with cols[i]:
            st.markdown(f"#### {clinic}")
            dates = st.multiselect("Dates", [datetime.date(selected_year, selected_month, d) for d in range(1, calendar.monthrange(selected_year, selected_month)[1] + 1)], key=f"out_date_{clinic}")
            capable = st.multiselect("Capable Doctors", doctors_list, key=f"out_doc_{clinic}")
            outpatient_setup[clinic] = {
                'days': [d.strftime("%Y-%m-%d") for d in dates],
                'capable': capable
            }

with tab3:
    st.subheader("Generate File")
    st.markdown("Review your settings in the tabs above, then click the button below to calculate.")
    
    if st.button("⚙️ GENERATE SCHEDULE", use_container_width=True, type="primary"):
        with st.spinner("Calculating mathematical constraints... This may take up to 45 seconds."):
            
            # Static inputs that rarely change
            ward_preferred = ["BURGAZZI", "ROBERTI"]
            or_capable = ["CACCAMO", "NUCCI"]
            manual_festivities = []
            manual_shifts = []

            # Run Engine
            success, filename, warnings, message = generate_cardiology_schedule(
                year=selected_year, month=selected_month, 
                conditional_or_days=or_days_formatted, 
                manual_festivities=manual_festivities, manual_assignments=manual_shifts,
                ward_preferred=ward_preferred, or_capable=or_capable, 
                outpatient_configs=outpatient_setup, 
                ferie=ferie_dict, leave_weeks=leave_list, desiderate=desiderate_dict,
                private_practice_afternoons=pp_dict
            )
            
            if success:
                st.success(f"✅ {message}")
                if warnings:
                    st.warning(warnings)
                
                # Provide Download Button
                with open(filename, 'rb') as f:
                    st.download_button(
                        label="📥 Download Excel/CSV Schedule",
                        data=f,
                        file_name=filename,
                        mime="text/csv",
                        type="primary"
                    )
            else:
                st.error(f"❌ {message}")
                if warnings:
                    st.warning(warnings)
                st.info("Try removing a few 'Desiderate' or 'Ferie' requests to give the algorithm more breathing room.")