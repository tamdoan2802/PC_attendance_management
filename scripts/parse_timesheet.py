# -*- coding: utf-8 -*-
"""
Timesheet Parser Module
Parses MISA Timesheet Excel files ('Bảng chấm công chi tiết'), handling multiline cells,
shift quotas, check-in/out timestamps, leave markers, and 15 raw data defects.
"""

import os
import sys
import re
import glob
import datetime
import pandas as pd

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

try:
    from .config import (
        TIMESHEET_DIR, SHIFT_CATALOG, DEPRECATED_SHIFTS, EXCLUDED_POPULATION_PREFIXES,
        EXCLUDED_EMP_IDS, CHECKIN_GRACE_MINUTES, HOLIDAY_FILE
    )
except (ImportError, ValueError):
    from config import (
        TIMESHEET_DIR, SHIFT_CATALOG, DEPRECATED_SHIFTS, EXCLUDED_POPULATION_PREFIXES,
        EXCLUDED_EMP_IDS, CHECKIN_GRACE_MINUTES, HOLIDAY_FILE
    )

def load_holidays(holiday_file=HOLIDAY_FILE):
    """Loads public holidays from CSV into a dict of {datetime.date: holiday_name}."""
    holidays = {}
    if not os.path.exists(holiday_file):
        return holidays
    try:
        df = pd.read_csv(holiday_file, encoding='utf-8')
        for _, row in df.iterrows():
            d_str = str(row.get('Date', '')).strip()
            name = str(row.get('Holiday_Name', '')).strip()
            if not d_str:
                continue
            # format DD/MM/YYYY
            try:
                parts = d_str.split('/')
                if len(parts) == 3:
                    d = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
                    holidays[d] = name
            except Exception:
                pass
    except Exception as e:
        print(f"Warning loading holidays: {e}")
    return holidays

def extract_dates_from_filename(filename):
    """
    Extracts start_date and end_date from filename:
    'Bảng chấm công chi tiết_Bảng chấm công từ ngày 21_08_2026 đến ngày 20_09_2026.xlsx'
    Returns (start_date, end_date) or (None, None)
    """
    pattern = r"ngày\s+(\d{1,2})_(\d{1,2})_(\d{4})\s+đến\s+ngày\s+(\d{1,2})_(\d{1,2})_(\d{4})"
    match = re.search(pattern, filename)
    if match:
        d1, m1, y1, d2, m2, y2 = match.groups()
        try:
            start_date = datetime.date(int(y1), int(m1), int(d1))
            end_date = datetime.date(int(y2), int(m2), int(d2))
            return start_date, end_date
        except Exception:
            pass
    return None, None

def get_timesheet_files_for_range(start_date, end_date, timesheet_dir=TIMESHEET_DIR):
    """Finds all timesheet files that overlap with [start_date, end_date]."""
    files = glob.glob(os.path.join(timesheet_dir, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    
    matched = []
    for f in files:
        f_start, f_end = extract_dates_from_filename(os.path.basename(f))
        if f_start and f_end:
            # check overlap
            if not (end_date < f_start or start_date > f_end):
                matched.append((f_start, f_end, f))
                
    # Sort by f_start
    matched.sort(key=lambda x: x[0])
    return [m[2] for m in matched]

def parse_cell(cell_value):
    """
    Parses a single attendance cell:
      Line 1: {SHIFT_CODE}: {ASSIGNED_CREDIT}
      Line 2: {CHECK_IN} - {CHECK_OUT}
      Line 3: {LEAVE_TYPE / SYSTEM_NOTE}: {LEAVE_CREDIT}
      Line 4: additional info
    Returns a dict with parsed values.
    """
    res = {
        'shift_code': None,
        'shift_credit': 0.0,
        'check_in': None,
        'check_out': None,
        'has_ci': False,
        'has_co': False,
        'leave_type': None,
        'leave_credit': 0.0,
        'is_manual_override': False,
        'raw_text': str(cell_value) if pd.notna(cell_value) else ''
    }
    
    if pd.isna(cell_value):
        return res
        
    text = str(cell_value).strip()
    if not text:
        return res
        
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    if not lines:
        return res
        
    # Check for manual override "Cập nhật công" anywhere in cell
    if 'Cập nhật công' in text or 'Cap nhat cong' in text:
        res['is_manual_override'] = True
        
    # Line 1: Shift and credit, e.g. "BS01: 1", "WT: 0.5", "WE: 1"
    # Note: sometimes Line 1 might just be shift code
    l1 = lines[0]
    m_shift = re.match(r"^([A-Za-z0-9\-]+)(?::\s*([\d\.]+))?", l1)
    if m_shift:
        res['shift_code'] = m_shift.group(1).strip()
        if m_shift.group(2):
            try:
                res['shift_credit'] = float(m_shift.group(2))
            except ValueError:
                res['shift_credit'] = 1.0
        else:
            res['shift_credit'] = 1.0
            
    # Line 2: Check-in / Check-out, e.g. "07:27 - 17:04" or "-:- - -:-" or "08:00 - -:-"
    if len(lines) >= 2:
        l2 = lines[1]
        m_time = re.match(r"^(\d{1,2}:\d{2}|-:-)\s*-\s*(\d{1,2}:\d{2}|-:-)", l2)
        if m_time:
            ci = m_time.group(1)
            co = m_time.group(2)
            if ci != '-:-':
                res['check_in'] = ci
                res['has_ci'] = True
            if co != '-:-':
                res['check_out'] = co
                res['has_co'] = True
                
    # Lines 3 & 4: Leave info, manual override, or secondary shifts
    for line in lines[2:]:
        if 'Cập nhật công' in line:
            res['is_manual_override'] = True
            continue
        m_leave = re.match(r"^([^:]+):\s*([\d\.]+)", line)
        if m_leave:
            l_name = m_leave.group(1).strip()
            l_cred = 0.0
            try:
                l_cred = float(m_leave.group(2))
            except ValueError:
                l_cred = 0.5
            # Distinguish leave vs business trip vs second shift
            if any(k in l_name for k in ['Nghỉ', 'Nghi', 'BHXH', 'phép', 'bù']):
                res['leave_type'] = l_name
                res['leave_credit'] = l_cred
            elif any(k in l_name.lower() for k in ['công tác', 'cong tac']):
                res['mission_credit'] = l_cred
                res['is_mission'] = True
            elif l_name in SHIFT_CATALOG:
                # Split shift additional session
                pass
                
    # Defect 5: Double-count fix on half-day leave
    # If leave_credit > 0 and shift_credit == 1.0, real work credit is max(0.0, 1.0 - leave_credit)
    if res['leave_credit'] > 0 and res['shift_credit'] > 0:
        res['shift_credit'] = max(0.0, 1.0 - res['leave_credit'])
        
    return res

def calculate_time_metrics(ci_str, co_str, shift_code, is_holiday=False, holiday_name=''):
    """
    Calculates actual hours, late check-in minutes, and late check-out minutes.
    Uses CHECKIN_GRACE_MINUTES (5 minutes).
    """
    metrics = {
        'actual_hours': 0.0,
        'std_hours': 8.0,
        'early_ci_mins': 0.0,
        'late_ci_mins': 0.0,
        'early_co_mins': 0.0,
        'late_co_mins': 0.0,
        'shift_start': None,
        'shift_end': None,
        'is_early_ci': False,
        'is_late_ci': False,
        'is_early_co': False,
        'is_late_co': False,
        'is_chronic_late_ci': False
    }
    
    if is_holiday:
        metrics['std_hours'] = 0.0
        return metrics
        
    shift_info = SHIFT_CATALOG.get(shift_code)
    if not shift_info:
        # Default standard shift WT
        shift_info = SHIFT_CATALOG.get('WT', ('08:00', '17:00', 1.0, 8.0))
        
    s_start_str, s_end_str, break_hours, std_hours = shift_info
    metrics['std_hours'] = std_hours
    metrics['shift_start'] = s_start_str
    metrics['shift_end'] = s_end_str
    
    s_start_m = int(s_start_str.split(':')[0]) * 60 + int(s_start_str.split(':')[1])
    s_end_m = int(s_end_str.split(':')[0]) * 60 + int(s_end_str.split(':')[1])
    
    ci_m = None
    co_m = None
    if ci_str:
        ci_m = int(ci_str.split(':')[0]) * 60 + int(ci_str.split(':')[1])
    if co_str:
        co_m = int(co_str.split(':')[0]) * 60 + int(co_str.split(':')[1])
        
    # Check-in metrics
    if ci_m is not None:
        if ci_m < s_start_m:
            metrics['early_ci_mins'] = float(s_start_m - ci_m)
            metrics['is_early_ci'] = True
        else:
            raw_late = ci_m - s_start_m
            if raw_late > CHECKIN_GRACE_MINUTES:
                metrics['late_ci_mins'] = float(raw_late)
                metrics['is_late_ci'] = True
                if raw_late > 30:
                    metrics['is_chronic_late_ci'] = True
                
    # Check-out metrics
    if co_m is not None:
        if co_m < s_end_m:
            metrics['early_co_mins'] = float(s_end_m - co_m)
            metrics['is_early_co'] = True
        elif co_m > s_end_m:
            metrics['late_co_mins'] = float(co_m - s_end_m)
            metrics['is_late_co'] = True
            
    # Actual working hours calculation
    if ci_m is not None and co_m is not None and co_m > ci_m:
        total_mins = co_m - ci_m
        total_hours = total_mins / 60.0
        # Deduct break if worked over 5 hours and break applies
        if total_hours >= 5.0 and break_hours > 0:
            actual_hours = max(0.0, total_hours - break_hours)
        else:
            actual_hours = total_hours
        metrics['actual_hours'] = round(actual_hours, 2)
        
    return metrics

def parse_timesheet_file(file_path, target_start_date=None, target_end_date=None, holidays=None):
    """
    Parses a single Timesheet Excel file into a list of daily record dicts.
    Filters records within [target_start_date, target_end_date] if provided.
    """
    if holidays is None:
        holidays = load_holidays()
        
    f_start, f_end = extract_dates_from_filename(os.path.basename(file_path))
    if not f_start or not f_end:
        print(f"Warning: Could not parse dates from {os.path.basename(file_path)}")
        return []
        
    df = pd.read_excel(file_path, header=None)
    if df.shape[0] < 13:
        return []
        
    # Map column index to date
    # Row 10 has day numbers, Row 11 has DOW
    col_dates = {}
    for c in range(14, df.shape[1]):
        val = df.iloc[10, c]
        if pd.notna(val):
            try:
                day_num = int(str(val).strip())
                # If day_num >= 21: it belongs to f_start month/year
                # If day_num < 21: it belongs to f_end month/year
                if day_num >= 21:
                    d = datetime.date(f_start.year, f_start.month, day_num)
                else:
                    d = datetime.date(f_end.year, f_end.month, day_num)
                col_dates[c] = d
            except Exception:
                pass
                
    records = []
    for r in range(12, df.shape[0]):
        emp_id = str(df.iloc[r, 1]).strip() if pd.notna(df.iloc[r, 1]) else ''
        emp_name = str(df.iloc[r, 2]).strip() if pd.notna(df.iloc[r, 2]) else ''
        position = str(df.iloc[r, 3]).strip() if pd.notna(df.iloc[r, 3]) else ''
        dept = str(df.iloc[r, 4]).strip() if pd.notna(df.iloc[r, 4]) else ''
        
        if not emp_id or emp_id == 'nan':
            continue
            
        # Defect 14: Exclude interns and excluded personnel (Adrian, Chau Ha, Toan Mai, Huyen Le)
        if any(emp_id.startswith(pfx) for pfx in EXCLUDED_POPULATION_PREFIXES) or emp_id in EXCLUDED_EMP_IDS:
            continue
            
        for c, date_val in col_dates.items():
            if target_start_date and date_val < target_start_date:
                continue
            if target_end_date and date_val > target_end_date:
                continue
                
            cell_val = df.iloc[r, c]
            parsed = parse_cell(cell_val)
            
            # Defect 13: Exclude deprecated WL01
            if parsed['shift_code'] in DEPRECATED_SHIFTS:
                continue
                
            is_weekend = date_val.weekday() >= 5
            is_holiday = date_val in holidays
            holiday_name = holidays.get(date_val, '')
            
            # Defect 8: Weekday NaN means pre-hire or post-resignation
            is_empty = pd.isna(cell_val) or str(cell_val).strip() == ''
            if is_empty and is_weekend:
                continue # Normal weekend
                
            metrics = calculate_time_metrics(
                parsed['check_in'], parsed['check_out'], parsed['shift_code'],
                is_holiday=is_holiday, holiday_name=holiday_name
            )
            
            # Credit business trip working hours (Đi công tác)
            if parsed.get('mission_credit', 0) > 0:
                metrics['actual_hours'] = max(metrics['actual_hours'], parsed['mission_credit'] * metrics['std_hours'])
            
            # Type of date classification
            work_credit = parsed['shift_credit']
            if is_holiday:
                type_of_date = 'Holiday'
                working_credit = 1.0
            elif parsed.get('mission_credit', 0) >= 1.0:
                type_of_date = 'FullWorkDay'
                working_credit = 1.0
            elif parsed.get('mission_credit', 0) > 0:
                type_of_date = 'HalfWorkDay'
                working_credit = 1.0
            elif parsed['leave_credit'] >= 1.0:
                type_of_date = 'FullLeave'
                working_credit = 1.0
            elif parsed['leave_credit'] == 0.5:
                type_of_date = 'HalfWorkDay'
                working_credit = 1.0
            elif work_credit >= 1.0:
                type_of_date = 'FullWorkDay'
                working_credit = 1.0
            elif work_credit > 0:
                type_of_date = 'HalfWorkDay'
                working_credit = work_credit
            elif is_empty:
                type_of_date = 'NoContract'
                working_credit = 0.0
            else:
                type_of_date = 'Absent'
                working_credit = 0.0
                
            # Actual presence day (0.0, 0.5, 1.0)
            actual_day = 0.0
            if type_of_date == 'FullWorkDay':
                actual_day = 1.0
            elif type_of_date == 'HalfWorkDay':
                actual_day = 0.5
                
            rec = {
                'emp_id': emp_id,
                'emp_name': emp_name,
                'position': position,
                'department': dept,
                'date': date_val,
                'day_of_week': date_val.strftime('%a'),
                'is_weekend': is_weekend,
                'is_holiday': is_holiday,
                'holiday_name': holiday_name,
                'type_of_date': type_of_date,
                'shift_code': parsed['shift_code'],
                'work_credit': work_credit,
                'actual_day': actual_day,
                'working_credit': working_credit,
                'check_in': parsed['check_in'],
                'check_out': parsed['check_out'],
                'has_ci': parsed['has_ci'],
                'has_co': parsed['has_co'],
                'leave_type': parsed['leave_type'],
                'leave_credit': parsed['leave_credit'],
                'is_manual_override': parsed['is_manual_override'],
                'shift_start': metrics['shift_start'],
                'shift_end': metrics['shift_end'],
                'std_hours': metrics['std_hours'],
                'actual_hours': metrics['actual_hours'],
                'early_ci_mins': metrics['early_ci_mins'],
                'late_ci_mins': metrics['late_ci_mins'],
                'early_co_mins': metrics['early_co_mins'],
                'late_co_mins': metrics['late_co_mins'],
                'is_early_ci': metrics['is_early_ci'],
                'is_late_ci': metrics['is_late_ci'],
                'is_chronic_late_ci': metrics['is_chronic_late_ci'],
                'is_early_co': metrics['is_early_co'],
                'is_late_co': metrics['is_late_co'],
                'raw_cell': parsed['raw_text']
            }
            records.append(rec)
            
    return records

def load_attendance_dataset(start_date, end_date):
    """
    Loads and stitches attendance records across all relevant Timesheet files
    for the exact date range [start_date, end_date].
    Returns a pandas DataFrame.
    """
    files = get_timesheet_files_for_range(start_date, end_date)
    if not files:
        print(f"No timesheet files found covering {start_date} to {end_date}")
        return pd.DataFrame()
        
    holidays = load_holidays()
    all_records = []
    seen_keys = set()
    
    for f in files:
        recs = parse_timesheet_file(f, target_start_date=start_date, target_end_date=end_date, holidays=holidays)
        for r in recs:
            key = (r['emp_id'], r['date'])
            if key not in seen_keys:
                seen_keys.add(key)
                all_records.append(r)
                
    df = pd.DataFrame(all_records)
    if not df.empty:
        df['date'] = pd.to_datetime(df['date']).dt.date
        df = df.sort_values(by=['emp_id', 'date']).reset_index(drop=True)
    return df
