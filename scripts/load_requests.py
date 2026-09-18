# -*- coding: utf-8 -*-
"""
Requests and Master Data Loader Module
Loads, parses, and reconciles all 6 MISA request files and Master Data entities.
Handles recurring requests expansion, date inversion, OT coalesce fallback,
and multi-day leave slicing.
"""

import os
import re
import datetime
import pandas as pd
from .config import (
    REQ_LATE_EARLY, REQ_OVERTIME, REQ_WFH, REQ_SHIFT_CHANGE,
    REQ_BUSINESS_TRIP, REQ_LEAVE, ROUTING_MAP_FILE, EMPLOYEES_FILE,
    SHIFT_LABEL_MAP, NOTICE_SUDDEN_MAX_DAYS, NOTICE_SHORT_MAX_DAYS, NOTICE_PROMPT_MAX_DAYS,
    EXCLUDED_EMP_IDS, EXCLUDED_TOTAL
)

def load_master_entities(routing_map_file=ROUTING_MAP_FILE):
    """
    Parses entities/notification_routing_map.md into a DataFrame indexed by EmpID.
    Returns DataFrame with columns:
    [emp_id, full_name_en, emp_name_vn, email, team, tl_name, tl_email,
     senior_leader, sl_email, customer_group, contact_names, contact_emails]
    """
    if not os.path.exists(routing_map_file):
        print(f"Warning: routing map file not found: {routing_map_file}")
        return pd.DataFrame()
        
    rows = []
    headers = None
    with open(routing_map_file, 'r', encoding='utf-8') as f:
        for line in f:
            line_str = line.strip()
            if not line_str.startswith('|'):
                continue
            parts = [p.strip() for p in line_str.split('|')[1:-1]]
            if not parts or all(p == '' for p in parts) or all(set(p).issubset({'-', ':', ' '}) for p in parts) or '---' in parts[0]:
                continue
            if headers is None:
                headers = parts
            else:
                clean_parts = [p.replace('`', '').replace('**', '').strip() for p in parts]
                if len(clean_parts) == len(headers) and clean_parts[0].startswith('MTVN'):
                    rows.append(clean_parts)
                    
    col_map = {
        'Mã NV': 'emp_id', 'Employee ID': 'emp_id', 'EmpID': 'emp_id',
        'Tên Tiếng Việt': 'emp_name_vn', 'Vietnamese Name': 'emp_name_vn',
        'Tên Tiếng Anh': 'full_name_en', 'English Name': 'full_name_en',
        'Email Nhân Viên': 'email', 'Employee Email': 'email',
        'Team Trực Thuộc': 'team', 'Team': 'team',
        'Team Leader': 'tl_name',
        'Email TL': 'tl_email',
        'Senior Leader': 'senior_leader',
        'Email SL': 'sl_email',
        'Khách Hàng Phụ Trách': 'customer_group', 'Assigned Customer': 'customer_group',
        'Đầu Mối Khách Hàng': 'contact_names',
        'Email Khách Hàng': 'contact_emails',
        'Email CC Khách Hàng': 'client_cc'
    }
    
    df = pd.DataFrame(rows, columns=headers)
    df = df.rename(columns=col_map)
    df['emp_id'] = df['emp_id'].str.strip()
    df = df[~df['emp_id'].isin(EXCLUDED_TOTAL)]
    return df

def normalize_shift(shift_str):
    """Normalizes raw shift descriptions to canonical short codes."""
    if not shift_str or pd.isna(shift_str):
        return None
    s = str(shift_str).strip()
    return SHIFT_LABEL_MAP.get(s, s)

def load_late_early_requests(file_path=REQ_LATE_EARLY):
    """
    Loads 'Đăng ký đi muộn, về sớm', expands recurring day-of-week filters,
    and returns a DataFrame of daily approved late/early allowances:
    [emp_id, date, approved_late_ci_mins, approved_early_co_mins, reason, reason_group]
    """
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    df = pd.read_excel(file_path, skiprows=4)
    # Filter approved
    df = df[df['Trạng thái'] == 'Đã duyệt'].copy()
    
    # DOW map for Vietnamese: 'Thứ 2': 0, 'Thứ 3': 1, 'Thứ 4': 2, 'Thứ 5': 3, 'Thứ 6': 4, 'Thứ 7': 5, 'Chủ nhật': 6
    dow_map = {
        'thứ 2': 0, 'thu 2': 0, 't2': 0,
        'thứ 3': 1, 'thu 3': 1, 't3': 1,
        'thứ 4': 2, 'thu 4': 2, 't4': 2,
        'thứ 5': 3, 'thu 5': 3, 't5': 3,
        'thứ 6': 4, 'thu 6': 4, 't6': 4,
        'thứ 7': 5, 'thu 7': 5, 't7': 5,
        'chủ nhật': 6, 'chu nhat': 6, 'cn': 6
    }
    
    records = []
    for _, row in df.iterrows():
        emp_id = str(row.get('Mã nhân viên', '')).strip()
        if not emp_id or emp_id == 'nan':
            continue
            
        t_from = row.get('Từ ngày')
        t_to = row.get('Đến ngày')
        if pd.isna(t_from) or pd.isna(t_to):
            continue
            
        d_from = pd.to_datetime(t_from).date()
        d_to = pd.to_datetime(t_to).date()
        if d_from > d_to:
            d_from, d_to = d_to, d_from
            
        apply_for = str(row.get('Áp dụng cho', '')).lower().strip()
        all_week = ('cả tuần' in apply_for or 'ca tuan' in apply_for or not apply_for or apply_for == 'nan')
        
        # Identify specific target weekdays
        target_dows = set()
        if not all_week:
            for k, dow_int in dow_map.items():
                if k in apply_for:
                    target_dows.add(dow_int)
                    
        late_ci_m = float(row.get('Đi muộn đầu ca (phút)', 0) or 0)
        early_co_m = float(row.get('Về sớm cuối ca (phút)', 0) or 0)
        reason = str(row.get('Lý do đi muộn, về sớm', ''))
        reason_grp = str(row.get('Nhóm lý do', ''))
        
        cur = d_from
        while cur <= d_to:
            if all_week or (cur.weekday() in target_dows):
                records.append({
                    'emp_id': emp_id,
                    'date': cur,
                    'approved_late_ci_mins': late_ci_m,
                    'approved_early_co_mins': early_co_m,
                    'reason': reason,
                    'reason_group': reason_grp
                })
            cur += datetime.timedelta(days=1)
            
    res_df = pd.DataFrame(records)
    if not res_df.empty:
        # If multiple requests on same day, sum approved minutes
        res_df = res_df.groupby(['emp_id', 'date']).agg({
            'approved_late_ci_mins': 'max',
            'approved_early_co_mins': 'max',
            'reason': lambda x: '; '.join(str(v) for v in x if v and str(v) != 'nan'),
            'reason_group': lambda x: '; '.join(str(v) for v in x if v and str(v) != 'nan')
        }).reset_index()
    return res_df

def load_overtime_requests(file_path=REQ_OVERTIME):
    """
    Loads 'Đơn_đăng_ký_làm_thêm', applies COALESCE fallback for employee ID,
    and returns a DataFrame of approved OT sessions:
    [emp_id, date, ot_hours, ot_type, ot_timing, ot_start, ot_end]
    """
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    df = pd.read_excel(file_path, skiprows=4)
    df = df[df['Trạng thái'] == 'Đã duyệt'].copy()
    
    records = []
    for _, row in df.iterrows():
        # Defect 3: Fallback coalesce logic
        emp_subject = row.get('Mã nhân viên làm thêm')
        emp_submitter = row.get('Mã nhân viên')
        emp_id = str(emp_subject).strip() if pd.notna(emp_subject) and str(emp_subject).strip() != 'nan' else str(emp_submitter).strip()
        
        if not emp_id or emp_id == 'nan':
            continue
            
        ot_from = row.get('Làm thêm từ')
        ot_to = row.get('Làm thêm đến')
        if pd.isna(ot_from):
            continue
            
        dt_from = pd.to_datetime(ot_from)
        dt_to = pd.to_datetime(ot_to) if pd.notna(ot_to) else dt_from
        
        date_val = dt_from.date()
        ot_hours = float(row.get('Số giờ làm thêm', 0) or 0)
        ot_type = str(row.get('Loại làm thêm', 'Hưởng lương')).strip()
        ot_timing = str(row.get('Thời điểm làm thêm', 'Sau ca làm việc')).strip()
        
        records.append({
            'emp_id': emp_id,
            'date': date_val,
            'ot_hours': ot_hours,
            'ot_type': ot_type,
            'ot_timing': ot_timing,
            'ot_start': dt_from.strftime('%H:%M') if pd.notna(dt_from) else None,
            'ot_end': dt_to.strftime('%H:%M') if pd.notna(dt_to) else None
        })
        
    res_df = pd.DataFrame(records)
    if not res_df.empty:
        # Group by emp_id and date
        res_df = res_df.groupby(['emp_id', 'date']).agg({
            'ot_hours': 'sum',
            'ot_type': lambda x: '/'.join(set(str(v) for v in x)),
            'ot_timing': lambda x: '/'.join(set(str(v) for v in x)),
            'ot_start': 'min',
            'ot_end': 'max'
        }).reset_index()
    return res_df

def load_wfh_requests(file_path=REQ_WFH):
    """
    Loads 'Dang_ky_lam_viec_tu_xa' with skiprows=3 (Defect 1).
    Validates datetime columns and returns DataFrame:
    [emp_id, date, wfh_credit, shift_applied]
    """
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    df = pd.read_excel(file_path, skiprows=3)
    df = df[df['Trạng thái'] == 'Đã duyệt'].copy()
    
    records = []
    for _, row in df.iterrows():
        emp_id = str(row.get('Mã nhân viên', '')).strip()
        if not emp_id or emp_id == 'nan':
            continue
            
        t_from = row.get('Từ ngày')
        t_to = row.get('Đến ngày')
        if pd.isna(t_from):
            continue
            
        dt_from = pd.to_datetime(t_from)
        dt_to = pd.to_datetime(t_to) if pd.notna(t_to) else dt_from
        
        # Safeguard if timestamps swapped
        if dt_from > dt_to:
            dt_from, dt_to = dt_to, dt_from
            
        wfh_credit = float(row.get('Số ngày làm việc từ xa', 1.0) or 1.0)
        shift_app = normalize_shift(row.get('Ca áp dụng'))
        
        cur = dt_from.date()
        end_d = dt_to.date()
        while cur <= end_d:
            records.append({
                'emp_id': emp_id,
                'date': cur,
                'wfh_credit': wfh_credit,
                'shift_applied': shift_app
            })
            cur += datetime.timedelta(days=1)
            
    res_df = pd.DataFrame(records)
    if not res_df.empty:
        res_df = res_df.groupby(['emp_id', 'date']).agg({
            'wfh_credit': 'max',
            'shift_applied': 'first'
        }).reset_index()
    return res_df

def load_shift_change_requests(file_path=REQ_SHIFT_CHANGE):
    """Loads approved shift change requests."""
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    df = pd.read_excel(file_path, skiprows=4)
    df = df[df['Trạng thái'] == 'Đã duyệt'].copy()
    
    records = []
    for _, row in df.iterrows():
        emp_id = str(row.get('Mã nhân viên', '')).strip()
        if not emp_id or emp_id == 'nan':
            continue
        d_val = row.get('Ngày làm việc')
        if pd.isna(d_val):
            continue
        date_val = pd.to_datetime(d_val).date()
        old_shift = normalize_shift(row.get('Ca hiện tại'))
        new_shift = normalize_shift(row.get('Ca đăng ký đổi'))
        
        records.append({
            'emp_id': emp_id,
            'date': date_val,
            'old_shift': old_shift,
            'new_shift': new_shift
        })
    return pd.DataFrame(records)

def load_leave_applications(file_path=REQ_LEAVE):
    """
    Loads 'Don_xin_nghi', expands multi-day leaves into daily slices (resolving cross-week bug),
    and categorizes notice lead time (Sudden, Short, Prompt, Planned).
    """
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    df = pd.read_excel(file_path, skiprows=4)
    df = df[df['Trạng thái'] == 'Đã duyệt'].copy()
    
    records = []
    for _, row in df.iterrows():
        emp_id = str(row.get('Mã nhân viên', '')).strip()
        if not emp_id or emp_id == 'nan':
            continue
            
        t_from = row.get('Từ ngày')
        t_to = row.get('Đến ngày')
        if pd.isna(t_from):
            continue
            
        d_from = pd.to_datetime(t_from).date()
        d_to = pd.to_datetime(t_to).date() if pd.notna(t_to) else d_from
        if d_from > d_to:
            d_from, d_to = d_to, d_from
            
        leave_days = float(row.get('Số ngày nghỉ', 1.0) or 1.0)
        leave_type = str(row.get('Loại nghỉ', 'Nghỉ phép')).strip()
        reason = str(row.get('Lý do nghỉ', '')).strip()
        
        # Calculate notice lead time
        sub_date = row.get('Ngày nộp đơn')
        lead_days = 30
        if pd.notna(sub_date):
            d_sub = pd.to_datetime(sub_date).date()
            # Defect 11: Fix future year typo (e.g. 2027 instead of 2026)
            if d_sub.year > d_from.year:
                try:
                    d_sub = d_sub.replace(year=d_from.year)
                except Exception:
                    pass
            lead_days = (d_from - d_sub).days
            
        if lead_days <= NOTICE_SUDDEN_MAX_DAYS:
            notice_type = 'Sudden Notice (<=1d)'
        elif lead_days <= NOTICE_SHORT_MAX_DAYS:
            notice_type = 'Short Notice (2-7d)'
        elif lead_days <= NOTICE_PROMPT_MAX_DAYS:
            notice_type = 'Prompt Notice (8-29d)'
        else:
            notice_type = 'Planned/Normal (>=30d)'
            
        # Macro Leave Taxonomy (4 groups)
        if any(k in leave_type for k in ['bù', 'bu']):
            macro_cat = 'Compensatory Leave'
        elif any(k in leave_type for k in ['BHXH', 'ốm', 'thai sản', 'con ốm']):
            macro_cat = 'Health & Family Care'
        elif any(k in leave_type for k in ['kết hôn', 'ma chay', 'chế độ']):
            macro_cat = 'Statutory & Welfare'
        else:
            macro_cat = 'Annual & Personal Leave'
            
        # Slicing across all calendar days in range
        cur = d_from
        total_span_days = max(1, (d_to - d_from).days + 1)
        per_day_credit = round(leave_days / total_span_days, 2)
        if per_day_credit > 1.0:
            per_day_credit = 1.0
            
        while cur <= d_to:
            records.append({
                'emp_id': emp_id,
                'date': cur,
                'leave_type': leave_type,
                'macro_category': macro_cat,
                'leave_credit': per_day_credit,
                'notice_lead_days': lead_days,
                'notice_type': notice_type,
                'reason': reason,
                'orig_from': d_from,
                'orig_to': d_to
            })
            cur += datetime.timedelta(days=1)
            
    res_df = pd.DataFrame(records)
    return res_df

def load_business_trip_requests(file_path=REQ_BUSINESS_TRIP):
    """Loads approved business trip requests ('De_nghi_di_cong_tac')."""
    if not os.path.exists(file_path):
        return pd.DataFrame()
        
    df = pd.read_excel(file_path, skiprows=4)
    df = df[df['Trạng thái'] == 'Đã duyệt'].copy()
    
    records = []
    for _, row in df.iterrows():
        emp_id = str(row.get('Mã nhân viên', '')).strip()
        if not emp_id or emp_id == 'nan':
            continue
        t_from = row.get('Từ ngày')
        t_to = row.get('Đến ngày')
        if pd.isna(t_from):
            continue
        d_from = pd.to_datetime(t_from).date()
        d_to = pd.to_datetime(t_to).date() if pd.notna(t_to) else d_from
        if d_from > d_to:
            d_from, d_to = d_to, d_from
            
        trip_days = float(row.get('Số ngày đi công tác', 1.0) or 1.0)
        location = str(row.get('Địa điểm công tác', '')).strip()
        purpose = str(row.get('Mục đích công tác', '')).strip()
        
        cur = d_from
        while cur <= d_to:
            records.append({
                'emp_id': emp_id,
                'date': cur,
                'trip_days': trip_days,
                'location': location,
                'purpose': purpose
            })
            cur += datetime.timedelta(days=1)
            
    return pd.DataFrame(records)

def compute_weekly_request_summary(start_date, end_date):
    """
    Computes aggregated volume counts for all 6 MISA request categories
    falling within [start_date, end_date].
    """
    summary = []
    
    # 1. Leave
    leave_df = load_leave_applications()
    if not leave_df.empty:
        leave_df = leave_df[~leave_df['emp_id'].isin(EXCLUDED_TOTAL)]
        l_sub = leave_df[(leave_df['date'] >= start_date) & (leave_df['date'] <= end_date)]
        leave_count = l_sub['emp_id'].nunique() if not l_sub.empty else 0
        leave_days_sum = l_sub['leave_credit'].sum() if not l_sub.empty else 0.0
    else:
        leave_count, leave_days_sum = 0, 0.0
    summary.append({
        'request_type': 'Leave Applications (Đơn xin nghỉ)',
        'count': leave_count,
        'metric_detail': f"{leave_days_sum:.1f} man-days requested"
    })
    
    # 2. Overtime
    ot_df = load_overtime_requests()
    if not ot_df.empty:
        ot_df = ot_df[~ot_df['emp_id'].isin(EXCLUDED_TOTAL)]
        ot_sub = ot_df[(ot_df['date'] >= start_date) & (ot_df['date'] <= end_date)]
        ot_count = len(ot_sub)
        ot_hours_sum = ot_sub['ot_hours'].sum()
    else:
        ot_count, ot_hours_sum = 0, 0.0
    summary.append({
        'request_type': 'Overtime Requests (Đơn làm thêm)',
        'count': ot_count,
        'metric_detail': f"{ot_hours_sum:.1f} OT hours approved"
    })
    
    # 3. WFH
    wfh_df = load_wfh_requests()
    if not wfh_df.empty:
        wfh_df = wfh_df[~wfh_df['emp_id'].isin(EXCLUDED_TOTAL)]
        w_sub = wfh_df[(wfh_df['date'] >= start_date) & (wfh_df['date'] <= end_date)]
        wfh_count = len(w_sub)
        wfh_days = w_sub['wfh_credit'].sum()
    else:
        wfh_count, wfh_days = 0, 0.0
    summary.append({
        'request_type': 'Work From Home (Đăng ký làm việc từ xa)',
        'count': wfh_count,
        'metric_detail': f"{wfh_days:.1f} remote days"
    })
    
    # 4. Shift Change
    sc_df = load_shift_change_requests()
    if not sc_df.empty:
        sc_df = sc_df[~sc_df['emp_id'].isin(EXCLUDED_TOTAL)]
        sc_sub = sc_df[(sc_df['date'] >= start_date) & (sc_df['date'] <= end_date)]
        sc_count = len(sc_sub)
    else:
        sc_count = 0
    summary.append({
        'request_type': 'Shift Changes (Đề nghị đổi ca)',
        'count': sc_count,
        'metric_detail': f"{sc_count} shift transfers"
    })
    
    # 5. Late / Early
    le_df = load_late_early_requests()
    if not le_df.empty:
        le_df = le_df[~le_df['emp_id'].isin(EXCLUDED_TOTAL)]
        le_sub = le_df[(le_df['date'] >= start_date) & (le_df['date'] <= end_date)]
        le_count = len(le_sub)
    else:
        le_count = 0
    summary.append({
        'request_type': 'Late CI / Early CO (Đi muộn, về sớm có phép)',
        'count': le_count,
        'metric_detail': f"{le_count} excused arrivals/departures"
    })
    
    # 6. Business Trip
    bt_df = load_business_trip_requests()
    if not bt_df.empty:
        bt_df = bt_df[~bt_df['emp_id'].isin(EXCLUDED_TOTAL)]
        bt_sub = bt_df[(bt_df['date'] >= start_date) & (bt_df['date'] <= end_date)]
        bt_count = len(bt_sub)
    else:
        bt_count = 0
    summary.append({
        'request_type': 'Business Trips (Đề nghị đi công tác)',
        'count': bt_count,
        'metric_detail': f"{bt_count} business trip sessions"
    })
    
    return pd.DataFrame(summary)

