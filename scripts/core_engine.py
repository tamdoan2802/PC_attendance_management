# -*- coding: utf-8 -*-
"""
Core Attendance Analytics Engine
Reconciles raw daily attendance facts with approved requests and Master Entities.
Evaluates 5 operational pillars: Working Time, Punctuality, Departure Regularity,
Operational Stability, and Overtime Intensity.
Adheres strictly to User Rule: Chronic Lateness is >30m and strictly UNEXCUSED.
"""

import datetime
import pandas as pd
import os
import sys

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

try:
    from .config import (
        OT_TIER_1_MAX, OT_TIER_2_MAX, BURNOUT_LATE_CO_MINUTES, BURNOUT_CONSECUTIVE_DAYS,
        EXCLUDED_EMP_IDS
    )
    from .parse_timesheet import load_attendance_dataset
    from .load_requests import (
        load_master_entities, load_late_early_requests, load_overtime_requests,
        load_wfh_requests, load_shift_change_requests, load_leave_applications
    )
except (ImportError, ValueError):
    from config import (
        OT_TIER_1_MAX, OT_TIER_2_MAX, BURNOUT_LATE_CO_MINUTES, BURNOUT_CONSECUTIVE_DAYS,
        EXCLUDED_EMP_IDS
    )
    from parse_timesheet import load_attendance_dataset
    from load_requests import (
        load_master_entities, load_late_early_requests, load_overtime_requests,
        load_wfh_requests, load_shift_change_requests, load_leave_applications
    )

def reconcile_daily_attendance(start_date, end_date):
    """
    Synthesizes a single unified daily attendance truth for each employee and date.
    Performs cross-table joining:
    FACT_Attendance + Entities + Req_LateEarly + Req_OT + Req_WFH + Req_ShiftChange + Req_Leave
    """
    att_df = load_attendance_dataset(start_date, end_date)
    if att_df.empty:
        return pd.DataFrame()
    att_df = att_df[~att_df['emp_id'].isin(EXCLUDED_EMP_IDS)]
        
    ent_df = load_master_entities()
    if not ent_df.empty:
        att_df = att_df.merge(ent_df, on='emp_id', how='left')
        att_df['full_name_en'] = att_df['full_name_en'].fillna(att_df['emp_name'])
        att_df['team'] = att_df['team'].fillna(att_df['department'])
        att_df['customer_group'] = att_df['customer_group'].fillna('Internal')
    else:
        att_df['full_name_en'] = att_df['emp_name']
        att_df['team'] = att_df['department']
        att_df['customer_group'] = 'Internal'
        
    late_early_df = load_late_early_requests()
    ot_df = load_overtime_requests()
    wfh_df = load_wfh_requests()
    sc_df = load_shift_change_requests()
    leave_df = load_leave_applications()
    
    if not late_early_df.empty:
        att_df = att_df.merge(late_early_df, on=['emp_id', 'date'], how='left')
    else:
        att_df['approved_late_ci_mins'] = 0.0
        att_df['approved_early_co_mins'] = 0.0
        att_df['reason'] = ''
        att_df['reason_group'] = ''
        
    if not ot_df.empty:
        att_df = att_df.merge(ot_df, on=['emp_id', 'date'], how='left')
    else:
        att_df['ot_hours'] = 0.0
        att_df['ot_type'] = ''
        att_df['ot_timing'] = ''
        
    if not wfh_df.empty:
        att_df = att_df.merge(wfh_df, on=['emp_id', 'date'], how='left')
    else:
        att_df['wfh_credit'] = 0.0
        att_df['shift_applied'] = None
        
    if not leave_df.empty:
        l_sub = leave_df[(leave_df['date'] >= start_date) & (leave_df['date'] <= end_date)]
        if not l_sub.empty:
            l_dedup = l_sub.groupby(['emp_id', 'date']).agg({
                'macro_category': 'first',
                'notice_type': 'first',
                'notice_lead_days': 'first',
                'leave_credit': 'max'
            }).reset_index().rename(columns={'leave_credit': 'req_leave_credit'})
            att_df = att_df.merge(l_dedup, on=['emp_id', 'date'], how='left')
        else:
            att_df['macro_category'] = None
            att_df['notice_type'] = None
            att_df['notice_lead_days'] = None
    else:
        att_df['macro_category'] = None
        att_df['notice_type'] = None
        att_df['notice_lead_days'] = None
        
    att_df['approved_late_ci_mins'] = att_df['approved_late_ci_mins'].fillna(0.0)
    att_df['approved_early_co_mins'] = att_df['approved_early_co_mins'].fillna(0.0)
    att_df['ot_hours'] = att_df['ot_hours'].fillna(0.0)
    att_df['wfh_credit'] = att_df['wfh_credit'].fillna(0.0)
    
    # Check-in compliance
    att_df['is_excused_late'] = (
        att_df['is_late_ci'] & (att_df['approved_late_ci_mins'] >= att_df['late_ci_mins'])
    )
    att_df['is_unexcused_late'] = att_df['is_late_ci'] & (~att_df['is_excused_late'])
    
    # USER RULE: Chronic is >30 mins late WITHOUT valid reason or approved request (strictly unexcused)
    att_df['is_chronic_late_ci'] = att_df['is_unexcused_late'] & (att_df['late_ci_mins'] > 30)
    
    # Check-out compliance
    att_df['is_excused_early'] = (
        att_df['is_early_co'] & (att_df['approved_early_co_mins'] >= att_df['early_co_mins'])
    )
    att_df['is_unexcused_early'] = att_df['is_early_co'] & (~att_df['is_excused_early'])
    
    # Overtime vs Late Stay
    att_df['has_approved_ot'] = att_df['ot_hours'] > 0.0
    att_df['is_unapproved_late_stay'] = (
        att_df['is_late_co'] & (att_df['late_co_mins'] > BURNOUT_LATE_CO_MINUTES) & (~att_df['has_approved_ot'])
    )
    
    # Make-up Time Detection
    att_df['is_makeup_time'] = (
        att_df['is_late_co'] & (att_df['late_ci_mins'] > 0) & 
        (abs(att_df['late_co_mins'] - att_df['late_ci_mins']) <= 30) &
        (~att_df['has_approved_ot'])
    )
    
    # OT Tiers
    def get_ot_tier(row):
        mins = row['ot_hours'] * 60.0
        if mins <= 0:
            return None
        if mins <= OT_TIER_1_MAX:
            return 'Tier 1 (Light <=60m)'
        elif mins <= OT_TIER_2_MAX:
            return 'Tier 2 (Moderate 61-120m)'
        else:
            return 'Tier 3 (Heavy >120m)'
            
    att_df['ot_tier'] = att_df.apply(get_ot_tier, axis=1)
    
    # Clean Team Name with separate EA/PA designation
    def clean_team_name(t):
        if not t or pd.isna(t):
            return 'Other'
        t_str = str(t).strip()
        if 'Personal Assistant' in t_str or 'EA/PA' in t_str or 'Executive Assistant' in t_str:
            return 'EA/PA'
        return t_str
        
    att_df['team_clean'] = att_df['team'].apply(clean_team_name)
    
    return att_df

def compute_employee_summary(daily_df):
    """
    Computes aggregated performance metrics per employee across the 5 pillars.
    (Note: AQ Score omitted per user specification).
    """
    if daily_df.empty:
        return pd.DataFrame()
        
    groups = []
    for emp_id, grp in daily_df.groupby('emp_id'):
        full_name_en = grp['full_name_en'].iloc[0]
        emp_name_vn = grp['emp_name'].iloc[0]
        team = grp['team_clean'].iloc[0]
        customer = grp['customer_group'].iloc[0]
        
        actual_days = grp['actual_day'].sum()
        working_days = grp['working_credit'].sum()
        total_actual_hours = grp['actual_hours'].sum()
        total_std_hours = grp['std_hours'].sum()
        
        actual_avg_per_day = round(total_actual_hours / actual_days, 2) if actual_days > 0 else 0.0
        shift_std_per_day = 8.0
        daily_delta = round(actual_avg_per_day - shift_std_per_day, 2) if actual_days > 0 else 0.0
        net_surplus_hours = round(total_actual_hours - total_std_hours, 2)
        adherence_rate = round((total_actual_hours / total_std_hours) * 100.0, 1) if total_std_hours > 0 else 100.0
        
        total_late_ci_count = grp['is_late_ci'].sum()
        excused_late_count = grp['is_excused_late'].sum()
        unexcused_late_count = grp['is_unexcused_late'].sum()
        chronic_late_count = grp['is_chronic_late_ci'].sum()
        
        unapproved_late_stays = grp['is_unapproved_late_stay'].sum()
        makeup_days = grp['is_makeup_time'].sum()
        
        ot_sessions = (grp['ot_hours'] > 0).sum()
        total_ot_hours = grp['ot_hours'].sum()
        
        manual_overrides = grp['is_manual_override'].sum()
        leave_days = grp['leave_credit'].sum()
        
        unapproved_stay_hours = grp.loc[grp['is_unapproved_late_stay'], 'late_co_mins'].sum() / 60.0
        overwork_index = round(
            (total_ot_hours + unapproved_stay_hours) / total_std_hours, 2
        ) if total_std_hours > 0 else 0.0
        
        groups.append({
            'emp_id': emp_id,
            'full_name_en': full_name_en,
            'emp_name_vn': emp_name_vn,
            'team': team,
            'customer_group': customer,
            'actual_days': actual_days,
            'working_days': working_days,
            'total_actual_hours': total_actual_hours,
            'total_std_hours': total_std_hours,
            'actual_avg_per_day': actual_avg_per_day,
            'shift_std_per_day': shift_std_per_day,
            'daily_delta': daily_delta,
            'net_surplus_hours': net_surplus_hours,
            'adherence_rate': adherence_rate,
            'unexcused_late_count': unexcused_late_count,
            'excused_late_count': excused_late_count,
            'chronic_late_count': chronic_late_count,
            'unapproved_late_stays': unapproved_late_stays,
            'makeup_days': makeup_days,
            'ot_sessions': ot_sessions,
            'total_ot_hours': total_ot_hours,
            'manual_overrides': manual_overrides,
            'leave_days': leave_days,
            'overwork_index': overwork_index
        })
        
    return pd.DataFrame(groups)

def compute_team_summary(emp_summary_df):
    """Aggregates metrics by operational Team."""
    if emp_summary_df.empty:
        return pd.DataFrame()
        
    res = emp_summary_df.groupby('team').agg({
        'emp_id': 'count',
        'actual_days': 'sum',
        'total_actual_hours': 'sum',
        'total_std_hours': 'sum',
        'unexcused_late_count': 'sum',
        'unapproved_late_stays': 'sum',
        'total_ot_hours': 'sum',
        'leave_days': 'sum'
    }).reset_index().rename(columns={'emp_id': 'headcount'})
    
    res['adherence_rate'] = (res['total_actual_hours'] / res['total_std_hours'] * 100.0).round(1)
    res['actual_avg_per_day'] = (res['total_actual_hours'] / res['actual_days']).round(2)
    return res

def compute_customer_capacity_loss(daily_df):
    """
    Computes man-days of capacity loss per Customer Group
    and flags concentrated leave bottlenecks (>= 2 employees of same customer on leave).
    """
    if daily_df.empty:
        return pd.DataFrame()
        
    leave_records = daily_df[daily_df['leave_credit'] > 0].copy()
    if leave_records.empty:
        return pd.DataFrame()
        
    summary = []
    for (cust, date_val), grp in leave_records.groupby(['customer_group', 'date']):
        emps = grp['full_name_en'].unique().tolist()
        num_emps = len(emps)
        total_leave_credit = grp['leave_credit'].sum()
        is_bottleneck = (num_emps >= 2) and (cust != 'Internal')
        
        summary.append({
            'customer_group': cust,
            'date': date_val,
            'absent_employees': emps,
            'absent_count': num_emps,
            'total_leave_credit': total_leave_credit,
            'is_bottleneck': is_bottleneck
        })
        
    return pd.DataFrame(summary)

def compute_customer_personnel_summary(emp_summary_df):
    """
    Groups personnel by assigned Customer Account (Master Data: entities/notification_routing_map.md)
    and summarizes headcount, employee names, actual working hours, leave days, and OT.
    """
    if emp_summary_df.empty:
        return pd.DataFrame()
        
    cust_rows = []
    for cust, grp in emp_summary_df.groupby('customer_group'):
        emp_names = sorted(grp['full_name_en'].unique().tolist())
        headcount = len(emp_names)
        actual_hours = grp['total_actual_hours'].sum()
        ot_hours = grp['total_ot_hours'].sum()
        leave_days = grp['leave_days'].sum()
        unexcused_lates = grp['unexcused_late_count'].sum()
        late_stays = grp['unapproved_late_stays'].sum()
        
        cust_rows.append({
            'customer_group': cust,
            'headcount': headcount,
            'employees': ', '.join(emp_names),
            'total_actual_hours': round(actual_hours, 1),
            'total_ot_hours': round(ot_hours, 1),
            'total_leave_days': round(leave_days, 1),
            'unexcused_lates': int(unexcused_lates),
            'unapproved_late_stays': int(late_stays)
        })
        
    res_df = pd.DataFrame(cust_rows)
    return res_df.sort_values(by='headcount', ascending=False).reset_index(drop=True)

