# -*- coding: utf-8 -*-
"""
Anomaly & Risk Detection Module
Identifies Burnout Risk, Manual Override Abuse, Chronic Lateness,
Frequent Excused Requests, and Customer Resource Bottlenecks.
Outputs prioritized P&C Action Triggers.
"""

import pandas as pd
from .config import (
    BURNOUT_CONSECUTIVE_DAYS, BURNOUT_LATE_CO_MINUTES,
    CHRONIC_LATENESS_MINUTES, CHRONIC_LATENESS_MAX_MONTH,
    MANUAL_OVERRIDE_MAX_MONTH, EXCUSED_FREQ_MAX_WEEK
)

def detect_burnout_risks(daily_df):
    """
    Identifies employees clocking out > 30m without OT for >= 3 consecutive working days.
    Consolidates streaks so each employee has at most 1 primary Burnout Risk flag
    summarizing their maximum consecutive streak.
    """
    if daily_df.empty:
        return []
        
    flags = []
    present_df = daily_df[daily_df['actual_day'] > 0].sort_values(by=['emp_id', 'date'])
    
    for emp_id, grp in present_df.groupby('emp_id'):
        cur_count = 0
        cur_streak_dates = []
        max_streak_count = 0
        max_streak_dates = []
        
        full_name = grp['full_name_en'].iloc[0]
        team = grp['team_clean'].iloc[0]
        customer = grp['customer_group'].iloc[0]
        
        for _, row in grp.iterrows():
            is_unapproved_stay = row['is_unapproved_late_stay'] and not row['is_makeup_time']
            if is_unapproved_stay:
                cur_count += 1
                cur_streak_dates.append(row['date'])
                if cur_count > max_streak_count:
                    max_streak_count = cur_count
                    max_streak_dates = list(cur_streak_dates)
            else:
                cur_count = 0
                cur_streak_dates = []
                
        if max_streak_count >= BURNOUT_CONSECUTIVE_DAYS:
            date_range_str = f"{max_streak_dates[0].strftime('%d/%m')} to {max_streak_dates[-1].strftime('%d/%m')}"
            flags.append({
                'category': 'Burnout Risk',
                'severity': 'HIGH',
                'emp_id': emp_id,
                'full_name_en': full_name,
                'team': team,
                'customer_group': customer,
                'trigger_condition': f'Late Check-out > {BURNOUT_LATE_CO_MINUTES}m without approved OT for {max_streak_count} consecutive days ({date_range_str})',
                'dates': max_streak_dates,
                'recommended_action': 'Schedule 1-on-1 check-in; consult Team Leader to assess project workload & deadline pressure'
            })
            
    return flags

def detect_manual_override_abuse(emp_summary_df):
    """
    Flags employees with >= 2 manual override ('Cập nhật công') days in the period.
    """
    if emp_summary_df.empty:
        return []
        
    flags = []
    abusers = emp_summary_df[emp_summary_df['manual_overrides'] > MANUAL_OVERRIDE_MAX_MONTH]
    for _, row in abusers.iterrows():
        flags.append({
            'category': 'Manual Override Procedural Violation',
            'severity': 'MEDIUM',
            'emp_id': row['emp_id'],
            'full_name_en': row['full_name_en'],
            'team': row['team'],
            'customer_group': row['customer_group'],
            'trigger_condition': f'Manual override ("Cập nhật công") used {int(row["manual_overrides"])} times (Company policy limit is 1/month)',
            'dates': [],
            'recommended_action': 'Issue compliance reminder; require formal explanation for forgotten badge swipes'
        })
    return flags

def detect_chronic_lateness(emp_summary_df):
    """
    Flags employees with >= 2 chronic lateness (>30 mins late) occurrences.
    """
    if emp_summary_df.empty:
        return []
        
    flags = []
    chronic = emp_summary_df[emp_summary_df['chronic_late_count'] >= CHRONIC_LATENESS_MAX_MONTH]
    for _, row in chronic.iterrows():
        flags.append({
            'category': 'Chronic / Severe Lateness',
            'severity': 'HIGH',
            'emp_id': row['emp_id'],
            'full_name_en': row['full_name_en'],
            'team': row['team'],
            'customer_group': row['customer_group'],
            'trigger_condition': f'Unexcused Check-in > 30 mins late {int(row["chronic_late_count"])} times without approved request',
            'dates': [],
            'recommended_action': 'Conduct formal performance review; verify commute / family distress reasons'
        })
    return flags

def detect_frequent_excused_requests(emp_summary_df):
    """
    Flags employees with >= 2 excused late/early requests in a single calendar week.
    """
    if emp_summary_df.empty:
        return []
        
    flags = []
    frequent = emp_summary_df[emp_summary_df['excused_late_count'] >= EXCUSED_FREQ_MAX_WEEK]
    for _, row in frequent.iterrows():
        flags.append({
            'category': 'Frequent Excused Arrival/Departure',
            'severity': 'LOW',
            'emp_id': row['emp_id'],
            'full_name_en': row['full_name_en'],
            'team': row['team'],
            'customer_group': row['customer_group'],
            'trigger_condition': f'{int(row["excused_late_count"])} approved late/early requests in week (Limit: 1/week)',
            'dates': [],
            'recommended_action': 'Proactive P&C reach-out; assess if shift adjustment or flexible arrangement is warranted'
        })
    return flags

def detect_all_anomalies(daily_df, emp_summary_df, capacity_loss_df=None):
    """
    Consolidates all governance alerts and prioritizes them by severity:
    HIGH -> MEDIUM -> LOW.
    """
    all_flags = []
    all_flags.extend(detect_burnout_risks(daily_df))
    all_flags.extend(detect_chronic_lateness(emp_summary_df))
    all_flags.extend(detect_manual_override_abuse(emp_summary_df))
    all_flags.extend(detect_frequent_excused_requests(emp_summary_df))
    
    if capacity_loss_df is not None and not capacity_loss_df.empty:
        bottlenecks = capacity_loss_df[capacity_loss_df['is_bottleneck']]
        for _, row in bottlenecks.iterrows():
            all_flags.append({
                'category': 'Resource Bottleneck Risk',
                'severity': 'HIGH',
                'emp_id': 'MULTIPLE',
                'full_name_en': ', '.join(row['absent_employees']),
                'team': 'Client Group',
                'customer_group': row['customer_group'],
                'trigger_condition': f'{row["absent_count"]} employees on leave simultaneously on {row["date"]}',
                'dates': [row['date']],
                'recommended_action': f'Notify {row["customer_group"]} Team Leader to review delivery coverage'
            })
            
    sev_order = {'HIGH': 0, 'MEDIUM': 1, 'LOW': 2}
    all_flags.sort(key=lambda x: sev_order.get(x['severity'], 3))
    return all_flags
