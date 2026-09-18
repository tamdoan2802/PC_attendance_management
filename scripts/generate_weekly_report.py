# -*- coding: utf-8 -*-
"""
Weekly Attendance Report Generator
Generates comprehensive 5-pillar attendance reports in HTML and Markdown formats.
Strictly adheres to corporate invariants: FullNameEN, EA/PA isolation,
client-level capacity attribution, and respectful HR B1 phrasing.
"""

import sys
import os
import datetime
import io
import argparse

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from .config import (
    DEFAULT_PC_MANAGER_EMAIL, DEFAULT_SENDER_EMAIL, SKILL_DIR
)
from .core_engine import (
    reconcile_daily_attendance, compute_employee_summary,
    compute_team_summary, compute_customer_capacity_loss,
    compute_customer_personnel_summary
)
from .load_requests import compute_weekly_request_summary
from .detect_anomalies import detect_all_anomalies

def get_date_range_from_keyword(kw):
    today = datetime.date.today()
    if not kw or kw in ['last_week', 'prev_week']:
        cur_dow = today.weekday()
        this_mon = today - datetime.timedelta(days=cur_dow)
        last_mon = this_mon - datetime.timedelta(days=7)
        last_sun = last_mon + datetime.timedelta(days=6)
        return last_mon, last_sun
    elif kw in ['this_week', 'current_week']:
        cur_dow = today.weekday()
        this_mon = today - datetime.timedelta(days=cur_dow)
        return this_mon, today
    elif ':' in kw:
        parts = kw.split(':')
        d1 = datetime.datetime.strptime(parts[0].strip(), '%Y-%m-%d').date()
        d2 = datetime.datetime.strptime(parts[1].strip(), '%Y-%m-%d').date()
        return d1, d2
    else:
        try:
            d = datetime.datetime.strptime(kw.strip(), '%Y-%m-%d').date()
            return d, d
        except Exception:
            return get_date_range_from_keyword('last_week')

def build_markdown_report(start_date, end_date, team_sum, emp_sum, daily_df, loss_df, anomalies, req_sum=None, cust_sum=None):
    period_str = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"
    lines = []
    lines.append(f"# 📊 Workforce Attendance Quality Report: {period_str}")
    lines.append(f"> **Prepared For:** People & Culture Management ({DEFAULT_PC_MANAGER_EMAIL})")
    lines.append("> **Governance Benchmark:** Attendance Semantic Reference Framework (5 Core Operational Pillars)")
    lines.append("")
    
    # 1. Executive Summary
    lines.append("## 1. Executive Summary: Operational Team Overview")
    lines.append("| Operational Team | Headcount | Adherence % | Avg Hours/Day | Total OT Hours | Unexcused Lates | Unapproved Late Stays |")
    lines.append("|:---|:---:|:---:|:---:|:---:|:---:|:---:|")
    for _, r in team_sum.sort_values(by='adherence_rate', ascending=False).iterrows():
        lines.append(f"| **{r['team']}** | {r['headcount']} | {r['adherence_rate']}% | {r['actual_avg_per_day']}h | {r['total_ot_hours']}h | {r['unexcused_late_count']} | {r['unapproved_late_stays']} |")
    lines.append("")
    
    # 2. Punctuality & Check-In Violations
    lates = emp_sum[emp_sum['unexcused_late_count'] > 0].sort_values(by='unexcused_late_count', ascending=False)
    lines.append("## 2. Punctuality & Start-of-Shift Discipline")
    if lates.empty:
        lines.append("✅ **100% Punctuality Compliance:** No unexcused late check-ins recorded during this period.")
        lines.append("")
    else:
        lines.append("| Employee Name | Team | Client Account | Actual Days | Unexcused Lates | Chronic Lates (>30m unexcused) | Excused Lates |")
        lines.append("|:---|:---|:---|:---:|:---:|:---:|:---:|")
        for _, r in lates.iterrows():
            lines.append(f"| **{r['full_name_en']}** | {r['team']} | {r['customer_group']} | {r['actual_days']} | {int(r['unexcused_late_count'])} | {int(r['chronic_late_count'])} | {int(r['excused_late_count'])} |")
        lines.append("")
        
    # 3. Overtime & Departure Regularity
    ot_emps = emp_sum[(emp_sum['total_ot_hours'] > 0) | (emp_sum['unapproved_late_stays'] > 0)].sort_values(by='total_ot_hours', ascending=False)
    lines.append("## 3. Departure Regularity & Overtime Analysis")
    if ot_emps.empty:
        lines.append("ℹ️ Standard departures across all teams. No overtime or late staying clocked.")
        lines.append("")
    else:
        lines.append("| Employee Name | Team | Client Account | Approved OT Hours | OT Sessions | Unapproved Late Stays (>30m) | Make-up Days | Overwork Index |")
        lines.append("|:---|:---|:---|:---:|:---:|:---:|:---:|:---:|")
        for _, r in ot_emps.iterrows():
            lines.append(f"| **{r['full_name_en']}** | {r['team']} | {r['customer_group']} | {r['total_ot_hours']}h | {int(r['ot_sessions'])} | {int(r['unapproved_late_stays'])} | {int(r['makeup_days'])} | {r['overwork_index']} |")
        lines.append("")
        
    # 4. Leave & Client Capacity Loss
    lines.append("## 4. Leave Taxonomy & Capacity Impact")
    if loss_df.empty:
        lines.append("ℹ️ No planned or unplanned leave recorded in this period.")
        lines.append("")
    else:
        lines.append("| Customer / Project Group | Date | Absent Employees | Total Man-Days Lost | Bottleneck Risk |")
        lines.append("|:---|:---:|:---|:---:|:---:|")
        for _, r in loss_df.sort_values(by='date').iterrows():
            risk_badge = "🔴 **BOTTLENECK**" if r['is_bottleneck'] else "🟢 Normal"
            lines.append(f"| **{r['customer_group']}** | {r['date'].strftime('%d/%m')} | {', '.join(r['absent_employees'])} | {r['total_leave_credit']} | {risk_badge} |")
        lines.append("")
        
    # 5. Prioritized P&C Action Triggers
    lines.append("## 5. Prioritized People & Culture Action Triggers")
    if not anomalies:
        lines.append("✅ **No critical governance alerts.** Workforce operating smoothly within standard thresholds.")
        lines.append("")
    else:
        lines.append("| Priority | Issue Category | Employee / Subject | Client Group | Trigger Condition | Recommended P&C Action |")
        lines.append("|:---:|:---|:---|:---|:---|:---|")
        for a in anomalies:
            sev_icon = "🔴 HIGH" if a['severity'] == 'HIGH' else ("🟡 MEDIUM" if a['severity'] == 'MEDIUM' else "⚪ LOW")
            lines.append(f"| {sev_icon} | **{a['category']}** | {a['full_name_en']} | {a['customer_group']} | {a['trigger_condition']} | {a['recommended_action']} |")
        lines.append("")

    # 6. Weekly Request Volume Breakdown
    lines.append("## 6. Weekly Request Volume Breakdown")
    if req_sum is None or req_sum.empty:
        lines.append("ℹ️ No application requests submitted for this evaluation window.")
        lines.append("")
    else:
        lines.append("| Request Category | Total Volume | Operational Metric Detail |")
        lines.append("|:---|:---:|:---|")
        for _, r in req_sum.iterrows():
            lines.append(f"| **{r['request_type']}** | {r['count']} | {r['metric_detail']} |")
        lines.append("")

    # 7. Customer Account Capacity & Personnel Allocation
    lines.append("## 7. Personnel Distribution by Customer Account")
    if cust_sum is None or cust_sum.empty:
        lines.append("ℹ️ No customer group allocation available.")
        lines.append("")
    else:
        lines.append("| Customer Account | Headcount | Assigned Personnel | Actual Hours | Total OT | Leave Days | Unexcused Lates | Late Stays (>30m) |")
        lines.append("|:---|:---:|:---|:---:|:---:|:---:|:---:|:---:|")
        for _, r in cust_sum.iterrows():
            lines.append(f"| **{r['customer_group']}** | {r['headcount']} | {r['employees']} | {r['total_actual_hours']}h | {r['total_ot_hours']}h | {r['total_leave_days']} | {r['unexcused_lates']} | {r['unapproved_late_stays']} |")
        lines.append("")
        
    return "\n".join(lines)

def build_html_report(start_date, end_date, team_sum, emp_sum, daily_df, loss_df, anomalies, req_sum=None, cust_sum=None):
    period_str = f"{start_date.strftime('%d/%m/%Y')} &ndash; {end_date.strftime('%d/%m/%Y')}"
    
    html = f"""
    <html>
    <head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; color: #222; line-height: 1.5; }}
        h2 {{ color: #1a365d; border-bottom: 2px solid #2b6cb0; padding-bottom: 4px; margin-top: 24px; font-size: 16px; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 16px; font-size: 12px; }}
        th {{ background-color: #ebf8ff; color: #2b6cb0; font-weight: 600; text-align: left; padding: 8px; border: 1px solid #cbd5e0; }}
        td {{ padding: 6px 8px; border: 1px solid #e2e8f0; }}
        tr:nth-child(even) {{ background-color: #f7fafc; }}
        .badge-high {{ background-color: #fed7d7; color: #9b2c2c; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 11px; }}
        .badge-med {{ background-color: #feebc8; color: #9c4221; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 11px; }}
        .badge-low {{ background-color: #edf2f7; color: #4a5568; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 11px; }}
        .info-card {{ background-color: #f7fafc; border-left: 4px solid #3182ce; padding: 10px 14px; margin-bottom: 16px; }}
    </style>
    </head>
    <body>
        <div style="background-color: #1a365d; color: white; padding: 14px 18px; border-radius: 6px; margin-bottom: 18px;">
            <h1 style="margin: 0; font-size: 20px;">Workforce Attendance Quality Report</h1>
            <p style="margin: 4px 0 0 0; font-size: 13px; opacity: 0.9;">Evaluation Period: <strong>{period_str}</strong> | MyTeam Vietnam Co., Ltd.</p>
        </div>
        
        <div class="info-card">
            <strong>Operational Scope:</strong> Reconciled across timesheet punch clocks and 6 MISA application modules.
            Evaluated under corporate 5-minute grace period policy, burnout risk screening, and contractual capacity adherence.
        </div>

        <h2>1. Executive Summary: Operational Team Overview</h2>
        <table>
            <tr>
                <th>Operational Team</th>
                <th style="text-align:center;">Headcount</th>
                <th style="text-align:center;">Adherence %</th>
                <th style="text-align:center;">Avg Hours/Day</th>
                <th style="text-align:center;">Total OT</th>
                <th style="text-align:center;">Unexcused Lates</th>
                <th style="text-align:center;">Late Stays (&gt;30m)</th>
            </tr>
    """
    for _, r in team_sum.sort_values(by='adherence_rate', ascending=False).iterrows():
        html += f"""
            <tr>
                <td><strong>{r['team']}</strong></td>
                <td style="text-align:center;">{r['headcount']}</td>
                <td style="text-align:center;">{r['adherence_rate']}%</td>
                <td style="text-align:center;">{r['actual_avg_per_day']}h</td>
                <td style="text-align:center;">{r['total_ot_hours']}h</td>
                <td style="text-align:center;">{r['unexcused_late_count']}</td>
                <td style="text-align:center;">{r['unapproved_late_stays']}</td>
            </tr>
        """
    html += """
        </table>

        <h2>2. Punctuality & Start-of-Shift Discipline</h2>
    """
    lates = emp_sum[emp_sum['unexcused_late_count'] > 0].sort_values(by='unexcused_late_count', ascending=False)
    if lates.empty:
        html += "<p>✅ <em>100% Punctuality Compliance: No unexcused late check-ins recorded during this period.</em></p>"
    else:
        html += """
        <table>
            <tr>
                <th>Employee Name</th>
                <th>Team</th>
                <th>Client Account</th>
                <th style="text-align:center;">Actual Days</th>
                <th style="text-align:center;">Unexcused Lates</th>
                <th style="text-align:center;">Chronic Lates (&gt;30m unexcused)</th>
                <th style="text-align:center;">Excused Lates</th>
            </tr>
        """
        for _, r in lates.iterrows():
            html += f"""
            <tr>
                <td><strong>{r['full_name_en']}</strong></td>
                <td>{r['team']}</td>
                <td>{r['customer_group']}</td>
                <td style="text-align:center;">{r['actual_days']}</td>
                <td style="text-align:center; color:#c53030; font-weight:bold;">{int(r['unexcused_late_count'])}</td>
                <td style="text-align:center;">{int(r['chronic_late_count'])}</td>
                <td style="text-align:center;">{int(r['excused_late_count'])}</td>
            </tr>
            """
        html += "</table>"
        
    html += """
        <h2>3. Departure Regularity & Overtime Analysis</h2>
    """
    ot_emps = emp_sum[(emp_sum['total_ot_hours'] > 0) | (emp_sum['unapproved_late_stays'] > 0)].sort_values(by='total_ot_hours', ascending=False)
    if ot_emps.empty:
        html += "<p>ℹ️ <em>Standard departures across all teams. No overtime or late staying clocked.</em></p>"
    else:
        html += """
        <table>
            <tr>
                <th>Employee Name</th>
                <th>Team</th>
                <th>Client Account</th>
                <th style="text-align:center;">Approved OT</th>
                <th style="text-align:center;">OT Sessions</th>
                <th style="text-align:center;">Late Stays (&gt;30m)</th>
                <th style="text-align:center;">Make-up Days</th>
                <th style="text-align:center;">Overwork Index</th>
            </tr>
        """
        for _, r in ot_emps.iterrows():
            html += f"""
            <tr>
                <td><strong>{r['full_name_en']}</strong></td>
                <td>{r['team']}</td>
                <td>{r['customer_group']}</td>
                <td style="text-align:center;">{r['total_ot_hours']}h</td>
                <td style="text-align:center;">{int(r['ot_sessions'])}</td>
                <td style="text-align:center;">{int(r['unapproved_late_stays'])}</td>
                <td style="text-align:center;">{int(r['makeup_days'])}</td>
                <td style="text-align:center;">{r['overwork_index']}</td>
            </tr>
            """
        html += "</table>"
        
    html += """
        <h2>4. Leave Taxonomy & Capacity Impact</h2>
    """
    if loss_df.empty:
        html += "<p>ℹ️ <em>No planned or unplanned leave recorded in this period.</em></p>"
    else:
        html += """
        <table>
            <tr>
                <th>Customer / Project Group</th>
                <th style="text-align:center;">Date</th>
                <th>Absent Employees</th>
                <th style="text-align:center;">Man-Days Lost</th>
                <th style="text-align:center;">Status</th>
            </tr>
        """
        for _, r in loss_df.sort_values(by='date').iterrows():
            status_html = "<span class='badge-high'>BOTTLENECK</span>" if r['is_bottleneck'] else "<span class='badge-low'>Normal</span>"
            html += f"""
            <tr>
                <td><strong>{r['customer_group']}</strong></td>
                <td style="text-align:center;">{r['date'].strftime('%d/%m')}</td>
                <td>{', '.join(r['absent_employees'])}</td>
                <td style="text-align:center;">{r['total_leave_credit']}</td>
                <td style="text-align:center;">{status_html}</td>
            </tr>
            """
        html += "</table>"
        
    html += """
        <h2>5. Prioritized People & Culture Action Triggers</h2>
    """
    if not anomalies:
        html += "<p>✅ <em>No critical governance alerts. Workforce operating smoothly within standard thresholds.</em></p>"
    else:
        html += """
        <table>
            <tr>
                <th style="text-align:center;">Priority</th>
                <th>Issue Category</th>
                <th>Subject</th>
                <th>Client Group</th>
                <th>Trigger Condition</th>
                <th>Recommended HR Action</th>
            </tr>
        """
        for a in anomalies:
            sev_class = 'badge-high' if a['severity'] == 'HIGH' else ('badge-med' if a['severity'] == 'MEDIUM' else 'badge-low')
            html += f"""
            <tr>
                <td style="text-align:center;"><span class="{sev_class}">{a['severity']}</span></td>
                <td><strong>{a['category']}</strong></td>
                <td>{a['full_name_en']}</td>
                <td>{a['customer_group']}</td>
                <td>{a['trigger_condition']}</td>
                <td><em>{a['recommended_action']}</em></td>
            </tr>
            """
        html += "</table>"
        
    html += """
        <h2>6. Weekly Request Volume Breakdown</h2>
    """
    if req_sum is None or req_sum.empty:
        html += "<p>ℹ️ <em>No application requests submitted for this evaluation window.</em></p>"
    else:
        html += """
        <table>
            <tr>
                <th>Request Category</th>
                <th style="text-align:center;">Volume</th>
                <th>Operational Metric Detail</th>
            </tr>
        """
        for _, r in req_sum.iterrows():
            html += f"""
            <tr>
                <td><strong>{r['request_type']}</strong></td>
                <td style="text-align:center;">{r['count']}</td>
                <td>{r['metric_detail']}</td>
            </tr>
            """
        html += "</table>"
        
    html += """
        <h2>7. Personnel Distribution by Customer Account</h2>
    """
    if cust_sum is None or cust_sum.empty:
        html += "<p>ℹ️ <em>No customer group allocation available.</em></p>"
    else:
        html += """
        <table>
            <tr>
                <th>Customer Account</th>
                <th style="text-align:center;">Headcount</th>
                <th>Assigned Personnel</th>
                <th style="text-align:center;">Actual Hours</th>
                <th style="text-align:center;">Total OT</th>
                <th style="text-align:center;">Leave Days</th>
                <th style="text-align:center;">Unexcused Lates</th>
                <th style="text-align:center;">Late Stays (&gt;30m)</th>
            </tr>
        """
        for _, r in cust_sum.iterrows():
            html += f"""
            <tr>
                <td><strong>{r['customer_group']}</strong></td>
                <td style="text-align:center;">{r['headcount']}</td>
                <td>{r['employees']}</td>
                <td style="text-align:center;">{r['total_actual_hours']}h</td>
                <td style="text-align:center;">{r['total_ot_hours']}h</td>
                <td style="text-align:center;">{r['total_leave_days']}</td>
                <td style="text-align:center;">{r['unexcused_lates']}</td>
                <td style="text-align:center;">{r['unapproved_late_stays']}</td>
            </tr>
            """
        html += "</table>"
        
    html += """
        <br>
        <p style="font-size: 11px; color: #718096; border-top: 1px solid #e2e8f0; padding-top: 8px;">
            Report generated by <strong>employee-attendance-management</strong> skill.<br>
            Single Source of Truth: MISA Biometric Timesheets &amp; Entities Master Directory.
        </p>
    </body>
    </html>
    """
    return html

def generate_report(date_kw='last_week', output_format='all'):
    start_date, end_date = get_date_range_from_keyword(date_kw)
    print(f"Executing Attendance Quality Audit for: {start_date} to {end_date}")
    
    daily_df = reconcile_daily_attendance(start_date, end_date)
    if daily_df.empty:
        print("Error: No attendance data found for specified range.")
        return None, None
        
    emp_sum = compute_employee_summary(daily_df)
    team_sum = compute_team_summary(emp_sum)
    loss_df = compute_customer_capacity_loss(daily_df)
    anomalies = detect_all_anomalies(daily_df, emp_sum, loss_df)
    req_sum = compute_weekly_request_summary(start_date, end_date)
    cust_sum = compute_customer_personnel_summary(emp_sum)
    
    md_text = build_markdown_report(start_date, end_date, team_sum, emp_sum, daily_df, loss_df, anomalies, req_sum, cust_sum)
    html_text = build_html_report(start_date, end_date, team_sum, emp_sum, daily_df, loss_df, anomalies, req_sum, cust_sum)
    
    html_file = os.path.join(SKILL_DIR, "latest_attendance_report.html")
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_text)
        
    md_file = os.path.join(SKILL_DIR, "latest_attendance_report.md")
    with open(md_file, 'w', encoding='utf-8') as f:
        f.write(md_text)
        
    print(f"Report successfully generated and saved to:")
    print(f"  HTML: {html_file}")
    print(f"  Markdown: {md_file}")
    
    if output_format in ['md', 'all']:
        print("\n" + md_text)
        
    return html_file, md_file

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Generate Attendance Quality Report")
    parser.add_argument('period', nargs='?', default='last_week', help='Period: last_week, this_week, or YYYY-MM-DD:YYYY-MM-DD')
    parser.add_argument('--format', default='all', choices=['html', 'md', 'all'], help='Output format')
    args = parser.parse_args()
    
    generate_report(args.period, args.format)
