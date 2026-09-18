# -*- coding: utf-8 -*-
"""
All-in-one CLI Orchestrator for Employee Attendance Management Skill
Provides unified commands for attendance reporting, anomaly screening,
drill-downs by team/employee, and Outlook email draft generation.
"""

import sys
import os
import io
import argparse

# Ensure UTF-8 output
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Ensure scripts directory and parent skill directory are in sys.path
_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
_SKILL_DIR = os.path.dirname(_SCRIPTS_DIR)
for _p in [_SCRIPTS_DIR, _SKILL_DIR]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

try:
    from .config import DEFAULT_PC_MANAGER_EMAIL, DEFAULT_SENDER_EMAIL
    from .generate_weekly_report import generate_report, get_date_range_from_keyword
    from .core_engine import reconcile_daily_attendance, compute_employee_summary, compute_team_summary, compute_customer_capacity_loss
    from .detect_anomalies import detect_all_anomalies
    from .create_outlook_attendance_draft import create_attendance_email_draft
except (ImportError, ValueError):
    from config import DEFAULT_PC_MANAGER_EMAIL, DEFAULT_SENDER_EMAIL
    from generate_weekly_report import generate_report, get_date_range_from_keyword
    from core_engine import reconcile_daily_attendance, compute_employee_summary, compute_team_summary, compute_customer_capacity_loss
    from detect_anomalies import detect_all_anomalies
    from create_outlook_attendance_draft import create_attendance_email_draft

def show_anomalies_only(period_kw='last_week'):
    """Prints only the prioritized anomalies table to console."""
    start_date, end_date = get_date_range_from_keyword(period_kw)
    print(f"\n=======================================================")
    print(f"🔍 Priority Governance Alerts & Anomalies ({start_date} to {end_date})")
    print(f"=======================================================")
    
    daily_df = reconcile_daily_attendance(start_date, end_date)
    if daily_df.empty:
        print("No attendance records found.")
        return
        
    emp_sum = compute_employee_summary(daily_df)
    loss_df = compute_customer_capacity_loss(daily_df)
    anomalies = detect_all_anomalies(daily_df, emp_sum, loss_df)
    
    if not anomalies:
        print("✅ No critical governance alerts detected.")
        return
        
    for i, a in enumerate(anomalies, 1):
        sev_icon = "🔴" if a['severity'] == 'HIGH' else ("🟡" if a['severity'] == 'MEDIUM' else "⚪")
        print(f"\n{i}. {sev_icon} [{a['severity']}] {a['category']}")
        print(f"   Employee/Subject: {a['full_name_en']} ({a['emp_id']}) | Client: {a['customer_group']} | Team: {a['team']}")
        print(f"   Trigger: {a['trigger_condition']}")
        print(f"   Recommended Action: {a['recommended_action']}")
    print("")

def drilldown_employee(emp_identifier, period_kw='last_week'):
    """Drills down into individual employee attendance daily logs."""
    start_date, end_date = get_date_range_from_keyword(period_kw)
    daily_df = reconcile_daily_attendance(start_date, end_date)
    if daily_df.empty:
        print("No records found.")
        return
        
    emp_recs = daily_df[
        (daily_df['emp_id'].str.upper() == emp_identifier.upper()) |
        (daily_df['full_name_en'].str.lower().str.contains(emp_identifier.lower())) |
        (daily_df['emp_name'].str.lower().str.contains(emp_identifier.lower()))
    ]
    
    if emp_recs.empty:
        print(f"No records found matching employee '{emp_identifier}' in period {start_date} to {end_date}.")
        return
        
    full_name = emp_recs['full_name_en'].iloc[0]
    team = emp_recs['team_clean'].iloc[0]
    cust = emp_recs['customer_group'].iloc[0]
    print(f"\n=======================================================")
    print(f"👤 Individual Drill-Down: {full_name} ({emp_recs['emp_id'].iloc[0]})")
    print(f"   Team: {team} | Client: {cust} | Period: {start_date} to {end_date}")
    print(f"=======================================================")
    
    for _, r in emp_recs.sort_values(by='date').iterrows():
        ci_str = r['check_in'] or '-:-'
        co_str = r['check_out'] or '-:-'
        notes = []
        if r['is_late_ci']:
            exc = " (Approved)" if r['is_excused_late'] else " (UNEXCUSED)"
            notes.append(f"Late CI: +{int(r['late_ci_mins'])}m{exc}")
        if r['is_unapproved_late_stay']:
            notes.append(f"Late CO: +{int(r['late_co_mins'])}m (No OT)")
        elif r['has_approved_ot']:
            notes.append(f"OT: {r['ot_hours']}h ({r['ot_type']})")
        if r['is_makeup_time']:
            notes.append("Make-up Time")
        if r['leave_credit'] > 0:
            notes.append(f"Leave: {r['leave_credit']} day ({r['leave_type']})")
        if r['is_manual_override']:
            notes.append("Manual Override ('Cập nhật công')")
            
        note_str = " | ".join(notes) if notes else "Compliant"
        print(f"  {r['date'].strftime('%a %d/%m')}: Shift {r['shift_code']} ({r['shift_start']}-{r['shift_end']}) | Punch: {ci_str} - {co_str} | Worked: {r['actual_hours']}h | {note_str}")
    print("")

def launch_dashboard(refresh=True, deploy=True):
    """
    Refreshes dashboard dataset (data.js and data.json), automatically deploys
    to GitHub Pages, and opens the workforce attendance dashboard in default browser.
    """
    from pathlib import Path
    import webbrowser
    skill_dir = Path(__file__).parent.parent.resolve()
    reports_html = skill_dir / "reports" / "index.html"
    
    if refresh:
        print("\n[+] Regenerating dashboard analytics dataset from raw timesheets & Master Data entities...")
        try:
            from .generate_dashboard_data import main as regen_dash
        except (ImportError, ValueError):
            from generate_dashboard_data import main as regen_dash
        regen_dash(deploy=deploy)
    elif deploy:
        try:
            from .generate_dashboard_data import deploy_to_github
        except (ImportError, ValueError):
            from generate_dashboard_data import deploy_to_github
        deploy_to_github()
        
    if not reports_html.exists():
        print(f"Error: Dashboard HTML not found at {reports_html}")
        return
        
    abs_html = os.path.abspath(reports_html)
    url = f"file:///{abs_html.replace(os.sep, '/')}"
    print(f"\n🚀 Opening Workforce Attendance Dashboard:\n  Local : {url}")
    print(f"  Live  : https://tamdoan2802.github.io/PC_attendance_management/\n")
    try:
        webbrowser.open(url)
    except Exception as e:
        print(f"Could not automatically launch browser: {e}")

def main():
    parser = argparse.ArgumentParser(description="Employee Attendance Management CLI")
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # 1. report
    p_rep = subparsers.add_parser('report', help='Generate full attendance quality report')
    p_rep.add_argument('period', nargs='?', default='last_week', help='last_week, this_week, or YYYY-MM-DD:YYYY-MM-DD')
    p_rep.add_argument('--format', default='all', choices=['html', 'md', 'all'])
    
    # 2. anomalies
    p_anom = subparsers.add_parser('anomalies', help='Show prioritized governance alerts only')
    p_anom.add_argument('period', nargs='?', default='last_week', help='last_week, this_week, or YYYY-MM-DD:YYYY-MM-DD')
    
    # 3. employee drilldown
    p_emp = subparsers.add_parser('employee', help='Drilldown into single employee attendance')
    p_emp.add_argument('emp_id', help='Employee ID or English name')
    p_emp.add_argument('period', nargs='?', default='last_week')
    
    # 4. draft email
    p_draft = subparsers.add_parser('draft', help='Create Outlook email draft for P&C Manager')
    p_draft.add_argument('period', nargs='?', default='last_week')
    p_draft.add_argument('--to', default=DEFAULT_PC_MANAGER_EMAIL)
    p_draft.add_argument('--cc', default="")
    
    # 5. dashboard
    p_dash = subparsers.add_parser('dashboard', help='Regenerate dashboard data and launch Workforce Attendance Dashboard in browser')
    p_dash.add_argument('--no-refresh', action='store_true', help='Skip data regeneration and open dashboard directly')
    p_dash.add_argument('--no-deploy', action='store_true', help='Skip automatic deploy to GitHub Pages')

    # 6. deploy
    p_dep = subparsers.add_parser('deploy', help='Deploy current dashboard directly to GitHub Pages')
    
    args = parser.parse_args()
    
    if args.command == 'report':
        generate_report(args.period, args.format)
    elif args.command == 'anomalies':
        show_anomalies_only(args.period)
    elif args.command == 'employee':
        drilldown_employee(args.emp_id, args.period)
    elif args.command == 'draft':
        create_attendance_email_draft(args.period, args.to, args.cc)
    elif args.command == 'dashboard':
        launch_dashboard(refresh=not args.no_refresh, deploy=not args.no_deploy)
    elif args.command == 'deploy':
        try:
            from .generate_dashboard_data import deploy_to_github
        except (ImportError, ValueError):
            from generate_dashboard_data import deploy_to_github
        deploy_to_github()
    else:
        # Default action: run weekly report
        generate_report('last_week', 'all')

if __name__ == '__main__':
    main()
