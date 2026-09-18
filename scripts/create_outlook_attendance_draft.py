# -*- coding: utf-8 -*-
"""
Outlook Email Draft Generator for Attendance Quality Report
Adheres strictly to the MANDATORY SAFETY POLICY: 100% DRAFT-ONLY & HUMAN-IN-THE-LOOP.
Saves the email to Outlook Drafts (.Save()) and displays it (.Display()) for human inspection.
STRICTLY NO .Send() EXECUTION.
"""

import sys
import os
import io
import argparse
import win32com.client

# Ensure UTF-8 output
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from .config import DEFAULT_PC_MANAGER_EMAIL, DEFAULT_SENDER_EMAIL, SKILL_DIR
from .generate_weekly_report import generate_report, get_date_range_from_keyword

def create_attendance_email_draft(date_kw='last_week', to_email=DEFAULT_PC_MANAGER_EMAIL, cc_email=""):
    """
    Generates report and populates an Outlook Email Draft.
    """
    start_date, end_date = get_date_range_from_keyword(date_kw)
    period_str = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"
    
    html_file, md_file = generate_report(date_kw=date_kw, output_format='none')
    if not html_file or not os.path.exists(html_file):
        print("Error: Could not generate HTML report.")
        return False
        
    with open(html_file, 'r', encoding='utf-8') as f:
        html_body = f.read()
        
    subject = f"MyTeam _ Weekly Workforce Attendance Quality Report _ {period_str}"
    
    print(f"\n=======================================================")
    print(f"Creating Outlook Email Draft for P&C Management...")
    print(f"  To: {to_email}")
    if cc_email:
        print(f"  CC: {cc_email}")
    print(f"  Subject: {subject}")
    print(f"=======================================================")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to_email
        if cc_email:
            mail.CC = cc_email
        mail.Subject = subject
        mail.HTMLBody = html_body
        
        # MANDATORY SAFETY POLICY: Save and Display only. NO .Send()
        mail.Save()
        mail.Display()
        
        print("\n✅ Successfully created email draft in Outlook!")
        print("🛑 MANDATORY SAFETY POLICY ENFORCED: 100% DRAFT ONLY.")
        print("👉 Outlook window is now displayed for your review and manual sending.")
        return True
    except Exception as e:
        print(f"❌ Error communicating with Outlook COM interface: {e}")
        print("Please ensure Microsoft Outlook is running.")
        return False

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Create Outlook Draft for Attendance Report")
    parser.add_argument('period', nargs='?', default='last_week', help='Period: last_week, this_week, or YYYY-MM-DD:YYYY-MM-DD')
    parser.add_argument('--to', default=DEFAULT_PC_MANAGER_EMAIL, help='Recipient email address')
    parser.add_argument('--cc', default='', help='CC email address')
    args = parser.parse_args()
    
    create_attendance_email_draft(args.period, args.to, args.cc)
