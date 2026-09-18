# -*- coding: utf-8 -*-
"""
Configuration module for employee-attendance-management skill.
Establishes paths, standard shifts catalog, governance rules, and threshold constants.
"""

import os
import datetime

# --- DIRECTORY PATHS ---
SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE_DIR = os.path.abspath(os.path.join(SKILL_DIR, "..", "..", ".."))

DATA_DIR = os.path.join(WORKSPACE_DIR, "data", "Attendance")
TIMESHEET_DIR = os.path.join(DATA_DIR, "Timesheet")
ENTITIES_DIR = os.path.join(WORKSPACE_DIR, "entities")
SKILL_DATA_DIR = os.path.join(SKILL_DIR, "data")
HOLIDAY_FILE = os.path.join(DATA_DIR, "vietnamese_holidays.csv")

# Request file paths
REQ_LATE_EARLY = os.path.join(DATA_DIR, "Đăng ký đi muộn, về sớm_Tất cả đơn vị.xlsx")
REQ_OVERTIME = os.path.join(DATA_DIR, "Đơn_đăng_ký_làm_thêm_Tất cả đơn vị.xlsx")
REQ_WFH = os.path.join(DATA_DIR, "Dang_ky_lam_viec_tu_xa_Tất cả đơn vị.xlsx")
REQ_SHIFT_CHANGE = os.path.join(DATA_DIR, "De_nghi_doi_ca_Tất cả đơn vị.xlsx")
REQ_BUSINESS_TRIP = os.path.join(DATA_DIR, "De_nghi_di_cong_tac_Tất cả đơn vị.xlsx")
REQ_LEAVE = os.path.join(DATA_DIR, "Don_xin_nghi_Tất cả đơn vị.xlsx")

# Entities paths
ROUTING_MAP_FILE = os.path.join(ENTITIES_DIR, "notification_routing_map.md")
EMPLOYEES_FILE = os.path.join(ENTITIES_DIR, "employees.md")
TEAMS_FILE = os.path.join(ENTITIES_DIR, "teams.md")
CUSTOMERS_FILE = os.path.join(ENTITIES_DIR, "customers.md")

# --- MASTER SHIFTS CATALOG ---
# Code: (Start_Time_Str, End_Time_Str, Break_Hours, Std_Hours)
SHIFT_CATALOG = {
    "WT": ("08:00", "17:00", 1.0, 8.0),
    "BS": ("07:00", "16:00", 1.0, 8.0),
    "BS01": ("07:30", "16:30", 1.0, 8.0),
    "WE": ("06:30", "15:30", 1.0, 8.0),
    "WE05": ("06:45", "15:45", 1.0, 8.0),
    "WE06": ("07:30", "16:00", 0.5, 8.0),
    "WT01-WFH": ("07:30", "16:30", 1.0, 8.0),
    "WE07-WFH": ("05:30", "09:00", 0.0, 3.5),
    "WE08-WFH": ("05:30", "08:30", 0.0, 3.0),
    "WE08": ("09:30", "15:30", 0.0, 6.0),
}

# Mapping MISA request labels to canonical Short Codes
SHIFT_LABEL_MAP = {
    "WorkTime": "WT",
    "Branch Support": "BS",
    "Branch Support 1": "BS01",
    "BS(07H30-16H30)": "BS01",
    "Work Early": "WE",
    "Work Early 05": "WE05",
    "Work Early 06": "WE06",
    "WorkTime WFH 01": "WT01-WFH",
    "Bulk Group (6AM-9AM)": "WE08-WFH",
    "Bulk Group (6AM-9AM) - WFH": "WE08-WFH",
    "Bulk Group (9H30-15H30)": "WE08",
}

# Excluded shifts / populations
DEPRECATED_SHIFTS = ["WL01"]
EXCLUDED_POPULATION_PREFIXES = ["MTVN-TT-"]  # Interns (IA shift)

EXCLUDED_TOTAL = [
    "MTVN0059",  # Adrian (Director - excluded from everything)
    "MTVN0062",  # Chau Ha (P&C Manager - excluded from everything)
]
EXCLUDED_ATTENDANCE_METRICS = [
    "MTVN0037",  # Toan Mai (Logistics Coordinator - attendance punch/hours excluded, ALL requests retained)
    "MTVN0066",  # Huyen Le (Logistics Support Officer - attendance punch/hours excluded, ALL requests retained)
]
EXCLUDED_EMP_IDS = EXCLUDED_TOTAL + EXCLUDED_ATTENDANCE_METRICS

# --- GOVERNANCE THRESHOLDS & POLICY LIMITS ---
CHECKIN_GRACE_MINUTES = 5           # <= 5m is compliant
CHRONIC_LATENESS_MINUTES = 30       # > 30m unexcused is severe
CHRONIC_LATENESS_MAX_MONTH = 2      # >= 2 times unexcused/month -> flag
EXCUSED_FREQ_MAX_WEEK = 2           # >= 2 approved requests/week -> flag
MANUAL_OVERRIDE_MAX_MONTH = 1       # >= 2 times/month -> flag
BURNOUT_CONSECUTIVE_DAYS = 3        # >= 3 consecutive days late CO > 30m without OT -> flag
BURNOUT_LATE_CO_MINUTES = 30

# Overtime Tiers
OT_TIER_1_MAX = 60                  # <= 60m: Tier 1 (Light)
OT_TIER_2_MAX = 120                 # 61 - 120m: Tier 2 (Moderate)
                                    # > 120m: Tier 3 (Heavy/Red)

# Leave Notice Classification (Updated per User Rule)
NOTICE_SUDDEN_MAX_DAYS = 1          # <= 1 day (Trước 01 ngày)
NOTICE_SHORT_MAX_DAYS = 7           # 2 - 7 days (Từ 2 đến 7 ngày)
NOTICE_PROMPT_MAX_DAYS = 29         # 8 - 29 days
                                    # >= 30 days: Planned/Normal

# Default Email Recipients
DEFAULT_PC_MANAGER_EMAIL = "ChauH@myteamsolution.com.au"
DEFAULT_SENDER_EMAIL = "tam.doan@myteamsolution.com.vn"
