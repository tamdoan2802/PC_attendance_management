#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_dashboard_data.py
==========================
Reads HR_Fact_Attendance.xlsx, computes all KPIs and detail tables,
then injects a fresh DASH_DATA JSON block into Attedance_dashboard.html.

Run from the reports/ folder:
    python generate_dashboard_data.py

Dependencies:  pip install pandas openpyxl

Columns used per sheet — see kpi_formular.md for full metric logic.
Excluded employees: MTVN0059 (Adrian), MTVN0062 (Chau Ha).
"""

import os
import glob
import sys
import re
import json

# Force UTF-8 output on Windows (avoids cp1252 UnicodeEncodeError)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")
import warnings
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

warnings.filterwarnings("ignore")

try:
    import pandas as pd
except ImportError:
    sys.exit("ERROR: pandas not installed. Run: pip install pandas openpyxl")

# ═══════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════
THIS_DIR   = Path(__file__).parent.resolve()
if str(THIS_DIR) not in sys.path:
    sys.path.insert(0, str(THIS_DIR))
WORKSPACE_DIR = THIS_DIR.parent.parent.parent.parent
DATA_ATT_DIR = WORKSPACE_DIR / "data" / "Attendance"
ENTITIES_DIR = WORKSPACE_DIR / "entities"
TIMESHEET_DIR = DATA_ATT_DIR / "Timesheet"

DATA_JSON_PATH = THIS_DIR.parent / "references" / "data.json"
DATA_JS_PATH = THIS_DIR.parent / "references" / "data.js"

EXCLUDED_TOTAL = {
    "MTVN0059",  # Adrian (Director)
    "MTVN0062",  # Chau Ha (P&C Manager)
}
EXCLUDED_ATTENDANCE = {
    "MTVN0037",  # Toan Mai (Logistics Coordinator - attendance excluded, leave retained)
    "MTVN0066",  # Huyen Le (Logistics Support Officer - attendance excluded, leave retained)
}
EXCLUDED_ALL_ATTENDANCE = EXCLUDED_TOTAL | EXCLUDED_ATTENDANCE
LATE_CI_THRESHOLD_MINS  = 5
OVERWORK_THRESHOLD_HRS  = 1.5
EXTENSIVE_HOURS_THRESH  = 9.5    # working hours > this = extensive
LATE_CO_FLAG_MINS       = 90     # checkout mins past standard = incident
PLANNED_NOTICE_DAYS     = 30
URGENT_NOTICE_DAYS      = 2
TRAILING_WEEKS          = 4

# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════

def safe_float(val, default=0.0):
    try:
        v = float(val)
        return default if (v != v) else v          # nan check
    except (TypeError, ValueError):
        return default

def safe_int(val, default=0):
    try:
        v = int(float(val))
        return v
    except (TypeError, ValueError):
        return default

def fmt_date(val):
    """Return dd/mm/yyyy string or empty."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%d/%m/%Y")
    return str(val)

def fmt_time(val):
    """Return HH:MM string from a time/datetime object."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    try:
        if isinstance(val, datetime):
            return val.strftime("%H:%M")
        if hasattr(val, "hour"):
            return f"{val.hour:02d}:{val.minute:02d}"
    except Exception:
        pass
    return str(val)

def parse_week_mon_fri(week_period_str):
    """'13/07/2026 - 17/07/2026' → (datetime, datetime)."""
    parts = str(week_period_str).split(" - ")
    mon = datetime.strptime(parts[0].strip(), "%d/%m/%Y")
    fri = datetime.strptime(parts[1].strip(), "%d/%m/%Y")
    return mon, fri

def week_range_label(week_period_str):
    """'13/07/2026 - 17/07/2026' → 'Jul 13, 26'  (Monday of the week)."""
    try:
        mon, _ = parse_week_mon_fri(week_period_str)
        return mon.strftime("%b %-d, %y") if sys.platform != "win32" else \
               mon.strftime("%b %d, %y").replace(" 0", " ")
    except Exception:
        return week_period_str

def week_label(week_period_str):
    try:
        mon, fri = parse_week_mon_fri(week_period_str)
        if mon.month == fri.month:
            return f"Week {mon.strftime('%b %d')} - {fri.strftime('%d')}"
        else:
            return f"Week {mon.strftime('%b %d')} - {fri.strftime('%b %d')}"
    except Exception:
        return week_period_str

def pct(num, denom, decimals=1):
    if denom == 0:
        return 0
    return round(100.0 * num / denom, decimals)

def notice_category(notice_before_days):
    """Classify leave notice from (Leave_From - Submit_Date) in days."""
    try:
        n = float(notice_before_days)
    except (TypeError, ValueError):
        return "Unknown"
    if n > PLANNED_NOTICE_DAYS:
        return "Planned"
    elif n > URGENT_NOTICE_DAYS:
        return "Unplanned"
    else:
        return "Urgent"

def get_leave_days_in_week(lf, lt, total_days, mon, fri):
    """Calculate the number of leave weekdays falling within [mon, fri]."""
    if pd.isna(lf) or pd.isna(lt) or pd.isna(mon) or pd.isna(fri):
        return 0.0
    try:
        start = max(pd.to_datetime(lf).normalize(), pd.to_datetime(mon).normalize())
        end = min(pd.to_datetime(lt).normalize(), pd.to_datetime(fri).normalize())
        if start > end:
            return 0.0
        bdays = len(pd.bdate_range(start, end))
        if bdays == 0:
            return 0.0
        total_d = safe_float(total_days, default=float(bdays))
        if bdays == 1 and total_d < 1.0:
            return total_d
        return float(bdays)
    except Exception:
        return 0.0

# ═══════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════


def load_master_employees_from_entities(legacy_emp_df=None):
    """
    Constructs the canonical employee dimension directly from entities/:
    - entities/notification_routing_map.md
    - entities/employees.md
    - entities/teams.md
    - entities/customers.md
    Guarantees 100% complete personnel coverage (including MTVN0097 - Trung Vo).
    """
    master_df = pd.DataFrame()
    routing_map_file = ENTITIES_DIR / "notification_routing_map.md"
    if routing_map_file.exists():
        try:
            try:
                from .load_requests import load_master_entities
                master_df = load_master_entities(str(routing_map_file))
            except Exception:
                try:
                    from load_requests import load_master_entities
                    master_df = load_master_entities(str(routing_map_file))
                except Exception:
                    try:
                        from scripts.load_requests import load_master_entities
                        master_df = load_master_entities(str(routing_map_file))
                    except Exception:
                        pass
        except Exception:
            pass

    emp_meta = {}
    emp_md_path = ENTITIES_DIR / "employees.md"
    if emp_md_path.exists():
        with open(emp_md_path, "r", encoding="utf-8") as f:
            for line in f:
                if not line.strip().startswith("|"): continue
                parts = [p.strip() for p in line.strip().split("|")[1:-1]]
                if not parts or all(set(p).issubset({'-', ':', ' '}) for p in parts) or "---" in parts[0] or "Employee ID" in parts[0]: continue
                clean = [p.replace("`", "").replace("**", "").strip() for p in parts]
                if len(clean) >= 9 and clean[0].startswith("MTVN"):
                    emp_meta[clean[0]] = {
                        "FullNameVN": clean[1],
                        "FullNameEN": clean[2],
                        "Location": clean[5] if len(clean) > 5 else "Da Nang",
                        "Position": clean[9] if len(clean) > 9 else "",
                        "Status": clean[8] if len(clean) > 8 else "Active"
                    }

    master_dict = master_df.set_index("emp_id").to_dict("index") if not master_df.empty else {}
    legacy_dict = legacy_emp_df.set_index("EmployeeID").to_dict("index") if legacy_emp_df is not None and not legacy_emp_df.empty else {}
    
    all_eids = {eid for eid in (set(master_dict.keys()).union(emp_meta.keys()).union(legacy_dict.keys())) - EXCLUDED_TOTAL if eid and eid.startswith("MTVN")}

    emp_rows = []
    for eid in sorted(all_eids):
        m = master_dict.get(eid, {})
        e = emp_meta.get(eid, {})
        l = legacy_dict.get(eid, {})
        
        fn_en = m.get("full_name_en") or e.get("FullNameEN") or l.get("FullNameEN") or ""
        fn_vn = m.get("emp_name_vn") or e.get("FullNameVN") or l.get("FullNameVN") or ""
        team = m.get("team") or l.get("Team") or ""
        client = m.get("customer_group") or l.get("Client") or ""
        loc = e.get("Location") or l.get("Location") or "Da Nang"
        pos = e.get("Position") or l.get("Position") or ""
        role = l.get("Role in Team") or "Member"
        
        stat = e.get("Status", "Active")
        if any(k in str(stat).lower() for k in ["resigned", "terminated", "inactive"]):
            active = 0
        else:
            active = safe_int(l.get("IsActive", 1), default=1)
            
        emp_rows.append({
            "EmployeeID": eid,
            "FullNameVN": fn_vn,
            "FullNameEN": fn_en,
            "Team": team,
            "Department": team,
            "Role in Team": role,
            "Client": client,
            "Location": loc,
            "Position": pos,
            "IsActive": active
        })
    df_emp = pd.DataFrame(emp_rows)
    print(f"  OK Master DIM_Employee built from entities ({len(df_emp):,} employees, {len(df_emp[df_emp['IsActive'] == 1]):,} active)")
    return df_emp

def load_live_requests(raw):
    """
    Enriches or supersedes requests in raw with live MISA Excel files from data/Attendance/:
    - Don_xin_nghi_Tất cả đơn vị.xlsx (Req_Leave)
    - Đơn_đăng_ký_làm_thêm_Tất cả đơn vị.xlsx (Req_OT)
    - Dang_ky_lam_viec_tu_xa_Tất cả đơn vị.xlsx (Req_WFh)
    - Đăng ký đi muộn, về sớm_Tất cả đơn vị.xlsx (Req_LCin&ECout)
    - De_nghi_di_cong_tac_Tất cả đơn vị.xlsx (Req_BusinessTrip)
    - De_nghi_doi_ca_Tất cả đơn vị.xlsx (Req_ShiftChange)
    """
    # 1. Req_Leave
    leave_file = DATA_ATT_DIR / "Don_xin_nghi_Tất cả đơn vị.xlsx"
    if leave_file.exists():
        try:
            df_l = pd.read_excel(leave_file, skiprows=4)
            records = []
            for _, r in df_l.iterrows():
                eid = str(r.get("Mã nhân viên", "")).strip()
                if not eid or eid == "nan": continue
                t_from = pd.to_datetime(r.get("Từ ngày"), errors="coerce")
                t_to = pd.to_datetime(r.get("Đến ngày"), errors="coerce")
                sub_dt = pd.to_datetime(r.get("Ngày nộp đơn"), errors="coerce")
                if pd.isna(t_from): continue
                if pd.isna(t_to): t_to = t_from
                if t_from > t_to: t_from, t_to = t_to, t_from
                
                l_type = str(r.get("Loại nghỉ", "Nghỉ phép")).strip()
                if "phép" in l_type.lower(): l_mapped = "Annual Leave"
                elif "không lương" in l_type.lower(): l_mapped = "Unpaid Leave"
                elif "bù" in l_type.lower(): l_mapped = "Compensatory Leave"
                elif any(k in l_type.lower() for k in ["ốm", "bhxh", "con ốm"]): l_mapped = "Sick/Family Care"
                else: l_mapped = l_type
                
                status = str(r.get("Trạng thái", "")).strip()
                status_flag = 1 if status == "Đã duyệt" else 0
                days = float(r.get("Số ngày nghỉ", 1.0) or 1.0)
                
                lead_days = (t_from.date() - sub_dt.date()).days if pd.notna(sub_dt) else 30
                if lead_days > 30: notice_cat = "Planned"
                elif lead_days > 2: notice_cat = "Unplanned"
                else: notice_cat = "Urgent"
                
                records.append({
                    "Employee_ID": eid,
                    "Employee_Name": str(r.get("Người nộp đơn", "")).strip(),
                    "Leave_From": t_from,
                    "Leave_To": t_to,
                    "Submit_Date": sub_dt,
                    "Leave_Days": days,
                    "Leave Days in Week": days,
                    "Leave_Type": l_type,
                    "Leave_Type_Mapped": l_mapped,
                    "Reason": str(r.get("Lý do nghỉ", "")).strip(),
                    "Status": status,
                    "Status_Flag": status_flag,
                    "notice_category": notice_cat,
                    "Has_MonFri": True
                })
            if records:
                raw["Req_Leave"] = pd.DataFrame(records)
                print(f"  OK Live Req_Leave refreshed from {leave_file.name} ({len(records):,} rows)")
        except Exception as e:
            print(f"  WARN: Failed loading live Req_Leave: {e}")

    # 2. Req_OT
    ot_file = DATA_ATT_DIR / "Đơn_đăng_ký_làm_thêm_Tất cả đơn vị.xlsx"
    if ot_file.exists():
        try:
            df_ot = pd.read_excel(ot_file, skiprows=4)
            records = []
            for _, r in df_ot.iterrows():
                emp_sub = r.get("Mã nhân viên làm thêm")
                emp_subm = r.get("Mã nhân viên")
                eid = str(emp_sub).strip() if pd.notna(emp_sub) and str(emp_sub).strip() != "nan" else str(emp_subm).strip()
                if not eid or eid == "nan": continue
                ot_from = pd.to_datetime(r.get("Làm thêm từ"), errors="coerce")
                ot_to = pd.to_datetime(r.get("Làm thêm đến"), errors="coerce")
                if pd.isna(ot_from): continue
                if pd.isna(ot_to): ot_to = ot_from
                
                ot_hrs = float(r.get("Số giờ làm thêm", 0) or 0)
                status = str(r.get("Trạng thái", "")).strip()
                status_flag = 1 if status == "Đã duyệt" else 0
                timing = str(r.get("Thời điểm làm thêm", "Sau ca làm việc")).strip()
                
                records.append({
                    "Employee_ID": eid,
                    "Employee_Name": str(r.get("Người nộp đơn", "")).strip(),
                    "OT_From": ot_from,
                    "OT_To": ot_to,
                    "OT Date": ot_from.normalize(),
                    "_ot_date": ot_from.normalize(),
                    "OT_Hours": ot_hrs,
                    "OT_Timing": timing,
                    "Reason": str(r.get("Lý do làm thêm", "")).strip(),
                    "Status": status,
                    "Status_Flag": status_flag
                })
            if records:
                raw["Req_OT"] = pd.DataFrame(records)
                print(f"  OK Live Req_OT refreshed from {ot_file.name} ({len(records):,} rows)")
        except Exception as e:
            print(f"  WARN: Failed loading live Req_OT: {e}")

    # 3. Req_WFh
    wfh_file = DATA_ATT_DIR / "Dang_ky_lam_viec_tu_xa_Tất cả đơn vị.xlsx"
    if wfh_file.exists():
        try:
            df_wfh = pd.read_excel(wfh_file, skiprows=3)
            records = []
            for _, r in df_wfh.iterrows():
                eid = str(r.get("Mã nhân viên", "")).strip()
                if not eid or eid == "nan": continue
                t_from = pd.to_datetime(r.get("Từ ngày"), errors="coerce")
                t_to = pd.to_datetime(r.get("Đến ngày"), errors="coerce")
                if pd.isna(t_from): continue
                if pd.isna(t_to): t_to = t_from
                status = str(r.get("Trạng thái", "")).strip()
                status_flag = 1 if status == "Đã duyệt" else 0
                records.append({
                    "Employee_ID": eid,
                    "Employee_Name": str(r.get("Người nộp đơn", "")).strip(),
                    "WFH_From": t_from,
                    "WFH_To": t_to,
                    "Status": status,
                    "Status_Flag": status_flag,
                    "Reason": str(r.get("Lý do làm việc từ xa", "")).strip()
                })
            if records:
                raw["Req_WFh"] = pd.DataFrame(records)
                print(f"  OK Live Req_WFh refreshed from {wfh_file.name} ({len(records):,} rows)")
        except Exception as e:
            print(f"  WARN: Failed loading live Req_WFh: {e}")

    # 4. Req_LCin&ECout
    lc_file = DATA_ATT_DIR / "Đăng ký đi muộn, về sớm_Tất cả đơn vị.xlsx"
    if lc_file.exists():
        try:
            df_lc = pd.read_excel(lc_file, skiprows=4)
            records = []
            for _, r in df_lc.iterrows():
                eid = str(r.get("Mã nhân viên", "")).strip()
                if not eid or eid == "nan": continue
                t_from = pd.to_datetime(r.get("Từ ngày"), errors="coerce")
                if pd.isna(t_from): continue
                status = str(r.get("Trạng thái", "")).strip()
                status_flag = 1 if status == "Đã duyệt" else 0
                m_late = float(r.get("Đi muộn đầu ca (phút)", 0) or 0) + float(r.get("Đi muộn giữa ca (phút)", 0) or 0)
                m_early = float(r.get("Về sớm giữa ca (phút)", 0) or 0) + float(r.get("Về sớm cuối ca (phút)", 0) or 0)
                records.append({
                    "Employee_ID": eid,
                    "Employee_Name": str(r.get("Người nộp đơn", "")).strip(),
                    "Apply_From": t_from,
                    "Minutes": m_late + m_early,
                    "CI_Category_Mapped": "Late" if m_late > 0 else "On Time",
                    "CO_Category_Mapped": "Early" if m_early > 0 else "On Time",
                    "Reason_Detail": str(r.get("Lý do đi muộn, về sớm", "")).strip(),
                    "Reason_Group": str(r.get("Nhóm lý do", "")).strip(),
                    "Status": status,
                    "Status_Flag": status_flag
                })
            if records:
                raw["Req_LCin&ECout"] = pd.DataFrame(records)
                print(f"  OK Live Req_LCin&ECout refreshed from {lc_file.name} ({len(records):,} rows)")
        except Exception as e:
            print(f"  WARN: Failed loading live Req_LCin&ECout: {e}")

    # 5. Req_BusinessTrip
    trip_file = DATA_ATT_DIR / "De_nghi_di_cong_tac_Tất cả đơn vị.xlsx"
    if trip_file.exists():
        try:
            df_tr = pd.read_excel(trip_file, skiprows=4)
            records = []
            for _, r in df_tr.iterrows():
                eid = str(r.get("Mã nhân viên", "")).strip()
                if not eid or eid == "nan": continue
                t_from = pd.to_datetime(r.get("Từ ngày"), errors="coerce")
                t_to = pd.to_datetime(r.get("Đến ngày"), errors="coerce")
                if pd.isna(t_from): continue
                if pd.isna(t_to): t_to = t_from
                status = str(r.get("Trạng thái", "")).strip()
                status_flag = 1 if status == "Đã duyệt" else 0
                records.append({
                    "Employee_ID": eid,
                    "Employee_Name": str(r.get("Người nộp đơn", "")).strip(),
                    "Trip_From": t_from,
                    "Trip_To": t_to,
                    "Trip_Days": float(r.get("Số ngày đi công tác", 1.0) or 1.0),
                    "Destination": str(r.get("Địa điểm công tác", "")).strip(),
                    "Purpose": str(r.get("Mục đích công tác", "")).strip(),
                    "Status": status,
                    "Status_Flag": status_flag
                })
            if records:
                raw["Req_BusinessTrip"] = pd.DataFrame(records)
                print(f"  OK Live Req_BusinessTrip refreshed from {trip_file.name} ({len(records):,} rows)")
        except Exception as e:
            print(f"  WARN: Failed loading live Req_BusinessTrip: {e}")

    # 6. Req_ShiftChange
    sc_file = DATA_ATT_DIR / "De_nghi_doi_ca_Tất cả đơn vị.xlsx"
    if sc_file.exists():
        try:
            df_sc = pd.read_excel(sc_file, skiprows=4)
            records = []
            for _, r in df_sc.iterrows():
                eid = str(r.get("Mã nhân viên", "")).strip()
                if not eid or eid == "nan": continue
                w_date = pd.to_datetime(r.get("Ngày làm việc"), errors="coerce")
                if pd.isna(w_date): continue
                status = str(r.get("Trạng thái", "")).strip()
                status_flag = 1 if status == "Đã duyệt" else 0
                records.append({
                    "Employee_ID": eid,
                    "Employee_Name": str(r.get("Người nộp đơn", "")).strip(),
                    "Work_Date": w_date,
                    "Shift_Code_Old": str(r.get("Ca hiện tại", "")).strip(),
                    "Shift_Code_New": str(r.get("Ca đăng ký đổi", "")).strip(),
                    "Reason": str(r.get("Lý do đổi ca", "")).strip(),
                    "Status": status,
                    "Status_Flag": status_flag
                })
            if records:
                raw["Req_ShiftChange"] = pd.DataFrame(records)
                print(f"  OK Live Req_ShiftChange refreshed from {sc_file.name} ({len(records):,} rows)")
        except Exception as e:
            print(f"  WARN: Failed loading live Req_ShiftChange: {e}")

def build_fact_attendance_daily():
    """
    Parses all raw Timesheet Excel files directly from data/Attendance/Timesheet/
    into a standardized FACT_Attendance_Daily DataFrame.
    Completely eliminates any dependency on static HR_Fact_Attendance.xlsx!
    """
    try:
        from .parse_timesheet import parse_timesheet_file, load_holidays
    except (ImportError, ValueError):
        try:
            from parse_timesheet import parse_timesheet_file, load_holidays
        except (ImportError, ValueError):
            from scripts.parse_timesheet import parse_timesheet_file, load_holidays

    files = glob.glob(str(TIMESHEET_DIR / "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith(("~$", "HR_Fact"))]
    
    holidays = load_holidays()
    all_recs = []
    for f in sorted(files):
        try:
            recs = parse_timesheet_file(f, holidays=holidays)
            all_recs.extend(recs)
        except Exception as e:
            print(f"  WARN: Failed parsing timesheet {os.path.basename(f)}: {e}")

    df_raw = pd.DataFrame(all_recs)
    if df_raw.empty:
        print("  WARN: No timesheet records parsed from raw files!")
        return pd.DataFrame()

    # Deduplicate on (emp_id, date), keeping the latest parsed record
    df_att = df_raw.drop_duplicates(subset=["emp_id", "date"], keep="last").copy()
    df_att = df_att[~df_att["emp_id"].isin(EXCLUDED_ALL_ATTENDANCE)]

    fact_att = pd.DataFrame()
    fact_att["Employee_ID"] = df_att["emp_id"].astype(str).str.strip()
    fact_att["Employee_Name"] = df_att["emp_name"].astype(str).str.strip()
    fact_att["Position"] = df_att["position"].astype(str).str.strip()
    fact_att["Date_Text"] = pd.to_datetime(df_att["date"])
    fact_att["Day_Num"] = fact_att["Date_Text"].dt.day
    fact_att["Day_Of_Week"] = fact_att["Date_Text"].dt.strftime("%a")
    fact_att["Shift_Code"] = df_att["shift_code"].fillna("")
    fact_att["Working_Days"] = df_att["working_credit"].fillna(0.0)
    fact_att["CheckIn_Time"] = df_att["check_in"].fillna("")
    fact_att["CheckOut_Time"] = df_att["check_out"].fillna("")
    fact_att["DIM_Shift.Start_Time"] = df_att["shift_start"].fillna("")
    fact_att["DIM_Shift.End_Time"] = df_att["shift_end"].fillna("")
    fact_att["Late_CheckIn (mins)"] = df_att["late_ci_mins"].fillna(0.0)
    fact_att["Late_CheckOut(mins)"] = df_att["late_co_mins"].fillna(0.0)
    fact_att["Type of Date"] = df_att["type_of_date"].fillna("Absent")
    fact_att["Số ngày làm việc tiêu chuẩn"] = 22
    fact_att["Số giờ làm việc tiêu chuẩn"] = df_att["std_hours"].fillna(8.0)
    fact_att["Số giờ làm việc thực tế"] = df_att["actual_hours"].fillna(0.0)
    fact_att["Is_Weekend"] = df_att["is_weekend"].astype(int)
    fact_att["Early_CI (mins)"] = df_att["early_ci_mins"].fillna(0.0) if "early_ci_mins" in df_att.columns else 0.0
    fact_att["Late_CI (mins)"] = df_att["late_ci_mins"].fillna(0.0)
    fact_att["Early_CO (mins)"] = df_att["early_co_mins"].fillna(0.0)
    fact_att["Late_CO (mins)"] = df_att["late_co_mins"].fillna(0.0)

    def calc_wp(dt):
        if pd.isna(dt): return ""
        m = dt - pd.Timedelta(days=dt.weekday())
        f = m + pd.Timedelta(days=4)
        return f"{m.strftime('%d/%m/%Y')} - {f.strftime('%d/%m/%Y')}"

    fact_att["Week Period"] = fact_att["Date_Text"].apply(calc_wp)
    fact_att["Delta (Số giờ làm việc thực tế - Số giờ làm việc tiêu chuẩn)"] = (
        fact_att["Số giờ làm việc thực tế"] - fact_att["Số giờ làm việc tiêu chuẩn"]
    )

    print(f"  OK Parsed {len(files)} raw timesheet files -> FACT_Attendance_Daily ({len(fact_att):,} daily records)")
    return fact_att

def load_all_raw_data():
    """
    Loads all data 100% directly from raw files in data/Attendance/ and entities/:
    - Master DIM_Employee: from entities/notification_routing_map.md and entities/employees.md
    - 6 Live Request DataFrames: from raw files in data/Attendance/
    - FACT_Attendance_Daily: parsed directly from raw timesheets in data/Attendance/Timesheet/
    Completely eliminates any dependency on static HR_Fact_Attendance.xlsx!
    """
    print("\n[1] Loading Master Entities from entities/...")
    emp_df = load_master_employees_from_entities()

    print("\n[2] Loading Live Requests from data/Attendance/...")
    raw = {
        "DIM_Employee": emp_df,
        "Req_Leave": pd.DataFrame(),
        "Req_OT": pd.DataFrame(),
        "Req_WFh": pd.DataFrame(),
        "Req_LCin&ECout": pd.DataFrame(),
        "Req_BusinessTrip": pd.DataFrame(),
        "Req_ShiftChange": pd.DataFrame(),
    }
    load_live_requests(raw)

    print("\n[3] Parsing Raw Timesheets from data/Attendance/Timesheet/...")
    raw["FACT_Attendance_Daily"] = build_fact_attendance_daily()

    return raw

# ═══════════════════════════════════════════════════════════════
# PRE-PROCESSING
# ═══════════════════════════════════════════════════════════════

def preprocess(raw):
    """Cast types, filter excluded employees, add derived columns."""

    # -- FACT_Attendance_Daily ------------------------------
    att = raw["FACT_Attendance_Daily"].copy()
    att["Date_Text"]  = pd.to_datetime(att["Date_Text"],  errors="coerce")
    att["Is_Weekend"] = pd.to_numeric(att["Is_Weekend"],  errors="coerce").fillna(1).astype(int)
    att["Employee_ID"] = att["Employee_ID"].astype(str).str.strip()
    att = att[~att["Employee_ID"].isin(EXCLUDED_ALL_ATTENDANCE)]

    num_cols_att = [
        "Late_CI (mins)", "Early_CI (mins)", "Early_CO (mins)", "Late_CO (mins)",
        "Số giờ làm việc thực tế", "Số giờ làm việc tiêu chuẩn",
        "Delta (Số giờ làm việc thực tế - Số giờ làm việc tiêu chuẩn)",
    ]
    for c in num_cols_att:
        if c in att.columns:
            att[c] = pd.to_numeric(att[c], errors="coerce").fillna(0.0)

    # normalise Week Period (strip stray spaces)
    att["Week Period"] = att["Week Period"].astype(str).str.strip()

    # -- DIM_Employee ---------------------------------------
    emp = raw["DIM_Employee"].copy()
    emp["EmployeeID"] = emp["EmployeeID"].astype(str).str.strip()
    emp["IsActive"]   = pd.to_numeric(emp["IsActive"], errors="coerce").fillna(0).astype(int)
    emp = emp[~emp["EmployeeID"].isin(EXCLUDED_TOTAL)]

    # -- Req_Leave -----------------------------------------
    leave = raw["Req_Leave"].copy()
    if not leave.empty:
        leave["Employee_ID"]  = leave["Employee_ID"].astype(str).str.strip()
        leave["Leave_From"]   = pd.to_datetime(leave["Leave_From"],  errors="coerce")
        leave["Leave_To"]     = pd.to_datetime(leave["Leave_To"],    errors="coerce")
        leave["Submit_Date"]  = pd.to_datetime(leave["Submit_Date"], errors="coerce")
        leave["Leave_Days"]   = pd.to_numeric(leave["Leave_Days"],   errors="coerce").fillna(0)
        leave["Status_Flag"]  = pd.to_numeric(leave["Status_Flag"],  errors="coerce").fillna(0).astype(int)
        leave["Has_MonFri"]   = leave["Has_MonFri"].map(
            lambda x: str(x).strip().lower() in ["true", "1", "yes"]
        )
        # notice_category from (Leave_From - Submit_Date) in days
        leave["_notice_days"] = (leave["Leave_From"] - leave["Submit_Date"]).dt.days.fillna(999)
        leave["notice_category"] = leave["_notice_days"].apply(notice_category)

        # Leave Days in Week — use this for accurate partial-week leave
        if "Leave Days in Week" not in leave.columns:
            leave["Leave Days in Week"] = leave["Leave_Days"]
        leave["Leave Days in Week"] = pd.to_numeric(
            leave["Leave Days in Week"], errors="coerce"
        ).fillna(leave["Leave_Days"])

        if "Week Period" in leave.columns:
            leave["Week Period"] = leave["Week Period"].astype(str).str.strip()
        else:
            def to_wp(dt):
                if pd.isna(dt): return ""
                m = dt - pd.Timedelta(days=dt.weekday())
                f = m + pd.Timedelta(days=4)
                return f"{m.strftime('%d/%m/%Y')} - {f.strftime('%d/%m/%Y')}"
            leave["Week Period"] = leave["Leave_From"].apply(to_wp)
        leave = leave[~leave["Employee_ID"].isin(EXCLUDED_TOTAL)]

    # -- Req_OT --------------------------------------------
    ot = raw["Req_OT"].copy()
    if not ot.empty:
        ot["Employee_ID"] = ot["Employee_ID"].astype(str).str.strip()
        ot["OT_Hours"]    = pd.to_numeric(ot["OT_Hours"],   errors="coerce").fillna(0.0)
        ot["Status_Flag"] = pd.to_numeric(ot["Status_Flag"], errors="coerce").fillna(0).astype(int)
        ot["OT_From"]     = pd.to_datetime(ot["OT_From"],   errors="coerce")
        if "OT Date" in ot.columns:
            ot["OT Date"] = pd.to_datetime(ot["OT Date"],   errors="coerce")
            ot["_ot_date"] = ot["OT Date"]
        else:
            ot["_ot_date"] = ot["OT_From"].dt.normalize()
        ot["Week Period"] = ot["Week Period"].astype(str).str.strip() \
            if "Week Period" in ot.columns else ""
        ot = ot[~ot["Employee_ID"].isin(EXCLUDED_TOTAL)]

    # -- Req_WFh -------------------------------------------
    wfh = raw["Req_WFh"].copy()
    if not wfh.empty:
        wfh["Employee_ID"] = wfh["Employee_ID"].astype(str).str.strip()
        wfh["WFH_From"]    = pd.to_datetime(wfh["WFH_From"], errors="coerce")
        wfh["WFH_To"]      = pd.to_datetime(wfh["WFH_To"],   errors="coerce")
        wfh["Status_Flag"] = pd.to_numeric(wfh["Status_Flag"], errors="coerce").fillna(0).astype(int)
        wfh["Week Period"] = wfh["Week Period"].astype(str).str.strip() \
            if "Week Period" in wfh.columns else ""
        wfh = wfh[~wfh["Employee_ID"].isin(EXCLUDED_TOTAL)]

    # -- Req_LCin&ECout ------------------------------------
    lcec = raw["Req_LCin&ECout"].copy()
    if not lcec.empty:
        lcec["Employee_ID"] = lcec["Employee_ID"].astype(str).str.strip()
        lcec["Apply_From"]  = pd.to_datetime(lcec["Apply_From"], errors="coerce")
        lcec["Status_Flag"] = pd.to_numeric(lcec["Status_Flag"], errors="coerce").fillna(0).astype(int)
        lcec["Week Period"] = lcec["Week Period"].astype(str).str.strip() \
            if "Week Period" in lcec.columns else ""
        if "Minutes" in lcec.columns:
            lcec["Minutes"] = pd.to_numeric(lcec["Minutes"], errors="coerce").fillna(0)
        else:
            m1 = pd.to_numeric(lcec["Đi muộn đầu ca (Mins)"], errors="coerce").fillna(0) if "Đi muộn đầu ca (Mins)" in lcec.columns else 0
            m2 = pd.to_numeric(lcec["Đi muộn giữa ca (mins)"], errors="coerce").fillna(0) if "Đi muộn giữa ca (mins)" in lcec.columns else 0
            m3 = pd.to_numeric(lcec["Về sớm giữa ca"], errors="coerce").fillna(0) if "Về sớm giữa ca" in lcec.columns else 0
            m4 = pd.to_numeric(lcec["Về sớm cuối ca"], errors="coerce").fillna(0) if "Về sớm cuối ca" in lcec.columns else 0
            lcec["Minutes"] = m1 + m2 + m3 + m4
        lcec = lcec[~lcec["Employee_ID"].isin(EXCLUDED_TOTAL)]

    # -- Req_BusinessTrip ----------------------------------
    trip = raw["Req_BusinessTrip"].copy()
    if not trip.empty:
        trip["Employee_ID"] = trip["Employee_ID"].astype(str).str.strip()
        trip["Trip_From"]   = pd.to_datetime(trip["Trip_From"], errors="coerce")
        trip["Trip_To"]     = pd.to_datetime(trip["Trip_To"],   errors="coerce")
        trip["Trip_Days"]   = pd.to_numeric(trip["Trip_Days"],  errors="coerce").fillna(0)
        trip["Status_Flag"] = pd.to_numeric(trip["Status_Flag"], errors="coerce").fillna(0).astype(int)
        trip = trip[~trip["Employee_ID"].isin(EXCLUDED_TOTAL)]

    # -- Req_ShiftChange -----------------------------------
    sc = raw["Req_ShiftChange"].copy()
    if not sc.empty:
        sc["Employee_ID"] = sc["Employee_ID"].astype(str).str.strip()
        sc["Work_Date"]   = pd.to_datetime(sc["Work_Date"], errors="coerce")
        sc["Status_Flag"] = pd.to_numeric(sc["Status_Flag"], errors="coerce").fillna(0).astype(int)
        sc["Week Period"] = sc["Week Period"].astype(str).str.strip() \
            if "Week Period" in sc.columns else ""
        sc = sc[~sc["Employee_ID"].isin(EXCLUDED_TOTAL)]

    # -- Propagate Master Data across all tables --
    master_map = emp.set_index("EmployeeID").to_dict("index")
    for df in [att, leave, ot, wfh, lcec, trip, sc]:
        if df is None or df.empty or "Employee_ID" not in df.columns:
            continue
        df["DIM_Employee.FullNameEN"] = df["Employee_ID"].apply(lambda x: master_map.get(str(x), {}).get("FullNameEN", ""))
        df["DIM_Employee.Team"] = df["Employee_ID"].apply(lambda x: master_map.get(str(x), {}).get("Team", ""))
        df["DIM_Employee.Client"] = df["Employee_ID"].apply(lambda x: master_map.get(str(x), {}).get("Client", ""))
        df["DIM_Employee.IsActive"] = df["Employee_ID"].apply(lambda x: master_map.get(str(x), {}).get("IsActive", 1))
        df["DIM_Employee.Location"] = df["Employee_ID"].apply(lambda x: master_map.get(str(x), {}).get("Location", "Da Nang"))
        df["Department"] = df["DIM_Employee.Team"]

    return dict(att=att, emp=emp, leave=leave, ot=ot, wfh=wfh, lcec=lcec, trip=trip, sc=sc)

# ═══════════════════════════════════════════════════════════════
# WEEK & SCOPE DETECTION
# ═══════════════════════════════════════════════════════════════

def detect_weeks(att, n=TRAILING_WEEKS):
    """Return n most-recent Mon–Fri week period strings, ending with the current system week."""
    import pandas as pd
    now = pd.Timestamp.now()
    curr_mon = now - pd.Timedelta(days=now.weekday())
    curr_mon = curr_mon.normalize()
    start_mon = curr_mon - pd.Timedelta(days=7*(n-1))
    
    ranges = []
    for i in range(n):
        m = start_mon + pd.Timedelta(days=7*i)
        f = m + pd.Timedelta(days=4)
        s = f"{m.strftime('%d/%m/%Y')} - {f.strftime('%d/%m/%Y')}"
        ranges.append(s)
    return ranges

def build_scopes(emp):
    """Derive teams, clients, and their mappings from active DIM_Employee rows."""
    active = emp[emp["IsActive"] == 1]
    t2c: dict[str, set] = defaultdict(set)
    c2t: dict[str, set] = defaultdict(set)
    for _, row in active.iterrows():
        team   = str(row.get("Team",   "")).strip()
        client = str(row.get("Client", "")).strip()
        if team and client and team != "-" and client != "-":
            t2c[team].add(client)
            c2t[client].add(team)

    teams   = ["All Teams"]   + sorted(t2c)
    clients = ["All Clients"] + sorted(c2t)
    t2c_s   = {t: sorted(cs) for t, cs in t2c.items()}
    c2t_s   = {c: sorted(ts) for c, ts in c2t.items()}

    # All scope keys
    scope_keys = ["All Teams||All Clients"]
    for t in teams[1:]:
        scope_keys.append(f"{t}||All Clients")
    for c in clients[1:]:
        scope_keys.append(f"All Teams||{c}")
    for t, cs in t2c_s.items():
        for c in cs:
            scope_keys.append(f"{t}||{c}")

    return dict(teams=teams, clients=clients, t2c=t2c_s, c2t=c2t_s, scope_keys=scope_keys)

def emp_ids_for_scope(emp, team_scope, client_scope):
    """Set of active Employee_IDs matching the given team×client scope."""
    df = emp[emp["IsActive"] == 1]
    if team_scope != "All Teams":
        df = df[df["Team"] == team_scope]
    if client_scope != "All Clients":
        df = df[df["Client"] == client_scope]
    return set(df["EmployeeID"].astype(str).str.strip())

def build_emp_lookup(emp, scopes):
    """Pre-build {(team, client): set(emp_ids)} for all scope combos."""
    lookup = {}
    for sk in scopes["scope_keys"]:
        t, c = sk.split("||")
        lookup[(t, c)] = emp_ids_for_scope(emp, t, c)
    return lookup

# ═══════════════════════════════════════════════════════════════
# FILTER HELPERS
# ═══════════════════════════════════════════════════════════════

def att_scope(att_week, emp_ids, team_scope, client_scope):
    df = att_week
    if emp_ids:
        df = df[df["Employee_ID"].isin(emp_ids)]
    if team_scope != "All Teams":
        df = df[df["DIM_Employee.Team"] == team_scope]
    if client_scope != "All Clients":
        df = df[df["DIM_Employee.Client"] == client_scope]
    return df

def req_scope(df, emp_ids, team_scope, dept_col="Department"):
    """Filter a request DF to the scope via employee-ID set + optional team column."""
    out = df[df["Employee_ID"].isin(emp_ids)] if emp_ids else df
    if team_scope != "All Teams":
        if "DIM_Employee.Team" in out.columns:
            out = out[out["DIM_Employee.Team"] == team_scope]
        elif dept_col in out.columns:
            out = out[out[dept_col] == team_scope]
    return out

def req_week(df, week_period_str, date_col):
    """Slice a request DF to a specific week period string."""
    if "Week Period" in df.columns:
        r = df[df["Week Period"] == week_period_str]
        if not r.empty:
            return r
    # Fallback: filter by date range
    try:
        mon, fri = parse_week_mon_fri(week_period_str)
        return df[(df[date_col] >= mon) & (df[date_col] <= fri)]
    except Exception:
        return df.iloc[0:0]   # empty

def name_col(df, fallback_col="Employee_Name"):
    """Choose best available name column."""
    if "DIM_Employee.FullNameEN" in df.columns:
        return "DIM_Employee.FullNameEN"
    return fallback_col

# ═══════════════════════════════════════════════════════════════
# ATTENDANCE METRICS
# ═══════════════════════════════════════════════════════════════

def att_metrics(df_scope):
    """Compute all attendance-based KPIs from a filtered week+scope slice."""
    wkd = df_scope[df_scope["Is_Weekend"] == 0]
    tot = len(wkd)
    if tot == 0:
        return None

    work_mask = wkd["Type of Date"].isin(["FullWorkDay", "HalfWorkDay"])
    work = wkd[work_mask]

    actual_col = "Số giờ làm việc thực tế"
    std_col    = "Số giờ làm việc tiêu chuẩn"
    delta_col  = "Delta (Số giờ làm việc thực tế - Số giờ làm việc tiêu chuẩn)"

    # Attendance quality
    full_wd = (wkd["Type of Date"] == "FullWorkDay").sum()
    att_quality = pct(full_wd, tot)

    # Avg working hours (work records only, exclude zeros)
    if actual_col in work.columns and len(work) > 0:
        hrs_series = work[actual_col].replace(0, float("nan")).dropna()
        avg_hrs = round(float(hrs_series.mean()), 2) if len(hrs_series) > 0 else 0.0
    else:
        avg_hrs = 0.0

    # Adherence rate
    if actual_col in work.columns and std_col in work.columns and len(work) > 0:
        sum_std = work[std_col].sum()
        adh = pct(work[actual_col].sum(), sum_std) if sum_std > 0 else 0
    else:
        adh = 0.0

    # Late CI profile
    lci = wkd["Late_CI (mins)"] if "Late_CI (mins)" in wkd.columns else pd.Series(dtype=float)
    lci_le5   = int(((lci > 0) & (lci <= 5)).sum())
    lci_5to10 = int(((lci > 5) & (lci <= 10)).sum())
    lci_gt10  = int((lci > 10).sum())
    lci_total = lci_5to10 + lci_gt10
    late_ci_rate = pct(lci_total, tot)

    # Late CO profile (staying late)
    lco = wkd["Late_CO (mins)"] if "Late_CO (mins)" in wkd.columns else pd.Series(dtype=float)
    lco_lt30   = int(((lco > 0) & (lco < 30)).sum())
    lco_30to90 = int(((lco >= 30) & (lco <= LATE_CO_FLAG_MINS)).sum())
    lco_gt90   = int((lco > LATE_CO_FLAG_MINS).sum())
    lco_total  = int((lco > 0).sum())

    # Early CI signal
    eci = wkd["Early_CI (mins)"] if "Early_CI (mins)" in wkd.columns else pd.Series(dtype=float)
    eci_signal = int((eci > LATE_CI_THRESHOLD_MINS).sum())

    # Early CO
    eco = wkd["Early_CO (mins)"] if "Early_CO (mins)" in wkd.columns else pd.Series(dtype=float)
    eco_cnt = int((eco > LATE_CI_THRESHOLD_MINS).sum())
    eco_rate = pct(eco_cnt, tot)
    eco_avg  = round(float(eco[eco > LATE_CI_THRESHOLD_MINS].mean()), 1) if eco_cnt > 0 else 0.0

    # Extensive late working (actual hours > threshold)
    ext_late = int((wkd[actual_col] > EXTENSIVE_HOURS_THRESH).sum()) \
        if actual_col in wkd.columns else 0

    return dict(
        total_weekday_records=tot,
        total_weekday_employees=wkd["Employee_ID"].nunique(),
        work_record_count=len(work),
        work_employee_count=work["Employee_ID"].nunique(),
        attendance_quality=att_quality,
        avg_working_hours=avg_hrs,
        adherence_rate=adh,
        late_checkin_rate=late_ci_rate,
        late_ci_le5=lci_le5, late_ci_5to10=lci_5to10, late_ci_gt10=lci_gt10,
        late_ci_total=lci_total,
        late_ci_unexplained=0,           # filled in caller
        early_checkout_rate=eco_rate,
        avg_early_co=eco_avg,
        early_ci_signal_count=eci_signal,
        early_ci_explained_count=0,      # filled in caller
        early_ci_unexplained_count=eci_signal,
        late_co_lt30=lco_lt30, late_co_30to90=lco_30to90, late_co_gt90=lco_gt90,
        late_co_total=lco_total,
        extensive_late_count=ext_late,
    )

def fill_explained(am, wkd, lcec_lookup):
    """Populate explained/unexplained CI counts using the approved LCEC lookup."""
    lci_col = "Late_CI (mins)"
    if lci_col in wkd.columns:
        late_mask  = wkd[lci_col] > LATE_CI_THRESHOLD_MINS
        late_recs  = wkd[late_mask]
        explained  = sum(
            1 for _, r in late_recs.iterrows()
            if pd.notna(r["Date_Text"])
            and lcec_lookup.get((r["Employee_ID"], r["Date_Text"].date()), {}).get("ci_late", False)
        )
        am["late_ci_unexplained"] = am["late_ci_total"] - explained

    eci_col = "Early_CI (mins)"
    if eci_col in wkd.columns:
        early_mask = wkd[eci_col] > LATE_CI_THRESHOLD_MINS
        early_recs = wkd[early_mask]
        e_explained = sum(
            1 for _, r in early_recs.iterrows()
            if pd.notna(r["Date_Text"])
            and (r["Employee_ID"], r["Date_Text"].date()) in lcec_lookup
        )
        am["early_ci_explained_count"]   = e_explained
        am["early_ci_unexplained_count"] = am["early_ci_signal_count"] - e_explained

    return am

# ═══════════════════════════════════════════════════════════════
# MAIN DATA BUILD
# ═══════════════════════════════════════════════════════════════

def build_dash_data(proc, scopes, week_ranges):
    att   = proc["att"]
    emp   = proc["emp"]
    leave = proc["leave"]
    ot    = proc["ot"]
    wfh   = proc["wfh"]
    lcec  = proc["lcec"]
    trip  = proc["trip"]
    sc    = proc["sc"]

    teams      = scopes["teams"]
    clients    = scopes["clients"]
    t2c        = scopes["t2c"]
    c2t        = scopes["c2t"]
    scope_keys = scopes["scope_keys"]
    cur_week   = week_ranges[-1]

    print("\n[1] Building employee scope lookup…")
    emp_lkp = build_emp_lookup(emp, scopes)

    print("[2] Splitting attendance by week…")
    att_by_wk = {wp: att[att["Week Period"] == wp] for wp in week_ranges}

    # Approved OT set: (employee_id, date) for overwork detection
    if not ot.empty:
        ot_appr = ot[ot["Status_Flag"] == 1]
        ot_appr_set = {
            (str(r["Employee_ID"]), r["_ot_date"].date())
            for _, r in ot_appr.iterrows()
            if pd.notna(r["_ot_date"])
        }
    else:
        ot_appr_set = set()

    # Approved LCEC lookup: (employee_id, date) → {ci_late: bool}
    if not lcec.empty:
        lcec_appr = lcec[lcec["Status_Flag"] == 1]
        lcec_lkp  = {}
        for _, r in lcec_appr.iterrows():
            if pd.notna(r["Apply_From"]):
                key  = (r["Employee_ID"], r["Apply_From"].date())
                ci   = str(r.get("CI_Category_Mapped", "On Time")).strip()
                lcec_lkp[key] = {"ci_late": ci != "On Time"}
    else:
        lcec_lkp = {}

    # --- Per-scope × per-week KPI series -------------------
    print("[3] Computing KPI series for all scopes × weeks…")
    ZERO_SERIES = lambda: {k: [] for k in [
        "attendance_quality","overwork_employees","active_flags",
        "late_checkin_rate","avg_working_hours","extensive_late_count","adherence_rate",
        "leave_days","leave_approval","leave_headcount",
        "leave_planned_days","leave_unplanned_days","leave_urgent_days",
        "leave_planned_count","leave_unplanned_count","leave_urgent_count",
        "ot_hours","ot_weekend_hours","ot_weekday_hours","ot_employees","wfh_days",
        "lc_ec_events","trip_count","sc_count",
        "early_checkout_rate","avg_early_co",
        "early_ci_signal_count","early_ci_explained_count","early_ci_unexplained_count",
        "late_ci_le5","late_ci_5to10","late_ci_gt10","late_ci_total","late_ci_unexplained",
        "late_co_lt30","late_co_30to90","late_co_gt90","late_co_total",
        "total_weekday_records","total_weekday_employees",
        "work_record_count","work_employee_count",
    ]}

    data_out = {}
    for sk in scope_keys:
        t, c      = sk.split("||")
        emp_ids   = emp_lkp[(t, c)]
        series    = ZERO_SERIES()

        for wp in week_ranges:
            mon, fri   = parse_week_mon_fri(wp)
            week_dates = {(mon + timedelta(days=i)).date() for i in range(5)}

            # -- Attendance ------------------------------
            a_df  = att_scope(att_by_wk[wp], emp_ids, t, c)
            wkd   = a_df[a_df["Is_Weekend"] == 0]
            am    = att_metrics(a_df)

            if am is None:
                for k in series:
                    if k not in ("overwork_employees","active_flags",
                                 "leave_days","leave_approval","leave_headcount",
                                 "leave_planned_days","leave_unplanned_days","leave_urgent_days",
                                 "leave_planned_count","leave_unplanned_count","leave_urgent_count",
                                 "ot_hours","ot_weekend_hours","ot_weekday_hours","ot_employees","wfh_days",
                                 "lc_ec_events","trip_count","sc_count"):
                        series[k].append(0)
                    else:
                        series[k].append(0)
                continue

            am = fill_explained(am, wkd, lcec_lkp)
            for k, v in am.items():
                series[k].append(v)

            # -- Overwork (avg Delta > threshold, no OT filed) --
            delta_col = "Delta (Số giờ làm việc thực tế - Số giờ làm việc tiêu chuẩn)"
            if delta_col in wkd.columns:
                avg_delta    = wkd.groupby("Employee_ID")[delta_col].mean()
                ow_ids       = set(avg_delta[avg_delta > OVERWORK_THRESHOLD_HRS].index)
                ot_filed_wk  = {eid for (eid, d) in ot_appr_set if d in week_dates}
                series["overwork_employees"].append(len(ow_ids - ot_filed_wk))
            else:
                series["overwork_employees"].append(0)

            # -- Leave ------------------------------------
            if not leave.empty:
                l_wk = leave[(leave["Leave_From"] <= fri) & (leave["Leave_To"] >= mon)]
                l_sc = req_scope(l_wk, emp_ids, t)
                if not l_sc.empty:
                    l_sc = l_sc.copy()
                    l_sc["_days_in_wk"] = l_sc.apply(
                        lambda r: get_leave_days_in_week(r["Leave_From"], r["Leave_To"], r["Leave_Days"], mon, fri),
                        axis=1
                    )
                    l_sc = l_sc[l_sc["_days_in_wk"] > 0]
                    l_app = l_sc[l_sc["Status_Flag"] == 1]
                    days_col = "_days_in_wk"
                else:
                    l_app = l_sc
                    days_col = "Leave Days in Week"

                planned_m   = l_sc["notice_category"] == "Planned"
                unplanned_m = l_sc["notice_category"] == "Unplanned"
                urgent_m    = l_sc["notice_category"] == "Urgent"

                series["leave_days"].append(round(float(l_sc[days_col].sum()), 1) if not l_sc.empty else 0.0)
                series["leave_approval"].append(pct(len(l_app), len(l_sc)) if len(l_sc) > 0 else 0)
                series["leave_headcount"].append(l_sc["Employee_ID"].nunique() if not l_sc.empty else 0)
                series["leave_planned_days"].append(round(float(l_sc.loc[planned_m,   days_col].sum()), 1) if not l_sc.empty else 0.0)
                series["leave_unplanned_days"].append(round(float(l_sc.loc[unplanned_m, days_col].sum()), 1) if not l_sc.empty else 0.0)
                series["leave_urgent_days"].append(round(float(l_sc.loc[urgent_m,   days_col].sum()), 1) if not l_sc.empty else 0.0)
                series["leave_planned_count"].append(int(planned_m.sum()) if not l_sc.empty else 0)
                series["leave_unplanned_count"].append(int(unplanned_m.sum()) if not l_sc.empty else 0)
                series["leave_urgent_count"].append(int(urgent_m.sum()) if not l_sc.empty else 0)
            else:
                for k in ["leave_days","leave_approval","leave_headcount",
                          "leave_planned_days","leave_unplanned_days","leave_urgent_days",
                          "leave_planned_count","leave_unplanned_count","leave_urgent_count"]:
                    series[k].append(0)

            # -- OT ---------------------------------------
            if not ot.empty:
                o_wk  = req_week(ot, wp, "_ot_date")
                o_sc  = req_scope(o_wk, emp_ids, t)
                o_app = o_sc[o_sc["Status_Flag"] == 1]
                wknd_hrs  = round(float(o_app.loc[o_app["OT_Timing"] == "Ngày nghỉ", "OT_Hours"].sum()), 1)
                total_hrs = round(float(o_app["OT_Hours"].sum()), 1)
                emp_count = o_app["Employee_ID"].nunique() if "Employee_ID" in o_app.columns else 0
                series["ot_hours"].append(total_hrs)
                series["ot_weekend_hours"].append(wknd_hrs)
                series["ot_weekday_hours"].append(round(total_hrs - wknd_hrs, 1))
                series["ot_employees"].append(emp_count)
            else:
                series["ot_hours"].append(0); series["ot_weekend_hours"].append(0); series["ot_weekday_hours"].append(0); series["ot_employees"].append(0)

            # -- WFH --------------------------------------
            if not wfh.empty:
                w_wk = req_week(wfh, wp, "WFH_From")
                w_sc = req_scope(w_wk, emp_ids, t)
                series["wfh_days"].append(int((w_sc["Status_Flag"] == 1).sum()))
            else:
                series["wfh_days"].append(0)

            # -- LCEC events -------------------------------
            if not lcec.empty:
                lc_wk = req_week(lcec, wp, "Apply_From")
                lc_sc = req_scope(lc_wk, emp_ids, t)
                series["lc_ec_events"].append(int((lc_sc["Status_Flag"] == 1).sum()))
            else:
                series["lc_ec_events"].append(0)

            # -- Business trips ----------------------------
            if not trip.empty:
                tr_wk = req_week(trip, wp, "Trip_From")
                tr_sc = req_scope(tr_wk, emp_ids, t)
                series["trip_count"].append(int((tr_sc["Status_Flag"] == 1).sum()))
            else:
                series["trip_count"].append(0)

            # -- Shift changes -----------------------------
            if not sc.empty:
                sc_wk = req_week(sc, wp, "Work_Date")
                sc_sc = req_scope(sc_wk, emp_ids, t)
                series["sc_count"].append(int((sc_sc["Status_Flag"] == 1).sum()))
            else:
                series["sc_count"].append(0)

            # active_flags placeholder — filled after flag computation
            series["active_flags"].append(0)

        data_out[sk] = series

    # --- Weekly Data Evaluation -----------------------------
    print("[4-9] Computing flags, incidents, heatmaps, and details for ALL weeks...")

    flags_out = {w: {} for w in week_ranges}
    incidents_out = {w: {} for w in week_ranges}
    top_hours_out = {w: {} for w in week_ranges}
    team_ranking_out = {w: {} for w in week_ranges}
    leave_calendar = {w: {} for w in week_ranges}
    leave_calendar_next = {w: {} for w in week_ranges}
    leave_details_current = {w: {} for w in week_ranges}
    leave_details_next = {w: {} for w in week_ranges}
    ot_details = {w: {} for w in week_ranges}
    wfh_details = {w: {} for w in week_ranges}
    request_details = {w: {} for w in week_ranges}

    for eval_week in week_ranges:
        eval_mon, eval_fri = parse_week_mon_fri(eval_week)
        eval_dates = {(eval_mon + timedelta(days=i)).date() for i in range(5)}
        
        next_mon = eval_mon + timedelta(days=7)
        next_fri = eval_fri + timedelta(days=7)
        next_wp  = f"{next_mon.strftime('%d/%m/%Y')} - {next_fri.strftime('%d/%m/%Y')}"

        # 1. Flags
        for sk in scope_keys:
            t, c    = sk.split("||")
            emp_ids = emp_lkp[(t, c)]
            flags   = []

            # Overwork, no OT filed
            a_df = att_scope(att_by_wk[eval_week], emp_ids, t, c)
            wkd  = a_df[a_df["Is_Weekend"] == 0]
            delta_col = "Delta (Số giờ làm việc thực tế - Số giờ làm việc tiêu chuẩn)"
            if delta_col in wkd.columns and len(wkd) > 0:
                avg_delta = wkd.groupby("Employee_ID")[delta_col].mean()
                ot_filed  = {eid for (eid, d) in ot_appr_set if d in eval_dates}
                for eid, delta in avg_delta.items():
                    if delta > OVERWORK_THRESHOLD_HRS and eid not in ot_filed:
                        name_series = wkd.loc[wkd["Employee_ID"] == eid, "DIM_Employee.FullNameEN"]
                        name = name_series.iloc[0] if len(name_series) > 0 else eid
                        flags.append({"type": "Overwork, no OT filed", "severity": "danger",
                                      "text": f"{name}: avg +{delta:.1f}h/day this week, no OT registered"})

            # Weekend-bridge leave & Short-notice leave
            if not leave.empty:
                l_wk = leave[(leave["Leave_From"] <= eval_fri) & (leave["Leave_To"] >= eval_mon)]
                l_sc = req_scope(l_wk, emp_ids, t)
                for _, r in l_sc.iterrows():
                    nm = r.get("DIM_Employee.FullNameEN") or r.get("Employee_Name") or r["Employee_ID"]
                    date_str = fmt_date(r.get("Leave_From"))[:5]
                    if r.get("Has_MonFri"):
                        flags.append({"type": "Weekend-bridge leave", "severity": "warning",
                                      "text": f"{nm}: leave adjoins weekend ({date_str})"})
                    if r.get("_notice_days", 999) < 1:
                        flags.append({"type": "Short-notice leave", "severity": "warning",
                                      "text": f"{nm}: leave requested with <1 day notice"})

            # Friday shift-change cluster
            if not sc.empty:
                sc_wk = req_week(sc, eval_week, "Work_Date")
                sc_sc = req_scope(sc_wk, emp_ids, t)
                fri_sc = sc_sc[sc_sc["Work_Date"].dt.date == eval_fri.date()]
                if len(fri_sc) >= 3:
                    flags.append({"type": "Friday shift-change cluster", "severity": "warning",
                                  "text": f"{len(fri_sc)} shift changes on Friday — may thin end-of-week capacity"})

            # SC & LCEC FLAGS
            if not sc.empty:
                idx = week_ranges.index(eval_week)
                start_idx = max(0, idx - 1)
                s_2wk = sc[sc["Week Period"].isin(week_ranges[start_idx:idx+1])]
                s_sc = req_scope(s_2wk, emp_ids, t)
                sc_counts = s_sc.groupby("Employee_ID").size()
                for eid, cnt in sc_counts.items():
                    if cnt >= 2:
                        nm = emp[emp["EmployeeID"] == eid]["FullNameEN"].iloc[0] if eid in emp["EmployeeID"].values else eid
                        flags.append({"type": "Volatile Shift Switch", "severity": "warning", "text": f"{nm}: {cnt} shift changes in last 2 weeks"})
            
            if not lcec.empty:
                idx = week_ranges.index(eval_week)
                start_idx = max(0, idx - 1)
                lc_2wk = lcec[lcec["Week Period"].isin(week_ranges[start_idx:idx+1])]
                lc_sc = req_scope(lc_2wk, emp_ids, t)
                lc_counts = lc_sc.groupby("Employee_ID").size()
                for eid, cnt in lc_counts.items():
                    if cnt >= 2:
                        nm = emp[emp["EmployeeID"] == eid]["FullNameEN"].iloc[0] if eid in emp["EmployeeID"].values else eid
                        flags.append({"type": "Unstable Working Time", "severity": "warning", "text": f"{nm}: {cnt} late CI/early CO requests in last 2 weeks"})

            flags_out[eval_week][sk] = flags
            # active_flags update logic is already handled in KPI loop for cur_week, we won't rewrite that part.

        # 2. Extensive late incidents
        for sk in scope_keys:
            t, c    = sk.split("||")
            emp_ids = emp_lkp[(t, c)]
            rows    = []
            # we only do for eval_week, previously it was all 4 weeks aggregated, now it's per week.
            a_wk = att_scope(att_by_wk[eval_week], emp_ids, t, c)
            wkd  = a_wk[a_wk["Is_Weekend"] == 0]
            lco  = wkd["Late_CO (mins)"] if "Late_CO (mins)" in wkd.columns else pd.Series(dtype=float)
            bad  = wkd[lco > LATE_CO_FLAG_MINS]
            for _, r in bad.iterrows():
                rows.append({
                    "name":          str(r.get("DIM_Employee.FullNameEN") or r.get("Employee_Name") or ""),
                    "team":          str(r.get("DIM_Employee.Team", "")),
                    "date":          fmt_date(r.get("Date_Text")),
                    "checkout_time": fmt_time(r.get("CheckOut_Time")),
                    "working_hours": round(safe_float(r.get("Số giờ làm việc thực tế")), 2),
                    "late_co_mins":  safe_int(r.get("Late_CO (mins)")),
                })
            incidents_out[eval_week][sk] = rows

        # 3. Top avg working hours
        actual_col = "Số giờ làm việc thực tế"
        for sk in scope_keys:
            t, c    = sk.split("||")
            emp_ids = emp_lkp[(t, c)]
            a_wk    = att_scope(att_by_wk[eval_week], emp_ids, t, c)
            work    = a_wk[(a_wk["Is_Weekend"] == 0) & a_wk["Type of Date"].isin(["FullWorkDay", "HalfWorkDay"])]
            if work.empty or actual_col not in work.columns:
                top_hours_out[eval_week][sk] = []
                continue
            grp = work.groupby("Employee_ID").agg(
                avg_hours     =(actual_col,                "mean"),
                name          =("DIM_Employee.FullNameEN", "first"),
                team          =("DIM_Employee.Team",       "first"),
                client        =("DIM_Employee.Client",     "first"),
            ).sort_values("avg_hours", ascending=False).head(10)
            top_hours_out[eval_week][sk] = [
                {"name": str(r["name"]), "team": str(r["team"]),
                 "client": str(r["client"]), "value": round(float(r["avg_hours"]), 2)}
                for _, r in grp.iterrows()
            ]

        # 4. Team ranking
        for cs in clients:
            lci_rank, adh_rank = [], []
            for ts in teams[1:]:
                sk = f"{ts}||{cs}"
                if sk not in data_out:
                    continue
                # For eval_week, get the index
                wk_idx = week_ranges.index(eval_week)
                lci_v = data_out[sk]["late_checkin_rate"][wk_idx] if data_out[sk]["late_checkin_rate"] else 0
                adh_v = data_out[sk]["adherence_rate"][wk_idx]    if data_out[sk]["adherence_rate"]    else 0

                emp_ids   = emp_lkp[(ts, cs)]
                a_wk      = att_scope(att_by_wk[eval_week], emp_ids, ts, cs)
                wkd_r     = a_wk[a_wk["Is_Weekend"] == 0]
                work_r    = wkd_r[wkd_r["Type of Date"].isin(["FullWorkDay", "HalfWorkDay"])]

                top_lci_emps, top_adh_emps = [], []
                if "Late_CI (mins)" in wkd_r.columns:
                    late_r = wkd_r[wkd_r["Late_CI (mins)"] > LATE_CI_THRESHOLD_MINS]
                    if not late_r.empty:
                        g = late_r.groupby("Employee_ID").agg(
                            v=("Late_CI (mins)", "mean"), nm=("DIM_Employee.FullNameEN", "first")
                        ).nlargest(3, "v")
                        top_lci_emps = [{"name": str(r["nm"]), "value": round(float(r["v"]), 1)} for _, r in g.iterrows()]

                std_c = "Số giờ làm việc tiêu chuẩn"
                if actual_col in work_r.columns and std_c in work_r.columns and len(work_r) > 0:
                    def adh_emp(g):
                        s = g[std_c].sum()
                        return 100 * g[actual_col].sum() / s if s > 0 else 0
                    g = work_r.groupby("Employee_ID").apply(adh_emp).reset_index(name="adh")
                    g["nm"] = g["Employee_ID"].map(
                        work_r.groupby("Employee_ID")["DIM_Employee.FullNameEN"].first()
                    )
                    top_adh_emps = [
                        {"name": str(r["nm"]), "value": round(float(r["adh"]), 1)}
                        for _, r in g.nlargest(3, "adh").iterrows()
                    ]

                if lci_v > 0 or data_out.get(sk, {}).get("total_weekday_records", [0])[wk_idx] > 0:
                    lci_rank.append({"team": ts, "value": lci_v, "topEmployees": top_lci_emps})
                if adh_v > 0:
                    adh_rank.append({"team": ts, "value": adh_v, "topEmployees": top_adh_emps})

            lci_rank.sort(key=lambda x: -x["value"])
            adh_rank.sort(key=lambda x: x["value"])
            team_ranking_out[eval_week][cs] = {"late_checkin_rate": lci_rank, "adherence_rate": adh_rank}

        # 5. Leave calendar heatmaps
        def build_calendar(week_period_str, is_next=False):
            if leave.empty: return {}
            cal_mon, cal_fri = parse_week_mon_fri(week_period_str)
            weekdays = [(cal_mon + timedelta(days=i)) for i in range(5)]
            day_names = ["Mon", "Tue", "Wed", "Thu", "Fri"]
            l_wk = leave[(leave["Leave_From"] <= cal_fri) & (leave["Leave_To"] >= cal_mon)]
            l_appr = l_wk[l_wk["Status_Flag"] == 1]
            cal_keys = ([f"All Teams||{c}" for c in clients[1:]] + [f"All Teams||All Clients"] + [f"{t}||{c}" for t in teams[1:] for c in t2c.get(t, [])])
            cal = {}
            for sk in cal_keys:
                tt, cc = sk.split("||")
                emp_ids = emp_lkp[(tt, cc)]
                headcount = len(emp_ids)
                if headcount == 0: continue
                l_scope = l_appr[l_appr["Employee_ID"].isin(emp_ids)]
                days = []
                for day_dt, day_nm in zip(weekdays, day_names):
                    on_leave_names = []
                    for _, r in l_scope.iterrows():
                        lf = r.get("Leave_From"); lt = r.get("Leave_To")
                        if pd.notna(lf) and pd.notna(lt) and lf.date() <= day_dt.date() <= lt.date():
                            nm = r.get("DIM_Employee.FullNameEN") or r.get("Employee_Name") or ""
                            on_leave_names.append(str(nm))
                    cnt = len(on_leave_names)
                    days.append({
                        "weekday": day_nm, "date": day_dt.strftime("%d/%m"),
                        "on_leave_count": cnt, "total_headcount": headcount, "pct": pct(cnt, headcount),
                        "avg_prior_weeks": 0.0, "employees": on_leave_names, "elevated": False,
                    })
                cal[sk] = {"headcount": headcount, "days": days}
            return cal

        e_cal  = build_calendar(eval_week, is_next=False)
        n_cal  = build_calendar(next_wp, is_next=True)
        for sk, entry in e_cal.items():
            if sk not in n_cal:
                n_cal[sk] = {"headcount": entry["headcount"], "days": [{**d, "on_leave_count": 0, "pct": 0, "employees": [], "elevated": False} for d in entry["days"]]}
        
        # Compute avg_prior_weeks for e_cal
        if not leave.empty:
            wk_idx = week_ranges.index(eval_week)
            prior_weeks = week_ranges[:wk_idx] if wk_idx > 0 else []
            for sk, entry in e_cal.items():
                tt, cc = sk.split("||")
                emp_ids = emp_lkp[(tt, cc)]
                l_appr_all = leave[leave["Status_Flag"] == 1]
                l_scope_all = l_appr_all[l_appr_all["Employee_ID"].isin(emp_ids)]
                for i, day_info in enumerate(entry["days"]):
                    prior_cnts = []
                    for pw in prior_weeks:
                        pw_mon, _ = parse_week_mon_fri(pw)
                        pw_day    = pw_mon + timedelta(days=i)
                        cnt = sum(1 for _, r in l_scope_all.iterrows() if pd.notna(r.get("Leave_From")) and pd.notna(r.get("Leave_To")) and r["Leave_From"].date() <= pw_day.date() <= r["Leave_To"].date())
                        prior_cnts.append(cnt)
                    avg_p = round(sum(prior_cnts) / len(prior_cnts), 1) if prior_cnts else 0.0
                    cur_c = day_info["on_leave_count"]
                    entry["days"][i]["avg_prior_weeks"] = avg_p
                    entry["days"][i]["elevated"] = (cur_c > 0 and cur_c > avg_p * 1.5 and cur_c > avg_p + 0.5)

        leave_calendar[eval_week] = e_cal
        leave_calendar_next[eval_week] = n_cal

        # 6. Detail tables
        all_table_keys = (["All Teams||All Clients"] + [f"{t}||All Clients" for t in teams[1:]] + [f"All Teams||{c}" for c in clients[1:]] + [f"{t}||{c}" for t in teams[1:] for c in t2c.get(t, [])])
        
        # Leave tables
        l_cur_res, l_next_res = {}, {}
        for sk in all_table_keys:
            tt, cc = sk.split("||"); emp_ids = emp_lkp[(tt, cc)]
            # Current
            if leave.empty: l_cur_res[sk] = []
            else:
                l_wk = leave[(leave["Leave_From"] <= eval_fri) & (leave["Leave_To"] >= eval_mon)]
                l_sc = req_scope(l_wk, emp_ids, tt)
                l_cur_res[sk] = [{"name": str(r.get("DIM_Employee.FullNameEN") or r.get("Employee_Name") or ""), "from": fmt_date(r.get("Leave_From")), "to": fmt_date(r.get("Leave_To")), "days": get_leave_days_in_week(r.get("Leave_From"), r.get("Leave_To"), r.get("Leave_Days"), eval_mon, eval_fri), "type": str(r.get("Leave_Type_Mapped") or r.get("Leave_Type") or ""), "notice_category": str(r.get("notice_category", "Unknown")), "status": str(r.get("Status", ""))} for _, r in l_sc.iterrows() if get_leave_days_in_week(r.get("Leave_From"), r.get("Leave_To"), r.get("Leave_Days"), eval_mon, eval_fri) > 0]
            # Next
            if leave.empty: l_next_res[sk] = []
            else:
                l_wk = leave[(leave["Leave_From"] <= next_fri) & (leave["Leave_To"] >= next_mon)]
                l_sc = req_scope(l_wk, emp_ids, tt)
                l_next_res[sk] = [{"name": str(r.get("DIM_Employee.FullNameEN") or r.get("Employee_Name") or ""), "from": fmt_date(r.get("Leave_From")), "to": fmt_date(r.get("Leave_To")), "days": get_leave_days_in_week(r.get("Leave_From"), r.get("Leave_To"), r.get("Leave_Days"), next_mon, next_fri), "type": str(r.get("Leave_Type_Mapped") or r.get("Leave_Type") or ""), "notice_category": str(r.get("notice_category", "Unknown")), "status": str(r.get("Status", ""))} for _, r in l_sc.iterrows() if get_leave_days_in_week(r.get("Leave_From"), r.get("Leave_To"), r.get("Leave_Days"), next_mon, next_fri) > 0]
        leave_details_current[eval_week] = l_cur_res
        leave_details_next[eval_week] = l_next_res

        # OT & WFH
        ot_res, wfh_res = {}, {}
        for sk in all_table_keys:
            tt, cc = sk.split("||"); emp_ids = emp_lkp[(tt, cc)]
            if ot.empty: ot_res[sk] = []
            else:
                o_wk = req_week(ot, eval_week, "_ot_date")
                o_sc = req_scope(o_wk, emp_ids, tt); o_app = o_sc[o_sc["Status_Flag"] == 1]; nc = name_col(o_app)
                ot_res[sk] = [{"name": str(r.get(nc) or r.get("Employee_Name") or ""), "date": fmt_date(r.get("OT Date") or r.get("OT_From")), "hours": safe_float(r.get("OT_Hours")), "timing": str(r.get("OT_Timing", "")), "reason": str(r.get("Reason", "")), "status": str(r.get("Status", ""))} for _, r in o_app.iterrows()]
            
            if wfh.empty: wfh_res[sk] = []
            else:
                w_wk = req_week(wfh, eval_week, "WFH_From")
                w_sc = req_scope(w_wk, emp_ids, tt); w_app = w_sc[w_sc["Status_Flag"] == 1]; nc = name_col(w_app)
                wfh_res[sk] = [{"name": str(r.get(nc) or ""), "from": fmt_date(r.get("WFH_From")), "to": fmt_date(r.get("WFH_To")), "reason": str(r.get("Reason", "")), "status": str(r.get("Status", ""))} for _, r in w_app.iterrows()]
        ot_details[eval_week] = ot_res
        wfh_details[eval_week] = wfh_res

        # Request details (LCEC, Trip, SC)
        lcec_r, trip_r, sc_r = {}, {}, {}
        for sk in all_table_keys:
            tt, cc = sk.split("||"); emp_ids = emp_lkp[(tt, cc)]
            if not lcec.empty:
                lc_wk = req_week(lcec, eval_week, "Apply_From"); lc_sc = req_scope(lc_wk, emp_ids, tt); lc_app = lc_sc[lc_sc["Status_Flag"] == 1]; nc = name_col(lc_app)
                lcec_r[sk] = [{"name": str(r.get(nc) or ""), "date": fmt_date(r.get("Apply_From")), "ci_category": str(r.get("CI_Category_Mapped", "")), "co_category": str(r.get("CO_Category_Mapped", "")), "minutes": safe_float(r.get("Minutes")), "reason": str(r.get("Reason_Detail") or r.get("Reason_Group") or ""), "status": str(r.get("Status", ""))} for _, r in lc_app.iterrows()]
            else: lcec_r[sk] = []
            
            if not trip.empty:
                tr_wk = req_week(trip, eval_week, "Trip_From"); tr_sc = req_scope(tr_wk, emp_ids, tt); tr_app = tr_sc[tr_sc["Status_Flag"] == 1]
                trip_r[sk] = [{"name": str(r.get("Employee_Name") or ""), "from": fmt_date(r.get("Trip_From")), "to": fmt_date(r.get("Trip_To")), "days": safe_float(r.get("Trip_Days")), "destination": str(r.get("Destination", "")), "purpose": str(r.get("Purpose", "")), "status": str(r.get("Status", ""))} for _, r in tr_app.iterrows()]
            else: trip_r[sk] = []
            
            if not sc.empty:
                sc_wk = req_week(sc, eval_week, "Work_Date"); sc_sc = req_scope(sc_wk, emp_ids, tt); sc_app = sc_sc[sc_sc["Status_Flag"] == 1]; nc = name_col(sc_app)
                sc_r[sk] = [{"name": str(r.get(nc) or r.get("Employee_Name") or ""), "date": fmt_date(r.get("Work_Date")), "old_shift": str(r.get("Shift_Code_Old", "")), "new_shift": str(r.get("Shift_Code_New", "")), "reason": str(r.get("Reason", "")), "status": str(r.get("Status", ""))} for _, r in sc_app.iterrows()]
            else: sc_r[sk] = []
        request_details[eval_week] = {"lcec": lcec_r, "trip": trip_r, "shift_change": sc_r}

    def build_attendance_dashboard(att, emp, leave, week_ranges):
        import random
        import pandas as pd
        
        # 1. Absences
        absences = []
        lwk = leave[leave["Status_Flag"] == 1]
        for _, r in lwk.iterrows():
            try:
                lf = pd.to_datetime(r["Leave_From"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
                lt = pd.to_datetime(r["Leave_To"], format="%Y-%m-%d %H:%M:%S", errors="coerce")
                if pd.isnull(lf) or pd.isnull(lt): continue
                delta = (lt - lf).days
                for i in range(delta + 1):
                    d = lf + pd.Timedelta(days=i)
                    if d.weekday() < 5: # Mon-Fri
                        absences.append({
                            "date": d.strftime("%Y-%m-%d"),
                            "name": str(r.get("Employee_Name", "")),
                            "customer": str(r.get("DIM_Employee.Client", "")),
                            "team": str(r.get("DIM_Employee.Team", "")),
                            "reason": (str(r.get("Reason", "")).strip() if str(r.get("Reason", "")).strip() and str(r.get("Reason", "")).strip().lower() != "nan" else str(r.get("Leave_Type", "Leave"))),
                            "notice_category": str(r.get("notice_category", "Unknown"))
                        })
            except:
                pass
            
        # 2. Late Checkouts > 90 mins
        att["_L_CO"] = pd.to_numeric(att["Late_CO (mins)"], errors="coerce").fillna(0)
        late_co = att[att["_L_CO"] > 90]
        late_watch = []
        for _, r in late_co.iterrows():
            t = str(r.get("CheckOut_Time", ""))
            if len(t) >= 5: t = t[:5]
            late_watch.append({
                "week": week_label(str(r.get("Week Period", ""))),
                "name": str(r.get("Employee_Name", "")),
                "customer": str(r.get("DIM_Employee.Client", "")),
                "team": str(r.get("DIM_Employee.Team", "")),
                "mins": safe_float(r.get("_L_CO", 0)),
                "time": t
            })
            
        # 3. Average Working Hours per Customer per Week
        att["_Hrs"] = pd.to_numeric(att["Số giờ làm việc thực tế"], errors="coerce")
        work = att[(att["Is_Weekend"] == 0) & att["Type of Date"].isin(["FullWorkDay", "HalfWorkDay"])]
        avg_hrs = work.groupby(["Week Period", "DIM_Employee.Client"])["_Hrs"].mean().reset_index()
        
        avg_hrs_dict = {}
        for _, r in avg_hrs.iterrows():
            wk = week_label(str(r["Week Period"]))
            c = str(r["DIM_Employee.Client"])
            val = safe_float(r["_Hrs"])
            if wk not in avg_hrs_dict: avg_hrs_dict[wk] = {}
            avg_hrs_dict[wk][c] = val
            
        # 4. Generate continuous weeks & weekly leaves
        import pandas as pd
        min_date = pd.to_datetime(att["Date_Text"], format="%d/%m/%Y", errors="coerce").min()
        max_date = pd.to_datetime(leave["Leave_To"]).max()
        if pd.isnull(max_date): max_date = pd.to_datetime(att["Date_Text"], format="%d/%m/%Y", errors="coerce").max()
        if pd.isnull(min_date): min_date = pd.Timestamp.now() - pd.Timedelta(days=30)
        if pd.isnull(max_date): max_date = pd.Timestamp.now() + pd.Timedelta(days=30)
        
        start_mon = min_date - pd.Timedelta(days=min_date.weekday())
        end_mon = max_date - pd.Timedelta(days=max_date.weekday())
        
        continuous_weeks = []
        calendars = {}
        weekly_leaves = []
        curr = start_mon
        while curr <= end_mon:
            fri = curr + pd.Timedelta(days=4)
            wp_str = f"{curr.strftime('%d/%m/%Y')} - {fri.strftime('%d/%m/%Y')}"
            if curr.month == fri.month:
                cw = f"Week {curr.strftime('%b %d')} - {fri.strftime('%d')}"
            else:
                cw = f"Week {curr.strftime('%b %d')} - {fri.strftime('%b %d')}"
            continuous_weeks.append(cw)
            
            cal_data = build_calendar(wp_str, is_next=False)
            if cal_data:
                calendars[cw] = cal_data
                
            # Populate weekly leaves for this week cw
            l_wk_appr = lwk[(lwk["Leave_From"] <= fri) & (lwk["Leave_To"] >= curr)]
            for _, r in l_wk_appr.iterrows():
                days_in_wk = get_leave_days_in_week(
                    r.get("Leave_From"), r.get("Leave_To"), r.get("Leave_Days"), curr, fri
                )
                if days_in_wk > 0:
                    weekly_leaves.append({
                        "week": cw,
                        "name": str(r.get("Employee_Name", "")),
                        "customer": str(r.get("DIM_Employee.Client", "")),
                        "team": str(r.get("DIM_Employee.Team", "")),
                        "from": str(r.get("Leave_From", ""))[:10],
                        "to": str(r.get("Leave_To", ""))[:10],
                        "days": days_in_wk,
                        "notice_category": str(r.get("notice_category", "Unknown")),
                        "reason": (str(r.get("Reason", "")).strip()
                                   if str(r.get("Reason", "")).strip() and str(r.get("Reason", "")).strip().lower() != "nan"
                                   else str(r.get("Leave_Type_Mapped", "Leave")))
                    })
                
            curr += pd.Timedelta(days=7)
            
        return {
            "current_week": week_label(cur_week),
            "weeks": continuous_weeks,
            "calendars": calendars,
            "absences": absences,
            "weekly_leaves": weekly_leaves,
            "late_watch": late_watch,
            "avg_hrs": avg_hrs_dict
        }

    # Build employee mappings for client / scope tooltips
    active_emp = emp[emp["IsActive"] == 1]
    emp_dict = active_emp.set_index("EmployeeID").to_dict("index")

    scope_to_employees = {}
    for sk in scope_keys:
        t, c = sk.split("||")
        eids = emp_lkp.get((t, c), set())
        scope_to_employees[sk] = [
            {
                "id": eid,
                "name": str(emp_dict.get(eid, {}).get("FullNameEN", eid)),
                "name_vn": str(emp_dict.get(eid, {}).get("FullNameVN", "")),
                "team": str(emp_dict.get(eid, {}).get("Team", "")),
                "position": str(emp_dict.get(eid, {}).get("Position", "")),
                "client": str(emp_dict.get(eid, {}).get("Client", "")),
            }
            for eid in sorted(eids)
        ]

    client_to_employees = {}
    for c in clients:
        if c == "All Clients":
            ce_df = active_emp
        else:
            ce_df = active_emp[active_emp["Client"] == c]
        client_to_employees[c] = [
            {
                "id": str(r["EmployeeID"]),
                "name": str(r.get("FullNameEN", r["EmployeeID"])),
                "name_vn": str(r.get("FullNameVN", "")),
                "team": str(r.get("Team", "")),
                "position": str(r.get("Position", "")),
            }
            for _, r in ce_df.iterrows()
        ]

    # --- Assemble --------------------------------------------
    print("[10] Assembling DASH_DATA...")

    return {
        "weeks":                [week_label(w) for w in week_ranges],
        "week_ranges":          [week_range_label(w) for w in week_ranges],
        "week_keys":            list(week_ranges),
        "teams":                teams,
        "clients":              clients,
        "team_to_clients":      t2c,
        "client_to_teams":      c2t,
        "client_to_employees":  client_to_employees,
        "scope_to_employees":   scope_to_employees,
        "data":                 data_out,
        "flags":                flags_out,
        "extensive_late_incidents": incidents_out,
        "top_avg_working_hours":    top_hours_out,
        "team_ranking":             team_ranking_out,
        "leave_calendar":           leave_calendar,
        "leave_calendar_next":      leave_calendar_next,
        "leave_details_current":    leave_details_current,
        "leave_details_next":       leave_details_next,
        "ot_details":               ot_details,
        "wfh_details":              wfh_details,
        "request_details":          request_details,
        "attendance_dashboard":     build_attendance_dashboard(att, emp, leave, week_ranges),
        "next_week_range":          f"{next_mon.strftime('%b %d').replace(' 0',' ')}, "
                                    f"{next_fri.strftime('%b %d').replace(' 0',' ')}, {next_mon.strftime('%y')}",
        "late_ci_threshold":        LATE_CI_THRESHOLD_MINS,
        "overwork_threshold":       OVERWORK_THRESHOLD_HRS,
        "extensive_hours_threshold": EXTENSIVE_HOURS_THRESH,
    }

# ═══════════════════════════════════════════════════════════════
# JSON SERIALISATION
# ═══════════════════════════════════════════════════════════════

def clean(obj):
    """Recursively make obj JSON-safe."""
    if isinstance(obj, bool):                         return obj
    if isinstance(obj, dict):                         return {k: clean(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):                return [clean(v) for v in obj]
    if isinstance(obj, (datetime, pd.Timestamp)):     return obj.strftime("%d/%m/%Y")
    if isinstance(obj, float):
        if obj != obj or obj == float("inf"):         return 0   # nan / inf
        return round(obj, 4)
    if isinstance(obj, (pd.Series, pd.DataFrame)):   return clean(obj.tolist() if hasattr(obj, "tolist") else [])
    try:
        if pd.isna(obj):                              return 0
    except Exception:
        pass
    if obj is None:                                   return None
    if isinstance(obj, (int, str)):                   return obj
    return str(obj)

# ═══════════════════════════════════════════════════════════════
# HTML INJECTION
# ═══════════════════════════════════════════════════════════════

def export_json(dash_data):
    json_str = json.dumps(clean(dash_data), ensure_ascii=False, separators=(",", ":"))
    print(f"\n[>>] Writing  {DATA_JSON_PATH}")
    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        f.write(json_str)
    print(f"   OK data.json written  ({len(json_str):,} bytes)")
    print(f"\n[>>] Writing  {DATA_JS_PATH}")
    with open(DATA_JS_PATH, "w", encoding="utf-8") as f:
        f.write("window.DASH_DATA = " + json_str + ";")
    print(f"   OK data.js written")

# ═══════════════════════════════════════════════════════════════
# AUTO GITHUB PAGES DEPLOYMENT
# ═══════════════════════════════════════════════════════════════

def get_git_executable():
    """Locates git command from PATH or bundled GitHub Desktop installation."""
    import shutil
    if shutil.which("git"):
        return "git"
    github_desktop_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "GitHubDesktop"
    if github_desktop_dir.exists():
        app_dirs = list(github_desktop_dir.glob("app-*"))
        if app_dirs:
            latest_app = sorted(app_dirs)[-1]
            git_path = latest_app / "resources" / "app" / "git" / "cmd" / "git.exe"
            if git_path.exists():
                return str(git_path)
    return "git"

def deploy_to_github(repo_dir=None):
    """
    Automatically commits and pushes updated dashboard files to GitHub Pages.
    Repository: tamdoan2802/PC_attendance_management
    """
    import subprocess
    if repo_dir is None:
        repo_dir = THIS_DIR.parent

    git_dir = repo_dir / ".git"
    if not git_dir.exists():
        print(f"\n[GIT] No .git directory found at {repo_dir}. Skipping auto-deploy.")
        return False

    git_cmd = get_git_executable()
    print(f"\n[>>] Deploying updated dashboard to GitHub Pages...")
    try:
        # Stage dashboard artifacts and scripts
        subprocess.run(
            [git_cmd, "add", "-A"],
            cwd=str(repo_dir),
            check=True,
            capture_output=True,
            text=True
        )

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_msg = f"Auto-deploy workforce dashboard at {timestamp}"
        commit_res = subprocess.run(
            [git_cmd, "commit", "-m", commit_msg],
            cwd=str(repo_dir),
            capture_output=True,
            text=True
        )
        if "nothing to commit" in commit_res.stdout or "nothing to commit" in commit_res.stderr:
            print("   [INFO] GitHub Pages is already up to date with latest data.")
            print("   🌐 Live URL: https://tamdoan2802.github.io/PC_attendance_management/")
            return True

        print("   [>>] Pushing to GitHub (origin/master)...")
        subprocess.run(
            [git_cmd, "push", "origin", "master"],
            cwd=str(repo_dir),
            check=True,
            capture_output=True,
            text=True
        )
        print("   ✅ Successfully deployed to GitHub Pages!")
        print("   🌐 Live URL: https://tamdoan2802.github.io/PC_attendance_management/")
        return True
    except subprocess.CalledProcessError as e:
        err_msg = e.stderr.strip() if e.stderr else str(e)
        print(f"   [WARN] Git auto-deploy encountered an issue: {err_msg}")
        return False
    except Exception as e:
        print(f"   [WARN] Could not deploy to GitHub: {e}")
        return False

# ═══════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════

def main(deploy=True):
    print("=" * 62)
    print("  Attendance Dashboard — Data Regenerator (100% Raw Data & Entities)")
    print("=" * 62)

    raw       = load_all_raw_data()
    proc      = preprocess(raw)
    week_rngs = detect_weeks(proc["att"], n=TRAILING_WEEKS)

    print(f"\n[2] Detected weeks: {[week_range_label(w) for w in week_rngs]}")
    print(f"   Current week  : {week_rngs[-1]}")

    scopes = build_scopes(proc["emp"])
    print(f"   Teams: {len(scopes['teams'])-1}  |  Clients: {len(scopes['clients'])-1}"
          f"  |  Scope keys: {len(scopes['scope_keys'])}")

    dash = build_dash_data(proc, scopes, week_rngs)
    export_json(dash)

    print("\n" + "=" * 62)
    print(f"  [DONE]  Done! Data exported to {DATA_JSON_PATH.name}")
    print("=" * 62)

    if deploy:
        deploy_to_github()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate Workforce Attendance Dashboard Data")
    parser.add_argument("--no-deploy", action="store_true", help="Skip automatic deployment to GitHub Pages")
    args = parser.parse_args()
    main(deploy=not args.no_deploy)
