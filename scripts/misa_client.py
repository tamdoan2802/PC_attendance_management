# -*- coding: utf-8 -*-
"""
MISA AMIS Timesheet Direct REST API Client
=========================================
Direct API integration client for MyTeam Vietnam's MISA AMIS HR / Timesheet platform.
Completely replaces legacy Excel exports by streaming structured operational data
directly into memory as Pandas DataFrames.
"""

import os
import sys
import re
import json
import base64
import urllib.request
import urllib.error
import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple, Union

import pandas as pd

# Force UTF-8 on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

BASE_URL = "https://amisapp.misa.vn/APIS/g1/TimeSheetAPI/api"

STATUS_MAP = {
    1: "Chờ duyệt",
    2: "Đã duyệt",
    3: "Từ chối",
    4: "Hủy bỏ"
}

COLUMNS_MAP = {
    "Attendance": (
        "AttendanceID,EmployeeID,EmployeeCode,FullName,JobPositionID,JobPositionName,OrganizationUnitID,"
        "OrganizationUnitName,JobTitleID,RequestDate,DictionaryKey,FromDate,ToDate,LeaveDay,NumberOfHourLeave,"
        "AttendanceTypeID,AttendanceTypeName,AttendanceTypeName_EN,Reason,ApprovalName,Status,SubstituteName,"
        "RelationShipNames,Description,SalaryRate,DataSourceID,ApprovalToID,Step,NextStep,IsProcess,"
        "ProcessApprovalOldID,StepOld,NextStepOld,IsApplyProcessNew,UpdateWorkingProcess,SubstituteID,"
        "NoteAttendanceDay,RelationShipIDs,TotalLeaved,NumRemain,NumLeave,ThisMonth,NextMonth,AttendanceData,"
        "JobTitleName,ProcessApprovalOldValue,IsApply,EmployeeAttendance,EmployeeAttendanceIDs,"
        "EmployeeAttendanceCodes,EmployeeAttendanceNames,ShowEmployeeAttendance,WorkLocationName,IsAttachedFile,"
        "TimeZone,SyncID,ProcessDetail,IsViolateDeadline,ViolateDeadlineType,ViolateDeadlineValue"
    ),
    "OverTime": (
        "OverTimeID,EmployeeID,EmployeeCode,FullName,JobPositionName,ApprovalName,OrganizationUnitID,"
        "OrganizationUnitName,ApplyDate,BreakTimeFrom,BreakTimeTo,FromDate,ToDate,Reason,Status,"
        "WorkingShiftName,OverTimeInWorkingShiftName,Description,DataSourceID,ApprovalToID,Step,NextStep,"
        "IsProcess,ProcessApprovalOldID,StepOld,NextStepOld,IsApplyProcessNew,JobPositionID,JobTitleID,"
        "WorkingShiftID,OverTimeInWorkingShift,RelationShipNames,OvertimeType,OvertimeTypeName,TotalTime,"
        "EmployeeOverTimeNames,EmployeeOverTimeCodes,EmployeeOverTimeIDs,WorkingShiftCode,JobTitleName,"
        "ProcessApprovalOldValue,IsApply,WorkLocationName,ShowEmployeeOverTime,IsAttachedFile,TimeZone,"
        "IsMultipleDate,BreakTimeStart,BreakTimeEnd,RelationShipIDs,FromTime,ToTime,ProcessDetail,"
        "IsViolateDeadline,ViolateDeadlineType,ViolateDeadlineValue"
    ),
    "WorkRemote": (
        "WorkRemoteID,EmployeeID,EmployeeCode,FullName,OrganizationUnitID,OrganizationUnitName,JobPositionID,"
        "JobTitleID,JobPositionName,JobTitleName,RequestDate,FromDate,ToDate,Reason,WorkingShiftIDs,"
        "WorkingShiftNames,ApplyDay,ApplyDayText,ApprovalToID,ApprovalName,RelationShipIDs,RelationShipNames,"
        "Status,Step,NextStep,IsProcess,ProcessApprovalOldID,StepOld,NextStepOld,IsApplyProcessNew,"
        "DataSourceID,TenantID,ProcessApprovalOldValue,IsApply,Description,IsAttachedFile,IsAttachedFile,"
        "TimeZone,ProcessDetail,WorkingShiftCodes,NumberOfRemoteDays,ModifyRemoteDays,IsViolateDeadline,"
        "ViolateDeadlineType,ViolateDeadlineValue"
    ),
    "LateInEarlyOut": (
        "ReasonGroup,ReasonGroupText,LateInEarlyOutID,EmployeeID,EmployeeCode,FullName,JobPositionName,"
        "OrganizationUnitID,OrganizationUnitName,ApprovalName,ApplyDate,WorkingShiftNames,CheckInLateStartTime,"
        "CheckOutEarlyStartBreakTime,CheckInLateEndBreakTime,CheckOutEarlyEndTimeint,ApplyDay,ApplyDayText,"
        "Status,FromDate,ToDate,Reason,RelateNames,Description,DataSourceID,ApprovalToID,Step,NextStep,"
        "IsProcess,ProcessApprovalOldID,StepOld,NextStepOld,IsApplyProcessNew,JobPositionID,JobTitleID,"
        "WorkingShiftIDs,JobTitleName,EmployeeLateInEarlyOutCodes,EmployeeLateInEarlyOutNames,"
        "EmployeeLateInEarlyOutIDs,EmployeeLateInEarlyOut,ProcessApprovalOldValue,IsApply,"
        "ShowEmployeeLateInEarlyOut,IsAttachedFile,TimeZone,RelateIDs,ProcessDetail,WorkingShiftCodes,"
        "IsViolateDeadline,ViolateDeadlineType,ViolateDeadlineValue"
    ),
    "MissionAllowance": (
        "MissionAllowanceID,EmployeeID,EmployeeCode,FullName,JobPositionName,OrganizationUnitID,"
        "OrganizationUnitName,ApprovalName,RequestDate,FromDate,ToDate,Location,Purpose,Status,Request,"
        "SupportNames,RelationShipNames,Description,LeaveDay,DataSourceID,ApprovalToID,Step,NextStep,"
        "IsProcess,ProcessApprovalOldID,StepOld,NextStepOld,IsApplyProcessNew,JobPositionID,JobTitleID,"
        "LocationID,EmployeeMisionIDs,EmployeeMisionCodes,EmployeeMisionNames,RelationShipIDs,SupportIDs,"
        "ShowEmployeeMissionAllowance,JobTitleName,ProcessApprovalOldValue,IsApply,WorkLocationName,"
        "RequestedAdvanceAmount,EmployeeMissionAllowances,IsAttachedFile,TimeZone,ProcessStatus,ProcessInfo,"
        "AdvanceApproverID,ProcessDetail,IsViolateDeadline,ViolateDeadlineType,ViolateDeadlineValue"
    ),
    "ChangeShift": (
        "ChangeShiftID,EmployeeID,EmployeeCode,ApprovalToID,ApprovalName,Status,Description,TenantID,"
        "ModifiedDate,ModifiedBy,CreatedDate,CreatedBy,EditVersion,FullName,OrganizationUnitID,"
        "OrganizationUnitName,JobPositionID,JobPositionName,RequestDate,WorkingDate,WorkingShiftIDFrom,"
        "WorkingShiftCodeFrom,WorkingShiftNameFrom,ChangeDate,WorkingShiftIDTo,WorkingShiftCodeTo,"
        "WorkingShiftNameTo,EmployeeIDChange,EmployeeNameChange,Reason,RejectReason,Step,NextStep,IsProcess,"
        "ForwarderIDs,ForwarderNames,JobTitleID,JobTitleName,DataSourceID,EmployeeChangeShiftIDs,"
        "EmployeeChangeShiftCodes,EmployeeChangeShiftNames,ShowEmployeeChangeShift,ProcessApprovalOldID,"
        "StepOld,NextStepOld,IsApplyProcessNew,WorkLocationID,WorkLocationCode,WorkLocationName,AttachedFile,"
        "IsAttachedFile,TimeZone,RelationShipNames,RelationShipIDs,SyncID,ProcessDetail"
    )
}

class MisaAuthError(Exception):
    """Raised when MISA authentication fails or session expires."""
    pass

class MisaAmisClient:
    """Client for executing direct API operations on MISA AMIS Timesheet platform."""

    def __init__(self, token_file: Optional[str] = None):
        if token_file:
            self.token_file = Path(token_file)
        else:
            self.token_file = Path(__file__).parent / ".token_misa"
        
        self.cookie_string = self._load_token()
        self.session_id, self.tenant_id = self._extract_session_headers(self.cookie_string)

    def _load_token(self) -> str:
        """Load session cookie from environment variable or local .token_misa file."""
        env_token = os.environ.get("MISA_TOKEN")
        if env_token:
            return env_token.strip()

        candidate_paths = [
            self.token_file,
            Path(__file__).parent.parent / ".token_misa",
            Path(__file__).parent.parent.parent.parent.parent / "attendance reference" / "MisaSetup" / ".token_misa",
        ]
        for p in candidate_paths:
            if p and p.exists():
                with open(p, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                    if content:
                        return self.parse_token_string(content)

        raise MisaAuthError(
            "No MISA AMIS session token found! Please set MISA_TOKEN or save your session cookies into .token_misa"
        )

    @staticmethod
    def parse_token_string(raw_input: str) -> str:
        """Extract cookie string from raw string or cURL command."""
        raw_input = raw_input.strip()

        # Case 1: cURL command containing -b '...' or -H 'Cookie: ...'
        cookie_match = re.search(r"(?:-b|-H\s+['\"]Cookie:)\s+['\"]([^'\"]+)['\"]", raw_input, re.IGNORECASE)
        if cookie_match:
            return cookie_match.group(1).strip()

        # Case 2: Raw Cookie string containing x-sessionid=
        if "x-sessionid=" in raw_input:
            return raw_input

        return raw_input

    def update_token(self, new_token_or_curl: str) -> str:
        """Update active session token from raw cookie string or cURL command."""
        parsed = self.parse_token_string(new_token_or_curl)
        self.cookie_string = parsed
        self.session_id, self.tenant_id = self._extract_session_headers(parsed)
        with open(self.token_file, "w", encoding="utf-8") as f:
            f.write(parsed)
        print("MISA AMIS session token updated successfully!")
        return parsed

    @staticmethod
    def _extract_session_headers(cookie_str: str) -> Tuple[str, str]:
        sid = ""
        tid = ""
        m_sid = re.search(r"x-sessionid=([a-zA-Z0-9_\-]+)", cookie_str)
        if m_sid:
            sid = m_sid.group(1)
        m_tid = re.search(r"x-tenantid=([a-zA-Z0-9_\-]+)", cookie_str)
        if m_tid:
            tid = m_tid.group(1)
        return sid, tid

    def request(
        self,
        endpoint: str,
        payload: Optional[Dict[str, Any]] = None,
        method: str = "POST",
        timeout: int = 45,
    ) -> Dict[str, Any]:
        """Execute authenticated HTTPS request against MISA AMIS API."""
        url = f"{BASE_URL}/{endpoint}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0",
            "Cookie": self.cookie_string,
            "x-sessionid": self.session_id,
            "x-tenantid": self.tenant_id,
            "x-tenantsource": "AMIS",
            "Content-Type": "application/json; charset=UTF-8",
            "Accept": "application/json, text/plain, */*",
        }

        data = json.dumps(payload).encode("utf-8") if payload is not None else None
        req = urllib.request.Request(url, data=data, headers=headers, method=method)

        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8")
                return json.loads(raw)
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                raise MisaAuthError(
                    f"MISA AMIS Session Token has expired or is invalid (HTTP {e.code}).\n"
                    "How to refresh (10 seconds):\n"
                    "1. Open Chrome/Edge and log into https://amisapp.misa.vn\n"
                    "2. Press F12 -> Network tab -> click any request\n"
                    "3. Right-click -> Copy -> Copy as cURL (bash)\n"
                    "4. Run: client.update_token('PASTE_CURL_HERE') or save to attendance reference/MisaSetup/.token_misa"
                ) from e
            err_body = ""
            try:
                err_body = e.read().decode("utf-8")
            except Exception:
                pass
            raise RuntimeError(f"HTTP Error {e.code} on {endpoint}: {err_body}") from e

    # ═══════════════════════════════════════════════════════════════
    # DOMAIN METHODS
    # ═══════════════════════════════════════════════════════════════

    def get_timesheet_list(self, page_size: int = 20) -> List[Dict[str, Any]]:
        """Retrieve list of all timesheet periods from MISA."""
        payload = {
            "PageSize": page_size,
            "PageIndex": 1,
            "Filter": "",
            "CustomFilter": None,
            "QuickSearch": None,
            "CustomParam": {"OrganizationUnitIDs": None, "InActive": False, "JobPositionID": None},
            "Sort": base64.b64encode(b'[{"selector":"FromDate","desc":true}]').decode("utf-8"),
            "Columns": None,
        }
        res = self.request("TimeSheet/v2/datapaging", payload)
        return res.get("Data", [])

    def get_timesheet_daily_dataframe(self, timesheet_id: Optional[Union[int, List[int]]] = None, num_periods: int = 3) -> pd.DataFrame:
        """
        Stream daily attendance punch grid directly into a standardized FACT_Attendance_Daily DataFrame.
        Supports fetching multiple periods (default: latest 3 started periods) so that trailing weekly metrics
        and trend charts have complete multi-week historical attendance data.
        """
        import datetime as _dt
        today = _dt.date.today()

        ts_list = self.get_timesheet_list(page_size=15)
        ts_dict = {t["TimeSheetID"]: t for t in ts_list}

        target_periods = []
        if timesheet_id is not None:
            if isinstance(timesheet_id, int):
                target_ids = [timesheet_id]
            elif isinstance(timesheet_id, (list, tuple)):
                target_ids = list(timesheet_id)
            else:
                target_ids = [int(timesheet_id)]
            for tid in target_ids:
                if tid in ts_dict:
                    target_periods.append(ts_dict[tid])
                else:
                    target_periods.append({"TimeSheetID": tid})
        else:
            if not ts_list:
                raise RuntimeError("No timesheets found on MISA AMIS!")
            # Pick the most recent periods whose FromDate has already started (FromDate <= today)
            started_ts = [
                ts for ts in ts_list
                if pd.to_datetime(ts.get("FromDate", "")).date() <= today
            ]
            if not started_ts:
                started_ts = [ts_list[0]]
            target_periods = started_ts[:num_periods]

        print(f"  Ingesting {len(target_periods)} Timesheet period(s):")
        for tp in target_periods:
            print(f"    - ID {tp['TimeSheetID']} ({tp.get('TimeSheetName', '')[:45]}...) "
                  f"[{str(tp.get('FromDate','?'))[:10]} \u2192 {str(tp.get('ToDate','?'))[:10]}]")

        records = []
        for tp in target_periods:
            tid = tp["TimeSheetID"]
            ts_from = None
            ts_to = None
            if tp.get("FromDate"):
                try:
                    ts_from = pd.to_datetime(tp["FromDate"]).date()
                except Exception:
                    pass
            if tp.get("ToDate"):
                try:
                    ts_to = pd.to_datetime(tp["ToDate"]).date()
                except Exception:
                    pass

            filter_b64 = base64.b64encode(json.dumps([["TimeSheetID", "=", tid]]).encode("utf-8")).decode("utf-8")
            page_idx = 1
            all_emp_rows = []

            while True:
                payload = {
                    "PageIndex": page_idx,
                    "PageSize": 50,
                    "Filter": filter_b64,
                    "Sort": None,
                    "CustomParam": {"TimeSheetID": tid},
                }
                res = self.request("TimeSheetDetail/paging", payload)
                if not res.get("Success") or not res.get("Data"):
                    break
                pdata = res["Data"].get("PageData", [])
                if not pdata:
                    break
                all_emp_rows.extend(pdata)
                total = res["Data"].get("Total", 0)
                if len(all_emp_rows) >= total or len(pdata) < 50:
                    break
                page_idx += 1

            for emp_row in all_emp_rows:
                emp_id = str(emp_row.get("EmployeeCode") or "").strip()
                emp_name = str(emp_row.get("FullName") or "").strip()
                pos = str(emp_row.get("JobPositionName") or "").strip()
                org = str(emp_row.get("OrganizationUnitName") or "").strip()

                for d in range(1, 32):
                    expected_date = None
                    if ts_from is not None:
                        expected_date = ts_from + _dt.timedelta(days=d - 1)
                        if ts_to is not None and expected_date > ts_to:
                            continue

                    day_json = emp_row.get(f"Day{d}", "[]")
                    if not day_json or day_json == "[]":
                        continue
                    try:
                        shifts = json.loads(day_json)
                    except Exception:
                        continue

                    for s in shifts:
                        dt_str = s.get("CheckStartTime") or s.get("CheckEndTime")
                        date_val = None
                        if dt_str:
                            try:
                                date_val = pd.to_datetime(dt_str[:10])
                            except Exception:
                                pass
                        if date_val is None and expected_date is not None:
                            date_val = pd.to_datetime(expected_date)

                        if date_val is None:
                            continue

                        # Do not ingest future days from pre-generated timesheet templates
                        if date_val.date() > today:
                            continue

                        ci_time = ""
                        if s.get("CheckStartTime"):
                            try:
                                ci_time = s.get("CheckStartTime")[11:16]
                            except Exception:
                                pass

                        co_time = ""
                        if s.get("CheckEndTime"):
                            try:
                                co_time = s.get("CheckEndTime")[11:16]
                            except Exception:
                                pass

                        shift_start = s.get("StartTime", "")[:5] if s.get("StartTime") else ""
                        shift_end = s.get("EndTime", "")[:5] if s.get("EndTime") else ""

                        working_credit = float(s.get("Working", 0.0) or 0.0)
                        actual_hours = float(s.get("ActualHours", 0.0) or 0.0)
                        late_ci = float(s.get("LastInTime", 0.0) or 0.0)
                        early_co = float(s.get("FirstOutTime", 0.0) or 0.0)
                        is_manual = bool(s.get("IsUpdateTimeKeeper", False))
                        shift_code = str(s.get("WorkingShiftCode") or "").strip()

                        # Calculate Late CheckOut
                        late_co = 0.0
                        if co_time and shift_end:
                            try:
                                h_co, m_co = map(int, co_time.split(":"))
                                h_se, m_se = map(int, shift_end.split(":"))
                                diff_co = (h_co * 60 + m_co) - (h_se * 60 + m_se)
                                if diff_co > 0:
                                    late_co = float(diff_co)
                            except Exception:
                                pass

                        std_hours = 8.0
                        if shift_code in ["WE08"]:
                            std_hours = 6.0
                        elif "WFH" in shift_code:
                            std_hours = 3.5 if "WE07" in shift_code else (3.0 if "WE08" in shift_code else 8.0)

                        # Credit working hours for business trip (Đi công tác)
                        total_mission = float(s.get("TotalMissionAllowance", 0.0) or 0.0)
                        if total_mission > 0:
                            mission_hours = total_mission * std_hours
                            if total_mission >= 1.0:
                                actual_hours = max(actual_hours, mission_hours)
                            else:
                                actual_hours = max(actual_hours + mission_hours, mission_hours)

                        # Type of Date
                        is_wk = 1 if date_val.weekday() >= 5 else 0

                        type_date = "Absent"
                        tot_leave = float(s.get("TotalLeave", 0.0) or 0.0)
                        tot_actual_work = float(s.get("TotalWorkingActual", 0.0) or 0.0)

                        if total_mission > 0:
                            type_date = "FullWorkDay" if (total_mission >= 1.0 or working_credit >= 1.0) else "PartialWorkDay"
                        elif tot_leave >= 1.0 and tot_actual_work == 0:
                            type_date = "Leave"
                        elif working_credit >= 1.0:
                            type_date = "FullWorkDay"
                        elif working_credit > 0:
                            type_date = "PartialWorkDay"
                        elif tot_leave > 0:
                            type_date = "Leave"
                        elif is_wk == 1:
                            type_date = "Weekend"

                        records.append({
                            "Employee_ID": emp_id,
                            "Employee_Name": emp_name,
                            "Position": pos,
                            "Department": org,
                            "Date_Text": date_val,
                            "Day_Num": date_val.day,
                            "Day_Of_Week": date_val.strftime("%a"),
                            "Shift_Code": shift_code,
                            "Working_Days": working_credit,
                            "CheckIn_Time": ci_time,
                            "CheckOut_Time": co_time,
                            "DIM_Shift.Start_Time": shift_start,
                            "DIM_Shift.End_Time": shift_end,
                            "Late_CheckIn (mins)": late_ci,
                            "Late_CheckOut(mins)": late_co,
                            "Late_CI (mins)": late_ci,
                            "Late_CO (mins)": late_co,
                            "Early_CI (mins)": 0.0,
                            "Early_CO (mins)": early_co,
                            "Type of Date": type_date,
                            "Số ngày làm việc tiêu chuẩn": 22,
                            "Số giờ làm việc tiêu chuẩn": std_hours,
                            "Số giờ làm việc thực tế": actual_hours,
                            "Is_Weekend": is_wk,
                            "Is_Manual_Update": is_manual,
                            "Delta (Số giờ làm việc thực tế - Số giờ làm việc tiêu chuẩn)": (actual_hours - std_hours),
                        })

        df = pd.DataFrame(records)
        if not df.empty:
            df = df.drop_duplicates(subset=["Employee_ID", "Date_Text", "Shift_Code"], keep="last")
            for c in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[c]):
                    if hasattr(df[c].dt, "tz") and df[c].dt.tz is not None:
                        df[c] = df[c].dt.tz_localize(None)

            if "Date_Text" in df.columns:
                def calc_wp(dt):
                    if pd.isna(dt): return ""
                    m = dt - pd.Timedelta(days=dt.weekday())
                    f = m + pd.Timedelta(days=4)
                    return f"{m.strftime('%d/%m/%Y')} - {f.strftime('%d/%m/%Y')}"
                df["Week Period"] = df["Date_Text"].apply(calc_wp)
        return df

    def get_requests_dataframe(self, subsystem: str, from_date: str = "2026-01-01 00:00:00") -> pd.DataFrame:
        """
        Stream request subsystem directly into a standardized DataFrame matching generate_dashboard_data schema.
        Supported subsystems: 'Attendance', 'OverTime', 'WorkRemote', 'LateInEarlyOut', 'MissionAllowance', 'ChangeShift'.
        """
        field_map = {
            "Attendance": "RequestDate",
            "OverTime": "FromDate",
            "WorkRemote": "FromDate",
            "LateInEarlyOut": "FromDate",
            "MissionAllowance": "FromDate",
            "ChangeShift": "WorkingDate",
        }
        date_col = field_map.get(subsystem, "FromDate")
        filter_b64 = base64.b64encode(json.dumps([[date_col, ">=", from_date]]).encode("utf-8")).decode("utf-8")
        columns_str = COLUMNS_MAP.get(subsystem)

        all_items = []
        page_idx = 1
        while True:
            payload = {
                "PageIndex": page_idx,
                "PageSize": 200,
                "Filter": filter_b64,
                "CustomFilter": None,
                "QuickSearch": None,
                "CustomParam": {"OrganizationUnitID": None, "Status": 0},
                "Sort": None,
                "Columns": columns_str,
            }
            res = self.request(f"{subsystem}/v2/datapaging", payload)
            items = res.get("Data", [])
            if not items:
                break
            all_items.extend(items)
            if len(items) < 200 or page_idx >= 5: # safety cap 1,000 records
                break
            page_idx += 1

        def date_to_wp(dt):
            if pd.isna(dt): return ""
            try:
                dt = pd.to_datetime(dt)
                m = dt - pd.Timedelta(days=dt.weekday())
                f = m + pd.Timedelta(days=4)
                return f"{m.strftime('%d/%m/%Y')} - {f.strftime('%d/%m/%Y')}"
            except Exception:
                return ""

        rows = []
        for r in all_items:
            eid = str(r.get("EmployeeCode") or "").strip()
            if not eid or eid == "nan":
                continue
            ename = str(r.get("FullName") or "").strip()
            st_code = r.get("Status", 0)
            st_str = STATUS_MAP.get(st_code, "Đã duyệt" if st_code == 2 else "Chờ duyệt")
            st_flag = 1 if st_code == 2 else 0

            if subsystem == "Attendance":
                t_from = pd.to_datetime(r.get("FromDate"), errors="coerce")
                t_to = pd.to_datetime(r.get("ToDate"), errors="coerce")
                sub_dt = pd.to_datetime(r.get("RequestDate"), errors="coerce")
                if pd.isna(t_from): continue
                if pd.isna(t_to): t_to = t_from
                if t_from > t_to: t_from, t_to = t_to, t_from

                l_type = str(r.get("AttendanceTypeName") or "Nghỉ phép").strip()
                if "phép" in l_type.lower(): l_mapped = "Annual Leave"
                elif "không lương" in l_type.lower(): l_mapped = "Unpaid Leave"
                elif "bù" in l_type.lower(): l_mapped = "Compensatory Leave"
                elif any(k in l_type.lower() for k in ["ốm", "bhxh", "con ốm"]): l_mapped = "Sick/Family Care"
                else: l_mapped = l_type

                leave_days = float(r.get("LeaveDay", 1.0) or 1.0)
                notice_days = (t_from - sub_dt).total_seconds() / 86400.0 if pd.notna(sub_dt) else 0

                rows.append({
                    "Employee_ID": eid,
                    "Employee_Name": ename,
                    "Leave_From": t_from,
                    "Leave_To": t_to,
                    "Submit_Date": sub_dt,
                    "Leave_Days": leave_days,
                    "Leave Days in Week": leave_days,
                    "Leave_Type": l_type,
                    "Leave_Type_Raw": l_type,
                    "Leave_Type_Mapped": l_mapped,
                    "Reason": str(r.get("Reason") or "").strip(),
                    "Status": st_str,
                    "Status_Flag": st_flag,
                    "Notice before": notice_days,
                    "Has_MonFri": 1 if (t_from.weekday() in [0, 4] or t_to.weekday() in [0, 4]) else 0,
                    "Week Period": date_to_wp(t_from)
                })

            elif subsystem == "OverTime":
                t_from = pd.to_datetime(r.get("FromDate"), errors="coerce")
                t_to = pd.to_datetime(r.get("ToDate"), errors="coerce")
                if pd.isna(t_from): continue
                if pd.isna(t_to): t_to = t_from

                ot_hours = float(r.get("TotalTime", 0.0) or 0.0)
                ot_type = str(r.get("OvertimeTypeName") or "Hưởng lương").strip()
                ot_timing = str(r.get("OverTimeInWorkingShiftName") or ("Ngày nghỉ" if t_from.weekday() >= 5 else "Sau ca làm việc")).strip()

                rows.append({
                    "Employee_ID": eid,
                    "Employee_Name": ename,
                    "OT_From": t_from,
                    "OT_To": t_to,
                    "OT_Hours": ot_hours,
                    "OT_Type": ot_type,
                    "OT_Timing": ot_timing,
                    "Reason": str(r.get("Reason") or "").strip(),
                    "Status": st_str,
                    "Status_Flag": st_flag,
                    "Week Period": date_to_wp(t_from)
                })

            elif subsystem == "WorkRemote":
                t_from = pd.to_datetime(r.get("FromDate"), errors="coerce")
                t_to = pd.to_datetime(r.get("ToDate"), errors="coerce")
                if pd.isna(t_from): continue
                if pd.isna(t_to): t_to = t_from

                rows.append({
                    "Employee_ID": eid,
                    "Employee_Name": ename,
                    "WFH_From": t_from,
                    "WFH_To": t_to,
                    "Reason": str(r.get("Reason") or "").strip(),
                    "Status": st_str,
                    "Status_Flag": st_flag,
                    "Week Period": date_to_wp(t_from)
                })

            elif subsystem == "LateInEarlyOut":
                t_from = pd.to_datetime(r.get("ApplyDate") or r.get("FromDate"), errors="coerce")
                if pd.isna(t_from): continue

                m_late = float(r.get("CheckInLateStartTime", 0) or 0)
                m_early = float(r.get("CheckOutEarlyEndTimeint", 0) or 0)

                rows.append({
                    "Employee_ID": eid,
                    "Employee_Name": ename,
                    "Apply_From": t_from,
                    "Minutes": m_late + m_early,
                    "CI_Category_Mapped": "Late" if m_late > 0 else "On Time",
                    "CO_Category_Mapped": "Early" if m_early > 0 else "On Time",
                    "Reason_Detail": str(r.get("Reason") or "").strip(),
                    "Reason_Group": str(r.get("ReasonGroupText") or "").strip(),
                    "Status": st_str,
                    "Status_Flag": st_flag,
                    "Week Period": date_to_wp(t_from)
                })

            elif subsystem == "MissionAllowance":
                t_from = pd.to_datetime(r.get("FromDate"), errors="coerce")
                t_to = pd.to_datetime(r.get("ToDate"), errors="coerce")
                if pd.isna(t_from): continue
                if pd.isna(t_to): t_to = t_from

                rows.append({
                    "Employee_ID": eid,
                    "Employee_Name": ename,
                    "Trip_From": t_from,
                    "Trip_To": t_to,
                    "Trip_Days": float(r.get("LeaveDay", 1.0) or 1.0),
                    "Destination": str(r.get("Location") or "").strip(),
                    "Purpose": str(r.get("Purpose") or "").strip(),
                    "Status": st_str,
                    "Status_Flag": st_flag,
                    "Week Period": date_to_wp(t_from)
                })

            elif subsystem == "ChangeShift":
                w_date = pd.to_datetime(r.get("WorkingDate"), errors="coerce")
                if pd.isna(w_date): continue

                rows.append({
                    "Employee_ID": eid,
                    "Employee_Name": ename,
                    "Work_Date": w_date,
                    "Shift_Code_Old": str(r.get("WorkingShiftCodeFrom") or "").strip(),
                    "Shift_Code_New": str(r.get("WorkingShiftCodeTo") or "").strip(),
                    "Reason": str(r.get("Reason") or "").strip(),
                    "Status": st_str,
                    "Status_Flag": st_flag,
                    "Week Period": date_to_wp(w_date)
                })

        df = pd.DataFrame(rows)
        for c in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[c]):
                if hasattr(df[c].dt, "tz") and df[c].dt.tz is not None:
                    df[c] = df[c].dt.tz_localize(None)
        return df

    def get_attendance_dataframes(self, timesheet_id: Optional[Union[int, List[int]]] = None, num_periods: int = 3) -> Dict[str, pd.DataFrame]:
        """
        Directly stream all 7 attendance datasets into memory as clean Pandas DataFrames.
        Returns dictionary with keys:
          - 'FACT_Attendance_Daily'
          - 'Req_Leave'
          - 'Req_OT'
          - 'Req_WFh'
          - 'Req_LCin&ECout'
          - 'Req_BusinessTrip'
          - 'Req_ShiftChange'
        """
        print("  -> Ingesting Timesheet Detail directly from MISA API...")
        df_att = self.get_timesheet_daily_dataframe(timesheet_id, num_periods=num_periods)

        print("  -> Ingesting 6 Request Subsystems from MISA API...")
        df_leave = self.get_requests_dataframe("Attendance")
        df_ot = self.get_requests_dataframe("OverTime")
        df_wfh = self.get_requests_dataframe("WorkRemote")
        df_lc = self.get_requests_dataframe("LateInEarlyOut")
        df_trip = self.get_requests_dataframe("MissionAllowance")
        df_sc = self.get_requests_dataframe("ChangeShift")

        return {
            "FACT_Attendance_Daily": df_att,
            "Req_Leave": df_leave,
            "Req_OT": df_ot,
            "Req_WFh": df_wfh,
            "Req_LCin&ECout": df_lc,
            "Req_BusinessTrip": df_trip,
            "Req_ShiftChange": df_sc,
        }

if __name__ == "__main__":
    print("=" * 65)
    print("      TESTING MISA AMIS DIRECT REST API CLIENT")
    print("=" * 65)
    client = MisaAmisClient()
    dfs = client.get_attendance_dataframes()
    for name, df in dfs.items():
        print(f"  [✓] {name.ljust(25)}: {len(df):,} records loaded directly into RAM")
    print("=" * 65)
