# Workforce dashboard — metric formulas & source columns

All metrics are computed per **week** (Mon–Fri) and per **Team** scope (`All Teams` or a specific `DIM_Employee.Team`). Trend charts use a fixed 4-week trailing window: current week + 3 prior weeks. Current week in this build = **13/07/2026 – 17/07/2026**.

Team filter mechanic: `FACT_Attendance_Daily` already carries `DIM_Employee.Team` as a column. Every request sheet (`Req_Leave`, `Req_OT`, `Req_WFh`, `Req_LCin&ECout`, `Req_BusinessTrip`, `Req_ShiftChange`) is joined on `Employee_ID = DIM_Employee.EmployeeID` to bring in `Team` before filtering.

---

## Overview / outcome metrics

| Metric | Formula | Sheet | Columns used |
|---|---|---|---|
| **Attendance quality (%)** | `COUNT(Type of Date = "FullWorkDay") / COUNT(all weekday records) × 100` | `FACT_Attendance_Daily` | `Type of Date`, `Is_Weekend` (filter = 0), `Date_Text` (week filter), `DIM_Employee.Team` |
| **Unplanned absence (%)** | `COUNT(Type of Date = "Unpaid Leave") / COUNT(all weekday records) × 100` | `FACT_Attendance_Daily` | `Type of Date`, `Is_Weekend` (filter = 0), `Date_Text`, `DIM_Employee.Team` |
| **Overwork index (employees)** | `COUNT(DISTINCT Employee_ID)` where `AVG(Delta)` for that employee that week `> 1.5` AND `Employee_ID` NOT IN `Req_OT` for that week | `FACT_Attendance_Daily` + `Req_OT` | `Employee_ID`, `Delta (Số giờ làm việc thực tế − Số giờ làm việc tiêu chuẩn)`, `Date_Text`, `Is_Weekend` ; `Req_OT.Employee_ID`, `Req_OT.OT_From` |
| **Active flags** | Sum of the 4 flag counts below (Weekend-bridge leave + Short-notice leave + Overwork-no-OT + Friday shift-change cluster) | multiple | see Flags section |

---

## Leave

| Metric | Formula | Sheet | Columns used |
|---|---|---|---|
| **Leave days** | `SUM(Leave_Days)` for requests where `Leave_From` falls in the week | `Req_Leave` | `Leave_Days`, `Leave_From`, `Team` (joined) |
| **People on leave (headcount)** | `COUNT(DISTINCT Employee_ID)` for requests where `Leave_From` falls in the week | `Req_Leave` | `Employee_ID`, `Leave_From` |
| **Approval rate (%)** | `COUNT(Status = "Đã duyệt") / COUNT(all requests) × 100` | `Req_Leave` | `Status`, `Leave_From` |

---

## OT & WFH

| Metric | Formula | Sheet | Columns used |
|---|---|---|---|
| **OT hours** | `SUM(OT_Hours)` for requests where `OT_From` falls in the week | `Req_OT` | `OT_Hours`, `OT_From` |
| **Weekend OT hours** | `SUM(OT_Hours)` where `OT_Timing = "Ngày nghỉ"` AND `OT_From` in week | `Req_OT` | `OT_Hours`, `OT_Timing`, `OT_From` |
| **WFH days** | `COUNT(rows)` where `WFH_From` falls in the week | `Req_WFh` | `WFH_From` (row count = 1 request ≈ 1 WFH instance) |

---

## Requests ops

| Metric | Formula | Sheet | Columns used |
|---|---|---|---|
| **Late CI / Early CO events** | `COUNT(rows)` where `Apply_From` falls in the week | `Req_LCin&ECout` | `Apply_From` |
| **Average Late CI (mins)** | `AVG(Late_CI (mins))` across all weekday attendance records in the week (NaN treated as 0) | `FACT_Attendance_Daily` | `Late_CI (mins)`, `Is_Weekend`, `Date_Text` |
| **Average Late CO (mins)** | `AVG(Late_CO (mins))` across all weekday attendance records in the week (NaN treated as 0) | `FACT_Attendance_Daily` | `Late_CO (mins)`, `Is_Weekend`, `Date_Text` |
| **Business trips** | `COUNT(rows)` where `Trip_From` falls in the week | `Req_BusinessTrip` | `Trip_From` |
| **Shift changes** | `COUNT(rows)` where `Work_Date` falls in the week | `Req_ShiftChange` | `Work_Date` |

---

## Flags (Risk tab, current week only)

| Flag | Trigger condition | Sheet | Columns used |
|---|---|---|---|
| **Overwork, no OT filed** | Employee's `AVG(Delta)` for the week `> 1.5` hrs/day AND no matching row in `Req_OT` for that employee that week | `FACT_Attendance_Daily` + `Req_OT` | `Delta (...)`, `Employee_ID`, `Date_Text` ; `Req_OT.Employee_ID`, `OT_From` |
| **Weekend-bridge leave** | `Has_MonFri = TRUE` for a leave request in the week | `Req_Leave` | `Has_MonFri`, `Leave_From` |
| **Short-notice leave** | `Notice before < 1` (day) | `Req_Leave` | `Notice before`, `Leave_From` |
| **Friday capacity reduction** | `COUNT(shift-change requests where Work_Date is a Friday) ≥ 3` in the week | `Req_ShiftChange` | `Work_Date` (filtered to weekday = Friday) |
| **OT escalation** | Current week's weekend OT hours `≥ 2 ×` prior week's weekend OT hours (only fires if prior week > 0) | `Req_OT` | `OT_Hours`, `OT_Timing = "Ngày nghỉ"`, `OT_From` (current week vs. prior week) |

---

## Sparklines & trend charts

All sparklines and Trends-tab line/combo charts plot the same 4 weekly values used in the tables above — no separate calculation, just the array of 4 week-by-week results for that metric and scope.

**Combo chart (Leave days vs. People on leave)**: bars = `Leave days` (as above), line = `People on leave` (as above), same week buckets, same `Req_Leave` filter.

---

## Notes on source data quirks

- `Late_CI (mins)` / `Late_CO (mins)` are used over their `Late_CheckIn (mins)` / `Late_CheckOut(mins)` counterparts per the column-alias convention in the source workbook — use whichever is populated if adapting this elsewhere.
- `Delta (...)` full column name in the workbook is `Delta (Số giờ làm việc thực tế - Số giờ làm việc tiêu chuẩn)` — actual hours worked minus standard hours; positive = worked more than standard.
- Week boundaries are Mon–Fri (weekend rows excluded via `Is_Weekend = 0` for `FACT_Attendance_Daily`-based metrics); request-based metrics (Leave, OT, WFH, etc.) are bucketed by their `*_From` / `Work_Date` falling within the Mon–Fri calendar range, regardless of `Is_Weekend`.
- `Status` values are in Vietnamese as exported by MISA: `Đã duyệt` = Approved, `Chờ duyệt` = Pending, `Từ chối` = Rejected.
