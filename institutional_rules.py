"""
Institutional Attendance Rules & Special Exceptions Engine for RVS / SVCET
=============================================================================
Encapsulates all university-specific attendance policies, departmental shift rules,
and employee-specific exemptions:

1. 100% Full-Month Present Exemptions (Always injected even if missing in sheet):
   - Principal Sir (ID 101 / Dr. M Mohan Babu): All days present (Full month).
   - IT (ID 707 / R. Gunasekaran): No biometric needed, full month present.
   - HAS Department (ID 900 / Dr. I. Sudarsan Kumar): Full month present.
   - Administration (ID 1060 / Vivekand Adhikari / Veveka Thadikar): Full month present.
   - Exam Section (ID 1015 / R Hari Krishna): Full month present.
   - TAP Cell (ID 1019 / Bishal Kumar Sha / Vishal Kumar): Full month present.
   - Management Staff (ID 1021 / M.P.Balaji, ID 4001 / R Poorna Chandra, S. Shajahan / MGT_1095): Full month present.
   - Drivers:
       * Principal Driver (D. Siva / Sivaiah / ID 6623 / ID 603 / ID 4024): Full month present.
       * VC Driver (ID 1030 / Thangeeru Surendera): Full month present.

2. Custom Daily Punch Rules:
   - CSE ID 536 (Bala Subramaniyan): Punch in before 12:10 PM -> Full Day Present.
   - Civil ID 109 (M. Leelakar / Lilaakar): Punch in before 11:00 AM -> Full Day Present.
   - Media Team ID 1018 (B. Prudhvi Raj): Punch in up to 9:35 AM -> Full Day Present.
   - Electricians (IDs: 206, 243, 218, 217, 213):
       * Normal in time 9:25 AM.
       * Early shift: In <= 8:30 AM and Out >= 4:30 PM (16:30) -> Full Day Present.
   - Attenders: In <= 8:35 AM and Out >= 5:30 PM (17:30) -> Full Day Present.
   - Garden Staff: In <= 8:35 AM and Out >= 5:10 PM (17:10) -> Full Day Present.
   - Transport Dept (IDs: 625, 626, 627, 648, 1198, 628, 622, 6621, 606, 623, 603, 653, 605, 6623, 607, 6633, 610, 613):
       * Any day with valid IN and OUT punch counts as Full Day Present (no strict hour limits).

3. Schedule & Quota Rules:
   - IT ID 1053 (Pachayapan): Thu, Fri, Sat schedule. If present 3 days/week or >= 12 days in month -> Full Month Present.
   - IT HOD ID 1203 (Dr J Velmurugan): Wed-Sat schedule (Mon, Tue off). If present Wed-Sat or >= 14 days in month -> Full Month Present.
   - Admission Dept (IDs: 2005, 2006, 6001, 1040, 1017, 2011, 6000, 2010, 2007, 2013, 2514, 2512, 2511, 2503, 2502, 2505, 2051, 2508, 6004, 6005):
       * 6 working days/week (Sundays allowed as working day).
       * 2nd Saturday week has 5 working days.
       * If weekly/monthly quota is met -> Full Month Present.
"""

import re
from datetime import datetime

# Full month exemption records (guaranteed to be present in all generated sheets and views)
GUARANTEED_EXEMPT_EMPLOYEES = [
    {
        "emp_id": "101",
        "emp_name": "Dr. M Mohan Babu",
        "domain": "Teaching",
        "department": "Principal Office",
        "designation": "Principal Sir / Principal",
        "standard_salary": 210089.0,
        "basic": 137035.0,
        "account_no": "4017000402717199",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "707",
        "emp_name": "R. Gunasekaran",
        "domain": "Teaching",
        "department": "IT",
        "designation": "Asst.Prof",
        "standard_salary": 48000.0,
        "basic": 31309.0,
        "account_no": "4017000402728912",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "900",
        "emp_name": "Dr. I. Sudarsan Kumar",
        "domain": "Teaching",
        "department": "HAS",
        "designation": "Professor & DAP",
        "standard_salary": 185000.0,
        "basic": 120671.0,
        "account_no": "4017000402719834",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "1060",
        "emp_name": "Vivekand Adhikari (Veveka)",
        "domain": "Non-Teaching",
        "department": "Administration",
        "designation": "IR Officer",
        "standard_salary": 35000.0,
        "basic": 35000.0,
        "account_no": "4017000402734123",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "1015",
        "emp_name": "R Hari Krishna",
        "domain": "Non-Teaching",
        "department": "Exam Section",
        "designation": "Jr.Asst",
        "standard_salary": 16500.0,
        "basic": 16500.0,
        "account_no": "4017000402723145",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "1019",
        "emp_name": "Bishal Kumar Sha (Vishal Kumar)",
        "domain": "Non-Teaching",
        "department": "TAP",
        "designation": "Executive Assistant",
        "standard_salary": 25000.0,
        "basic": 25000.0,
        "account_no": "4017000402729988",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "1021",
        "emp_name": "M.P.Balaji",
        "domain": "Management",
        "department": "Management / Chairman Office",
        "designation": "Accts. Officer",
        "standard_salary": 43700.0,
        "basic": 43700.0,
        "account_no": "4017000402715566",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "4001",
        "emp_name": "R Poorna Chandra",
        "domain": "Management",
        "department": "Management / Chairman Office",
        "designation": "Estate Manager",
        "standard_salary": 35000.0,
        "basic": 35000.0,
        "account_no": "4017000402716677",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "MGT_1095",
        "emp_name": "S. Shajahan",
        "domain": "Management",
        "department": "Management / Chairman Office",
        "designation": "P.A to Chairman",
        "standard_salary": 24685.0,
        "basic": 24685.0,
        "account_no": "4017000402718899",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "6623",
        "emp_name": "D. Siva (Principal Driver)",
        "domain": "Support Staff",
        "department": "Transport",
        "designation": "Principal Driver",
        "standard_salary": 16500.0,
        "basic": 16500.0,
        "account_no": "4017000402726655",
        "ifsc_code": "PUNB0401700"
    },
    {
        "emp_id": "1030",
        "emp_name": "Thangeeru Surendera (VC Driver)",
        "domain": "Support Staff",
        "department": "Transport",
        "designation": "VC Driver",
        "standard_salary": 30000.0,
        "basic": 30000.0,
        "account_no": "4017000402725544",
        "ifsc_code": "PUNB0401700"
    }
]

# Set of IDs for quick lookup
AUTO_FULL_MONTH_IDS = {g["emp_id"] for g in GUARANTEED_EXEMPT_EMPLOYEES}

# Electricians IDs
ELECTRICIAN_IDS = {'206', '243', '218', '217', '213'}

# Admission Department IDs
ADMISSION_IDS = {
    '2005', '2006', '6001', '1040', '1017', '2011', '6000', '2010', '2007', '2013',
    '2514', '2512', '2511', '2503', '2502', '2505', '2051', '2508', '6004', '6005'
}

# Transport Department IDs
TRANSPORT_IDS = {
    '625', '626', '627', '648', '1198', '628', '622', '6621', '606', '623', '603',
    '653', '605', '6623', '607', '6633', '610', '613', '26', '27'
}

def parse_time_minutes(t_str):
    """Parses 'HH:MM' string to minutes from midnight."""
    if not t_str:
        return None
    s = str(t_str).strip()
    m = re.search(r'(\d{1,2}):(\d{2})', s)
    if m:
        hh = int(m.group(1))
        mm = int(m.group(2))
        return hh * 60 + mm
    return None

def is_second_saturday(dt):
    """Check if date is the 2nd Saturday of the month."""
    if dt.weekday() == 5: # Saturday
        return 8 <= dt.day <= 14
    return False

def is_sunday(dt):
    return dt.weekday() == 6

def evaluate_daily_punch_rule(emp_id, date_str, in_time_str, out_time_str, dept_name="", base_status="Present"):
    """
    Evaluates institutional custom punch rules for a single day.
    Returns: status string ("Present", "Absent", "Half Day", "Holiday", etc.) and credit (1.0, 0.5, 0.0).
    """
    eid = str(emp_id).strip()
    dept = (dept_name or "").upper()
    
    in_mins = parse_time_minutes(in_time_str)
    out_mins = parse_time_minutes(out_time_str)
    has_punch = bool(in_mins is not None or out_mins is not None)
    has_both = bool(in_mins is not None and out_mins is not None)

    # 1. CSE ID 536 (Bala Subramanyam): Punch in before 12:10 PM -> Full day
    if eid == '536':
        if in_mins is not None and in_mins <= (12 * 60 + 10):
            return "Present", 1.0

    # 2. Civil ID 109 (M. Leelakar / Lilaakar): Punch in before 11:00 AM -> Full day
    if eid == '109':
        if in_mins is not None and in_mins <= (11 * 60):
            return "Present", 1.0

    # 3. Media Team ID 1018 (Prudhvi Raj): In-time up to 9:35 AM -> Full day
    if eid == '1018':
        if in_mins is not None and in_mins <= (9 * 60 + 35):
            return "Present", 1.0

    # 4. Electricians (IDs 206, 243, 218, 217, 213):
    #    Normal 9:25 AM. Early shift: In <= 8:30 AM and Out >= 4:30 PM (16:30) -> Full day
    if eid in ELECTRICIAN_IDS or "ELECTRI" in dept:
        if in_mins is not None:
            if in_mins <= (8 * 60 + 30) and out_mins is not None and out_mins >= (16 * 60 + 30):
                return "Present", 1.0
            if in_mins <= (9 * 60 + 25):
                return "Present", 1.0

    # 5. Attenders: In <= 8:35 AM, Out >= 5:30 PM (17:30)
    if "ATTENDER" in dept:
        if in_mins is not None and in_mins <= (8 * 60 + 35) and out_mins is not None and out_mins >= (17 * 60 + 30):
            return "Present", 1.0

    # 6. Garden Staff: In <= 8:35 AM, Out >= 5:10 PM (17:10)
    if "GARDEN" in dept:
        if in_mins is not None and in_mins <= (8 * 60 + 35) and out_mins is not None and out_mins >= (17 * 60 + 10):
            return "Present", 1.0

    # 7. Transport Staff: Any valid IN and OUT punch with no time limit -> Full day
    if eid in TRANSPORT_IDS or "TRANSPORT" in dept or "DRIVER" in dept or "CLEANER" in dept:
        if has_both:
            return "Present", 1.0

    # Default fallback to base status
    st_lower = str(base_status).lower()
    if 'present' in st_lower and 'absent' not in st_lower:
        return "Present", 1.0
    elif 'half' in st_lower:
        return "Half Day", 0.5
    elif 'holiday' in st_lower or 'week' in st_lower:
        return "Holiday", 1.0
    elif 'leave' in st_lower or 'cl' in st_lower:
        return "Leave", 1.0
    elif 'od' in st_lower or 'duty' in st_lower:
        return "On Duty", 1.0
    
    return "Absent", 0.0

def apply_institutional_employee_overrides(emp_record, month_days=31):
    """
    Applies special monthly schedule rules, quota checks, and exemption overrides to an employee record.
    Modifies and returns the updated emp_record dictionary.
    """
    eid = str(emp_record.get('emp_id', '')).strip()
    name = (emp_record.get('emp_name', '') or '').lower()
    dept = (emp_record.get('department', '') or '').upper()
    domain = emp_record.get('domain', '')
    
    bio_days = float(emp_record.get('biometric_days', 0.0) or 0.0)
    hol_days = float(emp_record.get('holiday_days', 0.0) or 0.0)
    leaves_days = float(emp_record.get('availed_leaves', 0.0) or 0.0)
    od_days = float(emp_record.get('od_days', 0.0) or 0.0)
    curr_pay_days = float(emp_record.get('total_pay_days', 0.0) or (bio_days + hol_days + leaves_days + od_days))
    
    default_holidays = hol_days if hol_days > 0 else 4.0

    # =========================================================================
    # RULE 1: 100% Full-Month Present Exemptions
    # =========================================================================
    is_auto_full = False
    
    if eid in AUTO_FULL_MONTH_IDS:
        is_auto_full = True
    elif 'principal' in name or 'principal' in dept.lower():
        is_auto_full = True
    elif 'shajahan' in name:
        is_auto_full = True
    elif ('siva' in name or 'driver' in name) and ('principal' in dept.lower() or 'vc' in dept.lower() or 'management' in dept.lower()):
        is_auto_full = True
    elif eid == '1060' or 'viveka' in name or 'veveka' in name:
        is_auto_full = True
    elif eid == '1015' or 'hari krishna' in name:
        is_auto_full = True
    elif eid == '1019' or 'vishal' in name or 'bishal' in name:
        is_auto_full = True
    elif eid == '707' or 'gunasekaran' in name:
        is_auto_full = True
    elif eid == '900' or 'sudarsan' in name:
        is_auto_full = True

    if is_auto_full:
        emp_record['holiday_days'] = default_holidays
        emp_record['biometric_days'] = max(bio_days, float(month_days) - default_holidays)
        emp_record['total_pay_days'] = float(month_days)
        emp_record['remarks'] = '[Exempt / Full Month Present]'
        return emp_record

    # =========================================================================
    # RULE 2: IT ID 1053 (Pachayapan) - 3 days a week (Thu, Fri, Sat) / >= 12 days
    # =========================================================================
    if eid == '1053' or 'pachaiyappan' in name or 'pachayapan' in name:
        if bio_days >= 12.0 or curr_pay_days >= 12.0:
            emp_record['holiday_days'] = default_holidays
            emp_record['biometric_days'] = float(month_days) - default_holidays
            emp_record['total_pay_days'] = float(month_days)
            emp_record['remarks'] = 'Pachayapan 3-Day Shift Met -> Full Month'
            return emp_record

    # =========================================================================
    # RULE 3: IT HOD ID 1203 (Dr J Velmurugan) - Wed to Sat / >= 14 days
    # =========================================================================
    if eid == '1203' or 'velmurugan' in name:
        if bio_days >= 14.0 or curr_pay_days >= 14.0:
            emp_record['holiday_days'] = default_holidays
            emp_record['biometric_days'] = float(month_days) - default_holidays
            emp_record['total_pay_days'] = float(month_days)
            emp_record['remarks'] = 'IT HOD 4-Day Shift Met -> Full Month'
            return emp_record

    # =========================================================================
    # RULE 4: Admission Dept Staff (6 days/week, 2nd Sat week 5 days)
    # =========================================================================
    if eid in ADMISSION_IDS or 'ADMISSION' in dept or domain == 'Admission':
        if (bio_days + leaves_days + od_days) >= 22.0 or curr_pay_days >= 26.0:
            emp_record['holiday_days'] = default_holidays
            emp_record['biometric_days'] = float(month_days) - default_holidays
            emp_record['total_pay_days'] = float(month_days)
            emp_record['remarks'] = 'Admission 6-Day Quota Met -> Full Month'
            return emp_record

    # =========================================================================
    # RULE 5: Transport Staff (Any in & out punch)
    # =========================================================================
    if eid in TRANSPORT_IDS or 'TRANSPORT' in dept:
        emp_record['total_pay_days'] = min(float(month_days), bio_days + hol_days + leaves_days + od_days)
        return emp_record

    # Default calculation
    emp_record['total_pay_days'] = min(float(month_days), bio_days + hol_days + leaves_days + od_days)
    return emp_record

def inject_guaranteed_exempt_employees(attendance_map, month_days=31):
    """
    Guarantees that all 11 exempt employees are always present in the attendance map
    with full month present (31/31), even if they are completely missing from the uploaded file!
    """
    hol = 4.0
    if attendance_map:
        sample_hols = [v.get("holiday_days", 0.0) for v in attendance_map.values() if v.get("holiday_days", 0.0) > 0]
        if sample_hols:
            hol = max(sample_hols)

    bio_days = max(0.0, float(month_days) - hol)

    for g in GUARANTEED_EXEMPT_EMPLOYEES:
        eid = g["emp_id"]
        if eid not in attendance_map:
            attendance_map[eid] = {
                "emp_id": eid,
                "emp_name": g["emp_name"],
                "department": g["department"],
                "domain": g["domain"],
                "designation": g["designation"],
                "biometric_days": bio_days,
                "holiday_days": hol,
                "availed_leaves": 0.0,
                "od_days": 0.0,
                "total_pay_days": float(month_days),
                "daily_logs": [],
                "remarks": "[Exempt / Full Month Present]"
            }
        else:
            attendance_map[eid]["holiday_days"] = hol
            attendance_map[eid]["biometric_days"] = max(attendance_map[eid].get("biometric_days", 0.0), bio_days)
            attendance_map[eid]["total_pay_days"] = float(month_days)
            attendance_map[eid]["remarks"] = "[Exempt / Full Month Present]"

    return attendance_map
