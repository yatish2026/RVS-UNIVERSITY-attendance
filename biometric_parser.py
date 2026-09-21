import xlrd
import openpyxl
import os
import re
from institutional_rules import evaluate_daily_punch_rule, apply_institutional_employee_overrides, inject_guaranteed_exempt_employees

def clean_str(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith('.0') and s[:-2].isdigit():
        return s[:-2]
    return s

def clean_num(val):
    if val is None or val == "":
        return 0.0
    try:
        if isinstance(val, str):
            val = val.replace(",", "").strip()
        return float(val)
    except:
        return 0.0

class BiometricParser:
    def __init__(self, file_path):
        self.file_path = file_path

    def parse(self, month_days=31):
        if self.file_path.endswith('.xlsx'):
            att_map = self._parse_xlsx(month_days)
        else:
            att_map = self._parse_xls(month_days)
        # Always inject guaranteed exempt staff even if completely missing from the sheet
        return inject_guaranteed_exempt_employees(att_map, month_days)

    def _determine_domain(self, dept_name):
        d = (dept_name or "").upper()
        if any(k in d for k in ["CE", "CIVIL", "EEE", "ME", "MECH", "ECE", "CSE", "CSM", "CSD", "CAI", "MBA", "MCA", "HAS", "H&S", "TEACHING"]):
            return "Teaching"
        elif any(k in d for k in ["SECURITY", "TRANSPORT", "ATTENDER", "GARDEN", "HOUSEKEEPING"]):
            return "Support Staff"
        elif "SBF" in d:
            return "SBF Facility"
        elif "SLH" in d or "MESS" in d:
            return "Hostel & Mess"
        elif "ADMISSION" in d:
            return "Admission"
        elif "MANAGEMENT" in d:
            return "Management"
        return "Non-Teaching"

    def _parse_xls(self, month_days=31):
        wb = xlrd.open_workbook(self.file_path)
        
        # Check if this workbook is a Summary Report (e.g. "SVCET July 2026 - No. of Days.xls")
        # or a Raw Biometric Machine Log ("Daily Attendance Report 01-09-2026.xls")
        is_summary_file = False
        for sname in wb.sheet_names():
            sh = wb.sheet_by_name(sname)
            for r in range(min(10, sh.nrows)):
                row_str = " ".join([clean_str(sh.cell_value(r, c)).lower() for c in range(sh.ncols)])
                if "biometric days" in row_str or "availed leaves" in row_str or "total pay days" in row_str or "working days" in row_str:
                    is_summary_file = True
                    break
            if is_summary_file:
                break

        if is_summary_file:
            return self._parse_summary_xls(wb, month_days)
        else:
            return self._parse_raw_machine_xls(wb, month_days)

    def _parse_summary_xls(self, wb, month_days=31):
        employees = {}
        for sname in wb.sheet_names():
            sh = wb.sheet_by_name(sname)
            current_dept = "General"
            
            col_map = {
                "emp_id": -1, "name": -1, "desig": -1, "bio": -1,
                "holiday": -1, "leaves": -1, "od": -1, "total_pay": -1, "remarks": -1
            }

            for r in range(sh.nrows):
                row_vals = [clean_str(sh.cell_value(r, c)) for c in range(sh.ncols)]
                row_lower = [v.lower() for v in row_vals]

                # Check department row
                if "department:" in row_lower:
                    idx = row_lower.index("department:")
                    if idx + 1 < len(row_vals) and row_vals[idx+1]:
                        current_dept = row_vals[idx+1]
                    continue
                
                # Check for "Department: XYZ" in any cell
                for cell in row_vals:
                    if cell.startswith("Department:"):
                        current_dept = cell.replace("Department:", "").strip()

                # Detect header row
                if any("biometric days" in v or "biometric present" in v or "working days" in v for v in row_lower):
                    for c_idx, v in enumerate(row_lower):
                        if any(k in v for k in ["emp. code", "emp code", "biometric no", "emp id", "s.no"]) and col_map["emp_id"] == -1:
                            col_map["emp_id"] = c_idx
                        elif any(k in v for k in ["employeename", "name of the staff", "name of the sraff", "staff name", "name"]):
                            col_map["name"] = c_idx
                        elif "designation" in v:
                            col_map["desig"] = c_idx
                        elif any(k in v for k in ["biometric days", "biometric present", "working days"]):
                            col_map["bio"] = c_idx
                        elif any(k in v for k in ["holiday", "week off"]):
                            col_map["holiday"] = c_idx
                        elif any(k in v for k in ["availed leaves", "leaves", "cl"]):
                            col_map["leaves"] = c_idx
                        elif any(k in v for k in ["sv/od", "od", "on duty"]):
                            col_map["od"] = c_idx
                        elif any(k in v for k in ["total pay days", "total days", "pay days"]):
                            col_map["total_pay"] = c_idx
                        elif "remarks" in v:
                            col_map["remarks"] = c_idx
                    continue

                # Process employee row
                if col_map["bio"] != -1:
                    raw_id = row_vals[col_map["emp_id"]] if col_map["emp_id"] != -1 and col_map["emp_id"] < len(row_vals) else ""
                    if not raw_id or raw_id.lower() in ["emp. code", "s.no", "total", "grand total"]:
                        continue

                    emp_code = raw_id
                    emp_name = row_vals[col_map["name"]] if col_map["name"] != -1 and col_map["name"] < len(row_vals) else ""
                    desig = row_vals[col_map["desig"]] if col_map["desig"] != -1 and col_map["desig"] < len(row_vals) else ""
                    
                    bio_days = clean_num(sh.cell_value(r, col_map["bio"])) if col_map["bio"] != -1 and col_map["bio"] < sh.ncols else 0.0
                    hol_days = clean_num(sh.cell_value(r, col_map["holiday"])) if col_map["holiday"] != -1 and col_map["holiday"] < sh.ncols else 0.0
                    leaves_days = clean_num(sh.cell_value(r, col_map["leaves"])) if col_map["leaves"] != -1 and col_map["leaves"] < sh.ncols else 0.0
                    od_days = clean_num(sh.cell_value(r, col_map["od"])) if col_map["od"] != -1 and col_map["od"] < sh.ncols else 0.0
                    
                    if col_map["total_pay"] != -1 and col_map["total_pay"] < sh.ncols:
                        pay_days = clean_num(sh.cell_value(r, col_map["total_pay"]))
                        if pay_days <= 0:
                            pay_days = min(float(month_days), bio_days + hol_days + leaves_days + od_days)
                    else:
                        pay_days = min(float(month_days), bio_days + hol_days + leaves_days + od_days)

                    remarks = row_vals[col_map["remarks"]] if col_map["remarks"] != -1 and col_map["remarks"] < len(row_vals) else ""
                    domain = self._determine_domain(current_dept)

                    emp_dict = {
                        "emp_id": emp_code,
                        "emp_name": emp_name,
                        "designation": desig,
                        "department": current_dept,
                        "domain": domain,
                        "biometric_days": bio_days,
                        "holiday_days": hol_days,     # EXACT count from sheet
                        "availed_leaves": leaves_days, # EXACT count from sheet
                        "od_days": od_days,           # EXACT count from sheet
                        "total_pay_days": pay_days,   # EXACT count from sheet
                        "daily_logs": [],
                        "remarks": remarks
                    }

                    # Apply institutional overrides
                    emp_dict = apply_institutional_employee_overrides(emp_dict, month_days)
                    employees[emp_code] = emp_dict

        return employees

    def _parse_raw_machine_xls(self, wb, month_days=31):
        employees = {}
        for sname in wb.sheet_names():
            sh = wb.sheet_by_name(sname)
            current_dept = "General"
            r = 0
            while r < sh.nrows:
                c1 = clean_str(sh.cell_value(r, 1))
                c3 = clean_str(sh.cell_value(r, 3))
                c7 = clean_str(sh.cell_value(r, 7))

                if 'Department:' in c1:
                    current_dept = c3

                if 'Employee Code:' in c1:
                    emp_code = c3
                    emp_name = c7
                    daily_logs = []
                    present_count = 0.0
                    holiday_count = 0.0
                    leave_count = 0.0
                    od_count = 0.0
                    absent_count = 0.0
                    found_summary = False

                    r += 1
                    while r < sh.nrows:
                        row_c1 = clean_str(sh.cell_value(r, 1))
                        if 'Employee Code:' in row_c1:
                            r -= 1
                            break
                        if 'Department:' in row_c1:
                            current_dept = clean_str(sh.cell_value(r, 3))
                            r += 1
                            continue

                        # Check for total duration summary row
                        if 'Total Duration=' in row_c1:
                            found_summary = True
                            m_pres = re.search(r'PresentDays=([\d\.]+)', row_c1)
                            m_leaves = re.search(r'Leaves=([\d\.]+)', row_c1)
                            m_hol = re.search(r'Holiday=([\d\.]+)', row_c1)
                            m_abs = re.search(r'AbsentDays=([\d\.]+)', row_c1)
                            
                            if m_pres: present_count = float(m_pres.group(1))
                            if m_leaves: leave_count = float(m_leaves.group(1))
                            if m_hol: holiday_count = float(m_hol.group(1)) # EXACT count from raw sheet
                            if m_abs: absent_count = float(m_abs.group(1))
                            r += 1
                            break

                        date_val = clean_str(sh.cell_value(r, 1))
                        if date_val and date_val not in ['Date', '']:
                            in_time = clean_str(sh.cell_value(r, 3))
                            out_time = clean_str(sh.cell_value(r, 4))
                            shift = clean_str(sh.cell_value(r, 6))
                            duration = clean_str(sh.cell_value(r, 7))
                            status = clean_str(sh.cell_value(r, 8))
                            remarks = clean_str(sh.cell_value(r, 10))

                            # Evaluate institutional custom punch rules for this day
                            adj_status, adj_credit = evaluate_daily_punch_rule(
                                emp_code, date_val, in_time, out_time, current_dept, status
                            )

                            daily_logs.append({
                                "date": date_val,
                                "in_time": in_time,
                                "out_time": out_time,
                                "shift": shift,
                                "duration": duration,
                                "status": adj_status,
                                "credit": adj_credit,
                                "remarks": remarks
                            })

                        r += 1

                    # Count accurately from evaluated daily logs if available
                    if daily_logs:
                        p_c = 0.0
                        h_c = 0.0
                        l_c = 0.0
                        o_c = 0.0
                        for log in daily_logs:
                            st = log['status'].lower()
                            cr = log.get('credit', 1.0)
                            if 'present' in st and 'absent' not in st:
                                p_c += cr
                            elif 'half' in st:
                                p_c += 0.5
                            elif 'leave' in st or 'cl' in st:
                                l_c += 1.0
                            elif 'holiday' in st or 'weekly off' in st or 'week' in st:
                                h_c += 1.0
                            elif 'od' in st or 'duty' in st:
                                o_c += 1.0
                        
                        # Use greater of machine summary or institutional evaluated punches
                        present_count = max(present_count, p_c)
                        if holiday_count == 0.0: holiday_count = h_c
                        if leave_count == 0.0: leave_count = l_c
                        if od_count == 0.0: od_count = o_c

                    total_pay_days = min(float(month_days), present_count + holiday_count + leave_count + od_count)
                    domain = self._determine_domain(current_dept)

                    if emp_code:
                        emp_dict = {
                            "emp_id": emp_code,
                            "emp_name": emp_name,
                            "department": current_dept,
                            "domain": domain,
                            "biometric_days": present_count,
                            "holiday_days": holiday_count,     # EXACT count from sheet
                            "availed_leaves": leave_count,     # EXACT count from sheet
                            "od_days": od_count,               # EXACT count from sheet
                            "total_pay_days": total_pay_days,  # EXACT count from sheet
                            "daily_logs": daily_logs,
                            "remarks": ""
                        }
                        emp_dict = apply_institutional_employee_overrides(emp_dict, month_days)
                        employees[emp_code] = emp_dict
                r += 1

        return employees

    def _parse_xlsx(self, month_days=31):
        wb = openpyxl.load_workbook(self.file_path, data_only=True)
        
        # Check if summary or raw
        is_summary = False
        for sname in wb.sheetnames:
            sh = wb[sname]
            for r in range(1, min(10, sh.max_row+1)):
                row_str = " ".join([clean_str(sh.cell(r, c).value).lower() for c in range(1, sh.max_column+1)])
                if "biometric days" in row_str or "availed leaves" in row_str or "total pay days" in row_str or "working days" in row_str:
                    is_summary = True
                    break
            if is_summary:
                break

        if is_summary:
            return self._parse_summary_xlsx(wb, month_days)
        else:
            return self._parse_raw_xlsx(wb, month_days)

    def _parse_summary_xlsx(self, wb, month_days=31):
        employees = {}
        for sname in wb.sheetnames:
            sh = wb[sname]
            current_dept = "General"
            col_map = {
                "emp_id": -1, "name": -1, "desig": -1, "bio": -1,
                "holiday": -1, "leaves": -1, "od": -1, "total_pay": -1, "remarks": -1
            }

            for r in range(1, sh.max_row + 1):
                row_vals = [clean_str(sh.cell(r, c).value) for c in range(1, sh.max_column + 1)]
                row_lower = [v.lower() for v in row_vals]

                if "department:" in row_lower:
                    idx = row_lower.index("department:")
                    if idx + 1 < len(row_vals) and row_vals[idx+1]:
                        current_dept = row_vals[idx+1]
                    continue

                for cell in row_vals:
                    if cell.startswith("Department:"):
                        current_dept = cell.replace("Department:", "").strip()

                if any("biometric days" in v or "biometric present" in v or "working days" in v for v in row_lower):
                    for c_idx, v in enumerate(row_lower):
                        col_1_based = c_idx + 1
                        if any(k in v for k in ["emp. code", "emp code", "biometric no", "emp id", "sl. no", "s.no"]) and col_map["emp_id"] == -1:
                            col_map["emp_id"] = col_1_based
                        elif any(k in v for k in ["employeename", "name of the staff", "name of the sraff", "staff name", "name"]):
                            col_map["name"] = col_1_based
                        elif "designation" in v:
                            col_map["desig"] = col_1_based
                        elif any(k in v for k in ["biometric days", "biometric present", "working days"]):
                            col_map["bio"] = col_1_based
                        elif any(k in v for k in ["holiday", "week off"]):
                            col_map["holiday"] = col_1_based
                        elif any(k in v for k in ["availed leaves", "leaves", "cl"]):
                            col_map["leaves"] = col_1_based
                        elif any(k in v for k in ["sv/od", "od", "on duty"]):
                            col_map["od"] = col_1_based
                        elif any(k in v for k in ["total pay days", "total days", "pay days"]):
                            col_map["total_pay"] = col_1_based
                        elif "remarks" in v:
                            col_map["remarks"] = col_1_based
                    continue

                if col_map["bio"] != -1:
                    raw_id = clean_str(sh.cell(r, col_map["emp_id"]).value) if col_map["emp_id"] != -1 else ""
                    if not raw_id or raw_id.lower() in ["emp. code", "s.no", "total", "grand total"]:
                        continue

                    emp_code = raw_id
                    emp_name = clean_str(sh.cell(r, col_map["name"]).value) if col_map["name"] != -1 else ""
                    desig = clean_str(sh.cell(r, col_map["desig"]).value) if col_map["desig"] != -1 else ""
                    
                    bio_days = clean_num(sh.cell(r, col_map["bio"]).value) if col_map["bio"] != -1 else 0.0
                    hol_days = clean_num(sh.cell(r, col_map["holiday"]).value) if col_map["holiday"] != -1 else 0.0
                    leaves_days = clean_num(sh.cell_value(r, col_map["leaves"])) if col_map["leaves"] != -1 and col_map["leaves"] < sh.ncols else (clean_num(sh.cell(r, col_map["leaves"]).value) if col_map["leaves"] != -1 else 0.0)
                    od_days = clean_num(sh.cell(r, col_map["od"]).value) if col_map["od"] != -1 else 0.0
                    
                    if col_map["total_pay"] != -1:
                        pay_days = clean_num(sh.cell(r, col_map["total_pay"]).value)
                        if pay_days <= 0:
                            pay_days = min(float(month_days), bio_days + hol_days + leaves_days + od_days)
                    else:
                        pay_days = min(float(month_days), bio_days + hol_days + leaves_days + od_days)

                    remarks = clean_str(sh.cell(r, col_map["remarks"]).value) if col_map["remarks"] != -1 else ""
                    domain = self._determine_domain(current_dept)

                    emp_dict = {
                        "emp_id": emp_code,
                        "emp_name": emp_name,
                        "designation": desig,
                        "department": current_dept,
                        "domain": domain,
                        "biometric_days": bio_days,
                        "holiday_days": hol_days,     # EXACT count from sheet
                        "availed_leaves": leaves_days, # EXACT count from sheet
                        "od_days": od_days,           # EXACT count from sheet
                        "total_pay_days": pay_days,   # EXACT count from sheet
                        "daily_logs": [],
                        "remarks": remarks
                    }
                    emp_dict = apply_institutional_employee_overrides(emp_dict, month_days)
                    employees[emp_code] = emp_dict

        return employees

    def _parse_raw_xlsx(self, wb, month_days=31):
        employees = {}
        for sname in wb.sheetnames:
            sh = wb[sname]
            current_dept = "General"
            r = 1
            while r <= sh.max_row:
                c1 = clean_str(sh.cell(r, 2).value)
                c3 = clean_str(sh.cell(r, 4).value)
                c7 = clean_str(sh.cell(r, 8).value)

                if 'Department:' in c1:
                    current_dept = c3

                if 'Employee Code:' in c1:
                    emp_code = c3
                    emp_name = c7
                    present_count = 0.0
                    leave_count = 0.0
                    holiday_count = 0.0
                    od_count = 0.0
                    daily_logs = []

                    r += 1
                    while r <= sh.max_row:
                        row_c1 = clean_str(sh.cell(r, 2).value)
                        if 'Employee Code:' in row_c1:
                            r -= 1
                            break
                        if 'Total Duration=' in row_c1:
                            m_pres = re.search(r'PresentDays=([\d\.]+)', row_c1)
                            m_leaves = re.search(r'Leaves=([\d\.]+)', row_c1)
                            m_hol = re.search(r'Holiday=([\d\.]+)', row_c1)
                            if m_pres: present_count = float(m_pres.group(1))
                            if m_leaves: leave_count = float(m_leaves.group(1))
                            if m_hol: holiday_count = float(m_hol.group(1)) # EXACT count from sheet
                            r += 1
                            break
                        
                        date_val = clean_str(sh.cell(r, 2).value)
                        if date_val and date_val not in ['Date', '']:
                            in_time = clean_str(sh.cell(r, 4).value)
                            out_time = clean_str(sh.cell(r, 5).value)
                            status = clean_str(sh.cell(r, 9).value)
                            adj_status, adj_credit = evaluate_daily_punch_rule(
                                emp_code, date_val, in_time, out_time, current_dept, status
                            )
                            daily_logs.append({
                                "date": date_val,
                                "in_time": in_time,
                                "out_time": out_time,
                                "status": adj_status,
                                "credit": adj_credit
                            })
                        r += 1

                    if daily_logs:
                        p_c = 0.0
                        h_c = 0.0
                        l_c = 0.0
                        o_c = 0.0
                        for log in daily_logs:
                            st = log['status'].lower()
                            cr = log.get('credit', 1.0)
                            if 'present' in st and 'absent' not in st:
                                p_c += cr
                            elif 'half' in st:
                                p_c += 0.5
                            elif 'leave' in st or 'cl' in st:
                                l_c += 1.0
                            elif 'holiday' in st or 'weekly off' in st or 'week' in st:
                                h_c += 1.0
                            elif 'od' in st or 'duty' in st:
                                o_c += 1.0
                        present_count = max(present_count, p_c)
                        if holiday_count == 0.0: holiday_count = h_c
                        if leave_count == 0.0: leave_count = l_c
                        if od_count == 0.0: od_count = o_c

                    total_pay_days = min(float(month_days), present_count + holiday_count + leave_count + od_count)
                    domain = self._determine_domain(current_dept)

                    if emp_code:
                        emp_dict = {
                            "emp_id": emp_code,
                            "emp_name": emp_name,
                            "department": current_dept,
                            "domain": domain,
                            "biometric_days": present_count,
                            "holiday_days": holiday_count,     # EXACT count from sheet
                            "availed_leaves": leave_count,     # EXACT count from sheet
                            "od_days": od_count,               # EXACT count from sheet
                            "total_pay_days": total_pay_days,  # EXACT count from sheet
                            "daily_logs": daily_logs,
                            "remarks": ""
                        }
                        emp_dict = apply_institutional_employee_overrides(emp_dict, month_days)
                        employees[emp_code] = emp_dict
                r += 1
        return employees

if __name__ == "__main__":
    print("Testing Universal BiometricParser with Institutional Rules:")
    
    # Test on Summary Sheet
    p1 = BiometricParser("SVCET July 2026 - No. of Days.xls")
    r1 = p1.parse(31)
    print(f"Summary Sheet parsed: {len(r1)} employees")
    for sid in ['101', '707', '109', '1053', '1203', '900', '1060', '1018', '1015', '1019', '1021', '4001', '603', '626']:
        if sid in r1:
            e = r1[sid]
            print(f"  Emp {sid} ({e['emp_name']} - {e['department']}): Bio={e['biometric_days']}, Hol={e['holiday_days']}, PayDays={e['total_pay_days']} | Remarks={e['remarks']}")

    # Test on Raw Machine Sheet
    p2 = BiometricParser("Daily Attendance Report 01-09-2026 (2).xls")
    r2 = p2.parse(31)
    print(f"\nRaw Machine Sheet parsed: {len(r2)} employees")
    for sid in ['536', '109', '1018', '206', '243', '2005', '625', '626', '627']:
        if sid in r2:
            e = r2[sid]
            print(f"  Emp {sid} ({e['emp_name']} - {e['department']}): Bio={e['biometric_days']}, Hol={e['holiday_days']}, PayDays={e['total_pay_days']} | Remarks={e['remarks']}")
