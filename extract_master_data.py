import openpyxl
import xlrd
import pandas as pd
import json
import re

def clean_str(val):
    if val is None:
        return ""
    return str(val).strip()

def clean_num(val):
    if val is None or val == "":
        return 0.0
    try:
        if isinstance(val, str):
            val = val.replace(",", "").strip()
        return float(val)
    except:
        return 0.0

def extract_teaching_staff():
    print("--- Extracting Teaching Staff ---")
    wb = openpyxl.load_workbook('SVCET July2026 - Teaching Staff (2).xlsx', data_only=True)
    sh = wb['Teaching']
    
    current_dept = "General"
    records = []
    
    # Let's inspect rows
    for r in range(4, sh.max_row + 1):
        c1 = sh.cell(r, 1).value # Sl No or None
        c2 = sh.cell(r, 2).value # Name or Dept
        c3 = sh.cell(r, 3).value # Designation
        c4 = sh.cell(r, 4).value # Total Salary
        c5 = sh.cell(r, 5).value # Consolidated
        c6 = sh.cell(r, 6).value # Basic
        c7 = sh.cell(r, 7).value # AGP
        c8 = sh.cell(r, 8).value # Total
        c9 = sh.cell(r, 9).value # No of Days
        c10 = sh.cell(r, 10).value # Basic+AGP
        c11 = sh.cell(r, 11).value # DA
        c12 = sh.cell(r, 12).value # HRA
        c13 = sh.cell(r, 13).value # Arrears
        c14 = sh.cell(r, 14).value # FA
        c15 = sh.cell(r, 15).value # TA/SA
        c16 = sh.cell(r, 16).value # Gross Total
        c17 = sh.cell(r, 17).value # EPF
        c18 = sh.cell(r, 18).value # IT
        c19 = sh.cell(r, 19).value # PT
        c20 = sh.cell(r, 20).value # WF
        c21 = sh.cell(r, 21).value # EB
        c22 = sh.cell(r, 22).value # Mess
        c23 = sh.cell(r, 23).value # Bus/others
        c24 = sh.cell(r, 24).value # Tot Ded
        c25 = sh.cell(r, 25).value # Net Salary
        c26 = sh.cell(r, 26).value # PNB Account Number
        c27 = sh.cell(r, 27).value # IFSC Code
        
        name_str = clean_str(c2)
        if not name_str:
            continue
        
        if "TEACHING STAFF" in name_str or "Total" in name_str or "Grand Total" in name_str:
            continue
            
        if "Dept" in name_str or "DEPARTMENT" in name_str.upper() or "CIVIL" in name_str.upper() or "EEE" in name_str.upper() or "MECH" in name_str.upper() or "ECE" in name_str.upper() or "CSE" in name_str.upper() or "CSM" in name_str.upper() or "CSD" in name_str.upper() or "CAI" in name_str.upper() or "MBA" in name_str.upper() or "MCA" in name_str.upper() or "H&S" in name_str.upper() or "HAS" in name_str.upper():
            if c3 is None or clean_str(c3) == "":
                current_dept = name_str.replace("Dept. of", "").replace("Dept of", "").strip()
                continue
                
        if c3 is not None and clean_str(c3) != "":
            records.append({
                "sl_no": c1,
                "name": name_str,
                "department": current_dept,
                "designation": clean_str(c3),
                "total_salary": clean_num(c4),
                "consolidated": clean_num(c5),
                "basic": clean_num(c6),
                "agp": clean_num(c7),
                "month_days": clean_num(c9),
                "da": clean_num(c11),
                "hra": clean_num(c12),
                "arrears": clean_num(c13),
                "fa": clean_num(c14),
                "ta_sa": clean_num(c15),
                "gross": clean_num(c16),
                "epf": clean_num(c17),
                "it": clean_num(c18),
                "pt": clean_num(c19),
                "wf": clean_num(c20),
                "eb": clean_num(c21),
                "mess": clean_num(c22),
                "bus": clean_num(c23),
                "total_ded": clean_num(c24),
                "net_salary": clean_num(c25),
                "account_no": clean_str(c26),
                "ifsc": clean_str(c27),
                "category": "Teaching"
            })

    print(f"Extracted {len(records)} teaching staff records.")
    return records

def extract_id_map():
    # Map Employee Names to Emp. Code from 'SVCET July 2026 - No. of Days.xls' and 'SVCET July2026 - Teaching Staff (ID) (1).xlsx'
    id_map = {}
    
    # 1. From No. of Days
    try:
        wb = xlrd.open_workbook('SVCET July 2026 - No. of Days.xls')
        sh = wb.sheet_by_name('New')
        for r in range(sh.nrows):
            code = clean_str(sh.cell_value(r, 1))
            name = clean_str(sh.cell_value(r, 2))
            desig = clean_str(sh.cell_value(r, 3))
            dept = ""
            if code and code not in ["Emp. Code", "Department:", "S.No"]:
                try:
                    if float(code).is_integer():
                        code = str(int(float(code)))
                except:
                    pass
                if name:
                    id_map[name.lower().replace(" ", "").replace(".", "")] = {
                        "emp_code": code,
                        "raw_name": name,
                        "designation": desig
                    }
    except Exception as e:
        print("Error in No. of Days map:", e)
        
    print(f"Mapped {len(id_map)} employees from No. of Days report.")
    return id_map

if __name__ == "__main__":
    t_staff = extract_teaching_staff()
    id_map = extract_id_map()
    
    # Match teaching staff with emp codes
    matched = 0
    unmatched = []
    for s in t_staff:
        key = s['name'].lower().replace(" ", "").replace(".", "")
        if key in id_map:
            s['emp_code'] = id_map[key]['emp_code']
            matched += 1
        else:
            # try partial match
            found = False
            for k, v in id_map.items():
                if key in k or k in key:
                    s['emp_code'] = v['emp_code']
                    matched += 1
                    found = True
                    break
            if not found:
                s['emp_code'] = ""
                unmatched.append(s['name'])
                
    print(f"Matched {matched}/{len(t_staff)} teaching staff with Employee Codes.")
    if unmatched:
        print("Sample unmatched names:", unmatched[:10])
