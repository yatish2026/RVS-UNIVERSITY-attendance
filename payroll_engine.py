import math
import sqlite3
from institutional_rules import apply_institutional_employee_overrides, inject_guaranteed_exempt_employees
from supabase_sync import SupabaseSync

def calc_pt(gross_salary):
    """
    Andhra Pradesh Professional Tax Slabs:
    Gross > 20,000 -> Rs 200
    15,000 < Gross <= 20,000 -> Rs 150
    Gross <= 15,000 -> Rs 0
    """
    if gross_salary > 20000:
        return 200.0
    elif gross_salary > 15000:
        return 150.0
    return 0.0

class PayrollEngine:
    def __init__(self, db_path="payroll_master.db", supabase_client=None):
        self.db_path = db_path
        self.supabase_client = supabase_client or SupabaseSync(db_path)

    def get_employees_map(self):
        if self.supabase_client and self.supabase_client.is_configured():
            return self.supabase_client.get_employees_map()
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM employees")
        rows = cur.fetchall()
        conn.close()
        return {str(r["emp_id"]): dict(r) for r in rows}

    def compute_employee_salary(self, emp, attendance_days, month_days=31, deductions=None):
        if deductions is None:
            deductions = {}

        standard_salary = float(emp.get("standard_salary", 0.0) or 0.0)
        consolidated_salary = float(emp.get("consolidated_salary", 0.0) or 0.0)
        agp = float(emp.get("agp", 0.0) or 0.0)
        fa = float(emp.get("fa", 0.0) or 0.0)
        ta_sa = float(emp.get("ta_sa", 0.0) or 0.0)
        epf_fixed = float(emp.get("epf_fixed", 0.0) or 0.0)
        it_fixed = float(emp.get("it_fixed", 0.0) or 0.0)
        domain = emp.get("domain", "Teaching")
        category = emp.get("category", domain)
        department = emp.get("department", "General")
        salary_type = emp.get("salary_type", "REGULAR")

        # Apply institutional overrides to attendance days
        att_dict = dict(attendance_days)
        att_dict['emp_id'] = emp.get('emp_id')
        att_dict['emp_name'] = emp.get('emp_name')
        att_dict['department'] = department
        att_dict['domain'] = domain
        att_dict = apply_institutional_employee_overrides(att_dict, month_days)

        # Attendance Days
        total_pay_days = float(att_dict.get("total_pay_days", month_days))
        total_pay_days = max(0.0, min(float(month_days), total_pay_days))

        # Basic calculation
        basic_full = float(emp.get("basic", 0.0) or 0.0)
        if basic_full <= 0 and standard_salary > 0:
            if domain == "Teaching":
                basic_full = round(standard_salary / 1.5331)
            else:
                basic_full = standard_salary

        # Prorate Basic + AGP
        if month_days > 0 and standard_salary > 0:
            if domain == "Teaching":
                basic_agp = round((basic_full / month_days) * total_pay_days, 2)
            else:
                basic_agp = round((standard_salary / month_days) * total_pay_days, 2)
        else:
            basic_agp = 0.0

        # Allowances (DA 37.31%, HRA 16% for teaching faculty)
        if domain == "Teaching" and standard_salary > 0:
            da = round(basic_agp * 0.3731)
            hra = round(basic_agp * 0.1600)
        else:
            da = 0.0
            hra = 0.0

        arrears = float(deductions.get("arrears", 0.0) or 0.0)
        gross_total = math.ceil(basic_agp + da + hra + arrears + fa + ta_sa)

        # Deductions
        pt = calc_pt(gross_total)
        wf = 75.0 if gross_total > 0 and domain in ["Teaching", "Non-Teaching", "Support Staff", "Management"] else 0.0
        epf = epf_fixed
        it = float(deductions.get("income_tax", it_fixed) or 0.0)
        eb = float(deductions.get("eb_deduction", 0.0) or 0.0)
        mess = float(deductions.get("mess_deduction", 0.0) or 0.0)
        bus = float(deductions.get("bus_deduction", 0.0) or 0.0)
        other_ded = float(deductions.get("other_deductions", 0.0) or 0.0)

        total_ded = epf + it + pt + wf + eb + mess + bus + other_ded
        net_salary = max(0.0, gross_total - total_ded)

        return {
            "emp_id": emp.get("emp_id"),
            "emp_name": emp.get("emp_name"),
            "domain": domain,
            "category": category,
            "department": department,
            "designation": emp.get("designation", ""),
            "standard_salary": standard_salary,
            "consolidated_salary": consolidated_salary,
            "basic": basic_full,
            "agp": agp,
            "month_days": month_days,
            "biometric_days": float(att_dict.get("biometric_days", 0.0) or 0.0),
            "holiday_days": float(att_dict.get("holiday_days", 0.0) or 0.0),
            "availed_leaves": float(att_dict.get("availed_leaves", 0.0) or 0.0),
            "od_days": float(att_dict.get("od_days", 0.0) or 0.0),
            "total_pay_days": total_pay_days,
            "basic_agp": basic_agp,
            "da": da,
            "hra": hra,
            "arrears": arrears,
            "fa": fa,
            "ta_sa": ta_sa,
            "gross_total": gross_total,
            "epf": epf,
            "it": it,
            "pt": pt,
            "wf": wf,
            "eb": eb,
            "mess": mess,
            "bus": bus,
            "other_ded": other_ded,
            "tot_ded": total_ded,
            "net_salary": net_salary,
            "account_no": emp.get("account_no", ""),
            "ifsc_code": emp.get("ifsc_code", ""),
            "remarks": att_dict.get("remarks", "")
        }

    def process_payroll(self, attendance_map, month_year="2026-08", month_days=31, deductions_map=None):
        if deductions_map is None:
            deductions_map = {}

        # Always ensure guaranteed exempt employees exist in attendance_map
        attendance_map = inject_guaranteed_exempt_employees(dict(attendance_map), month_days)

        emp_master = self.get_employees_map()
        processed_items = []

        total_gross = 0.0
        total_ded = 0.0
        total_net = 0.0

        all_emp_ids = set(attendance_map.keys()) | set(emp_master.keys())

        # Group stats by domain and by department
        domain_stats = {}
        dept_stats = {}

        for eid in sorted(all_emp_ids, key=lambda x: (str(x).isdigit(), int(x) if str(x).isdigit() else 999999, str(x))):
            emp = emp_master.get(eid, {
                "emp_id": eid,
                "emp_name": attendance_map.get(eid, {}).get("emp_name", f"Emp {eid}"),
                "domain": "Non-Teaching",
                "category": "Non-Teaching",
                "department": attendance_map.get(eid, {}).get("department", "General"),
                "standard_salary": 0.0
            })

            att = attendance_map.get(eid, {
                "biometric_days": 0.0,
                "holiday_days": 0.0,
                "availed_leaves": 0.0,
                "od_days": 0.0,
                "total_pay_days": 0.0
            })

            ded = deductions_map.get(eid, {})

            item = self.compute_employee_salary(emp, att, month_days=month_days, deductions=ded)
            
            # Include all university employees from the database
            processed_items.append(item)
            total_gross += item["gross_total"]
            total_ded += item["tot_ded"]
            total_net += item["net_salary"]

            # Aggregate by Domain
            d = item["domain"] or "General"
            if d not in domain_stats:
                domain_stats[d] = {"count": 0, "gross": 0.0, "deductions": 0.0, "net": 0.0}
            domain_stats[d]["count"] += 1
            domain_stats[d]["gross"] += item["gross_total"]
            domain_stats[d]["deductions"] += item["tot_ded"]
            domain_stats[d]["net"] += item["net_salary"]

            # Aggregate by Department
            dept = item["department"] or "General"
            if dept not in dept_stats:
                dept_stats[dept] = {"domain": d, "count": 0, "gross": 0.0, "deductions": 0.0, "net": 0.0}
            dept_stats[dept]["count"] += 1
            dept_stats[dept]["gross"] += item["gross_total"]
            dept_stats[dept]["deductions"] += item["tot_ded"]
            dept_stats[dept]["net"] += item["net_salary"]

        summary = {
            "month_year": month_year,
            "month_days": month_days,
            "total_employees": len(processed_items),
            "total_gross": total_gross,
            "total_deductions": total_ded,
            "total_net": total_net,
            "domain_stats": domain_stats,
            "department_stats": dept_stats,
            "items": processed_items
        }
        return summary
