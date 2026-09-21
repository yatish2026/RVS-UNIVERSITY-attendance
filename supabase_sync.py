import os
import json
import sqlite3
import urllib.request
import urllib.parse
from datetime import datetime
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

class SupabaseSync:
    def __init__(self, db_path="payroll_master.db", url=None, key=None):
        self.db_path = db_path
        self.url = (url or os.environ.get("SUPABASE_URL", "")).rstrip("/")
        self.key = key or os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_KEY") or os.environ.get("SUPABASE_ANON_KEY", "")
        self._cached_employees = None

    def is_configured(self):
        return bool(self.url and self.key)

    def _headers(self, content_type="application/json", prefer=None):
        h = {
            "apikey": self.key,
            "Authorization": f"Bearer {self.key}"
        }
        if content_type:
            h["Content-Type"] = content_type
        if prefer:
            h["Prefer"] = prefer
        return h

    def test_connection(self):
        if not self.is_configured():
            return {"status": "unconfigured", "message": "Supabase URL and API Key not provided"}
        
        try:
            req_url = f"{self.url}/rest/v1/employees?select=count"
            req = urllib.request.Request(req_url, headers={
                "apikey": self.key,
                "Authorization": f"Bearer {self.key}",
                "Range": "0-0"
            })
            with urllib.request.urlopen(req, timeout=8) as response:
                return {"status": "connected", "message": "Successfully connected to Supabase cloud database!"}
        except urllib.error.HTTPError as e:
            if e.code == 404 or "PGRST205" in str(e):
                return {"status": "needs_schema", "message": "Connected to Supabase! Please run supabase_schema.sql in Supabase SQL Editor."}
            return {"status": "error", "message": f"HTTP Error {e.code}: {e.reason}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def get_all_employees(self, use_cache=False):
        """
        Fetches all employees directly from Supabase Cloud.
        """
        if use_cache and self._cached_employees is not None:
            return self._cached_employees

        if not self.is_configured():
            return self._get_local_employees()

        records = []
        page_size = 1000
        offset = 0

        try:
            while True:
                req_url = f"{self.url}/rest/v1/employees?select=*&order=emp_id.asc"
                req = urllib.request.Request(req_url, headers={
                    "apikey": self.key,
                    "Authorization": f"Bearer {self.key}",
                    "Range": f"{offset}-{offset + page_size - 1}"
                })

                with urllib.request.urlopen(req, timeout=15) as response:
                    batch = json.loads(response.read().decode("utf-8"))
                    if not batch:
                        break
                    records.extend(batch)
                    if len(batch) < page_size:
                        break
                    offset += page_size
            
            self._cached_employees = records
            return records
        except Exception as e:
            print(f"[SupabaseSync] Direct fetch error: {e}. Falling back to local cache if available.")
            return self._get_local_employees()

    def get_employees_map(self):
        """
        Returns a dictionary {emp_id: employee_dict} loaded directly from Supabase.
        """
        records = self.get_all_employees()
        return {str(r.get("emp_id")): r for r in records}

    def get_filtered_employees(self, search="", domain="ALL", department="ALL"):
        """
        Filters employees in-memory after fast cloud load.
        """
        records = self.get_all_employees(use_cache=True)
        results = []
        s_lower = search.strip().lower() if search else ""

        for r in records:
            if domain and domain != "ALL" and r.get("domain") != domain:
                continue
            if department and department != "ALL" and r.get("department") != department:
                continue
            if s_lower:
                eid = str(r.get("emp_id", "")).lower()
                name = str(r.get("emp_name", "")).lower()
                acc = str(r.get("account_no", "")).lower()
                if s_lower not in eid and s_lower not in name and s_lower not in acc:
                    continue
            results.append(r)

        # Sort logically
        def sort_key(x):
            eid = str(x.get("emp_id", ""))
            return (0, int(eid)) if eid.isdigit() else (1, eid)

        results.sort(key=sort_key)
        return results

    def save_employee(self, data):
        """
        Saves or updates an employee directly in Supabase Cloud.
        """
        if not self.is_configured():
            return self._save_local_employee(data)

        emp_id = str(data.get("emp_id")).strip()
        emp_name = str(data.get("emp_name")).strip()
        if not emp_id or not emp_name:
            return {"status": "error", "message": "Employee ID and Name are required"}

        allowed_keys = {
            "emp_id", "emp_name", "domain", "category", "department", "designation",
            "salary_type", "standard_salary", "consolidated_salary", "basic", "agp",
            "fa", "ta_sa", "epf_fixed", "it_fixed", "account_no", "ifsc_code"
        }

        payload = {k: v for k, v in data.items() if k in allowed_keys}
        payload["emp_id"] = emp_id
        payload["emp_name"] = emp_name
        payload["domain"] = payload.get("domain", "Teaching")
        payload["category"] = payload.get("category", payload["domain"])
        payload["department"] = payload.get("department", "General")
        payload["standard_salary"] = float(payload.get("standard_salary", 0) or 0)
        payload["consolidated_salary"] = float(payload.get("consolidated_salary", 0) or 0)
        payload["basic"] = float(payload.get("basic", 0) or 0)
        payload["agp"] = float(payload.get("agp", 0) or 0)
        payload["fa"] = float(payload.get("fa", 0) or 0)
        payload["ta_sa"] = float(payload.get("ta_sa", 0) or 0)
        payload["epf_fixed"] = float(payload.get("epf_fixed", 0) or 0)
        payload["it_fixed"] = float(payload.get("it_fixed", 0) or 0)

        try:
            req_url = f"{self.url}/rest/v1/employees"
            body = json.dumps([payload]).encode("utf-8")
            req = urllib.request.Request(req_url, data=body, headers=self._headers(prefer="resolution=merge-duplicates"), method="POST")
            with urllib.request.urlopen(req, timeout=10) as response:
                self._cached_employees = None  # Invalidate cache
                self._save_local_employee(data)  # Keep local mirror
                return {"status": "success", "message": f"Employee {emp_id} successfully saved in Supabase Cloud database"}
        except Exception as e:
            return {"status": "error", "message": f"Failed to save in Supabase: {str(e)}"}

    def delete_employee(self, emp_id):
        """
        Deletes an employee directly in Supabase Cloud.
        """
        if not self.is_configured():
            return self._delete_local_employee(emp_id)

        try:
            safe_id = urllib.parse.quote(str(emp_id))
            req_url = f"{self.url}/rest/v1/employees?emp_id=eq.{safe_id}"
            req = urllib.request.Request(req_url, headers=self._headers(), method="DELETE")
            with urllib.request.urlopen(req, timeout=10) as response:
                self._cached_employees = None
                self._delete_local_employee(emp_id)
                return {"status": "success", "message": f"Employee {emp_id} deleted from Supabase Cloud database"}
        except Exception as e:
            return {"status": "error", "message": f"Failed to delete in Supabase: {str(e)}"}

    def save_payroll_run(self, month_year, month_days, summary):
        """
        Saves a finalized monthly payroll run and all employee salary line items to Supabase Cloud.
        """
        if not self.is_configured():
            return {"status": "error", "message": "Supabase credentials not configured"}

        try:
            # 1. Delete existing run for this month if exists
            safe_month = urllib.parse.quote(month_year)
            del_url = f"{self.url}/rest/v1/payroll_runs?month_year=eq.{safe_month}"
            del_req = urllib.request.Request(del_url, headers=self._headers(), method="DELETE")
            try:
                urllib.request.urlopen(del_req, timeout=8)
            except Exception:
                pass

            # 2. Insert new payroll run
            run_payload = {
                "month_year": month_year,
                "total_days_in_month": month_days,
                "total_employees": summary.get("total_employees", 0),
                "total_gross": summary.get("total_gross", 0.0),
                "total_deductions": summary.get("total_deductions", 0.0),
                "total_net": summary.get("total_net", 0.0)
            }
            run_url = f"{self.url}/rest/v1/payroll_runs"
            req = urllib.request.Request(run_url, data=json.dumps([run_payload]).encode("utf-8"),
                                         headers=self._headers(prefer="return=representation"), method="POST")
            
            run_id = None
            with urllib.request.urlopen(req, timeout=10) as response:
                res_data = json.loads(response.read().decode("utf-8"))
                if res_data and isinstance(res_data, list):
                    run_id = res_data[0].get("id")

            # 3. Insert payroll items in chunks
            items = summary.get("items", [])
            if items:
                allowed_item_keys = {
                    "emp_id", "emp_name", "domain", "category", "department", "designation",
                    "total_salary", "basic", "agp", "total_pay_days", "basic_agp", "da",
                    "hra", "arrears", "fa", "ta_sa", "gross_total", "epf", "it", "pt",
                    "wf", "eb", "mess", "bus", "tot_ded", "net_salary", "account_no", "ifsc_code"
                }

                chunk_size = 1000
                for i in range(0, len(items), chunk_size):
                    chunk = []
                    for it in items[i:i + chunk_size]:
                        row = {k: (0.0 if it.get(k) is None else it.get(k)) for k in allowed_item_keys}
                        row["run_id"] = run_id
                        row["month_days"] = month_days
                        chunk.append(row)

                    items_url = f"{self.url}/rest/v1/payroll_items"
                    req_items = urllib.request.Request(items_url, data=json.dumps(chunk).encode("utf-8"),
                                                       headers=self._headers(), method="POST")
                    try:
                        with urllib.request.urlopen(req_items, timeout=15):
                            pass
                    except Exception as ie:
                        print(f"[Supabase Payroll Item Insert Warning]: {ie}")

            return {"status": "success", "run_id": run_id, "message": f"Payroll for {month_year} successfully saved and locked in Supabase Cloud!"}
        except Exception as e:
            return {"status": "error", "message": f"Failed to save payroll run in Supabase: {str(e)}"}

    def get_payroll_runs(self):
        """
        Fetches all saved payroll runs from Supabase Cloud.
        """
        if not self.is_configured():
            return self._get_local_runs()

        try:
            req_url = f"{self.url}/rest/v1/payroll_runs?select=*&order=id.desc"
            req = urllib.request.Request(req_url, headers=self._headers())
            with urllib.request.urlopen(req, timeout=10) as response:
                return json.loads(response.read().decode("utf-8"))
        except Exception as e:
            return self._get_local_runs()

    def get_payroll_run_data(self, month_year):
        """
        Loads the complete payroll run and all employee salary line items for a specific month from Supabase.
        """
        if not self.is_configured():
            return None

        try:
            safe_month = urllib.parse.quote(month_year)
            run_url = f"{self.url}/rest/v1/payroll_runs?month_year=eq.{safe_month}&select=*"
            req = urllib.request.Request(run_url, headers=self._headers())
            with urllib.request.urlopen(req, timeout=10) as response:
                runs = json.loads(response.read().decode("utf-8"))
                if not runs:
                    return None
                run = runs[0]
                run_id = run["id"]

            items = []
            page_size = 1000
            offset = 0
            while True:
                items_url = f"{self.url}/rest/v1/payroll_items?run_id=eq.{run_id}&select=*&order=id.asc"
                req_items = urllib.request.Request(items_url, headers={
                    **self._headers(),
                    "Range": f"{offset}-{offset + page_size - 1}"
                })
                with urllib.request.urlopen(req_items, timeout=15) as response:
                    batch = json.loads(response.read().decode("utf-8"))
                    if not batch:
                        break
                    items.extend(batch)
                    if len(batch) < page_size:
                        break
                    offset += page_size

            # Reconstruct domain_stats and department_stats
            domain_stats = {}
            dept_stats = {}
            for it in items:
                d = it.get("domain") or "General"
                dept = it.get("department") or "General"
                if d not in domain_stats:
                    domain_stats[d] = {"count": 0, "gross": 0.0, "deductions": 0.0, "net": 0.0}
                domain_stats[d]["count"] += 1
                domain_stats[d]["gross"] += float(it.get("gross_total", 0) or 0)
                domain_stats[d]["deductions"] += float(it.get("tot_ded", 0) or 0)
                domain_stats[d]["net"] += float(it.get("net_salary", 0) or 0)

                if dept not in dept_stats:
                    dept_stats[dept] = {"domain": d, "count": 0, "gross": 0.0, "deductions": 0.0, "net": 0.0}
                dept_stats[dept]["count"] += 1
                dept_stats[dept]["gross"] += float(it.get("gross_total", 0) or 0)
                dept_stats[dept]["deductions"] += float(it.get("tot_ded", 0) or 0)
                dept_stats[dept]["net"] += float(it.get("net_salary", 0) or 0)

            return {
                "month_year": run.get("month_year"),
                "month_days": run.get("total_days_in_month", 31),
                "total_employees": run.get("total_employees", len(items)),
                "total_gross": float(run.get("total_gross", 0) or 0),
                "total_deductions": float(run.get("total_deductions", 0) or 0),
                "total_net": float(run.get("total_net", 0) or 0),
                "domain_stats": domain_stats,
                "department_stats": dept_stats,
                "items": items
            }
        except Exception as e:
            print(f"[SupabaseSync] Error loading payroll run for {month_year}: {e}")
            return None

    def get_stats(self):
        """
        Calculates live institution statistics directly from Supabase Cloud data.
        """
        records = self.get_all_employees(use_cache=True)
        total_emp = len(records)
        total_base_sal = sum(float(r.get("standard_salary", 0) or 0) for r in records)
        
        domains = {}
        depts_map = {}

        for r in records:
            d = r.get("domain") or "General"
            dept = r.get("department") or "General"
            sal = float(r.get("standard_salary", 0) or 0)

            if d not in domains:
                domains[d] = {"count": 0, "base_salary": 0.0}
            domains[d]["count"] += 1
            domains[d]["base_salary"] += sal

            key = (dept, d)
            depts_map[key] = depts_map.get(key, 0) + 1

        departments = [{"department": k[0], "domain": k[1], "cnt": v} for k, v in depts_map.items()]
        departments.sort(key=lambda x: x["cnt"], reverse=True)
        recent_runs = self.get_payroll_runs()

        return {
            "total_employees": total_emp,
            "total_departments": len(depts_map),
            "total_base_payroll": total_base_sal,
            "domains": domains,
            "departments": departments,
            "recent_runs": recent_runs,
            "supabase_connected": True
        }

    # =========================================================================
    # Fallback / Local Mirror Helpers (only if network offline)
    # =========================================================================
    def _get_local_employees(self):
        if not os.path.exists(self.db_path):
            return []
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM employees")
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def _save_local_employee(self, data):
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("""
        INSERT OR REPLACE INTO employees
        (emp_id, emp_name, domain, category, department, designation, salary_type, standard_salary, consolidated_salary, basic, agp, fa, ta_sa, epf_fixed, it_fixed, account_no, ifsc_code)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data.get("emp_id"), data.get("emp_name"), data.get("domain", "Teaching"), data.get("category", data.get("domain", "Teaching")),
            data.get("department", "General"), data.get("designation", ""), data.get("salary_type", "REGULAR"),
            float(data.get("standard_salary", 0) or 0), float(data.get("consolidated_salary", 0) or 0),
            float(data.get("basic", 0) or 0), float(data.get("agp", 0) or 0), float(data.get("fa", 0) or 0),
            float(data.get("ta_sa", 0) or 0), float(data.get("epf_fixed", 0) or 0), float(data.get("it_fixed", 0) or 0),
            data.get("account_no", ""), data.get("ifsc_code", "PUNB0401700")
        ))
        conn.commit()
        conn.close()

    def _delete_local_employee(self, emp_id):
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        cur.execute("DELETE FROM employees WHERE emp_id = ?", (emp_id,))
        conn.commit()
        conn.close()

    def _get_local_runs(self):
        if not os.path.exists(self.db_path):
            return []
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM payroll_runs ORDER BY id DESC LIMIT 10")
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def sync_to_supabase(self):
        rows = self._get_local_employees()
        if not rows:
            return {"status": "empty", "message": "No local employees to sync"}

        allowed_keys = {
            "emp_id", "emp_name", "domain", "category", "department", "designation",
            "salary_type", "standard_salary", "consolidated_salary", "basic", "agp",
            "fa", "ta_sa", "epf_fixed", "it_fixed", "account_no", "ifsc_code"
        }
        sanitized_rows = [{k: v for k, v in r.items() if k in allowed_keys} for r in rows]
        chunk_size = 100
        total_synced = 0
        
        for i in range(0, len(sanitized_rows), chunk_size):
            chunk = sanitized_rows[i:i + chunk_size]
            try:
                req_url = f"{self.url}/rest/v1/employees"
                data = json.dumps(chunk).encode("utf-8")
                req = urllib.request.Request(req_url, data=data, headers=self._headers(prefer="resolution=merge-duplicates"), method="POST")
                with urllib.request.urlopen(req, timeout=15) as response:
                    total_synced += len(chunk)
            except Exception as e:
                return {"status": "error", "message": f"Failed to sync to Supabase: {str(e)}"}

        self._cached_employees = None
        return {"status": "success", "synced_count": total_synced, "message": f"Successfully synced {total_synced} employees to Supabase cloud database!"}

    def pull_from_supabase(self):
        records = self.get_all_employees()
        conn = sqlite3.connect(self.db_path)
        cur = conn.cursor()
        for r in records:
            cur.execute("""
            INSERT OR REPLACE INTO employees 
            (emp_id, emp_name, domain, category, department, designation, salary_type, standard_salary, consolidated_salary, basic, agp, fa, ta_sa, epf_fixed, it_fixed, account_no, ifsc_code)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                r.get("emp_id"), r.get("emp_name"), r.get("domain", "Teaching"), r.get("category", "Teaching"),
                r.get("department", "General"), r.get("designation", ""), r.get("salary_type", "REGULAR"),
                float(r.get("standard_salary", 0) or 0), float(r.get("consolidated_salary", 0) or 0),
                float(r.get("basic", 0) or 0), float(r.get("agp", 0) or 0), float(r.get("fa", 0) or 0),
                float(r.get("ta_sa", 0) or 0), float(r.get("epf_fixed", 0) or 0), float(r.get("it_fixed", 0) or 0),
                r.get("account_no", ""), r.get("ifsc_code", "PUNB0401700")
            ))
        conn.commit()
        conn.close()
        return {"status": "success", "pulled_count": len(records), "message": f"Updated {len(records)} local employee records from Supabase"}

if __name__ == "__main__":
    client = SupabaseSync()
    print("Supabase configured:", client.is_configured())
    print("Testing connection:", client.test_connection())
    stats = client.get_stats()
    print(f"Loaded {stats['total_employees']} employees from Supabase Cloud!")
