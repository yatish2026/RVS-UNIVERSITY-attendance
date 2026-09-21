import os
import json
import sqlite3
import urllib.request
import urllib.parse

DEFAULT_SUPABASE_URL = "https://fzdzaowjiapxvfrxrdzc.supabase.co"
DEFAULT_SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ6ZHphb3dqaWFweHZmcnhyZHpjIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4OTk3Nzk0MiwiZXhwIjoyMTA1NTUzOTQyfQ.RmQouMmkGYPl3NSFxu0Lf_gzGDxGs8zuXmHR7foLW6Y"

class SupabaseSync:
    def __init__(self, db_path="payroll_master.db", url=None, key=None):
        self.db_path = db_path
        self.url = (url or os.environ.get("SUPABASE_URL", DEFAULT_SUPABASE_URL)).rstrip("/")
        self.key = key or os.environ.get("SUPABASE_KEY", DEFAULT_SUPABASE_KEY)

    def is_configured(self):
        return bool(self.url and self.key)

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
            with urllib.request.urlopen(req, timeout=5) as response:
                return {"status": "connected", "message": "Successfully connected to Supabase project!"}
        except urllib.error.HTTPError as e:
            if e.code == 404 or "PGRST205" in str(e):
                return {"status": "needs_schema", "message": "Connected to Supabase! Please run supabase_schema.sql in Supabase SQL Editor to create the employees table."}
            return {"status": "error", "message": f"HTTP Error {e.code}: {e.reason}"}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def sync_to_supabase(self):
        """
        Pushes all local employees from SQLite to Supabase 'employees' table in batches.
        """
        if not self.is_configured():
            return {"status": "skipped", "message": "Supabase credentials not configured. Local database is active."}

        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT * FROM employees")
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()

        if not rows:
            return {"status": "empty", "message": "No local employees to sync"}

        chunk_size = 100
        total_synced = 0
        
        for i in range(0, len(rows), chunk_size):
            chunk = rows[i:i + chunk_size]
            try:
                req_url = f"{self.url}/rest/v1/employees"
                data = json.dumps(chunk).encode("utf-8")
                req = urllib.request.Request(req_url, data=data, headers={
                    "apikey": self.key,
                    "Authorization": f"Bearer {self.key}",
                    "Content-Type": "application/json",
                    "Prefer": "resolution=merge-duplicates"
                }, method="POST")

                with urllib.request.urlopen(req, timeout=15) as response:
                    total_synced += len(chunk)
            except urllib.error.HTTPError as e:
                err_body = e.read().decode('utf-8', errors='ignore')
                return {"status": "error", "message": f"Supabase sync failed (HTTP {e.code}): {err_body}"}
            except Exception as e:
                return {"status": "error", "message": f"Failed to sync to Supabase: {str(e)}"}

        return {"status": "success", "synced_count": total_synced, "message": f"Successfully synced {total_synced} employees to Supabase cloud database!"}

    def pull_from_supabase(self):
        """
        Pulls employees from Supabase down to SQLite.
        """
        if not self.is_configured():
            return {"status": "skipped", "message": "Supabase credentials not configured."}

        try:
            req_url = f"{self.url}/rest/v1/employees?select=*"
            req = urllib.request.Request(req_url, headers={
                "apikey": self.key,
                "Authorization": f"Bearer {self.key}"
            })

            with urllib.request.urlopen(req, timeout=15) as response:
                records = json.loads(response.read().decode("utf-8"))
                
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
        except urllib.error.HTTPError as e:
            err_body = e.read().decode('utf-8', errors='ignore')
            return {"status": "error", "message": f"Supabase pull failed: {err_body}"}
        except Exception as e:
            return {"status": "error", "message": f"Failed to pull from Supabase: {str(e)}"}

if __name__ == "__main__":
    client = SupabaseSync()
    print("Supabase configured:", client.is_configured())
    print("Testing connection:", client.test_connection())
