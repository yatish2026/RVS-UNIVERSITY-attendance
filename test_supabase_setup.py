import urllib.request
import json
import sqlite3

url = "https://fzdzaowjiapxvfrxrdzc.supabase.co"
service_key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ6ZHphb3dqaWFweHZmcnhyZHpjIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4OTk3Nzk0MiwiZXhwIjoyMTA1NTUzOTQyfQ.RmQouMmkGYPl3NSFxu0Lf_gzGDxGs8zuXmHR7foLW6Y"
anon_key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ6ZHphb3dqaWFweHZmcnhyZHpjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk5Nzc5NDIsImV4cCI6MjEwNTU1Mzk0Mn0.cvJnMO3fR7D7L9ys5ng8go7myi2F8gbC2h11GsjfgAc"

print("Checking endpoints...")

# Let's check direct SQL endpoint or pgmeta
endpoints = [
    f"{url}/rest/v1/rpc",
    f"{url}/pg/query",
    f"{url}/pg-meta/default/query"
]

for ep in endpoints:
    try:
        req = urllib.request.Request(ep, headers={
            "apikey": service_key,
            "Authorization": f"Bearer {service_key}",
            "Content-Type": "application/json"
        }, data=json.dumps({"query": "SELECT 1"}).encode("utf-8"), method="POST")
        with urllib.request.urlopen(req, timeout=5) as resp:
            print(f"{ep} -> Status: {resp.status}, Body: {resp.read().decode('utf-8')[:100]}")
    except Exception as e:
        print(f"{ep} -> Error: {e}")
