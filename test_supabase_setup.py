import os
import urllib.request
import json
import sqlite3
from dotenv import load_dotenv

load_dotenv()

url = os.environ.get("SUPABASE_URL", "")
service_key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
anon_key = os.environ.get("SUPABASE_ANON_KEY", "")

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
