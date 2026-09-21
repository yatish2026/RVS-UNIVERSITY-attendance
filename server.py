import os
import io
import json
import sqlite3
import shutil
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from biometric_parser import BiometricParser
from payroll_engine import PayrollEngine
from excel_exporter import ExcelExporter
from supabase_sync import SupabaseSync

import tempfile
import threading

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# Temporary directories for uploads and exports (works in local, Render, and Vercel serverless)
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "rvs_uploads")
EXPORT_DIR = os.path.join(tempfile.gettempdir(), "rvs_exports")
try:
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(EXPORT_DIR, exist_ok=True)
except Exception:
    pass

DB_PATH = os.environ.get("DB_PATH", os.path.join(tempfile.gettempdir(), "payroll_master.db") if os.environ.get("VERCEL") else "payroll_master.db")
payroll_engine = PayrollEngine(DB_PATH)
excel_exporter = ExcelExporter()
supabase_client = SupabaseSync(DB_PATH)

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route("/api/stats", methods=["GET"])
def get_stats():
    return jsonify(supabase_client.get_stats())

@app.route("/api/employees", methods=["GET"])
def get_employees():
    search = request.args.get("search", "")
    domain = request.args.get("domain", "ALL")
    department = request.args.get("department", "ALL")
    rows = supabase_client.get_filtered_employees(search=search, domain=domain, department=department)
    return jsonify({"employees": rows, "count": len(rows)})

@app.route("/api/employees", methods=["POST"])
def save_employee():
    data = request.get_json(force=True)
    res = supabase_client.save_employee(data)
    if res.get("status") == "error":
        return jsonify(res), 400
    return jsonify(res)

@app.route("/api/employees/<emp_id>", methods=["DELETE"])
def delete_employee(emp_id):
    res = supabase_client.delete_employee(emp_id)
    return jsonify(res)

@app.route("/api/upload-biometric", methods=["POST"])
def upload_biometric():
    if "file" not in request.files:
        return jsonify({"status": "error", "message": "No file uploaded"}), 400
    
    file = request.files["file"]
    month_year = request.form.get("month_year", "August 2026")
    month_days = int(request.form.get("month_days", 31))

    temp_path = os.path.join(UPLOAD_DIR, f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
    file.save(temp_path)

    try:
        parser = BiometricParser(temp_path)
        attendance_map = parser.parse(month_days=month_days)

        if not attendance_map:
            return jsonify({"status": "error", "message": "No valid attendance records found in uploaded file"}), 400

        summary = payroll_engine.process_payroll(attendance_map, month_year, month_days)

        # Auto-save run and items directly to Supabase Cloud in background thread to guarantee fast response
        if supabase_client.is_configured():
            def async_save_run(m_yr, m_days, p_sum):
                try:
                    supabase_client.save_payroll_run(m_yr, m_days, p_sum)
                except Exception as se:
                    print(f"[Supabase Auto-Save Warning]: {se}")

            threading.Thread(target=async_save_run, args=(month_year, month_days, summary), daemon=True).start()

        return jsonify({
            "status": "success",
            "filename": file.filename,
            "parsed_employees": len(attendance_map),
            "attendance": attendance_map,
            "payroll_summary": summary
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/calculate-payroll", methods=["POST"])
def calculate_payroll():
    data = request.get_json(force=True)
    month_year = data.get("month_year", "August 2026")
    month_days = int(data.get("month_days", 31))
    attendance = data.get("attendance", {})
    deductions = data.get("deductions", {})

    try:
        summary = payroll_engine.process_payroll(
            attendance,
            month_year=month_year,
            month_days=month_days,
            deductions_map=deductions
        )
        if supabase_client.is_configured():
            def async_save_run(m_yr, m_days, p_sum):
                try:
                    supabase_client.save_payroll_run(m_yr, m_days, p_sum)
                except Exception as se:
                    print(f"[Supabase Recalc Auto-Save Warning]: {se}")

            threading.Thread(target=async_save_run, args=(month_year, month_days, summary), daemon=True).start()

        return jsonify({"status": "success", "payroll_summary": summary})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/payroll-runs", methods=["GET"])
def get_payroll_runs():
    runs = supabase_client.get_payroll_runs()
    return jsonify({"runs": runs})

@app.route("/api/payroll-runs/<path:month_year>", methods=["GET"])
def get_payroll_run_month(month_year):
    data = supabase_client.get_payroll_run_data(month_year)
    if not data:
        return jsonify({"status": "error", "message": f"No saved payroll data found for {month_year}"}), 404
    return jsonify({"status": "success", "payroll_summary": data})

@app.route("/api/export-excel", methods=["POST"])
def export_excel():
    data = request.get_json(force=True)
    month_year = data.get("month_year", "August 2026")
    month_days = int(data.get("month_days", 31))
    attendance = data.get("attendance", {})
    deductions = data.get("deductions", {})
    domain = data.get("domain", "ALL")
    department = data.get("department", "ALL")

    try:
        summary = payroll_engine.process_payroll(
            attendance,
            month_year=month_year,
            month_days=month_days,
            deductions_map=deductions
        )
        safe_month = month_year.replace(" ", "_").replace("-", "_")
        if department and department != "ALL":
            clean_dept = department.replace(" ", "_").replace("/", "_")
            filename = f"{clean_dept}_Department_Salary_Bill_{safe_month}.xlsx"
        elif domain and domain != "ALL":
            filename = f"{domain.replace(' ', '_')}_Department_Separated_Salary_Bill_{safe_month}.xlsx"
        else:
            filename = f"Master_University_Salary_Bill_{safe_month}.xlsx"

        filepath = os.path.join(EXPORT_DIR, filename)
        excel_exporter.export_salary_bill(summary, filepath, selected_domain=domain, selected_department=department)
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/supabase/config", methods=["POST"])
def config_supabase():
    data = request.get_json(force=True)
    supabase_client.url = data.get("supabase_url", "").rstrip("/")
    supabase_client.key = data.get("supabase_key", "")
    res = supabase_client.test_connection()
    return jsonify(res)

@app.route("/api/supabase/sync", methods=["POST"])
def sync_supabase():
    direction = request.args.get("direction", "push")
    if direction == "push":
        res = supabase_client.sync_to_supabase()
    else:
        res = supabase_client.pull_from_supabase()
    return jsonify(res)

@app.route("/")
def index():
    return render_template_string(HTML_CONTENT)

HTML_CONTENT = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>University Biometric Attendance & Payroll Engine</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  <script src="https://unpkg.com/lucide@latest"></script>
  <style>
    :root {
      --bg: #080c14;
      --bg-surface: #0e1422;
      --bg-card: rgba(14, 20, 34, 0.85);
      --bg-card-hover: rgba(22, 31, 51, 0.9);
      --border: #1a2438;
      --border-subtle: rgba(255, 255, 255, 0.05);
      --border-accent: rgba(56, 189, 248, 0.25);
      --primary: #38bdf8;
      --primary-hover: #0284c7;
      --primary-gradient: linear-gradient(135deg, #131d2e 0%, #0b111d 100%);
      --accent-slate: #1e293b;
      --accent-blue: #3b82f6;
      --success: #34d399;
      --success-bg: rgba(52, 211, 153, 0.08);
      --success-border: rgba(52, 211, 153, 0.25);
      --warning: #fbbf24;
      --danger: #f87171;
      --danger-bg: rgba(248, 113, 113, 0.08);
      --text: #f1f5f9;
      --text-muted: #8291a5;
      --text-dim: #55657d;
      --font-display: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
      --font-body: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
      --font-mono: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", monospace;
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      background-color: var(--bg);
      color: var(--text);
      font-family: var(--font-body);
      min-height: 100vh;
      overflow-x: hidden;
      background-image: 
        radial-gradient(circle at 15% 15%, rgba(56, 189, 248, 0.025) 0%, transparent 50%),
        radial-gradient(circle at 85% 85%, rgba(130, 145, 165, 0.02) 0%, transparent 50%);
    }

    header {
      backdrop-filter: blur(20px);
      background: rgba(8, 12, 20, 0.92);
      border-bottom: 1px solid var(--border);
      position: sticky;
      top: 0;
      z-index: 50;
      padding: 0.85rem 2rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }

    .brand { display: flex; align-items: center; gap: 0.85rem; }
    .logo-icon {
      width: 40px; height: 40px; border-radius: 10px;
      background: #111827;
      border: 1px solid #233149;
      display: flex; align-items: center; justify-content: center;
      color: #38bdf8;
    }
    .brand-title {
      font-family: var(--font-display); font-size: 1.125rem; font-weight: 700;
      color: #f8fafc;
      letter-spacing: -0.01em;
    }
    .brand-subtitle { font-size: 0.75rem; color: var(--text-muted); font-weight: 500; }

    nav {
      display: flex; gap: 0.35rem;
      background: #0e1422; padding: 0.3rem;
      border-radius: 9px; border: 1px solid var(--border);
    }
    .nav-btn {
      background: transparent; border: 1px solid transparent; color: var(--text-muted);
      padding: 0.45rem 0.85rem; border-radius: 6px; font-size: 0.825rem;
      font-weight: 600; cursor: pointer; display: flex; align-items: center;
      gap: 0.45rem; transition: all 0.15s ease;
    }
    .nav-btn:hover { color: var(--text); background: rgba(255, 255, 255, 0.03); }
    .nav-btn.active {
      background: #162238;
      border-color: #26385a;
      color: #60a5fa;
    }

    .badge-supabase {
      display: flex; align-items: center; gap: 0.4rem;
      padding: 0.4rem 0.8rem; border-radius: 8px; font-size: 0.75rem;
      font-weight: 600; background: rgba(52, 211, 153, 0.08);
      border: 1px solid rgba(52, 211, 153, 0.2); color: var(--success);
      cursor: pointer; transition: all 0.15s ease;
    }
    .badge-supabase:hover {
      background: rgba(52, 211, 153, 0.14);
      border-color: rgba(52, 211, 153, 0.35);
    }
    .badge-supabase.offline {
      background: rgba(130, 145, 165, 0.08); border-color: rgba(130, 145, 165, 0.18);
      color: var(--text-muted);
    }

    .container { max-width: 1440px; margin: 0 auto; padding: 2rem; }
    .tab-content { display: none; }
    .tab-content.active { display: block; animation: fadeIn 0.25s ease; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(3px); } to { opacity: 1; transform: translateY(0); } }

    .metrics-grid {
      display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 1.25rem; margin-bottom: 2rem;
    }
    .metric-card {
      background: var(--bg-card); backdrop-filter: blur(12px);
      border: 1px solid var(--border); border-radius: 12px;
      padding: 1.35rem; position: relative; overflow: hidden; transition: all 0.2s ease;
    }
    .metric-card:hover { border-color: #2a3a5a; }
    .metric-card::after {
      content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
      background: #2563eb; opacity: 0.6;
    }
    .metric-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.75rem; }
    .metric-label { font-size: 0.75rem; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.04em; }
    .metric-icon {
      width: 34px; height: 34px; border-radius: 8px;
      background: #131c2e; border: 1px solid #1f2d48; color: #60a5fa;
      display: flex; align-items: center; justify-content: center;
    }
    .metric-value { font-family: var(--font-display); font-size: 1.75rem; font-weight: 700; color: var(--text); font-variant-numeric: tabular-nums; }
    .metric-desc { font-size: 0.75rem; color: var(--text-dim); margin-top: 0.35rem; }

    /* Domain Pills Bar */
    .domain-tabs {
      display: flex; gap: 0.35rem; flex-wrap: wrap; margin-bottom: 1.5rem;
      background: #0e1422; padding: 0.35rem; border-radius: 9px;
      border: 1px solid var(--border);
    }
    .domain-tab-btn {
      background: transparent; border: 1px solid transparent; color: var(--text-muted);
      padding: 0.4rem 0.85rem; border-radius: 6px; font-size: 0.8rem; font-weight: 600;
      cursor: pointer; transition: all 0.15s ease; display: flex; align-items: center; gap: 0.4rem;
    }
    .domain-tab-btn:hover { color: var(--text); background: rgba(255, 255, 255, 0.03); }
    .domain-tab-btn.active {
      background: #162238; border-color: #26385a; color: #60a5fa;
    }

    .glass-panel {
      background: var(--bg-card); backdrop-filter: blur(12px);
      border: 1px solid var(--border); border-radius: 12px;
      padding: 1.75rem; margin-bottom: 2rem;
    }
    .panel-header {
      display: flex; justify-content: space-between; align-items: center;
      margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px solid var(--border);
    }
    .panel-title { font-family: var(--font-display); font-size: 1.125rem; font-weight: 700; display: flex; align-items: center; gap: 0.6rem; color: #f8fafc; }

    .btn {
      background: #162032; color: #f8fafc; border: 1px solid #283752;
      padding: 0.55rem 1.1rem; border-radius: 7px; font-size: 0.825rem;
      font-weight: 600; cursor: pointer; display: inline-flex; align-items: center;
      gap: 0.5rem; transition: all 0.15s ease;
    }
    .btn:hover { background: #202d44; border-color: #3b5074; color: #ffffff; }
    .btn-secondary { background: rgba(255, 255, 255, 0.03); color: var(--text-muted); border-color: var(--border); }
    .btn-secondary:hover { background: rgba(255, 255, 255, 0.06); border-color: #283752; color: var(--text); }
    .btn-success { background: rgba(52, 211, 153, 0.1); border-color: rgba(52, 211, 153, 0.25); color: #6ee7b7; }
    .btn-success:hover { background: rgba(52, 211, 153, 0.18); border-color: rgba(52, 211, 153, 0.4); color: #a7f3d0; }

    .dropzone {
      border: 2px dashed #223049; border-radius: 12px;
      padding: 3.5rem 2rem; text-align: center; background: #0b101c;
      cursor: pointer; transition: all 0.2s ease; display: flex; flex-direction: column;
      align-items: center; justify-content: center; gap: 1rem;
    }
    .dropzone:hover { border-color: #3b82f6; background: rgba(37, 99, 235, 0.03); }
    .dropzone-icon {
      width: 52px; height: 52px; border-radius: 12px;
      background: #131c2e; border: 1px solid #1f2d48;
      display: flex; align-items: center; justify-content: center; color: #60a5fa;
    }
    .dropzone-title { font-family: var(--font-display); font-size: 1.1rem; font-weight: 700; color: #f8fafc; }
    .dropzone-desc { font-size: 0.825rem; color: var(--text-muted); max-width: 480px; }

    .form-row { display: flex; gap: 1rem; flex-wrap: wrap; margin-bottom: 1.25rem; }
    .form-group { flex: 1; min-width: 200px; display: flex; flex-direction: column; gap: 0.4rem; }
    label { font-size: 0.75rem; font-weight: 600; color: var(--text-muted); }
    input, select {
      background: #0a0f19; border: 1px solid var(--border);
      color: var(--text); padding: 0.55rem 0.8rem; border-radius: 6px;
      font-size: 0.825rem; font-family: var(--font-body); outline: none;
    }
    input:focus, select:focus { border-color: #3b82f6; box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.15); }

    .table-container { overflow-x: auto; border: 1px solid var(--border); border-radius: 9px; background: #0a0f19; }
    table { width: 100%; border-collapse: collapse; text-align: left; font-size: 0.8125rem; font-variant-numeric: tabular-nums; }
    th {
      background: #0e1524; padding: 0.75rem 0.95rem; font-weight: 700;
      color: var(--text-muted); text-transform: uppercase; font-size: 0.7rem; border-bottom: 1px solid var(--border);
      white-space: nowrap; letter-spacing: 0.04em;
    }
    td { padding: 0.65rem 0.95rem; border-bottom: 1px solid #141c2c; white-space: nowrap; }
    tr:nth-child(even) td { background: rgba(255, 255, 255, 0.01); }
    tr:hover td { background: rgba(255, 255, 255, 0.025); }

    .pill { display: inline-block; padding: 0.2rem 0.5rem; border-radius: 5px; font-size: 0.7rem; font-weight: 600; }
    .pill-teaching { background: rgba(56, 189, 248, 0.08); color: #93c5fd; border: 1px solid rgba(56, 189, 248, 0.2); }
    .pill-teaching-id { background: rgba(99, 102, 241, 0.08); color: #a5b4fc; border: 1px solid rgba(99, 102, 241, 0.2); }
    .pill-non-teaching { background: rgba(148, 163, 184, 0.08); color: #cbd5e1; border: 1px solid rgba(148, 163, 184, 0.2); }
    .pill-support-staff { background: rgba(217, 119, 6, 0.08); color: #fde047; border: 1px solid rgba(217, 119, 6, 0.2); }
    .pill-admission { background: rgba(129, 140, 248, 0.08); color: #c7d2fe; border: 1px solid rgba(129, 140, 248, 0.2); }
    .pill-management { background: rgba(241, 245, 249, 0.07); color: #f8fafc; border: 1px solid rgba(241, 245, 249, 0.22); }
    .pill-sbf-facility { background: rgba(20, 184, 166, 0.08); color: #5eead4; border: 1px solid rgba(20, 184, 166, 0.2); }
    .pill-hostel--mess { background: rgba(234, 88, 12, 0.08); color: #fed7aa; border: 1px solid rgba(234, 88, 12, 0.2); }

    .inline-edit-input {
      width: 52px; background: #0e1422; border: 1px solid var(--border);
      color: #f8fafc; padding: 0.25rem 0.35rem; border-radius: 4px; text-align: center; font-weight: 600;
      font-family: var(--font-mono); font-size: 0.8rem;
    }
    .inline-edit-input:focus { border-color: #3b82f6; }

    .modal-overlay {
      position: fixed; top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(3, 6, 12, 0.78); backdrop-filter: blur(8px);
      display: none; align-items: center; justify-content: center; z-index: 100;
    }
    .modal-overlay.active { display: flex; animation: fadeIn 0.18s ease; }
    .modal-content {
      background: var(--bg-surface); border: 1px solid var(--border);
      border-radius: 14px; width: 90%; max-width: 600px; max-height: 90vh;
      overflow-y: auto; padding: 1.75rem; box-shadow: 0 24px 48px -12px rgba(0, 0, 0, 0.75);
    }
  </style>
</head>
<body>

  <header>
    <div class="brand">
      <div class="logo-icon"><i data-lucide="layers"></i></div>
      <div>
        <div class="brand-title">RVS / SVCET Domain & Department Payroll Engine</div>
        <div class="brand-subtitle">Automated Biometric Processing & Domain-Separated Salary Bills</div>
      </div>
    </div>

    <nav>
      <button class="nav-btn active" onclick="switchTab('dashboard')"><i data-lucide="layout-dashboard"></i> Executive Summary</button>
      <button class="nav-btn" onclick="switchTab('biometric')"><i data-lucide="file-up"></i> Upload Biometric</button>
      <button class="nav-btn" onclick="switchTab('attendance')"><i data-lucide="calendar-check"></i> Attendance by Domain</button>
      <button class="nav-btn" onclick="switchTab('payroll')"><i data-lucide="receipt"></i> Salary Registry</button>
      <button class="nav-btn" onclick="switchTab('employees')"><i data-lucide="users"></i> Employee Master</button>
    </nav>

    <div class="header-actions">
      <div id="supabaseBadge" class="badge-supabase offline" onclick="openSupabaseModal()">
        <i data-lucide="database"></i> <span id="supabaseText">Supabase Sync</span>
      </div>
    </div>
  </header>

  <div class="container">

    <!-- TAB 1: DASHBOARD / EXECUTIVE SUMMARY -->
    <div id="tab-dashboard" class="tab-content active">
      <!-- Historical Month Switcher Bar -->
      <div style="background: rgba(14, 165, 233, 0.05); border: 1px solid rgba(14, 165, 233, 0.2); border-radius: 12px; padding: 0.9rem 1.25rem; margin-bottom: 1.5rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;">
        <div style="display: flex; align-items: center; gap: 0.75rem;">
          <div style="background: rgba(14, 165, 233, 0.15); width: 36px; height: 36px; border-radius: 8px; display: flex; align-items: center; justify-content: center; color: #38bdf8;">
            <i data-lucide="history" style="width: 20px;"></i>
          </div>
          <div>
            <div style="font-weight: 700; font-size: 0.95rem; color: #f8fafc;">Historical Payroll Month Viewer</div>
            <div style="font-size: 0.8rem; color: var(--text-muted);">Switch between past uploaded months stored in Supabase without losing previous data</div>
          </div>
        </div>
        <div style="display: flex; align-items: center; gap: 0.75rem;">
          <label style="font-size: 0.85rem; color: var(--text-muted); margin: 0;">Select Month:</label>
          <select id="savedMonthsSelect" style="width: auto; min-width: 220px; padding: 0.4rem 0.8rem;" onchange="loadSavedPayrollMonth(this.value)">
            <option value="">-- Select Saved Month --</option>
          </select>
          <button class="btn btn-secondary" style="padding: 0.4rem 0.75rem;" onclick="loadSavedMonthsDropdown()"><i data-lucide="refresh-cw"></i></button>
        </div>
      </div>

      <div class="metrics-grid">
        <div class="metric-card">
          <div class="metric-header">
            <div class="metric-label">Master Registered Staff</div>
            <div class="metric-icon"><i data-lucide="users"></i></div>
          </div>
          <div id="statTotalEmployees" class="metric-value">1156</div>
          <div class="metric-desc">Total cataloged university personnel</div>
        </div>

        <div class="metric-card">
          <div class="metric-header">
            <div class="metric-label" id="statMonthLabel">Selected Month Active Staff</div>
            <div class="metric-icon"><i data-lucide="user-check"></i></div>
          </div>
          <div id="statMonthActiveStaff" class="metric-value">--</div>
          <div class="metric-desc" id="statMonthDesc">Processed in current payroll run</div>
        </div>

        <div class="metric-card">
          <div class="metric-header">
            <div class="metric-label">Total Net Payout</div>
            <div class="metric-icon"><i data-lucide="banknote"></i></div>
          </div>
          <div id="statMonthNet" class="metric-value" style="color: #4ade80;">--</div>
          <div class="metric-desc">Take-home salary disbursal</div>
        </div>

        <div class="metric-card">
          <div class="metric-header">
            <div class="metric-label">Total Monthly Gross</div>
            <div class="metric-icon"><i data-lucide="indian-rupee"></i></div>
          </div>
          <div id="statMonthGross" class="metric-value">--</div>
          <div class="metric-desc" id="statMonthDeductions">Deductions: ₹0</div>
        </div>
      </div>

      <!-- Domain Cards Grid -->
      <div class="glass-panel" style="margin-bottom: 1.5rem;">
        <div class="panel-header">
          <div class="panel-title"><i data-lucide="pie-chart"></i> Domain-Wise Payroll & Staff Distribution</div>
        </div>
        <div id="domainSummaryGrid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 1rem;">
          <!-- Injected via JavaScript -->
        </div>
      </div>

      <!-- Saved Monthly Runs History Table -->
      <div class="glass-panel">
        <div class="panel-header">
          <div class="panel-title"><i data-lucide="archive"></i> Saved Monthly Payroll Archives (Supabase Cloud)</div>
          <button class="btn btn-secondary" style="padding: 0.35rem 0.75rem; font-size: 0.8rem;" onclick="loadSavedMonthsDropdown()"><i data-lucide="refresh-cw"></i> Refresh Archives</button>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Payroll Month</th>
                <th style="text-align: center;">Staff Count</th>
                <th>Total Gross</th>
                <th>Total Deductions</th>
                <th>Total Net Payout</th>
                <th>Saved On</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody id="savedRunsTbody">
              <tr><td colspan="7" style="text-align: center; color: var(--text-muted); padding: 1.5rem;">Loading past saved months from Supabase Cloud...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- TAB 2: BIOMETRIC UPLOAD -->
    <div id="tab-biometric" class="tab-content">
      <div class="glass-panel">
        <div class="panel-header">
          <div class="panel-title"><i data-lucide="file-up"></i> Upload Raw Biometric Machine Report</div>
        </div>

        <div class="form-row">
          <div class="form-group">
            <label>Select Payroll Month</label>
            <input type="text" id="payrollMonth" value="August 2026">
          </div>
          <div class="form-group">
            <label>Total Days in Month</label>
            <input type="number" id="monthDays" value="31">
          </div>
        </div>

        <div class="dropzone" id="dropzone" onclick="document.getElementById('biometricFileInput').click()">
          <input type="file" id="biometricFileInput" style="display: none;" accept=".xls,.xlsx" onchange="handleFileSelect(event)">
          <div class="dropzone-icon">
            <i data-lucide="upload" style="width: 32px; height: 32px;"></i>
          </div>
          <div class="dropzone-title">Click or Drag & Drop Raw Biometric File Here</div>
          <div class="dropzone-desc">Supports multi-sheet machine raw exports like <code>Daily Attendance Report 01-09-2026.xls</code></div>
        </div>

        <div id="uploadStatus" style="margin-top: 1.5rem; display: none;"></div>
      </div>
    </div>

    <!-- TAB 3: ATTENDANCE BY DOMAIN -->
    <div id="tab-attendance" class="tab-content">
      <div class="glass-panel">
        <div class="panel-header">
          <div class="panel-title"><i data-lucide="calendar-check"></i> Monthly Attendance & On-Duty (OD) by Domain</div>
          <div style="display: flex; gap: 0.75rem;">
            <button class="btn btn-secondary" onclick="recalculatePayroll()"><i data-lucide="refresh-cw"></i> Recalculate</button>
            <button class="btn" onclick="switchTab('payroll')"><i data-lucide="arrow-right"></i> View Salary Bills</button>
          </div>
        </div>

        <!-- Domain Selection Tabs -->
        <div class="domain-tabs" id="attendanceDomainTabs">
          <button class="domain-tab-btn active" onclick="setAttendanceDomainFilter('ALL')">All Domains</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('Teaching')">Teaching</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('Teaching ID')">Teaching ID</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('Non-Teaching')">Non-Teaching</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('Admission')">Admission</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('Support Staff')">Support Staff</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('Management')">Management</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('SBF Facility')">SBF Facility</button>
          <button class="domain-tab-btn" onclick="setAttendanceDomainFilter('Hostel & Mess')">Hostel & Mess</button>
        </div>

        <div class="form-row">
          <div class="form-group" style="flex: 2;">
            <input type="text" id="attendanceSearch" placeholder="Search by Employee ID, Name or Department..." oninput="filterAttendanceTable()">
          </div>
          <div class="form-group">
            <select id="attendanceDeptFilter" onchange="filterAttendanceTable()">
              <option value="ALL">All Departments</option>
            </select>
          </div>
        </div>

        <div class="table-container">
          <table id="attendanceTable">
            <thead>
              <tr>
                <th>Emp ID</th>
                <th>Employee Name</th>
                <th>Domain</th>
                <th>Department</th>
                <th>Biometric Days</th>
                <th>Holiday</th>
                <th>Leaves (CL/SL)</th>
                <th>SV / OD (On Duty)</th>
                <th>Total Pay Days</th>
                <th>Remarks</th>
              </tr>
            </thead>
            <tbody id="attendanceTbody">
              <tr><td colspan="10" style="text-align: center; color: var(--text-muted); padding: 2rem;">Please upload biometric file to load attendance records.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- TAB 4: SALARY REGISTRY BY DOMAIN -->
    <div id="tab-payroll" class="tab-content">
      <div class="glass-panel">
        <div class="panel-header">
          <div class="panel-title"><i data-lucide="receipt"></i> Domain & Department-Separated University Salary Registry</div>
          <div style="display: flex; gap: 0.6rem; align-items: center; flex-wrap: wrap;">
            <button class="btn btn-success" id="btnExportDomain" onclick="downloadDomainExcel()"><i data-lucide="download"></i> <span id="exportDomainBtnText">Export Master Salary Bill (.xlsx)</span></button>
            <button class="btn btn-secondary" onclick="downloadMasterExcel()"><i data-lucide="file-spreadsheet"></i> Full Master Workbook</button>
            <button class="btn btn-secondary" onclick="savePayrollRun()"><i data-lucide="save"></i> Save Run</button>
          </div>
        </div>

        <!-- Domain Selection Tabs -->
        <div class="domain-tabs" id="payrollDomainTabs">
          <button class="domain-tab-btn active" onclick="setPayrollDomainFilter('ALL')">All Domains</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('Teaching')">Teaching</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('Teaching ID')">Teaching ID</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('Non-Teaching')">Non-Teaching</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('Admission')">Admission</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('Support Staff')">Support Staff</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('Management')">Management</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('SBF Facility')">SBF Facility</button>
          <button class="domain-tab-btn" onclick="setPayrollDomainFilter('Hostel & Mess')">Hostel & Mess</button>
        </div>

        <div class="metrics-grid">
          <div class="metric-card">
            <div class="metric-label">Staff in Selected View</div>
            <div id="payrollTotalStaff" class="metric-value">--</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">Total Gross Salary</div>
            <div id="payrollTotalGross" class="metric-value">₹0</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">Total Deductions</div>
            <div id="payrollTotalDeductions" class="metric-value">₹0</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">Net Payout Amount</div>
            <div id="payrollTotalNet" class="metric-value" style="color: var(--success);">₹0</div>
          </div>
        </div>

        <div class="form-row">
          <div class="form-group" style="flex: 2;">
            <input type="text" id="payrollSearch" placeholder="Filter by Name, ID, Designation, Bank..." oninput="filterPayrollTable()">
          </div>
          <div class="form-group">
            <select id="payrollDeptFilter" onchange="filterPayrollTable()">
              <option value="ALL">All Departments</option>
            </select>
          </div>
        </div>

        <div class="table-container">
          <table id="payrollTable">
            <thead>
              <tr>
                <th>Emp ID</th>
                <th>Name</th>
                <th>Domain</th>
                <th>Department</th>
                <th>Designation</th>
                <th>Standard CTC</th>
                <th>Pay Days</th>
                <th>Basic+AGP</th>
                <th>DA</th>
                <th>HRA</th>
                <th>Gross Total</th>
                <th>PT</th>
                <th>WF</th>
                <th>EPF</th>
                <th>Mess/Other</th>
                <th>Tot Ded</th>
                <th>Net Salary</th>
                <th>Bank Account</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody id="payrollTbody">
              <tr><td colspan="19" style="text-align: center; color: var(--text-muted); padding: 2rem;">No payroll computed yet. Upload biometric data to begin.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- TAB 5: EMPLOYEES MASTER BY DOMAIN -->
    <div id="tab-employees" class="tab-content">
      <div class="glass-panel">
        <div class="panel-header">
          <div class="panel-title"><i data-lucide="users"></i> Master Employee Database</div>
          <button class="btn" onclick="openAddEmployeeModal()"><i data-lucide="user-plus"></i> Add New Employee</button>
        </div>

        <div class="domain-tabs" id="empMasterDomainTabs">
          <button class="domain-tab-btn active" onclick="setEmpMasterDomainFilter('ALL')">All Domains</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('Teaching')">Teaching</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('Teaching ID')">Teaching ID</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('Non-Teaching')">Non-Teaching</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('Admission')">Admission</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('Support Staff')">Support Staff</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('Management')">Management</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('SBF Facility')">SBF Facility</button>
          <button class="domain-tab-btn" onclick="setEmpMasterDomainFilter('Hostel & Mess')">Hostel & Mess</button>
        </div>

        <div class="form-row">
          <div class="form-group" style="flex: 2;">
            <input type="text" id="empMasterSearch" placeholder="Search employee ID, name, bank account..." oninput="loadEmployees()">
          </div>
          <div class="form-group">
            <select id="empDeptFilter" onchange="loadEmployees()">
              <option value="ALL">All Departments</option>
            </select>
          </div>
        </div>

        <div class="table-container">
          <table id="employeesTable">
            <thead>
              <tr>
                <th>Emp ID</th>
                <th>Name</th>
                <th>Domain</th>
                <th>Department</th>
                <th>Designation</th>
                <th>Standard Salary</th>
                <th>PNB Account No</th>
                <th>IFSC Code</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody id="employeesTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

  </div>

  <!-- Supabase Settings Modal -->
  <div id="supabaseModal" class="modal-overlay" onclick="closeModalOnBg(event, 'supabaseModal')">
    <div class="modal-content">
      <div class="panel-header">
        <div class="panel-title"><i data-lucide="database"></i> Supabase Master DB Settings</div>
        <button class="btn btn-secondary" onclick="closeModal('supabaseModal')">&times;</button>
      </div>
      <p style="font-size: 0.825rem; color: var(--text-muted); margin-bottom: 1.25rem;">
        Connect your Supabase project to automatically sync employee profiles, master salaries, and monthly payroll runs across all domains.
      </p>
      <div class="form-group" style="margin-bottom: 1rem;">
        <label>Supabase Project URL</label>
        <input type="text" id="supabaseUrlInput" placeholder="https://your-project.supabase.co">
      </div>
      <div class="form-group" style="margin-bottom: 1.5rem;">
        <label>Supabase Anon / Service API Key</label>
        <input type="password" id="supabaseKeyInput" placeholder="eyJhbGciOiJIUzI1NiIsInR5cCI6...">
      </div>
      <div style="display: flex; gap: 0.75rem; justify-content: flex-end;">
        <button class="btn btn-secondary" onclick="syncSupabase('pull')"><i data-lucide="download-cloud"></i> Pull from Cloud</button>
        <button class="btn btn-secondary" onclick="syncSupabase('push')"><i data-lucide="upload-cloud"></i> Push Local to Cloud</button>
        <button class="btn" onclick="saveSupabaseConfig()"><i data-lucide="check"></i> Test & Save</button>
      </div>
    </div>
  </div>

  <!-- Add/Edit Employee Modal -->
  <div id="employeeModal" class="modal-overlay" onclick="closeModalOnBg(event, 'employeeModal')">
    <div class="modal-content">
      <div class="panel-header">
        <div class="panel-title" id="empModalTitle"><i data-lucide="user"></i> Employee Profile</div>
        <button class="btn btn-secondary" onclick="closeModal('employeeModal')">&times;</button>
      </div>
      <form id="empForm" onsubmit="saveEmployeeForm(event)">
        <div class="form-row">
          <div class="form-group">
            <label>Employee ID (Primary Key)</label>
            <input type="text" id="formEmpId" required>
          </div>
          <div class="form-group">
            <label>Full Name</label>
            <input type="text" id="formEmpName" required>
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>Domain</label>
            <select id="formDomain">
              <option value="Teaching">Teaching</option>
              <option value="Teaching ID">Teaching ID</option>
              <option value="Non-Teaching">Non-Teaching</option>
              <option value="Admission">Admission</option>
              <option value="Support Staff">Support Staff</option>
              <option value="Management">Management</option>
              <option value="SBF Facility">SBF Facility</option>
              <option value="Hostel & Mess">Hostel & Mess</option>
            </select>
          </div>
          <div class="form-group">
            <label>Department</label>
            <input type="text" id="formDepartment" placeholder="e.g. CSE, CIVIL, ADMIN">
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>Designation</label>
            <input type="text" id="formDesignation" placeholder="e.g. Assoc.Prof & HOD">
          </div>
          <div class="form-group">
            <label>Standard Total Salary (CTC)</label>
            <input type="number" id="formSalary" step="0.01" required>
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label>PNB Account Number</label>
            <input type="text" id="formAccountNo">
          </div>
          <div class="form-group">
            <label>IFSC Code</label>
            <input type="text" id="formIfsc" value="PUNB0401700">
          </div>
        </div>
        <div style="display: flex; gap: 0.75rem; justify-content: flex-end; margin-top: 1.5rem;">
          <button type="button" class="btn btn-secondary" onclick="closeModal('employeeModal')">Cancel</button>
          <button type="submit" class="btn"><i data-lucide="save"></i> Save Employee</button>
        </div>
      </form>
    </div>
  </div>

  <script>
    let currentAttendance = {};
    let currentDeductions = {};
    let currentPayrollSummary = null;

    let selectedAttendanceDomain = 'ALL';
    let selectedPayrollDomain = 'ALL';
    let selectedEmpMasterDomain = 'ALL';

    document.addEventListener('DOMContentLoaded', () => {
      lucide.createIcons();
      loadStats();
      loadEmployees();
      loadSavedMonthsDropdown();
    });

    function switchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

      const targetTab = document.getElementById('tab-' + tabId);
      if (targetTab) targetTab.classList.add('active');

      const btn = Array.from(document.querySelectorAll('.nav-btn')).find(b => b.getAttribute('onclick')?.includes(tabId));
      if (btn) btn.classList.add('active');
      lucide.createIcons();
    }

    async function loadStats() {
      try {
        const res = await fetch('/api/stats');
        const data = await res.json();
        document.getElementById('statTotalEmployees').innerText = data.total_employees || 1156;
        
        // If recent runs exist, show latest run in the top month metrics
        if (data.recent_runs && data.recent_runs.length > 0) {
          const latest = data.recent_runs[0];
          document.getElementById('statMonthActiveStaff').innerText = latest.total_employees;
          document.getElementById('statMonthNet').innerText = '₹' + Math.round(latest.total_net || 0).toLocaleString('en-IN');
          document.getElementById('statMonthGross').innerText = '₹' + Math.round(latest.total_gross || 0).toLocaleString('en-IN');
          document.getElementById('statMonthDeductions').innerText = `Deductions: ₹${Math.round(latest.total_deductions || 0).toLocaleString('en-IN')}`;
          document.getElementById('statMonthLabel').innerText = `Active for ${latest.month_year}`;
          document.getElementById('statMonthDesc').innerText = `Archived in Supabase Cloud Database`;
        } else {
          document.getElementById('statMonthActiveStaff').innerText = data.total_employees || 1156;
          document.getElementById('statMonthNet').innerText = '₹' + Math.round(data.total_base_payroll || 0).toLocaleString('en-IN');
          document.getElementById('statMonthGross').innerText = '₹' + Math.round(data.total_base_payroll || 0).toLocaleString('en-IN');
          document.getElementById('statMonthDeductions').innerText = `Standard Base Commitment`;
        }
        
        // Render Domain Summary Cards
        const grid = document.getElementById('domainSummaryGrid');
        grid.innerHTML = '';
        const domainIcons = {
          'Teaching': 'graduation-cap',
          'Teaching ID': 'book-open',
          'Non-Teaching': 'briefcase',
          'Admission': 'user-check',
          'Support Staff': 'shield',
          'Management': 'crown',
          'SBF Facility': 'sparkles',
          'Hostel & Mess': 'utensils'
        };

        for (const [dom, info] of Object.entries(data.domains)) {
          const icon = domainIcons[dom] || 'folder';
          const card = document.createElement('div');
          card.style = 'background: rgba(255,255,255,0.03); border: 1px solid var(--border); padding: 1.15rem; border-radius: 12px; cursor: pointer;';
          card.onclick = () => {
            setPayrollDomainFilter(dom);
            switchTab('payroll');
          };
          card.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.5rem;">
              <span class="pill pill-${dom.toLowerCase().replace(/[^a-z]/g, '-')}">${dom}</span>
              <i data-lucide="${icon}" style="width: 18px; color: var(--text-muted);"></i>
            </div>
            <div style="font-size: 1.35rem; font-weight: 700; color: var(--text);">${info.count} <span style="font-size: 0.8rem; font-weight: 400; color: var(--text-muted);">Staff</span></div>
            <div style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.2rem;">Base: ₹${Math.round(info.base_salary).toLocaleString('en-IN')}</div>
          `;
          grid.appendChild(card);
        }

        // Populate Dept dropdowns
        const deptOptions = data.departments.map(d => `<option value="${d.department}">${d.department} (${d.domain})</option>`).join('');
        document.getElementById('attendanceDeptFilter').innerHTML = '<option value="ALL">All Departments</option>' + deptOptions;
        document.getElementById('payrollDeptFilter').innerHTML = '<option value="ALL">All Departments</option>' + deptOptions;
        document.getElementById('empDeptFilter').innerHTML = '<option value="ALL">All Departments</option>' + deptOptions;

        const supBadge = document.getElementById('supabaseBadge');
        if (data.supabase_connected) {
          supBadge.classList.remove('offline');
          document.getElementById('supabaseText').innerText = 'Supabase Connected';
        }
        lucide.createIcons();
      } catch (e) {
        console.error('Stats error', e);
      }
    }

    async function loadSavedMonthsDropdown() {
      try {
        const res = await fetch('/api/payroll-runs');
        const data = await res.json();
        const runs = data.runs || [];
        
        const sel = document.getElementById('savedMonthsSelect');
        if (sel) {
          sel.innerHTML = '<option value="">-- Select Saved Month --</option>' + 
            runs.map(r => `<option value="${r.month_year}">${r.month_year} (${r.total_employees} staff - ₹${Math.round(r.total_net).toLocaleString('en-IN')})</option>`).join('');
        }

        const tbody = document.getElementById('savedRunsTbody');
        if (tbody) {
          if (runs.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; color: var(--text-muted); padding: 1.5rem;">No saved monthly payroll runs found yet. Upload a biometric file to create one!</td></tr>';
          } else {
            tbody.innerHTML = '';
            runs.forEach(r => {
              const tr = document.createElement('tr');
              tr.innerHTML = `
                <td><strong>${r.month_year}</strong></td>
                <td style="text-align: center;"><span class="pill pill-teaching">${r.total_employees} Staff</span></td>
                <td>₹${Math.round(r.total_gross || 0).toLocaleString('en-IN')}</td>
                <td style="color: var(--danger);">₹${Math.round(r.total_deductions || 0).toLocaleString('en-IN')}</td>
                <td style="color: var(--success); font-weight: 700;">₹${Math.round(r.total_net || 0).toLocaleString('en-IN')}</td>
                <td style="color: var(--text-muted); font-size: 0.8rem;">${new Date(r.created_at).toLocaleDateString()}</td>
                <td>
                  <button class="btn btn-secondary" style="padding: 0.3rem 0.6rem; font-size: 0.75rem;" onclick="loadSavedPayrollMonth('${r.month_year}', true)"><i data-lucide="eye"></i> View</button>
                  <button class="btn btn-success" style="padding: 0.3rem 0.6rem; font-size: 0.75rem;" onclick="downloadSpecificMonthExcel('${r.month_year}')"><i data-lucide="download"></i> Excel</button>
                </td>
              `;
              tbody.appendChild(tr);
            });

            // Automatically pre-load the latest month into memory if not loaded yet
            if (!currentPayrollSummary && runs.length > 0) {
              await loadSavedPayrollMonth(runs[0].month_year, false);
            }
          }
        }
        lucide.createIcons();
      } catch (e) {
        console.error('Error loading saved months', e);
      }
    }

    async function loadSavedPayrollMonth(monthYear, switchTabToPayroll = true) {
      if (!monthYear) return;
      try {
        const res = await fetch(`/api/payroll-runs/${encodeURIComponent(monthYear)}`);
        const data = await res.json();
        if (res.ok && data.payroll_summary) {
          currentPayrollSummary = data.payroll_summary;
          document.getElementById('payrollMonth').value = monthYear;
          document.getElementById('monthDays').value = data.payroll_summary.month_days || 31;
          
          // Update Month Specific Summary Cards
          document.getElementById('statMonthActiveStaff').innerText = data.payroll_summary.total_employees;
          document.getElementById('statMonthGross').innerText = '₹' + Math.round(data.payroll_summary.total_gross).toLocaleString('en-IN');
          document.getElementById('statMonthNet').innerText = '₹' + Math.round(data.payroll_summary.total_net).toLocaleString('en-IN');
          document.getElementById('statMonthDeductions').innerText = `Deductions: ₹${Math.round(data.payroll_summary.total_deductions).toLocaleString('en-IN')}`;
          document.getElementById('statMonthLabel').innerText = `Active for ${monthYear}`;
          document.getElementById('statMonthDesc').innerText = `Loaded from Supabase Cloud`;

          // Re-populate Attendance Map from Items
          currentAttendance = {};
          (data.payroll_summary.items || []).forEach(it => {
            currentAttendance[it.emp_id] = {
              emp_id: it.emp_id,
              emp_name: it.emp_name,
              domain: it.domain,
              department: it.department,
              biometric_days: it.total_pay_days || 31,
              holiday_days: 0,
              availed_leaves: 0,
              od_days: 0,
              total_pay_days: it.total_pay_days || 31,
              remarks: ''
            };
          });

          renderPayrollTable();
          renderAttendanceTable();
          if (switchTabToPayroll) {
            switchTab('payroll');
          }
        }
      } catch (e) {
        alert('Failed to load month: ' + e.message);
      }
    }

    async function downloadSpecificMonthExcel(monthYear) {
      try {
        const exportRes = await fetch('/api/export-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            month_year: monthYear,
            month_days: 31,
            domain: 'ALL',
            attendance: currentAttendance,
            deductions: currentDeductions
          })
        });
        const blob = await exportRes.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Master_University_Salary_Bill_${monthYear.replace(/ /g, '_')}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
      } catch (e) {
        alert('Error downloading: ' + e.message);
      }
    }

    function setAttendanceDomainFilter(dom) {
      selectedAttendanceDomain = dom;
      document.querySelectorAll('#attendanceDomainTabs .domain-tab-btn').forEach(b => {
        b.classList.toggle('active', b.innerText.trim() === dom || (dom === 'ALL' && b.innerText.includes('All')));
      });
      filterAttendanceTable();
    }

    function setPayrollDomainFilter(dom) {
      selectedPayrollDomain = dom;
      document.querySelectorAll('#payrollDomainTabs .domain-tab-btn').forEach(b => {
        b.classList.toggle('active', b.innerText.trim() === dom || (dom === 'ALL' && b.innerText.includes('All')));
      });
      const btnText = document.getElementById('exportDomainBtnText');
      if (btnText) {
        if (dom === 'ALL') {
          btnText.innerText = 'Export Master Salary Bill (.xlsx)';
        } else {
          btnText.innerText = `Export ${dom} Salary Bill (.xlsx)`;
        }
      }
      filterPayrollTable();
    }

    function setEmpMasterDomainFilter(dom) {
      selectedEmpMasterDomain = dom;
      document.querySelectorAll('#empMasterDomainTabs .domain-tab-btn').forEach(b => {
        b.classList.toggle('active', b.innerText.trim() === dom || (dom === 'ALL' && b.innerText.includes('All')));
      });
      loadEmployees();
    }

    async function loadEmployees() {
      try {
        const q = document.getElementById('empMasterSearch')?.value || '';
        const dept = document.getElementById('empDeptFilter')?.value || 'ALL';
        const res = await fetch(`/api/employees?search=${encodeURIComponent(q)}&domain=${selectedEmpMasterDomain}&department=${dept}`);
        const data = await res.json();
        
        const tbody = document.getElementById('employeesTbody');
        tbody.innerHTML = '';
        data.employees.forEach(emp => {
          const tr = document.createElement('tr');
          const pillClass = `pill-${(emp.domain || 'teaching').toLowerCase().replace(/[^a-z]/g, '-')}`;
          tr.innerHTML = `
            <td><strong>${emp.emp_id}</strong></td>
            <td>${emp.emp_name}</td>
            <td><span class="pill ${pillClass}">${emp.domain || 'General'}</span></td>
            <td>${emp.department || '--'}</td>
            <td>${emp.designation || '--'}</td>
            <td>₹${(emp.standard_salary || 0).toLocaleString('en-IN')}</td>
            <td>${emp.account_no || '--'}</td>
            <td>${emp.ifsc_code || '--'}</td>
            <td>
              <button class="btn btn-secondary" style="padding: 0.3rem 0.6rem; font-size: 0.75rem;" onclick='editEmployee(${JSON.stringify(emp)})'>Edit</button>
            </td>
          `;
          tbody.appendChild(tr);
        });
      } catch (e) {
        console.error('Error loading employees', e);
      }
    }

    async function handleFileSelect(e) {
      const file = e.target.files[0];
      if (!file) return;

      const monthYear = document.getElementById('payrollMonth').value;
      const monthDays = parseInt(document.getElementById('monthDays').value) || 31;

      const formData = new FormData();
      formData.append('file', file);
      formData.append('month_year', monthYear);
      formData.append('month_days', monthDays);

      const statusDiv = document.getElementById('uploadStatus');
      statusDiv.style.display = 'block';
      statusDiv.innerHTML = '<div style="color: var(--primary);"><i data-lucide="loader-2"></i> Parsing punch logs across all departments and domains...</div>';
      lucide.createIcons();

      try {
        const res = await fetch('/api/upload-biometric', { method: 'POST', body: formData });
        let data = {};
        const rawText = await res.text();
        try {
          data = JSON.parse(rawText);
        } catch (pe) {
          if (res.status === 413 || rawText.includes('Request Entity Too Large') || rawText.includes('PAYLOAD_TOO_LARGE')) {
            throw new Error(`File size (${(file.size / (1024*1024)).toFixed(1)}MB) exceeds Vercel's 4.5MB serverless payload limit. Please use your Render deployment (https://rvs-university-attendance.onrender.com) which has no file size limit for large machine exports!`);
          }
          throw new Error(rawText.substring(0, 150) || `Server returned HTTP ${res.status}`);
        }

        if (res.ok) {
          currentAttendance = data.attendance;
          currentPayrollSummary = data.payroll_summary;
          
          // Update month metric cards
          document.getElementById('statMonthActiveStaff').innerText = data.payroll_summary.total_employees;
          document.getElementById('statMonthGross').innerText = '₹' + Math.round(data.payroll_summary.total_gross).toLocaleString('en-IN');
          document.getElementById('statMonthNet').innerText = '₹' + Math.round(data.payroll_summary.total_net).toLocaleString('en-IN');
          document.getElementById('statMonthDeductions').innerText = `Deductions: ₹${Math.round(data.payroll_summary.total_deductions).toLocaleString('en-IN')}`;
          document.getElementById('statMonthLabel').innerText = `Active for ${monthYear}`;
          document.getElementById('statMonthDesc').innerText = `Saved in Supabase Cloud Database`;

          statusDiv.innerHTML = `<div style="color: var(--success);"><i data-lucide="check-circle-2"></i> Successfully parsed & saved ${data.parsed_employees} employees from ${data.filename} to Supabase Cloud!</div>`;
          lucide.createIcons();
          renderAttendanceTable();
          renderPayrollTable();
          loadSavedMonthsDropdown();
          setTimeout(() => switchTab('attendance'), 700);
        } else {
          statusDiv.innerHTML = `<div style="color: var(--danger);"><i data-lucide="alert-triangle"></i> Error: ${data.message}</div>`;
          lucide.createIcons();
        }
      } catch (err) {
        statusDiv.innerHTML = `<div style="color: var(--danger);"><i data-lucide="alert-triangle"></i> Upload failed: ${err.message}</div>`;
        lucide.createIcons();
      }
    }

    function renderAttendanceTable() {
      filterAttendanceTable();
    }

    function filterAttendanceTable() {
      const q = (document.getElementById('attendanceSearch')?.value || '').toLowerCase();
      const dept = document.getElementById('attendanceDeptFilter')?.value || 'ALL';
      const tbody = document.getElementById('attendanceTbody');
      tbody.innerHTML = '';

      const list = Object.values(currentAttendance).filter(att => {
        const dMatch = (selectedAttendanceDomain === 'ALL' || att.domain === selectedAttendanceDomain);
        const deptMatch = (dept === 'ALL' || att.department === dept);
        const searchMatch = !q || (att.emp_id.toLowerCase().includes(q) || att.emp_name.toLowerCase().includes(q) || (att.department || '').toLowerCase().includes(q));
        return dMatch && deptMatch && searchMatch;
      });

      if (list.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; color: var(--text-muted); padding: 2rem;">No matching attendance records found.</td></tr>';
        return;
      }

      list.forEach(att => {
        const tr = document.createElement('tr');
        const pillClass = `pill-${(att.domain || 'teaching').toLowerCase().replace(/[^a-z]/g, '-')}`;
        tr.innerHTML = `
          <td><strong>${att.emp_id}</strong></td>
          <td>${att.emp_name}</td>
          <td><span class="pill ${pillClass}">${att.domain || 'General'}</span></td>
          <td>${att.department || 'General'}</td>
          <td>${att.biometric_days}</td>
          <td>${att.holiday_days}</td>
          <td><input type="number" step="0.5" class="inline-edit-input" value="${att.availed_leaves}" onchange="updateAttendance('${att.emp_id}', 'availed_leaves', this.value)"></td>
          <td><input type="number" step="0.5" class="inline-edit-input" value="${att.od_days}" onchange="updateAttendance('${att.emp_id}', 'od_days', this.value)"></td>
          <td><strong id="paydays-${att.emp_id}" style="color: var(--primary);">${att.total_pay_days}</strong></td>
          <td><input type="text" style="width: 100px; background: rgba(0,0,0,0.2); border:1px solid var(--border); color:white; padding:0.2rem;" value="${att.remarks || ''}" onchange="updateAttendance('${att.emp_id}', 'remarks', this.value)"></td>
        `;
        tbody.appendChild(tr);
      });
    }

    function updateAttendance(empId, field, val) {
      if (currentAttendance[empId]) {
        currentAttendance[empId][field] = (field === 'remarks') ? val : parseFloat(val) || 0;
        
        const mDays = parseInt(document.getElementById('monthDays').value) || 31;
        const b = currentAttendance[empId].biometric_days || 0;
        const h = currentAttendance[empId].holiday_days || 0;
        const l = currentAttendance[empId].availed_leaves || 0;
        const od = currentAttendance[empId].od_days || 0;
        const total = Math.min(mDays, b + h + l + od);
        currentAttendance[empId].total_pay_days = total;
        
        const el = document.getElementById(`paydays-${empId}`);
        if (el) el.innerText = total;
      }
    }

    async function recalculatePayroll() {
      const monthYear = document.getElementById('payrollMonth').value;
      const monthDays = parseInt(document.getElementById('monthDays').value) || 31;

      try {
        const res = await fetch('/api/calculate-payroll', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            month_year: monthYear,
            month_days: monthDays,
            attendance: currentAttendance,
            deductions: currentDeductions
          })
        });
        const data = await res.json();
        if (res.ok) {
          currentPayrollSummary = data.payroll_summary;
          renderPayrollTable();
          alert('Payroll recalculated across all domains successfully!');
        }
      } catch (e) {
        console.error('Recalculate error', e);
      }
    }

    function renderPayrollTable() {
      filterPayrollTable();
    }

    function filterPayrollTable() {
      if (!currentPayrollSummary) return;

      const q = (document.getElementById('payrollSearch')?.value || '').toLowerCase();
      const dept = document.getElementById('payrollDeptFilter')?.value || 'ALL';

      // Update Export Button Label Dynamically
      const btnText = document.getElementById('exportDomainBtnText');
      if (btnText) {
        if (dept && dept !== 'ALL') {
          btnText.innerText = `Export ${dept} Department Salary Bill (.xlsx)`;
        } else if (selectedPayrollDomain && selectedPayrollDomain !== 'ALL') {
          btnText.innerText = `Export ${selectedPayrollDomain} Salary Bill (.xlsx)`;
        } else {
          btnText.innerText = 'Export Master Salary Bill (.xlsx)';
        }
      }

      const filteredItems = currentPayrollSummary.items.filter(it => {
        const dMatch = (selectedPayrollDomain === 'ALL' || it.domain === selectedPayrollDomain);
        const deptMatch = (dept === 'ALL' || it.department === dept);
        const searchMatch = !q || (it.emp_id.toLowerCase().includes(q) || it.emp_name.toLowerCase().includes(q) || (it.designation || '').toLowerCase().includes(q) || (it.account_no || '').includes(q));
        return dMatch && deptMatch && searchMatch;
      });

      let sumGross = 0, sumDed = 0, sumNet = 0;
      filteredItems.forEach(it => {
        sumGross += it.gross_total || 0;
        sumDed += it.tot_ded || 0;
        sumNet += it.net_salary || 0;
      });

      document.getElementById('payrollTotalStaff').innerText = filteredItems.length;
      document.getElementById('payrollTotalGross').innerText = '₹' + Math.round(sumGross).toLocaleString('en-IN');
      document.getElementById('payrollTotalDeductions').innerText = '₹' + Math.round(sumDed).toLocaleString('en-IN');
      document.getElementById('payrollTotalNet').innerText = '₹' + Math.round(sumNet).toLocaleString('en-IN');

      const tbody = document.getElementById('payrollTbody');
      tbody.innerHTML = '';

      if (filteredItems.length === 0) {
        tbody.innerHTML = '<tr><td colspan="19" style="text-align: center; color: var(--text-muted); padding: 2rem;">No matching staff found in this domain/department.</td></tr>';
        return;
      }

      filteredItems.forEach(it => {
        const tr = document.createElement('tr');
        const pillClass = `pill-${(it.domain || 'teaching').toLowerCase().replace(/[^a-z]/g, '-')}`;
        tr.innerHTML = `
          <td><strong>${it.emp_id}</strong></td>
          <td>${it.emp_name}</td>
          <td><span class="pill ${pillClass}">${it.domain || 'General'}</span></td>
          <td>${it.department || 'General'}</td>
          <td>${it.designation || '--'}</td>
          <td>₹${it.standard_salary?.toLocaleString('en-IN')}</td>
          <td><strong>${it.total_pay_days}</strong></td>
          <td>₹${it.basic_agp?.toLocaleString('en-IN')}</td>
          <td>₹${it.da?.toLocaleString('en-IN')}</td>
          <td>₹${it.hra?.toLocaleString('en-IN')}</td>
          <td><strong>₹${it.gross_total?.toLocaleString('en-IN')}</strong></td>
          <td>₹${it.pt}</td>
          <td>₹${it.wf}</td>
          <td>₹${it.epf}</td>
          <td>₹${it.mess + it.eb + it.bus + it.other_ded}</td>
          <td style="color: var(--danger);">₹${it.tot_ded?.toLocaleString('en-IN')}</td>
          <td style="color: var(--success); font-weight: 700;">₹${it.net_salary?.toLocaleString('en-IN')}</td>
          <td>${it.account_no || '--'}</td>
          <td>
            <button class="btn btn-secondary" style="padding: 0.25rem 0.5rem; font-size: 0.75rem;" onclick='viewSlip(${JSON.stringify(it)})'>Slip</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    async function downloadDomainExcel(overrideDomain, overrideDept) {
      if (!currentPayrollSummary) {
        alert('Please upload biometric file first or select a saved month!');
        return;
      }
      const dom = overrideDomain || selectedPayrollDomain || 'ALL';
      const dept = overrideDept || (document.getElementById('payrollDeptFilter') ? document.getElementById('payrollDeptFilter').value : 'ALL');
      const monthYear = document.getElementById('payrollMonth').value;
      const monthDays = parseInt(document.getElementById('monthDays').value) || 31;

      try {
        const res = await fetch('/api/export-excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            month_year: monthYear,
            month_days: monthDays,
            domain: dom,
            department: dept,
            attendance: currentAttendance,
            deductions: currentDeductions
          })
        });
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        let prefix = 'Master_University';
        if (dept && dept !== 'ALL') {
          prefix = `${dept.replace(/ /g, '_').replace(/\//g, '_')}_Department`;
        } else if (dom && dom !== 'ALL') {
          prefix = dom.replace(/ /g, '_');
        }
        a.download = `${prefix}_Salary_Bill_${monthYear.replace(/ /g, '_')}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
      } catch (e) {
        alert('Error downloading Excel: ' + e.message);
      }
    }

    async function downloadMasterExcel() {
      await downloadDomainExcel('ALL', 'ALL');
    }

    async function savePayrollRun() {
      if (!currentPayrollSummary) return;
      const monthYear = document.getElementById('payrollMonth').value;
      const monthDays = parseInt(document.getElementById('monthDays').value) || 31;

      try {
        const res = await fetch('/api/save-payroll', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            month_year: monthYear,
            month_days: monthDays,
            summary: currentPayrollSummary
          })
        });
        const data = await res.json();
        if (res.ok) {
          alert('Payroll Run Saved to Database!');
          loadStats();
        }
      } catch (e) {
        alert('Error saving payroll: ' + e.message);
      }
    }

    function viewSlip(it) {
      alert(`Official Payslip for ${it.emp_name} (ID: ${it.emp_id}):\\n\\nDomain: ${it.domain}\\nDepartment: ${it.department}\\nDesignation: ${it.designation}\\n\\nStandard CTC: ₹${it.standard_salary}\\nTotal Pay Days: ${it.total_pay_days}\\nGross Salary: ₹${it.gross_total}\\nTotal Deductions: ₹${it.tot_ded}\\nNet Payable Salary: ₹${it.net_salary}\\n\\nBank Account: ${it.account_no}\\nIFSC: ${it.ifsc_code}`);
    }

    function openAddEmployeeModal() {
      document.getElementById('formEmpId').value = '';
      document.getElementById('formEmpName').value = '';
      document.getElementById('formDepartment').value = '';
      document.getElementById('formDesignation').value = '';
      document.getElementById('formSalary').value = '';
      document.getElementById('formAccountNo').value = '';
      document.getElementById('employeeModal').classList.add('active');
    }

    function editEmployee(emp) {
      document.getElementById('formEmpId').value = emp.emp_id;
      document.getElementById('formEmpName').value = emp.emp_name;
      document.getElementById('formDomain').value = emp.domain || 'Teaching';
      document.getElementById('formDepartment').value = emp.department || '';
      document.getElementById('formDesignation').value = emp.designation || '';
      document.getElementById('formSalary').value = emp.standard_salary || '';
      document.getElementById('formAccountNo').value = emp.account_no || '';
      document.getElementById('formIfsc').value = emp.ifsc_code || 'PUNB0401700';
      document.getElementById('employeeModal').classList.add('active');
    }

    async function saveEmployeeForm(e) {
      e.preventDefault();
      const payload = {
        emp_id: document.getElementById('formEmpId').value,
        emp_name: document.getElementById('formEmpName').value,
        domain: document.getElementById('formDomain').value,
        category: document.getElementById('formDomain').value,
        department: document.getElementById('formDepartment').value,
        designation: document.getElementById('formDesignation').value,
        standard_salary: parseFloat(document.getElementById('formSalary').value) || 0,
        account_no: document.getElementById('formAccountNo').value,
        ifsc_code: document.getElementById('formIfsc').value
      };

      try {
        const res = await fetch('/api/employees', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        if (res.ok) {
          closeModal('employeeModal');
          loadEmployees();
          loadStats();
        }
      } catch (err) {
        alert('Error saving employee: ' + err.message);
      }
    }

    function openSupabaseModal() {
      document.getElementById('supabaseModal').classList.add('active');
    }

    function closeModal(id) {
      document.getElementById(id).classList.remove('active');
    }

    function closeModalOnBg(e, id) {
      if (e.target.id === id) closeModal(id);
    }

    async function saveSupabaseConfig() {
      const url = document.getElementById('supabaseUrlInput').value;
      const key = document.getElementById('supabaseKeyInput').value;
      const res = await fetch('/api/supabase/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ supabase_url: url, supabase_key: key })
      });
      const data = await res.json();
      alert(data.message);
      if (data.status === 'connected') {
        closeModal('supabaseModal');
        loadStats();
      }
    }

    async function syncSupabase(dir) {
      const res = await fetch(`/api/supabase/sync?direction=${dir}`, { method: 'POST' });
      const data = await res.json();
      alert(data.message);
      loadEmployees();
      loadStats();
    }
  </script>
</body>
</html>
"""

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False)
