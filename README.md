# RVS / SVCET University Biometric Attendance & Payroll Engine

A production-grade, domain-separated biometric attendance processing and statutory payroll calculation engine built for **RVS / Sri Venkateswara College of Engineering & Technology (SVCET)**.

---

## 🌟 Key Features

### 1. 🏢 Multi-Domain & Department Structure
- **Teaching Faculty**: CSE, CIVIL, ECE, EEE, MECH, HAS, MBA, MCA, AI&DS, CS&DS, etc.
- **Non-Teaching Staff**: Administration, Exam Section, Library, Physical Education, Central Office, Placement (TAP Cell).
- **Management & Executive Staff**: Senior Leadership, Accounts, Estate Management.
- **Support & Campus Services**: Transport Drivers & Cleaners, Electricians, Workshop, Attenders, Gardeners, Housekeeping.
- **Campus Auxiliary Units**: SLH / Sharath Mess, SBF, Admission Wing.

### 2. ⚡ Institutional Rules & Shift Policy Engine
- **Principal & Senior Leadership (101, 707, 900, 1060, 1015, 1019, 1021, 4001, Shajahan, 1030, 6623)**: Guaranteed full attendance even if unlisted in punch machines.
- **CSE Dr. Bala Subramanyam (536)**: Grace period — check-in before 12:10 PM awarded full day.
- **Civil M. Lilaakar (109)**: Grace period — check-in before 11:00 AM awarded full day.
- **Media Team Purdvi Raj (1018)**: Shift check-in by 09:35 AM.
- **Electricians**: Early campus shift (08:30 – 16:30).
- **Attenders & Support**: 08:35 – 17:30.
- **Gardeners**: 08:35 – 17:10.
- **Transport Drivers / Cleaners**: Morning campus route + evening departure punch verification.
- **Visiting / Shift Faculty (1053 Pachayapan & 1203 IT HOD)**: 3-day weekly quota ($\ge 12$ days) & 14-day monthly quota converted to full month.
- **Admission Cell**: 6-day weekly coverage.

### 3. 📊 Universal Biometric Parser
- Supports raw machine punch logs (ZKTeco / eSSL / Matrix) and summary sheets.
- Automatic day-by-day punch analysis with in/out grace thresholds.
- Dynamic holiday & Sunday auto-detection.
- Month-length cap (max 28/29/30/31 days) preventing over-counting.

### 4. 💰 Statutory Payroll Calculations
- **UGC Scaled & Consolidated Pay Structures**.
- **Dearness Allowance (DA)**: 37.31%.
- **House Rent Allowance (HRA)**: 16.00%.
- **Andhra Pradesh Professional Tax (PT)** Slabs.
- **Employees' Provident Fund (EPF)** & Staff Welfare Fund (₹75).
- **Custom Deductions**: Electricity (EB), Mess, Bus Fare, Income Tax (TDS).

### 5. 📑 Multi-Sheet Domain Excel Export
- Individual Excel workbooks for specific domains (e.g. `Teaching_Faculty_Salary_Bill.xlsx`).
- Multi-tab workbooks with dedicated sheets for each academic department (`CSE`, `CIVIL`, `ECE`, `EEE`, `MECH`, `HAS`, `MBA`, `MCA`, etc.).
- Native Excel formula subtotals (`=SUM(...)`) and grand totals.

### 6. ☁️ Supabase Cloud Database Integration
- Real-time cloud sync for employee profiles, biometric configurations, and payroll runs.
- Schema definitions with Row Level Security (RLS) policies.
- Automatic batch synchronization from local SQLite database (`payroll_master.db`).

---

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install flask openpyxl xlrd supabase requests
```

### 2. Run the Dashboard
```bash
python server.py
```
Open [http://localhost:5050](http://localhost:5050) in your browser.

### 3. Supabase Cloud Configuration
1. Open your Supabase project: `https://fzdzaowjiapxvfrxrdzc.supabase.co`
2. Run `supabase_schema.sql` in the **SQL Editor**.
3. Sync master data to the cloud:
```bash
python supabase_sync.py
```

---

## 📂 Project Architecture

```
rvs-salary/
│
├── server.py                   # Web dashboard & REST API (Flask)
├── institutional_rules.py      # Custom shift policies, exempt staff & grace periods
├── biometric_parser.py         # Machine punch log parser & attendance evaluator
├── payroll_engine.py           # Statutory salary & deduction math
├── excel_exporter.py           # Domain & department multi-sheet Excel generator
│
├── supabase_sync.py            # Supabase Cloud batch sync client
├── supabase_schema.sql         # Cloud SQL DDL & RLS security policies
│
├── seed_master_db.py           # Master database seeder from official records
├── extract_master_data.py      # University salary template parser
├── payroll_master.db           # Local SQLite database (1,156 verified staff records)
│
└── tests/                      # Verification & rule validation test suite
```

---

## 🔒 Security & Privacy
- Sensitive personal data and API keys should be set via environment variables in production.
- Environment variables:
  - `SUPABASE_URL`: Your Supabase Project URL
  - `SUPABASE_KEY`: Supabase Service Role / Anon Key
  - `PORT`: Web server port (Default: `5050`)
