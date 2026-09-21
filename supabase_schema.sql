-- =============================================================================
-- Supabase Schema for RVS / SVCET Attendance & Payroll Engine
-- =============================================================================
-- Run this script in the Supabase SQL Editor (https://supabase.com/dashboard/project/fzdzaowjiapxvfrxrdzc/sql)

-- 1. Employees Master Table
CREATE TABLE IF NOT EXISTS public.employees (
    emp_id TEXT PRIMARY KEY,
    emp_name TEXT NOT NULL,
    domain TEXT DEFAULT 'Teaching',
    category TEXT DEFAULT 'Teaching',
    department TEXT DEFAULT 'General',
    designation TEXT DEFAULT '',
    salary_type TEXT DEFAULT 'REGULAR',
    standard_salary NUMERIC DEFAULT 0.0,
    consolidated_salary NUMERIC DEFAULT 0.0,
    basic NUMERIC DEFAULT 0.0,
    agp NUMERIC DEFAULT 0.0,
    fa NUMERIC DEFAULT 0.0,
    ta_sa NUMERIC DEFAULT 0.0,
    epf_fixed NUMERIC DEFAULT 0.0,
    it_fixed NUMERIC DEFAULT 0.0,
    account_no TEXT DEFAULT '',
    ifsc_code TEXT DEFAULT 'PUNB0401700',
    created_at TIMESTAMPTZ DEFAULT now(),
    updated_at TIMESTAMPTZ DEFAULT now()
);

-- Enable Row Level Security & Policies
ALTER TABLE public.employees ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all access to employees" ON public.employees;
CREATE POLICY "Allow all access to employees" ON public.employees FOR ALL USING (true) WITH CHECK (true);

-- 2. Payroll Runs Table
CREATE TABLE IF NOT EXISTS public.payroll_runs (
    id BIGSERIAL PRIMARY KEY,
    month_year TEXT NOT NULL,
    total_days_in_month INT DEFAULT 31,
    total_employees INT DEFAULT 0,
    total_gross NUMERIC DEFAULT 0.0,
    total_deductions NUMERIC DEFAULT 0.0,
    total_net NUMERIC DEFAULT 0.0,
    created_at TIMESTAMPTZ DEFAULT now()
);

ALTER TABLE public.payroll_runs ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all access to payroll_runs" ON public.payroll_runs;
CREATE POLICY "Allow all access to payroll_runs" ON public.payroll_runs FOR ALL USING (true) WITH CHECK (true);

-- 3. Payroll Items Table (Line-by-line employee monthly calculations)
CREATE TABLE IF NOT EXISTS public.payroll_items (
    id BIGSERIAL PRIMARY KEY,
    run_id BIGINT REFERENCES public.payroll_runs(id) ON DELETE CASCADE,
    emp_id TEXT,
    emp_name TEXT,
    domain TEXT,
    category TEXT,
    department TEXT,
    designation TEXT,
    total_salary NUMERIC DEFAULT 0.0,
    basic NUMERIC DEFAULT 0.0,
    agp NUMERIC DEFAULT 0.0,
    month_days INT DEFAULT 31,
    total_pay_days NUMERIC DEFAULT 0.0,
    basic_agp NUMERIC DEFAULT 0.0,
    da NUMERIC DEFAULT 0.0,
    hra NUMERIC DEFAULT 0.0,
    arrears NUMERIC DEFAULT 0.0,
    fa NUMERIC DEFAULT 0.0,
    ta_sa NUMERIC DEFAULT 0.0,
    gross_total NUMERIC DEFAULT 0.0,
    epf NUMERIC DEFAULT 0.0,
    it NUMERIC DEFAULT 0.0,
    pt NUMERIC DEFAULT 0.0,
    wf NUMERIC DEFAULT 0.0,
    eb NUMERIC DEFAULT 0.0,
    mess NUMERIC DEFAULT 0.0,
    bus NUMERIC DEFAULT 0.0,
    tot_ded NUMERIC DEFAULT 0.0,
    net_salary NUMERIC DEFAULT 0.0,
    account_no TEXT DEFAULT '',
    ifsc_code TEXT DEFAULT 'PUNB0401700',
    created_at TIMESTAMPTZ DEFAULT now()
);

ALTER TABLE public.payroll_items ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all access to payroll_items" ON public.payroll_items;
CREATE POLICY "Allow all access to payroll_items" ON public.payroll_items FOR ALL USING (true) WITH CHECK (true);

-- 4. Publish tables to realtime
ALTER PUBLICATION supabase_realtime ADD TABLE public.employees;
ALTER PUBLICATION supabase_realtime ADD TABLE public.payroll_runs;
