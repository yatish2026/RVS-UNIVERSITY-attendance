import requests

BASE = 'http://127.0.0.1:5050'

r = requests.get(f'{BASE}/api/stats')
stats = r.json()
print('Total Employees in DB:', stats['total_employees'])
print('Domains found:', list(stats['domains'].keys()))

for dom in ['Teaching', 'Teaching ID', 'Non-Teaching', 'Admission', 'Support Staff', 'Management', 'SBF Facility', 'Hostel & Mess']:
    r = requests.get(f'{BASE}/api/employees?domain={dom}')
    print(f'  Domain [{dom}]: {r.json()["count"]} employees')

with open('Daily Attendance Report 01-09-2026 (2).xls', 'rb') as f:
    files = {'file': ('Daily Attendance Report 01-09-2026 (2).xls', f, 'application/vnd.ms-excel')}
    data = {'month_year': 'August 2026', 'month_days': 31}
    r = requests.post(f'{BASE}/api/upload-biometric', files=files, data=data)
    res = r.json()
    summary = res.get('payroll_summary', {})
    print('\nProcessed Payroll across Domains:')
    for dom, dstat in summary.get('domain_stats', {}).items():
        g = dstat['gross']
        n = dstat['net']
        c = dstat['count']
        print(f'  {dom}: {c} staff | Gross: Rs {g:,.2f} | Net: Rs {n:,.2f}')

r = requests.post(f'{BASE}/api/export-excel', json={'month_year': 'August 2026', 'month_days': 31, 'attendance': res.get('attendance', {}), 'deductions': {}})
print(f'\nExport Excel Status: {r.status_code} | File size: {len(r.content)} bytes')
