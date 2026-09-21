import requests

BASE = 'http://127.0.0.1:5050'

# Test Upload 1: Summary Sheet (SVCET July 2026 - No. of Days.xls)
with open('SVCET July 2026 - No. of Days.xls', 'rb') as f:
    files = {'file': ('SVCET July 2026 - No. of Days.xls', f, 'application/vnd.ms-excel')}
    data = {'month_year': 'July 2026', 'month_days': 31}
    r = requests.post(f'{BASE}/api/upload-biometric', files=files, data=data)
    res = r.json()
    print('Uploaded Summary Sheet (No. of Days.xls):', res.get('parsed_employees'), 'staff parsed')
    att = res.get('attendance', {})
    for eid in ['350', '351', '354', '394', '408', '410', '412', '414', '415']:
        if eid in att:
            e = att[eid]
            print(f"  Emp {eid} ({e['emp_name']}): Bio={e['biometric_days']}, Hol={e['holiday_days']}, Leaves={e['availed_leaves']}, OD={e['od_days']}, PayDays={e['total_pay_days']}")

# Test Upload 2: Raw Machine Log (Daily Attendance Report 01-09-2026 (2).xls)
with open('Daily Attendance Report 01-09-2026 (2).xls', 'rb') as f:
    files = {'file': ('Daily Attendance Report 01-09-2026 (2).xls', f, 'application/vnd.ms-excel')}
    data = {'month_year': 'August 2026', 'month_days': 31}
    r = requests.post(f'{BASE}/api/upload-biometric', files=files, data=data)
    res = r.json()
    print('\nUploaded Raw Machine Sheet (Daily Attendance Report):', res.get('parsed_employees'), 'staff parsed')
    att = res.get('attendance', {})
    for eid in ['561', '509', '1008', '1021', '5001']:
        if eid in att:
            e = att[eid]
            print(f"  Emp {eid} ({e['emp_name']}): Bio={e['biometric_days']}, Hol={e['holiday_days']}, Leaves={e['availed_leaves']}, OD={e['od_days']}, PayDays={e['total_pay_days']}")
