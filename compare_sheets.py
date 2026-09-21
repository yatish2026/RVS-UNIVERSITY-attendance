import xlrd
from biometric_parser import BiometricParser

# Test 1: Daily Attendance Report
parser1 = BiometricParser('Daily Attendance Report 01-09-2026 (2).xls')
res1 = parser1.parse(31)

print('=== Test 1: Daily Attendance Report ===')
for eid in ['350', '351', '354', '394', '408', '410', '412', '413', '414', '415']:
    if eid in res1:
        e = res1[eid]
        print(f"Emp {eid} ({e['emp_name']}, {e['department']}): Bio={e['biometric_days']}, Hol={e['holiday_days']}, Leaves={e['availed_leaves']}, PayDays={e['total_pay_days']}")
    else:
        print(f"Emp {eid} not found in Daily Attendance Report")

# Test 2: SVCET July 2026 - No. of Days.xls
wb = xlrd.open_workbook('SVCET July 2026 - No. of Days.xls')
sh = wb.sheet_by_name('New')
print('\n=== Test 2: In SVCET July 2026 - No. of Days.xls ===')
for r in range(sh.nrows):
    code = str(sh.cell_value(r, 1)).strip()
    if code.endswith('.0'): code = code[:-2]
    if code in ['350', '351', '354', '394', '408', '410', '412', '413', '414', '415']:
        name = str(sh.cell_value(r, 2)).strip()
        bio = sh.cell_value(r, 4)
        hol = sh.cell_value(r, 5)
        leaves = sh.cell_value(r, 6)
        od = sh.cell_value(r, 7)
        pay = sh.cell_value(r, 8)
        print(f"Emp {code} ({name}): Bio={bio}, Hol={hol}, Leaves={leaves}, OD={od}, PayDays={pay}")
