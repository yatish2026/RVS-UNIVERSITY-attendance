import sys
from biometric_parser import BiometricParser
from payroll_engine import PayrollEngine
from excel_exporter import ExcelExporter

def run_tests():
    print("=" * 75)
    print("TESTING INSTITUTIONAL RULES ON RAW MACHINE ATTENDANCE LOGS")
    print("=" * 75)
    parser = BiometricParser("Daily Attendance Report 01-09-2026 (2).xls")
    att_map = parser.parse(31)

    engine = PayrollEngine("payroll_master.db")
    summary = engine.process_payroll(att_map, "August 2026", 31)

    print(f"Total processed employees: {summary['total_employees']}")
    print(f"Total Gross: Rs. {summary['total_gross']:,.2f}")
    print(f"Total Net:   Rs. {summary['total_net']:,.2f}")

    special_ids = [
        '101', '536', '707', '109', '1053', '1203', '900', '1060', '1018',
        '1015', '1019', '1021', '4001', 'MGT_1095', '603', '6623', '1030',
        '2005', '206', '243', '625', '626', '627'
    ]

    print("\n" + "-" * 75)
    print("VERIFYING SPECIAL EMPLOYEES & INSTITUTIONAL RULES:")
    print("-" * 75)
    found_map = {it['emp_id']: it for it in summary['items']}
    for sid in special_ids:
        if sid in found_map:
            it = found_map[sid]
            print(f"ID {sid:8s} | {it['emp_name'][:22]:22s} | {it['domain'][:12]:12s} | PayDays: {it['total_pay_days']:4.1f}/31 | Std: Rs.{it['standard_salary']:6.0f} | Gross: Rs.{it['gross_total']:6.0f} | Net: Rs.{it['net_salary']:6.0f} | {it['remarks']}")
        else:
            print(f"ID {sid:8s} | NOT FOUND in output summary")

    # Also test Excel export
    exporter = ExcelExporter()
    out_file = "exports/Institutional_Rules_Test_Salary_Bill.xlsx"
    exporter.export_salary_bill(summary, out_file)
    print(f"\nSuccessfully generated multi-sheet Excel file at {out_file}")

if __name__ == "__main__":
    run_tests()
