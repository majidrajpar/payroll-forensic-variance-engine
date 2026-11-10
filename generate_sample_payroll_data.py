#!/usr/bin/env python3
"""
Sample Payroll Data Generator for Forensic Variance and Integrity Testing.
Generates Month 01 (prior) and Month 02 (current) payroll registers along with
an authorized HR Lifecycle Movement Log for a 250-employee organization.

Seeded Forensic Test Cases:
1. Two ghost employees sharing an IBAN with an active Finance payroll specialist.
2. One terminated employee still drawing active salary in Month 02 (post-termination leakage).
3. Two employees with unapproved basic salary spikes (>15%) without promotion/adjustment.
4. One record with negative deductions causing Net Pay > Gross Pay.
5. Two managerial employees drawing unauthorized overtime contrary to exempt policy.
6. One employee with an unadjusted round-number manual net pay plug.
"""

import os
import random
from datetime import datetime, date
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Seed for reproducibility
random.seed(42)

DEPARTMENTS = {
    "Executive Management": [
        ("Chief Executive Officer", "EXEC", 55000, 18000, 0, 3500),
        ("Chief Financial Officer", "EXEC", 45000, 15000, 0, 3000),
        ("Chief Operating Officer", "EXEC", 42000, 14000, 0, 2800),
        ("VP Operations", "EXEC", 32000, 10000, 0, 2200),
    ],
    "Finance": [
        ("Finance Manager", "MGR", 22000, 6000, 0, 1500),
        ("Senior Financial Accountant", "SR", 14000, 4000, 0, 950),
        ("Payroll Specialist", "SPEC", 11000, 3500, 0, 750),
        ("Accounts Payable Officer", "SPEC", 9500, 3000, 0, 650),
        ("Treasury Analyst", "SPEC", 10500, 3200, 0, 700),
        ("Junior Accountant", "JR", 7500, 2500, 0, 500),
    ],
    "Operations": [
        ("Operations Manager", "MGR", 24000, 7000, 0, 1600),
        ("Shift Production Supervisor", "SUP", 12500, 3500, 800, 850),
        ("Quality Assurance Lead", "SR", 11500, 3200, 600, 750),
        ("Senior Machine Technician", "TECH", 8500, 2400, 1100, 550),
        ("Line Operator", "OPR", 5200, 1600, 950, 350),
        ("Packaging Specialist", "OPR", 4800, 1500, 850, 320),
        ("Maintenance Technician", "TECH", 7200, 2100, 900, 480),
    ],
    "Logistics": [
        ("Logistics & Fleet Director", "DIR", 28000, 8000, 0, 1900),
        ("Logistics Manager", "MGR", 19500, 5500, 0, 1300),
        ("Warehouse Supervisor", "SUP", 11000, 3000, 750, 750),
        ("Inventory Control Officer", "SPEC", 8500, 2400, 500, 550),
        ("Forklift Operator", "OPR", 4600, 1400, 800, 300),
        ("Dispatch Coordinator", "SPEC", 6800, 2000, 650, 450),
        ("Heavy Fleet Driver", "OPR", 5500, 1800, 1200, 380),
    ],
    "Kitchens": [
        ("Executive Pastry Chef", "EXEC", 21000, 6000, 0, 1400),
        ("Head Production Chef", "MGR", 18500, 5000, 0, 1200),
        ("Sous Chef", "SUP", 11500, 3200, 700, 750),
        ("Station Chef de Partie", "SPEC", 7800, 2200, 850, 520),
        ("Commis Chef", "JR", 4900, 1500, 750, 330),
        ("Kitchen Operations Steward", "OPR", 3800, 1200, 600, 250),
        ("Butchery & Prep Specialist", "SPEC", 6200, 1800, 800, 410),
    ]
}

FIRST_NAMES = [
    "Tariq", "Kareem", "Fatima", "Zaid", "Omar", "Amina", "Rashid", "Layla",
    "Hamza", "Mariam", "Sultan", "Noor", "Bilal", "Salma", "Khalid", "Huda",
    "Fahad", "Reem", "Yousef", "Sara", "Adil", "Samira", "Mustafa", "Zahra",
    "Nasir", "Dina", "Mansoor", "Lina", "Farhan", "Mona", "Jamil", "Yasmin",
    "Imran", "Nadia", "Waleed", "Rania", "Saeed", "Maha", "Hassan", "Aisha",
    "Ibrahim", "Hiba", "Bassam", "Ruba", "Usman", "Lubna", "Qasim", "Farah"
]

LAST_NAMES = [
    "Al-Mansoori", "Hashmi", "Al-Nuaimi", "Siddiqui", "Al-Falasi", "Khan",
    "Al-Zaabi", "Qureshi", "Al-Suwaidi", "Malik", "Al-Shamsi", "Farooqi",
    "Al-Kuwari", "Darwish", "Al-Otaibi", "Rahman", "Al-Dosari", "Abbasi",
    "Al-Ghamdi", "Bhatti", "Al-Mutawa", "Javed", "Al-Husseini", "Mahmoud",
    "Al-Balooshi", "Chaudhry", "Al-Ameri", "Rizvi", "Al-Mazrouei", "Saleem",
    "Al-Dhaheri", "Ansari", "Al-Ketbi", "Wahab", "Al-Qasimi", "Tahir"
]

def generate_iban(index):
    return f"AE{28 + (index % 70):02d}0330000{index + 1000000000:010d}"

def generate_base_population(target_count=250):
    employees = []
    emp_counter = 1001

    # Specific key characters
    # 1. Active Payroll Specialist who will be involved in IBAN collusion (Finance is 1005 - 1018)
    payroll_clerk_id = "EMP-1008"
    payroll_clerk_iban = "AE440330000100800001"

    # Pre-allocate departmental proportions
    dept_weights = {
        "Executive Management": 4,
        "Finance": 14,
        "Operations": 98,
        "Logistics": 64,
        "Kitchens": 70
    }

    used_names = set()
    global_index = 1

    for dept, count in dept_weights.items():
        roles = DEPARTMENTS[dept]
        for i in range(count):
            emp_id = f"EMP-{emp_counter}"
            emp_counter += 1

            # Pick name
            while True:
                fn = random.choice(FIRST_NAMES)
                ln = random.choice(LAST_NAMES)
                name = f"{fn} {ln}"
                if name not in used_names:
                    used_names.add(name)
                    break

            # Pick role
            if dept == "Executive Management":
                role_info = roles[i % len(roles)]
            else:
                # Distribution favoring operators and specialists
                if i < 2:
                    role_info = roles[0]  # Manager
                elif i < 5 and len(roles) > 2:
                    role_info = roles[1]  # Supervisor / Senior
                else:
                    role_info = random.choice(roles[2:])

            title, grade, base_sal, base_allow, base_ot, base_ded = role_info

            # Add subtle natural salary variance (+- 5%)
            jitter = round(random.uniform(-0.04, 0.04), 4)
            basic = round(base_sal * (1.0 + jitter), 2)
            allow = round(base_allow * (1.0 + jitter), 2)

            # OT applies only to non-exempt roles (OPR, TECH, JR, SPEC with base_ot > 0)
            if base_ot > 0:
                ot_hours_val = round(base_ot * random.uniform(0.7, 1.3), 2)
            else:
                ot_hours_val = 0.0

            ded = round(base_ded * (1.0 + jitter * 0.5), 2)

            gross = round(basic + allow + ot_hours_val, 2)
            net = round(gross - ded, 2)

            iban = generate_iban(global_index)
            global_index += 1

            # Tag specific employee for payroll clerk
            if emp_id == payroll_clerk_id:
                title = "Senior Payroll Officer"
                grade = "SPEC"
                iban = payroll_clerk_iban

            employees.append({
                "Employee_ID": emp_id,
                "Employee_Name": name,
                "Department": dept,
                "Designation": title,
                "Grade": grade,
                "Basic_Salary": basic,
                "Allowances": allow,
                "Overtime_Pay": ot_hours_val,
                "Total_Deductions": ded,
                "Gross_Pay": gross,
                "Net_Pay": net,
                "Bank_IBAN": iban,
                "Employment_Status": "Active"
            })

    return employees, payroll_clerk_iban

def apply_forensic_seeds(m01_list, payroll_clerk_iban):
    """
    Produce Month 02 records from Month 01 by applying natural monthly shifts
    and explicitly seeding all 6 forensic red flags.
    """
    m02_list = []
    hr_movement_log = []
    log_id = 5001

    # Map M01 by emp_id
    m01_map = {e["Employee_ID"]: dict(e) for e in m01_list}

    # Selected candidates for forensic injection:
    # Seed 1: 2 ghost records sharing IBAN with active payroll clerk (payroll_clerk_iban)
    # Seed 2: 1 terminated employee still drawing monthly pay
    # Seed 3: 2 employees with unapproved basic salary spike (>15%) without promotion
    # Seed 4: 1 employee with negative deductions (Net > Gross)
    # Seed 5: 2 managerial employees drawing heavy overtime contrary to policy
    # Seed 6: 1 employee with round-number manual net plug

    ghost_ids = ["EMP-1252", "EMP-1253"]  # Newly added fictitious employees in M02
    terminated_leaker_id = "EMP-1045"    # Operations Machine Tech terminated end of M01
    legitimate_leaver_id = "EMP-1088"    # Logistics Driver resigned legitimately end of M01
    spike_emp_1_id = "EMP-1062"          # Operations Line Operator basic spiked
    spike_emp_2_id = "EMP-1115"          # Logistics Warehouse Supervisor basic spiked
    net_exceeds_gross_id = "EMP-1170"    # Kitchen Station Chef negative deduction
    mgr_ot_1_id = "EMP-1019"             # Operations Manager drawing heavy OT
    mgr_ot_2_id = "EMP-1100"             # Logistics Manager drawing heavy OT
    round_plug_id = "EMP-1142"           # Logistics Dispatch Coordinator manual round net plug
    legit_promotion_id = "EMP-1016"      # Legitimate promotion with authorized salary increase
    legit_new_hire_id = "EMP-1251"       # Legitimate approved new hire

    # Step A: Build HR Movement Log with legitimate events
    # 1. Legitimate termination of EMP-1088
    hr_movement_log.append({
        "Log_ID": f"HR-MOV-{log_id}",
        "Employee_ID": legitimate_leaver_id,
        "Employee_Name": m01_map[legitimate_leaver_id]["Employee_Name"],
        "Department": m01_map[legitimate_leaver_id]["Department"],
        "Event_Type": "TERMINATION",
        "Effective_Date": "2026-01-31",
        "Approved_By": "HR Director (H. Al-Falasi)",
        "Old_Value": "Active",
        "New_Value": "Terminated",
        "Notes": "Voluntary resignation. Final settlement cleared."
    })
    log_id += 1

    # 2. Approved termination of EMP-1045 (BUT payroll office fails/refuses to stop payment!)
    hr_movement_log.append({
        "Log_ID": f"HR-MOV-{log_id}",
        "Employee_ID": terminated_leaker_id,
        "Employee_Name": m01_map[terminated_leaker_id]["Employee_Name"],
        "Department": m01_map[terminated_leaker_id]["Department"],
        "Event_Type": "TERMINATION",
        "Effective_Date": "2026-01-25",
        "Approved_By": "Operations VP (F. Al-Zaabi)",
        "Old_Value": "Active",
        "New_Value": "Terminated",
        "Notes": "Terminated for cause following safety disciplinary committee hearing. Do not process Feb payroll."
    })
    log_id += 1

    # 3. Legitimate promotion and salary adjustment for EMP-1016
    hr_movement_log.append({
        "Log_ID": f"HR-MOV-{log_id}",
        "Employee_ID": legit_promotion_id,
        "Employee_Name": m01_map[legit_promotion_id]["Employee_Name"],
        "Department": m01_map[legit_promotion_id]["Department"],
        "Event_Type": "PROMOTION",
        "Effective_Date": "2026-02-01",
        "Approved_By": "Chief Financial Officer (R. Al-Mansoori)",
        "Old_Value": "Junior Accountant - Basic: 7500.00",
        "New_Value": "General Ledger Accountant - Basic: 9800.00",
        "Notes": "Approved annual career progression cycle review."
    })
    log_id += 1

    # 4. Legitimate approved new hire EMP-1250
    hr_movement_log.append({
        "Log_ID": f"HR-MOV-{log_id}",
        "Employee_ID": legit_new_hire_id,
        "Employee_Name": "Zubair Al-Hashimi",
        "Department": "Operations",
        "Event_Type": "NEW_HIRE",
        "Effective_Date": "2026-02-01",
        "Approved_By": "VP Operations (M. Siddiqui)",
        "Old_Value": "N/A",
        "New_Value": "Maintenance Technician - Basic: 7200.00",
        "Notes": "Replacement for departed line maintenance staff."
    })
    log_id += 1

    # Step B: Build Month 02 Register
    for emp_id, rec in m01_map.items():
        # Exclude legitimate leaver EMP-1088 from M02 register
        if emp_id == legitimate_leaver_id:
            continue

        # Create copy for M02
        new_rec = dict(rec)

        # Baseline slight natural OT fluctuation for regular employees
        if new_rec["Grade"] in ["OPR", "TECH", "JR", "SPEC"] and new_rec["Overtime_Pay"] > 0:
            new_rec["Overtime_Pay"] = round(new_rec["Overtime_Pay"] * random.uniform(0.9, 1.15), 2)
            new_rec["Gross_Pay"] = round(new_rec["Basic_Salary"] + new_rec["Allowances"] + new_rec["Overtime_Pay"], 2)
            new_rec["Net_Pay"] = round(new_rec["Gross_Pay"] - new_rec["Total_Deductions"], 2)

        # Inject Seed 2: Terminated employee still drawing salary in M02
        if emp_id == terminated_leaker_id:
            # Still paid active full salary! But in M01 status was Active, now still Active in register
            new_rec["Employment_Status"] = "Active"
            # Leaves amounts identical or slightly modified

        # Inject Legitimate Promotion
        elif emp_id == legit_promotion_id:
            new_rec["Designation"] = "General Ledger Accountant"
            new_rec["Grade"] = "SPEC"
            new_rec["Basic_Salary"] = 9800.00
            new_rec["Allowances"] = 3200.00
            new_rec["Gross_Pay"] = round(new_rec["Basic_Salary"] + new_rec["Allowances"] + new_rec["Overtime_Pay"], 2)
            new_rec["Net_Pay"] = round(new_rec["Gross_Pay"] - new_rec["Total_Deductions"], 2)

        # Inject Seed 3a: Unapproved basic salary spike 1 (EMP-1062 from ~5200 to 7100, +36.5%)
        elif emp_id == spike_emp_1_id:
            orig_basic = new_rec["Basic_Salary"]
            new_rec["Basic_Salary"] = round(orig_basic * 1.35, 2)
            new_rec["Gross_Pay"] = round(new_rec["Basic_Salary"] + new_rec["Allowances"] + new_rec["Overtime_Pay"], 2)
            new_rec["Net_Pay"] = round(new_rec["Gross_Pay"] - new_rec["Total_Deductions"], 2)

        # Inject Seed 3b: Unapproved basic salary spike 2 (EMP-1115 from ~11000 to 13800, +25.5%)
        elif emp_id == spike_emp_2_id:
            orig_basic = new_rec["Basic_Salary"]
            new_rec["Basic_Salary"] = round(orig_basic * 1.255, 2)
            new_rec["Gross_Pay"] = round(new_rec["Basic_Salary"] + new_rec["Allowances"] + new_rec["Overtime_Pay"], 2)
            new_rec["Net_Pay"] = round(new_rec["Gross_Pay"] - new_rec["Total_Deductions"], 2)

        # Inject Seed 4: Negative deductions causing Net > Gross (EMP-1170)
        elif emp_id == net_exceeds_gross_id:
            new_rec["Total_Deductions"] = -1450.00  # Negative deduction plug
            new_rec["Net_Pay"] = round(new_rec["Gross_Pay"] - new_rec["Total_Deductions"], 2) # Net > Gross!

        # Inject Seed 5a: Managerial OT violation 1 (EMP-1019 Operations Manager)
        elif emp_id == mgr_ot_1_id:
            new_rec["Overtime_Pay"] = 4850.00  # Drawing heavy managerial OT
            new_rec["Gross_Pay"] = round(new_rec["Basic_Salary"] + new_rec["Allowances"] + new_rec["Overtime_Pay"], 2)
            new_rec["Net_Pay"] = round(new_rec["Gross_Pay"] - new_rec["Total_Deductions"], 2)

        # Inject Seed 5b: Managerial OT violation 2 (EMP-1100 Logistics Manager)
        elif emp_id == mgr_ot_2_id:
            new_rec["Overtime_Pay"] = 3900.00  # Drawing heavy managerial OT
            new_rec["Gross_Pay"] = round(new_rec["Basic_Salary"] + new_rec["Allowances"] + new_rec["Overtime_Pay"], 2)
            new_rec["Net_Pay"] = round(new_rec["Gross_Pay"] - new_rec["Total_Deductions"], 2)

        # Inject Seed 6: Manual round-number net plug (EMP-1142)
        elif emp_id == round_plug_id:
            # Overwrite Net Pay with a round 10,000.00 without balancing components
            new_rec["Net_Pay"] = 10000.00

        m02_list.append(new_rec)

    # Step C: Add legitimate new hire EMP-1251 to M02
    m02_list.append({
        "Employee_ID": legit_new_hire_id,
        "Employee_Name": "Zubair Al-Hashimi",
        "Department": "Operations",
        "Designation": "Maintenance Technician",
        "Grade": "TECH",
        "Basic_Salary": 7200.00,
        "Allowances": 2100.00,
        "Overtime_Pay": 850.00,
        "Total_Deductions": 480.00,
        "Gross_Pay": 10150.00,
        "Net_Pay": 9670.00,
        "Bank_IBAN": "AE910330000125100001",
        "Employment_Status": "Active"
    })

    # Step D: Inject Seed 1: Ghost records sharing IBAN with payroll clerk
    # Ghost 1: EMP-1252 in Operations
    m02_list.append({
        "Employee_ID": ghost_ids[0],
        "Employee_Name": "Majid Tariq Al-Balooshi",
        "Department": "Operations",
        "Designation": "Senior Machine Technician",
        "Grade": "TECH",
        "Basic_Salary": 8600.00,
        "Allowances": 2400.00,
        "Overtime_Pay": 1150.00,
        "Total_Deductions": 550.00,
        "Gross_Pay": 12150.00,
        "Net_Pay": 11600.00,
        "Bank_IBAN": payroll_clerk_iban,  # COLLUSION IBAN
        "Employment_Status": "Active"
    })

    # Ghost 2: EMP-1253 in Logistics
    m02_list.append({
        "Employee_ID": ghost_ids[1],
        "Employee_Name": "Kareem Farhan Darwish",
        "Department": "Logistics",
        "Designation": "Heavy Fleet Driver",
        "Grade": "OPR",
        "Basic_Salary": 5600.00,
        "Allowances": 1800.00,
        "Overtime_Pay": 1250.00,
        "Total_Deductions": 380.00,
        "Gross_Pay": 8650.00,
        "Net_Pay": 8270.00,
        "Bank_IBAN": payroll_clerk_iban,  # COLLUSION IBAN
        "Employment_Status": "Active"
    })

    return m02_list, hr_movement_log

def style_excel_sheet(ws, title_text):
    # Professional dark navy header styling
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "A2"

    # Style Header Row
    ws.row_dimensions[1].height = 26
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Style Data Rows
    num_cols = ["Basic_Salary", "Allowances", "Overtime_Pay", "Total_Deductions", "Gross_Pay", "Net_Pay"]
    header_names = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 19
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if is_even:
                cell.fill = zebra_fill

            col_name = header_names[col_idx - 1]
            if col_name in num_cols:
                cell.number_format = "$#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif "Date" in str(col_name):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in ["Employee_ID", "Grade", "Employment_Status", "Log_ID", "Event_Type"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

def save_workbook_with_styling(df, filepath, sheet_name="Payroll_Register"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Write headers
    headers = list(df.columns)
    ws.append(headers)

    # Write rows
    for row in df.itertuples(index=False):
        ws.append(list(row))

    style_excel_sheet(ws, sheet_name)
    wb.save(filepath)

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    sample_dir = os.path.join(base_dir, "sample_data")
    os.makedirs(sample_dir, exist_ok=True)

    print("Generating baseline 250-employee workforce...")
    m01_records, payroll_clerk_iban = generate_base_population(target_count=250)
    df_m01 = pd.DataFrame(m01_records)

    print("Simulating Month 02 movements and seeding forensic test cases...")
    m02_records, hr_log = apply_forensic_seeds(m01_records, payroll_clerk_iban)
    df_m02 = pd.DataFrame(m02_records)
    df_hr = pd.DataFrame(hr_log)

    # Output file paths
    m01_xlsx = os.path.join(sample_dir, "Payroll_Register_M01.xlsx")
    m02_xlsx = os.path.join(sample_dir, "Payroll_Register_M02.xlsx")
    hr_xlsx = os.path.join(sample_dir, "HR_Movement_Log.xlsx")

    m01_csv = os.path.join(sample_dir, "Payroll_Register_M01.csv")
    m02_csv = os.path.join(sample_dir, "Payroll_Register_M02.csv")
    hr_csv = os.path.join(sample_dir, "HR_Movement_Log.csv")

    print(f"Writing Excel files to {sample_dir}...")
    save_workbook_with_styling(df_m01, m01_xlsx, "M01_Payroll_Register")
    save_workbook_with_styling(df_m02, m02_xlsx, "M02_Payroll_Register")
    save_workbook_with_styling(df_hr, hr_xlsx, "HR_Lifecycle_Log")

    print(f"Writing CSV companion files...")
    df_m01.to_csv(m01_csv, index=False)
    df_m02.to_csv(m02_csv, index=False)
    df_hr.to_csv(hr_csv, index=False)

    print(f"\nPopulation Summary:")
    print(f"  Month 01 Headcount: {len(df_m01)} employees")
    print(f"  Month 02 Headcount: {len(df_m02)} employees")
    print(f"  HR Movement Log Entries: {len(df_hr)} records")
    print(f"  Seeded Anomalies Injected:")
    print(f"    - Duplicate IBAN Collusion: 2 ghost records (EMP-1252, EMP-1253) matching Payroll Specialist (EMP-1008) IBAN {payroll_clerk_iban}")
    print(f"    - Post-Termination Leakage: 1 employee (EMP-1045) terminated in HR Log but still drawing active pay")
    print(f"    - Unapproved Salary Spikes (>15%): EMP-1062 (+35.0%), EMP-1115 (+25.5%)")
    print(f"    - Net > Gross Violation: EMP-1170 (Negative deduction of -$1,450.00)")
    print(f"    - Managerial Overtime Breaches: Operations Manager (EMP-1019: $4,850) and Logistics Manager (EMP-1100: $3,900)")
    print(f"    - Round-Number Net Pay Plug: EMP-1142 ($10,000.00 manual override)")
    print("Done.")

if __name__ == "__main__":
    main()
