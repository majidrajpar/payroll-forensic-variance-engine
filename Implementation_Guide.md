# Payroll Forensic Variance and Integrity Engine: Implementation & Field Deployment Guide

## 1. Operational Scope and Fiduciary Context

Payroll represents between 25% and 65% of total operating expenses in service, logistics, manufacturing, and hospitality organizations. Despite this scale, routine internal and statutory audits regularly evaluate payroll through small 15-to-25 employee test samples. Such manual testing leaves systemic payroll leakages undetected, including ghost employees, collusion between payroll clerks and line supervisors, duplicate disbursement vouchers, unauthorized basic salary spikes, manual net plugs, and unapproved managerial overtime.

The Payroll Forensic Variance and Integrity Engine executes continuous controls monitoring across 100% of payroll records between any two comparative periods. It ingests prior and current monthly payroll registers along with an optional human resources movement log, standardizes varying column conventions via fuzzy normalization, reconciles headcount movements, quantifies component-level financial variances, and isolates transaction-level red flags before payments clear commercial banking channels.

---

## 2. Technical Architecture and Ingestion Mechanics

The engine executes in Python 3.9+ using standard data engineering libraries. Output workbooks are compiled directly through openpyxl with zero VBA macros, ensuring instant acceptance by enterprise IT security firewalls and corporate endpoint monitoring systems.

### File Format Support
The engine accepts raw register files in Microsoft Excel (.xlsx) and comma-separated values (.csv) formats. Both monthly payroll registers and the HR movement log can be supplied interchangeably in either format.

### Dynamic Schema Mapper
ERP exports from platforms such as SAP, Oracle HCM, Workday, ADP, Microsoft Dynamics 365, and proprietary local payroll packages use differing column names. The engine normalizes column headers using fuzzy string matching to map input files dynamically against canonical fields:

- Employee identification numbers map from emp_id, employee_id, empid, staff_id, personnel_no, payroll_id, badge_id, worker_id, employee_no, or staff_code.
- Employee names map from name, employee_name, staff_name, full_name, person_name, or worker_name.
- Department assignments map from department, dept, cost_center, cost_centre, business_unit, division, function, team, or branch.
- Designations and job grades map from designation, job_title, title, role, position, grade, level, or post.
- Basic salary figures map from basic_pay, basic_salary, base_pay, base_salary, basic, fixed_salary, or monthly_basic.
- Fixed and variable allowances map from allowances, allowance, total_allowances, housing_allowance, transport_allowance, or benefits.
- Overtime payments map from overtime, ot, ot_pay, overtime_pay, ot_amount, or extra_hours_pay.
- Voluntary and statutory deductions map from deductions, total_deductions, deduction, tax_deductions, or withholdings.
- Gross earnings map from gross_pay, gross_salary, gross, total_gross, or total_earnings.
- Net disbursements map from net_pay, net_salary, net, take_home, total_net, net_payable, or transfer_amount.
- Bank account details map from bank_iban, iban, account_number, bank_account, account_no, or bank_acc.
- Employment status indicators map from status, employment_status, term_date, termination_date, exit_date, or work_status.

### Hardened Data Sanitization and Imputation
Payroll files in field environments routinely contain formatting anomalies that cause unhardened scripts to fail:
- Currency formatting: Prefixes such as USD, AED, SAR, GBP, EUR, Dhs, $, £, and € are automatically stripped.
- Number formatting: Comma thousands separators and European decimal commas (e.g. 1.250,50) are normalized into valid floating-point numbers.
- Accounting brackets: Negative values formatted with accounting parentheses (e.g. ($500.00)) or trailing minus signs are converted to standard signed numbers.
- Formula errors: Cell strings containing Excel errors (#N/A, #DIV/0!, #REF!, #VALUE!, #NAME?, #NUM!, #NULL!) or null tokens are imputed as zero for calculations rather than aborting execution.
- Missing columns: Smaller payrolls that omit optional columns (such as Overtime, Allowances, Department, or Designation) automatically receive imputed defaults (0.0 for monetary values, "Unassigned" for departments, "General Staff" for designations).
- Date parsing: The engine parses HR event dates across ISO formats (YYYY-MM-DD), European formats (DD/MM/YYYY), US formats (MM/DD/YYYY), verbal dates (15-Feb-2024), and native Excel integer serial numbers (such as 45337). Invalid or missing dates evaluate to None without halting the pipeline.
- Bank IBANs: Account numbers are stripped of internal spaces and surrounding whitespace, then converted to uppercase. Placeholders such as "NAN", "NONE", "NULL", "0", or "UNASSIGNED" are excluded from duplicate cluster matching to prevent false positives.

If input registers lack gross or net columns, the engine calculates gross earnings from basic, allowance, and overtime components, and computes net disbursement by subtracting deductions.

---

## 3. Forensic Rules and Heuristics

The engine executes seven deterministic forensic rules designed to uncover payroll fraud, control overrides, and processing breakdowns:

### Rule 0: Intra-Register Duplicate Employee ID Collisions
The engine audits both current and prior payroll registers for duplicate employee IDs appearing multiple times within the same monthly file. When an ID appears across multiple rows, the engine flags a critical severity exception and quantifies the total disbursement at risk. To prevent silent data loss, financial components across duplicate rows are aggregated so that variance waterfalls and general ledger reconciliations reflect 100% of actual disbursed funds. This test isolates duplicate payment vouchers, unbatched manual adjustments, and ghost identity reuse.

### Rule 1: Duplicate IBAN Cluster Detection
The engine groups current payroll records by normalized International Bank Account Numbers. When two or more distinct employee identification numbers share a single bank account, the engine isolates the cluster, flags all associated records as critical risk, and calculates the total disbursements diverted. This test identifies ghost employees created by payroll officers, kickback arrangements where multiple workers funnel wages to a shared intermediary, or unmonitored bank detail edits made directly inside the master payroll database.

### Rule 2: Post-Termination Salary Leakage
Employees marked as terminated in the prior month register or identified with termination and exit dates in the human resources movement log are cross-checked against current payroll records. If a terminated individual continues to draw active gross or net pay, the engine issues a critical alert and records the unrecovered compensation. This targets offboarding breakdown where line management releases staff but fails to submit payroll cut-off notices, or where payroll operators deliberately leave departed employees active on the bank transfer file.

### Rule 3: Negative Deductions and Net Exceeding Gross
Net pay should never mathematically exceed gross earnings. The engine tests every row for negative deductions or inverted balances. Negative deductions are frequently used as unauthorized off-ledger supplements, advance clearance bypasses, or manual tax adjustments that artificially inflate cash transfers without triggering gross payroll budget alarms.

### Rule 4: Managerial Overtime Breaches
Statutory labor regulations and corporate remuneration frameworks routinely categorize management personnel, directors, department heads, and executives as exempt from hourly overtime. The engine parses designations and departmental assignments for managerial keywords and flags any overtime disbursements paid to exempt titles. This catches manual timesheet additions and supervisory overrides designed to top up executive remuneration without Remuneration Committee review.

### Rule 5: Unapproved Basic Salary Spikes
Base compensation adjustments should occur only during scheduled performance reviews or formal promotions documented by human resources. The engine flags any employee whose basic pay increases by more than 15% month-over-month. When an HR movement log is provided, the engine cross-references the employee against authorized promotion and compensation change records. Increases lacking written human resources approval are escalated as high-risk unauthorized payroll tampering.

### Rule 6: Manual Round-Number Net Pay Plugs
Standard payroll calculations yield odd fractional figures due to pro-rata daily rates, statutory social insurance, and tax withholdings. The engine scans net pay for large round numbers (multiples of $1,000 or local currency equivalents) where the net figure does not balance with earnings less deductions. These occurrences indicate manual plug figures entered at month-end to override systematic payroll computations.

---

## 4. Headcount Reconciliation Bridge

The engine compiles a complete workforce population bridge between comparative cycles:

Starting Active Headcount (Prior Month)  
+ Additions and New Entrants (IDs in current register but absent from prior register)  
- Departures and Leavers (IDs in prior register but absent from current register)  
= Ending Active Headcount (Current Month)  

When an HR movement log is supplied, the engine compares actual register additions against authorized recruitment entries. New employees appearing on the payroll register without corresponding HR recruitment records are flagged as prospective ghost entries.

---

## 5. Execution Modes and CLI Operations

### Interactive Mode
To run the engine interactively, execute:
```bash
python payroll_audit_engine.py
```
The program prompts for:
1. Prior month payroll register file path (.xlsx or .csv)
2. Current month payroll register file path (.xlsx or .csv)
3. Optional HR movement log file path (.xlsx or .csv, press Enter to bypass)
4. Target output report path (defaults to a timestamped Excel workbook)

File paths dragged and dropped directly from Windows Explorer into the terminal window have surrounding single and double quotation marks automatically stripped.

### Automated Batch Mode
For automated scheduled execution or continuous controls integration, run with command-line flags:
```bash
python payroll_audit_engine.py \
  --prior "./sample_data/Payroll_Register_M01.xlsx" \
  --current "./sample_data/Payroll_Register_M02.xlsx" \
  --hr "./sample_data/HR_Movement_Log.xlsx" \
  --output "./Audit_Reports/Payroll_Audit_Report.xlsx" \
  --spike-threshold 0.15 \
  --round-threshold 1000.0 \
  --non-interactive
```

---

## 6. Audit Report Structure and Working Paper Standards

The engine generates an institutional four-tab Excel workbook formatted with locked panes, corporate navy headers (#1F4E79), zebra-striped data rows, and standard accounting number formatting:

### Tab 1: Executive Dashboard
Contains summary key performance indicator cards: total gross payroll expense, month-over-month variance, active workforce headcount, net employee movement, total forensic exceptions flagged, and cumulative financial exposure at risk. The sheet also incorporates the headcount bridge table, component-by-component variance summary, and departmental cost center variance schedules.

### Tab 2: Exception Register
Presents every transaction-level anomaly classified by severity (Critical, High, Medium). Fields include Exception ID, Severity, Anomaly Category, Employee ID, Name, Department, Designation, Financial Exposure, Flag Reason, Recommended Action, Management Response, Audit Disposition, Auditor Signoff, and Follow-up Due Date. This sheet serves as the primary working paper for field interviews and audit committee reporting.

### Tab 3: Variance Waterfall
Lists 100% of employees across both periods. For each worker, prior, current, and variance amounts are displayed across basic salary, allowances, overtime, deductions, gross earnings, and net take-home pay, alongside individual variance percentages and audit exception flags.

### Tab 4: Duplicate IBAN Clusters
Isolates all shared bank accounts into dedicated investigation schedules. Displays cluster identifiers, IBAN details, record counts, associated employee names and departments, individual net amounts, total cluster cash paid, suspect classification, and risk commentary.

---

## 7. Forensic Investigation and Escalation Protocols

When the engine surfaces critical findings, audit teams should follow structured evidence gathering procedures before conducting subject interviews:

For duplicate IBAN clusters, obtain official bank account confirmation letters and government identification cards for every individual linked to the shared account. Compare hire dates, signature cards, and physical duty stations. If a payroll administrator, human resources officer, or finance clerk is part of the cluster, escalate immediately to the Chief Financial Officer and General Counsel.

For post-termination salary leakage, file a formal recall notice with the clearing bank immediately. Pull personnel exit forms, hand-over signoffs, and IT directory deactivation logs to establish whether the disbursement reflects an administrative cut-off omission or an intentional diversion of funds.

For unapproved salary spikes, inspect master personnel files for signed remuneration committee minutes, executive approval memos, or contractual addenda. Review ERP system audit trails to extract the exact user ID, workstation IP, and timestamp associated with the compensation modification.

For duplicate employee ID collisions, inspect the underlying batch file and bank clearance records to verify whether multiple cash disbursements actually cleared the bank or represent an unposted subledger draft.

Preserve digital chain of custody by exporting original payroll files, ERP audit logs, and the engine-generated Excel deliverable to write-protected storage before commencing formal inquiries.
