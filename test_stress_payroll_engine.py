#!/usr/bin/env python3
"""
Stress Test Suite for Product 13: Payroll Forensic Variance and Integrity Engine
Veritux Consulting Network - Rigorous Controls Verification
Author: Majid Mumtaz (CIA, ACA, FCCA, GRCP, COSO ERM Certified)

Executes rigorous stress testing across:
1. Dirty data parsing: currency symbols, comma separators, accounting brackets, percentages, Excel errors
2. Missing non-mandatory columns (no overtime, no allowances, no department)
3. Intra-register duplicate employee IDs (collision detection and anti-overwriting)
4. Extreme register sizes: zero-row empty files and single-employee registers
5. Mixed date formats and Excel integer serial timestamps in HR logs
6. Scalability benchmark on 10,000 employees (< 5.0 seconds execution target)
7. Case-insensitivity and whitespace resilience in IBANs and IDs
"""

import unittest
import os
import time
import tempfile
import pandas as pd
import numpy as np
from datetime import datetime, date

# Import engine components
from payroll_audit_engine import (
    clean_numeric_value,
    clean_date_value,
    clean_iban_value,
    clean_emp_id_value,
    clean_text_value,
    map_dataframe_schema,
    PayrollForensicEngine,
    ExcelAuditReportBuilder,
    CANONICAL_FIELDS,
    HR_LOG_FIELDS
)

class TestPayrollEngineStress(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def test_clean_numeric_value_dirty_data(self):
        """Test numeric parser resilience against symbols, separators, errors, and accounting format."""
        test_cases = [
            ("$5,000.00", 5000.0),
            ("AED 15,000.50", 15000.50),
            ("SAR 8,500", 8500.0),
            ("£3,200.75", 3200.75),
            ("€2,400.00", 2400.0),
            ("Dhs 4,500.00", 4500.0),
            ("USD 10,000.00", 10000.0),
            ("1,250.00", 1250.0),
            ("12,500", 12500.0),
            ("1.250,50", 1250.50),       # European format
            ("($500.00)", -500.0),       # Accounting brackets
            ("(1,200.50)", -1200.50),
            ("-$750.00", -750.0),
            ("-750.00", -750.0),
            ("750.00-", -750.0),
            ("15%", 0.15),
            ("25.5%", 0.255),
            ("#N/A", 0.0),               # Excel formula errors
            ("#DIV/0!", 0.0),
            ("#REF!", 0.0),
            ("#VALUE!", 0.0),
            ("#NAME?", 0.0),
            ("#NUM!", 0.0),
            ("#NULL!", 0.0),
            ("None", 0.0),
            ("nan", 0.0),
            ("NULL", 0.0),
            ("-", 0.0),
            ("--", 0.0),
            ("", 0.0),
            (None, 0.0),
            (np.nan, 0.0),
            (7500, 7500.0),
            (7500.25, 7500.25),
            ("  8500.00  ", 8500.0)
        ]

        for raw_val, expected in test_cases:
            result = clean_numeric_value(raw_val)
            self.assertAlmostEqual(
                result, expected, places=2,
                msg=f"Failed parsing {raw_val!r}: got {result}, expected {expected}"
            )

    def test_clean_date_value_formats_and_excel_serials(self):
        """Test date parser across ISO, European, US, verbal, and Excel integer serials."""
        test_cases = [
            ("2024-02-15", "2024-02-15"),
            ("2024/02/15", "2024-02-15"),
            ("15/02/2024", "2024-02-15"),
            ("15-02-2024", "2024-02-15"),
            ("02/15/2024", "2024-02-15"),
            ("15-Feb-2024", "2024-02-15"),
            ("15 February 2024", "2024-02-15"),
            (45337, "2024-02-15"),         # Excel integer timestamp for 15-Feb-2024
            ("45337", "2024-02-15"),       # String serial
            (date(2024, 2, 15), "2024-02-15"),
            (datetime(2024, 2, 15, 14, 30), "2024-02-15"),
            (pd.Timestamp("2024-02-15"), "2024-02-15"),
            ("#VALUE!", None),
            ("#N/A", None),
            ("corrupt_date_string", None),
            ("", None),
            (None, None)
        ]

        for raw_val, expected in test_cases:
            result = clean_date_value(raw_val)
            self.assertEqual(
                result, expected,
                msg=f"Failed parsing date {raw_val!r}: got {result}, expected {expected}"
            )

    def test_clean_iban_whitespace_and_casing(self):
        """Test IBAN cleaning removes internal spaces, external padding, and upper-cases."""
        test_cases = [
            ("ae440330000100800001", "AE440330000100800001"),
            ("  AE44 0330 0001 0080 0001  ", "AE440330000100800001"),
            ("Ae44 0330 0001 0080 0001", "AE440330000100800001"),
            ("NAN", ""),
            ("NONE", ""),
            ("0", ""),
            ("#N/A", ""),
            ("#VALUE!", ""),
            ("UNASSIGNED", ""),
            (None, "")
        ]

        for raw_val, expected in test_cases:
            result = clean_iban_value(raw_val)
            self.assertEqual(result, expected, msg=f"Failed cleaning IBAN {raw_val!r}")

    def test_missing_optional_columns_imputation(self):
        """Verify engine gracefully processes registers lacking Overtime, Allowances, or Department."""
        # Minimalist raw register: only ID, Name, Basic, Deductions, Gross, Net
        raw_prior = pd.DataFrame({
            "Employee ID": ["EMP-001", "EMP-002"],
            "Full Name": ["Alice Smith", "Bob Jones"],
            "Basic Salary": ["$5,000.00", "$6,000.00"],
            "Deductions": ["$500.00", "$600.00"],
            "Gross Pay": ["$5,000.00", "$6,000.00"],
            "Net Pay": ["$4,500.00", "$5,400.00"],
            "IBAN": ["AE01001", "AE01002"]
        })

        raw_curr = pd.DataFrame({
            "Employee ID": ["EMP-001", "EMP-002"],
            "Full Name": ["Alice Smith", "Bob Jones"],
            "Basic Salary": ["$5,200.00", "$6,000.00"],
            "Deductions": ["$520.00", "$600.00"],
            "Gross Pay": ["$5,200.00", "$6,000.00"],
            "Net Pay": ["$4,680.00", "$5,400.00"],
            "IBAN": ["AE01001", "AE01002"]
        })

        df_p, _ = map_dataframe_schema(raw_prior, CANONICAL_FIELDS, "Prior")
        df_c, _ = map_dataframe_schema(raw_curr, CANONICAL_FIELDS, "Current")

        # Verify imputed defaults
        self.assertTrue((df_p["overtime"] == 0.0).all())
        self.assertTrue((df_p["allowances"] == 0.0).all())
        self.assertTrue((df_p["department"] == "Unassigned").all())
        self.assertTrue((df_p["designation"] == "General Staff").all())

        engine = PayrollForensicEngine(df_prior=df_p, df_current=df_c)
        engine.run_all_checks()

        # Verify report generation succeeds without errors
        out_xlsx = os.path.join(self.temp_dir, "test_missing_cols.xlsx")
        builder = ExcelAuditReportBuilder(engine, out_xlsx)
        builder.build_report()
        self.assertTrue(os.path.exists(out_xlsx))

    def test_intra_register_duplicate_employee_ids(self):
        """Verify duplicate employee IDs within the same file are flagged as CRITICAL collisions."""
        raw_prior = pd.DataFrame({
            "emp_id": ["EMP-001", "EMP-002"],
            "name": ["Alice Smith", "Bob Jones"],
            "department": ["Finance", "Operations"],
            "designation": ["Analyst", "Operator"],
            "basic_pay": [5000.0, 4000.0],
            "allowances": [1000.0, 800.0],
            "overtime": [0.0, 0.0],
            "deductions": [500.0, 400.0],
            "gross_pay": [6000.0, 4800.0],
            "net_pay": [5500.0, 4400.0],
            "bank_iban": ["AE01", "AE02"],
            "status": ["Active", "Active"]
        })

        # Current register contains EMP-002 TWICE (duplicate voucher collision)
        raw_curr = pd.DataFrame({
            "emp_id": ["EMP-001", "EMP-002", "EMP-002"],
            "name": ["Alice Smith", "Bob Jones", "Bob Jones"],
            "department": ["Finance", "Operations", "Operations"],
            "designation": ["Analyst", "Operator", "Operator"],
            "basic_pay": [5000.0, 4000.0, 4000.0],
            "allowances": [1000.0, 800.0, 800.0],
            "overtime": [0.0, 0.0, 0.0],
            "deductions": [500.0, 400.0, 400.0],
            "gross_pay": [6000.0, 4800.0, 4800.0],
            "net_pay": [5500.0, 4400.0, 4400.0],
            "bank_iban": ["AE01", "AE02", "AE02"],
            "status": ["Active", "Active", "Active"]
        })

        engine = PayrollForensicEngine(df_prior=raw_prior, df_current=raw_curr)
        engine.run_all_checks()

        # Find duplicate ID exceptions
        dup_exceptions = [e for e in engine.exceptions if e["category"] == "Duplicate Employee ID Collision"]
        self.assertGreaterEqual(len(dup_exceptions), 1)
        self.assertEqual(dup_exceptions[0]["emp_id"], "EMP-002")
        self.assertEqual(dup_exceptions[0]["severity"], "CRITICAL")
        self.assertEqual(dup_exceptions[0]["impact_amount"], 8800.0)

        # Confirm lookup map aggregated the dollars rather than silently discarding row 2
        self.assertEqual(engine.curr_map["EMP-002"]["net_pay"], 8800.0)
        self.assertEqual(engine.curr_map["EMP-002"]["gross_pay"], 9600.0)

    def test_zero_row_and_single_employee_registers(self):
        """Verify zero division crashes are prevented when registers have 0 rows or 1 row."""
        empty_df = pd.DataFrame(columns=list(CANONICAL_FIELDS.keys()))
        single_df = pd.DataFrame({
            "emp_id": ["EMP-001"],
            "name": ["Solo Worker"],
            "department": ["Executive"],
            "designation": ["Director"],
            "basic_pay": [10000.0],
            "allowances": [2000.0],
            "overtime": [0.0],
            "deductions": [1000.0],
            "gross_pay": [12000.0],
            "net_pay": [11000.0],
            "bank_iban": ["AE999"],
            "status": ["Active"]
        })

        # Case 1: Both empty
        eng_empty = PayrollForensicEngine(df_prior=empty_df, df_current=empty_df)
        eng_empty.run_all_checks()
        out1 = os.path.join(self.temp_dir, "test_empty.xlsx")
        b1 = ExcelAuditReportBuilder(eng_empty, out1)
        b1.build_report()
        self.assertTrue(os.path.exists(out1))
        self.assertEqual(eng_empty.kpi_summary["active_headcount"], 0)

        # Case 2: Prior empty, Current single employee
        eng_add = PayrollForensicEngine(df_prior=empty_df, df_current=single_df)
        eng_add.run_all_checks()
        self.assertEqual(eng_add.headcount_bridge["starting_headcount"], 0)
        self.assertEqual(eng_add.headcount_bridge["ending_headcount"], 1)
        self.assertEqual(eng_add.headcount_bridge["additions_count"], 1)

        # Case 3: Prior single, Current empty (100% departure)
        eng_dep = PayrollForensicEngine(df_prior=single_df, df_current=empty_df)
        eng_dep.run_all_checks()
        self.assertEqual(eng_dep.headcount_bridge["departures_count"], 1)

    def test_hr_lifecycle_date_serials_and_spike_authorization(self):
        """Verify HR log with Excel serial date authorizes promotions and flags post-termination leakage."""
        prior_df = pd.DataFrame({
            "emp_id": ["EMP-001", "EMP-002", "EMP-003"],
            "name": ["Alice", "Bob", "Charlie"],
            "department": ["Finance", "Operations", "Sales"],
            "designation": ["Analyst", "Operator", "Rep"],
            "basic_pay": [5000.0, 4000.0, 3000.0],
            "allowances": [0.0, 0.0, 0.0],
            "overtime": [0.0, 0.0, 0.0],
            "deductions": [0.0, 0.0, 0.0],
            "gross_pay": [5000.0, 4000.0, 3000.0],
            "net_pay": [5000.0, 4000.0, 3000.0],
            "bank_iban": ["AE01", "AE02", "AE03"],
            "status": ["Active", "Active", "Active"]
        })

        curr_df = pd.DataFrame({
            "emp_id": ["EMP-001", "EMP-002", "EMP-003"],
            "name": ["Alice", "Bob", "Charlie"],
            "department": ["Finance", "Operations", "Sales"],
            "designation": ["Analyst", "Operator", "Rep"],
            "basic_pay": [7000.0, 6000.0, 3000.0],   # EMP-001 +40% (authorized), EMP-002 +50% (unauthorized)
            "allowances": [0.0, 0.0, 0.0],
            "overtime": [0.0, 0.0, 0.0],
            "deductions": [0.0, 0.0, 0.0],
            "gross_pay": [7000.0, 6000.0, 3000.0],
            "net_pay": [7000.0, 6000.0, 3000.0],
            "bank_iban": ["AE01", "AE02", "AE03"],
            "status": ["Active", "Active", "Active"]
        })

        raw_hr = pd.DataFrame({
            "Employee ID": ["EMP-001", "EMP-003"],
            "Event": ["SALARY ADJUSTMENT / PROMOTION", "TERMINATION / RESIGNED"],
            "Effective Date": [45337, "15/02/2024"], # Serial 45337 and European string
            "Approver": ["CFO", "HR Director"],
            "Notes": ["Annual review increase", "Left company on Feb 15"]
        })

        df_p, _ = map_dataframe_schema(prior_df, CANONICAL_FIELDS)
        df_c, _ = map_dataframe_schema(curr_df, CANONICAL_FIELDS)
        df_hr, _ = map_dataframe_schema(raw_hr, HR_LOG_FIELDS)

        # Check effective date was converted from serial 45337 to '2024-02-15'
        self.assertEqual(df_hr.iloc[0]["effective_date"], "2024-02-15")
        self.assertEqual(df_hr.iloc[1]["effective_date"], "2024-02-15")

        engine = PayrollForensicEngine(df_prior=df_p, df_current=df_c, df_hr=df_hr)
        engine.run_all_checks()

        spike_eids = [e["emp_id"] for e in engine.exceptions if e["category"] == "Unapproved Basic Salary Spike"]
        # EMP-001 had authorized promotion in HR log -> not flagged
        self.assertNotIn("EMP-001", spike_eids)
        # EMP-002 had no HR authorization -> flagged
        self.assertIn("EMP-002", spike_eids)

        # EMP-003 was terminated in HR log but received Net Pay -> flagged as zombie payroll
        zombie_eids = [e["emp_id"] for e in engine.exceptions if e["category"] == "Post-Termination Salary Leakage"]
        self.assertIn("EMP-003", zombie_eids)

    def test_duplicate_iban_casing_and_clustering(self):
        """Verify mixed-case and spaced IBANs are clustered together accurately."""
        prior_df = pd.DataFrame({
            "emp_id": ["EMP-101", "EMP-102"],
            "name": ["John", "Jane"],
            "department": ["Finance", "Logistics"],
            "designation": ["Payroll Officer", "Driver"],
            "basic_pay": [5000.0, 4000.0],
            "allowances": [0.0, 0.0],
            "overtime": [0.0, 0.0],
            "deductions": [0.0, 0.0],
            "gross_pay": [5000.0, 4000.0],
            "net_pay": [5000.0, 4000.0],
            "bank_iban": ["AE44 0330 0001 0080 0001", "ae440330000100800001"],
            "status": ["Active", "Active"]
        })

        curr_raw = pd.DataFrame({
            "emp_id": ["EMP-101", "EMP-102"],
            "name": ["John", "Jane"],
            "department": ["Finance", "Logistics"],
            "designation": ["Payroll Officer", "Driver"],
            "basic_pay": [5000.0, 4000.0],
            "allowances": [0.0, 0.0],
            "overtime": [0.0, 0.0],
            "deductions": [0.0, 0.0],
            "gross_pay": [5000.0, 4000.0],
            "net_pay": [5000.0, 4000.0],
            "bank_iban": ["AE44 0330 0001 0080 0001", "  ae440330000100800001  "],
            "status": ["Active", "Active"]
        })

        df_c, _ = map_dataframe_schema(curr_raw, CANONICAL_FIELDS)
        df_p, _ = map_dataframe_schema(prior_df, CANONICAL_FIELDS)

        engine = PayrollForensicEngine(df_prior=df_p, df_current=df_c)
        engine.run_all_checks()

        # Must find exactly 1 cluster with 2 distinct employees
        self.assertEqual(len(engine.iban_clusters), 1)
        cluster = engine.iban_clusters[0]
        self.assertEqual(cluster["record_count"], 2)
        self.assertEqual(cluster["iban"], "AE440330000100800001")
        self.assertIn("Critical Collusion Risk", cluster["suspect_classification"])

    def test_scalability_10000_employees_benchmark(self):
        """Benchmark: Processing 10,000 employees must execute in under 5.0 seconds."""
        n = 10000

        # Prior workforce
        prior_records = []
        for i in range(n):
            prior_records.append({
                "emp_id": f"EMP-{i:06d}",
                "name": f"Employee {i}",
                "department": f"Dept-{i % 15}",
                "designation": "Staff",
                "basic_pay": 6000.0 + (i % 200),
                "allowances": 1500.0,
                "overtime": 0.0,
                "deductions": 750.0,
                "gross_pay": 7500.0 + (i % 200),
                "net_pay": 6750.0 + (i % 200),
                "bank_iban": f"AE{20 + (i % 70):02d}033000{i:010d}",
                "status": "Active"
            })
        df_p = pd.DataFrame(prior_records)

        # Current workforce: 9,800 retentions, 200 departures, 200 additions = 10,000 total
        curr_records = []
        # 9,800 retained employees
        for i in range(n - 200):
            row = dict(prior_records[i])
            # Seed 5 unapproved spikes
            if i < 5:
                row["basic_pay"] = row["basic_pay"] * 1.30
                row["gross_pay"] = row["gross_pay"] * 1.30
                row["net_pay"] = row["net_pay"] * 1.30
            # Seed 5 managerial overtime
            if 10 <= i < 15:
                row["designation"] = "Operations Manager"
                row["overtime"] = 1200.0
                row["gross_pay"] += 1200.0
                row["net_pay"] += 1200.0
            # Seed 3 duplicate IBANs
            if 20 <= i < 23:
                row["bank_iban"] = "AE9903300000SHAREDIBAN"
            curr_records.append(row)

        # 200 new additions
        for i in range(n, n + 200):
            curr_records.append({
                "emp_id": f"EMP-{i:06d}",
                "name": f"New Hire {i}",
                "department": f"Dept-{i % 15}",
                "designation": "Staff",
                "basic_pay": 6500.0,
                "allowances": 1500.0,
                "overtime": 0.0,
                "deductions": 750.0,
                "gross_pay": 8000.0,
                "net_pay": 7250.0,
                "bank_iban": f"AE{20 + (i % 70):02d}033000{i:010d}",
                "status": "Active"
            })
        df_c = pd.DataFrame(curr_records)

        start_time = time.time()
        engine = PayrollForensicEngine(df_prior=df_p, df_current=df_c)
        engine.run_all_checks()
        duration = time.time() - start_time

        print(f"\n[BENCHMARK] Processed 10,000 employees in: {duration:.3f} seconds (Target: < 5.0 s)")

        # Assert speed
        self.assertLess(
            duration, 5.0,
            msg=f"Scalability test failed: took {duration:.2f}s, expected < 5.0s"
        )

        # Assert correctness
        self.assertEqual(engine.headcount_bridge["starting_headcount"], 10000)
        self.assertEqual(engine.headcount_bridge["ending_headcount"], 10000)
        self.assertEqual(engine.headcount_bridge["additions_count"], 200)
        self.assertEqual(engine.headcount_bridge["departures_count"], 200)
        self.assertGreaterEqual(len(engine.exceptions), 10)

if __name__ == "__main__":
    unittest.main()
