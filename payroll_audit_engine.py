#!/usr/bin/env python3
"""
Product 13: Payroll Forensic Variance and Integrity Engine
Veritux Consulting Network - Forensic Audit & Continuous Controls Monitoring
Author: Majid Mumtaz (CIA, ACA, FCCA, GRCP, COSO ERM Certified)

Ingests Prior Month Register, Current Month Register, and optional HR Movement Log.
Performs dynamic column schema mapping, headcount bridge reconciliation,
component-by-component variance calculations, and executes forensic integrity checks:
- Duplicate IBAN detection across distinct employee IDs
- Intra-register duplicate employee ID collisions
- Post-termination salary leakage (Zombie payroll)
- Negative deductions & Net > Gross arithmetic violations
- Managerial overtime breaches contrary to exempt policy
- Unapproved basic salary spikes (>15%) without promotion/adjustment
- Round-number manual net pay plugs

Outputs terminal executive telemetry and an institutional 4-tab Excel audit report.
"""

import sys
import os
import re
import argparse
from datetime import datetime, date, timedelta
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Try importing rich for enhanced terminal display
try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.text import Text
    HAS_RICH = True
    console = Console()
except ImportError:
    HAS_RICH = False
    console = None

# ==============================================================================
# SCHEMA MAPPER CONFIGURATION
# ==============================================================================

CANONICAL_FIELDS = {
    "emp_id": [
        "empid", "employeeid", "staffid", "personnelno", "personnelnumber",
        "payrollid", "badgeid", "id", "empno", "workerid", "employeeno",
        "employeenumber", "empnumber", "staffnumber", "badgenumber", "fileno",
        "filenumber", "payrollnumber", "employeecode", "empcode", "staffcode"
    ],
    "name": [
        "employeename", "staffname", "fullname", "workername", "personname",
        "name", "employee", "staff"
    ],
    "department": [
        "department", "dept", "costcenter", "costcentre", "businessunit",
        "division", "team", "function", "cost_center", "unit", "branch"
    ],
    "designation": [
        "designation", "jobtitle", "title", "role", "position",
        "grade", "jobrole", "level", "post", "occupation", "job"
    ],
    "basic_pay": [
        "basicsalary", "basicpay", "basepay", "basesalary", "basic",
        "salarybasic", "monthlybasic", "fixedsalary", "base"
    ],
    "allowances": [
        "allowances", "allowance", "totalallowances", "housingallowance",
        "transportallowance", "fixedallowances", "otherallowances", "allowancetotal",
        "benefits", "totalbenefits"
    ],
    "overtime": [
        "overtimepay", "overtime", "otpay", "ot", "otamount",
        "extrahourspay", "overtimeamount", "otallowance", "othoursamount"
    ],
    "deductions": [
        "totaldeductions", "deductions", "deduction", "taxdeductions",
        "statutorydeductions", "otherdeductions", "deductionstotal", "withholdings"
    ],
    "gross_pay": [
        "grosspay", "grosssalary", "gross", "totalgross",
        "grossearnings", "totalgrosssalary", "totalearnings", "earnings"
    ],
    "net_pay": [
        "netpay", "netsalary", "net", "takehome", "totalnet",
        "netpayable", "transferamount", "disbursement", "disbursedamount"
    ],
    "bank_iban": [
        "bankiban", "iban", "accountnumber", "bankaccount",
        "accountno", "bankacc", "routingaccount", "ibanaccount",
        "acctno", "acctnumber", "accno", "bankacct", "bankaccountno"
    ],
    "status": [
        "employmentstatus", "status", "termdate", "terminationdate",
        "exitdate", "active", "state", "workstatus", "empstatus"
    ]
}

HR_LOG_FIELDS = {
    "emp_id": ["empid", "employeeid", "staffid", "id", "personnelno", "empno", "workerid", "employeeno"],
    "name": ["employeename", "staffname", "fullname", "name", "workername"],
    "event_type": ["eventtype", "movementtype", "event", "action", "type", "actiontype", "reason"],
    "effective_date": ["effectivedate", "date", "effdate", "actiondate", "startdate", "eventdate"],
    "approved_by": ["approvedby", "approver", "authorizer", "signedoffby", "manager"],
    "old_value": ["oldvalue", "priorvalue", "from", "previousvalue"],
    "new_value": ["newvalue", "currentvalue", "to", "updatedvalue"],
    "notes": ["notes", "remarks", "reason", "comment", "description"]
}

def clean_str(val):
    """Normalize string for fuzzy column header matching."""
    if val is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(val).strip().lower())

def clean_file_path(path_str):
    """Strip quotes and surrounding whitespace from Windows Explorer drag-and-drop paths."""
    if not path_str:
        return ""
    p = str(path_str).strip()
    if (p.startswith('"') and p.endswith('"')) or (p.startswith("'") and p.endswith("'")):
        p = p[1:-1].strip()
    p = p.replace('"', '').replace("'", "")
    return os.path.normpath(p)

def clean_numeric_value(val):
    """
    Robust numeric parser for payroll registers:
    Strips currency symbols ($, AED, SAR, GBP, EUR, etc.),
    thousands separators, European dot/comma notation,
    percentages, negative accounting brackets, and Excel errors (#N/A, #VALUE!, etc.).
    """
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val) if not (isinstance(val, float) and np.isnan(val)) else 0.0

    s = str(val).strip()
    if not s:
        return 0.0

    # Common Excel formula error strings and null tokens
    if s.lower() in ("none", "nan", "null", "n/a", "-", "--", "—", "#n/a", "#div/0!", "#ref!", "#value!", "#name?", "#num!", "#null!", "unknown"):
        return 0.0

    # Negative accounting notation: (1,234.56) or ($1,234.56)
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        negative = True
        s = s[1:].strip()
    elif s.endswith("-"):
        negative = True
        s = s[:-1].strip()

    is_percent = False
    if s.endswith("%"):
        is_percent = True
        s = s[:-1].strip()

    # Strip currency code words: USD, AED, SAR, GBP, EUR, QAR, OMR, BHD, KWD, DHS, DH, RS, INR
    s = re.sub(r"(?i)\b(USD|AED|SAR|GBP|EUR|QAR|OMR|BHD|KWD|DHS|DH|RS|INR)\b", "", s)

    # Strip currency symbols and whitespace
    s = re.sub(r"[\$£€¥₹৳\s]", "", s)

    # Commas and period resolution
    if "," in s and "." in s:
        if s.rfind(",") < s.rfind("."):
            # Standard US/UK/GCC: 1,234.56
            s = s.replace(",", "")
        else:
            # European: 1.234,56
            s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        parts = s.split(",")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = "".join(parts)
        elif len(parts) == 2 and len(parts[1]) in (1, 2):
            s = parts[0] + "." + parts[1]
        else:
            s = s.replace(",", "")

    try:
        val_float = float(s)
        if is_percent:
            val_float = val_float / 100.0
        return -val_float if negative else val_float
    except ValueError:
        return 0.0

def clean_date_value(val):
    """
    Robust date parser:
    Converts Excel integer/float serial numbers (e.g. 45337),
    ISO (YYYY-MM-DD), European (DD/MM/YYYY), US (MM/DD/YYYY),
    and verbal dates into standard ISO YYYY-MM-DD.
    Returns None for invalid dates or Excel errors.
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")

    # Excel serial dates
    if isinstance(val, (int, float)):
        if 30000 <= val <= 85000:
            d = datetime(1899, 12, 30) + timedelta(days=float(val))
            return d.strftime("%Y-%m-%d")
        return None

    s = str(val).strip()
    if not s or s.lower() in ("none", "nan", "null", "n/a", "-", "--", "#n/a", "#value!", "#ref!", "#div/0!"):
        return None

    if s.isdigit():
        num = int(s)
        if 30000 <= num <= 85000:
            d = datetime(1899, 12, 30) + timedelta(days=num)
            return d.strftime("%Y-%m-%d")

    date_formats = [
        "%Y-%m-%d", "%Y/%m/%d",
        "%d/%m/%Y", "%d-%m-%Y",
        "%m/%d/%Y", "%m-%d-%Y",
        "%d-%b-%Y", "%d-%B-%Y",
        "%d %b %Y", "%d %B %Y",
        "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in date_formats:
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime("%Y-%m-%d")
        except ValueError:
            pass

    try:
        dt = pd.to_datetime(s, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass

    return None

def clean_text_value(val, default=""):
    """Strip whitespace and replace Excel formula error strings with a clean default."""
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    if s.lower() in ("none", "nan", "null", "n/a", "#n/a", "#value!", "#ref!", "#div/0!", "#name?", "#num!", "#null!", ""):
        return default
    return s

def clean_iban_value(val):
    """Strip all internal/external whitespace and uppercase IBAN string."""
    if val is None or pd.isna(val):
        return ""
    s = re.sub(r"\s+", "", str(val)).upper()
    if s in ("", "NAN", "NONE", "NULL", "0", "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "UNASSIGNED", "UNKNOWN", "TBD", "-"):
        return ""
    return s

def clean_emp_id_value(val, row_idx):
    """Strip whitespace from Employee ID. Generate deterministic placeholder if missing."""
    if val is None or pd.isna(val):
        return f"UNKNOWN-ROW-{row_idx + 1}"
    s = str(val).strip()
    if s.lower() in ("", "none", "nan", "null", "n/a", "#n/a", "#value!", "#ref!", "#div/0!"):
        return f"UNKNOWN-ROW-{row_idx + 1}"
    return s

def map_dataframe_schema(df, field_dict, source_name="File"):
    """
    Intelligently map columns in df to canonical field names using normalized string matching.
    Cleans and standardizes numeric, text, date, and IBAN fields.
    """
    if df is None or df.empty:
        empty_cols = list(field_dict.keys())
        empty_df = pd.DataFrame(columns=empty_cols)
        for nf in ["basic_pay", "allowances", "overtime", "deductions", "gross_pay", "net_pay"]:
            if nf in empty_df.columns:
                empty_df[nf] = empty_df[nf].astype(float)
        return empty_df, {}

    col_mapping = {}
    normalized_cols = {clean_str(col): col for col in df.columns}
    mapped_orig_cols = set()

    for canonical, aliases in field_dict.items():
        matched = False
        # Direct normalized match
        for alias in aliases:
            if alias in normalized_cols and normalized_cols[alias] not in mapped_orig_cols:
                col_mapping[canonical] = normalized_cols[alias]
                mapped_orig_cols.add(normalized_cols[alias])
                matched = True
                break

        # Substring / partial match if not found
        if not matched:
            for norm_col, original_col in normalized_cols.items():
                if original_col in mapped_orig_cols:
                    continue
                if any(alias in norm_col for alias in aliases):
                    col_mapping[canonical] = original_col
                    mapped_orig_cols.add(original_col)
                    matched = True
                    break

    # Construct standardized dataframe
    mapped_df = pd.DataFrame(index=df.index)
    for canonical in field_dict.keys():
        if canonical in col_mapping:
            mapped_df[canonical] = df[col_mapping[canonical]].copy()
        else:
            mapped_df[canonical] = None

    # Handle numeric columns imputation for payroll registers
    num_fields = ["basic_pay", "allowances", "overtime", "deductions", "gross_pay", "net_pay"]
    for nf in num_fields:
        if nf in mapped_df.columns:
            mapped_df[nf] = mapped_df[nf].apply(clean_numeric_value)

    # Arithmetic self-healing: gross and net calculation if missing
    if "gross_pay" in mapped_df.columns and "basic_pay" in mapped_df.columns:
        allow = mapped_df["allowances"] if "allowances" in mapped_df.columns else 0.0
        ot = mapped_df["overtime"] if "overtime" in mapped_df.columns else 0.0
        if mapped_df["gross_pay"].sum() == 0:
            mapped_df["gross_pay"] = mapped_df["basic_pay"] + allow + ot

    if "net_pay" in mapped_df.columns and "gross_pay" in mapped_df.columns and "deductions" in mapped_df.columns:
        if mapped_df["net_pay"].sum() == 0:
            mapped_df["net_pay"] = mapped_df["gross_pay"] - mapped_df["deductions"]

    if "deductions" in mapped_df.columns and "gross_pay" in mapped_df.columns and "net_pay" in mapped_df.columns:
        if mapped_df["deductions"].sum() == 0 and mapped_df["gross_pay"].sum() > 0:
            mapped_df["deductions"] = mapped_df["gross_pay"] - mapped_df["net_pay"]

    # String normalization for ID, IBAN, Status if present
    if "emp_id" in mapped_df.columns:
        mapped_df["emp_id"] = [clean_emp_id_value(val, idx) for idx, val in enumerate(mapped_df["emp_id"])]
    if "name" in mapped_df.columns:
        mapped_df["name"] = mapped_df["name"].apply(lambda v: clean_text_value(v, "Unknown"))
    if "department" in mapped_df.columns:
        mapped_df["department"] = mapped_df["department"].apply(lambda v: clean_text_value(v, "Unassigned"))
    if "designation" in mapped_df.columns:
        mapped_df["designation"] = mapped_df["designation"].apply(lambda v: clean_text_value(v, "General Staff"))
    if "bank_iban" in mapped_df.columns:
        mapped_df["bank_iban"] = mapped_df["bank_iban"].apply(clean_iban_value)
    if "status" in mapped_df.columns:
        mapped_df["status"] = mapped_df["status"].apply(lambda v: clean_text_value(v, "Active"))
    if "effective_date" in mapped_df.columns:
        mapped_df["effective_date"] = mapped_df["effective_date"].apply(clean_date_value)

    return mapped_df, col_mapping

def read_data_file(filepath):
    """Safely load .xlsx or .csv into a pandas DataFrame with encoding and empty file fallbacks."""
    clean_p = clean_file_path(filepath)
    if not os.path.exists(clean_p):
        raise FileNotFoundError(f"File not found: {clean_p}")

    ext = os.path.splitext(clean_p)[1].lower()
    try:
        if ext in [".xlsx", ".xlsm", ".xls"]:
            return pd.read_excel(clean_p)
        elif ext in [".csv", ".txt"]:
            try:
                return pd.read_csv(clean_p, encoding="utf-8")
            except UnicodeDecodeError:
                return pd.read_csv(clean_p, encoding="latin1")
            except pd.errors.EmptyDataError:
                return pd.DataFrame()
        else:
            raise ValueError(f"Unsupported file format: {ext}. Please provide .xlsx or .csv")
    except pd.errors.EmptyDataError:
        return pd.DataFrame()

# ==============================================================================
# AUDIT & FORENSIC ANALYTICS ENGINE
# ==============================================================================

class PayrollForensicEngine:
    def __init__(self, df_prior, df_current, df_hr=None, spike_threshold=0.15, round_threshold=1000.0):
        self.df_prior = df_prior.copy() if df_prior is not None else pd.DataFrame()
        self.df_current = df_current.copy() if df_current is not None else pd.DataFrame()
        self.df_hr = df_hr.copy() if df_hr is not None else None
        self.spike_threshold = spike_threshold
        self.round_threshold = round_threshold

        self.exceptions = []
        self.waterfall_records = []
        self.iban_clusters = []
        self.kpi_summary = {}

        # Detect internal duplicate Employee IDs before constructing aggregated lookup maps
        self._detect_duplicate_employee_ids()

        # Build lookup structures with financial aggregation for duplicates
        self.prior_map = self._build_lookup_map(self.df_prior, "Prior")
        self.curr_map = self._build_lookup_map(self.df_current, "Current")

    def _detect_duplicate_employee_ids(self):
        """
        Rule 0: Intra-Register Duplicate Employee ID Collision Detection.
        Flags duplicate employee IDs within the same monthly register.
        Prevents silent row overwriting, double disbursement fraud, and ghost identity reuse.
        """
        if self.df_current is not None and not self.df_current.empty and "emp_id" in self.df_current.columns:
            counts = self.df_current["emp_id"].value_counts()
            dup_ids = counts[counts > 1].index.tolist()
            for dup_id in dup_ids:
                dup_rows = self.df_current[self.df_current["emp_id"] == dup_id]
                tot_net = float(dup_rows["net_pay"].sum())
                tot_gross = float(dup_rows["gross_pay"].sum())
                first_row = dup_rows.iloc[0]
                self.exceptions.append({
                    "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                    "severity": "CRITICAL",
                    "category": "Duplicate Employee ID Collision",
                    "emp_id": dup_id,
                    "name": first_row["name"],
                    "department": first_row["department"],
                    "designation": first_row["designation"],
                    "impact_amount": tot_net,
                    "flag_reason": (
                        f"Employee ID '{dup_id}' appears {len(dup_rows)} times in current payroll register "
                        f"with cumulative Net Pay of ${tot_net:,.2f} (Gross: ${tot_gross:,.2f}). "
                        f"Internal collision indicates duplicate payment voucher or ghost identity reuse."
                    ),
                    "recommended_action": (
                        "Immediately halt payroll batch payment. Invalidate duplicate disbursement tokens "
                        "and cross-reference biometric attendance and personnel subledger."
                    )
                })

        if self.df_prior is not None and not self.df_prior.empty and "emp_id" in self.df_prior.columns:
            p_counts = self.df_prior["emp_id"].value_counts()
            p_dup_ids = p_counts[p_counts > 1].index.tolist()
            for p_dup_id in p_dup_ids:
                p_dup_rows = self.df_prior[self.df_prior["emp_id"] == p_dup_id]
                tot_net = float(p_dup_rows["net_pay"].sum())
                first_row = p_dup_rows.iloc[0]
                self.exceptions.append({
                    "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                    "severity": "HIGH",
                    "category": "Duplicate Employee ID Collision (Prior Month)",
                    "emp_id": p_dup_id,
                    "name": first_row["name"],
                    "department": first_row["department"],
                    "designation": first_row["designation"],
                    "impact_amount": tot_net,
                    "flag_reason": (
                        f"Employee ID '{p_dup_id}' appeared {len(p_dup_rows)} times in prior month register "
                        f"with cumulative Net Pay of ${tot_net:,.2f}."
                    ),
                    "recommended_action": (
                        "Perform retroactive audit of prior month bank clearing file to confirm whether "
                        "double payment was disbursed."
                    )
                })

    def _build_lookup_map(self, df, label):
        """
        Construct emp_id -> row lookup dict.
        Aggregates compensation amounts for duplicate records so no payroll dollars are dropped.
        """
        lookup = {}
        if df is None or df.empty:
            return lookup

        num_fields = ["basic_pay", "allowances", "overtime", "deductions", "gross_pay", "net_pay"]
        for row in df.to_dict("records"):
            eid = row["emp_id"]
            if eid not in lookup:
                lookup[eid] = dict(row)
            else:
                existing = lookup[eid]
                for nf in num_fields:
                    existing[nf] = float(existing.get(nf, 0.0)) + float(row.get(nf, 0.0))
                if "[COLLISION]" not in str(existing.get("designation", "")):
                    existing["designation"] = f"{existing.get('designation', '')} [COLLISION]"
        return lookup

    def run_all_checks(self):
        """Execute complete suite of forensic rules and reconciliation tests."""
        self._build_headcount_bridge()
        self._calculate_component_variances()
        self._detect_duplicate_ibans()
        self._detect_post_termination_leakage()
        self._detect_net_exceeds_gross()
        self._detect_managerial_overtime()
        self._detect_salary_spikes()
        self._detect_round_number_plugs()
        self._build_variance_waterfall()
        self._compile_executive_kpis()

    def _build_headcount_bridge(self):
        prior_ids = set(self.prior_map.keys())
        curr_ids = set(self.curr_map.keys())

        starting_count = len(self.df_prior) if self.df_prior is not None else 0
        ending_count = len(self.df_current) if self.df_current is not None else 0
        additions = curr_ids - prior_ids
        departures = prior_ids - curr_ids
        net_change = ending_count - starting_count

        self.headcount_bridge = {
            "starting_headcount": starting_count,
            "additions_count": len(additions),
            "departures_count": len(departures),
            "ending_headcount": ending_count,
            "net_change": net_change,
            "additions_list": sorted(list(additions)),
            "departures_list": sorted(list(departures))
        }

    def _calculate_component_variances(self):
        components = ["basic_pay", "allowances", "overtime", "deductions", "gross_pay", "net_pay"]
        summary = {}

        for comp in components:
            prior_total = float(self.df_prior[comp].sum()) if comp in self.df_prior.columns and not self.df_prior.empty else 0.0
            curr_total = float(self.df_current[comp].sum()) if comp in self.df_current.columns and not self.df_current.empty else 0.0
            var_amt = curr_total - prior_total
            var_pct = (var_amt / prior_total * 100.0) if prior_total != 0 else 0.0

            summary[comp] = {
                "prior": prior_total,
                "current": curr_total,
                "variance_amount": var_amt,
                "variance_percent": var_pct
            }

        self.component_variances = summary

        # Departmental Summary using vectorized groupby
        dept_summary = []
        p_depts = self.df_prior["department"].unique().tolist() if "department" in self.df_prior.columns and not self.df_prior.empty else []
        c_depts = self.df_current["department"].unique().tolist() if "department" in self.df_current.columns and not self.df_current.empty else []
        all_depts = sorted(list(set(p_depts) | set(c_depts)))

        p_grouped = self.df_prior.groupby("department") if not self.df_prior.empty and "department" in self.df_prior.columns else None
        c_grouped = self.df_current.groupby("department") if not self.df_current.empty and "department" in self.df_current.columns else None

        for dept in all_depts:
            if p_grouped is not None and dept in p_grouped.groups:
                p_sub = p_grouped.get_group(dept)
                p_count = len(p_sub)
                p_gross = float(p_sub["gross_pay"].sum())
                p_net = float(p_sub["net_pay"].sum())
            else:
                p_count = 0
                p_gross = 0.0
                p_net = 0.0

            if c_grouped is not None and dept in c_grouped.groups:
                c_sub = c_grouped.get_group(dept)
                c_count = len(c_sub)
                c_gross = float(c_sub["gross_pay"].sum())
                c_net = float(c_sub["net_pay"].sum())
            else:
                c_count = 0
                c_gross = 0.0
                c_net = 0.0

            gross_var = c_gross - p_gross
            gross_pct = (gross_var / p_gross * 100.0) if p_gross != 0 else 0.0

            net_var = c_net - p_net
            net_pct = (net_var / p_net * 100.0) if p_net != 0 else 0.0

            dept_summary.append({
                "department": dept,
                "prior_headcount": p_count,
                "current_headcount": c_count,
                "prior_gross": p_gross,
                "current_gross": c_gross,
                "gross_variance": gross_var,
                "gross_pct": gross_pct,
                "prior_net": p_net,
                "current_net": c_net,
                "net_variance": net_var,
                "net_pct": net_pct
            })

        self.department_variances = dept_summary

    def _detect_duplicate_ibans(self):
        """
        Rule 1: Duplicate IBAN Cluster Detection.
        Flags cases where multiple distinct Employee IDs share the same bank account.
        Identifies potential payroll clerk embezzlement, ghost payees, or kickback schemes.
        """
        if self.df_current is None or self.df_current.empty or "bank_iban" not in self.df_current.columns:
            return

        iban_groups = {}
        for row in self.df_current.to_dict("records"):
            iban = str(row.get("bank_iban", "")).strip().upper()
            if iban and iban not in ["", "NAN", "NONE", "NULL", "0", "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "UNASSIGNED", "UNKNOWN", "TBD", "-"]:
                iban_groups.setdefault(iban, []).append(row)

        cluster_idx = 1
        for iban, records in iban_groups.items():
            # Check unique employee IDs within this IBAN group
            distinct_emp_ids = set(r["emp_id"] for r in records)
            if len(distinct_emp_ids) > 1:
                total_cluster_net = sum(float(r.get("net_pay", 0.0)) for r in records)
                emp_ids = [r["emp_id"] for r in records]
                emp_names = [r["name"] for r in records]
                depts = [r["department"] for r in records]
                roles = [r["designation"] for r in records]

                has_finance = any("finance" in str(d).lower() or "payroll" in str(r).lower() for d, r in zip(depts, roles))
                suspect_type = "Critical Collusion Risk (Payroll/Finance Linked)" if has_finance else "Multiple Payees on Shared IBAN"

                cluster_info = {
                    "cluster_id": f"IBAN-CLUST-{cluster_idx:03d}",
                    "iban": iban,
                    "record_count": len(records),
                    "emp_ids": ", ".join(emp_ids),
                    "emp_names": ", ".join(emp_names),
                    "departments": ", ".join(set(depts)),
                    "designations": ", ".join(roles),
                    "individual_nets": ", ".join(f"${float(r.get('net_pay', 0.0)):,.2f}" for r in records),
                    "total_net": total_cluster_net,
                    "suspect_classification": suspect_type,
                    "risk_commentary": f"Bank account shared across {len(distinct_emp_ids)} unique employees. High probability of fictitious ghost payees or diverted disbursements."
                }
                self.iban_clusters.append(cluster_info)
                cluster_idx += 1

                for r in records:
                    self.exceptions.append({
                        "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                        "severity": "CRITICAL",
                        "category": "Duplicate IBAN Collision",
                        "emp_id": r["emp_id"],
                        "name": r["name"],
                        "department": r["department"],
                        "designation": r["designation"],
                        "impact_amount": float(r.get("net_pay", 0.0)),
                        "flag_reason": f"IBAN {iban} shared with {len(distinct_emp_ids) - 1} other employee(s): {', '.join([x for x in emp_ids if x != r['emp_id']])}",
                        "recommended_action": "Freeze bank disbursement file immediately. Verify physical employee identity, government ID, and I-9/residency file."
                    })

    def _detect_post_termination_leakage(self):
        """
        Rule 2: Post-Termination Salary Leakage (Zombie Payroll).
        Checks if an employee terminated in the prior register or marked as TERMINATED
        in the HR Movement Log is actively drawn in the current register with Net Pay > 0.
        """
        term_ids = set()

        # Prior month register check
        if not self.df_prior.empty and "status" in self.df_prior.columns:
            for row in self.df_prior.to_dict("records"):
                st = str(row.get("status", "")).strip().lower()
                if any(k in st for k in ["term", "inactive", "exit", "resigned", "dismissed"]):
                    term_ids.add(row["emp_id"])

        # HR Movement Log check
        if self.df_hr is not None and not self.df_hr.empty:
            for hrow in self.df_hr.to_dict("records"):
                evt = str(hrow.get("event_type", "")).strip().upper()
                if any(k in evt for k in ["TERM", "RESIGN", "EXIT", "DISMISS", "LEAV"]):
                    eid = str(hrow.get("emp_id", "")).strip()
                    if eid:
                        term_ids.add(eid)

        for emp_id in term_ids:
            if emp_id in self.curr_map:
                curr_row = self.curr_map[emp_id]
                net_val = float(curr_row.get("net_pay", 0.0))
                gross_val = float(curr_row.get("gross_pay", 0.0))
                if net_val > 0 or gross_val > 0:
                    self.exceptions.append({
                        "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                        "severity": "CRITICAL",
                        "category": "Post-Termination Salary Leakage",
                        "emp_id": emp_id,
                        "name": curr_row.get("name", "Unknown"),
                        "department": curr_row.get("department", "Unassigned"),
                        "designation": curr_row.get("designation", "General Staff"),
                        "impact_amount": net_val,
                        "flag_reason": f"Employee documented as terminated but received current net pay of ${net_val:,.2f} (Gross: ${gross_val:,.2f}).",
                        "recommended_action": "Execute immediate clawback demand with commercial bank. Review payroll master cut-off controls and HR offboarding workflows."
                    })

    def _detect_net_exceeds_gross(self):
        """
        Rule 3: Net Pay > Gross Pay Violation.
        Detects negative deductions or arithmetic inversion where take-home pay exceeds total earnings.
        """
        if self.df_current is None or self.df_current.empty:
            return

        for row in self.df_current.to_dict("records"):
            gross = float(row.get("gross_pay", 0.0))
            net = float(row.get("net_pay", 0.0))
            ded = float(row.get("deductions", 0.0))

            if net > gross or ded < 0:
                self.exceptions.append({
                    "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                    "severity": "HIGH",
                    "category": "Net Pay Exceeds Gross Pay",
                    "emp_id": row["emp_id"],
                    "name": row["name"],
                    "department": row["department"],
                    "designation": row["designation"],
                    "impact_amount": abs(net - gross),
                    "flag_reason": f"Net pay (${net:,.2f}) exceeds gross pay (${gross:,.2f}) with negative deductions of ${ded:,.2f}.",
                    "recommended_action": "Audit deduction subledger and formula references. Confirm absence of unauthorized off-ledger salary supplements."
                })

    def _detect_managerial_overtime(self):
        """
        Rule 4: Managerial Overtime Breaches.
        Flags employees holding managerial, executive, or supervisory titles drawing Overtime Pay.
        """
        if self.df_current is None or self.df_current.empty:
            return

        mgr_keywords = ["manager", "director", "head", "chief", "vp", "president", "controller", "executive"]

        for row in self.df_current.to_dict("records"):
            title = str(row.get("designation", "")).lower()
            dept = str(row.get("department", "")).lower()
            ot = float(row.get("overtime", 0.0))

            if ot > 0:
                if any(kw in title for kw in mgr_keywords) or "executive" in dept:
                    self.exceptions.append({
                        "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                        "severity": "HIGH",
                        "category": "Managerial Overtime Breach",
                        "emp_id": row["emp_id"],
                        "name": row["name"],
                        "department": row["department"],
                        "designation": row["designation"],
                        "impact_amount": ot,
                        "flag_reason": f"Managerial/Exempt role drawing ${ot:,.2f} in overtime compensation contrary to statutory/corporate policy.",
                        "recommended_action": "Require written explanation from authorizing executive. Reconcile timesheet records against biometric turnstile logs."
                    })

    def _detect_salary_spikes(self):
        """
        Rule 5: Unapproved Basic Salary Spikes.
        Flags employees whose basic salary increased by more than the threshold (>15%)
        month-over-month without an approved promotion or salary adjustment in the HR log.
        """
        approved_salary_adjustments = set()
        if self.df_hr is not None and not self.df_hr.empty:
            for hrow in self.df_hr.to_dict("records"):
                evt = str(hrow.get("event_type", "")).strip().upper()
                if any(k in evt for k in ["SALARY", "PROMOTION", "COMPENSATION", "ADJUSTMENT", "GRADE"]):
                    eid = str(hrow.get("emp_id", "")).strip()
                    if eid:
                        approved_salary_adjustments.add(eid)

        for emp_id, curr_row in self.curr_map.items():
            if emp_id in self.prior_map:
                prior_row = self.prior_map[emp_id]
                p_basic = float(prior_row.get("basic_pay", 0.0))
                c_basic = float(curr_row.get("basic_pay", 0.0))

                if p_basic > 0:
                    pct_increase = (c_basic - p_basic) / p_basic
                    if pct_increase > self.spike_threshold:
                        if emp_id not in approved_salary_adjustments:
                            var_val = c_basic - p_basic
                            self.exceptions.append({
                                "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                                "severity": "HIGH",
                                "category": "Unapproved Basic Salary Spike",
                                "emp_id": emp_id,
                                "name": curr_row.get("name", "Unknown"),
                                "department": curr_row.get("department", "Unassigned"),
                                "designation": curr_row.get("designation", "General Staff"),
                                "impact_amount": var_val,
                                "flag_reason": f"Basic salary surged {pct_increase * 100:.1f}% (+${var_val:,.2f}) without corresponding HR log authorization.",
                                "recommended_action": "Inspect signed compensation revision letter, Remuneration Committee minutes, and ERP audit change logs."
                            })

    def _detect_round_number_plugs(self):
        """
        Rule 6: Round-Number Manual Override Net Pay Plug.
        Flags manual round-number lump-sum plugs (e.g. exactly $5,000, $10,000)
        where net pay diverges from calculated gross - deductions, or represents an unusual exact plug.
        """
        if self.df_current is None or self.df_current.empty:
            return

        for row in self.df_current.to_dict("records"):
            gross = float(row.get("gross_pay", 0.0))
            net = float(row.get("net_pay", 0.0))
            ded = float(row.get("deductions", 0.0))

            calc_net = round(gross - ded, 2)
            is_round = (net >= self.round_threshold) and (net % 1000.0 == 0.0)
            is_unbalanced = abs(net - calc_net) > 0.05

            if is_round and is_unbalanced:
                self.exceptions.append({
                    "exception_id": f"EXC-{len(self.exceptions) + 1:03d}",
                    "severity": "MEDIUM",
                    "category": "Manual Round-Number Net Plug",
                    "emp_id": row["emp_id"],
                    "name": row["name"],
                    "department": row["department"],
                    "designation": row["designation"],
                    "impact_amount": abs(net - calc_net),
                    "flag_reason": f"Net pay is an exact round plug of ${net:,.2f} diverging from calculated earnings less deductions (${calc_net:,.2f}).",
                    "recommended_action": "Review payroll adjustment journal entry vouchers and manual ledger debit/credit memos."
                })

    def _build_variance_waterfall(self):
        """Construct granular employee-level variance waterfall table."""
        all_emp_ids = sorted(list(set(self.prior_map.keys()) | set(self.curr_map.keys())))

        # Pre-index exceptions by emp_id for O(1) fast lookup
        exceptions_by_emp = {}
        for e in self.exceptions:
            eid = e.get("emp_id")
            if eid:
                exceptions_by_emp.setdefault(eid, []).append(e)

        for emp_id in all_emp_ids:
            in_prior = emp_id in self.prior_map
            in_curr = emp_id in self.curr_map

            p = self.prior_map[emp_id] if in_prior else None
            c = self.curr_map[emp_id] if in_curr else None

            name = c.get("name") if c is not None else (p.get("name") if p is not None else "Unknown")
            dept = c.get("department") if c is not None else (p.get("department") if p is not None else "Unassigned")
            desig = c.get("designation") if c is not None else (p.get("designation") if p is not None else "General Staff")

            p_basic = float(p.get("basic_pay", 0.0)) if in_prior else 0.0
            c_basic = float(c.get("basic_pay", 0.0)) if in_curr else 0.0
            var_basic = c_basic - p_basic

            p_allow = float(p.get("allowances", 0.0)) if in_prior else 0.0
            c_allow = float(c.get("allowances", 0.0)) if in_curr else 0.0
            var_allow = c_allow - p_allow

            p_ot = float(p.get("overtime", 0.0)) if in_prior else 0.0
            c_ot = float(c.get("overtime", 0.0)) if in_curr else 0.0
            var_ot = c_ot - p_ot

            p_ded = float(p.get("deductions", 0.0)) if in_prior else 0.0
            c_ded = float(c.get("deductions", 0.0)) if in_curr else 0.0
            var_ded = c_ded - p_ded

            p_gross = float(p.get("gross_pay", 0.0)) if in_prior else 0.0
            c_gross = float(c.get("gross_pay", 0.0)) if in_curr else 0.0
            var_gross = c_gross - p_gross

            p_net = float(p.get("net_pay", 0.0)) if in_prior else 0.0
            c_net = float(c.get("net_pay", 0.0)) if in_curr else 0.0
            var_net = c_net - p_net
            pct_net = (var_net / p_net * 100.0) if p_net != 0 else (100.0 if c_net > 0 else 0.0)

            if not in_prior and in_curr:
                mvt_type = "New Hire / Addition"
            elif in_prior and not in_curr:
                mvt_type = "Termination / Departure"
            elif var_net > 0:
                mvt_type = "Pay Increase"
            elif var_net < 0:
                mvt_type = "Pay Decrease"
            else:
                mvt_type = "Unchanged"

            emp_exceptions = exceptions_by_emp.get(emp_id, [])
            audit_flag = ", ".join(sorted(list(set(e["category"] for e in emp_exceptions)))) if emp_exceptions else "CLEAN"

            self.waterfall_records.append({
                "emp_id": emp_id,
                "name": name,
                "department": dept,
                "designation": desig,
                "prior_basic": p_basic,
                "current_basic": c_basic,
                "basic_variance": var_basic,
                "prior_allowances": p_allow,
                "current_allowances": c_allow,
                "allowances_variance": var_allow,
                "prior_overtime": p_ot,
                "current_overtime": c_ot,
                "overtime_variance": var_ot,
                "prior_deductions": p_ded,
                "current_deductions": c_ded,
                "deductions_variance": var_ded,
                "prior_gross": p_gross,
                "current_gross": c_gross,
                "gross_variance": var_gross,
                "prior_net": p_net,
                "current_net": c_net,
                "net_variance": var_net,
                "net_pct_change": pct_net,
                "movement_type": mvt_type,
                "audit_flag": audit_flag
            })

    def _compile_executive_kpis(self):
        curr_payroll_total = float(self.df_current["gross_pay"].sum()) if not self.df_current.empty and "gross_pay" in self.df_current.columns else 0.0
        prior_payroll_total = float(self.df_prior["gross_pay"].sum()) if not self.df_prior.empty and "gross_pay" in self.df_prior.columns else 0.0
        payroll_var = curr_payroll_total - prior_payroll_total
        payroll_var_pct = (payroll_var / prior_payroll_total * 100.0) if prior_payroll_total != 0 else 0.0

        curr_net_total = float(self.df_current["net_pay"].sum()) if not self.df_current.empty and "net_pay" in self.df_current.columns else 0.0
        prior_net_total = float(self.df_prior["net_pay"].sum()) if not self.df_prior.empty and "net_pay" in self.df_prior.columns else 0.0
        net_var = curr_net_total - prior_net_total
        net_var_pct = (net_var / prior_net_total * 100.0) if prior_net_total != 0 else 0.0

        total_exposure = sum(e["impact_amount"] for e in self.exceptions)

        crit_count = sum(1 for e in self.exceptions if e["severity"] == "CRITICAL")
        high_count = sum(1 for e in self.exceptions if e["severity"] == "HIGH")
        med_count = sum(1 for e in self.exceptions if e["severity"] == "MEDIUM")

        self.kpi_summary = {
            "current_gross_payroll": curr_payroll_total,
            "prior_gross_payroll": prior_payroll_total,
            "payroll_variance": payroll_var,
            "payroll_variance_pct": payroll_var_pct,
            "current_net_payroll": curr_net_total,
            "prior_net_payroll": prior_net_total,
            "net_variance": net_var,
            "net_variance_pct": net_var_pct,
            "active_headcount": len(self.df_current),
            "prior_headcount": len(self.df_prior),
            "net_headcount_change": len(self.df_current) - len(self.df_prior),
            "total_exceptions": len(self.exceptions),
            "critical_exceptions": crit_count,
            "high_exceptions": high_count,
            "medium_exceptions": med_count,
            "total_exposure_at_risk": total_exposure,
            "duplicate_iban_clusters": len(self.iban_clusters)
        }

# ==============================================================================
# EXCEL AUDIT REPORT GENERATOR (OPENPYXL)
# ==============================================================================

class ExcelAuditReportBuilder:
    def __init__(self, engine, output_path):
        self.engine = engine
        self.output_path = output_path
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)

        # Visual styling palette
        self.navy_header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        self.navy_section_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
        self.sub_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        self.kpi_bg_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        self.zebra_fill = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid")

        # Severity fills
        self.crit_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        self.crit_font = Font(name="Calibri", size=10, bold=True, color="C00000")
        self.high_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        self.high_font = Font(name="Calibri", size=10, bold=True, color="B25900")
        self.med_fill = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")
        self.med_font = Font(name="Calibri", size=10, bold=True, color="0E6655")

        self.header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.title_font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
        self.bold_font = Font(name="Calibri", size=10, bold=True, color="000000")
        self.data_font = Font(name="Calibri", size=10, color="000000")
        self.kpi_val_font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
        self.kpi_lbl_font = Font(name="Calibri", size=9, bold=True, color="595959")

        self.thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )
        self.total_border = Border(
            top=Side(style="thin", color="000000"),
            bottom=Side(style="double", color="000000")
        )

    def build_report(self):
        self._create_executive_dashboard()
        self._create_exception_register()
        self._create_variance_waterfall()
        self._create_iban_clusters_tab()
        self.wb.save(self.output_path)
        return self.output_path

    def _auto_fit_columns(self, ws, min_width=12, padding=3):
        for col in ws.columns:
            max_len = 0
            for cell in col:
                val_str = str(cell.value or "")
                if cell.row < 4 and len(val_str) > 40:
                    continue
                max_len = max(max_len, len(val_str))
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + padding, min_width)

    def _create_executive_dashboard(self):
        ws = self.wb.create_sheet(title="Executive_Dashboard")
        ws.views.sheetView[0].showGridLines = True
        kpis = self.engine.kpi_summary

        # Title Block
        ws.merge_cells("A1:H1")
        title_cell = ws["A1"]
        title_cell.value = "PAYROLL FORENSIC VARIANCE & CONTINUOUS INTEGRITY DASHBOARD"
        title_cell.font = self.title_font
        title_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A2:H2")
        sub_cell = ws["A2"]
        sub_cell.value = f"Veritux Consulting Network | Continuous Controls Monitoring Engine | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
        ws.row_dimensions[2].height = 18

        # KPI Block (Row 4 to 5)
        kpi_cards = [
            ("Current Gross Payroll", f"${kpis['current_gross_payroll']:,.2f}", "A", "B"),
            ("Net MoM Payroll Variance", f"${kpis['net_variance']:+,.2f} ({kpis['net_variance_pct']:+.1f}%)", "C", "D"),
            ("Active Headcount", f"{kpis['active_headcount']:,} ({kpis['net_headcount_change']:+d})", "E", "F"),
            ("Forensic Exceptions", f"{kpis['total_exceptions']} Flagged (${kpis['total_exposure_at_risk']:,.2f})", "G", "H"),
        ]

        ws.row_dimensions[4].height = 18
        ws.row_dimensions[5].height = 28

        for title, val, start_col, end_col in kpi_cards:
            ws.merge_cells(f"{start_col}4:{end_col}4")
            ws.merge_cells(f"{start_col}5:{end_col}5")

            c_lbl = ws[f"{start_col}4"]
            c_lbl.value = title.upper()
            c_lbl.font = self.kpi_lbl_font
            c_lbl.fill = self.kpi_bg_fill
            c_lbl.alignment = Alignment(horizontal="center", vertical="center")
            c_lbl.border = self.thin_border

            c_val = ws[f"{start_col}5"]
            c_val.value = val
            c_val.font = self.kpi_val_font
            c_val.fill = self.kpi_bg_fill
            c_val.alignment = Alignment(horizontal="center", vertical="center")
            c_val.border = self.thin_border

        # Section 1: Headcount Reconciliation Bridge (Row 7)
        ws.cell(row=7, column=1, value="1. HEADCOUNT RECONCILIATION BRIDGE").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
        ws.row_dimensions[7].height = 24

        bridge_headers = ["Metric / Movement Category", "Headcount", "% of Starting Base", "Audit Commentary"]
        ws.row_dimensions[8].height = 22
        for col_idx, h_text in enumerate(bridge_headers, start=1):
            cell = ws.cell(row=8, column=col_idx, value=h_text)
            cell.fill = self.navy_header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        bridge = self.engine.headcount_bridge
        start_cnt = bridge["starting_headcount"]

        add_list = bridge["additions_list"]
        dep_list = bridge["departures_list"]
        add_more = f" and {len(add_list) - 8} more" if len(add_list) > 8 else ""
        dep_more = f" and {len(dep_list) - 8} more" if len(dep_list) > 8 else ""
        add_str = f"Employees added to active register ({', '.join(add_list[:8])}{add_more})" if add_list else "No new employee additions"
        dep_str = f"Employees dropped from active register ({', '.join(dep_list[:8])}{dep_more})" if dep_list else "No employee departures"

        bridge_rows = [
            ("Prior Month Active Register", start_cnt, 100.0 if start_cnt > 0 else 0.0, "Baseline starting workforce"),
            ("Additions / New Hires", bridge["additions_count"], (bridge["additions_count"] / start_cnt * 100.0) if start_cnt > 0 else 0.0, add_str),
            ("Departures / Terminations", -bridge["departures_count"], (-bridge["departures_count"] / start_cnt * 100.0) if start_cnt > 0 else 0.0, dep_str),
            ("Net Headcount Movement", bridge["net_change"], (bridge["net_change"] / start_cnt * 100.0) if start_cnt > 0 else 0.0, "Net monthly expansion/contraction"),
            ("Current Month Active Register", bridge["ending_headcount"], (bridge["ending_headcount"] / start_cnt * 100.0) if start_cnt > 0 else 0.0, "Verified active headcount closing position")
        ]

        curr_r = 9
        for label, count, pct, comment in bridge_rows:
            ws.row_dimensions[curr_r].height = 19
            is_bold = "Current Month" in label or "Net Headcount" in label
            fnt = self.bold_font if is_bold else self.data_font

            c1 = ws.cell(row=curr_r, column=1, value=label)
            c2 = ws.cell(row=curr_r, column=2, value=count)
            c3 = ws.cell(row=curr_r, column=3, value=pct / 100.0)
            c4 = ws.cell(row=curr_r, column=4, value=comment)

            c1.font = fnt; c1.border = self.thin_border
            c2.font = fnt; c2.border = self.thin_border; c2.number_format = "#,##0"; c2.alignment = Alignment(horizontal="right")
            c3.font = fnt; c3.border = self.thin_border; c3.number_format = "0.0%"; c3.alignment = Alignment(horizontal="right")
            c4.font = fnt; c4.border = self.thin_border

            if "Current Month" in label:
                c1.border = self.total_border; c2.border = self.total_border; c3.border = self.total_border; c4.border = self.total_border

            curr_r += 1

        # Section 2: Payroll Component Variance Table
        curr_r += 1
        ws.cell(row=curr_r, column=1, value="2. PAYROLL COMPONENT VARIANCE BREAKDOWN").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
        ws.row_dimensions[curr_r].height = 24
        curr_r += 1

        comp_headers = ["Compensation Component", "Prior Month ($)", "Current Month ($)", "Variance ($)", "MoM Change (%)", "Variance Classification"]
        ws.row_dimensions[curr_r].height = 22
        for col_idx, h_text in enumerate(comp_headers, start=1):
            cell = ws.cell(row=curr_r, column=col_idx, value=h_text)
            cell.fill = self.navy_header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        curr_r += 1
        comp_labels = [
            ("basic_pay", "Basic Salary"),
            ("allowances", "Allowances (Housing, Transport)"),
            ("overtime", "Overtime Compensation"),
            ("gross_pay", "Total Gross Earnings"),
            ("deductions", "Statutory & Voluntary Deductions"),
            ("net_pay", "Total Net Disbursement")
        ]

        for comp_key, comp_title in comp_labels:
            ws.row_dimensions[curr_r].height = 19
            v = self.engine.component_variances[comp_key]
            is_bold = "Total" in comp_title
            fnt = self.bold_font if is_bold else self.data_font

            classification = "Favorable" if (v["variance_amount"] <= 0 and comp_key != "deductions") else "Expansionary"

            c1 = ws.cell(row=curr_r, column=1, value=comp_title)
            c2 = ws.cell(row=curr_r, column=2, value=v["prior"])
            c3 = ws.cell(row=curr_r, column=3, value=v["current"])
            c4 = ws.cell(row=curr_r, column=4, value=v["variance_amount"])
            c5 = ws.cell(row=curr_r, column=5, value=v["variance_percent"] / 100.0)
            c6 = ws.cell(row=curr_r, column=6, value=classification)

            c1.font = fnt; c1.border = self.thin_border
            c2.font = fnt; c2.border = self.thin_border; c2.number_format = "$#,##0.00"; c2.alignment = Alignment(horizontal="right")
            c3.font = fnt; c3.border = self.thin_border; c3.number_format = "$#,##0.00"; c3.alignment = Alignment(horizontal="right")
            c4.font = fnt; c4.border = self.thin_border; c4.number_format = "$#,##0.00"; c4.alignment = Alignment(horizontal="right")
            c5.font = fnt; c5.border = self.thin_border; c5.number_format = "0.0%"; c5.alignment = Alignment(horizontal="right")
            c6.font = fnt; c6.border = self.thin_border; c6.alignment = Alignment(horizontal="center")

            if is_bold:
                for cell_obj in [c1, c2, c3, c4, c5, c6]:
                    cell_obj.border = self.total_border

            curr_r += 1

        # Section 3: Departmental Variance Breakdown
        curr_r += 1
        ws.cell(row=curr_r, column=1, value="3. DEPARTMENTAL COST CENTER SUMMARY").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
        ws.row_dimensions[curr_r].height = 24
        curr_r += 1

        dept_headers = ["Cost Center / Department", "Prior Headcount", "Current Headcount", "Prior Net ($)", "Current Net ($)", "Net Variance ($)", "MoM Net Change (%)"]
        ws.row_dimensions[curr_r].height = 22
        for col_idx, h_text in enumerate(dept_headers, start=1):
            cell = ws.cell(row=curr_r, column=col_idx, value=h_text)
            cell.fill = self.navy_header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        curr_r += 1
        for d in self.engine.department_variances:
            ws.row_dimensions[curr_r].height = 19
            c1 = ws.cell(row=curr_r, column=1, value=d["department"])
            c2 = ws.cell(row=curr_r, column=2, value=d["prior_headcount"])
            c3 = ws.cell(row=curr_r, column=3, value=d["current_headcount"])
            c4 = ws.cell(row=curr_r, column=4, value=d["prior_net"])
            c5 = ws.cell(row=curr_r, column=5, value=d["current_net"])
            c6 = ws.cell(row=curr_r, column=6, value=d["net_variance"])
            c7 = ws.cell(row=curr_r, column=7, value=d["net_pct"] / 100.0)

            for cell_obj in [c1, c2, c3, c4, c5, c6, c7]:
                cell_obj.font = self.data_font
                cell_obj.border = self.thin_border

            c2.number_format = "#,##0"; c2.alignment = Alignment(horizontal="right")
            c3.number_format = "#,##0"; c3.alignment = Alignment(horizontal="right")
            c4.number_format = "$#,##0.00"; c4.alignment = Alignment(horizontal="right")
            c5.number_format = "$#,##0.00"; c5.alignment = Alignment(horizontal="right")
            c6.number_format = "$#,##0.00"; c6.alignment = Alignment(horizontal="right")
            c7.number_format = "0.0%"; c7.alignment = Alignment(horizontal="right")

            curr_r += 1

        self._auto_fit_columns(ws, min_width=14)

    def _create_exception_register(self):
        ws = self.wb.create_sheet(title="Exception_Register")
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"

        headers = [
            "Exception ID", "Severity", "Anomaly Category", "Employee ID",
            "Employee Name", "Department", "Designation", "Exposure Amount ($)",
            "Flag Reason / Audit Finding", "Recommended Action", "Management Response",
            "Audit Disposition", "Auditor Signoff", "Follow-up Due Date"
        ]

        ws.row_dimensions[1].height = 26
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h_text)
            cell.fill = self.navy_header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, exc in enumerate(self.engine.exceptions, start=2):
            ws.row_dimensions[row_idx].height = 20
            sev = exc["severity"]

            ws.cell(row=row_idx, column=1, value=exc["exception_id"]).alignment = Alignment(horizontal="center")

            sev_cell = ws.cell(row=row_idx, column=2, value=sev)
            sev_cell.alignment = Alignment(horizontal="center")
            if sev == "CRITICAL":
                sev_cell.fill = self.crit_fill
                sev_cell.font = self.crit_font
            elif sev == "HIGH":
                sev_cell.fill = self.high_fill
                sev_cell.font = self.high_font
            else:
                sev_cell.fill = self.med_fill
                sev_cell.font = self.med_font

            ws.cell(row=row_idx, column=3, value=exc["category"])
            ws.cell(row=row_idx, column=4, value=exc["emp_id"]).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=5, value=exc["name"])
            ws.cell(row=row_idx, column=6, value=exc["department"])
            ws.cell(row=row_idx, column=7, value=exc["designation"])

            amt_cell = ws.cell(row=row_idx, column=8, value=exc["impact_amount"])
            amt_cell.number_format = "$#,##0.00"
            amt_cell.alignment = Alignment(horizontal="right")

            ws.cell(row=row_idx, column=9, value=exc["flag_reason"])
            ws.cell(row=row_idx, column=10, value=exc["recommended_action"])

            ws.cell(row=row_idx, column=11, value="Pending Management Inquiry")
            ws.cell(row=row_idx, column=12, value="OPEN")
            ws.cell(row=row_idx, column=13, value="")
            ws.cell(row=row_idx, column=14, value="")

            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=c)
                cell.border = self.thin_border
                if c != 2:
                    cell.font = self.data_font

        self._auto_fit_columns(ws, min_width=12)

    def _create_variance_waterfall(self):
        ws = self.wb.create_sheet(title="Variance_Waterfall")
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"

        headers = [
            "Employee ID", "Employee Name", "Department", "Designation",
            "Prior Basic ($)", "Current Basic ($)", "Basic Var ($)",
            "Prior Allow ($)", "Current Allow ($)", "Allow Var ($)",
            "Prior OT ($)", "Current OT ($)", "OT Var ($)",
            "Prior Deduct ($)", "Current Deduct ($)", "Deduct Var ($)",
            "Prior Gross ($)", "Current Gross ($)", "Gross Var ($)",
            "Prior Net ($)", "Current Net ($)", "Net Var ($)", "Net % Change",
            "Movement Category", "Audit Exception Flag"
        ]

        ws.row_dimensions[1].height = 26
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h_text)
            cell.fill = self.navy_header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        num_currency_cols = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]

        for row_idx, r in enumerate(self.engine.waterfall_records, start=2):
            ws.row_dimensions[row_idx].height = 19
            is_even = (row_idx % 2 == 0)

            ws.cell(row=row_idx, column=1, value=r["emp_id"]).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=2, value=r["name"])
            ws.cell(row=row_idx, column=3, value=r["department"])
            ws.cell(row=row_idx, column=4, value=r["designation"])

            ws.cell(row=row_idx, column=5, value=r["prior_basic"])
            ws.cell(row=row_idx, column=6, value=r["current_basic"])
            ws.cell(row=row_idx, column=7, value=r["basic_variance"])

            ws.cell(row=row_idx, column=8, value=r["prior_allowances"])
            ws.cell(row=row_idx, column=9, value=r["current_allowances"])
            ws.cell(row=row_idx, column=10, value=r["allowances_variance"])

            ws.cell(row=row_idx, column=11, value=r["prior_overtime"])
            ws.cell(row=row_idx, column=12, value=r["current_overtime"])
            ws.cell(row=row_idx, column=13, value=r["overtime_variance"])

            ws.cell(row=row_idx, column=14, value=r["prior_deductions"])
            ws.cell(row=row_idx, column=15, value=r["current_deductions"])
            ws.cell(row=row_idx, column=16, value=r["deductions_variance"])

            ws.cell(row=row_idx, column=17, value=r["prior_gross"])
            ws.cell(row=row_idx, column=18, value=r["current_gross"])
            ws.cell(row=row_idx, column=19, value=r["gross_variance"])

            ws.cell(row=row_idx, column=20, value=r["prior_net"])
            ws.cell(row=row_idx, column=21, value=r["current_net"])
            ws.cell(row=row_idx, column=22, value=r["net_variance"])

            pct_cell = ws.cell(row=row_idx, column=23, value=r["net_pct_change"] / 100.0)
            pct_cell.number_format = "0.0%"
            pct_cell.alignment = Alignment(horizontal="right")

            ws.cell(row=row_idx, column=24, value=r["movement_type"]).alignment = Alignment(horizontal="center")

            flag_cell = ws.cell(row=row_idx, column=25, value=r["audit_flag"])
            if r["audit_flag"] != "CLEAN":
                flag_cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")
                flag_cell.fill = self.crit_fill

            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=c)
                cell.border = self.thin_border
                if cell.font != self.crit_font:
                    cell.font = self.data_font
                if is_even and r["audit_flag"] == "CLEAN":
                    cell.fill = self.zebra_fill

                if c in num_currency_cols:
                    cell.number_format = "$#,##0.00"
                    cell.alignment = Alignment(horizontal="right")

        self._auto_fit_columns(ws, min_width=12)

    def _create_iban_clusters_tab(self):
        ws = self.wb.create_sheet(title="Duplicate_IBAN_Clusters")
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"

        headers = [
            "Cluster ID", "Bank IBAN / Account", "Record Count", "Employee IDs",
            "Employee Names", "Departments", "Designations", "Individual Net Amounts",
            "Total Disbursed ($)", "Classification", "Risk Commentary"
        ]

        ws.row_dimensions[1].height = 26
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h_text)
            cell.fill = self.navy_header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, cl in enumerate(self.engine.iban_clusters, start=2):
            ws.row_dimensions[row_idx].height = 22

            ws.cell(row=row_idx, column=1, value=cl["cluster_id"]).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=2, value=cl["iban"]).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=3, value=cl["record_count"]).alignment = Alignment(horizontal="center")
            ws.cell(row=row_idx, column=4, value=cl["emp_ids"])
            ws.cell(row=row_idx, column=5, value=cl["emp_names"])
            ws.cell(row=row_idx, column=6, value=cl["departments"])
            ws.cell(row=row_idx, column=7, value=cl["designations"])
            ws.cell(row=row_idx, column=8, value=cl["individual_nets"])

            tot_cell = ws.cell(row=row_idx, column=9, value=cl["total_net"])
            tot_cell.number_format = "$#,##0.00"
            tot_cell.alignment = Alignment(horizontal="right")
            tot_cell.font = self.bold_font

            c_cell = ws.cell(row=row_idx, column=10, value=cl["suspect_classification"])
            c_cell.fill = self.crit_fill
            c_cell.font = self.crit_font
            c_cell.alignment = Alignment(horizontal="center")

            ws.cell(row=row_idx, column=11, value=cl["risk_commentary"])

            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=c)
                cell.border = self.thin_border
                if c not in [9, 10]:
                    cell.font = self.data_font

        self._auto_fit_columns(ws, min_width=14)

# ==============================================================================
# TERMINAL DISPLAY & TELEMETRY
# ==============================================================================

def display_terminal_summary(engine, output_report_path):
    kpis = engine.kpi_summary
    bridge = engine.headcount_bridge

    if HAS_RICH and console is not None:
        console.print()
        console.print("[bold white on #1F4E79]  PAYROLL FORENSIC VARIANCE & CONTINUOUS INTEGRITY ENGINE  [/bold white on #1F4E79]")
        console.print("[dim]Veritux Consulting Network | Continuous Controls Monitoring Pipeline[/dim]\n")

        summary_tbl = Table(title="Executive Reconciliation & Variance Summary", header_style="bold cyan", show_lines=True)
        summary_tbl.add_column("Telemetry Metric", style="bold")
        summary_tbl.add_column("Prior Month", justify="right")
        summary_tbl.add_column("Current Month", justify="right")
        summary_tbl.add_column("Variance ($ / #)", justify="right")
        summary_tbl.add_column("MoM Change (%)", justify="right")

        summary_tbl.add_row(
            "Active Workforce Headcount",
            f"{bridge['starting_headcount']:,}",
            f"{bridge['ending_headcount']:,}",
            f"{bridge['net_change']:+d}",
            f"{(bridge['net_change'] / bridge['starting_headcount'] * 100.0) if bridge['starting_headcount'] > 0 else 0.0:+.1f}%"
        )
        summary_tbl.add_row(
            "Total Gross Payroll Expense",
            f"${kpis['prior_gross_payroll']:,.2f}",
            f"${kpis['current_gross_payroll']:,.2f}",
            f"${kpis['payroll_variance']:+,.2f}",
            f"{kpis['payroll_variance_pct']:+.1f}%"
        )
        summary_tbl.add_row(
            "Total Net Disbursement",
            f"${kpis['prior_net_payroll']:,.2f}",
            f"${kpis['current_net_payroll']:,.2f}",
            f"${kpis['net_variance']:+,.2f}",
            f"{kpis['net_variance_pct']:+.1f}%"
        )
        console.print(summary_tbl)
        console.print()

        exc_tbl = Table(title=f"Forensic Red Flag Exceptions Detected ({kpis['total_exceptions']} Total)", header_style="bold red", show_lines=True)
        exc_tbl.add_column("ID", justify="center", style="bold")
        exc_tbl.add_column("Severity", justify="center")
        exc_tbl.add_column("Category", style="yellow")
        exc_tbl.add_column("Emp ID", justify="center")
        exc_tbl.add_column("Name")
        exc_tbl.add_column("Department")
        exc_tbl.add_column("Exposure ($)", justify="right", style="bold red")
        exc_tbl.add_column("Flag Reason")

        for exc in engine.exceptions:
            sev_style = "bold red" if exc["severity"] == "CRITICAL" else ("bold yellow" if exc["severity"] == "HIGH" else "cyan")
            exc_tbl.add_row(
                exc["exception_id"],
                f"[{sev_style}]{exc['severity']}[/{sev_style}]",
                exc["category"],
                exc["emp_id"],
                exc["name"],
                exc["department"],
                f"${exc['impact_amount']:,.2f}",
                exc["flag_reason"]
            )
        console.print(exc_tbl)
        console.print()

        console.print(f"[bold green]✔ Excel Audit Report compiled successfully:[/bold green] [underline]{output_report_path}[/underline]")
        console.print(f"[dim]Total Exposure at Risk: ${kpis['total_exposure_at_risk']:,.2f} | Critical: {kpis['critical_exceptions']} | High: {kpis['high_exceptions']} | Medium: {kpis['medium_exceptions']}[/dim]\n")

    else:
        print("\n" + "=" * 80)
        print("  PAYROLL FORENSIC VARIANCE & CONTINUOUS INTEGRITY ENGINE")
        print("  Veritux Consulting Network | Continuous Controls Monitoring Pipeline")
        print("=" * 80)
        print(f"Headcount: {bridge['starting_headcount']} -> {bridge['ending_headcount']} (Net: {bridge['net_change']:+d})")
        print(f"Gross Payroll: ${kpis['prior_gross_payroll']:,.2f} -> ${kpis['current_gross_payroll']:,.2f} (Var: ${kpis['payroll_variance']:+,.2f}, {kpis['payroll_variance_pct']:+.1f}%)")
        print(f"Net Payroll:   ${kpis['prior_net_payroll']:,.2f} -> ${kpis['current_net_payroll']:,.2f} (Var: ${kpis['net_variance']:+,.2f}, {kpis['net_variance_pct']:+.1f}%)")
        print(f"Forensic Exceptions Flagged: {kpis['total_exceptions']} (Critical: {kpis['critical_exceptions']}, High: {kpis['high_exceptions']}, Medium: {kpis['medium_exceptions']})")
        print(f"Total Exposure at Risk: ${kpis['total_exposure_at_risk']:,.2f}")
        print("-" * 80)
        for exc in engine.exceptions:
            print(f"[{exc['severity']}] {exc['exception_id']} - {exc['category']} - {exc['emp_id']} ({exc['name']}, {exc['department']}): ${exc['impact_amount']:,.2f}")
            print(f"       Reason: {exc['flag_reason']}")
        print("=" * 80)
        print(f"Report written to: {output_report_path}\n")

# ==============================================================================
# CLI & INTERACTIVE ORCHESTRATION
# ==============================================================================

def prompt_interactive_paths():
    """Prompt user for files with automatic quote stripping for drag-and-drop support."""
    print("\n" + "=" * 70)
    print("  Payroll Forensic Variance and Integrity Engine - Interactive Mode")
    print("  Drag and drop files directly from Windows Explorer into this window.")
    print("=" * 70 + "\n")

    while True:
        p_raw = input("Enter PRIOR Month Payroll Register path (.xlsx / .csv): ").strip()
        p_clean = clean_file_path(p_raw)
        if os.path.exists(p_clean):
            prior_path = p_clean
            break
        print(f"  [ERROR] File does not exist: {p_clean}. Please try again.\n")

    while True:
        c_raw = input("Enter CURRENT Month Payroll Register path (.xlsx / .csv): ").strip()
        c_clean = clean_file_path(c_raw)
        if os.path.exists(c_clean):
            current_path = c_clean
            break
        print(f"  [ERROR] File does not exist: {c_clean}. Please try again.\n")

    hr_raw = input("Enter optional HR Movement Log path (.xlsx / .csv, or Enter to skip): ").strip()
    hr_path = clean_file_path(hr_raw) if hr_raw else None
    if hr_path and not os.path.exists(hr_path):
        print(f"  [WARNING] HR Movement file '{hr_path}' not found. Proceeding without HR log.")
        hr_path = None

    default_out = f"Payroll_Forensic_Audit_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out_raw = input(f"Enter Output Report path [Default: {default_out}]: ").strip()
    out_path = clean_file_path(out_raw) if out_raw else default_out

    return prior_path, current_path, hr_path, out_path

def main():
    parser = argparse.ArgumentParser(
        description="Product 13: Payroll Forensic Variance and Integrity Engine"
    )
    parser.add_argument("--prior", "-p", help="Path to Prior Month Payroll Register (.xlsx / .csv)")
    parser.add_argument("--current", "-c", help="Path to Current Month Payroll Register (.xlsx / .csv)")
    parser.add_argument("--hr", help="Optional path to HR Movement Log (.xlsx / .csv)")
    parser.add_argument("--output", "-o", help="Path for generated Excel Forensic Audit Report (.xlsx)")
    parser.add_argument("--non-interactive", "-n", action="store_true", help="Run without interactive prompts")
    parser.add_argument("--spike-threshold", "-s", type=float, default=0.15, help="Basic salary spike threshold (default: 0.15 for 15%%)")
    parser.add_argument("--round-threshold", "-r", type=float, default=1000.0, help="Round number plug threshold (default: 1000.0)")

    args = parser.parse_args()

    if args.prior and args.current:
        prior_path = clean_file_path(args.prior)
        current_path = clean_file_path(args.current)
        hr_path = clean_file_path(args.hr) if args.hr else None
        output_path = clean_file_path(args.output) if args.output else f"Payroll_Forensic_Audit_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    elif args.non_interactive:
        print("[ERROR] Batch mode requires both --prior and --current arguments.")
        sys.exit(1)
    else:
        prior_path, current_path, hr_path, output_path = prompt_interactive_paths()

    print(f"\n[1/4] Ingesting and mapping Prior Month Payroll Register: {prior_path}")
    raw_prior = read_data_file(prior_path)
    df_prior, map_prior = map_dataframe_schema(raw_prior, CANONICAL_FIELDS, "Prior Month")
    print(f"      Mapped {len(df_prior)} records across columns: {list(map_prior.keys())}")

    print(f"[2/4] Ingesting and mapping Current Month Payroll Register: {current_path}")
    raw_current = read_data_file(current_path)
    df_current, map_current = map_dataframe_schema(raw_current, CANONICAL_FIELDS, "Current Month")
    print(f"      Mapped {len(df_current)} records across columns: {list(map_current.keys())}")

    df_hr = None
    if hr_path:
        print(f"[3/4] Ingesting optional HR Movement Log: {hr_path}")
        raw_hr = read_data_file(hr_path)
        df_hr, map_hr = map_dataframe_schema(raw_hr, HR_LOG_FIELDS, "HR Log")
        print(f"      Mapped {len(df_hr)} HR lifecycle events.")
    else:
        print(f"[3/4] Skipping optional HR Movement Log.")

    print(f"[4/4] Executing forensic reconciliation and anomaly detection heuristics...")
    engine = PayrollForensicEngine(
        df_prior=df_prior,
        df_current=df_current,
        df_hr=df_hr,
        spike_threshold=args.spike_threshold,
        round_threshold=args.round_threshold
    )
    engine.run_all_checks()

    # Generate Excel Report
    report_builder = ExcelAuditReportBuilder(engine, output_path)
    final_report = report_builder.build_report()

    # Display Terminal Telemetry
    display_terminal_summary(engine, final_report)

if __name__ == "__main__":
    main()
