# Contributing to the Payroll Forensic Variance and Integrity Engine

We welcome contributions from internal audit professionals, forensic data analysts, and Python developers.

---

## Code of Conduct and Operating Philosophy

This project prioritizes deterministic accuracy, zero-infrastructure simplicity, and fiduciarily sound audit analytics. Contributions should preserve these core characteristics:

- Keep dependencies minimal. The core engine relies strictly on pandas, openpyxl, and rich. Avoid introducing heavy framework dependencies or database requirements.
- Maintain 100% macro-free Excel generation. Output workbooks must remain native `.xlsx` files without VBA macros or embedded code, ensuring immediate clearance by corporate security gateways.
- Ensure cross-platform execution. The code must execute reliably on Windows, macOS, and Linux across supported Python versions (3.9 through 3.13).
- Preserve high-throughput efficiency. Vectorized pandas operations should be preferred over iterative row looping so that 10,000+ employee registers process in under 5 seconds.

---

## Development Setup

1. Fork and clone the repository:

```bash
git clone https://github.com/majidrajpar/payroll-forensic-variance-engine.git
cd payroll-forensic-variance-engine
```

2. Create and activate a virtual environment:

```bash
python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On Linux/macOS:
source .venv/bin/activate
```

3. Install the project in editable mode with development dependencies:

```bash
pip install -r requirements.txt
pip install -e .
```

---

## Proposing New Forensic Rules

When proposing a new audit heuristic or fraud detection check:

1. Document the fiduciary risk scenario, including the scheme structure and how perpetrators typically obscure it in corporate payrolls.
2. Formulate the mathematical or logical rule to prevent false positives on legitimate compensation adjustments (e.g., retro-pay adjustments, pro-rated mid-month joiners).
3. Implement the rule method within the `PayrollForensicEngine` class in `payroll_audit_engine.py`.
4. Integrate the resulting exceptions into both the terminal telemetry summary and the Excel Exception Register worksheet.
5. Add corresponding test cases in `test_stress_payroll_engine.py` verifying both detection accuracy and non-regression on clean data.

---

## Running Tests

All pull requests must pass the complete automated test suite before merging. Run tests locally using Python's standard unittest runner:

```bash
python test_stress_payroll_engine.py
```

Ensure that:
- All 9 baseline tests pass with zero errors.
- The 10,000-employee stress benchmark completes well within the 5.0-second performance ceiling.
- Schema normalization tests succeed across mixed column naming conventions.

---

## Submitting Pull Requests

1. Create a focused topic branch from `main`:

```bash
git checkout -b feature/new-forensic-heuristic
```

2. Commit changes with clear, descriptive commit messages outlining the operational change and why it was necessary.
3. Push to your fork and submit a Pull Request against the `main` branch.
4. Describe the specific payroll scenarios tested, including synthetic edge cases or schema variants tested against.

---

## Security Vulnerabilities

If you identify a security flaw or unhandled edge case that could lead to financial miscalculation or data leakage, please report it directly to Majid Mumtaz via LinkedIn (https://www.linkedin.com/in/majid-m-4b097118/) or email `majid@veritux.com`.
