# 💰 HR Compensation Automation: CTC, Payslip & Budget Modeling Tool

![Excel](https://img.shields.io/badge/Excel-VBA%20Enabled-217346)
![Tabs](https://img.shields.io/badge/Tabs-13-blue)
![Type](https://img.shields.io/badge/Type-Macro--Enabled%20(.xlsm)-orange)

An Excel-based HR compensation engine that turns a single employee master list into individual CTC breakdowns, payslips, appraisal planning, budget simulation and salary benchmarking, all driven by formulas from one editable set of policy assumptions.

> ⚠️ **Before uploading:** this workbook's sample data (`01_Employee_Master` and other tabs) contains named individuals and exact salary figures. If these are real people, **do not upload this file publicly as-is**. Replace the names and CTC values with clearly fictional placeholders (e.g. "Employee A", round numbers) before pushing to GitHub. See [Data Privacy](#-data-privacy) below.

---

## 📌 Table of Contents
1. [What This Tool Does](#-what-this-tool-does)
2. [Workbook Structure](#-workbook-structure)
3. [How the Calculation Works](#-how-the-calculation-works)
4. [Data Privacy](#-data-privacy)
5. [Scope & Assumptions](#-scope--assumptions)
6. [How to Use](#-how-to-use)
7. [Tech Stack](#-tech-stack)
8. [Author](#-author)

---

## 🎯 What This Tool Does

HR and payroll teams often rebuild the same salary breakdown by hand for every employee, every appraisal cycle. This workbook automates that end to end:

- Stores every employee's CTC and role in one master list.
- Lets a user pick any Employee ID and instantly see their full salary breakup (basic, HRA, PF, gratuity, deductions, net pay).
- Generates a formatted CTC statement and a monthly payslip for that employee.
- Validates every calculation with built-in audit checks (does the breakup reconcile back to the original CTC?).
- Models salary increments across the whole workforce, department-wise budgets, and "what-if" increment scenarios against a fixed budget.
- Benchmarks each role's pay against a salary band and flags anyone outside range.

## 📁 Workbook Structure

| # | Tab | Purpose |
|---|---|---|
| 01 | `Employee_Master` | Source of truth: every employee's ID, name, department, designation, CTC, variable pay % and status. |
| 02 | `Employee_Calculator` | Pick an Employee ID here; every downstream tab recalculates for that person. |
| 03 | `Policy_Assumptions` | The editable levers: Basic %, HRA %, PF %, Gratuity %, etc. Change these to re-run the model under a different policy. |
| 04 | `Salary_Engine` | The core calculation: breaks CTC into its components and reconstructs it to prove nothing was lost. |
| 05 | `Validation_Audit` | Automated PASS/FAIL checks (valid ID, positive CTC, CTC reconciles, no negative components). |
| 06 | `CTC_Statement` | A formatted, presentable annual compensation statement for the selected employee. |
| 07 | `Documentation` | In-workbook notes on scope and design intent. |
| 07 | `Monthly_Payslip` | A formatted monthly payslip: earnings, deductions, net pay. |
| 08 | `Compensation_Planning` | Proposed increments for every employee, with revision status and reason. |
| 09 | `Compensation_Simulator` | "What if we gave X% increment?" scenario testing against a total workforce budget. |
| 10 | `Compensation_Dept Budget` | Rolls up current vs. proposed cost, and budget variance, by department. |
| 11 | `Salary_Benchmark` | Compares each employee's CTC to a salary band (min/mid/max) for their role and flags exceptions. |
| 12 | `Compensation_Equity` | Pay-equity view across the workforce. |

## 🔬 How the Calculation Works

The core logic (`04_Salary_Engine`), driven entirely by the assumptions in `03_Policy_Assumptions`:

```
Total CTC
 └─ Variable Pay      = Total CTC × Variable Pay %
 └─ Fixed CTC         = Total CTC − Variable Pay
     └─ Basic Salary  = Fixed CTC × Basic %
     └─ HRA           = Basic × HRA %
     └─ Employer PF   = Basic × Employer PF %      (employer cost, not deducted from pay)
     └─ Gratuity      = Basic × Gratuity %          (employer cost, not deducted from pay)
     └─ Other Allowance = balancing figure so components sum back to Fixed CTC
 
Gross Earnings = Basic + HRA + Other Allowance
Employee PF    = Basic × Employee PF %              (deducted from pay)
Net Pay        = Gross Earnings − Employee PF − Professional Tax − Other Deductions
```

**Built-in reconciliation:** every calculation re-adds the components and checks the total against the original CTC (`Reconciliation Difference` should always be 0). The `05_Validation_Audit` tab surfaces this as a plain PASS/FAIL, so an error in the formulas is caught immediately rather than silently shipped.

## 🔒 Data Privacy

**This is the most important section if you plan to make the repository public.**

- The sample data includes what appear to be real first and last names and exact CTC figures.
- Compensation data is sensitive personal information. Publishing it, even accidentally, can cause real harm and may violate your organisation's data policy or applicable privacy law.
- **Before uploading:**
  1. Open `01_Employee_Master` and replace every name with a placeholder (`Employee 01`, `Employee 02`, ...).
  2. Replace exact CTC figures with clearly illustrative round numbers (e.g. multiples of 50,000).
  3. Check every other tab (`02`, `06`, `07`, `08`, `09`, `10`, `11`, `12`) for the same names or figures carried over, and repeat the substitution there.
  4. Confirm no real department, designation or salary-band figures came from an actual employer's real pay structure, if this is meant purely as a portfolio demo.
- If this workbook was built using your own organisation's real data for a work project, **do not upload it to a public repository at all.** Keep it in a private repo, or build a separate version with synthetic data specifically for your portfolio.

## ⚙️ Scope & Assumptions

Documented on the in-workbook `07_Documentation` tab, and restated here:

- This is an **illustrative, configurable compensation model**, not a full statutory payroll compliance engine. It does not calculate income tax, statutory bonus, or region-specific labour-law deductions.
- PF, gratuity and HRA percentages are simplified, editable assumptions (see `03_Policy_Assumptions`), not a substitute for consulting current statutory rates.
- Variable pay is included in annual CTC but is **not** automatically paid out monthly, reflecting how variable/bonus pay typically works in practice (paid periodically, not every month).
- Employer contributions (PF, gratuity) are shown as employer cost and are **excluded** from the employee's net take-home pay.

## 🚀 How to Use

1. Download the `.xlsm` file and open it in Microsoft Excel. **Enable macros** when prompted, since this file uses VBA.
2. Go to `01_Employee_Master` and review or edit employee records.
3. Go to `02_Employee_Calculator`, select an Employee ID from the dropdown.
4. Every other tab (`04` to `07`) recalculates automatically for that employee.
5. To model policy changes (e.g. a different HRA %), edit `03_Policy_Assumptions` and every calculation updates.
6. For scenario planning, use `09_Compensation_Simulator`: set an increment % and total budget, and see whether the workforce-wide cost stays within budget.

> This workbook uses Excel formulas and VBA macros; it will not fully recalculate in Google Sheets or other spreadsheet software without adjustment.

## 🛠️ Tech Stack

- Microsoft Excel (formulas: `SUMIFS`, `INDEX`/`MATCH`, `IFERROR`)
- VBA (macro-enabled workbook, `.xlsm`)

## 👤 Author

- Sania Sheikh

## 📄 License

Released under the [MIT License](LICENSE).
