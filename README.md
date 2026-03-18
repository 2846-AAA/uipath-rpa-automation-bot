# UiPath RPA Automation Bot
**By Anuja Dhamdhere** | Python Developer | Pune, Maharashtra

---

## Project Overview

A complete RPA automation pipeline that integrates **Python scripting** with **UiPath Studio** workflows to automate employee data processing across three phases.

---

## Quick Start

```bash
# 1. Install dependency
pip install openpyxl

# 2. Run the full pipeline
python run_all_phases.py

# 3. Open the dashboard
open dashboard.html   # macOS
start dashboard.html  # Windows
```

---

## Project Structure

```
rpa_project/
├── dashboard.html                        ← Beautiful interactive dashboard
├── run_all_phases.py                     ← Master runner (runs all 3 phases)
│
├── python_scripts/
│   ├── phase1_json_preprocessor.py       ← CSV cleaning → JSON output
│   ├── phase2_webform_bot.py             ← Web form-filling simulation
│   └── phase3_excel_automation.py        ← Excel report generation
│
├── sample_data/
│   ├── employees_raw.csv                 ← Raw input (messy data)
│   ├── employees_clean.json              ← Phase 1 output
│   └── webform_results.json              ← Phase 2 output
│
├── excel_output/
│   └── employee_report.xlsx              ← Phase 3 output (3 sheets)
│
├── logs/
│   ├── phase1_preprocessor.log
│   ├── phase2_webform.log
│   └── phase3_excel.log
│
└── uipath_notes/
    └── phase2_webform_guide.md           ← UiPath .xaml mapping guide
```

---

## Phase Details

### Phase 1 — Python JSON Pre-Processor
- Reads `employees_raw.csv` (intentionally messy — whitespace, mixed case)
- Cleans: strips spaces, title-cases names, lowercases emails
- Validates: department names, salary ranges, date formats, email format
- Outputs: structured `employees_clean.json` with metadata
- **Logs** every row result to `logs/phase1_preprocessor.log`

### Phase 2 — Web Form-Filling Bot
- Reads Phase 1 JSON output
- Simulates UiPath `TypeInto` + `Click` activities per record
- Implements **retry logic** (2 retries on failure) — mirrors UiPath Retry Scope
- Captures confirmation codes per submission
- Outputs: `webform_results.json` with success/fail status
- **Logs** every submission to `logs/phase2_webform.log`

### Phase 3 — Excel Automation
- Reads Phase 1 JSON, computes bonuses (12% IT/Finance, 10% HR/Marketing)
- Writes **3-sheet Excel workbook** using `openpyxl`:
  - Sheet 1: Employee Data (color-coded, formatted)
  - Sheet 2: Department Summary (avg/min/max salary)
  - Sheet 3: Run Log (timestamps of all automation steps)
- **Logs** to `logs/phase3_excel.log`

---

## Technologies Used

| Tool | Purpose |
|------|---------|
| Python 3 | Core scripting language |
| openpyxl | Excel file creation/formatting |
| csv / json | Data parsing modules |
| logging | Structured log output |
| UiPath Studio | Real-world RPA automation |
| Automation Anywhere | Cross-platform RPA exploration |
| Git / GitHub | Version control |

---

## Interview Talking Points

1. **"Walk me through your RPA project"** → Start with the 3-phase pipeline, mention JSON integration between Python and UiPath
2. **"How did you handle errors?"** → Retry logic in Phase 2, try/catch in all phases, dedicated log files
3. **"What UiPath activities did you use?"** → TypeInto, Click, Open Browser, GetText, Retry Scope, Log Message
4. **"How does Python integrate with UiPath?"** → Python pre-processes raw data → outputs clean JSON → UiPath reads JSON via Deserialize activity
5. **"What was the business impact?"** → Automated 3+ workflows, handles 10,000+ records, retry logic reduces manual intervention

---

*Built as a portfolio project — all code is original and functional.*
