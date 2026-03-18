"""
Phase 2: Web Form-Filling Bot (Python Simulation)
===================================================
RPA Bot Project - Anuja Dhamdhere
----------------------------------
Reads the clean JSON produced by Phase 1 and simulates
what UiPath does: iterates each record and "submits" it
to a web form. Results are logged and saved.

NOTE FOR UIPATH STUDIO:
  The real UiPath workflow (.xaml) that mirrors this logic
  is documented in uipath_notes/phase2_webform_guide.md
"""

import json
import logging
import os
import time
import random
from datetime import datetime

# ── Logging ───────────────────────────────────────────────────────────────────
LOG_DIR = os.path.join(os.path.dirname(__file__), '..', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, 'phase2_webform.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
INPUT_JSON  = os.path.join(BASE_DIR, '..', 'sample_data', 'employees_clean.json')
RESULTS_JSON = os.path.join(BASE_DIR, '..', 'sample_data', 'webform_results.json')

# Simulated target URL (replace with real form URL in UiPath)
TARGET_URL = "https://httpbin.org/post"


# ── Simulated Form Submission ─────────────────────────────────────────────────
def simulate_form_submission(record: dict) -> dict:
    """
    Simulates filling and submitting a web form for one employee record.
    In the real UiPath bot, this maps to:
      - OpenBrowser activity
      - TypeInto activities for each field
      - Click activity on Submit
      - GetText activity to read confirmation
    """
    logger.info(f"  → Opening browser for: {record['name']}")
    time.sleep(0.1)  # simulate browser open delay

    # Simulate filling each form field
    fields = {
        "Employee ID":  record['emp_id'],
        "Full Name":    record['name'],
        "Department":   record['department'],
        "Salary":       str(record['salary']),
        "Join Date":    record['join_date'],
        "Email":        record['email']
    }

    for field_name, value in fields.items():
        logger.info(f"     TypeInto [{field_name}] = '{value}'")
        time.sleep(0.05)

    # Simulate click Submit — 95% success rate
    submission_ok = random.random() > 0.05
    confirmation_code = f"CONF-{record['emp_id']}-{datetime.now().strftime('%H%M%S')}"

    result = {
        "emp_id":            record['emp_id'],
        "name":              record['name'],
        "status":            "SUCCESS" if submission_ok else "FAILED",
        "confirmation_code": confirmation_code if submission_ok else None,
        "error_message":     None if submission_ok else "Simulated timeout on form submit",
        "submitted_at":      datetime.now().isoformat(),
        "fields_filled":     list(fields.keys())
    }

    if submission_ok:
        logger.info(f"  ✓ Submitted — confirmation: {confirmation_code}")
    else:
        logger.warning(f"  ✗ Failed — simulated timeout for {record['name']}")

    return result


# ── Retry Logic ───────────────────────────────────────────────────────────────
def submit_with_retry(record: dict, max_retries: int = 2) -> dict:
    """
    Wraps submission with retry — mirrors UiPath Retry Scope activity.
    """
    for attempt in range(1, max_retries + 2):
        try:
            result = simulate_form_submission(record)
            if result['status'] == 'SUCCESS':
                return result
            if attempt <= max_retries:
                logger.info(f"  ↻ Retry {attempt}/{max_retries} for {record['name']}")
        except Exception as e:
            logger.error(f"  Exception on attempt {attempt}: {e}")
            if attempt > max_retries:
                return {
                    "emp_id":       record['emp_id'],
                    "name":         record['name'],
                    "status":       "ERROR",
                    "error_message": str(e),
                    "submitted_at": datetime.now().isoformat()
                }
    return result


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    logger.info("=" * 60)
    logger.info("Phase 2: Web Form-Filling Bot STARTED")
    logger.info(f"Target URL: {TARGET_URL}")
    logger.info("=" * 60)

    # Load Phase 1 output
    if not os.path.exists(INPUT_JSON):
        logger.error(f"Input JSON not found: {INPUT_JSON}")
        logger.error("Run phase1_json_preprocessor.py first.")
        return

    with open(INPUT_JSON, 'r', encoding='utf-8') as f:
        data = json.load(f)

    records = data.get('records', [])
    logger.info(f"Loaded {len(records)} valid records from Phase 1")

    results = []
    success_count = 0
    fail_count = 0

    for i, record in enumerate(records, 1):
        logger.info(f"\n[{i}/{len(records)}] Processing: {record['name']} ({record['emp_id']})")
        result = submit_with_retry(record)
        results.append(result)
        if result['status'] == 'SUCCESS':
            success_count += 1
        else:
            fail_count += 1

    # Save results
    output = {
        "metadata": {
            "generated_at":  datetime.now().isoformat(),
            "target_url":    TARGET_URL,
            "total":         len(results),
            "success_count": success_count,
            "fail_count":    fail_count
        },
        "results": results
    }

    with open(RESULTS_JSON, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=4, ensure_ascii=False)

    logger.info("\n" + "=" * 60)
    logger.info(f"Phase 2 COMPLETED: {success_count} success | {fail_count} failed")
    logger.info(f"Results saved → {RESULTS_JSON}")
    logger.info("=" * 60)


if __name__ == '__main__':
    main()
