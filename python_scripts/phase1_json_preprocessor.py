"""
Phase 1: Python JSON Pre-Processing Script
==========================================
RPA Bot Project - Anuja Dhamdhere
----------------------------------
Reads raw employee CSV data, cleans/validates it,
and outputs a structured JSON file for UiPath to consume.
"""

import csv
import json
import logging
import os
from datetime import datetime

# ── Logging Setup ─────────────────────────────────────────────────────────────
LOG_DIR = os.path.join(os.path.dirname(__file__), '..', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, 'phase1_preprocessor.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_CSV  = os.path.join(BASE_DIR, '..', 'sample_data', 'employees_raw.csv')
OUTPUT_JSON = os.path.join(BASE_DIR, '..', 'sample_data', 'employees_clean.json')

VALID_DEPARTMENTS = {"IT", "HR", "Finance", "Marketing", "Operations",
                     "It", "Hr", "Information Technology"}


# ── Helpers ───────────────────────────────────────────────────────────────────
def clean_string(value: str) -> str:
    """Strip whitespace and title-case a string."""
    return value.strip().title() if value else ""


def clean_email(value: str) -> str:
    """Strip and lowercase an email address."""
    return value.strip().lower()


def clean_salary(value: str) -> float | None:
    """Convert salary string to float, return None on error."""
    try:
        return float(value.strip().replace(',', ''))
    except ValueError:
        return None


def validate_date(value: str) -> str | None:
    """Return ISO date string if valid, else None."""
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(value.strip(), fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return None


def validate_email(email: str) -> bool:
    """Basic email format check."""
    return '@' in email and '.' in email.split('@')[-1]


# ── Core Processing ───────────────────────────────────────────────────────────
def process_csv(input_path: str) -> tuple[list[dict], list[dict]]:
    """
    Read and clean the CSV.
    Returns (valid_records, error_records).
    """
    valid_records = []
    error_records = []

    logger.info(f"Reading CSV: {input_path}")

    with open(input_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        # Normalise header names (strip whitespace)
        reader.fieldnames = [h.strip() for h in reader.fieldnames]

        for row_num, row in enumerate(reader, start=2):  # row 1 = header
            errors = []

            emp_id     = row.get('emp_id', '').strip()
            name       = clean_string(row.get('name', ''))
            department = clean_string(row.get('department', ''))
            salary     = clean_salary(row.get('salary', ''))
            join_date  = validate_date(row.get('join_date', ''))
            email      = clean_email(row.get('email', ''))

            # ── Validation ────────────────────────────────────────────────────
            if not emp_id:
                errors.append("Missing emp_id")
            if not name:
                errors.append("Missing name")
            if department not in VALID_DEPARTMENTS:
                errors.append(f"Invalid department: '{department}'")
            if salary is None or salary <= 0:
                errors.append(f"Invalid salary: '{row.get('salary','')}'")
            if join_date is None:
                errors.append(f"Invalid date: '{row.get('join_date','')}'")
            if not validate_email(email):
                errors.append(f"Invalid email: '{email}'")

            record = {
                "emp_id":     emp_id,
                "name":       name,
                "department": department,
                "salary":     salary,
                "join_date":  join_date,
                "email":      email,
                "processed_at": datetime.now().isoformat()
            }

            if errors:
                record["errors"] = errors
                error_records.append(record)
                logger.warning(f"Row {row_num} [{emp_id}] — validation errors: {errors}")
            else:
                valid_records.append(record)
                logger.info(f"Row {row_num} [{emp_id}] — OK: {name}")

    return valid_records, error_records


def save_json(valid: list[dict], errors: list[dict], output_path: str) -> None:
    """Write structured JSON output."""
    output = {
        "metadata": {
            "generated_at":   datetime.now().isoformat(),
            "source_file":    os.path.basename(INPUT_CSV),
            "total_records":  len(valid) + len(errors),
            "valid_count":    len(valid),
            "error_count":    len(errors),
            "script_version": "1.0.0"
        },
        "records": valid,
        "error_records": errors
    }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=4, ensure_ascii=False)

    logger.info(f"JSON saved → {output_path}")
    logger.info(f"Summary: {len(valid)} valid | {len(errors)} errors")


# ── Entry Point ───────────────────────────────────────────────────────────────
def main():
    logger.info("=" * 60)
    logger.info("Phase 1: JSON Pre-Processor STARTED")
    logger.info("=" * 60)

    try:
        valid, errors = process_csv(INPUT_CSV)
        save_json(valid, errors, OUTPUT_JSON)
        logger.info("Phase 1: COMPLETED SUCCESSFULLY")
    except FileNotFoundError as e:
        logger.error(f"Input file not found: {e}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        raise

    logger.info("=" * 60)


if __name__ == '__main__':
    main()
