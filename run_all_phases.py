"""
run_all_phases.py
==================
Master runner — executes all 3 Python phases in sequence.
Run this single file to demonstrate the full pipeline.

Usage:
    python run_all_phases.py
"""

import subprocess
import sys
import os

BASE = os.path.dirname(os.path.abspath(__file__))

def run(script_name):
    path = os.path.join(BASE, 'python_scripts', script_name)
    print(f"\n{'='*60}")
    print(f"  Running: {script_name}")
    print(f"{'='*60}")
    result = subprocess.run([sys.executable, path], capture_output=False)
    if result.returncode != 0:
        print(f"[ERROR] {script_name} failed with code {result.returncode}")
        sys.exit(result.returncode)
    print(f"[DONE] {script_name} completed.\n")

if __name__ == '__main__':
    # Check openpyxl
    try:
        import openpyxl
    except ImportError:
        print("Installing openpyxl...")
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'openpyxl'], check=True)

    run('phase1_json_preprocessor.py')
    run('phase2_webform_bot.py')
    run('phase3_excel_automation.py')

    print("\n" + "="*60)
    print("  ALL PHASES COMPLETED SUCCESSFULLY!")
    print("  Check: sample_data/  and  excel_output/  folders")
    print("="*60 + "\n")
