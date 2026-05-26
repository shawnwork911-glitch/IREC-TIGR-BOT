"""
upload_to_powerautomate.py
==========================
Sends  output/combined_output.xlsx  to a Power Automate HTTP trigger,
which then saves it to SharePoint.

Required GitHub Secret:
  POWER_AUTOMATE_URL  — the HTTP POST URL from your Power Automate flow
                        (When an HTTP request is received → HTTP POST URL)
"""

import os
import sys
import base64
import requests
from pathlib import Path
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
POWER_AUTOMATE_URL = os.environ.get("POWER_AUTOMATE_URL", "")
LOCAL_FILE         = Path("output/combined_output.xlsx")

# ── Validate ──────────────────────────────────────────────────────────────────
def validate():
    if not POWER_AUTOMATE_URL:
        print("\n[ERROR] POWER_AUTOMATE_URL secret is not set.")
        print("  Go to: GitHub repo → Settings → Secrets and variables → Actions")
        print("  Add secret: POWER_AUTOMATE_URL")
        sys.exit(1)

    if not LOCAL_FILE.exists():
        print(f"\n[ERROR] Output file not found: {LOCAL_FILE}")
        print("  Make sure combined_bot_ci.py ran successfully first.")
        sys.exit(1)

# ── Upload ────────────────────────────────────────────────────────────────────
def upload():
    print("\n" + "=" * 60)
    print("  Power Automate Upload → SharePoint")
    print("=" * 60)

    validate()

    # Read and encode the Excel file as base64
    # Power Automate expects file content as a base64 string
    print(f"\n  Reading: {LOCAL_FILE.resolve()}")
    file_bytes   = LOCAL_FILE.read_bytes()
    file_b64     = base64.b64encode(file_bytes).decode("utf-8")
    file_size_kb = len(file_bytes) / 1024
    print(f"  Size   : {file_size_kb:.1f} KB")

    # Name the file so old files are overwritten
    today       = datetime.now().strftime("%Y-%m-%d")
    remote_name = "combined_output.xlsx"
    print(f"  Sending as: {remote_name}")

    # POST to Power Automate
    payload = {
        "filename":    remote_name,
        "filecontent": file_b64,
    }

    print("\n  Sending to Power Automate...")
    try:
        r = requests.post(
            POWER_AUTOMATE_URL,
            json=payload,
            timeout=120,
        )
    except requests.exceptions.Timeout:
        print("  [ERROR] Request timed out after 120 seconds.")
        sys.exit(1)
    except requests.exceptions.RequestException as e:
        print(f"  [ERROR] Request failed: {e}")
        sys.exit(1)

    # Power Automate returns 200 or 202 on success
    if r.status_code in (200, 202):
        print("  ✓ File sent successfully!")
        print("  ✓ Power Automate will save it to SharePoint now.")
        print(f"\n  Expected SharePoint filename: {remote_name}")
    else:
        print(f"  [ERROR] Power Automate returned: {r.status_code}")
        print(f"  Response: {r.text[:500]}")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  Upload complete!")
    print("=" * 60)

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    upload()
