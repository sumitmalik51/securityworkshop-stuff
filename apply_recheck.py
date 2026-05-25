#!/usr/bin/env python3
"""
Apply recheck results: update batch CSVs with new statuses and screenshots,
then update NOT_FOUND_REMARKS in create_status_excel.py, and regenerate Excel.
"""

import csv
import os

CSV_DIR = r"C:\certs"

# All recheck results compiled from Batch1-5 runs
RECHECK_RESULTS = {
    # Batch1
    "odl_user_2226992": {"batch": "Batch1", "new_status": "COMPLETED", "remark": "Recheck v2: COMPLETED - View insights (29 items found)"},
    "odl_user_2227021": {"batch": "Batch1", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Job in progress"},
    "odl_user_2227233": {"batch": "Batch1", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Job in progress"},
    "odl_user_2227843": {"batch": "Batch1", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    # Batch2
    "odl_user_2227039": {"batch": "Batch2", "new_status": "NOT_FOUND", "remark": "Recheck v2: LOGIN_FAILED - Stuck on login page"},
    "odl_user_2227043": {"batch": "Batch2", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227059": {"batch": "Batch2", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227242": {"batch": "Batch2", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227844": {"batch": "Batch2", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    # Batch3
    "odl_user_2227091": {"batch": "Batch3", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227093": {"batch": "Batch3", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227126": {"batch": "Batch3", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227129": {"batch": "Batch3", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227131": {"batch": "Batch3", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Job in progress"},
    # Batch4
    "odl_user_2227168": {"batch": "Batch4", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227178": {"batch": "Batch4", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227848": {"batch": "Batch4", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227855": {"batch": "Batch4", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    # Batch5
    "odl_user_2227195": {"batch": "Batch5", "new_status": "RUNNING", "remark": "Recheck v2: RUNNING - Pick up where you left off"},
    "odl_user_2227196": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227322": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227323": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227324": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227327": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227328": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227332": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227342": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227345": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227414": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227418": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227421": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227495": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227507": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227508": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227510": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
    "odl_user_2227513": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: NO_PERMISSION - Insufficient permissions"},
    "odl_user_2227821": {"batch": "Batch5", "new_status": "NOT_FOUND", "remark": "Recheck v2: CLASSIC_DIALOG - Classic portal dialog blocking"},
}


def update_batch_csv(batch):
    """Update a batch CSV with recheck results."""
    csv_file = os.path.join(CSV_DIR, f"dspm_status_{batch}.csv")
    if not os.path.exists(csv_file):
        print(f"  [WARN] {csv_file} not found")
        return

    # Read existing rows
    rows = []
    with open(csv_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        for row in reader:
            rows.append(row)

    # Update rows that changed status
    updated = 0
    for row in rows:
        username = row.get("username", "")
        user_key = username.split("@")[0] if "@" in username else username
        if user_key in RECHECK_RESULTS:
            info = RECHECK_RESULTS[user_key]
            if info["batch"] == batch and info["new_status"] != "NOT_FOUND":
                old_status = row["status"]
                row["status"] = info["new_status"]
                # Update screenshot path to v2 version
                screenshots_dir = rf"C:\certs\screenshots_{batch}"
                v2_screenshot = os.path.join(screenshots_dir, f"{user_key}_{info['new_status']}_v2.png")
                if os.path.exists(v2_screenshot):
                    row["screenshot"] = v2_screenshot
                updated += 1
                print(f"    {user_key}: {old_status} -> {info['new_status']}")

    # Write back
    with open(csv_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"  {batch}: {updated} rows updated")


def main():
    print("Updating batch CSVs with recheck results...")
    for batch in ["Batch1", "Batch2", "Batch3", "Batch4", "Batch5"]:
        update_batch_csv(batch)

    print("\nDone! Now run: python create_status_excel.py")


if __name__ == "__main__":
    main()
