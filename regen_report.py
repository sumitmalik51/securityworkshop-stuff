import csv, os, sys
from datetime import datetime

sheet = sys.argv[1] if len(sys.argv) > 1 else "Batch1"
csv_file = f"C:/certs/dspm_status_{sheet}.csv"
md_file = f"C:/certs/dspm_report_{sheet}.md"
ss_dir = f"C:/certs/screenshots_{sheet}"

rows = []
with open(csv_file, "r", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        rows.append(r)

rows.sort(key=lambda r: r.get("status", ""))

with open(md_file, "w", encoding="utf-8") as f:
    f.write(f"# DSPM Status Report - {sheet}\n\n")
    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
    comp = sum(1 for r in rows if r["status"] == "COMPLETED")
    run = sum(1 for r in rows if r["status"] == "RUNNING")
    nf = sum(1 for r in rows if r["status"] == "NOT_FOUND")
    err = sum(1 for r in rows if r["status"] in ("ERROR", "CHECK_ERROR"))
    f.write("| Status | Count |\n|--------|-------|\n")
    f.write(f"| COMPLETED | {comp} |\n| RUNNING | {run} |\n| NOT_FOUND | {nf} |\n| ERROR | {err} |\n| **TOTAL** | **{len(rows)}** |\n\n---\n\n")

    for r in rows:
        user = r.get("username", "")
        f.write(f"## {user}\n\n")
        f.write(f"- **Status:** {r.get('status', '')}\n")
        f.write(f"- **Tenant ID:** {r.get('tenant_id', '')}\n")
        f.write(f"- **Checked at:** {r.get('checked_at', '')}\n")
        if r.get("job_title"):
            f.write(f"- **Job:** {r['job_title']}\n")
        if r.get("items_found"):
            f.write(f"- **Items:** {r['items_found']}\n")
        if r.get("job_date"):
            f.write(f"- **Date:** {r['job_date']}\n")

        safe = user.split("@")[0]
        status = r.get("status", "")
        img = os.path.join(ss_dir, f"{safe}_{status}.png")
        if os.path.exists(img):
            rel_path = f"screenshots_{sheet}/{safe}_{status}.png"
            f.write(f"\n![{user}]({rel_path})\n")
        else:
            f.write(f"\n*No screenshot available*\n")
        f.write(f"\n---\n\n")

print(f"Report written: {md_file}")
print(f"Entries: {len(rows)}")
