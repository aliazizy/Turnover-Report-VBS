import sys, requests, os

api_url = sys.argv[1] if len(sys.argv) > 1 else "https://turnover-report-vbs.azurewebsites.net/report"
json_path = sys.argv[2] if len(sys.argv) > 2 else "turnover.json"
out_path  = sys.argv[3] if len(sys.argv) > 3 else "turnover-report.xlsx"

if not os.path.exists(json_path):
    print(f"JSON not found: {json_path}", file=sys.stderr)
    sys.exit(1)

with open(json_path, "rb") as f:
    files = {"file": ("turnover.json", f, "application/json")}
    r = requests.post(api_url, files=files)

if not r.ok:
    print(f"HTTP {r.status_code}: {r.text}", file=sys.stderr)
    sys.exit(2)

# Prefer filename from response if provided
cd = r.headers.get("Content-Disposition", "")
fname = out_path
if "filename=" in cd:
    fname = cd.split("filename=",1)[1].strip().strip('"')

with open(fname, "wb") as out:
    out.write(r.content)

print(f"✅ Saved: {os.path.abspath(fname)}")
