#!/usr/bin/env python3
"""
fetch_asana.py - Pull SaaS Growth tasks from Asana into data.xlsx (raw export).

Requires environment variable ASANA_TOKEN (a personal access token).
Output: data.xlsx with a single 'Raw Data' sheet, ready for build.py.

Pulls the Reviewed and Reported 2026 sections of the SaaS Growth project,
paginating through all results. Timestamps are converted to US Central time.
"""
import os, sys, time, json, urllib.request, urllib.parse
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl

TOKEN = os.environ.get('ASANA_TOKEN')
if not TOKEN:
    print("ERROR: ASANA_TOKEN environment variable not set")
    sys.exit(1)

PROJECT_GID = "1207186220930827"
SECTIONS = {
    "Reported 2026": "1212708838044909",
    "Reviewed": "1207186220930831",
}
OPT_FIELDS = ("name,assignee.name,created_at,due_on,modified_at,completed_at,completed,"
              "parent.name,memberships.project.gid,memberships.section.name,"
              "custom_fields.name,custom_fields.display_value")
CENTRAL = ZoneInfo("America/Chicago")

HEADERS = ["Task ID", "Type", "Section", "Task Name", "Subtask Name", "Assignee",
           "Created At", "Due Date", "Modified At", "Completed At", "Completed",
           "SaaS Growth Type", "EAID", "Date SaaS Growth Recommended",
           "Date SaaS Growth Occured", "Old SaaS Dollars", "New SaaS Dollars",
           "PCSM Submitter", "Type of Increase", "Was the BA on Recommendation Call",
           "Demo Requested", "Onboarding or Post Onboarding", "Thryv ID",
           "Kicker Payout Amount", "Account name"]


def api_get(url):
    req = urllib.request.Request(
        url, headers={"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"})
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode())


def fetch_section(gid):
    tasks = []
    url = (f"https://app.asana.com/api/1.0/tasks?section={gid}&limit=100"
           f"&opt_fields={urllib.parse.quote(OPT_FIELDS, safe=',.')}")
    while url:
        data = api_get(url)
        tasks.extend(data.get("data", []))
        nxt = data.get("next_page")
        url = nxt["uri"] if nxt else None
        time.sleep(0.2)
    return tasks


def to_central(iso):
    if not iso:
        return ""
    dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
    return dt.astimezone(CENTRAL).strftime("%Y-%m-%d %H:%M:%S")


def to_num(v):
    try:
        n = float(v)
        return int(n) if n == int(n) else n
    except (ValueError, TypeError):
        return ""


def row_for(t):
    cf = {c["name"]: (c.get("display_value") or "") for c in (t.get("custom_fields") or [])}
    section = ""
    for m in (t.get("memberships") or []):
        proj = (m.get("project") or {}).get("gid")
        if proj == PROJECT_GID and m.get("section"):
            section = m["section"]["name"]
    is_sub = t.get("parent") is not None
    assignee = (t.get("assignee") or {}).get("name", "") if t.get("assignee") else ""
    return [
        t.get("gid", ""),
        "saasGrowth",
        section,
        (t["parent"]["name"] if is_sub else t.get("name", "")),
        (t.get("name", "") if is_sub else ""),
        assignee,
        to_central(t.get("created_at")),
        t.get("due_on") or "",
        to_central(t.get("modified_at")),
        to_central(t.get("completed_at")),
        bool(t.get("completed")),
        cf.get("SaaS Growth Type", ""),
        cf.get("EAID", ""),
        (cf.get("Date SaaS Growth Recommended", "") or "")[:10],
        (cf.get("Date SaaS Growth Occured", "") or "")[:10],
        to_num(cf.get("Old SaaS Dollars")),
        to_num(cf.get("New SaaS Dollars")),
        cf.get("PCSM Submitter", ""),
        cf.get("Type of Increase", ""),
        cf.get("Was the BA on Recommendation Call", ""),
        cf.get("Demo Requested", ""),
        cf.get("Onboarding or Post Onboarding", ""),
        cf.get("Thryv ID", ""),
        to_num(cf.get("Kicker Payout Amount")),
        cf.get("Account name", ""),
    ]


def main():
    all_tasks = []
    for name, gid in SECTIONS.items():
        ts = fetch_section(gid)
        print(f"  {name}: {len(ts)} tasks")
        all_tasks.extend(ts)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Raw Data"
    ws.append(HEADERS)
    for t in all_tasks:
        ws.append(row_for(t))
    wb.save("data.xlsx")
    print(f"data.xlsx written: {len(all_tasks)} rows")


if __name__ == "__main__":
    main()
