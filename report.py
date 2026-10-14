"""
Builds job_report.xlsx from applied_jobs.csv.

Runs automatically at the end of every job_bot.py run. You can also run it
on its own any time:   python report.py

If email_report is set up in config.yaml, the report is also emailed to you.

Your own edits in the "Status" and "Notes" columns of the Applied Jobs sheet
are kept when the report is rebuilt (matched by job link).
"""

import csv
import os
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

HERE = os.path.dirname(os.path.abspath(__file__))
CSV_FILE = os.path.join(HERE, "applied_jobs.csv")
XLSX_FILE = os.path.join(HERE, "job_report.xlsx")

SITES = ["Naukri", "LinkedIn", "Foundit", "Shine", "Cutshort", "Indeed"]
STATUSES = ["Applied", "Viewed", "Interview", "Offer", "Rejected", "No reply"]

FONT = "Arial"
HEAD_FILL = PatternFill("solid", start_color="1F3A68")
HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name=FONT, size=10)
LINK_FONT = Font(name=FONT, size=10, color="0563C1", underline="single")
INPUT_FILL = PatternFill("solid", start_color="FFF7D6")    # cells you edit
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F3A68")
NOTE_FONT = Font(name=FONT, italic=True, size=9, color="666666")
THIN = Side(style="thin", color="D0D5DD")
BORDER = Border(bottom=THIN)


def _read_rows():
    """One row per job from the local database (jobs.db)."""
    import jobdb
    rows = jobdb.rows()
    for r in rows:
        try:
            r["_dt"] = datetime.strptime(r["date"], "%Y-%m-%d %H:%M")
        except Exception:
            r["_dt"] = None
        try:
            r["_pct"] = float(r.get("match_percent") or 0) / 100
        except ValueError:
            r["_pct"] = 0
    return rows


def _previous_edits():
    """Keep Status/Notes the user typed into the last report."""
    edits = {}
    if not os.path.exists(XLSX_FILE):
        return edits
    try:
        wb = load_workbook(XLSX_FILE)
        ws = wb["Applied Jobs"]
        hdr = [c.value for c in ws[3]]
        i_link, i_status, i_notes = hdr.index("Job Link"), hdr.index("Status"), hdr.index("Notes")
        for row in ws.iter_rows(min_row=4):
            cell = row[i_link]
            url = cell.hyperlink.target if cell.hyperlink else cell.value
            if url:
                edits[url] = (row[i_status].value, row[i_notes].value)
    except Exception:
        pass
    # Status/notes set in the dashboard win over the old spreadsheet
    try:
        import json
        with open(os.path.join(HERE, "tracking.json"), encoding="utf-8") as f:
            for url, t in json.load(f).items():
                old = edits.get(url, (None, None))
                edits[url] = (t.get("status") or old[0], t.get("notes") if "notes" in t else old[1])
    except (FileNotFoundError, ValueError):
        pass
    return edits


def _header(ws, row, headers, widths):
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 22
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def _title(ws, text, note):
    ws["A1"] = text
    ws["A1"].font = TITLE_FONT
    ws["A2"] = note
    ws["A2"].font = NOTE_FONT


def _link(cell, url):
    cell.value = "Open job"
    cell.hyperlink = url
    cell.font = LINK_FONT


def _job_sheet(wb, name, rows, note, reason=False):
    ws = wb.create_sheet(name)
    _title(ws, name, note)
    headers = ["S.No", "Date", "Site", "Job Title", "Company", "Location", "Match %",
               "What happened" if reason else "Result", "Job Link"]
    _header(ws, 3, headers, [6, 17, 11, 40, 28, 16, 9, 70 if reason else 38, 11])
    for n, r in enumerate(rows, start=1):
        rr = 3 + n
        vals = [n, r["_dt"] or r["date"], r["site"], r["title"], r["company"], r["location"], r["_pct"],
                (r.get("reason") or r["status"]) if reason else r["status"]]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=rr, column=col, value=v)
            c.font, c.border = BODY_FONT, BORDER
        ws.cell(row=rr, column=2).number_format = "dd-mmm-yyyy hh:mm"
        ws.cell(row=rr, column=7).number_format = "0%"
        _link(ws.cell(row=rr, column=9), r["url"])
        ws.cell(row=rr, column=9).border = BORDER
    if rows:
        ws.auto_filter.ref = f"A3:I{3 + len(rows)}"
    return ws


def build_report():
    rows = _read_rows()
    edits = _previous_edits()

    applied = sorted([r for r in rows if r["status"] == "applied"], key=lambda r: r["_dt"] or datetime.min, reverse=True)
    manual = [r for r in rows if r.get("reason")]
    manual = sorted(manual, key=lambda r: r["_pct"], reverse=True)

    wb = Workbook()
    wb.remove(wb.active)

    # ---------------- Applied Jobs (the main sheet) ----------------
    ws = wb.create_sheet("Applied Jobs")
    _title(ws, "Jobs applied successfully",
           "Yellow columns are yours to update; they are kept when the report is rebuilt. "
           "Follow-up date = 7 days after applying; rows turn amber when it's due.")
    headers = ["S.No", "Applied On", "Site", "Job Title", "Company", "Location", "Match %",
               "Job Link", "Follow Up On", "Status", "Notes"]
    _header(ws, 3, headers, [6, 17, 11, 40, 28, 16, 9, 11, 13, 13, 34])
    status_dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    ws.add_data_validation(status_dv)

    for n, r in enumerate(applied, start=1):
        rr = 3 + n
        prev_status, prev_notes = edits.get(r["url"], (None, None))
        vals = [n, r["_dt"] or r["date"], r["site"], r["title"], r["company"], r["location"], r["_pct"]]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=rr, column=col, value=v)
            c.font, c.border = BODY_FONT, BORDER
        ws.cell(row=rr, column=2).number_format = "dd-mmm-yyyy hh:mm"
        ws.cell(row=rr, column=7).number_format = "0%"
        _link(ws.cell(row=rr, column=8), r["url"])
        fu = ws.cell(row=rr, column=9, value=f"=IF(ISNUMBER(B{rr}),INT(B{rr})+7,\"\")")
        fu.number_format, fu.font, fu.border = "dd-mmm-yyyy", BODY_FONT, BORDER
        st = ws.cell(row=rr, column=10, value=prev_status or "Applied")
        nt = ws.cell(row=rr, column=11, value=prev_notes)
        for c in (st, nt):
            c.font, c.fill, c.border = BODY_FONT, INPUT_FILL, BORDER
        status_dv.add(st)

    last = 3 + len(applied)
    if applied:
        ws.auto_filter.ref = f"A3:K{last}"
        ws.conditional_formatting.add(
            f"A4:K{last}",
            FormulaRule(formula=[f'AND(ISNUMBER($I4),$I4<=TODAY(),OR($J4="Applied",$J4="Viewed"))'],
                        fill=PatternFill("solid", start_color="FCE4C4")))
    else:
        ws["A4"] = "No successful applications yet. Run job_bot.py with dry_run: false."
        ws["A4"].font = NOTE_FONT

    # ---------------- Summary ----------------
    sm = wb.create_sheet("Summary", 0)
    _title(sm, "Job application summary", f"Report built {datetime.now():%d-%b-%Y %H:%M} from the local job database (jobs.db)")
    _header(sm, 3, ["Site", "Applied", "Applied today", "Apply manually", "Interviews", "Jobs scanned"],
            [14, 11, 14, 16, 12, 14])
    sm.freeze_panes = None
    a_rng = f"'Applied Jobs'!$C$4:$C${max(last, 4)}"
    a_dt = f"'Applied Jobs'!$B$4:$B${max(last, 4)}"
    a_st = f"'Applied Jobs'!$J$4:$J${max(last, 4)}"
    m_last = 3 + max(len(manual), 1)
    all_last = 3 + max(len(rows), 1)
    for i, site in enumerate(SITES, start=4):
        sm.cell(row=i, column=1, value=site)
        sm.cell(row=i, column=2, value=f"=COUNTIF({a_rng},A{i})")
        sm.cell(row=i, column=3, value=f"=COUNTIFS({a_rng},A{i},{a_dt},\">=\"&TODAY(),{a_dt},\"<\"&TODAY()+1)")
        sm.cell(row=i, column=4, value=f"=COUNTIF('Apply Manually'!$C$4:$C${m_last},A{i})")
        sm.cell(row=i, column=5, value=f"=COUNTIFS({a_rng},A{i},{a_st},\"Interview\")")
        sm.cell(row=i, column=6, value=f"=COUNTIF('All Jobs Scanned'!$C$4:$C${all_last},A{i})")
        for col in range(1, 7):
            c = sm.cell(row=i, column=col)
            c.font, c.border = BODY_FONT, BORDER
    tr = 4 + len(SITES)
    sm.cell(row=tr, column=1, value="Total").font = Font(name=FONT, bold=True, size=10)
    for col in range(2, 7):
        L = get_column_letter(col)
        c = sm.cell(row=tr, column=col, value=f"=SUM({L}4:{L}{tr - 1})")
        c.font = Font(name=FONT, bold=True, size=10)
        c.border = Border(top=Side(style="thin", color="1F3A68"))

    # ---------------- Manual + All ----------------
    _job_sheet(wb, "Apply Manually", manual,
               "Good matches the bot couldn't apply to by itself. Open each link and apply manually. Best matches first.",
               reason=True)
    _job_sheet(wb, "All Jobs Scanned", sorted(rows, key=lambda r: r["_dt"] or datetime.min, reverse=True),
               "Every job the bot looked at, including skipped ones and the reason.")

    wb.active = 0
    wb.save(XLSX_FILE)
    print(f"Excel report: {XLSX_FILE}  ({len(applied)} applied, {len(manual)} to apply manually)")
    return XLSX_FILE


# ------------------------------------------------------------------ email
def _summary(rows, since=None):
    """Counts per site for the email body (all time, and this run if `since` is given)."""
    out = []
    for site in SITES:
        mine = [r for r in rows if r["site"] == site]
        applied = [r for r in mine if r["status"] == "applied"]
        new = [r for r in applied if since and r["_dt"] and r["_dt"] >= since]
        manual = [r for r in mine if r.get("reason")]
        out.append((site, len(new), len(applied), len(manual)))
    return out


def send_report(since=None):
    """Email job_report.xlsx using the Gmail settings in config.yaml."""
    import smtplib
    import ssl
    from email.message import EmailMessage
    import yaml
    import secret_store

    full = secret_store.reveal(yaml.safe_load(open(os.path.join(HERE, "config.yaml"), encoding="utf-8")) or {})
    cfg = full.get("email_report") or {}
    name = ((full.get("profile") or {}).get("name") or "").split(" ")[0] or "there"
    if not cfg.get("enabled"):
        return False
    pw = str(cfg.get("app_password") or "").replace(" ", "")
    to, frm = cfg.get("send_to"), cfg.get("send_from") or cfg.get("send_to")
    if not (pw and to and frm):
        print("Email not sent: add your Gmail app password under email_report in config.yaml.")
        return False
    if not os.path.exists(XLSX_FILE):
        build_report()

    rows = _read_rows()
    stats = _summary(rows, since)
    new_total = sum(s[1] for s in stats)
    if cfg.get("only_if_new_applications") and since and new_total == 0:
        print("Email skipped: no new applications this run.")
        return False

    today = datetime.now()
    new_jobs = [r for r in rows if r["status"] == "applied" and since and r["_dt"] and r["_dt"] >= since]
    manual_now = [r for r in rows if r.get("reason") and (not since or (r["_dt"] and r["_dt"] >= since))]
    manual_html = "".join(
        f"<tr><td>{r['site']}</td><td><a href='{r['url']}'>{r['title']}</a></td><td>{r['company']}</td>"
        f"<td align=center>{r['match_percent']}%</td><td>{r['reason']}</td></tr>" for r in manual_now[:50])

    table = "".join(
        f"<tr><td>{s}</td><td align=center><b>{n}</b></td><td align=center>{a}</td><td align=center>{m}</td></tr>"
        for s, n, a, m in stats)
    jobs_html = "".join(
        f"<tr><td>{r['site']}</td><td><a href='{r['url']}'>{r['title']}</a></td><td>{r['company']}</td>"
        f"<td align=center>{r['match_percent']}%</td></tr>" for r in new_jobs[:50])
    html = f"""<div style="font-family:Arial,sans-serif;font-size:14px;color:#222">
<p>Hi {name},</p>
<p>Your job bot finished a run on {today:%d %b %Y at %I:%M %p}.
<b>{new_total} new application{'s' if new_total != 1 else ''}</b> were sent this run.
The full Excel report is attached.</p>
<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#ddd;font-size:13px">
<tr style="background:#1F3A68;color:#fff"><th>Site</th><th>Applied this run</th><th>Applied (all time)</th><th>Apply manually</th></tr>
{table}</table>
{f'<h3 style="font-size:15px">Applied this run</h3><table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#ddd;font-size:13px"><tr style="background:#1F3A68;color:#fff"><th>Site</th><th>Role</th><th>Company</th><th>Match</th></tr>{jobs_html}</table>' if new_jobs else ''}
{f'<h3 style="font-size:15px;color:#b4610a">Please apply manually ({len(manual_now)})</h3><p style="margin-top:-6px">Good matches the bot tried but couldn&#39;t apply to automatically.</p><table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#ddd;font-size:13px"><tr style="background:#b4610a;color:#fff"><th>Site</th><th>Role</th><th>Company</th><th>Match</th><th>What happened</th></tr>{manual_html}</table>' if manual_now else ''}
<p style="color:#666;font-size:12px">The "Apply Manually" sheet in the attached report lists every job still waiting for you.</p>
</div>"""

    msg = EmailMessage()
    msg["Subject"] = f"Job report {today:%d %b %Y}: {new_total} new applications"
    msg["From"], msg["To"] = frm, to
    msg.set_content(f"{new_total} new applications this run. The Excel report is attached.")
    msg.add_alternative(html, subtype="html")
    with open(XLSX_FILE, "rb") as f:
        msg.add_attachment(f.read(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=f"job_report_{today:%Y-%m-%d}.xlsx")
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=30) as s:
            s.login(frm, pw)
            s.send_message(msg)
        print(f"Report emailed to {to}")
        return True
    except smtplib.SMTPAuthenticationError:
        print("Email not sent: Gmail rejected the login. Use an App Password (not your normal password) "
              "and make sure 2-Step Verification is on.")
    except Exception as e:
        print("Email not sent:", e)
    return False


if __name__ == "__main__":
    import sys
    build_report()
    if "--email" in sys.argv:
        send_report()
