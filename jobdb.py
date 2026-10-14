"""
Local job database (jobs.db, SQLite, next to the bot).

One row per job (by URL) with its latest result. The bot reads it to
  - never apply to the same job twice (applied jobs are permanent),
  - apply in the next live run to good matches that a test run found,
  - list the jobs it tried but couldn't finish, so you can apply manually.
applied_jobs.csv is still written as a plain run log.
"""

import csv
import json
import os
import re
import sqlite3
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(HERE, "jobs.db")
CSV_FILE = os.path.join(HERE, "applied_jobs.csv")

SCHEMA = """
CREATE TABLE IF NOT EXISTS jobs (
    url           TEXT PRIMARY KEY,
    site          TEXT,
    title         TEXT,
    company       TEXT,
    location      TEXT,
    match_percent INTEGER DEFAULT 0,
    status        TEXT,
    first_seen    TEXT,
    last_seen     TEXT,
    applied_at    TEXT,
    attempts      INTEGER DEFAULT 0
);
CREATE INDEX IF NOT EXISTS jobs_status ON jobs(status);
CREATE TABLE IF NOT EXISTS questions (
    q          TEXT PRIMARY KEY,   -- normalised question (same key as learned_answers.yaml)
    text       TEXT,               -- as the form showed it
    site       TEXT,
    kind       TEXT,               -- text / choice
    options    TEXT,               -- JSON list for choice questions
    first_seen TEXT,
    last_seen  TEXT,
    times      INTEGER DEFAULT 0,
    state      TEXT DEFAULT 'open' -- open / answered / ignored
);
CREATE TABLE IF NOT EXISTS question_jobs (q TEXT, url TEXT, PRIMARY KEY (q, url));
"""
RETRY = "match (retry)"   # a job that stopped on questions you have answered since; applied in the next live run
FMT = "%Y-%m-%d %H:%M"


def now():
    return datetime.now().strftime(FMT)


def category(status):
    s = (status or "").lower()
    if s in ("applied", "already applied") or s.startswith("already applied"):
        return "applied" if s == "applied" else "skipped"
    if s.startswith(("match:", "needs manual", "no apply button")):
        return "manual"
    if s.startswith(("match (dry run)", RETRY)):
        return "dry"
    if s.startswith("error"):
        return "error"
    return "skipped"


def manual_reason(status):
    """Plain-English reason for a job the bot couldn't apply to by itself."""
    s = (status or "").lower()
    if s.startswith("needs manual"):
        why = {"questions": "the form asked questions it couldn't answer",
               "unknown step": "the form had a step it didn't recognise",
               "too many steps": "the form had too many steps"}
        detail = next((v for k, v in why.items() if k in s), "it couldn't finish the form")
        return f"Tried automatically, but {detail}. Please apply manually."
    if "company site" in s:
        return "Can't apply automatically: this job is applied for on the company's own website. Please apply manually."
    if "not easy apply" in s:
        return "Can't apply automatically: LinkedIn Easy Apply isn't available. Please apply manually."
    if "no apply button" in s or "walk-in" in s:
        return "Can't apply automatically: there was no apply button (maybe a walk-in). Please apply manually."
    return "Couldn't apply automatically. Please apply manually."


def connect():
    first = not os.path.exists(DB_FILE)
    db = sqlite3.connect(DB_FILE, timeout=15)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    if first:
        _import_csv(db)
    return db


def _upsert(db, site, title, company, location, pct, status, url, when):
    cur = db.execute("SELECT status FROM jobs WHERE url = ?", (url,)).fetchone()
    tried = status == "applied" or status.startswith("needs manual")
    if cur is None:
        db.execute("INSERT INTO jobs (url, site, title, company, location, match_percent, status, first_seen, last_seen,"
                   " applied_at, attempts) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                   (url, site, title, company, location, int(float(pct or 0)), status, when, when,
                    when if status == "applied" else None, 1 if tried else 0))
    elif cur["status"] == "applied" and status != "applied":
        db.execute("UPDATE jobs SET last_seen = ? WHERE url = ?", (when, url))     # applied is permanent
    else:
        db.execute("UPDATE jobs SET site=?, title=?, company=?, location=?, match_percent=?, status=?, last_seen=?,"
                   " applied_at = CASE WHEN ? = 'applied' THEN ? ELSE applied_at END, attempts = attempts + ? WHERE url=?",
                   (site, title, company, location, int(float(pct or 0)), status, when, status, when, 1 if tried else 0, url))


def _import_csv(db):
    """First start: bring in the history from applied_jobs.csv."""
    if not os.path.exists(CSV_FILE):
        return
    with open(CSV_FILE, encoding="utf-8", errors="replace") as f:
        for r in csv.DictReader(f):
            if r.get("url"):
                try:
                    pct = int(float(r.get("match_percent") or 0))
                except ValueError:
                    pct = 0
                _upsert(db, r.get("site", ""), r.get("title", ""), r.get("company", ""), r.get("location", ""),
                        pct, r.get("status", ""), r["url"], r.get("date") or now())
    db.commit()


def record(site, title, company, location, pct, status, url):
    with connect() as db:
        _upsert(db, site, title, company, location, pct, status, url, now())


def skip_urls(recheck=False, dry=False):
    """Job URLs a run should not open again.
    Applied / already applied: never. Tried-and-failed (apply manually) and filtered-out jobs: not again,
    unless recheck_old_jobs is on. Test-run matches: skipped in test runs, applied in live runs. Errors: retried."""
    out = set()
    with connect() as db:
        for r in db.execute("SELECT url, status FROM jobs"):
            c = category(r["status"])
            if r["status"] == "applied" or r["status"].startswith("already applied"):
                out.add(r["url"])
            elif c == "dry":
                if dry:
                    out.add(r["url"])
            elif c in ("manual", "skipped") and not recheck:
                out.add(r["url"])
    return out


def applied_keys():
    """(title, company) of every applied job, to catch the same job re-posted or listed on another site."""
    norm = lambda s: re.sub(r"\W+", " ", (s or "").lower()).strip()
    with connect() as db:
        return {(norm(r["title"]), norm(r["company"])) for r in db.execute("SELECT title, company FROM jobs WHERE status='applied'")}


def pending_matches(site):
    """Good matches waiting to be applied to (found by test runs, or ready to retry), best first."""
    with connect() as db:
        return [dict(r) for r in db.execute("SELECT * FROM jobs WHERE site = ? AND status IN ('match (dry run)', ?)"
                                            " ORDER BY match_percent DESC, last_seen DESC", (site, RETRY))]


def applied_today(site):
    with connect() as db:
        return db.execute("SELECT COUNT(*) FROM jobs WHERE site = ? AND status = 'applied' AND applied_at LIKE ?",
                          (site, datetime.now().strftime("%Y-%m-%d") + "%")).fetchone()[0]


# ------------------------------------------------------------------ questions the bot got stuck on
def norm_q(q):
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9#+.? ]", " ", (q or "").lower())).strip()[:200]


def record_questions(site, url, questions):
    """questions: [(text, kind, options)] that stopped an application to `url`."""
    with connect() as db:
        for text, kind, options in questions:
            q = norm_q(text)
            if not q:
                continue
            db.execute("INSERT INTO questions (q, text, site, kind, options, first_seen, last_seen, times, state)"
                       " VALUES (?,?,?,?,?,?,?,1,'open') ON CONFLICT(q) DO UPDATE SET text=excluded.text,"
                       " site=excluded.site, kind=excluded.kind, options=excluded.options, last_seen=excluded.last_seen,"
                       " times=times+1, state=CASE WHEN state='ignored' THEN 'ignored' ELSE 'open' END",
                       (q, text.strip()[:500], site, kind, json.dumps(options or []), now(), now()))
            db.execute("INSERT OR IGNORE INTO question_jobs (q, url) VALUES (?, ?)", (q, url))


def open_questions():
    """Questions waiting for your answer, the ones blocking the most jobs first."""
    with connect() as db:
        out = []
        for r in db.execute("SELECT * FROM questions WHERE state = 'open'"):
            jobs = [dict(j) for j in db.execute(
                "SELECT j.url, j.title, j.company, j.site, j.status FROM question_jobs qj JOIN jobs j ON j.url = qj.url"
                " WHERE qj.q = ? ORDER BY j.last_seen DESC", (r["q"],))]
            waiting = [j for j in jobs if j["status"] != "applied"]
            d = dict(r, options=json.loads(r["options"] or "[]"), jobs=waiting[:5], blocked=len(waiting))
            out.append(d)
        return sorted(out, key=lambda d: (-d["blocked"], -d["times"], d["last_seen"] or ""))


def answer_question(q):
    """You answered q: jobs that stopped only on answered questions go back in the queue. Returns how many."""
    with connect() as db:
        db.execute("UPDATE questions SET state = 'answered' WHERE q = ?", (q,))
        urls = [r["url"] for r in db.execute(
            "SELECT qj.url FROM question_jobs qj JOIN jobs j ON j.url = qj.url"
            " WHERE qj.q = ? AND j.status LIKE 'needs manual%'", (q,))]
        n = 0
        for url in urls:
            still = db.execute("SELECT COUNT(*) FROM question_jobs qj JOIN questions x ON x.q = qj.q"
                               " WHERE qj.url = ? AND x.state = 'open'", (url,)).fetchone()[0]
            if not still:
                db.execute("UPDATE jobs SET status = ? WHERE url = ?", (RETRY, url))
                n += 1
        return n


def ignore_question(q):
    with connect() as db:
        db.execute("UPDATE questions SET state = 'ignored' WHERE q = ?", (q,))


def rows():
    """All jobs, newest first, in the same shape as applied_jobs.csv rows (+ reason, attempts, applied_at)."""
    with connect() as db:
        out = []
        for r in db.execute("SELECT * FROM jobs ORDER BY COALESCE(applied_at, last_seen) DESC"):
            d = dict(r)
            d["date"] = d["applied_at"] or d["last_seen"] or ""
            d["reason"] = manual_reason(d["status"]) if category(d["status"]) == "manual" else ""
            out.append(d)
        return out


def counts():
    with connect() as db:
        return {r["status"]: r["n"] for r in db.execute("SELECT status, COUNT(*) n FROM jobs GROUP BY status")}


if __name__ == "__main__":
    from collections import Counter
    c = Counter()
    for st, n in counts().items():
        c[category(st)] += n
    print(f"{DB_FILE}\n  applied: {c['applied']}   apply manually: {c['manual']}   "
          f"test-run matches waiting: {c['dry']}   skipped: {c['skipped']}   errors: {c['error']}")
