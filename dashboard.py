"""
Job Bot Dashboard: manage everything from your browser.

    python dashboard.py          (or double-click "Start Dashboard.bat")

Opens http://127.0.0.1:8765 . It only listens on your own computer, so your
passwords and keys never leave it. From the dashboard you can:
  - set up your profile, skills, keywords, answers, AI key and email (no YAML editing)
  - open the bot's Chrome, start a test run or a live run, stop it, watch the log
  - browse every job scanned, track interviews/offers, add notes
  - build / download / email the Excel report
  - schedule a daily run
"""

import csv
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from collections import deque
from datetime import datetime, timedelta

import yaml
import jobdb
import secret_store
import site_login
from flask import Flask, jsonify, request, send_file, send_from_directory

HERE = os.path.dirname(os.path.abspath(__file__))
UI_DIR = os.path.join(HERE, "dashboard_ui")
CONFIG_FILE = os.path.join(HERE, "config.yaml")
EXAMPLE_FILE = os.path.join(HERE, "config.example.yaml")
SITES_FILE = os.path.join(HERE, "sites.yaml")
CSV_FILE = os.path.join(HERE, "applied_jobs.csv")
XLSX_FILE = os.path.join(HERE, "job_report.xlsx")
LEARNED_FILE = os.path.join(HERE, "learned_answers.yaml")
TRACK_FILE = os.path.join(HERE, "tracking.json")
STOP_FILE = os.path.join(HERE, ".stop_bot")
LOG_DIR = os.path.join(HERE, "logs")
CHROME_PROFILE = os.path.join(HERE, "chrome_profile")
RESUME_DIR = os.path.join(HERE, "resume")
RESUME_TEXT = os.path.join(RESUME_DIR, "resume.txt")   # read by the bot's AI helper
RESUME_TYPES = {".pdf", ".doc", ".docx"}

PORT = int(os.environ.get("JOBBOT_PORT", "8765"))
SITE_NAMES = {"naukri": "Naukri", "linkedin": "LinkedIn", "foundit": "Foundit",
              "shine": "Shine", "cutshort": "Cutshort", "indeed": "Indeed"}

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
LOCK = threading.Lock()


# ------------------------------------------------------------------ files
def ensure_config():
    """First run on a new computer: start from the blank template."""
    if not os.path.exists(CONFIG_FILE) and os.path.exists(EXAMPLE_FILE):
        shutil.copy(EXAMPLE_FILE, CONFIG_FILE)


def load_config():
    ensure_config()
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


CONFIG_HEADER = ("# Job Auto-Apply settings. Edit them in the dashboard (python dashboard.py).\n"
                 "# Passwords and API keys are encrypted with your Windows login (enc:...), so they only work\n"
                 "# on this computer. Your credentials stay on YOUR computer. Never share this file.\n\n")
MASK = "••••••••"   # what the browser gets instead of a saved password/key


def save_config(cfg):
    secret_store.protect(cfg)
    if os.path.exists(CONFIG_FILE):
        shutil.copy(CONFIG_FILE, CONFIG_FILE + ".bak")
    tmp = CONFIG_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(CONFIG_HEADER)
        yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False, width=120)
    os.replace(tmp, CONFIG_FILE)


def encrypt_saved_secrets():
    """Older versions kept passwords/keys as plain text: encrypt them (config.yaml and its .bak)."""
    for path in (CONFIG_FILE, CONFIG_FILE + ".bak"):
        try:
            if secret_store.protect_file(path, CONFIG_HEADER):
                print(f"  Encrypted the passwords and keys in {os.path.basename(path)}.")
        except Exception as e:
            print(f"  Couldn't encrypt {os.path.basename(path)}: {e}")


def get_path(cfg, path):
    for k in path.split("."):
        cfg = cfg.get(k) if isinstance(cfg, dict) else None
    return cfg


def set_path(cfg, path, value):
    *parents, last = path.split(".")
    for k in parents:
        cfg = cfg.setdefault(k, {})
    cfg[last] = value


def masked_paths(cfg, path=""):
    out = []
    for k, v in cfg.items():
        if isinstance(v, dict):
            out += masked_paths(v, f"{path}{k}.")
        elif v == MASK:
            out.append(f"{path}{k}")
    return out


def load_tracking():
    try:
        with open(TRACK_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError):
        return {}


def save_tracking(data):
    tmp = TRACK_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, ensure_ascii=False)
    os.replace(tmp, TRACK_FILE)


category = jobdb.category


def read_jobs():
    """One row per job from the local database (jobs.db), oldest first."""
    rows = jobdb.rows()[::-1]
    track = load_tracking()
    for i, r in enumerate(rows):
        r["id"] = i
        r["category"] = category(r.get("status"))
        try:
            r["match"] = int(float(r.get("match_percent") or 0))
        except ValueError:
            r["match"] = 0
        t = track.get(r.get("url", ""), {})
        r["track"] = t.get("status") or ("Applied" if r["category"] == "applied" else "")
        r["notes"] = t.get("notes", "")
    return rows


# ------------------------------------------------------------------ bot process
# Lines job_bot.py prints, turned into live progress for the dashboard
RX_HEAD = re.compile(r"^=+\s+(.+?)\s+\((.*)\)\s+=+$")
RX_JOB = re.compile(r"^\s*\[(.+?)\]\s+(\d+)%\s+(.*?) \| (.*)$")
RX_SEARCH = re.compile(r"^(.+?) search: (.+?) / (.+?) / page (\d+)")
PRETTY_SITE = {v.upper(): v for v in SITE_NAMES.values()}


def blank_progress():
    return {"site": None, "keyword": None, "search": None, "waiting": None, "waiting_text": "",
            "counts": {k: 0 for k in ("checked", "applied", "manual", "skipped", "error", "dry")},
            "sites": {}, "events": [],
            "parallel": False, "running_sites": [], "queue": []}   # queue: bots waiting for you, oldest first


class Bot:
    """Runs job_bot.py: one process for all sites in turn, or (parallel) one process per site at the same time."""

    def __init__(self):
        self.procs = {}          # site key -> process ("" when one process runs every site)
        self.wrapping = False    # parallel run: all sites done, building the report
        self.lines = deque(maxlen=5000)
        self.offset = 0          # number of lines dropped from the front of the deque
        self.started = None
        self.finished = None
        self.exit_code = None
        self.mode = None
        self.trigger = None
        self.parallel = False
        self.logfile = None
        self.progress = blank_progress()

    @property
    def running(self):
        return self.wrapping or any(p.poll() is None for p in self.procs.values())

    @staticmethod
    def _name(key):
        if not key:
            return None
        if key in SITE_NAMES:
            return SITE_NAMES[key]
        try:
            return (yaml.safe_load(open(SITES_FILE, encoding="utf-8")) or {}).get(key, {}).get("name") or key.title()
        except Exception:
            return key.title()

    def _sync_waiting(self):
        p = self.progress
        head = p["queue"][0] if p["queue"] else None
        p["waiting"], p["waiting_text"] = (head["kind"], head["text"]) if head else (None, "")

    def _track(self, line, site=None):
        p = self.progress
        if (p["queue"] and line.strip() and "Finish logging in" not in line and "YOUR TURN" not in line
                and not line.startswith("      Answer it")):
            # more output from a bot that was waiting means it carried on
            p["queue"] = [w for w in p["queue"] if site and w["site"] != site]
            self._sync_waiting()
        m = RX_HEAD.match(line)
        if m:
            p["site"], p["keyword"], p["search"] = site or PRETTY_SITE.get(m.group(1), m.group(1).title()), m.group(2), None
            p["sites"].setdefault(p["site"], {k: 0 for k in p["counts"]})
            return
        m = RX_SEARCH.match(line)
        if m:
            p["search"] = {"keyword": m.group(2), "location": m.group(3), "page": int(m.group(4))}
            return
        m = RX_JOB.match(line)
        if m:
            cat = category(m.group(1))
            site = site or p["site"] or "?"
            per = p["sites"].setdefault(site, {k: 0 for k in p["counts"]})
            for bucket in (p["counts"], per):
                bucket["checked"] += 1
                bucket[cat] = bucket.get(cat, 0) + 1
            p["events"].append({"time": datetime.now().strftime("%H:%M"), "site": site, "status": m.group(1),
                                "category": cat, "match": int(m.group(2)), "title": m.group(3), "company": m.group(4)})
            del p["events"][:-200]
            return
        kind = "login" if "Finish logging in" in line else "answer" if "YOUR TURN" in line else None
        if kind:
            text = line.strip(" >.") if kind == "login" else line.split("YOUR TURN:", 1)[-1].strip()
            p["queue"] = [w for w in p["queue"] if w["site"] != site] + [
                {"site": site, "kind": kind, "text": f"{site}: {text}" if site and kind == "answer" else text}]
            self._sync_waiting()

    def _add(self, line, site=None, own=False):
        """own: a message from the dashboard itself, shown in the log but not read as bot output."""
        shown = f"[{site}] {line}" if site and line.strip() else line
        with LOCK:
            if len(self.lines) == self.lines.maxlen:
                self.offset += 1
            self.lines.append(shown)
            try:
                if not own:
                    self._track(line, site)
            except Exception:
                pass
            if self.logfile:
                try:
                    self.logfile.write(shown + "\n"); self.logfile.flush()
                except Exception:
                    pass

    def start(self, dry_run=None, trigger="manual", task=None, parallel=False, sites=None):
        """sites: only these site keys (from the ones switched on); parallel: one bot per site, all at once."""
        if self.running:
            return False, "The bot is already running."
        cfg = load_config()
        keys = [""]
        chosen = []
        if not task:
            chosen = [k for k in (cfg.get("run_order") or SITE_NAMES) if (cfg.get(k) or {}).get("enabled")]
            if sites:
                chosen = [k for k in chosen if k in sites]
            if not chosen:
                return False, "Pick at least one job site." if sites else "No job sites are switched on."
            if parallel:
                keys = chosen
        if os.path.exists(STOP_FILE):
            os.remove(STOP_FILE)
        dry = bool(cfg.get("dry_run", True)) if dry_run is None else bool(dry_run)
        if task == "naukri_resume":
            dry = False
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1",
                   JOBBOT_DRY_RUN="1" if dry else "0", JOBBOT_FROM_DASHBOARD="1")
        env.pop("JOBBOT_SITES", None); env.pop("JOBBOT_NO_REPORT", None); env.pop("JOBBOT_START_DELAY", None)
        if sites and not parallel and not task:   # one bot, only the sites you picked
            env["JOBBOT_SITES"] = ",".join(chosen)
        os.makedirs(LOG_DIR, exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        self.logfile = open(os.path.join(LOG_DIR, f"run_{stamp}{'_naukri-resume' if task else '_test' if dry else ''}.log"), "w", encoding="utf-8")
        self.parallel = keys != [""]
        with LOCK:
            self.lines.clear(); self.offset = 0
            self.progress = blank_progress()
            self.progress["parallel"] = self.parallel
            self.progress["running_sites"] = [self._name(k) for k in keys if k]
            self.progress["run_sites"] = [self._name(k) for k in chosen]   # the sites in this run, for the Activity page
        self.started, self.finished, self.exit_code = datetime.now(), None, None
        self.mode, self.trigger = ("resume" if task else "test" if dry else "live"), trigger
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        args = [sys.executable, "-u", os.path.join(HERE, "job_bot.py")] + (["--naukri-resume"] if task == "naukri_resume" else [])
        self._add(f"=== {'Updating your Naukri resume' if task else 'TEST run (no applications sent)' if dry else 'LIVE run'} started "
                  f"{self.started:%d %b %Y %H:%M} ({trigger}{', all sites at the same time' if self.parallel else ''}) ===", own=True)
        with LOCK:   # held until every site's bot has started, so none counts as "the last one done" too early
            self.procs = {}
            for i, key in enumerate(keys):
                # one bot per site, started 3 s apart so they don't all hit Chrome at once
                e = dict(env, JOBBOT_SITES=key, JOBBOT_NO_REPORT="1", JOBBOT_START_DELAY=str(3 * i)) if self.parallel else env
                proc = subprocess.Popen(args, cwd=HERE, env=e,
                                        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                        text=True, encoding="utf-8", errors="replace", bufsize=1, creationflags=flags)
                self.procs[key] = proc
                threading.Thread(target=self._pump, args=(key, proc), daemon=True).start()
        return True, "Started."

    def _pump(self, key, proc):
        name = self._name(key)
        for line in proc.stdout:
            self._add(line.rstrip("\r\n"), name)
        proc.wait()
        with LOCK:
            p = self.progress
            if name in p["running_sites"]:
                p["running_sites"].remove(name)
            p["queue"] = [w for w in p["queue"] if w["site"] != name]
            self._sync_waiting()
            last = not any(q.poll() is None for q in self.procs.values()) and not self.wrapping
            if last:
                self.wrapping = True
        if name:
            self._add(f">> {name}: done (exit code {proc.returncode})", own=True)
        if not last:
            return
        try:
            if self.parallel and self.mode != "resume":
                self._build_report()
        finally:
            codes = [q.returncode for q in self.procs.values()]
            self.exit_code, self.finished = next((c for c in codes if c), 0), datetime.now()
            self._add(f"=== Finished {self.finished:%H:%M} (exit code {self.exit_code}) ===", own=True)
            if os.path.exists(STOP_FILE):
                os.remove(STOP_FILE)
            with LOCK:
                if self.logfile:
                    self.logfile.close(); self.logfile = None
            self.wrapping = False

    def _build_report(self):
        """Parallel runs: each site's bot skips the report, so build and email it once here."""
        since = self.started.replace(second=0, microsecond=0)
        code = ("from datetime import datetime\nfrom report import build_report, send_report\n"
                "try:\n    build_report()\n    send_report(since=datetime.fromisoformat(%r))\n"
                "except Exception as e:\n    print(\"Couldn't build the Excel report:\", e, \"(run: python report.py)\")\n"
                % since.isoformat())
        try:
            flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            r = subprocess.run([sys.executable, "-u", "-c", code], cwd=HERE, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=300, creationflags=flags,
                               env=dict(os.environ, PYTHONIOENCODING="utf-8"))
            for line in (r.stdout + r.stderr).splitlines():
                self._add(line, own=True)
        except Exception as e:
            self._add(f"Couldn't build the Excel report: {e} (run: python report.py)", own=True)

    def stop(self):
        """Ask nicely: the bot finishes the current job, then builds the report and exits."""
        if not self.running:
            return False
        open(STOP_FILE, "w").close()
        self._add(">> Stop requested. The bot will stop after the current job and save the report...", own=True)
        return True

    def kill(self):
        if not self.running:
            return False
        for proc in self.procs.values():
            if proc.poll() is None:
                proc.kill()
        self._add(">> Bot force-stopped.", own=True)
        return True

    def send_enter(self, text=""):
        """Continue: answers the bot that has been waiting longest (in a parallel run, that site's bot)."""
        if not self.running:
            return False
        with LOCK:
            queue = self.progress["queue"]
            head = queue[0] if queue else None
            if head:
                queue.pop(0); self._sync_waiting()
        key = next((k for k in self.procs if self._name(k) == head["site"]), None) if head and head["site"] else None
        targets = [self.procs[key]] if key is not None else list(self.procs.values())
        sent = False
        for proc in targets:
            if proc.poll() is None:
                try:
                    proc.stdin.write(text + "\n"); proc.stdin.flush(); sent = True
                except Exception:
                    pass
        if sent:
            self._add(f">> (sent Enter from dashboard{' to ' + head['site'] if head and head['site'] else ''})", own=True)
        return sent

    def state(self):
        return {"running": self.running, "mode": self.mode, "trigger": self.trigger, "parallel": self.parallel,
                "started": self.started.isoformat() if self.started else None,
                "finished": self.finished.isoformat() if self.finished else None,
                "exit_code": self.exit_code, "stopping": os.path.exists(STOP_FILE) and self.running,
                "progress": self.progress if self.started else None}


BOT = Bot()


# ------------------------------------------------------------------ chrome
def find_chrome():
    cands = []
    if os.name == "nt":
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"), os.environ.get("LocalAppData")):
            if base:
                cands.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
        for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles")):
            if base:
                cands.append(os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"))
    elif sys.platform == "darwin":
        cands += ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                  "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"]
    else:
        cands += [shutil.which(n) or "" for n in ("google-chrome", "google-chrome-stable", "chromium", "microsoft-edge")]
    return next((c for c in cands if c and os.path.exists(c)), None)


def chrome_connected():
    try:
        with urllib.request.urlopen("http://127.0.0.1:9222/json/version", timeout=1.5) as r:
            return r.status == 200
    except Exception:
        return False


def login_urls(keys=None):
    """Login pages to open: the given site keys, or Google plus every switched-on site."""
    cfg = load_config()
    try:
        extra = yaml.safe_load(open(SITES_FILE, encoding="utf-8")) or {}
    except Exception:
        extra = {}
    return [s["url"] for s in site_login.sites(cfg, extra) if (s["key"] in keys if keys else s["enabled"])]


def open_chrome(with_logins=True, sites=None):
    """Open the bot's Chrome. with_logins: also a login tab per site (sites = only these keys),
    added as new tabs when Chrome is already open, so you can sign in again any time."""
    urls = login_urls(sites) if with_logins else []
    if chrome_connected():
        if not urls:
            return True, "Chrome is already open in bot mode."
        n = site_login.open_tabs(urls)
        if not n:
            return False, "Couldn't open the login tabs. Close ALL Chrome windows and try again."
        return True, f"Opened {n} login tab{'s' if n != 1 else ''} in the bot's Chrome."
    exe = find_chrome()
    if not exe:
        return False, "Google Chrome was not found. Install it from https://www.google.com/chrome"
    args = [exe, "--remote-debugging-port=9222", f"--user-data-dir={CHROME_PROFILE}", "--no-first-run"]
    subprocess.Popen(args + urls, cwd=HERE)
    for _ in range(20):
        time.sleep(0.5)
        if chrome_connected():
            return True, "Chrome opened."
    return False, ("Chrome started but isn't reachable in bot mode. Close ALL Chrome windows "
                   "(check the system tray too) and try again.")


# ------------------------------------------------------------------ scheduler
def scheduler():
    last_run_day = None
    while True:
        time.sleep(20)
        try:
            sch = (load_config().get("dashboard") or {}).get("schedule") or {}
            if not sch.get("enabled"):
                continue
            now = datetime.now()
            days = sch.get("days") or ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
            if now.strftime("%a").lower()[:3] not in [d.lower()[:3] for d in days]:
                continue
            hh, mm = [int(x) for x in str(sch.get("time", "10:00")).split(":")[:2]]
            due = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
            if last_run_day != now.date() and due <= now < due + timedelta(minutes=5) and not BOT.running:
                last_run_day = now.date()
                if load_config().get("browser_mode", "attach") == "attach" and not chrome_connected():
                    open_chrome(with_logins=False)
                BOT.start(trigger="schedule")
        except Exception as e:
            print("Scheduler:", e)


# ------------------------------------------------------------------ API
@app.get("/")
def index():
    return send_from_directory(UI_DIR, "index.html")


@app.get("/api/state")
def api_state():
    cfg = load_config()
    profile = cfg.get("profile") or {}
    dash = cfg.get("dashboard") or {}
    st = jobdb.counts()
    pending = st.get("match (dry run)", 0)
    has_live = bool(st.get("applied"))
    has_test = bool(pending) or any(k.startswith(("skipped", "match:", "needs manual")) for k in st)
    if os.path.isdir(LOG_DIR):
        has_test = has_test or any("_test" in n for n in os.listdir(LOG_DIR))
    ai, er = cfg.get("ai") or {}, cfg.get("email_report") or {}
    sites_on = [k for k in (cfg.get("run_order") or SITE_NAMES) if (cfg.get(k) or {}).get("enabled")]
    return jsonify(
        onboarding={"signed_in": bool(dash.get("signed_in")), "test_run": has_test, "live_run": has_live,
                    "ai": str(ai.get("provider", "none")).lower() not in ("", "none") and bool(ai.get("api_key")),
                    "email": bool(er.get("enabled") and er.get("app_password") and er.get("send_to")),
                    "schedule": bool((dash.get("schedule") or {}).get("enabled")),
                    "resume": bool(resume_path(cfg))},
        run_plan={"sites": [SITE_NAMES.get(k, k) for k in sites_on], "site_keys": sites_on,
                  "keywords": (cfg.get("search") or {}).get("keywords") or [],
                  "locations": (cfg.get("search") or {}).get("locations") or [],
                  "max_per_site": cfg.get("max_applications_per_site", 25),
                  "min_match": (cfg.get("match") or {}).get("min_match_percent", 50),
                  "pending": pending + st.get(jobdb.RETRY, 0),
                  "max_per_day": cfg.get("max_applications_per_day", 0)},
                   questions=len(jobdb.open_questions()),bot=BOT.state(), chrome=chrome_connected(), chrome_found=bool(find_chrome()),
                   configured=bool(profile.get("name")) and bool((cfg.get("search") or {}).get("keywords")),
                   name=profile.get("name", ""), dry_run=bool(cfg.get("dry_run", True)),
                   browser_mode=cfg.get("browser_mode", "attach"),
                   schedule=(cfg.get("dashboard") or {}).get("schedule") or {},
                   report_exists=os.path.exists(XLSX_FILE),
                   report_time=datetime.fromtimestamp(os.path.getmtime(XLSX_FILE)).isoformat() if os.path.exists(XLSX_FILE) else None)


@app.get("/api/stats")
def api_stats():
    rows = read_jobs()
    today = datetime.now().strftime("%Y-%m-%d")
    days = [(datetime.now() - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(13, -1, -1)]
    daily = {d: {"applied": 0, "scanned": 0, "manual": 0} for d in days}
    sites = {}
    pipeline = {}
    followups = []
    for r in rows:
        d = (r.get("date") or "")[:10]
        c = r["category"]
        s = sites.setdefault(r.get("site") or "?", {"site": r.get("site") or "?", "applied": 0, "manual": 0,
                                                     "scanned": 0, "skipped": 0, "error": 0, "dry": 0, "today": 0})
        s["scanned"] += 1
        s[c] = s.get(c, 0) + 1
        if c == "applied" and d == today:
            s["today"] += 1
        if d in daily:
            daily[d]["scanned"] += 1
            if c in ("applied", "manual"):
                daily[d][c] += 1
        if c == "applied":
            pipeline[r["track"]] = pipeline.get(r["track"], 0) + 1
            try:
                due = datetime.strptime(r["date"], "%Y-%m-%d %H:%M") + timedelta(days=7)
                if due <= datetime.now() and r["track"] in ("Applied", "Viewed"):
                    followups.append({**r, "due": due.strftime("%Y-%m-%d")})
            except ValueError:
                pass
    applied = [r for r in rows if r["category"] == "applied"]
    manual = sorted([r for r in rows if r["category"] == "manual" and r["track"] not in ("Done", "Not interested")],
                    key=lambda r: -r["match"])
    totals = {
        "scanned": len(rows),
        "applied": len(applied),
        "applied_today": sum(1 for r in applied if r["date"].startswith(today)),
        "manual": len(manual),
        "interviews": len({r["url"] for r in rows if r["track"] in ("Interview", "Offer")}),
        "offers": len({r["url"] for r in rows if r["track"] == "Offer"}),
        "errors": sum(1 for r in rows if r["category"] == "error"),
        "avg_match": round(sum(r["match"] for r in applied) / len(applied)) if applied else 0,
        "last_activity": rows[-1]["date"] if rows else None,
    }
    return jsonify(totals=totals, daily=[{"date": d, **daily[d]} for d in days],
                   sites=sorted(sites.values(), key=lambda s: -s["scanned"]), pipeline=pipeline,
                   followups=followups[:20], todo=manual[:8])


@app.get("/api/jobs")
def api_jobs():
    return jsonify(read_jobs())


@app.post("/api/jobs/track")
def api_track():
    body = request.get_json(force=True)
    url = body.get("url")
    if not url:
        return jsonify(ok=False, error="missing url"), 400
    with LOCK:
        data = load_tracking()
        entry = data.get(url, {})
        for k in ("status", "notes"):
            if k in body:
                entry[k] = body[k]
        entry["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        data[url] = entry
        save_tracking(data)
    return jsonify(ok=True)


@app.get("/api/config")
def api_get_config():
    cfg = load_config()
    for p in secret_store.secret_paths(cfg):      # saved passwords/keys never go to the browser
        set_path(cfg, p, MASK)
    return jsonify(cfg)


@app.post("/api/config")
def api_save_config():
    cfg = request.get_json(force=True)
    if not isinstance(cfg, dict):
        return jsonify(ok=False, error="Invalid settings"), 400
    old = load_config()
    for p in masked_paths(cfg):                   # left unchanged: keep the saved (encrypted) value
        set_path(cfg, p, get_path(old, p) or "")
    save_config(cfg)
    return jsonify(ok=True, note="Saved." + (" Changes apply from the next run." if BOT.running else ""))


@app.post("/api/setup")
def api_setup():
    """First-run wizard: merge the essentials into config.yaml."""
    b = request.get_json(force=True)
    cfg = load_config()
    cfg.setdefault("profile", {}).update(name=b.get("name", "").strip(), email=b.get("email", "").strip())
    s = cfg.setdefault("search", {})
    s["keywords"] = b.get("keywords") or s.get("keywords") or []
    s["locations"] = b.get("locations") or s.get("locations") or []
    s["experience_years"] = float(b.get("experience_years") or 0)
    cfg.setdefault("match", {})["skills"] = b.get("skills") or []
    yrs = str(b.get("experience_years") or "0")
    cfg["total_years_answer"] = yrs
    cfg["skill_years"] = {str(sk).split("|")[0].strip(): str(b.get("skill_years") or yrs) for sk in b.get("skills") or []}
    if b.get("title_must_include"):
        cfg["match"]["title_must_include"] = b["title_must_include"]
    ans = cfg.setdefault("answers", {})
    for key, field in (("phone", "phone"), ("notice period", "notice"), ("current ctc", "current_ctc"),
                       ("expected ctc", "expected_ctc"), ("current location", "current_location")):
        if b.get(field) not in (None, ""):
            ans[key] = str(b[field])
    ai = cfg.setdefault("ai", {})
    ai["about_me"] = b.get("about_me") or ai.get("about_me", "")
    if b.get("ai_provider"):
        ai["provider"] = b["ai_provider"]
        if b.get("ai_key") not in (None, "", MASK):
            ai["api_key"] = b["ai_key"]
    er = cfg.setdefault("email_report", {})
    if b.get("email"):
        er.setdefault("send_to", b["email"])
        if not er.get("send_to"):
            er["send_to"] = b["email"]
    cfg["dry_run"] = True       # always start safely
    save_config(cfg)
    return jsonify(ok=True)


@app.post("/api/prefs")
def api_prefs():
    """Small dashboard flags (e.g. "I've signed in to my job sites"), kept in config.yaml."""
    cfg = load_config()
    cfg.setdefault("dashboard", {}).update(request.get_json(force=True) or {})
    save_config(cfg)
    return jsonify(ok=True)


@app.get("/api/sites")
def api_get_sites():
    with open(SITES_FILE, encoding="utf-8") as f:
        return jsonify(text=f.read())


@app.post("/api/sites")
def api_save_sites():
    text = request.get_json(force=True).get("text", "")
    try:
        data = yaml.safe_load(text)
        assert isinstance(data, dict) and "buttons" in data, "sites.yaml must keep the 'buttons' section"
        for k, v in data.items():
            if k != "buttons":
                for need in ("name", "search_url", "job_link_pattern"):
                    assert need in v, f"site '{k}' is missing '{need}'"
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400
    shutil.copy(SITES_FILE, SITES_FILE + ".bak")
    with open(SITES_FILE, "w", encoding="utf-8") as f:
        f.write(text)
    return jsonify(ok=True)


@app.get("/api/learned")
def api_get_learned():
    try:
        with open(LEARNED_FILE, encoding="utf-8") as f:
            return jsonify(yaml.safe_load(f) or {})
    except FileNotFoundError:
        return jsonify({})


@app.get("/api/questions")
def api_questions():
    return jsonify(jobdb.open_questions())


@app.post("/api/questions/answer")
def api_question_answer():
    """Save your answer to a question the bot got stuck on; the jobs it blocked are retried in the next live run."""
    b = request.get_json(force=True) or {}
    q, ans = jobdb.norm_q(b.get("q")), str(b.get("answer") or "").strip()
    if not q or not ans:
        return jsonify(ok=False, error="Type an answer first."), 400
    try:
        with open(LEARNED_FILE, encoding="utf-8") as f:
            learned = yaml.safe_load(f) or {}
    except FileNotFoundError:
        learned = {}
    learned[q] = ans
    with open(LEARNED_FILE, "w", encoding="utf-8") as f:
        f.write("# Answers the bot learned. Edit them in the dashboard.\n")
        yaml.safe_dump({str(k): str(v) for k, v in learned.items()}, f, allow_unicode=True, sort_keys=True)
    return jsonify(ok=True, retry=jobdb.answer_question(q))


@app.post("/api/questions/ignore")
def api_question_ignore():
    jobdb.ignore_question(jobdb.norm_q((request.get_json(force=True) or {}).get("q")))
    return jsonify(ok=True)


@app.post("/api/learned")
def api_save_learned():
    data = request.get_json(force=True) or {}
    with open(LEARNED_FILE, "w", encoding="utf-8") as f:
        f.write("# Answers the bot learned. Edit them in the dashboard.\n")
        yaml.safe_dump({str(k): str(v) for k, v in data.items()}, f, allow_unicode=True, sort_keys=True)
    return jsonify(ok=True)


@app.post("/api/bot/start")
def api_bot_start():
    body = request.get_json(silent=True) or {}
    cfg = load_config()
    if cfg.get("browser_mode", "attach") == "attach" and not chrome_connected():
        ok, msg = open_chrome(with_logins=False)
        if not ok:
            return jsonify(ok=False, error=msg), 400
    ok, msg = BOT.start(dry_run=body.get("dry_run"), parallel=bool(body.get("parallel")), sites=body.get("sites") or None)
    return jsonify(ok=ok, error=None if ok else msg)


@app.post("/api/naukri/resume")
def api_naukri_resume():
    """Upload the dashboard resume to the Naukri profile now."""
    cfg = load_config()
    if not resume_path(cfg):
        return jsonify(ok=False, error="Upload your resume first"), 400
    if cfg.get("browser_mode", "attach") == "attach" and not chrome_connected():
        ok, msg = open_chrome(with_logins=False)
        if not ok:
            return jsonify(ok=False, error=msg), 400
    ok, msg = BOT.start(trigger="manual", task="naukri_resume")
    return jsonify(ok=ok, error=None if ok else msg)


@app.post("/api/bot/stop")
def api_bot_stop():
    return jsonify(ok=BOT.stop())


@app.post("/api/bot/kill")
def api_bot_kill():
    return jsonify(ok=BOT.kill())


@app.post("/api/bot/enter")
def api_bot_enter():
    return jsonify(ok=BOT.send_enter())


@app.get("/api/bot/log")
def api_bot_log():
    since = int(request.args.get("since", 0))
    with LOCK:
        start = max(since - BOT.offset, 0)
        lines = list(BOT.lines)[start:]
        nxt = BOT.offset + len(BOT.lines)
    return jsonify(lines=lines, next=nxt, bot=BOT.state())


@app.get("/api/logs")
def api_logs():
    if not os.path.isdir(LOG_DIR):
        return jsonify([])
    files = sorted(os.listdir(LOG_DIR), reverse=True)[:30]
    return jsonify([{"name": f, "size": os.path.getsize(os.path.join(LOG_DIR, f))} for f in files])


@app.get("/api/logs/<name>")
def api_log_file(name):
    path = os.path.join(LOG_DIR, os.path.basename(name))
    if not os.path.exists(path):
        return "not found", 404
    return send_file(path, mimetype="text/plain")


@app.post("/api/chrome/open")
def api_chrome_open():
    ok, msg = open_chrome()
    return jsonify(ok=ok, message=msg, error=None if ok else msg)


# ------------------------------------------------------------------ logins (sign in to the sites, separately from a run)
@app.get("/api/logins")
def api_logins():
    try:
        extra = yaml.safe_load(open(SITES_FILE, encoding="utf-8")) or {}
    except Exception:
        extra = {}
    return jsonify(sites=site_login.sites(load_config(), extra), chrome=chrome_connected(),
                   chrome_found=bool(find_chrome()))


@app.post("/api/logins/open")
def api_logins_open():
    """Open login tabs: {"sites": ["naukri"]} for one or more sites, nothing for Google + every switched-on site."""
    keys = (request.get_json(silent=True) or {}).get("sites") or None
    ok, msg = open_chrome(with_logins=True, sites=keys)
    return jsonify(ok=ok, message=msg, error=None if ok else msg), (200 if ok else 400)


@app.post("/api/logins/mark")
def api_logins_mark():
    """{"site": "naukri", "signed_in": true}, or {"all": true} for Google + every switched-on site."""
    body = request.get_json(silent=True) or {}
    cfg = load_config()
    on = body.get("signed_in", True)
    keys = [s["key"] for s in site_login.sites(cfg) if s["enabled"]] if body.get("all") else [body.get("site")]
    for k in filter(None, keys):
        site_login.mark(cfg, k, on)
    save_config(cfg)
    return jsonify(ok=True, signed_in=cfg["dashboard"]["signed_in"])


@app.post("/api/report/build")
def api_report_build():
    try:
        import report
        report.build_report()
        return jsonify(ok=True)
    except PermissionError:
        return jsonify(ok=False, error="job_report.xlsx is open in Excel. Close it and try again."), 400
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 500


@app.post("/api/report/email")
def api_report_email():
    try:
        import report
        report.build_report()
        ok = report.send_report()
        return jsonify(ok=bool(ok), error=None if ok else
                       "Email not sent. Check the Email section in Settings (Gmail app password, send to/from).")
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 500


@app.get("/api/report/download")
def api_report_download():
    if not os.path.exists(XLSX_FILE):
        import report
        report.build_report()
    return send_file(XLSX_FILE, as_attachment=True, download_name=f"job_report_{datetime.now():%Y-%m-%d}.xlsx")


@app.get("/api/report/csv")
def api_report_csv():
    """Every job in the database as a CSV file."""
    import io
    rows = jobdb.rows()
    if not rows:
        return "No jobs logged yet.", 404
    buf = io.StringIO()
    cols = ["date", "site", "title", "company", "location", "match_percent", "status", "reason", "attempts", "url"]
    w = csv.DictWriter(buf, fieldnames=cols, extrasaction="ignore")
    w.writeheader()
    w.writerows(rows)
    data = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(data, mimetype="text/csv", as_attachment=True, download_name=f"jobs_{datetime.now():%Y-%m-%d}.csv")


@app.get("/api/package")
def api_package():
    """A clean copy of the bot for someone else: program files only, none of your data."""
    import make_package
    path = make_package.main()
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.post("/api/ai/test")
def api_ai_test():
    """Ask the configured AI one sample question so you know the key works."""
    try:
        cfg = load_config()
        ai = cfg.get("ai") or {}
        code =("import job_bot, json; "
                "print(json.dumps(job_bot.ai_answer('How many years of total experience do you have?', None, True)))")
        out = subprocess.run([sys.executable, "-c", code], cwd=HERE, capture_output=True, text=True, timeout=60,
                             env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        lines = [l for l in out.stdout.strip().splitlines() if l.strip()]
        ans = json.loads(lines[-1]) if lines else None
        if ans:
            return jsonify(ok=True, message=f"AI works ({ai.get('provider')}). Sample answer: {ans}")
        err = out.stderr.strip().splitlines()
        detail = " ".join(lines[:-1]) or (err[-1] if err else "")
        return jsonify(ok=False, error=f"No answer from AI. {detail}")
    except Exception as e:
        return jsonify(ok=False, error=str(e))


# ------------------------------------------------------------------ resume
def resume_path(cfg=None):
    """Full path of the uploaded resume, or None."""
    rel = ((cfg or load_config()).get("resume") or {}).get("file")
    path = os.path.join(HERE, rel) if rel else None
    return path if path and os.path.isfile(path) else None


def extract_text(path):
    """Plain text of a PDF or DOCX resume ("" if it can't be read, e.g. a scanned PDF or an old .doc)."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            from pypdf import PdfReader
            return "\n".join((pg.extract_text() or "") for pg in PdfReader(path).pages[:10])
        if ext == ".docx":
            import zipfile
            from html import unescape
            xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
            xml = re.sub(r"</w:p>|<w:br/>|<w:tab/>", "\n", xml)
            return unescape(re.sub(r"<[^>]+>", "", xml))
    except Exception as e:
        print("Couldn't read resume text:", e)
    return ""


def resume_info(cfg=None):
    cfg = cfg or load_config()
    path = resume_path(cfg)
    if not path:
        return {"has": False}
    r = cfg.get("resume") or {}
    text = open(RESUME_TEXT, encoding="utf-8").read() if os.path.exists(RESUME_TEXT) else ""
    return {"has": True, "name": r.get("name") or os.path.basename(path), "size": os.path.getsize(path),
            "uploaded": r.get("uploaded"), "ext": os.path.splitext(path)[1].lower().lstrip("."),
            "text": text[:20000], "readable": len(text.strip()) > 100}


@app.get("/api/resume")
def api_resume():
    return jsonify(resume_info())


@app.post("/api/resume")
def api_resume_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(ok=False, error="No file chosen"), 400
    original = os.path.basename(f.filename.replace("\\", "/"))
    ext = os.path.splitext(original)[1].lower()
    if ext not in RESUME_TYPES:
        return jsonify(ok=False, error="Please upload a PDF or Word file (.pdf, .docx or .doc)"), 400
    # Job sites show this file name to recruiters, so keep it readable
    stem = re.sub(r"[^\w\-. ]+", "", os.path.splitext(original)[0]).strip(" .") or "Resume"
    os.makedirs(RESUME_DIR, exist_ok=True)
    tmp = os.path.join(RESUME_DIR, ".upload" + ext)
    f.save(tmp)
    with open(tmp, "rb") as fh:
        head = fh.read(8)
    if (ext == ".pdf" and not head.startswith(b"%PDF")) or (ext == ".docx" and not head.startswith(b"PK")):
        os.remove(tmp)
        return jsonify(ok=False, error="That file doesn't look like a real " + ext.upper().lstrip(".") + " file"), 400
    cfg = load_config()
    old = resume_path(cfg)
    final = os.path.join(RESUME_DIR, stem + ext)
    if old and os.path.abspath(old) != os.path.abspath(final):
        os.remove(old)                     # only ever the previous upload, inside resume/
    os.replace(tmp, final)
    with open(RESUME_TEXT, "w", encoding="utf-8") as fh:
        fh.write(re.sub(r"[ \t]+", " ", extract_text(final)).strip())
    cfg["resume"] = {"file": os.path.relpath(final, HERE).replace("\\", "/"), "name": stem + ext,
                     "uploaded": datetime.now().isoformat(timespec="seconds")}
    save_config(cfg)
    return jsonify(ok=True, **resume_info(cfg))


@app.delete("/api/resume")
def api_resume_delete():
    cfg = load_config()
    path = resume_path(cfg)
    if path:
        os.remove(path)
    if os.path.exists(RESUME_TEXT):
        os.remove(RESUME_TEXT)
    cfg.pop("resume", None)
    save_config(cfg)
    return jsonify(ok=True)


@app.get("/api/resume/file")
def api_resume_file():
    path = resume_path()
    if not path:
        return jsonify(ok=False, error="No resume uploaded"), 404
    return send_file(path, as_attachment=request.args.get("download") == "1", download_name=os.path.basename(path))


@app.errorhandler(413)
def too_big(_):
    return jsonify(ok=False, error="That file is too big. Resumes should be under 2 MB."), 413


def main():
    # Windows lets two servers share a port, and the browser then reaches whichever answers, often an old copy
    import socket
    with socket.socket() as s:
        s.settimeout(1)
        if s.connect_ex(("127.0.0.1", PORT)) == 0:
            print(f"\n  The dashboard is already running at http://127.0.0.1:{PORT}. Opening it.\n"
                  "  (To restart it after an update, close its window first.)\n")
            if "--no-browser" not in sys.argv:
                webbrowser.open(f"http://127.0.0.1:{PORT}")
            return
    ensure_config()
    encrypt_saved_secrets()
    threading.Thread(target=scheduler, daemon=True).start()
    url = f"http://127.0.0.1:{PORT}"
    print(f"\n  Job Bot Dashboard running at {url}\n  Keep this window open. Close it to stop the dashboard.\n")
    if "--no-browser" not in sys.argv:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)


if __name__ == "__main__":
    main()
