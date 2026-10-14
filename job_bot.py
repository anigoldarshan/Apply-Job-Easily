"""
Job Auto-Apply bot for Naukri, LinkedIn (Easy Apply), Foundit, Shine,
Cutshort and Indeed.

Runs on YOUR computer in a visible Chrome window, so you can watch it and
handle OTPs / captchas yourself. Settings live in config.yaml.

    pip install playwright pyyaml openpyxl
    python -m playwright install chromium
    python job_bot.py

Every job it looks at is written to applied_jobs.csv.
Naukri and LinkedIn have dedicated handlers below. Foundit, Shine, Cutshort
and Indeed are driven by sites.yaml, so you can fix them without code.
"""

import csv
import os
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import quote

import yaml
import jobdb
import secret_store
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = secret_store.reveal(yaml.safe_load(open(os.path.join(HERE, "config.yaml"), encoding="utf-8")))
if secret_store.FAILED:
    print("Couldn't unlock these saved passwords/keys (config.yaml came from another computer or Windows user):",
          ", ".join(secret_store.FAILED), "- enter them again in the dashboard Settings.")
SITES = yaml.safe_load(open(os.path.join(HERE, "sites.yaml"), encoding="utf-8"))
BTN = {k: re.compile(v, re.I) for k, v in SITES.pop("buttons").items()}
LOG_FILE = os.path.join(HERE, "applied_jobs.csv")
PROFILE_DIR = os.path.join(HERE, "browser_profile")  # keeps you logged in between runs

# The dashboard can force a test run without changing config.yaml
DRY = (os.environ["JOBBOT_DRY_RUN"] == "1") if "JOBBOT_DRY_RUN" in os.environ else bool(CFG.get("dry_run", True))
STOP_FILE = os.path.join(HERE, ".stop_bot")   # created by the dashboard's Stop button
MAX_PER_SITE = int(CFG.get("max_applications_per_site", 25))
MAX_PER_DAY = int(CFG.get("max_applications_per_day", 30) or 0)  # per site, counting earlier runs today; 0 = no limit
REST = (CFG.get("rest_after_apply_min", 20), CFG.get("rest_after_apply_max", 45))   # seconds, after each application
S, M = CFG["search"], CFG["match"]
ANSWERS = {str(k).lower(): str(v) for k, v in (CFG.get("answers") or {}).items()}
ANSWER_KEYS = sorted(ANSWERS, key=len, reverse=True)

# Answers you give by hand are remembered here and reused automatically next time.
LEARNED_FILE = os.path.join(HERE, "learned_answers.yaml")
try:
    LEARNED = yaml.safe_load(open(LEARNED_FILE, encoding="utf-8")) or {}
except FileNotFoundError:
    LEARNED = {}
WAIT_SECS = int(CFG.get("wait_for_me_minutes", 5)) * 60
MAX_MIN_PER_SITE = float(CFG.get("max_minutes_per_site", 0) or 0)

# Resume uploaded in the dashboard: attached to forms that ask for one, and read by the AI helper.
_resume_rel = (CFG.get("resume") or {}).get("file")
RESUME = os.path.join(HERE, _resume_rel) if _resume_rel and os.path.isfile(os.path.join(HERE, _resume_rel)) else None
try:
    RESUME_TEXT = open(os.path.join(HERE, "resume", "resume.txt"), encoding="utf-8").read()[:8000] if RESUME else ""
except FileNotFoundError:
    RESUME_TEXT = ""


class SkipJob(Exception):
    """Raised when you choose to skip a job while the bot is waiting."""


CURRENT = {}   # title/company of the job being processed, for the log


norm_q = jobdb.norm_q


def remember(question, answer):
    q = norm_q(question)
    if not q or not str(answer or "").strip():
        return
    if LEARNED.get(q) == answer:
        return
    LEARNED[q] = str(answer).strip()
    with open(LEARNED_FILE, "w", encoding="utf-8") as f:
        f.write("# Answers the bot learned from you. Edit or delete lines freely.\n")
        yaml.safe_dump(LEARNED, f, allow_unicode=True, sort_keys=True)
    print(f"  (learned: \"{question.strip()[:60]}\" -> {answer})")


def beep():
    try:
        import winsound
        winsound.Beep(1200, 250); winsound.Beep(1500, 250)
    except Exception:
        print("\a", end="", flush=True)


def site_time_up(started):
    return MAX_MIN_PER_SITE and (time.time() - started) > MAX_MIN_PER_SITE * 60


def pick_option(want, options):
    """Index of the option that matches `want` (exact first, then partial)."""
    if not want:
        return None
    w = want.strip().lower()
    for i, o in enumerate(options):
        if o.strip().lower() == w:
            return i
    for i, o in enumerate(options):
        if w in o.lower() or (len(o) > 1 and o.strip().lower() in w):
            return i
    return None


def is_yes_no(options):
    low = {o.strip().lower() for o in options}
    return bool(low) and low <= {"yes", "no"}


# ------------------------------------------------------------------ helpers
def pause(a=None, b=None):
    if os.path.exists(STOP_FILE):
        raise KeyboardInterrupt
    time.sleep(random.uniform(a or CFG.get("min_delay_seconds", 4), b or CFG.get("max_delay_seconds", 9)))


def slug(text):
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")


def has_term(term, text):
    """Whole-term match that also works for things like 'c#' and '.net core'."""
    return re.search(r"(?<![a-z0-9])" + re.escape(term) + r"(?![a-z0-9#+])", text) is not None


def match_percent(description):
    """Each skill may list alternatives separated by '|', e.g. 'web api|webapi|rest api'."""
    skills = M.get("skills", [])
    if not skills:
        return 100, []
    d = description.lower()
    found = []
    for s in skills:
        alts = [a.strip().lower() for a in str(s).split("|") if a.strip()]
        if any(has_term(a, d) for a in alts):
            found.append(alts[0])
    return round(100 * len(found) / len(skills)), found


ALL_KEYWORDS = list(S.get("keywords") or [])      # S["keywords"] changes per round when sites rotate
# Words that don't say what the job is: "Senior Android Engineer" is an "Android Developer" job
GENERIC_TITLE_WORDS = {"developer", "developers", "development", "engineer", "engineers", "engineering", "senior", "sr",
                       "junior", "jr", "lead", "software", "specialist", "consultant", "associate", "programmer",
                       "professional", "expert", "staff", "sde", "the", "and", "of", "for", "with", "in", "a", "an",
                       "i", "ii", "iii", "iv", "1", "2", "3", "remote", "hybrid", "role", "job", "opening", "hiring"}
COMPOUNDS = {"back end": "backend", "front end": "frontend", "full stack": "fullstack", "dev ops": "devops",
             "node js": "nodejs", "react js": "reactjs", "vue js": "vuejs", "next js": "nextjs", "dot net": ".net"}


def norm_title(text):
    t = re.sub(r"[\-_/|,()\[\]:]+", " ", (text or "").lower())
    t = re.sub(r"\s+", " ", t).strip()
    for a, b in COMPOUNDS.items():
        t = t.replace(a, b)
    return t


def title_core(keyword):
    """'Senior Android Developer' -> 'android'; 'Full Stack Developer' -> 'fullstack'; 'Software Engineer' stays."""
    words = norm_title(keyword).split()
    core = [w for w in words if w not in GENERIC_TITLE_WORDS]
    return " ".join(core) if core else " ".join(words)


TITLE_CORES = [c for c in dict.fromkeys(title_core(k) for k in ALL_KEYWORDS) if c]


def title_problem(title):
    """Why a job title doesn't fit your search, or "" if it does."""
    t = norm_title(title)
    if t in ("", "?"):
        return ""
    mine = " ".join(norm_title(k) for k in ALL_KEYWORDS)
    for w in M.get("title_exclude", []) or []:
        w = norm_title(str(w))
        # a skip word that is part of your own job titles (e.g. "android" while searching "Android Developer") is ignored
        if w and has_term(w, t) and not has_term(w, mine):
            return f"title has '{w}'"
    if M.get("title_matches_search", True) and TITLE_CORES:
        extra = [norm_title(str(w)) for w in (M.get("title_must_include") or [])]
        if not any(has_term(c, t) for c in TITLE_CORES + [x for x in extra if x]):
            return "title isn't one of your job titles"
    else:
        inc = [norm_title(str(w)) for w in (M.get("title_must_include") or [])]
        if inc and not any(has_term(w, t) for w in inc):
            return "title filter"
    return ""


def title_ok(title):
    return not title_problem(title)


def skip_reason(title, company):
    return f"skipped: {title_problem(title) or 'company filter'}"


def company_ok(company):
    c = company.lower()
    return not any(x.lower() in c for x in M.get("company_exclude", []))


def my_years():
    return float(S["experience_years"])


def search_years():
    """Whole years for site search filters (4.6 -> 4)."""
    return int(my_years())


def experience_ok(exp_text):
    """Sites show e.g. '3-6 Yrs'. Accept if your years fall inside (with 1 year slack)."""
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", exp_text or "")]
    if len(nums) < 2:
        return True
    return nums[0] - 1 <= my_years() <= nums[1] + 1


SKILL_YEARS = sorted(((str(k).lower(), str(v)) for k, v in (CFG.get("skill_years") or {}).items()),
                     key=lambda kv: len(kv[0]), reverse=True)
GENERAL_EXP = re.compile(r"\b(total|overall|professional|work|relevant|industry|it|software)\b")


def answer_for(question):
    if norm_q(question) in LEARNED:
        return str(LEARNED[norm_q(question)])
    q = (question or "").lower()
    if "year" in q or "how many" in q:
        # "How many years of C# experience?" -> your C# years
        for skill, yrs in SKILL_YEARS:
            if has_term(skill, q):
                return yrs
        # "How many years of (total/work) experience?" -> total years,
        # but "years of Kafka experience" / "experience with Kafka" -> a tech you didn't list
        generic = {"work", "total", "overall", "professional", "relevant", "industry", "it", "software",
                   "the", "your", "you", "do", "have", "a", "an", "experience", "years", "year", "in", "of"}
        named = []
        m = re.search(r"years? of (.*?)\bexperience", q)
        if m: named += re.findall(r"[a-z0-9#+.]+", m.group(1))
        m = re.search(r"experience\b.*?\b(?:with|in|using|on)\s+([a-z0-9#+. ]+)", q)
        if m: named += re.findall(r"[a-z0-9#+.]+", m.group(1))
        if "experience" in q and all(w in generic for w in named):
            return str(CFG.get("total_years_answer", S["experience_years"]))
        # a technology you didn't list -> honest default
        return str(CFG.get("default_years", "0"))
    for key in ANSWER_KEYS:
        if key in q:
            return ANSWERS[key]
    return None



# ------------------------------------------------------------------ AI answers
AI = CFG.get("ai") or {}
AI_PROVIDER = str(AI.get("provider", "none")).lower()
# tolerate values like "Gemini API Key" or "Anthropic Claude"
AI_PROVIDER = next((p for p, words in (("claude", ("claude", "anthropic")), ("gemini", ("gemini", "google")),
                                       ("openai", ("openai", "gpt", "chatgpt"))) if any(w in AI_PROVIDER for w in words)),
                   "none")
AI_DEFAULT_MODEL = {"claude": "claude-haiku-4-5-20251001", "gemini": "gemini-2.5-flash", "openai": "gpt-4o-mini"}
AI_FAILED = False


def _post_json(url, body, headers):
    import json
    import urllib.request
    req = urllib.request.Request(url, data=json.dumps(body).encode(), method="POST",
                                 headers={"Content-Type": "application/json", **headers})
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.loads(r.read().decode())


def ai_answer(question, options=None, numeric=False):
    """Ask an AI model to answer an application question from your profile. Returns text or None."""
    global AI_FAILED
    key = str(AI.get("api_key") or "").strip()
    if AI_PROVIDER in ("", "none") or not key or AI_FAILED:
        return None
    model = AI.get("model") or AI_DEFAULT_MODEL.get(AI_PROVIDER, "")
    opts = ("\nOptions (reply with exactly one of these, copied exactly):\n- " + "\n- ".join(options)) if options else ""
    kind = "\nThis field accepts a NUMBER only: reply with digits only." if numeric else ""
    prompt = f"""You fill in job application forms for the candidate below. Answer the question truthfully from the profile.
Rules:
- Reply with ONLY the answer. No quotes, no explanation.
- Never invent experience. Years in a technology not in the profile = 0.
- Willingness questions (relocate, shifts, work from office, notice, background check): answer positively unless the profile says otherwise.
- If the profile doesn't cover it, give the most sensible neutral answer (for text: NA).
- Text answers: at most 25 words. Dates: DD/MM/YYYY.

CANDIDATE PROFILE
{AI.get("about_me", "")}
{("RESUME" + chr(10) + RESUME_TEXT) if RESUME_TEXT else ""}

QUESTION: {question}{opts}{kind}"""
    try:
        if AI_PROVIDER == "claude":
            r = _post_json("https://api.anthropic.com/v1/messages",
                           {"model": model, "max_tokens": 100, "messages": [{"role": "user", "content": prompt}]},
                           {"x-api-key": key, "anthropic-version": "2023-06-01"})
            text = "".join(b.get("text", "") for b in r.get("content", []))
        elif AI_PROVIDER == "gemini":
            r = _post_json(f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}",
                           {"contents": [{"parts": [{"text": prompt}]}]}, {})
            text = r["candidates"][0]["content"]["parts"][0]["text"]
        elif AI_PROVIDER == "openai":
            r = _post_json("https://api.openai.com/v1/chat/completions",
                           {"model": model, "messages": [{"role": "user", "content": prompt}]},
                           {"Authorization": f"Bearer {key}"})
            text = r["choices"][0]["message"]["content"]
        else:
            return None
    except Exception as e:
        print(f"  (AI not available: {str(e)[:120]}. Check the ai section of config.yaml. Continuing without AI.)")
        AI_FAILED = True
        return None
    text = text.strip().strip('"').strip()
    if numeric:
        m = re.search(r"\d+(\.\d+)?", text)
        text = m.group(0) if m else "0"
    if options:
        i = pick_option(text, options)
        text = options[i] if i is not None else None
    if text:
        print(f"  (AI answered: \"{question.strip()[:60]}\" -> {text})")
    return text or None


def resolve_answer(question, options=None, numeric=False):
    """Learned answers -> your config answers -> AI -> safe default. Returns text or None."""
    if not (question or "").strip():
        return None
    ans = answer_for(question)
    if options:
        if ans is not None and pick_option(ans, options) is not None:
            return options[pick_option(ans, options)]
        ai = ai_answer(question, options)
        if ai:
            remember(question, ai)
            return ai
        if is_yes_no(options):
            return options[pick_option(CFG.get("default_yes_no", "Yes"), options)]
        return None
    if ans is not None:
        return ans
    ai = ai_answer(question, None, numeric)
    if ai:
        remember(question, ai)
        return ai
    if CFG.get("fill_unknown", True):
        return "0" if numeric else str(CFG.get("unknown_text_answer", "NA"))
    return None


APPLIED_KEYS = set()   # (title, company) pairs already applied to, on any site
RUN_MANUAL = []        # jobs this run couldn't apply to by itself


def job_key(title, company):
    return (re.sub(r"\W+", " ", (title or "").lower()).strip(), re.sub(r"\W+", " ", (company or "").lower()).strip())


def load_seen():
    """Job URLs to leave alone this run, from the local database (jobs.db)."""
    APPLIED_KEYS.update(jobdb.applied_keys())
    return jobdb.skip_urls(recheck=CFG.get("recheck_old_jobs", False), dry=DRY)


def duplicate(title, company):
    """Same role at the same company already applied (reposted ad, or listed on another site)."""
    return title not in ("", "?") and job_key(title, company) in APPLIED_KEYS


def log(site, title, company, location, pct, status, url):
    new = not os.path.exists(LOG_FILE)
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["date", "site", "title", "company", "location", "match_percent", "status", "url"])
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M"), site, title, company, location, pct, status, url])
    try:
        jobdb.record(site, title, company, location, pct, status, url)
    except Exception as e:
        print("  (couldn't save to jobs.db:", e, ")")
    if status == "applied":
        APPLIED_KEYS.add(job_key(title, company))
    if status.startswith(("match:", "needs manual")):
        RUN_MANUAL.append((site, title, company, status, url))
    print(f"  [{status}] {pct}%  {title} | {company}")
    if status.startswith("needs manual") and STUCK:
        try:
            jobdb.record_questions(site, url, list(STUCK.values()))
            print("   stuck on:", "; ".join(q for q, _, _ in STUCK.values())[:200],
                  "-> answer it in the dashboard (Questions) and the bot will retry this job")
        except Exception as e:
            print("  (couldn't save the questions:", e, ")")
    STUCK.clear()
    if status == "applied" and not DRY:
        rest_after_apply()


STUCK = {}   # questions the current job's form asked that the bot couldn't answer: {norm_q: (text, kind, options)}


def note_stuck(text, kind="text", options=None):
    text = (text or "").strip()
    if text and norm_q(text):
        STUCK[norm_q(text)] = (text, kind, [o for o in (options or []) if o])


def rest_after_apply():
    """Rest a little after each application, like a person would. Sites restrict accounts that apply too fast."""
    try:
        a, b = float(REST[0] or 0), float(REST[1] or 0)
    except (TypeError, ValueError):
        a, b = 20, 45
    end = time.time() + random.uniform(a, max(a, b))
    while time.time() < end:
        if os.path.exists(STOP_FILE):
            raise KeyboardInterrupt
        time.sleep(1)


def all_text(page, selectors):
    """Join the text of every selector that exists (description + key skills + details)."""
    parts = []
    for sel in selectors:
        try:
            loc = page.locator(sel)
            for i in range(min(loc.count(), 3)):
                t = loc.nth(i).inner_text(timeout=2000).strip()
                if t and t not in parts:
                    parts.append(t)
        except Exception:
            pass
    return "\n".join(parts)


def text_of(page, selectors, default=""):
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count():
                t = loc.inner_text(timeout=2000).strip()
                if t:
                    return t
        except Exception:
            pass
    return default


def safe_goto(page, url):
    """Open a page, tolerating slow sites: one retry, then carry on with whatever has loaded."""
    for attempt in (1, 2):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            return
        except PWTimeout:
            if attempt == 2:
                try:
                    page.wait_for_selector("h1, a", timeout=10000)
                except Exception:
                    raise
            else:
                print("  (slow page, retrying...)")


def wait_for_login(site):
    where = "click Continue in the dashboard" if os.environ.get("JOBBOT_FROM_DASHBOARD") else "press Enter here"
    print(f"\n>> Finish logging in to {site} in the browser window (OTP/captcha if asked), then {where}... ", flush=True)
    input()


# ------------------------------------------------------------------ NAUKRI
def naukri_login(page):
    if CFG.get("skip_login"):
        return
    cfg = CFG["naukri"]
    page.goto("https://www.naukri.com/nlogin/login", wait_until="domcontentloaded")
    pause(2, 4)
    if "mnjuser" in page.url:
        print("Naukri: already logged in.")
        return
    try:
        page.fill("#usernameField", cfg.get("email", ""), timeout=8000)
        if cfg.get("password"):
            page.fill("#passwordField", cfg["password"])
            page.click("button[type=submit]")
    except PWTimeout:
        pass
    wait_for_login("Naukri")


def naukri_search_urls(keyword, location, page_no):
    base = f"https://www.naukri.com/{slug(keyword)}-jobs-in-{slug(location)}"
    if page_no > 1:
        base += f"-{page_no}"
    return (f"{base}?k={quote(keyword)}&l={quote(location)}"
            f"&experience={search_years()}&jobAge={S.get('posted_within_days', 7)}")


NK_DRAWER = "[class*='chatbot_Drawer'], [class*='chatbot_drawer']"
def _last_text(page, selector):
    try:
        loc = page.locator(selector)
        return loc.last.inner_text(timeout=2000).strip() if loc.count() else ""
    except Exception:
        return ""


def nk_question(page):
    return _last_text(page, "li.botItem, [class*='botMsg']")


def nk_last_reply(page):
    return _last_text(page, "li.userItem, [class*='userMsg']")


def nk_open(page):
    d = page.locator(NK_DRAWER)
    return d.count() > 0 and d.first.is_visible()


def nk_options(page):
    """Clickable choices in the Naukri chatbot (radio buttons, checkboxes or chips)."""
    out, seen = [], set()
    for sel in ("[class*='chatbot'] label", "[class*='chatbot'] [class*='chip']",
                "[class*='chatbot'] [class*='radio'] span", "[class*='chatbot'] [class*='Chip']"):
        loc = page.locator(sel)
        for i in range(min(loc.count(), 20)):
            el = loc.nth(i)
            try:
                t = el.inner_text().strip()
                if t and t.lower() not in seen and el.is_visible() and len(t) < 80:
                    seen.add(t.lower()); out.append((t, el))
            except Exception:
                pass
    return out


def wait_for_you(page, still_waiting, what):
    """Pause until you answer in the browser. Returns when still_waiting() is False."""
    beep()
    print(f"\n  >>> YOUR TURN: {what}")
    print(f"      Answer it in the Chrome window. I'll continue automatically "
          f"(waiting up to {WAIT_SECS // 60} min).")
    end = time.time() + WAIT_SECS
    while time.time() < end:
        time.sleep(2)
        try:
            if not still_waiting():
                print("  Thanks, continuing...")
                return True
        except Exception:
            return True
    print("  No answer in time; moving to the next job.")
    return False


def naukri_handle_chatbot(page):
    """Answer Naukri's question drawer. Unknown questions wait for you, and your answer is learned."""
    last_q, repeats = None, 0
    for _ in range(25):
        if not nk_open(page):
            return True
        question = nk_question(page)
        if question == last_q:
            repeats += 1
        else:
            last_q, repeats = question, 0
        opts = nk_options(page)
        texts = [t for t, _ in opts]
        ans = resolve_answer(question, texts or None)
        idx = pick_option(ans, texts) if (opts and ans) else None

        need_you = repeats >= 2 or (opts and idx is None) or (not opts and ans is None)
        if need_you and CFG.get("wait_for_me", False):
            choices = f"  Options: {', '.join(texts)}" if texts else ""
            ok = wait_for_you(page, lambda: nk_open(page) and nk_question(page) == question,
                              f"Naukri asks: \"{question[:150]}\"{choices}")
            if not ok:
                return False
            reply = nk_last_reply(page)
            if reply and reply != question:
                remember(question, reply)
            continue
        if need_you:
            note_stuck(question, "choice" if texts else "text", texts)
            return False

        if opts:
            opts[idx][1].click()
        else:
            box = page.locator("[class*='chatbot'] [contenteditable='true'], [class*='chatbot'] input[type=text]")
            if box.count():
                box.first.click(); box.first.fill(ans)
        send = page.locator("[class*='sendMsg'], [class*='chatbot'] button:has-text('Save')")
        if send.count():
            send.first.click()
        pause(2, 3)
    return not nk_open(page)


def run_naukri(page, seen):
    naukri_login(page)
    applied, started = SITE_DONE.get("naukri", 0), time.time()
    for kw in S["keywords"]:
        for loc in S["locations"]:
            for p in range(1, int(S.get("pages_per_search", 2)) + 1):
                if applied >= site_cap("naukri") or site_time_up(started):
                    return applied
                url = naukri_search_urls(kw, loc, p)
                print(f"\nNaukri search: {kw} / {loc} / page {p}")
                safe_goto(page, url); pause(3, 5)
                links = page.eval_on_selector_all("a.title", "els => els.map(e => e.href)")
                if not links:
                    print("  No results (or the page layout changed)."); break
                for job_url in links:
                    if applied >= site_cap("naukri") or site_time_up(started):
                        return applied
                    if job_url in seen:
                        continue
                    seen.add(job_url)
                    CURRENT.clear()
                    try:
                        applied += naukri_job(page, job_url)
                    except Exception as e:
                        log("Naukri", CURRENT.get("title", "?"), CURRENT.get("company", "?"), loc,
                            CURRENT.get("pct", 0), f"error: {type(e).__name__}", job_url)
                    pause()
    return applied


def naukri_job(page, url):
    safe_goto(page, url); pause(2, 4)
    title = text_of(page, ["h1"], "?")
    company = text_of(page, ["[class*='jd-header-comp-name'] a", "[class*='comp-name']"], "?")
    location = text_of(page, ["[class*='location'] a", "[class*='loc']"], "")
    exp = text_of(page, ["[class*='exp'] span", "[class*='exp']"], "")
    desc = all_text(page, ["section[class*='job-desc']", "[class*='key-skill']", "[class*='keySkill']",
                           "[class*='other-details']", "[class*='JDC']"]) or text_of(page, ["main"], "")
    pct, _ = match_percent(desc + " " + title)
    CURRENT.update(title=title, company=company, pct=pct)

    if not title_ok(title) or not company_ok(company):
        log("Naukri", title, company, location, pct, skip_reason(title, company), url); return 0
    if not experience_ok(exp):
        log("Naukri", title, company, location, pct, f"skipped: experience {exp}", url); return 0
    if pct < M.get("min_match_percent", 50):
        log("Naukri", title, company, location, pct, "skipped: low match", url); return 0
    if duplicate(title, company):
        log("Naukri", title, company, location, pct, "already applied (same job, other post)", url); return 0

    if page.locator("#already-applied, button:has-text('Applied')").count():
        log("Naukri", title, company, location, pct, "already applied", url); return 0
    if page.locator("#company-site-button").count() and not page.locator("#apply-button").count():
        log("Naukri", title, company, location, pct, "match: apply on company site (manual)", url); return 0
    if DRY:
        log("Naukri", title, company, location, pct, "match (dry run)", url); return 0

    btn = page.locator("#apply-button")
    if not btn.count():
        log("Naukri", title, company, location, pct, "match: no apply button / walk-in (manual)", url); return 0
    btn.first.click(); pause(3, 5)
    finished = naukri_handle_chatbot(page)
    ok = (page.get_by_text(re.compile("successfully applied", re.I)).count() > 0
          or page.locator("button:has-text('Applied')").count() > 0)
    status = "applied" if (ok or finished) else "needs manual (questions)"
    log("Naukri", title, company, location, pct, status, url)
    return 1 if status == "applied" else 0


NAUKRI_RESUME_STAMP = os.path.join(HERE, ".naukri_resume_updated")   # date of the last update


def naukri_update_resume(page, force=False):
    """Upload your resume to your Naukri profile (Profile -> Update resume).
    Naukri ranks recently updated profiles higher in recruiter searches. Once a day unless force=True."""
    if not RESUME:
        print("Naukri resume: no resume uploaded in the dashboard (Settings -> About you).")
        return False
    today = datetime.now().strftime("%Y-%m-%d")
    if not force and os.path.exists(NAUKRI_RESUME_STAMP) and open(NAUKRI_RESUME_STAMP).read().strip() == today:
        print("Naukri resume: already updated today.")
        return True
    print(f"Naukri resume: uploading {os.path.basename(RESUME)} to your profile...")
    naukri_login(page)
    for attempt in range(2):
        safe_goto(page, "https://www.naukri.com/mnjuser/profile"); pause(3, 5)
        if "login" not in page.url:
            break
        wait_for_login("Naukri")
    try:
        box = page.locator("input#attachCV, input[type=file][id*='attach' i], input[type=file]").first
        try:
            box.wait_for(state="attached", timeout=15000)
        except PWTimeout:
            print("Naukri resume: couldn't find the 'Update resume' upload on your profile page "
                  "(Naukri may have changed it). Update it yourself at naukri.com/mnjuser/profile.")
            return False
        box.set_input_files(RESUME)
        done = re.compile(r"(resume|cv)\b.{0,60}(uploaded|updated)|successfully (uploaded|updated)", re.I)
        for _ in range(20):
            time.sleep(1)
            if page.get_by_text(done).count():
                break
        else:
            print("Naukri resume: uploaded, but Naukri didn't show a confirmation. Check your profile page.")
            return False
        with open(NAUKRI_RESUME_STAMP, "w") as f:
            f.write(today)
        print("Naukri resume: updated on your profile.")
        return True
    except Exception as e:
        print("Naukri resume: couldn't update it (", str(e).splitlines()[0][:120], "). Update it yourself on naukri.com.")
        return False


# ------------------------------------------------------------------ LINKEDIN
def linkedin_login(page):
    if CFG.get("skip_login"):
        return
    cfg = CFG["linkedin"]
    page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded"); pause(2, 4)
    if "/feed" in page.url:
        print("LinkedIn: already logged in.")
        return
    try:
        page.fill("#username", cfg.get("email", ""), timeout=8000)
        if cfg.get("password"):
            page.fill("#password", cfg["password"])
            page.click("button[type=submit]")
    except PWTimeout:
        pass
    wait_for_login("LinkedIn")


def linkedin_exp_filter():
    y = my_years()
    if y <= 1: return "1,2"
    if y <= 4: return "2,3"
    if y <= 8: return "3,4"
    return "4,5"


def run_linkedin(page, seen):
    linkedin_login(page)
    applied, started = SITE_DONE.get("linkedin", 0), time.time()
    secs = int(S.get("posted_within_days", 7)) * 86400
    for kw in S["keywords"]:
        for loc in S["locations"]:
            for p in range(int(S.get("pages_per_search", 2))):
                if applied >= site_cap("linkedin") or site_time_up(started):
                    return applied
                url = (f"https://www.linkedin.com/jobs/search/?keywords={quote(kw)}&location={quote(loc)}"
                       f"&f_AL=true&f_TPR=r{secs}&f_E={quote(linkedin_exp_filter())}&start={p * 25}")
                print(f"\nLinkedIn search: {kw} / {loc} / page {p + 1}")
                safe_goto(page, url); pause(3, 5)
                for _ in range(6):  # scroll the results list so all cards load
                    page.mouse.wheel(0, 1500); time.sleep(0.8)
                ids = page.eval_on_selector_all(
                    "[data-occludable-job-id], [data-job-id]",
                    "els => [...new Set(els.map(e => e.getAttribute('data-occludable-job-id') || e.getAttribute('data-job-id')))]")
                ids = [i for i in ids if i and i.isdigit()]
                if not ids:
                    print("  No results (or the page layout changed)."); break
                for jid in ids:
                    if applied >= site_cap("linkedin") or site_time_up(started):
                        return applied
                    job_url = f"https://www.linkedin.com/jobs/view/{jid}/"
                    if job_url in seen:
                        continue
                    seen.add(job_url)
                    CURRENT.clear()
                    try:
                        applied += linkedin_job(page, job_url)
                    except Exception as e:
                        log("LinkedIn", CURRENT.get("title", "?"), CURRENT.get("company", "?"), loc,
                            CURRENT.get("pct", 0), f"error: {type(e).__name__}", job_url)
                        li_discard(page)
                    pause()
    return applied


def linkedin_job(page, url):
    safe_goto(page, url); pause(2, 4)
    title = text_of(page, ["h1"], "?")
    company = text_of(page, [".job-details-jobs-unified-top-card__company-name", "[class*='company-name']"], "?")
    location = text_of(page, [".job-details-jobs-unified-top-card__primary-description-container",
                              "[class*='top-card'] [class*='bullet']"], "")[:60]
    desc = text_of(page, ["#job-details", ".jobs-description__content", "article"], "")
    pct, _ = match_percent(desc + " " + title)
    CURRENT.update(title=title, company=company, pct=pct)

    if not title_ok(title) or not company_ok(company):
        log("LinkedIn", title, company, location, pct, skip_reason(title, company), url); return 0
    if pct < M.get("min_match_percent", 50):
        log("LinkedIn", title, company, location, pct, "skipped: low match", url); return 0
    if duplicate(title, company):
        log("LinkedIn", title, company, location, pct, "already applied (same job, other post)", url); return 0

    btn = page.locator("button.jobs-apply-button:has-text('Easy Apply')")
    if not btn.count():
        status = "already applied" if page.locator("text=/Applied \\d/").count() else "match: not Easy Apply (manual)"
        log("LinkedIn", title, company, location, pct, status, url); return 0
    if DRY:
        log("LinkedIn", title, company, location, pct, "match (dry run)", url); return 0

    btn.first.click(); pause(2, 3)
    status = li_easy_apply(page)
    log("LinkedIn", title, company, location, pct, status, url)
    return 1 if status == "applied" else 0


def _label_for(dialog, el):
    fid = el.get_attribute("id") or ""
    if fid:
        lab = dialog.locator(f"label[for='{fid}']")
        if lab.count():
            return lab.first.inner_text().strip()
    return (el.get_attribute("aria-label") or el.get_attribute("placeholder") or "").strip()


def fill_form(dialog):
    """Fill empty inputs, dropdowns and radio groups inside `dialog`.
    Returns {question: element} for fields it couldn't answer."""
    unanswered = {}
    yes = CFG.get("default_yes_no", "Yes")
    for field in dialog.locator("input[type=text], input[type=tel], input[type=number], input[type=email], input:not([type]), textarea").all():
        try:
            if not field.is_visible() or field.input_value().strip():
                continue
            label = _label_for(dialog, field)
            numeric = (field.get_attribute("type") or "") == "number"
            ans = ANSWERS.get("phone") if "phone" in label.lower() or "mobile" in label.lower() else resolve_answer(label, None, numeric)
            if ans is not None:
                field.fill(ans)
            elif label:
                unanswered[label] = field
                note_stuck(label, "text")
        except Exception:
            pass
    for sel in dialog.locator("select").all():
        try:
            current = sel.evaluate("s => s.options[s.selectedIndex] ? s.options[s.selectedIndex].text : ''")
            if current and "select" not in current.lower():
                continue
            label = _label_for(dialog, sel)
            opts = [o for o in sel.evaluate("s => [...s.options].map(o => o.text.trim())") if "select" not in o.lower()]
            i = pick_option(resolve_answer(label, opts), opts)
            if i is not None:
                sel.select_option(label=opts[i])
            elif label:
                unanswered[label] = sel
                note_stuck(label, "choice", opts)
        except Exception:
            pass
    for fs in dialog.locator("fieldset").all():
        try:
            if fs.locator("input[type=radio]:checked").count() or not fs.locator("input[type=radio]").count():
                continue
            legend = fs.locator("legend").first.inner_text().strip() if fs.locator("legend").count() else ""
            labels = fs.locator("label").all()
            texts = [l.inner_text().strip() for l in labels]
            i = pick_option(resolve_answer(legend, texts), texts)
            if i is not None:
                labels[i].click()
            elif legend:
                unanswered[legend] = fs
                note_stuck(legend, "choice", texts)
        except Exception:
            pass
    attach_resume(dialog)
    try:
        cb = dialog.locator("#follow-company-checkbox")
        if cb.count() and cb.is_checked():
            dialog.locator("label[for='follow-company-checkbox']").click()
    except Exception:
        pass
    return unanswered


def attach_resume(scope):
    """Upload your resume into empty file fields (not cover letters / photos).
    Skipped when the site already shows a resume picked from your profile."""
    if not RESUME:
        return
    ext = os.path.splitext(RESUME)[1].lower()
    try:
        if scope.locator("[class*='document-upload'][class*='selected'], input[type=radio][id*='resume' i]:checked").count():
            return
        for inp in scope.locator("input[type=file]").all():
            if inp.get_attribute("data-jobbot-resume") or inp.evaluate("e => e.files.length"):
                continue
            hint = " ".join(filter(None, [_label_for(scope, inp), inp.get_attribute("name"), inp.get_attribute("id"),
                                          inp.evaluate("e => (e.closest('div,section,label') || e).innerText.slice(0, 200)")])).lower()
            accept = (inp.get_attribute("accept") or "").lower()
            if re.search(r"cover|photo|picture|image|avatar|transcript", hint):
                continue
            if accept and ext not in accept and "application/" not in accept and "*" not in accept:
                continue
            inp.set_input_files(RESUME)
            inp.evaluate("e => e.setAttribute('data-jobbot-resume', '1')")
            print(f"   attached resume: {os.path.basename(RESUME)}")
            pause(2, 3)
    except Exception as e:
        print("   couldn't attach resume:", str(e)[:80])


def learn_from(unanswered):
    """After you filled fields by hand, save what you entered."""
    for q, el in unanswered.items():
        try:
            tag = el.evaluate("e => e.tagName.toLowerCase()")
            if tag == "fieldset":
                chk = el.locator("input[type=radio]:checked")
                if chk.count():
                    rid = chk.first.get_attribute("id") or ""
                    lab = el.locator(f"label[for='{rid}']") if rid else el.locator("label").first
                    remember(q, lab.first.inner_text().strip())
            elif tag == "select":
                v = el.evaluate("s => s.options[s.selectedIndex] ? s.options[s.selectedIndex].text.trim() : ''")
                if v and "select" not in v.lower():
                    remember(q, v)
            else:
                remember(q, el.input_value())
        except Exception:
            pass


def all_filled(unanswered):
    for el in unanswered.values():
        try:
            tag = el.evaluate("e => e.tagName.toLowerCase()")
            if tag == "fieldset":
                if not el.locator("input[type=radio]:checked").count():
                    return False
            elif tag == "select":
                v = el.evaluate("s => s.options[s.selectedIndex] ? s.options[s.selectedIndex].text : ''")
                if not v or "select" in v.lower():
                    return False
            elif not el.input_value().strip():
                return False
        except Exception:
            pass
    return True


def ask_for_fields(site_name, unanswered):
    """Wait while you fill the questions the bot doesn't know, then learn your answers."""
    if not unanswered or not CFG.get("wait_for_me", False):
        return bool(not unanswered)
    qs = "; ".join(q[:70] for q in unanswered)
    ok = wait_for_you(None, lambda: not all_filled(unanswered),
                      f"{site_name} asks: {qs}. Just fill these in; don't click Next, I'll do that.")
    if ok:
        time.sleep(1.5)
        learn_from(unanswered)
    return ok


def li_easy_apply(page):
    dialog = page.locator("div[role=dialog]").first
    stuck = 0
    for _ in range(15):
        unanswered = fill_form(dialog)
        if unanswered and not ask_for_fields("LinkedIn", unanswered):
            li_discard(page); return "needs manual (questions)"
        for label in ("Submit application", "Review your application", "Continue to next step"):
            b = dialog.locator(f"button[aria-label='{label}']")
            if b.count():
                b.first.click(); pause(2, 3)
                if label == "Submit application":
                    pause(2, 3)
                    try:
                        page.locator("button[aria-label='Dismiss']").first.click(timeout=4000)
                    except Exception:
                        pass
                    return "applied"
                break
        else:
            li_discard(page); return "needs manual (unknown step)"
        if dialog.locator(".artdeco-inline-feedback--error").count():
            stuck += 1
            note_error_fields(dialog)
            if stuck >= 2:
                if CFG.get("wait_for_me", False) and wait_for_you(
                        page, lambda: dialog.locator(".artdeco-inline-feedback--error").count() > 0,
                        "LinkedIn shows an error on this step. Fix the red field(s); don't click Next."):
                    stuck = 0
                    continue
                li_discard(page); return "needs manual (questions)"
    li_discard(page)
    return "needs manual (too many steps)"


def note_error_fields(dialog):
    """LinkedIn rejected an answer (red message): remember which questions, so you can answer them."""
    try:
        for label in dialog.locator(".artdeco-inline-feedback--error").evaluate_all(
                """els => els.map(e => { const c = e.closest('[data-test-form-element], .fb-dash-form-element,'
                + ' .jobs-easy-apply-form-element, fieldset') || e.parentElement;
                const l = c && c.querySelector('label, legend'); return l ? l.innerText.trim() : ''; })"""):
            note_stuck(label, "text")
    except Exception:
        pass


def li_discard(page):
    try:
        page.locator("button[aria-label='Dismiss']").first.click(timeout=3000); pause(1, 2)
        page.locator("button:has-text('Discard')").first.click(timeout=3000)
    except Exception:
        pass


# ------------------------------------------------------------------ GENERIC SITES (sites.yaml)
def fill_template(tpl, kw, loc, page_no, page_size):
    return tpl.format(kw=quote(kw), kw_slug=slug(kw), loc=quote(loc), loc_slug=slug(loc),
                      exp=search_years(), days=S.get("posted_within_days", 7),
                      page=page_no, start=(page_no - 1) * (page_size or 0))


def generic_login(page, key, site):
    if CFG.get("skip_login"):
        return
    cfg = CFG.get(key, {})
    page.goto(site["login_url"], wait_until="domcontentloaded"); pause(2, 4)
    try:
        email = page.locator("input[type=email], input[name*=email i], input[id*=email i], input[name*=user i]").first
        if email.count() and email.is_visible():
            email.fill(cfg.get("email", ""))
        pw_box = page.locator("input[type=password]").first
        if cfg.get("password") and pw_box.count() and pw_box.is_visible():
            pw_box.fill(cfg["password"])
            page.locator("button[type=submit], button:has-text('Login'), button:has-text('Log in'), button:has-text('Sign in')").first.click()
    except Exception:
        pass
    wait_for_login(site["name"])


def collect_job_links(page, site):
    for _ in range(int(site.get("scrolls", 4))):
        page.mouse.wheel(0, 1800); time.sleep(0.9)
    hrefs = page.eval_on_selector_all("a[href]", "els => els.map(e => e.href)")
    pat = re.compile(site["job_link_pattern"], re.I)
    out = []
    for h in hrefs:
        if not pat.search(h):
            continue
        if site.get("job_id_regex"):
            m = re.search(site["job_id_regex"], h)
            if not m:
                continue
            h = site["job_url_template"].format(id=m.group(1))
        else:
            h = h.split("?")[0].split("#")[0]
        if h not in out:
            out.append(h)
    return out


def find_button(scope, rx):
    for role in ("button", "link"):
        loc = scope.get_by_role(role, name=rx)
        for i in range(loc.count()):
            el = loc.nth(i)
            try:
                if el.is_visible():
                    return el
            except Exception:
                pass
    return None


def run_generic(page, key, seen):
    site = SITES[key]
    generic_login(page, key, site)
    applied, started = SITE_DONE.get(key, 0), time.time()
    pages = 1 if not site.get("page_size") else int(S.get("pages_per_search", 2))
    for kw in S["keywords"]:
        for loc in S["locations"]:
            for p in range(1, pages + 1):
                if applied >= site_cap(key) or site_time_up(started):
                    return applied
                url = fill_template(site["search_url"], kw, loc, p, site.get("page_size"))
                print(f"\n{site['name']} search: {kw} / {loc} / page {p}")
                safe_goto(page, url); pause(3, 5)
                links = collect_job_links(page, site)
                if not links:
                    print("  No results (or the page layout changed — check sites.yaml)."); break
                for job_url in links:
                    if applied >= site_cap(key) or site_time_up(started):
                        return applied
                    if job_url in seen:
                        continue
                    seen.add(job_url)
                    CURRENT.clear()
                    try:
                        applied += generic_job(page, key, site, job_url, loc)
                    except Exception as e:
                        log(site["name"], CURRENT.get("title", "?"), CURRENT.get("company", "?"), loc,
                            CURRENT.get("pct", 0), f"error: {type(e).__name__}", job_url)
                    pause()
    return applied


def generic_job(page, key, site, url, search_loc):
    n = site["name"]
    safe_goto(page, url); pause(2, 4)
    title = text_of(page, site.get("title", ["h1"]), "?")[:120]
    company = text_of(page, site.get("company", []), "?").split("\n")[0][:80]
    exp = text_of(page, site.get("experience", []), "")[:40]
    desc = text_of(page, site.get("description", ["main"]), "") or page.locator("body").inner_text()
    pct, _ = match_percent(desc + " " + title)
    CURRENT.update(title=title, company=company, pct=pct)
    loc = search_loc

    if not title_ok(title) or not company_ok(company):
        log(n, title, company, loc, pct, skip_reason(title, company), url); return 0
    if exp and not experience_ok(exp):
        log(n, title, company, loc, pct, f"skipped: experience {exp}", url); return 0
    if pct < M.get("min_match_percent", 50):
        log(n, title, company, loc, pct, "skipped: low match", url); return 0
    if duplicate(title, company):
        log(n, title, company, loc, pct, "already applied (same job, other post)", url); return 0
    if find_button(page, BTN["already_applied"]):
        log(n, title, company, loc, pct, "already applied", url); return 0

    apply_btn = find_button(page, BTN["apply"])
    if not apply_btn:
        status = "match: apply on company site (manual)" if find_button(page, BTN["external"]) else "match: no apply button (manual)"
        log(n, title, company, loc, pct, status, url); return 0
    if DRY:
        log(n, title, company, loc, pct, "match (dry run)", url); return 0

    status = generic_apply(page, site, apply_btn)
    log(n, title, company, loc, pct, status, url)
    return 1 if status == "applied" else 0


def generic_apply(page, site, apply_btn):
    ctx = page.context
    before = len(ctx.pages)
    apply_btn.click(); pause(3, 5)
    target = page
    if len(ctx.pages) > before:                       # apply opened a new tab
        target = ctx.pages[-1]
        try:
            target.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass
        if site["home_domain"] not in target.url:     # external company site
            target.close()
            return "match: apply on company site (manual)"
    try:
        for _ in range(12):
            if target.get_by_text(BTN["success_text"]).count():
                return "applied"
            dialogs = target.locator("div[role=dialog]:visible")
            scope = dialogs.first if dialogs.count() else target
            unanswered = fill_form(scope)
            if unanswered and not ask_for_fields(site["name"], unanswered):
                return "needs manual (questions)"
            b = find_button(scope, BTN["next_step"]) or find_button(scope, BTN["submit"])
            if not b:
                break
            b.click(); pause(2, 4)
        done = lambda: target.get_by_text(BTN["success_text"]).count() or find_button(target, BTN["already_applied"])
        if done():
            return "applied"
        if CFG.get("wait_for_me", False) and wait_for_you(
                target, lambda: not done(),
                f"{site['name']}: I couldn't finish this form. Complete and submit it yourself."):
            return "applied"
        return "needs manual (questions)"
    finally:
        if target is not page:
            try: target.close()
            except Exception: pass


CAP = {}   # most applications allowed per site in this run


def site_name(key):
    return {"naukri": "Naukri", "linkedin": "LinkedIn"}.get(key) or SITES.get(key, {}).get("name", key)


def site_cap(key):
    """max_applications_per_site for this run, lowered by the daily limit (applications already sent today)."""
    if key not in CAP:
        cap = MAX_PER_SITE
        if MAX_PER_DAY and not DRY:
            done = jobdb.applied_today(site_name(key))
            cap = min(cap, max(0, MAX_PER_DAY - done))
            if cap == 0:
                print(f"{site_name(key)}: daily limit reached ({done} applied today, limit {MAX_PER_DAY}). Skipping until tomorrow.")
            elif cap < MAX_PER_SITE:
                print(f"{site_name(key)}: {done} applied today, daily limit {MAX_PER_DAY}, so up to {cap} more.")
        CAP[key] = cap
    return CAP[key]


def apply_pending(page, key, seen):
    """Live run: apply to the good matches earlier test runs found (stored in jobs.db)."""
    name = {"naukri": "Naukri", "linkedin": "LinkedIn"}.get(key) or SITES[key]["name"]
    todo = [r for r in jobdb.pending_matches(name) if r["url"] not in seen]
    applied = SITE_DONE.get(key, 0)
    if not todo or applied >= site_cap(key):
        return applied
    print(f"\n{name}: applying to {len(todo)} good match{'es' if len(todo) != 1 else ''} found by your test runs")
    if key == "naukri":
        naukri_login(page)
    elif key == "linkedin":
        linkedin_login(page)
    else:
        generic_login(page, key, SITES[key])
    for r in todo:
        if applied >= site_cap(key):
            break
        seen.add(r["url"])
        CURRENT.clear()
        try:
            if key == "naukri":
                applied += naukri_job(page, r["url"])
            elif key == "linkedin":
                applied += linkedin_job(page, r["url"])
            else:
                applied += generic_job(page, key, SITES[key], r["url"], r.get("location") or "")
        except KeyboardInterrupt:
            raise
        except Exception as e:
            log(name, CURRENT.get("title", r["title"]), CURRENT.get("company", r["company"]), r.get("location") or "",
                CURRENT.get("pct", r["match_percent"]), f"error: {type(e).__name__}", r["url"])
            if key == "linkedin":
                li_discard(page)
        pause()
    return applied


# ------------------------------------------------------------------ main
ATTACHED = False
SITE_DONE = {}   # applications per site in this run (all rounds)


def open_browser(pw):
    """
    browser_mode "attach" (recommended): connect to the Chrome you opened with
    start_chrome.bat. It's a normal Chrome, so Google sign-in works like usual.
    Otherwise: launch Chrome/Edge directly.
    """
    global ATTACHED
    if CFG.get("browser_mode", "attach") == "attach":
        try:
            browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222", timeout=8000)
            ATTACHED = True
            print("Browser: connected to your Chrome (start_chrome.bat)")
            return browser.contexts[0] if browser.contexts else browser.new_context()
        except Exception:
            sys.exit("\nChrome isn't open in bot mode.\n"
                     "  1. Close ALL Chrome windows.\n"
                     "  2. Double-click start_chrome.bat in this folder.\n"
                     "  3. Run python job_bot.py again.\n")
    wanted = CFG.get("browser", "chrome")
    order = [wanted] + [b for b in ("chrome", "msedge", "chromium") if b != wanted]
    for channel in order:
        try:
            ctx = pw.chromium.launch_persistent_context(
                PROFILE_DIR, channel=None if channel == "chromium" else channel,
                headless=False, viewport={"width": 1300, "height": 900},
                ignore_default_args=["--enable-automation"],
                args=["--disable-blink-features=AutomationControlled"])
            print(f"Browser: {channel}")
            return ctx
        except Exception as e:
            print(f"Browser '{channel}' not available ({str(e).splitlines()[0][:90]})")
    sys.exit("No browser found. Install Google Chrome from https://www.google.com/chrome and run again.")


def update_naukri_resume_only():
    """python job_bot.py --naukri-resume : just update the resume on your Naukri profile."""
    with sync_playwright() as pw:
        ctx = open_browser(pw)
        page = ctx.new_page() if ATTACHED else (ctx.pages[0] if ctx.pages else ctx.new_page())
        page.set_default_timeout(20000)
        try:
            return 0 if naukri_update_resume(page, force=True) else 1
        finally:
            if ATTACHED:
                try: page.close()
                except Exception: pass
            else:
                ctx.close()


def main():
    if "--naukri-resume" in sys.argv:
        return update_naukri_resume_only()
    print("Mode:", "DRY RUN (no applications sent)" if DRY else "LIVE (will apply)")
    missing = [k for k, v in ANSWERS.items() if v.strip().upper() == "FILL_ME"]
    if missing and not DRY:
        print("Please fill these answers in config.yaml before going live:", ", ".join(missing))
        return 1
    seen = load_seen()
    run_started = datetime.now().replace(second=0, microsecond=0)
    with sync_playwright() as pw:
        ctx = open_browser(pw)
        page = ctx.new_page() if ATTACHED else (ctx.pages[0] if ctx.pages else ctx.new_page())
        page.set_default_timeout(20000)
        totals = {}
        try:
            order = [k for k in (CFG.get("run_order") or ["naukri", "linkedin", *SITES.keys()])
                     if CFG.get(k, {}).get("enabled")]
            all_kw = list(S["keywords"])
            # rotate: keyword 1 on every site, then keyword 2 on every site ... so all sites get a turn
            rounds = [[kw] for kw in all_kw] if CFG.get("rotate_sites", True) else [all_kw]
            names = {"naukri": "Naukri", "linkedin": "LinkedIn", **{k: v["name"] for k, v in SITES.items()}}
            for n_round, kws in enumerate(rounds):
                S["keywords"] = kws
                for key in order:
                    if SITE_DONE.get(key, 0) >= site_cap(key):
                        continue
                    print(f"\n========== {names.get(key, key).upper()}  ({', '.join(kws)}) ==========")
                    try:
                        if key == "naukri" and n_round == 0 and not DRY and CFG.get("naukri", {}).get("update_resume"):
                            naukri_update_resume(page)
                        if not DRY and n_round == 0:
                            SITE_DONE[key] = apply_pending(page, key, seen)
                            if SITE_DONE.get(key, 0) >= site_cap(key):
                                continue
                        if key == "naukri":
                            SITE_DONE[key] = run_naukri(page, seen)
                        elif key == "linkedin":
                            SITE_DONE[key] = run_linkedin(page, seen)
                        elif key in SITES:
                            SITE_DONE[key] = run_generic(page, key, seen)
                        else:
                            print(f"Unknown site '{key}' in run_order; skipping.")
                    except KeyboardInterrupt:
                        raise
                    except Exception as e:
                        print(f"{key}: stopped after an error ({str(e).splitlines()[0][:150]}). Moving to the next site.")
                    totals[names.get(key, key)] = SITE_DONE.get(key, 0)
        except KeyboardInterrupt:
            print("\nStopped by you.")
        finally:
            print("\nDone. Applications sent:", totals or 0)
            if RUN_MANUAL:
                print(f"\nCouldn't apply automatically to {len(RUN_MANUAL)} good match(es). Please apply manually:")
                for site, title, company, status, url in RUN_MANUAL:
                    print(f"  - {site}: {title} | {company}\n      {jobdb.manual_reason(status)}\n      {url}")
            print("Saved in the local database:", jobdb.DB_FILE)
            if ATTACHED:
                try: page.close()
                except Exception: pass
            else:
                ctx.close()
            try:
                from report import build_report, send_report
                build_report()
                send_report(since=run_started)
            except Exception as e:
                print("Couldn't build the Excel report:", e, "(run: python report.py)")


if __name__ == "__main__":
    sys.exit(main())
