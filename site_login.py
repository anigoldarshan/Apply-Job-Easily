"""
Sign in to the job sites, separately from running the bot.

Used by the dashboard's Logins page:
  - open one login tab per site in the bot's Chrome (or just one site), even
    when that Chrome is already open, so you can sign in again any time;
  - remember which sites you've confirmed you're signed in to (config.yaml,
    dashboard.logins).

    python site_login.py            open every enabled site's login tab
    python site_login.py naukri     open just that one
"""

import json
import sys
import urllib.parse
import urllib.request
from datetime import datetime

CDP = "http://127.0.0.1:9222"
GOOGLE = {"key": "google", "name": "Google", "url": "https://accounts.google.com"}
LOGIN_URLS = {
    "naukri": "https://www.naukri.com/nlogin/login",
    "linkedin": "https://www.linkedin.com/login",
    "foundit": "https://www.foundit.in/rio/login",
    "shine": "https://www.shine.com/myshine/login/",
    "cutshort": "https://cutshort.io/login",
    "indeed": "https://secure.indeed.com/auth",
}
NAMES = {"naukri": "Naukri", "linkedin": "LinkedIn", "foundit": "Foundit",
         "shine": "Shine", "cutshort": "Cutshort", "indeed": "Indeed"}


def chrome_connected():
    try:
        with urllib.request.urlopen(CDP + "/json/version", timeout=1.5) as r:
            return r.status == 200
    except Exception:
        return False


def sites(cfg, sites_yaml=None):
    """Google plus every site in run order, with its login URL and whether it's switched on."""
    order = cfg.get("run_order") or list(LOGIN_URLS)
    logins = (cfg.get("dashboard") or {}).get("logins") or {}
    out = [dict(GOOGLE, enabled=True, signed_in=logins.get("google"))]
    for key in order:
        url = LOGIN_URLS.get(key) or ((sites_yaml or {}).get(key) or {}).get("login_url")
        if not url:
            continue
        out.append({"key": key, "name": NAMES.get(key) or (sites_yaml or {}).get(key, {}).get("name", key),
                    "url": url, "enabled": bool((cfg.get(key) or {}).get("enabled")),
                    "signed_in": logins.get(key)})
    return out


def open_tab(url):
    """Open a new tab in the bot's Chrome. Newer Chrome wants PUT, older accepts GET."""
    target = CDP + "/json/new?" + urllib.parse.quote(url, safe=":/?&=%")
    for method in ("PUT", "GET"):
        try:
            with urllib.request.urlopen(urllib.request.Request(target, method=method), timeout=5) as r:
                json.loads(r.read() or b"{}")
                return True
        except Exception:
            continue
    return False


def open_tabs(urls):
    """Open each URL in its own tab. Returns how many opened."""
    return sum(open_tab(u) for u in urls)


def mark(cfg, key, signed_in=True):
    """Remember (or forget) that you're signed in to a site. Updates dashboard.signed_in too."""
    dash = cfg.setdefault("dashboard", {})
    logins = dash.setdefault("logins", {})
    if signed_in:
        logins[key] = datetime.now().strftime("%Y-%m-%d %H:%M")
    else:
        logins.pop(key, None)
    on = [s["key"] for s in sites(cfg) if s["enabled"] and s["key"] != "google"]
    dash["signed_in"] = bool(on) and all(k in logins for k in on)
    return cfg


if __name__ == "__main__":
    import os
    import yaml
    here = os.path.dirname(os.path.abspath(__file__))
    cfg = yaml.safe_load(open(os.path.join(here, "config.yaml"), encoding="utf-8")) or {}
    if not chrome_connected():
        sys.exit("The bot's Chrome isn't open. Open it from the dashboard (Logins page) or start_chrome.bat first.")
    wanted = [a.lower() for a in sys.argv[1:]]
    urls = [s["url"] for s in sites(cfg) if (s["key"] in wanted if wanted else s["enabled"])]
    print(f"Opened {open_tabs(urls)} of {len(urls)} login tab(s).")
