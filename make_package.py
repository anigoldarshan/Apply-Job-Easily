"""
Makes a clean zip of the job bot to give to someone else.

    python make_package.py

Creates JobBot_<date>.zip with only the program files. Your personal data is
left out: config.yaml (passwords, API keys, phone, CTC), your resume, job database (jobs.db), job history,
learned answers, logs and the Chrome profiles with your logins. The other
person double-clicks Install.bat, and the dashboard's setup wizard creates
their own config.
"""

import os
import zipfile
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
FILES = ["job_bot.py", "jobdb.py", "secret_store.py", "site_login.py", "report.py", "dashboard.py", "make_package.py", "sites.yaml", "config.example.yaml",
         "requirements.txt", "Install.bat", "Start Dashboard.bat", "start_chrome.bat", "README.md",
         "dashboard_ui/index.html"]


def main():
    """Returns the path of the zip it made."""
    out = os.path.join(HERE, f"JobBot_{datetime.now():%Y-%m-%d}.zip")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in FILES:
            path = os.path.join(HERE, f)
            if os.path.exists(path):
                z.write(path, os.path.join("JobBot", f))
            else:
                print("missing (skipped):", f)
    print(f"Created {out}\nSend this zip. The other person unzips it and double-clicks Install.bat.")
    return out


if __name__ == "__main__":
    main()
