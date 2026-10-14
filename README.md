# Job Auto-Apply (Naukri, LinkedIn, Foundit, Shine, Cutshort, Indeed)

This bot logs in to your accounts, searches with your keywords, location and experience, keeps only the jobs that match your skills (50%+ by default), and applies. It runs through all six sites one after another, in the order you set. Everything runs on your own computer, so your passwords never leave it.

## The dashboard (easiest way)

1. Double-click **Install.bat** once. It installs everything and puts a **Job Bot Dashboard** shortcut on your desktop.
2. Open the dashboard from that shortcut (or double-click **Start Dashboard.bat**). It opens in your browser at http://127.0.0.1:8765 and only runs on your own computer.
3. The first time, a setup wizard asks for your name and experience, lets you pick your kind of job (it fills in typical titles and skills for you), and asks the common application answers. It creates your own `config.yaml`.
4. The **Home** page always shows the one thing to do next: sign in to your job sites, try a test run, then start applying.

| Page | What you can do |
|---|---|
| Home | Your next step, a setup checklist, applied / apply-yourself / interview counts, last 14 days, jobs to apply yourself, follow-ups due |
| Activity | Watch a run live: current site, jobs checked, applied, skipped, a plain-English feed, and a **Continue** button when the bot needs you |
| Jobs | Every job as a card with its match % and *why* it was applied or skipped. Mark "I applied", "Interview", "Offer"… and add notes. Export to Excel/CSV or email |
| Questions | Questions job forms asked that the bot couldn't answer. Answer once: those jobs are retried in the next live run, and the bot answers that question by itself from then on |
| Settings | Simple forms that save automatically: about you, jobs to find, skills & matching, answers, job sites, AI helper, email report, daily schedule, safety limits |

The wrench icon at the bottom left shows **advanced options** (extra settings, answer memory, site rules). The moon icon switches light/dark, and the **?** icon opens help.
You get a chime (and a desktop notification) when the bot needs you or a run finishes.

### Applying automatically, never twice

Every job the bot looks at is saved in a small local database, **jobs.db** (SQLite, next to the bot; your old applied_jobs.csv history is imported the first time).

- A **live run** applies automatically to every job whose match is at least your "How picky" setting (50% by default).
- It first applies to the good matches your **test runs** found, then searches for new jobs.
- Jobs already applied to are never opened again, and neither is the same role at the same company posted on another site.
- When the bot tries but can't finish (the job is on the company's own website, or the form asks something it can't answer), the job goes to **Apply manually** with the reason, e.g. "Tried automatically, but the form asked questions it couldn't answer. Please apply manually." These are listed on the Jobs page, at the end of the run log, in the Excel report's *Apply Manually* sheet and in the report email.

Check the database any time with `python jobdb.py`.

### Your resume

Upload it in the setup wizard or in **Settings → About you** (PDF or Word; under 2 MB is accepted everywhere). Then:
- whenever an application form asks for a resume file, the bot attaches it (cover-letter and photo fields are left alone, and LinkedIn's already-selected resume is kept);
- the AI helper reads it when answering unusual questions;
- skills found in it are offered as one-click suggestions in **Skills & matching**.

**Naukri profile:** turn on *Keep my Naukri resume up to date* (Settings → About you) and each live run first uploads this resume to your Naukri profile, at most once a day; recently updated profiles rank higher in recruiter searches. The *Update my resume on Naukri now* button does it right away.

It's stored in the `resume/` folder on your computer. Naukri and LinkedIn usually apply with the resume saved on your profile there, so keep that one current too.

### Giving the bot to someone else

Run `python make_package.py`. It creates `JobBot_<date>.zip` with only the program files. Your config, passwords, keys, resume, job history, learned answers and Chrome logins are **not** included. The other person unzips it, double-clicks **Install.bat**, and fills in their own details in the setup wizard.

## 1. One-time setup (without the dashboard) (Windows / Mac / Linux)

Install Python 3.9 or newer from python.org (on Windows, tick "Add to PATH"). Then open a terminal inside this folder and run:

```
pip install -r requirements.txt
```

The bot uses the **Google Chrome** already installed on your PC, so there's no separate browser download. If you don't have Chrome, it uses Microsoft Edge instead, which comes with Windows.

## 2. Fill in config.yaml

For each site, set `enabled: true` or `false`, plus your email and password. `run_order` controls which sites run and in what order. You can leave the password blank and type it yourself when the browser opens. Then set your keywords, locations, years of experience, skills and answers. The **skills** list drives the match %: if you list 8 skills and a job mentions 4 of them, that job is a 50% match.

## 3. First run: dry run

`dry_run: true` is on by default. Run:

```
python job_bot.py
```

A Chrome window opens. Complete any OTP or captcha, then press Enter in the terminal. The bot scans jobs and writes every one to **applied_jobs.csv** with its match % and what it *would* do. Open the CSV and check that the matches look right. If they don't, adjust your keywords, skills or title filters.

## 4. Go live

Set `dry_run: false` and run again. Your login is remembered in the `browser_profile` folder, so you usually won't need to log in again.

## Excel report (job_report.xlsx)

After every run the bot creates **job_report.xlsx** in this folder. It covers all your runs, not just the latest. It has 4 sheets:

| Sheet | What's in it |
|---|---|
| Summary | Per site: jobs applied (total and today), jobs to apply manually, interviews, jobs scanned |
| Applied Jobs | Every successful application: date, site, title, company, location, match %, a clickable job link, and a follow-up date 7 days later |
| Apply Manually | Good matches the bot couldn't finish (external company sites or unanswered questions), best matches first |
| All Jobs Scanned | Everything the bot looked at, including skips and the reason |

In **Applied Jobs**, the yellow **Status** (dropdown: Applied, Viewed, Interview, Offer, Rejected, No reply) and **Notes** columns are yours to update. Your edits are kept the next time the report is rebuilt. Rows turn amber when a follow-up is due.

Close the Excel file before running the bot, otherwise Windows won't let it save the new report. To rebuild the report without running the bot, run `python report.py`.

## Report by email

After each run, the report is emailed to the address in `email_report.send_to` (Settings → Email report in the dashboard). The email shows a per-site summary and the list of jobs applied in that run, with the Excel file attached.

Gmail needs an **App Password** for this. Your normal password won't work. It's a one-time setup:

1. Turn on 2-Step Verification at https://myaccount.google.com/security
2. Go to https://myaccount.google.com/apppasswords, create a password (name it "Job bot"), and copy the 16 letters.
3. Paste them into `app_password` under `email_report` in config.yaml.

To send the latest report by email without running the bot, run `python report.py --email`. If you'd rather only get an email when something was actually applied, set `only_if_new_applications: true`.

## What happens to each job

| Status in CSV | Meaning |
|---|---|
| applied | Application sent |
| match (dry run) | Would apply when live |
| match: apply on company site / not Easy Apply (manual) | A good match, but it needs an external form. Do these by hand |
| needs manual (questions) | The application asked something the bot couldn't answer. Add that question to `answers` in config |
| skipped: … | Filtered out by match %, title, company or experience |

## Questions during applying (fully automatic)

The bot answers each question in this order:

1. **learned_answers.yaml**: answers it has saved before.
2. **answers** in config.yaml, such as phone, notice period, CTC and skill years.
3. **AI**, if you add a key in the `ai` section. It answers from your `about_me` profile, picks one of the offered options, and never invents experience.
4. **A safe default**: Yes for yes/no questions, 0 for numbers, and NA for text.

Every AI answer is saved to learned_answers.yaml, so each new question needs only one AI call. You can open that file any time to check or correct answers.

If none of these fit (for example a dropdown whose choices don't match any of your answers), the bot leaves that application for you and lists the question on the dashboard's **Questions** page. Answer it there once: the jobs it blocked go back in the queue for the next live run, and the answer is saved to learned_answers.yaml.

## Keeping your accounts safe

- **Passwords and API keys are encrypted** in config.yaml with your Windows login (they look like `enc:AQAAANCM...`). They only work for your Windows account on this computer, and the dashboard never shows them again after you save them. Older plain-text values are encrypted automatically the next time the dashboard starts.
- **Daily limit**: `max_applications_per_day` (default 30) is per site and counts every run that day, so running twice doesn't double it. `max_applications_per_site` still limits a single run.
- **Rest after each application**: a random pause between `rest_after_apply_min` and `rest_after_apply_max` seconds (default 20–45), like a person would take.

## Every site gets a turn

With `rotate_sites: true`, the bot runs keyword 1 on all sites, then keyword 2 on all sites, and so on, instead of spending the whole session on Naukri. A job link is never opened twice, and the same title at the same company is never applied to twice, even when it's posted on another site.

## How each site applies

| Site | What the bot does |
|---|---|
| Naukri | Clicks Apply and answers the follow-up chatbot questions |
| LinkedIn | Easy Apply jobs only; fills each step and submits |
| Foundit, Shine, Cutshort | Clicks Apply and fills any form or popup that appears |
| Indeed | Jobs with "Apply now" (Indeed Apply); "Apply on company site" jobs go on your manual list |

Indeed has the strongest bot protection of the six and often shows "verify you are human" checks. If it keeps blocking you, set `indeed: enabled: false`.

## Fixing a site yourself (sites.yaml)

Foundit, Shine, Cutshort and Indeed are defined in **sites.yaml**, not in code. Each entry has a `search_url` template, a `job_link_pattern` that recognises job links on the results page, and lists of selectors for the title, company and description. The `buttons` section at the bottom holds the wording the bot looks for, such as "Apply", "Next" and "Submit".

To fix a site, do a search on it in your browser and copy the results-page URL into `search_url`, swapping your words for placeholders like `{kw}` and `{loc}`. Then open one job and check its URL matches `job_link_pattern`.

You can also add a new job site the same way, by copying an entry, renaming it, and adding the name to `run_order` and to the accounts in config.yaml.

## Keep your accounts safe

LinkedIn and Naukri don't allow automated applying and may restrict accounts that do it. To lower the risk, keep `max_applications_per_site` around 20–30 per run, run it at most once or twice a day, keep the delays, and stay near the computer for OTPs.

Answer questions truthfully in `answers`. Wrong numbers, such as claiming years you don't have, get your applications rejected at the screening stage anyway.

## If something stops working

These sites change their page layouts often. If the bot finds "No results" or a button stops being clicked, edit sites.yaml for Foundit, Shine, Cutshort or Indeed. For Naukri and LinkedIn, the selectors are in `job_bot.py` (for example `a.title`, `#apply-button`, `button.jobs-apply-button`). If one site fails, the bot logs the error and moves on to the next site. Send me the error or a screenshot and I'll fix it.
