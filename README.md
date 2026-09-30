# JobAgent

A job-search assistant that runs inside [Claude Code](https://claude.com/claude-code). You answer a short set of questions once and upload your resume. After that it finds real, currently open roles that fit you, writes a tailored version of your resume for each one, and keeps a spreadsheet of everything.

**You always submit the applications yourself.** JobAgent never fills out an employer's form, creates an account for you, types a password, or presses Submit.

## What you get

- **Setup interview** that captures your target roles, location, pay, and deal-breakers
- **Daily search** that checks employer career pages, screens out duplicates, and scores each role against your criteria
- **Tailored resume per role** that reorders and re-emphasizes what is already true about you, and never invents anything
- **Application folders** with the tailored resume, the apply link, and a saved copy of the posting
- **Tracker spreadsheet** (`Job_Search_Log.xlsx`) with status for every application
- **Optional Gmail add-on** that saves recruiter and confirmation emails to the right folder

## Install

You need [Claude Code](https://claude.com/claude-code), Python 3.10 or newer, and (for resume page checks) the free [LibreOffice](https://www.libreoffice.org/download).

In Claude Code:

```
/plugin marketplace add ryanhohnbaum/job-agent
/plugin install job-agent@job-agent
```

Then say: **"Set up job agent."**

## How it works

```
   Setup (once)          Search (daily or on demand)                 You                    Afterward
 ┌─────────────┐   ┌────────────────────────────────────┐   ┌───────────────────┐   ┌────────────────┐
 │ Interview   │   │ Find roles → verify on employer    │   │ Open Apply.url    │   │ Tell it you    │
 │ Upload      │──▶│ page → skip duplicates → score →   │──▶│ Upload resume     │──▶│ applied; it    │
 │ resume      │   │ tailor resume → make folder        │   │ Answer questions  │   │ updates tracker│
 └─────────────┘   └────────────────────────────────────┘   │ Press Submit      │   └────────────────┘
                                                            └───────────────────┘
```

## What is automatic and what needs you

| Step | Who does it |
|---|---|
| Answering the setup questions | **You** (about 10 minutes, once) |
| Confirming your resume facts are correct | **You** (once, and whenever your history changes) |
| Finding roles and checking they are still open | Agent |
| Skipping roles you already saw or applied to | Agent |
| Scoring fit against your criteria | Agent |
| Tailoring your resume and checking page count | Agent |
| Creating the folder, apply link, and saved posting | Agent |
| Reviewing the roles it recommends | **You** |
| Filling out the employer's application, creating accounts, passwords, screening questions, legal attestations, and clicking Submit | **You, always** |
| Telling the agent you submitted | **You** ("I applied to Acme") |
| Marking it Applied in the tracker | Agent, only after you say so |
| Interview, rejection, and other status updates | You tell it, or the Gmail add-on detects them |
| Running the search every day | Optional. Can be scheduled, but only runs while Claude Code is open on your computer |

## Everyday use

Talk to it in plain language. These all work:

| Say | What happens |
|---|---|
| "Run my job search" | Runs a full search cycle and prepares any strong matches |
| "What should I apply to?" | Shows your queue sorted by fit |
| "Tailor my resume for this job: `<link>`" | Makes a tailored resume for one role you found yourself |
| "I applied to Acme" | Marks it Applied |
| "I got an interview at Acme" / "Acme rejected me" | Updates the status |
| "Skip the Globex one" | Marks it Skipped so it never comes back |
| "Change my search: remote only" | Updates your saved criteria |

## The skills

| Skill | Purpose |
|---|---|
| `jobagent-setup` | One-time interview, resume intake, workspace creation |
| `jobagent-search` | Discover, verify, dedupe, score, and prepare roles |
| `jobagent-tailor` | Truthful resume tailoring with render verification |
| `jobagent-track` | Queue, mark applied, statuses, duplicate checks |
| `jobagent-gmail` | Optional. Read-only Gmail evidence capture |

## Your data stays yours

Everything lives in a folder on your computer (default `~/JobAgent`). The workspace is never part of this repository, and this repo's `.gitignore` blocks it in case you set it up inside a clone. Web pages and emails are treated as untrusted text; instructions found in them are ignored. Gmail access, if you enable it, is read-only and stored locally.

## Honesty rules the agent follows

- Your resume is the source of truth. It will not change employers, titles, dates, education, or numbers.
- It will not claim skills or experience you do not have, even if a posting asks for them. Gaps are reported to you instead.
- It will not lower your fit threshold to hand you more results. Zero results on a slow day is a normal outcome.
- Compensation and travel are reported exactly as posted. If a posting does not say, the agent tells you "not disclosed" or "cannot confirm".

## Limits to know about

- It uses public web search and employer career pages. Some sites block automated access, and search results can be stale, so every role is checked on the employer's own page first.
- Resume page counting needs LibreOffice. Without it, tailoring still works but the page-length check is skipped.
- Saved posting PDFs need the optional `playwright` package (`pip install playwright && playwright install chromium`).
- This is a preparation tool. Review every resume before you send it. You are responsible for what you submit.

## Troubleshooting

Run `doctor` to check your setup:

```
python "<plugin folder>/scripts/jobagent.py" doctor
```

| Problem | Fix |
|---|---|
| "workspace not found" | Start Claude Code from your JobAgent folder, or run setup again |
| "appears open in Excel" | Close `Job_Search_Log.xlsx` and retry |
| Page count check fails | Install LibreOffice |
| Duplicate blocked a role you want | Tell the agent; it will show the existing row and can mark it On Hold or create it as a different role |

## Development

```
scripts/        Python helpers (tracker, resume build/edit/check, posting capture, Gmail)
skills/         The five skills
templates/      Starter config copied into each new workspace
examples/       A fictional resume spec used by tests and by the setup skill
```

`python scripts/jobagent.py --help` lists every command. Pull requests are welcome; keep the manual-apply boundary and truthful-resume rules intact.

## License

MIT
