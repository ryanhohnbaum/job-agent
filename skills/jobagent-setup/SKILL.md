---
name: jobagent-setup
description: One-time onboarding for JobAgent. Interviews the user about their job-search criteria, takes their master resume (docx, PDF, or pasted text), creates the private workspace and application tracker, and configures everything so daily searching and resume tailoring work. Use when the user says "set up job agent", "start my job search", "onboard me", or when other jobagent skills report no workspace or unfilled config.
---

# JobAgent setup

Goal: leave the user with a private workspace (`~/JobAgent` by default) containing a tracker workbook, a filled-in profile and search rules, and a verified master resume. Be warm and efficient. Ask questions in small batches, not all at once. Never invent an answer for the user.

Script launcher used below (`python3` on macOS/Linux):

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/jobagent.py" <command>
```

## Ground rules

- Everything personal stays in the workspace on the user's machine. Nothing is uploaded anywhere.
- The master resume is the factual source of truth for every later resume. Facts get confirmed by the user before they become the baseline.
- Never write passwords, tokens, or ID numbers into any file.

## Step 1: Environment check (silent unless something is wrong)

1. Confirm Python 3.10+ (`python --version`).
2. Install core packages after telling the user what you are installing: `pip install openpyxl PyYAML python-docx pypdf`. Playwright and Google packages are optional; do not install them now.
3. Check for LibreOffice (`doctor` reports it). It is used to verify resume page count. If missing, tell the user it is free at libreoffice.org and that resume checking will not work until it is installed. Do not block setup on it.

## Step 2: Workspace

Ask where to keep it. Default `~/JobAgent`. Then run:

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/jobagent.py" init --root "<chosen path>"
```

Tell the user to start future sessions from that folder (`cd` into it, then run `claude`), or set `JOB_AGENT_ROOT` to it.

## Step 3: Master resume

Ask for the resume file path (docx preferred; PDF, text, or pasted text also fine).

**If it is a .docx:** copy it to `Templates/<First>_<Last>_Master.docx` (never modify their original). Run `resume dump` and `resume check --text` on it. Report page count. If the docx uses tables, text boxes, columns, or images for layout, warn that ATS systems may misread it and offer to rebuild it into the clean format (below) after they approve.

**If it is a PDF, text, or pasted:** extract the text, then build a clean master:
1. Draft a JSON spec following `${CLAUDE_PLUGIN_ROOT}/examples/master_resume.example.json` (fields: name, headline, contact, lines, skills, sections with paragraphs, bullets, or jobs). Use only content from the resume. Do not improve, embellish, or add metrics.
2. Show the user a readable summary and get corrections.
3. Run `resume build spec.json Templates/<First>_<Last>_Master.docx`, then `resume check` and read the rendered PDF to inspect it visually.

**Fact confirmation (required).** Read back a compact list: each employer, title, and dates; education and certifications; every quantified claim. Ask the user to confirm or correct. Apply corrections to the master. This confirmed list is the truth boundary.

**Extra facts (optional but valuable).** Ask: "Is there work you have done that is not on this resume but you would want available when tailoring for a specific job? For example projects, tools, or accomplishments cut for space." Save confirmed items to `Config/extra_facts.md`, one bullet each, with any numbers exactly as the user states them. The tailor skill may draw on these; it may never draw on anything that is not on the resume or in this file.

Record the page count of the final master. Resumes tailored later will match it.

## Step 4: Interview

Ask in these batches. Offer examples, accept short answers, and reflect back a summary at the end of each batch. Skip anything the user says does not apply.

**Batch A, the basics:** full name as it should appear, email, phone, city and state, LinkedIn URL, authorized to work in the U.S. (and whether sponsorship is needed now or later). If outside the U.S., ask for the country and adapt wording; the workflow is not U.S.-specific.

**Batch B, what they want:**
- Target titles and functions (be specific: "Product Marketing Manager", not "marketing").
- Seniority levels to target, and levels to avoid (too junior or too senior).
- Should roles with a different title but a clear fit be included? (default yes)
- Role types to exclude, in plain language.
- Companies to exclude (current employer, past bad experiences, conflicts).

**Batch C, where and how:**
- Home market and which nearby cities count as commutable.
- Remote OK? Hybrid OK, and within what area?
- Relocation: no, or where?
- Maximum travel percent.

**Batch D, money and industries:**
- Minimum compensation, and whether that means base or total. It is fine to have no minimum.
- Industries in ranked preference, and whether strong fits outside them are welcome.
- Any other deal-breakers.

**Batch E, honesty guardrails:**
- "Is there anything a posting might ask for that you do NOT have and do not want implied on your resume?" (These become `truth_boundary.unsupported_claims_to_avoid`.)
- One sentence on how they want to be positioned to employers.

**Batch F, cadence:**
- How selective? Explain the fit score (1 to 10) and that only roles at or above the threshold get a tailored resume. Recommend 8.0 to start; lower it (7.0 to 7.5) if they want more volume.
- Would they like a daily automatic search? Explain honestly: it can run on a schedule through Claude Code's scheduling features, but only while the app is running on their machine.

## Step 5: Write config

Edit `Config/profile.yaml` and `Config/search_rules.yaml` with the answers. Keep the template's keys. Remove no keys; empty lists are fine. Set `master_resume.file` and `required_length_pages`. Then show the user the two files in plain language (not raw YAML dumps of every key) and ask for corrections.

## Step 6: Verify and hand off

1. Run `jobagent.py doctor`. Fix any FAIL. Warnings about optional packages are fine.
2. Summarize what exists and where.
3. Explain the three things they will do next, in this order:
   - "Run `jobagent-search` to find and prepare your first roles."
   - "Open the application folders the agent creates, use `Apply.url` to reach the employer's page, and apply yourself."
   - "Tell me when you have submitted, and I will mark it Applied."
4. State clearly what the agent will not do: it never fills employer forms, creates employer accounts, enters passwords, or presses Submit.
5. Mention the optional Gmail skill in one sentence; do not push it.

Offer to run the first search immediately.
