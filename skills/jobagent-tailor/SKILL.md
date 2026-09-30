---
name: jobagent-tailor
description: Tailor the user's master resume to a specific job posting without inventing anything. Clones the master, edits emphasis, headline, and skills to match the role, keeps formatting and page count, renders it to verify, and audits every claim against the master. Use for "tailor my resume for this job", or when jobagent-search needs a resume for a selected role.
---

# JobAgent resume tailoring

The master resume is the factual source. A tailored resume is a truthful derivative: it changes emphasis and wording, never facts.

Launcher (`python3` on macOS/Linux), run from the workspace folder:

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/jobagent.py" resume <dump|apply|check> ...
```

## Inputs

- The job posting (text or URL), plus company, exact title, and requisition ID.
- `Config/profile.yaml` (master path, page count, `truth_boundary`), the master docx in `Templates/`, and `Config/extra_facts.md` if present.

## What you may change

- The headline
- Executive profile or summary wording and emphasis
- Skills line: selection and order
- Bullet order and which bullets get space
- Terminology, mapped honestly to the posting's vocabulary when the meaning is the same
- Which supported accomplishments are most prominent

## What you must never do

- Change employer names, titles, dates, education, certifications, or any number, team size, budget, or outcome
- Add experience, tools, or credentials that are not in the master or `extra_facts.md`
- Imply anything listed in `truth_boundary.unsupported_claims_to_avoid`
- Claim direct experience because the posting asks for it. If the posting requires something the resume does not support, leave it as a gap and report it.
- Adopt a job title the user has never held. A headline may describe positioning ("Operations Leader"), not claim a title.
- Write in a keyword-dump style. It must read like a person wrote it.

If a draft conflicts with the master, the master wins.

## Procedure

1. **Baseline.** `resume check "<master>"` and note the page count. The tailored resume must match it.
2. **Read the master structure.** `resume dump "<master>"` lists every paragraph with an index. Indices refer to this master.
3. **Analyze the posting.** Identify the 5 to 8 things it cares about most and map each to real evidence in the master. Identify unsupported requirements (gaps).
4. **Write an edits file** (JSON) in a temp folder:

```json
{
  "replace": {
    "1": "New headline text",
    "6": "New executive profile text",
    "12": [["Bold label: ", true], ["normal text after it.", false]]
  },
  "move": [{"idx": 14, "before": 10}],
  "delete": [11]
}
```

   - A plain string keeps the paragraph's existing formatting. A list of `[text, bold]` segments is for paragraphs that mix bold and normal text (`dump` flags these).
   - Bullet paragraphs begin with `•` then a tab. Include that prefix in your replacement text so the bullet and hanging indent survive.
   - Employer lines use a tab before the right-aligned dates. Do not touch those paragraphs.
   - Do not edit section headings.
5. **Apply.** `resume apply "<master>" edits.json "<workspace>/Templates/YYYYMMDD_Company_Role.docx"`. Filename rules: date, company, and role only; underscores instead of spaces; no punctuation; never the word "Resume". It never overwrites the master.
6. **Verify.** `resume check "<tailored docx>" --pages <master page count> --text --outdir "<temp>"`.
   - Page count must equal the master's. If it overflows, shorten or drop the least relevant bullet or tighten wording. If it has a lot of blank space after cutting, that is fine for a one-page resume; do not pad.
   - Read the rendered PDF pages yourself and inspect for clipping, broken bullets, orphan headings, awkward wraps, and crowding.
   - The extracted text must read cleanly in order (this is what applicant tracking systems see).
7. **Truth audit.** Compare the tailored docx to the master, line by line, and confirm: employers, titles, dates, education, certifications, and all numbers are identical; every claim traces to the master or `extra_facts.md`; nothing on the avoid list is implied. Fix anything that fails and re-run `check`.
8. **Report.** Give the path, page count, what emphasis changed and why (three bullets), and the gaps between the posting and the resume. Be direct about gaps.

## When there is no master yet

Invoke `jobagent-setup`.

## Formatting notes

- The build tool produces a plain single-column layout: no tables, text boxes, columns, icons, or images, and Calibri throughout. That is what ATS software parses best. Preserve it.
- LibreOffice renders with Carlito if Calibri is not installed. The metrics are identical, so page counts are reliable.
