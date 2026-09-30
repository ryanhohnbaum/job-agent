---
name: jobagent-track
description: Manage the JobAgent application tracker. Show the queue of roles ready to apply to, check a role for duplicates, mark an application Applied after the user confirms submission, update statuses (interview, rejected, withdrawn, skipped, on hold), and log searches. Use for "show my queue", "what should I apply to", "I applied to X", "mark X as applied", "I got an interview", "I was rejected", or "skip that one".
---

# JobAgent tracker

`Job_Search_Log.xlsx` in the workspace is the single record of every application. Always use the commands below to change it. Do not write ad hoc scripts against the workbook. The commands back up the file, check that Excel does not have it open, write by column name, verify the saved result, and only then replace the file.

Launcher (`python3` on macOS/Linux), run from the workspace folder:

```
python "${CLAUDE_PLUGIN_ROOT}/scripts/jobagent.py" tracker <command>
```

## Commands

| Need | Command |
|---|---|
| What is ready to apply to | `queue` (Qualified and On Hold). `queue --all` for every status. `--json` for structured output. |
| Is this role new | `check-duplicate --company C --role R [--reqid ID]` |
| Add a role | `create ...` (normally done by `jobagent-search`) |
| User submitted an application | `mark-applied --application-id APP-... --confirmed-submitted [--date YYYY-MM-DD]` |
| Any other status change | `set-status --application-id APP-... --status <Status> [--note "..."]` |
| Record a search run | `log-search --source "..." --reviewed N --qualified N --started N` |

Statuses: Qualified (ready for the user to apply), On Hold, Applied, Interview, Offer, Rejected, Withdrawn, Skipped, Closed, Do Not Reapply.

## The Applied rule

Mark an application Applied **only after the user tells you they successfully submitted it on the employer's site.** Never infer it, never mark it because a resume was prepared, and never mark it on a schedule. If the user's statement is ambiguous ("I think I applied"), ask. Marking is idempotent: repeating it on an already-Applied row changes nothing.

`mark-applied` refuses to run without `--confirmed-submitted`. Passing that flag asserts that the user confirmed submission in this conversation.

## Showing the queue

Run `queue`, then present it as a short list sorted by fit score: company, title, fit, location and work model, the apply link, and the folder path. Flag any role whose posting is more than about two weeks old as "may have closed; check the link first". Offer to skip or hold anything the user is not interested in.

## Identifying the right row

Match by Application ID when the user has it; otherwise by requisition ID or company plus role. If more than one row could match, list the candidates and ask. Never guess.

## After a status change

Report exactly what changed: which row, old and new status, and the backup file created. If a command fails (Excel open, duplicate, missing column), report the error verbatim and stop; do not work around it by editing the file another way.

## Manual application reminder

The user does the employer-side work: opening the link, uploading the resume, answering questions, handling accounts and attestations, and pressing Submit. You never do those steps.
