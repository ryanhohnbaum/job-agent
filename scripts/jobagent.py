#!/usr/bin/env python3
"""JobAgent command launcher.

  python jobagent.py doctor
  python jobagent.py init --root ~/JobAgent
  python jobagent.py tracker <check-duplicate|create|mark-applied|set-status|queue|log-search> ...
  python jobagent.py resume <build|dump|apply|check> ...
  python jobagent.py capture-posting <url> <out.pdf>
  python jobagent.py gmail-auth | gmail-scan [--days 14] [--dry-run]

The workspace is found via $JOB_AGENT_ROOT or by walking up from the current folder to Config/profile.yaml.
"""
from __future__ import annotations

import importlib
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from common import config_dir, jobagent_root, workbook_lock_path, workbook_path  # noqa: E402

ROUTES = {
    "init": ("init_workspace", "main"),
    "tracker": ("tracker", "main"),
    "resume": ("resume", "main"),
    "capture-posting": ("capture_posting", "main"),
    "gmail-auth": ("gmail_auth", "main"),
    "gmail-scan": ("gmail_archive", "main"),
}


def doctor() -> int:
    root = jobagent_root()
    ok = True

    def line(status: str, label: str, detail: str = "") -> None:
        nonlocal ok
        if status == "FAIL":
            ok = False
        print(f"[{status:<4}] {label}" + (f": {detail}" if detail else ""))

    print(f"Workspace: {root}")
    line("OK" if root.exists() else "FAIL", "workspace folder exists")
    line("OK" if workbook_path().exists() else "FAIL", "tracker workbook", str(workbook_path()))
    line("WARN" if workbook_lock_path().exists() else "OK", "tracker not open in Excel")
    profile = config_dir() / "profile.yaml"
    rules = config_dir() / "search_rules.yaml"
    for path in (profile, rules):
        if not path.exists():
            line("FAIL", f"config {path.name}", "missing")
            continue
        text = path.read_text(encoding="utf-8")
        line("WARN" if "REPLACE_ME" in text else "OK", f"config {path.name}",
             "still has REPLACE_ME placeholders (run setup)" if "REPLACE_ME" in text else "")
    masters = list((root / "Templates").glob("*.docx")) if (root / "Templates").exists() else []
    line("OK" if masters else "WARN", "master resume in Templates/", masters[0].name if masters else "none yet")

    for mod, pip_name, needed in (
        ("openpyxl", "openpyxl", True), ("yaml", "PyYAML", True), ("docx", "python-docx", True),
        ("pypdf", "pypdf", True), ("playwright", "playwright", False),
        ("googleapiclient", "google-api-python-client", False),
    ):
        try:
            importlib.import_module(mod)
            line("OK", f"python package {pip_name}")
        except ImportError:
            line("FAIL" if needed else "INFO", f"python package {pip_name}",
                 "missing" if needed else "missing (only needed for posting capture / Gmail)")

    from resume import find_soffice
    line("OK" if find_soffice() else "WARN", "LibreOffice (resume page-count check)",
         "" if find_soffice() else "not found; install from libreoffice.org")
    print("\nAll required checks passed." if ok else "\nFix the FAIL items above, then rerun doctor.")
    return 0 if ok else 1


def main() -> None:
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help", "help"):
        print(__doc__)
        return
    cmd, rest = sys.argv[1], sys.argv[2:]
    if cmd == "doctor":
        sys.exit(doctor())
    if cmd not in ROUTES:
        raise SystemExit(f"Unknown command {cmd!r}. Run with --help.")
    module_name, func = ROUTES[cmd]
    module = importlib.import_module(module_name)
    try:
        getattr(module, func)(rest)
    except (RuntimeError, FileNotFoundError, ValueError, KeyError, IndexError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
