"""Authorize read-only Gmail access (one time). Needs Config/credentials.json, a Google Desktop-app OAuth client."""
from __future__ import annotations

import argparse
import os

from common import config_dir

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(description="Authorize JobAgent read-only Gmail access.")
    parser.add_argument("--force", action="store_true", help="Replace an existing token.")
    args = parser.parse_args(argv)

    from google_auth_oauthlib.flow import InstalledAppFlow

    config = config_dir()
    credentials = config / "credentials.json"
    token = config / "gmail_token.json"
    if not credentials.exists():
        raise FileNotFoundError(
            f"Missing {credentials}. Follow skills/jobagent-gmail/SKILL.md to create a Google OAuth client first."
        )
    if token.exists() and not args.force:
        raise RuntimeError(f"{token} already exists. Rerun with --force only to reauthorize on purpose.")

    flow = InstalledAppFlow.from_client_secrets_file(str(credentials), SCOPES)
    creds = flow.run_local_server(port=0)
    temp = token.with_suffix(".json.tmp")
    temp.write_text(creds.to_json(), encoding="utf-8")
    os.replace(temp, token)
    print(f"Gmail read-only authorization saved to {token}")


if __name__ == "__main__":
    main()
