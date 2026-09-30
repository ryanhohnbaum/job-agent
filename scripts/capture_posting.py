"""Save a live job posting as Job_Posting.pdf evidence (real render of the page, not reconstructed text).

Requires: pip install playwright && playwright install chromium
Usage: capture_posting.py <url> <output.pdf>
"""
from __future__ import annotations

import argparse
from pathlib import Path


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(description="Capture a job posting URL to PDF.")
    parser.add_argument("url")
    parser.add_argument("out")
    parser.add_argument("--wait-ms", type=int, default=2500, help="Extra wait for client-rendered pages.")
    args = parser.parse_args(argv)

    from playwright.sync_api import sync_playwright

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1280, "height": 1600})
        page.goto(args.url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(args.wait_ms)
        title = page.title()
        text_len = len(page.inner_text("body").strip())
        if text_len < 400:
            browser.close()
            raise SystemExit(
                f"Page looks empty or blocked ({text_len} characters, title {title!r}). "
                "Not writing evidence. Open the posting manually and print it to PDF instead."
            )
        page.emulate_media(media="screen")
        page.pdf(path=str(out), format="Letter", print_background=True,
                 margin={"top": "0.4in", "bottom": "0.4in", "left": "0.4in", "right": "0.4in"})
        browser.close()
    print(f"WROTE {out} ({out.stat().st_size} bytes) title={title!r}")


if __name__ == "__main__":
    main()
