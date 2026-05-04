"""
take_snapshot.py
Takes a headless screenshot of the Africa dashboard MTD cards and saves it to
mtd_snapshot.png.  Called by publish.bat after each git push (GitHub Pages
deployment is already underway by then; we retry until the page is fresh).

Usage:
    python scripts/take_snapshot.py             # immediate attempt (no deploy wait)
    python scripts/take_snapshot.py --wait 90   # wait N seconds first (post-push)
"""

import argparse, datetime, os, sys, time
from playwright.sync_api import sync_playwright

BASE          = os.path.dirname(os.path.abspath(__file__))
SNAPSHOT_PATH = os.path.join(BASE, "mtd_snapshot.png")
LOG_PATH      = os.path.join(BASE, "snapshot.log")
DASHBOARD_URL = "https://benoit-aujan.github.io/africa-dashboard/"

MAX_ATTEMPTS   = 6
RETRY_WAIT_SEC = 30


def log(msg):
    ts   = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def take_snapshot():
    today  = datetime.date.today()
    target = today.strftime("%B %Y")   # e.g. "May 2026"

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page    = browser.new_page(viewport={"width": 1280, "height": 900})

                # Use "load" (not "networkidle") — avoids hanging on analytics/CDN
                page.goto(DASHBOARD_URL, wait_until="load", timeout=30_000)

                # Wait explicitly for the MTD section to exist (confirms JS ran + data loaded)
                page.wait_for_selector(".mtd-section", timeout=15_000)
                page.wait_for_timeout(2000)   # allow chart rendering

                # Set correct slicers
                page.locator("#btn-weekly").click()
                page.locator("#btn-var-abs").click()
                page.wait_for_timeout(800)

                # Navigate to current month
                for _ in range(24):
                    label = page.locator("#period-label").inner_text()
                    if label == target:
                        break
                    page.locator("button:has-text('›')").click()
                    page.wait_for_timeout(200)

                page.wait_for_timeout(500)
                page.locator(".mtd-section").screenshot(path=SNAPSHOT_PATH)
                browser.close()

            sz = os.path.getsize(SNAPSHOT_PATH)
            log(f"Snapshot saved: {sz:,} bytes  (attempt {attempt}/{MAX_ATTEMPTS})")
            return True

        except Exception as e:
            log(f"Attempt {attempt}/{MAX_ATTEMPTS} failed: {e}")
            if attempt < MAX_ATTEMPTS:
                log(f"Retrying in {RETRY_WAIT_SEC}s...")
                time.sleep(RETRY_WAIT_SEC)

    log("All snapshot attempts failed — mtd_snapshot.png NOT updated.")
    return False


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--wait", type=int, default=0,
                    help="Seconds to wait before first attempt (use after git push)")
    args = ap.parse_args()

    if args.wait > 0:
        log(f"Waiting {args.wait}s for GitHub Pages deployment...")
        time.sleep(args.wait)

    ok = take_snapshot()
    sys.exit(0 if ok else 1)
