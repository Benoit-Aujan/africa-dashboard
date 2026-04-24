"""
update_from_email.py
Scans Outlook inbox for Minor tracker files and Radisson Revenue Reports.
Saves attachments to Temp (no overwriting), runs the full data pipeline,
and drafts an Outlook email for any discrepancies vs confirmed figures.

Sources handled:
  Consolidated Minor (ABAZ + ASLV + Pemba):
    bvangent@minor.com       Bronwyn Van Gent (primary)
    awolfaardt@minor.com     Annalise Wolfaardt (backup)
    chartley@minor.com       Chantelle Hartley (backup)

  Individual ABAZ:
    fo.bazaruto@anantara.com   Anantara Bazaruto Front Office
    bazaruto@anantara.com      Anantara Bazaruto (alt address)

  Individual Pemba:
    hpasipamire@minorhotels.com  Hamilton Pasipamire (primary)
    bjubane@nhhotels.com         Busani Jubane (backup)

  Radisson Excel (Revenue Report):
    natalia.sitoe@radissonblu.com    Natalia Sitoe (primary)
    shirley.lapoule@radissonblu.com  Shirley Lapoule (backup)

  Radisson PDF (Daily Manager Flash Report):
    reception.maputo@radissonblu.com
    → PDF only: saved to Temp for manual review (not parsed automatically)

Usage:
    python scripts/update_from_email.py            # scan last 3 days
    python scripts/update_from_email.py --days 7   # scan last 7 days
    python scripts/update_from_email.py --dry-run  # preview, no changes
"""

import argparse, datetime, json, os, re, subprocess, sys
import win32com.client

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_raw_actuals import parse_tracker, parse_radisson, parse_abaz_daily, parse_vpem_daily, parse_radisson_pdf

BASE      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_JSON = os.path.join(BASE, "data", "data.json")
TEMP      = "C:/Users/benoit.haas/AppData/Local/Temp"

# ── Sender configuration ──────────────────────────────────────────────────────

# These send the consolidated VPEM/ABAZ/ASLV tracker Excel
CONSOLIDATED_SENDERS = {
    "bvangent@minor.com",        # Bronwyn Van Gent (primary)
    "awolfaardt@minor.com",      # Annalise Wolfaardt (backup)
    "chartley@minor.com",        # Chantelle Hartley (backup)
}

# These send individual ABAZ reports (Excel, tracker format)
ABAZ_SENDERS = {
    "fo.bazaruto@anantara.com",  # Anantara Bazaruto Front Office (primary)
    "bazaruto@anantara.com",     # alt address
}

# These send individual Pemba/VPEM reports (Excel, tracker format)
PEMBA_SENDERS = {
    "hpasipamire@minorhotels.com",  # Hamilton Pasipamire (primary)
    "bjubane@nhhotels.com",          # Busani Jubane (backup)
}

# These send Radisson Revenue Report Excel files
RADISSON_XLS_SENDERS = {
    "natalia.sitoe@radissonblu.com",   # Natalia Sitoe (primary)
    "shirley.lapoule@radissonblu.com", # Shirley Lapoule (backup)
}

# This address sends a PDF flash report — saved for manual review, not parsed
RADISSON_PDF_SENDER = "reception.maputo@radissonblu.com"

ALL_SENDERS = (
    CONSOLIDATED_SENDERS | ABAZ_SENDERS | PEMBA_SENDERS |
    RADISSON_XLS_SENDERS | {RADISSON_PDF_SENDER}
)

EXCEL_RE = re.compile(r"\.xlsx?b?$", re.IGNORECASE)   # .xls, .xlsx, .xlsb
PDF_RE   = re.compile(r"\.pdf$",    re.IGNORECASE)

# Any difference at all is a discrepancy — zero tolerance
TOLERANCE_ROOMS = 0
TOLERANCE_REV   = 0.0


# ── Outlook helpers ───────────────────────────────────────────────────────────

def connect_outlook():
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        print(f"ERROR: Cannot connect to Outlook: {e}")
        print("Make sure Classic Outlook is open.")
        sys.exit(1)


def save_attachment(att):
    """Save to Temp with a date prefix. Never overwrites an existing file."""
    ts = datetime.date.today().strftime("%Y-%m-%d")
    base, ext = os.path.splitext(att.FileName)
    path = os.path.join(TEMP, f"{ts}_{base}{ext}")
    counter = 1
    while os.path.exists(path):
        path = os.path.join(TEMP, f"{ts}_{base}_{counter}{ext}")
        counter += 1
    att.SaveAsFile(path)
    return path


def scan_inbox(outlook, days):
    """
    Scans inbox for relevant attachments from known senders.
    Returns list of (sender_email, subject, att_type, saved_path) where
    att_type is one of: 'consolidated', 'abaz', 'pemba', 'radisson_xls', 'radisson_pdf'.
    """
    ns     = outlook.GetNamespace("MAPI")
    inbox  = ns.GetDefaultFolder(6)   # olFolderInbox
    cutoff = datetime.datetime.now() - datetime.timedelta(days=days)
    found  = []

    for item in inbox.Items:
        try:
            if item.Class != 43:   # olMail
                continue
            received = item.ReceivedTime.replace(tzinfo=None)
            if received < cutoff:
                continue
            sender  = (item.SenderEmailAddress or "").lower().strip()
            subject = item.Subject or ""

            if sender not in ALL_SENDERS:
                continue

            for att in item.Attachments:
                fname = att.FileName

                if sender in CONSOLIDATED_SENDERS and EXCEL_RE.search(fname):
                    path = save_attachment(att)
                    found.append((sender, subject, "consolidated", path))
                    print(f"  Saved consolidated tracker : {os.path.basename(path)}")

                elif sender in ABAZ_SENDERS and EXCEL_RE.search(fname):
                    path = save_attachment(att)
                    found.append((sender, subject, "abaz", path))
                    print(f"  Saved ABAZ individual      : {os.path.basename(path)}")

                elif sender in PEMBA_SENDERS and EXCEL_RE.search(fname):
                    path = save_attachment(att)
                    found.append((sender, subject, "pemba", path))
                    print(f"  Saved Pemba individual     : {os.path.basename(path)}")

                elif sender in RADISSON_XLS_SENDERS and EXCEL_RE.search(fname):
                    path = save_attachment(att)
                    found.append((sender, subject, "radisson_xls", path))
                    print(f"  Saved Radisson Excel       : {os.path.basename(path)}")

                elif sender == RADISSON_PDF_SENDER and PDF_RE.search(fname):
                    path = save_attachment(att)
                    found.append((sender, subject, "radisson_pdf", path))
                    print(f"  Saved Radisson PDF         : {os.path.basename(path)}")

        except Exception:
            continue

    return found


# ── Parsing ───────────────────────────────────────────────────────────────────

def parse_attachment(att_type, path):
    """
    Returns {prop: {date_str: raw}} or None if not parseable.
    For individual senders, filters to only the relevant property.
    """
    try:
        if att_type == "consolidated":
            return parse_tracker(path)   # returns all 3: Pemba, ABAZ, ASLV

        elif att_type == "abaz":
            if path.lower().endswith(".xlsb"):
                days = parse_abaz_daily(path)
                return {"ABAZ": days} if days else None
            else:
                parsed = parse_tracker(path)
                return {p: d for p, d in parsed.items() if p == "ABAZ"} if parsed else None

        elif att_type == "pemba":
            if path.lower().endswith(".xlsb"):
                days = parse_vpem_daily(path)
                return {"Pemba": days} if days else None
            else:
                parsed = parse_tracker(path)
                return {p: d for p, d in parsed.items() if p == "Pemba"} if parsed else None

        elif att_type == "radisson_xls":
            days = parse_radisson(path)
            return {"Radisson": days} if days else None

        elif att_type == "radisson_pdf":
            days = parse_radisson_pdf(path)
            return {"Radisson": days} if days else None

    except Exception as e:
        print(f"    ERROR parsing {os.path.basename(path)}: {e}")
        return None


# ── Data helpers ──────────────────────────────────────────────────────────────

def load_data():
    with open(DATA_JSON, encoding="utf-8") as f:
        return json.load(f)


def check_discrepancies(incoming, current_daily):
    """
    Compares incoming data against existing data.json.
    Returns list of discrepancy dicts (only for days that already have confirmed data).
    """
    discs = []
    for prop, days in incoming.items():
        for date_str, raw in days.items():
            existing = current_daily.get(date_str, {}).get(prop, {})
            if existing.get("rooms_occ") is None:
                continue   # new date — not a discrepancy
            delta_rooms = abs((raw["rooms_occ"] or 0) - (existing["rooms_occ"] or 0))
            delta_rev   = abs((raw["rev_usd"]   or 0) - (existing["rev_usd"]   or 0))
            if delta_rooms > TOLERANCE_ROOMS or delta_rev > TOLERANCE_REV:
                discs.append({
                    "prop":  prop,
                    "date":  date_str,
                    "db":    {"rooms_occ": existing["rooms_occ"],
                              "rev_usd":   round(existing["rev_usd"], 2)},
                    "email": {"rooms_occ": raw["rooms_occ"],
                              "rev_usd":   round(raw["rev_usd"], 2)},
                })
    return discs


def run_pipeline(dry_run):
    """Runs build_raw_actuals.py to merge all tracker/Radisson files and refresh PY + benchmarks."""
    if dry_run:
        print("  [DRY RUN] Would run build_raw_actuals.py")
        return
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_raw_actuals.py")
    result = subprocess.run([sys.executable, script], capture_output=True, text=True)
    for line in result.stdout.strip().splitlines():
        print(f"  {line}")
    if result.returncode != 0:
        print(f"  WARNING: pipeline error:\n{result.stderr}")


DISC_TO = [
    "Benoit Haas <benoit.haas@aujan.com>",
    "Glenda Gallego <Glenda.Gallego@aujan.com>",
    "Muhammad Moiz Siddiqui <moiz.siddiqui@aujan.com>",
]


def draft_email(outlook, sender, subject, discs):
    """
    Sends a discrepancy alert to Benoit + Glenda + Moiz.
    The original property sender is noted at the top (not directly emailed —
    Benoit should review and forward if needed).
    """
    mail         = outlook.CreateItem(0)   # olMailItem
    mail.To      = "; ".join(DISC_TO)
    mail.Subject = f"Africa Dashboard — Data discrepancy alert ({subject})"

    rows = "\n".join(
        f"  - {d['prop']} on {d['date']}:\n"
        f"      On record : {d['db']['rooms_occ']} rooms / ${d['db']['rev_usd']:.0f} revenue\n"
        f"      New file  : {d['email']['rooms_occ']} rooms / ${d['email']['rev_usd']:.0f} revenue"
        for d in discs
    )
    mail.Body = (
        f"*** This alert should be forwarded to the property contact: {sender} ***\n"
        f"*** Subject of their original email: {subject} ***\n"
        f"{'─' * 60}\n\n"
        "Dear team,\n\n"
        "While processing the latest property file, the following figures differ "
        "from what was previously recorded. Please review and confirm which are correct "
        "before forwarding to the property:\n\n"
        f"{rows}\n\n"
        "The dashboard has been updated with the latest figures. "
        "Please let me know if any correction is needed.\n\n"
        "Best regards,\nBenoit"
    )
    mail.Send()
    print(f"  Discrepancy alert sent to: {', '.join(DISC_TO)}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--days",    type=int, default=3,
                        help="Scan emails from last N days (default: 3)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview only — no files saved, no data updated")
    args = parser.parse_args()

    print("Connecting to Outlook...")
    outlook = connect_outlook()

    print(f"Scanning inbox (last {args.days} day(s))...")
    if args.dry_run:
        print("[DRY RUN] No changes will be saved.\n")

    attachments = scan_inbox(outlook, days=args.days)

    if not attachments:
        print("No relevant attachments found.")
        return

    if not attachments:
        print("\nNo relevant attachments to parse.")
        return

    # Parse all attachments (Excel + PDF)
    print(f"\nParsing {len(attachments)} attachment(s)...")
    incoming       = {}   # prop -> {date: raw}
    sender_by_prop = {}   # prop -> (sender_email, subject)

    for sender, subject, att_type, path in attachments:
        result = parse_attachment(att_type, path)
        if not result:
            continue
        for prop, days in result.items():
            incoming.setdefault(prop, {}).update(days)
            sender_by_prop[prop] = (sender, subject)
        summary = ", ".join(f"{p}:{len(d)}d" for p, d in result.items())
        print(f"  [{att_type:>14}] {os.path.basename(path)} -> {summary}")

    if not incoming:
        print("No data parsed.")
        return

    # Load full data dict — keep reference so saves are reflected
    data          = load_data()
    current_daily = data.get("daily", {})

    # Count genuinely new records (not yet in data.json)
    new_dates = [
        (prop, date_str)
        for prop, days in incoming.items()
        for date_str in days
        if current_daily.get(date_str, {}).get(prop, {}).get("rooms_occ") is None
    ]

    # Check discrepancies vs confirmed data
    discs = check_discrepancies(incoming, current_daily)

    # Summary
    print(f"\nNew records    : {len(new_dates)}")
    for prop, date_str in sorted(new_dates)[-15:]:
        print(f"  + {prop} {date_str}")
    if len(new_dates) > 15:
        print(f"  ... and {len(new_dates) - 15} more")

    if discs:
        print(f"\nDiscrepancies  : {len(discs)}")
        for d in discs:
            print(f"  ! {d['prop']} {d['date']}: "
                  f"db={d['db']['rooms_occ']}rms/${d['db']['rev_usd']:.0f} "
                  f"-> email={d['email']['rooms_occ']}rms/${d['email']['rev_usd']:.0f}")
    else:
        print("Discrepancies  : none")

    # Insert new records into data.json BEFORE running the pipeline.
    # The pipeline handles tracker/Radisson dates automatically; records from
    # sources it doesn't know (e.g. ABAZ .xlsb individual files) must be saved
    # here so the pipeline doesn't erase them.
    if new_dates and not args.dry_run:
        for prop, days in incoming.items():
            for date_str, raw in days.items():
                if current_daily.get(date_str, {}).get(prop, {}).get("rooms_occ") is not None:
                    continue   # already confirmed — skip
                if date_str not in current_daily:
                    current_daily[date_str] = {}
                prop_rec = current_daily[date_str].setdefault(prop, {
                    "occ_le": None, "occ_bgt": None, "occ_py": None,
                    "adr_le": None, "adr_bgt": None, "adr_py": None,
                })
                ro, ra, ru = raw["rooms_occ"], raw["rooms_avail"], raw["rev_usd"]
                prop_rec.update({
                    "rooms_occ": ro, "rooms_avail": ra, "rev_usd": ru,
                    "occ": round(ro / ra, 6) if ra > 0 else None,
                    "adr": round(ru / ro, 4) if ro > 0 else None,
                })
        data["daily"] = dict(sorted(current_daily.items()))
        with open(DATA_JSON, "w", encoding="utf-8") as f:
            json.dump(data, f, separators=(",", ":"))
        print(f"\nPre-saved {len(new_dates)} new record(s) to data.json")

    # Run full pipeline (updates tracker/Radisson dates, PY, benchmarks)
    print("\nRunning data pipeline...")
    run_pipeline(args.dry_run)

    # Draft discrepancy emails
    if discs and not args.dry_run:
        print("\nDrafting discrepancy emails...")
        from collections import defaultdict
        by_sender = defaultdict(list)
        for d in discs:
            key = sender_by_prop.get(d["prop"], ("unknown@unknown.com", "Data update"))
            by_sender[key].append(d)
        for (sender, subject), sender_discs in by_sender.items():
            draft_email(outlook, sender, subject, sender_discs)

    print("\nDone.")


if __name__ == "__main__":
    main()
