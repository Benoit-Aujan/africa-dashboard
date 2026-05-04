"""
backfill_mgmt_packs.py
One-off script to populate monthly_benchmarks from the March 2026 management pack
files already saved in Temp. Bypasses Outlook scan and state-file gate.

Run once:
    python scripts/backfill_mgmt_packs.py
    python scripts/backfill_mgmt_packs.py --dry-run
"""

import argparse, datetime, json, os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_raw_actuals import parse_mgmt_pack_minor, parse_mgmt_pack_radisson

BASE      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_JSON = os.path.join(BASE, "data", "data.json")
EXTRACT   = "C:/Users/benoit.haas/AppData/Local/Temp/mgmt_pack_extract"

# The four March 2026 management pack files we already have
PACK_FILES = [
    {
        "path":       os.path.join(EXTRACT, "ABAZ_Hotel Finance Template 2026 V16.1 with link_March_31032026_09042026.xlsb"),
        "prop":       "ABAZ",
        "rpt_month":  datetime.date(2026, 3, 1),
        "is_mzn":     True,
        "parser":     "minor",
    },
    {
        "path":       os.path.join(EXTRACT, "ASLV Hotel Finance Template Mar 2026 Final NL.xlsb"),
        "prop":       "ASLV",
        "rpt_month":  datetime.date(2026, 3, 1),
        "is_mzn":     False,
        "parser":     "minor",
    },
    {
        "path":       os.path.join(EXTRACT, "VPEM Hotel Finance Template March 2026 SP NL Updated Commentary.xlsb"),
        "prop":       "Pemba",
        "rpt_month":  datetime.date(2026, 3, 1),
        "is_mzn":     True,
        "parser":     "minor",
    },
    {
        "path":       os.path.join(EXTRACT, "7.MPMZH USAH 12MONTHS REPORT MARCH 2026.xlsx"),
        "prop":       "Radisson",
        "rpt_month":  datetime.date(2026, 3, 1),
        "is_mzn":     None,
        "parser":     "radisson",
    },
]


def apply_benchmarks(data, prop, rpt_month, parsed, dry_run):
    """
    Merge parsed LE/Budget into data["monthly_benchmarks"].
    Budget: immutable once stored.
    LE: always updated (no preliminary-vs-final alert for backfill).
    Returns (n_le_written, n_bgt_written).
    """
    if "monthly_benchmarks" not in data:
        data["monthly_benchmarks"] = {}
    prop_mb = data["monthly_benchmarks"].setdefault(prop, {})
    n_le = n_bgt = 0

    for section in ("LE", "Budget"):
        for month_str, vals in parsed.get(section, {}).items():
            month_entry   = prop_mb.setdefault(month_str, {})
            section_entry = month_entry.setdefault(section, {})

            for key in ("occ", "adr"):
                new_val = vals.get(key)
                if new_val is None:
                    continue
                existing = section_entry.get(key)
                if section == "Budget" and existing is not None:
                    # Budget immutable — keep old
                    continue
                if not dry_run:
                    section_entry[key] = round(new_val, 6)
                if section == "LE":
                    n_le += 1
                else:
                    n_bgt += 1

    return n_le, n_bgt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if args.dry_run:
        print("[DRY RUN] No files will be modified.\n")

    with open(DATA_JSON, encoding="utf-8") as f:
        data = json.load(f)

    total_le = total_bgt = 0

    for cfg in PACK_FILES:
        path = cfg["path"]
        prop = cfg["prop"]
        rpt  = cfg["rpt_month"]

        if not os.path.exists(path):
            print(f"  MISSING: {os.path.basename(path)} — skipped")
            continue

        print(f"\nParsing {prop} ({rpt.strftime('%B %Y')}): {os.path.basename(path)}")

        if cfg["parser"] == "minor":
            parsed = parse_mgmt_pack_minor(path, rpt, is_mzn=cfg["is_mzn"])
        else:
            parsed = parse_mgmt_pack_radisson(path, rpt)

        if not parsed or (not parsed.get("LE") and not parsed.get("Budget")):
            print(f"  WARNING: No data extracted.")
            continue

        le_months  = sorted(parsed.get("LE",     {}).keys())
        bgt_months = sorted(parsed.get("Budget", {}).keys())
        print(f"  LE months ({len(le_months)}):  {', '.join(le_months)}")
        print(f"  Bgt months ({len(bgt_months)}): {', '.join(bgt_months)}")

        # Show a sample value
        if le_months:
            sample_m = le_months[0]
            sv = parsed["LE"][sample_m]
            print(f"  Sample LE {sample_m}: Occ={sv['occ']*100:.1f}%  ADR=${sv['adr']:.2f}")

        n_le, n_bgt = apply_benchmarks(data, prop, rpt, parsed, args.dry_run)
        total_le  += n_le
        total_bgt += n_bgt
        print(f"  Written: {n_le//2} LE months, {n_bgt//2} Budget months")

    if not args.dry_run:
        with open(DATA_JSON, "w", encoding="utf-8") as f:
            json.dump(data, f, separators=(",", ":"))
        print(f"\nSaved data.json  (LE fields: {total_le}, Budget fields: {total_bgt})")

        # Run pipeline to propagate to daily records
        print("\nRunning build_raw_actuals.py...")
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_raw_actuals.py")
        result = subprocess.run([sys.executable, script], capture_output=True, text=True)
        for line in result.stdout.strip().splitlines():
            print(f"  {line}")
        if result.returncode != 0:
            print(f"  WARNING: pipeline error:\n{result.stderr}")
    else:
        print(f"\n[DRY RUN] Would write {total_le} LE fields and {total_bgt} Budget fields.")


if __name__ == "__main__":
    main()
