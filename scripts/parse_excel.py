"""
Parse all Africa weekly Excel files into a unified JSON dataset.
Handles 3 format variants:
  - 42-col (Jun-Jul 2025): No LE column, 5 metrics per property
  - 61-col (Aug 2025 - Jan 2026): 7 metrics per property, 3 trailing empty cols
  - 58-col (Feb 2026+): 7 metrics per property, "Percentage" sheet or Sheet1
"""

import pandas as pd
import json
import glob
import os
from datetime import datetime, date

EXCEL_GLOB = "C:/Users/benoit.haas/AppData/Local/Temp/Africa *.xlsx"
OUTPUT_JSON = os.path.join(os.path.dirname(__file__), "../data/data.json")

PROPERTIES = ["ABAZ", "Pemba", "ASLV", "Radisson"]
PROPERTY_NAMES = {
    "ABAZ": "Anantara Bazaruto",
    "Pemba": "Avani Pemba",
    "ASLV": "Anantara Stanley & Livingstone",
    "Radisson": "Radisson Blu Maputo",
}

# Column layouts per format
FORMATS = {
    "42col": {
        # No LE; 5 cols per property
        "date_col": 21,
        "occ": {
            "ABAZ":     {"actual": 1,  "bgt": 2,  "py": 3,  "vs_bgt": 4,  "vs_py": 5},
            "Pemba":    {"actual": 6,  "bgt": 7,  "py": 8,  "vs_bgt": 9,  "vs_py": 10},
            "ASLV":     {"actual": 11, "bgt": 12, "py": 13, "vs_bgt": 14, "vs_py": 15},
            "Radisson": {"actual": 16, "bgt": 17, "py": 18, "vs_bgt": 19, "vs_py": 20},
        },
        "adr": {
            "ABAZ":     {"actual": 22, "bgt": 23, "py": 24, "vs_bgt": 25, "vs_py": 26},
            "Pemba":    {"actual": 27, "bgt": 28, "py": 29, "vs_bgt": 30, "vs_py": 31},
            "ASLV":     {"actual": 32, "bgt": 33, "py": 34, "vs_bgt": 35, "vs_py": 36},
            "Radisson": {"actual": 37, "bgt": 38, "py": 39, "vs_bgt": 40, "vs_py": 41},
        },
    },
    "7col": {
        # With LE; 7 cols per property
        "date_col": 29,
        "occ": {
            "ABAZ":     {"actual": 1,  "le": 2,  "bgt": 3,  "py": 4,  "vs_le": 5,  "vs_bgt": 6,  "vs_py": 7},
            "Pemba":    {"actual": 8,  "le": 9,  "bgt": 10, "py": 11, "vs_le": 12, "vs_bgt": 13, "vs_py": 14},
            "ASLV":     {"actual": 15, "le": 16, "bgt": 17, "py": 18, "vs_le": 19, "vs_bgt": 20, "vs_py": 21},
            "Radisson": {"actual": 22, "le": 23, "bgt": 24, "py": 25, "vs_le": 26, "vs_bgt": 27, "vs_py": 28},
        },
        "adr": {
            "ABAZ":     {"actual": 30, "le": 31, "bgt": 32, "py": 33, "vs_le": 34, "vs_bgt": 35, "vs_py": 36},
            "Pemba":    {"actual": 37, "le": 38, "bgt": 39, "py": 40, "vs_le": 41, "vs_bgt": 42, "vs_py": 43},
            "ASLV":     {"actual": 44, "le": 45, "bgt": 46, "py": 47, "vs_le": 48, "vs_bgt": 49, "vs_py": 50},
            "Radisson": {"actual": 51, "le": 52, "bgt": 53, "py": 54, "vs_le": 55, "vs_bgt": 56, "vs_py": 57},
        },
    },
}


def detect_format(df):
    if df.shape[1] <= 43:
        return "42col"
    return "7col"


def safe_float(val):
    try:
        f = float(val)
        return None if pd.isna(f) else round(f, 6)
    except (TypeError, ValueError):
        return None


def extract_row(df, row_idx, fmt_key, file_date):
    fmt = FORMATS[fmt_key]
    date_val = df.iloc[row_idx, fmt["date_col"]]

    is_mtd = False
    row_date = None

    if pd.isna(date_val):
        return None, None

    date_str = str(date_val).strip()
    if "MTD" in date_str.upper():
        is_mtd = True
        row_date = file_date.strftime("%Y-%m") + "-MTD"
    else:
        try:
            from dateutil.relativedelta import relativedelta
            if hasattr(date_val, "date"):
                parsed = date_val.date()
            else:
                parsed = pd.to_datetime(date_str).date()
            # Correct year-rollover errors: if row date is >60 days ahead of
            # the file date and the gap is ~365 days, the year is off by +1
            delta = (parsed - file_date.date()).days
            if delta > 60 and abs(delta - 365) < 30:
                parsed = parsed - relativedelta(years=1)
            row_date = parsed.isoformat()
        except Exception:
            return None, None

    record = {}
    for prop in PROPERTIES:
        occ_cols = fmt["occ"][prop]
        adr_cols = fmt["adr"][prop]

        occ_actual = safe_float(df.iloc[row_idx, occ_cols["actual"]])
        adr_actual = safe_float(df.iloc[row_idx, adr_cols["actual"]])

        occ_le  = safe_float(df.iloc[row_idx, occ_cols.get("le",  -1)]) if "le"  in occ_cols else None
        occ_bgt = safe_float(df.iloc[row_idx, occ_cols.get("bgt", -1)]) if "bgt" in occ_cols else None
        occ_py  = safe_float(df.iloc[row_idx, occ_cols.get("py",  -1)]) if "py"  in occ_cols else None

        adr_le  = safe_float(df.iloc[row_idx, adr_cols.get("le",  -1)]) if "le"  in adr_cols else None
        adr_bgt = safe_float(df.iloc[row_idx, adr_cols.get("bgt", -1)]) if "bgt" in adr_cols else None
        adr_py  = safe_float(df.iloc[row_idx, adr_cols.get("py",  -1)]) if "py"  in adr_cols else None

        # Relative variance (%) — from file columns
        occ_vs_le  = safe_float(df.iloc[row_idx, occ_cols.get("vs_le",  -1)]) if "vs_le"  in occ_cols else None
        occ_vs_bgt = safe_float(df.iloc[row_idx, occ_cols.get("vs_bgt", -1)]) if "vs_bgt" in occ_cols else None
        occ_vs_py  = safe_float(df.iloc[row_idx, occ_cols.get("vs_py",  -1)]) if "vs_py"  in occ_cols else None

        adr_vs_le  = safe_float(df.iloc[row_idx, adr_cols.get("vs_le",  -1)]) if "vs_le"  in adr_cols else None
        adr_vs_bgt = safe_float(df.iloc[row_idx, adr_cols.get("vs_bgt", -1)]) if "vs_bgt" in adr_cols else None
        adr_vs_py  = safe_float(df.iloc[row_idx, adr_cols.get("vs_py",  -1)]) if "vs_py"  in adr_cols else None

        # Absolute variance (PP for occ, $ for ADR) — calculated from actuals
        def pp_diff(actual, benchmark):
            if actual is not None and benchmark is not None:
                return round((actual - benchmark) * 100, 4)
            return None

        def abs_diff(actual, benchmark):
            if actual is not None and benchmark is not None:
                return round(actual - benchmark, 2)
            return None

        record[prop] = {
            "occ":        occ_actual,
            "occ_le":     occ_le,
            "occ_bgt":    occ_bgt,
            "occ_py":     occ_py,
            "occ_vs_le":  occ_vs_le,
            "occ_vs_bgt": occ_vs_bgt,
            "occ_vs_py":  occ_vs_py,
            # Absolute (PP)
            "occ_vs_le_pp":  pp_diff(occ_actual, occ_le),
            "occ_vs_bgt_pp": pp_diff(occ_actual, occ_bgt),
            "occ_vs_py_pp":  pp_diff(occ_actual, occ_py),
            "adr":        adr_actual,
            "adr_le":     adr_le,
            "adr_bgt":    adr_bgt,
            "adr_py":     adr_py,
            "adr_vs_le":  adr_vs_le,
            "adr_vs_bgt": adr_vs_bgt,
            "adr_vs_py":  adr_vs_py,
            # Absolute ($)
            "adr_vs_le_abs":  abs_diff(adr_actual, adr_le),
            "adr_vs_bgt_abs": abs_diff(adr_actual, adr_bgt),
            "adr_vs_py_abs":  abs_diff(adr_actual, adr_py),
        }

    return row_date, record


def get_file_date(filename):
    """Parse the date from the filename like 'Africa 06042026.xlsx' -> 2026-04-06."""
    base = os.path.splitext(os.path.basename(filename))[0]
    parts = base.split(" ")
    date_part = parts[-1] if len(parts) > 1 else parts[0]
    try:
        if len(date_part) == 6:  # DDMMYY
            return datetime.strptime(date_part, "%d%m%y")
        elif len(date_part) == 8:  # DDMMYYYY
            return datetime.strptime(date_part, "%d%m%Y")
    except ValueError:
        pass
    return datetime.today()


def parse_file(filepath):
    """Return dict of {date_key: {prop: data}} and list of MTD rows."""
    file_date = get_file_date(filepath)
    xl = pd.ExcelFile(filepath)

    # Pick the right sheet
    if "Percentage" in xl.sheet_names:
        sheet = "Percentage"
    else:
        sheet = xl.sheet_names[0]

    df = pd.read_excel(filepath, sheet_name=sheet, header=None)
    fmt_key = detect_format(df)

    daily = {}
    mtd = {}

    for row_idx in range(3, len(df)):
        date_key, record = extract_row(df, row_idx, fmt_key, file_date)
        if date_key is None or record is None:
            continue
        if "MTD" in date_key:
            mtd[date_key] = record
        else:
            daily[date_key] = record

    return daily, mtd, file_date


def merge_data():
    """
    Merge all Excel files, latest file wins on conflict.
    Returns (data dict, discrepancies list).
    """
    files = sorted(glob.glob(EXCEL_GLOB), key=lambda f: get_file_date(f))

    all_daily = {}   # date -> {prop: data, "_sources": [...]}
    all_mtd = {}
    discrepancies = []

    for filepath in files:
        try:
            daily, mtd, file_date = parse_file(filepath)
        except Exception as e:
            print(f"  WARNING: Could not parse {os.path.basename(filepath)}: {e}")
            continue

        fname = os.path.basename(filepath)

        for date_key, record in daily.items():
            if date_key in all_daily:
                existing = all_daily[date_key]
                # Check for discrepancies on key metrics
                for prop in PROPERTIES:
                    if prop not in existing or prop not in record:
                        continue
                    for metric in ["occ", "adr"]:
                        old_val = existing[prop].get(metric)
                        new_val = record[prop].get(metric)
                        if old_val is not None and new_val is not None:
                            if abs(old_val - new_val) > 0.001:
                                discrepancies.append({
                                    "date": date_key,
                                    "property": prop,
                                    "metric": metric,
                                    "old_value": old_val,
                                    "new_value": new_val,
                                    "old_source": existing.get("_source", "unknown"),
                                    "new_source": fname,
                                })
                # Latest file wins
                record["_source"] = fname
                all_daily[date_key] = record
            else:
                record["_source"] = fname
                all_daily[date_key] = record

        for date_key, record in mtd.items():
            record["_source"] = fname
            all_mtd[date_key] = record

    # Clean up _source from output data
    clean_daily = {}
    for k, v in all_daily.items():
        row = {p: v[p] for p in PROPERTIES if p in v}
        clean_daily[k] = row

    clean_mtd = {}
    for k, v in all_mtd.items():
        row = {p: v[p] for p in PROPERTIES if p in v}
        clean_mtd[k] = row

    return clean_daily, clean_mtd, discrepancies


def main():
    print("Parsing all Africa Excel files...")
    daily, mtd, discrepancies = merge_data()

    output = {
        "generated_at": datetime.now().isoformat(),
        "properties": PROPERTIES,
        "property_names": PROPERTY_NAMES,
        "daily": daily,
        "mtd": mtd,
    }

    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)
    with open(OUTPUT_JSON, "w") as f:
        json.dump(output, f, indent=2, default=str)

    print(f"Wrote {len(daily)} daily records and {len(mtd)} MTD records to {OUTPUT_JSON}")

    if discrepancies:
        disc_path = os.path.join(os.path.dirname(OUTPUT_JSON), "discrepancies.json")
        with open(disc_path, "w") as f:
            json.dump(discrepancies, f, indent=2, default=str)
        print(f"\n*** DISCREPANCIES FOUND: {len(discrepancies)} ***")
        print(f"    Saved to {disc_path}")
        print("    Draft email saved to discrepancy_draft_email.txt")
        write_discrepancy_email(discrepancies)
    else:
        print("No discrepancies found.")

    return output


def write_discrepancy_email(discrepancies):
    """Generate a draft email flagging discrepancies for Benoit to review."""
    # Group by property
    by_prop = {}
    for d in discrepancies:
        key = d["property"]
        by_prop.setdefault(key, []).append(d)

    lines = []
    lines.append("TO: [Reply to original sender — do not copy all]")
    lines.append("SUBJECT: RE: [Original subject] — Data Discrepancy Flagged")
    lines.append("")
    lines.append("Hi [Name],")
    lines.append("")
    lines.append("Thank you for the report. While consolidating the daily data, I noticed some discrepancies between figures reported at different dates. Could you please help clarify:")
    lines.append("")

    for prop, items in by_prop.items():
        prop_name = PROPERTY_NAMES.get(prop, prop)
        lines.append(f"• {prop_name} ({prop}):")
        for d in items[:5]:  # Cap at 5 per property
            metric_label = "Occupancy" if d["metric"] == "occ" else "ADR"
            old_fmt = f"{d['old_value']*100:.1f}%" if d["metric"] == "occ" else f"${d['old_value']:.2f}"
            new_fmt = f"{d['new_value']*100:.1f}%" if d["metric"] == "occ" else f"${d['new_value']:.2f}"
            lines.append(f"  - {d['date']} {metric_label}: previously reported as {old_fmt}, now showing {new_fmt}")
        lines.append("")

    lines.append("Could you confirm which figure is correct? If there was a correction, please let me know the reason so I can update accordingly.")
    lines.append("")
    lines.append("Thank you,")
    lines.append("Benoit")

    email_path = os.path.join(os.path.dirname(OUTPUT_JSON), "discrepancy_draft_email.txt")
    with open(email_path, "w") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    main()
