"""
build_raw_actuals.py
Enriches data.json daily records with raw room counts and revenue from:
  - VPEM_ABAZ_ASLV Pick Up Tracker files (2025 and 2026)
  - Radisson Blu Maputo Revenue Report XLS files (all months)

Adds per-property fields: rooms_occ, rooms_avail, rev_usd
Recomputes occ and adr from those raw figures.
Budget, LE, and PY fields are left untouched.
"""

import json, glob, re, datetime, os
import openpyxl
import xlrd
try:
    import pdfplumber
    _PDF_OK = True
except ImportError:
    _PDF_OK = False
try:
    from pyxlsb import open_workbook as open_xlsb
    _XLSB_OK = True
except ImportError:
    _XLSB_OK = False

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_JSON = os.path.join(BASE, "data", "data.json")
TEMP = "C:/Users/benoit.haas/AppData/Local/Temp"
FX = 65.0  # MZN → USD

# All tracker files in Temp (date-prefixed saves from email + legacy undated copies),
# sorted oldest→newest so newer files win on overlap.
TRACKER_FILES = sorted(
    glob.glob(f"{TEMP}/*VPEM_ABAZ_ASLV*Pick Up Tracker*.xlsx"),
    key=os.path.getmtime,
)

# Exclude the partial March 2026 file; include all others
RADISSON_EXCLUDE = {"Revenue Report 30 March 2026..xls"}

MONTH_MAP = {
    "JAN": 1, "FEB": 2, "MAR": 3, "MARCH": 3,
    "APR": 4, "MAY": 5, "JUN": 6, "JUNE": 6,
    "JUL": 7, "JULY": 7, "AUG": 8, "SEP": 9, "SEPT": 9,
    "OCT": 10, "NOV": 11, "DEC": 12,
}

# prop_key in sheet name → dashboard property key
PROP_NAME = {"VPEM": "Pemba", "ABAZ": "ABAZ", "ASLV": "ASLV"}


# ── VPEM / ABAZ / ASLV trackers ──────────────────────────────────────────────

def parse_tracker(path):
    """
    Returns {prop: {date_str: {rooms_occ, rooms_avail, rev_usd}}}
    VPEM  → Total Property cols: occ=9, occ%=10, rev=12  (MZN)
    ABAZ  → cols: occ=1, occ%=2, rev=4                   (MZN)
    ASLV  → cols: occ=1, occ%=2, rev=4                   (USD)
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    result = {}

    for shname in wb.sheetnames:
        m = re.match(r"^(VPEM|ABAZ|ASLV)\s+(\w+?)(\d{2})$", shname)
        if not m:
            continue
        prop_key, month_str, year_str = m.group(1), m.group(2).upper(), m.group(3)
        month_num = MONTH_MAP.get(month_str)
        if not month_num:
            continue

        year = 2000 + int(year_str)
        prop = PROP_NAME[prop_key]
        is_mzn = (prop_key != "ASLV")

        sh = wb[shname]
        all_rows = list(sh.iter_rows(min_row=3, values_only=True))

        # First pass: infer property capacity from non-zero occupancy rows.
        # Needed so zero-occ days (genuine closures) still contribute to the
        # rooms_avail denominator in weighted Occ calculations.
        occ_col = 9 if prop_key == "VPEM" else 1
        pct_col = 10 if prop_key == "VPEM" else 2
        capacity = 0
        for row in all_rows:
            if not row or row[0] is None: continue
            occ_v = row[occ_col]; pct_v = row[pct_col]
            if occ_v and pct_v and pct_v > 0:
                capacity = max(capacity, round(int(occ_v) / pct_v))

        # Second pass: build daily records
        for row in all_rows:
            if not row or row[0] is None:
                continue
            date = row[0]
            if not isinstance(date, datetime.datetime):
                continue
            # Year-rollover correction (e.g. ABAZ FEB26 with 2025 dates)
            if date.year != year:
                try:
                    date = date.replace(year=year)
                except ValueError:
                    continue
            date_str = date.strftime("%Y-%m-%d")

            if prop_key == "VPEM":
                rooms_occ = row[9]
                occ_pct   = row[10]
                rev       = row[12]
            else:  # ABAZ or ASLV
                rooms_occ = row[1]
                occ_pct   = row[2]
                rev       = row[4]

            if rooms_occ is None or occ_pct is None or rev is None:
                continue

            # Zero-occupancy day: only store for ASLV (boutique lodge with genuine closures).
            # VPEM and ABAZ are city/resort hotels — zero-occ rows are future empty tracker rows.
            if rooms_occ == 0 and rev == 0:
                if prop_key == "ASLV" and capacity > 0:
                    result.setdefault(prop, {})[date_str] = {
                        "rooms_occ": 0, "rooms_avail": capacity, "rev_usd": 0.0,
                    }
                continue

            try:
                rooms_occ   = int(rooms_occ)
                rooms_avail = round(rooms_occ / occ_pct) if occ_pct > 0 else capacity
                rev_usd     = float(rev) / FX if is_mzn else float(rev)
            except (TypeError, ZeroDivisionError):
                continue

            if rooms_avail == 0:
                continue

            result.setdefault(prop, {})[date_str] = {
                "rooms_occ":   rooms_occ,
                "rooms_avail": rooms_avail,
                "rev_usd":     round(rev_usd, 4),
            }

    return result


# ── Radisson Revenue Report XLS ───────────────────────────────────────────────

def _find_rows(sh):
    """Search col B for room-stat labels; return {normalised_label: row_index}."""
    targets = {
        "COMP ROOMS", "MAINT ROOMS", "VACANT ROOMS",
        "ROOMS SOLD", "HOUSE USE ROOMS",
    }
    found = {}
    for r in range(sh.nrows):
        raw = sh.cell_value(r, 1)
        if not isinstance(raw, str):
            continue
        norm = " ".join(raw.strip().upper().split())  # collapse double spaces
        if norm in targets:
            found[norm] = r
    return found


def _num(val, default=0.0):
    """Safely cast an xlrd cell value to float."""
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def parse_radisson(path):
    """Returns {date_str: {rooms_occ, rooms_avail, rev_usd}}"""
    wb  = xlrd.open_workbook(path)
    sh  = wb.sheet_by_name("Daily Input")
    rows = _find_rows(sh)
    REV_ROW = 37  # row 38, 0-indexed

    result = {}
    for c in range(2, sh.ncols):
        date_val = sh.cell_value(0, c)
        if not isinstance(date_val, float) or date_val == 0:
            continue
        try:
            date = datetime.date(*xlrd.xldate_as_tuple(date_val, wb.datemode)[:3])
            date_str = date.strftime("%Y-%m-%d")
        except Exception:
            continue

        def _v(label):
            r = rows.get(label)
            return _num(sh.cell_value(r, c)) if r is not None else 0.0

        comp   = _v("COMP ROOMS")
        maint  = _v("MAINT ROOMS")
        vacant = _v("VACANT ROOMS")
        sold   = _v("ROOMS SOLD")
        house  = _v("HOUSE USE ROOMS")
        rev    = _num(sh.cell_value(REV_ROW, c))

        rooms_occ   = int(round(sold + comp + house))
        rooms_avail = int(round(comp + maint + vacant + sold + house))

        if rooms_avail == 0 or rev == 0:
            continue

        result[date_str] = {
            "rooms_occ":   rooms_occ,
            "rooms_avail": rooms_avail,
            "rev_usd":     round(rev / FX, 4),
        }
    return result


# ── Radisson PDF Flash Report ─────────────────────────────────────────────────

def parse_radisson_pdf(path):
    """
    Parses a Radisson Blu Maputo Daily Manager Flash Report PDF.
    Returns {date_str: {rooms_occ, rooms_avail, rev_usd}} or {}.

    PDF columns (left to right): DAY_2026, MTD_2026, YTD_2026, DAY_2025, MTD_2025, YTD_2025
    We use the DAY_2026 column (first numeric value on each row).

    Date is taken from the "Filter Calendar/Month to Date DD.MM.YY" footer,
    which gives the last date included in the MTD column — the same date as DAY.

    Revenue is in MZN; converted to USD using the FX constant.
    """
    if not _PDF_OK:
        print("  WARNING: pdfplumber not installed — pip install pdfplumber")
        return {}

    text = ""
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception as e:
        print(f"  ERROR reading PDF {os.path.basename(path)}: {e}")
        return {}

    # Date from "Filter Calendar/Month to Date DD.MM.YY"
    m = re.search(r"Filter Calendar/Month to Date\s+(\d{2})\.(\d{2})\.(\d{2})", text)
    if not m:
        print(f"  WARNING: no date found in {os.path.basename(path)}")
        return {}
    day, month, yr2 = int(m.group(1)), int(m.group(2)), int(m.group(3))
    date_str = f"{2000 + yr2}-{month:02d}-{day:02d}"

    def _day_val(label):
        """Return the first numeric token after `label` at line start (DAY_2026 column)."""
        hit = re.search(
            rf"^{re.escape(label)}\s+([\d,]+(?:\.\d+)?)",
            text, re.MULTILINE,
        )
        return float(hit.group(1).replace(",", "")) if hit else None

    total_rooms = _day_val("Total Rooms in Hotel")
    rooms_occ   = _day_val("Rooms Occupied")
    room_rev    = _day_val("Room Revenue")

    if None in (total_rooms, rooms_occ, room_rev):
        print(f"  WARNING: missing fields in {os.path.basename(path)} "
              f"(rooms={rooms_occ}, avail={total_rooms}, rev={room_rev})")
        return {}

    return {
        date_str: {
            "rooms_occ":   int(round(rooms_occ)),
            "rooms_avail": int(round(total_rooms)),
            "rev_usd":     round(room_rev / FX, 4),
        }
    }


# ── ABAZ Daily Income .xlsb ───────────────────────────────────────────────────

def parse_abaz_daily(path):
    """
    Parses an ABAZ-Daily Income-YYYY.MM.DD.xlsb file.
    Returns {date_str: {rooms_occ, rooms_avail, rev_usd}} or {}.

    Sheet "Rooms Drivers":
      - Searches for the "Today" row to get Rooms Inventory (avail) and Rooms Sold.
      - Searches for Complementary and House Use rows (those without a 3-letter
        market-segment code in col 2) to get comp/house room counts.
    Sheet "Daily Income":
      - Finds the "Rooms" revenue row for today's MZN total.
    Currency: always MZN → USD via FX constant.
    """
    if not _XLSB_OK:
        print("  WARNING: pyxlsb not installed — cannot parse .xlsb files. Run: pip install pyxlsb")
        return {}

    # Date from filename: ABAZ-Daily Income-2026.04.21.xlsb (may have _N suffix)
    m = re.search(r"(\d{4})\.(\d{2})\.(\d{2})", path)
    if not m:
        return {}
    date_str = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    try:
        with open_xlsb(path) as wb:
            # ── Rooms Drivers sheet ──
            rooms_sold = comp = house = rooms_avail = None
            with wb.get_sheet("Rooms Drivers") as sh:
                rows = list(sh.rows())

            for row in rows:
                vals = [c.v for c in row]
                if len(vals) < 5:
                    continue
                label = vals[1]
                code  = vals[2]   # market-segment code like "CMP", "HOU", or None

                if label == "Today" and isinstance(vals[3], (int, float)) and vals[3] > 0:
                    # Row: Today | Rooms Inventory | Rooms Sold | % Occ | ...
                    rooms_avail = vals[3]
                    rooms_sold  = vals[4]

                elif label == "Complementary" and code is None:
                    # The Comp & House section row (no 3-letter code), col 3 = today's comp rooms
                    if isinstance(vals[3], (int, float)):
                        comp = vals[3]

                elif label == "House Use" and code is None:
                    if isinstance(vals[3], (int, float)):
                        house = vals[3]

            # ── Daily Income sheet — Rooms revenue today (MZN) ──
            rev_mzn = None
            with wb.get_sheet("Daily Income") as sh:
                for row in sh.rows():
                    vals = [c.v for c in row]
                    if len(vals) > 3 and vals[1] == "Rooms" and isinstance(vals[3], (int, float)):
                        rev_mzn = vals[3]
                        break

    except Exception as e:
        print(f"  ERROR parsing {os.path.basename(path)}: {e}")
        return {}

    if rooms_avail is None or rooms_sold is None or rev_mzn is None:
        print(f"  WARNING: could not extract all fields from {os.path.basename(path)}")
        return {}

    rooms_occ   = int(round(rooms_sold + (comp or 0) + (house or 0)))
    rooms_avail = int(round(rooms_avail))
    rev_usd     = round(rev_mzn / FX, 4)

    if rooms_avail == 0 or rev_usd == 0:
        return {}

    return {date_str: {"rooms_occ": rooms_occ, "rooms_avail": rooms_avail, "rev_usd": rev_usd}}


# ── VPEM (Pemba) Daily Income .xlsb ──────────────────────────────────────────

# Known physical room inventory for VPEM (Avani Pemba) including Residences.
# Used as rooms_avail when reading individual daily xlsb files.
VPEM_CAPACITY = 168

def parse_vpem_daily(path):
    """
    Parses a VPEM-Daily Income-YYYY.MM.DD.xlsb file sent by Hamilton Pasipamire.
    Returns {date_str: {rooms_occ, rooms_avail, rev_usd}} or {}.

    Sheet "Contents":  Year/Month/Day in col 14 (0-indexed) of rows 11/12/13.
    Sheet "Rooms Drivers":
      - Row with label "Total Rooms + Residences (ex. Comp & House)":
            col 3 = today rooms sold (excl comp/house),  col 4 = today rev MZN
      - Row "Complementary"       (code CMP): col 3 = comp rooms today
      - Row "Complimentary Owner" (code COO): col 3 = comp owner rooms today
      - Row "House Use"           (code HOU): col 3 = house use rooms today
    rooms_occ  = sold + comp + comp_owner + house
    rooms_avail = VPEM_CAPACITY (physical inventory)
    rev_usd     = rev_mzn / FX
    """
    if not _XLSB_OK:
        print("  WARNING: pyxlsb not installed — cannot parse VPEM xlsb")
        return {}

    # Date from filename first (more reliable than sheet content)
    m = re.search(r"(\d{4})\.(\d{2})\.(\d{2})", path)
    if m:
        date_str = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    else:
        # Fall back to Contents sheet
        try:
            with open_xlsb(path) as wb:
                with wb.get_sheet("Contents") as sh:
                    rows = list(sh.rows())
                year  = int(rows[11][14].v) if len(rows) > 11 else None
                month = int(rows[12][14].v) if len(rows) > 12 else None
                day   = int(rows[13][14].v) if len(rows) > 13 else None
                if not all([year, month, day]):
                    return {}
                date_str = f"{year}-{month:02d}-{day:02d}"
        except Exception:
            return {}

    try:
        with open_xlsb(path) as wb:
            with wb.get_sheet("Rooms Drivers") as sh:
                rows = list(sh.rows())

        sold = comp = comp_owner = house = rev_mzn = None
        for row in rows:
            vals = [c.v for c in row]
            if len(vals) < 5:
                continue
            label = vals[1]
            if label == "Total Rooms + Residences (ex. Comp & House)":
                if isinstance(vals[3], (int, float)):
                    sold    = vals[3]
                    rev_mzn = vals[4] if isinstance(vals[4], (int, float)) else None
            elif label == "Complementary":
                if isinstance(vals[3], (int, float)):
                    comp = vals[3]
            elif label == "Complimentary Owner":
                if isinstance(vals[3], (int, float)):
                    comp_owner = vals[3]
            elif label == "House Use":
                if isinstance(vals[3], (int, float)):
                    house = vals[3]

        if sold is None or rev_mzn is None:
            return {}

        rooms_occ = int(round(sold + (comp or 0) + (comp_owner or 0) + (house or 0)))
        rev_usd   = round(rev_mzn / FX, 4)

        return {
            date_str: {
                "rooms_occ":   rooms_occ,
                "rooms_avail": VPEM_CAPACITY,
                "rev_usd":     rev_usd,
            }
        }
    except Exception as e:
        print(f"  ERROR parsing VPEM xlsb {os.path.basename(path)}: {e}")
        return {}


# ── Merge into data.json ──────────────────────────────────────────────────────

def main():
    # 1. Load existing data.json
    with open(DATA_JSON, encoding="utf-8") as f:
        data = json.load(f)

    # 2. Build actuals dict from tracker files
    #    actuals[prop][date_str] = {rooms_occ, rooms_avail, rev_usd}
    actuals = {}
    for path in TRACKER_FILES:
        try:
            tracker = parse_tracker(path)
        except Exception as e:
            print(f"  SKIP tracker {os.path.basename(path)}: {e}")
            continue
        for prop, days in tracker.items():
            prop_store = actuals.setdefault(prop, {})
            prop_store.update(days)  # later file (2026) overwrites if overlap
    print(f"Tracker actuals loaded: "
          + ", ".join(f"{p}={len(d)}" for p, d in actuals.items()))

    # 3. Build Radisson actuals
    rad_actuals = {}
    rad_files = sorted(
        f for f in glob.glob(f"{TEMP}/Revenue Report*.xls")
        if os.path.basename(f) not in RADISSON_EXCLUDE
    )
    for path in rad_files:
        name = os.path.basename(path)
        try:
            days = parse_radisson(path)
            rad_actuals.update(days)
            print(f"  Radisson {name}: {len(days)} days")
        except Exception as e:
            print(f"  SKIP {name}: {e}")
    # Also load PDF flash reports for dates not yet covered by XLS files.
    # XLS Revenue Reports are more detailed, so XLS always takes priority.
    rad_pdf_files = sorted(
        glob.glob(f"{TEMP}/*manager_report*.pdf"),
        key=os.path.getmtime,
    )
    pdf_added = 0
    for path in rad_pdf_files:
        name = os.path.basename(path)
        try:
            days = parse_radisson_pdf(path)
            for date_str, raw in days.items():
                if date_str not in rad_actuals:  # XLS wins on overlap
                    rad_actuals[date_str] = raw
                    pdf_added += 1
                    print(f"  Radisson PDF {name}: added {date_str}")
        except Exception as e:
            print(f"  SKIP PDF {name}: {e}")

    actuals["Radisson"] = rad_actuals
    print(f"Radisson total: {len(rad_actuals)} days ({pdf_added} from PDF)")

    # 4. Update existing records AND insert new dates not yet in data.json
    daily = data.get("daily", {})
    # Capture the original start date BEFORE any insertions
    range_start = min(daily.keys()) if daily else "2025-06-01"

    def _apply_raw(prop_rec, raw):
        ro, ra, ru = raw["rooms_occ"], raw["rooms_avail"], raw["rev_usd"]
        prop_rec["rooms_occ"]   = ro
        prop_rec["rooms_avail"] = ra
        prop_rec["rev_usd"]     = ru
        prop_rec["occ"] = round(ro / ra, 6) if ra > 0 else None
        prop_rec["adr"] = round(ru / ro, 4) if ro > 0 else None

    updated = {p: 0 for p in actuals}
    inserted = 0

    # Update existing records
    for date_str, day_rec in daily.items():
        for prop, days in actuals.items():
            if date_str in days:
                _apply_raw(day_rec.setdefault(prop, {}), days[date_str])
                updated[prop] += 1

    # Insert new dates that have at least one property with raw data,
    # but only within or after the existing date range (don't backfill history
    # before the first Africa weekly report, which had Budget/LE/PY data).
    all_actuals_dates = set()
    for prop_days in actuals.values():
        all_actuals_dates.update(prop_days.keys())

    for date_str in sorted(all_actuals_dates - set(daily.keys())):
        if date_str < range_start:
            continue
        new_rec = {}
        for prop, days in actuals.items():
            if date_str in days:
                prop_rec = {"occ_le": None, "occ_bgt": None, "occ_py": None,
                            "adr_le": None, "adr_bgt": None, "adr_py": None}
                _apply_raw(prop_rec, days[date_str])
                new_rec[prop] = prop_rec
        if new_rec:
            daily[date_str] = new_rec
            inserted += 1

    # Re-sort daily by date
    data["daily"] = dict(sorted(daily.items()))

    print("\nDaily records updated:")
    for p, n in updated.items():
        print(f"  {p}: {n}")
    print(f"New dates inserted: {inserted}")

    # 4.5 Populate PY (prior year) for 2026 dates from 2025 actuals.
    # PY for 2025 dates (= 2024 actuals) does not exist — kept null.
    # First clear any stale PY values (e.g. from old Africa weekly reports),
    # then repopulate 2026 PY from the same tracker/Radisson source files.
    PROPS_PY = ["ABAZ", "Pemba", "ASLV", "Radisson"]
    for date_str, day_rec in daily.items():
        for prop in PROPS_PY:
            rec = day_rec.get(prop)
            if rec is not None:
                rec["occ_py"] = None
                rec["adr_py"] = None

    py_filled = {p: 0 for p in PROPS_PY}
    for date_str, day_rec in daily.items():
        if not date_str.startswith("2026"):
            continue
        py_date = "2025" + date_str[4:]  # e.g. "2026-03-15" -> "2025-03-15"
        for prop in PROPS_PY:
            prop_rec = day_rec.get(prop)
            if prop_rec is None:
                continue
            py_raw = actuals.get(prop, {}).get(py_date)
            if py_raw is None:
                continue
            ro, ra, ru = py_raw["rooms_occ"], py_raw["rooms_avail"], py_raw["rev_usd"]
            prop_rec["occ_py"] = round(ro / ra, 6) if ra > 0 else None
            prop_rec["adr_py"] = round(ru / ro, 4) if ro > 0 else None
            py_filled[prop] += 1

    print("\nPY (2025 actuals) filled for 2026 dates:")
    for p, n in py_filled.items():
        print(f"  {p}: {n}")

    # 5. Normalize benchmarks: LE and Budget are monthly targets — enforce a
    #    single consistent value per property per month across all days.
    #    PY is excluded: it is daily actual data, not a monthly target.
    BENCH_FIELDS = ["occ_le", "occ_bgt", "adr_le", "adr_bgt"]
    PROPS_ALL = ["ABAZ", "Pemba", "ASLV", "Radisson"]

    # Collect all non-null values per (prop, month, field)
    monthly_vals = {}
    for date_str, day_rec in daily.items():
        month = date_str[:7]
        for prop in PROPS_ALL:
            prop_rec = day_rec.get(prop, {})
            for field in BENCH_FIELDS:
                v = prop_rec.get(field)
                if v is not None:
                    monthly_vals.setdefault((prop, month, field), []).append(v)

    # Compute the monthly benchmark (average of available values)
    monthly_bench = {k: sum(vs) / len(vs) for k, vs in monthly_vals.items()}

    # Write back to every day
    bench_applied = 0
    for date_str, day_rec in daily.items():
        month = date_str[:7]
        for prop in PROPS_ALL:
            prop_rec = day_rec.get(prop)
            if prop_rec is None:
                continue
            for field in BENCH_FIELDS:
                key = (prop, month, field)
                if key in monthly_bench:
                    prop_rec[field] = round(monthly_bench[key], 6)
                    bench_applied += 1

    print(f"Benchmark cells normalised: {bench_applied}")

    # 6. Save
    with open(DATA_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, separators=(",", ":"))
    print(f"\nSaved -> {DATA_JSON}")


if __name__ == "__main__":
    main()
