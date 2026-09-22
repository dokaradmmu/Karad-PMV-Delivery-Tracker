"""
PMV Toolkit Delivery Tracker — Core Processing Logic (Karad Division, v1)
===========================================================================

This module implements the core data-processing pipeline described and
agreed with the user. It is deliberately kept separate from any UI
(Streamlit) code so it can be unit-tested standalone against the sample
files before wiring up the app.

Pipeline:
  1. Load Office Master File (Karad's 14-column structure) -> jurisdiction lookup by PIN
  2. Load OLD and NEW PMV report CSVs, matched by COLUMN NAME (not position),
     because the source website's export schema drifts (e.g. an extra blank
     column appeared in the newer sample file).
  3. Normalize PIN codes on both sides.
  4. Split each report into KARAD rows (PIN found in master) vs OTHER DIVISIONS
     rows (PIN not found in master) and tag Karad rows with Division /
     Sub Division / Taluka / PIN.
  5. Diff OLD vs NEW using "PMV Application Number" as the match key (unique
     per row within each file) to find newly-delivered toolkits.
  6. Apply a persistent, user-curated RTS (Returned to Sender) exclusion list
     to pull RTS items out of "Pending" into their own sheet.
  7. Produce a dict of DataFrames, one per output sheet, plus a summary dict
     for the Streamlit on-screen dashboard.

No staleness/aging flagging (explicitly dropped by the user - PMV "Pending"
status can originate from booking/transmission stage, not just delivery, so
a pending-duration heuristic would be misleading).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# Canonical column names used internally, decoupled from the exact source
# header text. If the PMV website tweaks header wording again, only these
# constants need updating.
# --------------------------------------------------------------------------

COL_BARCODE = "Bar Code ID"
COL_APP_NO = "PMV Application Number"
COL_ARTISAN_NAME = "Artisan Name"
COL_MOBILE = "Mobile No"

# The PMV export schema changed on/around 22.09.2026: "Mobile No" was
# renamed to "New Mobile No", with two extra columns inserted
# ("Previous Mobile No", "Mobile No Changed Date"). We normalize any of
# these newer-schema exports back down to the original schema at load
# time (see load_pmv_report), so everything downstream — jurisdiction
# tagging, the diff logic, the Excel writer, the Streamlit UI — stays
# unchanged regardless of which export schema a given file uses.
NEW_SCHEMA_MOBILE_COL = "New Mobile No"
NEW_SCHEMA_EXTRA_COLS = ["Previous Mobile No", "Mobile No Changed Date"]
COL_ADDRESS = "Artisan Current Address"
COL_CIRCLE = "Circle Name"
COL_DIVISION_RAW = "Division Name"  # NOTE: unusable in source data, see below
COL_PIN = "Artisan Pin Code"
COL_STAFF_ASSIGNED = "Delivery Staff Assigned UnAssigned"
COL_STATUS = "Toolkit Delivery Status"

STATUS_DELIVERED = "Delivered"
STATUS_PENDING = "Pending"

REQUIRED_PMV_COLUMNS = [
    COL_BARCODE, COL_APP_NO, COL_ARTISAN_NAME, COL_MOBILE, COL_ADDRESS,
    COL_CIRCLE, COL_PIN, COL_STAFF_ASSIGNED, COL_STATUS,
]

# Office Master File canonical columns (Karad-specific v1: fixed schema).
# v2 will replace this fixed list with a user-driven column-mapping step.
MASTER_COL_DIVISION = "Division Office Name"
MASTER_COL_SUBDIVISION = "Sub Division Name"
MASTER_COL_SUBOFFICE = "Sub Office Name"
MASTER_COL_OFFICE_NAME = "Office Name"
MASTER_COL_OFFICE_TYPE = "Office Type Code"
MASTER_COL_PIN = "PIN"
MASTER_COL_TALUKA = "Taluka"


# --------------------------------------------------------------------------
# Loaders
# --------------------------------------------------------------------------

def load_pmv_report(csv_path: str | Path) -> pd.DataFrame:
    """
    Load a PMV toolkit delivery report CSV.

    Reads by column NAME, not position, so an extra/missing/reordered column
    in a given export (observed between the two sample files) doesn't
    silently misalign data. Any unnamed/blank columns the website injects
    are dropped. Any genuinely missing required column raises immediately
    rather than producing a wrong-but-plausible output.
    """
    df = pd.read_csv(csv_path, encoding="utf-8-sig", dtype=str)

    # Drop injected blank/unnamed columns (e.g. the stray column seen in the
    # newer sample export). These carry no data and no usable header.
    unnamed = [c for c in df.columns if c.strip() == "" or c.startswith("Unnamed")]
    if unnamed:
        df = df.drop(columns=unnamed)

    df.columns = [c.strip() for c in df.columns]

    # Normalize newer-schema exports (22.09.2026 onward) back to the
    # original schema: rename "New Mobile No" -> COL_MOBILE, and discard
    # "Previous Mobile No" / "Mobile No Changed Date" entirely so they
    # never reach the output sheets. Older exports (pre-22.09, already
    # using COL_MOBILE directly) pass through unchanged.
    if NEW_SCHEMA_MOBILE_COL in df.columns:
        df = df.rename(columns={NEW_SCHEMA_MOBILE_COL: COL_MOBILE})
    extra_present = [c for c in NEW_SCHEMA_EXTRA_COLS if c in df.columns]
    if extra_present:
        df = df.drop(columns=extra_present)

    missing = [c for c in REQUIRED_PMV_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"{csv_path}: missing expected column(s) {missing}. "
            f"Found columns: {list(df.columns)}. "
            "The source export format may have changed - check before proceeding."
        )

    # Drop fully blank rows the export sometimes tails with.
    df = df.dropna(how="all")

    # Normalize whitespace on key text fields.
    for col in [COL_APP_NO, COL_STATUS, COL_PIN]:
        df[col] = df[col].astype(str).str.strip()

    return df


def normalize_pin(value) -> "str | None":
    """Coerce a PIN-like value to a clean 6-digit string, or None if invalid."""
    if value is None:
        return None
    s = re.sub(r"\D", "", str(value))
    if len(s) != 6:
        return None
    return s


def load_office_master(xlsx_path: str | Path, sheet_name: str = "Sheet1") -> pd.DataFrame:
    """
    Load the Karad Office Master File.

    v1 assumes the fixed 14-column Karad structure confirmed from the sample
    file. Returns one row per (PIN) with Division/Sub Division/Taluka
    de-duplicated - since PIN -> Division and PIN -> Taluka are effectively
    1:1 in this master (verified against the sample), but PIN -> Office Name
    is NOT 1:1 (multiple BPOs can share a PIN), so Office Name is
    intentionally NOT included in the per-PIN jurisdiction lookup.
    """
    master = pd.read_excel(xlsx_path, sheet_name=sheet_name, dtype=str)
    master.columns = [c.strip() for c in master.columns]

    required = [MASTER_COL_DIVISION, MASTER_COL_SUBDIVISION, MASTER_COL_PIN, MASTER_COL_TALUKA]
    missing = [c for c in required if c not in master.columns]
    if missing:
        raise ValueError(
            f"{xlsx_path}: missing expected master column(s) {missing}. "
            f"Found columns: {list(master.columns)}."
        )

    master[MASTER_COL_PIN] = master[MASTER_COL_PIN].apply(normalize_pin)
    master = master.dropna(subset=[MASTER_COL_PIN])

    # Collapse to one row per PIN. Flag (but do not silently hide) any PIN
    # that maps to more than one Division or Taluka - that would indicate a
    # data-quality problem in the master file worth surfacing to the user.
    grouped = master.groupby(MASTER_COL_PIN)
    div_counts = grouped[MASTER_COL_DIVISION].nunique()
    taluka_counts = grouped[MASTER_COL_TALUKA].nunique(dropna=True)
    ambiguous_division_pins = div_counts[div_counts > 1].index.tolist()
    ambiguous_taluka_pins = taluka_counts[taluka_counts > 1].index.tolist()
    if ambiguous_division_pins:
        raise ValueError(
            f"Office Master File has PIN(s) mapping to multiple Divisions: "
            f"{ambiguous_division_pins}. This must be resolved in the master "
            f"file before jurisdiction tagging can be trusted."
        )

    lookup = (
        master.drop_duplicates(subset=[MASTER_COL_PIN])
        .set_index(MASTER_COL_PIN)[[MASTER_COL_DIVISION, MASTER_COL_SUBDIVISION, MASTER_COL_TALUKA]]
        .rename(columns={
            MASTER_COL_DIVISION: "Division",
            MASTER_COL_SUBDIVISION: "Sub Division",
            MASTER_COL_TALUKA: "Taluka",
        })
    )
    lookup.attrs["ambiguous_taluka_pins"] = ambiguous_taluka_pins
    return lookup


def load_rts_exclusion_list(path: "str | Path | None") -> set[str]:
    """
    Load the persistent RTS (Returned to Sender) exclusion list.

    Expected as a simple CSV/Excel with one column: 'PMV Application Number'.
    Returns an empty set if the path is None or the file doesn't exist yet
    (first run before any RTS cases have been confirmed).
    """
    if path is None:
        return set()
    p = Path(path)
    if not p.exists():
        return set()
    if p.suffix.lower() == ".csv":
        df = pd.read_csv(p, dtype=str)
    else:
        df = pd.read_excel(p, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    if COL_APP_NO not in df.columns:
        raise ValueError(f"RTS exclusion list {path} must have a '{COL_APP_NO}' column.")
    return set(df[COL_APP_NO].astype(str).str.strip().dropna())


# --------------------------------------------------------------------------
# Core pipeline
# --------------------------------------------------------------------------

@dataclass
class ProcessResult:
    delivered: pd.DataFrame
    pending: pd.DataFrame
    returned_to_sender: pd.DataFrame
    newly_delivered: pd.DataFrame
    other_divisions: pd.DataFrame
    summary: dict = field(default_factory=dict)
    old_date_label: str = ""
    new_date_label: str = ""


def _tag_jurisdiction(df: pd.DataFrame, master_lookup: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Split a PMV report DataFrame into (karad_rows, other_division_rows), tagging Karad rows."""
    df = df.copy()
    df["_pin_norm"] = df[COL_PIN].apply(normalize_pin)

    tagged = df.join(master_lookup, on="_pin_norm")
    is_karad = tagged["_pin_norm"].isin(master_lookup.index)

    karad_rows = tagged[is_karad].drop(columns=["_pin_norm"])
    other_rows = tagged[~is_karad].drop(columns=["_pin_norm", "Division", "Sub Division", "Taluka"])
    return karad_rows, other_rows


def _extract_date_label(filename: str) -> str:
    """Best-effort extraction of a dd.mm.yyyy-style date label from a report filename."""
    m = re.search(r"(\d{1,2})[._-](\d{1,2})[._-](\d{4})", filename)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}.{int(mo):02d}.{y}"
    return filename


def process(
    old_csv_path: str | Path,
    new_csv_path: str | Path,
    master_xlsx_path: str | Path,
    rts_list_path: "str | Path | None" = None,
    rts_app_numbers: "set[str] | None" = None,
    old_date_label: "str | None" = None,
    new_date_label: "str | None" = None,
) -> ProcessResult:
    """
    Run the full pipeline and return all output sheets + a summary dict.

    RTS exclusions can be supplied either as a file path (`rts_list_path`,
    used for local/CLI testing) or directly as an in-memory set
    (`rts_app_numbers`, used by the Streamlit app after fetching the
    persistent list from GitHub). If both are given, `rts_app_numbers`
    takes precedence.

    `old_date_label` / `new_date_label` let a caller override the filename-
    derived date (Streamlit uploads don't always keep a parseable filename).
    """

    def _label_for(explicit_label, source, default_label):
        if explicit_label:
            return explicit_label
        # `source` may be a filesystem path (str/Path) or a file-like object
        # (e.g. Streamlit's UploadedFile, which exposes .name as the
        # original filename but is not itself path-like).
        name = source.name if hasattr(source, "name") else source
        try:
            return _extract_date_label(Path(name).name)
        except TypeError:
            return default_label

    old_date_label = _label_for(old_date_label, old_csv_path, "Old Report")
    new_date_label = _label_for(new_date_label, new_csv_path, "New Report")

    master_lookup = load_office_master(master_xlsx_path)
    old_df = load_pmv_report(old_csv_path)
    new_df = load_pmv_report(new_csv_path)

    old_karad, _old_other = _tag_jurisdiction(old_df, master_lookup)
    new_karad, new_other = _tag_jurisdiction(new_df, master_lookup)

    if rts_app_numbers is not None:
        rts_set = set(rts_app_numbers)
    else:
        rts_set = load_rts_exclusion_list(rts_list_path)

    # --- Delivered (new file, Karad only) ---
    delivered = new_karad[new_karad[COL_STATUS] == STATUS_DELIVERED].copy()

    # --- Pending (new file, Karad only), RTS split out ---
    pending_all = new_karad[new_karad[COL_STATUS] == STATUS_PENDING].copy()
    is_rts = pending_all[COL_APP_NO].isin(rts_set)
    returned_to_sender = pending_all[is_rts].copy()
    pending = pending_all[~is_rts].copy()

    # --- Newly delivered: Pending in OLD -> Delivered in NEW, matched on Application No ---
    old_pending_apps = set(
        old_karad.loc[old_karad[COL_STATUS] == STATUS_PENDING, COL_APP_NO]
    )
    newly_delivered = delivered[delivered[COL_APP_NO].isin(old_pending_apps)].copy()

    summary = {
        "old_report_date": old_date_label,
        "new_report_date": new_date_label,
        "new_report_total_rows": len(new_df),
        "new_report_karad_rows": len(new_karad),
        "new_report_other_division_rows": len(new_other),
        "delivered_count": len(delivered),
        "pending_count": len(pending),
        "returned_to_sender_count": len(returned_to_sender),
        "newly_delivered_count": len(newly_delivered),
        "old_report_karad_rows": len(old_karad),
        "old_report_pending_count": int((old_karad[COL_STATUS] == STATUS_PENDING).sum()),
    }

    if master_lookup.attrs.get("ambiguous_taluka_pins"):
        summary["warning_ambiguous_taluka_pins"] = master_lookup.attrs["ambiguous_taluka_pins"]

    return ProcessResult(
        delivered=delivered,
        pending=pending,
        returned_to_sender=returned_to_sender,
        newly_delivered=newly_delivered,
        other_divisions=new_other,
        summary=summary,
        old_date_label=old_date_label,
        new_date_label=new_date_label,
    )
