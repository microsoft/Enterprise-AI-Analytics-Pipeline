"""
Module 15: pax_append_merge — Append/Merge Infrastructure
==========================================================
Migrated from: PAX_Purview_Audit_Log_Processor_v1.11.2.ps1
Source lines: L7383–L7467 (ConvertTo-RecordIdExclusion), L7470–L7580 (Import-CsvDeduped),
              L7583–L7736 (Merge-UsersCsv), L7739–L7936 (Merge-FactCsv)

Level: 2 (depends on mod10_pax_csv_export for atomic writes)

Provides append/merge infrastructure for cross-run CSV union merges:
- BOM-tolerant CSV import with duplicate header deduplication
- Users CSV union merge (Retained / New / Departed classification)
- Fact CSV union merge (parameterized key column)
- Provenance tracking (Date_Added, Latest_Append_Date, In_Latest_Append)

External dependencies: None (stdlib only)
Design: Uses stdlib logging.getLogger(__name__).

PS-to-Python Function Mapping
──────────────────────────────────────────────────────────────────────────
│ #   │ PS Function                  │ PS Line │ Python Function            │
│─────│─────────────────────────────│─────────│────────────────────────────│
│ 101 │ Import-CsvDeduped            │ 7470    │ import_csv_deduped()       │
│ 102 │ Merge-UsersCsv               │ 7583    │ merge_users_csv()          │
│ 103 │ Merge-FactCsv                │ 7739    │ merge_fact_csv()           │
│     │ (provenance constants)       │         │ PROVENANCE_COLUMNS         │
│     │ (merge key constants)        │         │ USERS_MERGE_KEY, etc.      │
──────────────────────────────────────────────────────────────────────────

Test Results
────────────
Run: (pending)
"""

from __future__ import annotations

import csv
import io
import logging
import os
import shutil
from datetime import datetime
from typing import Optional

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PROVENANCE_COLUMNS: list[str] = ["Date_Added", "Latest_Append_Date", "In_Latest_Append"]

USERS_MERGE_KEY: str = "PersonId_Normalized"
FACT_MERGE_KEY_RAW: str = "Message_Id_Raw"
AUDIT_MERGE_KEY: str = "RecordId"


# ═══════════════════════════════════════════════════════════════════════════════
# IMPORT CSV DEDUPED — PS Import-CsvDeduped (L7470–L7580)
# ═══════════════════════════════════════════════════════════════════════════════

def import_csv_deduped(file_path: str) -> list[dict]:
    """Import CSV tolerating BOM-prefixed headers, duplicate columns, and
    empty trailing columns.

    PS Import-CsvDeduped (L7470–L7580):
    - Returns empty list if file missing or header-only
    - Strips UTF-8 BOM (\\ufeff) from header line
    - Trims whitespace from each header cell
    - Renames blank headers to '_blank'
    - Deduplicates headers: first occurrence keeps name,
      later duplicates get '_dup2', '_dup3', etc.
    - Data rows parsed normally against the rewritten header

    Returns:
        list[dict] — one dict per data row. Empty list when file is
        missing, empty, or header-only.
    """
    if not os.path.isfile(file_path):
        return []

    with open(file_path, "r", encoding="utf-8-sig") as fh:
        raw_lines = fh.readlines()

    if not raw_lines:
        return []

    # Strip BOM if encoding="utf-8-sig" didn't catch it (belt-and-suspenders)
    header_line = raw_lines[0]
    if header_line and header_line[0] == "\ufeff":
        header_line = header_line[1:]

    # Parse header respecting quoted fields
    header_cells = _parse_csv_line(header_line.rstrip("\r\n"))

    # Deduplicate header cells (PS logic: first keeps name, later get _dup2, _dup3…)
    seen: dict[str, int] = {}
    deduped_header: list[str] = []
    for cell in header_cells:
        name = cell.strip() if cell else ""
        if not name:
            name = "_blank"
        lower_name = name.lower()
        if lower_name in seen:
            seen[lower_name] += 1
            deduped_header.append(f"{name}_dup{seen[lower_name]}")
        else:
            seen[lower_name] = 1
            deduped_header.append(name)

    # No data rows → empty list
    data_lines = raw_lines[1:]
    if not data_lines:
        return []

    # Rebuild CSV text with the rewritten header for csv.DictReader
    new_header_line = ",".join(
        '"' + h.replace('"', '""') + '"' for h in deduped_header
    )
    rebuilt = new_header_line + "\n" + "".join(data_lines)
    reader = csv.DictReader(io.StringIO(rebuilt))
    return list(reader)


def _parse_csv_line(line: str) -> list[str]:
    """Parse a single CSV line respecting double-quoted fields.

    Returns list of unquoted field values.
    """
    fields: list[str] = []
    buf: list[str] = []
    in_quotes = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if in_quotes and i + 1 < len(line) and line[i + 1] == '"':
                buf.append('"')
                i += 2
                continue
            in_quotes = not in_quotes
        elif ch == "," and not in_quotes:
            fields.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    fields.append("".join(buf))
    return fields


# ═══════════════════════════════════════════════════════════════════════════════
# MERGE USERS CSV — PS Merge-UsersCsv (L7583–L7736)
# ═══════════════════════════════════════════════════════════════════════════════

def _normalize_person_id(row: dict) -> str:
    """Derive PersonId_Normalized from a row using the PS fallback chain.

    Fallback: PersonId_Normalized → userPrincipalName → PersonId
    Each lower() + strip() (matches Python normalization exactly).
    """
    for key in ("PersonId_Normalized", "userPrincipalName", "PersonId"):
        val = row.get(key)
        if val and str(val).strip():
            return str(val).strip().lower()
    return ""


def merge_users_csv(
    target_csv: str,
    current_csv: str,
    output_path: str,
    run_date: Optional[str] = None,
) -> dict:
    """Union-merge a target Users CSV with the current run's Users CSV.

    PS Merge-UsersCsv (L7583–L7736):
    - Merge key: PersonId_Normalized (with fallback derivation)
    - Retained users (in both): keep target's Date_Added + UserKey; In_Latest_Append=TRUE
    - New users (only in current): mint Date_Added=run_date; In_Latest_Append=TRUE
    - Departed users (only in target): carry forward; In_Latest_Append=FALSE
    - Latest_Append_Date = run_date on every row
    - TotalEmployees recomputed across the union
    - Atomic write via temp file + rename

    Args:
        target_csv: Path to the existing/seed Users CSV (may not exist).
        current_csv: Path to the current run's freshly-emitted Users CSV.
        output_path: Where to write the union CSV.
        run_date: Date stamp (default: today as 'yyyy-MM-dd').

    Returns:
        dict with keys: Retained, New, Departed, Union
    """
    if run_date is None:
        run_date = datetime.now().strftime("%Y-%m-%d")

    if not os.path.isfile(current_csv):
        raise FileNotFoundError(
            f"Merge-UsersCsv: current Users CSV not found: '{current_csv}'"
        )

    # Load both files with BOM/dup-header tolerance
    target_rows = import_csv_deduped(target_csv) if os.path.isfile(target_csv) else []
    current_rows = import_csv_deduped(current_csv)

    # Build dedup indexes with on-the-fly PersonId_Normalized derivation
    target_by_pid: dict[str, dict] = {}
    for row in target_rows:
        pid = _normalize_person_id(row)
        if not pid:
            continue
        # Back-fill PersonId_Normalized if absent
        if not row.get("PersonId_Normalized", "").strip():
            row["PersonId_Normalized"] = pid
        if pid not in target_by_pid:
            target_by_pid[pid] = row

    current_by_pid: dict[str, dict] = {}
    for row in current_rows:
        pid = _normalize_person_id(row)
        if not pid:
            continue
        if not row.get("PersonId_Normalized", "").strip():
            row["PersonId_Normalized"] = pid
        if pid not in current_by_pid:
            current_by_pid[pid] = row

    # Warning when target has no derivable key
    if target_rows and not target_by_pid:
        logger.warning(
            "Merge-UsersCsv: target Users CSV has no derivable dedup key "
            "(PersonId_Normalized / userPrincipalName / PersonId are all empty). "
            "Union will contain only the current-run rows. Target: %s",
            target_csv,
        )

    # Header union: current-run order first, then target-only columns, then provenance
    # PS uses [System.StringComparer]::OrdinalIgnoreCase — case-insensitive dedup
    hdr_set_lower: set[str] = set()  # lowercase keys for case-insensitive dedup
    hdr_order: list[str] = []

    def _add_headers(rows: list[dict]) -> None:
        if rows:
            for key in rows[0]:
                if key.lower() not in hdr_set_lower:
                    hdr_set_lower.add(key.lower())
                    hdr_order.append(key)

    _add_headers(current_rows)
    _add_headers(target_rows)
    for col in PROVENANCE_COLUMNS:
        if col.lower() not in hdr_set_lower:
            hdr_set_lower.add(col.lower())
            hdr_order.append(col)

    merged: list[dict] = []

    # 1. Current-run rows (retained + new)
    for row in current_rows:
        pid = _normalize_person_id(row)
        obj = {c: "" for c in hdr_order}
        obj.update(row)

        if pid and pid in target_by_pid:
            tr = target_by_pid[pid]
            # Target's UserKey wins (continuity)
            if tr.get("UserKey", "").strip():
                obj["UserKey"] = tr["UserKey"]
            # Preserve target's Date_Added
            tda = tr.get("Date_Added", "").strip()
            obj["Date_Added"] = tda if tda else run_date
        else:
            obj["Date_Added"] = run_date

        obj["Latest_Append_Date"] = run_date
        obj["In_Latest_Append"] = "TRUE"
        merged.append(obj)

    # 2. Departed users (in target, not in current)
    departed_count = 0
    for pid, tr in target_by_pid.items():
        if pid in current_by_pid:
            continue
        obj = {c: "" for c in hdr_order}
        obj.update(tr)
        tda = tr.get("Date_Added", "").strip()
        if not tda:
            obj["Date_Added"] = run_date
        obj["Latest_Append_Date"] = run_date
        obj["In_Latest_Append"] = "FALSE"
        merged.append(obj)
        departed_count += 1

    # 3. Recompute TotalEmployees
    union_count = len(merged)
    if "totalemployees" in hdr_set_lower:
        # Find the actual casing of the TotalEmployees column in hdr_order
        te_col = next((c for c in hdr_order if c.lower() == "totalemployees"), None)
        if te_col:
            for row in merged:
                row[te_col] = str(union_count)

    # 4. Atomic write
    tmp_path = output_path + ".merging"
    try:
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        with open(tmp_path, "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=hdr_order, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(merged)
        shutil.move(tmp_path, output_path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise

    # Compute tallies
    retained = sum(1 for pid in current_by_pid if pid in target_by_pid)
    new = sum(1 for pid in current_by_pid if pid not in target_by_pid)

    return {
        "Retained": retained,
        "New": new,
        "Departed": departed_count,
        "Union": union_count,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# MERGE FACT CSV — PS Merge-FactCsv (L7739–L7936)
# ═══════════════════════════════════════════════════════════════════════════════

def merge_fact_csv(
    target_csv: str,
    current_csv: str,
    output_path: str,
    key_column: str = "Message_Id_Raw",
    run_date: Optional[str] = None,
) -> dict:
    """Union-merge a target Fact CSV with the current run's Fact CSV.

    PS Merge-FactCsv (L7739–L7936):
    - Merge key: key_column parameter (default 'Message_Id_Raw' for rollup,
      pass 'RecordId' for raw audit CSV)
    - Retained (in both): keep target's Date_Added; for Message_Id_Raw key,
      also carry target's Message_Id INT (continuity)
    - New (only in current): mint Date_Added=run_date
    - Departed (only in target): carry forward; In_Latest_Append=FALSE
    - Latest_Append_Date = run_date on every row
    - Atomic write via temp file + rename

    Args:
        target_csv: Path to the existing/seed Fact CSV (may not exist).
        current_csv: Path to the current run's freshly-emitted Fact CSV.
        output_path: Where to write the union CSV.
        key_column: Column name to use as the dedup key.
        run_date: Date stamp (default: today as 'yyyy-MM-dd').

    Returns:
        dict with keys: Retained, New, Departed, Union
    """
    if run_date is None:
        run_date = datetime.now().strftime("%Y-%m-%d")

    if not os.path.isfile(current_csv):
        raise FileNotFoundError(
            f"Merge-FactCsv: current Fact CSV not found: '{current_csv}'"
        )

    target_rows = import_csv_deduped(target_csv) if os.path.isfile(target_csv) else []
    current_rows = import_csv_deduped(current_csv)

    # Warn if target is missing the key column
    if target_rows:
        target_headers = list(target_rows[0].keys())
        if key_column not in target_headers:
            logger.warning(
                "Merge-FactCsv: target Fact CSV is missing dedup key column '%s'. "
                "Cannot classify Retained / New / Departed rows reliably; "
                "treating ALL current-run rows as New. Target: %s",
                key_column,
                target_csv,
            )

    # Build dedup indexes
    target_by_key: dict[str, dict] = {}
    for row in target_rows:
        k = row.get(key_column, "").strip()
        if k and k not in target_by_key:
            target_by_key[k] = row

    current_by_key: dict[str, dict] = {}
    for row in current_rows:
        k = row.get(key_column, "").strip()
        if k and k not in current_by_key:
            current_by_key[k] = row

    # Header union: current-run order first, then target-only, then provenance
    # PS uses [System.StringComparer]::OrdinalIgnoreCase — case-insensitive dedup
    hdr_set_lower: set[str] = set()
    hdr_order: list[str] = []

    def _add_headers(rows: list[dict]) -> None:
        if rows:
            for key in rows[0]:
                if key.lower() not in hdr_set_lower:
                    hdr_set_lower.add(key.lower())
                    hdr_order.append(key)

    _add_headers(current_rows)
    _add_headers(target_rows)
    for col in PROVENANCE_COLUMNS:
        if col.lower() not in hdr_set_lower:
            hdr_set_lower.add(col.lower())
            hdr_order.append(col)

    merged: list[dict] = []

    # 1. Current-run rows (retained + new)
    for row in current_rows:
        k = row.get(key_column, "").strip()
        obj = {c: "" for c in hdr_order}
        obj.update(row)

        if k and k in target_by_key:
            tr = target_by_key[k]
            # For Message_Id_Raw key: target's Message_Id INT wins (continuity)
            if key_column == "Message_Id_Raw":
                tr_mid = tr.get("Message_Id", "").strip()
                if tr_mid:
                    obj["Message_Id"] = tr_mid
            # Preserve target's Date_Added
            tda = tr.get("Date_Added", "").strip()
            obj["Date_Added"] = tda if tda else run_date
        else:
            obj["Date_Added"] = run_date

        obj["Latest_Append_Date"] = run_date
        obj["In_Latest_Append"] = "TRUE"
        merged.append(obj)

    # 2. Departed rows (in target, not in current)
    departed_count = 0
    for k, tr in target_by_key.items():
        if k in current_by_key:
            continue
        obj = {c: "" for c in hdr_order}
        obj.update(tr)
        tda = tr.get("Date_Added", "").strip()
        if not tda:
            obj["Date_Added"] = run_date
        obj["Latest_Append_Date"] = run_date
        obj["In_Latest_Append"] = "FALSE"
        merged.append(obj)
        departed_count += 1

    # 3. Atomic write
    union_count = len(merged)
    tmp_path = output_path + ".merging"
    try:
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        with open(tmp_path, "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=hdr_order, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(merged)
        shutil.move(tmp_path, output_path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise

    # Compute tallies
    retained = sum(1 for k in current_by_key if k in target_by_key)
    new = sum(1 for k in current_by_key if k not in target_by_key)

    return {
        "Retained": retained,
        "New": new,
        "Departed": departed_count,
        "Union": union_count,
    }
