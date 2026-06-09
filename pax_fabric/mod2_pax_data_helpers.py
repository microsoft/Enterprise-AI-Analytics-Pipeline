"""
PAX Module 2: pax_data_helpers
================================
Stateless, zero-dependency utility functions for the data-processing hot path.

Migrated from PAX_Purview_Audit_Log_Processor_v1.11.1.ps1
Source lines: L12913-13035, L14761, L14968-14981, L15044, L15225-15242, L15867

Functions (10):
  1. parse_date_safe          — Culture-invariant date parsing (PS Parse-DateSafe L12913)
  2. format_date_purview_fast — UTC ISO 8601 formatter       (PS Format-DatePurviewFast L12988)
  3. bool_tf_fast             — Truthy → 'TRUE'/'FALSE'      (PS BoolTFFast L13003)
  4. to_json_if_object_fast   — Scalar passthrough / JSON     (PS ToJsonIfObjectFast L13021)
  5. get_array_fast           — Safe array extraction          (PS GetArrayFast L13028)
  6. new_fast_row             — Dict → object builder          (PS New-FastRow L15225)
  7. test_scalar_value        — Scalar type check              (PS Test-ScalarValue L14761)
  8. get_safe_property        — Safe dict/obj key access       (PS Get-SafeProperty L15044)
  9. select_first_non_null    — First non-null/non-empty       (PS Select-FirstNonNull L15867)
 10. to_record_array          — Normalize input to list        (PS To-RecordArray L14968)

Hard dependencies: None
"""

from __future__ import annotations

import json
import re
import uuid
from datetime import datetime, timedelta, timezone
from typing import Any, Optional


# ===========================================================================
# PRE-COMPILED REGEXES (match PS $script:RegexTrueFalse etc.)
# ===========================================================================

_RE_TRUE_FALSE = re.compile(r"^(?:true|false)$", re.IGNORECASE)
_RE_YES_1 = re.compile(r"^(?:yes|1)$", re.IGNORECASE)
_RE_NO_0 = re.compile(r"^(?:no|0)$", re.IGNORECASE)

# ISO 8601 date formats (tried in order — most specific first)
_ISO_FORMATS: list[str] = [
    "%Y-%m-%dT%H:%M:%S.%fZ",       # yyyy-MM-ddTHH:mm:ss.ffffffZ
    "%Y-%m-%dT%H:%M:%SZ",          # yyyy-MM-ddTHH:mm:ssZ
    "%Y-%m-%dT%H:%M:%S.%f",        # yyyy-MM-ddTHH:mm:ss.ffffff (no Z)
    "%Y-%m-%dT%H:%M:%S",           # yyyy-MM-ddTHH:mm:ss
    "%Y-%m-%d %H:%M:%S.%f",        # yyyy-MM-dd HH:mm:ss.fff
    "%Y-%m-%d %H:%M:%S",           # yyyy-MM-dd HH:mm:ss
    "%Y-%m-%d",                    # yyyy-MM-dd
]

# US date formats (Purview API returns M/d/yyyy regardless of client locale)
_US_FORMATS: list[str] = [
    "%m/%d/%Y %H:%M:%S",           # M/d/yyyy HH:mm:ss
    "%m/%d/%Y %I:%M:%S %p",        # M/d/yyyy h:mm:ss tt
    "%m/%d/%Y",                    # M/d/yyyy
]


# ===========================================================================
# 1. parse_date_safe — PS Parse-DateSafe (L11654-L11727)
# ===========================================================================

def parse_date_safe(date_value: Any) -> Optional[datetime]:
    """
    Culture-invariant date parsing that handles Purview API date formats.

    Purview API returns dates in US format (M/d/yyyy HH:mm:ss) regardless of
    client locale. This function safely parses such dates on any system.

    Tries ISO 8601 formats first, then US formats, then generic parsing.

    Returns:
        datetime (UTC-aware) on success, None on failure.
    """
    # Already a datetime? Return as-is
    if isinstance(date_value, datetime):
        return date_value

    # Null/empty? Return None
    if date_value is None:
        return None

    date_str = str(date_value).strip()
    if not date_str:
        return None

    # Detect and extract timezone offset (+HH:MM or -HH:MM) for AdjustToUniversal parity
    # PS uses DateTimeStyles.AdjustToUniversal — the parsed time is converted to UTC.
    tz_offset = None
    clean_str = date_str
    if len(date_str) > 6 and date_str[-6] in ('+', '-') and date_str[-3] == ':':
        try:
            sign = 1 if date_str[-6] == '+' else -1
            oh = int(date_str[-5:-3])
            om = int(date_str[-2:])
            tz_offset = timezone(timedelta(hours=sign * oh, minutes=sign * om))
            clean_str = date_str[:-6]
        except (ValueError, IndexError):
            pass

    # Try ISO 8601 formats first (most common from properly-formatted API responses)
    for fmt in _ISO_FORMATS:
        try:
            dt = datetime.strptime(clean_str, fmt)
            if tz_offset is not None:
                dt = dt.replace(tzinfo=tz_offset)
                return dt.astimezone(timezone.utc)
            return dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue

    # Try US formats (what Purview actually returns — causes UK locale issues)
    for fmt in _US_FORMATS:
        try:
            dt = datetime.strptime(clean_str, fmt)
            if tz_offset is not None:
                dt = dt.replace(tzinfo=tz_offset)
                return dt.astimezone(timezone.utc)
            return dt.replace(tzinfo=timezone.utc)
        except ValueError:
            continue

    # Last resort: try the original string with fromisoformat (Python 3.11+)
    try:
        dt = datetime.fromisoformat(date_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except (ValueError, TypeError):
        pass

    return None


# ===========================================================================
# 2. format_date_purview_fast — PS Format-DatePurviewFast (L11729-L11742)
# ===========================================================================

def format_date_purview_fast(dt: Any) -> str:
    """
    Format a datetime to Purview's canonical ISO 8601 UTC string.

    Returns: 'yyyy-MM-ddTHH:mm:ss.fffZ' or '' on failure.
    """
    if dt is None:
        return ""

    try:
        if isinstance(dt, datetime):
            utc_dt = dt.astimezone(timezone.utc) if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
            return utc_dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{utc_dt.microsecond // 1000:03d}Z"

        parsed = parse_date_safe(dt)
        if parsed is None:
            return ""
        utc_dt = parsed.astimezone(timezone.utc) if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)
        return utc_dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{utc_dt.microsecond // 1000:03d}Z"
    except Exception:
        return ""


# ===========================================================================
# 3. bool_tf_fast — PS BoolTFFast (L11744-L11752)
# ===========================================================================

def bool_tf_fast(v: Any) -> str:
    """
    Convert a truthy/falsy value to 'TRUE', 'FALSE', or passthrough string.

    Matches PS behavior:
      - None → ''
      - bool → 'TRUE'/'FALSE'
      - 'true'/'false' (case-insensitive) → uppercased
      - 'yes'/'1' → 'TRUE'
      - 'no'/'0' → 'FALSE'
      - anything else → str(v)
    """
    if v is None:
        return ""

    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"

    v_str = str(v)

    if _RE_TRUE_FALSE.match(v_str):
        return v_str.upper()

    if _RE_YES_1.match(v_str):
        return "TRUE"

    if _RE_NO_0.match(v_str):
        return "FALSE"

    return v_str


# ===========================================================================
# 4. to_json_if_object_fast — PS ToJsonIfObjectFast (L11762-L11767)
# ===========================================================================

def to_json_if_object_fast(v: Any) -> Any:
    """
    If v is a scalar, return as-is. If it's a complex object (dict/list), JSON-serialize it.

    Returns: scalar value or JSON string, '' for None.
    """
    if v is None:
        return ""

    if test_scalar_value(v):
        return v

    try:
        return json.dumps(v, default=str, ensure_ascii=False)
    except Exception:
        return str(v)


# ===========================================================================
# 5. get_array_fast — PS GetArrayFast (L11769-L11776)
# ===========================================================================

def get_array_fast(parent: Any, name: str) -> list:
    """
    Safely extract a named property as a list.

    Returns: list (possibly empty), never None.
    """
    val = get_safe_property(parent, name)
    if val is None:
        return []

    if isinstance(val, (list, tuple)):
        return list(val)

    if isinstance(val, (str, dict)):
        return [val]

    # Try to iterate
    try:
        return list(val)
    except (TypeError, ValueError):
        return [val]


# ===========================================================================
# 7. test_scalar_value — PS Test-ScalarValue (L13394)
# ===========================================================================

# Scalar types that match PS: string, char, bool, int, long, double,
# decimal, float, datetime, guid
_SCALAR_TYPES = (str, bool, int, float, datetime, uuid.UUID)


def test_scalar_value(v: Any) -> bool:
    """
    Check if a value is a scalar (simple) type.

    Matches PS: null, string, char, bool, int, long, double, decimal, float,
    datetime, guid are all scalar.
    """
    if v is None:
        return True

    return isinstance(v, _SCALAR_TYPES)


# ===========================================================================
# 8. get_safe_property — PS Get-SafeProperty (L13677)
# ===========================================================================

def get_safe_property(obj: Any, name: str) -> Any:
    """
    Safely access a named property from a dict-like or object.

    Returns None if the object is None, the key doesn't exist, or access fails.
    """
    if obj is None:
        return None

    # dict-like access (most common in Python)
    if isinstance(obj, dict):
        return obj.get(name)

    # object attribute access
    try:
        return getattr(obj, name, None)
    except Exception:
        return None


# ===========================================================================
# 9. select_first_non_null — PS Select-FirstNonNull (L14500)
# ===========================================================================

def select_first_non_null(*values: Any) -> Any:
    """
    Return the first value that is not None and not empty-string.

    Matches PS: iterates values, returns first where $null -ne $v and '' -ne [string]$v.
    """
    for v in values:
        if v is not None and str(v) != "":
            return v
    return None


# ===========================================================================
# 10. to_record_array — PS To-RecordArray (L13601-L13615)
# ===========================================================================

def to_record_array(records: Any) -> list:
    """
    Normalize input to a flat list of records.

    Handles: None → [], single item → [item], list → list, nested iterables.
    In PS, PSObject/PSCustomObject (dict equivalent) is NOT IEnumerable — it wraps.
    """
    if records is None:
        return []

    if isinstance(records, (list, tuple)):
        return list(records)

    # dict and str are treated as single items (PS: PSObject is not IEnumerable)
    if isinstance(records, (str, dict)):
        return [records]

    # Try to iterate (generators, etc.)
    try:
        return list(records)
    except (TypeError, ValueError):
        return [records]


# ===========================================================================
# ENTRY POINT (for standalone testing)
# ===========================================================================

if __name__ == "__main__":
    import sys

    # Quick self-test
    ok = True
    errors: list[str] = []

    # parse_date_safe
    d = parse_date_safe("2025-10-01T12:30:45.123Z")
    if d is None or d.year != 2025 or d.month != 10:
        errors.append(f"parse_date_safe ISO failed: {d}")
        ok = False

    d2 = parse_date_safe("10/1/2025 14:30:00")
    if d2 is None or d2.day != 1:
        errors.append(f"parse_date_safe US failed: {d2}")
        ok = False

    # format_date_purview_fast
    fmt = format_date_purview_fast(d)
    if not fmt.endswith("Z") or "2025-10-01" not in fmt:
        errors.append(f"format_date_purview_fast failed: {fmt}")
        ok = False

    # bool_tf_fast
    if bool_tf_fast(True) != "TRUE" or bool_tf_fast(False) != "FALSE":
        errors.append("bool_tf_fast bool failed")
        ok = False
    if bool_tf_fast("yes") != "TRUE" or bool_tf_fast("no") != "FALSE":
        errors.append("bool_tf_fast yes/no failed")
        ok = False

    # test_scalar_value
    if not test_scalar_value("hello") or not test_scalar_value(42) or not test_scalar_value(None):
        errors.append("test_scalar_value failed")
        ok = False
    if test_scalar_value([1, 2]) or test_scalar_value({"a": 1}):
        errors.append("test_scalar_value non-scalar failed")
        ok = False

    # get_safe_property
    if get_safe_property({"a": 1}, "a") != 1 or get_safe_property({"a": 1}, "b") is not None:
        errors.append("get_safe_property failed")
        ok = False

    # select_first_non_null
    if select_first_non_null(None, "", "hello") != "hello":
        errors.append("select_first_non_null failed")
        ok = False

    # to_record_array
    if to_record_array(None) != [] or to_record_array("x") != ["x"] or to_record_array([1, 2]) != [1, 2]:
        errors.append("to_record_array failed")
        ok = False

    if ok:
        print("PAX Data Helpers Module - OK (all self-tests passed)")
        sys.exit(0)
    else:
        print("PAX Data Helpers Module - FAILED:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
