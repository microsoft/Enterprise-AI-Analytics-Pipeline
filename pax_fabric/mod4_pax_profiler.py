"""
Module 4: pax_profiler — Record Shape Analysis
================================================
Migrated from: PAX_Purview_Audit_Log_Processor_v1.11.1.ps1 Lines 7725–7850
Level: 0 (no hard dependencies)

Provides optional diagnostic profiling for audit records:
- Shape key generation and caching
- Recursive JSON depth measurement
- Per-record structural statistics (operation counts, record types, depth distribution)
- Summary output

Design: Injected into pax_data_transform as an optional callback (profile_fn).
No import-level coupling to other PAX modules.
"""

from __future__ import annotations

from typing import Any, Optional

# ---------------------------------------------------------------------------
# Module-level state (mirrors PS $script:profiler and $script:shapeCache)
# ---------------------------------------------------------------------------

_profiler: dict[str, Any] = {
    "rows": 0,
    "operations": {},       # {operation_name: count}
    "record_types": {},     # {record_type: count}
    "has_copilot": 0,
    "max_depth": 0,
    "depth_counts": {},     # {depth_int: count}
    "max_array_len": 0,
}

_shape_cache: dict[str, dict[str, Any]] = {}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_record_shape_key(audit_data: Any) -> str:
    """
    Generate a composite cache key: 'RecordType|Operation|HasCopilotEventData'.

    PS equivalent: Get-RecordShapeKey (L7737)
    """
    try:
        rt = _safe_get(audit_data, "RecordType", "")
    except Exception:
        rt = ""
    try:
        op = _safe_get(audit_data, "Operation", "")
    except Exception:
        op = ""
    try:
        has_copilot = _has_key(audit_data, "CopilotEventData")
    except Exception:
        has_copilot = False
    return f"{rt}|{op}|{has_copilot}"


def get_record_shape(audit_data: Any) -> Optional[dict[str, Any]]:
    """
    Analyze and cache the structural shape of an audit record.

    Returns dict with: RecordType, Operation, HasCopilot, Depth, Mode.
    PS equivalent: Get-RecordShape (L7751)
    """
    if audit_data is None:
        return None

    key = get_record_shape_key(audit_data)
    if key in _shape_cache:
        return _shape_cache[key]

    shape: dict[str, Any] = {}
    try:
        shape["RecordType"] = _safe_get(audit_data, "RecordType", "")
        shape["Operation"] = _safe_get(audit_data, "Operation", "")
    except Exception:
        pass
    try:
        shape["HasCopilot"] = _has_key(audit_data, "CopilotEventData")
    except Exception:
        shape["HasCopilot"] = False
    try:
        shape["Depth"] = get_json_depth(audit_data, 0)
    except Exception:
        shape["Depth"] = 0
    shape["Mode"] = "Copilot" if shape.get("HasCopilot") else "AuditData"

    _shape_cache[key] = shape
    return shape


def reset_profiler() -> None:
    """
    Initialize/reset the profiler state to zero counts and empty tracking dicts.
    Does NOT clear the shape cache (matches PS Reset-Profiler which only resets
    $script:profiler, not $script:shapeCache).

    PS equivalent: Reset-Profiler (L7768)
    """
    _profiler.clear()
    _profiler.update({
        "rows": 0,
        "operations": {},
        "record_types": {},
        "has_copilot": 0,
        "max_depth": 0,
        "depth_counts": {},
        "max_array_len": 0,
    })


def get_json_depth(node: Any, d: int = 0) -> int:
    """
    Recursively measure the maximum nesting depth of a JSON-like object graph.
    Also tracks the longest array encountered in _profiler['max_array_len'].

    PS equivalent: Get-JsonDepth (L7780)
    """
    if node is None or _is_scalar(node):
        return d

    if isinstance(node, dict):
        max_d = d
        for v in node.values():
            max_d = max(max_d, get_json_depth(v, d + 1))
        return max_d

    if isinstance(node, (list, tuple)):
        max_d = d
        length = 0
        for el in node:
            max_d = max(max_d, get_json_depth(el, d + 1))
            length += 1
        if length > _profiler["max_array_len"]:
            _profiler["max_array_len"] = length
        return max_d

    # Object with attributes (e.g. dataclass) — treat like dict
    if hasattr(node, "__dict__") and not isinstance(node, type):
        max_d = d
        for v in vars(node).values():
            max_d = max(max_d, get_json_depth(v, d + 1))
        return max_d

    return d


def profile_audit_data(audit_data: Any) -> None:
    """
    Collect structural statistics about an audit record into the profiler state.

    PS equivalent: Profile-AuditData (L7797)
    """
    if audit_data is None:
        return

    try:
        _profiler["rows"] += 1

        # Operation
        try:
            op = _safe_get(audit_data, "Operation", None)
            if op and str(op).strip():
                op_str = str(op)
                _profiler["operations"][op_str] = _profiler["operations"].get(op_str, 0) + 1
        except Exception:
            pass

        # RecordType
        try:
            rt = _safe_get(audit_data, "RecordType", None)
            if rt is not None and str(rt).strip():
                rt_str = str(rt)
                _profiler["record_types"][rt_str] = _profiler["record_types"].get(rt_str, 0) + 1
        except Exception:
            pass

        # CopilotEventData presence
        try:
            if _has_key(audit_data, "CopilotEventData"):
                _profiler["has_copilot"] += 1
        except Exception:
            pass

        # Depth & arrays
        depth = get_json_depth(audit_data, 0)
        if depth > _profiler["max_depth"]:
            _profiler["max_depth"] = depth
        _profiler["depth_counts"][depth] = _profiler["depth_counts"].get(depth, 0) + 1

    except Exception:
        pass


def write_profiler_summary(
    top_ops: int = 20,
    top_depths: int = 10,
    *,
    log_fn=None,
) -> str:
    """
    Format and return a profiler summary string.

    If log_fn is provided (e.g. write_log_host from mod3), it will also be called
    for each line. Otherwise, lines are printed to stdout.

    PS equivalent: Write-ProfilerSummary (L7828)

    Returns the full summary as a string (for testing/capture).
    """
    lines: list[str] = []
    try:
        header = (
            f"Profiler: Rows={_profiler['rows']}, "
            f"MaxDepth={_profiler['max_depth']}, "
            f"MaxArrayLen={_profiler['max_array_len']}, "
            f"HasCopilot={_profiler['has_copilot']}"
        )
        lines.append(header)

        if _profiler["operations"]:
            lines.append(f"Profiler: Operations (top {top_ops}):")
            sorted_ops = sorted(
                _profiler["operations"].items(), key=lambda x: x[1], reverse=True
            )
            for name, count in sorted_ops[:top_ops]:
                lines.append(f"  {name}: {count}")

        if _profiler["depth_counts"]:
            lines.append(f"Profiler: Depth distribution (top {top_depths}):")
            sorted_depths = sorted(
                _profiler["depth_counts"].items(), key=lambda x: x[1], reverse=True
            )
            for depth_val, count in sorted_depths[:top_depths]:
                lines.append(f"  Depth {depth_val}: {count}")

    except Exception:
        pass

    output = "\n".join(lines)
    if log_fn:
        for line in lines:
            try:
                log_fn(line)
            except Exception:
                pass
    else:
        for line in lines:
            print(line)

    return output


# ---------------------------------------------------------------------------
# Accessor for profiler state (for testing / external reads)
# ---------------------------------------------------------------------------

def clear_shape_cache() -> None:
    """
    Explicitly clear the shape cache.
    PS equivalent: $script:shapeCache = @{} (done at script initialization, not in Reset-Profiler).
    """
    _shape_cache.clear()


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _safe_get(obj: Any, key: str, default: Any = None) -> Any:
    """Safely get a property from a dict or object."""
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _has_key(obj: Any, key: str) -> bool:
    """Check if a key/attribute exists on a dict or object."""
    if isinstance(obj, dict):
        return key in obj
    return hasattr(obj, key)


def _is_scalar(value: Any) -> bool:
    """
    Test whether a value is a scalar/primitive type.
    Mirrors PS Test-ScalarValue logic from mod2.
    """
    if value is None:
        return True
    return isinstance(value, (str, int, float, bool, bytes, complex))


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Basic smoke test
    reset_profiler()
    assert _profiler["rows"] == 0

    record = {
        "RecordType": "50",
        "Operation": "CopilotInteraction",
        "CopilotEventData": {"messages": [{"content": "hello"}]},
        "UserId": "user@contoso.com",
    }

    profile_audit_data(record)
    assert _profiler["rows"] == 1
    assert _profiler["operations"]["CopilotInteraction"] == 1
    assert _profiler["record_types"]["50"] == 1
    assert _profiler["has_copilot"] == 1
    assert _profiler["max_depth"] >= 2

    shape = get_record_shape(record)
    assert shape is not None
    assert shape["RecordType"] == "50"
    assert shape["Operation"] == "CopilotInteraction"
    assert shape["HasCopilot"] is True
    assert shape["Mode"] == "Copilot"
    assert shape["Depth"] >= 2

    key = get_record_shape_key(record)
    assert key == "50|CopilotInteraction|True"

    # Null input
    assert get_record_shape(None) is None
    profile_audit_data(None)
    assert _profiler["rows"] == 1  # unchanged

    # Summary output
    summary = write_profiler_summary(top_ops=5, top_depths=5)
    assert "Rows=1" in summary
    assert "CopilotInteraction" in summary

    # Reset
    reset_profiler()
    assert _profiler["rows"] == 0
    assert len(_shape_cache) > 0  # shape cache NOT cleared by reset_profiler (matches PS)

    # Explicit shape cache clear
    clear_shape_cache()
    assert len(_shape_cache) == 0

    print("PAX Profiler Module - OK (all self-tests passed)")
