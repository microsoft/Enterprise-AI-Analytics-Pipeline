"""
PAX Module 3: pax_logging
============================
Logging, progress tracking, and display-safe utilities.

Migrated from PAX_Purview_Audit_Log_Processor_v1.11.1.ps1
Source lines: L1853-1873, L2209-2328, L6510-6551, L13102-13188

Functions (11):
  1. write_log                — Console + file logging        (PS Write-Log L2209)
  2. write_log_file           — File-only logging             (PS Write-LogFile L2223)
  3. write_log_host           — Colored console + file log    (PS Write-LogHost L2233)
  4. setup_host_logging       — Logging system initializer    (PS global:Write-Host L2245)
  5. get_masked_username      — Email masking for logs        (PS Get-MaskedUsername L2271)
  6. send_prompt_notification — System beep for prompts       (PS Send-PromptNotification L1853)
  7. set_progress_phase       — Phase state setter            (PS Set-ProgressPhase L13102)
  8. update_progress          — Progress bar update           (PS Update-Progress L13103)
  9. complete_progress        — Progress finalization         (PS Complete-Progress L13183)
 10. write_progress_tick      — Heartbeat tick                (PS Write-ProgressTick L13188)
 11. get_display_path         — Remote path display helper    (PS Get-DisplayPath L6510) *NEW v1.11.1*

Hard dependencies: None
"""

from __future__ import annotations

import logging
import os
import sys
import tempfile
import threading
import time
from datetime import datetime
from typing import Any, Optional


# ===========================================================================
# MODULE-LEVEL STATE (mirrors PS $script: scoped variables)
# ===========================================================================

# Log file path — set by setup_host_logging(); None until then.
# Early log entries are buffered in _log_buffer.
_log_file: Optional[str] = None
_log_buffer: list[str] = []

# Bootstrap log state (PS L1538-1596: $script:LogFileIsBootstrap, $script:BootstrapLogDir)
_log_file_is_bootstrap: bool = False
_bootstrap_log_dir: Optional[str] = None

# Progress state — mirrors PS $script:progressState
_progress_state: dict[str, Any] = {
    "weights": {"Parsing": 0.05, "Query": 0.85, "Explosion": 0.08, "Export": 0.02},
    "phase": "Query",
    "Parsing": {"current": 0, "total": 0},
    "Query": {"current": 0, "total": 0},
    "Explode": {"current": 0, "total": 0},
    "Export": {"current": 0, "total": 1},
}

_original_weights: Optional[dict[str, float]] = {"Parsing": 0.05, "Query": 0.85, "Explosion": 0.08, "Export": 0.02}

# Logger for this module — other modules use logging.getLogger(__name__)
# Fabric package uses a distinct logger name so notebook-side handlers can
# attach without colliding with any legacy ``pax`` logger left in the kernel.
_logger = logging.getLogger("pax_fabric")

# Thread lock for stdout writes.  Fabric / Jupyter notebooks replace
# sys.stdout with a custom stream that funnels output through the IOPub
# ZMQ channel.  When multiple threads (e.g. parallel Graph partitions)
# write concurrently, the kernel may silently drop messages or interleave
# partial lines.  Serializing all stdout writes behind one lock AND
# flushing immediately afterwards eliminates this class of output loss.
_stdout_lock = threading.Lock()

# Thread lock for log-FILE writes.  Every log-file write (both write_log()'s
# _write_to_log_file() and _FlushingFileHandler.emit(), below) opens the file,
# writes one entry, and closes it — serialized behind this lock so concurrent
# partition threads never interleave partial lines within the same open/write.
_file_write_lock = threading.Lock()


# ===========================================================================
# INTERNAL HELPERS
# ===========================================================================

def _timestamp() -> str:
    """Return a formatted timestamp matching PS '[yyyy-MM-dd HH:mm:ss]'."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _write_to_log_file(entry: str) -> None:
    """Append a log entry to the log file, or buffer it if file not yet set.

    Opens, writes, flushes, fsyncs, and closes the file on every call. On the
    Fabric Lakehouse Files/ (OneLake) mount, a write is not reliably visible
    to other readers (the Files pane, a separate notebook) until the file
    handle is actually closed — a Python-level flush() alone is not enough.
    Closing on every entry (rather than holding one handle open for the
    whole run) guarantees each line lands in OneLake immediately.
    """
    try:
        if _log_file:
            with _file_write_lock:
                with open(_log_file, "a", encoding="utf-8") as f:
                    f.write(entry + "\n")
                    f.flush()
                    os.fsync(f.fileno())
        else:
            _log_buffer.append(entry)
    except Exception:
        pass


def _flush_log_buffer() -> None:
    """Flush buffered log entries to the log file once it's available."""
    if _log_file and _log_buffer:
        try:
            with open(_log_file, "a", encoding="utf-8") as f:
                for entry in _log_buffer:
                    f.write(entry + "\n")
            _log_buffer.clear()
        except Exception:
            pass


# ===========================================================================
# 1. write_log — PS Write-Log (L2209-L2219)
# ===========================================================================

def write_log(message: str, level: str = "INFO") -> None:
    """
    Write a message to both console and log file.

    Matches PS Write-Log: echoes to host via Write-Host,
    then appends timestamped entry to log file (or buffer).

    Args:
        message: The message to log.
        level: Log level string (INFO, WARN, ERROR, DEBUG).
    """
    ts = _timestamp()
    log_entry = f"[{ts}] [{level}] {message}"

    # Echo to console — serialized + flushed to avoid notebook output loss
    with _stdout_lock:
        print(message, flush=True)

    # Write to file or buffer
    _write_to_log_file(log_entry)


# ===========================================================================
# 2. write_log_file — PS Write-LogFile (L2223-L2231)
# ===========================================================================

def write_log_file(message: str, level: str = "INFO") -> None:
    """
    Write a message to the log file ONLY (never echoes to console).

    Matches PS Write-LogFile: used for diagnostic detail (temp paths,
    argument vectors, stack traces) that should be captured in the run log
    but not clutter the customer-facing console.

    Args:
        message: The message to log.
        level: Log level string.
    """
    ts = _timestamp()
    log_entry = f"[{ts}] [{level}] {message}"
    _write_to_log_file(log_entry)


# ===========================================================================
# 3. write_log_host — PS Write-LogHost (L2233-L2243)
# ===========================================================================

# ANSI color map matching PS ConsoleColor names
_ANSI_COLORS: dict[str, str] = {
    "Black": "\033[30m",
    "DarkBlue": "\033[34m",
    "DarkGreen": "\033[32m",
    "DarkCyan": "\033[36m",
    "DarkRed": "\033[31m",
    "DarkMagenta": "\033[35m",
    "DarkYellow": "\033[33m",
    "Gray": "\033[37m",
    "DarkGray": "\033[90m",
    "Blue": "\033[94m",
    "Green": "\033[92m",
    "Cyan": "\033[96m",
    "Red": "\033[91m",
    "Magenta": "\033[95m",
    "Yellow": "\033[93m",
    "White": "\033[97m",
}

_ANSI_RESET = "\033[0m"


def write_log_host(message: str, foreground_color: str = "White") -> None:
    """
    Write a colored message to console AND log file.

    Matches PS Write-LogHost: displays with ForegroundColor,
    then appends timestamped entry to log file.

    Args:
        message: The message to display and log.
        foreground_color: PS ConsoleColor name (e.g., 'Green', 'Red', 'Yellow').
    """
    # Colored console output — serialized + flushed
    color_code = _ANSI_COLORS.get(foreground_color, "")
    with _stdout_lock:
        if color_code and sys.stdout.isatty():
            print(f"{color_code}{message}{_ANSI_RESET}", flush=True)
        else:
            print(message, flush=True)

    # Write to log file
    ts = _timestamp()
    log_entry = f"[{ts}] [INFO] {message}"
    _write_to_log_file(log_entry)


# ===========================================================================
# 4. setup_host_logging — PS global:Write-Host (L2245-L2268)
# ===========================================================================

def setup_bootstrap_log() -> str:
    """
    Create a bootstrap log file before the final output path is known.

    Mirrors PS L1538-1596: tries PAX_BOOTSTRAP_LOG_DIR env var first,
    falls back to system temp. Performs a write-probe (zero-byte sentinel)
    to verify the directory is writable.

    Returns:
        Path to the bootstrap log file.

    Raises:
        RuntimeError: If no writable bootstrap-log directory is available.
    """
    global _log_file, _log_file_is_bootstrap, _bootstrap_log_dir

    candidate_dirs: list[str] = []
    env_dir = os.environ.get('PAX_BOOTSTRAP_LOG_DIR', '').strip()
    if env_dir:
        candidate_dirs.append(env_dir)
    candidate_dirs.append(tempfile.gettempdir())

    chosen_dir: Optional[str] = None
    for cand in candidate_dirs:
        try:
            os.makedirs(cand, exist_ok=True)
            # Write-probe: zero-byte sentinel file
            probe = os.path.join(
                cand,
                f".pax_writeprobe_{os.getpid()}_{datetime.now().strftime('%Y%m%d%H%M%S')}{datetime.now().strftime('%f')[:3]}"
            )
            with open(probe, 'w') as f:
                f.write('')
            try:
                os.remove(probe)
            except OSError:
                pass
            chosen_dir = cand
            break
        except Exception:
            continue

    if not chosen_dir:
        raise RuntimeError("No writable bootstrap-log directory available.")

    _bootstrap_log_dir = chosen_dir
    log_name = f"PAX_bootstrap_{os.getpid()}_{datetime.now().strftime('%Y%m%d%H%M%S')}.log"
    bootstrap_path = os.path.join(chosen_dir, log_name)

    _log_file = bootstrap_path
    _log_file_is_bootstrap = True

    # Flush any buffered entries
    _flush_log_buffer()

    source = 'PAX_BOOTSTRAP_LOG_DIR' if env_dir else 'system temp'
    _write_to_log_file(
        f"[{_timestamp()}] [INFO] Bootstrap log opened at {bootstrap_path} "
        f"(PID {os.getpid()}, source: {source})."
    )

    return bootstrap_path


def migrate_bootstrap_log(final_log_path: str) -> None:
    """
    Migrate bootstrap log content to the final log file.

    Mirrors PS L16598-16616: tries Move-Item first, falls back to
    copy+delete if cross-volume.

    Args:
        final_log_path: The final log file path.
    """
    global _log_file, _log_file_is_bootstrap

    if not _log_file_is_bootstrap or not _log_file:
        return

    bootstrap_path = _log_file
    if not os.path.exists(bootstrap_path):
        _log_file = final_log_path
        _log_file_is_bootstrap = False
        return

    # Ensure parent directory exists
    os.makedirs(os.path.dirname(final_log_path) or '.', exist_ok=True)

    try:
        # Try atomic rename (same filesystem)
        os.replace(bootstrap_path, final_log_path)
    except OSError:
        # Cross-volume: copy + delete (PS: Get-Content -Raw → Set-Content -Encoding UTF8)
        try:
            with open(bootstrap_path, 'r', encoding='utf-8') as src:
                content = src.read()
            with open(final_log_path, 'w', encoding='utf-8') as dst:
                dst.write(content)
            try:
                os.remove(bootstrap_path)
            except OSError:
                pass
        except Exception as migrate_ex:
            # PS: Yellow warning "WARNING: Could not migrate bootstrap log..."
            print(
                f"WARNING: Could not migrate bootstrap log to final location "
                f"({bootstrap_path} -> {final_log_path}): {migrate_ex}"
            )

    _log_file = final_log_path
    _log_file_is_bootstrap = False


class _FlushingStreamHandler(logging.Handler):
    """Handler that routes logger output through ``print()`` — the same
    code path that ``write_log()`` uses — so both streams are treated
    identically by Jupyter / Fabric notebook kernels.

    Background: Jupyter replaces ``sys.stdout`` with a custom OutStream
    backed by a ZMQ IOPub channel.  ``print()`` is handled as a
    first-class operation and flushed immediately, but direct
    ``sys.stdout.write()`` calls (as used by stdlib StreamHandler) can
    be deferred or batched into a separate output group.  This causes
    ``write_log()`` output and ``logger.info()`` output to appear as
    two non-interleaved blocks instead of chronological order.

    By using ``print(flush=True)`` under the same ``_stdout_lock``,
    this handler's output merges seamlessly with ``write_log()`` output
    in the notebook cell, preserving chronological ordering.
    """

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            with _stdout_lock:
                print(msg, flush=True)
        except Exception:
            self.handleError(record)


class _FlushingFileHandler(logging.Handler):
    """File handler that opens, writes, flushes, fsyncs, and closes the log
    file on every single emit — the same open/write/close pattern
    ``_write_to_log_file()`` already uses for ``write_log()`` calls.

    Background: stdlib ``logging.FileHandler`` opens ONE file handle and
    keeps it open for the lifetime of the logger, calling ``stream.flush()``
    after every record. On local disk that is enough to make the write
    visible immediately. On the Fabric Lakehouse ``Files/`` mount
    (OneLake-backed), a Python-level ``flush()`` only empties the process's
    own buffer — it does not guarantee the underlying network filesystem
    pushes the bytes to OneLake until the file handle is actually **closed**.
    With a handle held open for an entire multi-hour run, every
    ``logger.info()`` / ``logger.warning()`` call from mod7/mod11/mod12 etc.
    (which carries most of the granular CopilotInteraction and Agent 365
    progress detail) can sit invisible in the Lakehouse Files pane until the
    handle finally closes — producing exactly the "log lines all appear in
    one lump only after that phase finishes" symptom.

    Opening, writing one line, flushing, fsyncing, and closing on every emit
    (serialized via ``_file_write_lock``, shared with ``_write_to_log_file()``)
    guarantees each record is fully released back to the filesystem — and
    therefore visible in the Lakehouse Files pane — before the emitting
    thread continues.
    """

    def __init__(self, filename: str, encoding: str = "utf-8") -> None:
        super().__init__()
        self._filename = filename
        self._encoding = encoding

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            with _file_write_lock:
                with open(self._filename, "a", encoding=self._encoding) as f:
                    f.write(msg + "\n")
                    f.flush()
                    os.fsync(f.fileno())
        except Exception:
            self.handleError(record)


def setup_host_logging(log_file_path: str) -> None:
    """
    Initialize the logging system with a file path.

    Matches PS global:Write-Host override behavior: once the log file path
    is known, all console output is mirrored to the file. Also flushes
    any buffered log entries from before the path was set.

    If a bootstrap log is active, migrates its content to the final path.

    Also configures Python's stdlib logging for use by other modules via
    ``logging.getLogger(__name__)``.

    Args:
        log_file_path: Absolute path to the log file.
    """
    global _log_file

    # If bootstrap log is active, migrate it to the final path
    if _log_file_is_bootstrap:
        migrate_bootstrap_log(log_file_path)
    else:
        _log_file = log_file_path
        # Flush buffered entries
        _flush_log_buffer()

    # Configure stdlib logging so other modules can use logging.getLogger().
    # Uses _FlushingFileHandler (open/write/fsync/close per emit) instead of
    # stdlib logging.FileHandler so writes are immediately visible on the
    # Fabric Lakehouse Files/ (OneLake) mount — see class docstring.
    handler = _FlushingFileHandler(log_file_path)
    handler.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s",
                                           datefmt="%Y-%m-%d %H:%M:%S"))
    _logger.addHandler(handler)

    # Console handler: uses _FlushingStreamHandler (print-based) so output
    # from logger.info() and write_log() merges in chronological order
    # inside Jupyter / Fabric notebook cells.
    stream_handler = _FlushingStreamHandler()
    stream_handler.setFormatter(logging.Formatter("%(message)s"))
    _logger.addHandler(stream_handler)

    _logger.setLevel(logging.DEBUG)

    # CRITICAL: Stop propagation to the root logger.  Fabric / Jupyter
    # runtimes install their own handler on the root logger that writes
    # to a separate IOPub output stream.  Without this, every logger.info()
    # message from child modules (mod5, mod7, mod11 …) propagates up to
    # both our handler (→ print → stdout → top block) AND the root
    # handler (→ separate stream → bottom block), causing the notebook
    # output to be split into two non-chronological sections.
    _logger.propagate = False


# ===========================================================================
# 5. get_masked_username — PS Get-MaskedUsername (L2271-L2328)
# ===========================================================================

def get_masked_username(username: Optional[str]) -> str:
    """
    Mask a username/email for secure display in logs and screenshots.

    Converts "admin@contoso.com" → "a******n@contoso.com" to prevent
    accidental credential exposure in terminal output, screenshots, or logs.

    Matches PS behavior exactly:
      - None/empty/whitespace → return as-is
      - No '@' → return as-is
      - Multiple '@' → return as-is
      - Local part ≤ 2 chars → first char + '******' + '@domain'
      - Normal → first char + '******' + last char + '@domain'

    Args:
        username: The username or email address to mask.

    Returns:
        Masked string, or original if not maskable.
    """
    if not username or not username.strip():
        return username or ""

    if "@" not in username:
        return username

    parts = username.split("@")
    if len(parts) != 2:
        return username

    local_part = parts[0]
    domain = parts[1]

    if len(local_part) <= 2:
        # PS: "$($localPart[0])******@$domain"
        return f"{local_part[0]}******@{domain}"

    first = local_part[0]
    last = local_part[-1]
    return f"{first}******{last}@{domain}"


# ===========================================================================
# 6. send_prompt_notification — PS Send-PromptNotification (L1853-L1873)
# ===========================================================================

def send_prompt_notification() -> None:
    """
    Play system beeps to alert user that a prompt requires attention.

    Matches PS: 3 short beeps at 800Hz, 1000Hz, 1200Hz.
    Falls back silently if beep is not supported.
    """
    try:
        if sys.platform == "win32":
            import winsound
            winsound.Beep(800, 200)   # 800Hz for 200ms
            time.sleep(0.1)           # PS: Start-Sleep -Milliseconds 100
            winsound.Beep(1000, 200)  # 1000Hz for 200ms
            time.sleep(0.1)           # PS: Start-Sleep -Milliseconds 100
            winsound.Beep(1200, 300)  # 1200Hz for 300ms
        else:
            # Unix: print BEL character
            print("\a", end="", flush=True)
    except Exception:
        pass


# ===========================================================================
# 7. set_progress_phase — PS Set-ProgressPhase (L13102)
# ===========================================================================

_VALID_PHASES = {"Parsing", "Query", "Explosion", "Export", "Complete"}


def set_progress_phase(phase: str, status: str = "") -> None:
    """
    Set the current progress phase.

    Matches PS Set-ProgressPhase: validates phase name, updates state,
    then calls update_progress.

    Args:
        phase: One of 'Parsing', 'Query', 'Explosion', 'Export', 'Complete'.
        status: Optional status text to display.
    """
    if phase not in _VALID_PHASES:
        raise ValueError(f"Invalid phase '{phase}'. Must be one of: {_VALID_PHASES}")

    _progress_state["phase"] = phase
    update_progress(status=status)


# ===========================================================================
# 8. update_progress — PS Update-Progress (L13103-L13181)
# ===========================================================================

def update_progress(
    status: str = "",
    batch_current: int = 0,
    batch_total: int = 0,
    batch_range_start: int = 0,
    batch_range_end: int = 0,
    batch_start_percent: int = 0,
    batch_end_percent: int = 0,
    batch_total_is_estimate: bool = False,
    metrics: Optional[dict] = None,
    parsing_label: str = "Pre-parsing JSON",
) -> None:
    """
    Update progress bar state and compute composite progress string.

    Matches PS Update-Progress: computes weighted progress across phases
    (Parsing, Query, Explosion, Export) and builds a composite status string
    with per-phase detail.

    The progress string is stored in _progress_state['last_status'] for
    retrieval by callers (e.g., tqdm wrapper, custom progress display).

    Args:
        status: Optional status prefix text.
        batch_current/batch_total: Batch-level progress within a phase.
        batch_range_start/batch_range_end: Record range for explosion phase.
        batch_start_percent/batch_end_percent: Percent range for batch display.
        batch_total_is_estimate: If True, prefix batch_total with '~'.
        metrics: Optional metrics dict with 'TotalRecordsFetched' key.
        parsing_label: Label for Parsing phase. PS uses 'Pre-parsing + Filtering'
            when agent/prompt filters are active (L11915-11917).
    """
    w = _progress_state["weights"]
    ps = _progress_state["Parsing"]
    qs = _progress_state["Query"]
    es = _progress_state["Explode"]
    xs = _progress_state["Export"]

    p_pct = (ps["current"] / ps["total"]) if ps["total"] > 0 and w.get("Parsing", 0) > 0 else 0.0
    q_pct = (qs["current"] / qs["total"]) if qs["total"] > 0 else 0.0
    e_pct = (es["current"] / es["total"]) if es["total"] > 0 and w.get("Explosion", 0) > 0 else 0.0
    x_pct = (xs["current"] / xs["total"]) if xs["total"] > 0 else 0.0

    # Zero-record weighting: emphasize Query when no records yet
    total_fetched = (metrics or {}).get("TotalRecordsFetched", 0)
    phase = _progress_state["phase"]

    if phase == "Query" and total_fetched == 0:
        w["Query"] = 1.0
        w["Explosion"] = 0.0
        w["Export"] = 0.0
        if "Parsing" in w:
            w["Parsing"] = 0.0
    elif phase == "Query" and total_fetched > 0:
        # Restore weights if they were temporarily overridden
        if _original_weights and w["Query"] == 1.0 and w["Explosion"] == 0.0 and w["Export"] == 0.0:
            for key in _original_weights:
                w[key] = _original_weights[key]

    # Build per-phase detail strings
    p_detail = ""
    if w.get("Parsing", 0) > 0 and ps["total"] > 0:
        p_detail = f"{ps['current']}/{ps['total']}({int(round(p_pct * 100))}%)"

    q_detail = ""
    if w.get("Query", 0) > 0 and qs["total"] > 0:
        q_detail = f"{qs['current']}/{qs['total']}({int(round(q_pct * 100))}%)"

    # Explosion counts with optional batch info
    batch_total_display = f"~{batch_total}" if batch_total_is_estimate else str(batch_total)

    if batch_range_start >= 1 and batch_range_end >= 1 and es["total"] > 0:
        if batch_start_percent >= 0 and batch_end_percent > 0:
            batch_info = (f" Batch: {batch_current}/{batch_total_display}"
                          f"({batch_start_percent}%-{batch_end_percent}%)") if batch_total >= 1 else ""
        else:
            batch_pct = int(round((batch_current / batch_total) * 100)) if batch_total > 0 and batch_current > 0 else 0
            batch_info = f" Batch: {batch_current}/{batch_total_display}({batch_pct}%)" if batch_total >= 1 else ""
        explosion_counts = f"Records {batch_range_start}-{batch_range_end}/{es['total']}{batch_info}"
    elif batch_total >= 1:
        batch_pct = int(round((batch_current / batch_total) * 100)) if batch_total > 0 and batch_current > 0 else 0
        batch_info = f" Batch: {batch_current}/{batch_total_display}({batch_pct}%)"
        explosion_counts = (f"Records {es['current']}/{es['total']}{batch_info}"
                            if es["total"] > 0 else "0/0")
    else:
        explosion_counts = (f"{es['current']}/{es['total']}({int(round(e_pct * 100))}%)"
                            if es["total"] > 0 else "0/0")

    e_detail = ""
    if w.get("Explosion", 0) > 0:
        if phase == "Explosion":
            e_detail = f" | {explosion_counts}"
        else:
            e_detail = f" | Explosion: {explosion_counts}"

    x_detail = (f" | Export: {xs['current']}/{xs['total']}({int(round(x_pct * 100))}%)"
                if xs["total"] > 0 else " | Export: 0/0")

    # Phase prefix — PS L11915: conditional label for Parsing phase
    phase_prefix_map = {
        "Parsing": parsing_label,
        "Query": "Query",
        "Explosion": "Explosion",
        "Export": "Export",
        "Complete": "Complete",
    }
    phase_prefix = phase_prefix_map.get(phase, phase)

    # Build composite string
    if phase == "Parsing" and p_detail:
        composite = f"{phase_prefix}: {p_detail}{e_detail}{x_detail}"
    elif phase == "Explosion" and not q_detail:
        composite = f"Explosion: {explosion_counts}{x_detail}"
    else:
        if q_detail:
            composite = f"{phase_prefix}: {q_detail}{e_detail}{x_detail}"
        else:
            composite = f"{phase_prefix}:{e_detail}{x_detail}"

    status_text = f"{status} :: {composite}" if status else composite

    # Truncate to 180 chars max (PS parity)
    if len(status_text) > 180:
        status_text = status_text[:177] + "..."

    # Store for retrieval by progress display layer
    _progress_state["last_status"] = status_text


# ===========================================================================
# 9. complete_progress — PS Complete-Progress (L13183)
# ===========================================================================

def complete_progress() -> None:
    """
    Finalize and clear the progress display.

    Matches PS Complete-Progress: placeholder for progress display compatibility.
    In Python, this can be used to close a tqdm bar or clear terminal progress.
    """
    _progress_state["phase"] = "Complete"
    _progress_state["last_status"] = "Complete"


# ===========================================================================
# PUBLIC API: Get/set progress counters
# ===========================================================================

def get_progress_state() -> dict[str, Any]:
    """Return a copy of the current progress state dict."""
    return dict(_progress_state)


def set_progress_counters(
    phase: str,
    current: int,
    total: int,
) -> None:
    """
    Update the current/total counters for a specific phase.

    Args:
        phase: One of 'Parsing', 'Query', 'Explode', 'Export'.
        current: Current count.
        total: Total count.
    """
    phase_key = phase if phase in _progress_state else None
    if phase_key and isinstance(_progress_state.get(phase_key), dict):
        _progress_state[phase_key]["current"] = current
        _progress_state[phase_key]["total"] = total


# ===========================================================================
# MEMORY / HEARTBEAT OBSERVABILITY (v1.11.15 parity: PS Write-PaxMemoryObservation)
# ===========================================================================
# Lightweight, additive heartbeat: samples this process's working set and
# managed heap, an optional child-process working set, rows/pages processed,
# and temp-storage bytes, then appends ONE diagnostic line to the run log
# (write_log_file — never echoed to console). No new required output file is
# created and it NEVER surfaces identifying row values. Emission is throttled
# to ``interval_seconds`` (default 30s) unless ``force=True``. Every metric
# read is guarded so a missing child process or inaccessible counter can
# never crash the run. Observation only — changes no success/failure result.

_pax_last_mem_obs_utc: Optional[datetime] = None
_pax_mem_obs_list: Optional[list[dict[str, Any]]] = None


def enable_memory_observation_sink() -> list[dict[str, Any]]:
    """Enable the optional in-memory metrics sink and return it (append-only
    list of observation dicts). Mirrors PS ``$script:PaxMemObsList``."""
    global _pax_mem_obs_list
    if _pax_mem_obs_list is None:
        _pax_mem_obs_list = []
    return _pax_mem_obs_list


def _process_working_set_bytes(pid: Optional[int] = None) -> Optional[int]:
    """Best-effort RSS/working-set bytes for the given pid (current process
    when omitted). Tries ``psutil`` first (accurate, cross-platform), then
    falls back to ``resource.getrusage`` (Linux/macOS, self only). Returns
    ``None`` rather than raising when unavailable."""
    try:
        import psutil  # type: ignore[import-not-found]

        proc = psutil.Process(pid) if pid is not None else psutil.Process()
        return int(proc.memory_info().rss)
    except Exception:
        pass
    if pid is None or pid == os.getpid():
        try:
            import resource

            # ru_maxrss is KB on Linux, bytes on macOS (Darwin).
            raw = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
            return int(raw * 1024) if sys.platform != "darwin" else int(raw)
        except Exception:
            return None
    return None


def _directory_size_bytes(path: str) -> Optional[int]:
    """Sum of file sizes under ``path`` (recursive). Returns ``None`` on any
    error (missing dir, permission issue) rather than raising."""
    try:
        total = 0
        for root, _dirs, files in os.walk(path):
            for name in files:
                try:
                    total += os.path.getsize(os.path.join(root, name))
                except OSError:
                    continue
        return total
    except Exception:
        return None


def write_pax_memory_observation(
    stage: str = "",
    *,
    child_pid: Optional[int] = None,
    rows_processed: Optional[int] = None,
    pages_processed: Optional[int] = None,
    temp_dir: str = "",
    force: bool = False,
    interval_seconds: int = 30,
) -> None:
    """Sample process memory + progress counters and append one MEMOBS line
    to the log file. Throttled to ``interval_seconds`` unless ``force=True``.
    Never raises — every failure path is swallowed so this is pure
    observability with zero effect on run success/failure."""
    global _pax_last_mem_obs_utc
    try:
        now_utc = datetime.utcnow()
        if not force and _pax_last_mem_obs_utc is not None:
            if (now_utc - _pax_last_mem_obs_utc).total_seconds() < interval_seconds:
                return
        _pax_last_mem_obs_utc = now_utc

        pid = os.getpid()
        parent_ws = _process_working_set_bytes()
        heap_bytes: Optional[int] = None
        try:
            import tracemalloc

            if tracemalloc.is_tracing():
                heap_bytes = tracemalloc.get_traced_memory()[0]
        except Exception:
            heap_bytes = None
        child_ws = _process_working_set_bytes(child_pid) if child_pid is not None else None
        temp_bytes = _directory_size_bytes(temp_dir) if temp_dir and os.path.isdir(temp_dir) else None

        parts: list[str] = [f"ts={now_utc.isoformat()}Z"]
        if stage:
            parts.append(f"stage={stage}")
        parts.append(f"pid={pid}")
        if parent_ws is not None:
            parts.append(f"parentWS_MB={parent_ws / (1024 * 1024):.1f}")
        if heap_bytes is not None:
            parts.append(f"heap_MB={heap_bytes / (1024 * 1024):.1f}")
        if child_pid is not None:
            parts.append(f"childPID={child_pid}")
        if child_ws is not None:
            parts.append(f"childWS_MB={child_ws / (1024 * 1024):.1f}")
        if rows_processed is not None:
            parts.append(f"rows={rows_processed}")
        if pages_processed is not None:
            parts.append(f"pages={pages_processed}")
        if temp_bytes is not None:
            parts.append(f"tempBytes={temp_bytes}")

        try:
            write_log_file("MEMOBS " + " ".join(parts))
        except Exception:
            pass

        if _pax_mem_obs_list is not None:
            try:
                _pax_mem_obs_list.append(
                    {
                        "ts": now_utc.isoformat() + "Z",
                        "stage": stage,
                        "pid": pid,
                        "parentWSBytes": parent_ws,
                        "heapBytes": heap_bytes,
                        "childPID": child_pid,
                        "childWSBytes": child_ws,
                        "rows": rows_processed,
                        "pages": pages_processed,
                        "tempBytes": temp_bytes,
                    }
                )
            except Exception:
                pass
    except Exception:
        # Observation must never affect run success/failure.
        pass


if __name__ == "__main__":
    ok = True
    errors: list[str] = []

    # get_masked_username
    if get_masked_username("admin@contoso.com") != "a******n@contoso.com":
        errors.append(f"mask failed: {get_masked_username('admin@contoso.com')}")
        ok = False

    if get_masked_username("ab@x.com") != "a******@x.com":
        errors.append(f"short mask failed: {get_masked_username('ab@x.com')}")
        ok = False

    if get_masked_username("noemail") != "noemail":
        errors.append("no-@ mask failed")
        ok = False

    if get_masked_username("") != "":
        errors.append("empty mask failed")
        ok = False

    if get_masked_username(None) != "":
        errors.append("None mask failed")
        ok = False

    # write_log / write_log_file / write_log_host
    with tempfile.NamedTemporaryFile(mode="w", suffix=".log", delete=False, encoding="utf-8") as f:
        tmp_log = f.name

    try:
        setup_host_logging(tmp_log)
        write_log("test log message")
        write_log_file("file only message")
        write_log_host("host message", "Green")

        # Close all handlers before reading/deleting
        for h in _logger.handlers[:]:
            h.close()
            _logger.removeHandler(h)

        with open(tmp_log, "r", encoding="utf-8") as f:
            content = f.read()
        if "test log message" not in content:
            errors.append("write_log not in file")
            ok = False
        if "file only message" not in content:
            errors.append("write_log_file not in file")
            ok = False
        if "host message" not in content:
            errors.append("write_log_host not in file")
            ok = False
    finally:
        os.unlink(tmp_log)

    # progress
    set_progress_phase("Query")
    set_progress_counters("Query", 5, 10)
    update_progress()
    state = get_progress_state()
    if state["phase"] != "Query":
        errors.append("phase not set")
        ok = False

    complete_progress()
    if get_progress_state()["phase"] != "Complete":
        errors.append("complete_progress failed")
        ok = False

    # send_prompt_notification (just verify no crash)
    try:
        send_prompt_notification()
    except Exception as e:
        errors.append(f"send_prompt_notification crashed: {e}")
        ok = False

    if ok:
        print("PAX Logging Module - OK (all self-tests passed)")
    else:
        print("PAX Logging Module - FAILED:")
        for e in errors:
            print(f"  - {e}")
        import sys as _sys
        _sys.exit(1)
