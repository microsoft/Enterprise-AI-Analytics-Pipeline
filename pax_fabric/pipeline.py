"""
pax_fabric.pipeline — Notebook-friendly orchestration entry point.
==================================================================
This module replaces the ``__main__`` CLI shell with a function call.
Fabric notebooks build a parameters dict and call::

    from pax_fabric import run
    result = run({
        "Auth": "AppRegistration",
        "TenantId": "...",
        "ClientId": "...",
        "ClientSecret": "...",
        "StartDate": "2026-02-01",
        "EndDate": "2026-02-02",
        "Rollup": True,
    })

The pipeline is intentionally a thin wrapper that delegates to the existing
helpers in :mod:`pax_fabric.__main__` — those helpers contain the validated
query orchestration, explosion, and CSV export logic that already matches
the legacy PowerShell behaviour byte-for-byte.

Phase A0 (CSV-to-Files bridge) differences from legacy ``__main__.main()``:
    * No ``argparse`` — input is a dict via :func:`config_from_params`.
    * No ``signal.signal(SIGINT)`` / ``atexit`` — Fabric notebooks supply
      their own lifecycle.
    * Output paths are rebased onto ``/lakehouse/default/Files/pax/...``
      via :mod:`pax_fabric.files_io`.
    * Returns a ``dict`` describing the run (output paths, counts, timing)
      instead of an exit code.
"""

from __future__ import annotations

import time
import traceback
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

from . import files_io
from .models import PAXConfig, PAXRunContext
from .mod1_pax_config import (
    SCRIPT_VERSION,
    config_from_params,
    initialize_config,
)
from .mod3_pax_logging import (
    setup_host_logging,
    set_progress_phase,
    write_log,
)
from .mod5_pax_auth import (
    connect_purview_audit,
    is_connected,
    reset_auth_state,
)
from .mod6_pax_checkpoint import (
    find_checkpoints,
    get_checkpoint_data,
    get_checkpoint_path,
    read_checkpoint,
    remove_checkpoint,
    reset_checkpoint_state,
    select_checkpoint,
    set_checkpoint_enabled,
)
from .mod10_pax_csv_export import (
    CsvWriter,
)
from .mod9_pax_data_transform import (
    convert_to_purview_exploded_records,
    convert_to_structured_record,
)
from .mod13_pax_dual_mode import disconnect_purview_audit

# Re-use the heavy orchestration helpers from __main__ unchanged so the
# Fabric pipeline matches the CLI pipeline behaviour 1:1. Importing __main__
# is side-effect-free (the signal/atexit hooks are inside ``main()``).
from .__main__ import (
    _run_query_phase,
    _run_rollup_processors,
    _export_entra_users,
    _export_entra_users_only,
    _iter_jsonl_shards,
    _cleanup_spilled_shards,
    _SPILL_BATCH_SIZE,
    EXIT_SUCCESS,
    EXIT_ERROR,
)


def _resolve_run_id(cfg: PAXConfig) -> str:
    """Return ``cfg.run_id`` if explicitly set, otherwise fall back to the
    same UTC timestamp the legacy pipeline uses for output filenames.

    Uses ``getattr`` so the pipeline keeps working even against a stale
    ``PAXConfig`` snapshot that pre-dates the ``run_id`` field declaration
    (e.g. an older ``pax_fabric`` package still present on the lakehouse,
    or a cached ``__pycache__``).
    """
    existing = getattr(cfg, "run_id", None)
    if existing:
        return existing
    new_id = (
        getattr(cfg, "script_run_timestamp", None)
        or datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    )
    try:
        cfg.run_id = new_id  # populate so downstream callers can read it back
    except AttributeError:
        # Stale dataclass with __slots__ or otherwise immutable layout —
        # the run_id stays local; downstream code reads it via this return.
        pass
    return new_id


def _resolve_csv_output_root(cfg: PAXConfig, run_id: str, *,
                              scratch: bool = False) -> str:
    """Resolve the destination directory for the per-run CSV bundle and
    propagate it onto ``cfg.output_path`` so every downstream helper (mod10,
    mod6, mod12, processors) inherits the lakehouse-aware location without
    any further patching.

    When ``scratch=True`` (Phase B / output_mode='delta'), the CSV bundle is
    written to a transient ``Files/pax/_scratch/<run_id>/`` directory that
    the caller is expected to delete after draining to Delta.
    """
    if cfg.csv_output_root:
        target = cfg.csv_output_root
    elif scratch:
        target = files_io.scratch_root(run_id)
    else:
        target = files_io.csv_root(run_id)
    Path(target).mkdir(parents=True, exist_ok=True)
    cfg.csv_output_root = target
    cfg.output_path = target  # Force mod10/mod6/mod12 to use the lakehouse dir.
    cfg._output_path_explicit = True

    # v1.11.3 validator (mod1_pax_config.validate_config) requires EXACTLY ONE
    # of (OutputPath* | Append*) PER STREAM, but only when that stream is in
    # scope; supplying a destination for an out-of-scope stream is also an
    # error. In Fabric CSV-bundle mode the per-run directory IS the
    # destination for every in-scope stream, so we bind each per-stream
    # OutputPath* attribute here, but ONLY when the corresponding stream is
    # actually requested. Defensive setattr / hasattr keeps older v1.11.1-style
    # PAXConfig snapshots (without these fields) working unchanged.
    _stream_bindings = (
        # (attr to bind, scope predicate)
        ("output_path_user_info",
         getattr(cfg, "include_user_info", False)
         or getattr(cfg, "only_user_info", False)),
        ("output_path_agent365_info",
         getattr(cfg, "include_agent365_info", False)
         or getattr(cfg, "only_agent365_info", False)),
    )
    for attr, in_scope in _stream_bindings:
        if not in_scope:
            continue
        if hasattr(cfg, attr) and getattr(cfg, attr, None) is None:
            try:
                setattr(cfg, attr, target)
            except AttributeError:
                pass
    return target


def _build_output_filename(cfg: PAXConfig) -> str:
    """Replicate ``__main__._resolve_output_path``'s filename rule exactly
    (the directory has already been rebased to the lakehouse)."""
    import re as _re
    timestamp = cfg.script_run_timestamp or datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    activity_types = cfg.activity_types or ["CopilotInteraction"]
    if cfg.combine_output:
        if len(activity_types) == 1:
            safe = _re.sub(r'[\\/:*?"<>|]', "_", activity_types[0])
            return f"Purview_Audit_UsageActivity_{safe}_{timestamp}.csv"
        return f"Purview_Audit_UsageActivity_CombinedActivityTypes_{timestamp}.csv"
    return f"Purview_Audit_{timestamp}.csv"


def _iter_delta_as_dicts(uri: str, storage_options: dict | None):
    """Yield one dict[str, str] per row from a Delta table, streaming batch by batch.

    Peak RAM is bounded to one Arrow batch (~65K rows) plus the yielded dict.
    All values are cast to strings to match csv.DictReader output.
    """
    from deltalake import DeltaTable
    dt = DeltaTable(uri, storage_options=storage_options)
    for batch in dt.to_pyarrow_dataset().to_batches():
        columns = batch.schema.names
        col_arrays = {c: batch.column(c).to_pylist() for c in columns}
        for i in range(batch.num_rows):
            row = {}
            for c in columns:
                v = col_arrays[c][i]
                if v is None:
                    row[c] = ""
                elif isinstance(v, str):
                    row[c] = v
                elif isinstance(v, float):
                    # Avoid "5.0" → int("5.0") ValueError in downstream parsers
                    row[c] = str(int(v)) if v == int(v) else str(v)
                else:
                    row[c] = str(v)
            yield row


def _recompute_userstats_from_delta(
    delta_results: list[dict],
    schema: str,
    log_fn=None,
) -> list[dict]:
    """Recompute UserStats + SessionCohort from the accumulated Rollup Delta table.

    After the per-run CSV drain (step 7), the Rollup and SessionStats Delta
    tables contain the full accumulated history. This function streams them
    batch-by-batch into write_userstats_files() via the aggregated_rows /
    session_stats_rows parameters — no temp CSV files, no full-table PyArrow
    load into RAM.

    Only called when ``output_mode='delta'`` AND ``include_m365_usage=True``.
    Returns a list of result dicts for the recomputed tables.
    """
    import tempfile as _tempfile

    from . import files_io as _fio
    from . import mod16_pax_delta as _mod16
    from .processors.m365_processor import write_userstats_files

    def _log(msg: str, level: str = "INFO") -> None:
        if log_fn:
            log_fn(msg, level)

    # --- Locate the Rollup and SessionStats Delta tables from drain results ---
    rollup_info = None
    session_stats_info = None
    userstats_info = None
    session_cohort_info = None
    for entry in delta_results:
        tname = entry.get("table", "")
        if "UserStats" not in tname and "SessionCohort" not in tname and "SessionStats" not in tname:
            if "_Rollup" in tname:
                rollup_info = entry
        if "SessionStats" in tname:
            session_stats_info = entry
        if "UserStats" in tname:
            userstats_info = entry
        if "SessionCohort" in tname:
            session_cohort_info = entry

    if not rollup_info:
        _log("Recompute: No Rollup Delta table found in drain results — skipping.", "WARN")
        return []

    rollup_uri = rollup_info["path"]
    _log(f"Recompute: Streaming from accumulated Rollup Delta: {rollup_info['table']}")

    try:
        import deltalake  # noqa: F401
    except ImportError:
        _log("Recompute: deltalake not installed — skipping.", "WARN")
        return []

    storage_options = _fio.onelake_storage_options()

    # Build streaming iterators
    rollup_iter = _iter_delta_as_dicts(rollup_uri, storage_options)
    ss_iter = None
    if session_stats_info:
        _log(f"Recompute: Will stream SessionStats Delta: {session_stats_info['table']}")
        ss_iter = _iter_delta_as_dicts(session_stats_info["path"], storage_options)

    # Write output to a temp directory (output CSVs only — inputs are streamed)
    results: list[dict] = []
    with _tempfile.TemporaryDirectory(prefix="pax_recompute_") as tmpdir:
        userstats_csv = str(Path(tmpdir) / "userstats.csv")
        session_cohort_csv = str(Path(tmpdir) / "session_cohort.csv")

        _log("Recompute: Running write_userstats_files() on full accumulated data...")
        try:
            user_count, cohort_count = write_userstats_files(
                aggregated_csv_path="<unused>",
                userstats_csv_path=userstats_csv,
                session_csv_path=session_cohort_csv,
                quiet=False,
                aggregated_rows=rollup_iter,
                session_stats_rows=ss_iter,
            )
        except Exception as exc:
            _log(f"Recompute: write_userstats_files failed: {exc}", "ERROR")
            return []

        _log(f"Recompute: UserStats={user_count:,} users, SessionCohort={cohort_count:,} pairs")

        # --- Overwrite the UserStats and SessionCohort Delta tables ---
        for csv_path, info in [
            (userstats_csv, userstats_info),
            (session_cohort_csv, session_cohort_info),
        ]:
            if not info or not Path(csv_path).is_file():
                continue
            target_uri = info["path"]
            table_name = info["table"]
            _log(f"Recompute: Overwriting Delta table {table_name}")
            try:
                res = _mod16.convert_csv_to_delta(
                    input_csv=csv_path,
                    target_uri=target_uri,
                    mode="overwrite",
                    storage_options=storage_options,
                )
                rows = res.get("rows_written", 0)
                _log(f"Recompute: {table_name} overwritten — {rows:,} rows")
                results.append({
                    "csv": Path(csv_path).name,
                    "table": table_name,
                    "path": target_uri,
                    "rows_written": rows,
                    "is_init": False,
                    "added_cols": [],
                    "recomputed": True,
                })
            except Exception as exc:
                _log(f"Recompute: Failed to write {table_name}: {exc}", "WARN")

    return results


def run(params: Optional[dict] = None) -> dict:
    """Execute the full PAX pipeline against the provided parameter dict.

    Parameters dict supports one Phase-B-specific key in addition to all
    legacy options:

        ``OutputMode`` : {'csv', 'delta'}, default 'csv'
            ``'csv'``   — legacy behaviour: persistent CSV bundle under
                          ``Files/pax/csv/<run_id>/``.
            ``'delta'`` — Phase B: CSVs are written to a transient scratch
                          directory, then drained into Delta tables under
                          ``Tables/<TargetSchema>/`` and the scratch dir is
                          deleted.

        ``TargetSchema`` : str, default 'dbo'
            Lakehouse schema for Delta output (Phase B only).

        ``KeepScratch`` : bool, default False
            When True (Phase B only), the transient scratch CSV directory
            is preserved for debugging.

    Returns a dict with::

        {
          "success": bool,
          "exit_code": int,
          "run_id": str,
          "output_mode": str,
          "csv_output_root": str,
          "output_file": str | None,
          "log_file": str,
          "records_fetched": int,
          "output_rows": int,
          "delta_tables": list[dict],  # Phase B only
          "target_schema": str | None, # Phase B only
          "elapsed_seconds": float,
          "error": str | None,    # traceback summary on failure
        }
    """
    params = params or {}
    output_mode = str(params.get("OutputMode", "csv")).lower()
    if output_mode not in {"csv", "delta"}:
        raise ValueError(
            f"OutputMode must be 'csv' or 'delta', got {output_mode!r}"
        )
    target_schema = str(params.get("TargetSchema", "dbo")).strip()
    if not target_schema:
        target_schema = "dbo"
    # Validate schema name: must be a simple identifier (letters, digits,
    # underscores). Reject path separators, dots, spaces, etc. to prevent
    # path-traversal or invalid Delta paths.
    import re as _re
    if not _re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", target_schema):
        raise ValueError(
            f"TargetSchema must be a valid identifier "
            f"(letters, digits, underscores, cannot start with a digit), "
            f"got {target_schema!r}"
        )
    keep_scratch = bool(params.get("KeepScratch", False))
    name_overrides = params.get("TableNameOverrides") or {}

    result: dict[str, Any] = {
        "success": False,
        "exit_code": EXIT_ERROR,
        "run_id": None,
        "output_mode": output_mode,
        "csv_output_root": None,
        "output_file": None,
        "log_file": None,
        "records_fetched": 0,
        "output_rows": 0,
        "delta_tables": [],
        "target_schema": target_schema if output_mode == "delta" else None,
        "elapsed_seconds": 0.0,
        "error": None,
    }

    ctx = PAXRunContext()
    ctx.config = config_from_params(params)
    config = ctx.config

    # ------------------------------------------------------------------
    # 0a. Reset module-level state from any prior run in this session.
    # ------------------------------------------------------------------
    # Fabric notebooks keep Python modules loaded across cell executions,
    # so globals like _is_resume_mode persist from earlier runs. Without
    # this reset a fresh run after a resume would skip checkpoint init
    # (is_resume_mode() still returns True) and lose the "If interrupted,
    # resume with ..." log line.
    reset_checkpoint_state()
    reset_auth_state()

    # ------------------------------------------------------------------
    # 0. Per-run identifiers + lakehouse output directories.
    # ------------------------------------------------------------------
    run_id = _resolve_run_id(config)
    csv_root = _resolve_csv_output_root(
        config, run_id, scratch=(output_mode == "delta")
    )
    # Remember the initial scratch dir so we can clean it up if resume
    # re-resolves csv_root to a different (prior-run) directory.
    _initial_csv_root = csv_root
    result["run_id"] = run_id
    result["csv_output_root"] = csv_root

    # ------------------------------------------------------------------
    # 1. Logging — redirect to Files/pax/logs/<run_id>.log.
    # ------------------------------------------------------------------
    log_file = str(Path(files_io.logs_root()) / f"{run_id}.log")
    setup_host_logging(log_file)
    ctx.log_file = log_file
    result["log_file"] = log_file
    write_log(f"PAX Fabric pipeline v{SCRIPT_VERSION}  run_id={run_id}")
    write_log(f"OutputMode:      {output_mode}")
    if output_mode == "delta":
        write_log(f"TargetSchema:    {target_schema}")
    write_log(f"CSV output root: {csv_root}")
    write_log(f"Log file:        {log_file}")

    start_wall = time.perf_counter()

    try:
        # --------------------------------------------------------------
        # 2. Validate + populate computed config fields.
        # --------------------------------------------------------------
        set_progress_phase("Parsing")
        errors = initialize_config(config)
        if errors:
            for err in errors:
                write_log(err, level="ERROR")
            result["error"] = "; ".join(errors)
            return result

        write_log(f"Date range: {config.start_date} -> {config.end_date}")

        # --- Checkpoint self-gate (PS L17575) ---
        # Checkpointing disabled for replay-from-CSV and OnlyUserInfo modes.
        set_checkpoint_enabled(
            not getattr(config, "raw_input_csv", None)
            and not getattr(config, "only_user_info", False)
        )

        # --------------------------------------------------------------
        # 3. Authentication (skip in replay-from-CSV mode).
        # --------------------------------------------------------------
        if not config.raw_input_csv:
            auth_result = connect_purview_audit(
                auth_method=config.auth,
                tenant_id=config.tenant_id,
                client_id=config.client_id,
                client_secret=config.client_secret,
                scopes=None,
                http_client=None,
                remote_output_mode=config.remote_output_mode,
                include_agent365=getattr(config, "include_agent365_info", False),
            )
            ctx.auth_token = auth_result.get("token")
            ctx.auth_expires_on = auth_result.get("expires_on")
            ctx.auth_method = config.auth

        # --------------------------------------------------------------
        # 3b. Resume / Checkpoint Recovery.
        # --------------------------------------------------------------
        if config.resume is not None:
            set_progress_phase("Parsing")  # Resume is part of the parsing phase
            resume_path = config.resume  # '' = auto-discover, 'path' = explicit

            if resume_path and resume_path.strip():
                # Explicit checkpoint path provided
                if read_checkpoint(checkpoint_path=resume_path, running_script_version=SCRIPT_VERSION):
                    ctx.checkpoint_path = resume_path
                    write_log(f"Resumed from checkpoint: {resume_path}")
                else:
                    write_log(
                        f"Failed to load checkpoint: {resume_path}",
                        level="ERROR",
                    )
                    result["error"] = f"Failed to load checkpoint: {resume_path}"
                    return result
            else:
                # Auto-discover checkpoint in output directory.
                # In Fabric, each run gets its own subdirectory under
                # _scratch/<run_id>/ (delta mode) or csv/<run_id>/ (csv mode).
                # The prior run's checkpoint lives in a DIFFERENT subdirectory
                # than the current run_id, so we must search the parent
                # (_scratch/ or csv/) to find it.
                import os as _os
                search_parent = str(Path(csv_root).parent)
                all_checkpoints: list[dict] = []
                if _os.path.isdir(search_parent):
                    for entry in _os.listdir(search_parent):
                        sub = _os.path.join(search_parent, entry)
                        if _os.path.isdir(sub):
                            found = find_checkpoints(sub)
                            all_checkpoints.extend(found)
                if not all_checkpoints:
                    write_log(
                        f"No checkpoint found under {search_parent} to resume from.",
                        level="ERROR",
                    )
                    result["error"] = "No checkpoint found to resume from."
                    return result
                # Non-interactive notebook: pick the most recent checkpoint
                selected = select_checkpoint(all_checkpoints)
                if selected:
                    cp_path = (
                        selected.get("Path", "")
                        if isinstance(selected, dict)
                        else str(selected)
                    )
                    if read_checkpoint(checkpoint_path=cp_path, running_script_version=SCRIPT_VERSION):
                        ctx.checkpoint_path = cp_path
                        write_log(f"Resumed from checkpoint: {cp_path}")
                    else:
                        write_log(
                            f"Failed to load checkpoint: {cp_path}",
                            level="ERROR",
                        )
                        result["error"] = f"Failed to load checkpoint: {cp_path}"
                        return result
                else:
                    write_log("No checkpoint selected.", level="ERROR")
                    result["error"] = "No checkpoint selected for resume."
                    return result

            # Restore ALL parameters from checkpoint so the resumed run
            # matches the original run exactly (PS L20704-20790).
            cp_data = get_checkpoint_data()
            if cp_data:
                cp_params = cp_data.get("parameters", {})

                # Restore original run timestamp (PS L20710-20712)
                # Critical: incremental shard filenames embed this timestamp;
                # using a new one would prevent prior-run shards from being found.
                if cp_data.get("runTimestamp"):
                    restored_ts = cp_data["runTimestamp"]
                    config.script_run_timestamp = restored_ts
                    try:
                        config.run_id = restored_ts
                    except AttributeError:
                        pass
                    # Update pipeline-level run_id to match
                    run_id = restored_ts
                    result["run_id"] = run_id
                    write_log(
                        f"  [RESUME] Restored original run timestamp: {restored_ts}"
                    )

                # Restore date range (PS L20720-20734)
                if cp_params.get("startDate"):
                    restored_start = cp_params["startDate"]
                    restored_end = cp_params.get("endDate", "")
                    write_log(
                        f"  [RESUME] Restoring date range from checkpoint: "
                        f"{restored_start} -> {restored_end}"
                    )
                    config.start_date = restored_start[:10]
                    if restored_end:
                        config.end_date = restored_end[:10]

                # Restore auth method and identifiers (PS restores from checkpoint;
                # only ClientSecret must be supplied by the user since it's never stored)
                if cp_params.get("auth"):
                    config.auth = cp_params["auth"]
                if cp_params.get("tenantId"):
                    config.tenant_id = cp_params["tenantId"]
                if cp_params.get("clientId"):
                    config.client_id = cp_params["clientId"]

                # Restore activity/record filtering (PS L20736-20742)
                if cp_params.get("activityTypes"):
                    config.activity_types = list(cp_params["activityTypes"])
                if cp_params.get("recordTypes"):
                    config.record_types = list(cp_params["recordTypes"])
                if cp_params.get("serviceTypes"):
                    config.service_types = list(cp_params["serviceTypes"])
                if cp_params.get("userIds"):
                    config.user_ids = list(cp_params["userIds"])
                if cp_params.get("groupNames"):
                    config.group_names = list(cp_params["groupNames"])

                # Restore agent filtering (PS L20744-20748)
                if cp_params.get("agentId"):
                    config.agent_id = list(cp_params["agentId"])
                if cp_params.get("agentsOnly"):
                    config.agents_only = True
                if cp_params.get("excludeAgents"):
                    config.exclude_agents = True

                # Restore prompt filtering (PS L20750)
                if cp_params.get("promptFilter"):
                    config.prompt_filter = cp_params["promptFilter"]

                # Restore schema/explosion settings (PS L20752-20758)
                if cp_params.get("explodeArrays"):
                    config.explode_arrays = True
                if cp_params.get("explodeDeep"):
                    config.explode_deep = True
                if cp_params.get("flatDepth"):
                    config.flat_depth = int(cp_params["flatDepth"])
                if cp_params.get("explosionThreads"):
                    config.explosion_threads = int(cp_params["explosionThreads"])

                # Restore M365/UserInfo bundles (PS L20760-20770)
                if cp_params.get("includeM365Usage"):
                    config.include_m365_usage = True
                if cp_params.get("includeUserInfo"):
                    config.include_user_info = True
                if cp_params.get("includeCopilotInteraction"):
                    config.include_copilot_interaction = True
                if cp_params.get("excludeCopilotInteraction"):
                    config.exclude_copilot_interaction = True
                if cp_params.get("includeAgent365Info"):
                    config.include_agent365_info = True
                if cp_params.get("onlyAgent365Info"):
                    config.only_agent365_info = True

                # Restore partitioning (PS L20772-20776)
                if cp_params.get("blockHours"):
                    config.block_hours = float(cp_params["blockHours"])
                if cp_params.get("partitionHours"):
                    config.partition_hours = int(cp_params["partitionHours"])
                if cp_params.get("maxPartitions"):
                    config.max_partitions = int(cp_params["maxPartitions"])

                # Restore query settings (PS L20778-20782)
                if cp_params.get("resultSize"):
                    config.result_size = int(cp_params["resultSize"])
                if cp_params.get("maxConcurrency"):
                    config.max_concurrency = int(cp_params["maxConcurrency"])
                if cp_params.get("combineOutput"):
                    config.combine_output = True

                # Restore rollup mode (PS L20784-20790)
                rollup_mode = cp_params.get("rollupMode", "None")
                if rollup_mode == "Rollup":
                    config.rollup = True
                elif rollup_mode == "RollupPlusRaw":
                    config.rollup_plus_raw = True

                # Restore remaining settings (PS L20792+)
                if cp_params.get("useEOM"):
                    config.use_eom = True
                if cp_params.get("autoCompleteness"):
                    config.auto_completeness = True
                if cp_params.get("includeTelemetry"):
                    config.include_telemetry = True
                if cp_params.get("appendFile"):
                    config.append_file = cp_params["appendFile"]

                write_log("  [RESUME] All parameters restored from checkpoint")

                # Re-resolve output directory with the restored run_id
                # so we write to the same directory that has the .pax_incremental shards.
                # Clear the cached csv_output_root first — it was set in Phase 0
                # to the NEW run's directory, but we need the PRIOR run's directory.
                config.csv_output_root = None
                csv_root = _resolve_csv_output_root(
                    config, run_id, scratch=(output_mode == "delta")
                )
                result["csv_output_root"] = csv_root
                write_log(f"  [RESUME] CSV output root: {csv_root}")

                # Re-run config validation so trim boundaries are recalculated
                errors = initialize_config(config)
                if errors:
                    for err in errors:
                        write_log(err, level="ERROR")
                    result["error"] = "; ".join(errors)
                    return result

        # --------------------------------------------------------------
        # 4. Query orchestration.
        # --------------------------------------------------------------
        only_user_info = getattr(config, "only_user_info", False)
        only_agent365 = getattr(config, "only_agent365_info", False)
        if not only_user_info and not only_agent365:
            set_progress_phase("Query")
            ctx.metrics.query_ms = _run_query_phase(ctx)
        result["records_fetched"] = ctx.metrics.total_records_fetched

        # --------------------------------------------------------------
        # 5. Post-query: dedupe + trim + explosion + CSV export (STREAMING).
        # --------------------------------------------------------------
        # Records were spilled to JSONL shards by _run_query_phase to avoid
        # OOM on large pulls. Stream them back through dedup -> date-trim ->
        # structuring/explosion -> CSV writer in fixed-size batches.
        spilled_shards: list[str] = getattr(ctx, "spilled_shards", [])
        if spilled_shards:
            set_progress_phase("Explosion")
            start_explosion = time.perf_counter_ns()

            write_log(
                f"Streaming {len(spilled_shards)} JSONL shard(s) through "
                f"dedup \u2192 trim \u2192 transform \u2192 CSV ..."
            )

            trim_start = config.trim_start_date_utc
            trim_end = config.trim_end_date_utc
            do_trim = trim_start is not None or trim_end is not None
            if do_trim:
                from .mod2_pax_data_helpers import parse_date_safe

            enable_explosion = getattr(config, "explode_arrays", False)
            enable_deep = getattr(config, "explode_deep", False)
            prompt_filter_value = getattr(config, "prompt_filter", None)

            out_filename = _build_output_filename(config)
            output_path = str(Path(csv_root) / out_filename)
            ctx.output_file = output_path
            result["output_file"] = output_path

            seen_ids: set[str] = set()
            dup_skipped = 0
            trim_skipped = 0
            rows_written = 0
            writer: CsvWriter | None = None
            csv_columns: list[str] = []
            pending_rows: list[dict[str, Any]] = []

            def _flush_pending() -> None:
                nonlocal writer, csv_columns, rows_written
                if not pending_rows:
                    return
                if writer is None:
                    csv_columns = list(pending_rows[0].keys())
                    writer = CsvWriter(path=output_path, columns=csv_columns)
                    write_log(
                        f"  [CSV] Opened writer at {Path(output_path).name} "
                        f"with {len(csv_columns)} columns"
                    )
                writer.write_rows(pending_rows)
                rows_written += len(pending_rows)
                pending_rows.clear()

            for record in _iter_jsonl_shards(spilled_shards):
                rid = (
                    record.get("Identity")
                    or record.get("Id")
                    or record.get("RecordId", "")
                )
                if rid:
                    if rid in seen_ids:
                        dup_skipped += 1
                        continue
                    seen_ids.add(rid)

                if do_trim:
                    cd = parse_date_safe(record.get("CreationDate"))
                    if cd is not None:
                        if trim_start and cd < trim_start:
                            trim_skipped += 1
                            continue
                        if trim_end and cd >= trim_end:
                            trim_skipped += 1
                            continue

                try:
                    if enable_explosion or enable_deep:
                        rows = convert_to_purview_exploded_records(
                            record=record,
                            deep=enable_deep,
                            partial_explode=enable_explosion,
                            prompt_filter_value=prompt_filter_value,
                        )
                    else:
                        rows = convert_to_structured_record(
                            record=record,
                            enable_explosion=False,
                            explode_deep=False,
                        )
                except Exception as _ex:
                    write_log(f"Transform error for record: {_ex}", level="WARN")
                    continue

                for r in rows:
                    pending_rows.append(r)
                    if len(pending_rows) >= _SPILL_BATCH_SIZE:
                        _flush_pending()

            _flush_pending()
            if writer is not None:
                writer.close()

            if dup_skipped:
                ctx.metrics.total_records_fetched -= dup_skipped
                write_log(f"Dedup: removed {dup_skipped} duplicate record(s)")
            if trim_skipped:
                write_log(
                    f"Date-range trim: removed {trim_skipped} record(s) "
                    f"outside requested date boundaries"
                )

            ctx.metrics.explosion_ms = (
                time.perf_counter_ns() - start_explosion
            ) // 1_000_000

            set_progress_phase("Export")
            ctx.metrics.export_ms = 0
            result["output_rows"] = rows_written
            write_log(f"Exported {rows_written} rows to {output_path}")

            if getattr(config, "export_workbook", False):
                write_log(
                    "Excel export with streaming mode is not supported "
                    "(ExportWorkbook is deprecated). Use the CSV output instead.",
                    level="WARN",
                )

        # --------------------------------------------------------------
        # 6. Post-processing: Entra users + rollup.
        # --------------------------------------------------------------
        if only_user_info:
            set_progress_phase("Export")
            _export_entra_users_only(ctx)
            result["output_file"] = ctx.output_file
        elif getattr(config, "include_user_info", False):
            _export_entra_users(ctx)

        if getattr(config, "rollup", False) or getattr(config, "rollup_plus_raw", False):
            _run_rollup_processors(ctx)

        # --------------------------------------------------------------
        # 7. Phase B drain: scratch CSVs -> Delta tables (output_mode='delta').
        # --------------------------------------------------------------
        if output_mode == "delta":
            from . import delta_writer
            set_progress_phase("Export", status="Delta drain")
            write_log(
                f"Draining {csv_root} -> Tables/{target_schema}/ "
                f"(run_id={run_id})"
            )
            delta_results = delta_writer.csv_dir_to_delta(
                csv_dir=csv_root,
                schema=target_schema,
                run_id=run_id,
                write_mode="append",
                name_overrides=name_overrides,
                log_fn=lambda msg, lvl="INFO": write_log(msg, level=lvl),
            )
            result["delta_tables"] = delta_results
            write_log(
                f"Delta drain complete: {len(delta_results)} table(s) written."
            )

            # ----------------------------------------------------------
            # 7b. Recompute UserStats + SessionCohort from accumulated
            #     Rollup Delta (M365 usage mode only).
            #
            #     The per-run CSV-derived UserStats/SessionCohort were
            #     already drained in step 7 (overwrite strategy), but
            #     they only reflect the current run's data. Recomputing
            #     from the full accumulated Rollup Delta ensures the
            #     percentiles, tiers, and cohort buckets cover the
            #     entire history — not just the latest run.
            # ----------------------------------------------------------
            if getattr(config, "include_m365_usage", False):
                set_progress_phase("Export", status="Recompute UserStats")
                write_log(
                    "Recomputing UserStats/SessionCohort from accumulated "
                    "Rollup Delta..."
                )
                recomputed = _recompute_userstats_from_delta(
                    delta_results=delta_results,
                    schema=target_schema,
                    log_fn=lambda msg, lvl="INFO": write_log(msg, level=lvl),
                )
                if recomputed:
                    # Update delta_tables with recomputed entries
                    recomputed_tables = {r["table"] for r in recomputed}
                    result["delta_tables"] = [
                        dt for dt in result["delta_tables"]
                        if dt["table"] not in recomputed_tables
                    ] + recomputed
                    write_log(
                        f"Recompute complete: {len(recomputed)} table(s) refreshed."
                    )

        # --- Cleanup OOM-spill JSONL shards on successful completion ---
        _spilled = getattr(ctx, "spilled_shards", [])
        _incremental_dir_str = getattr(ctx, "incremental_dir", None)
        if _spilled:
            _cleanup_spilled_shards(
                _spilled,
                Path(_incremental_dir_str) if _incremental_dir_str else None,
            )

        ctx.script_completed = True
        result["success"] = True
        result["exit_code"] = EXIT_SUCCESS

    except KeyboardInterrupt:
        # Notebook cell cancelled / session stopped by user.
        # KeyboardInterrupt is NOT a subclass of Exception — needs its own handler.
        # The checkpoint file (if initialized) persists progress to the last
        # completed partition. show_checkpoint_exit_message uses logger.info
        # which is invisible in Fabric notebooks, so we build the banner
        # directly with write_log here.
        cp_path = get_checkpoint_path()
        cp_data = get_checkpoint_data()

        write_log("")
        write_log("=" * 80)
        write_log("  Script Interrupted — Performing Graceful Cleanup")
        write_log("=" * 80)

        # Best-effort auth disconnect inside the interrupt handler itself
        # (the finally block will also try, but may not run if the kernel dies)
        try:
            if is_connected():
                disconnect_purview_audit(
                    get_context_fn=None,
                    disconnect_fn=reset_auth_state,
                    log_fn=lambda msg, lvl="INFO": write_log(msg, level=lvl),
                )
                write_log("  Microsoft Graph disconnected")
        except Exception:
            pass

        # PS-style PROGRESS SAVED banner
        write_log("")
        bar = "\u2550" * 80  # ═
        write_log(bar)
        write_log("  PROGRESS SAVED")
        write_log(bar)
        write_log("")
        if cp_path:
            write_log(f"  Checkpoint:   {cp_path}")
        if cp_data:
            stats = cp_data.get("statistics", {})
            parts = cp_data.get("partitions", {})
            completed = stats.get("partitionsComplete", 0)
            query_created = stats.get("partitionsQueryCreated", 0)
            total = parts.get("total", 0)
            remaining = total - completed - query_created
            records_saved = stats.get("totalRecordsSaved", 0)

            write_log(
                "  Partial data: (incremental .jsonl shards under "
                "scratch .pax_incremental/)"
            )
            write_log(f"  Records saved: {records_saved:,}")

            part_line = f"  Partitions: {completed}/{total} complete"
            if query_created > 0:
                part_line += f", {query_created} queries pending"
            if remaining > 0:
                part_line += f", {remaining} not started"
            write_log(part_line)
        else:
            write_log(
                "  No checkpoint was created (interrupted before query phase)."
            )

        write_log("")
        write_log("  To resume later:")
        if cp_path:
            write_log(
                f'    Set  Resume = "{cp_path}"'
            )
            write_log(
                "    in the parameters cell, fill in ClientSecret, and re-run."
            )
        else:
            write_log("    Re-run the notebook for a fresh start.")
        write_log("")
        write_log(bar)
        write_log("")
        write_log("  Cleanup complete. Exiting...")

        result["error"] = "Session interrupted by user."
        result["exit_code"] = EXIT_ERROR

    except Exception as exc:
        write_log(f"Fatal error: {exc}", level="ERROR")
        tb = traceback.format_exc()
        write_log(tb, level="ERROR")
        result["error"] = f"{exc}\n{tb}"
        result["exit_code"] = EXIT_ERROR
        # Show checkpoint resume hint on failure (use write_log, not
        # show_checkpoint_exit_message which uses logger.info — invisible
        # in Fabric notebooks without the logging bridge from __main__).
        cp_path = get_checkpoint_path()
        cp_data_err = get_checkpoint_data()
        if cp_path and cp_data_err:
            stats = cp_data_err.get("statistics", {})
            parts = cp_data_err.get("partitions", {})
            completed = stats.get("partitionsComplete", 0)
            total = parts.get("total", 0)
            records_saved = stats.get("totalRecordsSaved", 0)
            write_log("")
            write_log("=" * 80)
            write_log("  PROGRESS SAVED (run failed — checkpoint preserved)")
            write_log("=" * 80)
            write_log(f"  Checkpoint:    {cp_path}")
            write_log(f"  Records saved: {records_saved:,}")
            write_log(f"  Partitions:    {completed}/{total} complete")
            write_log("")
            write_log("  To resume:")
            write_log(f'    Set  Resume = "{cp_path}"')
            write_log(
                "    in the parameters cell, fill in ClientSecret, and re-run."
            )
            write_log("=" * 80)
    finally:
        # No atexit / no signal handler — Fabric owns the lifecycle.
        # Best-effort cleanup so a re-run from the same notebook starts fresh.
        try:
            if is_connected():
                disconnect_purview_audit(
                    get_context_fn=None,
                    disconnect_fn=reset_auth_state,
                    log_fn=lambda msg, lvl="INFO": write_log(msg, level=lvl),
                )
        except Exception:
            pass
        if ctx.script_completed:
            try:
                remove_checkpoint()
            except Exception:
                pass
        # Phase B: clean up scratch CSV dir (only after a successful drain).
        if (
            output_mode == "delta"
            and result["success"]
            and not keep_scratch
            and csv_root
            and "_scratch" in csv_root
        ):
            try:
                import shutil as _shutil
                _shutil.rmtree(csv_root, ignore_errors=True)
                write_log(f"Scratch directory removed: {csv_root}")
                # On resume, Phase 0 created a new empty scratch dir that
                # differs from the restored csv_root. Clean it up too.
                if _initial_csv_root != csv_root and "_scratch" in _initial_csv_root:
                    _shutil.rmtree(_initial_csv_root, ignore_errors=True)
            except Exception as _cleanup_exc:
                write_log(
                    f"Scratch cleanup failed (non-fatal): {_cleanup_exc}",
                    level="WARN",
                )
        result["elapsed_seconds"] = round(time.perf_counter() - start_wall, 2)

        # Surface auth-induced data loss in the summary line so a clean
        # `success=True` cannot hide silent partition failures.
        m = ctx.metrics
        auth_fail = getattr(m, "auth_failures_total", 0)
        loss_parts = getattr(m, "partitions_with_data_loss", 0)
        salvaged = getattr(m, "records_salvaged_after_auth", 0)
        if auth_fail or loss_parts:
            result["data_loss_detected"] = True
            result["auth_failures_total"] = auth_fail
            result["partitions_with_data_loss"] = loss_parts
            result["records_salvaged_after_auth"] = salvaged
            result["data_loss_events"] = list(getattr(m, "data_loss_events", []))

        loss_suffix = ""
        if auth_fail or loss_parts:
            loss_suffix = (
                f" auth_failures={auth_fail} data_loss_partitions={loss_parts} "
                f"salvaged={salvaged} DATA_LOSS_DETECTED=True"
            )
        write_log(
            f"--- PAX Fabric Run Summary ---  "
            f"records={result['records_fetched']} rows={result['output_rows']} "
            f"elapsed={result['elapsed_seconds']}s success={result['success']}"
            f"{loss_suffix}"
        )

        # Emit each lost block on its own line so they're trivially greppable
        # (single concatenated summary lines tend to wrap or get truncated).
        for ev in getattr(m, "data_loss_events", []) or []:
            write_log(f"  [DATA-LOSS] {ev}", level="ERROR")

    return result
