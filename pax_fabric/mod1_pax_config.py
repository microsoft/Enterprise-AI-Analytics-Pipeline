"""
PAX Module 1: pax_config
=========================
Configuration, parameters, canonical maps, normalization, and validation logic.

Migrated from PAX_Purview_Audit_Log_Processor_v1.11.2.ps1

This module provides:
- PAXConfig dataclass holding all runtime parameters (replaces PS param() block)
- Canonical maps for recordType/service normalization
- M365 usage bundles (activity types, record types, service types)
- Input normalization (comma-separated splitting, canonical casing)
- Tier inference (get_path_tier) and per-data-type destination resolution
- Validation functions (PAYG billing, append-file compatibility, state contracts)
- Resolve activity types logic (M365 usage, exclusions)
- Noninteractive host detection
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional


# ===========================================================================
# SCRIPT METADATA
# ===========================================================================

SCRIPT_VERSION = "1.11.3"


# ===========================================================================
# CANONICAL MAPS — Used to normalize user input to correct API casing
# ===========================================================================

RECORD_TYPE_CANONICAL_MAP: dict[str, str] = {
    "azureactivedirectory": "AzureActiveDirectory",
    "azureactivedirectoryaccountlogon": "AzureActiveDirectoryAccountLogon",
    "azureactivedirectorystslogon": "AzureActiveDirectoryStsLogon",
    "exchangeadmin": "ExchangeAdmin",
    "exchangeitem": "ExchangeItem",
    "exchangemailbox": "ExchangeMailbox",
    "sharepointfileoperation": "SharePointFileOperation",
    "sharepointsharingoperation": "SharePointSharingOperation",
    "sharepoint": "SharePoint",
    "onedrive": "OneDrive",
    "microsoftteams": "MicrosoftTeams",
}

SERVICE_CANONICAL_MAP: dict[str, str] = {
    "azureactivedirectory": "AzureActiveDirectory",
    "exchange": "Exchange",
    "sharepoint": "SharePoint",
    "onedrive": "OneDrive",
    "teams": "Teams",
}

RECORD_TYPE_WORKLOAD_MAP: dict[str, list[str]] = {
    "azureActiveDirectory": ["AzureActiveDirectory"],
    "azureActiveDirectoryAccountLogon": ["AzureActiveDirectory"],
    "azureActiveDirectoryStsLogon": ["AzureActiveDirectory"],
    "exchangeAdmin": ["Exchange"],
    "exchangeItem": ["Exchange"],
    "exchangeMailbox": ["Exchange"],
    "sharePointFileOperation": ["SharePoint", "OneDrive"],
    "sharePointSharingOperation": ["SharePoint", "OneDrive"],
    "sharePoint": ["SharePoint", "OneDrive"],
    "onedrive": ["OneDrive"],
    "microsoftTeams": ["Teams"],
    # M365 usage record types mapped to Exchange for single-pass processing
    "officeNative": ["Exchange"],
    "microsoftForms": ["Exchange"],
    "microsoftStream": ["Exchange"],
    "plannerPlan": ["Exchange"],
    "plannerTask": ["Exchange"],
    "powerAppsApp": ["Exchange"],
}

SERVICE_OPERATION_MAP: dict[str, list[str]] = {
    "AzureActiveDirectory": [
        "UserLoggedIn", "UserLoginFailed", "AdminLoggedIn",
        "ResetUserPassword", "AddRegisteredUser", "UpdateUser", "ChangedUserSetting",
    ],
    "Exchange": [
        "MailItemsAccessed", "Send", "SendOnBehalf", "SoftDelete", "HardDelete",
        "MoveToDeletedItems", "CopyToFolder", "AddMailboxPermission", "RemoveMailboxPermission",
    ],
    "SharePoint": [
        "FileAccessed", "FileDownloaded", "FileUploaded", "FileModified", "FileDeleted",
        "FileMoved", "SharingInvitationCreated", "SharingInvitationAccepted",
        "SharedLinkCreated", "SharingRevoked", "AddMemberToUnifiedGroup", "RemoveMemberFromUnifiedGroup",
    ],
    "OneDrive": [
        "FileAccessed", "FileDownloaded", "FileUploaded", "FileModified", "FileDeleted",
        "FileMoved", "SharingInvitationCreated", "SharingInvitationAccepted",
        "SharedLinkCreated", "SharingRevoked", "AddMemberToUnifiedGroup", "RemoveMemberFromUnifiedGroup",
    ],
    "Teams": [
        "TeamMemberAdded", "TeamMemberRemoved", "ChannelAdded", "ChannelDeleted",
        "ChannelMessageSent", "ChannelMessageDeleted", "TeamDeleted", "TeamArchived",
        "AddMemberToUnifiedGroup", "RemoveMemberFromUnifiedGroup",
    ],
    "MicrosoftForms": [
        "CreateForm", "EditForm", "DeleteForm", "ViewForm",
        "CreateResponse", "SubmitResponse", "ViewResponse", "DeleteResponse",
    ],
    "MicrosoftStream": ["StreamModified", "StreamViewed", "StreamDeleted", "StreamDownloaded"],
    "MicrosoftPlanner": [
        "PlanCreated", "PlanDeleted", "PlanModified", "TaskCreated",
        "TaskDeleted", "TaskModified", "TaskAssigned", "TaskCompleted",
    ],
    "PowerApps": ["LaunchedApp", "CreatedApp", "EditedApp", "DeletedApp", "PublishedApp"],
}


# ===========================================================================
# M365 USAGE BUNDLES
# ===========================================================================

COPILOT_BASE_ACTIVITY_TYPE = "CopilotInteraction"

M365_USAGE_SERVICE_BUNDLE: list[str] = ["Exchange", "SharePoint", "OneDrive", "Teams"]

M365_USAGE_RECORD_BUNDLE: list[str] = [
    "ExchangeAdmin", "ExchangeItem", "ExchangeMailbox",
    "SharePointFileOperation", "SharePointSharingOperation", "SharePoint",
    "OneDrive", "MicrosoftTeams", "OfficeNative", "MicrosoftForms",
    "MicrosoftStream", "PlannerPlan", "PlannerTask", "PowerAppsApp",
]

# Curated, trimmed M365 usage operations targeted at the Analytics-Hub M365
# Usage Analytics dashboard (v1.11.3). Scope: Exchange mail access,
# SharePoint/OneDrive file access, Teams chat/messaging, Teams meeting
# lifecycle, and Copilot/Connected-AI interaction signals.
M365_USAGE_ACTIVITY_BUNDLE: list[str] = list(dict.fromkeys([
    # === Exchange / Email ===
    "MailItemsAccessed", "MailboxLogin", "Send",
    # === SharePoint / OneDrive - File access ===
    "FileAccessed", "FileViewed", "FilePreviewed", "FileModified", "FileDownloaded", "FileUploaded",
    # === Teams - Chat / Messaging ===
    "MessageSent", "MessageRead", "MessagesListed", "ChatRetrieved", "ChatCreated", "TeamsSessionStarted",
    # === Teams - Meeting lifecycle ===
    "MeetingParticipantJoined", "MeetingStarted", "MeetingEnded", "MeetingParticipantDetail", "MeetingDetail",
    # === Copilot / Connected AI ===
    "CopilotInteraction", "ConnectedAIAppInteraction",
]))


# ===========================================================================
# PAXConfig — Central configuration dataclass (replaces PS param() block)
# ===========================================================================

@dataclass
class PAXConfig:
    """All runtime parameters for a PAX execution run."""

    # --- Date range ---
    start_date: Optional[str] = None  # yyyy-MM-dd or '*'
    end_date: Optional[str] = None    # yyyy-MM-dd or '*'

    # --- Output ---
    output_path: Optional[str] = None
    output_path_user_info: Optional[str] = None       # Per-data-type: EntraUsers CSV
    output_path_agent365_info: Optional[str] = None   # Per-data-type: Agent 365 catalog
    output_path_log: Optional[str] = None              # Per-data-type: log file
    flat_depth: int = 120

    # --- Authentication ---
    # Only app-only client-credential flow is supported (AppRegistration + client_secret).
    # Other auth methods (WebLogin/DeviceCode/Silent/Credential/ManagedIdentity) and
    # certificate credentials were removed in v1.11.4 as unused.
    auth: str = "AppRegistration"
    tenant_id: Optional[str] = None
    client_id: Optional[str] = None
    client_secret: Optional[str] = None

    # --- Query tuning ---
    block_hours: float = 0.5
    partition_hours: int = 0
    max_partitions: int = 160
    result_size: int = 10000
    pacing_ms: int = 0
    max_concurrency: int = 10

    # --- Activity/Record/Service types ---
    activity_types: list[str] = field(default_factory=lambda: ["CopilotInteraction"])
    record_types: Optional[list[str]] = None
    service_types: Optional[list[str]] = None

    # --- Explosion modes ---
    explode_arrays: bool = False
    explode_deep: bool = False

    # --- Replay mode ---
    raw_input_csv: Optional[str] = None

    # --- Parallel processing ---
    enable_parallel: bool = False
    max_parallel_groups: int = 8
    parallel_mode: str = "Auto"  # Off | On | Auto
    explosion_threads: int = 0   # 0=auto, 1=serial, 2-32=explicit

    # --- Adaptive safeguards ---
    disable_adaptive: bool = False
    progress_smoothing_alpha: float = 0.3
    high_latency_ms: int = 90000
    memory_pressure_mb: int = 1500
    max_memory_mb: int = -1
    # Resolved at startup by initialize_config() — mirrors PS $script:ResolvedMaxMemoryMB
    # / $script:memoryFlushEnabled (PS L16213-16232). Page-flush is gated on the flag,
    # never on a live RSS comparison. See README "Memory Optimization" notes.
    resolved_max_memory_mb: int = 0
    memory_flush_enabled: bool = False
    status_interval_seconds: int = 60
    low_latency_ms: int = 20000
    low_latency_consecutive: int = 2
    throughput_drop_pct: int = 15
    throughput_smoothing_alpha: float = 0.3
    adaptive_concurrency_ceiling: int = 6

    # --- Export ---
    export_progress_interval: int = 10
    streaming_schema_sample: int = 5000
    streaming_chunk_size: int = 5000

    # --- Filtering ---
    agent_id: Optional[list[str]] = None
    agents_only: bool = False
    exclude_agents: bool = False
    prompt_filter: Optional[str] = None  # Prompt | Response | Both | Null
    user_ids: Optional[list[str]] = None
    group_names: Optional[list[str]] = None

    # --- Reliability ---
    circuit_breaker_threshold: int = 5
    circuit_breaker_cooldown_seconds: int = 120
    backoff_base_seconds: float = 1.0
    backoff_max_seconds: int = 45
    max_network_outage_minutes: int = 30
    # End-of-run partition retry (PS parity: v1.11.3 L25686 — up to 5 total
    # attempts per partition with reduced concurrency on retry passes).
    partition_max_attempts: int = 5
    partition_retry_max_concurrency: int = 3
    # HTTP 429 / Retry-After throttle handling (v1.11.3+).
    # On a Graph throttle response, sleep at least throttle_min_wait_seconds
    # (or whatever Retry-After tells us, whichever is larger), capped at
    # throttle_max_wait_seconds. The waiting thread also bumps a process-wide
    # throttle deadline so sibling parallel partitions yield instead of
    # piling on top of an already-rate-limited endpoint.
    respect_retry_after: bool = True
    throttle_min_wait_seconds: float = 30.0
    throttle_max_wait_seconds: float = 180.0

    # --- Feature switches ---
    include_copilot_interaction: bool = False
    include_m365_usage: bool = False
    exclude_copilot_interaction: bool = False
    export_workbook: bool = False
    append_file: Optional[str] = None
    append_user_info: Optional[str] = None
    append_agent365_info: Optional[str] = None
    combine_output: bool = False
    force: bool = False
    skip_diagnostics: bool = False
    use_eom: bool = False
    include_user_info: bool = False
    only_user_info: bool = False
    include_agent365_info: bool = False
    only_agent365_info: bool = False
    include_telemetry: bool = False
    rollup: bool = False
    rollup_plus_raw: bool = False
    emit_metrics_json: bool = False
    metrics_path: Optional[str] = None
    auto_completeness: bool = False

    # --- Resume ---
    resume: Optional[str] = None  # None=not resuming, ''=auto-discover, 'path'=explicit

    # --- Remote output (computed from tier inference) ---
    remote_output_mode: str = "None"            # 'None' | 'SharePoint' | 'Fabric'
    remote_output_url: Optional[str] = None     # Trimmed destination URL
    remote_scratch_dir: Optional[str] = None    # Temp local scratch folder (deleted on success)

    # --- Computed at runtime (populated by validate()) ---
    script_run_timestamp: str = field(default_factory=lambda: datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S"))
    trim_start_date_utc: Optional[datetime] = None
    trim_end_date_utc: Optional[datetime] = None

    # --- Phase A0 / Fabric integration (set by pax_fabric.pipeline) ---
    run_id: Optional[str] = None
    csv_output_root: Optional[str] = None

    # --- Track whether output_path was explicitly set by caller ---
    _output_path_explicit: bool = False

    # --- Track whether dates were explicitly provided (mirrors PS $PSBoundParameters.ContainsKey) ---
    _start_date_explicit: bool = field(init=False, default=False)
    _end_date_explicit: bool = field(init=False, default=False)

    def __post_init__(self):
        """Auto-detect whether dates were explicitly set at construction time."""
        self._start_date_explicit = self.start_date is not None
        self._end_date_explicit = self.end_date is not None


# ===========================================================================
# NORMALIZATION FUNCTIONS
# ===========================================================================

def resolve_comma_separated_values(values: Optional[list[str]]) -> Optional[list[str]]:
    """
    Split any comma-separated entries in a list and deduplicate.
    Equivalent to PS: splitting on ',', trimming quotes/whitespace, deduplicating.
    """
    if not values:
        return values

    result: list[str] = []
    for value in values:
        if value is None:
            continue
        for piece in value.split(","):
            token = piece.strip().strip("'\"")
            if token:
                result.append(token)

    # Deduplicate preserving order
    return list(dict.fromkeys(result)) or None


def normalize_record_types(record_types: Optional[list[str]]) -> Optional[list[str]]:
    """Normalize record type names to canonical casing (deduplicated)."""
    if not record_types:
        return None

    processed = resolve_comma_separated_values(record_types)
    if not processed:
        return None

    normalized: list[str] = []
    for rt in processed:
        key = rt.lower()
        normalized.append(RECORD_TYPE_CANONICAL_MAP.get(key, rt))

    return list(dict.fromkeys(normalized)) or None


def normalize_service_types(service_types: Optional[list[str]]) -> Optional[list[str]]:
    """Normalize service type names to canonical casing (deduplicated)."""
    if not service_types:
        return None

    processed = resolve_comma_separated_values(service_types)
    if not processed:
        return None

    normalized: list[str] = []
    for svc in processed:
        key = svc.lower()
        normalized.append(SERVICE_CANONICAL_MAP.get(key, svc))

    return list(dict.fromkeys(normalized)) or None


def resolve_activity_types(config: PAXConfig) -> list[str]:
    """
    Resolve the final list of activity types based on all switches and overrides.
    Mirrors PS Resolve-CommaSeparatedValues logic (Steps 1-5, DSPM removed in v1.11.2).
    
    Returns the deduplicated final activity type list.
    """
    final: list[str] = []

    # Step 1: Process explicit activity types
    if config.activity_types:
        processed = resolve_comma_separated_values(config.activity_types)
        if processed:
            final.extend(processed)

    # Step 2: Add CopilotInteraction when explicitly requested via switch
    if config.include_copilot_interaction and COPILOT_BASE_ACTIVITY_TYPE not in final:
        final.append(COPILOT_BASE_ACTIVITY_TYPE)

    # Step 3: Add M365 usage bundle when requested
    if config.include_m365_usage:
        final.extend(M365_USAGE_ACTIVITY_BUNDLE)

    # Step 4: Base activity type — add CopilotInteraction as default
    # Auto-add when user didn't explicitly provide activity types
    user_provided_custom = config.activity_types != ["CopilotInteraction"]
    if not config.exclude_copilot_interaction:
        if not user_provided_custom:
            if COPILOT_BASE_ACTIVITY_TYPE not in final:
                final.insert(0, COPILOT_BASE_ACTIVITY_TYPE)

    # Step 5: Exclusion override — remove CopilotInteraction if excluded
    if config.exclude_copilot_interaction:
        final = [at for at in final if at != COPILOT_BASE_ACTIVITY_TYPE]

    # Final deduplication
    return list(dict.fromkeys(final))


# ===========================================================================
# TIER INFERENCE & PATH RESOLUTION (v1.11.2)
# ===========================================================================

# URL patterns for tier detection
_SP_URL_PATTERN = re.compile(
    r'^https?://[^/]+\.sharepoint(?:-df|-mil)?\.[a-z]{2,3}(?:/.+)?$'
)
_FABRIC_ROOT_PATTERN = re.compile(
    r'^https://([a-z0-9-]+-)?onelake\.dfs\.fabric\.microsoft\.com/'
    r'[^/]+/[^/]+\.Lakehouse(/Tables(/[A-Za-z_][A-Za-z0-9_]*)?|/Files(/.+)?)?/?$'
)
_FABRIC_FILES_PATTERN = re.compile(
    r'^https://([a-z0-9-]+-)?onelake\.dfs\.fabric\.microsoft\.com/'
    r'[^/]+/[^/]+\.Lakehouse/Files(/.+)?/?$'
)


def get_path_tier(
    value: str,
    switch_name: str,
    *,
    allow_fabric_files_only: bool = False,
) -> Optional[str]:
    """
    Infer storage tier from a destination path value.

    Returns ``'Local'``, ``'SharePoint'``, or ``'Fabric'``.
    Raises ``ValueError`` on UNC paths or unrecognised URL forms.

    Maps to PS ``script:Get-PathTier`` at L2370.
    """
    if not value or not value.strip():
        return None

    v = value.strip()

    # UNC rejection
    if v.startswith("\\\\"):
        raise ValueError(
            f"-{switch_name} does not accept UNC paths ('{v}'). "
            "Provide a drive-rooted local path, a SharePoint URL, or a Fabric OneLake URL."
        )

    # URL-based detection
    if re.match(r'^https?://', v):
        if _SP_URL_PATTERN.match(v):
            return "SharePoint"

        if allow_fabric_files_only:
            if _FABRIC_FILES_PATTERN.match(v):
                return "Fabric"
            raise ValueError(
                f"-{switch_name} on a Fabric destination must point under Files/ "
                f"(logs are not tabular). Provided: {v}"
            )

        if _FABRIC_ROOT_PATTERN.match(v):
            return "Fabric"

        raise ValueError(
            f"-{switch_name} URL is not a recognized SharePoint or Fabric Lakehouse destination. "
            f"Provided: {v}"
        )

    # Local: require drive-rooted absolute path (Windows or Unix)
    if re.match(r'^[A-Za-z]:[\\/]', v) or v.startswith("/"):
        return "Local"

    raise ValueError(
        f"-{switch_name} must be a drive-rooted absolute path, a SharePoint URL, "
        f"or a Fabric OneLake URL. Provided: {v}"
    )


# Data-type keys used throughout the per-data-type destination model
_DATA_TYPE_KEYS = ("Purview", "UserInfo", "Agent365Info", "Log")

# Maps data-type key → (OutputPath* config attr, AllowFabricFilesOnly flag)
_DEST_SWITCH_MAP: dict[str, tuple[str, bool]] = {
    "Purview":      ("output_path",              False),
    "UserInfo":     ("output_path_user_info",    False),
    "Agent365Info": ("output_path_agent365_info", False),
    "Log":          ("output_path_log",           True),
}

# Maps data-type key → Append* config attr
_APPEND_SWITCH_MAP: dict[str, str] = {
    "Purview":      "append_file",
    "UserInfo":     "append_user_info",
    "Agent365Info": "append_agent365_info",
}


def resolve_data_type_paths(
    data_type: str,
    default_basename: str,
    config: PAXConfig,
    *,
    dest_tier: dict,
    dest_raw: dict,
    dest_is_bound: dict,
    append_is_bound: dict,
    append_raw: dict,
) -> dict:
    """
    Central per-data-type effective destination lookup.

    Returns ``{'tier': str, 'raw': str, 'is_bound': bool,
               'effective_dir': str, 'basename': str}``.

    Maps to PS ``script:Resolve-DataTypePaths`` at L2753.
    """
    is_bound = dest_is_bound.get(data_type, False)

    # Fall-through: if OutputPath* not bound but Append* promoted DestRaw, treat as bound
    bound_via_append_only = False
    if not is_bound and append_is_bound.get(data_type, False) and data_type in dest_raw:
        is_bound = True
        bound_via_append_only = True

    if not is_bound:
        # Inherit from Purview
        tier = dest_tier.get("Purview", "Local")
        raw = dest_raw.get("Purview", config.output_path or "")
    else:
        tier = dest_tier.get(data_type, "Local")
        raw = dest_raw.get(data_type, "")

    # Determine file-form vs folder-form
    is_file_form = False
    if tier == "Local":
        is_file_form = bool(
            re.search(r'\.[a-zA-Z0-9]{2,5}$', raw) and not raw.endswith(("/", "\\"))
        )
    elif tier in ("SharePoint", "Fabric"):
        last_seg = raw.rstrip("/").rsplit("/", 1)[-1] if raw else ""
        is_file_form = bool(re.search(r'\.[a-zA-Z0-9]{2,5}$', last_seg))

    if is_file_form:
        if tier == "Local":
            eff_dir = str(Path(raw).parent)
            basename = Path(raw).name if not bound_via_append_only else default_basename
        else:
            eff_dir = raw.rstrip("/").rsplit("/", 1)[0]
            basename = (
                default_basename
                if bound_via_append_only
                else raw.rstrip("/").rsplit("/", 1)[-1]
            )
        return {
            "tier": tier,
            "raw": raw,
            "is_bound": is_bound,
            "effective_dir": eff_dir,
            "basename": basename,
        }

    return {
        "tier": tier,
        "raw": raw,
        "is_bound": is_bound,
        "effective_dir": raw,
        "basename": default_basename,
    }


def test_is_non_interactive() -> bool:
    """
    Detect whether the current host is noninteractive.

    Checks (in order):
    - ``PAX_FORCE_INTERACTIVE`` env var → force interactive (return False)
    - ``PAX_NONINTERACTIVE`` env var → force noninteractive (return True)
    - ``sys.stdin.isatty()`` → False means redirected stdin
    - CI environment indicators (``CI``, ``TF_BUILD``, ``GITHUB_ACTIONS``,
      ``JENKINS_URL``, ``CONTAINER``)

    Maps to PS ``script:Test-IsNonInteractive`` at L3040.
    """
    _TRUTHY = {"1", "true", "True", "TRUE", "yes", "Yes", "YES"}

    if os.environ.get("PAX_FORCE_INTERACTIVE", "") in _TRUTHY:
        return False
    if os.environ.get("PAX_NONINTERACTIVE", "") in _TRUTHY:
        return True

    # stdin redirect detection
    try:
        if not sys.stdin.isatty():
            return True
    except Exception:
        return True

    # CI environment indicators
    ci_vars = ("CI", "TF_BUILD", "GITHUB_ACTIONS", "JENKINS_URL")
    for var in ci_vars:
        if os.environ.get(var):
            return True

    return False


# ===========================================================================
# VALIDATION FUNCTIONS
# ===========================================================================

def validate_config(config: PAXConfig) -> list[str]:
    """
    Validate configuration parameters. Returns list of error messages.
    Empty list means valid.
    
    Mirrors the PS early-exit validation blocks.
    """
    errors: list[str] = []

    # MaxConcurrency range (1-10)
    if not (1 <= config.max_concurrency <= 10):
        errors.append(
            f"MaxConcurrency must be between 1 and 10. "
            f"Microsoft Purview enforces a max of 10 concurrent search jobs per user. "
            f"Current value: {config.max_concurrency}"
        )

    # BlockHours must be positive
    if config.block_hours <= 0:
        errors.append("BlockHours must be positive.")

    # ExcludeAgents vs AgentId/AgentsOnly mutual exclusion
    if config.exclude_agents and (config.agent_id or config.agents_only):
        errors.append(
            "ExcludeAgents cannot be combined with AgentId or AgentsOnly. "
            "These filters are mutually exclusive."
        )

    # Rollup mutual exclusion
    if config.rollup and config.rollup_plus_raw:
        errors.append("Rollup and RollupPlusRaw are mutually exclusive.")

    # IncludeAgent365Info / OnlyAgent365Info mutual exclusion
    if config.include_agent365_info and config.only_agent365_info:
        errors.append("IncludeAgent365Info and OnlyAgent365Info are mutually exclusive.")

    # OnlyAgent365Info is unsupported under app-only auth.
    # The Agent Package Management API requires delegated permissions (signed-in user),
    # but the pipeline now only supports AppRegistration + client_secret (app-only).
    if config.only_agent365_info:
        errors.append(
            "OnlyAgent365Info is not supported. The Agent Package Management API "
            "requires delegated permissions, but only app-only AppRegistration auth "
            "is supported. Use IncludeAgent365Info on an interactive host instead."
        )

    # IncludeAgent365Info/OnlyAgent365Info incompatible with replay and EOM modes.
    # PS L2659-2680: both switches blocked with RAWInputCSV/UseEOM;
    # Resume only blocked with OnlyAgent365Info (IncludeAgent365Info IS compatible with Resume).
    if config.include_agent365_info or config.only_agent365_info:
        agent_switch = "OnlyAgent365Info" if config.only_agent365_info else "IncludeAgent365Info"
        incompat_modes: list[str] = []
        if config.raw_input_csv:
            incompat_modes.append("RAWInputCSV (replay mode)")
        if config.use_eom:
            incompat_modes.append("UseEOM (Exchange Online Management mode)")
        if config.only_agent365_info and config.resume is not None:
            incompat_modes.append("Resume")
        if incompat_modes:
            errors.append(
                f"{agent_switch} is not supported with: {', '.join(incompat_modes)}. "
                f"Agent 365 enrichment requires a fresh live Microsoft Graph context."
            )

    # OnlyAgent365Info conflicts with audit-implying switches (PS L2627-2650)
    if config.only_agent365_info:
        conflicting: list[str] = []
        if config.include_m365_usage:
            conflicting.append("IncludeM365Usage")
        if config.include_copilot_interaction:
            conflicting.append("IncludeCopilotInteraction")
        if config.agents_only:
            conflicting.append("AgentsOnly")
        if config.exclude_agents:
            conflicting.append("ExcludeAgents")
        if config.combine_output:
            conflicting.append("CombineOutput")
        if config.only_user_info:
            conflicting.append("OnlyUserInfo")
        if config.append_file:
            conflicting.append("AppendFile")
        if conflicting:
            errors.append(
                f"OnlyAgent365Info cannot be combined with: {', '.join(conflicting)}. "
                f"OnlyAgent365Info skips the Purview audit pull entirely."
            )

    # RAWInputCSV conflict params
    if config.raw_input_csv:
        conflict_fields = []
        if config.block_hours != 0.5:
            conflict_fields.append("BlockHours")
        if config.result_size != 10000:
            conflict_fields.append("ResultSize")
        if config.pacing_ms != 0:
            conflict_fields.append("PacingMs")
        if config.parallel_mode != "Auto":
            conflict_fields.append("ParallelMode")
        if config.max_parallel_groups != 8:
            conflict_fields.append("MaxParallelGroups")
        if config.max_concurrency != 10:
            conflict_fields.append("MaxConcurrency")
        if config.enable_parallel:
            conflict_fields.append("EnableParallel")
        if config.group_names:
            conflict_fields.append("GroupNames")
        if conflict_fields:
            errors.append(
                f"RAWInputCSV (replay mode) is incompatible with: {', '.join(conflict_fields)}. "
                f"These parameters require live Purview queries."
            )

    # UseEOM incompatible with parallel
    if config.use_eom:
        if config.enable_parallel:
            errors.append("UseEOM is incompatible with EnableParallel. EOM mode is serial-only.")
        if config.parallel_mode not in ("Off", "Auto"):
            errors.append("UseEOM is incompatible with ParallelMode=On. EOM mode is serial-only.")

    # AppendFile validation
    if config.append_file:
        if config.only_user_info:
            errors.append(
                "AppendFile cannot be used with OnlyUserInfo. "
                "AppendFile targets the Purview activity stream, which is out of scope "
                "in only-modes. Use AppendUserInfo to append the EntraUsers snapshot."
            )
        if config.only_agent365_info:
            errors.append(
                "AppendFile cannot be used with OnlyAgent365Info. "
                "AppendFile targets the Purview activity stream, which is out of scope "
                "in only-modes. Use AppendAgent365Info to append the Agent 365 catalog."
            )
        # AppendFile must be a filename, not a directory (PS L2697-2710)
        if config.append_file.endswith("/") or config.append_file.endswith("\\"):
            errors.append(
                "AppendFile must specify a filename, not a directory path."
            )
        else:
            append_ext = Path(config.append_file).suffix.lower()
            if not append_ext:
                errors.append(
                    "AppendFile must include a file extension (.csv or .xlsx)."
                )
            elif config.export_workbook and append_ext != ".xlsx":
                errors.append(
                    "AppendFile must use .xlsx extension when ExportWorkbook is specified."
                )
            elif not config.export_workbook and append_ext not in (".csv", ""):
                errors.append(
                    "AppendFile must use .csv extension for CSV mode."
                )

    # OutputPath folder-only validation (PS L2435-2455)
    if config.output_path:
        if re.search(r'\.[a-zA-Z0-9]{2,4}$', config.output_path) and not config.output_path.endswith("/") and not config.output_path.endswith("\\"):
            errors.append(
                "OutputPath must be a folder path only. Custom filenames are not supported. "
                "The script automatically generates timestamped filenames."
            )

    # OnlyUserInfo incompatible params (PS L1850-1970)
    if config.only_user_info:
        only_user_conflicts: list[str] = []
        # Date filtering (use explicit-tracking flags — mirrors PS $PSBoundParameters.ContainsKey)
        if config._start_date_explicit:
            only_user_conflicts.append("StartDate")
        if config._end_date_explicit:
            only_user_conflicts.append("EndDate")
        # Activity configuration
        if config.activity_types != ["CopilotInteraction"]:
            only_user_conflicts.append("ActivityTypes")
        if config.include_m365_usage:
            only_user_conflicts.append("IncludeM365Usage")
        if config.exclude_copilot_interaction:
            only_user_conflicts.append("ExcludeCopilotInteraction")
        # Audit retrieval settings
        if config.block_hours != 0.5:
            only_user_conflicts.append("BlockHours")
        if config.partition_hours != 0:
            only_user_conflicts.append("PartitionHours")
        if config.max_partitions != 160:
            only_user_conflicts.append("MaxPartitions")
        if config.result_size != 10000:
            only_user_conflicts.append("ResultSize")
        if config.pacing_ms != 0:
            only_user_conflicts.append("PacingMs")
        if config.auto_completeness:
            only_user_conflicts.append("AutoCompleteness")
        if config.streaming_schema_sample != 5000:
            only_user_conflicts.append("StreamingSchemaSample")
        if config.streaming_chunk_size != 5000:
            only_user_conflicts.append("StreamingChunkSize")
        if config.export_progress_interval != 10:
            only_user_conflicts.append("ExportProgressInterval")
        # Filtering
        if config.agent_id:
            only_user_conflicts.append("AgentId")
        if config.agents_only:
            only_user_conflicts.append("AgentsOnly")
        if config.exclude_agents:
            only_user_conflicts.append("ExcludeAgents")
        if config.prompt_filter:
            only_user_conflicts.append("PromptFilter")
        if config.user_ids:
            only_user_conflicts.append("UserIds")
        if config.group_names:
            only_user_conflicts.append("GroupNames")
        if config.record_types:
            only_user_conflicts.append("RecordTypes")
        if config.service_types:
            only_user_conflicts.append("ServiceTypes")
        # Processing mode
        if config.explode_arrays:
            only_user_conflicts.append("ExplodeArrays")
        if config.explode_deep:
            only_user_conflicts.append("ExplodeDeep")
        if config.raw_input_csv:
            only_user_conflicts.append("RAWInputCSV")
        # Parallel processing
        if config.enable_parallel:
            only_user_conflicts.append("EnableParallel")
        if config.max_concurrency != 10:
            only_user_conflicts.append("MaxConcurrency")
        if config.max_parallel_groups != 8:
            only_user_conflicts.append("MaxParallelGroups")
        if config.parallel_mode != "Auto":
            only_user_conflicts.append("ParallelMode")
        if config.disable_adaptive:
            only_user_conflicts.append("DisableAdaptive")
        # Adaptive tuning (only if non-default)
        if config.progress_smoothing_alpha != 0.3:
            only_user_conflicts.append("ProgressSmoothingAlpha")
        if config.high_latency_ms != 90000:
            only_user_conflicts.append("HighLatencyMs")
        if config.memory_pressure_mb != 1500:
            only_user_conflicts.append("MemoryPressureMB")
        if config.low_latency_ms != 20000:
            only_user_conflicts.append("LowLatencyMs")
        if config.low_latency_consecutive != 2:
            only_user_conflicts.append("LowLatencyConsecutive")
        if config.throughput_drop_pct != 15:
            only_user_conflicts.append("ThroughputDropPct")
        if config.throughput_smoothing_alpha != 0.3:
            only_user_conflicts.append("ThroughputSmoothingAlpha")
        if config.adaptive_concurrency_ceiling != 6:
            only_user_conflicts.append("AdaptiveConcurrencyCeiling")
        # Reliability (only if non-default)
        if config.circuit_breaker_threshold != 5:
            only_user_conflicts.append("CircuitBreakerThreshold")
        if config.circuit_breaker_cooldown_seconds != 120:
            only_user_conflicts.append("CircuitBreakerCooldownSeconds")
        if config.backoff_base_seconds != 1.0:
            only_user_conflicts.append("BackoffBaseSeconds")
        if config.backoff_max_seconds != 45:
            only_user_conflicts.append("BackoffMaxSeconds")
        # Alternative modes
        if config.use_eom:
            only_user_conflicts.append("UseEOM")
        # Output combination
        if config.combine_output:
            only_user_conflicts.append("CombineOutput")
        if config.append_file:
            only_user_conflicts.append("AppendFile")
        if only_user_conflicts:
            errors.append(
                f"OnlyUserInfo cannot be used with: {', '.join(only_user_conflicts)}. "
                f"OnlyUserInfo exports only Entra user directory and license information (no audit logs)."
            )

    # =========================================================================
    # AUTO-IMPLY IncludeUserInfo / IncludeAgent365Info from Append* switches
    # Must run BEFORE XOR validation so the in-scope determination is correct.
    # Maps to PS L2882-2883.
    # =========================================================================
    if config.append_user_info and not config.include_user_info:
        config.include_user_info = True
    if config.append_agent365_info and not config.include_agent365_info:
        config.include_agent365_info = True

    # =========================================================================
    # DESTINATION PAIR XOR VALIDATION (v1.11.2)
    # For each output stream, when in scope, the user must supply EXACTLY ONE
    # of (OutputPath* | Append*) — never both, never neither.
    # When out of scope, neither may be supplied.
    # Skipped under Resume: checkpoint rehydrates destinations.
    # Maps to PS L2895-2980.
    # =========================================================================
    if config.resume is None:
        pv_out_bound = config.output_path is not None
        pv_app_bound = config.append_file is not None
        ui_out_bound = config.output_path_user_info is not None
        ui_app_bound = config.append_user_info is not None
        ag_out_bound = config.output_path_agent365_info is not None
        ag_app_bound = config.append_agent365_info is not None

        purview_in_scope = not config.only_user_info and not config.only_agent365_info
        user_info_in_scope = config.include_user_info or config.only_user_info
        agent_in_scope = config.include_agent365_info or config.only_agent365_info

        # --- Purview stream ---
        if purview_in_scope:
            if pv_out_bound and pv_app_bound:
                errors.append(
                    "OutputPath and AppendFile cannot both be supplied. "
                    "For the Purview activity stream, provide EXACTLY ONE of the pair."
                )
            if not pv_out_bound and not pv_app_bound:
                errors.append(
                    "Purview audit output destination not specified. "
                    "Supply EXACTLY ONE of: OutputPath OR AppendFile."
                )

        # --- UserInfo stream ---
        if user_info_in_scope:
            if ui_out_bound and ui_app_bound:
                errors.append(
                    "OutputPathUserInfo and AppendUserInfo cannot both be supplied. "
                    "Provide EXACTLY ONE for the EntraUsers stream."
                )
            if not ui_out_bound and not ui_app_bound and not config.export_workbook:
                errors.append(
                    "IncludeUserInfo/OnlyUserInfo requires a destination for the EntraUsers stream. "
                    "Supply EXACTLY ONE of: OutputPathUserInfo OR AppendUserInfo."
                )
        else:
            if ui_out_bound or ui_app_bound:
                which = "OutputPathUserInfo" if ui_out_bound else "AppendUserInfo"
                errors.append(
                    f"{which} requires IncludeUserInfo or OnlyUserInfo to be in scope. "
                    "Drop the destination switch, or add IncludeUserInfo/OnlyUserInfo."
                )

        # --- Agent365Info stream ---
        if agent_in_scope:
            if ag_out_bound and ag_app_bound:
                errors.append(
                    "OutputPathAgent365Info and AppendAgent365Info cannot both be supplied. "
                    "Provide EXACTLY ONE for the Agent 365 stream."
                )
            if not ag_out_bound and not ag_app_bound and not config.export_workbook:
                errors.append(
                    "IncludeAgent365Info/OnlyAgent365Info requires a destination for the Agent 365 stream. "
                    "Supply EXACTLY ONE of: OutputPathAgent365Info OR AppendAgent365Info."
                )
        else:
            if ag_out_bound or ag_app_bound:
                which = "OutputPathAgent365Info" if ag_out_bound else "AppendAgent365Info"
                errors.append(
                    f"{which} requires IncludeAgent365Info or OnlyAgent365Info to be in scope. "
                    "Drop the destination switch, or add IncludeAgent365Info/OnlyAgent365Info."
                )

    # =========================================================================
    # PER-DATA-TYPE DESTINATION & TIER VALIDATION (v1.11.2)
    # Replaces the old OutputPathSP/OutputPathFabric mutual-exclusivity checks
    # with tier-inferred validation via get_path_tier().
    # =========================================================================

    # Validate each destination value via get_path_tier (catches UNC, bad URLs)
    _dest_switches = [
        ("OutputPath",           config.output_path,              False),
        ("OutputPathUserInfo",   config.output_path_user_info,    False),
        ("OutputPathAgent365Info", config.output_path_agent365_info, False),
        ("OutputPathLog",        config.output_path_log,           True),
    ]
    detected_tiers: list[str] = []
    for sw_name, sw_val, files_only in _dest_switches:
        if sw_val:
            try:
                tier = get_path_tier(sw_val, sw_name, allow_fabric_files_only=files_only)
                if tier:
                    detected_tiers.append(tier)
            except ValueError as exc:
                errors.append(str(exc))

    # Validate Append* values via get_path_tier (when they look rooted/URL)
    _append_switches = [
        ("AppendFile",         config.append_file),
        ("AppendUserInfo",     config.append_user_info),
        ("AppendAgent365Info", config.append_agent365_info),
    ]
    for sw_name, sw_val in _append_switches:
        if sw_val:
            v = sw_val.strip()
            is_url = v.startswith("http://") or v.startswith("https://")
            is_rooted = bool(re.match(r'^[A-Za-z]:[\\/]', v) or v.startswith("/"))
            is_unc = v.startswith("\\\\")
            if is_unc:
                errors.append(
                    f"-{sw_name} does not accept UNC paths ('{v}'). "
                    "Provide a relative filename, a drive-rooted local path, "
                    "a SharePoint URL, or a Fabric OneLake URL."
                )
            elif is_url or is_rooted:
                try:
                    a_tier = get_path_tier(v, sw_name)
                    if a_tier:
                        detected_tiers.append(a_tier)
                except ValueError as exc:
                    errors.append(str(exc))

    # Same-tier enforcement: all supplied destinations must resolve to the same tier.
    # Exception: OutputPathLog may be Fabric Files/ when data destinations are Tables/.
    unique_tiers = set(detected_tiers)
    if len(unique_tiers) > 1:
        errors.append(
            f"All destination paths must resolve to the same storage tier in a single run. "
            f"Detected tiers: {', '.join(sorted(unique_tiers))}. "
            f"Provide all Local, all SharePoint, or all Fabric destinations."
        )

    # AppRegistration credential check: client_id + client_secret are required
    # (only auth path supported in v1.11.4+).
    if not (config.client_id and config.client_secret):
        errors.append(
            "AppRegistration auth requires both client_id and client_secret. "
            "Supply them via CLI (-ClientId / -ClientSecret) or environment variables."
        )

    # Agent365 + AppRegistration on noninteractive host: rejected
    if (config.include_agent365_info or config.only_agent365_info):
        if test_is_non_interactive():
            errors.append(
                "Agent 365 + Auth=AppRegistration is not supported on a noninteractive host. "
                "Detected a noninteractive host (container, CI runner, scheduled task, or "
                "pipeline with redirected stdin). Agent 365 enrichment under AppRegistration "
                "requires an interactive delegated sign-in."
            )

    # Date validation
    if config.start_date and config.start_date != "*":
        try:
            datetime.strptime(config.start_date, "%Y-%m-%d")
        except ValueError:
            errors.append(f"StartDate must be yyyy-MM-dd format. Got: {config.start_date}")

    if config.end_date and config.end_date != "*":
        try:
            datetime.strptime(config.end_date, "%Y-%m-%d")
        except ValueError:
            errors.append(f"EndDate must be yyyy-MM-dd format. Got: {config.end_date}")

    if (config.start_date and config.start_date != "*" and
            config.end_date and config.end_date != "*"):
        try:
            s = datetime.strptime(config.start_date, "%Y-%m-%d")
            e = datetime.strptime(config.end_date, "%Y-%m-%d")
            if e < s:
                errors.append(f"EndDate ({config.end_date}) is earlier than StartDate ({config.start_date}).")
        except ValueError:
            pass  # Already caught above

    return errors


def apply_date_defaults(config: PAXConfig) -> None:
    """
    Apply date defaults matching PS logic:
    - Live mode with no dates: yesterday to today (UTC)
    - Replay mode: leave as '*' if unset
    """
    if config.raw_input_csv:
        if not config.start_date:
            config.start_date = "*"
        if not config.end_date:
            config.end_date = "*"
    else:
        if not config.start_date and not config.end_date:
            yesterday_utc = datetime.now(timezone.utc).date() - timedelta(days=1)
            config.start_date = yesterday_utc.strftime("%Y-%m-%d")
            config.end_date = (yesterday_utc + timedelta(days=1)).strftime("%Y-%m-%d")
        elif not config.start_date:
            config.start_date = "*"
        elif not config.end_date:
            config.end_date = "*"


def compute_trim_boundaries(config: PAXConfig) -> None:
    """
    Compute UTC trim boundaries for client-side date-range filtering.
    Purview may return records outside the requested range.
    """
    if config.start_date and config.start_date != "*":
        config.trim_start_date_utc = datetime.strptime(
            config.start_date, "%Y-%m-%d"
        ).replace(tzinfo=timezone.utc)
    else:
        config.trim_start_date_utc = None

    if config.end_date and config.end_date != "*":
        config.trim_end_date_utc = datetime.strptime(
            config.end_date, "%Y-%m-%d"
        ).replace(tzinfo=timezone.utc)
    else:
        config.trim_end_date_utc = None


# ===========================================================================
# M365 USAGE MODE SIDE-EFFECTS
# ===========================================================================

def apply_m365_usage_mode(config: PAXConfig) -> None:
    """
    When IncludeM365Usage is active, apply side-effects:
    - Auto-enable CombineOutput
    - Merge record types with M365 bundle
    - Set ServiceTypes to None (single-pass query)
    """
    if not config.include_m365_usage:
        return

    config.combine_output = True

    # Merge record types
    merged = list(config.record_types or []) + M365_USAGE_RECORD_BUNDLE
    config.record_types = list(dict.fromkeys(merged)) or None

    # Critical: null ServiceTypes for single-pass M365 query
    config.service_types = None


# ===========================================================================
# REMOTE OUTPUT SETUP
# ===========================================================================

def setup_remote_output(config: PAXConfig) -> None:
    """
    Resolve remote output mode from the per-data-type tier inference.
    When a remote tier is active, redirect output_path to a per-run scratch dir.

    v1.11.2: replaces the old output_path_sp / output_path_fabric approach.
    Tier is inferred from get_path_tier() on whichever destination values
    are supplied. If the dominant tier is 'SharePoint' or 'Fabric',
    create a local scratch dir and redirect output_path there.

    Must be called AFTER validate_config (which validates destinations)
    and BEFORE any downstream code derives paths from output_path.
    """
    # Determine the dominant tier. PS L2500-2525:
    # 1. Check Purview (output_path) first
    # 2. Fallback to UserInfo / Agent365Info when Purview is unset (only-modes)
    # Note: output_path_log is NOT checked — log follows the run's tier.
    dominant_tier = "None"
    dominant_url = None

    # Priority 1: Purview destination
    if config.output_path:
        try:
            tier = get_path_tier(config.output_path, "OutputPath")
            if tier in ("SharePoint", "Fabric"):
                dominant_tier = tier
                dominant_url = config.output_path.strip().rstrip("/")
        except ValueError:
            pass

    # Priority 2: Fallback from in-scope non-Purview streams (only-modes)
    if dominant_tier == "None":
        for attr, sw_name in (
            ("output_path_user_info", "OutputPathUserInfo"),
            ("output_path_agent365_info", "OutputPathAgent365Info"),
        ):
            val = getattr(config, attr, None)
            if val:
                try:
                    tier = get_path_tier(val, sw_name)
                    if tier in ("SharePoint", "Fabric"):
                        dominant_tier = tier
                        dominant_url = val.strip().rstrip("/")
                        break
                except ValueError:
                    pass

    # Priority 3: Append-side URLs promote tier (PS L2430-2455 absorbs into DestTier)
    if dominant_tier == "None":
        for attr, sw_name in (
            ("append_file", "AppendFile"),
            ("append_user_info", "AppendUserInfo"),
            ("append_agent365_info", "AppendAgent365Info"),
        ):
            val = getattr(config, attr, None)
            if val and val.strip().startswith("http"):
                try:
                    tier = get_path_tier(val.strip(), sw_name)
                    if tier in ("SharePoint", "Fabric"):
                        dominant_tier = tier
                        dominant_url = val.strip().rstrip("/")
                        break
                except ValueError:
                    pass

    if dominant_tier == "None":
        config.remote_output_mode = "None"
        config.remote_output_url = None
        config.remote_scratch_dir = None
        return

    config.remote_output_mode = dominant_tier
    config.remote_output_url = dominant_url

    # Create a per-run scratch dir under OS temp folder.
    scratch_prefix = "PAX_" + datetime.now().strftime("%Y%m%d_%H%M%S") + "_"
    config.remote_scratch_dir = tempfile.mkdtemp(prefix=scratch_prefix)

    # Redirect output_path to scratch dir
    config.output_path = config.remote_scratch_dir
    if not config.output_path.endswith(os.sep):
        config.output_path += os.sep


# ===========================================================================
# FULL INITIALIZATION PIPELINE
# ===========================================================================

def resolve_max_memory_mb(config: PAXConfig) -> None:
    """
    Resolve MaxMemoryMB and derive memory_flush_enabled.

    Mirrors PS L16213-16232:
        -1 -> 75% of total system RAM (4096 MB fallback if detection fails)
         0 -> disabled (no flushing)
        >0 -> use as-is

    The flush flag is True iff the resolved value is > 0 AND no explosion mode
    is active (explosion needs the full record set in memory). The number itself
    is never compared to live RSS at runtime — flushing is per-page once enabled.
    """
    requested = config.max_memory_mb
    if requested == 0:
        config.resolved_max_memory_mb = 0
    elif requested == -1:
        try:
            import psutil  # type: ignore
            total_mb = int(psutil.virtual_memory().total / (1024 * 1024))
            config.resolved_max_memory_mb = int(round(total_mb * 0.75))
        except Exception:
            config.resolved_max_memory_mb = 4096
    else:
        config.resolved_max_memory_mb = int(requested)

    config.memory_flush_enabled = (
        config.resolved_max_memory_mb > 0
        and not config.explode_deep
        and not config.explode_arrays
        and not config.raw_input_csv
    )


def initialize_config(config: PAXConfig) -> list[str]:
    """
    Run the full Module 1 initialization pipeline on a PAXConfig instance.
    
    1. Apply date defaults
    2. Normalize filter arrays
    3. Apply M365 usage side-effects
    4. Resolve final activity types
    5. Normalize record/service types
    6. Compute trim boundaries
    7. Validate all parameters
    8. Setup remote output (if validation passed)
    
    Returns list of validation errors (empty = success).
    """
    # 1. Date defaults
    apply_date_defaults(config)

    # 2. Normalize filter arrays
    config.activity_types = resolve_comma_separated_values(config.activity_types) or ["CopilotInteraction"]
    if config.user_ids:
        config.user_ids = resolve_comma_separated_values(config.user_ids)
    if config.group_names:
        config.group_names = resolve_comma_separated_values(config.group_names)
    if config.agent_id:
        config.agent_id = resolve_comma_separated_values(config.agent_id)

    # 3. M365 usage side-effects
    apply_m365_usage_mode(config)

    # 4. Resolve final activity types (with M365, exclusions)
    config.activity_types = resolve_activity_types(config)

    # 5. Normalize record/service types
    config.record_types = normalize_record_types(config.record_types)
    config.service_types = normalize_service_types(config.service_types)

    # 6. Compute trim boundaries
    compute_trim_boundaries(config)

    # 6b. Resolve MaxMemoryMB and derive memory_flush_enabled (PS L16213-16232)
    resolve_max_memory_mb(config)

    # 7. Validate
    errors = validate_config(config)

    # 7b. Rollup side-effects (PS L3498-3530)
    # MUST come AFTER validation: in PS, the XOR destination validator runs at
    # L2881-2930 (checking $IncludeUserInfo which is still $false), and the
    # rollup auto-enable of $IncludeUserInfo happens later at L3520. Moving
    # this before validation would cause a spurious "EntraUsers stream requires
    # destination" error when the user supplies only -OutputPath + -Rollup
    # (which is the normal PS usage).
    if not errors and (config.rollup or config.rollup_plus_raw):
        is_copilot_only = (
            not config.include_m365_usage
            and COPILOT_BASE_ACTIVITY_TYPE in config.activity_types
        )
        if is_copilot_only and not config.include_user_info:
            config.include_user_info = True
        if not config.combine_output:
            config.combine_output = True

    # 7c. OnlyUserInfo post-validation side-effects (PS L1970-1971)
    # Must come AFTER validation: validation checks activity_types against
    # the default ["CopilotInteraction"] to detect user-supplied conflicts.
    # PS uses $PSBoundParameters.ContainsKey('ActivityTypes') which checks
    # explicit user input, not the current value. Clearing activity_types
    # before validation would cause a false-positive conflict.
    if not errors and config.only_user_info:
        config.include_user_info = True
        config.activity_types = []  # No audit queries needed

    # 8. Setup remote output (only if no validation errors)
    if not errors:
        setup_remote_output(config)

    return errors


# ===========================================================================
# FABRIC NOTEBOOK ENTRY POINT — programmatic config from dict
# ===========================================================================
#
# Notebooks build a config from a plain dict instead of CLI argv. The keys
# accept both snake_case (Python convention) and PascalCase (legacy PS
# convention) so users can lift parameter blocks from their existing
# ``python -m pax -StartDate ... -EndDate ...`` invocations unchanged.

def _coerce_csv_list(value):
    """Convert "a,b,c" or ["a","b","c"] to a normalized list."""
    if value is None:
        return None
    if isinstance(value, str):
        return [v.strip() for v in value.split(",") if v.strip()]
    if isinstance(value, (list, tuple)):
        return [str(v).strip() for v in value if str(v).strip()]
    return None


def config_from_params(params: dict) -> "PAXConfig":
    """Build a ``PAXConfig`` from a notebook parameters dict.

    Recognised keys (case-insensitive, both ``start_date`` and ``StartDate``):
        StartDate, EndDate, Auth, TenantId, ClientId, ClientSecret,
        ActivityTypes (list or comma-string), RecordTypes, ServiceTypes,
        AgentId, UserIds, GroupNames, PromptFilter,
        Rollup, RollupPlusRaw, IncludeCopilotInteraction,
        ExportWorkbook, IncludeUserInfo, OnlyUserInfo,
        IncludeM365Usage, IncludeAgent365Info, OnlyAgent365Info,
        IncludeDspmForAi, IncludeTelemetry,
        BlockHours, PartitionHours, MaxPartitions, ResultSize, FlatDepth,
        ExplodeArrays, ExplodeDeep,
        RunId, CsvOutputRoot — Fabric lakehouse routing overrides.

    Unknown keys are silently ignored so notebooks remain forward-compatible.
    Returns an *unvalidated* config; callers should run ``initialize_config``
    afterward to populate computed fields and raise on errors.
    """
    if params is None:
        params = {}

    # Build a lower-case lookup once so we can support both naming styles.
    lc = {str(k).lower().replace("-", "_"): v for k, v in params.items()}

    def pick(*aliases, default=None):
        for a in aliases:
            key = a.lower()
            if key in lc and lc[key] is not None:
                return lc[key]
        return default

    cfg = PAXConfig()

    # --- Dates -------------------------------------------------------
    sd = pick("startdate", "start_date")
    ed = pick("enddate", "end_date")
    if sd is not None:
        cfg.start_date = str(sd)
    if ed is not None:
        cfg.end_date = str(ed)

    # --- Output (legacy local + new lakehouse) -----------------------
    op = pick("outputpath", "output_path")
    if op is not None:
        cfg.output_path = str(op)
        cfg._output_path_explicit = True
    cro = pick("csvoutputroot", "csv_output_root")
    if cro is not None:
        cfg.csv_output_root = str(cro)
    rid = pick("runid", "run_id")
    if rid is not None:
        cfg.run_id = str(rid)

    # --- Auth --------------------------------------------------------
    for src, dst in (
        ("auth", "auth"),
        ("tenantid", "tenant_id"),
        ("clientid", "client_id"),
        ("clientsecret", "client_secret"),
    ):
        v = pick(src, dst)
        if v is not None:
            setattr(cfg, dst, str(v) if not isinstance(v, str) else v)

    # --- Filtering lists --------------------------------------------
    for src, dst in (
        ("activitytypes", "activity_types"),
        ("recordtypes", "record_types"),
        ("servicetypes", "service_types"),
        ("agentid", "agent_id"),
        ("userids", "user_ids"),
        ("groupnames", "group_names"),
    ):
        v = pick(src, dst)
        if v is not None:
            coerced = _coerce_csv_list(v)
            if coerced is not None:
                setattr(cfg, dst, coerced)

    pf = pick("promptfilter", "prompt_filter")
    if pf is not None:
        cfg.prompt_filter = str(pf)

    # --- Scalars (boolean & numeric) --------------------------------
    bool_fields = (
        ("rollup", "rollup"),
        ("rollupplusraw", "rollup_plus_raw"),
        ("includecopilotinteraction", "include_copilot_interaction"),
        ("excludecopilotinteraction", "exclude_copilot_interaction"),
        ("includedspmforai", "include_dspm_for_ai"),
        ("includem365usage", "include_m365_usage"),
        ("exportworkbook", "export_workbook"),
        ("includeuserinfo", "include_user_info"),
        ("onlyuserinfo", "only_user_info"),
        ("includeagent365info", "include_agent365_info"),
        ("onlyagent365info", "only_agent365_info"),
        ("includetelemetry", "include_telemetry"),
        ("explodearrays", "explode_arrays"),
        ("explodedeep", "explode_deep"),
        ("agentsonly", "agents_only"),
        ("excludeagents", "exclude_agents"),
        ("force", "force"),
        ("useeom", "use_eom"),
        ("autocompleteness", "auto_completeness"),
        ("skipdiagnostics", "skip_diagnostics"),
        ("disableadaptive", "disable_adaptive"),
        ("emitmetricsjson", "emit_metrics_json"),
        ("enableparallel", "enable_parallel"),
        ("combineoutput", "combine_output"),
        ("respectretryafter", "respect_retry_after"),
    )
    for src, dst in bool_fields:
        v = pick(src, dst)
        if v is not None:
            setattr(cfg, dst, bool(v))

    numeric_fields = (
        ("blockhours", "block_hours", float),
        ("partitionhours", "partition_hours", int),
        ("maxpartitions", "max_partitions", int),
        ("resultsize", "result_size", int),
        ("pacingms", "pacing_ms", int),
        ("maxconcurrency", "max_concurrency", int),
        ("flatdepth", "flat_depth", int),
        ("maxparallelgroups", "max_parallel_groups", int),
        ("explosionthreads", "explosion_threads", int),
        ("highlatencyms", "high_latency_ms", int),
        ("lowlatencyms", "low_latency_ms", int),
        ("memorypressuremb", "memory_pressure_mb", int),
        ("maxmemorymb", "max_memory_mb", int),
        ("statusintervalseconds", "status_interval_seconds", int),
        ("circuitbreakerthreshold", "circuit_breaker_threshold", int),
        ("circuitbreakercooldownseconds", "circuit_breaker_cooldown_seconds", int),
        ("backoffmaxseconds", "backoff_max_seconds", int),
        ("maxnetworkoutageminutes", "max_network_outage_minutes", int),
        ("partitionmaxattempts", "partition_max_attempts", int),
        ("partitionretrymaxconcurrency", "partition_retry_max_concurrency", int),
        ("throttleminwaitseconds", "throttle_min_wait_seconds", float),
        ("throttlemaxwaitseconds", "throttle_max_wait_seconds", float),
    )
    for src, dst, caster in numeric_fields:
        v = pick(src, dst)
        if v is not None:
            try:
                setattr(cfg, dst, caster(v))
            except (TypeError, ValueError):
                pass

    # Other string fields
    for src, dst in (
        ("parallelmode", "parallel_mode"),
        ("appendfile", "append_file"),
        ("metricspath", "metrics_path"),
        ("rawinputcsv", "raw_input_csv"),
        ("resume", "resume"),
    ):
        v = pick(src, dst)
        if v is not None:
            setattr(cfg, dst, str(v))

    cfg.__post_init__()
    return cfg


# ===========================================================================
# ENTRY POINT (for standalone testing)
# ===========================================================================

if __name__ == "__main__":
    # Quick self-test: create default config and validate
    cfg = PAXConfig()
    errors = initialize_config(cfg)

    if errors:
        print("Validation errors:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    else:
        print(f"PAX Config Module v{SCRIPT_VERSION} - OK")
        print(f"  Activity Types: {cfg.activity_types}")
        print(f"  Date Range: {cfg.start_date} to {cfg.end_date}")
        print(f"  Output Path: {cfg.output_path}")
        print(f"  Remote Output Mode: {cfg.remote_output_mode}")
        print(f"  Remote Output URL: {cfg.remote_output_url}")
        print(f"  Remote Scratch Dir: {cfg.remote_scratch_dir}")
        print(f"  Trim Start UTC: {cfg.trim_start_date_utc}")
        print(f"  Trim End UTC: {cfg.trim_end_date_utc}")
        print(f"  Timestamp: {cfg.script_run_timestamp}")
