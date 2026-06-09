"""
PAX Module 9: pax_data_transform
==================================
Record Explosion & Structuring — the core data transformation engine.

Migrated from PAX_Purview_Audit_Log_Processor_v1.11.1.ps1
Source lines: L14724-14762, L14764, L14934-14966, L15246-15866, L15869-16082

Functions (5 public + 3 internal helpers):
  1. find_all_arrays               — Recursive array path discovery       (PS Find-AllArrays L14724)
  2. convert_to_flat_columns       — Recursive JSON → flat dict           (PS ConvertTo-FlatColumns L14934)
  3. convert_to_purview_exploded_records — Core record exploder           (PS Convert-ToPurviewExplodedRecords L15246)
  4. convert_to_structured_record  — Structured record builder            (PS Convert-ToStructuredRecord L15869)
  5. _get_num                      — Numeric coercion helper              (PS Local:Get-Num L15605/15876)
  6. _add_or_update                — Dict key add/update helper           (PS Local:Add-OrUpdate L15877)
  7. _measure_collection           — Collection metric aggregation        (PS Local:Measure-Collection L15949)

Hard dependencies: pax_data_helpers (Module 2)
Optional dependency: pax_profiler (Module 4) via profile_fn callback

Constants:
  - PURVIEW_EXPLODED_HEADER       — Fixed 153-column header              (PS $PurviewExplodedHeader L15047)
  - FLAT_DEPTH_STANDARD           — Standard flattening depth            (PS $FlatDepthStandard = 6)
  - FLAT_DEPTH_DEEP               — Deep flattening depth                (PS $FlatDepthDeep = 120)
  - EXPLOSION_PER_RECORD_ROW_CAP  — Max rows per record                  (PS $ExplosionPerRecordRowCap = 1000)
  - JSON_DEPTH                    — JSON serialization depth             (PS $JsonDepth = 60)
"""

from __future__ import annotations

import json
from typing import Any, Callable, Optional

from .mod2_pax_data_helpers import (
    bool_tf_fast,
    format_date_purview_fast,
    get_array_fast,
    get_safe_property,
    parse_date_safe,
    select_first_non_null,
    test_scalar_value,
    to_json_if_object_fast,
)


# ===========================================================================
# CONSTANTS (match PS $script:* variables at L12840-12843, L15047-15103)
# ===========================================================================

JSON_DEPTH: int = 60
FLAT_DEPTH_STANDARD: int = 6
FLAT_DEPTH_DEEP: int = 120
EXPLOSION_PER_RECORD_ROW_CAP: int = 1000

# Fixed 153-column header for Purview exploded output (PS $PurviewExplodedHeader L15047)
PURVIEW_EXPLODED_HEADER: list[str] = [
    'RecordId', 'CreationDate', 'RecordType', 'Operation', 'UserId',
    'AssociatedAdminUnits', 'AssociatedAdminUnitsNames',
    '@odata.type', 'CreationTime', 'Id', 'OrganizationId',
    'ResultStatus', 'UserKey', 'UserType', 'Version', 'Workload',
    'ClientIP', 'ObjectId', 'AzureActiveDirectoryEventType',
    'ActorContextId', 'ActorIpAddress', 'InterSystemsId', 'IntraSystemId',
    'SupportTicketId', 'TargetContextId', 'ApplicationId',
    'DeviceProperties.OS', 'DeviceProperties.BrowserType',
    'ErrorNumber',
    'SiteUrl', 'SourceRelativeUrl', 'SourceFileName', 'SourceFileExtension',
    'ListId', 'ListItemUniqueId', 'WebId', 'ApplicationDisplayName', 'EventSource',
    'ItemType', 'SiteSensitivityLabelId', 'GeoLocation', 'IsManagedDevice',
    'DeviceDisplayName', 'ListBaseType', 'ListServerTemplate',
    'AuthenticationType', 'Site', 'DoNotDistributeEvent', 'HighPriorityMediaProcessing',
    'BrowserName', 'BrowserVersion', 'CorrelationId', 'Platform', 'UserAgent',
    'ActorInfoString', 'AppId', 'AuthType', 'ClientAppId', 'ClientIPAddress',
    'ClientInfoString', 'ExternalAccess', 'InternalLogonType', 'LogonType',
    'LogonUserSid', 'MailboxGuid', 'MailboxOwnerSid', 'MailboxOwnerUPN',
    'OrganizationName', 'OriginatingServer', 'SessionId',
    'TokenObjectId', 'TokenTenantId', 'TokenType', 'SaveToSentItems',
    'OperationCount', 'FileSizeBytes',
    'MeetingId', 'MeetingType', 'EventSignature', 'EventData',
    'Permission', 'SensitivityLabelId', 'SharingLinkScope',
    'TargetUserOrGroupType', 'TargetUserOrGroupName',
    'MeetingURL', 'ChatId', 'MessageId', 'MessageSizeInBytes', 'MessageType',
    'FormId', 'FormName', 'VideoId', 'VideoName', 'ChannelId', 'ViewDuration',
    'ClientRegion', 'CopilotLogVersion', 'TargetId',
    'TeamName', 'TeamGuid', 'ResponseId', 'IsAnonymous', 'DeviceType',
    'ChannelName', 'ChannelGuid', 'ChannelType', 'AppName', 'EnvironmentName',
    'PlanId', 'PlanName', 'TaskId', 'TaskName', 'PercentComplete',
    'CrossMailboxOperation',
    'RecordTypeNum', 'ResultStatus_Audit',
    'ModelId', 'ModelProvider', 'ModelFamily',
    'TokensTotal', 'TokensInput', 'TokensOutput', 'DurationMs', 'OutcomeStatus',
    'ConversationId', 'TurnNumber', 'RetryCount', 'ClientVersion', 'ClientPlatform',
    'AgentId', 'AgentName', 'AgentVersion', 'AgentCategory', 'ApplicationName',
    'AppHost', 'ThreadId',
    'Context_Id', 'Context_Type',
    'Message_Id', 'Message_isPrompt',
    'AccessedResource_Action', 'AccessedResource_PolicyDetails', 'AccessedResource_SiteUrl',
    'AISystemPlugin_Id', 'AISystemPlugin_Name',
    'ModelTransparencyDetails_ModelName', 'MessageIds',
    'AccessedResource_Name', 'AccessedResource_SensitivityLabel',
    'AccessedResource_ResourceType', 'SensitivityLabel', 'Context_Item',
]


# ===========================================================================
# INTERNAL HELPERS (PS Local:Get-Num, Local:Add-OrUpdate, Local:Measure-Collection)
# ===========================================================================

def _get_num(v: Any) -> Optional[float]:
    """
    Numeric coercion helper — PS Local:Get-Num (L15605/L15876).

    Returns float on success, None on failure or null/empty input.
    """
    if v is None:
        return None
    try:
        if isinstance(v, str) and v.strip() == "":
            return None
        return float(v)
    except (ValueError, TypeError):
        return None


def _add_or_update(obj: dict, name: str, value: Any) -> None:
    """
    Dict key add/update helper — PS Local:Add-OrUpdate (L15877).

    Always sets obj[name] = value regardless of whether key exists.
    """
    obj[name] = value


def _measure_collection(items: Any, prefix: str) -> dict[str, Any]:
    """
    Collection metric aggregation — PS Local:Measure-Collection (L15949-15963).

    Analyzes a collection of items and returns aggregated metrics like counts,
    types, average latency, edits, acceptance rate, success/failure counts.
    """
    result: dict[str, Any] = {}
    if not items:
        return result

    arr = items if isinstance(items, list) else list(items) if hasattr(items, '__iter__') and not isinstance(items, (str, dict)) else [items]
    if len(arr) == 0:
        return result

    result[f"{prefix}Count"] = len(arr)
    types: set[str] = set()
    latencies: list[float] = []
    edits: list[float] = []
    accepted = 0
    success = 0
    failure = 0

    for s in arr:
        if not isinstance(s, dict):
            continue

        # Type detection (try multiple candidate keys)
        for cand in ('Type', 'SuggestionType', 'Name', 'Kind', 'ActionType'):
            if cand in s and s[cand] is not None:
                types.add(str(s[cand]))
                break

        # Latency extraction
        for lat in ('LatencyMs', 'DurationMs', 'ElapsedMs'):
            if lat in s:
                v = _get_num(s[lat])
                if v is not None:
                    latencies.append(v)
                    break

        # Edit count extraction
        for ed in ('EditCount', 'Edits', 'EditsCount'):
            if ed in s:
                v = _get_num(s[ed])
                if v is not None:
                    edits.append(v)
                    break

        # Acceptance detection
        for acc in ('Accepted', 'IsAccepted', 'Success', 'Succeeded'):
            if acc in s:
                val = s[acc]
                if isinstance(val, bool):
                    if val:
                        accepted += 1
                elif isinstance(val, str) and val.lower() in ('true', 'yes', '1', 'success'):
                    accepted += 1
                break

        # Success/failure detection
        for succ in ('Success', 'Succeeded'):
            if succ in s:
                val = s[succ]
                if isinstance(val, bool):
                    if val:
                        success += 1
                    else:
                        failure += 1
                elif isinstance(val, str):
                    if val.lower() in ('true', 'yes', '1', 'success'):
                        success += 1
                    else:
                        failure += 1
                break

    if types:
        result[f"{prefix}Types"] = ';'.join(sorted(types))
    if latencies:
        result[f"{prefix}AvgLatencyMs"] = round(sum(latencies) / len(latencies), 2)
    if edits:
        result[f"{prefix}AvgEdits"] = round(sum(edits) / len(edits), 2)
        result[f"{prefix}TotalEdits"] = sum(edits)
    if accepted > 0:
        result[f"{prefix}Accepted"] = accepted
        result[f"{prefix}AcceptanceRate"] = round((accepted / len(arr)) * 100, 2)
    if success > 0 or failure > 0:
        result[f"{prefix}Success"] = success
        result[f"{prefix}Failure"] = failure

    return result


# ===========================================================================
# AGENT CATEGORIZATION HELPER
# ===========================================================================

def _categorize_agent(agent_id: Any) -> str:
    """Categorize agent based on AgentId pattern — matches PS logic."""
    if not agent_id:
        return ""
    agent_id_str = str(agent_id)
    if agent_id_str.startswith("CopilotStudio.Declarative."):
        return "Declarative Agent"
    elif agent_id_str.startswith("CopilotStudio.CustomEngine."):
        return "Custom Engine Agent"
    elif agent_id_str.startswith("P_"):
        return "Declarative Agent (Purview)"
    else:
        return "Other Agent"


# ===========================================================================
# 1. find_all_arrays — PS Find-AllArrays (L14724-14762)
# ===========================================================================

def find_all_arrays(
    data: Any,
    path: str = '',
    depth: int = 0,
    arrays: Optional[dict[str, dict]] = None
) -> dict[str, dict]:
    """
    Recursive array path discovery.

    Walks a nested data structure and discovers all arrays (lists) at any depth,
    recording their path, data reference, and element count. Stops recursing at
    depth > 6. Does NOT recurse into array elements — arrays are terminal values.

    Args:
        data: The data structure to inspect.
        path: Current dotted path prefix.
        depth: Current recursion depth.
        arrays: Accumulator dict (created if None).

    Returns:
        dict mapping path strings to {Path, Data, Count} info dicts.
    """
    if data is None:
        return {} if arrays is None else arrays

    if arrays is None:
        arrays = {}

    if depth > 6:
        return arrays

    if data is None:
        return arrays

    # Check if data is an array (list/tuple, but not string or dict)
    is_array = (
        isinstance(data, (list, tuple))
        and not isinstance(data, (str, bytes))
    )

    if is_array:
        key = path if path else 'root'
        if key not in arrays:
            arrays[key] = {'Path': path, 'Data': data, 'Count': len(data)}

    # Get properties to recurse into
    props: Optional[list[tuple[str, Any]]] = None
    if isinstance(data, dict):
        props = list(data.items())
    elif hasattr(data, '__dict__') and not isinstance(data, (list, tuple, str, bytes)):
        props = list(vars(data).items())

    if props:
        for name, val in props:
            child_path = f"{path}.{name}" if path else name
            find_all_arrays(data=val, path=child_path, depth=depth + 1, arrays=arrays)

    # Note: Do NOT recurse into array elements - arrays are treated as terminal values
    # that will be converted to JSON strings for predictable column names
    return arrays


# ===========================================================================
# 2. convert_to_flat_columns — PS ConvertTo-FlatColumns (L14934-14966)
# ===========================================================================

def convert_to_flat_columns(
    node: Any,
    prefix: str = '',
    max_depth: int = 60
) -> dict[str, Any]:
    """
    Recursively flatten a nested object into a single-level dict with dotted keys.

    Handles:
      - None → key = None
      - Scalars → key = value
      - Single-element arrays → recurse into element (no index in path)
      - Multi-element arrays → serialize to JSON string
      - Empty arrays → key = ''
      - Dicts/objects → recurse with prefix

    Args:
        node: Root object to flatten.
        prefix: Current key prefix (dotted path).
        max_depth: Maximum recursion depth (default 60, matches PS).

    Returns:
        dict with dotted-path keys and scalar/JSON-string values.
    """
    cols: dict[str, Any] = {}

    def recurse(n: Any, p: str, d: int) -> None:
        if d > max_depth:
            return

        if n is None:
            if p:
                cols[p.rstrip('.')] = None
            return

        if test_scalar_value(n):
            if p:
                cols[p.rstrip('.')] = n
            return

        # Array/list check (not string, not dict)
        if isinstance(n, (list, tuple)) and not isinstance(n, (str, bytes)):
            arr = list(n)
            if len(arr) == 1:
                # Single element: recurse into it without adding index to path (clean column names)
                recurse(arr[0], p, d + 1)
            elif len(arr) > 1:
                # Multiple elements: serialize to JSON (row explosion handles important arrays separately)
                if p:
                    try:
                        cols[p.rstrip('.')] = json.dumps(arr, default=str, ensure_ascii=False)
                    except Exception:
                        cols[p.rstrip('.')] = ''
            else:
                # Empty array
                if p:
                    cols[p.rstrip('.')] = ''
            return

        # Dict/object — recurse into properties
        props: Optional[list[tuple[str, Any]]] = None
        if isinstance(n, dict):
            props = list(n.items())
        elif hasattr(n, '__dict__') and not isinstance(n, (list, tuple, str, bytes)):
            try:
                props = list(vars(n).items())
            except Exception:
                pass

        if props:
            for name, child in props:
                cp = f"{p}{name}." if p else f"{name}."
                recurse(child, cp, d + 1)

    recurse(node, prefix, 0)
    return cols


# ===========================================================================
# 3. convert_to_purview_exploded_records — PS Convert-ToPurviewExplodedRecords (L15246-15866)
# ===========================================================================

def convert_to_purview_exploded_records(
    record: dict[str, Any],
    deep: bool = False,
    partial_explode: bool = False,
    prompt_filter_value: Optional[str] = None,
    skip_metrics: bool = False,
    metrics: Optional[dict[str, Any]] = None,
    deep_extra_columns: Optional[list[str]] = None,
    flat_depth_deep: int = FLAT_DEPTH_DEEP,
    profile_fn: Optional[Callable] = None,
) -> list[dict[str, Any]]:
    """
    Core Copilot/DSPM record exploder — PS Convert-ToPurviewExplodedRecords (L15246-15866).

    Transforms a single raw audit record into one or more flattened output rows.
    For non-Copilot records, produces exactly 1 row with 153 fixed columns.
    For Copilot records, explodes arrays (Messages, Contexts, AccessedResources,
    AISystemPlugin, ModelTransparencyDetails, SensitivityLabels) into multiple rows.

    Args:
        record: Raw audit record dict with at minimum AuditData (JSON string or parsed).
        deep: If True, deep-flatten CED + AuditData into extra columns.
        partial_explode: If True, prompt-specific explosion only (preserves AuditData column).
        prompt_filter_value: Filter messages by isPrompt ('Prompt', 'Response', 'Null', 'Both').
        skip_metrics: If True, skip metrics aggregation (for parallel replay).
        metrics: Shared metrics dict (modified in-place if provided).
        deep_extra_columns: Shared list of discovered extra column names.
        flat_depth_deep: Flattening depth for deep mode.
        profile_fn: Optional profiler callback (receives audit_data dict).

    Returns:
        List of flattened row dicts (possibly empty on parse failure).
    """
    if metrics is None:
        metrics = {}
    if deep_extra_columns is None:
        deep_extra_columns = []

    try:
        # Parse AuditData
        audit_data = record.get('_ParsedAuditData')
        if audit_data is None:
            raw_audit = record.get('AuditData')
            if raw_audit and isinstance(raw_audit, str):
                try:
                    audit_data = json.loads(raw_audit)
                except (json.JSONDecodeError, ValueError):
                    audit_data = None
            elif isinstance(raw_audit, dict):
                audit_data = raw_audit

        if not audit_data:
            if not skip_metrics:
                metrics['FilteringSkippedRecords'] = metrics.get('FilteringSkippedRecords', 0) + 1
                metrics['FilteringMissingAuditData'] = metrics.get('FilteringMissingAuditData', 0) + 1
            return []

        # Profile audit data (optional callback)
        if profile_fn:
            try:
                profile_fn(audit_data)
            except Exception:
                pass

        # Extract CopilotEventData
        ced = get_safe_property(audit_data, 'CopilotEventData')

        if not ced:
            # ── Non-Copilot path: fixed 153-column extraction (no dynamic discovery) ──
            record_id = (
                record.get('RecordId')
                or record.get('Identity')
                or record.get('Id')
                or get_safe_property(audit_data, 'Id')
            )
            creation_date = format_date_purview_fast(record.get('CreationDate'))
            creation_time = format_date_purview_fast(get_safe_property(audit_data, 'CreationTime'))
            op_value = (
                get_safe_property(audit_data, 'Operation')
                or record.get('Operation')
                or record.get('Operations')
                or ''
            )
            uid_value = (
                get_safe_property(audit_data, 'UserId')
                or record.get('UserId')
                or record.get('UserIds')
                or ''
            )
            record_type = record.get('RecordType')
            result_status = get_safe_property(audit_data, 'ResultStatus')
            try:
                record_type_num = int(record_type) if record_type is not None else record_type
            except (ValueError, TypeError):
                record_type_num = record_type

            application_id = select_first_non_null(
                get_safe_property(audit_data, 'ApplicationId'),
                get_safe_property(audit_data, 'AppId'),
                get_safe_property(audit_data, 'ClientAppId'),
            )

            # DeviceProperties NV-pivot: only .OS and .BrowserType (matches M code GetNVProp)
            dev_props = get_safe_property(audit_data, 'DeviceProperties')
            dp_os = ''
            dp_browser = ''
            if dev_props and isinstance(dev_props, (list, tuple)):
                for dp in dev_props:
                    if isinstance(dp, dict):
                        if dp.get('Name') == 'OS':
                            dp_os = dp.get('Value', '')
                        elif dp.get('Name') == 'BrowserType':
                            dp_browser = dp.get('Value', '')

            # AgentCategory
            agent_id_val = get_safe_property(audit_data, 'AgentId')
            agent_cat = _categorize_agent(agent_id_val)

            # Associated admin units
            assoc_admin_units = (
                record.get('AssociatedAdminUnits')
                or get_safe_property(audit_data, 'AssociatedAdminUnits')
                or ''
            )
            assoc_admin_units_names = (
                record.get('AssociatedAdminUnitsNames')
                or get_safe_property(audit_data, 'AssociatedAdminUnitsNames')
                or ''
            )

            row_obj: dict[str, Any] = {
                'RecordId': record_id,
                'CreationDate': creation_date,
                'RecordType': record_type,
                'Operation': op_value,
                'UserId': uid_value,
                'AssociatedAdminUnits': assoc_admin_units,
                'AssociatedAdminUnitsNames': assoc_admin_units_names,
                '@odata.type': get_safe_property(audit_data, '@odata.type'),
                'CreationTime': creation_time,
                'Id': get_safe_property(audit_data, 'Id'),
                'OrganizationId': get_safe_property(audit_data, 'OrganizationId'),
                'ResultStatus': result_status,
                'UserKey': get_safe_property(audit_data, 'UserKey'),
                'UserType': get_safe_property(audit_data, 'UserType'),
                'Version': get_safe_property(audit_data, 'Version'),
                'Workload': get_safe_property(audit_data, 'Workload'),
                'ClientIP': get_safe_property(audit_data, 'ClientIP'),
                'ObjectId': get_safe_property(audit_data, 'ObjectId'),
                'AzureActiveDirectoryEventType': get_safe_property(audit_data, 'AzureActiveDirectoryEventType'),
                'ActorContextId': get_safe_property(audit_data, 'ActorContextId'),
                'ActorIpAddress': get_safe_property(audit_data, 'ActorIpAddress'),
                'InterSystemsId': get_safe_property(audit_data, 'InterSystemsId'),
                'IntraSystemId': get_safe_property(audit_data, 'IntraSystemId'),
                'SupportTicketId': get_safe_property(audit_data, 'SupportTicketId'),
                'TargetContextId': get_safe_property(audit_data, 'TargetContextId'),
                'ApplicationId': application_id,
                'DeviceProperties.OS': dp_os,
                'DeviceProperties.BrowserType': dp_browser,
                'ErrorNumber': get_safe_property(audit_data, 'ErrorNumber'),
                'SiteUrl': get_safe_property(audit_data, 'SiteUrl'),
                'SourceRelativeUrl': get_safe_property(audit_data, 'SourceRelativeUrl'),
                'SourceFileName': get_safe_property(audit_data, 'SourceFileName'),
                'SourceFileExtension': get_safe_property(audit_data, 'SourceFileExtension'),
                'ListId': get_safe_property(audit_data, 'ListId'),
                'ListItemUniqueId': get_safe_property(audit_data, 'ListItemUniqueId'),
                'WebId': get_safe_property(audit_data, 'WebId'),
                'ApplicationDisplayName': get_safe_property(audit_data, 'ApplicationDisplayName'),
                'EventSource': get_safe_property(audit_data, 'EventSource'),
                'ItemType': get_safe_property(audit_data, 'ItemType'),
                'SiteSensitivityLabelId': get_safe_property(audit_data, 'SiteSensitivityLabelId'),
                'GeoLocation': get_safe_property(audit_data, 'GeoLocation'),
                'IsManagedDevice': get_safe_property(audit_data, 'IsManagedDevice'),
                'DeviceDisplayName': get_safe_property(audit_data, 'DeviceDisplayName'),
                'ListBaseType': get_safe_property(audit_data, 'ListBaseType'),
                'ListServerTemplate': get_safe_property(audit_data, 'ListServerTemplate'),
                'AuthenticationType': get_safe_property(audit_data, 'AuthenticationType'),
                'Site': get_safe_property(audit_data, 'Site'),
                'DoNotDistributeEvent': get_safe_property(audit_data, 'DoNotDistributeEvent'),
                'HighPriorityMediaProcessing': get_safe_property(audit_data, 'HighPriorityMediaProcessing'),
                'BrowserName': get_safe_property(audit_data, 'BrowserName'),
                'BrowserVersion': get_safe_property(audit_data, 'BrowserVersion'),
                'CorrelationId': get_safe_property(audit_data, 'CorrelationId'),
                'Platform': get_safe_property(audit_data, 'Platform'),
                'UserAgent': get_safe_property(audit_data, 'UserAgent'),
                'ActorInfoString': get_safe_property(audit_data, 'ActorInfoString'),
                'AppId': get_safe_property(audit_data, 'AppId'),
                'AuthType': get_safe_property(audit_data, 'AuthType'),
                'ClientAppId': get_safe_property(audit_data, 'ClientAppId'),
                'ClientIPAddress': get_safe_property(audit_data, 'ClientIPAddress'),
                'ClientInfoString': get_safe_property(audit_data, 'ClientInfoString'),
                'ExternalAccess': get_safe_property(audit_data, 'ExternalAccess'),
                'InternalLogonType': get_safe_property(audit_data, 'InternalLogonType'),
                'LogonType': get_safe_property(audit_data, 'LogonType'),
                'LogonUserSid': get_safe_property(audit_data, 'LogonUserSid'),
                'MailboxGuid': get_safe_property(audit_data, 'MailboxGuid'),
                'MailboxOwnerSid': get_safe_property(audit_data, 'MailboxOwnerSid'),
                'MailboxOwnerUPN': get_safe_property(audit_data, 'MailboxOwnerUPN'),
                'OrganizationName': get_safe_property(audit_data, 'OrganizationName'),
                'OriginatingServer': get_safe_property(audit_data, 'OriginatingServer'),
                'SessionId': get_safe_property(audit_data, 'SessionId'),
                'TokenObjectId': get_safe_property(audit_data, 'TokenObjectId'),
                'TokenTenantId': get_safe_property(audit_data, 'TokenTenantId'),
                'TokenType': get_safe_property(audit_data, 'TokenType'),
                'SaveToSentItems': get_safe_property(audit_data, 'SaveToSentItems'),
                'OperationCount': get_safe_property(audit_data, 'OperationCount'),
                'FileSizeBytes': get_safe_property(audit_data, 'FileSizeBytes'),
                'MeetingId': get_safe_property(audit_data, 'MeetingId'),
                'MeetingType': get_safe_property(audit_data, 'MeetingType'),
                'EventSignature': get_safe_property(audit_data, 'EventSignature'),
                'EventData': get_safe_property(audit_data, 'EventData'),
                'Permission': get_safe_property(audit_data, 'Permission'),
                'SensitivityLabelId': get_safe_property(audit_data, 'SensitivityLabelId'),
                'SharingLinkScope': get_safe_property(audit_data, 'SharingLinkScope'),
                'TargetUserOrGroupType': get_safe_property(audit_data, 'TargetUserOrGroupType'),
                'TargetUserOrGroupName': get_safe_property(audit_data, 'TargetUserOrGroupName'),
                'MeetingURL': get_safe_property(audit_data, 'MeetingURL'),
                'ChatId': get_safe_property(audit_data, 'ChatId'),
                'MessageId': get_safe_property(audit_data, 'MessageId'),
                'MessageSizeInBytes': get_safe_property(audit_data, 'MessageSizeInBytes'),
                'MessageType': get_safe_property(audit_data, 'MessageType'),
                'FormId': get_safe_property(audit_data, 'FormId'),
                'FormName': get_safe_property(audit_data, 'FormName'),
                'VideoId': get_safe_property(audit_data, 'VideoId'),
                'VideoName': get_safe_property(audit_data, 'VideoName'),
                'ChannelId': get_safe_property(audit_data, 'ChannelId'),
                'ViewDuration': get_safe_property(audit_data, 'ViewDuration'),
                'ClientRegion': get_safe_property(audit_data, 'ClientRegion'),
                'CopilotLogVersion': get_safe_property(audit_data, 'CopilotLogVersion'),
                'TargetId': get_safe_property(audit_data, 'TargetId'),
                'TeamName': get_safe_property(audit_data, 'TeamName'),
                'TeamGuid': get_safe_property(audit_data, 'TeamGuid'),
                'ResponseId': get_safe_property(audit_data, 'ResponseId'),
                'IsAnonymous': get_safe_property(audit_data, 'IsAnonymous'),
                'DeviceType': get_safe_property(audit_data, 'DeviceType'),
                'ChannelName': get_safe_property(audit_data, 'ChannelName'),
                'ChannelGuid': get_safe_property(audit_data, 'ChannelGuid'),
                'ChannelType': get_safe_property(audit_data, 'ChannelType'),
                'AppName': get_safe_property(audit_data, 'AppName'),
                'EnvironmentName': get_safe_property(audit_data, 'EnvironmentName'),
                'PlanId': get_safe_property(audit_data, 'PlanId'),
                'PlanName': get_safe_property(audit_data, 'PlanName'),
                'TaskId': get_safe_property(audit_data, 'TaskId'),
                'TaskName': get_safe_property(audit_data, 'TaskName'),
                'PercentComplete': get_safe_property(audit_data, 'PercentComplete'),
                'CrossMailboxOperation': get_safe_property(audit_data, 'CrossMailboxOperation'),
                'RecordTypeNum': record_type_num,
                'ResultStatus_Audit': result_status,
                'ModelId': get_safe_property(audit_data, 'ModelId'),
                'ModelProvider': get_safe_property(audit_data, 'ModelProvider'),
                'ModelFamily': get_safe_property(audit_data, 'ModelFamily'),
                'TokensTotal': get_safe_property(audit_data, 'TokensTotal'),
                'TokensInput': get_safe_property(audit_data, 'TokensInput'),
                'TokensOutput': get_safe_property(audit_data, 'TokensOutput'),
                'DurationMs': get_safe_property(audit_data, 'DurationMs'),
                'OutcomeStatus': get_safe_property(audit_data, 'OutcomeStatus'),
                'ConversationId': get_safe_property(audit_data, 'ConversationId'),
                'TurnNumber': get_safe_property(audit_data, 'TurnNumber'),
                'RetryCount': get_safe_property(audit_data, 'RetryCount'),
                'ClientVersion': get_safe_property(audit_data, 'ClientVersion'),
                'ClientPlatform': get_safe_property(audit_data, 'ClientPlatform'),
                'AgentId': agent_id_val,
                'AgentName': get_safe_property(audit_data, 'AgentName'),
                'AgentVersion': get_safe_property(audit_data, 'AgentVersion'),
                'AgentCategory': agent_cat,
                'ApplicationName': get_safe_property(audit_data, 'ApplicationName'),
                'SensitivityLabel': get_safe_property(audit_data, 'SensitivityLabel'),
                # CED sub-fields — empty for non-Copilot records
                'AppHost': '',
                'ThreadId': '',
                'Context_Id': '',
                'Context_Type': '',
                'Message_Id': '',
                'Message_isPrompt': '',
                'AccessedResource_Action': '',
                'AccessedResource_PolicyDetails': '',
                'AccessedResource_SiteUrl': '',
                'AISystemPlugin_Id': '',
                'AISystemPlugin_Name': '',
                'ModelTransparencyDetails_ModelName': '',
                'MessageIds': '',
                'AccessedResource_Name': '',
                'AccessedResource_SensitivityLabel': '',
                'AccessedResource_ResourceType': '',
                'Context_Item': '',
            }

            if deep:
                # Deep flatten entire AuditData for each row (no raw JSON)
                flat_audit = convert_to_flat_columns(audit_data, prefix='', max_depth=flat_depth_deep)
                for k, v in flat_audit.items():
                    if k not in row_obj:
                        row_obj[k] = v

            return [row_obj]

        # ── Copilot path: array explosion ──
        messages = get_array_fast(ced, 'Messages')

        # Prompt filtering
        if prompt_filter_value:
            filtered_messages: list = []
            if prompt_filter_value == 'Null':
                for msg in messages:
                    if isinstance(msg, dict) and msg.get('isPrompt') is None:
                        filtered_messages.append(msg)
                    elif not isinstance(msg, dict):
                        filtered_messages.append(msg)
            elif prompt_filter_value == 'Both':
                for msg in messages:
                    if isinstance(msg, dict) and msg.get('isPrompt') is not None:
                        filtered_messages.append(msg)
            else:
                target_value = (prompt_filter_value == 'Prompt')
                for msg in messages:
                    if isinstance(msg, dict):
                        try:
                            if msg.get('isPrompt') == target_value:
                                filtered_messages.append(msg)
                        except Exception:
                            pass
            messages = filtered_messages
            if len(messages) == 0:
                if not skip_metrics:
                    metrics['FilteringSkippedRecords'] = metrics.get('FilteringSkippedRecords', 0) + 1
                    metrics['FilteringPromptFiltered'] = metrics.get('FilteringPromptFiltered', 0) + 1
                return []

        contexts = get_array_fast(ced, 'Contexts')
        resources = get_array_fast(ced, 'AccessedResources')
        plugins_raw = get_array_fast(ced, 'AISystemPlugin')
        model_det_raw = get_array_fast(ced, 'ModelTransparencyDetails')
        message_ids = get_array_fast(ced, 'MessageIds')

        # DSPM for AI: Extract SensitivityLabels array
        sensitivity_labels = get_array_fast(ced, 'SensitivityLabels')

        # DSPM for AI: Determine activity type for conditional 2-level explosion
        activity_type = get_safe_property(audit_data, 'Operation')

        # DSPM for AI: Extract 2nd-level arrays (for full explosion mode)
        plugins = None
        recording_sessions = None
        context_items_count = None

        # Need app_identity_raw for plugin extraction
        app_identity_raw = select_first_non_null(
            get_safe_property(audit_data, 'AppIdentity'),
            get_safe_property(ced, 'AppIdentity'),
        )

        if not partial_explode:
            # Full explosion mode: Extract 2nd-level arrays for row count calculation
            if activity_type == 'ConnectedAIAppInteraction' and app_identity_raw:
                if isinstance(app_identity_raw, dict):
                    plugins = get_array_fast(app_identity_raw, 'Plugins')

            if activity_type == 'CopilotInteraction' and len(contexts) > 0:
                # Find max Items[] count across all Contexts
                max_items_count = 0
                for ctx in contexts:
                    if ctx and isinstance(ctx, dict):
                        items = get_array_fast(ctx, 'Items')
                        if items and len(items) > max_items_count:
                            max_items_count = len(items)
                if max_items_count > 0:
                    context_items_count = max_items_count

        # Calculate row count
        if prompt_filter_value:
            row_count = max(1, len(messages))
        else:
            # DSPM for AI: Include all arrays in row count calculation
            array_counts = [
                1, len(messages), len(contexts), len(resources),
                len(sensitivity_labels), len(plugins_raw), len(model_det_raw)
            ]

            # Full explosion: include 2nd-level arrays in row count
            if not partial_explode:
                if plugins:
                    array_counts.append(len(plugins))
                if recording_sessions:
                    array_counts.append(len(recording_sessions))
                if context_items_count:
                    array_counts.append(context_items_count)

            row_count = max(array_counts)

        # Extract common fields
        creation_date = format_date_purview_fast(record.get('CreationDate'))
        creation_time = format_date_purview_fast(get_safe_property(audit_data, 'CreationTime'))

        application_id = select_first_non_null(
            get_safe_property(audit_data, 'ApplicationId'),
            get_safe_property(audit_data, 'AppId'),
            get_safe_property(audit_data, 'ClientAppId'),
        )

        # DeviceProperties NV-pivot
        dev_props = get_safe_property(audit_data, 'DeviceProperties')
        dp_os = ''
        dp_browser = ''
        if dev_props and isinstance(dev_props, (list, tuple)):
            for dp in dev_props:
                if isinstance(dp, dict):
                    if dp.get('Name') == 'OS':
                        dp_os = dp.get('Value', '')
                    elif dp.get('Name') == 'BrowserType':
                        dp_browser = dp.get('Value', '')

        app_host = select_first_non_null(
            get_safe_property(ced, 'AppHost'),
            get_safe_property(audit_data, 'AppHost'),
            get_safe_property(audit_data, 'Workload'),
        )
        client_region = get_safe_property(audit_data, 'ClientRegion')
        agent_id = get_safe_property(audit_data, 'AgentId')
        agent_name = get_safe_property(audit_data, 'AgentName')
        agent_version = select_first_non_null(
            get_safe_property(audit_data, 'AgentVersion'),
            get_safe_property(ced, 'AgentVersion'),
            get_safe_property(ced, 'Version'),
        )
        agent_category = _categorize_agent(agent_id)

        thread_id = get_safe_property(ced, 'ThreadId')
        audit_user_key = get_safe_property(audit_data, 'UserKey')
        client_ip = get_safe_property(audit_data, 'ClientIP')
        organization_id = get_safe_property(audit_data, 'OrganizationId')
        version = get_safe_property(audit_data, 'Version')
        user_type = get_safe_property(audit_data, 'UserType')
        copilot_log_version = get_safe_property(audit_data, 'CopilotLogVersion')
        workload = get_safe_property(audit_data, 'Workload')

        # Extract fields to match ExplodeArrays output for Power BI compatibility
        audit_data_id = get_safe_property(audit_data, 'Id')
        record_type_num = get_safe_property(audit_data, 'RecordType')
        result_status_audit = get_safe_property(audit_data, 'ResultStatus')
        app_id = get_safe_property(audit_data, 'AppId')
        client_app_id = get_safe_property(audit_data, 'ClientAppId')
        correlation_id = get_safe_property(audit_data, 'CorrelationId')

        # Model and token fields
        model_id = select_first_non_null(
            get_safe_property(ced, 'ModelId'),
            get_safe_property(ced, 'ModelID'),
            get_safe_property(audit_data, 'ModelId'),
        )
        model_provider = select_first_non_null(
            get_safe_property(ced, 'ModelProvider'),
            get_safe_property(ced, 'Provider'),
            get_safe_property(ced, 'ModelVendor'),
        )
        model_family = select_first_non_null(
            get_safe_property(ced, 'ModelFamily'),
            get_safe_property(ced, 'ModelType'),
        )

        usage_node = select_first_non_null(
            get_safe_property(ced, 'Usage'),
            get_safe_property(ced, 'TokenUsage'),
            get_safe_property(ced, 'Tokens'),
            get_safe_property(audit_data, 'Usage'),
        )
        tokens_total: Optional[float] = None
        tokens_input: Optional[float] = None
        tokens_output: Optional[float] = None
        if usage_node and isinstance(usage_node, dict):
            tokens_total = _get_num(select_first_non_null(
                get_safe_property(usage_node, 'Total'),
                get_safe_property(usage_node, 'TotalTokens'),
                get_safe_property(usage_node, 'TokensTotal'),
            ))
            tokens_input = _get_num(select_first_non_null(
                get_safe_property(usage_node, 'Input'),
                get_safe_property(usage_node, 'Prompt'),
                get_safe_property(usage_node, 'InputTokens'),
                get_safe_property(usage_node, 'TokensInput'),
            ))
            tokens_output = _get_num(select_first_non_null(
                get_safe_property(usage_node, 'Output'),
                get_safe_property(usage_node, 'Completion'),
                get_safe_property(usage_node, 'OutputTokens'),
                get_safe_property(usage_node, 'TokensOutput'),
            ))
        if not tokens_total and (tokens_input or tokens_output):
            try:
                tokens_total = (tokens_input or 0) + (tokens_output or 0)
            except Exception:
                pass

        # Duration, outcome, conversation fields
        duration_ms = _get_num(select_first_non_null(
            get_safe_property(ced, 'DurationMs'),
            get_safe_property(ced, 'ElapsedMs'),
            get_safe_property(ced, 'ProcessingTimeMs'),
            get_safe_property(ced, 'LatencyMs'),
        ))
        outcome_status = select_first_non_null(
            get_safe_property(ced, 'OutcomeStatus'),
            get_safe_property(ced, 'Outcome'),
            get_safe_property(ced, 'Result'),
            get_safe_property(ced, 'Status'),
        )
        if isinstance(outcome_status, bool):
            outcome_status = 'Success' if outcome_status else 'Failure'
        conversation_id = select_first_non_null(
            get_safe_property(ced, 'ConversationId'),
            get_safe_property(ced, 'ConversationID'),
            get_safe_property(ced, 'SessionId'),
        )
        turn_number = _get_num(select_first_non_null(
            get_safe_property(ced, 'TurnNumber'),
            get_safe_property(ced, 'TurnIndex'),
            get_safe_property(ced, 'MessageIndex'),
        ))
        retry_count = _get_num(select_first_non_null(
            get_safe_property(ced, 'RetryCount'),
            get_safe_property(ced, 'Retries'),
        ))
        client_version = select_first_non_null(
            get_safe_property(ced, 'ClientVersion'),
            get_safe_property(ced, 'Version'),
            get_safe_property(ced, 'Build'),
        )
        client_platform = select_first_non_null(
            get_safe_property(ced, 'ClientPlatform'),
            get_safe_property(ced, 'Platform'),
            get_safe_property(ced, 'OS'),
        )

        # Associated admin units
        assoc_admin_units = (
            record.get('AssociatedAdminUnits')
            or get_safe_property(audit_data, 'AssociatedAdminUnits')
            or ''
        )
        assoc_admin_units_names = (
            record.get('AssociatedAdminUnitsNames')
            or get_safe_property(audit_data, 'AssociatedAdminUnitsNames')
            or ''
        )

        base_set = set(PURVIEW_EXPLODED_HEADER)
        rows: list[dict[str, Any]] = []

        for i in range(row_count):
            # Message fields
            msg_id = ''
            msg_is_prompt = ''
            if i < len(messages):
                msg = messages[i]
                if isinstance(msg, dict):
                    msg_id = get_safe_property(msg, 'Id') or ''
                    msg_is_prompt = bool_tf_fast(get_safe_property(msg, 'isPrompt'))
                else:
                    msg_id = msg if msg else ''

            # Context fields
            ctx_id = ''
            ctx_type = ''
            if i < len(contexts) and contexts[i] and isinstance(contexts[i], dict):
                ctx_id = get_safe_property(contexts[i], 'Id') or ''
                ctx_type = get_safe_property(contexts[i], 'Type') or ''

            # Resource fields
            res_action = ''
            res_policy_details = ''
            res_site_url = ''
            res_name = ''
            res_sensitivity_label = ''
            res_resource_type = ''
            if i < len(resources) and resources[i] and isinstance(resources[i], dict):
                res_action = get_safe_property(resources[i], 'Action') or ''
                res_policy_details = to_json_if_object_fast(get_safe_property(resources[i], 'PolicyDetails'))
                res_site_url = get_safe_property(resources[i], 'SiteUrl') or ''
                res_name = get_safe_property(resources[i], 'Name') or ''
                res_sensitivity_label = get_safe_property(resources[i], 'SensitivityLabel') or ''
                res_resource_type = get_safe_property(resources[i], 'ResourceType') or ''

            # Plugin fields
            plugin_id = ''
            plugin_name = ''
            if i < len(plugins_raw) and plugins_raw[i] and isinstance(plugins_raw[i], dict):
                plugin_id = get_safe_property(plugins_raw[i], 'Id') or ''
                plugin_name = get_safe_property(plugins_raw[i], 'Name') or ''

            # Model transparency details
            model_name = ''
            if i < len(model_det_raw) and model_det_raw[i] and isinstance(model_det_raw[i], dict):
                model_name = get_safe_property(model_det_raw[i], 'ModelName') or ''

            # MessageIds - semicolon-joined
            message_ids_str = ';'.join(str(mid) for mid in message_ids) if message_ids else ''

            # SensitivityLabel (indexed)
            sensitivity_label = ''
            if i < len(sensitivity_labels):
                try:
                    sensitivity_label = str(sensitivity_labels[i]) if sensitivity_labels[i] is not None else ''
                except Exception:
                    sensitivity_label = ''

            # Context_Item
            context_item = ''
            if activity_type == 'CopilotInteraction':
                if partial_explode:
                    if i < len(contexts) and contexts[i] and isinstance(contexts[i], dict):
                        try:
                            items = get_array_fast(contexts[i], 'Items')
                            if items and len(items) > 0:
                                context_item = ';'.join(
                                    to_json_if_object_fast(item) or ''
                                    for item in items
                                )
                        except Exception:
                            context_item = ''
                else:
                    try:
                        found_item = None
                        for ctx in contexts:
                            if ctx and isinstance(ctx, dict):
                                items = get_array_fast(ctx, 'Items')
                                if items and i < len(items):
                                    found_item = items[i]
                                    break
                        if found_item:
                            context_item = to_json_if_object_fast(found_item) or ''
                    except Exception:
                        context_item = ''

            row_obj = {
                'RecordId': record.get('RecordId') or record.get('Identity') or record.get('Id') or audit_data_id,
                'CreationDate': creation_date,
                'RecordType': record.get('RecordType'),
                'Operation': get_safe_property(audit_data, 'Operation'),
                'UserId': get_safe_property(audit_data, 'UserId'),
                'AssociatedAdminUnits': assoc_admin_units,
                'AssociatedAdminUnitsNames': assoc_admin_units_names,
                '@odata.type': get_safe_property(audit_data, '@odata.type'),
                'CreationTime': creation_time,
                'Id': audit_data_id,
                'OrganizationId': organization_id,
                'ResultStatus': result_status_audit,
                'UserKey': audit_user_key,
                'UserType': user_type,
                'Version': version,
                'Workload': workload,
                'ClientIP': client_ip,
                'ObjectId': get_safe_property(audit_data, 'ObjectId'),
                'AzureActiveDirectoryEventType': get_safe_property(audit_data, 'AzureActiveDirectoryEventType'),
                'ActorContextId': get_safe_property(audit_data, 'ActorContextId'),
                'ActorIpAddress': get_safe_property(audit_data, 'ActorIpAddress'),
                'InterSystemsId': get_safe_property(audit_data, 'InterSystemsId'),
                'IntraSystemId': get_safe_property(audit_data, 'IntraSystemId'),
                'SupportTicketId': get_safe_property(audit_data, 'SupportTicketId'),
                'TargetContextId': get_safe_property(audit_data, 'TargetContextId'),
                'ApplicationId': application_id,
                'DeviceProperties.OS': dp_os,
                'DeviceProperties.BrowserType': dp_browser,
                'ErrorNumber': get_safe_property(audit_data, 'ErrorNumber'),
                'SiteUrl': get_safe_property(audit_data, 'SiteUrl'),
                'SourceRelativeUrl': get_safe_property(audit_data, 'SourceRelativeUrl'),
                'SourceFileName': get_safe_property(audit_data, 'SourceFileName'),
                'SourceFileExtension': get_safe_property(audit_data, 'SourceFileExtension'),
                'ListId': get_safe_property(audit_data, 'ListId'),
                'ListItemUniqueId': get_safe_property(audit_data, 'ListItemUniqueId'),
                'WebId': get_safe_property(audit_data, 'WebId'),
                'ApplicationDisplayName': get_safe_property(audit_data, 'ApplicationDisplayName'),
                'EventSource': get_safe_property(audit_data, 'EventSource'),
                'ItemType': get_safe_property(audit_data, 'ItemType'),
                'SiteSensitivityLabelId': get_safe_property(audit_data, 'SiteSensitivityLabelId'),
                'GeoLocation': get_safe_property(audit_data, 'GeoLocation'),
                'IsManagedDevice': get_safe_property(audit_data, 'IsManagedDevice'),
                'DeviceDisplayName': get_safe_property(audit_data, 'DeviceDisplayName'),
                'ListBaseType': get_safe_property(audit_data, 'ListBaseType'),
                'ListServerTemplate': get_safe_property(audit_data, 'ListServerTemplate'),
                'AuthenticationType': get_safe_property(audit_data, 'AuthenticationType'),
                'Site': get_safe_property(audit_data, 'Site'),
                'DoNotDistributeEvent': get_safe_property(audit_data, 'DoNotDistributeEvent'),
                'HighPriorityMediaProcessing': get_safe_property(audit_data, 'HighPriorityMediaProcessing'),
                'BrowserName': get_safe_property(audit_data, 'BrowserName'),
                'BrowserVersion': get_safe_property(audit_data, 'BrowserVersion'),
                'CorrelationId': correlation_id,
                'Platform': get_safe_property(audit_data, 'Platform'),
                'UserAgent': get_safe_property(audit_data, 'UserAgent'),
                'ActorInfoString': get_safe_property(audit_data, 'ActorInfoString'),
                'AppId': app_id,
                'AuthType': get_safe_property(audit_data, 'AuthType'),
                'ClientAppId': client_app_id,
                'ClientIPAddress': get_safe_property(audit_data, 'ClientIPAddress'),
                'ClientInfoString': get_safe_property(audit_data, 'ClientInfoString'),
                'ExternalAccess': get_safe_property(audit_data, 'ExternalAccess'),
                'InternalLogonType': get_safe_property(audit_data, 'InternalLogonType'),
                'LogonType': get_safe_property(audit_data, 'LogonType'),
                'LogonUserSid': get_safe_property(audit_data, 'LogonUserSid'),
                'MailboxGuid': get_safe_property(audit_data, 'MailboxGuid'),
                'MailboxOwnerSid': get_safe_property(audit_data, 'MailboxOwnerSid'),
                'MailboxOwnerUPN': get_safe_property(audit_data, 'MailboxOwnerUPN'),
                'OrganizationName': get_safe_property(audit_data, 'OrganizationName'),
                'OriginatingServer': get_safe_property(audit_data, 'OriginatingServer'),
                'SessionId': get_safe_property(audit_data, 'SessionId'),
                'TokenObjectId': get_safe_property(audit_data, 'TokenObjectId'),
                'TokenTenantId': get_safe_property(audit_data, 'TokenTenantId'),
                'TokenType': get_safe_property(audit_data, 'TokenType'),
                'SaveToSentItems': get_safe_property(audit_data, 'SaveToSentItems'),
                'OperationCount': get_safe_property(audit_data, 'OperationCount'),
                'FileSizeBytes': get_safe_property(audit_data, 'FileSizeBytes'),
                'MeetingId': get_safe_property(audit_data, 'MeetingId'),
                'MeetingType': get_safe_property(audit_data, 'MeetingType'),
                'EventSignature': get_safe_property(audit_data, 'EventSignature'),
                'EventData': get_safe_property(audit_data, 'EventData'),
                'Permission': get_safe_property(audit_data, 'Permission'),
                'SensitivityLabelId': get_safe_property(audit_data, 'SensitivityLabelId'),
                'SharingLinkScope': get_safe_property(audit_data, 'SharingLinkScope'),
                'TargetUserOrGroupType': get_safe_property(audit_data, 'TargetUserOrGroupType'),
                'TargetUserOrGroupName': get_safe_property(audit_data, 'TargetUserOrGroupName'),
                'MeetingURL': get_safe_property(audit_data, 'MeetingURL'),
                'ChatId': get_safe_property(audit_data, 'ChatId'),
                'MessageId': get_safe_property(audit_data, 'MessageId'),
                'MessageSizeInBytes': get_safe_property(audit_data, 'MessageSizeInBytes'),
                'MessageType': get_safe_property(audit_data, 'MessageType'),
                'FormId': get_safe_property(audit_data, 'FormId'),
                'FormName': get_safe_property(audit_data, 'FormName'),
                'VideoId': get_safe_property(audit_data, 'VideoId'),
                'VideoName': get_safe_property(audit_data, 'VideoName'),
                'ChannelId': get_safe_property(audit_data, 'ChannelId'),
                'ViewDuration': get_safe_property(audit_data, 'ViewDuration'),
                'ClientRegion': client_region,
                'CopilotLogVersion': copilot_log_version,
                'TargetId': get_safe_property(audit_data, 'TargetId'),
                'TeamName': get_safe_property(audit_data, 'TeamName'),
                'TeamGuid': get_safe_property(audit_data, 'TeamGuid'),
                'ResponseId': get_safe_property(audit_data, 'ResponseId'),
                'IsAnonymous': get_safe_property(audit_data, 'IsAnonymous'),
                'DeviceType': get_safe_property(audit_data, 'DeviceType'),
                'ChannelName': get_safe_property(audit_data, 'ChannelName'),
                'ChannelGuid': get_safe_property(audit_data, 'ChannelGuid'),
                'ChannelType': get_safe_property(audit_data, 'ChannelType'),
                'AppName': get_safe_property(audit_data, 'AppName'),
                'EnvironmentName': get_safe_property(audit_data, 'EnvironmentName'),
                'PlanId': get_safe_property(audit_data, 'PlanId'),
                'PlanName': get_safe_property(audit_data, 'PlanName'),
                'TaskId': get_safe_property(audit_data, 'TaskId'),
                'TaskName': get_safe_property(audit_data, 'TaskName'),
                'PercentComplete': get_safe_property(audit_data, 'PercentComplete'),
                'CrossMailboxOperation': get_safe_property(audit_data, 'CrossMailboxOperation'),
                'RecordTypeNum': record_type_num,
                'ResultStatus_Audit': result_status_audit,
                'ModelId': model_id,
                'ModelProvider': model_provider,
                'ModelFamily': model_family,
                'TokensTotal': tokens_total,
                'TokensInput': tokens_input,
                'TokensOutput': tokens_output,
                'DurationMs': duration_ms,
                'OutcomeStatus': outcome_status,
                'ConversationId': conversation_id,
                'TurnNumber': turn_number,
                'RetryCount': retry_count,
                'ClientVersion': client_version,
                'ClientPlatform': client_platform,
                'AgentId': agent_id,
                'AgentName': agent_name,
                'AgentVersion': agent_version,
                'AgentCategory': agent_category,
                'ApplicationName': get_safe_property(audit_data, 'ApplicationName'),
                'SensitivityLabel': sensitivity_label,
                'AppHost': app_host,
                'ThreadId': thread_id,
                'Context_Id': ctx_id,
                'Context_Type': ctx_type,
                'Message_Id': msg_id,
                'Message_isPrompt': msg_is_prompt,
                'AccessedResource_Action': res_action,
                'AccessedResource_PolicyDetails': res_policy_details,
                'AccessedResource_SiteUrl': res_site_url,
                'AISystemPlugin_Id': plugin_id,
                'AISystemPlugin_Name': plugin_name,
                'ModelTransparencyDetails_ModelName': model_name,
                'MessageIds': message_ids_str,
                'AccessedResource_Name': res_name,
                'AccessedResource_SensitivityLabel': res_sensitivity_label,
                'AccessedResource_ResourceType': res_resource_type,
                'Context_Item': context_item,
            }

            # Partial explosion mode: Preserve AuditData column
            if partial_explode:
                row_obj['AuditData'] = record.get('AuditData', '')

            # DSPM for AI: 2-level explosion for ConnectedAIAppInteraction (AppIdentity.Plugins[])
            if activity_type == 'ConnectedAIAppInteraction' and plugins:
                if partial_explode:
                    # Partial mode: Semi-colon-joined JSON for all plugins
                    plugins_list = ';'.join(
                        to_json_if_object_fast(p) or '' for p in plugins
                    )
                    if 'AppIdentity_Plugins' not in row_obj:
                        row_obj['AppIdentity_Plugins'] = plugins_list
                        if 'AppIdentity_Plugins' not in deep_extra_columns:
                            deep_extra_columns.append('AppIdentity_Plugins')
                else:
                    # Full mode: One plugin per row
                    if i < len(plugins):
                        plugin_json = to_json_if_object_fast(plugins[i]) or ''
                        if 'AppIdentity_Plugin' not in row_obj:
                            row_obj['AppIdentity_Plugin'] = plugin_json
                            if 'AppIdentity_Plugin' not in deep_extra_columns:
                                deep_extra_columns.append('AppIdentity_Plugin')

            if deep:
                if ced:
                    flat = convert_to_flat_columns(ced, prefix='', max_depth=flat_depth_deep)
                    for k, v in flat.items():
                        if k in base_set:
                            continue
                        if k not in row_obj:
                            if k not in deep_extra_columns:
                                deep_extra_columns.append(k)
                            row_obj[k] = v

                if audit_data:
                    # Clone audit_data without CopilotEventData
                    audit_data_clone = {
                        k: v for k, v in audit_data.items()
                        if k != 'CopilotEventData'
                    }
                    flat_audit = convert_to_flat_columns(audit_data_clone, prefix='', max_depth=flat_depth_deep)
                    for k, v in flat_audit.items():
                        if k in base_set:
                            continue
                        if k not in row_obj:
                            if k not in deep_extra_columns:
                                deep_extra_columns.append(k)
                            row_obj[k] = v

            rows.append(row_obj)

        # Update metrics for explosion events
        if not skip_metrics and len(rows) > 1:
            metrics['ExplosionEvents'] = metrics.get('ExplosionEvents', 0) + 1
            metrics['ExplosionRowsFromEvents'] = metrics.get('ExplosionRowsFromEvents', 0) + (len(rows) - 1)
            current_max = metrics.get('ExplosionMaxPerRecord', 0)
            if len(rows) > current_max:
                metrics['ExplosionMaxPerRecord'] = len(rows)

        return rows

    except Exception:
        if not skip_metrics:
            metrics['FilteringSkippedRecords'] = metrics.get('FilteringSkippedRecords', 0) + 1
            metrics['FilteringParseFailures'] = metrics.get('FilteringParseFailures', 0) + 1
        return []


# ===========================================================================
# 4. convert_to_structured_record — PS Convert-ToStructuredRecord (L15869-16082)
# ===========================================================================

def convert_to_structured_record(
    record: dict[str, Any],
    enable_explosion: bool = False,
    explode_deep: bool = False,
    flat_depth_standard: int = FLAT_DEPTH_STANDARD,
    flat_depth_deep_val: int = FLAT_DEPTH_DEEP,
    explosion_per_record_row_cap: int = EXPLOSION_PER_RECORD_ROW_CAP,
    json_depth: int = JSON_DEPTH,
    metrics: Optional[dict[str, Any]] = None,
) -> list[dict[str, Any]]:
    """
    Structured record builder — PS Convert-ToStructuredRecord (L15869-16082).

    In non-explosion mode: Returns 8-column compact record matching Purview UI export schema.
    In explosion mode: Extracts and flattens all fields from AuditData, optionally
    explodes arrays (Suggestions, Actions, References, Participants) into multiple rows.

    Args:
        record: Raw audit record dict.
        enable_explosion: If True, explode arrays into separate rows.
        explode_deep: If True, deep-flatten CED into extra columns.
        flat_depth_standard: Standard flattening depth (default 6).
        flat_depth_deep_val: Deep flattening depth (default 120).
        explosion_per_record_row_cap: Max rows per record (default 1000).
        json_depth: JSON serialization depth (default 60).
        metrics: Shared metrics dict (modified in-place).

    Returns:
        List of record dicts (possibly empty on parse failure).
    """
    if metrics is None:
        metrics = {}

    try:
        # Parse AuditData
        audit_data = record.get('_ParsedAuditData')
        if audit_data is None:
            raw_audit = record.get('AuditData')
            if raw_audit and isinstance(raw_audit, str):
                try:
                    audit_data = json.loads(raw_audit)
                except (json.JSONDecodeError, ValueError):
                    audit_data = None
            elif isinstance(raw_audit, dict):
                audit_data = raw_audit

        if not audit_data:
            metrics['FilteringSkippedRecords'] = metrics.get('FilteringSkippedRecords', 0) + 1
            metrics['FilteringMissingAuditData'] = metrics.get('FilteringMissingAuditData', 0) + 1
            return []

        # NON-EXPLOSION MODE: Return 8-column compact record matching Purview UI export schema
        if not enable_explosion and not explode_deep:
            op_value = (
                get_safe_property(audit_data, 'Operation')
                or record.get('Operation')
                or record.get('Operations')
                or ''
            )
            user_id = record.get('UserId') or record.get('UserIds') or ''

            # Format creation date
            creation_date_raw = record.get('CreationDate')
            if creation_date_raw:
                parsed_dt = parse_date_safe(creation_date_raw)
                creation_date = parsed_dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{parsed_dt.microsecond // 1000:03d}Z" if parsed_dt else str(creation_date_raw)
            else:
                creation_date = ''

            compact_record: dict[str, Any] = {
                'RecordId': record.get('RecordId') or record.get('Identity') or record.get('Id') or get_safe_property(audit_data, 'Id'),
                'CreationDate': creation_date,
                'RecordType': record.get('RecordType'),
                'Operation': op_value,
                'UserId': user_id,
                'AuditData': record.get('AuditData', ''),
                'AssociatedAdminUnits': (
                    get_safe_property(audit_data, 'AssociatedAdminUnits')
                    or record.get('AssociatedAdminUnits')
                    or ''
                ),
                'AssociatedAdminUnitsNames': (
                    get_safe_property(audit_data, 'AssociatedAdminUnitsNames')
                    or record.get('AssociatedAdminUnitsNames')
                    or ''
                ),
            }
            return [compact_record]

        # EXPLOSION MODE: Extract and flatten all fields from AuditData
        ced = get_safe_property(audit_data, 'CopilotEventData')

        model_id = select_first_non_null(
            get_safe_property(ced, 'ModelId') if ced else None,
            get_safe_property(ced, 'ModelID') if ced else None,
            get_safe_property(audit_data, 'ModelId'),
        )
        model_provider = select_first_non_null(
            get_safe_property(ced, 'ModelProvider') if ced else None,
            get_safe_property(ced, 'Provider') if ced else None,
            get_safe_property(ced, 'ModelVendor') if ced else None,
        )
        model_family = select_first_non_null(
            get_safe_property(ced, 'ModelFamily') if ced else None,
            get_safe_property(ced, 'ModelType') if ced else None,
        )

        usage_node = select_first_non_null(
            get_safe_property(ced, 'Usage') if ced else None,
            get_safe_property(ced, 'TokenUsage') if ced else None,
            get_safe_property(ced, 'Tokens') if ced else None,
            get_safe_property(audit_data, 'Usage'),
        )
        tokens_total: Optional[float] = None
        tokens_input: Optional[float] = None
        tokens_output: Optional[float] = None
        if usage_node and isinstance(usage_node, dict):
            tokens_total = _get_num(select_first_non_null(
                get_safe_property(usage_node, 'Total'),
                get_safe_property(usage_node, 'TotalTokens'),
                get_safe_property(usage_node, 'TokensTotal'),
            ))
            tokens_input = _get_num(select_first_non_null(
                get_safe_property(usage_node, 'Input'),
                get_safe_property(usage_node, 'Prompt'),
                get_safe_property(usage_node, 'InputTokens'),
                get_safe_property(usage_node, 'TokensInput'),
            ))
            tokens_output = _get_num(select_first_non_null(
                get_safe_property(usage_node, 'Output'),
                get_safe_property(usage_node, 'Completion'),
                get_safe_property(usage_node, 'OutputTokens'),
                get_safe_property(usage_node, 'TokensOutput'),
            ))
        if not tokens_total and (tokens_input or tokens_output):
            try:
                tokens_total = (tokens_input or 0) + (tokens_output or 0)
            except Exception:
                pass

        duration_ms = _get_num(select_first_non_null(
            get_safe_property(ced, 'DurationMs') if ced else None,
            get_safe_property(ced, 'ElapsedMs') if ced else None,
            get_safe_property(ced, 'ProcessingTimeMs') if ced else None,
            get_safe_property(ced, 'LatencyMs') if ced else None,
        ))
        outcome_status = select_first_non_null(
            get_safe_property(ced, 'OutcomeStatus') if ced else None,
            get_safe_property(ced, 'Outcome') if ced else None,
            get_safe_property(ced, 'Result') if ced else None,
            get_safe_property(ced, 'Status') if ced else None,
        )
        if isinstance(outcome_status, bool):
            outcome_status = 'Success' if outcome_status else 'Failure'

        conversation_id = select_first_non_null(
            get_safe_property(ced, 'ConversationId') if ced else None,
            get_safe_property(ced, 'ConversationID') if ced else None,
            get_safe_property(ced, 'SessionId') if ced else None,
        )
        turn_number = _get_num(select_first_non_null(
            get_safe_property(ced, 'TurnNumber') if ced else None,
            get_safe_property(ced, 'TurnIndex') if ced else None,
            get_safe_property(ced, 'MessageIndex') if ced else None,
        ))
        retry_count = _get_num(select_first_non_null(
            get_safe_property(ced, 'RetryCount') if ced else None,
            get_safe_property(ced, 'Retries') if ced else None,
        ))
        client_version = select_first_non_null(
            get_safe_property(ced, 'ClientVersion') if ced else None,
            get_safe_property(ced, 'Version') if ced else None,
            get_safe_property(ced, 'Build') if ced else None,
        )
        client_platform = select_first_non_null(
            get_safe_property(ced, 'ClientPlatform') if ced else None,
            get_safe_property(ced, 'Platform') if ced else None,
            get_safe_property(ced, 'OS') if ced else None,
        )
        agent_id_val = select_first_non_null(
            get_safe_property(ced, 'AgentId') if ced else None,
            get_safe_property(ced, 'AgentID') if ced else None,
            get_safe_property(ced, 'AssistantId') if ced else None,
        )
        agent_name = select_first_non_null(
            get_safe_property(ced, 'AgentName') if ced else None,
            get_safe_property(ced, 'AssistantName') if ced else None,
        )
        agent_version = select_first_non_null(
            get_safe_property(ced, 'AgentVersion') if ced else None,
            get_safe_property(ced, 'Version') if ced else None,
        )
        agent_category = _categorize_agent(agent_id_val)

        app_identity = select_first_non_null(
            get_safe_property(ced, 'AppIdentity') if ced else None,
            get_safe_property(ced, 'ApplicationId') if ced else None,
            get_safe_property(ced, 'HostAppId') if ced else None,
        )
        application_name = select_first_non_null(
            get_safe_property(ced, 'ApplicationName') if ced else None,
            get_safe_property(ced, 'HostAppName') if ced else None,
            get_safe_property(ced, 'ClientAppName') if ced else None,
        )

        suggestions = get_safe_property(ced, 'Suggestions') if ced else None
        if not suggestions and ced:
            suggestions = get_safe_property(ced, 'SuggestionList')
        actions = get_safe_property(ced, 'Actions') if ced else None
        references = select_first_non_null(
            get_safe_property(ced, 'References') if ced else None,
            get_safe_property(ced, 'Sources') if ced else None,
            get_safe_property(ced, 'Citations') if ced else None,
        )
        participants = get_safe_property(ced, 'Participants') if ced else None

        suggest_agg = _measure_collection(suggestions, 'Suggestions')
        action_agg = _measure_collection(actions, 'Actions')
        ref_agg = _measure_collection(references, 'References')
        part_agg = _measure_collection(participants, 'Participants')

        # Format creation date and time
        creation_date_raw = record.get('CreationDate')
        if creation_date_raw:
            parsed_dt = parse_date_safe(creation_date_raw)
            creation_date_str = parsed_dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{parsed_dt.microsecond // 1000:03d}Z" if parsed_dt else str(creation_date_raw)
        else:
            creation_date_str = ''

        creation_time_raw = get_safe_property(audit_data, 'CreationTime')
        if creation_time_raw:
            parsed_ct = parse_date_safe(creation_time_raw)
            creation_time_str = parsed_ct.strftime("%Y-%m-%dT%H:%M:%S.") + f"{parsed_ct.microsecond // 1000:03d}Z" if parsed_ct else str(creation_time_raw)
        else:
            creation_time_str = ''

        base_record: dict[str, Any] = {
            'RecordId': record.get('RecordId') or record.get('Identity') or record.get('Id') or get_safe_property(audit_data, 'Id'),
            'RecordType': record.get('RecordType'),
            'CreationDate': creation_date_str,
            'ResultStatus': record.get('ResultStatus'),
            'ResultCount': record.get('ResultCount'),
            'Identity': record.get('Identity'),
            'IsValid': record.get('IsValid'),
            'ObjectState': record.get('ObjectState'),
            'Id': get_safe_property(audit_data, 'Id'),
            'CreationTime': creation_time_str,
            'Operation': get_safe_property(audit_data, 'Operation'),
            'OrganizationId': get_safe_property(audit_data, 'OrganizationId'),
            'RecordTypeNum': get_safe_property(audit_data, 'RecordType'),
            'ResultStatus_Audit': get_safe_property(audit_data, 'ResultStatus'),
            'UserKey': get_safe_property(audit_data, 'UserKey'),
            'UserType': get_safe_property(audit_data, 'UserType'),
            'Version': get_safe_property(audit_data, 'Version'),
            'Workload': get_safe_property(audit_data, 'Workload'),
            'UserId': get_safe_property(audit_data, 'UserId'),
            'AppId': get_safe_property(audit_data, 'AppId'),
            'ClientAppId': get_safe_property(audit_data, 'ClientAppId'),
            'CorrelationId': get_safe_property(audit_data, 'CorrelationId'),
            'ModelId': model_id,
            'ModelProvider': model_provider,
            'ModelFamily': model_family,
            'TokensTotal': tokens_total,
            'TokensInput': tokens_input,
            'TokensOutput': tokens_output,
            'DurationMs': duration_ms,
            'OutcomeStatus': outcome_status,
            'ConversationId': conversation_id,
            'TurnNumber': turn_number,
            'RetryCount': retry_count,
            'ClientVersion': client_version,
            'ClientPlatform': client_platform,
            'AgentId': agent_id_val,
            'AgentName': agent_name,
            'AgentVersion': agent_version,
            'AgentCategory': agent_category,
            'AppIdentity': app_identity,
            'ApplicationName': application_name,
        }

        # Flatten AppAccessContext for Copilot/AI records
        aac = get_safe_property(audit_data, 'AppAccessContext')
        if aac and isinstance(aac, dict):
            flat_aac = convert_to_flat_columns(aac, prefix='AppAccessContext.', max_depth=flat_depth_standard)
            for k, v in flat_aac.items():
                if k not in base_record:
                    base_record[k] = v
            # Remove raw AppAccessContext if it was set
            base_record.pop('AppAccessContext', None)
        elif aac and test_scalar_value(aac):
            if 'AppAccessContext' not in base_record:
                base_record['AppAccessContext'] = aac

        # Add aggregation metrics
        for k, v in suggest_agg.items():
            _add_or_update(base_record, k, v)
        for k, v in action_agg.items():
            _add_or_update(base_record, k, v)
        for k, v in ref_agg.items():
            _add_or_update(base_record, k, v)
        for k, v in part_agg.items():
            _add_or_update(base_record, k, v)

        # If not doing array explosion, return base record now
        if not enable_explosion:
            return [base_record]

        # ARRAY EXPLOSION
        rows: list[dict[str, Any]] = [base_record]
        arrays_to_explode = [
            {'Name': 'Suggestions', 'Data': suggestions, 'Prefix': 'Suggestion', 'Enabled': bool(suggestions)},
            {'Name': 'Actions', 'Data': actions, 'Prefix': 'Action', 'Enabled': bool(actions)},
            {'Name': 'References', 'Data': references, 'Prefix': 'Reference', 'Enabled': bool(references)},
            {'Name': 'Participants', 'Data': participants, 'Prefix': 'Participant', 'Enabled': bool(participants)},
        ]

        max_rows = explosion_per_record_row_cap
        for entry in arrays_to_explode:
            if not entry['Enabled']:
                continue
            data_arr = list(entry['Data']) if hasattr(entry['Data'], '__iter__') and not isinstance(entry['Data'], (str, dict)) else [entry['Data']]
            if len(data_arr) == 0:
                continue

            new_rows: list[dict[str, Any]] = []
            for r in rows:
                for idx, el in enumerate(data_arr):
                    nr = dict(r)  # shallow copy
                    _add_or_update(nr, f"ArrayIndex_{entry['Name']}", idx)
                    if el and isinstance(el, dict):
                        for prop_name, prop_val in el.items():
                            pname = f"{entry['Prefix']}_{prop_name}"
                            if pname in nr:
                                continue
                            if test_scalar_value(prop_val):
                                _add_or_update(nr, pname, prop_val)
                            else:
                                try:
                                    _add_or_update(nr, pname, json.dumps(prop_val, default=str, ensure_ascii=False))
                                except Exception:
                                    pass
                    new_rows.append(nr)
                    if len(new_rows) > max_rows:
                        break
                if len(new_rows) > max_rows:
                    break

            rows = new_rows
            if len(rows) > max_rows:
                break

        # Truncation handling
        if len(rows) > max_rows:
            for r in rows:
                _add_or_update(r, 'ExplosionTruncated', True)
            rows = rows[:max_rows]
            metrics['ExplosionTruncated'] = True

        # Deep flatten CED into all rows
        if explode_deep and ced:
            for r in rows:
                flat = convert_to_flat_columns(ced, prefix='', max_depth=flat_depth_standard)
                for ck, cv in flat.items():
                    if ck not in r:
                        _add_or_update(r, ck, cv)

        return rows

    except Exception:
        metrics['FilteringSkippedRecords'] = metrics.get('FilteringSkippedRecords', 0) + 1
        metrics['FilteringParseFailures'] = metrics.get('FilteringParseFailures', 0) + 1
        return []
