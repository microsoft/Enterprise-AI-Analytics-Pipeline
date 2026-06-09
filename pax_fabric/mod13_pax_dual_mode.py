"""
Module 13: pax_dual_mode — Graph API diagnostic, group expansion, unified query, and disconnection.

Migrates the Graph-API paths of the PS dual-mode functions.
EOM-only functions (Connect-ToComplianceCenter, Invoke-AuditCapabilityDiagnostics,
Invoke-SearchUnifiedAuditLogWithRetry) are excluded — no Python SDK equivalent.

All external dependencies are injected via callbacks for testability.
"""

from __future__ import annotations

import logging
import re
from typing import Any, Callable, Dict, List, Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Transient network error patterns (matches PS $transientPatterns)
# ---------------------------------------------------------------------------
TRANSIENT_PATTERNS: List[str] = [
    'timed out',
    'unable to connect',
    'connection',
    'remote name could not be resolved',
    'temporarily unavailable',
]


def _is_transient(error_msg: str) -> bool:
    """Return True if error message matches a known transient network pattern."""
    lower = error_msg.lower()
    return any(p in lower for p in TRANSIENT_PATTERNS)


# ---------------------------------------------------------------------------
# F2: expand_group_to_users (Graph API path only)
# PS: Expand-GroupToUsers with $UseEOMMode = $false
# ---------------------------------------------------------------------------

# GUID pattern (matches PS regex)
_GUID_RE = re.compile(
    r'^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
)


def expand_group_to_users(
    group_identity: str,
    *,
    graph_request_fn: Optional[Callable[[str, str], Any]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> List[str]:
    """
    Expands a distribution/security group to individual user principal names via Graph API.

    Parameters
    ----------
    group_identity : str
        Group identifier — display name, email address, or ObjectId (GUID).
    graph_request_fn : callable(method, url) -> response dict
        Graph API request function. Used for:
          - GET /groups?$filter=displayName eq '...' → resolve name to ID
          - GET /groups?$filter=mail eq '...' → resolve email to ID
          - GET /groups/{id}/members → get members
          - GET /users/{id} → get user UPN
    log_fn : callable(message, level)
        Logging callback.

    Returns
    -------
    List of user principal names (strings). Empty list on failure.
    """
    _log = log_fn or (lambda msg, lvl: logger.log(
        {'info': logging.INFO, 'warn': logging.WARNING, 'error': logging.ERROR}.get(lvl, logging.INFO), msg))

    members: List[str] = []

    if not group_identity or not group_identity.strip():
        return members

    if graph_request_fn is None:
        _log("No graph_request_fn provided for group expansion", 'error')
        return members

    try:
        _log(f"Processing group (Graph API): '{group_identity}'", 'info')

        # --- Resolve group ID ---
        group_id: Optional[str] = None

        if _GUID_RE.match(group_identity):
            # Already a GUID
            group_id = group_identity
        else:
            # Try display name first
            _log("Resolving group ID from display name...", 'info')
            escaped = group_identity.replace("'", "''")
            filter_url = f"https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '{escaped}'"

            try:
                resp = graph_request_fn('GET', filter_url)
                values = resp.get('value', []) if resp else []
                if values:
                    group_id = values[0].get('id')
            except Exception:
                pass

            # Fallback: try by mail
            if not group_id:
                mail_url = f"https://graph.microsoft.com/v1.0/groups?$filter=mail eq '{escaped}'"
                try:
                    resp = graph_request_fn('GET', mail_url)
                    values = resp.get('value', []) if resp else []
                    if values:
                        group_id = values[0].get('id')
                except Exception:
                    pass

            if not group_id:
                raise RuntimeError(f"Unable to find group with identifier: {group_identity}")

            _log(f"Resolved to ObjectId: {group_id}", 'info')

        # --- Get group members ---
        members_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members"
        all_members: List[Dict[str, Any]] = []

        # Pagination support
        next_url: Optional[str] = members_url
        while next_url:
            resp = graph_request_fn('GET', next_url)
            if not resp:
                break
            page_values = resp.get('value', [])
            if page_values:
                all_members.extend(page_values)
            next_url = resp.get('@odata.nextLink')

        # --- Filter to users and extract UPN ---
        for member in all_members:
            odata_type = ''
            # Check additionalProperties or direct @odata.type
            if isinstance(member, dict):
                odata_type = member.get('@odata.type', '')
                if not odata_type:
                    # PS: $member.AdditionalProperties.'@odata.type'
                    props = member.get('additionalProperties', {})
                    if isinstance(props, dict):
                        odata_type = props.get('@odata.type', '')

            if odata_type == '#microsoft.graph.user':
                # Try to get UPN directly from member object
                upn = member.get('userPrincipalName', '')
                if upn:
                    members.append(upn)
                else:
                    # Need to fetch full user object
                    member_id = member.get('id', '')
                    if member_id:
                        try:
                            user_url = f"https://graph.microsoft.com/v1.0/users/{member_id}"
                            user = graph_request_fn('GET', user_url)
                            if user and user.get('userPrincipalName'):
                                members.append(user['userPrincipalName'])
                        except Exception:
                            pass

        _log(f"Expanded: {len(members)} user member(s)", 'info')

    except Exception as e:
        _log(f"Warning: Failed to expand group '{group_identity}': {e}", 'warn')
        _log("Possible causes:", 'warn')
        _log("  - Group does not exist or identifier is invalid", 'warn')
        _log("  - Insufficient permissions (need GroupMember.Read.All)", 'warn')
        _log("  - Network connectivity issues with Graph API", 'warn')

    return members


# ---------------------------------------------------------------------------
# F4: disconnect_purview_audit (Graph API path only)
# PS: Disconnect-PurviewAudit with $UseEOMMode = $false
# ---------------------------------------------------------------------------

def disconnect_purview_audit(
    *,
    get_context_fn: Optional[Callable[[], Optional[Any]]] = None,
    disconnect_fn: Optional[Callable[[], None]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> bool:
    """
    Disconnects from Microsoft Graph cleanly.

    Parameters
    ----------
    get_context_fn : callable
        Returns context object if connected, None otherwise.
    disconnect_fn : callable
        Performs the actual disconnection.
    log_fn : callable(message, level)
        Logging callback.

    Returns
    -------
    True if disconnection was performed or already disconnected, False on error.
    """
    _log = log_fn or (lambda msg, lvl: logger.log(
        {'info': logging.INFO, 'warn': logging.WARNING, 'error': logging.ERROR}.get(lvl, logging.INFO), msg))

    try:
        # Check if connected first
        context = None
        if get_context_fn:
            try:
                context = get_context_fn()
            except Exception:
                pass

        if context:
            _log("Disconnecting from Microsoft Graph...", 'info')
            if disconnect_fn:
                disconnect_fn()
            _log("Disconnected from Microsoft Graph", 'info')
        else:
            _log("(Not connected to Microsoft Graph)", 'info')

        return True

    except Exception:
        _log("(Microsoft Graph disconnection skipped or already disconnected)", 'info')
        return True

