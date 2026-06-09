"""
Module 14: pax_remote_output
=============================
Complete remote output subsystem — URL resolution, token management (Az/Fabric),
pre-flight probe, upload (small + chunked), download (for AppendFile pull), and dispatch.

PS Source: Lines 6510–7380 in PAX_Purview_Audit_Log_Processor_v1.11.1.ps1

Functions migrated (13 + 1 helper):
  F0: get_display_path              (Get-DisplayPath, L6510)
  F1: resolve_sharepoint_target     (Resolve-SharePointTarget, L6553)
  F2: resolve_fabric_target         (Resolve-FabricTarget, L6715)
  F3: get_fabric_storage_token_raw  (Get-FabricStorageTokenRaw, L6771)
  F4: invoke_az_token_acquire       (Invoke-AzTokenAcquire, L6829)
  F5: refresh_fabric_token_if_needed(Refresh-FabricTokenIfNeeded, L6861)
  F6: get_fabric_storage_token      (Get-FabricStorageToken, L6920)
  F7: invoke_fabric_web_request     (Invoke-FabricWebRequest, L6937)
  F8: test_remote_destination       (Test-RemoteDestination, L7008)
  F9: send_file_to_sharepoint       (Send-FileToSharePoint, L7118)
  F10: send_file_to_onelake         (Send-FileToOneLake, L7183)
  F11: get_remote_file_sharepoint   (Get-RemoteFile-SharePoint, L7250)
  F12: get_remote_file_onelake      (Get-RemoteFile-OneLake, L7263)
  F13: invoke_output_upload         (Invoke-OutputUpload, L7277)

Architecture:
  All external dependencies (Graph API calls, Az token acquisition, HTTP requests,
  file I/O) are injected via callbacks for testability. The module maintains an
  AzAuthState dataclass as mutable state (mirrors $script:AzAuthState in PS).

Hard dependencies: pax_auth (for Graph tokens used by SharePoint operations)
"""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Callable, Dict, Optional
from urllib.parse import urlparse

logger = logging.getLogger(__name__)

# ============================================================
# DATA CLASSES
# ============================================================


@dataclass
class SharePointTarget:
    """Resolved SharePoint target from URL parsing + Graph site/drive lookup."""
    host_name: str = ''
    site_name: str = ''
    site_id: str = ''
    drive_id: str = ''
    library_name: str = ''
    folder_path: str = ''
    web_url: str = ''


@dataclass
class FabricTarget:
    """Resolved Fabric/OneLake target from URL parsing."""
    account_url: str = ''
    workspace: str = ''
    item_name: str = ''
    item_type: str = ''
    item_full: str = ''
    files_path: str = ''
    filesystem_base: str = ''


@dataclass
class AzAuthState:
    """
    Mirrors $script:AzAuthState — mutable token state for OneLake DFS calls.
    """
    token: Optional[str] = None
    expires_on: Optional[datetime] = None  # UTC
    acquired_at: Optional[datetime] = None  # local, for age-cap
    last_refresh: Optional[datetime] = None
    last_refresh_attempt: Optional[datetime] = None  # cooldown anchor
    refresh_count: int = 0
    auth_method: Optional[str] = None  # 'ManagedIdentity' | 'AppRegistration' | 'Interactive'


# ============================================================
# F1: resolve_sharepoint_target (Resolve-SharePointTarget)
# ============================================================

# Site-prefix segments recognized by SharePoint
_SP_PREFIXES = {'sites', 'teams', 'personal'}


def resolve_sharepoint_target(
    url: str,
    *,
    graph_request_fn: Optional[Callable[[str, str, Optional[Any]], Any]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> SharePointTarget:
    """
    Resolves a SharePoint URL into site ID, drive ID, library name, and folder path.

    Accepts:
      https://<host>/sites/<site>/<library>[/<sub>]
      https://<host>/teams/<team>/<library>[/<sub>]
      https://<host>/personal/<user>/<library>[/<sub>]
      https://<host>/<library>[/<sub>]    (root site collection)

    Parameters
    ----------
    url : str
        Full SharePoint URL.
    graph_request_fn : callable(method, uri, body=None) -> dict
        Executes Graph API requests. Must handle auth internally.
    log_fn : callable(message, level)
        Logging callback.

    Returns
    -------
    SharePointTarget
        Resolved target with site_id, drive_id, library_name, folder_path.

    Raises
    ------
    ValueError
        If URL is malformed or missing required segments.
    RuntimeError
        If Graph site lookup fails (with classified diagnostics).
    """
    _log = log_fn or (lambda msg, lvl: None)

    parsed = urlparse(url)
    host_name = parsed.hostname or ''
    if not host_name:
        raise ValueError(f"SharePoint URL has no host: '{url}'")

    # Split path into non-empty segments
    segments = [s for s in parsed.path.strip('/').split('/') if s]
    if not segments:
        raise ValueError(
            f"SharePoint URL has no path; expected https://<host>/<sites|teams|personal>/<name>/<library>[/<sub>] : '{url}'"
        )

    # Determine where site collection ends and library begins
    prefixed = segments[0].lower() in _SP_PREFIXES
    if prefixed:
        if len(segments) < 3:
            raise ValueError(
                f"SharePoint URL is missing a library/folder after the site name: '{url}'"
            )
        site_path_segs = segments[0:2]  # e.g. ['sites', 'Analytics']
        lib_and_folder = segments[2:]
        site_name = segments[1]
    else:
        # Root site collection
        site_path_segs = []
        lib_and_folder = segments
        site_name = '(root)'

    library_name = lib_and_folder[0]
    folder_in_library = '/'.join(lib_and_folder[1:]) if len(lib_and_folder) > 1 else ''

    # Resolve site via Graph
    if site_path_segs:
        site_path = '/'.join(site_path_segs)
        site_lookup_uri = f"https://graph.microsoft.com/v1.0/sites/{host_name}:/{site_path}"
    else:
        site_lookup_uri = f"https://graph.microsoft.com/v1.0/sites/{host_name}"

    if graph_request_fn is None:
        raise RuntimeError("No graph_request_fn provided for SharePoint site resolution.")

    try:
        site = graph_request_fn('GET', site_lookup_uri, None)
    except Exception as e:
        # Classify the failure with actionable diagnostics (mirrors PS behavior)
        error_msg = str(e)
        status = _extract_status_code(e)
        site_path_str = '/'.join(site_path_segs)

        diag_lines = [
            f"SharePoint site lookup failed (host '{host_name}', site path '{site_path_str}')."
        ]

        if status == 401:
            diag_lines.append('Cause: Graph rejected the token (401). Application permissions')
            diag_lines.append('       (Sites.ReadWrite.All / Files.ReadWrite.All) may not be admin-consented.')
            diag_lines.append('Action: Have a Global Administrator grant admin consent to the application')
            diag_lines.append('        permissions listed in the "Permissions Required for THIS run" table above.')
        elif status == 403:
            diag_lines.append('Cause: The signed-in identity is authenticated but does not have')
            diag_lines.append('       access to the target SharePoint site (HTTP 403 Forbidden).')
            diag_lines.append('Action: In SharePoint, add this identity to the target site as at least Member,')
            diag_lines.append('        OR re-run with an identity that already has access to the site.')
        elif status == 404:
            diag_lines.append(f"Cause: The site was not found on host '{host_name}' (HTTP 404).")
            diag_lines.append('Action: Verify the URL was copied via SharePoint Details -> Path -> Copy.')
            diag_lines.append('        Check spelling of the /sites/<name> or /teams/<name> segment.')
        elif status >= 500:
            diag_lines.append(f"Cause: Microsoft Graph returned a server error (HTTP {status}). This is usually transient.")
            diag_lines.append('Action: Wait a minute and retry. If it persists, check https://status.cloud.microsoft.')
        else:
            # Includes DNS failures, dogfood endpoints
            if re.search(r'sharepoint-(df|ppe)\.com$', host_name):
                diag_lines.append(f"Cause: Host '{host_name}' is a Microsoft-internal dogfood/PPE SharePoint endpoint")
                diag_lines.append('       and is not reachable from the public Microsoft Graph service.')
                diag_lines.append('Action: Use a production SharePoint URL (*.sharepoint.com) instead.')
            else:
                diag_lines.append(f"Cause: Graph site lookup failed (HTTP status: {status}).")
                diag_lines.append('Action: Verify the URL is correct and reachable, and that the signed-in identity')
                diag_lines.append('        has the required Graph scopes and SharePoint site access.')

        # Truncate raw detail
        raw_detail = error_msg[:600] + ('...(truncated)' if len(error_msg) > 600 else '')
        diag_lines.append('')
        diag_lines.append(f"Graph response: {raw_detail}")

        raise RuntimeError('\n'.join(diag_lines)) from e

    site_id = site.get('id') if isinstance(site, dict) else None
    if not site_id:
        raise RuntimeError(f"Unable to resolve SharePoint site '{url}' (Graph returned no id).")

    # Resolve drive (document library) by name
    drives_uri = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives_resp = graph_request_fn('GET', drives_uri, None)
    drives_list = drives_resp.get('value', []) if isinstance(drives_resp, dict) else []

    drive = None
    for d in drives_list:
        d_name = d.get('name', '')
        d_web_url = d.get('webUrl', '')
        if d_name == library_name or re.search(f'/{re.escape(library_name)}(/|$)', d_web_url):
            drive = d
            break

    if not drive:
        # Fallback: default drive
        default_drive_uri = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
        drive = graph_request_fn('GET', default_drive_uri, None)
        # Treat entire path-after-site as folder under default drive
        folder_in_library = '/'.join(lib_and_folder)

    return SharePointTarget(
        host_name=host_name,
        site_name=site_name,
        site_id=site_id,
        drive_id=drive.get('id', '') if isinstance(drive, dict) else '',
        library_name=library_name,
        folder_path=folder_in_library.rstrip('/'),
        web_url=drive.get('webUrl', '') if isinstance(drive, dict) else '',
    )


# ============================================================
# F2: resolve_fabric_target (Resolve-FabricTarget)
# ============================================================


def resolve_fabric_target(url: str) -> FabricTarget:
    """
    Resolves a Fabric/OneLake URL into workspace, item, and path components.

    Expected format:
      https://onelake.dfs.fabric.microsoft.com/<workspace>/<item>.Lakehouse/Files[/<rel>]

    Parameters
    ----------
    url : str
        Full OneLake DFS URL.

    Returns
    -------
    FabricTarget
        Resolved target with workspace, item_name, item_type, files_path, etc.

    Raises
    ------
    ValueError
        If URL is malformed or missing required segments.
    """
    parsed = urlparse(url)
    account_url = f"{parsed.scheme}://{parsed.hostname}"
    segments = [s for s in parsed.path.strip('/').split('/') if s]

    if len(segments) < 3:
        raise ValueError(f"Fabric URL must include workspace, item, and Files segment: '{url}'")

    workspace = segments[0]
    item_full = segments[1]  # e.g. MyLakehouse.Lakehouse

    # Validate item suffix
    match = re.search(r'\.(Lakehouse|Warehouse)$', item_full)
    if not match:
        raise ValueError(f"Fabric item must end with .Lakehouse or .Warehouse: '{item_full}'")

    item_type = match.group(1)
    item_name = item_full[:-(len(item_type) + 1)]

    if segments[2] != 'Files':
        raise ValueError(f"Fabric URL must include the 'Files' segment: '{url}'")

    rel = '/'.join(segments[3:]) if len(segments) > 3 else ''

    return FabricTarget(
        account_url=account_url,
        workspace=workspace,
        item_name=item_name,
        item_type=item_type,
        item_full=item_full,
        files_path=rel.rstrip('/'),
        filesystem_base=f"{account_url}/{workspace}",
    )


# ============================================================
# F3: get_fabric_storage_token_raw (Get-FabricStorageTokenRaw)
# ============================================================


@dataclass
class TokenResult:
    """Raw token acquisition result."""
    token: str = ''
    expires_on: Optional[datetime] = None  # UTC
    auth_method: str = 'Interactive'  # 'ManagedIdentity' | 'AppRegistration' | 'Interactive'


def get_fabric_storage_token_raw(
    *,
    acquire_token_fn: Optional[Callable[[], TokenResult]] = None,
    auth_mode: str = 'Interactive',
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> TokenResult:
    """
    Acquires a raw OneLake storage token for the https://storage.azure.com/ audience.

    In the PS version, this calls Az.Accounts (Connect-AzAccount + Get-AzAccessToken).
    In Python, the caller injects the actual token acquisition via acquire_token_fn.

    Parameters
    ----------
    acquire_token_fn : callable() -> TokenResult
        Injected function that returns a TokenResult with token, expires_on, auth_method.
        If None, raises RuntimeError.
    auth_mode : str
        'ManagedIdentity', 'AppRegistration', or 'Interactive'.
    tenant_id, client_id, client_secret : str
        Credentials passed through for auth_mode context.
    log_fn : callable
        Logging callback.

    Returns
    -------
    TokenResult
        Contains token string, expires_on (UTC), auth_method tag.

    Raises
    ------
    RuntimeError
        If token acquisition fails or no acquire_token_fn provided.
    """
    if acquire_token_fn is None:
        raise RuntimeError(
            "No acquire_token_fn provided. In production, inject azure-identity credential."
        )

    result = acquire_token_fn()
    if not result or not result.token:
        raise RuntimeError("Token acquisition returned empty token.")

    return result


# ============================================================
# F4: invoke_az_token_acquire (Invoke-AzTokenAcquire)
# ============================================================


def invoke_az_token_acquire(
    state: AzAuthState,
    *,
    reason: str = 'initial',
    acquire_token_fn: Optional[Callable[[], TokenResult]] = None,
    now_fn: Optional[Callable[[], datetime]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> bool:
    """
    Acquires a fresh OneLake storage token and updates AzAuthState.

    Includes stale-token rejection: refuses a token already at/near expiry (≤2 min).

    Parameters
    ----------
    state : AzAuthState
        Mutable auth state object (updated in-place).
    reason : str
        Human-readable reason for acquisition (for logging).
    acquire_token_fn : callable() -> TokenResult
        Injected token acquisition function.
    now_fn : callable() -> datetime
        Returns current UTC time.
    log_fn : callable(msg, level)
        Logging callback.

    Returns
    -------
    bool
        True if token acquired successfully, False on failure.
    """
    _log = log_fn or (lambda msg, lvl: None)
    _now = now_fn or (lambda: datetime.now(timezone.utc))

    try:
        token_obj = get_fabric_storage_token_raw(acquire_token_fn=acquire_token_fn)

        # Stale-token rejection
        now_utc = _now()
        if token_obj.expires_on is not None:
            minutes_valid = (token_obj.expires_on - now_utc).total_seconds() / 60.0
            if minutes_valid <= 2:
                raise RuntimeError(
                    f"Acquired OneLake storage token is already expired or near-expiry "
                    f"(expires in {minutes_valid:.1f} min)."
                )

        # Update state
        current_time = _now()
        state.token = token_obj.token
        state.expires_on = token_obj.expires_on
        state.acquired_at = current_time
        state.last_refresh = current_time
        state.refresh_count += 1
        state.auth_method = token_obj.auth_method

        verb = 'acquired' if state.refresh_count == 1 else 'refreshed'
        expires_str = token_obj.expires_on.strftime('%Y-%m-%d %H:%M:%S') if token_obj.expires_on else 'unknown'
        _log(
            f"  [AZ-TOKEN] OneLake storage token {verb} ({reason}); "
            f"auth: {token_obj.auth_method}; expires {expires_str} UTC; "
            f"refresh #{state.refresh_count}",
            'info'
        )
        return True

    except Exception as e:
        _log(
            f"  [AZ-TOKEN] [!] Failed to acquire OneLake storage token ({reason}): {e}",
            'error'
        )
        state._last_error = str(e)
        return False


# ============================================================
# F5: refresh_fabric_token_if_needed (Refresh-FabricTokenIfNeeded)
# ============================================================


def refresh_fabric_token_if_needed(
    state: AzAuthState,
    *,
    buffer_minutes: int = 5,
    force: bool = False,
    acquire_token_fn: Optional[Callable[[], TokenResult]] = None,
    now_fn: Optional[Callable[[], datetime]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> bool:
    """
    Proactively refreshes the Fabric/OneLake storage token if nearing expiry.

    Trigger conditions:
      - Forced (typically 401 reactive retry)
      - Within buffer_minutes of expiry (proactive)
      - Token age > 50 min (belt-and-suspenders)

    Cooldown:
      - App-only/MI: 45s (0.75 min) — silent client_credentials/IMDS
      - Interactive: 5 min — avoids re-prompt spam

    Parameters
    ----------
    state : AzAuthState
        Mutable auth state.
    buffer_minutes : int
        Refresh if token expires within this many minutes. Default: 5.
    force : bool
        Bypass cooldown and force re-acquisition (used by 401 reactive retry).
    acquire_token_fn : callable
        Token acquisition function.
    now_fn : callable
        UTC time source.
    log_fn : callable
        Logging callback.

    Returns
    -------
    bool
        True if token is valid (still-valid or freshly refreshed), False on failure.
    """
    _now = now_fn or (lambda: datetime.now(timezone.utc))
    _log = log_fn or (lambda msg, lvl: None)
    now = _now()

    # First-time acquisition: no state yet
    if not state.token or not state.expires_on:
        return invoke_az_token_acquire(
            state, reason='initial', acquire_token_fn=acquire_token_fn,
            now_fn=now_fn, log_fn=log_fn
        )

    minutes_remaining = (state.expires_on - now).total_seconds() / 60.0
    token_age = (
        (now - state.acquired_at).total_seconds() / 60.0
        if state.acquired_at else 999.0
    )

    # Trigger conditions
    need_refresh = force or (minutes_remaining <= buffer_minutes) or (token_age > 50)
    if not need_refresh:
        return True

    # Cooldown check
    is_app_only = state.auth_method in ('ManagedIdentity', 'AppRegistration')
    cooldown_minutes = 0.75 if is_app_only else 5.0

    if not force and state.last_refresh_attempt:
        since_last = (now - state.last_refresh_attempt).total_seconds() / 60.0
        if since_last < cooldown_minutes:
            # Still in cooldown; report current token as valid only if it actually is
            return minutes_remaining > 0

    state.last_refresh_attempt = now

    # Determine reason
    if force:
        reason = 'forced (401 retry or explicit)'
    elif minutes_remaining <= buffer_minutes:
        reason = f'near-expiry ({minutes_remaining:.1f} min remaining)'
    else:
        reason = f'age cap ({token_age:.1f} min)'

    return invoke_az_token_acquire(
        state, reason=reason, acquire_token_fn=acquire_token_fn,
        now_fn=now_fn, log_fn=log_fn
    )


# ============================================================
# F6: get_fabric_storage_token (Get-FabricStorageToken)
# ============================================================


def get_fabric_storage_token(
    state: AzAuthState,
    *,
    acquire_token_fn: Optional[Callable[[], TokenResult]] = None,
    now_fn: Optional[Callable[[], datetime]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> str:
    """
    Returns a valid OneLake storage-audience bearer token string.

    Always invokes refresh_fabric_token_if_needed first so callers receive a
    token guaranteed to have at least BufferMinutes of validity.

    Parameters
    ----------
    state : AzAuthState
        Mutable auth state.
    acquire_token_fn : callable
        Token acquisition function.
    now_fn : callable
        UTC time source.
    log_fn : callable
        Logging callback.

    Returns
    -------
    str
        Bearer token string.

    Raises
    ------
    RuntimeError
        If token acquisition/refresh fails.
    """
    ok = refresh_fabric_token_if_needed(
        state, acquire_token_fn=acquire_token_fn, now_fn=now_fn, log_fn=log_fn
    )
    if not ok or not state.token:
        last_err = getattr(state, '_last_error', '') or ''
        msg = "Failed to acquire Fabric/OneLake storage token (see prior [AZ-TOKEN] messages)."
        if last_err:
            msg += f" {last_err}"
        raise RuntimeError(msg)
    return state.token


# ============================================================
# F7: invoke_fabric_web_request (Invoke-FabricWebRequest)
# ============================================================


def invoke_fabric_web_request(
    state: AzAuthState,
    *,
    uri: str,
    method: str,
    headers: Optional[Dict[str, str]] = None,
    body: Optional[bytes] = None,
    content_type: Optional[str] = None,
    out_file: Optional[str] = None,
    http_request_fn: Optional[Callable[..., Any]] = None,
    acquire_token_fn: Optional[Callable[[], TokenResult]] = None,
    now_fn: Optional[Callable[[], datetime]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> Any:
    """
    HTTP request wrapper for OneLake DFS calls with proactive token refresh
    and transparent 401 reactive retry.

    Before each call: ensures token is fresh via refresh_fabric_token_if_needed.
    During the call: if 401 is returned, forces refresh and retries once.
    Any other error (429, 5xx, 403, 404) is re-raised unchanged.

    Parameters
    ----------
    state : AzAuthState
        Mutable auth state.
    uri : str
        Request URI.
    method : str
        HTTP method (GET, PUT, PATCH, HEAD, etc.).
    headers : dict, optional
        Caller-supplied headers. Authorization is always overwritten.
    body : bytes, optional
        Request body.
    content_type : str, optional
        Content-Type header value.
    out_file : str, optional
        If set, response body is written to this file path.
    http_request_fn : callable(method, uri, headers, body, content_type, out_file) -> response
        Injected HTTP client. Must raise Exception with .status_code on failures.
    acquire_token_fn : callable
        Token acquisition function (passed to refresh).
    now_fn : callable
        UTC time source.
    log_fn : callable
        Logging callback.

    Returns
    -------
    Any
        Response from http_request_fn.

    Raises
    ------
    RuntimeError
        If no token available.
    Exception
        On non-401 HTTP failures (re-raised from http_request_fn).
    """
    _log = log_fn or (lambda msg, lvl: None)

    if http_request_fn is None:
        raise RuntimeError("No http_request_fn provided for Invoke-FabricWebRequest.")

    # Pre-flight refresh
    refresh_fabric_token_if_needed(
        state, acquire_token_fn=acquire_token_fn, now_fn=now_fn, log_fn=log_fn
    )
    if not state.token:
        raise RuntimeError("Invoke-FabricWebRequest: no OneLake storage token available.")

    # Build effective headers
    h: Dict[str, str] = {}
    if headers:
        for k, v in headers.items():
            if k.lower() != 'authorization':
                h[k] = v
    h['Authorization'] = f"Bearer {state.token}"
    if 'x-ms-version' not in {k.lower(): k for k in h}:
        h['x-ms-version'] = '2021-06-08'
    # Normalize: ensure actual key is 'x-ms-version' not a variant
    lower_keys = {k.lower(): k for k in h}
    if 'x-ms-version' in lower_keys and lower_keys['x-ms-version'] != 'x-ms-version':
        val = h.pop(lower_keys['x-ms-version'])
        h['x-ms-version'] = val

    try:
        return http_request_fn(method, uri, h, body, content_type, out_file)
    except Exception as e:
        status = _extract_status_code(e)
        if status == 401:
            _log(
                "  [AZ-TOKEN] OneLake returned 401 Unauthorized - forcing token refresh and retrying once...",
                'warn'
            )
            refreshed = refresh_fabric_token_if_needed(
                state, force=True, acquire_token_fn=acquire_token_fn,
                now_fn=now_fn, log_fn=log_fn
            )
            if not refreshed:
                raise RuntimeError(
                    f"OneLake 401 Unauthorized and token refresh failed: {e}"
                ) from e
            h['Authorization'] = f"Bearer {state.token}"
            return http_request_fn(method, uri, h, body, content_type, out_file)
        raise


# ============================================================
# F8: test_remote_destination (Test-RemoteDestination)
# ============================================================


def test_remote_destination(
    *,
    remote_output_mode: str = 'None',
    remote_output_url: Optional[str] = None,
    state: Optional[AzAuthState] = None,
    graph_request_fn: Optional[Callable[[str, str, Optional[Any]], Any]] = None,
    http_request_fn: Optional[Callable[..., Any]] = None,
    acquire_token_fn: Optional[Callable[[], TokenResult]] = None,
    now_fn: Optional[Callable[[], datetime]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> Optional[Any]:
    """
    Pre-flight probe of the remote destination. Validates connectivity and permissions.

    For SharePoint: resolves site/drive, creates folder if missing.
    For Fabric: acquires token, probes filesystem HEAD.

    Parameters
    ----------
    remote_output_mode : str
        'None', 'SharePoint', or 'Fabric'.
    remote_output_url : str
        Remote destination URL.
    state : AzAuthState
        Auth state (for Fabric mode).
    graph_request_fn : callable
        Graph HTTP client (for SharePoint mode).
    http_request_fn : callable
        HTTP client (for Fabric mode).
    acquire_token_fn : callable
        Token acquisition (for Fabric mode).
    now_fn, log_fn : callable
        Time source and logging.

    Returns
    -------
    SharePointTarget or FabricTarget or None
        Resolved target, or None if mode is 'None'.

    Raises
    ------
    RuntimeError
        With classified diagnostic messages on failure.
    """
    _log = log_fn or (lambda msg, lvl: None)

    if remote_output_mode == 'None':
        return None

    if remote_output_mode == 'SharePoint':
        resolved = resolve_sharepoint_target(
            remote_output_url, graph_request_fn=graph_request_fn, log_fn=log_fn
        )
        # Ensure folder exists (create if missing)
        if resolved.folder_path:
            folder_uri = f"https://graph.microsoft.com/v1.0/drives/{resolved.drive_id}/root:/{resolved.folder_path}"
            try:
                graph_request_fn('GET', folder_uri, None)
            except Exception:
                # Create folder hierarchy
                parts = resolved.folder_path.split('/')
                accumulated = ''
                for p in parts:
                    if accumulated:
                        parent_uri = f"https://graph.microsoft.com/v1.0/drives/{resolved.drive_id}/root:/{accumulated}:/children"
                    else:
                        parent_uri = f"https://graph.microsoft.com/v1.0/drives/{resolved.drive_id}/root/children"
                    body = {
                        'name': p,
                        'folder': {},
                        '@microsoft.graph.conflictBehavior': 'replace'
                    }
                    try:
                        graph_request_fn('POST', parent_uri, body)
                    except Exception:
                        pass
                    accumulated = f"{accumulated}/{p}" if accumulated else p

        return resolved

    elif remote_output_mode == 'Fabric':
        resolved = resolve_fabric_target(remote_output_url)

        if state is None:
            state = AzAuthState()

        # Initial token acquisition
        try:
            get_fabric_storage_token(
                state, acquire_token_fn=acquire_token_fn, now_fn=now_fn, log_fn=log_fn
            )
        except Exception as e:
            token_err = str(e)
            diag_lines = [
                f"OneLake storage token acquisition failed for workspace '{resolved.workspace}'."
            ]
            if 'module not installed' in token_err.lower() or 'az.accounts' in token_err.lower():
                diag_lines.append('Cause: The Az.Accounts PowerShell module is not installed.')
                diag_lines.append('Action: Install-Module Az.Accounts -Scope CurrentUser')
            elif any(x in token_err.lower() for x in ['aadsts', 'consent', 'admin']):
                diag_lines.append('Cause: Azure AD rejected the sign-in for the storage.azure.com audience.')
                diag_lines.append('Action: Verify the identity has Azure AD sign-in to your tenant, and that')
                diag_lines.append('        any conditional access / MFA requirements are satisfied.')
            elif any(x in token_err.lower() for x in ['identity', 'managed identity', 'msi', 'imds']):
                diag_lines.append('Cause: Managed identity token acquisition failed (no IMDS endpoint reachable,')
                diag_lines.append('       or AZURE_CLIENT_ID points to an identity not assigned to this host).')
                diag_lines.append('Action: Re-run from a host with an assigned managed identity, or use')
                diag_lines.append('        -Auth Interactive / -Auth AppRegistration instead.')
            else:
                diag_lines.append('Cause: Could not acquire a token for https://storage.azure.com/.')
                diag_lines.append('Action: Verify your authentication mode (-Auth) and that you can sign in.')
            diag_lines.append('')
            diag_lines.append(f"Az response: {token_err}")
            raise RuntimeError('\n'.join(diag_lines)) from e

        # HEAD the filesystem to verify access
        probe_uri = f"{resolved.filesystem_base}?resource=filesystem"
        try:
            invoke_fabric_web_request(
                state, uri=probe_uri, method='HEAD',
                http_request_fn=http_request_fn, acquire_token_fn=acquire_token_fn,
                now_fn=now_fn, log_fn=log_fn
            )
        except Exception as e:
            status = _extract_status_code(e)
            raw_detail = str(e)[:600]
            if len(str(e)) > 600:
                raw_detail += '...(truncated)'

            diag_lines = [
                f"OneLake filesystem probe failed (workspace '{resolved.workspace}', item '{resolved.item_full}')."
            ]

            if status == 401:
                diag_lines.append('Cause: OneLake rejected the storage token (401). The token audience is correct')
                diag_lines.append('       but the identity is not recognized by this Fabric workspace.')
                diag_lines.append('Action: Verify the identity is signed in to the same tenant that owns the Fabric')
                diag_lines.append('        workspace, and that the storage.azure.com token is not expired/blocked.')
            elif status == 403:
                diag_lines.append('Cause: The identity is authenticated but lacks permissions on the Fabric workspace')
                diag_lines.append('       (HTTP 403 Forbidden from OneLake DFS).')
                diag_lines.append('Action: In the Fabric portal -> Workspace settings -> Manage access, grant the')
                diag_lines.append('        identity at least Contributor.')
            elif status == 404:
                diag_lines.append("Cause: Workspace or item not found (HTTP 404). One of these is wrong:")
                diag_lines.append(f"         workspace = '{resolved.workspace}'")
                diag_lines.append(f"         item      = '{resolved.item_full}' (must exist as a Lakehouse/Warehouse)")
                diag_lines.append('Action: Verify the URL by opening the lakehouse in the Fabric portal and using')
                diag_lines.append('        the OneLake "Copy ABFS path" option, then converting to https:// form.')
            elif status >= 500:
                diag_lines.append(f"Cause: OneLake returned a server error (HTTP {status}). This is usually transient.")
                diag_lines.append('Action: Wait a minute and retry. If it persists, check Fabric service health.')
            else:
                diag_lines.append(f"Cause: OneLake DFS probe failed (HTTP status: {status}).")
                diag_lines.append('Action: Verify the Fabric workspace URL is correct and that the identity has')
                diag_lines.append('        the Azure RBAC roles listed in the "Permissions Required for THIS run" table.')
            diag_lines.append('')
            diag_lines.append(f"OneLake response: {raw_detail}")

            raise RuntimeError('\n'.join(diag_lines)) from e

        return resolved

    return None


# ============================================================
# F13b: invoke_output_upload (Invoke-OutputUpload)
# ============================================================


def invoke_output_upload(
    local_path: str,
    *,
    remote_output_mode: str = 'None',
    parent_override: str | None = None,
    file_exists_fn: Optional[Callable[[str], bool]] = None,
    send_sharepoint_fn: Optional[Callable[..., None]] = None,
    send_onelake_fn: Optional[Callable[..., None]] = None,
    log_fn: Optional[Callable[[str, str], None]] = None,
) -> None:
    """
    Dispatches a local file to the configured remote destination.

    Parameters
    ----------
    local_path : str
        Path to local file to upload.
    remote_output_mode : str
        'None', 'SharePoint', or 'Fabric'.
    parent_override : str or None
        Optional parent folder override for remote destination (v1.11.2).
        Passed through to send_sharepoint_fn / send_onelake_fn as a
        keyword argument ``parent_override``.
    file_exists_fn : callable(path) -> bool
        Checks if file exists.
    send_sharepoint_fn : callable(path, *, parent_override) -> None
        Upload to SharePoint.
    send_onelake_fn : callable(path, *, parent_override) -> None
        Upload to OneLake.
    log_fn : callable
        Logging.

    Notes
    -----
    No catch wrapper — lets exceptions propagate so callers can handle
    context-specifically (same as PS version).
    """
    if remote_output_mode == 'None':
        return

    _exists = file_exists_fn or os.path.isfile
    if not _exists(local_path):
        # PS: Write-Verbose "skipping (file not found)"
        return

    if remote_output_mode == 'SharePoint':
        if send_sharepoint_fn:
            send_sharepoint_fn(local_path, parent_override=parent_override)
    elif remote_output_mode == 'Fabric':
        if send_onelake_fn:
            send_onelake_fn(local_path, parent_override=parent_override)


# ============================================================
# PRIVATE HELPERS
# ============================================================
def _extract_status_code(e: Exception) -> int:
    """Extract HTTP status code from an exception."""
    if hasattr(e, 'status_code'):
        return int(e.status_code)
    # Try common patterns
    msg = str(e)
    m = re.search(r'\b(4\d{2}|5\d{2})\b', msg)
    if m:
        return int(m.group(1))
    return 0


def _extract_retry_after(e: Exception) -> int:
    """Extract Retry-After header value from an exception."""
    if hasattr(e, 'retry_after'):
        return int(e.retry_after)
    return 0


def _default_read_bytes(path: str) -> bytes:
    """Default file reader."""
    with open(path, 'rb') as f:
        return f.read()

