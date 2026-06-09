"""
Module 5: pax_auth — Authentication & Token Management
=======================================================
Migrated from: PAX_Purview_Audit_Log_Processor_v1.11.1.ps1 Lines 7854-8990
Level: 0 (no hard dependencies)

Provides the auth lifecycle for Microsoft Graph Security API:
- API version auto-detection (v1.0 vs beta)
- App-only client-credential authentication (AppRegistration + client_secret)
- Token extraction and JWT decoding
- Proactive token refresh with cooldown
- Thread-safe shared auth state
- SharePoint scope injection for remote output mode
- Deferred auth context display for dual-phase (AppReg + Agent365) runs

External dependencies: msal (Microsoft Authentication Library)
Design: Config values (TenantId, ClientId, secret) passed as parameters.
        Uses stdlib logging.getLogger(__name__).
"""

from __future__ import annotations

import base64
import json
import logging
import threading
from datetime import datetime, timezone, timedelta
from typing import Any, Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Module-level state (mirrors PS $script: auth variables)
# ---------------------------------------------------------------------------

# Detected API version: "v1.0" or "beta" (cached after first detection)
_graph_audit_api_version: Optional[str] = None

# Auth configuration (mirrors PS $script:AuthConfig)
_auth_config: dict[str, Any] = {
    "method": None,           # Always 'AppRegistration' when populated
    "tenant_id": None,
    "client_id": None,
    "client_secret": None,    # SecureString equivalent — stored in memory only
    "can_reauthenticate": False,
    "token_issue_time": None,  # datetime when token was acquired
}

# Shared auth state for thread-safe token access (mirrors PS $script:SharedAuthState)
_shared_auth_state: dict[str, Any] = {
    "token": None,
    "expires_on": None,       # datetime (UTC)
    "last_refresh": None,     # datetime
    "refresh_count": 0,
    "auth_method": None,
}

# Connection state
_connected: bool = False

# MSAL application instance (cached for token refresh)
_msal_app: Any = None

# Lock for thread-safe token operations
_token_lock = threading.Lock()

# Cooldown tracking
_last_proactive_refresh_attempt: Optional[datetime] = None

# Auth failure flags (mirrors PS $script:AuthFailureDetected, Auth401MessageShown)
_auth_failure_detected: bool = False
_auth_401_message_shown: bool = False

# Separate token acquired time (mirrors PS $script:TokenAcquiredTime, distinct from AuthConfig.TokenIssueTime)
_token_acquired_time: Optional[datetime] = None

# Deferred auth context display (NEW v1.11.1)
# When AppRegistration + Agent365 run, context display is deferred until Phase 2 completes
_defer_auth_context_display: bool = False

# Phase 1 context stored for dual-context display (NEW v1.11.1)
_phase1_context: Optional[dict[str, Any]] = None


# ---------------------------------------------------------------------------
# Version detection config (mirrors PS $script:GraphAuditApiVersion_Current/Previous)
# ---------------------------------------------------------------------------
_GRAPH_API_VERSION_CURRENT = "v1.0"
_GRAPH_API_VERSION_PREVIOUS = "beta"
_GRAPH_BASE_URL = "https://graph.microsoft.com"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_graph_audit_api_uri(path: str, *, http_client=None) -> str:
    """
    Build Graph API audit endpoint URI with automatic version detection.
    Version is detected on first call and cached for the session.

    PS equivalent: Get-GraphAuditApiUri (L7854)

    Args:
        path: The audit API path (e.g., "queries", "queries/{id}/records")
        http_client: Optional HTTP client with .get() method for version probing.
                     If None, uses the current version without probing.

    Returns:
        Full Graph API URI string.
    """
    global _graph_audit_api_version

    if _graph_audit_api_version is None:
        if http_client is not None:
            # Auto-detect: try current version first
            test_uri = f"{_GRAPH_BASE_URL}/{_GRAPH_API_VERSION_CURRENT}/security/auditLog/queries"
            try:
                resp = http_client.get(test_uri)
                if resp.status_code < 400 or resp.status_code == 403:
                    # 403 means endpoint exists but permission denied — still valid version
                    _graph_audit_api_version = _GRAPH_API_VERSION_CURRENT
                    logger.info(
                        "Graph API: security/auditLog endpoint using version %s",
                        _GRAPH_API_VERSION_CURRENT,
                    )
                else:
                    raise Exception(f"HTTP {resp.status_code}")
            except Exception:
                _graph_audit_api_version = _GRAPH_API_VERSION_PREVIOUS
                logger.warning(
                    "Graph API: security/auditLog endpoint using version %s (fallback from %s)",
                    _GRAPH_API_VERSION_PREVIOUS,
                    _GRAPH_API_VERSION_CURRENT,
                )
        else:
            # No http_client — default to current version
            _graph_audit_api_version = _GRAPH_API_VERSION_CURRENT

    return f"{_GRAPH_BASE_URL}/{_graph_audit_api_version}/security/auditLog/{path}"


def connect_purview_audit(
    auth_method: str = "AppRegistration",
    *,
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
    scopes: Optional[list[str]] = None,
    http_client=None,
    remote_output_mode: str = "None",
    include_agent365: bool = False,
) -> dict[str, Any]:
    """
    Authenticate to Microsoft Graph Security API using MSAL.

    PS equivalent: Connect-PurviewAudit (L8187) — Graph API mode only.

    Only app-only client-credential flow (AppRegistration + client_secret) is
    supported. Other auth methods were removed in v1.11.4 as unused.

    Args:
        auth_method: Must be 'AppRegistration' (only supported method).
        tenant_id: Azure AD tenant ID.
        client_id: App registration client ID.
        client_secret: Client secret (required).
        scopes: List of Graph API scopes to request.
        http_client: Optional HTTP client for API version detection.
        remote_output_mode: 'None' | 'SharePoint' | 'Fabric'. When 'SharePoint',
                            injects Sites.ReadWrite.All + Files.ReadWrite.All scopes.
        include_agent365: If True, defers auth context display until Phase 2
                          Agent365 sign-in completes.

    Returns:
        Dict with 'token', 'expires_on', 'tenant_id', 'client_id', 'auth_method',
        'account_display', 'deferred_context'.

    Raises:
        ValueError: For invalid parameters.
        RuntimeError: For authentication failures.
    """
    global _connected, _msal_app, _auth_config, _shared_auth_state
    global _defer_auth_context_display, _phase1_context

    if auth_method != "AppRegistration":
        raise ValueError(
            f"Invalid auth_method '{auth_method}'. "
            "Only 'AppRegistration' (with client_secret) is supported."
        )

    if not tenant_id:
        raise ValueError("tenant_id is required for Graph API authentication")
    if not client_id:
        raise ValueError("client_id is required for Graph API authentication")
    if not client_secret:
        raise ValueError("client_secret is required for AppRegistration authentication")

    # Default scopes
    if scopes is None:
        scopes = ["https://graph.microsoft.com/.default"]

    # SharePoint remote-output: inject drive write scopes (PS L8398)
    if remote_output_mode == "SharePoint":
        if "Sites.ReadWrite.All" not in scopes:
            scopes = list(scopes) + ["Sites.ReadWrite.All"]
        if "Files.ReadWrite.All" not in scopes:
            scopes = list(scopes) + ["Files.ReadWrite.All"]

    logger.info("Connecting to Microsoft Graph Security API (method: AppRegistration)...")

    try:
        result = _connect_app_registration(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            scopes=scopes,
        )

        # Update module state
        _auth_config.update({
            "method": "AppRegistration",
            "tenant_id": tenant_id,
            "client_id": client_id,
            "client_secret": client_secret,
            "can_reauthenticate": True,
            "token_issue_time": datetime.now(timezone.utc),
        })

        token = result.get("access_token", "")
        expires_in = result.get("expires_in", 3600)
        expires_on = datetime.now(timezone.utc) + timedelta(seconds=expires_in)

        # Try to decode actual expiry from JWT
        jwt_expiry = _decode_jwt_expiry(token)
        if jwt_expiry:
            expires_on = jwt_expiry

        with _token_lock:
            _shared_auth_state.update({
                "token": token,
                "expires_on": expires_on,
                "last_refresh": datetime.now(timezone.utc),
                "auth_method": "appregistration",
            })

        _connected = True
        logger.info("Successfully connected to Microsoft Graph (method: AppRegistration)")

        # Trigger API version detection
        get_graph_audit_api_uri("queries", http_client=http_client)

        # Build account display label (PS L8616-8625)
        account_display = "(app-only / AppRegistration - no interactive user)"

        # Deferred auth context display for dual-phase runs (PS L8607-8614)
        deferred_context = False
        if include_agent365:
            _defer_auth_context_display = True
            _phase1_context = {
                "tenant_id": tenant_id,
                "account": account_display,
                "granted_required": list(scopes),
                "required_scopes": list(scopes),
            }
            deferred_context = True
            logger.info("  Phase 1 (audit) connected.")
        else:
            _defer_auth_context_display = False

        return {
            "token": token,
            "expires_on": expires_on,
            "tenant_id": tenant_id,
            "client_id": client_id,
            "auth_method": "AppRegistration",
            "account_display": account_display,
            "deferred_context": deferred_context,
        }

    except Exception as e:
        logger.error("Graph API authentication failed: %s", str(e))
        raise RuntimeError(f"Authentication failed: {e}") from e


def get_graph_access_token() -> Optional[str]:
    """
    Extract the current access token from the shared auth state.

    PS equivalent: Get-GraphAccessToken (L8696)

    Returns:
        The access token string, or None if not available.
    """
    with _token_lock:
        return _shared_auth_state.get("token")


def get_graph_access_token_with_expiry() -> Optional[dict[str, Any]]:
    """
    Extract access token AND expiry time. Decodes JWT 'exp' claim if possible.

    PS equivalent: Get-GraphAccessTokenWithExpiry (L8744)

    Returns:
        Dict with 'token', 'expires_on' (datetime UTC), 'source' ('JWT'|'estimated'),
        or None if no token available.
    """
    with _token_lock:
        token = _shared_auth_state.get("token")

    if not token:
        return None

    result = {
        "token": token,
        "expires_on": None,
        "source": "unknown",
    }

    # Try JWT decode
    jwt_expiry = _decode_jwt_expiry(token)
    if jwt_expiry:
        result["expires_on"] = jwt_expiry
        result["source"] = "JWT"
        return result

    # Fallback: estimate 50-minute expiry from issue time
    issue_time = _auth_config.get("token_issue_time")
    if issue_time:
        result["expires_on"] = issue_time + timedelta(minutes=50)
    else:
        result["expires_on"] = datetime.now(timezone.utc) + timedelta(minutes=50)
    result["source"] = "estimated"

    return result


def invoke_token_refresh(*, force: bool = False) -> dict[str, Any]:
    """
    Force re-authentication for AppRegistration auth mode to get a fresh token.

    PS equivalent: Invoke-TokenRefresh (L8832)

    Args:
        force: Force re-authentication even if token doesn't appear expired.

    Returns:
        Dict with 'success' (bool), 'new_token' (str|None), 'message' (str),
        'auth_method' (str).
    """
    global _msal_app, _auth_failure_detected, _auth_401_message_shown, _token_acquired_time

    result = {
        "success": False,
        "new_token": None,
        "message": "",
        "auth_method": _auth_config.get("method"),
    }

    if not _auth_config.get("can_reauthenticate"):
        result["message"] = (
            f"Auth method '{_auth_config.get('method')}' does not support "
            "automatic re-authentication"
        )
        return result

    if _auth_config.get("method") != "AppRegistration":
        result["message"] = "Only AppRegistration auth mode supports automatic token refresh"
        return result

    tenant_id = _auth_config.get("tenant_id")
    client_id = _auth_config.get("client_id")
    if not tenant_id:
        result["message"] = "Missing tenant_id in stored auth config"
        return result
    if not client_id:
        result["message"] = "Missing client_id in stored auth config"
        return result

    logger.info("[TOKEN-REFRESH] Attempting re-authentication using AppRegistration...")

    try:
        # Re-acquire token using stored credentials (always force fresh to bypass MSAL cache)
        token_result = _acquire_token_app_registration(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=_auth_config.get("client_secret"),
            scopes=["https://graph.microsoft.com/.default"],
            force_fresh=True,
        )

        new_token = token_result.get("access_token")
        if new_token:
            result["success"] = True
            result["new_token"] = new_token
            result["message"] = "Successfully refreshed token"

            # Update state — mirrors PS: TokenIssueTime, TokenAcquiredTime,
            # AuthFailureDetected, Auth401MessageShown
            _auth_config["token_issue_time"] = datetime.now(timezone.utc)
            _token_acquired_time = datetime.now(timezone.utc)
            _auth_failure_detected = False
            _auth_401_message_shown = False  # Reset for next auth failure cycle

            logger.info("[TOKEN-REFRESH] Successfully obtained fresh access token")
        else:
            result["message"] = "Re-authenticated but could not extract access token"
            logger.error("[TOKEN-REFRESH] %s", result["message"])

    except Exception as e:
        result["message"] = f"Re-authentication failed: {e}"
        logger.error("[TOKEN-REFRESH] %s", result["message"])

    return result


def refresh_graph_token_if_needed(*, buffer_minutes: int = 5, force: bool = False) -> bool | str:
    """
    Proactively refresh the Graph access token if nearing expiry.

    PS equivalent: Refresh-GraphTokenIfNeeded (L8990)

    Args:
        buffer_minutes: Refresh if token expires within this many minutes. Default: 5.
        force: If True and AppReg refresh fails, returns 'Quit' (headless/fatal).
               If False and AppReg refresh fails, returns False (caller can prompt).
               Mirrors PS script-level $Force parameter behavior.

    Returns:
        True — Token was refreshed successfully.
        False — No refresh needed (token still valid), within cooldown, or non-fatal failure.
        'Quit' — Fatal failure (AppRegistration + force mode).
    """
    global _last_proactive_refresh_attempt, _shared_auth_state

    with _token_lock:
        expires_on = _shared_auth_state.get("expires_on")

    if not expires_on:
        return False

    now = datetime.now(timezone.utc)
    minutes_remaining = (expires_on - now).total_seconds() / 60.0

    # Proactive refresh for AppReg: refresh at 30-minute token age
    needs_proactive_refresh = False
    if (_auth_config.get("method") == "AppRegistration"
            and _auth_config.get("can_reauthenticate")):
        token_issue_time = _auth_config.get("token_issue_time")
        if token_issue_time:
            token_age_minutes = (datetime.now(timezone.utc) - token_issue_time).total_seconds() / 60.0
            if token_age_minutes > 30:
                needs_proactive_refresh = True
                logger.warning(
                    "[TOKEN] Token age: %.1f minutes - proactive refresh triggered",
                    token_age_minutes,
                )

    if minutes_remaining > buffer_minutes and not needs_proactive_refresh:
        return False  # Token still valid

    # Cooldown check (1 minute for AppRegistration)
    cooldown_minutes = 1.0
    if _last_proactive_refresh_attempt:
        time_since_last = (datetime.now(timezone.utc) - _last_proactive_refresh_attempt).total_seconds() / 60.0
        if time_since_last < cooldown_minutes:
            return False

    _last_proactive_refresh_attempt = datetime.now(timezone.utc)

    if not needs_proactive_refresh:
        logger.warning(
            "[TOKEN] Token expires in %.1f minutes - attempting proactive refresh...",
            minutes_remaining,
        )

    # Phase 1: Try silent token acquisition (mirrors PS Get-GraphAccessTokenWithExpiry check)
    # In MSAL-direct mode, this is handled by acquire_token_silent inside invoke_token_refresh.
    # But we also check if the token differs from current and validate staleness (>2 min).
    if _auth_config.get("can_reauthenticate"):
        refresh_result = invoke_token_refresh(force=True)
        if refresh_result["success"]:
            new_token = refresh_result["new_token"]
            jwt_expiry = _decode_jwt_expiry(new_token) if new_token else None
            new_expires = jwt_expiry or (datetime.now(timezone.utc) + timedelta(minutes=50))

            # Stale token validation: new token must have > 2 min remaining AND
            # its expiry must have actually advanced vs. the previous token.
            # The second check catches MSAL/AAD returning the SAME cached JWT
            # (which is what causes reactive 401 retries to spin in a loop).
            now_utc = datetime.now(timezone.utc)
            minutes_until_new_expiry = (new_expires - now_utc).total_seconds() / 60.0
            prior_expires_on = expires_on  # captured at top of function
            exp_advanced = (prior_expires_on is None) or (new_expires > prior_expires_on)
            if minutes_until_new_expiry <= 2:
                logger.error(
                    "[TOKEN] WARNING: Refreshed token is already expired or near-expiry "
                    "(expires in %.1f min) - token rejected",
                    minutes_until_new_expiry,
                )
                # Fall through to failure handling below
            elif not exp_advanced:
                logger.error(
                    "[TOKEN] WARNING: Refreshed token exp did not advance "
                    "(prior=%s, new=%s) — same cached JWT, rejected to prevent retry loop",
                    prior_expires_on.isoformat() if prior_expires_on else "none",
                    new_expires.isoformat(),
                )
                # Fall through to failure handling below
            else:
                with _token_lock:
                    _shared_auth_state["token"] = new_token
                    _shared_auth_state["expires_on"] = new_expires
                    _shared_auth_state["last_refresh"] = datetime.now(timezone.utc)
                    _shared_auth_state["refresh_count"] += 1

                _auth_config["token_issue_time"] = datetime.now(timezone.utc)
                logger.info(
                    "[TOKEN] Token refreshed via AppRegistration (refresh #%d, "
                    "new expiry in %.1f min)",
                    _shared_auth_state["refresh_count"],
                    minutes_until_new_expiry,
                )
                return True

    # Silent refresh failed
    logger.error("[TOKEN] Silent token refresh failed - interactive re-authentication required")

    # AppRegistration + force: FATAL (headless mode, no interactive fallback)
    if _auth_config.get("method") == "AppRegistration" and force:
        logger.error(
            "[TOKEN] FATAL: AppRegistration token refresh failed. "
            "Cannot continue headless (force mode)."
        )
        return "Quit"

    # AppReg without force: return False — caller may prompt
    return False


# ---------------------------------------------------------------------------
# Accessors for module state (for testing / external reads)
# ---------------------------------------------------------------------------

def get_auth_config() -> dict[str, Any]:
    """Return a reference to the current auth config."""
    return _auth_config


def get_shared_auth_state() -> dict[str, Any]:
    """Return a reference to the shared auth state."""
    return _shared_auth_state


def update_shared_auth_state(*, token: str, expires_on: datetime | None = None) -> None:
    """Update shared auth state with a refreshed token.

    Called by 401 handlers that use invoke_token_refresh directly,
    ensuring _shared_auth_state stays consistent.
    Mirrors PS Refresh-GraphTokenIfNeeded L9091-9098 SharedAuthState update.
    """
    jwt_expiry = _decode_jwt_expiry(token)
    new_expires = expires_on or jwt_expiry or (datetime.now(timezone.utc) + timedelta(minutes=50))
    with _token_lock:
        _shared_auth_state["token"] = token
        _shared_auth_state["expires_on"] = new_expires
        _shared_auth_state["last_refresh"] = datetime.now(timezone.utc)
        _shared_auth_state["refresh_count"] += 1
    _auth_config["token_issue_time"] = datetime.now(timezone.utc)
    logger.info(
        "[TOKEN] Shared auth state updated (refresh #%d, expires: %s)",
        _shared_auth_state["refresh_count"],
        new_expires.strftime("%H:%M:%S") if new_expires else "unknown",
    )


def is_connected() -> bool:
    """Return whether authentication was successful."""
    return _connected


def get_api_version() -> Optional[str]:
    """Return the detected Graph API version."""
    return _graph_audit_api_version


def reset_auth_state() -> None:
    """Reset all module-level auth state (for testing)."""
    global _graph_audit_api_version, _connected, _msal_app
    global _last_proactive_refresh_attempt, _auth_failure_detected
    global _auth_401_message_shown, _token_acquired_time
    global _defer_auth_context_display, _phase1_context

    _graph_audit_api_version = None
    _connected = False
    _msal_app = None
    _last_proactive_refresh_attempt = None
    _auth_failure_detected = False
    _auth_401_message_shown = False
    _token_acquired_time = None
    _defer_auth_context_display = False
    _phase1_context = None

    _auth_config.clear()
    _auth_config.update({
        "method": None,
        "tenant_id": None,
        "client_id": None,
        "client_secret": None,
        "can_reauthenticate": False,
        "token_issue_time": None,
    })

    _shared_auth_state.clear()
    _shared_auth_state.update({
        "token": None,
        "expires_on": None,
        "last_refresh": None,
        "refresh_count": 0,
        "auth_method": None,
    })


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _decode_jwt_expiry(token: Optional[str]) -> Optional[datetime]:
    """
    Decode JWT token to extract the 'exp' claim.

    PS equivalent: JWT decode logic in Get-GraphAccessTokenWithExpiry (L7575-7610)

    Returns:
        datetime (UTC) of token expiry, or None if decode fails.
    """
    if not token:
        return None

    try:
        parts = token.split(".")
        if len(parts) < 2:
            return None

        # Decode payload (second part) — base64url
        payload_b64 = parts[1]
        # Add padding
        padding_needed = 4 - (len(payload_b64) % 4)
        if padding_needed < 4:
            payload_b64 += "=" * padding_needed

        # base64url → standard base64
        payload_b64 = payload_b64.replace("-", "+").replace("_", "/")

        payload_bytes = base64.b64decode(payload_b64)
        payload = json.loads(payload_bytes.decode("utf-8"))

        exp = payload.get("exp")
        if exp:
            return datetime.fromtimestamp(int(exp), tz=timezone.utc)

    except Exception:
        pass

    return None


def _connect_app_registration(
    *,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    scopes: list[str],
) -> dict[str, Any]:
    """
    Authenticate using client-credentials flow with a client_secret.
    Uses MSAL ConfidentialClientApplication.
    """
    global _msal_app

    try:
        import msal
    except ImportError:
        raise RuntimeError(
            "The 'msal' package is required for authentication. "
            "Install it with: pip install msal"
        )

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    logger.info("  -> Authenticating with client secret")
    _msal_app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )

    return _acquire_token_app_registration(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        scopes=scopes,
    )


def _acquire_token_app_registration(
    *,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    scopes: list[str],
    force_fresh: bool = False,
) -> dict[str, Any]:
    """Acquire token using the cached or new MSAL app."""
    global _msal_app

    try:
        import msal
    except ImportError:
        raise RuntimeError("The 'msal' package is required. Install: pip install msal")

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    # Build MSAL app if not cached
    if _msal_app is None:
        _msal_app = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )

    # Try silent first, then fresh
    if force_fresh:
        # RC4 fix: bypass MSAL token cache — mirrors PS Disconnect-MgGraph + Connect-MgGraph.
        # PS destroys the MgGraph session (Disconnect-MgGraph) then creates a new one
        # (Connect-MgGraph) to force Azure AD to issue a fresh JWT. Python equivalent:
        # destroy the cached MSAL app so a new one is built, then call
        # acquire_token_for_client on the fresh instance (no cached tokens).
        _msal_app = None  # Mirrors PS Disconnect-MgGraph — destroy cached session
        _msal_app = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret,
        )
        # Fresh app, no cache — this always hits Azure AD for a new token
        result = _msal_app.acquire_token_for_client(scopes=scopes)
    else:
        result = _msal_app.acquire_token_silent(scopes, account=None)
        if not result or "access_token" not in result:
            result = _msal_app.acquire_token_for_client(scopes=scopes)

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Token acquisition failed: {error}")

    return result


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Test JWT decode
    # Create a fake JWT with exp claim
    import time as _time

    header = base64.urlsafe_b64encode(json.dumps({"alg": "RS256"}).encode()).rstrip(b"=")
    exp_time = int(_time.time()) + 3600  # 1 hour from now
    payload = base64.urlsafe_b64encode(
        json.dumps({"exp": exp_time, "sub": "test"}).encode()
    ).rstrip(b"=")
    fake_jwt = f"{header.decode()}.{payload.decode()}.fakesignature"

    expiry = _decode_jwt_expiry(fake_jwt)
    assert expiry is not None
    assert abs((expiry - datetime.now(timezone.utc)).total_seconds() - 3600) < 5

    # Test None input
    assert _decode_jwt_expiry(None) is None
    assert _decode_jwt_expiry("") is None
    assert _decode_jwt_expiry("not.a.jwt") is None

    # Test get_graph_audit_api_uri (no http_client — defaults to v1.0)
    reset_auth_state()
    uri = get_graph_audit_api_uri("queries")
    assert uri == "https://graph.microsoft.com/v1.0/security/auditLog/queries"

    uri2 = get_graph_audit_api_uri("queries/abc123/records")
    assert uri2 == "https://graph.microsoft.com/v1.0/security/auditLog/queries/abc123/records"

    # Test get_graph_access_token (no token set)
    reset_auth_state()
    assert get_graph_access_token() is None

    # Test get_graph_access_token_with_expiry (no token)
    assert get_graph_access_token_with_expiry() is None

    # Test invoke_token_refresh (not configured)
    reset_auth_state()
    result = invoke_token_refresh()
    assert result["success"] is False

    # Test refresh_graph_token_if_needed (no state)
    reset_auth_state()
    assert refresh_graph_token_if_needed() is False

    # Test state accessors
    assert get_auth_config()["method"] is None
    assert get_shared_auth_state()["token"] is None
    assert is_connected() is False
    assert get_api_version() is None

    # Test rejection of removed auth methods
    try:
        connect_purview_audit(auth_method="WebLogin", tenant_id="t", client_id="c", client_secret="s")
        assert False, "Should have rejected WebLogin"
    except ValueError:
        pass
    try:
        connect_purview_audit(auth_method="ManagedIdentity", tenant_id="t", client_id="c", client_secret="s")
        assert False, "Should have rejected ManagedIdentity"
    except ValueError:
        pass

    # Test missing client_secret
    reset_auth_state()
    try:
        connect_purview_audit(auth_method="AppRegistration", tenant_id="t", client_id="c", client_secret=None)
        assert False, "Should have rejected missing client_secret"
    except ValueError:
        pass

    print("PAX Auth Module - OK (all self-tests passed)")
