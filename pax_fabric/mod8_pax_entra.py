"""
Module 8: pax_entra — Entra ID User & License Data
====================================================
Migrated from: PAX_Purview_Audit_Log_Processor_v1.11.1.ps1 Lines 10319–10712
Level: 2 (depends on pax_graph_api)

Provides Entra ID (Azure AD) user directory enrichment:
- Fetch tenant SKU + Copilot service plan discovery from /subscribedSkus
- Paginated user fetch with license + assignedPlans data
- Flatten nested Graph user objects into 47-column flat dicts
- Enrich with MAC-format license columns (assignedLicenses, hasLicense)
- Schema validation against expected 47-column header

External dependencies: httpx (preferred) or requests for HTTP calls
Design: HTTP client is injected (same pattern as Module 7).
        Uses stdlib logging.getLogger(__name__).

PS-to-Python Function Mapping
──────────────────────────────────────────────────────────────────────────
│ # │ PS Function                │ PS Line │ Python Function                 │
│───│───────────────────────────│─────────│─────────────────────────────────│
│63 │ Get-UserLicenseData        │ 10319   │ get_user_license_data()         │
│64 │ ConvertTo-FlatEntraUsers   │ 10454   │ convert_to_flat_entra_users()   │
│65 │ Get-EntraUsersData         │ 10619   │ get_entra_users_data()          │
│   │ Test-EntraUsersSchema      │ 15203   │ test_entra_users_schema()       │
│   │ $EntraUsersHeader          │ 15190   │ ENTRA_USERS_HEADER              │
──────────────────────────────────────────────────────────────────────────

Test Results
────────────
Run: (pending)
"""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Any, Callable, Optional

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Reactive 401 retry helper (shared by both paged fetchers)
# ---------------------------------------------------------------------------

def _graph_get_with_refresh(
    http_client: Any,
    url: str,
    token_refresh_fn: Optional[Callable[[], bool]],
) -> Any:
    """GET ``url`` via ``http_client``; on HTTP 401, call ``token_refresh_fn``
    once and retry. Caller's refresh callback is responsible for updating
    the shared auth state and the http_client's Authorization header.
    Returns the response (already ``raise_for_status``-checked).
    """
    auth_retried = False
    while True:
        resp = http_client.get(url)
        status = getattr(resp, "status_code", None)
        if status == 401 and not auth_retried and token_refresh_fn is not None:
            auth_retried = True
            if token_refresh_fn():
                continue
        resp.raise_for_status()
        return resp


# ---------------------------------------------------------------------------
# Date formatting helper (PS parity)
# ---------------------------------------------------------------------------

def _format_graph_date(value: Any) -> Optional[str]:
    """Convert ISO 8601 date string to dd-MM-yyyy HH:mm:ss format.

    PS parity: When the Microsoft Graph PowerShell SDK deserializes JSON,
    date strings become [DateTime] objects. Export-Csv then serializes them
    using the system locale's default format, which is typically
    ``dd-MM-yyyy HH:mm:ss``.  Python's ``requests`` library returns raw
    JSON strings (ISO 8601), so we must explicitly reformat to match.

    Handles: '2026-01-14T12:58:29Z', '2026-01-14T12:58:29.1234567Z',
             '2024-07-22' (date-only, e.g. employeeHireDate).
    """
    if not value or not isinstance(value, str):
        return value
    try:
        # Try full ISO 8601 with time component
        # Strip trailing 'Z' and handle variable fractional-second precision
        clean = value.rstrip("Z")
        if "T" in clean:
            # Truncate fractional seconds to max 6 digits (Python limit)
            if "." in clean:
                main, frac = clean.rsplit(".", 1)
                frac = frac[:6]
                clean = f"{main}.{frac}"
            dt = datetime.fromisoformat(clean)
            return dt.strftime("%d-%m-%Y %H:%M:%S")
        else:
            # Date-only string like '2024-07-22'
            dt = datetime.strptime(clean, "%Y-%m-%d")
            return dt.strftime("%d-%m-%Y %H:%M:%S")
    except (ValueError, TypeError):
        return value  # Return as-is if unparseable


# ---------------------------------------------------------------------------
# Expected schema header (PS $EntraUsersHeader at L15190)
# 47 columns: 30 core + 5 manager + 2 license + 6 PBI alias + 4 Viva placeholder
# ---------------------------------------------------------------------------

ENTRA_USERS_HEADER: list[str] = [
    # Core Identity Properties
    "userPrincipalName", "displayName", "id", "mail", "givenName", "surname",
    # Job Properties
    "jobTitle", "department", "employeeType", "employeeId", "employeeHireDate",
    # Location Properties
    "officeLocation", "city", "state", "country", "postalCode", "companyName",
    # Organizational Properties (flattened from employeeOrgData)
    "employeeOrgData_division", "employeeOrgData_costCenter",
    # Status Properties
    "accountEnabled", "userType", "createdDateTime",
    # Usage Properties
    "usageLocation", "preferredLanguage",
    # Sync Properties
    "onPremisesSyncEnabled", "onPremisesImmutableId", "externalUserState",
    # Proxy Addresses (exploded)
    "proxyAddresses_Primary", "proxyAddresses_Count", "proxyAddresses_All",
    # Manager (flattened)
    "manager_id", "manager_displayName", "manager_userPrincipalName",
    "manager_mail", "manager_jobTitle",
    # License columns (added by Get-EntraUsersData via Get-UserLicenseData)
    "assignedLicenses", "hasLicense",
    # Power BI AI-in-One Dashboard 2701 Template Compatibility (alias mappings)
    "ManagerID", "BusinessAreaLabel", "CountryofEmployment",
    "CompanyCodeLabel", "CostCentreLabel", "UserName",
    # Viva Insights placeholders (not available from Graph API)
    "EffectiveDate", "FunctionType", "BusinessAreaCode", "OrgLevel_3Label",
]


# ---------------------------------------------------------------------------
# Graph API select fields for user fetch (PS $entraUserSelect at L10653)
# 27 properties + manager expansion
# ---------------------------------------------------------------------------

_ENTRA_USER_SELECT_FIELDS: list[str] = [
    "userPrincipalName", "displayName", "id", "mail", "givenName", "surname",
    "jobTitle", "department", "employeeType", "employeeId", "employeeHireDate",
    "officeLocation", "city", "state", "country", "postalCode", "companyName",
    "accountEnabled", "userType", "createdDateTime", "usageLocation",
    "preferredLanguage", "onPremisesSyncEnabled", "onPremisesImmutableId",
    "externalUserState", "employeeOrgData", "proxyAddresses",
]


# ---------------------------------------------------------------------------
# 64. ConvertTo-FlatEntraUsers (Line 10454)
# ---------------------------------------------------------------------------

def convert_to_flat_entra_users(
    users: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Flattens Entra user objects into CSV-friendly format.

    PS equivalent: ConvertTo-FlatEntraUsers (L10454)

    Converts Entra ID user objects with nested properties into flat tabular format.
    Filters out non-user accounts (rooms, resources) based on userType validation
    and name heuristics.

    Filtering rules (exact PS parity):
    1. Skip if userType is null/empty.
    2. Skip if userType is not 'Member' or 'Guest'.
    3. If Member/Guest but no givenName AND no surname AND accountEnabled:
       skip if no assignedLicenses (likely a room/resource).

    Args:
        users: List of user dicts from Microsoft Graph API
               (with 27 properties + manager expansion).

    Returns:
        List of flattened user dicts with 35 columns
        (30 core + 5 manager; license columns added later by get_entra_users_data).
    """
    flattened_users: list[dict[str, Any]] = []

    for user in users:
        # ── Filter: Only include real user accounts ──

        # PS: $userTypeValue = $user.userType
        user_type_value = user.get("userType")

        # PS: if ([string]::IsNullOrWhiteSpace($userTypeValue)) { continue }
        if not user_type_value or not str(user_type_value).strip():
            continue

        # PS: if ($userTypeValue -ne 'Member' -and $userTypeValue -ne 'Guest') { continue }
        if user_type_value not in ("Member", "Guest"):
            continue

        # PS: Additional heuristic — real users have givenName or surname
        has_given_name = bool(user.get("givenName") and str(user["givenName"]).strip())
        has_surname = bool(user.get("surname") and str(user["surname"]).strip())

        if not has_given_name and not has_surname and user.get("accountEnabled"):
            # No licenses and no name components → likely a room/resource
            assigned_licenses = user.get("assignedLicenses") or []
            if not assigned_licenses or len(assigned_licenses) == 0:
                continue

        # ── Build flat user dict ──
        flat_user: dict[str, Any] = {}

        # Core Identity Properties
        flat_user["userPrincipalName"] = user.get("userPrincipalName")
        flat_user["displayName"] = user.get("displayName")
        flat_user["id"] = user.get("id")
        flat_user["mail"] = user.get("mail")
        flat_user["givenName"] = user.get("givenName")
        flat_user["surname"] = user.get("surname")

        # Job Properties
        flat_user["jobTitle"] = user.get("jobTitle")
        flat_user["department"] = user.get("department")
        flat_user["employeeType"] = user.get("employeeType")
        flat_user["employeeId"] = user.get("employeeId")
        flat_user["employeeHireDate"] = _format_graph_date(user.get("employeeHireDate"))

        # Location Properties
        flat_user["officeLocation"] = user.get("officeLocation")
        flat_user["city"] = user.get("city")
        flat_user["state"] = user.get("state")
        flat_user["country"] = user.get("country")
        flat_user["postalCode"] = user.get("postalCode")
        flat_user["companyName"] = user.get("companyName")

        # Organizational Properties (flattened from nested object)
        emp_org = user.get("employeeOrgData")
        flat_user["employeeOrgData_division"] = (
            emp_org.get("division") if isinstance(emp_org, dict) else None
        )
        flat_user["employeeOrgData_costCenter"] = (
            emp_org.get("costCenter") if isinstance(emp_org, dict) else None
        )

        # Status Properties
        flat_user["accountEnabled"] = user.get("accountEnabled")
        flat_user["userType"] = user.get("userType")
        flat_user["createdDateTime"] = _format_graph_date(user.get("createdDateTime"))

        # Usage Properties
        flat_user["usageLocation"] = user.get("usageLocation")
        flat_user["preferredLanguage"] = user.get("preferredLanguage")

        # Sync Properties
        flat_user["onPremisesSyncEnabled"] = user.get("onPremisesSyncEnabled")
        flat_user["onPremisesImmutableId"] = user.get("onPremisesImmutableId")
        flat_user["externalUserState"] = user.get("externalUserState")

        # ── Proxy Addresses (explode array) ──
        proxy_addresses = user.get("proxyAddresses") or []
        if proxy_addresses and len(proxy_addresses) > 0:
            # PS: $user.proxyAddresses | Where-Object { $_ -like 'SMTP:*' } | Select-Object -First 1
            # NOTE: PS -like is CASE-INSENSITIVE, so 'smtp:*' also matches 'SMTP:*'.
            # This means the filter returns ALL proxy addresses (all start with smtp: or SMTP:),
            # and Select-Object -First 1 picks the first entry in the list.
            primary_smtp = None
            for addr in proxy_addresses:
                if isinstance(addr, str) and addr.upper().startswith("SMTP:"):
                    primary_smtp = addr[5:]  # Remove 'smtp:' or 'SMTP:' prefix
                    break
            flat_user["proxyAddresses_Primary"] = primary_smtp
            flat_user["proxyAddresses_Count"] = len(proxy_addresses)
            # PS: $user.proxyAddresses -join '; '
            flat_user["proxyAddresses_All"] = "; ".join(
                str(a) for a in proxy_addresses
            )
        else:
            flat_user["proxyAddresses_Primary"] = None
            flat_user["proxyAddresses_Count"] = 0
            flat_user["proxyAddresses_All"] = None

        # ── Manager (flatten nested object) ──
        manager = user.get("manager")
        if manager and isinstance(manager, dict):
            flat_user["manager_id"] = manager.get("id")
            flat_user["manager_displayName"] = manager.get("displayName")
            flat_user["manager_userPrincipalName"] = manager.get(
                "userPrincipalName"
            )
            flat_user["manager_mail"] = manager.get("mail")
            flat_user["manager_jobTitle"] = manager.get("jobTitle")
        else:
            flat_user["manager_id"] = None
            flat_user["manager_displayName"] = None
            flat_user["manager_userPrincipalName"] = None
            flat_user["manager_mail"] = None
            flat_user["manager_jobTitle"] = None

        # ── Power BI AI-in-One Dashboard 2701 Template Compatibility ──
        # Alias columns mapping existing data to 2701 template column names
        flat_user["ManagerID"] = flat_user["manager_id"]
        flat_user["BusinessAreaLabel"] = flat_user["employeeOrgData_division"]
        flat_user["CountryofEmployment"] = flat_user["country"]
        flat_user["CompanyCodeLabel"] = flat_user["companyName"]
        flat_user["CostCentreLabel"] = flat_user["employeeOrgData_costCenter"]
        flat_user["UserName"] = flat_user["displayName"]

        # Viva Insights placeholders (not available from Graph API)
        flat_user["EffectiveDate"] = None
        flat_user["FunctionType"] = None
        flat_user["BusinessAreaCode"] = None
        flat_user["OrgLevel_3Label"] = None

        flattened_users.append(flat_user)

    return flattened_users


# ---------------------------------------------------------------------------
# Test-EntraUsersSchema (Line 15203) — schema validation helper
# ---------------------------------------------------------------------------

def test_entra_users_schema(
    users: list[dict[str, Any]],
    *,
    quiet: bool = False,
) -> bool:
    """
    Validates that flattened user dicts have the expected schema.

    PS equivalent: Test-EntraUsersSchema (L15203)

    Non-fatal: logs warnings on schema mismatch, returns True/False.

    Args:
        users: List of flattened user dicts.
        quiet: If True, suppress success messages.

    Returns:
        True if schema matches expected header, False otherwise.
    """
    if not users:
        return True

    expected = set(ENTRA_USERS_HEADER)
    actual = set(users[0].keys())

    missing = sorted(expected - actual)
    extra = sorted(actual - expected)

    if missing or extra:
        logger.warning(
            "WARNING: EntraUsers schema mismatch. Missing: %s; Extra: %s",
            ", ".join(missing) if missing else "(none)",
            ", ".join(extra) if extra else "(none)",
        )
        return False
    elif not quiet:
        logger.info(
            "Validated EntraUsers schema (%d columns).", len(expected)
        )
    return True


# ---------------------------------------------------------------------------
# 65. Get-EntraUsersData (Line 10619)
# ---------------------------------------------------------------------------

def get_entra_users_data(
    *,
    http_client: Any = None,
    license_data: Optional[dict[str, dict]] = None,
    quiet: bool = False,
    token_refresh_fn: Optional[Callable[[], bool]] = None,
) -> list[dict[str, Any]]:
    """
    Collects and flattens Entra ID user directory data, then enriches
    with MAC-format license columns.

    PS equivalent: Get-EntraUsersData (L10619)

    Workflow:
    1. Paginate /users with 27 $select fields + $expand=manager.
    2. Flatten via convert_to_flat_entra_users (filters rooms/resources).
    3. Enrich each user with:
       - assignedLicenses: semicolon-separated SKU names
       - hasLicense: Copilot license boolean
    4. Validate schema via test_entra_users_schema (non-fatal).

    Args:
        http_client: HTTP client with .get() method.
        license_data: Optional dict with 'UserLicenses' and 'UserHasCopilot' keys.
        quiet: If True, suppress info-level log messages.

    Returns:
        List of flattened + enriched user dicts (47 columns),
        or empty list on failure.
    """
    entra_users: list[dict[str, Any]] = []

    try:
        if not quiet:
            logger.info(
                "Fetching Entra user directory (35 properties + manager)..."
            )

        if http_client is None:
            logger.error("ERROR: No HTTP client provided for Entra user fetch")
            return []

        # Build $select parameter (PS L10653)
        select_param = ",".join(_ENTRA_USER_SELECT_FIELDS)
        base_uri = (
            f"https://graph.microsoft.com/v1.0/users"
            f"?$select={select_param}"
            f"&$expand=manager"
            f"&$top=999"
        )

        next_link: Optional[str] = base_uri
        raw_users: list[dict[str, Any]] = []
        loops = 0

        while next_link:
            loops += 1
            # PS safety abort: if ($loops -gt 2000) { throw "..." }
            if loops > 2000:
                raise RuntimeError("Safety abort: excessive paging (>2000)")

            resp = _graph_get_with_refresh(
                http_client, next_link, token_refresh_fn
            )
            resp_data = resp.json()

            if resp_data.get("value"):
                raw_users.extend(resp_data["value"])

            next_link = resp_data.get("@odata.nextLink")

        if not quiet:
            logger.info("  Retrieved %d raw user objects", len(raw_users))

        # ── Flatten ──
        flattened = convert_to_flat_entra_users(raw_users)
        if not quiet:
            logger.info(
                "  Flattened to %d user rows (filtered)", len(flattened)
            )

        # ── License enrichment (MAC-format columns) ──
        for u in flattened:
            upn = u.get("userPrincipalName")
            user_id = u.get("id")
            assigned_names: Optional[str] = None
            has_copilot = False

            if license_data:
                ul = license_data.get("UserLicenses", {})
                uhc = license_data.get("UserHasCopilot", {})

                # Lookup by UPN first, then by id (PS L10691-10694)
                if upn and upn in ul:
                    assigned_names = ";".join(ul[upn])
                elif user_id and user_id in ul:
                    assigned_names = ";".join(ul[user_id])

                # Copilot lookup by UPN first, then by id (PS L10695-10698)
                if upn and upn in uhc:
                    has_copilot = bool(uhc[upn])
                elif user_id and user_id in uhc:
                    has_copilot = bool(uhc[user_id])

            u["assignedLicenses"] = assigned_names
            u["hasLicense"] = has_copilot

        entra_users = flattened

        # ── Schema validation (non-fatal) ──
        try:
            test_entra_users_schema(entra_users, quiet=quiet)
        except Exception:
            pass

    except Exception as e:
        logger.warning(
            "WARNING: Failed to collect Entra user directory: %s", e
        )

    return entra_users
