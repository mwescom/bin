"""
Configuration management for the SharePoint-LDAP sync tool.

All settings are read from environment variables.  A .env file in the
working directory (or any parent directory) is loaded automatically when
python-dotenv is installed.  See .env.example for the full list of
required and optional variables.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import List

from dotenv import load_dotenv

load_dotenv()


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

def _require(name: str) -> str:
    """Return the value of *name* from the environment, or raise on missing."""
    value = os.getenv(name)
    if not value:
        raise EnvironmentError(
            f"Required environment variable '{name}' is not set. "
            "Copy .env.example to .env and fill in the values."
        )
    return value


def _optional(name: str, default: str = "") -> str:
    return os.getenv(name, default)


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------

@dataclass
class SharePointConfig:
    """Settings for connecting to SharePoint Online."""

    site_url: str          # e.g. https://contoso.sharepoint.com/sites/MySite
    client_id: str         # Azure AD application (client) ID
    client_secret: str     # Azure AD application secret
    # Path to the Excel file *inside* SharePoint, relative to the document
    # library root.  Example: "Shared Documents/Users/user_list.xlsx"
    file_path: str


@dataclass
class LDAPConfig:
    """Settings for the LDAP / Active Directory server."""

    server: str            # hostname or IP
    port: int              # 389 = plain, 636 = LDAPS
    use_ssl: bool
    bind_dn: str           # service-account DN or "DOMAIN\\user"
    bind_password: str
    search_base: str       # e.g. "DC=contoso,DC=com"
    # Attributes to fetch for each user.  Defaults shown in .env.example.
    attributes: List[str] = field(default_factory=list)


@dataclass
class ExcelConfig:
    """Column layout of the Excel workbook."""

    username_column: str   # Column letter that holds the usernames, e.g. "A"
    header_row: int        # Row number of the header row, e.g. 1
    data_start_row: int    # First data row, e.g. 2


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

DEFAULT_LDAP_ATTRIBUTES = [
    "displayName",
    "mail",
    "department",
    "title",
    "manager",
    "telephoneNumber",
    "userAccountControl",
    "accountExpires",
    "whenCreated",
    "lastLogonTimestamp",
    "memberOf",
]


def load_config() -> tuple[SharePointConfig, LDAPConfig, ExcelConfig]:
    """
    Read all configuration from environment variables and return typed
    config objects.  Raises ``EnvironmentError`` if any required variable
    is missing.
    """

    sp = SharePointConfig(
        site_url=_require("SP_SITE_URL"),
        client_id=_require("SP_CLIENT_ID"),
        client_secret=_require("SP_CLIENT_SECRET"),
        file_path=_require("SP_FILE_PATH"),
    )

    raw_attrs = _optional("LDAP_ATTRIBUTES")
    attributes = (
        [a.strip() for a in raw_attrs.split(",") if a.strip()]
        if raw_attrs
        else DEFAULT_LDAP_ATTRIBUTES
    )

    ldap = LDAPConfig(
        server=_require("LDAP_SERVER"),
        port=int(_optional("LDAP_PORT", "389")),
        use_ssl=_optional("LDAP_USE_SSL", "false").lower() in ("1", "true", "yes"),
        bind_dn=_require("LDAP_BIND_DN"),
        bind_password=_require("LDAP_BIND_PASSWORD"),
        search_base=_require("LDAP_SEARCH_BASE"),
        attributes=attributes,
    )

    excel = ExcelConfig(
        username_column=_optional("EXCEL_USERNAME_COLUMN", "A"),
        header_row=int(_optional("EXCEL_HEADER_ROW", "1")),
        data_start_row=int(_optional("EXCEL_DATA_START_ROW", "2")),
    )

    return sp, ldap, excel
