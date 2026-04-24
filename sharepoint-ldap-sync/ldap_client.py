"""
LDAP / Active Directory client.

Responsibilities
----------------
* Open (and keep open) a single authenticated connection to the directory.
* Execute a search-by-sAMAccountName query for a given username.
* Parse the raw ldap3 entry and return a flat ``dict[str, str]`` of the
  requested attributes, suitable for writing straight into Excel.
* Handle "user not found" and connection errors gracefully.
"""

from __future__ import annotations

import datetime
import logging
import struct
from typing import Dict, List, Optional

import ldap3
from ldap3 import Connection, Server, SUBTREE, ALL_ATTRIBUTES, NTLM, SIMPLE
from ldap3.core.exceptions import LDAPException

from config import LDAPConfig

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Bit flags for userAccountControl
_UAC_DISABLED = 0x0002

# Windows FILETIME epoch offset (100-ns intervals from 1601-01-01)
_FILETIME_EPOCH_DIFF = 116444736000000000

# ldap3 uses this sentinel when an attribute contains a raw bytes value that
# cannot be decoded (e.g. objectGUID, objectSid).  We skip those safely.
_LDAP3_BYTES_TYPE = bytes


# ---------------------------------------------------------------------------
# Public class
# ---------------------------------------------------------------------------

class LDAPClient:
    """Thin wrapper around an ldap3 Connection."""

    def __init__(self, cfg: LDAPConfig) -> None:
        self.cfg = cfg
        self._conn: Optional[Connection] = None

    # ------------------------------------------------------------------
    # Connection lifecycle
    # ------------------------------------------------------------------

    def connect(self) -> None:
        """Open and bind the LDAP connection."""
        server = Server(
            self.cfg.server,
            port=self.cfg.port,
            use_ssl=self.cfg.use_ssl,
            get_info=ldap3.ALL,
        )

        # Determine authentication type: NTLM if bind_dn looks like
        # DOMAIN\user, otherwise SIMPLE.
        if "\\" in self.cfg.bind_dn:
            auth_type = NTLM
        else:
            auth_type = SIMPLE

        self._conn = Connection(
            server,
            user=self.cfg.bind_dn,
            password=self.cfg.bind_password,
            authentication=auth_type,
            auto_bind=True,
            raise_exceptions=True,
        )
        logger.info(
            "Connected to LDAP server %s:%s (SSL=%s)",
            self.cfg.server,
            self.cfg.port,
            self.cfg.use_ssl,
        )

    def disconnect(self) -> None:
        """Unbind the connection if open."""
        if self._conn and self._conn.bound:
            self._conn.unbind()
            logger.info("LDAP connection closed.")
        self._conn = None

    def __enter__(self) -> "LDAPClient":
        self.connect()
        return self

    def __exit__(self, *_: object) -> None:
        self.disconnect()

    # ------------------------------------------------------------------
    # Query
    # ------------------------------------------------------------------

    def query_user(self, username: str) -> Optional[Dict[str, Optional[str]]]:
        """
        Search for *username* (sAMAccountName) and return a dictionary of
        attribute values.  Returns ``None`` if the user is not found.

        Raises ``LDAPException`` on connectivity / search errors.
        """
        if self._conn is None:
            raise RuntimeError("Not connected.  Call connect() or use as context manager.")

        search_filter = f"(&(objectClass=user)(sAMAccountName={ldap3.utils.conv.escape_filter_chars(username)}))"

        attributes_to_fetch = list(self.cfg.attributes) if self.cfg.attributes else [ALL_ATTRIBUTES]

        logger.debug("Querying LDAP for user '%s' …", username)

        self._conn.search(
            search_base=self.cfg.search_base,
            search_filter=search_filter,
            search_scope=SUBTREE,
            attributes=attributes_to_fetch,
        )

        entries = self._conn.entries
        if not entries:
            logger.warning("User '%s' not found in LDAP.", username)
            return None

        if len(entries) > 1:
            logger.warning(
                "Multiple entries found for '%s'; using the first one.", username
            )

        entry = entries[0]
        return self._parse_entry(entry)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _parse_entry(self, entry: ldap3.abstract.entry.Entry) -> Dict[str, Optional[str]]:
        """Convert an ldap3 Entry to a flat string dictionary."""
        result: Dict[str, Optional[str]] = {}

        for attr_name in self.cfg.attributes:
            try:
                attr = getattr(entry, attr_name, None)
                if attr is None or not attr.values:
                    result[attr_name] = None
                    continue

                raw = attr.value  # Single value or first of multi-valued

                if attr_name == "userAccountControl":
                    result[attr_name] = self._decode_uac(raw)
                elif attr_name in ("accountExpires", "lastLogonTimestamp"):
                    result[attr_name] = self._decode_filetime(raw)
                elif attr_name == "manager":
                    result[attr_name] = self._extract_cn(raw)
                elif attr_name == "memberOf":
                    # Join all group CNs into a semicolon-separated string
                    result[attr_name] = "; ".join(
                        self._extract_cn(v) for v in attr.values
                    )
                elif isinstance(raw, _LDAP3_BYTES_TYPE):
                    result[attr_name] = None  # skip non-decodable binary
                else:
                    result[attr_name] = str(raw)

            except Exception as exc:  # noqa: BLE001
                logger.debug("Could not parse attribute '%s': %s", attr_name, exc)
                result[attr_name] = None

        return result

    @staticmethod
    def _decode_uac(value: object) -> str:
        """Return 'Enabled' or 'Disabled' based on userAccountControl."""
        try:
            flags = int(value)
            return "Disabled" if flags & _UAC_DISABLED else "Enabled"
        except (TypeError, ValueError):
            return str(value)

    @staticmethod
    def _decode_filetime(value: object) -> Optional[str]:
        """Convert a Windows FILETIME integer to an ISO-8601 date string."""
        try:
            ft = int(value)
            if ft in (0, 9223372036854775807):  # "never"
                return "Never"
            # Convert 100-ns intervals since 1601-01-01 to Unix epoch seconds
            unix_ts = (ft - _FILETIME_EPOCH_DIFF) / 10_000_000
            dt = datetime.datetime(1970, 1, 1, tzinfo=datetime.timezone.utc) + datetime.timedelta(seconds=unix_ts)
            return dt.strftime("%Y-%m-%d %H:%M:%S UTC")
        except (TypeError, ValueError, OverflowError, OSError):
            return str(value)

    @staticmethod
    def _extract_cn(dn: str) -> str:
        """Extract the CN component from a Distinguished Name."""
        for part in str(dn).split(","):
            if part.strip().upper().startswith("CN="):
                return part.strip()[3:]
        return str(dn)

    # ------------------------------------------------------------------
    # Batch helper
    # ------------------------------------------------------------------

    def query_users(
        self, usernames: List[str]
    ) -> Dict[str, Optional[Dict[str, Optional[str]]]]:
        """
        Query multiple users and return a ``{username: attributes}`` dict.
        Users that are not found are mapped to ``None``.
        """
        results: Dict[str, Optional[Dict[str, Optional[str]]]] = {}
        for username in usernames:
            try:
                results[username] = self.query_user(username)
            except LDAPException as exc:
                logger.error("LDAP error while querying '%s': %s", username, exc)
                results[username] = None
        return results
