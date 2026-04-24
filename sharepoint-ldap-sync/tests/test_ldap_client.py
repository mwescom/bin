"""
Unit tests for the LDAPClient class.

The LDAP server is mocked using ldap3's MockSyncStrategy so no real
directory is required.
"""

from __future__ import annotations

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest
import ldap3
from ldap3 import Server, Connection, MOCK_SYNC, OFFLINE_AD_2012_R2
from unittest.mock import MagicMock, patch

from config import LDAPConfig, DEFAULT_LDAP_ATTRIBUTES
from ldap_client import LDAPClient


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

DEFAULT_LDAP_CFG = LDAPConfig(
    server="ldap.example.com",
    port=389,
    use_ssl=False,
    bind_dn="CN=svc,DC=example,DC=com",
    bind_password="password",
    search_base="DC=example,DC=com",
    attributes=["displayName", "mail", "userAccountControl", "accountExpires"],
)


def _make_mock_connection(cfg: LDAPConfig) -> Connection:
    """Return an ldap3 connection using the in-memory MOCK_SYNC strategy."""
    server = Server("mock_server", get_info=OFFLINE_AD_2012_R2)
    conn = Connection(
        server,
        user=cfg.bind_dn,
        password=cfg.bind_password,
        client_strategy=MOCK_SYNC,
        auto_bind=True,
    )
    return conn


# ---------------------------------------------------------------------------
# Tests for _decode_uac
# ---------------------------------------------------------------------------

class TestDecodeUAC:
    def test_enabled(self):
        assert LDAPClient._decode_uac(512) == "Enabled"

    def test_disabled(self):
        # 0x0202 = NORMAL_ACCOUNT | DISABLED
        assert LDAPClient._decode_uac(514) == "Disabled"

    def test_bad_value(self):
        result = LDAPClient._decode_uac("not-a-number")
        assert result == "not-a-number"


# ---------------------------------------------------------------------------
# Tests for _decode_filetime
# ---------------------------------------------------------------------------

class TestDecodeFiletime:
    def test_never_zero(self):
        assert LDAPClient._decode_filetime(0) == "Never"

    def test_never_max(self):
        assert LDAPClient._decode_filetime(9223372036854775807) == "Never"

    def test_known_date(self):
        # 2023-01-01 00:00:00 UTC
        # Unix timestamp = 1672531200
        # FILETIME = (1672531200 * 10_000_000) + 116444736000000000
        ft = 1672531200 * 10_000_000 + 116444736000000000
        result = LDAPClient._decode_filetime(ft)
        assert "2023-01-01" in result


# ---------------------------------------------------------------------------
# Tests for _extract_cn
# ---------------------------------------------------------------------------

class TestExtractCN:
    def test_typical_dn(self):
        dn = "CN=John Smith,OU=Users,DC=example,DC=com"
        assert LDAPClient._extract_cn(dn) == "John Smith"

    def test_no_cn(self):
        dn = "OU=Users,DC=example,DC=com"
        assert LDAPClient._extract_cn(dn) == dn


# ---------------------------------------------------------------------------
# Tests for query_user (using mocked Connection)
# ---------------------------------------------------------------------------

class TestQueryUser:
    def _make_client_with_mock_conn(self, mock_conn: Connection) -> LDAPClient:
        client = LDAPClient(DEFAULT_LDAP_CFG)
        client._conn = mock_conn
        return client

    def test_not_connected_raises(self):
        client = LDAPClient(DEFAULT_LDAP_CFG)
        with pytest.raises(RuntimeError):
            client.query_user("jsmith")

    def test_user_not_found_returns_none(self):
        mock_conn = MagicMock()
        mock_conn.entries = []
        client = self._make_client_with_mock_conn(mock_conn)
        result = client.query_user("nonexistent")
        assert result is None

    def test_user_found_returns_dict(self):
        # Build a fake ldap3 entry
        mock_entry = MagicMock()

        def _make_attr(value):
            attr = MagicMock()
            attr.value = value
            attr.values = [value]
            return attr

        mock_entry.displayName = _make_attr("John Smith")
        mock_entry.mail = _make_attr("jsmith@example.com")
        mock_entry.userAccountControl = _make_attr(512)  # Enabled
        mock_entry.accountExpires = _make_attr(0)  # Never

        mock_conn = MagicMock()
        mock_conn.entries = [mock_entry]
        client = self._make_client_with_mock_conn(mock_conn)

        result = client.query_user("jsmith")
        assert result is not None
        assert result["displayName"] == "John Smith"
        assert result["mail"] == "jsmith@example.com"
        assert result["userAccountControl"] == "Enabled"
        assert result["accountExpires"] == "Never"

    def test_multiple_entries_uses_first(self):
        mock_entry1 = MagicMock()
        mock_entry1.displayName = MagicMock(value="First User", values=["First User"])
        mock_entry1.mail = MagicMock(value="first@example.com", values=["first@example.com"])
        mock_entry1.userAccountControl = MagicMock(value=512, values=[512])
        mock_entry1.accountExpires = MagicMock(value=0, values=[0])

        mock_entry2 = MagicMock()

        mock_conn = MagicMock()
        mock_conn.entries = [mock_entry1, mock_entry2]
        client = self._make_client_with_mock_conn(mock_conn)

        result = client.query_user("duplicate")
        assert result is not None
        assert result["displayName"] == "First User"


# ---------------------------------------------------------------------------
# Tests for query_users (batch)
# ---------------------------------------------------------------------------

class TestQueryUsers:
    def test_batch_returns_all_keys(self):
        client = LDAPClient(DEFAULT_LDAP_CFG)

        def _mock_query(username: str):
            if username == "jsmith":
                return {"displayName": "John Smith"}
            return None

        client._conn = MagicMock()  # prevent "not connected" check
        client.query_user = _mock_query  # type: ignore[method-assign]

        results = client.query_users(["jsmith", "missing"])
        assert "jsmith" in results
        assert "missing" in results
        assert results["jsmith"] == {"displayName": "John Smith"}
        assert results["missing"] is None
