"""
Unit tests for the config module.
"""

from __future__ import annotations

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import pytest


class TestLoadConfig:
    """Tests for load_config() environment variable handling."""

    REQUIRED_ENV = {
        "SP_SITE_URL": "https://contoso.sharepoint.com/sites/MySite",
        "SP_CLIENT_ID": "client-id",
        "SP_CLIENT_SECRET": "client-secret",
        "SP_FILE_PATH": "Shared Documents/file.xlsx",
        "LDAP_SERVER": "dc01.contoso.com",
        "LDAP_BIND_DN": "CONTOSO\\svc",
        "LDAP_BIND_PASSWORD": "pass",
        "LDAP_SEARCH_BASE": "DC=contoso,DC=com",
    }

    def test_all_required_present(self, monkeypatch):
        for k, v in self.REQUIRED_ENV.items():
            monkeypatch.setenv(k, v)

        from config import load_config
        sp, ldap, excel = load_config()

        assert sp.site_url == "https://contoso.sharepoint.com/sites/MySite"
        assert sp.client_id == "client-id"
        assert sp.file_path == "Shared Documents/file.xlsx"
        assert ldap.server == "dc01.contoso.com"
        assert ldap.search_base == "DC=contoso,DC=com"
        assert ldap.port == 389  # default
        assert ldap.use_ssl is False  # default
        assert excel.username_column == "A"  # default
        assert excel.header_row == 1          # default
        assert excel.data_start_row == 2      # default

    def test_missing_required_raises(self, monkeypatch):
        for k in self.REQUIRED_ENV:
            monkeypatch.delenv(k, raising=False)

        from config import load_config
        with pytest.raises(EnvironmentError):
            load_config()

    def test_custom_ldap_port_and_ssl(self, monkeypatch):
        for k, v in self.REQUIRED_ENV.items():
            monkeypatch.setenv(k, v)
        monkeypatch.setenv("LDAP_PORT", "636")
        monkeypatch.setenv("LDAP_USE_SSL", "true")

        from config import load_config
        _, ldap, _ = load_config()
        assert ldap.port == 636
        assert ldap.use_ssl is True

    def test_custom_excel_settings(self, monkeypatch):
        for k, v in self.REQUIRED_ENV.items():
            monkeypatch.setenv(k, v)
        monkeypatch.setenv("EXCEL_USERNAME_COLUMN", "C")
        monkeypatch.setenv("EXCEL_HEADER_ROW", "2")
        monkeypatch.setenv("EXCEL_DATA_START_ROW", "3")

        from config import load_config
        _, _, excel = load_config()
        assert excel.username_column == "C"
        assert excel.header_row == 2
        assert excel.data_start_row == 3

    def test_custom_ldap_attributes(self, monkeypatch):
        for k, v in self.REQUIRED_ENV.items():
            monkeypatch.setenv(k, v)
        monkeypatch.setenv("LDAP_ATTRIBUTES", "mail, displayName, department")

        from config import load_config
        _, ldap, _ = load_config()
        assert ldap.attributes == ["mail", "displayName", "department"]
