"""
Unit tests for the SharePointClient class.

All Office365 REST calls are mocked so no real SharePoint tenant is needed.
"""

from __future__ import annotations

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unittest.mock import MagicMock, patch, PropertyMock
import pytest

from config import SharePointConfig
from sharepoint_client import SharePointClient


DEFAULT_SP_CFG = SharePointConfig(
    site_url="https://contoso.sharepoint.com/sites/MySite",
    client_id="client-id",
    client_secret="client-secret",
    file_path="Shared Documents/user_list.xlsx",
)


class TestBuildServerRelativeUrl:
    def test_standard_path(self):
        client = SharePointClient(DEFAULT_SP_CFG)
        url = client._build_server_relative_url()
        assert url == "/sites/MySite/Shared Documents/user_list.xlsx"

    def test_file_path_with_leading_slash(self):
        cfg = SharePointConfig(
            site_url="https://contoso.sharepoint.com/sites/MySite",
            client_id="x",
            client_secret="y",
            file_path="/Shared Documents/user_list.xlsx",
        )
        client = SharePointClient(cfg)
        url = client._build_server_relative_url()
        assert url == "/sites/MySite/Shared Documents/user_list.xlsx"

    def test_root_site(self):
        cfg = SharePointConfig(
            site_url="https://contoso.sharepoint.com",
            client_id="x",
            client_secret="y",
            file_path="Documents/file.xlsx",
        )
        client = SharePointClient(cfg)
        url = client._build_server_relative_url()
        assert url == "/Documents/file.xlsx"


class TestSplitFolderAndFilename:
    def test_standard(self):
        client = SharePointClient(DEFAULT_SP_CFG)
        folder, filename = client._split_folder_and_filename()
        assert filename == "user_list.xlsx"
        assert folder == "/sites/MySite/Shared Documents"


class TestConnectNotConnectedErrors:
    def test_download_raises_when_not_connected(self):
        client = SharePointClient(DEFAULT_SP_CFG)
        with pytest.raises(RuntimeError):
            client.download_file()

    def test_upload_raises_when_not_connected(self):
        client = SharePointClient(DEFAULT_SP_CFG)
        with pytest.raises(RuntimeError):
            client.upload_file(b"data")


class TestDownloadFile:
    def test_download_returns_bytes(self):
        client = SharePointClient(DEFAULT_SP_CFG)

        mock_ctx = MagicMock()
        file_content_holder = []

        def _fake_download(container):
            container.append(b"PK fake xlsx content")
            return MagicMock(execute_query=lambda: None)

        mock_file = MagicMock()
        mock_file.download.side_effect = _fake_download
        mock_ctx.web.get_file_by_server_relative_url.return_value = mock_file

        client._ctx = mock_ctx
        result = client.download_file()
        assert isinstance(result, bytes)

    def test_download_concatenates_chunks(self):
        client = SharePointClient(DEFAULT_SP_CFG)

        mock_ctx = MagicMock()

        def _fake_download(container):
            container.append(b"chunk1")
            container.append(b"chunk2")
            return MagicMock(execute_query=lambda: None)

        mock_file = MagicMock()
        mock_file.download.side_effect = _fake_download
        mock_ctx.web.get_file_by_server_relative_url.return_value = mock_file

        client._ctx = mock_ctx
        result = client.download_file()
        assert result == b"chunk1chunk2"


class TestUploadFile:
    def test_upload_calls_folder_upload(self):
        client = SharePointClient(DEFAULT_SP_CFG)

        mock_ctx = MagicMock()
        mock_folder = MagicMock()
        upload_result = MagicMock()
        mock_folder.upload_file.return_value = upload_result

        mock_ctx.web.get_folder_by_server_relative_url.return_value = mock_folder
        client._ctx = mock_ctx

        client.upload_file(b"updated xlsx data")

        mock_ctx.web.get_folder_by_server_relative_url.assert_called_once_with(
            "/sites/MySite/Shared Documents"
        )
        mock_folder.upload_file.assert_called_once_with("user_list.xlsx", b"updated xlsx data")
        upload_result.execute_query.assert_called_once()
