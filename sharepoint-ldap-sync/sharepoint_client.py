"""
SharePoint Online client.

Responsibilities
----------------
* Authenticate to SharePoint Online using app-only credentials
  (client ID + client secret via the Office 365 REST Python Client library).
* Download the target Excel file as raw bytes.
* Upload (overwrite) the file with updated bytes.
"""

from __future__ import annotations

import logging
import os
from typing import Optional

from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

from config import SharePointConfig

logger = logging.getLogger(__name__)


class SharePointClient:
    """Wrapper around the Office 365 REST Python Client for file operations."""

    def __init__(self, cfg: SharePointConfig) -> None:
        self.cfg = cfg
        self._ctx: Optional[ClientContext] = None

    # ------------------------------------------------------------------
    # Authentication
    # ------------------------------------------------------------------

    def connect(self) -> None:
        """Authenticate and initialise the client context."""
        credentials = ClientCredential(self.cfg.client_id, self.cfg.client_secret)
        self._ctx = ClientContext(self.cfg.site_url).with_credentials(credentials)
        # Validate credentials by loading the web title
        web = self._ctx.web
        self._ctx.load(web)
        self._ctx.execute_query()
        logger.info("Connected to SharePoint site: %s", web.properties.get("Title"))

    def __enter__(self) -> "SharePointClient":
        self.connect()
        return self

    def __exit__(self, *_: object) -> None:
        # No persistent connection to close for the REST client.
        pass

    # ------------------------------------------------------------------
    # File operations
    # ------------------------------------------------------------------

    def download_file(self) -> bytes:
        """
        Download the Excel file configured in ``cfg.file_path`` and return
        its content as raw bytes.
        """
        if self._ctx is None:
            raise RuntimeError("Not connected.  Call connect() or use as context manager.")

        server_relative_url = self._build_server_relative_url()
        logger.info("Downloading file: %s", server_relative_url)

        file_content: list[bytes] = []

        def _on_chunk(chunk: bytes) -> None:
            file_content.append(chunk)

        (
            self._ctx.web
            .get_file_by_server_relative_url(server_relative_url)
            .download(file_content)
            .execute_query()
        )

        data = b"".join(file_content)
        logger.info("Downloaded %d bytes.", len(data))
        return data

    def upload_file(self, data: bytes) -> None:
        """
        Overwrite the Excel file in SharePoint with *data*.

        The file must already exist; this performs an in-place update.
        """
        if self._ctx is None:
            raise RuntimeError("Not connected.  Call connect() or use as context manager.")

        folder_url, filename = self._split_folder_and_filename()
        logger.info("Uploading %d bytes to %s / %s", len(data), folder_url, filename)

        folder = self._ctx.web.get_folder_by_server_relative_url(folder_url)
        folder.upload_file(filename, data).execute_query()

        logger.info("Upload complete.")

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _build_server_relative_url(self) -> str:
        """
        Combine the site URL path with the document-library-relative file
        path to produce a server-relative URL.

        Example:
            site_url  = https://contoso.sharepoint.com/sites/MySite
            file_path = Shared Documents/Users/user_list.xlsx
            result    = /sites/MySite/Shared Documents/Users/user_list.xlsx
        """
        from urllib.parse import urlparse

        site_path = urlparse(self.cfg.site_url).path.rstrip("/")
        file_path = self.cfg.file_path.lstrip("/")
        return f"{site_path}/{file_path}"

    def _split_folder_and_filename(self) -> tuple[str, str]:
        """Return ``(folder_server_relative_url, filename)``."""
        server_relative = self._build_server_relative_url()
        folder_url, _, filename = server_relative.rpartition("/")
        return folder_url, filename
