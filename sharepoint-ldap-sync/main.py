"""
Main entry point for the SharePoint-LDAP sync tool.

Workflow
--------
1. Load configuration from environment / .env file.
2. Connect to SharePoint and download the Excel workbook.
3. Parse the workbook to extract usernames.
4. Connect to LDAP and query each username.
5. Write LDAP attributes back into the workbook.
6. Upload the updated workbook back to SharePoint.
"""

from __future__ import annotations

import logging
import sys
from typing import Optional

from config import load_config
from excel_handler import ExcelHandler
from ldap_client import LDAPClient
from sharepoint_client import SharePointClient

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def run() -> int:
    """
    Execute the full sync workflow.

    Returns
    -------
    int
        Exit code: 0 on success, 1 on fatal error.
    """

    # 1. Configuration -------------------------------------------------------
    try:
        sp_cfg, ldap_cfg, excel_cfg = load_config()
    except EnvironmentError as exc:
        logger.error("Configuration error: %s", exc)
        return 1

    # 2. Download workbook from SharePoint ------------------------------------
    logger.info("Step 1/4 – Downloading workbook from SharePoint …")
    try:
        with SharePointClient(sp_cfg) as sp_client:
            workbook_bytes = sp_client.download_file()
    except Exception as exc:  # noqa: BLE001
        logger.error("Failed to download workbook: %s", exc)
        return 1

    # 3. Parse usernames ------------------------------------------------------
    logger.info("Step 2/4 – Parsing usernames from workbook …")
    handler = ExcelHandler(excel_cfg)
    try:
        handler.load_from_bytes(workbook_bytes)
    except Exception as exc:  # noqa: BLE001
        logger.error("Failed to load workbook: %s", exc)
        return 1

    user_rows = handler.get_usernames()
    if not user_rows:
        logger.warning("No usernames found in the workbook – nothing to do.")
        return 0

    logger.info("Found %d username(s) to process.", len(user_rows))

    # 4. LDAP queries & workbook update ---------------------------------------
    logger.info("Step 3/4 – Querying LDAP and updating workbook …")
    updated = 0
    not_found = 0

    try:
        with LDAPClient(ldap_cfg) as ldap_client:
            for row_num, username in user_rows:
                logger.info("  Processing user: %s (row %d)", username, row_num)
                attrs = ldap_client.query_user(username)
                if attrs is None:
                    not_found += 1
                    # Write a sentinel so the spreadsheet shows the status
                    handler.write_user_data(row_num, {"Status": "NOT FOUND"})
                else:
                    attrs["Status"] = "OK"
                    handler.write_user_data(row_num, attrs)
                    updated += 1

    except Exception as exc:  # noqa: BLE001
        logger.error("LDAP processing failed: %s", exc)
        return 1

    logger.info(
        "Updated: %d  |  Not found: %d  |  Total: %d",
        updated,
        not_found,
        len(user_rows),
    )

    # 5. Upload workbook back to SharePoint -----------------------------------
    logger.info("Step 4/4 – Uploading updated workbook to SharePoint …")
    try:
        updated_bytes = handler.to_bytes()
        with SharePointClient(sp_cfg) as sp_client:
            sp_client.upload_file(updated_bytes)
    except Exception as exc:  # noqa: BLE001
        logger.error("Failed to upload workbook: %s", exc)
        return 1

    logger.info("Sync complete.")
    return 0


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    sys.exit(run())
