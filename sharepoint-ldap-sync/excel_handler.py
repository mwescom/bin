"""
Excel workbook helper.

Responsibilities
----------------
* Load a workbook from a local file path (or from an in-memory ``bytes``
  buffer returned by the SharePoint download).
* Iterate over the username column and yield (row_number, username) pairs.
* Write a dictionary of LDAP attributes back into a given row, creating
  header columns as needed.
* Save the workbook back to bytes so the SharePoint client can upload it.
"""

from __future__ import annotations

import io
import logging
from typing import Dict, Generator, List, Optional, Tuple

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import ExcelConfig

logger = logging.getLogger(__name__)


class ExcelHandler:
    """Read and write the user-info spreadsheet."""

    def __init__(self, cfg: ExcelConfig) -> None:
        self.cfg = cfg
        self._wb: Optional[Workbook] = None
        self._ws: Optional[Worksheet] = None
        # Map attribute name -> column letter (built lazily / on first write)
        self._attr_columns: Dict[str, str] = {}

    # ------------------------------------------------------------------
    # Load / save
    # ------------------------------------------------------------------

    def load_from_bytes(self, data: bytes) -> None:
        """Load the workbook from raw bytes (e.g. downloaded from SharePoint)."""
        buf = io.BytesIO(data)
        self._wb = openpyxl.load_workbook(buf)
        self._ws = self._wb.active
        logger.info("Workbook loaded; active sheet: %s", self._ws.title)
        self._build_attr_column_map()

    def load_from_file(self, path: str) -> None:
        """Load the workbook from a local file (useful for testing / dev)."""
        self._wb = openpyxl.load_workbook(path)
        self._ws = self._wb.active
        logger.info("Workbook loaded from %s; active sheet: %s", path, self._ws.title)
        self._build_attr_column_map()

    def to_bytes(self) -> bytes:
        """Serialise the (potentially modified) workbook to bytes."""
        if self._wb is None:
            raise RuntimeError("No workbook loaded.")
        buf = io.BytesIO()
        self._wb.save(buf)
        return buf.getvalue()

    def save_to_file(self, path: str) -> None:
        """Save the workbook to a local file."""
        if self._wb is None:
            raise RuntimeError("No workbook loaded.")
        self._wb.save(path)
        logger.info("Workbook saved to %s", path)

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------

    def iter_usernames(self) -> Generator[Tuple[int, str], None, None]:
        """
        Yield ``(row_number, username)`` for every non-empty cell in the
        username column, starting from ``data_start_row``.
        """
        if self._ws is None:
            raise RuntimeError("No workbook loaded.")

        col_letter = self.cfg.username_column.upper()
        for row in range(self.cfg.data_start_row, self._ws.max_row + 1):
            cell = self._ws[f"{col_letter}{row}"]
            value = cell.value
            if value is not None and str(value).strip():
                yield row, str(value).strip()

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------

    def write_user_data(
        self,
        row: int,
        attributes: Dict[str, Optional[str]],
    ) -> None:
        """
        Write LDAP attribute values into *row*.

        Any attribute that does not yet have a header column is appended to
        the right of the existing columns.
        """
        if self._ws is None:
            raise RuntimeError("No workbook loaded.")

        for attr, value in attributes.items():
            col_letter = self._ensure_column(attr)
            self._ws[f"{col_letter}{row}"] = value

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _build_attr_column_map(self) -> None:
        """
        Scan the header row and map existing column headers to their
        column letters so we can reuse them on subsequent runs.
        """
        if self._ws is None:
            return

        self._attr_columns.clear()
        header_row = self.cfg.header_row

        for cell in self._ws[header_row]:
            if cell.value and cell.column_letter != self.cfg.username_column.upper():
                self._attr_columns[str(cell.value).strip()] = cell.column_letter

        logger.debug("Attribute column map: %s", self._attr_columns)

    def _ensure_column(self, attr_name: str) -> str:
        """
        Return the column letter for *attr_name*, creating a new header
        cell if the column does not exist yet.
        """
        if attr_name in self._attr_columns:
            return self._attr_columns[attr_name]

        # Find next available column
        ws = self._ws
        max_col = ws.max_column or 1

        # Make sure we don't overwrite the username column
        col_idx = max_col + 1
        from openpyxl.utils import get_column_letter

        col_letter = get_column_letter(col_idx)

        # Write header
        ws[f"{col_letter}{self.cfg.header_row}"] = attr_name
        self._attr_columns[attr_name] = col_letter
        logger.debug("Created column %s for attribute '%s'", col_letter, attr_name)
        return col_letter

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------

    def get_usernames(self) -> List[Tuple[int, str]]:
        """Return all (row, username) pairs as a list."""
        return list(self.iter_usernames())
