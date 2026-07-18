"""Shimo platform connector placeholder.

The concrete implementation should own only Shimo-specific listing and fetching
logic. Login, browser automation, and format conversion stay outside this
module so future platforms can reuse them.
"""

from __future__ import annotations

from personal_kb_backup.core.contracts import DocumentContent, DocumentRef


class ShimoConnector:
    """Adapter for 石墨 documents."""

    name = "shimo"

    def list_documents(self) -> list[DocumentRef]:
        """List Shimo documents configured for backup."""

        raise NotImplementedError("Connect Shimo document listing in the next implementation phase.")

    def fetch_document(self, ref: DocumentRef) -> DocumentContent:
        """Fetch a Shimo document and normalize it to Markdown."""

        raise NotImplementedError("Connect Shimo document fetching in the next implementation phase.")
