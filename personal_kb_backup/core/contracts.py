"""Core extension contracts for the knowledge base backup pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol, Sequence


@dataclass(frozen=True)
class DocumentRef:
    """A lightweight reference to a source document on a platform."""

    platform: str
    document_id: str
    title: str
    updated_at: str | None = None
    path: str | None = None


@dataclass(frozen=True)
class DocumentContent:
    """Normalized document content used between platform and format plugins."""

    ref: DocumentRef
    markdown: str
    assets: Sequence[Path] = field(default_factory=tuple)
    metadata: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class ExportResult:
    """Output produced by a format exporter."""

    ref: DocumentRef
    format_name: str
    output_path: Path


class AuthProvider(Protocol):
    """Authentication boundary for platforms that need login sessions."""

    name: str

    def ensure_login(self) -> None:
        """Create or refresh a usable login/session state."""


class PlatformConnector(Protocol):
    """Connector boundary for knowledge platforms such as Shimo or Notion."""

    name: str

    def list_documents(self) -> Sequence[DocumentRef]:
        """Return documents visible to the configured account/workspace."""

    def fetch_document(self, ref: DocumentRef) -> DocumentContent:
        """Fetch one document and normalize it to Markdown plus assets."""


class FormatExporter(Protocol):
    """Exporter boundary for output formats such as Word, PDF, HTML, or Obsidian."""

    name: str

    def export(self, document: DocumentContent, output_dir: Path) -> ExportResult:
        """Write one document to the requested format."""
