"""Composable backup pipeline orchestration."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .contracts import AuthProvider, ExportResult, FormatExporter, PlatformConnector


@dataclass
class BackupPipeline:
    """Run login, platform sync, incremental checks, export, and reporting."""

    auth_providers: list[AuthProvider]
    platforms: list[PlatformConnector]
    exporters: list[FormatExporter]
    output_dir: Path

    def run(self) -> list[ExportResult]:
        """Run a full backup pass and return exported files.

        Incremental change detection is intentionally kept as a replaceable
        concern for the next phase; current scaffolding exports every fetched
        document through every configured exporter.
        """

        for provider in self.auth_providers:
            provider.ensure_login()

        results: list[ExportResult] = []
        for platform in self.platforms:
            for ref in platform.list_documents():
                document = platform.fetch_document(ref)
                for exporter in self.exporters:
                    results.append(exporter.export(document, self.output_dir))
        return results
