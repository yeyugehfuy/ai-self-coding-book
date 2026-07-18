"""Export a readable Shimo document page to a Word file.

This script intentionally avoids waiting for legacy Shimo toolbar buttons. Newer
Shimo pages can render readable article content while old fixed controls never
appear, so the export flow now waits for document body candidates first and
writes debug artifacts when the page cannot be detected.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from playwright.sync_api import Page
else:
    Page = Any

try:
    from docx import Document
except ImportError:  # pragma: no cover - handled at runtime for friendly CLI output
    Document = None  # type: ignore[assignment]


class ShimoPageTimeoutError(RuntimeError):
    """Raised when a readable Shimo body cannot be found."""


BODY_SELECTORS = [
    "[data-testid*='editor']",
    "[data-test*='editor']",
    "[class*='editor']",
    "[class*='Editor']",
    "[class*='reader']",
    "[class*='Reader']",
    "[class*='article']",
    "[class*='Article']",
    "[class*='document']",
    "[class*='Document']",
    "[contenteditable='true']",
    "main",
    "article",
    "body",
]


@dataclass(frozen=True)
class BodyCandidate:
    selector: str
    text: str


@dataclass(frozen=True)
class ExportMetadata:
    url: str
    title: str
    selector: str
    output_path: str
    content_hash: str
    exported_at: str
    character_count: int
    skipped: bool = False


def debug_log(enabled: bool, message: str) -> None:
    if enabled:
        print(f"[debug] {message}")


def safe_filename(value: str, fallback: str = "shimo_export") -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", value).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:80] or fallback


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def content_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def load_backup_state(path: Path) -> dict[str, object]:
    if not path.exists():
        return {"version": 1, "documents": {}}
    return json.loads(path.read_text(encoding="utf-8"))


def save_backup_state(path: Path, state: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def update_backup_state(path: Path, metadata: ExportMetadata) -> None:
    state = load_backup_state(path)
    documents = state.setdefault("documents", {})
    if not isinstance(documents, dict):
        documents = {}
        state["documents"] = documents
    documents[metadata.url] = asdict(metadata)
    state["updated_at"] = metadata.exported_at
    save_backup_state(path, state)


def write_report(path: Path, metadata: ExportMetadata) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    status = "skipped" if metadata.skipped else "exported"
    lines = [
        "Personal Knowledge Hub backup report",
        f"status: {status}",
        f"time: {metadata.exported_at}",
        f"url: {metadata.url}",
        f"title: {metadata.title}",
        f"selector: {metadata.selector}",
        f"output: {metadata.output_path}",
        f"characters: {metadata.character_count}",
        f"content_hash: {metadata.content_hash}",
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def should_skip_export(backup_json: Path, url: str, digest: str, force: bool) -> bool:
    if force or not backup_json.exists():
        return False
    state = load_backup_state(backup_json)
    documents = state.get("documents", {})
    if not isinstance(documents, dict):
        return False
    previous = documents.get(url, {})
    return isinstance(previous, dict) and previous.get("content_hash") == digest


def save_timeout_debug(page: Page, debug_dir: Path) -> None:
    debug_dir.mkdir(parents=True, exist_ok=True)
    screenshot_path = debug_dir / "timeout.png"
    html_path = debug_dir / "timeout.html"
    page.screenshot(path=str(screenshot_path), full_page=True)
    html_path.write_text(page.content(), encoding="utf-8")
    print(f"[timeout] screenshot saved: {screenshot_path}")
    print(f"[timeout] html saved: {html_path}")


def selector_has_visible_text(page: Page, selector: str, min_chars: int) -> bool:
    try:
        return bool(
            page.locator(selector).evaluate_all(
                """
                (nodes, minChars) => nodes.some((node) => {
                    const style = window.getComputedStyle(node);
                    const rect = node.getBoundingClientRect();
                    const text = (node.innerText || node.textContent || '').trim();
                    return style.visibility !== 'hidden'
                        && style.display !== 'none'
                        && rect.width > 0
                        && rect.height > 0
                        && text.length >= minChars;
                })
                """,
                min_chars,
            )
        )
    except Exception:
        return False


def find_body_candidate(page: Page, debug: bool, min_chars: int) -> BodyCandidate | None:
    for selector in BODY_SELECTORS:
        debug_log(debug, f"checking selector: {selector}")
        found = selector_has_visible_text(page, selector, min_chars)
        debug_log(debug, f"selector={selector!r} found_visible_body={found}")
        if not found:
            continue
        texts = page.locator(selector).evaluate_all(
            """
            (nodes) => nodes
                .map((node) => (node.innerText || node.textContent || '').trim())
                .filter(Boolean)
                .sort((a, b) => b.length - a.length)
            """
        )
        if texts:
            return BodyCandidate(selector=selector, text=texts[0])
    return None


def wait_for_body(page: Page, debug: bool, timeout_ms: int, min_chars: int) -> BodyCandidate:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

    debug_log(debug, f"current url: {page.url}")
    debug_log(debug, f"page title: {page.title()}")

    immediate_candidate = find_body_candidate(page, debug, min_chars)
    if immediate_candidate:
        debug_log(debug, f"DOM found body node: true, selector={immediate_candidate.selector}")
        return immediate_candidate

    try:
        debug_log(debug, "waiting load state: networkidle")
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except PlaywrightTimeoutError:
        # New Shimo pages may keep background connections open. If readable
        # content is already visible, continue with DOM detection instead of
        # failing on the load-state wait.
        debug_log(debug, "networkidle timeout; continue checking visible body selectors")

    debug_log(debug, f"current url: {page.url}")
    debug_log(debug, f"page title: {page.title()}")

    deadline_selector = BODY_SELECTORS[-1]
    for selector in BODY_SELECTORS:
        debug_log(debug, f"waiting selector: {selector}")
        try:
            page.wait_for_function(
                """
                ({ selector, minChars }) => Array.from(document.querySelectorAll(selector)).some((node) => {
                    const style = window.getComputedStyle(node);
                    const rect = node.getBoundingClientRect();
                    const text = (node.innerText || node.textContent || '').trim();
                    return style.visibility !== 'hidden'
                        && style.display !== 'none'
                        && rect.width > 0
                        && rect.height > 0
                        && text.length >= minChars;
                })
                """,
                {"selector": selector, "minChars": min_chars},
                timeout=3000 if selector != deadline_selector else 5000,
            )
        except PlaywrightTimeoutError:
            debug_log(debug, f"selector timeout: {selector}")
            continue

        candidate = find_body_candidate(page, debug, min_chars)
        if candidate:
            debug_log(debug, f"DOM found body node: true, selector={candidate.selector}")
            return candidate

    candidate = find_body_candidate(page, debug, min_chars)
    if candidate:
        return candidate
    raise ShimoPageTimeoutError(f"No readable Shimo body found within {timeout_ms}ms")


def write_word(title: str, text: str, output_path: Path) -> None:
    if Document is None:
        raise RuntimeError("Missing dependency: install python-docx to write .docx files.")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    document.add_heading(title or "石墨导出", level=1)
    for block in re.split(r"\n{2,}", text):
        block = block.strip()
        if block:
            document.add_paragraph(block)
    document.save(output_path)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export a Shimo document to Word.")
    parser.add_argument("--url", required=True, help="Shimo document URL")
    parser.add_argument("--output", type=Path, default=None, help="Output .docx path")
    parser.add_argument("--backup-json", type=Path, default=Path("backup.json"), help="Incremental backup state file")
    parser.add_argument("--report", type=Path, default=Path("backup-report.txt"), help="Human-readable backup report path")
    parser.add_argument("--profile-dir", type=Path, default=Path(".shimo-browser-profile"), help="Persistent browser profile for Shimo login state")
    parser.add_argument("--force", action="store_true", help="Export even when backup.json says content is unchanged")
    parser.add_argument("--debug", action="store_true", help="Print selector, URL, title, and DOM diagnostics")
    parser.add_argument("--debug-dir", type=Path, default=Path("debug"), help="Directory for timeout.png and timeout.html")
    parser.add_argument("--timeout", type=int, default=45000, help="Overall page/body wait timeout in milliseconds")
    parser.add_argument("--min-chars", type=int, default=20, help="Minimum visible text length considered as document body")
    parser.add_argument("--headed", action="store_true", help="Run Chromium with a visible window")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright
    except ImportError as exc:
        print("[error] Missing dependency: install playwright and run `playwright install chromium`.")
        print(f"[error] {exc}")
        return 1

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(args.profile_dir),
            headless=not args.headed,
        )
        page = context.new_page()
        try:
            page.goto(args.url, wait_until="domcontentloaded", timeout=args.timeout)
            candidate = wait_for_body(page, args.debug, args.timeout, args.min_chars)
            title = page.title() or "石墨导出"
            digest = content_hash(candidate.text)
            output_path = args.output or Path("exports") / f"{safe_filename(title)}.docx"

            skipped = should_skip_export(args.backup_json, args.url, digest, args.force)
            if skipped:
                print("[ok] unchanged; export skipped")
            else:
                write_word(title, candidate.text, output_path)
                print(f"[ok] exported selector: {candidate.selector}")
                print(f"[ok] output: {output_path}")

            metadata = ExportMetadata(
                url=args.url,
                title=title,
                selector=candidate.selector,
                output_path=str(output_path),
                content_hash=digest,
                exported_at=now_iso(),
                character_count=len(candidate.text),
                skipped=skipped,
            )
            update_backup_state(args.backup_json, metadata)
            write_report(args.report, metadata)
            print(f"[ok] backup state: {args.backup_json}")
            print(f"[ok] report: {args.report}")
            return 0
        except (PlaywrightTimeoutError, ShimoPageTimeoutError) as exc:
            print(f"[timeout] {exc}")
            print(f"[timeout] current url: {page.url}")
            try:
                print(f"[timeout] page title: {page.title()}")
            except Exception as title_error:
                print(f"[timeout] page title unavailable: {title_error}")
            save_timeout_debug(page, args.debug_dir)
            if args.debug:
                candidate = find_body_candidate(page, True, args.min_chars)
                print(f"[debug] DOM found body node: {candidate is not None}")
            return 2
        finally:
            context.close()


if __name__ == "__main__":
    raise SystemExit(main())
