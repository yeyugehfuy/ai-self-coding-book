"""Export Shimo documents through Shimo's official Word download flow.

The tool intentionally does not parse Shimo page HTML into Word. It simulates the
manual browser operation instead: open document, click the top-right menu, choose
Download, choose Word, wait for the browser download, then rename/move the
official `.docx` into `exports/` using Shimo's suggested filename and update `backup.json`.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:
    from playwright.sync_api import Download, Locator, Page
else:
    Download = Any
    Locator = Any
    Page = Any


MENU_SELECTORS = [
    "button[aria-haspopup='menu']",
    "button[aria-expanded]",
    "[role='button'][aria-haspopup='menu']",
    "[role='button'][aria-expanded]",
    "button[aria-label*='more' i]",
    "button[aria-label*='menu' i]",
    "button[aria-label*='operation' i]",
    "button[aria-label*='操作']",
    "button[aria-label*='菜单']",
    "button[aria-label*='更多']",
    "[data-testid*='more' i]",
    "[data-test*='more' i]",
    "button:has(svg)",
    "[role='button']:has(svg)",
]

DOWNLOAD_SELECTORS = [
    "text=下载",
    "[aria-label*='下载']",
    "[role='menuitem']:has-text('下载')",
    "[role='button']:has-text('下载')",
    "button:has-text('下载')",
]

WORD_SELECTORS = [
    "text=Word",
    "text=.docx",
    "text=DOCX",
    "[role='menuitem']:has-text('Word')",
    "[role='menuitem']:has-text('docx')",
]

TITLE_DATE_RE = re.compile(r"(?P<yy>\d{2})[.年/-](?P<m>\d{1,2})[.月/-](?P<d>\d{1,2})")


@dataclass(frozen=True)
class ExportMetadata:
    title: str
    url: str
    local_filename: str
    local_path: str
    downloaded_at: str
    shimo_updated_at: str | None
    content_hash: str | None
    export_format: str = "word"
    skipped: bool = False


@dataclass
class SyncStats:
    added: int = 0
    updated: int = 0
    skipped: int = 0
    failed: int = 0


def debug_log(enabled: bool, message: str) -> None:
    if enabled:
        print(f"[debug] {message}")


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def safe_filename(value: str, fallback: str = "shimo_export") -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", value).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:120] or fallback


def title_date(title: str) -> date | None:
    match = TITLE_DATE_RE.search(title)
    if not match:
        return None
    year = 2000 + int(match.group("yy"))
    return date(year, int(match.group("m")), int(match.group("d")))


def dated_output_dir(base_dir: Path, title: str) -> Path:
    parsed = title_date(title)
    if parsed is None:
        parsed = datetime.now().date()
    return base_dir / f"{parsed.year:04d}" / f"{parsed.month:02d}"


def default_output_path(base_dir: Path, title: str) -> Path:
    return dated_output_dir(base_dir, title) / f"{safe_filename(title)}.docx"


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_backup_state(path: Path) -> dict[str, object]:
    if not path.exists():
        return {"version": 2, "documents": {}, "last_sync_at": None}
    return json.loads(path.read_text(encoding="utf-8"))


def save_backup_state(path: Path, state: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def backup_documents(state: dict[str, object]) -> dict[str, object]:
    documents = state.setdefault("documents", {})
    if not isinstance(documents, dict):
        documents = {}
        state["documents"] = documents
    return documents


def previous_metadata(backup_json: Path, url: str) -> dict[str, object] | None:
    state = load_backup_state(backup_json)
    previous = backup_documents(state).get(url)
    return previous if isinstance(previous, dict) else None


def update_backup_state(path: Path, metadata: ExportMetadata) -> None:
    state = load_backup_state(path)
    state["version"] = 2
    backup_documents(state)[metadata.url] = asdict(metadata)
    state["last_sync_at"] = metadata.downloaded_at
    save_backup_state(path, state)


def should_skip_export(backup_json: Path, url: str, shimo_updated_at: str | None, force: bool) -> bool:
    if force or not backup_json.exists():
        return False
    previous = previous_metadata(backup_json, url)
    if previous is None:
        return False
    local_path = previous.get("local_path")
    if not isinstance(local_path, str) or not Path(local_path).exists():
        return False
    if shimo_updated_at and previous.get("shimo_updated_at") != shimo_updated_at:
        return False
    return True


def parse_date_input(value: str) -> date:
    value = value.strip()
    match = TITLE_DATE_RE.fullmatch(value)
    if match:
        return date(2000 + int(match.group("yy")), int(match.group("m")), int(match.group("d")))
    return datetime.strptime(value, "%Y-%m-%d").date()


def in_date_range(title: str, start: date | None, end: date | None) -> bool:
    parsed = title_date(title)
    if parsed is None:
        return False
    if start and parsed < start:
        return False
    if end and parsed > end:
        return False
    return True


def save_timeout_debug(page: Page, debug_dir: Path) -> None:
    debug_dir.mkdir(parents=True, exist_ok=True)
    screenshot_path = debug_dir / "timeout.png"
    html_path = debug_dir / "timeout.html"
    page.screenshot(path=str(screenshot_path), full_page=True)
    html_path.write_text(page.content(), encoding="utf-8")
    print(f"[timeout] current url: {page.url}")
    try:
        print(f"[timeout] page title: {page.title()}")
    except Exception as exc:
        print(f"[timeout] page title unavailable: {exc}")
    print(f"[timeout] screenshot saved: {screenshot_path}")
    print(f"[timeout] html saved: {html_path}")



def describe_interactive_elements(page: Page, limit: int = 80) -> list[dict[str, object]]:
    return page.evaluate(
        """
        (limit) => Array.from(document.querySelectorAll('button,[role="button"],[role="menuitem"],a'))
            .map((node, index) => {
                const rect = node.getBoundingClientRect();
                const style = window.getComputedStyle(node);
                const visible = style.visibility !== 'hidden'
                    && style.display !== 'none'
                    && rect.width > 0
                    && rect.height > 0;
                return {
                    index,
                    tag: node.tagName.toLowerCase(),
                    role: node.getAttribute('role') || '',
                    ariaLabel: node.getAttribute('aria-label') || '',
                    title: node.getAttribute('title') || '',
                    text: (node.innerText || node.textContent || '').trim().replace(/\\s+/g, ' ').slice(0, 80),
                    className: typeof node.className === 'string' ? node.className.slice(0, 160) : '',
                    visible,
                    disabled: Boolean(node.disabled) || node.getAttribute('aria-disabled') === 'true',
                    x: Math.round(rect.x),
                    y: Math.round(rect.y),
                    width: Math.round(rect.width),
                    height: Math.round(rect.height),
                };
            })
            .filter((item) => item.visible)
            .sort((a, b) => (a.y - b.y) || (b.x - a.x))
            .slice(0, limit)
        """,
        limit,
    )


def print_interactive_elements(page: Page, title: str = "interactive elements") -> None:
    print(f"[debug] {title}:")
    try:
        for item in describe_interactive_elements(page):
            print(
                "[debug] "
                f"#{item['index']} tag={item['tag']} role={item['role']!r} "
                f"aria={item['ariaLabel']!r} title={item['title']!r} "
                f"text={item['text']!r} class={item['className']!r} "
                f"box=({item['x']},{item['y']},{item['width']},{item['height']}) "
                f"disabled={item['disabled']}"
            )
    except Exception as exc:
        print(f"[debug] cannot dump interactive elements: {exc}")


def visible_locator_candidates(page: Page, selector: str, timeout_ms: int) -> list[Locator]:
    page.locator(selector).first.wait_for(state="attached", timeout=timeout_ms)
    locator = page.locator(selector)
    count = min(locator.count(), 30)
    candidates: list[Locator] = []
    for index in range(count):
        item = locator.nth(index)
        try:
            if item.is_visible() and item.bounding_box() is not None:
                candidates.append(item)
        except Exception:
            continue
    return candidates


def top_right_button(page: Page, debug: bool) -> Locator | None:
    buttons = page.locator("button,[role='button']")
    best: tuple[float, Locator] | None = None
    count = min(buttons.count(), 80)
    viewport = page.viewport_size or {"width": 1280, "height": 720}
    for index in range(count):
        item = buttons.nth(index)
        try:
            box = item.bounding_box()
            if not item.is_visible() or box is None:
                continue
            # Prefer visible controls in the top-right document toolbar.
            if box["y"] > 160 or box["x"] < viewport["width"] * 0.45:
                continue
            text = (item.inner_text(timeout=200) or "").strip()
            aria = item.get_attribute("aria-label") or ""
            score = box["x"] - box["y"]
            # Avoid obvious user/avatar/share controls when possible.
            bad_words = ("分享", "share", "头像", "avatar", "comment", "评论")
            if any(word.lower() in f"{text} {aria}".lower() for word in bad_words):
                score -= 1000
            if best is None or score > best[0]:
                best = (score, item)
        except Exception:
            continue
    if best:
        debug_log(debug, "using top-right visible button fallback for menu")
        return best[1]
    return None

def first_visible(page: Page, selectors: Iterable[str], debug: bool, timeout_ms: int = 2500) -> Locator:
    last_error: Exception | None = None
    for selector in selectors:
        debug_log(debug, f"waiting/click candidate: {selector}")
        try:
            candidates = visible_locator_candidates(page, selector, timeout_ms)
            if candidates:
                debug_log(debug, f"found selector: {selector}, visible_count={len(candidates)}")
                return candidates[0]
        except Exception as exc:
            last_error = exc
            debug_log(debug, f"selector unavailable: {selector} ({exc})")
    raise RuntimeError(f"No visible selector found. Last error: {last_error}")


def click_menu_button(page: Page, debug: bool) -> None:
    try:
        menu = first_visible(page, MENU_SELECTORS, debug, timeout_ms=1200)
    except Exception as exc:
        debug_log(debug, f"stable menu selectors failed: {exc}")
        menu = top_right_button(page, debug)
        if menu is None:
            print_interactive_elements(page, "visible buttons before menu failure")
            raise
    menu.click()


def trigger_official_word_download(page: Page, debug: bool, timeout_ms: int) -> Download:
    if debug:
        print_interactive_elements(page, "visible buttons before opening menu")
    click_menu_button(page, debug)
    debug_log(debug, "clicked menu button")
    if debug:
        print_interactive_elements(page, "visible buttons/menuitems after opening menu")
    download_item = first_visible(page, DOWNLOAD_SELECTORS, debug)
    debug_log(debug, "found download menu")
    download_item.hover()
    debug_log(debug, "hovered download menu")
    if debug:
        print_interactive_elements(page, "visible buttons/menuitems after hovering download")
    word_item = first_visible(page, WORD_SELECTORS, debug)
    debug_log(debug, "found Word menu item")
    with page.expect_download(timeout=timeout_ms) as download_info:
        word_item.click()
    download = download_info.value
    debug_log(debug, f"download suggested filename: {download.suggested_filename}")
    return download


def downloaded_filename(download: Download) -> str:
    suggested = download.suggested_filename or "shimo.docx"
    if not suggested.lower().endswith(".docx"):
        suggested = f"{suggested}.docx"
    return safe_filename(suggested)


def save_download_to_temp(download: Download, temp_dir: Path) -> Path:
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_path = temp_dir / downloaded_filename(download)
    download.save_as(str(temp_path))
    return temp_path


def keep_or_skip_download(temp_path: Path, output_path: Path, backup_json: Path, url: str) -> tuple[Path, str, bool]:
    digest = file_hash(temp_path)
    previous = previous_metadata(backup_json, url)
    if previous and previous.get("content_hash") == digest and output_path.exists():
        temp_path.unlink(missing_ok=True)
        return output_path, digest, True
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()
    temp_path.replace(output_path)
    return output_path, digest, False


def write_report(path: Path, stats: SyncStats, export_dir: Path, last_sync: str, lines: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    report = [
        "Personal Knowledge Hub backup report",
        f"新增：{stats.added}",
        f"更新：{stats.updated}",
        f"跳过：{stats.skipped}",
        f"失败：{stats.failed}",
        f"导出目录：{export_dir}",
        f"最后同步：{last_sync}",
        "",
        "明细：",
        *lines,
    ]
    path.write_text("\n".join(report) + "\n", encoding="utf-8")


def export_one(page: Page, args: argparse.Namespace, url: str, stats: SyncStats, report_lines: list[str]) -> None:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

    try:
        print(f"[sync] open: {url}")
        page.goto(url, wait_until="domcontentloaded", timeout=args.timeout)
        try:
            page.wait_for_load_state("networkidle", timeout=args.timeout)
        except PlaywrightTimeoutError:
            debug_log(args.debug, "networkidle timeout; continue official download flow")

        title = safe_filename(page.title() or "石墨文档")
        if args.debug:
            debug_log(args.debug, f"current page title: {title}")
            debug_log(args.debug, f"current url: {page.url}")
        if args.dump_buttons:
            print_interactive_elements(page, "visible buttons/menuitems")
            stats.skipped += 1
            report_lines.append(f"调试按钮：{title} {url}")
            return
        shimo_updated_at = args.updated_at
        if args.start_date or args.end_date:
            if not in_date_range(title, args.start_date, args.end_date):
                stats.skipped += 1
                report_lines.append(f"跳过（日期范围外）：{title} {url}")
                return

        existed_before = previous_metadata(args.backup_json, url) is not None
        download = trigger_official_word_download(page, args.debug, args.timeout)
        output_path = args.output or (args.output_dir / downloaded_filename(download))
        temp_path = save_download_to_temp(download, args.output_dir / ".tmp")
        final_path, digest, unchanged = keep_or_skip_download(temp_path, output_path, args.backup_json, url)
        metadata = ExportMetadata(
            title=title,
            url=url,
            local_filename=final_path.name,
            local_path=str(final_path),
            downloaded_at=now_iso(),
            shimo_updated_at=shimo_updated_at,
            content_hash=digest,
        )
        update_backup_state(args.backup_json, metadata)
        if unchanged:
            stats.skipped += 1
            report_lines.append(f"跳过（Hash 未变化）：{title} -> {final_path}")
            print(f"[skip] downloaded Word hash unchanged: {final_path}")
        elif existed_before:
            stats.updated += 1
            report_lines.append(f"更新：{title} -> {final_path}")
            print(f"[ok] official Word updated: {final_path}")
        else:
            stats.added += 1
            report_lines.append(f"新增：{title} -> {final_path}")
            print(f"[ok] official Word downloaded: {final_path}")
        debug_log(args.debug, f"download path: {final_path}")
        debug_log(args.debug, f"download success: {final_path.exists()}")
    except Exception as exc:
        stats.failed += 1
        report_lines.append(f"失败：{url} {exc}")
        print(f"[error] {url}: {exc}")
        save_timeout_debug(page, args.debug_dir)


def urls_from_args(args: argparse.Namespace) -> list[str]:
    urls = list(args.urls or [])
    if args.url:
        urls.append(args.url)
    if args.url_file:
        urls.extend(line.strip() for line in args.url_file.read_text(encoding="utf-8").splitlines() if line.strip())
    seen: set[str] = set()
    unique: list[str] = []
    for url in urls:
        if url not in seen:
            seen.add(url)
            unique.append(url)
    return unique


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Sync Shimo documents through official Word downloads.")
    parser.add_argument("--url", help="Single Shimo document URL")
    parser.add_argument("--urls", nargs="*", help="One or more Shimo document URLs")
    parser.add_argument("--url-file", type=Path, help="Text file containing one Shimo URL per line")
    parser.add_argument("--output", type=Path, default=None, help="Single-document output .docx path")
    parser.add_argument("--output-dir", type=Path, default=Path("exports"), help="Base export directory; official Shimo filenames are kept here")
    parser.add_argument("--backup-json", type=Path, default=Path("backup.json"), help="Sync database path")
    parser.add_argument("--report", type=Path, default=Path("backup-report.txt"), help="Human-readable sync report path")
    parser.add_argument("--profile-dir", type=Path, default=Path(".shimo-browser-profile"), help="Persistent browser profile for Shimo login state")
    parser.add_argument("--mode", choices=["all", "new", "7days", "30days", "range", "single", "force"], default="new")
    parser.add_argument("--start", help="Start date for range mode, e.g. 26.7.01 or 2026-07-01")
    parser.add_argument("--end", help="End date for range mode, e.g. 26.7.15 or 2026-07-15")
    parser.add_argument("--updated-at", help="Known Shimo last modified time for a single document")
    parser.add_argument("--force", action="store_true", help="Export even when backup.json says content is unchanged")
    parser.add_argument("--debug", action="store_true", help="Print selector, URL, title, and DOM diagnostics")
    parser.add_argument("--dump-buttons", action="store_true", help="Print visible buttons/menuitems and exit after page load")
    parser.add_argument("--debug-dir", type=Path, default=Path("debug"), help="Directory for timeout.png and timeout.html")
    parser.add_argument("--timeout", type=int, default=45000, help="Page/download timeout in milliseconds")
    parser.add_argument("--headed", action="store_true", help="Run Chromium with a visible window")
    args = parser.parse_args(argv)
    args.start_date = parse_date_input(args.start) if args.start else None
    args.end_date = parse_date_input(args.end) if args.end else None
    if args.mode == "7days":
        args.start_date = datetime.now().date() - timedelta(days=7)
    elif args.mode == "30days":
        args.start_date = datetime.now().date() - timedelta(days=30)
    elif args.mode in {"all", "force"}:
        args.force = True
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    urls = urls_from_args(args)
    if not urls:
        print("[error] No Shimo URL provided. Use --url, --urls, or --url-file.")
        return 1
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        print("[error] Missing dependency: install playwright and run `playwright install chromium`.")
        print(f"[error] {exc}")
        return 1

    stats = SyncStats()
    report_lines: list[str] = []
    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(args.profile_dir),
            headless=not args.headed,
            accept_downloads=True,
        )
        page = context.new_page()
        try:
            for url in urls:
                export_one(page, args, url, stats, report_lines)
        finally:
            context.close()

    last_sync = datetime.now().strftime("%Y-%m-%d %H:%M")
    write_report(args.report, stats, args.output_dir, last_sync, report_lines)
    print("[summary]")
    print(f"新增：{stats.added}")
    print(f"更新：{stats.updated}")
    print(f"跳过：{stats.skipped}")
    print(f"失败：{stats.failed}")
    print(f"导出目录：{args.output_dir}")
    print(f"最后同步：{last_sync}")
    return 0 if stats.failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
