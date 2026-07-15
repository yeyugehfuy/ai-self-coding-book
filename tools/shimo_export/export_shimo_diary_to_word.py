#!/usr/bin/env python3
"""Export Shimo diary documents to Word/PDF files with Playwright.

This helper keeps account credentials out of the repository. Prefer using an
existing browser session created with --login-only. If password login is needed,
provide credentials through environment variables instead of command-line flags:
SHIMO_USERNAME and SHIMO_PASSWORD.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterable, Literal
from urllib.parse import urljoin, urlparse, urlunparse

if TYPE_CHECKING:
    from playwright.sync_api import BrowserContext, Page
else:
    BrowserContext = Any
    Page = Any

ExportFormat = Literal["docx", "pdf"]


class MissingPlaywrightError(RuntimeError):
    pass


@dataclass(frozen=True)
class ShimoItem:
    title: str
    url: str
    last_modified: str = ""


@dataclass(frozen=True)
class ExportFailure:
    title: str
    url: str
    reason: str


@dataclass
class BackupRecord:
    document_id: str
    title: str
    last_modified: str
    last_exported: str
    export_paths: dict[str, str] = field(default_factory=dict)


def load_playwright():
    try:
        from playwright.sync_api import TimeoutError, sync_playwright
    except ModuleNotFoundError as exc:
        raise MissingPlaywrightError(
            "Playwright is not installed. Run: pip install playwright && "
            "python -m playwright install chromium"
        ) from exc
    return TimeoutError, sync_playwright


SHIMO_HOME = "https://shimo.im/"
DOC_URL_PATTERN = re.compile(
    r"/(docs|doc|document|lizard|spreadsheet|sheet|sheets|presentation|slides)/",
    re.I,
)
FOLDER_URL_PATTERN = re.compile(r"/(folder|desktop/folder)/", re.I)
LOGIN_TEXT_PATTERN = re.compile(r"登录|login|sign in", re.I)
EXPORT_PATTERNS: dict[ExportFormat, list[re.Pattern[str]]] = {
    "docx": [
        re.compile(r"导出.*(Word|word|DOCX|docx|文档)"),
        re.compile(r"(Word|word|DOCX|docx|文档).*导出"),
        re.compile(r"下载.*(Word|word|DOCX|docx|文档)"),
        re.compile(r"Word|DOCX", re.I),
    ],
    "pdf": [
        re.compile(r"导出.*(PDF|pdf)"),
        re.compile(r"(PDF|pdf).*导出"),
        re.compile(r"下载.*(PDF|pdf)"),
        re.compile(r"PDF", re.I),
    ],
}
MENU_TEXTS = ["更多", "···", "...", "菜单", "文件", "导出", "下载"]
NEXT_PAGE_TEXTS = ["下一页", "Next", "下页", ">"]
MODIFIED_TIME_PATTERN = re.compile(
    r"(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}(?:日)?(?:\s+\d{1,2}:\d{2})?|"
    r"\d{1,2}[-/.月]\d{1,2}(?:日)?(?:\s+\d{1,2}:\d{2})?|"
    r"昨天\s*\d{0,2}:?\d{0,2}|今天\s*\d{0,2}:?\d{0,2}|"
    r"\d+\s*(?:分钟前|小时前|天前))"
)
SUPPORTED_SUFFIXES: dict[ExportFormat, tuple[str, ...]] = {
    "docx": (".doc", ".docx"),
    "pdf": (".pdf",),
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download Shimo diary article/document URLs as Word or PDF files."
    )
    parser.add_argument(
        "--folder",
        action="append",
        default=[],
        help="A Shimo folder URL to crawl recursively. Repeat for multiple folders.",
    )
    parser.add_argument(
        "--urls-file",
        type=Path,
        help="Text file containing one Shimo document URL per line.",
    )
    parser.add_argument(
        "--url",
        action="append",
        default=[],
        help="A Shimo document URL. Repeat this flag for multiple documents.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("exported-shimo-diaries"),
        help="Directory used for downloaded files.",
    )
    parser.add_argument(
        "--profile-dir",
        type=Path,
        default=Path(".shimo-browser-profile"),
        help="Persistent browser profile directory. Keep it out of Git.",
    )
    parser.add_argument(
        "--login-only",
        action="store_true",
        help="Open Shimo and stop after you finish logging in manually.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser in headless mode. Manual login requires headed mode.",
    )
    parser.add_argument(
        "--timeout-ms",
        type=int,
        default=45_000,
        help="Timeout for page actions and downloads.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing exported files instead of skipping them.",
    )
    parser.add_argument(
        "--format",
        choices=("docx", "pdf", "both"),
        default="docx",
        help="Export format: docx, pdf, or both. Defaults to docx.",
    )
    parser.add_argument(
        "--max-scrolls",
        type=int,
        default=80,
        help="Maximum scroll attempts per folder for infinite-loading folders.",
    )
    parser.add_argument(
        "--sync",
        action="store_true",
        default=True,
        help="Use backup.json metadata to export only new or changed documents. Enabled by default.",
    )
    parser.add_argument(
        "--backup-db",
        type=Path,
        help="Path to backup metadata JSON. Defaults to <output-dir>/backup.json.",
    )
    return parser.parse_args()


def requested_formats(args: argparse.Namespace) -> list[ExportFormat]:
    return ["docx", "pdf"] if args.format == "both" else [args.format]


def normalize_url(url: str) -> str:
    parsed = urlparse(urljoin(SHIMO_HOME, url))
    return urlunparse((parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), "", "", ""))


def is_folder_url(url: str) -> bool:
    return bool(FOLDER_URL_PATTERN.search(urlparse(url).path))


def is_document_url(url: str) -> bool:
    path = urlparse(url).path
    return bool(DOC_URL_PATTERN.search(path)) and not is_folder_url(url)


def safe_name(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "-", name).strip(" .-")
    return name or f"shimo-export-{int(time.time())}"


def document_id_from_url(url: str) -> str:
    parsed = urlparse(normalize_url(url))
    parts = [part for part in parsed.path.split("/") if part]
    return parts[-1] if parts else normalize_url(url)


def item_key(item: ShimoItem) -> str:
    return document_id_from_url(item.url)


def load_url_items(args: argparse.Namespace) -> list[ShimoItem]:
    urls: list[str] = list(args.url)
    if args.urls_file:
        urls.extend(
            line.strip()
            for line in args.urls_file.read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        )
    return [ShimoItem(title="", url=normalize_url(url), last_modified="") for url in urls]


def dedupe_items(items: Iterable[ShimoItem]) -> list[ShimoItem]:
    seen: set[str] = set()
    deduped: list[ShimoItem] = []
    for item in items:
        key = item_key(item)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)
    return deduped



def utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def default_backup_db_path(output_dir: Path) -> Path:
    return output_dir / "backup.json"


def load_backup_db(path: Path) -> dict[str, BackupRecord]:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    documents = data.get("documents", data)
    records: dict[str, BackupRecord] = {}
    for document_id, raw_record in documents.items():
        records[document_id] = BackupRecord(
            document_id=str(raw_record.get("document_id") or document_id),
            title=str(raw_record.get("title") or ""),
            last_modified=str(raw_record.get("last_modified") or ""),
            last_exported=str(raw_record.get("last_exported") or ""),
            export_paths=dict(raw_record.get("export_paths") or {}),
        )
    return records


def save_backup_db(path: Path, records: dict[str, BackupRecord]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 1,
        "updated_at": utc_now(),
        "documents": {
            document_id: asdict(record)
            for document_id, record in sorted(records.items())
        },
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def sync_status(
    item: ShimoItem,
    records: dict[str, BackupRecord],
    formats: Iterable[ExportFormat],
) -> str:
    record = records.get(item_key(item))
    if record is None:
        return "new"
    if item.last_modified and record.last_modified != item.last_modified:
        return "modified"
    if any(export_format not in record.export_paths for export_format in formats):
        return "modified"
    return "skipped"


def update_backup_record(
    records: dict[str, BackupRecord],
    item: ShimoItem,
    export_format: ExportFormat,
    exported_path: Path,
) -> None:
    document_id = item_key(item)
    record = records.get(document_id) or BackupRecord(
        document_id=document_id,
        title="",
        last_modified="",
        last_exported="",
    )
    record.title = item.title or record.title or exported_path.stem
    record.last_modified = item.last_modified or record.last_modified
    record.last_exported = utc_now()
    record.export_paths[export_format] = str(exported_path)
    records[document_id] = record



def try_password_login(
    page: Page, timeout_ms: int, timeout_error: type[Exception]
) -> None:
    username = os.getenv("SHIMO_USERNAME")
    password = os.getenv("SHIMO_PASSWORD")
    if not username or not password:
        return

    page.goto(SHIMO_HOME, wait_until="domcontentloaded")
    for text in ("登录", "Sign in", "Log in"):
        locator = page.get_by_text(text, exact=False).first
        try:
            locator.click(timeout=3_000)
            break
        except timeout_error:
            continue

    user_fields = [
        page.locator('input[type="text"]').first,
        page.locator('input[type="email"]').first,
        page.get_by_placeholder(
            re.compile("手机|邮箱|账号|email|phone|account", re.I)
        ).first,
    ]
    pass_fields = [
        page.locator('input[type="password"]').first,
        page.get_by_placeholder(re.compile("密码|password", re.I)).first,
    ]
    for field in user_fields:
        try:
            field.fill(username, timeout=3_000)
            break
        except timeout_error:
            continue
    for field in pass_fields:
        try:
            field.fill(password, timeout=3_000)
            break
        except timeout_error:
            continue

    for text in ("登录", "Sign in", "Log in"):
        try:
            page.get_by_text(text, exact=True).last.click(timeout=3_000)
            page.wait_for_load_state("networkidle", timeout=timeout_ms)
            return
        except timeout_error:
            continue


def ensure_logged_in(page: Page) -> None:
    if LOGIN_TEXT_PATTERN.search(page.title()) or "/login" in page.url:
        raise RuntimeError(
            "石墨登录状态可能已失效。请先运行："
            "python tools/shimo_export/export_shimo_diary_to_word.py --login-only"
        )


def click_first_visible_text(
    page: Page,
    texts: Iterable[str],
    timeout_ms: int,
    timeout_error: type[Exception],
) -> bool:
    for text in texts:
        locators = [
            page.get_by_role("button", name=re.compile(re.escape(text), re.I)).first,
            page.get_by_text(text, exact=False).first,
        ]
        for locator in locators:
            try:
                locator.click(timeout=min(timeout_ms, 5_000))
                return True
            except timeout_error:
                continue
    return False


def maybe_click_next_page(
    page: Page, timeout_ms: int, timeout_error: type[Exception]
) -> bool:
    for text in NEXT_PAGE_TEXTS:
        locators = [
            page.get_by_role("button", name=re.compile(re.escape(text), re.I)).last,
            page.get_by_text(text, exact=True).last,
        ]
        for locator in locators:
            try:
                if locator.is_disabled(timeout=1_000):
                    continue
                locator.click(timeout=min(timeout_ms, 4_000))
                page.wait_for_load_state("networkidle", timeout=timeout_ms)
                page.wait_for_timeout(800)
                return True
            except timeout_error:
                continue
            except Exception:
                continue
    return False


def collect_visible_items(page: Page) -> tuple[list[ShimoItem], list[ShimoItem]]:
    anchors = page.locator("a[href]").evaluate_all(
        """
        anchors => anchors.map(anchor => {
            const row = anchor.closest('[role="row"], tr, li, .ant-list-item, .file-item') || anchor.parentElement;
            return {
                href: anchor.href,
                text: (anchor.innerText || anchor.textContent || anchor.title || '').trim(),
                title: anchor.title || '',
                rowText: row ? (row.innerText || row.textContent || '').trim() : ''
            };
        })
        """
    )
    documents: list[ShimoItem] = []
    folders: list[ShimoItem] = []
    for anchor in anchors:
        url = normalize_url(str(anchor.get("href", "")))
        if not url.startswith("https://shimo.im/"):
            continue
        title = safe_name(str(anchor.get("text") or anchor.get("title") or ""))
        modified_match = MODIFIED_TIME_PATTERN.search(str(anchor.get("rowText") or ""))
        last_modified = modified_match.group(1).strip() if modified_match else ""
        if is_folder_url(url):
            folders.append(ShimoItem(title=title, url=url, last_modified=last_modified))
        elif is_document_url(url):
            documents.append(ShimoItem(title=title, url=url, last_modified=last_modified))
    return dedupe_items(documents), dedupe_items(folders)


def crawl_current_folder_page(
    page: Page,
    timeout_ms: int,
    timeout_error: type[Exception],
    max_scrolls: int,
) -> tuple[list[ShimoItem], list[ShimoItem]]:
    documents: list[ShimoItem] = []
    folders: list[ShimoItem] = []
    previous_count = -1
    stagnant_rounds = 0

    for _ in range(max_scrolls):
        current_documents, current_folders = collect_visible_items(page)
        documents = dedupe_items([*documents, *current_documents])
        folders = dedupe_items([*folders, *current_folders])
        current_count = len(documents) + len(folders)
        if current_count == previous_count:
            stagnant_rounds += 1
        else:
            stagnant_rounds = 0
        if stagnant_rounds >= 3:
            break
        previous_count = current_count
        page.mouse.wheel(0, 2500)
        page.wait_for_timeout(900)

    while maybe_click_next_page(page, timeout_ms, timeout_error):
        next_documents, next_folders = crawl_current_folder_page(
            page, timeout_ms, timeout_error, max_scrolls
        )
        documents = dedupe_items([*documents, *next_documents])
        folders = dedupe_items([*folders, *next_folders])

    return dedupe_items(documents), dedupe_items(folders)


def crawl_folder(
    context: BrowserContext,
    folder_url: str,
    timeout_ms: int,
    timeout_error: type[Exception],
    max_scrolls: int,
    visited_folders: set[str],
) -> list[ShimoItem]:
    normalized_folder_url = normalize_url(folder_url)
    if normalized_folder_url in visited_folders:
        return []
    visited_folders.add(normalized_folder_url)

    page = context.new_page()
    page.set_default_timeout(timeout_ms)
    print(f"正在读取石墨文件夹：{normalized_folder_url}")
    page.goto(normalized_folder_url, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=timeout_ms)
    ensure_logged_in(page)

    documents, folders = crawl_current_folder_page(
        page, timeout_ms, timeout_error, max_scrolls
    )
    page.close()

    all_documents = list(documents)
    child_folders = [folder for folder in folders if item_key(folder) not in visited_folders]
    if child_folders:
        print(f"发现 {len(child_folders)} 个子文件夹，继续递归读取...")
    for folder in child_folders:
        all_documents.extend(
            crawl_folder(
                context,
                folder.url,
                timeout_ms,
                timeout_error,
                max_scrolls,
                visited_folders,
            )
        )
    return dedupe_items(all_documents)


def trigger_export(
    page: Page,
    export_format: ExportFormat,
    timeout_ms: int,
    timeout_error: type[Exception],
) -> None:
    patterns = EXPORT_PATTERNS[export_format]
    for pattern in patterns:
        try:
            page.get_by_text(pattern).first.click(timeout=3_000)
            return
        except timeout_error:
            pass

    click_first_visible_text(page, MENU_TEXTS, timeout_ms, timeout_error)
    page.wait_for_timeout(500)

    for pattern in patterns:
        try:
            page.get_by_text(pattern).first.click(timeout=8_000)
            return
        except timeout_error:
            pass

    raise RuntimeError(
        f"Could not find Shimo's {export_format.upper()} export button. Log in, "
        "open the document, and verify the current account has export permission."
    )


def expected_output_path(
    output_dir: Path, item: ShimoItem, export_format: ExportFormat
) -> Path | None:
    if not item.title:
        return None
    suffix = ".pdf" if export_format == "pdf" else ".docx"
    return output_dir / f"{safe_name(item.title)}{suffix}"


def export_one(
    context: BrowserContext,
    item: ShimoItem,
    output_dir: Path,
    export_format: ExportFormat,
    timeout_ms: int,
    timeout_error: type[Exception],
) -> Path:
    page = context.new_page()
    page.set_default_timeout(timeout_ms)
    page.goto(item.url, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=timeout_ms)
    ensure_logged_in(page)

    page_title = page.title().replace("- 石墨文档", "").replace("石墨文档", "")
    title = safe_name(item.title or page_title)
    with page.expect_download(timeout=timeout_ms) as download_info:
        trigger_export(page, export_format, timeout_ms, timeout_error)
    download = download_info.value
    suggested = safe_name(download.suggested_filename)
    suffix = Path(suggested).suffix or (".pdf" if export_format == "pdf" else ".docx")
    filename = (
        suggested
        if suggested.lower().endswith(SUPPORTED_SUFFIXES[export_format])
        else f"{title}{suffix}"
    )
    target = output_dir / filename
    download.save_as(target)
    page.close()
    return target


def print_summary(
    new_count: int,
    modified_count: int,
    skipped_count: int,
    failures: list[ExportFailure],
) -> None:
    print("\n导出完成")
    print(f"新增：{new_count}")
    print(f"修改：{modified_count}")
    print(f"跳过：{skipped_count}")
    print(f"失败：{len(failures)}")
    if failures:
        print("\n失败列表：")
        for failure in failures:
            title = failure.title or failure.url
            print(f"- {title} ({failure.reason})")


def main() -> int:
    args = parse_args()
    if not args.login_only and not (args.folder or args.url or args.urls_file):
        print(
            "No documents found. Use --folder, --url, or --urls-file.",
            file=sys.stderr,
        )
        return 2

    args.output_dir.mkdir(parents=True, exist_ok=True)
    backup_db_path = args.backup_db or default_backup_db_path(args.output_dir)
    backup_records = load_backup_db(backup_db_path)

    url_items = load_url_items(args)
    try:
        timeout_error, sync_playwright = load_playwright()
    except MissingPlaywrightError as exc:
        print(str(exc), file=sys.stderr)
        return 3

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            str(args.profile_dir),
            accept_downloads=True,
            headless=args.headless,
        )
        page = context.new_page()
        page.set_default_timeout(args.timeout_ms)
        try_password_login(page, args.timeout_ms, timeout_error)

        if args.login_only:
            page.goto(SHIMO_HOME, wait_until="domcontentloaded")
            print("请在打开的浏览器中完成石墨登录；完成后回到终端按 Enter。")
            input()
            context.close()
            return 0

        folder_items: list[ShimoItem] = []
        if args.folder:
            print("正在读取石墨文件夹...\n")
            visited_folders: set[str] = set()
            for folder_url in args.folder:
                folder_items.extend(
                    crawl_folder(
                        context,
                        folder_url,
                        args.timeout_ms,
                        timeout_error,
                        args.max_scrolls,
                        visited_folders,
                    )
                )

        items = dedupe_items([*url_items, *folder_items])
        if not items:
            print(
                "No documents found. Use --folder, --url, or --urls-file.",
                file=sys.stderr,
            )
            context.close()
            return 2

        formats = requested_formats(args)
        total_jobs = len(items) * len(formats)
        print(f"\n发现 {len(items)} 篇日记")
        print(f"导出格式：{', '.join(formats)}")
        print(f"总任务：{total_jobs}\n")

        new_count = 0
        modified_count = 0
        skipped_count = 0
        failures: list[ExportFailure] = []
        completed_jobs = 0
        for item in items:
            status = "modified" if args.force else sync_status(item, backup_records, formats)
            if status == "skipped" and not args.force:
                skipped_count += 1
                completed_jobs += len(formats)
                print(f"[{completed_jobs}/{total_jobs}]")
                print(f"跳过未修改：{item.title or item.url}\n")
                continue

            exported_any_format = False
            for export_format in formats:
                completed_jobs += 1
                title = item.title or item.url
                remaining = total_jobs - completed_jobs
                print(f"[{completed_jobs}/{total_jobs}]")
                print(f"正在导出：{title}")
                print(f"状态：{status}，格式：{export_format}，剩余：{remaining}")

                try:
                    target = export_one(
                        context,
                        item,
                        args.output_dir,
                        export_format,
                        args.timeout_ms,
                        timeout_error,
                    )
                    update_backup_record(backup_records, item, export_format, target)
                    save_backup_db(backup_db_path, backup_records)
                    exported_any_format = True
                    print(f"✔ 完成：{target}\n")
                except Exception as exc:
                    failures.append(ExportFailure(title=title, url=item.url, reason=str(exc)))
                    print(f"✘ 失败，已跳过：{exc}\n")
                    continue

            if exported_any_format:
                if status == "new":
                    new_count += 1
                else:
                    modified_count += 1

        context.close()

    save_backup_db(backup_db_path, backup_records)
    print(f"备份索引：{backup_db_path}")
    print_summary(new_count, modified_count, skipped_count, failures)
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
