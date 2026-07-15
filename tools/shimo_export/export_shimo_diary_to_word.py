#!/usr/bin/env python3
"""Export Shimo diary documents to Word files with Playwright.

This helper keeps account credentials out of the repository. Prefer using an
existing browser session created with --login-only. If password login is needed,
provide credentials through environment variables instead of command-line flags:
SHIMO_USERNAME and SHIMO_PASSWORD.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:
    from playwright.sync_api import BrowserContext, Page
else:
    BrowserContext = Any
    Page = Any


class MissingPlaywrightError(RuntimeError):
    pass


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
EXPORT_WORD_PATTERNS = [
    re.compile(r"导出.*(Word|word|DOCX|docx|文档)"),
    re.compile(r"(Word|word|DOCX|docx|文档).*导出"),
    re.compile(r"下载.*(Word|word|DOCX|docx|文档)"),
]
MENU_TEXTS = ["更多", "···", "...", "菜单", "文件", "导出", "下载"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download Shimo diary article/document URLs as Word files."
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
        default=Path("shimo-word-export"),
        help="Directory used for downloaded Word files.",
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
    return parser.parse_args()


def load_urls(args: argparse.Namespace) -> list[str]:
    urls: list[str] = list(args.url)
    if args.urls_file:
        urls.extend(
            line.strip()
            for line in args.urls_file.read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        )
    return urls


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


def trigger_word_export(
    page: Page, timeout_ms: int, timeout_error: type[Exception]
) -> None:
    # Some Shimo pages expose the export action directly; others hide it in menus.
    for pattern in EXPORT_WORD_PATTERNS:
        try:
            page.get_by_text(pattern).first.click(timeout=3_000)
            return
        except timeout_error:
            pass

    click_first_visible_text(page, MENU_TEXTS, timeout_ms, timeout_error)
    page.wait_for_timeout(500)

    for pattern in EXPORT_WORD_PATTERNS:
        try:
            page.get_by_text(pattern).first.click(timeout=8_000)
            return
        except timeout_error:
            pass

    raise RuntimeError(
        "Could not find Shimo's Word export button. Log in, open the document, "
        "and verify the current account has export permission."
    )


def safe_name(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "-", name).strip()
    return name or f"shimo-export-{int(time.time())}"


def export_one(
    context: BrowserContext,
    url: str,
    output_dir: Path,
    timeout_ms: int,
    timeout_error: type[Exception],
) -> Path:
    page = context.new_page()
    page.set_default_timeout(timeout_ms)
    page.goto(url, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=timeout_ms)

    title = safe_name(page.title().replace("- 石墨文档", "").replace("石墨文档", ""))
    with page.expect_download(timeout=timeout_ms) as download_info:
        trigger_word_export(page, timeout_ms, timeout_error)
    download = download_info.value
    suggested = safe_name(download.suggested_filename)
    suffix = Path(suggested).suffix or ".docx"
    filename = (
        suggested
        if suggested.lower().endswith((".doc", ".docx"))
        else f"{title}{suffix}"
    )
    target = output_dir / filename
    download.save_as(target)
    page.close()
    return target


def main() -> int:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)

    urls = load_urls(args)
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

        if not urls:
            print("No URLs provided. Use --url or --urls-file.", file=sys.stderr)
            context.close()
            return 2

        exported: list[Path] = []
        for url in urls:
            print(f"Exporting {url} ...")
            exported.append(
                export_one(context, url, args.output_dir, args.timeout_ms, timeout_error)
            )
        context.close()

    print("Exported files:")
    for path in exported:
        print(f"- {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
