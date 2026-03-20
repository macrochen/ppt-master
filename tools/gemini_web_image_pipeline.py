#!/usr/bin/env python3
"""Automate Gemini Web image generation and local post-processing.

This tool reads image prompts from a markdown file, submits them to Gemini Web
through Playwright, waits for the generated image download button to appear,
downloads the image, removes the Gemini watermark, and saves the final file to
the target output directory.

It is designed to work with prompt files shaped like:
  ### 图片 01: cover_hero.png
  **提示词 (Prompt)**:
  ```text
  ...
  ```
  **负面提示词 (Negative Prompt)**:
  ```text
  ...
  ```
"""

from __future__ import annotations

import argparse
import asyncio
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List

from playwright.async_api import Locator, Page, TimeoutError, async_playwright


DOWNLOAD_BUTTON_SELECTOR = ",".join(
    [
        "button[aria-label*='Download']",
        "button[aria-label*='下载']",
        "[role='button'][aria-label*='Download']",
        "[role='button'][aria-label*='下载']",
        "button:has-text('Download')",
        "button:has-text('下载')",
        "[role='button']:has-text('Download')",
        "[role='button']:has-text('下载')",
    ]
)

SEND_BUTTON_SELECTOR = ",".join(
    [
        "button[aria-label*='Send']",
        "button[aria-label*='发送']",
        ".send-button",
    ]
)


@dataclass
class PromptItem:
    filename: str
    prompt: str
    negative_prompt: str


def parse_prompt_items(markdown_path: Path) -> List[PromptItem]:
    content = markdown_path.read_text(encoding="utf-8")
    pattern = re.compile(
        r"^### 图片 \d+: (?P<filename>[^\n]+)\n"
        r".*?\*\*提示词 \(Prompt\)\*\*:\n```text\n(?P<prompt>.*?)\n```\n\n"
        r"\*\*负面提示词 \(Negative Prompt\)\*\*:\n```text\n(?P<negative>.*?)\n```",
        re.S | re.M,
    )

    items: List[PromptItem] = []
    for match in pattern.finditer(content):
        items.append(
            PromptItem(
                filename=match.group("filename").strip(),
                prompt=match.group("prompt").strip(),
                negative_prompt=match.group("negative").strip(),
            )
        )

    if not items:
        raise ValueError(f"No prompt items found in {markdown_path}")

    return items


def filter_items(items: Iterable[PromptItem], names: set[str] | None, limit: int | None) -> List[PromptItem]:
    selected = [item for item in items if not names or item.filename in names]
    if limit is not None:
        selected = selected[:limit]
    return selected


def build_prompt_text(item: PromptItem) -> str:
    return (
        f"{item.prompt}\n\n"
        "Negative prompt:\n"
        f"{item.negative_prompt}\n"
    )


async def wait_for_input_box(page: Page) -> Locator:
    while True:
        locator = page.locator("div[contenteditable='true'], textarea").first
        if await locator.count() and await locator.is_visible():
            return locator
        await asyncio.sleep(2)


async def count_visible_download_buttons(page: Page) -> int:
    locator = page.locator(DOWNLOAD_BUTTON_SELECTOR)
    count = await locator.count()
    visible = 0
    for idx in range(count):
        button = locator.nth(idx)
        try:
            if await button.is_visible():
                visible += 1
        except Exception:
            continue
    return visible


async def get_last_new_download_button(page: Page, previous_visible: int, timeout_ms: int) -> Locator:
    end_time = asyncio.get_running_loop().time() + timeout_ms / 1000
    locator = page.locator(DOWNLOAD_BUTTON_SELECTOR)

    while asyncio.get_running_loop().time() < end_time:
        count = await locator.count()
        visible_buttons: List[Locator] = []
        for idx in range(count):
            button = locator.nth(idx)
            try:
                if await button.is_visible():
                    visible_buttons.append(button)
            except Exception:
                continue

        if len(visible_buttons) > previous_visible:
            return visible_buttons[-1]

        await asyncio.sleep(2)

    raise TimeoutError(f"Timed out waiting for a new download button after {timeout_ms}ms")


async def submit_prompt(page: Page, prompt_text: str) -> None:
    input_box = await wait_for_input_box(page)
    await input_box.fill(prompt_text)
    await asyncio.sleep(1)

    send_button = page.locator(SEND_BUTTON_SELECTOR).first
    if await send_button.count() and await send_button.is_visible():
        await send_button.click()
    else:
        await page.keyboard.press("Enter")


def remove_watermark(remover_python: Path, remover_script: Path, input_path: Path, output_path: Path) -> None:
    result = subprocess.run(
        [
            str(remover_python),
            str(remover_script),
            str(input_path),
            "-o",
            str(output_path),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "watermark removal failed")


async def run_item(
    page: Page,
    item: PromptItem,
    output_dir: Path,
    raw_dir: Path,
    remover_python: Path,
    remover_script: Path,
    timeout_ms: int,
) -> Path:
    prompt_text = build_prompt_text(item)
    previous_downloads = await count_visible_download_buttons(page)

    print(f"\n=== Generating: {item.filename} ===")
    await submit_prompt(page, prompt_text)
    print("Prompt sent, waiting for generated image...")

    download_button = await get_last_new_download_button(page, previous_downloads, timeout_ms)
    print("Download button detected, downloading...")

    async with page.expect_download(timeout=timeout_ms) as download_info:
        await download_button.click()
    download = await download_info.value

    raw_path = raw_dir / f"{Path(item.filename).stem}_raw.png"
    await download.save_as(str(raw_path))
    print(f"Raw download saved: {raw_path}")

    final_path = output_dir / item.filename
    remove_watermark(remover_python, remover_script, raw_path, final_path)
    print(f"Final image saved: {final_path}")
    return final_path


async def run_pipeline(args: argparse.Namespace) -> List[Path]:
    prompt_items = parse_prompt_items(Path(args.prompt_markdown))
    name_filter = set(args.names.split(",")) if args.names else None
    prompt_items = filter_items(prompt_items, name_filter, args.limit)
    if not prompt_items:
        raise ValueError("No prompt items selected")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    raw_dir = output_dir / "raw_downloads"
    raw_dir.mkdir(parents=True, exist_ok=True)

    remover_python = Path(args.remover_python)
    remover_script = Path(args.remover_script)
    profile_dir = Path(args.profile_dir)

    results: List[Path] = []
    async with async_playwright() as p:
        for item in prompt_items:
            print("Opening Gemini Web...")
            context = await p.chromium.launch_persistent_context(
                user_data_dir=str(profile_dir),
                headless=False,
                args=["--disable-blink-features=AutomationControlled"],
            )
            try:
                page = context.pages[0] if context.pages else await context.new_page()
                await page.goto("https://gemini.google.com/app", wait_until="domcontentloaded")
                final_path = await run_item(
                    page,
                    item,
                    output_dir,
                    raw_dir,
                    remover_python,
                    remover_script,
                    args.timeout_ms,
                )
                results.append(final_path)
            finally:
                await context.close()

            await asyncio.sleep(args.pause_seconds)

    return results


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Batch automate Gemini Web image generation")
    parser.add_argument("--prompt_markdown", required=True, help="Path to image_prompts.md")
    parser.add_argument("--output_dir", required=True, help="Where final images should be saved")
    parser.add_argument("--names", help="Comma-separated subset of filenames to generate")
    parser.add_argument("--limit", type=int, help="Only process the first N selected prompts")
    parser.add_argument("--timeout_ms", type=int, default=600000, help="Per-image timeout in milliseconds")
    parser.add_argument("--pause_seconds", type=float, default=2.0, help="Delay between generations")
    parser.add_argument(
        "--profile_dir",
        default=str(Path.home() / ".gemini_automation_profile"),
        help="Playwright persistent profile directory",
    )
    parser.add_argument("--remover_python", required=True, help="Python executable for watermark remover")
    parser.add_argument("--remover_script", required=True, help="Path to remover.py")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        results = asyncio.run(run_pipeline(args))
    except KeyboardInterrupt:
        print("\nInterrupted by user")
        return 130
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1

    print("\n=== Completed ===")
    for path in results:
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
