#!/usr/bin/env python3
"""Shared resource budgets for untrusted email and attachment processing."""

from __future__ import annotations

import multiprocessing
import queue
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable


MB = 1024 * 1024


class LimitExceeded(RuntimeError):
    """Raised when input exceeds a configured processing budget."""


@dataclass(frozen=True)
class ResourceLimits:
    force: bool = False

    max_email_bytes: int = 100 * MB
    max_single_attachment_bytes: int = 100 * MB
    max_total_attachment_bytes: int = 500 * MB
    max_attachments: int = 100
    max_inline_image_bytes: int = 10 * MB
    max_generated_html_bytes: int = 50 * MB
    max_pdf_pages_per_attachment: int = 1000
    max_archive_depth: int = 3
    max_archive_uncompressed_bytes: int = 500 * MB
    max_text_attachment_bytes: int = 100 * MB

    timeout_email_seconds: int = 120
    timeout_playwright_email_seconds: int = 45
    timeout_weasyprint_seconds: int = 60
    timeout_ghostscript_seconds: int = 120
    timeout_libreoffice_seconds: int = 90
    timeout_docx2pdf_seconds: int = 120
    timeout_html_attachment_seconds: int = 45
    timeout_image_seconds: int = 30
    timeout_pdf_seconds: int = 30
    timeout_archive_seconds: int = 30
    timeout_merge_email_seconds: int = 300

    def enabled(self) -> bool:
        return not self.force

    def timeout(self, seconds: int) -> int | None:
        return seconds if self.enabled() else None


DEFAULT_LIMITS = ResourceLimits()


def format_bytes(value: int) -> str:
    if value >= MB and value % MB == 0:
        return f"{value // MB} MB"
    return f"{value} bytes"


def check_file_size(path: Path, max_bytes: int, label: str, limits: ResourceLimits) -> None:
    if not limits.enabled():
        return
    size = path.stat().st_size
    if size > max_bytes:
        raise LimitExceeded(
            f"{label} is {format_bytes(size)}, exceeding the {format_bytes(max_bytes)} limit"
        )


def check_bytes_size(size: int, max_bytes: int, label: str, limits: ResourceLimits) -> None:
    if limits.enabled() and size > max_bytes:
        raise LimitExceeded(
            f"{label} is {format_bytes(size)}, exceeding the {format_bytes(max_bytes)} limit"
        )


def check_text_size(value: str, max_bytes: int, label: str, limits: ResourceLimits) -> None:
    if not limits.enabled():
        return
    check_bytes_size(len(value.encode("utf-8", errors="replace")), max_bytes, label, limits)


def check_count(count: int, max_count: int, label: str, limits: ResourceLimits) -> None:
    if limits.enabled() and count > max_count:
        raise LimitExceeded(f"{label} count {count} exceeds the {max_count} limit")


def _child_entry(
    result_queue: multiprocessing.Queue,
    func: Callable[..., Any],
    args: tuple,
    kwargs: dict,
) -> None:
    try:
        result_queue.put(("ok", func(*args, **kwargs)))
    except BaseException as exc:
        result_queue.put(("err", exc.__class__.__name__, str(exc)))


def run_with_timeout(
    func: Callable[..., Any],
    timeout_seconds: int | None,
    description: str,
    *args: Any,
    **kwargs: Any,
) -> Any:
    """Run a picklable function in a killable child process when a timeout is set."""
    if timeout_seconds is None:
        return func(*args, **kwargs)

    ctx = multiprocessing.get_context()
    result_queue: multiprocessing.Queue = ctx.Queue(maxsize=1)
    proc = ctx.Process(target=_child_entry, args=(result_queue, func, args, kwargs))
    proc.start()
    proc.join(timeout_seconds)

    if proc.is_alive():
        proc.terminate()
        proc.join(5)
        if proc.is_alive():
            proc.kill()
            proc.join()
        raise TimeoutError(f"{description} exceeded {timeout_seconds}s timeout")

    try:
        status = result_queue.get_nowait()
    except queue.Empty:
        if proc.exitcode == 0:
            return None
        raise RuntimeError(f"{description} failed with exit code {proc.exitcode}")

    if status[0] == "ok":
        return status[1]

    exc_name, message = status[1], status[2]
    raise RuntimeError(f"{description} failed: {exc_name}: {message}")
