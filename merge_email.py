#!/usr/bin/env python3
"""
merge_email.py — Merge each email PDF with its attachments into a single combined PDF.

For every email folder in output_dir, produces one merged PDF in a sibling
_combined folder:

    output/                                      <- source (output_dir)
      2024.03.15 09.30 - Q1 Budget Review/
        2024.03.15 09.30 - Q1 Budget Review.pdf  <- email body
        metadata.json
        attachments/
          report.xlsx
          contract.pdf
          photo.jpg

    output_combined/                             <- created automatically
      2024.03.15 09.30 - Q1 Budget Review.pdf   <- email + all attachments merged
      errors.txt                                 <- written on any failure

Merge order: email PDF first, then attachments in filename order.
Each attachment gets a named PDF bookmark pointing to its first page.

If any attachment in an email cannot be converted the entire email is skipped
and the failure is flagged to the console and appended to errors.txt.

Conversion support
------------------
  .pdf                     merged directly (pikepdf)
  .docx .doc .rtf .odt     docx2pdf (Word/LibreOffice) -> LibreOffice fallback
  .xlsx .xls .pptx .ppt    LibreOffice headless
  .jpg .jpeg .png .gif
  .bmp .tiff .tif .webp    img2pdf (lossless, preserves DPI)
  .txt .csv                Playwright (rendered as preformatted text)
  .html .htm               Playwright
  .eml .msg                emailtopdf pipeline (recursive)
  .zip                     extracted, contents converted; nested bookmarks
  .rar                     extracted via rarfile + UnRAR binary; nested bookmarks
                           (requires: pip install rarfile + WinRAR/UnRAR in PATH)
  anything else            unsupported -> error, email skipped

Usage:
    python merge_email.py                         # uses output/ next to script
    python merge_email.py path/to/output_dir/     # explicit path
"""

import sys
import shutil
import tempfile
import subprocess
import traceback
import zipfile
import re as _re
from datetime import datetime
from html import escape
from pathlib import Path

# UTF-8 console output — prevents crashes on non-Latin subjects/paths (Windows CP1252)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------------------
# Dependency check
# ---------------------------------------------------------------------------

try:
    import pikepdf
except ImportError:
    sys.exit("pikepdf is required: pip install pikepdf")


# ---------------------------------------------------------------------------
# File-type sets
# ---------------------------------------------------------------------------

OFFICE_WORD_EXTS  = {".docx", ".doc", ".rtf", ".odt"}   # .doc handled via docx2pdf -> LibreOffice
OFFICE_OTHER_EXTS = {".xlsx", ".xls", ".pptx", ".ppt", ".ods", ".odp"}
IMAGE_EXTS        = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
TEXT_EXTS         = {".txt", ".csv"}
HTML_EXTS         = {".html", ".htm"}
EMAIL_EXTS        = {".eml", ".msg"}
ARCHIVE_EXTS      = {".zip", ".rar"}

_ARCHIVE_MAX_DEPTH        = 3                  # maximum nesting depth for archives-within-archives
_ARCHIVE_MAX_UNCOMPRESSED = 500 * 1024 * 1024  # 500 MB uncompressed size guard

# Common Windows locations for the UnRAR binary (tried in order if not in PATH)
_UNRAR_SEARCH_PATHS = [
    r"C:\Program Files\WinRAR\UnRAR.exe",
    r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
    r"C:\Program Files\WinRAR\WinRAR.exe",
]


# ---------------------------------------------------------------------------
# Individual converters
# ---------------------------------------------------------------------------

def _libreoffice_to_pdf(src: Path, out_dir: Path) -> Path:
    """Convert any LibreOffice-compatible file to PDF via headless LibreOffice."""
    for exe in ("soffice", "libreoffice"):
        if shutil.which(exe):
            break
    else:
        raise RuntimeError(
            "LibreOffice not found in PATH — install from https://www.libreoffice.org/"
        )
    result = subprocess.run(
        [exe, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(src)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice exited {result.returncode}: {result.stderr.strip()}")
    pdf = out_dir / f"{src.stem}.pdf"
    if not pdf.exists():
        raise RuntimeError(f"LibreOffice did not produce expected output: {pdf}")
    return pdf


def _docx2pdf_to_pdf(src: Path, out_dir: Path) -> Path:
    """Convert a Word-compatible file to PDF via docx2pdf (uses Word on Windows).

    Retries once after a short delay — Word COM can disconnect mid-session when
    converting many documents in sequence.
    """
    try:
        from docx2pdf import convert
    except ImportError:
        raise RuntimeError("docx2pdf not installed: pip install docx2pdf")
    import time
    out = out_dir / f"{src.stem}.pdf"
    last_exc = None
    for attempt in range(2):
        try:
            convert(str(src), str(out))
            if out.exists():
                return out
            raise RuntimeError(f"docx2pdf produced no output for '{src.name}'")
        except Exception as exc:
            last_exc = exc
            if out.exists():
                try:
                    out.unlink()      # remove any partial file before retry
                except OSError:
                    pass              # Word may still hold the output open; ignore
            if attempt == 0:
                time.sleep(3)         # give Word time to recover before retry
    raise RuntimeError(f"docx2pdf failed after retry: {last_exc or 'no output produced'}")


def _image_to_pdf(src: Path, out_dir: Path) -> Path:
    """Convert an image to PDF via img2pdf (lossless, preserves DPI)."""
    try:
        import img2pdf
    except ImportError:
        raise RuntimeError("img2pdf not installed: pip install img2pdf")
    out = out_dir / f"{src.stem}.pdf"
    with open(out, "wb") as f:
        f.write(img2pdf.convert(str(src)))
    return out


def _playwright_html_to_pdf(html: str, out: Path) -> None:
    """Render an HTML string to PDF via Playwright (headless Chromium)."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "playwright not installed: pip install playwright && playwright install chromium"
        )
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page()
        page.route("**/*", lambda route: route.abort())
        page.set_content(html, wait_until="domcontentloaded")
        page.pdf(path=str(out), print_background=True)
        browser.close()


def _text_to_pdf(src: Path, out_dir: Path) -> Path:
    """Render a plain-text or CSV file to PDF via Playwright."""
    text = src.read_text(encoding="utf-8", errors="replace")
    html = (
        "<!DOCTYPE html><html><head><meta charset='UTF-8'>"
        "<style>body{font-family:monospace;font-size:11px;"
        "white-space:pre-wrap;word-wrap:break-word;margin:24px}</style>"
        f"</head><body>{escape(text)}</body></html>"
    )
    out = out_dir / f"{src.stem}.pdf"
    _playwright_html_to_pdf(html, out)
    return out


def _html_file_to_pdf(src: Path, out_dir: Path) -> Path:
    """Render an HTML file to PDF via Playwright."""
    html = src.read_text(encoding="utf-8", errors="replace")
    out = out_dir / f"{src.stem}.pdf"
    _playwright_html_to_pdf(html, out)
    return out


def _email_to_pdf(src: Path, out_dir: Path) -> Path:
    """Convert an embedded .eml/.msg file to PDF using the emailtopdf pipeline."""
    script_dir = Path(__file__).parent
    if str(script_dir) not in sys.path:
        sys.path.insert(0, str(script_dir))
    import emailtopdf as etp

    parse_fn = etp.parse_eml if src.suffix.lower() == ".eml" else etp.parse_msg
    data     = parse_fn(src)
    html     = etp.build_html(data)
    html     = etp._inline_css(html)
    out      = out_dir / f"{src.stem}.pdf"
    with etp.PlaywrightPool() as pool:
        pool.render_pdf(html, out)
    return out


# ---------------------------------------------------------------------------
# Magic-byte file-type detection (for extensionless attachments)
# ---------------------------------------------------------------------------

# (magic_bytes, guessed_extension)
# Ordered: more-specific signatures first.
# OLE2 (\xd0\xcf\x11\xe0) is NOT listed here — it is handled separately in
# convert_to_pdf because both .msg and .doc share the same magic bytes.
_MAGIC_SIGNATURES = [
    (b"%PDF",              ".pdf"),
    (b"Rar!\x1a\x07",     ".rar"),
    (b"PK\x03\x04",       ".zip"),    # ZIP / OOXML (.docx .xlsx .pptx are ZIP inside)
    (b"{\rtf",            ".rtf"),
    (b"<html",            ".html"),
    (b"<!doctype",        ".html"),
    (b"<?xml",            ".xml"),
]

_OLE2_MAGIC = b"\xd0\xcf\x11\xe0"   # OLE2 compound document header

# EML-style headers that appear at the start of plain-text email files
_EML_HEADER_PREFIXES = (b"MIME-Version:", b"Content-Type:", b"From:", b"Date:", b"Received:")


def _sniff_extension(src: Path) -> str:
    """
    Read the first 16 bytes of src and return a best-guess file extension
    (including the leading dot), or '' if the type cannot be determined.

    Also handles extensionless files that are plain-text emails (.eml)
    by scanning the first 512 bytes for RFC 2822 header keywords.
    """
    try:
        header = src.read_bytes()[:512]
    except OSError:
        return ""

    peek = header[:16]
    # OLE2 is ambiguous — signal it separately so convert_to_pdf can try both
    if peek.startswith(_OLE2_MAGIC):
        return ".ole2"

    for magic, ext in _MAGIC_SIGNATURES:
        if peek.lower().startswith(magic.lower()):
            return ext

    # Plain-text email check — look for RFC 2822 headers in the first 512 bytes
    for prefix in _EML_HEADER_PREFIXES:
        if prefix.lower() in header.lower():
            return ".eml"

    # If it decodes cleanly as UTF-8, treat as plain text
    try:
        header.decode("utf-8")
        return ".txt"
    except UnicodeDecodeError:
        pass

    return ""


# ---------------------------------------------------------------------------
# Dispatcher
# ---------------------------------------------------------------------------

def convert_to_pdf(src: Path, out_dir: Path) -> Path:
    """
    Convert src to a PDF placed in out_dir.
    Returns the path to the produced PDF.
    Raises RuntimeError (with a human-readable message) on failure.
    Raises ValueError for unsupported file types.
    """
    ext = src.suffix.lower()

    # Extensionless files — try to detect type from magic bytes before giving up
    if not ext:
        ext = _sniff_extension(src)
        if not ext:
            raise ValueError(
                f"Cannot determine file type for '{src.name}' "
                f"(no extension, unrecognised file signature)"
            )
        # Work on a renamed copy so converters that rely on the extension work correctly
        renamed = out_dir / f"{src.name}{ext}"
        shutil.copy2(src, renamed)
        return convert_to_pdf(renamed, out_dir)

    # OLE2 compound document — could be .msg (Outlook embedded email) or .doc/.xls/.ppt.
    # Both share the same magic bytes, so we try .msg first (more common as an
    # extensionless email attachment) then fall back to .doc.
    if ext == ".ole2":
        msg_path = out_dir / f"{src.stem}.msg"
        shutil.copy2(src, msg_path)
        try:
            return _email_to_pdf(msg_path, out_dir)
        except Exception:
            pass
        doc_path = out_dir / f"{src.stem}.doc"
        shutil.copy2(src, doc_path)
        return convert_to_pdf(doc_path, out_dir)

    if ext == ".pdf":
        return src  # no conversion needed; caller uses the original

    if ext in OFFICE_WORD_EXTS:
        # Try docx2pdf (Word) first; fall back to LibreOffice
        try:
            return _docx2pdf_to_pdf(src, out_dir)
        except Exception as primary:
            try:
                return _libreoffice_to_pdf(src, out_dir)
            except Exception as fallback:
                raise RuntimeError(
                    f"docx2pdf failed ({primary}); LibreOffice fallback also failed ({fallback})"
                )

    if ext in OFFICE_OTHER_EXTS:
        return _libreoffice_to_pdf(src, out_dir)

    if ext in IMAGE_EXTS:
        return _image_to_pdf(src, out_dir)

    if ext in TEXT_EXTS:
        return _text_to_pdf(src, out_dir)

    if ext in HTML_EXTS:
        return _html_file_to_pdf(src, out_dir)

    if ext in EMAIL_EXTS:
        return _email_to_pdf(src, out_dir)

    raise ValueError(f"Unsupported file type: '{src.suffix}' ({src.name})")


# ---------------------------------------------------------------------------
# Archive expansion (ZIP and RAR)
# ---------------------------------------------------------------------------

def _find_unrar() -> str:
    """
    Return the path to a usable UnRAR binary, or raise RuntimeError.
    Checks PATH first, then common Windows install locations.
    """
    for name in ("unrar", "UnRAR"):
        if shutil.which(name):
            return shutil.which(name)
    for path in _UNRAR_SEARCH_PATHS:
        if Path(path).exists():
            return path
    raise RuntimeError(
        "UnRAR binary not found. Install WinRAR (https://www.rarlab.com/) "
        "or the free UnRAR utility and ensure it is in PATH."
    )


def _safe_extractall(af, extract_dir: Path) -> None:
    """
    Extract all archive members to extract_dir, rejecting any member whose resolved
    destination path escapes extract_dir (zip-slip / path-traversal guard).
    Works with both zipfile.ZipFile and rarfile.RarFile instances.
    """
    resolved_base = extract_dir.resolve()
    for member in af.infolist():
        dest = (extract_dir / member.filename).resolve()
        try:
            dest.relative_to(resolved_base)
        except ValueError:
            raise RuntimeError(
                f"Path traversal in archive member '{member.filename}' — extraction aborted."
            )
        af.extract(member, extract_dir)


def _check_multipart_rar(src: Path) -> None:
    """
    Raise RuntimeError if src looks like a RAR continuation part (partN where N > 1).
    Only part1 (or a non-split RAR) can be used as an entry point for extraction.
    """
    # New-style: file.part2.rar, file.part3.rar, …
    m = _re.search(r"\.part(\d+)\.rar$", src.name, _re.IGNORECASE)
    if m and int(m.group(1)) > 1:
        raise RuntimeError(
            f"'{src.name}' is a multi-part RAR continuation (part {m.group(1)}). "
            "Only the first part (.part1.rar) can be extracted; "
            "subsequent parts must be present in the same folder."
        )
    # Old-style: file.r00, file.r01, … are continuation parts — .rar is the first
    # so nothing to check for old-style; the opener will fail if parts are missing.


def _expand_archive(src: Path, out_dir: Path, depth: int = 0) -> list:
    """
    Extract a ZIP or RAR archive and convert its contents to PDFs.

    Returns a list of merge-part tuples: (label, pdf_path_or_None, children).
      - Regular file      → (filename, pdf_path, [])
      - Nested archive    → (archname, None,     [child_parts …])

    Raises RuntimeError on unreadable / oversized / too-deeply-nested archives,
    missing unrar binary (RAR only), or if any contained file cannot be converted.
    """
    if depth > _ARCHIVE_MAX_DEPTH:
        raise RuntimeError(
            f"Archive nesting exceeds maximum depth ({_ARCHIVE_MAX_DEPTH}): {src.name}"
        )

    ext = src.suffix.lower()

    if ext == ".rar":
        _check_multipart_rar(src)
        try:
            import rarfile
        except ImportError:
            raise RuntimeError("rarfile not installed: pip install rarfile")
        # Point rarfile at the unrar binary (only needs to happen once per process,
        # but setting it each time is harmless and avoids global-state concerns)
        rarfile.UNRAR_TOOL = _find_unrar()
        archive_cls   = rarfile.RarFile
        bad_exc       = (rarfile.BadRarFile, rarfile.NotRarFile, rarfile.NeedFirstVolume)
        archive_label = "RAR"
    else:
        archive_cls   = zipfile.ZipFile
        bad_exc       = (zipfile.BadZipFile,)
        archive_label = "zip"

    try:
        with archive_cls(src) as af:
            total_bytes = sum(info.file_size for info in af.infolist())
            if total_bytes > _ARCHIVE_MAX_UNCOMPRESSED:
                raise RuntimeError(
                    f"{archive_label} uncompressed size ({total_bytes // (1024 * 1024)} MB) "
                    f"exceeds {_ARCHIVE_MAX_UNCOMPRESSED // (1024 * 1024)} MB safety limit: {src.name}"
                )
            extract_dir = out_dir / "extracted"
            extract_dir.mkdir()
            _safe_extractall(af, extract_dir)
    except bad_exc as exc:
        raise RuntimeError(f"Cannot read {archive_label} '{src.name}': {exc}")

    all_files = sorted(f for f in extract_dir.rglob("*") if f.is_file())
    if not all_files:
        raise RuntimeError(f"'{src.name}' is empty or contains no files")

    parts = []
    for idx, item in enumerate(all_files):
        label    = str(item.relative_to(extract_dir))  # preserves subdir paths
        item_out = out_dir / f"item_{idx}"
        item_out.mkdir()

        if item.suffix.lower() in ARCHIVE_EXTS:
            children = _expand_archive(item, item_out, depth + 1)
            parts.append((label, None, children))
        else:
            pdf = convert_to_pdf(item, item_out)
            parts.append((label, pdf, []))

    return parts


# ---------------------------------------------------------------------------
# PDF merging with bookmarks
# ---------------------------------------------------------------------------

def merge_pdfs(parts: list, output_path: Path) -> None:
    """
    Merge parts into one PDF with a nested bookmark outline.

    Each part is a tuple: (label, pdf_path_or_None, children).
      - pdf_path_or_None: Path to a PDF to append, or None for container entries
        (e.g. a zip whose pages come entirely from its children).
      - children: list of the same tuple structure (may be empty).

    A container entry's bookmark points to the page where its first child starts,
    which is correct because add_pages() records the page counter before
    processing children.

    All source PDFs are kept open until save() completes.
    """
    merged  = pikepdf.Pdf.new()
    sources = []

    def add_pages(items):
        """Recursively append pages; return outline data (label, page, children)."""
        result = []
        for label, pdf_path, children in items:
            page_num = len(merged.pages)          # bookmark points here

            if pdf_path is not None:
                src = pikepdf.Pdf.open(pdf_path)
                sources.append(src)
                merged.pages.extend(src.pages)

            child_data = add_pages(children) if children else []
            result.append((label, page_num, child_data))
        return result

    def build_outline(parent, items):
        """Recursively attach OutlineItems to parent (a list or .children)."""
        for label, page_num, child_data in items:
            oi = pikepdf.OutlineItem(label, page_num)
            if child_data:
                build_outline(oi.children, child_data)
            parent.append(oi)

    try:
        outline_data = add_pages(parts)

        with merged.open_outline() as outline:
            build_outline(outline.root, outline_data)

        merged.save(output_path)

    finally:
        for src in sources:
            src.close()
        merged.close()


# ---------------------------------------------------------------------------
# Per-folder processing
# ---------------------------------------------------------------------------

def _find_email_folders(output_dir: Path) -> list:
    """Return all subdirectories that contain an email PDF ({dir}/{dir}.pdf)."""
    return sorted(
        d for d in output_dir.rglob("*")
        if d.is_dir() and (d / f"{d.name}.pdf").exists()
    )


def process_email_folder(
    email_folder: Path,
    combined_dir: Path,
    errors: list,
) -> bool:
    """
    Produce a combined PDF for one email folder.
    Returns True on success, False if any conversion failed (error appended to errors).
    """
    email_pdf      = email_folder / f"{email_folder.name}.pdf"
    attachments_dir = email_folder / "attachments"
    attachments    = sorted(
        f for f in attachments_dir.iterdir() if f.is_file()
    ) if attachments_dir.exists() else []

    out_path = combined_dir / f"{email_folder.name}.pdf"

    # No attachments — copy the email PDF directly
    if not attachments:
        shutil.copy2(email_pdf, out_path)
        print(f"  (no attachments — email PDF copied)")
        return True

    with tempfile.TemporaryDirectory() as _tmp:
        tmp_dir = Path(_tmp)
        # Each part: (label, pdf_path_or_None, children)
        parts   = [("Email", email_pdf, [])]

        for i, att in enumerate(attachments):
            att_tmp = tmp_dir / f"att_{i}"
            att_tmp.mkdir()
            try:
                if att.suffix.lower() in ARCHIVE_EXTS:
                    children = _expand_archive(att, att_tmp)
                    parts.append((att.name, None, children))
                    n_files = sum(1 for _, p, _ in children if p is not None)
                    print(f"  [OK]     {att.name}  ({n_files} file(s) inside)")
                else:
                    pdf = convert_to_pdf(att, att_tmp)
                    parts.append((att.name, pdf, []))
                    print(f"  [OK]     {att.name}")
            except Exception as exc:
                msg = (
                    f"[{email_folder.name}] "
                    f"Failed to convert '{att.name}': {exc}"
                )
                errors.append(msg)
                print(f"  [ERROR]  {att.name}: {exc}")
                return False

        merge_pdfs(parts, out_path)

    n = len(parts) - 1  # attachments only (excludes email entry)
    print(f"  -> merged ({n} attachment{'s' if n != 1 else ''})")
    return True


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    if len(sys.argv) > 1:
        output_dir = Path(sys.argv[1]).resolve()
    else:
        output_dir = Path(__file__).parent / "output"

    if not output_dir.exists() or not output_dir.is_dir():
        print(f"Error: directory not found: {output_dir}", file=sys.stderr)
        sys.exit(1)

    combined_dir = output_dir.parent / f"{output_dir.name}_combined"
    combined_dir.mkdir(parents=True, exist_ok=True)
    errors_txt = combined_dir / "errors.txt"

    print(f"Source : {output_dir}")
    print(f"Dest   : {combined_dir}")
    print()

    folders = _find_email_folders(output_dir)
    if not folders:
        print("No email folders found.")
        return

    errors    = []
    succeeded = 0
    failed    = 0

    for folder in folders:
        print(f"[{folder.name}]")
        ok = process_email_folder(folder, combined_dir, errors)
        if ok:
            succeeded += 1
        else:
            failed += 1
        print()

    print(f"Done: {succeeded} merged, {failed} failed.")

    if errors:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(errors_txt, "a", encoding="utf-8") as f:
            f.write(f"\n=== {timestamp} ===\n")
            for msg in errors:
                f.write(f"{msg}\n")
        print(f"\n{failed} error(s) logged to: {errors_txt}")


if __name__ == "__main__":
    main()
