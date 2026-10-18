#!/usr/bin/env python3
"""
merge_email.py — Merge each email PDF with its attachments into a single combined PDF.

Walks <folder> recursively, at any depth, until it reaches an "email folder":
a leaf folder that holds the unpacked email (email PDF + metadata.json, plus an
optional attachments/ subfolder). Walking stops at that level, so nothing
inside an email folder is mistaken for another email.

For every email folder found, one merged PDF is written to a sibling
<folder>_combined tree that mirrors the source subfolder structure:

    Matter/                                          <- source (<folder>)
      Custodian A/
        Inbox/
          2020.06.19 16.41 - RE ORIGINAL.../         <- email folder
            2020.06.19 16.41 - RE ORIGINAL....pdf    <- email body
            metadata.json
            attachments/
              invoice.pdf
              contract.docx
      Custodian B/
        2021.01.04 09.12 - FW Draft.../              <- email folder
          ...

    Matter_combined/                                 <- created automatically
      Custodian A/
        Inbox/
          2020.06.19 16.41 - RE ORIGINAL....pdf      <- email + attachments merged
      Custodian B/
        2021.01.04 09.12 - FW Draft....pdf
      errors.txt                                     <- written on any failure

An email folder is identified as follows:
  1. It contains {folder name}.pdf, or
  2. It contains metadata.json and exactly one PDF at its top level.
A folder with metadata.json but zero or several top-level PDFs (none matching
the folder name) is flagged in errors.txt and skipped.

Merge order: email PDF first, then attachments in filename order.
Each attachment gets a named PDF bookmark pointing to its first page.

If any attachment in an email cannot be converted the entire email is skipped
and the failure is flagged to the console and appended to errors.txt.

Conversion support
------------------
  .pdf                     merged directly (pikepdf)
  .docx .doc .rtf .odt     Microsoft Word (pywin32 on Windows, docx2pdf on
                           macOS) -> LibreOffice fallback
  .xlsx .xls .pptx .ppt    LibreOffice headless
  .jpg .jpeg .png .gif
  .bmp .tiff .tif .webp    img2pdf (lossless, preserves DPI)
  .txt .csv                Playwright (rendered as preformatted text)
  .html .htm               Playwright
  .eml .msg                emailtopdf pipeline (recursive)
  anything else            unsupported -> error, email skipped

Usage:
    python merge_email.py <folder name>

Requirements on Windows: pip install pikepdf img2pdf pywin32
LibreOffice (optional fallback) is found on PATH, in its default install
folder, or via the LIBREOFFICE_PATH environment variable.
"""

import os
import sys
import shutil
import tempfile
import subprocess
import traceback
import zipfile
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

OFFICE_WORD_EXTS  = {".docx", ".doc", ".rtf", ".odt"}
OFFICE_OTHER_EXTS = {".xlsx", ".xls", ".pptx", ".ppt", ".ods", ".odp"}
IMAGE_EXTS        = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
TEXT_EXTS         = {".txt", ".csv"}
HTML_EXTS         = {".html", ".htm"}
EMAIL_EXTS        = {".eml", ".msg"}
ZIP_EXTS          = {".zip"}

_ZIP_MAX_DEPTH        = 3          # maximum nesting depth for zips-within-zips
_ZIP_MAX_UNCOMPRESSED = 500 * 1024 * 1024   # 500 MB uncompressed size guard


# ---------------------------------------------------------------------------
# Individual converters
# ---------------------------------------------------------------------------

def _find_libreoffice() -> str:
    """Locate soffice: LIBREOFFICE_PATH env var, then PATH, then default install folders."""
    env = os.environ.get("LIBREOFFICE_PATH")
    if env and Path(env).is_file():
        return env
    for exe in ("soffice", "libreoffice"):
        found = shutil.which(exe)
        if found:
            return found
    candidates = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "LibreOffice/program/soffice.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "LibreOffice/program/soffice.exe",
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]
    for c in candidates:
        if c.is_file():
            return str(c)
    raise RuntimeError(
        "LibreOffice not found (checked LIBREOFFICE_PATH, PATH and the default install folder)"
    )


def _short_copy(src: Path, out_dir: Path, stem: str) -> Path:
    """
    Copy src to a short path inside out_dir by rewriting its bytes.
    This avoids Windows 260-character path limits on deeply nested email folders,
    and drops the 'downloaded from the internet' mark that makes Word open
    attachments in Protected View (which breaks automation).
    """
    dst = out_dir / f"{stem}{src.suffix.lower()}"
    dst.write_bytes(src.read_bytes())
    return dst


def _libreoffice_to_pdf(src: Path, out_dir: Path) -> Path:
    """Convert any LibreOffice-compatible file to PDF via headless LibreOffice."""
    exe = _find_libreoffice()
    lo_dir = out_dir / "lo"
    lo_dir.mkdir(exist_ok=True)
    work = _short_copy(src, lo_dir, "lo_src")
    result = subprocess.run(
        [exe, "--headless", "--convert-to", "pdf", "--outdir", str(lo_dir), str(work)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice exited {result.returncode}: {result.stderr.strip()}")
    pdf = lo_dir / f"{work.stem}.pdf"
    if not pdf.exists():
        raise RuntimeError(f"LibreOffice did not produce a PDF for {src.name}")
    return pdf


def _word_com_to_pdf(src: Path, out_dir: Path) -> Path:
    """
    Convert a Word-compatible file to PDF by driving Microsoft Word directly
    (Windows only). Replaces docx2pdf, which fails with 'Open.SaveAs' whenever
    Word cannot open the file (Protected View, long path, password, corruption)
    and hides the real reason.
    """
    try:
        import pythoncom
        import win32com.client
        import win32com.client.dynamic
    except ImportError:
        raise RuntimeError("pywin32 not installed: pip install pywin32")

    word_dir = out_dir / "word"
    word_dir.mkdir(exist_ok=True)
    work = _short_copy(src, word_dir, "word_src")
    out  = word_dir / "word_src.pdf"

    pythoncom.CoInitialize()
    word = None
    doc  = None
    try:
        # Separate, invisible Word instance; late-bound so a stale gen_py cache cannot interfere
        word = win32com.client.dynamic.Dispatch(
            win32com.client.DispatchEx("Word.Application")._oleobj_
        )
        word.Visible = False
        word.DisplayAlerts = 0                     # wdAlertsNone: no dialogs
        try:
            word.AutomationSecurity = 3            # msoAutomationSecurityForceDisable: no macros
        except Exception:
            pass

        doc = word.Documents.Open(
            FileName=str(work),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            PasswordDocument="__no_password__",    # password-protected files fail fast, no prompt
            Visible=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        if doc is None:
            raise RuntimeError("Word did not return an open document (Protected View or blocked file)")

        doc.ExportAsFixedFormat(str(out), 17)      # wdExportFormatPDF

    except pythoncom.com_error as exc:
        info = exc.excepinfo[2] if exc.excepinfo and exc.excepinfo[2] else str(exc)
        hint = ""
        if "password" in str(info).lower():
            hint = " (file is password-protected)"
        raise RuntimeError(f"Word could not convert the file: {info.strip()}{hint}")
    finally:
        if doc is not None:
            try:
                doc.Close(0)                       # wdDoNotSaveChanges
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(0)
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not out.exists():
        raise RuntimeError("Word finished but no PDF was produced")
    return out


def _docx2pdf_to_pdf(src: Path, out_dir: Path) -> Path:
    """Convert a Word-compatible file to PDF via docx2pdf (macOS, uses Word)."""
    try:
        from docx2pdf import convert
    except ImportError:
        raise RuntimeError("docx2pdf not installed: pip install docx2pdf")
    d2p_dir = out_dir / "d2p"
    d2p_dir.mkdir(exist_ok=True)
    work = _short_copy(src, d2p_dir, "d2p_src")
    out  = d2p_dir / "d2p_src.pdf"
    convert(str(work), str(out))
    if not out.exists():
        raise RuntimeError("docx2pdf did not produce a PDF")
    return out


def _word_to_pdf(src: Path, out_dir: Path) -> Path:
    """Microsoft Word first (COM on Windows, docx2pdf on macOS), then LibreOffice."""
    primary = _word_com_to_pdf if sys.platform == "win32" else _docx2pdf_to_pdf
    try:
        return primary(src, out_dir)
    except Exception as first:
        try:
            return _libreoffice_to_pdf(src, out_dir)
        except Exception as fallback:
            raise RuntimeError(f"{first}; LibreOffice fallback also failed: {fallback}")


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

    if ext == ".pdf":
        return src  # no conversion needed; caller uses the original

    if ext in OFFICE_WORD_EXTS:
        return _word_to_pdf(src, out_dir)

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
# Zip expansion
# ---------------------------------------------------------------------------

def _expand_zip(src: Path, out_dir: Path, depth: int = 0) -> list:
    """
    Extract a zip and convert its contents to PDFs.

    Returns a list of merge-part tuples: (label, pdf_path_or_None, children).
      - Regular file  → (filename, pdf_path, [])
      - Nested zip    → (zipname,  None,     [child_parts …])

    Raises RuntimeError on unreadable / oversized / too-deeply-nested zips,
    or if any contained file cannot be converted.
    """
    if depth > _ZIP_MAX_DEPTH:
        raise RuntimeError(
            f"Zip nesting exceeds maximum depth ({_ZIP_MAX_DEPTH}): {src.name}"
        )

    try:
        with zipfile.ZipFile(src) as zf:
            total_bytes = sum(info.file_size for info in zf.infolist())
            if total_bytes > _ZIP_MAX_UNCOMPRESSED:
                raise RuntimeError(
                    f"Zip uncompressed size ({total_bytes // (1024 * 1024)} MB) "
                    f"exceeds {_ZIP_MAX_UNCOMPRESSED // (1024 * 1024)} MB safety limit: {src.name}"
                )
            extract_dir = out_dir / "extracted"
            extract_dir.mkdir()
            zf.extractall(extract_dir)
    except zipfile.BadZipFile as exc:
        raise RuntimeError(f"Cannot read zip '{src.name}': {exc}")

    # Walk contents sorted so order is deterministic
    all_files = sorted(f for f in extract_dir.rglob("*") if f.is_file())
    if not all_files:
        raise RuntimeError(f"Zip '{src.name}' is empty or contains no files")

    parts = []
    for idx, item in enumerate(all_files):
        label    = str(item.relative_to(extract_dir))   # preserves subdir paths
        item_out = out_dir / f"item_{idx}"
        item_out.mkdir()

        if item.suffix.lower() in ZIP_EXTS:
            children = _expand_zip(item, item_out, depth + 1)
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

METADATA_NAME = "metadata.json"


def _email_pdf_for(folder: Path, filenames: list):
    """
    Decide whether folder is an email folder.

    Returns (email_pdf_path, error_message):
      (Path, None)  -> email folder; Path is the email body PDF
      (None, str)   -> looks like an email folder (has metadata.json) but the
                       email PDF cannot be identified
      (None, None)  -> not an email folder; keep walking
    """
    names_lower = {n.lower(): n for n in filenames}

    own_pdf = names_lower.get(f"{folder.name}.pdf".lower())
    if own_pdf:
        return folder / own_pdf, None

    if METADATA_NAME not in names_lower:
        return None, None

    pdfs = sorted(n for n in filenames if n.lower().endswith(".pdf"))
    if len(pdfs) == 1:
        return folder / pdfs[0], None
    if not pdfs:
        return None, "has metadata.json but no email PDF"
    return None, (
        f"has metadata.json and {len(pdfs)} PDFs, none named "
        f"'{folder.name}.pdf'; cannot tell which is the email"
    )


def _find_email_folders(source_dir: Path, exclude: Path, errors: list) -> list:
    """
    Walk source_dir recursively and return [(email_folder, email_pdf), ...].

    Descent stops at each email folder, so its attachments/ subfolder (or
    anything else inside it) is never treated as a separate email.
    The exclude directory (the _combined output tree) is never walked.
    """
    found = []
    for root, dirs, files in os.walk(source_dir):
        root_path = Path(root)
        dirs.sort()

        # Never walk into the output tree if it sits inside the source
        dirs[:] = [d for d in dirs if (root_path / d).resolve() != exclude]

        if root_path == source_dir:
            continue  # the top-level folder itself is not an email

        email_pdf, problem = _email_pdf_for(root_path, files)
        if email_pdf is not None:
            found.append((root_path, email_pdf))
            dirs[:] = []          # stop here: this is the email level
        elif problem:
            rel = root_path.relative_to(source_dir)
            errors.append(f"[{rel}] Skipped: {problem}")
            print(f"[SKIP]   {rel}: {problem}")
            dirs[:] = []
    return found


def process_email_folder(
    email_folder: Path,
    email_pdf: Path,
    out_path: Path,
    errors: list,
    label: str,
) -> bool:
    """
    Produce a combined PDF for one email folder at out_path.
    Returns True on success, False if any conversion failed (error appended to errors).
    """
    attachments_dir = email_folder / "attachments"
    attachments    = sorted(
        f for f in attachments_dir.iterdir() if f.is_file()
    ) if attachments_dir.exists() else []

    out_path.parent.mkdir(parents=True, exist_ok=True)

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
                if att.suffix.lower() in ZIP_EXTS:
                    children = _expand_zip(att, att_tmp)
                    parts.append((att.name, None, children))
                    n_files = sum(1 for _, p, _ in children if p is not None)
                    print(f"  [OK]     {att.name}  ({n_files} file(s) inside)")
                else:
                    pdf = convert_to_pdf(att, att_tmp)
                    parts.append((att.name, pdf, []))
                    print(f"  [OK]     {att.name}")
            except Exception as exc:
                msg = (
                    f"[{label}] "
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
    if len(sys.argv) != 2:
        print("Usage: python merge_email.py <folder name>", file=sys.stderr)
        sys.exit(2)

    source_dir = Path(sys.argv[1]).expanduser().resolve()
    if not source_dir.is_dir():
        print(f"Error: directory not found: {source_dir}", file=sys.stderr)
        sys.exit(1)

    combined_dir = source_dir.parent / f"{source_dir.name}_combined"
    combined_dir.mkdir(parents=True, exist_ok=True)
    errors_txt = combined_dir / "errors.txt"

    print(f"Source : {source_dir}")
    print(f"Dest   : {combined_dir}")
    print()

    errors  = []
    folders = _find_email_folders(source_dir, combined_dir.resolve(), errors)
    skipped = len(errors)

    if not folders:
        print("No email folders found.")

    succeeded = 0
    failed    = 0

    for folder, email_pdf in folders:
        rel      = folder.relative_to(source_dir)
        # Mirror the source structure: <parent subfolders>/<email folder name>.pdf
        out_path = combined_dir / rel.parent / f"{folder.name}.pdf"
        print(f"[{rel}]")
        try:
            ok = process_email_folder(folder, email_pdf, out_path, errors, str(rel))
        except Exception as exc:
            errors.append(f"[{rel}] Unexpected error: {exc}")
            print(f"  [ERROR]  {exc}")
            ok = False
        if ok:
            succeeded += 1
        else:
            failed += 1
        print()

    print(f"Done: {succeeded} merged, {failed} failed, {skipped} skipped.")

    if errors:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(errors_txt, "a", encoding="utf-8") as f:
            f.write(f"\n=== {timestamp} ===\n")
            f.write(f"Source: {source_dir}\n")
            for msg in errors:
                f.write(f"{msg}\n")
        print(f"\n{len(errors)} issue(s) logged to: {errors_txt}")


if __name__ == "__main__":
    main()
