# emailtopdf

Convert `.eml` and `.msg` email files to PDF, with attachments extracted to a per-email folder. Optionally merge each email and its attachments into a single combined PDF.

Renders via **headless Chromium** (Playwright) to preserve the original email's fonts, colours, and layout. Handles both modern HTML emails and legacy Outlook `.msg` files. All processing is fully local — no email content or metadata is sent to the internet.

---

## Scripts

| Script | Purpose |
|--------|---------|
| `emailtopdf.py` | Convert `.eml`/`.msg` files to PDF and extract attachments |
| `collect_pdfs.py` | Copy email PDFs into a flat folder (attachment PDFs excluded) |
| `merge_email.py` | Merge each email PDF with its attachments into one combined PDF |

---

## Features

### `emailtopdf.py`

- Converts `.eml` (RFC 2822) and `.msg` (Outlook) files
- Preserves original formatting — fonts, colour, spacing, inline images
- Injects a clean header block (From / To / CC / BCC / Date / Subject / Attachments) above the email body
- Strips Outlook/Word CSS artefacts that would break PDF layout (`@page`, VML, MSO conditionals)
- Rescues VML background images from Outlook conditional comments and renders them in the non-Outlook fallback
- Inlines CSS via `css-inline` for better rendering consistency
- Blocks all external network requests during rendering — no tracking pixel loads, no data sent to the internet
- Reuses one Chromium instance across the full batch (~3–5× faster than spawning per email)
- Extracts non-inline attachments to an `attachments/` subfolder, including embedded `.msg` files
- Saves `metadata.json` per email (headers + attachment list)
- Injects XMP metadata into the PDF (title, author, date, recipients) via pikepdf
- Optional PDF/A-2b output via Ghostscript (`--pdfa`)
- WeasyPrint fallback renderer (`--renderer weasyprint`) — external URLs blocked

### `collect_pdfs.py`

Copies all email PDFs from the output folder into a flat `_emails_only/` sibling folder. Attachment PDFs (inside `attachments/` subfolders) are excluded. Name conflicts resolved automatically with a `_2`, `_3` suffix.

### `merge_email.py`

Merges each email's PDF and all its attachments into a single combined PDF, with named bookmarks for each attachment. Emails with no attachments are copied as-is.

Supported attachment types:

| Type | Method |
|------|--------|
| `.pdf` | Merged directly (pikepdf) |
| `.docx` `.doc` `.rtf` `.odt` | docx2pdf (Word COM) → LibreOffice fallback |
| `.xlsx` `.xls` `.pptx` `.ppt` `.ods` `.odp` | LibreOffice headless |
| `.jpg` `.jpeg` `.png` `.gif` `.bmp` `.tiff` `.webp` | img2pdf (lossless, preserves DPI) |
| `.txt` `.csv` | Playwright (preformatted text) |
| `.html` `.htm` | Playwright |
| `.eml` `.msg` | emailtopdf pipeline (recursive) |
| `.zip` | Extracted, contents converted; nested bookmarks |
| `.rar` | Extracted via rarfile + UnRAR binary; nested bookmarks |

Extensionless attachments are identified by magic byte signatures before conversion. Any attachment conversion failure causes that email to be skipped and logged to `errors.txt`.

---

## Output structure

```
output/
  YYYY.MM.DD HH.MM - Subject/
    YYYY.MM.DD HH.MM - Subject.pdf
    metadata.json
    attachments/
      report.xlsx
      ...
```

Folders and PDFs use a date+time prefix so they sort chronologically. When `-o` is omitted, the output folder is placed **next to** (not inside) the input folder:

```
emails/            <- input folder
emails_output/     <- created automatically next to it
  YYYY.MM.DD .../
```

---

## Requirements

- Python 3.9+
- Chromium (installed via Playwright)

```bash
pip install -r requirements.txt
playwright install chromium
```

### Optional dependencies

| Feature | Package / tool |
|---------|---------------|
| PDF/A-2b output (`--pdfa`) | [Ghostscript](https://www.ghostscript.com/) — `gs` or `gswin64c` in PATH |
| XMP metadata | `pikepdf` (in `requirements.txt`) |
| CSS inlining | `css-inline` (in `requirements.txt`) |
| WeasyPrint renderer | `pip install weasyprint` + GTK/Pango (Linux/macOS) |
| Merge Word attachments | `pip install docx2pdf` + Microsoft Word installed |
| Merge image attachments | `pip install img2pdf` |
| Merge RAR archives | `pip install rarfile` + WinRAR or UnRAR in PATH |

---

## Usage

### Convert emails to PDF

```bash
# Single file — output goes next to the input folder
python emailtopdf.py path/to/email.msg

# Multiple files
python emailtopdf.py file1.eml file2.msg

# Entire folder
python emailtopdf.py path/to/emails/

# Explicit output directory
python emailtopdf.py path/to/emails/ -o path/to/output/

# PDF/A-2b archival output (requires Ghostscript in PATH)
python emailtopdf.py path/to/emails/ --pdfa
```

### Collect email PDFs into a flat folder

```bash
python collect_pdfs.py                    # uses output/ next to the script
python collect_pdfs.py path/to/output/    # explicit path
```

Creates `output_emails_only/` next to the output folder containing a flat list of email PDFs.

### Merge emails with their attachments

```bash
python merge_email.py                     # uses output/ next to the script
python merge_email.py path/to/output/     # explicit path
```

Creates `output_combined/` next to the output folder. Each email's PDF and all its attachments are merged into a single PDF with a named bookmark outline.

---

## How it works

### HTML transform pipeline

Each email body passes through these steps before rendering:

1. **CID resolution** — `src="cid:xxx"` inline image references replaced with base64 data URIs so all images are self-contained
2. **VML/MSO stripping** — Outlook conditional comments processed:
   - Paired MSO + non-MSO blocks: VML background images extracted and injected as CSS `background-image` onto the non-MSO fallback; MSO block removed
   - Unpaired MSO-only blocks: stripped entirely
   - Non-MSO unwrap blocks: comment markers removed, content kept
3. **CSS sanitization** — `@page`, `page:`, `size:`, `mso-*`, `behavior:`, `panose-1:` stripped from `<style>` blocks (Outlook/Word artefacts that force unwanted Chromium page breaks)
4. **Template assembly** — email header block injected above the body; wrapped in a complete HTML document
5. **CSS inlining** — `<style>` rules inlined into element `style=""` attributes via `css-inline`

### Rendering

- Playwright (Chromium), `print_background=True`, A4 format
- All outbound network aborted before any request leaves the process
- One Chromium instance reused across the full batch

### Post-processing

- Optional PDF/A-2b via Ghostscript (`--pdfa`)
- XMP metadata injected via pikepdf: `dc:title` (subject), `dc:creator` (from), `xmp:CreateDate`, `dc:description` (recipients), `dc:identifier` (Message-ID)

---

## Known limitations

- **VML backgrounds**: Outlook 2007–2019 VML backgrounds (`<v:rect>`, `<v:fill>`) are not rendered by Chromium. The script rescues them where the standard paired MSO/non-MSO pattern is used; non-standard VML may be lost.
- **MSO conditional layout**: Table-based layouts inside `<!--[if mso]>` blocks are stripped. The non-Outlook fallback layout is used instead.
- **PDF/A compliance**: `--pdfa` is best-effort. Validate with [veraPDF](https://verapdf.org/) for strict archival use.
- **WeasyPrint on Windows**: Requires GTK+/Pango system libraries. Not recommended on Windows; use the default Playwright renderer.

---

## Renderer comparison

| | Playwright (default) | WeasyPrint |
|---|---|---|
| CSS fidelity | Excellent (Chromium engine) | Good |
| Inline images | Yes | Yes |
| Background images | Yes | Partial |
| Windows support | Yes | Difficult (GTK dependency) |
| Requires browser install | Yes | No |
| Speed — single email | ~2 s | ~1 s |
| Speed — batch of 100 | ~40 s (shared pool) | ~100 s |

---

## Security

All rendering is fully local. No email content, metadata, headers, or attachment data is sent to any external service.

### Network isolation

**Playwright (default renderer)**
`page.route("**/*", abort)` is applied to every page before any content loads. All outbound connections — tracking pixels, remote images, linked stylesheets, web fonts — are aborted at the browser level before leaving the process.

**WeasyPrint renderer**
A `url_fetcher` that raises `URLFetchingError` is passed to every `HTML()` call. Any URL WeasyPrint would normally fetch is blocked immediately.

**css-inline**
`load_remote_stylesheets=False` is set explicitly.

### Security fix log

| Date | Severity | Component | Description |
|------|----------|-----------|-------------|
| 2026-04-12 | High | `emailtopdf.py` | WeasyPrint renderer made unrestricted external network requests. Email HTML bodies routinely contain tracking pixels and remote resources. Fixed by supplying a `url_fetcher` that raises `URLFetchingError` for every URL, blocking all outbound connections. |
| 2026-04-12 | Medium | `merge_email.py` | ZIP and RAR extraction used `extractall()` without member-path validation. A maliciously crafted archive attachment could write files outside the extraction directory (zip-slip / path traversal). Fixed by resolving each member's destination path and verifying it falls within the extraction root before extracting. |
