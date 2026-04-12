#!/usr/bin/env python3
"""
emailtopdf - Convert .eml and .msg email files to PDF with attachment extraction.

Usage:
    python emailtopdf.py file.eml file.msg          # convert individual files
    python emailtopdf.py path/to/emails/              # convert entire folder
    python emailtopdf.py file.eml --renderer weasyprint
    python emailtopdf.py path/to/emails/ --pdfa      # PDF/A-2b output (needs Ghostscript)
"""

import email
import email.header
import email.policy
import email.utils
import re
import base64
import json
import mimetypes
import sys
import argparse
import subprocess
import shutil
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime
from typing import Optional
from html import escape

# Reconfigure stdout/stderr to UTF-8 so that subjects or paths containing
# non-Latin characters (e.g. Chinese, Arabic) don't crash print() on
# Windows consoles that default to CP1252.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Attachment:
    filename: str
    data: bytes
    content_type: str
    is_inline: bool = False
    cid: Optional[str] = None   # stripped of angle brackets


@dataclass
class EmailData:
    subject: str
    from_: str
    to: list
    cc: list
    bcc: list
    date: Optional[datetime]
    html_body: Optional[str]
    text_body: Optional[str]
    attachments: list           # list[Attachment]
    message_id: str
    source_file: str


# ---------------------------------------------------------------------------
# Header decoding helpers
# ---------------------------------------------------------------------------

def _decode_header(value: Optional[str]) -> str:
    """Decode an RFC 2047 encoded email header value."""
    if not value:
        return ""
    decoded_parts = email.header.decode_header(value)
    parts = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            parts.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            parts.append(part)
    return "".join(parts)


def _parse_address_list(value: Optional[str]) -> list:
    """Split a comma- or semicolon-separated list of email addresses."""
    if not value:
        return []
    value = value.replace(";", ",")
    return [addr.strip() for addr in value.split(",") if addr.strip()]


def _format_date(date: Optional[datetime]) -> str:
    if not date:
        return "Unknown"
    try:
        return date.strftime("%d %B %Y, %H:%M %Z").strip()
    except Exception:
        return str(date)


# ---------------------------------------------------------------------------
# .eml parser
# ---------------------------------------------------------------------------

def parse_eml(path: Path) -> EmailData:
    """Parse an .eml file and return EmailData."""
    raw = path.read_bytes()
    msg = email.message_from_bytes(raw, policy=email.policy.compat32)

    subject    = _decode_header(msg.get("Subject", ""))
    from_      = _decode_header(msg.get("From", ""))
    to         = _parse_address_list(_decode_header(msg.get("To", "")))
    cc         = _parse_address_list(_decode_header(msg.get("CC", "")))
    bcc        = _parse_address_list(_decode_header(msg.get("BCC", "")))
    message_id = msg.get("Message-ID", "")

    date = None
    date_str = msg.get("Date", "")
    if date_str:
        try:
            date = email.utils.parsedate_to_datetime(date_str)
        except Exception:
            pass

    html_body   = None
    text_body   = None
    attachments = []

    for part in msg.walk():
        content_type  = part.get_content_type()
        disposition   = str(part.get("Content-Disposition", ""))
        content_id    = part.get("Content-ID", "").strip("<>")
        is_attachment = "attachment" in disposition.lower()

        if content_type == "text/html" and not is_attachment and html_body is None:
            charset = part.get_content_charset() or "utf-8"
            payload = part.get_payload(decode=True)
            if payload:
                html_body = payload.decode(charset, errors="replace")

        elif content_type == "text/plain" and not is_attachment and text_body is None:
            charset = part.get_content_charset() or "utf-8"
            payload = part.get_payload(decode=True)
            if payload:
                text_body = payload.decode(charset, errors="replace")

        elif content_type.startswith("image/") and content_id and not is_attachment:
            payload = part.get_payload(decode=True)
            if payload:
                ext   = content_type.split("/")[-1].split(";")[0]
                fname = part.get_filename() or f"{content_id}.{ext}"
                fname = _decode_header(fname)
                attachments.append(Attachment(
                    filename=fname, data=payload,
                    content_type=content_type, is_inline=True, cid=content_id,
                ))

        elif is_attachment:
            payload = part.get_payload(decode=True)
            fname   = part.get_filename() or "attachment"
            fname   = _decode_header(fname)
            if payload:
                attachments.append(Attachment(
                    filename=fname, data=payload,
                    content_type=content_type, is_inline=False,
                    cid=content_id or None,
                ))

    return EmailData(
        subject=subject, from_=from_, to=to, cc=cc, bcc=bcc,
        date=date, html_body=html_body, text_body=text_body,
        attachments=attachments, message_id=message_id,
        source_file=str(path),
    )


# ---------------------------------------------------------------------------
# .msg parser
# ---------------------------------------------------------------------------

def parse_msg(path: Path) -> EmailData:
    """Parse an Outlook .msg file and return EmailData."""
    try:
        import extract_msg
    except ImportError:
        raise ImportError(
            "extract-msg is required for .msg files: pip install extract-msg"
        )

    msg = extract_msg.openMsg(str(path))

    subject    = msg.subject or ""
    from_      = msg.sender or ""
    to         = _parse_address_list(msg.to or "")
    cc         = _parse_address_list(msg.cc or "")
    bcc        = _parse_address_list(getattr(msg, "bcc", None) or "")
    message_id = getattr(msg, "messageId", "") or ""

    date = None
    raw_date = msg.date
    if raw_date:
        if isinstance(raw_date, datetime):
            date = raw_date
        else:
            try:
                date = email.utils.parsedate_to_datetime(str(raw_date))
            except Exception:
                pass

    html_body = None
    text_body = None

    try:
        raw_html = msg.htmlBody
        if raw_html:
            if isinstance(raw_html, bytes):
                # Sniff the charset declared in the HTML meta tag before decoding.
                # MSG htmlBody bytes can be Windows-1252 (or other legacy charsets);
                # blindly decoding as UTF-8 turns \x93/\x94 (curly quotes) into ?
                m = re.search(rb'charset=["\']?([^"\';\s>]+)', raw_html[:2048])
                charset = m.group(1).decode("ascii", errors="replace") if m else "utf-8"
                html_body = raw_html.decode(charset, errors="replace")
            else:
                html_body = raw_html
    except Exception:
        pass  # RTF-to-HTML conversion can fail on malformed/non-standard charsets

    try:
        if msg.body:
            text_body = msg.body
    except Exception:
        pass  # same RTF decoding issue can affect plain-text body extraction

    # If htmlBody is a near-empty stub (e.g. '<html><body><br /></body></head>' generated
    # by extract-msg for RTF-only emails), treat it as absent so we fall through to RTF.
    if html_body:
        _visible = re.sub(r'<[^>]+>', '', html_body)
        _visible = re.sub(r'&[a-zA-Z]+;|&#\d+;', ' ', _visible)
        if not _visible.strip():
            html_body = None

    if text_body and not text_body.strip():
        text_body = None

    # RTF fallback — used when extract-msg cannot produce usable HTML or plain text.
    # This happens for RTF-only MSG files where htmlBody is a stub and body is whitespace.
    if not html_body and not text_body:
        try:
            raw_rtf = getattr(msg, "rtfBody", None)
            if raw_rtf:
                from striprtf.striprtf import rtf_to_text
                rtf_str = raw_rtf.decode("cp1252", errors="replace")
                rtf_text = rtf_to_text(rtf_str)
                if rtf_text.strip():
                    text_body = rtf_text
        except Exception:
            pass

    attachments = []
    for att in msg.attachments:
        data = getattr(att, "data", None)
        if data is None:
            continue

        # Embedded .msg attachments: data is a Message object, not bytes.
        # Use exportBytes() to get the raw OLE2 compound document bytes.
        if not isinstance(data, bytes):
            try:
                data = data.exportBytes()
            except Exception:
                continue
            if not isinstance(data, bytes) or not data:
                continue

        fname = (
            att.getFilename()
            or getattr(att, "longFilename", None)
            or getattr(att, "shortFilename", None)
            or "attachment"
        )
        cid_raw = getattr(att, "cid", None) or ""
        cid     = cid_raw.strip("<>") if cid_raw else None

        content_type = (
            getattr(att, "mimetype", None)
            or mimetypes.guess_type(fname)[0]
            or "application/octet-stream"
        )

        is_inline = bool(cid and html_body and f"cid:{cid}" in html_body)

        attachments.append(Attachment(
            filename=fname, data=data,
            content_type=content_type, is_inline=is_inline, cid=cid,
        ))

    msg.close()  # release the OLE2 file handle; without this Windows keeps a lock

    return EmailData(
        subject=subject, from_=from_, to=to, cc=cc, bcc=bcc,
        date=date, html_body=html_body, text_body=text_body,
        attachments=attachments, message_id=message_id,
        source_file=str(path),
    )


# ---------------------------------------------------------------------------
# CID image resolution
# ---------------------------------------------------------------------------

def _build_cid_map(attachments: list) -> dict:
    """Build { cid -> data:... URI } for all inline images."""
    cid_map = {}
    for att in attachments:
        if att.is_inline and att.cid and att.data:
            b64 = base64.b64encode(att.data).decode("ascii")
            ct  = att.content_type or "image/png"
            cid_map[att.cid] = f"data:{ct};base64,{b64}"
    return cid_map


def _resolve_cid_images(html: str, cid_map: dict) -> str:
    """Replace src="cid:xxx" / src='cid:xxx' with base64 data URIs."""
    if not cid_map:
        return html

    def replacer(match):
        quote    = match.group(1)
        cid      = match.group(2).strip()
        data_uri = cid_map.get(cid)
        if data_uri:
            return f"src={quote}{data_uri}{quote}"
        return match.group(0)

    return re.sub(
        r'src=(["\'])cid:([^"\'>\s]+)\1',
        replacer,
        html,
        flags=re.IGNORECASE,
    )


# ---------------------------------------------------------------------------
# VML / MSO conditional comment handling
# ---------------------------------------------------------------------------

# Paired: MSO block (VML etc.) immediately followed by non-MSO fallback block.
# [^\]!]* ensures we don't match <!--[if !mso]> (the ! would break [^\]!]*).
_RE_PAIRED_MSO = re.compile(
    r'<!--\[if [^\]!]*mso[^\]]*\]>(.*?)<!\[endif\]-->'
    r'(\s*)'
    r'<!--\[if !mso\]><!-->(.*?)<!--<!\[endif\]-->',
    re.DOTALL | re.IGNORECASE,
)
# Remaining unpaired MSO-only blocks (strip entirely)
_RE_MSO_ONLY = re.compile(
    r'<!--\[if [^\]!]*mso[^\]]*\]>.*?<!\[endif\]-->',
    re.DOTALL | re.IGNORECASE,
)
# Remaining non-MSO unwrap blocks (keep inner content, remove comment markers)
_RE_NON_MSO_UNWRAP = re.compile(
    r'<!--\[if !mso\]><!-->(.*?)<!--<!\[endif\]-->',
    re.DOTALL | re.IGNORECASE,
)
# VML fill element with src attribute
_RE_VML_FILL_SRC = re.compile(
    r'<v:fill\b[^>]*\bsrc=["\']([^"\']+)["\']',
    re.IGNORECASE,
)


def _inject_background_style(html_fragment: str, img_src: str) -> str:
    """Inject background-image CSS on the first block element in the fragment."""
    bg_css = (
        f"background-image:url('{img_src}');"
        "background-size:cover;"
        "background-repeat:no-repeat;"
        "background-position:center;"
    )
    injected = [False]

    def inject(m):
        if injected[0]:
            return m.group(0)
        tag   = m.group(1)
        attrs = m.group(2)
        style_m = re.search(r'\bstyle=["\']([^"\']*)["\']', attrs, re.IGNORECASE)
        if style_m:
            existing  = style_m.group(1).rstrip(";")
            new_attrs = (
                attrs[: style_m.start()]
                + f'style="{existing};{bg_css}"'
                + attrs[style_m.end():]
            )
        else:
            new_attrs = attrs + f' style="{bg_css}"'
        injected[0] = True
        return f"<{tag}{new_attrs}>"

    return re.sub(
        r"<(div|table|td|p|section|span|article)\b([^>]*)>",
        inject,
        html_fragment,
        flags=re.IGNORECASE,
    )


def _strip_vml_and_mso_blocks(html: str, cid_map: dict) -> str:
    """
    Process Outlook MSO conditional comments:

    - Paired MSO + non-MSO blocks: extract any VML fill image from the MSO block
      and inject it as background-image on the first element of the non-MSO
      fallback, then keep only the fallback content (visible to Chromium).
    - Remaining unpaired MSO-only blocks: strip.
    - Remaining non-MSO unwrap blocks: unwrap (keep inner content).
    """
    def _handle_pair(m: re.Match) -> str:
        mso_content = m.group(1)
        fallback    = m.group(3)
        vfill = _RE_VML_FILL_SRC.search(mso_content)
        if vfill:
            img_src = vfill.group(1).strip()
            if img_src.lower().startswith("cid:"):
                cid     = img_src[4:].strip()
                img_src = cid_map.get(cid, img_src)
            fallback = _inject_background_style(fallback, img_src)
        return fallback

    html = _RE_PAIRED_MSO.sub(_handle_pair, html)
    html = _RE_MSO_ONLY.sub("", html)
    html = _RE_NON_MSO_UNWRAP.sub(lambda m: m.group(1), html)
    return html


# ---------------------------------------------------------------------------
# CSS inlining
# ---------------------------------------------------------------------------

def _inline_css(html: str) -> str:
    """
    Inline all <style> block rules into element style="" attributes via css-inline.
    Non-inlineable rules (media queries, pseudo-selectors) are kept in <style>.
    Falls through gracefully if css-inline is not installed or raises.
    """
    try:
        import css_inline
        inliner = css_inline.CSSInliner(
            remove_style_tags=False,
            load_remote_stylesheets=False,
        )
        return inliner.inline(html)
    except ImportError:
        return html
    except Exception:
        return html


# ---------------------------------------------------------------------------
# HTML document builder
# ---------------------------------------------------------------------------

def _sanitize_email_styles(styles_html: str) -> str:
    """
    Strip CSS that conflicts with our PDF layout or is Outlook/IE-only:
      @page rules       — override our Playwright margins/page size
      page: property    — triggers Chromium named-page forced breaks
      size: property    — pairs with @page; overrides page format
      mso-* properties  — Outlook-only, unrecognised by Chromium
      behavior:         — IE/VML only
      panose-1:         — font metadata, not a rendering property
    """
    def _clean(m: re.Match) -> str:
        block = m.group(0)
        block = re.sub(r"@page\b[^{]*\{[^}]*\}", "", block, flags=re.IGNORECASE | re.DOTALL)
        block = re.sub(r"\bpage\s*:\s*[^;}\n]+;?",      "", block, flags=re.IGNORECASE)
        block = re.sub(r"\bsize\s*:\s*[^;}\n]+;?",      "", block, flags=re.IGNORECASE)
        block = re.sub(r"\bmso-[a-z-]+\s*:\s*[^;}\n]+;?", "", block, flags=re.IGNORECASE)
        block = re.sub(r"\bbehavior\s*:\s*[^;}\n]+;?",  "", block, flags=re.IGNORECASE)
        block = re.sub(r"\bpanose-1\s*:\s*[^;}\n]+;?",  "", block, flags=re.IGNORECASE)
        return block

    return re.sub(
        r"<style[^>]*>.*?</style>",
        _clean,
        styles_html,
        flags=re.IGNORECASE | re.DOTALL,
    )


def _extract_head_styles(html: str) -> str:
    """Extract <style> blocks from the email's <head>, sanitized for PDF rendering."""
    head_match = re.search(r"<head[^>]*>(.*?)</head>", html, re.IGNORECASE | re.DOTALL)
    if not head_match:
        return ""
    raw_styles = "\n".join(
        re.findall(r"<style[^>]*>.*?</style>", head_match.group(1), re.IGNORECASE | re.DOTALL)
    )
    return _sanitize_email_styles(raw_styles)


def _extract_body_content(html: str) -> str:
    """Return the content between <body> tags, preserving any body inline styles."""
    body_tag_match     = re.search(r"<body([^>]*)>",              html, re.IGNORECASE)
    body_content_match = re.search(r"<body[^>]*>(.*?)</body>", html, re.IGNORECASE | re.DOTALL)

    if not body_content_match:
        return html

    inner = body_content_match.group(1)

    if body_tag_match:
        attrs       = body_tag_match.group(1)
        style_match = re.search(r'style=["\']([^"\']*)["\']', attrs, re.IGNORECASE)
        if style_match:
            return f'<div style="{style_match.group(1)}">{inner}</div>'

    return inner


_HEADER_CSS = """
.etp-header {
  font-family: Arial, Helvetica, sans-serif;
  font-size: 12px;
  color: #333;
  margin-bottom: 16px;
  background-color: #ffffff;
  background-image: none;
}
.etp-subject {
  font-size: 18px;
  font-weight: bold;
  margin-bottom: 8px;
  color: #111;
}
.etp-meta { border-collapse: collapse; }
.etp-meta td { padding: 2px 8px 2px 0; vertical-align: top; }
.etp-meta .lbl { font-weight: bold; color: #555; white-space: nowrap; min-width: 55px; }
.etp-divider { border: none; border-top: 1px solid #ddd; margin: 12px 0 16px; }
.etp-att { margin: 0; padding-left: 16px; }
.etp-att li { padding: 1px 0; }
"""


def _build_header_block(data: EmailData) -> str:
    rows = [
        f'<tr><td class="lbl">From:</td><td>{escape(data.from_ or "")}</td></tr>',
        f'<tr><td class="lbl">To:</td><td>{escape(", ".join(data.to))}</td></tr>',
    ]
    if data.cc:
        rows.append(f'<tr><td class="lbl">CC:</td><td>{escape(", ".join(data.cc))}</td></tr>')
    if data.bcc:
        rows.append(f'<tr><td class="lbl">BCC:</td><td>{escape(", ".join(data.bcc))}</td></tr>')
    rows.append(f'<tr><td class="lbl">Date:</td><td>{escape(_format_date(data.date))}</td></tr>')

    non_inline = [a for a in data.attachments if not a.is_inline]
    if non_inline:
        items = "".join(
            f'<li>{escape(a.filename)}</li>' for a in non_inline
        )
        rows.append(
            f'<tr><td class="lbl">Attachments:</td>'
            f'<td><ul class="etp-att">{items}</ul></td></tr>'
        )

    return (
        f'<div class="etp-header">'
        f'<div class="etp-subject">{escape(data.subject or "(no subject)")}</div>'
        f'<table class="etp-meta">{"".join(rows)}</table>'
        f'<hr class="etp-divider">'
        f'</div>'
    )


def build_html(data: EmailData) -> str:
    """
    Assemble a complete, self-contained HTML document for PDF rendering.

    HTML body transform pipeline (in order):
      1. CID  -> base64 data URI (inline images)
      2. VML/MSO block stripping (rescue VML backgrounds into fallback divs)
      3. Head style extraction + CSS sanitization (@page, mso-*, etc.)
      4. Assemble final template with injected email header block

    CSS inlining (_inline_css) is applied by convert_email on the assembled doc.
    """
    cid_map      = _build_cid_map(data.attachments)
    header_block = _build_header_block(data)

    if data.html_body:
        processed    = _resolve_cid_images(data.html_body, cid_map)
        processed    = _strip_vml_and_mso_blocks(processed, cid_map)
        head_styles  = _extract_head_styles(processed)
        body_content = _extract_body_content(processed)

        return (
            f"<!DOCTYPE html>\n<html>\n<head>\n"
            f'<meta charset="UTF-8">\n'
            f"<style>{_HEADER_CSS}</style>\n"
            f"{head_styles}\n"
            f"</head>\n<body>\n"
            f"{header_block}\n"
            f'<div class="etp-body">{body_content}</div>\n'
            f"</body>\n</html>"
        )

    if data.text_body:
        return (
            f"<!DOCTYPE html>\n<html>\n<head>\n"
            f'<meta charset="UTF-8">\n'
            f"<style>{_HEADER_CSS}\n"
            f"pre {{ font-family: monospace; white-space: pre-wrap; "
            f"word-wrap: break-word; font-size: 13px; line-height: 1.4; }}\n"
            f"</style>\n</head>\n<body>\n"
            f"{header_block}\n"
            f"<pre>{escape(data.text_body)}</pre>\n"
            f"</body>\n</html>"
        )

    return (
        f'<!DOCTYPE html>\n<html>\n<head><meta charset="UTF-8">'
        f"<style>{_HEADER_CSS}</style></head>\n<body>\n"
        f"{header_block}\n<p><em>(No body content)</em></p>\n"
        f"</body>\n</html>"
    )


# ---------------------------------------------------------------------------
# Playwright pool
# ---------------------------------------------------------------------------

class PlaywrightPool:
    """
    Reusable Chromium browser instance for batch PDF rendering.

    Improvements over spawning a new browser per email:
    - 3-5x faster on batches (browser startup cost paid once)
    - All external network blocked via page.route (eliminates tracking pixel hangs)
    - wait_until="domcontentloaded" is safe since all resources are inlined data URIs

    Usage:
        with PlaywrightPool() as pool:
            pool.render_pdf(html, output_path)
    """

    def __init__(self) -> None:
        self._pw      = None
        self._browser = None

    def __enter__(self) -> "PlaywrightPool":
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise ImportError(
                "playwright is required:\n"
                "  pip install playwright\n"
                "  playwright install chromium"
            )
        self._pw      = sync_playwright().start()
        self._browser = self._pw.chromium.launch()
        return self

    def __exit__(self, *args) -> None:
        if self._browser:
            self._browser.close()
        if self._pw:
            self._pw.stop()

    def render_pdf(self, html: str, output_path: Path) -> None:
        """Render HTML -> PDF in a fresh page (reuses the shared browser instance)."""
        page = self._browser.new_page()
        try:
            page.route("**/*", lambda route: route.abort())
            page.set_content(html, wait_until="domcontentloaded")
            page.pdf(
                path=str(output_path),
                format="A4",
                print_background=True,
                margin={"top": "15mm", "bottom": "15mm", "left": "15mm", "right": "15mm"},
            )
        finally:
            page.close()


# ---------------------------------------------------------------------------
# PDF renderers
# ---------------------------------------------------------------------------

def render_pdf_playwright(
    html: str,
    output_path: Path,
    pool: Optional[PlaywrightPool] = None,
) -> None:
    """Render HTML -> PDF via headless Chromium. Uses pool if provided."""
    if pool is not None:
        pool.render_pdf(html, output_path)
        return

    # Standalone path (no pool) — same network-blocking improvements applied
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise ImportError(
            "playwright is required:\n"
            "  pip install playwright\n"
            "  playwright install chromium"
        )

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page    = browser.new_page()
        try:
            page.route("**/*", lambda route: route.abort())
            page.set_content(html, wait_until="domcontentloaded")
            page.pdf(
                path=str(output_path),
                format="A4",
                print_background=True,
                margin={"top": "15mm", "bottom": "15mm", "left": "15mm", "right": "15mm"},
            )
        finally:
            page.close()
            browser.close()


def render_pdf_weasyprint(html: str, output_path: Path) -> None:
    """Render HTML -> PDF via WeasyPrint (pure Python, no browser needed)."""
    try:
        from weasyprint import HTML
        from weasyprint.urls import URLFetchingError
    except ImportError:
        raise ImportError("weasyprint is required: pip install weasyprint")

    def _block_network(url, timeout=10):
        """Block all external URL fetches — emails may contain tracking pixels."""
        raise URLFetchingError(f"Blocked external URL: {url}")

    HTML(string=html, url_fetcher=_block_network).write_pdf(str(output_path))


RENDERERS = {
    "playwright": render_pdf_playwright,
    "weasyprint": render_pdf_weasyprint,
}


# ---------------------------------------------------------------------------
# XMP metadata injection
# ---------------------------------------------------------------------------

def _inject_xmp_metadata(pdf_path: Path, data: EmailData) -> None:
    """
    Inject email metadata as XMP into the PDF via pikepdf (best-effort).
    Called after PDF/A conversion so our metadata has the final say.
    Silently skips if pikepdf is not installed or the PDF cannot be opened.
    """
    try:
        import pikepdf
    except ImportError:
        return

    try:
        with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
            with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
                if data.subject:
                    meta["dc:title"] = data.subject
                if data.from_:
                    meta["dc:creator"] = [data.from_]
                if data.date:
                    meta["xmp:CreateDate"] = data.date.isoformat()
                recipients = ", ".join(filter(None, data.to + data.cc))
                if recipients:
                    meta["dc:description"] = f"To: {recipients}"
                if data.message_id:
                    meta["dc:identifier"] = data.message_id
            pdf.save(pdf_path)
    except Exception:
        pass  # metadata injection is best-effort; never fail a conversion for this


# ---------------------------------------------------------------------------
# PDF/A conversion
# ---------------------------------------------------------------------------

def _convert_to_pdfa(pdf_path: Path) -> bool:
    """
    Convert PDF -> PDF/A-2b via Ghostscript CLI (best-effort).

    Uses -dPDFACompatibilityPolicy=1 so Ghostscript auto-fixes what it can rather
    than failing hard on every issue. Returns True on success, False if Ghostscript
    is unavailable or the conversion fails.

    Output should be validated with veraPDF for strict archival compliance.
    """
    gs_cmd = (
        shutil.which("gswin64c")
        or shutil.which("gswin32c")
        or shutil.which("gs")
    )
    if not gs_cmd:
        return False

    tmp_path = pdf_path.with_suffix(".pdfa_tmp.pdf")

    try:
        result = subprocess.run(
            [
                gs_cmd,
                "-dPDFA=2",
                "-dBATCH",
                "-dNOPAUSE",
                "-dNOOUTERSAVE",
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.7",
                "-dPDFACompatibilityPolicy=1",
                f"-sOutputFile={tmp_path}",
                str(pdf_path),
            ],
            capture_output=True,
            timeout=120,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError):
        tmp_path.unlink(missing_ok=True)
        return False

    if result.returncode == 0 and tmp_path.exists() and tmp_path.stat().st_size > 0:
        pdf_path.unlink()
        tmp_path.rename(pdf_path)
        return True

    tmp_path.unlink(missing_ok=True)
    return False


# ---------------------------------------------------------------------------
# Output folder naming
# ---------------------------------------------------------------------------

def _sanitize(name: str, max_len: int = 80) -> str:
    """Replace filesystem-unsafe characters and truncate."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"[\s_]+", "_", name).strip("._")
    return name[:max_len] if name else "unnamed"


def _deduplicate(path: Path) -> Path:
    """Append _2, _3, ... until an unused path is found."""
    if not path.exists():
        return path
    i = 2
    while True:
        candidate = path.parent / f"{path.name}_{i}"
        if not candidate.exists():
            return candidate
        i += 1


def make_output_folder_name(data: EmailData) -> str:
    """
    Produce a sortable, filesystem-safe folder name:
        YYYY.MM.DD HH.MM - Subject
    """
    date_str = data.date.strftime("%Y.%m.%d %H.%M") if data.date else "0000.00.00 00.00"

    subject = data.subject or "(no subject)"
    subject = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", subject)  # strip filesystem-unsafe chars
    subject = re.sub(r"\s+", " ", subject).strip()            # collapse whitespace

    prefix = f"{date_str} - "
    # Windows MAX_PATH is 260. The PDF sits at base\folder\folder.pdf so the folder
    # name appears twice in the path; keep it under 80 chars to stay well within limits.
    max_subject = 80 - len(prefix)
    if len(subject) > max_subject:
        subject = subject[:max_subject].rstrip()
    return f"{prefix}{subject}"


# ---------------------------------------------------------------------------
# Attachment extraction
# ---------------------------------------------------------------------------

def extract_attachments(data: EmailData, output_dir: Path) -> list:
    """Save non-inline attachments to output_dir/attachments/. Returns metadata list."""
    non_inline = [a for a in data.attachments if not a.is_inline]
    if not non_inline:
        return []

    att_dir = output_dir / "attachments"
    att_dir.mkdir(exist_ok=True)

    saved      = []
    seen_names: dict = {}

    for att in non_inline:
        fname  = _sanitize(att.filename, max_len=100) or "attachment"
        stem   = Path(fname).stem
        suffix = Path(fname).suffix

        if fname in seen_names:
            seen_names[fname] += 1
            fname = f"{stem}_{seen_names[fname]}{suffix}"
        else:
            seen_names[fname] = 1

        (att_dir / fname).write_bytes(att.data)
        saved.append({
            "filename":     fname,
            "content_type": att.content_type,
            "size_bytes":   len(att.data),
        })

    return saved


# ---------------------------------------------------------------------------
# Metadata
# ---------------------------------------------------------------------------

def save_metadata(data: EmailData, output_dir: Path, attachments: list) -> None:
    meta = {
        "from":        data.from_,
        "to":          data.to,
        "cc":          data.cc,
        "bcc":         data.bcc,
        "subject":     data.subject,
        "date":        data.date.isoformat() if data.date else None,
        "message_id":  data.message_id,
        "source_file": Path(data.source_file).name,
        "attachments": attachments,
    }
    (output_dir / "metadata.json").write_text(
        json.dumps(meta, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Main conversion function
# ---------------------------------------------------------------------------

def convert_email(
    input_path: Path,
    output_base: Path,
    renderer: str = "playwright",
    pool: Optional[PlaywrightPool] = None,
    pdfa: bool = False,
) -> Path:
    """
    Convert one .eml or .msg file to PDF and extract attachments.
    Returns the path to the created output folder.
    """
    suffix = input_path.suffix.lower()

    print(f"  Parsing {input_path.name} ...")
    if suffix == ".eml":
        data = parse_eml(input_path)
    elif suffix == ".msg":
        data = parse_msg(input_path)
    else:
        raise ValueError(f"Unsupported format: {suffix!r} (expected .eml or .msg)")

    folder_name = make_output_folder_name(data)
    output_dir  = _deduplicate(output_base / folder_name)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"  Building HTML ...")
    html = build_html(data)
    html = _inline_css(html)

    pdf_path = output_dir / f"{output_dir.name}.pdf"

    print(f"  Rendering PDF ({renderer}) ...")
    if renderer == "playwright":
        render_pdf_playwright(html, pdf_path, pool=pool)
    else:
        RENDERERS[renderer](html, pdf_path)

    if pdfa and renderer == "playwright":
        print(f"  Converting to PDF/A-2b ...")
        if _convert_to_pdfa(pdf_path):
            print(f"    PDF/A-2b conversion successful.")
        else:
            print(f"    PDF/A-2b: Ghostscript not found or conversion failed — skipped.")

    # XMP metadata injection runs last (after PDF/A) so it has the final say
    _inject_xmp_metadata(pdf_path, data)

    print(f"  Extracting attachments ...")
    saved = extract_attachments(data, output_dir)
    save_metadata(data, output_dir, saved)

    n = len(saved)
    print(f"  -> {output_dir}  ({n} attachment{'s' if n != 1 else ''})")
    return output_dir


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Convert .eml and .msg files to PDF with attachment extraction."
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        metavar="FILE_OR_DIR",
        help=".eml/.msg file(s) or directory containing them",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        metavar="DIR",
        help=(
            "Output base directory. "
            "Defaults to an 'output' folder next to the input file/directory."
        ),
    )
    parser.add_argument(
        "--renderer",
        choices=list(RENDERERS),
        default="playwright",
        help="PDF renderer: playwright (default, best fidelity) or weasyprint (lighter)",
    )
    parser.add_argument(
        "--pdfa",
        action="store_true",
        default=False,
        help="Convert output to PDF/A-2b via Ghostscript (requires gs / gswin64c in PATH)",
    )
    args = parser.parse_args()

    # Collect input files as (file_path, input_root) tuples so we can mirror
    # the directory structure in the output.
    input_files: list = []  # list of (Path, Path) — (file, input_root)
    for inp in args.inputs:
        p = Path(inp).resolve()
        if p.is_dir():
            found = sorted(p.rglob("*.eml")) + sorted(p.rglob("*.msg"))
            for f in found:
                input_files.append((f, p))
        elif p.is_file() and p.suffix.lower() in {".eml", ".msg"}:
            input_files.append((p, p.parent))
        else:
            print(f"Warning: skipping {inp!r} (not found or unsupported)", file=sys.stderr)

    # Resolve output directory — default to output/ next to the first input
    if args.output:
        output_base = Path(args.output).resolve()
    elif input_files:
        first     = Path(args.inputs[0]).resolve()
        input_dir = first if first.is_dir() else first.parent
        output_base = input_dir.parent / f"{input_dir.name}_output"
    else:
        output_base = Path("output").resolve()

    output_base.mkdir(parents=True, exist_ok=True)

    if not input_files:
        print("No .eml or .msg files found.", file=sys.stderr)
        sys.exit(1)

    print(f"Converting {len(input_files)} email(s) -> {output_base}/\n")
    errors = []

    def _email_output_base(f: Path, input_root: Path) -> Path:
        """Mirror the input subdirectory structure under output_base."""
        rel = f.parent.relative_to(input_root)
        dest = output_base / rel
        dest.mkdir(parents=True, exist_ok=True)
        return dest

    if args.renderer == "playwright":
        # Single Chromium instance shared across the whole batch
        with PlaywrightPool() as pool:
            for f, input_root in input_files:
                print(f"[{f.name}]")
                try:
                    convert_email(
                        f, _email_output_base(f, input_root),
                        renderer=args.renderer, pool=pool, pdfa=args.pdfa,
                    )
                except Exception as exc:
                    print(f"  ERROR: {exc}", file=sys.stderr)
                    errors.append((f, exc))
    else:
        for f, input_root in input_files:
            print(f"[{f.name}]")
            try:
                convert_email(
                    f, _email_output_base(f, input_root),
                    renderer=args.renderer, pdfa=args.pdfa,
                )
            except Exception as exc:
                print(f"  ERROR: {exc}", file=sys.stderr)
                errors.append((f, exc))

    ok = len(input_files) - len(errors)
    print(f"\nFinished: {ok}/{len(input_files)} converted successfully.")
    if errors:
        print(f"\n{len(errors)} error(s):")
        for f, exc in errors:
            print(f"  {f.name}: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
