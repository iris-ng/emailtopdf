#!/usr/bin/env python3
"""
collect_pdfs.py — Copy email PDFs from the output folder into a flat _emails_only folder.

Walks output_dir recursively. The email PDF for each converted email is named
{dirname}/{dirname}.pdf (matching the folder name). This script collects those
PDFs — ignoring attachment PDFs — and copies them flat into a sibling folder:

    output/                                      <- source (output_dir)
      2024.03.15 09.30 - Q1 Budget Review/
        2024.03.15 09.30 - Q1 Budget Review.pdf  <- copied
        metadata.json
        attachments/
          report.xlsx                            <- NOT copied (attachment)

    output_emails_only/                          <- destination (created automatically)
      2024.03.15 09.30 - Q1 Budget Review.pdf

Usage:
    python collect_pdfs.py                    # output/ next to this script
    python collect_pdfs.py path/to/output/    # explicit path
"""

import sys
import shutil
from pathlib import Path


def collect_email_pdfs(output_dir: Path) -> int:
    """
    Copy email PDFs from output_dir into a sibling _emails_only folder.
    Returns the number of files copied.
    """
    dest = output_dir.parent / f"{output_dir.name}_emails_only"
    dest.mkdir(parents=True, exist_ok=True)

    print(f"Source : {output_dir}")
    print(f"Dest   : {dest}")
    print()

    copied = 0
    skipped = 0

    # A subdirectory is an email folder when it contains a PDF of the same name.
    for subdir in sorted(output_dir.rglob("*")):
        if not subdir.is_dir():
            continue
        pdf = subdir / f"{subdir.name}.pdf"
        if not pdf.exists():
            continue

        target = dest / pdf.name
        if target.exists():
            # Avoid silent overwrites: append a counter suffix.
            stem, suffix = pdf.stem, pdf.suffix
            n = 2
            while target.exists():
                target = dest / f"{stem}_{n}{suffix}"
                n += 1
            print(f"  [RENAME] {pdf.name} -> {target.name}  (name conflict)")
            skipped += 1

        shutil.copy2(pdf, target)
        print(f"  [COPY]   {pdf.name}")
        copied += 1

    print(f"\nDone: {copied} file(s) copied, {skipped} rename(s) due to name conflicts.")
    return copied


def main() -> None:
    if len(sys.argv) > 1:
        output_dir = Path(sys.argv[1]).resolve()
    else:
        output_dir = Path(__file__).parent / "output"

    if not output_dir.exists():
        print(f"Error: directory not found: {output_dir}", file=sys.stderr)
        sys.exit(1)
    if not output_dir.is_dir():
        print(f"Error: not a directory: {output_dir}", file=sys.stderr)
        sys.exit(1)

    collect_email_pdfs(output_dir)


if __name__ == "__main__":
    main()
