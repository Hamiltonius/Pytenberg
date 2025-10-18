#!/usr/bin/env python3
"""
Pytenberg - Email to Project Automation
Turns inbox chaos into organized folder structures.

Author: Thomas Galarneau
Version: 1.0.0
License: MIT
"""

__version__ = "1.0.0"
__author__ = "Thomas Galarneau"

import os
import sys
import shutil
import re
import unicodedata
from pathlib import Path
from typing import Optional

try:
    import extract_msg
except ImportError:
    print("Error: extract_msg library not found.")
    print("Install with: pip install extract-msg")
    sys.exit(1)

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).parent
DROP_FOLDER = SCRIPT_DIR / "drop"
OUT_FOLDER = SCRIPT_DIR / "out"
REFS_FOLDER = SCRIPT_DIR / "refs"
LOGS_FOLDER = SCRIPT_DIR / "logs"

# Regex to remove prefixes like "Re:" or "Fwd:" before matching
SUBJECT_PREFIX_RE = re.compile(r'(?i)^(?:re|fwd?|aw|sv)\s*:\s*')

# Regex variant presets for various subject styles
VARIANTS = {
    "invoice": r"(?i)(invoice|inv)[:\s#-]*([A-Z0-9-]+)",
    "project": r"(?i)(project|proj)[:\s#-]*([A-Z0-9-]+)",
    "client": r"(?i)(client|customer)[:\s#-]*([A-Za-z0-9\s]+?)(?=\s*[-:]|\s*$)",
    "case": r"(?i)(case|ticket)[:\s#-]*([A-Z0-9-]+)",
    "order": r"(?i)(order|po)[:\s#-]*([A-Z0-9-]+)",
    "contract": r"(?i)(contract|agreement)[:\s#-]*([A-Z0-9-]+)",
    "quote": r"(?i)(quote|rfq)[:\s#-]*([A-Z0-9-]+)",
    "proposal": r"(?i)(proposal|rfp)[:\s#-]*([A-Z0-9-]+)",
    "homework": r"(?i)(hw|homework|assignment)[:\s#-]*(\d+|[A-Z]+\d+)",
    "class": r"(?i)(class|course)[:\s#-]*([A-Z]{2,4}\s?\d{3,4})",
    "default": r"(?i)^([A-Za-z0-9&'().\s]+?)(?=\s*[-:])",
    "aerospace_code": r"(?<![A-Za-z0-9])[0-9][A-Za-z0-9]{9}(?![A-Za-z0-9])",
    "generic": r"[\[\(]([A-Za-z0-9\-_\s]+)[\]\)]|:\s*([A-Za-z0-9\-_]+)"
}

# Select active variant
ACTIVE_VARIANT = "default"  # e.g., "whole_subject_extract" or "aerospace_code"

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def sanitize_filename(name: str) -> str:
    """Ensure a filename is safe across all operating systems."""
    if not name:
        return "attachment.bin"
    name = unicodedata.normalize("NFKC", name)
    name = ''.join('_' if (ord(c) < 32 or c in '<>:"/\\|?*') else c for c in name)
    name = name.replace('\x00', '').replace('\n', '').replace('\r', '')
    name = name.strip().strip('_').strip('.')
    return name or "attachment.bin"


def dedupe_path(filepath: Path) -> Path:
    """Return a unique filepath by appending (1), (2), etc., if file exists."""
    if not filepath.exists():
        return filepath
    stem, suffix, parent = filepath.stem, filepath.suffix, filepath.parent
    counter = 1
    while True:
        new_path = parent / f"{stem} ({counter}){suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def clean_subject(subject: Optional[str]) -> str:
    """Strip prefixes and whitespace from an email subject line."""
    s = (subject or "No Subject").strip()
    s = SUBJECT_PREFIX_RE.sub("", s)
    return s


def extract_project_code(subject: str, variant: str) -> Optional[str]:
    """
    Extract a project or identifier code from an email subject using a regex variant.
    Returns a sanitized string safe for folder names, or None if no match.
    """
    subj = clean_subject(subject)
    match = re.search(variant, subj)
    if not match:
        return None

    groups = [g for g in match.groups() if g]
    if not groups:
        return None

    return sanitize_filename(groups[-1]).replace(" ", "_")


# -----------------------------------------------------------------------------
# CORE PROCESS
# -----------------------------------------------------------------------------
def process_msg_file(msg_path: Path, variant: str) -> bool:
    """
    Process a single .msg file:
      - Extract project code
      - Create folders
      - Copy reference files
      - Save attachments
    Returns True if processed successfully, False otherwise.
    """
    msg = None
    try:
        msg = extract_msg.Message(str(msg_path))
        subject = msg.subject or "No Subject"
        project_code = extract_project_code(subject, variant)

        if not project_code:
            print(f"⚠️  No match in: {subject}")
            return False

        project_folder = OUT_FOLDER / project_code
        archive_folder = project_folder / "archive"

        project_folder.mkdir(exist_ok=True)
        archive_folder.mkdir(exist_ok=True)

        # Copy reference files from refs/ if present
        if REFS_FOLDER.exists():
            for ref_file in REFS_FOLDER.iterdir():
                if ref_file.is_file():
                    dest = project_folder / ref_file.name
                    if not dest.exists():
                        shutil.copy2(ref_file, dest)

        # Copy original .msg file into archive
        msg_dest = dedupe_path(archive_folder / msg_path.name)
        shutil.copy2(msg_path, msg_dest)

        # Extract and save attachments
        attachment_count = 0
        for attachment in getattr(msg, "attachments", []) or []:
            raw_name = (
                getattr(attachment, "longFilename", None)
                or getattr(attachment, "shortFilename", None)
                or "attachment.bin"
            )
            safe_name = sanitize_filename(raw_name)
            attachment_path = dedupe_path(project_folder / safe_name)
            with open(attachment_path, "wb") as f:
                f.write(attachment.data)
            attachment_count += 1

        print(f"✅ {project_code} ({attachment_count} attachments)")
        return True

    except Exception as e:
        print(f"❌ Error processing {msg_path.name}: {e}")
        return False
    finally:
        try:
            if msg:
                msg.close()
        except:
            pass


# -----------------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------------
def main():
    """Main entry point — processes all .msg files in drop/ folder."""
    for folder in [DROP_FOLDER, OUT_FOLDER, REFS_FOLDER, LOGS_FOLDER]:
        folder.mkdir(exist_ok=True)

    variant = VARIANTS.get(ACTIVE_VARIANT)
    if not variant:
        print(f"Variant '{ACTIVE_VARIANT}' not found in VARIANTS: {list(VARIANTS.keys())}")
        return

    print(f"{'='*60}")
    print(f"Pytenberg v{__version__}")
    print(f"{'='*60}")
    print(f"Active variant: {ACTIVE_VARIANT}")
    print(f"Drop folder:    {DROP_FOLDER}")
    print(f"Output folder:  {OUT_FOLDER}")
    print(f"{'='*60}\n")

    msg_files = list(DROP_FOLDER.glob("*.msg"))
    if not msg_files:
        print("No .msg files found in drop/ folder.")
        print("\n💡 Example: save a test email with a subject that fits your variant.")
        print("   Drop it into 'drop/' and rerun this script.")
        return

    processed = failed = 0
    for msg_file in msg_files:
        if process_msg_file(msg_file, variant):
            processed += 1
        else:
            failed += 1

    print(f"\n{'='*60}")
    print(f"✅ Successfully processed: {processed}")
    print(f"❌ Failed: {failed}")
    print(f"📁 Output folder: {OUT_FOLDER}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
