"""Where the app keeps its files, outside the app folder so updates keep them.

Downloading a new version of the app gives a fresh folder; anything stored
inside the old one (uploaded templates and price lists, running numbers,
generated PDFs) would be left behind. Everything is therefore kept in
"DM Optical App" in the user's Documents folder (or home folder), or in
DMO_APP_HOME when that is set.
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path


def app_home() -> Path:
    override = os.environ.get("DMO_APP_HOME")
    if override:
        return Path(override)
    documents = Path.home() / "Documents"
    return (documents if documents.is_dir() else Path.home()) / "DM Optical App"


HOME = app_home()
SAVED_FILES = HOME / "Saved files"  # templates, price lists, signature & stamp
RECORDS = HOME / "Records"  # running numbers, remembered customers
OUTPUT = HOME / "Generated documents"  # the PDFs, one folder per date

# Folders inside the app folder used by earlier versions, and where they go now.
LEGACY_FOLDERS = {"private": SAVED_FILES, "data": RECORDS, "output": OUTPUT}


def migrate_legacy(app_root: Path) -> list[Path]:
    """Copy files from an earlier version's in-app folders, never overwriting newer ones."""
    copied = []
    for old_name, new_dir in LEGACY_FOLDERS.items():
        old_dir = Path(app_root) / old_name
        if not old_dir.is_dir():
            continue
        for src in old_dir.rglob("*"):
            dst = new_dir / src.relative_to(old_dir)
            if src.is_file() and not dst.exists():
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
                copied.append(dst)
    return copied
