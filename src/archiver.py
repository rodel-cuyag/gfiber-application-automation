"""
archiver.py
-----------
Moves processed input files into a dated archive folder after a
successful run, so the input directory empties out and is ready
for the next drop.
"""

import shutil
from pathlib import Path


def archive_files(paths: list[Path], archive_dir: Path) -> None:
    """Move each file in *paths* into *archive_dir* (created if needed)."""
    archive_dir.mkdir(parents=True, exist_ok=True)
    for path in paths:
        shutil.move(str(path), str(archive_dir / path.name))
