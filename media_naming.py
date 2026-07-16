"""Gemeinsame Pruef-Logik fuer Bild-/Video-Dateinamen nach dem Muster
YYYYMMDD_HHMMSS, optional mit Duplikat-Suffix ("_1" oder " (1)")."""

from __future__ import annotations

import re
from datetime import datetime

NAME_PATTERN = re.compile(r"^(\d{8})_(\d{6})(?:_\d+| \(\d+\))?$")


def is_valid_name(stem: str) -> bool:
    match = NAME_PATTERN.match(stem)
    if not match:
        return False
    date_part, time_part = match.group(1), match.group(2)
    try:
        datetime.strptime(f"{date_part}_{time_part}", "%Y%m%d_%H%M%S")
    except ValueError:
        return False
    return True
