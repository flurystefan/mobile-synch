"""Wiederverwendbare Hilfsfunktion zum Lesen des Aufnahmedatums aus EXIF-Daten.

Wird von mehreren Scripts in diesem Repo genutzt, die Datei-Zeitstempel
mit dem tatsaechlichen Aufnahmedatum abgleichen wollen.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PIL import Image

EXIF_TAG_DATETIME = 306  # IFD0 DateTime
EXIF_TAG_DATETIME_ORIGINAL = 36867  # Exif SubIFD DateTimeOriginal
EXIF_TAG_DATETIME_DIGITIZED = 36868  # Exif SubIFD DateTimeDigitized
EXIF_SUBIFD_TAG = 0x8769


def get_exif_datetime(file_path: Path) -> datetime | None:
    """Liest das Aufnahmedatum aus den EXIF-Daten einer Bilddatei.

    Bevorzugt DateTimeOriginal/DateTimeDigitized (Exif SubIFD) vor dem
    generischen DateTime (IFD0). Gibt None zurueck, wenn die Datei keine
    lesbaren EXIF-Zeitstempel hat (z.B. Videos, Screenshots, PNGs ohne EXIF).
    """
    try:
        with Image.open(file_path) as img:
            exif = img.getexif()
            if not exif:
                return None

            value = None
            try:
                exif_subifd = exif.get_ifd(EXIF_SUBIFD_TAG)
                value = exif_subifd.get(EXIF_TAG_DATETIME_ORIGINAL) or exif_subifd.get(
                    EXIF_TAG_DATETIME_DIGITIZED
                )
            except (KeyError, ValueError):
                pass

            if not value:
                value = exif.get(EXIF_TAG_DATETIME)

            if not value:
                return None

            return datetime.strptime(value.strip(), "%Y:%m:%d %H:%M:%S")
    except Exception:
        return None
