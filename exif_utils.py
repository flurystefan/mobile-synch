"""Wiederverwendbare Hilfsfunktionen zum Lesen von EXIF-Daten (Aufnahmedatum,
GPS-Koordinaten, Kamera-Metadaten) aus Bilddateien.

Wird von mehreren Scripts in diesem Repo genutzt. Die low-level Funktionen
(extract_*) arbeiten auf einem bereits geoeffneten Exif-Objekt (img.getexif()),
damit ein Aufrufer, der mehrere Werte braucht (z.B. fuers GeoParquet-Script,
das zusaetzlich noch ein Thumbnail generiert), die Bilddatei nur einmal
oeffnen muss.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image

EXIF_TAG_DATETIME = 306  # IFD0 DateTime
EXIF_TAG_DATETIME_ORIGINAL = 36867  # Exif SubIFD DateTimeOriginal
EXIF_TAG_DATETIME_DIGITIZED = 36868  # Exif SubIFD DateTimeDigitized
EXIF_TAG_MAKE = 271
EXIF_TAG_MODEL = 272
EXIF_TAG_ORIENTATION = 274
EXIF_TAG_LENS_MODEL = 42036  # Exif SubIFD
EXIF_TAG_FOCAL_LENGTH = 37386  # Exif SubIFD
EXIF_TAG_F_NUMBER = 33437  # Exif SubIFD
EXIF_TAG_EXPOSURE_TIME = 33434  # Exif SubIFD
EXIF_TAG_ISO_SPEED = 34855  # Exif SubIFD

EXIF_SUBIFD_TAG = 0x8769
GPS_IFD_TAG = 0x8825

GPS_TAG_LATITUDE_REF = 1
GPS_TAG_LATITUDE = 2
GPS_TAG_LONGITUDE_REF = 3
GPS_TAG_LONGITUDE = 4
GPS_TAG_ALTITUDE_REF = 5
GPS_TAG_ALTITUDE = 6


def _get_exif_subifd(exif) -> dict:
    try:
        return exif.get_ifd(EXIF_SUBIFD_TAG) or {}
    except (KeyError, ValueError):
        return {}


def extract_datetime_original(exif) -> datetime | None:
    """Bevorzugt DateTimeOriginal/DateTimeDigitized (Exif SubIFD) vor dem
    generischen DateTime (IFD0)."""
    if not exif:
        return None

    subifd = _get_exif_subifd(exif)
    value = subifd.get(EXIF_TAG_DATETIME_ORIGINAL) or subifd.get(EXIF_TAG_DATETIME_DIGITIZED)
    if not value:
        value = exif.get(EXIF_TAG_DATETIME)
    if not value:
        return None

    try:
        return datetime.strptime(value.strip(), "%Y:%m:%d %H:%M:%S")
    except ValueError:
        return None


def _dms_to_decimal(dms, ref: str) -> float | None:
    """Wandelt ein GPS Grad/Minuten/Sekunden-Tripel in Dezimalgrad um.

    Manche Kameras schreiben defekte Rationale (Nenner 0) in einzelne
    GPS-Felder. Dann kann keine Koordinate berechnet werden.
    """
    try:
        degrees, minutes, seconds = (float(v) for v in dms)
    except (ZeroDivisionError, TypeError, ValueError):
        return None
    decimal = degrees + minutes / 60 + seconds / 3600
    if ref in ("S", "W"):
        decimal = -decimal
    return decimal


def extract_gps(exif) -> tuple[float, float, float | None] | None:
    """Gibt (latitude, longitude, altitude_m) in WGS84/EPSG:4326 zurueck,
    oder None wenn keine (gueltigen) GPS-Koordinaten vorhanden sind.

    Ein defekter Hoehenwert (z.B. Rational mit Nenner 0) verhindert nicht
    das Setzen des Punkts - dann wird nur altitude_m auf None gesetzt.
    """
    if not exif:
        return None

    try:
        gps_ifd = exif.get_ifd(GPS_IFD_TAG)
    except (KeyError, ValueError):
        return None
    if not gps_ifd:
        return None

    lat_dms = gps_ifd.get(GPS_TAG_LATITUDE)
    lat_ref = gps_ifd.get(GPS_TAG_LATITUDE_REF)
    lon_dms = gps_ifd.get(GPS_TAG_LONGITUDE)
    lon_ref = gps_ifd.get(GPS_TAG_LONGITUDE_REF)
    if not (lat_dms and lat_ref and lon_dms and lon_ref):
        return None

    latitude = _dms_to_decimal(lat_dms, lat_ref)
    longitude = _dms_to_decimal(lon_dms, lon_ref)
    if latitude is None or longitude is None:
        return None

    altitude = None
    alt_value = gps_ifd.get(GPS_TAG_ALTITUDE)
    if alt_value is not None:
        try:
            altitude = float(alt_value)
        except (ZeroDivisionError, TypeError, ValueError):
            altitude = None
        else:
            if gps_ifd.get(GPS_TAG_ALTITUDE_REF) == 1:
                altitude = -altitude

    return latitude, longitude, altitude


def extract_camera_metadata(exif) -> dict[str, Any]:
    """Liest Kamera-/Aufnahme-Metadaten (Make, Model, Objektiv, Belichtung, ...)."""
    if not exif:
        return {}

    subifd = _get_exif_subifd(exif)
    focal_length = subifd.get(EXIF_TAG_FOCAL_LENGTH)
    f_number = subifd.get(EXIF_TAG_F_NUMBER)
    exposure_time = subifd.get(EXIF_TAG_EXPOSURE_TIME)
    iso_speed = subifd.get(EXIF_TAG_ISO_SPEED)

    return {
        "camera_make": exif.get(EXIF_TAG_MAKE),
        "camera_model": exif.get(EXIF_TAG_MODEL),
        "orientation": exif.get(EXIF_TAG_ORIENTATION),
        "lens_model": subifd.get(EXIF_TAG_LENS_MODEL),
        "focal_length_mm": float(focal_length) if focal_length is not None else None,
        "f_number": float(f_number) if f_number is not None else None,
        "exposure_time_s": float(exposure_time) if exposure_time is not None else None,
        "iso_speed": int(iso_speed) if iso_speed is not None else None,
    }


def get_exif_datetime(file_path: Path) -> datetime | None:
    """Liest nur das Aufnahmedatum einer Bilddatei (oeffnet die Datei selbst).

    Gibt None zurueck, wenn die Datei keine lesbaren EXIF-Zeitstempel hat
    (z.B. Videos, Screenshots, PNGs ohne EXIF).
    """
    try:
        with Image.open(file_path) as img:
            return extract_datetime_original(img.getexif())
    except Exception:
        return None
