"""Kopiert die generierten geoparquet/fgdb/AGP Ergebnisse vom lokalen Rechner
auf das NAS, damit die Fotokarten-Webapp (Docker auf dem NAS) und ArcGIS Pro
Projekte darauf direkt zugreifen koennen, ohne ueber diesen Windows-Rechner
zu gehen.

Nutzt robocopy (inkrementell - kopiert nur neue/geaenderte Dateien, laesst
Vorhandenes am Ziel unberuehrt statt es zu spiegeln/loeschen).
"""

from __future__ import annotations

import argparse
import logging
import subprocess
import sys
from pathlib import Path

DEFAULT_PERSON = "stefan"
DEFAULT_NAS_BASE = r"\\DS923Plus\PhoneMirror"
DEFAULT_FOLDERS = ["geoparquet", "fgdb", "AGP"]

REPO_ROOT = Path(__file__).resolve().parent
RESULTS_DIR = REPO_ROOT / "results"

logger = logging.getLogger("sync_results_to_nas")


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stdout,
    )


def prompt_choice(label: str, default: str) -> str:
    answer = input(f"{label}? [{default}]: ").strip()
    return answer or default


def sync_folder(source: Path, destination: Path) -> bool:
    if not source.exists():
        logger.warning("Quelle nicht gefunden, uebersprungen: %s", source)
        return False

    destination.mkdir(parents=True, exist_ok=True)
    logger.info("Synce %s -> %s", source, destination)
    result = subprocess.run(
        ["robocopy", str(source), str(destination), "/E", "/XO", "/R:2", "/W:2", "/NFL", "/NDL"],
        capture_output=True,
        text=True,
    )
    # robocopy: Exit-Codes 0-7 sind Erfolg, ab 8 ein Fehler.
    if result.returncode >= 8:
        logger.error("robocopy Fehler (Code %d): %s", result.returncode, result.stdout[-2000:])
        return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Kopiert geoparquet/fgdb/AGP Ergebnisse aufs NAS fuer die Fotokarten-Webapp."
    )
    parser.add_argument("--person", default=None, help="Person, deren Ergebnisse synchronisiert werden (Default: stefan)")
    parser.add_argument("--nas-base", default=DEFAULT_NAS_BASE, help=f"NAS-Basispfad (Default: {DEFAULT_NAS_BASE})")
    parser.add_argument(
        "--folders",
        default=",".join(DEFAULT_FOLDERS),
        help=f"Kommagetrennte Liste der zu synchronisierenden Unterordner (Default: {','.join(DEFAULT_FOLDERS)})",
    )
    args = parser.parse_args()

    setup_logging()

    person = args.person or prompt_choice("Person", DEFAULT_PERSON)
    folders = [f.strip() for f in args.folders.split(",") if f.strip()]

    source_root = RESULTS_DIR / person
    destination_root = Path(args.nas_base) / person / "gis"

    logger.info("Person: %s  Ordner: %s", person, ", ".join(folders))
    logger.info("Quelle: %s", source_root)
    logger.info("Ziel:   %s", destination_root)

    all_ok = True
    for folder in folders:
        all_ok = sync_folder(source_root / folder, destination_root / folder) and all_ok

    if all_ok:
        logger.info("Fertig. Alle Ordner synchronisiert.")
        return 0

    logger.error("Es gab Fehler beim Synchronisieren, siehe Log oben.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
