"""Prueft Datei-Namen von Bildern/Videos auf dem Handy (MTP) gegen das Muster
YYYYMMDD_HHMMSS[_N].ext und listet alle Dateien auf, die davon abweichen.

Es wird angenommen, dass die Zeitstempel im Dateinamen (vom Handy selbst
vergeben) korrekt sind - es findet kein Abgleich mit MTP-Metadaten statt
(deren "Geaendert"-Datum ist auf diesem Geraet nur minutengenau und war im
Test sogar um eine Stunde verschoben, also nicht vertrauenswuerdig).

Durchsucht werden nur die Verzeichnisse, die in json/<person>-mobile-dirs.json
gelistet sind.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import pywintypes
import win32com.client

from scan_phone_media import CSIDL_DRIVES, IMAGE_EXTENSIONS, VIDEO_EXTENSIONS, classify

DEFAULT_DEVICE_NAME = "S23 Ultra von Stefan"
DEFAULT_PERSON = "stefan"

REPO_ROOT = Path(__file__).resolve().parent
JSON_DIR = REPO_ROOT / "json"
RESULTS_DIR = REPO_ROOT / "results"

NAME_PATTERN = re.compile(r"^(\d{8})_(\d{6})(?:_\d+)?$")

logger = logging.getLogger("check_phone_media_names")


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        stream=sys.stdout,
    )


def available_persons() -> list[str]:
    return sorted(
        p.name[: -len("-mobile-dirs.json")]
        for p in JSON_DIR.glob("*-mobile-dirs.json")
    )


def prompt_person(default: str) -> str:
    persons = available_persons()
    choices = ", ".join(persons) if persons else "keine gefunden"
    answer = input(f"Person? [{default}] (verfuegbar: {choices}): ").strip()
    return answer or default


def load_target_dirs(person: str) -> list[str]:
    dirs_file = JSON_DIR / f"{person}-mobile-dirs.json"
    if not dirs_file.exists():
        raise RuntimeError(f"Datei nicht gefunden: {dirs_file}")

    data = json.loads(dirs_file.read_text(encoding="utf-8"))
    paths: list[str] = []
    for group_paths in data.values():
        paths.extend(group_paths)
    return paths


def get_device_root(device_name: str):
    shell = win32com.client.Dispatch("Shell.Application")
    this_pc = shell.Namespace(CSIDL_DRIVES)
    if this_pc is None:
        raise RuntimeError("Konnte 'Dieser PC' nicht oeffnen.")

    device_item = next((item for item in this_pc.Items() if item.Name == device_name), None)
    if device_item is None:
        available = ", ".join(item.Name for item in this_pc.Items())
        raise RuntimeError(
            f"Geraet '{device_name}' nicht unter 'Dieser PC' gefunden. "
            f"Ist das Handy eingesteckt und entsperrt? Verfuegbar: {available}"
        )
    return device_item.GetFolder


def resolve_path(device_root, relative_path: str):
    folder = device_root
    for part in relative_path.split("/"):
        item = next((i for i in folder.Items() if i.Name == part), None)
        if item is None:
            available = ", ".join(i.Name for i in folder.Items())
            raise RuntimeError(
                f"Teilpfad '{part}' von '{relative_path}' nicht gefunden. Verfuegbar: {available}"
            )
        folder = item.GetFolder
    return folder


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


def check_folder(folder, relative_path: str, violations: list[dict]) -> None:
    for item in folder.Items():
        if item.IsFolder:
            check_folder(item.GetFolder, f"{relative_path}/{item.Name}", violations)
            continue

        kind = classify(item.Name)
        if kind is None:
            continue

        stem = Path(item.Name).stem
        if not is_valid_name(stem):
            violations.append({"path": relative_path, "name": item.Name, "type": kind})
            logger.info("Abweichend: %s/%s", relative_path, item.Name)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Prueft Bild-/Video-Dateinamen auf dem Handy gegen das Muster YYYYMMDD_HHMMSS."
    )
    parser.add_argument("--device-name", default=DEFAULT_DEVICE_NAME, help="Name unter 'Dieser PC'")
    parser.add_argument("--person", default=None, help="Person, deren Verzeichnisliste genutzt wird (z.B. stefan, pia)")
    args = parser.parse_args()

    setup_logging()

    person = args.person or prompt_person(DEFAULT_PERSON)
    logger.info("Person: %s", person)

    try:
        target_dirs = load_target_dirs(person)
        device_root = get_device_root(args.device_name)
    except RuntimeError as exc:
        logger.error(str(exc))
        return 1

    violations: list[dict] = []
    for relative_path in target_dirs:
        logger.info("Durchsuche: %s", relative_path)
        try:
            folder = resolve_path(device_root, relative_path)
            check_folder(folder, relative_path, violations)
        except (RuntimeError, pywintypes.com_error) as exc:
            logger.warning("Konnte '%s' nicht durchsuchen: %s", relative_path, exc)

    logger.info("Fertig. %d abweichende Datei(en) gefunden.", len(violations))

    RESULTS_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_device_name = args.device_name.replace(" ", "_")
    output_file = RESULTS_DIR / f"{safe_device_name}_{person}_name_check_{timestamp}.json"
    output_file.write_text(
        json.dumps(
            {
                "device_name": args.device_name,
                "person": person,
                "checked_at": timestamp,
                "pattern": "YYYYMMDD_HHMMSS[_N].ext",
                "directories": target_dirs,
                "violations": violations,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    logger.info("Ergebnis gespeichert: %s", output_file)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
