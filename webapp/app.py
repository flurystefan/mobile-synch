"""Leichte Flask-App, die pro Verzeichnis (Album) die Bilder aus einem
GeoParquet als Punkte auf einer Karte anzeigt.

Liest die Parquet-Dateien direkt ueber pyarrow (keine geopandas/GDAL
Abhaengigkeit), damit das Docker-Image auf dem NAS klein bleibt. latitude/
longitude/thumbnail liegen bereits als eigene Spalten im Parquet (siehe
build_nas_geoparquet.py), ein Dekodieren der WKB-Geometrie ist nicht noetig.

Fuer die Vollbildansicht wird nicht das eingebettete 600px-Thumbnail
verwendet (zu grob, wirkt beim Hochskalieren "milchig"), sondern das
Originalbild direkt von PHOTOS_ROOT. Der full_path im Parquet ist ein
Windows-Pfad (so wie build_nas_geoparquet.py ihn geschrieben hat); SOURCE_ROOT
ist der Praefix davon, der durch PHOTOS_ROOT ersetzt wird, um den Pfad
innerhalb des Containers aufzuloesen.
"""

from __future__ import annotations

import os
from pathlib import Path

import pyarrow.parquet as pq
from flask import Flask, Response, abort, jsonify, render_template, send_file

PARQUET_ROOT = Path(os.environ.get("PARQUET_ROOT", "/data/geoparquet")).resolve()
PHOTOS_ROOT = Path(os.environ.get("PHOTOS_ROOT", "/data/photos"))
SOURCE_ROOT = os.environ.get("SOURCE_ROOT", r"\\DS923Plus\PhoneMirror\stefan\fotos\DCIM")

app = Flask(__name__)


def resolve_original_path(full_path: str) -> Path:
    normalized = full_path.replace("\\", "/")
    prefix = SOURCE_ROOT.replace("\\", "/").rstrip("/")
    if normalized.startswith(prefix):
        relative = normalized[len(prefix):].lstrip("/")
    else:
        relative = Path(normalized).name
    return PHOTOS_ROOT / relative


def list_albums() -> list[str]:
    if not PARQUET_ROOT.exists():
        return []
    albums = [
        str(path.relative_to(PARQUET_ROOT).with_suffix("")).replace(os.sep, "/")
        for path in PARQUET_ROOT.rglob("*.parquet")
    ]
    return sorted(albums)


def resolve_album_path(album: str) -> Path:
    candidate = (PARQUET_ROOT / album).with_suffix(".parquet")
    resolved = candidate.resolve()
    if PARQUET_ROOT not in resolved.parents and resolved != PARQUET_ROOT:
        abort(404)
    if not resolved.is_file():
        abort(404)
    return resolved


@app.route("/")
def index():
    return render_template("index.html", albums=list_albums())


@app.route("/map/<path:album>")
def map_view(album: str):
    resolve_album_path(album)
    return render_template("map.html", album=album)


@app.route("/api/points/<path:album>")
def api_points(album: str):
    path = resolve_album_path(album)
    table = pq.read_table(path, columns=["filename", "latitude", "longitude", "datetime_original"])
    points = [
        {
            "filename": row["filename"],
            "lat": row["latitude"],
            "lon": row["longitude"],
            "datetime": row["datetime_original"],
        }
        for row in table.to_pylist()
        if row["latitude"] is not None and row["longitude"] is not None
    ]
    return jsonify(points)


@app.route("/api/thumbnail/<path:album>/<filename>")
def api_thumbnail(album: str, filename: str):
    path = resolve_album_path(album)
    table = pq.read_table(path, columns=["filename", "thumbnail"])
    names = table.column("filename").to_pylist()
    try:
        index = names.index(filename)
    except ValueError:
        abort(404)
    thumbnail = table.column("thumbnail")[index].as_py()
    return Response(thumbnail, mimetype="image/jpeg")


@app.route("/api/photo/<path:album>/<filename>")
def api_photo(album: str, filename: str):
    path = resolve_album_path(album)
    table = pq.read_table(path, columns=["filename", "full_path"])
    names = table.column("filename").to_pylist()
    try:
        index = names.index(filename)
    except ValueError:
        abort(404)
    full_path = table.column("full_path")[index].as_py()
    source = resolve_original_path(full_path)
    if not source.is_file():
        abort(404)
    return send_file(source)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
