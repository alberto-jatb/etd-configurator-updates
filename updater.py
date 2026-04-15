"""
Auto-updater for ETD Configurator.

On startup, fetches a manifest.json from UPDATE_MANIFEST_URL.
If a newer version is available, downloads the changed source files
into the local app-data directory. On the next launch (or after restart)
the updated files are loaded automatically.

To publish an update:
  1. Edit the source files.
  2. Run generate_manifest.py  →  produces manifest.json.
  3. Upload manifest.json + changed .py files to the same URL base.
"""

import os
import sys
import json
import hashlib
import urllib.request
import urllib.error
from pathlib import Path

# ── Update feed (public repo, no token required) ──────────────────────────────
UPDATE_BASE_URL = "https://raw.githubusercontent.com/alberto-jatb/etd-configurator-updates/main/"
GITHUB_TOKEN = ""
# ──────────────────────────────────────────────────────────────────────────────

MANIFEST_URL    = UPDATE_BASE_URL + "manifest.json"
UPDATABLE_FILES = ["app.py", "geo_data.py", "ps_generator.py", "executor.py", "version.py", "theme_neutral.json"]
TIMEOUT         = 6  # seconds


def local_dir() -> Path:
    """Return (and create) the directory where updated files are stored."""
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home())) / "ETD_Configurator"
    else:
        base = Path.home() / "Library" / "Application Support" / "ETD_Configurator"
    base.mkdir(parents=True, exist_ok=True)
    return base


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()


def check_and_update(current_version: str):
    """
    Check for a newer version and download changed files.

    Returns:
        (new_version: str, updated_files: list[str])
        or (None, []) if already up to date or unreachable.
    """
    def _open(url):
        req = urllib.request.Request(url)
        if GITHUB_TOKEN and GITHUB_TOKEN != "PASTE_YOUR_TOKEN_HERE":
            req.add_header("Authorization", f"token {GITHUB_TOKEN}")
        return urllib.request.urlopen(req, timeout=TIMEOUT)

    try:
        with _open(MANIFEST_URL) as resp:
            manifest = json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None, []

    remote_version = manifest.get("version", "0")
    if remote_version <= current_version:
        return None, []

    base_url  = manifest.get("base_url", UPDATE_BASE_URL)
    dest      = local_dir()
    updated   = []

    for filename, remote_sha in manifest.get("files", {}).items():
        if filename not in UPDATABLE_FILES:
            continue
        local_path = dest / filename
        if local_path.exists() and _sha256(local_path) == remote_sha:
            continue
        try:
            url = base_url + filename
            with _open(url) as resp:
                local_path.write_bytes(resp.read())
            updated.append(filename)
        except Exception:
            pass

    return remote_version, updated
