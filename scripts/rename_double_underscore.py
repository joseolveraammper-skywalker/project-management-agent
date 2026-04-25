"""One-off cleanup: rename "Daily_José__YYYYMMDD_..." files to use a single
underscore. Touches SharePoint (rawDailyJoseTXT + rawDailyJoseVTT), the local
.vtt_cache/ directory, and manifest.json. Idempotent — safe to re-run; files
already on the canonical name are skipped.
"""

import json
import re
import sys
from pathlib import Path

import requests

from vars import BEARER

SCRIPTS_DIR = Path(__file__).parent
VTT_CACHE_DIR = SCRIPTS_DIR / ".vtt_cache"
MANIFEST_PATH = VTT_CACHE_DIR / "manifest.json"

SYSTEMS_BASE = "https://ammper.sharepoint.com/sites/Systems/_api"
TXT_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseTXT"
VTT_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseVTT"

HEADERS = {
    "Authorization": f"Bearer {BEARER}",
    "Accept": "application/json;odata=nometadata",
}


def canonical(name: str) -> str:
    """Collapse any run of underscores to one, then strip leading/trailing
    underscores from the *base* (preserving the file extension)."""
    if "." in name:
        stem, ext = name.rsplit(".", 1)
        stem = re.sub(r"_+", "_", stem).strip("_")
        return f"{stem}.{ext}"
    return re.sub(r"_+", "_", name).strip("_")


def rename_in_sharepoint(folder_path: str, label: str) -> int:
    print(f"\n--- {label} ({folder_path}) ---")
    list_url = (
        f"{SYSTEMS_BASE}/web/GetFolderByServerRelativeUrl('{folder_path}')/Files"
    )
    r = requests.get(list_url, headers=HEADERS)
    if r.status_code == 401:
        print("  ✗ 401 Unauthorized — refresh vars.py via get_token.py.")
        sys.exit(1)
    files = r.json().get("value", [])

    renamed = 0
    for f in files:
        old_name = f["Name"]
        new_name = canonical(old_name)
        if old_name == new_name:
            continue
        old_url = f"{folder_path}/{old_name}"
        new_url = f"{folder_path}/{new_name}"
        move_url = (
            f"{SYSTEMS_BASE}/web/getfilebyserverrelativeurl('{old_url}')"
            f"/moveto(newurl='{new_url}',flags=1)"
        )
        resp = requests.post(move_url, headers=HEADERS)
        if resp.status_code in (200, 201, 204):
            print(f"  ✓ {old_name} → {new_name}")
            renamed += 1
        else:
            print(
                f"  ✗ {old_name} → {new_name}: "
                f"HTTP {resp.status_code}: {resp.text[:200]}"
            )
    if renamed == 0:
        print("  (nothing to rename)")
    return renamed


def rename_local_cache() -> int:
    print(f"\n--- Local cache ({VTT_CACHE_DIR}) ---")
    if not VTT_CACHE_DIR.exists():
        print("  (no .vtt_cache directory; skipping)")
        return 0

    renamed = 0
    for vtt_path in sorted(VTT_CACHE_DIR.glob("*.vtt")):
        new_name = canonical(vtt_path.name)
        if vtt_path.name == new_name:
            continue
        new_path = vtt_path.parent / new_name
        if new_path.exists():
            print(f"  ⚠ {vtt_path.name} → {new_name}: target exists, skipped")
            continue
        vtt_path.rename(new_path)
        print(f"  ✓ {vtt_path.name} → {new_name}")
        renamed += 1

    if MANIFEST_PATH.exists():
        with open(MANIFEST_PATH) as fh:
            manifest = json.load(fh)
        new_manifest = [canonical(n) for n in manifest]
        if new_manifest != manifest:
            with open(MANIFEST_PATH, "w") as fh:
                json.dump(new_manifest, fh, indent=2)
            print(f"  ✓ manifest.json updated ({len(new_manifest)} entries)")
        else:
            print("  (manifest.json already canonical)")

    if renamed == 0:
        print("  (no .vtt files needed renaming)")
    return renamed


def main():
    total = 0
    total += rename_in_sharepoint(TXT_FOLDER, "SharePoint TXT folder")
    total += rename_in_sharepoint(VTT_FOLDER, "SharePoint VTT folder")
    total += rename_local_cache()
    print(f"\nDone. {total} item(s) renamed.")


if __name__ == "__main__":
    main()
