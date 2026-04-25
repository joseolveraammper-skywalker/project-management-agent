import argparse
import json
import os
import subprocess
import sys
from urllib.parse import parse_qsl
import requests
from datetime import datetime
from vars import *

HEADERS = {
    "Authorization": f"Bearer {BEARER}",
    "Accept": "application/json",
}

DRIVE_ID = "b!RBHDsJ-Kl0WZ6cEyf0jmmagT-TaeirNNiWr7WlVTYAI0Exing_vnT7kZf75UE4qs"
PERSONAL_BASE = "https://ammper-my.sharepoint.com/personal/jolvera_ammper_com/_api"
SYSTEMS_BASE = "https://ammper.sharepoint.com/sites/Systems/_api"
VTT_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseVTT"

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
VTT_CACHE_DIR = os.path.join(SCRIPTS_DIR, ".vtt_cache")
MANIFEST_PATH = os.path.join(VTT_CACHE_DIR, "manifest.json")

NAME_PREFIXES = ("Daily José", "Daily Jose")


def list_daily_recordings(verbose=True):
    """Pull every row in /Documents/Grabaciones (across all pages) and keep
    the ones whose filename starts with 'Daily José' or 'Daily Jose'.
    Sorted oldest-first."""
    url = f"{PERSONAL_BASE}/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream"

    base_params = {
        "@a1": "'/personal/jolvera_ammper_com/Documents'",
        "RootFolder": "/personal/jolvera_ammper_com/Documents/Grabaciones",
        "TryNewExperienceSingle": "TRUE",
    }

    matching = []
    params = dict(base_params)
    page = 0
    total_rows = 0

    while True:
        page += 1
        r = requests.post(
            url,
            headers={**HEADERS, "Content-Type": "application/json"},
            params=params,
            json={},
        )
        if r.status_code == 401:
            raise RuntimeError(
                "401 Unauthorized — bearer token in vars.py is expired. "
                "Run `python get_token.py` to refresh it."
            )
        data = r.json()
        if "Row" not in data:
            raise RuntimeError(
                f"Unexpected response (status {r.status_code}):\n{data}"
            )

        rows = data["Row"]
        total_rows += len(rows)
        for row in rows:
            name = row.get("FileLeafRef") or row.get("Title") or ""
            if any(name.startswith(p) for p in NAME_PREFIXES):
                matching.append(row)

        if verbose:
            print(
                f"  page {page}: scanned {len(rows)} row(s), "
                f"running total {total_rows}, matches so far {len(matching)}"
            )

        next_href = data.get("NextHref")
        if not next_href:
            break

        # NextHref is a query string like "?Paged=TRUE&p_ID=120&...".
        # Merge its tokens onto the base params for the next page.
        next_params = dict(base_params)
        for k, v in parse_qsl(next_href.lstrip("?")):
            next_params[k] = v
        params = next_params

        # Safety stop in case the API misbehaves.
        if page >= 200:
            print("⚠️  Stopped at 200 pages — bailing out to avoid an infinite loop.")
            break

    matching.sort(key=lambda r: r["Created_x0020_Date."])
    return matching


def print_listing(matching):
    if not matching:
        print("No matching 'Daily José' / 'Daily Jose' recordings found.")
        return
    print(f"\nFound {len(matching)} matching recording(s):\n")
    for i, row in enumerate(matching, 1):
        name = row.get("FileLeafRef") or row.get("Title") or "(unknown)"
        created = row["Created_x0020_Date."]
        print(f"  [{i:3d}] {created}  {name}")
    print()


def parse_range(range_str, total):
    """Parse '1-5' or '3' or '1,3,7' into a list of 0-indexed positions."""
    indices = set()
    for part in range_str.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            for i in range(int(a), int(b) + 1):
                indices.add(i - 1)
        else:
            indices.add(int(part) - 1)
    out = sorted(i for i in indices if 0 <= i < total)
    skipped = sorted(i + 1 for i in indices if not (0 <= i < total))
    if skipped:
        print(f"⚠️  Out-of-range indices ignored: {skipped}")
    return out


def download_and_upload_vtt(row):
    """Download the recording's transcript VTT, cache it locally, and upload
    it to SharePoint. Returns the local cache path or None if no transcript."""
    unique_id = row["UniqueId"].strip("{}")
    name = row.get("FileLeafRef") or row.get("Title") or "(unknown)"

    # UniqueId → itemId
    r = requests.get(
        f"{PERSONAL_BASE}/v2.1/drives/{DRIVE_ID}/items/{unique_id}",
        headers=HEADERS,
    )
    item_id = r.json()["id"]

    # Get transcript URL
    r = requests.get(
        f"{PERSONAL_BASE}/v2.1/drives/{DRIVE_ID}/items/{item_id}/media/transcripts",
        headers=HEADERS,
    )
    payload = r.json()
    if not payload.get("value"):
        print(f"⚠️  No transcript available for: {name}")
        return None

    transcript = payload["value"][0]
    download_url = transcript["temporaryDownloadUrl"]
    ts = datetime.strptime(
        row["Created_x0020_Date."], "%Y-%m-%dT%H:%M:%SZ"
    ).strftime("%Y%m%d_%H%M%S")
    base_name = "_".join(transcript["displayName"].replace(".json", "").split())
    vtt_filename = f"{base_name}_{ts}.vtt"
    cache_path = os.path.join(VTT_CACHE_DIR, vtt_filename)

    if os.path.exists(cache_path):
        print(f"📦 Cached, skipping download/upload: {vtt_filename}")
        return cache_path

    # Download VTT
    r = requests.get(download_url + "&is=1", headers={"Authorization": f"Bearer {BEARER}"})
    vtt_content = r.content
    with open(cache_path, "wb") as f:
        f.write(vtt_content)

    # Upload VTT to SharePoint
    requests.post(
        f"{SYSTEMS_BASE}/web/GetFolderByServerRelativeUrl('{VTT_FOLDER}')"
        f"/Files/add(url='{vtt_filename}',overwrite=true)",
        headers={**HEADERS, "Content-Type": "application/octet-stream"},
        data=vtt_content,
    )
    print(f"⬆️  Uploaded VTT: {vtt_filename}")
    return cache_path


def main():
    parser = argparse.ArgumentParser(
        description="Fetch all 'Daily José' / 'Daily Jose' recordings from OneDrive."
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List all matching recordings (default if no other action given).",
    )
    parser.add_argument(
        "--range",
        type=str,
        help="Indices to process, e.g. '1-5', '3', or '1,3,7'.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Process every matching recording.",
    )
    args = parser.parse_args()

    matching = list_daily_recordings()

    # Default action is --list when no range/all is given
    if args.list or (not args.range and not args.all):
        print_listing(matching)
        return 0

    if not matching:
        print("No matching recordings to process.")
        return 0

    if args.all:
        indices = list(range(len(matching)))
    else:
        indices = parse_range(args.range, len(matching))

    if not indices:
        print("No valid indices selected.")
        return 1

    os.makedirs(VTT_CACHE_DIR, exist_ok=True)

    print(f"\nProcessing {len(indices)} recording(s)...\n")
    manifest = []
    for idx in indices:
        row = matching[idx]
        cache_path = download_and_upload_vtt(row)
        if cache_path:
            manifest.append(os.path.basename(cache_path))

    if not manifest:
        print("No VTTs cached; nothing for ETL to do.")
        return 1

    # Write manifest of files in this batch so ETL knows what to process.
    with open(MANIFEST_PATH, "w") as f:
        json.dump(manifest, f, indent=2)
    print(f"\n📝 Wrote manifest with {len(manifest)} file(s): {MANIFEST_PATH}\n")

    # Hand off to ETL
    print("--- Running etl_date_range.py ---\n")
    result = subprocess.run(
        [sys.executable, os.path.join(SCRIPTS_DIR, "etl_date_range.py")],
    )
    return result.returncode


if __name__ == "__main__":
    sys.exit(main())
