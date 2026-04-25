import re
import json
import html
import sys
import requests
from pathlib import Path

SCRIPTS_DIR = Path(__file__).parent
PEOPLE_FILE = SCRIPTS_DIR / "people.json"
VTT_CACHE_DIR = SCRIPTS_DIR / ".vtt_cache"
MANIFEST_PATH = VTT_CACHE_DIR / "manifest.json"

SYSTEMS_BASE = "https://ammper.sharepoint.com/sites/Systems/_api"
TXT_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseTXT"


def load_people():
    with open(PEOPLE_FILE) as f:
        return json.load(f)


def resolve_name(raw_name, people):
    decoded = html.unescape(raw_name.strip())
    if decoded in people:
        return people[decoded]["alias"]
    # Fallback (shouldn't happen after pre-flight, but kept for safety)
    return decoded.split()[0]


def extract_speakers(vtt_bytes):
    """Return the set of every distinct <v Speaker> name in this VTT."""
    text = vtt_bytes.decode("utf-8-sig")
    matches = re.findall(r"<v ([^>]+)>", text)
    return {html.unescape(m.strip()) for m in matches}


def vtt_to_txt(vtt_bytes, people):
    text = vtt_bytes.decode("utf-8-sig")
    blocks = re.split(r"\n\s*\n", text.strip())
    lines = []
    current_speaker = None
    current_text = []

    for block in blocks:
        block = block.strip()
        if block == "WEBVTT":
            continue
        block_lines = block.splitlines()
        if not any("-->" in l for l in block_lines):
            continue

        content_lines = []
        past_timestamp = False
        for l in block_lines:
            if "-->" in l:
                past_timestamp = True
                continue
            if past_timestamp:
                content_lines.append(l.strip())

        content = " ".join(content_lines)
        match = re.match(r"<v ([^>]+)>(.*?)</v>\s*$", content, re.DOTALL)
        if not match:
            continue

        speaker = resolve_name(match.group(1), people)
        speech = html.unescape(match.group(2).strip())
        speech = " ".join(speech.split())

        if speaker == current_speaker:
            current_text.append(speech)
        else:
            if current_speaker:
                lines.append(f"{current_speaker}: {' '.join(current_text)}")
            current_speaker = speaker
            current_text = [speech]

    if current_speaker:
        lines.append(f"{current_speaker}: {' '.join(current_text)}")

    return "\n".join(lines)


def upload_txt(filename, txt_content, bearer):
    headers = {"Authorization": f"Bearer {bearer}", "Accept": "application/json"}
    requests.post(
        f"{SYSTEMS_BASE}/web/GetFolderByServerRelativeUrl('{TXT_FOLDER}')"
        f"/Files/add(url='{filename}',overwrite=true)",
        headers={**headers, "Content-Type": "application/octet-stream"},
        data=txt_content.encode("utf-8"),
    )
    print(f"⬆️  Uploaded TXT: {filename}")


def load_batch_files():
    """Get the list of VTT cache paths to process for this batch.

    Prefers manifest.json (written by fetch_date_range.py); falls back to
    every .vtt in the cache dir if no manifest exists."""
    if MANIFEST_PATH.exists():
        with open(MANIFEST_PATH) as f:
            names = json.load(f)
        return [VTT_CACHE_DIR / n for n in names if (VTT_CACHE_DIR / n).exists()]
    return sorted(VTT_CACHE_DIR.glob("*.vtt"))


def main():
    if not VTT_CACHE_DIR.exists():
        print("No VTT cache found. Run fetch_date_range.py first.")
        return 1

    vtt_files = load_batch_files()
    if not vtt_files:
        print("No VTT files in batch. Run fetch_date_range.py first.")
        return 1

    people = load_people()

    # ---- Pre-flight: scan every VTT for unknown speakers ----
    all_speakers = set()
    speaker_to_files = {}
    for vtt_path in vtt_files:
        with open(vtt_path, "rb") as f:
            vtt_bytes = f.read()
        speakers = extract_speakers(vtt_bytes)
        all_speakers |= speakers
        for s in speakers:
            speaker_to_files.setdefault(s, []).append(vtt_path.name)

    missing = sorted(s for s in all_speakers if s not in people)

    if missing:
        print(f"\n⚠️  {len(missing)} speaker(s) missing from people.json:\n")
        for name in missing:
            files = speaker_to_files[name]
            preview = ", ".join(files[:3]) + ("..." if len(files) > 3 else "")
            print(f"  • {name}")
            print(f"      seen in: {preview}")
        print(
            f"\nAdd these entries to {PEOPLE_FILE}, then re-run:\n"
            f"  python etl_date_range.py\n"
            f"\nNo TXT was uploaded — the whole batch is blocked until "
            f"every speaker resolves."
        )
        return 1

    print(
        f"✓ All {len(all_speakers)} unique speaker(s) resolved against "
        f"people.json. Processing {len(vtt_files)} file(s)...\n"
    )

    try:
        from vars import BEARER
    except ImportError:
        print("Cannot import BEARER from vars.py. Run get_token.py first.")
        return 1

    for vtt_path in vtt_files:
        with open(vtt_path, "rb") as f:
            vtt_bytes = f.read()
        txt_content = vtt_to_txt(vtt_bytes, people)
        txt_filename = vtt_path.name.replace(".vtt", ".txt")
        upload_txt(txt_filename, txt_content, BEARER)

    print(f"\n✓ Done. Uploaded {len(vtt_files)} TXT file(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
