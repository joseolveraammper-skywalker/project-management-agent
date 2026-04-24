import re
import json
import html
import requests
from pathlib import Path

SCRIPTS_DIR = Path(__file__).parent
PEOPLE_FILE = SCRIPTS_DIR / "people.json"

SYSTEMS_BASE = "https://ammper.sharepoint.com/sites/Systems/_api"
VTT_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseVTT"
TXT_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseTXT"

with open(PEOPLE_FILE) as f:
    people = json.load(f)


def resolve_name(raw_name):
    decoded = html.unescape(raw_name.strip())
    if decoded in people:
        return people[decoded]["alias"]
    return decoded.split()[0]


def vtt_to_txt(vtt_bytes):
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

        speaker = resolve_name(match.group(1))
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
        f"{SYSTEMS_BASE}/web/GetFolderByServerRelativeUrl('{TXT_FOLDER}')/Files/add(url='{filename}',overwrite=true)",
        headers={**headers, "Content-Type": "application/octet-stream"},
        data=txt_content.encode("utf-8"),
    )
    print(f"Uploaded TXT: {filename}")