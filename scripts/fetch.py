import requests
from vars import *
from datetime import datetime

HEADERS = {
    "Authorization": f"Bearer {BEARER}",
    "Accept": "application/json",
}

DRIVE_ID = "b!RBHDsJ-Kl0WZ6cEyf0jmmagT-TaeirNNiWr7WlVTYAI0Exing_vnT7kZf75UE4qs"
PERSONAL_BASE = "https://ammper-my.sharepoint.com/personal/jolvera_ammper_com/_api"
SYSTEMS_BASE = "https://ammper.sharepoint.com/sites/Systems/_api"
UPLOAD_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseVTT"


today = datetime.now().strftime("%Y-%m-%d")

# Step 1: List recordings folder, find today's recording
r = requests.post(
    f"{PERSONAL_BASE}/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream",
    headers={**HEADERS, "Content-Type": "application/json"},
    params={
        "@a1": "'/personal/jolvera_ammper_com/Documents'",
        "RootFolder": "/personal/jolvera_ammper_com/Documents/Grabaciones",
        "TryNewExperienceSingle": "TRUE",
    },
    json={},
)
data = r.json()
if "Row" not in data:
    raise RuntimeError(f"Unexpected response (status {r.status_code}):\n{data}")
rows = data["Row"]
row = next(r for r in rows if r["Created_x0020_Date."].startswith(today))
unique_id = row["UniqueId"].strip("{}")

# Step 2: Convert UniqueId to itemId
r = requests.get(
    f"{PERSONAL_BASE}/v2.1/drives/{DRIVE_ID}/items/{unique_id}",
    headers=HEADERS,
)
item_id = r.json()["id"]

# Step 3: Get transcript download URL
r = requests.get(
    f"{PERSONAL_BASE}/v2.1/drives/{DRIVE_ID}/items/{item_id}/media/transcripts",
    headers=HEADERS,
)
transcript = r.json()["value"][0]
download_url = transcript["temporaryDownloadUrl"]
filename = transcript["displayName"].replace(" ", "_") + ".vtt"

# Step 4: Download transcript content (with speaker names)
r = requests.get(download_url + "&is=1", headers={"Authorization": f"Bearer {BEARER}"})
vtt_content = r.content

# Step 5: Upload to SharePoint Systems site
requests.post(
    f"{SYSTEMS_BASE}/web/GetFolderByServerRelativeUrl('{UPLOAD_FOLDER}')/Files/add(url='{filename}',overwrite=true)",
    headers={**HEADERS, "Content-Type": "application/octet-stream"},
    data=vtt_content,
)

print(f"Done: {filename}")
