import requests
from vars import *
from datetime import datetime
from etl import vtt_to_txt, upload_txt

HEADERS = {
    "Authorization": f"Bearer {BEARER}",
    "Accept": "application/json",
}

DRIVE_ID = "b!RBHDsJ-Kl0WZ6cEyf0jmmagT-TaeirNNiWr7WlVTYAI0Exing_vnT7kZf75UE4qs"
PERSONAL_BASE = "https://ammper-my.sharepoint.com/personal/jolvera_ammper_com/_api"
SYSTEMS_BASE = "https://ammper.sharepoint.com/sites/Systems/_api"
VTT_FOLDER = "/sites/Systems/Documentos compartidos/dailyJose/rawDailyJoseVTT"

today = datetime.now().strftime("%Y-%m-%d")

# Step 1: Find today's recording
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

# Step 2: UniqueId → itemId
r = requests.get(f"{PERSONAL_BASE}/v2.1/drives/{DRIVE_ID}/items/{unique_id}", headers=HEADERS)
item_id = r.json()["id"]

# Step 3: Get transcript download URL
r = requests.get(f"{PERSONAL_BASE}/v2.1/drives/{DRIVE_ID}/items/{item_id}/media/transcripts", headers=HEADERS)
transcript = r.json()["value"][0]
download_url = transcript["temporaryDownloadUrl"]
ts = datetime.strptime(row["Created_x0020_Date."], "%Y-%m-%dT%H:%M:%SZ").strftime("%Y%m%d_%H%M%S")
base_name = transcript["displayName"].replace(".json", "").replace(" ", "_")
vtt_filename = f"{base_name}_{ts}.vtt"
txt_filename = f"{base_name}_{ts}.txt"

# Step 4: Download VTT into memory
r = requests.get(download_url + "&is=1", headers={"Authorization": f"Bearer {BEARER}"})
vtt_content = r.content

# Step 5: Upload VTT to SharePoint
requests.post(
    f"{SYSTEMS_BASE}/web/GetFolderByServerRelativeUrl('{VTT_FOLDER}')/Files/add(url='{vtt_filename}',overwrite=true)",
    headers={**HEADERS, "Content-Type": "application/octet-stream"},
    data=vtt_content,
)
print(f"Uploaded VTT: {vtt_filename}")

# Step 6: ETL in memory → upload TXT
txt_content = vtt_to_txt(vtt_content)
upload_txt(txt_filename, txt_content, BEARER)