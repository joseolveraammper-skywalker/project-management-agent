import asyncio
import subprocess
import os
from playwright.async_api import async_playwright

GRABACIONES_URL = "https://ammper-my.sharepoint.com/personal/jolvera_ammper_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fjolvera%5Fammper%5Fcom%2FDocuments%2FGrabaciones"
SESSION_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "session.json")
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))

bearer_token = None

def handle_request(request):
    global bearer_token
    auth = request.headers.get("authorization", "")
    if (auth.startswith("Bearer eyJhbGciOiJSUzI1Ni") and 
        "ammper-my.sharepoint.com" in request.url and
        "_api" in request.url):
        bearer_token = auth.replace("Bearer ", "")

async def main():
    global bearer_token
    async with async_playwright() as p:
        browser = await p.chromium.launch(channel="chrome", headless=False)

        ctx_kwargs = {"storage_state": SESSION_FILE} if os.path.exists(SESSION_FILE) else {}
        context = await browser.new_context(**ctx_kwargs)

        page = await context.new_page()
        page.on("request", handle_request)
        await page.goto(GRABACIONES_URL)
        #await asyncio.sleep(10)

        # Poll up to 60s for a bearer token (covers manual login time)
        for _ in range(60):
            if bearer_token:
                break
            await asyncio.sleep(1)

        await context.storage_state(path=SESSION_FILE)
        await browser.close()
        

asyncio.run(main())

if not bearer_token:
    raise RuntimeError("Could not extract Bearer token — did the page load and authenticate?")

with open(os.path.join(SCRIPTS_DIR, "vars.py"), "w") as f:
    f.write(f'BEARER = "{bearer_token}"\n')

print("Token extracted, running fetch script...")
subprocess.run(["python", os.path.join(SCRIPTS_DIR, "fetch.py")], check=True)
