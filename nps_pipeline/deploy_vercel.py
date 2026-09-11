"""
deploy_vercel.py — деплой сайта на Vercel v13 через inline files.
"""
import os, json, urllib.request, urllib.error, base64

TOKEN = os.environ["VERCEL_TOKEN"]
TEAM  = os.environ["VERCEL_TEAM_ID"]
PROJ  = os.environ["VERCEL_PROJECT_ID"]
SITE  = os.environ.get("NPS_SITE", os.path.join(os.path.dirname(__file__), "site"))

def api(method, path, body=None):
    url = f"https://api.vercel.com{path}"
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {TOKEN}")
    if body: req.add_header("Content-Type", "application/json")
    try:
        r = urllib.request.urlopen(req, timeout=60)
        return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        return {"HTTP_ERROR": e.code, "body": e.read().decode('utf-8','ignore')}

# Собираем список файлов с содержимым в base64 (v13 API формат)
files = []
for name in ["index.html","data.json","vercel.json"]:
    with open(f"{SITE}/{name}","rb") as f:
        raw = f.read()
    files.append({
        "file": name,
        "data": base64.b64encode(raw).decode(),
        "encoding": "base64",
    })

body = {
    "name": "nps-2026-pik",
    "project": PROJ,
    "files": files,
    "target": "production",
    "projectSettings": {
        "framework": None,
    },
}

print(f"Деплою {len(files)} файлов ({sum(len(f['data']) for f in files)//1024} KB)...")
r = api("POST", f"/v13/deployments?teamId={TEAM}&forceNew=1", body)
if "HTTP_ERROR" in r:
    print("ОШИБКА:", r["HTTP_ERROR"])
    print(r["body"][:2000])
else:
    print(f"Deployment id: {r.get('id')}")
    print(f"URL: https://{r.get('url')}")
    print(f"Alias: {r.get('alias')}")
    print(f"Ready state: {r.get('readyState')}")
    print(f"Target: {r.get('target')}")
