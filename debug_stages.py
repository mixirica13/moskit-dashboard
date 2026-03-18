import requests, json, os

env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(env_path):
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                os.environ.setdefault(key.strip(), val.strip())

API_KEY = os.environ.get("MOSKIT_API_KEY", "")
BASE = "https://api.ms.prod.moskit.services/v2"
headers = {"apikey": API_KEY, "Content-Type": "application/json"}

resp = requests.get(f"{BASE}/stages", headers=headers, params={"quantity": 200, "start": 0})
print(f"Status: {resp.status_code}")
stages = resp.json()
print(f"Total stages: {len(stages)}")
for s in stages:
    print(f"  ID: {s['id']}, Nome: {s.get('name','?')}, Pipeline: {s.get('pipeline',{}).get('id','?')}")
