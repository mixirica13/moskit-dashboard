"""Testa expressões de correspondência exata para custom fields."""
import requests
import json

API_BASE = "https://api.ms.prod.moskit.services/v2"
api_key = input("Cole sua API Key do Moskit: ").strip()
headers = {"apikey": api_key, "Content-Type": "application/json"}

CF_KEY = "CF_VRAqd1idCjgvPMbL"
VALOR = "[C014][CBO][LEADS][CRIATIVOS-22-01][USD]"

# Primeiro: ver se contains acha esse valor
print(f"\n=== contains com valor completo ===")
body = [{"field": CF_KEY, "expression": "contains", "values": [VALOR]}]
resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                     params={"quantity": 3}, timeout=30)
print(f"Status: {resp.status_code}, Deals: {len(resp.json()) if resp.status_code == 200 else resp.text[:200]}")

# Tentar com partes do valor
partes = ["C014", "CBO", "LEADS", "CRIATIVOS", "USD", "CRIATIVOS-22-01"]
for p in partes:
    body = [{"field": CF_KEY, "expression": "contains", "values": [p]}]
    resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                         params={"quantity": 3}, timeout=30)
    count = len(resp.json()) if resp.status_code == 200 else 0
    print(f"  contains '{p}' -> {count} deal(s)")

# Ver valores reais nos ultimos deals
print(f"\n=== Valores reais de campaign_name nos ultimos deals ===")
body3 = [{"field": "dateCreated", "expression": "gt", "values": ["2026-01-01T00:00:00Z"]}]
resp3 = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body3,
                      params={"quantity": 10}, timeout=30)
if resp3.status_code == 200:
    for d in resp3.json():
        for ecf in d.get("entityCustomFields", []):
            if ecf.get("id") == CF_KEY:
                print(f"  Deal {d.get('id')}: '{ecf.get('textValue', '')}'")
                break
