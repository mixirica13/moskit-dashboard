"""Mostra a estrutura completa de um deal e lista custom fields."""
import requests
import json

API_BASE = "https://api.ms.prod.moskit.services/v2"
api_key = input("Cole sua API Key do Moskit: ").strip()
headers = {"apikey": api_key, "Content-Type": "application/json"}

# Buscar 1 deal para ver a estrutura
print("\n=== Estrutura de um deal ===\n")
resp = requests.post(
    f"{API_BASE}/deals/search",
    headers=headers,
    json=[{"field": "dateCreated", "expression": "gt", "values": ["2026-01-01T00:00:00Z"]}],
    params={"quantity": 1},
    timeout=30,
)
if resp.status_code == 200:
    deals = resp.json()
    if deals:
        print(json.dumps(deals[0], indent=2, ensure_ascii=False))

# Buscar custom fields
print("\n\n=== Custom Fields cadastrados ===\n")
resp2 = requests.get(f"{API_BASE}/customFields", headers=headers, params={"quantity": 200}, timeout=30)
if resp2.status_code == 200:
    fields = resp2.json()
    print(f"Total: {len(fields)}\n")
    for f in fields[:10]:
        print(json.dumps(f, indent=2, ensure_ascii=False))
        print("---")
