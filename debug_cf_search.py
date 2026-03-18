"""Testa busca por custom field SEM filtro de data para verificar se contains funciona."""
import requests
import json

API_BASE = "https://api.ms.prod.moskit.services/v2"
api_key = input("Cole sua API Key do Moskit: ").strip()
headers = {"apikey": api_key, "Content-Type": "application/json"}

valor = input("Digite o valor do campaign_name para buscar: ").strip()

CF_KEY = "CF_VRAqd1idCjgvPMbL"  # campaign_name (LONG_TEXT)

# Teste 1: Somente custom field, sem filtro de data
print(f"\n=== Teste 1: Somente CF contains '{valor}' ===")
body = [{"field": CF_KEY, "expression": "contains", "values": [valor]}]
resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                     params={"quantity": 5}, timeout=30)
print(f"Status: {resp.status_code}, Deals: {len(resp.json()) if resp.status_code == 200 else resp.text[:200]}")
if resp.status_code == 200 and resp.json():
    deal = resp.json()[0]
    cfs = {ecf.get("id"): ecf.get("textValue", "") for ecf in deal.get("entityCustomFields", [])}
    print(f"Deal ID: {deal.get('id')}, Nome: {deal.get('name')}")
    print(f"Custom field {CF_KEY}: '{cfs.get(CF_KEY, 'NAO ENCONTRADO')}'")

# Teste 2: CF + data ampla (ano inteiro)
print(f"\n=== Teste 2: CF contains + data 2025-01-01 a 2026-12-31 ===")
body2 = [
    {"field": CF_KEY, "expression": "contains", "values": [valor]},
    {"field": "dateCreated", "expression": "gt", "values": ["2025-01-01T00:00:00Z"]},
    {"field": "dateCreated", "expression": "lt", "values": ["2026-12-31T23:59:59Z"]},
]
resp2 = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body2,
                      params={"quantity": 5}, timeout=30)
print(f"Status: {resp2.status_code}, Deals: {len(resp2.json()) if resp2.status_code == 200 else resp2.text[:200]}")

# Teste 3: Buscar deals recentes e mostrar os valores de campaign_name
print(f"\n=== Teste 3: Ultimos 5 deals - valores de campaign_name ===")
body3 = [{"field": "dateCreated", "expression": "gt", "values": ["2026-01-01T00:00:00Z"]}]
resp3 = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body3,
                      params={"quantity": 5}, timeout=30)
if resp3.status_code == 200:
    for d in resp3.json():
        cfs = {ecf.get("id"): ecf.get("textValue", "") for ecf in d.get("entityCustomFields", [])}
        print(f"  Deal {d.get('id')}: campaign_name = '{cfs.get(CF_KEY, '')}'")
