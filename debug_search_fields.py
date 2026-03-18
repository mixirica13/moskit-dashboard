"""Mostra todos os campos pesquisáveis, incluindo custom fields."""
import requests
import json

API_BASE = "https://api.ms.prod.moskit.services/v2"
api_key = input("Cole sua API Key do Moskit: ").strip()
headers = {"apikey": api_key, "Content-Type": "application/json"}

resp = requests.get(f"{API_BASE}/deals/search", headers=headers, timeout=30)
campos = resp.json()

print(f"Total: {len(campos)}\n")

# Mostrar campos que parecem ser custom fields (CF_)
print("=== Campos personalizados (pesquisáveis) ===\n")
for c in campos:
    if c.get("key", "").startswith("CF_") or c.get("type") in ("CUSTOM_FIELD",):
        print(json.dumps(c, indent=2, ensure_ascii=False))
        print("---")

# Mostrar todos os tipos disponíveis
print("\n=== Tipos de campos encontrados ===")
tipos = set(c.get("type") for c in campos)
print(tipos)

# Buscar campos que tenham 'campaign' no nome
print("\n=== Campos com 'campaign' no nome ===")
for c in campos:
    if "campaign" in c.get("name", "").lower() or "campaign" in c.get("key", "").lower():
        print(json.dumps(c, indent=2, ensure_ascii=False))
