"""Script para testar a API do Moskit e ver os campos disponíveis."""
import requests
import json
import sys

API_BASE = "https://api.ms.prod.moskit.services/v2"

api_key = input("Cole sua API Key do Moskit: ").strip()

headers = {"apikey": api_key, "Content-Type": "application/json"}

print("\n=== Buscando campos pesquisáveis de deals ===\n")
resp = requests.get(f"{API_BASE}/deals/search", headers=headers, timeout=30)

if resp.status_code != 200:
    print(f"Erro: {resp.status_code} - {resp.text}")
    sys.exit(1)

campos = resp.json()
print(f"Total de campos: {len(campos)}\n")

# Mostrar primeiros 5 campos completos
print("=== Primeiros 5 campos (estrutura completa) ===\n")
for c in campos[:5]:
    print(json.dumps(c, indent=2, ensure_ascii=False))
    print("---")

# Filtrar campos de data
print("\n=== Campos que parecem ser de data ===\n")
for c in campos:
    field = str(c.get("field", c.get("name", c.get("id", "")))).lower()
    tipo = str(c.get("type", "")).lower()
    if "date" in field or "date" in tipo or "data" in field:
        print(json.dumps(c, indent=2, ensure_ascii=False))
        print("---")
