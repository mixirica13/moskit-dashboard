"""Testa correspondência exata e variações com colchetes."""
import requests
import json

API_BASE = "https://api.ms.prod.moskit.services/v2"
api_key = input("Cole sua API Key do Moskit: ").strip()
headers = {"apikey": api_key, "Content-Type": "application/json"}

CF_KEY = "CF_VRAqd1idCjgvPMbL"

# Primeiro: pegar um valor real que sabemos que existe
print("=== Buscando um deal com C017 para pegar valor exato ===")
body = [{"field": CF_KEY, "expression": "contains", "values": ["C017"]}]
resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                     params={"quantity": 1}, timeout=30)
if resp.status_code != 200 or not resp.json():
    print("Nao encontrou deal com C017")
    exit()

deal = resp.json()[0]
valor_real = ""
for ecf in deal.get("entityCustomFields", []):
    if ecf.get("id") == CF_KEY:
        valor_real = ecf.get("textValue", "")
        break

print(f"Valor real encontrado: '{valor_real}'")

# Testar contains com o valor exato
print(f"\n=== contains com valor exato '{valor_real}' ===")
body = [{"field": CF_KEY, "expression": "contains", "values": [valor_real]}]
resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                     params={"quantity": 3}, timeout=30)
print(f"Resultado: {len(resp.json()) if resp.status_code == 200 else resp.text[:200]}")

# Testar sem colchetes
sem_colchetes = valor_real.replace("[", "").replace("]", "")
print(f"\n=== contains sem colchetes '{sem_colchetes}' ===")
body = [{"field": CF_KEY, "expression": "contains", "values": [sem_colchetes]}]
resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                     params={"quantity": 3}, timeout=30)
print(f"Resultado: {len(resp.json()) if resp.status_code == 200 else resp.text[:200]}")

# Testar partes maiores com hifen e colchete
partes = [
    valor_real,
    sem_colchetes,
    valor_real.split("][")[0] + "]",  # primeiro bloco com colchetes
    valor_real.split("][")[0].replace("[",""),  # primeiro bloco sem colchetes
]
# Adicionar partes progressivamente maiores
blocos = valor_real.replace("[", "").replace("]", " ").split()
for i in range(1, len(blocos) + 1):
    partes.append(" ".join(blocos[:i]))

print(f"\n=== Testes progressivos ===")
for p in partes:
    body = [{"field": CF_KEY, "expression": "contains", "values": [p]}]
    resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                         params={"quantity": 3}, timeout=30)
    count = len(resp.json()) if resp.status_code == 200 else f"ERR {resp.status_code}"
    print(f"  '{p}' -> {count}")

# Testar expressoes de match exato
print(f"\n=== Expressoes de match exato com valor real ===")
for expr in ["eq", "is", "match", "equals", "exact", "ct"]:
    body = [{"field": CF_KEY, "expression": expr, "values": [valor_real]}]
    resp = requests.post(f"{API_BASE}/deals/search", headers=headers, json=body,
                         params={"quantity": 3}, timeout=30)
    if resp.status_code == 200:
        print(f"  *** SUCESSO: '{expr}' -> {len(resp.json())} deal(s) ***")
    else:
        print(f"  FALHOU: '{expr}' -> {resp.status_code}")
