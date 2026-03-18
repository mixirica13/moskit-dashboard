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

# Testar expressoes para campo status
print("=== Testando expressoes para status ===")
for expr in ["eq", "match", "contains", "is", "equals"]:
    try:
        conditions = [{"field": "status", "expression": expr, "values": ["LOST"]}]
        resp = requests.post(f"{BASE}/deals/search", headers=headers, json=conditions, params={"quantity": 1, "start": 0})
        if resp.status_code == 200:
            deals = resp.json()
            print(f"  SUCESSO: '{expr}' -> {len(deals)} deal(s)")
            break
        else:
            print(f"  FALHOU: '{expr}' -> {resp.status_code}")
    except Exception as e:
        print(f"  ERRO: '{expr}' -> {e}")

# Buscar deals recentes e filtrar os perdidos
print("\n=== Buscando deals recentes para achar perdidos ===")
conditions = [
    {"field": "dateCreated", "expression": "gt", "values": ["2025-01-01T00:00:00Z"]},
    {"field": "dateCreated", "expression": "lt", "values": ["2026-03-17T23:59:59Z"]}
]
resp = requests.post(f"{BASE}/deals/search", headers=headers, json=conditions, params={"quantity": 50, "start": 0})
print(f"Status: {resp.status_code}")
if resp.status_code == 200:
    deals = resp.json()
    print(f"Total deals: {len(deals)}")

    perdidos = [d for d in deals if d.get("status") == "LOST"]
    print(f"Perdidos: {len(perdidos)}")

    for d in perdidos[:3]:
        print(f"\n=== Deal PERDIDO {d.get('id')}: {d.get('name')} ===")
        print(f"  status: {d.get('status')}")
        print(f"  stage: {json.dumps(d.get('stage'), ensure_ascii=False)}")
        print(f"  Chaves: {list(d.keys())}")
        # Mostrar campos que possam ter info de stage anterior
        for key in sorted(d.keys()):
            val = d[key]
            if isinstance(val, (str, int, float, bool)) or val is None:
                kl = key.lower()
                if any(x in kl for x in ['stage', 'lost', 'prev', 'last', 'reason', 'phase']):
                    print(f"  {key}: {val}")

    # Mostrar um deal ativo para comparar
    ativos = [d for d in deals if d.get("status") != "LOST"]
    if ativos:
        d = ativos[0]
        print(f"\n=== Deal ATIVO {d.get('id')}: {d.get('name')} ===")
        print(f"  status: {d.get('status')}")
        print(f"  stage: {json.dumps(d.get('stage'), ensure_ascii=False)}")
else:
    print(resp.text)
