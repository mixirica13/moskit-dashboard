"""Testa diferentes expressões para busca em custom fields."""
import requests
import json

API_BASE = "https://api.ms.prod.moskit.services/v2"
api_key = input("Cole sua API Key do Moskit: ").strip()
headers = {"apikey": api_key, "Content-Type": "application/json"}

CF_KEY = "CF_VRAqd1idCjgvPMbL"  # campaign_name (LONG_TEXT)

expressions = ["eq", "ne", "ct", "sw", "ew", "like", "contains", "has",
               "in", "nn", "nl", "is", "EQ", "NE", "CT", "SW", "EW",
               "LIKE", "CONTAINS", "IN", "NOT_NULL", "IS_NULL",
               "EQUALS", "NOT_EQUALS", "STARTS_WITH", "ENDS_WITH"]

for expr in expressions:
    body = [{"field": CF_KEY, "expression": expr, "values": ["google"]}]
    try:
        resp = requests.post(
            f"{API_BASE}/deals/search",
            headers=headers,
            json=body,
            params={"quantity": 1},
            timeout=15,
        )
        status = resp.status_code
        if status == 200:
            data = resp.json()
            print(f"  *** SUCESSO: expression='{expr}' -> {len(data)} deal(s) ***")
        else:
            msg = resp.text[:120]
            print(f"  FALHOU: expression='{expr}' -> {status}: {msg}")
    except Exception as e:
        print(f"  ERRO: expression='{expr}' -> {e}")
