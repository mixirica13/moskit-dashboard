"""Testa diferentes formatos de condição na API do Moskit."""
import requests
import json
import sys

API_BASE = "https://api.ms.prod.moskit.services/v2"

api_key = input("Cole sua API Key do Moskit: ").strip()
headers = {"apikey": api_key, "Content-Type": "application/json"}

# Testar diferentes formatos de condição
testes = [
    # Teste 1: GREATER_EQUALS + LESS_EQUALS
    {
        "desc": "GREATER_EQUALS + LESS_EQUALS com field",
        "body": [
            {"field": "dateCreated", "expression": "GREATER_EQUALS", "values": ["2026-01-01T00:00:00Z"]},
            {"field": "dateCreated", "expression": "LESS_EQUALS", "values": ["2026-03-16T23:59:59Z"]},
        ]
    },
    # Teste 2: >= e <=
    {
        "desc": ">= e <=",
        "body": [
            {"field": "dateCreated", "expression": ">=", "values": ["2026-01-01T00:00:00Z"]},
            {"field": "dateCreated", "expression": "<=", "values": ["2026-03-16T23:59:59Z"]},
        ]
    },
    # Teste 3: gt/lt
    {
        "desc": "gt e lt",
        "body": [
            {"field": "dateCreated", "expression": "gt", "values": ["2026-01-01T00:00:00Z"]},
            {"field": "dateCreated", "expression": "lt", "values": ["2026-03-16T23:59:59Z"]},
        ]
    },
    # Teste 4: value em vez de values
    {
        "desc": "BETWEEN com value (singular)",
        "body": [
            {"field": "dateCreated", "expression": "BETWEEN", "value": ["2026-01-01T00:00:00Z", "2026-03-16T23:59:59Z"]},
        ]
    },
    # Teste 5: between minúsculo com value
    {
        "desc": "between minusculo com value singular",
        "body": [
            {"field": "dateCreated", "expression": "between", "value": "2026-01-01T00:00:00Z,2026-03-16T23:59:59Z"},
        ]
    },
    # Teste 6: equals simples para ver se funciona
    {
        "desc": "EQUALS simples no campo name",
        "body": [
            {"field": "name", "expression": "EQUALS", "values": ["teste"]},
        ]
    },
    # Teste 7: equals com value singular
    {
        "desc": "EQUALS com value singular",
        "body": [
            {"field": "name", "expression": "EQUALS", "value": "teste"},
        ]
    },
    # Teste 8: eq
    {
        "desc": "eq com value",
        "body": [
            {"field": "name", "expression": "eq", "value": "teste"},
        ]
    },
]

for t in testes:
    print(f"\n=== Teste: {t['desc']} ===")
    print(f"Body: {json.dumps(t['body'])}")
    try:
        resp = requests.post(
            f"{API_BASE}/deals/search",
            headers=headers,
            json=t["body"],
            params={"quantity": 1},
            timeout=30,
        )
        print(f"Status: {resp.status_code}")
        if resp.status_code == 200:
            data = resp.json()
            print(f"SUCESSO! Retornou {len(data)} deal(s)")
            if data:
                print(f"Primeiro: {json.dumps(data[0], indent=2, ensure_ascii=False)[:200]}")
            break
        else:
            print(f"Erro: {resp.text[:200]}")
    except Exception as e:
        print(f"Exceção: {e}")
