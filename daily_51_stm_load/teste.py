import requests
import json
from urllib.parse import quote
from io import BytesIO
import pandas as pd

# ===============================
# CONFIG FIXA (APENAS TESTE)
# ===============================
BASE_URL = "https://amazonet.hubsoft.com.br:8443"
CARD_ID = "fd9bfef1-920c-47cb-970a-ed5d2c930c18"

DATA_INICIO = "2025-11-20"
DATA_FIM = "2026-01-18"

# IDs reais do Metabase (vindos do DevTools)
PARAMETERS = [
    {
        "type": "date/single",
        "value": DATA_INICIO,
        "id": "2592675a-dee8-eb0e-0eae-977ae51be3e4",
        "target": ["variable", ["template-tag", "datainicio"]],
    },
    {
        "type": "date/single",
        "value": DATA_FIM,
        "id": "84d0a0bf-af9b-8ee5-d601-07e203be6f9a",
        "target": ["variable", ["template-tag", "datafim"]],
    },
]

# ===============================
# BUILD URL
# ===============================
params_encoded = quote(json.dumps(PARAMETERS))

url = (
    f"{BASE_URL}/api/public/card/{CARD_ID}/query/xlsx"
    f"?parameters={params_encoded}"
)

print("URL FINAL:")
print(url)

# ===============================
# REQUEST
# ===============================
response = requests.get(url, timeout=60)
response.raise_for_status()

# ===============================
# VALIDAR SE É XLSX
# ===============================
content = response.content

if content[:2] != b"PK":
    print("❌ NÃO É XLSX")
    print(content[:300])
    exit()

# ===============================
# SALVAR E LER
# ===============================
with open("amazonet_teste.xlsx", "wb") as f:
    f.write(content)

print("✅ XLSX salvo como amazonet_teste.xlsx")

df = pd.read_excel(BytesIO(content), engine="openpyxl")

print("✅ DataFrame carregado")
print(df.head())
print(df.dtypes)
