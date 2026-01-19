import os
import json
import logging
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
import openpyxl  # noqa


# ======================================================
# ENV
# ======================================================
ROOT_DIR = Path(__file__).resolve().parent
load_dotenv(ROOT_DIR / ".env")

REQUIRED_VARS = [
    "MANIA_BASE_URL", "MANIA_CARD_ID", "MANIA_SHEET_TAB",
    "AMAZONET_BASE_URL", "AMAZONET_CARD_ID", "AMAZONET_SHEET_TAB",
    "GOOGLE_PROJECT_ID", "GOOGLE_PRIVATE_KEY_ID", "GOOGLE_PRIVATE_KEY",
    "GOOGLE_CLIENT_EMAIL", "GOOGLE_CLIENT_ID", "GOOGLE_SPREADSHEET_ID",
]

for var in REQUIRED_VARS:
    if not os.getenv(var):
        raise RuntimeError(f"Variável obrigatória não carregada: {var}")


# ======================================================
# LOGGING
# ======================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | metabase_gsheet | %(message)s",
)
logger = logging.getLogger("metabase_gsheet")


# ======================================================
# DATACLASSES
# ======================================================
@dataclass(frozen=True)
class MetabaseConfig:
    name: str
    base_url: str
    card_id: str
    sheet_tab: str


@dataclass(frozen=True)
class GoogleSheetsConfig:
    project_id: str
    private_key_id: str
    private_key: str
    client_email: str
    client_id: str
    spreadsheet_id: str


# ======================================================
# HELPERS
# ======================================================
def env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise EnvironmentError(f"Variável obrigatória não definida: {name}")
    return value


def get_metabase_config(empresa: str) -> MetabaseConfig:
    empresa = empresa.upper()
    return MetabaseConfig(
        name=empresa,
        base_url=env(f"{empresa}_BASE_URL"),
        card_id=env(f"{empresa}_CARD_ID"),
        sheet_tab=env(f"{empresa}_SHEET_TAB"),
    )


def get_gs_config() -> GoogleSheetsConfig:
    return GoogleSheetsConfig(
        project_id=env("GOOGLE_PROJECT_ID"),
        private_key_id=env("GOOGLE_PRIVATE_KEY_ID"),
        private_key=env("GOOGLE_PRIVATE_KEY").replace("\\n", "\n"),
        client_email=env("GOOGLE_CLIENT_EMAIL"),
        client_id=env("GOOGLE_CLIENT_ID"),
        spreadsheet_id=env("GOOGLE_SPREADSHEET_ID"),
    )


# ======================================================
# METABASE
# ======================================================
def build_url(cfg: MetabaseConfig, inicio: str, fim: str) -> str:
    if cfg.name == "MANIA":
        return (
            f"{cfg.base_url}/public/question/{cfg.card_id}.xlsx"
            f"?data_inicio={inicio}&data_fim={fim}"
        )

    # AMAZONET
    parameters = [
        {
            "type": "date/single",
            "value": inicio,
            "target": ["variable", ["template-tag", "datainicio"]],
        },
        {
            "type": "date/single",
            "value": fim,
            "target": ["variable", ["template-tag", "datafim"]],
        },
    ]

    return (
        f"{cfg.base_url}/api/public/card/{cfg.card_id}/query/xlsx"
        f"?parameters={quote(json.dumps(parameters))}"
    )


def extract_metabase(cfg: MetabaseConfig, inicio: str, fim: str) -> pd.DataFrame:
    url = build_url(cfg, inicio, fim)
    logger.info(f"[{cfg.name}] Endpoint | {url}")

    r = requests.get(url, timeout=60)
    r.raise_for_status()

    if r.content[:2] != b"PK":
        raise ValueError(f"[{cfg.name}] Retorno não é XLSX válido")

    df = pd.read_excel(BytesIO(r.content), engine="openpyxl")
    logger.info(f"[{cfg.name}] Extração concluída | linhas={len(df)}")

    return df


# ======================================================
# GOOGLE SHEETS
# ======================================================
def connect_gs(cfg: GoogleSheetsConfig):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(
        {
            "type": "service_account",
            "project_id": cfg.project_id,
            "private_key_id": cfg.private_key_id,
            "private_key": cfg.private_key,
            "client_email": cfg.client_email,
            "client_id": cfg.client_id,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
            "client_x509_cert_url": f"https://www.googleapis.com/robot/v1/metadata/x509/{cfg.client_email}",
        },
        scope,
    )
    return gspread.authorize(credentials)


def update_sheet(df: pd.DataFrame, aba: str):
    if df.empty:
        logger.warning(f"[{aba}] DataFrame vazio. Ignorado.")
        return

    gs_cfg = get_gs_config()
    gc = connect_gs(gs_cfg)

    ws = gc.open_by_key(gs_cfg.spreadsheet_id).worksheet(aba)
    ws.clear()
    ws.update([df.columns.tolist()] + df.values.tolist())

    logger.info(f"[{aba}] Google Sheets atualizado")


# ======================================================
# MAIN
# ======================================================
def main():
    inicio_exec = datetime.now()
    logger.info("🚀 INÍCIO")

    data_fim = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    data_inicio = (datetime.today() - timedelta(days=60)).strftime("%Y-%m-%d")

    for empresa in ("MANIA", "AMAZONET"):
        try:
            cfg = get_metabase_config(empresa)
            df = extract_metabase(cfg, data_inicio, data_fim)

            # ===============================
            # REGRAS POR EMPRESA
            # ===============================
            if cfg.name == "MANIA":
                # Coluna A → string simples
                df.iloc[:, 0] = df.iloc[:, 0].astype(str)

                # Coluna D → string com padding (18 dígitos)
                df.iloc[:, 3] = df.iloc[:, 3].apply(
                    lambda x: str(x).zfill(18) if pd.notna(x) else x
                )

            elif cfg.name == "AMAZONET":
                # Coluna A → MESMA REGRA da coluna D da MANIA
                df.iloc[:, 0] = df.iloc[:, 0].apply(
                    lambda x: str(x).zfill(18) if pd.notna(x) else x
                )

            update_sheet(df, cfg.sheet_tab)

        except Exception:
            logger.exception(f"❌ Erro ao processar {empresa}")

    logger.info(f"✅ FIM | Duração: {datetime.now() - inicio_exec}")


if __name__ == "__main__":
    main()
