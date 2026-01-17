import os
import logging
import requests
import pandas as pd
from io import BytesIO, StringIO
from datetime import datetime, timedelta
from dataclasses import dataclass
from pathlib import Path

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
import openpyxl  # garante suporte XLSX


# ======================================================
# LOAD ENV
# ======================================================
ROOT_DIR = Path(__file__).resolve().parent
ENV_PATH = ROOT_DIR / ".env"

if not ENV_PATH.exists():
    raise FileNotFoundError(f"Arquivo .env não encontrado em: {ENV_PATH}")

load_dotenv(ENV_PATH)

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
# CONFIGS
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
# ENV HELPERS
# ======================================================
def get_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise EnvironmentError(f"Variável obrigatória não definida: {name}")
    return value


# ======================================================
# FACTORIES
# ======================================================
def get_metabase_config(empresa: str) -> MetabaseConfig:
    empresa = empresa.upper()
    return MetabaseConfig(
        name=empresa,
        base_url=get_env(f"{empresa}_BASE_URL"),
        card_id=get_env(f"{empresa}_CARD_ID"),
        sheet_tab=get_env(f"{empresa}_SHEET_TAB"),
    )


def get_google_sheets_config() -> GoogleSheetsConfig:
    return GoogleSheetsConfig(
        project_id=get_env("GOOGLE_PROJECT_ID"),
        private_key_id=get_env("GOOGLE_PRIVATE_KEY_ID"),
        private_key=get_env("GOOGLE_PRIVATE_KEY").replace("\\n", "\n"),
        client_email=get_env("GOOGLE_CLIENT_EMAIL"),
        client_id=get_env("GOOGLE_CLIENT_ID"),
        spreadsheet_id=get_env("GOOGLE_SPREADSHEET_ID"),
    )


# ======================================================
# CORE FUNCTIONS
# ======================================================
def build_metabase_url(cfg: MetabaseConfig, data_inicio: str, data_fim: str) -> str:
    return (
        f"{cfg.base_url}/public/question/{cfg.card_id}.xlsx"
        f"?data_inicio={data_inicio}&data_fim={data_fim}"
    )


def extrair_metabase(cfg: MetabaseConfig, data_inicio: str, data_fim: str) -> pd.DataFrame:
    url = build_metabase_url(cfg, data_inicio, data_fim)
    logger.info(f"[{cfg.name}] Endpoint | {url}")

    response = requests.get(url, timeout=60)
    response.raise_for_status()

    # --------------------------
    # DEBUG: salvar relatório bruto
    # --------------------------
    raw_file_path = ROOT_DIR / f"{cfg.name}_raw.xlsx"
    with open(raw_file_path, "wb") as f:
        f.write(response.content)
    logger.info(f"[{cfg.name}] Relatório bruto salvo em: {raw_file_path}")

    content = response.content
    is_xlsx = content[:2] == b"PK"  # validação simples XLSX

    try:
        if is_xlsx:
            df = pd.read_excel(BytesIO(content), engine="openpyxl")
        else:
            logger.warning(f"[{cfg.name}] Retorno não é XLSX válido. Fallback para CSV/texto.")
            df = pd.read_csv(StringIO(response.text))
    except Exception:
        logger.exception(f"[{cfg.name}] Falha ao interpretar retorno do Metabase")
        raise

    logger.info(f"[{cfg.name}] Extração concluída | linhas={len(df)} | colunas={len(df.columns)}")
    return df


def conectar_google_sheets(cfg: GoogleSheetsConfig):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    credentials_dict = {
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
    }

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    return gspread.authorize(credentials)


def atualizar_google_sheets(df: pd.DataFrame, aba: str):
    if df.empty:
        logger.warning(f"[{aba}] Relatório vazio. Atualização ignorada.")
        return

    gs_cfg = get_google_sheets_config()
    gc = conectar_google_sheets(gs_cfg)

    sh = gc.open_by_key(gs_cfg.spreadsheet_id)
    worksheet = sh.worksheet(aba)

    worksheet.clear()
    worksheet.update([df.columns.tolist()] + df.values.tolist())

    logger.info(f"[{aba}] Google Sheets atualizado com sucesso")


# ======================================================
# MAIN
# ======================================================
def main():
    inicio = datetime.now()
    logger.info("=" * 38)
    logger.info("🚀 INÍCIO DA EXECUÇÃO")
    logger.info(f"Horário início: {inicio}")
    logger.info("=" * 38)

    data_fim = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    data_inicio = (datetime.today() - timedelta(days=60)).strftime("%Y-%m-%d")
    logger.info(f"Período global | início={data_inicio} | fim={data_fim}")
    logger.info("-" * 38)

    for empresa in ["MANIA", "AMAZONET"]:
        try:
            cfg = get_metabase_config(empresa)
            logger.info(f"🏢 PROCESSANDO EMPRESA: {empresa}")

            df = extrair_metabase(cfg, data_inicio, data_fim)

            # ===============================
            # AJUSTE ESPECÍFICO MANIA
            # ===============================
            if cfg.name == "MANIA" and not df.empty:
                # Coluna A (index 0)
                df.iloc[:, 0] = df.iloc[:, 0].astype(str)
                # Coluna D (index 3)
                df.iloc[:, 3] = df.iloc[:, 3].apply(lambda x: str(x).zfill(18))
                logger.info("[MANIA] Colunas A e D convertidas para string e protegidas contra truncamento")

            atualizar_google_sheets(df, cfg.sheet_tab)

        except Exception:
            logger.exception(f"❌ [{empresa}] Erro durante processamento")

        logger.info("-" * 38)

    fim = datetime.now()
    logger.info("=" * 38)
    logger.info("✅ FIM DA EXECUÇÃO")
    logger.info(f"Horário fim: {fim}")
    logger.info(f"Duração total: {fim - inicio}")
    logger.info("=" * 38)


if __name__ == "__main__":
    main()
