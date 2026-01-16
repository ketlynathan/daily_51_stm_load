import os
import logging
import requests
import pandas as pd
import gspread

from io import StringIO
from dotenv import load_dotenv
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials

# =====================
# LOAD ENV
# =====================
load_dotenv()

def env(var, default=None, required=False):
    value = os.getenv(var, default)
    if required and not value:
        raise EnvironmentError(f"Variável obrigatória não definida: {var}")
    return value

# =====================
# LOGGER SETUP
# =====================
LOG_PATH = env("LOG_PATH", "logs/metabase_gsheet.log")
LOG_LEVEL = env("LOG_LEVEL", "INFO").upper()

os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("metabase_gsheet")

# =====================
# MAIN PROCESS
# =====================
def main():
    logger.info("=== INÍCIO DA EXECUÇÃO ===")

    try:
        # 1️⃣ Datas dinâmicas
        data_inicio = (
            datetime.today() - timedelta(days=int(env("METABASE_DAYS_START", 60)))
        ).strftime("%Y-%m-%d")

        data_fim = (
            datetime.today() - timedelta(days=int(env("METABASE_DAYS_END", 1)))
        ).strftime("%Y-%m-%d")

        logger.info(f"Datas calculadas | início={data_inicio} | fim={data_fim}")

        # 2️⃣ Extrair Metabase
        metabase_url = env("METABASE_REPORT_URL", required=True)
        url = f"{metabase_url}.csv?data_inicio={data_inicio}&data_fim={data_fim}"

        logger.info("Baixando relatório do Metabase")
        response = requests.get(url, timeout=60)
        response.raise_for_status()

        df = pd.read_csv(StringIO(response.text))
        logger.info(f"Relatório carregado | linhas={len(df)}")

        if df.empty:
            logger.warning("Relatório retornou vazio. Processo encerrado.")
            return

        # 3️⃣ Google Sheets
        logger.info("Conectando ao Google Sheets")

        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        creds = ServiceAccountCredentials.from_json_keyfile_name(
            env("GOOGLE_CREDENTIALS_PATH", required=True),
            scope
        )

        client = gspread.authorize(creds)

        sheet = client.open_by_url(
            env("GSHEET_URL", required=True)
        ).worksheet(
            env("GSHEET_TAB_NAME", "Sheet1")
        )

        logger.info("Atualizando planilha")
        sheet.clear()
        sheet.update([df.columns.tolist()] + df.values.tolist())

        logger.info("Planilha atualizada com sucesso")

    except Exception as e:
        logger.exception("Erro crítico durante a execução")
        raise

    finally:
        logger.info("=== FIM DA EXECUÇÃO ===")

# =====================
# ENTRY POINT
# =====================
if __name__ == "__main__":
    main()
