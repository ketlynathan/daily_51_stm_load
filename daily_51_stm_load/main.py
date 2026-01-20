import subprocess
import sys
import logging
from pathlib import Path

# ======================================================
# LOGGING
# ======================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | main_runner | %(message)s"
)
logger = logging.getLogger("main_runner")

# ======================================================
# FUNÇÃO PARA RODAR SCRIPT
# ======================================================
def run_script(script_path: Path):
    logger.info(f"🚀 Iniciando {script_path.name}")
    try:
        # Executa o script Python
        result = subprocess.run(
            [sys.executable, str(script_path)],
            check=True,
            capture_output=True,
            text=True
        )
        logger.info(f"✅ Concluído {script_path.name}")
        logger.info(result.stdout)
    except subprocess.CalledProcessError as e:
        logger.error(f"❌ Erro ao executar {script_path.name}")
        logger.error(e.stdout)
        logger.error(e.stderr)
        sys.exit(1)  # encerra o main se algum script falhar

# ======================================================
# FUNÇÃO PARA RODAR EXECUTÁVEL
# ======================================================
def run_executable(exe_path: Path):
    logger.info(f"🚀 Iniciando executável {exe_path.name}")
    try:
        result = subprocess.run(
            [str(exe_path)],
            check=True,
            capture_output=True,
            text=True
        )
        logger.info(f"✅ Concluído {exe_path.name}")
        logger.info(result.stdout)
    except subprocess.CalledProcessError as e:
        logger.error(f"❌ Erro ao executar {exe_path.name}")
        logger.error(e.stdout)
        logger.error(e.stderr)
        sys.exit(1)

# ======================================================
# MAIN
# ======================================================
def main():
    root_dir = Path(__file__).parent

    # Scripts Python
    script_stm = root_dir / "script_stm.py"
    script_mao = root_dir / "script_mao.py"
    script_60  = root_dir / "script_60.py"  # ou pode ser um .exe

    # Rodando scripts em sequência
    run_script(script_stm)
    run_script(script_mao)

    # Se script_60 for Python
    run_script(script_60)

    # Se script_60 for executável, descomente:
    # run_executable(script_60)

    logger.info("🎉 Todos os scripts executados com sucesso!")

# ======================================================
# ENTRY POINT
# ======================================================
if __name__ == "__main__":
    main()
