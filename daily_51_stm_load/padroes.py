import pandas as pd

def padrao_18(valor) -> str:
    """
    Converte valor para string de 18 dígitos com zeros à esquerda.
    Remove casas decimais caso venha como float.
    """
    if pd.isna(valor):
        return ""

    valor_limpo = str(valor).split(".")[0]
    return valor_limpo.zfill(18)
