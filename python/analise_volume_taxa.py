"""
Conferencia independente da base e decomposicao da variacao do rendimento.

Le as colunas brutas da aba Base (nao depende das formulas do Excel) e mostra, por mes:
saldo, rendimento, saldo medio, taxa, efeito volume e efeito taxa, alem de conferencias
de fechamento da movimentacao e de continuidade do saldo.

Uso:  python analise_volume_taxa.py [caminho_da_base.xlsx|.xlsm]
"""
import sys
from pathlib import Path

import pandas as pd

PADRAO = Path(__file__).resolve().parent.parent / "excel" / "base_relatorio_exemplo.xlsx"


def carregar(caminho):
    df = pd.read_excel(caminho, sheet_name="Base", usecols="A:AZ")
    df["Data_Base"] = pd.to_datetime(df["Data_Base"])
    return df


def conferir_movimentacao(df):
    """Saldo final = anterior + aplicacoes - resgates - juros + rendimento (- IR/IOF retido, se saldo liquido)."""
    liquido = (df["Saldo_Liquido_de_IR"] == "S").astype(int)
    calc = (df["Saldo_Anterior"] + df["Aplicacoes"] - df["Resgates"] - df["Juros_Recebidos"]
            + df["Rendimento_Bruto"] - liquido * (df["IR_Retido"] + df["IOF_Retido"]))
    df = df.assign(Diferenca=(df["Saldo_Final"] - calc).round(2))
    return df[df["Diferenca"].abs() > 0.01][["Data_Base", "ID_Aplicacao", "Diferenca"]]


def decompor(df):
    g = df.groupby("Data_Base").agg(
        Saldo_Anterior=("Saldo_Anterior", "sum"), Saldo_Final=("Saldo_Final", "sum"),
        Rendimento=("Rendimento_Bruto", "sum")).sort_index()
    g["Saldo_Medio"] = (g["Saldo_Anterior"] + g["Saldo_Final"]) / 2
    g["Taxa"] = g["Rendimento"] / g["Saldo_Medio"]
    g["Variacao"] = g["Rendimento"].diff()
    g["Efeito_Volume"] = g["Saldo_Medio"].diff() * g["Taxa"].shift(1)
    g["Efeito_Taxa"] = g["Variacao"] - g["Efeito_Volume"]
    g["Continuidade"] = (g["Saldo_Anterior"] - g["Saldo_Final"].shift(1)).round(2)
    return g


def main():
    caminho = Path(sys.argv[1]) if len(sys.argv) > 1 else PADRAO
    df = carregar(caminho)
    print(f"Base: {caminho.name} | {len(df)} linhas | {df['Data_Base'].nunique()} mes(es)\n")

    erros = conferir_movimentacao(df)
    print("Movimentacao:", "OK - todas as linhas fecham" if erros.empty else f"{len(erros)} linha(s) nao fecham")
    if not erros.empty:
        print(erros.to_string(index=False))

    g = decompor(df)
    pd.set_option("display.float_format", lambda v: f"{v:,.2f}")
    print("\nDecomposicao da variacao do rendimento (R$):")
    print(g[["Saldo_Medio", "Rendimento", "Variacao", "Efeito_Volume", "Efeito_Taxa", "Continuidade"]].to_string())
    print("\nTaxa mensal sobre saldo medio:")
    print((g["Taxa"] * 100).map(lambda v: f"{v:.4f}%").to_string())


if __name__ == "__main__":
    main()
