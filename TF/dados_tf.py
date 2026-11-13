"""
Dados e cálculos usados pelos gráficos e pela apresentação do TF.

Fontes (pasta dados/):
  resumo_esforcos_maximos_<condição>.csv  resumo exportado do software (pico de cada elemento)
  estatisticas_colab.csv                  tabela "Resumo estatístico geral" do notebook (Colab)
  eventos_cheio_colab.csv                 contagem de picos por evento, condição cheio (Colab)
  ca_eventos_governantes.csv              linhas do CA_394_pontos.csv dos eventos citados

A geometria é calculada aqui com as mesmas coordenadas e fórmulas de
calcular_geometria_amarracao (notebook), com os cabos 11 e 12 no cabeço R.
"""
import os
import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
DADOS = os.path.join(BASE, "dados")

# Condição de carga -> calado [m]
CONDICOES = {"lastro": 9.0, "25%": 10.25, "50%": 11.5, "75%": 12.75, "cheio": 14.0}
ROTULO = {"lastro": "Lastro", "25%": "25%", "50%": "50%", "75%": "75%", "cheio": "Cheio"}

# Cabo de amarração
MBL_KN, DIAMETRO_MM, EA_KN = 736.0, 80.0, 5632.0
MBL_TF = MBL_KN / 9.80665                 # 75,05 tf
LIMITE_MBL, LIMITE_SWL = 55.0, 100.0

# Faixa de maré do CA (394 condições)
MARE_MIN, MARE_MAX = -0.20, 1.60

SECOES = {"Movimentos do Navio": "Movimentos", "Tracao nos Cabos": "Cabos",
          "Reacao nas Defensas": "Reacao", "Compressao nas Defensas": "Compressao",
          "Forca nos Cabecos": "Cabecos"}


def _numero(serie):
    return pd.to_numeric(serie.astype(str).str.extract(r"([-+]?\d+(?:\.\d+)?)")[0], errors="coerce")


def carregar_resumo(condicao):
    """Uma linha por elemento: Secao, Elemento, Valor, Limite, % do Limite, ID e ambiente do pico."""
    caminho = os.path.join(DADOS, f"resumo_esforcos_maximos_{condicao}.csv")
    df = pd.read_csv(caminho, sep=";", encoding="utf-8-sig", dtype=str)
    df.columns = df.columns.str.strip()
    df["Elemento"] = df["Elemento"].str.strip()
    titulo = df["Valor"].isna()
    df["Secao"] = df["Elemento"].where(titulo).ffill().map(SECOES)
    df = df[~titulo].copy()
    for col in ["Valor", "Limite", "V. vento [nos]", "V. corrente [nos]", "Hs [m]", "Tp [s]", "Mare [m]", "ID"]:
        df[col] = _numero(df[col])
    df["ID"] = df["ID"].astype(int)
    df["% do Limite"] = df["Valor"] / df["Limite"] * 100
    return df.reset_index(drop=True)


def carregar_estatisticas():
    return pd.read_csv(os.path.join(DADOS, "estatisticas_colab.csv"), sep=";")


def carregar_eventos_cheio():
    return pd.read_csv(os.path.join(DADOS, "eventos_cheio_colab.csv"), sep=";").set_index("ID")


def carregar_ca_eventos():
    return pd.read_csv(os.path.join(DADOS, "ca_eventos_governantes.csv"), sep=";").set_index("ID")


# ── Geometria (mesmas coordenadas e fórmulas do notebook) ─────────────────────
CABECOS_X = {c: -135 + 15 * i for i, c in enumerate("ABCDEFGHIJKLMNOPQRS")}
Y_CABECO, Z_CABECO, Z_BUZINA = 3.5, 4.6, 18.5
BUZINAS = {2: (-113.2, 5.9), 3: (-113.2, 7.5), 14: (104.8, 9.2), 15: (105.5, 8.8), 17: (111.7, 0.0),
           18: (109.4, -5.3), 23: (73.1, -16.1), 24: (72.0, -16.1), 28: (-56.4, -16.0),
           29: (-86.1, -16.1), 31: (-113.2, -7.5), 32: (-113.2, -5.9)}
CENTRO_Y = 1.4 + 16.2        # defensas a 1,4 m + meia-boca máxima do navio (|Y| máximo das 32 buzinas)
# Cabo do resumo -> buzina; o cabeço vem do nome do cabo
MAPA_CABOS = {"1-S": 2, "2-S": 3, "3-A": 14, "4-A": 15, "5-B": 17, "6-B": 18,
              "7-G": 23, "8-G": 24, "9-N": 28, "10-N": 29, "11-R": 31, "12-R": 32}


def geometria_linha(buzina, cabeco, calado, mare):
    x_local, y_local = BUZINAS[buzina]
    xb = -x_local                                   # navio com proa invertida
    yb = CENTRO_Y - y_local                         # atracação por boreste
    zb = Z_BUZINA - calado + mare
    dx, dy, dz = CABECOS_X[cabeco] - xb, Y_CABECO - yb, Z_CABECO - zb
    horiz = np.hypot(dx, dy)
    return {"comp": float(np.hypot(horiz, dz)), "ang": float(np.degrees(np.arctan(abs(dz) / horiz))),
            "dz_buzina_cabeco": float(zb - Z_CABECO)}


def tabela_geometria(calado, mare):
    linhas = []
    for nome, buzina in MAPA_CABOS.items():
        g = geometria_linha(buzina, nome.split("-")[1], calado, mare)
        linhas.append({"Cabo": int(nome.split("-")[0]), "Nome": nome, "Buzina": buzina,
                       "Cabeco": nome.split("-")[1], **g})
    return pd.DataFrame(linhas)
