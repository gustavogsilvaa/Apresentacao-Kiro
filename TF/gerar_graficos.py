"""
Gera os gráficos novos da apresentação (pasta imagens/), no mesmo estilo dos gráficos
redesenhados da v7: fonte STIX (desenho de Times), paleta suave e vírgula decimal.

Todos os valores vêm de dados/ (ver dados_tf.py) ou da geometria calculada.
"""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize
from matplotlib.ticker import FuncFormatter

import dados_tf as D

OUT = os.path.join(D.BASE, "imagens")
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family": "STIXGeneral", "mathtext.fontset": "stix", "font.size": 9.5,
    "axes.titlesize": 9.5, "axes.titleweight": "bold", "axes.titlecolor": "#1C2430",
    "axes.edgecolor": "#B9B2A6", "axes.labelcolor": "#3A4250",
    "xtick.color": "#4A5260", "ytick.color": "#4A5260",
    "axes.grid": True, "grid.color": "#ECE7DE", "grid.linewidth": 0.6,
    "axes.spines.top": False, "axes.spines.right": False,
    "figure.facecolor": "white", "axes.facecolor": "white",
    "legend.frameon": False, "legend.fontsize": 7.5,
})
LIM = "#D0574B"      # vermelho suave para limites
ESC = "#2F3A48"      # contorno escuro
AZUL, LARANJA, ROXO = "#7F9CC8", "#F0B27A", "#7A63C2"
FIG = (6.85, 4.6)    # mesma proporção do quadro de imagem nos slides
DPI = 220
CMAP_MARE = plt.get_cmap("coolwarm")
NORM_MARE = Normalize(D.MARE_MIN, D.MARE_MAX)
ORDEM = list(D.CONDICOES)
CALADOS = [D.CONDICOES[c] for c in ORDEM]
ROT_CALADO = [f"{D.ROTULO[c]}\n" + f"{D.CONDICOES[c]:g}".replace(".", ",") + " m" for c in ORDEM]


def br(v, d=1):
    return f"{v:.{d}f}".replace(".", ",").replace("-", "−")


def virgula(eixo, d=None):
    eixo.set_major_formatter(FuncFormatter(
        lambda v, _: (f"{v:g}" if d is None else f"{v:.{d}f}").replace("-", "−").replace(".", ",")))


def save(fig, nome):
    fig.savefig(os.path.join(OUT, nome), dpi=DPI, bbox_inches="tight", pad_inches=0.06)
    plt.close(fig)


R75 = D.carregar_resumo("75%")
EST = D.carregar_estatisticas()


def serie(grandeza, estat):
    t = EST[(EST["Grandeza"] == grandeza) & (EST["Condicao"].isin(ORDEM))].set_index("Condicao")
    return t.loc[ORDEM, estat].to_numpy()


# ─────────────────────────────────────────────────────────────
# Slide 12 · Distribuição dos picos por condição de carga
# ─────────────────────────────────────────────────────────────
def estatistica_por_calado():
    fig, axs = plt.subplots(2, 1, figsize=FIG, layout="constrained", sharex=True)
    for ax, g, lim, cor, titulo, yl in [
            (axs[0], "% MBL", D.LIMITE_MBL, AZUL, "Cabos · picos dos 12 cabos em cada condição", "% MBL"),
            (axs[1], "% SWL", D.LIMITE_SWL, LARANJA, "Cabeços · picos dos 6 cabeços em cada condição", "% SWL")]:
        med, p90, mx = serie(g, "Mediana"), serie(g, "P90"), serie(g, "Maximo")
        ax.fill_between(CALADOS, med, mx, color=cor, alpha=0.22, lw=0, label="mediana a máximo")
        ax.plot(CALADOS, mx, "o-", color=ESC, lw=1.4, ms=4.5, label="máximo")
        ax.plot(CALADOS, p90, "s--", color=cor, lw=1.2, ms=3.5, label="P90")
        ax.plot(CALADOS, med, "o-", color=cor, lw=1.4, ms=4, label="mediana")
        for x, v in zip(CALADOS, mx):
            ax.annotate(br(v, 1) + "%", (x, v), xytext=(0, 5), textcoords="offset points",
                        ha="center", fontsize=7.5, weight="bold", color=ESC)
        ax.axhline(lim, color=LIM, ls="--", lw=1.3, label=f"limite ({lim:g}%)")
        ax.set_ylim(0, lim * 1.15)
        ax.set_ylabel(yl)
        ax.set_title(titulo, loc="left")
        ax.legend(loc="lower right", ncol=5, fontsize=7)
    axs[1].set_xticks(CALADOS, ROT_CALADO)
    save(fig, "12_picos_por_calado.png")


# ─────────────────────────────────────────────────────────────
# Slide 13 · Influência da maré (75%)
# ─────────────────────────────────────────────────────────────
def influencia_mare_75():
    cab = R75[R75["Secao"] == "Cabos"]
    cbc = R75[R75["Secao"] == "Cabecos"]
    fig, axs = plt.subplots(2, 1, figsize=FIG, layout="constrained")
    for ax, d, col, lim, mk, s, titulo, yl, leg, topo in [
            (axs[0], cab, "Valor", D.LIMITE_MBL, "o", 70, "Cabos · pico de tração e maré no instante do pico",
             "Pico (% MBL)", "Limite operacional (55%)", 70),
            (axs[1], cbc, "% do Limite", D.LIMITE_SWL, "^", 85, "Cabeços · pico de esforço e maré no instante do pico",
             "Pico (% SWL)", "Limite estrutural (100%)", 125)]:
        x = np.arange(len(d))
        ax.scatter(x, d[col], c=d["Mare [m]"], cmap=CMAP_MARE, norm=NORM_MARE, s=s, marker=mk,
                   edgecolors=ESC, linewidths=0.5, zorder=3)
        ax.axhline(lim, color=LIM, ls="--", lw=1.3, label=leg)
        i = int(np.argmax(d[col].to_numpy()))
        ax.annotate(br(d[col].iloc[i], 1) + "%", (i, d[col].iloc[i]), xytext=(9, -3),
                    textcoords="offset points", fontsize=7.5, weight="bold", color=ESC)
        ax.set_xticks(x, d["Elemento"])
        ax.set_xlim(-0.5, len(d) - 0.5)
        ax.set_ylim(0, topo)
        ax.set_ylabel(yl)
        ax.set_title(titulo, loc="left")
        ax.legend(loc="upper left")
    cb = fig.colorbar(ScalarMappable(norm=NORM_MARE, cmap=CMAP_MARE), ax=axs, fraction=0.035, pad=0.02)
    virgula(cb.ax.yaxis, 1)
    cb.set_label("Nível da maré [m]")
    cb.outline.set_visible(False)
    save(fig, "13_influencia_mare_75.png")


# ─────────────────────────────────────────────────────────────
# Slide 14 · Eventos governantes (75% e cheio)
# ─────────────────────────────────────────────────────────────
SECOES_EVT = [("Movimentos", "Movimentos", "#8CC7B5"), ("Cabos", "Cabos", "#9DB5DE"),
              ("Reacao", "Reação nas defensas", "#B7DB8F"), ("Compressao", "Compressão nas defensas", "#F4D3A1"),
              ("Cabecos", "Cabeços", "#C9C3B8")]


def eventos_governantes():
    ca = D.carregar_ca_eventos()
    c75 = R75.groupby(["ID", "Secao"]).size().unstack(fill_value=0)
    c75 = c75.reindex(columns=[s for s, _, _ in SECOES_EVT], fill_value=0)
    cch = D.carregar_eventos_cheio().rename(columns={"Reacao Defensas": "Reacao", "Compressao Defensas": "Compressao"})
    fig, axs = plt.subplots(1, 2, figsize=FIG, layout="constrained", sharey=True)
    for ax, cont, titulo in [(axs[0], c75, "75% (12,75 m)"), (axs[1], cch, "Cheio (14 m)")]:
        cont = cont.loc[cont.sum(axis=1).sort_values(ascending=False).index]
        x = np.arange(len(cont))
        base = np.zeros(len(cont))
        for s, nome, cor in SECOES_EVT:
            ax.bar(x, cont[s], bottom=base, color=cor, edgecolor="white", lw=0.8, width=0.68, label=nome)
            base += cont[s].to_numpy()
        for xi, tot in zip(x, base):
            ax.text(xi, tot + 0.25, f"{int(tot)}", ha="center", fontsize=8, weight="bold", color=ESC)
        rot = [f"ID {i}\n{br(ca.loc[i, 'V. corrente [nos]'], 2)} nó\n{br(ca.loc[i, 'V. vento [nos]'], 0)} nós"
               for i in cont.index]
        ax.set_xticks(x, rot, fontsize=6.6)
        ax.grid(axis="x", visible=False)
        ax.set_title(f"{titulo} · picos por evento do CA", loc="left")
    axs[0].set_ylabel("Elementos com pico no evento")
    axs[0].set_ylim(0, 15)
    axs[0].legend(loc="upper right", fontsize=6.8)
    fig.text(0.01, -0.02, "Eixo: ID do evento, corrente e vento no instante do pico", fontsize=7, color="#6F7684")
    save(fig, "14_eventos_governantes.png")


# ─────────────────────────────────────────────────────────────
# Slide 15 · Tração e comprimento de cada cabo (75%)
# ─────────────────────────────────────────────────────────────
def cabos_tracao_comprimento():
    cab = R75[R75["Secao"] == "Cabos"].copy()
    cab["comp"] = [D.geometria_linha(D.MAPA_CABOS[n], n.split("-")[1], D.CONDICOES["75%"], m)["comp"]
                   for n, m in zip(cab["Elemento"], cab["Mare [m]"])]
    quatro = set(cab.nsmallest(4, "comp")["Elemento"])
    fig, ax = plt.subplots(figsize=FIG, layout="constrained")
    x = np.arange(len(cab))
    cores = [ROXO if n in quatro else "#C9C3B8" for n in cab["Elemento"]]
    barras = ax.bar(x, cab["Valor"], color=cores, edgecolor="white", width=0.66, zorder=2)
    ax.bar_label(barras, labels=[f"{v:.0f}%" for v in cab["Valor"]], padding=2, fontsize=7.5, color=ESC)
    ax.axhline(D.LIMITE_MBL, color=LIM, ls="--", lw=1.3, label="Limite (55% MBL)")
    ax.set_ylim(0, 70)
    ax.set_ylabel("Pico de tração (% MBL)")
    ax.set_xticks(x, cab["Elemento"])
    ax.grid(axis="x", visible=False)
    ax2 = ax.twinx()
    ax2.plot(x, cab["comp"], "o-", color=ESC, lw=1.1, ms=4, label="Comprimento 3D [m]")
    ax2.set_ylim(0, 50)
    ax2.set_ylabel("Comprimento 3D no pico [m]")
    ax2.grid(False)
    ax2.spines["right"].set_visible(True)
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    from matplotlib.patches import Patch
    ax.legend(h1 + h2 + [Patch(color=ROXO), Patch(color="#C9C3B8")],
              l1 + l2 + ["4 cabos mais curtos", "demais cabos"], loc="upper left", ncol=2)
    ax.set_title("75% (12,75 m) · pico de tração e comprimento de cada cabo", loc="left")
    save(fig, "15_tracao_comprimento_75.png")
    return cab


# ─────────────────────────────────────────────────────────────
# Slide 17 · Envelope crítico (75%)
# ─────────────────────────────────────────────────────────────
def envelope_75():
    fig, axs = plt.subplots(1, 2, figsize=FIG, layout="constrained")
    fator = 150
    tam = lambda m: (np.asarray(m) - D.MARE_MIN + 0.3) * fator
    for ax, secao, col, vmax, titulo, cbl in [
            (axs[0], "Cabos", "Valor", D.LIMITE_MBL, "Cabos", "Pico (% MBL)"),
            (axs[1], "Cabecos", "% do Limite", D.LIMITE_SWL, "Cabeços", "Pico (% SWL)")]:
        d = R75[R75["Secao"] == secao].sort_values(col)
        norm = Normalize(0, vmax)
        ax.scatter(d["V. corrente [nos]"], d["V. vento [nos]"], s=tam(d["Mare [m]"]), c=d[col],
                   cmap="coolwarm", norm=norm, edgecolors=ESC, linewidths=0.6, zorder=3)
        for (c, v, m), g in d.groupby(["V. corrente [nos]", "V. vento [nos]", "Mare [m]"]):
            labs = list(g.sort_index()["Elemento"])
            txt = ", ".join(labs)
            if len(labs) > 4:
                meio = (len(labs) + 1) // 2
                txt = ", ".join(labs[:meio]) + ",\n" + ", ".join(labs[meio:])
            ax.annotate(txt, (c, v), xytext=(np.sqrt(tam(m)) / 2 + 3, 0), textcoords="offset points",
                        fontsize=6.5, va="center", color=ESC)
        ax.set_xlim(0.8, 2.15)
        virgula(ax.xaxis, 1)
        ax.set_ylim(0, 57)
        ax.set_xlabel("Corrente [nós]")
        ax.set_ylabel("Vento [nós]")
        ax.set_title(f"{titulo} · condição de pico", loc="left")
        cb = fig.colorbar(ScalarMappable(norm=norm, cmap="coolwarm"), ax=ax, fraction=0.06, pad=0.02)
        cb.set_label(cbl, fontsize=7.5)
        cb.outline.set_visible(False)
        cb.ax.tick_params(labelsize=7)
    marcas = [-0.2, 0.7, 1.6]
    leg = [axs[0].scatter([], [], s=tam(m), c="#C9C3B8", edgecolors=ESC, linewidths=0.5) for m in marcas]
    axs[0].legend(leg, [f"{br(m, 2)} m" for m in marcas], title="Maré", loc="lower left",
                  labelspacing=1.1, fontsize=6.5, title_fontsize=7, borderpad=0.6)
    save(fig, "17_envelope_75.png")


# ─────────────────────────────────────────────────────────────
# Slide novo · Efeito do carregamento
# ─────────────────────────────────────────────────────────────
def efeito_carregamento():
    util_cabos = serie("% MBL", "Maximo") / D.LIMITE_MBL * 100
    util_cabecos = serie("% SWL", "Maximo")
    ang = {m: [D.tabela_geometria(c, m)["ang"].max() for c in CALADOS] for m in (D.MARE_MIN, D.MARE_MAX)}
    fig, axs = plt.subplots(2, 1, figsize=FIG, layout="constrained", sharex=True)
    ax = axs[0]
    ax.plot(CALADOS, util_cabos, "o-", color=AZUL, lw=1.6, ms=5, label="Cabos (pico % MBL ÷ 55%)")
    ax.plot(CALADOS, util_cabecos, "^-", color="#E39A55", lw=1.6, ms=5.5, label="Cabeços (pico % SWL)")
    for serie_, cor, dy in [(util_cabos, AZUL, -12), (util_cabecos, "#C77B36", 6)]:
        for x, v in zip(CALADOS, serie_):
            ax.annotate(br(v, 1) + "%", (x, v), xytext=(0, dy), textcoords="offset points",
                        ha="center", fontsize=7.2, color=cor, weight="bold")
    ax.axhline(100, color=LIM, ls="--", lw=1.3, label="Limite (100% do admissível)")
    ax.set_ylim(60, 108)
    ax.set_ylabel("% do limite")
    ax.set_title("Esforço máximo em relação ao limite", loc="left")
    ax.legend(loc="lower right", ncol=3, fontsize=7)
    ax = axs[1]
    ax.plot(CALADOS, ang[D.MARE_MAX], "o-", color=ROXO, lw=1.6, ms=5, label=f"maré {br(D.MARE_MAX, 2)} m")
    ax.plot(CALADOS, ang[D.MARE_MIN], "o--", color="#A99BD8", lw=1.4, ms=4.5, label=f"maré {br(D.MARE_MIN, 2)} m")
    for x, v in zip(CALADOS, ang[D.MARE_MAX]):
        ax.annotate(br(v, 1) + "°", (x, v), xytext=(0, 6), textcoords="offset points", ha="center",
                    fontsize=7.2, color=ROXO, weight="bold")
    ax.axhline(25, color=LIM, ls="--", lw=1.3, label="Limite (25°)")
    ax.set_ylim(0, 28)
    ax.set_ylabel("Âng. vertical máx. [°]")
    ax.set_title("Maior ângulo vertical entre os 12 cabos", loc="left")
    ax.legend(loc="upper right", ncol=3, fontsize=7)
    ax.set_xticks(CALADOS, ROT_CALADO)
    save(fig, "22_efeito_carregamento.png")


# ─────────────────────────────────────────────────────────────
# Slide 6 · O navio: medidas principais e movimentos (esquema)
# ─────────────────────────────────────────────────────────────
TEAL, CINZA, AGUA = "#2F8F86", "#9A938A", "#DCE8F2"
LOA, BOCA, PONTAL, T_LASTRO, T_CHEIO = 225.0, 32.26, 18.5, 9.0, 14.0


def _cota(ax, p0, p1, txt, offset=(0, 0), ha="center", va="center", cor=ESC, fs=7.5):
    ax.annotate("", p1, p0, arrowprops=dict(arrowstyle="<->", color=cor, lw=0.9, shrinkA=0, shrinkB=0))
    xm, ym = (p0[0] + p1[0]) / 2 + offset[0], (p0[1] + p1[1]) / 2 + offset[1]
    ax.text(xm, ym, txt, ha=ha, va=va, fontsize=fs, color=cor)


def navio_medidas_movimentos():
    from matplotlib.patches import Polygon, Rectangle, FancyArrowPatch
    fig, axs = plt.subplots(1, 3, figsize=(11.7, 1.95), layout="constrained",
                            gridspec_kw={"width_ratios": [1.35, 0.85, 1.0]})
    for ax in axs:
        ax.axis("off")

    # ── A · Perfil: LOA, pontal, calado e borda livre (fora de escala na vertical)
    ax = axs[0]
    ax.add_patch(Rectangle((-12, -2.5), 252, T_LASTRO + 2.5, fc=AGUA, ec="none", zorder=0))
    casco = [(0, PONTAL), (0, 4), (7, 0), (212, 0), (222, 6), (229, PONTAL)]
    ax.add_patch(Polygon(casco, closed=True, fc="#F4EFE7", ec="#A49C90", lw=1.0, zorder=2))
    ax.add_patch(Polygon([(0, T_LASTRO), (0, 4), (7, 0), (212, 0), (220.8, T_LASTRO * 0.6 + 0.6),
                          (222.9, T_LASTRO)], closed=True, fc="#C9D7E6", ec="none", alpha=0.55, zorder=3))
    ax.add_patch(Rectangle((6, PONTAL), 24, 13, fc="#E9E3D8", ec="#A49C90", lw=0.8, zorder=2))
    for k in range(7):                                        # escotilhas dos porões
        ax.add_patch(Rectangle((48 + 22 * k, PONTAL), 15, 2.2, fc="#E9E3D8", ec="#A49C90", lw=0.6, zorder=2))
    ax.plot([-12, 240], [T_LASTRO, T_LASTRO], color="#6E9CC4", lw=1.0, zorder=4)
    ax.text(118, 4.3, "parte submersa: área sob a corrente", ha="center", fontsize=7, style="italic",
            color="#3F5F96", zorder=5)
    ax.text(118, 13.4, "parte emersa: área ao vento", ha="center", fontsize=7, style="italic",
            color="#7A7266", zorder=5)
    _cota(ax, (0, 37), (229, 37), "comprimento total  LOA = 225 m", offset=(0, 3.6))
    _cota(ax, (237, 0), (237, PONTAL), "pontal\nD = 18,5 m", offset=(5, 0), ha="left")
    _cota(ax, (-6, 0), (-6, T_LASTRO), "calado T = 9 a 14 m", offset=(-4, -0.5), ha="right", fs=7.2)
    _cota(ax, (-6, T_LASTRO), (-6, PONTAL), "borda livre = 9,5 a 4,5 m", offset=(-4, 0.8), ha="right", fs=7.2)
    ax.text(-100, 47, "Perfil · medidas principais", fontsize=8.5, weight="bold", color="#1C2430")
    ax.text(-100, -6.5, "lastro (T = 9 m) · escala vertical ampliada", fontsize=6.5,
            style="italic", color="#8A8275")
    ax.set_xlim(-100, 272)
    ax.set_ylim(-8, 50)

    # ── B · Seção junto ao cais: boca, buzina, cabeço e cabo (em escala, maré 0)
    ax = axs[1]
    fundo, z_cais, y_def = -16.5, 4.6, 1.4
    z_q, z_conves = -T_LASTRO, PONTAL - T_LASTRO
    ax.add_patch(Rectangle((-10, fundo), 54, -fundo, fc=AGUA, ec="none", zorder=0))
    ax.add_patch(Rectangle((-10, fundo), 10, z_cais - fundo, fc="#E9E3D8", ec="#B9B2A6", lw=0.8, zorder=1))
    y0, y1 = y_def, y_def + BOCA
    secao = [(y0, z_conves), (y0, z_q + 2.5), (y0 + 2.5, z_q), (y1 - 2.5, z_q), (y1, z_q + 2.5), (y1, z_conves)]
    ax.add_patch(Polygon(secao, closed=True, fc="#F4EFE7", ec="#A49C90", lw=1.0, zorder=2))
    ax.add_patch(Rectangle((0, 0.8), y_def, 2.6, fc="#5E574D", ec="none", zorder=3))      # defensa
    ax.plot([-10, 44], [0, 0], color="#6E9CC4", lw=1.0, zorder=4)
    ax.plot(-3.5, z_cais + 0.8, "s", ms=4.5, color=ESC, zorder=5)                          # cabeço
    ax.plot(y0 + 0.4, z_conves, "o", ms=4, color=ESC, zorder=5)                            # buzina
    ax.plot([-3.5, y0 + 0.4], [z_cais + 0.8, z_conves], color=TEAL, lw=1.8, zorder=4)
    ax.text(-9.5, z_cais + 1.6, "cais", fontsize=7, color="#8A8275", weight="bold")
    ax.text(y0 + 1.2, z_conves + 1.6, "buzina", fontsize=7, color=ESC)
    ax.text(y0 + 1.0, 2.1, "← defensa", fontsize=6.5, color="#5E574D", ha="left", va="center")
    _cota(ax, (y0, z_conves + 4.5), (y1, z_conves + 4.5), "boca B = 32,26 m", offset=(0, 1.9))
    _cota(ax, (y1 + 2.2, z_q), (y1 + 2.2, 0), "T", offset=(1.2, 0), ha="left")
    ax.text(-10, 21.5, "Seção junto ao cais", fontsize=8.5, weight="bold", color="#1C2430")
    ax.text(-10, fundo - 2.6, "em escala · lastro, maré 0", fontsize=6.5, style="italic", color="#8A8275")
    ax.set_xlim(-10.5, 41)
    ax.set_ylim(fundo - 3.5, 24)
    ax.set_aspect("equal")

    # ── C · Os 6 movimentos (vista de planta)
    ax = axs[2]
    ax.add_patch(Rectangle((-15, -24), 167, 8, fc="#E9E3D8", ec="none", zorder=0))
    ax.text(-13, -21.6, "cais", fontsize=7, color="#8A8275", weight="bold")
    planta = [(0, -10), (88, -10), (100, -5), (106, 0), (100, 5), (88, 10), (0, 10), (-3, 6), (-3, -6)]
    ax.add_patch(Polygon(planta, closed=True, fc="#F4EFE7", ec="#A49C90", lw=1.0, zorder=1))
    ax.text(14, 0, "popa", fontsize=6.5, color="#A49C90", weight="bold", va="center")
    ax.text(84, 0, "proa", fontsize=6.5, color="#A49C90", weight="bold", va="center", ha="right")
    seta = dict(arrowstyle="<|-|>", color=TEAL, lw=1.6, mutation_scale=9)
    ax.add_patch(FancyArrowPatch((30, 0), (66, 0), **seta, zorder=3))
    ax.text(48, -5.2, "avanço (surge) · x", fontsize=7.2, color=TEAL, ha="center", weight="bold")
    ax.add_patch(FancyArrowPatch((118, -14), (118, 14), **seta, zorder=3))
    ax.text(121, 0, "deriva\n(sway) · y", fontsize=7.2, color=TEAL, va="center", weight="bold")
    ax.add_patch(FancyArrowPatch((36, 17), (62, 17), connectionstyle="arc3,rad=-0.35", **seta, zorder=3))
    ax.text(49, 24.5, "guinada (yaw) · em torno de z", fontsize=7.2, color=TEAL, ha="center", weight="bold")
    ax.text(-13, -31, "fora do plano: afundamento (heave), jogo (roll) e arfagem (pitch)",
            fontsize=6.8, color=CINZA, style="italic")
    ax.text(-13, 35, "Os 6 movimentos · vista de planta", fontsize=8.5, weight="bold", color="#1C2430")
    ax.set_xlim(-15, 152)
    ax.set_ylim(-34, 38)
    ax.set_aspect("equal")
    save(fig, "06_navio_medidas_movimentos.png")


if __name__ == "__main__":
    navio_medidas_movimentos()
    estatistica_por_calado()
    influencia_mare_75()
    eventos_governantes()
    cabos_tracao_comprimento()
    envelope_75()
    efeito_carregamento()
    print("Gráficos salvos em", OUT)
