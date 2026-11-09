"""
Recria os gráficos do notebook TF4.0 - Colab (calado 9 m, 394 condições)
com uma paleta suave para a apresentação.

Valores EXATOS (copiados das saídas do notebook):
  - contagens dos histogramas, matriz de correlação, matriz de operabilidade,
    picos de 75,3% MBL (cabo 9) e 110,8% SWL (cabeço S), coordenadas do plano.
Valores APROXIMADOS (lidos dos gráficos originais):
  - demais picos e níveis de maré do gráfico de influência da maré,
  - posições do envelope crítico, quartis do boxplot.
Se os PNGs originais forem colocados em ../imagens/, o script da apresentação
usa os originais no lugar destes.
"""
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.transforms
from matplotlib.cm import ScalarMappable
from matplotlib.colors import LinearSegmentedColormap, Normalize, TwoSlopeNorm
from matplotlib.lines import Line2D
from matplotlib.patches import Polygon

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "imagens_recriadas")
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    # STIX é a fonte com desenho de Times que acompanha o matplotlib (Times New Roman nos slides)
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
FIG = (6.85, 4.6)    # mesma proporção do quadro de imagem nos slides
DPI = 220


def br(v, d=1):
    return f"{v:.{d}f}".replace(".", ",")


def virgula(eixo, d=None):
    """Rótulos numéricos do eixo com vírgula decimal (padrão brasileiro)."""
    from matplotlib.ticker import FuncFormatter

    def f(v, _):
        if d is None:
            txt = f"{v:g}"
        else:
            txt = f"{v:.{d}f}"
        return txt.replace("-", "−").replace(".", ",")
    eixo.set_major_formatter(FuncFormatter(f))


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=DPI, bbox_inches="tight", pad_inches=0.06)
    plt.close(fig)


# ─────────────────────────────────────────────────────────────
# 00 · Esquema do plano de amarração (coordenadas do notebook)
# ─────────────────────────────────────────────────────────────
def plano_amarracao():
    cabecos_x = {c: -135 + 15 * i for i, c in enumerate("ABCDEFGHIJKLMNOPQRS")}
    casco = [(1, -113.2, 0), (2, -113.2, 5.9), (3, -113.2, 7.5), (4, -99.6, 15.3),
             (5, -86.1, 16.1), (6, -56.4, 16.0), (7, -55.3, 16.0), (8, -30.7, 16.1),
             (9, 20.9, 16.2), (10, 72.0, 16.1), (11, 73.1, 16.1), (12, 87.0, 15.1),
             (13, 99.4, 12.1), (14, 104.8, 9.2), (15, 105.5, 8.8), (16, 109.4, 5.3),
             (17, 111.7, 0), (18, 109.4, -5.3), (19, 105.5, -8.8), (20, 104.8, -9.2),
             (21, 99.4, -12.1), (22, 87.0, -15.1), (23, 73.1, -16.1), (24, 72.0, -16.1),
             (25, 20.9, -16.0), (26, -30.7, -16.1), (27, -55.3, -16.0), (28, -56.4, -16.0),
             (29, -86.1, -16.1), (30, -99.6, -15.3), (31, -113.2, -7.5), (32, -113.2, -5.9)]
    # Esquema: posições longitudinais (X) do notebook, com o X espelhado como no notebook.
    # Na transversal o desenho é fora de escala: o vão entre costado e cais foi ampliado
    # para que as linhas fiquem visíveis, e as buzinas ficam no costado de atracação.
    GAP = 10.0                          # vão desenhado entre a face do cais e o costado
    yc = GAP + 16.2                     # linha de centro do navio no desenho
    contorno = [(-x, yc - y) for _, x, y in casco]
    pos = {n: (-x, yc - abs(y)) for n, x, y in casco}
    plano = [(2, "S"), (3, "S"), (14, "A"), (15, "A"), (17, "B"), (18, "B"),
             (23, "G"), (24, "G"), (28, "N"), (29, "N"), (31, "S"), (32, "S")]
    y_cab = -3.5                        # cabeços no cais, atrás da face

    fig, ax = plt.subplots(figsize=(11.7, 2.6))
    ax.axhspan(-14, 0, color="#E9E3D8", lw=0)
    ax.text(-148, -12, "CAIS", fontsize=8, color="#8A8275", weight="bold")
    ax.text(-148, 46, "Esquema fora de escala na direção transversal", ha="left", va="center",
            fontsize=7, style="italic", color="#8A8275")
    ax.add_patch(Polygon(contorno, closed=True,
                         fc="#F4EFE7", ec="#A49C90", lw=1.0, zorder=1))
    ax.text(0, yc, "Navio Panamax · LOA 225 m · atracado por boreste",
            ha="center", va="center", fontsize=8, color="#7A7266")

    usados = {c for _, c in plano}
    for c, x in cabecos_x.items():
        if c in usados:
            ax.plot(x, y_cab, "o", ms=6, color=ESC, zorder=4)
            ax.text(x, -10.5, c, ha="center", fontsize=8.5, weight="bold", color=ESC)
        else:
            ax.plot(x, y_cab, "o", ms=2.5, color="#B9B2A6", zorder=2)
    ax.text(-100, yc, "PROA", ha="center", va="center", fontsize=7.5, weight="bold", color="#A49C90")
    ax.text(100, yc, "POPA", ha="center", va="center", fontsize=7.5, weight="bold", color="#A49C90")

    for i, (fl, cab) in enumerate(plano, start=1):
        (x0, y0), x1 = pos[fl], cabecos_x[cab]
        critico = i == 9
        ax.plot([x0, x1], [y0, y_cab], color=LIM if critico else "#7F9CC8",
                lw=2.4 if critico else 1.2, zorder=5 if critico else 3)

    # rótulos (pares que vão ao mesmo cabeço recebem um rótulo só)
    rot = {"1–2, 11–12": (2, 32, "S"), "3–4": (14, 15, "A"), "5–6": (17, 18, "B"),
           "7–8": (23, 24, "G"), "9": (28, 28, "N"), "10": (29, 29, "N")}
    for txt, (fa, fb, cab) in rot.items():
        xm = (pos[fa][0] + pos[fb][0]) / 2 * 0.5 + cabecos_x[cab] * 0.5
        ym = (pos[fa][1] + pos[fb][1]) / 2 * 0.5 + y_cab * 0.5
        crit = txt == "9"
        ax.text(xm, ym, txt, ha="center", va="center", fontsize=7.5, weight="bold",
                color=LIM if crit else "#3F5F96", zorder=6,
                bbox=dict(boxstyle="round,pad=0.18", fc="white", ec="none", alpha=0.9))

    ax.legend(handles=[Line2D([], [], color=LIM, lw=2.4, label="Cabo 9 (buzina 28 → N): crítico"),
                       Line2D([], [], color="#7F9CC8", lw=1.2, label="Demais cabos")],
              loc="upper right", bbox_to_anchor=(1.0, 1.12), ncol=2, fontsize=7.5)
    ax.set_xlim(-150, 150)
    ax.set_ylim(-14, 48)
    ax.set_aspect("equal")
    ax.axis("off")
    save(fig, "00_plano_amarracao.png")


# ─────────────────────────────────────────────────────────────
# 01 · Histogramas (% MBL e % SWL) — contagens exatas
# ─────────────────────────────────────────────────────────────
def histogramas():
    fig, axs = plt.subplots(2, 1, figsize=FIG)
    specs = [
        (axs[0], 9.80, 75.30, [92, 51, 40, 40, 42, 36, 37, 32, 19, 5], 55,
         "#9DB5DE", "Cabos · distribuição de % MBL", "% MBL", "Limite operacional (55%)"),
        (axs[1], 24.20, 110.80, [37, 103, 59, 52, 38, 40, 34, 10, 13, 1, 3, 2, 1, 1], 100,
         "#F5C996", "Cabeços · distribuição de % SWL", "% SWL", "Limite estrutural (100%)"),
    ]
    for ax, lo, hi, counts, lim, cor, titulo, xl, leg in specs:
        edges = np.linspace(lo, hi, len(counts) + 1)
        bars = ax.bar(edges[:-1], counts, width=np.diff(edges), align="edge",
                      color=cor, edgecolor="white", linewidth=0.9)
        ax.bar_label(bars, labels=[str(c) for c in counts], padding=1.5, fontsize=7, color="#3A4250")
        span = hi - lo
        ax.axvspan(lim, hi + span * 0.03, color=LIM, alpha=0.07, lw=0)
        ax.axvline(lim, color=LIM, ls="--", lw=1.3, label=leg)
        ax.set_title(titulo, loc="left")
        ax.set_xlabel(xl)
        ax.set_ylabel("Condições")
        ax.set_ylim(0, max(counts) * 1.2)
        ax.set_xlim(lo - span * 0.02, hi + span * 0.03)
        ax.grid(axis="x", visible=False)
        ax.legend(loc="upper right")
    fig.tight_layout(h_pad=1.0)
    save(fig, "01_histogramas.png")


# ─────────────────────────────────────────────────────────────
# 02 · Influência da maré (pico por elemento × maré no pico)
# ─────────────────────────────────────────────────────────────
CABOS = ["1-S", "2-S", "3-A", "4-A", "5-B", "6-B", "7-G", "8-G", "9-N", "10-N", "11-S", "12-S"]
CABO_PICO = [34, 34, 25, 25, 30, 38, 16, 17, 75.3, 39, 21, 23]
CABO_MARE = [0.62, 0.62, 0.33, 0.33, 0.33, 1.00, 0.88, 0.88, 1.58, 0.15, 0.33, 0.33]
CABECOS = ["A", "B", "G", "N", "S"]
CAB_PICO = [49.5, 67, 33, 79, 110.8]
CAB_MARE = [0.33, 0.33, 0.88, 1.58, 0.33]


def influencia_mare():
    cmap, norm = plt.get_cmap("coolwarm"), Normalize(0, 1.6)
    fig, axs = plt.subplots(2, 1, figsize=FIG, layout="constrained")

    ax = axs[0]
    x = np.arange(len(CABOS))
    ax.scatter(x, CABO_PICO, c=CABO_MARE, cmap=cmap, norm=norm, s=70,
               edgecolors=ESC, linewidths=0.5, zorder=3)
    ax.axhline(55, color=LIM, ls="--", lw=1.3, label="Limite operacional (55%)")
    ax.annotate("75,3%", (8, 75.3), xytext=(9, -3), textcoords="offset points",
                fontsize=7.5, weight="bold", color=ESC)
    ax.set_xticks(x, CABOS)
    ax.set_ylim(0, 92)
    ax.set_ylabel("Pico (% MBL)")
    ax.set_title("Cabos · pico de tração e maré no instante do pico", loc="left")
    ax.legend(loc="upper left")

    ax = axs[1]
    x = np.arange(len(CABECOS))
    ax.scatter(x, CAB_PICO, c=CAB_MARE, cmap=cmap, norm=norm, s=85, marker="^",
               edgecolors=ESC, linewidths=0.5, zorder=3)
    ax.axhline(100, color=LIM, ls="--", lw=1.3, label="Limite estrutural (100%)")
    ax.annotate("110,8%", (4, 110.8), xytext=(-40, -3), textcoords="offset points",
                fontsize=7.5, weight="bold", color=ESC)
    ax.set_xticks(x, CABECOS)
    ax.set_xlim(-0.5, 4.5)
    ax.set_ylim(0, 125)
    ax.set_ylabel("Pico (% SWL)")
    ax.set_title("Cabeços · pico de esforço e maré no instante do pico", loc="left")
    ax.legend(loc="upper left")

    cb = fig.colorbar(ScalarMappable(norm=norm, cmap=cmap), ax=axs, fraction=0.035, pad=0.02)
    virgula(cb.ax.yaxis, 1)
    cb.set_label("Nível da maré [m]")
    cb.outline.set_visible(False)
    save(fig, "02_influencia_mare.png")


# ─────────────────────────────────────────────────────────────
# 04 · Matriz de correlação (exata) + boxplot (aproximado)
# ─────────────────────────────────────────────────────────────
def correlacao_boxplot():
    vars_ = ["Vento", "Corrente", "Hs", "Tp", "% MBL", "% SWL"]
    C = np.array([[1.00, 0.03, 0.05, 0.01, 0.02, 0.25],
                  [0.03, 1.00, 0.01, 0.05, 0.05, 0.20],
                  [0.05, 0.01, 1.00, -0.51, 0.08, 0.06],
                  [0.01, 0.05, -0.51, 1.00, -0.07, -0.02],
                  [0.02, 0.05, 0.08, -0.07, 1.00, 0.70],
                  [0.25, 0.20, 0.06, -0.02, 0.70, 1.00]])
    fig, (a1, a2) = plt.subplots(1, 2, figsize=FIG, layout="constrained",
                                 gridspec_kw={"width_ratios": [1.0, 1.05]})
    im = a1.imshow(C, cmap="coolwarm", vmin=-1, vmax=1)
    for i in range(6):
        for j in range(6):
            a1.text(j, i, br(C[i, j], 2), ha="center", va="center", fontsize=6.5,
                    color="white" if abs(C[i, j]) > 0.6 else "#2F3A48")
    a1.set_xticks(range(6), vars_, rotation=45, ha="right")
    a1.set_yticks(range(6), vars_)
    a1.grid(False)
    for s in a1.spines.values():
        s.set_visible(False)
    a1.set_title("Correlação: ambiente × esforços", loc="left")
    cb = fig.colorbar(im, ax=a1, fraction=0.046, pad=0.03, shrink=0.75)
    virgula(cb.ax.yaxis, 2)
    cb.outline.set_visible(False)
    cb.ax.tick_params(labelsize=7)

    stats = [
        dict(label="9", whislo=11.5, q1=27.2, med=39.3, q3=52.0, whishi=75.3, fliers=[]),
        dict(label="2", whislo=18.5, q1=21.5, med=23.5, q3=26.5, whishi=33.0, fliers=[]),
        dict(label="5", whislo=14.8, q1=14.9, med=15.0, q3=15.0, whishi=15.1, fliers=[]),
        dict(label="6", whislo=11.2, q1=13.0, med=13.8, q3=15.5, whishi=19.0, fliers=[21.3]),
        dict(label="12", whislo=12.4, q1=12.5, med=12.6, q3=12.6, whishi=12.7, fliers=[]),
        dict(label="10", whislo=11.0, q1=11.8, med=12.8, q3=16.8, whishi=23.3,
             fliers=[24.5, 25.5, 26.0, 31.8, 39.3]),
        dict(label="8", whislo=10.5, q1=11.2, med=11.6, q3=11.9, whishi=12.1, fliers=[]),
        dict(label="3", whislo=10.0, q1=10.3, med=10.5, q3=10.7, whishi=10.9, fliers=[]),
        dict(label="7", whislo=9.8, q1=9.85, med=9.9, q3=9.95, whishi=10.0, fliers=[]),
    ]
    cores = ["#8CC7B5", "#F4A98A", "#A9B9DD", "#E8A9CF", "#B7DB8F",
             "#F6D77A", "#D9C3A0", "#BDBDBD", "#9FCFE0"]
    bp = a2.bxp(stats, patch_artist=True, widths=0.6,
                medianprops=dict(color=ESC, lw=1.1),
                flierprops=dict(marker="o", ms=3.5, mfc="white", mec="#6B7280"),
                whiskerprops=dict(color="#6B7280"), capprops=dict(color="#6B7280"))
    for patch, cor in zip(bp["boxes"], cores):
        patch.set_facecolor(cor)
        patch.set_edgecolor("#6B7280")
    a2.axhline(55, color=LIM, ls="--", lw=1.3, label="Limite operacional (55%)")
    a2.set_xlabel("Cabo mais carregado na condição")
    a2.set_ylabel("% MBL")
    a2.set_ylim(5, 80)
    a2.set_title("% MBL por cabo crítico", loc="left")
    a2.legend(loc="upper right")
    save(fig, "04_correlacao_boxplot.png")


# ─────────────────────────────────────────────────────────────
# 05 · Envelope crítico (corrente × vento, cor = esforço, tamanho = maré)
# ─────────────────────────────────────────────────────────────
ENV_CABOS = [(0.88, 49.1, 0.62, 34, "1-S"), (0.88, 49.1, 0.62, 34, "2-S"),
             (1.35, 49.0, 0.88, 16, "7-G"), (1.35, 49.0, 0.88, 17, "8-G"),
             (1.27, 46.1, 0.15, 39, "10-N"), (1.48, 43.7, 1.00, 38, "6-B"),
             (1.58, 36.7, 0.33, 25, "3-A"), (1.58, 36.7, 0.33, 25, "4-A"),
             (1.58, 36.7, 0.33, 30, "5-B"), (1.58, 36.7, 0.33, 21, "11-S"),
             (1.58, 36.7, 0.33, 23, "12-S"), (1.57, 8.2, 1.58, 75.3, "9-N")]
ENV_CABECOS = [(1.35, 49.0, 0.88, 33, "G"), (1.58, 36.7, 0.33, 49.5, "A"),
               (1.58, 36.7, 0.33, 67, "B"), (1.58, 36.7, 0.33, 110.8, "S"),
               (1.57, 8.2, 1.58, 79, "N")]


def envelope():
    fig, axs = plt.subplots(1, 2, figsize=FIG, layout="constrained")
    fator, mare_min = 190, 0.15
    tam = lambda m: (m - mare_min + 0.3) * fator
    for ax, pts, vmin, vmax, titulo, cbl in [
        (axs[0], ENV_CABOS, 15, 76, "Cabos", "Pico (% MBL)"),
        (axs[1], ENV_CABECOS, 30, 112, "Cabeços", "Pico (% SWL)")]:
        norm = Normalize(vmin, vmax)
        for c, v, m, p, _ in sorted(pts, key=lambda t: t[3]):
            ax.scatter(c, v, s=tam(m), c=[p], cmap="coolwarm", norm=norm,
                       edgecolors=ESC, linewidths=0.6, zorder=3)
        grupos = {}
        for c, v, m, p, lab in pts:
            grupos.setdefault((c, v), [m, []])[1].append(lab)
        for (c, v), (m, labs) in grupos.items():
            txt = ", ".join(labs)
            if len(txt) > 14:
                meio = (len(labs) + 1) // 2
                txt = ", ".join(labs[:meio]) + ",\n" + ", ".join(labs[meio:])
            raio = np.sqrt(tam(m)) / 2          # raio da bolha em pontos
            ax.annotate(txt, (c, v), xytext=(raio + 3, 0), textcoords="offset points",
                        fontsize=6.5, va="center", color=ESC)
        ax.set_xlim(0.75, 1.92)
        virgula(ax.xaxis, 1)
        ax.set_ylim(0, 57)
        ax.set_xlabel("Corrente [nós]")
        ax.set_ylabel("Vento [nós]")
        ax.set_title(f"{titulo} · condição de pico", loc="left")
        cb = fig.colorbar(ScalarMappable(norm=norm, cmap="coolwarm"), ax=ax,
                          fraction=0.06, pad=0.02)
        cb.set_label(cbl, fontsize=7.5)
        cb.outline.set_visible(False)
        cb.ax.tick_params(labelsize=7)
    leg = [axs[0].scatter([], [], s=tam(m), c="#C9C3B8", edgecolors=ESC, linewidths=0.5)
           for m in (0.0, 0.75, 1.5)]
    axs[0].legend(leg, ["0,00 m", "0,75 m", "1,50 m"], title="Maré", loc="lower left",
                  labelspacing=1.1, fontsize=6.5, title_fontsize=7, borderpad=0.6)
    save(fig, "05_envelope_critico.png")


# ─────────────────────────────────────────────────────────────
# 06 · Matriz de operabilidade (valores exatos)
# ─────────────────────────────────────────────────────────────
def operabilidade():
    val = np.array([[61.2, 64.4, 71.9, 75.3],     # vento 0–10
                    [64.0, 66.1, 64.1, 38.8],     # vento 10–20
                    [64.9, 64.6, 65.9, 69.1],     # vento 20–30
                    [69.1, 60.7, 62.5, 67.4],     # vento 30–40
                    [59.4, 61.2, 73.0, 67.1]])    # vento 40–50
    cmap = LinearSegmentedColormap.from_list("op", ["#86C995", "#F6E58D", "#EC7B67"])
    norm = TwoSlopeNorm(vmin=38.8, vcenter=55, vmax=75.3)
    fig, ax = plt.subplots(figsize=FIG, layout="constrained")
    im = ax.imshow(val, cmap=cmap, norm=norm, origin="lower", aspect="auto")
    ax.set_xticks(np.arange(-0.5, 4), minor=True)
    ax.set_yticks(np.arange(-0.5, 5), minor=True)
    ax.grid(False)
    ax.grid(which="minor", color="white", lw=2.5)
    ax.tick_params(which="minor", length=0)
    for s in ax.spines.values():
        s.set_visible(False)
    for i in range(5):
        for j in range(4):
            ax.text(j, i, br(val[i, j]), ha="center", va="center", fontsize=10,
                    weight="bold" if val[i, j] > 55 else "normal", color="#2F3A48")
    ax.set_xticks(range(4), ["0 – 0,5", "0,5 – 1,0", "1,0 – 1,5", "1,5 – 2,0"])
    ax.set_yticks(range(5), ["0 – 10", "10 – 20", "20 – 30", "30 – 40", "40 – 50"])
    ax.set_xlabel("Faixa de corrente [nós]")
    ax.set_ylabel("Faixa de vento [nós]")
    ax.set_title("Pior tração (% MBL) em cada faixa de vento × corrente", loc="left")
    cb = fig.colorbar(im, ax=ax, fraction=0.05, pad=0.02)
    cb.set_label("% MBL (limite 55%)")
    cb.ax.axhline(55, color=LIM, ls="--", lw=1.3)
    cb.outline.set_visible(False)
    save(fig, "06_matriz_operabilidade.png")


# ─────────────────────────────────────────────────────────────
# 07 · Arranjo típico de amarração de um Panamax (4-1-1 por extremidade)
#      Esquema elaborado a partir da Figura 9 do relatório do TF I
# ─────────────────────────────────────────────────────────────
def layout_referencia():
    COR_L, COR_T, COR_E = "#7F9CC8", "#CC6A58", "#2F8F86"
    gap, boca, y_cab = 12.0, 30.0, -4.0
    y0, yc = gap, gap + boca / 2
    fig, ax = plt.subplots(figsize=(11.7, 2.75))
    ax.axhspan(-17, 0, color="#E9E3D8", lw=0)
    ax.text(-178, -13.5, "CAIS", fontsize=9, color="#8A8275", weight="bold")
    # casco simplificado: proa à esquerda
    casco = [(-112, yc), (-104, y0 + 3), (-90, y0), (95, y0), (108, y0 + 4), (112, y0 + 8),
             (112, y0 + boca - 8), (108, y0 + boca - 4), (95, y0 + boca), (-90, y0 + boca),
             (-104, y0 + boca - 3)]
    ax.add_patch(Polygon(casco, closed=True, fc="#F4EFE7", ec="#A49C90", lw=1.0, zorder=1))
    ax.text(0, yc, "4 lançantes + 1 través + 1 espringue em cada extremidade (6 + 6 cabos)",
            ha="center", va="center", fontsize=10, color="#5E574D")
    ax.text(-97, yc, "PROA", ha="center", va="center", fontsize=8.5, weight="bold", color="#A49C90")
    ax.text(97, yc, "POPA", ha="center", va="center", fontsize=8.5, weight="bold", color="#A49C90")
    for x in range(-165, 166, 15):
        ax.plot(x, y_cab, "o", ms=2.6, color="#B9B2A6", zorder=2)

    def cabo(xa, ya, xb, cor, lw=1.6):
        ax.plot([xa, xb], [ya, y_cab], color=cor, lw=lw, zorder=3)
        ax.plot(xb, y_cab, "o", ms=5.5, color="#2F3A48", zorder=4)

    for s in (-1, 1):                     # -1 = proa (esquerda), +1 = popa (direita)
        for xa, ya, xb in [(108, yc - 3, 165), (108, yc + 3, 165), (104, y0 + 3, 150), (101, y0 + 2, 150)]:
            cabo(s * xa, ya, s * xb, COR_L)
        cabo(s * 88, y0, s * 90, COR_T, lw=2.2)
        cabo(s * 70, y0, s * 30, COR_E, lw=2.0)
    lab = dict(fontsize=8.5, weight="bold", ha="center", va="center", zorder=6,
               bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.92))
    for s, nome in ((-1, "proa"), (1, "popa")):
        ax.text(s * 143, 13, f"lançantes\nde {nome}", color="#3F5F96", **lab)
        ax.text(s * 89, 4.5, "través", color=COR_T, **lab)
        ax.text(s * 50, 4.5, "espringue", color=COR_E, **lab)
    seta = dict(arrowstyle="<->", color="#5E574D", lw=1.1)
    ax.annotate("", xy=(-125, y0 + boca + 10), xytext=(-125, y0 + boca + 1), arrowprops=seta)
    ax.text(-121, y0 + boca + 5.5, "afastamento (sway): través e lançantes", fontsize=9.5,
            va="center", color="#5E574D")
    ax.annotate("", xy=(55, y0 + boca + 5.5), xytext=(25, y0 + boca + 5.5), arrowprops=seta)
    ax.text(59, y0 + boca + 5.5, "avanço e recuo (surge): espringues", fontsize=9.5,
            va="center", color="#5E574D")
    ax.text(178, -13.5, "Esquema fora de escala na direção transversal", ha="right",
            fontsize=8, style="italic", color="#8A8275")
    ax.set_xlim(-180, 180)
    ax.set_ylim(-17, y0 + boca + 12)
    ax.set_aspect("equal")
    ax.axis("off")
    save(fig, "07_layout_panamax.png")


# ─────────────────────────────────────────────────────────────
# 08 · Maré em Santos, 2025 (Tabela 7 do TF I, dados da DHN)
# ─────────────────────────────────────────────────────────────
MARE_2025 = [  # mês, mínima, máxima, média [m]
    ("jan", -0.12, 1.51, 0.77), ("fev", -0.16, 1.52, 0.77), ("mar", -0.12, 1.49, 0.77),
    ("abr", -0.01, 1.49, 0.78), ("mai", 0.06, 1.46, 0.78), ("jun", 0.07, 1.43, 0.77),
    ("jul", -0.02, 1.44, 0.77), ("ago", -0.10, 1.48, 0.77), ("set", -0.14, 1.50, 0.78),
    ("out", -0.09, 1.50, 0.77), ("nov", 0.02, 1.50, 0.78), ("dez", 0.03, 1.48, 0.77)]


def mare_dhn():
    meses = [m[0] for m in MARE_2025]
    mn = np.array([m[1] for m in MARE_2025])
    mx = np.array([m[2] for m in MARE_2025])
    md = np.array([m[3] for m in MARE_2025])
    x = np.arange(12)
    fig, ax = plt.subplots(figsize=FIG, layout="constrained")
    ax.bar(x, mx - mn, bottom=mn, width=0.56, color="#BFDCD7", edgecolor="#2F8F86", lw=0.8,
           label="Faixa mensal (mínima a máxima)", zorder=2)
    ax.plot(x, md, "o", ms=5, color="#2F8F86", label="Média mensal", zorder=3)
    ax.axhline(1.58, color="#7A63C2", ls="--", lw=1.2, label="Maré alta usada na geometria (1,58 m)")
    ax.axhline(0, color="#7B8290", lw=0.9, zorder=1)
    ax.plot([1, 1], [-0.16, 1.52], ls="none", marker="D", ms=5.5, color="#C8605A",
            label="Extremos de 2025 (fev): −0,16 m e 1,52 m", zorder=4)
    ax.text(-0.012, -0.5, "amplitude", fontsize=9, color="#7B8290", va="center", ha="right",
            transform=matplotlib.transforms.blended_transform_factory(ax.transAxes, ax.transData))
    for xi, a, b in zip(x, mn, mx):
        ax.text(xi, -0.5, br(b - a, 2), ha="center", va="center", fontsize=9,
                weight="bold" if xi in (1, 5) else "normal", color="#3A4250")
    ax.set_xticks(x, meses)
    ax.set_yticks([0, 0.5, 1.0, 1.5, 2.0])
    virgula(ax.yaxis, 1)
    ax.set_xlim(-0.6, 11.6)
    ax.set_ylabel("Nível da maré [m]")
    ax.set_ylim(-0.65, 2.25)
    ax.set_title("Maré prevista em Santos em 2025: faixa de cada mês [m]", loc="left")
    ax.grid(axis="x", visible=False)
    ax.legend(loc="upper left", fontsize=8.5, ncol=2)
    save(fig, "08_mare_dhn.png")


if __name__ == "__main__":
    layout_referencia()
    mare_dhn()
    plano_amarracao()
    histogramas()
    influencia_mare()
    correlacao_boxplot()
    envelope()
    operabilidade()
    print("Gráficos gerados em", OUT)
    for f in sorted(os.listdir(OUT)):
        print("  ", f)
