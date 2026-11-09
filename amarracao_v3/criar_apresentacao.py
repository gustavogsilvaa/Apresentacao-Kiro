"""
Apresentação: leitura dos gráficos da análise de amarração (TF4.0 - Colab).
Estilo baseado no PDF do PINN: fundo creme, barra lateral colorida por seção,
título serifado, cards claros com bordas suaves e formas decorativas no canto.

Imagens: usa imagens_originais/<nome>.png se existir (PNGs exportados do Colab);
caso contrário usa imagens_recriadas/<nome>.png (gerar_graficos.py).
"""
import os
import re
import sys

sys.path.insert(0, "/projects/sandbox/pptx_pkgs")
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
SAIDA = os.path.join(BASE, "Apresentacao_Amarracao_v3.pptx")

# ── Paleta ────────────────────────────────────────────────────
def hx(s):
    return RGBColor.from_string(s.lstrip("#"))

CREME, CARD, BORDA = "#F4F1EA", "#FBFAF7", "#DDD6C8"
TXT, TXT2, TXT3 = "#1C2430", "#4A5363", "#7B8290"
SECAO = {                      # cor de destaque por seção
    "capa":      "#F2B57E",
    "contexto":  "#B9862C",
    "visao":     "#2F8F86",
    "fatores":   "#4C78C8",
    "riscos":    "#CC6A58",
    "geometria": "#7A63C2",
    "fim":       "#2F8F86",
}
ALERTA_FILL, DESTAQUE_FILL = "#F8E1DB", "#E9E3F6"


def tint(cor, f):
    a, b = hx(CREME), hx(cor)
    return RGBColor(*[round(a[i] * (1 - f) + b[i] * f) for i in range(3)])


W, H = 13.333, 7.5
X0, Y_CONT, Y_FIM = 0.75, 2.0, 6.95
TOTAL = 13
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)


# ── Primitivas ────────────────────────────────────────────────
def forma(sl, tipo, x, y, w, h, fill, linha=None, raio=None):
    s = sl.shapes.add_shape(tipo, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill if isinstance(fill, RGBColor) else hx(fill)
    if linha:
        s.line.color.rgb = hx(linha)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    if raio is not None:
        s.adjustments[0] = raio
    s.shadow.inherit = False
    s.text_frame.text = ""
    return s


def runs(par, texto, size, cor, fonte="Arial", bold=False, spc=None):
    """Escreve texto num parágrafo; trechos entre **...** ficam em negrito."""
    for i, pedaco in enumerate(re.split(r"\*\*(.+?)\*\*", texto)):
        if not pedaco:
            continue
        r = par.add_run()
        r.text = pedaco
        f = r.font
        f.size, f.name, f.bold = Pt(size), fonte, bold or (i % 2 == 1)
        f.color.rgb = hx(cor)
        if spc:
            r._r.get_or_add_rPr().set("spc", str(spc))


def texto(sl, x, y, w, h, paragrafos, size=11, cor=TXT2, fonte="Arial", bold=False,
          align=PP_ALIGN.LEFT, spc=None, entre=4, ancora=MSO_ANCHOR.TOP, linha=1.1):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = ancora
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(paragrafos, str):
        paragrafos = [paragrafos]
    for i, ptxt in enumerate(paragrafos):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = linha
        if i:
            p.space_before = Pt(entre)
        runs(p, ptxt, size, cor, fonte, bold, spc)
    return tb


def card(sl, x, y, w, h, fill=CARD):
    return forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, BORDA, raio=0.045)


def base(sl_secao, n, rotulo=None, titulo=None, subtitulo=None, fonte_nota=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    cor = SECAO[sl_secao]
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
    # formas decorativas no canto inferior direito (como no PDF de referência)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 10.25, 4.45, 3.6, 3.6, tint(cor, 0.10), raio=0.3)
    forma(sl, MSO_SHAPE.OVAL, 11.3, 5.6, 3.8, 3.8, tint(cor, 0.20))
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, cor)
    if rotulo:
        texto(sl, X0, 0.55, 10, 0.3, rotulo.upper(), 10, cor, bold=True, spc=250)
    if titulo:
        texto(sl, X0, 0.82, 11.8, 0.7, titulo, 30, TXT, fonte="Georgia")
    if subtitulo:
        texto(sl, X0, 1.5, 11.8, 0.4, subtitulo, 13, TXT3)
    if fonte_nota:
        texto(sl, X0, 7.08, 9.5, 0.25, fonte_nota, 8.5, TXT3)
    texto(sl, 11.6, 7.06, 1.4, 0.25, f"{n} / {TOTAL}", 10, TXT3, align=PP_ALIGN.RIGHT)
    return sl, cor


def card_texto(sl, x, y, w, h, rotulo, corpo, cor, kpi=None, kpi_legenda=None, size=10.5):
    card(sl, x, y, w, h)
    pad = 0.2
    texto(sl, x + pad, y + 0.18, w - 2 * pad, 0.25, rotulo.upper(), 9, cor, bold=True, spc=150)
    y_corpo = y + 0.48
    if kpi:
        texto(sl, x + pad, y + 0.42, w - 2 * pad, 0.45, kpi, 22, cor, fonte="Georgia", bold=True)
        texto(sl, x + pad, y + 0.88, w - 2 * pad, 0.25, kpi_legenda, 9.5, TXT3)
        y_corpo = y + 1.22
    texto(sl, x + pad, y_corpo, w - 2 * pad, y + h - y_corpo - 0.12, corpo, size, TXT2)


def imagem(nome):
    for pasta in ("imagens_originais", "imagens_recriadas"):
        p = os.path.join(BASE, pasta, nome)
        if os.path.exists(p):
            return p, pasta == "imagens_originais"
    raise FileNotFoundError(nome)


def por_imagem(sl, nome, x, y, w, h, pad=0.14):
    """Card branco com a imagem centralizada, mantendo a proporção."""
    card(sl, x, y, w, h, fill="#FFFFFF")
    caminho, original = imagem(nome)
    iw, ih = Image.open(caminho).size
    bw, bh = w - 2 * pad, h - 2 * pad
    esc = min(bw / iw, bh / ih)
    pw, ph = iw * esc, ih * esc
    sl.shapes.add_picture(caminho, Inches(x + pad + (bw - pw) / 2), Inches(y + pad + (bh - ph) / 2),
                          Inches(pw), Inches(ph))
    return original


def tabela(sl, x, y, w, cabec, linhas, larguras, alt_linha=0.34, size=10,
           cor_cabec=TXT, destaques=None):
    """destaques: {(linha, coluna|None): cor_fundo}; coluna None = linha inteira."""
    destaques = destaques or {}
    shp = sl.shapes.add_table(len(linhas) + 1, len(cabec), Inches(x), Inches(y),
                              Inches(w), Inches(alt_linha * (len(linhas) + 1)))
    tbl = shp.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("bandRow", "0")
    tblPr.set("firstRow", "0")
    soma = sum(larguras)
    for j, lw in enumerate(larguras):
        tbl.columns[j].width = Inches(w * lw / soma)
    for i in range(len(linhas) + 1):
        tbl.rows[i].height = Inches(alt_linha)
        for j in range(len(cabec)):
            c = tbl.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.05)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            if i == 0:
                val, fill, cor_t, b = cabec[j], cor_cabec, "#FFFFFF", True
            else:
                val = linhas[i - 1][j]
                fill = destaques.get((i - 1, j)) or destaques.get((i - 1, None)) or \
                    (CARD if i % 2 else "#F3EFE7")
                cor_t, b = TXT, (i - 1, j) in destaques
            c.fill.solid()
            c.fill.fore_color.rgb = hx(fill)
            tf = c.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            runs(p, str(val), size if i else size - 0.5, cor_t, bold=b)
    return shp


def nota_imagem(original, extra=""):
    base_txt = "Fonte: TF4.0 – Colab (calado 9 m, 394 condições)"
    if not original:
        base_txt += " · gráfico redesenhado a partir das saídas do notebook"
    return base_txt + extra


# ── Layout padrão: gráfico à esquerda, dois cards à direita ───
IMG = dict(x=X0, y=Y_CONT, w=7.05, h=Y_FIM - Y_CONT)
COL = dict(x=8.0, w=4.6)
C1_H, GAP = 2.05, 0.15


def slide_grafico(secao, n, rotulo, titulo, subtitulo, nome_img, le, resultado, kpi, kpi_leg):
    sl, cor = base(secao, n, rotulo, titulo, subtitulo)
    original = por_imagem(sl, nome_img, **IMG)
    colunas_texto(sl, cor, le, resultado, kpi, kpi_leg)
    texto(sl, X0, 7.08, 9.5, 0.25, nota_imagem(original), 8.5, TXT3)
    return sl


def colunas_texto(sl, cor, le, resultado, kpi, kpi_leg, rotulo="O que o gráfico lê"):
    card_texto(sl, COL["x"], Y_CONT, COL["w"], C1_H, rotulo, le, cor)
    y2 = Y_CONT + C1_H + GAP
    card_texto(sl, COL["x"], y2, COL["w"], Y_FIM - y2, "Resultado obtido", resultado, cor,
               kpi=kpi, kpi_legenda=kpi_leg)


# ══════════════════════════════════════════════════════════════
# 1 · Capa
# ══════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(prs.slide_layouts[6])
forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 9.7, 3.9, 4.6, 4.6, tint(SECAO["capa"], 0.18), raio=0.25)
forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, SECAO["capa"])
texto(sl, X0 + 0.05, 1.95, 11, 0.3, "TF4.0 · ANÁLISE DE SISTEMAS DE AMARRAÇÃO", 11, TXT,
      bold=True, spc=300)
texto(sl, X0, 2.4, 11.5, 1.9, ["Sistema de amarração", "de um navio Panamax"], 44, TXT,
      fonte="Georgia", bold=True, entre=0, linha=1.0)
texto(sl, X0 + 0.05, 4.45, 11, 0.4,
      "Leitura dos gráficos de esforços · calado 9 m · 394 condições simuladas", 16, TXT, bold=True)
texto(sl, X0 + 0.05, 5.1, 8, 0.35, "Gustavo Guilherme Silva", 13.5, TXT2)
texto(sl, 11.6, 7.06, 1.4, 0.25, f"1 / {TOTAL}", 10, TXT3, align=PP_ALIGN.RIGHT)

# ══════════════════════════════════════════════════════════════
# 2 · Contexto: plano de amarração e critérios
# ══════════════════════════════════════════════════════════════
sl, cor = base("contexto", 2, "Etapa 1 · Contexto", "Plano de amarração e critérios",
               "O arranjo analisado e os dois limites usados em todos os gráficos",
               "Fonte: TF4.0 – Colab · esquema desenhado com as coordenadas de buzinas e cabeços "
               "do bloco de geometria")
por_imagem(sl, "00_plano_amarracao.png", X0, Y_CONT, 11.85, 2.4, pad=0.1)
cw, cy, ch = (11.85 - 0.3) / 3, 4.55, Y_FIM - 4.55
card_texto(sl, X0, cy, cw, ch, "O que foi simulado",
           ["394 combinações de **vento, corrente, onda (Hs, Tp) e maré**, com o navio atracado "
            "por boreste e calado de 9 m.",
            "Em cada condição o software registra a tração nos **12 cabos**, a reação nas "
            "defensas e a força nos **5 cabeços** usados (A, B, G, N, S)."], cor)
card_texto(sl, X0 + cw + 0.15, cy, cw, ch, "Critério · cabos",
           "MBL é a carga mínima de ruptura do cabo. Acima de 55% MBL a condição "
           "é tratada como risco de ruptura (recomendação OCIMF).", cor,
           kpi="55% MBL", kpi_legenda="limite operacional")
card_texto(sl, X0 + 2 * (cw + 0.15), cy, cw, ch, "Critério · cabeços",
           "SWL é a carga de trabalho segura do cabeço (75 t neste cais, pelos valores "
           "do notebook). Acima de 100% SWL há risco de arrancamento.", cor,
           kpi="100% SWL", kpi_legenda="limite estrutural")

# ══════════════════════════════════════════════════════════════
# 3 · Resumo estatístico (tabela)
# ══════════════════════════════════════════════════════════════
sl, cor = base("visao", 3, "Etapa 2 · Visão geral", "Resumo estatístico geral",
               "Escala dos esforços nas 394 condições, antes de olhar cada gráfico",
               "Fonte: TF4.0 – Colab · tabela “Resumo estatístico geral” (describe com "
               "percentis 50, 90, 95 e 99)")
card(sl, **IMG)
cab = ["Variável", "Média", "Mediana", "P90", "P95", "P99", "Máximo"]
lin = [["Tração máx. [tf]", "25,15", "23,56", "44,73", "47,25", "51,82", "56,44"],
       ["% MBL", "33,54", "31,45", "59,61", "63,00", "69,10", "75,30"],
       ["Reação máx. [tf]", "17,43", "16,61", "28,48", "32,41", "43,72", "51,86"],
       ["Compressão máx. [m]", "0,04", "0,04", "0,07", "0,08", "0,11", "0,13"],
       ["Cabeço máx. [tf]", "34,50", "31,88", "48,56", "55,63", "68,17", "83,13"],
       ["% SWL", "46,00", "42,50", "64,77", "74,17", "90,91", "110,80"]]
dest = {(1, j): ALERTA_FILL for j in (3, 4, 5, 6)}
dest[(5, 6)] = ALERTA_FILL
tabela(sl, X0 + 0.25, Y_CONT + 0.35, IMG["w"] - 0.5, cab, lin, [2.7, 1, 1.1, 0.95, 0.95, 0.95, 1.05],
       alt_linha=0.5, size=11, cor_cabec=cor, destaques=dest)
forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, X0 + 0.25, Y_CONT + 4.05, 0.28, 0.2, ALERTA_FILL, BORDA, raio=0.2)
texto(sl, X0 + 0.62, Y_CONT + 4.03, 6, 0.25,
      "acima do limite (55% MBL nos cabos, 100% SWL nos cabeços)", 9.5, TXT3)
colunas_texto(sl, cor,
              ["Estatísticas de seis grandezas: tração nos cabos (tf e % MBL), reação e "
               "compressão nas defensas, força nos cabeços (tf e % SWL).",
               "**Intuito:** dar a ordem de grandeza dos esforços. Os percentis P90 a P99 "
               "mostram o comportamento nas condições mais severas."],
              ["Metade das condições fica abaixo de 31,5% MBL, mas o **P90 já passa do limite** "
               "(59,6%) e o máximo chega a 75,3%.",
               "Nos cabeços a mediana é 42,5% SWL e só o máximo (110,8%) passa de 100%. "
               "O notebook não define limite para as defensas."],
              "31,5% MBL", "mediana da tração nos cabos",
              rotulo="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 4 · Histogramas
# ══════════════════════════════════════════════════════════════
slide_grafico("visao", 4, "Etapa 2 · Visão geral", "Distribuição dos esforços",
              "Com que frequência cabos e cabeços operam perto do limite",
              "01_histogramas.png",
              ["Conta quantas condições caem em cada faixa de carregamento: % MBL nos cabos e "
               "% SWL nos cabeços. A linha tracejada marca o limite.",
               "**Intuito:** mostrar se as violações são casos isolados ou frequentes."],
              ["Nos cabos, a faixa mais comum é 10–16% MBL (92 condições), mas a distribuição "
               "se estende até 75%: **58 condições ficam à direita do limite**.",
               "Nos cabeços, a faixa mais comum é 30–37% SWL (103) e **só 2 condições passam "
               "de 100%**. O risco nos cabos é recorrente; nos cabeços, pontual."],
              "58 de 394", "condições acima de 55% MBL (14,7%)")

# ══════════════════════════════════════════════════════════════
# 5 · Influência da maré
# ══════════════════════════════════════════════════════════════
slide_grafico("fatores", 5, "Etapa 3 · Fatores de carga", "Influência da maré",
              "Pico de cada elemento e o nível de maré em que ele acontece",
              "02_influencia_mare.png",
              ["Mostra o maior esforço de cada cabo (círculos) e de cada cabeço (triângulos). "
               "A cor indica a maré no instante desse pico.",
               "**Intuito:** apontar quais elementos chegam ao limite e se isso ocorre com "
               "maré alta ou baixa."],
              ["**Só o cabo 9 passa de 55% MBL**; os outros 11 ficam abaixo de 40%. O pico "
               "dele ocorre com maré perto de 1,5 m.",
               "Nos cabeços, só o S passa de 100% SWL, e com **maré baixa** (cerca de 0,3 m). "
               "A maré não age no mesmo sentido em todos os elementos."],
              "75,3% MBL", "pico do cabo 9 (ligado ao cabeço N)")

# ══════════════════════════════════════════════════════════════
# 6 · Correlação + boxplot
# ══════════════════════════════════════════════════════════════
slide_grafico("fatores", 6, "Etapa 3 · Fatores de carga", "Clima ou arranjo de cabos?",
              "Matriz de correlação e distribuição de tração por cabo crítico",
              "04_correlacao_boxplot.png",
              ["À esquerda, a correlação linear entre ambiente e esforços (perto de ±1 = forte, "
               "perto de 0 = fraca). À direita, a % MBL agrupada pelo cabo mais carregado em "
               "cada condição.",
               "**Intuito:** separar o efeito do clima do efeito do arranjo."],
              ["Nenhuma variável de ambiente explica sozinha a tração nos cabos (|r| < 0,1). "
               "Vento (0,25) e corrente (0,20) têm relação fraca com os cabeços. As direções "
               "não entram na matriz.",
               "No boxplot, **só o cabo 9 cruza 55%**: o problema está concentrado nele, o que "
               "aponta para o arranjo das linhas."],
              "r = 0,70", "correlação entre % MBL e % SWL")

# ══════════════════════════════════════════════════════════════
# 7 · Combinações críticas (tabelas)
# ══════════════════════════════════════════════════════════════
sl, cor = base("riscos", 7, "Etapa 4 · Riscos", "Combinações críticas",
               "As condições de ambiente que levam cabos e cabeços além do limite",
               "Fonte: TF4.0 – Colab · bloco “Filtro de riscos e piores combinações”")
card(sl, **IMG)
tx, tw = X0 + 0.25, IMG["w"] - 0.5
texto(sl, tx, Y_CONT + 0.22, tw, 0.3, "Cabos · 5 piores condições (de 58 acima de 55% MBL)",
      11, TXT, bold=True)
tabela(sl, tx, Y_CONT + 0.58, tw,
       ["Cond.", "Vento [nós]", "Corrente [nós]", "Hs [m]", "Tração [tf]", "% MBL", "Cabo"],
       [["259", "8,20", "1,57", "1,00", "56,44", "75,30", "9"],
        ["310", "41,40", "1,16", "1,19", "54,72", "73,00", "9"],
        ["78", "5,30", "1,43", "0,35", "53,95", "71,90", "9"],
        ["282", "29,40", "1,58", "1,08", "51,82", "69,10", "9"],
        ["272", "39,90", "0,47", "1,05", "51,85", "69,10", "9"]],
       [0.8, 1.1, 1.25, 0.85, 1.1, 0.9, 0.7], alt_linha=0.36, cor_cabec=cor,
       destaques={(i, 5): ALERTA_FILL for i in range(5)})
texto(sl, tx, Y_CONT + 2.95, tw, 0.3, "Cabeços · todas as condições acima de 100% SWL",
      11, TXT, bold=True)
tabela(sl, tx, Y_CONT + 3.31, tw,
       ["Cond.", "Vento [nós]", "Corrente [nós]", "Hs [m]", "Força [tf]", "% SWL", "Cabeço"],
       [["198", "36,70", "1,58", "0,78", "83,13", "110,80", "S"],
        ["55", "49,10", "0,88", "0,30", "78,34", "104,50", "S"]],
       [0.8, 1.1, 1.25, 0.85, 1.1, 0.9, 0.7], alt_linha=0.36, cor_cabec=cor,
       destaques={(i, 5): ALERTA_FILL for i in range(2)})
colunas_texto(sl, cor,
              ["Lista as condições que mais carregam os cabos e todas as que passam do "
               "limite nos cabeços, com vento, corrente e altura de onda (Hs).",
               "**Intuito:** ver quais combinações de ambiente levam ao limite e qual "
               "elemento atinge o limite primeiro."],
              ["O pior caso (cond. 259) ocorre com **vento de só 8,2 nós**, mas corrente de "
               "1,57 nós e Hs de 1,0 m. Quatro das cinco têm Hs de 1 m ou mais.",
               "As duas violações do cabeço S têm **vento forte** (36,7 e 49,1 nós)."],
              "5 de 5", "piores condições têm o cabo 9 como crítico",
              rotulo="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 8 · Envelope crítico
# ══════════════════════════════════════════════════════════════
slide_grafico("riscos", 8, "Etapa 4 · Riscos", "Envelope crítico: vento × corrente",
              "Em que combinação de ambiente cada elemento atinge seu pico",
              "05_envelope_critico.png",
              ["Cada bolha é a condição em que um elemento atingiu o pico, posicionada por "
               "corrente (X) e vento (Y). A cor indica o esforço e o tamanho, a maré.",
               "**Intuito:** mapear em que ambiente cada parte do sistema é mais exigida."],
              ["A maioria dos cabos atinge o pico com vento de 37 a 49 nós. **O cabo 9 é a "
               "exceção**: vento fraco, corrente de 1,57 nós e maré alta.",
               "Os cabeços A, B e S têm o pico na mesma condição (36,7 nós, maré baixa). O "
               "cabeço N tem o pico junto com o cabo 9, que está amarrado nele."],
              "8,2 nós", "vento na condição de pico do cabo 9")

# ══════════════════════════════════════════════════════════════
# 9 · Matriz de operabilidade
# ══════════════════════════════════════════════════════════════
slide_grafico("riscos", 9, "Etapa 4 · Riscos", "Matriz de operabilidade",
              "Pior tração em cada faixa de vento e corrente: base para limites de paragem",
              "06_matriz_operabilidade.png",
              ["Agrupa as condições em faixas de vento (10 nós) e corrente (0,5 nó). Cada "
               "célula mostra a pior tração da faixa: verde abaixo de 55%, vermelho acima.",
               "**Intuito:** apoiar a definição de quando suspender a operação no berço."],
              ["Quase toda faixa tem ao menos uma condição acima do limite, **inclusive com "
               "vento e corrente fracos** (61,2%).",
               "Vento e corrente sozinhos não definem uma janela segura: onda, direção e maré "
               "precisam entrar no critério. A célula verde (38,8%) merece conferir quantas "
               "condições tem."],
              "19 de 20", "faixas com pior caso acima de 55% MBL")

# ══════════════════════════════════════════════════════════════
# 10–11 · Geometria das linhas
# ══════════════════════════════════════════════════════════════
GEO = [  # cabo, buzina, cabeço, comp 1,58 m, âng 1,58 m, comp 0 m, âng 0 m
    (1, 2, "S", 24.18, 15.55, 23.80, 11.88), (2, 3, "S", 23.68, 15.88, 23.30, 12.14),
    (3, 14, "A", 31.27, 11.96, 30.98, 9.10), (4, 15, "A", 30.66, 12.20, 30.37, 9.28),
    (5, 17, "B", 17.60, 21.61, 17.08, 16.67), (6, 18, "B", 23.04, 16.34, 22.64, 12.50),
    (7, 23, "G", 41.76, 8.93, 41.54, 6.77), (8, 24, "G", 41.02, 9.09, 40.81, 6.90),
    (9, 28, "N", 31.00, 12.07, 30.71, 9.18), (10, 29, "N", 40.44, 9.22, 40.22, 7.00),
    (11, 31, "S", 31.37, 11.92, 31.08, 9.07), (12, 32, "S", 30.29, 12.35, 29.99, 9.40)]
f = lambda v, d=2: f"{v:.{d}f}".replace(".", ",")
deltas = [g[4] - g[6] for g in GEO]
assert abs(min(deltas) - 2.16) < 1e-9 and abs(max(deltas) - 4.94) < 1e-9


def legenda_geo(sl, y):
    for k, (fill, txt) in enumerate([(ALERTA_FILL, "cabo 9: pico de 75,3% MBL"),
                                     (DESTAQUE_FILL, "cabo 5: maior ângulo vertical")]):
        xx = X0 + 0.25 + k * 3.3
        forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, y, 0.28, 0.2, fill, BORDA, raio=0.2)
        texto(sl, xx + 0.37, y - 0.02, 2.9, 0.25, txt, 9.5, TXT3)


sl, cor = base("geometria", 10, "Etapa 5 · Geometria", "Geometria das linhas · maré 1,58 m",
               "Calado 9 m, maré alta, atracação por boreste",
               "Fonte: TF4.0 – Colab · calcular_geometria_amarracao(calado 9 m, maré 1,58 m) · "
               "coluna “Cabo” segue a ordem do plano de amarração")
card(sl, **IMG)
tabela(sl, X0 + 0.25, Y_CONT + 0.2, IMG["w"] - 0.5,
       ["Cabo", "Buzina", "Cabeço", "Comp. 3D [m]", "Âng. vertical [°]"],
       [[g[0], g[1], g[2], f(g[3]), f(g[4])] for g in GEO],
       [0.8, 0.9, 0.9, 1.4, 1.5], alt_linha=0.31, cor_cabec=cor,
       destaques={(8, None): ALERTA_FILL, (4, None): DESTAQUE_FILL, (8, 3): ALERTA_FILL,
                  (9, 3): "#F3EFE7", (4, 4): DESTAQUE_FILL})
legenda_geo(sl, Y_CONT + 4.45)
colunas_texto(sl, cor,
              ["Comprimento 3D e ângulo vertical de cada linha, da buzina do navio ao cabeço, "
               "no cenário de maior pico do cabo 9.",
               "**Intuito:** ligar o esforço de cada cabo à sua geometria. Linhas mais curtas "
               "são mais rígidas; linhas mais inclinadas seguram menos na horizontal."],
              ["O **cabo 9 é 9,4 m mais curto que o cabo 10**, que divide o cabeço N. Mais "
               "rígido, tende a absorver mais carga, coerente com os picos.",
               "A linha mais inclinada é a do cabo 5 (21,6°). O cabeço S recebe **4 das 12 "
               "linhas** (cabos 1, 2, 11 e 12)."],
              "31,0 × 40,4 m", "comprimento dos cabos 9 e 10 (cabeço N)",
              rotulo="O que a tabela lê")

sl, cor = base("geometria", 11, "Etapa 5 · Geometria", "Geometria das linhas · maré 0 m",
               "Mesmo calado com maré baixa: efeito isolado da maré",
               "Fonte: TF4.0 – Colab · calcular_geometria_amarracao(calado 9 m, maré 0 m) · "
               "Δ calculado contra a tabela de maré 1,58 m")
card(sl, **IMG)
tabela(sl, X0 + 0.25, Y_CONT + 0.2, IMG["w"] - 0.5,
       ["Cabo", "Buzina", "Cabeço", "Comp. 3D [m]", "Âng. vertical [°]", "Δ ângulo [°]"],
       [[g[0], g[1], g[2], f(g[5]), f(g[6]), "−" + f(g[4] - g[6])] for g in GEO],
       [0.75, 0.85, 0.85, 1.25, 1.4, 1.15], alt_linha=0.31, cor_cabec=cor,
       destaques={(8, None): ALERTA_FILL, (4, None): DESTAQUE_FILL, (4, 5): DESTAQUE_FILL})
legenda_geo(sl, Y_CONT + 4.45)
colunas_texto(sl, cor,
              ["Repete o cálculo com maré 0 m. A última coluna mostra quanto o ângulo "
               "vertical caiu em relação à maré de 1,58 m.",
               "**Intuito:** isolar o efeito da maré na geometria, com o calado fixo."],
              ["Com maré baixa a buzina fica 1,58 m mais perto da altura do cabeço e **todas "
               "as linhas ficam menos inclinadas**. O comprimento muda só 0,2 a 0,5 m.",
               "A diferença entre os cabos 9 e 10 continua: a maré muda a inclinação, mas "
               "não resolve a assimetria."],
              "−2,2° a −4,9°", "queda do ângulo vertical com a maré baixa",
              rotulo="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 12 · Conclusões
# ══════════════════════════════════════════════════════════════
sl, cor = base("fim", 12, "Etapa 6 · Síntese", "Conclusões",
               "O que os gráficos mostram em conjunto",
               "Fonte: TF4.0 – Colab · próximos passos são sugestões a partir dos resultados")
achados = [
    ("C1", "O cabo 9 concentra o risco",
     "Único cabo acima de 55% MBL e presente nas 5 piores condições (pico de 75,3%)."),
    ("C2", "O cabeço S passa do SWL",
     "2 condições acima de 100% (até 110,8%, cerca de 83 t para SWL de 75 t). Recebe 4 das 12 linhas."),
    ("C3", "O ambiente sozinho não explica",
     "Correlações fracas e 19 de 20 faixas vento × corrente acima do limite."),
    ("C4", "A geometria explica parte",
     "Cabo 9 é 9,4 m mais curto que o cabo 10 no mesmo cabeço. Maré alta inclina as linhas 2 a 5° a mais."),
]
gw, gh = (11.85 - 0.45) / 4, 3.0
for k, (cid, tit, corpo) in enumerate(achados):
    xx = X0 + k * (gw + 0.15)
    card(sl, xx, Y_CONT, gw, gh)
    texto(sl, xx + 0.2, Y_CONT + 0.2, gw - 0.4, 0.25, cid, 10, TXT3, bold=True)
    texto(sl, xx + 0.2, Y_CONT + 0.48, gw - 0.4, 0.6, tit, 14, cor, bold=True)
    texto(sl, xx + 0.2, Y_CONT + 1.2, gw - 0.4, gh - 1.3, corpo, 12, TXT2)
py = Y_CONT + gh + 0.2
card(sl, X0, py, 11.85, Y_FIM - py)
texto(sl, X0 + 0.25, py + 0.2, 11, 0.25, "PRÓXIMOS PASSOS", 9, cor, bold=True, spc=150)
texto(sl, X0 + 0.25, py + 0.55, 11.3, Y_FIM - py - 0.65,
      ["•  Rever comprimento ou pré-tensão do **cabo 9** para dividir a carga com o cabo 10.",
       "•  Redistribuir as linhas do **cabeço S** ou confirmar sua capacidade.",
       "•  Incluir **Hs, direções e maré** na matriz de operabilidade antes de definir limites de paragem."],
      11.5, TXT2, entre=6)

# ══════════════════════════════════════════════════════════════
# 13 · Encerramento
# ══════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(prs.slide_layouts[6])
forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 9.7, 3.9, 4.6, 4.6, tint(SECAO["capa"], 0.18), raio=0.25)
forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, SECAO["capa"])
texto(sl, X0, 2.45, 11, 1.0, "Obrigado pela atenção", 44, TXT, fonte="Georgia", bold=True)
texto(sl, X0, 3.85, 11, 1.0, "Perguntas?", 44, TXT, fonte="Georgia", bold=True)
texto(sl, 11.6, 7.06, 1.4, 0.25, f"{TOTAL} / {TOTAL}", 10, TXT3, align=PP_ALIGN.RIGHT)

assert len(prs.slides) == TOTAL, len(prs.slides)
prs.save(SAIDA)
print("Salvo:", SAIDA, "·", len(prs.slides), "slides")
