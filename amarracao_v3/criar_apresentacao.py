"""
Apresentação parcial do TF: sistema de amarração em conjuntos portuários.
Junta o relatório do TF I (motivação, justificativa, cenário, abordagem) com os
resultados do notebook TF4.0 - Colab (calado 9 m, 394 condições).

Estilo baseado no PDF do PINN: fundo creme, barra lateral colorida por etapa,
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
SAIDA = os.path.join(BASE, "Apresentacao_TF_Amarracao_parcial.pptx")


# ── Paleta ────────────────────────────────────────────────────
def hx(s):
    return RGBColor.from_string(s.lstrip("#"))


CREME, CARD, BORDA = "#F4F1EA", "#FBFAF7", "#DDD6C8"
TXT, TXT2, TXT3 = "#1C2430", "#4A5363", "#7B8290"
SECAO = {                      # uma cor por etapa, como no PDF de referência
    "capa":          "#F2B57E",
    "roteiro":       "#B08A2E",
    "motivacao":     "#DB7F3B",
    "justificativa": "#B08A2E",
    "cenario":       "#2F8F86",
    "abordagem":     "#4C78C8",
    "resultados":    "#7A63C2",
    "conclusoes":    "#C8605A",
}
ALERTA_FILL, DESTAQUE_FILL = "#F8E1DB", "#E9E3F6"


def tint(cor, f):
    a, b = hx(CREME), hx(cor)
    return RGBColor(*[round(a[i] * (1 - f) + b[i] * f) for i in range(3)])


W, H = 13.333, 7.5
X0, Y_CONT, Y_FIM = 0.75, 2.0, 6.95
LARG = 11.85                   # largura útil do conteúdo
TOTAL = 19
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


def rotulo(sl, x, y, w, txt, cor, size=9):
    texto(sl, x, y, w, 0.25, txt.upper(), size, cor, bold=True, spc=150)


def seta(sl, x, y, w, h, cor):
    forma(sl, MSO_SHAPE.CHEVRON, x, y, w, h, tint(cor, 0.55))


def numero(sl, n):
    texto(sl, 11.6, 7.06, 1.4, 0.25, f"{n} / {TOTAL}", 10, TXT3, align=PP_ALIGN.RIGHT)


def base(secao, n, rot=None, titulo=None, subtitulo=None, nota=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    cor = SECAO[secao]
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
    # formas decorativas no canto inferior direito (como no PDF de referência)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 10.25, 4.45, 3.6, 3.6, tint(cor, 0.10), raio=0.3)
    forma(sl, MSO_SHAPE.OVAL, 11.3, 5.6, 3.8, 3.8, tint(cor, 0.20))
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, cor)
    if rot:
        texto(sl, X0, 0.55, 10, 0.3, rot.upper(), 10, cor, bold=True, spc=250)
    if titulo:
        texto(sl, X0, 0.82, LARG, 0.7, titulo, 30, TXT, fonte="Georgia")
    if subtitulo:
        texto(sl, X0, 1.5, LARG, 0.4, subtitulo, 13, TXT3)
    if nota:
        texto(sl, X0, 7.08, 10.6, 0.25, nota, 8.5, TXT3)
    numero(sl, n)
    return sl, cor


def card_texto(sl, x, y, w, h, rot, corpo, cor, kpi=None, kpi_legenda=None, size=10.5):
    card(sl, x, y, w, h)
    pad = 0.2
    rotulo(sl, x + pad, y + 0.18, w - 2 * pad, rot, cor)
    y_corpo = y + 0.48
    if kpi:
        texto(sl, x + pad, y + 0.42, w - 2 * pad, 0.45, kpi, 22, cor, fonte="Georgia", bold=True)
        texto(sl, x + pad, y + 0.88, w - 2 * pad, 0.25, kpi_legenda, 9.5, TXT3)
        y_corpo = y + 1.22
    texto(sl, x + pad, y_corpo, w - 2 * pad, y + h - y_corpo - 0.12, corpo, size, TXT2)


def kpi_tile(sl, x, y, w, h, valor, legenda, cor):
    card(sl, x, y, w, h)
    texto(sl, x + 0.2, y + 0.16, w - 0.4, 0.5, valor, 24, cor, fonte="Georgia", bold=True)
    texto(sl, x + 0.2, y + 0.7, w - 0.4, h - 0.8, legenda, 10, TXT3)


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


def nota_imagem(original):
    t = "Fonte: TF4.0 – Colab (calado 9 m, 394 condições)"
    if not original:
        t += " · gráfico redesenhado a partir das saídas do notebook"
    return t


# ── Layout padrão de resultados: gráfico à esquerda, dois cards à direita ──
IMG = dict(x=X0, y=Y_CONT, w=7.05, h=Y_FIM - Y_CONT)
COL = dict(x=8.0, w=4.6)
C1_H, GAP = 2.05, 0.15
ROT_RES = "Etapa 5 · Resultados"


def colunas_texto(sl, cor, le, resultado, kpi, kpi_leg, rot="O que o gráfico lê"):
    card_texto(sl, COL["x"], Y_CONT, COL["w"], C1_H, rot, le, cor)
    y2 = Y_CONT + C1_H + GAP
    card_texto(sl, COL["x"], y2, COL["w"], Y_FIM - y2, "Resultado obtido", resultado, cor,
               kpi=kpi, kpi_legenda=kpi_leg)


def slide_grafico(n, titulo, subtitulo, nome_img, le, resultado, kpi, kpi_leg):
    sl, cor = base("resultados", n, ROT_RES, titulo, subtitulo)
    original = por_imagem(sl, nome_img, **IMG)
    colunas_texto(sl, cor, le, resultado, kpi, kpi_leg)
    texto(sl, X0, 7.08, 9.5, 0.25, nota_imagem(original), 8.5, TXT3)


def capa_fundo(sl):
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 9.7, 3.9, 4.6, 4.6, tint(SECAO["capa"], 0.18), raio=0.25)
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, SECAO["capa"])


# ══════════════════════════════════════════════════════════════
# 1 · Capa
# ══════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(prs.slide_layouts[6])
capa_fundo(sl)
texto(sl, X0 + 0.05, 1.7, 11, 0.3, "PNV3510 · TRABALHO DE FORMATURA · APRESENTAÇÃO PARCIAL", 11, TXT,
      bold=True, spc=300)
texto(sl, X0, 2.15, 11.5, 1.9, ["Sistema de amarração em", "conjuntos portuários"], 44, TXT,
      fonte="Georgia", bold=True, entre=0, linha=1.0)
texto(sl, X0 + 0.05, 4.2, 11, 0.4,
      "Esforços em um navio Panamax no terminal T-Grão, Porto de Santos", 16, TXT, bold=True)
texto(sl, X0 + 0.05, 4.85, 8, 0.65,
      ["Gustavo Guilherme Silva · NUSP 11820555",
       "Escola Politécnica da Universidade de São Paulo"], 13.5, TXT2, entre=2)
numero(sl, 1)

# ══════════════════════════════════════════════════════════════
# 2 · Roteiro
# ══════════════════════════════════════════════════════════════
sl, cor = base("roteiro", 2, "Apresentação", "Roteiro", "Do problema aos resultados parciais")
itens = [("motivacao", "Motivação", "Por que a amarração é crítica nos terminais de Santos"),
         ("justificativa", "Justificativa", "A lacuna no monitoramento dos cabos e o objetivo desta etapa"),
         ("cenario", "Cenário analisado", "Terminal T-Grão, navio Panamax e arranjo de amarração"),
         ("abordagem", "Abordagem", "Da caracterização do berço ao pós-processamento em Python"),
         ("resultados", "Resultados", "O que cada gráfico da simulação lê e o que mostra"),
         ("conclusoes", "Conclusões", "Achados parciais e próximos passos")]
gw, gh = (LARG - 0.3) / 3, (Y_FIM - Y_CONT - 0.15) / 2
for k, (sec, tit, desc) in enumerate(itens):
    xx = X0 + (k % 3) * (gw + 0.15)
    yy = Y_CONT + (k // 3) * (gh + 0.15)
    c = SECAO[sec]
    card(sl, xx, yy, gw, gh)
    forma(sl, MSO_SHAPE.RECTANGLE, xx, yy + 0.3, 0.07, gh - 0.6, c)
    texto(sl, xx + 0.3, yy + 0.25, 1.2, 0.6, f"{k + 1:02d}", 28, c, fonte="Georgia", bold=True)
    texto(sl, xx + 0.3, yy + 0.95, gw - 0.5, 0.35, tit, 15, TXT, bold=True)
    texto(sl, xx + 0.3, yy + 1.4, gw - 0.5, gh - 1.5, desc, 11, TXT2)

# ══════════════════════════════════════════════════════════════
# 3 · Motivação
# ══════════════════════════════════════════════════════════════
sl, cor = base("motivacao", 3, "Etapa 1 · Motivação", "Por que estudar a amarração?",
               "Durante o carregamento, o navio Panamax fica preso ao cais apenas pelos cabos",
               "Fonte: relatório do TF I (Tabela 1 e seções 1.3, 1.9 e 6.3)")
tw_ = (LARG - 0.45) / 4
for k, (v, leg) in enumerate([(">100 Mt", "de carga por ano no Porto de Santos, o maior da América Latina"),
                              ("14 de 60", "berços de Santos dedicados a granéis vegetais"),
                              ("USD 34 mi", "em acidentes com equipamentos de amarração em 20 anos (UK P&I Club)"),
                              ("até 40%", "de aumento da tração nos cabos causado por navios passantes")]):
    kpi_tile(sl, X0 + k * (tw_ + 0.15), Y_CONT, tw_, 1.35, v, leg, cor)
cw = (LARG - 0.3) / 3
cy = Y_CONT + 1.5
mh = 2.05
for k, (rot, corpo) in enumerate([
        ("Segurança",
         ["Um cabo que rompe sob tração libera de uma vez a energia acumulada (**snap-back**).",
          "Por isso os amarradores precisam ficar a pelo menos 7 m dos cabeços."]),
        ("Operação",
         ["Durante o carregamento, o navio fica sujeito a vento, corrente de maré e navios passantes.",
          "Se ele se afasta do cais, shiploaders e esteiras podem ser danificados e o terminal para."]),
        ("Meio ambiente",
         ["Grãos despejados no estuário elevam a demanda bioquímica de oxigênio (DBO) da água.",
          "O acidente exige dragagem de limpeza e afeta a meta de **Porto Verde** (ESG)."])]):
    card_texto(sl, X0 + k * (cw + 0.15), cy, cw, mh, rot, corpo, cor, size=11.5)
sy = cy + mh + 0.15
card(sl, X0, sy, LARG, Y_FIM - sy, fill=tint(cor, 0.08))
forma(sl, MSO_SHAPE.RECTANGLE, X0, sy + 0.2, 0.07, Y_FIM - sy - 0.4, cor)
texto(sl, X0 + 0.3, sy, LARG - 0.5, Y_FIM - sy,
      "**Em síntese:** uma falha de amarração atinge ao mesmo tempo pessoas, operação e meio ambiente. "
      "Conhecer os esforços nos cabos é o primeiro passo para evitá-la.", 12.5, TXT2, ancora=MSO_ANCHOR.MIDDLE)

# ══════════════════════════════════════════════════════════════
# 4 · Justificativa e objetivo
# ══════════════════════════════════════════════════════════════
sl, cor = base("justificativa", 4, "Etapa 2 · Justificativa", "Justificativa e objetivo",
               "Medir a tração nos cabos de forma contínua ainda é uma lacuna nos terminais",
               "Fonte: relatório do TF I (Resumo e seção 1)")
fw, fh, fgap = (LARG - 2 * 0.5) / 3, 2.35, 0.5
fluxo = [("O problema",
          "A tração nos cabos muda o tempo todo com vento, corrente, maré e navios passantes, "
          "mas quase nunca é medida durante a operação."),
         ("A lacuna",
          "Células de carga e extensômetros degradam rápido com sal e umidade. O custo de "
          "manutenção leva a inspeções visuais e reativas."),
         ("A proposta do TF",
          "Monitoramento de saúde estrutural (**SHM**) sem contato, por visão computacional e "
          "machine learning, integrado a simulações dos cenários críticos.")]
for k, (rot, corpo) in enumerate(fluxo):
    xx = X0 + k * (fw + fgap)
    card_texto(sl, xx, Y_CONT, fw, fh, rot, corpo, cor, size=12)
    if k < 2:
        seta(sl, xx + fw + 0.14, Y_CONT + fh / 2 - 0.2, 0.22, 0.4, cor)
oy = Y_CONT + fh + 0.2
card(sl, X0, oy, LARG, Y_FIM - oy)
rotulo(sl, X0 + 0.25, oy + 0.2, 6, "Objetivo desta etapa", cor)
texto(sl, X0 + 0.25, oy + 0.52, LARG - 0.5, 0.75,
      "Mapear, por simulação, em quais condições de vento, corrente, onda e maré os cabos e os "
      "cabeços do berço chegam ao limite. Esses cenários críticos são a base quantitativa que o "
      "monitoramento futuro precisa detectar.", 12, TXT2)
chips = ["Quantificar os esforços nas 394 condições",
         "Identificar condições e elementos críticos",
         "Ligar a carga à geometria das linhas"]
chw = (LARG - 0.5 - 0.3) / 3
for k, c_txt in enumerate(chips):
    xx = X0 + 0.25 + k * (chw + 0.15)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, Y_FIM - 0.72, chw, 0.48, tint(cor, 0.16), raio=0.3)
    texto(sl, xx, Y_FIM - 0.72, chw, 0.48, c_txt, 10.5, TXT, bold=True, align=PP_ALIGN.CENTER,
          ancora=MSO_ANCHOR.MIDDLE)

# ══════════════════════════════════════════════════════════════
# 5 · Cenário analisado
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", 5, "Etapa 3 · Cenário analisado", "Cenário analisado",
               "Um Panamax atracado por boreste no terminal T-Grão, no canal do Porto de Santos",
               "Fonte: relatório do TF I (seções 1.2, 1.4, 1.8 e 7) · DHN (2025) · notebook TF4.0")
ch = 3.55
for k, (rot, kpi, leg, corpo) in enumerate([
        ("Terminal T-Grão", "149°", "orientação do píer (329° na entrada do software)",
         ["•  Atracação por boreste",
          "•  Cabeços e defensas a cada 15 m (45,08 m em 3 vãos; a PIANC indica 12 a 25 m)",
          "•  Profundidade de cerca de 16,5 m"]),
        ("Navio Panamax", "225 m", "comprimento total (LOA)",
         ["•  Boca de 32,26 m e pontal de 18,5 m",
          "•  Calado analisado de 9 m, abaixo do carregado (12 a 14,3 m)",
          "•  12 cabos, 6 em cada extremidade, como no arranjo típico do MEG4"]),
        ("Ambiente", "−0,16 a 1,52 m", "faixa da maré em Santos em 2025 (DHN)",
         ["•  Águas abrigadas: as ondas chegam atenuadas e vento e corrente de maré dominam",
          "•  Navios passantes também carregam os cabos, mas não entram nas 394 condições"])]):
    card_texto(sl, X0 + k * (cw + 0.15), Y_CONT, cw, ch, rot, corpo, cor, kpi=kpi, kpi_legenda=leg,
               size=11.5)
ny = Y_CONT + ch + 0.15
card(sl, X0, ny, LARG, Y_FIM - ny, fill=tint(cor, 0.08))
forma(sl, MSO_SHAPE.RECTANGLE, X0, ny + 0.2, 0.07, Y_FIM - ny - 0.4, cor)
texto(sl, X0 + 0.3, ny, LARG - 0.5, Y_FIM - ny,
      "**Caso mais desfavorável para a inclinação dos cabos:** maré alta com calado reduzido, a maior "
      "altura do navio em relação ao cais (Tabela 4 do TF I). Por isso o caso de "
      "maré de 1,58 m e calado de 9 m é o cenário crítico da geometria.", 12.5, TXT2, ancora=MSO_ANCHOR.MIDDLE)

# ══════════════════════════════════════════════════════════════
# 6 · Arranjo de amarração
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", 6, "Etapa 3 · Cenário analisado", "Arranjo de amarração",
               "12 linhas distribuídas em 5 cabeços, 6 em cada extremidade do navio",
               "Fonte: TF4.0 – Colab · esquema com as posições longitudinais de buzinas e cabeços do notebook; "
               "buzinas desenhadas no costado de atracação")
por_imagem(sl, "00_plano_amarracao.png", X0, Y_CONT, LARG, 2.4, pad=0.1)
cy = 4.55
card_texto(sl, X0, cy, cw, Y_FIM - cy, "À vante · 6 cabos",
           ["Cabos **3 a 8**, em pares nos cabeços **A, B e G**.",
            "Mesma quantidade do arranjo típico do MEG4 para Panamax (6 cabos por extremidade)."],
           cor, size=11.5)
card_texto(sl, X0 + cw + 0.15, cy, cw, Y_FIM - cy, "À ré · 6 cabos",
           ["Cabos **1, 2, 11 e 12** no cabeço **S**.",
            "Cabos **9 e 10** no cabeço **N**."], cor, size=11.5)
card_texto(sl, X0 + 2 * (cw + 0.15), cy, cw, Y_FIM - cy, "Por que isso importa",
           ["Cada cabeço soma a carga das linhas presas nele: o **S recebe 4 das 12**.",
            "Linhas no mesmo cabeço com comprimentos diferentes tendem a dividir a carga de forma "
            "desigual, como os cabos 9 e 10."], cor, size=11.5)

# ══════════════════════════════════════════════════════════════
# 7 · Como foi abordado
# ══════════════════════════════════════════════════════════════
sl, cor = base("abordagem", 7, "Etapa 4 · Abordagem", "Como foi abordado",
               "Da caracterização do berço à leitura dos resultados",
               "Base teórica do TF I: forças de vento e corrente pelo MEG4 (F = ½·C·ρ·V²·A) e "
               "decomposição da tração nos cabos (T·cos θ e T·sen θ)")
sw, sh, sgap = (LARG - 3 * 0.45) / 4, 2.75, 0.45
passos = [("Caracterização",
           "Orientação do píer, espaçamento de cabeços e defensas, maré da DHN e dados do "
           "Panamax, levantados no TF I."),
          ("Simulação",
           "394 combinações de vento, corrente, onda (Hs, Tp), maré e direções, com calado de "
           "9 m. O software exporta esforços, picos e condições em três CSV."),
          ("Pós-processamento",
           "Notebook em Python (TF4.0): junta os arquivos, calcula estatísticas, filtra riscos e "
           "monta o envelope crítico e a matriz de operabilidade."),
          ("Geometria",
           "Comprimento 3D e ângulo vertical de cada linha, a partir das coordenadas de buzinas e "
           "cabeços, com calado de 9 e 14 m e maré de 0 e 1,58 m.")]
for k, (rot, corpo) in enumerate(passos):
    xx = X0 + k * (sw + sgap)
    card(sl, xx, Y_CONT, sw, sh)
    texto(sl, xx + 0.2, Y_CONT + 0.15, 0.8, 0.5, str(k + 1), 24, cor, fonte="Georgia", bold=True)
    rotulo(sl, xx + 0.2, Y_CONT + 0.72, sw - 0.4, rot, cor)
    texto(sl, xx + 0.2, Y_CONT + 1.05, sw - 0.4, sh - 1.15, corpo, 10.5, TXT2)
    if k < 3:
        seta(sl, xx + sw + 0.12, Y_CONT + sh / 2 - 0.2, 0.21, 0.4, cor)
ky = Y_CONT + sh + 0.2
card(sl, X0, ky, LARG, Y_FIM - ky)
rotulo(sl, X0 + 0.25, ky + 0.2, 6, "Critérios de avaliação", cor)
kw = (LARG - 0.5) / 3
for k, (v, leg) in enumerate([("55% MBL", "tração máxima nos cabos (OCIMF); MBL é a carga mínima de ruptura"),
                              ("100% SWL", "força máxima nos cabeços; SWL de 75 t, pela razão força/SWL do notebook"),
                              ("≤ 25°", "ângulo vertical das linhas em relação ao cais (MEG4)")]):
    xx = X0 + 0.25 + k * kw
    texto(sl, xx, ky + 0.55, kw - 0.3, 0.5, v, 22, cor, fonte="Georgia", bold=True)
    texto(sl, xx, ky + 1.1, kw - 0.3, 0.6, leg, 10, TXT3)

# ══════════════════════════════════════════════════════════════
# 8 · Resumo estatístico (tabela)
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", 8, ROT_RES, "Resumo estatístico geral",
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
              "31,5% MBL", "mediana da tração nos cabos", rot="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 9 · Histogramas
# ══════════════════════════════════════════════════════════════
slide_grafico(9, "Distribuição dos esforços",
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
# 10 · Influência da maré
# ══════════════════════════════════════════════════════════════
slide_grafico(10, "Influência da maré",
              "Pico de cada elemento e o nível de maré em que ele acontece",
              "02_influencia_mare.png",
              ["Mostra o maior esforço de cada cabo (círculos) e de cada cabeço (triângulos). "
               "A cor indica a maré no instante desse pico.",
               "**Intuito:** apontar quais elementos chegam ao limite e se isso ocorre com "
               "maré alta ou baixa."],
              ["**Só o cabo 9 passa de 55% MBL**; os outros 11 ficam abaixo de 40%. O pico "
               "dele ocorre com maré perto de 1,5 m, quando as linhas ficam mais inclinadas.",
               "Nos cabeços, só o S passa de 100% SWL, e com **maré baixa** (cerca de 0,3 m). "
               "A maré não age no mesmo sentido em todos os elementos."],
              "75,3% MBL", "pico do cabo 9 (ligado ao cabeço N)")

# ══════════════════════════════════════════════════════════════
# 11 · Correlação + boxplot
# ══════════════════════════════════════════════════════════════
slide_grafico(11, "Clima ou arranjo de cabos?",
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
# 12 · Combinações críticas (tabelas)
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", 12, ROT_RES, "Combinações críticas",
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
              "5 de 5", "piores condições têm o cabo 9 como crítico", rot="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 13 · Envelope crítico
# ══════════════════════════════════════════════════════════════
slide_grafico(13, "Envelope crítico: vento × corrente",
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
# 14 · Matriz de operabilidade
# ══════════════════════════════════════════════════════════════
slide_grafico(14, "Matriz de operabilidade",
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
# 15–16 · Geometria das linhas
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
assert max(g[4] for g in GEO) == 21.61 < 25


def legenda_geo(sl, y):
    for k, (fill, txt) in enumerate([(ALERTA_FILL, "cabo 9: pico de 75,3% MBL"),
                                     (DESTAQUE_FILL, "cabo 5: maior ângulo vertical")]):
        xx = X0 + 0.25 + k * 3.55
        forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, y, 0.28, 0.2, fill, BORDA, raio=0.2)
        texto(sl, xx + 0.37, y - 0.02, 3.1, 0.25, txt, 9.5, TXT3)


sl, cor = base("resultados", 15, ROT_RES, "Geometria das linhas · maré 1,58 m",
               "Calado 9 m e maré alta: a maior altura do navio em relação ao cais",
               "Fonte: TF4.0 – Colab · calcular_geometria_amarracao(calado 9 m, maré 1,58 m) · "
               "1,58 m é a maré alta do notebook (a máxima DHN 2025 é 1,52 m)")
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
               "no caso de maior inclinação dos cabos.",
               "**Intuito:** conferir o limite de 25° do MEG4 e ligar o esforço de cada cabo à "
               "sua geometria."],
              ["Pelo cálculo do notebook, **todas as linhas ficam abaixo de 25°**; a mais inclinada é o cabo 5.",
               "O cabo 9 é **9,4 m mais curto** que o cabo 10, que divide o cabeço N. Mais "
               "rígido, tende a absorver mais carga, coerente com os picos."],
              "21,6° < 25°", "maior ângulo vertical (cabo 5) e limite do MEG4",
              rot="O que a tabela lê")

sl, cor = base("resultados", 16, ROT_RES, "Geometria das linhas · maré 0 m",
               "Mesmo calado com maré baixa: efeito isolado da maré",
               "Fonte: TF4.0 – Colab · calcular_geometria_amarracao(calado 9 m, maré 0 m) · "
               "Δ calculado contra a tabela de maré 1,58 m · calado 14 m pela mesma função")
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
              ["Com maré baixa a buzina desce 1,58 m e **todas as linhas ficam menos "
               "inclinadas**. O comprimento muda só 0,2 a 0,5 m.",
               "Com calado de 14 m o ângulo cai para 0,1° a 5,2°. A maré muda a inclinação, "
               "mas não a diferença entre os cabos 9 e 10."],
              "−2,2° a −4,9°", "queda do ângulo vertical com a maré baixa",
              rot="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 17 · Conclusões parciais
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", 17, "Etapa 6 · Conclusões", "Conclusões parciais",
               "O que os resultados mostram em conjunto",
               "Fonte: TF4.0 – Colab e relatório do TF I")
achados = [
    ("C1", "O cabo 9 concentra o risco",
     "Único cabo acima de 55% MBL (pico de 75,3%) e crítico nas 5 piores condições."),
    ("C2", "O cabeço S passa do SWL",
     "2 condições acima de 100% (até 110,8%, cerca de 83 t para SWL de 75 t). Recebe 4 das 12 linhas."),
    ("C3", "O ambiente sozinho não explica",
     "Correlações fracas e 19 de 20 faixas de vento × corrente com pior caso acima do limite."),
    ("C4", "A maré muda a geometria",
     "Maré alta inclina as linhas 2° a 5° a mais. Pelo cálculo atual, todas ficam abaixo dos 25° do MEG4."),
]
gw, gh = (LARG - 0.45) / 4, 2.9
for k, (cid, tit, corpo) in enumerate(achados):
    xx = X0 + k * (gw + 0.15)
    card(sl, xx, Y_CONT, gw, gh)
    texto(sl, xx + 0.2, Y_CONT + 0.2, gw - 0.4, 0.25, cid, 10, TXT3, bold=True)
    texto(sl, xx + 0.2, Y_CONT + 0.48, gw - 0.4, 0.65, tit, 14, cor, bold=True)
    texto(sl, xx + 0.2, Y_CONT + 1.2, gw - 0.4, gh - 1.3, corpo, 11.5, TXT2)
py = Y_CONT + gh + 0.2
card(sl, X0, py, LARG, Y_FIM - py, fill=tint(cor, 0.08))
forma(sl, MSO_SHAPE.RECTANGLE, X0, py + 0.25, 0.07, Y_FIM - py - 0.5, cor)
rotulo(sl, X0 + 0.3, py + 0.22, 6, "Resposta ao objetivo", cor)
texto(sl, X0 + 0.3, py + 0.55, LARG - 0.6, Y_FIM - py - 0.65,
      "A simulação mostrou em quais condições e em quais elementos o berço chega ao limite: o "
      "cabo 9, no cabeço N, e o cabeço S. São esses os pontos que o monitoramento por visão "
      "computacional deve acompanhar primeiro.", 12.5, TXT2)

# ══════════════════════════════════════════════════════════════
# 18 · Próximos passos
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", 18, "Etapa 6 · Conclusões", "Próximos passos",
               "Do diagnóstico parcial à entrega final",
               "Próximos passos propostos a partir dos resultados desta etapa e do plano do relatório do TF I")
etapas = [("TF I · concluído",
           "Base teórica: equipamentos, defensas, cabos PPL/PE, forças ambientais (MEG4) e "
           "caracterização do T-Grão."),
          ("Etapa atual",
           "Simulação de 394 condições com calado de 9 m e leitura dos esforços, dos riscos e "
           "da geometria."),
          ("Próxima etapa",
           "Ampliar os cenários simulados e desenvolver o modelo de monitoramento.")]
ew, eh = (LARG - 2 * 0.5) / 3, 1.9
for k, (rot, corpo) in enumerate(etapas):
    xx = X0 + k * (ew + 0.5)
    card(sl, xx, Y_CONT, ew, eh, fill=tint(cor, 0.12) if k == 1 else CARD)
    rotulo(sl, xx + 0.2, Y_CONT + 0.18, ew - 0.4, rot, cor)
    texto(sl, xx + 0.2, Y_CONT + 0.5, ew - 0.4, eh - 0.6, corpo, 11, TXT2)
    if k < 2:
        seta(sl, xx + ew + 0.14, Y_CONT + eh / 2 - 0.2, 0.22, 0.4, cor)
ly = Y_CONT + eh + 0.2
card(sl, X0, ly, LARG, Y_FIM - ly)
rotulo(sl, X0 + 0.25, ly + 0.2, 6, "O que vem a seguir", cor)
passos_a = ["•  Simular o calado carregado (14 m) e a maré mínima de 2025 (−0,16 m)",
            "•  Incluir Hs, direções e maré na matriz de operabilidade",
            "•  Avaliar o efeito de navios passantes no canal"]
passos_b = ["•  Revisar o cálculo de geometria (lado das buzinas e posição dos cabeços) e validar com o berço real",
            "•  Testar ajustes no arranjo: comprimento ou pré-tensão do cabo 9 e menos linhas no cabeço S",
            "•  Desenvolver o monitoramento dos cabos por visão computacional e machine learning"]
bw = (LARG - 0.75) / 2
texto(sl, X0 + 0.25, ly + 0.55, bw, Y_FIM - ly - 0.65, passos_a, 11.5, TXT2, entre=7)
texto(sl, X0 + 0.5 + bw, ly + 0.55, bw, Y_FIM - ly - 0.65, passos_b, 11.5, TXT2, entre=7)

# ══════════════════════════════════════════════════════════════
# 19 · Encerramento
# ══════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(prs.slide_layouts[6])
capa_fundo(sl)
texto(sl, X0, 2.3, 11, 1.0, "Obrigado pela atenção", 44, TXT, fonte="Georgia", bold=True)
texto(sl, X0, 3.7, 11, 1.0, "Perguntas?", 44, TXT, fonte="Georgia", bold=True)
texto(sl, X0 + 0.05, 5.0, 8, 0.35, "Gustavo Guilherme Silva · NUSP 11820555", 13.5, TXT2)
numero(sl, TOTAL)

assert len(prs.slides) == TOTAL, len(prs.slides)
prs.save(SAIDA)
print("Salvo:", SAIDA, "·", len(prs.slides), "slides")
