"""
Apresentação parcial do TF: sistema de amarração em conjuntos portuários.
Junta o relatório do TF I (motivação, justificativa, cenário, abordagem) com os
resultados do notebook TF4.0 (calado 9 m, 394 condições).

Estilo baseado no PDF do PINN: fundo creme, barra lateral colorida por etapa,
títulos serifados, cards claros e formas decorativas no canto. Toda a tipografia
usa Times New Roman.

Fluxo de leitura em cada slide:
  rótulo da etapa → título (tema) → subtítulo (mensagem principal) → evidência → cards 1 e 2
A barra de progresso no topo mostra em qual etapa a apresentação está.

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
FONTE = "Times New Roman"


# ── Paleta ────────────────────────────────────────────────────
def hx(s):
    return RGBColor.from_string(s.lstrip("#"))


CREME, CARD, BORDA = "#F4F1EA", "#FBFAF7", "#DDD6C8"
TXT, TXT2, TXT3 = "#1C2430", "#3F4757", "#6F7684"
SECAO = {                      # uma cor por etapa, como no PDF de referência
    "capa":          "#F2B57E",
    "roteiro":       "#B08A2E",
    "motivacao":     "#DB7F3B",
    "justificativa": "#B08A2E",
    "cenario":       "#2F8F86",
    "abordagem":     "#4C78C8",
    "resultados":    "#7A63C2",
    "conclusoes":    "#C8605A",
    "referencias":   "#6F7684",
}
ETAPAS = [("motivacao", "Motivação"), ("justificativa", "Justificativa"), ("cenario", "Cenário"),
          ("abordagem", "Abordagem"), ("resultados", "Resultados"), ("conclusoes", "Conclusões")]
ALERTA_FILL, DESTAQUE_FILL = "#F8E1DB", "#E9E3F6"


def tint(cor, f):
    a, b = hx(CREME), hx(cor)
    return RGBColor(*[round(a[i] * (1 - f) + b[i] * f) for i in range(3)])


W, H = 13.333, 7.5
X0, Y_CONT, Y_FIM = 0.75, 2.05, 6.95
LARG = 11.85                   # largura útil do conteúdo
TOTAL = 22
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


def runs(par, texto, size, cor, bold=False, italic=False, spc=None):
    """Escreve texto num parágrafo; trechos entre **...** ficam em negrito."""
    for i, pedaco in enumerate(re.split(r"\*\*(.+?)\*\*", texto)):
        if not pedaco:
            continue
        r = par.add_run()
        r.text = pedaco
        f = r.font
        f.size, f.name, f.bold, f.italic = Pt(size), FONTE, bold or (i % 2 == 1), italic
        f.color.rgb = hx(cor)
        if spc:
            r._r.get_or_add_rPr().set("spc", str(spc))


def texto(sl, x, y, w, h, paragrafos, size=12, cor=TXT2, bold=False, italic=False,
          align=PP_ALIGN.LEFT, spc=None, entre=5, ancora=MSO_ANCHOR.TOP, linha=1.08):
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
        runs(p, ptxt, size, cor, bold, italic, spc)
    return tb


def card(sl, x, y, w, h, fill=CARD):
    return forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, BORDA, raio=0.045)


def rotulo(sl, x, y, w, txt, cor, size=10):
    texto(sl, x, y, w, 0.25, txt.upper(), size, cor, bold=True, spc=120)


def selo(sl, x, y, n, cor, d=0.27):
    """Círculo numerado que indica a ordem de leitura."""
    forma(sl, MSO_SHAPE.OVAL, x, y, d, d, cor)
    texto(sl, x, y, d, d, str(n), 11, "#FFFFFF", bold=True, align=PP_ALIGN.CENTER,
          ancora=MSO_ANCHOR.MIDDLE, linha=1.0)


def seta(sl, x, y, w, h, cor):
    forma(sl, MSO_SHAPE.CHEVRON, x, y, w, h, tint(cor, 0.55))


def numero(sl, n):
    texto(sl, 11.6, 7.07, 1.4, 0.25, f"{n} / {TOTAL}", 11, TXT3, align=PP_ALIGN.RIGHT)


def trilha(sl, atual):
    """Barra de progresso com as seis etapas; a etapa atual fica destacada."""
    cw, gap, h = 0.92, 0.05, 0.27
    x = X0 + LARG - (len(ETAPAS) * cw + (len(ETAPAS) - 1) * gap)
    idx = [k for k, _ in ETAPAS].index(atual)
    for i, (k, nome) in enumerate(ETAPAS):
        xx = x + i * (cw + gap)
        if i == idx:
            forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, 0.4, cw, h, SECAO[k], raio=0.5)
            cor_t, b = "#FFFFFF", True
        else:
            forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, 0.4, cw, h,
                  tint(SECAO[k], 0.22 if i < idx else 0.07), raio=0.5)
            cor_t, b = (TXT2 if i < idx else TXT3), False
        texto(sl, xx, 0.4, cw, h, nome, 9.5, cor_t, bold=b, align=PP_ALIGN.CENTER,
              ancora=MSO_ANCHOR.MIDDLE, linha=1.0)


def base(secao, n, rot=None, titulo=None, mensagem=None, nota=None, etapa=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    cor = SECAO[secao]
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
    # formas decorativas no canto inferior direito (como no PDF de referência)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 10.25, 4.45, 3.6, 3.6, tint(cor, 0.10), raio=0.3)
    forma(sl, MSO_SHAPE.OVAL, 11.3, 5.6, 3.8, 3.8, tint(cor, 0.20))
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, cor)
    if etapa:
        trilha(sl, etapa)
    if rot:
        texto(sl, X0, 0.45, 6.3, 0.3, rot.upper(), 10.5, cor, bold=True, spc=200)
    if titulo:
        texto(sl, X0, 0.8, LARG, 0.7, titulo, 32, TXT)
    if mensagem:
        texto(sl, X0, 1.47, LARG, 0.45, mensagem, 15, TXT2, italic=True)
    if nota:
        texto(sl, X0, 7.06, 10.7, 0.3, nota, 9.5, TXT3)
    numero(sl, n)
    return sl, cor


def card_texto(sl, x, y, w, h, rot, corpo, cor, kpi=None, kpi_legenda=None, size=12, num=None):
    card(sl, x, y, w, h)
    pad = 0.2
    xl = x + pad
    if num is not None:
        selo(sl, x + pad, y + 0.15, num, cor)
        xl = x + pad + 0.37
    rotulo(sl, xl, y + 0.19, w - pad - (xl - x), rot, cor)
    y_corpo = y + 0.55
    if kpi:
        texto(sl, x + pad, y + 0.5, w - 2 * pad, 0.5, kpi, 26, cor, bold=True)
        texto(sl, x + pad, y + 1.0, w - 2 * pad, 0.3, kpi_legenda, 11, TXT3)
        y_corpo = y + 1.36
    texto(sl, x + pad, y_corpo, w - 2 * pad, y + h - y_corpo - 0.12, corpo, size, TXT2)


def kpi_tile(sl, x, y, w, h, valor, legenda, cor):
    card(sl, x, y, w, h)
    texto(sl, x + 0.2, y + 0.14, w - 0.4, 0.5, valor, 27, cor, bold=True)
    texto(sl, x + 0.2, y + 0.7, w - 0.4, h - 0.8, legenda, 11.5, TXT3)


def faixa(sl, y, h, txt, cor, size=13.5):
    """Faixa de síntese no rodapé do conteúdo."""
    card(sl, X0, y, LARG, h, fill=tint(cor, 0.09))
    forma(sl, MSO_SHAPE.RECTANGLE, X0, y + 0.18, 0.07, h - 0.36, cor)
    texto(sl, X0 + 0.32, y, LARG - 0.55, h, txt, size, TXT2, ancora=MSO_ANCHOR.MIDDLE)


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


def tabela(sl, x, y, w, cabec, linhas, larguras, alt_linha=0.34, size=11.5,
           cor_cabec=TXT, destaques=None, alinhar=None):
    """destaques: {(linha, coluna|None): cor_fundo}; coluna None = linha inteira.
    alinhar: lista de PP_ALIGN por coluna (padrão: centralizado)."""
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
            c.margin_left = c.margin_right = Inches(0.06)
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
            p.alignment = (alinhar[j] if alinhar and i else PP_ALIGN.CENTER)
            runs(p, str(val), size if i else size - 0.5, cor_t, bold=b)
    return shp


NOTEBOOK = "SILVA (2026b), notebook TF4.0"


def nota_imagem(original):
    t = f"Fonte: {NOTEBOOK} (calado 9 m, 394 condições)"
    if not original:
        t += " · gráfico redesenhado a partir das saídas do notebook"
    return t


# ── Layout dos slides de gráfico: imagem à esquerda, cards 1 e 2 à direita ──
IMG = dict(x=X0, y=Y_CONT, w=7.05, h=Y_FIM - Y_CONT)
COL = dict(x=8.0, w=4.6)
C1_H, GAP = 2.05, 0.15


def colunas_texto(sl, cor, le, resultado, kpi, kpi_leg, rot1="O que o gráfico lê",
                  rot2="Resultado obtido"):
    card_texto(sl, COL["x"], Y_CONT, COL["w"], C1_H, rot1, le, cor, num=1)
    y2 = Y_CONT + C1_H + GAP
    card_texto(sl, COL["x"], y2, COL["w"], Y_FIM - y2, rot2, resultado, cor,
               kpi=kpi, kpi_legenda=kpi_leg, num=2)


def slide_grafico(n, sub, titulo, mensagem, nome_img, le, resultado, kpi, kpi_leg):
    sl, cor = base("resultados", n, f"Etapa 5 · Resultados · {sub}", titulo, mensagem,
                   etapa="resultados")
    original = por_imagem(sl, nome_img, **IMG)
    colunas_texto(sl, cor, le, resultado, kpi, kpi_leg)
    texto(sl, X0, 7.06, 10.7, 0.3, nota_imagem(original), 9.5, TXT3)


def capa_fundo(sl):
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 9.7, 3.9, 4.6, 4.6, tint(SECAO["capa"], 0.18), raio=0.25)
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, SECAO["capa"])


cw = (LARG - 0.3) / 3          # largura de card em grade de 3 colunas

# ══════════════════════════════════════════════════════════════
# 1 · Capa
# ══════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(prs.slide_layouts[6])
capa_fundo(sl)
texto(sl, X0 + 0.05, 1.65, 11, 0.3, "PNV3510 · TRABALHO DE FORMATURA · APRESENTAÇÃO PARCIAL", 12, TXT,
      bold=True, spc=250)
texto(sl, X0, 2.1, 11.5, 2.0, ["Sistema de amarração em", "conjuntos portuários"], 50, TXT,
      bold=True, entre=0, linha=0.95)
texto(sl, X0 + 0.05, 4.3, 11, 0.45,
      "Análise dos esforços em um navio Panamax atracado no terminal T-Grão, Porto de Santos",
      18, TXT, bold=True)
texto(sl, X0 + 0.05, 4.95, 8, 0.7,
      ["Gustavo Guilherme Silva · NUSP 11820555",
       "Escola Politécnica da Universidade de São Paulo"], 15, TXT2, entre=2)
numero(sl, 1)

# ══════════════════════════════════════════════════════════════
# 2 · Roteiro
# ══════════════════════════════════════════════════════════════
sl, cor = base("roteiro", 2, "Apresentação", "Roteiro",
               "Do problema aos resultados parciais, em seis etapas")
itens = [("motivacao", "Motivação", "Por que a amarração é crítica nos terminais graneleiros de Santos"),
         ("justificativa", "Justificativa", "A lacuna no monitoramento dos cabos e o objetivo desta etapa"),
         ("cenario", "Cenário analisado", "Terminal T-Grão, maré de Santos, layout típico do Panamax e arranjo simulado"),
         ("abordagem", "Abordagem", "Etapas do trabalho e critérios de avaliação"),
         ("resultados", "Resultados", "Leitura dos gráficos da simulação: esforços, riscos e geometria"),
         ("conclusoes", "Conclusões", "Conclusões parciais e próximos passos")]
gh = (Y_FIM - Y_CONT - 0.15) / 2
for k, (sec, tit, desc) in enumerate(itens):
    xx = X0 + (k % 3) * (cw + 0.15)
    yy = Y_CONT + (k // 3) * (gh + 0.15)
    c = SECAO[sec]
    card(sl, xx, yy, cw, gh)
    forma(sl, MSO_SHAPE.RECTANGLE, xx, yy + 0.3, 0.07, gh - 0.6, c)
    texto(sl, xx + 0.3, yy + 0.22, 1.2, 0.6, f"{k + 1:02d}", 32, c, bold=True)
    texto(sl, xx + 0.3, yy + 0.92, cw - 0.5, 0.4, tit, 17, TXT, bold=True)
    texto(sl, xx + 0.3, yy + 1.38, cw - 0.5, gh - 1.48, desc, 13, TXT2)

# ══════════════════════════════════════════════════════════════
# 3 · Motivação
# ══════════════════════════════════════════════════════════════
sl, cor = base("motivacao", 3, "Etapa 1 · Motivação", "Por que estudar a amarração?",
               "Durante o carregamento, o navio é mantido no berço apenas pelo sistema de amarração",
               "Fontes: SILVA (2026a, Tabela 1 e seções 1.3, 1.9 e 6.3); UK P&I CLUB (2009)",
               etapa="motivacao")
tw_ = (LARG - 0.45) / 4
for k, (v, leg) in enumerate([(">100 Mt", "de carga por ano no Porto de Santos, o maior da América Latina"),
                              ("14 de 60", "berços de Santos são dedicados a granéis vegetais"),
                              ("US$ 34 mi", "em sinistros com equipamentos de amarração em 20 anos"),
                              ("até 40%", "de aumento da tração nos cabos com a passagem de outro navio")]):
    kpi_tile(sl, X0 + k * (tw_ + 0.15), Y_CONT, tw_, 1.4, v, leg, cor)
cy, mh = Y_CONT + 1.55, 2.15
for k, (rot, corpo) in enumerate([
        ("Segurança",
         ["A ruptura de um cabo tracionado libera de uma só vez a energia elástica acumulada (**snap-back**).",
          "Por isso, os amarradores devem ficar a pelo menos 7 m dos cabeços."]),
        ("Operação",
         ["O navio atracado sofre a ação contínua do vento, da corrente de maré e de navios passantes.",
          "Se ele se afasta do cais, shiploaders e esteiras podem ser danificados e o embarque é interrompido."]),
        ("Meio ambiente",
         ["O despejo de grãos no estuário eleva a demanda bioquímica de oxigênio (DBO) da água.",
          "A limpeza exige dragagem e compromete as metas de **Porto Verde** (ESG)."])]):
    card_texto(sl, X0 + k * (cw + 0.15), cy, cw, mh, rot, corpo, cor)
sy = cy + mh + 0.15
faixa(sl, sy, Y_FIM - sy,
      "**Em síntese:** uma falha de amarração atinge, ao mesmo tempo, pessoas, operação e meio ambiente. "
      "Conhecer os esforços nos cabos é o primeiro passo para preveni-la.", cor)

# ══════════════════════════════════════════════════════════════
# 4 · Justificativa e objetivo
# ══════════════════════════════════════════════════════════════
sl, cor = base("justificativa", 4, "Etapa 2 · Justificativa", "Justificativa e objetivo",
               "A tração nos cabos ainda não é medida de forma contínua nos terminais",
               "Fonte: SILVA (2026a, Resumo e seção 1)", etapa="justificativa")
fw, fh, fgap = (LARG - 2 * 0.5) / 3, 2.15, 0.5
fluxo = [("O problema",
          "A tração nos cabos varia continuamente com vento, corrente, maré e navios passantes, "
          "mas raramente é medida durante a operação."),
         ("A lacuna",
          "Células de carga e extensômetros se degradam rapidamente com a salinidade e a umidade. "
          "O custo de manutenção restringe o acompanhamento a inspeções visuais."),
         ("A proposta do TF",
          "Monitoramento de saúde estrutural (**SHM**) sem contato, por visão computacional e "
          "machine learning, integrado a simulações dos cenários críticos.")]
for k, (rot, corpo) in enumerate(fluxo):
    xx = X0 + k * (fw + fgap)
    card_texto(sl, xx, Y_CONT, fw, fh, rot, corpo, cor, num=k + 1, size=13.5)
    if k < 2:
        seta(sl, xx + fw + 0.14, Y_CONT + fh / 2 - 0.2, 0.22, 0.4, cor)
oy = Y_CONT + fh + 0.2
card(sl, X0, oy, LARG, Y_FIM - oy)
rotulo(sl, X0 + 0.25, oy + 0.2, 6, "Objetivo desta etapa", cor)
texto(sl, X0 + 0.25, oy + 0.52, LARG - 0.5, 0.8,
      "Identificar, por simulação, as condições de vento, corrente, onda e maré que levam os cabos e "
      "os cabeços do berço ao limite. Esses cenários formam a base quantitativa que o monitoramento "
      "deverá detectar.", 13.5, TXT2)
chips = ["Quantificar os esforços nas 394 condições",
         "Identificar condições e elementos críticos",
         "Relacionar a carga à geometria dos cabos"]
chw = (LARG - 0.5 - 0.3) / 3
for k, c_txt in enumerate(chips):
    xx = X0 + 0.25 + k * (chw + 0.15)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, Y_FIM - 0.72, chw, 0.48, tint(cor, 0.16), raio=0.3)
    texto(sl, xx, Y_FIM - 0.72, chw, 0.48, c_txt, 12, TXT, bold=True, align=PP_ALIGN.CENTER,
          ancora=MSO_ANCHOR.MIDDLE)

# ══════════════════════════════════════════════════════════════
# 5 · Cenário analisado
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", 5, "Etapa 3 · Cenário analisado", "Cenário analisado",
               "Navio Panamax atracado por boreste no terminal T-Grão, no canal do Porto de Santos",
               "Fontes: SILVA (2026a, seções 1.4, 1.7, 1.8 e 7), com dados de ACP (2026), PIANC (2017) e "
               "T-GRÃO (2026); PIANC (2025); MARINHA DO BRASIL (2025); SILVA (2026b)", etapa="cenario")
ch = 3.3
for k, (rot, kpi, leg, corpo) in enumerate([
        ("Terminal T-Grão", "149°", "orientação do píer (329° na entrada do software)",
         ["•  Atracação por boreste",
          "•  Cabeços e defensas a cada 15 m (45,08 m em três vãos; a PIANC indica 12 a 25 m)",
          "•  Profundidade de cerca de 16,5 m"]),
        ("Navio Panamax", "225 m", "comprimento total (LOA)",
         ["•  Boca de 32,26 m e pontal de 18,5 m",
          "•  Calado analisado de 9 m, abaixo do calado carregado (12 a 14,3 m)",
          "•  12 cabos, 6 em cada extremidade"]),
        ("Ambiente", "−0,16 a 1,52 m", "faixa de maré prevista para Santos em 2025",
         ["•  Águas abrigadas: as ondas chegam atenuadas; vento e corrente de maré predominam",
          "•  Navios passantes também solicitam os cabos, mas não fazem parte das 394 condições"])]):
    card_texto(sl, X0 + k * (cw + 0.15), Y_CONT, cw, ch, rot, corpo, cor, kpi=kpi, kpi_legenda=leg,
               size=13)
ny = Y_CONT + ch + 0.15
faixa(sl, ny, Y_FIM - ny,
      "**Cenário mais desfavorável para a inclinação dos cabos:** maré alta com calado reduzido, que "
      "coloca o navio na maior altura em relação ao cais (SILVA, 2026a, Tabela 4). Por isso, a geometria "
      "crítica foi calculada com maré de 1,58 m e calado de 9 m.", cor, size=13)

# ══════════════════════════════════════════════════════════════
# 6 · Estudo de maré
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", 6, "Etapa 3 · Cenário analisado", "Estudo de maré em Santos",
               "A maré sobe até 1,68 m em menos de 6 horas e muda a altura do navio em relação ao cais",
               etapa="cenario")
por_imagem(sl, "08_mare_dhn.png", **IMG)
colunas_texto(sl, cor,
              ["Previsões da tábua de marés da DHN para Santos em 2025 (1.453 valores). Cada barra vai "
               "da mínima à máxima do mês; o ponto indica a média.",
               "**Intuito:** definir a faixa de maré que o sistema de amarração precisa acomodar."],
              ["A amplitude mensal varia de 1,36 m (junho) a 1,68 m (fevereiro); a média se mantém "
               "entre 0,77 e 0,78 m.",
               "Os cabos precisam ser ajustados ao longo da operação (OCIMF, 2018). A geometria usou "
               "1,58 m como maré alta, pouco acima da máxima de 2025."],
              "1,68 m", "subida entre 27/fev (21h19) e 28/fev (03h14)", rot2="O que isso significa")
texto(sl, X0, 7.06, 10.7, 0.3,
      "Fontes: MARINHA DO BRASIL (2025), compilado em SILVA (2026a, Tabelas 6 e 7); OCIMF (2018, seção 1.3) "
      "· gráfico elaborado com os valores da Tabela 7", 9.5, TXT3)

# ══════════════════════════════════════════════════════════════
# 7 · Layout de amarração de um Panamax (referência)
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", 7, "Etapa 3 · Cenário analisado", "Layout de amarração de um Panamax",
               "Cada grupo de cabos resiste principalmente a uma direção de esforço",
               "Fontes: SILVA (2026a, Figura 9, Tabela 5 e cap. 5), com base em OCIMF (2018); "
               "OCIMF (2018, seções 1.3 e 1.4); SOUTHERN PORTS (2026)", etapa="cenario")
dh = 2.2
por_imagem(sl, "07_layout_panamax.png", X0, Y_CONT, LARG, dh, pad=0.08)
by = Y_CONT + dh + 0.15
bh = Y_FIM - by
tw7 = 6.95
card(sl, X0, by, tw7, bh)
selo(sl, X0 + 0.2, by + 0.15, 1, cor)
rotulo(sl, X0 + 0.57, by + 0.19, 5, "Função de cada grupo", cor)
tabela(sl, X0 + 0.2, by + 0.55, tw7 - 0.4,
       ["Grupo", "Cabos", "Direção de referência", "Esforço principal"],
       [["Lançantes", "4 + 4", "oblíquos, a vante e a ré", "afastamento das extremidades; avanço e recuo"],
        ["Través", "1 + 1", "até 15° da perpendicular ao navio", "afastamento (sway) e guinada (yaw)"],
        ["Espringues", "1 + 1", "até 10° da direção do costado", "avanço e recuo (surge)"]],
       [1.05, 0.65, 1.75, 2.35], alt_linha=0.47, size=11, cor_cabec=cor,
       alinhar=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT])
rx = X0 + tw7 + 0.15
card_texto(sl, rx, by, LARG - tw7 - 0.15, bh, "Por que este arranjo",
           ["•  O esforço transversal é o maior: o vento de través gera força muito maior que o de proa.",
            "•  Distribuir os cabos entre proa e popa forma um arranjo equilibrado, que também resiste à guinada.",
            "•  As escotilhas ocupam o convés central; guinchos e cabeços ficam na proa e na popa.",
            "•  O berço pode pedir outra combinação com 6 cabos por extremidade (ex.: Bunbury, 5 + 1)."],
           cor, size=11.5, num=2)

# ══════════════════════════════════════════════════════════════
# 8 · Arranjo adotado na simulação
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", 8, "Etapa 3 · Cenário analisado", "Arranjo adotado na simulação",
               "Mesmo número de cabos do layout típico: 6 a vante e 6 a ré, em 5 cabeços",
               f"Fonte: {NOTEBOOK} · esquema com as posições longitudinais de buzinas e cabeços; "
               "buzinas desenhadas no costado de atracação", etapa="cenario")
por_imagem(sl, "00_plano_amarracao.png", X0, Y_CONT, LARG, 2.75, pad=0.1)
cy = Y_CONT + 2.9
for k, (rot, corpo) in enumerate([
        ("A vante · 6 cabos", ["Cabos **3 a 8**, em pares nos cabeços **A, B e G**.",
                               "Cada cabeço recebe dois cabos."]),
        ("A ré · 6 cabos", ["Cabos **1, 2, 11 e 12** no cabeço **S**.",
                            "Cabos **9 e 10** no cabeço **N**."]),
        ("O que observar", ["O cabeço S recebe 4 dos 12 cabos e soma a carga de todos eles.",
                            "Os cabos 9 e 10 dividem o cabeço N com comprimentos diferentes, o que tende "
                            "a desequilibrar a carga."])]):
    card_texto(sl, X0 + k * (cw + 0.15), cy, cw, Y_FIM - cy, rot, corpo, cor, size=12.5)

# ══════════════════════════════════════════════════════════════
# 9 · Como o trabalho foi conduzido
# ══════════════════════════════════════════════════════════════
sl, cor = base("abordagem", 9, "Etapa 4 · Abordagem", "Como o trabalho foi conduzido",
               "Quatro etapas, da caracterização do berço à leitura dos resultados",
               "Fontes: SILVA (2026a; 2026b); OCIMF (2018, Figura 1.4). Base teórica do TF I: forças de vento "
               "e corrente pelo MEG4 (F = ½·C·ρ·V²·A) e decomposição da tração (T·cos θ e T·sen θ)",
               etapa="abordagem")
sw, sh, sgap = (LARG - 3 * 0.45) / 4, 2.7, 0.45
passos = [("Caracterização",
           "Orientação do píer, espaçamento de cabeços e defensas, maré da DHN e dados do Panamax, "
           "levantados no TF I."),
          ("Simulação",
           "394 combinações de vento, corrente, onda (Hs, Tp), maré e direções, com calado de 9 m, "
           "exportadas em três arquivos CSV."),
          ("Pós-processamento",
           "Notebook em Python (TF4.0): consolida os arquivos, calcula estatísticas, filtra riscos e "
           "monta o envelope crítico e a matriz de operabilidade."),
          ("Geometria",
           "Comprimento 3D e ângulo vertical de cada cabo, para calados de 9 e 14 m e marés de 0 e 1,58 m.")]
for k, (rot, corpo) in enumerate(passos):
    xx = X0 + k * (sw + sgap)
    card(sl, xx, Y_CONT, sw, sh)
    texto(sl, xx + 0.2, Y_CONT + 0.12, 0.8, 0.55, str(k + 1), 28, cor, bold=True)
    rotulo(sl, xx + 0.2, Y_CONT + 0.72, sw - 0.4, rot, cor)
    texto(sl, xx + 0.2, Y_CONT + 1.05, sw - 0.4, sh - 1.15, corpo, 12, TXT2)
    if k < 3:
        seta(sl, xx + sw + 0.12, Y_CONT + sh / 2 - 0.2, 0.21, 0.4, cor)
ky = Y_CONT + sh + 0.2
card(sl, X0, ky, LARG, Y_FIM - ky)
rotulo(sl, X0 + 0.25, ky + 0.18, 6, "Critérios de avaliação", cor)
kw = (LARG - 0.5) / 3
for k, (v, leg) in enumerate([("55% MBL", "tração máxima nos cabos adotada no notebook; o MEG4 indica WLL de "
                                          "55% (aço) e 50% (sintéticos) do MBL de projeto"),
                              ("100% SWL", "força máxima nos cabeços; SWL de 75 t, obtido da razão "
                                           "força/SWL do notebook"),
                              ("≤ 25°", "ângulo vertical dos cabos em relação ao plano do cais (MEG4)")]):
    xx = X0 + 0.25 + k * kw
    texto(sl, xx, ky + 0.5, kw - 0.3, 0.5, v, 26, cor, bold=True)
    texto(sl, xx, ky + 1.05, kw - 0.3, 0.75, leg, 11.5, TXT3)

# ══════════════════════════════════════════════════════════════
# 10 · Resumo estatístico (tabela)
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", 10, "Etapa 5 · Resultados · Visão geral", "Resumo estatístico geral",
               "A mediana fica em 31,5% MBL, mas o percentil 90 já ultrapassa o limite de 55%",
               f"Fonte: {NOTEBOOK}, tabela “Resumo estatístico geral” (percentis 50, 90, 95 e 99)",
               etapa="resultados")
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
tabela(sl, X0 + 0.25, Y_CONT + 0.35, IMG["w"] - 0.5, cab, lin, [2.6, 1, 1.1, 0.95, 0.95, 0.95, 1.05],
       alt_linha=0.52, size=12, cor_cabec=cor, destaques=dest)
forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, X0 + 0.25, Y_CONT + 4.2, 0.28, 0.2, ALERTA_FILL, BORDA, raio=0.2)
texto(sl, X0 + 0.62, Y_CONT + 4.17, 6, 0.3,
      "acima do limite (55% MBL nos cabos; 100% SWL nos cabeços)", 11, TXT3)
colunas_texto(sl, cor,
              ["Estatísticas de seis grandezas nas 394 condições: tração nos cabos (tf e % MBL), reação "
               "e compressão nas defensas e força nos cabeços (tf e % SWL).",
               "**Intuito:** dar a ordem de grandeza dos esforços; os percentis P90 a P99 representam as "
               "condições mais severas."],
              ["Metade das condições fica abaixo de 31,5% MBL, mas o P90 já ultrapassa o limite (59,6%) "
               "e o máximo chega a 75,3%.",
               "Nos cabeços, apenas o máximo (110,8%) passa de 100% SWL. O notebook não define limite "
               "para as defensas."],
              "31,5% MBL", "mediana da tração nos cabos", rot1="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 11–13 · Visão geral e fatores de carga
# ══════════════════════════════════════════════════════════════
slide_grafico(11, "Visão geral", "Distribuição dos esforços",
              "58 das 394 condições (14,7%) passam do limite nos cabos; nos cabeços, apenas 2",
              "01_histogramas.png",
              ["Número de condições em cada faixa de carregamento: % MBL nos cabos e % SWL nos "
               "cabeços. A linha tracejada marca o limite.",
               "**Intuito:** verificar se as violações são casos isolados ou frequentes."],
              ["Nos cabos, a faixa mais frequente é 10–16% MBL (92 condições), mas a distribuição se "
               "estende até 75%.",
               "Nos cabeços, a faixa mais frequente é 30–37% SWL. O risco é recorrente nos cabos e "
               "pontual nos cabeços."],
              "58 de 394", "condições acima de 55% MBL (14,7%)")

slide_grafico(12, "Fatores de carga", "Influência da maré",
              "Só o cabo 9 passa de 55% MBL, com maré alta; o cabeço S excede o SWL com maré baixa",
              "02_influencia_mare.png",
              ["Maior esforço de cada cabo (círculos) e de cada cabeço (triângulos); a cor indica o "
               "nível de maré no instante desse pico.",
               "**Intuito:** identificar os elementos que chegam ao limite e o nível de maré em que "
               "isso ocorre."],
              ["Os outros 11 cabos ficam abaixo de 40% MBL. O pico do cabo 9 ocorre com maré próxima "
               "de 1,5 m, quando os cabos estão mais inclinados.",
               "O cabeço S chega a 110,8% SWL com cerca de 0,3 m: a maré não age no mesmo sentido em "
               "todos os elementos."],
              "75,3% MBL", "pico do cabo 9, ligado ao cabeço N")

slide_grafico(13, "Fatores de carga", "Clima ou arranjo de cabos?",
              "Nenhuma variável ambiental isolada explica a tração; a sobrecarga se concentra no cabo 9",
              "04_correlacao_boxplot.png",
              ["À esquerda, a correlação linear entre ambiente e esforços (perto de ±1 = forte; perto de "
               "0 = fraca). À direita, a % MBL agrupada pelo cabo mais carregado em cada condição.",
               "**Intuito:** separar o efeito do clima do efeito do arranjo."],
              ["Cada variável ambiental tem correlação fraca com a tração nos cabos (|r| < 0,1). As "
               "direções não entram na matriz.",
               "No boxplot, apenas o cabo 9 ultrapassa 55%, o que aponta para o arranjo dos cabos."],
              "r = 0,70", "correlação entre % MBL e % SWL")

# ══════════════════════════════════════════════════════════════
# 14 · Combinações críticas (tabelas)
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", 14, "Etapa 5 · Resultados · Riscos", "Combinações críticas",
               "O cabo 9 é o elemento crítico nas cinco piores condições; o cabeço S excede o SWL com vento forte",
               f"Fonte: {NOTEBOOK}, bloco “Filtro de riscos e piores combinações”", etapa="resultados")
card(sl, **IMG)
tx, tw = X0 + 0.25, IMG["w"] - 0.5
texto(sl, tx, Y_CONT + 0.2, tw, 0.3, "Cabos · cinco piores condições (de 58 acima de 55% MBL)",
      12.5, TXT, bold=True)
tabela(sl, tx, Y_CONT + 0.58, tw,
       ["Cond.", "Vento [nós]", "Corrente [nós]", "Hs [m]", "Tração [tf]", "% MBL", "Cabo"],
       [["259", "8,20", "1,57", "1,00", "56,44", "75,30", "9"],
        ["310", "41,40", "1,16", "1,19", "54,72", "73,00", "9"],
        ["78", "5,30", "1,43", "0,35", "53,95", "71,90", "9"],
        ["282", "29,40", "1,58", "1,08", "51,82", "69,10", "9"],
        ["272", "39,90", "0,47", "1,05", "51,85", "69,10", "9"]],
       [0.8, 1.1, 1.3, 0.85, 1.1, 0.9, 0.7], alt_linha=0.37, cor_cabec=cor,
       destaques={(i, 5): ALERTA_FILL for i in range(5)})
texto(sl, tx, Y_CONT + 2.98, tw, 0.3, "Cabeços · todas as condições acima de 100% SWL",
      12.5, TXT, bold=True)
tabela(sl, tx, Y_CONT + 3.36, tw,
       ["Cond.", "Vento [nós]", "Corrente [nós]", "Hs [m]", "Força [tf]", "% SWL", "Cabeço"],
       [["198", "36,70", "1,58", "0,78", "83,13", "110,80", "S"],
        ["55", "49,10", "0,88", "0,30", "78,34", "104,50", "S"]],
       [0.8, 1.1, 1.3, 0.85, 1.1, 0.9, 0.7], alt_linha=0.37, cor_cabec=cor,
       destaques={(i, 5): ALERTA_FILL for i in range(2)})
colunas_texto(sl, cor,
              ["As cinco condições que mais carregam os cabos e todas as que ultrapassam o limite nos "
               "cabeços, com vento, corrente e altura de onda (Hs).",
               "**Intuito:** identificar as combinações ambientais que levam ao limite e o elemento que "
               "falha primeiro."],
              ["O pior caso (condição 259) ocorre com vento de apenas 8,2 nós, mas com corrente de "
               "1,57 nó e Hs de 1,0 m.",
               "As duas violações do cabeço S ocorrem com vento forte (36,7 e 49,1 nós)."],
              "5 de 5", "piores condições têm o cabo 9 como crítico", rot1="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 15–16 · Riscos
# ══════════════════════════════════════════════════════════════
slide_grafico(15, "Riscos", "Envelope crítico: vento × corrente",
              "A maioria dos cabos atinge o pico com vento forte; o cabo 9, com corrente forte e maré alta",
              "05_envelope_critico.png",
              ["Cada bolha é a condição de pico de um elemento, posicionada pela corrente (eixo x) e "
               "pelo vento (eixo y); a cor indica o esforço e o tamanho, a maré.",
               "**Intuito:** mapear em que combinação ambiental cada parte do sistema é mais exigida."],
              ["Os demais cabos atingem o pico com vento de 37 a 49 nós.",
               "Os cabeços A, B e S têm o pico na mesma condição (36,7 nós, maré baixa); o cabeço N, "
               "junto com o cabo 9, que está preso a ele."],
              "8,2 nós", "vento na condição de pico do cabo 9")

slide_grafico(16, "Riscos", "Matriz de operabilidade",
              "19 das 20 faixas de vento e corrente têm ao menos uma condição acima do limite",
              "06_matriz_operabilidade.png",
              ["As condições são agrupadas em faixas de vento (10 nós) e corrente (0,5 nó); cada célula "
               "mostra a pior tração da faixa (verde abaixo de 55%; vermelho acima).",
               "**Intuito:** apoiar a definição de limites de paragem do berço."],
              ["O limite é ultrapassado até com vento e corrente fracos (61,2%). Onda, direção e maré "
               "precisam entrar no critério.",
               "A célula verde (38,8%) deve reunir poucas condições e merece verificação."],
              "19 de 20", "faixas com pior caso acima de 55% MBL")

# ══════════════════════════════════════════════════════════════
# 17–18 · Geometria dos cabos
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
        xx = X0 + 0.25 + k * 3.3
        forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, y, 0.28, 0.2, fill, BORDA, raio=0.2)
        texto(sl, xx + 0.37, y - 0.03, 2.9, 0.3, txt, 11, TXT3)


sl, cor = base("resultados", 17, "Etapa 5 · Resultados · Geometria", "Geometria dos cabos · maré de 1,58 m",
               "Pelo cálculo do notebook, todos os cabos ficam abaixo do limite de 25° do MEG4",
               f"Fonte: {NOTEBOOK}, calcular_geometria_amarracao (calado 9 m, maré 1,58 m)",
               etapa="resultados")
card(sl, **IMG)
tabela(sl, X0 + 0.25, Y_CONT + 0.18, IMG["w"] - 0.5,
       ["Cabo", "Buzina", "Cabeço", "Comp. 3D [m]", "Âng. vertical [°]"],
       [[g[0], g[1], g[2], f(g[3]), f(g[4])] for g in GEO],
       [0.8, 0.9, 0.9, 1.4, 1.5], alt_linha=0.31, size=11.5, cor_cabec=cor,
       destaques={(8, None): ALERTA_FILL, (4, None): DESTAQUE_FILL, (8, 3): ALERTA_FILL,
                  (9, 3): "#F3EFE7", (4, 4): DESTAQUE_FILL})
legenda_geo(sl, Y_CONT + 4.42)
colunas_texto(sl, cor,
              ["Comprimento 3D e ângulo vertical de cada cabo, da buzina ao cabeço, no cenário de maior "
               "inclinação (calado de 9 m e maré alta).",
               "**Intuito:** verificar o limite de 25° e relacionar o esforço de cada cabo à sua geometria."],
              ["O cabo 9 é 9,4 m mais curto que o cabo 10, que divide o cabeço N. Mais rígido, tende a "
               "absorver mais carga, o que é coerente com os picos.",
               "Esses valores dependem da revisão da geometria (próximos passos)."],
              "21,6° < 25°", "maior ângulo vertical (cabo 5) e limite do MEG4", rot1="O que a tabela lê")

sl, cor = base("resultados", 18, "Etapa 5 · Resultados · Geometria", "Geometria dos cabos · maré de 0 m",
               "A maré baixa reduz o ângulo vertical de todos os cabos em 2,2° a 4,9°",
               f"Fonte: {NOTEBOOK}, calcular_geometria_amarracao (calado 9 m, maré 0 m) · Δ em relação à "
               "maré de 1,58 m · calado de 14 m pela mesma função", etapa="resultados")
card(sl, **IMG)
tabela(sl, X0 + 0.25, Y_CONT + 0.18, IMG["w"] - 0.5,
       ["Cabo", "Buzina", "Cabeço", "Comp. 3D [m]", "Âng. vertical [°]", "Δ ângulo [°]"],
       [[g[0], g[1], g[2], f(g[5]), f(g[6]), "−" + f(g[4] - g[6])] for g in GEO],
       [0.75, 0.85, 0.85, 1.25, 1.4, 1.15], alt_linha=0.31, size=11.5, cor_cabec=cor,
       destaques={(8, None): ALERTA_FILL, (4, None): DESTAQUE_FILL, (4, 5): DESTAQUE_FILL})
legenda_geo(sl, Y_CONT + 4.42)
colunas_texto(sl, cor,
              ["O mesmo cálculo com maré de 0 m; a última coluna mostra a redução do ângulo vertical em "
               "relação à maré de 1,58 m.",
               "**Intuito:** isolar o efeito da maré na geometria, com o calado fixo."],
              ["A buzina desce 1,58 m e o comprimento dos cabos muda apenas 0,2 a 0,5 m.",
               "Com calado de 14 m, o ângulo cai para 0,1° a 5,2°. A maré altera a inclinação, mas não "
               "a diferença entre os cabos 9 e 10."],
              "−2,2° a −4,9°", "variação do ângulo vertical com a maré baixa", rot1="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# 19 · Conclusões parciais
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", 19, "Etapa 6 · Conclusões", "Conclusões parciais",
               "Os resultados respondem ao objetivo desta etapa",
               "Fontes: SILVA (2026a; 2026b)", etapa="conclusoes")
achados = [
    ("C1", "O cabo 9 concentra o risco",
     "Único cabo acima de 55% MBL (pico de 75,3%) e elemento crítico nas cinco piores condições."),
    ("C2", "O cabeço S excede o SWL",
     "Duas condições acima de 100% (até 110,8%, cerca de 83 t para SWL de 75 t); recebe 4 dos 12 cabos."),
    ("C3", "O ambiente isolado não explica",
     "Correlações fracas e 19 de 20 faixas de vento × corrente com pior caso acima do limite."),
    ("C4", "A maré altera a geometria",
     "A maré alta inclina os cabos de 2° a 5° a mais; pelo cálculo atual, todos ficam abaixo de 25°."),
]
gw, gh = (LARG - 0.45) / 4, 2.85
for k, (cid, tit, corpo) in enumerate(achados):
    xx = X0 + k * (gw + 0.15)
    card(sl, xx, Y_CONT, gw, gh)
    texto(sl, xx + 0.2, Y_CONT + 0.18, gw - 0.4, 0.3, cid, 11, TXT3, bold=True)
    texto(sl, xx + 0.2, Y_CONT + 0.48, gw - 0.4, 0.7, tit, 15.5, cor, bold=True)
    texto(sl, xx + 0.2, Y_CONT + 1.25, gw - 0.4, gh - 1.35, corpo, 12.5, TXT2)
py = Y_CONT + gh + 0.2
card(sl, X0, py, LARG, Y_FIM - py, fill=tint(cor, 0.08))
forma(sl, MSO_SHAPE.RECTANGLE, X0, py + 0.25, 0.07, Y_FIM - py - 0.5, cor)
rotulo(sl, X0 + 0.32, py + 0.22, 6, "Resposta ao objetivo", cor)
texto(sl, X0 + 0.32, py + 0.56, LARG - 0.6, Y_FIM - py - 0.66,
      "A simulação identificou as condições e os elementos que levam o berço ao limite: o cabo 9, no "
      "cabeço N, e o cabeço S. São esses os pontos que o monitoramento por visão computacional deve "
      "acompanhar prioritariamente.", 14, TXT2)

# ══════════════════════════════════════════════════════════════
# 20 · Próximos passos
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", 20, "Etapa 6 · Conclusões", "Próximos passos",
               "Do diagnóstico parcial à entrega final",
               "Próximos passos propostos a partir dos resultados desta etapa e do plano de SILVA (2026a)",
               etapa="conclusoes")
etapas = [("TF I · concluído",
           "Base teórica: equipamentos, defensas, cabos PPL/PE, forças ambientais (MEG4) e caracterização "
           "do T-Grão."),
          ("Etapa atual",
           "Simulação de 394 condições com calado de 9 m e análise dos esforços, dos riscos e da geometria."),
          ("Próxima etapa",
           "Ampliar os cenários simulados e desenvolver o modelo de monitoramento.")]
ew, eh = (LARG - 2 * 0.5) / 3, 1.85
for k, (rot, corpo) in enumerate(etapas):
    xx = X0 + k * (ew + 0.5)
    card(sl, xx, Y_CONT, ew, eh, fill=tint(cor, 0.12) if k == 1 else CARD)
    rotulo(sl, xx + 0.2, Y_CONT + 0.18, ew - 0.4, rot, cor)
    texto(sl, xx + 0.2, Y_CONT + 0.52, ew - 0.4, eh - 0.6, corpo, 12, TXT2)
    if k < 2:
        seta(sl, xx + ew + 0.14, Y_CONT + eh / 2 - 0.2, 0.22, 0.4, cor)
ly = Y_CONT + eh + 0.2
card(sl, X0, ly, LARG, Y_FIM - ly)
rotulo(sl, X0 + 0.25, ly + 0.2, 6, "O que vem a seguir", cor)
passos_a = ["•  Simular o calado carregado (14 m) e a maré mínima de 2025 (−0,16 m)",
            "•  Reavaliar os resultados com 50% do MBL, caso os cabos sejam sintéticos (MEG4)",
            "•  Incluir Hs, direções e maré na matriz de operabilidade",
            "•  Avaliar o efeito de navios passantes no canal"]
passos_b = ["•  Revisar o cálculo de geometria (lado das buzinas e posição dos cabeços) e validá-lo com o berço real",
            "•  Testar ajustes no arranjo: comprimento ou pré-tensão do cabo 9 e menos cabos no cabeço S",
            "•  Desenvolver o monitoramento dos cabos por visão computacional e machine learning"]
bw = (LARG - 0.75) / 2
texto(sl, X0 + 0.25, ly + 0.55, bw, Y_FIM - ly - 0.65, passos_a, 12.5, TXT2, entre=7)
texto(sl, X0 + 0.5 + bw, ly + 0.55, bw, Y_FIM - ly - 0.65, passos_b, 12.5, TXT2, entre=7)

# ══════════════════════════════════════════════════════════════
# 21 · Referências (ABNT NBR 6023)
# ══════════════════════════════════════════════════════════════
sl, cor = base("referencias", 21, "Apresentação", "Referências",
               "Fontes citadas nos slides, conforme a ABNT NBR 6023")
REFS_A = [
    "ACP – AUTORIDAD DEL CANAL DE PANAMÁ. **OP Notice to Shipping N-1-2026**. Panamá, 2026.",
    "MARINHA DO BRASIL. Diretoria de Hidrografia e Navegação. **Tábuas das marés 2025**: Porto de Santos "
    "(SP). Niterói: DHN, 2025.",
    "OCIMF – OIL COMPANIES INTERNATIONAL MARINE FORUM. **Mooring Equipment Guidelines (MEG4)**. 4. ed. "
    "Londres: Witherby Publishing Group, 2018.",
    "PIANC. **Design principles for dry bulk marine terminals**. Bruxelas: PIANC, 2017.",
    "PIANC. **Mooring of large ships at quay walls**. Bruxelas: PIANC, 2025. (MarCom Working Group "
    "Report, n. 186).",
    "SILVA, G. G. **Sistema de amarração em conjuntos portuários**. 2026a. Trabalho de Formatura I "
    "(PNV3510) – Escola Politécnica, Universidade de São Paulo, São Paulo, 2026.",
]
REFS_B = [
    "SILVA, G. G. **TF4.0**: análise dos resultados de simulação de amarração (calado 9 m, 394 "
    "condições). São Paulo, 2026b. Notebook Python.",
    "SOUTHERN PORTS. **Mooring guide to Panamax ships**: loading grain – Berth 3 Bunbury. Bunbury: "
    "Southern Ports Authority, 2026. Disponível em: https://www.southernports.com.au. Acesso em: "
    "9 nov. 2026.",
    "T-GRÃO CARGO TERMINAL DE GRANÉIS S/A. **Página inicial**. Santos, 2026. Disponível em: "
    "https://tgrao.com.br/. Acesso em: 19 jun. 2026.",
    "UK P&I CLUB. Understanding mooring incidents. **LP News**, Londres, jan. 2009. Disponível em: "
    "https://www.ukpandi.com. Acesso em: 9 nov. 2026.",
]
rw = (LARG - 0.15) / 2
for k, refs in enumerate((REFS_A, REFS_B)):
    xx = X0 + k * (rw + 0.15)
    card(sl, xx, Y_CONT, rw, Y_FIM - Y_CONT)
    texto(sl, xx + 0.25, Y_CONT + 0.22, rw - 0.5, Y_FIM - Y_CONT - 0.35, refs, 12, TXT2, entre=10)

# ══════════════════════════════════════════════════════════════
# 22 · Encerramento
# ══════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(prs.slide_layouts[6])
capa_fundo(sl)
texto(sl, X0, 2.25, 11, 1.1, "Obrigado pela atenção", 50, TXT, bold=True)
texto(sl, X0, 3.65, 11, 1.1, "Perguntas?", 50, TXT, bold=True)
texto(sl, X0 + 0.05, 5.0, 8, 0.4, "Gustavo Guilherme Silva · NUSP 11820555", 15, TXT2)
numero(sl, TOTAL)

assert len(prs.slides) == TOTAL, len(prs.slides)
prs.save(SAIDA)
print("Salvo:", SAIDA, "·", len(prs.slides), "slides")
