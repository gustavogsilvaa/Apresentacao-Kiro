"""
Apresentação parcial do TF: sistema de amarração em conjuntos portuários (v8, pasta TF).
Mesma estrutura e layout da v7 (amarracao_v3), atualizada com a simulação nova:
5 condições de carga (lastro 9 m a cheio 14 m), 394 condições ambientais cada,
cabos 11 e 12 no cabeço R e cabo com MBL de 736 kN (75,05 t).

Dados em dados/ (ver dados_tf.py); gráficos novos em imagens/ (gerar_graficos.py).

Estilo baseado no PDF do PINN: fundo creme, barra lateral colorida por etapa,
títulos serifados, cards claros e formas decorativas no canto. Toda a tipografia
usa Times New Roman.

Fluxo de leitura em cada slide:
  rótulo da etapa → título (tema) → subtítulo (mensagem principal) → evidência → cards 1 e 2
A barra de progresso no topo mostra em qual etapa a apresentação está.
A numeração dos slides é automática.

Imagens: usa imagens_originais/<nome>.png se existir (PNGs exportados do Colab);
caso contrário usa imagens_recriadas/<nome>.png (gerar_graficos.py).
"""
import os
import re

import numpy as np
import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

import dados_tf as D

BASE = os.path.dirname(os.path.abspath(__file__))
SAIDA = os.path.join(BASE, "Apresentacao_TF_Amarracao_parcial_v8.pptx")
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
TOTAL = 27
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


def numero(sl):
    texto(sl, 11.6, 7.07, 1.4, 0.25, f"{len(prs.slides)} / {TOTAL}", 11, TXT3, align=PP_ALIGN.RIGHT)


def trilha(sl, atual):
    """Barra de progresso com as seis etapas; a etapa atual fica destacada."""
    cw_, gap, h = 0.92, 0.05, 0.27
    x = X0 + LARG - (len(ETAPAS) * cw_ + (len(ETAPAS) - 1) * gap)
    idx = [k for k, _ in ETAPAS].index(atual)
    for i, (k, nome) in enumerate(ETAPAS):
        xx = x + i * (cw_ + gap)
        if i == idx:
            forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, 0.4, cw_, h, SECAO[k], raio=0.5)
            cor_t, b = "#FFFFFF", True
        else:
            forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, 0.4, cw_, h,
                  tint(SECAO[k], 0.22 if i < idx else 0.07), raio=0.5)
            cor_t, b = (TXT2 if i < idx else TXT3), False
        texto(sl, xx, 0.4, cw_, h, nome, 9.5, cor_t, bold=b, align=PP_ALIGN.CENTER,
              ancora=MSO_ANCHOR.MIDDLE, linha=1.0)


def base(secao, rot=None, titulo=None, mensagem=None, nota=None, etapa=None):
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
    numero(sl)
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
    """imagens/ (gráficos novos), imagens_originais/ (do autor) ou imagens_recriadas/ (esquemas da v7)."""
    for pasta in ("imagens", "imagens_originais", "imagens_recriadas"):
        p = os.path.join(BASE, pasta, nome)
        if os.path.exists(p):
            return p, pasta == "imagens_originais"
    raise FileNotFoundError(nome)


def figura(sl, nome, x, y, w, h):
    """Imagem centralizada na caixa (x, y, w, h), mantendo a proporção."""
    caminho, original = imagem(nome)
    iw, ih = Image.open(caminho).size
    esc = min(w / iw, h / ih)
    pw, ph = iw * esc, ih * esc
    sl.shapes.add_picture(caminho, Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2),
                          Inches(pw), Inches(ph))
    return original


def por_imagem(sl, nome, x, y, w, h, pad=0.14):
    """Card branco com a imagem centralizada."""
    card(sl, x, y, w, h, fill="#FFFFFF")
    return figura(sl, nome, x + pad, y + pad, w - 2 * pad, h - 2 * pad)


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


def nota_imagem(original, condicao=None):
    if condicao:
        t = f"Fonte: {NOTEBOOK}, resumo de esforços máximos da condição {condicao} (394 condições)"
    else:
        t = f"Fonte: {NOTEBOOK} (5 condições de carga × 394 condições ambientais)"
    if not original:
        t += " · gráfico elaborado a partir das saídas do notebook"
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


def slide_grafico(sub, titulo, mensagem, nome_img, le, resultado, kpi, kpi_leg, rodape=True, condicao=None):
    sl, cor = base("resultados", f"Etapa 5 · Resultados · {sub}", titulo, mensagem, etapa="resultados")
    original = por_imagem(sl, nome_img, **IMG)
    colunas_texto(sl, cor, le, resultado, kpi, kpi_leg)
    if rodape:
        texto(sl, X0, 7.06, 10.7, 0.3, nota_imagem(original, condicao), 9.5, TXT3)
    return sl, cor


def existe_original(nome):
    return os.path.exists(os.path.join(BASE, "imagens_originais", nome))


def capa_fundo(sl):
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CREME)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, 9.7, 3.9, 4.6, 4.6, tint(SECAO["capa"], 0.18), raio=0.25)
    forma(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.14, H, SECAO["capa"])


cw = (LARG - 0.3) / 3          # largura de card em grade de 3 colunas

# ══════════════════════════════════════════════════════════════
# Capa
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
numero(sl)

# ══════════════════════════════════════════════════════════════
# Roteiro
# ══════════════════════════════════════════════════════════════
sl, cor = base("roteiro", "Apresentação", "Roteiro", "Do problema aos resultados parciais, em seis etapas")
itens = [("motivacao", "Motivação", "Por que a amarração é crítica nos terminais graneleiros de Santos"),
         ("justificativa", "Justificativa", "A lacuna no monitoramento dos cabos e o objetivo desta etapa"),
         ("cenario", "Cenário analisado", "Terminal T-Grão, o navio e seus movimentos, maré de Santos, layout "
                                          "típico do Panamax e arranjo simulado"),
         ("abordagem", "Abordagem", "Etapas do trabalho e critérios de avaliação"),
         ("resultados", "Resultados", "Leitura dos gráficos da simulação: esforços, riscos, geometria e "
                                      "efeito do carregamento"),
         ("conclusoes", "Conclusões", "Conclusões parciais, alternativas de segurança e próximos passos")]
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
# Etapa 1 · Motivação
# ══════════════════════════════════════════════════════════════
sl, cor = base("motivacao", "Etapa 1 · Motivação", "Por que estudar a amarração?",
               "Durante o carregamento, o navio é mantido no berço apenas pelo sistema de amarração",
               etapa="motivacao")
tw_ = (LARG - 0.45) / 4
for k, (v, leg) in enumerate([(">100 Mt", "de carga por ano no Porto de Santos (BRASIL, 2026)."),
                              ("14 de 60", "berços dedicados a granéis vegetais (SILVA, 2026a)."),
                              ("US$ 34 mi", "em sinistros com equipamentos de amarração em 20 anos "
                                            "(UK P&I CLUB, 2009)."),
                              ("até 40%", "de aumento da tração nos cabos com a passagem de outro navio "
                                          "(SILVA, 2026a).")]):
    kpi_tile(sl, X0 + k * (tw_ + 0.15), Y_CONT, tw_, 1.4, v, leg, cor)
cy, mh = Y_CONT + 1.55, 1.65
for k, (rot, corpo) in enumerate([
        ("Segurança",
         ["A ruptura de um cabo tracionado libera de uma só vez a energia elástica acumulada (**snap-back**)."]),
        ("Operação",
         ["Se o navio se afasta do cais, shiploaders e esteiras podem ser danificados e o embarque é "
          "interrompido."]),
        ("Meio ambiente",
         ["O despejo de grãos no estuário eleva a demanda bioquímica de oxigênio (DBO) da água.",
          "Isso compromete as metas de **Porto Verde** (ESG)."])]):
    card_texto(sl, X0 + k * (cw + 0.15), cy, cw, mh, rot, corpo, cor, size=12)
sy = cy + mh + 0.15
faixa(sl, sy, 1.15,
      "**Em síntese:** uma falha de amarração atinge, ao mesmo tempo, pessoas, operação e meio ambiente. "
      "Conhecer os esforços nos cabos é o primeiro passo para preveni-la.", cor)

# ══════════════════════════════════════════════════════════════
# Etapa 2 · Justificativa e objetivo
# ══════════════════════════════════════════════════════════════
sl, cor = base("justificativa", "Etapa 2 · Justificativa", "Justificativa e objetivo",
               "A tração nos cabos ainda não é medida de forma contínua nos terminais",
               etapa="justificativa")
fw, fh, fgap = (LARG - 2 * 0.5) / 3, 2.45, 0.5
fluxo = [("O problema", "A tração muda o tempo todo",
          ["Vento, corrente, maré e navios passantes alteram a carga nos cabos, que raramente é medida "
           "durante a operação."]),
         ("A lacuna", "Sensores se degradam no mar",
          ["Células de carga e extensômetros sofrem com o sal e a umidade; o acompanhamento fica "
           "restrito a inspeções visuais."]),
         ("A proposta do TF", "Medir a tração sem contato",
          ["Monitoramento de saúde estrutural (**SHM**): câmeras acompanham os cabos e um modelo de "
           "machine learning estima a tração.",
           "As simulações indicam os cenários críticos."])]
for k, (rot, manchete, corpo) in enumerate(fluxo):
    xx = X0 + k * (fw + fgap)
    destaque = k == 2                    # a proposta é o ponto de chegada do raciocínio
    card(sl, xx, Y_CONT, fw, fh, fill=tint(cor, 0.12) if destaque else CARD)
    selo(sl, xx + 0.2, Y_CONT + 0.15, k + 1, cor)
    rotulo(sl, xx + 0.57, Y_CONT + 0.19, fw - 0.8, rot, cor)
    texto(sl, xx + 0.2, Y_CONT + 0.6, fw - 0.4, 0.35, manchete, 15.5, TXT, bold=True)
    texto(sl, xx + 0.2, Y_CONT + 1.02, fw - 0.4, fh - 1.1, corpo, 12.5, TXT2, entre=4)
    if k < 2:
        seta(sl, xx + fw + 0.14, Y_CONT + fh / 2 - 0.2, 0.22, 0.4, cor)
oy = Y_CONT + fh + 0.2
oh = 1.95
card(sl, X0, oy, LARG, oh, fill=tint(cor, 0.09))
forma(sl, MSO_SHAPE.RECTANGLE, X0, oy + 0.2, 0.07, oh - 0.4, cor)
rotulo(sl, X0 + 0.32, oy + 0.2, 4.3, "Objetivo desta etapa", cor)
texto(sl, X0 + 0.32, oy + 0.5, 4.3, oh - 0.6,
      "Identificar, por simulação, as condições e os elementos que levam o sistema de amarração ao "
      "limite: os cenários que o monitoramento deverá detectar.", 14, TXT)
mx0 = X0 + 4.85
mw = (X0 + LARG - 0.2 - mx0 - 2 * 0.12) / 3
for k, (verbo, compl) in enumerate([("Quantificar", "os esforços nas 394 condições, em 5 condições de carga."),
                                   ("Identificar", "as condições e os elementos críticos."),
                                   ("Relacionar", "a carga à geometria dos cabos.")]):
    xx = mx0 + k * (mw + 0.12)
    card(sl, xx, oy + 0.2, mw, oh - 0.4)
    texto(sl, xx + 0.18, oy + 0.35, mw - 0.36, 0.35, verbo, 15.5, cor, bold=True)
    texto(sl, xx + 0.18, oy + 0.78, mw - 0.36, oh - 1.05, compl, 12.5, TXT2)

# ══════════════════════════════════════════════════════════════
# Etapa 3 · Cenário analisado (2 cards de contexto + 3 cards de ambiente)
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", "Etapa 3 · Cenário analisado", "Cenário analisado",
               "Panamax atracado por boreste no T-Grão, sob a maré, o vento e a corrente de Santos",
               etapa="cenario")
rh = (Y_FIM - Y_CONT - 0.15) / 2
x_top = X0 + (LARG - (2 * cw + 0.15)) / 2
for k, (rot, kpi, leg, corpo) in enumerate([
        ("Terminal T-Grão", "149°", "orientação do píer",
         ["•  Atracação por boreste",
          "•  Cabeços e defensas a cada 15 m (a PIANC indica 12 a 25 m)"]),
        ("Navio Panamax", "225 m", "comprimento total (LOA)",
         ["•  Boca de 32,26 m e pontal de 18,5 m",
          "•  5 condições de carga, do lastro (9 m) à plena carga (14 m)"])]):
    card_texto(sl, x_top + k * (cw + 0.15), Y_CONT, cw, rh, rot, corpo, cor, kpi=kpi, kpi_legenda=leg,
               size=12.5)
yb = Y_CONT + rh + 0.15
for k, (rot, kpi, leg, corpo) in enumerate([
        ("Ambiente · maré", "−0,16 a 1,52 m", "faixa de maré prevista para Santos em 2025",
         ["Águas abrigadas: as ondas chegam atenuadas; vento e corrente de maré predominam.",
          "Na simulação: −0,20 a 1,60 m."]),
        ("Ambiente · vento", "50 nós", "máximo simulado, com direções variáveis",
         ["Inclui vento forte de propósito, para avaliar o efeito da força aerodinâmica no navio e nos "
          "cabos. De través, essa força é muito maior que de proa (OCIMF, 2018)."]),
        ("Ambiente · corrente", "1,6 nó", "corrente máxima simulada, em águas abrigadas",
         ["Enchente e vazante, com desalinhamento de 10° em relação ao eixo do navio: mesmo um ângulo "
          "pequeno gera força transversal relevante (OCIMF, 2018)."])]):
    card_texto(sl, X0 + k * (cw + 0.15), yb, cw, rh, rot, corpo, cor, kpi=kpi, kpi_legenda=leg,
               size=12.5)

# ══════════════════════════════════════════════════════════════
# NOVO · O navio: medidas principais e movimentos (nivelamento)
# ══════════════════════════════════════════════════════════════
MOV75 = D.carregar_resumo("75%").set_index("Elemento")["Valor"]
f = lambda v, d=2: f"{v:.{d}f}".replace(".", ",")
sl, cor = base("cenario", "Etapa 3 · Cenário analisado", "O navio: medidas e movimentos",
               "Como olhamos o navio: as medidas que definem a geometria dos cabos e os movimentos que a "
               "amarração restringe",
               "Fontes: dimensões do Panamax: SILVA (2026a); movimentos máximos: resumo de esforços máximos da "
               "condição 75% (SILVA, 2026b)", etapa="cenario")
dh6 = 2.15
por_imagem(sl, "06_navio_medidas_movimentos.png", X0, Y_CONT, LARG, dh6, pad=0.08)
by6 = Y_CONT + dh6 + 0.15
bh6 = Y_FIM - by6
lw6 = 6.45
card(sl, X0, by6, lw6, bh6)
selo(sl, X0 + 0.2, by6 + 0.15, 1, cor)
rotulo(sl, X0 + 0.57, by6 + 0.19, 5, "Medidas principais", cor)
tabela(sl, X0 + 0.2, by6 + 0.52, lw6 - 0.4,
       ["Medida", "Neste estudo", "O que define na amarração"],
       [["Comprimento total (LOA)", "225 m", "cabeços alcançados e posição das buzinas"],
        ["Boca (B)", "32,26 m", "largura do casco; distância das buzinas ao cais"],
        ["Pontal (D)", "18,5 m", "altura do convés, onde ficam as buzinas"],
        ["Calado (T)", "9 a 14 m", "parte submersa: área sob a corrente"],
        ["Borda livre (D − T)", "9,5 a 4,5 m", "área ao vento e altura da buzina sobre o cais"]],
       [1.75, 1.05, 3.2], alt_linha=0.33, size=10.5, cor_cabec=cor,
       alinhar=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT])
rx6 = X0 + lw6 + 0.15
rw6 = LARG - lw6 - 0.15
card(sl, rx6, by6, rw6, bh6)
selo(sl, rx6 + 0.2, by6 + 0.15, 2, cor)
rotulo(sl, rx6 + 0.57, by6 + 0.19, 4.5, "Movimentos no porto", cor)
tabela(sl, rx6 + 0.2, by6 + 0.52, rw6 - 0.4,
       ["Movimento", "Eixo", "Máximo com 75% de carga"],
       [["Avanço (surge)", "x · longitudinal", f"{f(MOV75['Longitudinal [m]'])} m"],
        ["Deriva (sway)", "y · transversal", f"{f(MOV75['Transversal [m]'])} m"],
        ["Guinada (yaw)", "rotação em z", f"{f(MOV75['Yaw [deg]'])}°"],
        ["Afundamento (heave)", "z · vertical", "lento, com maré e calado"],
        ["Jogo (roll)", "rotação em x", "fora do resumo"],
        ["Arfagem (pitch)", "rotação em y", "fora do resumo"]],
       [1.55, 1.35, 1.75], alt_linha=0.28, size=10.5, cor_cabec=cor,
       destaques={(i, None): DESTAQUE_FILL for i in range(3)},
       alinhar=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER])

# ══════════════════════════════════════════════════════════════
# Estudo de maré
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", "Etapa 3 · Cenário analisado", "Estudo de maré em Santos",
               "A maré sobe até 1,68 m em menos de 6 horas e muda a altura do navio em relação ao cais",
               etapa="cenario")
por_imagem(sl, "08_mare_dhn.png", **IMG)
card_texto(sl, COL["x"], Y_CONT + 0.45, COL["w"], 1.35, "O que o gráfico lê",
           ["Previsões da tábua de marés da DHN para Santos em 2025 (1.453 valores)."], cor, size=12.5, num=1)
card_texto(sl, COL["x"], Y_CONT + 1.95, COL["w"], 2.5, "O que isso significa",
           ["Os cabos precisam ser ajustados ao longo da operação (OCIMF, 2018).",
            "A geometria usa a faixa simulada: −0,20 m e 1,60 m, pouco além da DHN 2025."], cor,
           kpi="1,60 m", kpi_legenda="maré alta usada no cálculo da geometria", size=12.5, num=2)

# ══════════════════════════════════════════════════════════════
# Layout de amarração de um Panamax (referência)
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", "Etapa 3 · Cenário analisado", "Layout de amarração de um Panamax",
               "Cada grupo de cabos resiste principalmente a uma direção de esforço", etapa="cenario")
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
       [["Lançantes", "4 + 4", "Oblíquos, a vante e a ré", "Afastamento das extremidades; avanço e recuo"],
        ["Través", "1 + 1", "Até 15° da perpendicular ao navio", "Afastamento (sway) e guinada (yaw)"],
        ["Espringues", "1 + 1", "Até 10° da direção do costado", "Avanço e recuo (surge)"]],
       [1.05, 0.65, 1.75, 2.35], alt_linha=0.47, size=11, cor_cabec=cor,
       alinhar=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT])
rx = X0 + tw7 + 0.15
card_texto(sl, rx, by, LARG - tw7 - 0.15, bh, "Por que este arranjo",
           ["•  **Vento de través** gera a maior força (OCIMF, 2018) → 4 lançantes e 1 través por "
            "extremidade.",
            "•  **Esforço longitudinal** é menor → 1 espringue por extremidade segura o avanço e o recuo.",
            "•  **Guinada (yaw)** → cabos nas extremidades, longe do centro, com maior braço de alavanca.",
            "•  **Escotilhas** no convés central → guinchos e cabeços do navio na proa e na popa "
            "(SILVA, 2026a)."],
           cor, size=11.5, num=2)

# ══════════════════════════════════════════════════════════════
# Arranjo adotado na simulação
# ══════════════════════════════════════════════════════════════
sl, cor = base("cenario", "Etapa 3 · Cenário analisado", "Arranjo adotado na simulação",
               "Mesmo número de cabos do layout típico: 6 a vante e 6 a ré",
               f"Fonte: {NOTEBOOK} · arranjo de projeto exportado do software de simulação", etapa="cenario")
aw8 = 6.45
por_imagem(sl, "00_arranjo_projeto.png", X0, Y_CONT, aw8, Y_FIM - Y_CONT, pad=0.12)
rx8 = X0 + aw8 + 0.2
rw8 = X0 + LARG - rx8
card_texto(sl, rx8, Y_CONT, rw8, 1.75, "Conjunto de cabos",
           ["**Lançantes de proa:** cabos 3, 4, 5 e 6, em duplas nos cabeços A e B.",
            "**Espringues:** cabos 7, 8, 9 e 10, em duplas nos cabeços G e N.",
            "**Lançantes de popa:** cabos 1, 2, 11 e 12, em duplas nos cabeços S e R."], cor, size=12, num=1)
card_texto(sl, rx8, Y_CONT + 1.9, rw8, Y_FIM - Y_CONT - 1.9, "O que observar",
           ["•  **Arranjo 4-0-2 em cada extremidade:** 4 lançantes, nenhum través e 2 espringues, em vez do "
            "4-1-1 do layout típico.",
            "•  **Por quê:** quando o espaçamento dos cabeços não permite um través eficiente, os portos "
            "adotam o 4-0-2 (SILVA, 2026a, seção 2.3).",
            "•  **Efeito:** sem través, o afastamento do navio fica a cargo dos lançantes.",
            "•  Os cabos 9 e 10 dividem o cabeço N com comprimentos diferentes, o que tende a desequilibrar "
            "a carga."], cor, size=12, num=2)

# ══════════════════════════════════════════════════════════════
# Etapa 4 · Como o trabalho foi conduzido
# ══════════════════════════════════════════════════════════════
sl, cor = base("abordagem", "Etapa 4 · Abordagem", "Como o trabalho foi conduzido",
               "Quatro etapas, da caracterização do berço à leitura dos resultados", etapa="abordagem")
sw, sh, sgap = (LARG - 3 * 0.45) / 4, 3.25, 0.45
passos = [("Caracterização",
           "Levantamento do píer T-Grão, dos cabeços e defensas, da maré (DHN) e das dimensões do Panamax.",
           "Parâmetros de entrada do modelo (TF I)"),
          ("Simulação",
           "394 combinações de vento, corrente, maré e direções, em 5 condições de carga (9 a 14 m).",
           "Resumo de esforços máximos por condição (CSV)"),
          ("Pós-processamento",
           "Notebook em Python: estatísticas, riscos, eventos governantes, envelope crítico e efeito do "
           "carregamento.",
           "Gráficos e tabelas da etapa 5"),
          ("Geometria",
           "Comprimento 3D e ângulo vertical de cada cabo, nos 5 calados e com marés de −0,20 e 1,60 m.",
           "Verificação do ângulo e da carga por cabo")]
for k, (rot, feito, produto) in enumerate(passos):
    xx = X0 + k * (sw + sgap)
    card(sl, xx, Y_CONT, sw, sh)
    texto(sl, xx + 0.2, Y_CONT + 0.12, 0.8, 0.55, str(k + 1), 28, cor, bold=True)
    rotulo(sl, xx + 0.2, Y_CONT + 0.72, sw - 0.4, rot, cor)
    texto(sl, xx + 0.2, Y_CONT + 1.05, sw - 0.4, 1.2, feito, 12.5, TXT2)
    forma(sl, MSO_SHAPE.RECTANGLE, xx + 0.2, Y_CONT + 2.32, sw - 0.4, 0.012, BORDA)
    rotulo(sl, xx + 0.2, Y_CONT + 2.42, sw - 0.4, "Produto", cor, size=9)
    texto(sl, xx + 0.2, Y_CONT + 2.68, sw - 0.4, sh - 2.78, produto, 12.5, TXT, bold=True)
    if k < 3:
        seta(sl, xx + sw + 0.12, Y_CONT + sh / 2 - 0.2, 0.21, 0.4, cor)
ky = Y_CONT + sh + 0.2
kh = Y_FIM - ky
card(sl, X0, ky, LARG, kh)
rotulo(sl, X0 + 0.25, ky + 0.17, 8, "Base teórica e ferramentas", cor)
kw = (LARG - 0.5) / 3
for k, (tit, conteudo, tam) in enumerate([
        ("Forças ambientais (MEG4)", ["F = ½ · C · ρ · V² · A"], 17),
        ("Componentes da tração no cabo", ["T · cos θ (horizontal)", "T · sen θ (vertical)"], 14),
        ("Ferramentas e dados", ["Software de simulação · Python · Google Earth"], 12.5)]):
    xx = X0 + 0.25 + k * kw
    texto(sl, xx, ky + 0.48, kw - 0.3, 0.3, tit, 12.5, TXT, bold=True)
    texto(sl, xx, ky + 0.8, kw - 0.3, kh - 0.9, conteudo, tam, cor if tam > 13 else TXT2,
          bold=tam > 13, entre=2)

# ══════════════════════════════════════════════════════════════
# Etapa 4 · Critérios de avaliação
# ══════════════════════════════════════════════════════════════
sl, cor = base("abordagem", "Etapa 4 · Abordagem", "Critérios de avaliação",
               "Três limites orientam a leitura dos resultados: tração nos cabos, força nos cabeços e ângulo",
               etapa="abordagem")
ch3 = 3.95
criterios = [
    ("Cabos · tração", "≤ 55% MBL", "crit_mbl.png",
     ["Limite de trabalho de 55% do MBL para cabos.",
      "**Neste estudo:** limite de 55%, com MBL de 736 kN (75 t); cabo de Ø 80 mm e EA de 5.632 kN.",
      "**Se exceder:** risco de ruptura do cabo e de snap-back."]),
    ("Cabeços · força", "≤ 100% SWL", "crit_swl.png",
     ["Cabeços em terra dimensionados pelo maior MBL de projeto que o berço recebe.",
      "**Neste estudo:** SWL de 75 t.",
      "**Se exceder:** risco de dano ou de arrancamento do cabeço."]),
    ("Cabos · ângulo vertical", "θ ≤ 25°", "crit_angulo.png",
     ["Até 25° em relação ao plano do cais.",
      "**Neste estudo:** ângulo de cada cabo nos 5 calados, com marés de −0,20 e 1,60 m.",
      "**Se exceder:** a parcela horizontal (T · cos θ) diminui e o cabo segura menos o navio."]),
]
for k, (rot, kpi, img, corpo) in enumerate(criterios):
    xx = X0 + k * (cw + 0.15)
    card(sl, xx, Y_CONT, cw, ch3)
    selo(sl, xx + 0.2, Y_CONT + 0.15, k + 1, cor)
    rotulo(sl, xx + 0.57, Y_CONT + 0.19, cw - 0.8, rot, cor)
    texto(sl, xx + 0.2, Y_CONT + 0.5, cw - 0.4, 0.5, kpi, 27, cor, bold=True)
    figura(sl, img, xx + 0.15, Y_CONT + 1.02, cw - 0.3, 1.2)
    texto(sl, xx + 0.2, Y_CONT + 2.32, cw - 0.4, ch3 - 2.42, corpo, 12, TXT2, entre=4)

# ══════════════════════════════════════════════════════════════
# Etapa 5 · Resumo estatístico (tabela)
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", "Etapa 5 · Resultados · Visão geral", "Resumo estatístico geral",
               "Nenhum elemento passa do limite; os picos crescem com o carregamento",
               f"Fonte: {NOTEBOOK}, tabela “Resumo estatístico geral” (picos de cada elemento nas 5 condições "
               "de carga)", etapa="resultados")
card(sl, **IMG)
EST = D.carregar_estatisticas()
fv = lambda v, d=2: f"{v:.{d}f}".replace(".", ",").replace("-", "−")
ORDEM = list(D.CONDICOES)


def est(grandeza, cond="todos"):
    return EST[(EST["Grandeza"] == grandeza) & (EST["Condicao"] == cond)].iloc[0]


def calado_do_max(grandeza):
    t = EST[(EST["Grandeza"] == grandeza) & (EST["Condicao"].isin(ORDEM))]
    return f"{D.CONDICOES[t.loc[t['Maximo'].idxmax(), 'Condicao']]:g} m"


VARS = [("Tração máx. [tf]", "Tracao max [tf]", "% MBL", D.LIMITE_MBL * D.MBL_TF / 100),
        ("% MBL", "% MBL", "% MBL", D.LIMITE_MBL),
        ("Cabeço máx. [tf]", "Cabeco max [tf]", "% SWL", 75.0),
        ("% SWL", "% SWL", "% SWL", D.LIMITE_SWL)]
lin, dest = [], {}
for i, (nome, g, g_cal, lim) in enumerate(VARS):
    e = est(g)
    vals = [e["Media"], e["Mediana"], e["P90"], e["P99"], e["Maximo"]]
    lin.append([nome] + [fv(v) for v in vals] + [calado_do_max(g_cal)])
    for j, v in enumerate(vals, start=1):
        if v > 0.9 * lim:                          # acima de 90% do limite
            dest[(i, j)] = ALERTA_FILL
tx11, tw11 = X0 + 0.25, IMG["w"] - 0.5
texto(sl, tx11, Y_CONT + 0.2, tw11, 0.3, "Todas as condições de carga", 12.5, TXT, bold=True)
tabela(sl, tx11, Y_CONT + 0.55, tw11, ["Variável", "Média", "Mediana", "P90", "P99", "Máximo", "Calado do máx."],
       lin, [2.3, 1, 1.1, 0.95, 0.95, 1.05, 1.35], alt_linha=0.44, size=12, cor_cabec=cor, destaques=dest)
lin2, dest2 = [], {}
for i, (nome, g, lim) in enumerate([("% MBL máx.", "% MBL", D.LIMITE_MBL), ("% SWL máx.", "% SWL", D.LIMITE_SWL)]):
    vals = [est(g, c)["Maximo"] for c in ORDEM]
    lin2.append([nome] + [fv(v, 1) for v in vals])
    for j, v in enumerate(vals, start=1):
        if v > 0.9 * lim:
            dest2[(i, j)] = ALERTA_FILL
texto(sl, tx11, Y_CONT + 2.95, tw11, 0.3, "Máximo em cada condição de carga", 12.5, TXT, bold=True)
tabela(sl, tx11, Y_CONT + 3.3, tw11,
       ["Variável"] + [f"{D.ROTULO[c]} · {D.CONDICOES[c]:g} m".replace(".", ",") for c in ORDEM],
       lin2, [1.6, 1.15, 1.25, 1.15, 1.25, 1.15], alt_linha=0.4, size=12, cor_cabec=cor, destaques=dest2)
forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, X0 + 0.25, Y_CONT + 4.55, 0.28, 0.2, ALERTA_FILL, BORDA, raio=0.2)
texto(sl, X0 + 0.62, Y_CONT + 4.52, 6.2, 0.3,
      "acima de 90% do limite (55% MBL nos cabos; SWL de 75 t nos cabeços).", 11, TXT3)
colunas_texto(sl, cor,
              ["Pico de cada elemento nas 394 condições, em cada um dos 5 calados: 12 cabos e 6 cabeços por "
               "condição.",
               "**Intuito:** dar a ordem de grandeza dos esforços e ver como mudam com o carregamento; os "
               "percentis descrevem os elementos, não as condições."],
              ["Nenhum elemento passa do limite em nenhuma condição de carga.",
               "Do lastro ao navio cheio, o pico sobe de 42% para 50% MBL nos cabos e de 81,5% para "
               "**98,1% SWL** nos cabeços: com o navio cheio, a margem do cabeço mais carregado é de 2%."],
              "50% MBL", "maior tração, com o navio cheio (14 m)", rot1="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# Distribuição dos esforços (histogramas originais do Colab, quando disponíveis)
# ══════════════════════════════════════════════════════════════
# O maior pico de cada condição fica abaixo de 55% MBL, então nenhum dos 60 picos de cabos passa do limite
assert all(est("% MBL", c)["Maximo"] <= D.LIMITE_MBL for c in ORDEM)
n_cabos_acima = 0
slide_grafico("Visão geral", "Distribuição dos esforços",
              "A mediana e o pico sobem a cada etapa do carregamento; o cabeço mais carregado chega a 98% do SWL",
              "12_picos_por_calado.png",
              ["Para cada condição de carga: mediana, P90 e máximo dos picos dos 12 cabos (% MBL) e dos "
               "6 cabeços (% SWL). A faixa colorida vai da mediana ao máximo.",
               "**Intuito:** verificar se a aproximação do limite é pontual ou acompanha o carregamento."],
              [f"Mediana dos cabos: {fv(est('% MBL', 'lastro')['Mediana'], 1)}% MBL no lastro e "
               f"{fv(est('% MBL', 'cheio')['Mediana'], 1)}% no navio cheio; nos cabeços, "
               f"{fv(est('% SWL', 'lastro')['Mediana'], 1)}% e {fv(est('% SWL', 'cheio')['Mediana'], 1)}% SWL.",
               "O P90 fica perto do máximo: em cada condição, mais de um elemento se aproxima do pico."],
              f"{n_cabos_acima} de 60", "picos de cabos acima de 55% MBL (12 cabos × 5 calados)")

R75 = D.carregar_resumo("75%")
cab75 = R75[R75["Secao"] == "Cabos"]
cbc75 = R75[R75["Secao"] == "Cabecos"]
n_198 = int((cab75["ID"] == 198).sum())
mares = R75["Mare [m]"]
slide_grafico("Fatores de carga", "Influência da maré",
              f"Com 75% de carga, {n_198} dos 12 cabos têm o pico com maré de 0,24 m, no mesmo evento",
              "13_influencia_mare_75.png",
              ["Maior esforço de cada cabo (círculos) e de cada cabeço (triângulos) na condição 75% "
               "(12,75 m); a cor indica o nível de maré no instante desse pico.",
               "**Intuito:** identificar os elementos mais carregados e o nível de maré em que isso ocorre."],
              ["Os cabos 5-B, 6-B, 11-R e 12-R ficam entre 45% e 47% MBL; os demais, entre 20% e 27%.",
               f"O cabeço R chega a 92,9% SWL com a mesma maré. Os picos de todos os elementos ocorrem com "
               f"marés de {fv(mares.min())} a {fv(mares.max())} m: a maré sozinha não define o pior caso."],
              "47% MBL", "pico dos cabos 6-B e 11-R (75%, 12,75 m)", condicao="75% (12,75 m)")

# ══════════════════════════════════════════════════════════════
# Clima ou arranjo (1/2): matriz de correlação
# ══════════════════════════════════════════════════════════════
CA_EVT = D.carregar_ca_eventos()
EVT_CHEIO = D.carregar_eventos_cheio()
e198 = CA_EVT.loc[198]
assert n_198 == 8 and EVT_CHEIO.loc[198, "Cabos"] == 8
ids_def = set(R75.loc[R75["Secao"].isin(["Reacao", "Compressao"]), "ID"]) | \
    set(EVT_CHEIO.index[(EVT_CHEIO["Reacao Defensas"] + EVT_CHEIO["Compressao Defensas"]) > 0])
corr_def = CA_EVT.loc[sorted(ids_def), "V. corrente [nos]"]
sl, cor = base("resultados", "Etapa 5 · Resultados · Fatores de carga", "Clima ou arranjo de cabos?",
               "Um único evento, com corrente de 1,58 nó, gera o pico de 8 dos 12 cabos", etapa="resultados")
orig = por_imagem(sl, "14_eventos_governantes.png", **IMG)
c1h = 4.3
card_texto(sl, COL["x"], Y_CONT + (IMG["h"] - c1h) / 2, COL["w"], c1h, "O que o gráfico conta",
           ["•  Cada barra é um evento do CA (uma hora de vento, corrente, onda e maré); a altura é o número "
            "de elementos que têm o pico nesse evento.",
            f"•  **ID 198** (09/01, 05h): vento de {fv(e198['V. vento [nos]'], 1)} nós, corrente de "
            f"{fv(e198['V. corrente [nos]'])} nó e maré de {fv(e198['Mare [m]'])} m. Gera o pico de 13 elementos, "
            "**8 cabos e 4 cabeços**, com 75% e com o navio cheio.",
            f"•  **Defensas:** todos os picos ocorrem com corrente de {fv(corr_def.min())} a "
            f"{fv(corr_def.max())} nó.",
            "•  A corrente do ID 198 está perto do máximo simulado (1,6 nó); o vento, não (máximo de 50 nós).",
            "•  **Intuito:** separar o efeito do clima do efeito do arranjo."], cor, size=12, num=1)
texto(sl, X0, 7.06, 10.7, 0.3, f"Fonte: {NOTEBOOK}, gráfico “eventos governantes” das condições 75% e cheio "
      "· ambiente de cada evento: CA_394_pontos.csv", 9.5, TXT3)

# ══════════════════════════════════════════════════════════════
# Clima ou arranjo (2/2): % MBL por cabo mais carregado
# ══════════════════════════════════════════════════════════════
comp75 = {n: D.geometria_linha(D.MAPA_CABOS[n], n.split("-")[1], D.CONDICOES["75%"], m)["comp"]
          for n, m in zip(cab75["Elemento"], cab75["Mare [m]"])}
curtos = sorted(comp75, key=comp75.get)[:4]
assert set(curtos) == {"5-B", "6-B", "11-R", "12-R"}
c_curtos = [comp75[n] for n in curtos]
sl, cor = base("resultados", "Etapa 5 · Resultados · Fatores de carga", "Clima ou arranjo de cabos?",
               "No mesmo evento, os quatro cabos mais curtos recebem quase o dobro da tração dos demais",
               etapa="resultados")
orig = por_imagem(sl, "15_tracao_comprimento_75.png", **IMG)
card_texto(sl, COL["x"], Y_CONT, COL["w"], 1.85, "O que o gráfico lê",
           ["Barras: pico de tração de cada cabo (% MBL) na condição 75%. Linha: comprimento 3D do cabo no "
            "instante do pico."], cor, size=12, num=1)
card_texto(sl, COL["x"], Y_CONT + 2.0, COL["w"], IMG["h"] - 2.0, "Resultado obtido",
           [f"•  No ID 198, os quatro cabos mais curtos (5-B, 6-B, 11-R e 12-R, de {fv(min(c_curtos), 1)} a "
            f"{fv(max(c_curtos), 1)} m) ficam em 45–47% MBL; 1-S a 4-A, no mesmo evento, em 25–27%.",
            "•  Com EA de 5.632 kN, a rigidez do cabo é k = EA/L: o cabo mais curto é mais rígido e "
            "absorve mais carga.",
            f"•  A direção também conta: 2-S ({fv(comp75['2-S'], 1)} m) tem quase o comprimento de 11-R e "
            "recebe 25%. O arranjo pesa mais que o clima."], cor, size=12, num=2)
texto(sl, X0, 7.06, 10.7, 0.3, f"Fonte: {NOTEBOOK}, condição 75% (12,75 m) · comprimento: "
      "calcular_geometria_amarracao · gráfico elaborado a partir das saídas do notebook", 9.5, TXT3)

# ══════════════════════════════════════════════════════════════
# Riscos
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", "Etapa 5 · Resultados · Riscos", "Combinações críticas",
               "Nenhuma combinação passa do limite; com 75% de carga, as maiores trações vêm do mesmo evento",
               f"Fonte: {NOTEBOOK}, bloco “Combinações críticas” · condição 75% (12,75 m)", etapa="resultados")
card(sl, **IMG)
tx, tw = X0 + 0.25, IMG["w"] - 0.5
top_c = cab75.sort_values("Valor", ascending=False, kind="stable").head(5)
top_b = cbc75.sort_values("% do Limite", ascending=False, kind="stable").head(5)
assert (top_c["Valor"] <= D.LIMITE_MBL).all() and (cbc75["% do Limite"] <= D.LIMITE_SWL).all()
texto(sl, tx, Y_CONT + 0.2, tw, 0.3, "Cabos · cinco maiores trações (nenhuma acima de 55% MBL)",
      12.5, TXT, bold=True)
tabela(sl, tx, Y_CONT + 0.58, tw,
       ["ID", "Vento [nós]", "Corrente [nós]", "Tração [tf]", "% MBL", "Cabo"],
       [[r["ID"], fv(r["V. vento [nos]"]), fv(r["V. corrente [nos]"]), fv(r["Valor"] / 100 * D.MBL_TF),
         fv(r["Valor"]), r["Elemento"]] for _, r in top_c.iterrows()],
       [0.85, 1.15, 1.35, 1.15, 0.95, 0.75], alt_linha=0.34, cor_cabec=cor,
       destaques={(i, 4): ALERTA_FILL for i, v in enumerate(top_c["Valor"]) if v > 0.8 * D.LIMITE_MBL})
texto(sl, tx, Y_CONT + 2.72, tw, 0.3, "Cabeços · cinco maiores esforços (nenhum acima de 100% SWL)",
      12.5, TXT, bold=True)
tabela(sl, tx, Y_CONT + 3.05, tw,
       ["ID", "Vento [nós]", "Corrente [nós]", "Força [tf]", "% SWL", "Cabeço"],
       [[r["ID"], fv(r["V. vento [nos]"]), fv(r["V. corrente [nos]"]), fv(r["Valor"]),
         fv(r["% do Limite"]), r["Elemento"]] for _, r in top_b.iterrows()],
       [0.85, 1.15, 1.35, 1.15, 0.95, 0.75], alt_linha=0.28, size=11, cor_cabec=cor,
       destaques={(i, 4): ALERTA_FILL for i, v in enumerate(top_b["% do Limite"]) if v > 90})
quatro = cab75.nlargest(4, "Valor")
assert (quatro["ID"] == 198).all() and quatro["Valor"].min() > cab75.drop(quatro.index)["Valor"].max()
colunas_texto(sl, cor,
              ["As cinco maiores trações nos cabos e os cinco maiores esforços nos cabeços na condição 75%, "
               "com o vento e a corrente do evento. Destaque: acima de 80% do limite nos cabos e de 90% nos "
               "cabeços.",
               "**Intuito:** identificar o evento e os elementos mais próximos do limite."],
              ["Os cabeços R e B (92,9% e 91,6% SWL) recebem os quatro cabos mais tracionados. O 5º lugar "
               "empata em 27% entre 3-A e 4-A (ID 198) e 7-G e 8-G (ID 320).",
               "Com o navio cheio, o pico sobe para 50% MBL nos cabos e 98,1% SWL nos cabeços."],
              "45–47% MBL", "as quatro maiores trações, todas no ID 198", rot1="O que a tabela lê")

slide_grafico("Riscos", "Envelope crítico: vento × corrente",
              "Com 75% de carga, 8 dos 12 cabos atingem o pico na mesma condição: 36,7 nós e 1,58 nó",
              "17_envelope_75.png",
              ["Cada bolha é a condição de pico de um elemento, posicionada pela corrente (eixo x) e "
               "pelo vento (eixo y); a cor indica o esforço e o tamanho, a maré.",
               "**Intuito:** mapear em que combinação ambiental cada parte do sistema é mais exigida."],
              ["7-G e 8-G atingem o pico com 49 nós e 1,35 nó; 9-N e 10-N, com 46,1 nós e 1,27 nó.",
               "Os cabeços A, B, R e S também têm o pico no ID 198; o G e o N, junto com os cabos presos a eles."],
              "1,58 nó", "corrente no pico de 8 dos 12 cabos", condicao="75% (12,75 m)")

# ══════════════════════════════════════════════════════════════
# Geometria dos cabos
# ══════════════════════════════════════════════════════════════
f = fv
MB, MA = D.MARE_MIN, D.MARE_MAX                      # −0,20 m e 1,60 m (faixa do CA)
G_BAIXA = D.tabela_geometria(9.0, MB)                 # calado 9 m, maré baixa
G_ALTA = D.tabela_geometria(9.0, MA)                  # calado 9 m, maré alta (maior inclinação)
delta = G_ALTA["ang"] - G_BAIXA["ang"]
dcomp = G_ALTA["comp"] - G_BAIXA["comp"]
i_ang = int(G_ALTA["ang"].idxmax())                   # cabo 5 (buzina 17 → B)
assert G_ALTA.loc[i_ang, "Cabo"] == 5 and G_ALTA["ang"].max() < 25
i_tracao = [5, 10]                                    # cabos 6-B e 11-R: maiores trações (47% MBL, com 75%)
dif_9_10 = G_ALTA.loc[9, "comp"] - G_ALTA.loc[8, "comp"]


def legenda_geo(sl, y):
    for k, (fill, txt) in enumerate([(ALERTA_FILL, "cabos 6 e 11: maior tração (47% MBL, com 75%)"),
                                     (DESTAQUE_FILL, "cabo 5: maior ângulo vertical")]):
        xx = X0 + 0.25 + k * 3.55
        forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, y, 0.28, 0.2, fill, BORDA, raio=0.2)
        texto(sl, xx + 0.37, y - 0.03, 3.15, 0.3, txt, 11, TXT3)


def destaques_geo(col_ang):
    d = {(i, None): ALERTA_FILL for i in i_tracao}
    d[(i_ang, None)] = DESTAQUE_FILL
    d[(i_ang, col_ang)] = DESTAQUE_FILL
    return d


# ── Maré baixa (−0,20 m) ───────────────────────────────────────
sl, cor = base("resultados", "Etapa 5 · Resultados · Geometria", f"Geometria dos cabos · maré de {f(MB)} m",
               f"A maré baixa reduz o ângulo vertical de todos os cabos em {f(delta.min(), 1)}° a "
               f"{f(delta.max(), 1)}°", etapa="resultados")
card(sl, **IMG)
tabela(sl, X0 + 0.25, Y_CONT + 0.18, IMG["w"] - 0.5,
       ["Cabo", "Buzina", "Cabeço", "Comp. 3D [m]", "Âng. vertical [°]", "Δ ângulo [°]"],
       [[g.Cabo, g.Buzina, g.Cabeco, f(g.comp), f(g.ang), f(-d)] for g, d in zip(G_BAIXA.itertuples(), delta)],
       [0.75, 0.85, 0.85, 1.25, 1.4, 1.15], alt_linha=0.31, size=11.5, cor_cabec=cor,
       destaques=destaques_geo(5))
legenda_geo(sl, Y_CONT + 4.42)
colunas_texto(sl, cor,
              [f"Comprimento 3D e ângulo vertical de cada cabo com calado de 9 m e maré baixa ({f(MB)} m). A "
               f"última coluna mostra quanto o ângulo diminui em relação à maré alta ({f(MA)} m).",
               "**Intuito:** isolar o efeito da maré na geometria, com o calado fixo."],
              [f"A buzina desce {f(MA - MB)} m e o comprimento dos cabos muda apenas {f(dcomp.min(), 1)} a "
               f"{f(dcomp.max(), 1)} m.",
               "Com o navio em plena carga (14 m), os cabos ficam quase horizontais; a matriz calado × maré, "
               "adiante, detalha esse efeito."],
              f"−{f(delta.min(), 1)}° a −{f(delta.max(), 1)}°", "variação do ângulo vertical com a maré baixa",
              rot1="O que a tabela lê")

# ── Maré alta (1,60 m · maior inclinação) ──────────────────────
sl, cor = base("resultados", "Etapa 5 · Resultados · Geometria", f"Geometria dos cabos · maré de {f(MA)} m",
               "Pelo cálculo do notebook, todos os cabos ficam abaixo do limite de 25° do MEG4",
               etapa="resultados")
card(sl, **IMG)
tabela(sl, X0 + 0.25, Y_CONT + 0.18, IMG["w"] - 0.5,
       ["Cabo", "Buzina", "Cabeço", "Comp. 3D [m]", "Âng. vertical [°]"],
       [[g.Cabo, g.Buzina, g.Cabeco, f(g.comp), f(g.ang)] for g in G_ALTA.itertuples()],
       [0.8, 0.9, 0.9, 1.4, 1.5], alt_linha=0.31, size=11.5, cor_cabec=cor, destaques=destaques_geo(4))
legenda_geo(sl, Y_CONT + 4.42)
colunas_texto(sl, cor,
              ["Comprimento 3D e ângulo vertical de cada cabo, da buzina ao cabeço, no cenário de maior "
               f"inclinação (calado de 9 m e maré de {f(MA)} m).",
               "**Intuito:** verificar o limite de 25° e relacionar o esforço de cada cabo à sua geometria."],
              [f"Com os cabos 11 e 12 no cabeço R, eles medem {f(G_ALTA.loc[10, 'comp'], 1)} e "
               f"{f(G_ALTA.loc[11, 'comp'], 1)} m e ficam entre os mais inclinados (16° a 17°).",
               f"O cabo 9 é {f(dif_9_10, 1)} m mais curto que o cabo 10, que divide o cabeço N. Esses valores "
               "dependem da revisão da geometria (próximos passos)."],
              f"{f(G_ALTA['ang'].max(), 1)}° < 25°", "maior ângulo vertical (cabo 5) e limite do MEG4",
              rot1="O que a tabela lê")

# ── NOVO · Matriz relacional calado × maré ─────────────────────
# Valores conferidos com calcular_geometria_amarracao (notebook TF4.0) para os 12 cabos.
REL = {(9, MB): "ponto médio de altura", (9, MA): "maior ponto de altura",
       (14, MB): "menor ponto de altura", (14, MA): "ponto médio de altura"}
MATRIZ = {}  # (calado, maré): (θ mín, θ máx, buzina − cabeço [m], cos θ máx, altura relativa no TF I)
for (cal, mar), rel in REL.items():
    g = D.tabela_geometria(cal, mar)
    MATRIZ[(cal, mar)] = (g["ang"].min(), g["ang"].max(), g["dz_buzina_cabeco"].iloc[0],
                          np.cos(np.radians(g["ang"].max())), rel)
G14_B, G14_A = D.tabela_geometria(14.0, MB), D.tabela_geometria(14.0, MA)
d_mare = pd.concat([G_ALTA["ang"] - G_BAIXA["ang"], G14_A["ang"] - G14_B["ang"]])
d_mare_comp = pd.concat([G_ALTA["comp"] - G_BAIXA["comp"], G14_A["comp"] - G14_B["comp"]])
d_cal = pd.concat([G14_B["ang"] - G_BAIXA["ang"], G14_A["ang"] - G_ALTA["ang"]])
d_cal_comp = pd.concat([G_BAIXA["comp"] - G14_B["comp"], G_ALTA["comp"] - G14_A["comp"]])
dif_9_10_todos = [D.tabela_geometria(c, m)["comp"].iloc[9] - D.tabela_geometria(c, m)["comp"].iloc[8]
                  for c in D.CONDICOES.values() for m in (MB, MA)]
cmin_pior = MATRIZ[(9, MA)][3]
sl, cor = base("resultados", "Etapa 5 · Resultados · Geometria", "Calado × maré no ângulo dos cabos",
               f"O calado pesa mais que a maré: de 9 m para 14 m, a buzina desce 5 m; a maré a move {f(MA - MB)} m",
               f"Fonte: {NOTEBOOK}, função calcular_geometria_amarracao · altura relativa: SILVA (2026a, "
               "Tabela 4)", etapa="resultados")
mx, my = X0, Y_CONT
hw, hh_ = 1.6, 0.5                     # coluna de rótulos das linhas e altura do cabeçalho
cw2 = (IMG["w"] - hw - 0.2) / 2
ch2 = 1.55
for j, (tit, sub) in enumerate([("Maré baixa", f"{f(MB)} m"), ("Maré alta", f"{f(MA)} m")]):
    xx = mx + hw + 0.1 + j * (cw2 + 0.1)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, my, cw2, hh_, cor, raio=0.2)
    texto(sl, xx, my, cw2, hh_, f"**{tit}** · {sub}", 13, "#FFFFFF", align=PP_ALIGN.CENTER,
          ancora=MSO_ANCHOR.MIDDLE, linha=1.0)
for i, (cal, tit, sub) in enumerate([(9, "Lastro", "calado 9 m"), (14, "Plena carga", "calado 14 m")]):
    yy = my + hh_ + 0.1 + i * (ch2 + 0.1)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, mx, yy, hw, ch2, tint(cor, 0.30), raio=0.08)
    texto(sl, mx + 0.1, yy, hw - 0.2, ch2, [f"**{tit}**", sub], 13.5, TXT, align=PP_ALIGN.CENTER,
          ancora=MSO_ANCHOR.MIDDLE, entre=2)
    for j, mar in enumerate((MB, MA)):
        tmin, tmax, dz, cmin, rel = MATRIZ[(cal, mar)]
        xx = mx + hw + 0.1 + j * (cw2 + 0.1)
        intens = {(9, MA): 0.34, (9, MB): 0.22, (14, MA): 0.10, (14, MB): 0.04}[(cal, mar)]
        forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, yy, cw2, ch2, tint(cor, intens), BORDA, raio=0.06)
        texto(sl, xx + 0.15, yy + 0.12, cw2 - 0.3, 0.25, rel.upper(), 9, TXT3, bold=True, spc=80)
        texto(sl, xx + 0.15, yy + 0.4, cw2 - 0.3, 0.5, f"{f(tmin, 1)}° a {f(tmax, 1)}°", 22, TXT, bold=True)
        lado = "acima" if dz > 0 else "abaixo"
        horiz = "≈ 100%" if cmin > 0.9995 else f"≥ {f(cmin * 100, 1)}%"
        texto(sl, xx + 0.15, yy + 0.95, cw2 - 0.3, 0.55,
              [f"buzina {f(abs(dz), 1)} m {lado} do cabeço",
               f"parcela horizontal {horiz}"], 11.5, TXT2, entre=1)
# efeitos de cada variável, abaixo da matriz
ey = my + hh_ + 0.1 + 2 * (ch2 + 0.1) + 0.02
eh = Y_FIM - ey
for j, (tit, corpo) in enumerate([
        (f"→ Maré {f(MB)} → {f(MA)} m",
         f"ângulo +{f(d_mare.min(), 1)}° a +{f(d_mare.max(), 1)}°; cabos até {f(d_mare_comp.max(), 1)} m mais longos"),
        ("↓ Calado 9 → 14 m",
         f"ângulo −{f(-d_cal.max(), 1)}° a −{f(-d_cal.min(), 1)}°; cabos {f(d_cal_comp.min(), 1)} a "
         f"{f(d_cal_comp.max(), 1)} m mais curtos")]):
    xx = mx + j * (IMG["w"] / 2 + 0.05)
    ww = IMG["w"] / 2 - 0.05
    card(sl, xx, ey, ww, eh)
    texto(sl, xx + 0.15, ey + 0.08, ww - 0.3, eh - 0.12, [f"**{tit}**", corpo], 11.5, TXT2, entre=1,
          ancora=MSO_ANCHOR.MIDDLE)
card_texto(sl, COL["x"], Y_CONT, COL["w"], 1.8, "O que a matriz relaciona",
           [f"Cruza o calado (lastro e plena carga) com a maré ({f(MB)} e {f(MA)} m). Cada célula traz a faixa "
            "de ângulo vertical dos 12 cabos e a altura da buzina em relação ao cabeço.",
            "**Intuito:** separar o efeito do calado do efeito da maré."], cor, size=12, num=1)
card_texto(sl, COL["x"], Y_CONT + 1.95, COL["w"], IMG["h"] - 1.95, "O que isso significa",
           [f"•  Todos os casos ficam abaixo de 25°: no pior ({f(MATRIZ[(9, MA)][1], 1)}°), "
            f"{cmin_pior * 100:.0f}% da tração age na horizontal.",
            "•  Com a buzina acima do cabeço, subir o navio alonga os cabos; carregar os encurta.",
            f"•  A diferença de {f(min(dif_9_10_todos), 1)} a {f(max(dif_9_10_todos), 1)} m entre os cabos 9 e "
            "10 se mantém em todos os calados."], cor,
           kpi=f"5 m × {f(MA - MB)} m", kpi_legenda="deslocamento da buzina: calado × maré", size=12, num=2)

# ══════════════════════════════════════════════════════════════
# NOVO · Efeito do carregamento (síntese dos 5 calados)
# ══════════════════════════════════════════════════════════════
mbl_l, mbl_c = est("% MBL", "lastro")["Maximo"], est("% MBL", "cheio")["Maximo"]
swl_l, swl_c = est("% SWL", "lastro")["Maximo"], est("% SWL", "cheio")["Maximo"]
ang_l, ang_c = MATRIZ[(9, MA)][1], MATRIZ[(14, MA)][1]
area_sub = (14.0 / 9.0 - 1) * 100                     # área lateral submersa ∝ calado
borda_livre = (1 - (18.5 - 14.0) / (18.5 - 9.0)) * 100  # área exposta ao vento ∝ pontal − calado
slide_grafico("Carregamento", "Efeito do carregamento",
              "Ao carregar, os cabos ficam mais horizontais, mas os esforços sobem",
              "22_efeito_carregamento.png",
              ["Em cima: maior esforço de cabos e cabeços em relação ao limite, do lastro ao navio cheio. "
               "Embaixo: maior ângulo vertical entre os 12 cabos, nas marés baixa e alta.",
               "**Intuito:** separar o efeito da geometria do efeito das forças ambientais."],
              [f"Cabeços: de {f(swl_l, 1)}% para {f(swl_c, 1)}% SWL. O ângulo máximo cai de {f(ang_l, 1)}° "
               f"para {f(ang_c, 1)}°: o aumento não vem da geometria.",
               f"A área submersa cresce {area_sub:.0f}% e a exposta ao vento cai {borda_livre:.0f}%: a corrente "
               "tende a dominar (OCIMF, 2018)."],
              f"+{f(mbl_c - mbl_l, 0)} pontos", f"de {f(mbl_l, 0)}% para {f(mbl_c, 0)}% MBL, de 9 m para 14 m")

# ══════════════════════════════════════════════════════════════
# Etapa 6 · Conclusões parciais
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", "Etapa 6 · Conclusões", "Conclusões parciais",
               "Os resultados respondem ao objetivo desta etapa",
               "Fontes: SILVA (2026a; 2026b)", etapa="conclusoes")
achados = [
    ("C1", "Nenhum elemento passa do limite",
     f"Pico de {f(mbl_c, 0)}% MBL nos cabos e {f(swl_c, 1)}% SWL nos cabeços, ambos com o navio cheio (14 m)."),
    ("C2", "O carregamento aumenta os esforços",
     f"De 9 m para 14 m: +{f(mbl_c - mbl_l, 0)} pontos de % MBL e +{f(swl_c - swl_l, 1)} pontos de % SWL, "
     "com a corrente agindo numa área submersa maior."),
    ("C3", "O arranjo pesa mais que o clima",
     "Um único evento (ID 198) gera o pico de 8 dos 12 cabos; nele, os 4 cabos mais curtos recebem quase "
     "o dobro da tração."),
    ("C4", "A geometria melhora ao carregar",
     f"O ângulo máximo cai de {f(ang_l, 1)}° (lastro, maré alta) para {f(ang_c, 1)}° (cheio); todos os casos "
     "ficam abaixo de 25°."),
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
      "A simulação mostra que o berço fica mais perto do limite com o navio cheio, quando os cabeços chegam "
      "a 98% do SWL. Os lançantes ligados aos cabeços B e R, os mais curtos e mais tracionados, são os pontos "
      "que o monitoramento por visão computacional deve acompanhar prioritariamente.", 14, TXT2)

# ══════════════════════════════════════════════════════════════
# Alternativas para aumentar a segurança operacional
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", "Etapa 6 · Conclusões", "Alternativas para aumentar a segurança",
               "Medidas voltadas aos pontos mais carregados: lançantes curtos, cabeços B e R e navio cheio",
               "Fontes: OCIMF (2018, seções 1.3 e 1.4; Figuras 1.3 e 1.4); PIANC (2025, cap. 4); "
               "SILVA (2026a, seção 2.2.1)", etapa="conclusoes")
POR_QUE = [
    "os 4 cabos mais curtos chegam a 45–47% MBL; os demais, a 25–27%.",
    f"com o navio cheio, o cabeço mais carregado chega a {f(swl_c, 1)}% SWL.",
    "os esforços sobem a cada etapa do carregamento.",
    "a tração não é medida, e a margem nos cabeços cai para 2%.",
]
alternativas = [
    ("Arranjo dos cabos", "C3",
     ["•  Igualar a rigidez dos lançantes: aproximar os comprimentos ou usar caudas (tails) nos cabos "
      "mais curtos (5, 6, 11 e 12).",
      "•  Reduzir a diferença de comprimento entre os cabos 9 e 10, que dividem o cabeço N.",
      "•  Manter um arranjo equilibrado entre proa e popa (OCIMF, 2018, seção 1.4.1)."]),
    ("Capacidade do berço", "C1 · C2",
     ["•  Verificar a capacidade dos cabeços B e R, que recebem os cabos mais tracionados: nos exemplos do "
      "MEG4, um único gancho ou cabeço tem SWL de 100% a 118% do MBL (OCIMF, 2018, Figuras 1.3 e 1.4).",
      "•  Avaliar cabeços de maior capacidade ou adicionais (PIANC, 2025, seção 4.9)."]),
    ("Gestão da operação", "C2 · C4",
     ["•  Ajustar a tensão dos cabos ao longo da maré e do carregamento (OCIMF, 2018, seção 1.3).",
      "•  Definir limites de paragem com corrente, vento, direção, maré e calado.",
      "•  Retirar os cabos de uso quando a resistência residual chegar a 75% do MBL de projeto "
      "(OCIMF, 2018, Figuras 1.3 e 1.4)."]),
    ("Monitoramento contínuo", "C1 · C3",
     ["•  Medir a carga em tempo real com cabeços instrumentados (smart bollards) ou ganchos com "
      "sensor de carga (PIANC, 2025, seção 4.10.6).",
      "•  Aplicar a visão computacional proposta no TF, começando pelos lançantes dos cabeços B e R."]),
]
aw = (LARG - 0.15) / 2
ah = (Y_FIM - Y_CONT - 0.15) / 2
for k, (rot, chip, corpo) in enumerate(alternativas):
    xx = X0 + (k % 2) * (aw + 0.15)
    yy = Y_CONT + (k // 2) * (ah + 0.15)
    card(sl, xx, yy, aw, ah)
    selo(sl, xx + 0.2, yy + 0.15, k + 1, cor)
    rotulo(sl, xx + 0.57, yy + 0.19, aw - 2.4, rot, cor, size=10.5)
    cx = xx + aw - 1.75
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, cx, yy + 0.14, 1.55, 0.3, tint(cor, 0.16), raio=0.5)
    texto(sl, cx, yy + 0.14, 1.55, 0.3, f"responde a {chip}", 10, TXT2, bold=True,
          align=PP_ALIGN.CENTER, ancora=MSO_ANCHOR.MIDDLE, linha=1.0)
    texto(sl, xx + 0.25, yy + 0.56, aw - 0.45, ah - 1.11, corpo, 12, TXT2, entre=4)
    forma(sl, MSO_SHAPE.RECTANGLE, xx + 0.25, yy + ah - 0.52, aw - 0.5, 0.012, BORDA)
    texto(sl, xx + 0.25, yy + ah - 0.45, aw - 0.45, 0.38, f"**Por quê:** {POR_QUE[k]}", 11.5, TXT3,
          italic=True)

# ══════════════════════════════════════════════════════════════
# Próximos passos
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", "Etapa 6 · Conclusões", "Próximos passos",
               "Do diagnóstico parcial à entrega final",
               "Próximos passos propostos a partir dos resultados desta etapa e do plano de SILVA (2026a)",
               etapa="conclusoes")
etapas = [("TF I · concluído",
           "Base teórica: equipamentos, defensas, cabos PPL/PE, forças ambientais (MEG4) e caracterização "
           "do T-Grão."),
          ("Etapa atual",
           "Simulação de 394 condições em 5 condições de carga (9 a 14 m): esforços, riscos, geometria e "
           "efeito do carregamento."),
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
passos_a = ["•  Exportar os resultados caso a caso (394 × 5 condições) para refazer histogramas, correlações "
            "e a matriz de operabilidade",
            f"•  Reavaliar os resultados com 50% do MBL: com o navio cheio, o pico já chega a {f(mbl_c, 0)}%",
            "•  Avaliar o efeito de navios passantes no canal"]
passos_b = ["•  Revisar o cálculo de geometria (lado das buzinas e posição dos cabeços) e validá-lo com o berço real",
            "•  Simular as alternativas de arranjo e de capacidade para medir o ganho de segurança",
            "•  **Desenvolver o monitoramento dos cabos por visão computacional e machine learning**"]
bw = (LARG - 0.75) / 2
texto(sl, X0 + 0.25, ly + 0.55, bw, Y_FIM - ly - 0.65, passos_a, 12.5, TXT2, entre=7)
texto(sl, X0 + 0.5 + bw, ly + 0.55, bw, Y_FIM - ly - 0.65, passos_b, 12.5, TXT2, entre=7)

# ══════════════════════════════════════════════════════════════
# Referências (ABNT NBR 6023)
# ══════════════════════════════════════════════════════════════
sl, cor = base("referencias", "Apresentação", "Referências",
               "Fontes citadas nos slides, conforme a ABNT NBR 6023")
REFS_A = [
    "ACP – AUTORIDAD DEL CANAL DE PANAMÁ. **OP Notice to Shipping N-1-2026**. Panamá, 2026.",
    "BRASIL. Ministério de Portos e Aeroportos. **Porto de Santos movimenta 186,4 milhões de toneladas em "
    "2025 e avança para ampliar sua capacidade**. Brasília, 16 jan. 2026. Disponível em: https://www.gov.br/"
    "portos-e-aeroportos. Acesso em: 13 nov. 2026.",
    "MARINHA DO BRASIL. Diretoria de Hidrografia e Navegação. **Tábuas das marés 2025**: Porto de Santos "
    "(SP). Niterói: DHN, 2025.",
    "OCIMF – OIL COMPANIES INTERNATIONAL MARINE FORUM. **Mooring Equipment Guidelines (MEG4)**. 4. ed. "
    "Londres: Witherby Publishing Group, 2018.",
    "PIANC. **Design principles for dry bulk marine terminals**. Bruxelas: PIANC, 2017.",
    "PIANC. **Mooring of large ships at quay walls**. Bruxelas: PIANC, 2025. (MarCom Working Group "
    "Report, n. 186).",
]
REFS_B = [
    "SILVA, G. G. **Sistema de amarração em conjuntos portuários**. 2026a. Trabalho de Formatura I "
    "(PNV3510) – Escola Politécnica, Universidade de São Paulo, São Paulo, 2026.",
    "SILVA, G. G. **TF4.0**: análise dos resultados de simulação de amarração (5 condições de carga, "
    "394 condições ambientais). São Paulo, 2026b. Notebook Python.",
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
# Encerramento
# ══════════════════════════════════════════════════════════════
sl = prs.slides.add_slide(prs.slide_layouts[6])
capa_fundo(sl)
texto(sl, X0, 2.25, 11, 1.1, "Obrigado pela atenção", 50, TXT, bold=True)
texto(sl, X0, 3.65, 11, 1.1, "Perguntas?", 50, TXT, bold=True)
texto(sl, X0 + 0.05, 5.0, 8, 0.4, "Gustavo Guilherme Silva · NUSP 11820555", 15, TXT2)
numero(sl)

assert len(prs.slides) == TOTAL, len(prs.slides)
prs.save(SAIDA)
print("Salvo:", SAIDA, "·", len(prs.slides), "slides")
