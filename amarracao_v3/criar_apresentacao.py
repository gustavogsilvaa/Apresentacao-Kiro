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
A numeração dos slides é automática.

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
TOTAL = 26
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
    for pasta in ("imagens_originais", "imagens_recriadas"):
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


def slide_grafico(sub, titulo, mensagem, nome_img, le, resultado, kpi, kpi_leg, rodape=True):
    sl, cor = base("resultados", f"Etapa 5 · Resultados · {sub}", titulo, mensagem, etapa="resultados")
    original = por_imagem(sl, nome_img, **IMG)
    colunas_texto(sl, cor, le, resultado, kpi, kpi_leg)
    if rodape:
        texto(sl, X0, 7.06, 10.7, 0.3, nota_imagem(original), 9.5, TXT3)
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
         ("cenario", "Cenário analisado", "Terminal T-Grão, maré de Santos, layout típico do Panamax e arranjo simulado"),
         ("abordagem", "Abordagem", "Etapas do trabalho e critérios de avaliação"),
         ("resultados", "Resultados", "Leitura dos gráficos da simulação: esforços, riscos e geometria"),
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
for k, (v, leg) in enumerate([(">100 Mt", "de carga por ano no Porto de Santos."),
                              ("14 de 60", "berços dedicados a granéis vegetais."),
                              ("US$ 34 mi", "em sinistros com equipamentos de amarração em 20 anos."),
                              ("até 40%", "de aumento da tração nos cabos com a passagem de outro navio.")]):
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
for k, (verbo, compl) in enumerate([("Quantificar", "os esforços nas 394 condições simuladas."),
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
          "•  Calados analisados: 9 m (lastro) e 14 m (plena carga)"])]):
    card_texto(sl, x_top + k * (cw + 0.15), Y_CONT, cw, rh, rot, corpo, cor, kpi=kpi, kpi_legenda=leg,
               size=12.5)
yb = Y_CONT + rh + 0.15
for k, (rot, kpi, leg, corpo) in enumerate([
        ("Ambiente · maré", "−0,16 a 1,52 m", "faixa de maré prevista para Santos em 2025",
         ["Águas abrigadas: as ondas chegam atenuadas; vento e corrente de maré predominam."]),
        ("Ambiente · vento", "50 nós", "máximo simulado, com direções variáveis",
         ["Inclui vento forte de propósito, para avaliar o efeito da força aerodinâmica no navio e nos "
          "cabos. De través, essa força é muito maior que de proa (OCIMF, 2018)."]),
        ("Ambiente · corrente", "1,6 nó", "corrente máxima simulada, em águas abrigadas",
         ["Enchente e vazante, com desalinhamento de 10° em relação ao eixo do navio: mesmo um ângulo "
          "pequeno gera força transversal relevante (OCIMF, 2018)."])]):
    card_texto(sl, X0 + k * (cw + 0.15), yb, cw, rh, rot, corpo, cor, kpi=kpi, kpi_legenda=leg,
               size=12.5)

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
            "A geometria usou 1,58 m como média de maré alta."], cor,
           kpi="1,58 m", kpi_legenda="maré alta usada no cálculo da geometria", size=12.5, num=2)

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
           "394 combinações de vento, corrente, maré e direções, com calado de 9 m e 14 m.",
           "Arquivos CSV com esforços e condições"),
          ("Pós-processamento",
           "Notebook em Python: estatísticas, filtro de riscos, envelope crítico e matriz de operabilidade.",
           "Gráficos e tabelas da etapa 5"),
          ("Geometria",
           "Comprimento 3D e ângulo vertical de cada cabo, com calados de 9 e 14 m e marés de 0 e 1,58 m.",
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
      "**Neste estudo:** limite de 55%, com MBL de 75 t.",
      "**Se exceder:** risco de ruptura do cabo e de snap-back."]),
    ("Cabeços · força", "≤ 100% SWL", "crit_swl.png",
     ["Cabeços em terra dimensionados pelo maior MBL de projeto que o berço recebe.",
      "**Neste estudo:** SWL de 75 t.",
      "**Se exceder:** risco de dano ou de arrancamento do cabeço."]),
    ("Cabos · ângulo vertical", "θ ≤ 25°", "crit_angulo.png",
     ["Até 25° em relação ao plano do cais.",
      "**Neste estudo:** ângulo de cada cabo para calados de 9 e 14 m e marés de 0 e 1,58 m.",
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
               "A mediana fica em 31,5% MBL, mas o percentil 90 já ultrapassa o limite de 55%",
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
      "acima do limite (55% MBL nos cabos; 100% SWL nos cabeços).", 11, TXT3)
colunas_texto(sl, cor,
              ["Tração nos cabos (tf e % MBL), reação e compressão nas defensas e força nos cabeços "
               "(tf e % SWL).",
               "**Intuito:** dar a ordem de grandeza dos esforços; os percentis P90 a P99 representam as "
               "condições mais severas."],
              ["Metade das condições fica abaixo de 31,5% MBL, mas o P90 já ultrapassa o limite (59,6%) "
               "e o máximo chega a 75,3%.",
               "Nos cabeços, só o máximo (110,8%) passa de 100% SWL: são **duas condições com risco de "
               "arrancamento**, ambas no cabeço S."],
              "31,5% MBL", "mediana da tração nos cabos", rot1="O que a tabela lê")

# ══════════════════════════════════════════════════════════════
# Distribuição dos esforços (histogramas originais do Colab, quando disponíveis)
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", "Etapa 5 · Resultados · Visão geral", "Distribuição dos esforços",
               "58 das 394 condições (14,7%) passam do limite nos cabos; nos cabeços, apenas 2",
               etapa="resultados")
if existe_original("01a_hist_mbl.png") and existe_original("01b_hist_swl.png"):
    card(sl, **IMG, fill="#FFFFFF")
    meia = (IMG["h"] - 0.3) / 2
    figura(sl, "01a_hist_mbl.png", IMG["x"] + 0.12, IMG["y"] + 0.1, IMG["w"] - 0.24, meia)
    figura(sl, "01b_hist_swl.png", IMG["x"] + 0.12, IMG["y"] + 0.2 + meia, IMG["w"] - 0.24, meia)
else:
    por_imagem(sl, "01_histogramas.png", **IMG)
colunas_texto(sl, cor,
              ["Número de condições em cada faixa de carregamento: % MBL nos cabos e % SWL nos "
               "cabeços. A linha tracejada marca o limite.",
               "**Intuito:** verificar se as violações são casos isolados ou frequentes."],
              ["Nos cabos, a faixa mais frequente é 10–16% MBL (92 condições), mas a distribuição se "
               "estende até 75%.",
               "Nos cabeços, a faixa mais frequente é 30–37% SWL. O risco é recorrente nos cabos e "
               "pontual nos cabeços."],
              "58 de 394", "condições acima de 55% MBL (14,7%)")

slide_grafico("Fatores de carga", "Influência da maré",
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
              "75,3% MBL", "pico do cabo 9, ligado ao cabeço N", rodape=False)

# ══════════════════════════════════════════════════════════════
# Clima ou arranjo (1/2): matriz de correlação
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", "Etapa 5 · Resultados · Fatores de carga", "Clima ou arranjo de cabos?",
               "Nenhuma variável ambiental isolada explica a tração nos cabos", etapa="resultados")
orig = por_imagem(sl, "04_correlacao_matriz.png", **IMG)
c1h = 3.75
card_texto(sl, COL["x"], Y_CONT + (IMG["h"] - c1h) / 2, COL["w"], c1h, "O que o gráfico conta",
           ["•  Cada célula é a correlação linear (r) entre duas variáveis: perto de ±1, forte; perto de 0, "
            "fraca.",
            "•  **Ambiente × tração (% MBL):** todos os valores ficam abaixo de 0,1 em módulo.",
            "•  **Ambiente × cabeços (% SWL):** vento (0,25) e corrente (0,20) pesam um pouco mais, mas a "
            "relação ainda é fraca.",
            "•  **% MBL × % SWL = 0,70:** os mesmos eventos que tracionam os cabos carregam os cabeços.",
            "•  **Hs × Tp = −0,51:** relação entre duas entradas da simulação, não um efeito no sistema.",
            "•  **Intuito:** separar o efeito do clima do efeito do arranjo."], cor, size=12, num=1)
texto(sl, X0, 7.06, 10.7, 0.3, nota_imagem(orig) + " · direções não entram na correlação", 9.5, TXT3)

# ══════════════════════════════════════════════════════════════
# Clima ou arranjo (2/2): % MBL por cabo mais carregado
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", "Etapa 5 · Resultados · Fatores de carga", "Clima ou arranjo de cabos?",
               "Nenhuma variável ambiental isolada explica a tração; a sobrecarga se concentra no cabo 9",
               etapa="resultados")
orig = por_imagem(sl, "04c_boxplot_cabos_alto.png", **IMG)
card_texto(sl, COL["x"], Y_CONT, COL["w"], 1.85, "O que o gráfico lê",
           ["Cada caixa reúne as condições em que aquele cabo foi o mais carregado: a linha central é a "
            "mediana e as hastes vão até os extremos."], cor, size=12, num=1)
card_texto(sl, COL["x"], Y_CONT + 2.0, COL["w"], IMG["h"] - 2.0, "Resultado obtido",
           ["•  Nenhuma variável explica sozinha a tração; nos cabeços, vento (0,25) e corrente (0,20) "
            "pesam um pouco mais.",
            "•  Só o cabo 9 passa de 55%, o que aponta para o arranjo dos cabos."], cor, size=12.5, num=2)
texto(sl, X0, 7.06, 10.7, 0.3, nota_imagem(orig), 9.5, TXT3)

# ══════════════════════════════════════════════════════════════
# Riscos
# ══════════════════════════════════════════════════════════════
sl, cor = base("resultados", "Etapa 5 · Resultados · Riscos", "Combinações críticas",
               "O cabo 9 é o elemento crítico nas cinco piores condições; o cabeço S excede o SWL com vento forte",
               f"Fonte: {NOTEBOOK}, bloco “Filtro de riscos e piores combinações”", etapa="resultados")
card(sl, **IMG)
tx, tw = X0 + 0.25, IMG["w"] - 0.5
texto(sl, tx, Y_CONT + 0.2, tw, 0.3, "Cabos · cinco piores condições (de 58 acima de 55% MBL)",
      12.5, TXT, bold=True)
tabela(sl, tx, Y_CONT + 0.58, tw,
       ["Cond.", "Vento [nós]", "Corrente [nós]", "Tração [tf]", "% MBL", "Cabo"],
       [["259", "8,20", "1,57", "56,44", "75,30", "9"],
        ["310", "41,40", "1,16", "54,72", "73,00", "9"],
        ["78", "5,30", "1,43", "53,95", "71,90", "9"],
        ["282", "29,40", "1,58", "51,82", "69,10", "9"],
        ["272", "39,90", "0,47", "51,85", "69,10", "9"]],
       [0.85, 1.15, 1.35, 1.15, 0.95, 0.75], alt_linha=0.37, cor_cabec=cor,
       destaques={(i, 4): ALERTA_FILL for i in range(5)})
texto(sl, tx, Y_CONT + 2.98, tw, 0.3, "Cabeços · todas as condições acima de 100% SWL",
      12.5, TXT, bold=True)
tabela(sl, tx, Y_CONT + 3.36, tw,
       ["Cond.", "Vento [nós]", "Corrente [nós]", "Força [tf]", "% SWL", "Cabeço"],
       [["198", "36,70", "1,58", "83,13", "110,80", "S"],
        ["55", "49,10", "0,88", "78,34", "104,50", "S"]],
       [0.85, 1.15, 1.35, 1.15, 0.95, 0.75], alt_linha=0.37, cor_cabec=cor,
       destaques={(i, 4): ALERTA_FILL for i in range(2)})
colunas_texto(sl, cor,
              ["As cinco condições que mais carregam os cabos e todas as que ultrapassam o limite nos "
               "cabeços, com vento e corrente.",
               "**Intuito:** identificar as combinações ambientais que levam ao limite e o elemento que "
               "falha primeiro."],
              ["O pior caso (condição 259) ocorre com vento de apenas 8,2 nós, mas com corrente de "
               "1,57 nó.",
               "As duas violações do cabeço S ocorrem com vento forte (36,7 e 49,1 nós)."],
              "5 de 5", "piores condições têm o cabo 9 como crítico", rot1="O que a tabela lê")

slide_grafico("Riscos", "Envelope crítico: vento × corrente",
              "A maioria dos cabos atinge o pico com vento forte; o cabo 9, com corrente forte e maré alta",
              "05_envelope_critico.png",
              ["Cada bolha é a condição de pico de um elemento, posicionada pela corrente (eixo x) e "
               "pelo vento (eixo y); a cor indica o esforço e o tamanho, a maré.",
               "**Intuito:** mapear em que combinação ambiental cada parte do sistema é mais exigida."],
              ["Os demais cabos atingem o pico com vento de 37 a 49 nós.",
               "Os cabeços A, B e S têm o pico na mesma condição (36,7 nós, maré baixa); o cabeço N, "
               "junto com o cabo 9, que está preso a ele."],
              "8,2 nós", "vento na condição de pico do cabo 9")

slide_grafico("Riscos", "Matriz de operabilidade",
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
# Geometria dos cabos
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


# ── Maré de 0 m (base) ─────────────────────────────────────────
sl, cor = base("resultados", "Etapa 5 · Resultados · Geometria", "Geometria dos cabos · maré de 0 m",
               "A maré baixa reduz o ângulo vertical de todos os cabos em 2,2° a 4,9°", etapa="resultados")
card(sl, **IMG)
tabela(sl, X0 + 0.25, Y_CONT + 0.18, IMG["w"] - 0.5,
       ["Cabo", "Buzina", "Cabeço", "Comp. 3D [m]", "Âng. vertical [°]", "Δ ângulo [°]"],
       [[g[0], g[1], g[2], f(g[5]), f(g[6]), "−" + f(g[4] - g[6])] for g in GEO],
       [0.75, 0.85, 0.85, 1.25, 1.4, 1.15], alt_linha=0.31, size=11.5, cor_cabec=cor,
       destaques={(8, None): ALERTA_FILL, (4, None): DESTAQUE_FILL, (4, 5): DESTAQUE_FILL})
legenda_geo(sl, Y_CONT + 4.42)
colunas_texto(sl, cor,
              ["Comprimento 3D e ângulo vertical de cada cabo com calado de 9 m e maré baixa (0 m). A última "
               "coluna mostra quanto o ângulo diminui em relação à maré alta (1,58 m).",
               "**Intuito:** isolar o efeito da maré na geometria, com o calado fixo."],
              ["A buzina desce 1,58 m e o comprimento dos cabos muda apenas 0,2 a 0,5 m.",
               "Com o navio em plena carga (14 m), os cabos ficam quase horizontais; a matriz calado × maré, "
               "adiante, detalha esse efeito."],
              "−2,2° a −4,9°", "variação do ângulo vertical com a maré baixa", rot1="O que a tabela lê")

# ── Maré de 1,58 m (maior inclinação) ──────────────────────────
sl, cor = base("resultados", "Etapa 5 · Resultados · Geometria", "Geometria dos cabos · maré de 1,58 m",
               "Pelo cálculo do notebook, todos os cabos ficam abaixo do limite de 25° do MEG4",
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

# ── NOVO · Matriz relacional calado × maré ─────────────────────
# Valores conferidos com calcular_geometria_amarracao (notebook TF4.0) para os 12 cabos.
MATRIZ = {  # (calado, maré): (θ mín, θ máx, buzina − cabeço [m], cos θ máx, altura relativa no TF I)
    # θ com 3 casas: com 2 casas, 0,35 e 2,05 seriam arredondados para baixo (0,3° e 2,0°).
    (9, 0): (6.774, 16.672, 4.90, 0.958, "ponto médio de altura"),
    (9, 1.58): (8.928, 21.606, 6.48, 0.930, "maior ponto de altura"),
    (14, 0): (0.139, 0.350 + 2e-4, -0.10, 1.000, "menor ponto de altura"),
    (14, 1.58): (2.055, 5.169, 1.48, 0.996, "ponto médio de altura"),
}
assert round(MATRIZ[(9, 1.58)][1], 2) == max(g[4] for g in GEO)
assert round(MATRIZ[(9, 0)][1], 2) == max(g[6] for g in GEO)
sl, cor = base("resultados", "Etapa 5 · Resultados · Geometria", "Calado × maré no ângulo dos cabos",
               "O calado pesa mais que a maré: de 9 m para 14 m, a buzina desce 5 m; a maré a move 1,58 m",
               f"Fonte: {NOTEBOOK}, função calcular_geometria_amarracao · altura relativa: SILVA (2026a, "
               "Tabela 4)", etapa="resultados")
mx, my = X0, Y_CONT
hw, hh_ = 1.6, 0.5                     # coluna de rótulos das linhas e altura do cabeçalho
cw2 = (IMG["w"] - hw - 0.2) / 2
ch2 = 1.55
for j, (tit, sub) in enumerate([("Maré baixa", "0 m"), ("Maré alta", "1,58 m")]):
    xx = mx + hw + 0.1 + j * (cw2 + 0.1)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, xx, my, cw2, hh_, cor, raio=0.2)
    texto(sl, xx, my, cw2, hh_, f"**{tit}** · {sub}", 13, "#FFFFFF", align=PP_ALIGN.CENTER,
          ancora=MSO_ANCHOR.MIDDLE, linha=1.0)
for i, (cal, tit, sub) in enumerate([(9, "Lastro", "calado 9 m"), (14, "Plena carga", "calado 14 m")]):
    yy = my + hh_ + 0.1 + i * (ch2 + 0.1)
    forma(sl, MSO_SHAPE.ROUNDED_RECTANGLE, mx, yy, hw, ch2, tint(cor, 0.30), raio=0.08)
    texto(sl, mx + 0.1, yy, hw - 0.2, ch2, [f"**{tit}**", sub], 13.5, TXT, align=PP_ALIGN.CENTER,
          ancora=MSO_ANCHOR.MIDDLE, entre=2)
    for j, mar in enumerate((0, 1.58)):
        tmin, tmax, dz, cmin, rel = MATRIZ[(cal, mar)]
        xx = mx + hw + 0.1 + j * (cw2 + 0.1)
        intens = {(9, 1.58): 0.34, (9, 0): 0.22, (14, 1.58): 0.10, (14, 0): 0.04}[(cal, mar)]
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
        ("→ Maré 0 → 1,58 m", "ângulo +1,9° a +4,9°; cabos até 0,5 m mais longos"),
        ("↓ Calado 9 → 14 m", "ângulo −6,6° a −16,4°; cabos 0,3 a 1,2 m mais curtos")]):
    xx = mx + j * (IMG["w"] / 2 + 0.05)
    ww = IMG["w"] / 2 - 0.05
    card(sl, xx, ey, ww, eh)
    texto(sl, xx + 0.15, ey + 0.08, ww - 0.3, eh - 0.12, [f"**{tit}**", corpo], 11.5, TXT2, entre=1,
          ancora=MSO_ANCHOR.MIDDLE)
card_texto(sl, COL["x"], Y_CONT, COL["w"], 1.8, "O que a matriz relaciona",
           ["Cruza o calado (lastro e plena carga) com a maré (0 e 1,58 m). Cada célula traz a faixa de "
            "ângulo vertical dos 12 cabos e a altura da buzina em relação ao cabeço.",
            "**Intuito:** separar o efeito do calado do efeito da maré."], cor, size=12, num=1)
card_texto(sl, COL["x"], Y_CONT + 1.95, COL["w"], IMG["h"] - 1.95, "O que isso significa",
           ["•  Todos os casos ficam abaixo de 25°: no pior (21,6°), 93% da tração age na horizontal.",
            "•  Com a buzina acima do cabeço, subir o navio alonga os cabos e aumenta a tração; carregar "
            "encurta e afrouxa.",
            "•  A diferença de cerca de 9,5 m entre os cabos 9 e 10 se mantém nos quatro casos."], cor,
           kpi="5 m × 1,58 m", kpi_legenda="deslocamento da buzina: calado × maré", size=12, num=2)

# ══════════════════════════════════════════════════════════════
# Etapa 6 · Conclusões parciais
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", "Etapa 6 · Conclusões", "Conclusões parciais",
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
# Alternativas para aumentar a segurança operacional
# ══════════════════════════════════════════════════════════════
sl, cor = base("conclusoes", "Etapa 6 · Conclusões", "Alternativas para aumentar a segurança",
               "Medidas voltadas aos pontos críticos: cabo 9, cabeço S e limites de operação",
               "Fontes: OCIMF (2018, seções 1.3 e 1.4; Figuras 1.3 e 1.4); PIANC (2025, cap. 4); "
               "SILVA (2026a, seção 2.2.1)", etapa="conclusoes")
POR_QUE = [
    "o cabo 9 chega a 75,3% MBL e o cabeço S recebe 4 dos 12 cabos.",
    "o cabeço S chega a 110,8% SWL, cerca de 83 t para SWL de 75 t.",
    "19 de 20 faixas passam do limite, e a maré varia até 1,68 m.",
    "a tração não é medida, e o risco se concentra em dois elementos.",
]
alternativas = [
    ("Arranjo dos cabos", "C1 · C2",
     ["•  Reduzir a diferença de rigidez entre os cabos 9 e 10, que dividem o cabeço N (comprimento, "
      "material ou pré-tensão).",
      "•  Redistribuir os 4 cabos do cabeço S entre cabeços vizinhos.",
      "•  Manter um arranjo equilibrado entre proa e popa (OCIMF, 2018, seção 1.4.1)."]),
    ("Capacidade do berço", "C2",
     ["•  Verificar a capacidade do cabeço S, que recebe 4 cabos: nos exemplos do MEG4, pontos com "
      "4 ganchos têm SWL de 240% a 310% do MBL; um único gancho ou cabeço, de 100% a 118% "
      "(OCIMF, 2018, Figuras 1.3 e 1.4).",
      "•  Avaliar cabeços de maior capacidade ou adicionais (PIANC, 2025, seção 4.9)."]),
    ("Gestão da operação", "C3 · C4",
     ["•  Ajustar a tensão dos cabos ao longo da maré e do carregamento (OCIMF, 2018, seção 1.3).",
      "•  Definir limites de paragem com vento, corrente, Hs, direção e maré.",
      "•  Retirar os cabos de uso quando a resistência residual chegar a 75% do MBL de projeto "
      "(OCIMF, 2018, Figuras 1.3 e 1.4)."]),
    ("Monitoramento contínuo", "C1 · C2",
     ["•  Medir a carga em tempo real com cabeços instrumentados (smart bollards) ou ganchos com "
      "sensor de carga (PIANC, 2025, seção 4.10.6).",
      "•  Aplicar a visão computacional proposta no TF, começando pelo cabo 9 e pelo cabeço S."]),
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
           "Simulação de 394 condições com calado de 9 m e 14 m, análise dos esforços, dos riscos e da "
           "geometria."),
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
            "•  Reavaliar os resultados com 50% do MBL",
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
