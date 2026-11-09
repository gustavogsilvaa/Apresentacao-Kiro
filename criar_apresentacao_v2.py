import sys
sys.path.insert(0, '/projects/sandbox/pptx_pkgs')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

# ══════════════════════════════════════════════════════════════════
#  PALETA  — Design moderno escuro/gradiente
# ══════════════════════════════════════════════════════════════════
NAVY        = RGBColor(0x0D, 0x1B, 0x2A)   # fundo escuro principal
AZUL_MED    = RGBColor(0x1B, 0x4F, 0x8A)   # azul médio
AZUL_VIVO   = RGBColor(0x14, 0x8F, 0xD0)   # azul destaque
CIANO       = RGBColor(0x00, 0xC8, 0xD7)   # accent ciano
VERDE_NEO   = RGBColor(0x06, 0xD6, 0x8A)   # accent verde
AMARELO     = RGBColor(0xFF, 0xC3, 0x00)   # accent amarelo/alerta
VERMELHO    = RGBColor(0xE5, 0x3E, 0x3E)   # perigo
BRANCO      = RGBColor(0xFF, 0xFF, 0xFF)
BRANCO_SUB  = RGBColor(0xCC, 0xDD, 0xEE)   # texto secundário branco
CINZA_CARD  = RGBColor(0x16, 0x2A, 0x3F)   # fundo card
CINZA_BDR   = RGBColor(0x2A, 0x4A, 0x6A)   # borda card

W = Inches(16)
H = Inches(9)

# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════
def nova_prs():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs

def rect(slide, l, t, w, h, fill_rgb, alpha=None, line_rgb=None, line_pt=0):
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        s.line.color.rgb = line_rgb
        s.line.width = Pt(line_pt)
    else:
        s.line.fill.background()
    return s

def txt_box(slide, texto, l, t, w, h, size=11, bold=False, cor=BRANCO,
            align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = texto
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = cor
    return tb

def add_img(slide, path, l, t, w, h):
    if path and os.path.exists(path):
        slide.shapes.add_picture(path, l, t, w, h)
        return True
    # placeholder se não tiver imagem
    r = rect(slide, l, t, w, h, CINZA_CARD, line_rgb=CINZA_BDR, line_pt=1)
    txt_box(slide, "[ Gráfico ]", l, t + h//2 - Inches(0.2),
            w, Inches(0.4), size=10, italic=True,
            cor=RGBColor(0x55,0x88,0xAA), align=PP_ALIGN.CENTER)
    return False

def linha_decorativa(slide, cor=CIANO, espessura_pt=3):
    s = rect(slide, Inches(0), Inches(1.18), W, Pt(espessura_pt), cor)

def barra_topo(slide, titulo, subtitulo=None, cor_accent=CIANO):
    # fundo total
    rect(slide, Inches(0), Inches(0), W, H, NAVY)
    # faixa topo gradiente simulada (2 retângulos)
    rect(slide, Inches(0), Inches(0), W, Inches(1.22), AZUL_MED)
    rect(slide, Inches(0), Inches(0), W, Inches(0.06), cor_accent)
    # linha inferior da barra
    rect(slide, Inches(0), Inches(1.16), W, Inches(0.04), cor_accent)
    # título
    txt_box(slide, titulo, Inches(0.45), Inches(0.10),
            Inches(13), Inches(0.65),
            size=24, bold=True, cor=BRANCO)
    if subtitulo:
        txt_box(slide, subtitulo, Inches(0.45), Inches(0.72),
                Inches(13), Inches(0.38),
                size=10.5, cor=BRANCO_SUB)

def card(slide, l, t, w, h, titulo=None, corpo=None,
         cor_titulo=CIANO, tamanho_titulo=11.5, tamanho_corpo=10,
         fundo=CINZA_CARD, borda=CINZA_BDR):
    rect(slide, l, t, w, h, fundo, line_rgb=borda, line_pt=0.8)
    # faixa colorida no topo do card
    rect(slide, l, t, w, Pt(3), cor_titulo)
    if titulo:
        txt_box(slide, titulo,
                l + Inches(0.15), t + Inches(0.12),
                w - Inches(0.3), Inches(0.35),
                size=tamanho_titulo, bold=True, cor=cor_titulo)
    if corpo:
        txt_box(slide, corpo,
                l + Inches(0.15), t + Inches(0.52),
                w - Inches(0.3), h - Inches(0.65),
                size=tamanho_corpo, cor=BRANCO_SUB, wrap=True)

def rodape(slide, num, total):
    rect(slide, Inches(0), Inches(8.75), W, Inches(0.25), AZUL_MED)
    txt_box(slide,
            f"Análise de Amarração Marítima — Calado 9 m  |  Navio Panamax  |  394 Condições Simuladas  |  {num}/{total}",
            Inches(0.3), Inches(8.76), Inches(15.4), Inches(0.22),
            size=7.5, cor=RGBColor(0x80,0xA8,0xCC), align=PP_ALIGN.CENTER)

def kpi_box(slide, l, t, w, h, valor, label, cor_valor=CIANO):
    rect(slide, l, t, w, h, CINZA_CARD, line_rgb=CINZA_BDR, line_pt=0.8)
    rect(slide, l, t, w, Pt(3), cor_valor)
    txt_box(slide, valor, l, t + Inches(0.15), w, Inches(0.55),
            size=26, bold=True, cor=cor_valor, align=PP_ALIGN.CENTER)
    txt_box(slide, label, l, t + Inches(0.68), w, Inches(0.32),
            size=9, cor=BRANCO_SUB, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════
#  IMAGENS
# ══════════════════════════════════════════════════════════════════
IMG = {
    "hist"      : "/projects/sandbox/imgs_graficos/01_histogramas.png",
    "mare"      : "/projects/sandbox/imgs_graficos/02_influencia_mare.png",
    "risco"     : "/projects/sandbox/imgs_graficos/03_combinacoes_criticas.png",
    "corr"      : "/projects/sandbox/imgs_graficos/04_correlacao_boxplot.png",
    "envelope"  : "/projects/sandbox/imgs_graficos/05_envelope_critico.png",
    "operab"    : "/projects/sandbox/imgs_graficos/06_matriz_operabilidade.png",
    "geo1"      : "/projects/sandbox/imgs_graficos/07_geometria_9m_alta.png",
    "geo2"      : "/projects/sandbox/imgs_graficos/08_geometria_9m_zero.png",
}

prs = nova_prs()
TOTAL_SLIDES = 11

# ══════════════════════════════════════════════════════════════════
#  SLIDE 1 — CAPA
# ══════════════════════════════════════════════════════════════════
def slide_capa(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, Inches(0), Inches(0), W, H, NAVY)
    # faixas decorativas horizontais
    rect(slide, Inches(0), Inches(0),   W, Inches(0.08), CIANO)
    rect(slide, Inches(0), Inches(8.92),W, Inches(0.08), CIANO)
    rect(slide, Inches(0), Inches(3.5), W, Inches(0.04), AZUL_MED)
    rect(slide, Inches(0), Inches(6.2), W, Inches(0.04), AZUL_MED)
    # bloco central
    rect(slide, Inches(1.0), Inches(1.2), Inches(14), Inches(6.5), CINZA_CARD,
         line_rgb=AZUL_VIVO, line_pt=1.5)
    rect(slide, Inches(1.0), Inches(1.2), Inches(14), Pt(4), CIANO)

    # título principal
    txt_box(slide, "ANÁLISE DE SISTEMA DE AMARRAÇÃO MARÍTIMA",
            Inches(1.5), Inches(2.0), Inches(13), Inches(1.0),
            size=34, bold=True, cor=BRANCO, align=PP_ALIGN.CENTER)
    # linha separadora
    rect(slide, Inches(3.5), Inches(3.25), Inches(9), Inches(0.04), CIANO)
    # subtítulo
    txt_box(slide, "Resultados de Simulação Meteoceanográfica — Calado 9 m",
            Inches(1.5), Inches(3.45), Inches(13), Inches(0.55),
            size=17, cor=BRANCO_SUB, align=PP_ALIGN.CENTER)
    # KPIs centrais
    kpi_box(slide, Inches(2.5),  Inches(4.3), Inches(2.5), Inches(1.2),
            "394", "Condições Simuladas", CIANO)
    kpi_box(slide, Inches(5.5),  Inches(4.3), Inches(2.5), Inches(1.2),
            "58",  "Condições > 55% MBL", AMARELO)
    kpi_box(slide, Inches(8.5),  Inches(4.3), Inches(2.5), Inches(1.2),
            "2",   "Condições > 100% SWL", VERMELHO)
    kpi_box(slide, Inches(11.5), Inches(4.3), Inches(2.5), Inches(1.2),
            "9m",  "Calado Analisado", VERDE_NEO)
    # rodapé capa
    txt_box(slide,
            "Software de Simulação Dinâmica  •  Navio Panamax  •  Berço Monobóia  •  Análise Automatizada em Python",
            Inches(1.0), Inches(7.8), Inches(14), Inches(0.4),
            size=9, cor=RGBColor(0x50,0x78,0xA0), align=PP_ALIGN.CENTER)

slide_capa(prs)

# ══════════════════════════════════════════════════════════════════
#  FUNÇÃO GENÉRICA: slide com imagem + 2 cards laterais
# ══════════════════════════════════════════════════════════════════
def slide_img_cards(prs, titulo, subtitulo, img_key,
                    txt_analisa, txt_resultado,
                    cor_accent=CIANO, num=1):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    barra_topo(slide, titulo, subtitulo, cor_accent)
    rodape(slide, num, TOTAL_SLIDES)

    # área imagem
    add_img(slide, IMG.get(img_key), Inches(0.3), Inches(1.28), Inches(9.6), Inches(7.2))

    # cards laterais
    card(slide,
         Inches(10.2), Inches(1.28), Inches(5.5), Inches(3.4),
         titulo="O QUE ESTE GRÁFICO ANALISA",
         corpo=txt_analisa,
         cor_titulo=AZUL_VIVO, tamanho_titulo=11, tamanho_corpo=10)

    card(slide,
         Inches(10.2), Inches(4.88), Inches(5.5), Inches(3.6),
         titulo="RESULTADO OBTIDO",
         corpo=txt_resultado,
         cor_titulo=VERDE_NEO, tamanho_titulo=11, tamanho_corpo=10)

# ══════════════════════════════════════════════════════════════════
#  FUNÇÃO GENÉRICA: slide com tabela + 2 cards
# ══════════════════════════════════════════════════════════════════
def slide_tabela_cards(prs, titulo, subtitulo, headers, rows,
                       txt_analisa, txt_resultado,
                       cor_accent=CIANO, num=1):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    barra_topo(slide, titulo, subtitulo, cor_accent)
    rodape(slide, num, TOTAL_SLIDES)

    n_rows = len(rows) + 1
    n_cols = len(headers)
    row_h  = min(Inches(0.46), Inches(6.8 / n_rows))
    tbl_h  = row_h * n_rows

    tbl = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(0.3), Inches(1.30),
        Inches(9.6), tbl_h
    ).table

    # cabeçalho
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = AZUL_VIVO
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.runs[0]
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = BRANCO

    # linhas
    for ri, row in enumerate(rows):
        bg = CINZA_CARD if ri % 2 == 0 else RGBColor(0x12, 0x22, 0x35)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.runs[0]
            r.font.size = Pt(9)
            r.font.color.rgb = BRANCO_SUB

    card(slide,
         Inches(10.2), Inches(1.28), Inches(5.5), Inches(3.4),
         titulo="O QUE ESTA TABELA ANALISA",
         corpo=txt_analisa,
         cor_titulo=AZUL_VIVO, tamanho_titulo=11, tamanho_corpo=10)

    card(slide,
         Inches(10.2), Inches(4.88), Inches(5.5), Inches(3.6),
         titulo="RESULTADO OBTIDO",
         corpo=txt_resultado,
         cor_titulo=VERDE_NEO, tamanho_titulo=11, tamanho_corpo=10)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 2 — Resumo Estatístico
# ══════════════════════════════════════════════════════════════════
headers_resumo = ["Variável", "Média", "Mediana", "P90", "P95", "P99", "Máximo"]
rows_resumo = [
    ["Tração máx. [tf]",    "25,15", "23,56", "44,73", "47,25", "51,82", "56,44"],
    ["% MBL",               "33,54", "31,45", "59,61", "63,00", "69,10", "75,30"],
    ["Reação máx. [tf]",    "17,43", "16,61", "28,48", "32,41", "43,72", "51,86"],
    ["Compressão máx. [m]", "0,04",  "0,04",  "0,07",  "0,08",  "0,11",  "0,13"],
    ["Cabeço máx. [tf]",    "34,50", "31,88", "48,56", "55,63", "68,17", "83,13"],
    ["% SWL",               "46,00", "42,50", "64,77", "74,17", "90,91", "110,80"],
]
slide_tabela_cards(prs,
    "Resumo Estatístico Geral",
    "Estatísticas descritivas dos esforços do sistema — 394 condições simuladas",
    headers_resumo, rows_resumo,
    txt_analisa=(
        "Síntese quantitativa dos esforços estruturais: tração nos cabos "
        "(tf e % MBL), reação e compressão nas defensas, e força nos cabeços "
        "(tf e % SWL).\n\n"
        "Os percentis P90, P95 e P99 revelam o comportamento do sistema nas "
        "condições mais severas, fundamentais para a definição de limites "
        "operacionais seguros."
    ),
    txt_resultado=(
        "A tração mediana é de 23,6 tf (31,5% MBL) — sistema opera com folga "
        "na maior parte do tempo.\n\n"
        "No P99, a tração atinge 51,8 tf (69,1% MBL), acima do limite OCIMF "
        "de 55% MBL.\n\n"
        "O % SWL máximo de 110,8% confirma 2 condições com risco estrutural "
        "real nos cabeços."
    ),
    cor_accent=CIANO, num=2
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 3 — Histogramas
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Distribuição de Frequência — Cabos e Cabeços",
    "Histogramas de % MBL (Cabos) e % SWL (Cabeços) | 394 condições simuladas",
    "hist",
    txt_analisa=(
        "Mostra com que frequência o sistema operou em cada faixa de "
        "carregamento ao longo das 394 condições.\n\n"
        "O eixo X representa o nível de utilização da capacidade (% MBL "
        "para cabos; % SWL para cabeços) e o eixo Y a quantidade de "
        "ocorrências.\n\n"
        "A linha vermelha tracejada marca os limites operacionais: 55% MBL "
        "(recomendação OCIMF) e 100% SWL."
    ),
    txt_resultado=(
        "92 das 394 condições concentram-se na faixa 10–20% MBL, confirmando "
        "que o regime habitual é de baixa solicitação.\n\n"
        "Entretanto, 58 condições (≈15%) ultrapassam 55% MBL — risco "
        "operacional recorrente.\n\n"
        "Para os cabeços, apenas 2 condições excedem 100% SWL, mas o "
        "histograma mostra uma cauda longa até 110%, sinalizando "
        "proximidade sistemática do limite estrutural."
    ),
    cor_accent=AZUL_VIVO, num=3
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 4 — Influência da Maré
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Influência da Maré nos Esforços",
    "Scatter Plot: Pico de Tensão (% MBL) e Pico de Esforço (% SWL) × Nível da Maré",
    "mare",
    txt_analisa=(
        "Correlaciona o nível de maré (escala de cores azul→vermelho) com "
        "o pico de tensão de cada cabo e de esforço de cada cabeço.\n\n"
        "Permite identificar se a maré é um fator agravante — cores quentes "
        "em posição elevada indicam que marés altas concentram os maiores "
        "esforços em elementos específicos do sistema."
    ),
    txt_resultado=(
        "O Cabo 9 se destaca com tensões consistentemente acima de 65% MBL "
        "(ponto vermelho intenso), independentemente do nível de maré.\n\n"
        "Marés altas (tons avermelhados) coincidem com os picos mais críticos, "
        "indicando que a geometria das linhas em maré alta penaliza "
        "especialmente o Cabo 9.\n\n"
        "O Cabeço S é o mais crítico: atinge 110% SWL com maré de 1,58 m, "
        "único ponto acima do limite estrutural."
    ),
    cor_accent=CIANO, num=4
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 5 — Combinações Críticas (com imagem da tabela)
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Filtro de Riscos — Combinações Críticas",
    "TOP 5 condições > 55% MBL (Cabos) e condições > 100% SWL (Cabeços)",
    "risco",
    txt_analisa=(
        "Filtra e classifica as condições meteoceanográficas que violam as "
        "margens de segurança (55% MBL para cabos; 100% SWL para cabeços).\n\n"
        "Exibe as combinações reais de vento, corrente e onda que representam "
        "risco iminente de ruptura de cabo ou arrancamento de cabeço, "
        "fornecendo os dados de entrada para os limites operacionais do berço."
    ),
    txt_resultado=(
        "O Cabo 9 aparece em TODOS os 5 piores cenários de tração:\n"
        "• Cond. 259: 75,3% MBL (vento 8,2 nós, corrente 1,57 nós)\n"
        "• Cond. 310: 73,0% MBL (vento 41,4 nós, corrente 1,16 nós)\n\n"
        "O Cabeço S ultrapassou 100% SWL em 2 condições:\n"
        "• Cond. 198: 110,8% SWL (vento 36,7 nós, corrente 1,58 nós)\n"
        "• Cond. 55:  104,5% SWL (vento 49,1 nós, corrente 0,88 nós)"
    ),
    cor_accent=VERMELHO, num=5
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 6 — Correlação + Boxplot
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Diagnóstico de Carga — Correlação e Distribuição por Cabo",
    "Matriz de Correlação (Ambiente × Esforços) e Boxplot de % MBL por Cabo Crítico",
    "corr",
    txt_analisa=(
        "① Matriz de Correlação — Mede a relação linear entre variáveis "
        "ambientais (vento, corrente, Hs, Tp) e os esforços (% MBL, % SWL). "
        "Valores próximos de ±1 indicam forte dependência.\n\n"
        "② Boxplot por Cabo — Exibe a distribuição estatística (mediana, "
        "quartis e outliers) da tensão por cabo, ordenados do mais ao menos "
        "crítico, revelando assimetrias no plano de linhas."
    ),
    txt_resultado=(
        "Correlação mais relevante: % MBL × % SWL = 0,70 — os mesmos eventos "
        "que tensionam os cabos sobrecarregam os cabeços.\n\n"
        "Vento tem baixa correlação com % MBL (0,02), confirmando que corrente "
        "e onda são as forças dominantes nos cabos.\n\n"
        "No Boxplot, o Cabo 9 tem mediana e amplitude muito superiores aos "
        "demais: absorve desproporcionalmente a carga dinâmica do sistema — "
        "indicativo de assimetria no plano de linhas."
    ),
    cor_accent=AMARELO, num=6
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 7 — Envelope Crítico
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Envelope Crítico — Vento × Corrente",
    "Mapa de esforços: cor = intensidade (% MBL/SWL) | tamanho = nível de maré",
    "envelope",
    txt_analisa=(
        "Gráfico de dispersão com eixos Corrente [nós] × Vento [nós]. Cada "
        "ponto representa um elemento estrutural em determinada combinação "
        "ambiental.\n\n"
        "Cor (fria→quente) = intensidade do esforço. "
        "Tamanho do ponto = nível de maré (maior = maré mais alta).\n\n"
        "Constrói o 'envelope de paragem': região onde o sistema ultrapassa "
        "os limites de segurança."
    ),
    txt_resultado=(
        "O ponto mais crítico dos cabos (vermelho, maior tamanho) aparece em "
        "corrente ≈ 1,6 nós e vento ≈ 10 nós, tensão > 70% MBL — "
        "demonstrando que corrente elevada domina sobre vento.\n\n"
        "Para os cabeços, o Cabeço S ultrapassa 100% SWL com corrente ≈ 1,6 "
        "nós e vento ≈ 37 nós (maré alta).\n\n"
        "A maré alta amplifica os esforços ao alterar a geometria das linhas "
        "e reduzir a eficiência horizontal dos cabos."
    ),
    cor_accent=VERMELHO, num=7
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 8 — Matriz de Operabilidade
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Matriz de Operabilidade — Limites de Paragem",
    "Mapa de calor: Tensão máxima (% MBL) por Faixa de Vento × Faixa de Corrente",
    "operab",
    txt_analisa=(
        "Transforma 394 simulações em um mapa operacional de fácil leitura, "
        "cruzando faixas de vento (eixo Y, intervalos de 10 nós) com faixas "
        "de corrente (eixo X, intervalos de 0,5 nós).\n\n"
        "Cada célula exibe a pior tensão (% MBL) registrada naquela combinação. "
        "Verde → Amarelo → Vermelho indica o grau de criticidade, com "
        "limite de paragem em 55% MBL."
    ),
    txt_resultado=(
        "Quase todas as células estão acima de 59% MBL (laranja/vermelho), "
        "confirmando que o Cabo 9 ultrapassa o limite operacional na vasta "
        "maioria das combinações avaliadas.\n\n"
        "O pico absoluto (75,3% MBL) ocorre com vento 0–10 nós e corrente "
        "1,5–2,0 nós — corrente domina sobre vento.\n\n"
        "A única célula verde (38,8% MBL) ocorre com vento 10–20 nós e "
        "corrente 1,5–2,0 nós, sendo uma exceção que merece verificação."
    ),
    cor_accent=VERDE_NEO, num=8
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 9 — Geometria 9m / Maré Alta
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Geometria das Linhas — Calado 9 m | Maré Alta 1,58 m",
    "Comprimento 3D e ângulo vertical calculados para o cenário mais crítico",
    "geo1",
    txt_analisa=(
        "Calcula analiticamente a geometria espacial (comprimento 3D e "
        "ângulo vertical) de cada linha de amarração para o cenário de "
        "calado 9 m com maré alta (1,58 m) — o cenário mais crítico "
        "identificado na análise.\n\n"
        "O ângulo vertical é determinante: acima de ~20° a componente "
        "horizontal de retenção do navio cai significativamente, "
        "aumentando a tração exigida."
    ),
    txt_resultado=(
        "Fairlead 17 → Cabeço B: maior ângulo vertical (21,6°) — linha "
        "em condição geométrica desfavorável.\n\n"
        "Fairleads 23 e 24 → Cabeço G: maiores comprimentos (≈ 41,8 m) "
        "e menores ângulos (≈ 9°) — configuração mais eficiente.\n\n"
        "Fairleads 2 e 3 → Cabeço S: ângulos de 15–16°, comprimentos "
        "de ~24 m — par crítico por concentrar o maior esforço no "
        "Cabeço S (110,8% SWL)."
    ),
    cor_accent=AZUL_VIVO, num=9
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 10 — Geometria 9m / Maré Zero
# ══════════════════════════════════════════════════════════════════
slide_img_cards(prs,
    "Geometria das Linhas — Calado 9 m | Maré Zero (0 m)",
    "Comparativo: efeito da remoção da maré sobre os ângulos verticais",
    "geo2",
    txt_analisa=(
        "Repete o cálculo geométrico para o mesmo calado (9 m) com maré "
        "zero, permitindo isolar o efeito exclusivo da variação de maré "
        "sobre a geometria das linhas.\n\n"
        "A comparação direta com o cenário de maré alta (1,58 m) quantifica "
        "o quanto a maré altera a inclinação das linhas e, "
        "consequentemente, a distribuição de esforços no sistema."
    ),
    txt_resultado=(
        "Com maré zero, todos os ângulos reduzem 2–5° em relação ao "
        "cenário de maré alta:\n"
        "• Fairlead 17 → B: 16,7° (era 21,6° com maré alta)\n"
        "• Fairleads 2/3 → S: 11,9–12,1° (eram 15,6–15,9°)\n\n"
        "A redução da maré melhora a geometria mas não elimina a "
        "assimetria: o Cabo 9 continua sendo o elemento mais solicitado, "
        "independentemente do nível de maré."
    ),
    cor_accent=AZUL_VIVO, num=10
)

# ══════════════════════════════════════════════════════════════════
#  SLIDE 11 — Conclusões
# ══════════════════════════════════════════════════════════════════
def slide_conclusoes(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    barra_topo(slide, "Conclusões e Recomendações",
               "Síntese executiva dos resultados — Análise de Amarração Marítima | Calado 9 m",
               cor_accent=CIANO)
    rodape(slide, TOTAL_SLIDES, TOTAL_SLIDES)

    conclusoes = [
        (VERMELHO,    "⚠  ELEMENTO CRÍTICO IDENTIFICADO",
         "O Cabo 9 e o Cabeço S são os elementos mais vulneráveis. "
         "O Cabo 9 aparece nos 5 piores cenários de tração (até 75,3% MBL). "
         "O Cabeço S ultrapassou o limite estrutural absoluto (100% SWL) em "
         "2 condições, com pico de 110,8% SWL."),

        (CIANO,       "📊  PERFIL OPERACIONAL",
         "O sistema opera com segurança na maioria do tempo (mediana 31,5% MBL). "
         "Contudo, 58 condições (≈15% do total) violam o limite de 55% MBL, "
         "indicando necessidade de limites operacionais formais para o berço."),

        (AMARELO,     "🌊  FATORES DOMINANTES",
         "Corrente elevada (> 1,4 nós) é o principal fator de risco, "
         "independente do vento. A maré alta (1,58 m) amplifica os esforços "
         "ao alterar a geometria das linhas — especialmente no Fairlead 17 "
         "(ângulo de 21,6°, o maior registrado)."),

        (VERDE_NEO,   "📐  GEOMETRIA DAS LINHAS",
         "Com calado 14 m e maré zero, os ângulos caem para < 0,4° — linhas "
         "praticamente horizontais. Esta condição minimiza a absorção vertical "
         "e pode favorecer movimentos de surgimento. Recomenda-se revisão do "
         "plano de amarração para calados superiores a 12 m."),
    ]

    positions = [
        (Inches(0.3),  Inches(1.30)),
        (Inches(8.15), Inches(1.30)),
        (Inches(0.3),  Inches(5.05)),
        (Inches(8.15), Inches(5.05)),
    ]

    for (cor, tit, corpo), (l, t) in zip(conclusoes, positions):
        cw = Inches(7.55)
        ch = Inches(3.45)
        rect(slide, l, t, cw, ch, CINZA_CARD, line_rgb=CINZA_BDR, line_pt=0.8)
        rect(slide, l, t, cw, Pt(4), cor)
        # ícone colorido lateral
        rect(slide, l, t, Inches(0.1), ch, cor)
        txt_box(slide, tit, l + Inches(0.22), t + Inches(0.12),
                cw - Inches(0.3), Inches(0.35),
                size=12, bold=True, cor=cor)
        txt_box(slide, corpo, l + Inches(0.22), t + Inches(0.55),
                cw - Inches(0.3), ch - Inches(0.65),
                size=10.5, cor=BRANCO_SUB, wrap=True)

slide_conclusoes(prs)

# ══════════════════════════════════════════════════════════════════
#  SALVAR
# ══════════════════════════════════════════════════════════════════
OUTPUT = "/projects/sandbox/Apresentacao_Amarracao_Maritima_v2.pptx"
prs.save(OUTPUT)
print(f"✅  Apresentação salva: {OUTPUT}")
print(f"📊  Total de slides: {len(prs.slides)}")
