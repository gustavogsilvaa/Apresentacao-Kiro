import sys
sys.path.insert(0, '/projects/sandbox/pptx_pkgs')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ─────────────────────────────────────────
#  PALETA DE CORES  (suave, clara, objetiva)
# ─────────────────────────────────────────
AZUL_ESCURO   = RGBColor(0x1A, 0x3A, 0x5C)   # título principal
AZUL_MEDIO    = RGBColor(0x2E, 0x6D, 0xA8)   # destaque
AZUL_CLARO    = RGBColor(0xD6, 0xE8, 0xF7)   # fundo de caixa
CINZA_CLARO   = RGBColor(0xF4, 0xF6, 0xF9)   # fundo de slide
CINZA_MEDIO   = RGBColor(0xB0, 0xBE, 0xCE)   # linha separadora
VERDE         = RGBColor(0x2E, 0x8B, 0x57)   # accent positivo
TEXTO_ESCURO  = RGBColor(0x22, 0x2E, 0x3E)   # texto geral
BRANCO        = RGBColor(0xFF, 0xFF, 0xFF)
VERMELHO_SUAVE= RGBColor(0xC0, 0x3B, 0x3B)   # alerta

def nova_apresentacao():
    prs = Presentation()
    prs.slide_width  = Inches(16)
    prs.slide_height = Inches(9)
    return prs

# ─────────────────────────────────────────
#  HELPER: adicionar fundo ao slide
# ─────────────────────────────────────────
def fundo_slide(slide, cor=CINZA_CLARO):
    from pptx.util import Inches, Pt
    from pptx.oxml.ns import qn
    import lxml.etree as etree
    # retângulo cobrindo o slide inteiro
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(0), Inches(0),
        Inches(16), Inches(9)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = cor
    shape.line.fill.background()
    shape.zorder = 0
    return shape

# ─────────────────────────────────────────
#  HELPER: caixa de texto formatada
# ─────────────────────────────────────────
def caixa(slide, texto, left, top, width, height,
          fundo=AZUL_CLARO, titulo_texto=None, titulo_cor=AZUL_MEDIO,
          texto_cor=TEXTO_ESCURO, tamanho_titulo=13, tamanho_texto=11,
          bold_titulo=True):
    from pptx.util import Inches, Pt
    # fundo da caixa
    box_bg = slide.shapes.add_shape(1, left, top, width, height)
    box_bg.fill.solid()
    box_bg.fill.fore_color.rgb = fundo
    box_bg.line.color.rgb = CINZA_MEDIO
    box_bg.line.width = Pt(0.75)

    # caixa de texto sobre o fundo
    txBox = slide.shapes.add_textbox(
        left + Inches(0.18), top + Inches(0.15),
        width - Inches(0.36), height - Inches(0.25)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    if titulo_texto:
        p0 = tf.paragraphs[0]
        p0.alignment = PP_ALIGN.LEFT
        run0 = p0.add_run()
        run0.text = titulo_texto
        run0.font.bold = bold_titulo
        run0.font.size = Pt(tamanho_titulo)
        run0.font.color.rgb = titulo_cor
        # parágrafo do texto principal
        p1 = tf.add_paragraph()
    else:
        p1 = tf.paragraphs[0]

    p1.alignment = PP_ALIGN.LEFT
    run1 = p1.add_run()
    run1.text = texto
    run1.font.size = Pt(tamanho_texto)
    run1.font.color.rgb = texto_cor

# ─────────────────────────────────────────
#  HELPER: retângulo de título do slide
# ─────────────────────────────────────────
def barra_titulo(slide, titulo, subtitulo=None):
    # faixa azul no topo
    barra = slide.shapes.add_shape(
        1, Inches(0), Inches(0), Inches(16), Inches(1.15)
    )
    barra.fill.solid()
    barra.fill.fore_color.rgb = AZUL_ESCURO
    barra.line.fill.background()

    txBox = slide.shapes.add_textbox(Inches(0.4), Inches(0.10), Inches(14), Inches(0.65))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = titulo
    run.font.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = BRANCO

    if subtitulo:
        txBox2 = slide.shapes.add_textbox(Inches(0.4), Inches(0.72), Inches(14), Inches(0.38))
        tf2 = txBox2.text_frame
        p2 = tf2.paragraphs[0]
        run2 = p2.add_run()
        run2.text = subtitulo
        run2.font.size = Pt(11)
        run2.font.color.rgb = RGBColor(0xB0, 0xD0, 0xF0)

# ─────────────────────────────────────────
#  HELPER: adicionar imagem (placeholder cinza se não houver img)
# ─────────────────────────────────────────
def bloco_imagem(slide, left, top, width, height, label="Gráfico gerado pelo Python"):
    # placeholder cinza médio
    ph = slide.shapes.add_shape(1, left, top, width, height)
    ph.fill.solid()
    ph.fill.fore_color.rgb = RGBColor(0xD8, 0xE5, 0xF0)
    ph.line.color.rgb = AZUL_MEDIO
    ph.line.width = Pt(1)

    tx = slide.shapes.add_textbox(
        left + Inches(0.2), top + height // 2 - Inches(0.3),
        width - Inches(0.4), Inches(0.6)
    )
    tf = tx.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = AZUL_MEDIO

# ─────────────────────────────────────────
#  SLIDE 0: CAPA
# ─────────────────────────────────────────
def slide_capa(prs):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    fundo_slide(slide, AZUL_ESCURO)

    # Linha de detalhe decorativa
    det = slide.shapes.add_shape(1, Inches(0), Inches(5.8), Inches(16), Inches(0.06))
    det.fill.solid()
    det.fill.fore_color.rgb = AZUL_MEDIO
    det.line.fill.background()

    tx_titulo = slide.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(13), Inches(1.5))
    tf = tx_titulo.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "Análise de Sistema de Amarração Marítima"
    run.font.bold = True
    run.font.size = Pt(36)
    run.font.color.rgb = BRANCO

    tx_sub = slide.shapes.add_textbox(Inches(2), Inches(3.75), Inches(12), Inches(0.7))
    tf2 = tx_sub.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = "Resultados de Simulação Meteoceanográfica — Calado 9m"
    run2.font.size = Pt(17)
    run2.font.color.rgb = RGBColor(0xB0, 0xD0, 0xF0)

    tx_info = slide.shapes.add_textbox(Inches(2), Inches(5.0), Inches(12), Inches(0.5))
    tf3 = tx_info.text_frame
    p3 = tf3.paragraphs[0]
    p3.alignment = PP_ALIGN.CENTER
    run3 = p3.add_run()
    run3.text = "394 Condições Simuladas  •  Navio Panamax  •  Berço Monobóia"
    run3.font.size = Pt(12)
    run3.font.color.rgb = RGBColor(0x88, 0xAA, 0xCC)

    tx_rodape = slide.shapes.add_textbox(Inches(0.5), Inches(8.3), Inches(15), Inches(0.5))
    tf4 = tx_rodape.text_frame
    p4 = tf4.paragraphs[0]
    p4.alignment = PP_ALIGN.CENTER
    run4 = p4.add_run()
    run4.text = "Software de Simulação Dinâmica  |  Análise Automatizada em Python"
    run4.font.size = Pt(9)
    run4.font.color.rgb = RGBColor(0x55, 0x77, 0x99)

# ─────────────────────────────────────────
#  LAYOUT PADRÃO: imagem esquerda, caixas direita
# ─────────────────────────────────────────
def slide_padrao(prs, titulo, subtitulo,
                 txt_o_que, txt_resultado,
                 img_path=None, img_label="Gráfico"):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    fundo_slide(slide, CINZA_CLARO)
    barra_titulo(slide, titulo, subtitulo)

    # ── Área da imagem ──────────────────────
    img_left   = Inches(0.35)
    img_top    = Inches(1.30)
    img_width  = Inches(9.40)
    img_height = Inches(7.35)

    if img_path:
        try:
            slide.shapes.add_picture(img_path, img_left, img_top, img_width, img_height)
        except Exception:
            bloco_imagem(slide, img_left, img_top, img_width, img_height, img_label)
    else:
        bloco_imagem(slide, img_left, img_top, img_width, img_height, img_label)

    # ── Caixas de texto (direita) ────────────
    cx_left  = Inches(10.1)
    cx_width = Inches(5.55)
    cx_top1  = Inches(1.30)
    cx_h1    = Inches(3.40)
    cx_top2  = Inches(4.90)
    cx_h2    = Inches(3.75)

    caixa(slide, txt_o_que,
          cx_left, cx_top1, cx_width, cx_h1,
          fundo=AZUL_CLARO,
          titulo_texto="O QUE ESTE GRÁFICO ANALISA",
          titulo_cor=AZUL_MEDIO, tamanho_titulo=12, tamanho_texto=10.5)

    caixa(slide, txt_resultado,
          cx_left, cx_top2, cx_width, cx_h2,
          fundo=RGBColor(0xE8, 0xF5, 0xEB),
          titulo_texto="RESULTADO OBTIDO",
          titulo_cor=VERDE, tamanho_titulo=12, tamanho_texto=10.5)

    return slide

# ─────────────────────────────────────────
#  LAYOUT TABELA: tabela esquerda, caixas direita
# ─────────────────────────────────────────
def slide_tabela(prs, titulo, subtitulo,
                 headers, rows,
                 txt_o_que, txt_resultado):
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    fundo_slide(slide, CINZA_CLARO)
    barra_titulo(slide, titulo, subtitulo)

    # ── Tabela ──────────────────────────────
    n_rows = len(rows) + 1
    n_cols = len(headers)
    tbl = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(0.35), Inches(1.35),
        Inches(9.40), Inches(min(7.2, 0.45 * n_rows + 0.5))
    ).table

    # Estilo cabeçalho
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = AZUL_ESCURO
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = BRANCO

    # Linhas
    for ri, row in enumerate(rows):
        bg = BRANCO if ri % 2 == 0 else RGBColor(0xEA, 0xF2, 0xFB)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.runs[0]
            run.font.size = Pt(9)
            run.font.color.rgb = TEXTO_ESCURO

    # ── Caixas de texto ──────────────────────
    cx_left  = Inches(10.1)
    cx_width = Inches(5.55)
    cx_top1  = Inches(1.30)
    cx_h1    = Inches(3.40)
    cx_top2  = Inches(4.90)
    cx_h2    = Inches(3.75)

    caixa(slide, txt_o_que,
          cx_left, cx_top1, cx_width, cx_h1,
          fundo=AZUL_CLARO,
          titulo_texto="O QUE ESTA TABELA ANALISA",
          titulo_cor=AZUL_MEDIO, tamanho_titulo=12, tamanho_texto=10.5)

    caixa(slide, txt_resultado,
          cx_left, cx_top2, cx_width, cx_h2,
          fundo=RGBColor(0xE8, 0xF5, 0xEB),
          titulo_texto="RESULTADO OBTIDO",
          titulo_cor=VERDE, tamanho_titulo=12, tamanho_texto=10.5)

    return slide

# ═══════════════════════════════════════════════════════════════
#  CONSTRUÇÃO DOS SLIDES
# ═══════════════════════════════════════════════════════════════
prs = nova_apresentacao()

# ── CAPA ────────────────────────────────────────────────────────
slide_capa(prs)

# ── SLIDE 1: Resumo Estatístico Geral ───────────────────────────
headers_resumo = ["Variável", "Média", "Mediana", "90%", "95%", "99%", "Máximo"]
rows_resumo = [
    ["Tração máx. [tf]",    "25,15", "23,56", "44,73", "47,25", "51,82", "56,44"],
    ["% MBL",               "33,54", "31,45", "59,61", "63,00", "69,10", "75,30"],
    ["Reação máx. [tf]",    "17,43", "16,61", "28,48", "32,41", "43,72", "51,86"],
    ["Compressão máx. [m]", "0,04",  "0,04",  "0,07",  "0,08",  "0,11",  "0,13"],
    ["Cabeço máx. [tf]",    "34,50", "31,88", "48,56", "55,63", "68,17", "83,13"],
    ["% SWL",               "46,00", "42,50", "64,77", "74,17", "90,91", "110,80"],
]
slide_tabela(
    prs,
    titulo="Resumo Estatístico Geral",
    subtitulo="Comportamento macro do sistema de amarração — 394 condições simuladas",
    headers=headers_resumo, rows=rows_resumo,
    txt_o_que=(
        "Apresenta uma síntese quantitativa dos esforços estruturais do sistema de "
        "amarração: tração nos cabos (tf e % MBL), reação e compressão nas defensas, "
        "e força nos cabeços (tf e % SWL).\n\n"
        "O objetivo é oferecer uma visão panorâmica do nível de solicitação que o "
        "sistema experimenta ao longo de todas as condições meteorológicas simuladas, "
        "com destaque para os percentis elevados (90%, 95%, 99%) que representam os "
        "eventos mais extremos."
    ),
    txt_resultado=(
        "A tração mediana nos cabos é de 23,6 tf (31,5% MBL), indicando que o sistema "
        "opera com folga na maior parte do tempo.\n\n"
        "Contudo, no percentil 99%, a tração alcança 51,8 tf (69,1% MBL), acima do "
        "limite operacional recomendado pela OCIMF de 55% MBL.\n\n"
        "Para os cabeços, o valor máximo de 110,8% SWL sinaliza que 2 condições "
        "ultrapassaram o limite estrutural absoluto de 100% SWL."
    )
)

# ── SLIDE 2: Histogramas MBL e SWL ──────────────────────────────
slide_padrao(
    prs,
    titulo="Distribuição de Frequência — Cabos e Cabeços",
    subtitulo="Histogramas de % MBL (Cabos) e % SWL (Cabeços) | 394 condições simuladas",
    txt_o_que=(
        "Os histogramas mostram com que frequência o sistema de amarração operou em "
        "cada faixa de carregamento ao longo das 394 condições simuladas.\n\n"
        "O eixo X representa o nível de utilização da capacidade (% MBL para cabos, "
        "% SWL para cabeços) e o eixo Y a quantidade de ocorrências.\n\n"
        "A linha tracejada vermelha indica os limites operacionais: 55% MBL para "
        "cabos (recomendação OCIMF) e 100% SWL para cabeços."
    ),
    txt_resultado=(
        "A maioria das condições (92 de 394) situa-se na faixa de 10–20% MBL, "
        "confirmando que o regime habitual é de baixa solicitação.\n\n"
        "Entretanto, 58 condições (≈ 15% do total) ultrapassam a linha de 55% MBL, "
        "configurando risco operacional.\n\n"
        "Para os cabeços, a distribuição é mais concentrada entre 20–60% SWL, com "
        "apenas 2 condições acima de 100% SWL — situações de risco estrutural severo."
    ),
    img_label="Histogramas: Distribuição (% MBL) — Cabos | Distribuição (% SWL) — Cabeços"
)

# ── SLIDE 3: Influência da Maré ──────────────────────────────────
slide_padrao(
    prs,
    titulo="Influência da Maré nos Esforços",
    subtitulo="Scatter Plot: Pico de Tensão e Esforço × Nível da Maré",
    txt_o_que=(
        "Gráfico de dispersão que correlaciona o nível da maré (codificado pela escala "
        "de cores fria-quente) com o pico de tensão (% MBL) em cada cabo e o pico de "
        "esforço (% SWL) em cada cabeço.\n\n"
        "Permite identificar se o estado de maré é um fator agravante dos esforços, "
        "revelando tendências de concentração de carga em determinados elementos "
        "estruturais conforme o nível d'água varia."
    ),
    txt_resultado=(
        "O Cabo 9 concentra consistentemente as tensões mais elevadas, com valores "
        "acima de 65% MBL, independentemente do nível de maré.\n\n"
        "Marés mais altas (tons avermelhados) tendem a coincidir com picos mais "
        "críticos, sugerindo que a geometria das linhas em maré alta penaliza "
        "principalmente o cabo 9.\n\n"
        "Para os cabeços, o Cabeço S é o elemento crítico, alcançando 110% SWL "
        "em condição de maré alta (1,58 m)."
    ),
    img_label="Scatter: Influência da Maré — Cabos (% MBL) | Cabeços (% SWL)"
)

# ── SLIDE 4: Análise de Risco — Tabela de Combinações Críticas ──
headers_risco = ["Cond.", "Vento [nos]", "Corrente [nos]", "Hs [m]", "Tração [tf]", "% MBL / % SWL", "Elemento Crítico"]
rows_risco = [
    ["259", "8,2",  "1,57", "1,00", "56,44 tf", "75,3% MBL", "Cabo 9"],
    ["310", "41,4", "1,16", "1,19", "54,72 tf", "73,0% MBL", "Cabo 9"],
    ["78",  "5,3",  "1,43", "0,35", "53,95 tf", "71,9% MBL", "Cabo 9"],
    ["282", "29,4", "1,58", "1,08", "51,82 tf", "69,1% MBL", "Cabo 9"],
    ["272", "39,9", "0,47", "1,05", "51,85 tf", "69,1% MBL", "Cabo 9"],
    ["198", "36,7", "1,58", "0,78", "83,13 tf", "110,8% SWL","Cabeço S"],
    ["55",  "49,1", "0,88", "0,30", "78,34 tf", "104,5% SWL","Cabeço S"],
]
slide_tabela(
    prs,
    titulo="Filtro de Riscos — Combinações Críticas",
    subtitulo="Condições que ultrapassam 55% MBL (cabos) e 100% SWL (cabeços)",
    headers=headers_risco, rows=rows_risco,
    txt_o_que=(
        "Filtra e classifica as condições meteoceanográficas que violam as margens de "
        "segurança estabelecidas (55% MBL para cabos e 100% SWL para cabeços), "
        "isolando os cenários de risco iminente.\n\n"
        "O objetivo é identificar quais combinações de vento, corrente e onda "
        "representam risco de ruptura de cabo ou arrancamento de cabeço, fornecendo "
        "os dados de entrada para a definição dos limites operacionais do berço."
    ),
    txt_resultado=(
        "Das 394 condições avaliadas, 58 ultrapassam 55% MBL nos cabos e apenas "
        "2 ultrapassam 100% SWL nos cabeços.\n\n"
        "O Cabo 9 é o elemento mais vulnerável: aparece em TODOS os 5 piores cenários "
        "de tração, com pico de 75,3% MBL na Cond. 259 (corrente de 1,57 nós, "
        "Hs = 1,0 m).\n\n"
        "O Cabeço S ultrapassa o limite estrutural em 2 condições, com força máxima "
        "de 83,1 tf (110,8% SWL) na Cond. 198."
    )
)

# ── SLIDE 5: Matriz de Correlação + Boxplot ──────────────────────
slide_padrao(
    prs,
    titulo="Inteligência de Dados — Diagnóstico de Carga",
    subtitulo="Matriz de Correlação (Ambiente × Esforços) e Boxplot por Cabo Crítico",
    txt_o_que=(
        "Dois diagnósticos complementares:\n\n"
        "① Matriz de Correlação — Mede a relação linear entre as variáveis "
        "ambientais (vento, corrente, Hs, Tp) e os esforços (% MBL, % SWL). "
        "Valores próximos de ±1 indicam dependência forte.\n\n"
        "② Boxplot por Cabo — Exibe a distribuição estatística (mediana, quartis "
        "e outliers) da tensão (% MBL) em cada cabo, ordenados do mais para o menos "
        "crítico, revelando assimetrias no plano de linhas."
    ),
    txt_resultado=(
        "A correlação mais relevante é entre % MBL e % SWL (0,70), sinalizando "
        "que os mesmos eventos que tensionam os cabos também sobrecarregam os "
        "cabeços.\n\n"
        "O vento apresenta correlação moderada com % SWL (0,25), mas correlações "
        "baixas com % MBL (0,02), sugerindo que a corrente e a onda são forças "
        "dominantes nos cabos.\n\n"
        "No Boxplot, o Cabo 9 tem mediana e amplitude significativamente maiores "
        "que os demais, indicando assimetria no plano de linhas: este cabo absorve "
        "desproporcionalmente a carga dinâmica do sistema."
    ),
    img_label="Esquerda: Matriz de Correlação | Direita: Boxplot % MBL por Cabo Crítico"
)

# ── SLIDE 6: Envelope Crítico ────────────────────────────────────
slide_padrao(
    prs,
    titulo="Envelope Crítico — Vento × Corrente",
    subtitulo="Mapa de esforços com codificação por cor e tamanho (nível de maré)",
    txt_o_que=(
        "Gráfico de dispersão bidimensional com eixos Corrente [nós] × Vento [nós]. "
        "Cada ponto representa um elemento estrutural (cabo ou cabeço) em determinada "
        "combinação ambiental.\n\n"
        "A cor do ponto (escala fria→quente) indica a intensidade do esforço "
        "(% MBL ou % SWL), e o tamanho do ponto representa o nível de maré — "
        "pontos maiores correspondem a marés mais elevadas.\n\n"
        "O objetivo é construir o 'envelope de paragem': a região do espaço "
        "vento-corrente onde o sistema ultrapassa os limites de segurança."
    ),
    txt_resultado=(
        "O ponto mais crítico dos cabos (vermelho intenso, maré alta) aparece em "
        "corrente ≈ 1,6 nós e vento ≈ 50 nós, com tensão acima de 70% MBL.\n\n"
        "Para os cabeços, o Cabeço S ultrapassa 100% SWL em corrente ≈ 1,6 nós "
        "e vento ≈ 37 nós — condição de maré alta.\n\n"
        "A maré alta amplifica os esforços ao alterar a geometria das linhas, "
        "aumentando o ângulo vertical e reduzindo a eficiência horizontal dos cabos."
    ),
    img_label="Envelope Crítico — Cabos (% MBL) | Envelope Crítico — Cabeços (% SWL)"
)

# ── SLIDE 7: Matriz de Operabilidade ────────────────────────────
slide_padrao(
    prs,
    titulo="Matriz de Operabilidade (Limites de Paragem)",
    subtitulo="Mapa de calor: Tensão máxima (% MBL) por Faixa de Vento × Faixa de Corrente",
    txt_o_que=(
        "Transforma os dados de 394 simulações em um mapa de calor de fácil "
        "leitura operacional, cruzando faixas de vento (eixo Y, em intervalos "
        "de 10 nós) com faixas de corrente (eixo X, em intervalos de 0,5 nós).\n\n"
        "Cada célula exibe a pior tensão (% MBL) registrada naquela combinação. "
        "A escala de cores Verde → Amarelo → Vermelho indica o grau de criticidade, "
        "com o limite de paragem fixado em 55% MBL."
    ),
    txt_resultado=(
        "Todas as células da matriz apresentam valores acima de 55% MBL, indicando "
        "que o Cabo 9 ultrapassa o limite operacional em praticamente qualquer "
        "combinação de vento e corrente.\n\n"
        "A faixa (0–10 nós vento / 1,5–2,0 nós corrente) registra o maior pico "
        "(75,3% MBL), mostrando que a corrente elevada é o principal fator de risco "
        "independentemente da velocidade do vento.\n\n"
        "A única célula próxima ao verde (38,8% MBL) ocorre com vento 10–20 nós "
        "e corrente 1,5–2,0 nós — uma exceção que requer investigação da geometria."
    ),
    img_label="Matriz de Operabilidade — Tensão Máxima (% MBL) por Vento e Corrente"
)

# ── SLIDE 8: Geometria das Linhas — Calado 9m ───────────────────
headers_geo = ["Fairlead", "Cabeço", "Comp. 3D [m]", "Âng. Vertical [°]"]
rows_geo_9m_alta = [
    ["2",  "S", "24,18", "15,55°"],
    ["3",  "S", "23,68", "15,88°"],
    ["14", "A", "31,27", "11,96°"],
    ["15", "A", "30,66", "12,20°"],
    ["17", "B", "17,60", "21,61°"],
    ["18", "B", "23,04", "16,34°"],
    ["23", "G", "41,76",  "8,93°"],
    ["24", "G", "41,02",  "9,09°"],
    ["28", "N", "31,00", "12,07°"],
    ["29", "N", "40,44",  "9,22°"],
    ["31", "S", "31,37", "11,92°"],
    ["32", "S", "30,29", "12,35°"],
]
slide_tabela(
    prs,
    titulo="Geometria das Linhas de Amarração — Calado 9 m | Maré Alta (1,58 m)",
    subtitulo="Comprimento 3D e ângulo vertical calculados para o cenário mais crítico",
    headers=headers_geo, rows=rows_geo_9m_alta,
    txt_o_que=(
        "Calcula analiticamente a geometria espacial (comprimento 3D e ângulo "
        "vertical) de cada linha de amarração para um dado estado de calado e maré.\n\n"
        "O ângulo vertical é determinante para a eficiência da linha: ângulos "
        "maiores que ~20° reduzem significativamente a componente horizontal "
        "de retenção do navio, aumentando a tração necessária para resistir "
        "às forças ambientais.\n\n"
        "A análise é repetida para 4 cenários: calado 9 m e 14 m, com maré alta "
        "(1,58 m) e sem maré (0 m)."
    ),
    txt_resultado=(
        "No cenário crítico (calado 9 m, maré 1,58 m), o Fairlead 17 → Cabeço B "
        "apresenta o maior ângulo vertical: 21,6°, sinalizando que esta linha "
        "trabalha em condição geométrica desfavorável.\n\n"
        "As linhas para o Cabeço G (Fairleads 23 e 24) têm os maiores comprimentos "
        "(≈ 41,8 m) e os menores ângulos (≈ 9°), indicando configuração mais eficiente.\n\n"
        "Com calado 14 m e maré zero, os ângulos caem drasticamente para < 1°, "
        "o que pode gerar esforços horizontais excessivos por falta de componente "
        "vertical de absorção."
    )
)

# ── SLIDE 9: Comparativo de Geometria entre Calados ─────────────
headers_comp = ["Fairlead", "Cabeço",
                "Âng. 9m/1,58m", "Âng. 9m/0m",
                "Âng. 14m/1,58m", "Âng. 14m/0m"]
rows_comp = [
    ["2→3",   "S", "~15,7°", "~12,0°", "~3,7°",  "~0,25°"],
    ["14→15", "A", "~12,1°", "~9,2°",  "~2,8°",  "~0,19°"],
    ["17→18", "B", "~19,0°", "~14,6°", "~4,5°",  "~0,35°"],
    ["23→24", "G",  "~9,0°", "~6,8°",  "~2,1°",  "~0,26°"],
    ["28→29", "N", "~10,6°", "~8,1°",  "~2,5°",  "~0,21°"],
    ["31→32", "S", "~12,1°", "~9,2°",  "~2,8°",  "~0,19°"],
]
slide_tabela(
    prs,
    titulo="Comparativo Geométrico — Calados 9 m vs. 14 m",
    subtitulo="Variação do ângulo vertical conforme calado e maré",
    headers=headers_comp, rows=rows_comp,
    txt_o_que=(
        "Consolida os resultados geométricos dos quatro cenários avaliados "
        "(calado 9 m e 14 m × maré alta 1,58 m e maré zero), permitindo "
        "comparar diretamente como a variação de calado e maré altera a "
        "inclinação vertical das linhas de amarração.\n\n"
        "A análise é fundamental para o planejamento operacional do berço, "
        "pois indica em quais condições a geometria das linhas é mais ou "
        "menos favorável à absorção de esforços horizontais."
    ),
    txt_resultado=(
        "A transição do calado 9 m para 14 m reduz os ângulos verticais em "
        "cerca de 75–95%, aproximando as linhas de uma configuração horizontal.\n\n"
        "Com calado 14 m e maré zero, todos os ângulos ficam abaixo de 0,4°, "
        "indicando linhas praticamente paralelas ao cais — situação que minimiza "
        "a capacidade de amortecimento vertical e pode favorecer movimentos de "
        "surgimento (surge) do navio.\n\n"
        "Recomenda-se revisão do plano de amarração para o cenário de calado "
        "máximo, com ênfase nas linhas de vante (Cabeços A e B)."
    )
)

# ── SLIDE 10: Sumário de Conclusões ─────────────────────────────
blank = prs.slide_layouts[6]
slide_conc = prs.slides.add_slide(blank)
fundo_slide(slide_conc, CINZA_CLARO)
barra_titulo(slide_conc, "Conclusões e Recomendações", "Síntese dos resultados da análise de amarração — Calado 9 m")

conclusoes = [
    ("⚠  Elemento Crítico Identificado",
     "O Cabo 9 e o Cabeço S são os elementos mais solicitados do sistema. "
     "O Cabo 9 aparece em todas as 5 piores condições de tração (até 75,3% MBL). "
     "O Cabeço S ultrapassou o limite estrutural absoluto (100% SWL) em 2 condições."),
    ("📊  Perfil Operacional",
     "O sistema opera com segurança na maior parte do tempo: mediana de 31,5% MBL. "
     "Contudo, 58 condições (≈15% do total) violam o limite operacional de 55% MBL, "
     "configurando um portfólio de cenários que requer atenção no planejamento."),
    ("🌊  Fatores Dominantes",
     "A corrente elevada (> 1,4 nós) é o principal fator de risco independente do vento. "
     "A maré alta (1,58 m) amplifica os esforços ao alterar a geometria das linhas, "
     "especialmente nas linhas curtas (Fairlead 17 → Cabeço B, ângulo 21,6°)."),
    ("📐  Geometria das Linhas",
     "Com calado 14 m e maré zero, os ângulos verticais caem para < 0,4°, "
     "comprometendo a eficiência de absorção vertical. "
     "Recomenda-se revisão do plano de amarração para calados superiores a 12 m."),
]

col_w = Inches(7.2)
col_h = Inches(3.3)
positions = [
    (Inches(0.35), Inches(1.35)),
    (Inches(8.05), Inches(1.35)),
    (Inches(0.35), Inches(4.90)),
    (Inches(8.05), Inches(4.90)),
]
cores_fundo = [
    RGBColor(0xFF, 0xF0, 0xD0),   # amarelo suave
    AZUL_CLARO,
    RGBColor(0xE8, 0xF5, 0xEB),   # verde suave
    RGBColor(0xF0, 0xE8, 0xF5),   # lilás suave
]
cores_titulo = [
    RGBColor(0xB8, 0x76, 0x10),
    AZUL_MEDIO,
    VERDE,
    RGBColor(0x70, 0x40, 0xA0),
]

for i, (titulo_c, texto_c) in enumerate(conclusoes):
    left, top = positions[i]
    caixa(slide_conc, texto_c,
          left, top, col_w, col_h,
          fundo=cores_fundo[i],
          titulo_texto=titulo_c,
          titulo_cor=cores_titulo[i],
          tamanho_titulo=12, tamanho_texto=10.5)

# ── RODAPÉ em todos os slides de conteúdo ───────────────────────
for i, slide in enumerate(prs.slides):
    if i == 0:
        continue
    rodape = slide.shapes.add_textbox(Inches(0.35), Inches(8.65), Inches(15), Inches(0.28))
    tf = rodape.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = f"Análise de Amarração — Calado 9 m  |  Navio Panamax  |  394 Condições Simuladas  |  Slide {i + 1} / {len(prs.slides) - 1}"
    run.font.size = Pt(7.5)
    run.font.color.rgb = CINZA_MEDIO
    p.alignment = PP_ALIGN.CENTER

# ── SALVAR ───────────────────────────────────────────────────────
output_path = "/projects/sandbox/Apresentacao_Amarracao_Maritima.pptx"
prs.save(output_path)
print(f"✅  Apresentação salva em: {output_path}")
