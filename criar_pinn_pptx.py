import sys
sys.path.insert(0, '/projects/sandbox/pptx_pkgs')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ══════════════════════════════════════════════════════════════
#  PALETA — Fiel ao PDF original (creme, bege, tipografia escura)
# ══════════════════════════════════════════════════════════════
CREME       = RGBColor(0xF5, 0xF0, 0xE8)   # fundo principal
CREME_CARD  = RGBColor(0xFF, 0xFF, 0xFF)   # card branco
CREME_CARD2 = RGBColor(0xF0, 0xEB, 0xE2)   # card bege levemente mais escuro
TEXTO_ESC   = RGBColor(0x1A, 0x1F, 0x2E)   # texto principal escuro
TEXTO_MED   = RGBColor(0x4A, 0x4A, 0x5A)   # texto secundário
BORDA_CARD  = RGBColor(0xD8, 0xD0, 0xC4)   # borda suave
# Cores de destaque (usadas como barra esquerda e títulos de seção)
LARANJA     = RGBColor(0xD4, 0x7B, 0x2A)   # laranja quente (step labels)
TEAL        = RGBColor(0x1A, 0x7A, 0x78)   # teal (subtítulos verdes)
AZUL_ESC    = RGBColor(0x1B, 0x3A, 0x6B)   # azul escuro (títulos principais)
ROXO        = RGBColor(0x5B, 0x3D, 0xA0)   # roxo (step 6 results)
AMARELO_D   = RGBColor(0xB8, 0x8A, 0x00)   # dourado (step labels)
BRANCO      = RGBColor(0xFF, 0xFF, 0xFF)
CINZA_CLARO = RGBColor(0xE8, 0xE4, 0xDC)   # linha divisória

W = Inches(16)
H = Inches(9)

# ══════════════════════════════════════════════════════════════
#  HELPERS BASE
# ══════════════════════════════════════════════════════════════
def nova_prs():
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs

def rect(slide, l, t, w, h, fill, line_rgb=None, line_pt=0.75, radius=False):
    from pptx.util import Pt
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line_rgb:
        s.line.color.rgb = line_rgb
        s.line.width = Pt(line_pt)
    else:
        s.line.fill.background()
    return s

def tb(slide, texto, l, t, w, h,
       size=11, bold=False, cor=TEXTO_ESC,
       align=PP_ALIGN.LEFT, italic=False, wrap=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text  = texto
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = cor
    return box

def tb_multi(slide, linhas, l, t, w, h,
             size=10.5, cor=TEXTO_MED, bold_primeiro=False, espaco=True):
    """Adiciona múltiplas linhas com bullet • em um textbox"""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = True
    for i, linha in enumerate(linhas):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(4) if espaco and i > 0 else Pt(0)
        run = p.add_run()
        run.text = linha
        run.font.size = Pt(size)
        run.font.color.rgb = cor
        run.font.bold = bold_primeiro and i == 0
    return box

def fundo(slide, cor=CREME):
    rect(slide, Inches(0), Inches(0), W, H, cor)

def barra_lateral(slide, cor=LARANJA, espessura=Inches(0.09)):
    """Barra colorida na lateral esquerda — elemento icônico do template"""
    rect(slide, Inches(0), Inches(0), espessura, H, cor)

def num_slide(slide, n, total, cor_barra=TEXTO_MED):
    """Número de slide no canto inferior direito"""
    tb(slide, f"{n} / {total}",
       Inches(14.5), Inches(8.55), Inches(1.3), Inches(0.35),
       size=9, cor=cor_barra, align=PP_ALIGN.RIGHT)

def step_label(slide, texto, cor=LARANJA):
    """Label 'STEP X · NOME' no topo esquerdo"""
    tb(slide, texto,
       Inches(0.25), Inches(0.18), Inches(10), Inches(0.30),
       size=8.5, bold=True, cor=cor)

def titulo_principal(slide, texto, cor=TEXTO_ESC, size=36):
    tb(slide, texto,
       Inches(0.25), Inches(0.55), Inches(13.5), Inches(1.1),
       size=size, bold=False, cor=cor)

def subtitulo_slide(slide, texto, cor=TEXTO_MED, size=12):
    tb(slide, texto,
       Inches(0.25), Inches(1.55), Inches(13.5), Inches(0.4),
       size=size, cor=cor)

def card_branco(slide, l, t, w, h, raio=False):
    """Card branco com borda suave — estilo do PDF"""
    r = rect(slide, l, t, w, h, CREME_CARD, line_rgb=BORDA_CARD, line_pt=0.8)
    return r

def card_bege(slide, l, t, w, h):
    r = rect(slide, l, t, w, h, CREME_CARD2, line_rgb=BORDA_CARD, line_pt=0.8)
    return r

TOTAL = 20

# ══════════════════════════════════════════════════════════════
#  SLIDE 1 — CAPA
# ══════════════════════════════════════════════════════════════
def s01_capa(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)

    # círculo decorativo canto inferior direito (imita o PDF)
    circ = slide.shapes.add_shape(9,  # OVAL
        Inches(12.5), Inches(6.0), Inches(4.0), Inches(4.0))
    circ.fill.solid()
    circ.fill.fore_color.rgb = RGBColor(0xEE, 0xE5, 0xD8)
    circ.line.fill.background()

    # Conteúdo
    tb(slide, "PME5242 - Physics-Informed Machine Learning",
       Inches(0.55), Inches(1.8), Inches(12), Inches(0.4),
       size=11, bold=True, cor=TEXTO_MED)

    tb(slide, "Physics Informed Neural Network\nfor vessel manoeuvring",
       Inches(0.55), Inches(2.4), Inches(12), Inches(2.2),
       size=42, bold=True, cor=TEXTO_ESC)

    tb(slide, "Vessel: Panamax (LOA 225 m)",
       Inches(0.55), Inches(5.05), Inches(10), Inches(0.45),
       size=16, bold=True, cor=TEXTO_ESC)

    tb(slide, "Gustavo Guilherme Silva  ·  NUSP: 11820555\nGabriel Testte Leite  ·  NUSP: 11259288",
       Inches(0.55), Inches(5.65), Inches(10), Inches(0.75),
       size=12, cor=TEXTO_MED)

    num_slide(slide, 1, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 2 — Motivation and Objective
# ══════════════════════════════════════════════════════════════
def s02_motivation(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)
    step_label(slide, "STEP 1  ·  MOTIVATION")
    titulo_principal(slide, "Motivation and objective")

    # Card esquerdo
    card_branco(slide, Inches(0.25), Inches(1.90), Inches(7.0), Inches(5.6))

    tb(slide, "Accurate trajectory prediction is vital to prevent maritime accidents and support terminal operators.",
       Inches(0.45), Inches(2.05), Inches(6.6), Inches(0.75),
       size=12, cor=TEXTO_MED)

    # Bullet 1
    tb(slide, "The Objective:",
       Inches(0.55), Inches(3.00), Inches(6.3), Inches(0.30),
       size=12, bold=True, cor=TEXTO_ESC)
    tb(slide, "Develop a Physics-Informed Neural Network (PINN) to predict 3-DOF maneuvering states (surge, sway, yaw) for a 225m Panamax.",
       Inches(0.55), Inches(3.32), Inches(6.3), Inches(0.75),
       size=11.5, cor=TEXTO_MED)

    # Bullet 2
    tb(slide, "The Innovation:",
       Inches(0.55), Inches(4.25), Inches(6.3), Inches(0.30),
       size=12, bold=True, cor=TEXTO_ESC)
    tb(slide, "Embed the nonlinear Abkowitz model to ensure strict physical consistency, bridging classical hydrodynamics with machine learning.",
       Inches(0.55), Inches(4.57), Inches(6.3), Inches(0.75),
       size=11.5, cor=TEXTO_MED)

    # Placeholder imagem direita
    card_bege(slide, Inches(7.55), Inches(1.90), Inches(8.1), Inches(5.6))
    tb(slide, "[ Imagem: acidente marítimo ]",
       Inches(7.55), Inches(4.45), Inches(8.1), Inches(0.4),
       size=10, italic=True, cor=BORDA_CARD, align=PP_ALIGN.CENTER)

    num_slide(slide, 2, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 3 — Work Hypotheses
# ══════════════════════════════════════════════════════════════
def s03_hypotheses(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)
    step_label(slide, "STEP 1  ·  HYPOTHESES")
    titulo_principal(slide, "WORK HYPOTHESES")
    tb(slide, "Everything the code assumes",
       Inches(0.25), Inches(1.60), Inches(8), Inches(0.35),
       size=12, cor=TEXTO_MED, italic=True)

    hipoteses = [
        ("H1", "Rigid hull, 3-DOF",
         "Planar motion: surge, sway and yaw.\nIgnores roll, pitch and heave."),
        ("H2", "Quasi-steady fluid",
         "Forces depend only on the current state.\nWithout fluid memory."),
        ("H3", "Calm and Deep Waters",
         "No wind, wave or seabed effects.\nOperational simplification."),
        ("H4", "Constant propulsion",
         "Only the deviation Δu relative to U₀ is modeled."),
    ]

    posicoes = [
        (Inches(0.25), Inches(2.05)),
        (Inches(4.55), Inches(2.05)),
        (Inches(0.25), Inches(5.00)),
        (Inches(4.55), Inches(5.00)),
    ]

    for (hid, htit, hcorp), (l, t) in zip(hipoteses, posicoes):
        card_branco(slide, l, t, Inches(4.0), Inches(2.65))
        tb(slide, hid,
           l + Inches(0.18), t + Inches(0.18), Inches(1.5), Inches(0.28),
           size=9, bold=True, cor=TEAL)
        tb(slide, htit,
           l + Inches(0.18), t + Inches(0.48), Inches(3.6), Inches(0.32),
           size=12, bold=True, cor=TEXTO_ESC)
        tb(slide, hcorp,
           l + Inches(0.18), t + Inches(0.85), Inches(3.6), Inches(1.5),
           size=10.5, cor=TEXTO_MED)

    # Placeholder imagem 3D DOF
    card_bege(slide, Inches(9.0), Inches(2.05), Inches(6.6), Inches(5.6))
    tb(slide, "[ Diagrama: 6 DOF — surge, sway, yaw, roll, pitch, heave ]",
       Inches(9.0), Inches(4.6), Inches(6.6), Inches(0.4),
       size=10, italic=True, cor=BORDA_CARD, align=PP_ALIGN.CENTER)

    tb(slide, "Based on: Abkowitz (1964); Fossen (2011)  |  Bäckström & Skårbratt (2010)",
       Inches(0.25), Inches(8.50), Inches(12), Inches(0.3),
       size=8.5, cor=TEXTO_MED)
    num_slide(slide, 3, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 4 — Abkowitz 3-DOF Model
# ══════════════════════════════════════════════════════════════
def s04_abkowitz(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)
    step_label(slide, "STEP 2  ·  PHYSICAL MODEL")
    titulo_principal(slide, "Abkowitz 3-DOF Model")
    tb(slide, "Nondimensional Newton-Euler model: fluid forces expanded via Taylor series as continuous functions of vessel velocities, accelerations, and rudder angle.",
       Inches(0.25), Inches(1.58), Inches(15), Inches(0.55),
       size=11.5, cor=TEXTO_MED)

    # Card 1 — Equação matricial
    card_branco(slide, Inches(0.25), Inches(2.30), Inches(15.3), Inches(2.25))
    tb(slide, "1",
       Inches(0.45), Inches(2.45), Inches(0.4), Inches(0.4),
       size=22, bold=True, cor=TEAL)
    tb(slide, "Matrix form of the 3-DOF equations of motion (Newton-Euler, nondimensional):",
       Inches(1.00), Inches(2.45), Inches(14), Inches(0.35),
       size=10.5, cor=TEXTO_MED)
    tb(slide, "[ M' - M_A' ] · { u̇', v̇', ṙ' }ᵀ  =  { X'(u',v',r',δ) + m'(v'r' + xG'r'²),  Y'(u',v',r',δ) − m'u'r',  N'(u',v',r',δ) − m'xG'u'r' }ᵀ",
       Inches(1.00), Inches(2.90), Inches(14), Inches(1.2),
       size=11, bold=True, cor=TEXTO_ESC)

    # Card 2 — Expansão Taylor
    card_branco(slide, Inches(0.25), Inches(4.75), Inches(15.3), Inches(2.60))
    tb(slide, "2",
       Inches(0.45), Inches(4.90), Inches(0.4), Inches(0.4),
       size=22, bold=True, cor=TEAL)
    tb(slide, "Taylor series expansion example for longitudinal force (surge):",
       Inches(1.00), Inches(4.90), Inches(14), Inches(0.35),
       size=10.5, cor=TEXTO_MED)
    tb(slide, "X = X₀' + Xᵤ'Δu' + Xᵤᵤ'Δu'² + Xᵤᵤᵤ'Δu'³ + Xᵥᵥ'v'² + Xᵣᵣ'r'² + Xδδ'δ² + Xᵥᵣ'v'r' + Xᵥδ'v'δ + Xᵣδ'r'δ",
       Inches(1.00), Inches(5.35), Inches(14), Inches(0.65),
       size=13, bold=True, cor=TEXTO_ESC)
    tb(slide, "Model Characteristics: High-fidelity physical representation, although requiring the identification of a large number of hydrodynamic coefficients to fully define X, Y, and N.",
       Inches(1.00), Inches(6.10), Inches(14), Inches(0.85),
       size=10.5, cor=TEXTO_MED)

    tb(slide, "Based on: Abkowitz (1964, 1980); Fossen (2011)",
       Inches(0.25), Inches(8.50), Inches(10), Inches(0.3),
       size=8.5, cor=TEXTO_MED)
    num_slide(slide, 4, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 5 — RK45 and Prime System
# ══════════════════════════════════════════════════════════════
def s05_rk45(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)
    step_label(slide, "STEP 2  ·  PHYSICAL MODEL")
    titulo_principal(slide, "RK45 and Prime System")

    # RK45
    tb(slide, "RK45:", Inches(0.25), Inches(1.60), Inches(4), Inches(0.35),
       size=13, bold=True, cor=TEAL)
    card_branco(slide, Inches(0.25), Inches(2.05), Inches(15.3), Inches(2.05))
    itens_rk = [
        "• Adaptive Time-Stepping:  Handles strong non-linearities while optimizing computational efficiency.",
        "• Full-Scale Physical Domain:  Integrates the equations of motion directly with physical quantities.",
        "• Precise Error Control:  Enables 4th and 5th-order error estimation to adjust time steps with high precision.",
    ]
    for i, item in enumerate(itens_rk):
        partes = item.split(":", 1)
        t_pos = Inches(2.15) + Pt(18) * i + Inches(0.55 * i)
        tb(slide, item,
           Inches(0.50), Inches(2.22) + Inches(0.58 * i),
           Inches(14.8), Inches(0.50),
           size=11, cor=TEXTO_MED)

    # Prime System
    tb(slide, "Prime System", Inches(0.25), Inches(4.30), Inches(5), Inches(0.35),
       size=13, bold=True, cor=TEAL)
    card_branco(slide, Inches(0.25), Inches(4.75), Inches(15.3), Inches(1.55))
    tb(slide, "• Input Scaling & Gradient Stability:  Normalizes variables into dimensionless form to prevent gradient pathologies for the PINN.",
       Inches(0.50), Inches(4.90), Inches(14.8), Inches(0.45),
       size=11, cor=TEXTO_MED)
    tb(slide, "• Numerical Conditioning:  Eliminates ill-conditioning caused by vast scale discrepancies.",
       Inches(0.50), Inches(5.38), Inches(14.8), Inches(0.40),
       size=11, cor=TEXTO_MED)

    # Fórmulas — 4 cards pequenos
    formulas = [
        ("Yaw rate",             "r' = r × L / U"),
        ("Dimensionless time",   "t' = t × U / L"),
        ("Translational u",      "u' = u / U"),
        ("Translational v",      "v' = v / U"),
    ]
    for i, (label, formula) in enumerate(formulas):
        cl = Inches(0.25) + Inches(3.8) * i
        card_branco(slide, cl, Inches(6.55), Inches(3.65), Inches(1.85))
        tb(slide, label, cl + Inches(0.15), Inches(6.65), Inches(3.35), Inches(0.30),
           size=9.5, bold=True, cor=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, formula, cl + Inches(0.15), Inches(7.05), Inches(3.35), Inches(0.55),
           size=14, bold=True, cor=TEXTO_ESC, align=PP_ALIGN.CENTER)

    tb(slide, "Based on: SNAME (1950), Prime System; Fossen (2011)",
       Inches(0.25), Inches(8.50), Inches(10), Inches(0.3),
       size=8.5, cor=TEXTO_MED)
    num_slide(slide, 5, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 6 — Synthetic Data (Turning Circle)
# ══════════════════════════════════════════════════════════════
def s06_synthetic_a(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)
    step_label(slide, "STEP 3  ·  SYNTHETIC DATA")
    titulo_principal(slide, "Synthetic Vessel Simulation")

    # Placeholder imagem turning circle
    card_branco(slide, Inches(0.25), Inches(1.80), Inches(8.60), Inches(6.70))
    tb(slide, "[ Imagem: Turning Circle — Tactical diameter, Advance, Radius, Drift angle ]",
       Inches(0.25), Inches(5.0), Inches(8.60), Inches(0.4),
       size=10, italic=True, cor=BORDA_CARD, align=PP_ALIGN.CENTER)

    # Card de texto
    card_branco(slide, Inches(9.15), Inches(1.80), Inches(6.5), Inches(3.80))
    tb(slide, "Trajectory:",
       Inches(9.35), Inches(2.00), Inches(6.0), Inches(0.30),
       size=12, bold=True, cor=TEXTO_ESC)
    tb(slide, "The vessel enters a transient phase upon rudder application, eventually stabilizing into a perfectly circular path.",
       Inches(9.35), Inches(2.35), Inches(6.0), Inches(0.80),
       size=11.5, cor=TEXTO_MED)
    tb(slide, "Disclaimer:",
       Inches(9.35), Inches(3.35), Inches(6.0), Inches(0.30),
       size=12, bold=True, cor=TEXTO_ESC)
    tb(slide, "This project uses a simple maneuver to study the development of the PINN.",
       Inches(9.35), Inches(3.70), Inches(6.0), Inches(0.65),
       size=11.5, cor=TEXTO_MED)

    num_slide(slide, 6, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 7 — Synthetic Data (Velocity Plots)
# ══════════════════════════════════════════════════════════════
def s07_synthetic_b(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)
    step_label(slide, "STEP 3  ·  SYNTHETIC DATA")
    titulo_principal(slide, "Synthetic Vessel Simulation")

    # 3 placeholders de gráficos
    graficos = [
        (Inches(0.25),  Inches(1.80), Inches(7.4),  Inches(3.20), "Longitudinal Velocity (Surge) — u [m/s] vs Time [s]"),
        (Inches(8.10),  Inches(1.80), Inches(7.4),  Inches(3.20), "Transverse Velocity (Sway) — v [m/s] vs Time [s]"),
        (Inches(3.80),  Inches(5.30), Inches(8.10), Inches(3.20), "Yaw Rate — r [°/s] vs Time [s]"),
    ]
    for l, t, w, h, label in graficos:
        card_branco(slide, l, t, w, h)
        tb(slide, label, l, t + h/2 - Inches(0.2), w, Inches(0.4),
           size=9, italic=True, cor=BORDA_CARD, align=PP_ALIGN.CENTER)

    num_slide(slide, 7, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 8 — Neural Network Architecture
# ══════════════════════════════════════════════════════════════
def s08_architecture(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, AZUL_ESC)
    step_label(slide, "STEP 4  ·  FORMULATION", cor=AZUL_ESC)
    titulo_principal(slide, "Neural network architecture")
    tb(slide, "The proposed neural network consists of three multilayer perceptrons (MLPs), one for each estimated variable.",
       Inches(0.25), Inches(1.58), Inches(12), Inches(0.45),
       size=11.5, cor=TEXTO_MED)

    # Placeholder diagrama NN
    card_branco(slide, Inches(0.25), Inches(2.15), Inches(10.5), Inches(6.35))
    tb(slide, "[ Diagrama: 3× MLP (1 → 16 → 16 → 1)  |  Entradas: t' adimensional  |  Saídas: u', v', r' ]",
       Inches(0.25), Inches(5.1), Inches(10.5), Inches(0.4),
       size=9.5, italic=True, cor=BORDA_CARD, align=PP_ALIGN.CENTER)
    tb(slide, "16 neurons          16 neurons",
       Inches(2.5), Inches(8.05), Inches(5), Inches(0.3),
       size=10, cor=TEXTO_MED, align=PP_ALIGN.CENTER)

    # Card de nota lateral
    card_branco(slide, Inches(11.1), Inches(2.15), Inches(4.65), Inches(2.40))
    tb(slide, "Additionally, the network includes a trainable parameter, δ, representing the rudder angle.",
       Inches(11.3), Inches(2.45), Inches(4.2), Inches(1.6),
       size=12, cor=TEXTO_MED)

    num_slide(slide, 8, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 9 — Boundary Conditions & Data Loss
# ══════════════════════════════════════════════════════════════
def s09_bc_loss(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, AZUL_ESC)
    step_label(slide, "STEP 4  ·  FORMULATION", cor=AZUL_ESC)

    # Dois títulos lado a lado
    tb(slide, "Boundary conditions", Inches(0.25), Inches(0.28), Inches(7.5), Inches(0.75),
       size=32, bold=False, cor=TEXTO_ESC)
    tb(slide, "Data loss", Inches(8.30), Inches(0.28), Inches(7.0), Inches(0.75),
       size=32, bold=False, cor=TEXTO_ESC)

    # Card esquerdo — Boundary
    card_branco(slide, Inches(0.25), Inches(1.20), Inches(7.65), Inches(7.30))
    tb(slide, "Initial conditions:",
       Inches(0.50), Inches(1.45), Inches(7.0), Inches(0.30),
       size=12, bold=True, cor=TEXTO_ESC)
    tb(slide, "The problem is formulated as an initial value problem governed by an ODE system. The vessel starts each maneuver with u' = 1, corresponding to y(0) = [1, 0, 0].",
       Inches(0.50), Inches(1.80), Inches(7.0), Inches(1.0),
       size=11, cor=TEXTO_MED)

    tb(slide, "Hard constraint implementation:",
       Inches(0.50), Inches(3.00), Inches(7.0), Inches(0.30),
       size=12, bold=True, cor=TEXTO_ESC)
    tb(slide, "Initial conditions are enforced through a hard constraint similar to boundary encoding:",
       Inches(0.50), Inches(3.35), Inches(7.0), Inches(0.60),
       size=11, cor=TEXTO_MED)

    # Fórmula hard constraint
    card_bege(slide, Inches(0.75), Inches(4.15), Inches(6.9), Inches(2.60))
    tb(slide, "y(t') = y₀ + (1 − e^{−t'}) · NN(−t')",
       Inches(0.75), Inches(5.15), Inches(6.9), Inches(0.55),
       size=18, bold=True, cor=TEXTO_ESC, align=PP_ALIGN.CENTER)
    tb(slide, "Distance function",
       Inches(0.75), Inches(5.80), Inches(6.9), Inches(0.35),
       size=10.5, bold=True, cor=TEAL, align=PP_ALIGN.CENTER)

    # Card direito — Data Loss
    card_branco(slide, Inches(8.30), Inches(1.20), Inches(7.35), Inches(7.30))
    tb(slide, "Sampling:",
       Inches(8.55), Inches(1.45), Inches(6.8), Inches(0.30),
       size=12, bold=True, cor=TEXTO_ESC)
    tb(slide, "30 data points are uniformly distributed within time windows at the beginning and end of the simulation. This approach aims to emulate AIS data sampling from vessels.",
       Inches(8.55), Inches(1.80), Inches(6.8), Inches(1.20),
       size=11, cor=TEXTO_MED)

    tb(slide, "Loss function:  Mean Squared Error (MSE)",
       Inches(8.55), Inches(3.25), Inches(6.8), Inches(0.45),
       size=12, bold=True, cor=TEXTO_ESC)

    card_bege(slide, Inches(8.80), Inches(3.90), Inches(6.6), Inches(2.20))
    tb(slide, "MSE = (1/N) × Σᵢ (yᵢ − ŷᵢ)²",
       Inches(8.80), Inches(4.75), Inches(6.6), Inches(0.55),
       size=18, bold=True, cor=TEXTO_ESC, align=PP_ALIGN.CENTER)

    num_slide(slide, 9, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 10 — Physics Informed Loss
# ══════════════════════════════════════════════════════════════
def s10_physics_loss(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, AZUL_ESC)
    step_label(slide, "STEP 4  ·  FORMULATION", cor=AZUL_ESC)
    titulo_principal(slide, "Physics informed loss")

    card_branco(slide, Inches(0.25), Inches(1.70), Inches(15.3), Inches(6.75))

    itens = [
        ("Physics-informed loss function:",
         "Mean squared error based on the acceleration equations of the maneuvering model used in the simulations."),
        ("Dimensional consistency:",
         "Network outputs and their derivatives are dimensionalized using variable-specific scaling factors to ensure physically consistent residuals."),
        ("Loss balancing:",
         "Weighting factors balance the contributions of the three equations of motion. The λ parameters are calculated based on the observed range of each variable."),
        ("Collocation point sampling:",
         "Random resampling at each training epoch considerably reduced the number of points needed to deliver great results. 3000 collocation points per epoch."),
    ]

    for i, (titulo, corpo) in enumerate(itens):
        y = Inches(1.90) + Inches(1.32) * i
        tb(slide, f"• {titulo}",
           Inches(0.55), y, Inches(14.5), Inches(0.32),
           size=12, bold=True, cor=TEXTO_ESC)
        tb(slide, f"  {corpo}",
           Inches(0.55), y + Inches(0.33), Inches(14.5), Inches(0.55),
           size=11, cor=TEXTO_MED)

    # Fórmula
    card_bege(slide, Inches(2.5), Inches(7.00), Inches(10.5), Inches(1.05))
    tb(slide, "ℒphys = λᵤ MSE(Rᵤ) + λᵥ MSE(Rᵥ) + λᵣ MSE(Rᵣ)",
       Inches(2.5), Inches(7.25), Inches(10.5), Inches(0.55),
       size=17, bold=True, cor=TEXTO_ESC, align=PP_ALIGN.CENTER)

    num_slide(slide, 10, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 11 — PINN Training
# ══════════════════════════════════════════════════════════════
def s11_training(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, AZUL_ESC)
    step_label(slide, "STEP 4  ·  FORMULATION", cor=AZUL_ESC)
    titulo_principal(slide, "PINN Training")
    tb(slide, "Two-stage training process:",
       Inches(0.25), Inches(1.60), Inches(8), Inches(0.35),
       size=12, cor=TEXTO_MED)

    # Card Estágio 1
    card_branco(slide, Inches(0.25), Inches(2.10), Inches(6.20), Inches(2.80))
    tb(slide, "Rudder Angle Identification",
       Inches(0.45), Inches(2.28), Inches(5.8), Inches(0.40),
       size=13, bold=True, cor=TEXTO_ESC)
    tb(slide, "• 30% of the training epochs\n• Loss function evaluated exclusively on data points",
       Inches(0.45), Inches(2.78), Inches(5.8), Inches(0.90),
       size=11, cor=TEXTO_MED)

    # Seta
    tb(slide, "→",
       Inches(6.65), Inches(3.20), Inches(0.8), Inches(0.6),
       size=28, bold=True, cor=TEXTO_ESC, align=PP_ALIGN.CENTER)

    # Card Estágio 2
    card_branco(slide, Inches(7.65), Inches(2.10), Inches(7.9), Inches(2.80))
    tb(slide, "Maneuvering Velocities Calibration",
       Inches(7.85), Inches(2.28), Inches(7.5), Inches(0.40),
       size=13, bold=True, cor=TEXTO_ESC)
    tb(slide, "• 70% of the training epochs\n• Loss function evaluated on both data points and collocation points",
       Inches(7.85), Inches(2.78), Inches(7.5), Inches(0.90),
       size=11, cor=TEXTO_MED)

    # Card observações
    card_branco(slide, Inches(0.25), Inches(5.20), Inches(15.3), Inches(3.10))
    tb(slide, "• Splitting the training process into two stages accelerated network convergence. Simultaneously optimizing the maneuvering velocities and rudder angle required significantly more epochs to achieve comparable performance.\n\n• Each training stage is optimized using Adam, followed by an additional 300-iteration optimization using L-BFGS at the end of the corresponding training epochs.",
       Inches(0.55), Inches(5.45), Inches(14.5), Inches(2.6),
       size=11.5, cor=TEXTO_MED)

    num_slide(slide, 11, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 12 — Analyzed Cases
# ══════════════════════════════════════════════════════════════
def s12_cases(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, AZUL_ESC)
    step_label(slide, "STEP 5  ·  IMPLEMENTATION", cor=AZUL_ESC)
    titulo_principal(slide, "Analyzed cases")
    tb(slide, "The neural network was trained using different configurations to evaluate its ability to learn the vessel's maneuvering dynamics.",
       Inches(0.25), Inches(1.58), Inches(15), Inches(0.55),
       size=11.5, cor=TEXTO_MED)

    # Tabela centralizada
    headers = ["ID", "Speed [kn]", "Rudder angle [deg]", "Data window", "Physics Loss"]
    rows = [
        ["1", "10", "30", "0:200 s",       "No"],
        ["2", "10", "30", "0:200 s",       "Yes"],
        ["3", "15", "10", "0:200 s",       "Yes"],
        ["4", "15", "10", "1800:2000 s",   "No"],
        ["5", "15", "10", "1800:2000 s",   "Yes"],
        ["6", "10", "30", "1800:2000 s",   "Yes"],
    ]

    tl = Inches(2.5)
    tt = Inches(2.50)
    tw = Inches(11.0)
    n_cols = len(headers)
    n_rows = len(rows) + 1
    tbl = slide.shapes.add_table(n_rows, n_cols, tl, tt, tw, Inches(0.48 * n_rows + 0.55)).table

    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = TEXTO_ESC
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(11)
        p.runs[0].font.color.rgb = BRANCO

    for ri, row in enumerate(rows):
        bg = CREME_CARD if ri % 2 == 0 else CREME_CARD2
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.runs[0].font.size = Pt(11)
            p.runs[0].font.color.rgb = TEXTO_ESC

    num_slide(slide, 12, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDES 13-18 — Resultados (Cases 1-6) — layout uniforme
# ══════════════════════════════════════════════════════════════
def slide_resultado(prs, num, titulo_case, cor_barra=ROXO):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, cor_barra)
    step_label(slide, "STEP 6  ·  RESULTS", cor=cor_barra)
    titulo_principal(slide, titulo_case, cor=TEXTO_ESC, size=30)

    # Placeholder para os 6 subgráficos (3 em cima, 3 embaixo)
    graficos_top = [
        "Longitudinal velocity u(t)\n[m/s] vs Time [s]\nRK45 | PINN | Train Data",
        "Transversal velocity v(t)\n[m/s] vs Time [s]\nRK45 | PINN | Train Data",
        "Yaw rate r(t)\n[°/s] vs Time [s]\nRK45 | PINN | Train Data",
    ]
    graficos_bot = [
        "Trajectory (2000 s)\nX [m] vs Y [m]\nRK45 | PINN | Train data",
        "Rudder angle history\nduring training\n[deg] vs Epoch",
        "PINN Absolute Error\n|Error| (log) vs Tempo [s]\n|Err u'| | |Err v'| | |Err r'|",
    ]
    w_g = Inches(4.85)
    h_g = Inches(3.25)
    for i, label in enumerate(graficos_top):
        l = Inches(0.25) + w_g * i + Inches(0.07) * i
        card_bege(slide, l, Inches(1.65), w_g, h_g)
        tb(slide, label, l, Inches(1.65) + h_g/2 - Inches(0.35),
           w_g, Inches(0.7), size=8.5, italic=True,
           cor=BORDA_CARD, align=PP_ALIGN.CENTER)
    for i, label in enumerate(graficos_bot):
        l = Inches(0.25) + w_g * i + Inches(0.07) * i
        card_bege(slide, l, Inches(5.15), w_g, h_g)
        tb(slide, label, l, Inches(5.15) + h_g/2 - Inches(0.35),
           w_g, Inches(0.7), size=8.5, italic=True,
           cor=BORDA_CARD, align=PP_ALIGN.CENTER)

    num_slide(slide, num, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 19 — Conclusions + Future Work
# ══════════════════════════════════════════════════════════════
def s19_conclusions(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, ROXO)
    step_label(slide, "STEP 6  ·  RESULTS", cor=ROXO)

    tb(slide, "Conclusions", Inches(0.25), Inches(0.28), Inches(7), Inches(0.65),
       size=34, bold=False, cor=TEXTO_ESC)

    card_branco(slide, Inches(0.25), Inches(1.10), Inches(15.3), Inches(3.30))
    conclusoes = [
        "• Developed a Physics-Informed Neural Network (PINN) for maneuverability estimation based on the 3-DOF Abkowitz model.",
        "• The physics-informed loss function required careful tuning, including an additional hyperparameter to balance the contributions of the governing equations.",
        "• The PINN achieved excellent results in parameter identification and variable extrapolation across all analyzed scenarios.",
    ]
    for i, c in enumerate(conclusoes):
        tb(slide, c,
           Inches(0.50), Inches(1.28) + Inches(0.95) * i,
           Inches(14.8), Inches(0.75),
           size=11.5, cor=TEXTO_MED)

    tb(slide, "Future work", Inches(0.25), Inches(4.75), Inches(7), Inches(0.65),
       size=34, bold=False, cor=TEXTO_ESC)

    card_branco(slide, Inches(0.25), Inches(5.55), Inches(15.3), Inches(2.85))
    futuros = [
        "• More complex manouvers: time varying rudder angle.",
        "• Neural operator mapping past/present state to future state.",
    ]
    for i, f in enumerate(futuros):
        tb(slide, f,
           Inches(0.50), Inches(5.75) + Inches(1.0) * i,
           Inches(14.8), Inches(0.65),
           size=11.5, cor=TEXTO_MED)

    num_slide(slide, 19, TOTAL)

# ══════════════════════════════════════════════════════════════
#  SLIDE 20 — Thank You
# ══════════════════════════════════════════════════════════════
def s20_thankyou(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fundo(slide)
    barra_lateral(slide, LARANJA)

    # Círculo decorativo
    circ = slide.shapes.add_shape(9,
        Inches(11.5), Inches(5.5), Inches(5.0), Inches(5.0))
    circ.fill.solid()
    circ.fill.fore_color.rgb = RGBColor(0xEE, 0xE5, 0xD8)
    circ.line.fill.background()

    tb(slide, "Thank you for the attention",
       Inches(0.55), Inches(2.2), Inches(12), Inches(1.3),
       size=46, bold=True, cor=TEXTO_ESC)

    tb(slide, "Any questions?",
       Inches(0.55), Inches(4.2), Inches(10), Inches(1.1),
       size=38, bold=True, cor=TEXTO_ESC)

    num_slide(slide, 20, TOTAL)

# ══════════════════════════════════════════════════════════════
#  CONSTRUÇÃO COMPLETA
# ══════════════════════════════════════════════════════════════
prs = nova_prs()

s01_capa(prs)
s02_motivation(prs)
s03_hypotheses(prs)
s04_abkowitz(prs)
s05_rk45(prs)
s06_synthetic_a(prs)
s07_synthetic_b(prs)
s08_architecture(prs)
s09_bc_loss(prs)
s10_physics_loss(prs)
s11_training(prs)
s12_cases(prs)

# Slides de resultados (Cases 1-6)
cases = [
    (13, "Case 1: U₀ = 10 kn, δ = 30°, w/o physics"),
    (14, "Case 2: U₀ = 10 kn, δ = 30°, with physics"),
    (15, "Case 3: U₀ = 15 kn, δ = 10°, with physics"),
    (16, "Case 4: U₀ = 15 kn, δ = 10°, w/o physics, end points"),
    (17, "Case 5: U₀ = 15 kn, δ = 10°, with physics, end points"),
    (18, "Case 6: U₀ = 10 kn, δ = 30°, with physics, end points"),
]
for num, titulo in cases:
    slide_resultado(prs, num, titulo)

s19_conclusions(prs)
s20_thankyou(prs)

OUTPUT = "/projects/sandbox/PINN_Vessel_Manoeuvring.pptx"
prs.save(OUTPUT)
print(f"✅  Salvo: {OUTPUT}")
print(f"📊  Slides: {len(prs.slides)}")
