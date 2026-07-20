"""
Generador de la presentación final del trabajo de investigación MCP.
Formato: .pptx — editable directo en PowerPoint / Google Slides.

Cubre ambas partes del curso (pauta: "Presentación... ambas partes del curso
y comparativa"):
  - Parte 1 (Informe 1): método exacto — ILP resuelto por Branch & Bound.
  - Parte 2 (Informe 2): metaheurística — Búsqueda Tabú con reinicio (ILS).
  - Contraste explícito de resultados entre ambos enfoques.

Estilo: colores USACH (azul oscuro + rojo) sobre fondo blanco.
Duración objetivo: ~20 minutos (~21 slides, ritmo de ~57 s/slide).
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pathlib import Path

# ============================================================
# CONFIGURACIÓN
# ============================================================
OUT_DIR = Path(__file__).parent
LOGO_PNG = OUT_DIR / "logo_usach.png"
OUTPUT = OUT_DIR / "Presentacion_MCP.pptx"

NB1 = OUT_DIR.parent / "notebook"
NB2 = OUT_DIR.parent / "notebook_metaheuristica"

# Colores USACH
AZUL_USACH = RGBColor(0x00, 0x36, 0x7C)
ROJO_USACH = RGBColor(0xDA, 0x29, 0x1C)
GRIS_TEXTO = RGBColor(0x33, 0x33, 0x33)
GRIS_CLARO = RGBColor(0xEA, 0xEA, 0xEA)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)

# Tamaño slide 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


# ============================================================
# HELPERS
# ============================================================
def add_header_bar(slide, title_text, page_num=None, total=None):
    """Barra superior azul con título + número de página."""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.9))
    bar.fill.solid()
    bar.fill.fore_color.rgb = AZUL_USACH
    bar.line.fill.background()

    strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.9), SLIDE_W, Inches(0.08))
    strip.fill.solid()
    strip.fill.fore_color.rgb = ROJO_USACH
    strip.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.12), SLIDE_W - Inches(2), Inches(0.7))
    tf = tb.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Calibri"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = BLANCO

    if page_num is not None:
        pn = slide.shapes.add_textbox(SLIDE_W - Inches(1.5), Inches(0.25), Inches(1.2), Inches(0.4))
        ptf = pn.text_frame
        pp = ptf.paragraphs[0]
        pp.text = f"{page_num} / {total}"
        pp.alignment = PP_ALIGN.RIGHT
        pp.font.name = "Calibri"
        pp.font.size = Pt(12)
        pp.font.color.rgb = BLANCO


def add_page_number(slide, page_num, total, color=BLANCO):
    """Número de página en la esquina inferior derecha (para slides sin add_header_bar:
    portada, divisores de sección, cierre)."""
    pn = slide.shapes.add_textbox(SLIDE_W - Inches(1.2), SLIDE_H - Inches(0.5), Inches(1.0), Inches(0.35))
    tf = pn.text_frame
    p = tf.paragraphs[0]
    p.text = f"{page_num} / {total}"
    p.alignment = PP_ALIGN.RIGHT
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.color.rgb = color


def add_footer(slide, text="Maximum Clique Problem — Optimización en Ingeniería — USACH"):
    fb = slide.shapes.add_textbox(Inches(0.5), SLIDE_H - Inches(0.4), SLIDE_W - Inches(1), Inches(0.3))
    tf = fb.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Calibri"
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = GRIS_TEXTO


def add_bullets(slide, items, left=Inches(0.7), top=Inches(1.4),
                width=None, height=None, font_size=20, line_spacing=1.3):
    if width is None:
        width = SLIDE_W - Inches(1.4)
    if height is None:
        height = SLIDE_H - top - Inches(0.6)
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(item, tuple):
            text, level = item
        else:
            text, level = item, 0
        p.text = ("•  " if level == 0 else "–  ") + text
        p.font.name = "Calibri"
        p.font.size = Pt(font_size if level == 0 else font_size - 2)
        p.font.color.rgb = GRIS_TEXTO
        p.level = level
        p.space_after = Pt(8)
        p.line_spacing = line_spacing


def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=GRIS_TEXTO,
                 align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.name = "Calibri"
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    return tb


def add_box_card(slide, title, content_lines, left, top, width, height,
                 accent_color=AZUL_USACH, content_font_size=14):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = GRIS_CLARO
    bg.line.color.rgb = accent_color
    bg.line.width = Pt(1.5)

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.12), height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent_color
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(0.15), width - Inches(0.4), Inches(0.45))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "Calibri"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = accent_color

    cb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(0.75), width - Inches(0.5), height - Inches(0.9))
    ctf = cb.text_frame
    ctf.word_wrap = True
    for i, line in enumerate(content_lines):
        p = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
        p.text = line
        p.font.name = "Calibri"
        p.font.size = Pt(content_font_size)
        p.font.color.rgb = GRIS_TEXTO
        p.space_after = Pt(4)


def make_table(slide, data, left, top, width, height, header_color=AZUL_USACH, font_size=14):
    rows = len(data)
    cols = len(data[0])
    tbl = slide.shapes.add_table(rows, cols, left, top, width, height).table
    for ri, row in enumerate(data):
        for ci, cell_text in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = str(cell_text)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(font_size)
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(font_size)
                    if ri == 0:
                        run.font.bold = True
                        run.font.color.rgb = BLANCO
                    else:
                        run.font.color.rgb = GRIS_TEXTO
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_color
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = BLANCO if ri % 2 else GRIS_CLARO
    return tbl


def add_image(slide, path, left, top, width=None, height=None):
    if not Path(path).exists():
        print(f"  [aviso] imagen no encontrada: {path}")
        return None
    kwargs = {}
    if width is not None:
        kwargs["width"] = width
    if height is not None:
        kwargs["height"] = height
    return slide.shapes.add_picture(str(path), left, top, **kwargs)


def add_section_divider(slide, part_label, title_text, subtitle_text, accent_color=AZUL_USACH):
    """Slide de transición a pantalla completa (marca el cambio de parte)."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = accent_color
    bg.line.fill.background()

    strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.15), SLIDE_W, Inches(0.1))
    strip.fill.solid()
    strip.fill.fore_color.rgb = ROJO_USACH if accent_color == AZUL_USACH else AZUL_USACH
    strip.line.fill.background()

    add_text_box(slide, part_label, Inches(0.5), Inches(2.4), SLIDE_W - Inches(1), Inches(0.6),
                 font_size=22, bold=True, color=RGBColor(0xCF, 0xD8, 0xE8), align=PP_ALIGN.CENTER)
    add_text_box(slide, title_text, Inches(0.5), Inches(3.0), SLIDE_W - Inches(1), Inches(1.1),
                 font_size=40, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
    add_text_box(slide, subtitle_text, Inches(1.5), Inches(4.4), SLIDE_W - Inches(3), Inches(0.8),
                 font_size=18, italic=True, color=BLANCO, align=PP_ALIGN.CENTER)


# ============================================================
# CONSTRUCCIÓN DE LA PRESENTACIÓN
# ============================================================
prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]
TOTAL_SLIDES = 21
n = 0


def next_n():
    global n
    n += 1
    return n


# -------------------------------------------------------------
# 1 — PORTADA
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
pn = next_n()

top_bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(2.5))
top_bg.fill.solid()
top_bg.fill.fore_color.rgb = AZUL_USACH
top_bg.line.fill.background()

strip = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.5), SLIDE_W, Inches(0.12))
strip.fill.solid()
strip.fill.fore_color.rgb = ROJO_USACH
strip.line.fill.background()

add_image(s, LOGO_PNG, Inches(0.5), Inches(0.4), height=Inches(1.7))

add_text_box(s, "UNIVERSIDAD DE SANTIAGO DE CHILE", Inches(3), Inches(0.6), Inches(10), Inches(0.5),
             font_size=20, bold=True, color=BLANCO)
add_text_box(s, "Facultad de Ingeniería — Departamento de Ingeniería Informática",
             Inches(3), Inches(1.1), Inches(10), Inches(0.4), font_size=14, color=BLANCO, italic=True)
add_text_box(s, "Magíster en Ingeniería Informática", Inches(3), Inches(1.5), Inches(10), Inches(0.4),
             font_size=14, color=BLANCO)

add_text_box(s, "Maximum Clique Problem", Inches(0.5), Inches(2.9), SLIDE_W - Inches(1), Inches(0.9),
             font_size=44, bold=True, color=AZUL_USACH, align=PP_ALIGN.CENTER)
add_text_box(s, "Método exacto, metaheurística y contraste de resultados",
             Inches(0.5), Inches(3.75), SLIDE_W - Inches(1), Inches(0.5),
             font_size=22, color=GRIS_TEXTO, italic=True, align=PP_ALIGN.CENTER)
add_text_box(s, "Optimización en Ingeniería  —  Informes 1 y 2",
             Inches(0.5), Inches(4.3), SLIDE_W - Inches(1), Inches(0.4),
             font_size=16, color=ROJO_USACH, bold=True, align=PP_ALIGN.CENTER)

add_text_box(s, "Ricardo Riveros  •  Gonzalo Ahumada  •  Daniel Muñoz",
             Inches(0.5), Inches(5.6), SLIDE_W - Inches(1), Inches(0.4),
             font_size=18, color=GRIS_TEXTO, align=PP_ALIGN.CENTER)
add_text_box(s, "Profesores: Mónica Villanueva, Mario Inostroza  —  Ayudante: Sebastián Aliaga",
             Inches(0.5), Inches(6.0), SLIDE_W - Inches(1), Inches(0.4),
             font_size=14, color=GRIS_TEXTO, italic=True, align=PP_ALIGN.CENTER)
add_text_box(s, "Santiago, julio 2026", Inches(0.5), Inches(6.5), SLIDE_W - Inches(1), Inches(0.4),
             font_size=14, color=GRIS_TEXTO, align=PP_ALIGN.CENTER)
add_page_number(s, pn, TOTAL_SLIDES, color=GRIS_TEXTO)

# -------------------------------------------------------------
# 2 — AGENDA
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Agenda", next_n(), TOTAL_SLIDES)
add_bullets(s, [
    "El problema: Maximum Clique Problem (MCP) y su complejidad",
    "Contextos de aplicación y objetivos",
    ("Parte 1 — Método exacto: formulación ILP, Branch & Bound, resultados", 0),
    ("Parte 2 — Metaheurística: diseño, Búsqueda Tabú + reinicio, resultados", 0),
    "Contraste de resultados: calidad, tiempo y el compromiso garantía–escalabilidad",
    "Conclusiones y trabajo futuro",
], top=Inches(1.6), font_size=22, line_spacing=1.5)
add_footer(s)

# -------------------------------------------------------------
# 3 — EL PROBLEMA Y SU COMPLEJIDAD
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "El problema: Maximum Clique Problem (MCP)", next_n(), TOTAL_SLIDES)
add_text_box(
    s,
    "Dado un grafo no dirigido G = (V, E), un clique es un subconjunto C ⊆ V donde "
    "todos los vértices están mutuamente conectados. El MCP busca el clique de mayor "
    "cardinalidad:  ω(G) = |C*|",
    Inches(0.7), Inches(1.35), SLIDE_W - Inches(1.4), Inches(1.1),
    font_size=18, color=GRIS_TEXTO,
)
add_box_card(
    s, "Complejidad computacional",
    [
        "NP-duro (Bomze et al., 1999; Pardalos & Xue, 1994).",
        "No aproximable en tiempo polinomial dentro de un factor",
        "constante, salvo que P = NP.",
        "Resoluble en tiempo polinomial solo en clases especiales",
        "(grafos perfectos, bipartitos, de intervalos).",
    ],
    Inches(0.6), Inches(2.7), Inches(5.9), Inches(3.9), accent_color=ROJO_USACH,
)
add_box_card(
    s, "Equivalencias clásicas",
    [
        "Clique máximo en G = Independent Set máximo",
        "en el complemento de G.",
        "",
        "Vertex Cover mínimo en G = V menos el",
        "Independent Set máximo (mismo grafo G).",
        "",
        "Esta dualidad es la base de varias familias",
        "de algoritmos exactos y heurísticos.",
    ],
    Inches(6.8), Inches(2.7), Inches(5.9), Inches(3.9), accent_color=AZUL_USACH, content_font_size=13,
)
add_footer(s)

# -------------------------------------------------------------
# 4 — CONTEXTOS DE APLICACIÓN
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Contextos de aplicación", next_n(), TOTAL_SLIDES)
add_box_card(
    s, "Teoría de códigos correctores de errores",
    [
        "Grafo G_{n,d}: vértices = palabras binarias de longitud n;",
        "aristas = pares con distancia de Hamming ≥ d.",
        "Clique máximo en G_{n,d} = código óptimo de tamaño A(n,d).",
        "",
        "Familia Keller (usada en este trabajo): teselaciones de",
        "R^n por hipercubos unitarios trasladados.",
    ],
    Inches(0.5), Inches(1.4), Inches(6.2), Inches(5.4), accent_color=AZUL_USACH,
)
add_box_card(
    s, "Bioinformática — Alineamiento molecular",
    [
        "Subestructura común máxima (MCS) entre dos moléculas.",
        "Grafo de productos G_P: vértices = pares (a,b) atómicamente",
        "compatibles; aristas = distancias internas consistentes.",
        "",
        "Aplicación: docking molecular, descubrimiento de fármacos,",
        "identificación de farmacóforos comunes.",
    ],
    Inches(7), Inches(1.4), Inches(5.8), Inches(5.4), accent_color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# 5 — OBJETIVOS Y PREGUNTA DE INVESTIGACIÓN
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Objetivos y pregunta de investigación", next_n(), TOTAL_SLIDES)
add_box_card(
    s, "Objetivo general",
    [
        "Resolver el MCP mediante dos paradigmas —un método exacto (ILP + Branch & Bound)",
        "y una metaheurística de búsqueda local (Búsqueda Tabú + reinicio)— y contrastar",
        "sus resultados en términos de calidad de solución y tiempo de cómputo.",
    ],
    Inches(0.5), Inches(1.4), SLIDE_W - Inches(1), Inches(1.6), accent_color=AZUL_USACH,
    content_font_size=16,
)
add_text_box(s, "Objetivos específicos", Inches(0.7), Inches(3.2), SLIDE_W - Inches(1.4), Inches(0.4),
             font_size=18, bold=True, color=AZUL_USACH)
add_bullets(s, [
    "Formular e implementar el ILP del MCP y resolverlo con HiGHS sobre keller4 (DIMACS).",
    "Diseñar e implementar una Búsqueda Tabú con reinicio sobre la vecindad Add/Drop/Swap.",
    "Evaluar ambos enfoques sobre instancias de distinto tamaño y densidad.",
    "Contrastar sistemáticamente calidad, tiempo y garantías de optimalidad.",
], top=Inches(3.65), font_size=16, line_spacing=1.2)
add_box_card(
    s, "Pregunta de investigación",
    [
        "¿En qué medida una metaheurística de búsqueda local relativamente simple iguala la calidad",
        "de un método exacto para el MCP mientras reduce drásticamente el tiempo de cómputo,",
        "y qué se sacrifica —en términos de garantías— al hacerlo?",
    ],
    Inches(0.5), Inches(5.7), SLIDE_W - Inches(1), Inches(1.3), accent_color=ROJO_USACH,
    content_font_size=15,
)
add_footer(s)

# -------------------------------------------------------------
# 6 — DIVISOR: PARTE 1
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
pn = next_n()
add_section_divider(
    s, "PARTE 1 · INFORME 1", "Método Exacto",
    "Programación Lineal Entera (ILP) resuelta por Ramificación y Acotamiento",
    accent_color=AZUL_USACH,
)
add_page_number(s, pn, TOTAL_SLIDES)

# -------------------------------------------------------------
# 7 — FORMULACIÓN ILP
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Formulación del MCP como ILP", next_n(), TOTAL_SLIDES)
add_text_box(s, "Variable de decisión:  xᵢ = 1 ⇔ vértice i ∈ clique,  xᵢ = 0 en caso contrario",
             Inches(0.7), Inches(1.45), SLIDE_W - Inches(1.4), Inches(0.5), font_size=17, color=GRIS_TEXTO)

fo_bg = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(2.15), Inches(12), Inches(0.9))
fo_bg.fill.solid()
fo_bg.fill.fore_color.rgb = GRIS_CLARO
fo_bg.line.color.rgb = AZUL_USACH
fo_bg.line.width = Pt(1.5)
add_text_box(s, "Función objetivo:   máx   z = Σᵢ∈V  xᵢ", Inches(0.7), Inches(2.3), Inches(12), Inches(0.6),
             font_size=22, bold=True, color=AZUL_USACH, align=PP_ALIGN.CENTER)

add_text_box(s, "Restricciones de adyacencia:", Inches(0.7), Inches(3.35), Inches(6), Inches(0.4),
             font_size=16, bold=True, color=GRIS_TEXTO)
add_text_box(s, "xᵢ + xⱼ ≤ 1     ∀ (i,j) ∈ Ē      (Ē = pares NO adyacentes)",
             Inches(0.7), Inches(3.75), Inches(12), Inches(0.5), font_size=18, color=GRIS_TEXTO)
add_text_box(s, "Restricciones de integralidad:", Inches(0.7), Inches(4.4), Inches(6), Inches(0.4),
             font_size=16, bold=True, color=GRIS_TEXTO)
add_text_box(s, "xᵢ ∈ {0, 1}     ∀ i ∈ V", Inches(0.7), Inches(4.8), Inches(12), Inches(0.5),
             font_size=18, color=GRIS_TEXTO)

add_box_card(
    s, "Resolución",
    [
        "Python + Pyomo (modelado algebraico) + HiGHS (solver MIP — Mixed Integer",
        "Programming —: Branch & Bound con planos de corte y heurísticas de factibilidad internas).",
    ],
    Inches(0.7), Inches(5.6), SLIDE_W - Inches(1.4), Inches(1.3), accent_color=ROJO_USACH,
    content_font_size=15,
)
add_footer(s)

# -------------------------------------------------------------
# 8 — RESULTADOS keller4: HiGHS vs GLPK
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Resultados: keller4 (171v, 9 435 aristas, ω=11)", next_n(), TOTAL_SLIDES)
data = [
    ["Métrica", "HiGHS", "GLPK"],
    ["Mejor solución (ω)", "11", "11"],
    ["Tiempo", "12–18 s", "> 300 s (timeout)"],
    ["Estado", "Óptimo probado", "Factible, sin probar optimalidad"],
]
make_table(s, data, Inches(0.9), Inches(1.5), Inches(11.5), Inches(1.9))
add_box_card(
    s, "Hallazgo clave",
    [
        "Ambos solvers encuentran ω = 11, pero solo HiGHS certifica optimalidad en tiempo razonable.",
        "La diferencia se explica por preprocesamiento, cortes (Gomory/MIR), heurísticas de",
        "factibilidad y ramificación adaptativa — técnicas ausentes en GLPK (GNU Linear Programming Kit).",
    ],
    Inches(0.9), Inches(3.7), Inches(11.5), Inches(1.6), accent_color=ROJO_USACH, content_font_size=15,
)
add_image(s, NB1 / "resultados_keller4_pyomo.png", Inches(4.85), Inches(5.35), height=Inches(1.55))
add_footer(s)

# -------------------------------------------------------------
# 9 — RELAJACIÓN LP
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Relajación LP y brecha de integralidad", next_n(), TOTAL_SLIDES)
add_text_box(s, "Se reemplaza xᵢ ∈ {0,1} por xᵢ ∈ [0,1] y se resuelve el LP (Linear Programming) continuo.",
             Inches(0.7), Inches(1.45), SLIDE_W - Inches(1.4), Inches(0.5), font_size=16, italic=True, color=GRIS_TEXTO)
data = [
    ["Métrica", "Valor"],
    ["Cota LP (relajación)", "85.50"],
    ["Solución entera (ILP)", "11"],
    ["Brecha de integralidad", "87.1 %"],
]
make_table(s, data, Inches(2.7), Inches(2.15), Inches(7.9), Inches(1.6))
add_box_card(
    s, "Implicación",
    [
        "La cota LP es muy floja (87.1 % de brecha): no captura la combinatoria del problema.",
        "→ El solver debe explorar un árbol B&B extenso para cerrar la brecha.",
        "Esta debilidad de la formulación estándar del MCP motiva explorar métodos alternativos —",
        "como la metaheurística de la Parte 2— para instancias donde el árbol se vuelve inviable.",
    ],
    Inches(1.2), Inches(4.1), Inches(10.9), Inches(2.2), accent_color=AZUL_USACH, content_font_size=15,
)
add_footer(s)

# -------------------------------------------------------------
# 10 — DIVISOR: PARTE 2
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
pn = next_n()
add_section_divider(
    s, "PARTE 2 · INFORME 2", "Metaheurística",
    "Búsqueda Tabú con reinicio por perturbación (estilo Búsqueda Local Iterada)",
    accent_color=AZUL_USACH,
)
add_page_number(s, pn, TOTAL_SLIDES)

# -------------------------------------------------------------
# 11 — DISEÑO: REPRESENTACIÓN, OBJETIVO, VECINDAD (TAREA 1)
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Diseño de la metaheurística (base: Tarea 1)", next_n(), TOTAL_SLIDES)
add_box_card(
    s, "Representación y objetivo",
    [
        "S ⊆ V, siempre un clique válido de G (factibilidad por construcción).",
        "f(S) = |S|,  a maximizar.",
        "C(S) = vértices fuera de S adyacentes a TODO S (candidatos a agregar).",
    ],
    Inches(0.5), Inches(1.4), SLIDE_W - Inches(1), Inches(1.9), accent_color=AZUL_USACH, content_font_size=16,
)
add_text_box(s, "Vecindad N(S): tres movimientos, todos cerrados en el espacio de cliques",
             Inches(0.7), Inches(3.5), SLIDE_W - Inches(1.4), Inches(0.5), font_size=18, bold=True, color=AZUL_USACH)
add_bullets(s, [
    ("Add(v): agrega v ∈ C(S) — incrementa f en 1.", 0),
    ("Drop(u): elimina u ∈ S — decrementa f en 1.", 0),
    ("Swap(u,v): intercambia u por v — mantiene f, cambia C(S) (permite escapar de óptimos locales).", 0),
], top=Inches(4.05), font_size=17, line_spacing=1.3)
add_box_card(
    s, "Complejidad",
    ["Cálculo de N(S) vectorizado con matriz de adyacencia: O(n·k), n=|V|, k=|S|."],
    Inches(0.7), Inches(6.0), SLIDE_W - Inches(1.4), Inches(0.9), accent_color=ROJO_USACH, content_font_size=15,
)
add_footer(s)

# -------------------------------------------------------------
# 12 — ALGORITMO: BÚSQUEDA TABÚ + REINICIO
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Algoritmo: Búsqueda Tabú con reinicio (ILS)", next_n(), TOTAL_SLIDES)
add_bullets(s, [
    "En cada iteración, se prioriza en orden:",
    ("Si C(S) ≠ ∅ → aplicar Add (siempre mejorante, casi nunca se rechaza).", 1),
    ("Si S es maximal pero hay Swap disponible → aplicar el que maximiza |C(S')|.", 1),
    ("Si no hay Add ni Swap → forzar un Drop aleatorio (diversificación).", 1),
    "Lista tabú: un vértice recién eliminado no puede re-agregarse durante tenure iteraciones.",
    "Reinicio (ILS): si el incumbente no mejora en stall_limit iteraciones, se elimina una fracción aleatoria de S (10–30 %) y se reconstruye de forma voraz.",
], top=Inches(1.5), font_size=18, line_spacing=1.3)
add_box_card(
    s, "Parámetros usados",
    ["tenure = 10        stall_limit = 200        perturbación ∈ [0.1, 0.3]        max_iter = 8 000"],
    Inches(0.7), Inches(6.0), SLIDE_W - Inches(1.4), Inches(0.9), accent_color=AZUL_USACH, content_font_size=16,
)
add_footer(s)

# -------------------------------------------------------------
# 13 — INSTANCIAS Y PROTOCOLO EXPERIMENTAL
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Instancias y protocolo experimental", next_n(), TOTAL_SLIDES)
data = [
    ["Instancia", "|V|", "|E|", "Densidad", "ω(G) conocido"],
    ["keller4", "171", "9 435", "0.6491", "11"],
    ["C125.9", "125", "6 963", "0.8985", "34"],
]
make_table(s, data, Inches(1.3), Inches(1.5), Inches(10.7), Inches(1.6))
add_text_box(
    s,
    "C125.9 se incorpora respecto al Informe 1 por ser considerablemente más densa: un Branch & Bound "
    "manual (PyBnB) no logró certificar optimalidad en 300 s sobre esta instancia (232 953 nodos explorados).",
    Inches(0.9), Inches(3.4), SLIDE_W - Inches(1.8), Inches(1.0), font_size=16, italic=True, color=GRIS_TEXTO,
)
add_box_card(
    s, "Protocolo",
    [
        "30 corridas independientes por instancia (semillas 0–29).",
        "Se registra: tamaño final, iteración/tiempo hasta la mejor solución, N° de perturbaciones.",
        "Cada solución se valida verificando C(|S*|,2) aristas en el subgrafo inducido.",
    ],
    Inches(0.9), Inches(4.6), SLIDE_W - Inches(1.8), Inches(1.9), accent_color=ROJO_USACH, content_font_size=16,
)
add_footer(s)

# -------------------------------------------------------------
# 14 — RESULTADOS AGREGADOS
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Resultados agregados (30 corridas)", next_n(), TOTAL_SLIDES)
data = [
    ["Métrica", "keller4", "C125.9"],
    ["Tasa de éxito (f(S*)=ω)", "93.3 % (28/30)", "96.7 % (29/30)"],
    ["Tamaño medio ± desv. estándar", "10.90 ± 0.40", "33.97 ± 0.18"],
    ["Tiempo medio hasta mejor solución", "0.0211 s", "0.0251 s"],
    ["Iteraciones medias hasta mejor solución", "556.3", "632.3"],
]
make_table(s, data, Inches(1.1), Inches(1.5), Inches(11.1), Inches(2.3))
add_box_card(
    s, "Lectura",
    [
        "En ambas instancias, el óptimo se alcanza en la gran mayoría de las corridas; cuando no,",
        "la solución queda a 1-2 vértices de distancia. El tiempo hasta la mejor solución es",
        "del orden de centésimas de segundo — prácticamente insensible a la dificultad de la instancia.",
    ],
    Inches(1.1), Inches(4.1), Inches(11.1), Inches(1.7), accent_color=AZUL_USACH, content_font_size=16,
)
add_footer(s)

# -------------------------------------------------------------
# 15 — ABLACIÓN
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Experimento de ablación: efecto del mecanismo de reinicio", next_n(), TOTAL_SLIDES)
add_text_box(
    s, "Comparación controlada: Búsqueda Tabú pura (stall_limit→∞) vs. con reinicio tipo ILS, mismas 30 semillas.",
    Inches(0.7), Inches(1.45), SLIDE_W - Inches(1.4), Inches(0.6), font_size=16, italic=True, color=GRIS_TEXTO,
)
data = [
    ["Variante", "Tasa éxito keller4", "Tamaño medio", "Tasa éxito C125.9", "Tamaño medio"],
    ["Sin reinicio (Tabú puro)", "73.3 %", "10.57 ± 0.76", "66.7 %", "33.40 ± 1.11"],
    ["Con reinicio (Tabú+ILS)", "93.3 %", "10.90 ± 0.40", "96.7 %", "33.97 ± 0.18"],
]
make_table(s, data, Inches(0.6), Inches(2.3), Inches(12.1), Inches(1.6), font_size=13)
add_box_card(
    s, "Conclusión del experimento de ablación",
    [
        "El reinicio por perturbación mejora la tasa de éxito en 20–30 puntos porcentuales y reduce",
        "la desviación estándar del tamaño final en 48 % (keller4) y 84 % (C125.9). El efecto es mayor",
        "en C125.9 (más densa): sin diversificación explícita, el Tabú puro queda atrapado en plateaus.",
    ],
    Inches(0.6), Inches(4.3), Inches(12.1), Inches(1.9), accent_color=ROJO_USACH, content_font_size=15,
)
add_footer(s)

# -------------------------------------------------------------
# 16 — CONVERGENCIA + CLIQUE ENCONTRADO (slide fusionada, colchón de tiempo)
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Convergencia y clique encontrado", next_n(), TOTAL_SLIDES)
add_text_box(
    s,
    "Patrón característico de plateaus: expansión voraz, estancamiento en Swap y salto al óptimo tras una perturbación.",
    Inches(0.5), Inches(0.98), SLIDE_W - Inches(1), Inches(0.35), font_size=13, italic=True, color=GRIS_TEXTO,
    align=PP_ALIGN.CENTER,
)

COL_W = Inches(6.15)
COL1_X = Inches(0.3)
COL2_X = Inches(6.88)
CONV_W = Inches(3.42)
CLIQUE_H = Inches(2.75)

add_text_box(s, "keller4 (ω = 11)", COL1_X, Inches(1.38), COL_W, Inches(0.32),
             font_size=15, bold=True, color=AZUL_USACH, align=PP_ALIGN.CENTER)
add_text_box(s, "C125.9 (ω = 34)", COL2_X, Inches(1.38), COL_W, Inches(0.32),
             font_size=15, bold=True, color=AZUL_USACH, align=PP_ALIGN.CENTER)

add_image(s, NB2 / "convergencia_keller4.png",
          COL1_X + Inches((6.15 - 3.42) / 2), Inches(1.72), width=CONV_W)
add_image(s, NB2 / "convergencia_C125.9.png",
          COL2_X + Inches((6.15 - 3.42) / 2), Inches(1.72), width=CONV_W)

add_image(s, NB2 / "grafo_clique_keller4_tabu.png",
          COL1_X + Inches((6.15 - 2.75) / 2), Inches(4.02), height=CLIQUE_H)
add_image(s, NB2 / "grafo_clique_C125.9_tabu.png",
          COL2_X + Inches((6.15 - 2.75) / 2), Inches(4.02), height=CLIQUE_H)
add_footer(s)

# -------------------------------------------------------------
# 17 — DIVISOR: CONTRASTE
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
pn = next_n()
add_section_divider(
    s, "COMPARATIVA · INFORME 2", "Contraste de Resultados",
    "Método exacto (Informe 1) vs. metaheurística (Informe 2)",
    accent_color=AZUL_USACH,
)
add_page_number(s, pn, TOTAL_SLIDES)

# -------------------------------------------------------------
# 18 — TABLA COMPARATIVA
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Contraste: método exacto vs. metaheurística", next_n(), TOTAL_SLIDES)
data = [
    ["Instancia", "Método", "ω", "Tiempo", "¿Óptimo probado?"],
    ["keller4", "ILP + HiGHS (Informe 1)", "11", "12–18 s", "Sí"],
    ["keller4", "Tabú + ILS (Informe 2)", "11", "0.021 s", "No (28/30 óptimas)"],
    ["C125.9", "B&B manual, PyBnB", "34", "300 s (timeout)", "No"],
    ["C125.9", "Tabú + ILS (Informe 2)", "34", "0.025 s", "No (29/30 óptimas)"],
]
make_table(s, data, Inches(0.6), Inches(1.5), Inches(12.1), Inches(2.5))
add_box_card(
    s, "Nota metodológica",
    [
        "HiGHS combina B&B con cortes y heurísticas (B&B asistido); el B&B manual de C125.9 es",
        "Ramificación y Acotamiento pura (solo relajación LP + branching), la comparación más limpia.",
    ],
    Inches(0.6), Inches(4.3), Inches(12.1), Inches(1.2), accent_color=ROJO_USACH, content_font_size=14,
)
add_footer(s)

# -------------------------------------------------------------
# 19 — TIEMPO DE CÓMPUTO Y TRADE-OFF
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Tiempo de cómputo: el compromiso garantía–escalabilidad", next_n(), TOTAL_SLIDES)
add_image(s, NB2 / "comparacion_tiempos.png", Inches(3.4), Inches(1.25), height=Inches(4.1))
add_box_card(
    s, "Hallazgo central",
    [
        "En keller4 la metaheurística es ~711× más rápida que HiGHS; en C125.9, ~11 900× más rápida",
        "que el B&B manual (que ni siquiera certifica optimalidad). El método exacto certifica el óptimo",
        "cuando logra cerrar el árbol, pero su costo crece sin control con la dificultad de la instancia;",
        "la metaheurística es casi insensible a esa dificultad, a costa de no ofrecer esa garantía.",
    ],
    Inches(0.7), Inches(5.45), SLIDE_W - Inches(1.4), Inches(1.55), accent_color=AZUL_USACH, content_font_size=14,
)
add_footer(s)

# -------------------------------------------------------------
# 20 — CONCLUSIONES Y TRABAJO FUTURO
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Conclusiones y trabajo futuro", next_n(), TOTAL_SLIDES)
add_text_box(s, "Conclusiones principales", Inches(0.7), Inches(1.4), Inches(8), Inches(0.5),
             font_size=20, bold=True, color=AZUL_USACH)
add_bullets(s, [
    "El método exacto certifica ω=11 en keller4 (12–18 s), pero su costo crece sin control en instancias densas.",
    "La metaheurística Tabú+ILS, construida sobre el diseño de Tarea 1, alcanza el óptimo en 93–97 % de las corridas.",
    "El mecanismo de reinicio (ILS) es indispensable: sin él, la tasa de éxito cae 20–30 puntos.",
    "El contraste confirma el compromiso teórico garantía vs. escalabilidad entre ambos paradigmas.",
], top=Inches(1.9), font_size=16, line_spacing=1.15)

add_text_box(s, "Trabajo futuro", Inches(0.7), Inches(4.75), Inches(8), Inches(0.5),
             font_size=20, bold=True, color=ROJO_USACH)
add_bullets(s, [
    "Evaluar sobre instancias DIMACS de mayor tamaño (n > 500).",
    "Comparar contra VNS (Hansen & Mladenović, 2007) sobre las mismas vecindades Add/Drop/Swap.",
    "Ajuste reactivo del tenure; certificar C125.9 con un solver MIP formal (HiGHS).",
], top=Inches(5.2), font_size=16, line_spacing=1.15)
add_footer(s)

# -------------------------------------------------------------
# 21 — CIERRE
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
pn = next_n()
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
bg.fill.solid()
bg.fill.fore_color.rgb = AZUL_USACH
bg.line.fill.background()

strip = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.0), SLIDE_W, Inches(0.15))
strip.fill.solid()
strip.fill.fore_color.rgb = ROJO_USACH
strip.line.fill.background()

add_text_box(s, "¿Preguntas?", Inches(0.5), Inches(2.5), SLIDE_W - Inches(1), Inches(1.2),
             font_size=72, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
add_text_box(s, "Gracias por su atención", Inches(0.5), Inches(4.3), SLIDE_W - Inches(1), Inches(0.6),
             font_size=28, italic=True, color=BLANCO, align=PP_ALIGN.CENTER)
add_text_box(s, "Ricardo Riveros  •  Gonzalo Ahumada  •  Daniel Muñoz",
             Inches(0.5), Inches(5.3), SLIDE_W - Inches(1), Inches(0.5),
             font_size=18, color=BLANCO, align=PP_ALIGN.CENTER)
add_text_box(s, "Optimización en Ingeniería — USACH 2026", Inches(0.5), Inches(5.8), SLIDE_W - Inches(1), Inches(0.5),
             font_size=14, italic=True, color=BLANCO, align=PP_ALIGN.CENTER)
add_page_number(s, pn, TOTAL_SLIDES)

# ============================================================
# GUARDAR
# ============================================================
assert n == TOTAL_SLIDES, f"Slides construidos ({n}) != TOTAL_SLIDES ({TOTAL_SLIDES})"
prs.save(str(OUTPUT))
print(f"OK: presentación guardada en {OUTPUT}")
print(f"    {len(prs.slides)} slides")
