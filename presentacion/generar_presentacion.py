"""
Generador de la presentación del trabajo de investigación MCP.
Formato: .pptx — importable directo a Google Slides.

Estilo: colores USACH (azul oscuro + rojo) sobre fondo blanco.
Duración objetivo: ~20 minutos (~17 slides).
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pathlib import Path

# ============================================================
# CONFIGURACIÓN
# ============================================================
OUT_DIR = Path(__file__).parent
LOGO_PNG = "/tmp/logo_usach-1.png"
OUTPUT = OUT_DIR / "Presentacion_MCP.pptx"

# Colores USACH
AZUL_USACH = RGBColor(0x00, 0x36, 0x7C)
ROJO_USACH = RGBColor(0xDA, 0x29, 0x1C)
GRIS_TEXTO = RGBColor(0x33, 0x33, 0x33)
GRIS_CLARO = RGBColor(0xEA, 0xEA, 0xEA)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)

# Tamaño slide 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def hex_color(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ============================================================
# HELPERS
# ============================================================
def add_header_bar(slide, title_text, page_num=None, total=None):
    """Barra superior azul con título + número de página."""
    # Barra azul
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.9)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = AZUL_USACH
    bar.line.fill.background()

    # Strip rojo delgado
    strip = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, Inches(0.9), SLIDE_W, Inches(0.08)
    )
    strip.fill.solid()
    strip.fill.fore_color.rgb = ROJO_USACH
    strip.line.fill.background()

    # Título dentro de la barra
    tb = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.15), SLIDE_W - Inches(2), Inches(0.6)
    )
    tf = tb.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Calibri"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = BLANCO

    # Número de página
    if page_num is not None:
        pn = slide.shapes.add_textbox(
            SLIDE_W - Inches(1.5), Inches(0.25), Inches(1.2), Inches(0.4)
        )
        ptf = pn.text_frame
        pp = ptf.paragraphs[0]
        pp.text = f"{page_num} / {total}"
        pp.alignment = PP_ALIGN.RIGHT
        pp.font.name = "Calibri"
        pp.font.size = Pt(12)
        pp.font.color.rgb = BLANCO


def add_footer(slide):
    """Pie de página con nombre del trabajo."""
    fb = slide.shapes.add_textbox(
        Inches(0.5), SLIDE_H - Inches(0.4),
        SLIDE_W - Inches(1), Inches(0.3)
    )
    tf = fb.text_frame
    p = tf.paragraphs[0]
    p.text = "Maximum Clique Problem — Optimización en Ingeniería — USACH"
    p.font.name = "Calibri"
    p.font.size = Pt(10)
    p.font.italic = True
    p.font.color.rgb = GRIS_TEXTO


def add_bullets(slide, items, left=Inches(0.7), top=Inches(1.4),
                width=None, height=None, font_size=20, line_spacing=1.3):
    """Agrega lista con bullets al slide."""
    if width is None:
        width = SLIDE_W - Inches(1.4)
    if height is None:
        height = SLIDE_H - top - Inches(0.6)
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        # item puede ser str o (texto, nivel)
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
                 accent_color=AZUL_USACH):
    """Caja con borde lateral coloreado + título + contenido."""
    # Fondo gris claro
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = GRIS_CLARO
    bg.line.color.rgb = accent_color
    bg.line.width = Pt(1.5)

    # Borde lateral grueso
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, Inches(0.12), height
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent_color
    bar.line.fill.background()

    # Título
    tb = slide.shapes.add_textbox(
        left + Inches(0.3), top + Inches(0.15), width - Inches(0.4), Inches(0.45)
    )
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "Calibri"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = accent_color

    # Contenido
    cb = slide.shapes.add_textbox(
        left + Inches(0.3), top + Inches(0.75),
        width - Inches(0.5), height - Inches(0.9)
    )
    ctf = cb.text_frame
    ctf.word_wrap = True
    for i, line in enumerate(content_lines):
        if i == 0:
            p = ctf.paragraphs[0]
        else:
            p = ctf.add_paragraph()
        p.text = line
        p.font.name = "Calibri"
        p.font.size = Pt(14)
        p.font.color.rgb = GRIS_TEXTO
        p.space_after = Pt(4)


def make_table(slide, data, left, top, width, height, header_color=AZUL_USACH):
    """Agrega una tabla con la primera fila como header."""
    rows = len(data)
    cols = len(data[0])
    tbl = slide.shapes.add_table(rows, cols, left, top, width, height).table
    for ri, row in enumerate(data):
        for ci, cell_text in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = str(cell_text)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(14)
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


# ============================================================
# CONSTRUCCIÓN DE LA PRESENTACIÓN
# ============================================================
prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]
TOTAL_SLIDES = 17

# -------------------------------------------------------------
# SLIDE 1 — PORTADA
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)

# Fondo azul superior
top_bg = s.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(2.5)
)
top_bg.fill.solid()
top_bg.fill.fore_color.rgb = AZUL_USACH
top_bg.line.fill.background()

# Strip rojo
strip = s.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, 0, Inches(2.5), SLIDE_W, Inches(0.12)
)
strip.fill.solid()
strip.fill.fore_color.rgb = ROJO_USACH
strip.line.fill.background()

# Logo
try:
    s.shapes.add_picture(LOGO_PNG, Inches(0.5), Inches(0.4),
                         height=Inches(1.7))
except Exception:
    pass

# Universidad
add_text_box(
    s, "UNIVERSIDAD DE SANTIAGO DE CHILE",
    Inches(3), Inches(0.6), Inches(10), Inches(0.5),
    font_size=20, bold=True, color=BLANCO
)
add_text_box(
    s, "Facultad de Ingeniería — Departamento de Ingeniería Informática",
    Inches(3), Inches(1.1), Inches(10), Inches(0.4),
    font_size=14, color=BLANCO, italic=True
)
add_text_box(
    s, "Magíster en Ingeniería Informática",
    Inches(3), Inches(1.5), Inches(10), Inches(0.4),
    font_size=14, color=BLANCO
)

# Título central
add_text_box(
    s, "Maximum Clique Problem",
    Inches(0.5), Inches(3.0), SLIDE_W - Inches(1), Inches(0.9),
    font_size=44, bold=True, color=AZUL_USACH, align=PP_ALIGN.CENTER
)
add_text_box(
    s, "Resolución exacta mediante Programación Lineal Entera",
    Inches(0.5), Inches(3.9), SLIDE_W - Inches(1), Inches(0.5),
    font_size=22, color=GRIS_TEXTO, italic=True, align=PP_ALIGN.CENTER
)
add_text_box(
    s, "Optimización en Ingeniería  —  Informe 1",
    Inches(0.5), Inches(4.5), SLIDE_W - Inches(1), Inches(0.4),
    font_size=16, color=ROJO_USACH, bold=True, align=PP_ALIGN.CENTER
)

# Autores
add_text_box(
    s, "Ricardo Riveros  •  Gonzalo Ahumada  •  Daniel Muñoz",
    Inches(0.5), Inches(5.7), SLIDE_W - Inches(1), Inches(0.4),
    font_size=18, color=GRIS_TEXTO, align=PP_ALIGN.CENTER
)
add_text_box(
    s, "Profesora: Mónica Villanueva  —  Ayudante: Sebastián Aliaga",
    Inches(0.5), Inches(6.1), SLIDE_W - Inches(1), Inches(0.4),
    font_size=14, color=GRIS_TEXTO, italic=True, align=PP_ALIGN.CENTER
)
add_text_box(
    s, "Santiago, mayo 2026",
    Inches(0.5), Inches(6.6), SLIDE_W - Inches(1), Inches(0.4),
    font_size=14, color=GRIS_TEXTO, align=PP_ALIGN.CENTER
)

# -------------------------------------------------------------
# SLIDE 2 — AGENDA
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Agenda", 2, TOTAL_SLIDES)
agenda = [
    "Definición del problema y complejidad",
    "Contextos de aplicación",
    "Objetivos y pregunta de investigación",
    "Estado del arte",
    "Formulación matemática (ILP y Motzkin–Straus)",
    "Implementación: Pyomo + HiGHS",
    "Resultados experimentales sobre keller4 (DIMACS)",
    "Comparativa de solvers HiGHS vs GLPK",
    "Análisis de la relajación LP",
    "Discusión, conclusiones y trabajo futuro",
]
add_bullets(s, agenda, top=Inches(1.5), font_size=20)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 3 — DEFINICIÓN DEL PROBLEMA
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "¿Qué es el Maximum Clique Problem?", 3, TOTAL_SLIDES)

add_text_box(
    s,
    "Dado un grafo no dirigido G = (V, E), un clique es un subconjunto C ⊆ V "
    "tal que todos sus vértices están conectados entre sí.",
    Inches(0.7), Inches(1.4), SLIDE_W - Inches(1.4), Inches(1),
    font_size=18, color=GRIS_TEXTO
)
add_text_box(
    s,
    "El MCP busca encontrar el clique de mayor cardinalidad:  ω(G) = |C*|",
    Inches(0.7), Inches(2.4), SLIDE_W - Inches(1.4), Inches(0.6),
    font_size=20, bold=True, color=AZUL_USACH
)

# Caja con ejemplo
add_box_card(
    s, "Ejemplo intuitivo",
    [
        "Red social: cada persona es un nodo, cada amistad una arista.",
        "Un clique = grupo donde TODOS se conocen entre sí.",
        "El MCP busca el grupo más numeroso con esa propiedad.",
    ],
    Inches(0.7), Inches(3.4), Inches(6), Inches(2.5),
    accent_color=AZUL_USACH,
)
add_box_card(
    s, "Equivalencias clásicas",
    [
        "• Maximum Independent Set",
        "• Minimum Vertex Cover",
        "(equivalentes vía complemento del grafo)",
    ],
    Inches(7), Inches(3.4), Inches(5.7), Inches(2.5),
    accent_color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 4 — COMPLEJIDAD
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Complejidad computacional", 4, TOTAL_SLIDES)
add_bullets(s, [
    "MCP pertenece a la clase NP-dura (Bomze et al., 1999; Pardalos & Xue, 1994).",
    "El tiempo computacional crece exponencialmente con |V|.",
    "No se conoce un algoritmo polinomial para instancias arbitrarias.",
    "Más aún: no es aproximable en tiempo polinomial dentro de un factor constante, a menos que P = NP.",
    "Sin embargo, es resoluble en tiempo polinomial en clases especiales: grafos perfectos, bipartitos, de intervalos, triangulados.",
], top=Inches(1.5), font_size=20)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 5 — CONTEXTOS DE APLICACIÓN
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Contextos de aplicación", 5, TOTAL_SLIDES)
add_box_card(
    s, "Teoría de códigos correctores de errores",
    [
        "Grafo G_{n,d}: vértices = palabras binarias de longitud n,",
        "aristas = pares con distancia Hamming ≥ d.",
        "",
        "Clique máximo en G_{n,d} = código óptimo de tamaño A(n,d).",
        "",
        "Familia Keller (usada en este trabajo): teselaciones",
        "del espacio R^n por hipercubos unitarios trasladados.",
    ],
    Inches(0.5), Inches(1.4), Inches(6.2), Inches(5.4),
    accent_color=AZUL_USACH,
)
add_box_card(
    s, "Bioinformática — Alineamiento molecular",
    [
        "Subestructura común máxima (MCS) entre dos moléculas.",
        "",
        "Grafo de productos G_P: vértices = pares (a,b)",
        "atómicamente compatibles; aristas = distancias internas",
        "consistentes.",
        "",
        "Aplicación: docking molecular, descubrimiento de fármacos,",
        "identificación de farmacóforos comunes.",
    ],
    Inches(7), Inches(1.4), Inches(5.8), Inches(5.4),
    accent_color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 6 — OBJETIVOS
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Objetivos y pregunta de investigación", 6, TOTAL_SLIDES)

add_box_card(
    s, "Objetivo general",
    [
        "Formular el MCP como un Programa Lineal Entero (ILP) y resolverlo",
        "de manera exacta sobre una instancia representativa del benchmark DIMACS,",
        "evaluando la calidad de la formulación mediante su relajación lineal continua.",
    ],
    Inches(0.5), Inches(1.4), SLIDE_W - Inches(1), Inches(1.7),
    accent_color=AZUL_USACH,
)

add_text_box(
    s, "Objetivos específicos",
    Inches(0.7), Inches(3.3), SLIDE_W - Inches(1.4), Inches(0.4),
    font_size=18, bold=True, color=AZUL_USACH,
)
add_bullets(s, [
    "Implementar el modelo ILP en Python con Pyomo.",
    "Resolver la instancia keller4 con el solver HiGHS (Branch & Bound + cortes + heurísticas).",
    "Calcular y analizar la brecha de integralidad entre la solución entera y la relajación continua.",
], top=Inches(3.8), font_size=16)

add_box_card(
    s, "Pregunta de investigación",
    [
        "¿Es suficiente la formulación ILP estándar del MCP, resuelta con un solver MIP moderno,",
        "para alcanzar la solución óptima conocida de una instancia densa DIMACS en tiempo razonable,",
        "y qué calidad ofrece su relajación lineal continua como cota superior?",
    ],
    Inches(0.5), Inches(5.5), SLIDE_W - Inches(1), Inches(1.5),
    accent_color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 7 — ESTADO DEL ARTE
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Estado del arte — Métodos de solución", 7, TOTAL_SLIDES)

add_box_card(
    s, "Métodos exactos",
    [
        "Branch & Bound — base algorítmica desde Carraghan & Pardalos (1990).",
        "",
        "Solvers especializados:",
        "  • cliquer (Östergård)",
        "  • MCR (Tomita)",
        "  • BBMC (San Segundo)",
        "",
        "Reformulación reciente:",
        "decisión parametrizada por k (Szabó & Zaválnij, 2018).",
    ],
    Inches(0.5), Inches(1.4), Inches(6.2), Inches(5.6),
    accent_color=AZUL_USACH,
)
add_box_card(
    s, "Métodos heurísticos",
    [
        "Greedy secuencial: Best in / Worst out",
        "  (Pardalos & Xue, 1994).",
        "",
        "Mejoras:",
        "  • Búsqueda local (k-intercambios)",
        "  • Randomized search",
        "  • Búsqueda Tabú",
        "  • Redes neuronales",
        "",
        "Conjuntos Independientes Mínimos para poda",
        "  (Singh & Govinda, 2014).",
    ],
    Inches(7), Inches(1.4), Inches(5.8), Inches(5.6),
    accent_color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 8 — MOTZKIN-STRAUS
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Variante teórica: Teorema de Motzkin–Straus", 8, TOTAL_SLIDES)

add_text_box(
    s,
    "El MCP también admite una formulación como problema de optimización cuadrática continua "
    "sobre el simplex Δ:",
    Inches(0.7), Inches(1.4), SLIDE_W - Inches(1.4), Inches(0.9),
    font_size=18, color=GRIS_TEXTO,
)

# Fórmula en caja destacada
formula_bg = s.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(2.5), Inches(2.7), Inches(8.3), Inches(1.5)
)
formula_bg.fill.solid()
formula_bg.fill.fore_color.rgb = GRIS_CLARO
formula_bg.line.color.rgb = AZUL_USACH
formula_bg.line.width = Pt(2)

add_text_box(
    s, "½ ( 1 − 1/ω(G) )  =  máx     ½ xᵀ A x",
    Inches(2.5), Inches(2.9), Inches(8.3), Inches(0.6),
    font_size=26, bold=True, color=AZUL_USACH, align=PP_ALIGN.CENTER,
)
add_text_box(
    s, "donde A es la matriz de adyacencia y x ∈ Δ = {x ∈ Rⁿ : Σxᵢ = 1, xᵢ ≥ 0}",
    Inches(2.5), Inches(3.5), Inches(8.3), Inches(0.6),
    font_size=14, italic=True, color=GRIS_TEXTO, align=PP_ALIGN.CENTER,
)

add_bullets(s, [
    "Establece equivalencia exacta entre un problema combinatorio discreto y uno continuo no lineal.",
    "Permite aplicar técnicas de programación cuadrática (QP).",
    "En este trabajo se documenta por completitud — no se implementa.",
    "Motivó extensiones modernas: relajaciones semidefinidas, variantes regularizadas.",
], top=Inches(4.5), font_size=16)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 9 — FORMULACIÓN ILP
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Formulación del MCP como ILP", 9, TOTAL_SLIDES)

add_text_box(
    s, "Variable de decisión binaria",
    Inches(0.7), Inches(1.4), Inches(6), Inches(0.5),
    font_size=18, bold=True, color=AZUL_USACH,
)
add_text_box(
    s, "xᵢ = 1  ⇔  vértice i ∈ clique           xᵢ = 0  en caso contrario",
    Inches(0.7), Inches(1.85), Inches(11), Inches(0.5),
    font_size=16, color=GRIS_TEXTO,
)

# Función objetivo
fo_bg = s.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(0.7), Inches(2.6), Inches(12), Inches(1.0)
)
fo_bg.fill.solid()
fo_bg.fill.fore_color.rgb = GRIS_CLARO
fo_bg.line.color.rgb = AZUL_USACH
fo_bg.line.width = Pt(1.5)
add_text_box(
    s, "Función objetivo:   máx   z = Σᵢ∈V  xᵢ",
    Inches(0.7), Inches(2.75), Inches(12), Inches(0.7),
    font_size=22, bold=True, color=AZUL_USACH, align=PP_ALIGN.CENTER,
)

# Restricciones
add_text_box(
    s, "Restricciones de adyacencia:",
    Inches(0.7), Inches(3.9), Inches(6), Inches(0.4),
    font_size=16, bold=True, color=GRIS_TEXTO,
)
add_text_box(
    s, "xᵢ + xⱼ ≤ 1     ∀ (i,j) ∈ Ē     con Ē = pares NO adyacentes (i<j)",
    Inches(0.7), Inches(4.3), Inches(12), Inches(0.5),
    font_size=18, color=GRIS_TEXTO,
)

add_text_box(
    s, "Restricciones de integralidad:",
    Inches(0.7), Inches(5.0), Inches(6), Inches(0.4),
    font_size=16, bold=True, color=GRIS_TEXTO,
)
add_text_box(
    s, "xᵢ ∈ {0, 1}     ∀ i ∈ V",
    Inches(0.7), Inches(5.4), Inches(12), Inches(0.5),
    font_size=18, color=GRIS_TEXTO,
)

add_text_box(
    s,
    "Si los vértices i, j no son adyacentes, no pueden estar ambos en el clique.",
    Inches(0.7), Inches(6.3), Inches(12), Inches(0.5),
    font_size=14, italic=True, color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 10 — IMPLEMENTACIÓN
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Implementación: Pyomo + HiGHS", 10, TOTAL_SLIDES)

add_box_card(
    s, "Stack tecnológico",
    [
        "• Python 3.12 — lenguaje de implementación",
        "• Pyomo 6.7 — modelado algebraico (declarativo)",
        "• HiGHS 1.7 — solver MIP de alto rendimiento",
        "• NetworkX 3.2 — manejo del grafo",
        "• Jupyter Notebook — reproducibilidad",
    ],
    Inches(0.5), Inches(1.4), Inches(6.2), Inches(3.0),
    accent_color=AZUL_USACH,
)
add_box_card(
    s, "Ventajas de Pyomo",
    [
        "• Separación modelo–solver (cambiar a CPLEX/Gurobi = 1 línea)",
        "• Legibilidad: refleja la formulación matemática",
        "• Extensibilidad: agregar cortes o simetrías sin reescribir",
    ],
    Inches(7), Inches(1.4), Inches(5.8), Inches(3.0),
    accent_color=ROJO_USACH,
)
add_box_card(
    s, "Por qué HiGHS",
    [
        "• Solver open-source de alto rendimiento (Universidad de Edimburgo).",
        "• Aplica internamente Branch & Bound + planos de corte (Gomory, MIR)",
        "  + heurísticas de factibilidad + estrategias de ramificación adaptativas.",
        "• Significativamente más rápido que GLPK en MIP densos (lo veremos).",
    ],
    Inches(0.5), Inches(4.6), SLIDE_W - Inches(1), Inches(2.4),
    accent_color=AZUL_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 11 — BRANCH & BOUND
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Algoritmo Branch & Bound (interno de HiGHS)", 11, TOTAL_SLIDES)

add_bullets(s, [
    "Inicialización: lista de subproblemas L = {P₀};  incumbente ẑ = 0",
    "Mientras L ≠ ∅:",
    ("Tomar subproblema P y resolver su relajación LP → obtener (x', z')", 1),
    ("Podar por infactibilidad si P no tiene solución", 1),
    ("Podar por cota si z' ≤ ẑ", 1),
    ("Si x' es entera, actualizar incumbente: ẑ ← z',  x* ← x'", 1),
    ("Si no, ramificar sobre variable fraccionaria xₖ → {xₖ=0, xₖ=1}", 1),
    "Retornar (ẑ, x*)",
], top=Inches(1.4), font_size=18, line_spacing=1.2)

# Caja final con cota
add_box_card(
    s, "Propiedad clave",
    [
        "z*_ILP ≤ z*_LP    (la relajación da cota superior siempre)",
        "Si la cota no mejora la incumbente → poda segura.",
    ],
    Inches(2), Inches(5.7), Inches(9.3), Inches(1.3),
    accent_color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 12 — INSTANCIA KELLER4
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Instancia experimental: keller4 (DIMACS)", 12, TOTAL_SLIDES)

add_text_box(
    s,
    "Grafo de la familia Keller, basado en la conjetura sobre teselaciones de R⁴ "
    "por hipercubos unitarios trasladados. Benchmark estándar para algoritmos MCP exactos.",
    Inches(0.7), Inches(1.4), SLIDE_W - Inches(1.4), Inches(1),
    font_size=16, italic=True, color=GRIS_TEXTO,
)

# Tabla con características
data = [
    ["Propiedad", "Valor"],
    ["Vértices  (|V|)", "171"],
    ["Aristas  (|E|)", "9 435"],
    ["Densidad", "0.6491"],
    ["Clique máximo conocido (ω)", "11"],
    ["Aristas del complemento (|Ē|)", "5 100"],
    ["Variables binarias del ILP", "171"],
    ["Restricciones del ILP", "5 100"],
    ["Razón restricciones / variable", "29.8 → modelo fuertemente restringido"],
]
make_table(
    s, data,
    Inches(2.5), Inches(2.7), Inches(8.3), Inches(4.0),
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 13 — RESULTADOS: HIGHS vs GLPK
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Resultados: comparativa HiGHS vs GLPK", 13, TOTAL_SLIDES)

add_text_box(
    s,
    "Mismo modelo ILP, ambos solvers ejecutados con Pyomo. Límite de tiempo GLPK = 300 s.",
    Inches(0.7), Inches(1.4), SLIDE_W - Inches(1.4), Inches(0.5),
    font_size=16, italic=True, color=GRIS_TEXTO,
)

data = [
    ["Métrica", "HiGHS", "GLPK"],
    ["Mejor solución (ω)", "11", "11"],
    ["Óptimo conocido", "11", "11"],
    ["Tiempo", "12–18 s", "> 300 s (timeout)"],
    ["Estado", "Óptimo", "Factible (sin probar optimalidad)"],
]
make_table(
    s, data,
    Inches(1.5), Inches(2.2), Inches(10.3), Inches(2.5),
)

add_box_card(
    s, "Hallazgo clave",
    [
        "Ambos solvers encuentran el clique óptimo ω=11.",
        "Solo HiGHS prueba optimalidad en tiempo razonable.",
        "Diferencia explicada por: preprocesamiento, cortes (Gomory/MIR),",
        "heurísticas de factibilidad y ramificación adaptativa en HiGHS.",
    ],
    Inches(1.5), Inches(5.0), Inches(10.3), Inches(2.0),
    accent_color=ROJO_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 14 — RELAJACIÓN LP
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Análisis de la relajación LP", 14, TOTAL_SLIDES)

add_text_box(
    s,
    "Reemplazamos xᵢ ∈ {0,1} por xᵢ ∈ [0,1] y resolvemos el LP continuo.",
    Inches(0.7), Inches(1.4), SLIDE_W - Inches(1.4), Inches(0.6),
    font_size=16, italic=True, color=GRIS_TEXTO,
)

data = [
    ["Métrica", "Valor"],
    ["Cota LP (relajación)", "85.50"],
    ["Solución entera (ILP)", "11"],
    ["Brecha de integralidad", "87.1 %"],
]
make_table(
    s, data,
    Inches(2.5), Inches(2.3), Inches(8.3), Inches(2.0),
)

add_box_card(
    s, "Implicación",
    [
        "La cota LP es muy floja (87.1 % de brecha): no captura la combinatoria del problema.",
        "→ El solver debe explorar un árbol B&B extenso para cerrar la brecha.",
        "",
        "Esta es una debilidad conocida de la formulación estándar del MCP.",
        "Se aborda con cortes (HiGHS los aplica automáticamente).",
    ],
    Inches(1.5), Inches(4.7), Inches(10.3), Inches(2.3),
    accent_color=AZUL_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 15 — DISCUSIÓN / DIFICULTADES
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Discusión y dificultades experimentales", 15, TOTAL_SLIDES)

add_box_card(
    s, "Dificultad 1 — Debilidad de la relajación LP",
    [
        "Brecha del 87.1 % → árbol B&B muy extenso.",
        "Solución: delegar a HiGHS (cortes Gomory/MIR + heurísticas).",
    ],
    Inches(0.5), Inches(1.4), SLIDE_W - Inches(1), Inches(1.6),
    accent_color=AZUL_USACH,
)
add_box_card(
    s, "Dificultad 2 — Falta de convergencia de GLPK",
    [
        "GLPK timeout en 300 s sin probar optimalidad.",
        "Solución: HiGHS como solver principal; GLPK como punto de comparación.",
    ],
    Inches(0.5), Inches(3.2), SLIDE_W - Inches(1), Inches(1.6),
    accent_color=ROJO_USACH,
)
add_box_card(
    s, "Dificultad 3 — Validación post-hoc del clique",
    [
        "Verificación: el subgrafo inducido por los 11 vértices",
        "debe contener exactamente C(11,2) = 55 aristas, todas en G.",
    ],
    Inches(0.5), Inches(5.0), SLIDE_W - Inches(1), Inches(1.6),
    accent_color=AZUL_USACH,
)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 16 — CONCLUSIONES Y TRABAJO FUTURO
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)
add_header_bar(s, "Conclusiones y trabajo futuro", 16, TOTAL_SLIDES)

add_text_box(
    s, "Conclusiones principales",
    Inches(0.7), Inches(1.4), Inches(6), Inches(0.5),
    font_size=20, bold=True, color=AZUL_USACH,
)
add_bullets(s, [
    "Se obtuvo el óptimo conocido ω = 11 sobre keller4 en 12–18 s.",
    "La elección del solver es determinante: HiGHS >> GLPK en este tipo de modelos.",
    "La formulación ILP estándar es suficiente para instancias moderadas, pero su relajación LP es débil.",
    "Pyomo permite traducir la formulación matemática a código sin perder claridad.",
], top=Inches(1.9), font_size=16, line_spacing=1.2)

add_text_box(
    s, "Trabajo futuro",
    Inches(0.7), Inches(4.5), Inches(6), Inches(0.5),
    font_size=20, bold=True, color=ROJO_USACH,
)
add_bullets(s, [
    "Implementar la metaheurística (Informe 2): algoritmo genético, tabú o GRASP.",
    "Contrastar enfoque exacto vs. metaheurístico sobre múltiples familias DIMACS.",
    "Explorar formulaciones reforzadas (desigualdades de clique, cortes odd-hole).",
    "Evaluar la formulación cuadrática de Motzkin–Straus.",
], top=Inches(5.0), font_size=16, line_spacing=1.2)
add_footer(s)

# -------------------------------------------------------------
# SLIDE 17 — CIERRE
# -------------------------------------------------------------
s = prs.slides.add_slide(BLANK)

# Fondo azul completo
bg = s.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H
)
bg.fill.solid()
bg.fill.fore_color.rgb = AZUL_USACH
bg.line.fill.background()

# Strip rojo central
strip = s.shapes.add_shape(
    MSO_SHAPE.RECTANGLE, 0, Inches(4.0), SLIDE_W, Inches(0.15)
)
strip.fill.solid()
strip.fill.fore_color.rgb = ROJO_USACH
strip.line.fill.background()

add_text_box(
    s, "¿Preguntas?",
    Inches(0.5), Inches(2.5), SLIDE_W - Inches(1), Inches(1.2),
    font_size=72, bold=True, color=BLANCO, align=PP_ALIGN.CENTER,
)
add_text_box(
    s, "Gracias por su atención",
    Inches(0.5), Inches(4.3), SLIDE_W - Inches(1), Inches(0.6),
    font_size=28, italic=True, color=BLANCO, align=PP_ALIGN.CENTER,
)
add_text_box(
    s,
    "Ricardo Riveros  •  Gonzalo Ahumada  •  Daniel Muñoz",
    Inches(0.5), Inches(5.3), SLIDE_W - Inches(1), Inches(0.5),
    font_size=18, color=BLANCO, align=PP_ALIGN.CENTER,
)
add_text_box(
    s, "Optimización en Ingeniería — USACH 2026",
    Inches(0.5), Inches(5.8), SLIDE_W - Inches(1), Inches(0.5),
    font_size=14, italic=True, color=BLANCO, align=PP_ALIGN.CENTER,
)

# ============================================================
# GUARDAR
# ============================================================
prs.save(str(OUTPUT))
print(f"OK: presentación guardada en {OUTPUT}")
print(f"    {len(prs.slides)} slides")
