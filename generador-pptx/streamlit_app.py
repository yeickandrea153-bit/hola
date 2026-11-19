"""Generador automático de presentaciones PowerPoint (.pptx) con Streamlit + python-pptx.

Pega un texto, elige una paleta y descarga la presentación: portada, índice,
diapositivas de desarrollo, conclusiones y fuentes. Todo se genera en memoria (io.BytesIO).
"""
import io
import math
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date

import streamlit as st
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ----------------------------------------------------------------------------- paletas
PALETAS = {
    "Azul Corporativo": dict(fondo="FFFFFF", portada="1F3A5F", titulo="1F3A5F", texto="2B2B2B",
                             acento="2E75B6", sobre_portada="FFFFFF", tenue="7F8C9A"),
    "Oscuro Moderno": dict(fondo="1E1E2E", portada="11111B", titulo="FFFFFF", texto="D9D9E3",
                           acento="00C2A8", sobre_portada="FFFFFF", tenue="8A8AA3"),
    "Verde Minimalista": dict(fondo="F7FAF7", portada="2F5D50", titulo="2F5D50", texto="333333",
                              acento="6BAF92", sobre_portada="FFFFFF", tenue="7C8C84"),
    "Naranja Cálido": dict(fondo="FFFaf5", portada="7A2E0E", titulo="7A2E0E", texto="3A2A22",
                           acento="E8710A", sobre_portada="FFFFFF", tenue="9A8478"),
}

ANCHO, ALTO = Inches(13.333), Inches(7.5)
CHARS_POR_LINEA = 62      # estimación de caracteres por línea con 24 pt en el cuadro de texto
MAX_LINEAS = 9            # líneas de texto máximas por diapositiva (evita desbordes)
PALABRAS_CONCLUSION = ("conclusi", "recomend")

# ----------------------------------------------------------------------------- análisis del texto
@dataclass
class Seccion:
    titulo: str
    puntos: list = field(default_factory=list)


RE_MD = re.compile(r"^\s{0,3}#{1,6}\s+(.*\S)\s*$")
RE_VINETA = re.compile(r"^\s*(?:[-*•–·]|\d{1,2}[.)])\s+(.+\S)\s*$")
RE_ORACION = re.compile(r"(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÜÑ¿¡0-9\"“])")


def _limpiar_titulo(t):
    t = re.sub(r"^\s*(?:#{1,6}|\d{1,2}[.)])\s*", "", t).strip()
    return t.rstrip(":").strip(" *_")


def _es_titulo(linea, previa_vacia, siguiente):
    s = linea.strip()
    if RE_MD.match(linea):
        return True
    if len(s) > 80 or not s:
        return False
    if s.endswith(":") and s.count(":") == 1:
        return True
    letras = [c for c in s if c.isalpha()]
    if len(letras) >= 4 and s.isupper():
        return True
    # línea corta, sola, sin puntuación final y seguida de contenido
    if previa_vacia and siguiente and len(s) <= 60 and not s.endswith((".", ",", ";", "!", "?")):
        return not RE_VINETA.match(linea) or bool(re.match(r"^\s*\d{1,2}[.)]\s+", linea))
    return False


def _oraciones(texto):
    return [o.strip() for o in RE_ORACION.split(texto.strip()) if o.strip()]


def _puntos_de_linea(linea):
    m = RE_VINETA.match(linea)
    if m:
        return [m.group(1).strip()]
    return _oraciones(linea)


def _resumen_titulo(oracion, max_palabras=8):
    o = oracion.strip().rstrip(".:;")
    palabras = o.split()
    if len(palabras) <= max_palabras:
        return o
    return " ".join(palabras[:max_palabras]).rstrip(",;:") + "…"


def analizar_texto(texto):
    """Convierte el texto pegado en una lista de Seccion (título + puntos)."""
    lineas = texto.replace("\r", "").split("\n")
    n = len(lineas)
    secciones, actual = [], None

    def siguiente_con_contenido(i):
        return i + 1 < n and lineas[i + 1].strip() != ""

    hay_titulos = False
    for i, l in enumerate(lineas):
        previa_vacia = i == 0 or lineas[i - 1].strip() == ""
        if l.strip() and _es_titulo(l, previa_vacia, siguiente_con_contenido(i)):
            hay_titulos = True
            break

    if hay_titulos:
        for i, l in enumerate(lineas):
            if not l.strip():
                continue
            previa_vacia = i == 0 or lineas[i - 1].strip() == ""
            if _es_titulo(l, previa_vacia, siguiente_con_contenido(i)):
                actual = Seccion(_limpiar_titulo(l))
                secciones.append(actual)
            else:
                if actual is None:
                    actual = Seccion("Introducción")
                    secciones.append(actual)
                actual.puntos.extend(_puntos_de_linea(l))
    else:  # sin títulos: cada párrafo (separado por línea en blanco) es una diapositiva
        bloques, buf = [], []
        for l in lineas + [""]:
            if l.strip():
                buf.append(l)
            elif buf:
                bloques.append(buf)
                buf = []
        for k, b in enumerate(bloques, 1):
            puntos = [p for l in b for p in _puntos_de_linea(l)]
            if len(puntos) >= 2 and len(puntos[0]) <= 70:
                secciones.append(Seccion(_resumen_titulo(puntos[0]), puntos[1:]))
            elif len(puntos) >= 2:
                secciones.append(Seccion(_resumen_titulo(puntos[0]), puntos))
            elif puntos:
                secciones.append(Seccion(f"Idea {k}", puntos))
    return [s for s in secciones if s.puntos]


def separar_conclusiones(secciones):
    desarrollo, conclusion = [], []
    for s in secciones:
        (conclusion if any(p in s.titulo.lower() for p in PALABRAS_CONCLUSION) else desarrollo).append(s)
    return desarrollo, conclusion


def _lineas(texto):
    return max(1, math.ceil(len(texto) / CHARS_POR_LINEA))


def _partir_largo(texto, max_chars=CHARS_POR_LINEA * 4):
    """Divide un punto muy largo en trozos por palabras para que quepa en la diapositiva."""
    if len(texto) <= max_chars:
        return [texto]
    trozos, actual = [], ""
    for palabra in texto.split():
        if len(actual) + len(palabra) + 1 > max_chars and actual:
            trozos.append(actual + "…")
            actual = "…" + palabra
        else:
            actual = f"{actual} {palabra}".strip()
    trozos.append(actual)
    return trozos


def dividir_puntos(puntos, max_puntos=5):
    """Agrupa puntos en diapositivas sin superar max_puntos ni MAX_LINEAS."""
    planos = [t for p in puntos for t in _partir_largo(p)]
    grupos, actual, lineas = [], [], 0
    for p in planos:
        l = _lineas(p)
        if actual and (len(actual) >= max_puntos or lineas + l > MAX_LINEAS):
            grupos.append(actual)
            actual, lineas = [], 0
        actual.append(p)
        lineas += l
    if actual:
        grupos.append(actual)
    return grupos


# ----------------------------------------------------------------------------- utilidades pptx
def _rgb(hex6):
    return RGBColor.from_string(hex6.upper())


def _fondo(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(color)


def _rect(slide, x, y, w, h, color):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    r.fill.solid()
    r.fill.fore_color.rgb = _rgb(color)
    r.line.fill.background()
    r.shadow.inherit = False
    return r


def _estilo_run(run, size, color, bold=False, nombre="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = nombre
    run.font.color.rgb = _rgb(color)


def _vineta(parrafo, color):
    """Viñeta real de PowerPoint con sangría francesa (se agrega después del espaciado)."""
    pPr = parrafo._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(0.38))))
    pPr.set("indent", str(-int(Inches(0.38))))
    clr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(clr, qn("a:srgbClr")).set("val", color.upper())
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def _cuadro(slide, x, y, w, h):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = Inches(0.05)
    return tb, tf


def _titulo(slide, texto, pal):
    texto = texto if len(texto) <= 90 else texto[:87] + "…"
    t = slide.shapes.title
    t.left, t.top, t.width, t.height = Inches(0.9), Inches(0.4), Inches(11.9), Inches(1.15)
    tf = t.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.text = texto
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _estilo_run(p.runs[0], 36 if len(texto) <= 42 else 28, pal["titulo"], bold=True)
    _rect(slide, Inches(0.5), Inches(0.55), Inches(0.12), Inches(0.85), pal["acento"])
    _rect(slide, Inches(0.9), Inches(1.65), Inches(11.9), Inches(0.03), pal["acento"])


def _pie(slide, pal, autor, numero):
    _, tf = _cuadro(slide, Inches(0.9), Inches(6.95), Inches(8), Inches(0.35))
    r = tf.paragraphs[0].add_run()
    r.text = autor
    _estilo_run(r, 11, pal["tenue"])
    _, tf = _cuadro(slide, Inches(11.8), Inches(6.95), Inches(1.0), Inches(0.35))
    tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
    r = tf.paragraphs[0].add_run()
    r.text = str(numero)
    _estilo_run(r, 11, pal["tenue"])


def _diapositiva_base(prs, pal, titulo, autor, numero):
    s = prs.slides.add_slide(prs.slide_layouts[5])  # "Solo el título"
    _fondo(s, pal["fondo"])
    _titulo(s, titulo, pal)
    _pie(s, pal, autor, numero)
    return s


def _lista_puntos(slide, puntos, pal, tam=24):
    _, tf = _cuadro(slide, Inches(0.9), Inches(1.95), Inches(11.6), Inches(4.85))
    for i, texto in enumerate(puntos):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(14)
        p.line_spacing = 1.05
        r = p.add_run()
        r.text = texto
        _estilo_run(r, tam, pal["texto"])
        _vineta(p, pal["acento"])


def _lista_numerada(slide, items, pal):
    _, tf = _cuadro(slide, Inches(0.9), Inches(1.95), Inches(11.6), Inches(4.85))
    for i, (num, texto) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(14)
        a = p.add_run()
        a.text = f"{num:02d}   "
        _estilo_run(a, 26, pal["acento"], bold=True)
        b = p.add_run()
        b.text = texto if len(texto) <= 85 else texto[:82] + "…"
        _estilo_run(b, 26, pal["texto"])


# ----------------------------------------------------------------------------- generación
def generar_pptx(titulo, autor, fuente, paleta, contenido, max_puntos=5, incluir_indice=True):
    """Devuelve (bytes_del_pptx, esquema) sin escribir nada en disco."""
    pal = PALETAS[paleta]
    secciones = analizar_texto(contenido)
    if not secciones:
        raise ValueError("No se encontró contenido para crear diapositivas.")
    desarrollo, conclusion = separar_conclusiones(secciones)
    if not desarrollo:  # todo el texto eran conclusiones
        desarrollo, conclusion = conclusion, []

    prs = Presentation()
    prs.slide_width, prs.slide_height = ANCHO, ALTO
    prs.core_properties.title = titulo
    prs.core_properties.author = autor
    esquema = []
    n = 1

    # --- portada
    s = prs.slides.add_slide(prs.slide_layouts[0])
    _fondo(s, pal["portada"])
    _rect(s, Inches(0), Inches(0), Inches(0.35), ALTO, pal["acento"])
    t = s.shapes.title
    t.left, t.top, t.width, t.height = Inches(1.0), Inches(1.9), Inches(11.3), Inches(2.4)
    tf = t.text_frame
    tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.BOTTOM
    tf.text = titulo
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    _estilo_run(tf.paragraphs[0].runs[0], 44 if len(titulo) <= 45 else 34, pal["sobre_portada"], bold=True)
    _rect(s, Inches(1.05), Inches(4.5), Inches(1.6), Inches(0.07), pal["acento"])
    sub = s.placeholders[1]
    sub.left, sub.top, sub.width, sub.height = Inches(1.0), Inches(4.8), Inches(11.3), Inches(1.4)
    tf = sub.text_frame
    tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.TOP
    tf.text = autor or ""
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    if tf.paragraphs[0].runs:
        _estilo_run(tf.paragraphs[0].runs[0], 24, pal["sobre_portada"])
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.LEFT
    r = p2.add_run()
    r.text = date.today().strftime("%d/%m/%Y")
    _estilo_run(r, 16, pal["acento"])
    esquema.append(("Portada", titulo))

    # --- índice
    if incluir_indice:
        titulos = [x.titulo for x in desarrollo]
        paginas = [titulos[i:i + 8] for i in range(0, len(titulos), 8)]
        for k, pag in enumerate(paginas):
            n += 1
            nombre = "Contenido" if len(paginas) == 1 else f"Contenido ({k + 1}/{len(paginas)})"
            sl = _diapositiva_base(prs, pal, nombre, autor, n)
            _lista_numerada(sl, [(k * 8 + j + 1, t) for j, t in enumerate(pag)], pal)
            esquema.append(("Índice", nombre))

    # --- desarrollo
    for sec in desarrollo:
        grupos = dividir_puntos(sec.puntos, max_puntos)
        for g, pts in enumerate(grupos, 1):
            n += 1
            nombre = sec.titulo if len(grupos) == 1 else f"{sec.titulo} ({g}/{len(grupos)})"
            sl = _diapositiva_base(prs, pal, nombre, autor, n)
            _lista_puntos(sl, pts, pal)
            esquema.append(("Desarrollo", nombre))

    # --- conclusiones / recomendaciones
    if conclusion:
        for sec in conclusion:
            for g, pts in enumerate(dividir_puntos(sec.puntos, max_puntos), 1):
                n += 1
                sl = _diapositiva_base(prs, pal, sec.titulo if g == 1 else f"{sec.titulo} ({g})", autor, n)
                _lista_puntos(sl, pts, pal)
                esquema.append(("Conclusiones", sec.titulo))
    else:
        ideas = [x.puntos[0] for x in desarrollo[:5]]
        n += 1
        sl = _diapositiva_base(prs, pal, "Conclusiones", autor, n)
        _lista_puntos(sl, [t for i in ideas for t in _partir_largo(i)][:5], pal)
        esquema.append(("Conclusiones", "Conclusiones (ideas clave generadas automáticamente)"))

    # --- fuentes
    n += 1
    sl = _diapositiva_base(prs, pal, "Fuentes y referencias", autor, n)
    refs = [l.strip(" -•\t") for l in re.split(r"[\n;]+", fuente or "") if l.strip(" -•\t")]
    _lista_puntos(sl, refs or ["Fuente no especificada."], pal, tam=20)
    esquema.append(("Fuentes", "Fuentes y referencias"))

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue(), esquema


# ----------------------------------------------------------------------------- interfaz Streamlit
EJEMPLO = """Introducción
El agua es un recurso esencial para la vida. Su calidad depende de factores físicos, químicos y biológicos.

Parámetros fisicoquímicos
- El pH indica la acidez o alcalinidad del agua.
- La conductividad eléctrica refleja la cantidad de sales disueltas.
- Los sólidos suspendidos totales afectan la turbidez.

Monitoreo y análisis
Se recolectan muestras en distintos puntos y se comparan estadísticamente. Los resultados se contrastan con los Estándares de Calidad Ambiental.

Conclusiones
- Los parámetros deben compararse con el ECA vigente.
- Se recomienda un monitoreo periódico en todos los puntos."""


def _slug(texto):
    t = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-zA-Z0-9]+", "_", t).strip("_").lower() or "presentacion"


def main():
    st.set_page_config(page_title="Generador de presentaciones", page_icon="📊", layout="centered")
    st.title("📊 Generador automático de presentaciones")
    st.caption("Pega tu contenido, elige una paleta y descarga un archivo .pptx listo para editar.")

    if "texto" not in st.session_state:
        st.session_state["texto"] = ""

    c1, c2 = st.columns(2)
    titulo = c1.text_input("Título de la presentación", placeholder="Ej.: Calidad del agua en la cuenca")
    autor = c2.text_input("Autor", placeholder="Ej.: Tu nombre")
    fuente = st.text_input("Fuente / Referencia", placeholder="Ej.: MINAM (2017). D.S. N° 004-2017-MINAM")
    c3, c4 = st.columns(2)
    paleta = c3.selectbox("Paleta de colores", list(PALETAS))
    max_puntos = c4.slider("Viñetas máximas por diapositiva", 3, 7, 5)
    incluir_indice = st.checkbox("Incluir diapositiva de índice", value=True)

    if st.button("Usar texto de ejemplo"):
        st.session_state["texto"] = EJEMPLO
    st.text_area(
        "Contenido o puntos clave",
        key="texto",
        height=320,
        help="Usa títulos (líneas cortas, terminadas en ':' o con #) para separar secciones, "
             "y guiones para las viñetas. Si no hay títulos, cada párrafo será una diapositiva.",
    )

    if st.button("Generar presentación", type="primary"):
        if not titulo.strip() or not st.session_state["texto"].strip():
            st.error("Escribe al menos el título y el contenido.")
        else:
            try:
                datos, esquema = generar_pptx(titulo.strip(), autor.strip(), fuente.strip(), paleta,
                                              st.session_state["texto"], max_puntos, incluir_indice)
                st.session_state["resultado"] = (datos, esquema, f"{_slug(titulo)}.pptx")
            except ValueError as e:
                st.session_state.pop("resultado", None)
                st.error(str(e))

    if "resultado" in st.session_state:
        datos, esquema, nombre = st.session_state["resultado"]
        st.success(f"Presentación lista: {len(esquema)} diapositivas.")
        st.download_button(
            "⬇️ Descargar .pptx", data=datos, file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        with st.expander("Ver esquema de diapositivas"):
            for i, (tipo, nom) in enumerate(esquema, 1):
                st.write(f"**{i}.** {nom}  ·  _{tipo}_")
        if not fuente.strip():
            st.warning("No indicaste fuente: la última diapositiva dirá «Fuente no especificada».")


if __name__ == "__main__":
    main()
