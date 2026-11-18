import io
import json
import os
import re
import time
import requests
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types

# ==============================================================================
# 1. CONFIGURACIÓN INTERNA Y ESTILO VISUAL
# ==============================================================================
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    api_key = os.environ.get("GEMINI_API_KEY", "")

st.set_page_config(
    page_title="PROFE IA | Suite CNB Guatemala",
    page_icon="🦜",
    layout="wide",
    initial_sidebar_state="expanded"
)

ICONO_QUETZAL_SVG = '<svg width="48" height="48" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="32" cy="32" r="30" fill="#ECFDF5" stroke="#059669" stroke-width="2.5"/><path d="M26 39C20 47 16 54 14 59C19 55 24 49 29 41" fill="#047857"/><path d="M31 40C27 49 24 56 23 61C27 56 31 50 34 41" fill="#10B981"/><path d="M23 33C22 24 27 16 35 16C41 16 44 21 43 28C42 35 36 41 28 41C25 41 23 37 23 33Z" fill="#059669"/><path d="M34 25C39 25 42 29 41 34C39 39 34 41 30 40C32 36 33 30 34 25Z" fill="#DC2626"/><path d="M30 16C31 11 35 9 38 11C36 12 35 14 36 16" stroke="#047857" stroke-width="2.5" stroke-linecap="round"/><path d="M33 15C35 10 39 9 41 12" stroke="#10B981" stroke-width="2" stroke-linecap="round"/><path d="M25 27C28 24 33 25 34 30C32 34 27 35 25 32Z" fill="#047857"/><path d="M42 19L48 21L42 23Z" fill="#F59E0B"/><circle cx="39" cy="19.5" r="1.6" fill="#FFFFFF"/><circle cx="39.3" cy="19.5" r="0.8" fill="#111827"/></svg>'

st.markdown("""
<style>
.hero-box {
    background: linear-gradient(120deg, #1e3a8a 0%, #2563eb 55%, #0ea5e9 100%);
    padding: 1.8rem 2rem; border-radius: 16px; color: white; margin-bottom: 1.5rem;
}
.hero-box h1 { color: white; margin: 0; font-size: 2rem; font-weight: 700; }
.hero-box p { color: #e0f2fe; margin: 0.4rem 0 0 0; font-size: 1.05rem; }
.tool-card {
    background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 4px solid #2563eb;
    border-radius: 12px; padding: 1rem; margin-bottom: 0.6rem; min-height: 135px;
}
.cnb-pill {
    background-color: #eff6ff; border-left: 4px solid #2563eb;
    padding: 0.9rem 1.1rem; border-radius: 8px; margin: 0.8rem 0; color: #1e3a8a;
}
.brand-header {
    display: flex; align-items: center; gap: 12px; padding: 0.2rem 0 0.4rem 0;
}
.brand-title {
    font-size: 1.65rem; font-weight: 800; color: #065f46; letter-spacing: 0.6px; line-height: 1.1;
}
.brand-sub {
    font-size: 0.82rem; color: #475569; margin-top: 2px;
}
.login-card {
    background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px;
    padding: 2.2rem; max-width: 480px; margin: 2rem auto; text-align: center;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.06);
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. CONTROL DE SESIÓN DEL DOCENTE
# ==============================================================================
if "docente_invitado" not in st.session_state:
    st.session_state.docente_invitado = ""

if not st.session_state.docente_invitado:
    st.markdown(f"""
    <div class="login-card">
        <div style="display:flex; justify-content:center; margin-bottom:12px;">{ICONO_QUETZAL_SVG}</div>
        <h2 style="color:#065f46; margin-bottom:4px;">Bienvenido a PROFE IA</h2>
        <p style="color:#475569; font-size:0.95rem; margin-bottom:1.5rem;">
            Tu asistente pedagógico alineado al CNB de Guatemala.<br>
            Escribe tu nombre para personalizar tus planificaciones y documentos.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_izq, col_centro, col_der = st.columns([1, 1.3, 1])
    with col_centro:
        with st.form("form_ingreso_rapido"):
            nombre_ingresado = st.text_input("Tu nombre y apellido:", placeholder="Ej. María Fernanda López")
            entrar = st.form_submit_button("Ingresar a la Plataforma →", type="primary", use_container_width=True)
            if entrar:
                if nombre_ingresado.strip():
                    st.session_state.docente_invitado = nombre_ingresado.strip()
                    st.rerun()
                else:
                    st.warning("Por favor escribe tu nombre para personalizar tus documentos.")
    st.stop()

nombre_usuario_actual = st.session_state.docente_invitado
primer_nombre = nombre_usuario_actual.split()[0]

# ==============================================================================
# 3. CATÁLOGO CURRICULAR (PREPRIMARIA A DIVERSIFICADO)
# ==============================================================================
CURRICULO_GT = {
    "Preprimaria (4, 5 y 6 años)": {
        "grados": ["Párvulos 1 (4 años)", "Párvulos 2 (5 años)", "Párvulos 3 (6 años / Preparatoria)"],
        "areas": ["Destrezas de Aprendizaje", "Comunicación y Lenguaje", "Medio Social y Natural", "Expresión Artística", "Educación Física"]
    },
    "Primaria (1ro a 6to Grado)": {
        "grados": ["1ro Primaria", "2do Primaria", "3ro Primaria", "4to Primaria", "5to Primaria", "6to Primaria"],
        "areas": ["Comunicación y Lenguaje L1", "Comunicación y Lenguaje L2", "Comunicación y Lenguaje L3 (Inglés)", "Matemáticas", "Medio Social y Natural (1ro-3ro)", "Ciencias Naturales y Tecnología (4to-6to)", "Ciencias Sociales (4to-6to)", "Expresión Artística", "Educación Física", "Formación Ciudadana"]
    },
    "Ciclo Básico (1ro a 3ro Básico)": {
        "grados": ["1ro Básico", "2do Básico", "3ro Básico"],
        "areas": ["Comunicación y Lenguaje Idioma Español", "Inglés", "Matemáticas", "Ciencias Naturales", "Ciencias Sociales y Formación Ciudadana", "Culturas e Idiomas Mayas, Garífuna o Xinka", "Educación Artística", "Emprendimiento para la Productividad", "TAC", "Educación Física"]
    },
    "Ciclo Diversificado (Bachillerato y Carreras)": {
        "grados": ["4to Diversificado", "5to Diversificado", "6to Diversificado"],
        "areas": ["Lengua y Literatura", "Matemática", "Física", "Química", "Biología", "Ciencias Sociales", "Filosofía", "Psicología", "Estadística", "Contabilidad / Administración", "TIC", "Seminario / Investigación"]
    }
}

# ==============================================================================
# 4. BUSCADOR CNB EN VIVO
# ==============================================================================
@st.cache_data(show_spinner=False)
def buscar_en_cnb_vivo(grado, area, tema):
    url = "https://cnbguatemala.org/api.php"
    headers = {"User-Agent": "ProfeIAGuatemala/2.0"}
    query = f"Malla curricular {area} {grado} {tema}".strip()
    try:
        r_search = requests.get(url, params={"action": "query", "list": "search", "srsearch": query, "format": "json", "srlimit": 3}, headers=headers, timeout=6)
        resultados = r_search.json().get("query", {}).get("search", [])
        if not resultados:
            return []
        titulo = resultados[0]["title"]
        r_parse = requests.get(url, params={"action": "parse", "page": titulo, "prop": "text", "format": "json"}, headers=headers, timeout=8)
        soup = BeautifulSoup(r_parse.json().get("parse", {}).get("text", {}).get("*", ""), "html.parser")
        registros = []
        for tabla in soup.find_all("table", class_="wikitable"):
            for fila in tabla.find_all("tr")[1:]:
                celdas = fila.find_all(["td", "th"])
                if len(celdas) >= 3:
                    comp = re.sub(r'\s+', ' ', celdas[0].get_text(strip=True))
                    ind = re.sub(r'\s+', ' ', celdas[1].get_text(strip=True))
                    cont = [re.sub(r'\s+', ' ', li.get_text(strip=True)) for li in celdas[2].find_all("li")] or [re.sub(r'\s+', ' ', celdas[2].get_text(strip=True))]
                    if comp and ind and len(ind) > 8:
                        registros.append({"fuente": f"CNB Oficial ({titulo})", "competencia": comp, "indicador_logro": ind, "contenidos": cont})
        return registros
    except Exception:
        return []

# ==============================================================================
# 5. MOTOR DE IA (TEXTO, JSON E ILUSTRACIONES VECTORIALES SVG)
# ==============================================================================
def obtener_modelos_activos(client):
    modelos = ["gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-3.5-flash", "gemini-3.8-pro"]
    try:
        for m in client.models.list():
            n = m.name.replace("models/", "")
            if "gemini" in n and ("flash" in n or "pro" in n) and n not in modelos:
                if not any(x in n for x in ["embedding", "vision", "image", "audio", "tts", "2.0", "2.5"]):
                    modelos.append(n)
    except Exception:
        pass
    return modelos

def llamar_gemini_json(key, prompt):
    if not key:
        raise RuntimeError("Falta configurar GEMINI_API_KEY en los Secrets de Streamlit Cloud.")
    client = genai.Client(api_key=key)
    ultimo_error = ""
    for mod in obtener_modelos_activos(client)[:5]:
        try:
            r = client.models.generate_content(model=mod, contents=prompt, config={"response_mime_type": "application/json", "temperature": 0.7})
            data = json.loads(r.text)
            return data[0] if isinstance(data, list) else data
        except Exception as e:
            ultimo_error = str(e)
            time.sleep(1)
    raise RuntimeError(f"Error de conexión IA: {ultimo_error}")

def llamar_gemini_texto(key, prompt):
    if not key:
        raise RuntimeError("Falta configurar GEMINI_API_KEY en los Secrets de Streamlit Cloud.")
    client = genai.Client(api_key=key)
    for mod in obtener_modelos_activos(client)[:5]:
        try:
            return client.models.generate_content(model=mod, contents=prompt).text
        except Exception:
            time.sleep(1)
    raise RuntimeError("No se pudo obtener respuesta del asistente en este momento.")

def generar_imagen_educativa(key, descripcion_imagen, estilo="Ilustración libro de texto"):
    if not key:
        return None, "Falta configurar GEMINI_API_KEY."

    estilos_svg = {
        "Ilustración libro de texto": "Usa colores vivos, degradados suaves, formas limpias y fondo claro (#F8FAFC). Estilo libro de texto escolar moderno.",
        "Dibujo para colorear (Blanco y negro)": "Usa EXCLUSIVAMENTE fondo blanco (#FFFFFF), líneas negras gruesas (stroke='#000000' stroke-width='3') y rellenos blancos (fill='#FFFFFF'). Sin colores ni grises, listo para que los niños coloreen.",
        "Esquema / Mapa visual de Ciencias": "Diseña un diagrama científico claro con flechas, iconos geométricos limpios y etiquetas de texto cortas y legibles en español.",
        "Contexto Cultural de Guatemala": "Incluye elementos visuales representativos de Guatemala (volcanes, quetzal, textiles, ceiba o paisajes locales) con colores vivos y estilo educativo."
    }
    instruccion_estilo = estilos_svg.get(estilo, estilos_svg["Ilustración libro de texto"])

    prompt_svg = f"""Eres un ilustrador educativo experto. Crea una ilustración vectorial en código SVG completa, detallada y estéticamente agradable sobre: "{descripcion_imagen}".
Estilo requerido: {instruccion_estilo}

Reglas técnicas obligatorias:
1. Devuelve ÚNICAMENTE el código XML del SVG empezando con <svg viewBox="0 0 800 600" width="800" height="600" xmlns="http://www.w3.org/2000/svg"> y terminando con </svg>.
2. NO incluyas bloques markdown (```svg o ```xml), ni explicaciones antes o después.
3. Usa formas geométricas bien proporcionadas (<circle>, <rect>, <path>, <ellipse>, <polygon>, <g>) y etiquetas <text> claras en español si son necesarias para explicar el tema (con font-family="Arial, sans-serif" y buen contraste).
4. Asegúrate de que ningún elemento se salga del área 800x600."""

    try:
        respuesta = llamar_gemini_texto(key, prompt_svg)
        match = re.search(r'<svg[\s\S]*?</svg>', respuesta, re.IGNORECASE)
        if match:
            svg_code = match.group(0)
            return svg_code.encode("utf-8"), None
        return None, "La IA no devolvió un formato gráfico válido. Intenta describir la imagen con más detalle."
    except Exception as e:
        return None, str(e)

# ==============================================================================
# 6. GENERADORES DE DOCUMENTOS WORD (.DOCX)
# ==============================================================================
def crear_word_planificacion(datos_doc, cnb_info, mat):
    doc = Document()
    t = doc.add_heading("PLANIFICACIÓN DE APRENDIZAJES - CNB GUATEMALA", level=1)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_info = doc.add_table(rows=4, cols=2)
    t_info.style = 'Table Grid'
    filas_info = [
        ("Establecimiento:", datos_doc["establecimiento"]),
        ("Docente:", datos_doc["docente"]),
        ("Nivel / Grado / Área:", f"{datos_doc['nivel']} | {datos_doc['grado']} - {datos_doc['area']}"),
        ("Tema / Unidad:", f"{datos_doc['tema']} ({datos_doc['unidad']} - {datos_doc['duracion']})")
    ]
    for i, (k, v) in enumerate(filas_info):
        t_info.cell(i, 0).text = k
        t_info.cell(i, 0).paragraphs[0].runs[0].bold = True
        t_info.cell(i, 1).text = v

    doc.add_paragraph("")
    t_cnb = doc.add_table(rows=2, cols=4)
    t_cnb.style = 'Table Grid'
    for idx, h in enumerate(["Competencia", "Indicador de Logro", "Contenidos", "Secuencia Didáctica y Evaluación"]):
        t_cnb.cell(0, idx).text = h
        t_cnb.cell(0, idx).paragraphs[0].runs[0].bold = True

    t_cnb.cell(1, 0).text = mat.get("competencia_oficial", cnb_info.get("competencia", ""))
    t_cnb.cell(1, 1).text = mat.get("indicador_oficial", cnb_info.get("indicador_logro", ""))
    t_cnb.cell(1, 2).text = "\n• ".join([""] + mat.get("contenidos_oficiales", cnb_info.get("contenidos", []))).strip()
    criterios = "\n".join([f"- {c}" for c in mat.get("evaluacion_criterios", [])])
    t_cnb.cell(1, 3).text = (
        f"INICIO:\n{mat.get('actividad_inicio', '')}\n\n"
        f"DESARROLLO:\n{mat.get('actividad_desarrollo', '')}\n\n"
        f"CIERRE:\n{mat.get('actividad_cierre', '')}\n\n"
        f"RECURSOS:\n{', '.join(mat.get('materiales_necesarios', []))}\n\n"
        f"EVALUACIÓN:\n{criterios}"
    )

    doc.add_paragraph("\n\n")
    firmas = doc.add_table(rows=1, cols=2)
    f1 = firmas.cell(0, 0).paragraphs[0]
    f1.text = f"_______________________________\n{datos_doc['docente']}\nDocente"
    f1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    f2 = firmas.cell(0, 1).paragraphs[0]
    f2.text = f"_______________________________\n{datos_doc['director']}\nVo. Bo. Dirección"
    f2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def crear_word_rubrica(datos_doc, rub):
    doc = Document()
    t = doc.add_heading(f"INSTRUMENTO DE EVALUACIÓN: {rub.get('titulo', 'RÚBRICA ANALÍTICA').upper()}", level=1)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Establecimiento: {datos_doc['establecimiento']}   |   Docente: {datos_doc['docente']}")
    doc.add_paragraph(f"Grado y Área: {datos_doc['grado']} - {datos_doc['area']}   |   Tema: {datos_doc['tema']}")
    doc.add_paragraph("Nombre del Estudiante: ___________________________________________   Punteo: ________")

    criterios = rub.get("criterios", [])
    tabla = doc.add_table(rows=len(criterios) + 1, cols=5)
    tabla.style = 'Table Grid'
    for idx, c in enumerate(["Criterio a Evaluar", "Excelente (4 pts)", "Muy Bueno (3 pts)", "Satisfactorio (2 pts)", "Debe Mejorar (1 pt)"]):
        tabla.cell(0, idx).text = c
        tabla.cell(0, idx).paragraphs[0].runs[0].bold = True

    for r_idx, item in enumerate(criterios, start=1):
        tabla.cell(r_idx, 0).text = item.get("criterio", "")
        tabla.cell(r_idx, 0).paragraphs[0].runs[0].bold = True
        tabla.cell(r_idx, 1).text = item.get("excelente", "")
        tabla.cell(r_idx, 2).text = item.get("muy_bueno", "")
        tabla.cell(r_idx, 3).text = item.get("satisfactorio", "")
        tabla.cell(r_idx, 4).text = item.get("debe_mejorar", "")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def crear_word_hoja_trabajo(datos_doc, hoja):
    doc = Document()
    t = doc.add_heading(f"HOJA DE TRABAJO: {hoja.get('titulo', '').upper()}", level=1)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Establecimiento: {datos_doc['establecimiento']}   |   Área: {datos_doc['area']} ({datos_doc['grado']})")
    doc.add_paragraph("Nombre del Estudiante: ___________________________________________   Fecha: ____________")
    p = doc.add_paragraph()
    p.add_run("Instrucciones Generales: ").bold = True
    p.add_run(hoja.get("instrucciones", ""))

    for i, ej in enumerate(hoja.get("ejercicios", []), start=1):
        pe = doc.add_paragraph()
        pe.add_run(f"{i}. ").bold = True
        pe.add_run(ej)
        doc.add_paragraph("Respuesta: ________________________________________________________________________\n")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ==============================================================================
# 7. BARRA LATERAL (PERFIL DEL DOCENTE Y DATOS OFICIALES)
# ==============================================================================
if "herramienta_activa" not in st.session_state:
    st.session_state.herramienta_activa = "🏠 Inicio / Todas las Herramientas"

def cambiar_herramienta(nombre):
    st.session_state.herramienta_activa = nombre

with st.sidebar:
    st.markdown(
        f'<div class="brand-header">{ICONO_QUETZAL_SVG}<div><div class="brand-title">PROFE IA</div><div class="brand-sub">Asistente Pedagógico CNB Guatemala</div></div></div>',
        unsafe_allow_html=True
    )

    st.caption(f"👤 Sesión activa: **{nombre_usuario_actual}**")
    if st.button("Cambiar de docente", use_container_width=True):
        st.session_state.docente_invitado = ""
        st.session_state.pop("mensajes_chat", None)
        st.rerun()

    st.divider()
    with st.expander("🏫 Datos Oficiales del Docente", expanded=True):
        establecimiento = st.text_input("Establecimiento:", placeholder="Ej. E.O.R.M. / Instituto / Colegio...")
        docente = st.text_input("Docente:", value=nombre_usuario_actual)
        director = st.text_input("Director(a):", placeholder="Nombre del Director(a)")
        unidad = st.selectbox("Unidad / Bimestre:", ["I Bimestre", "II Bimestre", "III Bimestre", "IV Bimestre"])
        duracion = st.selectbox("Duración:", ["45 minutos (1 período)", "90 minutos (2 períodos)", "Plan Semanal (5 días)"])

herramienta = st.session_state.herramienta_activa
nombre_saludo = docente.split()[0] if docente.strip() else primer_nombre

# ==============================================================================
# 8. SELECTOR CURRICULAR UNIVERSAL
# ==============================================================================
def bloque_selector_cnb():
    st.markdown("### 📚 Paso 1: Nivel, Grado y Tema del CNB")
    c1, c2, c3 = st.columns([1, 1, 1.3])
    with c1:
        nivel = st.selectbox("Nivel Educativo:", list(CURRICULO_GT.keys()))
    with c2:
        grado = st.selectbox("Grado:", CURRICULO_GT[nivel]["grados"])
    with c3:
        area = st.selectbox("Área Curricular:", CURRICULO_GT[nivel]["areas"])

    col_tema, col_ctx = st.columns([1.3, 1])
    with col_tema:
        tema = st.text_input("🔎 Tema específico o contenido a trabajar:", placeholder="Ej. Las fracciones, La célula, Derechos Humanos...")
    with col_ctx:
        contexto = st.text_input("Contexto del aula (opcional):", "Aula en Guatemala, recursos accesibles, enfoque participativo.")

    cnb_encontrado = {}
    if tema.strip():
        registros = buscar_en_cnb_vivo(grado, area, tema)
        if registros:
            st.success(f"✅ Se encontraron {len(registros)} referencias en {registros[0]['fuente']}.")
            opciones = [f"{i+1}. {r['indicador_logro'][:90]}..." for i, r in enumerate(registros)]
            idx = st.selectbox("Indicador de Logro encontrado:", range(len(opciones)), format_func=lambda x: opciones[x])
            cnb_encontrado = registros[idx]
        else:
            st.info("💡 La IA alineará automáticamente este tema a las Competencias e Indicadores oficiales del CNB MINEDUC.")

    return {
        "establecimiento": establecimiento or "Establecimiento Educativo",
        "docente": docente or nombre_usuario_actual,
        "director": director or "Dirección",
        "unidad": unidad, "duracion": duracion, "nivel": nivel, "grado": grado,
        "area": area, "tema": tema.strip() if tema.strip() else f"Contenidos fundamentales de {area}",
        "contexto": contexto, "cnb_vivo": cnb_encontrado
    }

# ==============================================================================
# 9. VISTAS PRINCIPALES
# ==============================================================================
if herramienta != "🏠 Inicio / Todas las Herramientas":
    st.button("⬅️ Volver a Todas las Herramientas", on_click=cambiar_herramienta, args=("🏠 Inicio / Todas las Herramientas",))

if herramienta == "🏠 Inicio / Todas las Herramientas":
    st.markdown(
        f'<div class="hero-box"><h1>👋 ¡Hola Profe {nombre_saludo}, qué deseas crear hoy?</h1><p>Haz clic directamente en cualquiera de las herramientas alineadas al CNB de Guatemala (Preprimaria a Diversificado).</p></div>',
        unsafe_allow_html=True
    )

    st.subheader("✨ Selecciona una herramienta para comenzar")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="tool-card"><h4>📋 Planificador de Clase CNB</h4><p>Secuencias didácticas (Inicio, Desarrollo y Cierre) alineadas al CNB en Word listo para firmar.</p></div>', unsafe_allow_html=True)
        st.button("Abrir Planificador →", key="btn_plan", on_click=cambiar_herramienta, args=("📋 Planificador de Clase CNB",), use_container_width=True, type="primary")
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="tool-card"><h4>🎨 Recursos Visuales y Pizarrón</h4><p>Ilustraciones educativas o guías paso a paso para dibujar esquemas didácticos en tu aula.</p></div>', unsafe_allow_html=True)
        st.button("Abrir Recursos Visuales →", key="btn_img", on_click=cambiar_herramienta, args=("🎨 Generador de Imágenes Educativas",), use_container_width=True)

    with col2:
        st.markdown('<div class="tool-card"><h4>📊 Generador de Rúbricas</h4><p>Rúbricas analíticas con 4 niveles de desempeño (Excelente a Debe Mejorar) en tabla Word.</p></div>', unsafe_allow_html=True)
        st.button("Abrir Rúbricas →", key="btn_rub", on_click=cambiar_herramienta, args=("📊 Generador de Rúbricas",), use_container_width=True, type="primary")
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="tool-card"><h4>💬 Chat Asistente Pedagógico</h4><p>Pide dinámicas, adaptaciones curriculares, ejemplos guatemaltecos o ideas multigrado.</p></div>', unsafe_allow_html=True)
        st.button("Abrir Chat Pedagógico →", key="btn_chat", on_click=cambiar_herramienta, args=("💬 Chat con Asistente Pedagógico",), use_container_width=True)

    with col3:
        st.markdown('<div class="tool-card"><h4>📝 Hojas de Trabajo</h4><p>Guías de ejercicios y actividades listas para imprimir y entregar a tus estudiantes.</p></div>', unsafe_allow_html=True)
        st.button("Abrir Hojas de Trabajo →", key="btn_hoja", on_click=cambiar_herramienta, args=("📝 Hojas de Trabajo y Ejercicios",), use_container_width=True, type="primary")

elif herramienta == "📋 Planificador de Clase CNB":
    st.header("📋 Planificador de Clase Oficial (CNB Guatemala)")
    datos = bloque_selector_cnb()
    if st.button("🚀 GENERAR PLANIFICACIÓN DE CLASE (.DOCX)", type="primary", use_container_width=True):
        with st.spinner("🧠 Diseñando planificación alineada al CNB de Guatemala..."):
            try:
                ref_cnb = f"Referencia CNB: {datos['cnb_vivo']}" if datos["cnb_vivo"] else "Deduce Competencia, Indicador de Logro y Contenidos oficiales del CNB MINEDUC Guatemala."
                prompt = f"""Experto pedagogo MINEDUC Guatemala (CNB). Crea planificación para {datos['nivel']} | {datos['grado']} | {datos['area']}, Tema: {datos['tema']}, Duración: {datos['duracion']}, Contexto: {datos['contexto']}. {ref_cnb}.
                Responde SOLO JSON válido con claves: "competencia_oficial", "indicador_oficial", "contenidos_oficiales" (lista), "actividad_inicio", "actividad_desarrollo", "actividad_cierre", "materiales_necesarios" (lista), "evaluacion_criterios" (lista)."""
                res = llamar_gemini_json(api_key, prompt)
                st.success("✅ ¡Planificación lista!")
                st.download_button("📥 DESCARGAR PLANIFICACIÓN EN WORD (.DOCX)", data=crear_word_planificacion(datos, datos["cnb_vivo"], res), file_name=f"Planificacion_{datos['grado'].replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
                st.markdown(f'<div class="cnb-pill"><b>Competencia CNB:</b> {res.get("competencia_oficial")}<br><b>Indicador de Logro:</b> {res.get("indicador_oficial")}</div>', unsafe_allow_html=True)
                st.subheader("1. Inicio")
                st.write(res.get("actividad_inicio"))
                st.subheader("2. Desarrollo")
                st.write(res.get("actividad_desarrollo"))
                st.subheader("3. Cierre")
                st.write(res.get("actividad_cierre"))
            except Exception as e:
                st.error(f"❌ {e}")

elif herramienta == "📊 Generador de Rúbricas":
    st.header("📊 Generador de Rúbricas de Evaluación (CNB)")
    datos = bloque_selector_cnb()
    actividad_eval = st.text_input("Actividad o producto a evaluar:", "Exposición oral / Maqueta / Resolución de problemas")
    if st.button("🚀 GENERAR RÚBRICA EN WORD (.DOCX)", type="primary", use_container_width=True):
        with st.spinner("📊 Construyendo matriz de evaluación..."):
            try:
                prompt = f"""Crea rúbrica analítica CNB Guatemala para {datos['grado']} - {datos['area']}, Tema: {datos['tema']}, Actividad: {actividad_eval}.
                Responde SOLO JSON válido: {{"titulo": "...", "criterios": [{{"criterio": "...", "excelente": "...", "muy_bueno": "...", "satisfactorio": "...", "debe_mejorar": "..."}}]}} con 4 criterios."""
                rub = llamar_gemini_json(api_key, prompt)
                st.success("✅ ¡Rúbrica generada!")
                st.download_button("📥 DESCARGAR RÚBRICA EN WORD (.DOCX)", data=crear_word_rubrica(datos, rub), file_name=f"Rubrica_{datos['tema'][:20].replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
                st.table(rub.get("criterios", []))
            except Exception as e:
                st.error(f"❌ {e}")

elif herramienta == "📝 Hojas de Trabajo y Ejercicios":
    st.header("📝 Generador de Hojas de Trabajo para Estudiantes")
    datos = bloque_selector_cnb()
    cant_ej = st.slider("Cantidad de ejercicios:", 3, 10, 5)
    if st.button("🚀 GENERAR HOJA DE TRABAJO (.DOCX)", type="primary", use_container_width=True):
        with st.spinner("📝 Redactando ejercicios contextualizados a Guatemala..."):
            try:
                prompt = f"""Diseña hoja de trabajo para estudiantes de Guatemala: {datos['nivel']} | {datos['grado']} | {datos['area']}, Tema: {datos['tema']}, Contexto: {datos['contexto']}, Cantidad: {cant_ej} ejercicios.
                Responde SOLO JSON válido: {{"titulo": "...", "instrucciones": "...", "ejercicios": ["..."], "clave_docente": ["..."]}}."""
                hoja = llamar_gemini_json(api_key, prompt)
                st.success("✅ ¡Hoja de trabajo lista!")
                st.download_button("📥 DESCARGAR HOJA DE TRABAJO (.DOCX)", data=crear_word_hoja_trabajo(datos, hoja), file_name=f"Hoja_Trabajo_{datos['tema'][:20].replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
                st.subheader(hoja.get("titulo", ""))
                st.write(f"**Instrucciones:** {hoja.get('instrucciones', '')}")
                for i, ej in enumerate(hoja.get("ejercicios", []), 1):
                    st.markdown(f"**{i}.** {ej}")
                with st.expander("🔑 Ver Clave de Respuestas (Solo para el Docente)"):
                    for i, resp in enumerate(hoja.get("clave_docente", []), 1):
                        st.write(f"**{i}.** {resp}")
            except Exception as e:
                st.error(f"❌ {e}")

elif herramienta == "🎨 Generador de Imágenes Educativas":
    st.header("🎨 Generador de Recursos Visuales e Ilustraciones para el Aula")

    col_est, col_desc = st.columns([1, 2])
    with col_est:
        estilo_img = st.selectbox(
            "Estilo del recurso visual:",
            [
                "Ilustración libro de texto",
                "Dibujo para colorear (Blanco y negro)",
                "Esquema / Mapa visual de Ciencias",
                "Contexto Cultural de Guatemala"
            ]
        )
    with col_desc:
        desc_img = st.text_area(
            "¿Qué imagen necesitas para tu clase?",
            placeholder="Ej. Las partes de la planta con sus nombres, El ciclo del agua con volcanes de Guatemala, Las fracciones 1/2 y 1/4...",
            height=100
        )

    if st.button("🎨 GENERAR ILUSTRACIÓN EDUCATIVA", type="primary", use_container_width=True):
        if not desc_img.strip():
            st.warning("Por favor describe qué imagen deseas crear.")
        else:
            with st.spinner("🖌️ Dibujando recurso visual educativo..."):
                svg_bytes, error_img = generar_imagen_educativa(api_key, desc_img, estilo_img)
                if svg_bytes:
                    svg_texto = svg_bytes.decode("utf-8")
                    st.markdown(
                        f'<div style="background:white; padding:16px; border-radius:12px; border:1px solid #cbd5e1; text-align:center;">{svg_texto}</div>',
                        unsafe_allow_html=True
                    )
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.download_button(
                        "📥 DESCARGAR ILUSTRACIÓN VECTORIAL (.SVG - SE ABRE EN WORD Y NAVEGADOR)",
                        data=svg_bytes,
                        file_name="Recurso_Visual_ProfeIA.svg",
                        mime="image/svg+xml",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    st.error(f"❌ No se pudo generar la ilustración: {error_img}")

elif herramienta == "💬 Chat con Asistente Pedagógico":
    st.header("💬 Chat con tu Asistente Pedagógico CNB")
    if "mensajes_chat" not in st.session_state:
        st.session_state.mensajes_chat = [{"role": "assistant", "content": f"¡Hola Profe {nombre_saludo}! Soy tu asistente pedagógico especializado en el CNB de Guatemala. ¿En qué te ayudo hoy?"}]
    for msg in st.session_state.mensajes_chat:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])
    pregunta = st.chat_input("Escribe tu consulta pedagógica aquí...")
    if pregunta:
        st.session_state.mensajes_chat.append({"role": "user", "content": pregunta})
        with st.chat_message("user"):
            st.write(pregunta)
        with st.chat_message("assistant"):
            with st.spinner("Pensando..."):
                respuesta = llamar_gemini_texto(api_key, f"Eres un asesor pedagógico experto en el CNB de Guatemala (MINEDUC). Responde de forma práctica y aplicable.\n\nConsulta: {pregunta}")
                st.write(respuesta)
                st.session_state.mensajes_chat.append({"role": "assistant", "content": respuesta})
