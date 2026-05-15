"""
Generador de Auditorías GEO en PPTX
Entradas: empresa, fecha, logo, screenshot LLMs Pulse
Claude extrae todo lo demás y genera el PPTX automáticamente.
"""

import os
import io
import json
import copy
import base64
import requests
import concurrent.futures
import streamlit as st
from datetime import date
from anthropic import Anthropic
from pptx import Presentation
from pptx.util import Emu
from pptx.oxml.ns import qn
from lxml import etree

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.pptx")

# API key: Streamlit Secrets (cloud) o variable de entorno (local)
def _load_env_file():
    """Lee el .env del directorio del script y setea os.environ."""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_path):
        return
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            os.environ[key.strip()] = val.strip()

_load_env_file()


def _get_key(name: str) -> str:
    val = os.environ.get(name, "")
    if not val:
        try:
            val = st.secrets.get(name, "")
        except Exception:
            pass
    return val


ANTHROPIC_KEY = _get_key("ANTHROPIC_API_KEY")
OPENAI_KEY    = _get_key("OPENAI_API_KEY")
GEMINI_KEY    = _get_key("GEMINI_API_KEY")
GROQ_KEY      = _get_key("GROQ_API_KEY")
EMU_PER_SCORE_PT = 54864

MONTHS_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# ─────────────────────────────────────────────────────────────
# PROMPT
# ─────────────────────────────────────────────────────────────

PROMPT = """\
Eres un consultor GEO senior redactando un informe profesional para el cliente final.

REGLA ABSOLUTA — NUNCA menciones en ningún campo del JSON:
- Nombres de herramientas o APIs: LLMs Pulse, Ahrefs, OpenAI, Anthropic, Groq, Llama, Claude API, ChatGPT API, Gemini API
- Errores técnicos, fallos de módulos o problemas de conexión
- El hecho de que los datos vienen de capturas de pantalla o análisis automatizados
- Frases como "la herramienta indica", "según LLMs Pulse", "el análisis confirma"

Escribe siempre como un consultor experto: "La marca no aparece en los modelos de IA", "Los modelos no reconocen la empresa como referente", "Se recomienda crear contenido estructurado para mejorar la visibilidad en IA".

Analiza esta captura de pantalla de visibilidad de marca en modelos de IA.

{contexto}

ANTES DE NADA — COMPETITORS (lee esto primero):
Busca en la imagen una sección titulada "Top Competitors" (columna derecha, parte superior). Anota los nombres exactos de las empresas que aparecen ahí y sus descripciones en inglés. Serán empresas del mismo sector y país que la empresa auditada — nombres locales/regionales, NUNCA consultoras globales como McKinsey, Deloitte, Accenture, etc. Si no lees la sección con claridad, devuelve "competitors": []. PROHIBIDO inventar o usar conocimiento propio: solo lo que está escrito en la imagen.

Ahora extrae todos los datos numéricos visibles y genera un análisis de auditoría GEO completo y profesional en español.

Devuelve EXCLUSIVAMENTE un JSON válido, sin texto antes ni después:

{{
  "score_global": <entero 0-100>,
  "chatgpt_score": <entero 0-100>,
  "gemini_score": <entero 0-100>,
  "claude_score": <entero 0-100>,
  "perplexity_score": <entero 0-100>,
  "resumen": "<2-3 frases sobre el estado de visibilidad en IA con el score y el nivel>",
  "competitive_desc": "<frase corta sobre la posición competitiva en búsquedas de IA>",
  "competitors": [
    {{"name": "<nombre exacto leído en la imagen, sección Top Competitors>", "stars": <entero 1-5>, "desc": "<descripción leída en la imagen, traducida al español>"}},
    {{"name": "<nombre exacto leído en la imagen, sección Top Competitors>", "stars": <entero 1-5>, "desc": "<descripción leída en la imagen, traducida al español>"}},
    {{"name": "<nombre exacto leído en la imagen, sección Top Competitors>", "stars": <entero 1-5>, "desc": "<descripción leída en la imagen, traducida al español>"}}
  ],
  "chatgpt_hallazgos": [
    {{"title": "<hallazgo crítico 1, máx 40 chars>", "detail": "<2-3 líneas separadas por \\n, máx 160 chars total>"}},
    {{"title": "<hallazgo crítico 2, máx 40 chars>", "detail": "<2-3 líneas separadas por \\n, máx 160 chars total>"}},
    {{"title": "<hallazgo crítico 3, máx 40 chars>", "detail": "<2-3 líneas separadas por \\n, máx 160 chars total>"}}
  ],
  "chatgpt_diagnosticos": [
    {{"label": "Visibilidad ChatGPT", "value": "<Nula|Baja|Media|Alta>"}},
    {{"label": "<diagnóstico clave 2>", "value": "<valor corto, máx 20 chars>"}},
    {{"label": "<diagnóstico clave 3>", "value": "<valor corto, máx 20 chars>"}}
  ],
  "chatgpt_prioridad": "<ACCIÓN PRIORITARIA CHATGPT EN MAYÚSCULAS, máx 55 chars>",
  "gemini_hallazgos": [
    {{"title": "<hallazgo crítico 1, máx 40 chars>", "detail": "<2-3 líneas separadas por \\n, máx 160 chars total>"}},
    {{"title": "<hallazgo crítico 2, máx 40 chars>", "detail": "<2-3 líneas separadas por \\n, máx 160 chars total>"}},
    {{"title": "<hallazgo crítico 3, máx 40 chars>", "detail": "<2-3 líneas separadas por \\n, máx 160 chars total>"}}
  ],
  "gemini_diagnosticos": [
    {{"label": "Visibilidad Gemini", "value": "<Nula|Baja|Media|Alta>"}},
    {{"label": "<diagnóstico clave 2>", "value": "<valor corto, máx 20 chars>"}},
    {{"label": "<diagnóstico clave 3>", "value": "<valor corto, máx 20 chars>"}}
  ],
  "gemini_prioridad": "<ACCIÓN PRIORITARIA GEMINI EN MAYÚSCULAS, máx 55 chars>",
  "strengths": [
    "<fortaleza actual 1>",
    "<fortaleza actual 2>",
    "<fortaleza actual 3>",
    "<fortaleza actual 4>"
  ],
  "opportunities": [
    "<oportunidad de mejora 1>",
    "<oportunidad de mejora 2>",
    "<oportunidad de mejora 3>",
    "<oportunidad de mejora 4>"
  ],
  "recommendations": [
    {{"title": "<acción concreta 1, máx 50 caracteres>", "desc": "<1 frase con acción + plazo, máx 90 caracteres>", "priority": 3}},
    {{"title": "<acción concreta 2, máx 50 caracteres>", "desc": "<1 frase con acción + plazo, máx 90 caracteres>", "priority": 3}},
    {{"title": "<acción concreta 3, máx 50 caracteres>", "desc": "<1 frase con acción + plazo, máx 90 caracteres>", "priority": 3}},
    {{"title": "<acción concreta 4, máx 50 caracteres>", "desc": "<1 frase con acción + plazo, máx 90 caracteres>", "priority": 2}},
    {{"title": "<acción concreta 5, máx 50 caracteres>", "desc": "<1 frase con acción + plazo, máx 90 caracteres>", "priority": 2}}
  ],
  "steps": [
    {{"title": "Revisar hallazgos", "desc": "<1 frase, máx 90 caracteres>"}},
    {{"title": "Priorizar iniciativas", "desc": "<1 frase, máx 90 caracteres>"}},
    {{"title": "Implementar Quick Wins", "desc": "<1 frase con 2-3 acciones clave, máx 90 caracteres>"}},
    {{"title": "Monitorear progreso", "desc": "<1 frase, máx 90 caracteres>"}}
  ]
}}

REGLA ESTRICTA para "competitors": Localiza en la imagen la sección llamada "Top Competitors". Contiene entre 2 y 5 empresas, cada una con su nombre y una descripción en inglés. DEBES copiar literalmente los nombres y descripciones que leas en esa sección de la imagen — NO uses conocimiento externo, NO inventes competidores, NO rellenes con empresas que no estén visibles. Si la sección no está visible con claridad en la imagen, devuelve "competitors": []. Para "stars" (1-5): estima según la descripción visible — presencia fuerte/nacional = 4-5, moderada/local = 2-3. Toma únicamente los 3 primeros que aparezcan. Las descripciones tradúcelas al español.
Si no ves un score numérico exacto, estímalo por los indicadores visuales.
Textos en español, concretos y accionables. Priority: 1=baja, 2=media, 3=alta.

REGLA CRÍTICA — COHERENCIA SCORES vs TEXTO: El campo "resumen" y todos los textos de análisis DEBEN ser coherentes con los scores numéricos que asignas. Escala obligatoria: ≥80 = presencia alta/amplia, 60-79 = presencia media-alta/reconocimiento notable, 40-59 = presencia media/reconocimiento parcial, 20-39 = presencia baja/menciones mínimas, <20 = sin presencia/desconocida. NUNCA escribas que un modelo "no reconoce" o "no tiene información" sobre la empresa si le has dado un score ≥60. NUNCA escribas que un modelo "reconoce" o "menciona" la empresa si le has dado un score <30. El modelo con el score más alto debe describirse siempre como el que mejor reconoce la marca; el de score más bajo, como el que menos.\
"""


# ─────────────────────────────────────────────────────────────
# FETCH COMPETITORS FROM URL
# ─────────────────────────────────────────────────────────────

def fetch_report_data_from_url(url: str) -> dict:
    """Fetches LLMs Pulse report and extracts Top Competitors + Recommendations."""
    empty = {"competitors": [], "recommendations": []}
    try:
        resp = requests.get(url, timeout=15, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
        })
        resp.raise_for_status()
        html  = resp.text
        lower = html.lower()

        def _extract_chunk(keyword, window=8000):
            idx = lower.find(keyword)
            if idx == -1:
                return ""
            return html[max(0, idx - 300): min(len(html), idx + window)]

        # Localizar la sección de competidores como ancla
        comp_idx = lower.find("top competitor")
        if comp_idx == -1:
            comp_idx = lower.find("competitor")

        if comp_idx != -1:
            # Oportunidades (~9000 chars antes) y recomendaciones (~19000 chars después)
            # Extraer un chunk grande que engloba las tres secciones
            big_start = max(0, comp_idx - 12000)
            big_end   = min(len(html), comp_idx + 28000)
            combined  = f"--- REPORT ANALYSIS SECTIONS ---\n{html[big_start:big_end]}"
        else:
            # Fallback: keywords individuales
            comp_chunk = _extract_chunk("competitor")
            rec_chunk  = (
                _extract_chunk("next step") or _extract_chunk("recommend") or
                _extract_chunk("action") or _extract_chunk("suggest")
            )
            opp_chunk  = (
                _extract_chunk("improvement") or _extract_chunk("opportunit") or
                _extract_chunk("gap") or _extract_chunk("area")
            )
            combined = (
                f"--- COMPETITORS ---\n{comp_chunk}\n\n"
                f"--- RECOMMENDATIONS ---\n{rec_chunk}\n\n"
                f"--- OPPORTUNITIES ---\n{opp_chunk}"
            )

        client = Anthropic(api_key=ANTHROPIC_KEY)
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=1200,
            messages=[{
                "role": "user",
                "content": (
                    "Analiza este HTML de un informe LLMs Pulse y extrae:\n"
                    "1) COMPETITORS: hasta 3 competidores con nombre y descripción.\n"
                    "2) RECOMMENDATIONS: hasta 3 recomendaciones (texto exacto de cada una).\n"
                    "3) OPPORTUNITIES: hasta 4 oportunidades de mejora (título + descripción corta).\n\n"
                    "Devuelve SOLO este JSON, sin texto extra:\n"
                    '{"competitors": [{"name": "...", "desc": "..."}], '
                    '"recommendations": ["texto 1", "texto 2", "texto 3"], '
                    '"opportunities": ["oportunidad 1", "oportunidad 2", "oportunidad 3", "oportunidad 4"]}\n\n'
                    f"HTML:\n{combined}"
                ),
            }],
        )
        raw = response.content[0].text.strip()
        if "```" in raw:
            raw = raw.split("```")[1].split("```")[0].strip()
            if raw.startswith("json"):
                raw = raw[4:].strip()
        return json.loads(raw)
    except Exception:
        return empty


# ─────────────────────────────────────────────────────────────
# CLAUDE API
# ─────────────────────────────────────────────────────────────

def analyze_with_claude(image_bytes: bytes, empresa: str, competitors_data: list, recommendations_data: list, opportunities_data: list, observaciones: str = "") -> dict:
    client   = Anthropic(api_key=ANTHROPIC_KEY)
    media_type = "image/png" if image_bytes[:4] == b"\x89PNG" else "image/jpeg"
    b64      = base64.standard_b64encode(image_bytes).decode("utf-8")
    contexto = f"La empresa auditada es: {empresa.strip()}." if empresa.strip() else ""
    if observaciones.strip():
        contexto += f" OBSERVACIONES DEL CLIENTE (tenlas muy en cuenta para el análisis y las recomendaciones): {observaciones.strip()}"
    if competitors_data:
        partes = []
        for i, c in enumerate(competitors_data, 1):
            if isinstance(c, dict):
                partes.append(f"{i}) {c.get('name', '')} — {c.get('desc', '')}")
            else:
                partes.append(f"{i}) {c}")
        contexto += (
            " Los competidores principales extraídos del informe LLMs Pulse son EXACTAMENTE estos "
            f"(úsalos en el campo 'competitors' en este orden, sin inventar otros): {'; '.join(partes)}."
        )
    if recommendations_data:
        recs = "; ".join(f"{i+1}) {r}" for i, r in enumerate(recommendations_data))
        contexto += (
            f" Las siguientes recomendaciones están extraídas DIRECTAMENTE del informe LLMs Pulse "
            f"— úsalas como base para los primeros {len(recommendations_data)} elementos del campo 'recommendations' "
            f"(tradúcelas al español y añade título corto y prioridad): {recs}."
        )
    if opportunities_data:
        opps = "; ".join(f"{i+1}) {o}" for i, o in enumerate(opportunities_data))
        contexto += (
            f" Las siguientes oportunidades de mejora están extraídas DIRECTAMENTE del informe LLMs Pulse "
            f"— úsalas en el campo 'opportunities' (tradúcelas al español): {opps}."
        )

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": PROMPT.format(contexto=contexto)},
            ],
        }],
    )

    raw = response.content[0].text.strip()
    if "```json" in raw:
        raw = raw.split("```json")[1].split("```")[0].strip()
    elif "```" in raw:
        raw = raw.split("```")[1].split("```")[0].strip()

    return json.loads(raw)


# ─────────────────────────────────────────────────────────────
# HELPERS PPTX
# ─────────────────────────────────────────────────────────────

def score_level(s: int) -> str:
    if s >= 80: return "alto"
    if s >= 60: return "medio-alto"
    if s >= 40: return "medio"
    if s >= 20: return "bajo"
    return "muy bajo"


def score_stars(s: int) -> str:
    return "⭐" * max(1, round(s / 20))


def model_label(score: int, model: str) -> str:
    opts = {
        "ChatGPT":    {80: "Alto - Amplia presencia espontánea", 60: "Medio-Alto - Buena cobertura",   40: "Medio - Menciones moderadas",     20: "Bajo - Menciones mínimas",       0: "Muy Bajo - Sin presencia"},
        "Gemini":     {80: "Alto - Reconocimiento sólido",       60: "Medio-Alto - Presencia notable", 40: "Medio - Reconocimiento parcial",   20: "Bajo - Poco reconocimiento",    0: "Muy Bajo - Sin presencia"},
        "Claude":     {80: "Alto - Información completa",        60: "Medio-Alto - Buena información", 40: "Medio - Información parcial",      20: "Bajo - Información limitada",   0: "Muy Bajo - Sin información"},
        "Perplexity": {80: "Alto - Fuentes abundantes",          60: "Medio-Alto - Buenas fuentes",    40: "Medio - Fuentes moderadas",        20: "Bajo - Pocas fuentes disponibles", 0: "Muy Bajo - Sin fuentes"},
        "Grok / Llama": {80: "Alto - Amplia cobertura",         60: "Medio-Alto - Buena cobertura",   40: "Medio - Cobertura parcial",        20: "Bajo - Cobertura limitada",        0: "Muy Bajo - Sin presencia"},
    }
    for thr, label in sorted(opts.get(model, {}).items(), reverse=True):
        if score >= thr:
            return label
    return f"{score_level(score).capitalize()} - {model}"


def find_shape(slide, name: str):
    for s in slide.shapes:
        if s.name == name:
            return s
    return None


def set_run(slide, name: str, text: str):
    shape = find_shape(slide, name)
    if shape and hasattr(shape, "text_frame"):
        paras = shape.text_frame.paragraphs
        if paras and paras[0].runs:
            paras[0].runs[0].text = text


def set_body(slide, name: str, lines: list):
    shape = find_shape(slide, name)
    if not shape or not hasattr(shape, "text_frame"):
        return
    txBody = shape.text_frame._txBody

    template_rPr = None
    for p in txBody.findall(qn("a:p")):
        for r in p.findall(qn("a:r")):
            rPr = r.find(qn("a:rPr"))
            if rPr is not None:
                template_rPr = copy.deepcopy(rPr)
            break
        if template_rPr is not None:
            break

    template_pPr = None
    fp = txBody.find(qn("a:p"))
    if fp is not None:
        pPr = fp.find(qn("a:pPr"))
        if pPr is not None:
            template_pPr = copy.deepcopy(pPr)

    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)

    for line in lines:
        is_header = line.strip().endswith(":")
        new_p = etree.SubElement(txBody, qn("a:p"))
        if template_pPr is not None:
            pPr_copy = copy.deepcopy(template_pPr)
            pPr_copy.set("algn", "just")
            new_p.append(pPr_copy)
        else:
            pPr_new = etree.SubElement(new_p, qn("a:pPr"))
            pPr_new.set("algn", "just")
        if line:
            new_r = etree.SubElement(new_p, qn("a:r"))
            if template_rPr is not None:
                rPr_copy = copy.deepcopy(template_rPr)
                if is_header:
                    rPr_copy.set("b", "1")
                new_r.append(rPr_copy)
            elif is_header:
                rPr_new = etree.SubElement(new_r, qn("a:rPr"))
                rPr_new.set("b", "1")
            new_t = etree.SubElement(new_r, qn("a:t"))
            new_t.text = line

    if not txBody.findall(qn("a:p")):
        etree.SubElement(txBody, qn("a:p"))


def update_bar(slide, name: str, score: int):
    shape = find_shape(slide, name)
    if shape:
        shape.width = Emu(max(1, score) * EMU_PER_SCORE_PT)


def replace_logo(slide, name: str, img_bytes: bytes):
    from PIL import Image as PILImage
    shape = find_shape(slide, name)
    if not shape or shape.shape_type != 13:
        return
    ns_a  = "http://schemas.openxmlformats.org/drawingml/2006/main"
    blip  = shape._element.find(f".//{{{ns_a}}}blip")
    if blip is None:
        return
    r_ns  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rId   = blip.get(f"{{{r_ns}}}embed")
    if not rId or rId not in slide.part.rels:
        return

    # Ajustar shape al aspect ratio real del logo sin distorsionar
    try:
        img = PILImage.open(io.BytesIO(img_bytes))
        img_w, img_h = img.size
        max_size = 2651760  # ancho/alto máximo del placeholder (EMU)
        ratio = img_w / img_h
        if ratio >= 1:  # más ancho que alto
            new_w = max_size
            new_h = int(max_size / ratio)
        else:           # más alto que ancho
            new_h = max_size
            new_w = int(max_size * ratio)
        # Posicionar: centrado horizontalmente, en el tercio superior del panel
        orig_left = shape.left
        shape.width  = Emu(new_w)
        shape.height = Emu(new_h)
        shape.left   = orig_left + (max_size - new_w) // 2
        shape.top    = Emu(1400000)  # tercio superior del panel navy
    except Exception:
        pass  # si falla, deja las dimensiones originales

    slide.part.rels[rId].target_part._blob = img_bytes


# ─────────────────────────────────────────────────────────────
# HELPER — arco SVG para gauge circular
# ─────────────────────────────────────────────────────────────

def add_arc_xml(slide, left, top, width, height, adj1, adj2, color_hex, line_w_emu):
    """Inserta un arco (prstGeom arc) vía XML directo en el spTree del slide."""
    sp_tree = slide.shapes._spTree
    max_id = 0
    for el in sp_tree.iter():
        try:
            max_id = max(max_id, int(el.get("id", 0)))
        except (ValueError, TypeError):
            pass
    sp_id = max_id + 1
    xml = (
        f'<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="arc{sp_id}"/>'
        f'<p:cNvSpPr/><p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        f'<a:prstGeom prst="arc"><a:avLst>'
        f'<a:gd name="adj1" fmla="val {adj1}"/>'
        f'<a:gd name="adj2" fmla="val {adj2}"/>'
        f'</a:avLst></a:prstGeom>'
        f'<a:noFill/>'
        f'<a:ln w="{line_w_emu}" cap="rnd">'
        f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
        f'</a:ln>'
        f'</p:spPr>'
        f'</p:sp>'
    )
    from lxml import etree as _etree
    sp_tree.append(_etree.fromstring(xml))


# ─────────────────────────────────────────────────────────────
# HELPER — diseño panel score + hallazgos (slides 5 y 6)
# ─────────────────────────────────────────────────────────────

def build_analysis_slide(s, score: int, model_name: str, hallazgos: list, diagnosticos: list, prioridad: str):
    """Dibuja slides 5/6: panel izquierdo (score+gauge+diagnósticos) y panel derecho (hallazgos)."""
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    NAVY  = RGBColor(0x0F, 0x1B, 0x3D)
    BLUE  = RGBColor(0x23, 0x6F, 0xAB)
    LGRAY = RGBColor(0xF0, 0xF2, 0xF5)
    DGRAY = RGBColor(0xCC, 0xCC, 0xCC)
    GRAY  = RGBColor(0x88, 0x88, 0x88)
    DARK  = RGBColor(0x22, 0x22, 0x22)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    SW, SH = 9144000, 5143500

    # Limpiar shapes del template
    sp_tree = s.shapes._spTree
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    # Fondo blanco global
    bg = s.shapes.add_shape(1, Emu(0), Emu(0), Emu(SW), Emu(SH))
    bg.fill.solid(); bg.fill.fore_color.rgb = WHITE; bg.line.fill.background()

    # ── PANEL IZQUIERDO ──────────────────────────────────────
    PL_L, PL_W, PL_T = 70000, 2560000, 70000
    PL_H = SH - 140000
    ML   = PL_L + 200000          # margen contenido dentro del panel
    PC   = PL_L + PL_W // 2       # centro x del panel

    panel = s.shapes.add_shape(1, Emu(PL_L), Emu(PL_T), Emu(PL_W), Emu(PL_H))
    panel.fill.solid(); panel.fill.fore_color.rgb = LGRAY; panel.line.fill.background()

    # "PUNTUACIÓN" label
    tb = s.shapes.add_textbox(Emu(ML), Emu(200000), Emu(PL_W - 300000), Emu(160000))
    p = tb.text_frame.paragraphs[0]; p.text = "PUNTUACIÓN"
    r = p.runs[0]; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = GRAY

    # Score número grande
    sb = s.shapes.add_textbox(Emu(ML), Emu(360000), Emu(900000), Emu(520000))
    sb.text_frame.word_wrap = False
    p = sb.text_frame.paragraphs[0]; p.text = str(score)
    r = p.runs[0]; r.font.size = Pt(52); r.font.bold = True; r.font.color.rgb = NAVY

    # "/100"
    lb = s.shapes.add_textbox(Emu(ML + 800000), Emu(490000), Emu(400000), Emu(220000))
    p = lb.text_frame.paragraphs[0]; p.text = "/100"
    r = p.runs[0]; r.font.size = Pt(13); r.font.color.rgb = GRAY

    # Gauge circular
    GD = 800000   # diámetro
    GL = PC - GD // 2
    GT = 950000
    # Arco gris (pista completa: 300°, de 7 en punto a 5 en punto)
    # En OOXML: 0°=3h, aumenta en sentido horario.  7h=120°, 5h=60° (largo=300° horario)
    A1 = 7200000   # 120° × 60000
    A2 = 25200000  # 120° + 300° = 420° × 60000 (wrap a 60°)
    add_arc_xml(s, GL, GT, GD, GD, A1, A2, "CCCCCC", 110000)
    # Arco azul (score %)
    score_sweep = int(score / 100 * (A2 - A1))
    if score_sweep > 0:
        add_arc_xml(s, GL, GT, GD, GD, A1, A1 + score_sweep, "236FAB", 110000)

    # "DIAGNÓSTICO" heading
    DIAG_T = GT + GD + 140000
    dtb = s.shapes.add_textbox(Emu(ML), Emu(DIAG_T), Emu(PL_W - 300000), Emu(190000))
    p = dtb.text_frame.paragraphs[0]; p.text = "DIAGNÓSTICO"
    r = p.runs[0]; r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = BLUE

    # Separador
    sep = s.shapes.add_shape(1, Emu(ML), Emu(DIAG_T + 200000), Emu(PL_W - 350000), Emu(12000))
    sep.fill.solid(); sep.fill.fore_color.rgb = DGRAY; sep.line.fill.background()

    # 3 ítems de diagnóstico
    for j, diag in enumerate(diagnosticos[:3]):
        it = DIAG_T + 260000 + j * 440000
        ltb = s.shapes.add_textbox(Emu(ML), Emu(it), Emu(PL_W - 300000), Emu(160000))
        p = ltb.text_frame.paragraphs[0]; p.text = diag.get("label", "")[:35]
        r = p.runs[0]; r.font.size = Pt(10); r.font.color.rgb = GRAY

        vtb = s.shapes.add_textbox(Emu(ML), Emu(it + 165000), Emu(PL_W - 300000), Emu(210000))
        p = vtb.text_frame.paragraphs[0]; p.text = diag.get("value", "")[:25]
        r = p.runs[0]; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = DARK

    # ── PANEL DERECHO ────────────────────────────────────────
    RP_L = PL_L + PL_W + 120000
    RP_W = SW - RP_L - 80000

    # Título "Hallazgos críticos"
    ttb = s.shapes.add_textbox(Emu(RP_L), Emu(130000), Emu(RP_W), Emu(540000))
    ttb.text_frame.word_wrap = False
    p = ttb.text_frame.paragraphs[0]; p.text = "Hallazgos críticos"
    r = p.runs[0]; r.font.size = Pt(34); r.font.bold = True; r.font.color.rgb = NAVY

    # 3 hallazgos
    BULL_D  = 175000
    ITEM_T  = 760000
    ITEM_S  = 1380000
    TXT_L   = RP_L + BULL_D + 160000
    TXT_W   = RP_W - BULL_D - 160000

    for i, h in enumerate(hallazgos[:3]):
        it = ITEM_T + i * ITEM_S

        # Bala azul rellena
        bull = s.shapes.add_shape(9, Emu(RP_L), Emu(it + 20000), Emu(BULL_D), Emu(BULL_D))
        bull.fill.solid(); bull.fill.fore_color.rgb = BLUE; bull.line.fill.background()

        # Título del hallazgo (negrita)
        htb = s.shapes.add_textbox(Emu(TXT_L), Emu(it), Emu(TXT_W), Emu(220000))
        p = htb.text_frame.paragraphs[0]; p.text = h.get("title", "")[:60]
        r = p.runs[0]; r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = DARK

        # Detalle (multi-línea)
        detail_raw   = h.get("detail", "").replace("\\n", "\n")
        detail_lines = [l.strip() for l in detail_raw.split("\n") if l.strip()]
        det_tb = s.shapes.add_textbox(Emu(TXT_L), Emu(it + 225000), Emu(TXT_W), Emu(ITEM_S - 260000))
        det_tf = det_tb.text_frame; det_tf.word_wrap = True
        for li, line in enumerate(detail_lines[:3]):
            if li == 0:
                p = det_tf.paragraphs[0]
            else:
                p = det_tf.add_paragraph()
            p.text = line[:160]
            if p.runs:
                r = p.runs[0]; r.font.size = Pt(11); r.font.color.rgb = DARK

    # Texto de prioridad al fondo
    ptb = s.shapes.add_textbox(Emu(RP_L), Emu(SH - 460000), Emu(RP_W), Emu(210000))
    p = ptb.text_frame.paragraphs[0]
    p.text = f"PRIORIDAD: {prioridad[:55]}"
    r = p.runs[0]; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = BLUE


# ─────────────────────────────────────────────────────────────
# HELPER — diseño de lista numerada (slides 7 y 8)
# ─────────────────────────────────────────────────────────────

def build_list_slide(s, title: str, subtitle: str, items: list):
    """Dibuja slides 7/8: fondo blanco, título, línea, subtítulo, 4 ítems con círculo."""
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    NAVY  = RGBColor(0x0F, 0x1B, 0x3D)
    LGRAY = RGBColor(0xBB, 0xBB, 0xBB)
    DARK  = RGBColor(0x22, 0x22, 0x22)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    SW, SH = 9144000, 5143500
    ML = 500000  # margen izquierdo

    # Eliminar todos los shapes del template para evitar solapamiento
    sp_tree = s.shapes._spTree
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    # Fondo blanco
    bg = s.shapes.add_shape(1, Emu(0), Emu(0), Emu(SW), Emu(SH))
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE
    bg.line.fill.background()

    # Título
    tb = s.shapes.add_textbox(Emu(ML), Emu(260000), Emu(SW - 2 * ML), Emu(570000))
    tb.text_frame.word_wrap = False
    p = tb.text_frame.paragraphs[0]
    p.text = title
    r = p.runs[0]
    r.font.bold = True
    r.font.size = Pt(38)
    r.font.color.rgb = NAVY

    # Línea separadora corta
    sep = s.shapes.add_shape(1, Emu(ML), Emu(865000), Emu(700000), Emu(16000))
    sep.fill.solid()
    sep.fill.fore_color.rgb = LGRAY
    sep.line.fill.background()

    # Subtítulo
    sb = s.shapes.add_textbox(Emu(ML), Emu(940000), Emu(SW - 2 * ML), Emu(230000))
    p = sb.text_frame.paragraphs[0]
    p.text = subtitle
    r = p.runs[0]
    r.font.size = Pt(12)
    r.font.color.rgb = DARK

    # Ítems numerados
    CD = 370000    # diámetro del círculo
    IT = 1100000   # top del primer ítem
    IS = 1010000   # paso vertical entre ítems
    TL = ML + CD + 130000  # left del texto

    for i, text in enumerate(items[:4]):
        top = IT + i * IS

        # Círculo con número
        circ = s.shapes.add_shape(9, Emu(ML), Emu(top), Emu(CD), Emu(CD))
        circ.fill.background()
        circ.line.color.rgb = NAVY
        circ.line.width = Pt(1.5)
        tf = circ.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = str(i + 1)
        r = p.runs[0]
        r.font.size = Pt(12)
        r.font.color.rgb = NAVY
        r.font.bold = True

        # Texto del ítem
        txb = s.shapes.add_textbox(Emu(TL), Emu(top), Emu(SW - TL - ML), Emu(CD + 260000))
        tf2 = txb.text_frame
        tf2.word_wrap = True
        tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf2.paragraphs[0]
        p.text = text[:280]
        r = p.runs[0]
        r.font.size = Pt(12)
        r.font.color.rgb = DARK


# ─────────────────────────────────────────────────────────────
# HELPER — slide próximos pasos (slide 10)
# ─────────────────────────────────────────────────────────────

def _build_steps_slide(s, steps: list):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    NAVY   = RGBColor(0x0F, 0x1B, 0x3D)
    ORANGE = RGBColor(0xFF, 0x66, 0x00)
    WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    LGRAY  = RGBColor(0xCC, 0xD6, 0xE8)
    SW, SH = 9144000, 5143500

    sp_tree = s.shapes._spTree
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    # Fondo navy
    bg = s.shapes.add_shape(1, Emu(0), Emu(0), Emu(SW), Emu(SH))
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()

    # Título
    ML = 500000
    tb = s.shapes.add_textbox(Emu(ML), Emu(100000), Emu(SW - 2*ML), Emu(500000))
    tb.text_frame.word_wrap = False
    p = tb.text_frame.paragraphs[0]; p.text = "Próximos pasos"
    r = p.runs[0]; r.font.size = Pt(40); r.font.bold = True; r.font.color.rgb = WHITE

    NUM_W  = 260000
    ROW_T  = 720000
    ROW_H  = 980000

    for i, step in enumerate(steps[:4]):
        top   = ROW_T + i * ROW_H
        title = step.get("title", "")[:70]
        desc  = step.get("desc", "")[:180]

        # Número naranja
        ntb = s.shapes.add_textbox(Emu(ML), Emu(top), Emu(NUM_W), Emu(ROW_H))
        ntb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = ntb.text_frame.paragraphs[0]; p.text = str(i + 1)
        r = p.runs[0]; r.font.size = Pt(36); r.font.bold = True; r.font.color.rgb = ORANGE

        TEXT_L = ML + NUM_W + 120000
        TEXT_W = SW - TEXT_L - ML

        # Título blanco
        ttb = s.shapes.add_textbox(Emu(TEXT_L), Emu(top + 80000), Emu(TEXT_W), Emu(280000))
        ttb.text_frame.word_wrap = True
        p = ttb.text_frame.paragraphs[0]; p.text = title
        r = p.runs[0]; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = WHITE

        # Descripción gris claro itálica
        dtb = s.shapes.add_textbox(Emu(TEXT_L), Emu(top + 360000), Emu(TEXT_W), Emu(540000))
        dtb.text_frame.word_wrap = True
        p = dtb.text_frame.paragraphs[0]; p.text = desc
        r = p.runs[0]; r.font.size = Pt(13); r.font.italic = True; r.font.color.rgb = LGRAY


# ─────────────────────────────────────────────────────────────
# HELPER — slide recomendaciones (slide 9)
# ─────────────────────────────────────────────────────────────

def build_recommendations_slide(s, recommendations: list):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    NAVY   = RGBColor(0x0F, 0x1B, 0x3D)
    ORANGE = RGBColor(0xFF, 0x66, 0x00)
    LGRAY  = RGBColor(0xF5, 0xF5, 0xF5)
    DGRAY  = RGBColor(0x88, 0x88, 0x88)
    DARK   = RGBColor(0x22, 0x22, 0x22)
    WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    SW, SH = 9144000, 5143500

    # Limpiar template
    sp_tree = s.shapes._spTree
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    # Fondo blanco
    bg = s.shapes.add_shape(1, Emu(0), Emu(0), Emu(SW), Emu(SH))
    bg.fill.solid(); bg.fill.fore_color.rgb = WHITE; bg.line.fill.background()

    # Título
    ML = 400000
    tb = s.shapes.add_textbox(Emu(ML), Emu(120000), Emu(SW - 2*ML), Emu(500000))
    tb.text_frame.word_wrap = False
    p = tb.text_frame.paragraphs[0]; p.text = "Top 5 recomendaciones"
    r = p.runs[0]; r.font.size = Pt(36); r.font.bold = True; r.font.color.rgb = NAVY

    # Filas
    NUM_W  = 300000
    ROW_H  = 820000
    ROW_T  = 700000
    ROW_GAP = 20000
    STAR_W = 500000

    for i, rec in enumerate(recommendations[:5]):
        top    = ROW_T + i * (ROW_H + ROW_GAP)
        title  = rec.get("title", "")[:70]
        desc   = rec.get("desc", "")[:160]
        stars  = "⭐" * max(1, min(5, rec.get("priority", 3)))

        # Fondo fila alternado
        row_bg = s.shapes.add_shape(1, Emu(ML), Emu(top), Emu(SW - 2*ML), Emu(ROW_H))
        row_bg.fill.solid(); row_bg.fill.fore_color.rgb = LGRAY; row_bg.line.fill.background()

        # Número
        ntb = s.shapes.add_textbox(Emu(ML + 60000), Emu(top), Emu(NUM_W), Emu(ROW_H))
        ntb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = ntb.text_frame.paragraphs[0]; p.text = str(i + 1)
        r = p.runs[0]; r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = ORANGE

        # Título + descripción
        TEXT_L = ML + NUM_W + 120000
        TEXT_W = SW - TEXT_L - ML - STAR_W - 80000

        ttb = s.shapes.add_textbox(Emu(TEXT_L), Emu(top + 120000), Emu(TEXT_W), Emu(260000))
        ttb.text_frame.word_wrap = True
        p = ttb.text_frame.paragraphs[0]; p.text = title
        r = p.runs[0]; r.font.size = Pt(13); r.font.bold = True; r.font.color.rgb = DARK

        dtb = s.shapes.add_textbox(Emu(TEXT_L), Emu(top + 380000), Emu(TEXT_W), Emu(360000))
        dtb.text_frame.word_wrap = True
        p = dtb.text_frame.paragraphs[0]; p.text = desc
        r = p.runs[0]; r.font.size = Pt(11); r.font.color.rgb = DGRAY

        # Estrellas
        stb = s.shapes.add_textbox(Emu(SW - ML - STAR_W), Emu(top), Emu(STAR_W), Emu(ROW_H))
        stb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = stb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT; p.text = stars
        r = p.runs[0]; r.font.size = Pt(16)


# ─────────────────────────────────────────────────────────────
# GENERADOR PPTX
# ─────────────────────────────────────────────────────────────

def generate_pptx(empresa: str, fecha: date, logo_bytes, d: dict, fourth_model: str = "Perplexity") -> bytes:
    prs      = Presentation(TEMPLATE_PATH)
    date_str = f"{MONTHS_ES[fecha.month]} {fecha.year}"
    gpt      = d["chatgpt_score"]
    gem      = d["gemini_score"]
    cla      = d["claude_score"]
    per      = d["perplexity_score"]
    score    = d["score_global"]
    avg      = round((gpt + gem + cla + per) / 4)

    # Slide 1 — Portada
    s = prs.slides[0]
    set_run(s, "Text 2", empresa)
    set_run(s, "Text 4", "")

    # Recuadro blanco en la franja inferior
    from pptx.dml.color import RGBColor
    WHITE_BOX_H = 1000000
    white_box_top = 5143500 - WHITE_BOX_H
    wb = s.shapes.add_shape(1, Emu(0), Emu(white_box_top), Emu(9144000), Emu(WHITE_BOX_H))
    wb.fill.solid()
    wb.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    wb.line.fill.background()
    # Mover el recuadro al fondo (detrás de todos los shapes)
    sp_elem  = wb._element
    sp_tree  = s.shapes._spTree
    sp_tree.remove(sp_elem)
    sp_tree.insert(2, sp_elem)  # índice 2 = justo después de nvGrpSpPr y grpSpPr

    # Logo La Fábrica del SEO dentro del recuadro blanco
    fabrica_logo = os.path.join(BASE_DIR, "logo_fabrica.png")
    if os.path.exists(fabrica_logo):
        from PIL import Image as PILImage
        with open(fabrica_logo, "rb") as f:
            fl_bytes = f.read()
        img = PILImage.open(io.BytesIO(fl_bytes))
        img_w, img_h = img.size
        logo_h_emu = int(WHITE_BOX_H * 0.60)
        logo_w_emu = int(logo_h_emu * img_w / img_h)
        logo_left = 9144000 - logo_w_emu - 350000
        logo_top  = white_box_top + (WHITE_BOX_H - logo_h_emu) // 2
        s.shapes.add_picture(fabrica_logo, Emu(logo_left), Emu(logo_top), Emu(logo_w_emu), Emu(logo_h_emu))

    # Slide 2 — Resumen ejecutivo
    s     = prs.slides[1]
    nivel = score_level(score)
    set_run(s, "Text 1", f"Score: {score}/100")
    set_body(s, "Text 2", [
        f"{empresa} obtiene una puntuación de {score}/100 en AI Visibility, "
        f"lo que indica un nivel {nivel} de presencia en búsquedas de IA. "
        f"Los principales modelos de lenguaje (ChatGPT, Gemini, Claude, Perplexity) "
        f"{'reconocen' if score >= 60 else 'no reconocen'} la empresa como referente en su sector.",
        "",
        d["resumen"],
    ])

    # Slide 3 — Visibilidad por modelo (barras proporcionales al score, full height del track)
    s = prs.slides[2]
    for bar_name, track_name, score_val in [
        ("Shape 2",  "Shape 1",  gpt),
        ("Shape 7",  "Shape 6",  gem),
        ("Shape 12", "Shape 11", cla),
        ("Shape 17", "Shape 16", per),
    ]:
        track = find_shape(s, track_name)
        bar   = find_shape(s, bar_name)
        if track and bar:
            bar.top    = track.top
            bar.height = track.height
            bar.width  = Emu(max(1, score_val) * EMU_PER_SCORE_PT)
    set_run(s, "Text 4",  f"{gpt}/100"); set_run(s, "Text 5",  model_label(gpt, "ChatGPT"))
    set_run(s, "Text 9",  f"{gem}/100"); set_run(s, "Text 10", model_label(gem, "Gemini"))
    set_run(s, "Text 14", f"{cla}/100"); set_run(s, "Text 15", model_label(cla, "Claude"))
    set_run(s, "Text 18", fourth_model)
    set_run(s, "Text 19", f"{per}/100"); set_run(s, "Text 20", model_label(per, fourth_model))
    set_run(s, "Text 21", f"Promedio: {avg}/100")

    # Slide 4 — Contexto competitivo (cajas más grandes, texto justificado)
    s = prs.slides[3]
    NEW_BG_H   = 731520   # antes 502920
    NEW_TEXT_H = 457200   # antes 228600
    TEXT_OFFSET = 73152   # desplazamiento título dentro de la caja
    row_tops = [1097280, 1920240, 2743200, 3566160]
    bg_names   = ["Shape 1", "Shape 5", "Shape 9",  "Shape 13"]
    text_groups = [
        ("Text 2",  "Text 3",  "Text 4"),
        ("Text 6",  "Text 7",  "Text 8"),
        ("Text 10", "Text 11", "Text 12"),
        ("Text 14", "Text 15", "Text 16"),
    ]
    for row_i, (bg_name, texts, row_top) in enumerate(zip(bg_names, text_groups, row_tops)):
        bg = find_shape(s, bg_name)
        if bg:
            bg.top    = Emu(row_top)
            bg.height = Emu(NEW_BG_H)
        for t_name in texts:
            sh = find_shape(s, t_name)
            if sh:
                sh.top    = Emu(row_top + TEXT_OFFSET)
                sh.height = Emu(NEW_TEXT_H)

    comps = d.get("competitors", [])
    for i, (tn, sn, dn) in enumerate([("Text 2","Text 3","Text 4"),("Text 6","Text 7","Text 8"),("Text 10","Text 11","Text 12")]):
        if i < len(comps):
            c = comps[i]
            set_run(s, tn, c.get("name", ""))
            set_run(s, sn, "⭐" * max(1, min(5, c.get("stars", 3))))
            set_run(s, dn, c.get("desc", "")[:120])
    set_run(s, "Text 14", empresa)
    set_run(s, "Text 15", score_stars(score))
    set_run(s, "Text 16", d["competitive_desc"][:120])

    # Slide 5 — Análisis ChatGPT (diseño panel score + hallazgos)
    build_analysis_slide(
        prs.slides[4],
        score       = gpt,
        model_name  = "ChatGPT",
        hallazgos   = d.get("chatgpt_hallazgos", []),
        diagnosticos= d.get("chatgpt_diagnosticos", []),
        prioridad   = d.get("chatgpt_prioridad", ""),
    )

    # Slide 6 — Análisis Gemini (diseño panel score + hallazgos)
    build_analysis_slide(
        prs.slides[5],
        score       = gem,
        model_name  = "Gemini",
        hallazgos   = d.get("gemini_hallazgos", []),
        diagnosticos= d.get("gemini_diagnosticos", []),
        prioridad   = d.get("gemini_prioridad", ""),
    )

    # Slide 7 — Fortalezas (diseño lista numerada)
    build_list_slide(
        prs.slides[6],
        title    = "Fortalezas actuales",
        subtitle = "Lo que está funcionando:",
        items    = d.get("strengths", []),
    )

    # Slide 8 — Oportunidades (diseño lista numerada)
    build_list_slide(
        prs.slides[7],
        title    = "Oportunidades de mejora",
        subtitle = "Áreas con potencial de crecimiento:",
        items    = d.get("opportunities", []),
    )

    # Slide 9 — Recomendaciones (diseño dinámico)
    build_recommendations_slide(prs.slides[8], d.get("recommendations", []))

    # Slide 10 — Próximos pasos (diseño dinámico)
    _build_steps_slide(prs.slides[9], d.get("steps", []))

    # Slide 11 — Logo
    if logo_bytes:
        replace_logo(prs.slides[10], "Picture 2", logo_bytes)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# ESCÁNER GEO — FUNCIONES
# ─────────────────────────────────────────────────────────────

def geo_detect_brand(domain: str, observaciones: str = "") -> dict:
    obs_hint = f"\nContexto adicional: {observaciones.strip()}" if observaciones.strip() else ""
    client = Anthropic(api_key=ANTHROPIC_KEY)
    r = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=200,
        messages=[{"role": "user", "content": (
            f"Dado el dominio: {domain}{obs_hint}\n"
            "Devuelve SOLO este JSON (sin texto extra):\n"
            '{"brand": "<nombre comercial>", "sector": "<sector en 3-4 palabras>", "pais": "<país en español>"}\n'
            'Ejemplo: {"brand": "La Fabrica del SEO", "sector": "agencia SEO", "pais": "Espana"}'
        )}]
    )
    raw = r.content[0].text.strip()
    if "```" in raw:
        raw = raw.split("```")[1].split("```")[0].strip()
        if raw.startswith("json"):
            raw = raw[4:].strip()
    return json.loads(raw)


SECTION_KEYS = ["competitors", "reputation", "strengths", "opportunities", "search_prompts"]

def geo_section_prompts(brand: str, sector: str, pais: str) -> dict:
    return {
        "competitors": (
            f"¿Qué empresas de {sector} en {pais} son las más conocidas y recomendadas? "
            f"Lista las principales con una breve descripción de cada una."
        ),
        "reputation": (
            f"¿Conoces la empresa '{brand}' del sector {sector} en {pais}? "
            f"¿Qué sabes de ella, cómo la describirías y qué reputación tiene?"
        ),
        "strengths": (
            f"¿Qué aspectos positivos o fortalezas destacarías de '{brand}' "
            f"en el sector {sector} en {pais}? Si no la conoces, indícalo."
        ),
        "opportunities": (
            f"¿En qué aspectos podría mejorar '{brand}' comparado con otras empresas "
            f"de {sector} en {pais}? Si no la conoces, indícalo."
        ),
        "search_prompts": (
            f"¿Qué preguntas o búsquedas haría alguien que necesita contratar "
            f"servicios de {sector} en {pais}? Dame 5 ejemplos realistas."
        ),
    }


def _query_safe(fn, prompt: str) -> str:
    try:
        return fn(prompt)
    except Exception as e:
        return f"[ERROR: {e}]"


def geo_query_claude(prompt: str) -> str:
    client = Anthropic(api_key=ANTHROPIC_KEY)
    r = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=500,
        messages=[{"role": "user", "content": prompt}]
    )
    return r.content[0].text


def geo_query_openai(prompt: str) -> str:
    from openai import OpenAI
    client = OpenAI(api_key=OPENAI_KEY)
    r = client.chat.completions.create(
        model="gpt-4o-mini",
        max_tokens=500,
        messages=[{"role": "user", "content": prompt}]
    )
    return r.choices[0].message.content


def geo_query_gemini(prompt: str) -> str:
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"gemini-flash-latest:generateContent?key={GEMINI_KEY}"
    )
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": 500},
    }
    r = requests.post(url, json=body, timeout=30)
    r.raise_for_status()
    return r.json()["candidates"][0]["content"]["parts"][0]["text"]


def geo_query_groq(prompt: str) -> str:
    from openai import OpenAI
    client = OpenAI(api_key=GROQ_KEY, base_url="https://api.groq.com/openai/v1")
    r = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        max_tokens=500,
        messages=[{"role": "user", "content": prompt}]
    )
    return r.choices[0].message.content


def geo_mentions(brand: str, text: str) -> bool:
    """
    Devuelve True SOLO si la marca específica aparece mencionada en el texto.
    Requiere que TODAS las palabras significativas (>2 chars) del nombre de marca
    aparezcan en el texto — evita falsos positivos por palabras genéricas del sector.
    """
    if "[ERROR" in text:
        return False
    bl = text.lower()
    # Primero intenta match exacto del nombre completo
    if brand.lower() in bl:
        return True
    # Si no, exige que TODAS las palabras significativas aparezcan
    significant = [w for w in brand.lower().split() if len(w) > 2]
    if len(significant) >= 2:
        return all(w in bl for w in significant)
    # Marca de una sola palabra significativa: match directo ya cubierto arriba
    return False


# Señales de desconocimiento — se buscan en TODO el texto de la respuesta
_STRONG_DENIAL = [
    "no conozco", "no la conozco", "no estoy familiarizado",
    "desconozco esta empresa", "no tengo conocimiento de esta",
    "i don't know", "i do not know", "not familiar with",
    "no tengo información sobre esta empresa",
    "no tengo información acerca de esta empresa",
    "no tengo datos sobre esta empresa",
    "no tengo información específica sobre esta empresa",
    "no dispongo de información",
    "no tengo datos suficientes sobre esta marca",
    "no cuento con información sobre esta empresa",
    "no puedo proporcionar información específica sobre esta empresa",
    "sin información verificada sobre esta marca",
    "no he encontrado información sobre esta empresa",
    "no tengo información verificada sobre",
]


def _model_knows(text: str, min_chars: int = 150) -> bool:
    """
    Devuelve True si el modelo dio una respuesta sustancial sobre la marca.
    Criterios:
    1. Respuesta de al menos `min_chars` caracteres
    2. Sin negación de conocimiento en TODO el texto (no solo la apertura)
    """
    if not text or "[ERROR" in text:
        return False
    text = text.strip()
    if len(text) < min_chars:
        return False
    full = text.lower()
    return not any(sig in full for sig in _STRONG_DENIAL)


def geo_score_from_sections(brand: str, sections: dict) -> int:
    """
    Score de visibilidad GEO por modelo (0-100):

    20 pts — competitors:   marca mencionada espontáneamente (pregunta genérica sin nombrarla)
    50 pts — reputation:    modelo demuestra conocimiento real de la marca
    15 pts — strengths:     modelo da fortalezas específicas (no dice "no la conozco")
    15 pts — opportunities: modelo da oportunidades específicas (no dice "no la conozco")
    """
    if not sections:
        return 0

    score = 0

    # 20 pts — mención espontánea en competitors
    competitors_resp = sections.get("competitors", "")
    if competitors_resp and "[ERROR" not in competitors_resp:
        if geo_mentions(brand, competitors_resp):
            score += 20

    # 50 pts — conocimiento real en reputation
    if _model_knows(sections.get("reputation", "")):
        score += 50

    # 15 pts — fortalezas específicas
    if _model_knows(sections.get("strengths", "")):
        score += 15

    # 15 pts — oportunidades específicas
    if _model_knows(sections.get("opportunities", "")):
        score += 15

    return min(score, 100)


def geo_analyze_results(brand: str, sector: str, pais: str, scores: dict, all_responses: dict, observaciones: str = "", ahrefs_img: bytes = None) -> dict:
    resps_text = ""
    for model, sections in all_responses.items():
        resps_text += f"\n\n=== {model} ===\n"
        if isinstance(sections, dict):
            for section, resp in sections.items():
                resps_text += f"[{section.upper()}] {resp[:250]}\n"
        else:
            for i, r in enumerate(sections, 1):
                resps_text += f"[P{i}] {r[:200]}\n"

    obs_block = f"\nOBSERVACIONES DEL CLIENTE (tenlas muy en cuenta): {observaciones.strip()}\n" if observaciones.strip() else ""
    ahrefs_block = ""
    if ahrefs_img:
        ahrefs_block = "\nADEMÁS se adjunta una captura de pantalla de Ahrefs AI Citations para este dominio. Extrae de ella: qué plataformas de IA citan el dominio, cuántas veces, qué páginas son las más citadas, y cualquier otro dato relevante visible. Integra esta información en el análisis, especialmente en brand_reputation, ai_search_insights y recommendations.\n"

    client = Anthropic(api_key=ANTHROPIC_KEY)
    prompt = f"""Eres un experto en GEO (Generative Engine Optimization) redactando un informe profesional para el cliente final.

REGLA ABSOLUTA — NUNCA menciones en ningún campo del JSON:
- Nombres de herramientas o APIs: Ahrefs, OpenAI, Anthropic, Groq, Llama, Claude API, ChatGPT API, Gemini API
- Errores técnicos, fallos de módulos, límites de tokens, problemas de conexión
- El hecho de que los datos vienen de un análisis automatizado o de capturas de pantalla
- Frases como "el análisis indica", "la herramienta confirma", "según los datos de X"

Escribe siempre en primera persona experta: "La marca no aparece en las respuestas de los modelos de IA", "Los modelos de IA no reconocen la marca", "Se recomienda crear contenido que los modelos puedan referenciar". El informe debe sonar como si lo hubiera redactado un consultor GEO senior, no una máquina.
{obs_block}{ahrefs_block}
Analiza los resultados de visibilidad de "{brand}" ({sector}, {pais}) en modelos de IA.

Scores obtenidos:
- Claude: {scores.get('Claude', 0)}/100
- ChatGPT: {scores.get('ChatGPT', 0)}/100
- Gemini: {scores.get('Gemini', 0)}/100
- Groq/Llama: {scores.get('Groq', 0)}/100

Respuestas de los modelos cuando se les preguntó qué empresas de {sector} recomiendan en {pais}:
{resps_text[:4000]}

INSTRUCCIONES:

TOP COMPETITORS: Lee las respuestas e identifica qué empresas mencionan los modelos. Esas son los competidores de "{brand}" en IA. Toma los 3 más mencionados, excluye a "{brand}". Para cada uno: nombre exacto, qué hace según los modelos, estrellas (4-5 si muy mencionado, 2-3 si poco).

BRAND REPUTATION: Describe cómo perciben los modelos a "{brand}" basándote EXCLUSIVAMENTE en los scores numéricos ya calculados (no en el texto literal de las respuestas, que puede ser parcial o ambiguo). El score es un indicador compuesto que refleja la visibilidad real de la marca en cada modelo. Usa esta escala OBLIGATORIA para redactar el resumen y elegir el tono:
- Score 80-100 → tono "Positiva": el modelo conoce bien la marca, la menciona, describe atributos concretos.
- Score 60-79  → tono "Neutra": el modelo reconoce la marca con información moderada, sin diferenciación clara.
- Score 40-59  → tono "Neutra": el modelo tiene información muy limitada, presencia incipiente.
- Score 20-39  → tono "Negativa": el modelo apenas reconoce la marca, menciones mínimas o confusas.
- Score 0-19   → tono "Ausente": el modelo no reconoce la marca en absoluto.
NUNCA uses frases como "admite no disponer de información" o "no reconoce" para modelos con score ≥60. El modelo con mayor score SIEMPRE debe describirse como el que mejor conoce la marca.

KEY STRENGTHS: Fortalezas reales de "{brand}" que debería capitalizar para mejorar su GEO (basadas en su sector y lo que valoran los modelos de empresas similares).

IMPROVEMENT OPPORTUNITIES: Oportunidades concretas basadas en lo que SÍ dicen los modelos de los competidores pero NO de "{brand}".

SAMPLE SEARCH PROMPTS: Ejemplos de búsquedas que los usuarios reales hacen en IA y donde "{brand}" debería aparecer pero no aparece. Hazlos realistas, como preguntaría un cliente potencial.

AI SEARCH INSIGHTS: Para cada modelo (ChatGPT, Gemini, Claude, Groq/Llama) un insight específico sobre cómo trata a "{brand}" y qué debería hacer para mejorar en ese modelo concreto. El campo "visibility" DEBE coincidir con el score: ≥80=Alta, 60-79=Media, 40-59=Media, 20-39=Baja, <20=Nula. El insight debe ser coherente con ese nivel de visibilidad.

Devuelve SOLO este JSON:
{{
  "resumen": "<2-3 frases sobre visibilidad global>",
  "competitive_desc": "<frase corta sobre posicion vs competidores>",
  "brand_reputation": {{
    "tone": "<Ausente|Negativa|Neutra|Positiva>",
    "summary": "<2-3 frases sobre cómo describen los modelos a la marca, o por qué no la mencionan>",
    "attributes": ["<atributo 1>", "<atributo 2>", "<atributo 3>"]
  }},
  "competitors": [
    {{"name": "<nombre>", "stars": <1-5>, "desc": "<que hace segun modelos, max 100 chars>"}},
    {{"name": "<nombre>", "stars": <1-5>, "desc": "<que hace segun modelos, max 100 chars>"}},
    {{"name": "<nombre>", "stars": <1-5>, "desc": "<que hace segun modelos, max 100 chars>"}}
  ],
  "strengths": ["<fortaleza 1>", "<fortaleza 2>", "<fortaleza 3>", "<fortaleza 4>"],
  "opportunities": ["<oportunidad 1>", "<oportunidad 2>", "<oportunidad 3>", "<oportunidad 4>"],
  "sample_prompts": [
    "<pregunta realista que haría un cliente potencial en ChatGPT/Gemini>",
    "<pregunta realista 2>",
    "<pregunta realista 3>",
    "<pregunta realista 4>",
    "<pregunta realista 5>"
  ],
  "ai_search_insights": [
    {{"model": "ChatGPT", "visibility": "<Nula|Baja|Media|Alta>", "insight": "<qué dice o no dice ChatGPT sobre la marca y por qué>"}},
    {{"model": "Gemini",  "visibility": "<Nula|Baja|Media|Alta>", "insight": "<insight específico de Gemini>"}},
    {{"model": "Claude",  "visibility": "<Nula|Baja|Media|Alta>", "insight": "<insight específico de Claude>"}},
    {{"model": "Groq/Llama", "visibility": "<Nula|Baja|Media|Alta>", "insight": "<insight específico de Llama>"}}
  ],
  "recommendations": [
    {{"title": "<accion 1, max 50 chars>", "desc": "<1 frase, max 90 chars>", "priority": 3}},
    {{"title": "<accion 2, max 50 chars>", "desc": "<1 frase, max 90 chars>", "priority": 3}},
    {{"title": "<accion 3, max 50 chars>", "desc": "<1 frase, max 90 chars>", "priority": 3}},
    {{"title": "<accion 4, max 50 chars>", "desc": "<1 frase, max 90 chars>", "priority": 2}},
    {{"title": "<accion 5, max 50 chars>", "desc": "<1 frase, max 90 chars>", "priority": 2}}
  ],
  "chatgpt_hallazgos": [
    {{"title": "<hallazgo 1, max 40 chars>", "detail": "<2 lineas con \\n, max 160 chars>"}},
    {{"title": "<hallazgo 2, max 40 chars>", "detail": "<2 lineas con \\n, max 160 chars>"}},
    {{"title": "<hallazgo 3, max 40 chars>", "detail": "<2 lineas con \\n, max 160 chars>"}}
  ],
  "chatgpt_diagnosticos": [
    {{"label": "Visibilidad ChatGPT", "value": "<Nula|Baja|Media|Alta>"}},
    {{"label": "<diagnostico 2>", "value": "<valor>"}},
    {{"label": "<diagnostico 3>", "value": "<valor>"}}
  ],
  "chatgpt_prioridad": "<ACCION PRIORITARIA EN MAYUSCULAS, max 55 chars>",
  "gemini_hallazgos": [
    {{"title": "<hallazgo 1, max 40 chars>", "detail": "<2 lineas con \\n, max 160 chars>"}},
    {{"title": "<hallazgo 2, max 40 chars>", "detail": "<2 lineas con \\n, max 160 chars>"}},
    {{"title": "<hallazgo 3, max 40 chars>", "detail": "<2 lineas con \\n, max 160 chars>"}}
  ],
  "gemini_diagnosticos": [
    {{"label": "Visibilidad Gemini", "value": "<Nula|Baja|Media|Alta>"}},
    {{"label": "<diagnostico 2>", "value": "<valor>"}},
    {{"label": "<diagnostico 3>", "value": "<valor>"}}
  ],
  "gemini_prioridad": "<ACCION PRIORITARIA EN MAYUSCULAS, max 55 chars>",
  "steps": [
    {{"title": "Revisar hallazgos", "desc": "<1 frase, max 90 chars>"}},
    {{"title": "Priorizar iniciativas", "desc": "<1 frase, max 90 chars>"}},
    {{"title": "Implementar Quick Wins", "desc": "<1 frase, max 90 chars>"}},
    {{"title": "Monitorear progreso", "desc": "<1 frase, max 90 chars>"}}
  ]
}}"""

    content = []
    if ahrefs_img:
        media_type = "image/png" if ahrefs_img[:4] == b"\x89PNG" else "image/jpeg"
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": base64.standard_b64encode(ahrefs_img).decode()}
        })
    content.append({"type": "text", "text": prompt})

    r = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=6000,
        messages=[{"role": "user", "content": content}]
    )
    raw = r.content[0].text.strip()
    if "```json" in raw:
        raw = raw.split("```json")[1].split("```")[0].strip()
    elif "```" in raw:
        raw = raw.split("```")[1].split("```")[0].strip()

    # Si el JSON viene truncado, intentar cerrarlo
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Truncar al último campo completo y cerrar el JSON
        last_brace = raw.rfind('},')
        if last_brace == -1:
            last_brace = raw.rfind('}')
        if last_brace != -1:
            raw = raw[:last_brace + 1] + "\n]}\n" if '"recommendations"' in raw[last_brace:] else raw[:last_brace + 1] + "\n}"
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            # Devolver estructura mínima para no romper la app
            return {
                "resumen": "Análisis generado parcialmente por límite de tokens.",
                "competitive_desc": "", "brand_reputation": {"tone": "Neutra", "summary": "", "attributes": []},
                "competitors": [], "strengths": [], "opportunities": [], "sample_prompts": [],
                "ai_search_insights": [], "recommendations": [], "steps": [],
                "chatgpt_hallazgos": [], "chatgpt_diagnosticos": [], "chatgpt_prioridad": "",
                "gemini_hallazgos": [], "gemini_diagnosticos": [], "gemini_prioridad": "",
            }


# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="Generador Auditoría GEO", page_icon="🔍", layout="wide")

st.title("Generador de Auditorías GEO")

tab1, tab2 = st.tabs(["📊 Desde Screenshot (LLMs Pulse)", "🔍 Escáner GEO por Dominio"])

# ══════════════════════════════════════════════════════════════
# TAB 1 — desde screenshot
# ══════════════════════════════════════════════════════════════
with tab1:
    st.caption("Sube el screenshot de LLMs Pulse → Claude analiza → descarga el PPTX.")

    with st.sidebar:
        st.header("Datos de la auditoría")

        empresa   = st.text_input("Nombre de la empresa *", placeholder="Ej: Acme Corp")
        fecha     = st.date_input("Fecha de auditoría", value=date.today())
        logo_file = st.file_uploader("Logo de la empresa (opcional)", type=["jpg", "jpeg", "png"])
        if logo_file:
            st.image(logo_file, caption="Logo cargado", use_container_width=True)

        st.divider()
        pulse_file = st.file_uploader(
            "Screenshot de LLMs Pulse *",
            type=["jpg", "jpeg", "png"],
            help="Captura de pantalla con los scores de visibilidad por modelo",
        )

        st.divider()
        report_url = st.text_input(
            "URL del informe LLMs Pulse (opcional)",
            placeholder="https://llmpulse.ai/ai-visibility-report/...",
            help="Si pegas la URL del informe, los competidores se extraen automáticamente.",
        )

        st.divider()
        observaciones_t1 = st.text_area(
            "Observaciones del cliente (opcional)",
            placeholder="Ej: El cliente quiere potenciar su presencia en ChatGPT. Acaba de lanzar un nuevo servicio de auditoría técnica. Compite principalmente con X e Y en Madrid...",
            height=150,
            help="Claude usará estas notas para personalizar el análisis y las recomendaciones.",
        )

    col_img, col_action = st.columns([3, 2], gap="large")

    with col_img:
        if pulse_file:
            st.subheader("Vista previa LLMs Pulse")
            pulse_file.seek(0)
            st.image(pulse_file, use_container_width=True)
        else:
            st.info("Sube el screenshot de LLMs Pulse en el sidebar para previsualizarlo aquí.")

    with col_action:
        st.subheader("Generar auditoría")

        if empresa:
            st.markdown(f"**Empresa:** {empresa}")
        st.markdown(f"**Fecha:** {MONTHS_ES[fecha.month]} {fecha.year}")
        if logo_file:
            st.markdown("**Logo:** cargado")
        if pulse_file:
            st.markdown("**LLMs Pulse:** imagen cargada")

        st.divider()

        if st.button("Analizar y Generar PPTX", type="primary", use_container_width=True):
            errors = []
            if not empresa.strip():
                errors.append("El nombre de la empresa es obligatorio.")
            if not pulse_file:
                errors.append("El screenshot de LLMs Pulse es obligatorio.")
            if not ANTHROPIC_KEY:
                errors.append("ANTHROPIC_API_KEY no configurada.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                competitors_data     = []
                recommendations_data = []
                opportunities_data   = []
                if report_url.strip():
                    with st.spinner("Leyendo datos desde el informe LLMs Pulse..."):
                        report_data = fetch_report_data_from_url(report_url.strip())
                        competitors_data     = report_data.get("competitors", [])
                        recommendations_data = report_data.get("recommendations", [])
                        opportunities_data   = report_data.get("opportunities", [])
                        msgs = []
                        if competitors_data:
                            msgs.append(f"Competidores: {', '.join(c['name'] for c in competitors_data)}")
                        if recommendations_data:
                            msgs.append(f"{len(recommendations_data)} recomendaciones extraídas")
                        if opportunities_data:
                            msgs.append(f"{len(opportunities_data)} oportunidades extraídas")
                        if msgs:
                            st.success(" · ".join(msgs))
                        else:
                            st.warning("No se pudieron extraer datos de la URL. Claude los generará desde el screenshot.")

                with st.spinner("Claude analizando el screenshot..."):
                    try:
                        pulse_file.seek(0)
                        result = analyze_with_claude(pulse_file.read(), empresa, competitors_data, recommendations_data, opportunities_data, observaciones_t1)
                    except Exception as exc:
                        st.error(f"Error al analizar la imagen: {exc}")
                        import traceback
                        with st.expander("Detalle"):
                            st.code(traceback.format_exc())
                        st.stop()

                score = result.get("score_global", 0)
                gpt   = result.get("chatgpt_score", 0)
                gem   = result.get("gemini_score", 0)
                cla   = result.get("claude_score", 0)
                per   = result.get("perplexity_score", 0)

                st.success(f"Análisis completado — Score global: **{score}/100**")
                with st.expander("DEBUG: competidores extraídos por Claude"):
                    st.json(result.get("competitors", "campo 'competitors' no encontrado"))
                m1, m2 = st.columns(2)
                m1.metric("ChatGPT", f"{gpt}/100")
                m2.metric("Gemini",  f"{gem}/100")
                m3, m4 = st.columns(2)
                m3.metric("Claude",    f"{cla}/100")
                m4.metric("Perplexity", f"{per}/100")

                with st.spinner("Generando presentación..."):
                    try:
                        logo_bytes = logo_file.getvalue() if logo_file else None
                        pptx_bytes = generate_pptx(empresa.strip(), fecha, logo_bytes, result)
                        filename   = f"Auditoria_GEO_{empresa.strip().replace(' ', '_')}_{fecha.strftime('%Y%m')}.pptx"
                    except Exception as exc:
                        st.error(f"Error al generar el PPTX: {exc}")
                        import traceback
                        with st.expander("Detalle"):
                            st.code(traceback.format_exc())
                        st.stop()

                st.download_button(
                    label="Descargar PPTX",
                    data=pptx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    type="primary",
                )

# ══════════════════════════════════════════════════════════════
# TAB 2 — escáner GEO por dominio
# ══════════════════════════════════════════════════════════════
with tab2:
    st.caption("Introduce un dominio → consultamos ChatGPT, Gemini, Claude y Llama → score de visibilidad → PPTX.")

    col_dom, col_pais = st.columns([3, 1])
    domain_input = col_dom.text_input("Dominio de la empresa", placeholder="Ej: lafabricadelseo.com", key="domain_input")
    pais_manual  = col_pais.text_input("País", placeholder="Ej: España", key="pais_manual")
    st.caption("Se enviarán 5 preguntas específicas a cada modelo: competidores, reputación, fortalezas, oportunidades y prompts de búsqueda.")

    with st.expander("Estado de API keys"):
        st.write(f"{'✅' if ANTHROPIC_KEY else '❌'} Anthropic (Claude)")
        st.write(f"{'✅' if OPENAI_KEY else '❌'} OpenAI (ChatGPT)")
        st.write(f"{'✅' if GEMINI_KEY else '❌'} Gemini")
        st.write(f"{'✅' if GROQ_KEY else '❌'} Groq (Grok)")

    col_logo, col_ahrefs = st.columns(2)
    with col_logo:
        logo_geo = st.file_uploader("Logo empresa (opcional, para el PPTX)", type=["jpg", "jpeg", "png"], key="logo_geo")
    with col_ahrefs:
        ahrefs_file = st.file_uploader(
            "Screenshot Ahrefs AI Citations (opcional)",
            type=["jpg", "jpeg", "png"],
            key="ahrefs_file",
            help="Captura de la sección AI Citations de Ahrefs. Claude la interpretará y enriquecerá el informe con datos reales de citaciones."
        )
        if ahrefs_file:
            st.image(ahrefs_file, caption="Ahrefs AI Citations", use_container_width=True)

    observaciones_t2 = st.text_area(
        "Observaciones del cliente (opcional)",
        placeholder="Ej: El cliente quiere potenciar su presencia en ChatGPT. Acaba de lanzar un nuevo servicio. Compite principalmente con X e Y...",
        height=120,
        key="obs_t2",
        help="Claude usará estas notas para personalizar el análisis y las recomendaciones.",
    )

    if st.button("Analizar visibilidad GEO", type="primary", use_container_width=True):
        errors = []
        if not domain_input.strip():
            errors.append("Introduce un dominio.")
        if not ANTHROPIC_KEY:
            errors.append("ANTHROPIC_API_KEY no configurada.")
        if not OPENAI_KEY:
            errors.append("OPENAI_API_KEY no configurada.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            # Paso 1 — Detectar marca y sector
            with st.spinner("Detectando marca y sector..."):
                try:
                    info   = geo_detect_brand(domain_input.strip(), observaciones_t2)
                    brand  = info.get("brand", domain_input)
                    sector = info.get("sector", "empresa")
                    pais   = pais_manual.strip() if pais_manual.strip() else info.get("pais", "España")
                except Exception as e:
                    st.error(f"Error detectando marca: {e}")
                    st.stop()

            st.info(f"**Marca:** {brand}  ·  **Sector:** {sector}  ·  **País:** {pais}")

            # Paso 2 — Generar preguntas por sección
            section_prompts = geo_section_prompts(brand, sector, pais)
            with st.expander("Preguntas enviadas a cada modelo"):
                for key, p in section_prompts.items():
                    st.write(f"**{key.upper()}:** {p}")

            # Paso 3 — Consultar modelos (una pregunta por sección)
            active_models = {"Claude": geo_query_claude, "ChatGPT": geo_query_openai}
            if GEMINI_KEY:
                active_models["Gemini"] = geo_query_gemini
            if GROQ_KEY:
                active_models["Grok"]   = geo_query_groq

            all_responses = {m: {} for m in active_models}
            scores        = {}
            total_calls   = len(active_models) * len(section_prompts)
            done          = 0
            progress_bar  = st.progress(0, text="Consultando modelos...")

            for model_name, query_fn in active_models.items():
                for section_key, prompt in section_prompts.items():
                    resp = _query_safe(query_fn, prompt)
                    all_responses[model_name][section_key] = resp
                    done += 1
                    progress_bar.progress(done / total_calls, text=f"{model_name} → {section_key}... ({done}/{total_calls})")
                scores[model_name] = geo_score_from_sections(brand, all_responses[model_name])

            progress_bar.empty()

            # Detectar modelos con errores y mostrar avisos
            _ERROR_HINTS = {
                "429": "cuota agotada — añade créditos en la plataforma del proveedor",
                "insufficient_quota": "cuota agotada — añade créditos en la plataforma del proveedor",
                "401": "API key inválida o revocada — revisa el valor en Secrets",
                "403": "acceso denegado — verifica permisos de la API key",
                "404": "modelo no encontrado — contacta soporte",
            }
            for model_name, sections in all_responses.items():
                errs = [r for r in sections.values() if "[ERROR" in r]
                if errs:
                    sample = errs[0]
                    hint = next((msg for key, msg in _ERROR_HINTS.items() if key in sample), "error desconocido")
                    st.warning(f"**{model_name}** no pudo responder ({hint}). Su score será 0.")

            # Paso 4 — Mostrar scores
            score_global = round(sum(scores.values()) / len(scores))
            st.subheader("Visibilidad por modelo")

            score_cols = st.columns(len(scores) + 1)
            for i, (model, s) in enumerate(scores.items()):
                score_cols[i].metric(model, f"{s}/100")
            score_cols[-1].metric("Score Global", f"{score_global}/100")

            with st.expander("Respuestas completas de los modelos"):
                for model_name, sections in all_responses.items():
                    st.markdown(f"### {model_name}")
                    for section_key, resp in sections.items():
                        st.markdown(f"**{section_key.upper()}**")
                        st.write(resp)

            # Guardar todo en session_state para persistir entre reruns
            with st.spinner("Claude generando el informe completo..."):
                try:
                    ahrefs_bytes = ahrefs_file.getvalue() if ahrefs_file else None
                    analysis = geo_analyze_results(brand, sector, pais, scores, all_responses, observaciones_t2, ahrefs_bytes)
                    analysis["score_global"]     = score_global
                    analysis["chatgpt_score"]    = scores.get("ChatGPT", 0)
                    analysis["gemini_score"]     = scores.get("Gemini", 0)
                    analysis["claude_score"]     = scores.get("Claude", 0)
                    analysis["perplexity_score"] = scores.get("Grok", 0)
                    st.session_state["geo_analysis"]     = analysis
                    st.session_state["geo_scores"]       = scores
                    st.session_state["geo_score_global"] = score_global
                    st.session_state["geo_brand"]        = brand
                except Exception as e:
                    st.error(f"Error generando el informe: {e}")
                    import traceback
                    with st.expander("Detalle"):
                        st.code(traceback.format_exc())

    # ── Informe y PPTX (fuera del botón Analizar, persiste entre reruns) ──
    if st.session_state.get("geo_analysis"):
        analysis     = st.session_state["geo_analysis"]
        scores_ss    = st.session_state.get("geo_scores", {})
        score_global = st.session_state.get("geo_score_global", 0)
        brand_ss     = st.session_state.get("geo_brand", "")

        comps      = analysis.get("competitors", [])
        opps       = analysis.get("opportunities", [])
        strengths  = analysis.get("strengths", [])
        reputation = analysis.get("brand_reputation", {})
        insights   = analysis.get("ai_search_insights", [])
        s_prompts  = analysis.get("sample_prompts", [])
        recs       = analysis.get("recommendations", [])

        st.divider()

        # Scores
        st.subheader("Visibilidad por modelo")
        score_cols = st.columns(len(scores_ss) + 1)
        for i, (model, s) in enumerate(scores_ss.items()):
            score_cols[i].metric(model, f"{s}/100")
        score_cols[-1].metric("Score Global", f"{score_global}/100")

        st.divider()

        # 1. Top Competitors
        st.subheader("Top Competitors")
        if comps:
            cols_c = st.columns(len(comps))
            for i, c in enumerate(comps):
                with cols_c[i]:
                    stars = "⭐" * max(1, min(5, c.get("stars", 3)))
                    st.markdown(f"**{c.get('name', '')}**")
                    st.write(stars)
                    st.caption(c.get("desc", ""))
        else:
            st.info("No se detectaron competidores en las respuestas.")

        st.divider()

        # 2. Brand Reputation
        st.subheader("Brand Reputation")
        tone = reputation.get("tone", "")
        tone_color = {"Ausente": "🔴", "Negativa": "🔴", "Neutra": "🟡", "Positiva": "🟢"}.get(tone, "⚪")
        st.markdown(f"{tone_color} **{tone}** — {reputation.get('summary', '')}")
        attrs = reputation.get("attributes", [])
        if attrs:
            st.write(" · ".join(f"`{a}`" for a in attrs))

        st.divider()

        # 3. Key Strengths
        st.subheader("Key Strengths")
        cols_s = st.columns(2)
        for i, s in enumerate(strengths):
            cols_s[i % 2].markdown(f"✅ {s}")

        st.divider()

        # 4. Improvement Opportunities
        st.subheader("Improvement Opportunities")
        cols_o = st.columns(2)
        for i, opp in enumerate(opps):
            cols_o[i % 2].markdown(f"🔧 {opp}")

        st.divider()

        # 5. Sample Search Prompts
        st.subheader("Sample Search Prompts")
        st.caption("Búsquedas donde debería aparecer la marca pero no aparece:")
        for p in s_prompts:
            st.markdown(f"💬 _{p}_")

        st.divider()

        # 6. AI Search Insights
        st.subheader("AI Search Insights")
        vis_color = {"Nula": "🔴", "Baja": "🟠", "Media": "🟡", "Alta": "🟢"}
        for ins in insights:
            icon = vis_color.get(ins.get("visibility", ""), "⚪")
            st.markdown(f"{icon} **{ins.get('model', '')}** ({ins.get('visibility', '')}) — {ins.get('insight', '')}")

        st.divider()

        # 7. Recommendations
        st.subheader("Recommendations")
        pri_label = {3: "🔴 Alta", 2: "🟡 Media", 1: "🟢 Baja"}
        for r in recs:
            st.markdown(f"{pri_label.get(r.get('priority', 2), '')} **{r.get('title', '')}** — {r.get('desc', '')}")

        st.divider()

        # Generar PPTX
        st.subheader("Generar auditoría PPTX")
        col_e, col_f = st.columns(2)
        empresa_geo = col_e.text_input("Nombre empresa", value=brand_ss, key="empresa_geo")
        fecha_geo   = col_f.date_input("Fecha", value=date.today(), key="fecha_geo")

        if st.button("Generar PPTX con estos resultados", key="btn_pptx_geo"):
            with st.spinner("Generando PPTX..."):
                try:
                    logo_bytes = logo_geo.getvalue() if logo_geo else None
                    pptx_bytes = generate_pptx(empresa_geo.strip(), fecha_geo, logo_bytes, analysis, fourth_model="Grok / Llama")
                    filename   = f"Auditoria_GEO_{empresa_geo.strip().replace(' ', '_')}_{fecha_geo.strftime('%Y%m')}.pptx"
                except Exception as e:
                    st.error(f"Error generando PPTX: {e}")
                    import traceback
                    with st.expander("Detalle"):
                        st.code(traceback.format_exc())
                    st.stop()

            st.download_button(
                label="Descargar PPTX",
                data=pptx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary",
            )
