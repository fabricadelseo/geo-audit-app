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
try:
    ANTHROPIC_KEY = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
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
Analiza esta captura de pantalla de LLMs Pulse, herramienta de medición de visibilidad de marca en modelos de IA.

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
  "chatgpt_analysis": "Hallazgos:\\n• <hallazgo 1>\\n• <hallazgo 2>\\n• <hallazgo 3>\\n• <hallazgo 4>\\n\\nBrechas identificadas:\\n\\u2717 <brecha 1>\\n\\u2717 <brecha 2>\\n\\u2717 <brecha 3>",
  "gemini_analysis": "Hallazgos:\\n• <hallazgo 1>\\n• <hallazgo 2>\\n• <hallazgo 3>\\n\\nBrechas identificadas:\\n\\u2717 <brecha 1>\\n\\u2717 <brecha 2>\\n\\u2717 <brecha 3>",
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
Textos en español, concretos y accionables. Priority: 1=baja, 2=media, 3=alta.\
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

        comp_chunk = _extract_chunk("top competitor") or _extract_chunk("competitor")
        rec_chunk  = _extract_chunk("recommended next step") or _extract_chunk("recommendation")
        opp_chunk  = _extract_chunk("improvement opportunit") or _extract_chunk("opportunity")
        combined   = (
            f"--- COMPETITORS SECTION ---\n{comp_chunk}\n\n"
            f"--- RECOMMENDATIONS SECTION ---\n{rec_chunk}\n\n"
            f"--- OPPORTUNITIES SECTION ---\n{opp_chunk}"
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

def analyze_with_claude(image_bytes: bytes, empresa: str, competitors_data: list, recommendations_data: list, opportunities_data: list) -> dict:
    client   = Anthropic(api_key=ANTHROPIC_KEY)
    media_type = "image/png" if image_bytes[:4] == b"\x89PNG" else "image/jpeg"
    b64      = base64.standard_b64encode(image_bytes).decode("utf-8")
    contexto = f"La empresa auditada es: {empresa.strip()}." if empresa.strip() else ""
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
                new_r.append(copy.deepcopy(template_rPr))
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
        # Centrar dentro del placeholder original
        orig_left = shape.left
        orig_top  = shape.top
        shape.width  = Emu(new_w)
        shape.height = Emu(new_h)
        shape.left   = orig_left + (max_size - new_w) // 2
        shape.top    = orig_top  + (max_size - new_h) // 2
    except Exception:
        pass  # si falla, deja las dimensiones originales

    slide.part.rels[rId].target_part._blob = img_bytes


# ─────────────────────────────────────────────────────────────
# GENERADOR PPTX
# ─────────────────────────────────────────────────────────────

def generate_pptx(empresa: str, fecha: date, logo_bytes, d: dict) -> bytes:
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
    set_run(s, "Text 4", f"Análisis de visibilidad en LLMs · {date_str}")

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
    set_run(s, "Text 19", f"{per}/100"); set_run(s, "Text 20", model_label(per, "Perplexity"))
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

    # Slide 5 — Análisis ChatGPT
    s = prs.slides[4]
    set_run(s, "Text 1", f"Score: {gpt}/100")
    set_body(s, "Text 2", d["chatgpt_analysis"].split("\n"))

    # Slide 6 — Análisis Gemini
    s = prs.slides[5]
    set_run(s, "Text 1", f"Score: {gem}/100")
    set_body(s, "Text 2", d["gemini_analysis"].split("\n"))

    # Slide 7 — Fortalezas
    s = prs.slides[6]
    for i, name in enumerate(["Text 2", "Text 3", "Text 4", "Text 5"]):
        txt = d["strengths"][i] if i < len(d["strengths"]) else ""
        if txt.strip():
            set_run(s, name, f"{i+1}. {txt}")

    # Slide 8 — Oportunidades
    s = prs.slides[7]
    for i, name in enumerate(["Text 2", "Text 3", "Text 4", "Text 5"]):
        txt = d["opportunities"][i] if i < len(d["opportunities"]) else ""
        if txt.strip():
            set_run(s, name, f"{i+1}. {txt}")

    # Slide 9 — Recomendaciones
    s = prs.slides[8]
    rec_map = [
        ("Text 3",  "Text 4",  "Text 5"),
        ("Text 8",  "Text 9",  "Text 10"),
        ("Text 13", "Text 14", "Text 15"),
        ("Text 18", "Text 19", "Text 20"),
        ("Text 23", "Text 24", "Text 25"),
    ]
    for i, (tn, dn, sn) in enumerate(rec_map):
        if i < len(d["recommendations"]):
            r = d["recommendations"][i]
            if r.get("title"): set_run(s, tn, r["title"][:55])
            if r.get("desc"):  set_run(s, dn, r["desc"][:95])
            set_run(s, sn, "⭐" * max(1, r.get("priority", 3)))

    # Slide 10 — Próximos pasos
    s = prs.slides[9]
    for i, (tn, dn) in enumerate([("Text 2","Text 3"),("Text 5","Text 6"),("Text 8","Text 9"),("Text 11","Text 12")]):
        if i < len(d["steps"]):
            st_ = d["steps"][i]
            if st_.get("title"): set_run(s, tn, st_["title"][:55])
            if st_.get("desc"):  set_run(s, dn, st_["desc"][:95])

    # Slide 11 — Logo
    if logo_bytes:
        replace_logo(prs.slides[10], "Picture 2", logo_bytes)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="Generador Auditoría GEO", page_icon="🔍", layout="wide")

st.title("Generador de Auditorías GEO")
st.caption("Rellena 4 datos → Claude analiza el screenshot → descarga el PPTX.")

# ── SIDEBAR — solo 4 inputs ───────────────────────────────────
with st.sidebar:
    st.header("Datos de la auditoría")

    empresa   = st.text_input("Nombre de la empresa *", placeholder="Ej: Acme Corp")
    fecha     = st.date_input("Fecha de auditoría", value=date.today())
    logo_file = st.file_uploader("Logo de la empresa (opcional)", type=["jpg", "jpeg", "png", "svg"])
    if logo_file:
        if logo_file.name.lower().endswith(".svg"):
            st.success(f"Logo SVG cargado: {logo_file.name}")
        else:
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
        help="Si pegas la URL del informe, los competidores se extraen automáticamente y con precisión. Si no, Claude intentará leerlos del screenshot.",
    )


# ── MAIN ─────────────────────────────────────────────────────
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
            # Paso 1 — Extraer competidores y recomendaciones de la URL
            competitors_data    = []
            recommendations_data = []
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

            # Paso 2 — Claude analiza la imagen
            with st.spinner("Claude analizando el screenshot..."):
                try:
                    pulse_file.seek(0)
                    result = analyze_with_claude(pulse_file.read(), empresa, competitors_data, recommendations_data, opportunities_data)
                except Exception as exc:
                    st.error(f"Error al analizar la imagen: {exc}")
                    import traceback
                    with st.expander("Detalle"):
                        st.code(traceback.format_exc())
                    st.stop()

            # Paso 2 — Mostrar scores extraídos
            score = result.get("score_global", 0)
            gpt   = result.get("chatgpt_score", 0)
            gem   = result.get("gemini_score", 0)
            cla   = result.get("claude_score", 0)
            per   = result.get("perplexity_score", 0)
            avg   = round((gpt + gem + cla + per) / 4)

            st.success(f"Análisis completado — Score global: **{score}/100**")
            with st.expander("DEBUG: competidores extraídos por Claude"):
                st.json(result.get("competitors", "⚠️ campo 'competitors' no encontrado"))
            m1, m2 = st.columns(2)
            m1.metric("ChatGPT", f"{gpt}/100")
            m2.metric("Gemini",  f"{gem}/100")
            m3, m4 = st.columns(2)
            m3.metric("Claude",    f"{cla}/100")
            m4.metric("Perplexity", f"{per}/100")

            # Paso 3 — Generar PPTX
            with st.spinner("Generando presentación..."):
                try:
                    if logo_file:
                        logo_bytes = logo_file.getvalue()
                        if logo_file.name.lower().endswith(".svg"):
                            from svglib.svglib import svg2rlg
                            from reportlab.graphics import renderPM
                            rlg = svg2rlg(io.BytesIO(logo_bytes))
                            logo_bytes = renderPM.drawToString(rlg, fmt="PNG")
                    else:
                        logo_bytes = None
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
