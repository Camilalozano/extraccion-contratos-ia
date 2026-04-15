import io
import json
import os
import re
import sys
import unicodedata
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openai import OpenAI

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

INPUT_ZIP_PATH = r""
OUTPUT_EXCEL_PATH = r"D:\Users\Usuario\Documents\ArchivosExtraidos.xlsx"
USE_AI = True
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

TARGET_FIELDS = [
    "numero_contrato",
    "Tipo_contrato",
    "nombre_contratista",
    "numero_documento_contratista",
    "obligaciones_especificas",
    "nombre_supervisor",
]


# =========================
# Utilidades
# =========================
def safe_json_loads(raw: str) -> dict:
    raw = (raw or "").strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


def normalize_nullable_text(x: Optional[str]) -> str:
    if x is None:
        return ""
    x = str(x).strip()
    x = re.sub(r"[ \t]+", " ", x)
    x = re.sub(r"\n{3,}", "\n\n", x)
    return x.strip()


def only_digits(x: Optional[str]) -> str:
    if not x:
        return ""
    x = str(x).translate(str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1"}))
    return re.sub(r"\D", "", x)


def normalize_spaces(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def limpiar_texto_para_llm(text: str) -> str:
    if not text:
        return ""
    t = unicodedata.normalize("NFKC", text)
    for a, b in [("\u00A0", " "), ("\u200B", ""), ("\u200E", ""), ("\u200F", "")]:
        t = t.replace(a, b)
    cleaned = []
    for ch in t:
        cat = unicodedata.category(ch)
        if cat.startswith("C") and ch not in ["\n", "\t"]:
            continue
        cleaned.append(ch)
    t = "".join(cleaned)
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def normalize_text(text: str) -> str:
    if not text:
        return ""
    return normalize_spaces(limpiar_texto_para_llm(text))


def search_first(patterns: List[str], text: str, flags: int = re.IGNORECASE | re.DOTALL):
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return match
    return None


def looks_like_person_name(value: str) -> bool:
    if not value:
        return False
    value = re.sub(r"\s+", " ", value).strip(" ,.;:\n\t")
    words = [w for w in value.split() if w]
    if len(words) < 2:
        return False
    upper_words = sum(1 for w in words if re.fullmatch(r"[A-ZÁÉÍÓÚÑ]+(?:[-'][A-ZÁÉÍÓÚÑ]+)?", w))
    return upper_words >= min(2, len(words))


def looks_like_entity_name(value: str) -> bool:
    if not value:
        return False
    v = value.upper()
    entity_markers = [
        "S.A.S", "S.A.", "LTDA", "E.S.P", "UNIVERSIDAD", "CORPORACIÓN", "CORPORACION",
        "FUNDACIÓN", "FUNDACION", "CAJA DE COMPENSACIÓN", "CAJA DE COMPENSACION",
        "EMPRESA", "ASOCIACIÓN", "ASOCIACION", "COLEGIO", "INSTITUTO", "ETB", "CAFAM",
        "NACIONAL", "DISTRITAL"
    ]
    return any(marker in v for marker in entity_markers)


def clean_contract_name(value: str) -> str:
    value = normalize_nullable_text(value)
    value = re.sub(r"^(la|el)\s+otra,?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^la\s+tecnolog[íi]a\s+y\s+", "", value, flags=re.IGNORECASE)
    return value.strip(" ,.;:\n\t")


def cut_text(text: str, limit: int = 16000) -> str:
    return text[:limit] if len(text) > limit else text


# =========================
# PDF
# =========================
def extract_text_with_pymupdf(pdf_bytes: bytes) -> str:
    if fitz is None:
        return ""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        pages = []
        for page in doc:
            try:
                pages.append(page.get_text("text") or "")
            except Exception:
                pages.append("")
        return "\n".join(pages)
    except Exception:
        return ""


def extract_text_with_pypdf(pdf_bytes: bytes) -> str:
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        pages = []
        for page in reader.pages:
            try:
                pages.append(page.extract_text() or "")
            except Exception:
                pages.append("")
        return "\n".join(pages)
    except Exception:
        return ""


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    text = extract_text_with_pymupdf(pdf_bytes)
    if len(normalize_text(text)) >= 200:
        return text
    text_alt = extract_text_with_pypdf(pdf_bytes)
    if len(normalize_text(text_alt)) > len(normalize_text(text)):
        text = text_alt
    return text


# =========================
# Clasificación básica
# =========================
def classify_document(text: str, filename: str = "") -> str:
    txt = (filename + "\n" + text[:4000]).upper()
    if "MEMORANDO" in txt:
        return "memorando"
    if "DOCUMENTOS PREVIOS" in txt or "SOLICITUD ORDENACIÓN DE CONTRATACIÓN" in txt or "SOLICITUD ORDENACION DE CONTRATACION" in txt:
        return "estudios_previos"
    if "ACTO ADMINISTRATIVO" in txt or "MEDIANTE EL CUAL SE JUSTIFICA" in txt or "POR LA CUAL SE JUSTIFICA" in txt:
        return "acto_justificacion"
    if "INFORME DE VERIFICACIÓN" in txt or "INFORME DE VERIFICACION" in txt or "PROPONENTE" in txt:
        return "evaluacion"
    if "CONTRATO" in txt or "CONVENIO" in txt or "CLAUSULADO" in txt:
        return "contractual"
    return "otro"


# =========================
# Extracción por reglas
# =========================
def extract_contract_number(text: str, filename: str = "") -> str:
    patterns = [
        r"(?:CONTRATO|CONVENIO)[^\n]{0,120}?No\.?\s*([A-Z0-9\-_/]+(?:\s*[-–]\s*\d{4})?)",
        r"No\.?\s*(ATENEA\s*[-–]\s*\d+\s*[-–]\s*\d{4})",
        r"(ATENEA\s*[-–]\s*\d+\s*[-–]\s*\d{4})",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            value = re.sub(r"\s+", "", m.group(1)).replace("–", "-")
            return value.strip(" .,:;\n\t")
    m_file = re.search(r"(ATENEA[-_]\d+[-_]\d{4})", filename, re.IGNORECASE)
    if m_file:
        return m_file.group(1).replace("_", "-")
    return ""


def extract_contract_type(text: str) -> str:
    patterns = [
        r"(?:CONTRATO|CLAUSULADO DEL CONTRATO)\s+DE\s+([A-ZÁÉÍÓÚÑ\s]+?)\s+No\.?",
        r"CONVENIO\s+([A-ZÁÉÍÓÚÑ\s]+?)\s+No\.?",
        r"presente\s+contrato\s+de\s+([A-ZÁÉÍÓÚÑ\s]+?)(?:\s+el\s+cual|\s+que\s+se\s+regir[áa]|\s+de\s+conformidad)",
    ]
    m = search_first(patterns, text)
    if not m:
        return ""
    return re.sub(r"\s+", " ", m.group(1)).strip(" ,.;:\n\t").upper()


def extract_name_from_header(text: str) -> str:
    patterns = [
        r"celebrado\s+entre\s+[^\n]{0,180}?\s+y\s+([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?)\.",
        r"entre\s+la\s+AGENCIA[^\n]{0,220}?\s+y\s+([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?)\.",
        r"entre\s+la\s+AGENCIA[^\n]{0,220}?\s+y\s+([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?)(?:,|\.)",
    ]
    m = search_first(patterns, text[:3500])
    if not m:
        return ""
    return clean_contract_name(m.group(1))


def extract_name_from_party_block(text: str) -> str:
    patterns = [
        r"y\s+por\s+la\s+otra,\s*([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?),\s*mayor\s+de\s+edad",
        r"y\s+por\s+la\s+otra,\s*([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?),\s*identificad[oa]",
        r"por\s+la\s+otra,\s*la\s+([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?),\s*con\s+NIT",
        r"por\s+la\s+otra,\s*el\s+([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?),\s*con\s+NIT",
        r"por\s+la\s+otra\s+parte,\s*([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?),\s*(?:mayor\s+de\s+edad|identificad[oa]|actuando)",
        r"entre\s+la\s+AGENCIA[^\n]{0,220}?\s+y\s+la\s+([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?),\s*con\s+NIT",
        r"entre\s+la\s+AGENCIA[^\n]{0,220}?\s+y\s+el\s+([A-ZÁÉÍÓÚÑ0-9\.\-\s]+?),\s*con\s+NIT",
    ]
    m = search_first(patterns, text[:9000])
    if not m:
        return ""
    return clean_contract_name(m.group(1))


def extract_contractor_name(text: str, filename: str = "") -> str:
    candidates = []
    for candidate in [extract_name_from_header(text), extract_name_from_party_block(text)]:
        if candidate:
            candidates.append(candidate)

    # Prefer entity names when there is NIT nearby; otherwise prefer person names.
    if candidates:
        candidates = list(dict.fromkeys(candidates))
        entity_candidates = [c for c in candidates if looks_like_entity_name(c)]
        person_candidates = [c for c in candidates if looks_like_person_name(c)]
        if entity_candidates:
            return entity_candidates[0]
        if person_candidates:
            return person_candidates[0]
        return candidates[0]

    # Fallback from filename on minuta contractual person-based docs
    stem = Path(filename).stem.upper()
    m = re.search(r"MINUTA\s+CONTRACTUAL\s+(.+)$", stem)
    if m:
        return clean_contract_name(m.group(1).replace("ATENEA", "").replace("-", " "))
    return ""


def build_document_candidates(text: str, contractor_name: str = "") -> List[str]:
    text_norm = normalize_spaces(text)
    candidates: List[str] = []

    num_pattern = r"([0-9OIl][0-9OIl\.\,\-\s]{5,}[0-9OIl])"

    if contractor_name:
        contractor_name_esc = re.escape(normalize_spaces(contractor_name))
        m_name = re.search(contractor_name_esc, text_norm, re.IGNORECASE)
        if m_name:
            start = max(0, m_name.start() - 120)
            end = min(len(text_norm), m_name.end() + 450)
            window = text_norm[start:end]

            local_patterns = [
                rf"{contractor_name_esc}.{{0,180}}?c[ée]dula\s+de\s+ciudadan[íi]a\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
                rf"{contractor_name_esc}.{{0,180}}?C\.?\s*C\.?\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
                rf"{contractor_name_esc}.{{0,220}}?NIT\s*(?:No\.?|#|:)?\s*{num_pattern}",
                rf"{contractor_name_esc}.{{0,220}}?identificad[oa].{{0,80}}?{num_pattern}",
            ]
            for pattern in local_patterns:
                for m in re.finditer(pattern, window, re.IGNORECASE | re.DOTALL):
                    candidates.append(only_digits(m.group(1)))

    party_block_patterns = [
        rf"por\s+la\s+otra[^\n]{{0,600}}?c[ée]dula\s+de\s+ciudadan[íi]a\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
        rf"por\s+la\s+otra[^\n]{{0,600}}?C\.?\s*C\.?\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
        rf"por\s+la\s+otra[^\n]{{0,600}}?NIT\s*(?:No\.?|#|:)?\s*{num_pattern}",
        rf"representada\s+legalmente\s+por\s+[^\n]{{0,180}}?c[ée]dula\s+de\s+ciudadan[íi]a\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
    ]
    head = text_norm[:12000]
    for pattern in party_block_patterns:
        for m in re.finditer(pattern, head, re.IGNORECASE | re.DOTALL):
            candidates.append(only_digits(m.group(1)))

    # Header candidate for juridical entities.
    for m in re.finditer(r"\bNIT\s*(?:No\.?|#|:)?\s*([0-9\.\-\s]{6,25})", head, re.IGNORECASE):
        candidates.append(only_digits(m.group(1)))

    return [c for c in candidates if 6 <= len(c) <= 12]


def score_document_candidate(candidate: str, contractor_name: str = "") -> int:
    score = 0
    if not candidate:
        return -999
    if 8 <= len(candidate) <= 10:
        score += 3
    if len(candidate) == 9:
        score += 1  # often NIT base
    if len(candidate) == 10:
        score += 1  # often CC
    if looks_like_entity_name(contractor_name) and len(candidate) in (9, 10):
        score += 2
    if looks_like_person_name(contractor_name) and len(candidate) in (8, 10):
        score += 2
    return score


def extract_contractor_document(text: str, contractor_name: str = "") -> str:
    candidates = build_document_candidates(text, contractor_name)
    if not candidates:
        return ""
    ranked = sorted(candidates, key=lambda c: (score_document_candidate(c, contractor_name), len(c)), reverse=True)
    return ranked[0]


def extract_obligaciones_especificas(text: str) -> str:
    start_patterns = [
        r"B\)\s*OBLIGACIONES\s+ESPEC[ÍI]FICAS\s*:\s*A\s*EL\s+CONTRATISTA\s+le\s+corresponde\s+el\s+cumplimiento\s+de\s+las\s+siguientes\s+obligaciones\s*:",
        r"B\)\s*OBLIGACIONES\s+ESPEC[ÍI]FICAS\s*:",
        r"OBLIGACIONES\s+ESPEC[ÍI]FICAS\s*:",
    ]
    end_patterns = [
        r"C\)\s*OBLIGACIONES\s+DE\s+LA\s+CONTRATANTE",
        r"OBLIGACIONES\s+DE\s+LA\s+CONTRATANTE",
        r"CL[ÁA]USULA\s+TERCERA\s*:",
        r"CL[ÁA]USULA\s+CUARTA\s*:",
        r"CL[ÁA]USULA\s+DE\s+SUPERVISI[ÓO]N",
    ]
    start_match = search_first(start_patterns, text)
    if not start_match:
        return ""
    tail = text[start_match.end():]
    end_match = search_first(end_patterns, tail)
    obligaciones = tail[: end_match.start()] if end_match else tail
    obligaciones = obligaciones.strip()
    obligaciones = re.sub(r"\n{3,}", "\n\n", obligaciones)
    return obligaciones.strip()


def extract_supervisor_name(text: str) -> str:
    patterns = [
        r"(?:la\s+)?supervisi[óo]n\s+(?:del\s+presente\s+contrato|contractual)?\s*(?:ser[áa]\s+ejercida|estar[áa]\s+a\s+cargo|corresponder[áa])\s+por\s+([^\n\.]+)",
        r"supervisor(?:a)?\s+del\s+contrato\s*(?:ser[áa]|es|:)?\s*([^\n\.]+)",
        r"la\s+supervisi[óo]n\s+ser[áa]\s+ejercida\s+por\s+([^\n\.]+)",
    ]
    m = search_first(patterns, text)
    if not m:
        return ""
    value = m.group(1)
    value = re.split(r";|,\s*quien\s+|,\s*o\s+por\s+quien|\s+o\s+por\s+quien", value, maxsplit=1, flags=re.IGNORECASE)[0]
    value = re.sub(r"\s+", " ", value).strip(" ,.;:\n\t")
    value = re.sub(r"^(el|la|los|las)\s+", "", value, flags=re.IGNORECASE)
    return value


# =========================
# IA
# =========================
def get_openai_client(api_key: Optional[str]) -> Optional[OpenAI]:
    if not api_key:
        return None
    try:
        return OpenAI(api_key=api_key)
    except Exception:
        return None


def build_focus_context(text: str, contractor_name_rule: str, contractor_doc_rule: str, obligaciones_regla: str, supervisor_regla: str) -> str:
    parts = []
    head = cut_text(text[:9000], 9000)
    parts.append("=== INICIO DEL DOCUMENTO ===\n" + head)

    if contractor_name_rule or contractor_doc_rule:
        parts.append(
            "=== CANDIDATOS POR REGLAS PARA CONTRATISTA ===\n"
            f"nombre_contratista: {contractor_name_rule}\n"
            f"numero_documento_contratista: {contractor_doc_rule}"
        )

    if obligaciones_regla:
        parts.append("=== BLOQUE DETECTADO DE OBLIGACIONES ESPECÍFICAS ===\n" + cut_text(obligaciones_regla, 7000))

    if supervisor_regla:
        parts.append("=== CANDIDATO DE SUPERVISIÓN ===\n" + supervisor_regla)

    party = re.search(r"por\s+la\s+otra", text, re.IGNORECASE)
    if party:
        start = max(0, party.start() - 500)
        end = min(len(text), party.start() + 3000)
        parts.append("=== CONTEXTO DEL CONTRATISTA ===\n" + text[start:end])

    return "\n\n".join(parts)


def extract_contract_fields_raw(client: OpenAI, text: str, filename: str, rule_candidates: dict, doc_class: str) -> str:
    focus_text = build_focus_context(
        text=text,
        contractor_name_rule=rule_candidates.get("nombre_contratista", ""),
        contractor_doc_rule=rule_candidates.get("numero_documento_contratista", ""),
        obligaciones_regla=rule_candidates.get("obligaciones_especificas", ""),
        supervisor_regla=rule_candidates.get("nombre_supervisor", ""),
    )

    prompt = f"""
Analiza el siguiente documento contractual en español y devuelve SOLO JSON válido con estos campos:
- numero_contrato
- Tipo_contrato
- nombre_contratista
- numero_documento_contratista
- obligaciones_especificas
- nombre_supervisor

Reglas obligatorias:
1. No inventes datos.
2. Si un campo no aparece claramente, devuelve "".
3. nombre_contratista debe ser la persona o entidad contratista, NO el funcionario de ATENEA.
4. Si el contratista es una entidad, devuelve el nombre de la entidad contratista y en numero_documento_contratista devuelve su NIT si aparece.
5. Si el contratista es persona natural, devuelve su nombre completo y su cédula.
6. numero_documento_contratista debe quedar SOLO con dígitos.
7. obligaciones_especificas debe conservar el texto del bloque, sin resumir.
8. nombre_supervisor debe ser persona o cargo supervisor.
9. Tipo de documento detectado: {doc_class}. Si el documento no es propiamente un contrato/minuta pero sí menciona claramente un contratista o entidad contratista, extrae esa información. Si no, deja campos vacíos.
10. Devuelve únicamente JSON válido, sin explicación, sin markdown.

Archivo: {filename}

Candidatos por reglas:
{json.dumps(rule_candidates, ensure_ascii=False, indent=2)}

Texto:
{focus_text}
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Devuelve SOLO JSON válido. Sin markdown."},
            {"role": "user", "content": prompt},
        ],
        temperature=0,
    )
    return resp.choices[0].message.content


def extract_party_only_raw(client: OpenAI, text: str, filename: str, doc_class: str, rule_name: str, rule_doc: str) -> str:
    party = re.search(r"por\s+la\s+otra", text, re.IGNORECASE)
    if party:
        start = max(0, party.start() - 500)
        end = min(len(text), party.start() + 3500)
        focus = text[start:end]
    else:
        focus = text[:5000]

    prompt = f"""
Extrae SOLO estos campos del siguiente documento y devuelve SOLO JSON válido:
- nombre_contratista
- numero_documento_contratista

Reglas:
1. No inventes datos.
2. nombre_contratista debe ser el contratista o la entidad contratista, NO el funcionario de ATENEA.
3. Si es persona natural, numero_documento_contratista debe ser su cédula.
4. Si es persona jurídica, numero_documento_contratista debe ser su NIT si aparece; si no aparece el NIT pero sí la cédula del representante legal, puedes devolver esa cédula SOLO si es la única identificación explícita asociada al contratista.
5. numero_documento_contratista debe quedar SOLO con dígitos.
6. Tipo de documento detectado: {doc_class}.
7. Devuelve únicamente JSON válido, sin explicación, sin markdown.

Archivo: {filename}
Candidatos por reglas:
- nombre_contratista: {rule_name}
- numero_documento_contratista: {rule_doc}

Texto:
{focus}
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Devuelve SOLO JSON válido. Sin markdown."},
            {"role": "user", "content": prompt},
        ],
        temperature=0,
    )
    return resp.choices[0].message.content


def normalize_ai_result(data: dict) -> dict:
    data = data or {}
    normalized = {field: normalize_nullable_text(data.get(field, "")) for field in TARGET_FIELDS}
    normalized["numero_documento_contratista"] = only_digits(normalized.get("numero_documento_contratista"))
    normalized["numero_contrato"] = normalized.get("numero_contrato", "").replace("–", "-")
    return normalized


def should_override_name(rule_name: str, ai_name: str) -> bool:
    if not ai_name:
        return False
    if not rule_name:
        return True
    if looks_like_entity_name(ai_name) and not looks_like_entity_name(rule_name):
        return True
    if looks_like_person_name(ai_name) and not looks_like_person_name(rule_name):
        return True
    if len(ai_name) > len(rule_name) + 6:
        return True
    return False


def should_override_document(rule_name: str, rule_doc: str, ai_name: str, ai_doc: str) -> bool:
    if not ai_doc:
        return False
    if not rule_doc:
        return True
    if looks_like_entity_name(ai_name or rule_name):
        return len(ai_doc) in (9, 10) and len(rule_doc) not in (9, 10)
    if looks_like_person_name(ai_name or rule_name):
        return len(ai_doc) in (8, 10) and len(rule_doc) not in (8, 10)
    return False


def merge_results(rule_result: dict, ai_result: Optional[dict], ai_party_result: Optional[dict]) -> dict:
    result = {field: normalize_nullable_text(rule_result.get(field, "")) for field in TARGET_FIELDS}
    result["numero_documento_contratista"] = only_digits(result.get("numero_documento_contratista"))

    for ai in [ai_result, ai_party_result]:
        if not ai:
            continue
        ai = normalize_ai_result(ai)

        for field in ["numero_contrato", "Tipo_contrato"]:
            if ai.get(field):
                result[field] = ai[field]

        if should_override_name(result.get("nombre_contratista", ""), ai.get("nombre_contratista", "")):
            result["nombre_contratista"] = ai["nombre_contratista"]

        if should_override_document(
            result.get("nombre_contratista", ""),
            result.get("numero_documento_contratista", ""),
            ai.get("nombre_contratista", ""),
            ai.get("numero_documento_contratista", ""),
        ):
            result["numero_documento_contratista"] = ai["numero_documento_contratista"]
        elif not result.get("numero_documento_contratista") and ai.get("numero_documento_contratista"):
            result["numero_documento_contratista"] = ai["numero_documento_contratista"]

        reglas_obl = result.get("obligaciones_especificas", "")
        ia_obl = ai.get("obligaciones_especificas", "")
        if len(reglas_obl) < 80 and ia_obl:
            result["obligaciones_especificas"] = ia_obl
        elif len(ia_obl) > len(reglas_obl) and len(reglas_obl) < 300:
            result["obligaciones_especificas"] = ia_obl

        supervisor_reglas = result.get("nombre_supervisor", "")
        supervisor_ia = ai.get("nombre_supervisor", "")
        if looks_like_person_name(supervisor_ia):
            result["nombre_supervisor"] = supervisor_ia
        elif not supervisor_reglas and supervisor_ia:
            result["nombre_supervisor"] = supervisor_ia

    return result


# =========================
# Procesamiento
# =========================
def process_single_pdf(pdf_bytes: bytes, filename: str, client: Optional[OpenAI] = None, use_ai: bool = True) -> Dict:
    raw_text = extract_text_from_pdf_bytes(pdf_bytes)
    text = normalize_text(raw_text)
    doc_class = classify_document(text, filename)

    contractor_name = extract_contractor_name(text, filename)
    contractor_doc = extract_contractor_document(text, contractor_name)

    rule_result = {
        "numero_contrato": extract_contract_number(text, filename),
        "Tipo_contrato": extract_contract_type(text),
        "nombre_contratista": contractor_name,
        "numero_documento_contratista": contractor_doc,
        "obligaciones_especificas": extract_obligaciones_especificas(raw_text or text),
        "nombre_supervisor": extract_supervisor_name(text),
    }

    ai_result = None
    ai_party_result = None
    metodo = "reglas"
    error_ia = ""

    if use_ai and client is not None:
        try:
            raw_ai = extract_contract_fields_raw(client, text=text, filename=filename, rule_candidates=rule_result, doc_class=doc_class)
            ai_result = normalize_ai_result(safe_json_loads(raw_ai))
            metodo = "hibrido_reglas_ia"

            # segunda pasada enfocada en contratista cuando quedó vacío o sospechoso
            if (
                not rule_result.get("nombre_contratista")
                or not rule_result.get("numero_documento_contratista")
                or (looks_like_person_name(rule_result.get("nombre_contratista", "")) and len(rule_result.get("numero_documento_contratista", "")) < 8)
            ):
                raw_party = extract_party_only_raw(
                    client,
                    text=text,
                    filename=filename,
                    doc_class=doc_class,
                    rule_name=rule_result.get("nombre_contratista", ""),
                    rule_doc=rule_result.get("numero_documento_contratista", ""),
                )
                ai_party_result = safe_json_loads(raw_party)
        except Exception as e:
            error_ia = str(e)
            metodo = "reglas_con_fallo_ia"

    final_result = merge_results(rule_result, ai_result, ai_party_result)
    final_result.update(
        {
            "archivo": Path(filename).name,
            "tipo_documento_origen": doc_class,
            "metodo_extraccion": metodo,
            "error_ia": error_ia,
            "texto_extraido_chars": len(text),
        }
    )
    return final_result


def process_zip(zip_path: Path, client: Optional[OpenAI] = None, use_ai: bool = True) -> List[Dict]:
    results = []
    with zipfile.ZipFile(zip_path, "r") as zf:
        pdf_files = [name for name in zf.namelist() if name.lower().endswith(".pdf")]
        total = len(pdf_files)
        if total == 0:
            raise ValueError("El ZIP no contiene archivos PDF.")

        for i, name in enumerate(pdf_files, start=1):
            print(f"Procesando {i}/{total}: {Path(name).name}")
            pdf_bytes = zf.read(name)
            try:
                results.append(process_single_pdf(pdf_bytes, name, client=client, use_ai=use_ai))
            except Exception as e:
                results.append(
                    {
                        "archivo": Path(name).name,
                        "tipo_documento_origen": "error",
                        "numero_contrato": "",
                        "Tipo_contrato": "",
                        "nombre_contratista": "",
                        "numero_documento_contratista": "",
                        "obligaciones_especificas": "",
                        "nombre_supervisor": "",
                        "metodo_extraccion": "error",
                        "error_ia": "",
                        "texto_extraido_chars": 0,
                        "error": str(e),
                    }
                )
    return results


def save_results_to_excel(data: List[Dict], output_path: Path) -> None:
    df = pd.DataFrame(data)
    preferred_columns = [
        "archivo",
        "tipo_documento_origen",
        "numero_contrato",
        "Tipo_contrato",
        "nombre_contratista",
        "numero_documento_contratista",
        "nombre_supervisor",
        "obligaciones_especificas",
        "metodo_extraccion",
        "texto_extraido_chars",
        "error_ia",
        "error",
    ]
    df = df[[c for c in preferred_columns if c in df.columns]]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="contratos")


def print_configuration(zip_path: Path, output_path: Path, use_ai: bool, api_key: str) -> None:
    print("=" * 80)
    print("EXTRACCIÓN DE CONTRATOS A EXCEL")
    print("=" * 80)
    print(f"ZIP de entrada : {zip_path}")
    print(f"Excel de salida: {output_path}")
    print(f"Usar IA        : {'Sí' if use_ai else 'No'}")
    print(f"OpenAI API Key : {'Detectada' if api_key else 'No detectada'}")
    print("=" * 80)


def prompt_zip_path(default_path: str = "") -> Path:
    while True:
        print("\nIngresa la ruta completa del archivo .zip que quieres procesar.")
        if default_path:
            print(f"Presiona Enter para usar la ruta por defecto: {default_path}")
        user_input = input("Ruta del ZIP: ").strip().strip('"')
        selected = user_input or default_path
        if not selected:
            print("Debes escribir una ruta.")
            continue
        zip_path = Path(selected)
        if not zip_path.exists():
            print(f"No existe la ruta: {zip_path}")
            continue
        if not zip_path.is_file():
            print(f"La ruta no corresponde a un archivo: {zip_path}")
            continue
        if zip_path.suffix.lower() != ".zip":
            print(f"El archivo no es .zip: {zip_path.name}")
            continue
        return zip_path


def prompt_output_path(default_path: str) -> Path:
    while True:
        print("\nIngresa la ruta completa del Excel de salida.")
        print(f"Presiona Enter para usar la ruta por defecto: {default_path}")
        user_input = input("Ruta del Excel de salida (.xlsx): ").strip().strip('"')
        selected = user_input or default_path
        if not selected:
            print("Debes escribir una ruta de salida.")
            continue
        output_path = Path(selected)
        if output_path.suffix.lower() != ".xlsx":
            print("La ruta de salida debe terminar en .xlsx")
            continue
        return output_path


def ask_yes_no(message: str, default: bool = True) -> bool:
    hint = "S/n" if default else "s/N"
    answer = input(f"{message} [{hint}]: ").strip().lower()
    if not answer:
        return default
    return answer in {"s", "si", "sí", "y", "yes"}


def main() -> None:
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        zip_path = Path(sys.argv[1].strip().strip('"'))
    else:
        zip_path = prompt_zip_path(INPUT_ZIP_PATH)

    if len(sys.argv) >= 3 and sys.argv[2].strip():
        output_path = Path(sys.argv[2].strip().strip('"'))
    else:
        output_path = prompt_output_path(OUTPUT_EXCEL_PATH)

    use_ai = ask_yes_no("¿Quieres usar IA para fortalecer la extracción?", default=USE_AI)

    api_key = OPENAI_API_KEY.strip()
    client = get_openai_client(api_key) if use_ai and api_key else None
    if use_ai and not client:
        print("\nAdvertencia: no se detectó OPENAI_API_KEY válida. Se procesará solo con reglas.")
        use_ai = False

    print_configuration(zip_path, output_path, use_ai, api_key)
    data = process_zip(zip_path=zip_path, client=client, use_ai=use_ai)
    save_results_to_excel(data, output_path)
    print("\nProceso terminado correctamente.")
    print(f"Excel guardado en: {output_path}")


if __name__ == "__main__":
    main()
