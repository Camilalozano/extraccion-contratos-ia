import io
import json
import os
import re
import sys
import unicodedata
import zipfile
from pathlib import Path
from typing import Dict, List, Optional

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


# =========================
# Configuración editable
# =========================
# Ruta sugerida del ZIP de entrada (se puede cambiar al iniciar el script).
INPUT_ZIP_PATH = r"Seleccion_500_Archivos_contratos.zip"

# Ruta por defecto del Excel de salida. Puedes cambiarla.
OUTPUT_EXCEL_PATH = r"D:\Users\Usuario\Documents\ArchivosExtraidos.xlsx"

# Activa o desactiva el uso de IA.
USE_AI = True

# Toma la API Key desde variable de entorno OPENAI_API_KEY.
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

# =========================
# Utilidades generales
# =========================
TARGET_FIELDS = [
    "numero_contrato",
    "Tipo_contrato",
    "nombre_contratista",
    "numero_documento_contratista",
    "obligaciones_especificas",
]


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



def limpiar_texto_para_llm(text: str) -> str:
    if not text:
        return ""

    t = unicodedata.normalize("NFKC", text)
    t = t.replace("\u00A0", " ")
    t = t.replace("\u200B", "")
    t = t.replace("\u200E", "")
    t = t.replace("\u200F", "")

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
    text = limpiar_texto_para_llm(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()



def search_first(patterns: List[str], text: str, flags: int = re.IGNORECASE | re.DOTALL) -> Optional[re.Match]:
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return match
    return None



def cut_text(text: str, limit: int = 15000) -> str:
    return text[:limit] if len(text) > limit else text


# =========================
# Extracción de texto PDF
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
# Reglas de extracción
# =========================
def extract_contract_number(text: str, filename: str = "") -> str:
    patterns = [
        r"CONTRATO\s+DE\s+[A-ZÁÉÍÓÚÑ\s]+?\s+No\.?\s*([A-Z0-9\-_/]+)",
        r"CONTRATO\s+No\.?\s*([A-Z0-9\-_/]+)",
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
        r"CONTRATO\s+DE\s+([A-ZÁÉÍÓÚÑ\s]+?)\s+No\.?,",
        r"CONTRATO\s+DE\s+([A-ZÁÉÍÓÚÑ\s]+?)\s+No\.?",
        r"presente\s+contrato\s+de\s+([A-ZÁÉÍÓÚÑ\s]+?)(?:\s+el\s+cual|\s+que\s+se\s+regir[áa]|\s+de\s+conformidad)",
    ]
    m = search_first(patterns, text)
    if not m:
        return ""
    value = re.sub(r"\s+", " ", m.group(1)).strip(" ,.;:\n\t")
    return value.upper()



def extract_contractor_name(text: str) -> str:
    patterns = [
        r"y\s+por\s+la\s+otra,\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+?),\s*mayor\s+de\s+edad,\s*identificad[oa]",
        r"AGENCIA\s+ATENEA\s+y\s+por\s+la\s+otra,\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+?),\s*mayor\s+de\s+edad",
        r"celebrado\s+entre.*?y\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+?)\.\s*ANG[ÉE]LICA",
        r"celebrado\s+entre.*?y\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+?)\.\s*[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+,\s*identificad",
        r"quien\s+en\s+adelante\s+se\s+denominar[áa]\s+EL\s+CONTRATISTA.*?por\s+la\s+otra,\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+?),",
    ]
    m = search_first(patterns, text)
    if m:
        contractor_name = re.sub(r"\s+", " ", m.group(1)).strip(" ,.;:\n\t")
        contractor_name = re.sub(r"^la\s+tecnolog[íi]a\s+y\s+", "", contractor_name, flags=re.IGNORECASE)
        return contractor_name
    return ""



def extract_contractor_document(text: str, contractor_name: str = "") -> str:
    def normalize_for_search(s: str) -> str:
        s = s.replace("\xa0", " ")
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    text_norm = normalize_for_search(text)
    num_pattern = r"([0-9OIl][0-9OIl\.\,\-\s]{5,}[0-9OIl])"

    id_patterns = [
        rf"c[ée]dula\s+de\s+ciudadan[íi]a\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
        rf"\bC\.?\s*C\.?\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
        rf"\bc[ée]dula\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
        rf"identificad[oa]\s*(?:\(a\))?\s+con\s+(?:la\s+)?c[ée]dula\s+de\s+ciudadan[íi]a\s*(?:No\.?|N°|Nº|#|:)?\s*{num_pattern}",
    ]

    if contractor_name:
        contractor_name_esc = re.escape(normalize_for_search(contractor_name))
        m_name = re.search(contractor_name_esc, text_norm, re.IGNORECASE)
        if m_name:
            start = max(0, m_name.start() - 80)
            end = min(len(text_norm), m_name.end() + 300)
            window = text_norm[start:end]
            for pattern in id_patterns:
                m = re.search(pattern, window, re.IGNORECASE)
                if m:
                    return only_digits(m.group(1))

    m_other = re.search(
        r"por\s+la\s+otra,?(.*?)(?:actuando\s+en\s+nombre\s+propio|qu[ií]en\s+declara|EL\s+CONTRATISTA)",
        text_norm,
        re.IGNORECASE | re.DOTALL,
    )
    if m_other:
        block = m_other.group(1)
        for pattern in id_patterns:
            m = re.search(pattern, block, re.IGNORECASE)
            if m:
                return only_digits(m.group(1))

    found = []
    for pattern in id_patterns:
        for m in re.finditer(pattern, text_norm, re.IGNORECASE):
            value = only_digits(m.group(1))
            if 6 <= len(value) <= 12:
                found.append(value)

    return found[-1] if found else ""



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



# =========================
# Extracción con IA
# =========================
def get_openai_client(api_key: Optional[str]) -> Optional[OpenAI]:
    if not api_key:
        return None
    try:
        return OpenAI(api_key=api_key)
    except Exception:
        return None



def build_focus_context(text: str, obligaciones_regla: str) -> str:
    parts = []
    header = cut_text(text, 12000)
    parts.append("=== INICIO DEL CONTRATO / CONTEXTO GENERAL ===\n" + header)

    if obligaciones_regla:
        parts.append("=== BLOQUE DETECTADO DE OBLIGACIONES ESPECÍFICAS ===\n" + cut_text(obligaciones_regla, 8000))

    return "\n\n".join(parts)



def extract_contract_fields_raw(client: OpenAI, text: str, filename: str, rule_candidates: dict) -> str:
    focus_text = build_focus_context(
        text=text,
        obligaciones_regla=rule_candidates.get("obligaciones_especificas", ""),
    )

    prompt = f"""
A partir del siguiente contrato en español, extrae SOLO estos campos y devuelve SOLO JSON válido:
- numero_contrato
- Tipo_contrato
- nombre_contratista
- numero_documento_contratista
- obligaciones_especificas

Reglas obligatorias:
1. No inventes datos.
2. Si un campo no aparece claramente, devuelve "".
3. numero_documento_contratista debe quedar SOLO con dígitos.
4. obligaciones_especificas debe devolver el bloque textual de obligaciones específicas del contratista; conserva el contenido sustancial del contrato y no resumas.
5. Devuelve únicamente JSON válido, sin explicación, sin markdown.

Archivo: {filename}

Candidatos por reglas:
{json.dumps(rule_candidates, ensure_ascii=False, indent=2)}

Texto del contrato:
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



def normalize_ai_result(data: dict) -> dict:
    data = data or {}
    normalized = {field: normalize_nullable_text(data.get(field, "")) for field in TARGET_FIELDS}
    normalized["numero_documento_contratista"] = only_digits(normalized.get("numero_documento_contratista"))
    normalized["numero_contrato"] = normalized.get("numero_contrato", "").replace("–", "-")
    return normalized



def merge_results(rule_result: dict, ai_result: Optional[dict]) -> dict:
    result = {field: normalize_nullable_text(rule_result.get(field, "")) for field in TARGET_FIELDS}
    result["numero_documento_contratista"] = only_digits(result.get("numero_documento_contratista"))

    if not ai_result:
        return result

    ai_result = normalize_ai_result(ai_result)

    for field in ["numero_contrato", "Tipo_contrato", "nombre_contratista", "obligaciones_especificas"]:
        if ai_result.get(field):
            result[field] = ai_result[field]

    if ai_result.get("numero_documento_contratista"):
        result["numero_documento_contratista"] = ai_result["numero_documento_contratista"]

    return result


# =========================
# Procesamiento principal
# =========================
def process_single_pdf(pdf_bytes: bytes, filename: str, client: Optional[OpenAI] = None, use_ai: bool = True) -> Dict:
    raw_text = extract_text_from_pdf_bytes(pdf_bytes)
    text = normalize_text(raw_text)

    contractor_name = extract_contractor_name(text)
    rule_result = {
        "numero_contrato": extract_contract_number(text, filename),
        "Tipo_contrato": extract_contract_type(text),
        "nombre_contratista": contractor_name,
        "numero_documento_contratista": extract_contractor_document(text, contractor_name),
        "obligaciones_especificas": extract_obligaciones_especificas(raw_text or text),
    }

    ai_result = None
    metodo = "reglas"
    error_ia = ""

    if use_ai and client is not None:
        try:
            raw_ai = extract_contract_fields_raw(client, text=text, filename=filename, rule_candidates=rule_result)
            ai_result = normalize_ai_result(safe_json_loads(raw_ai))
            metodo = "hibrido_reglas_ia"
        except Exception as e:
            error_ia = str(e)
            metodo = "reglas_con_fallo_ia"

    final_result = merge_results(rule_result, ai_result)
    final_result.update(
        {
            "archivo": Path(filename).name,
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

        for i, name in enumerate(pdf_files, start=1):
            print(f"Procesando {i}/{total}: {Path(name).name}")
            pdf_bytes = zf.read(name)
            try:
                results.append(process_single_pdf(pdf_bytes, name, client=client, use_ai=use_ai))
            except Exception as e:
                results.append(
                    {
                        "archivo": Path(name).name,
                        "numero_contrato": "",
                        "Tipo_contrato": "",
                        "nombre_contratista": "",
                        "numero_documento_contratista": "",
                        "obligaciones_especificas": "",
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
        "numero_contrato",
        "Tipo_contrato",
        "nombre_contratista",
        "numero_documento_contratista",
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
    print(f"Usar IA       : {'Sí' if use_ai else 'No'}")
    print(f"OpenAI API Key: {'Detectada' if api_key else 'No detectada'}")
    print("=" * 80)


def prompt_input_zip_path(default_zip_path: Path) -> Path:
    print("\nIngresa la ruta completa del archivo .zip a procesar.")
    print(f"Presiona ENTER para usar la ruta sugerida: {default_zip_path}")

    while True:
        user_value = input("Ruta del .zip: ").strip().strip('"').strip("'")
        zip_path = Path(user_value) if user_value else default_zip_path

        if zip_path.suffix.lower() != ".zip":
            print("La ruta ingresada no corresponde a un archivo .zip. Intenta de nuevo.")
            continue

        if not zip_path.exists():
            print(f"No se encontró el archivo: {zip_path}. Intenta de nuevo.")
            continue

        return zip_path



def main() -> None:
    zip_path = Path(INPUT_ZIP_PATH)
    output_path = Path(OUTPUT_EXCEL_PATH)
    use_ai = USE_AI

    # Permite sobreescribir por línea de comandos:
    # python Extract.py [ruta_zip] [ruta_salida_excel]
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        zip_path = Path(sys.argv[1])
    else:
        zip_path = prompt_input_zip_path(zip_path)
    if len(sys.argv) >= 3 and sys.argv[2].strip():
        output_path = Path(sys.argv[2])

    if not zip_path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo ZIP: {zip_path}\n"
            f"Verifica la ruta o ejecútalo así:\n"
            f"python Extract.py \"ruta_del_zip\" \"ruta_de_salida.xlsx\""
        )

    api_key = OPENAI_API_KEY.strip()
    client = get_openai_client(api_key) if use_ai and api_key else None
    if use_ai and not client:
        print("Advertencia: no se detectó OPENAI_API_KEY válida. Se procesará solo con reglas.")
        use_ai = False

    print_configuration(zip_path, output_path, use_ai, api_key)

    data = process_zip(zip_path=zip_path, client=client, use_ai=use_ai)
    save_results_to_excel(data, output_path)

    print("\nProceso terminado correctamente.")
    print(f"Excel guardado en: {output_path}")


if __name__ == "__main__":
    main()
