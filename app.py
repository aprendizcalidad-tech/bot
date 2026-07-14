from __future__ import annotations

import base64
import html
import json
import os
import re
import random
import shutil
import subprocess
import tempfile
import time
from io import BytesIO
from pathlib import Path
from typing import Literal
from xml.sax.saxutils import escape

import pdfplumber
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from google import genai
from google.genai import types
from pydantic import BaseModel, ConfigDict, Field
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

ALLOWED_EXTENSIONS = ["docx", "pdf", "txt"]
MODEL_DEFAULT = "gemini-flash-latest"
MAX_FILE_SIZE_MB = 20
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
MAX_GUIDES = 8
MAX_TEXT_CHARS_PER_FILE = 80_000
MAX_TOTAL_TEXT_CHARS = 300_000
MAX_TOTAL_UPLOAD_MB = 80
MAX_TOTAL_UPLOAD_BYTES = MAX_TOTAL_UPLOAD_MB * 1024 * 1024

STATUS_LABELS = {
    "completo": "Completo",
    "parcial": "Parcial",
    "faltante": "Faltante",
}

MIME_TYPES = {
    "pdf": "application/pdf",
    "docx": (
        "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.document"
    ),
    "txt": "text/plain",
}


# =========================================================
# MODELOS DE DATOS PARA RESPUESTAS ESTRUCTURADAS DE GEMINI
# =========================================================


class MissingQuestion(BaseModel):
    model_config = ConfigDict(extra="forbid")

    question: str = Field(
        description=(
            "Pregunta específica que debe responder el usuario. "
            "No debe ser genérica ni repetir otra pregunta."
        )
    )
    why_needed: str = Field(
        description="Razón breve por la cual el dato es necesario."
    )
    required: bool = Field(
        description="Indica si la respuesta es indispensable para generar el documento."
    )


class SectionAssessment(BaseModel):
    model_config = ConfigDict(extra="forbid")

    order: int = Field(ge=1)
    title: str
    guide_instruction: str = Field(
        description=(
            "Instrucción normativa interpretada desde la guía. "
            "No es contenido final del documento."
        )
    )
    criteria: list[str]
    required: bool
    status: Literal["completo", "parcial", "faltante"]
    evidence: list[str] = Field(
        description=(
            "Fragmentos o referencias breves del documento de origen "
            "que sustentan la sección."
        )
    )
    draft_content: str = Field(
        description=(
            "Borrador construido únicamente con información sustentada. "
            "Debe quedar vacío cuando no sea posible redactar sin inventar."
        )
    )
    output_format: Literal["parrafo", "lista", "tabla", "mixto"]
    missing_questions: list[MissingQuestion]


class GuideAnalysis(BaseModel):
    model_config = ConfigDict(extra="forbid")

    detected_process: str
    selected_guide_index: int = Field(ge=0)
    selected_guide_name: str
    supporting_guide_indices: list[int]
    selection_reason: str
    proposed_document_title: str
    general_requirements: list[str]
    sections: list[SectionAssessment]
    warnings: list[str]


class UserAnswer(BaseModel):
    model_config = ConfigDict(extra="forbid")

    section_title: str
    question: str
    answer: str


class FinalTable(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str
    headers: list[str]
    rows: list[list[str]]


class FinalSection(BaseModel):
    model_config = ConfigDict(extra="forbid")

    order: int = Field(ge=1)
    title: str
    paragraphs: list[str]
    bullets: list[str]
    tables: list[FinalTable]
    source_basis: list[str]


class ValidationItem(BaseModel):
    model_config = ConfigDict(extra="forbid")

    section_title: str
    criterion: str
    status: Literal["cumple", "parcial", "no_aplica"]
    note: str


class FinalDocument(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str
    subtitle: str
    introductory_note: str
    sections: list[FinalSection]
    validation: list[ValidationItem]
    warnings: list[str]


class AppError(Exception):
    """Error controlado que puede mostrarse al usuario."""


# =========================================================
# UTILIDADES DE CONFIGURACIÓN Y ARCHIVOS
# =========================================================


def read_secret(name: str, default: str = "") -> str:
    """Lee un secreto desde Streamlit o una variable de entorno."""

    try:
        value = st.secrets.get(name, default)
        return str(value) if value is not None else default
    except Exception:
        return os.getenv(name, default)


def safe_filename(filename: str) -> str:
    """Evita utilizar rutas incluidas en el nombre del archivo."""

    return Path(filename).name


def file_extension(filename: str) -> str:
    """Devuelve la extensión sin punto y en minúsculas."""

    return Path(filename).suffix.lower().lstrip(".")


def format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.1f} MB"


def validate_uploaded_file(filename: str, content: bytes) -> None:
    """Valida extensión, contenido y tamaño."""

    extension = file_extension(filename)

    if extension not in ALLOWED_EXTENSIONS:
        raise AppError(
            f"{filename}: el formato .{extension or 'desconocido'} no está permitido."
        )

    if not content:
        raise AppError(f"{filename}: el archivo está vacío.")

    if len(content) > MAX_FILE_SIZE_BYTES:
        raise AppError(
            f"{filename}: supera el límite de {MAX_FILE_SIZE_MB} MB."
        )


def decode_txt(content: bytes) -> str:
    """Lee TXT probando codificaciones frecuentes."""

    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue

    raise AppError("No fue posible determinar la codificación del archivo TXT.")


def extract_docx_text(content: bytes) -> str:
    """Extrae títulos, párrafos, tablas, encabezados y pies de página."""

    try:
        document = Document(BytesIO(content))
    except Exception as error:
        raise AppError("El archivo DOCX está dañado o no es válido.") from error

    parts: list[str] = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        style_name = (paragraph.style.name or "").lower()
        if "heading" in style_name or "título" in style_name or "titulo" in style_name:
            parts.append(f"[TÍTULO] {text}")
        else:
            parts.append(text)

    for table_number, table in enumerate(document.tables, start=1):
        parts.append(f"[TABLA {table_number}]")
        for row in table.rows:
            cells = [
                re.sub(r"\s+", " ", cell.text).strip()
                for cell in row.cells
            ]
            if any(cells):
                parts.append(" | ".join(cells))

    seen_headers: set[str] = set()
    seen_footers: set[str] = set()

    for section in document.sections:
        header_text = "\n".join(
            paragraph.text.strip()
            for paragraph in section.header.paragraphs
            if paragraph.text.strip()
        )
        if header_text and header_text not in seen_headers:
            parts.append(f"[ENCABEZADO]\n{header_text}")
            seen_headers.add(header_text)

        footer_text = "\n".join(
            paragraph.text.strip()
            for paragraph in section.footer.paragraphs
            if paragraph.text.strip()
        )
        if footer_text and footer_text not in seen_footers:
            parts.append(f"[PIE DE PÁGINA]\n{footer_text}")
            seen_footers.add(footer_text)

    return "\n".join(parts).strip()


def extract_pdf_text(content: bytes) -> tuple[str, int]:
    """Extrae texto local de PDF para vista previa; Gemini recibe el PDF nativo."""

    pages: list[str] = []

    try:
        with pdfplumber.open(BytesIO(content)) as pdf:
            page_count = len(pdf.pages)
            for page_number, page in enumerate(pdf.pages, start=1):
                text = (page.extract_text() or "").strip()
                if text:
                    pages.append(f"[PÁGINA {page_number}]\n{text}")
    except Exception as error:
        raise AppError("El PDF está dañado, protegido o no puede leerse.") from error

    return "\n\n".join(pages).strip(), page_count


def extract_text(filename: str, content: bytes) -> tuple[str, int | None]:
    """Extrae texto para DOCX, PDF o TXT."""

    extension = file_extension(filename)

    if extension == "txt":
        return decode_txt(content).strip(), None

    if extension == "docx":
        return extract_docx_text(content), None

    if extension == "pdf":
        try:
            return extract_pdf_text(content)
        except AppError:
            # La extracción local es solo para vista previa. Gemini recibirá
            # el PDF completo de forma nativa y hará el análisis principal.
            return "", None

    raise AppError(f"Formato no soportado: .{extension}")


def build_file_record(uploaded_file, role: str, index: int) -> dict:
    """Convierte UploadedFile en una estructura persistente en sesión."""

    name = safe_filename(uploaded_file.name)
    content = uploaded_file.getvalue()
    validate_uploaded_file(name, content)
    text, page_count = extract_text(name, content)

    warnings: list[str] = []
    extension = file_extension(name)

    if extension == "pdf" and len(text) < 30:
        warnings.append(
            "El PDF tiene poco texto extraíble localmente. "
            "Gemini lo analizará visualmente como PDF nativo."
        )

    if extension in {"docx", "txt"} and not text:
        warnings.append("No se encontró texto legible en el archivo.")

    return {
        "role": role,
        "index": index,
        "name": name,
        "extension": extension,
        "mime_type": MIME_TYPES[extension],
        "content": content,
        "size_bytes": len(content),
        "text": text,
        "page_count": page_count,
        "warnings": warnings,
    }


def clipped_text(text: str, limit: int = MAX_TEXT_CHARS_PER_FILE) -> str:
    """Limita texto para evitar solicitudes excesivas."""

    if len(text) <= limit:
        return text

    return (
        text[:limit]
        + "\n\n[CONTENIDO RECORTADO POR LÍMITE TÉCNICO DE LA APLICACIÓN]"
    )


def pdf_document_part(record: dict) -> types.Part:
    """Construye una entrada PDF nativa para generateContent."""

    return types.Part.from_bytes(
        data=record["content"],
        mime_type="application/pdf",
    )


def append_record_to_gemini_input(
    parts: list[types.Part],
    record: dict,
    label: str,
) -> None:
    """Agrega un archivo al input multimodal como objetos Part válidos."""

    parts.append(
        types.Part.from_text(
            text=(
                f"\n===== INICIO {label}: {record['name']} =====\n"
                f"Rol del archivo: {record['role']}\n"
            )
        )
    )

    if record["extension"] == "pdf":
        parts.append(pdf_document_part(record))
    else:
        extracted_text = clipped_text(record["text"]).strip()
        if not extracted_text:
            extracted_text = "[ARCHIVO SIN TEXTO EXTRAÍBLE]"
        parts.append(types.Part.from_text(text=extracted_text))

    parts.append(
        types.Part.from_text(
            text=f"\n===== FIN {label}: {record['name']} =====\n"
        )
    )


# =========================================================
# PROMPTS Y LLAMADAS A GEMINI
# =========================================================


def analysis_prompt(guide_names: list[str], source_name: str) -> str:
    catalog = "\n".join(
        f"- Índice {index}: {name}"
        for index, name in enumerate(guide_names)
    )

    return f"""
Actúa como analista documental, profesional de calidad y redactor técnico.

OBJETIVO DEL ANÁLISIS
Comparar las GUÍAS NORMATIVAS cargadas por el usuario con el DOCUMENTO DE
ORIGEN, identificar la guía principal aplicable y convertir sus instrucciones
en una estructura de documento que pueda completarse sin inventar datos.

CATÁLOGO DE GUÍAS
{catalog}

DOCUMENTO DE ORIGEN
{source_name}

INTERPRETACIÓN OBLIGATORIA DE LAS GUÍAS
Las guías NO son plantillas con marcadores y NO contienen necesariamente los
datos del caso. Son instrucciones sobre lo que debe tener cada sección.

Ejemplo:
Si una guía dice: "OBJETIVO: describir la intención del documento en términos
del qué y el para qué, de forma clara", debes interpretar que la sección
Objetivo debe explicar QUÉ se realizará y PARA QUÉ se realizará. No debes copiar
esa instrucción como si fuera el objetivo final.

REGLAS DE ANÁLISIS
1. Selecciona una guía principal utilizando su índice real del catálogo.
2. Puedes indicar guías complementarias si agregan reglas compatibles.
3. Extrae las secciones, el orden, las instrucciones y los criterios exigidos.
4. Compara cada criterio con el documento de origen.
5. Distingue entre:
   - Información explícita: aparece directamente en el origen.
   - Información derivable: puede redactarse razonablemente usando hechos del
     origen, sin agregar hechos nuevos.
   - Información ausente: no puede inferirse con seguridad.
6. Nunca inventes nombres, cargos, fechas, cifras, sedes, periodos, resultados,
   responsables, conclusiones ni decisiones.
7. Para información ausente, formula preguntas específicas. Ejemplo correcto:
   "¿Cuál fue el periodo evaluado?". Ejemplo incorrecto: "Complete el alcance".
8. Reduce preguntas duplicadas: una respuesta puede servir para varias secciones.
9. draft_content debe ser una propuesta redactada, no una copia de la instrucción.
10. Si una sección requiere tabla, lista o combinación de formatos, indícalo en
    output_format.
11. evidence debe contener referencias breves al documento de origen, no a la guía.
12. Conserva literalmente identificaciones, radicados, fechas y valores encontrados.
13. Ordena sections según la guía seleccionada.
14. selected_guide_name debe coincidir exactamente con el archivo del catálogo.
15. supporting_guide_indices solo puede contener índices válidos y distintos del
    índice principal.

ESTADOS
- completo: existe información suficiente para redactar y cumplir los criterios.
- parcial: existe un borrador sustentado, pero faltan uno o más datos críticos.
- faltante: no hay información suficiente para redactar sin inventar.

Devuelve únicamente el objeto estructurado solicitado.
""".strip()


def generation_prompt(
    analysis: GuideAnalysis,
    answers: list[UserAnswer],
) -> str:
    return f"""
Actúa como redactor técnico y auditor de cumplimiento documental.

Debes generar el DOCUMENTO FINAL utilizando:
1. Las guías normativas adjuntas.
2. El documento de origen adjunto.
3. El análisis estructurado previo.
4. Las respuestas suministradas por el usuario.

ANÁLISIS PREVIO
{analysis.model_dump_json(indent=2)}

RESPUESTAS DEL USUARIO
{json.dumps([answer.model_dump() for answer in answers], ensure_ascii=False, indent=2)}

REGLAS OBLIGATORIAS
1. La guía contiene instrucciones, no contenido factual. No copies las
   instrucciones como texto final.
2. Redacta cada sección siguiendo exactamente su finalidad y criterios.
3. Usa únicamente hechos del documento de origen y respuestas del usuario.
4. Puedes mejorar redacción, cohesión y orden, pero no inventar hechos.
5. Conserva literalmente nombres, documentos, radicados, fechas, cifras y valores.
6. Cuando una instrucción pida un objetivo en términos del qué y el para qué,
   sintetiza ambos elementos en una redacción clara y profesional.
7. Respeta el orden de las secciones establecido en el análisis.
8. Usa párrafos, viñetas o tablas según lo exigido por la guía.
9. No agregues una sección de "datos faltantes" al documento final.
10. Si todavía existe una limitación no resoluble, regístrala en warnings sin
    fabricar contenido.
11. En source_basis indica de forma breve qué información sustenta cada sección.
12. validation debe revisar cada criterio de la guía y marcar cumple, parcial o
    no_aplica con una justificación breve.
13. No incluyas el informe de validación dentro del contenido normal del documento.
14. Devuelve únicamente el objeto estructurado solicitado.
""".strip()


def translate_gemini_error(error: Exception) -> AppError:
    """Convierte errores técnicos frecuentes en mensajes comprensibles."""

    message = str(error)
    upper_message = message.upper()

    if "API_KEY_INVALID" in upper_message or "INVALID API KEY" in upper_message:
        return AppError(
            "La clave de Gemini no es válida. Revisa .streamlit/secrets.toml."
        )

    if "RESOURCE_EXHAUSTED" in upper_message or "429" in upper_message:
        return AppError(
            "Se alcanzó el límite temporal o gratuito de Gemini. "
            "Espera unos minutos y vuelve a intentar."
        )

    if "NOT_FOUND" in upper_message or "404" in upper_message:
        return AppError(
            "El modelo configurado no está disponible. "
            "Revisa GEMINI_MODEL en secrets.toml."
        )

    if "PERMISSION_DENIED" in upper_message or "403" in upper_message:
        return AppError(
            "La clave no tiene permiso para usar Gemini o el servicio no está "
            "habilitado para ese proyecto."
        )

    if "INVALID_ARGUMENT" in upper_message or "400" in upper_message:
        return AppError(
            "Gemini rechazó la solicitud por un argumento inválido. "
            f"Detalle técnico: {message[:500]}"
        )

    return AppError(f"Gemini no pudo completar el proceso. Detalle: {message[:700]}")


def _inline_and_clean_json_schema(schema_model: type[BaseModel]) -> dict:
    """Genera un JSON Schema compatible con Gemini.

    Pydantic agrega ``additionalProperties: false`` cuando se usa
    ``ConfigDict(extra="forbid")``. Algunas versiones del endpoint antiguo
    ``response_schema`` no aceptan esa propiedad. Esta función expande las
    referencias internas y elimina metadatos incompatibles antes de enviar el
    esquema mediante ``response_json_schema``.
    """

    schema_model.model_rebuild()
    raw_schema = schema_model.model_json_schema()
    definitions = raw_schema.get("$defs", {})

    def resolve(node):
        if isinstance(node, list):
            return [resolve(item) for item in node]

        if not isinstance(node, dict):
            return node

        if "$ref" in node:
            reference = str(node["$ref"])
            prefix = "#/$defs/"
            if reference.startswith(prefix):
                definition_name = reference[len(prefix):]
                definition = definitions.get(definition_name)
                if definition is None:
                    raise AppError(
                        f"No fue posible resolver el esquema interno {reference}."
                    )

                merged = dict(definition)
                merged.update(
                    {
                        key: value
                        for key, value in node.items()
                        if key != "$ref"
                    }
                )
                return resolve(merged)

        cleaned: dict = {}
        for key, value in node.items():
            if key in {
                "$defs",
                "$schema",
                "additionalProperties",
                "examples",
                "default",
            }:
                continue
            cleaned[key] = resolve(value)

        return cleaned

    compatible_schema = resolve(raw_schema)

    if not isinstance(compatible_schema, dict):
        raise AppError("No fue posible construir el esquema de respuesta.")

    return compatible_schema


def _clean_json_text(text: str) -> str:
    """Elimina cercas Markdown que algunos modelos agregan alrededor del JSON."""

    cleaned = text.strip()

    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    return cleaned.strip()


def _validate_structured_output(
    schema_model: type[BaseModel],
    response,
) -> BaseModel:
    """Valida la respuesta de Gemini con el modelo Pydantic solicitado."""

    parsed = getattr(response, "parsed", None)

    if isinstance(parsed, schema_model):
        return parsed

    if isinstance(parsed, dict):
        return schema_model.model_validate(parsed)

    output_text = getattr(response, "text", None)
    if not output_text:
        raise AppError(
            "Gemini respondió, pero no devolvió contenido estructurado."
        )

    try:
        return schema_model.model_validate_json(_clean_json_text(output_text))
    except Exception as error:
        raise AppError(
            "Gemini devolvió una respuesta JSON que no cumple la estructura "
            f"requerida. Detalle: {str(error)[:500]}"
        ) from error


def _schema_fallback_instruction(
    schema_model: type[BaseModel],
    compatible_schema: dict,
) -> types.Part:
    """Crea una instrucción de respaldo si el endpoint rechaza el esquema."""

    schema_text = json.dumps(
        compatible_schema,
        ensure_ascii=False,
        separators=(",", ":"),
    )

    return types.Part.from_text(
        text=(
            "\nFORMATO JSON OBLIGATORIO DE RESPUESTA\n"
            f"Devuelve un único objeto JSON válido para {schema_model.__name__}. "
            "No uses Markdown, comentarios ni texto antes o después del JSON. "
            f"Respeta este esquema: {schema_text}"
        )
    )


def _is_schema_endpoint_error(error: Exception) -> bool:
    """Detecta errores de compatibilidad del esquema del endpoint."""

    message = str(error).upper()
    return (
        "RESPONSE_SCHEMA" in message
        or "RESPONSE_JSON_SCHEMA" in message
        or "ADDITIONAL_PROPERTIES" in message
        or "ADDITIONALPROPERTIES" in message
        or "UNKNOWN NAME" in message
        or "CANNOT FIND FIELD" in message
    )


def _normalize_model_name(name: str | None) -> str:
    """Normaliza los nombres devueltos por models.list()."""

    normalized = (name or "").strip()
    if normalized.startswith("models/"):
        normalized = normalized[len("models/"):]
    return normalized


def _is_suitable_generation_model(model_info) -> bool:
    """Filtra modelos de texto/multimodales aptos para este bot documental."""

    name = _normalize_model_name(getattr(model_info, "name", ""))
    actions = set(getattr(model_info, "supported_actions", None) or [])

    if not name or "generateContent" not in actions:
        return False

    lowered = name.lower()
    excluded_terms = (
        "embedding",
        "image",
        "live",
        "tts",
        "audio",
        "robotics",
        "computer-use",
        "veo",
        "imagen",
        "lyria",
    )
    return lowered.startswith("gemini") and not any(
        term in lowered for term in excluded_terms
    )


def _model_priority(name: str) -> tuple[int, str]:
    """Ordena primero modelos Flash de menor latencia y mayor disponibilidad."""

    lowered = name.lower()
    preferred = {
        "gemini-flash-lite-latest": 0,
        "gemini-flash-latest": 1,
        "gemini-3.1-flash-lite-preview": 2,
        "gemini-3-flash-preview": 3,
        "gemini-2.0-flash-lite": 4,
        "gemini-2.0-flash-lite-001": 5,
        "gemini-2.0-flash": 6,
        "gemini-2.0-flash-001": 7,
    }
    if lowered in preferred:
        return preferred[lowered], lowered
    if "flash-lite" in lowered and "latest" in lowered:
        return 8, lowered
    if "flash" in lowered and "latest" in lowered:
        return 9, lowered
    if "flash-lite" in lowered:
        return 10, lowered
    if "flash" in lowered and "preview" not in lowered:
        return 11, lowered
    if "flash" in lowered:
        return 12, lowered
    if "pro" in lowered and "preview" not in lowered:
        return 13, lowered
    return 20, lowered


def _discover_candidate_models(client, configured_model: str) -> list[str]:
    """Obtiene modelos aptos y los ordena para evitar quedarse en uno saturado."""

    configured = _normalize_model_name(configured_model)
    preferred_defaults = [
        configured,
        "gemini-flash-lite-latest",
        "gemini-flash-latest",
        "gemini-3.1-flash-lite-preview",
        "gemini-3-flash-preview",
        "gemini-2.0-flash-lite",
        "gemini-2.0-flash",
    ]

    discovered: list[str] = []
    try:
        for model_info in client.models.list():
            if _is_suitable_generation_model(model_info):
                name = _normalize_model_name(getattr(model_info, "name", ""))
                if name and name not in discovered:
                    discovered.append(name)
    except Exception:
        discovered = []

    ordered: list[str] = []

    def add(candidate: str) -> None:
        candidate = _normalize_model_name(candidate)
        if candidate and candidate not in ordered:
            ordered.append(candidate)

    if discovered:
        available = set(discovered)
        for candidate in preferred_defaults:
            if candidate in available:
                add(candidate)
        for candidate in sorted(discovered, key=_model_priority):
            add(candidate)
    else:
        for candidate in preferred_defaults:
            add(candidate)

    # Evita esperas excesivas si el listado contiene demasiados modelos.
    return ordered[:8]


def _is_retryable_gemini_error(error: Exception) -> bool:
    """Detecta errores temporales de capacidad, cuota o red."""

    message = str(error).upper()
    retryable_terms = (
        "503",
        "UNAVAILABLE",
        "HIGH DEMAND",
        "429",
        "RESOURCE_EXHAUSTED",
        "DEADLINE_EXCEEDED",
        "504",
        "GATEWAY_TIMEOUT",
        "CONNECTION RESET",
        "CONNECTION ERROR",
        "TIMED OUT",
        "TIMEOUT",
    )
    return any(term in message for term in retryable_terms)


def _generate_with_retries(
    client,
    model: str,
    contents: list[types.Content],
    config: types.GenerateContentConfig,
    max_attempts: int = 3,
):
    """Reintenta errores temporales con espera exponencial y variación aleatoria."""

    last_error: Exception | None = None

    for attempt in range(max_attempts):
        try:
            return client.models.generate_content(
                model=model,
                contents=contents,
                config=config,
            )
        except Exception as error:
            last_error = error
            if not _is_retryable_gemini_error(error) or attempt == max_attempts - 1:
                raise

            delay = min(8.0, 1.5 * (2 ** attempt)) + random.uniform(0.1, 0.7)
            time.sleep(delay)

    if last_error is not None:
        raise last_error
    raise AppError("Gemini no respondió y no entregó un error identificable.")


def call_gemini_structured(
    api_key: str,
    model: str,
    input_parts: list[types.Part],
    schema_model: type[BaseModel],
) -> BaseModel:
    """Ejecuta Gemini con reintentos y cambio automático de modelo."""

    if not api_key:
        raise AppError(
            "No se encontró GEMINI_API_KEY. Configúrala en "
            ".streamlit/secrets.toml."
        )

    client = genai.Client(api_key=api_key)
    compatible_schema = _inline_and_clean_json_schema(schema_model)
    candidate_models = _discover_candidate_models(client, model)

    if not candidate_models:
        raise AppError(
            "La clave fue aceptada, pero la API no devolvió modelos compatibles "
            "con generateContent para este proyecto."
        )

    attempted_models: list[str] = []
    model_errors: list[str] = []

    try:
        for candidate_model in candidate_models:
            attempted_models.append(candidate_model)
            user_content = types.Content(role="user", parts=list(input_parts))

            try:
                response = _generate_with_retries(
                    client=client,
                    model=candidate_model,
                    contents=[user_content],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_json_schema=compatible_schema,
                        temperature=0.1,
                    ),
                )
                st.session_state["active_gemini_model"] = candidate_model
                return _validate_structured_output(schema_model, response)

            except AppError:
                raise
            except Exception as error:
                error_message = str(error)
                upper_message = error_message.upper()
                model_errors.append(f"{candidate_model}: {error_message[:180]}")

                # Modelo retirado/no habilitado o temporalmente saturado:
                # se prueba el siguiente modelo compatible.
                if (
                    "404" in upper_message
                    or "NOT_FOUND" in upper_message
                    or _is_retryable_gemini_error(error)
                ):
                    continue

                if not _is_schema_endpoint_error(error):
                    raise translate_gemini_error(error) from error

                fallback_parts = list(input_parts)
                fallback_parts.append(
                    _schema_fallback_instruction(schema_model, compatible_schema)
                )
                fallback_content = types.Content(role="user", parts=fallback_parts)

                try:
                    fallback_response = _generate_with_retries(
                        client=client,
                        model=candidate_model,
                        contents=[fallback_content],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.1,
                        ),
                    )
                    st.session_state["active_gemini_model"] = candidate_model
                    return _validate_structured_output(
                        schema_model,
                        fallback_response,
                    )
                except AppError:
                    raise
                except Exception as fallback_error:
                    fallback_message = str(fallback_error)
                    fallback_upper = fallback_message.upper()
                    model_errors.append(
                        f"{candidate_model} (respaldo): {fallback_message[:180]}"
                    )
                    if (
                        "404" in fallback_upper
                        or "NOT_FOUND" in fallback_upper
                        or _is_retryable_gemini_error(fallback_error)
                    ):
                        continue
                    raise translate_gemini_error(fallback_error) from fallback_error

        attempts = ", ".join(attempted_models) or "ninguno"
        detail = " | ".join(model_errors[-4:])
        raise AppError(
            "Gemini está temporalmente saturado o no respondió con los modelos "
            f"disponibles. Modelos probados: {attempts}. "
            "Espera uno o dos minutos y vuelve a presionar Analizar. "
            f"Detalle: {detail[:550]}"
        )

    finally:
        try:
            client.close()
        except Exception:
            pass


def analyze_guides_and_source(
    api_key: str,
    model: str,
    guides: list[dict],
    source: dict,
) -> GuideAnalysis:
    """Selecciona la guía y evalúa el documento de origen."""

    parts: list[types.Part] = []

    for index, guide in enumerate(guides):
        append_record_to_gemini_input(parts, guide, f"GUÍA {index}")

    append_record_to_gemini_input(parts, source, "DOCUMENTO DE ORIGEN")
    parts.append(
        types.Part.from_text(
            text=analysis_prompt(
                [guide["name"] for guide in guides],
                source["name"],
            )
        )
    )

    result = call_gemini_structured(
        api_key=api_key,
        model=model,
        input_parts=parts,
        schema_model=GuideAnalysis,
    )

    assert isinstance(result, GuideAnalysis)

    if not 0 <= result.selected_guide_index < len(guides):
        raise AppError("Gemini seleccionó un índice de guía inválido.")

    result.selected_guide_name = guides[result.selected_guide_index]["name"]

    valid_supporting: list[int] = []
    for index in result.supporting_guide_indices:
        if (
            0 <= index < len(guides)
            and index != result.selected_guide_index
            and index not in valid_supporting
        ):
            valid_supporting.append(index)
    result.supporting_guide_indices = valid_supporting

    result.sections = sorted(result.sections, key=lambda section: section.order)

    if not result.sections:
        raise AppError(
            "Gemini no identificó secciones en la guía seleccionada. "
            "Verifica que la guía describa el contenido esperado."
        )

    return result


def generate_final_with_gemini(
    api_key: str,
    model: str,
    guides: list[dict],
    source: dict,
    analysis: GuideAnalysis,
    answers: list[UserAnswer],
) -> FinalDocument:
    """Redacta y valida el documento final."""

    relevant_indices = [analysis.selected_guide_index]
    relevant_indices.extend(analysis.supporting_guide_indices)

    parts: list[types.Part] = []

    for index in relevant_indices:
        append_record_to_gemini_input(parts, guides[index], f"GUÍA APLICABLE {index}")

    append_record_to_gemini_input(parts, source, "DOCUMENTO DE ORIGEN")
    parts.append(
        types.Part.from_text(
            text=generation_prompt(analysis, answers)
        )
    )

    result = call_gemini_structured(
        api_key=api_key,
        model=model,
        input_parts=parts,
        schema_model=FinalDocument,
    )

    assert isinstance(result, FinalDocument)
    result.sections = sorted(result.sections, key=lambda section: section.order)

    if not result.sections:
        raise AppError("Gemini no generó las secciones del documento final.")

    return result


# =========================================================
# GENERACIÓN DE DOCX Y PDF
# =========================================================


def configure_docx(document: Document, preserve_guide_layout: bool = False) -> None:
    """Configura tipografía base y, si no hay guía visual, márgenes estándar."""

    if not preserve_guide_layout:
        for section in document.sections:
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)

    normal_style = document.styles["Normal"]
    if not normal_style.font.name:
        normal_style.font.name = "Arial"
    if not normal_style.font.size:
        normal_style.font.size = Pt(11)

    for style_name in ("Title", "Heading 1", "Heading 2"):
        if style_name in document.styles:
            style = document.styles[style_name]
            if not style.font.name:
                style.font.name = "Arial"


def clear_docx_body_keep_layout(document: Document) -> None:
    """
    Elimina el contenido corporal de una guía DOCX, pero conserva sus secciones,
    encabezados, pies de página, logos, márgenes y configuración de página.

    La propiedad sectPr debe permanecer porque enlaza encabezados y pies.
    """

    body = document._element.body
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def load_visual_guide_document(style_guide: dict | None) -> tuple[Document, bool]:
    """Crea el documento base usando la guía DOCX seleccionada cuando sea posible."""

    if style_guide and style_guide.get("extension") == "docx":
        try:
            document = Document(BytesIO(style_guide["content"]))
            clear_docx_body_keep_layout(document)
            return document, True
        except Exception:
            # Si la guía está dañada o usa una estructura no compatible, se usa
            # un documento limpio para no bloquear la generación del contenido.
            pass

    return Document(), False


def add_docx_table(document: Document, table_data: FinalTable) -> None:
    if table_data.title.strip():
        title_paragraph = document.add_paragraph()
        title_run = title_paragraph.add_run(table_data.title.strip())
        title_run.bold = True

    column_count = max(
        len(table_data.headers),
        max((len(row) for row in table_data.rows), default=0),
    )

    if column_count == 0:
        return

    table = document.add_table(rows=1, cols=column_count)
    table.style = "Table Grid"

    headers = list(table_data.headers)
    headers.extend([""] * (column_count - len(headers)))

    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = header
        for run in cell.paragraphs[0].runs:
            run.bold = True

    for source_row in table_data.rows:
        row_values = list(source_row)
        row_values.extend([""] * (column_count - len(row_values)))
        cells = table.add_row().cells
        for index, value in enumerate(row_values):
            cells[index].text = value

    document.add_paragraph("")


def create_docx(
    final_document: FinalDocument,
    style_guide: dict | None = None,
) -> bytes:
    """
    Crea el Word final. Si la guía seleccionada es DOCX, reutiliza ese archivo
    como base visual para conservar encabezados, pies, logos, márgenes y estilos.
    """

    document, inherited_layout = load_visual_guide_document(style_guide)
    configure_docx(document, preserve_guide_layout=inherited_layout)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(final_document.title.strip())
    title_run.bold = True
    title_run.font.name = "Arial"
    title_run.font.size = Pt(16)

    if final_document.subtitle.strip():
        subtitle = document.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle_run = subtitle.add_run(final_document.subtitle.strip())
        subtitle_run.italic = True
        subtitle_run.font.name = "Arial"
        subtitle_run.font.size = Pt(11)

    document.add_paragraph("")

    if final_document.introductory_note.strip():
        document.add_paragraph(final_document.introductory_note.strip())

    for section in final_document.sections:
        document.add_heading(section.title.strip(), level=1)

        for paragraph_text in section.paragraphs:
            if paragraph_text.strip():
                document.add_paragraph(paragraph_text.strip())

        for bullet in section.bullets:
            if bullet.strip():
                document.add_paragraph(bullet.strip(), style="List Bullet")

        for table_data in section.tables:
            add_docx_table(document, table_data)

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def convert_docx_to_pdf(docx_content: bytes) -> bytes | None:
    """Convierte el DOCX a PDF con LibreOffice para conservar encabezado y pie."""

    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice:
        return None

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            docx_path = temp_path / "documento_generado.docx"
            pdf_path = temp_path / "documento_generado.pdf"
            docx_path.write_bytes(docx_content)

            result = subprocess.run(
                [
                    libreoffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(temp_path),
                    str(docx_path),
                ],
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )

            if result.returncode == 0 and pdf_path.exists():
                pdf_bytes = pdf_path.read_bytes()
                if pdf_bytes.startswith(b"%PDF"):
                    return pdf_bytes
    except Exception:
        return None

    return None


def pdf_paragraph(text: str, style: ParagraphStyle) -> Paragraph:
    safe_text = escape(text).replace("\n", "<br/>")
    return Paragraph(safe_text, style)


def create_pdf(final_document: FinalDocument) -> bytes:
    """Crea una versión PDF estructurada del documento final."""

    output = BytesIO()
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocumentTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        spaceAfter=10,
    )
    subtitle_style = ParagraphStyle(
        "DocumentSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=10,
        leading=14,
        alignment=TA_CENTER,
        spaceAfter=16,
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        spaceBefore=12,
        spaceAfter=7,
    )
    body_style = ParagraphStyle(
        "BodyTextCustom",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        spaceAfter=7,
    )
    bullet_style = ParagraphStyle(
        "BulletCustom",
        parent=body_style,
        leftIndent=14,
        firstLineIndent=-8,
        bulletIndent=4,
    )
    table_cell_style = ParagraphStyle(
        "TableCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )
    table_header_style = ParagraphStyle(
        "TableHeader",
        parent=table_cell_style,
        fontName="Helvetica-Bold",
    )

    story: list = [
        pdf_paragraph(final_document.title.strip(), title_style),
    ]

    if final_document.subtitle.strip():
        story.append(pdf_paragraph(final_document.subtitle.strip(), subtitle_style))

    if final_document.introductory_note.strip():
        story.append(
            pdf_paragraph(final_document.introductory_note.strip(), body_style)
        )
        story.append(Spacer(1, 6))

    available_width = A4[0] - 5 * cm

    for section in final_document.sections:
        section_story: list = [
            pdf_paragraph(section.title.strip(), heading_style),
        ]

        for paragraph_text in section.paragraphs:
            if paragraph_text.strip():
                section_story.append(pdf_paragraph(paragraph_text.strip(), body_style))

        for bullet in section.bullets:
            if bullet.strip():
                section_story.append(
                    Paragraph(
                        f"• {escape(bullet.strip())}",
                        bullet_style,
                    )
                )

        story.append(KeepTogether(section_story[:2]))
        story.extend(section_story[2:])

        for table_data in section.tables:
            if table_data.title.strip():
                story.append(
                    pdf_paragraph(table_data.title.strip(), body_style)
                )

            column_count = max(
                len(table_data.headers),
                max((len(row) for row in table_data.rows), default=0),
            )
            if column_count == 0:
                continue

            headers = list(table_data.headers)
            headers.extend([""] * (column_count - len(headers)))
            rows = [
                headers,
                *[
                    list(row) + [""] * (column_count - len(row))
                    for row in table_data.rows
                ],
            ]

            formatted_rows: list[list[Paragraph]] = []
            for row_index, row in enumerate(rows):
                cell_style = table_header_style if row_index == 0 else table_cell_style
                formatted_rows.append(
                    [pdf_paragraph(str(value), cell_style) for value in row]
                )

            col_widths = [available_width / column_count] * column_count
            table = Table(
                formatted_rows,
                colWidths=col_widths,
                repeatRows=1,
                hAlign="LEFT",
            )
            table.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 5),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            story.append(table)
            story.append(Spacer(1, 10))

    pdf = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=2.5 * cm,
        leftMargin=2.5 * cm,
        topMargin=2.5 * cm,
        bottomMargin=2.5 * cm,
        title=final_document.title,
        author="Generador inteligente de documentos",
    )
    pdf.build(story)
    return output.getvalue()


def output_filename(title: str, suffix: str) -> str:
    """Genera un nombre de archivo seguro."""

    normalized = re.sub(r"[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_-]+", "_", title).strip("_")
    normalized = normalized[:80] or "documento_generado"
    return f"{normalized}.{suffix}"


def save_generated_file(filename: str, content: bytes) -> Path:
    """Guarda una copia local como respaldo de descarga en Codespaces."""

    output_dir = Path("output")
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / Path(filename).name
    path.write_bytes(content)
    return path


def direct_download_link(
    label: str,
    data: bytes,
    filename: str,
    mime_type: str,
) -> str:
    """Crea un enlace de descarga embebido que no depende del endpoint temporal de Streamlit."""

    encoded = base64.b64encode(data).decode("ascii")
    safe_filename = html.escape(Path(filename).name, quote=True)
    safe_label = html.escape(label)

    return (
        '<a download="' + safe_filename + '" '
        'href="data:' + mime_type + ';base64,' + encoded + '" '
        'style="display:block;text-align:center;padding:0.65rem 1rem;'
        'border-radius:0.5rem;border:1px solid rgba(250,250,250,0.25);'
        'text-decoration:none;font-weight:600;color:inherit;">'
        + safe_label + '</a>'
    )


# =========================================================
# ESTADO Y COMPONENTES DE INTERFAZ
# =========================================================


def initialize_state() -> None:
    defaults = {
        "guide_records": None,
        "source_record": None,
        "analysis": None,
        "final_document": None,
        "docx_output": None,
        "pdf_output": None,
        "answers": None,
        "upload_key": 0,
        "active_gemini_model": None,
    }

    for key, default in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default


def reset_workflow() -> None:
    st.session_state.guide_records = None
    st.session_state.source_record = None
    st.session_state.analysis = None
    st.session_state.final_document = None
    st.session_state.docx_output = None
    st.session_state.pdf_output = None
    st.session_state.answers = None
    st.session_state.active_gemini_model = None
    st.session_state.upload_key += 1


def count_required_questions(analysis: GuideAnalysis) -> int:
    return sum(
        1
        for section in analysis.sections
        for question in section.missing_questions
        if question.required
    )


def show_file_record(record: dict, title: str, key: str) -> None:
    """Muestra información y vista previa de un archivo."""

    with st.expander(title):
        columns = st.columns(4)
        columns[0].metric("Formato", record["extension"].upper())
        columns[1].metric("Tamaño", format_file_size(record["size_bytes"]))
        columns[2].metric("Caracteres", len(record["text"]))
        columns[3].metric(
            "Páginas",
            record["page_count"] if record["page_count"] is not None else "N/A",
        )

        for warning in record["warnings"]:
            st.warning(warning)

        preview = record["text"][:6_000]
        st.text_area(
            "Vista previa local",
            value=preview or "No hay texto local para mostrar.",
            height=220,
            disabled=True,
            key=f"preview_{key}",
        )


def show_analysis(analysis: GuideAnalysis, guides: list[dict]) -> None:
    st.subheader("Resultado del análisis")

    columns = st.columns(4)
    columns[0].metric("Proceso", analysis.detected_process)
    columns[1].metric("Guía principal", analysis.selected_guide_name)
    columns[2].metric("Secciones", len(analysis.sections))
    columns[3].metric("Preguntas obligatorias", count_required_questions(analysis))

    st.info(analysis.selection_reason)

    if analysis.supporting_guide_indices:
        supporting_names = [
            guides[index]["name"]
            for index in analysis.supporting_guide_indices
        ]
        st.write("**Guías complementarias:** " + ", ".join(supporting_names))

    if analysis.general_requirements:
        with st.expander("Requisitos generales identificados"):
            for requirement in analysis.general_requirements:
                st.write(f"- {requirement}")

    for warning in analysis.warnings:
        st.warning(warning)

    st.markdown("### Evaluación por sección")

    for section_index, section in enumerate(analysis.sections):
        icon = {
            "completo": "✅",
            "parcial": "🟠",
            "faltante": "🔴",
        }[section.status]

        with st.expander(
            f"{icon} {section.order}. {section.title} — "
            f"{STATUS_LABELS[section.status]}"
        ):
            st.write("**Instrucción interpretada de la guía**")
            st.write(section.guide_instruction)

            st.write("**Criterios**")
            for criterion in section.criteria:
                st.write(f"- {criterion}")

            if section.evidence:
                st.write("**Sustento encontrado en el documento de origen**")
                for evidence in section.evidence:
                    st.write(f"- {evidence}")

            if section.draft_content.strip():
                st.write("**Borrador sustentado**")
                st.write(section.draft_content)

            if section.missing_questions:
                st.write("**Información que debe completarse**")
                for question in section.missing_questions:
                    required = "Obligatoria" if question.required else "Opcional"
                    st.write(f"- **{question.question}** ({required})")
                    st.caption(question.why_needed)

            st.caption(f"Formato previsto: {section.output_format}")


def collect_answers_form(analysis: GuideAnalysis) -> list[UserAnswer] | None:
    """Muestra preguntas dinámicas y devuelve respuestas al enviar el formulario."""

    all_questions = [
        (section, question_index, question)
        for section in analysis.sections
        for question_index, question in enumerate(section.missing_questions)
    ]

    if not all_questions:
        st.success(
            "La información disponible es suficiente. "
            "Puedes generar directamente el documento final."
        )

    answers_by_key: dict[str, str] = {}

    with st.form("missing_information_form", clear_on_submit=False):
        if all_questions:
            st.markdown("### Completar información faltante")
            st.caption(
                "Responde únicamente con información verificable. "
                "Los campos marcados con * son obligatorios."
            )

        for item_index, (section, question_index, question) in enumerate(all_questions):
            label = question.question + (" *" if question.required else "")
            help_text = f"Sección: {section.title}. {question.why_needed}"
            key = f"answer_{section.order}_{question_index}_{item_index}"

            answers_by_key[key] = st.text_area(
                label,
                height=90,
                help=help_text,
                key=key,
            )

        submitted = st.form_submit_button(
            "Generar y validar documento final",
            type="primary",
            width="stretch",
        )

    if not submitted:
        return None

    missing_required: list[str] = []
    answers: list[UserAnswer] = []

    for item_index, (section, question_index, question) in enumerate(all_questions):
        key = f"answer_{section.order}_{question_index}_{item_index}"
        answer_text = answers_by_key.get(key, "").strip()

        if question.required and not answer_text:
            missing_required.append(question.question)

        if answer_text:
            answers.append(
                UserAnswer(
                    section_title=section.title,
                    question=question.question,
                    answer=answer_text,
                )
            )

    if missing_required:
        st.error(
            "Debes responder los siguientes datos obligatorios:\n\n- "
            + "\n- ".join(missing_required)
        )
        return None

    return answers


def show_final_document(final_document: FinalDocument) -> None:
    st.subheader("Documento final generado")
    st.markdown(f"## {final_document.title}")

    if final_document.subtitle.strip():
        st.caption(final_document.subtitle)

    if final_document.introductory_note.strip():
        st.write(final_document.introductory_note)

    for section in final_document.sections:
        with st.expander(f"{section.order}. {section.title}", expanded=False):
            for paragraph_text in section.paragraphs:
                st.write(paragraph_text)

            for bullet in section.bullets:
                st.write(f"- {bullet}")

            for table_data in section.tables:
                if table_data.title.strip():
                    st.write(f"**{table_data.title}**")

                if table_data.headers and table_data.rows:
                    normalized_rows = []
                    for row in table_data.rows:
                        padded = list(row) + [""] * (
                            len(table_data.headers) - len(row)
                        )
                        normalized_rows.append(padded[: len(table_data.headers)])

                    st.dataframe(
                        {
                            header: [row[index] for row in normalized_rows]
                            for index, header in enumerate(table_data.headers)
                        },
                        width="stretch",
                        hide_index=True,
                    )

            if section.source_basis:
                st.caption("Sustento: " + " | ".join(section.source_basis))

    st.markdown("### Validación contra la guía")

    compliance_counts = {"cumple": 0, "parcial": 0, "no_aplica": 0}
    for item in final_document.validation:
        compliance_counts[item.status] += 1

    columns = st.columns(3)
    columns[0].metric("Cumple", compliance_counts["cumple"])
    columns[1].metric("Parcial", compliance_counts["parcial"])
    columns[2].metric("No aplica", compliance_counts["no_aplica"])

    with st.expander("Ver validación detallada"):
        for item in final_document.validation:
            icon = {
                "cumple": "✅",
                "parcial": "🟠",
                "no_aplica": "⚪",
            }[item.status]
            st.write(
                f"{icon} **{item.section_title}** — {item.criterion}: {item.note}"
            )

    for warning in final_document.warnings:
        st.warning(warning)


# =========================================================
# DISEÑO VISUAL DE LA INTERFAZ
# =========================================================


def inject_app_styles() -> None:
    """Aplica un diseño más moderno, legible y amigable."""

    st.markdown(
        """
        <style>
        :root {
            --brand: #ff4b4b;
            --brand-dark: #d9363e;
            --panel: rgba(255,255,255,0.055);
            --line: rgba(255,255,255,0.12);
            --muted: rgba(255,255,255,0.72);
        }

        .block-container {
            max-width: 1180px;
            padding-top: 2rem;
            padding-bottom: 4rem;
        }

        [data-testid="stSidebar"] {
            border-right: 1px solid var(--line);
        }

        .app-hero {
            padding: 1.6rem 1.8rem;
            border: 1px solid rgba(255,75,75,0.30);
            border-radius: 22px;
            background:
                radial-gradient(circle at 90% 20%, rgba(255,75,75,0.22), transparent 34%),
                linear-gradient(135deg, rgba(255,75,75,0.12), rgba(255,255,255,0.035));
            box-shadow: 0 16px 42px rgba(0,0,0,0.18);
            margin-bottom: 1.4rem;
        }

        .app-hero h1 {
            margin: 0 0 .45rem 0;
            font-size: clamp(2rem, 4vw, 3.2rem);
            line-height: 1.05;
        }

        .app-hero p {
            margin: 0;
            max-width: 830px;
            color: var(--muted);
            font-size: 1.04rem;
        }

        .step-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: .75rem;
            margin: .5rem 0 1.5rem;
        }

        .step-card {
            padding: .9rem 1rem;
            border-radius: 15px;
            border: 1px solid var(--line);
            background: var(--panel);
            min-height: 82px;
        }

        .step-card.active {
            border-color: rgba(255,75,75,0.68);
            background: rgba(255,75,75,0.13);
        }

        .step-number {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 25px;
            height: 25px;
            border-radius: 50%;
            background: rgba(255,255,255,.10);
            font-size: .8rem;
            font-weight: 700;
            margin-bottom: .35rem;
        }

        .step-card.active .step-number { background: var(--brand); color: white; }
        .step-card strong { display: block; font-size: .92rem; }
        .step-card span { color: var(--muted); font-size: .78rem; }

        [data-testid="stFileUploader"] {
            background: var(--panel);
            border: 1px dashed rgba(255,255,255,.22);
            border-radius: 16px;
            padding: .55rem .7rem;
        }

        [data-testid="stMetric"] {
            border: 1px solid var(--line);
            background: var(--panel);
            padding: .8rem 1rem;
            border-radius: 14px;
        }

        .friendly-note {
            border-left: 4px solid var(--brand);
            padding: .8rem 1rem;
            border-radius: 0 12px 12px 0;
            background: rgba(255,75,75,.08);
            margin: .75rem 0 1rem;
        }

        .sidebar-brand {
            padding: 1rem;
            border-radius: 16px;
            background: linear-gradient(135deg, rgba(255,75,75,.20), rgba(255,255,255,.05));
            border: 1px solid rgba(255,75,75,.25);
            margin-bottom: 1rem;
        }

        .sidebar-brand strong { font-size: 1.05rem; }
        .sidebar-brand small { color: var(--muted); }

        .stButton > button, .stDownloadButton > button {
            border-radius: 12px;
            min-height: 46px;
            font-weight: 650;
        }

        @media (max-width: 850px) {
            .step-grid { grid-template-columns: 1fr 1fr; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero() -> None:
    st.markdown(
        """
        <section class="app-hero">
          <h1>Generador documental inteligente</h1>
          <p>
            Carga las guías institucionales y el documento de origen. El sistema
            interpreta las reglas, identifica información faltante y genera un
            Word y un PDF listos para revisión.
          </p>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_workflow_steps(current_step: int) -> None:
    steps = [
        (1, "Cargar archivos", "Guías y documento de origen"),
        (2, "Analizar", "Comparación con Gemini"),
        (3, "Completar", "Solo datos críticos faltantes"),
        (4, "Generar", "Word y PDF institucionales"),
    ]
    cards = []
    for number, title, subtitle in steps:
        active = " active" if number <= current_step else ""
        cards.append(
            f'<div class="step-card{active}"><div class="step-number">{number}</div>'
            f'<strong>{title}</strong><span>{subtitle}</span></div>'
        )
    st.markdown('<div class="step-grid">' + ''.join(cards) + '</div>', unsafe_allow_html=True)


def current_workflow_step() -> int:
    if st.session_state.get("final_document") is not None:
        return 4
    if st.session_state.get("analysis") is not None:
        return 3
    if st.session_state.get("guide_records"):
        return 2
    return 1

# =========================================================
# APLICACIÓN PRINCIPAL
# =========================================================


def main() -> None:
    st.set_page_config(
        page_title="Generador documental con guías",
        page_icon="📄",
        layout="wide",
    )

    initialize_state()
    inject_app_styles()

    api_key = read_secret("GEMINI_API_KEY")
    configured_model = read_secret("GEMINI_MODEL", MODEL_DEFAULT)
    if configured_model in {"gemini-2.5-flash", "gemini-2.5-flash-lite"}:
        configured_model = MODEL_DEFAULT

    render_hero()
    render_workflow_steps(current_workflow_step())

    with st.expander("¿Cómo funciona este proceso?", expanded=False):
        st.markdown(
            """
            1. **Carga las guías institucionales** y el documento con la información real.
            2. **Gemini interpreta las reglas** de cada sección y selecciona la guía aplicable.
            3. El bot **pregunta únicamente lo indispensable** cuando falta información.
            4. Genera un **Word con el encabezado y pie de la guía DOCX seleccionada** y su PDF.
            """
        )

    with st.sidebar:
        st.markdown(
            """<div class="sidebar-brand"><strong>Centro de control</strong><br>
            <small>Estado del servicio y opciones del proceso</small></div>""",
            unsafe_allow_html=True,
        )
        st.subheader("Configuración")

        if api_key:
            st.success("Clave de Gemini detectada")
        else:
            st.error("Falta GEMINI_API_KEY")

        model = st.text_input(
            "Modelo Gemini preferido",
            value=configured_model,
            help=(
                "El bot verifica los modelos habilitados para tu clave y cambia "
                "automáticamente a uno disponible si este nombre ya fue retirado."
            ),
        ).strip()

        active_model = st.session_state.get("active_gemini_model")
        if active_model:
            st.caption(f"Modelo usado en el último proceso: {active_model}")
        else:
            st.caption("Selección automática de modelo habilitada.")

        st.caption(
            f"Hasta {MAX_GUIDES} guías y {MAX_FILE_SIZE_MB} MB por archivo."
        )

        if st.button("Reiniciar proceso", width="stretch"):
            reset_workflow()
            st.rerun()

    st.subheader("📁 1. Cargar guías e información de origen")

    st.markdown(
        """<div class="friendly-note"><strong>Consejo:</strong> para conservar exactamente
        el encabezado, pie de página, logos y márgenes, carga al menos una guía en formato
        <strong>DOCX</strong>.</div>""",
        unsafe_allow_html=True,
    )

    guide_uploads = st.file_uploader(
        "Guías del proceso",
        type=ALLOWED_EXTENSIONS,
        accept_multiple_files=True,
        key=f"guides_{st.session_state.upload_key}",
        help=(
            "Carga las guías que describen los apartados, criterios y reglas "
            "del documento final."
        ),
    )

    source_upload = st.file_uploader(
        "Documento de origen",
        type=ALLOWED_EXTENSIONS,
        accept_multiple_files=False,
        key=f"source_{st.session_state.upload_key}",
        help="Archivo que contiene los datos reales del proceso o caso.",
    )

    analyze_clicked = st.button(
        "Analizar guías y documento de origen",
        type="primary",
        width="stretch",
        disabled=not api_key,
    )

    if analyze_clicked:
        try:
            if not guide_uploads:
                raise AppError("Debes cargar al menos una guía.")

            if len(guide_uploads) > MAX_GUIDES:
                raise AppError(f"Solo se permiten hasta {MAX_GUIDES} guías.")

            if source_upload is None:
                raise AppError("Debes cargar el documento de origen.")

            total_upload_bytes = sum(
                len(upload.getvalue()) for upload in guide_uploads
            ) + len(source_upload.getvalue())
            if total_upload_bytes > MAX_TOTAL_UPLOAD_BYTES:
                raise AppError(
                    f"El conjunto de archivos supera {MAX_TOTAL_UPLOAD_MB} MB. "
                    "Reduce la cantidad o el tamaño de los documentos."
                )

            names = [safe_filename(upload.name).lower() for upload in guide_uploads]
            if len(names) != len(set(names)):
                raise AppError("Hay guías con nombres duplicados.")

            with st.spinner("Leyendo archivos..."):
                guide_records = [
                    build_file_record(upload, role="guía normativa", index=index)
                    for index, upload in enumerate(guide_uploads)
                ]
                source_record = build_file_record(
                    source_upload,
                    role="documento de origen",
                    index=0,
                )

                total_text = sum(len(record["text"]) for record in guide_records)
                total_text += len(source_record["text"])

                if total_text > MAX_TOTAL_TEXT_CHARS:
                    st.warning(
                        "Los archivos contienen mucho texto. La aplicación limitará "
                        "el contenido textual enviado por archivo, pero los PDF se "
                        "seguirán enviando completos de forma nativa."
                    )

            with st.spinner(
                "Gemini está interpretando las guías y comparando la información..."
            ):
                analysis = analyze_guides_and_source(
                    api_key=api_key,
                    model=model or MODEL_DEFAULT,
                    guides=guide_records,
                    source=source_record,
                )

            st.session_state.guide_records = guide_records
            st.session_state.source_record = source_record
            st.session_state.analysis = analysis
            st.session_state.final_document = None
            st.session_state.docx_output = None
            st.session_state.pdf_output = None
            st.session_state.answers = None

            st.success("Análisis completado correctamente.")

        except AppError as error:
            st.error(str(error))
        except Exception as error:
            st.error(f"Ocurrió un error inesperado: {error}")

    guide_records = st.session_state.guide_records
    source_record = st.session_state.source_record
    analysis = st.session_state.analysis

    if guide_records and source_record:
        st.divider()
        st.subheader("Archivos procesados")

        for index, guide in enumerate(guide_records):
            show_file_record(
                guide,
                title=f"Guía {index}: {guide['name']}",
                key=f"guide_{index}",
            )

        show_file_record(
            source_record,
            title=f"Documento de origen: {source_record['name']}",
            key="source",
        )

    if analysis and guide_records and source_record:
        st.divider()
        show_analysis(analysis, guide_records)

        st.divider()
        answers = collect_answers_form(analysis)

        if answers is not None:
            try:
                with st.spinner(
                    "Gemini está redactando y validando el documento final..."
                ):
                    final_document = generate_final_with_gemini(
                        api_key=api_key,
                        model=model or MODEL_DEFAULT,
                        guides=guide_records,
                        source=source_record,
                        analysis=analysis,
                        answers=answers,
                    )

                    selected_guide = guide_records[analysis.selected_guide_index]
                    docx_output = create_docx(
                        final_document,
                        style_guide=selected_guide,
                    )
                    pdf_output = convert_docx_to_pdf(docx_output)
                    if pdf_output is None:
                        pdf_output = create_pdf(final_document)

                st.session_state.answers = answers
                st.session_state.final_document = final_document
                st.session_state.docx_output = docx_output
                st.session_state.pdf_output = pdf_output

                if selected_guide.get("extension") == "docx":
                    st.success(
                        "Documento generado con el encabezado, pie de página, márgenes "
                        "y elementos visuales de la guía DOCX seleccionada."
                    )
                else:
                    st.success("Documento final generado y validado.")
                    st.info(
                        "La guía seleccionada no es DOCX. El contenido fue respetado, pero "
                        "para copiar exactamente encabezado y pie debes cargar la versión Word."
                    )

            except AppError as error:
                st.error(str(error))
            except Exception as error:
                st.error(f"No fue posible generar el documento: {error}")

    final_document = st.session_state.final_document

    if final_document:
        st.divider()
        show_final_document(final_document)

        st.markdown("### Descargar archivos")

        docx_output = st.session_state.get("docx_output")
        pdf_output = st.session_state.get("pdf_output")

        files_are_valid = True

        if not isinstance(docx_output, (bytes, bytearray)) or not docx_output:
            st.error(
                "El archivo Word no está disponible en memoria. "
                "Vuelve a generar el documento."
            )
            files_are_valid = False

        if not isinstance(pdf_output, (bytes, bytearray)) or not pdf_output:
            st.error(
                "El archivo PDF no está disponible en memoria. "
                "Vuelve a generar el documento."
            )
            files_are_valid = False

        if files_are_valid:
            # Convertimos explícitamente a bytes para evitar problemas con
            # bytearray/BytesIO y mantenemos claves estables para los widgets.
            docx_bytes = bytes(docx_output)
            pdf_bytes = bytes(pdf_output)

            # Validación mínima de firmas para detectar archivos corruptos antes
            # de ofrecerlos al navegador. DOCX es un contenedor ZIP (PK) y PDF
            # debe iniciar con %PDF.
            if not docx_bytes.startswith(b"PK"):
                st.error(
                    "El Word generado no tiene una estructura DOCX válida. "
                    "Vuelve a generar el documento."
                )
                files_are_valid = False

            if not pdf_bytes.startswith(b"%PDF"):
                st.error(
                    "El PDF generado no tiene una estructura PDF válida. "
                    "Vuelve a generar el documento."
                )
                files_are_valid = False

        if files_are_valid:
            docx_name = output_filename(final_document.title, "docx")
            pdf_name = output_filename(final_document.title, "pdf")

            # Copia local de respaldo. En Codespaces también puede descargarse
            # con clic derecho sobre el archivo en la carpeta output.
            docx_path = save_generated_file(docx_name, docx_bytes)
            pdf_path = save_generated_file(pdf_name, pdf_bytes)

            st.markdown("#### Descarga directa")
            st.caption(
                "Estos enlaces llevan el archivo embebido en la página y evitan "
                "el endpoint temporal que puede fallar en puertos privados de Codespaces."
            )

            direct_columns = st.columns(2)
            direct_columns[0].markdown(
                direct_download_link(
                    "⬇️ Descargar Word",
                    docx_bytes,
                    docx_name,
                    MIME_TYPES["docx"],
                ),
                unsafe_allow_html=True,
            )
            direct_columns[1].markdown(
                direct_download_link(
                    "⬇️ Descargar PDF",
                    pdf_bytes,
                    pdf_name,
                    MIME_TYPES["pdf"],
                ),
                unsafe_allow_html=True,
            )

            with st.expander("Descarga alternativa de Streamlit"):
                download_columns = st.columns(2)
                download_columns[0].download_button(
                    "Descargar Word con Streamlit",
                    data=docx_bytes,
                    file_name=docx_name,
                    mime=MIME_TYPES["docx"],
                    width="stretch",
                    key="download_final_docx",
                    on_click="ignore",
                )
                download_columns[1].download_button(
                    "Descargar PDF con Streamlit",
                    data=pdf_bytes,
                    file_name=pdf_name,
                    mime=MIME_TYPES["pdf"],
                    width="stretch",
                    key="download_final_pdf",
                    on_click="ignore",
                )

            st.caption(
                f"Word: {format_file_size(len(docx_bytes))} · "
                f"PDF: {format_file_size(len(pdf_bytes))}"
            )
            st.info(
                "También se guardaron copias en "
                f"`{docx_path.as_posix()}` y `{pdf_path.as_posix()}`. "
                "En Codespaces puedes abrir la carpeta `output`, hacer clic derecho "
                "sobre cada archivo y seleccionar Download."
            )

    st.divider()
    st.caption(
        "La aplicación utiliza las guías como reglas de construcción. No copia sus "
        "instrucciones como contenido factual y no debe inventar información que no "
        "esté sustentada en el documento de origen o en las respuestas del usuario."
    )


if __name__ == "__main__":
    main()