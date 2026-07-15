from __future__ import annotations

import base64
from copy import deepcopy
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
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
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
from reportlab.pdfgen import canvas as pdf_canvas
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

    category: str = Field(
        default="Información del proceso",
        description=(
            "Categoría breve para agrupar la pregunta, por ejemplo: "
            "Datos operativos, Responsables, Fechas y periodos o Control documental."
        ),
    )
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
    numbered_items: list[str] = Field(
        default_factory=list,
        description=(
            "Pasos o actividades que deben presentarse como lista numerada. "
            "No deben incluir viñetas ni el número dentro del texto."
        ),
    )
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


class AuditReport(BaseModel):
    model_config = ConfigDict(extra="forbid")

    validation: list[ValidationItem]
    warnings: list[str]
    editorial_summary: str = Field(
        description=(
            "Resumen breve de la calidad de redacción, coherencia, trazabilidad "
            "y cumplimiento del documento revisado."
        )
    )


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
Actúa como analista documental, profesional de calidad y redactor técnico institucional.

OBJETIVO
Comparar las GUÍAS cargadas por el usuario con el DOCUMENTO DE ORIGEN,
seleccionar la guía principal aplicable, interpretar sus instrucciones y
determinar qué contenido puede redactarse con evidencia verificable.

CATÁLOGO DE GUÍAS
{catalog}

DOCUMENTO DE ORIGEN
{source_name}

MÉTODO OBLIGATORIO EN CUATRO ETAPAS
ETAPA 1 — INTERPRETAR LA GUÍA
- Identifica el tipo de documento, las secciones, el orden, los criterios,
  las tablas, las listas y las reglas de presentación.
- Las guías son instrucciones de construcción; no copies sus explicaciones
  como contenido factual del documento final.

ETAPA 2 — EXTRAER EVIDENCIA
- Localiza hechos explícitos en el documento de origen.
- Distingue entre información explícita, información derivable sin agregar
  hechos nuevos e información ausente.
- Conserva literalmente nombres, cargos, radicados, fechas, cifras,
  plataformas y denominaciones oficiales.

ETAPA 3 — PROPONER REDACCIÓN
- Redacta draft_content con lenguaje institucional claro, directo y preciso.
- No inventes datos, responsables, decisiones, resultados, periodos ni
  controles.
- Evita repeticiones, muletillas, frases ambiguas y contenido genérico.

ETAPA 4 — IDENTIFICAR BRECHAS
- Formula únicamente preguntas críticas que no puedan resolverse con el
  origen.
- Cada pregunta debe ser específica, breve y responder a un solo dato.
- Agrupa cada pregunta en una categoría útil: Datos operativos,
  Responsables, Fechas y periodos, Control documental u otra equivalente.
- No preguntes por datos de encabezado que la guía ya deja como XX o
  pendiente de asignación, salvo que sean indispensables para el contenido.

INTERPRETACIÓN
Ejemplo: si la guía dice “OBJETIVO: describir la intención del documento en
términos del qué y el para qué”, debes redactar un objetivo que explique QUÉ
se hará y PARA QUÉ se hará. No copies esa instrucción.

REGLAS
1. Selecciona una guía principal usando su índice real.
2. Incluye guías complementarias solo si aportan reglas compatibles.
3. Ordena las secciones según la guía principal.
4. evidence debe referirse al documento de origen, no a la guía.
5. draft_content debe ser contenido final propuesto.
6. Evita preguntas duplicadas; una respuesta puede servir a varias secciones.
7. Para procedimientos secuenciales usa output_format “lista” o “mixto”.
8. selected_guide_name debe coincidir exactamente con el archivo del catálogo.
9. supporting_guide_indices solo puede contener índices válidos y distintos del principal.

ESTADOS
- completo: información suficiente para cumplir todos los criterios.
- parcial: existe un borrador sustentado, pero falta información crítica.
- faltante: no es posible redactar sin inventar.

Devuelve únicamente el objeto estructurado solicitado.
""".strip()


def generation_prompt(
    analysis: GuideAnalysis,
    answers: list[UserAnswer],
) -> str:
    return f"""
Actúa como redactor técnico institucional y auditor de cumplimiento documental.

Debes producir un DOCUMENTO FINAL profesional utilizando:
1. Las guías aplicables adjuntas.
2. El documento de origen.
3. El análisis estructurado previo.
4. Las respuestas verificables del usuario.

ANÁLISIS PREVIO
{analysis.model_dump_json(indent=2)}

RESPUESTAS DEL USUARIO
{json.dumps([answer.model_dump() for answer in answers], ensure_ascii=False, indent=2)}

PROCESO OBLIGATORIO
1. INTERPRETAR: confirma la finalidad y criterios de cada sección.
2. REDACTAR: construye el contenido con trazabilidad y lenguaje institucional.
3. NORMALIZAR: corrige ortografía, concordancia, puntuación, repeticiones,
   mayúsculas, fechas y denominaciones.
4. AUDITAR: evalúa cada criterio y registra el resultado en validation.

REGLAS DE REDACCIÓN
- Usa un estilo formal, preciso, coherente y eficiente.
- Prioriza oraciones claras y directas; evita párrafos inflados.
- Mantén una idea principal por oración cuando sea posible.
- Usa voz institucional o impersonal.
- Para procedimientos, usa verbos de acción y orden lógico.
- No inventes hechos, responsables, fechas, resultados, decisiones ni controles.
- Conserva literalmente datos sensibles y denominaciones oficiales.
- No copies las instrucciones de la guía como contenido factual.
- El objetivo debe integrar claramente el QUÉ y el PARA QUÉ cuando la guía lo exija.
- El alcance debe indicar inicio, límite y finalización solo cuando estén sustentados.
- Las definiciones deben presentarse en orden alfabético cuando la guía lo solicite.
- No incluyas una sección de datos faltantes en el documento final.

ESTRUCTURA DE LISTAS
- Usa numbered_items exclusivamente para pasos, actividades secuenciales,
  procedimientos, etapas o instrucciones ordenadas.
- Cada elemento de numbered_items debe ir sin “1.”, “2.” ni viñeta inicial.
- Usa bullets para responsabilidades, definiciones, condiciones, referencias
  y listados sin secuencia.
- No mezcles una viñeta con un número, por ejemplo “• 1.”.

TABLAS
- Usa tablas solo cuando la guía las solicite o mejoren claramente la lectura.
- En control de cambios conserva las columnas de la guía.
- Evita textos con espacios repetidos o saltos manuales innecesarios.

VALIDACIÓN
- validation debe revisar cada criterio de la guía.
- Marca cumple, parcial o no_aplica y explica brevemente.
- source_basis debe indicar el sustento de cada sección.
- Registra en warnings cualquier limitación que no pueda resolverse sin inventar.

Devuelve únicamente el objeto estructurado solicitado.
""".strip()


def audit_prompt(
    analysis: GuideAnalysis,
    final_document: FinalDocument,
) -> str:
    return f"""
Actúa como auditor documental independiente.

Compara el borrador final contra las instrucciones y criterios identificados
en la guía. Evalúa contenido, redacción y trazabilidad.

ANÁLISIS DE LA GUÍA
{analysis.model_dump_json(indent=2)}

BORRADOR A REVISAR
{final_document.model_dump_json(indent=2)}

CRITERIOS DE AUDITORÍA
1. Cada sección cumple su finalidad y criterios.
2. La redacción es formal, clara, precisa y sin repeticiones.
3. No existen hechos inventados ni afirmaciones sin sustento.
4. Nombres, cargos, fechas, cifras y sistemas son consistentes.
5. Los pasos están ordenados y no mezclan viñetas con numeración.
6. Las tablas son coherentes y no contienen espacios artificiales.
7. El documento mantiene trazabilidad con el origen y las respuestas.
8. Las limitaciones reales se registran como advertencias.

Devuelve validation para todos los criterios relevantes, warnings concretos
y un editorial_summary breve. No reescribas el documento.
""".strip()


def section_regeneration_prompt(
    analysis: GuideAnalysis,
    answers: list[UserAnswer],
    final_document: FinalDocument,
    section: FinalSection,
) -> str:
    return f"""
Actúa como redactor técnico institucional.

Regenera ÚNICAMENTE la sección indicada, manteniendo el orden y la finalidad
definidos por la guía. Usa exclusivamente hechos sustentados en los archivos,
el análisis y las respuestas del usuario.

ANÁLISIS
{analysis.model_dump_json(indent=2)}

RESPUESTAS
{json.dumps([answer.model_dump() for answer in answers], ensure_ascii=False, indent=2)}

CONTEXTO DEL DOCUMENTO
Título: {final_document.title}
Subtítulo: {final_document.subtitle}

SECCIÓN A REGENERAR
{section.model_dump_json(indent=2)}

REGLAS
- Mejora claridad, precisión, cohesión y eficiencia.
- No inventes información.
- Conserva el número de orden de la sección.
- Usa numbered_items para pasos secuenciales y bullets para listas no ordenadas.
- No incluyas números dentro del texto de numbered_items.
- Mantén las tablas solo si son necesarias o exigidas por la guía.
- Devuelve únicamente la estructura FinalSection.
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
    "gemini-3.1-flash-lite",
    "gemini-flash-latest",
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
    return ordered[:3]


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


def _normalized_question_key(question: str) -> str:
    """Normaliza una pregunta para detectar duplicados semánticos simples."""

    cleaned = re.sub(r"[¿?¡!.,;:()]+", " ", question.lower())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    stopwords = {
        "el", "la", "los", "las", "un", "una", "de", "del", "en",
        "para", "por", "que", "cuál", "cual", "qué", "como", "cómo",
    }
    tokens = [token for token in cleaned.split() if token not in stopwords]
    return " ".join(tokens)


def _deduplicate_missing_questions(analysis: GuideAnalysis) -> GuideAnalysis:
    """Elimina preguntas repetidas sin perder la primera sección que las requiere."""

    seen: set[str] = set()

    for section in analysis.sections:
        unique_questions: list[MissingQuestion] = []

        for question in section.missing_questions:
            key = _normalized_question_key(question.question)
            if not key or key in seen:
                continue

            seen.add(key)
            question.category = (
                question.category.strip() or "Información del proceso"
            )
            unique_questions.append(question)

        section.missing_questions = unique_questions

    return analysis


def _strip_list_prefix(text: str) -> str:
    """Elimina viñetas y numeración ya incorporadas en el texto."""

    cleaned = re.sub(r"^\s*[•●▪◦\-–—]\s*", "", text.strip())
    cleaned = re.sub(r"^\s*\(?\d+\)?[.)\-:]\s*", "", cleaned)
    return re.sub(r"\s+", " ", cleaned).strip()


def normalize_final_document_structure(
    final_document: FinalDocument,
) -> FinalDocument:
    """Normaliza listas, espacios, tablas y orden antes de generar."""

    normalized_sections: list[FinalSection] = []

    def explicit_number(value: str) -> int | None:
        match = re.match(
            r"^\s*[•●▪◦\-–—]?\s*\(?(\d+)\)?[.)\-:]\s+",
            value,
        )
        return int(match.group(1)) if match else None

    for section in sorted(final_document.sections, key=lambda item: item.order):
        title = re.sub(r"\s+", " ", section.title).strip()

        paragraphs = [
            re.sub(r"[ \t]+", " ", paragraph).strip()
            for paragraph in section.paragraphs
            if paragraph.strip()
        ]

        title_key = title.lower()
        sequential_section = any(
            token in title_key
            for token in (
                "paso a paso",
                "procedimiento",
                "actividades",
                "metodología",
                "metodologia",
                "etapas",
            )
        )

        numbered_candidates: list[tuple[int | None, int, int, str]] = []
        remaining_bullets: list[str] = []

        for item_index, original in enumerate(section.numbered_items):
            cleaned = _strip_list_prefix(original)
            if cleaned:
                numbered_candidates.append(
                    (
                        explicit_number(original),
                        0,
                        item_index,
                        cleaned,
                    )
                )

        for item_index, original in enumerate(section.bullets):
            cleaned = _strip_list_prefix(original)
            if not cleaned:
                continue

            detected_number = explicit_number(original)
            if detected_number is not None or sequential_section:
                numbered_candidates.append(
                    (
                        detected_number,
                        1,
                        item_index,
                        cleaned,
                    )
                )
            else:
                remaining_bullets.append(cleaned)

        if any(item[0] is not None for item in numbered_candidates):
            numbered_candidates.sort(
                key=lambda item: (
                    item[0] is None,
                    item[0] if item[0] is not None else 10**9,
                    item[1],
                    item[2],
                )
            )
        else:
            numbered_candidates.sort(key=lambda item: (item[1], item[2]))

        numbered_items = list(
            dict.fromkeys(item[3] for item in numbered_candidates)
        )
        bullets = list(dict.fromkeys(remaining_bullets))

        tables: list[FinalTable] = []
        for table in section.tables:
            headers = [
                re.sub(r"\s+", " ", str(header)).strip()
                for header in table.headers
            ]
            rows = [
                [
                    re.sub(r"\s+", " ", str(value)).strip()
                    for value in row
                ]
                for row in table.rows
            ]
            tables.append(
                table.model_copy(
                    update={
                        "title": re.sub(r"\s+", " ", table.title).strip(),
                        "headers": headers,
                        "rows": rows,
                    }
                )
            )

        normalized_sections.append(
            section.model_copy(
                update={
                    "title": title,
                    "paragraphs": paragraphs,
                    "bullets": bullets,
                    "numbered_items": numbered_items,
                    "tables": tables,
                }
            )
        )

    return final_document.model_copy(
        update={
            "title": re.sub(r"\s+", " ", final_document.title).strip(),
            "subtitle": re.sub(r"\s+", " ", final_document.subtitle).strip(),
            "sections": normalized_sections,
        }
    )


def audit_final_document_with_gemini(
    api_key: str,
    model: str,
    guides: list[dict],
    source: dict,
    analysis: GuideAnalysis,
    final_document: FinalDocument,
) -> AuditReport:
    """Audita el documento editado contra las guías y el origen."""

    relevant_indices = [analysis.selected_guide_index]
    relevant_indices.extend(analysis.supporting_guide_indices)

    parts: list[types.Part] = []
    for index in relevant_indices:
        append_record_to_gemini_input(
            parts,
            guides[index],
            f"GUÍA PARA AUDITORÍA {index}",
        )

    append_record_to_gemini_input(parts, source, "DOCUMENTO DE ORIGEN")
    parts.append(
        types.Part.from_text(
            text=audit_prompt(analysis, final_document)
        )
    )

    result = call_gemini_structured(
        api_key=api_key,
        model=model,
        input_parts=parts,
        schema_model=AuditReport,
    )

    assert isinstance(result, AuditReport)
    return result


def regenerate_section_with_gemini(
    api_key: str,
    model: str,
    guides: list[dict],
    source: dict,
    analysis: GuideAnalysis,
    answers: list[UserAnswer],
    final_document: FinalDocument,
    section_order: int,
) -> FinalSection:
    """Regenera una sección específica sin rehacer el documento completo."""

    section = next(
        (
            item
            for item in final_document.sections
            if item.order == section_order
        ),
        None,
    )
    if section is None:
        raise AppError("No se encontró la sección seleccionada para regenerar.")

    relevant_indices = [analysis.selected_guide_index]
    relevant_indices.extend(analysis.supporting_guide_indices)

    parts: list[types.Part] = []
    for index in relevant_indices:
        append_record_to_gemini_input(
            parts,
            guides[index],
            f"GUÍA PARA REGENERACIÓN {index}",
        )

    append_record_to_gemini_input(parts, source, "DOCUMENTO DE ORIGEN")
    parts.append(
        types.Part.from_text(
            text=section_regeneration_prompt(
                analysis,
                answers,
                final_document,
                section,
            )
        )
    )

    result = call_gemini_structured(
        api_key=api_key,
        model=model,
        input_parts=parts,
        schema_model=FinalSection,
    )

    assert isinstance(result, FinalSection)
    result.order = section_order
    return result

def analyze_guides_and_source(
    api_key: str,
    model: str,
    guides: list[dict],
    source: dict,
) -> GuideAnalysis:
    """Selecciona la guía, evalúa el origen y depura preguntas repetidas."""

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
    result = _deduplicate_missing_questions(result)

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
    """Redacta el borrador y normaliza su estructura documental."""

    relevant_indices = [analysis.selected_guide_index]
    relevant_indices.extend(analysis.supporting_guide_indices)

    parts: list[types.Part] = []

    for index in relevant_indices:
        append_record_to_gemini_input(
            parts,
            guides[index],
            f"GUÍA APLICABLE {index}",
        )

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
    result = normalize_final_document_structure(result)

    if not result.sections:
        raise AppError("Gemini no generó las secciones del documento final.")

    return result


# =========================================================
# GENERACIÓN DE DOCX Y PDF
# =========================================================


PAGE_TOTAL_PATTERN = re.compile(
    r"(?i)\b(p[áa]gina)(\s*:?\s*)\d+\s*(?:de|/)\s*\d+\b"
)
PAGE_SINGLE_PATTERN = re.compile(
    r"(?i)\b(p[áa]gina)(\s*:?\s*)\d+\b"
)


def _set_update_fields_on_open(document: Document) -> None:
    """Solicita a Word/LibreOffice actualizar PAGE y NUMPAGES."""

    settings = document.settings._element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")


def _copy_run_format(run, source_rpr) -> None:
    if source_rpr is not None:
        run._r.insert(0, deepcopy(source_rpr))


def _append_word_field(paragraph, field_code: str, source_rpr=None) -> None:
    """Inserta un campo de Word, por ejemplo PAGE o NUMPAGES."""

    run = paragraph.add_run()
    _copy_run_format(run, source_rpr)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")

    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = f" {field_code} "

    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")

    display = OxmlElement("w:t")
    display.text = "1"

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    run._r.extend([begin, instruction, separate, display, end])


def _clear_paragraph_content(paragraph) -> None:
    """Limpia runs conservando las propiedades del párrafo."""

    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def _replace_page_markers_in_paragraph(paragraph) -> bool:
    """Reemplaza paginación estática por campos PAGE/NUMPAGES."""

    original_text = paragraph.text
    if not original_text.strip():
        return False

    tokenized = PAGE_TOTAL_PATTERN.sub(
        lambda match: (
            f"{match.group(1)}{match.group(2)}"
            "[[PAGE_CURRENT]] de [[PAGE_TOTAL]]"
        ),
        original_text,
    )

    if "[[PAGE_CURRENT]]" not in tokenized:
        tokenized = PAGE_SINGLE_PATTERN.sub(
            lambda match: (
                f"{match.group(1)}{match.group(2)}[[PAGE_CURRENT]]"
            ),
            tokenized,
        )

    if "[[PAGE_CURRENT]]" not in tokenized:
        return False

    source_rpr = None
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        source_rpr = deepcopy(paragraph.runs[0]._r.rPr)

    _clear_paragraph_content(paragraph)

    tokens = re.split(
        r"(\[\[PAGE_CURRENT\]\]|\[\[PAGE_TOTAL\]\])",
        tokenized,
    )
    for token in tokens:
        if not token:
            continue
        if token == "[[PAGE_CURRENT]]":
            _append_word_field(paragraph, "PAGE", source_rpr)
        elif token == "[[PAGE_TOTAL]]":
            _append_word_field(paragraph, "NUMPAGES", source_rpr)
        else:
            run = paragraph.add_run(token)
            _copy_run_format(run, source_rpr)

    return True


def _iter_container_paragraphs(container):
    for paragraph in container.paragraphs:
        yield paragraph

    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_container_paragraphs(cell)


def apply_dynamic_page_numbering(document: Document) -> int:
    """
    Convierte textos como “PÁGINA: 1 de 2” y “Página 1” en campos
    dinámicos que se actualizan al abrir o convertir el archivo.
    """

    updated = 0
    visited_parts: set[int] = set()

    for section in document.sections:
        containers = (
            section.header,
            section.first_page_header,
            section.even_page_header,
            section.footer,
            section.first_page_footer,
            section.even_page_footer,
        )

        for container in containers:
            part_id = id(container.part)
            if part_id in visited_parts:
                continue
            visited_parts.add(part_id)

            for paragraph in _iter_container_paragraphs(container):
                if _replace_page_markers_in_paragraph(paragraph):
                    updated += 1

    _set_update_fields_on_open(document)
    return updated


def _set_row_repeat_header(row) -> None:
    """Repite la primera fila de una tabla al pasar de página."""

    tr_pr = row._tr.get_or_add_trPr()
    element = tr_pr.find(qn("w:tblHeader"))
    if element is None:
        element = OxmlElement("w:tblHeader")
        tr_pr.append(element)
    element.set(qn("w:val"), "true")


def _set_row_cant_split(row) -> None:
    """Evita que una fila se parta entre dos páginas."""

    tr_pr = row._tr.get_or_add_trPr()
    element = tr_pr.find(qn("w:cantSplit"))
    if element is None:
        element = OxmlElement("w:cantSplit")
        tr_pr.append(element)


def _normalize_document_text(text: str) -> str:
    return re.sub(r"\s+", " ", str(text)).strip()


def _format_body_paragraph(paragraph, justify: bool = True) -> None:
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.widow_control = True
    if justify:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _format_heading_paragraph(paragraph) -> None:
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.keep_together = True
    paragraph.paragraph_format.widow_control = True

def has_docx_style(document: Document, style_name: str) -> bool:
    """Indica si el estilo está materializado dentro del DOCX seleccionado."""

    try:
        document.styles[style_name]
        return True
    except (KeyError, ValueError):
        return False


def configure_docx(document: Document, preserve_guide_layout: bool = False) -> None:
    """Configura tipografía y espaciado base sin romper la guía institucional."""

    if not preserve_guide_layout:
        for section in document.sections:
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)

    if has_docx_style(document, "Normal"):
        normal_style = document.styles["Normal"]
        if not normal_style.font.name:
            normal_style.font.name = "Arial"
        if not normal_style.font.size:
            normal_style.font.size = Pt(10.5)
        normal_style.paragraph_format.space_after = Pt(6)
        normal_style.paragraph_format.line_spacing = 1.15
        normal_style.paragraph_format.widow_control = True

    for style_name in ("Title", "Heading 1", "Heading 2"):
        if has_docx_style(document, style_name):
            style = document.styles[style_name]
            if not style.font.name:
                style.font.name = "Arial"
            style.paragraph_format.keep_with_next = True
            style.paragraph_format.widow_control = True


def add_safe_heading(document: Document, text: str):
    """Agrega un título unido al primer elemento de su sección."""

    clean_text = _normalize_document_text(text)
    if not clean_text:
        return None

    if has_docx_style(document, "Heading 1"):
        paragraph = document.add_paragraph(clean_text, style="Heading 1")
        _format_heading_paragraph(paragraph)
        return paragraph

    paragraph = document.add_paragraph()
    _format_heading_paragraph(paragraph)
    run = paragraph.add_run(clean_text)
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(12)
    return paragraph


def add_safe_bullet(document: Document, text: str):
    """Agrega una viñeta limpia sin depender de estilos opcionales."""

    clean_text = _strip_list_prefix(text)
    if not clean_text:
        return None

    if has_docx_style(document, "List Bullet"):
        paragraph = document.add_paragraph(clean_text, style="List Bullet")
    else:
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.left_indent = Cm(0.65)
        paragraph.paragraph_format.first_line_indent = Cm(-0.35)
        bullet_run = paragraph.add_run("• ")
        bullet_run.bold = True
        paragraph.add_run(clean_text)

    _format_body_paragraph(paragraph, justify=False)
    paragraph.paragraph_format.keep_together = True
    return paragraph


def add_safe_numbered(document: Document, text: str, number: int):
    """Agrega numeración limpia, sin producir combinaciones como “• 1.”."""

    clean_text = _strip_list_prefix(text)
    if not clean_text:
        return None

    paragraph = document.add_paragraph()
    paragraph.paragraph_format.left_indent = Cm(0.72)
    paragraph.paragraph_format.first_line_indent = Cm(-0.52)
    paragraph.add_run(f"{number}. ").bold = True
    paragraph.add_run(clean_text)
    _format_body_paragraph(paragraph, justify=False)
    paragraph.paragraph_format.keep_together = True
    return paragraph

def apply_safe_table_borders(table) -> None:
    """Dibuja bordes aunque la guía no contenga el estilo Table Grid."""

    table_properties = table._tbl.tblPr
    borders = table_properties.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table_properties.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        border = borders.find(qn(tag))
        if border is None:
            border = OxmlElement(tag)
            borders.append(border)
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "B7B7B7")


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
    """Agrega una tabla con alineación, filas repetibles y altura automática."""

    if table_data.title.strip():
        title_paragraph = document.add_paragraph()
        _format_heading_paragraph(title_paragraph)
        title_run = title_paragraph.add_run(
            _normalize_document_text(table_data.title)
        )
        title_run.bold = True

    column_count = max(
        len(table_data.headers),
        max((len(row) for row in table_data.rows), default=0),
    )

    if column_count == 0:
        return

    table = document.add_table(rows=1, cols=column_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    if has_docx_style(document, "Table Grid"):
        table.style = "Table Grid"
    else:
        apply_safe_table_borders(table)

    headers = [
        _normalize_document_text(header)
        for header in table_data.headers
    ]
    headers.extend([""] * (column_count - len(headers)))

    header_row = table.rows[0]
    _set_row_repeat_header(header_row)
    _set_row_cant_split(header_row)

    for index, header in enumerate(headers):
        cell = header_row.cells[index]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.keep_together = True
        run = paragraph.add_run(header) if not paragraph.text else None
        if run is None:
            paragraph.text = header
        for existing_run in paragraph.runs:
            existing_run.bold = True
            existing_run.font.name = "Arial"
            existing_run.font.size = Pt(9)

    normalized_headers = [
        re.sub(r"[^a-záéíóúñ]", "", header.lower())
        for header in headers
    ]

    for source_row in table_data.rows:
        row_values = [
            _normalize_document_text(value)
            for value in source_row
        ]
        row_values.extend([""] * (column_count - len(row_values)))

        row = table.add_row()
        _set_row_cant_split(row)

        for index, value in enumerate(row_values):
            cell = row.cells[index]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.text = value
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.keep_together = True

            header_key = normalized_headers[index] if index < len(normalized_headers) else ""
            if header_key in {"fecha", "versión", "version"}:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            for run in paragraph.runs:
                run.font.name = "Arial"
                run.font.size = Pt(9)

    spacer = document.add_paragraph("")
    spacer.paragraph_format.space_after = Pt(4)


def create_docx(
    final_document: FinalDocument,
    style_guide: dict | None = None,
) -> bytes:
    """
    Crea el Word final y conserva el diseño institucional de la guía DOCX.
    También convierte la paginación estática del encabezado y pie en campos
    PAGE/NUMPAGES dinámicos.
    """

    final_document = normalize_final_document_structure(final_document)
    document, inherited_layout = load_visual_guide_document(style_guide)
    configure_docx(document, preserve_guide_layout=inherited_layout)
    apply_dynamic_page_numbering(document)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(6)
    title.paragraph_format.keep_with_next = True
    title_run = title.add_run(final_document.title.strip())
    title_run.bold = True
    title_run.font.name = "Arial"
    title_run.font.size = Pt(15)

    if final_document.subtitle.strip():
        subtitle = document.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.paragraph_format.space_after = Pt(10)
        subtitle_run = subtitle.add_run(final_document.subtitle.strip())
        subtitle_run.italic = True
        subtitle_run.font.name = "Arial"
        subtitle_run.font.size = Pt(10.5)

    if final_document.introductory_note.strip():
        introductory = document.add_paragraph(
            final_document.introductory_note.strip()
        )
        _format_body_paragraph(introductory)

    for section in final_document.sections:
        add_safe_heading(document, section.title)

        first_content_paragraph = None

        for paragraph_text in section.paragraphs:
            if paragraph_text.strip():
                paragraph = document.add_paragraph(paragraph_text.strip())
                _format_body_paragraph(paragraph)
                if first_content_paragraph is None:
                    first_content_paragraph = paragraph

        for number, item in enumerate(section.numbered_items, start=1):
            paragraph = add_safe_numbered(document, item, number)
            if first_content_paragraph is None and paragraph is not None:
                first_content_paragraph = paragraph

        for bullet in section.bullets:
            paragraph = add_safe_bullet(document, bullet)
            if first_content_paragraph is None and paragraph is not None:
                first_content_paragraph = paragraph

        for table_data in section.tables:
            add_docx_table(document, table_data)

    _set_update_fields_on_open(document)

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


class NumberedCanvas(pdf_canvas.Canvas):
    """Canvas de respaldo con numeración dinámica Página X de Y."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: list[dict] = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self._saved_page_states)
        for page_number, state in enumerate(self._saved_page_states, start=1):
            self.__dict__.update(state)
            self.setFont("Helvetica", 8)
            self.drawCentredString(
                A4[0] / 2,
                1.25 * cm,
                f"Página {page_number} de {page_count}",
            )
            super().showPage()
        super().save()

def create_pdf(final_document: FinalDocument) -> bytes:
    """Crea una versión PDF estructurada con numeración dinámica."""

    final_document = normalize_final_document_structure(final_document)
    output = BytesIO()
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocumentTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        alignment=TA_CENTER,
        spaceAfter=8,
    )
    subtitle_style = ParagraphStyle(
        "DocumentSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=10,
        leading=13,
        alignment=TA_CENTER,
        spaceAfter=12,
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=11.5,
        leading=14,
        spaceBefore=10,
        spaceAfter=5,
        keepWithNext=True,
    )
    body_style = ParagraphStyle(
        "BodyTextCustom",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        spaceAfter=5,
        alignment=0,
        allowWidows=0,
        allowOrphans=0,
    )
    bullet_style = ParagraphStyle(
        "BulletCustom",
        parent=body_style,
        leftIndent=14,
        firstLineIndent=-8,
        bulletIndent=4,
    )
    numbered_style = ParagraphStyle(
        "NumberedCustom",
        parent=body_style,
        leftIndent=18,
        firstLineIndent=-14,
    )
    table_cell_style = ParagraphStyle(
        "TableCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        alignment=0,
    )
    table_center_style = ParagraphStyle(
        "TableCellCenter",
        parent=table_cell_style,
        alignment=TA_CENTER,
    )
    table_header_style = ParagraphStyle(
        "TableHeader",
        parent=table_cell_style,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
    )

    story: list = [
        pdf_paragraph(final_document.title.strip(), title_style),
    ]

    if final_document.subtitle.strip():
        story.append(
            pdf_paragraph(final_document.subtitle.strip(), subtitle_style)
        )

    if final_document.introductory_note.strip():
        story.append(
            pdf_paragraph(final_document.introductory_note.strip(), body_style)
        )
        story.append(Spacer(1, 4))

    available_width = A4[0] - 5 * cm

    for section in final_document.sections:
        section_story: list = [
            pdf_paragraph(section.title.strip(), heading_style),
        ]

        for paragraph_text in section.paragraphs:
            if paragraph_text.strip():
                section_story.append(
                    pdf_paragraph(paragraph_text.strip(), body_style)
                )

        for index, item in enumerate(section.numbered_items, start=1):
            if item.strip():
                section_story.append(
                    Paragraph(
                        f"<b>{index}.</b> {escape(_strip_list_prefix(item))}",
                        numbered_style,
                    )
                )

        for bullet in section.bullets:
            if bullet.strip():
                section_story.append(
                    Paragraph(
                        f"• {escape(_strip_list_prefix(bullet))}",
                        bullet_style,
                    )
                )

        if len(section_story) >= 2:
            story.append(KeepTogether(section_story[:2]))
            story.extend(section_story[2:])
        else:
            story.extend(section_story)

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

            headers = [
                _normalize_document_text(header)
                for header in table_data.headers
            ]
            headers.extend([""] * (column_count - len(headers)))

            normalized_headers = [
                re.sub(r"[^a-záéíóúñ]", "", header.lower())
                for header in headers
            ]

            rows = [
                headers,
                *[
                    [
                        _normalize_document_text(value)
                        for value in list(row)
                    ]
                    + [""] * (column_count - len(row))
                    for row in table_data.rows
                ],
            ]

            formatted_rows: list[list[Paragraph]] = []
            for row_index, row in enumerate(rows):
                formatted_cells: list[Paragraph] = []
                for column_index, value in enumerate(row):
                    if row_index == 0:
                        cell_style = table_header_style
                    elif normalized_headers[column_index] in {
                        "fecha",
                        "versión",
                        "version",
                    }:
                        cell_style = table_center_style
                    else:
                        cell_style = table_cell_style
                    formatted_cells.append(
                        pdf_paragraph(str(value), cell_style)
                    )
                formatted_rows.append(formatted_cells)

            col_widths = [available_width / column_count] * column_count
            table = Table(
                formatted_rows,
                colWidths=col_widths,
                repeatRows=1,
                hAlign="LEFT",
                splitByRow=1,
            )
            table.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 5),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            story.append(table)
            story.append(Spacer(1, 8))

    pdf = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=2.5 * cm,
        leftMargin=2.5 * cm,
        topMargin=2.5 * cm,
        bottomMargin=2.1 * cm,
        title=final_document.title,
        author="Generador inteligente de documentos",
    )
    pdf.build(story, canvasmaker=NumberedCanvas)
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
        "audit_summary": None,
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
    st.session_state.audit_summary = None
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
    columns[3].metric(
        "Preguntas obligatorias",
        count_required_questions(analysis),
    )

    status_counts = {
        status: sum(
            1 for section in analysis.sections if section.status == status
        )
        for status in STATUS_LABELS
    }
    completed_ratio = (
        status_counts["completo"] / len(analysis.sections)
        if analysis.sections
        else 0
    )
    st.progress(
        completed_ratio,
        text=(
            f"{status_counts['completo']} de {len(analysis.sections)} "
            "secciones tienen información completa"
        ),
    )

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

    comparison_rows = []
    for section in analysis.sections:
        comparison_rows.append(
            {
                "Sección": section.title,
                "Estado": STATUS_LABELS[section.status],
                "Criterios": len(section.criteria),
                "Evidencias": len(section.evidence),
                "Preguntas": len(section.missing_questions),
            }
        )

    st.markdown("### Comparación guía vs. información encontrada")
    st.dataframe(
        comparison_rows,
        width="stretch",
        hide_index=True,
        column_config={
            "Sección": st.column_config.TextColumn(width="large"),
            "Estado": st.column_config.TextColumn(width="small"),
        },
    )

    st.markdown("### Evaluación detallada por sección")

    for section in analysis.sections:
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
                    st.write(
                        f"- **{question.question}** "
                        f"({question.category} · {required})"
                    )
                    st.caption(question.why_needed)

            st.caption(f"Formato previsto: {section.output_format}")


def collect_answers_form(analysis: GuideAnalysis) -> list[UserAnswer] | None:
    """Agrupa preguntas críticas por categoría y valida las obligatorias."""

    all_questions = [
        (section, question_index, question)
        for section in analysis.sections
        for question_index, question in enumerate(section.missing_questions)
    ]

    if not all_questions:
        st.success(
            "La información disponible es suficiente. "
            "Puedes generar el borrador directamente."
        )

    grouped_questions: dict[str, list[tuple]] = {}
    for item in all_questions:
        category = item[2].category.strip() or "Información del proceso"
        grouped_questions.setdefault(category, []).append(item)

    answers_by_key: dict[str, str] = {}

    with st.form("missing_information_form", clear_on_submit=False):
        st.markdown("### Completar información faltante")
        st.caption(
            "Responde únicamente con información verificable. "
            "Los campos marcados con * son obligatorios."
        )

        item_index = 0
        for category, category_items in grouped_questions.items():
            st.markdown(f"#### {category}")

            for section, question_index, question in category_items:
                label = question.question + (" *" if question.required else "")
                help_text = f"Sección: {section.title}. {question.why_needed}"
                key = (
                    f"answer_{section.order}_{question_index}_{item_index}"
                )

                answers_by_key[key] = st.text_area(
                    label,
                    height=90,
                    help=help_text,
                    key=key,
                    placeholder="Escribe una respuesta concreta y verificable.",
                )
                item_index += 1

        submitted = st.form_submit_button(
            "Crear borrador para revisión",
            type="primary",
            width="stretch",
        )

    if not submitted:
        return None

    missing_required: list[str] = []
    answers: list[UserAnswer] = []
    item_index = 0

    for category_items in grouped_questions.values():
        for section, question_index, question in category_items:
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
            item_index += 1

    if missing_required:
        st.error(
            "Debes responder los siguientes datos obligatorios:\n\n- "
            + "\n- ".join(missing_required)
        )
        return None

    return answers


def _split_paragraph_editor(value: str) -> list[str]:
    """Convierte bloques separados por línea en blanco en párrafos."""

    return [
        re.sub(r"[ \t]+", " ", item).strip()
        for item in re.split(r"\n\s*\n", value.strip())
        if item.strip()
    ]


def _split_list_editor(value: str) -> list[str]:
    """Convierte una línea por elemento en una lista limpia."""

    return [
        _strip_list_prefix(item)
        for item in value.splitlines()
        if _strip_list_prefix(item)
    ]


def _records_from_data_editor(edited_data) -> list[dict]:
    if hasattr(edited_data, "to_dict"):
        return edited_data.to_dict(orient="records")
    if isinstance(edited_data, list):
        return [dict(item) for item in edited_data]
    return []


def replace_section_in_document(
    final_document: FinalDocument,
    regenerated_section: FinalSection,
) -> FinalDocument:
    sections = [
        regenerated_section
        if section.order == regenerated_section.order
        else section
        for section in final_document.sections
    ]
    return normalize_final_document_structure(
        final_document.model_copy(update={"sections": sections})
    )


def edit_final_document_form(
    final_document: FinalDocument,
) -> FinalDocument | None:
    """Permite editar títulos, texto, listas y celdas antes de generar."""

    final_document = normalize_final_document_structure(final_document)
    section_values: list[dict] = []

    with st.form("final_document_editor", clear_on_submit=False):
        st.markdown("### Vista previa editable")
        st.caption(
            "Revisa el contenido antes de crear los archivos. "
            "Los cambios guardados reemplazarán el borrador de la IA."
        )

        edited_title = st.text_input(
            "Título del documento",
            value=final_document.title,
        )
        edited_subtitle = st.text_input(
            "Subtítulo",
            value=final_document.subtitle,
        )
        edited_intro = st.text_area(
            "Nota introductoria",
            value=final_document.introductory_note,
            height=100,
        )

        for section_index, section in enumerate(final_document.sections):
            with st.expander(
                f"{section.order}. {section.title}",
                expanded=section_index == 0,
            ):
                title_value = st.text_input(
                    "Título de la sección",
                    value=section.title,
                    key=f"edit_title_{section.order}",
                )
                paragraphs_value = st.text_area(
                    "Párrafos",
                    value="\n\n".join(section.paragraphs),
                    height=max(120, min(320, 80 + 35 * len(section.paragraphs))),
                    key=f"edit_paragraphs_{section.order}",
                    help="Separa los párrafos con una línea en blanco.",
                )
                numbered_value = st.text_area(
                    "Pasos numerados",
                    value="\n".join(section.numbered_items),
                    height=max(90, min(250, 70 + 25 * len(section.numbered_items))),
                    key=f"edit_numbered_{section.order}",
                    help=(
                        "Escribe un paso por línea. No agregues 1., 2. ni viñetas; "
                        "el bot aplicará la numeración."
                    ),
                )
                bullets_value = st.text_area(
                    "Viñetas",
                    value="\n".join(section.bullets),
                    height=max(90, min(250, 70 + 25 * len(section.bullets))),
                    key=f"edit_bullets_{section.order}",
                    help="Escribe un elemento por línea, sin el símbolo de viñeta.",
                )

                edited_tables: list[tuple[str, list[str], object]] = []
                for table_index, table_data in enumerate(section.tables):
                    st.markdown(f"**Tabla {table_index + 1}**")
                    table_title = st.text_input(
                        "Título de la tabla",
                        value=table_data.title,
                        key=f"table_title_{section.order}_{table_index}",
                    )

                    headers = list(table_data.headers)
                    if headers:
                        records = []
                        for row in table_data.rows:
                            padded = list(row) + [""] * (len(headers) - len(row))
                            records.append(
                                {
                                    header: padded[column_index]
                                    for column_index, header in enumerate(headers)
                                }
                            )

                        editor_source = (
                            records
                            if records
                            else {header: [] for header in headers}
                        )
                        edited_data = st.data_editor(
                            editor_source,
                            num_rows="dynamic",
                            width="stretch",
                            hide_index=True,
                            key=f"table_editor_{section.order}_{table_index}",
                        )
                        edited_tables.append(
                            (table_title, headers, edited_data)
                        )
                    else:
                        st.caption("La tabla no tiene encabezados editables.")

                section_values.append(
                    {
                        "section": section,
                        "title": title_value,
                        "paragraphs": paragraphs_value,
                        "numbered": numbered_value,
                        "bullets": bullets_value,
                        "tables": edited_tables,
                    }
                )

        submitted = st.form_submit_button(
            "Guardar cambios, auditar y generar archivos",
            type="primary",
            width="stretch",
        )

    if not submitted:
        return None

    rebuilt_sections: list[FinalSection] = []

    for values in section_values:
        original_section: FinalSection = values["section"]
        rebuilt_tables: list[FinalTable] = []

        for table_title, headers, edited_data in values["tables"]:
            records = _records_from_data_editor(edited_data)
            rows = [
                [
                    _normalize_document_text(record.get(header, ""))
                    for header in headers
                ]
                for record in records
            ]
            rebuilt_tables.append(
                FinalTable(
                    title=table_title.strip(),
                    headers=headers,
                    rows=rows,
                )
            )

        if not values["tables"]:
            rebuilt_tables = original_section.tables

        rebuilt_sections.append(
            original_section.model_copy(
                update={
                    "title": values["title"].strip(),
                    "paragraphs": _split_paragraph_editor(values["paragraphs"]),
                    "numbered_items": _split_list_editor(values["numbered"]),
                    "bullets": _split_list_editor(values["bullets"]),
                    "tables": rebuilt_tables,
                }
            )
        )

    edited_document = final_document.model_copy(
        update={
            "title": edited_title.strip() or final_document.title,
            "subtitle": edited_subtitle.strip(),
            "introductory_note": edited_intro.strip(),
            "sections": rebuilt_sections,
        }
    )

    return normalize_final_document_structure(edited_document)


def show_error_panel(
    user_message: str,
    error: Exception,
    technical_label: str = "Ver detalle técnico",
) -> None:
    """Muestra un mensaje comprensible y oculta el detalle técnico."""

    st.error(user_message)
    with st.expander(technical_label):
        st.code(str(error)[:2000], language=None)

def show_final_document(final_document: FinalDocument) -> None:
    st.subheader("Documento final")
    st.markdown(f"## {final_document.title}")

    if final_document.subtitle.strip():
        st.caption(final_document.subtitle)

    if final_document.introductory_note.strip():
        st.write(final_document.introductory_note)

    for section in final_document.sections:
        with st.expander(f"{section.order}. {section.title}", expanded=False):
            for paragraph_text in section.paragraphs:
                st.write(paragraph_text)

            for index, item in enumerate(section.numbered_items, start=1):
                st.write(f"{index}. {item}")

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
                        normalized_rows.append(
                            padded[: len(table_data.headers)]
                        )

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

    total_checks = sum(compliance_counts.values())
    compliance_ratio = (
        compliance_counts["cumple"] / total_checks
        if total_checks
        else 0
    )
    st.progress(
        compliance_ratio,
        text=(
            f"{compliance_counts['cumple']} de {total_checks} "
            "criterios cumplen completamente"
        ),
    )

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
                f"{icon} **{item.section_title}** — "
                f"{item.criterion}: {item.note}"
            )

    audit_summary = st.session_state.get("audit_summary")
    if audit_summary:
        st.info(f"**Resumen editorial:** {audit_summary}")

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
    if st.session_state.get("docx_output") is not None:
        return 4
    if st.session_state.get("final_document") is not None:
        return 3
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
            1. **Carga las guías institucionales** y el documento de origen.
            2. El sistema **interpreta, extrae evidencia, redacta y audita**.
            3. Solo solicita **datos críticos que realmente estén ausentes**.
            4. Presenta un **borrador editable por secciones** antes de generar.
            5. Crea un **Word institucional** y su PDF con paginación dinámica.
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
        """<div class="friendly-note"><strong>Consejo:</strong> para conservar
        exactamente el encabezado, pie de página, logos y márgenes, carga al
        menos una guía en formato <strong>DOCX</strong>.</div>""",
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

            names = [
                safe_filename(upload.name).lower()
                for upload in guide_uploads
            ]
            if len(names) != len(set(names)):
                raise AppError("Hay guías con nombres duplicados.")

            with st.status(
                "Analizando el proceso documental...",
                expanded=True,
            ) as status:
                st.write("1/4 · Leyendo y validando los archivos")
                guide_records = [
                    build_file_record(
                        upload,
                        role="guía normativa",
                        index=index,
                    )
                    for index, upload in enumerate(guide_uploads)
                ]
                source_record = build_file_record(
                    source_upload,
                    role="documento de origen",
                    index=0,
                )

                total_text = sum(
                    len(record["text"]) for record in guide_records
                )
                total_text += len(source_record["text"])

                if total_text > MAX_TOTAL_TEXT_CHARS:
                    st.warning(
                        "Los archivos contienen mucho texto. Se limitará el "
                        "contenido textual por archivo; los PDF se enviarán "
                        "completos de forma nativa."
                    )

                st.write("2/4 · Interpretando las reglas de las guías")
                st.write("3/4 · Extrayendo evidencia del documento de origen")
                analysis = analyze_guides_and_source(
                    api_key=api_key,
                    model=model or MODEL_DEFAULT,
                    guides=guide_records,
                    source=source_record,
                )
                st.write("4/4 · Identificando brechas y preguntas críticas")
                status.update(
                    label="Análisis completado",
                    state="complete",
                    expanded=False,
                )

            st.session_state.guide_records = guide_records
            st.session_state.source_record = source_record
            st.session_state.analysis = analysis
            st.session_state.final_document = None
            st.session_state.docx_output = None
            st.session_state.pdf_output = None
            st.session_state.answers = None
            st.session_state.audit_summary = None

            st.success("Análisis completado correctamente.")

        except AppError as error:
            show_error_panel(
                "No fue posible completar el análisis. "
                "Revisa el mensaje técnico y vuelve a intentarlo.",
                error,
            )
        except Exception as error:
            show_error_panel(
                "Ocurrió un error inesperado durante el análisis.",
                error,
            )

    guide_records = st.session_state.guide_records
    source_record = st.session_state.source_record
    analysis = st.session_state.analysis
    final_document = st.session_state.final_document

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

        if final_document is None:
            st.divider()
            answers = collect_answers_form(analysis)

            if answers is not None:
                try:
                    with st.status(
                        "Construyendo el borrador profesional...",
                        expanded=True,
                    ) as status:
                        st.write("1/3 · Redactando las secciones")
                        draft_document = generate_final_with_gemini(
                            api_key=api_key,
                            model=model or MODEL_DEFAULT,
                            guides=guide_records,
                            source=source_record,
                            analysis=analysis,
                            answers=answers,
                        )
                        st.write("2/3 · Normalizando listas, tablas y redacción")
                        draft_document = normalize_final_document_structure(
                            draft_document
                        )
                        st.write("3/3 · Preparando la vista previa editable")
                        status.update(
                            label="Borrador listo para revisión",
                            state="complete",
                            expanded=False,
                        )

                    st.session_state.answers = answers
                    st.session_state.final_document = draft_document
                    st.session_state.docx_output = None
                    st.session_state.pdf_output = None
                    st.session_state.audit_summary = None
                    st.success(
                        "Borrador creado. Revísalo y edítalo antes de generar "
                        "los archivos."
                    )
                    st.rerun()

                except AppError as error:
                    show_error_panel(
                        "Gemini no pudo construir el borrador en este intento.",
                        error,
                    )
                except Exception as error:
                    show_error_panel(
                        "Ocurrió un error inesperado al construir el borrador.",
                        error,
                    )

    final_document = st.session_state.final_document

    if (
        final_document
        and analysis
        and guide_records
        and source_record
    ):
        st.divider()
        st.subheader("✍️ 3. Revisar y editar el borrador")

        st.markdown(
            """<div class="friendly-note"><strong>Control del usuario:</strong>
            puedes editar directamente cada sección o regenerar únicamente una
            sección sin rehacer todo el documento.</div>""",
            unsafe_allow_html=True,
        )

        section_options = {
            f"{section.order}. {section.title}": section.order
            for section in final_document.sections
        }
        selected_section_label = st.selectbox(
            "Sección para regenerar",
            options=list(section_options),
            help=(
                "La IA regenerará solo la sección seleccionada usando las "
                "mismas guías, el origen y tus respuestas."
            ),
        )

        if st.button(
            "Regenerar únicamente esta sección",
            width="stretch",
        ):
            try:
                with st.spinner("Regenerando la sección seleccionada..."):
                    regenerated = regenerate_section_with_gemini(
                        api_key=api_key,
                        model=model or MODEL_DEFAULT,
                        guides=guide_records,
                        source=source_record,
                        analysis=analysis,
                        answers=st.session_state.answers or [],
                        final_document=final_document,
                        section_order=section_options[selected_section_label],
                    )

                st.session_state.final_document = replace_section_in_document(
                    final_document,
                    regenerated,
                )
                st.session_state.docx_output = None
                st.session_state.pdf_output = None
                st.session_state.audit_summary = None
                st.success("Sección regenerada correctamente.")
                st.rerun()

            except AppError as error:
                show_error_panel(
                    "No fue posible regenerar la sección.",
                    error,
                )
            except Exception as error:
                show_error_panel(
                    "Ocurrió un error inesperado al regenerar la sección.",
                    error,
                )

        edited_document = edit_final_document_form(
            st.session_state.final_document
        )

        if edited_document is not None:
            selected_guide = guide_records[analysis.selected_guide_index]

            try:
                with st.status(
                    "Auditando y generando los archivos...",
                    expanded=True,
                ) as status:
                    st.write("1/4 · Guardando tus modificaciones")
                    edited_document = normalize_final_document_structure(
                        edited_document
                    )

                    st.write("2/4 · Auditando el contenido contra la guía")
                    try:
                        audit_report = audit_final_document_with_gemini(
                            api_key=api_key,
                            model=model or MODEL_DEFAULT,
                            guides=guide_records,
                            source=source_record,
                            analysis=analysis,
                            final_document=edited_document,
                        )
                        combined_warnings = list(
                            dict.fromkeys(
                                edited_document.warnings
                                + audit_report.warnings
                            )
                        )
                        edited_document = edited_document.model_copy(
                            update={
                                "validation": audit_report.validation,
                                "warnings": combined_warnings,
                            }
                        )
                        st.session_state.audit_summary = (
                            audit_report.editorial_summary
                        )
                    except Exception as audit_error:
                        st.warning(
                            "No fue posible completar la revalidación con "
                            "Gemini. El documento se generará con la validación "
                            "del borrador."
                        )
                        with st.expander("Ver detalle de la revalidación"):
                            st.code(str(audit_error)[:1600], language=None)

                    st.write("3/4 · Aplicando formato institucional y paginación")
                    docx_output = create_docx(
                        edited_document,
                        style_guide=selected_guide,
                    )

                    st.write("4/4 · Creando la versión PDF")
                    pdf_output = convert_docx_to_pdf(docx_output)
                    if pdf_output is None:
                        pdf_output = create_pdf(edited_document)

                    status.update(
                        label="Documento auditado y generado",
                        state="complete",
                        expanded=False,
                    )

                st.session_state.final_document = edited_document
                st.session_state.docx_output = docx_output
                st.session_state.pdf_output = pdf_output

                if selected_guide.get("extension") == "docx":
                    st.success(
                        "Documento generado con encabezado, pie, márgenes, "
                        "paginación dinámica y formato de la guía seleccionada."
                    )
                else:
                    st.success("Documento final generado y validado.")
                    st.info(
                        "La guía principal no es DOCX. Para reproducir exactamente "
                        "encabezado y pie, carga su versión Word."
                    )
                st.rerun()

            except AppError as error:
                show_error_panel(
                    "No fue posible generar el documento final.",
                    error,
                )
            except Exception as error:
                show_error_panel(
                    "Ocurrió un error inesperado durante la generación.",
                    error,
                )

    final_document = st.session_state.final_document
    docx_output = st.session_state.get("docx_output")
    pdf_output = st.session_state.get("pdf_output")

    if final_document and docx_output and pdf_output:
        st.divider()
        st.subheader("✅ 4. Documento listo")
        show_final_document(final_document)

        st.markdown("### Descargar archivos")
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
            docx_bytes = bytes(docx_output)
            pdf_bytes = bytes(pdf_output)

            if not docx_bytes.startswith(b"PK"):
                st.error(
                    "El Word generado no tiene una estructura DOCX válida."
                )
                files_are_valid = False

            if not pdf_bytes.startswith(b"%PDF"):
                st.error(
                    "El PDF generado no tiene una estructura PDF válida."
                )
                files_are_valid = False

        if files_are_valid:
            docx_name = output_filename(final_document.title, "docx")
            pdf_name = output_filename(final_document.title, "pdf")

            docx_path = save_generated_file(docx_name, docx_bytes)
            pdf_path = save_generated_file(pdf_name, pdf_bytes)

            st.markdown("#### Descarga directa")
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
                f"`{docx_path.as_posix()}` y `{pdf_path.as_posix()}`."
            )

    st.divider()
    st.caption(
        "La aplicación utiliza las guías como reglas de construcción, conserva "
        "la trazabilidad del origen y evita inventar información."
    )


if __name__ == "__main__":
    main()