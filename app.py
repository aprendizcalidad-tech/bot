from __future__ import annotations

import streamlit as st

from src.document_reader import (
    DocumentReadError,
    DocumentReadResult,
    read_document,
)


st.set_page_config(
    page_title="Generador Inteligente de Documentos",
    page_icon="📄",
    layout="centered",
)


def initialize_session_state() -> None:
    if "source_result" not in st.session_state:
        st.session_state.source_result = None


def show_document_result(result: DocumentReadResult) -> None:
    """Muestra el resultado de la extracción documental."""

    st.subheader("Resultado de la lectura")

    column_1, column_2, column_3 = st.columns(3)

    with column_1:
        st.metric("Formato", result.extension.upper())

    with column_2:
        st.metric("Caracteres", len(result.text))

    with column_3:
        if result.page_count is not None:
            st.metric("Páginas", result.page_count)
        else:
            st.metric("Páginas", "N/A")

    for warning in result.warnings:
        st.warning(warning)

    if result.text:
        st.markdown("#### Vista previa del texto")

        preview = result.text[:8000]

        st.text_area(
            label="Contenido extraído",
            value=preview,
            height=400,
            disabled=True,
        )

        if len(result.text) > len(preview):
            st.info(
                "La vista previa muestra los primeros 8.000 caracteres. "
                f"El documento completo contiene {len(result.text)} caracteres."
            )


def main() -> None:
    initialize_session_state()

    st.title("📄 Generador Inteligente de Documentos")

    st.write(
        "Carga las plantillas disponibles y el documento que contiene "
        "la información del proceso."
    )

    st.divider()

    st.subheader("1. Cargar plantillas")

    plantillas = st.file_uploader(
        label="Selecciona una o varias plantillas",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=True,
        help="Formatos permitidos: DOCX, PDF y TXT.",
    )

    st.subheader("2. Cargar documento de origen")

    documento_origen = st.file_uploader(
        label="Selecciona el documento con la información del proceso",
        type=["docx", "pdf", "txt"],
        accept_multiple_files=False,
        help="Solo se admite un documento de origen.",
    )

    st.divider()

    if plantillas:
        st.success(f"Plantillas cargadas: {len(plantillas)}")

        for plantilla in plantillas:
            st.write(f"📄 {plantilla.name}")

    if documento_origen:
        st.success("Documento de origen cargado correctamente")
        st.write(f"📑 {documento_origen.name}")

    archivos_completos = bool(plantillas) and documento_origen is not None

    if st.button(
        "Leer documento de origen",
        type="primary",
        use_container_width=True,
        disabled=not archivos_completos,
    ):
        try:
            with st.spinner("Extrayendo contenido del documento..."):
                result = read_document(
                    filename=documento_origen.name,
                    content=documento_origen.getvalue(),
                )

                st.session_state.source_result = result

            st.success("El documento fue leído correctamente.")

        except DocumentReadError as error:
            st.session_state.source_result = None
            st.error(str(error))

        except Exception:
            st.session_state.source_result = None
            st.error(
                "Ocurrió un error inesperado durante la lectura del archivo."
            )

    if st.session_state.source_result is not None:
        st.divider()
        show_document_result(st.session_state.source_result)


if __name__ == "__main__":
    main()