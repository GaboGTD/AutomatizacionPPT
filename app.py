import io
import os
import re
from datetime import datetime
from typing import Dict, Any, Tuple

import openpyxl
import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Pt

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="Generador de Presentaciones GTD",
    page_icon="📊",
    layout="centered"
)

TAMANO_FUENTE = Pt(9)
TABLAS_KV = ['DatosGenerales', 'Riesgo1']
TABLAS_QA = ['TablaCalificacion', 'Riesgo2']
RUTA_PLANTILLA = os.path.join(os.path.dirname(__file__), "Plantilla_Limpia.pptx")

# ==========================================
# FUNCIONES AUXILIARES
# ==========================================
def limpiar_texto(texto: Any) -> str:
    """Elimina espacios redundantes y normaliza a minúsculas."""
    if pd.isna(texto) or texto is None:
        return ""
    texto_str = str(texto)
    texto_limpio = re.sub(r'\s+', ' ', texto_str)
    return texto_limpio.strip().lower()

def aplicar_formato_texto(celda: Any, texto: Any) -> None:
    """Inserta texto y fuerza el tamaño de fuente fijo."""
    celda.text = str(texto) if pd.notna(texto) else ''
    if celda.text and celda.text_frame.paragraphs:
        celda.text_frame.paragraphs[0].font.size = TAMANO_FUENTE

def encontrar_tabla(wb: openpyxl.Workbook, nombre_tabla: str):
    """Localiza una tabla con nombre en cualquier hoja del libro."""
    for sheet in wb.worksheets:
        if nombre_tabla in sheet.tables:
            return sheet, sheet.tables[nombre_tabla]
    return None, None

def extraer_tabla_kv(wb: openpyxl.Workbook, nombre_tabla: str) -> Dict[str, Any]:
    """Extrae datos de tablas horizontales (Fila 1: Nombres, Fila 2: Valores)."""
    sheet, tabla = encontrar_tabla(wb, nombre_tabla)
    if not sheet:
        return {}
    data_rows = [[cell.value for cell in row] for row in sheet[tabla.ref]]
    datos = {}
    if len(data_rows) > 1:
        keys, values = data_rows[0], data_rows[1]
        for i, key in enumerate(keys):
            if key and i < len(values) and pd.notna(values[i]):
                datos[limpiar_texto(key)] = values[i]
                datos[str(key).strip()] = values[i]
    return datos

def extraer_tabla_qa(wb: openpyxl.Workbook, nombre_tabla: str) -> pd.DataFrame:
    """Extrae tablas de declaraciones/preguntas."""
    sheet, tabla = encontrar_tabla(wb, nombre_tabla)
    if not sheet:
        return pd.DataFrame()
    data_rows = [[cell.value for cell in row] for row in sheet[tabla.ref]]
    if len(data_rows) > 1:
        return pd.DataFrame(data_rows[1:], columns=data_rows[0])
    return pd.DataFrame()

def preparar_dict_qa(df: pd.DataFrame) -> Dict[str, pd.Series]:
    """Indexa un DataFrame por su primera columna limpia."""
    if df.empty:
        return {}
    col_clave = df.columns[0]
    diccionario = {}
    for _, fila in df.iterrows():
        llave = limpiar_texto(fila[col_clave])
        if llave:
            diccionario[llave] = fila
    return diccionario

# ==========================================
# MOTOR DE GENERACIÓN
# ==========================================
def procesar_presentacion(excel_bytes: bytes) -> Tuple[io.BytesIO, str]:
    excel_stream = io.BytesIO(excel_bytes)
    wb = openpyxl.load_workbook(excel_stream, data_only=True)

    # 1. Extracción de tablas generales
    datos_kv = {}
    for nombre in TABLAS_KV:
        datos_kv.update(extraer_tabla_kv(wb, nombre))

    if not datos_kv:
        raise ValueError("No se encontraron tablas de datos generales ('DatosGenerales' o 'Riesgo1').")

    # 2. Extracción de tablas de preguntas
    df_cal = extraer_tabla_qa(wb, TABLAS_QA[0])
    df_riesgo = extraer_tabla_qa(wb, TABLAS_QA[1])
    dict_cal = preparar_dict_qa(df_cal)
    dict_riesgo = preparar_dict_qa(df_riesgo)

    if not os.path.exists(RUTA_PLANTILLA):
        raise FileNotFoundError(f"No se encontró el archivo de plantilla: {RUTA_PLANTILLA}")

    prs = Presentation(RUTA_PLANTILLA)

    # 3. Procesamiento de diapositivas
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            tabla = shape.table

            # Identificación estadística de tablas de preguntas
            preguntas_ppt = []
            for i, row in enumerate(tabla.rows):
                if i > 0:
                    t = limpiar_texto(row.cells[0].text)
                    if t:
                        preguntas_ppt.append(t)

            matches_cal = sum(1 for p in preguntas_ppt if p in dict_cal)
            matches_riesgo = sum(1 for p in preguntas_ppt if p in dict_riesgo)

            if matches_cal > 0 or matches_riesgo > 0:
                tipo = "CALIFICACIÓN" if matches_cal >= matches_riesgo else "RIESGO"
                datos_excel = dict_cal if tipo == "CALIFICACIÓN" else dict_riesgo
                df_usar = df_cal if tipo == "CALIFICACIÓN" else df_riesgo

                for i, row in enumerate(tabla.rows):
                    if i == 0:
                        continue
                    clave = limpiar_texto(row.cells[0].text)
                    if clave in datos_excel:
                        fila = datos_excel[clave]
                        respuesta = fila.get('Respuesta', '')
                        if len(row.cells) > 1:
                            aplicar_formato_texto(row.cells[1], respuesta)
                        comentario = fila.get('Comentario/Mitigación', fila.get('Comentario/Mitigaciónn', ''))
                        if len(row.cells) > 2:
                            aplicar_formato_texto(row.cells[2], comentario)
            else:
                # Tabla superior (Clave-Valor)
                for row in tabla.rows:
                    for c, cell in enumerate(row.cells):
                        label = limpiar_texto(cell.text)
                        if label in datos_kv and c + 1 < len(row.cells):
                            aplicar_formato_texto(row.cells[c + 1], datos_kv[label])

    # 4. Construcción del nombre del archivo de salida
    nombre_op = "Oportunidad_Sin_Nombre"
    for k, v in datos_kv.items():
        if 'nombre' in k.lower() and 'oportunidad' in k.lower():
            nombre_op = str(v)
            break

    nombre_limpio = "".join(c for c in nombre_op if c.isalnum() or c in (' ', '_')).rstrip()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_archivo = f"Calificación_{nombre_limpio}_{timestamp}.pptx"

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer, nombre_archivo

# ==========================================
# INTERFAZ WEB (STREAMLIT)
# ==========================================
st.title("Generador Automático de Presentaciones")
st.markdown("Carga la matriz de calificación y riesgo en formato Excel para generar la presentación corporativa.")

archivo_subido = st.file_uploader("Selecciona el archivo Excel", type=["xlsx", "xlsm"])

if archivo_subido is not None:
    st.info(f"Archivo cargado: **{archivo_subido.name}**")

    if st.button("Generar Presentación PPTX", type="primary"):
        with st.spinner("Procesando datos y formateando diapositivas..."):
            try:
                buffer, nombre_salida = procesar_presentacion(archivo_subido.read())
                st.success("¡Presentación generada con éxito!")

                st.download_button(
                    label="📥 Descargar Presentación",
                    data=buffer,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"Error al generar la presentación: {e}")
