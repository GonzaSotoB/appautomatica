from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

from app_resumen_asistencias import generar_resumen


st.set_page_config(
    page_title="Resumen de asistencias",
    layout="wide",
)

st.title("Resumen automatico de asistencias")

archivo = st.file_uploader(
    "Sube el Excel de asistencia",
    type=["xlsx", "xls"],
)

if archivo is None:
    st.info("Carga un archivo Excel para generar el resumen.")
    st.stop()

try:
    resumen, detalle, formato_mensual = generar_resumen(archivo)
except Exception as exc:
    st.error(f"No se pudo procesar el archivo: {exc}")
    st.stop()

total_inscritos = int(formato_mensual["Inscritos"].sum())
total_asisten = int(formato_mensual["Presentes"].sum())
total_no_asisten = int(formato_mensual["Ausentes"].sum())

col1, col2, col3 = st.columns(3)
col1.metric("Inscritos", f"{total_inscritos:,}".replace(",", "."))
col2.metric("Asisten", f"{total_asisten:,}".replace(",", "."))
col3.metric("No asisten", f"{total_no_asisten:,}".replace(",", "."))

st.subheader("Tabla final")
st.dataframe(formato_mensual, use_container_width=True, hide_index=True)

salida = BytesIO()
with pd.ExcelWriter(salida, engine="openpyxl") as writer:
    formato_mensual.to_excel(writer, index=False, sheet_name="Tabla_final")

st.download_button(
    "Descargar resumen en Excel",
    data=salida.getvalue(),
    file_name="resumen_asistencias.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
