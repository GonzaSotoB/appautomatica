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

cupos = st.number_input("Cupos por actividad", min_value=1, value=40, step=1)

archivo = st.file_uploader(
    "Sube el Excel de asistencia",
    type=["xlsx", "xls"],
)

if archivo is None:
    st.info("Carga un archivo Excel para generar el resumen.")
    st.stop()

try:
    resumen, detalle, formato_mensual = generar_resumen(archivo, cupos_default=int(cupos))
except Exception as exc:
    st.error(f"No se pudo procesar el archivo: {exc}")
    st.stop()

total_inscritos = int(resumen["inscritos"].sum())
total_asisten = int(resumen["total_asisten"].sum())
total_no_asisten = int(resumen["total_no_asisten"].sum())

col1, col2, col3 = st.columns(3)
col1.metric("Inscritos", f"{total_inscritos:,}".replace(",", "."))
col2.metric("Asisten", f"{total_asisten:,}".replace(",", "."))
col3.metric("No asisten", f"{total_no_asisten:,}".replace(",", "."))

st.subheader("Formato mensual")
st.dataframe(formato_mensual, use_container_width=True, hide_index=True)

st.subheader("Resumen de control")
st.dataframe(resumen, use_container_width=True, hide_index=True)

with st.expander("Ver detalle persona por persona"):
    st.dataframe(detalle, use_container_width=True, hide_index=True)

salida = BytesIO()
with pd.ExcelWriter(salida, engine="openpyxl") as writer:
    formato_mensual.to_excel(writer, index=False, sheet_name="Formato_mensual")
    resumen.to_excel(writer, index=False, sheet_name="Resumen")
    detalle.to_excel(writer, index=False, sheet_name="Detalle_clasificacion")

st.download_button(
    "Descargar resumen en Excel",
    data=salida.getvalue(),
    file_name="resumen_asistencias.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
