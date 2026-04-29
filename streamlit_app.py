from __future__ import annotations

import re
import unicodedata
from io import BytesIO

import pandas as pd


DIAS_ES = {
    0: "lunes",
    1: "martes",
    2: "miercoles",
    3: "jueves",
    4: "viernes",
    5: "sabado",
    6: "domingo",
}


def normalizar(texto: object) -> str:
    texto = "" if texto is None else str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")


def buscar_columna(columnas: list[str], opciones: list[str], requerida: bool = True) -> str | None:
    mapa = {normalizar(col): col for col in columnas}
    for opcion in opciones:
        encontrada = mapa.get(normalizar(opcion))
        if encontrada:
            return encontrada
    if requerida:
        raise ValueError(f"No se encontro ninguna de estas columnas: {', '.join(opciones)}")
    return None


def buscar_columnas_clase(columnas: list[str]) -> list[str]:
    clases = []
    for col in columnas:
        if re.fullmatch(r"clase_?\d+", normalizar(col)):
            clases.append(col)

    def numero_clase(col: str) -> int:
        match = re.search(r"\d+", str(col))
        return int(match.group()) if match else 999

    return sorted(clases, key=numero_clase)


def formato_hora(valor: object) -> str:
    if pd.isna(valor):
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%H:%M")
    texto = str(valor).strip()
    match = re.search(r"(\d{1,2}):(\d{2})", texto)
    if match:
        return f"{int(match.group(1)):02d}:{match.group(2)}"
    return texto


def valor_es_uno(valor: object) -> bool:
    if pd.isna(valor):
        return False
    if isinstance(valor, str):
        valor = valor.strip().replace(",", ".")
    try:
        return float(valor) == 1
    except (TypeError, ValueError):
        return False


def leer_excel_completo(archivo_excel) -> pd.DataFrame:
    hojas = pd.read_excel(archivo_excel, sheet_name=None)
    datos = []
    for _, df in hojas.items():
        if not df.empty:
            datos.append(df.copy())
    if not datos:
        raise ValueError("El archivo no tiene hojas con datos.")
    return pd.concat(datos, ignore_index=True)


def generar_tabla_final(archivo_excel) -> pd.DataFrame:
    df = leer_excel_completo(archivo_excel)
    columnas = list(df.columns)

    col_actividad = buscar_columna(columnas, ["Actividades_Centro_club", "Actividad", "Nombre actividad"])
    col_fecha = buscar_columna(columnas, ["Fecha_actividad", "Fecha actividad", "Fecha"])
    col_horario = buscar_columna(columnas, ["Horario", "Hora"], requerida=False)
    col_total = buscar_columna(columnas, ["Total", "Total asistencia", "Total asistencias"], requerida=False)
    col_rut = buscar_columna(columnas, ["RUT", "Rut afiliado", "Documento"], requerida=False)
    col_nombre = buscar_columna(columnas, ["Nombre_afiliado", "Nombre afiliado", "Nombre"], requerida=False)
    columnas_clase = buscar_columnas_clase(columnas)

    if not col_total and not columnas_clase:
        raise ValueError("No se encontro la columna Total ni columnas tipo Clase 1, Clase 2, etc.")
    if not col_rut and not col_nombre:
        raise ValueError("No se encontro una columna para identificar personas, por ejemplo RUT o Nombre_afiliado.")

    trabajo = df.copy()
    columnas_base = [col_actividad, col_fecha]
    if col_total:
        columnas_base.append(col_total)
    else:
        columnas_base.extend(columnas_clase)
    trabajo = trabajo.loc[~trabajo[columnas_base].isna().all(axis=1)].copy()
    trabajo = trabajo.loc[
        trabajo[col_actividad].notna() & trabajo[col_fecha].notna()
    ].copy()

    trabajo[col_fecha] = pd.to_datetime(trabajo[col_fecha], errors="coerce").dt.date
    trabajo["Identificador_persona"] = ""
    if col_rut:
        trabajo["Identificador_persona"] = trabajo[col_rut].astype("string").fillna("").str.strip()
    if col_nombre:
        sin_rut = trabajo["Identificador_persona"].eq("")
        trabajo.loc[sin_rut, "Identificador_persona"] = (
            trabajo.loc[sin_rut, col_nombre].astype("string").fillna("").str.strip()
        )
    sin_id = trabajo["Identificador_persona"].eq("")
    trabajo.loc[sin_id, "Identificador_persona"] = "fila_" + (trabajo.index[sin_id] + 2).astype(str)

    if col_total:
        total = pd.to_numeric(trabajo[col_total], errors="coerce")
        trabajo["Presente"] = total.gt(0).astype(int)
        trabajo["Ausente"] = total.eq(0).astype(int)
    else:
        asistencia = trabajo[columnas_clase].map(valor_es_uno)
        trabajo["Presente"] = asistencia.any(axis=1).astype(int)
        trabajo["Ausente"] = asistencia.sum(axis=1).eq(0).astype(int)

    agrupadores = [col_fecha, col_actividad]
    if col_horario:
        agrupadores.append(col_horario)

    persona_clase = (
        trabajo.groupby(agrupadores + ["Identificador_persona"], dropna=False)
        .agg(Presente=("Presente", "max"), Ausente=("Ausente", "max"))
        .reset_index()
    )
    persona_clase.loc[persona_clase["Presente"].eq(1), "Ausente"] = 0

    tabla = (
        persona_clase.groupby(agrupadores, dropna=False)
        .agg(
            Inscritos=("Identificador_persona", "nunique"),
            Presentes=("Presente", "sum"),
            Ausentes=("Ausente", "sum"),
        )
        .reset_index()
    )
    tabla["Porcentaje de asistencia"] = (
        tabla["Presentes"] / tabla["Inscritos"].replace(0, pd.NA) * 100
    ).fillna(0).round().astype(int)
    tabla["Clase"] = tabla[col_actividad]
    tabla["Hora"] = [formato_hora(v) for v in (tabla[col_horario] if col_horario else [""] * len(tabla))]
    tabla["Dia"] = tabla[col_fecha].map(
        lambda fecha: "" if pd.isna(fecha) else DIAS_ES.get(pd.Timestamp(fecha).weekday(), "")
    )

    return tabla[
        ["Clase", "Hora", "Dia", "Inscritos", "Presentes", "Ausentes", "Porcentaje de asistencia"]
    ]


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="Resumen de asistencias", layout="wide")
    st.title("Resumen automatico de asistencias")

    archivo = st.file_uploader("Sube el Excel de asistencia", type=["xlsx", "xls"])

    if archivo is None:
        st.info("Carga un archivo Excel para generar el resumen.")
        st.stop()

    try:
        tabla_final = generar_tabla_final(archivo)
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        st.stop()

    col1, col2, col3 = st.columns(3)
    col1.metric("Inscritos", f"{int(tabla_final['Inscritos'].sum()):,}".replace(",", "."))
    col2.metric("Presentes", f"{int(tabla_final['Presentes'].sum()):,}".replace(",", "."))
    col3.metric("Ausentes", f"{int(tabla_final['Ausentes'].sum()):,}".replace(",", "."))

    st.subheader("Tabla final")
    st.caption("Regla: Total = 0 es Ausente; Total mayor que 0 es Presente.")
    st.dataframe(tabla_final, use_container_width=True, hide_index=True)

    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        tabla_final.to_excel(writer, index=False, sheet_name="Tabla_final")

    st.download_button(
        "Descargar resumen en Excel",
        data=salida.getvalue(),
        file_name="resumen_asistencias.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
