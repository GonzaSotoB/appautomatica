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

COLUMNAS_ESTANDAR_POR_POSICION = [
    "N",
    "Nombre_afiliado",
    "RUT",
    "Telefono",
    "Email_afiliado",
    "Actividades_Centro_club",
    "Fecha_actividad",
    "Horario",
    "sucursal",
    "mes_oferta",
    "Clase 1",
    "Clase 2",
    "Clase 3",
    "Clase 4",
    "Total",
    "%",
]

ORDEN_PREFACTURA = {
    "Virtual": [
        "Gimnasia funcional lunes 9:00",
        "Mindfulness avanzado lunes 9:00",
        "Ágilmente lunes 9:00",
        "Entrenando la mente lunes 11:00",
        "Autoestima y Conocimiento",
        "Inglés Intermedio lunes 14:00",
        "Tarot lunes 16:00",
        "Juegos cognitivos martes 9:00",
        "Entrenamiento funcional senior avanzado martes 9:00",
        "Técnicas de relajación martes 11:00",
        "Uso de Redes Sociales martes 11:00",
        "Baile entretenido martes 11:00",
        "Mente activa martes 11:00",
        "Inglés martes 14:00",
        "Yoga martes 16:00",
        "Cocina entretenida martes 16:00",
        "Taller de Autobiografía martes 16:00",
        "Entrenamiento funcional senior inicial miércoles 9:00",
        "Aprende sobre inteligencia artificial miércoles 9:00",
        "Yoga miércoles 11:00",
        "Reflexología miércoles 11:00",
        "Gimnasia funcional miércoles 11:00",
        "Mindfulness inicial miércoles 14:00",
        "Frances Inicial miércoles 14:00",
        "Amasandería miércoles 16:00",
        "Aromaterapia miércoles 16:00",
        "Jardineria y bienestar miércoles 16:00",
        "Ajedrez inicial jueves 09:00",
        "Emociones y Resolución de conflictos jueves 09:00",
        "Ágilmente jueves 11:00",
        "Viaja sin salir de tu casa jueves 11:00",
        "Gimnasia funcional jueves 11:00",
        "Introducción a la pintura jueves 14:00",
        "Bienestar emocional jueves 16:00",
        "Tejido a Palillo jueves 16:00",
        "Astrología jueves 16:00",
        "Entrenando la mente viernes 9:00",
        "Gimnasia funcional viernes 9:00",
        "Baile entretenido viernes 11:00",
        "Habilidades Digitales para la vida diaria viernes 11:00",
        "Tejido viernes 11:00",
        "Pintura intermedia viernes 14:00",
        "Uso de celulares Smartphone viernes 16:00",
    ],
    "Alonso de Ovalle": [
        "Baile entretenido lunes 10:00",
        "Gimnasia kinésica lunes 11:15",
        "Entrenamiento funcional lunes 11:15",
        "Ágilmente lunes 11:15",
        "Gimnasia kinésica lunes 12:30",
        "Crochet avanzado lunes 12:30",
        "Gimnasia kinésica lunes 15:30",
        "Yoga lunes 15:30",
        "Gimnasia kinésica martes 10:00",
        "Pintura martes 10:00",
        "Bailes tropicales martes 10:00",
        "Baile entretenido martes 11:15",
        "Aprende a Usar  tu sucursal Virtual martes 11:15",
        "Folklore chileno martes 11:15",
        "Teatro: mente, cuerpo y creatividad martes 11:15",
        "Inglés martes 12:30",
        "Gimnasia kinésica martes 12:30",
        "Pilates martes 12:30",
        "Taichi martes 15:30",
        "Chikung Terapeútico martes 15:30",
        "Gimnasia kinésica miércoles 10:00",
        "Uso de Whastapp miércoles 10:00",
        "Gimnasia kinésica miércoles 11:15",
        "Uso de Smartphone miércoles 11:15",
        "Danzas del mundo miércoles 11:15",
        "Gimnasia kinésica miércoles 12:30",
        "Tango miércoles 12:30",
        "Baile entretenido miércoles 15:30",
        "Gimnasia kinésica jueves 10:00",
        "Manos creativas: Arte Terapeútico jueves 11:15",
        "Gimnasia kinésica jueves 11:15",
        "Yoga jueves 11:15",
        "Gimnasia kinésica jueves 12:30",
        "Pilates jueves 12:30",
        "Taichi jueves 15:30",
        "Ágilmente jueves 15:30",
        "Baile entretenido viernes 10:00",
        "Gimnasia kinésica viernes 11:15",
        "Jueguemos Ajedrez viernes 11:15",
        "Entrenamiento funcional viernes 12:00",
        "Yoga viernes 12:00",
    ],
    "Ñuñoa": [
        "Baile entretenido lunes 10:00",
        "Viaje a través de la literatura lunes 11:15",
        "Uso de smartphone lunes 12:00",
        "Yoga lunes 15:15",
        "Gimnasia kinésica martes 10:00",
        "Agilmente martes 15:15",
        "Agilmente miércoles 10:00",
        "Entrenamiento Funcional miércoles 11:15",
        "Baile entretenido miércoles 15:15",
        "Gimnasia kinésica jueves 10:00",
        "Inglés Avanzado jueves 11:15",
        "Taichi jueves intermedio 12:00",
        "Taichi inicial jueves 15:15",
        "Baile entretenido viernes 10:00",
        "Yoga viernes 13:00",
    ],
    "La Florida": [
        "Gimnasia kinésica lunes 10:00",
        "Ágilmente lunes 11:15",
        "Entrenamiento funcional martes 12:00",
        "Gimnasia kinésica miércoles 10:00",
        "Baile entretenido viernes 10:00",
        "Gimnasia kinésica viernes 11:15",
    ],
    "Puente Alto": [
        "Baile entretenido lunes 11:15",
        "Gimnasia kinésica martes 15:00",
        "Uso de Smartphone miércoles 15:00",
        "Entrenamiento funcional jueves 12:00",
        "Baile entretenido jueves 15:00",
        "Gimnasia kinésica viernes 10:00",
    ],
    "San Bernardo": [
        "Agilmente martes 15:15",
        "Gimnasia kinésica miércoles 15:15",
        "Baile entretenido jueves 15:15",
    ],
    "Buin": [
        "Gimnasia kinésica martes 10:00",
        "Baile entretenido jueves 10:00",
    ],
    "Melipilla": [
        "Baile entretenido lunes 10:00",
        "Yoga miércoles 10:00",
        "Inglés Intermedio jueves 10:00",
        "Pilates viernes 10:00",
    ],
    "San Miguel": [
        "Gimnasia kinésica lunes 15:15",
        "Baile entretenido miércoles 10:00",
        "Ágilmente jueves 10:00",
    ],
}


def normalizar(texto: object) -> str:
    texto = "" if texto is None else str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower().strip()
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")


def clave_orden(texto: object) -> str:
    texto = "" if texto is None else str(texto)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.lower().strip()
    texto = re.sub(r"\b(\d{1,2}):(\d{2})\b", lambda m: f"{int(m.group(1)):02d}:{m.group(2)}", texto)
    texto = re.sub(r"[^a-z0-9:]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def clave_orden_compacta(texto: object) -> str:
    return clave_orden(texto).replace(" ", "")


ORDEN_SUCURSALES = {clave_orden(sucursal): i for i, sucursal in enumerate(ORDEN_PREFACTURA)}
ORDEN_CLASES = {
    clave_orden(sucursal): {
        clave_orden(clase): i for i, clase in enumerate(clases)
    }
    for sucursal, clases in ORDEN_PREFACTURA.items()
}
ORDEN_CLASES_COMPACTO = {
    clave_orden(sucursal): {
        clave_orden_compacta(clase): i for i, clase in enumerate(clases)
    }
    for sucursal, clases in ORDEN_PREFACTURA.items()
}


def buscar_columna(columnas: list[str], opciones: list[str], requerida: bool = True) -> str | None:
    mapa = {normalizar(col): col for col in columnas}
    for opcion in opciones:
        encontrada = mapa.get(normalizar(opcion))
        if encontrada:
            return encontrada
    if requerida:
        raise ValueError(f"No se encontro ninguna de estas columnas: {', '.join(opciones)}")
    return None


def sugerir_columna(columnas: list[str], palabras_clave: list[str]) -> str | None:
    for col in columnas:
        nombre = normalizar(col)
        if any(palabra in nombre for palabra in palabras_clave):
            return col
    return None


def sugerir_columna_actividad(columnas: list[str]) -> str | None:
    candidatas = []
    for col in columnas:
        nombre = normalizar(col)
        if any(excluir in nombre for excluir in ["fecha", "hora", "horario", "total", "rut", "nombre"]):
            continue
        if any(palabra in nombre for palabra in ["actividad", "actividades", "activi", "curso", "taller"]):
            candidatas.append(col)
    return candidatas[0] if candidatas else None


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


def estandarizar_columnas_por_posicion(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    columnas_nuevas = []
    usados = set()
    for indice, columna in enumerate(df.columns):
        if indice < len(COLUMNAS_ESTANDAR_POR_POSICION):
            nombre = COLUMNAS_ESTANDAR_POR_POSICION[indice]
        else:
            nombre = str(columna)

        if nombre in usados:
            contador = 2
            nombre_base = nombre
            while f"{nombre_base}_{contador}" in usados:
                contador += 1
            nombre = f"{nombre_base}_{contador}"

        usados.add(nombre)
        columnas_nuevas.append(nombre)

    df.columns = columnas_nuevas
    return df


def generar_tabla_final(
    archivo_excel,
    col_actividad: str | None = None,
    col_fecha: str | None = None,
    col_horario: str | None = None,
    col_total: str | None = None,
    col_persona_usuario: str | None = None,
) -> pd.DataFrame:
    df = estandarizar_columnas_por_posicion(leer_excel_completo(archivo_excel))
    columnas = list(df.columns)

    col_actividad = col_actividad or buscar_columna(
        columnas,
        ["Actividades_Centro_club", "activi", "Actividad", "Nombre actividad"],
        requerida=False,
    ) or sugerir_columna_actividad(columnas)
    col_fecha = col_fecha or buscar_columna(
        columnas,
        ["Fecha_actividad", "Fecha actividad", "Fecha"],
        requerida=False,
    ) or sugerir_columna(columnas, ["fecha"])
    col_sucursal = buscar_columna(columnas, ["sucursal", "sede", "centro"], requerida=False)
    col_horario = col_horario or buscar_columna(columnas, ["Horario", "Hora"], requerida=False)
    col_total = col_total or buscar_columna(columnas, ["Total", "Total asistencia", "Total asistencias"], requerida=False)
    col_rut = buscar_columna(columnas, ["RUT", "Rut afiliado", "Documento"], requerida=False)
    col_nombre = buscar_columna(columnas, ["Nombre_afiliado", "Nombre afiliado", "Nombre"], requerida=False)
    col_persona = col_persona_usuario or col_rut or col_nombre
    columnas_clase = buscar_columnas_clase(columnas)

    if not col_actividad:
        raise ValueError("No se pudo identificar la columna de actividad.")
    if not col_fecha:
        raise ValueError("No se pudo identificar la columna de fecha.")
    if not col_total and not columnas_clase:
        raise ValueError("No se encontro la columna Total ni columnas tipo Clase 1, Clase 2, etc.")
    if not col_persona:
        raise ValueError("No se encontro una columna para identificar personas, por ejemplo RUT o Nombre.")

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

    trabajo["_fecha_actividad"] = pd.to_datetime(trabajo[col_fecha], errors="coerce").dt.date
    trabajo["_actividad"] = trabajo[col_actividad]
    trabajo["_sucursal"] = trabajo[col_sucursal] if col_sucursal else ""
    trabajo["_horario"] = trabajo[col_horario] if col_horario else ""
    trabajo["Identificador_persona"] = trabajo[col_persona].astype("string").fillna("").str.strip()
    if col_persona_usuario is None and col_rut and col_nombre:
        sin_rut = trabajo["Identificador_persona"].eq("")
        trabajo.loc[sin_rut, "Identificador_persona"] = trabajo.loc[sin_rut, col_nombre].astype("string").fillna("").str.strip()
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

    agrupadores = ["_sucursal", "_fecha_actividad", "_actividad", "_horario"]

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
    tabla["Clase"] = tabla["_actividad"]
    tabla["Sucursal"] = tabla["_sucursal"]
    tabla["Hora"] = [formato_hora(v) for v in tabla["_horario"]]
    tabla["Dia"] = tabla["_fecha_actividad"].map(
        lambda fecha: "" if pd.isna(fecha) else DIAS_ES.get(pd.Timestamp(fecha).weekday(), "")
    )
    tabla["_orden_sucursal"] = tabla["Sucursal"].map(lambda s: ORDEN_SUCURSALES.get(clave_orden(s), 999))
    tabla["_clave_clase"] = (
        tabla["Clase"].astype(str) + " " + tabla["Dia"].astype(str) + " " + tabla["Hora"].astype(str)
    ).map(clave_orden)
    tabla["_orden_clase"] = [
        ORDEN_CLASES.get(clave_orden(sucursal), {}).get(
            clase,
            ORDEN_CLASES_COMPACTO.get(clave_orden(sucursal), {}).get(clase.replace(" ", ""), 9999),
        )
        for sucursal, clase in zip(tabla["Sucursal"], tabla["_clave_clase"])
    ]
    tabla = tabla.sort_values(
        by=["_orden_sucursal", "Sucursal", "_orden_clase", "Dia", "Hora", "Clase"],
        kind="stable",
    )

    return tabla[
        ["Sucursal", "Clase", "Hora", "Dia", "Inscritos", "Presentes", "Ausentes", "Porcentaje de asistencia"]
    ]


def detectar_columnas(archivo_excel) -> tuple[pd.DataFrame, dict[str, str | None]]:
    df = estandarizar_columnas_por_posicion(leer_excel_completo(archivo_excel))
    columnas = list(df.columns)
    deteccion = {
        "actividad": buscar_columna(columnas, ["Actividades_Centro_club", "activi", "Actividad", "Nombre actividad"], requerida=False)
        or sugerir_columna_actividad(columnas),
        "fecha": buscar_columna(columnas, ["Fecha_actividad", "Fecha actividad", "Fecha"], requerida=False)
        or sugerir_columna(columnas, ["fecha"]),
        "hora": buscar_columna(columnas, ["Horario", "Hora"], requerida=False),
        "total": buscar_columna(columnas, ["Total", "Total asistencia", "Total asistencias"], requerida=False),
        "persona": buscar_columna(columnas, ["RUT", "Rut afiliado", "Documento"], requerida=False)
        or buscar_columna(columnas, ["Nombre_afiliado", "Nombre afiliado", "Nombre"], requerida=False),
    }
    return df, deteccion


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="Resumen de asistencias", layout="wide")
    st.title("Resumen automatico de asistencias")

    archivo = st.file_uploader("Sube el Excel de asistencia", type=["xlsx", "xls"])

    if archivo is None:
        st.info("Carga un archivo Excel para generar el resumen.")
        st.stop()

    try:
        df_preview, deteccion = detectar_columnas(archivo)
    except Exception as exc:
        st.error(f"No se pudo procesar el archivo: {exc}")
        st.stop()

    columnas = list(df_preview.columns)
    opciones_requeridas = columnas
    opciones_opcionales = [""] + columnas

    with st.expander("Revisar o cambiar columnas detectadas", expanded=False):
        st.write("Si una columna viene con otro nombre, elige aqui cual corresponde.")
        col_actividad = st.selectbox(
            "Columna de clase / actividad",
            opciones_requeridas,
            index=opciones_requeridas.index(deteccion["actividad"]) if deteccion["actividad"] in opciones_requeridas else 0,
        )
        col_fecha = st.selectbox(
            "Columna de fecha",
            opciones_requeridas,
            index=opciones_requeridas.index(deteccion["fecha"]) if deteccion["fecha"] in opciones_requeridas else 0,
        )
        col_horario = st.selectbox(
            "Columna de hora",
            opciones_opcionales,
            index=opciones_opcionales.index(deteccion["hora"]) if deteccion["hora"] in opciones_opcionales else 0,
        )
        col_total = st.selectbox(
            "Columna total de asistencia",
            opciones_opcionales,
            index=opciones_opcionales.index(deteccion["total"]) if deteccion["total"] in opciones_opcionales else 0,
        )
        col_persona = st.selectbox(
            "Columna para identificar persona",
            opciones_requeridas,
            index=opciones_requeridas.index(deteccion["persona"]) if deteccion["persona"] in opciones_requeridas else 0,
        )

    try:
        archivo.seek(0)
        tabla_final = generar_tabla_final(
            archivo,
            col_actividad=col_actividad,
            col_fecha=col_fecha,
            col_horario=col_horario or None,
            col_total=col_total or None,
            col_persona_usuario=col_persona,
        )
    except Exception as exc:
        st.error(f"No se pudo generar la tabla: {exc}")
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
