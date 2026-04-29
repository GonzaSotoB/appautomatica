from __future__ import annotations

import re
import unicodedata
from pathlib import Path

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
        nombre = normalizar(col)
        if re.fullmatch(r"clase_?\d+", nombre):
            clases.append(col)

    def numero_clase(col: str) -> int:
        match = re.search(r"\d+", str(col))
        return int(match.group()) if match else 999

    return sorted(clases, key=numero_clase)


def valor_es_uno(valor: object) -> bool:
    if pd.isna(valor):
        return False
    if isinstance(valor, str):
        valor = valor.strip().replace(",", ".")
    try:
        return float(valor) == 1
    except (TypeError, ValueError):
        return False


def valor_es_cero(valor: object) -> bool:
    if pd.isna(valor):
        return False
    if isinstance(valor, str):
        valor = valor.strip().replace(",", ".")
    try:
        return float(valor) == 0
    except (TypeError, ValueError):
        return False


def formato_hora(valor: object) -> str:
    if pd.isna(valor):
        return ""
    if hasattr(valor, "strftime"):
        return valor.strftime("%H:%M")
    texto = str(valor).strip()
    if not texto:
        return ""
    match = re.search(r"(\d{1,2}):(\d{2})", texto)
    if match:
        return f"{int(match.group(1)):02d}:{match.group(2)}"
    return texto


def nombre_reporte(actividad: object, fecha: object, horario: object) -> str:
    actividad_txt = str(actividad).strip().capitalize()
    if pd.isna(fecha):
        dia = ""
    else:
        dia = DIAS_ES.get(pd.Timestamp(fecha).weekday(), "")
    hora = formato_hora(horario)
    partes = [actividad_txt, dia, hora]
    return " ".join(parte for parte in partes if parte).strip()


def leer_excel_completo(archivo_excel) -> pd.DataFrame:
    hojas = pd.read_excel(archivo_excel, sheet_name=None)
    datos = []
    for nombre_hoja, df in hojas.items():
        if df.empty:
            continue
        df = df.copy()
        df["Hoja_origen"] = nombre_hoja
        datos.append(df)

    if not datos:
        raise ValueError("El archivo no tiene hojas con datos.")

    return pd.concat(datos, ignore_index=True)


def generar_resumen(archivo_excel) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = leer_excel_completo(archivo_excel)
    columnas = list(df.columns)

    col_actividad = buscar_columna(columnas, ["Actividades_Centro_club", "Actividad", "Nombre actividad"])
    col_fecha = buscar_columna(columnas, ["Fecha_actividad", "Fecha actividad", "Fecha"])
    col_sucursal = buscar_columna(columnas, ["sucursal", "sede", "centro"], requerida=False)
    col_horario = buscar_columna(columnas, ["Horario", "Hora"], requerida=False)
    col_total = buscar_columna(columnas, ["Total", "Total asistencia", "Total asistencias"], requerida=False)
    col_rut = buscar_columna(columnas, ["RUT", "Rut afiliado", "Documento"], requerida=False)
    col_nombre = buscar_columna(columnas, ["Nombre_afiliado", "Nombre afiliado", "Nombre"], requerida=False)
    columnas_clase = buscar_columnas_clase(columnas)

    if not columnas_clase:
        raise ValueError("No se encontraron columnas de asistencia tipo 'Clase 1', 'Clase 2', etc.")

    if not col_rut and not col_nombre:
        raise ValueError("No se encontro una columna para identificar personas, por ejemplo RUT o Nombre_afiliado.")

    col_persona = "Identificador_persona"
    trabajo = df.copy()
    filas_vacias = (
        trabajo[[col_actividad, col_fecha] + columnas_clase]
        .isna()
        .all(axis=1)
    )
    trabajo = trabajo.loc[~filas_vacias].copy()
    trabajo[col_fecha] = pd.to_datetime(trabajo[col_fecha], errors="coerce").dt.date
    trabajo[col_persona] = ""
    if col_rut:
        trabajo[col_persona] = trabajo[col_rut].astype("string").fillna("").str.strip()
    if col_nombre:
        sin_rut = trabajo[col_persona].eq("")
        trabajo.loc[sin_rut, col_persona] = (
            trabajo.loc[sin_rut, col_nombre].astype("string").fillna("").str.strip()
        )
    sin_identificador = trabajo[col_persona].eq("")
    trabajo.loc[sin_identificador, col_persona] = (
        "fila_" + (trabajo.index[sin_identificador] + 2).astype(str)
    )

    asistencia = trabajo[columnas_clase].map(valor_es_uno)
    ceros = trabajo[columnas_clase].map(valor_es_cero)
    if col_total:
        total_asistencia = pd.to_numeric(trabajo[col_total], errors="coerce")
        trabajo["Asiste"] = total_asistencia.gt(0).astype(int)
        trabajo["No_asiste"] = total_asistencia.eq(0).astype(int)
        trabajo["Ausente_0000"] = trabajo["No_asiste"]
        trabajo["Asistencias_mes"] = total_asistencia.fillna(0)
    else:
        trabajo["Asiste"] = asistencia.any(axis=1).astype(int)
        trabajo["Ausente_0000"] = ceros.all(axis=1).astype(int)
        trabajo["No_asiste"] = ((trabajo["Asiste"].eq(0)) & (trabajo["Ausente_0000"].eq(1))).astype(int)
        trabajo["Asistencias_mes"] = asistencia.sum(axis=1)
    trabajo["Clasificacion_asistencia"] = "Sin clasificar"
    trabajo.loc[trabajo["Asiste"].eq(1), "Clasificacion_asistencia"] = "Presente"
    trabajo.loc[trabajo["No_asiste"].eq(1), "Clasificacion_asistencia"] = "Ausente"

    agrupadores = []
    if col_sucursal:
        agrupadores.append(col_sucursal)
    agrupadores.extend([col_fecha, col_actividad])

    persona_actividad = (
        trabajo.groupby(agrupadores + [col_persona], dropna=False)
        .agg(
            Asiste=("Asiste", "max"),
            Ausente_0000=("Ausente_0000", "min"),
            Asistencias_mes=("Asistencias_mes", "sum"),
        )
        .reset_index()
    )
    persona_actividad["No_asiste"] = (
        persona_actividad["Asiste"].eq(0) & persona_actividad["Ausente_0000"].eq(1)
    ).astype(int)

    resumen = (
        persona_actividad.groupby(agrupadores, dropna=False)
        .agg(
            inscritos=(col_persona, "nunique"),
            total_asisten=("Asiste", "sum"),
            total_no_asisten=("No_asiste", "sum"),
            ausentes_0000=("Ausente_0000", "sum"),
            asistencias_mes=("Asistencias_mes", "sum"),
        )
        .reset_index()
    )
    resumen["porcentaje_asistencia"] = (
        resumen["total_asisten"] / resumen["inscritos"]
    ).fillna(0)
    resumen["columnas_asistencia_usadas"] = ", ".join(map(str, columnas_clase))

    agrupadores_mensual = []
    if col_sucursal:
        agrupadores_mensual.append(col_sucursal)
    agrupadores_mensual.extend([col_fecha, col_actividad])
    if col_horario:
        agrupadores_mensual.append(col_horario)

    persona_mensual = (
        trabajo.groupby(agrupadores_mensual + [col_persona], dropna=False)
        .agg(
            Asiste=("Asiste", "max"),
            Ausente_0000=("Ausente_0000", "min"),
        )
        .reset_index()
    )
    persona_mensual["No_asiste"] = (
        persona_mensual["Asiste"].eq(0) & persona_mensual["Ausente_0000"].eq(1)
    ).astype(int)

    mensual_base = (
        persona_mensual.groupby(agrupadores_mensual, dropna=False)
        .agg(
            Inscritos=(col_persona, "nunique"),
            Presentes=("Asiste", "sum"),
            Ausentes=("No_asiste", "sum"),
        )
        .reset_index()
    )
    mensual_base["Prom. % de asistencia"] = (
        mensual_base["Presentes"] / mensual_base["Inscritos"].replace(0, pd.NA) * 100
    ).fillna(0).round().astype(int)
    mensual_base["Fecha actividad"] = mensual_base[col_fecha]
    mensual_base["Dia"] = mensual_base[col_fecha].map(
        lambda fecha: "" if pd.isna(fecha) else DIAS_ES.get(pd.Timestamp(fecha).weekday(), "")
    )
    mensual_base["Horario"] = [
        formato_hora(valor) for valor in (mensual_base[col_horario] if col_horario else [""] * len(mensual_base))
    ]
    mensual_base["Clase"] = mensual_base[col_actividad]
    columnas_formato = [
        "Clase",
        "Horario",
        "Dia",
        "Inscritos",
        "Presentes",
        "Ausentes",
        "Prom. % de asistencia",
    ]
    formato_mensual = mensual_base[columnas_formato].rename(
        columns={
            "Horario": "Hora",
            "Prom. % de asistencia": "Porcentaje de asistencia",
        }
    )

    columnas_detalle = agrupadores + [col_persona]
    if col_nombre and col_nombre not in columnas_detalle:
        columnas_detalle.append(col_nombre)
    detalle = trabajo[
        columnas_detalle
        + columnas_clase
        + ["Clasificacion_asistencia", "Asiste", "No_asiste", "Ausente_0000"]
    ].copy()

    return resumen, detalle, formato_mensual


def guardar_resultado(ruta_excel: str | Path, ruta_salida: str | Path | None = None) -> Path:
    ruta_excel = Path(ruta_excel)
    if ruta_salida is None:
        ruta_salida = ruta_excel.with_name(f"{ruta_excel.stem}_RESUMEN.xlsx")
    ruta_salida = Path(ruta_salida)

    resumen, detalle, formato_mensual = generar_resumen(ruta_excel)
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        formato_mensual.to_excel(writer, index=False, sheet_name="Formato_mensual")
        ws = writer.sheets["Formato_mensual"]
        for cell in ws[1]:
            cell.style = "Headline 3"
        for column_cells in ws.columns:
            width = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 45)
            ws.column_dimensions[column_cells[0].column_letter].width = width

    return ruta_salida


def main() -> None:
    from tkinter import Tk, filedialog, messagebox

    root = Tk()
    root.withdraw()

    ruta_excel = filedialog.askopenfilename(
        title="Selecciona el Excel de asistencia",
        filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
    )
    if not ruta_excel:
        return

    try:
        ruta_salida = guardar_resultado(ruta_excel)
    except Exception as exc:
        messagebox.showerror("No se pudo generar el resumen", str(exc))
        raise

    messagebox.showinfo(
        "Resumen generado",
        f"Listo. Se creo el archivo:\n{ruta_salida}",
    )


if __name__ == "__main__":
    main()
