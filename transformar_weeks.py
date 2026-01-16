#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script para transformar el archivo de Brand Franchise / Int Fresh:

- A partir de la columna 'Time' (ej. 'Week Ending 01-05-25'):
    * Agrega columnas: Week, Mes#, Mes name, Mes code, Year
    * Las nuevas columnas se insertan justo después de 'Time'
- Crea una segunda hoja en el Excel con un diccionario:
    * Week No, Week Ending

Formato esperado en Time:
    'Week Ending MM-DD-YY'
    Ejemplo: 'Week Ending 01-05-25' -> 01 = mes, 05 = día, 25 = año (2025)
"""

from pathlib import Path
from time import perf_counter
import pandas as pd

# ==========================
# CONFIGURACIÓN EDITABLE
# ==========================

# Rango de meses a procesar (inclusive). 1 = enero, 12 = diciembre.
# Ejemplo: (1, 6) procesa enero a junio.
MONTH_RANGE = (1, 12)

# Mapa de archivos de entrada por mes. Ajusta las rutas existentes en tu carpeta.
# Solo se procesan los meses cuyo archivo exista en este diccionario y estén dentro de MONTH_RANGE.
INPUT_FILES = {
    1: r"Entrada\\enero.xlsx",
    2: r"Entrada\\febrero.xlsx",
    3: r"Entrada\\marzo.xlsx",
    4: r"Entrada\\abril.xlsx",
    5: r"Entrada\\mayo.xlsx",
    6: r"Entrada\\junio.xlsx",
    7: r"Entrada\\julio.xlsx",
    8: r"Entrada\\agosto.xlsx",
    9: r"Entrada\\septiembre.xlsx",
    10: r"Entrada\\octubre.xlsx",
    11: r"Entrada\\noviembre.xlsx",
    12: r"Entrada\\diciembre.xlsx",
}

# Carpeta de salida y plantilla de nombre de archivo.
# Se infiere el año de la columna Time; si no se logra, se usa YEAR_FALLBACK.
OUTPUT_DIR = Path("Salida")
OUTPUT_TEMPLATE = "SM_CIRCANA_{month:02d}_{year}.xlsx"
YEAR_FALLBACK = 2024

# Hoja a leer: índice (0 = primera hoja) o nombre de la hoja
SHEET_NAME = 0

# Nombre de la columna que contiene el texto tipo 'Week Ending 01-05-25'
TIME_COL = "Time"


def parse_time_to_datetime(time_str: str):
    """Convierte 'Week Ending MM-DD-YY' a datetime; devuelve NaT si no se puede parsear."""
    if pd.isna(time_str):
        return pd.NaT
    date_part = str(time_str).replace("Week Ending", "").strip()
    return pd.to_datetime(date_part, format="%m-%d-%y", errors="coerce")


def print_progress(done: int, total: int, prefix: str = "Progreso"):
    """Imprime una barra de progreso simple en consola."""
    if total <= 0:
        return
    pct = int((done / total) * 100)
    bar_len = 30
    filled = int(bar_len * pct / 100)
    bar = "#" * filled + "-" * (bar_len - filled)
    print(f"\r{prefix}: [{bar}] {pct:3d}%", end="", flush=True)


def parse_week_info(time_str: str) -> pd.Series:
    """
    Recibe un string como 'Week Ending 01-05-25'
    y devuelve Week, Mes#, Mes name, Mes code, Year.
    """
    dt = parse_time_to_datetime(time_str)
    if pd.isna(dt):
        return pd.Series(
            {
                "Week": pd.NA,
                "Mes#": pd.NA,
                "Mes name": pd.NA,
                "Mes code": pd.NA,
                "Year": pd.NA,
            }
        )

    # Número de semana ISO
    week_no = dt.isocalendar().week

    # Número de mes
    mes_num = dt.month

    # Nombre de mes (en inglés: January, February, etc.)
    mes_name = dt.strftime("%B")

    # Código de mes tipo "1. Jan"
    mes_code = f"{mes_num}. {dt.strftime('%b')}"

    year = dt.year

    return pd.Series(
        {
            "Week": week_no,
            "Mes#": mes_num,
            "Mes name": mes_name,
            "Mes code": mes_code,
            "Year": year,
        }
    )


def add_calendar_columns(df: pd.DataFrame, time_col: str = "Time") -> pd.DataFrame:
    """
    Agrega las columnas Week, Mes#, Mes name, Mes code, Year al DataFrame original
    e inserta estas columnas inmediatamente después de la columna 'Time',
    conservando TODAS las columnas originales en su orden.
    """
    if time_col not in df.columns:
        raise KeyError(f"La columna '{time_col}' no existe en el archivo.")

    # Calcular la info de calendario
    week_info = df[time_col].apply(parse_week_info)

    # Unir al DataFrame original (añadimos las columnas nuevas al final de momento)
    df_out = pd.concat([df, week_info], axis=1)

    original_cols = list(df.columns)
    new_cols = ["Week", "Mes#", "Mes name", "Mes code", "Year"]

    # Solo columnas nuevas que realmente existen (por seguridad)
    new_cols = [c for c in new_cols if c in df_out.columns]

    # Construimos el orden:
    # 1. Todas las columnas originales, en orden
    # 2. Justo después de 'Time', insertamos las nuevas columnas
    ordered_cols = []
    for col in original_cols:
        ordered_cols.append(col)
        if col == time_col:
            ordered_cols.extend(new_cols)

    # 3. Si quedara alguna columna extra (por ejemplo, de week_info),
    #    que no esté aún en ordered_cols, la agregamos al final.
    for col in df_out.columns:
        if col not in ordered_cols:
            ordered_cols.append(col)

    df_out = df_out[ordered_cols]

    return df_out


def build_week_dictionary(df: pd.DataFrame, time_col: str = "Time") -> pd.DataFrame:
    """
    Construye un diccionario de semanas con:
        - Week No
        - Week Ending (texto exacto de la columna Time)

    La numeración de Week No viene de la columna 'Week' generada previamente.
    """
    if time_col not in df.columns:
        raise KeyError(f"La columna '{time_col}' no existe en el archivo.")
    if "Week" not in df.columns:
        raise KeyError("La columna 'Week' no existe. Primero debes generarla.")

    tmp = df[[time_col, "Week"]].drop_duplicates().sort_values("Week")

    # Renombrar columnas como pide el usuario
    week_dict = tmp.rename(columns={time_col: "Time", "Week": "Week"})

    # Reordenar columnas por si acaso
    week_dict = week_dict[["Week", "Time"]]

    return week_dict


def infer_month_year(df: pd.DataFrame, time_col: str, default_month: int) -> tuple[int, int]:
    """Obtiene mes y año desde la primera fila válida de Time; usa defaults si falla."""
    valid_times = df[time_col].dropna()
    if not valid_times.empty:
        dt = parse_time_to_datetime(valid_times.iloc[0])
        if not pd.isna(dt):
            return dt.month, dt.year
    return default_month, YEAR_FALLBACK


def process_file_for_month(month: int, input_file: str, sheet_name=SHEET_NAME, time_col: str = TIME_COL) -> Path:
    """Lee, transforma y guarda un archivo; devuelve la ruta de salida."""
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # Agregar columnas de calendario (quedarán detrás de 'Time')
    df_transformed = add_calendar_columns(df, time_col=time_col)

    # Crear diccionario de semanas
    week_dict = build_week_dictionary(df_transformed, time_col=time_col)

    # Determinar mes/año para el nombre de salida
    month_for_name, year_for_name = infer_month_year(df_transformed, time_col=time_col, default_month=month)

    output_path = OUTPUT_DIR / OUTPUT_TEMPLATE.format(month=month_for_name, year=year_for_name)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_transformed.to_excel(writer, sheet_name="Datos", index=False)
        week_dict.to_excel(writer, sheet_name="Diccionario_Semanas", index=False)

    return output_path


def main():
    start_time = perf_counter()
    start_month, end_month = MONTH_RANGE
    months_to_process = [
        m for m in sorted(INPUT_FILES.keys())
        if start_month <= m <= end_month
    ]

    if not months_to_process:
        print("No hay meses configurados en INPUT_FILES dentro del rango indicado.")
        return

    total_months = len(months_to_process)
    processed = 0

    for month in months_to_process:
        input_file = INPUT_FILES[month]
        try:
            output_path = process_file_for_month(month, input_file)
            print(f"[OK] Mes {month:02d}: {input_file} -> {output_path}")
        except FileNotFoundError:
            print(f"[ADVERTENCIA] No se encontró el archivo para mes {month:02d}: {input_file}")
        except Exception as e:
            print(f"[ERROR] Falló el procesamiento de {input_file}: {e}")
        finally:
            processed += 1
            print_progress(processed, total_months, prefix="Procesando meses")

    elapsed = perf_counter() - start_time
    print(f"\nTiempo total de procesamiento: {elapsed:0.2f} segundos")


if __name__ == "__main__":
    main()
