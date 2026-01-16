import pandas as pd
import numpy as np
from pathlib import Path

# ----------------------------------------------------------
# CONFIGURACIÓN BÁSICA
# ----------------------------------------------------------
INPUT_FILE = "Salida/00. SM-SourceOfTruth.xlsx"        # <- Cambia por el nombre de tu archivo
SHEET_NAME = 0                         # Hoja a leer (0 = primera hoja)
OUTPUT_FILE = "reporte_calidad_datos.xlsx"

# Si tu archivo está en otra carpeta, puedes poner la ruta completa, ej:
# INPUT_FILE = r"C:\ruta\a\tu\archivo.xlsx"

# ----------------------------------------------------------
# FUNCIONES AUXILIARES
# ----------------------------------------------------------

def limpiar_nombre_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Elimina espacios al inicio/fin y normaliza algunos caracteres.
    """
    df = df.copy()
    df.columns = df.columns.str.strip()
    return df

def a_numerico(col: pd.Series) -> pd.Series:
    """
    Intenta convertir una serie a numérico:
    - Quita $, comas, %, espacios
    - Devuelve float con NaN donde no se pueda convertir
    """
    return (
        col.astype(str)
        .str.replace(r"[\$,]", "", regex=True)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace({"": np.nan, "nan": np.nan, "None": np.nan})
        .astype(float)
    )

def inferir_tipo(col: pd.Series) -> str:
    """
    Devuelve un tipo sugerido simple para el análisis.
    """
    if pd.api.types.is_numeric_dtype(col):
        return "numérico"
    if pd.api.types.is_datetime64_any_dtype(col):
        return "fecha"
    # intento de conversión a numérico
    try:
        _ = pd.to_numeric(col.dropna().head(20).str.replace(r"[\$,]", "", regex=True), errors="raise")
        return "texto (numérico en texto)"
    except Exception:
        return "texto"

def resumen_columna(col: pd.Series) -> dict:
    """
    Genera un resumen de calidad por columna.
    """
    total = len(col)
    n_null = col.isna().sum()
    n_non_null = total - n_null
    n_unique = col.nunique(dropna=True)

    info = {
        "non_null": n_non_null,
        "nulls": n_null,
        "pct_nulls": round(n_null * 100 / total, 2) if total > 0 else np.nan,
        "unique_values": n_unique,
        "sample_values": ", ".join(map(str, col.dropna().unique()[:5])),
        "inferred_type": inferir_tipo(col),
    }

    # Stats numéricas si aplica
    if pd.api.types.is_numeric_dtype(col):
        info.update(
            {
                "min": col.min(),
                "max": col.max(),
                "mean": col.mean(),
                "std": col.std(),
            }
        )
    else:
        info.update({"min": np.nan, "max": np.nan, "mean": np.nan, "std": np.nan})

    return info

def analizar_duplicados(df: pd.DataFrame) -> pd.DataFrame:
    """
    Analiza duplicados a nivel de fila completa.
    """
    total = len(df)
    dup_mask = df.duplicated(keep=False)
    n_dup_rows = dup_mask.sum()
    pct_dup = round(n_dup_rows * 100 / total, 2) if total > 0 else 0

    resumen = pd.DataFrame(
        {
            "metrica": ["filas_totales", "filas_duplicadas", "%_filas_duplicadas"],
            "valor": [total, n_dup_rows, pct_dup],
        }
    )

    # Se devuelven también los detalles de duplicados (limitando a 200 filas)
    detalles = df[dup_mask].copy()
    if len(detalles) > 200:
        detalles = detalles.head(200)
    return resumen, detalles

def crear_registro_regla(nombre_regla, descripcion, n_errores, total_filas):
    pct = round(n_errores * 100 / total_filas, 2) if total_filas > 0 else np.nan
    return {
        "Regla": nombre_regla,
        "Descripción": descripcion,
        "Filas con error": n_errores,
        "% filas con error": pct,
    }

# ----------------------------------------------------------
# CARGA DE DATOS
# ----------------------------------------------------------

print(f"Leyendo archivo: {INPUT_FILE}")

if not Path(INPUT_FILE).is_file():
    raise FileNotFoundError(f"No se encontró el archivo: {INPUT_FILE}")

df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
df = limpiar_nombre_columnas(df)

total_filas = len(df)
print(f"Filas leídas: {total_filas}")
print(f"Columnas: {len(df.columns)}")

# ----------------------------------------------------------
# RESUMEN POR COLUMNA
# ----------------------------------------------------------

resumen_cols = []
for colname in df.columns:
    info = resumen_columna(df[colname])
    info["columna"] = colname
    resumen_cols.append(info)

df_resumen_columnas = pd.DataFrame(resumen_cols)[
    [
        "columna",
        "inferred_type",
        "non_null",
        "nulls",
        "pct_nulls",
        "unique_values",
        "min",
        "max",
        "mean",
        "std",
        "sample_values",
    ]
].sort_values("columna")

# ----------------------------------------------------------
# ANÁLISIS DE DUPLICADOS
# ----------------------------------------------------------

df_resumen_dup, df_detalle_dup = analizar_duplicados(df)

# ----------------------------------------------------------
# REGLAS DE NEGOCIO / CONSISTENCIA
# ----------------------------------------------------------

reglas_resumen = []
detalles_errores = {}

# --- 1. Consistencia del precio actual: Price per Unit ≈ Dollar Sales / Unit Sales

cols_precio = ["Dollar Sales", "Unit Sales", "Price per Unit"]
for c in cols_precio:
    if c not in df.columns:
        print(f"Aviso: columna '{c}' no encontrada. Se omite la regla de Precio actual.")
        cols_precio = []
        break

if cols_precio:
    sales_num = a_numerico(df["Dollar Sales"])
    units_num = a_numerico(df["Unit Sales"])
    price_num = a_numerico(df["Price per Unit"])

    mask_valido = (~sales_num.isna()) & (~units_num.isna()) & (units_num != 0) & (~price_num.isna())
    precio_calc = sales_num / units_num

    # error relativo absoluto
    error_rel = (precio_calc - price_num).abs() / price_num.replace(0, np.nan)
    # tolerancia 15%
    mask_error = mask_valido & (error_rel > 0.15)

    n_errores = mask_error.sum()

    reglas_resumen.append(
        crear_registro_regla(
            "Precio_actual_vs_ventas",
            "Price per Unit ≈ Dollar Sales / Unit Sales (tolerancia 15%)",
            n_errores,
            total_filas,
        )
    )

    detalles_errores["Precio_actual_vs_ventas"] = pd.DataFrame(
        {
            "Dollar Sales": df["Dollar Sales"],
            "Unit Sales": df["Unit Sales"],
            "Price per Unit": df["Price per Unit"],
            "Precio_calc_DollarSales_div_UnitSales": precio_calc,
            "Error_relativo": error_rel,
        }
    )[mask_error].head(300)

# --- 2. Consistencia del precio año anterior: Price per Unit Year Ago ≈ Dollar Sales Year Ago / Unit Sales Year Ago

cols_precio_ya = ["Dollar Sales Year Ago", "Unit Sales Year Ago", "Price per Unit Year Ago"]
for c in cols_precio_ya:
    if c not in df.columns:
        print(f"Aviso: columna '{c}' no encontrada. Se omite la regla de Precio año anterior.")
        cols_precio_ya = []
        break

if cols_precio_ya:
    sales_ya = a_numerico(df["Dollar Sales Year Ago"])
    units_ya = a_numerico(df["Unit Sales Year Ago"])
    price_ya = a_numerico(df["Price per Unit Year Ago"])

    mask_valido_ya = (~sales_ya.isna()) & (~units_ya.isna()) & (units_ya != 0) & (~price_ya.isna())
    precio_calc_ya = sales_ya / units_ya

    error_rel_ya = (precio_calc_ya - price_ya).abs() / price_ya.replace(0, np.nan)
    mask_error_ya = mask_valido_ya & (error_rel_ya > 0.15)

    n_errores_ya = mask_error_ya.sum()

    reglas_resumen.append(
        crear_registro_regla(
            "Precio_ya_vs_ventas_ya",
            "Price per Unit Year Ago ≈ Dollar Sales Year Ago / Unit Sales Year Ago (tolerancia 15%)",
            n_errores_ya,
            total_filas,
        )
    )

    detalles_errores["Precio_ya_vs_ventas_ya"] = pd.DataFrame(
        {
            "Dollar Sales Year Ago": df["Dollar Sales Year Ago"],
            "Unit Sales Year Ago": df["Unit Sales Year Ago"],
            "Price per Unit Year Ago": df["Price per Unit Year Ago"],
            "Precio_calc_ya": precio_calc_ya,
            "Error_relativo": error_rel_ya,
        }
    )[mask_error_ya].head(300)

# --- 3. Consistencia de ACV Weighted Distribution vs ACV Weighted Distribution Year Ago (rango razonable 0–100)

for col in ["ACV Weighted Distribution", "ACV Weighted Distribution Year Ago"]:
    if col not in df.columns:
        print(f"Aviso: columna '{col}' no encontrada. Se omiten reglas de ACV.")
        break
else:
    acv = a_numerico(df["ACV Weighted Distribution"])
    acv_ya = a_numerico(df["ACV Weighted Distribution Year Ago"])

    mask_acv_fuera_rango = (acv < 0) | (acv > 100)
    mask_acv_ya_fuera_rango = (acv_ya < 0) | (acv_ya > 100)

    n_acv_fuera = (mask_acv_fuera_rango | mask_acv_ya_fuera_rango).sum()

    reglas_resumen.append(
        crear_registro_regla(
            "ACV_rango",
            "ACV Weighted Distribution y Year Ago deben estar entre 0 y 100",
            n_acv_fuera,
            total_filas,
        )
    )

    detalles_errores["ACV_rango"] = pd.DataFrame(
        {
            "ACV Weighted Distribution": df["ACV Weighted Distribution"],
            "ACV Weighted Distribution Year Ago": df["ACV Weighted Distribution Year Ago"],
        }
    )[mask_acv_fuera_rango | mask_acv_ya_fuera_rango].head(300)

# --- 4. Consistencia de Year vs Time (cuando Time tiene la forma 'Week Ending mm-dd-yy')

if ("Year" in df.columns) and ("Time" in df.columns):
    year_col = a_numerico(df["Year"])
    time_col = df["Time"].astype(str)

    # intentamos extraer los últimos 2 dígitos del año de la cadena
    # asumiendo formato tipo 'Week Ending 01-05-25'
    year_from_time = []
    for val in time_col:
        # tomamos lo último que parezca año (2 dígitos)
        # fallback: np.nan
        try:
            parts = val.strip().split()
            # buscamos algo que parezca fecha con guiones
            candidate = None
            for p in parts[::-1]:
                if "-" in p:
                    candidate = p
                    break
            if candidate is not None:
                # último segmento después del último '-'
                yy = candidate.split("-")[-1]
                yy_int = int(yy)
                if yy_int < 50:  # asumimos 2000+
                    yy_full = 2000 + yy_int
                else:
                    yy_full = 1900 + yy_int
                year_from_time.append(yy_full)
            else:
                year_from_time.append(np.nan)
        except Exception:
            year_from_time.append(np.nan)

    year_from_time = pd.Series(year_from_time, index=df.index, name="Year_from_Time")

    mask_both = (~year_col.isna()) & (~year_from_time.isna())
    mask_diff = mask_both & (year_col != year_from_time)

    n_diff = mask_diff.sum()

    reglas_resumen.append(
        crear_registro_regla(
            "Year_vs_Time",
            "Year debe coincidir con el año de la fecha en Time (cuando se puede interpretar)",
            n_diff,
            total_filas,
        )
    )

    detalles_errores["Year_vs_Time"] = pd.DataFrame(
        {
            "Time": df["Time"],
            "Year": df["Year"],
            "Year_from_Time": year_from_time,
        }
    )[mask_diff].head(300)

# ----------------------------------------------------------
# RESUMEN FINAL DE REGLAS
# ----------------------------------------------------------

df_reglas_resumen = pd.DataFrame(reglas_resumen) if reglas_resumen else pd.DataFrame(
    columns=["Regla", "Descripción", "Filas con error", "% filas con error"]
)

# ----------------------------------------------------------
# EXPORTAR A EXCEL
# ----------------------------------------------------------

print(f"Generando archivo de salida: {OUTPUT_FILE}")

with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
    df_resumen_columnas.to_excel(writer, sheet_name="Resumen_columnas", index=False)
    df_resumen_dup.to_excel(writer, sheet_name="Duplicados_resumen", index=False)
    df_detalle_dup.to_excel(writer, sheet_name="Duplicados_detalle", index=False)

    df_reglas_resumen.to_excel(writer, sheet_name="Reglas_resumen", index=False)

    # hojas con detalle de errores de reglas
    for regla, df_err in detalles_errores.items():
        # nombre de hoja máx 31 caracteres
        sheet_name = f"Err_{regla}"[:31]
        df_err.to_excel(writer, sheet_name=sheet_name, index=False)

print("Proceso terminado. Revisa el archivo de reporte generado.")
