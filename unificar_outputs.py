#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Unifica múltiples archivos (con hojas 'Datos' y 'Diccionario_Semanas') en un solo Excel,
y además incorpora dimensiones externas:
- Brand Dictionary
- Category Dictionary

Salida (4 hojas):
- 'Source of truth'
- 'Week Dictionary'
- 'Brand Dictionary'
- 'Category Dictionary'

También agrega en 'Source of truth' (si existen las columnas necesarias):
- Brand SM:     mapea 'Brand-Int Fresh Value' -> Brand Dictionary[brand] y retorna [name]
- Category SM:  mapea 'Product' -> Category Dictionary[Product] y retorna [category]
- Subcategory SM: mapea 'Product' -> Category Dictionary[Product] y retorna [subcategory]
"""

from pathlib import Path
from time import perf_counter
import pandas as pd

# ==========================
# CONFIGURACIÓN EDITABLE
# ==========================

# Carpeta donde están los archivos de salida individuales
INPUT_FOLDER = r"Salida"

# Patrón de nombres de archivos a unificar (dentro de INPUT_FOLDER)
FILE_PATTERN = "SM_CIRCANA_*.xlsx"

# Archivo Excel de salida unificado
OUTPUT_FILE = r"Salida\00. SM-SourceOfTruth.xlsx"

# Hojas que existen en los archivos de entrada
DATOS_SHEET_NAME = "Datos"
WEEK_DICT_SHEET_NAME = "Diccionario_Semanas"

# Nombres de hojas en el archivo final
OUT_SOT_SHEET = "Source of Truth"
OUT_WEEK_SHEET = "Week Dictionary"
OUT_BRAND_SHEET = "Brand Dictionary"
OUT_CATEGORY_SHEET = "Category Dictionary"

# Dimensiones (excels aparte)
BRAND_DICT_FILE = r"Dimensiones\Brand_Dictionary.xlsx"
BRAND_DICT_SHEET = 0  # o el nombre de hoja, ej: "brand dictionary"

CATEGORY_DICT_FILE = r"Dimensiones\Category_Dictionary.xlsx"
CATEGORY_DICT_SHEET = 0  # o el nombre de hoja, ej: "category dictionary"

# Mapeos (columna origen -> diccionario)
SOT_BRAND_KEY_COL = "Brand-Int Fresh Value"  # en Source of truth
BRAND_DICT_KEY_COL = "Brand"                 # en Brand Dictionary
BRAND_DICT_VALUE_COL = "Name"                # lo que retorna a Brand SM

SOT_PRODUCT_KEY_COL = "Product"              # en Source of truth
CATEGORY_DICT_KEY_COL = "Product"            # en Category Dictionary
CATEGORY_DICT_CAT_COL = "Category"           # retorna a Category SM
CATEGORY_DICT_SUBCAT_COL = "Subcategory"     # retorna a Subcategory SM

# Si quieres agregar una columna indicando el archivo de origen en Source of truth
ADD_SOURCE_COLUMN = False
SOURCE_COLUMN_NAME = "Source_File"


def safe_read_excel(path, sheet):
    """Lee un Excel con manejo de errores y retorna DataFrame (o vacío)."""
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo leer '{path}' (sheet={sheet}): {e}")
        return pd.DataFrame()


def print_progress(done: int, total: int, prefix: str = "Progreso"):
    """Imprime una barra de progreso simple en consola."""
    if total <= 0:
        return
    pct = int((done / total) * 100)
    bar_len = 30
    filled = int(bar_len * pct / 100)
    bar = "#" * filled + "-" * (bar_len - filled)
    print(f"\r{prefix}: [{bar}] {pct:3d}%", end="", flush=True)


def main():
    start_time = perf_counter()
    input_path = Path(INPUT_FOLDER)
    files = sorted(input_path.glob(FILE_PATTERN))

    if not files:
        print(f"No se encontraron archivos en '{INPUT_FOLDER}' con patrón '{FILE_PATTERN}'.")
        return

    print("Archivos encontrados para unificar:")
    for f in files:
        print(" -", f)

    datos_list = []
    week_list = []
    total_files = len(files)
    processed = 0

    # --------------------------
    # 1) Unificar inputs
    # --------------------------
    for f in files:
        # Datos
        try:
            df_datos = pd.read_excel(f, sheet_name=DATOS_SHEET_NAME)
            if ADD_SOURCE_COLUMN:
                df_datos[SOURCE_COLUMN_NAME] = f.name
            datos_list.append(df_datos)
        except Exception as e:
            print(f"[ADVERTENCIA] No se pudo leer hoja '{DATOS_SHEET_NAME}' de {f}: {e}")

        # Diccionario semanas
        try:
            df_week = pd.read_excel(f, sheet_name=WEEK_DICT_SHEET_NAME)
            week_list.append(df_week)
        except Exception as e:
            print(f"[ADVERTENCIA] No se pudo leer hoja '{WEEK_DICT_SHEET_NAME}' de {f}: {e}")

        processed += 1
        print_progress(processed, total_files, prefix="Unificando archivos")

    if not datos_list and not week_list:
        print("No se pudo leer ninguna hoja de entrada.")
        return

    # Source of truth (antes: Datos)
    if datos_list:
        sot = pd.concat(datos_list, ignore_index=True)
    else:
        sot = pd.DataFrame()
        print("[INFO] No se unificó 'Datos' porque no se encontró ninguna hoja válida.")

    # Week Dictionary (antes: Diccionario_Semanas)
    if week_list:
        week_dict = pd.concat(week_list, ignore_index=True)

        # Renombrar y limpiar columnas para que queden como ('Week', 'Time')
        if {"Week No", "Week Ending"}.issubset(set(week_dict.columns)):
            week_dict = week_dict.rename(columns={"Week No": "Week", "Week Ending": "Time"})

        # Intentar eliminar duplicados por ('Week', 'Time') si existen
        cols_for_dupes = [c for c in ["Week", "Time", "Time "] if c in week_dict.columns]
        if not cols_for_dupes:
            cols_for_dupes = [c for c in ["Week No", "Week Ending"] if c in week_dict.columns]
        if cols_for_dupes:
            week_dict = week_dict.drop_duplicates(subset=cols_for_dupes)
        else:
            week_dict = week_dict.drop_duplicates()

        # Dejar solo las columnas relevantes si existen
        # Orden final: Time primero, luego Week
        desired_cols = [c for c in ["Time", "Time ", "Week"] if c in week_dict.columns]
        if desired_cols:
            week_dict = week_dict[desired_cols]
    else:
        week_dict = pd.DataFrame()
        print("[INFO] No se unificó 'Diccionario_Semanas' porque no se encontró ninguna hoja válida.")

    # --------------------------
    # 2) Leer dimensiones
    # --------------------------
    brand_dict = safe_read_excel(BRAND_DICT_FILE, BRAND_DICT_SHEET)
    category_dict = safe_read_excel(CATEGORY_DICT_FILE, CATEGORY_DICT_SHEET)

    # --------------------------
    # 3) Enriquecer Source of truth con columnas SM
    # --------------------------
    if not sot.empty:
        # --- Brand SM (map) ---
        if (
            not brand_dict.empty
            and SOT_BRAND_KEY_COL in sot.columns
            and BRAND_DICT_KEY_COL in brand_dict.columns
            and BRAND_DICT_VALUE_COL in brand_dict.columns
        ):
            brand_map = (
                brand_dict[[BRAND_DICT_KEY_COL, BRAND_DICT_VALUE_COL]]
                .dropna(subset=[BRAND_DICT_KEY_COL])
                .drop_duplicates(subset=[BRAND_DICT_KEY_COL])
                .set_index(BRAND_DICT_KEY_COL)[BRAND_DICT_VALUE_COL]
            )
            sot["Brand SM"] = sot[SOT_BRAND_KEY_COL].map(brand_map)
        else:
            print("[INFO] No se pudo calcular 'Brand SM' (faltan columnas o Brand Dictionary).")
            if "Brand SM" not in sot.columns:
                sot["Brand SM"] = pd.NA

        # --- Category SM / Subcategory SM (merge por Product) ---
        needed_cat_cols = {CATEGORY_DICT_KEY_COL, CATEGORY_DICT_CAT_COL, CATEGORY_DICT_SUBCAT_COL}
        if (
            not category_dict.empty
            and SOT_PRODUCT_KEY_COL in sot.columns
            and needed_cat_cols.issubset(set(category_dict.columns))
        ):
            cat_small = (
                category_dict[[CATEGORY_DICT_KEY_COL, CATEGORY_DICT_CAT_COL, CATEGORY_DICT_SUBCAT_COL]]
                .dropna(subset=[CATEGORY_DICT_KEY_COL])
                .drop_duplicates(subset=[CATEGORY_DICT_KEY_COL])
            )

            # Evitar colisiones si ya existen
            for col in ["Category SM", "Subcategory SM"]:
                if col in sot.columns:
                    sot = sot.drop(columns=[col])

            sot = sot.merge(
                cat_small,
                how="left",
                left_on=SOT_PRODUCT_KEY_COL,
                right_on=CATEGORY_DICT_KEY_COL
            )

            # Renombrar a las columnas finales
            sot = sot.rename(columns={
                CATEGORY_DICT_CAT_COL: "Category SM",
                CATEGORY_DICT_SUBCAT_COL: "Subcategory SM",
            })

            # Si el key col del diccionario no es exactamente igual al de SOT, lo quitamos
            if CATEGORY_DICT_KEY_COL != SOT_PRODUCT_KEY_COL and CATEGORY_DICT_KEY_COL in sot.columns:
                sot = sot.drop(columns=[CATEGORY_DICT_KEY_COL])
        else:
            print("[INFO] No se pudo calcular 'Category SM/Subcategory SM' (faltan columnas o Category Dictionary).")
            if "Category SM" not in sot.columns:
                sot["Category SM"] = pd.NA
            if "Subcategory SM" not in sot.columns:
                sot["Subcategory SM"] = pd.NA

    # --------------------------
    # 4) Guardar archivo final
    # --------------------------
    output_path = Path(OUTPUT_FILE)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Siempre intentamos crear las 4 hojas (aunque vayan vacías)
        sot.to_excel(writer, sheet_name=OUT_SOT_SHEET, index=False)
        week_dict.to_excel(writer, sheet_name=OUT_WEEK_SHEET, index=False)
        brand_dict.to_excel(writer, sheet_name=OUT_BRAND_SHEET, index=False)
        category_dict.to_excel(writer, sheet_name=OUT_CATEGORY_SHEET, index=False)

    elapsed = perf_counter() - start_time
    print(f"\rArchivo unificado guardado en: {output_path}")
    print(f"Tiempo total de procesamiento: {elapsed:0.2f} segundos")


if __name__ == "__main__":
    main()
