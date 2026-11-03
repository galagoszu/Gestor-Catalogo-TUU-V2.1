
from pathlib import Path
from datetime import datetime
import os

import pandas as pd


# ==========================================
# COLUMNAS OFICIALES DE TUU
# ==========================================

COLUMNAS_TUU = [
    "NOMBRE*",
    "Descripción",
    "PESO",
    "PRECIO*",
    "PRECIO COSTO",
    "¿EXENTO?*",
    "CÓDIGO SKU",
    "CÓDIGO DE BARRAS",
    "MARCAS",
    "CATEGORIAS",
    "¿HABILITADO?*"
]


# ==========================================
# CARPETA DE TRABAJO
# ==========================================

def obtener_carpeta_trabajo():
    """
    Carpeta permanente de trabajo para cada
    usuario de Windows.

    Los Excel no dependen de dónde se instale
    el programa.
    """

    base = os.environ.get("LOCALAPPDATA")

    if base:
        carpeta = Path(base) / "GestorTUU" / "trabajo"
    else:
        carpeta = Path.home() / "GestorTUU" / "trabajo"

    carpeta.mkdir(parents=True, exist_ok=True)

    return carpeta


# ==========================================
# CREAR CATALOGO VACIO - NUEVO V2.0
# ==========================================

def crear_catalogo_vacio():
    """
    Crea un catálogo nuevo utilizando exactamente
    las 11 columnas oficiales de TUU.

    Inicialmente no contiene productos.

    La primera fila editable será administrada
    desde la interfaz de Streamlit.
    """

    return pd.DataFrame(columns=COLUMNAS_TUU).fillna("")


# ==========================================
# CARGAR EXCEL EXISTENTE
# ==========================================

def cargar_excel(archivo):
    try:
        df = pd.read_excel(
            archivo,
            dtype=str,
            keep_default_na=False,
            engine="openpyxl"
        )

        if list(df.columns) != COLUMNAS_TUU:
            return None, (
                "El archivo no coincide con las 11 columnas "
                "oficiales de TUU. Revisa los encabezados y su orden."
            )

        return df, None

    except Exception as e:
        return None, f"Error al cargar Excel: {e}"


# ==========================================
# GUARDAR EXCEL
# ==========================================

def guardar_excel(df, ruta):
    try:
        ruta = Path(ruta)

        ruta.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        df = df[COLUMNAS_TUU].copy()

        temporal = ruta.with_name(
            ruta.stem + "_temporal.xlsx"
        )

        df.to_excel(
            temporal,
            index=False,
            engine="openpyxl"
        )

        temporal.replace(ruta)

        return True, None

    except Exception as e:
        return False, f"Error al guardar Excel: {e}"


# ==========================================
# CREAR RUTA DE TRABAJO
# ==========================================

def crear_ruta_trabajo(nombre_archivo):
    carpeta = obtener_carpeta_trabajo()

    fecha = datetime.now().strftime(
        "%Y%m%d_%H%M%S_%f"
    )

    nombre = Path(nombre_archivo).stem

    return carpeta / f"{nombre}_{fecha}.xlsx"


# ==========================================
# LISTAR CATALOGOS DE TRABAJO
# ==========================================

def listar_catalogos_trabajo():
    carpeta = obtener_carpeta_trabajo()

    return sorted(
        (
            ruta
            for ruta in carpeta.glob("*.xlsx")
            if not ruta.stem.endswith("_temporal")
        ),
        key=lambda ruta: ruta.stat().st_mtime,
        reverse=True
    )
