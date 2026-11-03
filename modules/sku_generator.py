
import re
import unicodedata

import pandas as pd


# ==========================================
# CONFIGURACION DEL GENERADOR SKU
# ==========================================

MAX_CATEGORIAS = 999
MAX_CORRELATIVO = 99999

LONGITUD_MARCA = 3
LONGITUD_PRODUCTO = 6

PATRON_SKU = re.compile(
    r"^(\d{3})-([A-Z0-9]{1,3})-"
    r"([A-Z0-9]{1,6})-(\d{5})$"
)


# ==========================================
# NORMALIZACION DE TEXTO
# ==========================================

def normalizar_texto_sku(texto):
    """
    Convierte textos a mayusculas, elimina
    acentos y caracteres especiales.
    """

    texto = str(texto or "").strip().upper()

    texto = "".join(
        caracter
        for caracter in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(caracter)
    )

    texto = re.sub(r"[^A-Z0-9\s]", " ", texto)

    return " ".join(texto.split())


# ==========================================
# ABREVIATURA DE MARCA
# ==========================================

def abreviar_marca(marca):
    """
    Genera una abreviatura de hasta 3 caracteres.

    Si no existe marca, utiliza GEN.
    """

    texto = normalizar_texto_sku(marca)

    if not texto:
        return "GEN"

    palabras = texto.split()

    if len(palabras) == 1:
        return palabras[0][:LONGITUD_MARCA]

    iniciales = "".join(
        palabra[0] for palabra in palabras
    )

    if len(iniciales) >= LONGITUD_MARCA:
        return iniciales[:LONGITUD_MARCA]

    return (
        iniciales
        + palabras[0][1:]
    )[:LONGITUD_MARCA]


# ==========================================
# ABREVIATURA DEL PRODUCTO
# ==========================================

def abreviar_producto(nombre, abreviatura=None):
    """
    Genera una descripcion de hasta 6 caracteres.

    Permite proporcionar una abreviatura manual
    para diferenciar variedades y presentaciones.
    """

    if abreviatura:
        texto = normalizar_texto_sku(abreviatura)
        codigo = "".join(texto.split())

        if not codigo:
            raise ValueError(
                "La abreviatura del producto no es valida."
            )

        return codigo[:LONGITUD_PRODUCTO]

    texto = normalizar_texto_sku(nombre)

    palabras = texto.split()

    if not palabras:
        raise ValueError(
            "Debes indicar el nombre del producto."
        )

    return "".join(palabras)[:LONGITUD_PRODUCTO]


# ==========================================
# OBTENER CATEGORIAS EXISTENTES
# ==========================================

def obtener_categorias(df, registro=None):
    """
    Obtiene la correspondencia entre categorias
    y numeros asignados.

    El registro externo permite conservar los
    identificadores entre sesiones.
    """

    categorias = dict(registro or {})

    if df is None or df.empty:
        return categorias

    if "CATEGORIAS" not in df.columns:
        return categorias

    for categoria in df["CATEGORIAS"]:
        nombre = normalizar_texto_sku(categoria)

        if not nombre or nombre in categorias:
            continue

        usados = set(categorias.values())

        for numero in range(1, MAX_CATEGORIAS + 1):
            if numero not in usados:
                categorias[nombre] = numero
                break
        else:
            raise ValueError(
                "Se alcanzo el limite de categorias."
            )

    return categorias


# ==========================================
# ASIGNAR NUMERO DE CATEGORIA
# ==========================================

def obtener_numero_categoria(
    categoria,
    categorias
):
    """
    Obtiene o asigna un identificador numerico
    unico a la categoria.
    """

    nombre = normalizar_texto_sku(categoria)

    if not nombre:
        raise ValueError(
            "Debes seleccionar una categoria."
        )

    if nombre in categorias:
        return categorias[nombre]

    usados = set(categorias.values())

    for numero in range(1, MAX_CATEGORIAS + 1):
        if numero not in usados:
            categorias[nombre] = numero
            return numero

    raise ValueError(
        "No quedan identificadores de categoria."
    )


# ==========================================
# OBTENER SKU EXISTENTES
# ==========================================

def obtener_sku_existentes(df):
    """
    Devuelve un conjunto de SKU existentes,
    normalizados en mayusculas.
    """

    if df is None or df.empty:
        return set()

    if "CÓDIGO SKU" not in df.columns:
        return set()

    return {
        str(valor).strip().upper()
        for valor in df["CÓDIGO SKU"]
        if pd.notna(valor) and str(valor).strip()
    }


# ==========================================
# SIGUIENTE CORRELATIVO DISPONIBLE
# ==========================================

def siguiente_correlativo(
    numero_categoria,
    sku_existentes,
    correlativos_reservados=None
):
    """
    Busca el siguiente correlativo para la
    categoria, respetando los SKU existentes.

    Los correlativos reservados permiten evitar
    reutilizar identificadores ya registrados.
    """

    mayor = 0

    prefijo = f"{numero_categoria:03d}-"

    for sku in sku_existentes:
        sku = str(sku).strip().upper()

        if not sku.startswith(prefijo):
            continue

        coincidencia = PATRON_SKU.fullmatch(sku)

        if coincidencia:
            correlativo = int(coincidencia.group(4))
            mayor = max(mayor, correlativo)

    if correlativos_reservados:
        mayor = max(
            mayor,
            max(correlativos_reservados, default=0)
        )

    siguiente = mayor + 1

    if siguiente > MAX_CORRELATIVO:
        raise ValueError(
            "Se alcanzo el limite de productos "
            "para esta categoria."
        )

    return siguiente


# ==========================================
# GENERAR SKU
# ==========================================

def generar_sku(
    nombre,
    marca,
    categoria,
    df,
    categorias=None,
    abreviatura_producto=None,
    correlativos_reservados=None
):
    """
    Genera un SKU con el formato:

    001-COC-ZER15-00001

    Devuelve el SKU y el registro actualizado
    de categorias.

    No modifica el DataFrame original.
    """

    registro = obtener_categorias(
        df,
        registro=categorias
    )

    numero_categoria = obtener_numero_categoria(
        categoria,
        registro
    )

    codigo_marca = abreviar_marca(marca)

    codigo_producto = abreviar_producto(
        nombre,
        abreviatura=abreviatura_producto
    )

    existentes = obtener_sku_existentes(df)

    correlativo = siguiente_correlativo(
        numero_categoria,
        existentes,
        correlativos_reservados
    )

    sku = (
        f"{numero_categoria:03d}-"
        f"{codigo_marca}-"
        f"{codigo_producto}-"
        f"{correlativo:05d}"
    )

    if sku in existentes:
        raise ValueError(
            "El SKU generado ya existe."
        )

    return sku, registro
