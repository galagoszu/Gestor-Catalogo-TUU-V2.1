
from pathlib import Path
import json
import os


# ==========================================
# CONFIGURACION DEL REGISTRO SKU
# ==========================================

NOMBRE_ARCHIVO = "sku_registry.json"

VERSION_REGISTRO = 1

MAX_CATEGORIAS = 999
MAX_CORRELATIVO = 99999


# ==========================================
# UBICACION DEL REGISTRO
# ==========================================

def obtener_ruta_registro():
    """
    Devuelve la ubicación permanente del
    registro de SKU en Windows.

    El archivo se almacena fuera de la
    carpeta donde está instalado el programa.
    """

    base = os.environ.get("LOCALAPPDATA")

    if base:
        carpeta = Path(base) / "GestorTUU"
    else:
        carpeta = Path.home() / "GestorTUU"

    carpeta.mkdir(
        parents=True,
        exist_ok=True
    )

    return carpeta / NOMBRE_ARCHIVO


# ==========================================
# CREAR REGISTRO INICIAL
# ==========================================

def crear_registro_vacio():
    """
    Crea la estructura inicial del registro.

    categorias:
        Nombre normalizado -> número único.

    correlativos:
        Número de categoría -> lista de
        correlativos utilizados.
    """

    return {
        "version": VERSION_REGISTRO,
        "categorias": {},
        "correlativos": {}
    }


# ==========================================
# VALIDAR REGISTRO
# ==========================================

def validar_registro(registro):
    """
    Comprueba que el registro tenga una
    estructura válida antes de utilizarlo.
    """

    if not isinstance(registro, dict):
        raise ValueError(
            "El registro SKU debe ser un objeto JSON."
        )

    categorias = registro.get("categorias")
    correlativos = registro.get("correlativos")

    if not isinstance(categorias, dict):
        raise ValueError(
            "El registro de categorías no es válido."
        )

    if not isinstance(correlativos, dict):
        raise ValueError(
            "El registro de correlativos no es válido."
        )

    numeros_asignados = set()

    for nombre, numero in categorias.items():

        if not isinstance(nombre, str) or not nombre.strip():
            raise ValueError(
                "Existe una categoría sin nombre válido."
            )

        if type(numero) is not int:
            raise ValueError(
                f"Identificador inválido para {nombre}."
            )

        if not 1 <= numero <= MAX_CATEGORIAS:
            raise ValueError(
                f"Identificador fuera de rango: {nombre}."
            )

        if numero in numeros_asignados:
            raise ValueError(
                "Existen categorías con el mismo identificador."
            )

        numeros_asignados.add(numero)

    for categoria, valores in correlativos.items():

        if not categoria.isdigit():
            raise ValueError(
                "Identificador de categoría inválido "
                "en correlativos."
            )

        numero_categoria = int(categoria)

        if not 1 <= numero_categoria <= MAX_CATEGORIAS:
            raise ValueError(
                "Número de categoría fuera de rango."
            )

        if not isinstance(valores, list):
            raise ValueError(
                "Los correlativos deben ser una lista."
            )

        numeros_vistos = set()

        for valor in valores:

            if type(valor) is not int:
                raise ValueError(
                    "Existe un correlativo no numérico."
                )

            if not 1 <= valor <= MAX_CORRELATIVO:
                raise ValueError(
                    "Existe un correlativo fuera de rango."
                )

            if valor in numeros_vistos:
                raise ValueError(
                    "Existen correlativos repetidos "
                    "en el registro."
                )

            numeros_vistos.add(valor)

    return True


# ==========================================
# CARGAR REGISTRO
# ==========================================

def cargar_registro():
    """
    Recupera el registro guardado.

    Si todavía no existe, devuelve un
    registro vacío.

    Si está dañado, informa el error sin
    sobrescribir los datos existentes.
    """

    ruta = obtener_ruta_registro()

    if not ruta.exists():
        return crear_registro_vacio()

    try:
        with ruta.open(
            "r",
            encoding="utf-8"
        ) as archivo:
            registro = json.load(archivo)

        validar_registro(registro)

        return registro

    except (OSError, ValueError, TypeError) as error:
        raise RuntimeError(
            f"No se pudo cargar el registro SKU: {error}"
        ) from error


# ==========================================
# GUARDAR REGISTRO
# ==========================================

def guardar_registro(registro):
    """
    Guarda el registro de forma segura.

    Primero escribe un archivo temporal y
    después reemplaza el anterior.
    """

    validar_registro(registro)

    ruta = obtener_ruta_registro()

    temporal = ruta.with_name(
        ruta.name + ".tmp"
    )

    try:
        with temporal.open(
            "w",
            encoding="utf-8"
        ) as archivo:
            json.dump(
                registro,
                archivo,
                ensure_ascii=False,
                indent=4
            )

        os.replace(
            temporal,
            ruta
        )

        return True, None

    except OSError as error:
        return False, (
            f"No se pudo guardar el registro SKU: {error}"
        )


# ==========================================
# CONSULTAR CATEGORIAS
# ==========================================

def obtener_categorias_registradas(registro):
    """
    Devuelve una copia de las categorías
    registradas.
    """

    return dict(
        registro["categorias"]
    )


# ==========================================
# CONSULTAR CORRELATIVOS
# ==========================================

def obtener_correlativos_reservados(
    registro,
    numero_categoria
):
    """
    Devuelve los correlativos reservados
    para una categoría específica.
    """

    clave = str(int(numero_categoria))

    return list(
        registro["correlativos"].get(
            clave,
            []
        )
    )


# ==========================================
# REGISTRAR SKU GENERADO
# ==========================================

def registrar_sku(
    registro,
    sku,
    categorias_actualizadas
):
    """
    Incorpora al registro un SKU generado
    y la asignación actual de categorías.

    No guarda automáticamente el archivo.
    La persistencia se realiza mediante
    guardar_registro().
    """

    import re

    patron = re.compile(
        r"^(\d{3})-([A-Z0-9]{1,3})-"
        r"([A-Z0-9]{1,6})-(\d{5})$"
    )

    coincidencia = patron.fullmatch(
        str(sku).strip().upper()
    )

    if not coincidencia:
        raise ValueError(
            "El SKU no cumple el formato oficial "
            "definido para el gestor."
        )

    numero_categoria = int(
        coincidencia.group(1)
    )

    correlativo = int(
        coincidencia.group(4)
    )

    if correlativo < 1:
        raise ValueError(
            "El correlativo debe ser mayor que cero."
        )

    nuevo_registro = {
        "version": VERSION_REGISTRO,
        "categorias": dict(
            categorias_actualizadas
        ),
        "correlativos": {
            clave: list(valores)
            for clave, valores
            in registro["correlativos"].items()
        }
    }

    clave_categoria = str(numero_categoria)

    reservados = nuevo_registro[
        "correlativos"
    ].setdefault(
        clave_categoria,
        []
    )

    if correlativo in reservados:
        raise ValueError(
            "El correlativo ya está registrado."
        )

    reservados.append(correlativo)
    reservados.sort()

    validar_registro(nuevo_registro)

    return nuevo_registro
