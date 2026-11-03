
import re
import unicodedata

import pandas as pd


def normalizar_texto(texto):
    """Normaliza texto sin eliminar diferencias de producto."""

    texto = str(texto).lower().strip()

    texto = "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )

    texto = re.sub(r"(?<=\d),(?=\d)", ".", texto)

    texto = re.sub(
        r"\b(litros?|lts?)\b", "l", texto
    )

    texto = re.sub(
        r"\b(kilogramos?|kilos?)\b", "kg", texto
    )

    texto = re.sub(
        r"\bgramos?\b", "g", texto
    )

    texto = re.sub(r"(\d)([a-z])", r"\1 \2", texto)
    texto = re.sub(r"[^\w\s.]", " ", texto)

    return " ".join(texto.split())


def normalizar_codigo(codigo):
    return str(codigo).strip().upper()


def detectar_grupos_duplicados(df, umbral=90):
    """
    Detecta:
    - SKU exactamente repetidos.
    - EAN exactamente repetidos.
    - Nombres equivalentes tras normalización.

    No agrupa por parecido aproximado para evitar
    confundir tamaños, sabores y variedades.

    'umbral' se conserva por compatibilidad.
    """

    productos = df.reset_index(drop=True).fillna("")
    cantidad = len(productos)

    padres = list(range(cantidad))
    motivos_por_par = {}

    def buscar(i):
        while padres[i] != i:
            padres[i] = padres[padres[i]]
            i = padres[i]
        return i

    def unir(i, j):
        a = buscar(i)
        b = buscar(j)

        if a != b:
            padres[b] = a

    indices_sku = {}
    indices_ean = {}
    indices_nombre = {}

    def registrar(diccionario, clave, indice, motivo):
        if not clave:
            return

        if clave in diccionario:
            anterior = diccionario[clave]
            par = tuple(sorted((anterior, indice)))

            motivos_por_par.setdefault(par, set()).add(
                motivo
            )

            unir(anterior, indice)
        else:
            diccionario[clave] = indice

    for i, producto in productos.iterrows():
        registrar(
            indices_sku,
            normalizar_codigo(
                producto.get("CÓDIGO SKU", "")
            ),
            i,
            "SKU repetido"
        )

        registrar(
            indices_ean,
            normalizar_codigo(
                producto.get("CÓDIGO DE BARRAS", "")
            ),
            i,
            "Código de barras repetido"
        )

        registrar(
            indices_nombre,
            normalizar_texto(
                producto.get("NOMBRE*", "")
            ),
            i,
            "Nombre equivalente"
        )

    grupos_indices = {}

    for i in range(cantidad):
        raiz = buscar(i)
        grupos_indices.setdefault(raiz, []).append(i)

    resultados = []

    for indices in grupos_indices.values():
        if len(indices) < 2:
            continue

        productos_grupo = []
        motivos_grupo = set()

        for (a, b), motivos in motivos_por_par.items():
            if buscar(a) == buscar(indices[0]):
                motivos_grupo.update(motivos)

        for indice in indices:
            producto = productos.iloc[indice]

            productos_grupo.append({
                "indice": indice,
                "fila_excel": indice + 2,
                "nombre": producto.get("NOMBRE*", ""),
                "sku": producto.get("CÓDIGO SKU", ""),
                "ean": producto.get(
                    "CÓDIGO DE BARRAS", ""
                ),
                "precio": producto.get("PRECIO*", "")
            })

        resultados.append({
            "grupo": 0,
            "cantidad": len(productos_grupo),
            "motivos": sorted(motivos_grupo),
            "productos": productos_grupo
        })

    resultados.sort(
        key=lambda grupo: (
            -grupo["cantidad"],
            grupo["productos"][0]["indice"]
        )
    )

    for numero, grupo in enumerate(resultados, 1):
        grupo["grupo"] = numero

    return resultados


def detectar_duplicados(df, umbral=90):
    """Formato de tabla compatible con app.py."""

    grupos = detectar_grupos_duplicados(df, umbral)
    filas = []

    for grupo in grupos:
        for producto in grupo["productos"]:
            filas.append({
                "Grupo": grupo["grupo"],
                "Cantidad en grupo": grupo["cantidad"],
                "Motivo": ", ".join(grupo["motivos"]),
                "Fila Excel": producto["fila_excel"],
                "Producto": producto["nombre"],
                "SKU": producto["sku"],
                "EAN": producto["ean"],
                "Precio": producto["precio"]
            })

    return pd.DataFrame(
        filas,
        columns=[
            "Grupo",
            "Cantidad en grupo",
            "Motivo",
            "Fila Excel",
            "Producto",
            "SKU",
            "EAN",
            "Precio"
        ]
    )
