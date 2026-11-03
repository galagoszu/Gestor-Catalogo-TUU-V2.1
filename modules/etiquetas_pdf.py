
"""Generación de etiquetas y cartillas de códigos EAN-13 para Gestor TUU."""

from io import BytesIO
import math
import re
import unicodedata

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.graphics.barcode.eanbc import Ean13BarcodeWidget
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF


def _texto(valor):
    return "" if valor is None else str(valor).strip()


def _orden(valor):
    return unicodedata.normalize("NFKD", _texto(valor)).casefold()


def _precio(valor):
    try:
        n = float(
            _texto(valor)
            .replace("$", "")
            .replace(".", "")
            .replace(",", ".")
        )
        if not math.isfinite(n) or n < 0:
            return "Precio no disponible"
        return "$" + f"{n:,.0f}".replace(",", ".")
    except (ValueError, TypeError):
        return "Precio no disponible"


def _ean_valido(codigo):
    if not re.fullmatch(r"\d{13}", codigo):
        return False

    digitos = [int(x) for x in codigo]
    suma = sum(digitos[:12:2]) + 3 * sum(digitos[1:12:2])

    return (10 - suma % 10) % 10 == digitos[12]


def validar_etiquetas(df):
    """Devuelve (filas imprimibles, problemas) sin modificar el catálogo."""

    validas, problemas = [], []

    for indice, fila in df.iterrows():
        nombre = _texto(fila.get("NOMBRE*"))
        codigo = _texto(fila.get("CÓDIGO DE BARRAS"))

        faltas = []

        if not nombre:
            faltas.append("nombre vacío")

        if not codigo:
            faltas.append("sin código de barras")
        elif not _ean_valido(codigo):
            faltas.append("código no es EAN-13 válido")

        if faltas:
            problemas.append({
                "Fila": int(indice) + 2,
                "Producto": nombre or "(Sin nombre)",
                "Problema": ", ".join(faltas)
            })
        else:
            validas.append({
                "nombre": nombre,
                "sku": _texto(fila.get("CÓDIGO SKU")),
                "codigo": codigo,
                "categoria": _texto(fila.get("CATEGORIAS")) or "Sin categoría",
                "precio": _precio(fila.get("PRECIO*"))
            })

    validas.sort(
        key=lambda x: (
            _orden(x["categoria"]),
            _orden(x["nombre"]),
            x["codigo"]
        )
    )

    return validas, problemas


def _ajustar(c, texto, x, y, ancho, fuente="Helvetica", tam=9, minimo=6.5):
    texto = _texto(texto)
    tam_actual = tam

    while tam_actual > minimo and c.stringWidth(texto, fuente, tam_actual) > ancho:
        tam_actual -= 0.5

    if c.stringWidth(texto, fuente, tam_actual) > ancho:
        while texto and c.stringWidth(texto + "…", fuente, tam_actual) > ancho:
            texto = texto[:-1]
        texto += "…"

    c.setFont(fuente, tam_actual)
    c.drawString(x, y, texto)


def generar_pdf_etiquetas(
    df,
    titulo="Etiquetas de productos",
    cantidades=None,
    modo="cartilla"
):
    """Carta vertical, 3 columnas, etiquetas con borde, agrupadas por categoría."""

    filas, problemas = validar_etiquetas(df)

    if not filas:
        raise ValueError(
            "No hay productos con nombre y código EAN-13 válido para imprimir."
        )

    if modo not in ("cartilla", "etiquetas"):
        raise ValueError("Modo de impresión no válido.")

    if modo == "etiquetas":
        cantidades = cantidades or {}
        filas_con_cantidades = []

        for indice, fila in df.iterrows():
            cantidad = int(cantidades.get(int(indice), 0))

            if cantidad <= 0:
                continue

            validas_fila, _ = validar_etiquetas(df.loc[[indice]])
            filas_con_cantidades.extend(validas_fila * cantidad)

        filas = sorted(
            filas_con_cantidades,
            key=lambda x: (
                _orden(x["categoria"]),
                _orden(x["nombre"]),
                x["codigo"]
            )
        )

        if not filas:
            raise ValueError("No hay etiquetas válidas seleccionadas.")

    salida = BytesIO()
    c = canvas.Canvas(salida, pagesize=letter)
    c.setTitle(titulo)

    ancho, alto = letter

    margen = 27
    separacion = 7
    ancho_et = (ancho - 2 * margen - 2 * separacion) / 3
    alto_et = 111
    cabecera = 34
    pie = 28

    y = alto - 54
    columna = 0
    pagina = 0

    def nueva_pagina():
        nonlocal y, columna, pagina

        if pagina:
            c.showPage()

        pagina += 1

        c.setFont("Helvetica-Bold", 12)
        c.drawString(margen, alto - 29, titulo[:75])

        c.setStrokeColorRGB(.78, .78, .78)
        c.line(margen, alto - 37, ancho - margen, alto - 37)

        c.setFont("Helvetica", 8)
        c.drawRightString(ancho - margen, 17, f"Página {pagina}")

        y = alto - 52
        columna = 0

    nueva_pagina()
    categoria_actual = None

    for fila in filas:
        if fila["categoria"] != categoria_actual:
            if columna:
                y -= alto_et + separacion
                columna = 0

            if y - cabecera - alto_et < pie:
                nueva_pagina()

            categoria_actual = fila["categoria"]

            c.setFillColorRGB(.92, .94, .97)
            c.roundRect(
                margen,
                y - 23,
                ancho - 2 * margen,
                22,
                4,
                fill=1,
                stroke=0
            )

            c.setFillColorRGB(.12, .17, .24)

            _ajustar(
                c,
                categoria_actual,
                margen + 9,
                y - 16,
                ancho - 2 * margen - 18,
                "Helvetica-Bold",
                10
            )

            y -= cabecera

        if y - alto_et < pie:
            nueva_pagina()

            c.setFont("Helvetica-Bold", 9)

            _ajustar(
                c,
                categoria_actual + " (continuación)",
                margen,
                y - 12,
                ancho - 2 * margen,
                "Helvetica-Bold",
                9
            )

            y -= 23

        x = margen + columna * (ancho_et + separacion)
        fondo_y = y - alto_et

        c.setStrokeColorRGB(.65, .68, .73)
        c.setLineWidth(.65)

        c.roundRect(
            x,
            fondo_y,
            ancho_et,
            alto_et,
            4,
            stroke=1,
            fill=0
        )

        _ajustar(
            c,
            fila["categoria"].upper(),
            x + 7,
            y - 12,
            ancho_et - 14,
            "Helvetica-Bold",
            7
        )

        _ajustar(
            c,
            fila["nombre"],
            x + 7,
            y - 25,
            ancho_et - 14,
            "Helvetica-Bold",
            9
        )

        _ajustar(
            c,
            "SKU: " + (fila["sku"] or "Sin SKU"),
            x + 7,
            y - 37,
            ancho_et - 14,
            "Helvetica",
            7.4
        )

        # Codifica los primeros 12 dígitos y calcula el control EAN-13.
        widget = Ean13BarcodeWidget(
            fila["codigo"][:12],
            barHeight=34,
            barWidth=.9,
            humanReadable=False,
            quiet=False
        )

        dibujo = Drawing(0, 0)
        dibujo.add(widget)

        # Centrar sin alterar el ancho de las barras.
        ancho_barras = float(
            widget.getBounds()[2] - widget.getBounds()[0]
        )

        renderPDF.draw(
            dibujo,
            c,
            x + (ancho_et - ancho_barras) / 2,
            y - 76
        )

        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", 8)

        c.drawCentredString(
            x + ancho_et / 2,
            y - 87,
            fila["codigo"]
        )

        c.setFont("Helvetica-Bold", 10)

        c.drawCentredString(
            x + ancho_et / 2,
            y - 103,
            fila["precio"]
        )

        columna += 1

        if columna == 3:
            columna = 0
            y -= alto_et + separacion

    c.save()

    return salida.getvalue(), problemas

