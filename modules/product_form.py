
import math

import streamlit as st

from modules.excel_manager import COLUMNAS_TUU


# ==========================================
# VALORES PERMITIDOS POR TUU
# ==========================================

OPCIONES_SI_NO = ["S", "N"]


# ==========================================
# FUNCIONES DE APOYO
# ==========================================

def obtener_valor(producto, columna):
    """
    Obtiene un valor existente para edición.
    En modo creación devuelve texto vacío.
    """

    if producto is None:
        return ""

    valor = producto.get(columna, "")

    if valor is None:
        return ""

    return str(valor)


def normalizar_si_no(valor, predeterminado):
    """
    Convierte valores antiguos a los
    códigos oficiales de TUU.

    SI -> S
    NO -> N

    Si el valor es desconocido, no lo
    selecciona automáticamente.
    """

    valor = str(valor).strip().upper()

    if valor in ("S", "SI", "SÍ"):
        return "S"

    if valor in ("N", "NO"):
        return "N"

    if not valor:
        return predeterminado

    return ""


def validar_producto(datos):
    """
    Verifica los campos obligatorios,
    precios y valores oficiales de TUU.
    """

    errores = []

    nombre = datos["NOMBRE*"].strip()
    precio = datos["PRECIO*"].strip()
    exento = datos["¿EXENTO?*"].strip()
    habilitado = datos["¿HABILITADO?*"].strip()

    if not nombre:
        errores.append(
            "El nombre del producto es obligatorio."
        )

    if not precio:
        errores.append(
            "El precio de venta es obligatorio."
        )
    else:
        try:
            valor_precio = float(precio)

            if (
                not math.isfinite(valor_precio)
                or valor_precio < 0
            ):
                errores.append(
                    "El precio de venta debe ser "
                    "un número válido no negativo."
                )

        except ValueError:
            errores.append(
                "El precio debe ser numérico."
            )

    costo = datos["PRECIO COSTO"].strip()

    if costo:
        try:
            valor_costo = float(costo)

            if (
                not math.isfinite(valor_costo)
                or valor_costo < 0
            ):
                errores.append(
                    "El precio de costo debe ser "
                    "un número válido no negativo."
                )

        except ValueError:
            errores.append(
                "El precio de costo debe ser numérico."
            )

    if exento not in OPCIONES_SI_NO:
        errores.append(
            "El campo ¿Exento? debe ser S o N."
        )

    if habilitado not in OPCIONES_SI_NO:
        errores.append(
            "El campo ¿Habilitado? debe ser S o N."
        )

    return errores


# ==========================================
# FORMULARIO COMPARTIDO
# ==========================================

def mostrar_formulario_producto(
    producto=None,
    clave="producto"
):
    """
    Muestra el formulario para crear o editar.

    Retorna:
        (datos, True) si se guarda correctamente.
        (None, False) en otro caso.
    """

    modo_edicion = producto is not None
    borrador = st.session_state.get("sku_borrador")
    if borrador is not None and borrador.get("clave") != clave:
        borrador = None

    def valor(campo):
        if borrador is not None:
            return borrador["datos"].get(campo, "")
        return obtener_valor(producto, campo)


    if modo_edicion:
        st.subheader("✏️ Editar producto")
    else:
        st.subheader("➕ Nuevo producto")

    st.caption(
        "Los campos marcados con * "
        "son obligatorios."
    )

    with st.form(
        key=f"formulario_{clave}",
        clear_on_submit=False,
        enter_to_submit=False
    ):

        # ==================================
        # INFORMACION PRINCIPAL
        # ==================================

        st.markdown("#### Información del producto")

        nombre = st.text_input(
            "Nombre del producto *",
            value=valor(
                "NOMBRE*"
            )
        )

        descripcion = st.text_area(
            "Descripción",
            value=valor(
                "Descripción"
            ),
            height=90
        )

        peso = st.text_input(
            "Peso",
            value=valor(
                "PESO"
            )
        )

        # ==================================
        # PRECIOS
        # ==================================

        st.markdown("#### Precios")

        col_precio, col_costo = st.columns(2)

        with col_precio:
            precio = st.text_input(
                "Precio de venta *",
                value=valor(
                    "PRECIO*"
                ),
                placeholder="Ej. 1990"
            )

        with col_costo:
            precio_costo = st.text_input(
                "Precio de costo",
                value=valor(
                    "PRECIO COSTO"
                ),
                placeholder="Ej. 1200"
            )

        # ==================================
        # IDENTIFICADORES
        # ==================================

        st.markdown("#### Identificación")

        sku = st.text_input(
            "Código SKU",
            value=valor(
                "CÓDIGO SKU"
            )
        )

        codigo_barras = st.text_input(
            "Código de barras",
            value=valor(
                "CÓDIGO DE BARRAS"
            ),
            placeholder="Escribe o escanea el código",
            help=(
                "Compatible con lectores de códigos "
                "de barras que funcionan como teclado."
            )
        )

        generar_barras = st.form_submit_button(
            "🏷️ Generar código de barras interno",
            help="Solo para productos sin código comercial. No reemplaza códigos existentes."
        )

        mensaje_barras = st.session_state.pop('barcode_mensaje', None)
        if mensaje_barras:
            tipo, texto = mensaje_barras
            getattr(st, tipo)(texto)

        st.markdown("##### Generador automático de SKU")
        abreviatura_sku = st.text_input(
            "Abreviatura del producto (opcional, máximo 6 caracteres)",
            value=(borrador or {}).get("abreviatura", ""),
            help="Ejemplo: ZER15. Si queda vacía se calcula desde el nombre."
        )
        generar = st.form_submit_button("🔢 Generar y previsualizar SKU")
        mensaje_sku = st.session_state.pop('sku_mensaje', None)
        if mensaje_sku:
            tipo, texto = mensaje_sku
            getattr(st, tipo)(texto)

        st.caption(
            "Puedes escribir el código o escanearlo "
            "con un lector USB/Bluetooth configurado "
            "como teclado."
        )

        # ==================================
        # CLASIFICACION
        # ==================================

        st.markdown("#### Clasificación")

        marca = st.text_input(
            "Marca",
            value=valor(
                "MARCAS"
            )
        )

        categoria = st.text_input(
            "Categoría",
            value=valor(
                "CATEGORIAS"
            )
        )

        # ==================================
        # CONFIGURACION TUU
        # ==================================

        st.markdown("#### Configuración")

        valor_exento = normalizar_si_no(
            valor(
                "¿EXENTO?*"
            ),
            predeterminado=(
                "" if modo_edicion else "N"
            )
        )

        valor_habilitado = normalizar_si_no(
            valor(
                "¿HABILITADO?*"
            ),
            predeterminado=(
                "" if modo_edicion else "S"
            )
        )

        # Si existe un dato antiguo inválido,
        # se obliga a elegir S o N.
        opciones_exento = (
            [""] + OPCIONES_SI_NO
            if not valor_exento
            else OPCIONES_SI_NO
        )

        opciones_habilitado = (
            [""] + OPCIONES_SI_NO
            if not valor_habilitado
            else OPCIONES_SI_NO
        )

        col_exento, col_habilitado = st.columns(2)

        with col_exento:
            exento = st.selectbox(
                "¿Exento? *",
                options=opciones_exento,
                index=opciones_exento.index(
                    valor_exento
                ),
                format_func=lambda valor: (
                    "Seleccionar..."
                    if valor == ""
                    else valor
                ),
                help=(
                    "S = Sí, está exento. "
                    "N = No está exento."
                )
            )

        with col_habilitado:
            habilitado = st.selectbox(
                "¿Habilitado? *",
                options=opciones_habilitado,
                index=opciones_habilitado.index(
                    valor_habilitado
                ),
                format_func=lambda valor: (
                    "Seleccionar..."
                    if valor == ""
                    else valor
                ),
                help=(
                    "S = Sí, está habilitado. "
                    "N = No está habilitado."
                )
            )

        st.caption(
            "Estos campos se guardan únicamente "
            "como S o N, según el formato de TUU."
        )

        st.divider()

        texto_boton = (
            "💾 Guardar cambios"
            if modo_edicion
            else "➕ Crear producto"
        )

        guardar = st.form_submit_button(
            texto_boton,
            type="primary",
            use_container_width=True
        )

    # ======================================
    # RECOPILAR DATOS
    # ======================================

    datos = {
        "NOMBRE*": nombre.strip(),
        "Descripción": descripcion.strip(),
        "PESO": peso.strip(),
        "PRECIO*": precio.strip(),
        "PRECIO COSTO": precio_costo.strip(),
        "¿EXENTO?*": exento,
        "CÓDIGO SKU": sku.strip(),
        "CÓDIGO DE BARRAS": codigo_barras.strip(),
        "MARCAS": marca.strip(),
        "CATEGORIAS": categoria.strip(),
        "¿HABILITADO?*": habilitado
    }

    # Mantener las 11 columnas oficiales TUU.
    datos = {
        columna: datos[columna]
        for columna in COLUMNAS_TUU
    }

    if generar_barras:
        st.session_state["barcode_solicitud"] = {
            "clave": clave, "datos": datos.copy(),
            "abreviatura": abreviatura_sku.strip()
        }
        return None, False

    if generar:
        st.session_state["sku_solicitud"] = {
            "clave": clave, "datos": datos.copy(),
            "abreviatura": abreviatura_sku.strip()
        }
        return None, False

    if guardar:
        previo_barras = st.session_state.get('barcode_previsualizado')
        st.session_state['barcode_generado_guardar'] = bool(
            previo_barras and previo_barras.get('clave') == clave and
            datos['CÓDIGO DE BARRAS'] == previo_barras.get('codigo')
        )
        # Si se cambió manualmente el SKU de la vista previa, no se reserva.
        previo = st.session_state.get("sku_previsualizado")
        if previo and previo.get("clave") == clave:
            st.session_state["sku_generado_guardar"] = (
                datos["CÓDIGO SKU"].upper() == previo["sku"]
            )
        else:
            st.session_state["sku_generado_guardar"] = False
        errores = validar_producto(datos)

        if errores:
            for error in errores:
                st.error(error)

            return None, False

        return datos, True

    return None, False
