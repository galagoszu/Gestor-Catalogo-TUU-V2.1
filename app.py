
from pathlib import Path
import hashlib
import math

import pandas as pd
import streamlit as st

from modules.excel_manager import (
    COLUMNAS_TUU,
    cargar_excel,
    guardar_excel,
    crear_ruta_trabajo,
    crear_catalogo_vacio
)

from modules.duplicate_detector import (
    detectar_grupos_duplicados,
    normalizar_texto,
    normalizar_codigo
)

from modules.product_form import mostrar_formulario_producto
from modules.etiquetas_pdf import generar_pdf_etiquetas, validar_etiquetas
from modules.barcode_generator import generar_ean13_interno
from modules.barcode_registry import cargar_codigos, guardar_codigos, filtrar_por_origen, origen_codigo
from modules.sku_generator import generar_sku, obtener_categorias, normalizar_texto_sku, PATRON_SKU
from modules.sku_registry import (
    cargar_registro, guardar_registro, registrar_sku,
    obtener_correlativos_reservados,
)


# ==========================================
# CONFIGURACION GENERAL
# ==========================================

st.set_page_config(
    page_title="Gestor de Catálogo TUU",
    page_icon="📦",
    layout="wide"
)

LIMITE_HISTORIAL = 30

# ==========================================
# ESTILO VISUAL V2.1 — sin alterar la lógica
# ==========================================
st.markdown("""
<style>
/* Espaciado general y ancho cómodo para tablas */
.block-container { padding-top: 1.5rem; padding-bottom: 2.5rem; }
/* Jerarquía tipográfica consistente */
h1 { font-size: 2rem !important; letter-spacing: -.025em; line-height: 1.2; }
h2 { font-size: 1.45rem !important; letter-spacing: -.015em; line-height: 1.3; }
h3 { font-size: 1.22rem !important; line-height: 1.35; }
h4 { font-size: 1.1rem !important; line-height: 1.4; margin-top: .7rem !important; }
h5 { font-size: 1rem !important; line-height: 1.4; }
/* Secciones desplegables con encabezados legibles */
[data-testid="stExpander"] summary {
  font-weight: 700;
  font-size: 1.16rem;
  position: relative;
}
/* Centra el título dentro del encabezado sin desplazar la flecha. */
[data-testid="stExpander"] summary > div[data-testid="stMarkdownContainer"] {
  width: 100%;
  text-align: center;
}
[data-testid="stExpander"] summary p {
  font-size: inherit;
  font-weight: inherit;
  text-align: center;
}
/* Botones uniformes, sin imponer colores ajenos al tema */
.stButton button, .stDownloadButton button { border-radius: .55rem; min-height: 2.5rem; font-weight: 550; }
/* Formularios y tablas con separaciones equilibradas */
[data-testid="stVerticalBlock"] > div:has(> [data-testid="stDataFrame"]) { margin-top: .3rem; }
[data-testid="stAlert"] { border-radius: .6rem; }
[data-testid="stMetric"] { padding-block: .2rem; }
@media (max-width: 768px) {
  .block-container { padding-top: 1rem; }
  h1 { font-size: 1.65rem !important; }
  h2 { font-size: 1.3rem !important; }
}
</style>
""", unsafe_allow_html=True)


st.title("📦 Gestor de Catálogo TUU")
st.caption("Administración de productos para minimarket")


# ==========================================
# ESTADO DE SESION
# ==========================================

VALORES_INICIALES = {
    "vista": "principal",
    "archivo_id": None,
    "modo_catalogo": "Importar Excel existente",
    "df_productos": None,
    "ruta_trabajo": None,
    "version_editor": 0,
    "version_formulario": 0,
    "indice_edicion": None,
    "historial": [],
    "posicion_historial": 0,
    "grupos_duplicados": None,
    "producto_eliminar": None,
    "datos_pendientes": None,
    "modo_pendiente": None,
    "indice_pendiente": None,
    "mensaje": None,
    "version_tabla": 0,
    "filtro_edicion": "Todos los productos",
    "busqueda_edicion": "",
    "busqueda_catalogo": "",
    "marca_catalogo": "Todas las marcas",
    "categoria_catalogo": "Todas las categorías",
    "origen_catalogo": "Todos los códigos",
    "origen_edicion": "Todos los códigos",
    "masivo_confirmacion": None,
    "confirmar_reinicio_global": False,
    "mensaje_inicio": None,
}

for clave, valor in VALORES_INICIALES.items():
    if clave not in st.session_state:
        st.session_state[clave] = valor


def reiniciar_sesion_completa():
    """Descarta el catálogo activo y devuelve la app a su pantalla inicial."""
    st.session_state.clear()
    st.session_state["mensaje_inicio"] = (
        "La sesión se reinició. Puedes importar otro Excel o crear un catálogo nuevo."
    )


def cancelar_reinicio_global():
    st.session_state["confirmar_reinicio_global"] = False


def solicitar_reinicio_global():
    st.session_state["confirmar_reinicio_global"] = True


def limpiar_sesion_al_quitar_excel():
    """
    Callback del file_uploader al quitar el Excel con la X.
    No modifica la clave del propio widget: Streamlit la administra y es
    de solo lectura desde session_state. Limpia únicamente el estado de la app.
    """
    if (
        st.session_state.get("archivo_excel_catalogo") is None
        and st.session_state.get("archivo_id") is not None
    ):
        # Restablecer los datos de la aplicación sin tocar el estado interno
        # del file_uploader (archivo_excel_catalogo).
        for clave, valor in VALORES_INICIALES.items():
            if clave != "archivo_excel_catalogo":
                st.session_state[clave] = valor

        st.session_state["mensaje_inicio"] = (
            "Se quitó el archivo Excel y se limpió el catálogo de la sesión."
        )


# ==========================================
# REINICIO GLOBAL DEL CATALOGO
# ==========================================

if st.session_state.get("df_productos") is not None:
    col_reinicio_texto, col_reinicio_boton = st.columns([5, 1.6])
    with col_reinicio_texto:
        st.caption(
            "¿Terminaste de trabajar o deseas comenzar con otro archivo? "
            "Puedes volver a la pantalla inicial."
        )
    with col_reinicio_boton:
        st.button(
            "🔄 Reiniciar programa",
            key="solicitar_reinicio_programa",
            on_click=solicitar_reinicio_global,
            width="stretch",
        )

if st.session_state.get("confirmar_reinicio_global", False):
    st.warning(
        "**¿Deseas reiniciar el programa y volver a la pantalla inicial?**\n\n"
        "Se quitará el Excel activo y se descartarán los cambios de esta sesión "
        "que todavía no hayas descargado. Asegúrate de guardar o descargar tu "
        "catálogo antes de continuar. Los archivos guardados en tu computador "
        "no se eliminarán."
    )
    col_confirmar_reinicio, col_cancelar_reinicio = st.columns(2)
    with col_confirmar_reinicio:
        st.button(
            "⚠️ Sí, reiniciar y descartar la sesión",
            key="confirmar_reinicio_programa",
            type="primary",
            on_click=reiniciar_sesion_completa,
            width="stretch",
        )
    with col_cancelar_reinicio:
        st.button(
            "Cancelar",
            key="cancelar_reinicio_programa",
            on_click=cancelar_reinicio_global,
            width="stretch",
        )


def limpiar_busqueda_catalogo():
    st.session_state["busqueda_catalogo"] = ""
    st.session_state["marca_catalogo"] = "Todas las marcas"
    st.session_state["categoria_catalogo"] = "Todas las categorías"
    st.session_state["origen_catalogo"] = "Todos los códigos"


def limpiar_filtros_edicion():
    st.session_state["filtro_edicion"] = "Todos los productos"
    st.session_state["busqueda_edicion"] = ""
    st.session_state["origen_edicion"] = "Todos los códigos"


def opciones_columna(df, columna, titulo):
    """Lista ordenada de valores reales, ignorando vacíos."""
    valores = sorted({
        str(valor).strip()
        for valor in df[columna].fillna("")
        if str(valor).strip()
    }, key=str.casefold)
    return [titulo] + valores


# ==========================================
# FUNCIONES GENERALES
# ==========================================

def limpiar_pendientes():
    st.session_state["datos_pendientes"] = None
    st.session_state["modo_pendiente"] = None
    st.session_state["indice_pendiente"] = None


def limpiar_duplicados():
    st.session_state["grupos_duplicados"] = None
    st.session_state["producto_eliminar"] = None


def cambiar_vista(vista, indice=None):
    st.session_state["vista"] = vista
    st.session_state["indice_edicion"] = indice
    st.session_state["version_formulario"] += 1
    limpiar_pendientes()
    st.session_state.pop("sku_borrador", None)
    st.session_state.pop("sku_previsualizado", None)
    st.session_state.pop("sku_solicitud", None)
    st.session_state.pop("sku_generado_guardar", None)
    st.session_state.pop('barcode_generado_guardar', None)
    st.session_state.pop('barcode_previsualizado', None)
    st.session_state.pop('barcode_mensaje', None)
    st.session_state.pop('sku_mensaje', None)


def iniciar_historial(df):
    st.session_state["historial"] = [
        df.copy(deep=True)
    ]
    st.session_state["posicion_historial"] = 0


def registrar_historial(df):
    historial = st.session_state["historial"]
    posicion = st.session_state["posicion_historial"]

    historial = historial[:posicion + 1]
    historial.append(df.copy(deep=True))

    if len(historial) > LIMITE_HISTORIAL + 1:
        historial.pop(0)

    st.session_state["historial"] = historial
    st.session_state["posicion_historial"] = (
        len(historial) - 1
    )


def guardar_cambios(df_nuevo):
    df_nuevo = (
        df_nuevo[COLUMNAS_TUU]
        .copy()
        .reset_index(drop=True)
        .fillna("")
        .astype(str)
    )

    correcto, error = guardar_excel(
        df_nuevo,
        st.session_state["ruta_trabajo"]
    )

    if correcto:
        st.session_state["df_productos"] = df_nuevo
        registrar_historial(df_nuevo)

        st.session_state["version_editor"] += 1
        st.session_state["version_tabla"] += 1

        limpiar_duplicados()

    return correcto, error


def mover_historial(direccion):
    historial = st.session_state["historial"]
    posicion = st.session_state["posicion_historial"]
    nueva_posicion = posicion + direccion

    if not 0 <= nueva_posicion < len(historial):
        return False, "No hay más cambios disponibles."

    restaurado = historial[
        nueva_posicion
    ].copy(deep=True)

    correcto, error = guardar_excel(
        restaurado,
        st.session_state["ruta_trabajo"]
    )

    if correcto:
        st.session_state["df_productos"] = restaurado
        st.session_state["posicion_historial"] = (
            nueva_posicion
        )

        st.session_state["version_editor"] += 1
        st.session_state["version_tabla"] += 1

        limpiar_duplicados()
        limpiar_pendientes()

    return correcto, error


# ==========================================
# VALIDACION DEL CATALOGO
# ==========================================

def obtener_problemas(producto):
    """
    Revisa únicamente los campos
    obligatorios para TUU.
    """

    problemas = []

    nombre = str(
        producto["NOMBRE*"]
    ).strip()

    precio = str(
        producto["PRECIO*"]
    ).strip()

    exento = str(
        producto["¿EXENTO?*"]
    ).strip()

    habilitado = str(
        producto["¿HABILITADO?*"]
    ).strip()

    if not nombre:
        problemas.append("Nombre vacío")

    if not precio:
        problemas.append("Precio vacío")
    else:
        try:
            valor = float(precio)

            if not math.isfinite(valor):
                problemas.append(
                    "Precio no válido"
                )
            elif valor < 0:
                problemas.append(
                    "Precio negativo"
                )

        except ValueError:
            problemas.append(
                "Precio no numérico"
            )

    if exento not in ("S", "N"):
        problemas.append(
            "Exento debe ser S o N"
        )

    if habilitado not in ("S", "N"):
        problemas.append(
            "Habilitado debe ser S o N"
        )

    return problemas


def validar_catalogo(df):
    errores = []

    for indice, producto in df.iterrows():
        problemas = obtener_problemas(
            producto
        )

        if problemas:
            nombre = str(
                producto["NOMBRE*"]
            ).strip()

            errores.append({
                "Fila Excel": indice + 2,
                "Producto": (
                    nombre or "(Sin nombre)"
                ),
                "Problemas": ", ".join(problemas)
            })

    return errores


def calcular_avance(df):
    total = len(df)
    errores = validar_catalogo(df)

    pendientes = len(errores)
    completos = total - pendientes

    porcentaje = (
        (completos / total) * 100
        if total > 0 else 0.0
    )

    return {
        "total": total,
        "completos": completos,
        "pendientes": pendientes,
        "porcentaje": porcentaje,
        "errores": errores
    }


# ==========================================
# FILTROS DE EDICION RAPIDA
# ==========================================

OPCIONES_FILTRO = [
    "Todos los productos",
    "Productos completos",
    "Productos incompletos",
    "Sin nombre",
    "Sin precio",
    "Sin SKU",
    "Sin código de barras",
    "Sin descripción",
    "Sin marca",
    "Sin categoría",
    "Sin peso",
    "Sin precio de costo",
    "Exento sin definir",
    "Habilitado sin definir",
]

CAMPOS_FALTANTES = {
    "Sin nombre": "NOMBRE*",
    "Sin precio": "PRECIO*",
    "Sin SKU": "CÓDIGO SKU",
    "Sin código de barras": "CÓDIGO DE BARRAS",
    "Sin descripción": "Descripción",
    "Sin marca": "MARCAS",
    "Sin categoría": "CATEGORIAS",
    "Sin peso": "PESO",
    "Sin precio de costo": "PRECIO COSTO",
}


def filtrar_productos(df, filtro, busqueda):
    """
    Mantiene los índices originales
    de los productos.

    Esto permite guardar correctamente
    las modificaciones de una tabla filtrada.
    """

    resultado = df.copy()

    if filtro == "Productos completos":
        mascara = resultado.apply(
            lambda producto: (
                len(obtener_problemas(producto)) == 0
            ),
            axis=1
        )

        resultado = resultado.loc[mascara]

    elif filtro == "Productos incompletos":
        mascara = resultado.apply(
            lambda producto: (
                len(obtener_problemas(producto)) > 0
            ),
            axis=1
        )

        resultado = resultado.loc[mascara]

    elif filtro in CAMPOS_FALTANTES:
        columna = CAMPOS_FALTANTES[filtro]

        mascara = (
            resultado[columna]
            .fillna("")
            .astype(str)
            .str.strip()
            .eq("")
        )

        if filtro == "Sin precio":
            # Incluye precios vacíos, inválidos
            # o negativos para facilitar revisión.
            mascara = resultado.apply(
                lambda producto: any(
                    problema.startswith("Precio")
                    for problema in obtener_problemas(
                        producto
                    )
                ),
                axis=1
            )

        resultado = resultado.loc[mascara]

    elif filtro == "Exento sin definir":
        mascara = ~resultado[
            "¿EXENTO?*"
        ].astype(str).str.strip().isin(["S", "N"])

        resultado = resultado.loc[mascara]

    elif filtro == "Habilitado sin definir":
        mascara = ~resultado[
            "¿HABILITADO?*"
        ].astype(str).str.strip().isin(["S", "N"])

        resultado = resultado.loc[mascara]

    texto = busqueda.strip()

    if texto:
        mascara_busqueda = (
            resultado["NOMBRE*"]
            .str.contains(
                texto,
                case=False,
                regex=False,
                na=False
            )
            |
            resultado["CÓDIGO SKU"]
            .str.contains(
                texto,
                case=False,
                regex=False,
                na=False
            )
            |
            resultado["CÓDIGO DE BARRAS"]
            .str.contains(
                texto,
                case=False,
                regex=False,
                na=False
            )
        )

        resultado = resultado.loc[
            mascara_busqueda
        ]

    return resultado


def guardar_edicion_filtrada(
    df_original,
    df_mostrado,
    df_editado
):
    """
    Guarda los cambios en las filas
    originales del catálogo.

    No modifica las filas ocultas
    por el filtro.
    """

    df_nuevo = df_original.copy()

    if len(df_mostrado) != len(df_editado):
        return (
            False,
            "No se puede cambiar la cantidad de "
            "filas mientras hay un filtro activo."
        )

    for indice_original in df_mostrado.index:
        for columna in COLUMNAS_TUU:
            valor_anterior = (
                df_mostrado.at[
                    indice_original,
                    columna
                ]
            )

            valor_nuevo = (
                df_editado.at[
                    indice_original,
                    columna
                ]
            )

            if pd.isna(valor_nuevo):
                valor_nuevo = ""

            if str(valor_anterior) != str(valor_nuevo):
                df_nuevo.at[
                    indice_original,
                    columna
                ] = str(valor_nuevo)

    return guardar_cambios(df_nuevo)


# ==========================================
# COMPROBAR DUPLICADOS AL GUARDAR
# ==========================================

def buscar_coincidencias(datos, df, excluir=None):
    coincidencias = []

    nombre = normalizar_texto(datos["NOMBRE*"])
    sku = normalizar_codigo(datos["CÓDIGO SKU"])
    ean = normalizar_codigo(
        datos["CÓDIGO DE BARRAS"]
    )

    for indice, producto in df.iterrows():
        if indice == excluir:
            continue

        motivos = []

        nombre_actual = normalizar_texto(
            producto["NOMBRE*"]
        )

        sku_actual = normalizar_codigo(
            producto["CÓDIGO SKU"]
        )

        ean_actual = normalizar_codigo(
            producto["CÓDIGO DE BARRAS"]
        )

        if sku and sku == sku_actual:
            motivos.append("SKU repetido")

        if ean and ean == ean_actual:
            motivos.append(
                "Código de barras repetido"
            )

        if nombre and nombre == nombre_actual:
            motivos.append("Nombre equivalente")

        if motivos:
            coincidencias.append({
                "indice": indice,
                "nombre": producto["NOMBRE*"],
                "motivos": motivos
            })

    return coincidencias


# ==========================================
# CREAR O EDITAR PRODUCTOS
# ==========================================

def aplicar_producto(datos, modo, indice=None):
    df_nuevo = st.session_state[
        "df_productos"
    ].copy()

    if modo == "crear":
        nuevo = pd.DataFrame(
            [datos],
            columns=COLUMNAS_TUU
        )

        df_nuevo = pd.concat(
            [df_nuevo, nuevo],
            ignore_index=True
        )

    elif modo == "editar":
        if indice is None or not (
            0 <= indice < len(df_nuevo)
        ):
            return False, "Producto no encontrado."

        for columna in COLUMNAS_TUU:
            df_nuevo.at[
                indice,
                columna
            ] = datos[columna]

    else:
        return False, "Operación no válida."

    # Reservar el SKU únicamente cuando se guarda el producto.
    # Los códigos introducidos manualmente conservan el comportamiento anterior.
    generado = st.session_state.get("sku_generado_guardar", False)
    if generado:
        try:
            registro_anterior = cargar_registro()
            categorias = obtener_categorias(
                df_nuevo, registro_anterior["categorias"]
            )
            sku = str(datos["CÓDIGO SKU"]).strip().upper()
            if not PATRON_SKU.fullmatch(sku):
                return False, "El SKU generado no tiene un formato válido."
            numero = int(sku[:3])
            nombre_categoria = normalizar_texto_sku(datos["CATEGORIAS"])
            if categorias.get(nombre_categoria) != numero:
                return False, "La categoría del SKU no coincide con su identificador. Genera nuevamente el SKU."
            for i, fila in st.session_state["df_productos"].iterrows():
                if modo == "editar" and i == indice:
                    continue
                if str(fila["CÓDIGO SKU"]).strip().upper() == sku:
                    return False, "El SKU ya existe. Genera uno nuevo."
            nuevo_registro = registrar_sku(registro_anterior, sku, categorias)
            correcto, error = guardar_registro(nuevo_registro)
            if not correcto:
                return False, error
        except (ValueError, RuntimeError, OSError) as error:
            return False, f"No se pudo registrar el SKU: {error}"

    correcto, error = guardar_cambios(df_nuevo)
    if correcto and st.session_state.get('barcode_generado_guardar', False):
        try:
            internos = cargar_codigos()
            internos.add(str(datos['CÓDIGO DE BARRAS']).strip())
            guardado, err_reg = guardar_codigos(internos)
            if not guardado:
                st.session_state['mensaje'] = f'Producto guardado; atención: no se registró el origen del código interno: {err_reg}'
        except RuntimeError as exc:
            st.session_state['mensaje'] = f'Producto guardado; atención: {exc}'
    if not correcto and generado:
        # Intentar restaurar el registro previo si falla el guardado Excel.
        restaurado, error_restauracion = guardar_registro(registro_anterior)
        if not restaurado:
            error = f"{error}. Además, no se pudo restaurar el registro SKU: {error_restauracion}"
    return correcto, error


def finalizar_guardado(datos, modo, indice=None):
    correcto, error = aplicar_producto(
        datos,
        modo,
        indice
    )

    if correcto:
        cambiar_vista("principal")

        st.session_state["mensaje"] = (
            "Producto guardado correctamente."
        )

        st.rerun()
    else:
        st.error(error)


# ==========================================
# FORMULARIO COMPARTIDO
# ==========================================

def mostrar_vista_formulario(modo, indice=None):
    df = st.session_state["df_productos"]

    producto = None

    if modo == "editar":
        if indice is None or not (
            0 <= indice < len(df)
        ):
            st.error("El producto ya no existe.")

            if st.button("← Volver al catálogo"):
                cambiar_vista("principal")
                st.rerun()

            return

        producto = df.iloc[indice]

    col_titulo, col_volver = st.columns([5, 1])

    with col_titulo:
        if modo == "crear":
            st.subheader("➕ Nuevo Producto")
        else:
            st.subheader("✏️ Editar Producto")
            st.caption(
                f"Producto: {producto['NOMBRE*']}"
            )

    with col_volver:
        if st.button(
            "← Volver",
            width="stretch"
        ):
            cambiar_vista("principal")
            st.rerun()

    clave_formulario = f"{modo}_{indice}_{st.session_state['version_formulario']}"
    with st.container(border=True):
        datos, enviado = mostrar_formulario_producto(
            producto=producto,
            clave=clave_formulario
        )

    solicitud_barras = st.session_state.pop("barcode_solicitud", None)
    if solicitud_barras and solicitud_barras["clave"] == clave_formulario:
        borrador_barras = solicitud_barras["datos"]
        if borrador_barras["CÓDIGO DE BARRAS"].strip():
            st.session_state['barcode_mensaje'] = ('warning', 'Este producto ya tiene código de barras. Bórralo antes de generar uno interno.')
            st.rerun()
        else:
            try:
                ocupados = df["CÓDIGO DE BARRAS"].fillna("").astype(str).tolist()
                nuevo_codigo = generar_ean13_interno(ocupados)
                borrador_barras["CÓDIGO DE BARRAS"] = nuevo_codigo
                st.session_state["version_formulario"] += 1
                nueva_clave = f"{modo}_{indice}_{st.session_state['version_formulario']}"
                st.session_state["sku_borrador"] = {
                    "clave": nueva_clave,
                    "datos": borrador_barras,
                    "abreviatura": solicitud_barras["abreviatura"]
                }
                # La vista previa SKU anterior sigue siendo válida si el SKU no cambió.
                previo = st.session_state.get("sku_previsualizado")
                if previo and previo.get("clave") == clave_formulario:
                    previo["clave"] = nueva_clave
                st.session_state['barcode_mensaje'] = ('success', f'Código interno generado: {nuevo_codigo}. Guarda el producto para registrarlo.')
                st.session_state['barcode_previsualizado'] = {'clave': nueva_clave, 'codigo': nuevo_codigo}
                st.rerun()
            except (ValueError, RuntimeError) as error:
                st.session_state['barcode_mensaje'] = ('error', f'No se pudo generar el código: {error}')
                st.rerun()

    solicitud = st.session_state.pop("sku_solicitud", None)
    if solicitud and solicitud["clave"] == clave_formulario:
        try:
            registro = cargar_registro()
            categorias = obtener_categorias(df, registro["categorias"])
            categoria = normalizar_texto_sku(solicitud["datos"]["CATEGORIAS"])
            if not categoria:
                raise ValueError("Debes indicar la categoría antes de generar el SKU.")
            # No considerar el SKU del mismo producto al editarlo.
            df_base = df.drop(index=indice) if modo == "editar" else df
            numero = categorias.get(categoria)
            reservados = obtener_correlativos_reservados(registro, numero) if numero else []
            sku, categorias = generar_sku(
                solicitud["datos"]["NOMBRE*"],
                solicitud["datos"]["MARCAS"],
                solicitud["datos"]["CATEGORIAS"],
                df_base, categorias=categorias,
                abreviatura_producto=solicitud["abreviatura"] or None,
                correlativos_reservados=reservados,
            )
            solicitud["datos"]["CÓDIGO SKU"] = sku
            st.session_state["sku_borrador"] = solicitud
            st.session_state["sku_previsualizado"] = {"clave": clave_formulario, "sku": sku}
            st.session_state["version_formulario"] += 1
            # Mantener el borrador asociado a la nueva clave del formulario.
            nueva_clave = f"{modo}_{indice}_{st.session_state['version_formulario']}"
            st.session_state["sku_borrador"]["clave"] = nueva_clave
            st.session_state["sku_previsualizado"]["clave"] = nueva_clave
            previo_barras = st.session_state.get('barcode_previsualizado')
            if previo_barras and previo_barras.get('clave') == clave_formulario:
                previo_barras['clave'] = nueva_clave
            st.session_state['sku_mensaje'] = ('success', f'SKU generado: {sku}. Se reservará al guardar.')
            st.rerun()
        except (ValueError, RuntimeError, OSError) as error:
            st.session_state['sku_mensaje'] = ('error', f'No se pudo generar el SKU: {error}')
            st.rerun()

    previo = st.session_state.get("sku_previsualizado")
    if previo and previo.get("clave") == clave_formulario:
        st.info(f"SKU generado (aún no reservado): {previo['sku']}")

    if enviado:
        coincidencias = buscar_coincidencias(
            datos,
            df,
            excluir=(
                indice if modo == "editar"
                else None
            )
        )

        if coincidencias:
            st.session_state["datos_pendientes"] = (
                datos.copy()
            )
            st.session_state["modo_pendiente"] = modo
            st.session_state["indice_pendiente"] = indice

        else:
            finalizar_guardado(
                datos,
                modo,
                indice
            )

    pendientes = st.session_state["datos_pendientes"]

    if pendientes is not None:
        coincidencias = buscar_coincidencias(
            pendientes,
            st.session_state["df_productos"],
            excluir=(
                indice if modo == "editar"
                else None
            )
        )

        st.warning(
            "⚠️ Se encontraron posibles duplicados. "
            "Revisa antes de guardar."
        )

        for coincidencia in coincidencias:
            st.write(
                f"**{coincidencia['nombre']}** — "
                + ", ".join(
                    coincidencia["motivos"]
                )
            )

        col_si, col_no = st.columns(2)

        with col_si:
            if st.button(
                "Guardar de todas formas",
                type="primary",
                width="stretch"
            ):
                finalizar_guardado(
                    pendientes,
                    st.session_state["modo_pendiente"],
                    st.session_state["indice_pendiente"]
                )

        with col_no:
            if st.button(
                "Cancelar guardado",
                width="stretch"
            ):
                limpiar_pendientes()
                st.rerun()


# ==========================================
# IMPORTAR CATALOGO
# ==========================================

with st.expander(
    "📂 Archivo Excel del catálogo",
    expanded=True
):
    modo_catalogo = st.radio(
        "¿Cómo deseas comenzar?",
        ["Importar Excel existente", "Crear catálogo nuevo"],
        key="modo_catalogo",
        horizontal=True,
    )

    if modo_catalogo == "Crear catálogo nuevo":
        nombre_nuevo = st.text_input(
            "Nombre del nuevo catálogo",
            value="Mi_catalogo_TUU",
            key="nombre_catalogo_nuevo",
        )
        if st.button("➕ Crear catálogo", type="primary"):
            nombre_limpio = Path(nombre_nuevo.strip()).stem.strip()
            if not nombre_limpio or nombre_limpio in (".", ".."):
                st.error("Escribe un nombre válido para el catálogo.")
            else:
                df_nuevo = crear_catalogo_vacio()
                ruta = crear_ruta_trabajo(nombre_limpio + ".xlsx")
                correcto, error = guardar_excel(df_nuevo, ruta)
                if not correcto:
                    st.error(error)
                else:
                    st.session_state["archivo_id"] = None
                    st.session_state["df_productos"] = df_nuevo
                    st.session_state["ruta_trabajo"] = str(ruta)
                    st.session_state["version_editor"] += 1
                    st.session_state["version_tabla"] += 1
                    st.session_state["filtro_edicion"] = "Todos los productos"
                    st.session_state["busqueda_edicion"] = ""
                    st.session_state["marca_catalogo"] = "Todas las marcas"
                    st.session_state["categoria_catalogo"] = "Todas las categorías"
                    st.session_state["busqueda_catalogo"] = ""
                    iniciar_historial(df_nuevo)
                    limpiar_duplicados()
                    cambiar_vista("principal")
                    st.session_state["mensaje"] = "Catálogo nuevo creado correctamente."
                    st.rerun()

    archivo = None
    if modo_catalogo == "Importar Excel existente":
        archivo = st.file_uploader(
            "Selecciona tu archivo Excel de TUU",
            type=["xlsx"],
            key="archivo_excel_catalogo",
            on_change=limpiar_sesion_al_quitar_excel,
        )

    if archivo is not None:
        identificador = hashlib.sha256(
            archivo.getvalue()
        ).hexdigest()

        if (
            st.session_state["archivo_id"]
            != identificador
        ):
            df_importado, error = cargar_excel(
                archivo
            )

            if error:
                st.error(error)
                st.stop()

            ruta = crear_ruta_trabajo(
                archivo.name
            )

            correcto, error = guardar_excel(
                df_importado,
                ruta
            )

            if not correcto:
                st.error(error)
                st.stop()

            df_importado = (
                df_importado.copy()
                .reset_index(drop=True)
                .fillna("")
                .astype(str)
            )

            st.session_state["archivo_id"] = (
                identificador
            )
            st.session_state["df_productos"] = (
                df_importado
            )
            st.session_state["ruta_trabajo"] = str(ruta)

            st.session_state["version_editor"] += 1
            st.session_state["version_tabla"] += 1

            st.session_state["filtro_edicion"] = (
                "Todos los productos"
            )
            st.session_state["busqueda_edicion"] = ""
            st.session_state["marca_catalogo"] = "Todas las marcas"
            st.session_state["categoria_catalogo"] = "Todas las categorías"
            st.session_state["busqueda_catalogo"] = ""

            iniciar_historial(df_importado)
            limpiar_duplicados()
            cambiar_vista("principal")

            st.success(
                "Catálogo importado correctamente."
            )


if st.session_state["df_productos"] is None:
    mensaje_inicio = st.session_state.get("mensaje_inicio")
    if mensaje_inicio:
        st.success(mensaje_inicio)
        st.session_state["mensaje_inicio"] = None

    st.info(
        "Importa un archivo Excel o crea un catálogo nuevo para comenzar."
    )
    st.stop()


# ==========================================
# INFORMACION GENERAL
# ==========================================

df = st.session_state["df_productos"]

col_total, col_ruta = st.columns([1, 4])

with col_total:
    st.metric("Total de productos", len(df))

with col_ruta:
    st.caption(
        "Copia de trabajo: "
        + str(st.session_state["ruta_trabajo"])
    )

if st.session_state["mensaje"]:
    st.success(st.session_state["mensaje"])
    st.session_state["mensaje"] = None

st.divider()


# ==========================================
# VISTA INDEPENDIENTE DEL FORMULARIO
# ==========================================

vista = st.session_state["vista"]

if vista in ("crear", "editar"):
    mostrar_vista_formulario(
        modo=vista,
        indice=st.session_state["indice_edicion"]
    )

    st.stop()


# ==========================================
# SECCION 1: CATALOGO
# ==========================================

with st.expander(
    "📋 Catálogo de productos",
    expanded=True
):
    col_titulo, col_nuevo = st.columns([5, 1])

    with col_titulo:
        st.caption(
            "Busca, selecciona y edita "
            "los productos de tu catálogo."
        )

    with col_nuevo:
        if st.button(
            "➕ Nuevo Producto",
            type="primary",
            width="stretch"
        ):
            cambiar_vista("crear")
            st.rerun()

    col_busqueda_catalogo, col_limpiar_catalogo = st.columns([5, 1])
    with col_busqueda_catalogo:
        busqueda = st.text_input(
            "🔎 Buscar producto",
            placeholder="Nombre, SKU o código de barras",
            key="busqueda_catalogo"
        )
    with col_limpiar_catalogo:
        st.write("")
        st.button(
            "🔄 Mostrar todos",
            key="limpiar_catalogo",
            width="stretch",
            on_click=limpiar_busqueda_catalogo,
        )

    df_filtrado = st.session_state[
        "df_productos"
    ]

    opciones_marcas_catalogo = opciones_columna(df_filtrado, "MARCAS", "Todas las marcas")
    opciones_categorias_catalogo = opciones_columna(df_filtrado, "CATEGORIAS", "Todas las categorías")
    if st.session_state.get("marca_catalogo") not in opciones_marcas_catalogo:
        st.session_state["marca_catalogo"] = "Todas las marcas"
    if st.session_state.get("categoria_catalogo") not in opciones_categorias_catalogo:
        st.session_state["categoria_catalogo"] = "Todas las categorías"

    col_marca_catalogo, col_categoria_catalogo = st.columns(2)
    with col_marca_catalogo:
        marca_catalogo = st.selectbox(
            "Filtrar por marca",
            options=opciones_marcas_catalogo,
            key="marca_catalogo",
        )
    with col_categoria_catalogo:
        categoria_catalogo = st.selectbox(
            "Filtrar por categoría",
            options=opciones_categorias_catalogo,
            key="categoria_catalogo",
        )

    origen_catalogo = st.selectbox('Filtrar por origen del código de barras',
        ['Todos los códigos', 'Código comercial / externo', 'Código interno generado', 'Sin código de barras'],
        key='origen_catalogo')
    try:
        df_filtrado = filtrar_por_origen(df_filtrado, origen_catalogo, cargar_codigos())
    except RuntimeError as exc:
        st.error(str(exc))

    if marca_catalogo != "Todas las marcas":
        df_filtrado = df_filtrado.loc[
            df_filtrado["MARCAS"].fillna("").astype(str).str.strip()
            .str.casefold().eq(marca_catalogo.casefold())
        ]
    if categoria_catalogo != "Todas las categorías":
        df_filtrado = df_filtrado.loc[
            df_filtrado["CATEGORIAS"].fillna("").astype(str).str.strip()
            .str.casefold().eq(categoria_catalogo.casefold())
        ]

    if busqueda.strip():
        texto = busqueda.strip()

        mascara = (
            df_filtrado["NOMBRE*"]
            .str.contains(
                texto,
                case=False,
                regex=False,
                na=False
            )
            |
            df_filtrado["CÓDIGO SKU"]
            .str.contains(
                texto,
                case=False,
                regex=False,
                na=False
            )
            |
            df_filtrado["CÓDIGO DE BARRAS"]
            .str.contains(
                texto,
                case=False,
                regex=False,
                na=False
            )
        )

        df_filtrado = df_filtrado.loc[mascara]

    st.caption(
        f"Mostrando {len(df_filtrado)} "
        f"de {len(df)} productos"
    )

    columnas_visibles = [
        "NOMBRE*",
        "PRECIO*",
        "CÓDIGO SKU",
        "CÓDIGO DE BARRAS",
        "MARCAS",
        "CATEGORIAS"
    ]

    tabla_catalogo = df_filtrado[
        columnas_visibles
    ].copy()

    indices_reales = list(
        tabla_catalogo.index
    )

    tabla_catalogo = (
        tabla_catalogo.reset_index(drop=True)
    )

    clave_busqueda = hashlib.md5(
        (busqueda.strip().lower() + "|" + marca_catalogo + "|" + categoria_catalogo + "|" + origen_catalogo).encode("utf-8")
    ).hexdigest()[:10]

    seleccion = st.dataframe(
        tabla_catalogo,
        hide_index=True,
        width="stretch",
        height=400,
        on_select="rerun",
        selection_mode="single-row",
        key=(
            f"tabla_catalogo_"
            f"{st.session_state['version_tabla']}_"
            f"{clave_busqueda}"
        )
    )

    filas_seleccionadas = (
        seleccion.selection.rows
    )

    indice_real = None

    if filas_seleccionadas:
        posicion_fila = filas_seleccionadas[0]

        if 0 <= posicion_fila < len(indices_reales):
            indice_real = indices_reales[
                posicion_fila
            ]

    col_editar, col_seleccion = st.columns([2, 5])

    with col_editar:
        if st.button(
            "✏️ Editar seleccionado",
            disabled=indice_real is None,
            type="primary",
            width="stretch"
        ):
            cambiar_vista(
                "editar",
                indice_real
            )
            st.rerun()

    with col_seleccion:
        if indice_real is not None:
            producto = (
                st.session_state["df_productos"]
                .iloc[indice_real]
            )

            st.caption(
                f"Seleccionado: {producto['NOMBRE*']}"
            )
        else:
            st.caption(
                "Selecciona una fila para editar."
            )


# ==========================================
# SECCION 2: EDICION RAPIDA CON FILTROS
# ==========================================

with st.expander(
    "📝 Edición rápida de la planilla",
    expanded=True
):
    st.caption(
        "Filtra los productos y edita directamente "
        "sus datos. Los cambios se guardan "
        "automáticamente."
    )

    # --------------------------------------
    # FILTROS
    # --------------------------------------

    df_actual = st.session_state["df_productos"]

    col_filtro, col_buscar, col_limpiar = st.columns([2, 3, 1.3])
    with col_filtro:
        filtro = st.selectbox(
            "Filtrar productos",
            options=OPCIONES_FILTRO,
            key="filtro_edicion"
        )
    with col_buscar:
        busqueda_rapida = st.text_input(
            "Buscar dentro de los resultados",
            placeholder="Nombre, SKU o código de barras",
            key="busqueda_edicion"
        )
    with col_limpiar:
        st.write("")
        st.button(
            "🔄 Mostrar todos",
            key="limpiar_edicion",
            width="stretch",
            on_click=limpiar_filtros_edicion,
        )

    origen_edicion = st.selectbox('Filtrar por origen del código de barras',
        ['Todos los códigos', 'Código comercial / externo', 'Código interno generado', 'Sin código de barras'],
        key='origen_edicion')

    df_mostrado = filtrar_productos(df_actual, filtro, busqueda_rapida)
    try:
        df_mostrado = filtrar_por_origen(df_mostrado, origen_edicion, cargar_codigos())
    except RuntimeError as exc:
        st.error(str(exc))

    st.caption(
        f"Mostrando {len(df_mostrado)} "
        f"de {len(df_actual)} productos"
    )

    if filtro in CAMPOS_FALTANTES:
        if filtro not in ("Sin nombre", "Sin precio"):
            st.caption(
                "Este campo es opcional para la "
                "validación básica de TUU, pero "
                "puedes revisarlo para mejorar "
                "tu catálogo."
            )

    # --------------------------------------
    # LECTOR DE CODIGOS EN EDICION RAPIDA
    # --------------------------------------
    # Algunos lectores envian Enter al terminar. st.data_editor no
    # siempre confirma la celda con ese Enter. Este campo dedicado
    # permite escanear sin abrir el formulario de cada producto.
    with st.expander("🔫 Escanear código de barras en edición rápida"):
        st.caption(
            "Selecciona el producto de la planilla y escanea. "
            "El lector debe enviar Enter al finalizar; el código "
            "se guarda automáticamente. También puedes escribirlo."
        )
        if df_mostrado.empty:
            st.info("Primero agrega un producto a la planilla.")
        else:
            indices_visibles = list(df_mostrado.index)

            def etiqueta_producto(indice):
                fila = df_actual.loc[indice]
                nombre = str(fila.get("NOMBRE*", "")).strip() or "Sin nombre"
                codigo = str(fila.get("CÓDIGO DE BARRAS", "")).strip()
                return f"Fila {indice + 1}: {nombre}" + (f" | {codigo}" if codigo else "")

            def guardar_lectura_rapida():
                codigo = str(st.session_state.get("lectura_rapida_codigo", "")).strip()
                indice = st.session_state.get("lectura_rapida_producto")
                if not codigo or indice not in st.session_state["df_productos"].index:
                    return
                actual = st.session_state["df_productos"]
                # Evita asignar accidentalmente el mismo código a otro producto.
                repetidos = actual.index[
                    actual["CÓDIGO DE BARRAS"].astype(str).str.strip().eq(codigo)
                    & (actual.index != indice)
                ]
                if len(repetidos):
                    st.session_state["lectura_rapida_mensaje"] = (
                        "⚠️ Código ya asignado a otra fila. No se guardó."
                    )
                    return
                nuevo = actual.copy()
                nuevo.at[indice, "CÓDIGO DE BARRAS"] = codigo
                correcto, error = guardar_cambios(nuevo)
                if correcto:
                    st.session_state["lectura_rapida_codigo"] = ""
                    st.session_state["lectura_rapida_mensaje"] = (
                        f"✅ Código {codigo} guardado en {etiqueta_producto(indice)}"
                    )
                else:
                    st.session_state["lectura_rapida_mensaje"] = f"❌ {error}"

            st.selectbox(
                "Producto que recibirá el código",
                options=indices_visibles,
                format_func=etiqueta_producto,
                key="lectura_rapida_producto",
            )
            st.text_input(
                "Escanear código de barras aquí",
                key="lectura_rapida_codigo",
                placeholder="Haz clic aquí y utiliza la pistola lectora",
                on_change=guardar_lectura_rapida,
            )
            if st.session_state.get("lectura_rapida_mensaje"):
                st.caption(st.session_state["lectura_rapida_mensaje"])

    # --------------------------------------
    # CONTROLES
    # --------------------------------------

    historial = st.session_state["historial"]
    posicion = st.session_state[
        "posicion_historial"
    ]

    # Solo permitimos agregar/eliminar filas
    # cuando no existe ningún filtro.
    vista_completa = (
        filtro == "Todos los productos"
        and not busqueda_rapida.strip()
        and origen_edicion == "Todos los códigos"
    )

    col_agregar, col_atras, col_adelante, col_estado = (
        st.columns([1.7, 1.3, 1.3, 3])
    )

    with col_agregar:
        if st.button(
            "➕ Agregar fila",
            type="primary",
            width="stretch",
            disabled=not vista_completa
        ):
            df_nuevo = df_actual.copy()

            nueva_fila = {
                columna: ""
                for columna in COLUMNAS_TUU
            }

            nueva_fila["¿EXENTO?*"] = "N"
            nueva_fila["¿HABILITADO?*"] = "S"

            df_nuevo = pd.concat(
                [
                    df_nuevo,
                    pd.DataFrame(
                        [nueva_fila],
                        columns=COLUMNAS_TUU
                    )
                ],
                ignore_index=True
            )

            correcto, error = guardar_cambios(
                df_nuevo
            )

            if correcto:
                st.toast(
                    "Fila agregada al final",
                    icon="✅"
                )
                st.rerun()
            else:
                st.error(error)

    with col_atras:
        if st.button(
            "↶ Deshacer",
            disabled=posicion == 0,
            width="stretch"
        ):
            correcto, error = mover_historial(-1)

            if correcto:
                st.rerun()
            else:
                st.error(error)

    with col_adelante:
        if st.button(
            "↷ Rehacer",
            disabled=(
                posicion >= len(historial) - 1
            ),
            width="stretch"
        ):
            correcto, error = mover_historial(1)

            if correcto:
                st.rerun()
            else:
                st.error(error)

    with col_estado:
        st.caption(
            f"{posicion} para deshacer | "
            f"{len(historial) - posicion - 1} "
            "para rehacer"
        )

    # --------------------------------------
    # TABLA EDITABLE
    # --------------------------------------

    if df_mostrado.empty and not vista_completa:
        st.info(
            "No se encontraron productos "
            "con los filtros seleccionados."
        )

    else:
        version = st.session_state[
            "version_editor"
        ]

        # Clave diferente por filtro/búsqueda
        # para evitar estados anteriores
        # del editor.
        identidad_filtro = (
            f"{filtro}|{busqueda_rapida}"
        )

        clave_filtro = hashlib.md5(
            identidad_filtro.encode("utf-8")
        ).hexdigest()[:10]

        # En un catálogo nuevo, mostrar una primera fila editable
        # sin guardarla como producto hasta que el usuario escriba datos.
        primera_fila = df_mostrado.empty and vista_completa
        if primera_fila:
            fila_inicial = {columna: "" for columna in COLUMNAS_TUU}
            fila_inicial["¿EXENTO?*"] = "N"
            fila_inicial["¿HABILITADO?*"] = "S"
            df_para_editar = pd.DataFrame([fila_inicial], columns=COLUMNAS_TUU)
        else:
            df_para_editar = df_mostrado.copy()

        df_editado = st.data_editor(
            df_para_editar,
            width="stretch",
            hide_index=True,
            num_rows=(
                "dynamic"
                if vista_completa
                else "fixed"
            ),
            key=(
                f"editor_productos_"
                f"{version}_{clave_filtro}"
            )
        )

        if primera_fila:
            # Los valores predeterminados S/N no cuentan como producto.
            columnas_contenido = [
                columna for columna in COLUMNAS_TUU
                if columna not in ("¿EXENTO?*", "¿HABILITADO?*")
            ]
            hay_contenido = any(
                str(valor).strip()
                for _, fila in df_editado.iterrows()
                for valor in (fila[columna] for columna in columnas_contenido)
                if pd.notna(valor)
            )
            hay_cambios = hay_contenido
        else:
            hay_cambios = not df_editado.equals(df_mostrado)

        if hay_cambios:

            if vista_completa:
                # Sin filtros, se puede agregar,
                # editar o eliminar filas.
                # No registrar filas sin información de producto
                # al iniciar un catálogo desde cero.
                if primera_fila:
                    columnas_contenido = [
                        c for c in COLUMNAS_TUU
                        if c not in ("¿EXENTO?*", "¿HABILITADO?*")
                    ]
                    mascara = df_editado[columnas_contenido].apply(
                        lambda fila: any(
                            pd.notna(v) and str(v).strip()
                            for v in fila
                        ),
                        axis=1
                    )
                    df_a_guardar = df_editado.loc[mascara].copy()
                else:
                    df_a_guardar = df_editado

                correcto, error = guardar_cambios(
                    df_a_guardar
                )

            else:
                # Con filtros, solo se modifican
                # las filas originales visibles.
                correcto, error = (
                    guardar_edicion_filtrada(
                        df_actual,
                        df_mostrado,
                        df_editado
                    )
                )

            if correcto:
                st.toast(
                    "Cambios guardados",
                    icon="✅"
                )
                st.rerun()
            else:
                st.error(error)

        else:
            st.caption(
                "Todos los cambios están guardados."
            )

    if not vista_completa:
        st.info(
            "Estás trabajando con una vista filtrada. "
            "Puedes editar las celdas, pero para "
            "agregar o eliminar filas debes "
            "seleccionar «Todos los productos» "
            "y limpiar la búsqueda."
        )

    else:
        st.info(
            "Puedes editar, agregar o eliminar "
            "filas directamente en la tabla. "
            "Los cambios se guardan automáticamente."
        )


# ==========================================
# NUEVA SECCION: GENERACION MASIVA
# ==========================================

def _filtrar_masivo(df_base, prefijo):
    col_cat, col_marca = st.columns(2)
    with col_cat:
        cat = st.selectbox('Filtrar por categoría', opciones_columna(df_base, 'CATEGORIAS', 'Todas las categorías'), key=f'{prefijo}_categoria')
    with col_marca:
        marca = st.selectbox('Filtrar por marca', opciones_columna(df_base, 'MARCAS', 'Todas las marcas'), key=f'{prefijo}_marca')
    filtrado = df_base.copy()
    if cat != 'Todas las categorías':
        filtrado = filtrado.loc[filtrado['CATEGORIAS'].astype(str).str.strip().str.casefold().eq(cat.casefold())]
    if marca != 'Todas las marcas':
        filtrado = filtrado.loc[filtrado['MARCAS'].astype(str).str.strip().str.casefold().eq(marca.casefold())]
    return filtrado


def _vista_previa(df, indices, columna):
    vista = df.loc[indices, ['NOMBRE*', 'MARCAS', 'CATEGORIAS', columna]].copy()
    vista.index.name = 'Fila interna'
    st.dataframe(vista, width="stretch", height=min(360, 40 + 36 * max(len(vista), 1)))


with st.expander('🔢 Generación masiva de SKU', expanded=True):
    st.subheader('Generar SKU faltantes')
    st.caption('Puedes procesar todo el catálogo o limitarlo por categoría y marca. Nunca se sobrescriben SKU existentes.')
    df_m = st.session_state['df_productos']
    filtrado = _filtrar_masivo(df_m, 'masivo_sku')
    sin_sku = filtrado.loc[filtrado['CÓDIGO SKU'].fillna('').astype(str).str.strip().eq('')]
    # Mostrar TODOS los productos sin SKU, incluidos los que requieren correcciones.
    # La descripción se exige aquí para asegurar un catálogo bien documentado;
    # el algoritmo SKU utiliza nombre, marca y categoría.
    requisitos_sku = {
        'NOMBRE*': 'nombre',
        'MARCAS': 'marca',
        'CATEGORIAS': 'categoría',
        'Descripción': 'descripción',
    }

    def faltantes_sku(fila):
        return [etiqueta for columna, etiqueta in requisitos_sku.items()
                if not str(fila.get(columna, '') or '').strip()]

    tabla_sku = sin_sku[['NOMBRE*', 'Descripción', 'MARCAS', 'CATEGORIAS', 'CÓDIGO SKU']].copy()
    tabla_sku.insert(0, 'Fila Excel', tabla_sku.index + 2)
    tabla_sku['Estado / información faltante'] = [
        ('Falta: ' + ', '.join(faltantes_sku(fila))) if faltantes_sku(fila) else 'Listo para generar'
        for _, fila in sin_sku.iterrows()
    ]
    incompletos = sin_sku.loc[[i for i, fila in sin_sku.iterrows() if faltantes_sku(fila)]]
    validos = sin_sku.drop(index=incompletos.index)

    st.caption(f'{len(sin_sku)} productos sin SKU en el filtro actual: '
               f'{len(validos)} listos y {len(incompletos)} con información faltante.')
    if sin_sku.empty:
        st.info('No hay productos sin SKU con los filtros seleccionados.')
    else:
        st.caption('Puedes corregir nombre, descripción, marca y categoría directamente en la tabla. '
                   'Después presiona «Guardar correcciones».')
        editada_sku = st.data_editor(
            tabla_sku.reset_index(drop=True),
            key=f"tabla_editar_sku_{st.session_state['version_editor']}",
            hide_index=True,
            width='stretch',
            height=min(450, 75 + 35 * len(tabla_sku)),
            num_rows='fixed',
            disabled=['Fila Excel', 'CÓDIGO SKU', 'Estado / información faltante'],
        )
        columnas_editables = ['NOMBRE*', 'Descripción', 'MARCAS', 'CATEGORIAS']
        hay_cambios_sku = any(
            str(editada_sku.at[pos, columna]) != str(tabla_sku.iloc[pos][columna])
            for pos in range(len(tabla_sku)) for columna in columnas_editables
        )
        # Todas las vistas comparten el mismo DataFrame de sesión y la
        # misma copia de trabajo Excel. Guardar inmediatamente los cambios
        # de la tabla evita que una vista quede desactualizada.
        if hay_cambios_sku:
            copia_corregida = df_m.copy(deep=True)
            for pos, idx in enumerate(tabla_sku.index):
                for columna in columnas_editables:
                    valor = editada_sku.at[pos, columna]
                    copia_corregida.at[idx, columna] = (
                        "" if pd.isna(valor) else str(valor).strip()
                    )
            ok, error = guardar_cambios(copia_corregida)
            if ok:
                st.session_state['masivo_confirmacion'] = None
                st.toast('Corrección guardada en el catálogo único', icon='✅')
                st.rerun()
            else:
                st.error(f'No se pudieron guardar las correcciones: {error}')
        else:
            st.caption('Los datos mostrados están sincronizados con el catálogo.')
    if not incompletos.empty:
        st.warning('Hay productos incompletos. Completa los campos indicados directamente '
                   'en la tabla y guarda las correcciones antes de generar los SKU.')

    if st.button('Preparar generación masiva de SKU',
                 disabled=sin_sku.empty, key='preparar_sku'):
        if not incompletos.empty:
            st.error(f'No se puede generar: {len(incompletos)} productos tienen datos '
                     'faltantes. Revisa la columna «Estado / información faltante».')
            st.session_state['masivo_confirmacion'] = None
        else:
            try:
                registro = cargar_registro()
                categorias = obtener_categorias(df_m, registro['categorias'])
                copia = df_m.copy()
                cambios = {}
                reservados = {int(k): set(v) for k, v in registro['correlativos'].items()}
                for idx in validos.index:
                    fila = copia.loc[idx]
                    numero_cat = categorias.get(normalizar_texto_sku(fila['CATEGORIAS']))
                    sku, categorias = generar_sku(fila['NOMBRE*'], fila['MARCAS'], fila['CATEGORIAS'], copia,
                        categorias=categorias, correlativos_reservados=reservados.get(numero_cat, set()))
                    copia.at[idx, 'CÓDIGO SKU'] = sku
                    cambios[int(idx)] = sku
                    numero_cat = int(sku[:3]); reservados.setdefault(numero_cat, set()).add(int(sku[-5:]))
                st.session_state['masivo_confirmacion'] = {'tipo': 'sku', 'cambios': cambios, 'categorias': categorias}
            except (ValueError, RuntimeError, OSError) as exc:
                st.error(f'No se pudo preparar la generación: {exc}')
    pendiente = st.session_state.get('masivo_confirmacion')
    if pendiente and pendiente['tipo'] == 'sku':
        st.warning(f"Se generarán {len(pendiente['cambios'])} SKU. Revisa antes de confirmar.")
        st.dataframe(pd.DataFrame([{'Producto': df_m.at[i, 'NOMBRE*'], 'Nuevo SKU': sku} for i, sku in pendiente['cambios'].items()]), hide_index=True)
        if st.button('Confirmar y guardar SKU', type='primary', key='confirmar_sku'):
            try:
                registro_anterior = cargar_registro()
                nuevo_registro = registro_anterior
                for sku in pendiente['cambios'].values():
                    nuevo_registro = registrar_sku(nuevo_registro, sku, pendiente['categorias'])
                copia = df_m.copy()
                for idx, sku in pendiente['cambios'].items():
                    if faltantes_sku(copia.loc[idx]):
                        raise ValueError('Hay datos incompletos desde la vista previa. Prepara nuevamente.')
                    if str(copia.at[idx, 'CÓDIGO SKU']).strip():
                        raise ValueError('Un producto cambió desde la vista previa. Prepara nuevamente.')
                    copia.at[idx, 'CÓDIGO SKU'] = sku
                ok, err = guardar_registro(nuevo_registro)
                if not ok:
                    raise RuntimeError(err)
                ok, err = guardar_cambios(copia)
                if not ok:
                    guardar_registro(registro_anterior)
                    raise RuntimeError(err)
                st.session_state['masivo_confirmacion'] = None
                st.session_state['mensaje'] = 'SKU masivos guardados correctamente.'
                st.rerun()
            except (ValueError, RuntimeError, OSError) as exc:
                st.error(f'No se pudo guardar: {exc}')
        if st.button('Cancelar generación SKU', key='cancelar_sku'):
            st.session_state['masivo_confirmacion'] = None
            st.rerun()


with st.expander('🏷️ Generación masiva de códigos de barras internos', expanded=True):
    st.subheader('Generar códigos de barras internos')
    st.caption('Solo para productos sin código. Desmarca los que podrían tener un código comercial pendiente de registrar.')
    df_m = st.session_state['df_productos']
    filtrado = _filtrar_masivo(df_m, 'masivo_barras')
    sin_barras = filtrado.loc[filtrado['CÓDIGO DE BARRAS'].fillna('').astype(str).str.strip().eq('')]
    # Solo los códigos INTERNOS exigen una ficha completa. Los comerciales
    # continúan ingresándose por teclado o lector en edición rápida.
    def faltantes_codigo_interno(fila):
        faltan = faltantes_sku(fila)
        if not str(fila.get('CÓDIGO SKU', '')).strip():
            faltan.append('SKU')
        faltan.extend(obtener_problemas(fila))
        return faltan

    st.caption('Para generar códigos internos se requiere nombre, descripción, marca, categoría, SKU, precio válido, exento y habilitado. Los códigos comerciales se ingresan con lector o teclado.')
    if sin_barras.empty:
        st.info('No hay productos sin código de barras con estos filtros.')
    else:
        # Conservar los índices originales para actualizar las filas correctas.
        columnas_corregibles = ['NOMBRE*', 'Descripción', 'MARCAS', 'CATEGORIAS', 'CÓDIGO SKU', 'PRECIO*', '¿EXENTO?*', '¿HABILITADO?*']
        tabla = sin_barras[columnas_corregibles].copy()
        tabla.insert(0, 'Generar', True)
        tabla.insert(0, 'Índice', tabla.index)
        tabla['Estado / información faltante'] = [
            'Falta: ' + ', '.join(faltantes_codigo_interno(fila)) if faltantes_codigo_interno(fila)
            else 'Listo para generar' for _, fila in sin_barras.iterrows()
        ]
        st.caption('Corrige los campos directamente aquí. Los cambios se guardan en el mismo catálogo que utilizan todas las vistas.')
        seleccionados = st.data_editor(
            tabla.reset_index(drop=True), hide_index=True, width='stretch',
            disabled=['Índice', 'Estado / información faltante'],
            column_config={'Generar': st.column_config.CheckboxColumn('Generar código')},
            key=f'editor_seleccion_barras_{st.session_state["version_tabla"]}_{hashlib.md5(str(tuple(sin_barras.index)).encode()).hexdigest()[:10]}'
        )
        hay_edicion = any(
            str(seleccionados.at[pos, col]) != str(tabla.iloc[pos][col])
            for pos in range(len(tabla)) for col in columnas_corregibles
        )
        if hay_edicion:
            copia_corregida = df_m.copy(deep=True)
            for pos, idx in enumerate(sin_barras.index):
                for col in columnas_corregibles:
                    valor = seleccionados.at[pos, col]
                    copia_corregida.at[idx, col] = '' if pd.isna(valor) else str(valor).strip()
            ok, error = guardar_cambios(copia_corregida)
            if ok:
                st.session_state['masivo_confirmacion'] = None
                st.toast('Correcciones guardadas en el catálogo único', icon='✅')
                st.rerun()
            else:
                st.error(f'No se pudieron guardar los cambios: {error}')
        indices = [int(i) for i in seleccionados.loc[seleccionados['Generar'], 'Índice']]
        incompletos_seleccionados = {i: faltantes_codigo_interno(df_m.loc[i]) for i in indices if faltantes_codigo_interno(df_m.loc[i])}
        st.caption(f'{len(indices)} de {len(sin_barras)} productos seleccionados. '
                   f'{len(incompletos_seleccionados)} seleccionados requieren correcciones.')
        if st.button('Preparar códigos internos', disabled=not indices, key='preparar_barras'):
            if incompletos_seleccionados:
                st.error(f'No se pueden generar códigos: {len(incompletos_seleccionados)} productos seleccionados tienen datos faltantes. Corrige la columna «Estado / información faltante» o desmárcalos.')
                st.session_state['masivo_confirmacion'] = None
            else:
                try:
                    ocupados = set(df_m['CÓDIGO DE BARRAS'].fillna('').astype(str)) | cargar_codigos()
                    cambios = {}
                    for idx in indices:
                        codigo = generar_ean13_interno(ocupados)
                        cambios[idx] = codigo
                        ocupados.add(codigo)
                    st.session_state['masivo_confirmacion'] = {'tipo': 'barras', 'cambios': cambios}
                except (ValueError, RuntimeError, OSError) as exc:
                    st.error(f'No se pudo preparar: {exc}')
    pendiente = st.session_state.get('masivo_confirmacion')
    if pendiente and pendiente['tipo'] == 'barras':
        st.warning(f"Se generarán {len(pendiente['cambios'])} códigos internos. No se modificarán códigos existentes.")
        st.dataframe(pd.DataFrame([{'Producto': df_m.at[i, 'NOMBRE*'], 'Nuevo código': cod} for i, cod in pendiente['cambios'].items()]), hide_index=True)
        if st.button('Confirmar y guardar códigos internos', type='primary', key='confirmar_barras'):
            try:
                copia = df_m.copy()
                for idx, codigo in pendiente['cambios'].items():
                    faltantes = faltantes_codigo_interno(copia.loc[idx])
                    if faltantes:
                        raise ValueError(f'Fila {idx + 2}: faltan {", ".join(faltantes)}. Prepara nuevamente.')
                    if str(copia.at[idx, 'CÓDIGO DE BARRAS']).strip():
                        raise ValueError('Un producto ya tiene código. Prepara nuevamente.')
                    copia.at[idx, 'CÓDIGO DE BARRAS'] = codigo
                internos_antes = cargar_codigos()
                ok, err = guardar_codigos(internos_antes | set(pendiente['cambios'].values()))
                if not ok:
                    raise RuntimeError(err)
                ok, err = guardar_cambios(copia)
                if not ok:
                    guardar_codigos(internos_antes)
                    raise RuntimeError(err)
                st.session_state['masivo_confirmacion'] = None
                st.session_state['mensaje'] = 'Códigos internos guardados correctamente.'
                st.rerun()
            except (ValueError, RuntimeError, OSError) as exc:
                st.error(f'No se pudo guardar: {exc}')
        if st.button('Cancelar generación de códigos', key='cancelar_barras'):
            st.session_state['masivo_confirmacion'] = None
            st.rerun()


# ==========================================
# SIMULADOR DE CAJA — SOLO PRUEBAS
# ==========================================
# Esta vista consulta el mismo DataFrame que todas las demás secciones.
# No registra ventas ni modifica inventario, precios o archivos.

def _precio_simulador(valor):
    texto = str(valor).strip().replace("$", "").replace(" ", "")
    if not texto:
        raise ValueError("Precio faltante")
    # Formatos habituales: 2500, 2.500, 2.500,00 y 2500.00.
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    elif texto.count(".") == 1 and len(texto.split(".")[1]) == 3:
        texto = texto.replace(".", "")
    precio = float(texto)
    if not math.isfinite(precio) or precio < 0:
        raise ValueError("Precio inválido")
    return precio


def _formato_pesos_simulador(precio):
    return "$" + f"{precio:,.0f}".replace(",", ".")


def _registrar_problema_simulador(codigo, encontrados, mensaje):
    """Conserva cada lectura problemática sin alterar el catálogo."""
    fila = encontrados.iloc[0]
    st.session_state["caja_encontrados_incompletos"].append({
        "Código de barras": codigo,
        "Producto": str(fila.get("NOMBRE*", "")).strip() or "(Sin nombre)",
        "SKU": str(fila.get("CÓDIGO SKU", "")).strip(),
        "Marca": str(fila.get("MARCAS", "")).strip(),
        "Categoría": str(fila.get("CATEGORIAS", "")).strip(),
        "Problema": mensaje,
    })


def _procesar_lectura_simulador():
    codigo = str(st.session_state.get("caja_lectura", "")).strip()
    if not codigo:
        return
    st.session_state["caja_lectura"] = ""
    st.session_state["caja_lecturas_realizadas"] = st.session_state.get("caja_lecturas_realizadas", 0) + 1
    df_caja = st.session_state["df_productos"]
    encontrados = df_caja.loc[
        df_caja["CÓDIGO DE BARRAS"].fillna("").astype(str).str.strip().eq(codigo)
    ]
    if encontrados.empty:
        st.session_state["caja_mensaje"] = ("error", f"Código {codigo}: no existe en el catálogo.")
        st.session_state["caja_no_encontrados"].append(codigo)
        return
    if len(encontrados) > 1:
        mensaje = f"Código duplicado: aparece en {len(encontrados)} productos. Corrige el duplicado."
        for indice in range(len(encontrados)):
            _registrar_problema_simulador(codigo, encontrados.iloc[[indice]], mensaje)
        st.session_state["caja_mensaje"] = ("error", f"Código {codigo}: {mensaje}")
        return
    fila = encontrados.iloc[0]
    nombre = str(fila["NOMBRE*"]).strip()
    if not nombre:
        mensaje = "Falta nombre del producto."
        _registrar_problema_simulador(codigo, encontrados, mensaje)
        st.session_state["caja_mensaje"] = ("error", f"Código {codigo}: {mensaje}")
        return
    if str(fila["¿HABILITADO?*"]).strip().upper() != "S":
        mensaje = "Producto no habilitado."
        _registrar_problema_simulador(codigo, encontrados, mensaje)
        st.session_state["caja_mensaje"] = ("warning", f"{nombre}: {mensaje}")
        return
    try:
        precio = _precio_simulador(fila["PRECIO*"])
    except (ValueError, TypeError):
        mensaje = "Precio faltante o inválido."
        _registrar_problema_simulador(codigo, encontrados, mensaje)
        st.session_state["caja_mensaje"] = ("error", f"{nombre}: {mensaje}")
        return
    carrito = st.session_state["caja_carrito"]
    if codigo in carrito:
        carrito[codigo]["cantidad"] += 1
    else:
        carrito[codigo] = {
            "nombre": nombre, "sku": str(fila["CÓDIGO SKU"]).strip(),
            "marca": str(fila["MARCAS"]).strip(),
            "categoria": str(fila["CATEGORIAS"]).strip(),
            "precio": precio, "cantidad": 1,
        }
    st.session_state["caja_mensaje"] = ("success", f"Agregado: {nombre} — {_formato_pesos_simulador(precio)}")


for _clave, _valor in (
    ("caja_lectura", ""), ("caja_carrito", {}),
    ("caja_no_encontrados", []), ("caja_encontrados_incompletos", []),
    ("caja_mensaje", None), ("caja_lecturas_realizadas", 0),
):
    if _clave not in st.session_state:
        st.session_state[_clave] = _valor

def _reiniciar_prueba_simulador():
    """Se ejecuta antes de crear el campo del lector en el siguiente rerun."""
    st.session_state["caja_carrito"] = {}
    st.session_state["caja_no_encontrados"] = []
    st.session_state["caja_encontrados_incompletos"] = []
    st.session_state["caja_mensaje"] = None
    st.session_state["caja_lectura"] = ""
    st.session_state["caja_lecturas_realizadas"] = 0


with st.expander("🛒 Simulador de caja — Probar catálogo", expanded=True):
    st.caption("Prueba el lector de códigos de barras sin registrar ventas ni alterar el catálogo. "
               "El lector debe enviar Enter al terminar; también puedes escribir el código.")
    st.text_input(
        "🔫 Escanear código de barras", key="caja_lectura",
        placeholder="Haz clic aquí, escanea y presiona Enter",
        on_change=_procesar_lectura_simulador,
    )
    mensaje_caja = st.session_state.get("caja_mensaje")
    if mensaje_caja:
        nivel, mensaje = mensaje_caja
        if nivel == "success":
            st.success(mensaje)
        elif nivel == "warning":
            st.warning(mensaje)
        else:
            st.error(mensaje)

    carrito_caja = st.session_state["caja_carrito"]
    if carrito_caja:
        filas_caja = [
            {"Producto": item["nombre"], "SKU": item["sku"],
             "Marca": item["marca"], "Categoría": item["categoria"],
             "Código de barras": codigo, "Cantidad": item["cantidad"],
             "Precio unitario": _formato_pesos_simulador(item["precio"]),
             "Subtotal": _formato_pesos_simulador(item["precio"] * item["cantidad"])}
            for codigo, item in carrito_caja.items()
        ]
        st.dataframe(pd.DataFrame(filas_caja), hide_index=True, width="stretch")
        unidades_caja = sum(item["cantidad"] for item in carrito_caja.values())
        total_caja = sum(item["precio"] * item["cantidad"] for item in carrito_caja.values())
        col_unidades, col_total_caja = st.columns(2)
        col_unidades.metric("Unidades escaneadas", unidades_caja)
        col_total_caja.metric("Total simulado", _formato_pesos_simulador(total_caja))
        col_quitar, col_vaciar = st.columns([3, 1])
        with col_quitar:
            codigo_quitar = st.selectbox(
                "Quitar producto del carrito", list(carrito_caja),
                format_func=lambda c: f"{carrito_caja[c]['nombre']} ({c})",
                key="caja_codigo_quitar",
            )
            if st.button("➖ Quitar una unidad", key="caja_quitar_unidad"):
                carrito_caja[codigo_quitar]["cantidad"] -= 1
                if carrito_caja[codigo_quitar]["cantidad"] <= 0:
                    del carrito_caja[codigo_quitar]
                st.rerun()
        with col_vaciar:
            st.write("")
            if st.button("🗑️ Vaciar carrito", key="caja_vaciar"):
                st.session_state["caja_carrito"] = {}
                st.session_state["caja_mensaje"] = None
                st.rerun()
    else:
        st.info("Carrito vacío. Escanea un producto para comenzar.")

    st.divider()
    if st.session_state["caja_lecturas_realizadas"] == 0:
        st.subheader("Listos para iniciar prueba")
        st.info("Todavía no se han realizado lecturas. Escanea un código para comenzar.")
    else:
        st.subheader("Resultados de la prueba")
        if not st.session_state["caja_no_encontrados"] and not st.session_state["caja_encontrados_incompletos"]:
            st.success("No se registraron lecturas con problemas.")
        else:
            st.caption("Revisa las lecturas con problemas detectadas durante la prueba.")
    if st.session_state["caja_lecturas_realizadas"] > 0:
        col_no_encontrados, col_incompletos = st.columns(2)
        with col_no_encontrados:
            st.markdown("**Códigos no encontrados**")
            faltantes_caja = st.session_state["caja_no_encontrados"]
            st.caption(f"{len(faltantes_caja)} lecturas sin coincidencia en el catálogo.")
            if faltantes_caja:
                st.dataframe(pd.DataFrame({"Código no encontrado": faltantes_caja}),
                             hide_index=True, width="stretch")
            else:
                st.caption("Sin códigos desconocidos registrados.")
            if st.button("Limpiar códigos no encontrados", key="caja_limpiar_faltantes",
                         disabled=not faltantes_caja, width="stretch"):
                st.session_state["caja_no_encontrados"] = []
                st.rerun()

        with col_incompletos:
            st.markdown("**Productos encontrados con problemas**")
            incompletos_caja = st.session_state["caja_encontrados_incompletos"]
            st.caption(f"{len(incompletos_caja)} lecturas con producto encontrado, pero no agregado.")
            if incompletos_caja:
                st.dataframe(pd.DataFrame(incompletos_caja),
                             hide_index=True, width="stretch")
            else:
                st.caption("Sin productos con problemas registrados.")
            if st.button("Limpiar productos encontrados incompletos",
                         key="caja_limpiar_incompletos", disabled=not incompletos_caja,
                         width="stretch"):
                st.session_state["caja_encontrados_incompletos"] = []
                st.rerun()

    st.button("🔄 Reiniciar prueba completa", key="caja_reiniciar_prueba",
              width="stretch", on_click=_reiniciar_prueba_simulador)
    st.caption("Importante: esta prueba valida datos del catálogo local, no la aceptación de códigos por TUU Desk.")


# ==========================================
# SECCION 3: DUPLICADOS
# ==========================================

with st.expander(
    "🔎 Revisión de duplicados",
    expanded=True
):
    st.caption(
        "La detección es manual. "
        "No se eliminarán productos automáticamente."
    )

    if st.button(
        "🔍 Analizar duplicados",
        type="primary"
    ):
        with st.spinner(
            "Analizando productos..."
        ):
            grupos = detectar_grupos_duplicados(
                st.session_state["df_productos"]
            )

        st.session_state[
            "grupos_duplicados"
        ] = grupos

        st.session_state[
            "producto_eliminar"
        ] = None

    grupos = st.session_state[
        "grupos_duplicados"
    ]

    if grupos is None:
        st.info(
            "Presiona Analizar duplicados "
            "para revisar el catálogo."
        )

    elif not grupos:
        st.success(
            "No se encontraron duplicados."
        )

    else:
        total = sum(
            grupo["cantidad"]
            for grupo in grupos
        )

        st.warning(
            f"Se encontraron {len(grupos)} "
            f"grupos con {total} registros "
            "que requieren revisión."
        )

        for grupo in grupos:
            with st.container(border=True):
                st.markdown(
                    f"### Grupo {grupo['grupo']} "
                    f"— {grupo['cantidad']} productos"
                )

                st.caption(
                    "Motivo: "
                    + ", ".join(grupo["motivos"])
                )

                for producto in grupo["productos"]:
                    indice = producto["indice"]

                    st.markdown(
                        f"**{producto['nombre']}**"
                    )

                    st.caption(
                        f"Fila Excel: "
                        f"{producto['fila_excel']} | "
                        f"SKU: "
                        f"{producto['sku'] or 'Sin SKU'} | "
                        f"EAN: "
                        f"{producto['ean'] or 'Sin EAN'} | "
                        f"Precio: "
                        f"{producto['precio'] or 'Sin precio'}"
                    )

                    col_editar, col_eliminar, _ = (
                        st.columns([1, 1, 4])
                    )

                    with col_editar:
                        if st.button(
                            "✏️ Editar",
                            key=f"dup_editar_{indice}",
                            width="stretch"
                        ):
                            cambiar_vista(
                                "editar",
                                indice
                            )
                            st.rerun()

                    with col_eliminar:
                        if st.button(
                            "❌ Eliminar",
                            key=f"dup_eliminar_{indice}",
                            width="stretch"
                        ):
                            st.session_state[
                                "producto_eliminar"
                            ] = indice
                            st.rerun()

                    if (
                        st.session_state[
                            "producto_eliminar"
                        ] == indice
                    ):
                        st.error(
                            "¿Confirmas la eliminación "
                            "de este producto?"
                        )

                        st.caption(
                            "Puedes recuperarlo "
                            "con Deshacer."
                        )

                        col_si, col_no = (
                            st.columns(2)
                        )

                        with col_si:
                            if st.button(
                                "Sí, eliminar",
                                type="primary",
                                key=f"si_eliminar_{indice}"
                            ):
                                df_nuevo = (
                                    st.session_state[
                                        "df_productos"
                                    ]
                                    .drop(index=indice)
                                    .reset_index(drop=True)
                                )

                                correcto, error = (
                                    guardar_cambios(
                                        df_nuevo
                                    )
                                )

                                if correcto:
                                    st.toast(
                                        "Producto eliminado",
                                        icon="✅"
                                    )
                                    st.rerun()
                                else:
                                    st.error(error)

                        with col_no:
                            if st.button(
                                "Cancelar",
                                key=f"no_eliminar_{indice}"
                            ):
                                st.session_state[
                                    "producto_eliminar"
                                ] = None
                                st.rerun()

                    st.divider()


# ==========================================
# ETIQUETAS Y CARTILLA DE CODIGOS DE BARRAS
# ==========================================
def _filtros_impresion(df_base, prefijo):
    """Filtros de una vista sobre el catálogo único, sin modificarlo."""
    col_cat, col_marca, col_origen = st.columns(3)
    with col_cat:
        categoria = st.selectbox("Categoría", opciones_columna(df_base, "CATEGORIAS", "Todas las categorías"), key=f"{prefijo}_cat")
    with col_marca:
        marca = st.selectbox("Marca", opciones_columna(df_base, "MARCAS", "Todas las marcas"), key=f"{prefijo}_marca")
    with col_origen:
        origen = st.selectbox("Origen del código de barras", ["Todos los códigos", "Código interno generado", "Código comercial / externo"], key=f"{prefijo}_origen")
    buscar = st.text_input("Buscar producto", placeholder="Nombre, SKU o código de barras", key=f"{prefijo}_buscar")
    resultado = df_base.copy()
    if categoria != "Todas las categorías":
        resultado = resultado.loc[resultado["CATEGORIAS"].astype(str).str.strip().str.casefold().eq(categoria.casefold())]
    if marca != "Todas las marcas":
        resultado = resultado.loc[resultado["MARCAS"].astype(str).str.strip().str.casefold().eq(marca.casefold())]
    if origen != "Todos los códigos":
        try:
            resultado = filtrar_por_origen(resultado, origen, cargar_codigos())
        except RuntimeError as exc:
            st.error(f"No se pudo consultar el registro de códigos internos: {exc}")
            return resultado.iloc[0:0]
    if buscar.strip():
        mascara = pd.Series(False, index=resultado.index)
        for campo in ("NOMBRE*", "CÓDIGO SKU", "CÓDIGO DE BARRAS"):
            mascara |= resultado[campo].astype(str).str.contains(buscar.strip(), case=False, regex=False, na=False)
        resultado = resultado.loc[mascara]
    return resultado


with st.expander("🏷️ Impresión de etiquetas y cartillas", expanded=True):
    st.caption("Ambos documentos usan el mismo catálogo. Formato fijo: carta vertical, tres columnas, categoría, nombre, SKU, código escaneable, número y precio.")
    df_impresion = st.session_state["df_productos"]

    with st.expander("🏷️ Imprimir etiquetas de productos", expanded=True):
        st.caption("Selecciona productos y cambia la cantidad de etiquetas que necesitas de cada uno. Una cantidad de 0 excluye el producto.")
        df_et = _filtros_impresion(df_impresion, "imp_et")
        if df_et.empty:
            st.info("No hay productos con los filtros seleccionados.")
        else:
            tabla = df_et[["NOMBRE*", "CÓDIGO SKU", "CÓDIGO DE BARRAS", "PRECIO*", "CATEGORIAS"]].copy()
            tabla.insert(0, "Cantidad", 1)
            tabla.insert(0, "Fila", tabla.index)
            clave = hashlib.md5(str(tuple(df_et.index)).encode()).hexdigest()[:12]
            seleccion = st.data_editor(tabla, hide_index=True, width="stretch", height=320,
                disabled=["Fila", "NOMBRE*", "CÓDIGO SKU", "CÓDIGO DE BARRAS", "PRECIO*", "CATEGORIAS"],
                column_config={"Cantidad": st.column_config.NumberColumn("Etiquetas", min_value=0, max_value=500, step=1, required=True)},
                key=f"cantidad_etiquetas_{st.session_state['version_tabla']}_{clave}")
            cantidades = {}
            for _, fila in seleccion.iterrows():
                cantidad = int(fila["Cantidad"] or 0)
                if cantidad > 0:
                    cantidades[int(fila["Fila"])] = cantidad
            st.caption(f"{len(cantidades)} productos; {sum(cantidades.values())} etiquetas solicitadas.")
            if cantidades:
                seleccion_df = df_impresion.loc[list(cantidades)]
                validas, problemas = validar_etiquetas(seleccion_df)
                if problemas:
                    st.warning("Los productos siguientes no se imprimirán hasta corregir sus datos.")
                    st.dataframe(pd.DataFrame(problemas), hide_index=True, width="stretch")
                if validas:
                    try:
                        pdf, _ = generar_pdf_etiquetas(seleccion_df, titulo="Etiquetas de productos", cantidades=cantidades, modo="etiquetas")
                        st.download_button("📄 Descargar etiquetas PDF", pdf, file_name="etiquetas_TUU.pdf", mime="application/pdf", type="primary", width="stretch")
                    except (ValueError, OSError) as exc:
                        st.error(f"No se pudo generar el PDF: {exc}")

    with st.expander("📋 Generar cartilla de códigos de barras", expanded=True):
        st.caption("Cartilla agrupada por categoría y ordenada alfabéticamente. Cada producto se imprime una sola vez.")
        df_cartilla = _filtros_impresion(df_impresion, "imp_cart")
        if df_cartilla.empty:
            st.info("No hay productos con los filtros seleccionados.")
        else:
            tabla = df_cartilla[["NOMBRE*", "CÓDIGO SKU", "CÓDIGO DE BARRAS", "PRECIO*", "CATEGORIAS"]].copy()
            tabla.insert(0, "Incluir", True)
            tabla.insert(0, "Fila", tabla.index)
            clave = hashlib.md5(str(tuple(df_cartilla.index)).encode()).hexdigest()[:12]
            seleccion = st.data_editor(tabla, hide_index=True, width="stretch", height=320,
                disabled=["Fila", "NOMBRE*", "CÓDIGO SKU", "CÓDIGO DE BARRAS", "PRECIO*", "CATEGORIAS"],
                column_config={"Incluir": st.column_config.CheckboxColumn("Incluir")},
                key=f"seleccion_cartilla_{st.session_state['version_tabla']}_{clave}")
            indices = [int(i) for i in seleccion.loc[seleccion["Incluir"], "Fila"]]
            st.caption(f"{len(indices)} productos seleccionados de {len(df_cartilla)}.")
            if indices:
                seleccion_df = df_impresion.loc[indices]
                validas, problemas = validar_etiquetas(seleccion_df)
                if problemas:
                    st.warning("Los siguientes productos no se imprimirán hasta corregir sus datos.")
                    st.dataframe(pd.DataFrame(problemas), hide_index=True, width="stretch")
                if validas:
                    try:
                        pdf, _ = generar_pdf_etiquetas(seleccion_df, titulo="Cartilla de códigos de barras", modo="cartilla")
                        st.download_button("📄 Descargar cartilla PDF", pdf, file_name="cartilla_TUU.pdf", mime="application/pdf", type="primary", width="stretch")
                    except (ValueError, OSError) as exc:
                        st.error(f"No se pudo generar el PDF: {exc}")


# ==========================================
# SECCION 4: REVISION DE PRODUCTOS INCOMPLETOS
# ==========================================

with st.expander(
    "📊 Revisión de productos incompletos",
    expanded=True
):
    df_revision = st.session_state["df_productos"]
    avance = calcular_avance(df_revision)

    total = avance["total"]
    completos = avance["completos"]
    pendientes = avance["pendientes"]
    porcentaje = avance["porcentaje"]

    col_completos, col_pendientes, col_avance = st.columns(3)
    with col_completos:
        st.metric("Productos completos", completos)
    with col_pendientes:
        st.metric("Productos pendientes", pendientes)
    with col_avance:
        st.metric("Avance del catálogo", f"{porcentaje:.1f}%")

    st.progress(porcentaje / 100 if total else 0.0)
    st.caption(
        f"{completos} de {total} productos cumplen las validaciones básicas "
        "de campos obligatorios para TUU."
    )

    if total == 0:
        st.info("El catálogo está vacío. Agrega productos para comenzar.")
    elif pendientes == 0:
        st.success("✅ Todos los productos cumplen las validaciones básicas.")
        st.caption("Este indicador no garantiza por sí solo la aceptación del archivo por TUU.")
    else:
        st.warning(
            f"Hay {pendientes} productos con información pendiente. "
            "Puedes seleccionarlos y corregirlos desde esta sección."
        )

        filas_revision = []
        indices_revision = []
        for indice, producto in df_revision.iterrows():
            problemas_producto = obtener_problemas(producto)
            if problemas_producto:
                indices_revision.append(indice)
                nombre_producto = str(producto.get("NOMBRE*", "")).strip()
                filas_revision.append({
                    "Fila Excel": int(indice) + 2,
                    "Producto": nombre_producto or "(Sin nombre)",
                    "Problemas": ", ".join(problemas_producto),
                })

        tabla_revision = pd.DataFrame(filas_revision)
        clave_revision = f"revision_incompletos_{st.session_state['version_tabla']}"
        seleccion_revision = st.dataframe(
            tabla_revision,
            hide_index=True,
            width="stretch",
            height=min(420, max(150, 44 * len(tabla_revision) + 42)),
            on_select="rerun",
            selection_mode="single-row",
            key=clave_revision,
        )

        filas_seleccionadas_revision = seleccion_revision.selection.rows
        indice_revision_seleccionado = None
        if filas_seleccionadas_revision:
            posicion_seleccionada = filas_seleccionadas_revision[0]
            if 0 <= posicion_seleccionada < len(indices_revision):
                indice_revision_seleccionado = indices_revision[posicion_seleccionada]

        col_editar_revision, col_info_revision = st.columns([2, 5])
        with col_editar_revision:
            if st.button(
                "✏️ Editar seleccionado",
                key="editar_incompleto_seleccionado",
                type="primary",
                width="stretch",
                disabled=indice_revision_seleccionado is None,
            ):
                cambiar_vista("editar", indice_revision_seleccionado)
                st.rerun()
        with col_info_revision:
            if indice_revision_seleccionado is not None:
                producto_seleccionado_revision = df_revision.loc[indice_revision_seleccionado]
                st.caption(
                    f"Seleccionado: {producto_seleccionado_revision.get('NOMBRE*', '')}"
                )
            else:
                st.caption("Selecciona una fila y pulsa «Editar seleccionado» para corregirla.")


# ==========================================
# SECCION 5: DESCARGAR CATALOGO — AL FINAL
# ==========================================

with st.expander(
    "💾 Guardar como Excel",
    expanded=True
):
    ruta_trabajo = Path(st.session_state["ruta_trabajo"])

    st.subheader("📥 Descargar catálogo")
    nombre_sugerido = ruta_trabajo.stem + "_editado.xlsx"
    nombre_descarga = st.text_input(
        "Nombre del archivo",
        value=nombre_sugerido,
        key="nombre_descarga_catalogo",
    )

    if not nombre_descarga.lower().endswith(".xlsx"):
        nombre_descarga += ".xlsx"

    if ruta_trabajo.exists():
        st.download_button(
            "📥 Descargar catálogo actualizado",
            data=ruta_trabajo.read_bytes(),
            file_name=Path(nombre_descarga).name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            width="stretch",
        )
        st.caption(
            "La descarga está siempre disponible, aunque existan productos pendientes. "
            "El Excel original permanece intacto."
        )
