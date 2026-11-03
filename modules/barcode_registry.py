
"""Registro local de códigos internos emitidos por Gestor TUU.

No altera las 11 columnas oficiales de TUU. Solo se consideran internos los
códigos efectivamente registrados por esta aplicación.
"""
import json
import os
from pathlib import Path


def ruta_registro():
    base = os.environ.get('LOCALAPPDATA')
    carpeta = (Path(base) if base else Path.home()) / 'GestorTUU'
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta / 'barcode_registry.json'


def cargar_codigos():
    ruta = ruta_registro()
    if not ruta.exists():
        return set()
    try:
        with ruta.open(encoding='utf-8') as f:
            datos = json.load(f)
        if not isinstance(datos, dict) or not isinstance(datos.get('codigos'), list):
            raise ValueError('Formato de registro inválido')
        return {str(x) for x in datos['codigos'] if str(x).strip()}
    except (OSError, ValueError, TypeError) as exc:
        raise RuntimeError(f'No se pudo leer el registro de códigos internos: {exc}') from exc


def guardar_codigos(codigos):
    ruta = ruta_registro()
    temporal = ruta.with_suffix('.tmp')
    try:
        with temporal.open('w', encoding='utf-8') as f:
            json.dump({'version': 1, 'codigos': sorted(set(codigos))}, f, indent=2)
        temporal.replace(ruta)
        return True, None
    except OSError as exc:
        return False, str(exc)


def origen_codigo(codigo, internos):
    codigo = str(codigo).strip()
    if not codigo:
        return 'Sin código de barras'
    return 'Código interno generado' if codigo in internos else 'Código comercial / externo'


def filtrar_por_origen(df, opcion, internos):
    if opcion == 'Todos los códigos':
        return df
    mascara = df['CÓDIGO DE BARRAS'].fillna('').astype(str).map(
        lambda x: origen_codigo(x, internos) == opcion
    )
    return df.loc[mascara]
