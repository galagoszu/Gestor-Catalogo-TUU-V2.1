"""Generador de EAN-13 internos para uso exclusivo del establecimiento."""
import secrets


def digito_control_ean13(doce_digitos):
    texto = str(doce_digitos)
    if len(texto) != 12 or not texto.isascii() or not texto.isdigit():
        raise ValueError('Se requieren exactamente 12 dígitos ASCII.')
    total = sum(int(n) * (1 if i % 2 == 0 else 3) for i, n in enumerate(texto))
    return str((-total) % 10)


def validar_ean13(codigo):
    texto = str(codigo).strip()
    return (len(texto) == 13 and texto.isascii() and texto.isdigit()
            and digito_control_ean13(texto[:12]) == texto[-1])


def generar_ean13_interno(codigos_existentes=()):
    """Genera un EAN-13 con prefijo 20, comprobando colisiones locales.

    No constituye un GTIN comercial registrado en GS1.
    """
    ocupados = {str(c).strip() for c in codigos_existentes if str(c).strip()}
    for _ in range(1000):
        base = '20' + f'{secrets.randbelow(10**10):010d}'
        codigo = base + digito_control_ean13(base)
        if codigo not in ocupados:
            return codigo
    raise RuntimeError('No se pudo generar un código libre; vuelve a intentarlo.')
