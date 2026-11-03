
import os
import sys
import time
import socket
import threading
import webbrowser
from pathlib import Path


def obtener_directorio_aplicacion():
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)

    return Path(__file__).resolve().parent


def esperar_y_abrir_navegador(puerto):
    direccion = f"http://127.0.0.1:{puerto}"

    for _ in range(60):
        try:
            with socket.create_connection(
                ("127.0.0.1", puerto),
                timeout=1
            ):
                webbrowser.open(direccion)
                return
        except OSError:
            time.sleep(0.5)


def main():
    if getattr(sys, "frozen", False):
        import multiprocessing
        multiprocessing.freeze_support()

    # Desactivar el modo de desarrollo de Streamlit
    # antes de cargar sus módulos.
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    from streamlit.web import cli as stcli

    carpeta_app = obtener_directorio_aplicacion()
    archivo_app = carpeta_app / "app.py"

    if not archivo_app.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo: {archivo_app}"
        )

    puerto = 8501

    threading.Thread(
        target=esperar_y_abrir_navegador,
        args=(puerto,),
        daemon=True
    ).start()

    sys.argv = [
        "streamlit",
        "run",
        str(archivo_app),
        "--global.developmentMode=false",
        "--server.address=127.0.0.1",
        f"--server.port={puerto}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--server.fileWatcherType=none",
        "--server.enableCORS=true",
        "--server.enableXsrfProtection=true"
    ]

    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
