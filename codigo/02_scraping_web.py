# Nombres y apellidos: ASTETE ROSALES GRAYCE NICOL
# Código de matrícula: 2024200482A
# Tema N.º 95 del temario: Criptoactivos como instrumento alternativo: formación de precios y volatilidad
# Fecha de extracción: 2026-11-21
"""
02_scraping_web.py — Vía 2 (descarga programática, numeral 2.4.3 de la consigna).

El script solicita a FRED (Banco de la Reserva Federal de St. Louis) el archivo CSV oficial
de cada serie, lo guarda intacto en /datos_crudos/csv_fred y lo lee con código para
comprobar que trae la ventana pedida. Series:
  VIXCLS   : índice de volatilidad implícita del S&P 500 (VIX), puntos.
  DGS10    : rendimiento del bono del Tesoro de EE. UU. a 10 años, porcentaje.
  DTWEXBGS : índice amplio nominal del dólar estadounidense (enero 2006 = 100).

Buenas prácticas de rastreo: se revisa robots.txt antes de descargar, se declara un
User-Agent identificable y se deja una pausa de 1,5 s entre solicitudes.

Entradas: ninguna.   Salidas: datos_crudos/csv_fred/<serie>.csv y log_ejecucion.txt.
Uso:      desde la carpeta del proyecto ->  python codigo/02_scraping_web.py
"""

# ---------------------------------------------------------------------------
# 1. Librerías
# ---------------------------------------------------------------------------
import io
import logging
import sys
import time
from pathlib import Path
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# 2. Parámetros congelados de la consulta
# ---------------------------------------------------------------------------
CODIGO_MATRICULA = "2024200482A"
FECHA_INICIO = "2018-01-01"
FECHA_CORTE = "2025-12-31"

SERIES_FRED = {
    "VIXCLS": "Índice de volatilidad CBOE (VIX), puntos",
    "DGS10": "Rendimiento del Tesoro de EE. UU. a 10 años, %",
    "DTWEXBGS": "Índice amplio nominal del dólar (ene-2006 = 100)",
}
# cosd y coed fijan la fecha inicial y final del archivo que entrega FRED
URL_FRED = "https://fred.stlouisfed.org/graph/fredgraph.csv?id={serie}&cosd={inicio}&coed={fin}"
URL_ROBOTS = "https://fred.stlouisfed.org/robots.txt"

CABECERAS = {"User-Agent": "Investigacion academica UNCP Finanzas I - e_2024200482A@uncp.edu.pe"}
PAUSA = 1.5

# ---------------------------------------------------------------------------
# 3. Rutas relativas a la carpeta del proyecto
# ---------------------------------------------------------------------------
RAIZ = Path(__file__).resolve().parent.parent
DIR_FRED = RAIZ / "datos_crudos" / "csv_fred"
DIR_FRED.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 4. Registro (log)
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | 02_scraping_web | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(RAIZ / "log_ejecucion.txt", encoding="utf-8"),
        logging.StreamHandler(),
    ],
    force=True,
)
log = logging.getLogger()


# ---------------------------------------------------------------------------
# 5. Revisión de robots.txt
# ---------------------------------------------------------------------------
def permiso_robots(url_prueba):
    """Descarga robots.txt con nuestro User-Agent y pregunta si la URL está permitida."""
    respuesta = requests.get(URL_ROBOTS, headers=CABECERAS, timeout=60)
    lector = RobotFileParser()
    lector.parse(respuesta.text.splitlines())
    permitido = lector.can_fetch(CABECERAS["User-Agent"], url_prueba)
    log.info("robots.txt de FRED: HTTP %s | ¿se permite %s? %s",
             respuesta.status_code, url_prueba.split("?")[0], "SÍ" if permitido else "NO")
    return permitido


# ---------------------------------------------------------------------------
# 6. Descarga y lectura de cada CSV
# ---------------------------------------------------------------------------
def descargar_serie(serie, descripcion):
    """Pide el CSV oficial, lo guarda sin cambios y lo lee para verificar su contenido."""
    url = URL_FRED.format(serie=serie, inicio=FECHA_INICIO, fin=FECHA_CORTE)
    respuesta = requests.get(url, headers=CABECERAS, timeout=60)
    if respuesta.status_code != 200:
        log.error("FRED %s: HTTP %s | URL: %s", serie, respuesta.status_code, url)
        return
    # Se guardan los bytes exactos que entregó el servidor (evidencia original)
    ruta = DIR_FRED / f"{serie}.csv"
    ruta.write_bytes(respuesta.content)

    # Lectura con código: FRED marca los días sin dato con "." o con celda vacía
    tabla = pd.read_csv(io.StringIO(respuesta.text), na_values=["."])
    tabla.columns = ["fecha", "valor"]
    tabla["fecha"] = pd.to_datetime(tabla["fecha"])
    en_ventana = tabla[(tabla["fecha"] >= FECHA_INICIO) & (tabla["fecha"] <= FECHA_CORTE)]
    log.info(
        "FRED %s (%s): HTTP %s, %d filas en el archivo, %d dentro de %s a %s (%d sin dato) -> %s | URL: %s",
        serie, descripcion, respuesta.status_code, len(tabla), len(en_ventana),
        FECHA_INICIO, FECHA_CORTE, en_ventana["valor"].isna().sum(), ruta.relative_to(RAIZ), url,
    )


# ---------------------------------------------------------------------------
# 7. Ejecución
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    log.info("=== Inicio de la descarga programática (FRED) | ventana %s a %s ===", FECHA_INICIO, FECHA_CORTE)
    url_ejemplo = URL_FRED.format(serie="VIXCLS", inicio=FECHA_INICIO, fin=FECHA_CORTE)
    if not permiso_robots(url_ejemplo):
        log.error("robots.txt no permite la descarga: se detiene el script sin descargar nada")
        sys.exit(1)
    time.sleep(PAUSA)
    for serie, descripcion in SERIES_FRED.items():
        descargar_serie(serie, descripcion)
        time.sleep(PAUSA)
    log.info("=== Fin de la descarga programática ===")
