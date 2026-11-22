# Nombres y apellidos: ASTETE ROSALES GRAYCE NICOL
# Código de matrícula: 2024200482A
# Tema N.º 95 del temario: Criptoactivos como instrumento alternativo: formación de precios y volatilidad
# Fecha de extracción: 2026-11-22
"""
01_extraccion_api.py — Vía 1 (API).

Descarga, con fechas fijas, las series diarias 2018-2025 de tres API:
  1) Yahoo Finance (librería yfinance): Bitcoin, Ethereum, S&P 500 y oro.
  2) BCRPData (API REST del BCRP): índice general de la BVL y tipo de cambio interbancario.
  3) Coin Metrics Community API: precio de referencia y capitalización de mercado del Bitcoin.

Entradas: ninguna (todo sale de internet).
Salidas:  respuestas originales, sin tocar, en /datos_crudos:
            csv_yahoo/<ticker>.csv · json_bcrp/<codigo>.json · json_coinmetrics/pagina_XX.json
          y el registro de cada consulta en log_ejecucion.txt.
Uso:      desde la carpeta del proyecto ->  python codigo/01_extraccion_api.py
"""

# ---------------------------------------------------------------------------
# 1. Librerías
# ---------------------------------------------------------------------------
import json
import logging
import time
from pathlib import Path

import pandas as pd
import requests
import yfinance as yf

# ---------------------------------------------------------------------------
# 2. Parámetros congelados de la consulta (constantes, nunca "hoy")
# ---------------------------------------------------------------------------
CODIGO_MATRICULA = "2024200482A"
FECHA_INICIO = "2018-01-01"
FECHA_CORTE = "2025-12-31"

# Yahoo Finance: ticker -> descripción
TICKERS_YAHOO = {
    "BTC-USD": "Bitcoin en dólares",
    "ETH-USD": "Ethereum en dólares",
    "^GSPC": "Índice S&P 500",
    "GC=F": "Futuro del oro (COMEX)",
}

# BCRPData: código de serie -> descripción
SERIES_BCRP = {
    "PD38026MD": "Índice General Bursátil de la BVL (base 31/12/91 = 100)",
    "PD04638PD": "Tipo de cambio interbancario venta (S/ por US$)",
}
URL_BCRP = "https://estadisticas.bcrp.gob.pe/estadisticas/series/api/{codigo}/json/{inicio}/{fin}/esp"

# Coin Metrics Community API (no requiere clave)
URL_COINMETRICS = "https://community-api.coinmetrics.io/v4/timeseries/asset-metrics"
PARAMETROS_COINMETRICS = {
    "assets": "btc",
    "metrics": "PriceUSD,CapMrktCurUSD",
    "frequency": "1d",
    "start_time": FECHA_INICIO,
    "end_time": FECHA_CORTE,
    "page_size": 10000,
}

# Identificación ante los servidores y pausa entre consultas (en segundos)
CABECERAS = {"User-Agent": "Investigacion academica UNCP Finanzas I - e_2024200482A@uncp.edu.pe"}
PAUSA = 1.5

# ---------------------------------------------------------------------------
# 3. Rutas relativas a la carpeta del proyecto
# ---------------------------------------------------------------------------
RAIZ = Path(__file__).resolve().parent.parent
DIR_YAHOO = RAIZ / "datos_crudos" / "csv_yahoo"
DIR_BCRP = RAIZ / "datos_crudos" / "json_bcrp"
DIR_COINMETRICS = RAIZ / "datos_crudos" / "json_coinmetrics"
for carpeta in (DIR_YAHOO, DIR_BCRP, DIR_COINMETRICS):
    carpeta.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 4. Registro (log): se escribe en log_ejecucion.txt y en pantalla
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | 01_extraccion_api | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(RAIZ / "log_ejecucion.txt", encoding="utf-8"),
        logging.StreamHandler(),
    ],
    force=True,
)
log = logging.getLogger()


# ---------------------------------------------------------------------------
# 5. Yahoo Finance
# ---------------------------------------------------------------------------
def nombre_archivo(ticker):
    """Convierte el ticker en un nombre de archivo válido (^GSPC -> GSPC, GC=F -> GC_F)."""
    return ticker.replace("^", "").replace("=", "_").replace("-", "_")


def extraer_yahoo():
    """Descarga cada ticker y guarda la tabla tal como la entrega yfinance."""
    # yfinance trata la fecha final como excluida: se pide un día después del corte
    fin_exclusivo = (pd.Timestamp(FECHA_CORTE) + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    for ticker, descripcion in TICKERS_YAHOO.items():
        datos = yf.Ticker(ticker).history(
            start=FECHA_INICIO, end=fin_exclusivo, interval="1d", auto_adjust=False, actions=False
        )
        if datos.empty:
            log.error("Yahoo %s (%s): la consulta no devolvió filas", ticker, descripcion)
            continue
        ruta = DIR_YAHOO / f"{nombre_archivo(ticker)}.csv"
        datos.to_csv(ruta, encoding="utf-8")
        log.info(
            "Yahoo %s (%s): %d filas, de %s a %s -> %s",
            ticker, descripcion, len(datos),
            datos.index.min().date(), datos.index.max().date(), ruta.relative_to(RAIZ),
        )
        time.sleep(PAUSA)


# ---------------------------------------------------------------------------
# 6. BCRPData
# ---------------------------------------------------------------------------
def extraer_bcrp():
    """Consulta una serie por llamada y guarda el JSON exacto que devuelve el BCRP."""
    for codigo, descripcion in SERIES_BCRP.items():
        url = URL_BCRP.format(codigo=codigo, inicio=FECHA_INICIO, fin=FECHA_CORTE)
        respuesta = requests.get(url, headers=CABECERAS, timeout=60)
        try:
            contenido = respuesta.json()
        except json.JSONDecodeError:
            # Si el código no existe, el BCRP responde con una página HTML en vez de JSON
            log.error("BCRP %s: la respuesta no es JSON (HTTP %s). URL: %s", codigo, respuesta.status_code, url)
            continue
        ruta = DIR_BCRP / f"{codigo}.json"
        ruta.write_text(respuesta.text, encoding="utf-8")
        periodos = contenido.get("periods", [])
        nombre_oficial = contenido["config"]["series"][0]["name"]
        log.info(
            "BCRP %s (%s): HTTP %s, %d periodos, de %s a %s | nombre oficial: %s | URL: %s",
            codigo, descripcion, respuesta.status_code, len(periodos),
            periodos[0]["name"] if periodos else "-", periodos[-1]["name"] if periodos else "-",
            nombre_oficial, url,
        )
        time.sleep(PAUSA)


# ---------------------------------------------------------------------------
# 7. Coin Metrics Community
# ---------------------------------------------------------------------------
def extraer_coinmetrics():
    """Recorre todas las páginas de la respuesta y guarda cada una sin modificarla."""
    url, parametros, pagina, total = URL_COINMETRICS, dict(PARAMETROS_COINMETRICS), 1, 0
    while url:
        respuesta = requests.get(url, params=parametros, headers=CABECERAS, timeout=60)
        if respuesta.status_code != 200:
            log.error("Coin Metrics página %d: HTTP %s | %s", pagina, respuesta.status_code, respuesta.text[:200])
            break
        contenido = respuesta.json()
        ruta = DIR_COINMETRICS / f"pagina_{pagina:02d}.json"
        ruta.write_text(respuesta.text, encoding="utf-8")
        filas = contenido.get("data", [])
        total += len(filas)
        log.info(
            "Coin Metrics página %d: HTTP %s, %d filas, de %s a %s | URL: %s",
            pagina, respuesta.status_code, len(filas),
            filas[0]["time"][:10] if filas else "-", filas[-1]["time"][:10] if filas else "-",
            respuesta.url,
        )
        # La API indica si hay otra página; esa URL ya trae todos los parámetros
        url, parametros, pagina = contenido.get("next_page_url"), None, pagina + 1
        time.sleep(PAUSA)
    log.info("Coin Metrics: %d filas en total", total)


# ---------------------------------------------------------------------------
# 8. Ejecución
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    log.info("=== Inicio de la extracción por API | ventana %s a %s ===", FECHA_INICIO, FECHA_CORTE)
    extraer_yahoo()
    extraer_bcrp()
    extraer_coinmetrics()
    log.info("=== Fin de la extracción por API ===")
