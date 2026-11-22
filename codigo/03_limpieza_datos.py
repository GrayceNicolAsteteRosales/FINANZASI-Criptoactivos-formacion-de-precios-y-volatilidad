# Nombres y apellidos: ASTETE ROSALES GRAYCE NICOL
# Código de matrícula: 2024200482A
# Tema N.º 95 del temario: Criptoactivos como instrumento alternativo: formación de precios y volatilidad
# Fecha de extracción: 2026-11-21
"""
03_limpieza_datos.py — Depuración, unión por la llave común (fecha) y variables derivadas.

Entradas (generadas por 01 y 02, nunca editadas a mano):
  datos_crudos/csv_yahoo/*.csv · datos_crudos/json_bcrp/*.json
  datos_crudos/json_coinmetrics/pagina_*.json · datos_crudos/csv_fred/*.csv
Salidas:
  datos_crudos/datos_crudos_2024200482A.csv (+ .xlsx)      unión de las fuentes, valores tal como vienen
  datos_procesados/datos_procesados_2024200482A.csv (+ .xlsx)  base limpia del artículo
  salidas/reporte_calidad.csv · salidas/revision_extremos.csv · salidas/hash_sha256.txt
Reglas: calendario diario del Bitcoin (cotiza todos los días); los días sin mercado de los
activos tradicionales (fines de semana y feriados) quedan vacíos: NO se imputa ningún valor.
Uso:    desde la carpeta del proyecto ->  python codigo/03_limpieza_datos.py
"""

# ---------------------------------------------------------------------------
# 1. Librerías
# ---------------------------------------------------------------------------
import hashlib
import json
import logging
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# 2. Parámetros
# ---------------------------------------------------------------------------
CODIGO_MATRICULA = "2024200482A"
FECHA_INICIO = "2018-01-01"
FECHA_CORTE = "2025-12-31"
DECIMALES = 6                 # redondeo antes de guardar y calcular el hash
VENTANA_BTC = 30              # días calendario para la volatilidad del Bitcoin
VENTANA_BURSATIL = 21         # días hábiles (~30 días calendario) para S&P 500 y BVL
UMBRAL_EXTREMO = 15.0         # |retorno diario| en % que se manda a revisar contra la fuente
TOLERANCIA_REGRESO = 2.0      # % máximo entre el día anterior y el siguiente para llamar "aislado" a un pico
SERIES_REVISADAS = ["btc_precio_usd", "eth_precio_usd", "sp500_indice", "oro_usd_onza",
                    "igbvl_indice", "tc_pen_usd", "dolar_indice"]

RAIZ = Path(__file__).resolve().parent.parent
CRUDOS = RAIZ / "datos_crudos"
PROCESADOS = RAIZ / "datos_procesados"
SALIDAS = RAIZ / "salidas"
for carpeta in (PROCESADOS, SALIDAS):
    carpeta.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | 03_limpieza_datos | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.FileHandler(RAIZ / "log_ejecucion.txt", encoding="utf-8"), logging.StreamHandler()],
    force=True,
)
log = logging.getLogger()


# ---------------------------------------------------------------------------
# 3. Lectura de cada fuente (una función por formato)
# ---------------------------------------------------------------------------
def leer_yahoo(archivo, prefijo):
    """CSV de yfinance: la fecha viene con zona horaria ('2018-01-02 00:00:00-05:00'); se toma el día."""
    tabla = pd.read_csv(CRUDOS / "csv_yahoo" / archivo)
    tabla["fecha"] = pd.to_datetime(tabla["Date"].astype(str).str[:10])
    tabla = tabla.rename(columns={"Close": f"{prefijo}_close", "Adj Close": f"{prefijo}_adj_close",
                                  "Volume": f"{prefijo}_volume"})
    return tabla[["fecha", f"{prefijo}_close", f"{prefijo}_adj_close", f"{prefijo}_volume"]]


MESES = {"Ene": 1, "Feb": 2, "Mar": 3, "Abr": 4, "May": 5, "Jun": 6, "Jul": 7,
         "Ago": 8, "Set": 9, "Sep": 9, "Oct": 10, "Nov": 11, "Dic": 12}


def fecha_bcrp(texto):
    """Convierte '04.Ene.21' en la fecha 2021-01-04."""
    dia, mes, anio = texto.split(".")
    return pd.Timestamp(year=2000 + int(anio), month=MESES[mes], day=int(dia))


def leer_bcrp(codigo, columna):
    """JSON del BCRP: periods[].name es la fecha y periods[].values[0] el valor ('n.d.' = sin dato)."""
    contenido = json.loads((CRUDOS / "json_bcrp" / f"{codigo}.json").read_text(encoding="utf-8"))
    filas = [(fecha_bcrp(p["name"]), p["values"][0]) for p in contenido["periods"]]
    tabla = pd.DataFrame(filas, columns=["fecha", columna])
    tabla[columna] = pd.to_numeric(tabla[columna].replace("n.d.", np.nan), errors="coerce")
    return tabla


def leer_coinmetrics():
    """Une todas las páginas JSON de Coin Metrics; 'time' viene como '2018-01-01T00:00:00.000000000Z'."""
    filas = []
    for pagina in sorted((CRUDOS / "json_coinmetrics").glob("pagina_*.json")):
        filas += json.loads(pagina.read_text(encoding="utf-8"))["data"]
    tabla = pd.DataFrame(filas)
    tabla["fecha"] = pd.to_datetime(tabla["time"].str[:10])
    tabla["cm_btc_priceusd"] = pd.to_numeric(tabla["PriceUSD"], errors="coerce")
    tabla["cm_btc_capmrktcurusd"] = pd.to_numeric(tabla["CapMrktCurUSD"], errors="coerce")
    return tabla[["fecha", "cm_btc_priceusd", "cm_btc_capmrktcurusd"]]


def leer_fred(serie, columna):
    """CSV de FRED: columnas observation_date y la serie; '.' o celda vacía = sin dato."""
    tabla = pd.read_csv(CRUDOS / "csv_fred" / f"{serie}.csv", na_values=["."])
    tabla.columns = ["fecha", columna]
    tabla["fecha"] = pd.to_datetime(tabla["fecha"])
    tabla[columna] = pd.to_numeric(tabla[columna], errors="coerce")
    return tabla


# ---------------------------------------------------------------------------
# 4. Funciones auxiliares
# ---------------------------------------------------------------------------
def guardar_csv(tabla, ruta):
    """CSV para Excel en español: separador ';', coma decimal y BOM UTF-8."""
    tabla.to_csv(ruta, sep=";", decimal=",", index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")


def retorno_log(serie):
    """Retorno logarítmico en % entre observaciones consecutivas CON dato (sin rellenar huecos)."""
    disponible = serie.dropna()
    return (100 * np.log(disponible / disponible.shift(1))).reindex(serie.index)


def diferencia(serie, factor=1.0):
    """Diferencia entre observaciones consecutivas con dato (multiplicada por un factor)."""
    disponible = serie.dropna()
    return (factor * disponible.diff()).reindex(serie.index)


def anular_picos_aislados(p, columna):
    """
    Detecta valores mal publicados en la fuente: el nivel salta más de UMBRAL_EXTREMO % un día
    y al día siguiente regresa (en sentido contrario, también más de UMBRAL_EXTREMO %) a menos de
    TOLERANCIA_REGRESO % del nivel previo. Ese valor se deja VACÍO (no se imputa) y se registra.
    """
    s = p[columna].dropna()
    r = 100 * np.log(s / s.shift(1))                       # salto de ida
    r_siguiente = r.shift(-1)                              # salto de vuelta
    regreso = (100 * np.log(s.shift(-1) / s.shift(1))).abs()
    marca = ((r.abs() > UMBRAL_EXTREMO) & (r_siguiente.abs() > UMBRAL_EXTREMO)
             & (np.sign(r) != np.sign(r_siguiente)) & (regreso < TOLERANCIA_REGRESO))
    registros = []
    for fecha in s.index[marca]:
        posicion = s.index.get_loc(fecha)
        registros.append({
            "fecha": fecha.date(), "variable": columna, "valor_en_fuente": s.iloc[posicion],
            "valor_dia_anterior": s.iloc[posicion - 1], "valor_dia_siguiente": s.iloc[posicion + 1],
            "salto_ida_pct": round(r.iloc[posicion], 4), "salto_vuelta_pct": round(r_siguiente.iloc[posicion], 4),
            "tratamiento": "se deja vacío en datos_procesados; el crudo conserva el valor original",
        })
        p.loc[fecha, columna] = np.nan
        log.warning("Pico aislado en %s el %s: valor en la fuente %s (anterior %s, siguiente %s) -> vacío",
                    columna, fecha.date(), s.iloc[posicion], s.iloc[posicion - 1], s.iloc[posicion + 1])
    return registros


def volatilidad(retornos, ventana, dias_anio):
    """Desviación estándar móvil de los retornos con dato, anualizada, en %."""
    disponible = retornos.dropna()
    return (disponible.rolling(ventana, min_periods=ventana).std() * np.sqrt(dias_anio)).reindex(retornos.index)


# ---------------------------------------------------------------------------
# 5. Base cruda unificada (los valores tal como vienen de cada fuente)
# ---------------------------------------------------------------------------
def construir_crudos():
    fuentes = [
        leer_yahoo("BTC_USD.csv", "yahoo_btc_usd"),
        leer_yahoo("ETH_USD.csv", "yahoo_eth_usd"),
        leer_yahoo("GSPC.csv", "yahoo_gspc"),
        leer_yahoo("GC_F.csv", "yahoo_gc_f_usd"),
        leer_coinmetrics(),
        leer_bcrp("PD38026MD", "bcrp_pd38026md_igbvl"),
        leer_bcrp("PD04638PD", "bcrp_pd04638pd_tc_pen_usd"),
        leer_fred("VIXCLS", "fred_vixcls"),
        leer_fred("DGS10", "fred_dgs10_pct"),
        leer_fred("DTWEXBGS", "fred_dtwexbgs"),
    ]
    for tabla in fuentes:   # control: ninguna fuente debe traer fechas repetidas
        repetidas = tabla["fecha"].duplicated().sum()
        if repetidas:
            raise ValueError(f"Fechas repetidas en {tabla.columns[1]}: {repetidas}")
    crudos = fuentes[0]
    for tabla in fuentes[1:]:
        crudos = crudos.merge(tabla, on="fecha", how="outer")   # unión por la llave común
    crudos = crudos[(crudos["fecha"] >= FECHA_INICIO) & (crudos["fecha"] <= FECHA_CORTE)]
    crudos = crudos.sort_values("fecha").reset_index(drop=True)
    crudos.insert(0, "id", range(1, len(crudos) + 1))
    log.info("Base cruda unificada: %d filas x %d columnas (%s a %s)", len(crudos), crudos.shape[1],
             crudos["fecha"].min().date(), crudos["fecha"].max().date())
    return crudos


# ---------------------------------------------------------------------------
# 6. Base procesada
# ---------------------------------------------------------------------------
def construir_procesados(crudos):
    # 6.1 Calendario: los días en que existe precio del Bitcoin en Yahoo
    base = crudos[crudos["yahoo_btc_usd_close"].notna()].copy().set_index("fecha")
    log.info("Calendario del Bitcoin: %d días (se descartan %d fechas sin precio de BTC)",
             len(base), len(crudos) - len(base))

    # 6.2 Variables en niveles, con nombres normalizados y unidad en el nombre
    p = pd.DataFrame(index=base.index)
    p["btc_precio_usd"] = base["yahoo_btc_usd_close"]
    p["btc_volumen_musd"] = base["yahoo_btc_usd_volume"] / 1e6
    p["btc_capitalizacion_mmusd"] = base["cm_btc_capmrktcurusd"] / 1e9
    p["btc_precio_cm_usd"] = base["cm_btc_priceusd"]
    p["eth_precio_usd"] = base["yahoo_eth_usd_close"]
    p["sp500_indice"] = base["yahoo_gspc_close"]
    p["oro_usd_onza"] = base["yahoo_gc_f_usd_close"]
    p["igbvl_indice"] = base["bcrp_pd38026md_igbvl"]
    p["tc_pen_usd"] = base["bcrp_pd04638pd_tc_pen_usd"]
    p["vix_puntos"] = base["fred_vixcls"]
    p["ust10_pct"] = base["fred_dgs10_pct"]
    p["dolar_indice"] = base["fred_dtwexbgs"]
    p["dia_habil_eeuu"] = p["sp500_indice"].notna().astype(int)

    # 6.3 Valores mal publicados en la fuente (picos aislados): se dejan vacíos y se documentan
    correcciones = []
    for columna in SERIES_REVISADAS:
        correcciones += anular_picos_aislados(p, columna)
    columnas_corr = ["fecha", "variable", "valor_en_fuente", "valor_dia_anterior", "valor_dia_siguiente",
                     "salto_ida_pct", "salto_vuelta_pct", "tratamiento"]
    guardar_csv(pd.DataFrame(correcciones, columns=columnas_corr), SALIDAS / "correcciones_fuente.csv")
    log.info("Picos aislados dejados vacíos: %d (ver salidas/correcciones_fuente.csv)", len(correcciones))

    # 6.4 Variables derivadas (retornos y variaciones entre observaciones con dato)
    p["ret_btc_pct"] = retorno_log(p["btc_precio_usd"])
    p["ret_eth_pct"] = retorno_log(p["eth_precio_usd"])
    p["ret_sp500_pct"] = retorno_log(p["sp500_indice"])
    p["ret_oro_pct"] = retorno_log(p["oro_usd_onza"])
    p["ret_igbvl_pct"] = retorno_log(p["igbvl_indice"])
    p["var_tc_pct"] = retorno_log(p["tc_pen_usd"])
    p["ret_dolar_pct"] = retorno_log(p["dolar_indice"])
    p["var_vix_puntos"] = diferencia(p["vix_puntos"])
    p["var_ust10_pbs"] = diferencia(p["ust10_pct"], factor=100)
    p["log_volumen_btc"] = np.log(p["btc_volumen_musd"].where(p["btc_volumen_musd"] > 0))
    p["vol30_btc_anual_pct"] = volatilidad(p["ret_btc_pct"], VENTANA_BTC, 365)
    p["vol21_sp500_anual_pct"] = volatilidad(p["ret_sp500_pct"], VENTANA_BURSATIL, 252)
    p["vol21_igbvl_anual_pct"] = volatilidad(p["ret_igbvl_pct"], VENTANA_BURSATIL, 252)

    p = p.reset_index()
    p.insert(0, "id", range(1, len(p) + 1))
    return p


# ---------------------------------------------------------------------------
# 7. Controles de calidad
# ---------------------------------------------------------------------------
def controles(p):
    # 7.1 Resumen por variable
    numericas = [c for c in p.columns if c not in ("id", "fecha")]
    filas = []
    for c in numericas:
        s = p.set_index("fecha")[c]
        con_dato = s.dropna()
        filas.append({
            "variable": c, "observaciones": int(con_dato.size), "faltantes": int(s.isna().sum()),
            "minimo": con_dato.min(), "maximo": con_dato.max(), "media": con_dato.mean(),
            "fecha_minimo": con_dato.idxmin().date() if con_dato.size else None,
            "fecha_maximo": con_dato.idxmax().date() if con_dato.size else None,
        })
    calidad = pd.DataFrame(filas).round(DECIMALES)
    guardar_csv(calidad, SALIDAS / "reporte_calidad.csv")
    log.info("Reporte de calidad guardado (%d variables)", len(calidad))

    # 7.2 Cotejo interno: precio de Yahoo frente al de Coin Metrics (dos fuentes distintas)
    dif = 100 * (p["btc_precio_usd"] / p["btc_precio_cm_usd"] - 1)
    log.info("Diferencia Yahoo vs Coin Metrics en el precio del BTC: mediana %.3f %%, máx. absoluta %.3f %%",
             dif.median(), dif.abs().max())

    # 7.3 Retornos extremos: no se borran; se listan para cotejarlos contra la fuente
    columnas_ret = [c for c in p.columns if c.startswith(("ret_", "var_tc"))]
    extremos = []
    for c in columnas_ret:
        marca = p[c].abs() > UMBRAL_EXTREMO
        for _, fila in p[marca].iterrows():
            extremos.append({"fecha": fila["fecha"].date(), "variable": c, "valor_pct": round(fila[c], 4)})
    extremos = pd.DataFrame(extremos, columns=["fecha", "variable", "valor_pct"])
    guardar_csv(extremos, SALIDAS / "revision_extremos.csv")
    log.info("Retornos con |valor| > %.0f %%: %d (ver salidas/revision_extremos.csv)", UMBRAL_EXTREMO, len(extremos))

    # 7.4 Observaciones conjuntas (días con todas las variables del modelo)
    conjuntas = p[["ret_btc_pct", "ret_sp500_pct", "ret_oro_pct", "ret_igbvl_pct", "var_tc_pct",
                   "var_vix_puntos", "var_ust10_pbs", "ret_dolar_pct"]].dropna()
    log.info("Días con todas las variables a la vez: %d", len(conjuntas))


# ---------------------------------------------------------------------------
# 8. Ejecución
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    log.info("=== Inicio de la limpieza ===")
    crudos = construir_crudos()
    ruta_crudos = CRUDOS / f"datos_crudos_{CODIGO_MATRICULA}.csv"
    guardar_csv(crudos, ruta_crudos)
    crudos.to_excel(ruta_crudos.with_suffix(".xlsx"), index=False)

    procesados = construir_procesados(crudos)
    numericas = procesados.select_dtypes("number").columns
    procesados[numericas] = procesados[numericas].round(DECIMALES)
    ruta_proc = PROCESADOS / f"datos_procesados_{CODIGO_MATRICULA}.csv"
    guardar_csv(procesados, ruta_proc)
    procesados.to_excel(ruta_proc.with_suffix(".xlsx"), index=False)
    log.info("Base procesada: %d filas x %d columnas -> %s", len(procesados), procesados.shape[1],
             ruta_proc.relative_to(RAIZ))

    controles(procesados)

    # Huella digital del archivo procesado (se coteja con el README)
    huella = hashlib.sha256(ruta_proc.read_bytes()).hexdigest()
    (SALIDAS / "hash_sha256.txt").write_text(f"{ruta_proc.name}  SHA-256: {huella}\n", encoding="utf-8")
    log.info("Hash SHA-256 de %s: %s", ruta_proc.name, huella)
    log.info("=== Fin de la limpieza ===")
