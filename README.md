# Criptoactivos como instrumento alternativo: formación de precios y volatilidad

- **Autora:** ASTETE ROSALES GRAYCE NICOL · **Código:** 2024200482A
- **Curso:** Finanzas I (055D) · Escuela Profesional de Economía · UNCP · 2026-II · Docente: Dr. Ciro Iván Machacuay Meza
- **Repositorio:** https://github.com/GrayceNicolAsteteRosales/FINANZASI-Criptoactivos-formacion-de-precios-y-volatilidad
- **Fecha de corte:** 2025-12-31 · **Python:** 3.13.18

Base de datos y código reproducible del artículo de la Unidad II. Se analiza la formación de precios y la
volatilidad del Bitcoin frente a activos tradicionales (S&P 500, oro), al mercado peruano (IGBVL, tipo de
cambio) y a variables financieras globales (VIX, Tesoro a 10 años, índice del dólar), con datos diarios de
2018-01-01 a 2025-12-31.

## 1. Fuentes y vías de extracción

| Vía | Fuente | Series | Script |
|---|---|---|---|
| 1 · API | Yahoo Finance (librería `yfinance`) | BTC-USD, ETH-USD, ^GSPC, GC=F | `codigo/01_extraccion_api.py` |
| 1 · API | BCRPData (API REST del BCRP) | PD38026MD (IGBVL), PD04638PD (TC interbancario venta) | `codigo/01_extraccion_api.py` |
| 1 · API | Coin Metrics Community API | btc: PriceUSD, CapMrktCurUSD | `codigo/01_extraccion_api.py` |
| 2 · Descarga programática | FRED, Fed de St. Louis (CSV oficial) | VIXCLS, DGS10, DTWEXBGS | `codigo/02_scraping_web.py` |

Ninguna fuente requiere clave (`.env.example` lo documenta). Endpoints completos en `diccionario_variables.xlsx`.

**Fecha de extracción** (hora UTC del servidor de Colab, registrada en `log_ejecucion.txt`):
- Vía 1 (API): 2026-11-21 19:18:27
- Vía 2 (FRED): 2026-11-21 19:34:21

Detalle de la vía 1:
- Yahoo BTC-USD (Bitcoin en dólares): 2922 filas, de 2018-01-01 a 2025-12-31 -> datos_crudos/csv_yahoo/BTC_USD.csv
- Yahoo ETH-USD (Ethereum en dólares): 2922 filas, de 2018-01-01 a 2025-12-31 -> datos_crudos/csv_yahoo/ETH_USD.csv
- Yahoo ^GSPC (Índice S&P 500): 2011 filas, de 2018-01-02 a 2025-12-31 -> datos_crudos/csv_yahoo/GSPC.csv
- Yahoo GC=F (Futuro del oro (COMEX)): 2013 filas, de 2018-01-02 a 2025-12-31 -> datos_crudos/csv_yahoo/GC_F.csv
- BCRP PD38026MD (Índice General Bursátil de la BVL (base 31/12/91 = 100)): HTTP 200, 2088 periodos, de 01.Ene.18 a 31.Dic.25 | nombre oficial: Indicadores bursátiles - Índice General Bursátil BVL (índice) | URL: https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PD38026MD/json/2018-01-01/2025-12-31/esp
- BCRP PD04638PD (Tipo de cambio interbancario venta (S/ por US$)): HTTP 200, 2088 periodos, de 01.Ene.18 a 31.Dic.25 | nombre oficial: Tipo de cambio - TC Interbancario (S/ por US$) - Venta | URL: https://estadisticas.bcrp.gob.pe/estadisticas/series/api/PD04638PD/json/2018-01-01/2025-12-31/esp
- Coin Metrics página 1: HTTP 200, 2922 filas, de 2018-01-01 a 2025-12-31 | URL: https://community-api.coinmetrics.io/v4/timeseries/asset-metrics?assets=btc&metrics=PriceUSD%2CCapMrktCurUSD&frequency=1d&start_time=2018-01-01&end_time=2025-12-31&page_size=10000
- Coin Metrics: 2922 filas en total

Detalle de la vía 2:
- FRED VIXCLS (Índice de volatilidad CBOE (VIX), puntos): HTTP 200, 2087 filas en el archivo, 2087 dentro de 2018-01-01 a 2025-12-31 (49 sin dato) -> datos_crudos/csv_fred/VIXCLS.csv | URL: https://fred.stlouisfed.org/graph/fredgraph.csv?id=VIXCLS&cosd=2018-01-01&coed=2025-12-31
- FRED DGS10 (Rendimiento del Tesoro de EE. UU. a 10 años, %): HTTP 200, 2087 filas en el archivo, 2087 dentro de 2018-01-01 a 2025-12-31 (88 sin dato) -> datos_crudos/csv_fred/DGS10.csv | URL: https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10&cosd=2018-01-01&coed=2025-12-31
- FRED DTWEXBGS (Índice amplio nominal del dólar (ene-2006 = 100)): HTTP 200, 2087 filas en el archivo, 2087 dentro de 2018-01-01 a 2025-12-31 (92 sin dato) -> datos_crudos/csv_fred/DTWEXBGS.csv | URL: https://fred.stlouisfed.org/graph/fredgraph.csv?id=DTWEXBGS&cosd=2018-01-01&coed=2025-12-31

**Rastreo responsable:** antes de descargar se consulta `robots.txt` de FRED (permite `/graph/fredgraph.csv`;
`Crawl-delay: 1`), se usa un User-Agent identificable y una pausa de 1,5 s entre solicitudes. Solo se usa
información pública. Las incidencias (CoinGecko, SBS) están en `incidencias_fuente.md`.

## 2. Estructura

```
codigo/            01_extraccion_api.py · 02_scraping_web.py · 03_limpieza_datos.py · 04_analisis.py
datos_crudos/      respuestas originales (csv_yahoo, json_bcrp, json_coinmetrics, csv_fred) y datos_crudos_2024200482A.csv
datos_procesados/  datos_procesados_2024200482A.csv (+ .xlsx de revisión)
salidas/           tablas (.xlsx .csv .tex), figuras (.png .pdf), dashboard, reporte de calidad, hash
diccionario_variables.xlsx · requirements.txt · .env.example · log_ejecucion.txt · incidencias_fuente.md
```

## 3. Base procesada

- **Archivo:** `datos_procesados/datos_procesados_2024200482A.csv` (separador `;`, coma decimal, UTF-8 con BOM).
- **Dimensiones:** 2922 filas × 28 columnas (id, fecha y 26 variables).
- **Llave común:** `fecha`. **Calendario:** días con precio del Bitcoin (todos los días del periodo).
- **Faltantes:** fines de semana y feriados de los mercados tradicionales quedan vacíos; **no se imputa ningún valor**.
  Cada variable tiene entre 1990 y 2922 observaciones con dato.
- **Valores mal publicados en la fuente:** 1 caso(s) detectado(s) por regla explícita y documentado(s)
  en `salidas/correcciones_fuente.csv`; el crudo conserva el valor original.

### Huella digital (SHA-256)

```
c96517f8181ccec431d586c6df64e860b8d1bb4c39d06ac0d4a174e70777111c  datos_procesados_2024200482A.csv
```

Verificación: Windows PowerShell `Get-FileHash datos_procesados\datos_procesados_2024200482A.csv -Algorithm SHA256`;
Linux/Mac `sha256sum datos_procesados/datos_procesados_2024200482A.csv`.

## 4. Cómo reproducir

```bash
pip install -r requirements.txt
python codigo/01_extraccion_api.py      # vía 1: API
python codigo/02_scraping_web.py        # vía 2: descarga programática
python codigo/03_limpieza_datos.py      # unión, limpieza, variables derivadas, controles y hash
python codigo/04_analisis.py            # tablas, figuras y dashboard
```

**Versión del lenguaje y de las librerías:** pandas==2.2.3 · numpy==2.4.6 · requests==2.32.4 · yfinance==0.2.66 · statsmodels==0.15.0 · arch==9.0.0 · scipy==1.16.3 · matplotlib==3.10.0 · plotly==5.24.1 · openpyxl==3.1.5
(lista completa en `requirements.txt`).

Los cuatro scripts se ejecutan desde la carpeta del proyecto, usan rutas relativas y parámetros fijos
(`FECHA_INICIO = "2018-01-01"`, `FECHA_CORTE = "2025-12-31"`). No hay procesos aleatorios.
Si se vuelven a descargar los datos, las fuentes podrían haber revisado valores históricos; por eso los
crudos originales se conservan en `datos_crudos/`.

## 5. Resultados (generados por `04_analisis.py`)

- Tablas: tabla1_descriptivos, tabla2_adf, tabla3_correlaciones, tabla3b_matriz_correlaciones, tabla4_mco_newey_west, tabla5_vif, tabla6_garch, tabla7_granger, tabla8_resumen_anual, tablas_articulo y el libro `tablas_articulo.xlsx`.
- Figuras: figura1_precio_capitalizacion.png, figura2_desempeno_acumulado.png, figura3_volatilidad_realizada.png, figura4_correlacion_movil.png, figura5_volatilidad_condicional.png, figura6_distribucion_retornos.png, figura7_curva_impacto_noticias.png, figura8_matrices_correlacion.png.
- Tablero interactivo: `salidas/dashboard_criptoactivos.html` (se abre en cualquier navegador, sin internet).
- Cifras citadas en el artículo: `salidas/resultados_clave.txt`.

## 6. Referencias de las fuentes (APA 7)

- Banco Central de Reserva del Perú. (2026). *BCRPData: Series estadísticas* [Base de datos]. Recuperado el 21 de noviembre de 2026, de https://estadisticas.bcrp.gob.pe/estadisticas/series/
- Coin Metrics. (2026). *Coin Metrics Community API* (versión 4) [Base de datos]. Recuperado el 21 de noviembre de 2026, de https://community-api.coinmetrics.io/v4
- Federal Reserve Bank of St. Louis. (2026). *FRED, Federal Reserve Economic Data* [Base de datos]. Recuperado el 21 de noviembre de 2026, de https://fred.stlouisfed.org/
- Yahoo. (2026). *Yahoo Finance* [Base de datos]. Recuperado el 21 de noviembre de 2026, de https://finance.yahoo.com/

Uso académico. Los datos pertenecen a sus respectivas fuentes.

## Cómo reproducir los resultados

1. Clonar o descargar este repositorio y abrir una terminal en la carpeta del proyecto.
2. Instalar las librerías: `pip install -r requirements.txt`
3. Copiar `.env.example` como `.env` (ninguna fuente usada requiere clave).
4. Ejecutar en orden:
   - `python codigo/01_extraccion_api.py` (Yahoo Finance, BCRPData y Coin Metrics, vía API)
   - `python codigo/02_scraping_web.py` (FRED, descarga programática con revisión de robots.txt)
   - `python codigo/03_limpieza_datos.py` (integración, control de calidad y hash SHA-256)
   - `python codigo/04_analisis.py` (tablas, figuras y dashboard en `/salidas`)
5. Comparar el hash de `salidas/hash_sha256.txt` con el indicado en este README.

Nota: una nueva extracción puede diferir si las fuentes revisan sus datos históricos.
La base entregada corresponde a la extracción registrada en `log_ejecucion.txt`.

## Historial de versiones

| Fecha | Cambio |
|---|---|
| 2026-11-21 | Extracción de datos (vías API y descarga programática), limpieza y análisis |
| 2026-11-22 | Documentación: README, diccionario de variables, requirements, .env.example e incidencias de fuente |
| 2026-11-25 | Verificación de datos contra las fuentes originales e instrucciones de reproducción |
