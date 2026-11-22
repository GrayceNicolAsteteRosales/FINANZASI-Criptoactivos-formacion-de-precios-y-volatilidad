# Incidencias de fuentes

| Fecha | Fuente | Incidencia | Decisión |
|---|---|---|---|
| 2026-11-21 | CoinGecko API (`/coins/bitcoin/market_chart/range`), sugerida en el temario | Sin clave, la consulta de 2018 devolvió HTTP 401: el plan gratuito solo permite 365 días de historia | Se reemplazó por Coin Metrics Community API (sin clave, historia completa) para precio de referencia y capitalización de mercado |
| 2026-11-21 | SBS, página `TipoCambioPromedio.aspx` | La página respondió HTTP 503 (*Service Unavailable*) en dos intentos | El tipo de cambio se tomó de BCRPData (serie PD04638PD, API); la segunda vía de extracción es la descarga programática de los CSV oficiales de FRED |
| 2026-11-21 | BCRPData, serie PD38026MD (IGBVL) | El 21/06/2024 la fuente publica 22 521,03, entre 29 857,57 (20/06) y 30 058,66 (24/06): salto aislado de −28,2 % y +28,9 % | `03_limpieza_datos.py` lo detecta con una regla explícita (pico aislado), lo deja vacío en la base procesada y lo registra en `salidas/correcciones_fuente.csv`; el dato crudo se conserva intacto |
