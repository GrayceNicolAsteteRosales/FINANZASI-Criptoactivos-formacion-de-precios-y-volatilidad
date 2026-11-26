# Nombres y apellidos: ASTETE ROSALES GRAYCE NICOL
# Código de matrícula: 2024200482A
# Tema N.º 95 del temario: Criptoactivos como instrumento alternativo: formación de precios y volatilidad
# Fecha de extracción: 2026-11-21
"""
04_analisis.py — Estimaciones, tablas, figuras y dashboard del artículo.

Entrada:  datos_procesados/datos_procesados_2024200482A.csv (generado por 03)
Salidas (en /salidas):
  Tablas: cada una en .xlsx con formato profesional, .csv (';' y coma decimal) y .tex (booktabs),
          más el libro consolidado tablas_articulo.xlsx (una hoja por tabla y un índice).
    tabla1_descriptivos · tabla2_adf · tabla3_correlaciones · tabla3b_matriz_correlaciones
    tabla4_mco_newey_west · tabla5_vif · tabla6_garch · tabla7_granger · tabla8_resumen_anual
    tabla9_chow · tabla10_portafolio
  Figuras (.png a 300 dpi y .pdf):
    figura1_precio_capitalizacion · figura2_desempeno_acumulado · figura3_volatilidad_realizada
    figura4_correlacion_movil · figura5_volatilidad_condicional · figura6_distribucion_retornos
    figura7_curva_impacto_noticias · figura8_matrices_correlacion · figura9_portafolio
  dashboard_criptoactivos.html (tablero interactivo; se abre en cualquier navegador, sin internet)
  resultados_clave.txt (cifras que se citan en el texto del artículo)
Uso: desde la carpeta del proyecto ->  python codigo/04_analisis.py
No hay procesos aleatorios: los resultados son idénticos en cada ejecución.
"""

# ---------------------------------------------------------------------------
# 1. Librerías
# ---------------------------------------------------------------------------
import html
import logging
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")                     # genera figuras sin abrir ventanas
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import statsmodels.api as sm
from arch import arch_model
from matplotlib.transforms import blended_transform_factory
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from plotly.subplots import make_subplots
from scipy import stats
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.tsa.stattools import adfuller, grangercausalitytests

warnings.filterwarnings("ignore")         # avisos de convergencia y de versiones, no de datos

# ---------------------------------------------------------------------------
# 2. Parámetros
# ---------------------------------------------------------------------------
CODIGO_MATRICULA = "2024200482A"
AUTORA = "Grayce Nicol Astete Rosales"
FECHA_INICIO = "2018-01-01"
FECHA_CORTE = "2025-12-31"
CORTE_SUBPERIODO = "2021-12-31"           # 2018-2021 frente a 2022-2025
REZAGOS_NEWEY_WEST = 5
REZAGOS_GRANGER = 5
VENTANA_CORRELACION = 63                  # días hábiles (~ 3 meses calendario)
ENLACE_REPOSITORIO = "https://github.com/GrayceNicolAsteteRosales/FINANZASI-Criptoactivos-formacion-de-precios-y-volatilidad"

# Hitos del mercado cripto que se marcan en las figuras (fechas de conocimiento público)
EVENTOS = [
    ("2020-03-12", "Jueves negro\n(COVID-19)"),
    ("2021-11-10", "Máximo\nde 2021"),
    ("2022-05-09", "Colapso\nTerra-LUNA"),
    ("2022-11-11", "Quiebra\nde FTX"),
    ("2024-01-10", "ETF spot de\nBTC (SEC)"),
    ("2024-04-20", "4.º halving"),
]

FUENTE_GENERAL = ("Fuente: elaboración propia con datos de Yahoo Finance, Coin Metrics, "
                  "BCRPData (BCRP) y FRED (Fed de St. Louis), 2018-2025.")

RAIZ = Path(__file__).resolve().parent.parent
ARCHIVO = RAIZ / "datos_procesados" / f"datos_procesados_{CODIGO_MATRICULA}.csv"
SALIDAS = RAIZ / "salidas"
SALIDAS.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | 04_analisis | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.FileHandler(RAIZ / "log_ejecucion.txt", encoding="utf-8"), logging.StreamHandler()],
    force=True,
)
log = logging.getLogger()

# Colores fijos por activo (el mismo activo lleva siempre el mismo color)
AZUL, NARANJA, VERDE_AGUA, AMARILLO = "#2a78d6", "#eb6834", "#1baf7a", "#eda100"
GRIS, GRIS_CLARO, TINTA, TINTA_2 = "#8a8a85", "#e3e2dd", "#1f1f1d", "#5f5e5a"
COLOR = {"Bitcoin": AZUL, "S&P 500": NARANJA, "Oro": VERDE_AGUA, "IGBVL": AMARILLO}
AZUL_TABLA, FILA_ALTERNA = "1F3864", "EEF3FA"

# Estilo de revista: tipografía serif (similar a Times, como el texto LaTeX del artículo),
# ancho exacto de la columna de texto, ejes finos y marcas hacia afuera.
ANCHO = 6.3                               # pulgadas = ancho de texto de una página A4 con márgenes de 2,5 cm
LINEAS = {"Bitcoin": "-", "S&P 500": (0, (5, 1.5)), "Oro": (0, (1.2, 1.2)), "IGBVL": (0, (6, 1.5, 1.2, 1.5))}
plt.rcParams.update({
    "font.family": "serif", "font.serif": ["STIXGeneral", "DejaVu Serif"], "mathtext.fontset": "stix",
    "font.size": 8.5, "axes.titlesize": 9, "axes.titleweight": "bold", "axes.labelsize": 8.5,
    "xtick.labelsize": 8, "ytick.labelsize": 8, "legend.fontsize": 7.5,
    "axes.titlecolor": TINTA, "axes.labelcolor": TINTA, "xtick.color": TINTA_2, "ytick.color": TINTA_2,
    "axes.spines.top": False, "axes.spines.right": False, "axes.edgecolor": TINTA_2, "axes.linewidth": 0.6,
    "xtick.direction": "out", "ytick.direction": "out", "xtick.major.size": 3, "ytick.major.size": 3,
    "xtick.major.width": 0.6, "ytick.major.width": 0.6,
    "axes.grid": True, "axes.grid.axis": "y", "grid.color": "#ebeae6", "grid.linewidth": 0.5,
    "axes.axisbelow": True, "legend.frameon": False, "lines.linewidth": 0.9,
    "savefig.bbox": "tight", "savefig.pad_inches": 0.04, "pdf.fonttype": 42,
})

CLAVE = []                                # líneas para resultados_clave.txt
TABLAS = []                               # (nombre, título, tabla, nota) para el libro consolidado


# ---------------------------------------------------------------------------
# 3. Tablas: formato Excel profesional, CSV y LaTeX
# ---------------------------------------------------------------------------
def escapar_tex(texto):
    """Escapa los caracteres especiales de LaTeX."""
    texto = str(texto)
    for a, b in [("\\", r"\textbackslash{}"), ("%", r"\%"), ("_", r"\_"), ("&", r"\&"),
                 ("#", r"\#"), ("$", r"\$"), ("→", r"$\rightarrow$"), ("²", r"$^2$")]:
        texto = texto.replace(a, b)
    return texto


def formato_tex(valor, decimales):
    """Número con decimales fijos y espacio fino como separador de miles."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return ""
    if isinstance(valor, (int, np.integer)):
        return f"{valor:,}".replace(",", r"\,")
    if isinstance(valor, (float, np.floating)):
        return f"{valor:,.{decimales}f}".replace(",", r"\,")
    return escapar_tex(valor)


def formato_numero_excel(columna, serie, decimales):
    """Elige el formato de número de Excel según la columna."""
    if pd.api.types.is_integer_dtype(serie):
        return "#,##0"
    if columna.lower().startswith("p ") or columna.lower() == "p":
        return "0.0000"
    if serie.abs().max() >= 1000:
        return "#,##0.00"
    return "0." + "0" * decimales


def escribir_hoja(hoja, tabla, titulo, nota, decimales):
    """Escribe una tabla con título, encabezado azul, filas alternas, formatos de número y nota."""
    n_col = tabla.shape[1]
    fino, grueso = Side(style="thin", color="BFBFBF"), Side(style="medium", color=AZUL_TABLA)
    hoja.sheet_view.showGridLines = False

    # Título y subtítulo
    hoja.cell(row=1, column=1, value=titulo).font = Font(bold=True, size=13, color=AZUL_TABLA)
    hoja.cell(row=2, column=1, value=f"{AUTORA} · Finanzas I (UNCP) · Tema 95 · {FECHA_INICIO} a {FECHA_CORTE}"
              ).font = Font(italic=True, size=9, color="595959")

    # Encabezados
    for j, columna in enumerate(tabla.columns, start=1):
        celda = hoja.cell(row=4, column=j, value=str(columna))
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=AZUL_TABLA)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = Border(top=grueso, bottom=grueso)
    hoja.row_dimensions[4].height = 34

    # Cuerpo
    formatos = {c: formato_numero_excel(c, tabla[c], decimales)
                for c in tabla.columns if pd.api.types.is_numeric_dtype(tabla[c])}
    for i, fila in enumerate(tabla.itertuples(index=False), start=5):
        relleno = PatternFill("solid", fgColor=FILA_ALTERNA) if (i - 5) % 2 == 1 else None
        for j, (columna, valor) in enumerate(zip(tabla.columns, fila), start=1):
            if isinstance(valor, (float, np.floating)):
                valor = None if np.isnan(valor) else float(valor)
            elif isinstance(valor, np.integer):
                valor = int(valor)
            celda = hoja.cell(row=i, column=j, value=valor)
            celda.alignment = Alignment(horizontal="left" if j == 1 else "center", vertical="center")
            celda.border = Border(bottom=fino)
            if j == 1:
                celda.font = Font(bold=True, color="262626")
            if columna in formatos:
                celda.number_format = formatos[columna]
            if relleno:
                celda.fill = relleno
    ultima = len(tabla) + 4
    for j in range(1, n_col + 1):
        hoja.cell(row=ultima, column=j).border = Border(bottom=grueso)

    # Nota al pie (celdas combinadas, texto ajustado)
    fila_nota = ultima + 2
    hoja.merge_cells(start_row=fila_nota, start_column=1, end_row=fila_nota, end_column=max(n_col, 4))
    celda = hoja.cell(row=fila_nota, column=1, value=f"Nota. {nota}")
    celda.font = Font(italic=True, size=9, color="404040")
    celda.alignment = Alignment(wrap_text=True, vertical="top")
    hoja.row_dimensions[fila_nota].height = 15 * (1 + len(nota) // 120)

    # Ancho de columnas según el contenido
    for j, columna in enumerate(tabla.columns, start=1):
        largo = max([len(str(columna)) * 0.8] + [len(f"{v:,.2f}") if isinstance(v, (float, np.floating))
                                                  else len(str(v)) for v in tabla[columna]])
        hoja.column_dimensions[get_column_letter(j)].width = min(max(12, largo + 3), 42)
    hoja.freeze_panes = "B5"
    hoja.page_setup.orientation = "landscape"
    hoja.sheet_properties.pageSetUpPr.fitToPage = True   # al imprimir o exportar, cabe en el ancho de la hoja
    hoja.page_setup.fitToWidth = 1
    hoja.page_setup.fitToHeight = 0


def guardar_tabla(tabla, nombre, titulo, nota, decimales=3):
    """Guarda la tabla en .xlsx (formato profesional), .csv (redondeado) y .tex (booktabs)."""
    redondeada = tabla.copy()
    numericas = redondeada.select_dtypes("number").columns
    redondeada[numericas] = redondeada[numericas].round(decimales)

    # CSV para Excel en español (abrirlo con Excel, no con la vista previa de Drive)
    redondeada.to_csv(SALIDAS / f"{nombre}.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")

    # Excel con formato
    libro = Workbook()
    escribir_hoja(libro.active, redondeada, titulo, nota, decimales)
    libro.active.title = nombre[:31]
    libro.save(SALIDAS / f"{nombre}.xlsx")
    TABLAS.append((nombre, titulo, redondeada, nota, decimales))

    # LaTeX: tabla completa lista para \input{} en Overleaf
    alineacion = "l" + "c" * (tabla.shape[1] - 1)
    lineas = [r"\begin{table}[H]", r"\centering", r"\footnotesize",
              rf"\caption{{{escapar_tex(titulo)}}}", rf"\label{{tab:{nombre}}}",
              r"\resizebox{\textwidth}{!}{%", rf"\begin{{tabular}}{{{alineacion}}}", r"\toprule",
              " & ".join(rf"\textbf{{{escapar_tex(c)}}}" for c in tabla.columns) + r" \\", r"\midrule"]
    for fila in redondeada.itertuples(index=False):
        lineas.append(" & ".join(formato_tex(v, decimales) for v in fila) + r" \\")
    lineas += [r"\bottomrule", r"\end{tabular}}",
               rf"\par\vspace{{3pt}}\parbox{{\textwidth}}{{\scriptsize \textit{{Nota.}} {escapar_tex(nota)}}}",
               r"\end{table}"]
    (SALIDAS / f"{nombre}.tex").write_text("\n".join(lineas), encoding="utf-8")
    log.info("Tabla guardada: %s (%d filas)", nombre, len(tabla))


def guardar_libro_consolidado():
    """Un solo Excel con un índice y una hoja por tabla (útil para revisar y para el anexo)."""
    libro = Workbook()
    indice = libro.active
    indice.title = "Indice"
    indice.sheet_view.showGridLines = False
    indice["A1"] = "Tablas del artículo: Criptoactivos como instrumento alternativo"
    indice["A1"].font = Font(bold=True, size=14, color=AZUL_TABLA)
    indice["A2"] = f"{AUTORA} · Código {CODIGO_MATRICULA} · Generado por codigo/04_analisis.py"
    indice["A2"].font = Font(italic=True, size=9, color="595959")
    for j, texto in enumerate(["Hoja", "Título"], start=1):
        celda = indice.cell(row=4, column=j, value=texto)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=AZUL_TABLA)
    for i, (nombre, titulo, tabla, nota, decimales) in enumerate(TABLAS, start=5):
        hoja = libro.create_sheet(nombre[:31])
        escribir_hoja(hoja, tabla, titulo, nota, decimales)
        enlace = indice.cell(row=i, column=1, value=nombre)
        enlace.hyperlink = f"#'{nombre[:31]}'!A1"
        enlace.font = Font(color="0563C1", underline="single")
        indice.cell(row=i, column=2, value=titulo)
    indice.column_dimensions["A"].width = 32
    indice.column_dimensions["B"].width = 95
    libro.save(SALIDAS / "tablas_articulo.xlsx")
    log.info("Libro consolidado guardado: tablas_articulo.xlsx (%d tablas)", len(TABLAS))


def estrellas(p):
    """Asteriscos de significancia: *** 1 %, ** 5 %, * 10 %."""
    return "***" if p < 0.01 else "**" if p < 0.05 else "*" if p < 0.10 else ""


# ---------------------------------------------------------------------------
# 4. Figuras: funciones de apoyo
# ---------------------------------------------------------------------------
NOTA_EVENTOS = ("Hitos: " + "; ".join(f"({k}) {texto.replace(chr(10), ' ')}, {pd.Timestamp(f):%d/%m/%Y}"
                                       for k, (f, texto) in enumerate(EVENTOS, start=1)) + ".")


def guardar_figura(figura, nombre, fuente=FUENTE_GENERAL):
    """Agrega la nota de fuente al pie y guarda en .png (300 dpi) y .pdf vectorial."""
    figura.text(0.0, -0.005, fuente, fontsize=7, color=TINTA_2, ha="left", va="top", style="italic", wrap=True,
                transform=figura.transFigure)
    figura.savefig(SALIDAS / f"{nombre}.png", dpi=300, facecolor="white")
    figura.savefig(SALIDAS / f"{nombre}.pdf", facecolor="white")
    plt.close(figura)
    log.info("Figura guardada: %s", nombre)


def eje_fechas(eje):
    """Marcas anuales en el eje x."""
    eje.xaxis.set_major_locator(mdates.YearLocator())
    eje.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    eje.xaxis.set_minor_locator(mdates.MonthLocator(bymonth=(4, 7, 10)))
    eje.tick_params(axis="x", which="minor", length=1.5, width=0.4)
    eje.set_xlim(pd.Timestamp(FECHA_INICIO), pd.Timestamp(FECHA_CORTE) + pd.Timedelta(days=15))


def marcar_subperiodos(eje, etiquetar=False):
    """Sombrea suavemente 2022-2025 y rotula los dos subperiodos del análisis."""
    inicio2 = pd.Timestamp(CORTE_SUBPERIODO) + pd.Timedelta(days=1)
    eje.axvspan(inicio2, pd.Timestamp(FECHA_CORTE) + pd.Timedelta(days=15), color="#f5f4f0", zorder=0, lw=0)
    eje.axvline(inicio2, color=TINTA_2, linewidth=0.6, zorder=1)
    if etiquetar:
        mezcla = blended_transform_factory(eje.transData, eje.transAxes)
        for x, texto in [(pd.Timestamp("2020-01-01"), "2018-2021"), (pd.Timestamp("2024-01-01"), "2022-2025")]:
            eje.text(x, 0.02, texto, transform=mezcla, ha="center", fontsize=7, color=TINTA_2, style="italic")


def marcar_eventos(eje, rotulos=True):
    """Línea punteada en cada hito con su número en un círculo (la clave va en la nota de la figura)."""
    mezcla = blended_transform_factory(eje.transData, eje.transAxes)
    for k, (fecha, _) in enumerate(EVENTOS, start=1):
        x = pd.Timestamp(fecha)
        eje.axvline(x, color=GRIS, linewidth=0.5, linestyle=(0, (1.5, 1.5)), zorder=1)
        if rotulos:
            eje.text(x, 1.0, str(k), transform=mezcla, fontsize=6.5, color=TINTA, ha="center", va="center",
                     bbox=dict(boxstyle="circle,pad=0.18", fc="white", ec=TINTA_2, lw=0.5), zorder=6, clip_on=False)


def formato_log(eje):
    """Eje logarítmico con marcas 1-2-5, números completos y sin rótulos menores."""
    eje.yaxis.set_major_locator(mticker.LogLocator(base=10, subs=(1, 2, 5)))
    eje.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    eje.yaxis.set_minor_locator(mticker.NullLocator())


def miles(eje):
    eje.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))


def titulo_panel(eje, texto):
    eje.set_title(texto, loc="left", pad=9)


def rotulos_finales(eje, tabla, formato="{:,.0f}", log_y=False, separacion=0.07):
    """Rótulos al final de cada serie, separados para que no se encimen."""
    finales = tabla.ffill().iloc[-1].sort_values()
    y = np.log10(finales.values) if log_y else finales.values.astype(float)
    rango = (np.nanmax(np.log10(tabla.values)) - np.nanmin(np.log10(tabla.values))) if log_y else \
        (np.nanmax(tabla.values) - np.nanmin(tabla.values))
    for i in range(1, len(y)):
        y[i] = max(y[i], y[i - 1] + separacion * rango)
    x_rotulo = tabla.index[-1] + pd.Timedelta(days=40)
    for (nombre, valor), yr in zip(finales.items(), (10 ** y) if log_y else y):
        eje.annotate(f"{nombre} " + formato.format(valor), xy=(tabla.index[-1], valor), xytext=(x_rotulo, yr),
                     fontsize=7.5, color=TINTA, va="center", annotation_clip=False,
                     arrowprops=dict(arrowstyle="-", color=COLOR.get(nombre, GRIS), lw=0.6))
# ---------------------------------------------------------------------------
# 5. Carga de la base procesada y muestra conjunta de días hábiles
# ---------------------------------------------------------------------------
d = pd.read_csv(ARCHIVO, sep=";", decimal=",", encoding="utf-8-sig", parse_dates=["fecha"]).set_index("fecha")
log.info("Base procesada leída: %d filas x %d columnas", *d.shape)

# Días en que existen TODOS los niveles: los retornos se recalculan entre esos días,
# para que el retorno del BTC cubra el mismo intervalo que el de los activos bursátiles.
NIVELES = ["btc_precio_usd", "sp500_indice", "oro_usd_onza", "igbvl_indice", "tc_pen_usd",
           "vix_puntos", "ust10_pct", "dolar_indice"]
h = d[NIVELES].dropna()
conjunta = pd.DataFrame({
    "BTC": 100 * np.log(h["btc_precio_usd"]).diff(),
    "S&P 500": 100 * np.log(h["sp500_indice"]).diff(),
    "Oro": 100 * np.log(h["oro_usd_onza"]).diff(),
    "IGBVL": 100 * np.log(h["igbvl_indice"]).diff(),
    "TC PEN/USD": 100 * np.log(h["tc_pen_usd"]).diff(),
    "Var. VIX": h["vix_puntos"].diff(),
    "Var. UST10 (pbs)": 100 * h["ust10_pct"].diff(),
    "Índice dólar": 100 * np.log(h["dolar_indice"]).diff(),
}).dropna()
log.info("Muestra conjunta de días hábiles: %d observaciones", len(conjunta))
CLAVE.append(f"Muestra conjunta (días con todas las variables): {len(conjunta)} observaciones, "
             f"{conjunta.index.min().date()} a {conjunta.index.max().date()}")
inicio2 = pd.Timestamp(CORTE_SUBPERIODO) + pd.Timedelta(days=1)
periodos = {"2018-2025": conjunta, "2018-2021": conjunta.loc[:CORTE_SUBPERIODO], "2022-2025": conjunta.loc[inicio2:]}

# ---------------------------------------------------------------------------
# 6. Tabla 1: estadísticos descriptivos de los retornos (hechos estilizados)
# ---------------------------------------------------------------------------
SERIES_RET = [("ret_btc_pct", "Bitcoin", 365), ("ret_eth_pct", "Ethereum", 365),
              ("ret_sp500_pct", "S&P 500", 252), ("ret_oro_pct", "Oro", 252),
              ("ret_igbvl_pct", "IGBVL (BVL)", 252), ("var_tc_pct", "Tipo de cambio PEN/USD", 252),
              ("ret_dolar_pct", "Índice del dólar", 252)]
filas = []
for columna, nombre, dias in SERIES_RET:
    s = d[columna].dropna()
    jb = stats.jarque_bera(s)
    filas.append({"Serie (retorno diario, %)": nombre, "N": len(s), "Media": s.mean(),
                  "Desv. est.": s.std(), "Mínimo": s.min(), "Máximo": s.max(),
                  "VaR 5 % histórico": s.quantile(0.05), "Días con |r| > 5 %": int((s.abs() > 5).sum()),
                  "Asimetría": stats.skew(s), "Curtosis exceso": stats.kurtosis(s),
                  "Jarque-Bera": jb.statistic, "p (JB)": jb.pvalue,
                  "Vol. anual (%)": s.std() * np.sqrt(dias)})
t1 = pd.DataFrame(filas)
guardar_tabla(t1, "tabla1_descriptivos", "Tabla 1. Estadísticos descriptivos de los retornos logarítmicos diarios",
              "Retornos logarítmicos ×100 entre observaciones consecutivas con dato. VaR 5 %: percentil 5 de la "
              "distribución empírica. Volatilidad anualizada con √365 para criptoactivos (cotizan todos los días) y "
              "√252 para los demás. " + FUENTE_GENERAL)
for _, f in t1.iterrows():
    CLAVE.append(f"{f.iloc[0]}: vol. anual {f['Vol. anual (%)']:.1f} %, curtosis exceso {f['Curtosis exceso']:.2f}, "
                 f"asimetría {f['Asimetría']:.2f}, mínimo {f['Mínimo']:.2f} %, VaR 5 % {f['VaR 5 % histórico']:.2f} %")

# ---------------------------------------------------------------------------
# 7. Tabla 2: prueba de raíz unitaria ADF (niveles y primeras diferencias)
# ---------------------------------------------------------------------------
SERIES_ADF = [("btc_precio_usd", "ln Bitcoin", True), ("sp500_indice", "ln S&P 500", True),
              ("oro_usd_onza", "ln Oro", True), ("igbvl_indice", "ln IGBVL", True),
              ("tc_pen_usd", "ln TC PEN/USD", True), ("vix_puntos", "VIX", False),
              ("ust10_pct", "UST10 (%)", False), ("log_volumen_btc", "ln Volumen BTC", False)]
filas = []
for columna, nombre, usar_log in SERIES_ADF:
    s = d[columna].dropna()
    s = np.log(s) if usar_log else s
    for etiqueta, serie in [("Nivel", s), ("Primera diferencia", s.diff().dropna())]:
        estad, p, rezagos, nobs, _, _ = adfuller(serie, autolag="AIC", regression="c")
        filas.append({"Variable": nombre, "Transformación": etiqueta, "Estadístico ADF": estad,
                      "p": p, "Rezagos (AIC)": int(rezagos), "N": int(nobs),
                      "Conclusión (5 %)": "Estacionaria" if p < 0.05 else "Raíz unitaria"})
t2 = pd.DataFrame(filas)
guardar_tabla(t2, "tabla2_adf", "Tabla 2. Prueba de Dickey-Fuller aumentada",
              "H0: la serie tiene raíz unitaria. Regresión con constante; rezagos elegidos por AIC. " + FUENTE_GENERAL)

# ---------------------------------------------------------------------------
# 8. Tabla 3: correlaciones de Pearson en la muestra conjunta (total y subperiodos)
# ---------------------------------------------------------------------------
filas, correlaciones = [], {}
for variable in conjunta.columns.drop("BTC"):
    fila = {"Correlación del BTC con": variable}
    for etiqueta, datos in periodos.items():
        r, p = stats.pearsonr(datos["BTC"], datos[variable])
        fila[etiqueta] = f"{r:.3f}{estrellas(p)}"
        correlaciones[(variable, etiqueta)] = r
    filas.append(fila)
t3 = pd.DataFrame(filas)
guardar_tabla(t3, "tabla3_correlaciones", "Tabla 3. Correlación de los retornos diarios del Bitcoin con activos tradicionales",
              f"Muestra conjunta de días hábiles (N = {len(conjunta)}). *** p < 0,01; ** p < 0,05; * p < 0,10. "
              + FUENTE_GENERAL)
matriz = conjunta.corr().reset_index().rename(columns={"index": "Variable"})
guardar_tabla(matriz, "tabla3b_matriz_correlaciones", "Tabla 3b. Matriz de correlaciones de los retornos (2018-2025)",
              f"Muestra conjunta de días hábiles (N = {len(conjunta)}). " + FUENTE_GENERAL)

# ---------------------------------------------------------------------------
# 9. Tabla 4: MCO del retorno del BTC con errores Newey-West (total y subperiodos)
# ---------------------------------------------------------------------------
REGRESORES = ["S&P 500", "Oro", "IGBVL", "TC PEN/USD", "Var. VIX", "Var. UST10 (pbs)", "Índice dólar"]
resultados = {}
for etiqueta, datos in periodos.items():
    X = sm.add_constant(datos[REGRESORES])
    resultados[etiqueta] = sm.OLS(datos["BTC"], X).fit(cov_type="HAC", cov_kwds={"maxlags": REZAGOS_NEWEY_WEST})
filas = []
for variable in ["const"] + REGRESORES:
    coef = {"Variable": "Constante" if variable == "const" else variable}
    error = {"Variable": ""}
    for etiqueta, res in resultados.items():
        coef[etiqueta] = f"{res.params[variable]:.4f}{estrellas(res.pvalues[variable])}"
        error[etiqueta] = f"({res.bse[variable]:.4f})"
    filas += [coef, error]
for nombre_fila, valor in [("R² ajustado", lambda r: f"{r.rsquared_adj:.4f}"),
                           ("Estadístico F (p)", lambda r: f"{r.fvalue:.2f} ({r.f_pvalue:.3f})"),
                           ("N", lambda r: f"{int(r.nobs)}")]:
    filas.append({"Variable": nombre_fila, **{e: valor(r) for e, r in resultados.items()}})
t4 = pd.DataFrame(filas)
guardar_tabla(t4, "tabla4_mco_newey_west", "Tabla 4. Determinantes del retorno diario del Bitcoin (MCO)",
              f"Variable dependiente: retorno logarítmico diario del BTC (%). Errores estándar Newey-West con "
              f"{REZAGOS_NEWEY_WEST} rezagos entre paréntesis. *** p < 0,01; ** p < 0,05; * p < 0,10. " + FUENTE_GENERAL)
r_total = resultados["2018-2025"]
CLAVE.append(f"MCO 2018-2025: beta S&P 500 = {r_total.params['S&P 500']:.3f} (p = {r_total.pvalues['S&P 500']:.4f}), "
             f"beta oro = {r_total.params['Oro']:.3f} (p = {r_total.pvalues['Oro']:.4f}), R² aj. = {r_total.rsquared_adj:.3f}")
for e in ["2018-2021", "2022-2025"]:
    CLAVE.append(f"MCO {e}: beta S&P 500 = {resultados[e].params['S&P 500']:.3f} "
                 f"(p = {resultados[e].pvalues['S&P 500']:.4f}), R² aj. = {resultados[e].rsquared_adj:.3f}")

# Tabla 5: factor de inflación de varianza (multicolinealidad; criterio VIF < 10)
X = sm.add_constant(conjunta[REGRESORES]).values
t5 = pd.DataFrame({"Regresor": REGRESORES,
                   "VIF": [variance_inflation_factor(X, i + 1) for i in range(len(REGRESORES))]})
t5["Diagnóstico"] = np.where(t5["VIF"] < 10, "Sin multicolinealidad grave", "Revisar")
guardar_tabla(t5, "tabla5_vif", "Tabla 5. Factor de inflación de la varianza de los regresores",
              "Criterio: VIF < 10 indica ausencia de multicolinealidad grave. " + FUENTE_GENERAL)

# ---------------------------------------------------------------------------
# 10. Tabla 6: GARCH(1,1) y GJR-GARCH(1,1) con distribución t de Student
# ---------------------------------------------------------------------------
SERIES_GARCH = [("ret_btc_pct", "BTC", 365), ("ret_sp500_pct", "S&P 500", 252), ("ret_igbvl_pct", "IGBVL", 252)]
PARAMETROS = [("mu", "μ (media)"), ("omega", "ω"), ("alpha[1]", "α (ARCH)"),
              ("gamma[1]", "γ (asimetría)"), ("beta[1]", "β (GARCH)"), ("nu", "ν (grados de libertad)")]
columnas, garch = {}, {}
for columna, nombre, dias in SERIES_GARCH:
    y = d[columna].dropna()
    for asimetria, modelo in [(0, "GARCH"), (1, "GJR")]:
        res = arch_model(y, mean="Constant", vol="GARCH", p=1, o=asimetria, q=1, dist="t").fit(disp="off")
        clave = f"{nombre} {modelo}"
        garch[clave] = (res, dias)
        valores = {}
        for param, etiqueta in PARAMETROS:
            valores[etiqueta] = (f"{res.params[param]:.4f}{estrellas(res.pvalues[param])}"
                                 if param in res.params.index else "—")
        persistencia = res.params["alpha[1]"] + res.params["beta[1]"] + res.params.get("gamma[1]", 0) / 2
        valores["Persistencia (α+β+γ/2)"] = f"{persistencia:.4f}"
        valores["Vida media del choque (días)"] = (f"{np.log(0.5) / np.log(persistencia):.1f}"
                                                   if persistencia < 0.99999 else "Muy larga (α+β ≈ 1)")
        valores["Log-verosimilitud"] = f"{res.loglikelihood:.2f}"
        valores["AIC"] = f"{res.aic:.2f}"
        valores["BIC"] = f"{res.bic:.2f}"
        valores["N"] = f"{int(res.nobs)}"
        columnas[clave] = valores
        CLAVE.append(f"{clave}: α = {res.params['alpha[1]']:.4f}, β = {res.params['beta[1]']:.4f}, "
                     f"γ = {res.params.get('gamma[1]', float('nan')):.4f}, persistencia {persistencia:.5f}, "
                     f"ν = {res.params['nu']:.2f}, AIC = {res.aic:.2f}")
t6 = pd.DataFrame(columnas).reset_index().rename(columns={"index": "Parámetro"})
guardar_tabla(t6, "tabla6_garch", "Tabla 6. Modelos de volatilidad condicional GARCH(1,1) y GJR-GARCH(1,1)",
              "Media constante; innovaciones t de Student; estimación por máxima verosimilitud. γ > 0 indica que los "
              "choques negativos elevan más la volatilidad. *** p < 0,01; ** p < 0,05; * p < 0,10. " + FUENTE_GENERAL)
volatilidad_btc = garch["BTC GJR"][0].conditional_volatility * np.sqrt(365)

# ---------------------------------------------------------------------------
# 11. Tabla 7: causalidad de Granger entre volumen y retorno del BTC (formación de precios)
# ---------------------------------------------------------------------------
g = pd.DataFrame({"retorno": d["ret_btc_pct"], "dvolumen": 100 * d["log_volumen_btc"].diff()}).dropna()
vol_a_ret = grangercausalitytests(g[["retorno", "dvolumen"]], maxlag=REZAGOS_GRANGER)
ret_a_vol = grangercausalitytests(g[["dvolumen", "retorno"]], maxlag=REZAGOS_GRANGER)
filas = []
for k in range(1, REZAGOS_GRANGER + 1):
    f1, p1 = vol_a_ret[k][0]["ssr_ftest"][:2]
    f2, p2 = ret_a_vol[k][0]["ssr_ftest"][:2]
    filas.append({"Rezagos": k, "F: volumen → retorno": f1, "p volumen → retorno": p1,
                  "Decisión (5 %)": "Rechaza H0" if p1 < 0.05 else "No rechaza H0",
                  "F: retorno → volumen": f2, "p retorno → volumen": p2,
                  "Decisión (5 %) ": "Rechaza H0" if p2 < 0.05 else "No rechaza H0"})
t7 = pd.DataFrame(filas)
guardar_tabla(t7, "tabla7_granger", "Tabla 7. Causalidad de Granger entre la variación del volumen y el retorno del Bitcoin",
              f"H0: la primera variable no causa en el sentido de Granger a la segunda. Variación del log-volumen ×100. "
              f"N = {len(g)} días. " + FUENTE_GENERAL, decimales=4)
CLAVE.append("Granger (5 rezagos): p volumen→retorno = %.4f; p retorno→volumen = %.4f"
             % (t7.iloc[-1]["p volumen → retorno"], t7.iloc[-1]["p retorno → volumen"]))

# ---------------------------------------------------------------------------
# 12. Tabla 8: resumen anual del mercado del Bitcoin (precio, capitalización, volumen, riesgo)
# ---------------------------------------------------------------------------
anual = d.groupby(d.index.year).agg(
    precio_medio=("btc_precio_usd", "mean"), precio_cierre=("btc_precio_usd", "last"),
    capitalizacion_media=("btc_capitalizacion_mmusd", "mean"), volumen_medio=("btc_volumen_musd", "mean"),
    retorno_anual=("ret_btc_pct", "sum"), desv=("ret_btc_pct", "std"))
tc_cierre = d["tc_pen_usd"].dropna().groupby(d["tc_pen_usd"].dropna().index.year).last()
t8 = pd.DataFrame({
    "Año": anual.index.astype(str),
    "Precio medio (USD)": anual["precio_medio"].values, "Precio de cierre (USD)": anual["precio_cierre"].values,
    "Precio de cierre (S/)": (anual["precio_cierre"] * tc_cierre.reindex(anual.index)).values,
    "Capitalización media (miles de mill. USD)": anual["capitalizacion_media"].values,
    "Volumen diario medio (mill. USD)": anual["volumen_medio"].values,
    "Retorno log. anual (%)": anual["retorno_anual"].values,
    "Volatilidad anual (%)": (anual["desv"] * np.sqrt(365)).values,
})
guardar_tabla(t8, "tabla8_resumen_anual", "Tabla 8. Evolución anual del mercado del Bitcoin, 2018-2025",
              "Capitalización de mercado de Coin Metrics (CapMrktCurUSD); precio y volumen de Yahoo Finance; "
              "precio en soles con el último tipo de cambio interbancario del año (BCRP). " + FUENTE_GENERAL, decimales=2)
# ---------------------------------------------------------------------------
# 12b. Tabla 9: prueba de Chow de cambio estructural en enero de 2022
# ---------------------------------------------------------------------------
# ¿La ecuación del retorno del BTC (Tabla 4) es distinta antes y después del 01/01/2022?
# (a) Chow clásico: compara la suma de residuos al cuadrado de la regresión conjunta con la de las
#     dos submuestras. (b) Versión robusta: regresión con una dummy D (1 desde 2022) e interacciones
#     D × regresor, con errores Newey-West; se prueba con Wald que todos los términos con D valen cero.
#     (c) Solo el cambio en la pendiente del S&P 500 (prueba t robusta de D × S&P 500).
X_total = sm.add_constant(conjunta[REGRESORES])
k_param = X_total.shape[1]                                   # 8 parámetros: constante + 7 regresores
n_total = len(conjunta)
ssr_conjunta = sm.OLS(conjunta["BTC"], X_total).fit().ssr
ssr_1 = sm.OLS(periodos["2018-2021"]["BTC"], sm.add_constant(periodos["2018-2021"][REGRESORES])).fit().ssr
ssr_2 = sm.OLS(periodos["2022-2025"]["BTC"], sm.add_constant(periodos["2022-2025"][REGRESORES])).fit().ssr
gl_den = n_total - 2 * k_param
f_chow = ((ssr_conjunta - ssr_1 - ssr_2) / k_param) / ((ssr_1 + ssr_2) / gl_den)
p_chow = stats.f.sf(f_chow, k_param, gl_den)

dummy = (conjunta.index >= inicio2).astype(float)
X_quiebre = X_total.copy()
X_quiebre["D"] = dummy
for variable in REGRESORES:
    X_quiebre[f"D x {variable}"] = dummy * conjunta[variable]
res_quiebre = sm.OLS(conjunta["BTC"], X_quiebre).fit(cov_type="HAC", cov_kwds={"maxlags": REZAGOS_NEWEY_WEST})
terminos_d = [c for c in X_quiebre.columns if c.startswith("D")]
R = np.zeros((len(terminos_d), X_quiebre.shape[1]))
for i, c in enumerate(terminos_d):
    R[i, list(X_quiebre.columns).index(c)] = 1
wald = res_quiebre.f_test(R)
f_wald, p_wald = float(np.squeeze(wald.fvalue)), float(np.squeeze(wald.pvalue))
coef_dsp, p_dsp = res_quiebre.params["D x S&P 500"], res_quiebre.pvalues["D x S&P 500"]

t9 = pd.DataFrame([
    {"Prueba": "Chow clásico (todos los coeficientes)", "Estadístico": f_chow,
     "Grados de libertad": f"F({k_param}, {gl_den})", "p": p_chow},
    {"Prueba": "Wald robusto Newey-West (D e interacciones)", "Estadístico": f_wald,
     "Grados de libertad": f"F({len(terminos_d)}, {int(res_quiebre.df_resid)})", "p": p_wald},
    {"Prueba": "Cambio en la pendiente del S&P 500 (D × S&P 500)", "Estadístico": coef_dsp,
     "Grados de libertad": "coeficiente (t robusta)", "p": p_dsp},
])
t9["Decisión (5 %)"] = np.where(t9["p"] < 0.05, "Hay cambio estructural", "Sin evidencia de cambio")
guardar_tabla(t9, "tabla9_chow", "Tabla 9. Pruebas de cambio estructural en enero de 2022",
              f"H0: los coeficientes de la ecuación de la Tabla 4 son iguales en 2018-2021 y 2022-2025. D = 1 desde el "
              f"{inicio2:%d/%m/%Y}. La prueba robusta usa errores Newey-West con {REZAGOS_NEWEY_WEST} rezagos; la última "
              f"fila es el cambio estimado en la sensibilidad del BTC al S&P 500. N = {n_total}. " + FUENTE_GENERAL,
              decimales=4)
CLAVE.append(f"Chow (quiebre 01/01/2022): F = {f_chow:.3f} (p = {p_chow:.4f}); Wald robusto F = {f_wald:.3f} "
             f"(p = {p_wald:.4f}); cambio pendiente S&P 500 = {coef_dsp:.3f} (p = {p_dsp:.4f})")

# ---------------------------------------------------------------------------
# 12c. Tabla 10: portafolio de un inversionista peruano en soles (BVL con y sin Bitcoin)
# ---------------------------------------------------------------------------
# Días hábiles de la BVL con dato del IGBVL y del tipo de cambio; el BTC se convierte a soles con el
# tipo de cambio del mismo día. Pesos fijos rebalanceados cada día. Rendimientos simples.
PESOS_BTC = [0.0, 0.05, 0.10]
p_sol = d[["igbvl_indice", "tc_pen_usd", "btc_precio_usd"]].dropna()
r_bvl = p_sol["igbvl_indice"].pct_change()
r_btc_sol = (p_sol["btc_precio_usd"] * p_sol["tc_pen_usd"]).pct_change()
rendimientos = pd.DataFrame({
    f"{100 - 100 * w:.0f} % BVL + {100 * w:.0f} % BTC" if w > 0 else "100 % BVL": (1 - w) * r_bvl + w * r_btc_sol
    for w in PESOS_BTC}).dropna()
valor_portafolio = 100 * (1 + rendimientos).cumprod()


def indicadores_portafolio(r):
    """Rendimiento y volatilidad anualizados, razón rendimiento/riesgo, VaR 5 % diario y caída máxima."""
    anios = (r.index[-1] - r.index[0]).days / 365.25
    dias_por_anio = len(r) / anios
    valor = (1 + r).cumprod()
    rend = valor.iloc[-1] ** (1 / anios) - 1
    vol = r.std() * np.sqrt(dias_por_anio)
    return {"Rend. anual (%)": 100 * rend, "Volatilidad anual (%)": 100 * vol, "Rend./riesgo": rend / vol,
            "VaR 5 % diario (%)": 100 * r.quantile(0.05), "Caída máxima (%)": 100 * (valor / valor.cummax() - 1).min(),
            "Valor final de S/ 100": 100 * valor.iloc[-1]}


filas = []
for etiqueta, (ini, fin) in {"2018-2025": (FECHA_INICIO, FECHA_CORTE), "2018-2021": (FECHA_INICIO, CORTE_SUBPERIODO),
                             "2022-2025": (inicio2, FECHA_CORTE)}.items():
    for portafolio in rendimientos.columns:
        filas.append({"Periodo": etiqueta, "Portafolio": portafolio,
                      **indicadores_portafolio(rendimientos.loc[ini:fin, portafolio])})
t10 = pd.DataFrame(filas)
guardar_tabla(t10, "tabla10_portafolio", "Tabla 10. Portafolio de un inversionista peruano en soles: BVL con y sin Bitcoin",
              f"Rendimientos simples diarios en soles en {len(rendimientos)} días hábiles de la BVL; el BTC se convierte "
              f"con el tipo de cambio interbancario del mismo día (BCRP). Pesos fijos rebalanceados diariamente; sin costos "
              f"de transacción. Rendimiento anual compuesto; volatilidad anualizada según los días hábiles por año. "
              + FUENTE_GENERAL, decimales=2)
for _, f in t10[t10["Periodo"] == "2018-2025"].iterrows():
    CLAVE.append(f"Portafolio {f['Portafolio']} (2018-2025, soles): rend. anual {f['Rend. anual (%)']:.2f} %, "
                 f"vol. anual {f['Volatilidad anual (%)']:.2f} %, VaR 5 % {f['VaR 5 % diario (%)']:.2f} %, "
                 f"caída máx. {f['Caída máxima (%)']:.2f} %, S/ 100 -> S/ {f['Valor final de S/ 100']:.0f}")

guardar_libro_consolidado()

# ---------------------------------------------------------------------------
# 13. Figuras para el artículo (estilo de revista, ancho de columna exacto)
# ---------------------------------------------------------------------------
# Figura 1: precio, capitalización y caída desde el máximo (drawdown) del BTC
fig, ejes = plt.subplots(3, 1, figsize=(ANCHO, 6.6), sharex=True, gridspec_kw={"height_ratios": [1.45, 1, 0.85]})
eje = ejes[0]
marcar_subperiodos(eje, etiquetar=False)
marcar_eventos(eje)
eje.plot(d.index, d["btc_precio_usd"], color=AZUL, linewidth=0.9)
eje.set_yscale("log")
formato_log(eje)
eje.set_ylabel("USD por BTC (log.)")
eje.set_ylim(top=d["btc_precio_usd"].max() * 1.6)          # espacio para el rótulo del máximo
titulo_panel(eje, "(a) Precio de cierre diario")
maximo, minimo = d["btc_precio_usd"].idxmax(), d["btc_precio_usd"].idxmin()
for fecha, texto, desp, alin in [
        (maximo, f"Máximo: {d.loc[maximo, 'btc_precio_usd']:,.0f} USD ({maximo:%d/%m/%Y})", (-12, 5), "right"),
        (minimo, f"Mínimo: {d.loc[minimo, 'btc_precio_usd']:,.0f} USD ({minimo:%d/%m/%Y})", (10, -4), "left")]:
    eje.plot([fecha], [d.loc[fecha, "btc_precio_usd"]], "o", ms=3.5, mfc="white", mec=AZUL, mew=1, zorder=5)
    eje.annotate(texto, xy=(fecha, d.loc[fecha, "btc_precio_usd"]), xytext=desp, textcoords="offset points",
                 fontsize=7.2, ha=alin, va="center", color=TINTA,
                 bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.9),
                 arrowprops=dict(arrowstyle="-", color=TINTA_2, lw=0.5, shrinkA=0, shrinkB=3))
eje = ejes[1]
marcar_subperiodos(eje, etiquetar=False)
marcar_eventos(eje, rotulos=False)
eje.fill_between(d.index, d["btc_capitalizacion_mmusd"], color=AZUL, alpha=0.15, linewidth=0)
eje.plot(d.index, d["btc_capitalizacion_mmusd"], color=AZUL, linewidth=0.8)
eje.set_ylim(bottom=0)
miles(eje)
eje.set_ylabel("Miles de mill. de USD")
titulo_panel(eje, "(b) Capitalización de mercado")
eje = ejes[2]
caida = 100 * (d["btc_precio_usd"] / d["btc_precio_usd"].cummax() - 1)
marcar_subperiodos(eje)
marcar_eventos(eje, rotulos=False)
eje.fill_between(d.index, caida, 0, color=NARANJA, alpha=0.25, linewidth=0)
eje.plot(d.index, caida, color=NARANJA, linewidth=0.7)
peor = caida.idxmin()
eje.annotate(f"{caida.min():.0f} % ({peor:%m/%Y})".replace("-", "−"), xy=(peor, caida.min()), xytext=(8, 4), textcoords="offset points",
             fontsize=7.2, color=TINTA)
eje.set_ylim(top=0)
eje.set_ylabel("% bajo el máximo")
titulo_panel(eje, "(c) Caída desde el máximo previo de la muestra (drawdown)")
eje_fechas(eje)
fig.align_ylabels(ejes)
fig.tight_layout(h_pad=0.6)
guardar_figura(fig, "figura1_precio_capitalizacion",
               "Fuente: elaboración propia con datos de Yahoo Finance (precio) y Coin Metrics (capitalización). "
               "En (c), el máximo previo se mide desde el 01/01/2018. Zona sombreada: subperiodo 2022-2025. " + NOTA_EVENTOS)
CLAVE.append(f"Máxima caída del BTC desde su máximo previo: {caida.min():.1f} % ({peor:%d/%m/%Y})")

# Figura 2: desempeño acumulado, base 100 = primer día común de 2018
base = h[["btc_precio_usd", "sp500_indice", "oro_usd_onza", "igbvl_indice"]]
base = 100 * base / base.iloc[0]
base.columns = ["Bitcoin", "S&P 500", "Oro", "IGBVL"]
fig, eje = plt.subplots(figsize=(ANCHO, 3.3))
marcar_subperiodos(eje)
for nombre in base.columns:
    eje.plot(base.index, base[nombre], color=COLOR[nombre], linestyle=LINEAS[nombre], linewidth=1.0, label=nombre)
eje.axhline(100, color=TINTA_2, linewidth=0.5)
eje.set_yscale("log")
formato_log(eje)
rotulos_finales(eje, base, log_y=True)
eje.set_ylabel("Índice, base 100 = 04/01/2018 (log.)")
eje.legend(loc="upper left", ncol=4, handlelength=3)
eje_fechas(eje)
fig.tight_layout()
guardar_figura(fig, "figura2_desempeno_acumulado",
               "Fuente: elaboración propia con datos de Yahoo Finance y BCRPData. Días hábiles con información de todos "
               "los activos; el valor al final de cada línea es el índice al 31/12/2025. Zona sombreada: subperiodo "
               "2022-2025.")
CLAVE.append("Índice base 100 al cierre de 2025: " + ", ".join(f"{c} {base[c].iloc[-1]:,.0f}" for c in base.columns))

# Figura 3: volatilidad realizada, un panel por activo con la misma escala (comparación directa)
SERIES_VOL = [("vol30_btc_anual_pct", "Bitcoin", "ventana de 30 días"),
              ("vol21_sp500_anual_pct", "S&P 500", "ventana de 21 días hábiles"),
              ("vol21_igbvl_anual_pct", "IGBVL", "ventana de 21 días hábiles")]
tope = max(d[c].max() for c, _, _ in SERIES_VOL) * 1.05
fig, ejes = plt.subplots(3, 1, figsize=(ANCHO, 5.4), sharex=True, sharey=True)
for k, (eje, (columna, nombre, ventana)) in enumerate(zip(ejes, SERIES_VOL)):
    s3 = d[columna].dropna()
    marcar_subperiodos(eje, etiquetar=(k == 2))
    marcar_eventos(eje, rotulos=(k == 0))
    eje.fill_between(s3.index, s3, color=COLOR[nombre], alpha=0.15, linewidth=0)
    eje.plot(s3.index, s3, color=COLOR[nombre], linewidth=0.8)
    eje.axhline(s3.mean(), color=TINTA, linewidth=0.6, linestyle=(0, (4, 2)))
    eje.text(1.0, s3.mean(), f" media {s3.mean():.1f} %", transform=blended_transform_factory(eje.transAxes, eje.transData),
             fontsize=7, va="bottom", ha="right", color=TINTA,
             bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.9))
    titulo_panel(eje, f"({'abc'[k]}) {nombre} ({ventana})")
    eje.set_ylim(0, tope)
ejes[1].set_ylabel("Volatilidad anualizada (%)")
eje_fechas(ejes[-1])
fig.tight_layout(h_pad=0.5)
guardar_figura(fig, "figura3_volatilidad_realizada",
               "Fuente: elaboración propia con datos de Yahoo Finance y BCRPData. Desviación estándar móvil de los "
               "retornos, anualizada con raíz de 365 (BTC) y raíz de 252 (S&P 500 e IGBVL); misma escala en los tres "
               "paneles. Zona sombreada: subperiodo 2022-2025. "
               + NOTA_EVENTOS)

# Figura 4: correlación móvil del BTC con el S&P 500 y el oro, con promedios por subperiodo
fig, eje = plt.subplots(figsize=(ANCHO, 3.3))
marcar_subperiodos(eje)
for variable in ["S&P 500", "Oro"]:
    r = conjunta["BTC"].rolling(VENTANA_CORRELACION).corr(conjunta[variable])
    eje.plot(r.index, r, color=COLOR[variable], linestyle=LINEAS[variable], linewidth=0.9, label=f"BTC – {variable}")
    for etiqueta, (ini, fin) in {"2018-2021": (FECHA_INICIO, CORTE_SUBPERIODO),
                                 "2022-2025": (inicio2, FECHA_CORTE)}.items():
        valor = correlaciones[(variable, etiqueta)]
        eje.hlines(valor, pd.Timestamp(ini), pd.Timestamp(fin), color=COLOR[variable], linewidth=1.8, alpha=0.9, zorder=4)
        eje.annotate(f"ρ = {valor:.2f}".replace("-", "−"), xy=(pd.Timestamp(fin), valor), xytext=(-3, 3 if variable == "S&P 500" else -9),
                     textcoords="offset points", ha="right", fontsize=7.2, color=TINTA, zorder=6,
                     bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.85))
eje.axhline(0, color=TINTA_2, linewidth=0.5)
eje.set_ylabel(f"Correlación móvil ({VENTANA_CORRELACION} días hábiles)")
eje.legend(loc="lower left", ncol=2, handlelength=3, bbox_to_anchor=(0, 1.0))
eje_fechas(eje)
fig.tight_layout()
guardar_figura(fig, "figura4_correlacion_movil",
               "Fuente: elaboración propia con datos de Yahoo Finance. Segmentos horizontales: correlación de Pearson "
               "de cada subperiodo (Tabla 3). Zona sombreada: subperiodo 2022-2025.")

# Figura 5: retornos diarios y volatilidad condicional GJR-GARCH del BTC
fig, ejes = plt.subplots(2, 1, figsize=(ANCHO, 4.6), sharex=True, gridspec_kw={"height_ratios": [1, 1.15]})
r_btc = d["ret_btc_pct"].dropna()
marcar_eventos(ejes[0])
ejes[0].vlines(r_btc.index, 0, r_btc, color=np.where(r_btc >= 0, AZUL, NARANJA), linewidth=0.45)
ejes[0].axhline(0, color=TINTA_2, linewidth=0.4)
ejes[0].set_ylabel("Retorno diario (%)")
titulo_panel(ejes[0], "(a) Retornos logarítmicos diarios del BTC")
realizada = d["vol30_btc_anual_pct"].dropna()
marcar_eventos(ejes[1], rotulos=False)
ejes[1].plot(realizada.index, realizada, color=GRIS, linewidth=0.8, label="Realizada (30 días)")
ejes[1].plot(volatilidad_btc.index, volatilidad_btc, color=AZUL, linewidth=0.9, label="Condicional GJR-GARCH(1,1)")
ejes[1].set_ylabel("Volatilidad anualizada (%)")
ejes[1].set_ylim(bottom=0)
titulo_panel(ejes[1], "(b) Volatilidad condicional estimada y volatilidad realizada")
ejes[1].legend(loc="upper right", ncol=2)
eje_fechas(ejes[1])
fig.align_ylabels(ejes)
fig.tight_layout(h_pad=0.6)
guardar_figura(fig, "figura5_volatilidad_condicional",
               "Fuente: elaboración propia con datos de Yahoo Finance; modelo de la Tabla 6. " + NOTA_EVENTOS)

# Figura 6: distribución de los retornos del BTC (histograma y gráfico Q-Q)
s = d["ret_btc_pct"].dropna()
x = np.linspace(-20, 20, 500)
t_param = stats.t.fit(s)
fig, ejes = plt.subplots(1, 2, figsize=(ANCHO, 2.9), gridspec_kw={"width_ratios": [1.35, 1]})
eje = ejes[0]
eje.hist(s, bins=np.arange(-20, 20.5, 0.5), density=True, color=AZUL, alpha=0.35, edgecolor=AZUL, linewidth=0.3,
         label="Retornos del BTC")
eje.plot(x, stats.norm.pdf(x, s.mean(), s.std()), color=NARANJA, linewidth=1.1, linestyle=LINEAS["S&P 500"],
         label="Normal")
eje.plot(x, stats.t.pdf(x, *t_param), color=TINTA, linewidth=0.9, label=f"t de Student (ν = {t_param[0]:.1f})")
eje.set_xlim(-15, 15)
eje.set_xlabel("Retorno logarítmico diario (%)")
eje.set_ylabel("Densidad")
titulo_panel(eje, "(a) Densidad empírica y teóricas")
eje.legend(loc="upper left")
eje = ejes[1]
(teoricos, observados), (pendiente, intercepto, _) = stats.probplot(s, dist="norm")
eje.plot(teoricos, pendiente * teoricos + intercepto, color=NARANJA, linewidth=1.0, linestyle=LINEAS["S&P 500"])
eje.plot(teoricos, observados, "o", ms=2, mfc="none", mec=AZUL, mew=0.5)
eje.set_xlabel("Cuantiles teóricos de la normal")
eje.set_ylabel("Cuantiles observados (%)")
titulo_panel(eje, "(b) Gráfico Q-Q")
eje.grid(axis="x")
fig.tight_layout(w_pad=1.5)
guardar_figura(fig, "figura6_distribucion_retornos",
               "Fuente: elaboración propia con datos de Yahoo Finance. En (a) el eje se recorta a ±15 %; los puntos "
               "alejados de la recta en (b) evidencian colas pesadas.")

# Figura 7: curvas de impacto de noticias (news impact curve) del GJR-GARCH
fig, eje = plt.subplots(figsize=(ANCHO * 0.62, 3.0))
z = np.linspace(-5, 5, 401)
for nombre_modelo, nombre in [("BTC GJR", "Bitcoin"), ("S&P 500 GJR", "S&P 500"), ("IGBVL GJR", "IGBVL")]:
    res, _ = garch[nombre_modelo]
    p = res.params
    var_ref = np.median(res.conditional_volatility ** 2)          # varianza condicional típica
    eps = z * np.sqrt(var_ref)
    curva = p["omega"] + (p["alpha[1]"] + p["gamma[1]"] * (eps < 0)) * eps ** 2 + p["beta[1]"] * var_ref
    eje.plot(z, curva / curva[200], color=COLOR[nombre], linestyle=LINEAS[nombre], linewidth=1.1, label=nombre)
eje.axvline(0, color=TINTA_2, linewidth=0.5)
eje.set_xlabel("Choque de ayer (en desviaciones estándar)")
eje.set_ylabel("Varianza de hoy (choque nulo = 1)")
eje.grid(axis="x")
eje.legend(loc="upper center")
fig.tight_layout()
guardar_figura(fig, "figura7_curva_impacto_noticias",
               "Fuente: elaboración propia; parámetros GJR-GARCH(1,1) de la Tabla 6. Una curva asimétrica indica que las "
               "caídas elevan más la volatilidad que las subidas.")

# Figura 8: matrices de correlación por subperiodo (mapa de calor divergente)
from matplotlib.colors import LinearSegmentedColormap
divergente = LinearSegmentedColormap.from_list("divergente", [AZUL, "#f2f1ec", NARANJA])
cortos = ["BTC", "S&P", "Oro", "IGBVL", "TC", "VIX", "UST10", "Dólar"]
fig, ejes = plt.subplots(1, 2, figsize=(ANCHO, 3.3))
for eje, etiqueta in zip(ejes, ["2018-2021", "2022-2025"]):
    m = periodos[etiqueta].corr().values[1:, :-1]             # sin la diagonal (siempre 1)
    mascara = np.triu(np.ones_like(m, dtype=bool), k=1)
    imagen = eje.imshow(np.ma.masked_where(mascara, m), cmap=divergente, vmin=-1, vmax=1)
    for i in range(m.shape[0]):
        for j in range(i + 1):
            eje.text(j, i, f"{m[i, j]:.2f}".replace("-", "−"), ha="center", va="center",
                     fontsize=6.5, color="white" if abs(m[i, j]) > 0.6 else TINTA)
    eje.set_xticks(range(len(cortos) - 1), cortos[:-1], rotation=45, ha="right")
    eje.set_yticks(range(len(cortos) - 1), cortos[1:])
    eje.grid(False)
    for lado in eje.spines.values():
        lado.set_visible(False)
    eje.tick_params(length=0)
    titulo_panel(eje, f"({'ab'[ejes.tolist().index(eje)]}) {etiqueta}")
barra = fig.colorbar(imagen, ax=ejes, shrink=0.8, pad=0.02)
barra.outline.set_visible(False)
barra.ax.tick_params(labelsize=7)
guardar_figura(fig, "figura8_matrices_correlacion",
               "Fuente: elaboración propia con datos de Yahoo Finance, BCRPData y FRED. Correlaciones de Pearson de los "
               "retornos diarios en la muestra conjunta de días hábiles.")

# Figura 9: valor de S/ 100 invertidos en los tres portafolios del inversionista peruano
COLOR_PORTAFOLIO = [AMARILLO, VERDE_AGUA, AZUL]
LINEA_PORTAFOLIO = [(0, (6, 1.5, 1.2, 1.5)), (0, (5, 1.5)), "-"]
fig, eje = plt.subplots(figsize=(ANCHO, 3.3))
marcar_subperiodos(eje)
for color, linea, portafolio in zip(COLOR_PORTAFOLIO, LINEA_PORTAFOLIO, valor_portafolio.columns):
    eje.plot(valor_portafolio.index, valor_portafolio[portafolio], color=color, linestyle=linea, linewidth=1.0,
             label=portafolio)
    final = valor_portafolio[portafolio].iloc[-1]
    eje.annotate(f"S/ {final:,.0f}", xy=(valor_portafolio.index[-1], final), xytext=(4, 0),
                 textcoords="offset points", fontsize=7.2, color=TINTA, va="center", annotation_clip=False)
eje.axhline(100, color=TINTA_2, linewidth=0.5)
eje.set_ylabel("Valor de S/ 100 invertidos")
eje.legend(loc="upper left", ncol=3, handlelength=3)
eje_fechas(eje)
fig.tight_layout()
guardar_figura(fig, "figura9_portafolio",
               "Fuente: elaboración propia con datos de BCRPData (IGBVL y tipo de cambio) y Yahoo Finance (BTC). Valor en "
               "soles con pesos fijos rebalanceados diariamente (Tabla 10). Zona sombreada: subperiodo 2022-2025.")
# ---------------------------------------------------------------------------
# 14. Dashboard interactivo (HTML autónomo, funciona sin internet)
# ---------------------------------------------------------------------------
def estilo_plotly(figura, alto, titulo_y=None, log_y=False):
    """Estilo común de los gráficos del tablero."""
    figura.update_layout(
        template="plotly_white", height=alto, margin=dict(l=60, r=20, t=40, b=40),
        font=dict(family="Segoe UI, Roboto, Arial, sans-serif", size=12, color=TINTA),
        hovermode="x unified", legend=dict(orientation="h", y=1.08, x=0),
        paper_bgcolor="white", plot_bgcolor="white",
    )
    figura.update_xaxes(showgrid=False, linecolor=GRIS)
    figura.update_yaxes(gridcolor=GRIS_CLARO, zeroline=False, title_text=titulo_y,
                        type="log" if log_y else "linear")
    return figura


def eventos_plotly(figura, filas=1):
    """Líneas punteadas de los hitos con su rótulo al pasar el mouse."""
    for fecha, texto in EVENTOS:
        if filas > 1:
            for fila in range(1, filas + 1):
                figura.add_vline(x=fecha, line_width=1, line_dash="dot", line_color=GRIS, row=fila, col=1)
        else:
            figura.add_vline(x=fecha, line_width=1, line_dash="dot", line_color=GRIS)
        figura.add_annotation(x=fecha, y=1.0, yref="paper", text=texto.replace("\n", " "), showarrow=False,
                              font=dict(size=9, color=TINTA_2), textangle=-90, xanchor="right", yanchor="top")
    return figura


def ticks_log(figura, valores, **kwargs):
    """Eje logarítmico con marcas 1-2-5 y números completos (evita los rótulos menores 2, 3, 4...)."""
    minimo, maximo = np.nanmin(valores), np.nanmax(valores)
    marcas = [m * 10 ** e for e in range(-1, 8) for m in (1, 2, 5) if minimo * 0.8 <= m * 10 ** e <= maximo * 1.25]
    figura.update_yaxes(type="log", tickvals=marcas, ticktext=[f"{m:,.0f}" for m in marcas], **kwargs)
    return figura


def selector_rango(figura):
    """Botones para elegir el periodo visible."""
    figura.update_xaxes(rangeselector=dict(buttons=[
        dict(count=6, label="6 m", step="month", stepmode="backward"),
        dict(count=1, label="1 año", step="year", stepmode="backward"),
        dict(count=3, label="3 años", step="year", stepmode="backward"),
        dict(step="all", label="Todo")], bgcolor="#f3f2ee", activecolor="#cde2fb"), row=1, col=1)
    return figura


def html_figura(figura, primera=False):
    return figura.to_html(full_html=False, include_plotlyjs=True if primera else False,
                          config={"displaylogo": False, "responsive": True})


def html_tabla(tabla, decimales=3):
    """Tabla HTML con números formateados."""
    t = tabla.copy()
    for c in t.select_dtypes("number").columns:
        t[c] = t[c].map(lambda v: "" if pd.isna(v) else (f"{v:,.0f}" if float(v).is_integer() and abs(v) >= 1
                                                          else f"{v:,.{decimales}f}"))
    return t.to_html(index=False, classes="tabla", border=0, escape=True)


def crear_dashboard():
    # Indicadores principales (tarjetas)
    precio_fin = d["btc_precio_usd"].dropna().iloc[-1]
    cap_fin = d["btc_capitalizacion_mmusd"].dropna().iloc[-1]
    vol_btc = t1.loc[t1.iloc[:, 0] == "Bitcoin", "Vol. anual (%)"].iloc[0]
    vol_sp = t1.loc[t1.iloc[:, 0] == "S&P 500", "Vol. anual (%)"].iloc[0]
    beta1, beta2 = resultados["2018-2021"].params["S&P 500"], resultados["2022-2025"].params["S&P 500"]
    p2 = resultados["2022-2025"].pvalues["S&P 500"]
    res_gjr = garch["BTC GJR"][0]
    persist = res_gjr.params["alpha[1]"] + res_gjr.params["beta[1]"] + res_gjr.params["gamma[1]"] / 2
    p_granger = t7.iloc[-1]["p retorno → volumen"]
    tarjetas = [
        ("Precio del BTC al 31-12-2025", f"{precio_fin:,.0f} USD", f"Máximo del periodo: {d['btc_precio_usd'].max():,.0f} USD"),
        ("Capitalización de mercado", f"{cap_fin:,.0f} mil mill. USD", "Coin Metrics · CapMrktCurUSD"),
        ("Volatilidad anual del BTC", f"{vol_btc:.1f} %", f"{vol_btc / vol_sp:.1f} veces la del S&P 500 ({vol_sp:.1f} %)"),
        ("Beta del BTC con el S&P 500", f"{beta1:.2f} → {beta2:.2f}", f"2018-2021 → 2022-2025 (p = {p2:.4f})"),
        ("Persistencia GJR-GARCH", f"{persist:.4f}",
         f"α + β ≈ 1: choques de volatilidad casi permanentes · γ = {res_gjr.params['gamma[1]']:.4f}"),
        ("Granger retorno → volumen", f"p = {p_granger:.4f}", "El precio se forma antes que el volumen"),
    ]
    html_tarjetas = "".join(
        f'<div class="tarjeta"><div class="etq">{html.escape(a)}</div><div class="valor">{html.escape(b)}</div>'
        f'<div class="nota">{html.escape(c)}</div></div>' for a, b, c in tarjetas)

    # Gráfico 1: precio y capitalización
    f1 = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.07,
                       subplot_titles=("Precio de cierre (USD, escala log.)", "Capitalización (miles de millones de USD)"))
    f1.add_trace(go.Scatter(x=d.index, y=d["btc_precio_usd"], name="Precio BTC", line=dict(color=AZUL, width=1.6),
                            hovertemplate="%{y:,.0f} USD"), row=1, col=1)
    f1.add_trace(go.Scatter(x=d.index, y=d["btc_capitalizacion_mmusd"], name="Capitalización", fill="tozeroy",
                            line=dict(color=AZUL, width=1.2), fillcolor="rgba(42,120,214,0.15)",
                            hovertemplate="%{y:,.0f} mil mill. USD"), row=2, col=1)
    estilo_plotly(f1, 620)
    f1.update_layout(showlegend=False)          # cada panel tiene su título; así no choca con los botones
    ticks_log(f1, d["btc_precio_usd"], row=1, col=1)
    selector_rango(f1)
    eventos_plotly(f1, filas=2)

    # Gráfico 2: desempeño acumulado
    f2 = go.Figure([go.Scatter(x=base.index, y=base[c], name=c, line=dict(color=COLOR[c], width=1.6),
                               hovertemplate="%{y:,.1f}") for c in base.columns])
    estilo_plotly(f2, 430, "Base 100 = enero 2018 (log.)")
    ticks_log(f2, base.values)

    # Gráfico 3: volatilidad realizada
    f3 = go.Figure()
    for columna, nombre in [("vol30_btc_anual_pct", "Bitcoin"), ("vol21_sp500_anual_pct", "S&P 500"),
                            ("vol21_igbvl_anual_pct", "IGBVL")]:
        s3 = d[columna].dropna()
        f3.add_trace(go.Scatter(x=s3.index, y=s3, name=nombre, line=dict(color=COLOR[nombre], width=1.4),
                                hovertemplate="%{y:.1f} %"))
    estilo_plotly(f3, 430, "Volatilidad anualizada (%)")
    eventos_plotly(f3)

    # Gráfico 4: correlación móvil
    f4 = go.Figure()
    for variable in ["S&P 500", "Oro"]:
        r = conjunta["BTC"].rolling(VENTANA_CORRELACION).corr(conjunta[variable])
        f4.add_trace(go.Scatter(x=r.index, y=r, name=f"BTC – {variable}", line=dict(color=COLOR[variable], width=1.4),
                                hovertemplate="%{y:.2f}"))
    f4.add_hline(y=0, line_color=GRIS, line_width=1)
    f4.add_vrect(x0=inicio2.strftime("%Y-%m-%d"), x1=FECHA_CORTE, fillcolor="#f3f2ee", line_width=0, layer="below")
    estilo_plotly(f4, 430, f"Correlación móvil ({VENTANA_CORRELACION} días hábiles)")

    # Gráfico 5: volatilidad condicional
    f5 = go.Figure([
        go.Scatter(x=realizada.index, y=realizada, name="Realizada (30 días)", line=dict(color=GRIS, width=1.2),
                   hovertemplate="%{y:.1f} %"),
        go.Scatter(x=volatilidad_btc.index, y=volatilidad_btc, name="Condicional GJR-GARCH", line=dict(color=AZUL, width=1.6),
                   hovertemplate="%{y:.1f} %")])
    estilo_plotly(f5, 430, "Volatilidad anualizada del BTC (%)")
    eventos_plotly(f5)

    # Gráfico 6: distribución
    f6 = go.Figure([go.Histogram(x=s, histnorm="probability density", nbinsx=160, name="Retornos BTC",
                                 marker_color=AZUL, opacity=0.8, hovertemplate="%{x:.1f} %: %{y:.3f}"),
                    go.Scatter(x=x, y=stats.norm.pdf(x, s.mean(), s.std()), name="Normal",
                               line=dict(color=NARANJA, width=2))])
    estilo_plotly(f6, 430, "Densidad")
    f6.update_layout(hovermode="closest", bargap=0.02)
    f6.update_xaxes(range=[-20, 20], title_text="Retorno logarítmico diario (%)")

    # Tablas del tablero
    tabla_desc = t1[["Serie (retorno diario, %)", "N", "Media", "Desv. est.", "Mínimo", "Máximo",
                     "Curtosis exceso", "Vol. anual (%)"]]

    plantilla = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard · Criptoactivos como instrumento alternativo</title>
<style>
 body{{margin:0;background:#f6f5f1;font-family:"Segoe UI",Roboto,Arial,sans-serif;color:{TINTA}}}
 header{{background:#{AZUL_TABLA};color:#fff;padding:26px 40px}}
 header h1{{margin:0;font-size:24px}} header p{{margin:6px 0 0;opacity:.85;font-size:14px}}
 main{{max-width:1280px;margin:0 auto;padding:24px 28px 40px}}
 .tarjetas{{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:22px}}
 .tarjeta{{background:#fff;border-radius:12px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,0,0,.08);border-top:4px solid {AZUL}}}
 .etq{{font-size:12px;color:{TINTA_2};text-transform:uppercase;letter-spacing:.03em}}
 .valor{{font-size:24px;font-weight:700;margin:6px 0 4px}} .nota{{font-size:12px;color:{TINTA_2}}}
 .panel{{background:#fff;border-radius:12px;padding:18px 20px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.08)}}
 .panel h2{{font-size:17px;margin:0 0 4px}} .panel .desc{{font-size:13px;color:{TINTA_2};margin:0 0 8px}}
 .rejilla{{display:grid;grid-template-columns:1fr 1fr;gap:20px}} @media(max-width:900px){{.rejilla{{grid-template-columns:1fr}}}}
 table.tabla{{border-collapse:collapse;width:100%;font-size:13px}}
 table.tabla th{{background:#{AZUL_TABLA};color:#fff;padding:8px;text-align:center}}
 table.tabla td{{padding:7px 8px;text-align:center;border-bottom:1px solid #e6e6e6}}
 table.tabla td:first-child{{text-align:left;font-weight:600}} table.tabla tr:nth-child(even) td{{background:#{FILA_ALTERNA}}}
 footer{{font-size:12px;color:{TINTA_2};text-align:center;padding:10px 20px 30px}}
</style></head><body>
<header><h1>Criptoactivos como instrumento alternativo: formación de precios y volatilidad</h1>
<p>{AUTORA} · Código {CODIGO_MATRICULA} · Finanzas I (055D), UNCP · Datos diarios {FECHA_INICIO} a {FECHA_CORTE}</p></header>
<main>
<section class="tarjetas">{html_tarjetas}</section>
<section class="panel"><h2>1. Precio y capitalización de mercado del Bitcoin</h2>
<p class="desc">Use los botones para cambiar el periodo y pase el mouse para ver los valores diarios. Líneas punteadas: hitos del mercado.</p>
{html_figura(f1, primera=True)}</section>
<div class="rejilla">
<section class="panel"><h2>2. Desempeño acumulado</h2><p class="desc">Días hábiles comunes; base 100 = enero de 2018.</p>{html_figura(f2)}</section>
<section class="panel"><h2>3. Volatilidad realizada anualizada</h2><p class="desc">BTC: 30 días calendario; S&amp;P 500 e IGBVL: 21 días hábiles.</p>{html_figura(f3)}</section>
<section class="panel"><h2>4. ¿Refugio o activo de riesgo?</h2><p class="desc">Correlación móvil del BTC; zona sombreada: 2022-2025.</p>{html_figura(f4)}</section>
<section class="panel"><h2>5. Volatilidad condicional (GJR-GARCH)</h2><p class="desc">Modelo con innovaciones t de Student.</p>{html_figura(f5)}</section>
</div>
<section class="panel"><h2>6. Distribución de los retornos diarios del BTC</h2><p class="desc">Colas más pesadas que la normal (curtosis en exceso).</p>{html_figura(f6)}</section>
<div class="rejilla">
<section class="panel"><h2>Estadísticos descriptivos</h2>{html_tabla(tabla_desc)}</section>
<section class="panel"><h2>Determinantes del retorno del BTC (MCO, Newey-West)</h2>{t4.to_html(index=False, classes="tabla", border=0)}</section>
</div>
<section class="panel"><h2>Evolución anual del mercado del Bitcoin</h2>{html_tabla(t8, 2)}</section>
</main>
<footer>{html.escape(FUENTE_GENERAL)}<br>Generado automáticamente por codigo/04_analisis.py · Repositorio: {html.escape(ENLACE_REPOSITORIO)}</footer>
</body></html>"""
    (SALIDAS / "dashboard_criptoactivos.html").write_text(plantilla, encoding="utf-8")
    log.info("Dashboard guardado: dashboard_criptoactivos.html")


crear_dashboard()

# ---------------------------------------------------------------------------
# 15. Cifras clave para el texto del artículo
# ---------------------------------------------------------------------------
(SALIDAS / "resultados_clave.txt").write_text("\n".join(CLAVE) + "\n", encoding="utf-8")
for linea in CLAVE:
    log.info("CLAVE | %s", linea)
log.info("=== Fin del análisis: tablas, figuras y dashboard en /salidas ===")
