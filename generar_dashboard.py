#!/usr/bin/env python3
"""
Estadincho-Gen v3.0 - Convierte cualquier archivo de datos en un dashboard visual.
Formatos soportados: .xlsx, .xls, .xlsm, .csv, .tsv, .ods, .sav, .dta, .rds, .RData
Uso: python generar_dashboard.py archivo.xlsx
     python generar_dashboard.py datos.sav
     python generar_dashboard.py datos.dta
"""

import sys
import os
import json
import webbrowser
import re
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────
# DEPENDENCIAS
# ─────────────────────────────────────────────

def verificar_dependencias():
    faltantes = []
    try:
        import pandas
    except ImportError:
        faltantes.append("pandas")
    try:
        import openpyxl
    except ImportError:
        faltantes.append("openpyxl")
    try:
        import odf
    except ImportError:
        faltantes.append("odfpy")
    try:
        import pyreadstat
    except ImportError:
        faltantes.append("pyreadstat")  # para .sav y .dta

    if faltantes:
        print("ERROR: Faltan las siguientes librerias:")
        for f in faltantes:
            print(f"  - {f}")
        print("\nEjecuta:")
        print(f"  pip install {' '.join(faltantes)}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

verificar_dependencias()

import pandas as pd

# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def limpiar_nombre(col):
    return str(col).strip()

def es_numerica(series):
    return pd.api.types.is_numeric_dtype(series)

def es_texto_col(series):
    return pd.api.types.is_string_dtype(series) or series.dtype == object

def es_fecha(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    if es_texto_col(series):
        sample = series.dropna().head(20).astype(str)
        hits = sample.str.match(r'\d{1,4}[-/]\d{1,2}[-/]\d{1,4}').sum()
        return hits > len(sample) * 0.5
    return False

def es_categorica(series, umbral=0.5):
    if es_texto_col(series):
        return True
    if pd.api.types.is_integer_dtype(series):
        ratio = series.nunique() / max(len(series), 1)
        return ratio < umbral and series.nunique() < 50
    return False

def top_n(series, n=10):
    vc = series.value_counts().head(n)
    return {
        "labels": [str(x) for x in vc.index.tolist()],
        "values": [int(x) for x in vc.values.tolist()]
    }

def serie_temporal(df, col_fecha, col_valor=None):
    df2 = df.copy()
    df2["__mes__"] = pd.to_datetime(df2[col_fecha], errors="coerce").dt.to_period("M")
    df2 = df2.dropna(subset=["__mes__"])
    if col_valor and es_numerica(df[col_valor]):
        grp = df2.groupby("__mes__")[col_valor].sum()
    else:
        grp = df2.groupby("__mes__").size()
    grp = grp.sort_index()
    return {
        "labels": [str(p) for p in grp.index],
        "values": [float(v) for v in grp.values]
    }

# ─────────────────────────────────────────────
# DETECCION DE ENCODING
# ─────────────────────────────────────────────

def detectar_encoding(ruta):
    encodings = ["utf-8", "utf-8-sig", "latin-1", "cp1252", "iso-8859-1"]
    for enc in encodings:
        try:
            with open(ruta, "r", encoding=enc) as f:
                f.read(2048)
            return enc
        except (UnicodeDecodeError, Exception):
            continue
    return "utf-8"

# ─────────────────────────────────────────────
# DETECCION DE DELIMITADOR
# ─────────────────────────────────────────────

def detectar_delimitador(ruta, encoding):
    delimitadores = [",", ";", "\t", "|"]
    mejor_sep = ","
    mejor_cols = 1
    try:
        with open(ruta, "r", encoding=encoding, errors="replace") as f:
            muestra = f.read(4096)
        for sep in delimitadores:
            cols = len(muestra.split("\n")[0].split(sep))
            if cols > mejor_cols:
                mejor_cols = cols
                mejor_sep = sep
    except Exception:
        pass
    return mejor_sep

# ─────────────────────────────────────────────
# APLICAR ETIQUETAS (SPSS / Stata)
# ─────────────────────────────────────────────

def aplicar_etiquetas(df, meta):
    """
    Reemplaza códigos numéricos por sus etiquetas de valor cuando existen.
    meta es el objeto pyreadstat.metadata_container.
    Devuelve df con columnas etiquetadas y un dict col -> etiqueta de variable.
    """
    etiquetas_col = {}  # col -> nombre largo de la variable

    # Etiquetas de variables (nombres largos)
    if hasattr(meta, "column_labels") and meta.column_labels:
        for i, col in enumerate(df.columns):
            if i < len(meta.column_labels) and meta.column_labels[i]:
                etiquetas_col[col] = meta.column_labels[i]

    # Etiquetas de valores (códigos -> texto)
    if hasattr(meta, "variable_value_labels") and meta.variable_value_labels:
        for col, mapa in meta.variable_value_labels.items():
            if col in df.columns and mapa:
                try:
                    df[col] = df[col].map(lambda x: mapa.get(x, x) if pd.notna(x) else x)
                except Exception:
                    pass

    return df, etiquetas_col

# ─────────────────────────────────────────────
# CARGA DEL ARCHIVO
# ─────────────────────────────────────────────

def cargar_archivo(ruta):
    ruta = Path(ruta)
    if not ruta.exists():
        print(f"ERROR: No se encontro el archivo: {ruta}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

    ext = ruta.suffix.lower()
    print(f"Cargando archivo: {ruta.name} ...")

    formatos_soportados = [".xlsx", ".xls", ".xlsm", ".csv", ".tsv", ".ods",
                           ".sav", ".dta", ".rds", ".rdata"]
    if ext not in formatos_soportados:
        print(f"ERROR: Formato no soportado: {ext}")
        print(f"Formatos soportados: {', '.join(formatos_soportados)}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

    etiquetas_col = {}

    try:
        # ── Excel ──────────────────────────────────────
        if ext in [".xlsx", ".xls", ".xlsm"]:
            xl = pd.ExcelFile(ruta)
            mejor = None
            for hoja in xl.sheet_names:
                try:
                    tmp = xl.parse(hoja)
                    if mejor is None or len(tmp) > len(mejor):
                        mejor = tmp
                except Exception:
                    pass
            if mejor is None:
                raise ValueError("No se pudo leer ninguna hoja.")
            df = mejor

        # ── ODS ────────────────────────────────────────
        elif ext == ".ods":
            try:
                df = pd.read_excel(ruta, engine="odf")
            except Exception as e:
                print(f"ERROR al leer .ods: {e}")
                print("Asegurate de tener instalado: pip install odfpy")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

        # ── CSV / TSV ──────────────────────────────────
        elif ext in [".csv", ".tsv"]:
            encoding = detectar_encoding(ruta)
            print(f"  -> Encoding detectado: {encoding}")
            if ext == ".tsv":
                sep = "\t"
            else:
                sep = detectar_delimitador(ruta, encoding)
                nombres_sep = {",": "coma", ";": "punto y coma", "\t": "tabulador", "|": "pipe"}
                print(f"  -> Delimitador detectado: {nombres_sep.get(sep, sep)}")

            df = pd.read_csv(ruta, sep=sep, encoding=encoding,
                             on_bad_lines="skip", low_memory=False)
            if df.shape[1] == 1:
                for sep_alt in [",", ";", "\t", "|"]:
                    if sep_alt != sep:
                        try:
                            df2 = pd.read_csv(ruta, sep=sep_alt, encoding=encoding,
                                              on_bad_lines="skip", low_memory=False)
                            if df2.shape[1] > df.shape[1]:
                                df = df2
                                sep = sep_alt
                        except Exception:
                            continue

        # ── SPSS (.sav) ────────────────────────────────
        elif ext == ".sav":
            import pyreadstat
            print("  -> Formato SPSS (.sav) detectado")
            df, meta = pyreadstat.read_sav(str(ruta), apply_value_formats=False)
            df, etiquetas_col = aplicar_etiquetas(df, meta)
            if etiquetas_col:
                print(f"  -> {len(etiquetas_col)} etiquetas de variables cargadas")

        # ── Stata (.dta) ───────────────────────────────
        elif ext == ".dta":
            import pyreadstat
            print("  -> Formato Stata (.dta) detectado")
            df, meta = pyreadstat.read_dta(str(ruta))
            df, etiquetas_col = aplicar_etiquetas(df, meta)
            if etiquetas_col:
                print(f"  -> {len(etiquetas_col)} etiquetas de variables cargadas")

        # ── R (.rds / .RData) ──────────────────────────
        elif ext in [".rds", ".rdata"]:
            try:
                import pyreadr
            except ImportError:
                print("ERROR: Falta la libreria pyreadr.")
                print("Ejecuta: pip install pyreadr")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

            print(f"  -> Formato R ({ext}) detectado")
            resultado = pyreadr.read_r(str(ruta))
            # read_r devuelve un dict de dataframes; tomamos el más grande
            if not resultado:
                raise ValueError("El archivo R no contiene dataframes.")
            df = max(resultado.values(), key=len)
            print(f"  -> Objeto R cargado ({list(resultado.keys())[0]})")

        else:
            print(f"ERROR: Formato no soportado: {ext}")
            input("Presiona Enter para cerrar...")
            sys.exit(1)

        # ── Post-proceso común ─────────────────────────
        df.columns = [limpiar_nombre(c) for c in df.columns]
        df = df.dropna(how="all").reset_index(drop=True)

        # Convertir columnas fecha automáticamente
        for col in df.columns:
            if es_texto_col(df[col]) and not es_numerica(df[col]):
                muestra = df[col].dropna().head(20).astype(str)
                hits = muestra.str.match(r'\d{1,4}[-/]\d{1,2}[-/]\d{1,4}').sum()
                if hits > len(muestra) * 0.5:
                    df[col] = pd.to_datetime(df[col], errors="coerce")

        print(f"  -> {len(df)} filas, {len(df.columns)} columnas cargadas.")
        return df, ruta.stem, etiquetas_col

    except Exception as e:
        print(f"ERROR al leer el archivo: {e}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

# ─────────────────────────────────────────────
# ANÁLISIS INTELIGENTE
# ─────────────────────────────────────────────

def analizar(df, etiquetas_col=None):
    if etiquetas_col is None:
        etiquetas_col = {}

    cols = list(df.columns)
    total_filas = len(df)

    col_fecha = None
    cols_num = []
    cols_cat = []

    for c in cols:
        s = df[c].dropna()
        if len(s) == 0:
            continue
        if es_fecha(s):
            if col_fecha is None:
                col_fecha = c
        elif es_numerica(s):
            cols_num.append(c)
        elif es_categorica(s):
            cols_cat.append(c)

    # Nombre visible de columna (etiqueta si existe)
    def nombre_visible(c):
        return etiquetas_col.get(c, c)

    # ── KPIs ──────────────────────────────────────────
    kpis = [{"icon": "📊", "label": "Total Registros", "value": str(total_filas), "clase": "total"}]

    for c in cols_num[:3]:
        s = df[c].dropna()
        if len(s) == 0:
            continue
        val = s.sum()
        label = f"Total {nombre_visible(c)}"
        fmt = f"{val:,.0f}" if val == int(val) else f"{val:,.2f}"
        kpis.append({"icon": "🔢", "label": label, "value": fmt, "clase": "num"})
        prom = s.mean()
        kpis.append({"icon": "📈", "label": f"Promedio {nombre_visible(c)}", "value": f"{prom:,.1f}", "clase": "promedio"})

    # ── Gráficas ──────────────────────────────────────
    graficas = []

    if col_fecha:
        col_val = cols_num[0] if cols_num else None
        data_ts = serie_temporal(df, col_fecha, col_val)
        if len(data_ts["labels"]) >= 2:
            graficas.append({
                "id": "chartFecha",
                "titulo": f"Evolución por Período ({nombre_visible(col_fecha)})",
                "tipo": "line",
                "labels": data_ts["labels"],
                "datasets": [{"label": nombre_visible(col_val) if col_val else "Registros",
                               "data": data_ts["values"]}],
                "ancho": "half"
            })

    tipo_ciclo = ["doughnut", "bar", "doughnut", "bar", "bar", "bar"]
    for i, c in enumerate(cols_cat[:6]):
        data_cat = top_n(df[c], n=8 if i > 0 else 5)
        if len(data_cat["labels"]) < 2:
            continue
        t = tipo_ciclo[i % len(tipo_ciclo)]
        graficas.append({
            "id": f"chartCat{i}",
            "titulo": f"Distribución por {nombre_visible(c)}",
            "tipo": t,
            "labels": data_cat["labels"],
            "datasets": [{"label": nombre_visible(c), "data": data_cat["values"]}],
            "ancho": "half"
        })

    if cols_num and cols_cat:
        try:
            c_cat = cols_cat[0]
            c_num = cols_num[0]
            agr = df.groupby(c_cat)[c_num].sum().nlargest(10).reset_index()
            if len(agr) >= 3:
                graficas.append({
                    "id": "chartTop",
                    "titulo": f"Top 10: {nombre_visible(c_cat)} por {nombre_visible(c_num)}",
                    "tipo": "bar",
                    "labels": [str(x) for x in agr[c_cat].tolist()],
                    "datasets": [{"label": nombre_visible(c_num),
                                   "data": [float(x) for x in agr[c_num].tolist()]}],
                    "ancho": "full"
                })
        except Exception:
            pass

    # ── Tabla resumen ─────────────────────────────────
    tabla = None
    if cols_cat:
        c = cols_cat[0]
        vc = df[c].value_counts().head(12)
        total = vc.sum()
        filas = []
        for val, cnt in vc.items():
            pct = cnt / total * 100
            estado = "Alto" if pct >= 10 else ("Medio" if pct >= 5 else "Bajo")
            filas.append({"nombre": str(val), "cantidad": int(cnt),
                          "porcentaje": f"{pct:.1f}%", "estado": estado})
        tabla = {"columna": nombre_visible(c), "filas": filas}

    # ── Estadísticas descriptivas ─────────────────────
    tabla_stats = None
    if cols_num:
        filas_s = []
        for c in cols_num[:10]:
            s = df[c].dropna()
            if len(s) == 0:
                continue
            q1 = s.quantile(0.25)
            q3 = s.quantile(0.75)
            iqr = q3 - q1
            outliers = int(((s < q1 - 1.5 * iqr) | (s > q3 + 1.5 * iqr)).sum())
            nulos = int(df[c].isna().sum())
            filas_s.append({
                "columna": nombre_visible(c),
                "media": f"{s.mean():,.2f}",
                "mediana": f"{s.median():,.2f}",
                "desv": f"{s.std():,.2f}",
                "min": f"{s.min():,.2f}",
                "max": f"{s.max():,.2f}",
                "nulos": str(nulos),
                "outliers": str(outliers)
            })
        tabla_stats = {"filas": filas_s}

    return {
        "titulo": "",
        "subtitulo": f"Análisis automático • {total_filas:,} registros • {len(df.columns)} columnas",
        "kpis": kpis[:8],
        "graficas": graficas,
        "tabla": tabla,
        "tabla_stats": tabla_stats,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "formato_detectado": {
            "col_fecha": col_fecha,
            "cols_numericas": cols_num,
            "cols_categoricas": cols_cat
        }
    }

# ─────────────────────────────────────────────
# GENERACIÓN HTML
# ─────────────────────────────────────────────

COLORS = [
    "#00b4d8","#f72585","#7209b7","#4361ee",
    "#fb8500","#06d6a0","#ffd60a","#ef476f",
    "#3a86ff","#8338ec"
]

def color_datasets(datasets, tipo):
    result = []
    for i, ds in enumerate(datasets):
        c = COLORS[i % len(COLORS)]
        entry = dict(ds)
        if tipo == "line":
            entry["borderColor"] = c
            entry["_fillColor"] = c
            entry["borderWidth"] = 3
            entry["fill"] = True
            entry["tension"] = 0.4
            entry["pointBackgroundColor"] = c
            entry["pointBorderColor"] = "#fff"
            entry["pointBorderWidth"] = 2
            entry["pointRadius"] = 6
        elif tipo in ["doughnut", "pie"]:
            entry["backgroundColor"] = COLORS[:len(ds["data"])]
            entry["borderWidth"] = 0
        else:
            if len(datasets) == 1:
                entry["backgroundColor"] = COLORS[:len(ds["data"])]
            else:
                entry["backgroundColor"] = c
            entry["borderRadius"] = 8
        result.append(entry)
    return result

def grafica_js(g):
    tipo = g["tipo"]
    datasets = color_datasets(g["datasets"], tipo)

    if tipo == "line":
        opts = """{
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
                x: { grid: { display: false } }
            }
        }"""
    elif tipo in ["doughnut", "pie"]:
        opts = """{
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { padding: 20 } } },
            cutout: '65%'
        }"""
    else:
        max_label = max((len(str(l)) for l in g["labels"]), default=0)
        horiz = len(g["labels"]) > 4 or max_label > 12
        axis_extra = "indexAxis: 'y'," if horiz else ""
        multi_legend = "position: 'bottom', labels: { padding: 20 }" if len(datasets) > 1 else "display: false"
        opts = f"""{{
            responsive: true, maintainAspectRatio: false,
            {axis_extra}
            plugins: {{ legend: {{ {multi_legend} }} }},
            scales: {{
                x: {{ beginAtZero: true, grid: {{ color: 'rgba(255,255,255,0.05)' }} }},
                y: {{ grid: {{ display: false }} }}
            }}
        }}"""

    ds_json = json.dumps(datasets, ensure_ascii=False)
    return f"""
    (function() {{
        var ctx = document.getElementById('{g["id"]}').getContext('2d');
        new Chart(ctx, {{
            type: '{tipo}',
            data: {{
                labels: {json.dumps(g["labels"], ensure_ascii=False)},
                datasets: {ds_json}
            }},
            options: {opts}
        }});
    }})();"""

def generar_html(datos, nombre_archivo):
    titulo = datos["titulo"] or nombre_archivo.replace("_", " ").replace("-", " ").title()

    # KPIs
    kpi_html = ""
    for k in datos["kpis"]:
        cls = k.get("clase", "total")
        kpi_html += f"""
        <div class="kpi-card {cls}">
            <div class="icon">{k["icon"]}</div>
            <div class="value">{k["value"]}</div>
            <div class="label">{k["label"]}</div>
        </div>"""

    # Gráficas
    charts_html = ""
    charts_js = ""
    for g in datos["graficas"]:
        ancho = g.get("ancho", "half")
        tall = ' tall' if ancho == "full" else ""
        charts_html += f"""
        <div class="chart-card {ancho}-width">
            <h3>{g["titulo"]}</h3>
            <div class="chart-container{tall}">
                <canvas id="{g["id"]}"></canvas>
            </div>
        </div>"""
        charts_js += grafica_js(g)

    # Tabla resumen
    tabla_html = ""
    if datos.get("tabla"):
        t = datos["tabla"]
        filas_html = ""
        for f in t["filas"]:
            badge = "badge-success" if f["estado"] == "Alto" else ("badge-warning" if f["estado"] == "Medio" else "badge-danger")
            filas_html += f"""
                    <tr>
                        <td>{f["nombre"]}</td>
                        <td>{f["cantidad"]}</td>
                        <td>{f["porcentaje"]}</td>
                        <td><span class="badge {badge}">{f["estado"]}</span></td>
                    </tr>"""
        tabla_html = f"""
        <section class="chart-card" style="margin-top:25px">
            <h3>📋 Distribución por {t["columna"]}</h3>
            <table class="stats-table">
                <thead>
                    <tr><th>{t["columna"]}</th><th>Cantidad</th><th>Porcentaje</th><th>Nivel</th></tr>
                </thead>
                <tbody>{filas_html}</tbody>
            </table>
        </section>"""

    # Tabla estadísticas descriptivas
    tabla_stats_html = ""
    if datos.get("tabla_stats"):
        filas_s = ""
        for f in datos["tabla_stats"]["filas"]:
            outlier_badge = f'<span class="badge badge-warning">{f["outliers"]}</span>' if int(f["outliers"]) > 0 else f'<span class="badge badge-success">0</span>'
            nulos_badge = f'<span class="badge badge-danger">{f["nulos"]}</span>' if int(f["nulos"]) > 0 else f'<span class="badge badge-success">0</span>'
            filas_s += f"""
                    <tr>
                        <td><strong>{f["columna"]}</strong></td>
                        <td>{f["media"]}</td>
                        <td>{f["mediana"]}</td>
                        <td>{f["desv"]}</td>
                        <td>{f["min"]}</td>
                        <td>{f["max"]}</td>
                        <td>{nulos_badge}</td>
                        <td>{outlier_badge}</td>
                    </tr>"""
        tabla_stats_html = f"""
        <section class="chart-card" style="margin-top:25px">
            <h3>📐 Estadísticas Descriptivas</h3>
            <div style="overflow-x:auto">
            <table class="stats-table">
                <thead>
                    <tr>
                        <th>Variable</th><th>Media</th><th>Mediana</th>
                        <th>Desv. Est.</th><th>Mín</th><th>Máx</th>
                        <th>Nulos</th><th>Outliers (IQR)</th>
                    </tr>
                </thead>
                <tbody>{filas_s}</tbody>
            </table>
            </div>
        </section>"""

    # Badges de columnas detectadas
    fmt = datos.get("formato_detectado", {})
    badge_cols = ""
    if fmt.get("col_fecha"):
        badge_cols += f'<span class="badge badge-success">📅 Fecha: {fmt["col_fecha"]}</span> '
    if fmt.get("cols_numericas"):
        badge_cols += f'<span class="badge badge-warning">🔢 Numéricas: {len(fmt["cols_numericas"])}</span> '
    if fmt.get("cols_categoricas"):
        badge_cols += f'<span class="badge badge-success">🏷️ Categorías: {len(fmt["cols_categoricas"])}</span> '

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard - {titulo}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <style>
        * {{ margin:0; padding:0; box-sizing:border-box; }}
        :root {{
            --bg-primary:#0f0f23; --bg-secondary:#1a1a2e; --bg-card:#16213e;
            --text-primary:#eaeaea; --text-secondary:#a0a0a0;
            --accent-blue:#4361ee; --accent-purple:#7209b7; --accent-pink:#f72585;
            --accent-orange:#fb8500; --accent-green:#06d6a0; --accent-cyan:#00b4d8;
            --border-color:#2a2a4a;
        }}
        body {{
            font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;
            background:linear-gradient(135deg,var(--bg-primary) 0%,var(--bg-secondary) 100%);
            color:var(--text-primary); min-height:100vh;
        }}
        .container {{ max-width:1600px; margin:0 auto; padding:20px; }}
        header {{ text-align:center; padding:30px 0; border-bottom:1px solid var(--border-color); margin-bottom:20px; }}
        header h1 {{
            font-size:2.5rem;
            background:linear-gradient(90deg,var(--accent-cyan),var(--accent-pink));
            -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text;
            margin-bottom:10px;
        }}
        header p {{ color:var(--text-secondary); font-size:1.1rem; margin-bottom:12px; }}
        .col-badges {{ display:flex; justify-content:center; gap:10px; flex-wrap:wrap; margin-top:10px; }}
        .kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:20px; margin-bottom:30px; }}
        .kpi-card {{
            background:var(--bg-card); border-radius:16px; padding:25px; text-align:center;
            border:1px solid var(--border-color); transition:transform .3s,box-shadow .3s;
        }}
        .kpi-card:hover {{ transform:translateY(-5px); box-shadow:0 10px 30px rgba(67,97,238,.2); }}
        .kpi-card .icon {{ font-size:2.5rem; margin-bottom:10px; }}
        .kpi-card .value {{ font-size:2.2rem; font-weight:bold; margin-bottom:5px; }}
        .kpi-card .label {{ color:var(--text-secondary); font-size:.95rem; }}
        .kpi-card.total    {{ border-left:4px solid var(--accent-cyan); }}
        .kpi-card.activos  {{ border-left:4px solid var(--accent-orange); }}
        .kpi-card.tiempo   {{ border-left:4px solid var(--accent-green); }}
        .kpi-card.retraso  {{ border-left:4px solid var(--accent-pink); }}
        .kpi-card.tasa     {{ border-left:4px solid var(--accent-purple); }}
        .kpi-card.promedio {{ border-left:4px solid var(--accent-blue); }}
        .kpi-card.num      {{ border-left:4px solid var(--accent-cyan); }}
        .charts-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(500px,1fr)); gap:25px; margin-bottom:30px; }}
        .chart-card {{ background:var(--bg-card); border-radius:16px; padding:25px; border:1px solid var(--border-color); }}
        .chart-card h3 {{ color:var(--text-primary); margin-bottom:20px; font-size:1.2rem; display:flex; align-items:center; gap:10px; }}
        .chart-card h3::before {{ content:''; width:4px; height:20px; background:linear-gradient(180deg,var(--accent-cyan),var(--accent-pink)); border-radius:2px; }}
        .chart-container {{ position:relative; height:300px; }}
        .chart-container.tall {{ height:400px; }}
        .full-width {{ grid-column:1/-1; }}
        .half-width {{ grid-column:span 1; }}
        footer {{ text-align:center; padding:20px; margin-top:30px; border-top:1px solid var(--border-color); color:var(--text-secondary); font-size:.9rem; }}
        .stats-table {{ width:100%; margin-top:15px; border-collapse:collapse; font-size:.9rem; }}
        .stats-table th,.stats-table td {{ padding:12px 15px; text-align:left; border-bottom:1px solid var(--border-color); }}
        .stats-table th {{ background:var(--bg-secondary); color:var(--accent-cyan); font-weight:600; }}
        .stats-table tr:hover {{ background:var(--bg-secondary); }}
        .badge {{ display:inline-block; padding:4px 12px; border-radius:20px; font-size:.85rem; font-weight:500; }}
        .badge-success {{ background:rgba(6,214,160,.2); color:var(--accent-green); }}
        .badge-warning {{ background:rgba(251,133,0,.2); color:var(--accent-orange); }}
        .badge-danger  {{ background:rgba(247,37,133,.2); color:var(--accent-pink); }}
        @media(max-width:1100px){{ .charts-grid{{grid-template-columns:1fr;}} .chart-container{{height:280px;}} }}
        @media(max-width:600px){{ .kpi-grid{{grid-template-columns:repeat(2,1fr);}} header h1{{font-size:1.8rem;}} }}
    </style>
</head>
<body>
<div class="container">
    <header>
        <h1>Dashboard — {titulo}</h1>
        <p>{datos["subtitulo"]} &nbsp;|&nbsp; Generado: {datos["generado"]}</p>
        <div class="col-badges">{badge_cols}</div>
    </header>

    <section class="kpi-grid">
{kpi_html}
    </section>

    <section class="charts-grid">
{charts_html}
    </section>

{tabla_html}
{tabla_stats_html}

    <footer>
        <p>Estadincho-Gen v3.0 &nbsp;|&nbsp; Archivo: {nombre_archivo} &nbsp;|&nbsp; {datos["generado"]}</p>
    </footer>
</div>
<script>
    Chart.defaults.color = '#a0a0a0';
    Chart.defaults.borderColor = '#2a2a4a';
{charts_js}
</script>
</body>
</html>"""
    return html

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("=" * 55)
        print("  ESTADINCHO-GEN v3.0")
        print("=" * 55)
        print()
        print("Arrastra tu archivo sobre este script, o ejecuta:")
        print()
        print("  python generar_dashboard.py mi_archivo.xlsx")
        print("  python generar_dashboard.py datos.csv")
        print("  python generar_dashboard.py datos.tsv")
        print("  python generar_dashboard.py datos.ods")
        print("  python generar_dashboard.py encuesta.sav")
        print("  python generar_dashboard.py resultados.dta")
        print("  python generar_dashboard.py objeto.rds")
        print()
        input("Presiona Enter para cerrar...")
        sys.exit(0)

    ruta_archivo = sys.argv[1]
    df, nombre, etiquetas_col = cargar_archivo(ruta_archivo)

    print("Analizando datos...")
    datos = analizar(df, etiquetas_col)

    print("Generando HTML...")
    html = generar_html(datos, nombre)

    carpeta = Path(ruta_archivo).parent
    salida = (carpeta / f"dashboard_{nombre}.html").resolve()
    salida.write_text(html, encoding="utf-8")

    print(f"\n✔ Dashboard creado: {salida}")
    print("  Abriendo en el navegador...")

    import subprocess
    ruta_str = str(salida)
    abierto = False

    for chrome in [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]:
        if os.path.exists(chrome):
            subprocess.Popen([chrome, ruta_str])
            abierto = True
            break

    if not abierto:
        for firefox in [
            r"C:\Program Files\Mozilla Firefox\firefox.exe",
            r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
        ]:
            if os.path.exists(firefox):
                subprocess.Popen([firefox, ruta_str])
                abierto = True
                break

    if not abierto:
        os.startfile(ruta_str)

    print("\nListo.")

if __name__ == "__main__":
    main()
