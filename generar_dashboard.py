#!/usr/bin/env python3
"""
Estadincho-Gen v4.0 - Convierte cualquier archivo de datos en un dashboard visual.
Formatos soportados: .xlsx, .xls, .xlsm, .csv, .tsv, .ods, .sav, .dta, .rds, .RData
Uso: python generar_dashboard.py archivo.xlsx
     python generar_dashboard.py datos.csv
"""

import sys
import os
import json
import webbrowser
import re
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: Falta la libreria pandas.")
    print("Ejecuta: pip install pandas openpyxl")
    input("Presiona Enter para cerrar...")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────────────────────────

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
    return {"labels": [str(x) for x in vc.index.tolist()],
            "values": [int(x) for x in vc.values.tolist()]}

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

# ─────────────────────────────────────────────────────────────────
# CARGA DEL ARCHIVO  (Fase 1 + Fase 2)
# ─────────────────────────────────────────────────────────────────

def cargar_archivo(ruta):
    ruta = Path(ruta)
    if not ruta.exists():
        print(f"ERROR: No se encontro el archivo: {ruta}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

    ext = ruta.suffix.lower()
    print(f"Cargando archivo: {ruta.name} ...")

    var_labels = {}   # {col_nombre: etiqueta_larga}
    val_labels = {}   # {col_nombre: {codigo: etiqueta}}

    try:
        # ── Excel ──────────────────────────────────────────────
        if ext in [".xlsx", ".xls", ".xlsm"]:
            xl = pd.ExcelFile(ruta)
            mejor = None
            for hoja in xl.sheet_names:
                try:
                    tmp = xl.parse(hoja)
                    if mejor is None or len(tmp) > len(mejor):
                        mejor = tmp
                except:
                    pass
            if mejor is None:
                raise ValueError("No se pudo leer ninguna hoja.")
            df = mejor

        # ── ODS ────────────────────────────────────────────────
        elif ext == ".ods":
            df = pd.read_excel(ruta, engine="odf")

        # ── CSV ────────────────────────────────────────────────
        elif ext == ".csv":
            df = None
            for sep in [",", ";", "\t", "|"]:
                for enc in ["utf-8", "latin-1"]:
                    try:
                        tmp = pd.read_csv(ruta, sep=sep, encoding=enc, on_bad_lines="skip")
                        if tmp.shape[1] > 1:
                            df = tmp
                            break
                    except:
                        continue
                if df is not None:
                    break
            if df is None:
                raise ValueError("No se pudo detectar el separador del CSV.")

        # ── TSV ────────────────────────────────────────────────
        elif ext == ".tsv":
            df = None
            for enc in ["utf-8", "latin-1"]:
                try:
                    df = pd.read_csv(ruta, sep="\t", encoding=enc, on_bad_lines="skip")
                    break
                except:
                    continue
            if df is None:
                raise ValueError("No se pudo leer el archivo TSV.")

        # ── SPSS (.sav) ────────────────────────────────────────
        elif ext == ".sav":
            try:
                import pyreadstat
                df, meta = pyreadstat.read_sav(str(ruta), apply_value_formats=False)
                var_labels = meta.column_labels_and_names  # {nombre: etiqueta}
                val_labels = meta.variable_value_labels    # {nombre: {cod: etiqueta}}
                # Reemplazar códigos por etiquetas donde aplique
                for col, mapping in val_labels.items():
                    if col in df.columns:
                        df[col] = df[col].map(lambda x: mapping.get(x, x))
            except ImportError:
                print("AVISO: pyreadstat no instalado. Ejecuta: pip install pyreadstat")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

        # ── Stata (.dta) ───────────────────────────────────────
        elif ext == ".dta":
            try:
                import pyreadstat
                df, meta = pyreadstat.read_dta(str(ruta), apply_value_formats=False)
                var_labels = meta.column_labels_and_names
                val_labels = meta.variable_value_labels
                for col, mapping in val_labels.items():
                    if col in df.columns:
                        df[col] = df[col].map(lambda x: mapping.get(x, x))
            except ImportError:
                print("AVISO: pyreadstat no instalado. Ejecuta: pip install pyreadstat")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

        # ── R (.rds / .RData) ──────────────────────────────────
        elif ext in [".rds", ".rdata"]:
            try:
                import pyreadr
                result = pyreadr.read_r(str(ruta))
                # Tomar el DataFrame más grande
                df = max(result.values(), key=lambda x: len(x) if hasattr(x, '__len__') else 0)
            except ImportError:
                print("AVISO: pyreadr no instalado. Ejecuta: pip install pyreadr")
                input("Presiona Enter para cerrar...")
                sys.exit(1)

        else:
            print(f"ERROR: Formato no soportado: {ext}")
            print("Formatos soportados: .xlsx .xls .xlsm .csv .tsv .ods .sav .dta .rds .RData")
            input("Presiona Enter para cerrar...")
            sys.exit(1)

        # Aplicar etiquetas largas como nombres de columna si las hay
        if var_labels:
            rename_map = {}
            for col in df.columns:
                etiqueta = var_labels.get(col, "")
                if etiqueta and etiqueta != col:
                    rename_map[col] = etiqueta
            if rename_map:
                df = df.rename(columns=rename_map)

        df.columns = [limpiar_nombre(c) for c in df.columns]
        df = df.dropna(how="all").reset_index(drop=True)
        print(f"  -> {len(df)} filas, {len(df.columns)} columnas cargadas.")
        return df, ruta.stem

    except Exception as e:
        print(f"ERROR al leer el archivo: {e}")
        input("Presiona Enter para cerrar...")
        sys.exit(1)

# ─────────────────────────────────────────────────────────────────
# FASE 3 — CALIDAD DE DATOS
# ─────────────────────────────────────────────────────────────────

def calcular_calidad(df):
    """
    Retorna:
      - filas_calidad: lista de dicts por columna numérica
      - score_global: int 0-100
    """
    import math

    cols_num = [c for c in df.columns if es_numerica(df[c])]
    if not cols_num:
        return [], 100

    total = len(df)
    filas = []
    scores = []

    for col in cols_num:
        s = df[col]
        nulos = int(s.isna().sum())
        pct_nulos = nulos / total * 100

        s_clean = s.dropna()
        n = len(s_clean)

        # Outliers IQR
        if n >= 4:
            q1 = s_clean.quantile(0.25)
            q3 = s_clean.quantile(0.75)
            iqr = q3 - q1
            out_iqr = int(((s_clean < q1 - 1.5 * iqr) | (s_clean > q3 + 1.5 * iqr)).sum())
        else:
            out_iqr = 0

        # Outliers Z-score
        if n >= 4:
            mean = s_clean.mean()
            std  = s_clean.std()
            if std > 0:
                zscores = ((s_clean - mean) / std).abs()
                out_z = int((zscores > 3).sum())
            else:
                out_z = 0
        else:
            out_z = 0

        # Score 0-100
        penalizacion = pct_nulos * 0.7 + (out_iqr / max(n, 1)) * 100 * 0.3
        score = max(0, round(100 - penalizacion))
        scores.append(score)

        if score >= 80:
            estado = "Buena"
        elif score >= 50:
            estado = "Regular"
        else:
            estado = "Revisar"

        filas.append({
            "col": col,
            "nulos": nulos,
            "pct_nulos": f"{pct_nulos:.1f}%",
            "out_iqr": out_iqr,
            "out_z": out_z,
            "score": score,
            "estado": estado
        })

    score_global = round(sum(scores) / len(scores)) if scores else 100
    return filas, score_global

# ─────────────────────────────────────────────────────────────────
# FASE 3 — MATRIZ DE CORRELACIÓN
# ─────────────────────────────────────────────────────────────────

def calcular_correlacion(df):
    """
    Retorna dict con labels (columnas) y matriz 2D de valores redondeados,
    o None si hay menos de 2 columnas numéricas.
    """
    cols_num = [c for c in df.columns if es_numerica(df[c]) and df[c].dropna().nunique() > 1]
    if len(cols_num) < 2:
        return None

    cols_num = cols_num[:10]  # máx 10x10
    corr = df[cols_num].corr().round(2)
    matrix = []
    for _, row in corr.iterrows():
        matrix.append([None if pd.isna(v) else float(v) for v in row])

    return {
        "labels": cols_num,
        "matrix": matrix
    }

# ─────────────────────────────────────────────────────────────────
# ANÁLISIS INTELIGENTE
# ─────────────────────────────────────────────────────────────────

def analizar(df):
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

    # Calidad y correlación (Fase 3)
    filas_calidad, score_global = calcular_calidad(df)
    correlacion = calcular_correlacion(df)

    # ── KPIs ────────────────────────────────────────────────────
    icono_score = "🟢" if score_global >= 80 else ("🟡" if score_global >= 50 else "🔴")
    kpis = [
        {"icon": "📊", "label": "Total Registros", "value": str(total_filas), "clase": "total"},
        {"icon": icono_score, "label": "Calidad de Datos", "value": f"{score_global}/100", "clase": "calidad"}
    ]

    for c in cols_num[:3]:
        s = df[c].dropna()
        if len(s) == 0:
            continue
        val = s.sum()
        fmt = f"{val:,.0f}" if val == int(val) else f"{val:,.2f}"
        kpis.append({"icon": "💰", "label": f"Total {c}", "value": fmt, "clase": "num"})
        prom = s.mean()
        kpis.append({"icon": "📈", "label": f"Promedio {c}", "value": f"{prom:,.1f}", "clase": "promedio"})

    # ── Gráficas ─────────────────────────────────────────────────
    graficas = []

    if col_fecha:
        col_val = cols_num[0] if cols_num else None
        data_ts = serie_temporal(df, col_fecha, col_val)
        if len(data_ts["labels"]) >= 2:
            graficas.append({
                "id": "chartFecha",
                "titulo": f"Evolución por Período ({col_fecha})",
                "tipo": "line",
                "labels": data_ts["labels"],
                "datasets": [{"label": col_val or "Registros", "data": data_ts["values"]}],
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
            "titulo": f"Distribución por {c}",
            "tipo": t,
            "labels": data_cat["labels"],
            "datasets": [{"label": c, "data": data_cat["values"]}],
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
                    "titulo": f"Top 10: {c_cat} por {c_num}",
                    "tipo": "bar",
                    "labels": [str(x) for x in agr[c_cat].tolist()],
                    "datasets": [{"label": c_num, "data": [float(x) for x in agr[c_num].tolist()]}],
                    "ancho": "full"
                })
        except:
            pass

    # ── Tabla resumen ─────────────────────────────────────────────
    tabla = None
    if cols_cat:
        c = cols_cat[0]
        vc = df[c].value_counts().head(12)
        total = vc.sum()
        filas = []
        for val, cnt in vc.items():
            pct = cnt / total * 100
            estado = "Alto" if pct >= 10 else ("Medio" if pct >= 5 else "Bajo")
            filas.append({
                "nombre": str(val),
                "cantidad": int(cnt),
                "porcentaje": f"{pct:.1f}%",
                "estado": estado
            })
        tabla = {"columna": c, "filas": filas}

    return {
        "titulo": "",
        "subtitulo": f"Análisis automático · {total_filas:,} registros · {len(df.columns)} columnas",
        "kpis": kpis[:8],
        "graficas": graficas,
        "tabla": tabla,
        "calidad": filas_calidad,
        "correlacion": correlacion,
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M")
    }

# ─────────────────────────────────────────────────────────────────
# GENERACIÓN HTML
# ─────────────────────────────────────────────────────────────────

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

# ─── Tabla de calidad HTML ────────────────────────────────────────
def html_calidad(filas_calidad):
    if not filas_calidad:
        return ""
    filas_html = ""
    for f in filas_calidad:
        badge = ("badge-success" if f["estado"] == "Buena"
                 else "badge-warning" if f["estado"] == "Regular"
                 else "badge-danger")
        bar_color = ("#06d6a0" if f["score"] >= 80
                     else "#fb8500" if f["score"] >= 50
                     else "#f72585")
        filas_html += f"""
            <tr>
                <td>{f["col"]}</td>
                <td>{f["nulos"]} ({f["pct_nulos"]})</td>
                <td>{f["out_iqr"]}</td>
                <td>{f["out_z"]}</td>
                <td>
                    <div class="score-bar-bg">
                        <div class="score-bar" style="width:{f["score"]}%;background:{bar_color}"></div>
                    </div>
                    <span style="font-size:.85rem;color:{bar_color}">{f["score"]}</span>
                </td>
                <td><span class="badge {badge}">{f["estado"]}</span></td>
            </tr>"""
    return f"""
    <section class="chart-card full-width" style="margin-top:25px">
        <h3>Calidad de Datos por Variable</h3>
        <table class="stats-table">
            <thead>
                <tr>
                    <th>Variable</th>
                    <th>Nulos</th>
                    <th>Outliers IQR</th>
                    <th>Outliers Z-score</th>
                    <th>Score</th>
                    <th>Estado</th>
                </tr>
            </thead>
            <tbody>{filas_html}</tbody>
        </table>
    </section>"""

# ─── Canvas de correlación HTML + JS ─────────────────────────────
def html_correlacion(correlacion):
    if not correlacion:
        return "", ""
    labels_json = json.dumps(correlacion["labels"], ensure_ascii=False)
    matrix_json = json.dumps(correlacion["matrix"])
    n = len(correlacion["labels"])

    html = """
    <section class="chart-card full-width" style="margin-top:25px">
        <h3>Matriz de Correlación</h3>
        <div style="overflow-x:auto">
            <canvas id="canvasCorr"></canvas>
        </div>
    </section>"""

    js = f"""
    (function() {{
        var labels  = {labels_json};
        var matrix  = {matrix_json};
        var n       = labels.length;
        var cell    = 70;
        var pad     = 130;
        var cvs     = document.getElementById('canvasCorr');
        cvs.width   = pad + n * cell;
        cvs.height  = pad + n * cell;
        var ctx     = cvs.getContext('2d');

        function lerp(a, b, t) {{ return a + (b - a) * t; }}
        function corrColor(v) {{
            if (v === null) return '#2a2a4a';
            var t = (v + 1) / 2;        // 0=rojo, 0.5=gris, 1=azul
            var r, g, b;
            if (t < 0.5) {{
                r = Math.round(lerp(220,  50, t * 2));
                g = Math.round(lerp( 50,  50, t * 2));
                b = Math.round(lerp( 50, 120, t * 2));
            }} else {{
                r = Math.round(lerp( 50,  30, (t-0.5)*2));
                g = Math.round(lerp( 50,  80, (t-0.5)*2));
                b = Math.round(lerp(120, 220, (t-0.5)*2));
            }}
            return 'rgb(' + r + ',' + g + ',' + b + ')';
        }}

        ctx.fillStyle = '#0f0f23';
        ctx.fillRect(0, 0, cvs.width, cvs.height);

        // Etiquetas columnas (arriba)
        ctx.fillStyle = '#a0a0a0';
        ctx.font      = '11px Segoe UI';
        ctx.textAlign = 'center';
        for (var j = 0; j < n; j++) {{
            ctx.save();
            ctx.translate(pad + j * cell + cell/2, pad - 10);
            ctx.rotate(-Math.PI / 4);
            ctx.fillText(labels[j].substring(0, 14), 0, 0);
            ctx.restore();
        }}
        // Etiquetas filas (izquierda)
        ctx.textAlign = 'right';
        for (var i = 0; i < n; i++) {{
            ctx.fillStyle = '#a0a0a0';
            ctx.fillText(labels[i].substring(0, 16), pad - 8, pad + i * cell + cell/2 + 4);
        }}

        // Celdas
        ctx.textAlign = 'center';
        for (var i = 0; i < n; i++) {{
            for (var j = 0; j < n; j++) {{
                var v  = matrix[i][j];
                var x  = pad + j * cell;
                var y  = pad + i * cell;
                ctx.fillStyle = corrColor(v);
                ctx.beginPath();
                ctx.roundRect(x+2, y+2, cell-4, cell-4, 6);
                ctx.fill();
                if (v !== null) {{
                    ctx.fillStyle = (Math.abs(v) > 0.5) ? '#ffffff' : '#cccccc';
                    ctx.font = 'bold 12px Segoe UI';
                    ctx.fillText(v.toFixed(2), x + cell/2, y + cell/2 + 4);
                }}
            }}
        }}
    }})();"""
    return html, js

def generar_html(datos, nombre_archivo):
    titulo = datos["titulo"] or nombre_archivo.replace("_", " ").replace("-", " ").title()

    # KPIs
    clases_kpi = ["total","calidad","num","promedio","activos","tiempo","retraso","tasa"]
    kpi_html = ""
    for i, k in enumerate(datos["kpis"]):
        cls = k.get("clase", clases_kpi[i % len(clases_kpi)])
        kpi_html += f"""
        <div class="kpi-card {cls}">
            <div class="icon">{k["icon"]}</div>
            <div class="value">{k["value"]}</div>
            <div class="label">{k["label"]}</div>
        </div>"""

    # Gráficas
    charts_html = ""
    charts_js   = ""
    for g in datos["graficas"]:
        ancho = g.get("ancho", "half")
        tall  = ' tall' if ancho == "full" else ""
        charts_html += f"""
        <div class="chart-card {ancho}-width">
            <h3>{g["titulo"]}</h3>
            <div class="chart-container{tall}">
                <canvas id="{g["id"]}"></canvas>
            </div>
        </div>"""
        charts_js += grafica_js(g)

    # Tabla resumen categorías
    tabla_html = ""
    if datos.get("tabla"):
        t = datos["tabla"]
        filas_html = ""
        for f in t["filas"]:
            badge = ("badge-success" if f["estado"] == "Alto"
                     else "badge-warning" if f["estado"] == "Medio"
                     else "badge-danger")
            filas_html += f"""
                    <tr>
                        <td>{f["nombre"]}</td>
                        <td>{f["cantidad"]}</td>
                        <td>{f["porcentaje"]}</td>
                        <td><span class="badge {badge}">{f["estado"]}</span></td>
                    </tr>"""
        tabla_html = f"""
        <section class="chart-card" style="margin-top:25px">
            <h3>Distribución por {t["columna"]}</h3>
            <table class="stats-table">
                <thead>
                    <tr><th>{t["columna"]}</th><th>Cantidad</th><th>Porcentaje</th><th>Nivel</th></tr>
                </thead>
                <tbody>{filas_html}</tbody>
            </table>
        </section>"""

    # Calidad y correlación (Fase 3)
    calidad_html = html_calidad(datos.get("calidad", []))
    corr_html, corr_js = html_correlacion(datos.get("correlacion"))

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
        header {{ text-align:center; padding:30px 0; border-bottom:1px solid var(--border-color); margin-bottom:30px; }}
        header h1 {{
            font-size:2.5rem;
            background:linear-gradient(90deg,var(--accent-cyan),var(--accent-pink));
            -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text;
            margin-bottom:10px;
        }}
        header p {{ color:var(--text-secondary); font-size:1.1rem; }}
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
        .kpi-card.calidad  {{ border-left:4px solid var(--accent-green); }}
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
        .full-width  {{ grid-column:1/-1; }}
        .half-width  {{ grid-column:span 1; }}
        footer {{ text-align:center; padding:20px; margin-top:30px; border-top:1px solid var(--border-color); color:var(--text-secondary); font-size:.9rem; }}
        .stats-table {{ width:100%; margin-top:15px; border-collapse:collapse; }}
        .stats-table th,.stats-table td {{ padding:12px 15px; text-align:left; border-bottom:1px solid var(--border-color); }}
        .stats-table th {{ background:var(--bg-secondary); color:var(--accent-cyan); font-weight:600; }}
        .stats-table tr:hover {{ background:var(--bg-secondary); }}
        .badge {{ display:inline-block; padding:4px 12px; border-radius:20px; font-size:.85rem; font-weight:500; }}
        .badge-success {{ background:rgba(6,214,160,.2);   color:var(--accent-green); }}
        .badge-warning {{ background:rgba(251,133,0,.2);   color:var(--accent-orange); }}
        .badge-danger  {{ background:rgba(247,37,133,.2);  color:var(--accent-pink); }}
        .score-bar-bg  {{ display:inline-block; width:80px; height:8px; background:#2a2a4a; border-radius:4px; vertical-align:middle; margin-right:6px; }}
        .score-bar     {{ height:8px; border-radius:4px; transition:width .5s; }}
        @media(max-width:1100px){{ .charts-grid{{grid-template-columns:1fr;}} .chart-container{{height:280px;}} }}
        @media(max-width:600px){{ .kpi-grid{{grid-template-columns:repeat(2,1fr);}} header h1{{font-size:1.8rem;}} }}
    </style>
</head>
<body>
<div class="container">
    <header>
        <h1>Dashboard — {titulo}</h1>
        <p>{datos["subtitulo"]} &nbsp;|&nbsp; Generado: {datos["generado"]}</p>
    </header>

    <section class="kpi-grid">
{kpi_html}
    </section>

    <section class="charts-grid">
{charts_html}
    </section>

{tabla_html}
{calidad_html}
{corr_html}

    <footer>
        <p>Estadincho-Gen v4.0 &nbsp;|&nbsp; Archivo: {nombre_archivo} &nbsp;|&nbsp; {datos["generado"]}</p>
    </footer>
</div>
<script>
    Chart.defaults.color = '#a0a0a0';
    Chart.defaults.borderColor = '#2a2a4a';
{charts_js}
{corr_js}
</script>
</body>
</html>"""
    return html

# ─────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("=" * 55)
        print("  ESTADINCHO-GEN v4.0")
        print("=" * 55)
        print()
        print("Arrastra tu archivo sobre este script, o ejecuta:")
        print()
        print("  python generar_dashboard.py mi_archivo.xlsx")
        print("  python generar_dashboard.py datos.csv")
        print("  python generar_dashboard.py encuesta.sav")
        print()
        input("Presiona Enter para cerrar...")
        sys.exit(0)

    ruta_archivo = sys.argv[1]
    df, nombre  = cargar_archivo(ruta_archivo)

    print("Analizando datos...")
    datos = analizar(df)

    print("Generando HTML...")
    html = generar_html(datos, nombre)

    carpeta = Path(ruta_archivo).parent
    salida  = (carpeta / f"dashboard_{nombre}.html").resolve()
    salida.write_text(html, encoding="utf-8")

    print(f"\n✔ Dashboard creado: {salida}")
    print("  Abriendo en el navegador...")
    webbrowser.open(salida.as_uri())
    print("\nListo.")

if __name__ == "__main__":
    main()
