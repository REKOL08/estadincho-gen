# 📊 Estadincho-Gen

**Convierte cualquier archivo de datos en un dashboard visual interactivo con un solo clic.**

Desarrollado para el sistema de bibliotecas de la Fundación Universitaria del Área Andina — pero funciona con cualquier dataset.

---

## ¿Qué hace?

Arrastra un archivo de datos sobre el ejecutable y genera automáticamente un dashboard HTML completo con:

- **KPIs** — total de registros, sumas y promedios de variables numéricas, score de calidad
- **Gráficas** — línea de tiempo, distribuciones por categoría, top 10, dona
- **Tabla de resumen** — distribución por la variable categórica principal
- **Calidad de datos** — nulos, outliers por IQR y Z-score, score 0–100 por variable
- **Matriz de correlación** — heatmap en Canvas puro, sin librerías externas

Todo en un solo archivo `.html` que se abre directamente en el navegador. **Sin instalaciones, sin configuración.**

---

## Uso

### Para el usuario final — solo dos archivos necesarios

generar_dashboard.exe
generar_dashboard.bat

**Opción 1 — Arrastrar y soltar**
Arrastra tu archivo de datos directamente sobre `generar_dashboard.bat` en el Explorador de Windows.

**Opción 2 — Doble clic**
Ejecuta `generar_dashboard.bat`, escribe la ruta del archivo cuando se solicite.

El dashboard se guarda en la misma carpeta del archivo original como `dashboard_<nombre>.html` y se abre automáticamente en el navegador.

---

## Formatos soportados

| Formato | Extensión |
|---|---|
| Excel | `.xlsx` `.xls` `.xlsm` |
| CSV | `.csv` |
| TSV | `.tsv` |
| OpenDocument | `.ods` |
| SPSS | `.sav` |
| Stata | `.dta` |
| R | `.rds` `.RData` |

---

## Estructura del proyecto

estadincho-gen/
├── generar_dashboard.exe  # Ejecutable — distribuir al usuario final
├── generar_dashboard.bat  # Lanzador por arrastre
├── generar_dashboard.py   # Código fuente
└── README.md

---

## Fases de desarrollo

| Fase | Descripción | Estado |
|---|---|---|
| Base | Excel y CSV, gráficas automáticas, KPIs | ✅ |
| Fase 1 | Soporte TSV y ODS, encoding/delimitador automático, tabla de estadísticas | ✅ |
| Fase 2 | Soporte SPSS, Stata y R con etiquetas de variables | ✅ |
| Fase 3 | Calidad de datos por variable + matriz de correlación | ✅ |

---

## Desarrollado por

**REKOL08** — Biblioteca Fundación Universitaria del Área Andina  
Sedes: Bogotá · Pereira · Valledupar