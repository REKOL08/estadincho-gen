# 📊 Dashboard Generator

Convierte **cualquier Excel o CSV** en un dashboard visual profesional con un solo clic.  
Sin conocimientos técnicos. Sin configuración. Solo arrastra el archivo.

---

## ✅ Requisitos

- Windows 7 / 10 / 11
- Python 3.8 o superior → [descargar aquí](https://www.python.org/downloads/)  
  ⚠️ Durante la instalación de Python marca **"Add Python to PATH"**

---

## 🚀 Instalación (una sola vez por PC)

### Opción A — GitHub (recomendada para distribuir)

```cmd
git clone https://github.com/TU_USUARIO/dashboard-gen.git
cd dashboard-gen
pip install pandas openpyxl
```

### Opción B — Descarga ZIP

1. Descarga el ZIP del repositorio
2. Extrae en cualquier carpeta
3. Abre CMD en esa carpeta y ejecuta:

```cmd
pip install pandas openpyxl
```

---

## 🖱️ Uso diario (sin CMD)

1. **Arrastra** tu archivo `.xlsx` o `.csv` encima del ícono `generar_dashboard.bat`
2. Espera unos segundos
3. El dashboard se abre **automáticamente** en tu navegador

**Así de simple.** El archivo HTML generado queda en la misma carpeta que tu Excel/CSV.

---

## 💻 Uso desde CMD

```cmd
python generar_dashboard.py ventas_2024.xlsx
python generar_dashboard.py clientes.csv
python generar_dashboard.py C:\Documentos\reporte.xlsx
```

---

## 📁 Formatos soportados

| Formato | Extensión |
|---------|-----------|
| Excel   | `.xlsx` `.xls` `.xlsm` |
| CSV     | `.csv` (comas, punto y coma, tabulaciones) |

---

## 🎨 ¿Qué genera el dashboard?

El script detecta **automáticamente** el tipo de cada columna:

| Tipo de columna | Qué genera |
|----------------|------------|
| Fechas | Gráfica de línea por período |
| Texto / Categorías | Gráfica de dona o barras horizontales |
| Números | KPIs con suma y promedio |
| Combinación | Top 10, comparativas, tabla resumen |

**Estilo:** fondo oscuro, colores neón, Chart.js — idéntico al dashboard de referencia.

---

## ❓ Preguntas frecuentes

**¿El archivo se sube a internet?**  
No. Todo se procesa localmente en tu PC. El dashboard es un HTML que vive en tu carpeta.

**¿Funciona con archivos con varias hojas?**  
Sí, toma automáticamente la hoja con más datos.

**¿El CSV tiene punto y coma como separador?**  
Sí, detecta automáticamente: comas, punto y coma, tabulaciones y pipes (`|`).

**Primera vez tarda mucho.**  
Solo la primera vez instala `pandas` y `openpyxl`. Las siguientes veces es instantáneo.

---

## 🗂️ Archivos del proyecto

```
dashboard-gen/
├── generar_dashboard.py   ← Script principal (cerebro)
├── generar_dashboard.bat  ← Lanzador Windows (drag & drop)
└── README.md              ← Este archivo
```

---

*Dashboard Generator — uso libre interno*
