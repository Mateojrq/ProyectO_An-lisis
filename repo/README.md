# Análisis de la Educación Superior en Ecuador (2015-2023)
### Cobertura, Equidad de Acceso y Eficiencia Académica

Proyecto de fin de semestre — Análisis de Datos, Escuela de Formación de Tecnólogos, Escuela Politécnica Nacional (2026-A).

## Descripción

Integración de 7 fuentes de datos abiertos del Ecuador (SENESCYT y Ministerio de Educación) para diagnosticar cobertura, equidad de acceso y eficiencia académica de la educación superior en Ecuador, usando **Python**, **KNIME** y **Power BI**.

## Estructura del repositorio

```
├── data/
│   ├── raw/            # Datasets originales, sin modificar (los 7 .xlsx/.csv/.pdf descargados)
│   └── processed/       # Datasets ya limpios/unidos, listos para Power BI
├── notebooks/            # Notebooks .ipynb: inspección, EDA, limpieza avanzada, clustering
├── knime/                 # Workflow .knwf del proceso ETL
├── powerbi/               # Archivo .pbix del dashboard
├── sql/                   # Scripts de creación de tablas / carga a SQL y NoSQL
├── scripts/               # Scripts .py auxiliares (estandarización de nombres, validaciones)
├── docs/                  # Diccionario de datos, capturas del workflow y del dashboard
├── requirements.txt       # Dependencias de Python
└── README.md
```

## Fuentes de datos

Todas provienen del [Portal de Datos Abiertos del Ecuador](https://www.datosabiertos.gob.ec/), grupo Educación.

| # | Fuente | Institución | Formato |
|---|--------|-------------|---------|
| 1 | Base estadística de matrícula UEP 2015-2023 | SENESCYT | XLSX / PDF |
| 2 | Base de datos de oferta académica de las IES | SENESCYT | XLSX / PDF |
| 3 | Base estadística de docentes UEP 2015-2022 | SENESCYT | XLSX / PDF |
| 4 | Artículos publicados en revistas indexadas | SENESCYT | XLS / PDF |
| 5 | Base de datos de becarios (corte agosto 2024) | SENESCYT | CSV / PDF |
| 6 | Registro de Matrícula | Ministerio de Educación | CSV / XLSX |
| 7 | Descomposición de la Matrícula | Ministerio de Educación | CSV / XLSX |

## Herramientas

- **Python**: extracción, limpieza avanzada, EDA, técnica de segmentación/clustering.
- **KNIME**: workflow ETL (lectura, estandarización de llaves, unión, exportación).
- **Power BI**: modelo de datos tipo constelación, medidas DAX, dashboard de 4 páginas.

## Cómo reproducir

1. Clonar el repositorio y colocar los archivos originales en `data/raw/` (no se incluyen en el repo por tamaño; ver sección de descarga abajo).
2. Instalar dependencias: `pip install -r requirements.txt`
3. Ejecutar los notebooks en `notebooks/` en orden numerado.
4. Abrir `knime/etl_workflow.knwf` en KNIME Analytics Platform y ejecutar el flujo completo.
5. Abrir `powerbi/dashboard.pbix` en Power BI Desktop.

## Descarga de los datos originales

Por su tamaño, los archivos originales no se suben al repositorio. Descárgalos directamente desde:
https://www.datosabiertos.gob.ec/ (grupo **Educación**) y colócalos en `data/raw/` con los nombres indicados en `docs/diccionario_datos.md`.

## Autores

- [Nombre completo 1]
- [Nombre completo 2]

## Docente

Ing. Juan Pablo Zaldumbide — Análisis de Datos, período 2026-A.
