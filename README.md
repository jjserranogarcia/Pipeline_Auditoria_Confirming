# Pipeline de auditoría de bonificaciones de confirming

**Pipeline ETL en Python que automatiza la extracción, comprobación y monitorización de bonificaciones de confirming**, utilizando Google Drive para gestionar los documentos y GitHub Actions para ejecutar el proceso.

> **Datos ficticios:** Grupo Prisma, las sociedades, los bancos, los proveedores y las operaciones se han simulado con fines demostrativos.

## 1. 🎯 Contexto y objetivo

En el *confirming*, los proveedores pueden **anticipar el cobro de sus facturas**. El banco cobra una comisión e intereses por el anticipo y entrega una parte de esos ingresos a la sociedad pagadora en concepto de **bonificación**.

![Funcionamiento del confirming y reparto de las bonificaciones](imagenes/01_Explicacion_Confirming.png)

El proyecto sustituye la revisión manual de esas bonificaciones en Excel por un proceso automatizado. El supuesto contempla **8 sociedades y 8 bancos**, es decir, **64 documentos por ejecución mensual completa**.

## 2. 📥 Documentos de entrada

Cada banco presenta la información de forma diferente. El pipeline identifica el banco, la sociedad y la fecha a partir del nombre del archivo (`BANCO_SOCIEDAD_FECHA`), y adapta la extracción al formato recibido.

![Sociedades, bancos, estructuras documentales y tecnologías de extracción](imagenes/02_Contexto_Sociedades_Bancos_Documentos.png)

- **6 bancos:** PDF con texto extraíble, tratados con `pypdf` y expresiones regulares.
- **1 banco:** PDF escaneado, tratado con Poppler y Tesseract (OCR).
- **1 banco:** archivo Excel, leído mediante `openpyxl`.

<details>
<summary><strong>Ver ejemplos de los cuatro formatos de entrada</strong></summary>

**PDF con listado de facturas (Oro)**

![PDF con listado de facturas](imagenes/03_Documento_PDF_Lista.png)

**PDF agrupado por proveedor (Plata)**

![PDF con facturas agrupadas por proveedor](imagenes/04_Documento_PDF_Proveedor.png)

**PDF escaneado con datos en dos tablas (Esmeralda)**

![PDF con datos repartidos en dos tablas](imagenes/05_Documento_PDF_Separado.png)

**Libro de Excel (Zafiro)**

![Documento Excel recibido de una entidad bancaria](imagenes/06_Documento_Excel.png)

</details>

## 3. 🗺️ Arquitectura del proceso

El departamento financiero deposita los documentos en Google Drive. **GitHub Actions ejecuta el pipeline**, que identifica, extrae y normaliza la información, realiza los cálculos y devuelve los resultados a Drive.

![Arquitectura funcional del pipeline automatizado](imagenes/07_Arquitectura_Funcional.png)

[Ver también el diagrama técnico del repositorio en GitDiagram](https://gitdiagram.com/jjserranogarcia/pipeline_auditoria_confirming?utm_source=readme&utm_medium=picture).

## 4. 🔎 Comprobación de bonificaciones

Para cada factura, el programa **recalcula la parte de comisión e interés que corresponde a Grupo Prisma**, aplicando el porcentaje de reparto y una base de **360 días** para el interés. Después compara ambos importes con los comunicados por el banco y registra las diferencias.

![Documento procesado con cálculos y diferencias de bonificaciones](imagenes/08_Documento_Generado.png)

El resultado permite al departamento financiero localizar discrepancias y decidir cuáles requieren revisión o reclamación.

## 5. 📤 Resultados y monitorización

Cada documento de entrada genera un **Excel procesado** y un **PDF de revisión**. Además, el programa actualiza el **histórico acumulado en CSV**, archiva los originales en Google Drive y crea un **PDF de monitorización** de la ejecución.

![Informe de monitorización que registra los documentos procesados](imagenes/09_Comprobacion_Generada.png)

El informe de monitorización identifica los documentos procesados por sociedad y banco y comprueba si se han recibido las **64 combinaciones esperadas**.

## 6. ⚙️ Automatización y tecnologías

El workflow [`mensual.yml`](.github/workflows/mensual.yml) programa la ejecución para el **día 8 de cada mes** y también permite lanzarla manualmente desde **GitHub Actions**.

![Ejecuciones del pipeline en GitHub Actions y opción de ejecución manual](imagenes/10_Pipeline_Automatizado.png)

**Tecnologías:** Python 3.11, `pypdf`, `openpyxl`, Poppler, Tesseract, LibreOffice, ReportLab, Google Drive API, PyDrive2, OAuth 2.0 y GitHub Actions.

La conexión con Drive utiliza **credenciales OAuth 2.0 guardadas en los Secrets de GitHub**, sin publicarlas en el código.

## 7. ✅ Validación y limitaciones

La prueba completa procesó **64 documentos simulados sin incidencias**, generó los informes y actualizó el histórico. La auditoría automática detecta diferencias, pero **la decisión sobre una posible reclamación corresponde al departamento financiero**.

La extracción depende del formato de los documentos bancarios y los PDF escaneados pueden contener errores de reconocimiento OCR.

## 8. 📊 Proyecto complementario: Power BI

El **histórico CSV producido por este pipeline** alimenta el dashboard **[Reporte_Auditoria_Confirming](https://github.com/jjserranogarcia/Reporte_Auditoria_Confirming)**, que permite analizar importes anticipados y bonificaciones por sociedad, banco, proveedor y región.
