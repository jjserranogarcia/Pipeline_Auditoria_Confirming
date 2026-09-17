Pipeline de extracción, transformación y carga de datos para automatizar la revisión de las bonificaciones de confirming de un holding ficticio llamado Grupo Prisma. 

Este holding empresarial está compuesto por ocho sociedades las cuales trabajan con ocho bancos, por lo que una ejecución mensual completa incluye 64 combinaciones sociedad-banco. 

Este proyecto sustituye el proceso manual de copia de los datos en un documento Excel en el que se aplican fórmulas para realizar las comprobaciones por el departamento financiero, por un proceso automatizado y periódico más eficiente que minimiza la carga de trabajo de los empleados y reduce el riesgo de errores humanos.

Los documentos de entrada (bonificaciones de confirming) son simulados y reproducen tres situaciones habituales: 
1.	documentos PDF con texto extraíble (emitidos por 6 bancos)
2.	PDF escaneado sin texto extraíble (emitido por 1 banco)
3.	documento de Excel (emitido por 1 banco)

El programa identifica el banco, la sociedad y la fecha, y a continuación extrae y normaliza los datos con reglas adaptadas a cada banco, recalcula la parte de comisión e interés correspondiente a la empresa del holding, y compara ese importe con el comunicado y abonado por el banco.

Finalmente, el pipeline genera como resultado un Excel y un PDF procesados por cada documento de entrada, actualiza una base de datos en formato CSV, archiva el original y crea un PDF para monitorizar el correcto funcionamiento de cada ejecución del código. 

Para llevar a cabo este proceso, GitHub Actions prepara una máquina virtual temporal y ejecuta el programa desarrollado de forma periódica y automática, y Google Drive junto con OAuth 2.0 actúa como espacio de entrada y salida de documentos.
