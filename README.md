# Seguimiento de Consultas 2026

Herramienta de la Supervisión Escolar (SEMS) para dar seguimiento diario a la consulta a estudiantes y a la consulta a docentes aplicadas en LimeSurvey, del 13 de noviembre al 11 de diciembre de 2026.

## Cómo funciona

- El Excel que exporta LimeSurvey se procesa **en el navegador**. Las respuestas individuales nunca salen del equipo.
- A la nube solo se envían conteos agregados por plantel, turno y semestre.
- Los cortes se guardan en Google Sheets mediante un servicio de Google Apps Script (`apps-script/Codigo.gs`).
- Hay dos bases de datos separadas: **PRUEBAS** y **PRODUCCIÓN**. El indicador de la esquina superior derecha muestra siempre a cuál está conectada la herramienta.

## Reglas que protegen los datos

- Un solo corte vigente por encuesta y fecha. Guardar otra vez la misma fecha reemplaza el corte anterior, que se conserva marcado como "reemplazado".
- El servidor rechaza planteles fuera del catálogo, fechas fuera del periodo y conteos que no suman el total del archivo.
- Pide confirmación si un corte trae menos respuestas que el anterior o si el mismo archivo ya se guardó con otra fecha.
- Toda operación queda registrada en la hoja Bitácora.

## Archivos

- `index.html`: la herramienta.
- `apps-script/Codigo.gs`: el servicio de datos. La clave de acceso no se publica aquí; vive solo en el script de cada hoja de Google.
