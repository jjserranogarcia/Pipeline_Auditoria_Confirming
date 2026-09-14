# Pipeline de auditoría de confirming

El pipeline procesa los documentos de Google Drive y genera un PDF diario de monitorización en:

`Bonificaciones Confirming/Historial Bonificaciones de Confirming`

El informe se llama `Documentos Procesados DD-MM-AAAA.pdf`. Se crea aunque no haya documentos pendientes. Si la acción se ejecuta varias veces el mismo día, se actualiza el mismo archivo usando la fecha de España.

## Archivos que van en GitHub

- `Pipeline_Auditoria_Confirming.py`, en la raíz del repositorio.
- `mensual.yml`, dentro de `.github/workflows/`.
- `.gitignore`, en la raíz del repositorio.

`client_secret.json` y `obtener_refresh_token.py` se usan solo en tu ordenador para generar las credenciales. No subas `client_secret.json` a GitHub.

## Credenciales necesarias

El pipeline sigue necesitando OAuth de Google Drive, pero ya no usa la API de Gmail.

En Google Cloud debe estar habilitada `Google Drive API` y debes tener un cliente OAuth de tipo `Desktop app`. El archivo descargado se guarda junto a `obtener_refresh_token.py` con el nombre `client_secret.json`.

Si los Secrets actuales permiten que el pipeline acceda a Drive, no tienes que generar un refresh token nuevo. En ese caso, conserva estos tres Secrets del repositorio:

- `GDRIVE_CLIENT_ID`
- `GDRIVE_CLIENT_SECRET`
- `GDRIVE_REFRESH_TOKEN`

Solo si faltan, han sido revocados o ya no funcionan, ejecuta una vez en tu ordenador:

```bash
python obtener_refresh_token.py
```

El script abrirá el navegador para autorizar únicamente Google Drive y mostrará los tres valores que debes guardar en `Settings > Secrets and variables > Actions`.

## Puesta en marcha

1. Coloca los archivos en las rutas indicadas y súbelos al repositorio.
2. Comprueba que existen los tres Secrets de Drive.
3. Abre la pestaña `Actions` del repositorio.
4. Selecciona `Pipeline confirming mensual` y pulsa `Run workflow`.
5. Revisa la carpeta `Historial Bonificaciones de Confirming` en Drive.

La ejecución mensual automática se mantiene el día 8 a las 09:00 UTC. El nombre y la hora mostrados en el PDF usan la zona horaria de España (`Europe/Madrid`).

## Resultado esperado

Además de los Excel y PDF generados para cada documento y del CSV histórico, aparecerá un informe diario como:

`Documentos Procesados 14-09-2026.pdf`

El contenido conserva el control anterior por sociedades y bancos e indica si el conjunto recibido ha sido habitual o extraordinario.
