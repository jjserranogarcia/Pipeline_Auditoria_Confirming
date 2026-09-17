# -*- coding: utf-8 -*-
"""
Este programa obtiene las credenciales que permiten al pipeline acceder a
Google Drive. Se ejecuta una sola vez en el ordenador del responsable del
departamento financiero.

Antes de iniciarlo debe crearse en Google Cloud un cliente OAuth de tipo
"Desktop app" y guardarse el archivo descargado como client_secret.json en
la misma carpeta. También debe estar activada la API de Google Drive.

Al finalizar se muestran el client id, el client secret y el refresh token,
que deben guardarse como Secrets del repositorio de GitHub. Si el token existente sigue
permitiendo el acceso a Drive, no es necesario repetir esta autorización ni
modificar los Secrets utilizados por el workflow.
"""
try:
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    # Si la biblioteca no está disponible, se instala con el mismo intérprete.
    # Así se evita que el comando pip apunte a otra instalación de Python.
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "google-auth-oauthlib"])
    from google_auth_oauthlib.flow import InstalledAppFlow

# PERMISO SOLICITADO A GOOGLE DRIVE
"""
Este alcance permite consultar, crear y mover los archivos del proyecto.
La autorización se limita a Drive porque el resultado de la ejecución se
guarda como PDF y no se envía mediante correo electrónico.
"""
ALCANCE = [
    "https://www.googleapis.com/auth/drive",
]

# Crea el flujo de autorización con la configuración de client_secret.json.
flujo = InstalledAppFlow.from_client_secrets_file("client_secret.json", ALCANCE)

# AUTORIZACIÓN Y PRESENTACIÓN DE LAS CREDENCIALES
"""
El navegador permite seleccionar la cuenta y aceptar el acceso a Drive.
El bloque try/except muestra cualquier fallo de autorización sin cerrar
bruscamente el proceso. Si finaliza correctamente, se presentan los tres
valores que deberán copiarse en los Secrets de GitHub.
"""
try:
    credenciales = flujo.run_local_server(port=0)
except Exception as error:
    print(f"\nAlgo ha fallado al autorizar: {error}")
else:
    print("\n¡Listo! Guarda estos tres valores como Secrets en GitHub")
    print("(Settings > Secrets and variables > Actions > New repository secret):\n")
    print("GDRIVE_CLIENT_ID     =", credenciales.client_id)
    print("GDRIVE_CLIENT_SECRET =", credenciales.client_secret)
    print("GDRIVE_REFRESH_TOKEN =", credenciales.refresh_token)
    print("\nEste refresh token no caduca (mientras el proyecto siga en modo")
    print("\"In production\" y yo no revoque el acceso a mano), así que solo")
    print("tengo que hacer esto una vez.")

# PAUSA FINAL
"""
La espera impide que la ventana se cierre inmediatamente cuando el archivo
se abre con doble clic. Así, las credenciales obtenidas o el mensaje de error
permanecen visibles hasta que la persona responsable pulse la tecla Enter.
"""
input("\nPulsa Enter para cerrar esta ventana...")
