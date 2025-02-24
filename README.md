# Scripts de Automatizacion

## Proyecto Aun en construccion

Este proyecto consisite en varios scripts de automatizacion en Pyton para mejorar la eficiencia de un Analista de Impuestos con la finalidad de realizar varias tareas cotidianas en un solo click. Si tu trabajo consiste en descargar las opiniones de mas de 10 empresas al dia estos scripts te ayudaran

### Descarga Masiva d las opiniones de cumplimiento IMSS e INFONAVIT
 - Archivos que realizan la descarga de las opiniones de cumplimiento de los portales correspondeitnes
 - Ambos scripts necesitan de la URL de un archivo de Google Sheets (Proximamente lo agregare el link de Google Drive)
 - Las opiniones de IMSS reunen la informacion necesaria de sheets, que incluyen las rutas a cada certificado asi como los usuarios y contraseñas
 - Una vez reunida la informacion de Google Sheets en automatico el script ingresara al Buzon IMSS y descargara las opiniones de cada uno de tus Registros Patronales
 - En caso de que no se puedan descargar por algun error en la pagina web o del script, este tomara una captura de pantalla antes de cerrar el navegador
 - Para las opiniones de INFONAVIT el script esta preprado para realizar el captcha de acceso al portal empresarial y de manera automatica hara todo el proceso de descarga de las opiniones

Proximamente se subira un script para realizar las confrontas mensuales y bimestrales de manera automatica
