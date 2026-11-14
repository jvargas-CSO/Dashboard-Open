# Copias por ejecutivo (Apps Script) — instalación

Este script crea, para cada ejecutivo, un archivo de Google Sheets con **solo su información** y lo comparte en lectura con sus correos. El dashboard lee esos archivos cuando alguien entra sin acceso a los Sheets principales.

## Instalar (una sola vez, ~10 minutos)

1. **Crea un Google Sheet nuevo** en tu Drive (por ejemplo, «Accesos Dashboard»). Renombra la primera pestaña a **`Accesos`** y pon estos encabezados en la fila 1:

   | Correo | Rol | Ejecutivo(s) | Estado |
   |---|---|---|---|
   | ejemplo@opent.com.mx | Ejecutivo | Ricardo Ramírez | *(lo llena el script)* |
   | adv@opent.com.mx | ADV | Ricardo Ramírez; Marie Tenorio | |

   - **Ejecutivo(s)**: el nombre como aparece en Data Comercial (acentos y mayúsculas no importan). Si una persona ve a varios (ADV, mentor), sepáralos con `;`.
   - **Rol** es solo informativo; lo que da acceso es la columna Ejecutivo(s).
   - Quien **no** esté en la lista no podrá ver nada.
2. En ese Sheet: **Extensiones → Apps Script**. Borra lo que haya, pega **todo** el contenido de `dashboard-ejecutivos.gs` y guarda.
3. Recarga el Sheet. Aparece el menú **Dashboard**. Elige **Actualizar copias ahora**. La primera vez Google pide autorización (Drive y Sheets): acéptala con tu cuenta.
4. Menú **Dashboard → Programar actualización diaria** (corre cada día cerca de las 5:00).
5. En **Google Cloud Console → Pantalla de consentimiento OAuth**, si sigue en modo *Testing*, agrega a cada ejecutivo como **Test user** (o publica la app / ponla como *Internal*). Sin eso, Google no los deja iniciar sesión.

## Qué pasa después
- Se crea la carpeta **Dashboard Ejecutivos** con un archivo por ejecutivo: `Dashboard Ejecutivo - <Nombre>`. **No compartas la carpeta**: cada archivo se comparte solo con sus correos.
- La columna **Estado** de Accesos dice cuántas líneas llevó cada archivo, o el error. Si dice «Sin líneas», el nombre del ejecutivo no coincide con Data Comercial.
- Google enviará un aviso por correo a cada persona cuando le comparta su archivo.
- Para **quitar acceso**: borra la fila (o el ejecutivo) y ejecuta **Actualizar copias ahora**; el script retira los permisos.
- Si agregas un ejecutivo nuevo o cambia su nombre, ejecuta **Actualizar copias ahora** (o espera a la actualización diaria).

## Qué lleva cada archivo
- Sus líneas de venta (todas las columnas) en las pestañas por año y en «Concentrado 2024».
- El historial de sus **clientes actuales** (los que tienen venta suya este año o el anterior) vendido por otros ejecutivos, **solo con venta y fechas**: sin costos, utilidad, ROI, comisiones ni proveedor, y con el nombre «Otro ejecutivo».
- Su forecast (pestañas «Forecast 2026», «Forecast 2027»…).
- Una pestaña «Info» con la fecha de la última actualización.

## Límites que conviene conocer
- Los datos se refrescan una vez al día (o con el botón manual), no al instante.
- El ejecutivo puede abrir su archivo directamente en Google Sheets y ver ahí lo mismo que ve en el dashboard; nada más.
- Los coordinadores y la pestaña Campañas son la etapa 2.
