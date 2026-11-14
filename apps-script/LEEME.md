# Copias por ejecutivo (Apps Script) — instalación

Este script crea, para cada ejecutivo, un archivo de Google Sheets con **solo su información** y lo comparte en lectura con sus correos. El dashboard lee esos archivos cuando alguien entra sin acceso a los Sheets principales.

## Instalar (una sola vez, ~10 minutos)

1. **Usa tu hoja de accesos** («accesos dashboard»). Mantén sus columnas **Nombre | Puesto | Correo | Acceso** y agrega una más al final: **Ve a**.

   | Nombre | Puesto | Correo | Acceso | Ve a |
   |---|---|---|---|---|
   | (dirección) | Director… | … | Total | *(vacío)* |
   | (ADV) | ADV | … | propio y de su ejecutivo | Nombre del ejecutivo de su dúo |
   | (gerente) | Gerente Comercial | … | Propio y su ejecutivo jr | Nombre(s) de su ejecutivo jr, separados por `;` |
   | (ejecutivo) | Ejecutivo Comercial | … | propio | *(vacío)* |

   - **Total** = dirección: ve todo entrando con su cuenta (necesita acceso de lectura a los Sheets principales). El script no les crea archivo.
   - **ADV**: ve solo los ejecutivos de «Ve a».
   - **Gerente / Ejecutivo**: ve su propio nombre más los de «Ve a».
   - Los nombres deben ser los que aparecen en Data Comercial (acentos y mayúsculas no importan).
   - El script escribe en la columna **Estado** (la crea si no existe) cuántas líneas llevó cada archivo o el error. Quien no esté en la lista no podrá ver nada.
2. En esa hoja: **Extensiones → Apps Script**. Borra lo que haya, pega **todo** el contenido de `dashboard-ejecutivos.gs` y guarda.
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

## Probar sin las personas: menú «Vista de»
Quien tiene acceso total (y es dueño de los archivos generados) ve arriba del dashboard un menú **👁 Vista de**. Al elegir a una persona (ADV, gerente o ejecutivo) el dashboard carga **sus archivos reales** y muestra exactamente lo que ella vería. «↩ Volver a mi vista» regresa a la vista completa. Requiere haber ejecutado antes «Actualizar copias ahora». No prueba los permisos de Google (eso lo hace Google al compartir cada archivo).
