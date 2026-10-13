# Sistema de Control de Asistencia — Santa Lucía

Aplicación web para el registro y gestión de asistencia de estudiantes y personal
de un instituto (3 turnos), con panel de administración, reportes y cierres automáticos.

> Portafolio: código fuente para lectura y descarga. Para ejecutarlo se despliega
> en Google Apps Script (`clasp push`) sobre un Google Sheet propio.

## Stack
Google Apps Script (V8) · Google Sheets (base de datos) · Google Drive (archivos) ·
Google Forms (solicitudes) · HTML/CSS/JS · SheetJS · qrcodejs

## Estructura
- `Código.js` — backend (~80 funciones: asistencia, control diario, alumnos,
  personal, solicitudes, reportes, usuarios/permisos, configuración, caché, triggers).
- `index.html` + `java.html` — portada de marcación (reloj de servidor, padrón
  instantáneo, responsive móvil).
- `admin.html` + 8 `admin-*.html` — panel (inicio, alumnos, personal, reportes,
  usuarios, configuración).
- `css.html` — estilos portada + panel + responsive.
- `appsscript.json` — manifiesto (V8, zona America/Lima).

## Despliegue (resumen)
1. Hoja nueva + `Extensiones > Apps Script`.
2. Copiar estos archivos (respetando nombres) o `clasp push`.
3. Crear hojas con el instalador incluido en el proyecto original.
4. Configurar `CONFIGURACION`, usuarios, triggers y `Implementar > Aplicación web.
