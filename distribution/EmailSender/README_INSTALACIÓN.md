# EmailSender - Instrucciones de Instalación

¡Gracias por descargar EmailSender!

## INSTALACIÓN

1. Extrae todos los archivos de este ZIP a una carpeta de tu elección
2. Asegúrate de mantener la estructura de carpetas:
   - EmailSender.exe (archivo principal)
   - templates/ (plantillas de correo)
   - config/ (archivos de configuración)
   - data/ (archivos de datos)
   - icon.ico (icono de la aplicación)

## CONFIGURACIÓN INICIAL

1. Ejecuta EmailSender.exe
2. Ve a la pestaña "Configuración"
3. Completa los siguientes campos:
   - Nombre del remitente: Tu nombre
   - Email del remitente: Tu dirección de Gmail
   - Contraseña de aplicación: Contraseña específica de aplicación de Gmail
   
   **IMPORTANTE**: Debes usar una "contraseña de aplicación" de Gmail, no tu contraseña normal.
   [Guía para crear contraseña de aplicación](https://support.google.com/mail/answer/185833)

4. Configura el asunto y mensaje predeterminados
5. Haz clic en "Guardar Configuración"
6. Usa "Probar Conexión" para verificar que todo funciona

## USO BÁSICO

1. Carga tu archivo Excel/CSV con datos de estudiantes en la pestaña "Gestión de Datos"
2. Ve a "Vista Previa de Correos" para verificar los datos
3. En la pestaña "Envío de Correos", usa "Envío de Prueba" primero
4. Si todo funciona bien, configura el rango y envía los correos masivos

## ESTRUCTURA DE DATOS REQUERIDA

Tu archivo Excel/CSV debe tener estas columnas:
- id
- Cod.Universitario
- Nombres
- Apellidos
- Facultad
- Escuela
- Correo

## SOLUCIÓN DE PROBLEMAS

- Si hay errores de conexión, revisa tu conexión a Internet
- Si hay errores de autenticación, verifica tu contraseña de aplicación
- Para más ayuda, consulta SOLUCION_ERRORES.md

## ARCHIVOS INCLUIDOS

- EmailSender.exe: Aplicación principal
- templates/: Plantillas HTML para los correos
- config/user_config_ejemplo.json: Ejemplo de configuración
- data/estudiantes_ejemplo.xlsx: Ejemplo de datos
- README_INSTALACION.txt: Este archivo
- README.md: Documentación completa
- SOLUCION_ERRORES.md: Guía de solución de problemas

## REQUISITOS DEL SISTEMA

- Windows 10 o superior
- Conexión a Internet
- Cuenta de Gmail con autenticación de 2 factores habilitada

## CONTACTO

Para soporte técnico o reportar problemas, consulta la documentación incluida.

¡Disfruta usando EmailSender!
