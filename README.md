# python-send-mail
python-send-mail es una aplicación sencilla y eficiente desarrollada en Python que permite el envío automatizado de correos electrónicos a través del protocolo SMTP. 

**Ahora disponible en dos versiones:**
- **Versión de línea de comandos** (`index.py`): Para uso automatizado y scripts
- **Versión de escritorio** (`app.py`): Interfaz gráfica completa para gestión visual de correos

## 🚀 Características de la Aplicación de Escritorio (app.py)

- **Vista previa de correos**: Visualiza todos los correos con filtros por rango
- **Gestión completa de datos**: Agregar, editar y eliminar correos individualmente
- **Importación masiva**: Pega listas de correos para importar en lote
- **Envío controlado**: Selecciona rangos específicos para envío
- **Plantillas HTML**: Soporte completo para plantillas personalizadas
- **Log en tiempo real**: Monitoreo del proceso de envío
- **Configuración integrada**: Gestión de credenciales desde la interfaz 

# Paso 01: Clonar este repositorio
```sh
git clone https://github.com/Rodrious73/python-send-mail.git
```

# Paso 02: Crea el entorno virtual
```sh
py -3 -m venv venv
```

# Paso 03: Activar entorno virtual
```sh
.\venv\Scripts\activate
```

# Paso 04: Actualizamos pip
```sh
pip install --upgrade pip
```
Si sale error, prueba con este:
```sh
python.exe -m pip install --upgrade pip
```

# Paso 05: Instala dependencias desde requirements.txt
```sh
pip install -r requirements.txt
```

# Paso 06:
Accede a la [página de verificación en dos pasos](https://myaccount.google.com/signinoptions/two-step-verification) de tu cuenta de Gmail para activar esta opción.
![image](https://github.com/user-attachments/assets/b3d8cdae-c610-4759-9c96-ed7739515899)

# Paso 07:
Luego, ve a la [página de contraseñas de la aplicación](https://myaccount.google.com/apppasswords) para generar una contraseña.
Escribe un nombre de aplicación y crear.
![image](https://github.com/user-attachments/assets/82ba401f-1f07-4338-999c-266c901ed3bb)

# Paso 08:
Finalmente se mostrará una contraseña de 16 caracteres, completamente necesaria para iniciar sesión en Gmail.
![image](https://github.com/user-attachments/assets/56b3479f-4e96-4860-950e-d2355dc7a283)

# Paso 09: Archivo .env
-- Crea un archivo .env con las siguientes caracteristicas:

```name_account="nombre-que-saldra-al-momento-de-enviar-un-correo"```

```email_account="tucorreo@gmail.com"```

```password_account="contraseña-de-16-caracteres-generada-en-el-paso-02"```

# Paso 10: Poner los correos en el Excel
Dependiendo la cantidad de correos que hay en el Excel, sera la cantidad de envio. Los campos como Name, Asunto, Mensaje son opcionales.
<img width="724" alt="image" src="https://github.com/user-attachments/assets/c14dab46-849c-4791-9423-9ed73d9175cd" />

# Recomendaciones: Cambiar los mensajes por default en el archivo index.py
![image](https://github.com/user-attachments/assets/cf715357-39d4-4fa9-913d-27b05f012cb6)

# Paso 11: Ejecutar la aplicación

## Versión de Escritorio (Recomendada)
```sh
python app.py
```

## Versión de Línea de Comandos
```sh
python index.py
```

## 📱 Uso de la Aplicación de Escritorio

### Pestaña "Vista Previa de Correos"
- Selecciona rangos de correos para visualizar (ej: 1-50, 51-100)
- Ve un resumen de todos los correos con nombres, emails, asuntos y mensajes
- Haz clic en cualquier correo para ver los detalles completos

### Pestaña "Gestión de Datos"
- **Agregar individual**: Completa los campos y presiona "Agregar"
- **Editar**: Selecciona un correo de la vista previa, modifica los campos y presiona "Actualizar"
- **Eliminar**: Selecciona un correo y presiona "Eliminar"
- **Importación masiva**: Pega una lista de emails (uno por línea) y presiona "Importar"
- **Guardar**: Presiona "Guardar Cambios" para actualizar el archivo Excel

### Pestaña "Envío de Correos"
- Selecciona el rango de correos a enviar
- Activa/desactiva el uso de plantillas HTML
- Usa "Envío de Prueba" para probar con el primer correo
- Monitorea el progreso en tiempo real con la barra de progreso y log

### Pestaña "Configuración"
- Configura tus credenciales de correo
- Selecciona plantillas HTML personalizadas
- Prueba la conexión con el servidor

## 🎯 Ventajas de la Versión de Escritorio
- **Visualización completa**: Ve todos tus correos antes de enviar
- **Control granular**: Envía rangos específicos de correos
- **Gestión visual**: Agrega, edita y elimina correos con facilidad
- **Importación rápida**: Pega listas completas de emails
- **Monitoreo en tiempo real**: Ve el progreso del envío
- **Sin necesidad de editar código**: Todo se gestiona desde la interfaz
