# python-send-mail
python-send-mail es una aplicación sencilla y eficiente desarrollada en Python que permite el envío automatizado de correos electrónicos a través del protocolo SMTP. 

# Paso 01: Crea el entorno virtual
```sh
py -3 -m venv venv
```

# Paso 02: Activar entorno virtual
```sh
.\venv\Scripts\activate
```

# Paso 03: Actualizamos pip
```sh
pip install --upgrade pip
```
Si sale error, prueba con este:
```sh
python.exe -m pip install --upgrade pip
```

# Paso 04: Instala dependencias desde requirements.txt
```sh
pip install -r requirements.txt
```

# Paso 05:
Accede a la [página de verificación en dos pasos](https://myaccount.google.com/signinoptions/two-step-verification) de tu cuenta de Gmail para activar esta opción.
![image](https://github.com/user-attachments/assets/b3d8cdae-c610-4759-9c96-ed7739515899)

# Paso 06:
Luego, ve a la [página de contraseñas de la aplicación](https://myaccount.google.com/apppasswords) para generar una contraseña.
Escribe un nombre de aplicación y crear.
![image](https://github.com/user-attachments/assets/82ba401f-1f07-4338-999c-266c901ed3bb)

# Paso 07:
Finalmente se mostrará una contraseña de 16 caracteres, completamente necesaria para iniciar sesión en Gmail.
![image](https://github.com/user-attachments/assets/56b3479f-4e96-4860-950e-d2355dc7a283)

# Paso 08: Archivo .env
-- Crea un archivo .env con las siguientes caracteristicas:

```name_account="nombre-que-saldra-al-momento-de-enviar-un-correo"```

```email_account="tucorreo@gmail.com"```

```password_account="contraseña-de-16-caracteres-generada-en-el-paso-02"```

# Paso 09: Poner los correos en el Excel
Dependiendo la cantidad de correos que hay en el Excel, sera la cantidad de envio. Los campos como Name, Asunto, Mensaje son opcionales.
<img width="724" alt="image" src="https://github.com/user-attachments/assets/c14dab46-849c-4791-9423-9ed73d9175cd" />

# Recomendaciones: Cambiar los mensajes por default en el archivo index.py
![image](https://github.com/user-attachments/assets/cf715357-39d4-4fa9-913d-27b05f012cb6)

# Paso 10: Ejecutar el archivo index.py
```sh
python index.py
```
