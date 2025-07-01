import pandas as pd
import smtplib
import os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import codecs
load_dotenv()

# Función para cargar la plantilla HTML
def cargar_plantilla_html(ruta_plantilla):
    """Carga la plantilla HTML desde un archivo"""
    try:
        with codecs.open(ruta_plantilla, 'r', encoding='utf-8') as file:
            return file.read()
    except FileNotFoundError:
        print(f"No se encontró la plantilla HTML en: {ruta_plantilla}")
        return None
    except Exception as e:
        print(f"Error al cargar la plantilla HTML: {str(e)}")
        return None

def personalizar_plantilla(plantilla, nombre, mensaje, nombre_remitente, titulo="Mensaje Personalizado"):
    """Personaliza la plantilla HTML con los datos del destinatario"""
    if plantilla is None:
        return None
    
    # Reemplazar las variables en la plantilla
    plantilla_personalizada = plantilla.replace("{{NOMBRE}}", nombre)
    plantilla_personalizada = plantilla_personalizada.replace("{{MENSAJE_PRINCIPAL}}", mensaje)
    plantilla_personalizada = plantilla_personalizada.replace("{{NOMBRE_REMITENTE}}", nombre_remitente)
    plantilla_personalizada = plantilla_personalizada.replace("{{TITULO_PRINCIPAL}}", titulo)
    
    return plantilla_personalizada

# Obtenemos el nombre del remitente desde las variables de entorno
name_account = os.getenv("name_account")
# Obtenemos el correo del remitente
email_account = os.getenv("email_account")
# Obtenemos la contraseña del correo (idealmente una contraseña de aplicación)
password_account = os.getenv("password_account")

# Configuración para usar plantilla HTML
usar_html = True  # Cambia a False si quieres usar texto plano
ruta_plantilla = "templates/email_template.html"

# Cargar la plantilla HTML
plantilla_html = None
if usar_html:
    plantilla_html = cargar_plantilla_html(ruta_plantilla)
    if plantilla_html is None:
        print("Fallback: Se usará formato de texto plano")
        usar_html = False

# Creamos una conexión segura SSL con el servidor SMTP de Gmail, en el puerto 465
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
# Realizamos el saludo (handshake) con el servidor
server.ehlo()
# Iniciamos sesión en la cuenta de correo con las credenciales
server.login(email_account, password_account)

# Leemos el archivo Excel que contiene la información de los correos a enviar
email_df = pd.read_excel("data/correos.xlsx")

# Extraemos las columnas relevantes del archivo Excel
all_names = email_df['Name']       # Nombres de los destinatarios
all_emails = email_df['Email']     # Correos de los destinatarios
all_subjects = email_df['Asunto']  # Asuntos personalizados (si existen)
all_messages = email_df['Mensaje'] # Mensajes personalizados (si existen)

# Definimos valores por defecto en caso de que no se proporcione asunto, mensaje o nombre
default_subject = "Asunto por defecto"
default_message = "Este es un mensaje por defecto."
default_name = "Estimado/a cliente"

# Iteramos sobre cada fila del archivo Excel
for i in range(len(email_df)):
    # Si no hay nombre, usamos el nombre por defecto
    if pd.isna(all_names[i]) or all_names[i] == '':
        name = default_name
    else:
        name = all_names[i]    # Nombre del destinatario actual
    
    email = all_emails[i]  # Correo del destinatario actual

    # Si no hay asunto personalizado, usamos el asunto por defecto
    if pd.isna(all_subjects[i]) or all_subjects[i] == '':
        subject = default_subject + ', ' + name + '!'
    else:
        subject = all_subjects[i] + ', ' + name + '!'

    # Si no hay mensaje personalizado, usamos el mensaje por defecto
    if pd.isna(all_messages[i]) or all_messages[i] == '':
        message_body = default_message
    else:
        message_body = all_messages[i]

    # Construimos el cuerpo del mensaje final
    if usar_html and plantilla_html:
        # Crear mensaje con plantilla HTML
        mensaje_html = personalizar_plantilla(
            plantilla_html, 
            name, 
            message_body, 
            name_account,
            subject.replace(', ' + name + '!', '')  # Remover el nombre del título
        )
        
        # Crear mensaje MIME multipart
        msg = MIMEMultipart('alternative')
        msg['From'] = f"{name_account} <{email_account}>"
        msg['To'] = f"{name} <{email}>"
        msg['Subject'] = subject
        
        # Versión en texto plano (fallback)
        texto_plano = ('Hey, ' + name + '!\n\n' +
                      message_body + '\n\n'
                      'Te deseamos lo mejor,\n' +
                      name_account)
        
        # Adjuntar ambas versiones
        part1 = MIMEText(texto_plano, 'plain', 'utf-8')
        part2 = MIMEText(mensaje_html, 'html', 'utf-8')
        
        msg.attach(part1)
        msg.attach(part2)
        
        sent_email = msg.as_string()
    else:
        # Formato tradicional de texto plano
        message = ('Hey, ' + name + '!\n\n' +
                  message_body + '\n\n'
                  'Te deseamos lo mejor,\n' +
                  name_account)

        # Construimos el formato completo del correo, incluyendo remitente, destinatario, asunto y mensaje
        sent_email = ("From: {0} <{1}>\n"
                      "To: {2} <{3}>\n"
                      "Subject: {4}\n\n"
                      "{5}"
                      .format(name_account, email_account, name, email, subject, message))
    
    # Intentamos enviar el correo, y si hay un error, lo mostramos
    try:
        server.sendmail(email_account, [email], sent_email)
        print(f'✅ Correo enviado exitosamente a {name} ({email})')
    except Exception as e:
        print('❌ No se pudo enviar el correo a {}. Error: {}\n'.format(email, str(e)))

# Cerramos la conexión con el servidor de correo
server.close()
