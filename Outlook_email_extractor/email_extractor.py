import os
import time
import win32com.client
import json
from datetime import datetime
import sys
from pathlib import Path, PureWindowsPath, PurePosixPath
import subprocess  # Para ejecutar el script file_analizer.py

# Ruta del archivo de log
LOG_FILE = "C:\\Program Files (x86)\\ossec-agent\\active-response\\active-responses.log" if os.name == 'nt' else "/var/ossec/logs/active-responses.log"

# Carpeta para guardar los archivos adjuntos
ATTACHMENTS_FOLDER = os.path.join(Path.home(), "Documents", "attached_email_file_tmp")

# Función para escribir en el archivo de depuración
def write_debug_file(ar_name, msg):
    with open(LOG_FILE, mode="a") as log_file:
        ar_name_posix = str(PurePosixPath(PureWindowsPath(ar_name[ar_name.find("active-response"):])))
        log_file.write(str(datetime.now().strftime('%Y/%m/%d %H:%M:%S')) + " " + ar_name_posix + ": " + msg + "\n")

# Función para obtener la dirección de correo electrónico del remitente
def get_sender_email_address(mail):
    sender = mail.Sender
    sender_email_address = ""
    
    if sender.AddressEntryUserType == 0 or sender.AddressEntryUserType == 5:
        exch_user = sender.GetExchangeUser()
        if exch_user is not None:
            sender_email_address = exch_user.PrimarySmtpAddress
        else:
            sender_email_address = mail.SenderEmailAddress
    else:
        sender_email_address = mail.SenderEmailAddress
    
    return sender_email_address

# Función para obtener el nombre de usuario y el nombre del PC
def get_user_and_pc():
    user = os.getlogin()
    pc = os.getenv('COMPUTERNAME')
    return user, pc

# Función para asegurar que la carpeta 'Eventos' existe
def ensure_event_folder_exists(folder_name='Eventos'):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

# Función para guardar en el archivo JSON dentro de la carpeta 'Eventos'
def save_to_json(data, folder_name='Eventos', filename='correos.json'):
    ensure_event_folder_exists(folder_name)
    file_path = os.path.join(folder_name, 'Correos',filename)
    
    try:
        with open(file_path, 'a') as f:
            for entry in data:
                json.dump(entry, f)
                f.write('\n')  # Agregar una línea nueva después de cada objeto JSON
        write_debug_file(sys.argv[0],  f"Datos guardados exitosamente en {file_path}")
    except Exception as e:
        write_debug_file(sys.argv[0],  f"Error al guardar datos en {file_path}: {e}")

# Función para cargar los mensajes existentes desde el archivo JSON
def load_existing_messages(folder_name='Eventos', filename='correos.json'):
    write_debug_file(sys.argv[0],  f"Dentro de load_existing_messages()")
    file_path = os.path.join(folder_name, 'Correos', filename)
    if not os.path.exists(file_path):
        write_debug_file(sys.argv[0],  f"Archivo no encontrado: {file_path}. Se retorna una lista vacía.")
        return []
    
    try:
        with open(file_path, 'r') as f:
            write_debug_file(sys.argv[0],  f"Cargando mensajes desde {file_path}")
            return [json.loads(line) for line in f]
    except Exception as e:
        write_debug_file(sys.argv[0],  f"Error al cargar mensajes desde {file_path}: {e}")
        return []

# Función para obtener la carpeta de Enviados de Outlook
def get_sent_folder_for_outlook(namespace):
    try:
        sent = namespace.GetDefaultFolder(5)  # 5 representa la carpeta de enviados en Outlook
        return sent
    except Exception as e:
        write_debug_file(sys.argv[0],  f"Error accessing Sent folder: {e}")
        return None

# Función para determinar el tipo de cuenta y devolver la carpeta de Enviados correspondiente
def get_sent_folder():
    attempt_number = 0
    max_attempts = 5
    while attempt_number < max_attempts:
        try:
            write_debug_file(sys.argv[0],  "Dentro de get_sent_folder()")
            time.sleep(5)
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            namespace.Logon(ShowDialog=True)  # Solicitar el perfil
            
            sent_folder = get_sent_folder_for_outlook(namespace)
            if sent_folder:
                return sent_folder
        except Exception as e:
            attempt_number += 1
            write_debug_file(sys.argv[0],  f"Attempt {attempt_number}: Error initializing Outlook or accessing folders: {e}")
            if attempt_number >= max_attempts:
                write_debug_file(sys.argv[0],  "Max attempts reached. Exiting.")
                return None
            time.sleep(5)  # Esperar 5 segundos antes de reintentar
    return None

# Modificación en la función process_new_messages para guardar los archivos adjuntos
def process_new_messages(sent_folder):
    try:
        existing_messages = load_existing_messages()  # Cargar mensajes existentes desde el archivo JSON
        existing_subjects = {entry['Subject'] for entry in existing_messages}
        new_entries = []

        # Iterar sobre los mensajes no procesados en la carpeta de Enviados
        for message in sent_folder.Items:
            if message.Subject not in existing_subjects:
                sender = get_sender_email_address(message)
                user, pc = get_user_and_pc()

                # Limitar el contenido a 3000 caracteres
                truncated_content = message.Body[:3000] if message.Body else ""

                # Construir el objeto JSON para cada mensaje
                data = {
                    "Date": message.CreationTime.strftime('%Y-%m-%d %H:%M:%S'),
                    "actividad": "email_sent",
                    "User": user,
                    "PC": pc,
                    "To": message.To,
                    "CC": message.CC,
                    "BCC": message.BCC,
                    "From": sender,
                    "Size": message.Size,
                    "Attachments": [attachment.FileName for attachment in message.Attachments],
                    "Content": truncated_content,
                    "Subject": message.Subject
                }
                new_entries.append(data)

                # Si el correo tiene archivos adjuntos, procesarlos
                if message.Attachments.Count > 0:
                    write_debug_file(sys.argv[0], f"Correo con archivos adjuntos detectado: {message.Subject}")
                    save_attachments(message)

        # Guardar la información solo si hay nuevos mensajes
        if new_entries:
            save_to_json(new_entries)

    except Exception as e:
        write_debug_file(sys.argv[0], f"Error processing emails: {e}")

# Función para registrar el usuario actual
def log_current_user():
    try:
        user = os.getlogin()  # Obtiene el usuario que ha iniciado sesión
        write_debug_file(sys.argv[0], f"Script ejecutado por el usuario: {user}")
    except Exception as e:
        write_debug_file(sys.argv[0], f"Error al obtener el usuario actual: {e}")

# Función principal
def main():
    write_debug_file(sys.argv[0],  f"Llamando a sent_folder = get_sent_folder()")
    log_current_user()
    sent_folder = get_sent_folder()

    if not sent_folder:
        write_debug_file(sys.argv[0],  "Sent folder not found.")
        return
    else:
        write_debug_file(sys.argv[0],  f"Total items in Sent folder: {sent_folder.Items.Count}")

    # Bucle principal para verificar nuevos mensajes cada 10 segundos
    while True:
        try:
            write_debug_file(sys.argv[0],  "Ejecutando Código")
            process_new_messages(sent_folder)
            time.sleep(5)  # Esperar 5 segundos antes de volver a verificar
        except Exception as e:
            write_debug_file(sys.argv[0],  f"Error in main loop: {e}")
            time.sleep(5)

# Función para guardar los archivos adjuntos
def save_attachments(mail):
    try:
        # Crear carpeta base si no existe
        os.makedirs(ATTACHMENTS_FOLDER, exist_ok=True)

        # Crear el nombre de la subcarpeta con la fecha de envío del correo, remitente, destinatario y tamaño
        size_kb = round(mail.Size / 1024, 2)  # Convertir tamaño a KB con dos decimales
        subfolder_name = f"{mail.CreationTime.strftime('%Y%m%d%H%M%S')}_{size_kb}KB_{mail.SenderEmailAddress.replace(' ', '_')}_to_{mail.To.replace(' ', '_')}"
        subfolder_path = os.path.join(ATTACHMENTS_FOLDER, subfolder_name)
        os.makedirs(subfolder_path, exist_ok=True)

        # Guardar cada archivo adjunto
        for attachment in mail.Attachments:
            attachment_path = os.path.join(subfolder_path, attachment.FileName)
            attachment.SaveAsFile(attachment_path)
            write_debug_file(sys.argv[0], f"Archivo adjunto guardado en: {attachment_path}")

        # Llamar al script file_analizer.py pasando la ruta de la carpeta como argumento
        analyze_files(subfolder_path, mail)

    except Exception as e:
        write_debug_file(sys.argv[0], f"Error al guardar archivos adjuntos: {e}")

# Función para analizar los archivos adjuntos
def analyze_files(folder_path, mail):
    try:
        user, pc = get_user_and_pc()  # Obtener usuario y PC
        email_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # O usa un valor relevante para el correo
        
        script_path = os.path.join(os.getcwd(), "file_analyzer.py")
        subprocess.run([sys.executable, script_path, folder_path, user, pc, mail.CreationTime.strftime('%Y-%m-%d %H:%M:%S')], check=True)
        write_debug_file(sys.argv[0], f"Script file_analyzer.py ejecutado con los parámetros: {folder_path}, {user}, {pc}, {email_date}")
    except Exception as e:
        write_debug_file(sys.argv[0], f"Error al ejecutar file_analyzer.py: {e}")

        write_debug_file(sys.argv[0], f"Error al ejecutar file_analyzer.py: {e}")

# Ejecutar la función principal
if __name__ == "__main__":
    main()
