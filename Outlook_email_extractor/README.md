Email Monitor & File Analyzer

Este proyecto permite detectar cuando se envía un correo mediante Outlook y generar un evento en formato JSON.
Además, si el correo contiene archivos adjuntos, estos son extraídos y analizados para registrar también su contenido.

Archivos incluidos

email_extractor.py

Monitorea la carpeta de Enviados en Outlook.

Registra la información principal de cada correo enviado (fecha, remitente, destinatarios, asunto, cuerpo, tamaño, etc.).

Si detecta archivos adjuntos, los guarda temporalmente y llama al script file_analyzer.py para analizarlos.

file_analyzer.py

Procesa los archivos adjuntos extraídos (PDF, TXT y DOCX).

Extrae texto (limitado a 3000 caracteres por archivo).

Genera un evento JSON con los datos del archivo (nombre, usuario, PC, fecha, contenido).

Elimina la carpeta temporal de adjuntos tras el análisis.

