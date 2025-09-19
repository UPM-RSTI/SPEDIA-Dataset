import os
import json
import sys
from datetime import datetime
import time
from PyPDF2 import PdfReader
import docx
import shutil

# Función para asegurar que la carpeta 'Eventos' existe
def ensure_event_folder_exists(folder_name='Eventos'):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

# Función para extraer texto de archivos PDF, TXT y DOCX
def extract_text_from_file(file_path):
    """
    Extract text content from PDF, TXT, or DOCX files.
    """
    try:
        if file_path.lower().endswith('.pdf'):
            # For PDF files
            with open(file_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                num_pages = len(pdf_reader.pages)
                return "\n".join([pdf_reader.pages[page_num].extract_text() for page_num in range(num_pages)])

        elif file_path.lower().endswith('.txt'):
            # For text files
            with open(file_path, 'r') as file:
                return file.read()

        elif file_path.lower().endswith('.docx'):
            # For Word documents
            doc = docx.Document(file_path)
            return "\n".join([paragraph.text for paragraph in doc.paragraphs])

        else:
            return "Unsupported file format. Only PDF, TXT, and DOCX files are supported."

    except Exception as e:
        return f"Error reading file {file_path}: {e}"

# Función para guardar los datos del email y archivo en un archivo JSON
def save_email_data(message, user, pc, email_date, file_path, content):
    """
    Save the analyzed email and file data to a JSON file.
    """
    try:
        # Path to the 'Eventos' folder and the JSON file
        eventos_folder = os.path.join('Eventos')
        json_file_path = os.path.join(eventos_folder, 'Files', "files_analyzed.json")

        # Ensure the 'Eventos' directory exists, create it if it doesn't
        ensure_event_folder_exists(eventos_folder)

        # Create the entry with email data and file content
        entry = {
            "Date": email_date,
            "actividad": "email_sent_with_file",
            "User": user,
            "PC": pc,
            "filename": os.path.basename(file_path),
            "content": content
        }

        # Append the entry to the JSON file (without overwriting it)
        with open(json_file_path, 'a', encoding='utf-8') as json_file:
            json.dump(entry, json_file, ensure_ascii=False)  # No indentación ni salto de línea adicional
            json_file.write('\n')  # Agrega un salto de línea después de cada objeto

        print(f"Analysis complete. Results saved in {json_file_path}")

    except Exception as e:
        print(f"Error saving data to JSON: {e}")

# Función para eliminar una carpeta y su contenido, excepto 'attached_email_file_tmp'
def delete_folder_except_tmp(folder_path, exception_folder='attached_email_file_tmp'):
    """
    Delete the folder and its contents, except the 'attached_email_file_tmp' folder.
    """
    try:
        # Check if the folder exists
        if os.path.exists(folder_path):
            for root, dirs, files in os.walk(folder_path, topdown=False):
                for name in files:
                    os.remove(os.path.join(root, name))
                for name in dirs:
                    # Skip deletion of the exception folder
                    if name != exception_folder:
                        shutil.rmtree(os.path.join(root, name))

            # If the folder is empty, remove it as well
            if os.path.isdir(folder_path) and folder_path != exception_folder:
                os.rmdir(folder_path)
            print(f"Folder {folder_path} deleted successfully.")
        else:
            print(f"Folder {folder_path} not found.")
    except Exception as e:
        print(f"Error deleting folder {folder_path}: {e}")

# Función para analizar todos los archivos en una carpeta
def analyze_files_in_folder(folder_path, user, pc, email_date):
    """
    Analyze all files in the specified folder and save results in a JSON file.
    """
    try:
        # Analyze each file in the folder
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                if file_path.lower().endswith(('.pdf', '.txt', '.docx')):
                    # Extract content
                    content = extract_text_from_file(file_path)

                    # Call the function to save the email and file data
                    save_email_data(None, user, pc, email_date, file_path, content[:3000])

        # After analyzing all files, delete the folder and its content, except 'attached_email_file_tmp'
        delete_folder_except_tmp(folder_path)

    except Exception as e:
        print(f"Error analyzing files in folder {folder_path}: {e}")


if __name__ == "__main__":
    # Get folder path from command-line arguments
    if len(sys.argv) < 5:
        print("Usage: python file_analyzer.py <folder_path> <user> <pc> <email_date>")
        sys.exit(1)

    folder_path = sys.argv[1]
    user = sys.argv[2]
    pc = sys.argv[3]
    email_date = sys.argv[4]

    # Analyze files in the folder
    analyze_files_in_folder(folder_path, user, pc, email_date)
