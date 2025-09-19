#!/usr/bin/python3
# Copyright (C) 2015-2022, Wazuh Inc.
# All rights reserved.

# This program is free software; you can redistribute it
# and/or modify it under the terms of the GNU General Public
# License (version 2) as published by the FSF - Free Software
# Foundation.

import os
import sys
import json
import datetime
from pathlib import PureWindowsPath, PurePosixPath

import win32api
import win32file
import time
import socket
import getpass
import threading  # Importar el módulo threading

if os.name == 'nt':
    LOG_FILE = "C:\\Program Files (x86)\\ossec-agent\\active-response\\active-responses.log"
else:
    LOG_FILE = "/var/ossec/logs/active-responses.log"

ADD_COMMAND = 0
DELETE_COMMAND = 1
CONTINUE_COMMAND = 2
ABORT_COMMAND = 3

OS_SUCCESS = 0
OS_INVALID = -1

class message:
    def __init__(self):
        self.alert = ""
        self.command = 0


def list_files_in_drive(drive):
    print("Files in Drive %s:" % drive)
    try:
        for root, dirs, files in os.walk(drive):
            print(f"\nRoot folder: {root}")
            print("Files:")
            for file in files:
                print(os.path.join(root, file))
            print("Folders:")
            for dir in dirs:
                print(os.path.join(root, dir))
    except Exception as e:
        print("Error accessing drive:", e)

def monitor_usb(argv):
    write_debug_file(argv[0], 'DENTRO')
    while True:
        drives = win32api.GetLogicalDriveStrings().split('\x00')[:-1]
        write_debug_file(argv[0], 'DENTRO')
        usb_detected = False  # Variable para verificar si se detectó un USB

        for device in drives:
            type = win32file.GetDriveType(device)
            print("Drive: %s" % device)
            #print(drive_types[type])
            print("-" * 72)
            if type == win32file.DRIVE_REMOVABLE:
                print("USB detected:", device)
                usb_detected = True
                list_files_in_drive(device)
                monitor_files_in_drive(device)

        if not usb_detected:  # Si no se detectó ningún USB, salir del bucle
            break

        time.sleep(2)  # Espera 2 segundos antes de volver a verificar


def monitor_files_in_drive(drive):
    monitored_files = {}
    scan_folder(drive, monitored_files)

    while True:
        new_files = {}
        scan_folder(drive, new_files)

        added_files = set(new_files.keys()) - set(monitored_files.keys())
        if added_files:
            print("New files added:")
            with open("usb_log.txt", mode="a") as test_file:
                test_file.write("Archivo añadido" + "\n")
                
            for file in added_files:
                print(os.path.join(drive, file))
                size = new_files[file]
                with open("usb_log.txt", mode="a") as test_file:
                    test_file.write("Archivo añadido: " + file + "\n")
                generate_json(file, drive, 'file_add', size)
            monitored_files.update(new_files)

        deleted_files = set(monitored_files.keys()) - set(new_files.keys())
        if deleted_files:
            # Check if all monitored files are deleted
            all_monitored_files_deleted = all(file in deleted_files for file in monitored_files)
            print("all_monitored_files_deleted:" + str(all_monitored_files_deleted))
            if all_monitored_files_deleted:
                print("USB disconnected. Stopping monitoring.")
                with open("usb_log.txt", mode="a") as test_file:
                    test_file.write("USB desconectado" + "\n")
                break
            else:
                print("Files deleted:")
                with open("usb_log.txt", mode="a") as test_file:
                    test_file.write("Archivo eliminado" + "\n")
                for file in deleted_files:
                    print(os.path.join(drive, file))
                    size = monitored_files[file]  # Obtener el tamaño del archivo eliminado
                    with open("usb_log.txt", mode="a") as test_file:
                        test_file.write("Archivo eliminado: " + file + "\n")
                    generate_json(file, drive, 'file_delete', size)
                for file in deleted_files:
                    monitored_files.pop(file)  # Eliminar el archivo de la lista de archivos monitoreados

        time.sleep(0.5)  # Espera 0.5 segundos antes de volver a verificar


def scan_folder(folder, file_dict):
    for root, dirs, files in os.walk(folder):
        for file in files:
            file_path = os.path.relpath(os.path.join(root, file), folder)
            file_size = os.path.getsize(os.path.join(root, file))
            file_dict[file_path] = file_size

def generate_json(file_path, drive, event_type, size):
    timestamp = datetime.datetime.now().strftime('%Y-%m-%dT%H-%M-%S')
    user = getpass.getuser()
    hostname = socket.gethostname()

    data = {
        "timestamp": timestamp,
        "user": user,
        "hostname": hostname,
        "filepath": os.path.dirname(file_path),
        "filename": os.path.basename(file_path),
        "size": size,
        "activity": event_type
    }

    json_entry = json.dumps(data, separators=(',', ':'))  # Formatear JSON en una sola línea

    # Agregar un salto de línea después de cada entrada de JSON
    json_entry += '\n'

    json_folder = os.path.join(os.getcwd(), 'active-response', 'bin', 'Eventos')
    os.makedirs(json_folder, exist_ok=True)
    json_path = os.path.join(json_folder, 'eventos.json')

    # Escribir la entrada JSON en el archivo
    with open(json_path, 'a') as json_file:
        json_file.write(json_entry)

def write_debug_file(ar_name, msg):
    with open(LOG_FILE, mode="a") as log_file:
        ar_name_posix = str(PurePosixPath(PureWindowsPath(ar_name[ar_name.find("active-response"):])))
        log_file.write(str(datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')) + " " + ar_name_posix + ": " + msg +"\n")


def setup_and_check_message(argv):
    # get alert from stdin
    input_str = ""
    for line in sys.stdin:
        input_str = line
        break

    write_debug_file(argv[0], input_str)

    try:
        data = json.loads(input_str)
    except ValueError:
        write_debug_file(argv[0], 'Decoding JSON has failed, invalid input format')
        message.command = OS_INVALID
        return message

    message.alert = data

    command = data.get("command")

    if command == "add":
        message.command = ADD_COMMAND
    elif command == "delete":
        message.command = DELETE_COMMAND
    else:
        message.command = OS_INVALID
        write_debug_file(argv[0], 'Not valid command: ' + command)

    return message


def send_keys_and_check_message(argv, keys):
    # build and send message with keys
    keys_msg = json.dumps({"version": 1,"origin":{"name": argv[0],"module":"active-response"},"command":"usb-monitor-autoclose","parameters":{"keys":keys}})

    write_debug_file(argv[0], keys_msg)

    print(keys_msg)
    sys.stdout.flush()

    # read the response of previous message
    input_str = ""
    while True:
        line = sys.stdin.readline()
        if line:
            input_str = line
            break

    write_debug_file(argv[0], input_str)

    try:
        data = json.loads(input_str)
    except ValueError:
        write_debug_file(argv[0], 'Decoding JSON has failed, invalid input format')
        return message

    action = data.get("command")

    if "continue" == action:
        ret = CONTINUE_COMMAND
    elif "abort" == action:
        ret = ABORT_COMMAND
    else:
        ret = OS_INVALID
        write_debug_file(argv[0], "Invalid value of 'command'")

    return ret


def main(argv):
    write_debug_file(argv[0], "Started")

    # validate json and get command
    msg = setup_and_check_message(argv)

    if msg.command < 0:
        sys.exit(OS_INVALID)

    if msg.command == ADD_COMMAND:
        alert = msg.alert["parameters"]["alert"]
        keys = [alert["rule"]["id"]]

        action = send_keys_and_check_message(argv, keys)

        if action != CONTINUE_COMMAND:
            if action == ABORT_COMMAND:
                write_debug_file(argv[0], "Aborted")
                sys.exit(OS_SUCCESS)
            else:
                write_debug_file(argv[0], "Invalid command")
                sys.exit(OS_INVALID)

        # Ejecutar monitor_usb en un hilo separado
        write_debug_file(argv[0], 'ANTES DE LLAMAR')
        usb_monitor_thread = threading.Thread(target=monitor_usb, args=(argv,))
        usb_monitor_thread.start()
        usb_monitor_thread.join()  # Espera a que el hilo termine
        write_debug_file(argv[0], 'DESPUES DE LLAMAR')

    elif msg.command == DELETE_COMMAND:
        # Similar a lo anterior, pero para el comando DELETE_COMMAND
        pass


if __name__ == "__main__":
    main(sys.argv)
