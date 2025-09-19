usb_monitor — Monitorización de USB

Este proyecto contiene un programa en Python diseñado para detectar la inserción y actividad en memorias USB extraíbles, con el objetivo de monitorizar posibles intentos de exfiltración de información.

Archivos incluidos

usb_monitor_autoclose.py
Versión principal. Preparada para integrarse en el sistema de Active Response de Wazuh, detecta memorias USB conectadas, registra la creación/eliminación de archivos y genera logs en formato JSON.

usb_monitor_script.py
Versión sencilla para ejecución manual. No está pensada para integrarse en Wazuh, útil solo para pruebas locales.

Funcionalidades principales

Detecta la conexión y desconexión de dispositivos USB.

Monitorea archivos añadidos o eliminados en la memoria.

Registra la actividad en un archivo de log (usb_log.txt) y en un archivo JSON estructurado (eventos.json).

Integra información de usuario, host y marca temporal en los registros. 

