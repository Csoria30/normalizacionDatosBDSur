# Normalizacion de datos Sistema Sur

## Proceso de normalización de datos
* Se separaron los valores en la columna de número de internos en registros individuales, debido a la existencia de registros con múltiples internos registrados, resultado de la falta de validación en la base de datos y el formulario de edición.

* Se eliminaron los registros con formato inválido, incluyendo internos que comienzan con 0 (no válidos en el sistema de telefonía) y registros que contienen letras en su cuerpo (no válidos como números de interno), debido a la falta de validación en la base de datos.

* Se normalizaron los nombres de los registros, eliminando acentos, símbolos y espacios innecesarios generados durante el ingreso de datos, debido a la configuración de la base de datos.

* Se reemplazaron palabras por sinónimos para mejorar la compatibilidad con el sistema de transcripción de voz que procesará la información, aumentando así el porcentaje de entendimiento y precisión en las transferencias telefónicas automáticas por coincidencia de voz.

* Se generó una hoja de datos procesados dinámicamente con el nombre y formato solicitados por el sistema de transferencia por voz, lista para ser utilizada en el proceso de transferencia automática.
