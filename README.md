# Normalizacion de datos Sistema SUR
Se trata del Sistema de Gestión de Reclamos y Base de Datos de Instituciones utilizado por la Autopista de la Información del Gobierno de la Provincia de San Luis. En mi tarea, trabajé con la información exportada del módulo de instituciones, la cual fue solicitada para ser procesada, normalizada y adaptada según las reglas previamente establecidas.

## Proceso de normalización de datos
* Eliminación de Registros Inválidos: Se eliminaron los registros con formato inválido, incluyendo números de interno que comenzaban con 0 (no válidos en el sistema de telefonía) y registros que contenían letras en su cuerpo (no válidos como números de interno), debido a la falta de validación en la base de datos.

* Normalización de Nombres: Se normalizaron los nombres de los registros, eliminando acentos, símbolos y espacios innecesarios generados durante el ingreso de datos, lo que mejoró la consistencia y calidad de la información.

* Optimización para Transcripción de Voz: Se reemplazaron palabras por sinónimos para mejorar la compatibilidad con el sistema de transcripción de voz, lo que aumentó el porcentaje de entendimiento y precisión en las transferencias telefónicas automáticas por coincidencia de voz.

* Generación de Hoja de Datos Procesados: Se generó una hoja de datos procesados dinámicamente con el nombre y formato solicitados por el sistema de transferencia por voz, lista para ser utilizada en el proceso de transferencia automática.

### Optimización del Procesamiento de Datos

Inicialmente, se solicitó que los datos procesados fueran entregados en una hoja de cálculo de Google. Sin embargo, al notar que el procesamiento tardaba alrededor de 15 minutos, propuse implementar una solución en el backend para mejorar la eficiencia. Se me proporcionó un endpoint en el sistema de transcripción de voz para enviar los datos procesados en formato JSON.

### Implementación de Node.js
Implementé una solución utilizando Node.js con JavaScript, lo que permitió reducir significativamente el tiempo de procesamiento a solo segundos. Los datos procesados fueron enviados al endpoint solicitado y posteriormente guardados en una base de datos.

### Resultados
La implementación de Node.js logró:
* Reducir el tiempo de procesamiento de 15 minutos a segundos
* Enviar los datos procesados al endpoint solicitado en formato JSON
* Guardar los datos procesados en una base de datos para su posterior uso.