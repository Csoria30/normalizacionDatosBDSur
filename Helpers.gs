// FUNCIONES NORMALIZACION DE DATOS
function actualizarValoresEnHoja(datos, filaInicio, columna) {
    hojaActiva.getRange(filaInicio, columna, datos.length, 1).setValues(datos);
}

function eliminarPalabrasDuplicadas(datos) {
    const datosUnicos = [];

    for (let j = 0; j < datos.length; j++) {
        let texto = datos[j][0];

        // Eliminar palabras duplicadas
        const palabras = texto.split(' ');
        const palabrasUnicas = [...new Set(palabras)];
        texto = palabrasUnicas.join(' ');

        datosUnicos.push([texto]);
    }

    return datosUnicos;
}

function concatenarColumnas(datosNombre, datosLocalidades, datosInternos) {
    const datosConcatenados = [];

    for (let j = 0; j < datosNombre.length; j++) {
        let texto = datosNombre[j][0] + ' ' + datosLocalidades[j][0] + ' ' + datosInternos[j][0];
        datosConcatenados.push([texto]);
    }

    return datosConcatenados;
}

function obtenerDatosEnMinusculas(hoja, filaMin, columna) {
    const datos = getDatos(hoja, filaMin, columna);
    const esNumerica = datos.every(fila => !isNaN(parseFloat(fila[0])));

    if (esNumerica) {
        return datos;
    } else {
        return normalizarTexto(datos);
    }
}

function normalizarTexto(texto) {
    if (typeof texto === 'string') {
        //throw new Error('La variable texto no es una cadena de texto');
        return texto
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[áàâä]/g, "a")
            .replace(/[éèêë]/g, "e")
            .replace(/[íìîï]/g, "i")
            .replace(/[óòôö]/g, "o")
            .replace(/[úùûü]/g, "u")
            .replace(/\s+/g, " ")
            .replace(/-/g, " ") // Quita simbolo -
            .trim()
            .replace(/<[^>]*>/g, "")
            .replace(/[^\w\s]/g, "");
    }
    else
        return texto;
}

function buscarPrioridad() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaInstituciones = spreadsheet.getSheetByName("Instituciones");
    const COLUMNA_OBS = 6;

    // Selecciona toda la columna
    let colF = hojaInstituciones.getRange(1, COLUMNA_OBS, hojaInstituciones.getLastRow());
    colF.clearContent(); // Borra el contenido de la columna

    const datos = hojaInstituciones.getDataRange().getValues();
    const prioridades = {
        "prioridad1": "*****",
        "prioridad2": "****",
        "prioridad3": "***",
        "prioridad4": "**",
        "prioridad5": "*"
    };

    for (let i = 1; i < datos.length; i++) {
        const texto = datos[i][4].toString().toLowerCase();
        for (const prioridad in prioridades) {
            if (texto.includes(prioridad)) {
                hojaInstituciones.getRange(i + 1, COLUMNA_OBS).setValue(prioridades[prioridad]);
                break;
            }
        }
    }
}

function debuguear(e) {
    Logger.log(e);
    throw new Error('Debugueando...');
}

function getDatos(hoja, filaMin, columna) {
    return hoja.getRange(filaMin, columna, ultimaFila - 1).getValues();
}

function getDato(hoja, filaMin, columna) {
    return hoja.getRange(filaMin, columna, ultimaFila - 1).getValue();
}