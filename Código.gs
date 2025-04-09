// Funciones Principales
function separarValores() {
    // Expresión regular para identificar delimitadores
    var delimitadores = /[\/\-\s]+|\/\/+/;
    var filaActual = 2;
 try {
    while (filaActual <= ultimaFila) {
        var valorD = getDato(hojaActiva, filaActual, COL_INTERNOS);
        var valorB = getDato(hojaActiva, filaActual, COL_NOMBRE);
        var valorE = getDato(hojaActiva, filaActual, COL_OBS);

        if (delimitadores.test(valorD)) {
            var valoresD = valorD.split(delimitadores);
            var j = 1;

            while (j < valoresD.length) {
                hojaInstituciones.insertRowAfter(filaActual + j - 1);
                ultimaFila++; // Actualizar última fila
                var valores = [
                    hojaInstituciones.getRange(filaActual, 1).getValue(),
                    valorB,
                    hojaInstituciones.getRange(filaActual, 3).getValue(),
                    valoresD[j],
                    valorE
                ];
                hojaInstituciones.getRange(filaActual + j, 1, 1, COL_INTERNOS + 1).setValues([valores]);
                j++;
            }

            // Reemplazar el valor de la columna D en la fila original
            hojaInstituciones.getRange(filaActual, COL_INTERNOS).setValue(valoresD[0]);
        }

        filaActual++;
    }

   
    } catch (error) {
        SpreadsheetApp.getUi().alert("Error al separa filas: " + error.message);
    }
}