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

// Elimina filas que no cumple con el formato indicado

function eliminarFilas() {
    let datos = hojaInstituciones.getDataRange().getValues();
    let filaActual = 2;
    let filasEliminar = [];
    let valoresDuplicados = ["3530", "6880", "8928", "5201", "5960"];
    let internosEliminar = ["4840", "4841"];
    let valoresEncontrados = {};
    let MSG_INICIO = "Estamos eliminando los registros con formato incorrecto, por favor espere.";

    //Mensaje
    //Browser.msgBox(MSG_INICIO);

    try {

        while (filaActual <= datos.length) {
            let valorD = datos[filaActual - 1][3].toString().trim();
            let valorE = datos[filaActual - 1][4].toString().toLowerCase();
            let valorC = datos[filaActual - 1][2];

            // Verificando duplicados de internos
            /*
            if (valoresDuplicados.indexOf(valorD) != -1 && valoresEncontrados[valorD]) {
                filasEliminar.push(filaActual);
            }
            */
            //Eliminar internos 
            if (internosEliminar.includes(valorD)) 
            {
              filasEliminar.push(filaActual); // Agrega el número de fila a filasEliminar
            }

            if (valoresDuplicados.indexOf(valorD) != -1) 
            {
                valoresEncontrados[valorD] = true;
            }
            // Verificando de internos no validos
            else if (valorD == "" || isNaN(valorD) || (valorD != "" && (valorD.length < 4 || valorD.charAt(0) == "0" || valorD.charAt(0) == "9" || valorD.charAt(0) == "*"))) {
                if (valorC != "") {
                    filasEliminar.push(filaActual);
                }
            }
            // Verificando Observaciones
            else if (/no transferir|interno pasivo/i.test(valorE)) {
                if (valorC != "") {
                    filasEliminar.push(filaActual);
                }
            }

          filaActual++;
        }

        //Arreglo con indices a eliminar
        if (filasEliminar.length > 0) {
            filasEliminar.sort(function (a, b) { return b - a; });
            for (let i = 0; i < filasEliminar.length; i++) {
                hojaInstituciones.deleteRow(filasEliminar[i]);
            }
        }

        //Browser.msgBox("Eliminación completada");

    } catch (error) {
        SpreadsheetApp.getUi().alert("Error Eliminar Filas: " + error.message);
    }

}