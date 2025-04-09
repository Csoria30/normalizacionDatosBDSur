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

// 2 - Elimina filas que no cumple con el formato indicado

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

// 3 - Depurar nombres datos
function depurarDatos() {
  const ministerios = {
    "m e": "ministerio educacion del ",
    "m ciencia e inn": "ministerio ciencia innovacion del ",
    "m des humano": "ministerio desarrollo humano del ",
    "m des hum": "ministerio desarrollo humano del ",
    "m des productivo": "ministerio desarrollo productivo del ",
    "m gobierno ": "ministerio gobierno del ",
    "m hacienda inf pub": "ministerio hacienda publica del ",
    "m hac.inf.pub": "ministerio hacienda publica del ",
    "m hac.inf pub": "ministerio hacienda publica del ",
    "m hacinfpub": "ministerio hacienda publica del ",
    "m jefe gabinete ministros": "ministerio jefe gabinete ministros del ",
    "m p": "ministerio desarrollo productivo del ",
    "m sa": "ministerio salud del ",
    "m turismo": "ministerio turismo del ",
    "m seguridad": "ministerio seguridad del ",
    "se ambiente des sus": "secretaria ambiente desarrollo sustentable del ",
    "se act logisticas": "secretaria actividades logisticas del ",
    "se comunicacion": "secretaria comunicacion del ",
    "se deporte": "secretaria deporte del ",
    "se d": "secretaria deporte del ",
    "sg gobernacion": "secretaria general gobernacion del ",
    "se gral gob": "secretaria general gobernacion del ",
  };

  const abreviaturas = {
    "caps": "centro salud sala salita caps ",
    "colegio": "escuela colegio educativo ",
    "centro periferico": "centro salud sala salita caps ",
    "centro salud": "centro salud sala salita caps ",
    "cid": "centro integral desarrollo ",
    "cipe": "cipe centro emision ",
    "ctro": "centro ",
    "consultorio periferico": "centro salud sala salita caps ",
    "esc": "escuela colegio educativo ",
    "gral": "general ",
    "hrc": "hospital ramon carrillo ",
    "ifdc": "instituto formacion docente continua ",
    "mediaciã³n": "mediacion ",
    "penitenciario": "penitenciario penitenciaria ",
    "prog": "programa ",
    "sgelt": "secretaria estado legal tecnica ",
    "sist": "sistema ",
    "sl": " ",
    "ulp": "ulp universidad punta ",
    "perã³n": "peron ",
    "sempro": "sempro emergencia ambulancia ",
    "vm": " ",
  };

  const articulos = ["el", "la", "los", "las", "de", "del", "y", "con"];
  const signos = ["-", "/"];
  let comisarias = {
    "5501" : " primera 5501 ",
    "5502" : " segunda 5502 ",
    "5503" : " tercera 5503 ",
    "5504" : " cuarta 5504 ",
    "5525" : " quinta 5525 ",
    "5505" : " sexta 5505 ",
    "5506" : " septima 5506 ",
    "5902" : " octava 5902 ",
    "5903" : " novena 5903 ",
    "5904" : " decima 5904 ",
  };

  let palabrasSingulares = {
    "barranca colorada" : " barranca barrancas colorada coloradas ",
    "docente" : " docente docentes ",
    "docentes" : " docente docentes ",
  };
  

    try {
        // Obtener datos de todas las columnas
        const datosNombre = obtenerDatosEnMinusculas(hojaInstituciones, FILA_MIN, COL_NOMBRE);
        const datosLocalidades = obtenerDatosEnMinusculas(hojaInstituciones, FILA_MIN, COL_LOCALIDAD);
        const datosInternos = obtenerDatosEnMinusculas(hojaInstituciones, FILA_MIN, COL_INTERNOS);

        //Concatenar datos con Localidades e insertar datos actualizados
        const datosConcatenados = concatenarColumnas(datosNombre, datosLocalidades, datosInternos);
        actualizarValoresEnHoja(datosConcatenados, FILA_MIN, COL_NOMBRE);

        for (let i = 0; i < datosConcatenados.length; i++) {
            let texto = datosConcatenados[i][0];

            //Remplazo de Articulos
            for (let articulo of articulos) {
                texto = texto.replace(new RegExp(`(^|\\s+)${articulo}(\\s|$|\\b)`, 'g'), ' ');
            }

            //Remplazo de Signos
            for (let signo of signos) {
                texto = texto.replace(new RegExp(`\\s*${signo}\\s*`, 'g'), ' ');
            }

            //Remplazo de Ministerios
            for (let ministerio in ministerios) {
                const regexQuitar = new RegExp(`\\b${ministerio}\\b(?!\\s*(mesa|privada|despacho))`, 'g');
                const regexRemplazar = new RegExp(`\\b${ministerio}\\b`, 'g');

                if (texto.match(regexQuitar)) {
                    texto = texto.replace(regexQuitar, '');
                } else {
                    texto = texto.replace(regexRemplazar, ministerios[ministerio]);
                }
            }

            //Remplazo de abreviaturas
            for (let abreviatura in abreviaturas) {
                const regex = new RegExp('\\b' + abreviatura + '\\b', 'g');
                texto = texto.replace(regex, abreviaturas[abreviatura]);
            }

            // Verificar si la clave del objeto comisarias coincide con el interno
            for (let comisaria in comisarias) {
                const regex = new RegExp('\\b' + comisaria + '\\b', 'g');
                texto = texto.replace(regex, comisarias[comisaria]);
            }

            // Singulares - Plurales
            for (let palabra in palabrasSingulares) {
                const regex = new RegExp('\\b' + palabra + '\\b', 'g');
                texto = texto.replace(regex, palabrasSingulares[palabra]);
            }

            // Aplicar función normalizarTexto
            if (typeof texto === 'string') {
              texto = normalizarTexto(texto);
            } 

            // Actualizar datosNombre[i][0] con el texto normalizado
            datosConcatenados[i][0] = texto;
        } // Fin primer For Principal 

        //Valores unicos en nombres
        const datosUnicos = eliminarPalabrasDuplicadas(datosConcatenados);
        actualizarValoresEnHoja(datosUnicos, FILA_MIN, COL_NOMBRE);

        buscarPrioridad();

        //Browser.msgBox(MENSAJE_CONFIRMACION);
    } catch (error) {
        Logger.log(error.message);
        //SpreadsheetApp.getUi().alert("Error al normalizar nombres: " + error.message);
    }

}

// 4 - Crear Hoja Data
function crearHojaData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOriginal = spreadsheet.getSheetByName("Instituciones");
  
  try {
    // Verificar que la hoja original tenga al menos una fila de datos
    if (hojaOriginal.getLastRow() < 2) {
      throw new Error("La hoja original no tiene suficientes filas de datos.");
    }
    
    const datos = hojaOriginal.getDataRange().getValues();
    
    // Verificar si la hoja "Data" ya existe
    let hojaData = spreadsheet.getSheetByName("Data");
    if (!hojaData) {
      // Crear la hoja "Data" si no existe
      hojaData = spreadsheet.insertSheet("Data");
    }
    
    // Insertar la cabecera
    const cabecera = ["prioridad", "interno", "institucion"];
    hojaData.appendRow(cabecera);
    
    // Procesar los datos
    const datosProcesados = [];
    for (let i = 1; i < datos.length; i++) {
      const fila = [datos[i][5], datos[i][3], datos[i][1]];
      datosProcesados.push(fila);
    }
    
    // Insertar los datos procesados
    const rango = hojaData.getRange(2, 1, datosProcesados.length, datosProcesados[0].length);
    rango.setValues(datosProcesados);
    
    // Estilos de la hoja 
    miEstilo();
  } catch (error) {
    SpreadsheetApp.getUi().alert("Error al crear la hoja 'Data': " + error.message);
  }
}