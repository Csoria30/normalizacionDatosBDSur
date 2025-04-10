function onOpen() {
    crearMenu();
}

function separarDatos() {
    const MSG_INICIO = "Estamos trabajando, por favor espere.";
    const MSG_FIN = "Listo, los datos los encontrara en la hoja Data";
    var respuesta = Browser.msgBox("Normalizar Datos", "¿Estás seguro de continuar?", Browser.Buttons.YES_NO);



    if (respuesta == "yes") {
        Browser.msgBox(MSG_INICIO);
        formatoInstituciones();
        separarValores();
        eliminarFilas();
        Utilities.sleep(5000);
        depurarDatos();
        crearHojaData();
        Browser.msgBox(MSG_FIN);
    }
}

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
            if (internosEliminar.includes(valorD)) {
                filasEliminar.push(filaActual); // Agrega el número de fila a filasEliminar
            }

            if (valoresDuplicados.indexOf(valorD) != -1) {
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

    const datosReemplazo = [
        {
            tipo: 'bibliotecas',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'educativos',
            datos: [
                { clave: "", valor: "" },
                { clave: "esc", valor: "escuela colegio educativo" },
                { clave: "escuela", valor: "escuela colegio educativo" },
            ]
        },
        {
            tipo: 'entre clases',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'gobernacion',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'entes de gobierno',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'militar',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'planta potabilizadora',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'policia',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'salud',
            datos: [
                { clave: "", valor: "" },
            ]
        },
        {
            tipo: 'terrazas del portezuelo',
            datos: [
                { clave: "", valor: "" },
            ]
        },

    ];

    try {
        // Obtener datos de todas las columnas
        const datosNombre = obtenerDatosEnMinusculas(hojaInstituciones, FILA_MIN, COL_NOMBRE);
        const datosLocalidades = obtenerDatosEnMinusculas(hojaInstituciones, FILA_MIN, COL_LOCALIDAD);
        const datosInternos = obtenerDatosEnMinusculas(hojaInstituciones, FILA_MIN, COL_INTERNOS);
        const datosTipoOrganismo = obtenerDatosEnMinusculas(hojaInstituciones, FILA_MIN, COL_ORGANISMO);

        // Objeto principal - Informacion de la fila 
        const objDatos = datosTipoOrganismo.map((valor, indice) => {
            return {
                tipoOrganismo: valor[0],
                localidad: datosLocalidades[indice][0],
                interno: datosInternos[indice][0],
                nombre: datosNombre[indice][0],
                texto: `${datosNombre[indice][0]} ${datosLocalidades[indice][0]} ${datosInternos[indice][0]}`
            }
        });


        objDatos.forEach((fila) => {
            let texto = fila.texto; // Nombre del organismo
            let tipoDeOrganismo = fila.tipoOrganismo; // Tipo de organismo

            // Aplicar función normalizarTexto
            texto = normalizarTexto(texto);
            tipoDeOrganismo = normalizarTexto(tipoDeOrganismo);

            // Buscar el objeto de reemplazo correspondiente al tipo de organismo
            const reemplazo = datosReemplazo.find((reemplazo) => reemplazo.tipo === tipoDeOrganismo);
            if (reemplazo) {
                // Reemplazar palabras utilizando el objeto de reemplazo
                reemplazo.datos.forEach((dato) => {
                    const palabras = texto.split(' ');
                    palabras.forEach((palabra, indice) => {
                        if (palabra.toLowerCase() === dato.clave.toLowerCase()) {
                            palabras[indice] = dato.valor;
                        }
                    });
                    texto = palabras.join(' ');
                });
            }



            // Actualizar datosNombre[i][0] con el texto normalizado
            fila.texto = texto;
        })

        // Valores unicos en nombres
        const datosUnicos = eliminarPalabrasDuplicadas(objDatos.map((fila) => fila.texto));

        actualizarValoresEnHoja(datosUnicos, FILA_MIN, COL_NOMBRE);

        buscarPrioridad();
        formatoInstituciones();
    } catch (error) {
        Logger.log(error.message);
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