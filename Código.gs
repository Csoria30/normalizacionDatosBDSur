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

// 1 - Separar registros por medio de internos
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
    let valoresDuplicados = ['3530', '6880', '8928', '5201', '5960'];
    let internosEliminar = ['4840', '4841', '3269', '2033', '2022', '2055', '47216', '47217', '47875'];
    let valoresEncontrados = {};
    
    // Valores prohibidos en valorB
    let valoresProhibidos = ["data center", "verificar", "shelter", "sd", "no hay area asignada", "no registra area", "no responde"]; 
    let MSG_INICIO = "Estamos eliminando los registros con formato incorrecto, por favor espere.";

    try {
        while (filaActual <= datos.length) {
            let valorB = datos[filaActual - 1][1].toString().toLowerCase();
            let valorC = datos[filaActual - 1][2];
            let valorD = datos[filaActual - 1][3].toString().trim();
            let valorE = datos[filaActual - 1][4].toString().toLowerCase();

            // Verificando duplicados de internos
            
            if (valoresDuplicados.indexOf(valorD) != -1 && valoresEncontrados[valorD]) {
                filasEliminar.push(filaActual);
            }
            
            //Eliminar internos 
            if (internosEliminar.includes(valorD)) {
                filasEliminar.push(filaActual); // Agrega el número de fila a filasEliminar
            }

            if (valoresDuplicados.indexOf(valorD) != -1) {
                valoresEncontrados[valorD] = true;
            }

            // Comprueba valores prohibidos en valorB (convirtiendo a minúsculas)
            if (valoresProhibidos.some(valor => valorB.includes(valor.toLowerCase()))) {
                filasEliminar.push(filaActual);
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

    const datosReemplazoGenerales = [
      { clave: "", valor: "" },
      { clave: "de", valor: "" },
      { clave: "del", valor: "" },
      { clave: "con", valor: "" },
      { clave: "y", valor: "" },
      { clave: "los", valor: "" },
      { clave: "la", valor: "" },
      { clave: "b", valor: "" },
      { clave: "style", valor: "" },
      { clave: "gral", valor: "general" },
      { clave: "sempro", valor: "sempro emergencia ambulancia" },
      { clave: "ulp", valor: "ulp universidad punta" },
      { clave: "prog", valor: "programa" },
      { clave: "mosca", valor: "mosca moscas" },
      { clave: "frutos", valor: "fruto frutos" },
      { clave: "docente", valor: "docente docentes" },
      { clave: "cipe", valor: "cipe sipe centro emision" },
    ];

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
                { clave: "xxi", valor: "21" },
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
                { clave: "gob", valor: "gobernacion" },
                { clave: "subp", valor: "sub programa" },
                { clave: "subprogr", valor: "sub programa" },
                { clave: "vicegobernacion", valor: "vicegobernacion vice gobernacion" }, // Tambien posee regla de string 
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
                { clave: "5501", valor: "primera 5501" },
                { clave: "5502", valor: "segunda 5502" },
                { clave: "5503", valor: "tercera 5503" },
                { clave: "5504", valor: "cuarta 5504" },
                { clave: "5525", valor: "quinta 5525" },
                { clave: "5505", valor: "sexta 5505" },
                { clave: "5506", valor: "septima 5506" },
                { clave: "5902", valor: "octava 5902" },
                { clave: "5903", valor: "novena 5903" },
                { clave: "5904", valor: "decima 5904" },
            ]
        },
        {
            tipo: 'salud',
            datos: [
                { clave: "", valor: "" },
                { clave: "caps", valor: "centro salud sala salita caps" },
                { clave: "centro", valor: "centro salud sala salita caps" },
                { clave: "ctro", valor: "centro salud sala salita caps" },
                { clave: "periferico", valor: "centro salud sala salita caps" },
            ]
        },
        {
            tipo: 'terrazas del portezuelo',
            datos: [
                { clave: "", valor: "" },
                { clave: "pane", valor: "pane panes" },
            ]
        },

    ];

    const reemplazosPorString = [
      { clave: "m ciencia e inn", valor: "ministerio ciencia innovacion del" },
      { clave: "m des productivo", valor: "ministerio desarrollo productivo del" },
      { clave: "m des humano", valor: "ministerio desarrollo humano del" },
      { clave: "m e", valor: "ministerio educacion" },
      { clave: "m gobierno", valor: "ministerio gobierno del" },
      { clave: "m gob", valor: "ministerio gobierno del" },
      { clave: "m jefe gabinete", valor: "ministerio jefe gabinete" },
      { clave: "m hacienda inf pub", valor: "ministerio hacienda publica del" },
      { clave: "m hacinfpub", valor: "ministerio hacienda publica del" },
      { clave: "m p", valor: "ministerio produccion" },
      { clave: "m sa", valor: "ministerio salud" },
      { clave: "m seguridad", valor: "ministerio seguridad" },
      { clave: "m turismo", valor: "ministerio turismo" },
      { clave: "se act logisticas", valor: "secretaria actividades logisticas" },
      { clave: "se ambiente des sus", valor: "secretaria ambiente desarrollo sustentable" },
      { clave: "se comunicacion", valor: "secretaria comunicacion" },
      { clave: "m rise", valor: "relaciones institucionales seguridad" },
      { clave: "se d", valor: "secretaria desarrollo" },
      { clave: "se deporte", valor: "secretaria deporte deportes" },
      { clave: "se general gob", valor: "secretaria general gobernacion" },
      { clave: "sg gobernacion", valor: "secretaria general gobernacion" },
      { clave: "se transporte", valor: "secretaria transporte" },
      { clave: "se estado transp", valor: "secretaria transporte" },
      { clave: "se v", valor: "secretaria vivienda viviendas" },
      { clave: "sec estado general legal tec", valor: "secretaria estado legal tecnica" },
      // Policia
      { clave: "complejo provincial penitenciario 1", valor: "complejo provincial servicio penitenciario 1" },
      //Otros
      { clave: "vice gobernacion", valor: "vicegobernacion vice gobernacion" }, // Tambien posee regla por organismo 'gobernacion'

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

            // Aplicar reglas de reemplazo generales
            datosReemplazoGenerales.forEach((dato) => {
              const regex = new RegExp(`\\b${dato.clave}\\b`, 'gi');
              texto = texto.replace(regex, dato.valor);
            });

            // Aplicar reemplazos por string
            reemplazosPorString.forEach((dato) => {
              const regex = new RegExp(`\\b${dato.clave}\\b`, 'gi');
              texto = texto.replace(regex, dato.valor);
            });

            // Agrega el nombre biblioteca
            if (tipoDeOrganismo === 'bibliotecas') {
              texto = 'biblioteca ' + texto;
            }
            
            if (tipoDeOrganismo === 'terrazas del portezuelo') {
              texto = 'terrasas terrasa de del portesuelo ' + texto;
            }

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

        // Busca e inserta prioridad
        buscarPrioridad();

        // Renombra registros condicional = nInterno
        buscarYReemplazar();

        // Formato condicional a la hoja de instituciones
        formatoInstituciones();
    } catch (error) {
        Logger.log(error.message);
    }

}

// 4 - Remplazo nombres particulares
function buscarYReemplazar() {
  const valoresABuscar = [
    { valor: "4073", reemplazo: 'ministerio desarrollo humano del direccion viviendas inscripciones san luis 4073' },
    { valor: "5960", reemplazo: 'hospital central ramon carillo mesa entrada turnos san luis 5960' },
    { valor: "5201", reemplazo: 'hospital san francisco mesa entrada turnos 5201' },
    { valor: "8928", reemplazo: 'maternidad doctor carlos alberto luco mesa entrada turnos villa mercedes 8928' },
    { valor: "3530", reemplazo: 'maternidad teresita baigorria conmutador san luis mesa turnos entrada 3530' },
    { valor: "6880", reemplazo: 'terminal ediro nueva informe informes mesa entrada 6880 san luis' },
    { valor: "3206", reemplazo: 'ministerio desarrollo productivo del direccion industrial energias sustentables san luis instalacion paneles solares 3206' },
  ];

  const hoja = hojaInstituciones;
  const ultimaFila = hoja.getLastRow();
  const columnaD = hoja.getRange(1, 4, ultimaFila, 1).getValues();
  const columnaB = hoja.getRange(1, 2, ultimaFila, 1);

  valoresABuscar.forEach((valor) => {
    for (let i = 0; i < columnaD.length; i++) 
    {
      if (columnaD[i][0] === valor.valor) {
        columnaB.getCell(i + 1, 1).setValue(valor.reemplazo);
      }
    }

  });
}

// 5 - Crear Hoja Data
function crearHojaData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaOriginal = spreadsheet.getSheetByName("Instituciones");

    try {
        // Eliminar la hoja "Data" si existe
        const hojaDataExistente = spreadsheet.getSheetByName("Data");
        if (hojaDataExistente){
          spreadsheet.deleteSheet(hojaDataExistente);
          // Utilities.sleep(2000); // Agregar un retraso de 2 segundos
        }
        
        //Crea la hoja data y obtener la referencia
        const hojaData = spreadsheet.insertSheet('Data');
        
        // Verificar que la hoja original tenga al menos una fila de datos
        if (hojaOriginal.getLastRow() < 2) {
            throw new Error("La hoja original no tiene suficientes filas de datos.");
        }

        //Obteniendo informacion de istituciones
        const data = hojaOriginal.getDataRange().offset(1, 0, hojaOriginal.getLastRow() - 1, hojaOriginal.getLastColumn()).getValues();

        // Insertar la cabecera
        const cabecera = ["prioridad", "interno", "institucion"];
        hojaData.appendRow(cabecera);

        // Procesar los datos
        var resultados = data.map(function(fila) {
          return [fila[5], fila[3], fila[1]];
        });

        
        // Insertar los datos procesados
        hojaData.getRange(2, 1, resultados.length, resultados[0].length).setValues(resultados);


        // Estilos de la hoja 
        //miEstilo();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Error al crear la hoja 'Data': " + error.message);
    }
}