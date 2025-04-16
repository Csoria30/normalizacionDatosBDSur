import XLSX from 'xlsx';
import fs from 'fs/promises';

//! Obteniendo los registros
async function leerDatos() {
    try {
        const nombreLibro = './InstitucionReporte.xls';
        const contenido = await fs.readFile(nombreLibro);
        const libro = XLSX.read(contenido);
        const hoja = libro.SheetNames[0];

        const data = XLSX.utils.sheet_to_json(libro.Sheets[hoja], {
            header: [
                "pop",
                "organismo",
                "nombre",
                "email",
                "direccion",
                "localidad",
                "departamento",
                "nInterno",
                "tecnologia",
                "telefono",
                "observacion",
                "anotaciones"
            ]
        });

        const datosDesdeFila4 = data.slice(3); // Elimina las primeras 3 filas

        //* Mofidicacion de Data
        const data_eliminandoPropiedades = await eliminandoPropiedades(datosDesdeFila4); //Quitando prioridades
        const data_agregandoPrioridad = await agregandoPrioridad(data_eliminandoPropiedades); //Agregando prioridad
        const data_eliminandoOrganismos = await eliminandoOrganismos(data_agregandoPrioridad); //Quitando organismos
        const data_valorAMinuscula = await valorAMinuscula(data_eliminandoOrganismos); //Valores a miunsculas
        const data_separandoValores = await separandoValores(data_valorAMinuscula); // Separa internos
        const data_eliminacionPorNombre = await eliminacionPorNombre(data_separandoValores); // Eliminacion nombre
        const data_depurarNombre = await depurarNombre(data_eliminacionPorNombre); //Depura nombre 
        const data_remplazoEnNombre = await remplazoEnNombre(data_depurarNombre);
        const data_concatenarLocalidad = await concatenarLocalidad(data_remplazoEnNombre);
        const data_eliminandoObs = await eliminandoObs(data_concatenarLocalidad);

        
        return data_eliminandoObs;
    } catch (error) {
        console.log(`Error al leer el archivo Excel: ${error}`);
    }
}

//! Modificacion de Datos
// 2 - Modificacion del objeto data
async function eliminandoPropiedades(data) {
    try {

        const propiedadesAEliminar = [
            "pop",
            "email",
            "departamento",
            "direccion",
            "tecnologia",
            "telefono"
        ];

        const datosModificados = data.map(objeto => {
            propiedadesAEliminar.forEach(propiedad => {
                if (objeto.hasOwnProperty(propiedad)) {
                    delete objeto[propiedad];
                }
            });
            return objeto;
        });

        return datosModificados;

    } catch (error) {
        console.log(`Error al eliminar propiedades: ${error}`);
    }
}

// 3 - Agregando la propieda prioridad
async function agregandoPrioridad(data) {
    try {
        const prioridadMap = {
            'prioridad1': '*****',
            'prioridad2': '****',
            'prioridad3': '***',
            'prioridad4': '**',
            'prioridad5': '*'
        };

        const datosConPrioridad = data.map(objeto => {
            const prioridadMatch = objeto.observacion?.match(/prioridad\d+/);
            const prioridad = prioridadMatch ? prioridadMap[prioridadMatch[0]] : null;
            return { ...objeto, prioridad };
        });

        return datosConPrioridad;

    } catch (error) {
        console.log(`Error al agregar prioridades: ${error}`);
    }
}

// 4 - Eliminando Organismos
async function eliminandoOrganismos(data) {
    try {

        const organismosAExcluir = [
            'Privados AUI ( Serv. Dedicado )',
            'Cámaras CCC',
            'Privados',
            'Otros',
            'Alarmas comunitarias',
            'Cybers',
            'Confidencial',
            'Planta potabilizadora',
            'Privados AVL'
        ];

        const datosFiltrados = data.filter(objeto => !organismosAExcluir.includes(objeto.organismo));

        return datosFiltrados;

    } catch (error) {
        console.log(`Error al eliminar organismos: ${error}`);
    }
}

// 5 - Convirtiendo data a minusculas
async function valorAMinuscula(data) {
    try {

        const datosEnMinusculas = data.map(objeto => {
            return Object.fromEntries(
                Object.entries(objeto).map(([clave, valor]) => {
                    if (typeof valor === 'string') {
                        return [clave, valor.toLowerCase()];
                    } else {
                        return [clave, valor];
                    }
                })
            );
        });

        return datosEnMinusculas;

    } catch (error) {
        console.log(`Error al transformar en minuscula: ${error}`);
    }
}

// 6 - Separa registros que tengan dos internos 
async function separandoValores(data) {
    try {

        const datosModificados = data.flatMap(objeto => {
            const vozIps = objeto.nInterno ? objeto.nInterno.split(/[-/]+/).map(vozIp => vozIp.trim()) : [];
            return vozIps.map(vozIp => ({ ...objeto, nInterno: vozIp }));
        });

        return datosModificados;

    } catch (error) {
        console.log(`Error al separar internos: ${error}`);
    }
}

// 7 - Eliminacion registros nombres no validos
async function eliminacionPorNombre(data) {
    try {

        const nombresExcluir = ['data center', 'verificar', 'shelter', 'sd', 'no hay area asignada', 'no registra area', 'no responde'];
        const observacionesExcluir = ['no transferir', 'sin servicio'];

        const datosFiltrados = data.filter(objeto =>
            !nombresExcluir.some(palabra => objeto.nombre.toLowerCase().includes(palabra.toLowerCase())) &&
            !(objeto.observacion && observacionesExcluir.some(palabra => objeto.observacion.toLowerCase().includes(palabra.toLowerCase())))
        );

        return datosFiltrados;

    } catch (error) {
        console.log(`Error al eliminar por nombre: ${error}`);
    }
}

// 8 - Depuracion de data nombre
async function depurarNombre(data) {
    try {

        const dataDepurada = data.map(objeto => {
            const objetoDepurado = {};

            /*
                Normaliza informacion del objeto
                Quita acentos
                Deja un espacio entre palabras 
                Quita simbolos no permitidos: / , //, -, 
            */
            Object.keys(objeto).forEach(propiedad => {
                if (typeof objeto[propiedad] === 'string') {
                    objetoDepurado[propiedad] = objeto[propiedad]
                        .normalize("NFD")
                        .replace(/[\u0300-\u036f]/g, "")
                        .replace(/[\/\/-]/g, ' ')
                        .replace(/\s+/g, ' ')
                        .trim();
                } else {
                    objetoDepurado[propiedad] = objeto[propiedad];
                }
            });
            return objetoDepurado;
        });


        return dataDepurada;

    } catch (error) {
        console.log(`Error al depurar nombres: ${error}`);
    }
}

// 9 - Remplazo de nombres en data
async function remplazoEnNombre(data) {

    try {

        //* Objetos de validacion
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

            //! Policia
            { clave: "complejo provincial penitenciario 1", valor: "complejo provincial servicio penitenciario 1" },

            //! Otros
            { clave: "vice gobernacion", valor: "vicegobernacion vice gobernacion" }, //* Tambien posee regla por organismo 'gobernacion'
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

        //* Bucle remplazo data
        const dataConReemplazos = data.map(objeto => {
            let nombre = objeto.nombre;

            // Aplicar reglas de reemplazo generales
            datosReemplazoGenerales.forEach((dato) => {
                const regex = new RegExp(`\\b${dato.clave}\\b`, 'gi');
                nombre = nombre.replace(regex, dato.valor);
            });

            // Aplicar reemplazos por string
            reemplazosPorString.forEach(dato => {
                const regex = new RegExp(`\\b${dato.clave}\\b`, 'gi');
                nombre = nombre.replace(regex, dato.valor);
            });

            // Buscar el objeto de reemplazo correspondiente al tipo de organismo
            const reemplazo = datosReemplazo.find((reemplazo) => reemplazo.tipo === objeto.organismo);
            if (reemplazo) {
                // Reemplazar palabras utilizando el objeto de reemplazo
                reemplazo.datos.forEach((dato) => {
                    const palabras = nombre.split(' ');
                    palabras.forEach((palabra, indice) => {
                        if (palabra.toLowerCase() === dato.clave.toLowerCase()) {
                            palabras[indice] = dato.valor;
                        }
                    });
                    nombre = palabras.join(' ');
                });
            }

            // Dejar palabras únicas
            const palabras = nombre.split(' ');
            const palabrasUnicas = [...new Set(palabras.map(palabra => palabra.toLowerCase()))];
            nombre = palabrasUnicas.join(' ');


            return { ...objeto, nombre: nombre };
        });

        return dataConReemplazos;

    } catch (error) {
        console.log(`Error al renombrar los registros: ${error}`);
    }
}

// 10 - Concatenar localidad
async function concatenarLocalidad(data) {
    try {

        const dataConLocalidad = data.map(objeto => ({
            ...objeto,
            nombre: `${objeto.nombre} (${objeto.localidad})`
        }));

        return dataConLocalidad;

    } catch (error) {
        console.log(`Error al concatenar localidades: ${error}`);
    }
}

// 11 - Eliminando Obs post validaciones
async function eliminandoObs(data) {
    try {

        const propiedadesAEliminar = ["observacion"];

        const datosModificados = data.map(objeto => {
            propiedadesAEliminar.forEach(propiedad => {
                if (objeto.hasOwnProperty(propiedad)) {
                    delete objeto[propiedad];
                }
            });
            return objeto;
        });

        return datosModificados;

    } catch (error) {
        console.log(`Error al eliminar propiedades: ${error}`);
    }
}

//* Funciones Helpers
async function exportarAJson(datos, nombreArchivo) {
    try {
        const json = JSON.stringify(datos, null, 2);
        await fs.writeFile(nombreArchivo, json, 'utf8');
        console.log(`Archivo ${nombreArchivo} creado con éxito`);
    } catch (err) {
        console.error(`Error al escribir archivo ${nombreArchivo}:`, err);
    }
}


export {
    exportarAJson,
    leerDatos
}