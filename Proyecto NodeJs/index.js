import XLSX from 'xlsx';
import fs from 'fs/promises';
import { exportarAJson } from './helpers.js';

async function leerDatos() {
    try {
        const nombreLibro = './InstitucionReporte.xls';
        const contenido = await fs.readFile(nombreLibro);
        const libro = XLSX.read(contenido);
        const hoja = libro.SheetNames[0];

        const data = XLSX.utils.sheet_to_json(libro.Sheets[hoja], {
            header: [
                "Pop",
                "Organismo",
                "Nombre",
                "Email",
                "Direccion",
                "Localidad",
                "Departamento",
                "nInterno",
                "Tecnologia",
                "Telefono",
                "Observacion",
                "Anotaciones"
            ]
        });

        const datosDesdeFila4 = data.slice(3); // Elimina las primeras 3 filas

        return datosDesdeFila4;
    } catch (error) {
        console.log(`Error al leer el archivo Excel: ${error}`);
    }
}

async function eliminandoPropiedades() {
    try {
        const data = await leerDatos();

        const propiedadesAEliminar = [
            "Pop",
            "Email",
            "Departamento",
            "Direccion",
            "Tecnologia",
            "Telefono"
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

async function agregandoPrioridad() {
    try{
        const data = await eliminandoPropiedades();
        const datosConPrioridad = data.map(objeto => ({ ...objeto, prioridad: null }));
        return datosConPrioridad;
    }catch(error){
        console.log(`Error al agregar prioridades: ${error}`);
    }
}

async function eliminandoOrganismos() {
    try {
        const data = await agregandoPrioridad();
        
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
    
        const datosFiltrados = data.filter(objeto => !organismosAExcluir.includes(objeto.Organismo));
    
        return datosFiltrados;

    } catch (error) {
        console.log(`Error al eliminar organismos: ${error}`);
    }
}

async function valorAMinuscula() {
    try {
        
        const data = await eliminandoOrganismos();
    
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

async function separandoValores() {
    try {
        const data = await valorAMinuscula();
    
        const datosModificados = data.flatMap(objeto => {
            const vozIps = objeto.nInterno ? objeto.nInterno.split(/[-/]+/).map(vozIp => vozIp.trim()) : [];
            return vozIps.map(vozIp => ({ ...objeto, nInterno: vozIp }));
        });
    
        return datosModificados;

    } catch (error) {
        console.log(`Error al separar internos: ${error}`);
    }
}

async function eliminacionPorNombre() {
    try {
        
        const data = await separandoValores();
    
        const nombresExcluir = ['data center', 'verificar', 'shelter', 'sd', 'no hay area asignada', 'no registra area', 'no responde'];
        const observacionesExcluir = ['no transferir', 'sin servicio'];
    
        const datosFiltrados = data.filter(objeto =>
            !nombresExcluir.some(palabra => objeto.Nombre.toLowerCase().includes(palabra.toLowerCase())) &&
            !(objeto.Observacion && observacionesExcluir.some(palabra => objeto.Observacion.toLowerCase().includes(palabra.toLowerCase())))
        );
    
        return datosFiltrados;

    } catch (error) {
        console.log(`Error al eliminar por nombre: ${error}`);
    }
}

async function depurarNombre() {
    try {
        const data = await eliminacionPorNombre();
    
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

async function remplazoEnNombre() {

    try {
        const data = await depurarNombre();
        
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
            let nombre = objeto.Nombre;
    
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
            const reemplazo = datosReemplazo.find((reemplazo) => reemplazo.tipo === objeto.Organismo);
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
    
    
            return { ...objeto, Nombre: nombre };
        });
    
        return dataConReemplazos;

    } catch (error) {
        console.log(`Error al renombrar los registros: ${error}`);
    }
}

async function concatenarLocalidad() {
    try {
        const data = await remplazoEnNombre();    
        return data;
    } catch (error) {
        console.log(`Error al concatenar localidades: ${error}`);
    }
}

const data = await eliminandoOrganismos();
exportarAJson(data, 'Data.json');