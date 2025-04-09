// Variables globales
var libro = SpreadsheetApp.getActiveSpreadsheet();
var hojaActiva = libro.getActiveSheet();
var hojaData = libro.getSheetByName("Data");
var hojaInstituciones = libro.getSheetByName("Instituciones");
var celdaActiva = libro.getActiveCell();
var nombreLibro = libro.getName();
var idLibro = libro.getId();
var urlLibro = libro.getUrl();

var valor = celdaActiva.getValue();
var filaActiva = celdaActiva.getRow();
var colActiva = celdaActiva.getColumn();
var ultimaFila = hojaActiva.getLastRow();

//Columnas 
const COL_ORGANISMO = 1; // Columna A
const COL_NOMBRE = 2; // Columna B
const COL_LOCALIDAD = 3; // Columna C
const COL_INTERNOS = 4; // Columna D
const COL_OBS = 5; // Columna E
const COL_PRIORIDAD = 6; // Columna F

const FILA_ACTUAL = 2;
const FILA_MIN = 2;
const ULTIMAF_INSTITUCIONES = hojaInstituciones.getLastRow();