import { leerDatos, exportarAJson } from './helpers.js'

//Get datos
const data = await leerDatos();

exportarAJson(data, 'Data.json');

