import XLSX from 'xlsx';
import fs from 'fs/promises';
import axios from 'axios';
import { leerDatos, exportarAJson } from './helpers.js'

const data = await leerDatos();
const url = 'http://192.168.64.93:9005/api/cargar-instituciones-sur';
exportarAJson(data, 'Data.json');

const config = {
    headers: {
        'Content-Type': 'application/json'
    }
};


axios.post(url, data, config)
    .then((response) => {
        console.log('Respuesta de la API:', response.data);
        console.log('Estado de la respuesta:', response.status);
    })
    .catch((error) => {
        console.error('Error enviando la solicitud:', error.message);
        if (error.response) {
            console.log('Datos de la respuesta:', error.response.data);
            console.log('Estado de la respuesta:', error.response.status);
        }
    });