import axios from 'axios';
import dotenv from 'dotenv';
import { leerDatos } from './helpers.js'

//Get datos
const data = await leerDatos();

//Variables env
dotenv.config();
const dbHost = process.env.DB_HOST;
const dbPort = process.env.DB_PORT;
const url =  `${dbHost}:${dbPort}/api/cargar-instituciones-sur`;



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