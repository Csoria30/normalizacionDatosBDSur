import fs from 'fs/promises';

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
    exportarAJson
}