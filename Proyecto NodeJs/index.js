import XLSX from 'xlsx';
import fs from 'fs/promises';

async function leerDatos(){
    try {
        const nombreLibro = './InstitucionReporte.xls';
        const contenido = await fs.readFile(nombreLibro);
        const libro = XLSX.read(contenido);
        const hoja = libro.SheetNames[0];

        const data = XLSX.utils.sheet_to_json(libro.Sheets[hoja],{
            header:["Pop", "Organismo", "Nombre", "Email", "Direccion", "Localidad", "Departamento", "N_VozIP", "Tecnologia", "Telefono", "Observacion", "Anotaciones"]
        });

        const datosDesdeFila4 = data.slice(3); // Elimina las primeras 3 filas

        return datosDesdeFila4;
    }catch (error) {
        console.error('Error al leer archivo Excel:', error);
    }
}
    

