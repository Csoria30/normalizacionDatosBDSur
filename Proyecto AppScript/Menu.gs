function crearMenu() {
    const MENU_PRINCIPAL = 'Mis funciones';
    const ITEM_SEPARAR_INTERNOS = 'Separar y Eliminar Internos';
    const ITEM_RENOMBRAR_REGISTROS = 'Renombrar Registros';
    const ITEM_NEW_DATA = 'Generar Hoja Data';
    const ITEM_FORMATO_CONDICIONAL = 'Aplicar Formato Condicional';

    // Creación de Menu - submenús
    var menu = SpreadsheetApp.getUi().createMenu(MENU_PRINCIPAL);
    
    menu.addItem(ITEM_SEPARAR_INTERNOS, 'separarValores');
    menu.addItem(ITEM_RENOMBRAR_REGISTROS, 'depurarDatos');
    menu.addItem(ITEM_NEW_DATA, 'crearHojaData');
    menu.addItem(ITEM_FORMATO_CONDICIONAL, 'aplicarFormatoCondicional');

    // Adición del menú a la interfaz de usuario
    menu.addToUi();
}