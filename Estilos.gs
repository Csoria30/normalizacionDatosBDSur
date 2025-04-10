function miEstilo()
{
  
  let colores = {
    "*****": "#58ce74",
    "****": "#dceb5b",
    "***": "#e4b302",
    "**": "#b31d6a",
    "*": "#f21d1f"
  };

  let colorAzul = "#4180ab";
  let colorBlanco = "#FFF";
  let rango = hojaData.getRange("A1:C1");
  let ultimaColumna = hojaData.getLastColumn();
  const columnaB = hojaData.getRange(1, 2, hojaData.getLastRow(), 1);


  rango
    .setBackground (colorAzul)
    .setFontColor (colorBlanco)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontSize(13)
  
  hojaData
    .setFrozenRows(1)
  
  columnaB
    .setHorizontalAlignment("center");

  hojaData
    .deleteColumns(4, ultimaColumna);

  // Formato Condicional 
  for (var fila = 1; fila <= ultimaFila; fila++) {
    var valorColumnaA = hojaData.getRange(fila, 1).getValue();
    
    for (var prioridad in colores) {
      if (valorColumnaA === (prioridad)) {
        hojaData.getRange(fila, 1, 1, 3).setBackground(colores[prioridad]);
      }
    }
  }

  // Auto Size Columnas 
  for (let i = 1; i <= 3; i++) {
    hojaData.autoResizeColumn(i);
  }
}

function formatoInstituciones(){
  let colorAzul = "#4180ab";
  let colorBlanco = "#FFF";
  let rango =hojaInstituciones.getRange("A1:E1");

  rango
    .setBackground (colorAzul)
    .setFontColor (colorBlanco)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontSize(13)

  hojaInstituciones
    .setFrozenRows(1)
}


function aplicarFormatoCondicional() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instituciones");
  var ultimaFila = hoja.getLastRow();
  
  var colores = {
    "prioridad1": "#58ce74",
    "prioridad2": "#dceb5b",
    "prioridad3": "#e4b302",
    "prioridad4": "#b31d6a",
    "prioridad5": "#f21d1f"
  };
  
  for (var fila = 1; fila <= ultimaFila; fila++) {
    var valorColumnaE = hoja.getRange(fila, 5).getValue();
    
    for (var prioridad in colores) {
      if (valorColumnaE.includes(prioridad)) {
        hoja.getRange(fila, 1, 1, 4).setBackground(colores[prioridad]);
      }
    }
  }
}
