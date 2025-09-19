function importarCSV() {
  const props = PropertiesService.getScriptProperties();
  const url   = props.getProperty('KOBO_URL');
  const token = props.getProperty('KOBO_TOKEN');

  var sheetName = "DatosImportados";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  var options = {
    method: "get",
    headers: {
      Authorization: "Token " + token
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  var csvContent = response.getContentText();

  // Detectar separador automáticamente (Kobo suele usar ;)
  var delimiter = csvContent.indexOf(";") !== -1 ? ";" : ",";

  // Parsear TODO el CSV en una sola llamada
  var filas = Utilities.parseCsv(csvContent, delimiter);

  // Obtener encabezados y localizar las columnas 'start', 'end' y 'Ubicar en el mapa'
  var headers = filas[0];
  var startIndex = headers.indexOf("start");
  var endIndex = headers.indexOf("end");
  var ubicacionIndex = headers.indexOf("Ubicar en el mapa");

  if (startIndex === -1 || endIndex === -1) {
    console.warn("No se encontraron las columnas 'start' y 'end'.");
    return;
  }

  // Datos actuales
  var data = sheet.getDataRange().getValues();
  var existingPairs = new Set();

  for (var i = 1; i < data.length; i++) {
    existingPairs.add(data[i][startIndex] + "|" + data[i][endIndex]);
  }

  var nuevasFilas = [];
  var numColumnas = headers.length;

  // Filtrar filas nuevas
  for (var j = 1; j < filas.length; j++) {
    var row = filas[j];
    var pairKey = row[startIndex] + "|" + row[endIndex];

    if (!existingPairs.has(pairKey)) {
      // Rellenar si faltan columnas
      while (row.length < numColumnas) row.push("");
      while (row.length > numColumnas) row.pop();

      // Limpiar "Ubicar en el mapa"
      if (ubicacionIndex !== -1 && row[ubicacionIndex]) {
        var partes = row[ubicacionIndex].trim().split(" ");
        if (partes.length === 4 && partes[2] === "0" && partes[3] === "0") {
          row[ubicacionIndex] = partes.slice(0, 2).join(" ");
        }
      }

      nuevasFilas.push(row);
    }
  }

  // Agregar encabezados si es nuevo
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Agregar filas nuevas
  if (nuevasFilas.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, nuevasFilas.length, numColumnas).setValues(nuevasFilas);
  } else {
    console.warn("No hay nuevas filas para agregar.");
  }
}

function sanitize(v){
  if (typeof v !== 'string') return v;
  const s = v.trim();
  return /^[=+\-@]/.test(s) ? ("'"+s) : s;
}
