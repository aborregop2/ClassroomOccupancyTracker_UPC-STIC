//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//IMPORTANT: Aquest script ha de ser executat des de Google Sheets a Google Apps Script per a que pugui funcionar correctament.//
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const llistaAules = [                  // Llista d'aules
    'VGA105',
    'VGA108',
    'VGA109',
    'VGA111',
    'VGA114',
];

const nomFullaFranjas = 'horari tic'; // Nom de la fulla a pintar 
var filaDeInici = 1;                  // Fila inicial per la que comença a pintar
var columnaDeInici = 'A';             // Columna inicial per la que comença a pintar

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Funció principal per processar i pintar ocupacio de les aules per s1 i s2
function processarOcupacioAules() {
  var baseUrl = 'https://web4.epsevg.upc.edu/app/ocupacioespais/index.php?aula=';

  // Obtenir la data actual i la data a 7 dies
  var dataActual = new Date();
  var dataFutura = new Date();
  dataFutura.setDate(dataActual.getDate() + 7);

  var opcionsData = { day: '2-digit', month: '2-digit', year: 'numeric' };
  var dataActualStr = dataActual.toLocaleDateString('es-ES', opcionsData).split('/').reverse().join('/');
  var dataFuturaStr = dataFutura.toLocaleDateString('es-ES', opcionsData).split('/').reverse().join('/');

  llistaAules.forEach(function(aula) {
    Logger.log(`Processant aula: ${aula}`);

    var urlActual = baseUrl + aula + '&data=' + dataActualStr;
    var urlFutura = baseUrl + aula + '&data=' + dataFuturaStr;

    // Extreure dades per a la data actual
    var nomFullaActual = extreureICrearFulla(urlActual);

    // Extreure dades per a la data a 7 dies
    var nomFullaFutura = extreureICrearFulla(urlFutura);

    // Comparar les dues fulles
    var nomFullaComparacio = compararFulles(nomFullaActual, nomFullaFutura);

    // Pintar franjes horàries a la fulla de comparació
    marcarFranjasHoraries(nomFullaComparacio, filaDeInici, columnaDeInici);

    // Borrar les fulles originals
    eliminarFullaPerNom(nomFullaActual);
    eliminarFullaPerNom(nomFullaFutura);

    // Borrar la fulla de comparació
    eliminarFullaPerNom(nomFullaComparacio);

    filaDeInici++;
  });
}

// Funció per extreure i crear una nova fulla
function extreureICrearFulla(url) {
  var resposta = UrlFetchApp.fetch(url);
  var html = resposta.getContentText();

  // Buscar el nom de l'aula en el span amb l'estil específic
  var aulaMatch = html.match(/<span[^>]*style=["']color:#f49114;["'][^>]*>([^<]+)<\/span>/i);
  var nomAula = aulaMatch ? aulaMatch[1].trim() : 'Nova Fulla';

  // Determinar si té classes s1 o s2
  var teS1 = /<span[^>]*class=["']s1["']/.test(html);
  var teS2 = /<span[^>]*class=["']s2["']/.test(html);
  nomAula += teS1 ? ' s1' : (teS2 ? ' s2' : ' s1/2');

  // Verificar si el nom de la fulla ja existeix i modificar el nom si és necessari
  var fulla = SpreadsheetApp.getActiveSpreadsheet();
  var nomFulla = nomAula;
  var contador = 1;

  while (fulla.getSheetByName(nomFulla)) {
    nomFulla = nomAula + ' ' + contador++;
  }

  // Crear una nova fulla amb el nom de l'aula
  var novaFulla = fulla.insertSheet(nomFulla);

  // Extreure la taula HTML
  var taulaHtml = html.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
  if (!taulaHtml) {
    console.log('Pàgina incorrecta :)');
    return 'Error'; // Retorna un valor d'error si no s'ha pogut extreure la taula
  }

  // Processar les files i cel·les de la taula
  var files = taulaHtml[0].match(/<tr[^>]*>([\s\S]*?)<\/tr>/gi);
  var dades = files.map(function(file) {
    return (file.match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi) || []).map(function(cel) {
      return cel.replace(/<\/?[^>]+(>|$)/g, "").trim();
    });
  });

  // Col·locar les dades a la nova fulla
  if (dades.length > 0) {
    novaFulla.getRange(1, 1, dades.length, dades[0].length).setValues(dades);
  }

  return nomFulla; // Retorna el nom de la fulla creada
}

// Funció per comparar dues fulles
function compararFulles(nomFulla1, nomFulla2) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fulla1 = ss.getSheetByName(nomFulla1);
  var fulla2 = ss.getSheetByName(nomFulla2);

  if (!fulla1 || !fulla2) {
    Logger.log('Una o ambdues fulles no existeixen.');
    return;
  }

  var dades1 = fulla1.getDataRange().getValues();
  var dades2 = fulla2.getDataRange().getValues();
  
  // Determinar la mida mínima comuna per a les comparacions
  var numFiles = Math.min(dades1.length, dades2.length);
  var numCols = Math.min(dades1[0].length, dades2[0].length);

  // Crear una nova fulla per a la comparació
  var nomFullaComparacio = 'FullaComparacio';
  var fullaComparacio = ss.getSheetByName(nomFullaComparacio) || ss.insertSheet(nomFullaComparacio);

  // Netejar contingut previ a la fulla de comparació
  fullaComparacio.clear();

  var rang = fullaComparacio.getRange(1, 1, numFiles, numCols);
  var valors = [];
  var colorsFons = [];

  for (var i = 0; i < numFiles; i++) {
    var valorsFila = [];
    var colorsFila = [];

    for (var j = 0; j < numCols; j++) {
      var valor1 = dades1[i][j];
      var valor2 = dades2[i][j];
      
      valorsFila.push(valor1);

      // Definir color basat en la comparació de cel·les
      if (!valor1 && !valor2) {
        colorsFila.push('#93c47d');  // Verd
      } else if (!valor1 && valor2) {
        colorsFila.push(nomFulla1.includes("s1") ? '#ffd966' : '#c27ba0');  // Groc o Morat
      } else if (valor1 && !valor2) {
        colorsFila.push(nomFulla1.includes("s1") ? '#c27ba0' : '#ffd966');  // Morat o Groc
      } else {
        colorsFila.push('#FFFFFF');  // Blanc
      }
    }
    valors.push(valorsFila);
    colorsFons.push(colorsFila);
  }

  rang.setValues(valors);
  rang.setBackgrounds(colorsFons);

  return nomFullaComparacio;
}

// Funció per marcar les franjes horàries
function marcarFranjasHoraries(nomFullaComparacio, fila, columna) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fulla = ss.getSheetByName(nomFullaComparacio);
  if (!fulla) {
    throw new Error(`La fulla "${nomFullaComparacio}" no existeix.`);
  }

  const franjas = [
    { start: 2, end: 6, cells: [getColumnLetter(columna, 0) + fila, getColumnLetter(columna, 6) + fila, getColumnLetter(columna, 12) + fila, getColumnLetter(columna, 18) + fila, getColumnLetter(columna, 24) + fila] }, // 8:30 a 10:30
    { start: 7, end: 10, cells: [getColumnLetter(columna, 1) + fila, getColumnLetter(columna, 7) + fila, getColumnLetter(columna, 13) + fila, getColumnLetter(columna, 19) + fila, getColumnLetter(columna, 25) + fila] }, // 10:30 a 12:30
    { start: 11, end: 14, cells: [getColumnLetter(columna, 2) + fila, getColumnLetter(columna, 8) + fila, getColumnLetter(columna, 14) + fila, getColumnLetter(columna, 20) + fila, getColumnLetter(columna, 26) + fila] }, // 12:30 a 14:30
    { start: 15, end: 19, cells: [getColumnLetter(columna, 3) + fila, getColumnLetter(columna, 9) + fila, getColumnLetter(columna, 15) + fila, getColumnLetter(columna, 21) + fila, getColumnLetter(columna, 27) + fila] }, // 15:00 a 17:00
    { start: 20, end: 23, cells: [getColumnLetter(columna, 4) + fila, getColumnLetter(columna, 10) + fila, getColumnLetter(columna, 16) + fila, getColumnLetter(columna, 22) + fila, getColumnLetter(columna, 28) + fila] }, // 17:00 a 19:00
    { start: 24, end: 27, cells: [getColumnLetter(columna, 5) + fila, getColumnLetter(columna, 11) + fila, getColumnLetter(columna, 17) + fila, getColumnLetter(columna, 23) + fila, getColumnLetter(columna, 29) + fila] } // 19:00 a 21:00
  ];

  let novaFulla = ss.getSheetByName(nomFullaFranjas);
  if (!novaFulla) {
    novaFulla = ss.insertSheet(nomFullaFranjas);
  }

  const colorsToSet = [];
  franjas.forEach((franja) => {
    const columnes = ['B', 'C', 'D', 'E', 'F']; // Dilluns a Divendres
    columnes.forEach((columna, colIndex) => {
      try {
        const rang = fulla.getRange(`${columna}${franja.start}:${columna}${franja.end}`);
        const colors = rang.getBackgrounds();

        const comptaColors = {};
        colors.forEach(row => row.forEach(color => {
          if (color) {
            comptaColors[color] = (comptaColors[color] || 0) + 1;
          }
        }));

        let maxColor = '';
        let maxCount = 0;
        for (const [color, count] of Object.entries(comptaColors)) {
          if (count > maxCount) {
            maxColor = color;
            maxCount = count;
          }
        }

        colorsToSet.push({ cell: franja.cells[colIndex], color: maxColor && maxColor !== '#ffffff' && maxColor !== 'white' ? maxColor : '#ffffff' });
      } catch (e) {
        Logger.log(`Error al processar el rang ${columna}${franja.start}:${columna}${franja.end} - ${e.message}`);
      }
    });
  });

  // Establir fons en grup
  colorsToSet.forEach(({ cell, color }) => {
    novaFulla.getRange(cell).setBackground(color);
  });
}

// Funció per eliminar una fulla per nom
function eliminarFullaPerNom(nomFulla) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fulla = ss.getSheetByName(nomFulla);
  if (fulla) {
    ss.deleteSheet(fulla);
  }
}

// Funció per obtenir la lletra de columna
function getColumnLetter(baseCol, index) {
  let baseCode = baseCol.charCodeAt(0) - 65; // Convertir 'A' a 0, 'B' a 1, etc.
  let colIndex = baseCode + index;

  let result = '';
  while (colIndex >= 0) {
    result = String.fromCharCode((colIndex % 26) + 65) + result;
    colIndex = Math.floor(colIndex / 26) - 1;
  }

  return result;
}