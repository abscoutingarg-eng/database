const CONFIG = {
  SHEET_ID: '1Ti82rOPecPgkKKQXTT8nXUGlbo_euDsOLvTSHPKI154',
  SHEET_NAME: 'Respuestas de formulario 1'
};

const COLUMNAS = {
  fecha: 'Marca temporal',
  email: 'Dirección de correo electrónico',
  nombre: 'Nombre',
  apellido: 'Apellido',
  edad: 'Edad',
  foto: 'Foto perfil',
  localidad: 'Localidad',
  provincia: 'Provincia / Departamento',
  pais: 'País',
  pasaporte: 'Pasaporte',
  nacionalidad2: 'Posee 2da nacionalidad? Mencione',
  contacto: 'Número de contacto',
  formacion: 'Nivel de formación',
  idiomas: 'Idiomas - Complementario al nativo',
  altura: 'Altura',
  peso: 'Peso',
  posicion: 'Posición de juego',
  perfil: 'Perfil',
  club: '¿Actualmente estás en actividad? Mencionar el Club',
  contrato: 'Bajo contrato',
  detalleContrato: 'En casi de haber respondido "Si" tiempo estimado, y cláusulas de contratación y baja del mismo.',
  video1: 'Video 1',
  video2: 'Video 2',
  video3: 'Video 3 - movimiento en el campo sin pelota',
  video4: 'Video 49',
  video5: 'Video 5',
  terminos: 'Acepto términos'
};

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Athlon Base Scouting')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function limpiarTexto(valor) {
  return String(valor || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[“”]/g, '"')
    .replace(/\s+/g, ' ')
    .trim();
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'No se encontró la hoja "' + CONFIG.SHEET_NAME + '". Hojas disponibles: ' +
      ss.getSheets().map(s => s.getName()).join(', ')
    );
  }

  return sheet;
}

function getPlayers() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getDisplayValues();

  if (!data || data.length <= 1) {
    return [];
  }

  const headers = data[0].map(limpiarTexto);

  const players = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (row.every(cell => limpiarTexto(cell) === '')) {
      continue;
    }

    const player = {
      rowNumber: i + 1
    };

    Object.keys(COLUMNAS).forEach(key => {
      const columnaBuscada = limpiarTexto(COLUMNAS[key]);
      const index = headers.indexOf(columnaBuscada);

      player[key] = index >= 0 ? row[index] : '';
    });

    player.nombreCompleto = `${player.nombre || ''} ${player.apellido || ''}`.trim();

    players.push(player);
  }

  return players;
}

function saveEvaluation(payload) {
  const sheet = getSheet_();
  const row = Number(payload.rowNumber);

  if (!row || row < 2) {
    throw new Error('Fila inválida para guardar evaluación.');
  }

  const nuevasColumnas = [
    'Evaluación Técnica',
    'Evaluación Táctica',
    'Evaluación Física',
    'Evaluación Mental',
    'Evaluación Proyección',
    'Observaciones Scouting',
    'Condición'
  ];

  asegurarColumnas_(sheet, nuevasColumnas);

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(limpiarTexto);

  escribirPorColumna_(sheet, row, headers, 'Evaluación Técnica', payload.tecnica);
  escribirPorColumna_(sheet, row, headers, 'Evaluación Táctica', payload.tactica);
  escribirPorColumna_(sheet, row, headers, 'Evaluación Física', payload.fisico);
  escribirPorColumna_(sheet, row, headers, 'Evaluación Mental', payload.mental);
  escribirPorColumna_(sheet, row, headers, 'Evaluación Proyección', payload.proyeccion);
  escribirPorColumna_(sheet, row, headers, 'Observaciones Scouting', payload.observaciones);
  escribirPorColumna_(sheet, row, headers, 'Condición', payload.condicion);

  return {
    ok: true,
    message: 'Evaluación guardada correctamente.'
  };
}

function asegurarColumnas_(sheet, columnas) {
  let headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(limpiarTexto);

  columnas.forEach(columna => {
    if (!headers.includes(columna)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(columna);
      headers.push(columna);
    }
  });
}

function escribirPorColumna_(sheet, row, headers, columna, valor) {
  const index = headers.indexOf(limpiarTexto(columna));

  if (index >= 0) {
    sheet.getRange(row, index + 1).setValue(valor || '');
  }
}

function diagnostico() {
  const sheet = getSheet_();

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0].map(limpiarTexto);

  Logger.log('Hoja usada: ' + sheet.getName());
  Logger.log('Filas: ' + sheet.getLastRow());
  Logger.log('Columnas: ' + sheet.getLastColumn());
  Logger.log('Encabezados encontrados:');
  Logger.log(headers);
  Logger.log('Jugadores detectados: ' + getPlayers().length);
}
