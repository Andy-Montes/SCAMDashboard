// ==========================================
// 0. SEGURIDAD
// ==========================================

const API_SECRET = 'Scm-xK9p-rT4q-2026';

const USER_PASSWORDS = {
  'Andrea Montes':      'am2026',
  'Camila Acuña':       'ca2026',
  'Victor Catalán':     'vc2026',
  'Cael Tarifa':        'ct2026',
  'Ricardo Ulloa':      'ru2026',
  'Katherine Martinez': 'km2026'
};

// ==========================================
// 1. CONFIGURACIÓN Y MAPEOS
// ==========================================

const MAP_TIPOS = {
  // Tipos actuales
  'Camara mal posicionada/Traslado':  'CT',
  'Camara mal posicionada/Reenfoque': 'CR',
  'Error de configuración':           'EC',
  'Obstrucción de cámara':            'OB',
  'Problema de imagen':               'PI',
  'Sin acceso a cámara':              'SA',
  'Problema de Red':                  'PR',
  // Legacy (resueltas históricas)
  'Cámara desalineada / mal posicionada': 'CD',
  'Sin acceso a Camara':              'SA',
  'Faltan datos de configuración':    'FD'
};

const MAP_EMAILS = {
  'Andrea Montes':      'amontes@sumatoid.com',
  'Victor Catalán':     'vcatalan@sumatoid.com',
  'Cael Tarifa':        'ctarifa@sumatoid.com',
  'Camila Acuña':       'cacuna@sumatoid.com',
  'Ricardo Ulloa':      'ext_rjulloah@sodimac.cl',
  'Katherine Martinez': 'kathmartinezg@sodimac.cl'
};

const MAP_TIENDAS = {
  'Cerrillos':        'CE',
  'Chillan':          'CH',
  'Coyhaique':        'CO',
  'El Belloto':       'EB',
  'Estación Central': 'EC',
  'Huechuraba':       'HU',
  'Las Condes':       'LC',
  'Maipú':            'MP',
  'Nueva La Florida': 'LF',
  'Open Kennedy':     'OK',
  'Puerto Montt':     'PM',
  'Punta Arenas':     'PA',
  'Quilicura':        'QU',
  'Ñuble':            'NU',
  'Ñuñoa':            'NO',
  'SUMATO TEST':      'XX'
};

const SHEETS = {
  INCIDENTES:         'Anomalías Deploy',
  CAMARAS:            'Total Cámaras',
  MANAGER:            'CamarasManager',
  LOG_SESIONES:       'Log_Sesiones',
  LOG_MODIFICACIONES: 'Log_Modificaciones',
  USUARIOS:           'Usuarios'
};

// ==========================================
// 2. FUNCIÓN ONEDIT (ingreso manual en Sheets)
// ==========================================

function onEdit(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  if (sheet.getName() !== SHEETS.INCIDENTES || row === 1) return;

  if (col === 1) {
    const deviceId = range.getValue();
    if (!deviceId) return;
    completarDatosCamara(ss, sheet, row, deviceId);
  }

  if (col === 7) {
    generarIdIncidencia(sheet, row);
  }

  // Auto-completar email cuando se selecciona responsable asignado (col S = 19)
  if (col === 19) {
    const responsable = normalizeText(range.getValue());
    const email = MAP_EMAILS[responsable] || '';
    sheet.getRange(row, 20).setValue(email); // T Email_Responsable
  }

  if (col === 13 && String(range.getValue()) === 'Si') {
    sheet.getRange(row, 14).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    );
  }
}

// ==========================================
// 3. HELPERS
// ==========================================

function normalizeText(value) {
  return String(value || '').trim();
}

function safeJsonParse(text) {
  try {
    return JSON.parse(text || '{}');
  } catch (err) {
    return null;
  }
}

function getLastDataRow_(sheet, colIndex) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(1, colIndex, maxRows, 1).getDisplayValues();
  for (let r = values.length - 1; r >= 0; r--) {
    const v = String(values[r][0] || '').trim();
    if (v !== '') return r + 1;
  }
  return 1;
}

function ensureSheetSize_(sheet, requiredRow, requiredCol) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (requiredRow > maxRows) sheet.insertRowsAfter(maxRows, requiredRow - maxRows);
  if (requiredCol > maxCols) sheet.insertColumnsAfter(maxCols, requiredCol - maxCols);
}

function copiarFormatoYFormulasFilaAnterior(sheet, newRow) {
  const prevRow = getLastDataRow_(sheet, 1);
  if (prevRow < 2) return;
  ensureSheetSize_(sheet, newRow, 21);
  const lastCol = Math.max(sheet.getLastColumn(), 21);
  const sourceRange = sheet.getRange(prevRow, 1, 1, lastCol);
  const targetRange = sheet.getRange(newRow, 1, 1, lastCol);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  const formulas = sourceRange.getFormulas()[0];
  for (let c = 1; c <= lastCol; c++) {
    if (c === 15) continue; // O = ID, no copiar fórmula
    const f = formulas[c - 1];
    if (f) sheet.getRange(newRow, c).setFormula(f);
  }
}

// ==========================================
// 4. FUNCIONES DE NEGOCIO
// ==========================================

function completarDatosCamara(ss, sheet, row, deviceId) {
  const stcSheet = ss.getSheetByName(SHEETS.CAMARAS);
  const scmSheet = ss.getSheetByName(SHEETS.MANAGER);
  if (!stcSheet || !scmSheet) return;

  const stc = stcSheet.getDataRange().getValues();
  const scm = scmSheet.getDataRange().getValues();

  const camData    = stc.find(r => String(r[0]) === String(deviceId));
  const managerData = scm.find(r => String(r[5]) === String(deviceId));

  if (camData) {
    sheet.getRange(row, 2).setValue(camData[3] || 'SIN INFO');  // B Tienda
    sheet.getRange(row, 4).setValue(camData[4] || 'SIN INFO');  // D Camara
    sheet.getRange(row, 5).setValue(camData[5] || 'SIN INFO');  // E PhysicalZoneName
    sheet.getRange(row, 6).setValue(camData[8] || 'SIN INFO');  // F IP
    sheet.getRange(row, 10).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    ); // J Fecha
    sheet.getRange(row, 13).setValue('No'); // M Resuelto
  }

  if (managerData) {
    sheet.getRange(row, 3).setValue(managerData[2] || 'SIN INFO'); // C Codigo
  }
}

function generarIdIncidencia(sheet, row) {
  const tipo   = normalizeText(sheet.getRange(row, 7).getValue());  // G
  const tienda = normalizeText(sheet.getRange(row, 2).getValue());  // B

  if (!tipo || !tienda) return '';

  const tipoCod   = MAP_TIPOS[tipo]   || 'OT';
  const tiendaCod = MAP_TIENDAS[tienda] || 'XX';
  const prefijo   = `${tipoCod}-${tiendaCod}-`;

  const lastRow = getLastDataRow_(sheet, 15);
  if (lastRow < 2) {
    const id = `${prefijo}001`;
    sheet.getRange(row, 15).setValue(id);
    return id;
  }

  const ids = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
  let maxNum = 0;

  for (let i = 0; i < ids.length; i++) {
    const idString = String(ids[i][0] || '');
    if (!idString.startsWith(prefijo)) continue;
    const partes = idString.split('-');
    if (partes.length !== 3) continue;
    const num = parseInt(partes[2], 10);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  }

  const nuevoId = `${prefijo}${String(maxNum + 1).padStart(3, '0')}`;
  sheet.getRange(row, 15).setValue(nuevoId);
  return nuevoId;
}

// ==========================================
// 5. LOGS
// ==========================================

function registrarSesion(usuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(SHEETS.LOG_SESIONES);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEETS.LOG_SESIONES);
    logSheet.appendRow(['Timestamp', 'Usuario', 'Acción']);
  }
  logSheet.appendRow([
    Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm:ss'),
    usuario || '',
    'Login'
  ]);
}

function registrarModificacion(usuario, accion, idIncidencia, detalle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(SHEETS.LOG_MODIFICACIONES);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEETS.LOG_MODIFICACIONES);
    logSheet.appendRow(['Timestamp', 'Usuario', 'Acción', 'ID Incidencia', 'Detalle']);
  }
  logSheet.appendRow([
    Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm:ss'),
    usuario     || '',
    accion      || '',
    idIncidencia || '',
    detalle     || ''
  ]);
}

// ==========================================
// 6. API PARA EL DASHBOARD
// ==========================================

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const incSheet  = ss.getSheetByName(SHEETS.INCIDENTES);
    const dataInc   = incSheet.getDataRange().getValues();
    const headersInc = dataInc[0] || [];

    const rowsInc = dataInc.slice(1).map(r => {
      const obj = {};
      headersInc.forEach((h, i) => (obj[h] = r[i]));
      obj.Cliente             = r[16] || ''; // Q
      obj.Registrado_Por      = r[17] || ''; // R
      obj.Responsable_Asignado = r[18] || ''; // S
      obj.Email_Responsable   = r[19] || ''; // T
      obj.Comentarios         = r[20] || ''; // U
      return obj;
    });

    const devSheet  = ss.getSheetByName(SHEETS.CAMARAS);
    const dataDev   = devSheet.getDataRange().getValues();
    const headersDev = dataDev[0] || [];

    const rowsDev = dataDev.slice(1).map(r => {
      const obj = {};
      headersDev.forEach((h, i) => (obj[h] = r[i]));
      return obj;
    });

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        incidencias: rowsInc,
        devices: rowsDev,
        lastUpdated: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = safeJsonParse(rawBody);

    if (!data) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', error: 'JSON inválido en body' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (normalizeText(data.apiKey) !== API_SECRET) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', error: 'No autorizado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const action  = normalizeText(data.action);
    const usuario = normalizeText(data.registradoPor || data.Registrado_Por || data.usuario || '');

    Logger.log('[doPost] action=%s usuario=%s', action, usuario);

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.INCIDENTES);

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', error: 'No existe la hoja Anomalías Deploy' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ------------------------------------------
    // LOGIN (registro de sesión)
    // ------------------------------------------
    if (action === 'login') {
      const usr = normalizeText(data.usuario || '');
      const pwd = normalizeText(data.password || '');
      if (!usr || !pwd) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'success', valid: false }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // --- Verificar bloqueo server-side ---
      const props = PropertiesService.getScriptProperties();
      const lockKey = 'lockout_' + usr;
      const lockRaw = props.getProperty(lockKey);
      if (lockRaw) {
        const lock = JSON.parse(lockRaw);
        const elapsed = (Date.now() - lock.since) / 60000; // minutos
        if (elapsed < 15) {
          const minutesLeft = Math.ceil(15 - elapsed);
          return ContentService
            .createTextOutput(JSON.stringify({ status: 'success', valid: false, locked: true, minutesLeft: minutesLeft }))
            .setMimeType(ContentService.MimeType.JSON);
        } else {
          props.deleteProperty(lockKey);
        }
      }

      // Buscar usuario en hoja Usuarios
      let usuariosSheet = ss.getSheetByName(SHEETS.USUARIOS);
      if (!usuariosSheet) {
        usuariosSheet = ss.insertSheet(SHEETS.USUARIOS);
        usuariosSheet.appendRow(['Usuario', 'Password', 'EsInicial']);
      }

      const usrData = usuariosSheet.getDataRange().getValues();
      let userRow = -1;
      let storedPwd = '';
      let esInicial = true;

      for (let i = 1; i < usrData.length; i++) {
        if (normalizeText(String(usrData[i][0])) === usr) {
          userRow = i + 1;
          storedPwd = normalizeText(String(usrData[i][1]));
          esInicial = usrData[i][2] === true || String(usrData[i][2]).toLowerCase() === 'true';
          break;
        }
      }

      const attKey = 'attempts_' + usr;

      if (userRow === -1) {
        // Primera vez: validar contra contraseñas iniciales
        const initialPwd = USER_PASSWORDS[usr] || '';
        if (!initialPwd || pwd !== initialPwd) {
          const attempts = parseInt(props.getProperty(attKey) || '0') + 1;
          if (attempts >= 3) {
            props.setProperty(lockKey, JSON.stringify({ since: Date.now() }));
            props.deleteProperty(attKey);
            return ContentService
              .createTextOutput(JSON.stringify({ status: 'success', valid: false, locked: true, minutesLeft: 15 }))
              .setMimeType(ContentService.MimeType.JSON);
          }
          props.setProperty(attKey, String(attempts));
          return ContentService
            .createTextOutput(JSON.stringify({ status: 'success', valid: false }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        // Crear entrada en hoja Usuarios
        props.deleteProperty(attKey);
        usuariosSheet.appendRow([usr, pwd, true]);
        registrarSesion(usr);
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'success', valid: true, mustChangePassword: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // Usuario ya registrado
      if (pwd !== storedPwd) {
        const attempts = parseInt(props.getProperty(attKey) || '0') + 1;
        if (attempts >= 3) {
          props.setProperty(lockKey, JSON.stringify({ since: Date.now() }));
          props.deleteProperty(attKey);
          return ContentService
            .createTextOutput(JSON.stringify({ status: 'success', valid: false, locked: true, minutesLeft: 15 }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        props.setProperty(attKey, String(attempts));
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'success', valid: false }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      props.deleteProperty(attKey);
      registrarSesion(usr);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', valid: true, mustChangePassword: esInicial }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ------------------------------------------
    // CAMBIAR CONTRASEÑA
    // ------------------------------------------
    if (action === 'cambiar_password') {
      const usr        = normalizeText(data.usuario || '');
      const pwdActual  = normalizeText(data.passwordActual || '');
      const pwdNueva   = normalizeText(data.passwordNueva || '');

      if (!usr || !pwdActual || !pwdNueva) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', error: 'Faltan datos' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const usuariosSheet = ss.getSheetByName(SHEETS.USUARIOS);
      if (!usuariosSheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', error: 'Hoja Usuarios no encontrada' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const usrData = usuariosSheet.getDataRange().getValues();
      for (let i = 1; i < usrData.length; i++) {
        if (normalizeText(String(usrData[i][0])) === usr) {
          if (normalizeText(String(usrData[i][1])) !== pwdActual) {
            return ContentService
              .createTextOutput(JSON.stringify({ status: 'error', error: 'Contraseña actual incorrecta' }))
              .setMimeType(ContentService.MimeType.JSON);
          }
          usuariosSheet.getRange(i + 1, 2).setValue(pwdNueva);  // nueva contraseña
          usuariosSheet.getRange(i + 1, 3).setValue(false);     // ya no es inicial
          registrarModificacion(usr, 'CambiarPassword', '', '');
          return ContentService
            .createTextOutput(JSON.stringify({ status: 'success', message: 'Contraseña actualizada' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', error: 'Usuario no encontrado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ------------------------------------------
    // AGREGAR INCIDENCIA
    // ------------------------------------------
    if (action === 'agregar_incidencia') {
      const deviceId          = normalizeText(data.deviceId || data.DeviceId || '');
      const esManual          = data.manual === true;
      const tipo              = normalizeText(data.tipoAnomalia || data.tipo || data.Tipo || '');
      const descripcion       = normalizeText(data.descripcion || data.Descripcion || '');
      const estadoCamara      = normalizeText(data.estadoCamara || data.estado || data.Estado || '');
      const responsableTipo   = normalizeText(data.responsableTipo || data.ResponsableTipo || '');
      const cliente           = normalizeText(data.cliente || data.Cliente || '');
      const registradoPor     = normalizeText(data.registradoPor || data.Registrado_Por || '');
      const responsableAsignado = normalizeText(data.responsable || data.Responsable_Asignado || '');
      const emailResponsable  = MAP_EMAILS[responsableAsignado] || normalizeText(data.emailResponsable || '');

      // Campos exclusivos del modo manual
      const tiendaManual  = normalizeText(data.tienda || '');
      const camaraManual  = normalizeText(data.camara || '');
      const ipManual      = normalizeText(data.ip || '');

      if (!esManual && !deviceId) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', error: 'Falta deviceId' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      if (esManual && !tiendaManual) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', error: 'Falta tienda en modo manual' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const lastDataRow = getLastDataRow_(sheet, 1);
      const newRow      = lastDataRow + 1;

      ensureSheetSize_(sheet, newRow, 21);
      copiarFormatoYFormulasFilaAnterior(sheet, newRow);

      if (esManual) {
        sheet.getRange(newRow, 1).setValue('');              // A DeviceId vacío
        sheet.getRange(newRow, 2).setValue(tiendaManual);    // B Tienda
        sheet.getRange(newRow, 3).setValue('');              // C Codigo
        sheet.getRange(newRow, 4).setValue(camaraManual);    // D Camara
        sheet.getRange(newRow, 5).setValue('');              // E PhysicalZoneName
        sheet.getRange(newRow, 6).setValue(ipManual);        // F IP
        sheet.getRange(newRow, 10).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
        );                                                   // J Fecha
        sheet.getRange(newRow, 13).setValue('No');            // M Resuelto
      } else {
        sheet.getRange(newRow, 1).setValue(deviceId);
        completarDatosCamara(ss, sheet, newRow, deviceId);  // B,C,D,E,F,J,M
      }

      sheet.getRange(newRow, 7).setValue(tipo);                  // G Tipo
      sheet.getRange(newRow, 8).setValue(descripcion);           // H Descripcion
      sheet.getRange(newRow, 9).setValue(estadoCamara);          // I Estado
      sheet.getRange(newRow, 11).setValue(responsableTipo);      // K Responsable (Sodimac/Sumato)
      sheet.getRange(newRow, 17).setValue(cliente);              // Q Cliente
      sheet.getRange(newRow, 18).setValue(registradoPor);        // R Registrado_Por
      sheet.getRange(newRow, 19).setValue(responsableAsignado);  // S Responsable Asignado
      sheet.getRange(newRow, 20).setValue(emailResponsable);     // T Email_Responsable
      sheet.getRange(newRow, 21).setValue('');                    // U Comentarios (vacío al crear)

      const idGenerado = generarIdIncidencia(sheet, newRow) || String(sheet.getRange(newRow, 15).getValue() || '');

      registrarModificacion(registradoPor, esManual ? 'Crear-Manual' : 'Crear', idGenerado,
        `Device:${deviceId || 'MANUAL'} Tienda:${tiendaManual || ''} Tipo:${tipo} Responsable:${responsableAsignado}`);

      Logger.log('[agregar_incidencia] OK row=%s id=%s manual=%s', newRow, idGenerado, esManual);

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', message: 'Incidencia agregada correctamente', id: idGenerado
     }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ------------------------------------------
    // MARCAR RESUELTA
    // ------------------------------------------
    if (action === 'marcar_resuelta') {
      const idIncidencia = normalizeText(data.idIncidencia || data.ID || '');
      if (!idIncidencia) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', error: 'Falta idIncidencia' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const lastRow = getLastDataRow_(sheet, 15);
      const allData = sheet.getRange(1, 1, lastRow, 21).getValues();

      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][14]) === idIncidencia) {
          sheet.getRange(i + 1, 13).setValue('Si'); // M Resuelto
          sheet.getRange(i + 1, 14).setValue(
            Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
          ); // N FechaResolucion

          registrarModificacion(usuario, 'Resolver', idIncidencia, '');

          return ContentService
            .createTextOutput(JSON.stringify({ status: 'success', message: 'Incidencia marcada como resuelta' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', error: 'ID de incidencia no encontrado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ------------------------------------------
    // EDITAR INCIDENCIA
    // ------------------------------------------
    if (action === 'editar_incidencia') {
      const idIncidencia        = normalizeText(data.idIncidencia || data.ID || '');
      const tipo                = normalizeText(data.tipoAnomalia || data.tipo || data.Tipo || '');
      const descripcion         = normalizeText(data.descripcion || data.Descripcion || '');
      const estadoCamara        = normalizeText(data.estadoCamara || data.estado || data.Estado || '');
      const registradoPor       = normalizeText(data.registradoPor || data.Registrado_Por || '');
      const responsableTipo     = normalizeText(data.responsableTipo || data.ResponsableTipo || '');
      const responsableAsignado = normalizeText(data.responsable || data.Responsable_Asignado || '');
      const emailResponsable    = MAP_EMAILS[responsableAsignado] || normalizeText(data.emailResponsable || '');
      const comentarios         = normalizeText(data.comentarios || data.Comentarios || '');

      if (!idIncidencia) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', error: 'Falta idIncidencia' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const lastRow = getLastDataRow_(sheet, 15);
      const allData = sheet.getRange(1, 1, lastRow, 21).getValues();

      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][14]) === idIncidencia) {
          sheet.getRange(i + 1, 7).setValue(tipo);                  // G Tipo
          sheet.getRange(i + 1, 8).setValue(descripcion);           // H Descripcion
          sheet.getRange(i + 1, 9).setValue(estadoCamara);          // I Estado
          sheet.getRange(i + 1, 11).setValue(responsableTipo);      // K Responsable (Sodimac/Sumato)
          sheet.getRange(i + 1, 18).setValue(registradoPor);        // R Registrado_Por
          sheet.getRange(i + 1, 19).setValue(responsableAsignado);  // S Responsable Asignado
          sheet.getRange(i + 1, 20).setValue(emailResponsable);     // T Email_Responsable
          sheet.getRange(i + 1, 21).setValue(comentarios);           // U Comentarios

          registrarModificacion(usuario, 'Editar', idIncidencia,
            `Tipo:${tipo} Responsable:${responsableAsignado}`);

          return ContentService
            .createTextOutput(JSON.stringify({ status: 'success', message: 'Incidencia editada' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', error: 'ID de incidencia no encontrado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ------------------------------------------
    // ELIMINAR INCIDENCIA
    // ------------------------------------------
    if (action === 'eliminar_incidencia') {
      const idIncidencia = normalizeText(data.idIncidencia || data.ID || '');
      if (!idIncidencia) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', error: 'Falta idIncidencia' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const lastRow = getLastDataRow_(sheet, 15);
      const allData = sheet.getRange(1, 1, lastRow, 21).getValues();

      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][14]) === idIncidencia) {
          registrarModificacion(usuario, 'Eliminar', idIncidencia,
            `Device:${allData[i][0]} Tipo:${allData[i][6]}`);
          sheet.deleteRow(i + 1);
          return ContentService
            .createTextOutput(JSON.stringify({ status: 'success', message: 'Incidencia eliminada' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', error: 'ID de incidencia no encontrado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', error: 'Acción no válida' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('[doPost] ERROR: %s', String(err && err.stack ? err.stack : err));
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 7. CORRECCIÓN MASIVA DE IDs
// ==========================================

function corregirTodosLosIdsMasivo() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.INCIDENTES);
  const lastRow = getLastDataRow_(sheet, 1);
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const contadores = {};

  for (let i = 0; i < data.length; i++) {
    const tiendaRaw = String(data[i][1] || '').trim();
    const tipoRaw   = String(data[i][6] || '').trim();
    if (!tiendaRaw || !tipoRaw) continue;

    const tipoCod   = MAP_TIPOS[tipoRaw]   || 'OT';
    const tiendaCod = MAP_TIENDAS[tiendaRaw] || 'XX';
    const prefijo   = `${tipoCod}-${tiendaCod}-`;

    contadores[prefijo] = (contadores[prefijo] || 0) + 1;
    const nuevoId = `${prefijo}${String(contadores[prefijo]).padStart(3, '0')}`;
    sheet.getRange(i + 2, 15).setValue(nuevoId);
  }

  Logger.log('✅ IDs corregidos exitosamente.');
}
