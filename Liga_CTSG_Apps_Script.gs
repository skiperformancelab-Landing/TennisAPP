// ============================================================
// LIGA CTSG - Google Apps Script Backend
// Versión 1.0 | Junio 2026
// Pega este código en Google Sheets → Extensiones → Apps Script
// ============================================================

const SPREADSHEET_ID = '18wMAfIKZ7tQ85BJMkvgjG4wONeF8FfxUdX17XjcXNwc';
const ADMIN_KEY = 'ctsg2026admin'; // Cambia esto por seguridad

// ------------------------------------------------------------
// PUNTO DE ENTRADA - Maneja todas las peticiones de la app
// ------------------------------------------------------------
function doGet(e) {
  const p = e.parameter;
  let result;
  try {
    switch(p.action) {
      case 'setup':       result = setupSheets(); break;
      case 'login':       result = login(p.username, p.password); break;
      case 'getData':     result = getData(p.userId); break;
      case 'saveResult':  result = saveResult(p); break;
      case 'cancelMatch':      result = cancelMatch(p); break;
      case 'changePassword':   result = changePassword(p); break;
      case 'getConfig':   result = getConfig(); break;
      case 'advanceFase':
        if(p.adminKey !== ADMIN_KEY) { result = {ok:false, error:'No autorizado'}; break; }
        result = advanceFase();
        break;
      case 'simularResultados':
        if(p.adminKey !== ADMIN_KEY) { result = {ok:false, error:'No autorizado'}; break; }
        result = simularResultados();
        break;
      case 'resetFase':
        if(p.adminKey !== ADMIN_KEY) { result = {ok:false, error:'No autorizado'}; break; }
        result = resetFase(parseInt(p.fase||1));
        break;
      default: result = {ok:false, error:'Acción desconocida: ' + p.action};
    }
  } catch(err) {
    result = {ok:false, error: err.message};
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// SETUP - Ejecutar UNA VEZ para inicializar la hoja
// Ir a Apps Script → Ejecutar → setupSheets
// ------------------------------------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Crear hojas necesarias
  ['Jugadores','Partidos','Config'].forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) ss.insertSheet(name);
    else sh.clearContents();
  });

  // ---- Config ----
  const cfg = ss.getSheetByName('Config');
  cfg.getRange('A1:B1').setValues([['clave','valor']]);
  cfg.getRange('A2:B5').setValues([
    ['fase_activa','1'],
    ['deadline_f1','2026-07-31'],
    ['deadline_f2','2026-08-31'],
    ['ultima_actualizacion', new Date().toISOString()]
  ]);

  // ---- Jugadores ----
  const jugSh = ss.getSheetByName('Jugadores');
  jugSh.getRange('A1:H1').setValues([['id','nombre','telefono','nivel','username','password','grupo_f1','activo']]);
  const jugadores = getJugadoresData();
  const jugRows = jugadores.map((j, i) => [
    i,
    j.nombre,
    j.telefono,
    j.nivel,
    j.username,
    j.password,
    String.fromCharCode(65 + Math.floor(i / 4)), // grupo_f1: A,A,A,A,B,B,B,B...
    1
  ]);
  jugSh.getRange(2, 1, jugRows.length, jugRows[0].length).setValues(jugRows);

  // ---- Partidos ----
  const partSh = ss.getSheetByName('Partidos');
  partSh.getRange('A1:N1').setValues([
    ['id','fase','grupo','p1_id','p2_id','s1p1','s1p2','s2p1','s2p2','s3p1','s3p2','estado','fecha','anotado_por']
  ]);
  const partidos = [];
  let mid = 0;
  for (let g = 0; g < 25; g++) {
    const grupo = String.fromCharCode(65 + g);
    for (let i = 0; i < 4; i++) {
      for (let j = i + 1; j < 4; j++) {
        const p1 = g * 4 + i;
        const p2 = g * 4 + j;
        if (p1 < jugadores.length && p2 < jugadores.length) {
          partidos.push([mid++, 1, grupo, p1, p2, '','','','','','','pending','','']);
        }
      }
    }
  }
  partSh.getRange(2, 1, partidos.length, partidos[0].length).setValues(partidos);

  return {ok:true, msg:`✅ Setup completo: ${jugadores.length} jugadores, ${partidos.length} partidos de Fase 1 generados`};
}

// ------------------------------------------------------------
// LOGIN
// ------------------------------------------------------------
function login(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = ss.getSheetByName('Jugadores').getDataRange().getValues();
  const h = data[0];
  const iUser = h.indexOf('username');
  const iPass = h.indexOf('password');
  const iId   = h.indexOf('id');
  const iNom  = h.indexOf('nombre');
  const iNiv  = h.indexOf('nivel');
  const iGf1  = h.indexOf('grupo_f1');
  const iAct  = h.indexOf('activo');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (
      row[iUser].toString().toLowerCase() === username.toLowerCase() &&
      row[iPass].toString() === password &&
      row[iAct] == 1
    ) {
      // Obtener grupo de fase activa
      const cfg = getConfigData(ss);
      const fase = parseInt(cfg.fase_activa || 1);
      const grupoCol = fase === 1 ? iGf1 : h.indexOf('grupo_f' + fase);
      return {
        ok: true,
        userId: row[iId],
        nombre: row[iNom],
        nivel: row[iNiv],
        grupo: grupoCol >= 0 ? row[grupoCol] : row[iGf1],
        faseActiva: fase
      };
    }
  }
  return {ok:false, error:'Usuario o contraseña incorrectos'};
}

// ------------------------------------------------------------
// CHANGE PASSWORD
// ------------------------------------------------------------
function changePassword(p) {
  const userId    = parseInt(p.userId);
  const oldPass   = (p.oldPassword || '').toString();
  const newPass   = (p.newPassword || '').toString();
  if (!oldPass || !newPass) return {ok:false, error:'Parámetros incompletos'};
  if (newPass.length < 6)   return {ok:false, error:'La contraseña debe tener al menos 6 caracteres'};

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh   = ss.getSheetByName('Jugadores');
  const data = sh.getDataRange().getValues();
  const h    = data[0];
  const iId  = h.indexOf('id');
  const iPass = h.indexOf('password');

  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][iId]) === userId) {
      if (data[i][iPass].toString() !== oldPass) {
        return {ok:false, error:'La contraseña actual no es correcta'};
      }
      sh.getRange(i + 1, iPass + 1).setValue(newPass);
      return {ok:true};
    }
  }
  return {ok:false, error:'Usuario no encontrado'};
}

// ------------------------------------------------------------
// GET DATA - Carga toda la información del usuario
// ------------------------------------------------------------
function getData(userId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg = getConfigData(ss);
  const fase = parseInt(cfg.fase_activa || 1);
  const deadline = cfg['deadline_f' + fase] || '';

  // Siempre calculamos ranking F1 (histórico)
  const rankingF1 = calcularRanking(ss, 1);

  let ranking;
  if (fase === 1) {
    ranking = rankingF1;
  } else {
    // Para fase 2+: pts acumulados = F1_pts + F2_pts
    const f1Map = {};
    rankingF1.forEach(function(r) { f1Map[r.id] = r.pts || 0; });
    const rankingFaseActual = calcularRanking(ss, fase);
    ranking = rankingFaseActual.map(function(r) {
      const ptsF1 = f1Map[r.id] || 0;
      return {
        id: r.id, nombre: r.nombre, telefono: r.telefono, nivel: r.nivel,
        grupo: r.grupo, pj: r.pj, v: r.v, d: r.d, ds: r.ds, dj: r.dj,
        pts_f1: ptsF1,
        pts: ptsF1 + (r.pts || 0)   // total acumulado
      };
    });
  }

  // Todos los partidos del jugador (todas las fases)
  const partidos = getMisPartidos(ss, parseInt(userId));
  const jugInfo  = getJugadorInfo(ss, parseInt(userId), fase);

  return {
    ok: true, faseActiva: fase, deadline,
    ranking,
    rankingF1: fase > 1 ? rankingF1 : null,
    misPartidos: partidos,
    jugador: jugInfo
  };
}

// ------------------------------------------------------------
// SAVE RESULT - Guarda el marcador de un partido
// ------------------------------------------------------------
function saveResult(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg = getConfigData(ss);
  const fase = parseInt(cfg.fase_activa || 1);

  // Verificar plazo
  const deadlineStr = cfg['deadline_f' + fase];
  if (deadlineStr) {
    const deadline = new Date(deadlineStr);
    deadline.setHours(23, 59, 59);
    if (new Date() > deadline) return {ok:false, error:'⏰ Plazo cerrado para esta fase'};
  }

  const sh = ss.getSheetByName('Partidos');
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const p1id = parseInt(p.p1id);
  const p2id = parseInt(p.p2id);
  const fasePart = parseInt(p.fase || fase);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rFase   = parseInt(row[h.indexOf('fase')]);
    const rP1     = parseInt(row[h.indexOf('p1_id')]);
    const rP2     = parseInt(row[h.indexOf('p2_id')]);
    const rEstado = row[h.indexOf('estado')];

    if (rFase === fasePart && rEstado === 'pending' &&
        ((rP1 === p1id && rP2 === p2id) || (rP1 === p2id && rP2 === p1id))) {

      const swapped = (rP1 === p2id); // el orden en sheet está al revés
      const rowNum  = i + 1;

      sh.getRange(rowNum, h.indexOf('s1p1')+1).setValue(swapped ? p.s1p2 : p.s1p1);
      sh.getRange(rowNum, h.indexOf('s1p2')+1).setValue(swapped ? p.s1p1 : p.s1p2);
      sh.getRange(rowNum, h.indexOf('s2p1')+1).setValue(swapped ? p.s2p2 : p.s2p1);
      sh.getRange(rowNum, h.indexOf('s2p2')+1).setValue(swapped ? p.s2p1 : p.s2p2);
      if (p.s3p1 && p.s3p2) {
        sh.getRange(rowNum, h.indexOf('s3p1')+1).setValue(swapped ? p.s3p2 : p.s3p1);
        sh.getRange(rowNum, h.indexOf('s3p2')+1).setValue(swapped ? p.s3p1 : p.s3p2);
      }
      sh.getRange(rowNum, h.indexOf('estado')+1).setValue('played');
      sh.getRange(rowNum, h.indexOf('fecha')+1).setValue(new Date().toISOString());
      sh.getRange(rowNum, h.indexOf('anotado_por')+1).setValue(p.userId || '');

      return {ok:true, msg:'✅ Resultado guardado'};
    }
  }
  return {ok:false, error:'Partido no encontrado o ya anotado'};
}

// ------------------------------------------------------------
// ADVANCE FASE - Solo admin
// Llamar: URL?action=advanceFase&adminKey=ctsg2026admin
// ------------------------------------------------------------
function advanceFase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg = getConfigData(ss);
  const faseActual = parseInt(cfg.fase_activa || 1);
  const nuevaFase  = faseActual + 1;

  // Calcular ranking de fase actual y ordenar jugadores
  const ranking = calcularRanking(ss, faseActual);
  const sorted  = ranking.sort((a,b) => (b.pts-a.pts) || (b.ds-a.ds) || (b.dj-a.dj));

  // Actualizar grupo_f{n} en hoja Jugadores
  const jugSh   = ss.getSheetByName('Jugadores');
  const jugData = jugSh.getDataRange().getValues();
  const jugH    = jugData[0];
  const colName = 'grupo_f' + nuevaFase;
  let grupoColIdx = jugH.indexOf(colName);

  if (grupoColIdx === -1) {
    // Añadir nueva columna
    const newCol = jugSh.getLastColumn() + 1;
    jugSh.getRange(1, newCol).setValue(colName);
    grupoColIdx = newCol - 1;
  }

  sorted.forEach((player, idx) => {
    const newGrupo = String.fromCharCode(65 + Math.floor(idx / 4));
    for (let i = 1; i < jugData.length; i++) {
      if (parseInt(jugData[i][jugH.indexOf('id')]) === parseInt(player.id)) {
        jugSh.getRange(i + 1, grupoColIdx + 1).setValue(newGrupo);
        break;
      }
    }
  });

  // Generar partidos de la nueva fase
  const partSh    = ss.getSheetByName('Partidos');
  let   mid       = partSh.getLastRow(); // continuar IDs
  const newParts  = [];

  for (let g = 0; g < 25; g++) {
    const grupo   = String.fromCharCode(65 + g);
    const grPlayers = sorted.slice(g * 4, g * 4 + 4);
    for (let i = 0; i < grPlayers.length; i++) {
      for (let j = i + 1; j < grPlayers.length; j++) {
        newParts.push([mid++, nuevaFase, grupo, grPlayers[i].id, grPlayers[j].id, '','','','','','','pending','','']);
      }
    }
  }
  if (newParts.length > 0) {
    partSh.getRange(partSh.getLastRow()+1, 1, newParts.length, newParts[0].length).setValues(newParts);
  }

  // Actualizar deadline para nueva fase
  const deadlines = {
    2: '2026-08-31', 3: '2026-09-30', 4: '2026-10-31',
    5: '2026-11-30', 6: '2026-12-31', 7: '2027-01-31',
    8: '2027-02-28', 9: '2027-03-31', 10: '2027-04-30',
    11: '2027-05-31', 12: '2027-06-30'
  };
  if (deadlines[nuevaFase]) setConfigValue(ss, 'deadline_f'+nuevaFase, deadlines[nuevaFase]);

  // Actualizar fase_activa
  setConfigValue(ss, 'fase_activa', nuevaFase.toString());
  setConfigValue(ss, 'ultima_actualizacion', new Date().toISOString());

  return {
    ok: true,
    msg: `🎾 Fase ${nuevaFase} activada. ${newParts.length} nuevos partidos generados.`,
    nuevaFase,
    totalPartidos: newParts.length
  };
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------
function getConfigData(ss) {
  const data = ss.getSheetByName('Config').getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < data.length; i++) cfg[data[i][0]] = data[i][1];
  return cfg;
}

function getConfig() {
  return {ok:true, config: getConfigData(SpreadsheetApp.openById(SPREADSHEET_ID))};
}

function setConfigValue(ss, key, value) {
  const sh   = ss.getSheetByName('Config');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sh.getRange(i+1, 2).setValue(value); return; }
  }
  sh.getRange(sh.getLastRow()+1, 1, 1, 2).setValues([[key, value]]);
}

function getJugadorInfo(ss, userId, fase) {
  const data = ss.getSheetByName('Jugadores').getDataRange().getValues();
  const h = data[0];
  for (let i = 1; i < data.length; i++) {
    if (parseInt(data[i][h.indexOf('id')]) === userId) {
      const grupoCol = h.indexOf('grupo_f' + fase);
      return {
        id: data[i][h.indexOf('id')],
        nombre: data[i][h.indexOf('nombre')],
        telefono: data[i][h.indexOf('telefono')],
        nivel: data[i][h.indexOf('nivel')],
        grupo: grupoCol >= 0 ? data[i][grupoCol] : data[i][h.indexOf('grupo_f1')]
      };
    }
  }
  return null;
}

function calcularRanking(ss, fase) {
  const jugData  = ss.getSheetByName('Jugadores').getDataRange().getValues();
  const jugH     = jugData[0];
  const partData = ss.getSheetByName('Partidos').getDataRange().getValues();
  const partH    = partData[0];
  const grupoCol = jugH.indexOf('grupo_f' + fase) >= 0 ? jugH.indexOf('grupo_f' + fase) : jugH.indexOf('grupo_f1');

  // Inicializar stats
  const stats = {};
  for (let i = 1; i < jugData.length; i++) {
    const id = parseInt(jugData[i][jugH.indexOf('id')]);
    stats[id] = {
      id,
      nombre:    jugData[i][jugH.indexOf('nombre')],
      telefono:  jugData[i][jugH.indexOf('telefono')],
      nivel:     jugData[i][jugH.indexOf('nivel')],
      grupo:     jugData[i][grupoCol] || jugData[i][jugH.indexOf('grupo_f1')],
      pj:0, v:0, d:0, ds:0, dj:0, pts:0
    };
  }

  // Procesar partidos jugados y cancelados
  for (let i = 1; i < partData.length; i++) {
    const row    = partData[i];
    const estado = row[partH.indexOf('estado')];
    if (parseInt(row[partH.indexOf('fase')]) !== fase) continue;
    if (estado !== 'played' && estado !== 'cancelled') continue;

    const p1id = parseInt(row[partH.indexOf('p1_id')]);
    const p2id = parseInt(row[partH.indexOf('p2_id')]);
    if (!stats[p1id] || !stats[p2id]) continue;

    // Partido cancelado → penalización -2 DS y -2 DJ a cada jugador (= -4 pts)
    if (estado === 'cancelled') {
      stats[p1id].pj++; stats[p2id].pj++;
      stats[p1id].d++;  stats[p2id].d++;
      stats[p1id].ds -= 2; stats[p2id].ds -= 2;
      stats[p1id].dj -= 2; stats[p2id].dj -= 2;
      stats[p1id].pts = stats[p1id].ds + stats[p1id].dj;
      stats[p2id].pts = stats[p2id].ds + stats[p2id].dj;
      continue;
    }

    const s1p1 = Number(row[partH.indexOf('s1p1')]) || 0;
    const s1p2 = Number(row[partH.indexOf('s1p2')]) || 0;
    const s2p1 = Number(row[partH.indexOf('s2p1')]) || 0;
    const s2p2 = Number(row[partH.indexOf('s2p2')]) || 0;
    const s3p1 = Number(row[partH.indexOf('s3p1')]) || 0;
    const s3p2 = Number(row[partH.indexOf('s3p2')]) || 0;

    let sets1 = 0, sets2 = 0;
    if (s1p1 > s1p2) sets1++; else if (s1p2 > s1p1) sets2++;
    if (s2p1 > s2p2) sets1++; else if (s2p2 > s2p1) sets2++;
    if (row[partH.indexOf('s3p1')]) {
      if (s3p1 > s3p2) sets1++; else if (s3p2 > s3p1) sets2++;
    }

    const g1 = s1p1+s2p1+s3p1, g2 = s1p2+s2p2+s3p2;

    stats[p1id].pj++; stats[p2id].pj++;
    if (sets1 > sets2) { stats[p1id].v++; stats[p2id].d++; }
    else               { stats[p1id].d++; stats[p2id].v++; }
    stats[p1id].ds += sets1-sets2; stats[p2id].ds += sets2-sets1;
    stats[p1id].dj += g1-g2;       stats[p2id].dj += g2-g1;
    stats[p1id].pts = stats[p1id].ds + stats[p1id].dj;
    stats[p2id].pts = stats[p2id].ds + stats[p2id].dj;
  }

  return Object.values(stats);
}

function getMisPartidos(ss, userId) {
  // Devuelve TODOS los partidos del jugador en todas las fases
  const partData = ss.getSheetByName('Partidos').getDataRange().getValues();
  const partH    = partData[0];
  const jugData  = ss.getSheetByName('Jugadores').getDataRange().getValues();
  const jugH     = jugData[0];

  const nombresMap = {};
  const telMap     = {};
  for (let i = 1; i < jugData.length; i++) {
    const id = parseInt(jugData[i][jugH.indexOf('id')]);
    nombresMap[id] = jugData[i][jugH.indexOf('nombre')];
    telMap[id]     = jugData[i][jugH.indexOf('telefono')];
  }

  const result = [];
  for (let i = 1; i < partData.length; i++) {
    const row   = partData[i];
    const rFase = parseInt(row[partH.indexOf('fase')]);
    const rP1   = parseInt(row[partH.indexOf('p1_id')]);
    const rP2   = parseInt(row[partH.indexOf('p2_id')]);
    if (rP1 !== userId && rP2 !== userId) continue; // todas las fases

    result.push({
      id:       row[partH.indexOf('id')],
      fase:     rFase,
      grupo:    row[partH.indexOf('grupo')],
      p1_id:    rP1,
      p2_id:    rP2,
      p1_nombre: nombresMap[rP1] || '',
      p2_nombre: nombresMap[rP2] || '',
      p1_tel:   telMap[rP1] || '',
      p2_tel:   telMap[rP2] || '',
      s1p1: row[partH.indexOf('s1p1')], s1p2: row[partH.indexOf('s1p2')],
      s2p1: row[partH.indexOf('s2p1')], s2p2: row[partH.indexOf('s2p2')],
      s3p1: row[partH.indexOf('s3p1')], s3p2: row[partH.indexOf('s3p2')],
      estado: row[partH.indexOf('estado')],
      fecha:  row[partH.indexOf('fecha')]
    });
  }
  return result;
}

// ------------------------------------------------------------
// DATOS DE LOS 101 JUGADORES
// Contraseña inicial = últimas 4 cifras del teléfono
// ------------------------------------------------------------
function getJugadoresData() {
  return [
    {nombre:'Marc Pascual Segarra',       telefono:'+34 612 345 678', nivel:5, username:'marc.pascual',       password:'5678'},
    {nombre:'Guillermo Ortega Jiménez',   telefono:'+34 687 750 275', nivel:5, username:'guillermo.ortega',   password:'0275'},
    {nombre:'Jordi Vila Font',            telefono:'+34 730 652 897', nivel:5, username:'jordi.vila',         password:'2897'},
    {nombre:'Manuel Sánchez Muñoz',       telefono:'+34 672 935 170', nivel:5, username:'manuel.sanchez',     password:'5170'},
    {nombre:'Toni González Domínguez',    telefono:'+34 643 640 993', nivel:5, username:'toni.gonzalez',      password:'0993'},
    {nombre:'Cristian Mas Cano',          telefono:'+34 618 446 121', nivel:5, username:'cristian.mas',       password:'6121'},
    {nombre:'Miguel Vega Pérez',          telefono:'+34 795 597 319', nivel:5, username:'miguel.vega',        password:'7319'},
    {nombre:'Roger Muñoz Ramos',          telefono:'+34 684 708 140', nivel:5, username:'roger.munoz',        password:'8140'},
    {nombre:'Sergio Muñoz Molina',        telefono:'+34 779 820 409', nivel:5, username:'sergio.munoz',       password:'0409'},
    {nombre:'Luis Vega Romero',           telefono:'+34 792 621 509', nivel:5, username:'luis.vega',          password:'1509'},
    {nombre:'Xavier Font Santos',         telefono:'+34 766 791 318', nivel:5, username:'xavier.font',        password:'1318'},
    {nombre:'Nicolás Mendoza Ramírez',    telefono:'+34 775 779 748', nivel:5, username:'nicolas.mendoza',    password:'9748'},
    {nombre:'Matías Muñoz Fernández',     telefono:'+34 795 643 672', nivel:5, username:'matias.munoz',       password:'3672'},
    {nombre:'Alejandro Navarro Vega',     telefono:'+34 669 230 737', nivel:5, username:'alejandro.navarro',  password:'0737'},
    {nombre:'Toni Molina Ramírez',        telefono:'+34 630 418 210', nivel:5, username:'toni.molina',        password:'8210'},
    {nombre:'Pablo Blanco Molina',        telefono:'+34 665 280 851', nivel:5, username:'pablo.blanco',       password:'0851'},
    {nombre:'Marc Romero Camps',          telefono:'+34 765 945 848', nivel:5, username:'marc.romero',        password:'5848'},
    {nombre:'Fernando Castillo Muñoz',    telefono:'+34 741 868 575', nivel:5, username:'fernando.castillo',  password:'8575'},
    {nombre:'Cristian López Gómez',       telefono:'+34 798 349 413', nivel:5, username:'cristian.lopez',     password:'9413'},
    {nombre:'Fernando Díaz Álvarez',      telefono:'+34 650 222 860', nivel:5, username:'fernando.diaz',      password:'2860'},
    {nombre:'Xavi Vargas Flores',         telefono:'+34 745 841 703', nivel:5, username:'xavi.vargas',        password:'1703'},
    {nombre:'Rubén Sánchez Álvarez',      telefono:'+34 661 603 175', nivel:5, username:'ruben.sanchez',      password:'3175'},
    {nombre:'Jordi Cruz Sánchez',         telefono:'+34 720 354 221', nivel:5, username:'jordi.cruz',         password:'4221'},
    {nombre:'Oriol Ruiz Vargas',          telefono:'+34 643 906 910', nivel:5, username:'oriol.ruiz',         password:'6910'},
    {nombre:'Antonio Vila Font',          telefono:'+34 693 818 658', nivel:4, username:'antonio.vila',       password:'8658'},
    {nombre:'Roberto González Muñoz',     telefono:'+34 784 538 697', nivel:4, username:'roberto.gonzalez',   password:'8697'},
    {nombre:'Jordi Ramírez Cruz',         telefono:'+34 659 490 710', nivel:4, username:'jordi.ramirez',      password:'0710'},
    {nombre:'David Prat Mendoza',         telefono:'+34 717 151 698', nivel:4, username:'david.prat',         password:'1698'},
    {nombre:'Guillermo Martínez Ortega',  telefono:'+34 655 898 673', nivel:4, username:'guillermo.martinez', password:'8673'},
    {nombre:'Héctor López Fernández',     telefono:'+34 714 211 710', nivel:4, username:'hector.lopez',       password:'1710'},
    {nombre:'Arnau Castillo Domínguez',   telefono:'+34 795 518 434', nivel:4, username:'arnau.castillo',     password:'8434'},
    {nombre:'Arnau Serra Díaz',           telefono:'+34 648 393 315', nivel:4, username:'arnau.serra',        password:'3315'},
    {nombre:'Víctor Aguilar Cano',        telefono:'+34 780 556 262', nivel:4, username:'victor.aguilar',     password:'6262'},
    {nombre:'Cristian Álvarez Aguilar',   telefono:'+34 790 344 381', nivel:4, username:'cristian.alvarez',   password:'4381'},
    {nombre:'Xavi Serra Navarro',         telefono:'+34 748 871 499', nivel:4, username:'xavi.serra',         password:'1499'},
    {nombre:'Tomás Mendoza Bosch',        telefono:'+34 738 379 546', nivel:4, username:'tomas.mendoza',      password:'9546'},
    {nombre:'Alejandro Ramos Prat',       telefono:'+34 692 538 238', nivel:4, username:'alejandro.ramos',    password:'8238'},
    {nombre:'Rubén Sánchez Álvarez',      telefono:'+34 668 434 445', nivel:4, username:'ruben.sanchez2',     password:'4445'},
    {nombre:'Tomás Blanco Domínguez',     telefono:'+34 711 969 530', nivel:4, username:'tomas.blanco',       password:'9530'},
    {nombre:'Javier Muñoz Cabrera',       telefono:'+34 698 587 761', nivel:4, username:'javier.munoz',       password:'7761'},
    {nombre:'Guillermo Soler Costa',      telefono:'+34 780 438 460', nivel:4, username:'guillermo.soler',    password:'8460'},
    {nombre:'Xavi Serrano Font',          telefono:'+34 676 489 561', nivel:4, username:'xavi.serrano',       password:'9561'},
    {nombre:'Antonio Cano Prat',          telefono:'+34 699 532 448', nivel:3, username:'antonio.cano',       password:'2448'},
    {nombre:'Diego Jiménez Aguilar',      telefono:'+34 658 199 467', nivel:3, username:'diego.jimenez',      password:'9467'},
    {nombre:'Gerard Álvarez Martínez',    telefono:'+34 658 180 665', nivel:3, username:'gerard.alvarez',     password:'0665'},
    {nombre:'Jordi Camps Molina',         telefono:'+34 615 777 333', nivel:3, username:'jordi.camps',        password:'7333'},
    {nombre:'Rodrigo Álvarez Costa',      telefono:'+34 792 448 214', nivel:3, username:'rodrigo.alvarez',    password:'8214'},
    {nombre:'Daniel Ramírez Medina',      telefono:'+34 674 208 991', nivel:3, username:'daniel.ramirez',     password:'8991'},
    {nombre:'Carlos Font Castillo',       telefono:'+34 656 999 951', nivel:3, username:'carlos.font',        password:'9951'},
    {nombre:'Daniel Jiménez Bosch',       telefono:'+34 761 787 765', nivel:3, username:'daniel.jimenez',     password:'7765'},
    {nombre:'Javier Serrano Domínguez',   telefono:'+34 743 309 785', nivel:3, username:'javier.serrano',     password:'9785'},
    {nombre:'Óscar Camps Serra',          telefono:'+34 674 371 235', nivel:3, username:'oscar.camps',        password:'1235'},
    {nombre:'Tomás Martínez Molina',      telefono:'+34 665 230 142', nivel:3, username:'tomas.martinez',     password:'0142'},
    {nombre:'Arnau Mendoza Torres',       telefono:'+34 662 125 283', nivel:3, username:'arnau.mendoza',      password:'5283'},
    {nombre:'Mario Cabrera Castillo',     telefono:'+34 659 690 294', nivel:3, username:'mario.cabrera',      password:'0294'},
    {nombre:'Miguel Reyes Blanco',        telefono:'+34 656 541 171', nivel:3, username:'miguel.reyes',       password:'1171'},
    {nombre:'Diego Moreno Navarro',       telefono:'+34 618 524 517', nivel:3, username:'diego.moreno',       password:'4517'},
    {nombre:'Arnau Torres Muñoz',         telefono:'+34 682 323 576', nivel:3, username:'arnau.torres',       password:'3576'},
    {nombre:'Pablo Vargas Delgado',       telefono:'+34 682 901 142', nivel:3, username:'pablo.vargas',       password:'1142'},
    {nombre:'Jordi Aguilar Puig',         telefono:'+34 770 590 551', nivel:3, username:'jordi.aguilar',      password:'0551'},
    {nombre:'Iván García Serrano',        telefono:'+34 747 133 337', nivel:3, username:'ivan.garcia',        password:'3337'},
    {nombre:'Roger Mendoza Bosch',        telefono:'+34 714 130 195', nivel:2, username:'roger.mendoza',      password:'0195'},
    {nombre:'Javier Moreno Rodríguez',    telefono:'+34 791 954 473', nivel:2, username:'javier.moreno',      password:'4473'},
    {nombre:'Biel Muñoz Ramírez',         telefono:'+34 791 804 670', nivel:2, username:'biel.munoz',         password:'4670'},
    {nombre:'Jorge Serra Reyes',          telefono:'+34 760 758 569', nivel:2, username:'jorge.serra',        password:'8569'},
    {nombre:'Guillermo Moreno González',  telefono:'+34 616 981 212', nivel:2, username:'guillermo.moreno',   password:'1212'},
    {nombre:'Marcos Cabrera Serra',       telefono:'+34 670 926 516', nivel:2, username:'marcos.cabrera',     password:'6516'},
    {nombre:'Roger Santos Rodríguez',     telefono:'+34 723 354 296', nivel:2, username:'roger.santos',       password:'4296'},
    {nombre:'Cristian Costa Rodríguez',   telefono:'+34 621 871 969', nivel:2, username:'cristian.costa',     password:'1969'},
    {nombre:'Jaime Álvarez Medina',       telefono:'+34 729 294 403', nivel:2, username:'jaime.alvarez',      password:'4403'},
    {nombre:'Marc Mas García',            telefono:'+34 794 206 999', nivel:2, username:'marc.mas',           password:'6999'},
    {nombre:'Roberto Fernández Rodríguez',telefono:'+34 644 388 719', nivel:2, username:'roberto.fernandez',  password:'8719'},
    {nombre:'Jordi Blanco Romero',        telefono:'+34 726 752 368', nivel:2, username:'jordi.blanco',       password:'2368'},
    {nombre:'Ferran Cano Costa',          telefono:'+34 624 177 807', nivel:2, username:'ferran.cano',        password:'7807'},
    {nombre:'Nicolás Moreno Herrera',     telefono:'+34 739 328 124', nivel:2, username:'nicolas.moreno',     password:'8124'},
    {nombre:'Pau Ruiz Costa',             telefono:'+34 795 488 793', nivel:2, username:'pau.ruiz',           password:'8793'},
    {nombre:'Manuel Soler Camps',         telefono:'+34 763 744 689', nivel:2, username:'manuel.soler',       password:'4689'},
    {nombre:'Sergio Reyes Ruiz',          telefono:'+34 750 653 182', nivel:2, username:'sergio.reyes',       password:'3182'},
    {nombre:'Alejandro Díaz Serra',       telefono:'+34 735 177 706', nivel:2, username:'alejandro.diaz',     password:'7706'},
    {nombre:'Arnau Domínguez Navarro',    telefono:'+34 751 286 599', nivel:2, username:'arnau.dominguez',    password:'6599'},
    {nombre:'Mario Álvarez Vega',         telefono:'+34 781 110 629', nivel:2, username:'mario.alvarez',      password:'0629'},
    {nombre:'Oriol Bosch González',       telefono:'+34 680 399 814', nivel:2, username:'oriol.bosch',        password:'9814'},
    {nombre:'Raúl Prat Blanco',           telefono:'+34 688 857 201', nivel:2, username:'raul.prat',          password:'7201'},
    {nombre:'Jordi Fernández López',      telefono:'+34 638 242 854', nivel:1, username:'jordi.fernandez',    password:'2854'},
    {nombre:'Roger Castillo Delgado',     telefono:'+34 650 510 374', nivel:1, username:'roger.castillo',     password:'0374'},
    {nombre:'Matías Moreno Prat',         telefono:'+34 690 160 334', nivel:1, username:'matias.moreno',      password:'0334'},
    {nombre:'Pablo Rodríguez Domínguez',  telefono:'+34 762 578 984', nivel:1, username:'pablo.rodriguez',    password:'8984'},
    {nombre:'Tomás Cano González',        telefono:'+34 769 355 995', nivel:1, username:'tomas.cano',         password:'5995'},
    {nombre:'Iván Serrano Puig',          telefono:'+34 717 268 488', nivel:1, username:'ivan.serrano',       password:'6488'},
    {nombre:'Santiago Mas Ramírez',       telefono:'+34 633 170 709', nivel:1, username:'santiago.mas',       password:'0709'},
    {nombre:'Manuel Álvarez Ramos',       telefono:'+34 750 869 174', nivel:1, username:'manuel.alvarez',     password:'9174'},
    {nombre:'Oriol Vega Jiménez',         telefono:'+34 742 966 152', nivel:1, username:'oriol.vega',         password:'6152'},
    {nombre:'Eduardo Domínguez Mendoza',  telefono:'+34 699 210 491', nivel:1, username:'eduardo.dominguez',  password:'0491'},
    {nombre:'Héctor Gómez Aguilar',       telefono:'+34 635 250 125', nivel:1, username:'hector.gomez',       password:'0125'},
    {nombre:'Oriol Vargas Navarro',       telefono:'+34 628 771 804', nivel:1, username:'oriol.vargas',       password:'1804'},
    {nombre:'Xavi Delgado Rodríguez',     telefono:'+34 699 630 575', nivel:1, username:'xavi.delgado',       password:'0575'},
    {nombre:'Tomás Medina Serrano',       telefono:'+34 759 888 698', nivel:1, username:'tomas.medina',       password:'8698'},
    {nombre:'Fernando Aguilar Serrano',   telefono:'+34 712 867 652', nivel:1, username:'fernando.aguilar',   password:'7652'},
    {nombre:'Mario Moreno Santos',        telefono:'+34 641 304 959', nivel:1, username:'mario.moreno',       password:'4959'},
    {nombre:'Gerard Ruiz Rodríguez',      telefono:'+34 656 283 409', nivel:1, username:'gerard.ruiz',        password:'3409'},
    {nombre:'Jordi Jiménez Álvarez',      telefono:'+34 640 277 665', nivel:1, username:'jordi.jimenez',      password:'7665'},
  ];
}

// ------------------------------------------------------------
// RESET FASE — Fuerza la fase activa a un número concreto
// Usar con: ?action=resetFase&fase=2&adminKey=ctsg2026admin
// ------------------------------------------------------------
function resetFase(fase) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  setConfigValue(ss, 'fase_activa', String(fase));
  return {ok: true, msg: `✅ Fase activa reseteada a ${fase}`};
}

function cancelMatch(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Partidos');
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const p1id = parseInt(p.p1id);
  const p2id = parseInt(p.p2id);
  const fase = parseInt(p.fase || 1);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rFase = parseInt(row[h.indexOf('fase')]);
    const rP1   = parseInt(row[h.indexOf('p1_id')]);
    const rP2   = parseInt(row[h.indexOf('p2_id')]);
    const rEst  = row[h.indexOf('estado')];
    if (rFase !== fase) continue;
    if (rEst === 'played') continue;
    if (!((rP1 === p1id && rP2 === p2id) || (rP1 === p2id && rP2 === p1id))) continue;
    // Guardar scores -1/-1 -1/-1 (penalización de -4 pts a cada jugador)
    sh.getRange(i + 1, h.indexOf('s1p1') + 1).setValue(-1);
    sh.getRange(i + 1, h.indexOf('s1p2') + 1).setValue(-1);
    sh.getRange(i + 1, h.indexOf('s2p1') + 1).setValue(-1);
    sh.getRange(i + 1, h.indexOf('s2p2') + 1).setValue(-1);
    sh.getRange(i + 1, h.indexOf('s3p1') + 1).setValue('');
    sh.getRange(i + 1, h.indexOf('s3p2') + 1).setValue('');
    sh.getRange(i + 1, h.indexOf('estado') + 1).setValue('cancelled');
    sh.getRange(i + 1, h.indexOf('fecha') + 1).setValue(new Date().toISOString());
    sh.getRange(i + 1, h.indexOf('anotado_por') + 1).setValue('cancelado');
    return {ok: true, msg: '❌ Partido cancelado'};
  }
  return {ok: false, error: 'Partido no encontrado o ya jugado'};
}

// ------------------------------------------------------------
// SIMULAR RESULTADOS — Genera marcadores aleatorios para todos
// los partidos pendientes excepto los de Marc (id=0)
// Ejecutar desde: Apps Script → seleccionar simularResultados → Ejecutar
// O llamar vía URL con ?action=simularResultados&adminKey=ctsg2026admin
// ------------------------------------------------------------
function simularResultados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const partSh = ss.getSheetByName('Partidos');
  const data = partSh.getDataRange().getValues();
  const h = data[0];

  const iEstado = h.indexOf('estado');
  const iP1     = h.indexOf('p1_id');
  const iP2     = h.indexOf('p2_id');
  const iS1p1   = h.indexOf('s1p1');
  const iS1p2   = h.indexOf('s1p2');
  const iS2p1   = h.indexOf('s2p1');
  const iS2p2   = h.indexOf('s2p2');
  const iS3p1   = h.indexOf('s3p1');
  const iS3p2   = h.indexOf('s3p2');
  const iFecha  = h.indexOf('fecha');
  const iAnotBy = h.indexOf('anotado_por');

  function randomSet() {
    const scores = [[6,0],[6,1],[6,2],[6,3],[6,4],[7,5],[7,6]];
    return scores[Math.floor(Math.random() * scores.length)];
  }

  let actualizados = 0;
  const ahora = new Date().toISOString();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const p1id = parseInt(row[iP1]);
    const p2id = parseInt(row[iP2]);
    const estado = row[iEstado];

    // Saltar partidos ya jugados o los de Marc (id=0)
    if (estado === 'played' || p1id === 0 || p2id === 0) continue;

    // Generar sets
    const goTo3 = Math.random() < 0.30;
    let setsP1 = 0, setsP2 = 0;

    const [w1, l1] = randomSet();
    const p1WinsS1 = Math.random() < 0.5;
    const s1p1v = p1WinsS1 ? w1 : l1;
    const s1p2v = p1WinsS1 ? l1 : w1;
    if (p1WinsS1) setsP1++; else setsP2++;

    const p1WinsS2 = goTo3 ? !p1WinsS1 : p1WinsS1;
    const [w2, l2] = randomSet();
    const s2p1v = p1WinsS2 ? w2 : l2;
    const s2p2v = p1WinsS2 ? l2 : w2;
    if (p1WinsS2) setsP1++; else setsP2++;

    let s3p1v = '', s3p2v = '';
    if (setsP1 === 1 && setsP2 === 1) {
      const p1WinsS3 = Math.random() < 0.5;
      const [w3, l3] = randomSet();
      s3p1v = p1WinsS3 ? w3 : l3;
      s3p2v = p1WinsS3 ? l3 : w3;
    }

    // Escribir en la fila
    partSh.getRange(i + 1, iS1p1 + 1).setValue(s1p1v);
    partSh.getRange(i + 1, iS1p2 + 1).setValue(s1p2v);
    partSh.getRange(i + 1, iS2p1 + 1).setValue(s2p1v);
    partSh.getRange(i + 1, iS2p2 + 1).setValue(s2p2v);
    partSh.getRange(i + 1, iS3p1 + 1).setValue(s3p1v);
    partSh.getRange(i + 1, iS3p2 + 1).setValue(s3p2v);
    partSh.getRange(i + 1, iEstado + 1).setValue('played');
    partSh.getRange(i + 1, iFecha  + 1).setValue(ahora);
    partSh.getRange(i + 1, iAnotBy + 1).setValue('simulado');
    actualizados++;
  }

  return {ok: true, msg: `✅ ${actualizados} partidos simulados y guardados en Google Sheets`};
}
