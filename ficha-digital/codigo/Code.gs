/**
 * Ficha digital de examen con IA — plantilla para instalar en su propia cuenta de Google
 * Autor: Ph.D. José Herrera · https://orcid.org/0000-0002-8207-9714 · https://jherreraxp.github.io
 * Uso educativo libre citando al autor.
 *
 * Script vinculado a una hoja de cálculo de Google. Menú «Examen» (se crea al abrir la hoja):
 *   Preparar hojas · Configurar clave de IA · Elegir proveedor · Cambiar código docente · Probar IA.
 * Secretos (clave de IA y código docente) en Propiedades del script: nunca en las hojas ni en el navegador.
 */

var HOJAS = {
  Config: ['examen_id', 'titulo', 'curso', 'n_preg', 'minutos', 'max_infracciones', 'nota_aprobatoria',
           'mostrar_correctas', 'mostrar_ranking', 'activo', 'clave'],
  Alumnos: ['codigo', 'dni', 'apellidos', 'nombres', 'seccion', 'email', 'origen'],
  Respuestas: ['marca_tiempo', 'examen_id', 'codigo', 'dni', 'apellidos', 'nombres', 'aula', 'verificacion', 'respuestas',
               'correctas', 'n_preg', 'nota', 'estado', 'infracciones', 'desbloqueos', 'tiempo_seg', 'inicio', 'fin'],
  Eventos: ['marca_tiempo', 'examen_id', 'codigo', 'tipo', 'detalle', 'fuera_seg'],
  Intentos: ['examen_id', 'codigo', 'dni', 'apellidos', 'nombres', 'aula', 'verificacion', 'token', 'inicio_ms', 'estado',
             'respuestas', 'infracciones', 'desbloqueos', 'ultimo_guardado'],
  Archivo: ['archivado', 'motivo', 'hoja', 'examen_id', 'codigo', 'datos']
};

var CONFIG_INICIAL = [
  ['DEMO-A', 'Examen de prueba · Versión A', 'Curso de ejemplo', 20, 90, 5, 10.5, 'SI', 'SI', 'SI', 'ABCDEABCDEABCDEABCDE'],
  ['DEMO-B', 'Examen de prueba · Versión B', 'Curso de ejemplo', 20, 90, 5, 10.5, 'SI', 'SI', 'SI', 'EDCBAEDCBAEDCBAEDCBA']
];

var GRACIA_SEG = 90;
var CODIGO_DOCENTE_INICIAL = '';   // se genera al azar en preparar(); cámbielo en el panel

/* ================================================================== menú y preparación */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Examen')
    .addItem('1. Preparar hojas', 'preparar')
    .addItem('2. Configurar clave de IA (Claude o Gemini)', 'menuClaveIA')
    .addItem('3. Elegir proveedor de IA', 'menuProveedor')
    .addItem('4. Cambiar código docente', 'menuCodigoDocente')
    .addItem('5. Probar la IA', 'menuProbarIA')
    .addToUi();
}

function preparar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try { ss.setSpreadsheetTimeZone('America/Lima'); } catch (e) {}
  Object.keys(HOJAS).forEach(function (nombre) {
    var sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HOJAS[nombre].length).setValues([HOJAS[nombre]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  var cfg = ss.getSheetByName('Config');
  cfg.getRange(2, HOJAS.Config.indexOf('clave') + 1, 50, 1).setNumberFormat('@');
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, CONFIG_INICIAL.length, CONFIG_INICIAL[0].length).setValues(CONFIG_INICIAL);
  ss.getSheetByName('Alumnos').getRange(2, 1, 500, 2).setNumberFormat('@');   // código y DNI como texto (ceros a la izquierda)
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('CODIGO_DOCENTE')) p.setProperty('CODIGO_DOCENTE', String(Math.floor(100000 + Math.random() * 900000)));
  if (!p.getProperty('IA_PROVEEDOR')) p.setProperty('IA_PROVEEDOR', 'auto');
  var al = ss.getSheetByName('Alumnos');
  if (al.getLastRow() < 2) al.appendRow(['DEMO0001', '', 'Estudiante de Prueba', 'Ana', 'A', '', 'Ejemplo (bórrelo)']);
  var m = 'Listo: hojas creadas con un examen de prueba. Su código docente inicial es ' + p.getProperty('CODIGO_DOCENTE') +
          ' (también es la clave del panel hasta que la cambie). Anótelo y cámbielo en el panel de administración.';
  Logger.log(m);
  try { SpreadsheetApp.getUi().alert(m); } catch (e) {}
  return m;
}

function menuClaveIA() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Clave de IA', 'Pegue una clave de Claude (sk-ant-…) o de Gemini (AIza…). Se guarda en las propiedades del script.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  ui.alert(configurarIA(r.getResponseText()));
}
function menuProveedor() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Proveedor de IA', 'Escriba: auto (Claude si hay clave; si falla, Gemini), claude o gemini.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var v = String(r.getResponseText()).trim().toLowerCase();
  if (['auto', 'claude', 'gemini'].indexOf(v) < 0) { ui.alert('Valor no válido.'); return; }
  PropertiesService.getScriptProperties().setProperty('IA_PROVEEDOR', v);
  ui.alert('Proveedor: ' + v);
}
function menuCodigoDocente() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Código docente', 'Nuevo código numérico (4 a 8 dígitos):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var v = String(r.getResponseText()).trim();
  if (!/^\d{4,8}$/.test(v)) { ui.alert('Debe tener entre 4 y 8 dígitos.'); return; }
  PropertiesService.getScriptProperties().setProperty('CODIGO_DOCENTE', v);
  ui.alert('Código docente actualizado.');
}
function menuProbarIA() { SpreadsheetApp.getUi().alert(probarIA()); }

/** Guarda la clave en la ranura del proveedor que corresponde (detección por prefijo). */
function configurarIA(clave) {
  clave = String(clave || '').trim();
  var prov = proveedorDeClave_(clave);
  if (!prov) return 'No reconozco la clave: las de Claude empiezan con «sk-ant-» y las de Gemini con «AIza».';
  var p = PropertiesService.getScriptProperties();
  p.setProperty(prov === 'claude' ? 'IA_CLAVE_CLAUDE' : 'IA_CLAVE_GEMINI', clave);
  CacheService.getScriptCache().remove('IA_MODELO_' + prov);
  return 'Clave de ' + (prov === 'claude' ? 'Claude' : 'Gemini') + ' guardada. ' + probarIA();
}

function probarIA() {
  var out = [];
  ['claude', 'gemini'].forEach(function (prov) {
    if (!claveDe_(prov)) return;
    try { out.push(prov + ': modelo ' + modelo_(prov, true)); } catch (e) { out.push(prov + ': ' + e.message); }
  });
  return out.length ? out.join(' · ') : 'No hay claves de IA configuradas.';
}

function doGet(e) {
  if (e && e.parameter && e.parameter.admin !== undefined) {
    return HtmlService.createHtmlOutputFromFile('Admin').setTitle('Administración · Ficha digital')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ficha de respuestas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ================================================================== utilidades de hoja */
function hoja_(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(nombre);
  if (!sh && HOJAS[nombre]) {                       // hojas nuevas de versiones posteriores se crean solas
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, HOJAS[nombre].length).setValues([HOJAS[nombre]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function filas_(nombre) {
  var sh = hoja_(nombre), n = sh.getLastRow();
  if (n < 2) return [];
  var h = HOJAS[nombre], v = sh.getRange(2, 1, n - 1, h.length).getValues();
  return v.map(function (r, i) { var o = { _fila: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o; });
}
function agregar_(nombre, obj) { hoja_(nombre).appendRow(HOJAS[nombre].map(function (k) { return obj[k] === undefined ? '' : obj[k]; })); }
function actualizar_(nombre, fila, obj) {
  var sh = hoja_(nombre), h = HOJAS[nombre];
  Object.keys(obj).forEach(function (k) { var j = h.indexOf(k); if (j >= 0) sh.getRange(fila, j + 1).setValue(obj[k]); });
}
function borrarFilas_(nombre, filas) {
  var sh = hoja_(nombre);
  filas.slice().sort(function (a, b) { return b - a; }).forEach(function (f) { sh.deleteRow(f); });
}
function archivar_(motivo, nombre, filas) {
  filas.forEach(function (r) {
    var o = {}; HOJAS[nombre].forEach(function (k) { o[k] = r[k] instanceof Date ? r[k].toISOString() : r[k]; });
    agregar_('Archivo', { archivado: new Date(), motivo: motivo, hoja: nombre, examen_id: String(r.examen_id || ''), codigo: String(r.codigo || ''), datos: JSON.stringify(o) });
  });
}
function si_(v) { return String(v).trim().toUpperCase() === 'SI'; }
function limpiar_(s, max) { return String(s === undefined || s === null ? '' : s).replace(/[\u0000-\u001f<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, max || 80); }
function normCod_(s) { return String(s || '').toUpperCase().replace(/[^0-9A-Z]/g, ''); }
function normDni_(s) { var d = String(s || '').replace(/\D/g, ''); return d ? ('00000000' + d).slice(-8) : ''; }
function normTxt_(s) { return String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-ZÑ ]/g, ' ').replace(/\s+/g, ' ').trim(); }
function config_(id) {
  var c = filas_('Config').filter(function (r) { return String(r.examen_id) === String(id); })[0];
  if (!c) throw new Error('El examen no existe.');
  return c;
}
function grupo_(c) { return String(c.curso) + '|' + String(c.titulo).split('·')[0].trim(); }
function conLock_(fn) { var l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
function evento_(examen, codigo, tipo, detalle, seg) {
  agregar_('Eventos', { marca_tiempo: new Date(), examen_id: examen, codigo: codigo, tipo: tipo, detalle: limpiar_(detalle, 160), fuera_seg: seg === undefined || seg === null ? '' : seg });
}

/* ================================================================== código docente (validado en el servidor) */
function codigoDocenteOk_(clave, contexto) {
  var cache = CacheService.getScriptCache(), k = 'FALLOS_' + normCod_(contexto);
  var fallos = Number(cache.get(k) || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos con el código docente. Espere unos minutos.');
  var ok = String(clave || '').trim() !== '' && String(clave).trim() === String(PropertiesService.getScriptProperties().getProperty('CODIGO_DOCENTE') || CODIGO_DOCENTE_INICIAL || '\u0000');
  if (!ok && String(clave || '').trim() !== '') cache.put(k, String(fallos + 1), 600);
  return ok;
}

/* ================================================================== IA: Claude o Gemini, modelo detectado automáticamente */
function proveedorDeClave_(k) { return /^sk-ant-/.test(k) ? 'claude' : /^AIza/.test(k) ? 'gemini' : null; }
function claveDe_(prov) { return PropertiesService.getScriptProperties().getProperty(prov === 'claude' ? 'IA_CLAVE_CLAUDE' : 'IA_CLAVE_GEMINI'); }
function proveedores_() {
  var pref = String(PropertiesService.getScriptProperties().getProperty('IA_PROVEEDOR') || 'auto');
  var lista = pref === 'gemini' ? ['gemini', 'claude'] : ['claude', 'gemini'];
  if (pref !== 'auto') lista = [pref];
  return lista.filter(function (p) { return !!claveDe_(p); });
}

/** Elige el modelo más reciente con entrada de imagen. Nunca fija una versión: la consulta a la API. */
function modelo_(prov, forzar) {
  var cache = CacheService.getScriptCache(), ck = 'IA_MODELO_' + prov;
  if (!forzar) { var c = cache.get(ck); if (c) return c; }
  var m;
  if (prov === 'claude') {
    var r = UrlFetchApp.fetch('https://api.anthropic.com/v1/models?limit=100', {
      headers: { 'x-api-key': claveDe_('claude'), 'anthropic-version': '2023-06-01' }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('Claude respondió ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 160));
    var ms = JSON.parse(r.getContentText()).data || [];
    var vis = ms.filter(function (x) { var c = x.capabilities || {}; return !c.image_input || c.image_input.supported !== false; });
    vis.sort(function (a, b) { return String(b.created_at).localeCompare(String(a.created_at)); });
    var rapidos = vis.filter(function (x) { return x.line === 'haiku' || /haiku/.test(x.id); });  // lectura rápida y barata
    m = (rapidos[0] || vis[0] || {}).id;
  } else {
    var g = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200&key=' + encodeURIComponent(claveDe_('gemini')), { muteHttpExceptions: true });
    if (g.getResponseCode() !== 200) throw new Error('Gemini respondió ' + g.getResponseCode() + ': ' + g.getContentText().slice(0, 160));
    var gs = (JSON.parse(g.getContentText()).models || []).filter(function (x) {
      return (x.supportedGenerationMethods || []).indexOf('generateContent') >= 0 && /gemini/.test(x.name) &&
             !/(tts|image|live|embedding|audio|robotics|computer|exp)/.test(x.name);
    });
    var ver = function (n) { var k = /gemini-(\d+(?:\.\d+)?)/.exec(n); return k ? parseFloat(k[1]) : 0; };
    var peso = function (x) { return ver(x.name) * 10 + (/flash/.test(x.name) ? 2 : 0) - (/lite/.test(x.name) ? 1 : 0) - (/preview/.test(x.name) ? 0.5 : 0); };
    gs.sort(function (a, b) { return peso(b) - peso(a); });
    m = gs[0] ? gs[0].name.replace(/^models\//, '') : null;
  }
  if (!m) throw new Error('No encontré un modelo con visión para ' + prov + '.');
  cache.put(ck, m, 21600);
  return m;
}

var PROMPT_DOC = 'La imagen debería ser un DNI peruano o un carné universitario. Extrae los datos del titular y responde SOLO con un objeto JSON, sin texto adicional, con estas claves: ' +
  '{"tipo":"DNI|carné universitario|otro|ninguno","dni":"8 dígitos o vacío","codigo":"código universitario (por ejemplo 20231234A) o vacío",' +
  '"apellidos":"apellido paterno y materno","nombres":"pre-nombres","legible":true}. ' +
  'Si un dato no se lee con seguridad, déjalo vacío. Si la imagen no es un documento de identidad o no se puede leer, usa "legible":false. No inventes datos.';

function llamarIA_(prov, imagenes, prompt) {
  var modelo = modelo_(prov), txt;
  if (prov === 'claude') {
    var content = imagenes.map(function (im) { return { type: 'image', source: { type: 'base64', media_type: im.mime, data: im.b64 } }; });
    content.push({ type: 'text', text: prompt });
    var r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': claveDe_('claude'), 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: modelo, max_tokens: 3000, messages: [{ role: 'user', content: content }] })   // margen para el razonamiento del modelo
    });
    if (r.getResponseCode() === 404) { CacheService.getScriptCache().remove('IA_MODELO_claude'); }
    if (r.getResponseCode() !== 200) throw new Error('Claude ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
    txt = (JSON.parse(r.getContentText()).content || []).map(function (c) { return c.type === 'text' ? c.text : ''; }).join('');
  } else {
    var parts = imagenes.map(function (im) { return { inline_data: { mime_type: im.mime, data: im.b64 } }; });
    parts.push({ text: prompt });
    var g = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent?key=' + encodeURIComponent(claveDe_('gemini')), {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ contents: [{ parts: parts }], generationConfig: { temperature: 0, responseMimeType: 'application/json' } })
    });
    if (g.getResponseCode() === 404) { CacheService.getScriptCache().remove('IA_MODELO_gemini'); }
    if (g.getResponseCode() !== 200) throw new Error('Gemini ' + g.getResponseCode() + ': ' + g.getContentText().slice(0, 200));
    var cand = (JSON.parse(g.getContentText()).candidates || [])[0];
    txt = cand && cand.content ? (cand.content.parts || []).map(function (q) { return q.text || ''; }).join('') : '';
  }
  var j = /\{[\s\S]*\}/.exec(txt);
  if (!j) throw new Error('La IA no devolvió datos legibles.');
  return { datos: JSON.parse(j[0]), modelo: modelo };
}

/** Recorre los proveedores configurados hasta que uno responda. */
function conIA_(imagenes, prompt) {
  var provs = proveedores_(), err = [];
  if (!provs.length) throw new Error('La IA no está configurada (panel de administración).');
  for (var i = 0; i < provs.length; i++) {
    try { var r = llamarIA_(provs[i], imagenes, prompt); r.proveedor = provs[i]; return r; } catch (e) { err.push(e.message); }
  }
  throw new Error(err.join(' | ').slice(0, 220));
}
function imagenOk_(b64, mime) {
  b64 = String(b64 || ''); mime = /^image\/(jpeg|png|webp)$/.test(mime) ? mime : 'image/jpeg';
  if (b64.length < 1000 || b64.length > 4e6) throw new Error('La foto no es válida o es demasiado grande.');
  return { b64: b64, mime: mime };
}

/** Lee el DNI o carné de la foto. La imagen no se guarda: solo se envía a la IA para leerla. */
function leerDocumento(b64, mime) {
  try {
    var r = conIA_([imagenOk_(b64, mime)], PROMPT_DOC), d = r.datos || {};
    var datos = { tipo: limpiar_(d.tipo, 30), dni: normDni_(d.dni), codigo: normCod_(d.codigo).slice(0, 12),
                  apellidos: limpiar_(d.apellidos), nombres: limpiar_(d.nombres) };
    var legible = d.legible !== false && !!(datos.dni || datos.codigo || datos.apellidos);
    var coincide = legible ? buscarAlumno_(datos) : null;
    if (coincide) datos.seccion = String(coincide.seccion || '');
    return { ok: legible, datos: datos, proveedor: r.proveedor, modelo: r.modelo, enBase: !!coincide,
             error: legible ? '' : 'No se pudo leer el documento. Tome otra foto con buena luz o escriba sus datos.' };
  } catch (e) { return { ok: false, error: e.message + ' Escriba sus datos.' }; }
}

/* ================================================================== comparación de rostros (documento vs. foto actual) */
var PROMPT_ROSTRO = 'Control de identidad en un examen universitario, con consentimiento del estudiante. La primera imagen es la foto de su documento ' +
  '(DNI o carné) y la segunda es una foto tomada ahora con la cámara del celular. No identifiques a nadie ni digas quién es: solo compara ' +
  'si las dos fotos parecen mostrar a la misma persona. Si alguna imagen no muestra un rostro con claridad, responde "dudoso". ' +
  'Responde SOLO JSON: {"rostro_documento":true,"rostro_actual":true,"misma_persona":"si|no|dudoso","motivo":"breve"}';

function rostroRequerido_() { return String(PropertiesService.getScriptProperties().getProperty('VERIFICAR_ROSTRO') || 'SI') === 'SI'; }
function estadoRostro_(codigo) {
  if (!rostroRequerido_()) return 'off';
  return CacheService.getScriptCache().get('ROSTRO_' + normCod_(codigo)) || 'pendiente';
}

/** Compara la foto del documento con la foto actual. Ninguna imagen se guarda. */
function compararRostros(d, docB64, selfieB64) {
  d = d || {};
  var a = buscarAlumno_(d), cod = normCod_(a ? a.codigo : d.codigo);
  if (!cod) return { ok: false, resultado: 'pendiente', error: 'Escriba primero su código (o tome la foto del documento).' };
  if (!rostroRequerido_()) return { ok: true, resultado: 'off', motivo: 'La comparación de rostros está desactivada.' };
  var res = 'dudoso', motivo = '';
  try {
    var r = conIA_([imagenOk_(docB64, 'image/jpeg'), imagenOk_(selfieB64, 'image/jpeg')], PROMPT_ROSTRO), x = r.datos || {};
    var v = String(x.misma_persona || '').toLowerCase();
    res = (v === 'si' || v === 'sí') && x.rostro_documento !== false && x.rostro_actual !== false ? 'si' : v === 'no' ? 'no' : 'dudoso';
    motivo = limpiar_(x.motivo, 200);
  } catch (e) { motivo = 'No se pudo comparar: ' + e.message; }
  CacheService.getScriptCache().put('ROSTRO_' + cod, res, 10800);
  evento_('', cod, 'rostro: ' + res, motivo);
  return { ok: res === 'si', resultado: res, motivo: motivo };
}

/* ================================================================== verificación con la base de alumnos */
function buscarAlumno_(d) {
  var cod = normCod_(d.codigo), dni = normDni_(d.dni), ape = normTxt_(d.apellidos);
  var base = filas_('Alumnos');
  var por = function (f) { return base.filter(f)[0]; };
  return (cod && por(function (a) { return normCod_(a.codigo) === cod; })) ||
         (dni && por(function (a) { return normDni_(a.dni) === dni; })) ||
         (ape && ape.length > 5 && normTxt_(d.nombres) && por(function (a) {
            return normTxt_(a.apellidos) === ape && normTxt_(a.nombres).split(' ')[0] === normTxt_(d.nombres).split(' ')[0]; })) || null;
}

/** Devuelve el alumno oficial (de la base) o exige la conformidad del docente. */
function verificar_(d) {
  var a = buscarAlumno_(d);
  var rostro = estadoRostro_(a ? a.codigo : d.codigo), rostroOk = rostro === 'si' || rostro === 'off';
  var docente = null;
  var docenteOk = function () { if (docente === null) docente = !!d.claveDocente && codigoDocenteOk_(d.claveDocente, d.codigo || d.dni || 'anon'); return docente; };
  var msgRostro = rostro === 'pendiente' ? 'Falta la foto de su rostro para compararla con su documento.' :
                  rostro === 'no' ? 'Su foto no coincide con la de su documento. El docente debe comparar ambas fotos y dar conformidad.' :
                  'La comparación de su rostro no fue concluyente. El docente debe comparar ambas fotos y dar conformidad.';
  if (a) {
    var al = { codigo: normCod_(a.codigo) || normCod_(d.codigo), dni: normDni_(a.dni) || normDni_(d.dni),
               apellidos: limpiar_(a.apellidos), nombres: limpiar_(a.nombres), seccion: limpiar_(a.seccion, 20) };
    var fuente = /docente/i.test(String(a.origen)) ? 'Agregado por el docente' : 'Base de alumnos';
    if (rostroOk) return { ok: true, fuente: fuente + (rostro === 'si' ? ' · rostro verificado' : ''), rostro: rostro, alumno: al };
    if (docenteOk()) return { ok: true, fuente: fuente + ' · identidad confirmada por el docente', rostro: rostro, alumno: al };
    return { ok: false, requiereDocente: true, rostro: rostro, error: d.claveDocente ? 'Código docente incorrecto.' : msgRostro };
  }
  var vacia = filas_('Alumnos').length === 0;
  if (docenteOk()) {
    var cod = normCod_(d.codigo), ape = limpiar_(d.apellidos), nom = limpiar_(d.nombres);
    if (!cod || !ape || !nom) return { ok: false, error: 'Complete código, apellidos y nombres.' };
    // Conformidad del docente en el aula: el alumno se agrega a la base (una sola vez)
    var aula = limpiar_(d.aula, 20);
    conLock_(function () {
      if (!buscarAlumno_({ codigo: cod })) {
        agregar_('Alumnos', { codigo: cod, dni: normDni_(d.dni), apellidos: ape, nombres: nom, seccion: aula,
                              email: '', origen: 'Agregado por el docente · ' + Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm') });
        evento_('', cod, 'alta en la base', ape + ', ' + nom + (aula ? ' · ' + aula : ''));
      }
    });
    return { ok: true, fuente: 'Agregado por el docente', rostro: rostro, alumno: { codigo: cod, dni: normDni_(d.dni), apellidos: ape, nombres: nom, seccion: aula } };
  }
  return { ok: false, requiereDocente: true, rostro: rostro,
           error: d.claveDocente ? 'Código docente incorrecto.' : vacia ? 'La base de alumnos aún no está cargada: pida al docente su conformidad.' :
                  'No está en la lista del curso. Revise su código o DNI, o pida al docente que lo agregue.' };
}

function verificarAlumno(d) {
  try { return verificar_(d || {}); } catch (e) { return { ok: false, requiereDocente: true, error: e.message }; }
}

/* ================================================================== API del examen */
function getExamenes() {
  return filas_('Config').filter(function (c) { return si_(c.activo); }).map(function (c) {
    return { id: String(c.examen_id), titulo: String(c.titulo), curso: String(c.curso), n: Number(c.n_preg),
             minutos: Number(c.minutos), maxInf: Number(c.max_infracciones) };
  });
}

function iniciar(d) {
  d = d || {};
  var c = config_(d.examen_id);
  if (!si_(c.activo)) throw new Error('El examen no está activo.');
  var v = verificar_(d);
  if (!v.ok) throw new Error(v.error);
  var al = v.alumno, aula = limpiar_(d.aula, 20) || al.seccion;
  return conLock_(function () {
    var grupo = grupo_(c);
    var ids = filas_('Config').filter(function (x) { return grupo_(x) === grupo; }).map(function (x) { return String(x.examen_id); });
    var env = filas_('Respuestas').filter(function (r) { return normCod_(r.codigo) === al.codigo && ids.indexOf(String(r.examen_id)) >= 0; })[0];
    if (env) return { yaEnviado: true, resultado: resultado_(env) };
    var otros = filas_('Intentos').filter(function (r) { return normCod_(r.codigo) === al.codigo && ids.indexOf(String(r.examen_id)) >= 0; });
    var prev = otros.filter(function (r) { return String(r.examen_id) === String(c.examen_id); })[0];
    if (!prev && otros.length) throw new Error('Este código ya inició otra versión de este examen (' + otros[0].examen_id + ').');
    var ahora = Date.now();
    if (!prev) {
      prev = { examen_id: String(c.examen_id), codigo: al.codigo, dni: al.dni, apellidos: al.apellidos, nombres: al.nombres, aula: aula,
               verificacion: v.fuente, token: Utilities.getUuid(), inicio_ms: ahora, estado: 'En curso', respuestas: '',
               infracciones: 0, desbloqueos: 0, ultimo_guardado: new Date() };
      agregar_('Intentos', prev);
      evento_(prev.examen_id, al.codigo, 'inicio', v.fuente + (aula ? ' · ' + aula : ''));
    } else evento_(prev.examen_id, al.codigo, 'reanudación', String(prev.estado));
    var restante = Math.max(0, Math.round(Number(c.minutos) * 60 - (ahora - Number(prev.inicio_ms)) / 1000));
    return { token: String(prev.token), examen_id: String(c.examen_id), titulo: String(c.titulo), curso: String(c.curso),
             n: Number(c.n_preg), maxInf: Number(c.max_infracciones), restante: restante, respuestas: String(prev.respuestas || ''),
             infracciones: Number(prev.infracciones || 0), bloqueado: String(prev.estado) === 'Bloqueado',
             alumno: al.nombres + ' ' + al.apellidos, verificacion: v.fuente, reanudado: otros.length > 0 };
  });
}

function intento_(token) {
  var it = filas_('Intentos').filter(function (r) { return String(r.token) === String(token); })[0];
  if (!it) throw new Error('Intento no encontrado.');
  return it;
}
function normalizar_(r, n) { return String(r || '').toUpperCase().replace(/[^A-E\-]/g, '-').slice(0, n); }

function guardarProgreso(token, respuestas) {
  return conLock_(function () {
    var it = intento_(token);
    if (['En curso', 'Bloqueado'].indexOf(String(it.estado)) < 0) return { ok: false };
    actualizar_('Intentos', it._fila, { respuestas: normalizar_(respuestas, 40), ultimo_guardado: new Date() });
    return { ok: true };
  });
}

/** Registra un evento. Las fugas se cuentan en el servidor; al llegar al máximo la hoja queda bloqueada. */
function registrarEvento(token, tipo, detalle, fueraSeg) {
  return conLock_(function () {
    var it = intento_(token), c = config_(it.examen_id);
    evento_(it.examen_id, it.codigo, limpiar_(tipo, 30), detalle, fueraSeg === undefined || fueraSeg === null ? '' : Number(fueraSeg));
    var inf = Number(it.infracciones || 0), estado = String(it.estado);
    if (/fuga/i.test(tipo) && estado === 'En curso') {
      inf++;
      var cambios = { infracciones: inf };
      if (inf >= Number(c.max_infracciones)) { cambios.estado = estado = 'Bloqueado'; evento_(it.examen_id, it.codigo, 'bloqueo', inf + ' fugas'); }
      actualizar_('Intentos', it._fila, cambios);
    }
    return { infracciones: inf, bloqueado: estado === 'Bloqueado' };
  });
}

function desbloquear(token, clave) {
  return conLock_(function () {
    var it = intento_(token);
    if (!codigoDocenteOk_(clave, it.codigo)) { evento_(it.examen_id, it.codigo, 'código docente incorrecto', 'desbloqueo'); return { ok: false, error: 'Código incorrecto.' }; }
    var n = Number(it.desbloqueos || 0) + 1;
    if (String(it.estado) === 'Bloqueado') actualizar_('Intentos', it._fila, { estado: 'En curso', desbloqueos: n });
    evento_(it.examen_id, it.codigo, 'desbloqueo docente', 'desbloqueo ' + n);
    return { ok: true, desbloqueos: n };
  });
}

function enviar(token, respuestas, motivo) {
  return conLock_(function () {
    var it = intento_(token);
    var prev = filas_('Respuestas').filter(function (r) { return String(r.examen_id) === String(it.examen_id) && normCod_(r.codigo) === normCod_(it.codigo); })[0];
    if (prev) return resultado_(prev);
    var c = config_(it.examen_id), n = Number(c.n_preg), clave = String(c.clave).toUpperCase();
    var r = normalizar_(respuestas, n); while (r.length < n) r += '-';
    var ok = 0; for (var i = 0; i < n; i++) if (r[i] === clave[i]) ok++;
    var fin = Date.now(), seg = Math.round((fin - Number(it.inicio_ms)) / 1000);
    var estado = motivo === 'tiempo' ? 'Auto-envío (tiempo)' : 'Enviado';
    if (String(it.estado) === 'Bloqueado') estado += ' · bloqueado';
    if (seg > Number(c.minutos) * 60 + GRACIA_SEG) estado = 'Fuera de tiempo';
    var fila = { marca_tiempo: new Date(), examen_id: String(it.examen_id), codigo: String(it.codigo), dni: String(it.dni || ''),
                 apellidos: it.apellidos, nombres: it.nombres, aula: it.aula, verificacion: it.verificacion, respuestas: r, correctas: ok,
                 n_preg: n, nota: Math.round(ok / n * 20 * 100) / 100, estado: estado, infracciones: Number(it.infracciones || 0),
                 desbloqueos: Number(it.desbloqueos || 0), tiempo_seg: seg, inicio: new Date(Number(it.inicio_ms)), fin: new Date(fin) };
    agregar_('Respuestas', fila);
    actualizar_('Intentos', it._fila, { estado: estado, respuestas: r, ultimo_guardado: new Date() });
    evento_(it.examen_id, it.codigo, 'envío', estado);
    return resultado_(fila);
  });
}

function resultado_(fila) {
  var c = config_(fila.examen_id), n = Number(c.n_preg), clave = String(c.clave).toUpperCase();
  var nota = Number(fila.nota), aprobado = nota >= Number(c.nota_aprobatoria);
  var res = { nota: nota, correctas: Number(fila.correctas), n: n, aprobado: aprobado, estado: String(fila.estado),
              tiempo: Number(fila.tiempo_seg), examen_id: String(fila.examen_id), codigo: String(fila.codigo) };
  if (si_(c.mostrar_correctas)) res.revision = { marcadas: String(fila.respuestas), clave: clave.slice(0, n) };
  if (aprobado && si_(c.mostrar_ranking)) res.ranking = ranking_(c, normCod_(fila.codigo));
  return res;
}

function ranking_(c, yo) {
  var grupo = grupo_(c), apr = {};
  filas_('Config').forEach(function (x) { if (grupo_(x) === grupo) apr[String(x.examen_id)] = Number(x.nota_aprobatoria); });
  return filas_('Respuestas')
    .filter(function (r) { return apr[String(r.examen_id)] !== undefined && Number(r.nota) >= apr[String(r.examen_id)] && r.estado !== 'Fuera de tiempo'; })
    .sort(function (a, b) { return Number(b.nota) - Number(a.nota) || Number(a.tiempo_seg) - Number(b.tiempo_seg); })
    .map(function (r) { return { nombre: String(r.nombres).split(' ')[0] + ' ' + String(r.apellidos).charAt(0) + '.', nota: Number(r.nota),
                                 tiempo: Number(r.tiempo_seg), yo: normCod_(r.codigo) === yo }; });
}

function getRanking(token) {
  var it = intento_(token);
  var r = filas_('Respuestas').filter(function (x) { return String(x.examen_id) === String(it.examen_id) && normCod_(x.codigo) === normCod_(it.codigo); })[0];
  if (!r) throw new Error('Aún no envió su ficha.');
  return resultado_(r).ranking || [];
}

/* ================================================================== panel de administración (?admin=1) */
function claveAdmin_() {
  var p = PropertiesService.getScriptProperties();
  return String(p.getProperty('CLAVE_ADMIN') || p.getProperty('CODIGO_DOCENTE') || CODIGO_DOCENTE_INICIAL || '\u0000');
}
function adminOk_(t) {
  if (!t || CacheService.getScriptCache().get('ADM_' + t) !== '1') throw new Error('Sesión vencida: vuelva a ingresar.');
  CacheService.getScriptCache().put('ADM_' + t, '1', 21600);
}
function enmascarar_(k) { return k ? String(k).slice(0, 7) + '…' + String(k).slice(-4) : ''; }

function adminLogin(clave) {
  var cache = CacheService.getScriptCache(), f = Number(cache.get('ADM_FALLOS') || 0);
  if (f >= 5) throw new Error('Demasiados intentos. Espere 10 minutos.');
  if (String(clave || '').trim() !== claveAdmin_()) { cache.put('ADM_FALLOS', String(f + 1), 600); return { ok: false, error: 'Clave incorrecta.' }; }
  var t = Utilities.getUuid(); cache.put('ADM_' + t, '1', 21600); cache.remove('ADM_FALLOS');
  var p = PropertiesService.getScriptProperties();
  return { ok: true, token: t, claveInicial: !p.getProperty('CLAVE_ADMIN') };
}

function adminEstado(t) {
  adminOk_(t);
  var p = PropertiesService.getScriptProperties(), cache = CacheService.getScriptCache();
  var intentos = filas_('Intentos'), resp = filas_('Respuestas');
  var notas = {}; resp.forEach(function (r) { notas[String(r.examen_id) + '|' + normCod_(r.codigo)] = r; });
  var lista = intentos.map(function (it) {
    var r = notas[String(it.examen_id) + '|' + normCod_(it.codigo)];
    var marcadas = String(it.respuestas || '').replace(/-/g, '').length;
    return { token: String(it.token), examen: String(it.examen_id).slice(-1), codigo: String(it.codigo),
             nombre: String(it.nombres) + ' ' + String(it.apellidos), estado: String(it.estado), fugas: Number(it.infracciones || 0),
             desbloqueos: Number(it.desbloqueos || 0), marcadas: marcadas, nota: r ? Number(r.nota) : null,
             verificacion: String(it.verificacion), min: Math.round((Date.now() - Number(it.inicio_ms)) / 60000) };
  }).reverse();
  var cuenta = function (f) { return lista.filter(f).length; };
  var url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  return {
    ia: { proveedor: String(p.getProperty('IA_PROVEEDOR') || 'auto'),
          claude: enmascarar_(p.getProperty('IA_CLAVE_CLAUDE')), gemini: enmascarar_(p.getProperty('IA_CLAVE_GEMINI')),
          modeloClaude: cache.get('IA_MODELO_claude') || '', modeloGemini: cache.get('IA_MODELO_gemini') || '' },
    rostro: rostroRequerido_(), claveAdminInicial: !p.getProperty('CLAVE_ADMIN'),
    examenes: filas_('Config').map(function (c) { return { id: String(c.examen_id), titulo: String(c.titulo), activo: si_(c.activo),
                                                         minutos: Number(c.minutos), maxInf: Number(c.max_infracciones) }; }),
    resumen: { alumnos: filas_('Alumnos').length, enCurso: cuenta(function (x) { return x.estado === 'En curso'; }),
               bloqueados: cuenta(function (x) { return x.estado === 'Bloqueado'; }), enviados: resp.length },
    intentos: lista.slice(0, 300),
    perfil: perfil_(),
    proyector: (function () {
      var activos = intentos.filter(function (it) { return String(it.estado) === 'En curso' || String(it.estado) === 'Bloqueado'; });
      var fin = 0;
      activos.forEach(function (it) { var c = filas_('Config').filter(function (x) { return String(x.examen_id) === String(it.examen_id); })[0];
        if (c) fin = Math.max(fin, Number(it.inicio_ms) + Number(c.minutos) * 60000); });
      return { finMs: fin, ahoraMs: Date.now() };
    })(),
    enlaces: { alumnos: url, hoja: SpreadsheetApp.getActiveSpreadsheet().getUrl() }
  };
}

function adminGuardarClave(t, clave) { adminOk_(t); return configurarIA(clave); }
function adminBorrarClave(t, prov) {
  adminOk_(t);
  PropertiesService.getScriptProperties().deleteProperty(prov === 'gemini' ? 'IA_CLAVE_GEMINI' : 'IA_CLAVE_CLAUDE');
  CacheService.getScriptCache().remove('IA_MODELO_' + (prov === 'gemini' ? 'gemini' : 'claude'));
  return 'Clave de ' + (prov === 'gemini' ? 'Gemini' : 'Claude') + ' eliminada.';
}
function adminProbarIA(t) { adminOk_(t); return probarIA(); }
function adminAjuste(t, k, v) {
  adminOk_(t);
  var p = PropertiesService.getScriptProperties();
  if (k === 'proveedor') { if (['auto', 'claude', 'gemini'].indexOf(v) < 0) throw new Error('Proveedor no válido.'); p.setProperty('IA_PROVEEDOR', v); return 'Proveedor: ' + v; }
  if (k === 'rostro') { p.setProperty('VERIFICAR_ROSTRO', v ? 'SI' : 'NO'); return 'Comparación de rostros ' + (v ? 'activada' : 'desactivada') + '.'; }
  if (k === 'docente') { if (!/^\d{4,8}$/.test(String(v))) throw new Error('El código docente debe tener de 4 a 8 dígitos.'); p.setProperty('CODIGO_DOCENTE', String(v)); return 'Código docente actualizado.'; }
  if (k === 'admin') { if (String(v).length < 6) throw new Error('La clave de administración debe tener al menos 6 caracteres.'); p.setProperty('CLAVE_ADMIN', String(v)); return 'Clave de administración actualizada.'; }
  throw new Error('Ajuste desconocido.');
}
function adminActivo(t, id, activo) {
  adminOk_(t);
  return conLock_(function () {
    var c = config_(id); actualizar_('Config', c._fila, { activo: activo ? 'SI' : 'NO' });
    return c.examen_id + (activo ? ' abierto.' : ' cerrado.');
  });
}
function adminDesbloquear(t, tokenIntento) {
  adminOk_(t);
  return conLock_(function () {
    var it = intento_(tokenIntento), n = Number(it.desbloqueos || 0) + 1;
    if (String(it.estado) !== 'Bloqueado') return 'Ese intento no está bloqueado.';
    actualizar_('Intentos', it._fila, { estado: 'En curso', desbloqueos: n });
    evento_(it.examen_id, it.codigo, 'desbloqueo docente', 'desde el panel · ' + n);
    return 'Desbloqueado: ' + it.nombres + ' ' + it.apellidos + '.';
  });
}

/** Consulta ligera del estado del intento (la hoja bloqueada pregunta si el docente la desbloqueó desde el panel). */
function estadoIntento(token) {
  var it = filas_('Intentos').filter(function (r) { return String(r.token) === String(token); })[0];
  if (!it) return { existe: false };
  return { existe: true, bloqueado: String(it.estado) === 'Bloqueado', estado: String(it.estado), infracciones: Number(it.infracciones || 0) };
}

/* ================================================================== v4: perfil, reinicios, carga de alumnos, proyector */
var PERFIL_INICIAL = { universidad: 'Mi universidad', facultad: 'Mi facultad', curso: 'Mi curso', docente: 'Nombre del docente', semestre: 'Semestre' };
function perfil_() {
  var s = PropertiesService.getScriptProperties().getProperty('PERFIL');
  var p = {}; try { p = s ? JSON.parse(s) : {}; } catch (e) {}
  Object.keys(PERFIL_INICIAL).forEach(function (k) { if (!p[k]) p[k] = PERFIL_INICIAL[k]; });
  return p;
}
/** Datos públicos de cabecera (sin secretos). */
function getPerfil() { return perfil_(); }

function adminGuardarPerfil(t, p) {
  adminOk_(t); p = p || {};
  var out = {}; Object.keys(PERFIL_INICIAL).forEach(function (k) { out[k] = limpiar_(p[k], 120) || PERFIL_INICIAL[k]; });
  PropertiesService.getScriptProperties().setProperty('PERFIL', JSON.stringify(out));
  return 'Perfil guardado.';
}

/** Reinicia a un alumno: «cero» borra su intento y su nota; «reabrir» anula el envío y le devuelve la hoja con sus marcas y su tiempo. */
function adminResetAlumno(t, tokenIntento, modo) {
  adminOk_(t);
  return conLock_(function () {
    var it = intento_(tokenIntento);
    var resp = filas_('Respuestas').filter(function (r) { return String(r.examen_id) === String(it.examen_id) && normCod_(r.codigo) === normCod_(it.codigo); });
    var quien = it.nombres + ' ' + it.apellidos;
    if (modo === 'reabrir') {
      if (!resp.length) return quien + ' no ha enviado: no hay nada que reabrir.';
      archivar_('reabrir', 'Respuestas', resp);
      borrarFilas_('Respuestas', resp.map(function (r) { return r._fila; }));
      actualizar_('Intentos', intento_(tokenIntento)._fila, { estado: 'En curso', ultimo_guardado: new Date() });
      evento_(it.examen_id, it.codigo, 'reapertura docente', 'conserva marcas y tiempo');
      return 'Reabierto: ' + quien + ' vuelve a su hoja con sus marcas (el tiempo sigue contando desde su inicio).';
    }
    archivar_('reinicio', 'Respuestas', resp);
    archivar_('reinicio', 'Intentos', [it]);
    borrarFilas_('Respuestas', resp.map(function (r) { return r._fila; }));
    borrarFilas_('Intentos', [intento_(tokenIntento)._fila]);
    CacheService.getScriptCache().remove('ROSTRO_' + normCod_(it.codigo));
    evento_(it.examen_id, it.codigo, 'reinicio docente', 'intento y nota archivados');
    return 'Reiniciado: ' + quien + ' puede volver a ingresar desde cero (puede elegir otra versión).';
  });
}

/** Reinicia un examen completo (o todos). Todo se copia a la hoja Archivo antes de borrarse. */
function adminResetExamen(t, examenId, confirmacion) {
  adminOk_(t);
  if (String(confirmacion || '').trim().toUpperCase() !== 'REINICIAR') throw new Error('Escriba REINICIAR para confirmar.');
  return conLock_(function () {
    var cae = function (r) { return examenId === 'TODOS' || String(r.examen_id) === String(examenId); };
    var resp = filas_('Respuestas').filter(cae), ints = filas_('Intentos').filter(cae);
    archivar_('reinicio de examen', 'Respuestas', resp); archivar_('reinicio de examen', 'Intentos', ints);
    borrarFilas_('Respuestas', resp.map(function (r) { return r._fila; }));
    borrarFilas_('Intentos', ints.map(function (r) { return r._fila; }));
    ints.forEach(function (r) { CacheService.getScriptCache().remove('ROSTRO_' + normCod_(r.codigo)); });
    evento_(examenId, '', 'reinicio de examen', ints.length + ' intentos y ' + resp.length + ' notas archivados');
    return (examenId === 'TODOS' ? 'Todos los exámenes' : examenId) + ' reiniciado: ' + ints.length + ' intentos y ' + resp.length + ' notas pasaron a la hoja Archivo.';
  });
}

/** Carga la lista de alumnos desde el Excel (filas ya mapeadas en el panel). */
function adminCargarAlumnos(t, filas, reemplazar) {
  adminOk_(t);
  return conLock_(function () {
    var vistos = {}, nuevas = [];
    (filas || []).forEach(function (r) {
      var cod = normCod_(r.codigo), ape = limpiar_(r.apellidos), nom = limpiar_(r.nombres);
      if (!cod || !ape || vistos[cod]) return;
      vistos[cod] = 1;
      nuevas.push({ codigo: cod, dni: normDni_(r.dni), apellidos: ape, nombres: nom, seccion: limpiar_(r.seccion, 20), email: limpiar_(r.email, 120),
                    origen: 'Lista oficial · ' + Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd') });
    });
    if (!nuevas.length) throw new Error('No encontré filas válidas (cada fila necesita código y apellidos).');
    var actuales = filas_('Alumnos');
    if (reemplazar) { archivar_('lista reemplazada', 'Alumnos', actuales); borrarFilas_('Alumnos', actuales.map(function (r) { return r._fila; })); actuales = []; }
    var ya = {}; actuales.forEach(function (a) { ya[normCod_(a.codigo)] = a; });
    var agregados = 0, actualizados = 0;
    nuevas.forEach(function (n) {
      if (ya[n.codigo]) { actualizar_('Alumnos', ya[n.codigo]._fila, { apellidos: n.apellidos, nombres: n.nombres, seccion: n.seccion || ya[n.codigo].seccion, email: n.email || ya[n.codigo].email, dni: n.dni || ya[n.codigo].dni }); actualizados++; }
      else { agregar_('Alumnos', n); agregados++; }
    });
    evento_('', '', 'carga de alumnos', agregados + ' nuevos, ' + actualizados + ' actualizados' + (reemplazar ? ' (lista reemplazada)' : ''));
    return 'Lista cargada: ' + agregados + ' nuevos' + (actualizados ? ', ' + actualizados + ' actualizados' : '') + '. Total: ' + filas_('Alumnos').length + '.';
  });
}
