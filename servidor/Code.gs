/**
 * Servidor de la app «Actas de Bienes Públicos · CIIP».
 * Guarda actas, bitácora, usuarios y configuración en una hoja de Google Sheets.
 * Instalación: ver servidor/INSTALAR.md. Ejecuta primero la función instalar().
 * Los demás usuarios y los datos de la responsable se cargan desde la app (Configuración y Restaurar respaldo).
 */
const VERSION = "3.2.0";
const APP_ID = "ciip-actas-bienes";
const SESSION_DAYS = 30;
const HASH_ROUNDS = 1500;

const SHEETS = {
  actas: ["id", "correlativo", "tipo", "fecha", "dependencia", "resumen", "bienes", "estado", "revision", "emitida", "emitida_por", "actualizada", "json"],
  bitacora: ["id", "fecha", "accion", "correlativo", "motivo", "detalle", "responsable", "json"],
  usuarios: ["usuario", "nombre", "rol", "activo", "debe_cambiar", "ultimo_acceso", "creado", "creado_por", "salt", "hash"],
  config: ["clave", "json"]
};
const ROLES = { admin: "Administrador", operador: "Operador", consulta: "Consulta" };
const PERMS = {
  emitir: ["admin", "operador"], corregir: ["admin", "operador"], anular: ["admin", "operador"], importar: ["admin", "operador"], registrar: ["admin", "operador"], log: ["admin", "operador"],
  numero: ["admin"], eliminar: ["admin"], revertir: ["admin"], restaurar: ["admin"], config: ["admin"], usuarios: ["admin"]
};
const TIPOS = { ASG: "asig", DES: "desin", MOV: "mov", SAL: "sal", INV: "inv" };

const DEFAULT_CONFIG = {
  inst: "Centro Internacional de Inversión Productiva", siglas: "CIIP", ciudad: "Caracas", entidad: "Distrito Capital",
  prefijo: "CIIP-GBP", lugarSalida: "las instalaciones del Centro Internacional de Inversión Productiva (CIIP)", adminNombre: "CEMK85",
  resp: { trat: "f", nombre: "RAHELENYS JOSUÉ BURGOS SANTAELLA", ci: "", cargo: "Gerente de Bienes Públicos (E)", desig: "mediante Providencia Administrativa N.º 014-2026" },
  aut: [
    { dep: "PRESIDENCIA", n: "ALEJANDRO ENRIQUE PUGLIA COSTAS", c: "Presidente" },
    { dep: "VICEPRESIDENCIA", n: "", c: "" },
    { dep: "AUDITORÍA INTERNA", n: "RAYDA YSABEL VERA PARRA", c: "Auditor Interno" },
    { dep: "CONSULTORÍA JURÍDICA", n: "MARCO ANTONIO MAGALLANES GRILLET", c: "Consultor Jurídico" },
    { dep: "GERENCIA GENERAL DE PLANIFICACIÓN Y PRESUPUESTO", n: "KATIUSKA DEL VALLE DIAZ YANEZ", c: "Gerente General" },
    { dep: "GERENCIA GENERAL DE TECNOLOGÍA DE LA INFORMACIÓN Y COMUNICACIÓN", n: "JESUS ANTONIO ARELLANO NATERA", c: "Gerente General" },
    { dep: "GERENCIA GENERAL DE GESTIÓN HUMANA", n: "KRISTINA DUGNAS SULCAS", c: "Gerente General" },
    { dep: "GERENCIA GENERAL DE GESTIÓN ADMINISTRATIVA", n: "RAHELENYS JOSUÉ BURGOS SANTAELLA", c: "Gerente General" },
    { dep: "GERENCIA GENERAL DE SEGURIDAD INTEGRAL", n: "JOSE MANUEL BARRIOS RIERA", c: "Gerente General" },
    { dep: "GERENCIA GENERAL DEL OBSERVATORIO VENEZOLANO ANTIBLOQUEO", n: "WILLIAM ALFREDO CASTILLO BOLLE", c: "Gerente General" },
    { dep: "GERENCIA GENERAL DE PROMOCIÓN DE INVERSIONES", n: "ENRY JOSUE DURAN", c: "Gerente General" }
  ],
  ultimos: {}
};
const USUARIOS_INICIALES = [
  { usuario: "cemk85", nombre: "CEMK85", rol: "admin" }
];

/* ───────── Instalación (ejecutar una vez desde el editor) ───────── */
function instalar() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SHEET_ID"), ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create("BD · Actas de Bienes Públicos CIIP");
    id = ss.getId(); props.setProperty("SHEET_ID", id);
  }
  const first = ss.getSheets()[0];
  Object.keys(SHEETS).forEach(function (name, i) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = (i === 0 && first.getName() !== name && first.getLastRow() === 0) ? first.setName(name) : ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight("bold"); }
  });
  if (!readConfig_(ss, true)) writeConfig_(ss, DEFAULT_CONFIG);
  const creds = [];
  USUARIOS_INICIALES.forEach(function (u) {
    if (findUser_(ss, u.usuario)) return;
    const pass = tempPass_(); const salt = Utilities.getUuid().replace(/-/g, "");
    upsertUser_(ss, { usuario: u.usuario, nombre: u.nombre, rol: u.rol, activo: true, debeCambiar: true, ultimoAcceso: "", creado: new Date().toISOString(), creadoPor: "instalación", salt: salt, hash: hashPass_(pass, salt) });
    creds.push(u.nombre + " · usuario: " + u.usuario + " · clave temporal: " + pass);
  });
  bump_();
  Logger.log("Base de datos: " + ss.getUrl());
  if (creds.length) Logger.log("CLAVES TEMPORALES (cópialas ahora; no se vuelven a mostrar):\n" + creds.join("\n"));
  else Logger.log("Los usuarios iniciales ya existían; no se cambiaron sus claves.");
  return ss.getUrl();
}

/** Genera una clave temporal nueva para un usuario (ejecutar desde el editor si nadie puede entrar). */
function restablecerClave(usuario) {
  const ss = book_(); const u = findUser_(ss, normUser_(usuario || "cemk85"));
  if (!u) throw new Error("No existe el usuario " + usuario);
  const pass = tempPass_(); u.salt = Utilities.getUuid().replace(/-/g, ""); u.hash = hashPass_(pass, u.salt); u.debeCambiar = true; u.activo = true;
  upsertUser_(ss, u); bump_();
  Logger.log(u.nombre + " · usuario: " + u.usuario + " · clave temporal: " + pass);
}

/* ───────── Web app ───────── */
function doGet() { return out_({ ok: true, app: APP_ID, version: VERSION }); }
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: "Solicitud no válida." }); }
  try { const r = route_(req) || {}; r.ok = true; r.version = VERSION; return out_(r); }
  catch (err) { return out_({ ok: false, error: String((err && err.message) || err), code: (err && err.code) || "" }); }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fail_(msg, code) { const e = new Error(msg); e.code = code || "error"; throw e; }

function route_(req) {
  const a = req.action;
  if (a === "ping") return { ping: true };
  if (a === "login") return login_(req);
  const ss = book_(); const me = session_(ss, req.token);
  switch (a) {
    case "yo": return { me: publicUser_(me) };
    case "logout": PropertiesService.getScriptProperties().deleteProperty("S_" + req.token); return {};
    case "sync": return sync_(ss, me, req.ver);
    case "clave": return changePass_(ss, me, req);
    case "emitir": need_(me, "emitir"); return withLock_(function () { return { rec: emitir_(ss, me, req.rec) }; });
    case "guardar": need_(me, req.accion === "registrar" ? "registrar" : req.accion === "restaurar" ? "restaurar" : req.accion === "importar" ? "importar" : "corregir"); return withLock_(function () { return { rec: guardar_(ss, me, req.rec, !!req.nuevo) }; });
    case "renumerar": need_(me, "numero"); return withLock_(function () { return { rec: renumerar_(ss, req.id, Number(req.num), Number(req.anio)) }; });
    case "eliminar": need_(me, "eliminar"); return withLock_(function () { deleteRow_(ss, "actas", req.id); bump_(); return {}; });
    case "actualizar": need_(me, req.data && req.data.estado === "emitida" ? "revertir" : "anular"); return withLock_(function () { return { rec: patch_(ss, req.id, req.data || {}) }; });
    case "importar": need_(me, "importar"); return withLock_(function () { return importar_(ss, req); });
    case "log": need_(me, "log"); return withLock_(function () { return { entry: log_(ss, me, req.entry || {}) }; });
    case "logPatch": need_(me, "restaurar"); return withLock_(function () { const all = rows_(ss, "bitacora"); const r = all.find(function (x) { return x.obj.id === req.id; }); if (r) { Object.assign(r.obj, req.data || {}); writeLogRow_(ss, r.obj, r.row); bump_(); } return {}; });
    case "config": need_(me, "config"); return withLock_(function () { writeConfig_(ss, req.config); bump_(); return {}; });
    case "usuario": need_(me, "usuarios"); return withLock_(function () { return userOp_(ss, me, req); });
  }
  fail_("Acción desconocida: " + a);
}
function need_(me, perm) { if ((PERMS[perm] || ["admin"]).indexOf(me.rol) < 0) fail_("Tu rol (" + (ROLES[me.rol] || me.rol) + ") no permite esta acción.", "forbidden"); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
function bump_() { const p = PropertiesService.getScriptProperties(); p.setProperty("VER", String(Number(p.getProperty("VER") || 0) + 1)); }
function book_() { const id = PropertiesService.getScriptProperties().getProperty("SHEET_ID"); if (!id) fail_("El servidor no está instalado. Ejecuta instalar() en el editor de Apps Script.", "setup"); return SpreadsheetApp.openById(id); }

/* ───────── Sesiones ───────── */
function login_(req) {
  const ss = book_(); const u = normUser_(req.usuario); const cache = CacheService.getScriptCache();
  const fk = "fail_" + u; const fails = Number(cache.get(fk) || 0);
  if (fails >= 6) fail_("Demasiados intentos fallidos. Espera 15 minutos o pide a un administrador que restablezca tu clave.", "locked");
  const rec = findUser_(ss, u);
  if (!rec || rec.activo === false || hashPass_(String(req.clave || ""), rec.salt) !== rec.hash) {
    cache.put(fk, String(fails + 1), 900);
    fail_(rec && rec.activo === false ? "Este usuario está desactivado." : "Usuario o clave incorrectos.", "auth");
  }
  cache.remove(fk);
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, "");
  const p = PropertiesService.getScriptProperties();
  p.setProperty("S_" + token, JSON.stringify({ u: rec.usuario, h: rec.hash.slice(0, 16), exp: Date.now() + SESSION_DAYS * 864e5 }));
  cleanSessions_(p);
  rec.ultimoAcceso = new Date().toISOString(); upsertUser_(ss, rec);
  return { token: token, me: publicUser_(rec) };
}
function session_(ss, token) {
  if (!token) fail_("Inicia sesión para continuar.", "auth");
  const p = PropertiesService.getScriptProperties(); const raw = p.getProperty("S_" + token);
  if (!raw) fail_("Tu sesión expiró. Vuelve a iniciar sesión.", "auth");
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { p.deleteProperty("S_" + token); fail_("Tu sesión expiró. Vuelve a iniciar sesión.", "auth"); }
  const u = findUser_(ss, s.u);
  if (!u || u.activo === false || u.hash.slice(0, 16) !== s.h) { p.deleteProperty("S_" + token); fail_("Tu acceso cambió. Vuelve a iniciar sesión.", "auth"); }
  return u;
}
function cleanSessions_(p) { const all = p.getProperties(); Object.keys(all).forEach(function (k) { if (k.indexOf("S_") === 0) { try { if (JSON.parse(all[k]).exp < Date.now()) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); } } }); }
function publicUser_(u) { return { usuario: u.usuario, nombre: u.nombre, rol: u.rol, debeCambiar: !!u.debeCambiar }; }
function changePass_(ss, me, req) {
  const nueva = String(req.nueva || "");
  if (!me.debeCambiar && hashPass_(String(req.actual || ""), me.salt) !== me.hash) fail_("La clave actual no es correcta.");
  if (nueva.length < 8) fail_("La clave debe tener al menos 8 caracteres.");
  return withLock_(function () {
    me.salt = Utilities.getUuid().replace(/-/g, ""); me.hash = hashPass_(nueva, me.salt); me.debeCambiar = false;
    upsertUser_(ss, me); bump_();
    const p = PropertiesService.getScriptProperties(); const all = p.getProperties();
    Object.keys(all).forEach(function (k) { if (k.indexOf("S_") === 0) { try { if (JSON.parse(all[k]).u === me.usuario) p.deleteProperty(k); } catch (e) {} } });
    const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, "");
    p.setProperty("S_" + token, JSON.stringify({ u: me.usuario, h: me.hash.slice(0, 16), exp: Date.now() + SESSION_DAYS * 864e5 }));
    return { token: token, me: publicUser_(me) };
  });
}

/* ───────── Datos ───────── */
function sync_(ss, me, ver) {
  const cur = PropertiesService.getScriptProperties().getProperty("VER") || "0";
  const base = { ver: cur, me: publicUser_(me) };
  if (String(ver) === cur) { base.changed = false; return base; }
  base.changed = true;
  base.actas = rows_(ss, "actas").map(function (r) { return r.obj; });
  base.bitacora = rows_(ss, "bitacora").map(function (r) { return r.obj; });
  base.config = readConfig_(ss) || DEFAULT_CONFIG;
  base.usuarios = me.rol === "admin" ? rows_(ss, "usuarios").map(function (r) { const o = r.obj; return { usuario: o.usuario, nombre: o.nombre, rol: o.rol, activo: o.activo, debeCambiar: o.debeCambiar, ultimoAcceso: o.ultimoAcceso, creado: o.creado }; }) : [];
  return base;
}
function pad4_(n) { return ("0000" + n).slice(-Math.max(4, String(n).length)); }
function correlativo_(cfg, code, anio, num) { return (cfg.prefijo || "CIIP") + "-" + code + "-" + anio + "-" + pad4_(num); }
function emitir_(ss, me, rec) {
  if (!rec || !TIPOS[rec.code] || !rec.anio) fail_("Datos del acta incompletos.");
  const cfg = readConfig_(ss) || DEFAULT_CONFIG; const anio = Number(rec.anio);
  let max = Number((cfg.ultimos || {})[rec.code + "-" + anio] || 0);
  rows_(ss, "actas").forEach(function (r) { if (r.obj.code === rec.code && Number(r.obj.anio) === anio && Number(r.obj.num) > max) max = Number(r.obj.num); });
  rec.num = max + 1; rec.anio = anio; rec.id = rec.code + "-" + anio + "-" + pad4_(rec.num);
  rec.correlativo = correlativo_(rec.cfgSnap || cfg, rec.code, anio, rec.num);
  rec.emitidaEn = new Date().toISOString(); rec.emisor = me.nombre; rec.emisorUsuario = me.usuario; rec.rev = 0; rec.estado = "emitida";
  writeActa_(ss, rec, null); bump_();
  return rec;
}
function guardar_(ss, me, rec, nuevo) {
  if (!rec || !TIPOS[rec.code] || !rec.num || !rec.anio) fail_("Datos del acta incompletos.");
  rec.id = rec.code + "-" + rec.anio + "-" + pad4_(rec.num);
  const ex = rows_(ss, "actas").find(function (r) { return r.obj.id === rec.id; });
  if (nuevo && ex) fail_(rec.correlativo + " ya existe en el historial.", "exists");
  if (!nuevo && !ex) fail_("El acta " + rec.id + " ya no está en el historial.", "missing");
  writeActa_(ss, rec, ex ? ex.row : null); bump_();
  return rec;
}
function renumerar_(ss, id, num, anio) {
  if (!(num > 0) || !(anio > 2000)) fail_("Número o año no válido.");
  const all = rows_(ss, "actas"); const ex = all.find(function (r) { return r.obj.id === id; });
  if (!ex) fail_("El acta ya no está en el historial.", "missing");
  const nid = ex.obj.code + "-" + anio + "-" + pad4_(num);
  if (all.some(function (r) { return r.obj.id === nid; })) fail_("Ese número ya está ocupado.", "exists");
  const cfg = readConfig_(ss) || DEFAULT_CONFIG; const nr = ex.obj;
  nr.numerosAnteriores = (nr.numerosAnteriores || []).concat(nr.correlativo);
  nr.num = num; nr.anio = anio; nr.id = nid; nr.correlativo = correlativo_(nr.cfgSnap || cfg, nr.code, anio, num);
  writeActa_(ss, nr, ex.row); bump_();
  return nr;
}
function patch_(ss, id, data) {
  const ex = rows_(ss, "actas").find(function (r) { return r.obj.id === id; });
  if (!ex) fail_("El acta ya no está en el historial.", "missing");
  Object.assign(ex.obj, data); writeActa_(ss, ex.obj, ex.row); bump_();
  return ex.obj;
}
function importar_(ss, req) {
  const have = {}; rows_(ss, "actas").forEach(function (r) { have[r.obj.id] = 1; });
  let n = 0, m = 0;
  (req.actas || []).forEach(function (rec) {
    if (!rec || !TIPOS[rec.code] || !rec.num || !rec.anio) return;
    rec.id = rec.code + "-" + rec.anio + "-" + pad4_(rec.num);
    if (have[rec.id]) return; writeActa_(ss, rec, null); have[rec.id] = 1; n++;
  });
  const logs = {}; rows_(ss, "bitacora").forEach(function (r) { logs[r.obj.id] = 1; });
  (req.bitacora || []).forEach(function (e) { if (!e || !e.id || logs[e.id]) return; writeLogRow_(ss, e, null); logs[e.id] = 1; m++; });
  if (n || m) bump_();
  return { actas: n, bitacora: m };
}
function log_(ss, me, e) {
  e.id = e.id || (Date.now().toString(36) + "-" + Utilities.getUuid().slice(0, 5));
  e.en = new Date().toISOString(); e.por = me.nombre; e.porUsuario = me.usuario;
  writeLogRow_(ss, e, null); bump_();
  return e;
}
function userOp_(ss, me, req) {
  const op = req.op; const u = normUser_(req.usuario);
  if (op === "crear") {
    if (u.length < 3) fail_("El usuario debe tener al menos 3 caracteres.");
    if (findUser_(ss, u)) fail_("Ese usuario ya existe.", "exists");
    if (!ROLES[req.rol]) fail_("Rol no válido.");
    const pass = tempPass_(); const salt = Utilities.getUuid().replace(/-/g, "");
    const rec = { usuario: u, nombre: String(req.nombre || "").toUpperCase().trim(), rol: req.rol, activo: true, debeCambiar: true, ultimoAcceso: "", creado: new Date().toISOString(), creadoPor: me.nombre, salt: salt, hash: hashPass_(pass, salt) };
    if (!rec.nombre) fail_("Escribe el nombre.");
    upsertUser_(ss, rec); log_(ss, me, { accion: "Usuario creado", correlativo: "", motivo: rec.nombre + " (" + u + ")", detalle: "Rol: " + ROLES[rec.rol] });
    return { usuario: publicUser_(rec), clave: pass };
  }
  const rec = findUser_(ss, u); if (!rec) fail_("No existe el usuario.");
  if (op === "reset") {
    const pass = tempPass_(); rec.salt = Utilities.getUuid().replace(/-/g, ""); rec.hash = hashPass_(pass, rec.salt); rec.debeCambiar = true; rec.activo = true;
    upsertUser_(ss, rec); log_(ss, me, { accion: "Clave restablecida", correlativo: "", motivo: rec.nombre + " (" + u + ")" });
    return { usuario: publicUser_(rec), clave: pass };
  }
  if (op === "activo") {
    if (u === me.usuario) fail_("No puedes desactivar tu propio usuario.");
    rec.activo = !!req.activo; upsertUser_(ss, rec); log_(ss, me, { accion: rec.activo ? "Usuario activado" : "Usuario desactivado", correlativo: "", motivo: rec.nombre + " (" + u + ")" });
    return {};
  }
  if (op === "rol") {
    if (!ROLES[req.rol]) fail_("Rol no válido.");
    if (u === me.usuario && req.rol !== "admin") fail_("No puedes quitarte el rol de administrador.");
    const ant = rec.rol; rec.rol = req.rol; upsertUser_(ss, rec); log_(ss, me, { accion: "Rol cambiado", correlativo: "", motivo: rec.nombre + " (" + u + ")", detalle: ROLES[ant] + " → " + ROLES[rec.rol] });
    return {};
  }
  fail_("Operación de usuario desconocida.");
}

/* ───────── Hojas ───────── */
function rows_(ss, name) {
  const sh = ss.getSheetByName(name); const vals = sh.getDataRange().getValues(); const H = SHEETS[name]; const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i]; if (!r[0]) continue; let obj;
    if (name === "usuarios") obj = { usuario: String(r[0]), nombre: String(r[1]), rol: String(r[2]), activo: r[3] !== false && r[3] !== "FALSE", debeCambiar: r[4] === true || r[4] === "TRUE", ultimoAcceso: String(r[5] || ""), creado: String(r[6] || ""), creadoPor: String(r[7] || ""), salt: String(r[8]), hash: String(r[9]) };
    else { try { obj = JSON.parse(r[H.length - 1]); } catch (e) { continue; } }
    out.push({ row: i + 1, obj: obj });
  }
  return out;
}
const safe_ = v => { v = v == null ? "" : String(v); return /^[=+\-@]/.test(v) ? "'" + v : v; };
function putRow_(ss, name, arr, row) {
  const sh = ss.getSheetByName(name);
  if (row) sh.getRange(row, 1, 1, arr.length).setValues([arr]); else sh.appendRow(arr);
}
function trimRec_(rec) {
  let s = JSON.stringify(rec);
  const revs = rec.revisiones || [];
  for (let i = 0; s.length > 45000 && i < revs.length; i++) { delete revs[i].anterior; revs[i].anteriorRecortada = true; s = JSON.stringify(rec); }
  if (s.length > 49000) fail_("El acta es demasiado grande para guardarse.");
  return s;
}
function writeActa_(ss, rec, row) {
  rec.actualizadaEn = new Date().toISOString();
  const json = trimRec_(rec);
  putRow_(ss, "actas", [rec.id, rec.correlativo, safe_(rec.code), rec.fecha || "", safe_(rec.dep), safe_(rec.resumen), rec.nBienes || 0, rec.estado || "emitida", rec.rev || 0, rec.emitidaEn || "", safe_(rec.emisor), rec.actualizadaEn, json], row);
}
function writeLogRow_(ss, e, row) { putRow_(ss, "bitacora", [e.id, e.en || "", safe_(e.accion), safe_(e.correlativo), safe_(e.motivo), safe_(e.detalle), safe_(e.por), JSON.stringify(e)], row); }
function deleteRow_(ss, name, id) { const ex = rows_(ss, name).find(function (r) { return r.obj.id === id; }); if (!ex) fail_("El acta ya no está en el historial.", "missing"); ss.getSheetByName(name).deleteRow(ex.row); }
function findUser_(ss, u) { const r = rows_(ss, "usuarios").find(function (x) { return x.obj.usuario === u; }); return r ? Object.assign(r.obj, { _row: r.row }) : null; }
function upsertUser_(ss, u) {
  const ex = u._row ? { row: u._row } : rows_(ss, "usuarios").find(function (x) { return x.obj.usuario === u.usuario; });
  putRow_(ss, "usuarios", [u.usuario, safe_(u.nombre), u.rol, u.activo !== false, !!u.debeCambiar, u.ultimoAcceso || "", u.creado || "", safe_(u.creadoPor), u.salt, u.hash], ex ? ex.row : null);
}
function readConfig_(ss, quiet) { const r = ss.getSheetByName("config").getDataRange().getValues(); for (let i = 1; i < r.length; i++) if (r[i][0] === "general") { try { return JSON.parse(r[i][1]); } catch (e) { return null; } } return null; }
function writeConfig_(ss, cfg) {
  const sh = ss.getSheetByName("config"); const r = sh.getDataRange().getValues();
  for (let i = 1; i < r.length; i++) if (r[i][0] === "general") { sh.getRange(i + 1, 2).setValue(JSON.stringify(cfg)); return; }
  sh.appendRow(["general", JSON.stringify(cfg)]);
}

/* ───────── Claves ───────── */
function normUser_(s) { return String(s || "").trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9._-]/g, ""); }
function hex_(bytes) { return bytes.map(function (b) { return ((b < 0 ? b + 256 : b).toString(16)).padStart(2, "0"); }).join(""); }
function hashPass_(pass, salt) {
  let h = salt + ":" + pass;
  for (let i = 0; i < HASH_ROUNDS; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return h;
}
function tempPass_() { const A = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789"; const id = Utilities.getUuid().replace(/-/g, ""); let s = ""; for (let i = 0; i < 10; i++) s += A[parseInt(id.substr(i * 2, 2), 16) % A.length]; return "Ciip-" + s.slice(0, 5) + "-" + s.slice(5); }
