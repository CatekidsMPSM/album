// ============================================================
// ÁLBUM CATEKIDS — Google Apps Script Backend
// Versión 3.1 | Para usar con GitHub Pages + Apps Script
//
// INSTRUCCIONES RÁPIDAS:
// 1. Pega este código en tu proyecto de Apps Script
// 2. Configura DRIVE_FOLDER_ID con el ID de tu carpeta de Drive
// 3. Asegúrate de tener las hojas: "DATOS" y "Respuestas de formulario 1"
// 4. Implementar → Nueva implementación → Web App
//    - Ejecutar como: Yo
//    - Acceso: Cualquier persona
// 5. Copia la URL /exec y pégala en index.html en WEBAPP_URL
// ============================================================

// ── CONFIGURACIÓN GLOBAL ──────────────────────────────────
const CONFIG = {
  SHEET_RESPUESTAS: "Respuestas de formulario 1",
  SHEET_DATOS:      "DATOS",
  TOTAL_TARJETAS:   365,
  CACHE_SEGUNDOS:   300,
  // 🔧 REEMPLAZA con el ID de la carpeta de Drive donde se guardarán las fotos
  // (crea una carpeta en Drive, ábrela, y el ID está en la URL: drive.google.com/drive/folders/ESTE_ID)
  DRIVE_FOLDER_ID:  "TU_ID_DE_CARPETA_DRIVE",
};

// Columnas en "Respuestas de formulario 1" (base 1)
const COL = {
  TIMESTAMP: 1,
  EMAIL:     2,
  CODIGO:    3,
  NOMBRE:    4,
  TARJETA:   5,
  IMAGEN:    6,
};

// ── HEADERS CORS ─────────────────────────────────────────
function corsHeaders() {
  return {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

// ── PUNTO DE ENTRADA GET ──────────────────────────────────
function doGet(e) {
  const params = e.parameter;

  if (params.action === "getAlbum" && params.codigo) {
    return jsonResponse(getAlbum(params.codigo));
  }
  if (params.action === "validateCodigo" && params.codigo) {
    return jsonResponse(validateCodigo(params.codigo));
  }
  if (params.action === "getRanking") {
    return jsonResponse(getRanking());
  }
  if (params.action === "getGrupoStats" && params.grupo) {
    return jsonResponse(getGrupoStats(params.grupo));
  }

  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ÁlbumKids 🌟")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ── PUNTO DE ENTRADA POST (subir imagen) ─────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === "subirTarjeta")   return jsonResponse(subirTarjeta(body));
    if (body.action === "subirInsignia")  return jsonResponse(subirInsignia(body));
    if (body.action === "borrarInsignia") return jsonResponse(borrarInsignia(body));

    return jsonResponse({ ok: false, mensaje: "Acción no reconocida" });
  } catch (err) {
    return jsonResponse({ ok: false, mensaje: "Error en doPost: " + err.message });
  }
}

// ── SUBIR INSIGNIA CATEKIDS A DRIVE ─────────────────────
function subirInsignia(body) {
  try {
    const codigo    = String(body.codigo  || "").trim().toUpperCase();
    const insId     = parseInt(body.insigniaId, 10);
    const insNombre = String(body.insigniaNombre || "insignia");
    const dataUrl   = String(body.imagen   || "");
    const mimeType  = String(body.mimeType || "image/jpeg");

    if (!codigo || isNaN(insId))              return { ok: false, mensaje: "Datos incompletos" };
    if (!dataUrl.startsWith("data:"))         return { ok: false, mensaje: "Imagen inválida" };

    const perfil = validateCodigo(codigo);
    if (!perfil.ok) return perfil;

    const base64Data = dataUrl.split(",")[1];
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType,
      codigo + "_insignia" + insId + "_" + Date.now() + ".jpg"
    );

    // Carpeta base
    let folder;
    try { folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID); }
    catch (e) { folder = DriveApp.getRootFolder(); }

    // Subcarpeta por niño
    let subFolder;
    const it = folder.getFoldersByName(codigo);
    subFolder = it.hasNext() ? it.next() : folder.createFolder(codigo);

    // Subcarpeta insignias dentro de la del niño
    let insFolder;
    const it2 = subFolder.getFoldersByName("insignias");
    insFolder = it2.hasNext() ? it2.next() : subFolder.createFolder("insignias");

    // Borrar foto anterior si existe (para no acumular archivos)
    const prevIt = insFolder.getFilesByName(codigo + "_insignia" + insId + ".jpg");
    while (prevIt.hasNext()) { prevIt.next().setTrashed(true); }

    const archivo = insFolder.createFile(blob);
    archivo.setName(codigo + "_insignia" + insId + ".jpg");
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const url = "https://drive.google.com/uc?export=view&id=" + archivo.getId();

    // Guardar en hoja INSIGNIAS (créala si no existe)
    guardarInsigniaEnHoja(codigo, perfil.nombre, insId, insNombre, url);

    invalidarCache(codigo);
    return { ok: true, url: url };

  } catch (err) {
    return { ok: false, mensaje: "Error al subir insignia: " + err.message };
  }
}

function guardarInsigniaEnHoja(codigo, nombre, insId, insNombre, url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("INSIGNIAS_CK");
  if (!hoja) {
    hoja = ss.insertSheet("INSIGNIAS_CK");
    hoja.appendRow(["Timestamp","Código","Nombre","InsigniaId","InsigniaNombre","URL"]);
  }
  const datos = hoja.getDataRange().getValues();
  // Actualizar fila existente si ya tiene esta insignia
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim().toUpperCase() === codigo && parseInt(datos[i][3]) === insId) {
      hoja.getRange(i + 1, 1, 1, 6).setValues([[new Date(), codigo, nombre, insId, insNombre, url]]);
      return;
    }
  }
  // Si no existe, agregar nueva fila
  hoja.appendRow([new Date(), codigo, nombre, insId, insNombre, url]);
}

// ── BORRAR INSIGNIA ────────────────────────────────────
function borrarInsignia(body) {
  try {
    const codigo = String(body.codigo || "").trim().toUpperCase();
    const insId  = parseInt(body.insigniaId, 10);
    if (!codigo || isNaN(insId)) return { ok: false, mensaje: "Datos incompletos" };

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("INSIGNIAS_CK");
    if (!hoja) return { ok: true };

    const datos = hoja.getDataRange().getValues();
    for (let i = datos.length - 1; i >= 1; i--) {
      if (String(datos[i][1]).trim().toUpperCase() === codigo && parseInt(datos[i][3]) === insId) {
        hoja.deleteRow(i + 1);
      }
    }
    invalidarCache(codigo);
    return { ok: true };
  } catch (err) {
    return { ok: false, mensaje: err.message };
  }
}

// ── SUBIR TARJETA CON IMAGEN A DRIVE ─────────────────────
function subirTarjeta(body) {
  try {
    const codigo    = String(body.codigo   || "").trim().toUpperCase();
    const nombre    = String(body.nombre   || "").trim();
    const numTarj   = parseInt(body.tarjeta, 10);
    const dataUrl   = String(body.imagen   || "");
    const mimeType  = String(body.mimeType || "image/jpeg");

    // Validaciones básicas
    if (!codigo)             return { ok: false, mensaje: "Código vacío" };
    if (isNaN(numTarj) || numTarj < 1 || numTarj > CONFIG.TOTAL_TARJETAS)
                             return { ok: false, mensaje: "Número de tarjeta inválido" };
    if (!dataUrl.startsWith("data:"))
                             return { ok: false, mensaje: "Imagen inválida" };

    // Verificar que el código existe
    const perfil = validateCodigo(codigo);
    if (!perfil.ok)          return perfil;

    // Decodificar base64
    const base64Data = dataUrl.split(",")[1];
    const blob       = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType,
      codigo + "_tarjeta" + numTarj + "_" + Date.now() + ".jpg"
    );

    // Subir a Drive
    let folder;
    try {
      folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    } catch (err) {
      // Si no existe la carpeta configurada, usar la raíz
      folder = DriveApp.getRootFolder();
    }

    // Subcarpeta por código (opcional pero ordenado)
    let subFolder;
    const subFolderIter = folder.getFoldersByName(codigo);
    if (subFolderIter.hasNext()) {
      subFolder = subFolderIter.next();
    } else {
      subFolder = folder.createFolder(codigo);
    }

    const archivo  = subFolder.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId   = archivo.getId();
    const urlImg   = "https://drive.google.com/uc?export=view&id=" + fileId;

    // Registrar en la hoja de respuestas
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.SHEET_RESPUESTAS);
    hoja.appendRow([
      new Date(),        // TIMESTAMP
      "",                // EMAIL (vacío, no requerido)
      codigo,            // CODIGO
      nombre,            // NOMBRE
      numTarj,           // TARJETA
      urlImg,            // IMAGEN
    ]);

    // Invalidar caché para que se recargue el álbum
    invalidarCache(codigo);

    return { ok: true, url: urlImg, fileId: fileId };

  } catch (err) {
    return { ok: false, mensaje: "Error al subir: " + err.message };
  }
}

// ── VALIDAR CÓDIGO DEL NIÑO ───────────────────────────────
function validateCodigo(codigo) {
  try {
    const codigoLimpio = String(codigo).trim().toUpperCase();
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const hoja  = ss.getSheetByName(CONFIG.SHEET_DATOS);
    const datos = hoja.getDataRange().getValues();

    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toUpperCase() === codigoLimpio) {
        return {
          ok:         true,
          codigo:     codigoLimpio,
          nombre:     datos[i][1],
          grupo:      datos[i][2],
          catequista: datos[i][3],
        };
      }
    }
    return { ok: false, mensaje: "Código no encontrado. Pide ayuda a tu catequista 🙏" };
  } catch (err) {
    return { ok: false, mensaje: "Error al validar: " + err.message };
  }
}

// ── OBTENER ÁLBUM DE UN NIÑO ──────────────────────────────
function getAlbum(codigo) {
  const codigoLimpio = String(codigo).trim().toUpperCase();
  const cacheKey     = "album_" + codigoLimpio;
  const cache        = CacheService.getScriptCache();

  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  try {
    const perfil = validateCodigo(codigoLimpio);
    if (!perfil.ok) return perfil;

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const hoja  = ss.getSheetByName(CONFIG.SHEET_RESPUESTAS);
    const datos = hoja.getDataRange().getValues();

    const tarjetasMap = {};

    for (let i = 1; i < datos.length; i++) {
      const fila    = datos[i];
      const codFila = String(fila[COL.CODIGO - 1]).trim().toUpperCase();
      if (codFila !== codigoLimpio) continue;

      const numTarjeta = parseInt(fila[COL.TARJETA - 1], 10);
      if (isNaN(numTarjeta) || numTarjeta < 1 || numTarjeta > CONFIG.TOTAL_TARJETAS) continue;

      const timestamp = new Date(fila[COL.TIMESTAMP - 1]).getTime();
      const urlImg    = String(fila[COL.IMAGEN - 1]).trim();

      if (!tarjetasMap[numTarjeta] || timestamp > tarjetasMap[numTarjeta].timestamp) {
        tarjetasMap[numTarjeta] = {
          numero:    numTarjeta,
          url:       urlImg ? convertirUrlDrive(urlImg) : null,
          timestamp: timestamp,
        };
      }
    }

    const tarjetas = [];
    for (let n = 1; n <= CONFIG.TOTAL_TARJETAS; n++) {
      tarjetas.push(tarjetasMap[n] || { numero: n, url: null });
    }

    // Cargar fotos de insignias CateKids
    const ckFotos = {};
    const hInsig = ss.getSheetByName("INSIGNIAS_CK");
    if (hInsig) {
      const insData = hInsig.getDataRange().getValues();
      for (let i = 1; i < insData.length; i++) {
        const codIns = String(insData[i][1]).trim().toUpperCase();
        if (codIns !== codigoLimpio) continue;
        const insId  = parseInt(insData[i][3], 10);
        const insUrl = String(insData[i][5]).trim();
        if (!isNaN(insId) && insUrl) {
          ckFotos[insId] = convertirUrlDrive(insUrl);
        }
      }
    }

    const resultado = {
      ok:          true,
      nombre:      perfil.nombre,
      grupo:       perfil.grupo,
      catequista:  perfil.catequista,
      codigo:      codigoLimpio,
      tarjetas:    tarjetas,
      ckFotos:     ckFotos,
      completadas: Object.keys(tarjetasMap).length,
      total:       CONFIG.TOTAL_TARJETAS,
      porcentaje:  Math.round((Object.keys(tarjetasMap).length / CONFIG.TOTAL_TARJETAS) * 100),
    };

    cache.put(cacheKey, JSON.stringify(resultado), CONFIG.CACHE_SEGUNDOS);
    return resultado;

  } catch (err) {
    return { ok: false, mensaje: "Error al cargar el álbum: " + err.message };
  }
}

// ── RANKING TOP 10 ────────────────────────────────────────
function getRanking() {
  const cacheKey = "ranking_global";
  const cache    = CacheService.getScriptCache();
  const cached   = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const hoja  = ss.getSheetByName(CONFIG.SHEET_RESPUESTAS);
    const datos = hoja.getDataRange().getValues();

    const conteo  = {};
    const nombres = {};

    for (let i = 1; i < datos.length; i++) {
      const fila    = datos[i];
      const codigo  = String(fila[COL.CODIGO - 1]).trim().toUpperCase();
      const nombre  = String(fila[COL.NOMBRE - 1]).trim();
      const tarjeta = parseInt(fila[COL.TARJETA - 1], 10);
      if (!codigo || isNaN(tarjeta)) continue;
      if (!conteo[codigo]) conteo[codigo] = new Set();
      conteo[codigo].add(tarjeta);
      nombres[codigo] = nombre;
    }

    const ranking = Object.entries(conteo)
      .map(([codigo, set]) => ({ codigo, nombre: nombres[codigo], tarjetas: set.size }))
      .sort((a, b) => b.tarjetas - a.tarjetas)
      .slice(0, 10);

    const resultado = { ok: true, ranking };
    cache.put(cacheKey, JSON.stringify(resultado), CONFIG.CACHE_SEGUNDOS);
    return resultado;
  } catch (err) {
    return { ok: false, mensaje: err.message };
  }
}

// ── ESTADÍSTICAS POR GRUPO ────────────────────────────────
function getGrupoStats(grupo) {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const hDatos     = ss.getSheetByName(CONFIG.SHEET_DATOS);
    const hResp      = ss.getSheetByName(CONFIG.SHEET_RESPUESTAS);
    const datosNinos = hDatos.getDataRange().getValues();
    const datosResp  = hResp.getDataRange().getValues();

    const ninosGrupo = {};
    for (let i = 1; i < datosNinos.length; i++) {
      if (String(datosNinos[i][2]).trim() === grupo) {
        ninosGrupo[String(datosNinos[i][0]).trim().toUpperCase()] = {
          nombre: datosNinos[i][1], tarjetas: new Set(),
        };
      }
    }
    for (let i = 1; i < datosResp.length; i++) {
      const codigo  = String(datosResp[i][COL.CODIGO - 1]).trim().toUpperCase();
      const tarjeta = parseInt(datosResp[i][COL.TARJETA - 1], 10);
      if (ninosGrupo[codigo] && !isNaN(tarjeta)) {
        ninosGrupo[codigo].tarjetas.add(tarjeta);
      }
    }

    const stats = Object.entries(ninosGrupo)
      .map(([codigo, data]) => ({ codigo, nombre: data.nombre, tarjetas: data.tarjetas.size }))
      .sort((a, b) => b.tarjetas - a.tarjetas);

    return { ok: true, grupo, stats };
  } catch (err) {
    return { ok: false, mensaje: err.message };
  }
}

// ── INVALIDAR CACHÉ ───────────────────────────────────────
function invalidarCache(codigo) {
  const cache = CacheService.getScriptCache();
  cache.remove("album_" + String(codigo).trim().toUpperCase());
  cache.remove("ranking_global");
}

// ── TRIGGER onFormSubmit ──────────────────────────────────
function onFormSubmit(e) {
  try {
    const respuestas = e.values;
    const codigo     = String(respuestas[COL.CODIGO - 1]).trim().toUpperCase();
    invalidarCache(codigo);
  } catch (err) {
    console.error("Error en trigger onFormSubmit:", err);
  }
}

// ── UTIL: Convertir URL de Drive ──────────────────────────
function convertirUrlDrive(url) {
  if (!url) return null;
  let match = url.match(/[?&]id=([^&]+)/);
  if (match) return "https://drive.google.com/uc?export=view&id=" + match[1];
  match = url.match(/\/d\/([^\/]+)/);
  if (match) return "https://drive.google.com/uc?export=view&id=" + match[1];
  return url;
}

// ── UTIL: JSON response ───────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
