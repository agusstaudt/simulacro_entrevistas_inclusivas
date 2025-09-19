function doGet() {
  const template = HtmlService.createTemplateFromFile("index");
  template.checkboxesRolTrabajado = generarCheckboxesRoles("rolTrabajado");
  template.checkboxesRolDeseado = generarCheckboxesRoles("rolDeseado");
  const html = template.evaluate();
  return html
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Formulario de inscripción - Simulacro de entrevistas")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function obtenerRoles() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roles");
  const valores = hoja.getRange("A2:A" + hoja.getLastRow()).getValues();
  return valores.flat().filter(r => r);
}

function escHtml_(s) {
  return String(s)
   .replace(/&/g,"&amp;").replace(/</g,"&lt;")
   .replace(/>/g,"&gt;").replace(/"/g,"&quot;")
   .replace(/'/g,"&#39;");
}

function generarCheckboxesRoles(prefix) {
  const roles = obtenerRoles();
  return roles.map(role => {
    const safeRole = String(role);
    const idSafe = safeRole.replace(/\s+/g, '-').replace(/[^\w-]/g, '');
    const id = `${prefix}-${idSafe}`;
    const labelTxt = escHtml_(safeRole);
    const valueAttr = escHtml_(safeRole);
    return `
      <li class="w-full border-b border-gray-200">
        <div class="flex items-center pl-3">
          <input id="${id}" name="${prefix}" type="checkbox" value="${valueAttr}"
            class="w-4 h-4 text-green-600 bg-gray-100 border-gray-300 rounded-sm focus:ring-green-500 focus:ring-2 checked:bg-green-600 checked:border-transparent">
          <label for="${id}" class="w-full py-3 ml-2 text-sm font-medium text-gray-900">${labelTxt}</label>
        </div>
      </li>
    `;
  }).join('\n');
}

/******************** Helpers de seguridad/sanitización ********************/

// Normaliza strings: quita control chars, colapsa espacios y recorta a max
function sanitizeStr_(s, max = 120) {
  return String(s || "")
    .replace(/[\u0000-\u001F\u007F]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max);
}

// Devuelve solo dígitos
function onlyDigits_(s) {
  return String(s || "").replace(/\D/g, "");
}

// Convierte valor a array (checkboxes), o [] si vacío
function toArr_(v) {
  if (Array.isArray(v)) return v;
  if (v == null || v === "") return [];
  return [String(v)];
}

// Throttle simple por clave (p.ej., por CUIL/email). Evita reintentos masivos.
function throttle_(key, seconds) {
  const cache = CacheService.getUserCache();
  const hit = cache.get(key);
  if (hit) throw new Error("Demasiadas solicitudes, por favor intentá en unos segundos.");
  cache.put(key, "1", seconds);
}

/******************** Función principal endurecida ********************/

function guardarEnDatosCursos_v3(datos) {
  // --- Normalizaciones iniciales crudas (para validar con seguridad) ---
  const cuilDigits = onlyDigits_(datos.cuil);
  const correo = String(datos.correo || "").trim().toLowerCase();

  // --- Rate limit: 1 envío cada 10s por usuario lógico ---
  const throttleKey = "form:" + (cuilDigits || correo || Utilities.getUuid().slice(0, 8));
  throttle_(throttleKey, 10);

  // --- Validaciones duras ---
  if (cuilDigits.length !== 11) {
    throw new Error("CUIL debe tener 11 dígitos.");
  }

  const f = String(datos.fechaNacimiento || "");
  const [d, m, a] = f.split("/");
  const dia = parseInt(d, 10), mes = parseInt(m, 10), anio = parseInt(a, 10);
  const currentYear = new Date().getFullYear();
  if (
    !(d && m && a && d.length === 2 && m.length === 2 && a.length === 4 &&
      dia >= 1 && dia <= 31 && mes >= 1 && mes <= 12 &&
      anio >= 1900 && anio <= currentYear)
  ) {
    throw new Error("Fecha de nacimiento inválida.");
  }

  const reEmail = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  if (!reEmail.test(correo)) {
    throw new Error("Correo inválido.");
  }

  // Validaciones condicionales de domicilio
  const vivePosadas = (datos.vivePosadas === "si" || datos.vivePosadas === "no") ? datos.vivePosadas : "";
  if (vivePosadas === "si" && !String(datos.barrio || "").trim()) {
    throw new Error("Completá barrio/zona.");
  }
  if (vivePosadas === "no" && !String(datos.otraLocalidad || "").trim()) {
    throw new Error("Indicá localidad.");
  }

  // --- Whitelist de roles válidos ---
  // Asumimos que obtenerRoles() devuelve un array de strings con la lista canónica
  const rolesValidos = new Set((obtenerRoles() || []).map(String));
  const rolesDeseadosRaw = toArr_(datos.rolDeseado);
  const rolesTrabajadosRaw = toArr_(datos.rolTrabajado);

  const rolesDeseados = rolesDeseadosRaw.filter(r => rolesValidos.has(String(r)));
  const rolesTrabajados = rolesTrabajadosRaw.filter(r => rolesValidos.has(String(r)));

  if (rolesDeseados.length === 0) {
    throw new Error("Seleccioná al menos un rol deseado válido.");
  }

  // --- Sanitización final y límites de longitud ---
  const limpio = {
    cuil: cuilDigits,                                           // 11 dígitos
    nombre: sanitizeStr_(datos.nombre, 80),
    apellido: sanitizeStr_(datos.apellido, 80),
    fechaNacimiento: `${d}/${m}/${a}`,                           // ya validada
    edad: onlyDigits_(datos.edad).slice(0, 3),                   // 0-999 (texto)
    telefono: onlyDigits_(datos.telefono).slice(0, 20),
    correo,                                                      // normalizado lower
    vivePosadas,                                                 // "si"/"no" o ""
    barrio: sanitizeStr_(datos.barrio, 120),
    otraLocalidad: sanitizeStr_(datos.otraLocalidad, 120),
    experienciaPrevia: (datos.experienciaPrevia === "si" || datos.experienciaPrevia === "no") ? datos.experienciaPrevia : "",
    rolTrabajado: rolesTrabajados.join(", "),
    otroTextoExperiencia: sanitizeStr_(datos.otroTextoExperiencia, 200),
    rolDeseado: rolesDeseados.join(", "),
    otroTextoDeseados: sanitizeStr_(datos.otroTextoDeseados, 200),
  };

  // --- Hoja destino local ---
  const columnas = [
    "fechaRegistro","cuil","nombre","apellido",
    "fechaNacimiento","edad","telefono","correo",
    "vivePosadas","barrio","otraLocalidad",
    "experienciaPrevia","rolTrabajado","otroTextoExperiencia",
    "rolDeseado","otroTextoDeseados"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("Registro");
  if (!hoja) hoja = ss.insertSheet("Registro");
  if (hoja.getLastRow() === 0) hoja.appendRow(columnas);

  const fechaRegistro = new Date();
  const fila = columnas.map(col => {
    if (col === "fechaRegistro") return fechaRegistro;
    return (col in limpio) ? limpio[col] : "";
  });
  hoja.appendRow(fila);

  // --- Guardado externo ---
  const externoOk = guardarEnExterno_v3(fechaRegistro, limpio);

  // --- Respuesta “segura” ---
  return {
    ok: true,
    externo: !!externoOk,
    ts: new Date().toISOString()
  };
}

function getProp_(k) {
  return PropertiesService.getScriptProperties().getProperty(k);
}

function guardarEnExterno_v3(datos, fechaRegistro) {
  const DEST_ID = getProp_("DEST_ID"); // <- ya no hardcodeado
  if (!DEST_ID) throw new Error("DEST_ID no configurado en Script Properties");
  const NOMBRE_HOJA = "Planilla de inscriptos";
  const ssExterno = SpreadsheetApp.openById(DEST_ID);
  let hoja = ssExterno.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = ssExterno.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(["fechaRegistro","cuil","apellido","nombre"]);
  }
  hoja.appendRow([ fechaRegistro, datos.cuil||"", datos.apellido||"", datos.nombre||"" ]);
  return true;
}