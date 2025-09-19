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

/******************** Helpers ********************/
function sanitizeStr_(s, max = 120) {
  return String(s || "")
    .replace(/[\u0000-\u001F\u007F]/g, "") // control chars
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max);
}
function onlyDigits_(s) { return String(s || "").replace(/\D/g, ""); }
function toArr_(v) {
  if (Array.isArray(v)) return v;
  if (v == null || v === "") return [];
  return [String(v)];
}
function throttle_(key, seconds) {
  const cache = CacheService.getUserCache();
  const hit = cache.get(key);
  if (hit) throw new Error("Demasiadas solicitudes, intentá en unos segundos.");
  cache.put(key, "1", seconds);
}

/******************** Guardado principal ********************/
function guardarEnDatosCursos(datos) {
  // Rate-limit: 1 envío/10s por usuario lógico
  const cuilDigits = onlyDigits_(datos.cuil);
  const correoNorm = String(datos.correo || "").trim().toLowerCase();
  const throttleKey = "form:" + (cuilDigits || correoNorm || Utilities.getUuid().slice(0,8));
  throttle_(throttleKey, 10);

  // Validaciones duras
  if (cuilDigits.length !== 11) throw new Error("CUIL debe tener 11 dígitos.");

  // Fecha: viene en campos separados (dia, mes, anio)
  const fDia = String(datos.dia || "");
  const fMes = String(datos.mes || "");
  const fAnio = String(datos.anio || "");
  const dia = parseInt(fDia, 10), mes = parseInt(fMes,10), anio = parseInt(fAnio,10);
  const currentYear = new Date().getFullYear();
  if (!(fDia.length===2 && fMes.length===2 && fAnio.length===4 &&
        dia>=1 && dia<=31 && mes>=1 && mes<=12 && anio>=1900 && anio<=currentYear)) {
    throw new Error("Fecha de nacimiento inválida.");
  }
  const fechaNacimiento = `${fDia}/${fMes}/${fAnio}`;

  const reEmail = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  if (!reEmail.test(correoNorm)) throw new Error("Correo inválido.");

  const vivePosadas = (datos.vivePosadas === "si" || datos.vivePosadas === "no") ? datos.vivePosadas : "";
  if (vivePosadas === "si" && !String(datos.barrio || "").trim()) throw new Error("Completá barrio/zona.");
  if (vivePosadas === "no" && !String(datos.otraLocalidad || "").trim()) throw new Error("Indicá localidad.");

  // Whitelist de roles válidos
  const rolesValidos = new Set((obtenerRoles() || []).map(String));
  const rolesTrabajados = toArr_(datos.rolTrabajado).filter(r => rolesValidos.has(String(r)));
  const rolesDeseados   = toArr_(datos.rolDeseado).filter(r => rolesValidos.has(String(r)));
  if (rolesDeseados.length === 0) throw new Error("Seleccioná al menos un rol deseado válido.");

  // Sanitización y límites
  const limpio = {
    cuil: cuilDigits,
    nombre: sanitizeStr_(datos.nombre, 80),
    apellido: sanitizeStr_(datos.apellido, 80),
    fechaNacimiento,
    telefono: onlyDigits_(datos.telefono).slice(0, 20),
    correo: correoNorm,
    vivePosadas,
    barrio: sanitizeStr_(datos.barrio, 120),
    otraLocalidad: sanitizeStr_(datos.otraLocalidad, 120),
    cud: (["si","no","enTramite"].includes(datos.cud)) ? datos.cud : "",
    tipoDiscapacidad: (["motriz","visual","auditiva","intelectual","psicosocial","otra"].includes(datos.tipoDiscapacidad)) ? datos.tipoDiscapacidad : "",
    otraDiscapacidad: sanitizeStr_(datos.otraDiscapacidad, 200),
    experienciaPrevia: (["si","no"].includes(datos.experienciaPrevia)) ? datos.experienciaPrevia : "",
    rolTrabajado: rolesTrabajados.join(", "),
    otroTextoExperiencia: sanitizeStr_(datos.otroTextoExperiencia, 200),
    rolDeseado: rolesDeseados.join(", "),
    otroTextoDeseados: sanitizeStr_(datos.otroTextoDeseados, 200),
    disponibilidadHoraria: (["mañana","tarde"].includes(datos.disponibilidadHoraria)) ? datos.disponibilidadHoraria : "",
    necesidadApoyo: (["si","no"].includes(datos.necesidadApoyo)) ? datos.necesidadApoyo : "",
    otroApoyo: sanitizeStr_(datos.otroApoyo, 200),
  };

  // Hoja local “Registro” con encabezados
  const columnas = [
    "fechaRegistro",
    "cuil","nombre","apellido",
    "fechaNacimiento",
    "telefono","correo",
    "vivePosadas","barrio","otraLocalidad",
    "cud","tipoDiscapacidad","otraDiscapacidad",
    "experienciaPrevia","rolTrabajado","otroTextoExperiencia",
    "rolDeseado","otroTextoDeseados",
    "disponibilidadHoraria","necesidadApoyo","otroApoyo"
  ];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("Registro");
  if (!hoja) hoja = ss.insertSheet("Registro");
  if (hoja.getLastRow() === 0) hoja.appendRow(columnas);

  const fechaRegistro = new Date();
  const fila = columnas.map(col => (col === "fechaRegistro") ? fechaRegistro : (limpio[col] ?? ""));
  hoja.appendRow(fila);

  // Guardar en hoja externa (ID desde Script Properties)
  try {
    const destId = getProp_("DEST_ID"); // configurado en Script Properties
    if (!destId) throw new Error("DEST_ID no configurado");
    const ssExterno = SpreadsheetApp.openById(destId);
    let hojaExterna = ssExterno.getSheetByName("Panilla de inscriptos");
    if (!hojaExterna) {
      hojaExterna = ssExterno.insertSheet("Panilla de inscriptos");
      hojaExterna.appendRow(["fechaRegistro","cuil","apellido","nombre"]);
    }
    hojaExterna.appendRow([fechaRegistro, limpio.cuil, limpio.apellido, limpio.nombre]);
  } catch (e) {
    console.warn("Error al guardar en la hoja externa: " + e);
  }

  return { ok: true, ts: new Date().toISOString() };
}

// Obtener roles de la hoja de g sheets
function obtenerRoles() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Roles");
  const valores = hoja.getRange("A2:A" + hoja.getLastRow()).getValues(); // Asume que los roles están desde A2
  const roles = valores.flat().filter(r => r); // Elimina valores vacíos
  return roles;
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
    const valueAttr = escHtml_(safeRole);
    const labelTxt = escHtml_(safeRole);
    return `
      <li class="w-full border-b border-gray-200">
        <div class="flex items-center pl-3">
          <input id="${id}" name="${prefix}" type="checkbox" value="${valueAttr}"
            class="w-4 h-4 text-green-600 bg-gray-100 border-gray-300 rounded-sm focus:ring-green-500 focus:ring-2 checked:bg-green-600 checked:border-transparent" required>
          <label for="${id}" class="w-full py-3 ml-2 text-sm font-medium text-gray-900">${labelTxt}</label>
        </div>
      </li>
    `;
  }).join('\n');
}