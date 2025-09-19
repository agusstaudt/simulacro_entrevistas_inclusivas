function doGet(e) {
  const pagina = e.parameter.pagina;
  const archivo = pagina === "formulario" ? "FormularioCapacitaciones" : "index";

  return HtmlService.createHtmlOutputFromFile(archivo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // opcional, si embebés el sitio
    .setTitle("Formulario")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

////////////////// Funciones para el formulario inicial en atención al público ////////////////////
function buscarPersona(cuil, nombre, apellido, dia, mes, anio) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DatosImportados");
  const data = sheet.getDataRange().getValues();

  // Índices de las columnas
  const headers = data[0];
  const cuilIdx = headers.indexOf("CUIL");
  const fnacIdx = headers.indexOf("Fecha de nacimiento");
  const nombreIdx = headers.indexOf("Nombre");
  const apIdx = headers.indexOf("Apellido");
  const primeraVezIdx = headers.indexOf("¿Primera vez que concurre a la Oficina de Empleo?");
  const endFechaIdx = headers.indexOf("end");

  if ([cuilIdx, fnacIdx, nombreIdx, apIdx, primeraVezIdx, endFechaIdx].includes(-1)) {
    return "Error: No se encontraron las columnas necesarias en la hoja.";
  }

  // Fecha mínima de filtro
  const fechaFiltro = new Date("2025-07-21T00:00:00");

  // Convertir valores ingresados para búsqueda
  const cuilBuscado = String(cuil).trim();
  const nombresBuscados = normalizeString(nombre).split(" ");
  const apellidoBuscado = normalizeString(apellido);
  const fechaIngresada = `${anio}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}`;

  let personaEncontrada = false;

  // Recorrer datos
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Filtrar por fecha 'end'
    let endValue = row[endFechaIdx];
    if (!(endValue instanceof Date)) {
      try {
        endValue = new Date(endValue); // Convierte desde string ISO
      } catch (e) {
        continue; // Si no es fecha válida, salta esta fila
      }
    }
    if (endValue < fechaFiltro) {
      continue; // Salta filas más antiguas
    }

    const cuilValue = String(row[cuilIdx]).trim();
    const fnacValue = row[fnacIdx] instanceof Date ? row[fnacIdx].toISOString().split("T")[0] : String(row[fnacIdx]).trim();
    const nombreValue = normalizeString(row[nombreIdx]);
    const apValue = normalizeString(row[apIdx]);
    const primeraVezValue = String(row[primeraVezIdx]).trim().toLowerCase();

    // Chequeo por CUIL con "sí"
    if (cuilValue === cuilBuscado && (primeraVezValue === "si" || primeraVezValue === "sí")) {
      return "La persona se encuentra registrada";
    }

    // Chequeo por apellido + fecha nacimiento + coincidencia de algún nombre
    if (apValue === apellidoBuscado && fnacValue === fechaIngresada) {
      const nombresEnBase = nombreValue.split(" ");
      if (nombresBuscados.some(nombre => nombresEnBase.includes(nombre))) {
        personaEncontrada = true;
      }
    }
  }

  return personaEncontrada
    ? "La persona se encuentra registrada"
    : "La persona NO se encuentra registrada aún, REGISTRAR";
}

// Normalizar strings (minúsculas y sin acentos)
function normalizeString(str) {
  return String(str || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/////////////////////// Formulario capacitaciones ///////////////////
function extraerAnioMesDesdeFecha(fechaTexto) {
  if (!fechaTexto || typeof fechaTexto !== "string") return "";
  const partesFecha = fechaTexto.split(" ")[0]; // "21/07/2025"
  if (!partesFecha.includes("/")) return "";
  const partes = partesFecha.split("/"); // ["21", "07", "2025"]
  if (partes.length !== 3) return "";
  const anio = partes[2];
  const mes = partes[1].padStart(2, "0");
  return `${anio}-${mes}`;
}

function datosYaExisten(datos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cursos");
  const valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return false;

  const encabezados = valores[0].map(e => e.toString().trim().toLowerCase());
  const idxCuil = encabezados.indexOf("cuil");
  const idxFecha = encabezados.indexOf("fecharegistro");
  const idxCurso = encabezados.indexOf("nombrecurso");

  if (idxCuil === -1 || idxFecha === -1 || idxCurso === -1) {
    console.warn("❌ No se encontraron las columnas necesarias.");
    return false;
  }

  const cursosExcepcion = [
    "taller tips de entrevistas laborales",
    "taller fortalecimiento de cv"
  ];

  const cuilNuevo = (datos["cuil"] || "").toString().trim().toLowerCase();
  const fechaTextoNuevo = datos["fechaRegistro"] || "";
  const nombreCursoNuevo = (datos["nombreCurso"] || "").toString().trim().toLowerCase();
  const fechaNueva = extraerAnioMesDesdeFecha(fechaTextoNuevo);

  const cursoNuevoEsExcepcion = cursosExcepcion.includes(nombreCursoNuevo);

  let cursosNormalesEnMes = 0;

  for (let i = 1; i < valores.length; i++) {
    const fila = valores[i];
    const cuilExistente = (fila[idxCuil] || "").toString().trim().toLowerCase();
    const fechaTextoExistente = fila[idxFecha] || "";
    const nombreCursoExistente = (fila[idxCurso] || "").toString().trim().toLowerCase();
    const fechaExistente = extraerAnioMesDesdeFecha(fechaTextoExistente);

    const cursoExistenteEsExcepcion = cursosExcepcion.includes(nombreCursoExistente);

    // Si el mismo CUIL ya tiene un curso normal en el mes, y el nuevo también es normal → rechazar
    if (
      cuilNuevo === cuilExistente &&
      fechaNueva === fechaExistente &&
      !cursoNuevoEsExcepcion &&
      !cursoExistenteEsExcepcion
    ) {
      console.warn("💥 Coincidencia en curso normal en mismo mes ➤ Bloquear");
      return true;
    }

    // Contamos cuántos cursos normales ya tiene en ese mes
    if (
      cuilNuevo === cuilExistente &&
      fechaNueva === fechaExistente &&
      !cursoExistenteEsExcepcion
    ) {
      cursosNormalesEnMes++;
    }
  }

  // Si ya hay 1 curso normal y el nuevo también es normal → rechazar
  if (cursosNormalesEnMes >= 1 && !cursoNuevoEsExcepcion) {
    console.warn("❌ Ya tiene un curso normal ese mes ➤ No se permite otro");
    return true;
  }

  console.warn("✅ No se encontraron coincidencias relevantes");
  return false;
}

function obtenerDatosPorCUIL(cuil) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DatosImportados");
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];

  const idxCUIL = headers.indexOf("CUIL");
  const idxNombre = headers.indexOf("Nombre");
  const idxApellido = headers.indexOf("Apellido");
  const idxFecha = headers.indexOf("Fecha de nacimiento");
  const idxTel = headers.indexOf("Teléfono de contacto personal");
  const idxEmail = headers.indexOf("Correo Electrónico (email)");
  const idxPrimeraVez = headers.indexOf("¿Primera vez que concurre a la Oficina de Empleo?");

  if (idxCUIL === -1 || idxNombre === -1 || idxApellido === -1 || idxFecha === -1 || idxTel === -1 || idxEmail === -1 || idxPrimeraVez === -1) {
    return null;
  }
  
  const cuilBuscado = String(cuil).trim();

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const valorCUIL = String(fila[idxCUIL]).trim();
    const primeraVez = String(fila[idxPrimeraVez]).trim().toLowerCase();

    // Solo considerar fila si es “sí”
    if (valorCUIL === cuilBuscado && (primeraVez === "sí" || primeraVez === "si")) {
      return {
        nombre: fila[idxNombre],
        apellido: fila[idxApellido],
        fechaNacimiento: fila[idxFecha] instanceof Date
          ? Utilities.formatDate(fila[idxFecha], Session.getScriptTimeZone(), "yyyy-MM-dd")
          : '',
        telefono: fila[idxTel],
        correo: fila[idxEmail]
      };
    }
  }

  // Si no encontró con “Sí”, retorna null
  return null;
}

function guardarEnDatosCursos(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("Cursos");

  if (!hoja) {
    hoja = ss.insertSheet("Cursos");
  }

  // Orden fijo deseado
  const ordenCampos = [
    "fechaRegistro", // Nueva columna al principio
    "nombreCurso", "comision", "turno", "institucion", "duracion",
    "encuentros", "horasClase", "fechaInicio", "cuil", "nombre",
    "apellido", "fechaNacimiento", "telefono", "correo", "metodoInscripcion"
  ];

  // Si la hoja está vacía, escribir encabezados en el orden deseado
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(ordenCampos);
  }

  // Crear fila con los datos en el orden correcto + fecha registro
  const fila = [new Date(), ...ordenCampos.slice(1).map(k => datos[k] || "")];

  hoja.appendRow(fila);
}


function obtenerCursos() {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_CURSOS_ID');
  const hoja = SpreadsheetApp.openById(id).getSheetByName("Nombre de cursos");
  const datos = hoja.getRange("A2:A" + hoja.getLastRow()).getValues();
  
  const cursosLimpios = datos
    .flat()
    .map(curso => curso.toString().trim().replace(/\s+/g, ' '))  // Normaliza espacios
    .filter(curso => curso !== "");

  return [...new Set(cursosLimpios)].sort();
}

// funcion para enviar datos para planilla de asistencia
function guardarEnHojaCurso(datos) {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_CURSOS_ID');
  const libro = SpreadsheetApp.openById(id);

  const nombreCurso = datos.nombreCurso.toString().trim().replace(/\s+/g, ' ');
  const fechaInicio = datos.fechaInicio.toString().trim().replace(/\s+/g, ' ');
  const comision =  datos.comision.toString().trim().replace(/\s+/g, ' ');
  // Buscar código del curso en la hoja "Nombre cursos"
  const hojaCodigos = libro.getSheetByName('Nombre de cursos');
  const valoresCodigos = hojaCodigos.getRange('A2:B' + hojaCodigos.getLastRow()).getValues();

  let codigoCurso = null;
  for (let i = 0; i < valoresCodigos.length; i++) {
    if (valoresCodigos[i][0].toString().trim().toLowerCase() === nombreCurso.toLowerCase()) {
      codigoCurso = valoresCodigos[i][1];
      break;
    }
  }

  if (!codigoCurso) {
    throw new Error(`No se encontró la codificación para el curso: "${nombreCurso}"`);
  }

  const nombreHoja = `${codigoCurso}_com ${comision}_${fechaInicio}`;
  let hoja = libro.getSheetByName(nombreHoja);

  if (!hoja) {
    hoja = libro.insertSheet(nombreHoja);

    const duracion = parseInt(datos.duracion, 10) || 0;
    const encabezados = [
      'Turno', 'CUIL', 'Nombre y Apellido','Telefono'
    ];

    for (let i = 1; i <= duracion; i++) {
      encabezados.push(`Encuentro ${i}`);
    }

    encabezados.push('Finalizado');
    hoja.appendRow(encabezados);
  }

  const nombreApellido = formatearTitulo(datos.nombre) + ' ' + formatearTitulo(datos.apellido);
  const fila = [
    datos.turno,
    datos.cuil,
    nombreApellido,
    datos.telefono
  ];

  hoja.appendRow(fila);
} 

function formatearTitulo(texto) {
  return texto.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());
}

function obtenerCantidadInscriptos(nombreCurso, comision, fechaInicio) {
  try {
    const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_CURSOS_ID');
    const libro = SpreadsheetApp.openById(id);

    // Buscar código y cupo del curso
    const hojaCodigos = libro.getSheetByName('Nombre de cursos');
    const valoresCodigos = hojaCodigos.getRange('A2:K' + hojaCodigos.getLastRow()).getValues(); // hasta columna K

    let codigoCurso = null;
    let cupoMaximo = null;
    for (let i = 0; i < valoresCodigos.length; i++) {
      if (valoresCodigos[i][0].toString().trim().toLowerCase() === nombreCurso.toLowerCase()) {
        codigoCurso = valoresCodigos[i][1];     // Columna B
        cupoMaximo = valoresCodigos[i][9];      // Columna J (índice 9)
        break;
      }
    }

    if (!codigoCurso || !cupoMaximo) return { error: -1 };

    const nombreHoja = `${codigoCurso}_com ${comision}_${fechaInicio}`;
    const hoja = libro.getSheetByName(nombreHoja);
    const cantidadInscriptos = hoja ? hoja.getLastRow() - 1 : 0;

    return {
      cantidad: cantidadInscriptos,
      cupoMaximo: cupoMaximo
    };

  } catch (error) {
    console.warn("Error en obtenerCantidadInscriptos: " + error);
    return { error: -2 };
  }
}

// Para especificar el turno del curso seleccionado en la inscripcion
function obtenerTurnosPorCurso(nombreCurso) {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_CURSOS_ID');
  const libro = SpreadsheetApp.openById(id);
  const hoja = libro.getSheetByName("Nombre de cursos");
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(e => e.toString().trim().toLowerCase());

  const idxCurso = encabezados.indexOf("listado de cursos");
  const idxTurno = encabezados.indexOf("turno");

  if (idxCurso === -1 || idxTurno === -1) return [];

  for (let i = 1; i < datos.length; i++) {
    const curso = datos[i][idxCurso].toString().trim().toLowerCase();
    if (curso === nombreCurso.trim().toLowerCase()) {
      const turnos = datos[i][idxTurno];
      return turnos ? turnos.split(",").map(t => t.trim()) : [];
    }
  }

  return [];
}

// Para especificar la comision segun el curso elegido
function obtenerComisionesPorCurso(nombreCurso) {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_CURSOS_ID');
  const libro = SpreadsheetApp.openById(id);
  const hoja = libro.getSheetByName("Nombre de cursos");

  if (!hoja) {
    console.warn("❌ Hoja 'Nombre de cursos' no encontrada");
    return [];
  }

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(e => e.toString().trim().toLowerCase());
  const idxCurso = encabezados.indexOf("listado de cursos");
  const idxComision = encabezados.indexOf("comisión");

  if (idxCurso === -1 || idxComision === -1) {
    console.warn("❌ Encabezados 'curso' o 'comision' no encontrados");
    return [];
  }

  for (let i = 1; i < datos.length; i++) {
    const curso = datos[i][idxCurso]?.toString().trim().toLowerCase();
    if (curso === nombreCurso.toLowerCase()) {
      const valor = datos[i][idxComision]?.toString().trim();
      if (!valor) return [];

      if (valor.includes('-')) {
        // Rango tipo 1-3
        const [inicio, fin] = valor.split('-').map(Number);
        if (!isNaN(inicio) && !isNaN(fin)) {
          return Array.from({ length: fin - inicio + 1 }, (_, i) => (inicio + i).toString());
        }
      } else if (valor.includes(',')) {
        // Lista tipo 1,2,3
        return valor.split(',').map(e => e.trim());
      } else {
        // Un solo número
        return [valor];
      }
    }
  }

  return [];
}

function obtenerDatosCurso(nombreCurso) {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_CURSOS_ID');
  const libro = SpreadsheetApp.openById(id);
  const hoja = libro.getSheetByName("Nombre de cursos");

  if (!hoja) return {};

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(e => e.toString().trim().toLowerCase());

  const idxCurso = encabezados.indexOf("listado de cursos");
  const idxInstitucion = encabezados.indexOf("institución");
  const idxDuracion = encabezados.indexOf("duración");
  const idxEncuentros = encabezados.indexOf("encuentros semanales");
  const idxHoras = encabezados.indexOf("horas clase");
  const idxFecha = encabezados.indexOf("fecha de inicio");

  if (idxCurso === -1) return {};

  const opciones = {
    institucion: new Set(),
    duracion: new Set(),
    encuentros: new Set(),
    horasClase: new Set(),
    fechaInicio: new Set()
  };

  for (let i = 1; i < datos.length; i++) {
    const cursoHoja = datos[i][idxCurso]?.toString().trim().toLowerCase().replace(/\s+/g, ' ');
    const cursoBuscado = nombreCurso.trim().toLowerCase().replace(/\s+/g, ' ');

    if (cursoHoja === cursoBuscado) {
      opciones.institucion.add(datos[i][idxInstitucion]);

      (datos[i][idxDuracion] || '').toString().split(';').forEach(v => opciones.duracion.add(v.trim()));
      (datos[i][idxEncuentros] || '').toString().split(';').forEach(v => opciones.encuentros.add(v.trim()));
      (datos[i][idxHoras] || '').toString().split(';').forEach(v => opciones.horasClase.add(v.trim()));
      (datos[i][idxFecha] || '').toString().split(';').forEach(v => opciones.fechaInicio.add(v.trim()));
    }
  }
  // Convertir sets a arrays únicos
  return {
    institucion: [...opciones.institucion].filter(Boolean),
    duracion: [...opciones.duracion].filter(Boolean),
    encuentros: [...opciones.encuentros].filter(Boolean),
    horasClase: [...opciones.horasClase].filter(Boolean),
    fechaInicio: [...opciones.fechaInicio].filter(Boolean)
  };
}