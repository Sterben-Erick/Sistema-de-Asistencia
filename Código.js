function doGet(e) {
  const pagina = e && e.parameter && e.parameter.pagina;
  return pagina === "admin"
    ? HtmlService.createTemplateFromFile("admin").evaluate().setTitle("Administración - Sistema de Asistencia")
    : HtmlService.createTemplateFromFile("index").evaluate().setTitle("Sistema de Asistencia");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function buscarEstudiante(dni) {
  dni = String(dni || "").trim();
  if (!dni) return { encontrado: false };
  const mapa = obtenerMapaEstudiantes_();
  const e = mapa[dni];
  if (!e) return { encontrado: false };
  return {
    encontrado: true,
    dni: dni,
    nombre: e.nombre,
    apellidos: e.apellidos,
    carrera: e.carrera,
    ciclo: e.ciclo
  };
}

function registrarAsistencia(dni, metodo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAsistencias = ss.getSheetByName("Asistencias");

  dni = normalizarDni_(dni, false); if (!/^\d{8}$/.test(dni)) return { estado: "ERROR", mensaje: "El DNI debe tener exactamente 8 dígitos." };
  metodo = String(metodo || "").trim();

  if (!dni) return { estado: "ERROR", mensaje: "El DNI es obligatorio." };

  const mapaEst = obtenerMapaEstudiantes_();
  const e = mapaEst[dni];
  let estudiante = e ? {
    dni: dni,
    nombre: e.nombre,
    apellidos: e.apellidos,
    carrera: e.carrera,
    ciclo: e.ciclo,
    turno: e.turno
  } : null;

  const mapaPer = obtenerMapaPersonal_();
  const p = mapaPer[dni];
  let personal = p ? {
    dni: dni,
    nombre: p.nombre,
    apellidos: p.apellidos,
    tipo: p.tipo,
    cargo: p.cargo,
    estado: p.estado
  } : null;

  if (!estudiante && !personal) return { estado: "NO_ENCONTRADO", versionPadron: versionPadron_() };

  const ahora = new Date();
  const fechaRegistro = new Date(ahora);
  fechaRegistro.setHours(0, 0, 0, 0);

  const horaRegistro = new Date(ahora);
  const vPadron = versionPadron_();
  const bloquearDuplicado = obtenerConfiguracionBoolean("BLOQUEAR_DUPLICADO_DIARIO", true);
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const lastA = hojaAsistencias.getLastRow();
    const datosAsistencias = lastA >= 2 ? hojaAsistencias.getRange(2, 1, lastA - 1, 8).getValues() : [];

    if (estudiante) {
      for (let i = 0; i < datosAsistencias.length; i++) {
        const fecha = datosAsistencias[i][1];
        const estadoRegistro = String(datosAsistencias[i][5]).trim().toUpperCase();

        if (
          normalizarDni_(datosAsistencias[i][0]) === dni &&
          fecha instanceof Date &&
          esMismaFecha(fecha, ahora) &&
          estadoRegistro === "ENTRADA" &&
          bloquearDuplicado
        ) {
          return {
            estado: "DUPLICADO",
            versionPadron: vPadron,
            tipoPersona: "ESTUDIANTE",
            estudiante: estudiante,
            horaAnterior: formatearFecha(datosAsistencias[i][2], "hh:mm:ss a")
          };
        }
      }

      hojaAsistencias.appendRow([
        estudiante.dni,
        fechaRegistro,
        horaRegistro,
        metodo,
        "ESTUDIANTE",
        "ENTRADA",
        estudiante.carrera,
        estudiante.ciclo
      ]);

      const estadoControl = obtenerEstadoAsistenciaPorTurno_(
        estudiante.turno,
        ahora
      );

      actualizarEstadoControlAsistencia_(
        ss,
        estudiante.dni,
        fechaRegistro,
        estadoControl
      );

      return {
        estado: "REGISTRADO",
        versionPadron: vPadron,
        tipoPersona: "ESTUDIANTE",
        estudiante: estudiante,
        fecha: formatearFecha(ahora, "dd/MM/yyyy"),
        hora: formatearFecha(ahora, "hh:mm:ss a")
      };
    }

    if (String(personal.estado).trim().toUpperCase() !== "ACTIVO") {
      return {
        estado: "PERSONAL_INACTIVO",
        versionPadron: vPadron,
        tipoPersona: personal.tipo,
        personal: personal,
        mensaje: "El personal se encuentra inactivo y no puede registrar asistencia."
      };
    }

    let tieneEntrada = false;
    let tieneSalida = false;
    let fechaEntrada = null;

    for (let i = 0; i < datosAsistencias.length; i++) {
      const fecha = datosAsistencias[i][1];
      const hora = datosAsistencias[i][2];
      const tipoRegistro = String(datosAsistencias[i][5]).trim().toUpperCase();

      if (
        normalizarDni_(datosAsistencias[i][0]) === dni &&
        fecha instanceof Date &&
        esMismaFecha(fecha, ahora)
      ) {
        if (tipoRegistro === "ENTRADA") {
          tieneEntrada = true;

          if (hora instanceof Date) {
            fechaEntrada = new Date(fecha);
            fechaEntrada.setHours(
              hora.getHours(),
              hora.getMinutes(),
              hora.getSeconds(),
              0
            );
          }
        }

        if (tipoRegistro === "SALIDA") {
          tieneSalida = true;
        }
      }
    }

    if (tieneEntrada && tieneSalida) {
      return {
        estado: "PERSONAL_COMPLETO",
        versionPadron: vPadron,
        tipoPersona: personal.tipo,
        personal: personal
      };
    }

    if (tieneEntrada && !tieneSalida) {
      const tiempoMinimoSalida = obtenerTiempoMinimoSalida();

      if (fechaEntrada) {
        const diferencia = (ahora.getTime() - fechaEntrada.getTime()) / (1000 * 60);

        if (diferencia < tiempoMinimoSalida) {
          return {
            estado: "PERSONAL_TIEMPO_MINIMO",
            versionPadron: vPadron,
            tipoPersona: personal.tipo,
            personal: personal,
            minutosRestantes: Math.ceil(tiempoMinimoSalida - diferencia)
          };
        }
      }

      const foto = obtenerFotoPersonal(personal.dni);

      hojaAsistencias.appendRow([
        personal.dni,
        fechaRegistro,
        horaRegistro,
        metodo,
        personal.tipo,
        "SALIDA",
        "",
        ""
      ]);

      return {
        estado: "PERSONAL_SALIDA_REGISTRADA",
        versionPadron: vPadron,
        tipoPersona: personal.tipo,
        tipoRegistro: "SALIDA",
        personal: personal,
        fecha: formatearFecha(ahora, "dd/MM/yyyy"),
        hora: formatearFecha(ahora, "hh:mm:ss a"),
        foto: foto
      };
    }

    const foto = obtenerFotoPersonal(personal.dni);

    hojaAsistencias.appendRow([
      personal.dni,
      fechaRegistro,
      horaRegistro,
      metodo,
      personal.tipo,
      "ENTRADA",
      "",
      ""
    ]);

    return {
      estado: "PERSONAL_ENTRADA_REGISTRADA",
      versionPadron: vPadron,
      tipoPersona: personal.tipo,
      tipoRegistro: "ENTRADA",
      personal: personal,
      fecha: formatearFecha(ahora, "dd/MM/yyyy"),
      hora: formatearFecha(ahora, "hh:mm:ss a"),
      foto: foto
    };
  } finally {
    lock.releaseLock();
  }
}

function cerrarControlAsistenciaHoy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaControl = ss.getSheetByName("Control de Asistencia");
  const lock = LockService.getScriptLock();

  if (!hojaControl) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "Control de Asistencia".'
    };
  }

  lock.waitLock(5000);

  try {
    const ahora = new Date();

    const fecha = new Date(ahora);
    fecha.setHours(0, 0, 0, 0);

    asegurarControlAsistenciaHoy_(ss, fecha);

    const turnos = obtenerConfiguracionTurnos_();

    const turnosCerrados = [];

    for (const turno of Object.keys(turnos)) {
      const horaCierre = turnos[turno].cierre;
      const [hora, minuto] = horaCierre.split(":").map(Number);

      const cierre = new Date(fecha);
      cierre.setHours(hora, minuto, 0, 0);

      if (ahora.getTime() >= cierre.getTime()) {
        turnosCerrados.push(turno);
      }
    }

    if (!turnosCerrados.length) {
      return {
        estado: "OK",
        fecha: formatearFecha(fecha, "dd/MM/yyyy"),
        faltas: 0,
        turnosCerrados: [],
        mensaje: "Todavía no corresponde cerrar ningún turno."
      };
    }

    const ultimaFila = hojaControl.getLastRow();

    if (ultimaFila < 2) {
      return {
        estado: "OK",
        fecha: formatearFecha(fecha, "dd/MM/yyyy"),
        faltas: 0,
        turnosCerrados: turnosCerrados
      };
    }

    const hojaEstudiantes = ss.getSheetByName("Estudiantes");
    const mapaCreacion = {};
    if (hojaEstudiantes && hojaEstudiantes.getLastRow() >= 2) {
      hojaEstudiantes.getRange(2, 1, hojaEstudiantes.getLastRow() - 1, 7).getValues().forEach(r => {
        const dni = normalizarDni_(r[0]);
        if (dni) mapaCreacion[dni] = r[6];
      });
    }

    const datos = hojaControl
      .getRange(2, 1, ultimaFila - 1, 12)
      .getValues();

    let faltas = 0;

    for (let i = 0; i < datos.length; i++) {
      const fechaControl = datos[i][1];
      const dniControl = normalizarDni_(datos[i][2]);
      const turno = String(datos[i][4] || "").trim().toUpperCase();
      const estado = String(datos[i][5] || "").trim().toUpperCase();

      if (
        fechaControl instanceof Date &&
        esMismaFecha(fechaControl, fecha) &&
        turnosCerrados.includes(turno) &&
        estado === "ESPERANDO"
      ) {
        const fc = mapaCreacion[dniControl];
        if (fc instanceof Date && esMismaFecha(fc, fecha)) continue;
        datos[i][5] = "FALTÓ";
        faltas++;
      }
    }

    if (faltas > 0) {
      hojaControl
        .getRange(2, 1, datos.length, 12)
        .setValues(datos);
    }

    return {
      estado: "OK",
      fecha: formatearFecha(fecha, "dd/MM/yyyy"),
      faltas: faltas,
      turnosCerrados: turnosCerrados,
      mensaje: faltas === 0
        ? "No quedaron alumnos pendientes en los turnos cerrados."
        : faltas + " alumno(s) marcado(s) como FALTÓ."
    };

  } finally {
    lock.releaseLock();
  }
}

function inicializarControlAsistencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();

  lock.waitLock(5000);

  try {
    const fecha = new Date();
    fecha.setHours(0, 0, 0, 0);

    asegurarControlAsistenciaHoy_(ss, fecha);

    return {
      estado: "OK",
      fecha: formatearFecha(fecha, "dd/MM/yyyy")
    };
  } finally {
    lock.releaseLock();
  }
}

function asegurarControlAsistenciaHoy_(ss, fecha) {
  const hojaControl = ss.getSheetByName("Control de Asistencia");
  const hojaEstudiantes = ss.getSheetByName("Estudiantes");

  if (!hojaControl || !hojaEstudiantes) {
    throw new Error("No se encontró Control de Asistencia o Estudiantes.");
  }

  const datosEstudiantes = hojaEstudiantes.getDataRange().getValues();

  if (datosEstudiantes.length < 2) return;

  const ultimaFilaControl = hojaControl.getLastRow();

  const datosControl = ultimaFilaControl >= 2
    ? hojaControl.getRange(2, 1, ultimaFilaControl - 1, 12).getValues()
    : [];

  const existentes = new Set();

  for (const fila of datosControl) {
    const fechaControl = fila[1];
    const dni = String(fila[2] || "").trim();
    const tipoPersona = String(fila[3] || "").trim().toUpperCase();

    if (
      fechaControl instanceof Date &&
      dni &&
      tipoPersona === "ESTUDIANTE" &&
      esMismaFecha(fechaControl, fecha)
    ) {
      existentes.add(dni);
    }
  }

  const nuevos = [];

  for (let i = 1; i < datosEstudiantes.length; i++) {
    const dni = String(datosEstudiantes[i][0] || "").trim();
    const turno = String(datosEstudiantes[i][5] || "").trim().toUpperCase();

    if (!dni || existentes.has(dni)) continue;

    nuevos.push([
      Utilities.getUuid(),
      new Date(fecha),
      dni,
      "ESTUDIANTE",
      turno,
      "ESPERANDO",
      "",
      "",
      "",
      "",
      "",
      ""
    ]);
  }

  if (!nuevos.length) return;

  const filaInicial = hojaControl.getLastRow() + 1;

  hojaControl
    .getRange(filaInicial, 1, nuevos.length, 12)
    .setValues(nuevos);

  hojaControl
    .getRange(filaInicial, 2, nuevos.length, 1)
    .setNumberFormat("dd/MM/yyyy");

  hojaControl
    .getRange(filaInicial, 10, nuevos.length, 1)
    .setNumberFormat("dd/MM/yyyy hh:mm:ss AM/PM");
}

function actualizarEstadoControlAsistencia_(ss, dni, fecha, estado) {
  const hojaControl = ss.getSheetByName("Control de Asistencia");
  if (!hojaControl) {
    throw new Error('No existe la hoja "Control de Asistencia".');
  }
  dni = String(dni).trim();
  const ultimaFila = hojaControl.getLastRow();
  const datos = ultimaFila >= 2
    ? hojaControl.getRange(2, 1, ultimaFila - 1, 12).getValues()
    : [];
  const existentes = new Set();
  let filaDni = -1;
  for (let i = 0; i < datos.length; i++) {
    const fechaControl = datos[i][1];
    const dniControl = String(datos[i][2] || "").trim();
    const tipoPersona = String(datos[i][3] || "").trim().toUpperCase();
    if (!(fechaControl instanceof Date) || !esMismaFecha(fechaControl, fecha)) continue;
    if (tipoPersona !== "ESTUDIANTE" || !dniControl) continue;
    existentes.add(dniControl);
    if (dniControl === dni && filaDni === -1) filaDni = i + 2;
  }
  const faltantes = [];
  if (filaDni === -1 || true) {
    const mapaEst = obtenerMapaEstudiantes_();
    Object.keys(mapaEst).forEach(d => {
      if (existentes.has(d)) return;
      faltantes.push([
        Utilities.getUuid(), new Date(fecha), d, "ESTUDIANTE",
        String((mapaEst[d] || {}).turno || "").trim().toUpperCase(),
        "ESPERANDO", "", "", "", "", "", ""
      ]);
      existentes.add(d);
      if (d === dni && filaDni === -1) filaDni = hojaControl.getLastRow() + faltantes.length;
    });
    if (faltantes.length) {
      hojaControl.getRange(hojaControl.getLastRow() + 1, 1, faltantes.length, 12).setValues(faltantes);
    }
  }
  if (filaDni === -1) return false;
  hojaControl.getRange(filaDni, 6).setValue(estado);
  return true;
}

function obtenerHoraServidor() {
  const ahora = new Date();
  return {
    estado: "OK",
    epoch: ahora.getTime(),
    fecha: Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    hora: Utilities.formatDate(ahora, Session.getScriptTimeZone(), "hh:mm:ss a")
  };
}

function esMismaFecha(fecha1, fecha2) {
  return fecha1.getFullYear() === fecha2.getFullYear() &&
    fecha1.getMonth() === fecha2.getMonth() &&
    fecha1.getDate() === fecha2.getDate();
}

function formatearFecha(fecha, formato) {
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), formato);
}

function obtenerConfiguracion(clave, valorDefecto = "") {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIGURACION");
  if (!hoja) return valorDefecto;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return valorDefecto;

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 2).getValues();

  for (const fila of datos) {
    if (String(fila[0]).trim() === clave) {
      return fila[1] !== "" ? fila[1] : valorDefecto;
    }
  }

  return valorDefecto;
}

function obtenerConfiguracionBoolean(clave, valorDefecto = false) {
  const valor = String(obtenerConfiguracion(clave, valorDefecto ? "SI" : "NO")).trim().toUpperCase();
  return valor === "SI";
}

function obtenerTiempoMinimoSalida() {
  const valor = Number(obtenerConfiguracion("TIEMPO_MINIMO_SALIDA", 60));
  return !isNaN(valor) && valor >= 0 ? valor : 60;
}

function obtenerHoraConfiguracion_(clave, horaDefecto) {
  const valor = String(obtenerConfiguracion(clave, horaDefecto)).trim();

  if (!/^\d{2}:\d{2}$/.test(valor)) {
    return horaDefecto;
  }

  const partes = valor.split(":");
  const hora = Number(partes[0]);
  const minuto = Number(partes[1]);

  if (
    isNaN(hora) ||
    isNaN(minuto) ||
    hora < 0 ||
    hora > 23 ||
    minuto < 0 ||
    minuto > 59
  ) {
    return horaDefecto;
  }

  return valor;
}

function obtenerConfiguracionTurnos_() {
  return {
    MAÑANA: {
      ingreso: obtenerHoraConfiguracion_("HORA_INGRESO_MANANA", "08:00"),
      cierre: obtenerHoraConfiguracion_("HORA_CIERRE_MANANA", "10:00")
    },
    TARDE: {
      ingreso: obtenerHoraConfiguracion_("HORA_INGRESO_TARDE", "14:00"),
      cierre: obtenerHoraConfiguracion_("HORA_CIERRE_TARDE", "16:00")
    },
    NOCHE: {
      ingreso: obtenerHoraConfiguracion_("HORA_INGRESO_NOCHE", "19:00"),
      cierre: obtenerHoraConfiguracion_("HORA_CIERRE_NOCHE", "21:00")
    }
  };
}

function obtenerConfiguracionesSistema(token) {
  const autorizacion = validarAutorizacion(token, "CONFIGURACION_VER");

  if (!autorizacion.autorizado) return autorizacion;

  return {
    estado: "OK",
    configuraciones: {
      TIEMPO_MINIMO_SALIDA: obtenerTiempoMinimoSalida(),
      HORA_INGRESO_MANANA: obtenerHoraConfiguracion_("HORA_INGRESO_MANANA", "08:00"),
      HORA_CIERRE_MANANA: obtenerHoraConfiguracion_("HORA_CIERRE_MANANA", "10:00"),
      HORA_INGRESO_TARDE: obtenerHoraConfiguracion_("HORA_INGRESO_TARDE", "14:00"),
      HORA_CIERRE_TARDE: obtenerHoraConfiguracion_("HORA_CIERRE_TARDE", "16:00"),
      HORA_INGRESO_NOCHE: obtenerHoraConfiguracion_("HORA_INGRESO_NOCHE", "19:00"),
      HORA_CIERRE_NOCHE: obtenerHoraConfiguracion_("HORA_CIERRE_NOCHE", "21:00"),
      PERMITIR_REGISTRO_MANUAL: obtenerConfiguracionBoolean("PERMITIR_REGISTRO_MANUAL", true),
      PERMITIR_REGISTRO_SCANNER: obtenerConfiguracionBoolean("PERMITIR_REGISTRO_SCANNER", true),
      BLOQUEAR_DUPLICADO_DIARIO: obtenerConfiguracionBoolean("BLOQUEAR_DUPLICADO_DIARIO", true),
      EXIGIR_ENTRADA_PREVIA_SALIDA: obtenerConfiguracionBoolean("EXIGIR_ENTRADA_PREVIA_SALIDA", true),
      BLOQUEAR_SALIDA_PERSONAL_INACTIVO: obtenerConfiguracionBoolean("BLOQUEAR_SALIDA_PERSONAL_INACTIVO", true),
      PERMITIR_ENTRADA_PERSONAL_INACTIVO: obtenerConfiguracionBoolean("PERMITIR_ENTRADA_PERSONAL_INACTIVO", false),
      MOSTRAR_FOTO_PERSONAL: obtenerConfiguracionBoolean("MOSTRAR_FOTO_PERSONAL", true),
      MOSTRAR_QR_REGISTRO_ALUMNO: obtenerConfiguracionBoolean("MOSTRAR_QR_REGISTRO_ALUMNO", true),
      URL_FORMULARIO_REGISTRO_ALUMNO: obtenerConfiguracion("URL_FORMULARIO_REGISTRO_ALUMNO", ""),
      NOMBRE_INSTITUCION: obtenerConfiguracion("NOMBRE_INSTITUCION", ""),
      LOGO_INSTITUCIONAL: obtenerConfiguracion("LOGO_INSTITUCIONAL", ""),
      IMAGEN_FONDO_ASISTENCIA: obtenerConfiguracion("IMAGEN_FONDO_ASISTENCIA", "")
    }
  };
}

function obtenerRecursosPublicos() {
  const cache = CacheService.getScriptCache();
  const g = cache.get("RECURSOS_PUB_V1");
  if (g) return JSON.parse(g);
  const convertirImagen = id => {
    if (!id) return "";

    try {
      const archivo = DriveApp.getFileById(id);
      const blob = archivo.getBlob();
      return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      return "";
    }
  };

  const r = {
    estado: "OK",
    logoInstitucional: convertirImagen(obtenerConfiguracion("LOGO_INSTITUCIONAL", "")),
    imagenFondoAsistencia: convertirImagen(obtenerConfiguracion("IMAGEN_FONDO_ASISTENCIA", ""))
  };
  try {
    const texto = JSON.stringify(r);
    if (texto.length < 90000) cache.put("RECURSOS_PUB_V1", texto, 600);
  } catch (e) { }
  return r;
}

function guardarConfiguracionesSistema(datos) {
  const autorizacion = validarAutorizacion(datos.token, "CONFIGURACION_EDITAR");

  if (!autorizacion.autorizado) return autorizacion;

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIGURACION");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "CONFIGURACION".'
    };
  }

  const configuraciones = datos.configuraciones || {};
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      estado: "ERROR",
      mensaje: "La hoja CONFIGURACION no contiene parámetros."
    };
  }

  const rango = hoja.getRange(2, 1, ultimaFila - 1, 2);
  const valores = rango.getValues();

  const permitidas = [
    "TIEMPO_MINIMO_SALIDA",
    "HORA_INGRESO_MANANA",
    "HORA_CIERRE_MANANA",
    "HORA_INGRESO_TARDE",
    "HORA_CIERRE_TARDE",
    "HORA_INGRESO_NOCHE",
    "HORA_CIERRE_NOCHE",
    "PERMITIR_REGISTRO_MANUAL",
    "PERMITIR_REGISTRO_SCANNER",
    "BLOQUEAR_DUPLICADO_DIARIO",
    "EXIGIR_ENTRADA_PREVIA_SALIDA",
    "BLOQUEAR_SALIDA_PERSONAL_INACTIVO",
    "PERMITIR_ENTRADA_PERSONAL_INACTIVO",
    "MOSTRAR_FOTO_PERSONAL",
    "MOSTRAR_QR_REGISTRO_ALUMNO",
    "URL_FORMULARIO_REGISTRO_ALUMNO",
    "NOMBRE_INSTITUCION",
    "LOGO_INSTITUCIONAL",
    "IMAGEN_FONDO_ASISTENCIA"
  ];

  const booleanas = new Set([
    "PERMITIR_REGISTRO_MANUAL",
    "PERMITIR_REGISTRO_SCANNER",
    "BLOQUEAR_DUPLICADO_DIARIO",
    "EXIGIR_ENTRADA_PREVIA_SALIDA",
    "BLOQUEAR_SALIDA_PERSONAL_INACTIVO",
    "PERMITIR_ENTRADA_PERSONAL_INACTIVO",
    "MOSTRAR_FOTO_PERSONAL",
    "MOSTRAR_QR_REGISTRO_ALUMNO"
  ]);

  const nuevosValores = {};

  for (const clave of permitidas) {
    if (!(clave in configuraciones)) continue;

    let valor = configuraciones[clave];

    if (clave === "TIEMPO_MINIMO_SALIDA") {
      valor = Number(valor);

      if (!Number.isFinite(valor) || valor < 0) {
        return {
          estado: "ERROR",
          mensaje: "El tiempo mínimo de salida debe ser un número igual o mayor que 0."
        };
      }

      valor = Math.floor(valor);
    }

    const horasTurno = new Set(["HORA_INGRESO_MANANA", "HORA_CIERRE_MANANA", "HORA_INGRESO_TARDE", "HORA_CIERRE_TARDE", "HORA_INGRESO_NOCHE", "HORA_CIERRE_NOCHE"]);
    if (horasTurno.has(clave)) {
      valor = String(valor || "").trim();
      if (!/^\d{2}:\d{2}$/.test(valor)) return { estado: "ERROR", mensaje: "Hora no válida en " + clave + ". Use HH:MM." };
    }

    if (booleanas.has(clave)) {
      valor = valor === true ? "SI" : "NO";
    }

    if (clave === "URL_FORMULARIO_REGISTRO_ALUMNO") {
      valor = String(valor || "").trim();

      if (valor && !/^https?:\/\/.+/i.test(valor)) {
        return {
          estado: "ERROR",
          mensaje: "La URL del formulario no es válida."
        };
      }
    }

    if (clave === "NOMBRE_INSTITUCION") {
      valor = String(valor || "").trim();
    }

    if (clave === "LOGO_INSTITUCIONAL" || clave === "IMAGEN_FONDO_ASISTENCIA") {
      valor = String(valor || "").trim();
    }

    nuevosValores[clave] = valor;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    for (let i = 0; i < valores.length; i++) {
      const clave = String(valores[i][0]).trim();

      if (clave in nuevosValores) {
        valores[i][1] = nuevosValores[clave];
      }
    }

    rango.setValues(valores);

    limpiarCacheMaestros_("TURNOS");

    return {
      estado: "GUARDADO",
      mensaje: "Configuración guardada correctamente."
    };
  } finally {
    lock.releaseLock();
  }
}

function guardarRecursoInstitucional(datos) {
  const autorizacion = validarAutorizacion(datos.token, "CONFIGURACION_EDITAR");

  if (!autorizacion.autorizado) return autorizacion;

  const tipo = String(datos.tipo || "").trim().toUpperCase();
  const nombre = String(datos.nombre || "").trim();
  const contenido = String(datos.contenido || "").trim();
  const mimeType = String(datos.mimeType || "").trim();

  const tiposPermitidos = {
    LOGO_INSTITUCIONAL: "LOGO_INSTITUCIONAL",
    IMAGEN_FONDO_ASISTENCIA: "IMAGEN_FONDO_ASISTENCIA"
  };

  if (!tiposPermitidos[tipo]) {
    return {
      estado: "ERROR",
      mensaje: "Tipo de recurso institucional no válido."
    };
  }

  if (!contenido || !mimeType) {
    return {
      estado: "ERROR",
      mensaje: "No se recibió la imagen."
    };
  }

  if (!/^image\/(jpeg|png|webp)$/i.test(mimeType)) {
    return {
      estado: "ERROR",
      mensaje: "Solo se permiten imágenes JPG, PNG o WEBP."
    };
  }

  const partes = contenido.split(",");

  if (partes.length !== 2) {
    return {
      estado: "ERROR",
      mensaje: "El formato de la imagen no es válido."
    };
  }

  const bytes = Utilities.base64Decode(partes[1]);

  if (bytes.length > 5 * 1024 * 1024) {
    return {
      estado: "ERROR",
      mensaje: "La imagen no puede superar los 5 MB."
    };
  }

  const ID_CARPETA_RECURSOS = obtenerConfiguracion("ID_CARPETA_RECURSOS", "1bD76_GGBY-RVlP0sKpK-2lTU0V9KpS1F");
  const carpeta = DriveApp.getFolderById(ID_CARPETA_RECURSOS);

  const extension = mimeType.split("/")[1].toLowerCase();
  const nombreArchivo = tipo + "." + extension;

  const archivos = carpeta.getFilesByName(nombreArchivo);

  while (archivos.hasNext()) {
    archivos.next().setTrashed(true);
  }

  const blob = Utilities.newBlob(bytes, mimeType, nombreArchivo);
  const archivo = carpeta.createFile(blob);

  CacheService.getScriptCache().remove("RECURSOS_PUB_V1");

  return {
    estado: "GUARDADO",
    id: archivo.getId(),
    nombre: archivo.getName()
  };
}

function obtenerFotoPersonal(dni) {
  const ID_CARPETA_FOTOS = obtenerConfiguracion("ID_CARPETA_FOTOS", "1UsMmVCAsC9ewAokf_cMHpm6J57qlroIy");
  const carpeta = DriveApp.getFolderById(ID_CARPETA_FOTOS);

  for (const extension of [".jpg", ".jpeg", ".png"]) {
    const archivos = carpeta.getFilesByName(dni + extension);

    if (archivos.hasNext()) {
      const blob = archivos.next().getBlob();
      return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
    }
  }

  return null;
}

function obtenerFotoPersonalPorDni(token, dni) {
  const autorizacion = validarAutorizacion(token, "PERSONAL_VER");
  if (!autorizacion.autorizado) return autorizacion;
  dni = String(dni || "").trim();
  if (!dni) return { estado: "ERROR", mensaje: "DNI obligatorio." };
  return { estado: "OK", dni: dni, foto: obtenerFotoPersonal(dni) };
}

function generarHashPassword(password, salt) {
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + salt,
    Utilities.Charset.UTF_8
  );

  return hash.map(byte => ("0" + (byte & 0xFF).toString(16)).slice(-2)).join("");
}

function generarIdUsuario(hoja) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return "USR-0001";

  const ids = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues().flat();
  let mayor = 0;

  ids.forEach(id => {
    const texto = String(id).trim();

    if (texto.startsWith("USR-")) {
      const numero = parseInt(texto.substring(4), 10);
      if (!isNaN(numero) && numero > mayor) mayor = numero;
    }
  });

  return "USR-" + String(mayor + 1).padStart(4, "0");
}

function validarLogin(usuario, password) {
  usuario = String(usuario || "").trim().toLowerCase();
  password = String(password || "");

  if (!usuario || !password) return {
    estado: "DATOS_INCOMPLETOS",
    mensaje: "Debe ingresar usuario y contraseña."
  };

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS");
  if (!hoja) throw new Error('No existe la hoja "USUARIOS".');

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) return {
    estado: "NO_VALIDO",
    mensaje: "No existen usuarios registrados."
  };

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 10).getValues();
  let usuarioEncontrado = null;

  for (let i = 0; i < datos.length; i++) {
    const fila = datos[i];

    if (String(fila[1]).trim().toLowerCase() === usuario) {
      usuarioEncontrado = {
        id: fila[0],
        usuario: fila[1],
        nombre: fila[2],
        apellidos: fila[3],
        rol: fila[4],
        cargo: fila[5],
        passwordHash: fila[6],
        salt: fila[7],
        estado: fila[8]
      };
      break;
    }
  }

  if (!usuarioEncontrado) return {
    estado: "NO_VALIDO",
    mensaje: "Usuario o contraseña incorrectos."
  };

  if (String(usuarioEncontrado.estado).trim().toUpperCase() !== "ACTIVO") {
    return {
      estado: "USUARIO_INACTIVO",
      mensaje: "El usuario se encuentra inactivo."
    };
  }

  if (generarHashPassword(password, usuarioEncontrado.salt) !== usuarioEncontrado.passwordHash) {
    return {
      estado: "NO_VALIDO",
      mensaje: "Usuario o contraseña incorrectos."
    };
  }

  const token = crearSesion(usuarioEncontrado);

  return {
    estado: "VALIDO",
    token: token,
    usuario: usuarioEncontrado.usuario,
    nombre: usuarioEncontrado.nombre,
    apellidos: usuarioEncontrado.apellidos,
    rol: usuarioEncontrado.rol,
    cargo: usuarioEncontrado.cargo,
    id: usuarioEncontrado.id
  };
}

function crearSesion(usuario) {
  const token = Utilities.getUuid();
  const propiedades = PropertiesService.getScriptProperties();
  const claveVersion = "VERSION_SESION_" + usuario.id;
  const version = propiedades.getProperty(claveVersion) || "1";

  CacheService.getScriptCache().put(
    "SESION_" + token,
    JSON.stringify({
      usuario: usuario.usuario,
      id: usuario.id,
      rol: usuario.rol,
      cargo: usuario.cargo,
      nombre: usuario.nombre,
      apellidos: usuario.apellidos,
      version: version
    }),
    7200
  );

  return token;
}

function validarSesion(token) {
  token = String(token || "").trim();

  if (!token) {
    return {
      estado: "SESION_INVALIDA",
      mensaje: "No existe una sesión activa."
    };
  }

  const cache = CacheService.getScriptCache();
  const datos = cache.get("SESION_" + token);

  if (!datos) {
    return {
      estado: "SESION_INVALIDA",
      mensaje: "La sesión ha expirado o no es válida."
    };
  }

  const sesion = JSON.parse(datos);
  const versionActual = PropertiesService
    .getScriptProperties()
    .getProperty("VERSION_SESION_" + sesion.id) || "1";

  if (sesion.version !== versionActual) {
    cache.remove("SESION_" + token);

    return {
      estado: "SESION_INVALIDA",
      mensaje: "La sesión ya no es válida. Inicie sesión nuevamente."
    };
  }

  return {
    estado: "SESION_VALIDA",
    usuario: sesion.usuario,
    id: sesion.id,
    rol: sesion.rol,
    cargo: sesion.cargo,
    nombre: sesion.nombre,
    apellidos: sesion.apellidos
  };
}

function obtenerPermisoUsuario(idUsuario, idPermiso, idCargo) {
  const m = obtenerMatricesPermisos_();
  const permisos = m.permisos;
  const permisosCargos = m.cargos;
  const permisosUsuarios = m.usuarios;
  if (!permisos.length && !permisosCargos.length && !permisosUsuarios.length) {
    return { existe: false, habilitado: false, protegido: false };
  }

  let protegido = false;

  for (let i = 0; i < permisos.length; i++) {
    if (String(permisos[i][0]).trim().toUpperCase() === idPermiso) {
      protegido = ["SI", "SÍ"].includes(String(permisos[i][4]).trim().toUpperCase());
      break;
    }
  }

  if (protegido) {
    return {
      existe: true,
      habilitado: true,
      protegido: true
    };
  }

  for (let i = 0; i < permisosUsuarios.length; i++) {
    if (
      String(permisosUsuarios[i][0]).trim() === String(idUsuario).trim() &&
      String(permisosUsuarios[i][1]).trim().toUpperCase() === idPermiso
    ) {
      return {
        existe: true,
        habilitado: ["SI", "SÍ"].includes(String(permisosUsuarios[i][2]).trim().toUpperCase()),
        protegido: false
      };
    }
  }

  for (let i = 0; i < permisosCargos.length; i++) {
    if (
      String(permisosCargos[i][0]).trim() === String(idCargo).trim() &&
      String(permisosCargos[i][1]).trim().toUpperCase() === idPermiso
    ) {
      return {
        existe: true,
        habilitado: ["SI", "SÍ"].includes(String(permisosCargos[i][2]).trim().toUpperCase()),
        protegido: false
      };
    }
  }

  return {
    existe: true,
    habilitado: false,
    protegido: false
  };
}

function validarAutorizacion(token, permisoOroles) {
  const sesion = validarSesion(token);

  if (sesion.estado !== "SESION_VALIDA") {
    return {
      autorizado: false,
      estado: "SESION_INVALIDA",
      mensaje: "La sesión ha expirado o no es válida."
    };
  }

  if (Array.isArray(permisoOroles)) {
    const rol = String(sesion.rol || "").trim().toUpperCase();

    if (!permisoOroles.includes(rol)) {
      return {
        autorizado: false,
        estado: "SIN_PERMISO",
        mensaje: "No tiene permisos para realizar esta operación."
      };
    }

    return {
      autorizado: true,
      sesion: sesion
    };
  }

  const idPermiso = String(permisoOroles || "").trim().toUpperCase();

  if (!idPermiso) {
    return {
      autorizado: false,
      estado: "SIN_PERMISO",
      mensaje: "No se especificó el permiso requerido."
    };
  }

  if (String(sesion.rol || "").trim().toUpperCase() === "ADMINISTRADOR") {
    return {
      autorizado: true,
      sesion: sesion
    };
  }

  const permiso = obtenerPermisoUsuario(
    sesion.id,
    idPermiso,
    sesion.cargo
  );

  if (!permiso.existe) {
    return {
      autorizado: false,
      estado: "SIN_PERMISO",
      mensaje: "El permiso solicitado no existe."
    };
  }

  if (permiso.protegido) {
    return {
      autorizado: false,
      estado: "SIN_PERMISO",
      mensaje: "Este permiso está reservado para el administrador."
    };
  }

  if (!permiso.habilitado) {
    return {
      autorizado: false,
      estado: "SIN_PERMISO",
      mensaje: "No tiene permisos para realizar esta operación."
    };
  }

  return {
    autorizado: true,
    sesion: sesion
  };
}

function cerrarSesion(token) {
  token = String(token || "").trim();

  if (token) {
    CacheService.getScriptCache().remove("SESION_" + token);
  }

  return {
    estado: "SESION_CERRADA"
  };
}

function obtenerUrlAplicacion() {
  return ScriptApp.getService().getUrl();
}

function obtenerCargos(token) {
  const autorizacion = validarAutorizacion(token, "USUARIOS_VER");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGOS");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "CARGOS".'
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      estado: "OK",
      cargos: []
    };
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 4).getDisplayValues();

  return {
    estado: "OK",
    cargos: datos
      .filter(fila =>
        String(fila[0]).trim() &&
        String(fila[1]).trim() &&
        String(fila[3]).trim().toUpperCase() === "ACTIVO"
      )
      .map(fila => ({
        id: String(fila[0]).trim(),
        nombre: String(fila[1]).trim(),
        descripcion: String(fila[2]).trim()
      }))
  };
}

function registrarUsuario(datos) {
  const autorizacion = validarAutorizacion(datos.token, "USUARIOS_CREAR");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  datos = datos || {};

  const usuario = String(datos.usuario || "").trim().toLowerCase();
  const nombre = String(datos.nombre || "").trim();
  const apellidos = String(datos.apellidos || "").trim();
  const rol = String(datos.rol || "").trim().toUpperCase();
  const cargo = String(datos.cargo || "").trim();
  const password = String(datos.password || "");
  const confirmarPassword = String(datos.confirmarPassword || "");

  if (!usuario || !nombre || !apellidos || !rol || !password) {
    return {
      estado: "ERROR",
      mensaje: "Complete todos los campos obligatorios."
    };
  }

  if (!/^[a-z0-9._-]{4,30}$/.test(usuario)) {
    return {
      estado: "ERROR",
      mensaje: "El usuario debe tener entre 4 y 30 caracteres y solo puede contener letras, números, punto, guion o guion bajo."
    };
  }

  if (password.length < 8) {
    return {
      estado: "ERROR",
      mensaje: "La contraseña debe tener como mínimo 8 caracteres."
    };
  }

  if (confirmarPassword && password !== confirmarPassword) {
    return {
      estado: "ERROR",
      mensaje: "Las contraseñas no coinciden."
    };
  }

  if (!["ADMINISTRADOR", "SECRETARIA", "COORDINADOR"].includes(rol)) {
    return {
      estado: "ERROR",
      mensaje: "El rol seleccionado no es válido."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("USUARIOS");
  const hojaCargos = ss.getSheetByName("CARGOS");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "USUARIOS".'
    };
  }

  if (!hojaCargos) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "CARGOS".'
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const ultimaFila = hoja.getLastRow();

    const datosUsuarios = ultimaFila >= 2
      ? hoja.getRange(2, 1, ultimaFila - 1, 10).getValues()
      : [];

    const usuarioExistente = datosUsuarios.some(fila =>
      String(fila[1]).trim().toLowerCase() === usuario
    );

    if (usuarioExistente) {
      return {
        estado: "DUPLICADO",
        mensaje: "El nombre de usuario ya está registrado."
      };
    }

    const datosCargos = hojaCargos.getLastRow() >= 2
      ? hojaCargos.getRange(2, 1, hojaCargos.getLastRow() - 1, 4).getDisplayValues()
      : [];

    let cargoFinal = cargo;

    if (!cargoFinal) {
      const cargoPorRol = datosCargos.find(fila =>
        String(fila[1]).trim().toUpperCase() === rol &&
        String(fila[3]).trim().toUpperCase() === "ACTIVO"
      );

      cargoFinal = cargoPorRol
        ? String(cargoPorRol[0]).trim()
        : "";
    }

    const cargoValido = datosCargos.some(fila =>
      String(fila[0]).trim() === cargoFinal &&
      String(fila[3]).trim().toUpperCase() === "ACTIVO"
    );

    if (!cargoValido) {
      return {
        estado: "ERROR",
        mensaje: "El cargo seleccionado no es válido o está inactivo."
      };
    }

    const id = generarIdUsuario(hoja);
    const salt = Utilities.getUuid();
    const passwordHash = generarHashPassword(password, salt);

    hoja.appendRow([
      id,
      usuario,
      nombre,
      apellidos,
      rol,
      cargoFinal,
      passwordHash,
      salt,
      "ACTIVO",
      new Date()
    ]);

    PropertiesService
      .getScriptProperties()
      .setProperty("VERSION_SESION_" + id, "1");

    return {
      estado: "REGISTRADO",
      mensaje: "Usuario registrado correctamente."
    };
  } finally {
    lock.releaseLock();
  }
}

function obtenerUsuarios(token) {
  const autorizacion = validarAutorizacion(token, "USUARIOS_VER");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje,
      usuarios: []
    };
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "USUARIOS".',
      usuarios: []
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      estado: "OK",
      usuarios: []
    };
  }

  const datos = hoja
    .getRange(2, 1, ultimaFila - 1, 10)
    .getDisplayValues();

  const usuarios = datos
    .filter(fila => String(fila[0]).trim() !== "")
    .map(fila => ({
      id: String(fila[0]).trim(),
      usuario: String(fila[1]).trim(),
      nombre: String(fila[2]).trim(),
      apellidos: String(fila[3]).trim(),
      rol: String(fila[4]).trim().toUpperCase(),
      cargo: String(fila[5]).trim(),
      estado: String(fila[8]).trim().toUpperCase(),
      fechaCreacion: String(fila[9]).trim()
    }));

  return {
    estado: "OK",
    usuarios
  };
}

function actualizarUsuario(datos) {
  const autorizacion = validarAutorizacion(datos.token, "USUARIOS_EDITAR");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  datos = datos || {};

  const id = String(datos.id || "").trim();
  const usuario = String(datos.usuario || "").trim().toLowerCase();
  const nombre = String(datos.nombre || "").trim();
  const apellidos = String(datos.apellidos || "").trim();
  const rol = String(datos.rol || "").trim().toUpperCase();
  const cargo = String(datos.cargo || "").trim();
  const estado = String(datos.estado || "").trim().toUpperCase();

  if (!id || !usuario || !nombre || !apellidos || !rol || !cargo || !estado) {
    return {
      estado: "ERROR",
      mensaje: "Complete todos los campos obligatorios."
    };
  }

  if (!/^[a-z0-9._-]{4,30}$/.test(usuario)) {
    return {
      estado: "ERROR",
      mensaje: "El nombre de usuario no es válido."
    };
  }

  if (!["ADMINISTRADOR", "SECRETARIA", "COORDINADOR"].includes(rol)) {
    return {
      estado: "ERROR",
      mensaje: "El rol seleccionado no es válido."
    };
  }

  if (!["ACTIVO", "INACTIVO"].includes(estado)) {
    return {
      estado: "ERROR",
      mensaje: "El estado seleccionado no es válido."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("USUARIOS");
  const hojaCargos = ss.getSheetByName("CARGOS");

  if (!hoja || !hojaCargos) {
    return {
      estado: "ERROR",
      mensaje: "No se encontraron las hojas necesarias."
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const ultimaFila = hoja.getLastRow();

    if (ultimaFila < 2) {
      return {
        estado: "NO_ENCONTRADO",
        mensaje: "No existen usuarios registrados."
      };
    }

    const datosUsuarios = hoja
      .getRange(2, 1, ultimaFila - 1, 10)
      .getValues();

    let filaUsuario = -1;
    let usuarioActual = null;

    for (let i = 0; i < datosUsuarios.length; i++) {
      if (String(datosUsuarios[i][0]).trim() === id) {
        filaUsuario = i + 2;
        usuarioActual = datosUsuarios[i];
        break;
      }
    }

    if (filaUsuario === -1) {
      return {
        estado: "NO_ENCONTRADO",
        mensaje: "No se encontró el usuario indicado."
      };
    }

    for (let i = 0; i < datosUsuarios.length; i++) {
      if (i + 2 === filaUsuario) continue;

      if (String(datosUsuarios[i][1]).trim().toLowerCase() === usuario) {
        return {
          estado: "DUPLICADO",
          mensaje: "El nombre de usuario ya está registrado."
        };
      }
    }

    const rolActual = String(usuarioActual[4]).trim().toUpperCase();
    const cargoActual = String(usuarioActual[5]).trim();
    const estadoActual = String(usuarioActual[8]).trim().toUpperCase();

    if (
      rolActual === "ADMINISTRADOR" &&
      estadoActual === "ACTIVO" &&
      (rol !== "ADMINISTRADOR" || estado !== "ACTIVO")
    ) {
      const administradoresActivos = datosUsuarios.filter(fila =>
        String(fila[4]).trim().toUpperCase() === "ADMINISTRADOR" &&
        String(fila[8]).trim().toUpperCase() === "ACTIVO" &&
        String(fila[0]).trim() !== id
      ).length;

      if (administradoresActivos === 0) {
        return {
          estado: "ERROR",
          mensaje: "No puede dejar el sistema sin un administrador activo."
        };
      }
    }

    if (
      String(autorizacion.sesion.id).trim() === id &&
      estado !== "ACTIVO"
    ) {
      return {
        estado: "ERROR",
        mensaje: "No puede desactivar su propia cuenta."
      };
    }

    const datosCargos = hojaCargos.getLastRow() >= 2
      ? hojaCargos.getRange(2, 1, hojaCargos.getLastRow() - 1, 4).getDisplayValues()
      : [];

    const cargoValido = datosCargos.some(fila =>
      String(fila[0]).trim() === cargo &&
      String(fila[3]).trim().toUpperCase() === "ACTIVO"
    );

    if (!cargoValido) {
      return {
        estado: "ERROR",
        mensaje: "El cargo seleccionado no es válido o está inactivo."
      };
    }

    hoja.getRange(filaUsuario, 2, 1, 5).setValues([[
      usuario,
      nombre,
      apellidos,
      rol,
      cargo
    ]]);

    hoja.getRange(filaUsuario, 9).setValue(estado);

    if (
      String(usuarioActual[1]).trim().toLowerCase() !== usuario ||
      rolActual !== rol ||
      cargoActual !== cargo ||
      estadoActual !== estado
    ) {
      invalidarSesionesUsuario(id);
    }

    return {
      estado: "ACTUALIZADO",
      mensaje: "Usuario actualizado correctamente."
    };
  } finally {
    lock.releaseLock();
  }
}

function cambiarPasswordUsuario(datos) {
  const autorizacion = validarAutorizacion(datos.token, "USUARIOS_EDITAR");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  datos = datos || {};

  const id = String(datos.id || "").trim();
  const password = String(datos.password || "");
  const confirmarPassword = String(datos.confirmarPassword || "");

  if (!id || !password || !confirmarPassword) {
    return {
      estado: "ERROR",
      mensaje: "Complete todos los campos."
    };
  }

  if (password.length < 8) {
    return {
      estado: "ERROR",
      mensaje: "La contraseña debe tener como mínimo 8 caracteres."
    };
  }

  if (password !== confirmarPassword) {
    return {
      estado: "ERROR",
      mensaje: "Las contraseñas no coinciden."
    };
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "USUARIOS".'
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const ultimaFila = hoja.getLastRow();

    if (ultimaFila < 2) {
      return {
        estado: "NO_ENCONTRADO",
        mensaje: "No existen usuarios registrados."
      };
    }

    const datosUsuarios = hoja
      .getRange(2, 1, ultimaFila - 1, 10)
      .getValues();

    for (let i = 0; i < datosUsuarios.length; i++) {
      if (String(datosUsuarios[i][0]).trim() !== id) continue;

      const filaUsuario = i + 2;
      const salt = Utilities.getUuid();
      const passwordHash = generarHashPassword(password, salt);

      hoja.getRange(filaUsuario, 7, 1, 2).setValues([[
        passwordHash,
        salt
      ]]);

      invalidarSesionesUsuario(id);

      return {
        estado: "ACTUALIZADO",
        mensaje: "Contraseña actualizada correctamente."
      };
    }

    return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró el usuario indicado."
    };
  } finally {
    lock.releaseLock();
  }
}

function invalidarSesionesUsuario(idUsuario) {
  const clave = "VERSION_SESION_" + String(idUsuario).trim();
  const propiedades = PropertiesService.getScriptProperties();
  const versionActual = parseInt(propiedades.getProperty(clave) || "1", 10);

  propiedades.setProperty(
    clave,
    String(isNaN(versionActual) ? 2 : versionActual + 1)
  );
}

function obtenerAlumnos(token) {
  const autorizacion = validarAutorizacion(token, "ALUMNOS_VER");

  if (!autorizacion.autorizado) return {
    estado: autorizacion.estado,
    mensaje: autorizacion.mensaje
  };

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estudiantes");

  if (!hoja) throw new Error('No existe la hoja "Estudiantes".');

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) return [];

  return hoja.getRange(2, 1, ultimaFila - 1, 6)
    .getDisplayValues()
    .filter(fila => fila[0].trim() !== "")
    .map(fila => ({
      dni: fila[0],
      nombre: fila[1],
      apellidos: fila[2],
      carrera: fila[3],
      ciclo: fila[4],
      turno: fila[5]
    }));
}

function obtenerSolicitudesPendientes(token) {
  const autorizacion = validarAutorizacion(token, "ALUMNOS_SOLICITUDES");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje,
      solicitudes: []
    };
  }

  const hoja = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("SOLICITUDES_REGISTRO");

  if (!hoja) {
    throw new Error('No existe la hoja "SOLICITUDES_REGISTRO".');
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      estado: "OK",
      solicitudes: []
    };
  }

  const datos = hoja
    .getRange(2, 1, ultimaFila - 1, 12)
    .getDisplayValues();

  const solicitudes = datos
    .filter(fila => String(fila[8]).trim().toUpperCase() === "PENDIENTE")
    .map(fila => ({
      idSolicitud: fila[0],
      fechaSolicitud: fila[1],
      dni: fila[2],
      nombre: fila[3],
      apellidos: fila[4],
      carrera: fila[5],
      ciclo: fila[6],
      turno: String(fila[7] || "").trim().toUpperCase(),
      estado: fila[8],
      observacion: fila[9],
      fechaValidacion: fila[10],
      validadoPor: fila[11]
    }));

  return {
    estado: "OK",
    solicitudes: solicitudes
  };
}

function resolverSolicitudRegistro(token, datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    datos = datos || {};

    const autorizacion = validarAutorizacion(token, "ALUMNOS_SOLICITUDES");

    if (!autorizacion.autorizado) return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };

    const idSolicitud = String(datos.idSolicitud || "").trim();
    const accion = String(datos.accion || "").trim().toUpperCase();
    const observacion = String(datos.observacion || "").trim();
    const sesion = autorizacion.sesion;

    if (!/^SOL-\d{6}$/.test(idSolicitud)) return {
      estado: "ERROR",
      mensaje: "El identificador de la solicitud no es válido."
    };

    if (!["APROBAR", "RECHAZAR"].includes(accion)) return {
      estado: "ERROR",
      mensaje: "La acción solicitada no es válida."
    };

    if (accion === "RECHAZAR" && !observacion) return {
      estado: "ERROR",
      mensaje: "Debe indicar el motivo del rechazo."
    };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaSolicitudes = ss.getSheetByName("SOLICITUDES_REGISTRO");
    const hojaEstudiantes = ss.getSheetByName("Estudiantes");

    if (!hojaSolicitudes) throw new Error('No existe la hoja "SOLICITUDES_REGISTRO".');
    if (!hojaEstudiantes) throw new Error('No existe la hoja "Estudiantes".');

    const ultimaFila = hojaSolicitudes.getLastRow();

    if (ultimaFila < 2) return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró la solicitud."
    };

    const datosSolicitudes = hojaSolicitudes
      .getRange(2, 1, ultimaFila - 1, 12)
      .getDisplayValues();

    let filaSolicitud = -1;
    let solicitud = null;

    for (let i = 0; i < datosSolicitudes.length; i++) {
      if (String(datosSolicitudes[i][0]).trim() === idSolicitud) {
        filaSolicitud = i + 2;
        solicitud = datosSolicitudes[i];
        break;
      }
    }

    if (filaSolicitud === -1) return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró la solicitud."
    };

    if (String(solicitud[8]).trim().toUpperCase() !== "PENDIENTE") return {
      estado: "SOLICITUD_YA_RESUELTA",
      mensaje: "La solicitud ya fue procesada anteriormente."
    };

    const dni = String(solicitud[2] || "").trim();
    const nombre = capitalizarNombre_(solicitud[3]);
    const apellidos = capitalizarNombre_(solicitud[4]);
    const carrera = String(solicitud[5] || "").trim();
    const ciclo = String(solicitud[6] || "").trim();
    const turno = String(solicitud[7] || "").trim().toUpperCase();

    if (!/^\d{8}$/.test(dni)) return {
      estado: "ERROR",
      mensaje: "El DNI de la solicitud no es válido."
    };

    if (!nombre) return {
      estado: "ERROR",
      mensaje: "El nombre de la solicitud está vacío."
    };

    if (!apellidos) return {
      estado: "ERROR",
      mensaje: "Los apellidos de la solicitud están vacíos."
    };

    if (!carrera) return {
      estado: "ERROR",
      mensaje: "La carrera de la solicitud está vacía."
    };

    if (!ciclo) return {
      estado: "ERROR",
      mensaje: "El ciclo de la solicitud está vacío."
    };

    if (!["MAÑANA", "TARDE", "NOCHE"].includes(turno)) return {
      estado: "ERROR",
      mensaje: "El turno de la solicitud no es válido."
    };

    const carrerasValidas = [
      "Asistencia Administrativa",
      "Enfermeria Tecnica",
      "Tecnica en Farmacia",
      "Computacion e Informatica"
    ];

    const ciclosValidos = ["I", "II", "III", "IV", "V", "VI"];

    if (!carrerasValidas.includes(carrera)) return {
      estado: "ERROR",
      mensaje: "La carrera de la solicitud no es válida."
    };

    if (!ciclosValidos.includes(ciclo)) return {
      estado: "ERROR",
      mensaje: "El ciclo de la solicitud no es válido."
    };

    if (!/^[A-Za-zÁÉÍÓÚÜÑáéíóúüñÀ-ÿ' -]+$/.test(nombre)) return {
      estado: "ERROR",
      mensaje: "El nombre de la solicitud contiene caracteres no válidos."
    };

    if (!/^[A-Za-zÁÉÍÓÚÜÑáéíóúüñÀ-ÿ' -]+$/.test(apellidos)) return {
      estado: "ERROR",
      mensaje: "Los apellidos de la solicitud contienen caracteres no válidos."
    };

    if (accion === "RECHAZAR") {
      const ahora = new Date();

      hojaSolicitudes.getRange(filaSolicitud, 9, 1, 4).setValues([[
        "RECHAZADO",
        observacion,
        ahora,
        sesion.usuario
      ]]);

      return {
        estado: "RECHAZADO",
        mensaje: "Solicitud rechazada correctamente.",
        idSolicitud: idSolicitud
      };
    }

    const ultimaFilaEstudiantes = hojaEstudiantes.getLastRow();

    if (ultimaFilaEstudiantes >= 2) {
      const dnis = hojaEstudiantes
        .getRange(2, 1, ultimaFilaEstudiantes - 1, 1)
        .getDisplayValues()
        .flat()
        .map(valor => normalizarDni_(valor));

      if (dnis.includes(dni)) return {
        estado: "DNI_EXISTENTE",
        mensaje: "El DNI ya se encuentra registrado como alumno."
      };
    }

    hojaEstudiantes.appendRow([
      dni,
      nombre,
      apellidos,
      carrera,
      ciclo,
      turno,
      new Date()
    ]);

    const ahora = new Date();

    hojaSolicitudes.getRange(filaSolicitud, 9, 1, 4).setValues([[
      "APROBADO",
      "",
      ahora,
      sesion.usuario
    ]]);

    limpiarCacheMaestros_("EST");
    subirVersionPadron_();

    return {
      estado: "APROBADO",
      mensaje: "Solicitud aprobada y alumno registrado correctamente.",
      idSolicitud: idSolicitud,
      alumno: {
        dni: dni,
        nombre: nombre,
        apellidos: apellidos,
        carrera: carrera,
        ciclo: ciclo,
        turno: turno
      }
    };

  } finally {
    lock.releaseLock();
  }
}

function obtenerPersonal(token) {
  const autorizacion = validarAutorizacion(token, "PERSONAL_VER");
  if (!autorizacion.autorizado) return autorizacion;

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
  const last = hoja.getLastRow();
  if (last < 2) return { estado: "OK", personal: [] };
  const datos = hoja.getRange(2, 1, last - 1, 6).getValues();
  const personal = datos.map(fila => ({
    dni: String(fila[0]).trim(),
    nombre: String(fila[1]).trim(),
    apellidos: String(fila[2]).trim(),
    tipo: String(fila[3]).trim(),
    cargo: String(fila[4]).trim(),
    estado: String(fila[5]).trim(),
    foto: null
  }));
  return { estado: "OK", personal: personal };
}

function registrarPersonal(datos) {
  const autorizacion = validarAutorizacion(datos.token, "PERSONAL_CREAR");
  if (!autorizacion.autorizado) return autorizacion;

  const dni = String(datos.dni || "").trim();
  const nombre = capitalizarNombre_(datos.nombre);
  const apellidos = capitalizarNombre_(datos.apellidos);
  const tipo = String(datos.tipo || "").trim();
  const cargo = String(datos.cargo || "").trim();

  if (!/^\d{8}$/.test(dni)) {
    return { estado: "ERROR", mensaje: "El DNI debe tener exactamente 8 dígitos." };
  }

  if (!dni || !nombre || !apellidos || !tipo || !cargo) {
    return { estado: "ERROR", mensaje: "Complete todos los campos obligatorios." };
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
  const datosHoja = hoja.getDataRange().getValues();

  const existe = datosHoja.slice(1).some(fila => String(fila[0]).trim() === dni);

  if (existe) {
    return { estado: "ERROR", mensaje: "Ya existe un personal registrado con ese DNI." };
  }

  if (datos.foto && datos.foto.base64) {
    const resultadoFoto = guardarFotoPersonal(dni, datos.foto.base64, datos.foto.tipo);

    if (resultadoFoto.estado !== "OK") {
      return resultadoFoto;
    }
  }

  hoja.appendRow([dni, nombre, apellidos, tipo, cargo, "ACTIVO"]);

  limpiarCacheMaestros_("PER");
  subirVersionPadron_();

  return {
    estado: "OK",
    mensaje: "Personal registrado correctamente.",
    foto: datos.foto && datos.foto.base64 ? true : false
  };
}

function guardarFotoPersonal(dni, base64, tipo) {
  try {
    const ID_CARPETA_FOTOS = obtenerConfiguracion("ID_CARPETA_FOTOS", "1UsMmVCAsC9ewAokf_cMHpm6J57qlroIy");
    const carpeta = DriveApp.getFolderById(ID_CARPETA_FOTOS);

    const contenido = base64.split(",")[1] || base64;
    const bytes = Utilities.base64Decode(contenido);

    let extension = ".jpg";
    if (tipo === "image/png") extension = ".png";
    if (tipo === "image/jpeg") extension = ".jpg";

    const extensiones = [".jpg", ".jpeg", ".png"];

    extensiones.forEach(function (ext) {
      const archivos = carpeta.getFilesByName(dni + ext);
      while (archivos.hasNext()) {
        archivos.next().setTrashed(true);
      }
    });

    const blob = Utilities.newBlob(bytes, tipo || "image/jpeg", dni + extension);
    carpeta.createFile(blob);

    return { estado: "OK" };
  } catch (error) {
    console.error(error);
    return { estado: "ERROR", mensaje: "No se pudo guardar la fotografía del personal." };
  }
}

function actualizarPersonal(datos) {
  const autorizacion = validarAutorizacion(datos.token, "PERSONAL_EDITAR");
  if (!autorizacion.autorizado) return autorizacion;

  const dni = String(datos.dni || "").trim();
  const nombre = capitalizarNombre_(datos.nombre);
  const apellidos = capitalizarNombre_(datos.apellidos);
  const tipo = String(datos.tipo || "").trim();
  const cargo = String(datos.cargo || "").trim();
  const estado = String(datos.estado || "").trim();

  if (!/^\d{8}$/.test(dni)) {
    return { estado: "ERROR", mensaje: "El DNI debe tener exactamente 8 dígitos." };
  }

  if (!dni || !nombre || !apellidos || !tipo || !cargo || !estado) {
    return { estado: "ERROR", mensaje: "Complete todos los campos obligatorios." };
  }

  if (estado !== "ACTIVO" && estado !== "INACTIVO") {
    return { estado: "ERROR", mensaje: "El estado seleccionado no es válido." };
  }

  if (datos.foto && datos.foto.base64) {
    const resultadoFoto = guardarFotoPersonal(dni, datos.foto.base64, datos.foto.tipo);
    if (resultadoFoto.estado !== "OK") return resultadoFoto;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPersonal = ss.getSheetByName("Personal");
  const hojaAsistencias = ss.getSheetByName("Asistencias");

  if (!hojaPersonal || !hojaAsistencias) {
    return { estado: "ERROR", mensaje: "No se encontraron las hojas necesarias." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const lastP = hojaPersonal.getLastRow();
    const datosPersonal = lastP >= 2 ? hojaPersonal.getRange(2, 1, lastP - 1, 6).getValues() : [];

    let filaPersona = -1;
    for (let i = 0; i < datosPersonal.length; i++) {
      if (String(datosPersonal[i][0]).trim() === dni) { filaPersona = i + 2; break; }
    }

    if (filaPersona === -1) {
      return { estado: "NO_ENCONTRADO", mensaje: "No se encontró el personal indicado." };
    }

    hojaPersonal.getRange(filaPersona, 2, 1, 5).setValues([[
      nombre, apellidos, tipo, cargo, estado
    ]]);

    let asistenciaAnulada = false;

    if (estado === "INACTIVO") {
      const lastA = hojaAsistencias.getLastRow();
      const datosAsistencias = lastA >= 2 ? hojaAsistencias.getRange(2, 1, lastA - 1, 8).getValues() : [];
      const ahora = new Date();
      let filaEntradaHoy = -1;
      let tieneSalida = false;

      for (let j = 0; j < datosAsistencias.length; j++) {
        if (String(datosAsistencias[j][0]).trim() !== dni) continue;
        const fecha = datosAsistencias[j][1];
        if (!(fecha instanceof Date) || !esMismaFecha(fecha, ahora)) continue;
        const tipoRegistro = String(datosAsistencias[j][5]).trim().toUpperCase();
        if (tipoRegistro === "ENTRADA" && filaEntradaHoy === -1) filaEntradaHoy = j + 2;
        if (tipoRegistro === "SALIDA") tieneSalida = true;
        if (filaEntradaHoy !== -1 && tieneSalida) break;
      }

      if (filaEntradaHoy !== -1 && !tieneSalida) {
        hojaAsistencias.getRange(filaEntradaHoy, 6).setValue("ANULADA");
        asistenciaAnulada = true;
      }
    }

    limpiarCacheMaestros_("PER");
    subirVersionPadron_();

    return {
      estado: "OK",
      mensaje: asistenciaAnulada
        ? "Personal actualizado correctamente. La entrada pendiente del día fue anulada."
        : "Personal actualizado correctamente.",
      foto: !!(datos.foto && datos.foto.base64),
      asistenciaAnulada: asistenciaAnulada
    };
  } finally {
    lock.releaseLock();
  }
}

function eliminarPersonal(datos) {
  const autorizacion = validarAutorizacion(datos.token, "PERSONAL_ELIMINAR");
  if (!autorizacion.autorizado) return autorizacion;

  const dni = String(datos.dni || "").trim();

  if (!dni) {
    return { estado: "ERROR", mensaje: "No se indicó el DNI del personal." };
  }

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName("Personal");
  const auditoria = libro.getSheetByName("Eliminaciones Personal");

  if (!hoja || !auditoria) {
    return { estado: "ERROR", mensaje: "No se encontraron las hojas necesarias." };
  }

  const datosHoja = hoja.getDataRange().getValues();

  for (let i = 1; i < datosHoja.length; i++) {
    const fila = datosHoja[i];

    if (String(fila[0]).trim() === dni) {
      const usuario = String(autorizacion.sesion.usuario || "").trim();
      const fecha = new Date();

      auditoria.appendRow([
        String(fila[0]).trim(),
        String(fila[1]).trim(),
        String(fila[2]).trim(),
        String(fila[3]).trim(),
        String(fila[4]).trim(),
        String(fila[5]).trim(),
        fecha,
        usuario
      ]);

      hoja.deleteRow(i + 1);

      limpiarCacheMaestros_("PER");
      subirVersionPadron_();

      return {
        estado: "OK",
        mensaje: "Personal eliminado correctamente."
      };
    }
  }

  return {
    estado: "NO_ENCONTRADO",
    mensaje: "No se encontró el personal indicado."
  };
}

function registrarAlumno(datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    datos = datos || {};

    const token = String(datos.token || "").trim();
    const autorizacion = validarAutorizacion(token, "ALUMNOS_CREAR");

    if (!autorizacion.autorizado) return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };

    const dni = String(datos.dni || "").trim();
    const nombre = capitalizarNombre_(datos.nombre);
    const apellidos = capitalizarNombre_(datos.apellidos);
    const carrera = String(datos.carrera || "").trim();
    const ciclo = String(datos.ciclo || "").trim().toUpperCase();
    const turno = String(datos.turno || "").trim().toUpperCase();

    if (!/^\d{8}$/.test(dni)) return {
      estado: "ERROR",
      mensaje: "El DNI debe tener exactamente 8 dígitos."
    };

    if (!nombre) return { estado: "ERROR", mensaje: "El nombre es obligatorio." };
    if (!apellidos) return { estado: "ERROR", mensaje: "Los apellidos son obligatorios." };
    if (!carrera) return { estado: "ERROR", mensaje: "Debe seleccionar una carrera." };
    if (!ciclo) return { estado: "ERROR", mensaje: "Debe seleccionar un ciclo." };

    if (!["I", "II", "III", "IV", "V", "VI"].includes(ciclo)) {
      return {
        estado: "ERROR",
        mensaje: "El ciclo seleccionado no es válido."
      };
    }

    if (!["MAÑANA", "TARDE", "NOCHE"].includes(turno)) {
      return {
        estado: "ERROR",
        mensaje: "El turno seleccionado no es válido."
      };
    }

    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estudiantes");

    if (!hoja) throw new Error('No existe la hoja "Estudiantes".');

    const ultimaFila = hoja.getLastRow();

    if (ultimaFila >= 2) {
      const dnis = hoja
        .getRange(2, 1, ultimaFila - 1, 1)
        .getDisplayValues();

      if (dnis.some(fila => String(fila[0]).trim() === dni)) {
        return {
          estado: "DNI_EXISTENTE",
          mensaje: "Ya existe un alumno registrado con ese DNI."
        };
      }
    }

    hoja.appendRow([
      dni,
      nombre,
      apellidos,
      carrera,
      ciclo,
      turno,
      new Date()
    ]);

    limpiarCacheMaestros_("EST");
    subirVersionPadron_();

    return {
      estado: "REGISTRADO",
      mensaje: "Alumno registrado correctamente.",
      alumno: {
        dni,
        nombre,
        apellidos,
        carrera,
        ciclo,
        turno
      }
    };

  } finally {
    lock.releaseLock();
  }
}

function importarAlumnos(datos) {
  datos = datos || {};

  const token = String(datos.token || "").trim();
  const autorizacion = validarAutorizacion(token, "ALUMNOS_IMPORTAR");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  const registros = Array.isArray(datos.registros) ? datos.registros : [];

  if (!registros.length) {
    return {
      estado: "ERROR",
      mensaje: "No hay registros para importar."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Estudiantes");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "Estudiantes".'
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ultimaFila = hoja.getLastRow();

    const datosExistentes = ultimaFila >= 2
      ? hoja.getRange(2, 1, ultimaFila - 1, 5).getDisplayValues()
      : [];

    const dnisExistentes = new Set(
      datosExistentes
        .map(fila => String(fila[0] || "").replace(/\D/g, ""))
        .filter(Boolean)
    );

    const carrerasValidas = new Set([
      "Asistencia Administrativa",
      "Secretariado Ejecutivo",
      "Enfermeria Tecnica",
      "Tecnica en Farmacia",
      "Farmacia Tecnica",
      "Computacion e Informatica"
    ]);

    const ciclosValidos = new Set([
      "I",
      "II",
      "III",
      "IV",
      "V",
      "VI"
    ]);

    const dnisLote = new Set();
    const filasValidas = [];
    const conflictos = [];

    registros.forEach(registro => {
      const filaExcel = Number(registro.fila) || 0;
      const dni = normalizarDni_(registro.dni, false);
      const nombre = capitalizarNombre_(registro.nombre);
      const apellidos = capitalizarNombre_(registro.apellidos);
      const carrera = String(registro.carrera || "").trim();
      const ciclo = String(registro.ciclo || "").trim().toUpperCase();
      const turno = String(registro.turno || "").trim().toUpperCase();

      const errores = [];

      if (!/^\d{8}$/.test(dni)) {
        errores.push("DNI inválido");
      }

      if (!nombre) {
        errores.push("Nombres obligatorios");
      } else if (!/^[A-Za-zÁÉÍÓÚÜÑáéíóúüñÀ-ÿ' -]+$/.test(nombre)) {
        errores.push("Nombres contienen caracteres no válidos");
      }

      if (!apellidos) {
        errores.push("Apellidos obligatorios");
      } else if (!/^[A-Za-zÁÉÍÓÚÜÑáéíóúüñÀ-ÿ' -]+$/.test(apellidos)) {
        errores.push("Apellidos contienen caracteres no válidos");
      }

      if (!carrera) {
        errores.push("Carrera obligatoria");
      } else if (!carrerasValidas.has(carrera)) {
        errores.push("Carrera no válida");
      }

      if (!ciclo) {
        errores.push("Ciclo obligatorio");
      } else if (!ciclosValidos.has(ciclo)) {
        errores.push("Ciclo no válido");
      }

      if (!turno) {
        errores.push("Turno obligatorio");
      } else if (!["MAÑANA", "TARDE", "NOCHE"].includes(turno)) {
        errores.push("Turno no válido");
      }

      if (dni && dnisExistentes.has(dni)) {
        errores.push("DNI ya registrado");
      }

      if (dni && dnisLote.has(dni)) {
        errores.push("DNI duplicado en el lote");
      }

      if (dni) {
        dnisLote.add(dni);
      }

      if (errores.length) {
        conflictos.push({
          fila: filaExcel,
          dni,
          nombre,
          apellidos,
          carrera,
          ciclo,
          turno,
          errores
        });
        return;
      }

      const hoy = new Date();
      filasValidas.push([
        dni,
        nombre,
        apellidos,
        carrera,
        ciclo,
        turno,
        hoy
      ]);
    });

    if (filasValidas.length) {
      const filaInicio = hoja.getLastRow() + 1;

      hoja
        .getRange(filaInicio, 1, filasValidas.length, 7)
        .setValues(filasValidas);
    }

    limpiarCacheMaestros_("EST");
    subirVersionPadron_();

    return {
      estado: "IMPORTADO",
      mensaje: filasValidas.length
        ? "La importación se completó correctamente."
        : "No se importó ningún alumno.",
      importados: filasValidas.length,
      omitidos: conflictos.length,
      conflictos
    };

  } finally {
    lock.releaseLock();
  }
}

function actualizarAlumno(datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    datos = datos || {};

    const token = String(datos.token || "").trim();
    const autorizacion = validarAutorizacion(token, "ALUMNOS_EDITAR");

    if (!autorizacion.autorizado) return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };

    const dni = String(datos.dni || "").trim();
    const nombre = capitalizarNombre_(datos.nombre);
    const apellidos = capitalizarNombre_(datos.apellidos);
    const carrera = String(datos.carrera || "").trim();
    const ciclo = String(datos.ciclo || "").trim().toUpperCase();
    const turno = String(datos.turno || "").trim().toUpperCase();

    if (!/^\d{8}$/.test(dni)) return {
      estado: "ERROR",
      mensaje: "El DNI debe tener exactamente 8 dígitos."
    };

    if (!nombre) return {
      estado: "ERROR",
      mensaje: "El nombre es obligatorio."
    };

    if (!apellidos) return {
      estado: "ERROR",
      mensaje: "Los apellidos son obligatorios."
    };

    if (!carrera) return {
      estado: "ERROR",
      mensaje: "Debe seleccionar una carrera."
    };

    if (!ciclo) return {
      estado: "ERROR",
      mensaje: "Debe seleccionar un ciclo."
    };

    if (!["I", "II", "III", "IV", "V", "VI"].includes(ciclo)) {
      return {
        estado: "ERROR",
        mensaje: "El ciclo seleccionado no es válido."
      };
    }

    if (!["MAÑANA", "TARDE", "NOCHE"].includes(turno)) {
      return {
        estado: "ERROR",
        mensaje: "El turno seleccionado no es válido."
      };
    }

    const hoja = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Estudiantes");

    if (!hoja) throw new Error('No existe la hoja "Estudiantes".');

    const ultimaFila = hoja.getLastRow();

    if (ultimaFila < 2) return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró el alumno."
    };

    const dnis = hoja
      .getRange(2, 1, ultimaFila - 1, 1)
      .getDisplayValues();

    let filaAlumno = -1;

    for (let i = 0; i < dnis.length; i++) {
      if (String(dnis[i][0]).trim() === dni) {
        filaAlumno = i + 2;
        break;
      }
    }

    if (filaAlumno === -1) return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró el alumno."
    };

    hoja.getRange(filaAlumno, 2, 1, 5).setValues([[
      nombre,
      apellidos,
      carrera,
      ciclo,
      turno
    ]]);

    limpiarCacheMaestros_("EST");
    subirVersionPadron_();

    return {
      estado: "ACTUALIZADO",
      mensaje: "Alumno actualizado correctamente.",
      alumno: {
        dni,
        nombre,
        apellidos,
        carrera,
        ciclo,
        turno
      }
    };

  } finally {
    lock.releaseLock();
  }
}

function eliminarAlumno(datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    datos = datos || {};

    const dni = String(datos.dni || "").trim();
    const motivo = String(datos.motivo || "").trim();
    const token = String(datos.token || "").trim();

    if (!/^\d{8}$/.test(dni)) return {
      estado: "ERROR",
      mensaje: "El DNI no es válido."
    };

    if (!motivo) return {
      estado: "ERROR",
      mensaje: "Debe indicar el motivo de la eliminación."
    };

    const autorizacion = validarAutorizacion(token, "ALUMNOS_ELIMINAR");

    if (!autorizacion.autorizado) return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };

    const sesion = autorizacion.sesion;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Estudiantes");
    const historial = ss.getSheetByName("Eliminaciones Alumnos");

    if (!hoja) throw new Error('No existe la hoja "Estudiantes".');
    if (!historial) throw new Error('No existe la hoja "Eliminaciones Alumnos".');

    const ultimaFila = hoja.getLastRow();

    if (ultimaFila < 2) return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró el alumno."
    };

    const datosAlumnos = hoja.getRange(2, 1, ultimaFila - 1, 6).getDisplayValues();
    let filaAlumno = -1;
    let alumno = null;

    for (let i = 0; i < datosAlumnos.length; i++) {
      if (String(datosAlumnos[i][0]).trim() === dni) {
        filaAlumno = i + 2;
        alumno = {
          dni: datosAlumnos[i][0],
          nombre: datosAlumnos[i][1],
          apellidos: datosAlumnos[i][2],
          carrera: datosAlumnos[i][3],
          ciclo: datosAlumnos[i][4],
          turno: datosAlumnos[i][5]
        };
        break;
      }
    }

    if (filaAlumno === -1) return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró el alumno."
    };

    const ahora = new Date();

    historial.appendRow([
      alumno.dni,
      alumno.nombre,
      alumno.apellidos,
      alumno.carrera,
      alumno.ciclo,
      alumno.turno,
      motivo,
      sesion.usuario,
      sesion.rol,
      formatearFecha(ahora, "dd/MM/yyyy"),
      formatearFecha(ahora, "hh:mm:ss a")
    ]);

    hoja.deleteRow(filaAlumno);
    limpiarCacheMaestros_("EST");
    subirVersionPadron_();

    return {
      estado: "ELIMINADO",
      mensaje: "Alumno eliminado correctamente."
    };

  } finally {
    lock.releaseLock();
  }
}

function obtenerReporteAsistencias(token, filtros) {
  const autorizacion = validarAutorizacion(token, "REPORTES_VER");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  filtros = filtros || {};

  const dniFiltro = normalizarDni_(filtros.dni, false)
  const nombreFiltro = String(filtros.nombre || "").trim();
  const fechaInicio = String(filtros.fechaInicio || "").trim();
  const fechaFin = String(filtros.fechaFin || "").trim();
  const carreraFiltro = String(filtros.carrera || "").trim();
  const cicloFiltro = String(filtros.ciclo || "").trim();
  const metodoFiltro = String(filtros.metodo || "").trim().toUpperCase();
  const tipoRegistroFiltro = String(filtros.tipoRegistro || "").trim().toUpperCase();

  if (fechaInicio && fechaFin && fechaInicio > fechaFin) {
    return {
      estado: "ERROR",
      mensaje: "La fecha desde no puede ser posterior a la fecha hasta."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Asistencias");
  const hojaEstudiantes = ss.getSheetByName("Estudiantes");
  const hojaPersonal = ss.getSheetByName("Personal");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "Asistencias".'
    };
  }

  if (!hojaEstudiantes) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "Estudiantes".'
    };
  }

  if (!hojaPersonal) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "Personal".'
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      estado: "OK",
      asistencias: []
    };
  }

  const mapaPersonas = new Map();
  const mapaEst = obtenerMapaEstudiantes_();
  Object.keys(mapaEst).forEach(dni => {
    const k = normalizarDni_(dni);
    if (!k) return;
    const e = mapaEst[dni];
    mapaPersonas.set(k, {
      nombre: `${String(e.nombre || "").trim()} ${String(e.apellidos || "").trim()}`.trim()
    });
  });

  const mapaPer = obtenerMapaPersonal_();
  Object.keys(mapaPer).forEach(dni => {
    const k = normalizarDni_(dni);
    if (!k) return;
    const p = mapaPer[dni];
    mapaPersonas.set(k, {
      nombre: `${String(p.nombre || "").trim()} ${String(p.apellidos || "").trim()}`.trim()
    });
  });

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 8).getValues();

  const zonaHoraria = ss.getSpreadsheetTimeZone();

  const normalizarTexto = valor =>
    String(valor || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim()
      .toUpperCase();

  const obtenerFechaClave = valor => {
    if (valor instanceof Date) {
      return Utilities.formatDate(
        valor,
        zonaHoraria,
        "yyyy-MM-dd"
      );
    }

    const texto = String(valor || "").trim();

    if (!texto) return "";

    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
      return texto;
    }

    const partes = texto.split("/");

    if (partes.length === 3) {
      return `${partes[2]}-${partes[1].padStart(2, "0")}-${partes[0].padStart(2, "0")}`;
    }

    return "";
  };

  const formatearFechaRep_ = valor => {
    if (!(valor instanceof Date)) return String(valor || "").trim();
    return Utilities.formatDate(valor, zonaHoraria, "dd/MM/yyyy");
  };
  const formatearHoraRep_ = valor => {
    if (!(valor instanceof Date)) return String(valor || "").trim();
    return Utilities.formatDate(valor, zonaHoraria, "hh:mm:ss a");
  };

  const nombreBuscado = normalizarTexto(nombreFiltro);

  const asistencias = datos
    .filter(fila => {
      const dni = normalizarDni_(fila[0]);
      const fechaRegistro = obtenerFechaClave(fila[1]);
      const metodoRegistro = String(fila[3] || "").trim().toUpperCase();
      const carreraRegistro = String(fila[6] || "").trim();
      const cicloRegistro = String(fila[7] || "").trim().toUpperCase();

      if (!dni) return false;

      if (dniFiltro && dni !== dniFiltro) return false;

      if (fechaInicio && (!fechaRegistro || fechaRegistro < fechaInicio)) {
        return false;
      }

      if (fechaFin && (!fechaRegistro || fechaRegistro > fechaFin)) {
        return false;
      }

      if (carreraFiltro && carreraRegistro !== carreraFiltro) {
        return false;
      }

      if (cicloFiltro && cicloRegistro !== cicloFiltro) {
        return false;
      }

      if (metodoFiltro && metodoRegistro !== metodoFiltro) {
        return false;
      }

      const tipoRegistro = String(fila[5] || "").trim().toUpperCase();

      if (tipoRegistroFiltro && tipoRegistro !== tipoRegistroFiltro) {
        return false;
      }

      if (nombreBuscado) {
        const persona = mapaPersonas.get(dni);

        if (!persona) return false;

        if (!normalizarTexto(persona.nombre).includes(nombreBuscado)) {
          return false;
        }
      }

      return true;
    })
    .map(fila => {
      const dni = normalizarDni_(fila[0]);
      const persona = mapaPersonas.get(dni);

      return {
        dni,
        nombre: persona ? persona.nombre : "",
        fecha: formatearFechaRep_(fila[1]),
        hora: formatearHoraRep_(fila[2]),
        metodo: String(fila[3] || "").trim().toUpperCase(),
        tipoPersona: String(fila[4] || "").trim().toUpperCase(),
        tipoRegistro: String(fila[5] || "").trim().toUpperCase(),
        carrera: String(fila[6] || "").trim(),
        ciclo: String(fila[7] || "").trim()
      };
    });

  return {
    estado: "OK",
    asistencias
  };
}

function obtenerReporteAusentes(token, filtros) {
  const aut = validarAutorizacion(token, "REPORTES_VER");
  if (!aut.autorizado) return aut;
  filtros = filtros || {};
  const fFecha = String(filtros.fecha || "").trim();
  const fTurno = String(filtros.turno || "").trim().toUpperCase();
  const fCarrera = String(filtros.carrera || "").trim();
  const fCiclo = String(filtros.ciclo || "").trim().toUpperCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaC = ss.getSheetByName("Control de Asistencia");
  const hojaE = ss.getSheetByName("Estudiantes");
  const hojaA = ss.getSheetByName("Asistencias");
  if (!hojaC || !hojaE || !hojaA) return { estado: "ERROR", mensaje: "Faltan hojas." };

  const mapEst = new Map();
  if (hojaE.getLastRow() >= 2) {
    hojaE.getRange(2, 1, hojaE.getLastRow() - 1, 6).getDisplayValues().forEach(r => {
      mapEst.set(normalizarDni_(r[0]), {
        nombre: String(r[1]).trim() + " " + String(r[2]).trim(),
        carrera: String(r[3]).trim(),
        ciclo: String(r[4]).trim().toUpperCase()
      });
    });
  }
  const mapHora = new Map();
  if (hojaA.getLastRow() >= 2) {
    hojaA.getRange(2, 1, hojaA.getLastRow() - 1, 8).getValues().forEach(r => {
      if (String(r[5] || "").trim().toUpperCase() !== "ENTRADA") return;
      if (!(r[1] instanceof Date) || !(r[2] instanceof Date)) return;
      const k = normalizarDni_(r[0]) + "|" + Utilities.formatDate(r[1], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      if (!mapHora.has(k)) mapHora.set(k, Utilities.formatDate(r[2], ss.getSpreadsheetTimeZone(), "hh:mm:ss a"));
    });
  }
  if (hojaC.getLastRow() < 2) return { estado: "OK", ausentes: [] };
  const datos = hojaC.getRange(2, 1, hojaC.getLastRow() - 1, 12).getValues();
  const zh = ss.getSpreadsheetTimeZone();
  const ausentes = datos.filter(r => {
    if (!(r[1] instanceof Date)) return false;
    const fecha = Utilities.formatDate(r[1], zh, "dd/MM/yyyy");
    const turno = String(r[4] || "").trim().toUpperCase();
    const estado = String(r[5] || "").trim().toUpperCase();
    if (fFecha && fecha !== fFecha) return false;
    if (fTurno && turno !== fTurno) return false;
    if (estado !== "ESPERANDO" && estado !== "FALTÓ") return false;
    const est = mapEst.get(String(r[2]).trim());
    if (fCarrera && (!est || est.carrera !== fCarrera)) return false;
    if (fCiclo && (!est || est.ciclo !== fCiclo)) return false;
    return true;
  }).map(r => {
    const dni = String(r[2]).trim();
    const fecha = Utilities.formatDate(r[1], zh, "dd/MM/yyyy");
    const est = mapEst.get(dni) || { nombre: "", carrera: "", ciclo: "" };
    return {
      dni: dni,
      nombre: est.nombre,
      carrera: est.carrera,
      ciclo: est.ciclo,
      turno: String(r[4] || "").trim().toUpperCase(),
      fecha: fecha,
      estado: String(r[5] || "").trim().toUpperCase(),
      horaEntrada: mapHora.get(dni + "|" + fecha) || "",
      motivo: String(r[6] || "").trim()
    };
  });
  return { estado: "OK", ausentes: ausentes };
}

function autorizarExportacionReporte(token) {
  const autorizacion = validarAutorizacion(token, "REPORTES_EXPORTAR");

  return {
    autorizado: autorizacion.autorizado,
    estado: autorizacion.estado,
    mensaje: autorizacion.mensaje
  };
}

function obtenerDashboard(token, filtros) {
  const autorizacion = validarAutorizacion(token, "REPORTES_VER");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAsistencias = libro.getSheetByName("Asistencias");
  const hojaEstudiantes = libro.getSheetByName("Estudiantes");
  const hojaPersonal = libro.getSheetByName("Personal");

  if (!hojaAsistencias || !hojaEstudiantes || !hojaPersonal) {
    return {
      estado: "ERROR",
      mensaje: "No se encontraron todas las hojas necesarias para el dashboard."
    };
  }

  const zonaHoraria = libro.getSpreadsheetTimeZone();
  const ahora = new Date();
  const hoy = Utilities.formatDate(ahora, zonaHoraria, "dd/MM/yyyy");

  filtros = filtros || {};
  const tipoPeriodo = String(filtros.tipo || "DIA").trim().toUpperCase();

  const datosAsistencias = hojaAsistencias.getLastRow() < 2
    ? []
    : hojaAsistencias.getRange(2, 1, hojaAsistencias.getLastRow() - 1, 8).getDisplayValues();

  const mapaEstDash = obtenerMapaEstudiantes_();
  const mapaPerDash = obtenerMapaPersonal_();

  const normalizarFechaCorta = texto => {
    const p = String(texto || "").trim().split("/");
    if (p.length !== 3) return "";
    return `${p[0].padStart(2, "0")}/${p[1].padStart(2, "0")}/${p[2]}`;
  };
  const registrosHoy = datosAsistencias.filter(fila =>
    normalizarFechaCorta(fila[1]) === hoy
  );

  const estudiantesRegistrados = Object.keys(mapaEstDash).filter(dni =>
    String(dni).trim()
  ).length;
  const personalActivo = Object.keys(mapaPerDash).filter(dni =>
    String(dni).trim() &&
    String((mapaPerDash[dni] || {}).estado || "").trim().toUpperCase() === "ACTIVO"
  ).length;

  const entradasPersonalHoy = registrosHoy.filter(fila =>
    String(fila[4]).trim().toUpperCase() !== "ESTUDIANTE" &&
    String(fila[5]).trim().toUpperCase() === "ENTRADA"
  ).length;

  const salidasPersonalHoy = registrosHoy.filter(fila =>
    String(fila[4]).trim().toUpperCase() !== "ESTUDIANTE" &&
    String(fila[5]).trim().toUpperCase() === "SALIDA"
  ).length;

  function convertirFecha(fechaTexto) {
    const partes = String(fechaTexto || "").trim().split("/");

    if (partes.length !== 3) return null;

    const dia = Number(partes[0]);
    const mes = Number(partes[1]);
    const anio = Number(partes[2]);

    if (!dia || !mes || !anio) return null;

    const fecha = new Date(anio, mes - 1, dia);
    fecha.setHours(0, 0, 0, 0);

    return fecha;
  }

  function formatearFechaDash_(fecha) {
    return Utilities.formatDate(fecha, zonaHoraria, "dd/MM/yyyy");
  }
  function formatearFechaCortaDash_(fecha) {
    return Utilities.formatDate(fecha, zonaHoraria, "dd/MM");
  }

  function inicioSemana(fecha) {
    const resultado = new Date(fecha);
    const dia = resultado.getDay();
    const diferencia = dia === 0 ? 6 : dia - 1;
    resultado.setDate(resultado.getDate() - diferencia);
    resultado.setHours(0, 0, 0, 0);
    return resultado;
  }

  function sumarDias(fecha, cantidad) {
    const resultado = new Date(fecha);
    resultado.setDate(resultado.getDate() + cantidad);
    return resultado;
  }

  function obtenerRangoPeriodo() {
    let inicio;
    let fin;

    if (tipoPeriodo === "PERSONALIZADO") {
      inicio = convertirFecha(filtros.fechaInicio);
      fin = convertirFecha(filtros.fechaFin);

      if (!inicio || !fin || inicio > fin) {
        return null;
      }

      return { inicio, fin };
    }

    if (tipoPeriodo === "SEMANA") {
      fin = inicioSemana(ahora);
      inicio = sumarDias(fin, -49);
      return { inicio, fin };
    }

    if (tipoPeriodo === "MES") {
      fin = new Date(ahora.getFullYear(), ahora.getMonth() + 1, 0);
      fin.setHours(0, 0, 0, 0);

      inicio = new Date(ahora.getFullYear(), ahora.getMonth() - 5, 1);
      inicio.setHours(0, 0, 0, 0);

      return { inicio, fin };
    }

    fin = new Date(ahora);
    fin.setHours(0, 0, 0, 0);

    inicio = sumarDias(fin, -6);

    return { inicio, fin };
  }

  function generarGraficoPeriodo() {
    const rango = obtenerRangoPeriodo();

    if (!rango) return [];

    const registros = datosAsistencias
      .map(fila => ({
        fecha: convertirFecha(fila[1]),
        fila
      }))
      .filter(item =>
        item.fecha &&
        item.fecha >= rango.inicio &&
        item.fecha <= rango.fin
      );

    const grupos = {};

    if (tipoPeriodo === "SEMANA") {
      let cursor = new Date(rango.inicio);

      while (cursor <= rango.fin) {
        const inicio = inicioSemana(cursor);
        const fin = sumarDias(inicio, 6);
        const clave = formatearFechaDash_(inicio);

        grupos[clave] = {
          fechaInicio: new Date(inicio),
          fechaFin: new Date(fin),
          cantidad: 0
        };

        cursor = sumarDias(inicio, 7);
      }

      registros.forEach(item => {
        const inicio = inicioSemana(item.fecha);
        const clave = formatearFechaDash_(inicio);

        if (grupos[clave]) {
          grupos[clave].cantidad++;
        }
      });

      return Object.values(grupos).map(grupo => ({
        etiqueta: `${formatearFechaCortaDash_(grupo.fechaInicio)} - ${formatearFechaCortaDash_(grupo.fechaFin)}`,
        cantidad: grupo.cantidad
      }));
    }

    if (tipoPeriodo === "MES") {
      let cursor = new Date(rango.inicio.getFullYear(), rango.inicio.getMonth(), 1);

      while (cursor <= rango.fin) {
        const clave = `${cursor.getFullYear()}-${cursor.getMonth()}`;

        grupos[clave] = {
          fecha: new Date(cursor),
          cantidad: 0
        };

        cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
      }

      registros.forEach(item => {
        const clave = `${item.fecha.getFullYear()}-${item.fecha.getMonth()}`;

        if (grupos[clave]) {
          grupos[clave].cantidad++;
        }
      });

      return Object.values(grupos).map(grupo => ({
        etiqueta: Utilities.formatDate(grupo.fecha, zonaHoraria, "MMM yyyy"),
        cantidad: grupo.cantidad
      }));
    }

    let cursor = new Date(rango.inicio);

    while (cursor <= rango.fin) {
      const clave = formatearFechaDash_(cursor);

      grupos[clave] = {
        fecha: new Date(cursor),
        cantidad: 0
      };

      cursor = sumarDias(cursor, 1);
    }

    registros.forEach(item => {
      const clave = formatearFechaDash_(item.fecha);

      if (grupos[clave]) {
        grupos[clave].cantidad++;
      }
    });

    return Object.values(grupos).map(grupo => ({
      etiqueta: formatearFechaCortaDash_(grupo.fecha),
      cantidad: grupo.cantidad
    }));
  }

  const asistenciasPorDia = {};

  datosAsistencias.forEach(fila => {
    const fecha = String(fila[1]).trim();

    if (!fecha) return;

    asistenciasPorDia[fecha] =
      (asistenciasPorDia[fecha] || 0) + 1;
  });

  const fechas = Object.keys(asistenciasPorDia)
    .sort((a, b) => {
      const fechaA = convertirFecha(a);
      const fechaB = convertirFecha(b);
      return fechaA - fechaB;
    })
    .slice(-7);

  const graficoDias = fechas.map(fecha => ({
    fecha,
    cantidad: asistenciasPorDia[fecha]
  }));

  const porCarrera = {};

  registrosHoy.forEach(fila => {
    const tipoPersona = String(fila[4]).trim().toUpperCase();
    const carrera = String(fila[6]).trim();

    if (tipoPersona !== "ESTUDIANTE" || !carrera) return;

    porCarrera[carrera] =
      (porCarrera[carrera] || 0) + 1;
  });

  const asistenciasPorCarrera = Object.entries(porCarrera)
    .map(([carrera, cantidad]) => ({
      carrera,
      cantidad
    }))
    .sort((a, b) => b.cantidad - a.cantidad);

  const porMetodo = {};

  registrosHoy.forEach(fila => {
    const metodo = String(fila[3]).trim().toUpperCase();

    if (!metodo) return;

    porMetodo[metodo] =
      (porMetodo[metodo] || 0) + 1;
  });

  const asistenciasPorMetodo = Object.entries(porMetodo)
    .map(([metodo, cantidad]) => ({
      metodo,
      cantidad
    }));

  const porTipoPersona = {};

  registrosHoy.forEach(fila => {
    const tipoPersona = String(fila[4]).trim().toUpperCase();

    if (!tipoPersona) return;

    porTipoPersona[tipoPersona] =
      (porTipoPersona[tipoPersona] || 0) + 1;
  });

  const asistenciasPorTipoPersona = Object.entries(porTipoPersona)
    .map(([tipoPersona, cantidad]) => ({
      tipoPersona,
      cantidad
    }));

  const ultimosRegistros = datosAsistencias
    .filter(fila => String(fila[0]).trim())
    .slice(-10)
    .reverse()
    .map(fila => ({
      dni: String(fila[0]).trim(),
      fecha: String(fila[1]).trim(),
      hora: String(fila[2]).trim(),
      metodo: String(fila[3]).trim().toUpperCase(),
      tipoPersona: String(fila[4]).trim().toUpperCase(),
      tipoRegistro: String(fila[5]).trim().toUpperCase(),
      carrera: String(fila[6]).trim(),
      ciclo: String(fila[7]).trim()
    }));

  return {
    estado: "OK",
    fecha: hoy,
    asistenciasHoy: registrosHoy.length,
    alumnosRegistrados: estudiantesRegistrados,
    personalActivo,
    entradasPersonalHoy,
    salidasPersonalHoy,
    asistenciasPorDia: graficoDias,
    graficoPeriodo: generarGraficoPeriodo(),
    periodoSeleccionado: tipoPeriodo,
    fechaInicioPeriodo: (() => {
      const rango = obtenerRangoPeriodo();
      return rango ? formatearFechaDash_(rango.inicio) : "";
    })(),
    fechaFinPeriodo: (() => {
      const rango = obtenerRangoPeriodo();
      return rango ? formatearFechaDash_(rango.fin) : "";
    })(),
    asistenciasPorCarrera,
    asistenciasPorMetodo,
    asistenciasPorTipoPersona,
    ultimosRegistros
  };
}

function eliminarUsuario(datos) {
  datos = datos || {};

  const autorizacion = validarAutorizacion(datos.token, "USUARIOS_ELIMINAR");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  const id = String(datos.id || "").trim();
  const passwordConfirmacion = String(datos.passwordConfirmacion || "");

  if (!id) {
    return {
      estado: "ERROR",
      mensaje: "No se indicó el usuario que se desea eliminar."
    };
  }

  if (!passwordConfirmacion) {
    return {
      estado: "ERROR",
      mensaje: "Debe ingresar su contraseña para confirmar."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaUsuarios = ss.getSheetByName("USUARIOS");
  const hojaEliminaciones = ss.getSheetByName("Eliminaciones Usuarios");

  if (!hojaUsuarios) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "USUARIOS".'
    };
  }

  if (!hojaEliminaciones) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "Eliminaciones Usuarios".'
    };
  }

  const ultimaFila = hojaUsuarios.getLastRow();

  if (ultimaFila < 2) {
    return {
      estado: "NO_ENCONTRADO",
      mensaje: "No existen usuarios registrados."
    };
  }

  const datosUsuarios = hojaUsuarios
    .getRange(2, 1, ultimaFila - 1, 10)
    .getValues();

  let filaUsuario = -1;
  let usuarioActual = null;
  let administrador = null;

  for (let i = 0; i < datosUsuarios.length; i++) {
    const fila = datosUsuarios[i];
    const idFila = String(fila[0]).trim();

    if (idFila === String(autorizacion.sesion.id).trim()) {
      administrador = fila;
    }

    if (idFila === id) {
      filaUsuario = i + 2;
      usuarioActual = fila;
    }
  }

  if (!administrador) {
    return {
      estado: "ERROR",
      mensaje: "No se pudo verificar la cuenta administradora."
    };
  }

  if (String(administrador[8]).trim().toUpperCase() !== "ACTIVO") {
    return {
      estado: "ERROR",
      mensaje: "La cuenta administradora no se encuentra activa."
    };
  }

  const hashConfirmacion = generarHashPassword(
    passwordConfirmacion,
    String(administrador[7])
  );

  if (hashConfirmacion !== String(administrador[6])) {
    return {
      estado: "ERROR",
      mensaje: "La contraseña de confirmación es incorrecta."
    };
  }

  if (filaUsuario === -1 || !usuarioActual) {
    return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró el usuario indicado."
    };
  }

  if (String(autorizacion.sesion.id).trim() === id) {
    return {
      estado: "ERROR",
      mensaje: "No puede eliminar su propia cuenta."
    };
  }

  const rol = String(usuarioActual[4]).trim().toUpperCase();
  const estado = String(usuarioActual[8]).trim().toUpperCase();

  if (rol === "ADMINISTRADOR" && estado === "ACTIVO") {
    const administradoresActivos = datosUsuarios.filter(fila =>
      String(fila[4]).trim().toUpperCase() === "ADMINISTRADOR" &&
      String(fila[8]).trim().toUpperCase() === "ACTIVO"
    ).length;

    if (administradoresActivos <= 1) {
      return {
        estado: "ERROR",
        mensaje: "No puede eliminar al único administrador activo del sistema."
      };
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    hojaEliminaciones.appendRow([
      String(usuarioActual[0]).trim(),
      String(usuarioActual[1]).trim(),
      String(usuarioActual[2]).trim(),
      String(usuarioActual[3]).trim(),
      String(usuarioActual[4]).trim(),
      new Date(),
      String(autorizacion.sesion.usuario).trim()
    ]);

    invalidarSesionesUsuario(id);
    hojaUsuarios.deleteRow(filaUsuario);

    return {
      estado: "ELIMINADO",
      mensaje: "Usuario eliminado correctamente."
    };
  } finally {
    lock.releaseLock();
  }
}

function procesarSolicitudRegistroForm(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaSolicitudes = ss.getSheetByName("SOLICITUDES_REGISTRO");
  const hojaEstudiantes = ss.getSheetByName("Estudiantes");

  if (!hojaSolicitudes) throw new Error('No existe la hoja "SOLICITUDES_REGISTRO".');
  if (!hojaEstudiantes) throw new Error('No existe la hoja "Estudiantes".');
  if (!e || !e.range) throw new Error("No se recibió correctamente el evento del formulario.");

  const hojaOrigen = e.range.getSheet();

  if (hojaOrigen.getName() !== "Respuestas de formulario 1") return;

  const fila = e.range.getRow();
  const datos = hojaOrigen.getRange(fila, 1, 1, 7).getDisplayValues()[0];

  const fechaSolicitud = new Date();
  const dni = String(datos[1] || "").trim();
  const nombre = capitalizarNombre_(datos[2]);
  const apellidos = capitalizarNombre_(datos[3]);
  const carrera = String(datos[4] || "").trim();
  const ciclo = String(datos[5] || "").trim();
  const turno = String(datos[6] || "").trim().toUpperCase();

  let estado = "PENDIENTE";
  let observacion = "";

  if (!/^\d{8}$/.test(dni)) {
    estado = "RECHAZADO";
    observacion = "El DNI debe tener exactamente 8 dígitos.";
  }

  if (estado === "PENDIENTE" && !nombre) {
    estado = "RECHAZADO";
    observacion = "El nombre es obligatorio.";
  }

  if (estado === "PENDIENTE" && !apellidos) {
    estado = "RECHAZADO";
    observacion = "Los apellidos son obligatorios.";
  }

  if (estado === "PENDIENTE" && !carrera) {
    estado = "RECHAZADO";
    observacion = "La carrera es obligatoria.";
  }

  if (estado === "PENDIENTE" && !ciclo) {
    estado = "RECHAZADO";
    observacion = "El ciclo es obligatorio.";
  }

  if (estado === "PENDIENTE" && !turno) {
    estado = "RECHAZADO";
    observacion = "El turno es obligatorio.";
  }

  if (
    estado === "PENDIENTE" &&
    !["MAÑANA", "TARDE", "NOCHE"].includes(turno)
  ) {
    estado = "RECHAZADO";
    observacion = "El turno seleccionado no es válido.";
  }

  const carrerasValidas = [
    "Asistencia Administrativa",
    "Enfermeria Tecnica",
    "Tecnica en Farmacia",
    "Computacion e Informatica"
  ];

  const ciclosValidos = ["I", "II", "III", "IV", "V", "VI"];

  if (estado === "PENDIENTE" && !carrerasValidas.includes(carrera)) {
    estado = "RECHAZADO";
    observacion = "La carrera seleccionada no es válida.";
  }

  if (estado === "PENDIENTE" && !ciclosValidos.includes(ciclo)) {
    estado = "RECHAZADO";
    observacion = "El ciclo seleccionado no es válido.";
  }

  if (estado === "PENDIENTE" && !/^[A-Za-zÁÉÍÓÚÜÑáéíóúüñÀ-ÿ' -]+$/.test(nombre)) {
    estado = "RECHAZADO";
    observacion = "El nombre contiene caracteres no válidos.";
  }

  if (estado === "PENDIENTE" && !/^[A-Za-zÁÉÍÓÚÜÑáéíóúüñÀ-ÿ' -]+$/.test(apellidos)) {
    estado = "RECHAZADO";
    observacion = "Los apellidos contienen caracteres no válidos.";
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const idSolicitud = generarIdSolicitud(hojaSolicitudes);

    if (estado === "PENDIENTE") {
      const ultimaFila = hojaEstudiantes.getLastRow();

      if (ultimaFila >= 2) {
        const dnis = hojaEstudiantes
          .getRange(2, 1, ultimaFila - 1, 1)
          .getDisplayValues()
          .flat()
          .map(valor => normalizarDni_(valor));

        if (dnis.includes(dni)) {
          estado = "RECHAZADO";
          observacion = "El DNI ya se encuentra registrado como alumno.";
        }
      }
    }

    hojaSolicitudes.appendRow([
      idSolicitud,
      fechaSolicitud,
      dni,
      nombre,
      apellidos,
      carrera,
      ciclo,
      turno,
      estado,
      observacion,
      "",
      ""
    ]);
  } finally {
    lock.releaseLock();
  }
}

function generarIdSolicitud(hojaSolicitudes) {
  const ultimaFila = hojaSolicitudes.getLastRow();

  if (ultimaFila < 2) return "SOL-000001";

  const ids = hojaSolicitudes
    .getRange(2, 1, ultimaFila - 1, 1)
    .getDisplayValues()
    .flat();

  let mayor = 0;

  ids.forEach(id => {
    const coincidencia = String(id).match(/^SOL-(\d+)$/);
    if (coincidencia) {
      mayor = Math.max(mayor, Number(coincidencia[1]));
    }
  });

  return `SOL-${String(mayor + 1).padStart(6, "0")}`;
}

function obtenerPermisosSistema(token, idUsuario) {
  const autorizacion = validarAutorizacion(token, "USUARIOS_PERMISOS");

  if (!autorizacion.autorizado) return autorizacion;

  idUsuario = String(idUsuario || "").trim();

  if (!idUsuario) {
    return {
      estado: "ERROR",
      mensaje: "No se indicó el usuario."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPermisos = ss.getSheetByName("PERMISOS");
  const hojaUsuarios = ss.getSheetByName("USUARIOS");
  const hojaCargos = ss.getSheetByName("CARGOS");
  const hojaPermisosCargos = ss.getSheetByName("PERMISOS_CARGOS");
  const hojaPermisosUsuarios = ss.getSheetByName("PERMISOS_USUARIOS");

  if (!hojaPermisos || !hojaUsuarios || !hojaCargos || !hojaPermisosCargos || !hojaPermisosUsuarios) {
    return {
      estado: "ERROR",
      mensaje: "No se encontraron todas las hojas de permisos necesarias."
    };
  }

  const datosUsuarios = hojaUsuarios.getLastRow() >= 2
    ? hojaUsuarios.getRange(2, 1, hojaUsuarios.getLastRow() - 1, 10).getDisplayValues()
    : [];

  const usuario = datosUsuarios.find(fila =>
    String(fila[0]).trim() === idUsuario
  );

  if (!usuario) {
    return {
      estado: "NO_ENCONTRADO",
      mensaje: "No se encontró el usuario indicado."
    };
  }

  const cargoId = String(usuario[5] || "").trim();

  const permisos = hojaPermisos.getLastRow() >= 2
    ? hojaPermisos.getRange(2, 1, hojaPermisos.getLastRow() - 1, 5).getDisplayValues()
    : [];

  const permisosCargos = hojaPermisosCargos.getLastRow() >= 2
    ? hojaPermisosCargos.getRange(2, 1, hojaPermisosCargos.getLastRow() - 1, 3).getDisplayValues()
    : [];

  const permisosUsuarios = hojaPermisosUsuarios.getLastRow() >= 2
    ? hojaPermisosUsuarios.getRange(2, 1, hojaPermisosUsuarios.getLastRow() - 1, 3).getDisplayValues()
    : [];

  const resultado = permisos.map(permiso => {
    const idPermiso = String(permiso[0]).trim().toUpperCase();
    const sistema = String(permiso[4]).trim().toUpperCase() === "SÍ";

    let cargoHabilitado = false;
    let usuarioHabilitado = null;

    const permisoCargo = permisosCargos.find(fila =>
      String(fila[0]).trim() === cargoId &&
      String(fila[1]).trim().toUpperCase() === idPermiso
    );

    if (permisoCargo) {
      cargoHabilitado =
        String(permisoCargo[2]).trim().toUpperCase() === "SÍ";
    }

    const permisoUsuario = permisosUsuarios.find(fila =>
      String(fila[0]).trim() === idUsuario &&
      String(fila[1]).trim().toUpperCase() === idPermiso
    );

    if (permisoUsuario) {
      usuarioHabilitado =
        String(permisoUsuario[2]).trim().toUpperCase() === "SÍ";
    }

    const habilitado = usuarioHabilitado === null
      ? cargoHabilitado
      : usuarioHabilitado;

    return {
      idPermiso,
      modulo: String(permiso[1]).trim(),
      accion: String(permiso[2]).trim(),
      descripcion: String(permiso[3]).trim(),
      sistema,
      cargoHabilitado,
      usuarioHabilitado,
      habilitado
    };
  });

  return {
    estado: "OK",
    usuario: {
      id: idUsuario,
      nombre: String(usuario[2]).trim(),
      apellidos: String(usuario[3]).trim(),
      rol: String(usuario[4]).trim().toUpperCase(),
      cargo: cargoId
    },
    permisos: resultado
  };
}

function guardarPermisosUsuario(datos) {
  datos = datos || {};

  const autorizacion = validarAutorizacion(
    datos.token,
    "USUARIOS_PERMISOS"
  );

  if (!autorizacion.autorizado) return autorizacion;

  const idUsuario = String(datos.idUsuario || "").trim();
  const permisos = Array.isArray(datos.permisos)
    ? datos.permisos
    : [];

  if (!idUsuario) {
    return {
      estado: "ERROR",
      mensaje: "No se indicó el usuario."
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPermisos = ss.getSheetByName("PERMISOS");
  const hojaPermisosUsuarios = ss.getSheetByName("PERMISOS_USUARIOS");

  if (!hojaPermisos || !hojaPermisosUsuarios) {
    return {
      estado: "ERROR",
      mensaje: "No se encontraron las hojas de permisos."
    };
  }

  const datosPermisos = hojaPermisos.getLastRow() >= 2
    ? hojaPermisos.getRange(
      2,
      1,
      hojaPermisos.getLastRow() - 1,
      5
    ).getDisplayValues()
    : [];

  const protegidos = new Set(
    datosPermisos
      .filter(fila =>
        String(fila[4]).trim().toUpperCase() === "SÍ"
      )
      .map(fila =>
        String(fila[0]).trim().toUpperCase()
      )
  );

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ultimaFila = hojaPermisosUsuarios.getLastRow();

    if (ultimaFila >= 2) {
      const existentes = hojaPermisosUsuarios
        .getRange(2, 1, ultimaFila - 1, 3)
        .getDisplayValues();

      for (let i = existentes.length - 1; i >= 0; i--) {
        if (String(existentes[i][0]).trim() === idUsuario) {
          hojaPermisosUsuarios.deleteRow(i + 2);
        }
      }
    }

    const nuevasFilas = [];

    permisos.forEach(item => {
      const idPermiso =
        String(item.idPermiso || "").trim().toUpperCase();

      if (!idPermiso || protegidos.has(idPermiso)) return;

      const valor =
        String(item.habilitado || "").trim().toUpperCase();

      if (valor !== "SÍ" && valor !== "NO") return;

      nuevasFilas.push([
        idUsuario,
        idPermiso,
        valor
      ]);
    });

    if (nuevasFilas.length) {
      hojaPermisosUsuarios
        .getRange(
          hojaPermisosUsuarios.getLastRow() + 1,
          1,
          nuevasFilas.length,
          3
        )
        .setValues(nuevasFilas);
      CacheService.getScriptCache().remove("MAP_PERMISOS_V1");
    }

    invalidarSesionesUsuario(idUsuario);

    return {
      estado: "ACTUALIZADO",
      mensaje: "Permisos del usuario actualizados correctamente."
    };
  } finally {
    lock.releaseLock();
  }
}

function obtenerEstadoAsistenciaPorTurno_(turno, fechaHora) {
  const configuracion = obtenerTurnosCache_();
  const turnoNormalizado = String(turno || "").trim().toUpperCase();

  const datosTurno = configuracion[turnoNormalizado];

  if (!datosTurno) {
    return "ASISTIÓ";
  }

  const [hora, minuto] = datosTurno.cierre.split(":").map(Number);

  const cierre = new Date(fechaHora);
  cierre.setHours(hora, minuto, 0, 0);

  return fechaHora.getTime() >= cierre.getTime()
    ? "ASISTIÓ_TARDE"
    : "ASISTIÓ";
}

function crearTriggerCierreAsistencia() {
  const nombreFuncion = "cerrarControlAsistenciaHoy";

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === nombreFuncion) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const cfg = obtenerConfiguracionTurnos_();
  const horas = [cfg.MAÑANA.cierre, cfg.TARDE.cierre, cfg.NOCHE.cierre]
    .map(h => Number(String(h).split(":")[0]))
    .filter(h => h >= 0 && h <= 23);
  if (!horas.length) {
    return { estado: "ERROR", mensaje: "Cierres inválidos en Configuración." };
  }
  [...new Set(horas)].forEach(hora => {
    ScriptApp.newTrigger(nombreFuncion)
      .timeBased()
      .everyDays(1)
      .atHour(hora)
      .create();
  });

  return {
    estado: "OK",
    mensaje: "Trigger de cierre de asistencia creado correctamente."
  };
}

function eliminarTriggerCierreAsistencia() {
  const nombreFuncion = "cerrarControlAsistenciaHoy";

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === nombreFuncion) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  return {
    estado: "OK",
    mensaje: "Trigger de cierre de asistencia eliminado correctamente."
  };
}

function obtenerEstadoTriggerCierre() {
  const n = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "cerrarControlAsistenciaHoy").length;
  return { estado: "OK", activos: n };
}

function justificarInasistencia(datos) {
  const autorizacion = validarAutorizacion(datos.token, "ASISTENCIA_JUSTIFICAR");

  if (!autorizacion.autorizado) {
    return {
      estado: autorizacion.estado,
      mensaje: autorizacion.mensaje
    };
  }

  datos = datos || {};

  const dni = String(datos.dni || "").trim();
  const fechaTexto = String(datos.fecha || "").trim();
  const motivo = String(datos.motivo || "").trim();
  const observacion = String(datos.observacion || "").trim();
  const documentoId = String(datos.documentoId || "").trim();
  const documentoNombre = String(datos.documentoNombre || "").trim();

  if (!dni || !fechaTexto || !motivo) {
    return {
      estado: "ERROR",
      mensaje: "DNI, fecha y motivo son obligatorios."
    };
  }

  const partes = fechaTexto.split("/");

  if (
    partes.length !== 3 ||
    isNaN(Number(partes[0])) ||
    isNaN(Number(partes[1])) ||
    isNaN(Number(partes[2]))
  ) {
    return {
      estado: "ERROR",
      mensaje: "La fecha no tiene un formato válido."
    };
  }

  const fecha = new Date(
    Number(partes[2]),
    Number(partes[1]) - 1,
    Number(partes[0])
  );

  fecha.setHours(0, 0, 0, 0);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Control de Asistencia");

  if (!hoja) {
    return {
      estado: "ERROR",
      mensaje: 'No existe la hoja "Control de Asistencia".'
    };
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return {
      estado: "NO_ENCONTRADO",
      mensaje: "No existen registros de asistencia."
    };
  }

  const datosControl = hoja
    .getRange(2, 1, ultimaFila - 1, 12)
    .getValues();

  for (let i = 0; i < datosControl.length; i++) {
    const fechaControl = datosControl[i][1];
    const dniControl = String(datosControl[i][2] || "").trim();
    const tipoPersona = String(datosControl[i][3] || "").trim().toUpperCase();
    const estado = String(datosControl[i][5] || "").trim().toUpperCase();

    if (
      dniControl === dni &&
      tipoPersona === "ESTUDIANTE" &&
      fechaControl instanceof Date &&
      esMismaFecha(fechaControl, fecha)
    ) {
      if (estado !== "FALTÓ") {
        return {
          estado: "ERROR",
          mensaje: "El registro seleccionado no está en estado FALTÓ."
        };
      }

      const fila = i + 2;
      const ahora = new Date();

      hoja.getRange(fila, 6, 1, 7).setValues([[
        "JUSTIFICADO",
        motivo,
        observacion,
        documentoId,
        documentoNombre,
        String(autorizacion.sesion.nombre || "") + " " +
        String(autorizacion.sesion.apellidos || ""),
        ahora
      ]]);

      hoja.getRange(fila, 12).setNumberFormat("dd/MM/yyyy hh:mm:ss AM/PM");

      return {
        estado: "JUSTIFICADO",
        mensaje: "La inasistencia fue justificada correctamente."
      };
    }
  }

  return {
    estado: "NO_ENCONTRADO",
    mensaje: "No se encontró una inasistencia para el alumno y fecha indicados."
  };
}

function guardarDocumentoJustificacion(datos) {
  const autorizacion = validarAutorizacion(datos.token, "ASISTENCIA_JUSTIFICAR");
  if (!autorizacion.autorizado) return autorizacion;
  datos = datos || {};
  const contenido = String(datos.contenido || "").trim();
  const mimeType = String(datos.mimeType || "").trim();
  const nombreOrig = String(datos.nombre || "").trim();
  if (!contenido || !mimeType) return { estado: "ERROR", mensaje: "No se recibió archivo." };
  if (!/^application\/pdf$|^image\/(jpeg|png|webp)$/i.test(mimeType)) return { estado: "ERROR", mensaje: "Solo PDF, JPG, PNG o WEBP." };
  const partes = contenido.split(",");
  if (partes.length !== 2) return { estado: "ERROR", mensaje: "Formato archivo inválido." };
  const bytes = Utilities.base64Decode(partes[1]);
  if (bytes.length > 5 * 1024 * 1024) return { estado: "ERROR", mensaje: "Máximo 5 MB." };
  const ID_CARPETA_JUSTIFICANTES = "14q_StnnMY_kCWWVzCgraxLTWtbTKd5lG";
  const carpeta = DriveApp.getFolderById(ID_CARPETA_JUSTIFICANTES);
  const ext = mimeType === "application/pdf" ? "pdf" : mimeType.split("/")[1].toLowerCase();
  const nombreSeguro = String(new Date().getTime()) + "_" + nombreOrig.replace(/[^\w.\-]+/g, "_") + "." + ext;
  const blob = Utilities.newBlob(bytes, mimeType, nombreSeguro);
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  return { estado: "GUARDADO", documentoId: archivo.getId(), documentoNombre: archivo.getName() };
}

function obtenerMapaEstudiantes_() {
  const cache = CacheService.getScriptCache();
  const g = cache.get("MAP_EST_V1");
  if (g) return JSON.parse(g);
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estudiantes");
  const last = hoja.getLastRow();
  const mapa = {};
  if (last >= 2) {
    hoja.getRange(2, 1, last - 1, 6).getValues().forEach(r => {
      const dni = normalizarDni_(r[0]);
      if (!dni) return;
      mapa[dni] = { nombre: r[1], apellidos: r[2], carrera: r[3], ciclo: r[4], turno: String(r[5] || "").trim().toUpperCase() };
    });
  }
  cache.put("MAP_EST_V1", JSON.stringify(mapa), 300);
  return mapa;
}

function obtenerMapaPersonal_() {
  const cache = CacheService.getScriptCache();
  const g = cache.get("MAP_PER_V1");
  if (g) return JSON.parse(g);
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal");
  const last = hoja.getLastRow();
  const mapa = {};
  if (last >= 2) {
    hoja.getRange(2, 1, last - 1, 6).getValues().forEach(r => {
      const dni = normalizarDni_(r[0]);
      if (!dni) return;
      mapa[dni] = { nombre: r[1], apellidos: r[2], tipo: r[3], cargo: r[4], estado: r[5] };
    });
  }
  cache.put("MAP_PER_V1", JSON.stringify(mapa), 300);
  return mapa;
}

function obtenerTurnosCache_() {
  const cache = CacheService.getScriptCache();
  const g = cache.get("TURNOS_V1");
  if (g) return JSON.parse(g);
  const t = obtenerConfiguracionTurnos_();
  cache.put("TURNOS_V1", JSON.stringify(t), 600);
  return t;
}

function limpiarCacheMaestros_(cual) {
  const cache = CacheService.getScriptCache();
  if (!cual || cual === "EST") cache.remove("MAP_EST_V1");
  if (!cual || cual === "PER") cache.remove("MAP_PER_V1");
  if (!cual || cual === "TURNOS") cache.remove("TURNOS_V1");
}

function versionPadron_() {
  return PropertiesService.getScriptProperties().getProperty("VERSION_PADRON") || "1";
}

function subirVersionPadron_() {
  const p = PropertiesService.getScriptProperties();
  const v = parseInt(p.getProperty("VERSION_PADRON") || "1", 10);
  p.setProperty("VERSION_PADRON", String((isNaN(v) ? 1 : v) + 1));
}

function obtenerPadronRapido() {
  const mapaEst = obtenerMapaEstudiantes_();
  const mapaPer = obtenerMapaPersonal_();
  const alumnos = {};
  Object.keys(mapaEst).forEach(dni => {
    const e = mapaEst[dni];
    alumnos[dni] = { nombre: (String(e.nombre || "").trim() + " " + String(e.apellidos || "").trim()).trim(), turno: e.turno || "" };
  });
  Object.keys(mapaPer).forEach(dni => {
    const p = mapaPer[dni];
    if (String(p.estado || "").trim().toUpperCase() !== "ACTIVO") return;
    alumnos[dni] = { nombre: (String(p.nombre || "").trim() + " " + String(p.apellidos || "").trim()).trim(), turno: "" };
  });
  return { estado: "OK", version: versionPadron_(), alumnos: alumnos };
}

function obtenerMatricesPermisos_() {
  const cache = CacheService.getScriptCache();
  const g = cache.get("MAP_PERMISOS_V1");
  if (g) return JSON.parse(g);
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const r = (nombre, cols) => {
    const h = libro.getSheetByName(nombre);
    if (!h || h.getLastRow() < 2) return [];
    return h.getRange(2, 1, h.getLastRow() - 1, cols).getValues();
  };
  const m = {
    permisos: r("PERMISOS", 5),
    cargos: r("PERMISOS_CARGOS", 3),
    usuarios: r("PERMISOS_USUARIOS", 3)
  };
  cache.put("MAP_PERMISOS_V1", JSON.stringify(m), 600);
  return m;
}

function capitalizarNombre_(texto) {
  const minusculas = new Set(["de", "del", "la", "las", "el", "los", "y", "e", "en", "da", "do", "dos"]);
  return String(texto || "").trim().replace(/\s+/g, " ")
    .toLowerCase().split(" ")
    .map((palabra, i) => (i > 0 && minusculas.has(palabra)) ? palabra : palabra.replace(/^[a-záéíóúüñ]/i, c => c.toUpperCase()))
    .join(" ");
}

function normalizarDni_(v, rellenar) {
  const d = String(v == null ? "" : v).replace(/\D/g, "");
  if (rellenar === false) return d;
  return d.length <= 8 ? d.padStart(8, "0") : d;
}