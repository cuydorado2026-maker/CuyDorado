function doPost(e) {
  const spreadsheetId = "1DdW74FjyLzFhgVrWJbsj1zrF0m5Hg6QMzkwZ6jE7OR4";
  const sheetName = "Hoja 1";
  const parametros = e && e.parameter ? e.parameter : {};

  if (String(parametros.accion || "") === "canjear") {
    const libro = SpreadsheetApp.openById(spreadsheetId);
    return respuesta(registrarDatosCanje(libro, parametros));
  }

  const codigo = e && e.parameter ? String(e.parameter.codigo || "").trim() : "";

  if (!codigo) {
    return respuesta({ ok: false, error: "El codigo es obligatorio." });
  }

  const libro = SpreadsheetApp.openById(spreadsheetId);
  const hoja = libro.getSheetByName(sheetName) || libro.getSheets()[0];

  if (!hoja) {
    return respuesta({ ok: false, error: "No se encontro una pestaña en la hoja." });
  }

  return respuesta(registrarCodigo(libro, hoja, codigo));
}

function registrarDatosCanje(libro, parametros) {
  parametros = parametros || {};
  if (!libro) {
    return { ok: false, error: "Ejecuta el canje desde la pagina web, no desde esta funcion." };
  }

  const codigo = String(parametros.codigo || "").trim().toUpperCase();
  const nombre = String(parametros.nombre || "").trim();
  const telefono = String(parametros.telefono || "").trim();
  const correo = String(parametros.correo || "").trim();

  if (!codigo || !nombre || !telefono || !correo) {
    return { ok: false, error: "Todos los datos son obligatorios." };
  }

  const registros = libro.getSheetByName("Registros");
  if (!registros || registros.getLastRow() < 2) {
    return { ok: false, error: "No se encontro el registro del codigo." };
  }

  const bloqueo = LockService.getScriptLock();
  bloqueo.waitLock(30000);

  try {
    const codigos = registros.getRange(2, 2, registros.getLastRow() - 1, 1)
      .getDisplayValues()
      .flat()
      .map(function (valor) { return String(valor).trim().toUpperCase(); });
    const indice = codigos.indexOf(codigo);

    if (indice === -1) {
      return { ok: false, error: "El codigo no esta registrado." };
    }

    registros.getRange(indice + 2, 3, 1, 3).setValues([[nombre, telefono, correo]]);
    return { ok: true, guardado: true };
  } finally {
    bloqueo.releaseLock();
  }
}

function doGet(e) {
  const spreadsheetId = "1DdW74FjyLzFhgVrWJbsj1zrF0m5Hg6QMzkwZ6jE7OR4";
  const sheetName = "Hoja 1";
  const codigo = e && e.parameter ? String(e.parameter.codigo || "").trim() : "";

  if (!codigo) {
    return respuesta({ ok: true, mensaje: "Receptor activo." });
  }

  const libro = SpreadsheetApp.openById(spreadsheetId);
  const hoja = libro.getSheetByName(sheetName) || libro.getSheets()[0];
  return hoja ? respuesta(registrarCodigo(libro, hoja, codigo)) : respuesta({ ok: false, error: "No se encontro la hoja." });
}

function registrarCodigo(libro, hoja, codigo) {
  const resultado = validarCodigo(hoja, codigo);

  if (!resultado.valido) {
    return resultado;
  }

  const bloqueo = LockService.getScriptLock();
  bloqueo.waitLock(30000);

  try {
    const registros = libro.getSheetByName("Registros") || libro.insertSheet("Registros");
    if (registros.getLastRow() === 0) {
      registros.appendRow(["Fecha", "Codigo", "Nombre", "Telefono", "Correo"]);
    }

    const codigoNormalizado = codigo.toUpperCase();
    const codigosRegistrados = registros.getRange(2, 2, Math.max(registros.getLastRow() - 1, 1), 1)
      .getDisplayValues()
      .flat()
      .map(function (valor) { return String(valor).trim().toUpperCase(); });

    if (codigosRegistrados.includes(codigoNormalizado)) {
      return Object.assign(resultado, { yaRegistrado: true });
    }

    registros.appendRow([new Date(), codigoNormalizado]);
    return Object.assign(resultado, { yaRegistrado: false });
  } finally {
    bloqueo.releaseLock();
  }
}

function validarCodigo(hoja, codigo) {
  const filas = hoja.getDataRange().getDisplayValues();
  const codigoNormalizado = codigo.toUpperCase();

  for (let indice = 0; indice < filas.length; indice++) {
    const fila = filas[indice];
    const codigoEnFila = String(fila[0] || "").trim().toUpperCase();

    if (codigoEnFila !== codigoNormalizado) {
      continue;
    }

    const montoTexto = String(fila[1] || "0").replace(/[S/$\s]/gi, "").replace(",", ".");
    const monto = Number(montoTexto) || 0;
    return {
      ok: true,
      valido: true,
      gano: monto > 0,
      monto: monto.toFixed(2)
    };
  }

  return { ok: true, valido: false, gano: false, monto: "0.00" };
}

function respuesta(datos) {
  return ContentService
    .createTextOutput(JSON.stringify(datos))
    .setMimeType(ContentService.MimeType.JSON);
}
