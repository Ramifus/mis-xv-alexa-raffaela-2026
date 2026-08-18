/**
 * Google Apps Script para recibir las confirmaciones de Confirmar.astro
 *
 * CÓMO USARLO
 * 1. Abre tu Google Sheet -> Extensiones -> Apps Script.
 * 2. Borra todo y pega este código.
 * 3. Cambia NOMBRE_HOJA si tu pestaña no se llama "Hoja 1".
 * 4. Implementar -> Nueva implementación -> Tipo: Aplicación web
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: CUALQUIER USUARIO  <-- importante
 * 5. Copia la URL /exec y pégala en SCRIPT_URL dentro de Confirmar.astro.
 *
 * OJO: cada vez que edites el script tienes que crear una NUEVA implementación
 * (o "Administrar implementaciones" -> editar -> Versión: Nueva) para que los
 * cambios se apliquen a la misma URL.
 */

var NOMBRE_HOJA = "Hoja 1";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // evita que dos invitados escriban en la misma fila

  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);

    if (!hoja) {
      hoja = SpreadsheetApp.getActiveSpreadsheet().insertSheet(NOMBRE_HOJA);
    }

    // encabezados la primera vez
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(["#", "Nombre", "Fecha de confirmación"]);
    }

    var nombre = "";

    // 1) urlencoded (lo que manda Confirmar.astro)
    if (e && e.parameter && e.parameter.nombre) {
      nombre = e.parameter.nombre;
    }
    // 2) respaldo: por si llega como JSON
    else if (e && e.postData && e.postData.contents) {
      try {
        var json = JSON.parse(e.postData.contents);
        nombre = json.nombre || "";
      } catch (err) {
        nombre = "";
      }
    }

    nombre = String(nombre).trim();

    if (!nombre) {
      return respuesta({ ok: false, error: "nombre vacio" });
    }

    var numero = hoja.getLastRow(); // fila 1 = encabezado, así queda 1, 2, 3...
    var fecha = Utilities.formatDate(
      new Date(),
      "America/Mexico_City",
      "dd/MM/yyyy HH:mm"
    );

    hoja.appendRow([numero, nombre, fecha]);

    return respuesta({ ok: true, nombre: nombre, total: numero });
  } catch (error) {
    return respuesta({ ok: false, error: String(error) });
  } finally {
    lock.releaseLock();
  }
}

/** Abrir la URL /exec en el navegador devuelve el total de confirmados. */
function doGet() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  var total = hoja ? Math.max(0, hoja.getLastRow() - 1) : 0;
  return respuesta({ ok: true, total: total });
}

function respuesta(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto)).setMimeType(
    ContentService.MimeType.JSON
  );
}
