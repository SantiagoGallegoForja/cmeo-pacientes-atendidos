/**
 * Apps Script para identificar pacientes atendidos hace 3 dias
 * y agregarles el tag "paciente-atendido-3d" en GoHighLevel
 *
 * Configuracion requerida (Propiedades del script):
 *   GHL_API_KEY   - API Key v2 de GoHighLevel
 *   GHL_LOCATION_ID - Location ID de la cuenta GHL
 */

var GHL_BASE_URL = "https://services.leadconnectorhq.com";
var TAG_NOMBRE = "paciente-atendido-3d";
var GHL_API_VERSION = "2021-07-28";

// Mapeo de nombre en el Sheet -> valor del custom field en GHL
var DOCTOR_MAP = {
  "Daniel": "Dr. Daniel Avila",
  "Carolina": "Dra. Carolina Campuzano"
};

// ============================================================
// FUNCION PRINCIPAL
// ============================================================

function tagPacientesAtendidos() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty("GHL_API_KEY");
  var locationId = props.getProperty("GHL_LOCATION_ID");

  if (!apiKey || !locationId) {
    Logger.log("ERROR: Configura GHL_API_KEY y GHL_LOCATION_ID en Propiedades del script");
    return;
  }

  // Buscar el ID del custom field "Doctor"
  var doctorFieldId = obtenerCustomFieldId(apiKey, locationId, "Doctor");
  if (!doctorFieldId) {
    Logger.log("ADVERTENCIA: No se encontro el custom field 'Doctor' en GHL. Se continuara sin asignar doctor.");
  } else {
    Logger.log("Custom field Doctor ID: " + doctorFieldId);
  }

  // Calcular fecha de hace 3 dias
  var hace3dias = new Date();
  hace3dias.setDate(hace3dias.getDate() - 3);
  var fechaStr = Utilities.formatDate(hace3dias, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var nombreHoja = "Pacientes_" + fechaStr;

  Logger.log("Buscando hoja: " + nombreHoja);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) {
    Logger.log("No se encontro la hoja: " + nombreHoja);
    return;
  }

  var datos = hoja.getDataRange().getValues();
  var encabezados = datos[0];

  // Encontrar indices de columnas
  var colCelular = encabezados.indexOf("Celular");
  var colTelefono = encabezados.indexOf("Teléfono");
  var colNombre = encabezados.indexOf("Nombre");
  var colDoctor = encabezados.indexOf("Doctor");

  Logger.log("Columnas - Celular:" + colCelular + " Telefono:" + colTelefono + " Nombre:" + colNombre + " Doctor:" + colDoctor);

  if (colCelular === -1 && colTelefono === -1) {
    Logger.log("ERROR: No se encontraron columnas Celular ni Telefono");
    return;
  }

  var totalPacientes = datos.length - 1;
  var tagExitoso = 0;
  var noEncontrado = 0;
  var errores = 0;

  Logger.log("Procesando " + totalPacientes + " pacientes de " + nombreHoja);

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var nombre = colNombre >= 0 ? fila[colNombre] : "Sin nombre";
    var celular = colCelular >= 0 ? limpiarTelefono(String(fila[colCelular])) : "";
    var telefono = colTelefono >= 0 ? limpiarTelefono(String(fila[colTelefono])) : "";
    var doctor = colDoctor >= 0 ? fila[colDoctor] : "";

    Logger.log("--- Paciente " + i + "/" + totalPacientes + ": " + nombre + " (Dr. " + doctor + ")");

    // Buscar primero por celular, luego por telefono
    var contacto = null;

    if (celular) {
      contacto = buscarContactoGHL(apiKey, locationId, celular);
    }

    if (!contacto && telefono && telefono !== celular) {
      Logger.log("  No encontrado por celular, intentando por telefono: " + telefono);
      contacto = buscarContactoGHL(apiKey, locationId, telefono);
    }

    if (contacto) {
      // Mapear nombre del doctor al valor del custom field
      var doctorGHL = DOCTOR_MAP[doctor] || "";
      var resultado = actualizarContactoGHL(apiKey, contacto, TAG_NOMBRE, doctorFieldId, doctorGHL);
      if (resultado) {
        tagExitoso++;
        Logger.log("  Tag + Doctor asignados exitosamente a: " + nombre + " -> " + doctorGHL);
      } else {
        errores++;
        Logger.log("  Error al actualizar contacto: " + nombre);
      }
    } else {
      noEncontrado++;
      Logger.log("  Contacto NO encontrado en GHL: " + nombre + " (cel:" + celular + " tel:" + telefono + ")");
    }

    // Pausa para no exceder rate limits de GHL
    Utilities.sleep(500);
  }

  Logger.log("============================================================");
  Logger.log("RESUMEN");
  Logger.log("============================================================");
  Logger.log("Total pacientes: " + totalPacientes);
  Logger.log("Tag agregado: " + tagExitoso);
  Logger.log("No encontrados en GHL: " + noEncontrado);
  Logger.log("Errores: " + errores);
}

// ============================================================
// FUNCIONES GHL API
// ============================================================

function buscarContactoGHL(apiKey, locationId, telefono) {
  try {
    var url = GHL_BASE_URL + "/contacts/?locationId=" + locationId + "&query=" + encodeURIComponent(telefono) + "&limit=1";

    var response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "Authorization": "Bearer " + apiKey,
        "Version": GHL_API_VERSION
      },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      Logger.log("  Error buscando contacto (HTTP " + code + "): " + response.getContentText().substring(0, 200));
      return null;
    }

    var data = JSON.parse(response.getContentText());
    var contactos = data.contacts || [];

    if (contactos.length > 0) {
      Logger.log("  Contacto encontrado: " + contactos[0].contactName + " (ID: " + contactos[0].id + ")");
      return contactos[0];
    }

    return null;

  } catch (e) {
    Logger.log("  Excepcion buscando contacto: " + e.message);
    return null;
  }
}

function actualizarContactoGHL(apiKey, contacto, nuevoTag, doctorFieldId, doctorValor) {
  try {
    var payload = {};

    // Agregar tag (sin duplicar)
    var tagsActuales = contacto.tags || [];
    if (tagsActuales.indexOf(nuevoTag) === -1) {
      tagsActuales.push(nuevoTag);
    }
    payload.tags = tagsActuales;

    // Agregar custom field Doctor si tenemos el ID y el valor
    if (doctorFieldId && doctorValor) {
      payload.customFields = [
        { id: doctorFieldId, value: doctorValor }
      ];
    }

    var url = GHL_BASE_URL + "/contacts/" + contacto.id;

    var response = UrlFetchApp.fetch(url, {
      method: "put",
      headers: {
        "Authorization": "Bearer " + apiKey,
        "Version": GHL_API_VERSION,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code === 200) {
      return true;
    } else {
      Logger.log("  Error actualizando contacto (HTTP " + code + "): " + response.getContentText().substring(0, 200));
      return false;
    }

  } catch (e) {
    Logger.log("  Excepcion actualizando contacto: " + e.message);
    return false;
  }
}

function obtenerCustomFieldId(apiKey, locationId, nombreCampo) {
  try {
    var url = GHL_BASE_URL + "/locations/" + locationId + "/customFields";

    var response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        "Authorization": "Bearer " + apiKey,
        "Version": GHL_API_VERSION
      },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      Logger.log("Error obteniendo custom fields (HTTP " + code + "): " + response.getContentText().substring(0, 200));
      return null;
    }

    var data = JSON.parse(response.getContentText());
    var fields = data.customFields || [];

    for (var i = 0; i < fields.length; i++) {
      if (fields[i].name === nombreCampo) {
        return fields[i].id;
      }
    }

    Logger.log("Custom field '" + nombreCampo + "' no encontrado entre " + fields.length + " campos");
    return null;

  } catch (e) {
    Logger.log("Excepcion obteniendo custom fields: " + e.message);
    return null;
  }
}

// ============================================================
// UTILIDADES
// ============================================================

function limpiarTelefono(tel) {
  if (!tel) return "";
  // Quitar espacios, guiones, puntos, parentesis
  tel = tel.replace(/[\s\-\.\(\)]/g, "");
  // Si no tiene codigo de pais y tiene 10 digitos, agregar +57
  if (/^\d{10}$/.test(tel)) {
    tel = "+57" + tel;
  }
  // Si empieza con 57 y tiene 12 digitos, agregar +
  if (/^57\d{10}$/.test(tel)) {
    tel = "+" + tel;
  }
  return tel;
}

// ============================================================
// CONFIGURACION INICIAL (ejecutar una sola vez)
// ============================================================

function configurarCredenciales() {
  var ui = SpreadsheetApp.getUi();

  var respApiKey = ui.prompt("Configuracion GHL", "Ingresa tu API Key de GoHighLevel:", ui.ButtonSet.OK_CANCEL);
  if (respApiKey.getSelectedButton() !== ui.Button.OK) return;

  var respLocationId = ui.prompt("Configuracion GHL", "Ingresa tu Location ID de GoHighLevel:", ui.ButtonSet.OK_CANCEL);
  if (respLocationId.getSelectedButton() !== ui.Button.OK) return;

  var props = PropertiesService.getScriptProperties();
  props.setProperty("GHL_API_KEY", respApiKey.getResponseText().trim());
  props.setProperty("GHL_LOCATION_ID", respLocationId.getResponseText().trim());

  ui.alert("Credenciales guardadas correctamente.");
}

// ============================================================
// TRIGGER AUTOMATICO DIARIO A LAS 9 AM
// ============================================================

function activarTriggerDiario() {
  // Eliminar triggers anteriores de esta funcion para no duplicar
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "tagPacientesAtendidos") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Crear trigger diario a las 9 AM (zona horaria del script/sheet)
  ScriptApp.newTrigger("tagPacientesAtendidos")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log("Trigger diario activado: tagPacientesAtendidos a las 9 AM");
  SpreadsheetApp.getUi().alert("Trigger activado: se ejecutara automaticamente todos los dias a las 9 AM.");
}

function desactivarTriggerDiario() {
  var eliminados = 0;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "tagPacientesAtendidos") {
      ScriptApp.deleteTrigger(triggers[i]);
      eliminados++;
    }
  }

  Logger.log("Triggers eliminados: " + eliminados);
  SpreadsheetApp.getUi().alert("Trigger diario desactivado.");
}

// ============================================================
// MENU PERSONALIZADO
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CMEO Pacientes")
    .addItem("Configurar credenciales GHL", "configurarCredenciales")
    .addSeparator()
    .addItem("Tag pacientes hace 3 dias (manual)", "tagPacientesAtendidos")
    .addSeparator()
    .addItem("Activar trigger diario 9 AM", "activarTriggerDiario")
    .addItem("Desactivar trigger diario", "desactivarTriggerDiario")
    .addToUi();
}
