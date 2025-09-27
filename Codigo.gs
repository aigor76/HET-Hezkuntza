// AÑADIR ESTA FUNCIÓN ANTES DE onFormSubmit
function extraerDatosFormulario(namedValues) {
  let email = null;
  let nombreFamilia = null;
  let horario = null;
  
  Object.keys(namedValues).forEach(campo => {
    const campoLower = campo.toLowerCase();
    const valor = namedValues[campo][0];
    
    // Identificar email
    if (!email && (
      campoLower.includes('email') ||
      campoLower.includes('correo') ||
      campoLower.includes('helbide') ||
      campoLower.includes('dirección') ||
      (valor && valor.includes('@'))
    )) {
      email = valor;
    }
    
    // Identificar nombre
    if (!nombreFamilia && (
      campoLower.includes('nombre') ||
      campoLower.includes('izena') ||
      campoLower.includes('hijo') ||
      campoLower.includes('hija') ||
      campoLower.includes('alumno')
    )) {
      nombreFamilia = valor;
    }
    
    // Identificar horario
    if (!horario && (
      campoLower.includes('aukeratu') ||
      campoLower.includes('elige') ||
      campoLower.includes('día') ||
      campoLower.includes('egun') ||
      campoLower.includes('hora') ||
      campoLower.includes('ordu')
    )) {
      horario = valor;
    }
  });
  
  return { email, nombreFamilia: nombreFamilia || 'Familia', horario };
}

const ss = SpreadsheetApp.getActiveSpreadsheet();
const hojaConfig = ss.getSheetByName('Ordutegiak');
const formUrl = ss.getFormUrl();
/// ===== FUNCIÓN PRINCIPAL - PROCESAR ENVÍO DEL FORMULARIO =====
function onFormSubmit(e) {
  try {
    console.log('=== FORMULARIO ENVIADO ===');
    console.log('Timestamp:', new Date());
    
    // Obtener información del evento
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    console.log('Orria:', sheet.getName());
    console.log('Erantzunaren lerroa:', row);
    
    // Obtener los datos de la respuesta
    const datosRespuesta = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('Erantzunaren datuak:', datosRespuesta);
    const preguntaHorario = 'Aukeratu egun eta ordu bat / Elige un día y hora';
    const firma = hojaConfig.getRange("A6").getValue() || "";
    const form = FormApp.openByUrl(formUrl);
    
    // Obtener datos del formulario
    const { email, nombreFamilia, horario } = extraerDatosFormulario(e.namedValues);
    
    if (!horario || !email) {
      console.log("Datuak falta dira: ordutegia edo emaila");
      console.log("horario:", horario);
      console.log("email:", email);
      return;
    }
    
    // Enviar correo de confirmación
    enviarCorreoConfirmacion(email, horario, firma);

    // Recordatorio
    programarRecordatorio(email, nombreFamilia, horario);
    
    // Eliminar franja del formulario y hoja
    eliminarFranjaDelFormulario(form, preguntaHorario, horario);
    eliminarFranjaDeLaHoja(hojaConfig, horario);
    
    console.log(`Hitzordua zuzen prozesatu da ${email}: ${horario}`);
    
  } catch (error) {
    console.error("Error en onFormSubmit:", error);
    console.error("Stack completo:", error.stack);
  }
}

// ===== FUNCIÓN AUXILIAR: ENCONTRAR HOJA DE RESPUESTAS =====
function encontrarHojaRespuestas(spreadsheet) {
  const hojas = spreadsheet.getSheets();
  
  // Buscar hoja que contenga "respuestas" en el nombre
  let hojaRespuestas = hojas.find(hoja => 
    hoja.getName().toLowerCase().includes('Respuestas') ||
    hoja.getName().toLowerCase().includes('responses') ||
    hoja.getName().toLowerCase().includes('form')
  );
  
  // Si no encuentra por nombre, buscar por contenido típico de formulario
  if (!hojaRespuestas) {
    hojaRespuestas = hojas.find(hoja => {
      try {
        const headers = hoja.getRange(1, 1, 1, Math.min(10, hoja.getLastColumn())).getValues()[0];
        return headers.some(header => 
          header && (
            header.toString().includes('Timestamp') ||
            header.toString().includes('Marca temporal') ||
            header.toString().includes('Email') ||
            header.toString().includes('Aukeratu egun')
          )
        );
      } catch (e) {
        return false;
      }
    });
  }
  
  return hojaRespuestas;
}

// ===== FUNCIÓN AUXILIAR: ENCONTRAR COLUMNA =====
function encontrarColumna(headers, palabrasClave) {
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().toLowerCase();
    if (palabrasClave.some(palabra => header.includes(palabra.toLowerCase()))) {
      return i;
    }
  }
  return -1;
}

// ===== FUNCIÓN: CREAR HORARIOS FÁCILMENTE =====
function crearHorariosRapido() {
  const html = HtmlService.createHtmlOutputFromFile('CrearHorarios')
    .setTitle('Ordutegiak Sortu')
    .setWidth(850)           // ← Más ancho
    .setHeight(750);         // ← Más alto
  SpreadsheetApp.getUi().showModalDialog(html, 'Ordutegiak Sortu');  // ← Modal en lugar de sidebar
}
function borrarSelecciones() {
    // Desmarcar todos los checkboxes
    document.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
    
    // Vaciar campo de horas personalizadas
    document.getElementById('horasPersonalizadas').value = '';

    // Vaciar fechas
    document.getElementById('fechaInicio').value = hoy;
    document.getElementById('fechaFin').value = '';

    // Ocultar la vista previa
    document.getElementById('preview').style.display = 'none';
}

// ===== NUEVA FUNCIÓN: ORDENAR HORARIOS POR FECHA =====
function ordenarHorariosPorFecha() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordutegiak");
    
    // Obtener todos los horarios desde A8
    const rangoCompleto = hoja.getRange("A8:A1000");
    const todosLosDatos = rangoCompleto.getValues();
    
    // Filtrar solo las celdas que contienen datos válidos
    const horariosConFecha = [];
    
    todosLosDatos.forEach((fila, index) => {
      const horario = fila[0];
      if (horario && horario.toString().trim() !== '') {
        const fechaExtraida = extraerFechaDeCita(horario.toString());
        const horaExtraida = extraerHoraDeCita(horario.toString());
        
        if (fechaExtraida && horaExtraida) {
          // Crear fecha completa con hora
          const fechaCompleta = new Date(fechaExtraida);
          const [horas, minutos] = horaExtraida.split(':').map(Number);
          fechaCompleta.setHours(horas, minutos, 0, 0);
          
          horariosConFecha.push({
            texto: horario.toString(),
            fechaCompleta: fechaCompleta,
            filaOriginal: index + 8
          });
        } else {
          // Si no se puede extraer la fecha, mantener al final
          horariosConFecha.push({
            texto: horario.toString(),
            fechaCompleta: new Date('2099-12-31'), // Fecha muy lejana
            filaOriginal: index + 8
          });
        }
      }
    });
    
    // Ordenar por fecha y hora
    horariosConFecha.sort((a, b) => a.fechaCompleta.getTime() - b.fechaCompleta.getTime());
    
    // Limpiar el rango existente
    if (horariosConFecha.length > 0) {
      const ultimaFilaConDatos = Math.max(...horariosConFecha.map(item => item.filaOriginal));
      hoja.getRange(8, 1, ultimaFilaConDatos - 7, 1).clear();
      
      // Insertar los datos ordenados
      const datosOrdenados = horariosConFecha.map(item => [item.texto]);
      hoja.getRange(8, 1, datosOrdenados.length, 1).setValues(datosOrdenados);
    }
    
    console.log(`✅ ${horariosConFecha.length} Ordutegiak, daten arabera ordenatuta`);
    
  } catch (error) {
    console.error("Errorea ordutegiak ordenatzen:", error);
  }
}

// ===== FUNCIÓN AUXILIAR MEJORADA: EXTRAER HORA DE LA CITA =====
function extraerHoraDeCita(horario) {
  try {
    // Buscar patrón de hora HH:MM
    const patronHora = /(\d{1,2}:\d{2})/;
    const match = horario.match(patronHora);
    
    if (match) {
      return match[1];
    }
    
    return null;
  } catch (error) {
    console.error("Errores orria lortzen:", error);
    return null;
  }
}

// ===== FUNCIÓN AUXILIAR MEJORADA: EXTRAER FECHA (YA EXISTÍA PERO LA MEJORO) =====
function extraerFechaDeCita(horario) {
  try {
    const hoy = new Date();
    const año = hoy.getFullYear();
    
    // Mapas de meses en español
    const mesesEs = {
      'enero': 0, 'febrero': 1, 'marzo': 2, 'abril': 3, 'mayo': 4, 'junio': 5,
      'julio': 6, 'agosto': 7, 'septiembre': 8, 'octubre': 9, 'noviembre': 10, 'diciembre': 11
    };
    
    // Mapas de meses en euskera
    const mesesEu = {
      'urtarrila': 0, 'urtarrilak': 0, 'otsaila': 1, 'otsailak': 1, 'martxoa': 2, 'martxoak': 2, 
      'apirila': 3, 'apirilak': 3, 'maiatza': 4, 'maiatzak': 4, 'ekaina': 5, 'ekainak': 5,
      'uztaila': 6, 'uztailak': 6, 'abuztua': 7, 'abuztuak': 7, 'iraila': 8, 'irailak': 8, 
      'urria': 9, 'urriak': 9, 'azaroa': 10, 'azaroak': 10, 'abendua': 11, 'abenduak': 11
    };
    
    // Patrones de fecha más amplios
    const patronesFecha = [
      /(\d{1,2})\s+de\s+(\w+)/i,  // "15 de marzo"
      /(\d{1,2})\s+(\w+)/i,       // "15 marzo" o "15 martxoak"
      /(\w+)\s+(\d{1,2})/i        // "martxoak 15"
    ];
    
    for (let patron of patronesFecha) {
      const match = horario.match(patron);
      if (match) {
        let dia, mesNombre;
        
        // Determinar si el día está en la primera o segunda posición
        if (isNaN(match[1])) {
          // El mes está primero: "martxoak 15"
          mesNombre = match[1].toLowerCase();
          dia = parseInt(match[2]);
        } else {
          // El día está primero: "15 de marzo" o "15 marzo"
          dia = parseInt(match[1]);
          mesNombre = match[2].toLowerCase();
        }
        
        const mes = mesesEs[mesNombre] || mesesEu[mesNombre];
        
        if (mes !== undefined && dia >= 1 && dia <= 31) {
          let fechaResultado = new Date(año, mes, dia);
          
          // Si la fecha ya pasó este año, usar el próximo año
          if (fechaResultado < hoy) {
            fechaResultado = new Date(año + 1, mes, dia);
          }
          
          return fechaResultado;
        }
      }
    }
    
    return null;
  } catch (error) {
    console.error("Errorea data lortzen:", error);
    return null;
  }
}

// ===== NUEVA FUNCIÓN DE MENÚ: ORDENAR HORARIOS EXISTENTES =====
function ordenarHorariosExistentes() {
  try {
    ordenarHorariosPorFecha();
    SpreadsheetApp.getUi().alert("✅ Ordutegiak zuzen ordenatu dira dataren arabera.");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Errorea ordenatzen: ${error.message}`);
  }
}
// ===== FUNCIÓN PARA EL HTML DE CREAR HORARIOS (MODIFICADA) =====
function generarHorarios(fechaInicio, fechaFin, horasSeleccionadas, diasSemana, horasPersonalizadas) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordutegiak");
    
    const inicio = new Date(fechaInicio);
    const fin = new Date(fechaFin);
    const horarios = [];
    
    // Combinar horas predefinidas y personalizadas
    let todasLasHoras = [...horasSeleccionadas];
    if (horasPersonalizadas && horasPersonalizadas.trim()) {
      const personalizadas = horasPersonalizadas.split(',')
        .map(h => h.trim())
        .filter(h => h && /^\d{1,2}:\d{2}$/.test(h)); // Validar formato HH:MM
      todasLasHoras = [...todasLasHoras, ...personalizadas];
    }
    
    // Eliminar duplicados y ordenar
    todasLasHoras = [...new Set(todasLasHoras)].sort();
    
    const nombresDias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    const nombresDiasEu = ['igandea', 'astelehena', 'asteartea', 'asteazkena', 'osteguna', 'ostirala', 'larunbata'];
    const mesesEs = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    const mesesEu = ['Urtarrilak', 'Otsailak', 'Martxoak', 'Apirilak', 'Maiatzak', 'Ekainak', 'Uztailak', 'Abuztuak', 'Irailak', 'Urriak', 'Azaroak', 'Abenduak'];
    
    // Array temporal para ordenar por fecha
    const horariosConFecha = [];
    
    for (let fecha = new Date(inicio); fecha <= fin; fecha.setDate(fecha.getDate() + 1)) {
      const diaSemana = fecha.getDay();
      
      if (diasSemana.includes(diaSemana)) {
        const nombreDia = nombresDias[diaSemana];
        const nombreDiaEu = nombresDiasEu[diaSemana];
        const dia = fecha.getDate();
        const mes = mesesEs[fecha.getMonth()];
        const mesEu = mesesEu[fecha.getMonth()];
        
        todasLasHoras.forEach(hora => {
          const horario = `${nombreDia} ${dia} de ${mes} - ${hora} / ${mesEu} ${dia}, ${nombreDiaEu} - ${hora}`;
          
          // Crear fecha y hora completa para ordenar
          const fechaCompleta = new Date(fecha);
          const [horas, minutos] = hora.split(':').map(Number);
          fechaCompleta.setHours(horas, minutos, 0, 0);
          
          horariosConFecha.push({
            texto: horario,
            fechaCompleta: fechaCompleta
          });
        });
      }
    }
    
    // ORDENAR POR FECHA Y HORA
    horariosConFecha.sort((a, b) => a.fechaCompleta.getTime() - b.fechaCompleta.getTime());
    
    // Convertir a formato para insertar en la hoja
    const horariosParaInsertar = horariosConFecha.map(item => [item.texto]);
    
    if (horariosParaInsertar.length > 0) {
      // Encontrar la primera fila vacía después de A7
      let filaInicio = 8;
      const datos = hoja.getRange("A8:A").getValues();
      for (let i = datos.length - 1; i >= 0; i--) {
        if (datos[i][0] && datos[i][0].toString().trim() !== '') {
          filaInicio = i + 9;
          break;
        }
      }
      
      hoja.getRange(filaInicio, 1, horariosParaInsertar.length, 1).setValues(horariosParaInsertar);
      
      // ORDENAR TODA LA COLUMNA A PARTIR DE A8
      ordenarHorariosPorFecha();
    }
    
    return `✅ ${horariosParaInsertar.length} ordutegiak zuzen sortu dira eta dataren arabera ordenatu`;
    
  } catch (error) {
    console.error("Ordutegiak sortzen errorea:", error);
    return `❌ Errorea: ${error.message}`;
  }
}

// ===== FUNCIÓN AUXILIAR: ENVIAR RECORDATORIO INDIVIDUAL (SIN IMAGEN EXTERNA) =====
function enviarRecordatorio(email, familia, horario, firma) {
  const asunto = "🔔 Gogorarazpena: Bihar hitzordua / Recordatorio: Cita mañana";
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.5; max-width: 600px;">
      
      <!-- HEADER CON ESTILO EN VEZ DE IMAGEN -->
      <table width="100%" style="background-color: #2a5298; margin-bottom: 20px; border-radius: 12px;" cellpadding="0" cellspacing="0">
        <tr>
          <td style="padding: 30px; text-align: left;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align: top; width: 70%;">
                  <!-- Logo con emojis del centro -->
                  <div style="margin-bottom: 20px; font-size: 24px;">
                    🏔️☀️⛰️🧒🏽
                  </div>

                  <!-- Título principal -->
                  <h1 style="margin: 0 0 10px 0; font-size: 28px; font-weight: bold; color: #ffffff;">
                    MENDIALDEA IPI
                  </h1>
                  
                  <!-- Subtítulo en caja -->
                  <div style="background-color: rgba(255,255,255,0.2); padding: 12px 18px; border-radius: 20px; display: inline-block;">
                    <p style="margin: 0; font-size: 14px; color: #ffffff; font-weight: 500;">
                      GOGORARAZPENA<br>
                      <span style="font-size: 13px; opacity: 0.9;">RECORDATORIO</span>
                    </p>
                  </div>
                </td>
                
                <td style="vertical-align: middle; text-align: center; width: 30%;">
                  <!-- Elemento visual escuela -->
                  <div style="width: 90px; height: 80px; background-color: rgba(255,255,255,0.1); border-radius: 40px; border: 2px solid rgba(255,255,255,0.3); text-align: center; line-height: 76px; margin: 0 auto;">
                    <span style="font-size: 60px;">🔔</span>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <!-- Barra de colores -->
        <tr>
          <td style="height: 4px; background-color: #ffb74d;"></td>
        </tr>
      </table>

      <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <h3 style="color: #856404; margin-top: 0;"></h3>
        
        <div style="background-color: #fff; padding: 15px; border-radius: 6px; margin: 15px 0;">
          <p style="margin: 0;"><strong>📅 Zuen hitzordua bihar da / Vuestra cita es mañana:</strong></p>
          <p style="color: #2e86de; font-size: 18px; font-weight: bold; margin: 10px 0;">${horario}</p>
        </div>
        
        <p><strong>🫱🏾‍🫲🏿 </strong><br>
        Gogoratu bihar hitzordua dugula. Eskerrik asko!</p>
        
        <p><strong>🫱🏾‍🫲🏿 </strong><br>
        Recordad que mañana tenemos cita. ¡Muchas gracias!</p>
      </div>

      <div style="border-top: 1px solid #e9ecef; padding-top: 20px; font-size: 14px; color: #6c757d;">
        <p>${firma}</p>
      </div>
    </div>`;

  MailApp.sendEmail({
    to: email,
    subject: asunto,
    htmlBody: htmlBody
  });
}


// ===== CONFIGURAR RECORDATORIOS AUTOMÁTICOS =====
function configurarRecordatoriosAutomaticos() {
  // Eliminar triggers existentes
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarRecordatoriosCitas') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger para las 8:00 AM todos los días
  ScriptApp.newTrigger('enviarRecordatoriosCitas')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  SpreadsheetApp.getUi().alert("✅ Egunero 8:00etarako konfiguratu dira gogorarazpen automatikoak");
}

// ===== FUNCIÓN: ENVIAR CORREO DE CONFIRMACIÓN (SIN IMAGEN EXTERNA) =====
function enviarCorreoConfirmacion(email, horario, firma) {
  const asunto = "📩 Hitzordua baieztatu da / Cita confirmada";
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.5; max-width: 600px;">
      
      <!-- HEADER CON ESTILO EN VEZ DE IMAGEN -->
      <table width="100%" style="background-color: #2a5298; margin-bottom: 20px; border-radius: 12px;" cellpadding="0" cellspacing="0">
        <tr>
          <td style="padding: 30px; text-align: left;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align: top; width: 70%;">
                  <!-- Logo con emojis del centro -->
                  <div style="margin-bottom: 20px; font-size: 24px;">
                    🏔️☀️⛰️🧒🏽
                  </div>
                  
                  <!-- Título principal -->
                  <h1 style="margin: 0 0 10px 0; font-size: 28px; font-weight: bold; color: #ffffff;">
                    MENDIALDEA IPI
                  </h1>
                  
                  <!-- Subtítulo en caja -->
                  <div style="background-color: rgba(255,255,255,0.2); padding: 12px 18px; border-radius: 20px; display: inline-block;">
                    <p style="margin: 0; font-size: 14px; color: #ffffff; font-weight: 500;">
                      Zuen hitzordua baieztatu da<br>
                      <span style="font-size: 13px; opacity: 0.9;">Se ha confirmado vuestra cita</span>
                    </p>
                  </div>
                </td>
                
                <td style="vertical-align: middle; text-align: center; width: 30%;">
                  <!-- Elemento visual escuela -->
                  <div style="width: 90px; height: 80px; background-color: rgba(255,255,255,0.1); border-radius: 40px; border: 2px solid rgba(255,255,255,0.3); text-align: center; line-height: 76px; margin: 0 auto;">
                    <span style="font-size: 60px;">✅</span>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <!-- Barra de colores -->
        <tr>
          <td style="height: 4px; background-color: #ffb74d;"></td>
        </tr>
      </table>

      <div style="background-color: #f8f9fa; padding: 20px; border-left: 4px solid #2e86de; margin: 20px 0;">
        <p style="margin: 0;"><strong>📅 Zuen hitzordua / Vuestra cita:</strong></p>
        <p style="color: #2e86de; font-size: 18px; font-weight: bold; margin: 10px 0;">${horario}</p>
      </div>

      <div style="margin: 30px 0;">
        <p><strong>🫱🏾‍🫲🏿 </strong><br>
        Mila esker. Zuen zain egongo gara aipatutako egunean eta orduan.</p>
        
        <p><strong>🫱🏾‍🫲🏿 </strong><br>
        Muchas gracias. Os estaremos esperando el día y hora indicados.</p>
      </div>

      <div style="border-top: 1px solid #e9ecef; padding-top: 20px; font-size: 14px; color: #6c757d;">
        <p>${firma}</p>
      </div>
    </div>`;

  MailApp.sendEmail({
    to: email,
    subject: asunto,
    htmlBody: htmlBody
  });
}
function eliminarFranjaDelFormulario(form, preguntaHorario, horario) {
  const items = form.getItems(FormApp.ItemType.LIST);
  const listItem = items.find(item => item.getTitle() === preguntaHorario)?.asListItem();
  
  if (!listItem) {
    console.log("Ez da aurkitu ordutegiari buruzko galdera formularioan");
    return;
  }
  
  const nuevasOpciones = listItem.getChoices()
    .map(choice => choice.getValue())
    .filter(value => value !== horario);  // ← PROBLEMA: comparación estricta
    
  listItem.setChoices(nuevasOpciones.map(opt => listItem.createChoice(opt)));
}

// ===== FUNCIÓN: ELIMINAR FRANJA DEL FORMULARIO =====
//function eliminarFranjaDelFormulario(form, preguntaHorario, horario) {
//  const items = form.getItems(FormApp.ItemType.LIST);
//  const listItem = items.find(item => item.getTitle() === preguntaHorario)?.asListItem();
  
//  if (!listItem) {
//    console.log("Ez da aurkitu ordutegiari buruzko galdera formularioan");
//    return;
//  }
  
//  const nuevasOpciones = listItem.getChoices()
//    .map(choice => choice.getValue())
//    .filter(value => value !== horario);
    
//  listItem.setChoices(nuevasOpciones.map(opt => listItem.createChoice(opt)));
//}

// ===== FUNCIÓN: ELIMINAR FRANJA DE LA HOJA =====
//function eliminarFranjaDeLaHoja(hojaConfig, horario) {
//  const datos = hojaConfig.getRange("A8:A").getValues();
  
//  for (let i = 0; i < datos.length; i++) {
//    if (datos[i][0] === horario) {
//      hojaConfig.deleteRow(i + 8);
//      break;
//    }
//  }
//}
function eliminarFranjaDeLaHoja(hojaConfig, horario) {
  try {
    console.log('🗑️ Eliminando horario de la hoja:', horario);
    
    // Obtener el rango completo desde A8 hacia abajo
    const ultimaFila = hojaConfig.getLastRow();
    
    if (ultimaFila < 8) {
      console.log('⚠️ No hay datos desde A8 para eliminar');
      return false;
    }
    
    const datos = hojaConfig.getRange(8, 1, ultimaFila - 7, 1).getValues();
    console.log(`📊 Revisando ${datos.length} horarios desde la fila 8`);
    
    // Buscar el horario de atrás hacia adelante para evitar problemas de índices
    for (let i = datos.length - 1; i >= 0; i--) {
      const horarioEnFila = datos[i][0];
      
      if (horarioEnFila && horarioEnFila.toString().trim() === horario.toString().trim()) {
        const filaAEliminar = i + 8; // +8 porque empezamos desde A8
        
        console.log(`🎯 Horario encontrado en fila ${filaAEliminar}: "${horarioEnFila}"`);
        
        // Eliminar la fila
        hojaConfig.deleteRow(filaAEliminar);
        
        console.log(`✅ Fila ${filaAEliminar} eliminada correctamente`);
        return true;
      }
    }
    
    console.log(`⚠️ No se encontró el horario "${horario}" en la hoja`);
    return false;
    
  } catch (error) {
    console.error('❌ Error eliminando horario de la hoja:', error);
    return false;
  }
}
// ===== FUNCIÓN MEJORADA: ACTUALIZAR OPCIONES DESDE HOJA (ORDEN CRONOLÓGICO) =====
function actualizarOpcionesDesdeHoja() {
  try {
    console.log("🚀 Aukerak eguneratzen hasten...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordutegiak");
    
    if (!hoja) {
      throw new Error("Ez da aurkitu 'Ordutegiak' orria");
    }
    
    // Obtener URL del formulario
    const formUrl = hoja.getRange("A2").getValue();
    console.log("📋 Formulariaren URLa:", formUrl);
    
    if (!formUrl || formUrl.toString().trim() === '') {
      throw new Error("Inprimakiaren URLa ez da aurkitu A2 gelaxkan. Mesedez, konfiguratu ezazu lehenik menu hau erabiliz: ⚙️ Konfigurazioa > ✏️ Editatzeko esteka");
    }

    // Validar que la URL sea correcta
    const urlStr = formUrl.toString().trim();
    if (!urlStr.includes('docs.google.com/forms') && !urlStr.includes('/edit')) {
      throw new Error("A2ko URLak ez dirudi Googlen inprimaki bat editatzeko esteka balioduna denik. Eduki hau izan behar du: 'docs.google.com/forms' y '/edit'");
    }

    // Obtener horarios desde A8 hacia abajo - MANTENER EL ORDEN DE LA HOJA
    console.log("📅 Orritik ordutegiak lortzen...");
    const rangoCompleto = hoja.getRange("A8:A1000");
    const todosLosDatos = rangoCompleto.getValues();
    
    // Filtrar solo las celdas que contienen datos válidos - SIN ORDENAR ALFABÉTICAMENTE
    const franjas = todosLosDatos
      .flat()
      .map(valor => valor ? valor.toString().trim() : '')
      .filter(valor => valor !== '' && valor !== null && valor !== undefined);
    // ❌ ELIMINADO: .sort() - para mantener el orden cronológico de la hoja

    console.log(`📊 ${franjas.length} Ordutegi aurkituak (ordena kronologikoa mantenduta)`);
    console.log("🔍 Lehen 3 hitzorduak:", franjas.slice(0, 3));

    if (franjas.length === 0) {
      throw new Error("Orrian ez da ordutegirik aurkitu 8. ilaratik aurrera (A8tik beherantz). Mesedez, gehitu ordutegi batzuk lehendabizi.");
    }

    // Abrir el formulario
    console.log("🔗 Formularioa irekitzen...");
    const form = FormApp.openByUrl(urlStr);
    
    // Buscar la pregunta específica
    const preguntaObjetivo = 'Aukeratu egun eta ordu bat / Elige un día y hora';
    console.log("🔍 Buscando pregunta:", preguntaObjetivo);
    
    const itemsLista = form.getItems(FormApp.ItemType.LIST);
    console.log(`📋 Preguntas de tipo lista: ${itemsLista.length}`);
    
    const desplegable = itemsLista.find(item => item.getTitle().trim() === preguntaObjetivo);
    
    if (!desplegable) {
      // Si no encuentra la pregunta exacta, buscar parcialmente
      const desplegableAlternativo = itemsLista.find(item => 
        item.getTitle().includes('Aukeratu egun') || 
        item.getTitle().includes('Elige un día') ||
        item.getTitle().includes('horario') ||
        item.getTitle().includes('ordu')
      );
      
      if (!desplegableAlternativo) {
        const titulosDisponibles = itemsLista.map(item => `"${item.getTitle()}"`).join(', ');
        throw new Error(`Ez da "${preguntaObjetivo}" galdera formularioan aurkitu. Zerrendako galdera erabilgarriak: ${titulosDisponibles || 'ninguna'}. Egiaztatu formularioak aukera anitzeko/goitibeherako galdera bat duela izenburu zehatz horrekin.`);
      } else {
        console.log(`⚠️ Ez da galdera zehatza aurkitu, hau erabiliz: "${desplegableAlternativo.getTitle()}"`);
        actualizarOpcionesCronologico(desplegableAlternativo.asListItem(), franjas);
      }
    } else {
      console.log("✅ Galdera zuzen aurkitu da");
      actualizarOpcionesCronologico(desplegable.asListItem(), franjas);
    }

    // Mostrar resultado exitoso
    const mensaje = `✅ ${franjas.length} aukera eguneratu dira orden kronológikoa mantenduz`;
    console.log("🎉", mensaje);
    
    SpreadsheetApp.getActiveSpreadsheet()
      .toast(mensaje, "Eguneraketa eginda / Actualización completada", 5);
      
  } catch (error) {
    const mensajeError = `❌ Errorea: ${error.message}`;
    console.error("💥 Error en actualizarOpcionesDesdeHoja:", error);
    
    SpreadsheetApp.getActiveSpreadsheet()
      .toast(mensajeError, "Errorea", 10);
  }
}

//===== FUNCIÓN AUXILIAR: ACTUALIZAR OPCIONES MANTENIENDO ORDEN CRONOLÓGICO =====
function actualizarOpcionesCronologico(desplegable, franjas) {
  try {
    console.log("🔄 Goitibeherako menuaren aukerak eguneratzen (ordenean )...");
    
    // Crear las nuevas opciones MANTENIENDO EL ORDEN ORIGINAL
    const choices = franjas.map(franja => desplegable.createChoice(franja));
    console.log(`📋 ${choices.length} aukera berri sortzen orden kronologikoa mantenduz`);
    
    // Aplicar las opciones al desplegable
    desplegable.setChoices(choices);
    console.log("✅ Formularioari zuzen gehitu zaizkio aukerak orden cronológico egokian");
    
  } catch (error) {
    throw new Error(`Errorea goitibeherako menuaren aukerak eguneratzean: ${error.message}`);
  }
}

// ===== FUNCIÓN: VERIFICAR ORDEN EN FORMULARIO =====
function verificarOrdenFormulario() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("Ordutegiak");
    const formUrl = hoja.getRange("A2").getValue();
    
    if (!formUrl) {
      SpreadsheetApp.getUi().alert("❌ Ez da aurkitu formularioaren URLa A2 gelaxkan");
      return;
    }
    
    const form = FormApp.openByUrl(formUrl);
    const preguntaObjetivo = 'Aukeratu egun eta ordu bat';
    const itemsLista = form.getItems(FormApp.ItemType.LIST);
    const desplegable = itemsLista.find(item => item.getTitle().trim() === preguntaObjetivo);
    
    if (desplegable) {
      const opciones = desplegable.asListItem().getChoices().map(choice => choice.getValue());
      console.log("📋 Orden actual en el formulario:");
      opciones.forEach((opcion, i) => console.log(`${i+1}. ${opcion}`));
      
      SpreadsheetApp.getUi().alert(`✅ Egiaztapena eginda.\n\nFormularioak ${opciones.length} aukera ditu.\nKontsola berrikusi  ordena xehatua ikusteko..`);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Errorea berrikusten: ${error.message}`);
  }
}


// ===== FUNCIÓN: OBTENER LISTA DE EMAILS =====
function lortuEmailZerrenda() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Familien emailak");
    const datuak = hoja.getRange("A2:B").getValues();
    const zerrenda = [];

    datuak.forEach(([email, izena]) => {
      if (email && izena && email.toString().trim() !== '' && izena.toString().trim() !== '') {
        zerrenda.push({
          email: email.toString().trim(),
          label: izena.toString().trim()
        });
      }
    });

    zerrenda.sort((a, b) => {
      const abizenaA = a.label.split(" ").slice(-1)[0].toLowerCase();
      const abizenaB = b.label.split(" ").slice(-1)[0].toLowerCase();
      return abizenaA.localeCompare(abizenaB, 'es', { sensitivity: 'accent' });
    });

    return zerrenda;
    
  } catch (error) {
    console.error("Posta elektronikoak lortzen errorea:", error);
    return [];
  }
}
// ===== FUNCIÓN: MOSTRAR SELECTOR DE FAMILIAS =====
function erakutsiFamiliaAukeraketa() {
  const html = HtmlService.createHtmlOutputFromFile('FamiliaAukeraketa')
    .setTitle('Familiak Hautatu')
    .setWidth(650)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Familiak Hautatu');
}
// ===== FUNCIÓN: ENVIAR EMAILS A FAMILIAS SELECCIONADAS (SIN IMAGEN EXTERNA) =====
function bidaliHautatuetara(hautatuak) {
  if (!hautatuak || hautatuak.length === 0) {
    return "❌ Ez da familiarik hautatu";
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordutegiakHoja = ss.getSheetByName("Ordutegiak");
    const galdetegiEsteka = ordutegiakHoja.getRange("A4").getValue();
    const sinadura = ordutegiakHoja.getRange("A6").getValue() || "";
    
    if (!galdetegiEsteka) {
      return "❌ Ez da aurkitu galdetegiaren esteka A4 gelaxkan";
    }
    
    const htmlMezua = `
    <div style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.6; max-width: 600px;">
      <!-- HEADER COMPATIBLE CON GMAIL -->
      <table width="100%" style="background-color: #2a5298; margin-bottom: 20px; border-radius: 12px;" cellpadding="0" cellspacing="0">
        <tr>
          <td style="padding: 30px; text-align: left;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align: top; width: 70%;">
                  <!-- Logo con emojis del centro -->
                  <div style="margin-bottom: 20px; font-size: 24px;">
                    🏔️☀️⛰️🧒🏽
                  </div>
                  
                  <!-- Título principal -->
                  <h1 style="margin: 0 0 10px 0; font-size: 28px; font-weight: bold; color: #ffffff;">
                    MENDIALDEA IPI
                  </h1>
                  
                  <!-- Subtítulo en caja -->
                  <div style="background-color: rgba(255,255,255,0.2); padding: 12px 18px; border-radius: 20px; display: inline-block;">
                    <p style="margin: 0; font-size: 14px; color: #ffffff; font-weight: 500;">
                      Hitzordu-eskaera tutoretzarako<br>
                      <span style="font-size: 13px; opacity: 0.9;">Solicitud de cita para tutoría</span>
                    </p>
                  </div>
                </td>
                
                <td style="vertical-align: middle; text-align: center; width: 30%;">
                  <!-- Elemento visual escuela -->
                  <div style="width: 90px; height: 80px; background-color: rgba(255,255,255,0.1); border-radius: 40px; border: 2px solid rgba(255,255,255,0.3); text-align: center; line-height: 76px; margin: 0 auto;">
                    <span style="font-size: 60px;">🗓️</span>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <!-- Barra de colores -->
        <tr>
          <td style="height: 4px; background-color: #ffb74d;"></td>
        </tr>
      </table>
      
      <!-- CONTENIDO RESTO -->
      <div style="background-color: #e3f2fd; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <h3 style="color: #1976d2; margin-top: 0;">🏫 Kaixo, familia!</h3>
        <p>Azpiko estekan, zuen seme-alabaren tutorearekin hitzordua eskatzeko galdetegi bat duzue. Aukeratu egokien zaizuen eguna eta ordua.</p>
        <div style="text-align: center; margin: 20px 0;">
          <a href="${galdetegiEsteka}" 
             style="background-color: #1976d2; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold;">
             📝 Sartu galdetegira
          </a>
        </div>
      </div>
      
      <div style="background-color: #f3e5f5; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <h3 style="color: #7b1fa2; margin-top: 0;">🏫 ¡Hola, familia!</h3>
        <p>A continuación encontraréis un formulario para pedir cita con el tutor/a de vuestro hijo o hija. Podéis elegir el día y la hora que mejor os convenga.</p>
        <div style="text-align: center; margin: 20px 0;">
          <a href="${galdetegiEsteka}" 
             style="background-color: #7b1fa2; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold;">
             📝 Acceder al formulario
          </a>
        </div>
      </div>
      
      <div style="border-top: 1px solid #e0e0e0; padding-top: 20px; font-size: 14px; color: #666;">
        <p><em>${sinadura}</em></p>
      </div>
    </div>`;
    
    const gaia = "📩 Hitzordua eskatzeko galdetegia / Formulario para pedir cita";
    let bidalitakoak = 0;
    
    hautatuak.forEach(email => {
      try {
        MailApp.sendEmail({
          to: email,
          subject: gaia,
          htmlBody: htmlMezua
        });
        bidalitakoak++;
      } catch (error) {
        console.error(`${email}-ri posta elektronikoa bidaltzean errorea:`, error);
      }
    });
    
    return `✅ ${bidalitakoak}/${hautatuak.length} mezu bidali dira`;
    
  } catch (error) {
    console.error("Error en bidaliHautatuetara:", error);
    return `❌ Errorea mezuak bidaltzean: ${error.message}`;
  }
}

// ===== FUNCIÓN: CREAR MENÚ AL ABRIR (MODIFICADA) =====
function onOpen() {
  const menu = SpreadsheetApp.getUi()
    .createMenu("🏫 Tutoretza Galdetegia")
    .addItem("📩 Mezuak bidali familiei", "eguneratuEtaBidali")
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("⏰ Ordutegia kudeatu")
      .addItem("⏰ Ordutegia sortu", "crearHorariosRapido")
      .addItem("🔄 Hitzordu-aukerak eguneratu", "actualizarOpcionesDesdeHoja")
      .addItem('❌ Aukerak ezabatu', 'borrarTodoCompleto')
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("🔔 Gogorarazpenak")
      .addItem("📤 Orain bidali gogorarazpenak", "enviarRecordatoriosCitas")
      .addItem("⚙️ Gogorarazpen automatikoak", "configurarRecordatoriosAutomaticos")
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("📅 Calendar sinkronizazioa")
      .addItem("📤 Orain sinkronizatu", "ejecutarSincronizacionManual")
      .addItem("⚙️ Sinkronizazio automatikoa", "configurarSincronizacion")
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu("⚙️ Konfigurazioa")
      .addItem("📝 Sinadura ezarri", "galdetuSinadura")
      .addItem("📋 Estekak eguneratu", "pegarAmbasURLs")  
      .addItem('📥 Tutoretza taldea aldatu', 'seleccionarClaseEImportarEmails')
      .addItem('🪢 Sheet-a eta galdetegia lotu', 'vincularSheetConFormulario')
      .addItem('💻Aktibatzailea sortu','configurarActivador')
      .addItem('🔗Estekak eguneratu','pegarAmbasURLs')
      .addItem('1️⃣Lehen konfigurazioa', 'primeraConfiguracion'))
    ;
    
  menu.addToUi();
}
// ===== FUNCIÓN: ACTUALIZAR Y ENVIAR =====
function eguneratuEtaBidali() {
  actualizarOpcionesDesdeHoja();
  Utilities.sleep(1000);
  erakutsiFamiliaAukeraketa();
}
//LEHEN KONFIGURAZIOA
function primeraConfiguracion() {
  var ui = SpreadsheetApp.getUi();
  
  // Hasierako mezua
  ui.alert(
    "Lehen konfigurazioa",
    "Programa honetan sartu zara lehen konfigurazioa egiteko.\n" +
    "Mendialdea IPIko tutoretza bileren kudeaketa automatikoa abiaraziko duzu.\n" +
    "Konfigurazioa pausoz pauso egingo da.",
    ui.ButtonSet.OK
  );
  
  // Pausoen zerrenda (izenak eta funtzioak)
  var pasos = [
    {nombre: "1/7: Sheeta eta galdetegia lotu", funcion: vincularSheetConFormulario},
    {nombre: "2/7: Aktibatzailea sortu", funcion: configurarActivador},
    {nombre: "3/7: Estekak itsatsi", funcion: pegarAmbasURLs},
    {nombre: "4/7: Sinadura ezarri", funcion: galdetuSinadura},
    {nombre: "5/7: Tutoretza gela aukeratu", funcion: seleccionarClaseEImportarEmails},
    {nombre: "6/7: Gogorarazpenak konfiguratu", funcion: configurarRecordatoriosAutomaticos},
    {nombre: "7/7: Calendarrekin sinkronizazioa", funcion: configurarSincronizacion}
  ];
  
  var currentStep = 0;
  
  try {
    for (currentStep = 0; currentStep < pasos.length; currentStep++) {
      var paso = pasos[currentStep];
      
      // Pausoaren hasiera jakinarazi
      ui.alert("Exekutatzen: " + paso.nombre);
      
      // Funtzioa exekutatu
      paso.funcion();
    }
  } catch (e) {
    ui.alert(
      "ERROREA",
      "Konfigurazioa eten egin da.\n" +
      "Pausoan huts egin du: " + pasos[currentStep].nombre + "\n" +
      "Xehetasunak: " + e.message,
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Amaierako mezua
  ui.alert(
    "Konfigurazioa amaituta",
    "Lehen konfigurazioa behar bezala egin da.\n" +
    "Gogorarazpenak automatikoki bidaliko dira egunero goizeko 8:00etan.",
    ui.ButtonSet.OK
  );
}

//Vincular el Google Sheet actual con un Google Form del mismo nombre en la misma carpeta
function vincularSheetConFormulario() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const archivoSheet = DriveApp.getFileById(ss.getId());
    const nombreSheet = archivoSheet.getName();
    
    console.log('Sheetarentzako galdetegia aurkitzen:', nombreSheet);
    
    // Obtener la carpeta donde está el sheet
    const carpetas = archivoSheet.getParents();
    
    if (!carpetas.hasNext()) {
      throw new Error('Ezin daiteke spreadsheetaren karpeta aurkitu');
    }
    
    const carpeta = carpetas.next();
    console.log('Karpeta aurkitu da:', carpeta.getName());
    
    // Buscar formulario con el mismo nombre en la misma carpeta
    const archivosEnCarpeta = carpeta.getFilesByType(MimeType.GOOGLE_FORMS);
    let formularioEncontrado = null;
    
    while (archivosEnCarpeta.hasNext()) {
      const archivo = archivosEnCarpeta.next();
      if (archivo.getName() === nombreSheet) {
        formularioEncontrado = archivo;
        break;
      }
    }
    
    if (!formularioEncontrado) {
      throw new Error(`Ez da aurkitu karpeta berean "${nombreSheet}" izena duen galdetegirik`);
    }
    
    console.log('Galdetegia aurkitu da:', formularioEncontrado.getName());
    
    // Abrir el formulario
    const form = FormApp.openById(formularioEncontrado.getId());
    
    // Verificar estado actual del destino
    let destinationId = null;
    try {
      destinationId = form.getDestinationId();
      console.log('Galdetegiaren egungo kokapena:', destinationId);
    } catch (e) {
      console.log('Galdetegiak ez du kokapenik:', e.message);
    }
    
    // Vincular formulario con el spreadsheet SIEMPRE (incluso si ya está vinculado)
    console.log('Sheeta galdetegira lotzen...');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    console.log('✅ Sheeta galdetgita zuzen lotu da.');
    
    // Verificar vinculación
    const nuevoDestino = form.getDestinationId();
    console.log('Kokapen berria konfiguratua:', nuevoDestino);
    
    if (nuevoDestino !== ss.getId()) {
      throw new Error('Loturak errorea- kokapenek ez dute kontziditzen');
    }
    
    // Obtener URL del formulario
    const urlFormulario = form.getPublishedUrl();
    console.log('Galdetegiaren URLa lortu da:', urlFormulario);
    
    // Mensaje de éxito
    const mensaje = `✅ LOTURA ZUZENA

Galdetegia aurkitu da: ${formularioEncontrado.getName()}
Sheeta: ${nombreSheet}
Destino ID: ${nuevoDestino}

URL del formulario:
${urlFormulario}

Galdetegia behar bezala lotuta dago sheet honetan erantzunak jasotzeko.`;
    
    SpreadsheetApp.getUi().alert('Lotura egin da', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return {
      success: true,
      formularioId: formularioEncontrado.getId(),
      urlFormulario: urlFormulario,
      destinationId: nuevoDestino
    };
    
  } catch (error) {
    console.error('Errorea galdetegia lotzen:', error);
    
    // Información adicional para debug
    console.error('Errorearen detalleak:');
    console.error('- Izena:', error.name);
    console.error('- Mezua:', error.message);
    console.error('- Stack:', error.stack);
    
    SpreadsheetApp.getUi().alert('Error', 'Errorea galdetegia lotzen: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

//Configurar activador/trigger para ejecutar función al enviar formulario
function configurarActivador() {
  console.log('=== AKTIBATZAILEAREN KONFIGURAZIOA HASTEN ===');
  
  try {
    console.log('1. PAUSUA: kalkulu orria lortzen...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✅ Kalkulu orria lortu da:', ss.getName());
    console.log('✅ ID:', ss.getId());
    
    console.log('2. PAUSUA: aurretiko aktibatzaileak lortzen...');
    const triggersExistentes = ScriptApp.getProjectTriggers();
    console.log('✅ Aktibatzaile aurkituak:', triggersExistentes.length);
    
    let triggersEliminados = 0;
    console.log('3. PAUSUA: Aurretiko aktibadoreak ezabatzen...');
    
    triggersExistentes.forEach((trigger, index) => {
      console.log(`Revisando trigger ${index + 1}:`, {
        tipo: trigger.getEventType(),
        funcion: trigger.getHandlerFunction(),
        id: trigger.getUniqueId()
      });
      
      if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
        console.log('Galdetegiko aktibatzailea ezabatzen...');
        ScriptApp.deleteTrigger(trigger);
        triggersEliminados++;
        console.log('✅ Aktibatzaile ezabatua:', trigger.getUniqueId());
      }
    });
    
    console.log(`Aktibatzaile ezabatua: ${triggersEliminados}`);
    
    console.log('4 PAUSUA: Aktibatzaile berria sortzen...');
    console.log('Aktibatzaileraren konfigurazioa');
    console.log('- Función: onFormSubmit');
    console.log('- Hoja:', ss.getName());
    console.log('- Tipo: ON_FORM_SUBMIT');
    
    // Crear el trigger para la función onFormSubmit
    const trigger = ScriptApp.newTrigger('onFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();
    
    console.log('✅ Aktibatzailea sortu da!');
    
    console.log('Detalles del trigger creado:');
    console.log('- ID:', trigger.getUniqueId());
    console.log('- Función:', trigger.getHandlerFunction());
    console.log('- Tipo de evento:', trigger.getEventType());
    console.log('- Fuente:', trigger.getTriggerSourceId());
    
    console.log('5. PAUSUA: Baieztapena erakusten...');
    SpreadsheetApp.getUi().alert(
      'Aktibatzaile konfiguratua', 
      `✅ Aktibatzailea egoki sortu da:

• Función a ejecutar: onFormSubmit
• Hoja de cálculo: ${ss.getName()}
• Tipo de evento: Al enviarse el formulario
• ID del trigger: ${trigger.getUniqueId()}

Aktibatzailea aktibo dago eta funtzioa automatikoki exekutatuko du lotutako galdetegiari erantzuten zaionean.

GARRANTZITSUA: Ziurtatu 'onFormSubmit' funtzioa dagoela zure proiektuan.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log('=== ONDO OSATUTAKO KONFIGURAZIOA ===');
    return trigger.getUniqueId();
    
  } catch (error) {
    console.error('=== ERROREA KONFIGURATZEN ===');
    console.error('Tipo de error:', error.name);
    console.error('Mensaje:', error.message);
    console.error('Stack trace:', error.stack);
    console.error('================================');
    
    SpreadsheetApp.getUi().alert(
      'Errorea aktibatzailea konfiguratzean', 
      `❌ Error detallado:

Tipo: ${error.name}
Mensaje: ${error.message}

Begiratu kontsola (Ikusi > Erregistroak) informazio tekniko gehiagorako.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    throw error;
  }
}

// Función que pega URLs del formulario en A2 (editable) y A4 (público) y asegura que esté publicado
function pegarAmbasURLs() {
  try {
    console.log('=== GALDETEGIAREN ESTEKAK KONFIGURATZEN ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const urlFormulario = ss.getFormUrl();
    
    if (!urlFormulario) {
      throw new Error('Ez da aurkitu sheet honi lotutako galdetegirik');
    }
    
    // Obtener el ID del formulario de la URL
    const formId = urlFormulario.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/)[1];
    console.log('ID del formulario:', formId);
    
    // Abrir el formulario para verificar y cambiar configuración
    const form = FormApp.openById(formId);
    console.log('Galdetegi irekia:', form.getTitle());
    
    // PASO 1: PUBLICAR el formulario primero (crítico)
    console.log('📤 Galdetegia publikatzen...');
    form.setAcceptingResponses(true);
    console.log('✅ Formulario publicado y acepta respuestas');
    
    // PASO 2: Configuraciones de acceso público
    console.log('🔓 Configurando acceso público...');
    
    
    try {
      // Configuraciones adicionales (si están disponibles)
      form.setAllowResponseEdits(false); // No permitir editar respuestas
      console.log('✅ Edición de respuestas deshabilitada');
    } catch (editError) {
      console.warn('⚠️ No se pudo configurar edición:', editError.message);
    }
    
    // PASO 3: Verificar estado final
    const estadoFinal = form.isAcceptingResponses();
    let loginFinal = null;
    
    try {
      loginFinal = form.requiresLogin();
    } catch (e) {
      console.warn('⚠️ No se pudo verificar requisito de login');
    }
    
    console.log('Estado final del formulario:');
    console.log('- Acepta respuestas:', estadoFinal);
    console.log('- Requiere login:', loginFinal);
    
    if (!estadoFinal) {
      throw new Error('El formulario no se pudo publicar correctamente');
    }
    
    // PASO 4: Crear URLs y pegarlas
    const urlEditable = urlFormulario.replace('/viewform', '/edit');
    
    // Buscar o crear hoja Ordutegiak
    let hoja = ss.getSheetByName('Ordutegiak');
    if (!hoja) {
      hoja = ss.insertSheet('Ordutegiak');
      console.log('Hoja Ordutegiak creada');
    }
    
    // Pegar ambas URLs en la hoja
    hoja.getRange('A2').setValue(urlEditable);  // URL editable
    hoja.getRange('A4').setValue(urlFormulario); // URL público
    
    console.log('✅ URLs pegadas:');
    console.log('- A2 (Editable):', urlEditable);
    console.log('- A4 (Público):', urlFormulario);
    
    // PASO 5: Mensaje de confirmación
    const loginStatus = loginFinal === null ? 'No verificado' : (loginFinal ? 'SÍ requiere' : 'NO requiere');
    
    SpreadsheetApp.getUi().alert(
      'URLs Configuradas', 
      `✅ Formulario configurado y URLs colocadas:

📝 URL Editable (A2):
${urlEditable}

👥 URL Público (A4):
${urlFormulario}

✅ Estado del formulario:
   - Acepta respuestas: ${estadoFinal}
   - Login: ${loginStatus}
   - Listo para compartir`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return { 
      success: true, 
      publicado: estadoFinal, 
      loginRequired: loginFinal,
      urls: { editable: urlEditable, publico: urlFormulario } 
    };
    
  } catch (error) {
    console.error('❌ Error en pegarAmbasURLs:', error);
    console.error('Detalles completos:', {
      name: error.name,
      message: error.message,
      stack: error.stack
    });
    
    SpreadsheetApp.getUi().alert(
      'Error', 
      `Error configurando URLs: ${error.message}\n\nRevisa los logs para más detalles.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

// ============================================
// SISTEMA DE IMPORTACIÓN DE EMAILS POR CLASE
// ============================================

// URL del Google Sheet público con los emails
const SHEET_PUBLICO_URL = "https://docs.google.com/spreadsheets/d/1w8VORB8b9mYv0UWJ2aIsiCRJTKJo71ReNHjX4laEx0A/edit?usp=sharing";
const HOJA_DESTINO = "Familien emailak"; // Nombre de la hoja donde se pegarán los emails

//Función principal para seleccionar e importar emails de una clase
function seleccionarClaseEImportarEmails() {
  try {
    // Obtener lista de clases (nombres de las pestañas)
    const clases = obtenerListaClases();
    
    if (clases.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'Ezin izan dira sheet publikoko gelak lortu.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Mostrar selector de clases al profesor
    const claseSeleccionada = mostrarSelectorClases(clases);
    
    if (claseSeleccionada) {
      // Importar emails de la clase seleccionada
      importarEmailsDeClase(claseSeleccionada);
    }
    
  } catch (error) {
    console.error('Errorea hautaketan eta inportazioan:', error);
    SpreadsheetApp.getUi().alert('Error', 'Errorea gertatu da: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

//Obtener la lista de clases (nombres de las pestañas) del sheet público
function obtenerListaClases() {
  try {
    console.log('Sheet publikoko gelen zerrenda lortzen...');
    
    // Abrir el sheet público
    const sheetPublico = SpreadsheetApp.openByUrl(SHEET_PUBLICO_URL);
    
    // Obtener todos los nombres de las hojas/pestañas
    const todasLasHojas = sheetPublico.getSheets();
    const nombresClases = [];
    
    todasLasHojas.forEach(hoja => {
      const nombreHoja = hoja.getName();
      // Filtrar hojas que puedan ser clases
      if (nombreHoja && !nombreHoja.toLowerCase().includes('config') && !nombreHoja.toLowerCase().includes('admin')) {
        nombresClases.push(nombreHoja);
      }
    });
    
    console.log('Gela aurkituak', nombresClases);
    return nombresClases.sort(); // Ordenar alfabéticamente
    
  } catch (error) {
    console.error('Errorea gelen zerrenda lortzen:', error);
    throw new Error('Ezin izan da sheet publikora sartu. Egiaztatu URLa zuzena dela eta sheeta publikoa dela.');
  }
}

//Mostrar selector de clases al profesor
function mostrarSelectorClases(clases) {
  const ui = SpreadsheetApp.getUi();
  
  // Crear mensaje con lista numerada de clases
  let mensaje = 'AUKERATU ZURE GELA:\n\n';
  clases.forEach((clase, index) => {
    mensaje += `${index + 1}. ${clase}\n`;
  });
  mensaje += '\nIDATZI ZURE GELAREN ZENBAKIA:';
  
  // Mostrar prompt al usuario
  const response = ui.prompt(
    'Gela aukeratzea',
    mensaje,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const numeroSeleccionado = parseInt(response.getResponseText().trim());
    
    if (numeroSeleccionado >= 1 && numeroSeleccionado <= clases.length) {
      const claseSeleccionada = clases[numeroSeleccionado - 1];
      console.log('Gela aukeratua:', claseSeleccionada);
      return claseSeleccionada;
    } else {
      ui.alert('Error', 'Zenbaki baliogabea. Zenbakiak 1 eta ' + clases.length + ' artean egon behar du.', ui.ButtonSet.OK);
      return null;
    }
  }
  
  return null; // Usuario canceló
}

//Importar emails y nombres de una clase específica
function importarEmailsDeClase(nombreClase) {
  try {
    console.log('Gelako emailak importatzen:', nombreClase);
    
    // Abrir el sheet público
    const sheetPublico = SpreadsheetApp.openByUrl(SHEET_PUBLICO_URL);
    
    // Obtener la hoja de la clase seleccionada
    const hojaClase = sheetPublico.getSheetByName(nombreClase);
    
    if (!hojaClase) {
      throw new Error(`Ez da orria aurkitu: ${nombreClase}`);
    }
    
    // Verificar que existe la celda A1 con "Member email"
    const tituloA1 = hojaClase.getRange('A1').getValue();
    console.log('Títuloa A1ean:', tituloA1);
    
    if (!tituloA1 || !tituloA1.toString().toLowerCase().includes('email')) {
      SpreadsheetApp.getUi().alert('Advertencia', `Klasearen A1 gelaxkak "${nombreClase}" ez du "Member email". Gela zuzena al da?`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    // Obtener emails (columna A, desde A2 hacia abajo)
    const ultimaFilaA = hojaClase.getLastRow();
    let emails = [];
    let nombres = [];
    
    if (ultimaFilaA > 1) {
      // Obtener emails (columna A desde A2)
      const rangoEmails = hojaClase.getRange(2, 1, ultimaFilaA - 1, 1);
      const datosEmails = rangoEmails.getValues();
      
      // Obtener nombres (columna B desde B2)
      const rangoNombres = hojaClase.getRange(2, 2, ultimaFilaA - 1, 1);
      const datosNombres = rangoNombres.getValues();
      
      // Filtrar emails válidos y sus nombres correspondientes
      for (let i = 0; i < datosEmails.length; i++) {
        const email = datosEmails[i][0];
        const nombre = datosNombres[i] ? datosNombres[i][0] : '';
        
        if (email && email.toString().trim() !== '' && email.toString().includes('@')) {
          emails.push(email.toString().trim());
          nombres.push(nombre ? nombre.toString().trim() : '');
        }
      }
    }
    
    if (emails.length === 0) {
      throw new Error(`Ez da emailik aurkitu "${nombreClase}"`);
    }
    
    console.log(`${emails.length} email aurkitu dira.`);
    
    // Pegar en la hoja destino
    pegarEmailsEnHojaDestino(nombreClase, emails, nombres);
    
  } catch (error) {
    console.error('Errorea emailak inportatzean:', error);
    throw error;
  }
}

//Pegar los emails y nombres en la hoja "Familien emailak"
function pegarEmailsEnHojaDestino(nombreClase, emails, nombres) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Buscar o crear la hoja destino
    let hojaDestino = ss.getSheetByName(HOJA_DESTINO);
    
    if (!hojaDestino) {
      hojaDestino = ss.insertSheet(HOJA_DESTINO);
      console.log('Orria "' + HOJA_DESTINO + '" sortu da');
    }
    
    // Limpiar contenido anterior (opcional - pregunta al usuario)
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert(
      'Aurreko datuak garbitu',
      'Nahi al duzu orri honetako datuak garbitu "' + HOJA_DESTINO + '"?',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (respuesta === ui.Button.CANCEL) {
      return; // Usuario canceló
    }
    
    if (respuesta === ui.Button.YES) {
      hojaDestino.clear();
      console.log('Orria garbitu da');
    }
    
    // Determinar dónde empezar a escribir
    let filaInicio = hojaDestino.getLastRow() + 1;
    
    if (filaInicio === 1 || respuesta === ui.Button.YES) {
      // Crear headers
      hojaDestino.getRange(1, 1, 1, 2).setValues([['Email', 'Nombre y Apellidos']]);
      hojaDestino.getRange(1, 1, 1, 2).setBackground('#4CAF50').setFontColor('white').setFontWeight('bold');
      filaInicio = 2;
    }
    
    // Preparar datos para insertar
    const datosParaInsertar = [];
    for (let i = 0; i < emails.length; i++) {
      datosParaInsertar.push([
        emails[i],            // Columna A: Email
        nombres[i] || ''      // Columna B: Nombre y Apellidos
      ]);
    }
    
    // Insertar los datos
    const rangoDestino = hojaDestino.getRange(filaInicio, 1, datosParaInsertar.length, 2);
    rangoDestino.setValues(datosParaInsertar);
    
    // Formatear
    hojaDestino.autoResizeColumns(1, 2);
    
    // Mensaje de éxito
    const mensaje = `✅ INPORTAZIOA GAUZATU DA

Gela: ${nombreClase}
Email inportatuak: ${emails.length}
Helburu-orria: ${HOJA_DESTINO}
Hasierako ilara: ${filaInicio}

Emailak behar bezala itsatsi dira eta erabiltzeko prest daude.`;
    
    ui.alert('Inportazioa gauzatu da.', mensaje, ui.ButtonSet.OK);
    
    console.log('Inportazio arrakastatsua');
    
  } catch (error) {
    console.error('Errorea emailak itsastean:', error);
    throw error;
  }
}

function galdetuSinadura() {
  const ui = SpreadsheetApp.getUi();
  const erantzuna = ui.prompt(
    "Konfigurazioa",
    "Sartu zure izen-abizenak eta tutoretza maila sinadurarako:",
    ui.ButtonSet.OK_CANCEL
  );

  if (erantzuna.getSelectedButton() == ui.Button.OK) {
    const sinadura = erantzuna.getResponseText().trim();
    if (sinadura) {
      // Buscar o crear hoja Ordutegiak
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName("Ordutegiak");
      if (!sheet) {
        sheet = ss.insertSheet("Ordutegiak");
        console.log('Hoja Ordutegiak creada para sinadura');
      }
      sheet.getRange("A6").setValue(sinadura);
      ui.alert("✅ Sinadura gorde da");
    }
  }
}
// Funciones auxiliares
function borrarCeldasA8() {
  try {
    console.log('🔄 Borrarceldas exekutatzen...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaOrdutegiak = spreadsheet.getSheetByName('Ordutegiak');
    
    if (hojaOrdutegiak) {
      const ultimaFila = hojaOrdutegiak.getLastRow();
      
      if (ultimaFila >= 8) {
        const rango = hojaOrdutegiak.getRange('A8:A' + ultimaFila);
        rango.clearContent();
        console.log('✅ borrarCeldasA8 completado');
        return true;
      } else {
        console.log('⚠️ No hay datos para borrar desde A8');
        return true;
      }
    } else {
      throw new Error('No se encontró la hoja Ordutegiak');
    }
    
  } catch (error) {
    console.log('❌ Error en borrarCeldasA8: ' + error.toString());
    throw error;
  }
}

function borrarRespuestasHoja() {
  try {
    console.log('🔄 Ejecutando borrarRespuestasHoja...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const todasLasHojas = spreadsheet.getSheets();
    let hojaEncontrada = false;
    
    for (let i = 0; i < todasLasHojas.length; i++) {
      const nombreHoja = todasLasHojas[i].getName();
      
      if (nombreHoja.startsWith('Respuestas')) {
        const hojaRespuestas = todasLasHojas[i];
        const ultimaFila = hojaRespuestas.getLastRow();
        
        if (ultimaFila > 1) {
          hojaRespuestas.deleteRows(2, ultimaFila - 1);
          console.log('✅ Respuestas de hoja eliminadas');
        } else {
          console.log('⚠️ No hay respuestas para eliminar');
        }
        
        hojaEncontrada = true;
        break;
      }
    }
    
    if (!hojaEncontrada) {
      throw new Error('No se encontró hoja que empiece con "Respuestas"');
    }
    
    console.log('✅ borrarRespuestasHoja completado');
    return true;
    
  } catch (error) {
    console.log('❌ Error en borrarRespuestasHoja: ' + error.toString());
    throw error;
  }
}

function borrarRespuestasForm() {
  try {
    console.log('🔄 Ejecutando borrarRespuestasForm...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const formUrl = spreadsheet.getFormUrl();
    
    if (formUrl) {
      const formulario = FormApp.openByUrl(formUrl);
      const respuestas = formulario.getResponses();
      const numRespuestas = respuestas.length;
      
      formulario.deleteAllResponses();
      console.log('✅ ' + numRespuestas + ' respuestas del formulario eliminadas');
    } else {
      console.log('⚠️ No hay formulario vinculado');
    }
    
    console.log('✅ borrarRespuestasForm completado');
    return true;
    
  } catch (error) {
    console.log('❌ Error en borrarRespuestasForm: ' + error.toString());
    throw error;
  }
}

// Función principal corregida
function borrarTodoCompleto() {
  console.log('🚀 Iniciando borrarTodoCompleto...');
  
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Konfirmazioa', 
    'FUNTZIO HAU BILERA GUZTIAK BUKATU ONDOREN ABIARAZTEA KOMENI DA. Hitzordu-aukera eta erantzun guztiak ezabatu nahi dituzu?',
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta !== ui.Button.YES) {
    console.log('❌ Usuario canceló la operación');
    return;
  }
  
  try {
    console.log('🔄 Ejecutando funciones de borrado...');
    borrarCeldasA8();
    borrarRespuestasHoja();
    borrarRespuestasForm();
    
    console.log('✅ Todas las operaciones completadas exitosamente');
    ui.alert('Eginda', 'Dena ondo ezabatu da!', ui.ButtonSet.OK);
    
  } catch (error) {
    console.log('❌ Error en borrarTodoCompleto: ' + error.toString());
    ui.alert('Errorea', 'Errore bat gertatu da: ' + error.toString(), ui.ButtonSet.OK);
  }
}
// ===== CORRECCIÓN 2: FUNCIÓN PROGRAMAR RECORDATORIO MODIFICADA =====
function programarRecordatorio(email, nombreCompleto, horarioCita) {
  try {
    console.log('📅 Programando recordatorio para:', email, '-', nombreCompleto, '-', horarioCita);
    
    // Extraer la fecha de la cita del string del horario
    const fechaCita = extraerFechaDeCita(horarioCita);
    
    if (!fechaCita) {
      console.log('⚠️ No se pudo extraer fecha del horario:', horarioCita);
      return;
    }
    
    // Calcular fecha de envío (día anterior a la cita)
    const fechaEnvio = new Date(fechaCita);
    fechaEnvio.setDate(fechaEnvio.getDate() - 1);
    
    console.log('📅 Fecha de la cita:', fechaCita.toLocaleDateString());
    console.log('📅 Fecha envío recordatorio:', fechaEnvio.toLocaleDateString());
    
    // Verificar que no sea fecha pasada
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    fechaEnvio.setHours(0, 0, 0, 0);
    
    if (fechaEnvio < hoy) {
      console.log('⚠️ La fecha de envío ya pasó, no se programa recordatorio');
      return;
    }
    
    // Obtener o crear hoja de recordatorios
    const hojaRecordatorios = obtenerHojaRecordatoriosMejorada();
    
    // Preparar datos para insertar
    const ahora = new Date();
    const datosRecordatorio = [
      fechaEnvio,           // A: Fecha de envío
      email,                // B: Email familia
      nombreCompleto,       // C: Nombre y apellidos del alumno (MEJORADO)
      horarioCita,          // D: Horario completo de la cita
      ahora                 // E: Timestamp de cuándo se programó
    ];
    
    // Insertar en la primera fila vacía
    const ultimaFila = hojaRecordatorios.getLastRow();
    hojaRecordatorios.getRange(ultimaFila + 1, 1, 1, 5).setValues([datosRecordatorio]);
    
    console.log('✅ Recordatorio programado correctamente para:', nombreCompleto);
    
  } catch (error) {
    console.error('❌ Error programando recordatorio:', error);
  }
}

// ===== CORRECCIÓN 3: FUNCIÓN OBTENER HOJA MEJORADA =====
function obtenerHojaRecordatoriosMejorada() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName('Gogorarazpenak');
    
    if (!hoja) {
      console.log('🆕 Creando hoja Gogorarazpenak...');
      hoja = ss.insertSheet('Gogorarazpenak');
    }
    
    // SIEMPRE verificar y crear headers (por si no existen)
    const primeraFila = hoja.getRange(1, 1, 1, 5).getValues()[0];
    const tieneHeaders = primeraFila.some(celda => celda && celda.toString().trim() !== '');
    
    if (!tieneHeaders) {
      console.log('📋 Añadiendo headers a Gogorarazpenak...');
      
      // Crear headers corregidos
      const headers = [
        'Fecha Envío',        // A
        'Email',              // B
        'Alumno',             // C - Nombre del alumno
        'Horario Cita',       // D  
        'Programado el'       // E
      ];
      
      hoja.getRange(1, 1, 1, 5).setValues([headers]);
      
      // Formatear headers
      const headerRange = hoja.getRange(1, 1, 1, 5);
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      // Ajustar ancho de columnas
      hoja.setColumnWidth(1, 120); // Fecha envío
      hoja.setColumnWidth(2, 200); // Email  
      hoja.setColumnWidth(3, 150); // Alumno
      hoja.setColumnWidth(4, 300); // Horario cita
      hoja.setColumnWidth(5, 150); // Programado el
      
      console.log('✅ Headers añadidos a Gogorarazpenak');
    }
    
    return hoja;
    
  } catch (error) {
    console.error('❌ Error obteniendo hoja recordatorios:', error);
    throw error;
  }
}


// ===== FUNCIÓN MEJORADA: ENVIAR RECORDATORIOS CON SINCRONIZACIÓN AUTOMÁTICA =====
function enviarRecordatoriosCitas() {
  try {
    console.log('🔔 === INICIANDO ENVÍO DE RECORDATORIOS CON SINCRONIZACIÓN ===');
    
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0); // Reset horas para comparar solo fechas
    
    console.log('📅 Buscando recordatorios para hoy:', hoy.toLocaleDateString());
    
    // Obtener hoja de recordatorios
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRecordatorios = ss.getSheetByName('Gogorarazpenak');
    
    if (!hojaRecordatorios) {
      console.log('⚠️ No existe hoja Gogorarazpenak, no hay recordatorios pendientes');
      return;
    }
    
    const ultimaFila = hojaRecordatorios.getLastRow();
    
    if (ultimaFila <= 1) {
      console.log('📭 No hay recordatorios pendientes');
      return;
    }
    
    // ===== NUEVA FUNCIONALIDAD: OBTENER RESPUESTAS ACTUALES DEL FORMULARIO =====
    console.log('🔄 Obteniendo respuestas actuales del formulario para sincronización...');
    const respuestasActuales = obtenerRespuestasFormularioActuales();
    console.log(`📊 Respuestas actuales en formulario: ${respuestasActuales.length}`);
    
    // Crear set de emails+horarios activos para comparación rápida
    const citasActivas = new Set();
    respuestasActuales.forEach(resp => {
      const clave = `${resp.email}|${resp.fechaHoraOriginal}`;
      citasActivas.add(clave);
    });
    console.log(`🔑 Claves de citas activas creadas: ${citasActivas.size}`);
    
    // Obtener todos los datos de recordatorios
    const datosCompletos = hojaRecordatorios.getRange(2, 1, ultimaFila - 1, 5).getValues();
    
    console.log(`📊 Revisando ${datosCompletos.length} recordatorios programados...`);
    
    // Obtener configuración para la firma
    const hojaConfig = ss.getSheetByName('Ordutegiak');
    const firma = hojaConfig ? hojaConfig.getRange("A6").getValue() || "" : "";
    
    let recordatoriosEnviados = 0;
    let recordatoriosHuerfanos = 0;
    const filasParaBorrar = []; // Guardar índices de filas a borrar (empezando desde el final)
    
    // Revisar cada recordatorio (desde el final hacia atrás para borrar correctamente)
    for (let i = datosCompletos.length - 1; i >= 0; i--) {
      const fila = datosCompletos[i];
      const [fechaEnvio, email, familia, horarioCita, programadoEl] = fila;
      
      // Normalizar fecha de envío para comparar
      const fechaEnvioNormalizada = new Date(fechaEnvio);
      fechaEnvioNormalizada.setHours(0, 0, 0, 0);
      
      console.log(`📋 Fila ${i + 2}: ${familia} - Envío: ${fechaEnvioNormalizada.toLocaleDateString()}`);
      
      // ===== VERIFICACIÓN DE SINCRONIZACIÓN =====
      const claveRecordatorio = `${email}|${horarioCita}`;
      const citaAunExiste = citasActivas.has(claveRecordatorio);
      
      console.log(`🔍 Verificando existencia de cita: ${claveRecordatorio}`);
      console.log(`✓ Cita aún existe en formulario: ${citaAunExiste}`);
      
      // Si la cita ya no existe en el formulario, marcar para borrar
      if (!citaAunExiste) {
        console.log(`🗑️ RECORDATORIO HUÉRFANO: La cita de ${familia} ya no existe en el formulario`);
        filasParaBorrar.push(i + 2);
        recordatoriosHuerfanos++;
        continue; // Saltar al siguiente recordatorio
      }
      
      // Si es para hoy Y la cita aún existe, enviar recordatorio
      if (fechaEnvioNormalizada.getTime() === hoy.getTime()) {
        try {
          console.log(`📤 Enviando recordatorio a ${familia} (${email}) - Cita confirmada en formulario`);
          
          // Enviar el recordatorio
          enviarRecordatorio(email, familia, horarioCita, firma);
          
          recordatoriosEnviados++;
          
          // Marcar fila para borrar (número real de fila en la hoja)
          filasParaBorrar.push(i + 2); // +2 porque: +1 por array base-0, +1 por header
          
          console.log(`✅ Recordatorio enviado a ${familia}`);
          
        } catch (errorEnvio) {
          console.error(`❌ Error enviando recordatorio a ${familia}:`, errorEnvio);
        }
      }
      // Si la fecha ya pasó, también borrar (limpieza)
      else if (fechaEnvioNormalizada < hoy) {
        console.log(`🗑️ Borrando recordatorio vencido: ${familia}`);
        filasParaBorrar.push(i + 2);
      }
    }
    
    // Borrar filas procesadas (desde el final para no alterar índices)
    filasParaBorrar.forEach(numeroFila => {
      console.log(`🗑️ Borrando fila ${numeroFila}`);
      hojaRecordatorios.deleteRow(numeroFila);
    });
    
    const mensaje = `🔔 Recordatorios procesados: ${recordatoriosEnviados} enviados, ${recordatoriosHuerfanos} huérfanos eliminados, ${filasParaBorrar.length} filas totales eliminadas`;
    console.log(mensaje);
    
    return {
      enviados: recordatoriosEnviados,
      huerfanos: recordatoriosHuerfanos,
      eliminados: filasParaBorrar.length,
      mensaje: mensaje
    };
    
  } catch (error) {
    console.error('❌ Error en enviarRecordatoriosCitas:', error);
    
    // Opcional: enviar email de error al administrador
    try {
      MailApp.sendEmail({
        to: 'zuzendaria@mendialdeaipi.net', // Cambiar por email real
        subject: '❌ Error en sistema de recordatorios',
        body: `Error ejecutando recordatorios automáticos:\n\n${error.message}\n\nStack:\n${error.stack}`
      });
    } catch (mailError) {
      console.error('No se pudo enviar email de error:', mailError);
    }
  }
}

// ===== FUNCIÓN AUXILIAR: OBTENER RESPUESTAS ACTUALES DEL FORMULARIO =====
function obtenerRespuestasFormularioActuales() {
  try {
    console.log('📋 Buscando hoja de respuestas del formulario...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojas = ss.getSheets();
    
    // Buscar hoja de respuestas automáticamente
    let hojaRespuestas = hojas.find(hoja => 
      hoja.getName().toLowerCase().includes('respuesta') ||
      hoja.getName().toLowerCase().includes('response') ||
      hoja.getName().toLowerCase().includes('form')
    );
    
    if (!hojaRespuestas) {
      // Si no encuentra por nombre, buscar por contenido típico de formulario
      hojaRespuestas = hojas.find(hoja => {
        try {
          const headers = hoja.getRange(1, 1, 1, Math.min(10, hoja.getLastColumn())).getValues()[0];
          return headers.some(header => 
            header && (
              header.toString().includes('Marca temporal') ||
              header.toString().includes('Timestamp') ||
              header.toString().includes('correo electrónico') ||
              header.toString().includes('Aukeratu egun')
            )
          );
        } catch (e) {
          return false;
        }
      });
    }
    
    if (!hojaRespuestas) {
      console.log('⚠️ No se encontró hoja de respuestas del formulario');
      return [];
    }
    
    console.log(`✅ Usando hoja de respuestas: ${hojaRespuestas.getName()}`);
    
    // Usar la función existente para obtener respuestas
    return obtenerRespuestasFormulario(hojaRespuestas);
    
  } catch (error) {
    console.error('❌ Error obteniendo respuestas actuales del formulario:', error);
    return [];
  }
}
// =====================================================
// SISTEMA DE SINCRONIZACIÓN CON GOOGLE CALENDAR
// Versión limpia y organizada
// =====================================================

// ===== CONFIGURACIÓN Y GESTIÓN DE TRIGGERS =====

/**
 * Configurar la sincronización automática diaria
 */
function configurarSincronizacion() {
  const ui = SpreadsheetApp.getUi();
  
  const triggerExistente = verificarTriggerExistente();
  
  const mensaje = triggerExistente ? 
    'Zer egin nahi duzu sinkronizazio automatikoarekin?\n\n' +
   '• BAI: Eguneroko sinkronizazioa mantendu/berrkonfiguratu\n' +
   '• EZ: Desaktibatu sinkronizazio automatikoa' :
   'Formularioaren eta zure egutegiaren arteko eguneroko sinkronizazio automatikoa aktibatu nahi duzu?\n\n' +
   '• BAI: Egunero sinkronizatuko da 8: 00etan AM\n' +
   '• EZ: Ez da sinkronizaziorik konfiguratuko';

  
  const respuesta = ui.alert('Configurar Sincronización', mensaje, ui.ButtonSet.YES_NO);
  
  if (respuesta == ui.Button.YES) {
    configurarTriggerDiario();
    ui.alert('✅ Sinkronizazio aktibatua', 
      'Sinkronizazioa egunero exekutatuko da 8: 00etan\n\n' +
      ' Funtzio hau berriro exekutatu dezakezu desaktibatzeko.', 
      ui.ButtonSet.OK);
  } else {
    eliminarTriggerDiario();
    ui.alert('❌ Sinkronizazio ezabatua', 
      'Egutegiko hitzorduak ez dira automatikoki sinkronizatuko.', 
      ui.ButtonSet.OK);
  }
}

/**
 * Crear trigger diario para sincronización automática
 */
function configurarTriggerDiario() {
  eliminarTriggerDiario(); // Eliminar existente primero
  
  ScriptApp.newTrigger('sincronizarCitasAutomatico')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  
  console.log('✅ Trigger diario configurado para las 8:00 AM');
}

/**
 * Eliminar triggers de sincronización existentes
 */
function eliminarTriggerDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sincronizarCitasAutomatico') {
      ScriptApp.deleteTrigger(trigger);
      console.log('🗑️ Trigger eliminado');
    }
  });
}

/**
 * Verificar si existe un trigger de sincronización
 */
function verificarTriggerExistente() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(trigger => 
    trigger.getHandlerFunction() === 'sincronizarCitasAutomatico'
  );
}

// ===== FUNCIONES PRINCIPALES DE SINCRONIZACIÓN =====

/**
 * Ejecutar sincronización manual (desde menú)
 */
function ejecutarSincronizacionManual() {
  try {
    console.log('🔄 Eskuzko sinkronizazioa abiarazten ');
    const resultado = sincronizarCitasAutomatico();
    
    const ui = SpreadsheetApp.getUi();
    if (resultado) {
      ui.alert('✅ Sinkronizazio ossatua', 
        `Prozasatuak: ${resultado.procesadas || 0} respuestas\n` +
        `Berriak: ${resultado.nuevas || 0} citas\n` +
        `Ezabatuak: ${resultado.eliminadas || 0} citas`, 
        ui.ButtonSet.OK);
    } else {
      ui.alert('ℹ️ Sinkronizazio osatua', 
        'Sinkronizazioa zuzen exekutatu da.', 
        ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('❌ Errorea eskuzko sinkronizazioan:', error);
    SpreadsheetApp.getUi().alert('❌ Error', 
      `Errorea sinkronizazioan: ${error.message}`, 
      ui.ButtonSet.OK);
  }
}

/**
 * Función principal de sincronización (automática y manual)
 */
function sincronizarCitasAutomatico() {
  try {
    console.log('🔄 === CALENDARREKIN SINKRONIZAZIOA HASTEN ===');
    
    // 1. Buscar hoja de respuestas del formulario
    const hojaRespuestas = encontrarHojaRespuestas();
    if (!hojaRespuestas) {
      throw new Error('No se encontró la hoja de respuestas del formulario');
    }
    console.log(`📋 Orria erabiltzen: ${hojaRespuestas.getName()}`);
    
    // 2. Obtener calendar por defecto
    const calendar = CalendarApp.getDefaultCalendar();
    console.log(`📅 Calendar: ${calendar.getName()}`);
    
    // 3. Obtener respuestas actuales del formulario
    const respuestasActuales = obtenerRespuestasFormulario(hojaRespuestas);
    console.log(`📊 Egungo erantzunak: ${respuestasActuales.length}`);
    
    if (respuestasActuales.length === 0) {
      console.log('ℹ️ Ez dago sinkronizatzeko erantzunik');
      return { procesadas: 0, nuevas: 0, eliminadas: 0 };
    }
    
    // 4. Obtener snapshot anterior para comparar cambios
    const snapshotAnterior = obtenerSnapshotAnterior();
    console.log(`📊 Aurreko erantzunak: ${snapshotAnterior.length}`);
    
    // 5. Comparar y detectar cambios
    const cambios = compararRespuestas(respuestasActuales, snapshotAnterior);
    console.log(`📈 Cambios: ${cambios.nuevas.length} nuevas, ${cambios.eliminadas.length} eliminadas`);
    
    // 6. Procesar citas nuevas
    if (cambios.nuevas.length > 0) {
      procesarNuevasCitas(cambios.nuevas, calendar);
    }
    
    // 7. Procesar citas eliminadas
    if (cambios.eliminadas.length > 0) {
      procesarCitasEliminadas(cambios.eliminadas, calendar);
    }
    
    // 8. Guardar nuevo snapshot
    guardarSnapshot(respuestasActuales);
    
    console.log('✅ Sincronización completada exitosamente');
    
    return {
      procesadas: respuestasActuales.length,
      nuevas: cambios.nuevas.length,
      eliminadas: cambios.eliminadas.length
    };
    
  } catch (error) {
    console.error('❌ Error en sincronización:', error);
    
    // Enviar email de error al administrador si está configurado
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const hojaConfig = ss.getSheetByName('Ordutegiak');
      const firma = hojaConfig ? hojaConfig.getRange("A6").getValue() : '';
      
      if (firma && firma.includes('@')) {
        MailApp.sendEmail({
          to: firma,
          subject: '❌ Error en sincronización de calendario',
          body: `Error en la sincronización automática del calendario:\n\n${error.message}\n\nStack:\n${error.stack}`
        });
      }
    } catch (mailError) {
      console.error('⚠️ No se pudo enviar email de error:', mailError);
    }
    
    throw error;
  }
}

// ===== FUNCIONES DE PROCESAMIENTO DE DATOS =====

/**
 * Encontrar automáticamente la hoja de respuestas del formulario
 */
function encontrarHojaRespuestas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ss.getSheets();
  
  // Buscar por nombre típico
  let hojaRespuestas = hojas.find(hoja => 
    hoja.getName().toLowerCase().includes('respuesta') ||
    hoja.getName().toLowerCase().includes('response') ||
    hoja.getName().toLowerCase().includes('form')
  );
  
  // Si no encuentra por nombre, buscar por contenido típico
  if (!hojaRespuestas) {
    hojaRespuestas = hojas.find(hoja => {
      try {
        const headers = hoja.getRange(1, 1, 1, Math.min(10, hoja.getLastColumn())).getValues()[0];
        return headers.some(header => 
          header && (
            header.toString().includes('Marca temporal') ||
            header.toString().includes('Timestamp') ||
            header.toString().includes('correo electrónico') ||
            header.toString().includes('Aukeratu egun')
          )
        );
      } catch (e) {
        return false;
      }
    });
  }
  
  return hojaRespuestas;
}

/**
 * Obtener y procesar respuestas del formulario
 */
function obtenerRespuestasFormulario(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log('ℹ️ No hay respuestas en el formulario');
      return [];
    }
    
    const headers = data[0];
    const respuestas = [];
    
    // Índices de columnas (ajustar según tu formulario)
    const indices = {
      timestamp: 0,
      email: 1,
      nombre: 2,
      horario: 3,
      tema: 4
    };
    
    // Procesar cada fila (saltando headers)
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      
      const textoHorario = fila[indices.horario] ? fila[indices.horario].toString() : '';
      const emailPersona = fila[indices.email] ? fila[indices.email].toString() : '';
      const nombrePersona = fila[indices.nombre] ? fila[indices.nombre].toString() : '';
      const temaConsulta = fila[indices.tema] ? fila[indices.tema].toString() : 'Tutoría';
      
      if (!textoHorario || !emailPersona) {
        continue; // Saltar filas incompletas
      }
      
      // Extraer fecha y hora del texto del horario
      const fechaCita = extraerFechaDeCita(textoHorario);
      const horaCita = extraerHoraDeCita(textoHorario);
      
      if (!fechaCita || !horaCita) {
        console.warn(`⚠️ No se pudo extraer fecha/hora de: "${textoHorario}"`);
        continue;
      }
      
      const respuesta = {
        id: crearIdUnico(fila[indices.timestamp], emailPersona),
        timestamp: fila[indices.timestamp],
        nombre: nombrePersona || 'Sin nombre',
        email: emailPersona,
        fechaHoraOriginal: textoHorario,
        fecha: fechaCita,
        hora: horaCita,
        motivo: temaConsulta || 'Tutoría',
        filaOriginal: i + 1
      };
      
      respuestas.push(respuesta);
    }
    
    console.log(`📊 Respuestas procesadas: ${respuestas.length}`);
    return respuestas;
    
  } catch (error) {
    console.error('❌ Error procesando respuestas del formulario:', error);
    return [];
  }
}

/**
 * Procesar nuevas citas y crear eventos en Calendar
 */
function procesarNuevasCitas(nuevasCitas, calendar) {
  console.log(`🆕 Bilera berriak ${nuevasCitas.length} prozesatzen...`);
  
  nuevasCitas.forEach(cita => {
    try {
      const fechaHora = parsearFechaHora(cita.fecha, cita.hora);
      
      if (fechaHora.inicio && fechaHora.fin) {
        const titulo = `📅 Guraso bilera: ${cita.nombre}`;
        const descripcion = 
          `👨‍🎓 Ikaslea: ${cita.nombre}\n` +
          `📧 Emaila: ${cita.email}\n` +
          `⏰ Ordutegia: ${cita.fechaHoraOriginal}\n` +
          `💬 Zergatia: ${cita.motivo}\n` +
          `🆔 ID: ${cita.id}\n` +
          `📋 Galdetegitik automatikoki sortua`;
        
        const evento = calendar.createEvent(titulo, fechaHora.inicio, fechaHora.fin, {
          description: descripcion,
          guests: cita.email,
          sendInvites: false
        });
        
        console.log(`✅ Hitzordu sortua: ${cita.nombre} - ${fechaHora.inicio.toLocaleString()}`);
        
      } else {
        console.error(`❌ Error parseando fecha/hora para ${cita.nombre}`);
      }
      
    } catch (error) {
      console.error(`❌ Error creando cita para ${cita.nombre}:`, error);
    }
  });
}

/**
 * Procesar citas eliminadas y remover eventos de Calendar
 */
function procesarCitasEliminadas(citasEliminadas, calendar) {
  console.log(`🗑️ Procesando eliminación de ${citasEliminadas.length} citas...`);
  
  // Buscar eventos en un rango amplio
  const ahora = new Date();
  const pasado = new Date(ahora.getTime() - (60 * 24 * 60 * 60 * 1000)); // 60 días atrás
  const futuro = new Date(ahora.getTime() + (90 * 24 * 60 * 60 * 1000)); // 90 días adelante
  
  const eventos = calendar.getEvents(pasado, futuro);
  console.log(`📅 Revisando ${eventos.length} eventos en calendar`);
  
  citasEliminadas.forEach(cita => {
    try {
      let eventosEliminados = 0;
      
      eventos.forEach(evento => {
        try {
          const titulo = evento.getTitle() || '';
          const descripcion = evento.getDescription() || '';
          
          // Buscar por ID en descripción O por nombre+email
          const coincideID = descripcion.includes(cita.id);
          const coincideNombre = cita.nombre && titulo.includes(cita.nombre);
          const coincideEmail = cita.email && descripcion.includes(cita.email);
          
          if (coincideID || (coincideNombre && coincideEmail)) {
            evento.deleteEvent();
            eventosEliminados++;
            console.log(`🗑️ Evento eliminado: "${titulo}"`);
          }
          
        } catch (eventError) {
          console.error('⚠️ Error procesando evento individual:', eventError);
        }
      });
      
      if (eventosEliminados === 0) {
        console.log(`⚠️ No se encontraron eventos para eliminar: ${cita.nombre || 'Sin nombre'}`);
      }
      
    } catch (error) {
      console.error(`❌ Error eliminando eventos para ${cita.nombre || 'Sin nombre'}:`, error);
    }
  });
}

// ===== FUNCIONES AUXILIARES =====

/**
 * Crear ID único para cada respuesta
 */
function crearIdUnico(timestamp, email) {
  return Utilities.base64Encode(timestamp.toString() + '|' + email)
    .replace(/[^a-zA-Z0-9]/g, '');
}

/**
 * Parsear fecha y hora para crear eventos de Calendar
 */
function parsearFechaHora(fecha, hora) {
  try {
    if (!(fecha instanceof Date)) {
      return { inicio: null, fin: null };
    }
    
    if (isNaN(fecha.getTime())) {
      return { inicio: null, fin: null };
    }
    
    // Parsear hora
    if (typeof hora === 'string' && hora.includes(':')) {
      const partesHora = hora.split(':');
      const horas = parseInt(partesHora[0]);
      const minutos = parseInt(partesHora[1]);
      
      if (!isNaN(horas) && !isNaN(minutos)) {
        const inicio = new Date(fecha);
        inicio.setHours(horas, minutos, 0, 0);
        
        const fin = new Date(fecha);
        fin.setHours(horas, minutos + 30, 0, 0); // 30 minutos de duración
        
        return { inicio, fin };
      }
    }
    
    return { inicio: null, fin: null };
    
  } catch (error) {
    console.error('❌ Error parseando fecha/hora:', error);
    return { inicio: null, fin: null };
  }
}

// ===== GESTIÓN DE SNAPSHOTS =====

/**
 * Obtener snapshot anterior de respuestas
 */
function obtenerSnapshotAnterior() {
  try {
    const propiedades = PropertiesService.getScriptProperties();
    const snapshot = propiedades.getProperty('SNAPSHOT_RESPUESTAS');
    return snapshot ? JSON.parse(snapshot) : [];
  } catch (error) {
    console.warn('⚠️ Error leyendo snapshot anterior:', error);
    return [];
  }
}

/**
 * Guardar snapshot actual de respuestas
 */
function guardarSnapshot(respuestas) {
  try {
    const propiedades = PropertiesService.getScriptProperties();
    const snapshot = respuestas.map(r => ({
      id: r.id,
      timestamp: r.timestamp,
      email: r.email,
      nombre: r.nombre
    }));
    
    propiedades.setProperty('SNAPSHOT_RESPUESTAS', JSON.stringify(snapshot));
    console.log(`💾 Snapshot guardado: ${snapshot.length} respuestas`);
  } catch (error) {
    console.error('❌ Error guardando snapshot:', error);
  }
}

/**
 * Comparar respuestas actuales con anteriores
 */
function compararRespuestas(actuales, anteriores) {
  const idsActuales = actuales.map(r => r.id);
  const idsAnteriores = anteriores.map(r => r.id);
  
  const nuevas = actuales.filter(r => !idsAnteriores.includes(r.id));
  const eliminadas = anteriores.filter(r => !idsActuales.includes(r.id));
  
  return { nuevas, eliminadas };
}

// ===== FUNCIONES DE UTILIDAD Y DEBUG =====

/**
 * Limpiar snapshot guardado (para testing)
 */
function limpiarSnapshot() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('SNAPSHOT_RESPUESTAS');
    console.log('🧹 Snapshot limpiado');
    SpreadsheetApp.getUi().alert('✅ Snapshot limpiado', 
      'El snapshot se ha eliminado correctamente.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ Error limpiando snapshot:', error);
  }
}

/**
 * Mostrar estado actual de la sincronización
 */
function mostrarEstadoSincronizacion() {
  const triggerActivo = verificarTriggerExistente();
  const snapshot = obtenerSnapshotAnterior();
  
  console.log('=== ESTADO ACTUAL DE SINCRONIZACIÓN ===');
  console.log('Trigger activo:', triggerActivo);
  console.log('Respuestas en snapshot:', snapshot.length);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Estado de Sincronización', 
    `Sincronización automática: ${triggerActivo ? 'ACTIVADA' : 'DESACTIVADA'}\n` +
    `Respuestas registradas: ${snapshot.length}\n\n` +
    `${triggerActivo ? 'Se ejecuta diariamente a las 8:00 AM' : 'No hay sincronización automática'}`,
    ui.ButtonSet.OK);
}

/**
 * Limpiar todos los eventos de tutoría del calendar (uso con precaución)
 */
function limpiarEventosTutoriaCalendar() {
  const ui = SpreadsheetApp.getUi();
  const confirmacion = ui.alert('⚠️ Confirmación', 
    '¿Estás seguro de que quieres eliminar TODOS los eventos de tutoría del calendar?\n\n' +
    'Esta acción NO se puede deshacer.',
    ui.ButtonSet.YES_NO);
  
  if (confirmacion !== ui.Button.YES) {
    return;
  }
  
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const ahora = new Date();
    const pasado = new Date(ahora.getTime() - (60 * 24 * 60 * 60 * 1000));
    const futuro = new Date(ahora.getTime() + (90 * 24 * 60 * 60 * 1000));
    
    const eventos = calendar.getEvents(pasado, futuro);
    let eventosEliminados = 0;
     
    eventos.forEach(evento => {
      const titulo = evento.getTitle();
      if (titulo && titulo.includes('Tutoría')) {
        evento.deleteEvent();
        eventosEliminados++;
      }
    });
    
    ui.alert('✅ Limpieza completada', 
      `Se eliminaron ${eventosEliminados} eventos de tutoría del calendar`,
      ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ Error en limpieza:', error);
    ui.alert('❌ Error', 
      `Error durante la limpieza: ${error.message}`,
      ui.ButtonSet.OK);
  }
}
// ===== FUNCIÓN: MOSTRAR TUTORIAL DE AYUDA =====
function erakutsiLaguntza() {
  const html = HtmlService.createHtmlOutputFromFile('Laguntza')
    .setTitle('Laguntza - Tutorial')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Hitzordu-eskaera Tutoretzarako Laguntza');
}
