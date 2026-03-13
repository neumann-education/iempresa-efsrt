// --- CONFIGURACIÓN ---
const SPREADSHEET_ID = '1OVi4Hp2_5xv_sEH7QilaTa2-S0Ec5CPUXBoVH6JQkCs'
const SHEET_NAME = 'Respuestas'
const WEBHOOK_URL =
  'https://n8n.balticec.com/webhook-test/4f11bba3-dd22-4708-9312-7c3582553fb0'

const PERMISSIONS_URL =
  'https://script.google.com/macros/s/AKfycbw0XJcMKV6mXnnK827GoUTGNB5cYaKKG-rnbtT-3kRIfjArlVFfrmAnHUpE7wFPHZwTiQ/exec?action=formulariosPermisosEFSRT'

function isFormEnabled(formName) {
  try {
    const response = UrlFetchApp.fetch(PERMISSIONS_URL, {
      muteHttpExceptions: true,
    })
    const list = JSON.parse(response.getContentText())
    if (!Array.isArray(list)) return false
    const entry = list.find(
      (o) =>
        o['LISTADO EFSRT'] &&
        o['LISTADO EFSRT'].toString().trim().toLowerCase() ===
          formName.toString().trim().toLowerCase(),
    )
    return entry && Number(entry['Habilitado']) === 1
  } catch (e) {
    // Si falla la llamada, asumimos deshabilitado para ser conservadores
    Logger.log('Error comprobando permisos: ' + e)
    return false
  }
}
// ---------------------------------------------------

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents)
    const resultado = processForm(params)

    return ContentService.createTextOutput(
      JSON.stringify(resultado),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Error en el servidor: ' + error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}

// ---------------------------------------------------
function processForm(data) {
  if (!isFormEnabled('Procesos Institucionales')) {
    return {
      success: false,
      message:
        'El formulario de Procesos Institucionales no está habilitado en este momento.',
    }
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  let sheet = ss.getSheetByName(SHEET_NAME)

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME)

    const headers = [
      'Marca temporal',
      'ID Registro',
      'Nombres',
      'Apellidos',
      'DNI',
      'Correo Estudiantil',
      'Programa de Estudios',
      'Ciclo de Estudios',
      'Proceso Institucional',
      'Jefe Inmediato',
      'Fecha de Inicio',
      'Horas Realizadas',
      'Cargo Estudiante',
      'Actividades Estudiante',
      'Nota',
      'Estado',
      'Justificación',
      'Documento Informe',
    ]

    sheet.appendRow(headers)
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold')
    sheet.setFrozenRows(1)
  }

  // ---------- 1. VALIDAR CORREO Y CICLO ----------

  const correo = (data.correo_estudiantil || '').toString().trim().toLowerCase()

  const ciclo = (data.ciclo || '').toString().trim().toLowerCase()

  if (!correo) {
    return {
      success: false,
      message: 'El correo estudiantil es obligatorio.',
    }
  }

  if (!ciclo) {
    return {
      success: false,
      message: 'El ciclo es obligatorio.',
    }
  }

  const values = sheet.getDataRange().getValues()
  const headers = values[0]

  const correoIndex = headers.indexOf('Correo Estudiantil')
  const cicloIndex = headers.indexOf('Ciclo de Estudios')
  const estadoIndex = headers.indexOf('Estado')

  if (correoIndex === -1) {
    throw new Error('No existe la columna "Correo Estudiantil"')
  }

  if (cicloIndex === -1) {
    throw new Error('No existe la columna "Ciclo de Estudios"')
  }

  if (estadoIndex === -1) {
    throw new Error('No existe la columna "Estado"')
  }

  // 🔎 Validar: correo + ciclo + estado aprobado
  const existeAprobado = values.slice(1).some((row) => {
    const correoSheet = (row[correoIndex] || '').toString().trim().toLowerCase()
    const cicloSheet = (row[cicloIndex] || '').toString().trim().toLowerCase()
    const estadoSheet = (row[estadoIndex] || '').toString().trim().toLowerCase()

    return (
      correoSheet === correo &&
      cicloSheet === ciclo &&
      estadoSheet === 'aprobado'
    )
  })

  if (existeAprobado) {
    return {
      success: false,
      message:
        'Ya existe un registro para este correo en este ciclo. No puedes registrar nuevamente.',
    }
  }

  // ---------- 2. GENERAR ID ----------
  const idRegistro = Utilities.getUuid()

  // ---------- 3. GUARDAR EN SHEET ----------

  const row = [
    new Date(),
    idRegistro,
    data.nombres || '',
    data.apellidos || '',
    "'" + (data.dni || ''),
    correo,
    data.programa || '',
    data.ciclo || '',
    data.proceso_institucional || '',
    data.jefe_inmediato || '',
    data.fecha_inicio || '',
    data.horas_realizadas || '',
    data.cargo_estudiante || '',
    data.actividades_estudiante || '',
    '', // Nota
    '', // Estado
    '', // Justificación
    '', // Documento Informe
  ]

  sheet.appendRow(row)

  // ---------- 4. ENVIAR CORREO ----------
  sendConfirmationEmail(data, idRegistro)

  // ---------- 5. WEBHOOK ----------
  if (WEBHOOK_URL) {
    const payload = {
      ...data,
      id: idRegistro,
    }

    sendWebhook(payload)
  }

  return {
    success: true,
    message: 'Registro exitoso',
    id: idRegistro,
  }
}

// ---------------------------------------------------

function sendConfirmationEmail(data, id) {
  const subject = `✅ Confirmación EFSRT - Procesos Institucionales (${id})`

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <div style="text-align: center; margin-bottom: 20px;">
           <h2 style="color: #7c3aed; margin: 0;">Registro Recibido</h2>
           <p style="color: #6b7280; font-size: 14px;">EFSRT - Procesos Institucionales Neumann</p>
        </div>
        
        <p style="color: #374151; font-size: 16px;">Hola <strong>${data.nombres}</strong>,</p>
        
        <p style="color: #374151; line-height: 1.6;">
          Hemos recibido tu registro correctamente. Tu información será validada por tu jefe inmediato (${data.jefe_inmediato}).
        </p>
        
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <ul style="list-style: none; padding: 0; margin: 0; color: #4b5563; font-size: 14px;">
            <li style="margin-bottom: 8px;"><strong>ID Registro:</strong> ${id}</li>
            <li style="margin-bottom: 8px;"><strong>Oficina:</strong> ${data.proceso_institucional}</li>
            <li style="margin-bottom: 8px;"><strong>Cargo:</strong> ${data.cargo_estudiante}</li>
            <li style="margin-bottom: 8px;"><strong>Horas:</strong> ${data.horas_realizadas}</li>
          </ul>
        </div>
        
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        
        <div style="text-align: center; color: #9ca3af; font-size: 12px;">
          <p>© ${new Date().getFullYear()} Instituto Neumann</p>
        </div>
      </div>
    </div>
  `

  try {
    MailApp.sendEmail({
      to: data.correo_estudiantil,
      subject: subject,
      htmlBody: htmlBody,
    })
  } catch (e) {
    console.error('Error enviando correo: ' + e.toString())
  }
}

// ---------------------------------------------------

function sendWebhook(payload) {
  try {
    UrlFetchApp.fetch(WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    })
  } catch (e) {
    console.log('Error webhook:', e)
  }
}
