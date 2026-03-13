// --- CONFIGURACIÓN ---
const SPREADSHEET_ID = '1D_thD9OcXtVCS2eUSuiPQ8jKm4Apu4kdRY_upFOOrKk'
const SHEET_NAME = 'Respuestas'
const UPLOAD_FOLDER_ID = '1r8SO-ejRKldlZSlOrs95H3-CWdpBiRS0' // carpeta de Drive de la constancia
const WEBHOOK_URL =
  'https://n8n.balticec.com/webhook-test/40ad027a-2b9d-406c-9dcb-33bf6db7407c'
// ---------------------------------------------------

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents)
    const resultado = processSocialForm(params)

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
function processSocialForm(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  let sheet = ss.getSheetByName(SHEET_NAME)
  if (!sheet) sheet = ss.getSheets()[0]

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

  const correoIndex = headers.indexOf('Correo estudiantil')
  const cicloIndex = headers.indexOf('Ciclo')
  const estadoIndex = headers.indexOf('Estado')

  if (correoIndex === -1) {
    throw new Error('No existe la columna "Correo estudiantil"')
  }

  if (cicloIndex === -1) {
    throw new Error('No existe la columna "Ciclo"')
  }

  if (estadoIndex === -1) {
    throw new Error('No existe la columna "Estado"')
  }

  // 🔎 Validación: correo + ciclo + estado aprobado
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

  // ---------- 2. ARCHIVO ----------

  let fileUrl = ''
  let fileId = ''

  const fileField = data.adjunto ? 'adjunto' : 'constancia'

  if (data[fileField] && data[fileField].data) {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID)

    const blob = Utilities.newBlob(
      Utilities.base64Decode(data[fileField].data),
      data[fileField].mimeType,
      data[fileField].name,
    )

    const ext = data[fileField].name.split('.').pop()

    const newName = `${data.dni || 'SINDNI'}_${data.apellidos || 'Estudiante'}_ProyectoSocial.${ext}`

    const file = folder.createFile(blob).setName(newName)

    fileUrl = file.getUrl()
    fileId = file.getId()

    delete data[fileField]
  }

  // ---------- 3. REGISTRO ----------

  const idRegistro = Utilities.getUuid()
  data.id = idRegistro

  sheet.appendRow([
    new Date(),
    idRegistro,
    data.nombres || '',
    data.apellidos || '',
    "'" + (data.dni || ''),
    correo,
    data.programa || '',
    data.ciclo || '',

    data.razon_social || '',
    "'" + (data.ruc || ''),
    data.objetivo_iniciativa || '',
    data.jefe_inmediato || '',

    data.cargo_estudiante || '',
    data.horas_realizadas || '',
    data.actividad_social || '',
    data.actividades_estudiante || '',
    data.convenio || '',
    data.representante || '',
    data.direccion || '',

    fileUrl,
  ])

  // ---------- 4. EMAIL ----------

  sendConfirmationEmail(data, idRegistro)

  // ---------- 5. WEBHOOK ----------

  if (WEBHOOK_URL) {
    const payload = {
      ...data,
      id: idRegistro,
      fileUrl,
      fileId,
    }

    sendWebhook(payload)
  }

  return {
    success: true,
    id: idRegistro,
    fileUrl,
    fileId,
  }
}
// ---------------------------------------------------

function sendConfirmationEmail(data, id) {
  const subject = `✅ Confirmación EFSRT - Proyecto Social (${id})`
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <div style="text-align: center; margin-bottom: 20px;">
           <h2 style="color: #7c3aed; margin: 0;">Registro Recibido</h2>
           <p style="color: #6b7280; font-size: 14px;">EFSRT - Proyecto Social</p>
        </div>
        <p style="color: #374151; font-size: 16px;">Hola <strong>${data.nombres}</strong>,</p>
        <p style="color: #374151; line-height: 1.6;">
          Hemos recibido tu registro para la modalidad Proyecto Social. A continuación un resumen de los datos:
        </p>
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <ul style="list-style: none; padding: 0; margin: 0; color: #4b5563; font-size: 14px;">
            <li style="margin-bottom: 8px;"><strong>ID:</strong> ${id}</li>
            <li style="margin-bottom: 8px;"><strong>Programa:</strong> ${data.programa}</li>
            <li style="margin-bottom: 8px;"><strong>Ciclo:</strong> ${data.ciclo}</li>
            <li style="margin-bottom: 8px;"><strong>Organización:</strong> ${data.razon_social || '-'}</li>
            <li style="margin-bottom: 8px;"><strong>Cargo:</strong> ${data.cargo_estudiante || '-'}</li>
            <li style="margin-bottom: 8px;"><strong>Horas:</strong> ${data.horas_realizadas || '-'}</li>
          </ul>
        </div>
        <p style="color: #374151; font-size: 14px;">
          Tu constancia ha sido adjuntada correctamente.
        </p>
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        <div style="text-align: center; color: #9ca3af; font-size: 12px;">
          <p>© ${new Date().getFullYear()} Instituto Neumann. Todos los derechos reservados.</p>
        </div>
      </div>
    </div>
  `

  MailApp.sendEmail({
    to: data.correo_estudiantil,
    subject: subject,
    htmlBody: htmlBody,
  })
}

function sendWebhook(payload) {
  try {
    UrlFetchApp.fetch(WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    })
  } catch (e) {
    Logger.log('Error webhook: ' + e)
  }
}
