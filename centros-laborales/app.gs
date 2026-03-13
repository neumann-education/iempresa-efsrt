// --- VARIABLES DE ENTORNO ---
const SPREADSHEET_ID = '1zvsYztL59i33RuliHFxo2CnUKAFzpUhyuhN3OkdYSYI'
const SHEET_NAME = 'Respuestas'
const UPLOAD_FOLDER_ID = '1SSqVsffohVTDRVgWic1ibEhT8MkL3mdX'
// Webhook URL (Pega tu URL aquí)
const WEBHOOK_URL =
  'https://n8n.balticec.com/webhook-test/3073d6f1-51ce-4282-9319-262a1c6bd200'
// ----------------------------

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

function doPost(e) {
  try {
    // Si los datos vienen como JSON string en el postData
    var params = JSON.parse(e.postData.contents)
    var resultado = processExternalForm(params)

    return ContentService.createTextOutput(
      JSON.stringify(resultado),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Error en doPost: ' + error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0]
    const values = sheet.getDataRange().getValues()
    if (values.length === 0) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, data: [] }),
      ).setMimeType(ContentService.MimeType.JSON)
    }

    const headers = values[0].map((h) => h.toString())
    const rows = values.slice(1)

    const data = rows.map(function (row) {
      const obj = {}
      headers.forEach(function (h, i) {
        obj[h] = row[i]
      })
      return obj
    })

    if (e.parameter && e.parameter.id) {
      const filtered = data.filter(function (item) {
        return (
          item['registroId'] === e.parameter.id ||
          item['ID'] === e.parameter.id ||
          item['id'] === e.parameter.id
        )
      })
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, data: filtered }),
      ).setMimeType(ContentService.MimeType.JSON)
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, data: data }),
    ).setMimeType(ContentService.MimeType.JSON)
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: 'Error en doGet: ' + error.toString(),
      }),
    ).setMimeType(ContentService.MimeType.JSON)
  }
}

function processExternalForm(data) {
  if (!isFormEnabled('Centros Laborales')) {
    return {
      success: false,
      message:
        'El formulario de Centros Laborales no está habilitado en este momento.',
    }
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  let sheet = ss.getSheetByName(SHEET_NAME)
  if (!sheet) {
    sheet = ss.getSheets()[0]
  }

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
  const cicloIndex = headers.indexOf('Ciclo')
  const estadoIndex = headers.indexOf('Estado')

  if (correoIndex === -1) {
    throw new Error('No se encontró la columna "Correo Estudiantil"')
  }

  if (cicloIndex === -1) {
    throw new Error('No se encontró la columna "Ciclo"')
  }

  if (estadoIndex === -1) {
    throw new Error('No se encontró la columna "Estado"')
  }

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

  // 3. Generar ID único
  const registroId = Utilities.getUuid()
  data.id = registroId

  // 4. Manejar archivo
  let fileUrl = ''

  if (data.constancia && data.constancia.data) {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID)

    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.constancia.data),
      data.constancia.mimeType,
      data.constancia.name,
    )

    const newName = `${data.dni || 'SIN_DNI'}_${data.apellidos || ''}_Constancia.pdf`

    const file = folder.createFile(blob).setName(newName)

    fileUrl = file.getUrl()

    data.fileUrl = fileUrl
    data.fileId = file.getId()

    delete data.constancia
  }

  // 5. Preparar fila
  const rowData = [
    new Date(), // Marca temporal
    registroId, // ID
    data.nombres || '',
    data.apellidos || '',
    "'" + (data.dni || ''),
    correo,
    data.programa || '',
    data.ciclo || '',
    data.razon_social || '',
    "'" + (data.ruc || ''),
    data.jefe || '',
    data.fecha_inicio || '',
    data.horas || '',
    data.desc || '',
    data.cargo || '',
    data.actividades || '',
    fileUrl,
    data.evaluacion_constancia || '',
    '', // Nota
    '', // Estado
    '', // Justificación
  ]

  sheet.appendRow(rowData)

  // 6. Enviar correo
  sendConfirmationEmail(data)

  // 7. Webhook
  if (WEBHOOK_URL) {
    sendToWebhook(data)
  }

  return {
    success: true,
    message: 'Registro guardado exitosamente.',
    id: registroId,
  }
}

function sendConfirmationEmail(data) {
  const subject =
    '✅ Confirmación de Registro EFSRT - IEmpresa (ID: ' + data.id + ')'

  // Plantilla HTML simple y limpia
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <div style="text-align: center; margin-bottom: 20px;">
           <h2 style="color: #7c3aed; margin: 0;">Registro Exitoso</h2>
           <p style="color: #6b7280; font-size: 14px;">Experiencia Formativa en Situaciones Reales de Trabajo</p>
        </div>
        
        <p style="color: #374151; font-size: 16px;">Hola <strong>${data.nombres}</strong>,</p>
        
        <p style="color: #374151; line-height: 1.6;">
          Tu formulario de registro EFSRT ha sido recibido correctamente. A continuación, un resumen de la información registrada:
        </p>
        
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <ul style="list-style: none; padding: 0; margin: 0; color: #4b5563; font-size: 14px;">
            <li style="margin-bottom: 8px;"><strong>ID registro:</strong> ${data.id}</li>
            <li style="margin-bottom: 8px;"><strong>Programa:</strong> ${data.programa}</li>
            <li style="margin-bottom: 8px;"><strong>Ciclo:</strong> ${data.ciclo}</li>
            <li style="margin-bottom: 8px;"><strong>Empresa:</strong> ${data.razon_social}</li>
            <li style="margin-bottom: 8px;"><strong>Cargo:</strong> ${data.cargo}</li>
            <li><strong>Fecha Inicio:</strong> ${data.fecha_inicio}</li>
          </ul>
        </div>
        
        <p style="color: #374151; font-size: 14px;">
          Hemos adjuntado tu constancia en nuestros registros. Si tienes alguna duda, contacta con tu coordinador de carrera.
        </p>
        
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        
        <div style="text-align: center; color: #9ca3af; font-size: 12px;">
          <p>© ${new Date().getFullYear()} IEmpresa. Todos los derechos reservados.</p>
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

function sendToWebhook(data) {
  try {
    const payload = JSON.stringify(data)
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: payload,
      muteHttpExceptions: true, // Para que no falle el script si el webhook responde error
    }
    UrlFetchApp.fetch(WEBHOOK_URL, options)
  } catch (e) {
    Logger.log('Error enviando webhook: ' + e.toString())
  }
}
