// --- CONFIGURACIÓN ---
const SPREADSHEET_ID = '1ePzu4sgTyap8QB0bDrkAWkFyxXmUaT21tffr96fp1hw'
const SHEET_NAME = 'Respuestas'
const UPLOAD_FOLDER_ID = '1Vz4pWhwWAVx04-TDaMonB_tYZVvPYmy5' //Para los archivos
const WEBHOOK_URL =
  'https://n8n.balticec.com/webhook-test/271def94-7f50-4d82-91db-b965e0271f27'

// ---------------------------------------------------
const PERMISSIONS_URL =
  'https://script.google.com/macros/s/AKfycbx3j5Q0h6v4IrMD4BLaLPyJZkDQLoyAnm1yHRU29-z-xJbc-JOd5_L-CV-5fIUJvm3D/exec?action=formulariosPermisosEFSRT'

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
  if (!isFormEnabled('Emprendimiento e Iniciativas de Negocio')) {
    return {
      success: false,
      message:
        'El formulario de Emprendimiento e Iniciativas de Negocio no está habilitado en este momento.',
    }
  }

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

  const correoIndex = headers.indexOf('Correo Estudiantil')
  const cicloIndex = headers.indexOf('Ciclo')
  const estadoIndex = headers.indexOf('Estado')

  if (correoIndex === -1) {
    throw new Error('No existe la columna "Correo Estudiantil"')
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
        'Ya existe un registro aprobado para este correo en este ciclo. No puedes registrar nuevamente.',
    }
  }

  // ---------- 2. ARCHIVO ----------

  let fileUrl = ''
  let fileId = ''

  if (data.adjunto && data.adjunto.data) {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID)

    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.adjunto.data),
      data.adjunto.mimeType,
      data.adjunto.name,
    )

    const ext = data.adjunto.name.split('.').pop()

    const tipoArchivo =
      data.tipo === 'Iniciativa de negocio' ? 'PlanNegocio' : 'Evidencias'

    const fileName = `${data.dni || 'SINDNI'}_${data.apellidos || 'Estudiante'}_${tipoArchivo}.${ext}`

    const file = folder.createFile(blob).setName(fileName)

    fileUrl = file.getUrl()
    fileId = file.getId()
  }

  // ---------- 3. REGISTRO ----------

  const idRegistro = Utilities.getUuid()

  sheet.appendRow([
    new Date(),
    idRegistro,
    data.nombres || '',
    data.apellidos || '',
    "'" + (data.dni || ''),
    correo,
    data.programa || '',
    data.ciclo || '',
    data.tipo || '',
    data.nombre_empresa || '-',
    data.marca || '-',
    data.mision || '-',
    data.vision || '-',
    data.estructura_legal || '-',
    fileUrl,
    '', // Nota
    '', // Estado
    '', // Justificación
    '', // Documento Informe
  ])

  // ---------- 4. EMAIL ----------

  sendEmail(data, idRegistro)

  // ---------- 5. WEBHOOK ----------

  if (WEBHOOK_URL) {
    const payload = {
      ...data,
      id: idRegistro,
      fileId: fileId,
      fileUrl: fileUrl,
    }

    delete payload.adjunto

    sendWebhook(payload)
  }

  return {
    success: true,
    id: idRegistro,
    fileId: fileId,
    fileUrl: fileUrl,
  }
}
// ---------------------------------------------------

function sendEmail(data, id) {
  const subject = `✅ Confirmación EFSRT - Emprendimiento (${id})`

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <div style="text-align: center; margin-bottom: 20px;">
           <h2 style="color: #7c3aed; margin: 0;">Registro Recibido</h2>
           <p style="color: #6b7280; font-size: 14px;">EFSRT - Emprendimiento e Iniciativas de Negocio</p>
        </div>
        
        <p style="color: #374151; font-size: 16px;">Hola <strong>${data.nombres}</strong>,</p>
        
        <p style="color: #374151; line-height: 1.6;">
          Hemos recibido tu solicitud de registro para EFSRT en la modalidad de Emprendimiento. A continuación el detalle:
        </p>
        
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <ul style="list-style: none; padding: 0; margin: 0; color: #4b5563; font-size: 14px;">
            <li style="margin-bottom: 8px;"><strong>ID Registro:</strong> ${id}</li>
            <li style="margin-bottom: 8px;"><strong>Tipo:</strong> ${data.tipo}</li>
            <li style="margin-bottom: 8px;"><strong>Programa:</strong> ${data.programa}</li>
            ${data.nombre_empresa ? `<li style="margin-bottom: 8px;"><strong>Empresa:</strong> ${data.nombre_empresa}</li>` : ''}
          </ul>
        </div>

        <p style="color: #374151; font-size: 14px;">
          Tu archivo adjunto (Plan de Negocio o Evidencias) ha sido almacenado correctamente.
        </p>
        
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        
        <div style="text-align: center; color: #9ca3af; font-size: 12px;">
          <p>© ${new Date().getFullYear()} Instituto Neumann</p>
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
