/*************************************************
 * CONFIGURACIÓN - IDS DE SHEETS (LISTADOS)
 *************************************************/
const SHEET_CENTROS_LABORALES_ID =
  '1r_RNnr_HPUytwDtSO71ldIVFaI6gVrFQ5RuDGgGZVvg'

const SHEET_EMPRENDIMIENTO_ID = '1FmhxiH15gLTdtuGtpdNSoVSYSHk56rv9sFGz8GITJAc'

const SHEET_PROCESOS_INSTITUCIONALES_ID =
  '15mvWzmVv9EiXgsixhfzVabErYJuNluELbMGSVhB24Ik'

const SHEET_PROYECTO_SOCIAL_ID = '1n-W1A-XDm6hwvPoKZhSTQ-au6cL_HmMd7u9uPjq416U'

//  ESTE SPREADSHEET CONTIENE:
// Sheet: Usuarios
// Sheet: FormulariosEFSRT
// Sheet: Logger
const SHEET_FORMULARIOS_PERMISOS_EFSRT =
  '10lFsej1Piz-sHlG0CUPdBtRzCWV09oXbDlqyfewjwao'

/*************************************************
 * LOGGER PERSONALIZADO (A SHEET)
 *************************************************/
function LoggerSheet(message, details = '') {
  try {
    const ss = SpreadsheetApp.openById(SHEET_FORMULARIOS_PERMISOS_EFSRT)
    let sheet = ss.getSheetByName('Logger')

    // Si la hoja no existe, la creamos con cabeceras iniciales
    if (!sheet) {
      sheet = ss.insertSheet('Logger')
      sheet.appendRow(['Timestamp', 'Mensaje', 'Detalles'])
      sheet.getRange('A1:C1').setFontWeight('bold')
    }

    const timestamp = new Date()

    // Convertir detalles a string si es un objeto
    const detString =
      typeof details === 'object'
        ? JSON.stringify(details, null, 2)
        : String(details)

    sheet.appendRow([timestamp, String(message), detString])
  } catch (e) {
    // Fallback normal si el Spreadsheet falla
    console.error('Error escribiendo en Logger sheet:', e)
  }
}

/*************************************************
 * HELPER  SHEET A JSON (POR ID)
 *************************************************/
function sheetToJson_(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId)
  const sheet = ss.getSheets()[0]
  const data = sheet.getDataRange().getValues()

  if (data.length < 2) return []

  const headers = data[0]

  return data
    .slice(1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) => {
      const obj = {}
      headers.forEach((h, i) => {
        obj[h] = row[i] ?? ''
      })
      return obj
    })
}

/*************************************************
 * HELPER  SHEET A JSON (POR NOMBRE)
 *************************************************/
function sheetToJsonByName_(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId)
  const sheet = ss.getSheetByName(sheetName)
  if (!sheet) return []

  const data = sheet.getDataRange().getValues()
  if (data.length < 2) return []

  const headers = data[0]

  return data
    .slice(1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) => {
      const obj = {}
      headers.forEach((h, i) => {
        obj[h] = row[i] ?? ''
      })
      return obj
    })
}

/*************************************************
 * RESPONSE JSON
 *************************************************/
function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  )
}

/*************************************************
 * AUTH / SESIÓN
 *************************************************/
function generateSessionToken(username) {
  const token = Utilities.getUuid()
  try {
    CacheService.getScriptCache().put(token, username, 21600)
  } catch (e) {}
  try {
    CacheService.getUserCache().put(token, username, 21600)
  } catch (e) {}
  return token
}

function validateSessionToken(token) {
  if (!token) return null

  let user = null
  try {
    user = CacheService.getScriptCache().get(token)
  } catch (e) {}
  if (user) return user

  try {
    user = CacheService.getUserCache().get(token)
  } catch (e) {}
  if (user) return user

  return null
}

/*************************************************
 * LOGIN API  DESDE SHEET "Usuarios"
 *************************************************/
function loginAPI(params) {
  try {
    const username = params.username || params.user
    const password = params.password || params.pass

    if (!username || !password) {
      return jsonResponse_({
        success: false,
        message: 'Usuario y contraseña requeridos',
      })
    }

    const users = sheetToJsonByName_(
      SHEET_FORMULARIOS_PERMISOS_EFSRT,
      'Usuarios',
    )

    const match = users.find(
      (u) =>
        String(u['Usuario']).toLowerCase() === String(username).toLowerCase() &&
        String(u['Contraseña']) === String(password),
    )

    if (!match) {
      return jsonResponse_({
        success: false,
        message: 'Credenciales inválidas',
      })
    }

    const token = generateSessionToken(match['Usuario'])

    return jsonResponse_({
      success: true,
      user: match['Usuario'],
      token,
    })
  } catch (err) {
    return jsonResponse_({
      success: false,
      message: err.message,
    })
  }
}

/*************************************************
 * ENDPOINT POST  FORMULARIOS EFSRT (PÚBLICO)
 *************************************************/
function formulariosPermisosEFSRT_API_() {
  const ss = SpreadsheetApp.openById(SHEET_FORMULARIOS_PERMISOS_EFSRT)
  const sheet = ss.getSheetByName('FormulariosEFSRT')

  if (!sheet) {
    return jsonResponse_({
      success: false,
      message: 'Sheet FormularioEFSRT no encontrado',
    })
  }

  const data = sheet.getDataRange().getValues()
  if (data.length < 2) return jsonResponse_([])

  const headers = data[0]

  const result = data
    .slice(1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null))
    .map((row) => {
      const obj = {}
      headers.forEach((h, i) => {
        obj[h] = row[i] ?? ''
      })
      return obj
    })

  return jsonResponse_(result)
}

/*************************************************
 * ENDPOINTS GET  LISTADOS (PROTEGIDOS)
 *************************************************/
function getCentrosLaborales_() {
  return sheetToJson_(SHEET_CENTROS_LABORALES_ID)
}

function getEmprendimientos_() {
  return sheetToJson_(SHEET_EMPRENDIMIENTO_ID)
}

function getProcesosInstitucionales_() {
  return sheetToJson_(SHEET_PROCESOS_INSTITUCIONALES_ID)
}

function getProyectoSocial_() {
  return sheetToJson_(SHEET_PROYECTO_SOCIAL_ID)
}

/*************************************************
 * ROUTER POST
 *************************************************/
function doPost(e) {
  const params = e.parameter || {}

  if (e.postData && e.postData.contents) {
    try {
      const data = JSON.parse(e.postData.contents)
      for (const key in data) params[key] = data[key]
    } catch (e) {}
  }

  switch (params.action) {
    case 'login':
      return loginAPI(params)

    case 'sendEmails':
      return sendEmails_(params)

    default:
      return jsonResponse_({
        success: false,
        message: 'Acción POST no válida',
      })
  }
}

/*************************************************
 * HELPER  SEND EMAILS (CENTROS LABORALES Y EMPRENDIMIENTO)
 *************************************************/
function sendEmails_(params) {
  const user = validateSessionToken(params.token)
  if (!user) {
    return jsonResponse_({
      success: false,
      message: 'Token inválido o expirado',
    })
  }

  const type = params.type || 'centros'

  let sheetId, idCol
  if (type === 'emprendimiento') {
    sheetId = SHEET_EMPRENDIMIENTO_ID
    idCol = 'ID Registro'
  } else if (type === 'procesos') {
    sheetId = SHEET_PROCESOS_INSTITUCIONALES_ID
    idCol = 'ID Registro'
  } else if (type === 'social') {
    sheetId = SHEET_PROYECTO_SOCIAL_ID
    idCol = 'ID'
  } else {
    sheetId = SHEET_CENTROS_LABORALES_ID
    idCol = 'ID'
  }

  let students = []
  try {
    students = JSON.parse(params.students || '[]')
  } catch (e) {
    return jsonResponse_({
      success: false,
      message: 'Formato de estudiantes inválido',
    })
  }

  if (!students.length) {
    return jsonResponse_({
      success: false,
      message: 'No se enviaron registros para procesar',
    })
  }

  // 1. Abrir Sheet para actualizar estados
  //    Asumimos que la hoja de datos es la primera (index 0)
  const ss = SpreadsheetApp.openById(sheetId)
  const sheet = ss.getSheets()[0]
  const data = sheet.getDataRange().getValues()

  if (data.length < 2) {
    return jsonResponse_({
      success: false,
      message: 'La hoja de cálculo está vacía o sin cabeceras',
    })
  }

  // Mapear cabeceras a índices
  const headers = data[0]
  const colIdIndex = headers.indexOf(idCol)
  const colSentIndex = headers.indexOf('Correo Enviado')
  const colDateIndex = headers.indexOf('Fecha Envío Correo')

  if (colIdIndex === -1 || colSentIndex === -1 || colDateIndex === -1) {
    return jsonResponse_({
      success: false,
      message: `No se encontraron las columnas requeridas (${idCol}, Correo Enviado, Fecha Envío Correo)`,
    })
  }

  // Mapa rápido de ID -> Fila (1-based para getRange, pero data es 0-based array)
  // Data index 1 es Row 2
  const idToRowMap = {}
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][colIdIndex])
    if (rowId) {
      idToRowMap[rowId] = i + 1 // Row number in Sheet
    }
  }

  const results = []
  let successCount = 0
  const now = new Date() // Fecha actual para todos

  students.forEach((st) => {
    LoggerSheet(JSON.stringify(st), 'debe traer')

    const id = String(st[idCol])
    const email = st['Correo estudiantil'] || st['Correo Estudiantil']
    const nombre = `${st['Nombres']} ${st['Apellidos']}`

    if (!email || !email.includes('@')) {
      results.push({
        id: id,
        email,
        status: 'error',
        error: 'Correo inválido',
      })
      return
    }

    const asunto = `Resultado Evaluación EFSRT - ${nombre}`

    let programaField = 'Programa'
    if (type === 'procesos') programaField = 'Programa de Estudios'

    let extraFields = ''
    if (type === 'emprendimiento') {
      const tipo = (st['Tipo EFSRT'] || '').toString().toLowerCase()
      extraFields = `
            <li style="margin-bottom: 8px;"><strong>Tipo EFSRT:</strong> ${st['Tipo EFSRT'] || '-'}</li>
      `
      // Sólo añadir nombre de empresa cuando el tipo sea "Negocio Propio"
      if (tipo === 'negocio propio') {
        extraFields += `
            <li style="margin-bottom: 8px;"><strong>Nombre Empresa:</strong> ${st['Nombre Empresa'] || '-'}</li>
        `
      }
    } else if (type === 'procesos') {
      extraFields = `
            <li style="margin-bottom: 8px;"><strong>Proceso Institucional:</strong> ${st['Proceso Institucional'] || '-'}</li>
            <li style="margin-bottom: 8px;"><strong>Jefe Inmediato:</strong> ${st['Jefe Inmediato'] || '-'}</li>
      `
    } else if (type === 'social') {
      extraFields = `
            <li style="margin-bottom: 8px;"><strong>Actividad Social:</strong> ${st['Actividad social'] || '-'}</li>
            <li style="margin-bottom: 8px;"><strong>Razón Social:</strong> ${st['Razón social / Nombre comercial'] || '-'}</li>
      `
    }

    const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; padding: 20px; border-radius: 10px;">
      <div style="background-color: #ffffff; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
        <div style="text-align: center; margin-bottom: 20px;">
           <h2 style="color: #7c3aed; margin: 0;">Resultado Evaluación</h2>
           <p style="color: #6b7280; font-size: 14px;">EFSRT - Universidad de Ciencias Empresariales</p>
        </div>
        
        <p style="color: #374151; font-size: 16px;">Hola <strong>${st['Nombres']}</strong>,</p>
        
        <p style="color: #374151; line-height: 1.6;">
          Adjuntamos el resultado de la revisión de tu expediente de EFSRT. A continuación el detalle:
        </p>
        
        <div style="background-color: #f3f4f6; padding: 15px; border-radius: 6px; margin: 20px 0;">
          <ul style="list-style: none; padding: 0; margin: 0; color: #4b5563; font-size: 14px;">
            <li style="margin-bottom: 8px;"><strong>Programa:</strong> ${st[programaField] || '-'}</li>
            ${extraFields}
            <li style="margin-bottom: 8px;"><strong>Estado:</strong> 
                <span style="color: ${(st['Estado'] || '').toString().toLowerCase().includes('aprobado') && !(st['Estado'] || '').toString().toLowerCase().includes('desaprobado') ? '#16a34a' : '#dc2626'}; font-weight: bold;">
                    ${st['Estado'] || 'Pendiente'}
                </span>
            </li>
            <li style="margin-bottom: 8px;"><strong>Nota:</strong> ${st['Nota'] || '-'}</li>
            <li style="margin-bottom: 8px;"><strong>Justificación:</strong> ${st['Justificación'] || 'Ninguna'}</li>
          </ul>
        </div>

        <div style="margin-top: 20px; text-align: center;">
            <p style="margin-bottom: 10px;"><strong>Documento Informe:</strong></p>
            ${
              st['Documento Informe']
                ? `<a href="${st['Documento Informe']}" style="background-color: #7c3aed; color: #ffffff; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-size: 14px;">Ver Documento</a>`
                : '<span style="color: #9ca3af;">No adjunto</span>'
            }
        </div>
        
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        
        <div style="text-align: center; color: #9ca3af; font-size: 12px;">
          <p>Si tienes consultas, por favor contacta a tu coordinador.</p>
          <p> ${new Date().getFullYear()} Universidad de Ciencias Empresariales</p>
        </div>
      </div>
    </div>
    `

    try {
      MailApp.sendEmail({
        to: email,
        subject: asunto,
        htmlBody: htmlBody,
      })

      // Intentar actualizar Sheet si existe el ID
      if (idToRowMap[id]) {
        const rowNum = idToRowMap[id]
        // data es 2D array, indices son colSentIndex y colDateIndex
        sheet.getRange(rowNum, colSentIndex + 1).setValue('SI')
        sheet.getRange(rowNum, colDateIndex + 1).setValue(now)
      }

      results.push({ id: id, email: email, status: 'success' })
      successCount++
    } catch (err) {
      results.push({
        id: id,
        email: email,
        status: 'error',
        error: err.toString(),
      })
    }
  })

  return jsonResponse_({
    success: true,
    processed: students.length,
    sent: successCount,
    details: results,
  })
}
/*************************************************
 * ROUTER GET (PROTEGIDO)
 *************************************************/
function doGet(e) {
  const action = e.parameter.action
  const token = e.parameter.token

  const PUBLIC_ACTIONS = ['formulariosPermisosEFSRT']

  if (!PUBLIC_ACTIONS.includes(action)) {
    const user = validateSessionToken(token)
    if (!user) {
      return jsonResponse_({
        success: false,
        message: 'Token inválido o expirado',
      })
    }
  }

  switch (action) {
    case 'formulariosPermisosEFSRT':
      return formulariosPermisosEFSRT_API_()

    case 'centros':
      return jsonResponse_(getCentrosLaborales_())

    case 'emprendimiento':
      return jsonResponse_(getEmprendimientos_())

    case 'procesos':
      return jsonResponse_(getProcesosInstitucionales_())

    case 'social':
      return jsonResponse_(getProyectoSocial_())

    default:
      return jsonResponse_({
        error: 'Ruta no valida',
      })
  }
}
