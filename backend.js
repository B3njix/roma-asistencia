// Google Apps Script - Backend Optimizado para Control de Asistencia
const SPREADSHEET_ID = '1FdunyeX_56sSUblUYA5IgaDMiamWDZSwwPKicHsOKLk';
const SHEET_NAME = 'Asistencia';
const EMPLOYEES_SHEET_NAME = 'Empleados';
const USERS_SHEET_NAME = 'Usuarios';

function doPost(e) {
  try {
    let data = {};
    if (e.parameter) {
      data = Object.keys(e.parameter).reduce((acc, key) => {
        acc[key] = e.parameter[key];
        return acc;
      }, {});
    }

    console.log('Acción:', data.action, '| Datos:', JSON.stringify(data));
    
    const actions = {
      'getEmployeeInfo': getEmployeeInfo,
      'registerAttendance': registerAttendance,
      'login': login,
      'getEmployeeList': getEmployeeList,
      'addEmployee': addEmployee,
      'updateEmployee': updateEmployee,
      'deleteEmployee': deleteEmployee,
      'getAttendanceRecords': getAttendanceRecords,
      'testConnection': testConnection,
      'getLastAttendanceRecord': getLastAttendanceRecord,
      'generateTimeReport': generateTimeReport,
      'getProjectDistribution': getProjectDistribution
    };

    const actionHandler = actions[data.action] || (() => ({
      success: false,
      message: 'Acción no válida',
      availableActions: Object.keys(actions)
    }));

    const result = actionHandler(data);
    console.log('Resultado:', JSON.stringify(result));
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('Error crítico:', error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Error del servidor: ' + error.toString(),
      stack: error.stack
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Función de autenticación unificada
function login(data) {
  try {
    if (!data.username || !data.password) {
      return { success: false, message: 'Usuario y contraseña requeridos' };
    }

    const userSheet = getOrCreateUserSheet();
    const userData = userSheet.getDataRange().getValues();
    
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === data.username && userData[i][1] === data.password) {
        return {
          success: true,
          user: {
            username: userData[i][0],
            name: userData[i][2] || 'Sin nombre',
            role: userData[i][3] || 'usuario',
            status: userData[i][4] || 'Activo',
            employeeId: userData[i][5] || ''
          }
        };
      }
    }
    
    return { success: false, message: 'Credenciales inválidas' };
  } catch (error) {
    return { success: false, message: 'Error en login: ' + error.toString() };
  }
}

// Función de prueba de conexión
function testConnection() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetNames = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    return {
      success: true,
      message: 'Conexión exitosa con Google Sheets',
      spreadsheetId: SPREADSHEET_ID,
      sheets: sheetNames
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error de conexión: ' + error.toString()
    };
  }
}

// Obtener información del empleado por ID
function getEmployeeInfo(data) {
  try {
    if (!data.employeeId) return { success: false, message: 'ID de empleado requerido' };

    const employeeSheet = getOrCreateEmployeeSheet();
    const employeeData = employeeSheet.getDataRange().getValues();
    
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] && employeeData[i][0].toString().trim() === data.employeeId.toString().trim()) {
        return {
          success: true,
          employee: {
            id: employeeData[i][0],
            name: employeeData[i][1] || 'Sin nombre',
            department: employeeData[i][2] || 'Sin departamento',
            position: employeeData[i][3] || 'Sin puesto',
            status: employeeData[i][4] || 'Activo'
          }
        };
      }
    }
    
    return { success: false, message: 'Empleado no encontrado' };
  } catch (error) {
    return { success: false, message: 'Error al buscar empleado: ' + error.toString() };
  }
}

// Obtener lista de empleados con paginación
function getEmployeeList(data) {
  try {
    const employeeSheet = getOrCreateEmployeeSheet();
    const allData = employeeSheet.getDataRange().getValues();
    
    if (allData.length <= 1) return {
      success: true,
      employees: [],
      total: 0
    };

    const employees = allData.slice(1).map(row => ({
      id: row[0],
      name: row[1] || 'Sin nombre',
      department: row[2] || 'Sin departamento',
      position: row[3] || 'Sin puesto',
      status: row[4] || 'Activo'
    }));

    // Paginación
    const page = parseInt(data.page) || 1;
    const limit = parseInt(data.limit) || 10;
    const startIndex = (page - 1) * limit;
    const paginatedEmployees = employees.slice(startIndex, startIndex + limit);
    
    return {
      success: true,
      employees: paginatedEmployees,
      total: employees.length,
      page: page,
      totalPages: Math.ceil(employees.length / limit)
    };
  } catch (error) {
    return {
      success: false,
      message: 'Error al obtener lista de empleados: ' + error.toString()
    };
  }
}

// Agregar un nuevo empleado
function addEmployee(data) {
  try {
    const employeeSheet = getOrCreateEmployeeSheet();
    const id = data.id;
    const name = data.name;
    const department = data.department;
    const position = data.position;
    const status = data.status;

    // Verificar si el ID ya existe
    const existingEmployee = getEmployeeInfo({ employeeId: id });
    if (existingEmployee.success) {
      return { success: false, message: 'El ID del empleado ya existe' };
    }

    // Agregar el nuevo empleado
    employeeSheet.appendRow([id, name, department, position, status]);

    return { success: true, message: 'Empleado agregado correctamente' };
  } catch (error) {
    return { success: false, message: 'Error al agregar empleado: ' + error.toString() };
  }
}

// Actualizar un empleado existente
function updateEmployee(data) {
  try {
    const employeeSheet = getOrCreateEmployeeSheet();
    const id = data.id;
    const name = data.name;
    const department = data.department;
    const position = data.position;
    const status = data.status;

    const employeeData = employeeSheet.getDataRange().getValues();
    let updated = false;

    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] && employeeData[i][0].toString() === id.toString()) {
        employeeSheet.getRange(i+1, 2, 1, 4).setValues([[name, department, position, status]]);
        updated = true;
        break;
      }
    }

    if (!updated) {
      return { success: false, message: 'Empleado no encontrado' };
    }

    return { success: true, message: 'Empleado actualizado correctamente' };
  } catch (error) {
    return { success: false, message: 'Error al actualizar empleado: ' + error.toString() };
  }
}

// Eliminar un empleado
function deleteEmployee(data) {
  try {
    const employeeSheet = getOrCreateEmployeeSheet();
    const employeeId = data.employeeId;

    const employeeData = employeeSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] && employeeData[i][0].toString() === employeeId.toString()) {
        rowIndex = i+1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Empleado no encontrado' };
    }

    employeeSheet.deleteRow(rowIndex);
    return { success: true, message: 'Empleado eliminado correctamente' };
  } catch (error) {
    return { success: false, message: 'Error al eliminar empleado: ' + error.toString() };
  }
}

// Registrar asistencia - MODIFICADO para aceptar "almuerzo"
function registerAttendance(data) {
  try {
    // Validaciones
    if (!data.employeeId || !data.type) {
      return { success: false, message: 'ID de empleado y tipo son requeridos' };
    }
    
    if (!['entrada', 'salida'].includes(data.type.toLowerCase())) {
      return { success: false, message: 'Tipo inválido. Use "entrada" o "salida".' };
    }

    // Validar tipo de hora - nuevos tipos
    const validHourTypes = ['hora normal', 'hora extra', 'almuerzo'];
    const hourType = validHourTypes.includes(data.hourType) 
        ? data.hourType 
        : 'hora normal';

    const employeeInfo = getEmployeeInfo({ employeeId: data.employeeId });
    const employeeName = employeeInfo.success ? employeeInfo.employee.name : 'Sin nombre';
    
    const now = new Date();
    const rowData = [
      now, // Timestamp
      data.employeeId,
      employeeName,
      data.project || 'Sin proyecto',
      hourType, // Usamos la variable validada
      data.type,
      now.toLocaleDateString('es-ES'),
      now.toLocaleTimeString('es-ES'),
      getWeekNumber(now),
      now.getMonth() + 1,
      now.getFullYear(),
      getDayOfWeek(now),
      data.notes || ''
    ];

    const sheet = getOrCreateSheet();
    sheet.appendRow(rowData);

    return {
      success: true,
      message: `✅ ${data.type === 'entrada' ? 'Entrada' : 'Salida'} registrada correctamente`,
      data: {
        employeeId: data.employeeId,
        employeeName: employeeName,
        type: data.type,
        hourType: hourType,
        timestamp: now.toISOString()
      }
    };
  } catch (error) {
    return { success: false, message: 'Error al registrar asistencia: ' + error.toString() };
  }
}

// Obtener último registro de asistencia
function getLastAttendanceRecord(data) {
  try {
    if (!data.employeeId) return { success: false, message: "ID de empleado requerido" };
    
    const sheet = getOrCreateSheet();
    if (!sheet) return { success: false, message: "Hoja de asistencia no encontrada" };

    const dataRange = sheet.getDataRange().getValues();
    
    // Buscar desde el final (últimos registros primero)
    for (let i = dataRange.length - 1; i >= 1; i--) {
      if (dataRange[i][1] === data.employeeId) {
        return { 
          success: true, 
          record: {
            timestamp: dataRange[i][0],
            employeeId: dataRange[i][1],
            employeeName: dataRange[i][2],
            project: dataRange[i][3],
            hourType: dataRange[i][4],
            type: dataRange[i][5],
            date: dataRange[i][6],
            time: dataRange[i][7]
          }
        };
      }
    }
    
    return { success: false, message: "No se encontraron registros" };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

// Obtener registros de asistencia con paginación
function getAttendanceRecords(data) {
  try {
    const sheet = getOrCreateSheet();
    if (!sheet) return { success: true, records: [], total: 0 };
    
    const allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return { success: true, records: [], total: 0 };
    
    // Mapear registros con estructura corregida
    let records = allData.slice(1).map(row => ({
      timestamp: row[0],
      employeeId: row[1],
      employeeName: row[2],
      project: row[3],
      hourType: row[4],
      type: row[5],
      date: row[6],
      time: row[7],
      week: row[8],
      month: row[9],
      year: row[10],
      dayOfWeek: row[11],
      notes: row[12]
    }));
    
    // Aplicar filtros
    if (data.employeeId) records = records.filter(r => r.employeeId === data.employeeId);
    if (data.project) records = records.filter(r => r.project === data.project);
    if (data.type) records = records.filter(r => r.type === data.type);
    
    if (data.startDate && data.endDate) {
      const start = new Date(data.startDate);
      const end = new Date(data.endDate);
      end.setDate(end.getDate() + 1); // Incluir día completo
      
      records = records.filter(r => {
        const recordDate = new Date(r.timestamp);
        return recordDate >= start && recordDate <= end;
      });
    }
    
    // Ordenar por timestamp descendente
    records.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    // Paginación
    const page = parseInt(data.page) || 1;
    const limit = parseInt(data.limit) || 10;
    const startIndex = (page - 1) * limit;
    const paginatedRecords = records.slice(startIndex, startIndex + limit);
    
    return {
      success: true,
      records: paginatedRecords,
      total: records.length,
      page: page,
      totalPages: Math.ceil(records.length / limit)
    };
  } catch (error) {
    return { success: false, message: 'Error al obtener registros: ' + error.toString() };
  }
}

// Generar reporte de horas trabajadas - MODIFICADO para excluir almuerzo
function generateTimeReport(data) {
  try {
    if (!data.startDate || !data.endDate) {
      return { success: false, message: 'Fechas de inicio y fin son requeridas' };
    }
    
    const reportType = data.period || 'semanal';
    const attendanceData = getAttendanceRecords({
      startDate: data.startDate,
      endDate: data.endDate,
      limit: 0 // Sin límite
    });
    
    if (!attendanceData.success) return attendanceData;
    
    const { individualReport, globalTotal } = processAttendanceRecords(
      attendanceData.records, 
      reportType
    );
    
    return {
      success: true,
      individualReport: individualReport,
      globalTotal: globalTotal
    };
  } catch (error) {
    return { success: false, message: 'Error al generar reporte: ' + error.toString() };
  }
}

// Procesar registros para cálculo de horas - MODIFICADO para excluir almuerzo
function processAttendanceRecords(records, reportType) {
  // Ordenar registros cronológicamente
  records.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
  
  const individualReport = {};
  let globalTotal = 0;

  // Agrupar por empleado y proyecto
  const groupedRecords = {};
  records.forEach(record => {
    // Excluir los registros de tipo "almuerzo"
    if (record.hourType === 'almuerzo') {
      return;
    }
    const key = `${record.employeeId}-${record.project}`;
    if (!groupedRecords[key]) groupedRecords[key] = [];
    groupedRecords[key].push(record);
  });

  // Procesar cada grupo
  for (const key in groupedRecords) {
    const [employeeId, project] = key.split('-');
    const records = groupedRecords[key];
    
    // Agrupar por período
    const periodRecords = {};
    records.forEach(record => {
      const date = new Date(record.timestamp);
      const period = reportType === 'semanal' ? getWeekNumber(date) : date.getMonth() + 1;
      
      if (!periodRecords[period]) periodRecords[period] = [];
      periodRecords[period].push(record);
    });

    // Calcular horas por período
    for (const period in periodRecords) {
      const periodData = periodRecords[period];
      let totalHours = 0;
      let lastEntrada = null;

      for (const record of periodData) {
        if (record.type === 'entrada') {
          lastEntrada = new Date(record.timestamp);
        } else if (record.type === 'salida' && lastEntrada) {
          const salida = new Date(record.timestamp);
          const diffHours = (salida - lastEntrada) / (1000 * 60 * 60);
          totalHours += diffHours;
          lastEntrada = null;
        }
      }

      // Actualizar reportes
      if (!individualReport[employeeId]) individualReport[employeeId] = {};
      if (!individualReport[employeeId][period]) individualReport[employeeId][period] = {};
      
      individualReport[employeeId][period][project] = 
        (individualReport[employeeId][period][project] || 0) + totalHours;
      
      globalTotal += totalHours;
    }
  }

  return { individualReport, globalTotal };
}

// Distribución por proyectos
function getProjectDistribution() {
  try {
    const sheet = getOrCreateSheet();
    if (!sheet) return { success: true, distribution: [] };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, distribution: [] };
    
    const distribution = {};
    for (let i = 1; i < data.length; i++) {
      const project = data[i][3];
      if (project) distribution[project] = (distribution[project] || 0) + 1;
    }
    
    const result = Object.keys(distribution).map(project => ({
      project: project,
      count: distribution[project]
    }));
    
    return { success: true, distribution: result };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// ======== FUNCIONES AUXILIARES ========
function getOrCreateSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      setupSheetHeaders(sheet);
    } else if (!hasValidHeaders(sheet)) {
      setupSheetHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    throw new Error('Error al obtener/crear hoja: ' + error.message);
  }
}

function getOrCreateEmployeeSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET_NAME);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(EMPLOYEES_SHEET_NAME);
      setupEmployeeSheetHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    throw new Error('Error al obtener hoja de empleados: ' + error.message);
  }
}

function getOrCreateUserSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(USERS_SHEET_NAME);
      setupUserSheetHeaders(sheet);
    }
    
    return sheet;
  } catch (error) {
    throw new Error('Error al obtener hoja de usuarios: ' + error.message);
  }
}

function setupSheetHeaders(sheet) {
  const headers = [
    'Timestamp', 'ID Empleado', 'Nombre', 'Proyecto', 
    'Tipo de Hora', 'Tipo', 'Fecha', 'Hora',
    'Semana', 'Mes', 'Año', 'Día de la Semana', 'Notas'
  ];
  
  sheet.clearContents();
  sheet.appendRow(headers);
  
  // Formato
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#667eea').setFontColor('white').setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function setupEmployeeSheetHeaders(sheet) {
  const headers = ['ID', 'Nombre', 'Departamento', 'Puesto', 'Estado'];
  sheet.clearContents();
  sheet.appendRow(headers);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#28a745').setFontColor('white').setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function setupUserSheetHeaders(sheet) {
  const headers = ['Usuario', 'Contraseña', 'Nombre', 'Rol', 'Estado', 'EmpleadoId'];
  sheet.clearContents();
  sheet.appendRow(headers);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a5568').setFontColor('white').setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function hasValidHeaders(sheet) {
  try {
    const headers = sheet.getRange(1, 1, 1, 13).getValues()[0];
    const expected = [
      'Timestamp', 'ID Empleado', 'Nombre', 'Proyecto', 
      'Tipo de Hora', 'Tipo', 'Fecha', 'Hora',
      'Semana', 'Mes', 'Año', 'Día de la Semana', 'Notas'
    ];
    
    return expected.every((h, i) => headers[i] === h);
  } catch (error) {
    return false;
  }
}

function getWeekNumber(date) {
  const startDate = new Date(date.getFullYear(), 0, 1);
  const days = Math.floor((date - startDate) / (24 * 60 * 60 * 1000));
  return Math.ceil((days + startDate.getDay() + 1) / 7);
}

function getDayOfWeek(date) {
  const days = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  return days[date.getDay()];
}

// Función para crear hojas si no existen
function createSheetsIfNotExist() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Crear/verificar hoja de Asistencia
    let asistenciaSheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!asistenciaSheet) {
      asistenciaSheet = spreadsheet.insertSheet(SHEET_NAME);
      setupSheetHeaders(asistenciaSheet);
    }

    // Crear/verificar hoja de Empleados
    let empleadosSheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET_NAME);
    if (!empleadosSheet) {
      empleadosSheet = spreadsheet.insertSheet(EMPLOYEES_SHEET_NAME);
      setupEmployeeSheetHeaders(empleadosSheet);
    }

    // Crear/verificar hoja de Usuarios
    let usuariosSheet = spreadsheet.getSheetByName(USERS_SHEET_NAME);
    if (!usuariosSheet) {
      usuariosSheet = spreadsheet.insertSheet(USERS_SHEET_NAME);
      setupUserSheetHeaders(usuariosSheet);
    }
    
    return { success: true, message: 'Hojas configuradas correctamente' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// Función para manejar GET (para reportes)
function doGet(e) {
  return HtmlService.createHtmlOutput(`
    <h1>Backend de Control de Asistencia</h1>
    <p>Funcionando correctamente</p>
    <p>${new Date().toISOString()}</p>
  `);
}

// Función para ejecutar manualmente
function initialSetup() {
  const result = createSheetsIfNotExist();
  console.log(result.message);
  return result;
}