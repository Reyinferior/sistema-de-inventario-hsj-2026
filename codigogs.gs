// ============================================================
//  INVENTARIO DE EQUIPOS DE CÓMPUTO — HOSPITAL SAN JOSÉ
//  Code.gs  |  Google Apps Script
//  Coloca este código en tu proyecto de Apps Script junto
//  con el archivo InventarioEquipos.html
// ============================================================

// Nombre de la hoja de cálculo donde se guardan los equipos
var HOJA_EQUIPOS = "Equipos";

// Columnas en el mismo orden en que se guardan las filas
var COLUMNAS = [
  "nombre_pc","numero_ip","unidad","oficina","area","usuario","cod_patrimonio","fecha_adq","n2",
  "marca_cpu","modelo_cpu","serie_cpu","estado_cpu",
  "marca_mb","modelo_mb",
  "marca_proc","procesador","velocidad_proc","gen_proc",
  "tipo_mem","total_mem",
  "marca_disco","cap_disco","modelo_disco","tipo_disco",
  "marca_tvideo","modelo_tvideo",
  "marca_dvd","tipo_dvd","velocidad_grabadora",
  "marca_teclado","modelo_teclado","serie_teclado","cod_teclado","estado_teclado",
  "marca_mouse","modelo_mouse","serie_mouse","estado_mouse",
  "marca_monitor","tipo_monitor","modelo_monitor","pulgadas_monitor",
  "serie_monitor","cod_monitor","estado_monitor",
  "marca_estab","modelo_estab","serie_estab","cod_estab","estado_estab",
  "nombre_so","num_so","key_so","software_oficina","key_software_oficina","internet","antivirus",
  "fecha_registro"   // se agrega automáticamente al guardar
];

// Encabezados legibles para la hoja (mismo orden que COLUMNAS)
var ENCABEZADOS = [
  "Nombre PC","Número de ip","Unidad","Oficina","área","Usuario","Cod. Patrimonio","Fecha Adquisición","N2",
  "Marca CPU","Modelo CPU","N° Serie CPU","Estado CPU",
  "Marca Mainboard","Modelo Mainboard",
  "Marca Procesador","Procesador","Velocidad Procesador","  ",
  "Tipo Memoria","Total Memoria (GB)",
  "Marca Disco","Capacidad Disco","Modelo Disco","Tipo Disco",
  "Marca T.Video","Modelo T.Video",
  "Marca Grabadora DVD","Tipo Grabadora DVD","Velocidad Grabadora",
  "Marca Teclado","Modelo Teclado","N° Serie Teclado","Cod. Teclado","Estado Teclado",
  "Marca Mouse","Modelo Mouse","N° Serie Mouse","Estado Mouse",
  "Marca Monitor","Tipo Monitor","Modelo Monitor","Monitor Pulgadas",
  "N° Serie Monitor","Cod. Monitor","Estado Monitor",
  "Marca Estabilizador","Modelo Estabilizador","N° Serie Estabilizador",
  "Cod. Estabilizador","Estado Estabilizador",
  "Sistema Operativo","N° Licencia SO","Key SO",
  "Software Oficina","key software de oficina","Internet","Antivirus",
  "Fecha Registro"
];

// ---- PUNTO DE ENTRADA WEB ----
function doGet(e) {
  var pagina = e && e.parameter && e.parameter.pagina;
  if (pagina === 'formulario') {
    return HtmlService.createHtmlOutputFromFile('formulario')
      .setTitle('Reportar Problema — Hospital San José')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Inventario Equipos — Hospital San José')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
//  INICIALIZACIÓN
// ============================================================
function inicializarHojasEquipos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);

    if (!hoja) {
      hoja = ss.insertSheet(HOJA_EQUIPOS);
      hoja.appendRow(ENCABEZADOS);

      // Formato de encabezados
      var rango = hoja.getRange(1, 1, 1, ENCABEZADOS.length);
      rango.setBackground('#1C5F7B');
      rango.setFontColor('#FFFFFF');
      rango.setFontWeight('bold');
      rango.setFontSize(10);
      hoja.setFrozenRows(1);
      hoja.setColumnWidths(1, ENCABEZADOS.length, 130);

      return 'Hoja "' + HOJA_EQUIPOS + '" creada correctamente con ' + ENCABEZADOS.length + ' columnas.';
    }

    return 'La hoja "' + HOJA_EQUIPOS + '" ya existe. No se realizaron cambios.';
  } catch (err) {
    throw new Error('Error al inicializar: ' + err.message);
  }
}

// ============================================================
//  REGISTRAR EQUIPO
// ============================================================
function registrarEquipo(equipo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);

    if (!hoja) {
      inicializarHojasEquipos();
      hoja = ss.getSheetByName(HOJA_EQUIPOS);
    }

    // Validación básica
    if (!equipo.nombre_pc || !equipo.usuario) {
      throw new Error('Nombre PC y Usuario son obligatorios.');
    }

    // Verificar que no exista el mismo nombre_pc
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().toUpperCase() === equipo.nombre_pc.toUpperCase()) {
        throw new Error('Ya existe un equipo con el nombre "' + equipo.nombre_pc + '". Use un nombre único.');
      }
    }

    // Agregar fecha de registro automáticamente
    equipo.fecha_registro = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // Construir fila en el orden de COLUMNAS
    var fila = COLUMNAS.map(function(col) {
      return equipo[col] !== undefined ? equipo[col] : '';
    });

    hoja.appendRow(fila);

    // Auto-ajustar alto de la nueva fila
    var ultimaFila = hoja.getLastRow();
    hoja.setRowHeight(ultimaFila, 22);

    // Alternar color de filas
    if (ultimaFila % 2 === 0) {
      hoja.getRange(ultimaFila, 1, 1, COLUMNAS.length).setBackground('#F0F8FC');
    }

    limpiarCache();


    return 'Equipo "' + equipo.nombre_pc + '" registrado correctamente.';
  } catch (err) {
    throw new Error(err.message);
  }
}

// ============================================================
//  OBTENER TODOS LOS EQUIPOS
// ============================================================
function obtenerEquipos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);

    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var resultado = [];

    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[0]) continue;

      var obj = {};
      COLUMNAS.forEach(function(col, idx) {
        var val = fila[idx];

        // Si el valor es una fecha, formatearla como dd/mm/yyyy
        if (val instanceof Date && !isNaN(val.getTime())) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        } else {
          val = val !== undefined && val !== null ? val.toString() : '';
        }

        obj[col] = val;
      });

      resultado.push(obj);
    }

     limpiarCache();

    return resultado;
  } catch (err) {
    throw new Error('Error al obtener equipos: ' + err.message);
  }
}
// ============================================================
//  RESUMEN PARA DASHBOARD
// ============================================================
function obtenerResumenEquipos() {
  try {
    var equipos = obtenerEquipos();
    var bueno = 0, regular = 0, malo = 0;

    equipos.forEach(function(e) {
      var est = (e.estado_cpu || '').toLowerCase();
      if (est === 'bueno') bueno++;
      else if (est === 'regular') regular++;
      else if (est === 'malo' || est === 'de baja') malo++;
    });

    return {
      total: equipos.length,
      bueno: bueno,
      regular: regular,
      malo: malo,
      equipos: equipos
    };
  } catch (err) {
    throw new Error('Error en resumen: ' + err.message);
  }
}

// ============================================================
//  BUSCAR EQUIPOS
// ============================================================
function buscarEquipo(texto) {
  try {
    var equipos = obtenerEquipos();
    if (!texto) return equipos;

    var buscar = texto.toString().toLowerCase();
    var camposBusqueda = ['nombre_pc','usuario','area','unidad','cod_patrimonio',
                          'marca_cpu','procesador','marca_monitor','nombre_so','numero_ip'];

    return equipos.filter(function(e) {
      return camposBusqueda.some(function(campo) {
        return e[campo] && e[campo].toString().toLowerCase().indexOf(buscar) !== -1;
      });
    });
  } catch (err) {
    throw new Error('Error en búsqueda: ' + err.message);
  }
}

// ============================================================
//  VALIDAR SISTEMA
// ============================================================
function validarSistemaEquipos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);

    if (!hoja) {
      return { ok: false, mensaje: 'La hoja "' + HOJA_EQUIPOS + '" no existe. Use "Inicializar Hojas".' };
    }

    var encHoja = hoja.getRange(1, 1, 1, COLUMNAS.length).getValues()[0];
    var totalEquipos = Math.max(0, hoja.getLastRow() - 1);

    return {
      ok: true,
      mensaje: 'Sistema OK. Hoja "' + HOJA_EQUIPOS + '" con ' + COLUMNAS.length +
               ' columnas y ' + totalEquipos + ' equipo(s) registrado(s).'
    };
  } catch (err) {
    return { ok: false, mensaje: 'Error en validación: ' + err.message };
  }
}

// ============================================================
//  FUNCIÓN DE PRUEBA (ejecutar desde el editor de Apps Script)
// ============================================================
function testRegistrar() {
  var equipoDemo = {
    nombre_pc: 'PC-TEST-001',
    unidad: 'Administración',
    area: 'Contabilidad',
    usuario: 'Juan Pérez',
    cod_patrimonio: 'PAT-2024-001',
    fecha_adq: '2024-01-15',
    n2: 'N2-001',
    marca_cpu: 'HP',
    modelo_cpu: 'ProDesk 400 G7',
    serie_cpu: 'SN123456',
    estado_cpu: 'Bueno',
    marca_mb: 'Intel',
    modelo_mb: 'H470',
    marca_proc: 'Intel',
    procesador: 'Core i5-10400',
    velocidad_proc: '2.90 GHz',
    gen_proc: '10ma generación',
    tipo_mem: 'DDR4',
    total_mem: '8',
    marca_disco: 'Seagate',
    cap_disco: '1TB',
    modelo_disco: 'ST1000DM010',
    tipo_disco: 'HDD',
    marca_tvideo: 'Intel',
    modelo_tvideo: 'UHD 630',
    marca_dvd: 'LG',
    tipo_dvd: 'DVD-RW',
    marca_teclado: 'HP',
    modelo_teclado: 'SK-2086',
    serie_teclado: 'TEC001',
    cod_teclado: 'TEC-PAT-001',
    estado_teclado: 'Bueno',
    marca_mouse: 'HP',
    modelo_mouse: 'MO-001',
    serie_mouse: 'MOU001',
    estado_mouse: 'Bueno',
    marca_monitor: 'LG',
    tipo_monitor: 'LED',
    modelo_monitor: '22MK430H',
    pulgadas_monitor: '21.5',
    serie_monitor: 'MON001',
    cod_monitor: 'MON-PAT-001',
    estado_monitor: 'Bueno',
    marca_estab: 'APC',
    modelo_estab: 'ES 550VA',
    serie_estab: 'EST001',
    cod_estab: 'EST-PAT-001',
    estado_estab: 'Bueno',
    nombre_so: 'Windows 10 Pro',
    num_so: 'LIC-001',
    key_so: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
    software_oficina: 'Microsoft Office 2021',
    internet: 'Sí',
    antivirus: 'Windows Defender'
  };

  Logger.log(inicializarHojasEquipos());
  Logger.log(registrarEquipo(equipoDemo));
  Logger.log('Total equipos: ' + obtenerEquipos().length);
}


// ============================================================
//  ACTUALIZAR EQUIPO
// ============================================================
function actualizarEquipo(equipo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);
    if (!hoja) throw new Error('Hoja no encontrada.');

    var datos = hoja.getDataRange().getValues();
    var filaIdx = -1;

    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().toUpperCase() === equipo.nombre_pc.toUpperCase()) {
        filaIdx = i + 1; // +1 porque las filas en Sheets empiezan en 1
        break;
      }
    }

    if (filaIdx === -1) throw new Error('No se encontró el equipo "' + equipo.nombre_pc + '".');

    // Conservar la fecha de registro original
    equipo.fecha_registro = datos[filaIdx - 1][COLUMNAS.indexOf('fecha_registro')];
    if (equipo.fecha_registro instanceof Date) {
      equipo.fecha_registro = Utilities.formatDate(equipo.fecha_registro, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }

    var fila = COLUMNAS.map(function(col) {
      return equipo[col] !== undefined ? equipo[col] : '';
    });

    hoja.getRange(filaIdx, 1, 1, fila.length).setValues([fila]);
     limpiarCache();

    return 'Equipo "' + equipo.nombre_pc + '" actualizado correctamente.';
  } catch (err) {
    throw new Error(err.message);
  }
}

// ============================================================
//  ELIMINAR EQUIPO
// ============================================================
function eliminarEquipo(nombre_pc) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);
    if (!hoja) throw new Error('Hoja no encontrada.');

    var datos = hoja.getDataRange().getValues();

    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().toUpperCase() === nombre_pc.toUpperCase()) {
        hoja.deleteRow(i + 1);
        limpiarCache();
        return 'Equipo "' + nombre_pc + '" eliminado correctamente.';
      }
    }

    throw new Error('No se encontró el equipo "' + nombre_pc + '".');
  } catch (err) {
    throw new Error(err.message);
  }
}


//OBTENER LISTA 


function obtenerListas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Listas");
    if (!hoja) return {};
 
    var datos = hoja.getDataRange().getValues();
    var resultado = {};
    var encabezados = datos[0];
 
    encabezados.forEach(function(enc, colIdx) {
      if (!enc) return;
      var valores = [];
      for (var i = 1; i < datos.length; i++) {
        var val = datos[i][colIdx];
        if (val && val.toString().trim() !== '') {
          valores.push(val.toString().trim());
        }
      }
      resultado[enc] = valores;
    });
 
    return resultado;
  } catch (err) {
    throw new Error('Error al obtener listas: ' + err.message);
  }
}




// ---- CONFIGURACIÓN ----
var HOJA_PERSONAL      = "Personal";
var HOJA_MANTENIMIENTO = "Mantenimiento";
//va en el js 
var COLUMNAS_PERSONAL = [
  "dni","nombres_completos",
  "cod_modalidad","des_modalidad",
  "cod_cargo","des_cargo",
  "cod_servicio","des_servicio","equipo_asignado","fecha_nac",
];
 

//en el gs 
var ENCABEZADOS_PERSONAL = [
  "DNI","Nombres y Apellidos",
  "Cod_Modalidad","DES_MODALIDAD",
  "COD_DESCARGO","DES_CARGO",
  "COD_SERVICIO","DES_SERVICIO",
  "Equipo Asignado"
  ,"FEC_NACIMIENTO"
];
 
var COLUMNAS_MANT = [
  "id_mant","equipo","usuario","tecnico_nombre",
  "tipo","estado",
  "fecha_prog","fecha_aten",
  "problema","solucion","observaciones",
  "fecha_registro"
];

var ENCABEZADOS_MANT = [
  "ID Mant.","Equipo","Usuario","Nombre Técnico",
  "Tipo","Estado",
  "Fecha Programada","Fecha Atención",
  "Problema","Solución","Observaciones",
  "Fecha Registro"
];
// ============================================================
//  INICIALIZAR HOJAS PERSONAL Y MANTENIMIENTO
//  (llama a esto una sola vez desde Configuración o el editor)
// ============================================================
function inicializarHojasPersonalMant() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msg = [];
 
  [
    { nombre: HOJA_PERSONAL,      encabezados: ENCABEZADOS_PERSONAL },
    { nombre: HOJA_MANTENIMIENTO, encabezados: ENCABEZADOS_MANT }
  ].forEach(function(cfg) {
    var hoja = ss.getSheetByName(cfg.nombre);
    if (!hoja) {
      hoja = ss.insertSheet(cfg.nombre);
      hoja.appendRow(cfg.encabezados);
      var rng = hoja.getRange(1, 1, 1, cfg.encabezados.length);
      rng.setBackground('#1C5F7B');
      rng.setFontColor('#FFFFFF');
      rng.setFontWeight('bold');
      rng.setFontSize(10);
      hoja.setFrozenRows(1);
      hoja.setColumnWidths(1, cfg.encabezados.length, 130);
      msg.push('Hoja "' + cfg.nombre + '" creada.');
    } else {
      msg.push('Hoja "' + cfg.nombre + '" ya existe.');
    }
  });
 
  // Asegurarse de que la hoja Listas tenga las columnas de catálogos
  var hojaListas = ss.getSheetByName("Listas");
  if (hojaListas) {
    msg.push('Recuerda agregar columnas "modalidades", "cargos", "servicios" en la hoja Listas con formato COD|DESCRIPCION por fila.');
  }
 
  return msg.join('\n');
}
 
// ============================================================
//  PERSONAL — REGISTRAR
// ============================================================
function registrarPersonal(persona) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = _getOrCreateSheet(ss, HOJA_PERSONAL, ENCABEZADOS_PERSONAL);
 
    if (!persona.dni) throw new Error('El DNI es obligatorio.');
 
    // Verificar DNI único
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString() === persona.dni.toString()) {
        throw new Error('Ya existe un registro con el DNI "' + persona.dni + '".');
      }
    }
 
    persona.fecha_registro = _timestamp();
    var fila = COLUMNAS_PERSONAL.map(function(c) { return persona[c] || ''; });
    hoja.appendRow(fila);
    _colorearFila(hoja, hoja.getLastRow(), COLUMNAS_PERSONAL.length);

     limpiarCache();
 
    return 'Personal "' + persona.apellidos + ', ' + persona.nombres + '" registrado correctamente.';
  } catch(err) { throw new Error(err.message); }
}
 
// ============================================================
//  PERSONAL — OBTENER TODOS
// ============================================================
function obtenerPersonal() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_PERSONAL);
    
    if (!hoja) {
      // Si no existe la hoja, inicializarla y devolver vacío
      inicializarHojasPersonalMant();
      return [];
    }
    
    var lastRow = hoja.getLastRow();
    var lastCol = hoja.getLastColumn();
    
    // Si solo tiene encabezados o está vacía
    if (lastRow < 2 || lastCol < 1) return [];

    // Leer solo el rango con datos reales (más eficiente)
    var numCols = Math.min(lastCol, COLUMNAS_PERSONAL.length);
    var datos = hoja.getRange(1, 1, lastRow, numCols).getValues();
    
    var resultado = [];
    for (var i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      var obj = {};
      COLUMNAS_PERSONAL.forEach(function(col, idx) {
        if (idx >= numCols) { obj[col] = ''; return; }
        var val = datos[i][idx];
        if (val instanceof Date && !isNaN(val)) {
          val = Utilities.formatDate(
            val, Session.getScriptTimeZone(), 'dd/MM/yyyy'
          );
        } else {
          val = (val !== undefined && val !== null) ? val.toString() : '';
        }
        obj[col] = val;
      });
      resultado.push(obj);
    }
    return resultado;

  } catch(err) {
    throw new Error('Error al obtener personal: ' + err.message);
  }
}
 
// ============================================================
//  PERSONAL — ACTUALIZAR
// ============================================================
function actualizarPersonal(persona) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_PERSONAL);
    if (!hoja) throw new Error('Hoja Personal no encontrada.');
 
    var datos = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString() === persona.dni.toString()) {
        filaIdx = i + 1; break;
      }
    }
    if (filaIdx === -1) throw new Error('No se encontró el DNI "' + persona.dni + '".');
 
    // Conservar fecha_registro original
    persona.fecha_registro = datos[filaIdx - 1][COLUMNAS_PERSONAL.indexOf('fecha_registro')];
    if (persona.fecha_registro instanceof Date) {
      persona.fecha_registro = Utilities.formatDate(persona.fecha_registro, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }
 
    var fila = COLUMNAS_PERSONAL.map(function(c) { return persona[c] || ''; });
    hoja.getRange(filaIdx, 1, 1, fila.length).setValues([fila]);

    limpiarCache();

    return 'Personal actualizado correctamente.';
  } catch(err) { throw new Error(err.message); }
}
 
// ============================================================
//  PERSONAL — ELIMINAR
// ============================================================
function eliminarPersonal(dni) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_PERSONAL);
    if (!hoja) throw new Error('Hoja Personal no encontrada.');
 
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString() === dni.toString()) {
        hoja.deleteRow(i + 1);

        limpiarCache();

        return 'Personal con DNI "' + dni + '" eliminado correctamente.';
      }
    }
    throw new Error('No se encontró el DNI "' + dni + '".');
  } catch(err) { throw new Error(err.message); }
}
 
// ============================================================
//  MANTENIMIENTO — REGISTRAR
// ============================================================
function registrarMantenimiento(mant) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = _getOrCreateSheet(ss, HOJA_MANTENIMIENTO, ENCABEZADOS_MANT);
 
    if (!mant.equipo) {
  throw new Error('El equipo es obligatorio.');
}
if (!mant.estado) mant.estado = 'Pendiente';
    // Generar ID único
    var lastRow = hoja.getLastRow();
    mant.id_mant = 'MANT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + (lastRow).toString().padStart(4, '0');
    mant.fecha_registro = _timestamp();
 
    var fila = COLUMNAS_MANT.map(function(c) { return mant[c] || ''; });
    hoja.appendRow(fila);
    _colorearFila(hoja, hoja.getLastRow(), COLUMNAS_MANT.length);


    limpiarCache();
 
    return 'Mantenimiento "' + mant.id_mant + '" registrado correctamente.';
  } catch(err) { throw new Error(err.message); }
}
 
// ============================================================
//  MANTENIMIENTO — OBTENER TODOS
// ============================================================
function obtenerMantenimientos() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_MANTENIMIENTO);
    if (!hoja || hoja.getLastRow() < 2) return [];
 
    var datos = hoja.getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      var obj = {};
      COLUMNAS_MANT.forEach(function(col, idx) {
        var val = datos[i][idx];
        if (val instanceof Date && !isNaN(val)) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        } else {
          val = val !== undefined && val !== null ? val.toString() : '';
        }
        obj[col] = val;
      });
      resultado.push(obj);
    }
    // Ordenar: Pendiente primero, luego Programado, luego Terminado
    var orden = { 'Pendiente': 0, 'Programado': 1, 'Terminado': 2 };
    resultado.sort(function(a, b) {
      return (orden[a.estado] || 99) - (orden[b.estado] || 99);
    });
    return resultado;
  } catch(err) { throw new Error('Error al obtener mantenimientos: ' + err.message); }
}
 
// ============================================================
//  MANTENIMIENTO — ACTUALIZAR
// ============================================================
function actualizarMantenimiento(mant) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_MANTENIMIENTO);
    if (!hoja) throw new Error('Hoja Mantenimiento no encontrada.');
 
    var datos = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString() === mant.id_mant.toString()) {
        filaIdx = i + 1; break;
      }
    }
    if (filaIdx === -1) throw new Error('No se encontró el mantenimiento "' + mant.id_mant + '".');
 
    mant.fecha_registro = datos[filaIdx - 1][COLUMNAS_MANT.indexOf('fecha_registro')];
    if (mant.fecha_registro instanceof Date) {
      mant.fecha_registro = Utilities.formatDate(mant.fecha_registro, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }
 
    var fila = COLUMNAS_MANT.map(function(c) { return mant[c] || ''; });
    hoja.getRange(filaIdx, 1, 1, fila.length).setValues([fila]);

    limpiarCache();


    return 'Mantenimiento actualizado correctamente.';
  } catch(err) { throw new Error(err.message); }
}
 
// ============================================================
//  MANTENIMIENTO — ELIMINAR
// ============================================================
function eliminarMantenimiento(id_mant) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_MANTENIMIENTO);
    if (!hoja) throw new Error('Hoja Mantenimiento no encontrada.');
 
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString() === id_mant.toString()) {
        hoja.deleteRow(i + 1);

         limpiarCache();

        return 'Mantenimiento "' + id_mant + '" eliminado correctamente.';
      }
    }
    throw new Error('No se encontró el mantenimiento "' + id_mant + '".');
  } catch(err) { throw new Error(err.message); }
}
 
// ============================================================
//  HELPERS INTERNOS
// ============================================================
function _getOrCreateSheet(ss, nombre, encabezados) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.appendRow(encabezados);
    var rng = hoja.getRange(1, 1, 1, encabezados.length);
    rng.setBackground('#1C5F7B');
    rng.setFontColor('#FFFFFF');
    rng.setFontWeight('bold');
    rng.setFontSize(10);
    hoja.setFrozenRows(1);
    hoja.setColumnWidths(1, encabezados.length, 130);
  }
  return hoja;
}
 
function _timestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
 
function _colorearFila(hoja, fila, numCols) {
  hoja.setRowHeight(fila, 22);
  if (fila % 2 === 0) {
    hoja.getRange(fila, 1, 1, numCols).setBackground('#F0F8FC');
  }
}
 
var HOJA_TECNICOS = "Tecnicos";

function obtenerTecnicos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_TECNICOS);
    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < datos.length; i++) {
      var nombre = datos[i][0];
      if (nombre && nombre.toString().trim() !== '') {
        resultado.push({ nombre: nombre.toString().trim() });
      }
    }
    return resultado;
  } catch(err) {
    throw new Error('Error al obtener técnicos: ' + err.message);
  }
}



function obtenerCargos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Cargo");
    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < datos.length; i++) {
      var cod = datos[i][0];
      var des = datos[i][1];
      if (cod && cod.toString().trim() !== '') {
        resultado.push({ 
          codigo: cod.toString().trim(), 
          descripcion: des ? des.toString().trim() : '' 
        });
      }
    }
    return resultado;
  } catch(err) {
    throw new Error('Error al obtener cargos: ' + err.message);
  }
}

function obtenerOficinas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("oficina");
    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < datos.length; i++) {
      var cod = datos[i][0];
      var des = datos[i][1];
      if (cod && cod.toString().trim() !== '') {
        resultado.push({ 
          codigo: cod.toString().trim(), 
          descripcion: des ? des.toString().trim() : '' 
        });
      }
    }
    return resultado;
  } catch(err) {
    throw new Error('Error al obtener oficinas: ' + err.message);
  }
}

// ============================================================
//  OBTENER MODALIDADES (fix nombre con 's' y typo)
// ============================================================
function obtenerModalidades() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("modalidad");
    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < datos.length; i++) {
      var cod = datos[i][0];
      var des = datos[i][1];
      if (cod && cod.toString().trim() !== '') {
        resultado.push({
          codigo: cod.toString().trim(),
          descripcion: des ? des.toString().trim() : ''
        });
      }
    }
    return resultado;
  } catch(err) {
    throw new Error('Error al obtener modalidades: ' + err.message);
  }
}

// ============================================================
//  OBTENER NOMBRES DE EQUIPOS PARA SELECTORES
// ============================================================
function obtenerEquiposParaSelector() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);
    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var resultado = [];
    for (var i = 1; i < datos.length; i++) {
      var nombre = datos[i][0]; // columna nombre_pc
      if (nombre && nombre.toString().trim() !== '') {
        resultado.push(nombre.toString().trim());
      }
    }
    return resultado;
  } catch(err) {
    throw new Error('Error al obtener equipos para selector: ' + err.message);
  }
}


function obtenerEquiposPorUsuario(nombreUsuario) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_EQUIPOS);
    if (!hoja || hoja.getLastRow() < 2) return [];

    var datos = hoja.getDataRange().getValues();
    var resultado = [];
    var buscar = nombreUsuario.toString().toLowerCase().trim();

    for (var i = 1; i < datos.length; i++) {
      var usuario = datos[i][5]; // columna Usuario (índice 5)
      var nombrePc = datos[i][0]; // columna Nombre PC (índice 0)
      if (usuario && usuario.toString().toLowerCase().indexOf(buscar) !== -1) {
        if (nombrePc && nombrePc.toString().trim() !== '') {
          resultado.push(nombrePc.toString().trim());
        }
      }
    }
    return resultado;
  } catch(err) {
    throw new Error('Error: ' + err.message);
  }
}


// ... todo tu código actual ...
// ... obtenerEquiposPorUsuario() es la última función que tienes ...

// ============================================================
//  CACHÉ — agregar al final del Code.gs
// ============================================================

var CACHE_TTL = 300; // 5 minutos

function obtenerPersonalCached() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('personal_data');
  if (cached) return JSON.parse(cached);
  
  var data = obtenerPersonal();
  try { cache.put('personal_data', JSON.stringify(data), CACHE_TTL); } catch(e) {}
  return data;
}

function obtenerEquiposCached() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('equipos_data');
  if (cached) return JSON.parse(cached);
  
  var data = obtenerEquipos();
  try { cache.put('equipos_data', JSON.stringify(data), CACHE_TTL); } catch(e) {}
  return data;
}

function obtenerMantenimientosCached() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('mant_data');
  if (cached) return JSON.parse(cached);
  
  var data = obtenerMantenimientos();
  try { cache.put('mant_data', JSON.stringify(data), CACHE_TTL); } catch(e) {}
  return data;
}

function limpiarCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['personal_data', 'equipos_data', 'mant_data']);
}


// =========================================================
//  MÓDULO IMPRESORAS (Pega esto en tu Code.gs de Google Apps Script)
// =========================================================

function inicializarHojasImpresoras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const encabInventario = ['N°', 'UNIDAD_DPTO', 'AREA_SERVICIO', 'USUARIO', 'CODPAT', 'COD_INV', 'TIPO_EQUIPO', 'MARCA', 'MODELO', 'TIPO_IMPRESORA', 'SERIE', 'MAC', 'UBICACION', 'OBSERVACION', 'ADM_ASIS', 'COINCIDENTE', 'BAJA', 'ESTA_SN', 'ESTADO', 'NOTA', 'FECHA_REGISTRO'];
  _getOrCreateSheet(ss, 'Inventario_de_Impresoras', encabInventario);
  
  const encabListas = ['marcas_de_impresoras', 'tipos_de_impresoras', 'estado', 'modelo'];
  _getOrCreateSheet(ss, 'Impresoras', encabListas);
  return 'Hoja de Impresoras configurada exitosamente.';
}

function registrarImpresora(imp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Inventario_de_Impresoras');
  if (!hoja) throw new Error('No existe la hoja Inventario_de_Impresoras');

  // Calcular el correlativo (N°)
  const ultimaFila = hoja.getLastRow();
  let numero = 1;
  if (ultimaFila > 1) {
    const ultimoNum = hoja.getRange(ultimaFila, 1).getValue();
    if (!isNaN(ultimoNum) && ultimoNum !== "") numero = Number(ultimoNum) + 1;
  }
  
  const fechaReg = _timestamp();
  const fila = [
    numero,
    imp.unidad_dpto || '',
    imp.area_servicio || '',
    imp.usuario || '',
    imp.codpat || '',
    imp.cod_inv || '',
    imp.tipo_equipo || '',
    imp.marca || '',
    imp.modelo || '',
    imp.tipo_impresora || '',
    imp.serie || '',
    imp.mac || '',
    imp.ubicacion || '',
    imp.observacion || '',
    imp.adm_asis || '',
    imp.coincidente || '',
    imp.baja || '',
    imp.esta_sn || '',
    imp.estado || '',
    imp.nota || '',
    fechaReg
  ];
  
  hoja.appendRow(fila);
  CacheService.getScriptCache().remove('impresoras_data');
  return 'Impresora guardada correctamente';
}

function obtenerImpresoras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Inventario_de_Impresoras');
  if (!hoja) return [];
  const datos = hoja.getDataRange().getDisplayValues();
  if (datos.length <= 1) return [];
  const headers = datos[0];
  const resultado = [];
  for (let i = 1; i < datos.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      // Normalizamos el header para que sirva de clave en objetos JS (elimina caracteres especiales)
      let head = headers[j].toLowerCase().replace(/[^a-z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
      if(head === 'n' || head === '') head = 'numero';
      obj[head] = datos[i][j];
    }
    obj._fila = i + 1;
    resultado.push(obj);
  }
  return resultado;
}

function obtenerImpresorasCached() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('impresoras_data');
  if (cached) return JSON.parse(cached);
  const data = obtenerImpresoras();
  cache.put('impresoras_data', JSON.stringify(data), 600); // 10 minutos
  return data;
}

function actualizarImpresora(imp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Inventario_de_Impresoras');
  const datos = hoja.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    const codpat = datos[i][4]; // Columna E = CODPAT (índice 4)
    const serie = datos[i][10]; // Columna K = SERIE (índice 10)
    
    // Verificamos por Número de Serie si existe, sino por CODPAT
    if ((imp.serie && serie === imp.serie) || (imp.codpat && codpat === imp.codpat)) {
      const filaUpdate = [
        datos[i][0], // N° se mantiene
        imp.unidad_dpto || '',
        imp.area_servicio || '',
        imp.usuario || '',
        imp.codpat || '',
        imp.cod_inv || '',
        imp.tipo_equipo || '',
        imp.marca || '',
        imp.modelo || '',
        imp.tipo_impresora || '',
        imp.serie || '',
        imp.mac || '',
        imp.ubicacion || '',
        imp.observacion || '',
        imp.adm_asis || '',
        imp.coincidente || '',
        imp.baja || '',
        imp.esta_sn || '',
        imp.estado || '',
        imp.nota || '',
        datos[i][20] // FECHA_REGISTRO se mantiene
      ];
      hoja.getRange(i + 1, 1, 1, 21).setValues([filaUpdate]);
      CacheService.getScriptCache().remove('impresoras_data');
      return 'Impresora actualizada correctamente';
    }
  }
  throw new Error('Impresora no encontrada en el inventario.');
}

function eliminarImpresora(identificador) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Inventario_de_Impresoras');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    const codpat = datos[i][4]; 
    const serie = datos[i][10]; 
    if (serie == identificador || codpat == identificador) {
      hoja.deleteRow(i + 1);
      CacheService.getScriptCache().remove('impresoras_data');
      return 'Impresora eliminada exitosamente';
    }
  }
  throw new Error('No se encontró la impresora a eliminar');
}

function obtenerListasImpresoras() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Impresoras');
  if (!hoja) return { marcas: [], tipos: [], estados: [], modelos: [] };
  
  const datos = hoja.getDataRange().getDisplayValues();
  const result = { marcas: [], tipos: [], estados: [], modelos: [] };
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0]) result.marcas.push(datos[i][0].trim());
    if (datos[i][1]) result.tipos.push(datos[i][1].trim());
    if (datos[i][2]) result.estados.push(datos[i][2].trim());
    if (datos[i][3]) result.modelos.push(datos[i][3].trim());
  }
  
  return {
    marcas: [...new Set(result.marcas)],
    tipos: [...new Set(result.tipos)],
    estados: [...new Set(result.estados)],
    modelos: [...new Set(result.modelos)]
  };
}

// ============================================================
//  MÓDULO TICKETS
// ============================================================

var HOJA_TICKETS = "Tickets";

var COLUMNAS_TICKETS = [
  "id_ticket","usuario","area","equipo","problema",
  "prioridad","estado","tecnico","nota_tecnico",
  "origen","fecha_registro","fecha_actualizacion"
];

var ENCABEZADOS_TICKETS = [
  "ID Ticket","Usuario","Área / Servicio","Equipo Afectado","Descripción del Problema",
  "Prioridad","Estado","Técnico Asignado","Nota del Técnico",
  "Origen","Fecha Registro","Última Actualización"
];

function _inicializarHojaTickets(ss) {
  var hoja = ss.getSheetByName(HOJA_TICKETS);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_TICKETS);
    hoja.appendRow(ENCABEZADOS_TICKETS);
    var rng = hoja.getRange(1, 1, 1, ENCABEZADOS_TICKETS.length);
    rng.setBackground('#1C5F7B');
    rng.setFontColor('#FFFFFF');
    rng.setFontWeight('bold');
    rng.setFontSize(10);
    hoja.setFrozenRows(1);
    hoja.setColumnWidths(1, ENCABEZADOS_TICKETS.length, 150);
  }
  return hoja;
}

function _generarIdTicket(hoja) {
  var lastRow = hoja.getLastRow();
  var num = lastRow; // fila 1 = encabezado, filas siguientes = datos
  return 'TKT-' + String(num).padStart(4, '0');
}

function obtenerTicketsCached() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('tickets_data');
    if (cached) return JSON.parse(cached);
    var datos = _leerTickets();
    cache.put('tickets_data', JSON.stringify(datos), 120);
    return datos;
  } catch(err) {
    throw new Error('Error al obtener tickets: ' + err.message);
  }
}

function _leerTickets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = _inicializarHojaTickets(ss);
  if (hoja.getLastRow() < 2) return [];
  var datos = hoja.getDataRange().getValues();
  var resultado = [];
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (!fila[0]) continue;
    var obj = {};
    COLUMNAS_TICKETS.forEach(function(col, idx) { obj[col] = fila[idx] ? fila[idx].toString() : ''; });
    resultado.push(obj);
  }
  return resultado.reverse(); // más recientes primero
}

function registrarTicket(ticket) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = _inicializarHojaTickets(ss);
    var id = _generarIdTicket(hoja);
    var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    var fila = [
      id,
      ticket.usuario || '',
      ticket.area || '',
      ticket.equipo || '',
      ticket.problema || '',
      ticket.prioridad || 'Media',
      'Nuevo',
      '',
      '',
      ticket.origen || 'Formulario',
      ahora,
      ahora
    ];
    hoja.appendRow(fila);
    CacheService.getScriptCache().remove('tickets_data');
    return 'Ticket ' + id + ' registrado correctamente.';
  } catch(err) {
    throw new Error('Error al registrar ticket: ' + err.message);
  }
}

function actualizarTicket(ticket) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = _inicializarHojaTickets(ss);
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString() === ticket.id_ticket) {
        var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        hoja.getRange(i + 1, 7).setValue(ticket.estado || datos[i][6]);
        hoja.getRange(i + 1, 8).setValue(ticket.tecnico || datos[i][7]);
        hoja.getRange(i + 1, 9).setValue(ticket.nota_tecnico || datos[i][8]);
        hoja.getRange(i + 1, 12).setValue(ahora);
        CacheService.getScriptCache().remove('tickets_data');
        return 'Ticket ' + ticket.id_ticket + ' actualizado correctamente.';
      }
    }
    throw new Error('Ticket no encontrado.');
  } catch(err) {
    throw new Error('Error al actualizar ticket: ' + err.message);
  }
}

function eliminarTicket(id) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = _inicializarHojaTickets(ss);
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString() === id) {
        hoja.deleteRow(i + 1);
        CacheService.getScriptCache().remove('tickets_data');
        return 'Ticket ' + id + ' eliminado correctamente.';
      }
    }
    throw new Error('Ticket no encontrado.');
  } catch(err) {
    throw new Error('Error al eliminar ticket: ' + err.message);
  }
}

// ============================================================
//  LEER TICKETS DESDE GMAIL (trigger automático)
//  Configura un trigger: Editar > Triggers > leerTicketsDeGmail
//  Tipo: Time-driven, cada 5 o 10 minutos
// ============================================================
function leerTicketsDeGmail() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = _inicializarHojaTickets(ss);
    // Busca correos no leídos con etiqueta "soporte-ti" o asunto que contenga [TICKET] o [SOPORTE]
    var hilos = GmailApp.search('is:unread subject:([TICKET] OR [SOPORTE] OR soporte ti) newer_than:1d', 0, 20);
    if (hilos.length === 0) return 'Sin correos nuevos.';
    var procesados = 0;
    hilos.forEach(function(hilo) {
      var mensajes = hilo.getMessages();
      var msg = mensajes[mensajes.length - 1]; // último mensaje del hilo
      var asunto = msg.getSubject();
      var cuerpo  = msg.getPlainBody().substring(0, 500);
      var remitente = msg.getFrom();
      var fecha = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      // Extraer nombre del remitente (antes del <email>)
      var nombreMatch = remitente.match(/^"?([^"<]+)"?\s*</);
      var nombre = nombreMatch ? nombreMatch[1].trim() : remitente;
      var id = _generarIdTicket(hoja);
      hoja.appendRow([
        id, nombre, '', '', asunto + '\n' + cuerpo,
        'Media', 'Nuevo', '', '',
        'Correo', fecha, fecha
      ]);
      hilo.markRead();
      procesados++;
    });
    CacheService.getScriptCache().remove('tickets_data');
    return 'Se crearon ' + procesados + ' ticket(s) desde Gmail.';
  } catch(err) {
    throw new Error('Error al leer Gmail: ' + err.message);
  }
}

// ============================================================
//  ENVIAR NOTIFICACIÓN POR CORREO AL TÉCNICO ASIGNADO
// ============================================================
function notificarTecnico(ticket) {
  try {
    // Mapa técnico → correo (ajustar según tu organización)
    var correosTecnicos = {
      'José Flores':   'jose.flores@hospitalSanJose.pe',
      'Rosa Mamani':   'rosa.mamani@hospitalSanJose.pe',
      'Pedro Quispe':  'pedro.quispe@hospitalSanJose.pe',
      'Luis Cárdenas': 'luis.cardenas@hospitalSanJose.pe'
    };
    var correo = correosTecnicos[ticket.tecnico];
    if (!correo) return;
    GmailApp.sendEmail(
      correo,
      '[TICKET ASIGNADO] ' + ticket.id_ticket + ' — ' + ticket.problema.substring(0, 60),
      'Hola ' + ticket.tecnico + ',\n\n' +
      'Se te ha asignado el siguiente ticket:\n\n' +
      'ID: ' + ticket.id_ticket + '\n' +
      'Usuario: ' + ticket.usuario + '\n' +
      'Área: ' + ticket.area + '\n' +
      'Equipo: ' + ticket.equipo + '\n' +
      'Problema: ' + ticket.problema + '\n' +
      'Prioridad: ' + ticket.prioridad + '\n\n' +
      'Por favor atender a la brevedad.\n\n' +
      '— Sistema de Tickets, Hospital San José'
    );
  } catch(err) {
    Logger.log('Error al notificar técnico: ' + err.message);
  }
}

