/**
 * SISTEMA DE CONTROL INTERNO - SERVICIOS VEHICULARES E INVENTARIO
 * BACKEND MAESTRO CENTRALIZADO (Codigo.gs)
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Main')
      .setTitle('Sistema CI Ciudad Maderas')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function incluir(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function obtenerUsuarioActual() {
  return Session.getActiveUser().getEmail();
}

function FORZAR_PERMISOS_DRIVE() {
  try {
    let folder = DriveApp.createFolder("CI_TEST_PERMISOS");
    folder.setTrashed(true);
  } catch(e) { Logger.log("Permisos validados."); }
}

function generarIdIncremental(nombreHoja, prefijo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  const lastRow = sheet.getLastRow();
  const nextId = lastRow > 0 ? lastRow : 1;
  return prefijo + "-" + nextId.toString().padStart(4, '0');
}

function formatoDDMMYYYY(fechaInput) {
  if (!fechaInput) return "";
  let d = (fechaInput instanceof Date) ? fechaInput : new Date(fechaInput);
  if (isNaN(d.getTime())) return "";
  let dia = ("0" + d.getDate()).slice(-2);
  let mes = ("0" + (d.getMonth() + 1)).slice(-2);
  let anio = d.getFullYear();
  return dia + "/" + mes + "/" + anio;
}

function limpiarHoraLectura(cadena) {
  if (!cadena) return "";
  if (cadena instanceof Date) return formatoDDMMYYYY(cadena);
  let fechaPura = cadena.toString().split(" ")[0].trim();
  if (fechaPura.includes("-")) {
    let p = fechaPura.split("-");
    if (p[0].length === 4) return p[2].padStart(2, '0') + "/" + p[1].padStart(2, '0') + "/" + p[0];
  } else if (fechaPura.includes("/")) {
    let p = fechaPura.split("/");
    if (p[2] && p[2].length === 4) return p[0].padStart(2, '0') + "/" + p[1].padStart(2, '0') + "/" + p[2];
  }
  return fechaPura;
}

function parseFechaToDate(str) {
  if (!str) return new Date(0);
  let s = limpiarHoraLectura(str);
  let parts = s.split('/');
  if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
  if (s.includes('-')) {
    let p = s.split('-');
    if (p[0].length === 4) return new Date(p[0], p[1] - 1, p[2]);
  }
  return new Date(s);
}

function validarLogin(usuario, contrasena) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Accesos");
    const datos = hoja.getDataRange().getDisplayValues();
    for (let i = 1; i < datos.length; i++) {
      let userRow = datos[i][2].toString().trim();
      let passRow = datos[i][3].toString().trim();
      if (userRow === usuario.trim() && passRow === contrasena) {
        return { exito: true, nombre: datos[i][0], rol: datos[i][4], correo: datos[i][5] };
      }
    }
    return { exito: false, error: "Usuario o contraseña incorrectos" };
  } catch(e) { return { exito: false, error: e.message }; }
}

function buscarInfoNuco(nuco) {
  try {
    if (!nuco || nuco.toString().trim() === '') return { exito: false, error: 'NUCO vacío' };
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vehiculos');
    if (!hoja) return { exito: false, error: 'Hoja Vehiculos no encontrada' };
    const datos = hoja.getDataRange().getValues();
    const nucoStr = nuco.toString().trim().toUpperCase();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][4].toString().trim().toUpperCase() === nucoStr) {
        return {
          exito: true,
          datos: {
            serie:        datos[i][2]  || '',
            nuco:         datos[i][4]  || '',
            responsable:  datos[i][5]  || '',
            departamento: datos[i][8]  || '',
            marca:        datos[i][19] || '',
            clase:        datos[i][20] || '',
            linea:        datos[i][21] || '',
            modelo:       datos[i][22] || '',
            color:        datos[i][23] || '',
            placa:        datos[i][24] || '',
            razonSocial:  datos[i][28] || '',
            sede:         datos[i][33] || '',
            oficina:      datos[i][34] || ''
          }
        };
      }
    }
    return { exito: false, error: 'NUCO no encontrado' };
  } catch(e) { return { exito: false, error: e.message }; }
}

function buscarVehiculoGeneral(identificador) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehiculos");
    const datos = hoja.getDataRange().getDisplayValues();
    const buscado = identificador.toString().toUpperCase().trim();
    for (let i = 1; i < datos.length; i++) {
      let nucoE  = datos[i][4].toString().toUpperCase().trim();
      let placaY = datos[i][24].toString().toUpperCase().trim();
      if (nucoE === buscado || placaY === buscado) {
        return { placa: datos[i][24], nuco: datos[i][4], serie: datos[i][2], marca: datos[i][19], modelo: datos[i][22] };
      }
    }
    return { error: "No encontrado" };
  } catch(e) { return { error: e.message }; }
}

function buscarInsumosPorTicket(ticketBuscado) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario");
    const datos = hoja.getDataRange().getDisplayValues();
    const buscado = ticketBuscado.toString().toUpperCase().trim();
    let encontrados = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      let ticket  = datos[i][1].toString().toUpperCase().trim();
      let estatus = datos[i][12].toString().toUpperCase().trim();
      if (ticket === buscado && estatus.indexOf("DISPONIBLE") > -1) {
        let cant = parseFloat(datos[i][9]) || 0;
        let cu   = parseFloat(datos[i][11].replace(/[^0-9.-]+/g,"")) || parseFloat(datos[i][11]) || 0;
        if (cant > 0) {
          encontrados.push({
            id:            datos[i][0].toString().toUpperCase().trim(),
            producto:      datos[i][4] || "S/N",
            vehiculo:      datos[i][6] || "S/V",
            cantidad:      cant,
            costoUnitario: cu
          });
        }
      }
    }
    return encontrados.length > 0
      ? { exito: true, datos: encontrados }
      : { exito: false, datos: [], error: "No se encontraron refacciones DISPONIBLES con ese ticket." };
  } catch(e) { return { exito: false, datos: [], error: "Fallo técnico: " + e.message }; }
}

function registrarTaller(d) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Taller");
    let prefijo = d.proveedor === "INTERNO" ? "T-INT" : (d.proveedor === "EXTERNO" ? "T-EXT" : "T-AMB");
    const idGenerado = generarIdIncremental("Taller", prefijo);
    hoja.appendRow([
      idGenerado, d.nuco, d.ticket, d.folioCpp, d.alertaTicket, d.serie, d.odometro,
      formatoDDMMYYYY(new Date()), limpiarHoraLectura(d.fechaServicio), d.tipoServicio,
      d.familia, d.numeroPartes, d.proveedor, d.descripcion, d.quienRegistra,
      d.servInterno, d.servExterno, d.costoTotal, d.estatusUnidad, d.evidencia
    ]);
    SpreadsheetApp.flush();
    return "Registro guardado con ID: " + idGenerado;
  } catch(e) { return "Error: " + e.message; }
}

function registrarInventario(g, arr) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario");
    const f = formatoDDMMYYYY(new Date());
    arr.forEach(item => {
      let idG  = generarIdIncremental("Inventario", "INV");
      let fCpp = g.folioCpp ? g.folioCpp : "PENDIENTE";
      hoja.appendRow([ idG, g.ticket, fCpp, g.obsTicket, item.producto, item.marca, item.vehiculo, item.unidadMedida, f, item.cantidad, item.costoTotal, item.costoUnitario, "DISPONIBLE", g.quienRegistra ]);
    });
    SpreadsheetApp.flush();
    return "Registrados " + arr.length + " artículo(s) con éxito.";
  } catch(e) { return "Error: " + e.message; }
}

function registrarConsumoMultiple(g, arr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hc = ss.getSheetByName("Consumos");
    const hi = ss.getSheetByName("Inventario");
    const dI = hi.getDataRange().getValues();
    let mapI = {};
    for (let i = 1; i < dI.length; i++) {
      if (dI[i][0]) mapI[dI[i][0].toString()] = {
        fila: i + 1,
        c:    parseFloat(dI[i][9])  || 0,
        cu:   parseFloat(dI[i][11]) || 0,
        p:    dI[i][4].toString(),
        v:    dI[i][6].toString()
      };
    }
    arr.forEach(d => {
      let it = mapI[d.idPieza.toString()];
      if (it) {
        let cons = parseFloat(d.cantidad) || 0;
        if (cons > it.c) cons = it.c;
        if (cons > 0) {
          let nC = it.c - cons;
          it.c = nC;
          hc.appendRow([ generarIdIncremental("Consumos", "CONS"), d.idPieza, it.p, g.quienRegistra, g.ticketInsumo, it.v, limpiarHoraLectura(g.fechaConsumo), cons, (cons * it.cu), d.evidencia ]);
          hi.getRange(it.fila, 10).setValue(nC);
          hi.getRange(it.fila, 11).setValue(nC * it.cu);
          if (nC <= 0) hi.getRange(it.fila, 13).setValue("CONSUMIDO");
        }
      }
    });
    SpreadsheetApp.flush();
    return "Consumos procesados.";
  } catch(e) { return "Error: " + e.message; }
}

function registrarProveedor(d) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proveedores");
    const idGenerado = generarIdIncremental("Proveedores", "PROV");
    hoja.appendRow([ idGenerado, d.proveedor, d.razonSocial, d.rfc, d.direccion, d.estado, d.contacto, d.numero, d.correo, d.segmentacion, d.diasCredito, d.totalCredito, d.servicio, d.razonesSociales ]);
    SpreadsheetApp.flush();
    return { msj: "Proveedor registrado con éxito." };
  } catch(e) { return { msj: "Error: " + e.message }; }
}

function eliminarFilaTaller(id) { return _eliminarRegistro('Taller', id); }
function eliminarFilaInventario(id) { return _eliminarRegistro('Inventario', id); }
function eliminarFilaConsumo(id) { return _eliminarRegistro('Consumos', id); }
function eliminarProveedor(id) { return _eliminarRegistro('Proveedores', id); }

function _eliminarRegistro(tipo, id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja;
    if (tipo === 'Taller')      hoja = ss.getSheetByName('Taller');
    if (tipo === 'Inventario')  hoja = ss.getSheetByName('Inventario');
    if (tipo === 'Consumos')    hoja = ss.getSheetByName('Consumos');
    if (tipo === 'Proveedores') hoja = ss.getSheetByName('Proveedores');
    if (!hoja) return { exito: false, msj: "Hoja no encontrada." };
    const datos = hoja.getDataRange().getValues();
    let fila = -1;
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === id.toString().trim()) { fila = i + 1; break; }
    }
    if (fila === -1) return { exito: false, msj: "Registro no encontrado." };
    if (tipo === 'Consumos') {
      let idP = datos[fila-1][1].toString();
      let cR  = parseFloat(datos[fila-1][7]) || 0;
      const hI = ss.getSheetByName('Inventario');
      const dI = hI.getDataRange().getValues();
      for (let j = 1; j < dI.length; j++) {
        if (dI[j][0].toString() === idP) {
          let nS = (parseFloat(dI[j][9]) || 0) + cR;
          hI.getRange(j+1, 10).setValue(nS);
          hI.getRange(j+1, 11).setValue(nS * (parseFloat(dI[j][11]) || 0));
          hI.getRange(j+1, 13).setValue("DISPONIBLE");
          break;
        }
      }
    } else if (tipo === 'Inventario') {
      const hC = ss.getSheetByName('Consumos');
      const dC = hC.getDataRange().getValues();
      for (let k = dC.length - 1; k >= 1; k--) {
        if (dC[k][1].toString().trim() === id.toString().trim()) hC.deleteRow(k + 1);
      }
    }
    hoja.deleteRow(fila);
    SpreadsheetApp.flush();
    return { exito: true, msj: "Registro eliminado exitosamente." };
  } catch(e) { return { exito: false, msj: "Fallo al borrar registro: " + e.message }; }
}

function registrarOC(datosGenerales, todasLasCotizaciones, configuracionCorreo, correoDestino) {
  try {
    const ss             = SpreadsheetApp.getActiveSpreadsheet();
    const hojaOC         = ss.getSheetByName("OC");
    const hojaInventario = ss.getSheetByName("Inventario");
    const folioGenerado  = generarIdIncremental("OC", "SV");
    const fRegistro      = formatoDDMMYYYY(new Date());
    let partidasGanadoras = todasLasCotizaciones.filter(c => c.esGanador === true);
    let fCpp = datosGenerales.folioCpp ? datosGenerales.folioCpp : "PENDIENTE";
    for (let i = 0; i < partidasGanadoras.length; i++) {
      let prod = partidasGanadoras[i];
      let idGenerado = generarIdIncremental("OC", "OC");
      hojaOC.appendRow([
        idGenerado, "SV", folioGenerado, datosGenerales.ticket, fRegistro, datosGenerales.usuario,
        datosGenerales.nuco, datosGenerales.placas, datosGenerales.linea, datosGenerales.modelo,
        datosGenerales.departamento, datosGenerales.sede, datosGenerales.oficina, prod.nombreProveedor,
        prod.rfc, prod.nombreProducto, prod.familia, prod.descripcion, prod.cantidad, prod.unidadMedida,
        prod.tiempoEntrega, prod.precioUnitario, prod.subtotal, prod.intercambio, prod.costoIntercambio,
        prod.total, datosGenerales.razonSocialCompra, datosGenerales.odometroActual,
        datosGenerales.odometroUltimo, datosGenerales.formaPago, datosGenerales.comentarios,
        datosGenerales.motivoCompra, datosGenerales.metodoPagoF, datosGenerales.usoCFDI,
        datosGenerales.puntosConsiderar, datosGenerales.tipoServicioPartida, datosGenerales.tipoSolicitud,
        datosGenerales.tipoServicio, datosGenerales.tiempoVida, datosGenerales.servicioIntExt
      ]);
      let idInv = generarIdIncremental("Inventario", "INV");
      hojaInventario.appendRow([ idInv, datosGenerales.ticket, fCpp, "ALTA AUTOMÁTICA DESDE OC: " + folioGenerado, prod.nombreProducto, prod.marca, datosGenerales.nuco, prod.unidadMedida, fRegistro, prod.cantidad, prod.total, prod.precioUnitario, "DISPONIBLE", datosGenerales.usuario ]);
    }
    SpreadsheetApp.flush();
    let resultadoArchivos = ejecutarCompilacionFormatosPDF(datosGenerales.ticket, folioGenerado, todasLasCotizaciones, configuracionCorreo, datosGenerales, correoDestino);
    return { exito: true, msj: "¡Éxito! Documentos PDF generados.", archivos: resultadoArchivos };
  } catch(e) { return { exito: false, error: "Error al procesar: " + e.message }; }
}

function ejecutarCompilacionFormatosPDF(ticket, folioOC, todasLasCotizaciones, configCorreo, datosGenerales, correoUsuarioDestino) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const datosProv = ss.getSheetByName("Proveedores").getDataRange().getDisplayValues();
  ss.getSheets().forEach(s => {
    if (s.getName().startsWith("TEMP_OC_") || s.getName().startsWith("TEMP_COMP_")) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });
  let masterFolder   = DriveApp.getFoldersByName("COMPARATIVAS Y OC - SV").hasNext() ? DriveApp.getFoldersByName("COMPARATIVAS Y OC - SV").next() : DriveApp.createFolder("COMPARATIVAS Y OC - SV");
  let folderPDF      = masterFolder.getFoldersByName("TICKET_" + ticket).hasNext() ? masterFolder.getFoldersByName("TICKET_" + ticket).next() : masterFolder.createFolder("TICKET_" + ticket);
  let resultadoLinks = { comparativa: "", ocs: [] };
  let filesToEmail   = [];
  let tsUnique       = new Date().getTime();
  const hojasBase    = ["Accesos","Vehiculos","Taller","Inventario","Consumos","Proveedores","OC","COMPARATIVA DOC","OC DOC","FORMATO ALTA VEHICULOS"];
  let rzCompra       = datosGenerales.razonSocialCompra ? datosGenerales.razonSocialCompra.toUpperCase() : "";
  let textoHomoclave = datosGenerales.departamento.toUpperCase().includes("OOAM") ? "HOMOCLAVE OOAM DE CONFORMIDAD CON EL TÍTULO DE CONCESIÓN OTORGADO POR LAS AUTORIDADES CORRESPONDIENTES A: " + rzCompra : "";
  let partidasGanadoras = todasLasCotizaciones.filter(c => c.esGanador === true);
  let aplicoIntercambio = partidasGanadoras.some(p => p.intercambio === "SI") ? "SÍ" : "NO";
  let tipoServicioF  = datosGenerales.servicioIntExt ? datosGenerales.servicioIntExt.toUpperCase().trim() : "";
  let marcaInterno   = (tipoServicioF === "INTERNO" || tipoServicioF === "AMBOS" || tipoServicioF === "AMBAS") ? "X" : "";
  let marcaExterno   = (tipoServicioF === "EXTERNO" || tipoServicioF === "AMBOS" || tipoServicioF === "AMBAS") ? "X" : "";
  let productosUnicos = [];
  todasLasCotizaciones.forEach(p => {
    if (!productosUnicos.find(u => u.nombre === p.nombreProducto)) {
      productosUnicos.push({ nombre: p.nombreProducto, familia: p.familia, cant: p.cantidad, um: p.unidadMedida });
    }
  });
  const hojaCompOriginal = ss.getSheetByName("COMPARATIVA DOC");
  let hojaComp = hojaCompOriginal.copyTo(ss).setName("TEMP_COMP_" + folioOC + "_" + tsUnique);
  hojaComp.showSheet();
  hojaComp.setHiddenGridlines(true);
  hojaComp.createTextFinder("{{TEXTO_HOMOCLAVE}}").replaceAllWith(textoHomoclave);
  hojaComp.createTextFinder("{{FECHA}}").replaceAllWith(formatoDDMMYYYY(new Date()));
  hojaComp.createTextFinder("{{TICKET}}").replaceAllWith(datosGenerales.ticket);
  hojaComp.createTextFinder("{{NUCO}}").replaceAllWith(datosGenerales.nuco);
  hojaComp.createTextFinder("{{DEPARTAMENTO}}").replaceAllWith(datosGenerales.departamento);
  hojaComp.createTextFinder("{{SEDE}}").replaceAllWith(datosGenerales.sede);
  hojaComp.createTextFinder("{{OFICINA}}").replaceAllWith(datosGenerales.oficina);
  hojaComp.createTextFinder("{{LINEA}}").replaceAllWith(datosGenerales.linea);
  hojaComp.createTextFinder("{{MODELO}}").replaceAllWith(datosGenerales.modelo);
  hojaComp.createTextFinder("{{PLACAS}}").replaceAllWith(datosGenerales.placas);
  hojaComp.createTextFinder("{{RAZON SOCIAL COMPRA}}").replaceAllWith(datosGenerales.razonSocialCompra);
  hojaComp.createTextFinder("{{ODOMETRO ACTUAL}}").replaceAllWith(datosGenerales.odometroActual);
  hojaComp.createTextFinder("{{ODOMETRO DEL ULTIMO SERVICIO REALIZADO}}").replaceAllWith(datosGenerales.odometroUltimo);
  hojaComp.createTextFinder("{{CHECK_INTERNO}}").replaceAllWith(marcaInterno);
  hojaComp.createTextFinder("{{CHECK_EXTERNO}}").replaceAllWith(marcaExterno);
  hojaComp.createTextFinder("{{COMENTARIOS}}").replaceAllWith(datosGenerales.comentarios);
  hojaComp.createTextFinder("{{MOTIVO DE COMPRA}}").replaceAllWith(datosGenerales.motivoCompra);
  hojaComp.createTextFinder("{{TIPO SERVICIO PARTIDA}}").replaceAllWith(datosGenerales.tipoServicioPartida);
  hojaComp.createTextFinder("{{TIEMPO VIDA ESTIMADO}}").replaceAllWith(datosGenerales.tiempoVida);
  hojaComp.createTextFinder("{{TIPO SOLICITUD}}").replaceAllWith(datosGenerales.tipoSolicitud);
  hojaComp.createTextFinder("{{TIPO SERVICIO}}").replaceAllWith(datosGenerales.tipoServicio);
  hojaComp.createTextFinder("{{QUIEN REGISTRA}}").replaceAllWith(datosGenerales.usuario || "");
  let compGranTotal = 0; let compAhorroIntercambio = 0; let totalProveedorContado = 0;
  partidasGanadoras.forEach(p => {
    let sub = parseFloat(p.subtotal) || 0; let tot = parseFloat(p.total) || 0;
    compGranTotal += tot;
    if (p.intercambio === "SI") compAhorroIntercambio += (sub * 0.30);
    totalProveedorContado += sub;
  });
  let difDinero     = compGranTotal - totalProveedorContado;
  let difPorcentaje = totalProveedorContado > 0 ? (difDinero / totalProveedorContado) : 0;
  let mensajeEconomia = difPorcentaje < 0 ? "MÁS ECONÓMICO A CRÉDITO" : (difPorcentaje <= 0.15 ? "SE ENCUENTRA DENTRO DEL 15% AUTORIZADO" : "SOBREPASA EL 15% AUTORIZADO");
  let colItem_C = [], colNombre_C = [], colFamilia_C = [], colCant_C = [], colUM_C = [];
  productosUnicos.forEach((u, i) => { colItem_C.push(i+1); colNombre_C.push(u.nombre); colFamilia_C.push(u.familia); colCant_C.push(u.cant); colUM_C.push(u.um); });
  hojaComp.createTextFinder("{{ITEM}}").replaceAllWith(colItem_C.join('\n'));
  hojaComp.createTextFinder("{{NOMBRE}}").replaceAllWith(colNombre_C.join('\n'));
  hojaComp.createTextFinder("{{FAMILIA}}").replaceAllWith(colFamilia_C.join('\n'));
  hojaComp.createTextFinder("{{CANTIDAD PRODUCTOS}}").replaceAllWith(colCant_C.join('\n'));
  hojaComp.createTextFinder("{{UNIDAD MEDIDA}}").replaceAllWith(colUM_C.join('\n'));
  let provKeys = [...new Set(todasLasCotizaciones.map(p => p.rfc))];
  for (let m = 0; m < 5; m++) {
    let pIdx = m + 1;
    if (m < provKeys.length) {
      let rfcActual = provKeys[m];
      let nombreP   = todasLasCotizaciones.find(p => p.rfc === rfcActual).nombreProveedor.toUpperCase();
      let provData  = datosProv.find(d => d[3].toString().toUpperCase().trim() === rfcActual) || [];
      let rzP       = provData[2] ? provData[2].toString().toUpperCase() : "";
      let colDesc_P = [], colSub_P = []; let tot_P = 0; let fp_P = datosGenerales.formaPago; let te_P = "";
      productosUnicos.forEach(u => {
        let pItem = todasLasCotizaciones.find(p => p.rfc === rfcActual && p.nombreProducto === u.nombre);
        if (pItem) { colDesc_P.push(pItem.descripcion); colSub_P.push("$" + parseFloat(pItem.subtotal).toLocaleString('es-MX',{minimumFractionDigits:2})); te_P = pItem.tiempoEntrega; tot_P += parseFloat(pItem.total); }
        else { colDesc_P.push("-"); colSub_P.push("-"); }
      });
      hojaComp.createTextFinder("{{PROVEEDOR_P"+pIdx+"}}").replaceAllWith(nombreP);
      hojaComp.createTextFinder("{{RAZON SOCIAL_P"+pIdx+"}}").replaceAllWith(rzP);
      hojaComp.createTextFinder("{{FORMA DE PAGO_P"+pIdx+"}}").replaceAllWith(fp_P);
      hojaComp.createTextFinder("{{TIEMPO ENTREGA_P"+pIdx+"}}").replaceAllWith(te_P);
      hojaComp.createTextFinder("{{DESCRIPCION_P"+pIdx+"}}").replaceAllWith(colDesc_P.join('\n'));
      hojaComp.createTextFinder("{{SUBTOTAL_P"+pIdx+"}}").replaceAllWith(colSub_P.join('\n'));
      hojaComp.createTextFinder("{{TOTAL_P"+pIdx+"}}").replaceAllWith("$" + tot_P.toLocaleString('es-MX',{minimumFractionDigits:2}));
    } else {
      ["PROVEEDOR_P","RAZON SOCIAL_P","FORMA DE PAGO_P","TIEMPO ENTREGA_P","DESCRIPCION_P","SUBTOTAL_P","TOTAL_P"].forEach(k => hojaComp.createTextFinder("{{"+k+pIdx+"}}").replaceAllWith("-"));
    }
  }
  hojaComp.createTextFinder("{{GRAN_TOTAL_COMPRA}}").replaceAllWith("$" + compGranTotal.toLocaleString('es-MX',{minimumFractionDigits:2}));
  hojaComp.createTextFinder("{{AHORRO_INTERCAMBIO}}").replaceAllWith(compAhorroIntercambio > 0 ? "$" + compAhorroIntercambio.toLocaleString('es-MX',{minimumFractionDigits:2}) : "$0.00");
  hojaComp.createTextFinder("{{TOTAL_CON_DESCUENTO}}").replaceAllWith("$" + (compGranTotal - compAhorroIntercambio).toLocaleString('es-MX',{minimumFractionDigits:2}));
  hojaComp.createTextFinder("{{DIF_DINERO}}").replaceAllWith((difDinero >= 0 ? "$" : "-$") + Math.abs(difDinero).toLocaleString('es-MX',{minimumFractionDigits:2}));
  hojaComp.createTextFinder("{{DIF_PORCENTAJE}}").replaceAllWith((difPorcentaje * 100).toFixed(0) + "%");
  hojaComp.createTextFinder("{{MENSAJE_ECONOMIA}}").replaceAllWith(mensajeEconomia);
  hojaComp.createTextFinder("{{¿INTERCAMBIO?}}").replaceAllWith(aplicoIntercambio);
  ss.setActiveSheet(hojaComp);
  ss.moveActiveSheet(1);
  ss.getSheets().forEach(h => { if (h.getName() !== hojaComp.getName()) h.hideSheet(); });
  SpreadsheetApp.flush();
  let fileComp = folderPDF.createFile(ss.getAs('application/pdf').setName("COMPARATIVA_TICKET_" + ticket + "_FOLIO_" + folioOC + ".pdf"));
  try { fileComp.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  resultadoLinks.comparativa = fileComp.getUrl();
  filesToEmail.push(fileComp);
  ss.getSheets().forEach(h => { if (hojasBase.includes(h.getName())) h.showSheet(); });
  ss.deleteSheet(hojaComp);
  const hojaOCOriginal   = ss.getSheetByName("OC DOC");
  let rfcUnicosGanadores = [...new Set(partidasGanadoras.map(p => p.rfc))];
  for (let k = 0; k < rfcUnicosGanadores.length; k++) {
    let rfc        = rfcUnicosGanadores[k];
    let pGanadores = partidasGanadoras.filter(p => p.rfc === rfc);
    let nombreProv = pGanadores[0].nombreProveedor.toUpperCase();
    let tempOC     = hojaOCOriginal.copyTo(ss).setName("TEMP_OC_" + rfc + "_" + tsUnique);
    tempOC.showSheet();
    tempOC.setHiddenGridlines(true);
    let pD      = datosProv.find(d => d[3].toString().toUpperCase().trim() === rfc) || [];
    let pRz     = pD[2]  ? pD[2].toString().toUpperCase()  : "";
    let pDir    = pD[4]  ? pD[4].toString().toUpperCase()  : "";
    let pCont   = pD[6]  ? pD[6].toString().toUpperCase() + " / " + pD[9].toString().toLowerCase() : "";
    let pCorreo = pD[9]  ? pD[9].toString().toLowerCase()  : "";
    let pCred   = pD[10] ? pD[10].toString()               : "";
    let lLeg    = "Mediante la aceptación vía correo electrónico de la presente orden de compra, el proveedor " + nombreProv + " asume de manera exclusiva y total la responsabilidad sobre la calidad, el estado y la integridad del producto hasta el momento en que se efectúe la entrega física y se firme por escrito la misma y conforme en el lugar estipulado en este documento.\n\nEn el supuesto de que la Orden de Compra sea cancelada, el proveedor " + nombreProv + " será responsable de cubrir todos los costos y gastos que se deriven de dicha cancelación en un plazo máximo de 5 días hábiles.";
    let pGTotal = 0; let pAhorro = 0;
    let cItm=[],cNom=[],cFam=[],cCant=[],cUM=[],cDes=[],cPU=[],cInt=[],cSub=[],cCInt=[],cTot=[];
    pGanadores.forEach((p, idx) => {
      let sub=parseFloat(p.subtotal)||0; let tot=parseFloat(p.total)||0; let ci=parseFloat(p.costoIntercambio)||0;
      pGTotal+=tot; pAhorro+=ci;
      cItm.push(idx+1); cNom.push(p.nombreProducto); cFam.push(p.familia);
      cCant.push(p.cantidad); cUM.push(p.unidadMedida); cDes.push(p.descripcion);
      cPU.push("$"+parseFloat(p.precioUnitario).toLocaleString('es-MX'));
      cInt.push(p.intercambio); cSub.push("$"+sub.toLocaleString('es-MX'));
      cCInt.push("$"+ci.toLocaleString('es-MX')); cTot.push("$"+tot.toLocaleString('es-MX'));
    });
    tempOC.createTextFinder("{{FOLIO}}").replaceAllWith(folioOC);
    tempOC.createTextFinder("{{FECHA}}").replaceAllWith(formatoDDMMYYYY(new Date()));
    tempOC.createTextFinder("{{TICKET}}").replaceAllWith(datosGenerales.ticket);
    tempOC.createTextFinder("{{DEPARTAMENTO}}").replaceAllWith(datosGenerales.departamento);
    tempOC.createTextFinder("{{SEDE}}").replaceAllWith(datosGenerales.sede);
    tempOC.createTextFinder("{{OFICINA}}").replaceAllWith(datosGenerales.oficina);
    tempOC.createTextFinder("{{LINEA}}").replaceAllWith(datosGenerales.linea);
    tempOC.createTextFinder("{{MODELO}}").replaceAllWith(datosGenerales.modelo);
    tempOC.createTextFinder("{{ODOMETRO ACTUAL}}").replaceAllWith(datosGenerales.odometroActual);
    tempOC.createTextFinder("{{PLACAS}}").replaceAllWith(datosGenerales.placas);
    tempOC.createTextFinder("{{NUCO}}").replaceAllWith(datosGenerales.nuco);
    tempOC.createTextFinder("{{RAZON SOCIAL}}").replaceAllWith(pRz);
    tempOC.createTextFinder("{{DIRECCION}}").replaceAllWith(pDir);
    tempOC.createTextFinder("{{NUMERO DE CONTACTO}}").replaceAllWith(pCont);
    tempOC.createTextFinder("{{DIAS DE CREDITO}}").replaceAllWith(pCred);
    tempOC.createTextFinder("{{RFC}}").replaceAllWith(rfc);
    tempOC.createTextFinder("{{CORREO}}").replaceAllWith(pCorreo);
    tempOC.createTextFinder("{{RAZON SOCIAL COMPRA}}").replaceAllWith(datosGenerales.razonSocialCompra);
    tempOC.createTextFinder("{{USO CFDI}}").replaceAllWith(datosGenerales.usoCFDI);
    tempOC.createTextFinder("{{METODO PAGO FACT}}").replaceAllWith(datosGenerales.metodoPagoF);
    tempOC.createTextFinder("{{FORMA DE PAGO}}").replaceAllWith(datosGenerales.formaPago);
    tempOC.createTextFinder("{{PUNTOS A CONSIDERAR}}").replaceAllWith(datosGenerales.puntosConsiderar);
    tempOC.createTextFinder("{{QUIEN REGISTRA}}").replaceAllWith(datosGenerales.usuario || "");
    tempOC.createTextFinder("{{ITEM}}").replaceAllWith(cItm.join('\n'));
    tempOC.createTextFinder("{{NOMBRE}}").replaceAllWith(cNom.join('\n'));
    tempOC.createTextFinder("{{FAMILIA}}").replaceAllWith(cFam.join('\n'));
    tempOC.createTextFinder("{{CANTIDAD}}").replaceAllWith(cCant.join('\n'));
    tempOC.createTextFinder("{{UNIDAD MEDIDA}}").replaceAllWith(cUM.join('\n'));
    tempOC.createTextFinder("{{DESCRIPCION}}").replaceAllWith(cDes.join('\n'));
    tempOC.createTextFinder("{{PRECIO UNITARIO}}").replaceAllWith(cPU.join('\n'));
    tempOC.createTextFinder("{{¿INTERCAMBIO?}}").replaceAllWith(cInt.join('\n'));
    tempOC.createTextFinder("{{SUBTOTAL}}").replaceAllWith(cSub.join('\n'));
    tempOC.createTextFinder("{{SUBTOTAL_OC}}").replaceAllWith("$"+(pGTotal-pAhorro).toLocaleString('es-MX',{minimumFractionDigits:2}));
    tempOC.createTextFinder("{{INTERCAMBIO_OC}}").replaceAllWith("$"+pAhorro.toLocaleString('es-MX',{minimumFractionDigits:2}));
    tempOC.createTextFinder("{{TOTAL_OC}}").replaceAllWith("$"+pGTotal.toLocaleString('es-MX',{minimumFractionDigits:2}));
    tempOC.createTextFinder("{{PROVEEDOR_GANADOR}}").replaceAllWith(nombreProv);
    tempOC.createTextFinder("{{LEYENDA_LEGAL_OC}}").replaceAllWith(lLeg);
    tempOC.createTextFinder("{{TEXTO_HOMOCLAVE}}").replaceAllWith(textoHomoclave);
    tempOC.createTextFinder("{{CHECK_INTERNO}}").replaceAllWith(marcaInterno);
    tempOC.createTextFinder("{{CHECK_EXTERNO}}").replaceAllWith(marcaExterno);
    ss.getSheets().forEach(h => { if (h.getName() !== tempOC.getName()) h.hideSheet(); });
    SpreadsheetApp.flush();
    let fileOC = folderPDF.createFile(ss.getAs('application/pdf').setName("OC_" + nombreProv.replace(/ /g,"_") + "_TICKET_" + ticket + "_FOLIO_" + folioOC + ".pdf"));
    try { fileOC.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    resultadoLinks.ocs.push({ proveedor: nombreProv, url: fileOC.getUrl() });
    filesToEmail.push(fileOC);
    ss.getSheets().forEach(h => { if (hojasBase.includes(h.getName())) h.showSheet(); });
    ss.deleteSheet(tempOC);
  }
  ss.getSheets().forEach(h => { if (hojasBase.includes(h.getName())) h.showSheet(); });
  if (configCorreo !== "NINGUNO" && correoUsuarioDestino) {
    let adjuntos = configCorreo === "AMBOS" ? filesToEmail : filesToEmail.filter(f => f.getName().indexOf("OC_") > -1);
    MailApp.sendEmail({
      to: correoUsuarioDestino,
      subject: "DOCUMENTOS GENERADOS - TICKET: " + ticket + " | FOLIO OC: " + folioOC,
      htmlBody: "<h3>SISTEMA CI - SERVICIOS VEHICULARES</h3><p>Se adjuntan los documentos para el Ticket: " + ticket + " / Folio: " + folioOC + "</p>",
      name: "Sistema de Control Interno",
      attachments: adjuntos
    });
  }
  return resultadoLinks;
}

function obtenerHistorialGlobalOC() {
  try {
    const d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OC").getDataRange().getDisplayValues();
    let filas = []; let foliosSet = new Set();
    for (let i = d.length - 1; i >= 1; i--) {
      if (!d[i][2]) continue;
      let folio = d[i][2];
      if (!foliosSet.has(folio)) {
        foliosSet.add(folio);
        filas.push({ folio: folio, ticket: d[i][3], fecha: limpiarHoraLectura(d[i][4]), nuco: d[i][6], razonSocial: d[i][25] });
      }
    }
    return { exito: true, datos: filas };
  } catch(e) { return { exito: false, error: "Fallo al leer la bitácora." }; }
}

function buscarDocumentosGenerados(ticket, folio) {
  try {
    let masterFolder = DriveApp.getFoldersByName("COMPARATIVAS Y OC - SV");
    if (!masterFolder.hasNext()) return { exito: false, msj: "No existe la carpeta de documentos." };
    let ticketFolder = masterFolder.next().getFoldersByName("TICKET_" + ticket);
    if (!ticketFolder.hasNext()) return { exito: false, msj: "No se encontraron documentos para este ticket." };
    let files = ticketFolder.next().getFiles();
    let res = [];
    while (files.hasNext()) { let f = files.next(); res.push({ nombre: f.getName(), url: f.getUrl() }); }
    return res.length > 0 ? { exito: true, archivos: res } : { exito: false, msj: "La carpeta está vacía." };
  } catch(e) { return { exito: false, error: "Problema conectando con Drive." }; }
}

function obtenerDatosInventario() {
  try {
    const d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario").getDataRange().getDisplayValues();
    let tot = 0; let filas = [];
    for (let i = 1; i < d.length; i++) {
      if (!d[i][1].toString().trim()) continue;
      let est = d[i][12].toString().toUpperCase().trim();
      let c   = parseFloat(d[i][9]) || 0;
      let cu  = parseFloat(d[i][11].replace(/[^0-9.-]+/g,"")) || 0;
      let ct  = parseFloat(d[i][10].replace(/[^0-9.-]+/g,"")) || (c * cu);
      if (est.indexOf("DISPONIBLE") > -1) tot += ct;
      filas.push({ id: d[i][0], ticket: d[i][1], folioCpp: d[i][2], obs: d[i][3], producto: d[i][4], marca: d[i][5], vehiculo: d[i][6], cantidad: c, medida: d[i][7], estatus: est, costoTotal: ct, costoUnitario: cu, fecha: limpiarHoraLectura(d[i][8]), quienRegistra: d[i][13] });
    }
    return { exito: true, datos: filas.reverse(), total: tot };
  } catch(e) { return { exito: false, error: "Error de lectura de inventario." }; }
}

function obtenerDatosTaller() {
  try {
    const d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Taller").getDataRange().getDisplayValues();
    let tI = 0; let tE = 0; let filas = [];
    for (let i = 1; i < d.length; i++) {
      if (!d[i][1].toString().trim()) continue;
      let cInt = parseFloat(d[i][15].replace(/[^0-9.-]+/g,"")) || 0;
      let cExt = parseFloat(d[i][16].replace(/[^0-9.-]+/g,"")) || 0;
      tI += cInt; tE += cExt;
      filas.push({
        idTaller: d[i][0], nuco: d[i][1], ticket: d[i][2], folioCpp: d[i][3],
        alertaTicket: d[i][4], serie: d[i][5], odometro: d[i][6],
        fecha: limpiarHoraLectura(d[i][7]), fechaServicio: limpiarHoraLectura(d[i][8]),
        tipoServicio: d[i][9], familia: d[i][10], numeroPartes: d[i][11],
        proveedor: d[i][12], descripcion: d[i][13], quienRegistra: d[i][14],
        costoTotal: d[i][17], estatus: d[i][18], servInterno: cInt, servExterno: cExt
      });
    }
    return { exito: true, datos: filas.reverse(), totalInterno: tI, totalExterno: tE };
  } catch(e) { return { exito: false, error: "Error de lectura de Taller." }; }
}

function obtenerDatosConsumos() {
  try {
    const d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Consumos").getDataRange().getDisplayValues();
    let tot = 0; let filas = [];
    for (let i = 1; i < d.length; i++) {
      if (!d[i][1].toString().trim()) continue;
      let ct = parseFloat(d[i][8].replace(/[^0-9.-]+/g,"")) || 0;
      tot += ct;
      filas.push({ idConsumo: d[i][0], producto: d[i][2], quienRegistra: d[i][3], ticketNuevo: d[i][4], vehiculo: d[i][5], fecha: limpiarHoraLectura(d[i][6]), cantidad: d[i][7], costoTotal: d[i][8] });
    }
    return { exito: true, datos: filas.reverse(), total: tot };
  } catch(e) { return { exito: false, error: "Error de lectura de consumos." }; }
}

function obtenerDatosProveedores() {
  try {
    const d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proveedores").getDataRange().getDisplayValues();
    let f = [];
    for (let i = 1; i < d.length; i++) {
      if (d[i][1].toString().trim()) {
        f.push({ id: d[i][0], proveedor: d[i][1], razonSocial: d[i][2], rfc: d[i][3], direccion: d[i][4], estado: d[i][5], contacto: d[i][6], telefono: d[i][7], correo: d[i][8], segmentacion: d[i][9], diasCredito: d[i][10], totalCredito: d[i][11], servicio: d[i][12], razonesSociales: d[i][13] });
      }
    }
    return { exito: true, datos: f.reverse() };
  } catch(e) { return { exito: false, error: "Error al leer proveedores." }; }
}

function obtenerHistorialNuco(nucoBusqueda) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const b  = nucoBusqueda.toString().toUpperCase().trim();
    let h = [];
    const dT = ss.getSheetByName("Taller").getDataRange().getDisplayValues();
    for (let i = 1; i < dT.length; i++) {
      if (dT[i][1] && dT[i][1].toString().toUpperCase().trim() === b) {
        let fS = limpiarHoraLectura(dT[i][8]);
        h.push({ modulo: 'TALLER', timestamp: parseFechaToDate(fS).getTime(), fecha: fS, ticket: dT[i][2], odometro: dT[i][6], tipoServicio: dT[i][9], proveedor: dT[i][12], descripcion: dT[i][13], costoTotal: parseFloat(dT[i][17].replace(/[^0-9.-]+/g,"")) || 0, quienRegistra: dT[i][14] });
      }
    }
    const dC = ss.getSheetByName("Consumos").getDataRange().getDisplayValues();
    for (let i = 1; i < dC.length; i++) {
      if (dC[i][5] && dC[i][5].toString().toUpperCase().trim().includes(b)) {
        let fC = limpiarHoraLectura(dC[i][6]);
        h.push({ modulo: 'CONSUMO', timestamp: parseFechaToDate(fC).getTime(), fecha: fC, ticket: dC[i][4] || dC[i][3], descripcion: dC[i][2], costoTotal: parseFloat(dC[i][8].replace(/[^0-9.-]+/g,"")) || 0, quienRegistra: dC[i][3] });
      }
    }
    return { exito: true, datos: h.sort((x, y) => y.timestamp - x.timestamp) };
  } catch(e) { return { exito: false, error: "Error de lectura de historial." }; }
}

function obtenerDetallesCompletosTicket(ticketBusqueda) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const b  = ticketBusqueda.toString().toUpperCase().trim();
    let dTal = [], dInv = [], dCon = [];
    const dT = ss.getSheetByName("Taller").getDataRange().getDisplayValues();
    for (let i = 1; i < dT.length; i++) {
      if (dT[i][2].toString().toUpperCase().trim() === b) {
        dTal.push({ id: dT[i][0], nuco: dT[i][1], ticket: dT[i][2], folioCpp: dT[i][3], alertaTicket: dT[i][4], serie: dT[i][5], odometro: dT[i][6], fechaReg: limpiarHoraLectura(dT[i][7]), fechaServ: limpiarHoraLectura(dT[i][8]), tipoServicio: dT[i][9], familia: dT[i][10], numeroPartes: dT[i][11], proveedor: dT[i][12], descripcion: dT[i][13], quienRegistra: dT[i][14], servInterno: dT[i][15], servExterno: dT[i][16], costoTotal: dT[i][17], estatusUnidad: dT[i][18], evidencia: dT[i][19] });
      }
    }
    const dI = ss.getSheetByName("Inventario").getDataRange().getDisplayValues();
    let mSI = {};
    for (let i = 1; i < dI.length; i++) {
      mSI[dI[i][0]] = { stockDisp: parseFloat(dI[i][9])||0, cu: parseFloat(dI[i][11].replace(/[^0-9.-]+/g,""))||0 };
      if (dI[i][1].toString().toUpperCase().trim() === b) {
        dInv.push({ id: dI[i][0], ticket: dI[i][1], folioCpp: dI[i][2], obsTicket: dI[i][3], producto: dI[i][4], marca: dI[i][5], vehiculo: dI[i][6], unidadMedida: dI[i][7], fechaReg: limpiarHoraLectura(dI[i][8]), cantidad: dI[i][9], costoTotal: dI[i][10], costoUnitario: dI[i][11], estatus: dI[i][12], quienRegistra: dI[i][13] });
      }
    }
    const dC = ss.getSheetByName("Consumos").getDataRange().getDisplayValues();
    for (let i = 1; i < dC.length; i++) {
      if (dC[i][4].toString().toUpperCase().trim() === b || dC[i][3].toString().toUpperCase().trim() === b || dC[i][1].toString().toUpperCase().trim() === b) {
        let idP = dC[i][1];
        dCon.push({ id: dC[i][0], idPieza: idP, producto: dC[i][2], quienRegistra: dC[i][3], ticketInsumo: dC[i][4], vehiculo: dC[i][5], fecha: limpiarHoraLectura(dC[i][6]), cantidad: dC[i][7], costoTotal: dC[i][8], evidencia: dC[i][9], costoUnitario: mSI[idP]?mSI[idP].cu:0, stockDisponibleActual: mSI[idP]?mSI[idP].stockDisp:0 });
      }
    }
    return { exito: true, taller: dTal, inventario: dInv, consumos: dCon };
  } catch(e) { return { exito: false, error: e.message }; }
}

function actualizarFilaTallerBackend(d) {
  try {
    const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Taller");
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === d.id.toString().trim()) {
        let f = i + 1;
        hoja.getRange(f,  5).setValue(d.alertaTicket);
        hoja.getRange(f,  7).setValue(d.odometro);
        hoja.getRange(f,  9).setValue(limpiarHoraLectura(d.fechaServ));
        hoja.getRange(f, 10).setValue(d.tipoServicio);
        hoja.getRange(f, 11).setValue(d.familia);
        hoja.getRange(f, 12).setValue(d.numeroPartes);
        hoja.getRange(f, 13).setValue(d.proveedor);
        hoja.getRange(f, 14).setValue(d.descripcion);
        hoja.getRange(f, 16).setValue(d.servInterno);
        hoja.getRange(f, 17).setValue(d.servExterno);
        hoja.getRange(f, 18).setValue(d.costoTotal);
        hoja.getRange(f, 19).setValue(d.estatusUnidad);
        hoja.getRange(f, 20).setValue(d.evidencia || "SIN EVIDENCIA");
        SpreadsheetApp.flush();
        return { exito: true, msj: "Taller actualizado con éxito." };
      }
    }
    return { exito: false, error: "Registro no localizado." };
  } catch(e) { return { exito: false, error: "Fallo al actualizar registro." }; }
}

function actualizarFilaInventarioBackend(d) {
  try {
    const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario");
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === d.id.toString().trim()) {
        let f = i + 1;
        hoja.getRange(f,  3).setValue(d.folioCpp);
        hoja.getRange(f,  4).setValue(d.obsTicket);
        hoja.getRange(f,  5).setValue(d.producto);
        hoja.getRange(f,  6).setValue(d.marca);
        hoja.getRange(f,  8).setValue(d.unidadMedida);
        hoja.getRange(f, 10).setValue(d.cantidad);
        hoja.getRange(f, 11).setValue(d.costoTotal);
        hoja.getRange(f, 12).setValue(d.costoUnitario);
        hoja.getRange(f, 13).setValue(d.estatus);
        SpreadsheetApp.flush();
        return { exito: true, msj: "Actualizado con éxito." };
      }
    }
    return { exito: false, error: "Registro no localizado." };
  } catch(e) { return { exito: false, error: "Fallo al actualizar registro." }; }
}

function actualizarFilaConsumoBackend(d) {
  try {
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const hojaConsumos = ss.getSheetByName("Consumos");
    const hojaInv      = ss.getSheetByName("Inventario");
    const datosCons    = hojaConsumos.getDataRange().getValues();
    const datosInv     = hojaInv.getDataRange().getValues();
    let fC = -1; let oldC = 0; let idP = "";
    for (let i = 1; i < datosCons.length; i++) {
      if (datosCons[i][0].toString().trim() === d.id.toString().trim()) { fC = i+1; idP = datosCons[i][1].toString(); oldC = parseFloat(datosCons[i][7]) || 0; break; }
    }
    if (fC === -1) return { exito: false, error: "Consumo no encontrado." };
    let fI = -1; let stA = 0; let cU = 0;
    for (let i = 1; i < datosInv.length; i++) {
      if (datosInv[i][0].toString().trim() === idP) { fI = i+1; stA = parseFloat(datosInv[i][9]) || 0; cU = parseFloat(datosInv[i][11]) || 0; break; }
    }
    let newC = parseFloat(d.cantidad); let diff = newC - oldC;
    if (fI !== -1) {
      if (diff > stA) return { exito: false, error: "Stock insuficiente." };
      let nS = stA - diff;
      hojaInv.getRange(fI, 10).setValue(nS);
      hojaInv.getRange(fI, 11).setValue(nS * cU);
      hojaInv.getRange(fI, 13).setValue(nS <= 0 ? "CONSUMIDO" : "DISPONIBLE");
    }
    hojaConsumos.getRange(fC, 7).setValue(limpiarHoraLectura(d.fecha));
    hojaConsumos.getRange(fC, 8).setValue(newC);
    hojaConsumos.getRange(fC, 9).setValue(newC * cU);
    SpreadsheetApp.flush();
    return { exito: true, msj: "Consumo actualizado." };
  } catch(e) { return { exito: false, error: "Fallo al actualizar registro." }; }
}

function actualizarProveedorCompleto(d) {
  try {
    const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proveedores");
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === d.id.toString().trim()) {
        const f = i + 1;
        hoja.getRange(f,  2).setValue(d.proveedor);
        hoja.getRange(f,  3).setValue(d.razonSocial);
        hoja.getRange(f,  4).setValue(d.rfc);
        hoja.getRange(f,  5).setValue(d.direccion);
        hoja.getRange(f,  6).setValue(d.estado);
        hoja.getRange(f,  7).setValue(d.contacto);
        hoja.getRange(f,  8).setValue(d.telefono);
        hoja.getRange(f,  9).setValue(d.correo);
        hoja.getRange(f, 10).setValue(d.segmentacion);
        hoja.getRange(f, 11).setValue(d.diasCredito);
        hoja.getRange(f, 12).setValue(d.totalCredito);
        hoja.getRange(f, 13).setValue(d.servicio);
        hoja.getRange(f, 14).setValue(d.razonesSociales);
        SpreadsheetApp.flush();
        return { exito: true, msj: "Proveedor actualizado con éxito." };
      }
    }
    return { exito: false, error: "Registro no localizado." };
  } catch(e) { return { exito: false, error: "Fallo al actualizar." }; }
}
