/**
 * Procesa los recibos PDF (agua y luz) desde la carpeta "Recibos_Procesar" en Drive,
 * extrae los datos vía OCR, calcula el pago de luz/agua/alquiler por local
 * y deja el resultado en las hojas "Calculadora" y "Resumen_Cobro".
 *
 * Requiere el Advanced Service "Drive API" activado en el proyecto de Apps Script.
 */

function procesarRecibosYCalcular() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCalc = ss.getSheetByName("Calculadora");
  var sheetResumen = ss.getSheetByName("Resumen_Cobro");
  var sheetRespuestas = ss.getSheetByName("Form_Responses") || ss.getSheetByName("Respuestas de formulario 1") || ss.getSheets()[0];

  function notificar(msg) {
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  }

  function redondear1Decimal(num){
    return Math.round(num*10)/10;
  }

  if (!sheetCalc || !sheetResumen) {
    notificar("Error: Confirma que existan las pestañas 'Calculadora' y 'Resumen_Cobro'.");
    return;
  }

  var carpetas = DriveApp.getFoldersByName("Recibos_Procesar");
  if (!carpetas.hasNext()) {
    notificar("No se encontró la carpeta 'Recibos_Procesar' en Google Drive.");
    return;
  }

  var carpeta = carpetas.next();
  var archivos = carpeta.getFilesByType(MimeType.PDF); // Ahora filtramos PDFs directamente

  var ocrMontoLuz = 0;
  var ocrPu = 0;
  var ocrAgua = 0;
  var ocrDiaAct = "";
  var ocrDiaAnt = "";
  var ocrMesActStr = "";
  var ocrMesAntStr = "";
  var ocrFechaAct = ""; 
  var ocrFechaAnt = ""; 

  var cargosAfectosValores = [0, 0, 0, 0];
  var cargosNoAfectosValores = [0, 0, 0, 0, 0, 0, 0];

  // ==========================================================
  // Función auxiliar: convierte un PDF a texto real usando OCR
  // ==========================================================
  function extraerTextoDePdf(archivoPdf) {
    var recurso = {
      title: archivoPdf.getName() + "_OCR_TEMP",
      mimeType: archivoPdf.getMimeType()
    };
    var opciones = {
      ocr: true,
      ocrLanguage: "es"
    };
    var archivoTemp = Drive.Files.insert(recurso, archivoPdf.getBlob(), opciones);
    var doc = DocumentApp.openById(archivoTemp.id);
    var texto = doc.getBody().getText();

    // Limpieza: eliminamos el Google Doc temporal generado por el OCR
    DriveApp.getFileById(archivoTemp.id).setTrashed(true);

    return texto;
  }

  function buscarMontoMd(texto, patron) {
    var m = texto.match(patron);
    return m ? parseFloat(m[1].replace(/,/g, "")) : 0;
  }

  // ==========================================
  // PASO 1: LECTURA DE CADA PDF (AGUA O LUZ)
  // ==========================================
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    var nombreArchivo = archivo.getName().toLowerCase();
    var docText = extraerTextoDePdf(archivo);

    // --- PROCESAMIENTO RECIBO DE AGUA (SEDAPAL) ---
    if (nombreArchivo.indexOf("agua") !== -1) {
      var matchAgua = docText.match(/Importe\s*total\s*a\s*pagar\s*:?\s*S\/\s*\S*?([\d\.,]+)/i) ||
                  docText.match(/total\s*a\s*pagar[\s\S]*?S\/\s*\S*?([\d\.,]+)/i) ||
                  docText.match(/S\/\s*\S*?([\d\.,]+)/i);

      if (matchAgua) ocrAgua = parseFloat(matchAgua[1].replace(/,/g, ""));
    }

    // --- PROCESAMIENTO RECIBO DE LUZ (PLUZ/ENEL) ---
    else if (nombreArchivo.indexOf("luz") !== -1) {
      var matchTotalLuz = docText.match(/S\/\**([\d\.,]+)/i);
      if (matchTotalLuz) ocrMontoLuz = parseFloat(matchTotalLuz[1].replace(/,/g, ""));

      var matchPu = docText.match(/precio\s*de\s*S\/?\s*([\d\.]+)/i) || docText.match(/S\/\s*(0\.\d+)\s*por\s*kwh/i);
      if (matchPu) ocrPu = parseFloat(matchPu[1]);

      var matchFechas = docText.match(/Lectura\s*Actual[\s\S]*?\((\d{2})\/(\d{2})\/(\d{4})\)/i);
      var matchFechasAnt = docText.match(/Lectura\s*Anterior[\s\S]*?\((\d{2})\/(\d{2})\/(\d{4})\)/i);

      if (matchFechas && matchFechasAnt) {
        ocrFechaAct = matchFechas[1] + "/" + matchFechas[2] + "/" + matchFechas[3];
        ocrFechaAnt = matchFechasAnt[1] + "/" + matchFechasAnt[2] + "/" + matchFechasAnt[3];
        
        ocrDiaAct = parseInt(matchFechas[1], 10);   // ej. 11
        ocrDiaAnt = parseInt(matchFechasAnt[1], 10); // ej. 09

        var mesesLista = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
        ocrMesActStr = mesesLista[parseInt(matchFechas[2], 10) - 1];
        ocrMesAntStr = mesesLista[parseInt(matchFechasAnt[2], 10) - 1];
      }

      // Extraemos únicamente el bloque del "DETALLE DE IMPORTES"---
      var posInicio = docText.search(/Alumbrado\s*P[uú]blico/i);
      var subTextoNumeros = posInicio !== -1 ? docText.substring(posInicio) : docText;

      // Extraemos todos los números con formato decimal (soporta negativos y comas de miles)
      var todosLosNumeros = subTextoNumeros.match(/-?\d+[\.,]\d{2}/g);

      if (todosLosNumeros && todosLosNumeros.length >= 5) {
        // Limpiamos las comas de miles antes de parsear a Float
        var num = todosLosNumeros.map(function(n) {
          return parseFloat(n.replace(/,/g, ""));
        });

      cargosAfectosValores[0] = num[0]; // Reposición y Mant.
      cargosAfectosValores[1] = num[1]; // Cargo Fijo
      cargosAfectosValores[2] = num[3]; // Interés Compensatorio (salta el cargo de energía num[2])
      cargosAfectosValores[3] = num[4]; // Alumbrado Público

      // Cargos No Afectos
       if (num.length >= 12) {
          //cargosNoAfectosValores[0] = num[8];  // Deuda Anterior 
          cargosNoAfectosValores[1] = num[9];  // Aporte Ley 
          cargosNoAfectosValores[2] = num[10]; // DS 020 
          cargosNoAfectosValores[3] = num[11]; // DL 25844 
          //cargosNoAfectosValores[4] = num[12]; // Recargo por Mora 
          cargosNoAfectosValores[5] = num[13]; // Redondeo Mes Anterior
          cargosNoAfectosValores[6] = num[14];
        }
       }
       
    }
  }

  // ==================================================
  // PASO 2: INYECCIÓN EN LA COLUMNA B DE "CALCULADORA"
  // ==================================================
  sheetCalc.getRange("B2").setValue(ocrMontoLuz);
  sheetCalc.getRange("B3").setValue(ocrPu);
  sheetCalc.getRange("B4").setValue(ocrFechaAct).setNumberFormat("@");
  sheetCalc.getRange("B5").setValue(ocrFechaAnt).setNumberFormat("@");
  sheetCalc.getRange("B6").setValue(ocrMesActStr);
  sheetCalc.getRange("B7").setValue(ocrMesAntStr);
  sheetCalc.getRange("B22").setValue(ocrAgua);

  // Colocamos los importes de la luz en la columna B
  sheetCalc.getRange("B9").setValue(cargosAfectosValores[0]);
  sheetCalc.getRange("B10").setValue(cargosAfectosValores[1]);
  sheetCalc.getRange("B11").setValue(cargosAfectosValores[2]);
  sheetCalc.getRange("B12").setValue(cargosAfectosValores[3]);

  //sheetCalc.getRange("B13").setValue(cargosNoAfectosValores[0]);
  sheetCalc.getRange("B14").setValue(cargosNoAfectosValores[1]);
  sheetCalc.getRange("B15").setValue(cargosNoAfectosValores[2]);
  sheetCalc.getRange("B16").setValue(cargosNoAfectosValores[3]);
  //sheetCalc.getRange("B17").setValue(cargosNoAfectosValores[4]);
  sheetCalc.getRange("B18").setValue(cargosNoAfectosValores[5]);
  sheetCalc.getRange("B19").setValue(cargosNoAfectosValores[6]);

  // ===========================================
  // PASO 3: ANÁLISIS DE LECTURAS DEL FORMULARIO
  // ===========================================
  var datosForm = sheetRespuestas.getDataRange().getValues();
  var encabezados = datosForm[0];
  var idxLocal = 1;
  var idxMesCol = encabezados.length - 1;

  // Mapa de columnas por día (9 al 15)
  var colPorDia = {};
  for (var c = 0; c < encabezados.length; c++) {
    var colTexto = encabezados[c].toString().toLowerCase();
    var numEnCol = colTexto.match(/día\s*(\d+)/i) || colTexto.match(/\d+/);
    if (numEnCol) {
      colPorDia[parseInt(numEnCol[1] || numEnCol[0], 10)] = c;
    }
  }

  function norm(t) {
    return t.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  }

  var listaLocales = [
    { nombre: "Chifa", fila: 2 },
    { nombre: "Local 4", fila: 3 },
    { nombre: "Botica", fila: 4 },
    { nombre: "Panadería", fila: 5 }
  ];
  //EL PAGO DEL AGUA SE ACORDÓ ENTRE LOS INQUILINOS QUE, SOLO BOTICA NO PAGARÍA LA 4 PAR DEL TOTAL, SINO 65 SOLES
  var cuartaParteAgua = ocrAgua / 4;
  var aguaPorLocalDict = {
    "Chifa": cuartaParteAgua,
    "Panadería": cuartaParteAgua,
    "Botica": 65.00,
    "Local 4": cuartaParteAgua + (cuartaParteAgua - 65.00)
  };

  // Función interna para obtener la lectura válida (>0) más cercana a un día objetivo
  function buscarLecturaPorDia(registros, mesObjetivo, diaObjetivo) {
    var colIdx = colPorDia[diaObjetivo];
    if (colIdx === undefined) return 0;

    for (var r = registros.length - 1; r >= 0; r--) {
      var fila = registros[r];
      var mesFila = fila[idxMesCol] ? fila[idxMesCol].toString().trim() : "";
      if (norm(mesFila) === norm(mesObjetivo)) {
        var val = parseFloat(fila[colIdx]);
        if (!isNaN(val) && val > 0) {
          return val;
        }
      }
    }
    return 0;
  }

  var mesesNumero = {
    "enero": 0, "febrero": 1, "marzo": 2, "abril": 3, "mayo": 4, "junio": 5,
    "julio": 6, "agosto": 7, "septiembre": 8, "octubre": 9, "noviembre": 10, "diciembre": 11
  };

  /**
   * Obtiene la lectura de un local para un mes y día específicos.
   * Retorna 0 si no existe o es vacía/cero.
   */
  function obtenerLecturaDirecta(registros, mesObjetivo, diaObjetivo) {
    var colIdx = colPorDia[diaObjetivo];
    if (colIdx === undefined) return 0;

    for (var r = registros.length - 1; r >= 0; r--) {
      var fila = registros[r];
      var mesFila = fila[idxMesCol] ? fila[idxMesCol].toString().trim() : "";
      if (norm(mesFila) === norm(mesObjetivo)) {
        var val = parseFloat(fila[colIdx]);
        if (!isNaN(val) && val > 0) {
          return val;
        }
      }
    }
    return 0;
  }

  /**
   * Busca en paralelo (sincronizado por desfase/offset) las lecturas.
   * Empieza en (diaActBase, diaAntBase). Si no existen AMBAS a la vez,
   * avanza en paralelo a (diaActBase+1, diaAntBase+1), (diaActBase+2, diaAntBase+2)...
   */
  function obtenerLecturasSincronizadas(registros, mesAct, diaActBase, mesAnt, diaAntBase) {
    // Probamos offsets desde 0 hasta el límite de días (máximo día 15)
    for (var offset = 0; offset <= 6; offset++) {
      var dAct = diaActBase + offset;
      var dAnt = diaAntBase + offset;

      // Asegurarnos de no sobrepasar el rango de columnas (hasta día 15)
      if (dAct <= 15 && dAnt <= 15) {
        var l1 = obtenerLecturaDirecta(registros, mesAct, dAct);
        var l2 = obtenerLecturaDirecta(registros, mesAnt, dAnt);

        // CONDICIÓN CLAVE: Deben existir REGISTROS A LA VEZ en ambos meses
        if (l1 > 0 && l2 > 0) {
          return { L1: l1, L2: l2, diaActElegido: dAct, diaAntElegido: dAnt };
        }
      }
    }

    // Si avanzando no encontró, intentamos retroceder en paralelo (-1, -2, etc.)
    for (var offsetNeg = 1; offsetNeg <= 5; offsetNeg++) {
      var dActNeg = diaActBase - offsetNeg;
      var dAntNeg = diaAntBase - offsetNeg;

      if (dActNeg >= 9 && dAntNeg >= 9) {
        var l1_neg = obtenerLecturaDirecta(registros, mesAct, dActNeg);
        var l2_neg = obtenerLecturaDirecta(registros, mesAnt, dAntNeg);

        if (l1_neg > 0 && l2_neg > 0) {
          return { L1: l1_neg, L2: l2_neg, diaActElegido: dActNeg, diaAntElegido: dAntNeg };
        }
      }
    }

    return { L1: 0, L2: 0, diaActElegido: 0, diaAntElegido: 0 };
  }

  // LECTURA DINÁMICA DE PARÁMETROS DEL RECIBO (Variables que cambian cada mes)
  // B9: Reposic. y Mant., B10: Cargo fijo, B11: Interés Comp., B12: Alumbrado Público
  var cargosAfectosSuma = (parseFloat(sheetCalc.getRange("B9").getValue()) || 0) +
                          (parseFloat(sheetCalc.getRange("B10").getValue()) || 0) +
                          (parseFloat(sheetCalc.getRange("B11").getValue()) || 0) +
                          (parseFloat(sheetCalc.getRange("B12").getValue()) || 0);

  // B14: Deuda Ant., B15: Aporte Ley, B16: DS 020, B18: DL 25844, B19: Recargo Mora
  var cargosNoAfectosSuma = (parseFloat(sheetCalc.getRange("B14").getValue()) || 0) +
                            (parseFloat(sheetCalc.getRange("B15").getValue()) || 0) +
                            (parseFloat(sheetCalc.getRange("B16").getValue()) || 0) +
                            (parseFloat(sheetCalc.getRange("B18").getValue()) || 0) +
                            (parseFloat(sheetCalc.getRange("B19").getValue()) || 0);

  // B3: Precio unitario extraído del recibo
  var ocrPrecioUnitario = parseFloat(sheetCalc.getRange("B3").getValue()) || 0;

  // Bucle por cada local
  for (var k = 0; k < listaLocales.length; k++) {
    var loc = listaLocales[k];
    var aguaEsteLocal = redondear1Decimal(aguaPorLocalDict[loc.nombre] || 0);
    sheetResumen.getRange(loc.fila, 3).setValue(aguaEsteLocal);

    // Filtrar registros de este local
    var registrosLocal = [];
    for (var r = 1; r < datosForm.length; r++) {
      var fila = datosForm[r];
      if (norm(fila[idxLocal] || "").indexOf(norm(loc.nombre)) !== -1) {
        registrosLocal.push(fila);
      }
    }

    // Búsqueda paralela sincronizada
    var resBusqueda = obtenerLecturasSincronizadas(registrosLocal, ocrMesActStr, ocrDiaAct, ocrMesAntStr, ocrDiaAnt);
    var L1 = resBusqueda.L1;
    var L2 = resBusqueda.L2;

    if (L1 > 0 && L2 > 0 && L1 >= L2) {
      var consumoKwh = L1 - L2;

      // 1. Diferencia * Precio Unitario
      var costoConsumo = consumoKwh * ocrPrecioUnitario;

      // 2. Subtotal Inquilino = Costo Consumo + (Cargos Afectos * 0.25)
      var subTotalInquilino = costoConsumo + (cargosAfectosSuma * 0.25);

      // 3. IGV = Subtotal Inquilino * 18%
      var igv = subTotalInquilino * 0.18;

      // 4. Subtotal con IGV
      var subTotalConIgv = subTotalInquilino + igv;

      // 5. Cargos Adicionales / No Afectos * 0.25
      var cargosAdicionalesInquilino = cargosNoAfectosSuma * 0.25;

      // 6. MONTO TOTAL LUZ
      var totalPagarLuzLocal = redondear1Decimal(subTotalConIgv + cargosAdicionalesInquilino);

      // Guardar en hoja Resumen
      sheetResumen.getRange(loc.fila, 2).setValue(totalPagarLuzLocal);

      var montoAlquiler = parseFloat(sheetResumen.getRange(loc.fila, 4).getValue()) || 0;
      var totalFinal = redondear1Decimal(totalPagarLuzLocal + aguaEsteLocal + montoAlquiler);
      
      sheetResumen.getRange(loc.fila, 5).setValue(totalFinal);

      // Mensaje estructurado para WhatsApp
      var msgWs = "LIQUIDACIÓN MENSUAL - " + loc.nombre.toUpperCase() + "\n\n" +
                  "• Luz: S/ " + totalPagarLuzLocal.toFixed(2) + " (Consumo: " + consumoKwh.toFixed(2) + " kWh | Lecturas: Día " + resBusqueda.diaActElegido + " " + ocrMesActStr + " vs Día " + resBusqueda.diaAntElegido + " " + ocrMesAntStr + ")\n" +
                  "• Agua: S/ " + aguaEsteLocal.toFixed(2) + "\n" +
                  "• Alquiler: S/ " + montoAlquiler.toFixed(2) + "\n" +
                  "TOTAL A PAGAR: S/ " + totalFinal.toFixed(2)

      sheetResumen.getRange(loc.fila, 6).setValue(msgWs);
      
    } else {
      sheetResumen.getRange(loc.fila, 2).setValue(0);
      sheetResumen.getRange(loc.fila, 6).setValue("No se encontraron lecturas simétricas registradas para " + loc.nombre);
    }
  }

  SpreadsheetApp.flush();
  notificar("¡Listo! Datos procesados correctamente desde los recibos PDF.");
}