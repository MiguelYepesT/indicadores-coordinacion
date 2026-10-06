/*
 * procesar.js
 * Lee los Excel mensuales de ingresos y retiros y devuelve SOLO datos agregados
 * (conteos por mes y empresa). Ningún dato personal sale de esta función.
 *
 * Replica la lógica del Power BI "MES A MES":
 *  - Lee la tabla de Excel "INGRESOS" e "RETIROS" de cada archivo.
 *  - Ignora archivos cuyo nombre contenga "ANTERIOR" y archivos temporales (~$).
 *  - SARY INNOVATION y GRANEL FOOD se cuentan como SARY.
 *  - COMPLEMENTOS HUMANOS se muestra como COMPLEMENTOS.
 *  - Ingresos se asignan al mes de INGRESO; retiros al mes de RETIRO.
 *  - Rotación temprana: retiros con 0 a UMBRAL_DIAS días entre ingreso y retiro.
 */
(function (root) {
  'use strict';

  var UMBRAL_DIAS = 60;

  var MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
    'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  // Orden fijo de empresas. Si aparece una empresa nueva, se agrega al final.
  var ORDEN_EMPRESAS = ['GI GROUP', 'MISION', 'SARY', 'COMPLEMENTOS'];

  var ALIAS_EMPRESA = {
    'SARY INNOVATION': 'SARY',
    'GRANEL FOOD': 'SARY',
    'COMPLEMENTOS HUMANOS': 'COMPLEMENTOS',
    'MISIÓN': 'MISION'
  };

  function limpiar(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\s+/g, ' ').trim().toUpperCase();
  }

  function normalizarEmpresa(v) {
    var e = limpiar(v);
    return ALIAS_EMPRESA[e] || e;
  }

  function numeroMes(v) {
    var i = MESES.indexOf(limpiar(v));
    return i >= 0 ? i + 1 : null;
  }

  function entero(v) {
    if (v === null || v === undefined || v === '') return null;
    var n = Number(v);
    return isFinite(n) ? Math.trunc(n) : null;
  }

  function fecha(anio, mes, dia) {
    if (!anio || !mes || !dia) return null;
    var d = new Date(Date.UTC(anio, mes - 1, dia));
    // Rechaza fechas imposibles (ej. 31 de febrero)
    if (d.getUTCMonth() !== mes - 1 || d.getUTCDate() !== dia) return null;
    return d;
  }

  // ---------- Lectura de tablas de Excel (ListObjects) ----------

  function texto(file) {
    if (!file) return '';
    var c = file.content;
    if (typeof c === 'string') return c;
    return new TextDecoder('utf-8').decode(c);
  }

  function resolverRuta(base, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var partes = base.split('/');
    partes.pop();
    target.split('/').forEach(function (p) {
      if (p === '..') partes.pop();
      else if (p !== '.') partes.push(p);
    });
    return partes.join('/');
  }

  function relaciones(wb, rutaRels) {
    var xml = texto(wb.files[rutaRels]);
    var mapa = {};
    var re = /<Relationship\b[^>]*>/g, m;
    while ((m = re.exec(xml))) {
      var id = /\bId="([^"]+)"/.exec(m[0]);
      var tg = /\bTarget="([^"]+)"/.exec(m[0]);
      if (id && tg) mapa[id[1]] = tg[1];
    }
    return mapa;
  }

  // Devuelve { NOMBRE_TABLA: { hoja: 'NombreHoja', ref: 'A1:K76' } }
  function tablasDelLibro(wb) {
    var res = {};
    if (!wb.files) return res;
    try {
      var wbXml = texto(wb.files['xl/workbook.xml']);
      var relsWb = relaciones(wb, 'xl/_rels/workbook.xml.rels');
      var re = /<sheet\b[^>]*>/g, m;
      while ((m = re.exec(wbXml))) {
        var nombre = /\bname="([^"]+)"/.exec(m[0]);
        var rid = /\br:id="([^"]+)"/.exec(m[0]) || /\bid="([^"]+)"/.exec(m[0]);
        if (!nombre || !rid || !relsWb[rid[1]]) continue;
        var hojaNombre = nombre[1].replace(/&amp;/g, '&');
        var rutaHoja = resolverRuta('xl/workbook.xml', relsWb[rid[1]]);
        var partes = rutaHoja.split('/');
        var archivo = partes.pop();
        var rutaRels = partes.join('/') + '/_rels/' + archivo + '.rels';
        if (!wb.files[rutaRels]) continue;
        var relsHoja = relaciones(wb, rutaRels);
        Object.keys(relsHoja).forEach(function (k) {
          var t = relsHoja[k];
          if (!/tables\//.test(t)) return;
          var rutaTabla = resolverRuta(rutaHoja, t);
          var xmlT = texto(wb.files[rutaTabla]);
          var tn = /<table\b[^>]*\bname="([^"]+)"/.exec(xmlT);
          var tr = /<table\b[^>]*\bref="([^"]+)"/.exec(xmlT);
          if (tn && tr) res[limpiar(tn[1])] = { hoja: hojaNombre, ref: tr[1] };
        });
      }
    } catch (e) { /* si falla, se usa el plan B */ }
    return res;
  }

  // Lee filas como objetos { ENCABEZADO: valor }. Usa la tabla de Excel si existe;
  // si no, usa la hoja con el mismo nombre y se detiene en la primera fila vacía.
  function leerFilas(XLSX, wb, nombre, avisos, archivo) {
    var tablas = tablasDelLibro(wb);
    var t = tablas[nombre];
    var hoja, matriz;
    if (t && wb.Sheets[t.hoja]) {
      hoja = wb.Sheets[t.hoja];
      matriz = XLSX.utils.sheet_to_json(hoja, { header: 1, range: t.ref, raw: true, defval: null, blankrows: true });
    } else {
      var nombreHoja = wb.SheetNames.filter(function (n) { return limpiar(n) === nombre; })[0];
      if (!nombreHoja) {
        avisos.push(archivo + ': no tiene la hoja ni la tabla ' + nombre + '.');
        return [];
      }
      avisos.push(archivo + ': no se encontró la tabla ' + nombre + ', se leyó la hoja hasta la primera fila vacía.');
      matriz = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja], { header: 1, raw: true, defval: null, blankrows: true });
      var fin = matriz.findIndex(function (f, i) {
        return i > 0 && f.every(function (v) { return v === null || v === ''; });
      });
      if (fin > 0) matriz = matriz.slice(0, fin);
    }
    if (!matriz.length) return [];
    var enc = matriz[0].map(limpiar);
    return matriz.slice(1).map(function (fila) {
      var o = {};
      enc.forEach(function (h, i) { if (h) o[h] = fila[i]; });
      return o;
    });
  }

  // ---------- Utilidades nuevas ----------

  function sinTildes(t) { return t.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }

  function encabezado(v) { return sinTildes(limpiar(v)); }

  // Clasifica el motivo de retiro
  function tipoMotivo(v) {
    var m = encabezado(v);
    if (m.indexOf('RENUNCIA') >= 0) return 'renuncias';
    if (m.indexOf('TERMINACION') >= 0 || m.indexOf('DESPIDO') >= 0) return 'terminaciones';
    return 'otros';
  }

  // Solo los subprocesos que se grafican
  var SUBPROCESOS = ['EMPAQUE', 'PISO 3', 'MERCADEO MEDELLÍN', 'MERCADEO BOGOTÁ', 'MERCADEO CALI'];
  function normalizarSubproceso(v) {
    var s = encabezado(v);
    if (!s) return null;
    if (s.indexOf('EMPAQUE') >= 0) return 'EMPAQUE';
    if (/PISO\s*3/.test(s)) return 'PISO 3';
    if (s.indexOf('MERCADEO') >= 0 || s.indexOf('VENTAS') >= 0) {
      if (s.indexOf('MEDELLIN') >= 0) return 'MERCADEO MEDELLÍN';
      if (s.indexOf('BOGOTA') >= 0) return 'MERCADEO BOGOTÁ';
      if (s.indexOf('CALI') >= 0) return 'MERCADEO CALI';
    }
    return null;
  }

  // Convierte una celda de fecha de Excel a número de día (serial). Acepta número, Date o texto.
  function serialFecha(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number' && isFinite(v) && v > 20000 && v < 80000) return Math.floor(v);
    if (v instanceof Date && !isNaN(v)) return Math.floor((Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
    var t = String(v).trim(), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t))) return serialDe(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/.exec(t))) return serialDe(+m[3], +m[2], +m[1]);
    return null;
  }
  function serialDe(a, m, d) {
    var f = fecha(a, m, d);
    return f ? Math.round((f - Date.UTC(1899, 11, 30)) / 86400000) : null;
  }
  function partesDeSerial(n) {
    var d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
    return { anio: d.getUTCFullYear(), mes: d.getUTCMonth() + 1, dia: d.getUTCDate() };
  }
  function idPeriodo(anio, mes) { return anio + '-' + String(mes).padStart(2, '0'); }
  function textoFecha(n) { var p = partesDeSerial(n); return p.dia + '/' + p.mes + '/' + p.anio; }

  // Columna de la tabla "motivo de retiro por subproceso": los 5 subprocesos o, si no, la dirección
  var COLUMNAS_SUB = { 'EMPAQUE': 'Empaque', 'MERCADEO MEDELLÍN': 'Mercadeo Medellín', 'MERCADEO BOGOTÁ': 'Mercadeo Bogotá', 'MERCADEO CALI': 'Mercadeo Cali', 'PISO 3': 'Piso 3' };
  function columnaArea(sub, direccion) {
    if (sub && COLUMNAS_SUB[sub]) return COLUMNAS_SUB[sub];
    var d = encabezado(direccion);
    if (!d) return 'Sin dato';
    if (d.indexOf('OPERACION') === 0) return 'Operaciones';
    if (d.indexOf('SISTEMA') === 0) return 'Sistema de gestión';
    if (d.indexOf('MANTENIMIENTO') === 0) return 'Mantenimiento';
    if (d.indexOf('INNOVACION') === 0 || d === 'I+D') return 'I+D';
    if (d.indexOf('DESARROLLO HUMANO') === 0) return 'Desarrollo humano';
    if (d.indexOf('PRODUCCION') === 0) return 'Producción (otros)';
    if (d.indexOf('MERCADEO') === 0 || d.indexOf('VENTAS') === 0) return 'Mercadeo (otros)';
    if (d.indexOf('ADMINISTRA') === 0) return 'Administrativo';
    return textoLimpio(direccion);
  }

  // Texto legible para motivos y causas: sin espacios raros, primera letra en mayúscula
  function textoLimpio(v) {
    if (v === null || v === undefined) return '';
    var t = String(v).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    if (!t) return '';
    t = t.toLowerCase();
    return t.charAt(0).toUpperCase() + t.slice(1);
  }

  // Busca en el libro una tabla (o una hoja) cuyos encabezados incluyan todos los requeridos
  function buscarPorEncabezados(XLSX, wb, requeridos) {
    var tablas = tablasDelLibro(wb);
    var candidatas = Object.keys(tablas).map(function (k) { return tablas[k]; });
    wb.SheetNames.forEach(function (n) { candidatas.push({ hoja: n, ref: null }); });
    for (var i = 0; i < candidatas.length; i++) {
      var c = candidatas[i];
      var hoja = wb.Sheets[c.hoja];
      if (!hoja) continue;
      var opts = { header: 1, raw: true, defval: null, blankrows: true };
      if (c.ref) opts.range = c.ref;
      var matriz = XLSX.utils.sheet_to_json(hoja, opts);
      if (!matriz.length) continue;
      var enc = matriz[0].map(encabezado);
      var ok = requeridos.every(function (r) { return enc.some(function (h) { return h.indexOf(r) === 0; }); });
      if (!ok) continue;
      if (!c.ref) {
        var fin = matriz.findIndex(function (f, j) { return j > 0 && f.every(function (v) { return v === null || v === ''; }); });
        if (fin > 0) matriz = matriz.slice(0, fin);
      }
      return { enc: enc, filas: matriz.slice(1) };
    }
    return null;
  }
  function columna(enc, prefijo) {
    for (var i = 0; i < enc.length; i++) if (enc[i].indexOf(prefijo) === 0) return i;
    return -1;
  }

  // Lee la planta (# total de empleados) de la hoja "Indicador de rotacion"
  function leerPlanta(XLSX, wb) {
    var nombre = wb.SheetNames.filter(function (n) { return encabezado(n).indexOf('INDICADOR') === 0; })[0];
    if (!nombre) return null;
    var m = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: null });
    var total = null, titulo = null;
    for (var i = 0; i < m.length; i++) {
      for (var j = 0; j < m[i].length; j++) {
        var v = m[i][j];
        if (titulo === null && typeof v === 'string') {
          var t = /^([A-ZÁÉÍÓÚ]+)\s+(?:DE\s+)?(\d{4})$/.exec(limpiar(v));
          if (t && MESES.indexOf(t[1]) >= 0) titulo = { mes: MESES.indexOf(t[1]) + 1, anio: +t[2], texto: limpiar(v) };
        }
        if (total === null && encabezado(v) === 'TOTAL' && typeof m[i][j + 1] === 'number' && m[i][j + 1] > 20) total = Math.round(m[i][j + 1]);
      }
    }
    return total ? { total: total, titulo: titulo } : null;
  }

  // Lee un archivo de "total de empleados" con bloques por mes:
  // una celda "MES DE AÑO" seguida de una fila "TOTAL | número".
  function nombreArea(v) {
    var a = encabezado(v);
    if (a.indexOf('MERCADEO') === 0 || a === 'VENTAS' || a === 'VENTAS NACIONALES') return 'Ventas nacionales';
    if (a.indexOf('INTERNACIONAL') >= 0) return 'Ventas internacionales';
    if (/\+/.test(a) || a.length <= 3) return limpiar(v);
    var t = limpiar(v).toLowerCase();
    return t.charAt(0).toUpperCase() + t.slice(1);
  }
  function leerArchivoPlanta(XLSX, wb) {
    var totales = {}, areas = {}, resumen = {}, n = 0;
    wb.SheetNames.forEach(function (nombre) {
      var m = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: null });
      var actual = null, enAreas = false;
      m.forEach(function (fila) {
        var celdas = fila.filter(function (v) { return v !== null && v !== ''; });
        if (!celdas.length) { enAreas = false; return; }
        for (var j = 0; j < fila.length; j++) {
          var v = fila[j];
          if (typeof v !== 'string') continue;
          var t = /^([A-ZÁÉÍÓÚ]+)\s+(?:DE\s+)?(\d{4})$/.exec(limpiar(v));
          if (t && MESES.indexOf(t[1]) >= 0) { actual = idPeriodo(+t[2], MESES.indexOf(t[1]) + 1); enAreas = false; return; }
          if (!actual) return;
          var et = encabezado(v), num = fila[j + 1], ret = fila[j + 2], R = resumen[actual];
          if (et === 'TOTAL' && typeof num === 'number' && num > 20) {
            if (!totales[actual]) {
              totales[actual] = Math.round(num); n++; areas[actual] = []; enAreas = true;
              resumen[actual] = { planta: Math.round(num), retiros: typeof ret === 'number' ? Math.round(ret) : null, renuncias: null, terminaciones: null, vinculacion: null };
            }
            return;
          }
          // Renuncias y terminaciones (dos formatos: fila con número o texto entre paréntesis)
          var mr = /\((\d+)\s*RENUNCIAS?\)/.exec(et), mt = /CONTROLADA\s*=\s*(\d+)/.exec(et);
          if (R && mr) R.renuncias = +mr[1];
          else if (R && mt) R.terminaciones = +mt[1];
          else if (R && /^RENUNCIAS?$/.test(et) && typeof num === 'number') R.renuncias = Math.round(num);
          else if (R && /^TERMINACION/.test(et) && /DIRECTA/.test(et) && typeof num === 'number') R.vinculacion = Math.round(num);
          else if (R && /^TERMINACION/.test(et) && typeof num === 'number') R.terminaciones = Math.round(num);
          if (/^(RENUNCIA|TERMINACION|AREAS|SIN ROTACION|ROTACION)/.test(et)) { enAreas = false; return; }
          if (enAreas && typeof num === 'number') {
            areas[actual].push({ area: nombreArea(v), planta: Math.round(num), retiros: typeof ret === 'number' ? Math.round(ret) : 0 });
          }
          return;
        }
      });
    });
    return n ? { totales: totales, areas: areas, resumen: resumen } : null;
  }



  // ---------- Procesamiento principal ----------

  /*
   * archivos: [{ nombre: 'ruta/archivo.xlsx', datos: ArrayBuffer|Uint8Array }]
   * Devuelve { data, usados, omitidos, avisos, planta }
   *  - Los Excel mensuales (con tabla INGRESOS) aportan ingresos, retiros y planta.
   *  - El consolidado de retiros (con FECHA DE RETIRO) reemplaza los retiros de la carpeta
   *    en todos los meses que contenga, y aporta el detalle por subproceso.
   */
  function procesar(XLSX, archivos) {
    var avisos = [], usados = [], omitidos = [];
    var ingresos = [];        // { periodo, empresa }
    var retCarpeta = [];      // { periodo, empresa, motivo, rt }
    var retConsol = [];       // { periodo, empresa, motivo, rt, sub }
    var planta = {};          // periodo -> { valor, origen, aviso }
    var consolidados = [];    // versiones encontradas del consolidado
    var archivosPlanta = [];  // versiones encontradas del archivo de total de empleados
    var areasPorMes = {};     // periodo -> [{ area, planta, retiros }] según el archivo de total de empleados

    archivos
      .slice()
      .sort(function (a, b) { return a.nombre.localeCompare(b.nombre); })
      .forEach(function (f) {
        var base = f.nombre.split('/').pop();
        if (!/\.xlsx?$/i.test(base)) return;
        if (base.indexOf('~$') === 0) { omitidos.push(f.nombre + ' (archivo temporal de Excel)'); return; }
        if (/ANTERIOR/i.test(base)) { omitidos.push(f.nombre + ' (versión anterior)'); return; }

        var wb;
        try { wb = XLSX.read(f.datos, { type: 'array', bookFiles: true }); }
        catch (e) { avisos.push(f.nombre + ': no se pudo abrir (' + e.message + ').'); return; }

        // ¿Es el consolidado de retiros?
        var consol = buscarPorEncabezados(XLSX, wb, ['FECHA DE RETIRO', 'EMPRESA']);
        var tieneIngresos = wb.SheetNames.some(function (n) { return limpiar(n) === 'INGRESOS'; }) || tablasDelLibro(wb).INGRESOS;
        if (consol && !tieneIngresos) {
          var e = consol.enc;
          var cEmp = columna(e, 'EMPRESA'), cFi = columna(e, 'FECHA DE INGRESO'), cFr = columna(e, 'FECHA DE RETIRO'),
            cMr = columna(e, 'MES DE RETIRO'), cAr = columna(e, 'ANO RETIRO'), cMot = columna(e, 'MOTIVO DE TERMINA'),
            cSub = columna(e, 'SUBPROCESO');
          if (cAr < 0) cAr = columna(e, 'ANO DE RETIRO');
          var cGen = columna(e, 'MOTIVO GENERICO'), cCausa = columna(e, 'CAUSA DE LA TERMINACION'), cDir = columna(e, 'DIRECCION');
          var n = 0, meses = {}, filasArchivo = [], avisosArchivo = [];
          consol.filas.forEach(function (r, i) {
            if (r.every(function (v) { return v === null || v === ''; })) return;
            var fila = i + 2;
            var sr = serialFecha(r[cFr]);
            var per;
            if (sr !== null) { var p = partesDeSerial(sr); per = idPeriodo(p.anio, p.mes); }
            else {
              var mr = numeroMes(r[cMr]), ar = entero(r[cAr]);
              if (!mr || !ar) { avisosArchivo.push(base + ' · fila ' + fila + ': sin fecha de retiro, no se contó.'); return; }
              per = idPeriodo(ar, mr);
            }
            var empresa = normalizarEmpresa(r[cEmp]);
            var si = serialFecha(r[cFi]);
            var rt = false, diasPerm = null;
            if (si !== null && sr !== null) {
              var dias = sr - si;
              if (dias >= 0) diasPerm = dias;
              rt = dias >= 0 && dias <= UMBRAL_DIAS;
              if (dias < 0) avisosArchivo.push(base + ' · fila ' + fila + ' (' + empresa + '): la fecha de retiro es anterior a la de ingreso.');
            } else {
              avisosArchivo.push(base + ' · fila ' + fila + ' (' + empresa + (sr !== null ? ', retiro ' + textoFecha(sr) : '') + '): sin fecha de ingreso, no entra en rotación temprana.');
            }
            var textoMotivo = encabezado(cMot >= 0 ? r[cMot] : '');
            var subN = cSub >= 0 ? normalizarSubproceso(r[cSub]) : null;
            filasArchivo.push({ periodo: per, empresa: empresa, motivo: tipoMotivo(cMot >= 0 ? r[cMot] : ''), rt: rt, dias: diasPerm,
              sub: subN, columna: columnaArea(subN, cDir >= 0 ? r[cDir] : ''),
              antecedentes: textoMotivo.indexOf('ANTECEDENTE') >= 0,
              motivoGen: cGen >= 0 ? textoLimpio(r[cGen]) : '', causa: cCausa >= 0 ? textoLimpio(r[cCausa]) : '' });
            meses[per] = true;
            n++;
          });
          consolidados.push({ nombre: f.nombre, fecha: f.fecha || 0, filas: filasArchivo, avisos: avisosArchivo,
            uso: { archivo: f.nombre, tipo: 'consolidado', ingresos: 0, retiros: n, meses: Object.keys(meses).sort() } });
          return;
        }

        // ¿Es el archivo de total de empleados?
        if (!tieneIngresos) {
          var valores = leerArchivoPlanta(XLSX, wb);
          if (valores) {
            archivosPlanta.push({ nombre: f.nombre, fecha: f.fecha || 0, valores: valores.totales, areas: valores.areas, resumen: valores.resumen });
            return;
          }
          avisos.push(f.nombre + ': no se reconoció como Excel mensual, consolidado de retiros ni total de empleados. No se usó.');
          return;
        }

        // Excel mensual de la carpeta
        var ing = leerFilas(XLSX, wb, 'INGRESOS', avisos, f.nombre);
        var ret = leerFilas(XLSX, wb, 'RETIROS', avisos, f.nombre);
        var nIng = 0, nRet = 0, conteoMes = {};

        ing.forEach(function (r, i) {
          var vacia = Object.keys(r).every(function (k) { return r[k] === null || r[k] === ''; });
          if (vacia) return;
          var mes = numeroMes(r['MES INGRESO']);
          var anio = entero(r['AÑO INGRESO']);
          if (!mes || !anio) {
            avisos.push(f.nombre + ' · INGRESOS fila ' + (i + 2) + ': sin mes o año de ingreso válido, no se contó.');
            return;
          }
          var per = idPeriodo(anio, mes);
          ingresos.push({ periodo: per, empresa: normalizarEmpresa(r['EMPRESA']) });
          conteoMes[per] = (conteoMes[per] || 0) + 1;
          nIng++;
        });

        ret.forEach(function (r, i) {
          var vacia = Object.keys(r).every(function (k) { return r[k] === null || r[k] === ''; });
          if (vacia) return;
          var mesR = numeroMes(r['MES RETIRO']);
          var anioR = entero(r['AÑO RETIRO']);
          if (!mesR || !anioR) {
            avisos.push(f.nombre + ' · RETIROS fila ' + (i + 2) + ': sin mes o año de retiro, no se contó.');
            return;
          }
          var fi = fecha(entero(r['AÑO INGRESO']), numeroMes(r['MES INGRESO']), entero(r['DIA INGRESO']));
          var fr = fecha(anioR, mesR, entero(r['DIA RETIRO']));
          var rt = false;
          if (fi && fr) { var dias = Math.round((fr - fi) / 86400000); rt = dias >= 0 && dias <= UMBRAL_DIAS; }
          else avisos.push(f.nombre + ' · RETIROS fila ' + (i + 2) + ': fecha de ingreso o retiro incompleta, no entra en rotación temprana.');
          retCarpeta.push({ periodo: idPeriodo(anioR, mesR), empresa: normalizarEmpresa(r['EMPRESA']), motivo: tipoMotivo(r['MOTIVO RETIRO']), rt: rt });
          nRet++;
        });

        // Mes del archivo = el mes con más ingresos; ahí se asigna la planta
        var mesArchivo = Object.keys(conteoMes).sort(function (a, b) { return conteoMes[b] - conteoMes[a]; })[0];
        var pl = leerPlanta(XLSX, wb);
        if (mesArchivo && pl) {
          var aviso = '';
          if (pl.titulo && idPeriodo(pl.titulo.anio, pl.titulo.mes) !== mesArchivo)
            aviso = 'La hoja "Indicador de rotación" de este archivo dice ' + pl.titulo.texto + '. Puede ser una copia del mes anterior: verifica.';
          planta[mesArchivo] = { valor: pl.total, origen: 'Hoja del Excel mensual', aviso: aviso };
        }
        usados.push({ archivo: f.nombre, tipo: 'mensual', ingresos: nIng, retiros: nRet, meses: mesArchivo ? [mesArchivo] : [] });
      });

    // Si hay varias versiones del consolidado se usa solo la más reciente (por fecha de modificación)
    if (consolidados.length) {
      consolidados.sort(function (a, b) { return a.fecha - b.fecha || a.nombre.localeCompare(b.nombre); });
      var elegido = consolidados[consolidados.length - 1];
      consolidados.slice(0, -1).forEach(function (c) {
        omitidos.push(c.nombre + ' (otra versión del consolidado de retiros; se usó ' + elegido.nombre.split('/').pop() + ', la más reciente)');
      });
      retConsol = elegido.filas;
      avisos = avisos.concat(elegido.avisos);
      usados.push(elegido.uso);
    }

    // Archivos de resumen mensual (TOTAL EMPLEADOS, resúmenes de años anteriores).
    // Tienen prioridad sobre las hojas "Indicador de rotación". Si dos traen el mismo mes, gana el más reciente.
    var resumenArchivo = {};
    archivosPlanta.sort(function (a, b) { return a.fecha - b.fecha || a.nombre.localeCompare(b.nombre); });
    archivosPlanta.forEach(function (ap) {
      Object.keys(ap.valores).forEach(function (per) {
        planta[per] = { valor: ap.valores[per], origen: 'Archivo de total de empleados', aviso: '' };
        areasPorMes[per] = ap.areas[per];
        resumenArchivo[per] = ap.resumen[per];
      });
      usados.push({ archivo: ap.nombre, tipo: 'planta', ingresos: 0, retiros: 0, meses: Object.keys(ap.valores).sort() });
    });

    // Los meses del consolidado reemplazan los retiros de la carpeta
    var mesesConsol = {};
    retConsol.forEach(function (r) { mesesConsol[r.periodo] = true; });
    var retiros = retCarpeta.filter(function (r) { return !mesesConsol[r.periodo]; }).concat(retConsol);

    // Agregación por periodo y empresa
    var celdas = {}, empresasVistas = {}, periodos = {};
    function celda(per, emp) {
      var k = per + '|' + emp;
      if (!celdas[k]) celdas[k] = { ingresos: 0, retiros: 0, rotTemprana: 0, renuncias: 0, terminaciones: 0 };
      empresasVistas[emp] = true;
      if (!periodos[per]) periodos[per] = { ingresos: false, retiros: false };
      return celdas[k];
    }
    ingresos.forEach(function (r) { celda(r.periodo, r.empresa).ingresos++; periodos[r.periodo].ingresos = true; });
    retiros.forEach(function (r) {
      var c = celda(r.periodo, r.empresa);
      c.retiros++;
      if (r.rt) c.rotTemprana++;
      if (r.motivo === 'renuncias') c.renuncias++;
      else if (r.motivo === 'terminaciones') c.terminaciones++;
      periodos[r.periodo].retiros = true;
    });

    var empresas = ORDEN_EMPRESAS.filter(function (e) { return empresasVistas[e]; });
    Object.keys(empresasVistas).sort().forEach(function (e) { if (empresas.indexOf(e) < 0) empresas.push(e); });

    var filas = Object.keys(celdas).map(function (k) {
      var p = k.split('|'), c = celdas[k];
      return { periodo: p[0], empresa: p[1], ingresos: c.ingresos, retiros: c.retiros, rotTemprana: c.rotTemprana, renuncias: c.renuncias, terminaciones: c.terminaciones };
    }).sort(function (a, b) {
      return a.periodo.localeCompare(b.periodo) || empresas.indexOf(a.empresa) - empresas.indexOf(b.empresa);
    });

    // Subprocesos (solo del consolidado)
    var sub = {};
    retConsol.forEach(function (r) {
      if (!r.sub) return;
      var k = r.periodo + '|' + r.sub;
      if (!sub[k]) sub[k] = { periodo: r.periodo, subproceso: r.sub, renuncias: 0, terminaciones: 0, otros: 0, rotTemprana: 0, antecedentes: 0, dias: [] };
      sub[k][r.motivo]++;
      if (r.antecedentes) sub[k].antecedentes++;
      if (r.rt) sub[k].rotTemprana++;
      if (r.dias !== null && r.dias !== undefined) sub[k].dias.push(r.dias);
    });

    // Motivos de renuncia y causas de terminación por mes (solo del consolidado)
    var retirosDetalle = {};
    retConsol.forEach(function (r) {
      var d = retirosDetalle[r.periodo] || (retirosDetalle[r.periodo] = { antecedentes: 0, motivosRenuncia: {}, causasTerminacion: {}, porArea: { renuncias: {}, terminaciones: {} } });
      if (r.antecedentes) d.antecedentes++;
      var col = r.columna || 'Sin dato', t;
      if (r.motivo === 'renuncias') {
        var mg = r.motivoGen || 'Sin dato';
        d.motivosRenuncia[mg] = (d.motivosRenuncia[mg] || 0) + 1;
        t = d.porArea.renuncias[mg] || (d.porArea.renuncias[mg] = {}); t[col] = (t[col] || 0) + 1;
      }
      if (r.motivo === 'terminaciones') {
        var ca = r.causa || 'Sin dato';
        d.causasTerminacion[ca] = (d.causasTerminacion[ca] || 0) + 1;
        t = d.porArea.terminaciones[ca] || (d.porArea.terminaciones[ca] = {}); t[col] = (t[col] || 0) + 1;
      }
    });

    var data = {
      version: 2,
      generado: new Date().toISOString(),
      umbralDias: UMBRAL_DIAS,
      periodos: Object.keys(periodos).sort().map(function (p) {
        var a = Number(p.slice(0, 4)), m = Number(p.slice(5, 7));
        return {
          id: p, anio: a, mes: m, nombre: MESES[m - 1],
          planta: planta[p] ? planta[p].valor : null,
          tieneIngresos: periodos[p].ingresos,
          tieneRetiros: periodos[p].retiros,
          fuenteRetiros: mesesConsol[p] ? 'consolidado' : 'carpeta',
          tieneSubprocesos: !!mesesConsol[p]
        };
      }),
      empresas: empresas,
      datos: filas,
      retirosDetalle: retirosDetalle,
      areasPorMes: areasPorMes,
      resumenArchivo: resumenArchivo,
      subprocesos: {
        lista: SUBPROCESOS,
        datos: Object.keys(sub).sort().map(function (k) { sub[k].dias.sort(function (x, y) { return x - y; }); return sub[k]; })
      }
    };

    return { data: data, usados: usados, omitidos: omitidos, avisos: avisos, planta: planta };
  }

  // ---------- Fusión: cargar un mes sobre los datos ya publicados ----------

  function totalesPeriodo(d, p) {
    var t = { ingresos: 0, retiros: 0, rotTemprana: 0, renuncias: 0, terminaciones: 0 };
    d.datos.forEach(function (x) { if (x.periodo === p) Object.keys(t).forEach(function (k) { t[k] += x[k] || 0; }); });
    var per = (d.periodos || []).filter(function (x) { return x.id === p; })[0];
    t.planta = per ? per.planta : null;
    t.tieneIngresos = per ? per.tieneIngresos !== false : false;
    t.fuenteRetiros = per ? per.fuenteRetiros : null;
    t.existe = !!per;
    return t;
  }

  /*
   * base: data.json ya publicado (descifrado)
   * r: resultado de procesar() con SOLO los archivos que se adjuntaron
   * mes: periodo elegido ('2026-10')
   * plantaManual: número escrito a mano para ese mes (opcional)
   */
  function fusionar(base, r, mes, plantaManual) {
    var d = JSON.parse(JSON.stringify(base));
    var avisos = [], tocados = {}, errores = [];
    d.datos = d.datos || [];
    d.subprocesos = d.subprocesos || { lista: SUBPROCESOS, datos: [] };
    d.subprocesos.lista = SUBPROCESOS;
    if (!d.plantaPorMes) {
      d.plantaPorMes = {}; d.plantaOrigen = {};
      (d.periodos || []).forEach(function (p) { if (p.planta) { d.plantaPorMes[p.id] = p.planta; d.plantaOrigen[p.id] = 'anterior'; } });
    }
    var perMap = {};
    (d.periodos || []).forEach(function (p) { perMap[p.id] = JSON.parse(JSON.stringify(p)); });
    function periodo(id) {
      if (!perMap[id]) perMap[id] = { id: id, tieneIngresos: false, tieneRetiros: false, fuenteRetiros: 'carpeta', tieneSubprocesos: false };
      return perMap[id];
    }
    function fila(p, e) {
      var f = d.datos.filter(function (x) { return x.periodo === p && x.empresa === e; })[0];
      if (!f) { f = { periodo: p, empresa: e, ingresos: 0, retiros: 0, rotTemprana: 0, renuncias: 0, terminaciones: 0 }; d.datos.push(f); }
      return f;
    }
    var CAMPOS_RET = ['retiros', 'rotTemprana', 'renuncias', 'terminaciones'];
    function copiarRetiros(p) {
      d.datos.forEach(function (x) { if (x.periodo === p) CAMPOS_RET.forEach(function (k) { x[k] = 0; }); });
      r.data.datos.forEach(function (x) { if (x.periodo === p) { var f = fila(p, x.empresa); CAMPOS_RET.forEach(function (k) { f[k] = x[k] || 0; }); } });
    }
    var nombreMes = MESES[Number(mes.slice(5, 7)) - 1] + ' ' + mes.slice(0, 4);
    var rPer = {}; r.data.periodos.forEach(function (p) { rPer[p.id] = p; });

    // 1. Ingresos del mes elegido (Excel mensual)
    var hayMensual = r.usados.some(function (u) { return u.tipo === 'mensual'; });
    if (hayMensual) {
      var mesesIng = r.data.periodos.filter(function (p) { return p.tieneIngresos; }).map(function (p) { return p.id; });
      if (mesesIng.indexOf(mes) < 0) {
        errores.push('El Excel mensual no trae ingresos de ' + nombreMes + (mesesIng.length ? ' (trae de ' + mesesIng.map(function (p) { return MESES[Number(p.slice(5, 7)) - 1]; }).join(', ') + ')' : '') + '. Revisa que sea el archivo correcto o que el mes elegido sea el correcto.');
      } else {
        d.datos.forEach(function (x) { if (x.periodo === mes) x.ingresos = 0; });
        r.data.datos.forEach(function (x) { if (x.periodo === mes && x.ingresos) fila(mes, x.empresa).ingresos = x.ingresos; });
        periodo(mes).tieneIngresos = true;
        tocados[mes] = true;
        mesesIng.filter(function (p) { return p !== mes; }).forEach(function (p) {
          var n = r.data.datos.filter(function (x) { return x.periodo === p; }).reduce(function (a, x) { return a + x.ingresos; }, 0);
          avisos.push('El Excel mensual tiene ' + n + ' ingreso(s) de ' + MESES[Number(p.slice(5, 7)) - 1] + '. No se contaron: solo se carga ' + nombreMes + '.');
        });
      }
    }

    // 2. Retiros: el consolidado reemplaza todos los meses que trae
    var consol = r.usados.filter(function (u) { return u.tipo === 'consolidado'; })[0];
    var mesesConsol = consol ? consol.meses : [];
    mesesConsol.forEach(function (p) {
      copiarRetiros(p);
      var per = periodo(p);
      per.tieneRetiros = true; per.fuenteRetiros = 'consolidado'; per.tieneSubprocesos = true;
      d.subprocesos.datos = d.subprocesos.datos.filter(function (x) { return x.periodo !== p; })
        .concat(r.data.subprocesos.datos.filter(function (x) { return x.periodo === p; }));
      d.retirosDetalle = d.retirosDetalle || {};
      if (r.data.retirosDetalle && r.data.retirosDetalle[p]) d.retirosDetalle[p] = r.data.retirosDetalle[p];
      tocados[p] = true;
    });
    if (hayMensual && mesesConsol.indexOf(mes) < 0 && errores.length === 0) {
      var perBase = perMap[mes];
      if (perBase && perBase.fuenteRetiros === 'consolidado') {
        avisos.push('Los retiros de ' + nombreMes + ' ya vienen del consolidado publicado; se ignoró la hoja RETIROS del Excel mensual.');
      } else if (rPer[mes] && rPer[mes].tieneRetiros) {
        copiarRetiros(mes);
        var pm = periodo(mes);
        pm.tieneRetiros = true; pm.fuenteRetiros = 'carpeta'; pm.tieneSubprocesos = false;
        avisos.push('No adjuntaste el consolidado de retiros: los retiros de ' + nombreMes + ' se tomaron de la hoja RETIROS del Excel mensual.');
      }
    }

    // 3. Planta: archivo de total de empleados > lo escrito a mano > hoja del Excel mensual
    Object.keys(r.planta).forEach(function (p) {
      var o = r.planta[p];
      if (/Archivo/.test(o.origen)) {
        if (d.plantaPorMes[p] !== o.valor && perMap[p]) tocados[p] = true;
        d.plantaPorMes[p] = o.valor; d.plantaOrigen[p] = 'archivo';
      } else if (p === mes && d.plantaOrigen[p] !== 'archivo' && !(plantaManual > 0)) {
        d.plantaPorMes[p] = o.valor; d.plantaOrigen[p] = 'hoja';
        if (o.aviso) avisos.push(o.aviso);
      }
    });
    var historico = [];
    var resNuevo = r.data.resumenArchivo || {};
    if (Object.keys(resNuevo).length) {
      d.resumenArchivo = d.resumenArchivo || {};
      Object.keys(resNuevo).forEach(function (p) {
        var cambia = JSON.stringify(d.resumenArchivo[p]) !== JSON.stringify(resNuevo[p]);
        d.resumenArchivo[p] = resNuevo[p];
        if (cambia && !perMap[p]) historico.push(p);
        if (cambia && perMap[p]) tocados[p] = true;
      });
    }
    var areasNuevas = r.data.areasPorMes || {};
    if (Object.keys(areasNuevas).length) {
      d.areasPorMes = d.areasPorMes || {};
      Object.keys(areasNuevas).forEach(function (p) {
        if (JSON.stringify(d.areasPorMes[p]) !== JSON.stringify(areasNuevas[p]) && perMap[p]) tocados[p] = true;
        d.areasPorMes[p] = areasNuevas[p];
      });
    }
    if (plantaManual > 0) {
      if (d.plantaPorMes[mes] !== plantaManual) tocados[mes] = true;
      d.plantaPorMes[mes] = plantaManual; d.plantaOrigen[mes] = 'manual';
    }

    // 4. Reconstruir periodos, empresas y orden
    d.datos = d.datos.filter(function (x) {
      return x.ingresos || x.retiros || x.rotTemprana || x.renuncias || x.terminaciones;
    });
    var vistas = {};
    d.datos.forEach(function (x) { vistas[x.empresa] = true; periodo(x.periodo); });
    var empresas = ORDEN_EMPRESAS.filter(function (e) { return vistas[e]; });
    Object.keys(vistas).sort().forEach(function (e) { if (empresas.indexOf(e) < 0) empresas.push(e); });
    d.empresas = empresas;
    d.periodos = Object.keys(perMap).sort().filter(function (id) {
      return d.datos.some(function (x) { return x.periodo === id; });
    }).map(function (id) {
      var p = perMap[id];
      var a = Number(id.slice(0, 4)), m = Number(id.slice(5, 7));
      return {
        id: id, anio: a, mes: m, nombre: MESES[m - 1],
        planta: d.plantaPorMes[id] || null,
        tieneIngresos: !!p.tieneIngresos,
        tieneRetiros: d.datos.some(function (x) { return x.periodo === id && x.retiros; }),
        fuenteRetiros: p.fuenteRetiros || 'carpeta',
        tieneSubprocesos: !!p.tieneSubprocesos
      };
    });
    d.datos.sort(function (a, b) { return a.periodo.localeCompare(b.periodo) || empresas.indexOf(a.empresa) - empresas.indexOf(b.empresa); });
    d.subprocesos.datos.sort(function (a, b) { return a.periodo.localeCompare(b.periodo) || SUBPROCESOS.indexOf(a.subproceso) - SUBPROCESOS.indexOf(b.subproceso); });
    d.version = 3;
    d.umbralDias = UMBRAL_DIAS;
    d.generado = new Date().toISOString();

    var cambios = Object.keys(tocados).sort().map(function (p) {
      return { periodo: p, nombre: MESES[Number(p.slice(5, 7)) - 1] + ' ' + p.slice(0, 4), antes: totalesPeriodo(base, p), despues: totalesPeriodo(d, p) };
    });
    return { data: d, cambios: cambios, historico: historico.sort(), avisos: avisos.concat(r.avisos), errores: errores };
  }

  var api = { procesar: procesar, fusionar: fusionar, MESES: MESES, UMBRAL_DIAS: UMBRAL_DIAS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ProcesarSary = api;
})(this);
