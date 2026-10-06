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
          var t = /^([A-ZÁÉÍÓÚ]+)\s+DE\s+(\d{4})$/.exec(limpiar(v));
          if (t && MESES.indexOf(t[1]) >= 0) titulo = { mes: MESES.indexOf(t[1]) + 1, anio: +t[2], texto: limpiar(v) };
        }
        if (total === null && encabezado(v) === 'TOTAL' && typeof m[i][j + 1] === 'number' && m[i][j + 1] > 20) total = Math.round(m[i][j + 1]);
      }
    }
    return total ? { total: total, titulo: titulo } : null;
  }

  // Lee un archivo de "total de empleados" con bloques por mes:
  // una celda "MES DE AÑO" seguida de una fila "TOTAL | número".
  function leerArchivoPlanta(XLSX, wb) {
    var res = {}, n = 0;
    wb.SheetNames.forEach(function (nombre) {
      var m = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: null });
      var actual = null;
      m.forEach(function (fila) {
        for (var j = 0; j < fila.length; j++) {
          var v = fila[j];
          if (typeof v === 'string') {
            var t = /^([A-ZÁÉÍÓÚ]+)\s+DE\s+(\d{4})$/.exec(limpiar(v));
            if (t && MESES.indexOf(t[1]) >= 0) { actual = idPeriodo(+t[2], MESES.indexOf(t[1]) + 1); return; }
            if (actual && encabezado(v) === 'TOTAL' && typeof fila[j + 1] === 'number' && fila[j + 1] > 20) {
              if (!res[actual]) { res[actual] = Math.round(fila[j + 1]); n++; }
              return;
            }
          }
        }
      });
    });
    return n ? res : null;
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
            var rt = false;
            if (si !== null && sr !== null) {
              var dias = sr - si;
              rt = dias >= 0 && dias <= UMBRAL_DIAS;
              if (dias < 0) avisosArchivo.push(base + ' · fila ' + fila + ' (' + empresa + '): la fecha de retiro es anterior a la de ingreso.');
            } else {
              avisosArchivo.push(base + ' · fila ' + fila + ' (' + empresa + (sr !== null ? ', retiro ' + textoFecha(sr) : '') + '): sin fecha de ingreso, no entra en rotación temprana.');
            }
            filasArchivo.push({ periodo: per, empresa: empresa, motivo: tipoMotivo(cMot >= 0 ? r[cMot] : ''), rt: rt, sub: cSub >= 0 ? normalizarSubproceso(r[cSub]) : null });
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
            archivosPlanta.push({ nombre: f.nombre, fecha: f.fecha || 0, valores: valores });
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

    // El archivo de total de empleados tiene prioridad sobre las hojas "Indicador de rotación"
    if (archivosPlanta.length) {
      archivosPlanta.sort(function (a, b) { return a.fecha - b.fecha || a.nombre.localeCompare(b.nombre); });
      var ap = archivosPlanta[archivosPlanta.length - 1];
      archivosPlanta.slice(0, -1).forEach(function (c) {
        omitidos.push(c.nombre + ' (otra versión del total de empleados; se usó ' + ap.nombre.split('/').pop() + ', la más reciente)');
      });
      Object.keys(ap.valores).forEach(function (per) {
        planta[per] = { valor: ap.valores[per], origen: 'Archivo de total de empleados', aviso: '' };
      });
      usados.push({ archivo: ap.nombre, tipo: 'planta', ingresos: 0, retiros: 0, meses: Object.keys(ap.valores).sort() });
    }

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
      if (!sub[k]) sub[k] = { periodo: r.periodo, subproceso: r.sub, renuncias: 0, terminaciones: 0, otros: 0 };
      sub[k][r.motivo]++;
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
      subprocesos: {
        lista: SUBPROCESOS,
        datos: Object.keys(sub).sort().map(function (k) { return sub[k]; })
      }
    };

    return { data: data, usados: usados, omitidos: omitidos, avisos: avisos, planta: planta };
  }

  var api = { procesar: procesar, MESES: MESES, UMBRAL_DIAS: UMBRAL_DIAS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ProcesarSary = api;
})(this);
