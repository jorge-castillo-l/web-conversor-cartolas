/* Conversor de cartolas bancarias chilenas al Excel de importación de Skualo.
   Sin datos de una empresa: el RUT de contraparte queda vacío; la cuenta se elige en Skualo. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CartolaSkualo = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var ENCABEZADOS = ["Fecha", "MontoCargo", "MontoAbono", "NumDoc", "Glosa", "RUT"];
  var NOMBRES_BANCO = {
    bci: "BCI",
    estado: "Banco Estado",
    itau: "Itaú",
    santander: "Santander",
  };

  function norm(value) {
    var text = value == null ? "" : String(value);
    text = text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    text = text.replace(/[°º]/g, " ").replace(/nº/gi, "n ").replace(/n°/gi, "n ");
    return text.toLowerCase().replace(/º/g, " ").replace(/\s+/g, " ").trim();
  }

  function parseMonto(value) {
    if (value == null || value === "") return 0;
    if (typeof value === "boolean") return 0;
    if (typeof value === "number") return Math.abs(value);
    if (value instanceof Date) return 0;
    var text = String(value)
      .trim()
      .replace(/\$/g, "")
      .replace(/ /g, "")
      .replace(/\u00a0/g, "")
      .replace(/−/g, "-");
    text = text.replace(/^-+|-+$/g, "");
    if (!text) return 0;
    if (text.indexOf(",") >= 0 && text.indexOf(".") >= 0) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else if ((text.match(/\./g) || []).length > 1) {
      text = text.replace(/\./g, "");
    } else if ((text.match(/\./g) || []).length === 1) {
      var partes = text.split(".");
      if (/^\d{3}$/.test(partes[1]) && /^-?\d+$/.test(partes[0])) text = text.replace(/\./g, "");
    } else if (text.indexOf(",") >= 0) {
      var c = text.split(",");
      text = /^\d{3}$/.test(c[1]) ? text.replace(/,/g, "") : text.replace(",", ".");
    }
    var n = parseFloat(text);
    return isFinite(n) ? Math.abs(n) : 0;
  }

  function parseNumdoc(value) {
    if (value == null || value === "") return 0;
    if (typeof value === "number") return Math.trunc(value);
    var digits = String(value).replace(/\D/g, "");
    return digits ? parseInt(digits, 10) : 0;
  }

  function ymd(y, m, d) {
    if (m < 1 || m > 12 || d < 1 || d > 31) return null;
    var dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return y + "-" + String(m).padStart(2, "0") + "-" + String(d).padStart(2, "0");
  }

  function desdeFechaJs(value) {
    if (!(value instanceof Date) || isNaN(value.getTime())) return null;
    return ymd(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate());
  }

  function parseFechaFlexible(value) {
    if (value == null || value === "") return null;
    var js = desdeFechaJs(value);
    if (js) return js;
    var text = String(value).trim();
    var formatos = [
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
      /^(\d{1,2})-(\d{1,2})-(\d{4})$/,
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/,
    ];
    var m = text.match(formatos[0]) || text.match(formatos[1]);
    if (m) return ymd(Number(m[3]), Number(m[2]), Number(m[1]));
    m = text.match(formatos[2]);
    if (m) return ymd(Number(m[1]), Number(m[2]), Number(m[3]));
    return null;
  }

  function anioDe(iso) {
    return Number(iso.slice(0, 4));
  }

  function parseFecha(value, periodoInicio, periodoFin) {
    if (value == null || value === "") return null;
    var js = desdeFechaJs(value);
    if (js) return js;
    var flex = parseFechaFlexible(value);
    if (flex) return flex;
    var text = String(value).trim();
    var corto = text.match(/^(\d{1,2})[/-](\d{1,2})$/);
    if (corto && periodoInicio && periodoFin) {
      var dia = Number(corto[1]);
      var mes = Number(corto[2]);
      var years = [anioDe(periodoFin), anioDe(periodoInicio)];
      var i;
      for (i = 0; i < years.length; i++) {
        var cand = ymd(years[i], mes, dia);
        if (!cand) continue;
        var t = Date.parse(cand + "T00:00:00Z");
        var ini = Date.parse(periodoInicio + "T00:00:00Z") - 3 * 86400000;
        var fin = Date.parse(periodoFin + "T00:00:00Z") + 3 * 86400000;
        if (t >= ini && t <= fin) return cand;
      }
      return ymd(anioDe(periodoFin), mes, dia);
    }
    var cortoAnio = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2})$/);
    if (cortoAnio) {
      var yy = Number(cortoAnio[3]);
      return ymd(yy < 70 ? 2000 + yy : 1900 + yy, Number(cortoAnio[2]), Number(cortoAnio[1]));
    }
    return null;
  }

  function extraerPeriodo(filas) {
    var texto = [];
    var r, c, fila;
    for (r = 0; r < Math.min(40, filas.length); r++) {
      fila = filas[r] || [];
      for (c = 0; c < fila.length; c++) texto.push(fila[c] == null ? "" : String(fila[c]));
    }
    var textoN = norm(texto.join(" | "));
    var reRango = /(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s*(?:al|-|a)\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})/g;
    var hit;
    while ((hit = reRango.exec(textoN))) {
      var iniR = parseFechaFlexible(hit[1]);
      var finR = parseFechaFlexible(hit[2]);
      if (iniR && finR) return [iniR, finR];
    }
    var ini = null;
    var fin = null;
    for (r = 0; r < Math.min(25, filas.length); r++) {
      fila = filas[r] || [];
      var etiqueta = norm(fila[0]);
      var valor = null;
      for (c = 1; c < Math.min(6, fila.length); c++) {
        var f = parseFechaFlexible(fila[c]);
        if (f) {
          valor = f;
          break;
        }
      }
      if (etiqueta.indexOf("fecha inicio") >= 0 || etiqueta.indexOf("fecha desde") >= 0) ini = valor;
      else if (etiqueta.indexOf("fecha final") >= 0 || etiqueta.indexOf("fecha hasta") >= 0) fin = valor;
      for (c = 0; c < fila.length; c++) {
        var n = norm(fila[c]);
        var mDesde = n.match(/fecha desde[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})/);
        var mHasta = n.match(/fecha hasta[:\s]+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})/);
        if (mDesde) ini = parseFechaFlexible(mDesde[1]);
        if (mHasta) fin = parseFechaFlexible(mHasta[1]);
      }
    }
    return [ini, fin];
  }

  function digitosCuenta(valor) {
    if (valor == null) return "";
    if (valor instanceof Date) return "";
    var texto = String(valor).trim();
    if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/.test(texto)) return "";
    var digits = texto.replace(/\D/g, "");
    if (digits.length >= 6 && digits.length <= 12) return digits;
    return "";
  }

  function extraerCuenta(filas) {
    var r, i, fila, n, mismo, cand, vecino;
    for (r = 0; r < Math.min(25, filas.length); r++) {
      fila = filas[r] || [];
      for (i = 0; i < fila.length; i++) {
        n = norm(fila[i]);
        if (n.indexOf("cuenta") < 0) continue;
        mismo = digitosCuenta(fila[i]);
        if (mismo) return mismo;
        for (cand = i + 1; cand < Math.min(i + 5, fila.length); cand++) {
          vecino = digitosCuenta(fila[cand]);
          if (vecino) return vecino;
        }
      }
    }
    return "";
  }

  function filaEncabezado(filas, pistas) {
    var pistasN = pistas.map(norm);
    var i, fila, celdas;
    for (i = 0; i < filas.length; i++) {
      fila = filas[i] || [];
      celdas = fila.map(norm);
      if (pistasN.every(function (p) { return celdas.some(function (c) { return c.indexOf(p) >= 0; }); })) {
        return i;
      }
    }
    return null;
  }

  function indiceCol(fila, pistas) {
    var celdas = (fila || []).map(norm);
    var p, i, c;
    for (p = 0; p < pistas.length; p++) {
      var pista = norm(pistas[p]);
      for (i = 0; i < celdas.length; i++) {
        c = celdas[i];
        if (c === pista || c.indexOf(pista) >= 0) return i;
      }
    }
    return null;
  }

  function celda(fila, idx) {
    if (idx == null || idx < 0 || idx >= fila.length) return null;
    return fila[idx];
  }

  function parsearBci(filas) {
    var i = filaEncabezado(filas, ["fecha", "cheques y otros cargos"]);
    if (i == null) throw new Error("BCI: no encuentro la fila de encabezados de movimientos.");
    var cabeza = filas[i];
    var cFecha = indiceCol(cabeza, ["fecha"]);
    var cGlosa = indiceCol(cabeza, ["descripcion"]);
    var cDoc = indiceCol(cabeza, ["n documento", "documento"]);
    var cCargo = indiceCol(cabeza, ["cheques y otros cargos"]);
    var cAbono = indiceCol(cabeza, ["depositos y abono", "depositos y abonos"]);
    if ([cFecha, cGlosa, cDoc, cCargo, cAbono].indexOf(null) >= 0) {
      throw new Error("BCI: faltan columnas de fecha, glosa, documento, cargo o abono.");
    }
    var per = extraerPeriodo(filas);
    var movs = [];
    var fila, fecha, glosa, cargo, abono;
    for (var r = i + 1; r < filas.length; r++) {
      fila = filas[r] || [];
      fecha = parseFecha(celda(fila, cFecha), per[0], per[1]);
      if (!fecha) continue;
      glosa = String(celda(fila, cGlosa) || "").trim();
      if (!glosa) continue;
      cargo = parseMonto(celda(fila, cCargo));
      abono = parseMonto(celda(fila, cAbono));
      if (cargo === 0 && abono === 0) continue;
      movs.push({ fecha: fecha, cargo: cargo, abono: abono, numdoc: parseNumdoc(celda(fila, cDoc)), glosa: glosa, rut: "" });
    }
    var totC = null;
    var totA = null;
    for (r = 0; r < Math.min(20, filas.length); r++) {
      fila = filas[r] || [];
      if (fila.length > 10 && fila[0] && /\d{2}[-/]\d{2}[-/]\d{4}/.test(String(fila[0]))) {
        if (fila[6] != null && fila[6] !== "" && fila[8] != null && fila[8] !== "") {
          totC = parseMonto(fila[6]);
          totA = parseMonto(fila[8]);
        }
      }
    }
    return [movs, totC, totA];
  }

  function parsearEstado(hojas) {
    var names = Object.keys(hojas);
    var movKey = names.filter(function (k) { return norm(k) === "movimientos"; })[0];
    var resKey = names.filter(function (k) { return norm(k) === "resumen"; })[0];
    if (!movKey) throw new Error("Banco Estado: no hay hoja Movimientos.");
    var filas = hojas[movKey];
    if (!filas || !filas.length) throw new Error("Banco Estado: la hoja Movimientos está vacía.");
    var cabeza = filas[0];
    var cFecha = indiceCol(cabeza, ["fecha"]);
    var cGlosa = indiceCol(cabeza, ["descripcion"]);
    var cDoc = indiceCol(cabeza, ["n operacion", "operacion"]);
    var cCargo = indiceCol(cabeza, ["cheques / cargos", "cargos"]);
    var cAbono = indiceCol(cabeza, ["depositos / abonos", "abonos"]);
    if ([cFecha, cGlosa, cDoc, cCargo, cAbono].indexOf(null) >= 0) {
      throw new Error("Banco Estado: faltan columnas en Movimientos.");
    }
    var per = extraerPeriodo(resKey ? hojas[resKey] : filas);
    var movs = [];
    var r, fila, fecha, glosa;
    for (r = 1; r < filas.length; r++) {
      fila = filas[r] || [];
      fecha = parseFecha(celda(fila, cFecha), per[0], per[1]);
      if (!fecha) continue;
      glosa = String(celda(fila, cGlosa) || "").trim();
      movs.push({
        fecha: fecha,
        cargo: parseMonto(celda(fila, cCargo)),
        abono: parseMonto(celda(fila, cAbono)),
        numdoc: parseNumdoc(celda(fila, cDoc)),
        glosa: glosa,
        rut: "",
      });
    }
    var totC = null;
    var totA = null;
    if (resKey) {
      var res = hojas[resKey];
      for (r = 0; r < res.length; r++) {
        fila = res[r] || [];
        var etiqueta = norm(fila[0]);
        var valor = fila.length > 4 ? fila[4] : null;
        if (etiqueta === "total cargos") totC = parseMonto(valor);
        else if (etiqueta === "total abonos") totA = parseMonto(valor);
      }
    }
    return [movs, totC, totA];
  }

  function parsearSantander(filas) {
    var i = filaEncabezado(filas, ["monto", "fecha", "cargo/abono"]);
    if (i == null) throw new Error("Santander: no encuentro el detalle de movimientos.");
    var cabeza = filas[i];
    var cMonto = indiceCol(cabeza, ["monto"]);
    var cGlosa = indiceCol(cabeza, ["descripcion movimiento", "descripcion"]);
    var cFecha = indiceCol(cabeza, ["fecha"]);
    var cDoc = indiceCol(cabeza, ["n documento", "documento"]);
    var cTipo = indiceCol(cabeza, ["cargo/abono"]);
    if ([cMonto, cGlosa, cFecha, cDoc, cTipo].indexOf(null) >= 0) {
      throw new Error("Santander: faltan columnas del detalle.");
    }
    var per = extraerPeriodo(filas);
    var movs = [];
    var r, fila, primera, fecha, glosa, tipo, monto, cargo, abono;
    for (r = i + 1; r < filas.length; r++) {
      fila = filas[r] || [];
      primera = norm(fila[0]);
      if (primera.indexOf("resumen") === 0 || primera.indexOf("saldos") === 0) break;
      if (primera === "monto" || primera === "saldo") break;
      fecha = parseFecha(celda(fila, cFecha), per[0], per[1]);
      if (!fecha) continue;
      glosa = String(celda(fila, cGlosa) || "").trim();
      if (!glosa) continue;
      tipo = norm(celda(fila, cTipo));
      monto = parseMonto(celda(fila, cMonto));
      cargo = tipo.indexOf("c") === 0 ? monto : 0;
      abono = tipo.indexOf("a") === 0 ? monto : 0;
      if (cargo === 0 && abono === 0) continue;
      movs.push({ fecha: fecha, cargo: cargo, abono: abono, numdoc: parseNumdoc(celda(fila, cDoc)), glosa: glosa, rut: "" });
    }
    var totC = null;
    var totA = null;
    for (r = 0; r < filas.length; r++) {
      fila = filas[r] || [];
      var celdas = fila.map(norm);
      if (celdas.indexOf("saldo inicial") >= 0 && celdas.indexOf("otros abonos") >= 0 && r + 1 < filas.length) {
        var vals = filas[r + 1] || [];
        totA = parseMonto(vals[1]) + parseMonto(vals[2]);
        totC = parseMonto(vals[3]) + parseMonto(vals[4]) + parseMonto(vals[5]);
        break;
      }
    }
    return [movs, totC, totA];
  }

  function parsearItau(filas) {
    var i = null;
    var idx, celdas;
    for (idx = 0; idx < filas.length; idx++) {
      celdas = (filas[idx] || []).map(norm);
      if (celdas.indexOf("fecha") >= 0 && celdas.some(function (c) { return c.indexOf("deposito") >= 0; })) {
        i = idx;
        break;
      }
    }
    if (i == null) throw new Error("Itaú: no encuentro la tabla de movimientos.");
    var cabeza = (filas[i] || []).slice();
    var extra;
    if (i + 1 < filas.length) {
      extra = filas[i + 1] || [];
      for (var j = 0; j < extra.length; j++) {
        if (extra[j]) {
          if (j < cabeza.length) cabeza[j] = String(cabeza[j] || "").trim() + " " + extra[j];
          else cabeza.push(extra[j]);
        }
      }
    }
    var cFecha = indiceCol(cabeza, ["fecha"]);
    var cDoc = indiceCol(cabeza, ["numero de operacion", "numero de"]);
    var cGlosa = indiceCol(cabeza, ["descripcion"]);
    var cAbono = indiceCol(cabeza, ["depositos o abonos", "depositos"]);
    var cCargo = indiceCol(cabeza, ["giros o cargos", "giros"]);
    if ([cFecha, cGlosa, cAbono, cCargo].indexOf(null) >= 0) {
      throw new Error("Itaú: faltan columnas de movimientos.");
    }
    var per = extraerPeriodo(filas);
    var start = i + 1;
    if (i + 1 < filas.length && norm((filas[i + 1] || [])[0]) === "") start = i + 2;
    var filaSig = filas[i + 1] || [];
    if (i + 1 < filas.length && norm(filaSig[1] || "").indexOf("operacion") >= 0) start = i + 2;
    var movs = [];
    var r, fila, fecha, glosa;
    for (r = start; r < filas.length; r++) {
      fila = filas[r] || [];
      if (fila[0] && norm(fila[0]).indexOf("resumen") === 0) break;
      fecha = parseFecha(celda(fila, cFecha), per[0], per[1]);
      if (!fecha) {
        if (movs.length) break;
        continue;
      }
      glosa = String(celda(fila, cGlosa) || "").trim();
      if (!glosa) continue;
      movs.push({
        fecha: fecha,
        cargo: parseMonto(celda(fila, cCargo)),
        abono: parseMonto(celda(fila, cAbono)),
        numdoc: parseNumdoc(cDoc == null ? 0 : celda(fila, cDoc)),
        glosa: glosa,
        rut: "",
      });
    }
    var totC = null;
    var totA = null;
    for (idx = 0; idx < filas.length; idx++) {
      celdas = (filas[idx] || []).map(norm);
      if (celdas.indexOf("total cargos") >= 0 && celdas.indexOf("total abono") >= 0 && idx + 1 < filas.length) {
        totC = parseMonto((filas[idx + 1] || [])[0]);
        totA = parseMonto((filas[idx + 1] || [])[1]);
        break;
      }
    }
    return [movs, totC, totA];
  }

  function textoMuestra(hojas) {
    var partes = Object.keys(hojas);
    Object.keys(hojas).forEach(function (k) {
      var filas = hojas[k] || [];
      var r, c, fila;
      for (r = 0; r < Math.min(30, filas.length); r++) {
        fila = filas[r] || [];
        for (c = 0; c < fila.length; c++) partes.push(fila[c] == null ? "" : String(fila[c]));
      }
    });
    return partes.join(" ");
  }

  function detectarBanco(nombreArchivo, hojas) {
    var n = norm(nombreArchivo + " " + textoMuestra(hojas));
    var nombres = Object.keys(hojas).map(norm);
    if (n.indexOf("itau") >= 0) return "itau";
    if (n.indexOf("bancoestado") >= 0 || n.indexOf("banco estado") >= 0 ||
        (nombres.indexOf("resumen") >= 0 && nombres.indexOf("movimientos") >= 0)) {
      return "estado";
    }
    if (n.indexOf("santander") >= 0 || n.indexOf("cargo/abono") >= 0 || n.indexOf("cartolas historicas de cuentas") >= 0) {
      return "santander";
    }
    if (n.indexOf("bci") >= 0 || n.indexOf("cheques y otros cargos") >= 0) return "bci";
    throw new Error(
      "No reconozco el banco de " + nombreArchivo + ". " +
      "Este conversor cubre BCI, Banco Estado, Itaú y Santander."
    );
  }

  function primeraHojaUtil(hojas) {
    var keys = Object.keys(hojas);
    var k;
    for (k = 0; k < keys.length; k++) {
      var filas = hojas[keys[k]] || [];
      var hay = filas.some(function (fila) {
        return (fila || []).some(function (c) { return c != null && String(c).trim() !== ""; });
      });
      if (hay) return hojas[keys[k]];
    }
    return hojas[keys[0]] || [];
  }

  function sumaCampo(movs, campo) {
    return Math.round(movs.reduce(function (a, m) { return a + m[campo]; }, 0) * 100) / 100;
  }

  function fmtEntero(n) {
    var s = String(Math.round(n));
    var neg = s.charAt(0) === "-";
    if (neg) s = s.slice(1);
    var out = s.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return neg ? "-" + out : out;
  }

  function cuadreOk(sumaCargos, sumaAbonos, totC, totA, tolerancia) {
    if (tolerancia == null) tolerancia = 1;
    var msgs = [];
    var ok = true;
    if (totC != null) {
      var dC = Math.abs(sumaCargos - totC);
      if (dC > tolerancia) {
        ok = false;
        msgs.push("Cargos no cuadran: archivo " + fmtEntero(sumaCargos) + " vs cartola " + fmtEntero(totC) + " (dif " + fmtEntero(dC) + ")");
      }
    }
    if (totA != null) {
      var dA = Math.abs(sumaAbonos - totA);
      if (dA > tolerancia) {
        ok = false;
        msgs.push("Abonos no cuadran: archivo " + fmtEntero(sumaAbonos) + " vs cartola " + fmtEntero(totA) + " (dif " + fmtEntero(dA) + ")");
      }
    }
    return { ok: ok, mensajes: msgs };
  }

  function convertir(nombreArchivo, hojas) {
    var banco = detectarBanco(nombreArchivo, hojas);
    var primera = primeraHojaUtil(hojas);
    var parsed;
    var cuenta = "";
    var filasPeriodo = primera;
    if (banco === "bci") {
      parsed = parsearBci(primera);
      cuenta = extraerCuenta(primera);
    } else if (banco === "estado") {
      parsed = parsearEstado(hojas);
      var resKey = Object.keys(hojas).filter(function (k) { return norm(k) === "resumen"; })[0];
      var res = resKey ? hojas[resKey] : primera;
      cuenta = extraerCuenta(res);
      filasPeriodo = res;
    } else if (banco === "santander") {
      parsed = parsearSantander(primera);
      cuenta = extraerCuenta(primera);
    } else if (banco === "itau") {
      parsed = parsearItau(primera);
      cuenta = extraerCuenta(primera);
    } else {
      throw new Error(banco);
    }
    var movs = parsed[0];
    var totC = parsed[1];
    var totA = parsed[2];
    if (!movs.length) throw new Error("No salió ningún movimiento de " + nombreArchivo + ".");
    var per = extraerPeriodo(filasPeriodo);
    if (!per[0]) per[0] = movs.reduce(function (a, m) { return a < m.fecha ? a : m.fecha; }, movs[0].fecha);
    if (!per[1]) per[1] = movs.reduce(function (a, m) { return a > m.fecha ? a : m.fecha; }, movs[0].fecha);
    var sumaCargos = sumaCampo(movs, "cargo");
    var sumaAbonos = sumaCampo(movs, "abono");
    var cuadre = cuadreOk(sumaCargos, sumaAbonos, totC, totA);
    var stem = nombreArchivo.replace(/\.[^.]+$/, "");
    var nombreSalida = stem + "_formato skualo";
    if (!cuadre.ok) nombreSalida += "_SUMAS_NO_CUADRAN";
    nombreSalida += ".xlsx";
    return {
      banco: banco,
      bancoNombre: NOMBRES_BANCO[banco] || banco,
      cuenta: cuenta,
      movimientos: movs,
      periodoInicio: per[0],
      periodoFin: per[1],
      sumaCargos: sumaCargos,
      sumaAbonos: sumaAbonos,
      totalCargosOrigen: totC,
      totalAbonosOrigen: totA,
      cuadrado: cuadre.ok,
      avisos: cuadre.mensajes.slice(),
      nombreSalida: nombreSalida,
    };
  }

  function filasSkualo(resultado) {
    var filas = [ENCABEZADOS.slice()];
    resultado.movimientos.forEach(function (m) {
      var p = m.fecha.split("-");
      filas.push([
        new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]))),
        m.cargo,
        m.abono,
        m.numdoc,
        m.glosa,
        "",
      ]);
    });
    return filas;
  }

  return {
    ENCABEZADOS: ENCABEZADOS,
    convertir: convertir,
    filasSkualo: filasSkualo,
    norm: norm,
  };
});
