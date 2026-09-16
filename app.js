(function () {
  "use strict";

  var listaArchivos = [];
  var convertidos = [];

  function $(id) {
    return document.getElementById(id);
  }

  function esExcel(nombre) {
    return /\.(xlsx|xls|xlsm)$/i.test(nombre);
  }

  function celdaValor(cell) {
    if (!cell) return null;
    if (cell.t === "d" && cell.v instanceof Date) return cell.v;
    if (cell.v == null || cell.v === "") return null;
    return cell.v;
  }

  function rangoHoja(sheet) {
    var maxR = 0;
    var maxC = 0;
    Object.keys(sheet).forEach(function (key) {
      if (key.charAt(0) === "!") return;
      var dec = XLSX.utils.decode_cell(key);
      if (dec.r > maxR) maxR = dec.r;
      if (dec.c > maxC) maxC = dec.c;
    });
    if (sheet["!ref"]) {
      var ref = XLSX.utils.decode_range(sheet["!ref"]);
      maxR = Math.max(maxR, ref.e.r);
      maxC = Math.max(maxC, ref.e.c);
    }
    return { r: maxR, c: maxC };
  }

  function hojaAFilas(sheet) {
    var lim = rangoHoja(sheet);
    var filas = [];
    var r, c, fila;
    for (r = 0; r <= lim.r; r++) {
      fila = [];
      for (c = 0; c <= lim.c; c++) {
        fila.push(celdaValor(sheet[XLSX.utils.encode_cell({ r: r, c: c })]));
      }
      filas.push(fila);
    }
    return filas;
  }

  function leerLibro(buffer) {
    var wb = XLSX.read(buffer, { type: "array", cellDates: true, raw: true });
    var hojas = {};
    wb.SheetNames.forEach(function (nombre) {
      hojas[nombre] = hojaAFilas(wb.Sheets[nombre]);
    });
    return hojas;
  }

  function libroSkualo(resultado) {
    var aoa = CartolaSkualo.filasSkualo(resultado);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var range = XLSX.utils.decode_range(ws["!ref"]);
    var r, addr;
    for (r = 1; r <= range.e.r; r++) {
      addr = XLSX.utils.encode_cell({ r: r, c: 0 });
      if (ws[addr]) {
        ws[addr].t = "d";
        ws[addr].z = "DD-MM-YYYY";
      }
      [1, 2].forEach(function (c) {
        var a = XLSX.utils.encode_cell({ r: r, c: c });
        if (ws[a]) ws[a].z = "#,##0.00";
      });
    }
    ws["!cols"] = [
      { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 55 }, { wch: 14 },
    ];
    ws["!autofilter"] = { ref: ws["!ref"] };
    ws["!freeze"] = { xSplit: 0, ySplit: 1 };
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Importar_Cartola");
    return wb;
  }

  function fmtMiles(n) {
    return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  }

  function escapeHtml(texto) {
    return String(texto)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function bancoDeNombre(nombre) {
    var n = String(nombre).toLowerCase();
    if (n.indexOf("itau") >= 0 || n.indexOf("itaú") >= 0) return "itau";
    if (n.indexOf("estado") >= 0) return "estado";
    if (n.indexOf("santander") >= 0) return "santander";
    if (n.indexOf("bci") >= 0) return "bci";
    return "";
  }

  function htmlLogo(codigo) {
    var src = {
      bci: "img/bci.svg",
      estado: "img/bancoestado.svg",
      itau: "img/itau.svg",
      santander: "img/santander.svg",
    }[codigo];
    if (!src) return "";
    var extra = codigo === "itau" ? " itau" : "";
    return '<img class="logo-mini' + extra + '" src="' + src + '" alt="">';
  }

  function mostrar(texto) {
    $("resultado").textContent = texto;
  }

  function refrescarListaOrigen() {
    var ul = $("lista-origen");
    ul.innerHTML = "";
    if (!listaArchivos.length) {
      ul.innerHTML = "<li class='vacio'>Ningún Excel todavía.</li>";
      return;
    }
    listaArchivos.forEach(function (item, i) {
      var li = document.createElement("li");
      var banco = bancoDeNombre(item.file.name);
      li.innerHTML =
        '<div class="item-archivo">' + htmlLogo(banco) +
        "<span>" + escapeHtml(item.file.name) + "</span></div>";
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn-texto";
      btn.textContent = "Quitar";
      btn.addEventListener("click", function () {
        listaArchivos.splice(i, 1);
        refrescarListaOrigen();
      });
      li.appendChild(btn);
      ul.appendChild(li);
    });
  }

  function refrescarConvertidos() {
    var ul = $("lista-convertidos");
    ul.innerHTML = "";
    if (!convertidos.length) {
      ul.innerHTML = "<li class='vacio'>Todavía no hay archivos convertidos. Quedan en este navegador, no se envían a internet.</li>";
      return;
    }
    convertidos.forEach(function (item, i) {
      var li = document.createElement("li");
      var alerta = item.cuadrado ? "" : '<span class="badge-alerta">No cuadra</span>';
      li.innerHTML =
        '<div class="item-archivo">' + htmlLogo(item.banco) + alerta +
        "<span>" + escapeHtml(item.nombre) + "</span></div>";
      var acciones = document.createElement("span");
      acciones.className = "acciones";
      var bAbrir = document.createElement("button");
      bAbrir.type = "button";
      bAbrir.className = "btn-texto";
      bAbrir.textContent = "Descargar";
      bAbrir.addEventListener("click", function () { descargarItem(item); });
      var bDel = document.createElement("button");
      bDel.type = "button";
      bDel.className = "btn-texto peligro";
      bDel.textContent = "Eliminar";
      bDel.addEventListener("click", function () {
        if (!confirm("¿Eliminar de Convertidos?\n\n" + item.nombre)) return;
        convertidos.splice(i, 1);
        refrescarConvertidos();
      });
      acciones.appendChild(bAbrir);
      acciones.appendChild(bDel);
      li.appendChild(acciones);
      ul.appendChild(li);
    });
  }

  function descargarItem(item) {
    var blob = new Blob([item.bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = item.nombre;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 500);
  }

  function agregarFiles(fileList) {
    var agregados = 0;
    Array.prototype.forEach.call(fileList || [], function (file) {
      if (!esExcel(file.name)) return;
      var ya = listaArchivos.some(function (x) {
        return x.file.name === file.name && x.file.size === file.size;
      });
      if (ya) return;
      listaArchivos.push({ file: file });
      agregados += 1;
    });
    refrescarListaOrigen();
    if (agregados) mostrar("Archivos listos. Pulse Convertir.");
    else if (fileList && fileList.length) mostrar("Solo se aceptan Excel del banco (.xlsx o .xls).");
  }

  function convertirUno(file) {
    return file.arrayBuffer().then(function (buf) {
      var hojas = leerLibro(buf);
      var resultado = CartolaSkualo.convertir(file.name, hojas);
      var wb = libroSkualo(resultado);
      var bytes = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      convertidos.unshift({
        nombre: resultado.nombreSalida,
        bytes: bytes,
        cuadrado: resultado.cuadrado,
        banco: resultado.banco,
      });
      var lineas = [
        "Archivo: " + file.name,
        "Banco: " + resultado.bancoNombre + "  ·  " + resultado.movimientos.length + " movimientos",
      ];
      if (resultado.cuenta) lineas.push("Cuenta detectada: " + resultado.cuenta + " (confírmela al importar en Skualo)");
      lineas.push("Cargos: " + fmtMiles(resultado.sumaCargos) + "  ·  Abonos: " + fmtMiles(resultado.sumaAbonos));
      lineas.push("Listo: " + resultado.nombreSalida);
      resultado.avisos.forEach(function (a) { lineas.push(a); });
      if (!resultado.cuadrado) {
        lineas.push("Hay un archivo que no cuadra. No lo importe en Skualo.");
      }
      return lineas.join("\n");
    });
  }

  function convertir() {
    if (!listaArchivos.length) {
      alert("Primero agregue el Excel que bajó del banco.");
      return;
    }
    $("btn-convertir").disabled = true;
    mostrar("Convirtiendo…");
    var cola = listaArchivos.slice();
    var partes = [];
    var cadena = Promise.resolve();
    cola.forEach(function (item) {
      cadena = cadena.then(function () {
        return convertirUno(item.file).then(
          function (texto) { partes.push(texto); },
          function (err) { partes.push("Archivo: " + item.file.name + "\nNo se pudo convertir.\n" + (err && err.message ? err.message : err)); }
        );
      });
    });
    cadena.then(function () {
      partes.push("El Excel listo quedó en Convertidos, a la derecha. Pulse Descargar e impórtelo en Tesorería → Cartolas.");
      mostrar(partes.join("\n\n"));
      refrescarConvertidos();
      $("btn-convertir").disabled = false;
    });
  }

  function iniciar() {
    var zona = $("zona");
    var input = $("input-archivos");
    $("btn-elegir").addEventListener("click", function () { input.click(); });
    input.addEventListener("change", function () {
      agregarFiles(input.files);
      input.value = "";
    });
    zona.addEventListener("dragover", function (ev) {
      ev.preventDefault();
      zona.classList.add("sobre");
    });
    zona.addEventListener("dragleave", function () { zona.classList.remove("sobre"); });
    zona.addEventListener("drop", function (ev) {
      ev.preventDefault();
      zona.classList.remove("sobre");
      agregarFiles(ev.dataTransfer.files);
    });
    document.addEventListener("paste", function (ev) {
      if (!ev.clipboardData || !ev.clipboardData.files) return;
      if (!ev.clipboardData.files.length) return;
      agregarFiles(ev.clipboardData.files);
    });
    $("btn-convertir").addEventListener("click", convertir);
    $("btn-vaciar").addEventListener("click", function () {
      if (!convertidos.length) return;
      if (!confirm("¿Eliminar todos los archivos de Convertidos?")) return;
      convertidos = [];
      refrescarConvertidos();
    });
    refrescarListaOrigen();
    refrescarConvertidos();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", iniciar);
  else iniciar();
})();
