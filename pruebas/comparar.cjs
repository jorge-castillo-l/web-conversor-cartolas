/* Compara el motor web con el conversor de escritorio. No forma parte de la Page. */
const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const XLSX = require("../vendor/xlsx.full.min.js");
const CartolaSkualo = require("../motor.js");

function celdaValor(cell) {
  if (!cell) return null;
  if (cell.t === "d" && cell.v instanceof Date) return cell.v;
  if (cell.v == null || cell.v === "") return null;
  return cell.v;
}

function rangoHoja(sheet) {
  let maxR = 0;
  let maxC = 0;
  Object.keys(sheet).forEach((key) => {
    if (key.charAt(0) === "!") return;
    const dec = XLSX.utils.decode_cell(key);
    if (dec.r > maxR) maxR = dec.r;
    if (dec.c > maxC) maxC = dec.c;
  });
  if (sheet["!ref"]) {
    const ref = XLSX.utils.decode_range(sheet["!ref"]);
    maxR = Math.max(maxR, ref.e.r);
    maxC = Math.max(maxC, ref.e.c);
  }
  return { r: maxR, c: maxC };
}

function hojaAFilas(sheet) {
  const lim = rangoHoja(sheet);
  const filas = [];
  for (let r = 0; r <= lim.r; r++) {
    const fila = [];
    for (let c = 0; c <= lim.c; c++) {
      fila.push(celdaValor(sheet[XLSX.utils.encode_cell({ r, c })]));
    }
    filas.push(fila);
  }
  return filas;
}

function convertirWeb(filePath) {
  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: "buffer", cellDates: true, raw: true });
  const hojas = {};
  wb.SheetNames.forEach((nombre) => {
    hojas[nombre] = hojaAFilas(wb.Sheets[nombre]);
  });
  return CartolaSkualo.convertir(path.basename(filePath), hojas);
}

function pythonInfo(filePath) {
  const src = path.resolve(__dirname, "..", "..", "herramientas", "conversor-cartolas-fuente");
  const code = `
import json, sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
import conversor
r = conversor.convertir_archivo(Path(sys.argv[2]))
ok, _ = conversor.cuadre_ok(r)
print(json.dumps({
  "banco": r.banco,
  "n": len(r.movimientos),
  "cargos": r.suma_cargos,
  "abonos": r.suma_abonos,
  "cuenta": r.cuenta,
  "cuadrado": ok,
}))
`;
  const out = spawnSync("python", ["-c", code, src, filePath], { encoding: "utf8" });
  if (out.status !== 0) {
    throw new Error(out.stderr || out.stdout || "python falló");
  }
  const line = out.stdout.trim().split("\n").filter(Boolean).pop();
  return JSON.parse(line);
}

const roots = [
  path.resolve(__dirname, "..", "..", "cartolas-enero"),
  path.resolve(__dirname, "..", "..", "cartolas-faltantes"),
];
const files = [];
for (const root of roots) {
  if (!fs.existsSync(root)) continue;
  for (const name of fs.readdirSync(root)) {
    if (/\.(xlsx|xls)$/i.test(name) && !name.startsWith("~")) {
      files.push(path.join(root, name));
    }
  }
}

let fallos = 0;
for (const file of files) {
  const etiqueta = path.relative(path.resolve(__dirname, "..", ".."), file);
  try {
    const web = convertirWeb(file);
    const py = pythonInfo(file);
    const mismos =
      web.banco === py.banco &&
      web.movimientos.length === py.n &&
      Math.abs(web.sumaCargos - py.cargos) < 1 &&
      Math.abs(web.sumaAbonos - py.abonos) < 1;
    if (!mismos) {
      fallos += 1;
      console.log("DIFF", etiqueta, {
        web: { banco: web.banco, n: web.movimientos.length, cargos: web.sumaCargos, abonos: web.sumaAbonos, cuenta: web.cuenta },
        py,
      });
    } else {
      console.log("OK  ", etiqueta, web.banco, web.movimientos.length, "cta", web.cuenta || "-");
    }
  } catch (err) {
    fallos += 1;
    console.log("ERR ", etiqueta, err.message);
  }
}
console.log(fallos ? `Fallos: ${fallos}` : `Listo: ${files.length} archivos coinciden`);
process.exit(fallos ? 1 : 0);
