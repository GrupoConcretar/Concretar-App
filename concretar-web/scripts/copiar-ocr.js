// Copia a public/tesseract lo que necesita el lector de DNI (motor de OCR + idioma
// español), así la app lo sirve ella misma en lugar de bajarlo de un CDN externo.
import { copyFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const destino = join(process.cwd(), "public", "tesseract");
mkdirSync(destino, { recursive: true });

const desde = (paquete, archivo) => join(dirname(require.resolve(`${paquete}/package.json`)), archivo);
const archivos = [
  [desde("tesseract.js", "dist/worker.min.js"), "worker.min.js"],
  ...["tesseract-core-lstm.wasm.js", "tesseract-core-simd-lstm.wasm.js", "tesseract-core-relaxedsimd-lstm.wasm.js"]
    .map((f) => [desde("tesseract.js-core", f), f]),
  [desde("@tesseract.js-data/spa", "4.0.0_best_int/spa.traineddata.gz"), "spa.traineddata.gz"],
];
for (const [origen, nombre] of archivos) copyFileSync(origen, join(destino, nombre));
console.log(`Lector de DNI: ${archivos.length} archivos copiados a public/tesseract`);
