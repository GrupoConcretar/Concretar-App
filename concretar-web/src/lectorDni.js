// Lector de DNI argentino (tarjeta). Lee las fotos con OCR dentro del navegador
// (tesseract.js, servido desde /tesseract): no se manda la imagen a ningún servidor.
//
// La fuente principal es la zona de lectura mecánica del dorso (las 3 líneas con
// "<<<"): trae DNI, apellido, nombres, sexo y fecha de nacimiento con dígitos
// verificadores, así se sabe si se leyó bien. Arriba de esa zona, en el mismo dorso,
// están impresos el domicilio y el CUIL. Del frente solo se toma el número de DNI,
// por si el dorso no se pudo leer.

const ANCHO_OCR = 2400;
const CARACTERES_MRZ = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<";

// ---------- Interpretación del texto (sin navegador: se puede probar suelta) ----------

const valorMrz = (c) => (c === "<" ? 0 : /\d/.test(c) ? Number(c) : c.charCodeAt(0) - 55);
export const digitoMrz = (s) => [...s].reduce((t, c, i) => t + valorMrz(c) * [7, 3, 1][i % 3], 0) % 10;

export function cuilValido(cuil) {
  const d = String(cuil).replace(/\D/g, "");
  if (d.length !== 11) return false;
  const suma = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2].reduce((t, p, i) => t + p * Number(d[i]), 0);
  const resto = 11 - (suma % 11);
  const verificador = resto === 11 ? 0 : resto === 10 ? 9 : resto;
  return verificador === Number(d[10]);
}

export function tituloNombre(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim()
    .replace(/(^|[\s\-(/])(\p{L})/gu, (_, sep, letra) => sep + letra.toUpperCase());
}

// AAMMDD → AAAA-MM-DD. El siglo se elige para que la fecha no quede en el futuro.
function fechaMrz(aammdd) {
  const aa = Number(aammdd.slice(0, 2));
  const siglo = aa > new Date().getFullYear() % 100 ? 1900 : 2000;
  const fecha = `${siglo + aa}-${aammdd.slice(2, 4)}-${aammdd.slice(4, 6)}`;
  return Number.isNaN(Date.parse(fecha)) ? "" : fecha;
}

// Recibe el texto leído (una o varias pasadas) y arma los datos de la zona mecánica.
export function interpretarMrz(texto) {
  const lineas = String(texto || "")
    .toUpperCase()
    .replace(/«/g, "<<")
    .split("\n")
    .map((l) => l.replace(/\s+/g, ""))
    .filter(Boolean);
  const datos = {};

  for (const l of lineas) {
    // Línea 1: IDARG + número de documento (9 posiciones) + dígito verificador.
    const m1 = l.replace(/O/g, "0").match(/ARG([0-9]{7,8})(<?)([0-9])?/);
    if (m1 && !datos.dni) {
      const campo = (m1[1] + "<").slice(0, 9).padEnd(9, "<");
      datos.dni = m1[1].replace(/^0+/, "");
      datos.dniVerificado = m1[3] !== undefined && digitoMrz(campo) === Number(m1[3]);
      continue;
    }
    // Línea 2: nacimiento AAMMDD + dígito, sexo, vencimiento AAMMDD + dígito, ARG.
    const m2 = l.replace(/O/g, "0").match(/(\d{6})(\d)<*([MFX])[MFX]?(\d{6})(\d)/);
    if (m2 && !datos.fechaNacimiento) {
      if (digitoMrz(m2[1]) === Number(m2[2])) datos.fechaNacimiento = fechaMrz(m2[1]);
      datos.sexo = m2[3];
      continue;
    }
    // Línea 3: APELLIDO<APELLIDO<<NOMBRE<NOMBRE (sin números).
    if (!datos.apellido && /<</.test(l) && /^[A-Z<]+$/.test(l) && !/^I[D<]ARG/.test(l) && !/^[<K]+$/.test(l)) {
      const limpio = l.replace(/[<K]{3,}$/, "").replace(/<+$/, "");
      const corte = limpio.indexOf("<<");
      if (corte > 1) {
        datos.apellido = tituloNombre(limpio.slice(0, corte).replace(/<+/g, " "));
        datos.nombre = tituloNombre(limpio.slice(corte + 2).replace(/<+/g, " "));
      }
    }
  }
  return datos;
}

// Busca un CUIL válido en el texto. Si se conoce el DNI, tiene que coincidir.
export function buscarCuil(texto, dni) {
  const candidatos = String(texto || "").replace(/[Oo]/g, "0").matchAll(/(2[0347]|3[034])\s*[-–—.]?\s*(\d{8})\s*[-–—.]?\s*(\d)/g);
  for (const m of candidatos) {
    const cuil = `${m[1]}-${m[2]}-${m[3]}`;
    if (!cuilValido(cuil)) continue;
    if (dni && Number(m[2]) !== Number(dni)) continue;
    return cuil;
  }
  return "";
}

// Busca el domicilio en el texto de la parte de arriba del dorso.
export function buscarDomicilio(texto) {
  const lineas = String(texto || "").split("\n").map((l) => l.trim()).filter(Boolean);
  const conEtiqueta = lineas.map((l) => l.match(/domic\w*\s*[:.]?\s*(.{6,})/i)).find(Boolean);
  let linea = conEtiqueta ? conEtiqueta[1] : lineas.find((l) => /\d/.test(l) && (l.match(/[A-ZÁÉÍÓÚÑ]{2,}/g) || []).length >= 2 && l.length >= 10 && !/CUIL|TRAMITE|TRÁMITE|<</i.test(l));
  if (!linea) return "";
  // Se descartan restos de la etiqueta o de bordes antes de la primera palabra en mayúsculas.
  const desde = linea.search(/[A-ZÁÉÍÓÚÑ0-9]{2,}/);
  linea = (desde >= 0 ? linea.slice(desde) : linea).replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ0-9.)]+$/, "");
  return tituloNombre(linea);
}

// ---------- OCR en el navegador ----------

async function cargarImagen(src) {
  const img = new Image();
  img.src = src;
  await img.decode();
  return img;
}

// Dibuja un recorte de la imagen (opcionalmente girada) en un lienzo de ANCHO_OCR,
// usando el canal rojo (el que mejor separa la tinta del fondo celeste/rosa del DNI)
// y estirando el contraste.
function prepararLienzo(fuente, { rotacion = 0, recorte = null } = {}) {
  const ancho0 = fuente.naturalWidth || fuente.width;
  const alto0 = fuente.naturalHeight || fuente.height;
  const r = recorte || { x: 0, y: 0, w: ancho0, h: alto0 };
  const girada = rotacion === 90 || rotacion === 270;
  const escala = ANCHO_OCR / (girada ? r.h : r.w);
  const lienzo = document.createElement("canvas");
  lienzo.width = Math.round((girada ? r.h : r.w) * escala);
  lienzo.height = Math.round((girada ? r.w : r.h) * escala);
  const ctx = lienzo.getContext("2d", { willReadFrequently: true });
  ctx.translate(lienzo.width / 2, lienzo.height / 2);
  ctx.rotate((rotacion * Math.PI) / 180);
  ctx.drawImage(fuente, r.x, r.y, r.w, r.h, (-r.w * escala) / 2, (-r.h * escala) / 2, r.w * escala, r.h * escala);
  const datos = ctx.getImageData(0, 0, lienzo.width, lienzo.height);
  const px = datos.data;
  const histograma = new Array(256).fill(0);
  for (let i = 0; i < px.length; i += 4) histograma[px[i]]++;
  const total = px.length / 4;
  let acumulado = 0, bajo = 0, alto = 255;
  for (let v = 0; v < 256; v++) { acumulado += histograma[v]; if (acumulado <= total * 0.01) bajo = v; if (acumulado <= total * 0.99) alto = v; }
  const rango = Math.max(1, alto - bajo);
  for (let i = 0; i < px.length; i += 4) {
    const v = Math.max(0, Math.min(255, ((px[i] - bajo) * 255) / rango));
    px[i] = px[i + 1] = px[i + 2] = v;
  }
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.putImageData(datos, 0, 0);
  return lienzo;
}

const lineasDe = (data) => (data.blocks || []).flatMap((b) => b.paragraphs.flatMap((p) => p.lines));

async function crearLector() {
  const { createWorker } = await import("tesseract.js");
  const base = new URL("/tesseract/", window.location.href).href;
  return createWorker("spa", 1, {
    workerPath: `${base}worker.min.js`,
    corePath: base,
    langPath: base.replace(/\/$/, ""),
    gzip: true,
  });
}

async function leerDorso(lector, img) {
  // 1) Se busca la zona mecánica, probando la foto girada por si se sacó de costado.
  let lienzo = null, lineasMrz = [], textoGeneral = "";
  for (const rotacion of [0, 90, 270, 180]) {
    lienzo = prepararLienzo(img, { rotacion });
    const { data } = await lector.recognize(lienzo, {}, { blocks: true, text: true });
    lineasMrz = lineasDe(data).filter((l) => /<</.test(l.text));
    textoGeneral = data.text;
    if (lineasMrz.length >= 2) break;
  }
  if (lineasMrz.length < 2) return null;

  const x0 = Math.min(...lineasMrz.map((l) => l.bbox.x0)), x1 = Math.max(...lineasMrz.map((l) => l.bbox.x1));
  const y0 = Math.min(...lineasMrz.map((l) => l.bbox.y0)), y1 = Math.max(...lineasMrz.map((l) => l.bbox.y1));

  // 2) Se relee solo la zona mecánica, permitiendo nada más que sus caracteres.
  const margenMrz = (y1 - y0) * 0.15;
  const recorteMrz = { x: Math.max(0, x0 - 20), y: Math.max(0, y0 - margenMrz), w: Math.min(lienzo.width, x1 + 20) - Math.max(0, x0 - 20), h: Math.min(lienzo.height, y1 + margenMrz) - Math.max(0, y0 - margenMrz) };
  await lector.setParameters({ tessedit_char_whitelist: CARACTERES_MRZ, tessedit_pageseg_mode: "6" });
  const { data: dataMrz } = await lector.recognize(prepararLienzo(lienzo, { recorte: recorteMrz }));
  await lector.setParameters({ tessedit_char_whitelist: "", tessedit_pageseg_mode: "3" });
  const mrzLimpia = interpretarMrz(dataMrz.text);
  const mrzGeneral = interpretarMrz(textoGeneral);
  const datos = { ...mrzGeneral, ...Object.fromEntries(Object.entries(mrzLimpia).filter(([, v]) => v !== "" && v !== undefined)) };
  if (!datos.dniVerificado && mrzGeneral.dniVerificado) { datos.dni = mrzGeneral.dni; datos.dniVerificado = true; }

  // 3) Arriba de la zona mecánica está impreso el domicilio y el CUIL. La tarjeta mide
  //    85,6 x 54 mm y la zona mecánica ocupa casi todo el ancho: con eso se ubica el resto.
  const anchoTarjeta = (x1 - x0) / 0.88;
  const altoTarjeta = (anchoTarjeta * 54) / 85.6;
  const arriba = Math.max(0, y1 - altoTarjeta * 0.97);
  const izquierda = Math.max(0, x0 - anchoTarjeta * 0.06);
  const recorteSuperior = { x: izquierda, y: arriba, w: Math.min(lienzo.width - izquierda, anchoTarjeta * 1.02), h: Math.max(10, y0 - arriba) };
  const { data: dataSuperior } = await lector.recognize(prepararLienzo(lienzo, { recorte: recorteSuperior }));
  // El domicilio va en el renglón de arriba de todo: se relee esa franja sola, como un
  // bloque de texto, para que no lo corte en pedazos el resto del dibujo de la tarjeta.
  const franjaArriba = Math.max(0, y1 - altoTarjeta);
  const franja = { x: izquierda, y: franjaArriba, w: recorteSuperior.w, h: Math.min(lienzo.height - franjaArriba, altoTarjeta * 0.24) };
  await lector.setParameters({ tessedit_pageseg_mode: "6" });
  const { data: dataFranja } = await lector.recognize(prepararLienzo(lienzo, { recorte: franja }));
  await lector.setParameters({ tessedit_pageseg_mode: "3" });

  datos.domicilio = buscarDomicilio(dataFranja.text) || buscarDomicilio(dataSuperior.text);
  datos.cuil = buscarCuil(`${dataSuperior.text}\n${textoGeneral}`, datos.dni);
  return datos;
}

async function leerFrente(lector, img) {
  for (const rotacion of [0, 90, 270]) {
    const { data } = await lector.recognize(prepararLienzo(img, { rotacion }));
    const m = data.text.match(/\b(\d{1,2})[.,](\d{3})[.,](\d{3})\b/);
    if (m) return { dni: `${m[1]}${m[2]}${m[3]}` };
  }
  return {};
}

// Lee las fotos del DNI y devuelve { apellido, nombre, dni, fechaNacimiento, sexo, domicilio, cuil }
// con lo que se haya podido leer (los campos que no se leyeron quedan vacíos).
export async function leerDni({ frente, dorso }) {
  const lector = await crearLector();
  try {
    let datos = dorso ? await leerDorso(lector, await cargarImagen(dorso)) : null;
    if ((!datos || !datos.dni) && frente) {
      const delFrente = await leerFrente(lector, await cargarImagen(frente));
      datos = { ...(datos || {}), ...delFrente };
      if (datos.dni && !datos.cuil) datos.cuil = "";
    }
    return datos || {};
  } finally {
    await lector.terminate();
  }
}
