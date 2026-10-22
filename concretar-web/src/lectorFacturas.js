// Lectura automática de facturas de compra (PDF o foto) para completar solos el N° de
// factura, el tipo (A/B/C), la fecha, el total y las percepciones.
//
// - PDF: se lee el texto con su posición en la hoja y se arman renglones, así cada
//   etiqueta ("Per IB SJ", "Importe Total:", "Comp. Nro:"...) queda junto a su valor.
//   Además, si el PDF trae el QR de ARCA como link, de ahí salen N°, tipo, fecha y total
//   (es lo más confiable).
// - Foto: solo se puede leer el QR de ARCA (N°, tipo, fecha y total). Las percepciones
//   no están en el QR, así que en una foto hay que cargarlas a mano.
//
// pdf.js y el lector de QR se cargan recién cuando se usan, para no hacer más pesada
// la app al abrirla.

let pdfjsPromise = null;
function cargarPdfjs() {
  if (!pdfjsPromise) {
    pdfjsPromise = Promise.all([
      import("pdfjs-dist"),
      import("pdfjs-dist/build/pdf.worker.min.js?url"),
    ]).then(([pdfjs, worker]) => {
      const lib = pdfjs.default?.getDocument ? pdfjs.default : pdfjs;
      lib.GlobalWorkerOptions.workerSrc = worker.default;
      return lib;
    });
  }
  return pdfjsPromise;
}
let jsqrPromise = null;
function cargarJsqr() {
  if (!jsqrPromise) jsqrPromise = import("jsqr").then((m) => m.default || m);
  return jsqrPromise;
}

// Código de comprobante de ARCA → letra, y si es nota de crédito.
const TIPOS_CMP = {
  1: { letra: "A" }, 2: { letra: "A", notaDebito: true }, 3: { letra: "A", notaCredito: true },
  6: { letra: "B" }, 7: { letra: "B", notaDebito: true }, 8: { letra: "B", notaCredito: true },
  11: { letra: "C" }, 12: { letra: "C", notaDebito: true }, 13: { letra: "C", notaCredito: true },
  51: { letra: "A" }, 52: { letra: "A", notaDebito: true }, 53: { letra: "A", notaCredito: true },
  201: { letra: "A" }, 203: { letra: "A", notaCredito: true }, 206: { letra: "B" }, 208: { letra: "B", notaCredito: true },
  211: { letra: "C" }, 213: { letra: "C", notaCredito: true },
};

// "25,551.66" / "5700,00" / "$ 160164.91" / "1.234.567,89" → número.
export function parseMontoFactura(texto) {
  const limpio = String(texto || "").replace(/\s/g, "").replace(/\$/g, "");
  const m = limpio.match(/-?\d[\d.,]*/);
  if (!m) return null;
  let t = m[0].replace(/[.,]$/, "");
  const ultimaComa = t.lastIndexOf(",");
  const ultimoPunto = t.lastIndexOf(".");
  if (ultimaComa > -1 && ultimoPunto > -1) {
    const decimal = ultimaComa > ultimoPunto ? "," : ".";
    const miles = decimal === "," ? "." : ",";
    t = t.split(miles).join("").replace(decimal, ".");
  } else if (ultimaComa > -1 || ultimoPunto > -1) {
    const sep = ultimaComa > -1 ? "," : ".";
    const partes = t.split(sep);
    const decimales = partes[partes.length - 1];
    // Un solo separador con 1 o 2 dígitos atrás es la coma/punto decimal; con 3 dígitos
    // (o varios separadores) son miles.
    t = partes.length === 2 && decimales.length <= 2 ? partes.join(".") : partes.join("");
  }
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

const esItemNumerico = (s) => /\d/.test(s) && /^[$\s]*-?[\d.,\s]+$/.test(s);
// "F A C T U R A A" → "FACTURAA": algunos sistemas escriben las letras separadas.
const sinEspaciado = (s) => {
  const letras = s.replace(/\s/g, "").length;
  const espacios = (s.match(/\s/g) || []).length;
  return letras > 3 && espacios >= letras * 0.6 ? s.replace(/\s+/g, "") : s;
};
const fechaAISO = (d, m, a) => `${a}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
function primeraFecha(texto) {
  const m = String(texto).match(/\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b/);
  return m ? fechaAISO(m[1], m[2], m[3]) : null;
}
function formatearNumero(letra, puntoVenta, numero) {
  const pv = String(puntoVenta).replace(/\D/g, "").replace(/^0+(?=\d{4})/, "").padStart(4, "0");
  const nro = String(numero).replace(/\D/g, "").padStart(8, "0");
  return `${letra ? `${letra} ` : ""}${pv}${nro}`;
}
function tipoPercepcion(texto) {
  if (/iva/i.test(texto)) return "IVA";
  if (/gan/i.test(texto)) return "Ganancias";
  return "Ingresos Brutos";
}

// Datos del QR de ARCA: https://www.afip.gob.ar/fe/qr/?p=<base64 de un JSON>.
function datosDeQr(url) {
  try {
    const p = new URL(url).searchParams.get("p");
    if (!p) return null;
    const json = JSON.parse(atob(decodeURIComponent(p).replace(/-/g, "+").replace(/_/g, "/")));
    const tipo = TIPOS_CMP[Number(json.tipoCmp)] || {};
    return {
      tipoFactura: tipo.letra || null,
      esNotaCredito: !!tipo.notaCredito,
      comprobante: json.ptoVta && json.nroCmp ? formatearNumero(tipo.letra, json.ptoVta, json.nroCmp) : null,
      fecha: /^\d{4}-\d{2}-\d{2}/.test(json.fecha || "") ? json.fecha.slice(0, 10) : null,
      total: Number(json.importe) || null,
      cuitEmisor: json.cuit ? String(json.cuit) : null,
    };
  } catch {
    return null;
  }
}

// Arma renglones con los textos que están a la misma altura, ordenados de izquierda a derecha.
function armarRenglones(items) {
  const ordenados = items
    .filter((it) => it.str && it.str.trim())
    .map((it) => ({ str: sinEspaciado(it.str.trim()), x: it.transform[4], y: it.transform[5] }))
    .sort((a, b) => b.y - a.y || a.x - b.x);
  const renglones = [];
  for (const it of ordenados) {
    const r = renglones.find((x) => Math.abs(x.y - it.y) <= 3);
    if (r) r.items.push(it); else renglones.push({ y: it.y, items: [it] });
  }
  renglones.forEach((r) => {
    r.items.sort((a, b) => a.x - b.x);
    r.texto = r.items.map((i) => i.str).join(" ");
  });
  return renglones;
}
const ultimoNumeroDelRenglon = (r, desdeX = -Infinity) => {
  const nums = r.items.filter((i) => i.x > desdeX && esItemNumerico(i.str)).map((i) => parseMontoFactura(i.str)).filter((n) => n !== null);
  return nums.length ? nums[nums.length - 1] : null;
};

export function datosDeTextoFactura(items) {
  const renglones = armarRenglones(items);
  const todo = renglones.map((r) => r.texto).join("\n");
  const datos = { tipoFactura: null, esNotaCredito: false, comprobante: null, fecha: null, total: null, percepciones: [] };

  // Tipo: "Cod.01" / "COD. 011" / "Código 06".
  const cod = todo.match(/c[oó]d(?:igo)?\.?\s*:?\s*0*(\d{1,3})\b/i);
  const tipo = cod ? TIPOS_CMP[Number(cod[1])] : null;
  if (tipo) {
    datos.tipoFactura = tipo.letra;
    datos.esNotaCredito = !!tipo.notaCredito;
  } else {
    const m = todo.match(/FACTURA\s*([ABC])\b/i);
    if (m) datos.tipoFactura = m[1].toUpperCase();
  }
  if (/nota\s*de\s*cr[eé]dito/i.test(todo)) datos.esNotaCredito = true;

  // N°: "Punto de Venta: 00004  Comp. Nro: 00000148" (formato ARCA)...
  const itemsTodos = renglones.flatMap((r) => r.items.map((i) => ({ ...i, renglon: r })));
  const valorALaDerecha = (re) => {
    const etiqueta = itemsTodos.find((i) => re.test(i.str));
    if (!etiqueta) return null;
    const enLinea = etiqueta.renglon.items.filter((i) => i.x > etiqueta.x && /^\d{1,8}$/.test(i.str));
    if (enLinea.length) return enLinea[0].str;
    const pegado = etiqueta.str.match(/(\d{1,8})\s*$/);
    return pegado ? pegado[1] : null;
  };
  const pv = valorALaDerecha(/punto\s*de\s*venta/i);
  const nro = valorALaDerecha(/comp\.?\s*n(ro|°|º)/i);
  if (pv && nro) {
    datos.comprobante = formatearNumero(datos.tipoFactura, pv, nro);
  } else {
    // ... o "N° 0030-00045459" ...
    const guion = todo.match(/\b(\d{4,5})\s*-\s*(\d{8})\b/);
    if (guion) {
      datos.comprobante = formatearNumero(datos.tipoFactura, guion[1], guion[2]);
    } else {
      // ... o un número de 12/13 dígitos junto al título "FACTURA" (ej. Trielec: 001900056387).
      const titulo = itemsTodos.find((i) => /^(FACTURA|NOTA\s*DE\s*CR[EÉ]DITO)$/i.test(i.str));
      const candidatos = itemsTodos.filter((i) => /^\d{12,13}$/.test(i.str));
      if (candidatos.length) {
        const ref = titulo || { x: 400, y: Math.max(...itemsTodos.map((i) => i.y)) };
        candidatos.sort((a, b) => Math.hypot(a.x - ref.x, a.y - ref.y) - Math.hypot(b.x - ref.x, b.y - ref.y));
        const n = candidatos[0].str;
        datos.comprobante = formatearNumero(datos.tipoFactura, n.slice(0, n.length - 8), n.slice(-8));
      }
    }
  }

  // Fecha de emisión.
  const renglonFecha = renglones.find((r) => /emisi[oó]n|^fecha\s*:/i.test(r.texto) && primeraFecha(r.texto));
  datos.fecha = renglonFecha ? primeraFecha(renglonFecha.texto) : null;
  if (!datos.fecha) {
    const etiquetaFecha = itemsTodos.find((i) => /fecha\s*(de\s*)?emisi[oó]n/i.test(i.str));
    if (etiquetaFecha) {
      const cercana = itemsTodos
        .filter((i) => primeraFecha(i.str))
        .sort((a, b) => Math.hypot(a.x - etiquetaFecha.x, a.y - etiquetaFecha.y) - Math.hypot(b.x - etiquetaFecha.x, b.y - etiquetaFecha.y))[0];
      if (cercana) datos.fecha = primeraFecha(cercana.str);
    }
  }
  if (!datos.fecha) datos.fecha = primeraFecha(todo);

  // Total: "Importe Total: $ 5700,00" / "TOTAL $ 25,551.66" (no "SUBTOTAL").
  const renglonesTotal = renglones.filter((r) => /total/i.test(r.texto) && !/sub\s*-?\s*total|otros\s*tributos/i.test(r.texto));
  const conImporte = renglonesTotal.find((r) => /importe\s*total/i.test(r.texto) && ultimoNumeroDelRenglon(r) !== null);
  const renglonTotal = conImporte || renglonesTotal.find((r) => ultimoNumeroDelRenglon(r) !== null);
  datos.total = renglonTotal ? ultimoNumeroDelRenglon(renglonTotal) : null;

  // Percepciones: "Per IB SJ 0 $ 618,19", "Percepción IVA ... 1.234,56", "Perc. IIBB".
  const porTipo = {};
  renglones
    .filter((r) => /percep|perc\.|\bperc\b|^per\.?\s|\bper\.?\s*(ib|iibb|iva|ing)/i.test(r.texto) && !/r[eé]gimen|sujeto\s+a|no\s+corresponde/i.test(r.texto))
    .forEach((r) => {
      const etiqueta = r.items.find((i) => /per/i.test(i.str));
      const monto = ultimoNumeroDelRenglon(r, etiqueta ? etiqueta.x : -Infinity);
      if (!monto || monto <= 0) return;
      // Si el "monto" es en realidad la alícuota (ej. "3,00 %"), no se toma.
      if (/%\s*$/.test(r.texto) && monto <= 10) return;
      const tipoP = tipoPercepcion(etiqueta ? etiqueta.str + " " + r.texto : r.texto);
      porTipo[tipoP] = (porTipo[tipoP] || 0) + monto;
    });
  datos.percepciones = Object.entries(porTipo).map(([tipoP, monto]) => ({ tipo: tipoP, monto: Math.round(monto * 100) / 100 }));
  return datos;
}

function dataUrlABytes(dataUrl) {
  const b64 = String(dataUrl).split(",")[1] || "";
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

// Busca un QR en un canvas, probando a distintos tamaños (en fotos grandes el QR
// se lee mejor achicando la imagen).
async function leerQrDeCanvas(canvas) {
  const jsQR = await cargarJsqr();
  const ctx = canvas.getContext("2d");
  const tamanios = [1, 0.6, 0.4];
  for (const escala of tamanios) {
    let fuente = canvas;
    if (escala !== 1) {
      fuente = document.createElement("canvas");
      fuente.width = Math.max(1, Math.round(canvas.width * escala));
      fuente.height = Math.max(1, Math.round(canvas.height * escala));
      fuente.getContext("2d").drawImage(canvas, 0, 0, fuente.width, fuente.height);
    }
    const c = escala === 1 ? ctx : fuente.getContext("2d");
    const img = c.getImageData(0, 0, fuente.width, fuente.height);
    const qr = jsQR(img.data, img.width, img.height, { inversionAttempts: "attemptBoth" });
    if (qr?.data && /\/fe\/qr\//i.test(qr.data)) return qr.data;
  }
  return null;
}

async function leerPdf(dataUrl) {
  const pdfjs = await cargarPdfjs();
  const pdf = await pdfjs.getDocument({ data: dataUrlABytes(dataUrl), isEvalSupported: false }).promise;
  const page = await pdf.getPage(1);
  const contenido = await page.getTextContent();
  const datos = datosDeTextoFactura(contenido.items);
  // QR de ARCA: primero como link del PDF; si no, dibujando la hoja y leyendo la imagen.
  let qrUrl = null;
  try {
    const anotaciones = await page.getAnnotations();
    qrUrl = anotaciones.map((a) => a.url || a.unsafeUrl).find((u) => u && /\/fe\/qr\//i.test(u)) || null;
  } catch { /* sin anotaciones */ }
  if (!qrUrl) {
    const textoQr = contenido.items.map((i) => i.str).join("").match(/https?:\/\/[^\s]*\/fe\/qr\/\?p=[A-Za-z0-9+/=_%-]{40,}/);
    if (textoQr) qrUrl = textoQr[0];
  }
  if (!qrUrl && typeof document !== "undefined") {
    try {
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
      qrUrl = await leerQrDeCanvas(canvas);
    } catch { /* no se pudo dibujar la hoja */ }
  }
  const qr = qrUrl ? datosDeQr(qrUrl) : null;
  return { ...datos, ...quitarVacios(qr || {}), percepciones: datos.percepciones, origen: qr ? "pdf+qr" : "pdf" };
}

async function leerFoto(dataUrl) {
  const img = await new Promise((resolve, reject) => {
    const i = new Image();
    i.onload = () => resolve(i);
    i.onerror = reject;
    i.src = dataUrl;
  });
  const max = 2000;
  const escala = Math.min(1, max / Math.max(img.width, img.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(img.width * escala);
  canvas.height = Math.round(img.height * escala);
  canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
  const qrUrl = await leerQrDeCanvas(canvas);
  const qr = qrUrl ? datosDeQr(qrUrl) : null;
  return qr ? { ...qr, percepciones: [], origen: "foto+qr" } : { origen: "foto", percepciones: [] };
}

const quitarVacios = (o) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== null && v !== undefined && v !== ""));

// Punto de entrada: devuelve lo que se pudo leer (los campos que no se encontraron
// quedan en null) y de dónde salió.
export async function leerFactura(dataUrl, tipoArchivo) {
  const esPdf = /pdf/i.test(tipoArchivo || "") || /^data:application\/pdf/i.test(dataUrl || "");
  try {
    return esPdf ? await leerPdf(dataUrl) : await leerFoto(dataUrl);
  } catch (err) {
    return { error: err?.message || "No se pudo leer el archivo", percepciones: [] };
  }
}
