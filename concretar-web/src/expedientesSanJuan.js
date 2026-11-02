// Consulta de movimientos de expedientes del Gobierno de San Juan: es la misma consulta
// que hace la página https://sj.sanjuan.gob.ar/?uid=CUMV-VIEW-ExpedientesComponent
// (el servidor permite que se la llame desde otra página, así que no hace falta un
// intermediario).

export const URL_CONSULTA_EXPEDIENTES = "https://sj.sanjuan.gob.ar/?uid=CUMV-VIEW-ExpedientesComponent";
const API = "https://syrahapi.sanjuan.gob.ar/data";
export const TIPOS_EXPEDIENTE = ["EXP", "OC", "OFI", "TP", "HIST"];

// "400-001234-2025", "400-1234/25", "Expte. N° 400 1234 2025" → { prefijo, numero, anio }
export function parsearNumeroExpediente(texto) {
  const partes = String(texto || "").match(/\d+/g) || [];
  if (partes.length < 3) return null;
  const [prefijo, numero, anioTexto] = partes.slice(-3).map(Number);
  const anio = anioTexto < 100 ? 2000 + anioTexto : anioTexto;
  if (!prefijo || !numero || anio < 1990 || anio > 2100) return null;
  return { prefijo, numero, anio };
}

export const textoExpediente = (e) => `${e.prefijo}-${String(e.numero).padStart(6, "0")}-${e.anio}`;

// "24-06-2025 12:45:50" → "2025-06-24T12:45:50"
const fechaIso = (f) => {
  const m = String(f || "").match(/(\d{2})-(\d{2})-(\d{4})\s+(\d{2}:\d{2}(?::\d{2})?)/);
  return m ? `${m[3]}-${m[2]}-${m[1]}T${m[4]}` : null;
};

// Devuelve los movimientos del expediente, del más nuevo al más viejo:
// [{ organismo, oficina, fecha }]. Si no hay ninguno, el número no existe (o no se movió).
export async function consultarExpediente({ prefijo, numero, anio, tipo = "EXP" }, cantidad = 10) {
  const resp = await fetch(API, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ site: "dlc.siged.consulta.exp", tipo: "DLC", Prefijo: prefijo, Numero: numero, Anio: String(anio), Tipo: tipo, results: cantidad }),
  });
  if (!resp.ok) throw new Error(`La página del Gobierno respondió con error ${resp.status}`);
  const data = await resp.json();
  if (data.error) throw new Error(data.error);
  // El último renglón ("MOVIMIENTOS AL: ...") es solo la hora de la consulta.
  return (data.items || [])
    .filter((it) => it && it.title)
    .map((it) => ({ organismo: String(it.title).trim(), oficina: String(it.description || "").trim(), fecha: fechaIso(it.fecha) }));
}
