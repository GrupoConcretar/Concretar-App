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

// ---------- Circuito del expediente (DMyOM) ----------

// Pasos del circuito administrativo de un expediente de obra contratada por DMyOM (una vez
// aprobado el presupuesto), según la planilla de la oficina. "dias" son días hábiles
// estimados (se pueden cambiar en cada expediente). "claves" reconocen la oficina tal como
// figura en la página del Gobierno (ej. "División Administrativa Contable DMyOM" es
// Administración OM y "Departamento Contable - SI" es Contable de obra).
const ADM_OM = ["administrativa contable", "administracion om", "administracion de obras menores"];
const CONTABLE = ["departamento contable", "contaduria"];
const SAF = ["administrativa financiera", "adm. financiera", "(saf)"];
const DESPACHO = ["despacho"];
const DELEGACION = ["delegacion fiscal"];
const DMYOM = ["direccion de mantenimiento", "mantenimiento y obras menores"];
export const PASOS_EXPEDIENTE = [
  { n: 1, oficina: "Administración OM", motivo: "Agregar documentación al expediente", dias: 3, claves: ADM_OM },
  { n: 2, oficina: "Contable de obra", motivo: "Imputación contable", dias: 3, claves: CONTABLE },
  { n: 3, oficina: "Secret. Adm. Financiera (SAF)", motivo: "V°B° en imputación contable", dias: 3, claves: SAF },
  { n: 4, oficina: "Despacho", motivo: "Confección de contrato", dias: 5, claves: DESPACHO },
  { n: 5, oficina: "Administración OM", motivo: "Firma de contrato y posterior aforado", dias: 5, claves: ADM_OM },
  { n: 6, oficina: "Despacho", motivo: "Proyecto de resolución", dias: 5, claves: DESPACHO },
  { n: 7, oficina: "Asesoría Letrada", motivo: "Dictamen", dias: 10, claves: ["asesoria letrada"] },
  { n: 8, oficina: "Despacho", motivo: "Resolución", dias: 5, claves: DESPACHO },
  { n: 9, oficina: "Administración OM", motivo: "Firma de resolución", dias: 3, claves: ADM_OM },
  { n: 10, oficina: "Delegación Fiscal", motivo: "V°B° de auditor técnico y delegados", dias: 7, claves: DELEGACION },
  { n: 11, oficina: "Administración OM", motivo: "Notificar resolución a la empresa", dias: 3, claves: ADM_OM },
  { n: 12, oficina: "DMyOM", motivo: "A la espera de certificación", dias: 20, claves: DMYOM },
  { n: 13, oficina: "DMyOM", motivo: "Con certificados aprobados", dias: 5, claves: DMYOM },
  { n: 14, oficina: "Contable de obra", motivo: "Compromiso", dias: 3, claves: CONTABLE },
  { n: 15, oficina: "Delegación Fiscal", motivo: "Aprobar compromiso", dias: 5, claves: DELEGACION },
  { n: 16, oficina: "Contable de obra", motivo: "A la espera de factura", dias: 3, claves: CONTABLE },
  { n: 17, oficina: "Secret. Adm. Financiera (SAF)", motivo: "Aprobar ordenado en curso", dias: 5, claves: SAF },
  { n: 18, oficina: "Administración OM", motivo: "Firma de ordenado aprobado", dias: 3, claves: ADM_OM },
  { n: 19, oficina: "Habilitación", motivo: "Para pago", dias: 7, claves: ["habilitacion"] },
  { n: 20, oficina: "Pagado", motivo: "Cobrado por la empresa", dias: 0, claves: [] },
];

export const sinAcentos = (t) => String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
const coincide = (paso, mov) => {
  const texto = sinAcentos(`${mov.organismo} ${mov.oficina}`);
  return paso.claves.some((c) => texto.includes(c));
};

// Feriados nacionales y días no laborables con fines turísticos (la administración
// pública no trabaja). Se suman a los feriados cargados en la app (Calendario), donde
// se pueden agregar los provinciales.
export const FERIADOS_NACIONALES = [
  "2026-01-01", "2026-02-16", "2026-02-17", "2026-03-23", "2026-03-24", "2026-04-02", "2026-04-03",
  "2026-05-01", "2026-05-25", "2026-06-15", "2026-06-20", "2026-07-09", "2026-07-10", "2026-08-17",
  "2026-10-12", "2026-11-23", "2026-12-07", "2026-12-08", "2026-12-25",
  "2027-01-01", "2027-02-08", "2027-02-09", "2027-03-24", "2027-03-26", "2027-04-02", "2027-05-01",
  "2027-05-25", "2027-06-21", "2027-06-20", "2027-07-09", "2027-08-16", "2027-10-11", "2027-11-22",
  "2027-12-08", "2027-12-25",
];

// Fechas "AAAA-MM-DD" manejadas en UTC, así no las corre el huso horario.
const aFecha = (iso) => new Date(`${String(iso).slice(0, 10)}T00:00:00Z`);
const aIso = (d) => d.toISOString().slice(0, 10);
export function esDiaHabil(iso, feriados) {
  const dia = aFecha(iso).getUTCDay();
  return dia !== 0 && dia !== 6 && !feriados.has(String(iso).slice(0, 10));
}
// El día hábil número "n" contando desde el día siguiente a "iso".
export function sumarDiasHabiles(iso, n, feriados) {
  const d = aFecha(iso);
  let faltan = n;
  while (faltan > 0) {
    d.setUTCDate(d.getUTCDate() + 1);
    if (esDiaHabil(aIso(d), feriados)) faltan--;
  }
  return aIso(d);
}
// Días hábiles desde el día siguiente a "desde" hasta "hasta" (inclusive).
export function diasHabilesEntre(desde, hasta, feriados) {
  let n = 0;
  const d = aFecha(desde);
  const fin = aFecha(hasta);
  while (d < fin) {
    d.setUTCDate(d.getUTCDate() + 1);
    if (esDiaHabil(aIso(d), feriados)) n++;
  }
  return n;
}

// Reconstruye por qué pasos fue el expediente leyendo sus movimientos en orden: cada
// llegada a una oficina se asigna al próximo paso de esa oficina (mirando hasta 3 pasos
// adelante), y las oficinas que no están en el circuito (ej. "Comunicación y Copia") o
// las idas y vueltas se ignoran. Un paso queda hecho el día que el expediente llega a la
// oficina del paso siguiente. Devuelve { hechos: { n: fecha }, actual, desde }.
export function pasosSegunMovimientos(movimientos) {
  const cronologicos = [...(movimientos || [])].filter((m) => m.fecha).sort((a, b) => a.fecha.localeCompare(b.fecha));
  const llegadas = {};
  let siguiente = 1;
  for (const m of cronologicos) {
    const paso = PASOS_EXPEDIENTE.slice(siguiente - 1, siguiente + 3).find((p) => coincide(p, m));
    if (!paso) continue;
    llegadas[paso.n] = m.fecha.slice(0, 10);
    siguiente = paso.n + 1;
  }
  const llegados = Object.keys(llegadas).map(Number).sort((a, b) => a - b);
  if (llegados.length === 0) return { hechos: {}, actual: null, desde: null };
  const actual = llegados[llegados.length - 1];
  const hechos = {};
  for (let n = 1; n < actual; n++) {
    const proxima = llegados.find((k) => k > n);
    hechos[n] = llegadas[proxima];
  }
  return { hechos, actual, desde: llegadas[actual] };
}

// Estado de cada paso: hecho (con su fecha), el actual (días hábiles que lleva y si está
// atrasado) y los que faltan, con la fecha estimada en que terminarían.
export function calcularPasosExpediente(exp, hoy, feriados) {
  const diasDe = (paso) => Number(exp.dias?.[paso.n] ?? paso.dias);
  let desde = exp.fechaInicio || hoy;
  let actualVisto = false;
  return PASOS_EXPEDIENTE.map((paso) => {
    const hecho = exp.pasos?.[paso.n] || null;
    const dias = diasDe(paso);
    if (hecho) {
      desde = hecho;
      return { ...paso, dias, estado: "hecho", fecha: hecho };
    }
    if (!actualVisto) {
      actualVisto = true;
      const vence = sumarDiasHabiles(desde, dias, feriados);
      const transcurridos = diasHabilesEntre(desde, hoy, feriados);
      const atrasado = transcurridos > dias;
      const inicio = desde;
      // Si ya está atrasado, lo que sigue se estima a partir de hoy.
      desde = atrasado ? hoy : vence;
      return { ...paso, dias, estado: atrasado ? "atrasado" : "actual", desdeFecha: inicio, transcurridos, fecha: vence };
    }
    desde = sumarDiasHabiles(desde, dias, feriados);
    return { ...paso, dias, estado: "pendiente", fecha: desde };
  });
}
