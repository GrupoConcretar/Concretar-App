// Página que ve el cliente con el link privado de su obra (?portal=<token>): avance,
// certificados y pagos, próximos materiales y fotos. Solo lectura y sin usuario: los datos
// salen de la función portal_obra de la base, que valida el link y devuelve únicamente lo
// que se le muestra al cliente (nunca costos, gastos ni proveedores).
import { useEffect, useState } from "react";
import { createClient } from "@supabase/supabase-js";
import { CalendarDays, Camera, CheckCircle2, Clock, Loader2, Package, Receipt, TrendingUp, X } from "lucide-react";
import { SUPABASE_URL, SUPABASE_ANON_KEY, isSupabaseConfigured } from "./supabaseConfig.js";

const NAVY = "#021d34";
const NAVY_700 = "#153f59";
const FUENTE = { fontFamily: "'Poppins', ui-sans-serif, system-ui, sans-serif" };

// Sin sesión guardada: el link funciona igual aunque en ese navegador haya alguien
// logueado en la app, y no toca su sesión.
const supabase = isSupabaseConfigured
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY, { auth: { persistSession: false, autoRefreshToken: false } })
  : null;

const fmtARS = (n) => new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", minimumFractionDigits: 2 }).format(Number(n) || 0);
function fechaLocal(iso) {
  if (!iso) return null;
  const [y, m, d] = String(iso).slice(0, 10).split("-").map(Number);
  return new Date(y, m - 1, d);
}
const fmtFecha = (iso) => (iso ? fechaLocal(iso).toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" }) : "—");
const fmtFechaLarga = (iso) => (iso ? fechaLocal(iso).toLocaleDateString("es-AR", { day: "numeric", month: "long", year: "numeric" }) : "—");
const hoyISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const diasEntre = (a, b) => Math.round((fechaLocal(b) - fechaLocal(a)) / 86400000);
function sumarMeses(iso, n) {
  const d = fechaLocal(iso);
  const r = new Date(d.getFullYear(), d.getMonth() + n, d.getDate());
  return `${r.getFullYear()}-${String(r.getMonth() + 1).padStart(2, "0")}-${String(r.getDate()).padStart(2, "0")}`;
}

// Avance ponderado por la duración de cada etapa (una etapa de 2 meses pesa más que una de 3 días).
function avanceDe(etapas) {
  let total = 0;
  let hecho = 0;
  etapas.forEach((e) => {
    const peso = e.inicio && e.fin ? Math.max(1, diasEntre(e.inicio, e.fin) + 1) : 1;
    total += peso;
    hecho += peso * Math.min(100, Number(e.avance) || 0);
  });
  return total ? hecho / total : null;
}
function estadoEtapa(e, hoy) {
  if ((Number(e.avance) || 0) >= 100) return "Terminada";
  if ((Number(e.avance) || 0) === 0 && e.inicio && e.inicio > hoy) return "Próxima";
  return "En curso";
}
const ESTILO_ETAPA = {
  Terminada: "bg-emerald-50 text-emerald-700 ring-emerald-600/20",
  "En curso": "bg-amber-50 text-amber-800 ring-amber-600/20",
  Próxima: "bg-slate-100 text-slate-600 ring-slate-500/20",
};
const ESTILO_CERTIFICADO = {
  Cobrado: "bg-emerald-50 text-emerald-700 ring-emerald-600/20",
  Pendiente: "bg-amber-50 text-amber-800 ring-amber-600/20",
  "A certificar": "bg-sky-50 text-sky-800 ring-sky-600/20",
};
const ETIQUETA_CERTIFICADO = { Cobrado: "Pagado", Pendiente: "A pagar", "A certificar": "A certificar" };
const ETIQUETA_PEDIDO = { Solicitado: "Pedido", Aprobado: "Confirmado", Facturado: "Comprado" };
const importe = (c) => (Number(c.monto) || 0) * (Number(c.factor) || 1);

function Chip({ children, className }) {
  return <span className={`inline-flex items-center whitespace-nowrap rounded-full px-2 py-0.5 text-[11px] font-semibold ring-1 ring-inset ${className}`}>{children}</span>;
}
function Seccion({ icono: Icono, titulo, children, extra }) {
  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <div className="mb-3 flex items-center justify-between gap-2">
        <h2 className="flex items-center gap-2 text-base font-semibold text-slate-900">
          <Icono size={18} style={{ color: NAVY_700 }} /> {titulo}
        </h2>
        {extra}
      </div>
      {children}
    </section>
  );
}
function Barra({ valor, alto = "h-2.5" }) {
  return (
    <div className={`${alto} w-full overflow-hidden rounded-full bg-slate-100`}>
      <div className="h-full rounded-full bg-amber-500 transition-[width] duration-700" style={{ width: `${Math.max(0, Math.min(100, valor || 0))}%` }} />
    </div>
  );
}

// Foto: se baja cuando se muestra (cada una por separado, para que la página cargue rápido).
function Foto({ foto, token, onAbrir }) {
  const [src, setSrc] = useState(foto.src || null);
  useEffect(() => {
    if (src || !supabase) return;
    let vivo = true;
    supabase.rpc("portal_obra_foto", { p_token: token, p_foto_id: foto.id }).then(({ data }) => { if (vivo && data) setSrc(data); });
    return () => { vivo = false; };
  }, [foto.id, token, src]);
  return (
    <button type="button" onClick={() => src && onAbrir({ ...foto, src })} className="group relative aspect-[4/3] overflow-hidden rounded-xl bg-slate-100 focus:outline-none focus-visible:ring-2 focus-visible:ring-amber-500">
      {src ? (
        <img src={src} alt={foto.descripcion || `Foto del ${fmtFecha(foto.fecha)}`} className="h-full w-full object-cover transition-transform duration-300 group-hover:scale-[1.03]" />
      ) : (
        <span className="flex h-full w-full items-center justify-center text-slate-300"><Loader2 size={20} className="animate-spin" /></span>
      )}
      <span className="absolute inset-x-0 bottom-0 bg-gradient-to-t from-black/60 to-transparent px-2 pb-1.5 pt-5 text-left text-[11px] font-medium text-white">
        {fmtFecha(foto.fecha)}{foto.descripcion ? ` · ${foto.descripcion}` : ""}
      </span>
    </button>
  );
}

export default function PortalCliente({ token }) {
  const [estado, setEstado] = useState({ cargando: true, datos: null, error: null });
  const [fotoAbierta, setFotoAbierta] = useState(null);

  useEffect(() => {
    if (!supabase) { setEstado({ cargando: false, datos: DEMO_PORTAL, error: null }); return; }
    supabase.rpc("portal_obra", { p_token: token }).then(({ data, error }) => {
      if (error) setEstado({ cargando: false, datos: null, error: "No se pudo cargar la información. Probá de nuevo en un rato." });
      else if (!data) setEstado({ cargando: false, datos: null, error: "Este link no está disponible. Pedile uno nuevo a Grupo Concretar." });
      else setEstado({ cargando: false, datos: data, error: null });
    });
  }, [token]);
  useEffect(() => {
    if (estado.datos?.obra?.nombre) document.title = `${estado.datos.obra.nombre} · Grupo Concretar`;
  }, [estado.datos]);
  useEffect(() => {
    if (!fotoAbierta) return;
    const cerrar = (e) => { if (e.key === "Escape") setFotoAbierta(null); };
    window.addEventListener("keydown", cerrar);
    return () => window.removeEventListener("keydown", cerrar);
  }, [fotoAbierta]);

  const encabezado = (
    <header style={{ backgroundColor: NAVY }} className="text-white">
      <div className="mx-auto flex max-w-3xl items-center gap-3 px-4 py-3">
        <img src="/gc-logo-white.png" alt="Grupo Concretar" className="h-7 w-auto" />
        <span className="text-sm font-medium tracking-wide text-slate-300">Grupo Concretar</span>
      </div>
    </header>
  );

  if (estado.cargando || estado.error) {
    return (
      <div className="min-h-screen bg-slate-50" style={FUENTE}>
        {encabezado}
        <div className="mx-auto flex max-w-3xl flex-col items-center gap-3 px-4 py-20 text-center text-slate-600">
          {estado.cargando ? <Loader2 size={28} className="animate-spin text-slate-400" /> : <p className="max-w-sm text-sm">{estado.error}</p>}
        </div>
      </div>
    );
  }

  const { obra, contratos, etapas, certificados, materiales, fotos } = estado.datos;
  const hoy = hoyISO();
  const avance = avanceDe(etapas);
  const finEtapas = etapas.map((e) => e.fin).filter(Boolean).sort().at(-1);
  const finEstimado = finEtapas || (obra.inicio && obra.meses ? sumarMeses(obra.inicio, obra.meses) : null);
  const varios = contratos.length > 1;
  const nombreContrato = (id) => {
    const c = contratos.find((x) => x.id === id);
    return c ? c.nombre.replace(`${obra.nombre} — `, "") : "";
  };
  const totalContrato = contratos.reduce((s, c) => s + (Number(c.presupuesto) || 0), 0);
  const pagado = certificados.filter((c) => c.estado === "Cobrado").reduce((s, c) => s + importe(c), 0);
  const porPagar = certificados.filter((c) => c.estado !== "Cobrado");
  const totalPorPagar = porPagar.reduce((s, c) => s + importe(c), 0);
  const proximo = [...porPagar].filter((c) => c.fechaCobro).sort((a, b) => a.fechaCobro.localeCompare(b.fechaCobro))[0];
  const conFactor = certificados.some((c) => c.factor && Number(c.factor) !== 1);
  const etapasOrdenadas = [...etapas].sort((a, b) => (a.inicio || "").localeCompare(b.inicio || ""));

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800" style={FUENTE}>
      {encabezado}

      <div style={{ backgroundColor: NAVY }} className="pb-16 text-white">
        <div className="mx-auto max-w-3xl px-4 pt-6">
          <p className="text-xs font-semibold uppercase tracking-[0.18em] text-amber-400">Seguimiento de obra</p>
          <h1 className="mt-1 text-2xl font-bold leading-tight sm:text-3xl" style={{ textWrap: "balance" }}>{obra.nombre}</h1>
          {obra.cliente && <p className="mt-1 text-sm text-slate-300">{obra.cliente}</p>}
          <p className="mt-3 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-300">
            <span className="flex items-center gap-1"><CalendarDays size={13} /> Inicio {fmtFecha(obra.inicio)}</span>
            {finEstimado && <span className="flex items-center gap-1"><Clock size={13} /> Fin estimado {fmtFecha(finEstimado)}</span>}
          </p>
        </div>
      </div>

      <main className="mx-auto -mt-12 max-w-3xl space-y-4 px-4 pb-10">
        <Seccion icono={TrendingUp} titulo="Avance de la obra">
          {avance === null ? (
            <p className="text-sm text-slate-500">Todavía no está cargada la planificación de la obra.</p>
          ) : (
            <>
              <div className="flex items-end justify-between gap-3">
                <div className="text-4xl font-bold tabular-nums" style={{ color: NAVY }}>{Math.round(avance)}<span className="text-2xl">%</span></div>
                <div className="pb-1 text-right text-xs text-slate-500">
                  {etapas.filter((e) => (Number(e.avance) || 0) >= 100).length} de {etapas.length} etapas terminadas
                </div>
              </div>
              <div className="mt-2"><Barra valor={avance} alto="h-3" /></div>

              {varios && (
                <div className="mt-4 grid gap-2 sm:grid-cols-2">
                  {contratos.map((c) => {
                    const a = avanceDe(etapas.filter((e) => e.obraId === c.id));
                    return (
                      <div key={c.id} className="rounded-xl bg-slate-50 p-3">
                        <div className="flex items-center justify-between gap-2 text-sm">
                          <span className="font-medium text-slate-800">{nombreContrato(c.id)}</span>
                          <span className="font-semibold tabular-nums text-slate-700">{a === null ? "—" : `${Math.round(a)}%`}</span>
                        </div>
                        <div className="mt-1.5"><Barra valor={a || 0} alto="h-1.5" /></div>
                      </div>
                    );
                  })}
                </div>
              )}

              <ul className="mt-4 divide-y divide-slate-100">
                {etapasOrdenadas.map((e, i) => {
                  const est = estadoEtapa(e, hoy);
                  return (
                    <li key={i} className="py-2.5">
                      <div className="flex flex-wrap items-center justify-between gap-x-3 gap-y-1">
                        <span className="text-sm font-medium text-slate-800">{e.nombre}</span>
                        <span className="flex items-center gap-2">
                          <span className="text-xs tabular-nums text-slate-500">{fmtFecha(e.inicio)} → {fmtFecha(e.fin)}</span>
                          <Chip className={ESTILO_ETAPA[est]}>{est}</Chip>
                        </span>
                      </div>
                      <div className="mt-1.5 flex items-center gap-2">
                        <Barra valor={e.avance} alto="h-1.5" />
                        <span className="w-9 shrink-0 text-right text-xs tabular-nums text-slate-500">{Math.round(Number(e.avance) || 0)}%</span>
                      </div>
                    </li>
                  );
                })}
              </ul>
            </>
          )}
        </Seccion>

        <Seccion icono={Receipt} titulo="Certificados y pagos">
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
            {totalContrato > 0 && (
              <div className="rounded-xl bg-slate-50 p-3">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Total del contrato</div>
                <div className="mt-0.5 text-lg font-bold tabular-nums text-slate-900">{fmtARS(totalContrato)}</div>
              </div>
            )}
            <div className="rounded-xl bg-emerald-50 p-3">
              <div className="text-[11px] font-semibold uppercase tracking-wide text-emerald-700">Pagado</div>
              <div className="mt-0.5 text-lg font-bold tabular-nums text-emerald-800">{fmtARS(pagado)}</div>
            </div>
            <div className="rounded-xl bg-amber-50 p-3">
              <div className="text-[11px] font-semibold uppercase tracking-wide text-amber-800">Por pagar</div>
              <div className="mt-0.5 text-lg font-bold tabular-nums text-amber-900">{fmtARS(totalPorPagar)}</div>
            </div>
          </div>

          {proximo && (
            <div className="mt-3 flex flex-wrap items-center justify-between gap-2 rounded-xl px-3 py-2.5 text-white" style={{ backgroundColor: NAVY_700 }}>
              <span className="text-sm">Próximo pago: <span className="font-semibold">{proximo.concepto || "Certificado"}</span></span>
              <span className="text-sm"><span className="font-bold tabular-nums">{fmtARS(importe(proximo))}</span> · {fmtFechaLarga(proximo.fechaCobro)}</span>
            </div>
          )}

          {certificados.length === 0 ? (
            <p className="mt-3 text-sm text-slate-500">Todavía no hay certificados cargados.</p>
          ) : (
            <>
              {/* Celular: una tarjeta por certificado. */}
              <ul className="mt-3 space-y-2 sm:hidden">
                {certificados.map((c, i) => (
                  <li key={i} className="rounded-xl border border-slate-200 p-3">
                    <div className="flex items-start justify-between gap-2">
                      <div className="min-w-0">
                        <div className="text-sm font-semibold text-slate-900">{c.concepto || "Certificado"}</div>
                        {varios && <div className="text-xs text-slate-500">{nombreContrato(c.obraId)}</div>}
                      </div>
                      <Chip className={ESTILO_CERTIFICADO[c.estado]}>{ETIQUETA_CERTIFICADO[c.estado]}</Chip>
                    </div>
                    <div className="mt-2 flex items-end justify-between gap-2 text-xs text-slate-500">
                      <div>
                        <div>Certificado: {fmtFecha(c.fecha)}</div>
                        <div>{c.estado === "Cobrado" ? "Pagado el" : "Día de cobro"}: <span className="font-medium text-slate-700">{fmtFecha(c.fechaCobro)}</span></div>
                      </div>
                      <div className="text-right">
                        {c.factor && Number(c.factor) !== 1 && (
                          <div className="tabular-nums">{fmtARS(c.monto)} × {Number(c.factor).toLocaleString("es-AR", { maximumFractionDigits: 4 })}</div>
                        )}
                        <div className="text-base font-bold tabular-nums text-slate-900">{fmtARS(importe(c))}</div>
                      </div>
                    </div>
                  </li>
                ))}
              </ul>
              {/* Tablet/PC: tabla. */}
              <div className="mt-3 hidden overflow-x-auto sm:block">
                <table className="w-full text-left text-sm">
                  <thead className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                    <tr className="border-b border-slate-200">
                      <th className="py-2 pr-3">Certificado</th>
                      <th className="py-2 pr-3">Fecha</th>
                      <th className="py-2 pr-3 text-right">Monto</th>
                      {conFactor && <th className="py-2 pr-3 text-right">Factor</th>}
                      {conFactor && <th className="py-2 pr-3 text-right">Actualizado</th>}
                      <th className="py-2 pr-3">Día de cobro</th>
                      <th className="py-2">Estado</th>
                    </tr>
                  </thead>
                  <tbody className="tabular-nums">
                    {certificados.map((c, i) => (
                      <tr key={i} className="border-b border-slate-100 last:border-0">
                        <td className="py-2 pr-3">
                          <div className="font-medium text-slate-900">{c.concepto || "Certificado"}</div>
                          {varios && <div className="text-xs text-slate-500">{nombreContrato(c.obraId)}</div>}
                        </td>
                        <td className="py-2 pr-3 text-slate-600">{fmtFecha(c.fecha)}</td>
                        <td className="py-2 pr-3 text-right text-slate-800">{fmtARS(c.monto)}</td>
                        {conFactor && <td className="py-2 pr-3 text-right text-slate-600">{c.factor ? Number(c.factor).toLocaleString("es-AR", { maximumFractionDigits: 4 }) : "—"}</td>}
                        {conFactor && <td className="py-2 pr-3 text-right font-semibold text-slate-900">{fmtARS(importe(c))}</td>}
                        <td className="py-2 pr-3 text-slate-600">{fmtFecha(c.fechaCobro)}</td>
                        <td className="py-2"><Chip className={ESTILO_CERTIFICADO[c.estado]}>{ETIQUETA_CERTIFICADO[c.estado]}</Chip></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {conFactor && <p className="mt-2 text-xs text-slate-500">Los montos con factor se actualizan multiplicando el certificado por su factor de actualización.</p>}
            </>
          )}
        </Seccion>

        {materiales.length > 0 && (
          <Seccion icono={Package} titulo="Próximos materiales">
            <ul className="space-y-3">
              {materiales.map((p, i) => (
                <li key={i} className="rounded-xl border border-slate-200 p-3">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <span className="text-sm font-semibold text-slate-900">
                      {p.fechaNecesaria ? `Para el ${fmtFechaLarga(p.fechaNecesaria)}` : `Pedido del ${fmtFecha(p.fecha)}`}
                      {varios && <span className="font-normal text-slate-500"> · {nombreContrato(p.obraId)}</span>}
                    </span>
                    <Chip className="bg-slate-100 text-slate-700 ring-slate-500/20">{ETIQUETA_PEDIDO[p.estado] || p.estado}</Chip>
                  </div>
                  <ul className="mt-2 space-y-1 text-sm">
                    {p.items.map((it, j) => (
                      <li key={j} className="flex items-start justify-between gap-3">
                        <span className="text-slate-700">
                          <span className="font-medium tabular-nums text-slate-900">{Number(it.cantidad || 0).toLocaleString("es-AR")} {it.unidad || ""}</span> {it.material}
                        </span>
                        {it.total != null && <span className="shrink-0 tabular-nums text-slate-600">{fmtARS(it.total)}</span>}
                      </li>
                    ))}
                  </ul>
                  {p.total != null && (
                    <div className="mt-2 flex justify-between border-t border-slate-100 pt-2 text-sm font-semibold text-slate-900">
                      <span>Total</span><span className="tabular-nums">{fmtARS(p.total)}</span>
                    </div>
                  )}
                </li>
              ))}
            </ul>
          </Seccion>
        )}

        {fotos.length > 0 && (
          <Seccion icono={Camera} titulo="Fotos de la obra" extra={<span className="text-xs text-slate-500">{fotos.length} foto{fotos.length === 1 ? "" : "s"}</span>}>
            <div className="grid grid-cols-2 gap-2 sm:grid-cols-3">
              {fotos.map((f) => <Foto key={f.id} foto={f} token={token} onAbrir={setFotoAbierta} />)}
            </div>
          </Seccion>
        )}

        <p className="flex items-center justify-center gap-1.5 pt-2 text-center text-xs text-slate-400">
          <CheckCircle2 size={13} /> Información al {fmtFechaLarga(hoy)} · se actualiza sola con los datos de la obra.
        </p>
      </main>

      {fotoAbierta && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 p-4" onClick={() => setFotoAbierta(null)}>
          <button type="button" aria-label="Cerrar" onClick={() => setFotoAbierta(null)} className="absolute right-3 top-3 rounded-full bg-white/10 p-2 text-white hover:bg-white/20"><X size={20} /></button>
          <figure className="max-h-full max-w-4xl" onClick={(e) => e.stopPropagation()}>
            <img src={fotoAbierta.src} alt={fotoAbierta.descripcion || "Foto de la obra"} className="max-h-[80vh] w-auto rounded-lg object-contain" />
            <figcaption className="mt-2 text-center text-sm text-slate-200">{fmtFechaLarga(fotoAbierta.fecha)}{fotoAbierta.descripcion ? ` · ${fotoAbierta.descripcion}` : ""}</figcaption>
          </figure>
        </div>
      )}
    </div>
  );
}

// Ejemplo para la demo (sin base conectada).
const fotoDemo = (color, texto) => `data:image/svg+xml;utf8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600"><rect width="800" height="600" fill="${color}"/><text x="400" y="310" font-family="sans-serif" font-size="42" fill="white" text-anchor="middle">${texto}</text></svg>`)}`;
const DEMO_PORTAL = {
  obra: { nombre: "Edificio Belgrano 450", cliente: "Consorcio Belgrano SA", inicio: "2026-02-01", meses: 12, estado: "En curso" },
  mostrarPreciosMateriales: true,
  contratos: [{ id: 1, nombre: "Edificio Belgrano 450", presupuesto: 85000000, inicio: "2026-02-01", meses: 12, estado: "En curso" }],
  etapas: [
    { obraId: 1, nombre: "Excavación y fundaciones", inicio: "2026-02-01", fin: "2026-03-31", avance: 100 },
    { obraId: 1, nombre: "Estructura de hormigón", inicio: "2026-04-01", fin: "2026-08-31", avance: 100 },
    { obraId: 1, nombre: "Mampostería", inicio: "2026-08-15", fin: "2026-11-30", avance: 65 },
    { obraId: 1, nombre: "Instalaciones", inicio: "2026-10-01", fin: "2026-12-31", avance: 20 },
    { obraId: 1, nombre: "Terminaciones", inicio: "2027-01-02", fin: "2027-01-31", avance: 0 },
  ],
  certificados: [
    { tipo: "cobro", obraId: 1, concepto: "Anticipo certificado 1", fecha: "2026-02-05", monto: 20000000, estado: "Cobrado", fechaCobro: "2026-02-05", factor: null },
    { tipo: "cobro", obraId: 1, concepto: "Certificado de avance 2", fecha: "2026-04-10", monto: 18000000, estado: "Cobrado", fechaCobro: "2026-04-10", factor: null },
    { tipo: "cobro", obraId: 1, concepto: "Certificado de avance 3", fecha: "2026-09-20", monto: 15000000, estado: "Pendiente", fechaCobro: "2026-11-10", factor: 1.085 },
    { tipo: "futuro", obraId: 1, concepto: "Certificado de avance 4", fecha: "2026-12-01", monto: 16000000, estado: "A certificar", fechaCobro: "2026-12-15", factor: 1.12 },
  ],
  materiales: [
    { fecha: "2026-11-01", fechaNecesaria: "2026-11-12", estado: "Aprobado", obraId: 1, total: 1250000, items: [
      { material: "Ladrillo hueco 12x18x33", cantidad: 1500, unidad: "u.", precioUnitario: 650, total: 975000 },
      { material: "Cemento x 50 kg", cantidad: 25, unidad: "bolsas", precioUnitario: 11000, total: 275000 },
    ] },
  ],
  fotos: [
    { id: 1, fecha: "2026-11-03", descripcion: "Mampostería 2° piso", src: fotoDemo("#3a5c66", "Mampostería") },
    { id: 2, fecha: "2026-10-20", descripcion: "Cañerías", src: fotoDemo("#153f59", "Instalaciones") },
    { id: 3, fecha: "2026-09-02", descripcion: "Losa terminada", src: fotoDemo("#7a6a4f", "Estructura") },
  ],
};
