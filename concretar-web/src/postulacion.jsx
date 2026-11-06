// Formulario público para que personal externo deje sus datos (y su CV, opcional) y
// quede en la sección Postulantes de la app. Se abre con el link fijo ?trabajo, sin
// usuario. Solo puede cargar una postulación nueva: no lee nada de la base (la función
// crear_postulacion valida todo del lado del servidor).
import { useEffect, useState } from "react";
import { createClient } from "@supabase/supabase-js";
import { CheckCircle2, FileText, ImagePlus, Loader2, X } from "lucide-react";
import { SUPABASE_URL, SUPABASE_ANON_KEY, isSupabaseConfigured } from "./supabaseConfig.js";
import { RUBROS_POSTULANTE, CATEGORIAS_POSTULANTE, MODALIDADES_POSTULANTE } from "./postulantesConfig.js";

const NAVY = "#021d34";
const FUENTE = { fontFamily: "'Poppins', ui-sans-serif, system-ui, sans-serif" };
const supabase = isSupabaseConfigured
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY, { auth: { persistSession: false, autoRefreshToken: false } })
  : null;
const MAX_CV_MB = 5;
const MAX_FOTOS = 3;

const inputCls = "w-full rounded-lg border border-slate-300 bg-white px-3 py-2.5 text-[15px] text-slate-900 outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/30";

function leerComoDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("No se pudo leer el archivo"));
    r.readAsDataURL(file);
  });
}
// Las fotos del celular pesan mucho: se achican a 1600 px en JPG antes de mandarlas.
async function fotoLiviana(file, ladoMax = 1600, calidad = 0.72) {
  const url = URL.createObjectURL(file);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    const escala = Math.min(1, ladoMax / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement("canvas");
    c.width = Math.round(img.naturalWidth * escala);
    c.height = Math.round(img.naturalHeight * escala);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
    return c.toDataURL("image/jpeg", calidad);
  } finally {
    URL.revokeObjectURL(url);
  }
}

function Campo({ label, opcional, children, ayuda }) {
  return (
    <label className="block">
      <span className="mb-1 block text-sm font-medium text-slate-700">
        {label}{opcional && <span className="font-normal text-slate-400"> (opcional)</span>}
      </span>
      {children}
      {ayuda && <span className="mt-1 block text-xs text-slate-500">{ayuda}</span>}
    </label>
  );
}
function Bloque({ titulo, children }) {
  return (
    <section className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <h2 className="mb-3 text-base font-semibold text-slate-900">{titulo}</h2>
      <div className="space-y-3">{children}</div>
    </section>
  );
}
function SiNo({ value, onChange }) {
  return (
    <div className="flex gap-2">
      {[["Sí", true], ["No", false]].map(([t, v]) => (
        <button key={t} type="button" onClick={() => onChange(value === v ? null : v)}
          className={`rounded-lg border px-4 py-2 text-sm font-medium ${value === v ? "border-amber-500 bg-amber-50 text-amber-900" : "border-slate-300 bg-white text-slate-700"}`}>
          {t}
        </button>
      ))}
    </div>
  );
}

export default function Postulacion() {
  const [f, setF] = useState({
    nombre: "", apellido: "", dni: "", telefono: "", email: "", localidad: "", edad: "",
    anios_experiencia: "", ultimas_obras: "", herramientas_propias: null, tarjeta_ieric: null,
    disponible_desde: "", modalidad: "", comentarios: "", consentimiento: false,
  });
  // { "Electricidad": "Oficial", ... }
  const [rubros, setRubros] = useState({});
  const [otroRubro, setOtroRubro] = useState("");
  const [cv, setCv] = useState(null); // { dataUrl, nombre, tipo }
  const [fotos, setFotos] = useState([]);
  const [procesando, setProcesando] = useState(false);
  const [enviando, setEnviando] = useState(false);
  const [error, setError] = useState("");
  const [enviado, setEnviado] = useState(false);
  const set = (k) => (e) => setF((x) => ({ ...x, [k]: e?.target ? e.target.value : e }));

  useEffect(() => { document.title = "Trabajá con nosotros · Grupo Concretar"; }, []);

  function toggleRubro(r) {
    setRubros((prev) => {
      const n = { ...prev };
      if (r in n) delete n[r]; else n[r] = "";
      return n;
    });
  }
  async function elegirCv(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setError("");
    setProcesando(true);
    try {
      if (file.type.startsWith("image/")) {
        setCv({ dataUrl: await fotoLiviana(file), nombre: file.name.replace(/\.[^.]+$/, "") + ".jpg", tipo: "image/jpeg" });
      } else {
        if (file.size > MAX_CV_MB * 1024 * 1024) { setError(`El CV pesa más de ${MAX_CV_MB} MB. Probá con un PDF más liviano o una foto.`); return; }
        setCv({ dataUrl: await leerComoDataUrl(file), nombre: file.name, tipo: file.type || "application/pdf" });
      }
    } catch {
      setError("No se pudo leer el CV. Probá con otro archivo.");
    } finally {
      setProcesando(false);
    }
  }
  async function elegirFotos(e) {
    const files = Array.from(e.target.files || []).filter((x) => x.type.startsWith("image/"));
    e.target.value = "";
    if (!files.length) return;
    setProcesando(true);
    try {
      const nuevas = [];
      for (const file of files.slice(0, MAX_FOTOS - fotos.length)) nuevas.push(await fotoLiviana(file, 1280, 0.7));
      setFotos((prev) => [...prev, ...nuevas].slice(0, MAX_FOTOS));
    } catch {
      setError("No se pudo leer alguna de las fotos.");
    } finally {
      setProcesando(false);
    }
  }

  async function enviar(e) {
    e.preventDefault();
    setError("");
    const elegidos = Object.entries(rubros);
    if (elegidos.length === 0) { setError("Elegí al menos un rubro en el que trabajás."); return; }
    if (elegidos.some(([, cat]) => !cat)) { setError("Indicá tu categoría (ayudante, oficial, etc.) en cada rubro que marcaste."); return; }
    if ("Otro" in rubros && !otroRubro.trim()) { setError("Contanos cuál es el otro rubro."); return; }
    if (!f.consentimiento) { setError("Para enviar tus datos tenés que aceptar que los guardemos."); return; }
    const payload = {
      ...f,
      edad: f.edad || null,
      anios_experiencia: f.anios_experiencia || null,
      rubros: elegidos.map(([rubro, categoria]) => ({ rubro: rubro === "Otro" ? `Otro: ${otroRubro.trim()}` : rubro, categoria })),
      cv_archivo: cv?.dataUrl || null, cv_nombre: cv?.nombre || null, cv_tipo: cv?.tipo || null,
      fotos,
    };
    setEnviando(true);
    try {
      if (supabase) {
        const { error: err } = await supabase.rpc("crear_postulacion", { p: payload });
        if (err) throw err;
      }
      setEnviado(true);
      window.scrollTo({ top: 0 });
    } catch (err) {
      setError(/Failed to fetch/i.test(err?.message || "") ? "No hay conexión. Revisá internet y probá de nuevo." : `No se pudo enviar: ${err?.message || "error"}`);
    } finally {
      setEnviando(false);
    }
  }

  const encabezado = (
    <header style={{ backgroundColor: NAVY }} className="text-white">
      <div className="mx-auto max-w-2xl px-4 pb-8 pt-4">
        <img src="/gc-logo-white.png" alt="Grupo Concretar" className="h-7 w-auto" />
        <p className="mt-6 text-xs font-semibold uppercase tracking-[0.18em] text-amber-400">Trabajá con nosotros</p>
        <h1 className="mt-1 text-2xl font-bold leading-tight sm:text-3xl" style={{ textWrap: "balance" }}>Dejanos tus datos y te llamamos cuando haya obra para vos</h1>
        <p className="mt-2 text-sm text-slate-300">Te lleva 3 minutos. El CV es opcional.</p>
      </div>
    </header>
  );

  if (enviado) {
    return (
      <div className="min-h-screen bg-slate-50" style={FUENTE}>
        {encabezado}
        <div className="mx-auto max-w-2xl px-4 py-12 text-center">
          <CheckCircle2 size={48} className="mx-auto text-emerald-600" />
          <h2 className="mt-3 text-xl font-bold text-slate-900">¡Listo, {f.nombre.trim()}! Recibimos tus datos.</h2>
          <p className="mx-auto mt-2 max-w-md text-sm text-slate-600">Quedás en nuestra lista de personal. Cuando tengamos una obra que necesite tu oficio, te vamos a contactar al {f.telefono}.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800" style={FUENTE}>
      {encabezado}
      <form onSubmit={enviar} className="mx-auto max-w-2xl space-y-4 px-4 py-5">
        <Bloque titulo="Tus datos">
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <Campo label="Nombre"><input required value={f.nombre} onChange={set("nombre")} autoComplete="given-name" className={inputCls} /></Campo>
            <Campo label="Apellido"><input required value={f.apellido} onChange={set("apellido")} autoComplete="family-name" className={inputCls} /></Campo>
            <Campo label="Teléfono / WhatsApp"><input required type="tel" value={f.telefono} onChange={set("telefono")} autoComplete="tel" placeholder="264 4123456" className={inputCls} /></Campo>
            <Campo label="DNI" opcional><input inputMode="numeric" value={f.dni} onChange={set("dni")} className={inputCls} /></Campo>
            <Campo label="Localidad o zona" opcional><input value={f.localidad} onChange={set("localidad")} placeholder="Ej: Rivadavia, Capital, Pocito" className={inputCls} /></Campo>
            <Campo label="Edad" opcional><input type="number" min="14" max="99" value={f.edad} onChange={set("edad")} className={inputCls} /></Campo>
            <div className="sm:col-span-2"><Campo label="Mail" opcional><input type="email" value={f.email} onChange={set("email")} autoComplete="email" className={inputCls} /></Campo></div>
          </div>
        </Bloque>

        <Bloque titulo="¿En qué trabajás?">
          <p className="text-sm text-slate-600">Marcá los rubros que manejás y, en cada uno, si sos ayudante, oficial, etc.</p>
          <div className="space-y-2">
            {RUBROS_POSTULANTE.map((r) => {
              const marcado = r in rubros;
              return (
                <div key={r} className={`rounded-xl border p-2.5 ${marcado ? "border-amber-400 bg-amber-50/60" : "border-slate-200"}`}>
                  <label className="flex cursor-pointer items-center gap-2.5 text-[15px]">
                    <input type="checkbox" checked={marcado} onChange={() => toggleRubro(r)} className="h-5 w-5 accent-amber-500" />
                    <span className="font-medium text-slate-800">{r}</span>
                  </label>
                  {marcado && (
                    <div className="mt-2 space-y-2 pl-7">
                      {r === "Otro" && <input value={otroRubro} onChange={(e) => setOtroRubro(e.target.value)} placeholder="¿Cuál?" className={inputCls} />}
                      <div className="flex flex-wrap gap-1.5">
                        {CATEGORIAS_POSTULANTE.map((c) => (
                          <button key={c} type="button" onClick={() => setRubros((p) => ({ ...p, [r]: c }))}
                            className={`rounded-full border px-3 py-1 text-xs font-semibold ${rubros[r] === c ? "border-slate-900 bg-slate-900 text-white" : "border-slate-300 bg-white text-slate-700"}`}>
                            {c}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </Bloque>

        <Bloque titulo="Experiencia">
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <Campo label="Años trabajando en construcción" opcional><input type="number" min="0" max="70" value={f.anios_experiencia} onChange={set("anios_experiencia")} className={inputCls} /></Campo>
            <div />
            <Campo label="¿Tenés herramientas propias?" opcional><SiNo value={f.herramientas_propias} onChange={set("herramientas_propias")} /></Campo>
            <Campo label="¿Tenés tarjeta IERIC?" opcional><SiNo value={f.tarjeta_ieric} onChange={set("tarjeta_ieric")} /></Campo>
          </div>
          <Campo label="Últimas obras o empresas donde trabajaste" opcional>
            <textarea rows={3} value={f.ultimas_obras} onChange={set("ultimas_obras")} placeholder="Ej: Constructora X (2024-2025), edificio en Capital…" className={inputCls} />
          </Campo>
        </Bloque>

        <Bloque titulo="Disponibilidad">
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <Campo label="¿Desde cuándo podés empezar?" opcional><input type="date" value={f.disponible_desde} onChange={set("disponible_desde")} className={inputCls} /></Campo>
            <Campo label="¿Cómo preferís trabajar?" opcional>
              <select value={f.modalidad} onChange={set("modalidad")} className={inputCls}>
                <option value="">Me da igual</option>
                {MODALIDADES_POSTULANTE.map((m) => <option key={m}>{m}</option>)}
              </select>
            </Campo>
          </div>
        </Bloque>

        <Bloque titulo="CV y fotos de trabajos (opcional)">
          <div className="flex flex-wrap items-center gap-2">
            <label className="flex cursor-pointer items-center gap-2 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-700">
              <FileText size={16} /> {cv ? "Cambiar CV" : "Subir CV (PDF o foto)"}
              <input type="file" accept="application/pdf,image/*,.doc,.docx" onChange={elegirCv} className="hidden" />
            </label>
            {cv && (
              <span className="flex items-center gap-1 rounded-full bg-emerald-50 px-2.5 py-1 text-xs font-medium text-emerald-800">
                {cv.nombre}
                <button type="button" onClick={() => setCv(null)} aria-label="Quitar CV"><X size={13} /></button>
              </span>
            )}
          </div>
          <div>
            <label className={`inline-flex cursor-pointer items-center gap-2 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-700 ${fotos.length >= MAX_FOTOS ? "pointer-events-none opacity-50" : ""}`}>
              <ImagePlus size={16} /> Fotos de trabajos que hiciste (hasta {MAX_FOTOS})
              <input type="file" accept="image/*" multiple onChange={elegirFotos} className="hidden" />
            </label>
            {fotos.length > 0 && (
              <div className="mt-2 flex flex-wrap gap-2">
                {fotos.map((src, i) => (
                  <div key={i} className="relative h-20 w-20 overflow-hidden rounded-lg">
                    <img src={src} alt={`Trabajo ${i + 1}`} className="h-full w-full object-cover" />
                    <button type="button" onClick={() => setFotos((p) => p.filter((_, j) => j !== i))} aria-label="Quitar foto" className="absolute right-1 top-1 rounded-full bg-black/60 p-0.5 text-white"><X size={12} /></button>
                  </div>
                ))}
              </div>
            )}
          </div>
          {procesando && <p className="flex items-center gap-1.5 text-xs text-slate-500"><Loader2 size={13} className="animate-spin" /> Preparando archivos…</p>}
        </Bloque>

        <Bloque titulo="¿Algo más que quieras contarnos?">
          <textarea rows={3} value={f.comentarios} onChange={set("comentarios")} placeholder="Opcional" className={inputCls} />
        </Bloque>

        <label className="flex cursor-pointer items-start gap-2.5 rounded-xl border border-slate-200 bg-white p-3 text-sm text-slate-700">
          <input type="checkbox" checked={f.consentimiento} onChange={(e) => setF((x) => ({ ...x, consentimiento: e.target.checked }))} className="mt-0.5 h-5 w-5 shrink-0 accent-amber-500" />
          <span>Acepto que Grupo Concretar guarde estos datos para contactarme por trabajo. Puedo pedir que los borren cuando quiera (Ley 25.326).</span>
        </label>

        {error && <p className="rounded-lg bg-rose-50 px-3 py-2 text-sm font-medium text-rose-700">{error}</p>}

        <button disabled={enviando || procesando} className="w-full rounded-xl px-4 py-3.5 text-base font-semibold text-white shadow-sm disabled:opacity-60" style={{ backgroundColor: NAVY }}>
          {enviando ? "Enviando…" : "Enviar mis datos"}
        </button>
        <p className="pb-6 text-center text-xs text-slate-400">Tus datos solo los ve Grupo Concretar.</p>
      </form>
    </div>
  );
}
