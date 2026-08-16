import { useState, useEffect } from "react";
import {
  LayoutDashboard, Building2, Users, ClipboardCheck, Wrench,
  ShoppingCart, Receipt, Plus, MapPin, TrendingUp, TrendingDown, X, AlertTriangle, CheckCircle2,
  Database, Loader2, RefreshCw, DollarSign, Check, Menu
} from "lucide-react";
import {
  ResponsiveContainer, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend
} from "recharts";

const fmtARS = (n) =>
  new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 }).format(n || 0);

const ESTADOS_HERRAMIENTA = ["Disponible", "En uso", "Mantenimiento", "Perdida"];
const ESTADOS_OC = ["Pendiente", "Requiere aprobación", "Aprobada", "Recibida"];
const ESTADOS_FACTURA = ["Pendiente", "Pagada"];
const CATEGORIAS_GASTO = ["Materiales", "Mano de obra", "Equipos", "Otros"];
const CATEGORIAS_HERRAMIENTA = ["Eléctrica", "Manual", "Estructura", "Medición", "Seguridad", "Otro"];
const CATEGORIAS_PERSONAL = ["Oficial Especializado", "Oficial", "Medio Oficial", "Ayudante", "Gerente", "HyS", "Recursos Humanos", "Capataz", "Logística"];
const ESTADOS_PERSONAL = ["Activo", "Licencia", "Baja"];
const ESTADOS_ASISTENCIA = ["Cargada", "Aprobada"];
const ESTADOS_LIQUIDACION = ["Pendiente", "Liquidada"];
const UMBRAL_APROBACION_OC = 3000000;
const DESVIO_ALERTA_PCT = 10;
const DESVIO_DANGER_PCT = 20;

// Roles que pueden "iniciar sesión" (simulado hasta que armemos el login real)
const ROLES = ["Gerente", "Recursos Humanos", "HyS", "Capataz", "Otro (sin acceso)"];
const ROLES_ALTA_PERSONAL = ["Gerente", "Recursos Humanos", "HyS", "Capataz"];
const ROLES_EDITAR_PERSONAL = ["Gerente", "Recursos Humanos"];
const ROLES_EDITAR_COSTOS = ["Gerente", "Recursos Humanos"];

// ============================================================
// CONFIGURACIÓN DE SUPABASE
// Reemplazá estos dos valores por los de tu proyecto:
// Project Settings > API > Project URL / anon public key
// ============================================================
const SUPABASE_URL = "https://ijkthelpxsyjicanduaj.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_YTPsdNI7QefbSm6YkDiaaA_1NCjtXnV";

const isSupabaseConfigured =
  SUPABASE_URL.startsWith("http") &&
  !SUPABASE_URL.includes("TU-PROYECTO") &&
  SUPABASE_ANON_KEY &&
  !SUPABASE_ANON_KEY.includes("TU_ANON_KEY");

const toCamel = (s) => s.replace(/_([a-z])/g, (_, c) => c.toUpperCase());
const toSnake = (s) => s.replace(/[A-Z]/g, (c) => "_" + c.toLowerCase());
const rowToCamel = (row) => {
  const out = {};
  for (const k in row) out[toCamel(k)] = row[k];
  return out;
};
const objToSnake = (obj) => {
  const out = {};
  for (const k in obj) out[toSnake(k)] = obj[k];
  return out;
};

async function sbSelect(table) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?select=*&order=id.asc`, {
    headers: { apikey: SUPABASE_ANON_KEY },
  });
  if (!res.ok) throw new Error(`No se pudo leer "${table}" (HTTP ${res.status})`);
  const rows = await res.json();
  return rows.map(rowToCamel);
}

async function sbInsert(table, obj) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
    method: "POST",
    headers: {
      apikey: SUPABASE_ANON_KEY,
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify(objToSnake(obj)),
  });
  if (!res.ok) throw new Error(`No se pudo guardar en "${table}" (HTTP ${res.status})`);
  const rows = await res.json();
  return rowToCamel(rows[0]);
}

async function sbUpdate(table, id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?id=eq.${id}`, {
    method: "PATCH",
    headers: {
      apikey: SUPABASE_ANON_KEY,
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: JSON.stringify(objToSnake(patch)),
  });
  if (!res.ok) throw new Error(`No se pudo actualizar "${table}" (HTTP ${res.status})`);
  const rows = await res.json();
  return rowToCamel(rows[0]);
}

async function sbDelete(table, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?id=eq.${id}`, {
    method: "DELETE",
    headers: { apikey: SUPABASE_ANON_KEY },
  });
  if (!res.ok) throw new Error(`No se pudo eliminar en "${table}" (HTTP ${res.status})`);
}

function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

const BADGE_STYLES = {
  Disponible: "border-emerald-600 text-emerald-700",
  "En uso": "border-amber-600 text-amber-700",
  Mantenimiento: "border-slate-400 text-slate-500",
  Perdida: "border-rose-600 text-rose-700",
  Pendiente: "border-slate-400 text-slate-500",
  "Requiere aprobación": "border-rose-600 text-rose-700",
  Aprobada: "border-amber-600 text-amber-700",
  Recibida: "border-emerald-600 text-emerald-700",
  Pagada: "border-emerald-600 text-emerald-700",
  "En curso": "border-amber-600 text-amber-700",
  Finalizada: "border-emerald-600 text-emerald-700",
  Activo: "border-emerald-600 text-emerald-700",
  Licencia: "border-amber-600 text-amber-700",
  Baja: "border-rose-600 text-rose-700",
  Cargada: "border-slate-400 text-slate-500",
  Liquidada: "border-emerald-600 text-emerald-700",
};

function Badge({ estado }) {
  return (
    <span className={`inline-block rounded-full border-2 px-2 py-0.5 text-[10px] font-bold uppercase tracking-widest whitespace-nowrap ${BADGE_STYLES[estado] || "border-slate-400 text-slate-500"}`}>
      {estado}
    </span>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      {children}
    </label>
  );
}

const inputCls = "rounded-md border border-stone-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/30";
const btnGhost = "rounded-md border border-slate-300 px-2.5 py-1 text-xs font-semibold text-slate-700 hover:bg-stone-100";
const btnGhostDanger = "rounded-md border border-rose-300 px-2.5 py-1 text-xs font-semibold text-rose-600 hover:bg-rose-50";

function PhotoInput({ label, value, onChange }) {
  return (
    <div className="flex flex-col gap-1 text-sm">
      <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <div className="flex items-center gap-3">
        {value ? (
          <img src={value} alt={label} className="h-14 w-14 rounded-md border border-stone-300 object-cover" />
        ) : (
          <div className="flex h-14 w-14 items-center justify-center rounded-md border border-dashed border-stone-300 text-center text-[9px] text-slate-400">Sin foto</div>
        )}
        <div className="flex flex-col gap-1">
          <input
            type="file"
            accept="image/*"
            onChange={async (e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              try {
                const dataUrl = await readFileAsDataURL(file);
                onChange(dataUrl);
              } catch {
                alert("No se pudo leer la imagen.");
              }
            }}
            className="w-40 text-xs text-slate-600"
          />
          {value && (
            <button type="button" onClick={() => onChange(null)} className="text-left text-xs text-rose-600 hover:underline">
              Quitar
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function Panel({ title, action, children }) {
  return (
    <div className="rounded-lg border border-stone-200 bg-white shadow-sm">
      <div className="flex items-center justify-between border-b border-stone-200 px-5 py-3">
        <h3 className="text-sm font-bold uppercase tracking-wide text-slate-700">{title}</h3>
        {action}
      </div>
      <div className="p-5">{children}</div>
    </div>
  );
}

function monthsBetween(d1, d2) {
  return (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
}

export default function ConcretarApp() {
  const [tab, setTab] = useState("dashboard");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  const DEMO_OBRAS = [
    { id: 1, nombre: "Edificio Belgrano 450", cliente: "Consorcio Belgrano SA", presupuesto: 85000000, meses: 10, inicio: "2026-02-01", estado: "En curso" },
    { id: 2, nombre: "Casa Quinta Yerba Buena", cliente: "Fam. Ledesma", presupuesto: 32000000, meses: 6, inicio: "2026-05-01", estado: "En curso" },
  ];
  const DEMO_PERSONAL = [
    { id: 1, nombreCompleto: "Facundo C", dni: "", telefono: "", categoria: "Ayudante", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null },
    { id: 2, nombreCompleto: "Eduardo Sr", dni: "", telefono: "", categoria: "Oficial", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null },
    { id: 3, nombreCompleto: "Daniel Tello", dni: "", telefono: "", categoria: "Oficial Especializado", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null },
    { id: 4, nombreCompleto: "Pablo Robles", dni: "", telefono: "", categoria: "Gerente", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null },
    { id: 5, nombreCompleto: "Pepito Chespirito", dni: "", telefono: "", categoria: "Ayudante", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null },
    { id: 6, nombreCompleto: "Emi Perez", dni: "", telefono: "", categoria: "Logística", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null },
  ];
  const DEMO_COSTOS = CATEGORIAS_PERSONAL.map((cat, i) => ({ id: i + 1, categoria: cat, costoHora: null }));

  const DEMO_ASISTENCIA = [
    { id: 1, fecha: "2026-08-03", obraId: 1, nombre: "Pablo Robles", horas: 8, cargadoPor: "Emi Perez", estado: "Aprobada", liquidacion: "Liquidada" },
    { id: 2, fecha: "2026-08-03", obraId: 1, nombre: "Eduardo Sr", horas: 9, cargadoPor: "Emi Perez", estado: "Aprobada", liquidacion: "Liquidada" },
    { id: 3, fecha: "2026-08-04", obraId: 1, nombre: "Daniel Tello", horas: 8, cargadoPor: "Emi Perez", estado: "Aprobada", liquidacion: "Pendiente" },
    { id: 4, fecha: "2026-08-05", obraId: 2, nombre: "Facundo C", horas: 6, cargadoPor: "Pablo Robles", estado: "Cargada", liquidacion: "Pendiente" },
    { id: 5, fecha: "2026-08-05", obraId: 2, nombre: "Pepito Chespirito", horas: 8, cargadoPor: "Pablo Robles", estado: "Cargada", liquidacion: "Pendiente" },
  ];
  const DEMO_HERRAMIENTAS = [
    { id: 1, nombre: "Amoladora angular", categoria: "Eléctrica", ubicacion: "Depósito central", responsable: "-", estado: "Disponible" },
    { id: 2, nombre: "Andamio tubular (juego x6)", categoria: "Estructura", ubicacion: "Edificio Belgrano 450", responsable: "Pablo Robles", estado: "En uso" },
    { id: 3, nombre: "Rotomartillo SDS", categoria: "Eléctrica", ubicacion: "Casa Quinta Yerba Buena", responsable: "Daniel Tello", estado: "En uso" },
    { id: 4, nombre: "Nivel láser", categoria: "Medición", ubicacion: "Depósito central", responsable: "-", estado: "Mantenimiento" },
  ];
  const DEMO_OC = [
    { id: 1, fecha: "2026-07-02", obraId: 1, proveedor: "Corralón San Martín", item: "Cemento x50, hierro 8mm x200", montoEstimado: 4200000, estado: "Recibida" },
    { id: 2, fecha: "2026-08-05", obraId: 1, proveedor: "Aberturas del Norte", item: "Ventanas de aluminio (12 unid.)", montoEstimado: 6800000, estado: "Requiere aprobación" },
    { id: 3, fecha: "2026-08-10", obraId: 2, proveedor: "Corralón San Martín", item: "Bloques cerámicos x1000", montoEstimado: 1500000, estado: "Pendiente" },
  ];
  const DEMO_FACTURAS = [
    { id: 1, fecha: "2026-02-15", obraId: 1, ordenCompraId: null, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 5200000, comprobante: "A-0001-00012345", estado: "Pagada" },
    { id: 2, fecha: "2026-03-18", obraId: 1, ordenCompraId: null, proveedor: "Jornales de la semana", categoria: "Mano de obra", monto: 4800000, comprobante: "A-0001-00012400", estado: "Pagada" },
    { id: 3, fecha: "2026-04-20", obraId: 1, ordenCompraId: null, proveedor: "Hierros del Sur", categoria: "Materiales", monto: 6100000, comprobante: "A-0002-00003321", estado: "Pagada" },
    { id: 4, fecha: "2026-05-22", obraId: 1, ordenCompraId: null, proveedor: "Jornales de la semana", categoria: "Mano de obra", monto: 5300000, comprobante: "A-0001-00012551", estado: "Pagada" },
    { id: 5, fecha: "2026-06-19", obraId: 1, ordenCompraId: 1, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 4200000, comprobante: "A-0003-00009087", estado: "Pagada" },
    { id: 6, fecha: "2026-07-25", obraId: 1, ordenCompraId: 2, proveedor: "Aberturas del Norte", categoria: "Materiales", monto: 6800000, comprobante: "B-0001-00000442", estado: "Pendiente" },
    { id: 7, fecha: "2026-05-10", obraId: 2, ordenCompraId: null, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 3800000, comprobante: "A-0001-00012470", estado: "Pagada" },
    { id: 8, fecha: "2026-06-14", obraId: 2, ordenCompraId: null, proveedor: "Jornales de la semana", categoria: "Mano de obra", monto: 2600000, comprobante: "A-0001-00012600", estado: "Pagada" },
    { id: 9, fecha: "2026-07-15", obraId: 2, ordenCompraId: 3, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 1500000, comprobante: "A-0004-00001180", estado: "Pendiente" },
  ];

  const [obras, setObras] = useState(isSupabaseConfigured ? [] : DEMO_OBRAS);
  const [selectedObraId, setSelectedObraId] = useState(1);
  const [personal, setPersonal] = useState(isSupabaseConfigured ? [] : DEMO_PERSONAL);
  const [costosCategoria, setCostosCategoria] = useState(isSupabaseConfigured ? [] : DEMO_COSTOS);
  const [asistencia, setAsistencia] = useState(isSupabaseConfigured ? [] : DEMO_ASISTENCIA);
  const [herramientas, setHerramientas] = useState(isSupabaseConfigured ? [] : DEMO_HERRAMIENTAS);
  const [ordenesCompra, setOrdenesCompra] = useState(isSupabaseConfigured ? [] : DEMO_OC);
  const [comprasFacturas, setComprasFacturas] = useState(isSupabaseConfigured ? [] : DEMO_FACTURAS);

  const [dbLoading, setDbLoading] = useState(isSupabaseConfigured);
  const [dbError, setDbError] = useState(null);
  const [reloadKey, setReloadKey] = useState(0);

  useEffect(() => {
    if (!isSupabaseConfigured) return;
    setDbLoading(true);
    setDbError(null);
    (async () => {
      try {
        const [o, p, cc, a, h, oc, cf] = await Promise.all([
          sbSelect("obras"), sbSelect("personal"), sbSelect("costos_categoria"), sbSelect("asistencia"),
          sbSelect("herramientas"), sbSelect("ordenes_compra"), sbSelect("compras_facturas"),
        ]);
        setObras(o);
        setPersonal(p);
        setCostosCategoria(cc);
        setAsistencia(a);
        setHerramientas(h);
        setOrdenesCompra(oc);
        setComprasFacturas(cf);
        if (o[0]) setSelectedObraId(o[0].id);
      } catch (err) {
        setDbError(err.message);
      } finally {
        setDbLoading(false);
      }
    })();
  }, [reloadKey]);

  let nextId = 200;
  const genId = () => nextId++;

  async function addRecord(table, obj, setter) {
    if (isSupabaseConfigured) {
      try {
        const row = await sbInsert(table, obj);
        setter((prev) => [...prev, row]);
      } catch (err) {
        alert("No se pudo guardar: " + err.message);
      }
    } else {
      setter((prev) => [...prev, { ...obj, id: genId() }]);
    }
  }

  async function updateRecord(table, id, patch, setter) {
    if (isSupabaseConfigured) {
      try {
        const row = await sbUpdate(table, id, patch);
        setter((prev) => prev.map((x) => (x.id === id ? row : x)));
      } catch (err) {
        alert("No se pudo actualizar: " + err.message);
      }
    } else {
      setter((prev) => prev.map((x) => (x.id === id ? { ...x, ...patch } : x)));
    }
  }

  async function deleteRecord(table, id, setter) {
    if (!window.confirm("¿Eliminar este registro? Esta acción no se puede deshacer.")) return;
    if (isSupabaseConfigured) {
      try {
        await sbDelete(table, id);
        setter((prev) => prev.filter((x) => x.id !== id));
      } catch (err) {
        alert("No se pudo eliminar: " + err.message);
      }
    } else {
      setter((prev) => prev.filter((x) => x.id !== id));
    }
  }

  // ---------- Rol actual (simula el login hasta que armemos uno real) ----------
  const [currentRole, setCurrentRole] = useState("Gerente");
  const canCrearPersonal = ROLES_ALTA_PERSONAL.includes(currentRole);
  const canEditarPersonal = ROLES_EDITAR_PERSONAL.includes(currentRole);

  const canEditarCostos = ROLES_EDITAR_COSTOS.includes(currentRole);
  const [costoDrafts, setCostoDrafts] = useState({});
  const [savedCostoId, setSavedCostoId] = useState(null);

  function guardarCosto(costo) {
    const draft = costoDrafts[costo.id];
    const nuevoValor = draft === undefined ? costo.costoHora : (draft === "" ? null : Number(draft));
    updateRecord("costos_categoria", costo.id, { costoHora: nuevoValor }, setCostosCategoria);
    setSavedCostoId(costo.id);
    setTimeout(() => setSavedCostoId((id) => (id === costo.id ? null : id)), 1500);
  }

  const emptyPersonalForm = {
    nombreCompleto: "", dni: "", categoria: CATEGORIAS_PERSONAL[0],
    estado: "Activo", direccion: "", fechaNacimiento: "", fotoPersona: null, dniFrente: null, dniDorso: null,
  };
  const [personalForm, setPersonalForm] = useState(emptyPersonalForm);
  const [editingPersonalId, setEditingPersonalId] = useState(null);
  const pf = (key) => (val) => setPersonalForm((f) => ({ ...f, [key]: val }));

  function startEditPersonal(p) {
    setPersonalForm({
      nombreCompleto: p.nombreCompleto || "",
      dni: p.dni || "",
      categoria: p.categoria || CATEGORIAS_PERSONAL[0],
      estado: p.estado || "Activo",
      direccion: p.direccion || "",
      fechaNacimiento: p.fechaNacimiento || "",
      fotoPersona: p.fotoPersona || null,
      dniFrente: p.dniFrente || null,
      dniDorso: p.dniDorso || null,
    });
    setEditingPersonalId(p.id);
    setShowPersonalForm(true);
  }

  function cancelPersonalForm() {
    setPersonalForm(emptyPersonalForm);
    setEditingPersonalId(null);
    setShowPersonalForm(false);
  }

  function submitPersonalForm(e) {
    e.preventDefault();
    const payload = {
      nombreCompleto: personalForm.nombreCompleto,
      dni: personalForm.dni,
      categoria: personalForm.categoria,
      estado: personalForm.estado,
      direccion: personalForm.direccion,
      fechaNacimiento: personalForm.fechaNacimiento || null,
      fotoPersona: personalForm.fotoPersona,
      dniFrente: personalForm.dniFrente,
      dniDorso: personalForm.dniDorso,
    };
    if (editingPersonalId) {
      updateRecord("personal", editingPersonalId, payload, setPersonal);
    } else {
      addRecord("personal", payload, setPersonal);
    }
    cancelPersonalForm();
  }

  const NAV = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "obras", label: "Obras", icon: Building2 },
    { id: "personal", label: "Personal", icon: Users },
    { id: "asistencia", label: "Asistencia", icon: ClipboardCheck },
    { id: "herramientas", label: "Herramientas", icon: Wrench },
    { id: "ordenes", label: "Órdenes de Compra", icon: ShoppingCart },
    { id: "facturas", label: "Compras y Facturas", icon: Receipt },
    { id: "costos", label: "Costos por Categoría", icon: DollarSign },
  ];

  // ---------- Dashboard calculations ----------
  const obraSel = obras.find((o) => o.id === selectedObraId) || obras[0] || null;
  const startDate = obraSel ? new Date(obraSel.inicio) : new Date();
  const meses = obraSel ? Array.from({ length: obraSel.meses }, (_, i) => new Date(startDate.getFullYear(), startDate.getMonth() + i, 1)) : [];
  const gastosObra = obraSel ? comprasFacturas.filter((c) => c.obraId === obraSel.id).sort((a, b) => new Date(a.fecha) - new Date(b.fecha)) : [];

  const chartData = meses.map((d, i) => {
    const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const real = gastosObra.filter((g) => new Date(g.fecha) <= monthEnd).reduce((s, g) => s + g.monto, 0);
    const planificado = Math.round((obraSel.presupuesto * (i + 1)) / obraSel.meses);
    return { mes: d.toLocaleDateString("es-AR", { month: "short", year: "2-digit" }), Planificado: planificado, Real: real };
  });

  const idxActual = obraSel ? Math.min(obraSel.meses - 1, Math.max(0, monthsBetween(startDate, new Date()))) : 0;
  const puntoActual = chartData[idxActual] || { Planificado: 0, Real: 0 };
  const desvioAbs = puntoActual.Real - puntoActual.Planificado;
  const desvioPct = puntoActual.Planificado ? (desvioAbs / puntoActual.Planificado) * 100 : 0;
  const herramientasEnUso = obraSel ? herramientas.filter((h) => h.ubicacion === obraSel.nombre && h.estado === "En uso").length : 0;

  // ---------- Alertas globales ----------
  const herramientasAtencion = herramientas.filter((h) => h.estado === "Mantenimiento" || h.estado === "Perdida");
  const ocPendientesAprobacion = ordenesCompra.filter((o) => o.estado === "Requiere aprobación");
  const hayDesvioAlerta = desvioPct > DESVIO_ALERTA_PCT;
  const totalAlertas = herramientasAtencion.length + ocPendientesAprobacion.length + (hayDesvioAlerta ? 1 : 0);

  // ---------- Forms state ----------
  const [showObraForm, setShowObraForm] = useState(false);
  const [showHerrForm, setShowHerrForm] = useState(false);
  const [showOcForm, setShowOcForm] = useState(false);
  const [showFacturaForm, setShowFacturaForm] = useState(false);
  const [showPersonalForm, setShowPersonalForm] = useState(false);
  const [showAsistenciaForm, setShowAsistenciaForm] = useState(false);
  const [filtroHerr, setFiltroHerr] = useState({ ubicacion: "Todas", estado: "Todos" });

  const aprobarOC = (id) => updateRecord("ordenes_compra", id, { estado: "Aprobada" }, setOrdenesCompra);
  const recibirOC = (id) => updateRecord("ordenes_compra", id, { estado: "Recibida" }, setOrdenesCompra);

  if (dbLoading) {
    return (
      <div className="flex h-full min-h-[720px] w-full flex-col items-center justify-center gap-3 rounded-xl border border-stone-200 bg-stone-100 text-slate-600">
        <Loader2 size={28} className="animate-spin text-amber-500" />
        <div className="text-sm font-medium">Conectando con la base de datos…</div>
      </div>
    );
  }

  if (dbError) {
    return (
      <div className="flex h-full min-h-[720px] w-full flex-col items-center justify-center gap-3 rounded-xl border border-stone-200 bg-stone-100 p-8 text-center">
        <AlertTriangle size={28} className="text-rose-500" />
        <div className="text-sm font-semibold text-slate-800">No se pudo conectar a Supabase</div>
        <div className="max-w-md text-xs text-slate-500">{dbError}</div>
        <div className="max-w-md text-xs text-slate-500">Revisá que hayas corrido schema.sql, y que SUPABASE_URL / SUPABASE_ANON_KEY estén bien copiados.</div>
        <button onClick={() => setReloadKey((k) => k + 1)} className="mt-2 flex items-center gap-1 rounded-md bg-slate-900 px-3 py-2 text-xs font-semibold text-white hover:bg-slate-700">
          <RefreshCw size={14} /> Reintentar
        </button>
      </div>
    );
  }

  return (
    <div className="relative flex h-full min-h-[720px] w-full overflow-hidden md:rounded-xl md:border md:border-stone-200 bg-stone-100 text-slate-800" style={{ fontFamily: "ui-sans-serif, system-ui, sans-serif" }}>
      {/* Barra superior solo en celular */}
      <div className="fixed inset-x-0 top-0 z-30 flex items-center justify-between bg-slate-900 px-4 py-3 text-slate-100 md:hidden">
        <button onClick={() => setMobileNavOpen(true)} aria-label="Abrir menú">
          <Menu size={22} />
        </button>
        <div className="text-sm font-bold tracking-tight">Concretar App</div>
        <div className="w-[22px]" />
      </div>

      {/* Fondo oscuro al abrir el menú en celular */}
      {mobileNavOpen && (
        <div className="fixed inset-0 z-30 bg-black/50 md:hidden" onClick={() => setMobileNavOpen(false)} />
      )}

      {/* Sidebar */}
      <aside
        className={`fixed inset-y-0 left-0 z-40 flex w-64 shrink-0 -translate-x-full flex-col bg-slate-900 text-slate-100 transition-transform duration-200 md:static md:z-auto md:w-56 md:translate-x-0 ${
          mobileNavOpen ? "translate-x-0" : ""
        }`}
      >
        <div className="flex items-center justify-between border-b border-slate-700/60 px-5 py-5">
          <div>
            <div className="text-[10px] font-bold uppercase tracking-[0.25em] text-amber-400">Concretar</div>
            <div className="text-lg font-bold tracking-tight">App de Obra</div>
          </div>
          <button onClick={() => setMobileNavOpen(false)} className="text-slate-400 md:hidden" aria-label="Cerrar menú">
            <X size={20} />
          </button>
        </div>
        <nav className="flex-1 space-y-1 overflow-y-auto px-2 py-4">
          {NAV.map((item) => (
            <button
              key={item.id}
              onClick={() => {
                setTab(item.id);
                setMobileNavOpen(false);
              }}
              className={`flex w-full items-center justify-between rounded-md px-3 py-2.5 text-left text-sm font-medium transition ${
                tab === item.id ? "bg-amber-500 text-slate-900" : "text-slate-300 hover:bg-slate-800 hover:text-white"
              }`}
            >
              <span className="flex items-center gap-3"><item.icon size={17} />{item.label}</span>
              {item.id === "dashboard" && totalAlertas > 0 && (
                <span className="flex h-5 w-5 items-center justify-center rounded-full bg-rose-500 text-[10px] font-bold text-white">{totalAlertas}</span>
              )}
            </button>
          ))}
        </nav>
        <div className="border-t border-slate-700/60 px-5 py-4">
          <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-slate-400">Ingresando como</div>
          <select
            value={currentRole}
            onChange={(e) => setCurrentRole(e.target.value)}
            className="w-full rounded-md border border-slate-700 bg-slate-800 px-2 py-1.5 text-xs text-slate-100"
          >
            {ROLES.map((r) => <option key={r}>{r}</option>)}
          </select>
          <div className="mt-1 text-[10px] text-slate-500">Simula el login hasta que armemos uno real</div>
        </div>
        <div className="border-t border-slate-700/60 px-5 py-4 text-[11px] text-slate-500">
          {isSupabaseConfigured ? (
            <span className="flex items-center gap-1.5 text-emerald-400"><Database size={12} /> Conectado a Supabase</span>
          ) : (
            <span className="flex items-center gap-1.5 text-amber-400"><Database size={12} /> Modo demo — configurá Supabase</span>
          )}
        </div>
      </aside>

      {/* Main content */}
      <main className="mt-12 flex-1 overflow-y-auto p-4 md:mt-0 md:p-8">
        {tab === "dashboard" && !obraSel && (
          <div className="flex flex-col items-center justify-center gap-3 rounded-lg border-2 border-dashed border-stone-300 bg-white p-10 text-center">
            <Building2 size={28} className="text-slate-400" />
            <div className="text-sm font-semibold text-slate-700">Todavía no hay obras cargadas</div>
            <div className="max-w-sm text-xs text-slate-500">
              {isSupabaseConfigured
                ? "La base está conectada pero la tabla \"obras\" está vacía. Revisá en Supabase (Table Editor) si corrió el schema.sql, o cargá tu primera obra desde la pestaña Obras."
                : "Cargá tu primera obra desde la pestaña Obras."}
            </div>
            <button onClick={() => setTab("obras")} className="mt-1 rounded-md bg-amber-500 px-3 py-2 text-xs font-semibold text-slate-900 hover:bg-amber-400">
              Ir a Obras
            </button>
          </div>
        )}

        {tab === "dashboard" && obraSel && (
          <div className="space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div>
                <div className="text-[11px] font-semibold uppercase tracking-widest text-amber-600">Dashboard financiero</div>
                <h2 className="text-2xl font-bold tracking-tight text-slate-900">{obraSel.nombre}</h2>
              </div>
              <select
                className={inputCls}
                value={selectedObraId}
                onChange={(e) => setSelectedObraId(Number(e.target.value))}
              >
                {obras.map((o) => (
                  <option key={o.id} value={o.id}>{o.nombre}</option>
                ))}
              </select>
            </div>

            <Panel title="Alertas">
              {totalAlertas === 0 ? (
                <div className="flex items-center gap-2 text-sm text-emerald-700"><CheckCircle2 size={16} /> Todo en orden, sin pendientes críticos.</div>
              ) : (
                <div className="space-y-2">
                  {hayDesvioAlerta && (
                    <div className={`flex items-center gap-2 rounded-md border px-3 py-2 text-sm ${desvioPct > DESVIO_DANGER_PCT ? "border-rose-300 bg-rose-50 text-rose-700" : "border-amber-300 bg-amber-50 text-amber-700"}`}>
                      <AlertTriangle size={16} />
                      {obraSel.nombre} está {desvioPct.toFixed(1)}% por encima de lo planificado a la fecha.
                    </div>
                  )}
                  {ocPendientesAprobacion.length > 0 && (
                    <div className="flex items-center gap-2 rounded-md border border-rose-300 bg-rose-50 px-3 py-2 text-sm text-rose-700">
                      <AlertTriangle size={16} />
                      {ocPendientesAprobacion.length} orden(es) de compra por encima de {fmtARS(UMBRAL_APROBACION_OC)} esperando tu aprobación.
                    </div>
                  )}
                  {herramientasAtencion.length > 0 && (
                    <div className="flex items-center gap-2 rounded-md border border-amber-300 bg-amber-50 px-3 py-2 text-sm text-amber-700">
                      <AlertTriangle size={16} />
                      {herramientasAtencion.length} herramienta(s) en mantenimiento o perdidas.
                    </div>
                  )}
                </div>
              )}
            </Panel>

            <div className="grid grid-cols-2 gap-4 md:grid-cols-4">
              <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Presupuesto</div>
                <div className="mt-1 font-mono text-lg font-bold text-slate-900">{fmtARS(obraSel.presupuesto)}</div>
              </div>
              <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Gastado a la fecha</div>
                <div className="mt-1 font-mono text-lg font-bold text-slate-900">{fmtARS(puntoActual.Real)}</div>
              </div>
              <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Desvío vs. plan</div>
                <div className={`mt-1 flex items-center gap-1 font-mono text-lg font-bold ${desvioAbs > 0 ? "text-rose-600" : "text-emerald-600"}`}>
                  {desvioAbs > 0 ? <TrendingUp size={16} /> : <TrendingDown size={16} />}
                  {desvioPct.toFixed(1)}%
                </div>
              </div>
              <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Herramientas en uso</div>
                <div className="mt-1 font-mono text-lg font-bold text-slate-900">{herramientasEnUso}</div>
              </div>
            </div>

            <Panel title="Curva de inversión — planificado vs. real">
              <div className="h-72 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={chartData} margin={{ top: 5, right: 20, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" />
                    <XAxis dataKey="mes" tick={{ fontSize: 12 }} stroke="#78716c" />
                    <YAxis tickFormatter={(v) => `${(v / 1000000).toFixed(0)}M`} tick={{ fontSize: 12 }} stroke="#78716c" />
                    <Tooltip formatter={(v) => fmtARS(v)} />
                    <Legend />
                    <Line type="monotone" dataKey="Planificado" stroke="#78716c" strokeDasharray="5 4" strokeWidth={2} dot={false} />
                    <Line type="monotone" dataKey="Real" stroke="#f59e0b" strokeWidth={3} dot={{ r: 3 }} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </Panel>
          </div>
        )}

        {tab === "obras" && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Obras</h2>
              <button onClick={() => setShowObraForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Nueva obra
              </button>
            </div>

            {showObraForm && (
              <Panel title="Añadir obra" action={<button onClick={() => setShowObraForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    addRecord("obras", {
                      nombre: f.get("nombre"),
                      cliente: f.get("cliente"),
                      presupuesto: Number(f.get("presupuesto")) || 0,
                      meses: Number(f.get("meses")) || 1,
                      inicio: f.get("inicio"),
                      estado: "En curso",
                    }, setObras);
                    e.target.reset();
                    setShowObraForm(false);
                  }}
                >
                  <Field label="Nombre de la obra"><input name="nombre" required className={inputCls} /></Field>
                  <Field label="Cliente"><input name="cliente" className={inputCls} /></Field>
                  <Field label="Presupuesto (ARS)"><input name="presupuesto" type="number" required className={inputCls} /></Field>
                  <Field label="Fecha de inicio"><input name="inicio" type="date" required className={inputCls} /></Field>
                  <Field label="Duración (meses)"><input name="meses" type="number" min="1" required className={inputCls} /></Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              {obras.map((o) => (
                <div key={o.id} className="rounded-lg border border-stone-200 bg-white p-5 shadow-sm">
                  <div className="flex items-start justify-between">
                    <div>
                      <div className="font-bold text-slate-900">{o.nombre}</div>
                      <div className="text-sm text-slate-500">{o.cliente}</div>
                    </div>
                    <Badge estado={o.estado} />
                  </div>
                  <div className="mt-3 flex justify-between text-sm">
                    <span className="text-slate-500">Presupuesto</span>
                    <span className="font-mono font-semibold">{fmtARS(o.presupuesto)}</span>
                  </div>
                  <div className="flex justify-between text-sm">
                    <span className="text-slate-500">Duración</span>
                    <span>{o.meses} meses desde {new Date(o.inicio).toLocaleDateString("es-AR")}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {tab === "personal" && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Personal</h2>
              {canCrearPersonal ? (
                <button
                  onClick={() => (showPersonalForm ? cancelPersonalForm() : setShowPersonalForm(true))}
                  className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400"
                >
                  <Plus size={16} /> Añadir elemento
                </button>
              ) : (
                <span className="text-xs text-slate-400">Tu rol no puede dar de alta personal</span>
              )}
            </div>

            <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
              Capataz, Gerentes, RRHH e HyS pueden dar de alta. Solo Gerentes y RRHH pueden editar o eliminar. El costo por hora ya no se carga por persona: se toma automáticamente de la pestaña "Costos por Categoría" según la categoría de cada uno.
            </div>

            {showPersonalForm && canCrearPersonal && (
              <Panel title={editingPersonalId ? "Editar elemento" : "Añadir elemento"} action={<button onClick={cancelPersonalForm}><X size={16} /></button>}>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitPersonalForm}>
                  <Field label="Nombre completo">
                    <input value={personalForm.nombreCompleto} onChange={(e) => pf("nombreCompleto")(e.target.value)} required className={inputCls} />
                  </Field>
                  <Field label="DNI">
                    <input value={personalForm.dni} onChange={(e) => pf("dni")(e.target.value)} className={inputCls} />
                  </Field>
                  <Field label="Categoría">
                    <select value={personalForm.categoria} onChange={(e) => pf("categoria")(e.target.value)} className={inputCls}>
                      {CATEGORIAS_PERSONAL.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label="Estado">
                    <select value={personalForm.estado} onChange={(e) => pf("estado")(e.target.value)} className={inputCls}>
                      {ESTADOS_PERSONAL.map((s) => <option key={s}>{s}</option>)}
                    </select>
                  </Field>
                  <Field label="Fecha de nacimiento">
                    <input type="date" value={personalForm.fechaNacimiento} onChange={(e) => pf("fechaNacimiento")(e.target.value)} className={inputCls} />
                  </Field>
                  <Field label="Dirección">
                    <input value={personalForm.direccion} onChange={(e) => pf("direccion")(e.target.value)} className={inputCls} />
                  </Field>
                  <PhotoInput label="Foto de la persona" value={personalForm.fotoPersona} onChange={pf("fotoPersona")} />
                  <PhotoInput label="DNI frente" value={personalForm.dniFrente} onChange={pf("dniFrente")} />
                  <PhotoInput label="DNI dorso" value={personalForm.dniDorso} onChange={pf("dniDorso")} />
                  <div className="flex items-end gap-2 md:col-span-3">
                    <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">
                      {editingPersonalId ? "Guardar cambios" : "Guardar"}
                    </button>
                    <button type="button" onClick={cancelPersonalForm} className={btnGhost}>Cancelar</button>
                  </div>
                </form>
              </Panel>
            )}

            <div className="overflow-hidden rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr>
                    <th className="px-4 py-3">Foto</th>
                    <th className="px-4 py-3">Nombre</th>
                    <th className="px-4 py-3">Categoría</th>
                    <th className="px-4 py-3">DNI</th>
                    <th className="px-4 py-3">Costo/hora</th>
                    <th className="px-4 py-3">Estado</th>
                    {canEditarPersonal && <th className="px-4 py-3">Acciones</th>}
                  </tr>
                </thead>
                <tbody>
                  {personal.map((p) => (
                    <tr key={p.id} className="border-t border-stone-100">
                      <td className="px-4 py-3">
                        {p.fotoPersona ? (
                          <img src={p.fotoPersona} alt={p.nombreCompleto} className="h-9 w-9 rounded-full border border-stone-200 object-cover" />
                        ) : (
                          <div className="flex h-9 w-9 items-center justify-center rounded-full bg-stone-100 text-[10px] font-semibold text-slate-400">
                            {(p.nombreCompleto || "?").slice(0, 1)}
                          </div>
                        )}
                      </td>
                      <td className="px-4 py-3 font-medium text-slate-900">{p.nombreCompleto}</td>
                      <td className="px-4 py-3 text-slate-600">{p.categoria}</td>
                      <td className="px-4 py-3 font-mono text-xs text-slate-500">{p.dni || "—"}</td>
                      <td className="px-4 py-3 font-mono text-slate-700">
                        {(() => {
                          const c = costosCategoria.find((x) => x.categoria === p.categoria)?.costoHora;
                          return c ? fmtARS(c) : <span className="text-slate-400">Sin definir</span>;
                        })()}
                      </td>
                      <td className="px-4 py-3"><Badge estado={p.estado} /></td>
                      {canEditarPersonal && (
                        <td className="px-4 py-3">
                          <div className="flex gap-2">
                            <button onClick={() => startEditPersonal(p)} className={btnGhost}>Editar</button>
                            <button onClick={() => deleteRecord("personal", p.id, setPersonal)} className={btnGhostDanger}>Eliminar</button>
                          </div>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab === "asistencia" && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Asistencia</h2>
              <button onClick={() => setShowAsistenciaForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Cargar asistencia
              </button>
            </div>

            {showAsistenciaForm && (
              <Panel title="Cargar asistencia" action={<button onClick={() => setShowAsistenciaForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    addRecord("asistencia", {
                      fecha: f.get("fecha"),
                      obraId: Number(f.get("obraId")),
                      nombre: f.get("nombre"),
                      horas: Number(f.get("horas")) || 0,
                      cargadoPor: f.get("cargadoPor"),
                      estado: f.get("estado"),
                      liquidacion: f.get("liquidacion"),
                    }, setAsistencia);
                    e.target.reset();
                    setShowAsistenciaForm(false);
                  }}
                >
                  <Field label="Fecha"><input name="fecha" type="date" required className={inputCls} /></Field>
                  <Field label="Obra">
                    <select name="obraId" className={inputCls}>{obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}</select>
                  </Field>
                  <Field label="Nombre y apellido">
                    <select name="nombre" className={inputCls}>{personal.map((p) => <option key={p.id}>{p.nombreCompleto}</option>)}</select>
                  </Field>
                  <Field label="Hs trabajadas"><input name="horas" type="number" defaultValue={8} required className={inputCls} /></Field>
                  <Field label="Cargado por">
                    <select name="cargadoPor" className={inputCls}>{personal.map((p) => <option key={p.id}>{p.nombreCompleto}</option>)}</select>
                  </Field>
                  <Field label="Estado">
                    <select name="estado" className={inputCls}>{ESTADOS_ASISTENCIA.map((s) => <option key={s}>{s}</option>)}</select>
                  </Field>
                  <Field label="Estado liquidación">
                    <select name="liquidacion" className={inputCls}>{ESTADOS_LIQUIDACION.map((s) => <option key={s}>{s}</option>)}</select>
                  </Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Enviar</button></div>
                </form>
              </Panel>
            )}

            <div className="overflow-hidden rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-4 py-3">Fecha</th><th className="px-4 py-3">Obra</th><th className="px-4 py-3">Nombre</th><th className="px-4 py-3">Hs</th><th className="px-4 py-3">Cargado por</th><th className="px-4 py-3">Estado</th><th className="px-4 py-3">Liquidación</th></tr>
                </thead>
                <tbody>
                  {[...asistencia].sort((a, b) => new Date(b.fecha) - new Date(a.fecha)).map((a) => {
                    const obra = obras.find((o) => o.id === a.obraId);
                    return (
                      <tr key={a.id} className="border-t border-stone-100">
                        <td className="px-4 py-3 text-slate-600">{new Date(a.fecha).toLocaleDateString("es-AR")}</td>
                        <td className="px-4 py-3 text-slate-600">{obra?.nombre}</td>
                        <td className="px-4 py-3 font-medium text-slate-900">{a.nombre}</td>
                        <td className="px-4 py-3 font-mono text-slate-700">{a.horas}</td>
                        <td className="px-4 py-3 text-slate-600">{a.cargadoPor}</td>
                        <td className="px-4 py-3"><Badge estado={a.estado} /></td>
                        <td className="px-4 py-3"><Badge estado={a.liquidacion} /></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab === "herramientas" && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Herramientas</h2>
              <button onClick={() => setShowHerrForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Nueva herramienta
              </button>
            </div>

            {herramientasAtencion.length > 0 && (
              <div className="rounded-md border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-800">
                <div className="mb-1 flex items-center gap-2 font-semibold"><AlertTriangle size={16} /> Requieren atención</div>
                <ul className="ml-6 list-disc space-y-0.5">
                  {herramientasAtencion.map((h) => <li key={h.id}>{h.nombre} — {h.estado} ({h.ubicacion})</li>)}
                </ul>
              </div>
            )}

            {showHerrForm && (
              <Panel title="Añadir herramienta" action={<button onClick={() => setShowHerrForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    addRecord("herramientas", {
                      nombre: f.get("nombre"),
                      categoria: f.get("categoria"),
                      ubicacion: f.get("ubicacion"),
                      responsable: f.get("responsable") || "-",
                      estado: f.get("estado"),
                    }, setHerramientas);
                    e.target.reset();
                    setShowHerrForm(false);
                  }}
                >
                  <Field label="Nombre"><input name="nombre" required className={inputCls} /></Field>
                  <Field label="Categoría">
                    <select name="categoria" className={inputCls}>{CATEGORIAS_HERRAMIENTA.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <Field label="Ubicación">
                    <select name="ubicacion" className={inputCls}>
                      <option>Depósito central</option>
                      {obras.map((o) => <option key={o.id}>{o.nombre}</option>)}
                    </select>
                  </Field>
                  <Field label="Responsable"><input name="responsable" className={inputCls} /></Field>
                  <Field label="Estado">
                    <select name="estado" className={inputCls}>{ESTADOS_HERRAMIENTA.map((s) => <option key={s}>{s}</option>)}</select>
                  </Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="flex flex-wrap gap-3">
              <select className={inputCls} value={filtroHerr.ubicacion} onChange={(e) => setFiltroHerr((f) => ({ ...f, ubicacion: e.target.value }))}>
                <option>Todas</option>
                <option>Depósito central</option>
                {obras.map((o) => <option key={o.id}>{o.nombre}</option>)}
              </select>
              <select className={inputCls} value={filtroHerr.estado} onChange={(e) => setFiltroHerr((f) => ({ ...f, estado: e.target.value }))}>
                <option>Todos</option>
                {ESTADOS_HERRAMIENTA.map((s) => <option key={s}>{s}</option>)}
              </select>
            </div>

            <div className="overflow-hidden rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-4 py-3">Herramienta</th><th className="px-4 py-3">Categoría</th><th className="px-4 py-3">Ubicación</th><th className="px-4 py-3">Responsable</th><th className="px-4 py-3">Estado</th></tr>
                </thead>
                <tbody>
                  {herramientas
                    .filter((h) => filtroHerr.ubicacion === "Todas" || h.ubicacion === filtroHerr.ubicacion)
                    .filter((h) => filtroHerr.estado === "Todos" || h.estado === filtroHerr.estado)
                    .map((h) => (
                      <tr key={h.id} className="border-t border-stone-100">
                        <td className="px-4 py-3 font-medium text-slate-900">{h.nombre}</td>
                        <td className="px-4 py-3 text-slate-600">{h.categoria}</td>
                        <td className="px-4 py-3 text-slate-600"><span className="inline-flex items-center gap-1"><MapPin size={13} className="text-amber-600" />{h.ubicacion}</span></td>
                        <td className="px-4 py-3 text-slate-600">{h.responsable}</td>
                        <td className="px-4 py-3"><Badge estado={h.estado} /></td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab === "ordenes" && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Órdenes de Compra</h2>
              <button onClick={() => setShowOcForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Nueva orden
              </button>
            </div>

            <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
              Las órdenes por {fmtARS(UMBRAL_APROBACION_OC)} o más quedan como "Requiere aprobación" hasta que las apruebes.
            </div>

            {showOcForm && (
              <Panel title="Añadir orden de compra" action={<button onClick={() => setShowOcForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    const monto = Number(f.get("montoEstimado")) || 0;
                    addRecord("ordenes_compra", {
                      fecha: f.get("fecha"),
                      obraId: Number(f.get("obraId")),
                      proveedor: f.get("proveedor"),
                      item: f.get("item"),
                      montoEstimado: monto,
                      estado: monto >= UMBRAL_APROBACION_OC ? "Requiere aprobación" : "Pendiente",
                    }, setOrdenesCompra);
                    e.target.reset();
                    setShowOcForm(false);
                  }}
                >
                  <Field label="Fecha"><input name="fecha" type="date" required className={inputCls} /></Field>
                  <Field label="Obra">
                    <select name="obraId" className={inputCls}>{obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}</select>
                  </Field>
                  <Field label="Proveedor"><input name="proveedor" required className={inputCls} /></Field>
                  <Field label="Ítems / detalle"><input name="item" className={inputCls} /></Field>
                  <Field label="Monto estimado (ARS)"><input name="montoEstimado" type="number" required className={inputCls} /></Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="space-y-3">
              {ordenesCompra.map((oc) => {
                const obra = obras.find((o) => o.id === oc.obraId);
                return (
                  <div key={oc.id} className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                    <div>
                      <div className="font-semibold text-slate-900">{oc.proveedor}</div>
                      <div className="text-sm text-slate-500">{oc.item}</div>
                      <div className="mt-1 text-xs text-slate-400">{obra?.nombre} · {new Date(oc.fecha).toLocaleDateString("es-AR")}</div>
                    </div>
                    <div className="flex items-center gap-3">
                      <span className="font-mono text-sm font-semibold text-slate-800">{fmtARS(oc.montoEstimado)}</span>
                      <Badge estado={oc.estado} />
                      {(oc.estado === "Pendiente" || oc.estado === "Requiere aprobación") && (
                        <button onClick={() => aprobarOC(oc.id)} className={btnGhost}>Aprobar</button>
                      )}
                      {oc.estado === "Aprobada" && (
                        <button onClick={() => recibirOC(oc.id)} className={btnGhost}>Marcar recibida</button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab === "facturas" && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Compras y Facturas</h2>
              <button onClick={() => setShowFacturaForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Cargar compra
              </button>
            </div>

            {showFacturaForm && (
              <Panel title="Cargar compra / factura" action={<button onClick={() => setShowFacturaForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    addRecord("compras_facturas", {
                      fecha: f.get("fecha"),
                      obraId: Number(f.get("obraId")),
                      ordenCompraId: f.get("ordenCompraId") ? Number(f.get("ordenCompraId")) : null,
                      proveedor: f.get("proveedor"),
                      categoria: f.get("categoria"),
                      monto: Number(f.get("monto")) || 0,
                      comprobante: f.get("comprobante"),
                      estado: f.get("estado"),
                    }, setComprasFacturas);
                    e.target.reset();
                    setShowFacturaForm(false);
                  }}
                >
                  <Field label="Fecha"><input name="fecha" type="date" required className={inputCls} /></Field>
                  <Field label="Obra">
                    <select name="obraId" className={inputCls}>{obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}</select>
                  </Field>
                  <Field label="Orden de compra (opcional)">
                    <select name="ordenCompraId" className={inputCls}>
                      <option value="">Sin orden asociada</option>
                      {ordenesCompra.map((oc) => <option key={oc.id} value={oc.id}>#{oc.id} · {oc.proveedor}</option>)}
                    </select>
                  </Field>
                  <Field label="Proveedor"><input name="proveedor" required className={inputCls} /></Field>
                  <Field label="Categoría">
                    <select name="categoria" className={inputCls}>{CATEGORIAS_GASTO.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <Field label="Monto (ARS)"><input name="monto" type="number" required className={inputCls} /></Field>
                  <Field label="N° comprobante"><input name="comprobante" className={inputCls} /></Field>
                  <Field label="Estado">
                    <select name="estado" className={inputCls}>{ESTADOS_FACTURA.map((s) => <option key={s}>{s}</option>)}</select>
                  </Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="overflow-hidden rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-4 py-3">Fecha</th><th className="px-4 py-3">Obra</th><th className="px-4 py-3">Proveedor</th><th className="px-4 py-3">Categoría</th><th className="px-4 py-3">Comprobante</th><th className="px-4 py-3">Monto</th><th className="px-4 py-3">Estado</th></tr>
                </thead>
                <tbody>
                  {[...comprasFacturas].sort((a, b) => new Date(b.fecha) - new Date(a.fecha)).map((c) => {
                    const obra = obras.find((o) => o.id === c.obraId);
                    return (
                      <tr key={c.id} className="border-t border-stone-100">
                        <td className="px-4 py-3 text-slate-600">{new Date(c.fecha).toLocaleDateString("es-AR")}</td>
                        <td className="px-4 py-3 text-slate-600">{obra?.nombre}</td>
                        <td className="px-4 py-3 font-medium text-slate-900">{c.proveedor}</td>
                        <td className="px-4 py-3 text-slate-600">{c.categoria}</td>
                        <td className="px-4 py-3 font-mono text-xs text-slate-500">{c.comprobante}</td>
                        <td className="px-4 py-3 text-right font-mono font-semibold text-slate-800">{fmtARS(c.monto)}</td>
                        <td className="px-4 py-3"><Badge estado={c.estado} /></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab === "costos" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Costos por Categoría</h2>
            <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
              Este valor se usa para calcular el costo por hora de cada persona en Personal, según su categoría.
              {!canEditarCostos && " Solo Gerente y RRHH pueden modificarlo."}
            </div>

            <div className="overflow-hidden rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr>
                    <th className="px-4 py-3">Categoría</th>
                    <th className="px-4 py-3">Costo por hora (ARS)</th>
                    {canEditarCostos && <th className="px-4 py-3"></th>}
                  </tr>
                </thead>
                <tbody>
                  {costosCategoria.map((c) => (
                    <tr key={c.id} className="border-t border-stone-100">
                      <td className="px-4 py-3 font-medium text-slate-900">{c.categoria}</td>
                      <td className="px-4 py-3">
                        {canEditarCostos ? (
                          <input
                            type="number"
                            className={`${inputCls} w-40`}
                            value={costoDrafts[c.id] ?? (c.costoHora ?? "")}
                            onChange={(e) => setCostoDrafts((d) => ({ ...d, [c.id]: e.target.value }))}
                          />
                        ) : c.costoHora ? (
                          fmtARS(c.costoHora)
                        ) : (
                          <span className="text-slate-400">Sin definir</span>
                        )}
                      </td>
                      {canEditarCostos && (
                        <td className="px-4 py-3">
                          <button onClick={() => guardarCosto(c)} className={btnGhost}>
                            {savedCostoId === c.id ? <span className="flex items-center gap-1 text-emerald-600"><Check size={14} /> Guardado</span> : "Guardar"}
                          </button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
