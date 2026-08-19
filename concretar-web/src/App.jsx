import { useState, useEffect } from "react";
import { createClient } from "@supabase/supabase-js";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  LayoutDashboard, Building2, Users, ClipboardCheck, Wrench,
  ShoppingCart, Receipt, Plus, MapPin, TrendingUp, TrendingDown, X, AlertTriangle, CheckCircle2,
  Database, Loader2, RefreshCw, DollarSign, Check, Menu, FileDown, ShieldCheck,
  Printer, HardHat, Zap, PaintRoller, Droplet, Hammer, Flame, Wallet,
  Landmark, Smartphone, Banknote
} from "lucide-react";
import {
  ResponsiveContainer, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend
} from "recharts";

const fmtARS = (n) =>
  new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 }).format(n || 0);

// Las fechas se guardan como texto "YYYY-MM-DD". Si se arman con
// `new Date("YYYY-MM-DD")` a secas, JS las interpreta como medianoche
// UTC — y en Argentina (UTC-3) eso "cae" al día anterior. Estas dos
// funciones arman/leen la fecha en horario LOCAL para evitar ese bug.
function fechaLocal(fechaStr) {
  if (!fechaStr) return null;
  const [y, m, d] = String(fechaStr).slice(0, 10).split("-").map(Number);
  return new Date(y, m - 1, d);
}
function fmtFecha(fechaStr) {
  const d = fechaLocal(fechaStr);
  return d ? d.toLocaleDateString("es-AR") : "—";
}
function hoyISO() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

const ESTADOS_HERRAMIENTA = ["Disponible", "En uso", "Mantenimiento", "Perdida"];
const ESTADOS_OC = ["Pendiente", "Requiere aprobación", "Aprobada", "Recibida"];
const ESTADOS_FACTURA = ["Pendiente", "Pagada"];
const CATEGORIAS_GASTO = ["Materiales", "Mano de obra", "Equipos", "Otros"];
const CATEGORIAS_HERRAMIENTA = ["Eléctrica", "Manual", "Estructura", "Medición", "Seguridad", "Otro"];
const CATEGORIAS_PERSONAL = ["Oficial Especializado", "Oficial", "Medio Oficial", "Ayudante", "Gerente", "HyS", "Recursos Humanos", "Capataz", "Logística"];
const TIPOS_TRABAJADOR = ["Empresa", "Tantero"];
const ESTADOS_PERSONAL = ["Activo", "Licencia", "Baja"];
const MANO_HABIL = ["Diestro", "Zurdo"];
const TIPOS_SANGRE = ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"];
const TALLES_PANTALON = ["38", "40", "42", "44", "46", "48", "50", "52", "54"];
const TALLES_CAMISA = ["S", "M", "L", "XL", "XXL"];
const TALLES_GUANTES = ["S", "M", "L", "XL"];
const TALLES_CALZADO = ["37", "38", "39", "40", "41", "42", "43", "44", "45", "46"];
const ESPECIALIDADES = ["Civil", "Metalúrgico", "Eléctrico", "Pintor", "Plomería", "Carpintero", "Hierrero"];
const ICONO_ESPECIALIDAD = {
  Civil: HardHat,
  "Metalúrgico": Wrench,
  "Eléctrico": Zap,
  Pintor: PaintRoller,
  "Plomería": Droplet,
  Carpintero: Hammer,
  Hierrero: Flame,
};
const ESTADOS_ASISTENCIA = ["Presente", "Ausente", "Tardanza"];
const UMBRAL_APROBACION_OC = 3000000;
const DESVIO_ALERTA_PCT = 10;
const DESVIO_DANGER_PCT = 20;

// Roles que pueden "iniciar sesión" (simulado hasta que armemos el login real)
const ROLES = ["Gerente", "Recursos Humanos", "HyS", "Capataz", "Contador", "Otro (sin acceso)"];
const ROLES_ALTA_PERSONAL = ["Gerente", "Recursos Humanos", "HyS", "Capataz"];
const ROLES_EDITAR_PERSONAL = ["Gerente", "Recursos Humanos"];
const ROLES_EDITAR_COSTOS = ["Gerente", "Recursos Humanos"];
const ROLES_LIQUIDACION = ["Gerente", "Contador"];
const ROLES_FINANZAS = ["Gerente", "Contador"];
const FORMALIDADES = ["Blanco", "Negro"];
const CUENTAS = ["Efectivo", "Banco", "Mercado Pago"];

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

// Cliente oficial. Se crea siempre (no rompe nada crearlo con valores
// de ejemplo); solo se usa de verdad cuando isSupabaseConfigured es true.
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

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
  const { data, error } = await supabase.from(table).select("*").order("id", { ascending: true });
  if (error) throw new Error(`No se pudo leer "${table}": ${error.message}`);
  return data.map(rowToCamel);
}

async function sbInsert(table, obj) {
  const { data, error } = await supabase.from(table).insert(objToSnake(obj)).select().single();
  if (error) throw new Error(`No se pudo guardar en "${table}": ${error.message}`);
  return rowToCamel(data);
}

async function sbUpdate(table, id, patch) {
  const { data, error } = await supabase.from(table).update(objToSnake(patch)).eq("id", id).select().single();
  if (error) throw new Error(`No se pudo actualizar "${table}": ${error.message}`);
  return rowToCamel(data);
}

async function sbDelete(table, id) {
  const { error } = await supabase.from(table).delete().eq("id", id);
  if (error) throw new Error(`No se pudo eliminar en "${table}": ${error.message}`);
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
  Pagado: "border-emerald-600 text-emerald-700",
  Blanco: "border-sky-600 text-sky-700",
  Negro: "border-slate-600 text-slate-700",
  "En curso": "border-amber-600 text-amber-700",
  Finalizada: "border-emerald-600 text-emerald-700",
  Activo: "border-emerald-600 text-emerald-700",
  Licencia: "border-amber-600 text-amber-700",
  Baja: "border-rose-600 text-rose-700",
  Presente: "border-emerald-600 text-emerald-700",
  Ausente: "border-rose-600 text-rose-700",
  Tardanza: "border-amber-600 text-amber-700",
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

function EspecialidadIcon({ especialidad, size = 13 }) {
  const IconComp = ICONO_ESPECIALIDAD[especialidad];
  if (!IconComp) return null;
  return (
    <span title={especialidad} className="inline-flex items-center text-slate-500">
      <IconComp size={size} />
    </span>
  );
}

const ICONO_CUENTA = { Efectivo: Banknote, Banco: Landmark, "Mercado Pago": Smartphone };

function CuentaIcon({ cuenta, size = 13 }) {
  const IconComp = ICONO_CUENTA[cuenta];
  if (!IconComp) return null;
  return (
    <span className="inline-flex items-center text-slate-500">
      <IconComp size={size} />
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
    { id: 1, nombre: "Facundo", apellido: "C", dni: "", telefono: "", categoria: "Ayudante", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Empresa" },
    { id: 2, nombre: "Eduardo", apellido: "Sr", dni: "", telefono: "", categoria: "Oficial", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Empresa" },
    { id: 3, nombre: "Daniel", apellido: "Tello", dni: "", telefono: "", categoria: "Oficial Especializado", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Empresa" },
    { id: 4, nombre: "Pablo", apellido: "Robles", dni: "", telefono: "", categoria: "Gerente", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Empresa" },
    { id: 5, nombre: "Pepito", apellido: "Chespirito", dni: "", telefono: "", categoria: "Ayudante", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Empresa" },
    { id: 6, nombre: "Emi", apellido: "Perez", dni: "", telefono: "", categoria: "Logística", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Empresa" },
    { id: 7, nombre: "Mario", apellido: "González", dni: "", telefono: "", categoria: "Oficial Especializado", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "Eléctrico", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Tantero" },
    { id: 8, nombre: "Raúl", apellido: "Medina", dni: "", telefono: "", categoria: "Oficial", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "Eléctrico", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Tantero" },
  ];
  const DEMO_COSTOS = CATEGORIAS_PERSONAL.map((cat, i) => ({ id: i + 1, categoria: cat, costoHora: null }));

  const DEMO_ASISTENCIA = [
    { id: 1, fecha: "2026-08-03", obraId: 1, nombre: "Pablo Robles", horas: 8, cargadoPor: "Gerente", estado: "Presente" },
    { id: 2, fecha: "2026-08-03", obraId: 1, nombre: "Eduardo Sr", horas: 9, cargadoPor: "Gerente", estado: "Presente" },
    { id: 3, fecha: "2026-08-04", obraId: 1, nombre: "Daniel Tello", horas: 8, cargadoPor: "Capataz", estado: "Tardanza" },
    { id: 4, fecha: "2026-08-05", obraId: 2, nombre: "Facundo C", horas: 6, cargadoPor: "Capataz", estado: "Presente" },
    { id: 5, fecha: "2026-08-05", obraId: 2, nombre: "Pepito Chespirito", horas: 0, cargadoPor: "Capataz", estado: "Ausente" },
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
    { id: 1, fecha: "2026-02-15", obraId: 1, ordenCompraId: null, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 5200000, comprobante: "A-0001-00012345", estado: "Pagada", formalidad: "Blanco", cuenta: "Banco" },
    { id: 2, fecha: "2026-03-18", obraId: 1, ordenCompraId: null, proveedor: "Jornales de la semana", categoria: "Mano de obra", monto: 4800000, comprobante: "A-0001-00012400", estado: "Pagada", formalidad: "Negro", cuenta: "Efectivo" },
    { id: 3, fecha: "2026-04-20", obraId: 1, ordenCompraId: null, proveedor: "Hierros del Sur", categoria: "Materiales", monto: 6100000, comprobante: "A-0002-00003321", estado: "Pagada", formalidad: "Blanco", cuenta: "Banco" },
    { id: 4, fecha: "2026-05-22", obraId: 1, ordenCompraId: null, proveedor: "Jornales de la semana", categoria: "Mano de obra", monto: 5300000, comprobante: "A-0001-00012551", estado: "Pagada", formalidad: "Negro", cuenta: "Efectivo" },
    { id: 5, fecha: "2026-06-19", obraId: 1, ordenCompraId: 1, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 4200000, comprobante: "A-0003-00009087", estado: "Pagada", formalidad: "Blanco", cuenta: "Banco" },
    { id: 6, fecha: "2026-07-25", obraId: 1, ordenCompraId: 2, proveedor: "Aberturas del Norte", categoria: "Materiales", monto: 6800000, comprobante: "B-0001-00000442", estado: "Pendiente", formalidad: "Blanco", cuenta: "Banco" },
    { id: 7, fecha: "2026-05-10", obraId: 2, ordenCompraId: null, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 3800000, comprobante: "A-0001-00012470", estado: "Pagada", formalidad: "Blanco", cuenta: "Mercado Pago" },
    { id: 8, fecha: "2026-06-14", obraId: 2, ordenCompraId: null, proveedor: "Jornales de la semana", categoria: "Mano de obra", monto: 2600000, comprobante: "A-0001-00012600", estado: "Pagada", formalidad: "Negro", cuenta: "Efectivo" },
    { id: 9, fecha: "2026-07-15", obraId: 2, ordenCompraId: 3, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 1500000, comprobante: "A-0004-00001180", estado: "Pendiente", formalidad: "Blanco", cuenta: "Banco" },
  ];

  const DEMO_INGRESOS = [
    { id: 1, fecha: "2026-02-05", obraId: 1, concepto: "Anticipo certificado 1", monto: 20000000, formalidad: "Blanco", cuenta: "Banco" },
    { id: 2, fecha: "2026-04-10", obraId: 1, concepto: "Certificado de avance 2", monto: 18000000, formalidad: "Blanco", cuenta: "Banco" },
    { id: 3, fecha: "2026-05-15", obraId: 1, concepto: "Adicional acordado con el cliente", monto: 6000000, formalidad: "Negro", cuenta: "Efectivo" },
    { id: 4, fecha: "2026-05-01", obraId: 2, concepto: "Anticipo Fam. Ledesma", monto: 12000000, formalidad: "Blanco", cuenta: "Mercado Pago" },
    { id: 5, fecha: "2026-06-20", obraId: 2, concepto: "Pago en mano acordado", monto: 4000000, formalidad: "Negro", cuenta: "Efectivo" },
  ];

  const DEMO_TANTEROS = [
    { id: 1, nombreGrupo: "Mario Electricista", obraId: 1, integrantes: [7, 8], precioTotal: 12000000 },
  ];
  const DEMO_AVANCES_TANTEROS = [
    { id: 1, tanteroId: 1, fecha: "2026-06-01", monto: 4000000, descripcion: "1er avance — cableado planta baja" },
    { id: 2, tanteroId: 1, fecha: "2026-07-10", monto: 3000000, descripcion: "2do avance — tablero principal" },
  ];

  const [obras, setObras] = useState(isSupabaseConfigured ? [] : DEMO_OBRAS);
  const [selectedObraId, setSelectedObraId] = useState(1);
  const [personal, setPersonal] = useState(isSupabaseConfigured ? [] : DEMO_PERSONAL);
  const [costosCategoria, setCostosCategoria] = useState(isSupabaseConfigured ? [] : DEMO_COSTOS);
  const [asistencia, setAsistencia] = useState(isSupabaseConfigured ? [] : DEMO_ASISTENCIA);
  const [herramientas, setHerramientas] = useState(isSupabaseConfigured ? [] : DEMO_HERRAMIENTAS);
  const [ordenesCompra, setOrdenesCompra] = useState(isSupabaseConfigured ? [] : DEMO_OC);
  const [comprasFacturas, setComprasFacturas] = useState(isSupabaseConfigured ? [] : DEMO_FACTURAS);
  const [ingresos, setIngresos] = useState(isSupabaseConfigured ? [] : DEMO_INGRESOS);
  const [tanteros, setTanteros] = useState(isSupabaseConfigured ? [] : DEMO_TANTEROS);
  const [avancesTanteros, setAvancesTanteros] = useState(isSupabaseConfigured ? [] : DEMO_AVANCES_TANTEROS);

  const [dbLoading, setDbLoading] = useState(isSupabaseConfigured);
  const [dbError, setDbError] = useState(null);
  const [reloadKey, setReloadKey] = useState(0);

  useEffect(() => {
    if (!isSupabaseConfigured) return;
    setDbLoading(true);
    setDbError(null);
    (async () => {
      try {
        const [o, p, cc, a, h, oc, cf, ing, tt, av] = await Promise.all([
          sbSelect("obras"), sbSelect("personal"), sbSelect("costos_categoria"), sbSelect("asistencia"),
          sbSelect("herramientas"), sbSelect("ordenes_compra"), sbSelect("compras_facturas"), sbSelect("ingresos"),
          sbSelect("tanteros"), sbSelect("avances_tanteros"),
        ]);
        setObras(o);
        setPersonal(p);
        setCostosCategoria(cc);
        setAsistencia(a);
        setHerramientas(h);
        setOrdenesCompra(oc);
        setComprasFacturas(cf);
        setIngresos(ing);
        setTanteros(tt);
        setAvancesTanteros(av);
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

  const [viewingPersonId, setViewingPersonId] = useState(null);
  const [modoSeleccionPdf, setModoSeleccionPdf] = useState(false);
  const [seleccionadosPdf, setSeleccionadosPdf] = useState([]);
  const toggleSeleccionPdf = (id) =>
    setSeleccionadosPdf((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  const viewingPerson = personal.find((p) => p.id === viewingPersonId) || null;

  function nombreCompletoDe(p) {
    return [p.nombre, p.apellido].filter(Boolean).join(" ") || "—";
  }

  function nombreCorto(p) {
    const n = (p.nombre || "").trim().split(/\s+/)[0] || "";
    const a = (p.apellido || "").trim().split(/\s+/)[0] || "";
    return [n, a].filter(Boolean).join(" ") || "—";
  }

  function ultimaObraDe(nombreCompleto) {
    const registros = asistencia
      .filter((a) => a.nombre === nombreCompleto)
      .sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));
    if (registros.length === 0) return null;
    return obras.find((o) => o.id === registros[0].obraId)?.nombre || null;
  }

  function formatoImagen(dataUrl) {
    const m = /^data:image\/(\w+);/.exec(dataUrl || "");
    if (!m) return "JPEG";
    const f = m[1].toUpperCase();
    return f === "JPG" ? "JPEG" : f;
  }

  function generarPdfSeguro(personas) {
    if (!personas || personas.length === 0) return;
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    const pageWidth = 210;
    const margin = 15;

    personas.forEach((p, idx) => {
      if (idx > 0) doc.addPage();
      let y = margin;

      doc.setFontSize(15);
      doc.setFont(undefined, "bold");
      doc.text("FICHA DE ALTA — SEGURO", margin, y);
      doc.setFontSize(9);
      doc.setFont(undefined, "normal");
      doc.text(new Date().toLocaleDateString("es-AR"), pageWidth - margin, y, { align: "right" });
      y += 10;

      if (p.fotoPersona) {
        try { doc.addImage(p.fotoPersona, formatoImagen(p.fotoPersona), pageWidth - margin - 32, margin + 4, 32, 32); } catch (e) {}
      }

      doc.setFontSize(13);
      doc.setFont(undefined, "bold");
      doc.text(nombreCompletoDe(p), margin, y);
      y += 9;

      doc.setFontSize(10);
      const campo = (label, valor) => {
        doc.setFont(undefined, "bold");
        doc.text(`${label}:`, margin, y);
        doc.setFont(undefined, "normal");
        doc.text(String(valor || "—"), margin + 42, y);
        y += 7;
      };

      campo("DNI", p.dni);
      campo("Categoría", p.categoria);
      campo("Fecha de nacimiento", p.fechaNacimiento ? fmtFecha(p.fechaNacimiento) : "—");
      campo("Dirección", p.direccion);
      campo("Mano hábil", p.manoHabil);
      campo("Tipo de sangre", p.tipoSangre);
      campo("Tarjeta IERIC", p.tarjetaIeric);
      campo("Estado", p.estado);
      y += 3;

      if (p.observaciones) {
        doc.setFont(undefined, "bold");
        doc.text("Observaciones:", margin, y);
        y += 6;
        doc.setFont(undefined, "normal");
        const lineas = doc.splitTextToSize(p.observaciones, pageWidth - margin * 2);
        doc.text(lineas, margin, y);
        y += lineas.length * 5 + 4;
      }

      y = Math.max(y, margin + 45) + 4;
      doc.setFont(undefined, "bold");
      doc.text("Documentación", margin, y);
      y += 4;

      const imgW = 80, imgH = 50;
      if (p.dniFrente) {
        try {
          doc.addImage(p.dniFrente, formatoImagen(p.dniFrente), margin, y, imgW, imgH);
          doc.setFontSize(8);
          doc.setFont(undefined, "normal");
          doc.text("DNI frente", margin, y + imgH + 4);
        } catch (e) {}
      }
      if (p.dniDorso) {
        try {
          doc.addImage(p.dniDorso, formatoImagen(p.dniDorso), margin + imgW + 10, y, imgW, imgH);
          doc.setFontSize(8);
          doc.setFont(undefined, "normal");
          doc.text("DNI dorso", margin + imgW + 10, y + imgH + 4);
        } catch (e) {}
      }
    });

    doc.save(`altas_seguro_${hoyISO()}.pdf`);
  }

  function generarPdfTalles(personas) {
    if (!personas || personas.length === 0) return;
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    doc.setFontSize(14);
    doc.setFont(undefined, "bold");
    doc.text("Resumen de talles — indumentaria", 15, 15);
    doc.setFontSize(9);
    doc.setFont(undefined, "normal");
    doc.text(new Date().toLocaleDateString("es-AR"), 195, 15, { align: "right" });

    autoTable(doc, {
      startY: 22,
      head: [["Nombre", "Categoría", "Pantalón", "Camisa", "Guantes", "Calzado"]],
      body: personas.map((p) => [
        nombreCompletoDe(p),
        p.categoria || "—",
        p.tallePantalon || "—",
        p.talleCamisa || "—",
        p.talleGuantes || "—",
        p.talleCalzado || "—",
      ]),
      styles: { fontSize: 9 },
      headStyles: { fillColor: [15, 23, 42] },
    });

    doc.save(`resumen_talles_${hoyISO()}.pdf`);
  }

  const emptyPersonalForm = {
    nombre: "", apellido: "", dni: "", categoria: CATEGORIAS_PERSONAL[0],
    estado: "Activo", direccion: "", fechaNacimiento: "", fotoPersona: null, dniFrente: null, dniDorso: null,
    manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "",
    especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "",
    tipoTrabajador: "Empresa",
  };
  const [personalForm, setPersonalForm] = useState(emptyPersonalForm);
  const [editingPersonalId, setEditingPersonalId] = useState(null);
  const pf = (key) => (val) => setPersonalForm((f) => ({ ...f, [key]: val }));

  function startEditPersonal(p) {
    setPersonalForm({
      nombre: p.nombre || "",
      apellido: p.apellido || "",
      dni: p.dni || "",
      categoria: p.categoria || CATEGORIAS_PERSONAL[0],
      estado: p.estado || "Activo",
      direccion: p.direccion || "",
      fechaNacimiento: p.fechaNacimiento || "",
      fotoPersona: p.fotoPersona || null,
      dniFrente: p.dniFrente || null,
      dniDorso: p.dniDorso || null,
      manoHabil: p.manoHabil || "Diestro",
      tipoSangre: p.tipoSangre || "",
      tarjetaIeric: p.tarjetaIeric || "No",
      observaciones: p.observaciones || "",
      especialidad: p.especialidad || "",
      tallePantalon: p.tallePantalon || "",
      talleCamisa: p.talleCamisa || "",
      talleGuantes: p.talleGuantes || "",
      talleCalzado: p.talleCalzado || "",
      tipoTrabajador: p.tipoTrabajador || "Empresa",
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
      nombre: personalForm.nombre,
      apellido: personalForm.apellido,
      dni: personalForm.dni,
      categoria: personalForm.categoria,
      estado: personalForm.estado,
      direccion: personalForm.direccion,
      fechaNacimiento: personalForm.fechaNacimiento || null,
      fotoPersona: personalForm.fotoPersona,
      dniFrente: personalForm.dniFrente,
      dniDorso: personalForm.dniDorso,
      manoHabil: personalForm.manoHabil,
      tipoSangre: personalForm.tipoSangre,
      tarjetaIeric: personalForm.tarjetaIeric,
      observaciones: personalForm.observaciones,
      especialidad: personalForm.especialidad,
      tallePantalon: personalForm.tallePantalon,
      talleCamisa: personalForm.talleCamisa,
      talleGuantes: personalForm.talleGuantes,
      talleCalzado: personalForm.talleCalzado,
      tipoTrabajador: personalForm.tipoTrabajador,
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
    { id: "liquidacion", label: "Liquidación", icon: Wallet },
    { id: "herramientas", label: "Herramientas", icon: Wrench },
    { id: "ordenes", label: "Órdenes de Compra", icon: ShoppingCart },
    { id: "ingresos", label: "Ingresos", icon: TrendingUp },
    { id: "facturas", label: "Compras y Facturas", icon: Receipt },
    { id: "cuentas", label: "Cuentas", icon: Landmark },
    { id: "costos", label: "Costos por Categoría", icon: DollarSign },
  ];

  // ---------- Dashboard calculations ----------
  const obraSel = obras.find((o) => o.id === selectedObraId) || obras[0] || null;
  const startDate = obraSel ? fechaLocal(obraSel.inicio) : new Date();
  const meses = obraSel ? Array.from({ length: obraSel.meses }, (_, i) => new Date(startDate.getFullYear(), startDate.getMonth() + i, 1)) : [];
  const gastosObra = obraSel ? comprasFacturas.filter((c) => c.obraId === obraSel.id).sort((a, b) => fechaLocal(a.fecha) - fechaLocal(b.fecha)) : [];

  const chartData = meses.map((d, i) => {
    const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const real = gastosObra.filter((g) => fechaLocal(g.fecha) <= monthEnd).reduce((s, g) => s + g.monto, 0);
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
  const asistenciasEditadas = asistencia.filter((a) => a.editado);
  const totalAlertas = herramientasAtencion.length + ocPendientesAprobacion.length + (hayDesvioAlerta ? 1 : 0) + asistenciasEditadas.length;

  // ---------- Forms state ----------
  const [showObraForm, setShowObraForm] = useState(false);
  const [showHerrForm, setShowHerrForm] = useState(false);
  const [showOcForm, setShowOcForm] = useState(false);
  const [showFacturaForm, setShowFacturaForm] = useState(false);
  const [showIngresoForm, setShowIngresoForm] = useState(false);
  const [showPersonalForm, setShowPersonalForm] = useState(false);
  const [showAsistenciaForm, setShowAsistenciaForm] = useState(false);
  const emptyAsistenciaForm = { fecha: hoyISO(), nombre: "", obraId: obras[0]?.id ?? "", horas: 8, estado: "Presente" };
  const [asistenciaForm, setAsistenciaForm] = useState(emptyAsistenciaForm);
  const [asistenciaSesion, setAsistenciaSesion] = useState([]);
  const asf = (key) => (val) => setAsistenciaForm((f) => ({ ...f, [key]: val }));

  function abrirCargaAsistencia() {
    setAsistenciaForm(emptyAsistenciaForm);
    setAsistenciaSesion([]);
    setShowAsistenciaForm(true);
  }

  function finalizarCargaAsistencia() {
    setShowAsistenciaForm(false);
    setAsistenciaSesion([]);
  }

  function submitAsistenciaForm(e) {
    e.preventDefault();
    if (!asistenciaForm.nombre) return;
    addRecord("asistencia", {
      fecha: asistenciaForm.fecha,
      nombre: asistenciaForm.nombre,
      obraId: Number(asistenciaForm.obraId),
      horas: Number(asistenciaForm.horas) || 0,
      estado: asistenciaForm.estado,
      cargadoPor: currentRole,
    }, setAsistencia);
    setAsistenciaSesion((prev) => [...prev, asistenciaForm.nombre]);
    // Mantiene fecha, obra, horas y estado; solo limpia el nombre para cargar a la próxima persona.
    setAsistenciaForm((f) => ({ ...f, nombre: "" }));
  }

  // ---------- Edición de asistencia (con motivo obligatorio) ----------
  const [editingAsistenciaId, setEditingAsistenciaId] = useState(null);
  const [editAsistenciaDraft, setEditAsistenciaDraft] = useState(null);
  const [motivoEdicionAsistencia, setMotivoEdicionAsistencia] = useState("");

  function startEditAsistencia(a) {
    setEditingAsistenciaId(a.id);
    setEditAsistenciaDraft({ fecha: a.fecha, nombre: a.nombre, obraId: a.obraId, horas: a.horas, estado: a.estado });
    setMotivoEdicionAsistencia("");
  }

  function cancelEditAsistencia() {
    setEditingAsistenciaId(null);
    setEditAsistenciaDraft(null);
    setMotivoEdicionAsistencia("");
  }

  function submitEditAsistencia(e) {
    e.preventDefault();
    if (!motivoEdicionAsistencia.trim()) return;
    updateRecord("asistencia", editingAsistenciaId, {
      fecha: editAsistenciaDraft.fecha,
      nombre: editAsistenciaDraft.nombre,
      obraId: Number(editAsistenciaDraft.obraId),
      horas: Number(editAsistenciaDraft.horas) || 0,
      estado: editAsistenciaDraft.estado,
      editado: true,
      motivoEdicion: motivoEdicionAsistencia.trim(),
      editadoPor: currentRole,
      fechaEdicion: new Date().toISOString(),
    }, setAsistencia);
    cancelEditAsistencia();
  }

  // ---------- Liquidación (pago de jornales) ----------
  const canVerLiquidacion = ROLES_LIQUIDACION.includes(currentRole);
  const [obraHistorialId, setObraHistorialId] = useState(obras[0]?.id ?? "");
  const [seleccionLiquidacion, setSeleccionLiquidacion] = useState([]); // claves "semana|obraId|nombre"
  const [vistaLiquidacion, setVistaLiquidacion] = useState("pendientes");

  function categoriaDe(nombreCompleto) {
    return personal.find((p) => nombreCompletoDe(p) === nombreCompleto)?.categoria || null;
  }

  function costoHoraDeCategoria(categoria) {
    return costosCategoria.find((c) => c.categoria === categoria)?.costoHora || 0;
  }

  function montoDe(a) {
    return (a.horas || 0) * costoHoraDeCategoria(categoriaDe(a.nombre));
  }

  // Lunes de la semana a la que pertenece una fecha "YYYY-MM-DD"
  function inicioSemana(fechaStr) {
    const d = fechaLocal(fechaStr);
    const dia = d.getDay(); // 0=domingo ... 6=sábado
    const diff = (dia === 0 ? -6 : 1) - dia;
    const lunes = new Date(d);
    lunes.setDate(d.getDate() + diff);
    return lunes;
  }
  function claveSemana(fechaStr) {
    const l = inicioSemana(fechaStr);
    const y = l.getFullYear();
    const m = String(l.getMonth() + 1).padStart(2, "0");
    const dd = String(l.getDate()).padStart(2, "0");
    return `${y}-${m}-${dd}`;
  }

  const pendientesTodasObras = asistencia.filter((a) => a.estadoPago !== "Pagado" && a.estado !== "Ausente" && (a.horas || 0) > 0);

  // Agrupa: semana -> obraId -> nombre -> { registros:[ids], horas, monto }
  const gruposSemana = {};
  pendientesTodasObras.forEach((a) => {
    const semanaKey = claveSemana(a.fecha);
    const obraId = a.obraId;
    const nombre = a.nombre;
    if (!gruposSemana[semanaKey]) gruposSemana[semanaKey] = {};
    if (!gruposSemana[semanaKey][obraId]) gruposSemana[semanaKey][obraId] = {};
    if (!gruposSemana[semanaKey][obraId][nombre]) gruposSemana[semanaKey][obraId][nombre] = { registros: [], horas: 0, monto: 0 };
    const g = gruposSemana[semanaKey][obraId][nombre];
    g.registros.push(a.id);
    g.horas += a.horas || 0;
    g.monto += montoDe(a);
  });
  const semanasOrdenadas = Object.keys(gruposSemana).sort((a, b) => new Date(b) - new Date(a));

  const toggleSeleccionLiquidacion = (key) =>
    setSeleccionLiquidacion((prev) => (prev.includes(key) ? prev.filter((x) => x !== key) : [...prev, key]));

  let totalSeleccionado = 0;
  const personasSeleccionadas = new Set();
  seleccionLiquidacion.forEach((key) => {
    const [semanaKey, obraId, nombre] = key.split("|");
    const g = gruposSemana[semanaKey]?.[obraId]?.[nombre];
    if (g) {
      totalSeleccionado += g.monto;
      personasSeleccionadas.add(nombre);
    }
  });

  const historialPagos = asistencia
    .filter((a) => a.obraId === Number(obraHistorialId) && a.estadoPago === "Pagado")
    .sort((a, b) => fechaLocal(b.fechaPago) - fechaLocal(a.fechaPago));
  const totalHistorico = historialPagos.reduce((s, a) => s + (a.montoAbonado || 0), 0);

  async function confirmarPago() {
    if (seleccionLiquidacion.length === 0) return;
    if (!window.confirm(`¿Confirmar el pago de ${fmtARS(totalSeleccionado)} para ${personasSeleccionadas.size} trabajador(es)?`)) return;
    const hoy = hoyISO();
    const idsAPagar = [];
    seleccionLiquidacion.forEach((key) => {
      const [semanaKey, obraId, nombre] = key.split("|");
      const registros = gruposSemana[semanaKey]?.[obraId]?.[nombre]?.registros || [];
      idsAPagar.push(...registros);
    });
    await Promise.all(
      idsAPagar.map((id) => {
        const a = asistencia.find((x) => x.id === id);
        return updateRecord("asistencia", id, { estadoPago: "Pagado", fechaPago: hoy, montoAbonado: montoDe(a) }, setAsistencia);
      })
    );
    setSeleccionLiquidacion([]);
  }

  // ---------- Tanteros (mano de obra por precio cerrado) ----------
  const [showTanteroForm, setShowTanteroForm] = useState(false);
  const emptyTanteroForm = { nombreGrupo: "", obraId: obras[0]?.id ?? "", precioTotal: "", integrantes: [] };
  const [tanteroForm, setTanteroForm] = useState(emptyTanteroForm);
  const [avanceAbiertoId, setAvanceAbiertoId] = useState(null);
  const [avanceForm, setAvanceForm] = useState({ fecha: hoyISO(), monto: "", descripcion: "" });

  const tanterosDisponibles = personal.filter((p) => p.tipoTrabajador === "Tantero");

  function toggleIntegranteTantero(id) {
    setTanteroForm((f) => ({
      ...f,
      integrantes: f.integrantes.includes(id) ? f.integrantes.filter((x) => x !== id) : [...f.integrantes, id],
    }));
  }

  function submitTanteroForm(e) {
    e.preventDefault();
    if (tanteroForm.integrantes.length === 0) {
      alert("Elegí al menos un integrante del grupo.");
      return;
    }
    addRecord("tanteros", {
      nombreGrupo: tanteroForm.nombreGrupo,
      obraId: Number(tanteroForm.obraId),
      precioTotal: Number(tanteroForm.precioTotal) || 0,
      integrantes: tanteroForm.integrantes,
    }, setTanteros);
    setTanteroForm(emptyTanteroForm);
    setShowTanteroForm(false);
  }

  function pagadoDeTantero(tanteroId) {
    return avancesTanteros.filter((a) => a.tanteroId === tanteroId).reduce((s, a) => s + (a.monto || 0), 0);
  }

  function submitAvanceForm(e, tanteroId) {
    e.preventDefault();
    addRecord("avances_tanteros", {
      tanteroId,
      fecha: avanceForm.fecha,
      monto: Number(avanceForm.monto) || 0,
      descripcion: avanceForm.descripcion,
    }, setAvancesTanteros);
    setAvanceForm({ fecha: hoyISO(), monto: "", descripcion: "" });
    setAvanceAbiertoId(null);
  }

  // ---------- Resumen de Cuentas (blanco/negro x efectivo/banco/MP) ----------
  const canVerFinanzas = ROLES_FINANZAS.includes(currentRole);

  function saldoCuenta(cuenta, formalidad) {
    const totalIngresos = ingresos
      .filter((i) => i.cuenta === cuenta && i.formalidad === formalidad)
      .reduce((s, i) => s + (i.monto || 0), 0);
    const totalEgresos = comprasFacturas
      .filter((c) => c.cuenta === cuenta && c.formalidad === formalidad)
      .reduce((s, c) => s + (c.monto || 0), 0);
    return totalIngresos - totalEgresos;
  }

  const saldosCuentas = CUENTAS.flatMap((cuenta) => FORMALIDADES.map((formalidad) => ({ cuenta, formalidad, saldo: saldoCuenta(cuenta, formalidad) })));
  const totalBlanco = FORMALIDADES[0] && CUENTAS.reduce((s, c) => s + saldoCuenta(c, "Blanco"), 0);
  const totalNegro = CUENTAS.reduce((s, c) => s + saldoCuenta(c, "Negro"), 0);

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
                  {asistenciasEditadas.length > 0 && (
                    <div className="rounded-md border border-sky-300 bg-sky-50 px-3 py-2 text-sm text-sky-800">
                      <div className="flex items-center gap-2 font-semibold">
                        <AlertTriangle size={16} />
                        {asistenciasEditadas.length} registro(s) de asistencia modificados — revisión sugerida.
                      </div>
                      <ul className="ml-6 mt-1 list-disc space-y-0.5 text-xs">
                        {asistenciasEditadas.slice(0, 5).map((a) => (
                          <li key={a.id}>
                            {a.nombre} ({fmtFecha(a.fecha)}) — {a.editadoPor}: "{a.motivoEdicion}"
                          </li>
                        ))}
                      </ul>
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
                    <span>{o.meses} meses desde {fmtFecha(o.inicio)}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {tab === "personal" && !viewingPerson && (
          <div className="space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Personal</h2>
              <div className="flex gap-2">
                {canEditarPersonal && (
                  <button
                    onClick={() => {
                      setModoSeleccionPdf((v) => !v);
                      setSeleccionadosPdf([]);
                    }}
                    className={`flex items-center gap-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                      modoSeleccionPdf ? "border-slate-400 bg-stone-100 text-slate-700" : "border-stone-300 bg-white text-slate-700 hover:bg-stone-50"
                    }`}
                  >
                    <ShieldCheck size={16} /> {modoSeleccionPdf ? "Cancelar selección" : "Reportes de personal"}
                  </button>
                )}
                {canCrearPersonal && (
                  <button
                    onClick={() => (showPersonalForm ? cancelPersonalForm() : setShowPersonalForm(true))}
                    className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400"
                  >
                    <Plus size={16} /> Añadir personal
                  </button>
                )}
              </div>
            </div>

            {!canCrearPersonal && !canEditarPersonal && (
              <span className="text-xs text-slate-400">Tu rol no puede dar de alta ni gestionar personal</span>
            )}

            {modoSeleccionPdf && (
              <div className="flex flex-wrap items-center justify-between gap-3 rounded-md border border-amber-300 bg-amber-50 px-4 py-3">
                <span className="text-sm text-amber-800">{seleccionadosPdf.length} persona(s) seleccionada(s) — tocá los nombres en la lista para elegirlas.</span>
                <div className="flex flex-wrap gap-2">
                  <button
                    disabled={seleccionadosPdf.length === 0}
                    onClick={() => generarPdfSeguro(personal.filter((p) => seleccionadosPdf.includes(p.id)))}
                    className="flex items-center gap-1 rounded-md bg-slate-900 px-3 py-2 text-sm font-semibold text-white hover:bg-slate-700 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    <FileDown size={16} /> PDF seguro
                  </button>
                  <button
                    disabled={seleccionadosPdf.length === 0}
                    onClick={() => generarPdfTalles(personal.filter((p) => seleccionadosPdf.includes(p.id)))}
                    className="flex items-center gap-1 rounded-md border border-slate-300 bg-white px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-stone-50 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    <Printer size={16} /> Resumen de talles
                  </button>
                </div>
              </div>
            )}

            <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
              Capataz, Gerentes, RRHH e HyS pueden dar de alta. Solo Gerentes y RRHH pueden editar, eliminar y generar el PDF para el seguro. El costo por hora ya no se carga por persona: se toma automáticamente de la pestaña "Costos por Categoría" según la categoría de cada uno.
            </div>

            {showPersonalForm && canCrearPersonal && (
              <Panel title={editingPersonalId ? "Editar personal" : "Añadir personal"} action={<button onClick={cancelPersonalForm}><X size={16} /></button>}>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitPersonalForm}>
                  <Field label="Nombre">
                    <input value={personalForm.nombre} onChange={(e) => pf("nombre")(e.target.value)} required className={inputCls} />
                  </Field>
                  <Field label="Apellido">
                    <input value={personalForm.apellido} onChange={(e) => pf("apellido")(e.target.value)} required className={inputCls} />
                  </Field>
                  <Field label="DNI">
                    <input value={personalForm.dni} onChange={(e) => pf("dni")(e.target.value)} className={inputCls} />
                  </Field>
                  <Field label="Categoría">
                    <select value={personalForm.categoria} onChange={(e) => pf("categoria")(e.target.value)} className={inputCls}>
                      {CATEGORIAS_PERSONAL.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label="Tipo de trabajador">
                    <select value={personalForm.tipoTrabajador} onChange={(e) => pf("tipoTrabajador")(e.target.value)} className={inputCls}>
                      {TIPOS_TRABAJADOR.map((t) => <option key={t}>{t}</option>)}
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
                  <Field label="Mano hábil">
                    <select value={personalForm.manoHabil} onChange={(e) => pf("manoHabil")(e.target.value)} className={inputCls}>
                      {MANO_HABIL.map((m) => <option key={m}>{m}</option>)}
                    </select>
                  </Field>
                  <Field label="Tipo de sangre">
                    <select value={personalForm.tipoSangre} onChange={(e) => pf("tipoSangre")(e.target.value)} className={inputCls}>
                      <option value="">Sin especificar</option>
                      {TIPOS_SANGRE.map((t) => <option key={t}>{t}</option>)}
                    </select>
                  </Field>
                  <Field label="¿Tiene tarjeta IERIC?">
                    <select value={personalForm.tarjetaIeric} onChange={(e) => pf("tarjetaIeric")(e.target.value)} className={inputCls}>
                      <option>No</option>
                      <option>Sí</option>
                    </select>
                  </Field>
                  <Field label="Especialidad en obra">
                    <select value={personalForm.especialidad} onChange={(e) => pf("especialidad")(e.target.value)} className={inputCls}>
                      <option value="">Sin especificar</option>
                      {ESPECIALIDADES.map((e) => <option key={e}>{e}</option>)}
                    </select>
                  </Field>
                  <Field label="Talle de pantalón">
                    <select value={personalForm.tallePantalon} onChange={(e) => pf("tallePantalon")(e.target.value)} className={inputCls}>
                      <option value="">Sin especificar</option>
                      {TALLES_PANTALON.map((t) => <option key={t}>{t}</option>)}
                    </select>
                  </Field>
                  <Field label="Talle de camisa">
                    <select value={personalForm.talleCamisa} onChange={(e) => pf("talleCamisa")(e.target.value)} className={inputCls}>
                      <option value="">Sin especificar</option>
                      {TALLES_CAMISA.map((t) => <option key={t}>{t}</option>)}
                    </select>
                  </Field>
                  <Field label="Talle de guantes">
                    <select value={personalForm.talleGuantes} onChange={(e) => pf("talleGuantes")(e.target.value)} className={inputCls}>
                      <option value="">Sin especificar</option>
                      {TALLES_GUANTES.map((t) => <option key={t}>{t}</option>)}
                    </select>
                  </Field>
                  <Field label="Talle de calzado">
                    <select value={personalForm.talleCalzado} onChange={(e) => pf("talleCalzado")(e.target.value)} className={inputCls}>
                      <option value="">Sin especificar</option>
                      {TALLES_CALZADO.map((t) => <option key={t}>{t}</option>)}
                    </select>
                  </Field>
                  <div className="md:col-span-3">
                    <Field label="Observaciones (alergias, lesiones previas, etc.)">
                      <textarea
                        value={personalForm.observaciones}
                        onChange={(e) => pf("observaciones")(e.target.value)}
                        rows={3}
                        placeholder="Ej: alérgico a la penicilina, lesión previa de hombro derecho..."
                        className={inputCls}
                      />
                    </Field>
                  </div>
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

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-[11px]">
                <thead className="bg-stone-50 text-[9px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr>
                    {modoSeleccionPdf && <th className="px-1.5 py-1.5"></th>}
                    <th className="px-1.5 py-1.5">Foto</th>
                    <th className="px-1.5 py-1.5">Nombre</th>
                    <th className="px-1.5 py-1.5">Categoría</th>
                    <th className="px-1.5 py-1.5">Última obra</th>
                    <th className="px-1.5 py-1.5">DNI</th>
                  </tr>
                </thead>
                <tbody>
                  {[...personal].sort((a, b) => b.id - a.id).map((p) => (
                    <tr
                      key={p.id}
                      onClick={() => modoSeleccionPdf && toggleSeleccionPdf(p.id)}
                      className={`border-t border-stone-100 ${modoSeleccionPdf ? "cursor-pointer" : ""} ${seleccionadosPdf.includes(p.id) ? "bg-amber-50" : ""}`}
                    >
                      {modoSeleccionPdf && (
                        <td className="px-1.5 py-1">
                          <input type="checkbox" checked={seleccionadosPdf.includes(p.id)} onChange={() => toggleSeleccionPdf(p.id)} className="h-3.5 w-3.5" />
                        </td>
                      )}
                      <td className="px-1.5 py-1">
                        {p.fotoPersona ? (
                          <img src={p.fotoPersona} alt={nombreCompletoDe(p)} className="h-6 w-6 rounded-full border border-stone-200 object-cover" />
                        ) : (
                          <div className="flex h-6 w-6 items-center justify-center rounded-full bg-stone-100 text-[8px] font-semibold text-slate-400">
                            {(p.nombre || "?").slice(0, 1)}
                          </div>
                        )}
                      </td>
                      <td className="px-1.5 py-1">
                        {modoSeleccionPdf ? (
                          <span className="flex items-center gap-1 font-medium text-slate-900">
                            <EspecialidadIcon especialidad={p.especialidad} />
                            {nombreCorto(p)}
                            {p.observaciones && (
                              <span title={p.observaciones}><AlertTriangle size={11} className="text-amber-500" /></span>
                            )}
                          </span>
                        ) : (
                          <button onClick={() => setViewingPersonId(p.id)} className="flex items-center gap-1 font-medium text-slate-900 underline decoration-dotted hover:text-amber-600">
                            <EspecialidadIcon especialidad={p.especialidad} />
                            {nombreCorto(p)}
                            {p.observaciones && (
                              <span title={p.observaciones}><AlertTriangle size={11} className="text-amber-500" /></span>
                            )}
                          </button>
                        )}
                      </td>
                      <td className="px-1.5 py-1 text-slate-600">{p.categoria}</td>
                      <td className="px-1.5 py-1 text-slate-600">{ultimaObraDe(nombreCompletoDe(p)) || <span className="text-slate-400">—</span>}</td>
                      <td className="px-1.5 py-1 font-mono text-slate-500">{p.dni || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab === "personal" && viewingPerson && (
          <div className="space-y-4">
            <button onClick={() => setViewingPersonId(null)} className="text-xs font-semibold text-slate-500 hover:text-slate-800">
              ← Volver a Personal
            </button>

            <div className="rounded-lg border border-stone-200 bg-white p-5 shadow-sm">
              <div className="flex flex-wrap items-center gap-4">
                {viewingPerson.fotoPersona ? (
                  <img src={viewingPerson.fotoPersona} alt={nombreCompletoDe(viewingPerson)} className="h-20 w-20 rounded-full border border-stone-200 object-cover" />
                ) : (
                  <div className="flex h-20 w-20 items-center justify-center rounded-full bg-stone-100 text-2xl font-semibold text-slate-400">
                    {(viewingPerson.nombre || "?").slice(0, 1)}
                  </div>
                )}
                <div>
                  <div className="text-xl font-bold text-slate-900">{nombreCompletoDe(viewingPerson)}</div>
                  <div className="mt-1 flex flex-wrap gap-2">
                    <Badge estado={viewingPerson.estado} />
                    <span className="text-sm text-slate-500">{viewingPerson.categoria}</span>
                  </div>
                </div>
                {canEditarPersonal && (
                  <div className="ml-auto flex gap-2">
                    <button onClick={() => generarPdfSeguro([viewingPerson])} className={btnGhost}>
                      <span className="flex items-center gap-1"><FileDown size={13} /> PDF seguro</span>
                    </button>
                    <button onClick={() => startEditPersonal(viewingPerson)} className={btnGhost}>Editar</button>
                    <button
                      onClick={() => {
                        deleteRecord("personal", viewingPerson.id, setPersonal);
                        setViewingPersonId(null);
                      }}
                      className={btnGhostDanger}
                    >
                      Eliminar
                    </button>
                  </div>
                )}
              </div>

              <div className="mt-6 grid grid-cols-1 gap-4 text-sm sm:grid-cols-2 md:grid-cols-3">
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">DNI</div><div className="font-mono text-slate-800">{viewingPerson.dni || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Costo por hora</div><div className="font-mono text-slate-800">{(() => { const c = costosCategoria.find((x) => x.categoria === viewingPerson.categoria)?.costoHora; return c ? fmtARS(c) : "Sin definir"; })()}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Última obra</div><div className="text-slate-800">{ultimaObraDe(nombreCompletoDe(viewingPerson)) || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Fecha de nacimiento</div><div className="text-slate-800">{viewingPerson.fechaNacimiento ? fmtFecha(viewingPerson.fechaNacimiento) : "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Dirección</div><div className="text-slate-800">{viewingPerson.direccion || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Mano hábil</div><div className="text-slate-800">{viewingPerson.manoHabil || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Tipo de sangre</div><div className="text-slate-800">{viewingPerson.tipoSangre || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Tarjeta IERIC</div><div className="text-slate-800">{viewingPerson.tarjetaIeric || "No"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Tipo de trabajador</div><div className="text-slate-800">{viewingPerson.tipoTrabajador || "Empresa"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Especialidad</div><div className="flex items-center gap-1.5 text-slate-800"><EspecialidadIcon especialidad={viewingPerson.especialidad} size={14} />{viewingPerson.especialidad || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Talle pantalón</div><div className="text-slate-800">{viewingPerson.tallePantalon || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Talle camisa</div><div className="text-slate-800">{viewingPerson.talleCamisa || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Talle guantes</div><div className="text-slate-800">{viewingPerson.talleGuantes || "—"}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Talle calzado</div><div className="text-slate-800">{viewingPerson.talleCalzado || "—"}</div></div>
              </div>

              {viewingPerson.observaciones && (
                <div className="mt-4 rounded-md border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800">
                  <div className="mb-1 flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide"><AlertTriangle size={13} /> Observaciones</div>
                  {viewingPerson.observaciones}
                </div>
              )}

              {(viewingPerson.dniFrente || viewingPerson.dniDorso) && (
                <div className="mt-4 flex flex-wrap gap-4">
                  {viewingPerson.dniFrente && (
                    <div>
                      <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">DNI frente</div>
                      <img src={viewingPerson.dniFrente} alt="DNI frente" className="h-24 rounded-md border border-stone-200 object-cover" />
                    </div>
                  )}
                  {viewingPerson.dniDorso && (
                    <div>
                      <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">DNI dorso</div>
                      <img src={viewingPerson.dniDorso} alt="DNI dorso" className="h-24 rounded-md border border-stone-200 object-cover" />
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        )}

        {tab === "asistencia" && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Asistencia</h2>
              <button onClick={abrirCargaAsistencia} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Cargar asistencia
              </button>
            </div>

            {showAsistenciaForm && (
              <Panel title="Cargar asistencia" action={<button onClick={finalizarCargaAsistencia}><X size={16} /></button>}>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitAsistenciaForm}>
                  <Field label="Fecha">
                    <input type="date" required value={asistenciaForm.fecha} onChange={(e) => asf("fecha")(e.target.value)} className={inputCls} />
                  </Field>
                  <Field label="Nombre y apellido">
                    <select required value={asistenciaForm.nombre} onChange={(e) => asf("nombre")(e.target.value)} className={inputCls}>
                      <option value="">-- Elegí a la persona --</option>
                      {personal.map((p) => <option key={p.id}>{nombreCompletoDe(p)}</option>)}
                    </select>
                  </Field>
                  <Field label="Centro de costo / Obra">
                    <select value={asistenciaForm.obraId} onChange={(e) => asf("obraId")(e.target.value)} className={inputCls}>
                      {obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                    </select>
                  </Field>
                  <Field label="Hs trabajadas">
                    <input type="number" required value={asistenciaForm.horas} onChange={(e) => asf("horas")(e.target.value)} className={inputCls} />
                  </Field>
                  <Field label="Estado">
                    <select value={asistenciaForm.estado} onChange={(e) => asf("estado")(e.target.value)} className={inputCls}>
                      {ESTADOS_ASISTENCIA.map((s) => <option key={s}>{s}</option>)}
                    </select>
                  </Field>
                  <div className="flex flex-col justify-end gap-1">
                    <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Cargado por</span>
                    <span className="rounded-md border border-stone-200 bg-stone-50 px-3 py-2 text-sm text-slate-500">{currentRole} (vos)</span>
                  </div>
                  <div className="flex items-end gap-2 md:col-span-3">
                    <button type="submit" className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">
                      Confirmar y cargar siguiente
                    </button>
                    <button type="button" onClick={finalizarCargaAsistencia} className={btnGhost}>Finalizar</button>
                  </div>
                </form>

                {asistenciaSesion.length > 0 && (
                  <div className="mt-4 rounded-md border border-emerald-200 bg-emerald-50 p-3 text-sm text-emerald-800">
                    <div className="mb-1 flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide"><CheckCircle2 size={13} /> Cargados en esta tanda ({asistenciaSesion.length})</div>
                    {asistenciaSesion.join(", ")}
                  </div>
                )}

                <div className="mt-3 text-[11px] text-slate-400">
                  La fecha, el centro de costo, las horas y el estado quedan fijos entre carga y carga — solo cambiá el nombre para ir sumando a toda la cuadrilla rápido. "Cargado por" se completa solo con tu rol actual; cuando armemos el login real, va a quedar el nombre de quien inició sesión.
                </div>
              </Panel>
            )}

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-4 py-3">Fecha</th><th className="px-4 py-3">Nombre</th><th className="px-4 py-3">Centro de costo</th><th className="px-4 py-3">Hs</th><th className="px-4 py-3">Estado</th><th className="px-4 py-3">Cargado por</th><th className="px-4 py-3"></th></tr>
                </thead>
                <tbody>
                  {[...asistencia].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((a) => {
                    const obra = obras.find((o) => o.id === a.obraId);
                    return (
                      <tr key={a.id} className="border-t border-stone-100">
                        <td className="px-4 py-3 text-slate-600">{fmtFecha(a.fecha)}</td>
                        <td className="px-4 py-3 font-medium text-slate-900">
                          <span className="flex items-center gap-1.5">
                            {a.nombre}
                            {a.editado && <span title={`Editado por ${a.editadoPor}: ${a.motivoEdicion}`}><AlertTriangle size={12} className="text-sky-500" /></span>}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-slate-600">{obra?.nombre}</td>
                        <td className="px-4 py-3 font-mono text-slate-700">{a.horas}</td>
                        <td className="px-4 py-3"><Badge estado={a.estado} /></td>
                        <td className="px-4 py-3 text-slate-500">{a.cargadoPor}</td>
                        <td className="px-4 py-3"><button onClick={() => startEditAsistencia(a)} className={btnGhost}>Editar</button></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {editingAsistenciaId && editAsistenciaDraft && (
              <Panel title="Editar asistencia" action={<button onClick={cancelEditAsistencia}><X size={16} /></button>}>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitEditAsistencia}>
                  <Field label="Fecha">
                    <input type="date" required value={editAsistenciaDraft.fecha} onChange={(e) => setEditAsistenciaDraft((d) => ({ ...d, fecha: e.target.value }))} className={inputCls} />
                  </Field>
                  <Field label="Nombre y apellido">
                    <select required value={editAsistenciaDraft.nombre} onChange={(e) => setEditAsistenciaDraft((d) => ({ ...d, nombre: e.target.value }))} className={inputCls}>
                      {personal.map((p) => <option key={p.id}>{nombreCompletoDe(p)}</option>)}
                    </select>
                  </Field>
                  <Field label="Centro de costo / Obra">
                    <select value={editAsistenciaDraft.obraId} onChange={(e) => setEditAsistenciaDraft((d) => ({ ...d, obraId: e.target.value }))} className={inputCls}>
                      {obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                    </select>
                  </Field>
                  <Field label="Hs trabajadas">
                    <input type="number" required value={editAsistenciaDraft.horas} onChange={(e) => setEditAsistenciaDraft((d) => ({ ...d, horas: e.target.value }))} className={inputCls} />
                  </Field>
                  <Field label="Estado">
                    <select value={editAsistenciaDraft.estado} onChange={(e) => setEditAsistenciaDraft((d) => ({ ...d, estado: e.target.value }))} className={inputCls}>
                      {ESTADOS_ASISTENCIA.map((s) => <option key={s}>{s}</option>)}
                    </select>
                  </Field>
                  <div className="md:col-span-3">
                    <Field label="Motivo de la modificación (obligatorio)">
                      <textarea
                        required
                        value={motivoEdicionAsistencia}
                        onChange={(e) => setMotivoEdicionAsistencia(e.target.value)}
                        rows={2}
                        placeholder="Ej: me equivoqué de persona, corrección de horas cargadas..."
                        className={inputCls}
                      />
                    </Field>
                  </div>
                  <div className="flex items-end gap-2 md:col-span-3">
                    <button type="submit" className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar corrección</button>
                    <button type="button" onClick={cancelEditAsistencia} className={btnGhost}>Cancelar</button>
                  </div>
                </form>
                <div className="mt-3 text-[11px] text-slate-400">
                  Esta corrección va a quedar visible para los gerentes en el panel de Alertas del Dashboard, junto con el motivo.
                </div>
              </Panel>
            )}
          </div>
        )}

        {tab === "liquidacion" && !canVerLiquidacion && (
          <div className="flex flex-col items-center justify-center gap-3 rounded-lg border-2 border-dashed border-stone-300 bg-white p-10 text-center">
            <Wallet size={28} className="text-slate-400" />
            <div className="text-sm font-semibold text-slate-700">Sección restringida</div>
            <div className="max-w-sm text-xs text-slate-500">Solo Gerente y Contador pueden ver y gestionar los pagos de jornales.</div>
          </div>
        )}

        {tab === "liquidacion" && canVerLiquidacion && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Liquidación</h2>

            <div className="flex flex-wrap gap-2">
              <button
                onClick={() => setVistaLiquidacion("pendientes")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaLiquidacion === "pendientes" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Pendientes de pago
              </button>
              <button
                onClick={() => setVistaLiquidacion("tanteros")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaLiquidacion === "tanteros" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Tanteros
              </button>
              <button
                onClick={() => setVistaLiquidacion("historial")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaLiquidacion === "historial" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Historial de la obra
              </button>
            </div>

            {vistaLiquidacion === "pendientes" && (
              <>
                {seleccionLiquidacion.length > 0 && (
                  <div className="sticky top-0 z-10 flex flex-wrap items-center justify-between gap-3 rounded-md border border-amber-300 bg-amber-50 px-4 py-3">
                    <div className="text-sm text-amber-800">
                      {seleccionLiquidacion.length} trabajador-semana seleccionados — {personasSeleccionadas.size} persona(s) — total <strong>{fmtARS(totalSeleccionado)}</strong>
                    </div>
                    <button onClick={confirmarPago} className="flex items-center gap-1 rounded-md bg-slate-900 px-3 py-2 text-sm font-semibold text-white hover:bg-slate-700">
                      <Check size={16} /> Confirmar pago
                    </button>
                  </div>
                )}

                {semanasOrdenadas.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    No hay días pendientes de pago en ninguna obra. 🎉
                  </div>
                ) : (
                  semanasOrdenadas.map((semanaKey) => {
                    const obrasDeSemana = gruposSemana[semanaKey];
                    const totalSemana = Object.values(obrasDeSemana).reduce(
                      (s, trabajadores) => s + Object.values(trabajadores).reduce((s2, t) => s2 + t.monto, 0),
                      0
                    );
                    return (
                      <Panel
                        key={semanaKey}
                        title={`Semana del ${fmtFecha(semanaKey)}`}
                        action={<span className="font-mono text-sm font-bold text-slate-800">{fmtARS(totalSemana)}</span>}
                      >
                        <div className="space-y-5">
                          {Object.keys(obrasDeSemana).map((obraId) => {
                            const obra = obras.find((o) => o.id === Number(obraId));
                            const trabajadores = obrasDeSemana[obraId];
                            const totalObra = Object.values(trabajadores).reduce((s, t) => s + t.monto, 0);
                            return (
                              <div key={obraId}>
                                <div className="mb-2 flex items-center justify-between border-b border-stone-200 pb-1.5">
                                  <span className="flex items-center gap-1.5 text-sm font-semibold text-slate-700"><Building2 size={14} className="text-amber-600" />{obra?.nombre || "Obra"}</span>
                                  <span className="font-mono text-sm font-semibold text-slate-700">{fmtARS(totalObra)}</span>
                                </div>
                                <div className="space-y-1">
                                  {Object.keys(trabajadores).map((nombre) => {
                                    const t = trabajadores[nombre];
                                    const key = `${semanaKey}|${obraId}|${nombre}`;
                                    const seleccionado = seleccionLiquidacion.includes(key);
                                    return (
                                      <div
                                        key={key}
                                        onClick={() => toggleSeleccionLiquidacion(key)}
                                        className={`flex cursor-pointer items-center justify-between rounded-md px-3 py-2 text-sm ${seleccionado ? "bg-amber-50" : "hover:bg-stone-50"}`}
                                      >
                                        <span className="flex items-center gap-2">
                                          <input type="checkbox" checked={seleccionado} onChange={() => toggleSeleccionLiquidacion(key)} className="h-3.5 w-3.5" />
                                          <span className="font-medium text-slate-900">{nombre}</span>
                                          <span className="text-xs text-slate-400">({t.horas} hs)</span>
                                        </span>
                                        <span className="font-mono text-slate-800">{fmtARS(t.monto)}</span>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </Panel>
                    );
                  })
                )}
                <div className="text-[11px] text-slate-400">
                  Cada línea es el total de esa persona, en esa obra, para toda la semana (lunes a domingo). Tocá una fila para incluirla en el pago — no hace falta elegir día por día.
                </div>
              </>
            )}

            {vistaLiquidacion === "tanteros" && (
              <>
                <div className="flex items-center justify-end">
                  <button onClick={() => setShowTanteroForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                    <Plus size={16} /> Nuevo grupo
                  </button>
                </div>

                <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
                  Para que alguien aparezca acá, primero marcalo como "Tantero" en su ficha de Personal. Estos grupos no cobran por hora — tienen un precio cerrado por el trabajo, y vos cargás los avances que les vas pagando.
                </div>

                {showTanteroForm && (
                  <Panel title="Nuevo grupo de tanteros" action={<button onClick={() => setShowTanteroForm(false)}><X size={16} /></button>}>
                    <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitTanteroForm}>
                      <Field label="Nombre del grupo">
                        <input value={tanteroForm.nombreGrupo} onChange={(e) => setTanteroForm((f) => ({ ...f, nombreGrupo: e.target.value }))} required placeholder="Ej: Mario Electricista" className={inputCls} />
                      </Field>
                      <Field label="Obra">
                        <select value={tanteroForm.obraId} onChange={(e) => setTanteroForm((f) => ({ ...f, obraId: e.target.value }))} className={inputCls}>
                          {obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                        </select>
                      </Field>
                      <Field label="Precio cerrado (ARS)">
                        <input type="number" value={tanteroForm.precioTotal} onChange={(e) => setTanteroForm((f) => ({ ...f, precioTotal: e.target.value }))} required className={inputCls} />
                      </Field>
                      <div className="md:col-span-3">
                        <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Integrantes del grupo</div>
                        {tanterosDisponibles.length === 0 ? (
                          <div className="rounded-md border border-dashed border-stone-300 p-3 text-xs text-slate-500">
                            Todavía no hay nadie marcado como "Tantero" en Personal. Andá a Personal, editá a la persona y elegí "Tantero" en Tipo de trabajador.
                          </div>
                        ) : (
                          <div className="flex flex-wrap gap-2">
                            {tanterosDisponibles.map((p) => (
                              <button
                                type="button"
                                key={p.id}
                                onClick={() => toggleIntegranteTantero(p.id)}
                                className={`rounded-full border px-3 py-1 text-xs font-medium ${
                                  tanteroForm.integrantes.includes(p.id) ? "border-amber-500 bg-amber-50 text-amber-800" : "border-stone-300 bg-white text-slate-600 hover:bg-stone-50"
                                }`}
                              >
                                {nombreCompletoDe(p)}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                      <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar grupo</button></div>
                    </form>
                  </Panel>
                )}

                {tanteros.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    Todavía no hay grupos de tanteros cargados.
                  </div>
                ) : (
                  <div className="space-y-4">
                    {tanteros.map((t) => {
                      const obra = obras.find((o) => o.id === t.obraId);
                      const pagado = pagadoDeTantero(t.id);
                      const saldo = (t.precioTotal || 0) - pagado;
                      const integrantesNombres = (t.integrantes || []).map((id) => {
                        const p = personal.find((x) => x.id === id);
                        return p ? nombreCompletoDe(p) : null;
                      }).filter(Boolean);
                      const avancesGrupo = avancesTanteros.filter((a) => a.tanteroId === t.id).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));
                      return (
                        <div key={t.id} className="rounded-lg border border-stone-200 bg-white p-5 shadow-sm">
                          <div className="flex flex-wrap items-start justify-between gap-3">
                            <div>
                              <div className="font-bold text-slate-900">{t.nombreGrupo}</div>
                              <div className="text-sm text-slate-500">{obra?.nombre} · {integrantesNombres.length} integrante(s): {integrantesNombres.join(", ")}</div>
                            </div>
                            <button onClick={() => setAvanceAbiertoId(avanceAbiertoId === t.id ? null : t.id)} className={btnGhost}>
                              {avanceAbiertoId === t.id ? "Cancelar" : "Cargar avance"}
                            </button>
                          </div>

                          <div className="mt-3 grid grid-cols-3 gap-3 text-sm">
                            <div>
                              <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Precio cerrado</div>
                              <div className="font-mono font-semibold text-slate-900">{fmtARS(t.precioTotal)}</div>
                            </div>
                            <div>
                              <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Pagado</div>
                              <div className="font-mono font-semibold text-emerald-700">{fmtARS(pagado)}</div>
                            </div>
                            <div>
                              <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Saldo</div>
                              <div className={`font-mono font-semibold ${saldo > 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(saldo)}</div>
                            </div>
                          </div>

                          {avanceAbiertoId === t.id && (
                            <form
                              className="mt-4 grid grid-cols-1 gap-3 rounded-md border border-stone-200 bg-stone-50 p-3 md:grid-cols-4"
                              onSubmit={(e) => submitAvanceForm(e, t.id)}
                            >
                              <Field label="Fecha">
                                <input type="date" required value={avanceForm.fecha} onChange={(e) => setAvanceForm((f) => ({ ...f, fecha: e.target.value }))} className={inputCls} />
                              </Field>
                              <Field label="Monto (ARS)">
                                <input type="number" required value={avanceForm.monto} onChange={(e) => setAvanceForm((f) => ({ ...f, monto: e.target.value }))} className={inputCls} />
                              </Field>
                              <div className="md:col-span-2">
                                <Field label="Descripción">
                                  <input value={avanceForm.descripcion} onChange={(e) => setAvanceForm((f) => ({ ...f, descripcion: e.target.value }))} placeholder="Ej: 3er avance" className={inputCls} />
                                </Field>
                              </div>
                              <div className="md:col-span-4">
                                <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar avance</button>
                              </div>
                            </form>
                          )}

                          {avancesGrupo.length > 0 && (
                            <div className="mt-4 space-y-1 border-t border-stone-100 pt-3">
                              {avancesGrupo.map((a) => (
                                <div key={a.id} className="flex items-center justify-between text-xs text-slate-500">
                                  <span>{fmtFecha(a.fecha)} — {a.descripcion || "Avance"}</span>
                                  <span className="font-mono text-slate-700">{fmtARS(a.monto)}</span>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </>
            )}

            {vistaLiquidacion === "historial" && (
              <>
                <select className={inputCls} value={obraHistorialId} onChange={(e) => setObraHistorialId(e.target.value)}>
                  {obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                </select>

                <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Total pagado en esta obra</div>
                  <div className="mt-1 font-mono text-xl font-bold text-slate-900">{fmtARS(totalHistorico)}</div>
                </div>

                {historialPagos.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    Todavía no se registraron pagos en esta obra.
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                        <tr>
                          <th className="px-3 py-3">Fecha trabajada</th>
                          <th className="px-3 py-3">Nombre</th>
                          <th className="px-3 py-3">Rubro</th>
                          <th className="px-3 py-3">Hs</th>
                          <th className="px-3 py-3 text-right">Monto pagado</th>
                          <th className="px-3 py-3">Fecha de pago</th>
                        </tr>
                      </thead>
                      <tbody>
                        {historialPagos.map((a) => (
                          <tr key={a.id} className="border-t border-stone-100">
                            <td className="px-3 py-2 text-slate-600">{fmtFecha(a.fecha)}</td>
                            <td className="px-3 py-2 font-medium text-slate-900">{a.nombre}</td>
                            <td className="px-3 py-2 text-slate-600">{categoriaDe(a.nombre) || "—"}</td>
                            <td className="px-3 py-2 font-mono text-slate-700">{a.horas}</td>
                            <td className="px-3 py-2 text-right font-mono text-slate-800">{fmtARS(a.montoAbonado)}</td>
                            <td className="px-3 py-2 text-slate-500"><Badge estado="Pagado" />{" "}{fmtFecha(a.fechaPago)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </>
            )}
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

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
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
                      <div className="mt-1 text-xs text-slate-400">{obra?.nombre} · {fmtFecha(oc.fecha)}</div>
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
                      formalidad: f.get("formalidad"),
                      cuenta: f.get("cuenta"),
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
                  <Field label="Formalidad">
                    <select name="formalidad" className={inputCls}>{FORMALIDADES.map((f) => <option key={f}>{f}</option>)}</select>
                  </Field>
                  <Field label="Cuenta de pago">
                    <select name="cuenta" className={inputCls}>{CUENTAS.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-4 py-3">Fecha</th><th className="px-4 py-3">Obra</th><th className="px-4 py-3">Proveedor</th><th className="px-4 py-3">Categoría</th><th className="px-4 py-3">Formalidad</th><th className="px-4 py-3">Cuenta</th><th className="px-4 py-3">Monto</th><th className="px-4 py-3">Estado</th></tr>
                </thead>
                <tbody>
                  {[...comprasFacturas].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((c) => {
                    const obra = obras.find((o) => o.id === c.obraId);
                    return (
                      <tr key={c.id} className="border-t border-stone-100">
                        <td className="px-4 py-3 text-slate-600">{fmtFecha(c.fecha)}</td>
                        <td className="px-4 py-3 text-slate-600">{obra?.nombre}</td>
                        <td className="px-4 py-3 font-medium text-slate-900">{c.proveedor}</td>
                        <td className="px-4 py-3 text-slate-600">{c.categoria}</td>
                        <td className="px-4 py-3"><Badge estado={c.formalidad || "Blanco"} /></td>
                        <td className="px-4 py-3 text-slate-600"><span className="flex items-center gap-1"><CuentaIcon cuenta={c.cuenta} />{c.cuenta || "—"}</span></td>
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

        {tab === "ingresos" && !canVerFinanzas && (
          <div className="flex flex-col items-center justify-center gap-3 rounded-lg border-2 border-dashed border-stone-300 bg-white p-10 text-center">
            <TrendingUp size={28} className="text-slate-400" />
            <div className="text-sm font-semibold text-slate-700">Sección restringida</div>
            <div className="max-w-sm text-xs text-slate-500">Solo Gerente y Contador pueden ver y cargar ingresos.</div>
          </div>
        )}

        {tab === "ingresos" && canVerFinanzas && (
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Ingresos</h2>
              <button onClick={() => setShowIngresoForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Cargar ingreso
              </button>
            </div>

            <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
              Lo que cobrás de cada obra — con esto más los gastos de "Compras y Facturas" se arma el resumen de "Cuentas".
            </div>

            {showIngresoForm && (
              <Panel title="Cargar ingreso" action={<button onClick={() => setShowIngresoForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    addRecord("ingresos", {
                      fecha: f.get("fecha"),
                      obraId: Number(f.get("obraId")),
                      concepto: f.get("concepto"),
                      monto: Number(f.get("monto")) || 0,
                      formalidad: f.get("formalidad"),
                      cuenta: f.get("cuenta"),
                    }, setIngresos);
                    e.target.reset();
                    setShowIngresoForm(false);
                  }}
                >
                  <Field label="Fecha"><input name="fecha" type="date" required className={inputCls} /></Field>
                  <Field label="Obra">
                    <select name="obraId" className={inputCls}>{obras.map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}</select>
                  </Field>
                  <Field label="Concepto"><input name="concepto" required placeholder="Ej: certificado de avance 3" className={inputCls} /></Field>
                  <Field label="Monto (ARS)"><input name="monto" type="number" required className={inputCls} /></Field>
                  <Field label="Formalidad">
                    <select name="formalidad" className={inputCls}>{FORMALIDADES.map((f) => <option key={f}>{f}</option>)}</select>
                  </Field>
                  <Field label="Cuenta">
                    <select name="cuenta" className={inputCls}>{CUENTAS.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-sm">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-4 py-3">Fecha</th><th className="px-4 py-3">Obra</th><th className="px-4 py-3">Concepto</th><th className="px-4 py-3">Formalidad</th><th className="px-4 py-3">Cuenta</th><th className="px-4 py-3">Monto</th></tr>
                </thead>
                <tbody>
                  {[...ingresos].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((i) => {
                    const obra = obras.find((o) => o.id === i.obraId);
                    return (
                      <tr key={i.id} className="border-t border-stone-100">
                        <td className="px-4 py-3 text-slate-600">{fmtFecha(i.fecha)}</td>
                        <td className="px-4 py-3 text-slate-600">{obra?.nombre}</td>
                        <td className="px-4 py-3 font-medium text-slate-900">{i.concepto}</td>
                        <td className="px-4 py-3"><Badge estado={i.formalidad || "Blanco"} /></td>
                        <td className="px-4 py-3 text-slate-600"><span className="flex items-center gap-1"><CuentaIcon cuenta={i.cuenta} />{i.cuenta || "—"}</span></td>
                        <td className="px-4 py-3 text-right font-mono font-semibold text-emerald-700">{fmtARS(i.monto)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab === "cuentas" && !canVerFinanzas && (
          <div className="flex flex-col items-center justify-center gap-3 rounded-lg border-2 border-dashed border-stone-300 bg-white p-10 text-center">
            <Landmark size={28} className="text-slate-400" />
            <div className="text-sm font-semibold text-slate-700">Sección restringida</div>
            <div className="max-w-sm text-xs text-slate-500">Solo Gerente y Contador pueden ver el resumen de cuentas.</div>
          </div>
        )}

        {tab === "cuentas" && canVerFinanzas && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Cuentas</h2>

            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
              <div className="rounded-lg border border-sky-200 bg-sky-50 p-4">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-sky-700">Total en blanco</div>
                <div className="mt-1 font-mono text-2xl font-bold text-sky-900">{fmtARS(totalBlanco)}</div>
              </div>
              <div className="rounded-lg border border-slate-300 bg-slate-50 p-4">
                <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-600">Total en negro</div>
                <div className="mt-1 font-mono text-2xl font-bold text-slate-800">{fmtARS(totalNegro)}</div>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
              {saldosCuentas.map(({ cuenta, formalidad, saldo }) => (
                <div key={`${cuenta}-${formalidad}`} className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                  <div className="flex items-center justify-between">
                    <span className="flex items-center gap-1.5 text-sm font-semibold text-slate-700"><CuentaIcon cuenta={cuenta} size={15} />{cuenta}</span>
                    <Badge estado={formalidad} />
                  </div>
                  <div className={`mt-2 font-mono text-lg font-bold ${saldo < 0 ? "text-rose-600" : "text-slate-900"}`}>{fmtARS(saldo)}</div>
                </div>
              ))}
            </div>

            <div className="text-[11px] text-slate-400">
              Saldo = Ingresos − Compras/Facturas de cada cuenta y formalidad. Un saldo negativo significa que se cargaron más gastos que ingresos en esa combinación.
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

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
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
