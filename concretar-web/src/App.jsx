import { useState, useEffect, useRef, Fragment } from "react";
import { createClient } from "@supabase/supabase-js";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Building2, Users, ClipboardCheck, Wrench,
  ShoppingCart, Receipt, Plus, MapPin, TrendingUp, TrendingDown, X, AlertTriangle, CheckCircle2,
  Database, Loader2, RefreshCw, DollarSign, Check, Menu, FileDown, ShieldCheck,
  Printer, HardHat, Zap, PaintRoller, Droplet, Hammer, Flame, Wallet,
  Landmark, Smartphone, Banknote, Briefcase, Info, Pencil, Truck, ArrowRightLeft, CalendarDays, Package, Upload, FileSpreadsheet, Trash2
} from "lucide-react";
import {
  ResponsiveContainer, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend
} from "recharts";

// Paleta oficial del Manual de Marca (Grupo Concretar S.A.S)
const BRAND = {
  navy900: "#021d34", // Tono principal
  navy700: "#153f59",
  navy400: "#3a5c66",
  font: "'Poppins', ui-sans-serif, system-ui, sans-serif",
};

const fmtARS = (n) =>
  new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n || 0);

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

const ESTADOS_HERRAMIENTA = ["Disponible", "En Obra", "En Reparación", "Mal Estado", "Rota"];
const ESTADOS_OBRA = ["En curso", "Pausada", "Finalizada"];
const ESTADOS_ITEM_COMBO = ["Entregado", "Roto", "Perdido", "Devuelto"];
const TIPOS_CAJA = ["Electricista", "Civil", "Pintor", "Metalúrgico"];
const DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const DIAS_SEMANA_JS = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
const ESTADOS_OC = ["Pendiente", "Requiere aprobación", "Aprobada", "Recibida"];
const ESTADOS_FACTURA = ["Pendiente", "Pagada"];
const ESTADOS_PEDIDO_MATERIAL = ["Solicitado", "Aprobado", "Rechazado", "Facturado", "Recibido"];
const CATEGORIAS_GASTO = ["Materiales", "Mano de obra", "Equipos", "Otros"];
const CATEGORIAS_PEDIDO = ["Materiales", "Herramientas", "Equipos", "Epps", "Consumibles", "Otros"];
// La pestaña "Pedidos de Obra" arma el pedido en pasos (rubros). Cada uno filtra
// las líneas del presupuesto importado (o el catálogo propio, en el caso de Epps
// y Consumibles) por estas categorías. Epps y Consumibles no se ordenan por rubro
// de obra porque un mismo ítem suele servir para varios rubros a la vez.
const CATEGORIAS_POR_VISTA = {
  materiales: ["Materiales"],
  equipos: ["Equipos", "Herramientas"],
  epps: ["Epps"],
  consumibles: ["Consumibles"],
};
// Cuando el capataz agrega un ítem a mano no elige categoría — se infiere de la pestaña
// en la que está parado (Materiales / Equipos y Herramientas / Epps / Consumibles).
const CATEGORIA_DE_VISTA = { materiales: "Materiales", equipos: "Equipos", epps: "Epps", consumibles: "Consumibles" };
const CATEGORIAS_HERRAMIENTA = ["Herramienta Eléctrica", "Herramienta Manual", "Equipo Eléctrico", "Equipo a Combustión"];
const SI_NO = ["No", "Sí"];
// Letra usada en el N° de serie automático (Tipo-Marca+N°). Cambiá acá si preferís otras letras.
const LETRA_TIPO_HERRAMIENTA = {
  "Herramienta Eléctrica": "E",
  "Herramienta Manual": "M",
  "Equipo Eléctrico": "Q",
  "Equipo a Combustión": "C",
};
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
const ICONO_CATEGORIA_HERR = {
  "Herramienta Eléctrica": Zap,
  "Herramienta Manual": Hammer,
  "Equipo Eléctrico": Wrench,
  "Equipo a Combustión": Flame,
};
const ESTADOS_ASISTENCIA = ["Presente", "Ausente", "Tardanza"];
const UMBRAL_APROBACION_OC = 3000000;
const DESVIO_ALERTA_PCT = 10;
const DESVIO_DANGER_PCT = 20;

// Roles que pueden "iniciar sesión" (simulado hasta que armemos el login real)
const ROLES = ["Gerente", "Recursos Humanos", "HyS", "Capataz", "Logística", "Contador", "Otro (sin acceso)"];
const ROLES_ALTA_PERSONAL = ["Gerente", "Recursos Humanos", "HyS", "Capataz"];
const ROLES_EDITAR_PERSONAL = ["Gerente", "Recursos Humanos"];
const ROLES_EDITAR_COSTOS = ["Gerente", "Recursos Humanos"];
const ROLES_LIQUIDACION = ["Gerente", "Contador"];
const ROLES_FINANZAS = ["Gerente", "Contador"];
// Precios de pedidos de obra: el capataz arma el pedido a ciegas (sin precios ni proveedor),
// eso lo ve y lo carga Logística cuando el pedido llega aprobado.
const ROLES_VEN_PRECIOS_PEDIDO = ["Gerente", "Contador", "Logística"];
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
  Disponible: "border-green-600 text-green-700",
  "En Obra": "border-amber-800 text-amber-900",
  "En Reparación": "border-cyan-500 text-cyan-600",
  "Mal Estado": "border-yellow-500 text-yellow-700",
  Rota: "border-red-600 text-red-700",
  Baja: "border-red-600 text-red-700",
  Entregado: "border-emerald-600 text-emerald-700",
  Roto: "border-rose-600 text-rose-700",
  Perdido: "border-rose-600 text-rose-700",
  Devuelto: "border-slate-400 text-slate-500",
  Pendiente: "border-slate-400 text-slate-500",
  Solicitado: "border-sky-600 text-sky-700",
  Aprobado: "border-emerald-600 text-emerald-700",
  Rechazado: "border-rose-600 text-rose-700",
  Facturado: "border-indigo-600 text-indigo-700",
  "Requiere aprobación": "border-rose-600 text-rose-700",
  Aprobada: "border-amber-600 text-amber-700",
  Recibida: "border-emerald-600 text-emerald-700",
  Pagada: "border-emerald-600 text-emerald-700",
  Pagado: "border-emerald-600 text-emerald-700",
  Blanco: "border-sky-600 text-sky-700",
  Negro: "border-slate-600 text-slate-700",
  "En curso": "border-amber-600 text-amber-700",
  Pausada: "border-rose-600 text-rose-700",
  Finalizada: "border-emerald-600 text-emerald-700",
  Papelera: "border-slate-400 text-slate-500",
  Activo: "border-emerald-600 text-emerald-700",
  Licencia: "border-amber-600 text-amber-700",
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

function CategoriaHerrIcon({ categoria, size = 13 }) {
  const IconComp = ICONO_CATEGORIA_HERR[categoria];
  if (!IconComp) return null;
  return (
    <span title={categoria} className="inline-flex items-center text-slate-500">
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

// Campo de plata: mientras escribís va agregando puntos de miles y ",00" de decimales
// en vivo (1 -> 1,00 -> 10 -> 10,00 -> 1000 -> 1.000,00), como en una caja registradora.
// Funciona tanto con formularios controlados (value+onChange) como con FormData (name).
function MoneyInput({ name, value, onChange, onBlur, className, placeholder, required, disabled }) {
  const [raw, setRaw] = useState(value !== undefined && value !== null && value !== "" ? String(Math.round(Number(value))) : "");
  function handleChange(e) {
    const digitos = e.target.value.replace(/\D/g, "").replace(/^0+(?=\d)/, "");
    setRaw(digitos);
    if (onChange) onChange(digitos === "" ? 0 : parseInt(digitos, 10));
  }
  const num = raw === "" ? 0 : parseInt(raw, 10);
  const display = raw === "" ? "" : new Intl.NumberFormat("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(num);
  return (
    <>
      <input
        type="text"
        inputMode="numeric"
        value={display}
        onChange={handleChange}
        onBlur={() => onBlur && onBlur(num)}
        placeholder={placeholder || "0,00"}
        disabled={disabled}
        className={`${className || inputCls} ${disabled ? "cursor-not-allowed bg-stone-100 text-slate-400" : ""}`}
      />
      {name && <input type="hidden" name={name} value={num} required={required} />}
    </>
  );
}
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

const ALERT_TONE_CARD = {
  rose: "border-rose-200 bg-rose-50 text-rose-900",
  amber: "border-amber-200 bg-amber-50 text-amber-900",
  sky: "border-sky-200 bg-sky-50 text-sky-900",
  emerald: "border-emerald-200 bg-emerald-50 text-emerald-900",
};
const ALERT_TONE_ICON = {
  rose: "bg-rose-500",
  amber: "bg-amber-500",
  sky: "bg-sky-500",
  emerald: "bg-emerald-500",
};

// Tarjeta visual para el panel de Alertas: ícono en círculo + título + detalle opcional.
// Reemplaza las franjas de texto planas por bloques que se agrupan en grilla en PC.
function AlertCard({ tone = "amber", icon: Icon = AlertTriangle, title, children }) {
  return (
    <div className={`rounded-xl border p-3.5 shadow-sm ${ALERT_TONE_CARD[tone]}`}>
      <div className="flex items-start gap-3">
        <span className={`mt-0.5 flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-white ${ALERT_TONE_ICON[tone]}`}>
          <Icon size={16} />
        </span>
        <div className="min-w-0 flex-1 space-y-1.5">
          <div className="text-sm font-semibold leading-snug">{title}</div>
          {children}
        </div>
      </div>
    </div>
  );
}

// Agrupa pedidos de materiales por obra para mostrar "N pedido(s) para tal obra" en vez
// de listar materiales sueltos. Ordena por la fecha necesaria más próxima primero.
function agruparPedidosPorObra(pedidos, obras) {
  const grupos = new Map();
  for (const p of pedidos) {
    if (!grupos.has(p.obraId)) grupos.set(p.obraId, []);
    grupos.get(p.obraId).push(p);
  }
  return [...grupos.entries()]
    .map(([obraId, lista]) => ({
      obraId,
      nombreObra: obras.find((o) => o.id === obraId)?.nombre || "Obra sin nombre",
      pedidos: [...lista].sort((a, b) => fechaLocal(a.fechaNecesaria) - fechaLocal(b.fechaNecesaria)),
    }))
    .sort((a, b) => fechaLocal(a.pedidos[0]?.fechaNecesaria) - fechaLocal(b.pedidos[0]?.fechaNecesaria));
}

function monthsBetween(d1, d2) {
  return (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
}

export default function ConcretarApp() {
  const [tab, setTab] = useState("general");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  const DEMO_OBRAS = [
    { id: 1, nombre: "Edificio Belgrano 450", cliente: "Consorcio Belgrano SA", clienteId: 1, presupuesto: 85000000, meses: 10, inicio: "2026-02-01", estado: "En curso", encargadoId: 4, diasLaborables: ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"], horaApertura: "08:00", diaCierre: "Viernes", horaCierre: "18:00" },
    { id: 2, nombre: "Casa Quinta Yerba Buena", cliente: "Fam. Ledesma", clienteId: 2, presupuesto: 32000000, meses: 6, inicio: "2026-05-01", estado: "En curso", encargadoId: null, diasLaborables: ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"], horaApertura: "07:30", diaCierre: "Viernes", horaCierre: "17:30" },
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
  const DEMO_COSTOS = [
    { id: 1, categoria: "Oficial Especializado", mes: "2026-07", costoHora: 7000 },
    { id: 2, categoria: "Oficial", mes: "2026-07", costoHora: 6100 },
    { id: 3, categoria: "Medio Oficial", mes: "2026-07", costoHora: 5200 },
    { id: 4, categoria: "Ayudante", mes: "2026-07", costoHora: 4200 },
    { id: 5, categoria: "Oficial Especializado", mes: "2026-08", costoHora: 7500 },
    { id: 6, categoria: "Oficial", mes: "2026-08", costoHora: 6500 },
    { id: 7, categoria: "Medio Oficial", mes: "2026-08", costoHora: 5500 },
    { id: 8, categoria: "Ayudante", mes: "2026-08", costoHora: 4500 },
  ];

  const DEMO_ASISTENCIA = [
    { id: 1, fecha: "2026-08-03", obraId: 1, nombre: "Pablo Robles", horas: 8, cargadoPor: "Gerente", estado: "Presente" },
    { id: 2, fecha: "2026-08-03", obraId: 1, nombre: "Eduardo Sr", horas: 9, cargadoPor: "Gerente", estado: "Presente" },
    { id: 3, fecha: "2026-08-04", obraId: 1, nombre: "Daniel Tello", horas: 8, cargadoPor: "Capataz", estado: "Tardanza" },
    { id: 4, fecha: "2026-08-05", obraId: 2, nombre: "Facundo C", horas: 6, cargadoPor: "Capataz", estado: "Presente" },
    { id: 5, fecha: "2026-08-05", obraId: 2, nombre: "Pepito Chespirito", horas: 0, cargadoPor: "Capataz", estado: "Ausente" },
  ];
  const DEMO_HERRAMIENTAS = [
    { id: 1, nombre: "Amoladora angular", numeroSerie: "E-MAK01", marca: "Makita", categoria: "Herramienta Eléctrica", ubicacion: "Oficina", estado: "Disponible", maletin: "Sí", accesorios: "Sí", detalleAccesorios: "2 discos de corte de repuesto", observaciones: "", fechaUltimoCambioEstado: null },
    { id: 2, nombre: "Andamio tubular (juego x6)", numeroSerie: "M-GEN01", marca: "", categoria: "Herramienta Manual", ubicacion: "Edificio Belgrano 450", estado: "En Obra", maletin: "No", accesorios: "No", detalleAccesorios: "", observaciones: "Faltan 2 crucetas, revisar al devolver", fechaUltimoCambioEstado: null },
    { id: 3, nombre: "Rotomartillo SDS", numeroSerie: "E-BOS01", marca: "Bosch", categoria: "Herramienta Eléctrica", ubicacion: "Casa Quinta Yerba Buena", estado: "Mal Estado", maletin: "Sí", accesorios: "Sí", detalleAccesorios: "3 puntas SDS, 1 cincel", observaciones: "Pierde potencia, revisar antes de que se rompa del todo", fechaUltimoCambioEstado: null },
    { id: 4, nombre: "Nivel láser", numeroSerie: "Q-STA01", marca: "Stanley", categoria: "Equipo Eléctrico", ubicacion: "Electromecánica Ríos", estado: "En Reparación", maletin: "Sí", accesorios: "No", detalleAccesorios: "", observaciones: "No enciende, revisar batería", fechaUltimoCambioEstado: "2026-08-18T15:00:00.000Z" },
  ];
  const DEMO_COMBOS = [
    {
      id: 1,
      tipo: "Electricista",
      numero: 1,
      obraId: 1,
      personaId: 2,
      fecha: "2026-06-01",
      items: [
        { nombre: "Martillo", cantidad: 1, estado: "Entregado" },
        { nombre: "Plomada", cantidad: 1, estado: "Entregado" },
        { nombre: "Tenaza", cantidad: 1, estado: "Roto" },
      ],
    },
    {
      id: 2,
      tipo: "Civil",
      numero: 1,
      obraId: null,
      personaId: null,
      fecha: "2026-07-10",
      items: [
        { nombre: "Cinta métrica", cantidad: 1, estado: "Entregado" },
        { nombre: "Nivel de mano", cantidad: 1, estado: "Entregado" },
      ],
    },
  ];
  const DEMO_CATALOGO_NOMBRES = [
    { id: 1, categoria: "Herramienta Eléctrica", nombre: "Amoladora angular" },
    { id: 2, categoria: "Herramienta Eléctrica", nombre: "Rotomartillo SDS" },
    { id: 3, categoria: "Herramienta Eléctrica", nombre: "Taladro percutor" },
    { id: 4, categoria: "Herramienta Manual", nombre: "Andamio tubular (juego x6)" },
    { id: 5, categoria: "Herramienta Manual", nombre: "Carretilla" },
    { id: 6, categoria: "Equipo Eléctrico", nombre: "Nivel láser" },
    { id: 7, categoria: "Equipo Eléctrico", nombre: "Soldadora" },
    { id: 8, categoria: "Equipo a Combustión", nombre: "Generador" },
    { id: 9, categoria: "Equipo a Combustión", nombre: "Compactadora" },
  ];
  const DEMO_CATALOGO_MARCAS = [
    { id: 1, nombre: "Makita" },
    { id: 2, nombre: "Bosch" },
    { id: 3, nombre: "Stanley" },
    { id: 4, nombre: "DeWalt" },
  ];
  const DEMO_CATALOGO_CHICAS = [
    { id: 1, nombre: "Martillo" },
    { id: 2, nombre: "Plomada" },
    { id: 3, nombre: "Tenaza" },
    { id: 4, nombre: "Destornillador Phillips" },
    { id: 5, nombre: "Destornillador Plano" },
    { id: 6, nombre: "Cinta métrica" },
    { id: 7, nombre: "Nivel de mano" },
    { id: 8, nombre: "Serrucho" },
  ];
  const DEMO_PROVEEDORES = [
    { id: 1, razonSocial: "Corralón San Martín", cuit: "30-12345678-9", domicilio: "Ruta 40 km 12, San Juan", contacto: "Marcos Díaz", telefono: "264-4000001", esTaller: "No" },
    { id: 2, razonSocial: "Electromecánica Ríos", cuit: "30-98765432-1", domicilio: "Av. Libertador 850, San Juan", contacto: "Ríos Hnos.", telefono: "264-4000002", esTaller: "Sí" },
  ];
  const DEMO_CLIENTES = [
    { id: 1, razonSocial: "Consorcio Belgrano SA", cuit: "30-55566677-8", domicilio: "Belgrano 450, San Juan", contacto: "Adm. Belgrano", telefono: "264-4100001" },
    { id: 2, razonSocial: "Fam. Ledesma", cuit: "", domicilio: "Yerba Buena, San Juan", contacto: "Roberto Ledesma", telefono: "264-4100002" },
  ];
  const DEMO_REMITOS = [
    {
      id: 1,
      fecha: "2026-08-10",
      origen: "Edificio Belgrano 450",
      destino: "Oficina",
      destinoEsTaller: false,
      destinoProveedorId: null,
      herramientaIds: [4],
      estado: "Recibido",
      creadoPor: "Capataz",
      fechaRecepcion: "2026-08-11",
      recibidoPor: "Gerente",
    },
  ];
  const DEMO_AUDITORIAS = [
    {
      id: 1,
      obraId: 1,
      fecha: "2026-08-14",
      tipo: "Cierre",
      realizadoPor: "Capataz",
      herramientasPresentes: [2],
      herramientasFaltantes: [],
      observaciones: "Todo en orden, semana cerrada sin novedades.",
    },
  ];
  const DEMO_FERIADOS = [
    { id: 1, fecha: "2026-12-25", descripcion: "Navidad" },
    { id: 2, fecha: "2026-01-01", descripcion: "Año Nuevo" },
    { id: 3, fecha: "2026-08-17", descripcion: "Paso a la Inmortalidad del Gral. San Martín" },
  ];
  const DEMO_SUBCATEGORIAS_MAT = [
    { id: 1, categoria: "Materiales", nombre: "Civil" },
    { id: 2, categoria: "Materiales", nombre: "Electricidad" },
    { id: 3, categoria: "Materiales", nombre: "Plomería" },
  ];
  const DEMO_TIPOS_MATERIAL = [
    { id: 1, categoria: "Materiales", subcategoria: "Civil", nombre: "Hierros Nervurados" },
    { id: 2, categoria: "Materiales", subcategoria: "Electricidad", nombre: "Daisa" },
    { id: 3, categoria: "Materiales", subcategoria: "Plomería", nombre: "Grifería" },
  ];
  const DEMO_CATALOGO_MATERIALES = [
    { id: 1, categoria: "Materiales", subcategoria: "Civil", tipo: "Hierros Nervurados", nombre: "Hierro fi 10mm", unidad: "Barra", ultimoPrecio: 12000, ultimoProveedor: "Corralón San Martín" },
    { id: 2, categoria: "Materiales", subcategoria: "Civil", tipo: "", nombre: "Cemento", unidad: "Bolsa", ultimoPrecio: 8500, ultimoProveedor: "Corralón San Martín" },
    { id: 3, categoria: "Materiales", subcategoria: "Electricidad", tipo: "Daisa", nombre: "Caño Daisa 3/4", unidad: "Metro", ultimoPrecio: 450, ultimoProveedor: "Electromecánica Ríos" },
    { id: 4, categoria: "Materiales", subcategoria: "Plomería", tipo: "Grifería", nombre: "Grifería FV monocomando", unidad: "Unidad", ultimoPrecio: 35000, ultimoProveedor: "Corralón San Martín" },
  ];
  const DEMO_PRESUPUESTO_GENERAL = [
    { id: 1, obraId: 1, totalManoObra: 16800869, totalEquipos: 873612.6, totalMateriales: 12187643.87, precioTotalSinIva: 43793871.95, precioTotalConIva: 52990585.06, fechaImportacion: "2026-08-01" },
  ];
  const DEMO_PRESUPUESTO_MATERIALES = [
    { id: 1, obraId: 1, categoria: "Materiales", subcategoria: "Materiales Civiles", tipo: "", material: "Cemento x 25kg", unidad: "und.", cantidad: 30, precioUnitario: 5775, total: 173250, observaciones: "", origen: "Excel", pedidoId: null },
    { id: 2, obraId: 1, categoria: "Materiales", subcategoria: "Materiales Civiles", tipo: "", material: "Barra hierro 8mm", unidad: "und.", cantidad: 5, precioUnitario: 8085, total: 40425, observaciones: "", origen: "Excel", pedidoId: null },
    { id: 3, obraId: 1, categoria: "Equipos", subcategoria: "", tipo: "", material: "Contenedor", unidad: "und.", cantidad: 2, precioUnitario: 150000, total: 300000, observaciones: "", origen: "Excel", pedidoId: null },
  ];
  const DEMO_PEDIDOS_MATERIALES = [];
  const DEMO_STOCK_MATERIALES = [];
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
  const [combosHerramientas, setCombosHerramientas] = useState(isSupabaseConfigured ? [] : DEMO_COMBOS);
  const [catalogoNombresHerr, setCatalogoNombresHerr] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_NOMBRES);
  const [catalogoMarcas, setCatalogoMarcas] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_MARCAS);
  const [catalogoChicas, setCatalogoChicas] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_CHICAS);
  const [proveedores, setProveedores] = useState(isSupabaseConfigured ? [] : DEMO_PROVEEDORES);
  const [clientes, setClientes] = useState(isSupabaseConfigured ? [] : DEMO_CLIENTES);
  const [remitos, setRemitos] = useState(isSupabaseConfigured ? [] : DEMO_REMITOS);
  const [auditorias, setAuditorias] = useState(isSupabaseConfigured ? [] : DEMO_AUDITORIAS);
  const [feriados, setFeriados] = useState(isSupabaseConfigured ? [] : DEMO_FERIADOS);
  const [subcategoriasMat, setSubcategoriasMat] = useState(isSupabaseConfigured ? [] : DEMO_SUBCATEGORIAS_MAT);
  const [tiposMaterial, setTiposMaterial] = useState(isSupabaseConfigured ? [] : DEMO_TIPOS_MATERIAL);
  const [catalogoMateriales, setCatalogoMateriales] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_MATERIALES);
  const [presupuestoGeneral, setPresupuestoGeneral] = useState(isSupabaseConfigured ? [] : DEMO_PRESUPUESTO_GENERAL);
  const [presupuestoMateriales, setPresupuestoMateriales] = useState(isSupabaseConfigured ? [] : DEMO_PRESUPUESTO_MATERIALES);
  const [pedidosMateriales, setPedidosMateriales] = useState(isSupabaseConfigured ? [] : DEMO_PEDIDOS_MATERIALES);
  const [stockMateriales, setStockMateriales] = useState(isSupabaseConfigured ? [] : DEMO_STOCK_MATERIALES);
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
        // Además del cron horario en Supabase, disparamos la purga acá para que
        // una obra vencida en Papelera desaparezca apenas alguien abre la app.
        try { await supabase.rpc("purgar_obras_papelera_vencidas"); } catch { /* el cron del servidor la va a agarrar igual */ }
        const [o, p, cc, a, h, oc, cf, ing, tt, av, ch, cn, cm, cch, pv, rm, au, fer, cli, sm, tm, cma, pma, ped, pg, stk] = await Promise.all([
          sbSelect("obras"), sbSelect("personal"), sbSelect("costos_categoria"), sbSelect("asistencia"),
          sbSelect("herramientas"), sbSelect("ordenes_compra"), sbSelect("compras_facturas"), sbSelect("ingresos"),
          sbSelect("tanteros"), sbSelect("avances_tanteros"), sbSelect("combos_herramientas"),
          sbSelect("catalogo_nombres_herramienta"), sbSelect("catalogo_marcas"), sbSelect("catalogo_herramientas_chicas"),
          sbSelect("proveedores"), sbSelect("remitos"), sbSelect("auditorias_herramientas"), sbSelect("feriados"), sbSelect("clientes"),
          sbSelect("subcategorias_material"), sbSelect("tipos_material"), sbSelect("catalogo_materiales"), sbSelect("presupuesto_materiales"),
          sbSelect("pedidos_materiales"), sbSelect("presupuesto_general"), sbSelect("stock_materiales"),
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
        setCombosHerramientas(ch);
        setCatalogoNombresHerr(cn);
        setCatalogoMarcas(cm);
        setCatalogoChicas(cch);
        setProveedores(pv);
        setClientes(cli);
        setRemitos(rm);
        setAuditorias(au);
        setFeriados(fer);
        setSubcategoriasMat(sm);
        setTiposMaterial(tm);
        setCatalogoMateriales(cma);
        setPresupuestoMateriales(pma);
        setPedidosMateriales(ped);
        setPresupuestoGeneral(pg);
        setStockMateriales(stk);
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
        return row;
      } catch (err) {
        alert("No se pudo guardar: " + err.message);
        return null;
      }
    } else {
      const row = { ...obj, id: genId() };
      setter((prev) => [...prev, row]);
      return row;
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
  const anioActual = Number(hoyISO().slice(0, 4));
  const mesActualNum = Number(hoyISO().slice(5, 7));
  const [anioCostos, setAnioCostos] = useState(anioActual);
  const aniosCostosDisponibles = [anioActual - 1, anioActual, anioActual + 1, anioActual + 2];

  // Año actual: de este mes a diciembre. Otro año (ej. 2027): el año completo, para planificar con anticipación.
  const mesesCostos = Array.from(
    { length: anioCostos === anioActual ? 13 - mesActualNum : 12 },
    (_, i) => {
      const mesNum = (anioCostos === anioActual ? mesActualNum + i : i + 1);
      return `${anioCostos}-${String(mesNum).padStart(2, "0")}`;
    }
  );

  function nombreMes(mesStr) {
    return fechaLocal(`${mesStr}-01`).toLocaleDateString("es-AR", { month: "long" });
  }
  function guardarCostoCelda(categoria, mes, valor) {
    const existente = costosCategoria.find((c) => c.categoria === categoria && c.mes === mes);
    if (existente) {
      if (existente.costoHora === valor) return;
      updateRecord("costos_categoria", existente.id, { costoHora: valor }, setCostosCategoria);
    } else {
      if (!valor) return;
      addRecord("costos_categoria", { categoria, mes, costoHora: valor }, setCostosCategoria);
    }
  }

  const [viewingPersonId, setViewingPersonId] = useState(null);
  const [modoSeleccionPdf, setModoSeleccionPdf] = useState(false);
  const [showCostosPanel, setShowCostosPanel] = useState(false);
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

  // Gerencia, RRHH, Logística, HyS trabajan "de todos lados" y por defecto
  // se imputan al Centro de Costos General, no a una obra puntual.
  const CATEGORIAS_CENTRO_GENERAL = ["Gerente", "Recursos Humanos", "Logística", "HyS"];

  // En qué obra está trabajando esta persona ahora mismo:
  // - Tantero: la obra del grupo al que pertenece.
  // - Empresa: la obra de su registro de asistencia más reciente.
  function obraActualDe(p) {
    if (p.tipoTrabajador === "Tantero") {
      const grupo = tanteros.find((t) => (t.integrantes || []).includes(p.id));
      if (!grupo) return null;
      const obraGrupo = obras.find((o) => o.id === grupo.obraId);
      return obraGrupo?.estado === "En curso" ? obraGrupo : null;
    }
    const registros = asistencia
      .filter((a) => a.nombre === nombreCompletoDe(p))
      .sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));
    if (registros.length === 0) return null;
    const obraReciente = obras.find((o) => o.id === registros[0].obraId);
    return obraReciente?.estado === "En curso" ? obraReciente : null;
  }
  function ultimaFechaActividad(p) {
    const registros = asistencia.filter((a) => a.nombre === nombreCompletoDe(p)).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));
    return registros[0]?.fecha || null;
  }
  // Personal activo, de obra (no Gerencia/RRHH/Logística que trabaja "de todos lados"),
  // sin ninguna asistencia registrada en los últimos 5 días — candidato a dar de baja.
  const personalSinObra5Dias = personal.filter((p) => {
    if (p.estado !== "Activo") return false;
    if (CATEGORIAS_CENTRO_GENERAL.includes(p.categoria)) return false;
    if (p.tipoTrabajador === "Tantero") return false; // se rastrean por grupo, no por asistencia diaria
    const ultima = ultimaFechaActividad(p);
    if (!ultima) return false;
    const dias = Math.round((fechaLocal(hoyISO()) - fechaLocal(ultima)) / 86400000);
    return dias >= 5;
  });
  function darDeBajaPersonal(p) {
    if (!window.confirm(`¿Dar de baja a ${nombreCompletoDe(p)}? Va a dejar de aparecer como personal activo.`)) return;
    updateRecord("personal", p.id, { estado: "Baja" }, setPersonal);
  }

  // ---------- Agrupación para "Personal/Cuadrillas" ----------
  const idsEnAlgunGrupoTantero = new Set(tanteros.flatMap((t) => t.integrantes || []));
  const personalCentroGeneral = personal.filter((p) => CATEGORIAS_CENTRO_GENERAL.includes(p.categoria));
  const cuadrillasPorObra = {}; // obraId -> { obra, empresa: [...], gruposTantero: [...] }
  const personalSinAsignar = [];

  personal
    .filter((p) => !CATEGORIAS_CENTRO_GENERAL.includes(p.categoria))
    .forEach((p) => {
      if (p.tipoTrabajador === "Tantero") {
        if (!idsEnAlgunGrupoTantero.has(p.id)) personalSinAsignar.push(p);
        return; // si está en un grupo, se muestra a través del grupo, no individualmente
      }
      const obraActual = obraActualDe(p);
      if (obraActual) {
        if (!cuadrillasPorObra[obraActual.id]) cuadrillasPorObra[obraActual.id] = { obra: obraActual, empresa: [], gruposTantero: [] };
        cuadrillasPorObra[obraActual.id].empresa.push(p);
      } else {
        personalSinAsignar.push(p);
      }
    });

  tanteros.forEach((t) => {
    if (!cuadrillasPorObra[t.obraId]) {
      const obra = obras.find((o) => o.id === t.obraId);
      if (obra) cuadrillasPorObra[t.obraId] = { obra, empresa: [], gruposTantero: [] };
    }
    if (cuadrillasPorObra[t.obraId]) cuadrillasPorObra[t.obraId].gruposTantero.push(t);
  });

  function renderPersonaRow(p) {
    const seleccionado = seleccionadosPdf.includes(p.id);
    return (
      <div
        key={p.id}
        onClick={() => modoSeleccionPdf && toggleSeleccionPdf(p.id)}
        className={`flex items-center justify-between gap-2 rounded-md px-2 py-1.5 text-sm ${modoSeleccionPdf ? "cursor-pointer" : ""} ${seleccionado ? "bg-amber-50" : "hover:bg-stone-50"}`}
      >
        <span className="flex items-center gap-2 min-w-0">
          {modoSeleccionPdf && <input type="checkbox" checked={seleccionado} onChange={() => toggleSeleccionPdf(p.id)} className="h-3.5 w-3.5 shrink-0" />}
          {p.fotoPersona ? (
            <img src={p.fotoPersona} alt={nombreCompletoDe(p)} className="h-6 w-6 shrink-0 rounded-full border border-stone-200 object-cover" />
          ) : (
            <div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-stone-100 text-[8px] font-semibold text-slate-400">{(p.nombre || "?").slice(0, 1)}</div>
          )}
          {modoSeleccionPdf ? (
            <span className="flex items-center gap-1 truncate font-medium text-slate-900">
              <EspecialidadIcon especialidad={p.especialidad} />
              {nombreCorto(p)}
              {p.observaciones && <span title={p.observaciones}><AlertTriangle size={11} className="text-amber-500 shrink-0" /></span>}
            </span>
          ) : (
            <button onClick={() => setViewingPersonId(p.id)} className="flex items-center gap-1 truncate font-medium text-slate-900 underline decoration-dotted hover:text-amber-600">
              <EspecialidadIcon especialidad={p.especialidad} />
              {nombreCorto(p)}
              {p.observaciones && <span title={p.observaciones}><AlertTriangle size={11} className="text-amber-500 shrink-0" /></span>}
            </button>
          )}
        </span>
        <span className="flex shrink-0 items-center gap-2 text-xs text-slate-500">
          <span>{p.categoria}</span>
          <Badge estado={p.estado} />
        </span>
      </div>
    );
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
    setViewingPersonId(null);
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
    { id: "general", label: "General", icon: LayoutDashboard },
    { id: "obras", label: "Obras", icon: Building2 },
    { id: "personal", label: "Personal/Cuadrillas", icon: Users },
    { id: "asistencia", label: "Asistencia", icon: ClipboardCheck },
    { id: "liquidacion", label: "Liquidación", icon: Wallet },
    { id: "herramientas", label: "Herramientas", icon: Wrench },
    { id: "materiales", label: "Pedidos de Obra", icon: Package },
    { id: "ordenes", label: "Órdenes de Compra", icon: ShoppingCart },
    { id: "ingresos", label: "Ingresos", icon: TrendingUp },
    { id: "facturas", label: "Compras y Facturas", icon: Receipt },
    { id: "cuentas", label: "Cuentas", icon: Landmark },
    { id: "proveedores", label: "Clientes/Proveedores", icon: Truck },
    { id: "calendario", label: "Calendario", icon: CalendarDays },
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
  const herramientasEnUso = obraSel ? herramientas.filter((h) => h.ubicacion === obraSel.nombre && h.estado === "En Obra").length : 0;

  // Obras en la Papelera: sus gastos/ingresos/pedidos dejan de contar en todos lados
  // hasta que se restauren o se borren del todo (24hs), como si nunca hubieran existido.
  const obraIdsPapelera = new Set(obras.filter((o) => o.estado === "Papelera").map((o) => o.id));

  // ---------- Alertas globales ----------
  const herramientasAtencion = herramientas.filter((h) => h.estado === "Mal Estado" || h.estado === "Rota");
  const herramientasReparadasRecientes = herramientas.filter((h) => {
    if (h.estado !== "Disponible" && h.estado !== "En Obra") return false;
    if (!h.fechaUltimoCambioEstado) return false;
    const horas = (Date.now() - new Date(h.fechaUltimoCambioEstado).getTime()) / 36e5;
    return horas >= 0 && horas < 48;
  });
  const ocPendientesAprobacion = ordenesCompra.filter((o) => o.estado === "Requiere aprobación" && !obraIdsPapelera.has(o.obraId));
  const hayDesvioAlerta = desvioPct > DESVIO_ALERTA_PCT;
  const asistenciasEditadas = asistencia.filter((a) => a.editado);
  function nombreDiaHoy() {
    return DIAS_SEMANA_JS[new Date().getDay()];
  }
  function esFeriadoHoy() {
    return feriados.some((f) => f.fecha === hoyISO());
  }
  function primerDiaLaborable(obra) {
    const dias = obra.diasLaborables || [];
    return DIAS_SEMANA.find((d) => dias.includes(d)) || null;
  }
  function dentroDeVentanaCierre(obra) {
    if (esFeriadoHoy()) return false;
    if (!obra.diaCierre || !obra.horaCierre) return false;
    if (nombreDiaHoy() !== obra.diaCierre) return false;
    const [h, m] = obra.horaCierre.split(":").map(Number);
    const ahora = new Date();
    const cierre = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate(), h, m || 0);
    const diffMin = (cierre - ahora) / 60000;
    return diffMin > 0 && diffMin <= 60;
  }
  function auditoriaHoy(obraId, tipo) {
    return auditorias.some((a) => a.obraId === obraId && a.tipo === tipo && a.fecha === hoyISO());
  }
  const obrasEnVentanaCierre = obras.filter((o) => dentroDeVentanaCierre(o) && !auditoriaHoy(o.id, "Cierre"));
  // "Apertura" dispara en el primer día laborable configurado de cada obra (normalmente lunes, pero se adapta si esa obra trabaja otros días), salvo feriado.
  const obrasSinAperturaLunes = esFeriadoHoy() ? [] : obras.filter((o) => nombreDiaHoy() === primerDiaLaborable(o) && !auditoriaHoy(o.id, "Apertura"));

  // Alarmas previas de materiales: según la "Fecha Necesaria" del presupuesto importado.
  function diasHasta(fechaStr) {
    return Math.round((fechaLocal(fechaStr) - fechaLocal(hoyISO())) / 86400000);
  }
  const pedidoEnCurso = (p) => p.estado === "Solicitado" || p.estado === "Aprobado";
  const materialesVencidos = pedidosMateriales.filter((p) => pedidoEnCurso(p) && p.fechaNecesaria && diasHasta(p.fechaNecesaria) < 0 && !obraIdsPapelera.has(p.obraId));
  const materialesProximos = pedidosMateriales.filter((p) => pedidoEnCurso(p) && p.fechaNecesaria && diasHasta(p.fechaNecesaria) >= 0 && diasHasta(p.fechaNecesaria) <= 2 && !obraIdsPapelera.has(p.obraId));
  const pedidosPorAprobar = pedidosMateriales.filter((p) => p.estado === "Solicitado" && !obraIdsPapelera.has(p.obraId));

  const totalAlertas =
    herramientasAtencion.length + herramientasReparadasRecientes.length + ocPendientesAprobacion.length +
    (hayDesvioAlerta ? 1 : 0) + asistenciasEditadas.length + obrasEnVentanaCierre.length + obrasSinAperturaLunes.length +
    materialesVencidos.length + materialesProximos.length + pedidosPorAprobar.length + personalSinObra5Dias.length;

  // ---------- Forms state ----------
  const [showObraForm, setShowObraForm] = useState(false);
  const [filtroObrasEstado, setFiltroObrasEstado] = useState("En curso");
  const [resumenObraImportado, setResumenObraImportado] = useState(null);
  const [itemsObraImportados, setItemsObraImportados] = useState([]);
  const [archivoObraNombre, setArchivoObraNombre] = useState("");
  const [creandoObra, setCreandoObra] = useState(false);
  const obraFileInputRef = useRef(null);

  function handleExcelUploadNuevaObra(e) {
    const file = e.target.files[0];
    if (!file) return;
    setArchivoObraNombre(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const { resumen, items } = parsePresupuestoGeneral(wb);
        if (items.length === 0) {
          alert("No se pudo leer ningún ítem de Equipos, Herramientas o Materiales. Revisá que sea la Planilla Interna de Costeo.");
        }
        setResumenObraImportado(resumen);
        setItemsObraImportados(items);
      } catch (err) {
        alert("No se pudo leer el archivo. Revisá que sea la Planilla Interna de Costeo en formato .xlsx.");
      }
    };
    reader.readAsArrayBuffer(file);
  }
  function quitarExcelNuevaObra() {
    setResumenObraImportado(null);
    setItemsObraImportados([]);
    setArchivoObraNombre("");
    if (obraFileInputRef.current) obraFileInputRef.current.value = "";
  }
  // Vuelca el Excel importado (Equipos/Herramientas/Materiales) a una obra puntual.
  // Se usa tanto al crear una obra nueva como al editar una que todavía no tenía presupuesto cargado.
  async function importarPresupuestoAObra(obraId) {
    if (!resumenObraImportado) return;
    await addRecord("presupuesto_general", { obraId, ...resumenObraImportado, fechaImportacion: hoyISO() }, setPresupuestoGeneral);
    for (const it of itemsObraImportados) {
      await addRecord("presupuesto_materiales", { obraId, ...it, origen: "Excel" }, setPresupuestoMateriales);
      if (it.subcategoria && !subcategoriasMat.some((s) => s.categoria === it.categoria && s.nombre === it.subcategoria)) {
        await addRecord("subcategorias_material", { categoria: it.categoria, nombre: it.subcategoria }, setSubcategoriasMat);
      }
      const existente = catalogoMateriales.find((m) => m.nombre.toLowerCase() === it.material.toLowerCase() && m.categoria === it.categoria);
      if (existente) {
        if (it.precioUnitario > 0) await updateRecord("catalogo_materiales", existente.id, { ultimoPrecio: it.precioUnitario, subcategoria: it.subcategoria || existente.subcategoria, unidad: it.unidad || existente.unidad }, setCatalogoMateriales);
      } else {
        await addRecord("catalogo_materiales", { categoria: it.categoria, subcategoria: it.subcategoria, tipo: "", nombre: it.material, unidad: it.unidad, ultimoPrecio: it.precioUnitario, ultimoProveedor: null }, setCatalogoMateriales);
      }
    }
  }

  async function submitNuevaObra(e) {
    e.preventDefault();
    const f = new FormData(e.target);
    setCreandoObra(true);
    const nuevaObra = await addRecord("obras", {
      nombre: f.get("nombre"),
      clienteId: f.get("clienteId") ? Number(f.get("clienteId")) : null,
      cliente: clientes.find((c) => c.id === Number(f.get("clienteId")))?.razonSocial || "",
      presupuesto: Number(f.get("presupuesto")) || resumenObraImportado?.precioTotalConIva || 0,
      meses: Number(f.get("meses")) || 1,
      inicio: f.get("inicio"),
      estado: "En curso",
      encargadoId: f.get("encargadoId") ? Number(f.get("encargadoId")) : null,
      diaCierre: f.get("diaCierre"),
      horaCierre: f.get("horaCierre"),
    }, setObras);
    if (nuevaObra) await importarPresupuestoAObra(nuevaObra.id);
    e.target.reset();
    quitarExcelNuevaObra();
    setCreandoObra(false);
    setShowObraForm(false);
  }
  function cambiarEstadoObra(obra, nuevoEstado) {
    if ((nuevoEstado === "Pausada" || nuevoEstado === "Finalizada") && !window.confirm(`¿${nuevoEstado === "Pausada" ? "Pausar" : "Finalizar"} "${obra.nombre}"? El personal que estaba afectado queda liberado para asignarse a otra obra.`)) return;
    updateRecord("obras", obra.id, { estado: nuevoEstado }, setObras);
  }

  // Manda la obra a la Papelera: libera las herramientas que estaban ahí (vuelven
  // a Oficina, Disponibles) y deja registrada la fecha de cancelación para que
  // el borrado automático (24hs) sepa cuándo purgarla. El personal se libera solo
  // porque obraActualDe() ya ignora cualquier obra que no esté "En curso".
  async function cancelarObra(obra) {
    if (!window.confirm(`¿Cancelar "${obra.nombre}"? Pasa a la Papelera: las herramientas que estén ahí vuelven a Oficina, el personal queda liberado, y los gastos/ingresos de esta obra dejan de contar en los balances. Se borra todo definitivamente en 24hs (podés restaurarla antes desde la Papelera). ¿Confirmar?`)) return false;
    const herramientasDeLaObra = herramientas.filter((h) => h.ubicacion === obra.nombre);
    for (const h of herramientasDeLaObra) {
      await updateRecord("herramientas", h.id, { ubicacion: "Oficina", estado: "Disponible", fechaUltimoCambioEstado: new Date().toISOString() }, setHerramientas);
    }
    await updateRecord("obras", obra.id, { estado: "Papelera", fechaCancelacion: new Date().toISOString() }, setObras);
    return true;
  }
  function restaurarObra(obra) {
    if (!window.confirm(`¿Restaurar "${obra.nombre}"? Vuelve a quedar "En curso" y sus gastos/ingresos vuelven a contar en los balances.`)) return;
    updateRecord("obras", obra.id, { estado: "En curso", fechaCancelacion: null }, setObras);
  }
  // Horas que le quedan a una obra en Papelera antes del borrado automático.
  function horasRestantesPapelera(obra) {
    if (!obra.fechaCancelacion) return null;
    const limite = new Date(obra.fechaCancelacion).getTime() + 24 * 3600 * 1000;
    return Math.max(0, Math.ceil((limite - Date.now()) / 3600000));
  }
  const [editandoObraId, setEditandoObraId] = useState(null);
  const [viewingObraId, setViewingObraId] = useState(null);
  function abrirObra(obra) {
    setViewingObraId(obra.id);
    setSelectedObraId(obra.id);
    setEditandoObraId(null);
  }
  function iniciarEdicionObra(obra) {
    setEditandoObraId(obra.id);
    setShowObraForm(false);
    setViewingObraId(null);
  }
  async function guardarEdicionObra(e, obra) {
    e.preventDefault();
    const f = new FormData(e.target);
    const nuevoClienteId = f.get("clienteId") ? Number(f.get("clienteId")) : null;
    await updateRecord("obras", obra.id, {
      nombre: f.get("nombre"),
      clienteId: nuevoClienteId,
      cliente: clientes.find((c) => c.id === nuevoClienteId)?.razonSocial || obra.cliente,
      presupuesto: Number(f.get("presupuesto")) || obra.presupuesto,
      encargadoId: f.get("encargadoId") ? Number(f.get("encargadoId")) : null,
    }, setObras);
    await importarPresupuestoAObra(obra.id);
    quitarExcelNuevaObra();
    setEditandoObraId(null);
  }
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

  function costoHoraDeCategoria(categoria, fechaStr) {
    const mes = (fechaStr || hoyISO()).slice(0, 7); // "YYYY-MM"
    const exacto = costosCategoria.find((c) => c.categoria === categoria && c.mes === mes);
    if (exacto) return exacto.costoHora || 0;
    // Si ese mes puntual no se cargó, usa el último mes cargado ANTERIOR (el sueldo no cambió hasta que se avise lo contrario).
    const anteriores = costosCategoria
      .filter((c) => c.categoria === categoria && c.mes && c.mes <= mes)
      .sort((a, b) => b.mes.localeCompare(a.mes));
    return anteriores[0]?.costoHora || 0;
  }

  function montoDe(a) {
    return (a.horas || 0) * costoHoraDeCategoria(categoriaDe(a.nombre), a.fecha);
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

  // Historial agrupado por semana, combinando pagos a Personal y avances a Tanteros
  const tanterosDeObraHistorial = tanteros.filter((t) => t.obraId === Number(obraHistorialId));
  const avancesDeObraHistorial = avancesTanteros.filter((av) => tanterosDeObraHistorial.some((t) => t.id === av.tanteroId));
  const totalHistoricoTanteros = avancesDeObraHistorial.reduce((s, av) => s + (av.monto || 0), 0);

  const semanasHistorial = {}; // semanaKey -> { personal: [...], tanteros: [...] }
  historialPagos.forEach((a) => {
    const key = claveSemana(a.fecha);
    if (!semanasHistorial[key]) semanasHistorial[key] = { personal: [], tanteros: [] };
    semanasHistorial[key].personal.push(a);
  });
  avancesDeObraHistorial.forEach((av) => {
    const key = claveSemana(av.fecha);
    if (!semanasHistorial[key]) semanasHistorial[key] = { personal: [], tanteros: [] };
    semanasHistorial[key].tanteros.push(av);
  });
  const semanasHistorialOrdenadas = Object.keys(semanasHistorial).sort((a, b) => new Date(b) - new Date(a));

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
  // El capataz arma el pedido de obra sin ver precios ni proveedor — eso es cosa de Logística.
  const canVerPreciosPedido = ROLES_VEN_PRECIOS_PEDIDO.includes(currentRole);

  function saldoCuenta(cuenta, formalidad) {
    const totalIngresos = ingresos
      .filter((i) => i.cuenta === cuenta && i.formalidad === formalidad && !obraIdsPapelera.has(i.obraId))
      .reduce((s, i) => s + (i.monto || 0), 0);
    const totalEgresos = comprasFacturas
      .filter((c) => c.cuenta === cuenta && c.formalidad === formalidad && !obraIdsPapelera.has(c.obraId))
      .reduce((s, c) => s + (c.monto || 0), 0);
    return totalIngresos - totalEgresos;
  }

  const saldosCuentas = CUENTAS.flatMap((cuenta) => FORMALIDADES.map((formalidad) => ({ cuenta, formalidad, saldo: saldoCuenta(cuenta, formalidad) })));
  const totalBlanco = FORMALIDADES[0] && CUENTAS.reduce((s, c) => s + saldoCuenta(c, "Blanco"), 0);
  const totalNegro = CUENTAS.reduce((s, c) => s + saldoCuenta(c, "Negro"), 0);

  const [filtroHerr, setFiltroHerr] = useState({ ubicacion: "Todas", estado: "Todos" });

  // N° de serie automático: letra de tipo + 3 letras de marca + N° correlativo (ej: E-BOS01)
  function generarNumeroSerie(categoria, marca) {
    const letraTipo = LETRA_TIPO_HERRAMIENTA[categoria] || "X";
    const letrasMarca = (marca || "GEN").replace(/[^a-zA-Zñ]/g, "").toUpperCase().padEnd(3, "X").slice(0, 3);
    const prefijo = `${letraTipo}-${letrasMarca}`;
    const existentes = herramientas.filter((h) => (h.numeroSerie || "").startsWith(prefijo));
    const numero = String(existentes.length + 1).padStart(2, "0");
    return `${prefijo}${numero}`;
  }

  const [vistaHerramientas, setVistaHerramientas] = useState("altoValor");
  const [viewingHerramientaId, setViewingHerramientaId] = useState(null);
  const viewingHerramienta = herramientas.find((h) => h.id === viewingHerramientaId) || null;

  // ---------- Alto Valor / Maquinaria (formulario controlado) ----------
  const emptyHerrForm = { categoria: CATEGORIAS_HERRAMIENTA[0], nombre: "", marca: "", maletin: "No", accesorios: "No", detalleAccesorios: "", observaciones: "" };
  const [herrForm, setHerrForm] = useState(emptyHerrForm);
  const [showAddNombreHerr, setShowAddNombreHerr] = useState(false);
  const [nuevoNombreHerr, setNuevoNombreHerr] = useState("");
  const [showAddMarca, setShowAddMarca] = useState(false);
  const [nuevaMarca, setNuevaMarca] = useState("");

  const nombresDisponiblesParaCategoria = catalogoNombresHerr.filter((c) => c.categoria === herrForm.categoria);

  function agregarNombreCatalogo() {
    if (!nuevoNombreHerr.trim()) return;
    addRecord("catalogo_nombres_herramienta", { categoria: herrForm.categoria, nombre: nuevoNombreHerr.trim() }, setCatalogoNombresHerr);
    setHerrForm((f) => ({ ...f, nombre: nuevoNombreHerr.trim() }));
    setNuevoNombreHerr("");
    setShowAddNombreHerr(false);
  }
  function agregarMarcaCatalogo() {
    if (!nuevaMarca.trim()) return;
    addRecord("catalogo_marcas", { nombre: nuevaMarca.trim() }, setCatalogoMarcas);
    setHerrForm((f) => ({ ...f, marca: nuevaMarca.trim() }));
    setNuevaMarca("");
    setShowAddMarca(false);
  }
  const [editingHerramientaId, setEditingHerramientaId] = useState(null);

  function cancelHerrForm() {
    setHerrForm(emptyHerrForm);
    setEditingHerramientaId(null);
    setShowHerrForm(false);
  }
  function startAddHerramienta() {
    setHerrForm(emptyHerrForm);
    setEditingHerramientaId(null);
    setShowHerrForm((v) => !v);
  }
  function startEditHerramienta(h) {
    setHerrForm({
      categoria: h.categoria || CATEGORIAS_HERRAMIENTA[0],
      nombre: h.nombre || "",
      marca: h.marca || "",
      maletin: h.maletin || "No",
      accesorios: h.accesorios || "No",
      detalleAccesorios: h.detalleAccesorios || "",
      observaciones: h.observaciones || "",
    });
    setEditingHerramientaId(h.id);
    setShowHerrForm(true);
    setViewingHerramientaId(null);
  }
  function submitHerrForm(e) {
    e.preventDefault();
    if (!herrForm.nombre) {
      alert("Elegí (o agregá) el nombre de la herramienta.");
      return;
    }
    const payload = {
      nombre: herrForm.nombre,
      marca: herrForm.marca,
      categoria: herrForm.categoria,
      maletin: herrForm.maletin,
      accesorios: herrForm.accesorios,
      detalleAccesorios: herrForm.detalleAccesorios,
      observaciones: herrForm.observaciones,
    };
    if (editingHerramientaId) {
      updateRecord("herramientas", editingHerramientaId, payload, setHerramientas);
    } else {
      addRecord("herramientas", {
        ...payload,
        numeroSerie: generarNumeroSerie(herrForm.categoria, herrForm.marca),
        ubicacion: "Oficina",
        estado: "Disponible",
        fechaUltimoCambioEstado: new Date().toISOString(),
      }, setHerramientas);
    }
    cancelHerrForm();
  }
  function cambiarEstadoHerramienta(h, nuevoEstado) {
    updateRecord("herramientas", h.id, { estado: nuevoEstado, fechaUltimoCambioEstado: new Date().toISOString() }, setHerramientas);
  }

  // ---------- Caja Herramientas Personal (por rubro, con asignación separada) ----------
  const emptyComboForm = { tipo: TIPOS_CAJA[0], items: [] };
  const [comboForm, setComboForm] = useState(emptyComboForm);
  const [comboItemDraft, setComboItemDraft] = useState({ nombre: "", cantidad: 1 });
  const [showComboForm, setShowComboForm] = useState(false);
  const [showAddChica, setShowAddChica] = useState(false);
  const [nuevaChica, setNuevaChica] = useState("");
  const [asignandoCajaId, setAsignandoCajaId] = useState(null);
  const [personaParaAsignar, setPersonaParaAsignar] = useState("");
  const [asignandoObraCajaId, setAsignandoObraCajaId] = useState(null);
  const [obraParaAsignar, setObraParaAsignar] = useState("");

  function nombreCaja(combo) {
    return `Caja ${combo.tipo} ${combo.numero}`;
  }
  function generarNumeroCaja(tipo) {
    const existentes = combosHerramientas.filter((c) => c.tipo === tipo);
    return existentes.length + 1;
  }
  function agregarChicaCatalogo() {
    if (!nuevaChica.trim()) return;
    addRecord("catalogo_herramientas_chicas", { nombre: nuevaChica.trim() }, setCatalogoChicas);
    setComboItemDraft((d) => ({ ...d, nombre: nuevaChica.trim() }));
    setNuevaChica("");
    setShowAddChica(false);
  }
  function agregarItemCombo() {
    if (!comboItemDraft.nombre.trim()) return;
    setComboForm((f) => ({
      ...f,
      items: [...f.items, { nombre: comboItemDraft.nombre.trim(), cantidad: Number(comboItemDraft.cantidad) || 1, estado: "Entregado" }],
    }));
    setComboItemDraft({ nombre: "", cantidad: 1 });
  }
  function quitarItemComboDraft(idx) {
    setComboForm((f) => ({ ...f, items: f.items.filter((_, i) => i !== idx) }));
  }
  function submitComboForm(e) {
    e.preventDefault();
    if (comboForm.items.length === 0) {
      alert("Agregá al menos una herramienta a la caja.");
      return;
    }
    addRecord("combos_herramientas", {
      tipo: comboForm.tipo,
      numero: generarNumeroCaja(comboForm.tipo),
      obraId: null,
      personaId: null,
      fecha: hoyISO(),
      items: comboForm.items,
    }, setCombosHerramientas);
    setComboForm(emptyComboForm);
    setShowComboForm(false);
  }
  function actualizarItemCombo(combo, idx, nuevoEstado) {
    const nuevosItems = combo.items.map((it, i) => (i === idx ? { ...it, estado: nuevoEstado } : it));
    updateRecord("combos_herramientas", combo.id, { items: nuevosItems }, setCombosHerramientas);
  }
  function confirmarAsignarObraCaja(combo) {
    if (!obraParaAsignar) return;
    updateRecord("combos_herramientas", combo.id, { obraId: Number(obraParaAsignar) }, setCombosHerramientas);
    setAsignandoObraCajaId(null);
    setObraParaAsignar("");
  }
  function confirmarAsignarCaja(combo) {
    if (!personaParaAsignar) return;
    updateRecord("combos_herramientas", combo.id, { personaId: Number(personaParaAsignar), fechaAsignacion: hoyISO() }, setCombosHerramientas);
    setAsignandoCajaId(null);
    setPersonaParaAsignar("");
  }
  function devolverCaja(combo) {
    if (!window.confirm(`¿Marcar "${nombreCaja(combo)}" como devuelta a depósito/oficina?`)) return;
    updateRecord("combos_herramientas", combo.id, { obraId: null, personaId: null }, setCombosHerramientas);
  }

  // ---------- Proveedores (incluye talleres de reparación) ----------
  const emptyProveedorForm = { razonSocial: "", cuit: "", domicilio: "", contacto: "", telefono: "", esTaller: "No" };
  const [proveedorForm, setProveedorForm] = useState(emptyProveedorForm);
  const [showProveedorForm, setShowProveedorForm] = useState(false);
  const talleres = proveedores.filter((p) => p.esTaller === "Sí");

  function submitProveedorForm(e) {
    e.preventDefault();
    addRecord("proveedores", { ...proveedorForm }, setProveedores);
    setProveedorForm(emptyProveedorForm);
    setShowProveedorForm(false);
  }
  function balanceProveedor(prov) {
    const facturas = comprasFacturas.filter((c) => c.proveedor === prov.razonSocial && !obraIdsPapelera.has(c.obraId));
    const totalFacturado = facturas.reduce((s, c) => s + (c.monto || 0), 0);
    const totalPagado = facturas.filter((c) => c.estado === "Pagada").reduce((s, c) => s + (c.monto || 0), 0);
    return { totalFacturado, totalPagado, saldo: totalFacturado - totalPagado, facturasPendientes: facturas.filter((c) => c.estado !== "Pagada") };
  }
  function marcarFacturaPagada(factura) {
    updateRecord("compras_facturas", factura.id, { estado: "Pagada" }, setComprasFacturas);
  }

  // ---------- Clientes ----------
  const emptyClienteForm = { razonSocial: "", cuit: "", domicilio: "", contacto: "", telefono: "" };
  const [clienteForm, setClienteForm] = useState(emptyClienteForm);
  const [showClienteForm, setShowClienteForm] = useState(false);
  const [vistaClientesProveedores, setVistaClientesProveedores] = useState("clientes");

  function submitClienteForm(e) {
    e.preventDefault();
    addRecord("clientes", { ...clienteForm }, setClientes);
    setClienteForm(emptyClienteForm);
    setShowClienteForm(false);
  }
  function balanceCliente(cli) {
    const obrasCliente = obras.filter((o) => o.clienteId === cli.id && o.estado !== "Papelera");
    const totalAcordado = obrasCliente.reduce((s, o) => s + (o.presupuesto || 0), 0);
    const totalCobrado = ingresos.filter((i) => obrasCliente.some((o) => o.id === i.obraId)).reduce((s, i) => s + (i.monto || 0), 0);
    return { obrasCliente, totalAcordado, totalCobrado, saldo: totalAcordado - totalCobrado };
  }

  // ---------- Remitos (traslados con doble aprobación: salida + recepción) ----------
  const emptyRemitoForm = { origen: "Oficina", destino: "", herramientaIds: [] };
  const [remitoForm, setRemitoForm] = useState(emptyRemitoForm);
  const [showRemitoForm, setShowRemitoForm] = useState(false);

  const herramientasEnOrigenRemito = herramientas.filter((h) => h.ubicacion === remitoForm.origen);
  const remitosPendientes = remitos.filter((r) => r.estado === "En tránsito" && r.herramientaIds?.length > 0);
  const remitosCompletados = remitos.filter((r) => r.estado === "Recibido" && r.herramientaIds?.length > 0);
  const remitosMaterialesPendientes = remitos.filter((r) => r.estado === "En tránsito" && r.materialItems?.length > 0);
  const remitosMaterialesCompletados = remitos.filter((r) => r.estado === "Recibido" && r.materialItems?.length > 0);

  function toggleHerramientaRemito(id) {
    setRemitoForm((f) => ({
      ...f,
      herramientaIds: f.herramientaIds.includes(id) ? f.herramientaIds.filter((x) => x !== id) : [...f.herramientaIds, id],
    }));
  }

  function submitRemitoForm(e) {
    e.preventDefault();
    if (!remitoForm.destino) {
      alert("Elegí el destino del remito.");
      return;
    }
    if (remitoForm.herramientaIds.length === 0) {
      alert("Elegí al menos una herramienta para enviar.");
      return;
    }
    const taller = talleres.find((t) => t.razonSocial === remitoForm.destino);
    addRecord("remitos", {
      fecha: hoyISO(),
      origen: remitoForm.origen,
      destino: remitoForm.destino,
      destinoEsTaller: !!taller,
      destinoProveedorId: taller ? taller.id : null,
      herramientaIds: remitoForm.herramientaIds,
      estado: "En tránsito",
      creadoPor: currentRole,
    }, setRemitos);
    setRemitoForm(emptyRemitoForm);
    setShowRemitoForm(false);
  }

  async function confirmarRecepcionRemito(remito) {
    // Los remitos de materiales pueden entregarse parciales (una parte a la obra, el resto a stock) — abren un formulario.
    if (remito.materialItems?.length > 0) {
      abrirRecepcionMaterial(remito);
      return;
    }
    if (!window.confirm(`¿Confirmar la recepción del remito en "${remito.destino}"?`)) return;
    if (remito.herramientaIds?.length > 0) {
      const nuevoEstado = remito.destinoEsTaller ? "En Reparación" : remito.destino === "Oficina" ? "Disponible" : "En Obra";
      await Promise.all(
        remito.herramientaIds.map((id) =>
          updateRecord("herramientas", id, { ubicacion: remito.destino, estado: nuevoEstado, fechaUltimoCambioEstado: new Date().toISOString() }, setHerramientas)
        )
      );
    }
    await updateRecord("remitos", remito.id, { estado: "Recibido", fechaRecepcion: hoyISO(), recibidoPor: currentRole }, setRemitos);
  }

  // ---------- Recepción parcial de materiales: parte a la obra, parte a Stock general de la empresa ----------
  const [recibiendoRemitoId, setRecibiendoRemitoId] = useState(null);
  const [cantidadesRecepcion, setCantidadesRecepcion] = useState([]);
  const [guardandoRecepcion, setGuardandoRecepcion] = useState(false);

  function abrirRecepcionMaterial(remito) {
    setRecibiendoRemitoId(remito.id);
    setCantidadesRecepcion(remito.materialItems.map((it) => it.cantidad)); // por defecto: entra todo a la obra
  }
  function actualizarCantidadRecepcion(idx, valor, cantidadMaxima) {
    const num = Math.max(0, Math.min(Number(valor) || 0, cantidadMaxima));
    setCantidadesRecepcion((arr) => arr.map((v, i) => (i === idx ? num : v)));
  }
  async function confirmarRecepcionMaterialConDivision(remito) {
    setGuardandoRecepcion(true);
    const pedido = remito.pedidoMaterialId ? pedidosMateriales.find((p) => p.id === remito.pedidoMaterialId) : null;

    for (let i = 0; i < remito.materialItems.length; i++) {
      const it = remito.materialItems[i];
      const cantidadAObra = cantidadesRecepcion[i] ?? it.cantidad;
      const cantidadAStock = Math.max(0, it.cantidad - cantidadAObra);
      const precio = it.precioUnitario || 0;

      if (cantidadAObra > 0) {
        await addRecord("compras_facturas", {
          fecha: hoyISO(),
          obraId: remito.destinoObraId ?? pedido?.obraId ?? null,
          ordenCompraId: null,
          proveedor: remito.origen,
          categoria: it.categoria || "Materiales",
          monto: cantidadAObra * precio,
          comprobante: pedido?.comprobante || "",
          estado: "Pendiente",
          formalidad: "Blanco",
          cuenta: "Banco",
        }, setComprasFacturas);
      }
      if (cantidadAStock > 0) {
        const facturaGeneral = await addRecord("compras_facturas", {
          fecha: hoyISO(),
          obraId: null,
          ordenCompraId: null,
          proveedor: remito.origen,
          categoria: it.categoria || "Materiales",
          monto: cantidadAStock * precio,
          comprobante: pedido?.comprobante || "",
          estado: "Pendiente",
          formalidad: "Blanco",
          cuenta: "Banco",
        }, setComprasFacturas);
        const loteExistente = stockMateriales.find((s) => s.material.toLowerCase() === it.material.toLowerCase() && s.categoria === it.categoria && s.precioUnitario === precio);
        if (loteExistente) {
          await updateRecord("stock_materiales", loteExistente.id, { cantidad: loteExistente.cantidad + cantidadAStock }, setStockMateriales);
        } else {
          await addRecord("stock_materiales", {
            material: it.material, categoria: it.categoria || "Materiales", subcategoria: it.subcategoria || "",
            unidad: it.unidad, cantidad: cantidadAStock, precioUnitario: precio,
            facturaGeneralId: facturaGeneral?.id ?? null, fechaIngreso: hoyISO(),
          }, setStockMateriales);
        }
      }
    }

    if (pedido) await updateRecord("pedidos_materiales", pedido.id, { estado: "Recibido" }, setPedidosMateriales);
    await updateRecord("remitos", remito.id, { estado: "Recibido", fechaRecepcion: hoyISO(), recibidoPor: currentRole }, setRemitos);
    setRecibiendoRemitoId(null);
    setGuardandoRecepcion(false);
  }

  // ---------- Stock general de materiales (lo que no entró completo a una obra) ----------
  const [asignandoStockId, setAsignandoStockId] = useState(null);
  const [obraParaStock, setObraParaStock] = useState("");
  const [cantidadParaStock, setCantidadParaStock] = useState(1);

  async function asignarStockAObra(lote) {
    const cantidad = Math.max(0, Math.min(Number(cantidadParaStock) || 0, lote.cantidad));
    if (cantidad <= 0 || !obraParaStock) return;
    const monto = cantidad * lote.precioUnitario;
    await updateRecord("stock_materiales", lote.id, { cantidad: lote.cantidad - cantidad }, setStockMateriales);
    if (lote.facturaGeneralId) {
      const facturaGeneral = comprasFacturas.find((c) => c.id === lote.facturaGeneralId);
      if (facturaGeneral) await updateRecord("compras_facturas", facturaGeneral.id, { monto: Math.max(0, facturaGeneral.monto - monto) }, setComprasFacturas);
    }
    await addRecord("compras_facturas", {
      fecha: hoyISO(),
      obraId: Number(obraParaStock),
      ordenCompraId: null,
      proveedor: "Stock interno",
      categoria: lote.categoria || "Materiales",
      monto,
      comprobante: "Reasignado desde stock",
      estado: "Pendiente",
      formalidad: "Blanco",
      cuenta: "Banco",
    }, setComprasFacturas);
    setAsignandoStockId(null);
    setObraParaStock("");
    setCantidadParaStock(1);
  }

  // ---------- Auditoría semanal de herramientas: formulario ----------
  const [obraAuditoriaId, setObraAuditoriaId] = useState(obras[0]?.id ?? "");
  const [tipoAuditoria, setTipoAuditoria] = useState("Cierre");
  const [presentesAuditoria, setPresentesAuditoria] = useState([]);
  const [obsAuditoria, setObsAuditoria] = useState("");
  const [showAuditoriaForm, setShowAuditoriaForm] = useState(false);

  const obraAuditoriaSel = obras.find((o) => o.id === Number(obraAuditoriaId));
  const herramientasDeObraAuditoria = herramientas.filter((h) => obraAuditoriaSel && h.ubicacion === obraAuditoriaSel.nombre);

  function abrirAuditoria(obraId, tipo) {
    setObraAuditoriaId(obraId);
    setTipoAuditoria(tipo);
    setPresentesAuditoria([]);
    setObsAuditoria("");
    setShowAuditoriaForm(true);
    setVistaHerramientas("auditoria");
    setTab("herramientas");
  }
  function togglePresenteAuditoria(id) {
    setPresentesAuditoria((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  }
  function submitAuditoria(e) {
    e.preventDefault();
    const faltantes = herramientasDeObraAuditoria.map((h) => h.id).filter((id) => !presentesAuditoria.includes(id));
    addRecord("auditorias_herramientas", {
      obraId: Number(obraAuditoriaId),
      fecha: hoyISO(),
      tipo: tipoAuditoria,
      realizadoPor: currentRole,
      herramientasPresentes: presentesAuditoria,
      herramientasFaltantes: faltantes,
      observaciones: obsAuditoria,
    }, setAuditorias);
    setShowAuditoriaForm(false);
  }

  // ---------- Calendario Corporativo: feriados y ficha horaria por obra ----------
  const emptyFeriadoForm = { fecha: hoyISO(), descripcion: "" };
  const [feriadoForm, setFeriadoForm] = useState(emptyFeriadoForm);
  const [showFeriadoForm, setShowFeriadoForm] = useState(false);

  function submitFeriadoForm(e) {
    e.preventDefault();
    if (!feriadoForm.descripcion.trim()) return;
    addRecord("feriados", { fecha: feriadoForm.fecha, descripcion: feriadoForm.descripcion.trim() }, setFeriados);
    setFeriadoForm(emptyFeriadoForm);
    setShowFeriadoForm(false);
  }
  function eliminarFeriado(f) {
    if (!window.confirm(`¿Sacar "${f.descripcion}" del calendario?`)) return;
    deleteRecord("feriados", f.id, setFeriados);
  }

  const [editandoHorarioObraId, setEditandoHorarioObraId] = useState(null);
  const [horarioForm, setHorarioForm] = useState({ diasLaborables: [], horaApertura: "08:00", diaCierre: "Viernes", horaCierre: "18:00" });

  function abrirEditarHorario(obra) {
    setEditandoHorarioObraId(obra.id);
    setHorarioForm({
      diasLaborables: obra.diasLaborables || [],
      horaApertura: obra.horaApertura || "08:00",
      diaCierre: obra.diaCierre || "Viernes",
      horaCierre: obra.horaCierre || "18:00",
    });
  }
  function toggleDiaLaborable(dia) {
    setHorarioForm((f) => ({
      ...f,
      diasLaborables: f.diasLaborables.includes(dia) ? f.diasLaborables.filter((d) => d !== dia) : [...f.diasLaborables, dia],
    }));
  }
  function guardarHorarioObra(e) {
    e.preventDefault();
    if (horarioForm.diasLaborables.length === 0) {
      alert("Elegí al menos un día laborable.");
      return;
    }
    updateRecord("obras", editandoHorarioObraId, { ...horarioForm }, setObras);
    setEditandoHorarioObraId(null);
  }

  // ---------- Materiales: subcategorías, catálogo de precios, presupuesto por Excel ----------
  const [vistaMateriales, setVistaMateriales] = useState("materiales");
  useEffect(() => {
    if (!canVerPreciosPedido && ["catalogo", "consolidar", "stock"].includes(vistaMateriales)) setVistaMateriales("materiales");
  }, [canVerPreciosPedido, vistaMateriales]);
  const [categoriaParaSubcat, setCategoriaParaSubcat] = useState(CATEGORIAS_PEDIDO[0]);
  const [nuevaSubcategoria, setNuevaSubcategoria] = useState("");
  const [categoriaParaTipo, setCategoriaParaTipo] = useState(CATEGORIAS_PEDIDO[0]);
  const [subcategoriaParaTipo, setSubcategoriaParaTipo] = useState("");
  const [nuevoTipo, setNuevoTipo] = useState("");
  const [obraPresupuestoId, setObraPresupuestoId] = useState(obras[0]?.id ?? "");
  const [mostrarListaHerramientas, setMostrarListaHerramientas] = useState(false);
  const [cantidadPedirStock, setCantidadPedirStock] = useState({});
  const [filtroEppParte, setFiltroEppParte] = useState("Todas");
  const [filtroEppTipo, setFiltroEppTipo] = useState("Todos");

  function agregarSubcategoria() {
    if (!nuevaSubcategoria.trim()) return;
    addRecord("subcategorias_material", { categoria: categoriaParaSubcat, nombre: nuevaSubcategoria.trim() }, setSubcategoriasMat);
    setNuevaSubcategoria("");
  }
  function agregarTipoMaterial() {
    if (!nuevoTipo.trim() || !subcategoriaParaTipo) return;
    addRecord("tipos_material", { categoria: categoriaParaTipo, subcategoria: subcategoriaParaTipo, nombre: nuevoTipo.trim() }, setTiposMaterial);
    setNuevoTipo("");
  }

  // ---------- Lector de la Planilla Interna de Costeo (estructura real de Concretar) ----------
  function buscarFilaYCol(rows, texto, desde = 0) {
    const t = texto.toUpperCase();
    for (let i = desde; i < rows.length; i++) {
      for (let j = 0; j < rows[i].length; j++) {
        if (typeof rows[i][j] === "string" && rows[i][j].toUpperCase().includes(t)) return [i, j];
      }
    }
    return [-1, -1];
  }
  function numeroDespuesDe(rows, texto, desde = 0) {
    const [fila, col] = buscarFilaYCol(rows, texto, desde);
    if (fila < 0) return { valor: 0, fila };
    for (let j = col + 1; j < rows[fila].length; j++) {
      if (typeof rows[fila][j] === "number") return { valor: rows[fila][j], fila };
    }
    return { valor: 0, fila };
  }
  function buscarHeaderItem(rows, desde) {
    for (let i = desde; i < rows.length; i++) {
      if (rows[i].some((v) => String(v).trim().toUpperCase() === "ITEM")) return i;
    }
    return -1;
  }
  function mapaColumnas(headerRow) {
    const mapa = {};
    (headerRow || []).forEach((val, idx) => {
      const t = String(val || "").toUpperCase().trim();
      if (t.includes("DESCRIP")) mapa.descripcion = idx;
      else if (t.includes("CANT")) mapa.cantidad = idx;
      else if (t.startsWith("UM")) mapa.unidad = idx;
      else if (t.includes("P/UNIT") || t.includes("COSTO UNIT")) mapa.precio = idx;
      else if (t.includes("P/TOTAL") || t.includes("COSTO TOTAL")) mapa.total = idx;
    });
    return mapa;
  }
  function leerItemsSeccion(rows, filaHeader, filaFin, categoria) {
    if (filaHeader < 0 || filaFin < 0) return [];
    const mapa = mapaColumnas(rows[filaHeader]);
    const items = [];
    let subcategoriaActual = "";
    for (let i = filaHeader + 1; i < filaFin; i++) {
      const fila = rows[i];
      const desc = mapa.descripcion !== undefined ? String(fila[mapa.descripcion] || "").trim() : "";
      if (!desc) continue;
      const precioRaw = mapa.precio !== undefined ? fila[mapa.precio] : 0;
      const precio = typeof precioRaw === "number" ? precioRaw : 0;
      const cantRaw = mapa.cantidad !== undefined ? fila[mapa.cantidad] : "";
      const sinCantidad = cantRaw === "" || cantRaw === null || cantRaw === undefined;
      if (precio === 0) {
        if (sinCantidad) subcategoriaActual = desc; // renglón "título" que agrupa el rubro siguiente
        continue; // sin precio cargado, no es un ítem pedible todavía
      }
      const cantidad = typeof cantRaw === "number" ? cantRaw : 0;
      const unidad = mapa.unidad !== undefined ? String(fila[mapa.unidad] || "").trim() : "";
      const totalRaw = mapa.total !== undefined ? fila[mapa.total] : null;
      const total = typeof totalRaw === "number" ? totalRaw : cantidad * precio;
      items.push({ categoria, subcategoria: subcategoriaActual, tipo: "", material: desc, unidad, cantidad, precioUnitario: precio, total });
    }
    return items;
  }
  function parsePresupuestoGeneral(workbook) {
    const nombreHoja = workbook.SheetNames.includes("Planilla final") ? "Planilla final" : workbook.SheetNames[0];
    const sheet = workbook.Sheets[nombreHoja];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    const { valor: totalManoObra, fila: filaSubtotalMO } = numeroDespuesDe(rows, "SUBTOTAL M.O.");
    const [filaInicioEquipos] = buscarFilaYCol(rows, "EQUIPOS/OTROS", filaSubtotalMO + 1);
    const filaHeaderEquipos = buscarHeaderItem(rows, filaInicioEquipos);
    const { valor: totalEquipos, fila: filaSubtotalEquipos } = numeroDespuesDe(rows, "SUBTOTAL EQUIPOS", filaHeaderEquipos);
    const filaHeaderMateriales = buscarHeaderItem(rows, filaSubtotalEquipos);
    const { valor: totalMateriales, fila: filaSubtotalMateriales } = numeroDespuesDe(rows, "SUBTOTAL MATERIALES", filaHeaderMateriales);

    // Sección opcional "Herramientas" — no todos los presupuestos la tienen. Si no aparece, se sigue de largo sin error.
    const [filaSubtotalHerr, colSubtotalHerr] = buscarFilaYCol(rows, "SUBTOTAL HERRAMIENTAS", filaSubtotalMateriales);
    let itemsHerramientas = [];
    let totalHerramientas = 0;
    if (filaSubtotalHerr >= 0) {
      // Busca el header ITEM más cercano ANTES de ese subtotal (puede estar antes o después de Materiales).
      let filaHeaderHerr = -1;
      for (let i = filaSubtotalHerr - 1; i >= 0; i--) {
        if (rows[i].some((v) => String(v).trim().toUpperCase() === "ITEM")) { filaHeaderHerr = i; break; }
      }
      const r = numeroDespuesDe(rows, "SUBTOTAL HERRAMIENTAS", filaSubtotalMateriales);
      totalHerramientas = r.valor;
      itemsHerramientas = leerItemsSeccion(rows, filaHeaderHerr, filaSubtotalHerr, "Herramientas");
    }

    // La cascada de Impuestos/Ganancia/IVA es opcional (obra "en negro" sin factura puede no tenerla).
    const filaDesdeParaCascada = Math.max(filaSubtotalMateriales, filaSubtotalHerr);
    const { valor: precioTotalSinIvaRaw, fila: filaPrecioSinIva } = numeroDespuesDe(rows, "PRECIO TOTAL $ SIN IVA", filaDesdeParaCascada);
    const { valor: precioTotalConIvaRaw } = numeroDespuesDe(rows, "PRECIO TOTAL $ CON IVA", filaPrecioSinIva >= 0 ? filaPrecioSinIva : filaDesdeParaCascada);
    const subtotal123 = totalManoObra + totalEquipos + totalMateriales + totalHerramientas;
    // Si no se encontró la cascada de IVA (obra en negro), el total es directo la suma de las 3-4 secciones.
    const precioTotalSinIva = precioTotalSinIvaRaw || subtotal123;
    const precioTotalConIva = precioTotalConIvaRaw || precioTotalSinIva;

    const itemsEquipos = leerItemsSeccion(rows, filaHeaderEquipos, filaSubtotalEquipos, "Equipos");
    const itemsMateriales = leerItemsSeccion(rows, filaHeaderMateriales, filaSubtotalMateriales, "Materiales");

    return {
      resumen: { totalManoObra, totalEquipos, totalHerramientas, totalMateriales, precioTotalSinIva, precioTotalConIva },
      items: [...itemsEquipos, ...itemsHerramientas, ...itemsMateriales],
    };
  }

  function eliminarLineaPresupuesto(id) {
    if (!window.confirm("¿Sacar esta línea del presupuesto?")) return;
    deleteRecord("presupuesto_materiales", id, setPresupuestoMateriales);
  }

  // ---------- Pedidos de materiales (armados desde el presupuesto + ítems manuales) ----------
  const [seleccionPresupuesto, setSeleccionPresupuesto] = useState([]);
  const [showPedidoForm, setShowPedidoForm] = useState(false);
  const [pedidoItems, setPedidoItems] = useState([]);
  const [pedidoProveedor, setPedidoProveedor] = useState("");
  const [pedidoFechaNecesaria, setPedidoFechaNecesaria] = useState("");
  const [itemManualDraft, setItemManualDraft] = useState({ categoria: "Materiales", subcategoria: "", tipo: "", material: "", unidad: "", cantidad: 1, precioUnitario: 0 });
  const [enviandoPedido, setEnviandoPedido] = useState(false);

  function toggleSeleccionPresupuesto(id) {
    setSeleccionPresupuesto((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  }

  // Arma (o completa) el pedido en curso a partir de las líneas de presupuesto tildadas.
  // No pisa lo que ya se haya agregado desde otro rubro (manual o del catálogo de EPPs) —
  // así podés ir sumando materiales, equipos/herramientas y EPPs al mismo pedido.
  function abrirArmadoPedido() {
    const seleccionadas = presupuestoMateriales.filter((m) => seleccionPresupuesto.includes(m.id));
    const itemsPresupuesto = seleccionadas.map((m) => ({
      presupuestoId: m.id, categoria: m.categoria, subcategoria: m.subcategoria, tipo: m.tipo,
      material: m.material, unidad: m.unidad, cantidad: m.cantidad, precioUnitario: m.precioUnitario,
    }));
    const itemsPrevios = pedidoItems.filter((it) => !it.presupuestoId);
    // Sugiere el proveedor más repetido según el catálogo (última vez que se compró ese material).
    const conteo = {};
    seleccionadas.forEach((m) => {
      const enCatalogo = catalogoMateriales.find((c) => c.nombre.toLowerCase() === m.material.toLowerCase() && c.categoria === m.categoria);
      if (enCatalogo?.ultimoProveedor) conteo[enCatalogo.ultimoProveedor] = (conteo[enCatalogo.ultimoProveedor] || 0) + 1;
    });
    const sugerido = Object.entries(conteo).sort((a, b) => b[1] - a[1])[0]?.[0] || pedidoProveedor;
    setPedidoItems([...itemsPresupuesto, ...itemsPrevios]);
    setPedidoProveedor(sugerido);
    setShowPedidoForm(true);
  }

  function actualizarCantidadPedido(idx, campo, valor) {
    setPedidoItems((items) => items.map((it, i) => (i === idx ? { ...it, [campo]: valor } : it)));
  }
  function quitarItemPedido(idx) {
    setPedidoItems((items) => items.filter((_, i) => i !== idx));
  }
  function agregarItemManualPedido() {
    if (!itemManualDraft.material.trim()) return;
    const categoriaFinal = canVerPreciosPedido ? itemManualDraft.categoria : (CATEGORIA_DE_VISTA[vistaMateriales] || itemManualDraft.categoria);
    const precioFinal = canVerPreciosPedido ? itemManualDraft.precioUnitario : 0;
    setPedidoItems((items) => [...items, { presupuestoId: null, ...itemManualDraft, categoria: categoriaFinal, precioUnitario: precioFinal, material: itemManualDraft.material.trim() }]);
    setItemManualDraft((d) => ({ ...d, subcategoria: d.subcategoria, tipo: "", material: "", unidad: "", cantidad: 1, precioUnitario: 0 }));
    setShowPedidoForm(true);
  }
  // Agrega un ítem del catálogo (ej. un EPP ya cargado antes) directo al pedido en curso.
  function agregarCatalogoAlPedido(item) {
    setPedidoItems((items) => [...items, {
      presupuestoId: null, categoria: item.categoria, subcategoria: item.subcategoria, tipo: item.tipo || "",
      material: item.nombre, unidad: item.unidad || "und.", cantidad: 1, precioUnitario: item.ultimoPrecio || 0,
    }]);
    setShowPedidoForm(true);
  }
  // Pedido puntual sobre una herramienta real del inventario (ej. reponer una rota,
  // o comprar una segunda igual) — a diferencia del catálogo, no tiene precio de referencia.
  function agregarHerramientaAlPedido(h) {
    setPedidoItems((items) => [...items, {
      presupuestoId: null, categoria: "Herramientas", subcategoria: h.categoria, tipo: "", tipoEquipo: "Propio",
      material: `${h.nombre} (${h.numeroSerie || "s/n"})`, unidad: "und.", cantidad: 1, precioUnitario: 0,
    }]);
    setShowPedidoForm(true);
  }

  async function confirmarPedido() {
    if (pedidoItems.length === 0) {
      alert("Agregá al menos un ítem al pedido.");
      return;
    }
    if (!pedidoFechaNecesaria) {
      alert("Elegí para cuándo lo necesitás.");
      return;
    }
    setEnviandoPedido(true);
    const itemsFinales = pedidoItems.map((it) => ({ ...it, total: (Number(it.cantidad) || 0) * (Number(it.precioUnitario) || 0) }));
    const totalPedido = itemsFinales.reduce((s, it) => s + it.total, 0);
    const pedidoCreado = await addRecord("pedidos_materiales", {
      obraId: obraPresupuestoId === "general" ? null : Number(obraPresupuestoId),
      fecha: hoyISO(),
      fechaNecesaria: pedidoFechaNecesaria,
      proveedor: pedidoProveedor,
      estado: "Solicitado",
      solicitadoPor: currentRole,
      items: itemsFinales,
      total: totalPedido,
    }, setPedidosMateriales);
    if (pedidoCreado) {
      for (const it of itemsFinales) {
        if (it.presupuestoId) {
          // Las líneas del presupuesto que se usaron quedan marcadas como "ya pedidas".
          await updateRecord("presupuesto_materiales", it.presupuestoId, { pedidoId: pedidoCreado.id }, setPresupuestoMateriales);
        } else if (!catalogoMateriales.some((c) => c.nombre.toLowerCase() === it.material.toLowerCase() && c.categoria === it.categoria)) {
          // Lo que se tipeó a mano y todavía no estaba en el catálogo queda guardado para la próxima vez.
          await addRecord("catalogo_materiales", {
            categoria: it.categoria, subcategoria: it.subcategoria || "", tipo: it.tipo || "", nombre: it.material,
            unidad: it.unidad, ultimoPrecio: it.precioUnitario, ultimoProveedor: pedidoProveedor || null,
          }, setCatalogoMateriales);
        }
      }
    }
    setSeleccionPresupuesto([]);
    setPedidoItems([]);
    setPedidoProveedor("");
    setPedidoFechaNecesaria("");
    setShowPedidoForm(false);
    setEnviandoPedido(false);
  }

  async function recibirPedidoMaterial(pedido) {
    await updateRecord("pedidos_materiales", pedido.id, { estado: "Recibido" }, setPedidosMateriales);
    await addRecord("compras_facturas", {
      fecha: hoyISO(),
      obraId: pedido.obraId,
      ordenCompraId: null,
      proveedor: pedido.proveedor || "Sin especificar",
      categoria: "Materiales",
      monto: pedido.total,
      comprobante: "",
      estado: "Pendiente",
      formalidad: "Blanco",
      cuenta: "Banco",
    }, setComprasFacturas);
    // Actualiza el "último proveedor" de cada material del catálogo, para que la próxima sugerencia de consolidación sea más precisa.
    if (pedido.proveedor) {
      for (const it of pedido.items) {
        const existente = catalogoMateriales.find((m) => m.nombre.toLowerCase() === it.material.toLowerCase() && m.categoria === it.categoria);
        if (existente) await updateRecord("catalogo_materiales", existente.id, { ultimoProveedor: pedido.proveedor }, setCatalogoMateriales);
      }
    }
  }
  async function marcarPedidoRecibido(pedido) {
    if (!window.confirm(`¿Marcar este pedido como recibido en obra? Esto suma ${fmtARS(pedido.total)} a la curva de inversión real de la obra.`)) return;
    await recibirPedidoMaterial(pedido);
  }
  const puedeAprobarPedidos = currentRole === "Gerente";
  function aprobarPedidoMaterial(pedido) {
    updateRecord("pedidos_materiales", pedido.id, { estado: "Aprobado", aprobadoPor: currentRole, fechaAprobacion: hoyISO() }, setPedidosMateriales);
  }
  function rechazarPedidoMaterial(pedido) {
    if (!window.confirm("¿Rechazar este pedido? El capataz va a tener que volver a armarlo si todavía lo necesita.")) return;
    updateRecord("pedidos_materiales", pedido.id, { estado: "Rechazado", aprobadoPor: currentRole, fechaAprobacion: hoyISO() }, setPedidosMateriales);
  }

  // Pedidos "Compra general" de Epps/Consumibles (sin obra) no pasan por remito: al
  // recibirse quedan directo en el depósito (Stock), listos para que cada obra los pida.
  async function recibirPedidoGeneralAStock(pedido) {
    if (!window.confirm(`¿Marcar recibido? ${fmtARS(pedido.total)} en ${pedido.items.length} ítem(s) van a quedar en el depósito (pestaña Stock) hasta que alguna obra los pida.`)) return;
    for (const it of pedido.items) {
      const precio = it.precioUnitario || 0;
      const facturaGeneral = await addRecord("compras_facturas", {
        fecha: hoyISO(), obraId: null, ordenCompraId: null, proveedor: pedido.proveedor || "Sin especificar",
        categoria: it.categoria, monto: (Number(it.cantidad) || 0) * precio, comprobante: pedido.comprobante || "",
        estado: "Pendiente", formalidad: "Blanco", cuenta: "Banco",
      }, setComprasFacturas);
      const loteExistente = stockMateriales.find((s) => s.material.toLowerCase() === it.material.toLowerCase() && s.categoria === it.categoria && s.precioUnitario === precio);
      if (loteExistente) {
        await updateRecord("stock_materiales", loteExistente.id, { cantidad: loteExistente.cantidad + (Number(it.cantidad) || 0) }, setStockMateriales);
      } else {
        await addRecord("stock_materiales", {
          material: it.material, categoria: it.categoria, subcategoria: it.subcategoria || "",
          unidad: it.unidad, cantidad: Number(it.cantidad) || 0, precioUnitario: precio,
          facturaGeneralId: facturaGeneral?.id ?? null, fechaIngreso: hoyISO(),
        }, setStockMateriales);
      }
    }
    await updateRecord("pedidos_materiales", pedido.id, { estado: "Recibido" }, setPedidosMateriales);
  }

  // Una obra pide una cantidad de algo que ya está en el depósito: descuenta el stock
  // y mueve ese gasto al centro de costos de la obra (reduciendo la factura general
  // en la misma proporción, para no contarlo dos veces).
  async function pedirDeStockParaObra(lote, cantidad, obraId) {
    const cant = Math.max(0, Math.min(Number(cantidad) || 0, lote.cantidad));
    if (cant <= 0 || !obraId) return;
    const monto = cant * lote.precioUnitario;
    await updateRecord("stock_materiales", lote.id, { cantidad: lote.cantidad - cant }, setStockMateriales);
    if (lote.facturaGeneralId) {
      const facturaGeneral = comprasFacturas.find((c) => c.id === lote.facturaGeneralId);
      if (facturaGeneral) await updateRecord("compras_facturas", facturaGeneral.id, { monto: Math.max(0, facturaGeneral.monto - monto) }, setComprasFacturas);
    }
    await addRecord("compras_facturas", {
      fecha: hoyISO(), obraId, ordenCompraId: null, proveedor: "Depósito interno", categoria: lote.categoria,
      monto, comprobante: "Pedido desde depósito", estado: "Pendiente", formalidad: "Blanco", cuenta: "Banco",
    }, setComprasFacturas);
  }

  // ---------- Orden de Compra en PDF ----------
  function generarOrdenCompraPDF(pedido) {
    const obra = obras.find((o) => o.id === pedido.obraId);
    const prov = proveedores.find((p) => p.razonSocial === pedido.proveedor);
    const doc = new jsPDF({ unit: "mm", format: "a4" });

    doc.setFillColor(2, 29, 52);
    doc.rect(0, 0, 210, 28, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.text("ORDEN DE COMPRA", 14, 14);
    doc.setFontSize(10);
    doc.text(`N° ${String(pedido.id).padStart(5, "0")}  ·  ${fmtFecha(pedido.fecha)}`, 14, 21);
    doc.text("Grupo Concretar S.A.S.", 150, 14);

    doc.setTextColor(20, 20, 20);
    doc.setFontSize(10);
    let y = 38;
    doc.setFont(undefined, "bold");
    doc.text("Proveedor", 14, y);
    doc.setFont(undefined, "normal");
    doc.text(pedido.proveedor || "Sin especificar", 40, y);
    y += 6;
    if (prov?.cuit) { doc.text(`CUIT: ${prov.cuit}`, 14, y); y += 6; }
    if (prov?.domicilio) { doc.text(`Domicilio: ${prov.domicilio}`, 14, y); y += 6; }
    if (prov?.contacto) { doc.text(`Contacto: ${prov.contacto}${prov.telefono ? " — " + prov.telefono : ""}`, 14, y); y += 6; }

    doc.setFont(undefined, "bold");
    doc.text("Obra de destino", 120, 38);
    doc.setFont(undefined, "normal");
    doc.text(obra?.nombre || "-", 120, 44);
    if (pedido.fechaNecesaria) doc.text(`Necesario para: ${fmtFecha(pedido.fechaNecesaria)}`, 120, 50);

    autoTable(doc, {
      startY: Math.max(y, 56) + 4,
      head: [["Material", "Unidad", "Cantidad", "P. Unitario (sin IVA)", "Total"]],
      body: pedido.items.map((it) => [it.material, it.unidad || "-", String(it.cantidad), fmtARS(it.precioUnitario), fmtARS(it.total)]),
      foot: [["", "", "", "TOTAL", fmtARS(pedido.total)]],
      headStyles: { fillColor: [2, 29, 52] },
      footStyles: { fillColor: [245, 245, 244], textColor: [20, 20, 20], fontStyle: "bold" },
      styles: { fontSize: 9 },
    });

    const finalY = doc.lastAutoTable.finalY || 100;
    doc.setFontSize(8);
    doc.setTextColor(120, 120, 120);
    doc.text("Precios estimados según presupuesto — sujetos a corrección con la factura real.", 14, finalY + 8);

    doc.save(`OC_${(obra?.nombre || "obra").replace(/\s+/g, "_")}_${pedido.id}.pdf`);
  }

  // ---------- Carga de factura real (corrige precios, imputa el gasto final a la obra) ----------
  const [facturandoPedidoId, setFacturandoPedidoId] = useState(null);
  const [itemsFacturaDraft, setItemsFacturaDraft] = useState([]);
  const [comprobanteDraft, setComprobanteDraft] = useState("");
  const [totalFacturaRapido, setTotalFacturaRapido] = useState(0);
  const [facturaRevision, setFacturaRevision] = useState(0);
  const [guardandoFactura, setGuardandoFactura] = useState(false);
  // El capataz arma el pedido sin proveedor — Logística confirma acá el que se cotizó
  // inicialmente (sugerido según el catálogo) o elige otro antes de cargar los precios.
  const [proveedorFacturaDraft, setProveedorFacturaDraft] = useState("");

  function abrirCargaFactura(pedido) {
    setFacturandoPedidoId(pedido.id);
    setItemsFacturaDraft(pedido.items.map((it) => ({ ...it })));
    setComprobanteDraft("");
    setTotalFacturaRapido(0);
    setFacturaRevision(0);
    setProveedorFacturaDraft(pedido.proveedor || "");
  }
  function actualizarPrecioFactura(idx, valor) {
    setItemsFacturaDraft((items) => items.map((it, i) => (i === idx ? { ...it, precioUnitario: valor, total: (Number(it.cantidad) || 0) * (Number(valor) || 0) } : it)));
  }
  // Logística marca cada equipo/herramienta como propio (ya lo tenemos, sin costo) o
  // alquilado (el alquiler se carga como gasto real de la obra).
  function actualizarTipoEquipoFactura(idx, valor) {
    setItemsFacturaDraft((items) => items.map((it, i) => (i === idx ? { ...it, tipoEquipo: valor, precioUnitario: valor === "Propio" ? 0 : it.precioUnitario, total: valor === "Propio" ? 0 : it.total } : it)));
  }
  // Para cuando hay apuro y no vale la pena discriminar precio por ítem: reparte un
  // total único entre los ítems a prorrata de cantidad, sin tocar los que son "Propio"
  // (esos no tienen costo, sea cual sea el total que se cargue).
  function aplicarTotalFactura(totalIngresado) {
    const total = Number(totalIngresado) || 0;
    const itemsAPrecio = itemsFacturaDraft.filter((it) => it.tipoEquipo !== "Propio");
    const sumaCantidades = itemsAPrecio.reduce((s, it) => s + (Number(it.cantidad) || 0), 0);
    setItemsFacturaDraft((items) => items.map((it) => {
      if (it.tipoEquipo === "Propio") return it;
      const cant = Number(it.cantidad) || 0;
      const proporcion = sumaCantidades > 0 ? cant / sumaCantidades : 1 / (itemsAPrecio.length || 1);
      const precio = cant > 0 ? (total * proporcion) / cant : 0;
      return { ...it, precioUnitario: precio, total: precio * cant };
    }));
    setFacturaRevision((r) => r + 1);
  }
  async function confirmarFacturaReal(pedido) {
    if (itemsFacturaDraft.some((it) => ["Equipos", "Herramientas"].includes(it.categoria) && !it.tipoEquipo)) {
      alert("Marcá cada equipo/herramienta como Propio o Alquilado antes de confirmar.");
      return;
    }
    setGuardandoFactura(true);
    const itemsFinal = itemsFacturaDraft.map((it) => ({ ...it, total: (Number(it.cantidad) || 0) * (Number(it.precioUnitario) || 0) }));
    const totalReal = itemsFinal.reduce((s, it) => s + it.total, 0);
    await updateRecord("pedidos_materiales", pedido.id, { items: itemsFinal, total: totalReal, proveedor: proveedorFacturaDraft, estado: "Facturado", comprobante: comprobanteDraft }, setPedidosMateriales);
    // El gasto real se imputa recién cuando se recibe el remito (ahí se sabe cuánto entró
    // a esta obra y cuánto quedó en stock general) — acá solo corregimos los precios.
    for (const it of itemsFinal) {
      const existente = catalogoMateriales.find((m) => m.nombre.toLowerCase() === it.material.toLowerCase() && m.categoria === it.categoria);
      if (existente && it.precioUnitario > 0) {
        await updateRecord("catalogo_materiales", existente.id, { ultimoPrecio: it.precioUnitario, ultimoProveedor: proveedorFacturaDraft || existente.ultimoProveedor }, setCatalogoMateriales);
      }
    }
    setFacturandoPedidoId(null);
    setGuardandoFactura(false);
  }

  // ---------- Consolidación de pedidos entre obras + generación de remitos por proveedor ----------
  const idsPedidosConRemito = new Set(remitos.filter((r) => r.pedidoMaterialId).map((r) => r.pedidoMaterialId));
  // Los pedidos "compra general" (obraId null) no tienen una obra destino a la que
  // mandar un remito — se reciben directo al depósito (recibirPedidoGeneralAStock).
  const pedidosSinEnviar = pedidosMateriales.filter((p) => (p.estado === "Aprobado" || p.estado === "Facturado") && p.obraId != null && !idsPedidosConRemito.has(p.id));

  function proveedorSugerido(pedido) {
    if (pedido.proveedor) return pedido.proveedor;
    const conteo = {};
    for (const it of pedido.items) {
      const mat = catalogoMateriales.find((m) => m.nombre.toLowerCase() === it.material.toLowerCase() && m.categoria === it.categoria);
      if (mat?.ultimoProveedor) conteo[mat.ultimoProveedor] = (conteo[mat.ultimoProveedor] || 0) + 1;
    }
    const entradas = Object.entries(conteo).sort((a, b) => b[1] - a[1]);
    return entradas[0]?.[0] || null;
  }

  const gruposPorProveedor = {}; // proveedor (o "Sin proveedor") -> [pedidos]
  pedidosSinEnviar.forEach((p) => {
    const clave = p.proveedor || proveedorSugerido(p) || "Sin proveedor asignado";
    if (!gruposPorProveedor[clave]) gruposPorProveedor[clave] = [];
    gruposPorProveedor[clave].push(p);
  });

  function asignarProveedorPedido(pedido, proveedor) {
    updateRecord("pedidos_materiales", pedido.id, { proveedor }, setPedidosMateriales);
  }

  function generarRemitoDesdePedido(pedido, proveedor) {
    const obra = obras.find((o) => o.id === pedido.obraId);
    if (!obra) return;
    addRecord("remitos", {
      fecha: hoyISO(),
      origen: proveedor,
      destino: obra.nombre,
      destinoObraId: obra.id,
      destinoEsTaller: false,
      destinoProveedorId: null,
      herramientaIds: [],
      materialItems: pedido.items,
      pedidoMaterialId: pedido.id,
      estado: "En tránsito",
      creadoPor: currentRole,
    }, setRemitos);
  }

  async function generarRemitosDelGrupo(proveedor, pedidosGrupo) {
    for (const p of pedidosGrupo) {
      if (!p.proveedor) await updateRecord("pedidos_materiales", p.id, { proveedor }, setPedidosMateriales);
      generarRemitoDesdePedido({ ...p, proveedor }, proveedor);
    }
  }

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
    <div className="relative flex h-full min-h-[720px] w-full overflow-hidden md:rounded-xl md:border md:border-stone-200 bg-stone-100 text-slate-800" style={{ fontFamily: BRAND.font }}>
      {/* Barra superior solo en celular */}
      <div className="fixed inset-x-0 top-0 z-30 flex items-center justify-between px-4 py-3 text-slate-100 md:hidden" style={{ backgroundColor: BRAND.navy900 }}>
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
        className={`fixed inset-y-0 left-0 z-40 flex w-64 shrink-0 -translate-x-full flex-col text-slate-100 transition-transform duration-200 md:static md:z-auto md:w-56 md:translate-x-0 ${
          mobileNavOpen ? "translate-x-0" : ""
        }`}
        style={{ backgroundColor: BRAND.navy900 }}
      >
        <div className="flex items-start justify-between border-b border-white/10 px-5 py-5">
          <div>
            <div className="text-2xl font-extrabold uppercase tracking-tight text-white">Concretar</div>
            <div className="mt-0.5 text-sm font-medium text-slate-300">Construyamos</div>
            <img src="/gc-logo-white.png" alt="Grupo Concretar" className="mt-3 h-7 w-auto" />
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
              style={tab === item.id ? { backgroundColor: BRAND.navy400 } : undefined}
              className={`flex w-full items-center justify-between rounded-md px-3 py-2.5 text-left text-sm font-medium transition ${
                tab === item.id ? "text-white" : "text-slate-300 hover:bg-white/5 hover:text-white"
              }`}
            >
              <span className="flex items-center gap-3"><item.icon size={17} />{item.label}</span>
              {item.id === "general" && totalAlertas > 0 && (
                <span className="flex h-5 w-5 items-center justify-center rounded-full bg-rose-500 text-[10px] font-bold text-white">{totalAlertas}</span>
              )}
            </button>
          ))}
        </nav>
        <div className="border-t border-white/10 px-5 py-4">
          <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-slate-400">Ingresando como</div>
          <select
            value={currentRole}
            onChange={(e) => setCurrentRole(e.target.value)}
            className="w-full rounded-md border border-white/20 bg-black/20 px-2 py-1.5 text-xs text-slate-100"
          >
            {ROLES.map((r) => <option key={r}>{r}</option>)}
          </select>
          <div className="mt-1 text-[10px] text-slate-500">Simula el login hasta que armemos uno real</div>
        </div>
        <div className="border-t border-white/10 px-5 py-4 text-[11px] text-slate-500">
          {isSupabaseConfigured ? (
            <span className="flex items-center gap-1.5 text-emerald-400"><Database size={12} /> Conectado a Supabase</span>
          ) : (
            <span className="flex items-center gap-1.5 text-amber-400"><Database size={12} /> Modo demo — configurá Supabase</span>
          )}
        </div>
      </aside>

      {/* Main content */}
      <main className="mt-12 flex-1 overflow-y-auto p-4 md:mt-0 md:p-8">
        {tab === "general" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">General</h2>

            <Panel title="Alertas de todas las obras">
              {totalAlertas === 0 ? (
                <div className="flex items-center gap-2 text-sm text-emerald-700"><CheckCircle2 size={16} /> Todo en orden, sin pendientes críticos.</div>
              ) : (
                <div className="grid grid-cols-1 gap-3 lg:grid-cols-2 xl:grid-cols-3">
                  {ocPendientesAprobacion.length > 0 && (
                    <AlertCard tone="rose" icon={AlertTriangle} title={`${ocPendientesAprobacion.length} orden(es) de compra por encima de ${fmtARS(UMBRAL_APROBACION_OC)} esperando aprobación.`} />
                  )}
                  {herramientasAtencion.length > 0 && (
                    <AlertCard tone="amber" icon={AlertTriangle} title={`${herramientasAtencion.length} herramienta(s) en mal estado o rota(s) — mandar a reparar.`}>
                      <ul className="space-y-0.5 text-xs">
                        {herramientasAtencion.slice(0, 5).map((h) => <li key={h.id} className="truncate">{h.nombre} ({h.numeroSerie}) — {h.estado}</li>)}
                      </ul>
                    </AlertCard>
                  )}
                  {herramientasReparadasRecientes.length > 0 && (
                    <AlertCard tone="emerald" icon={CheckCircle2} title={`${herramientasReparadasRecientes.length} herramienta(s) reparada(s) recientemente.`}>
                      <ul className="space-y-0.5 text-xs">
                        {herramientasReparadasRecientes.slice(0, 5).map((h) => <li key={h.id} className="truncate">{h.nombre} ({h.numeroSerie}) — ya disponible</li>)}
                      </ul>
                    </AlertCard>
                  )}
                  {obrasEnVentanaCierre.map((o) => (
                    <AlertCard key={`cierre-${o.id}`} tone="rose" icon={AlertTriangle} title={`Falta menos de 1hs para el cierre de "${o.nombre}" (${o.horaCierre}) — hacé el control de herramientas.`}>
                      <button onClick={() => abrirAuditoria(o.id, "Cierre")} className="text-xs font-semibold underline hover:no-underline">
                        Hacer control de cierre ahora →
                      </button>
                    </AlertCard>
                  ))}
                  {obrasSinAperturaLunes.map((o) => (
                    <AlertCard key={`apertura-${o.id}`} tone="amber" icon={AlertTriangle} title={`Falta validar el inventario inicial de "${o.nombre}" para arrancar la semana.`}>
                      <button onClick={() => abrirAuditoria(o.id, "Apertura")} className="text-xs font-semibold underline hover:no-underline">
                        Hacer control de apertura ahora →
                      </button>
                    </AlertCard>
                  ))}
                  {agruparPedidosPorObra(materialesVencidos, obras).map((grupo) => (
                    <AlertCard
                      key={`venc-${grupo.obraId}`}
                      tone="rose"
                      icon={Package}
                      title={`${grupo.pedidos.length} pedido${grupo.pedidos.length > 1 ? "s" : ""} pendiente${grupo.pedidos.length > 1 ? "s" : ""} para ${grupo.nombreObra} — vencido${grupo.pedidos.length > 1 ? "s" : ""}`}
                    >
                      <div className="space-y-1">
                        {grupo.pedidos.slice(0, 4).map((p) => (
                          <button
                            key={p.id}
                            onClick={() => { setTab("materiales"); setVistaMateriales("materiales"); setObraPresupuestoId(p.obraId); }}
                            className="flex w-full items-center justify-between gap-2 rounded-md bg-white/70 px-2 py-1 text-left text-xs hover:bg-white"
                          >
                            <span className="truncate">{p.items.map((it) => it.material).slice(0, 2).join(", ")}{p.items.length > 2 ? "…" : ""}</span>
                            <span className="flex shrink-0 items-center gap-1 font-semibold"><CalendarDays size={11} /> {fmtFecha(p.fechaNecesaria)}</span>
                          </button>
                        ))}
                      </div>
                    </AlertCard>
                  ))}
                  {agruparPedidosPorObra(materialesProximos, obras).map((grupo) => (
                    <AlertCard
                      key={`prox-${grupo.obraId}`}
                      tone="amber"
                      icon={Package}
                      title={`${grupo.pedidos.length} pedido${grupo.pedidos.length > 1 ? "s" : ""} pendiente${grupo.pedidos.length > 1 ? "s" : ""} para ${grupo.nombreObra} — llega${grupo.pedidos.length > 1 ? "n" : ""} pronto`}
                    >
                      <div className="space-y-1">
                        {grupo.pedidos.slice(0, 4).map((p) => (
                          <button
                            key={p.id}
                            onClick={() => { setTab("materiales"); setVistaMateriales("materiales"); setObraPresupuestoId(p.obraId); }}
                            className="flex w-full items-center justify-between gap-2 rounded-md bg-white/70 px-2 py-1 text-left text-xs hover:bg-white"
                          >
                            <span className="truncate">{p.items.map((it) => it.material).slice(0, 2).join(", ")}{p.items.length > 2 ? "…" : ""}</span>
                            <span className="flex shrink-0 items-center gap-1 font-semibold"><CalendarDays size={11} /> {fmtFecha(p.fechaNecesaria)}</span>
                          </button>
                        ))}
                      </div>
                    </AlertCard>
                  ))}
                  {agruparPedidosPorObra(pedidosPorAprobar, obras).map((grupo) => (
                    <AlertCard
                      key={`aprobar-${grupo.obraId}`}
                      tone="sky"
                      icon={ShoppingCart}
                      title={`${grupo.pedidos.length} pedido${grupo.pedidos.length > 1 ? "s" : ""} para ${grupo.nombreObra} esperando aprobación`}
                    >
                      <div className="space-y-1">
                        {grupo.pedidos.slice(0, 4).map((p) => (
                          <button
                            key={p.id}
                            onClick={() => { setTab("materiales"); setVistaMateriales("materiales"); setObraPresupuestoId(p.obraId); }}
                            className="flex w-full items-center justify-between gap-2 rounded-md bg-white/70 px-2 py-1 text-left text-xs hover:bg-white"
                          >
                            <span className="truncate">{p.items.map((it) => it.material).slice(0, 2).join(", ")}{p.items.length > 2 ? "…" : ""}</span>
                            <span className="flex shrink-0 items-center gap-1 font-semibold">
                              {p.fechaNecesaria ? <><CalendarDays size={11} /> {fmtFecha(p.fechaNecesaria)}</> : fmtARS(p.total)}
                            </span>
                          </button>
                        ))}
                      </div>
                    </AlertCard>
                  ))}
                  {personalSinObra5Dias.length > 0 && (
                    <AlertCard tone="rose" icon={Users} title={`${personalSinObra5Dias.length} persona(s) sin obra asignada hace 5 días o más.`}>
                      <ul className="space-y-1 text-xs">
                        {personalSinObra5Dias.slice(0, 5).map((p) => (
                          <li key={p.id} className="flex items-center justify-between gap-2">
                            <span className="truncate">{nombreCompletoDe(p)} — {fmtFecha(ultimaFechaActividad(p))}</span>
                            <button onClick={() => darDeBajaPersonal(p)} className="shrink-0 rounded-md border border-rose-300 px-2 py-0.5 text-[10px] font-semibold text-rose-700 hover:bg-rose-100">Dar de baja</button>
                          </li>
                        ))}
                      </ul>
                    </AlertCard>
                  )}
                  {asistenciasEditadas.length > 0 && (
                    <AlertCard tone="sky" icon={AlertTriangle} title={`${asistenciasEditadas.length} registro(s) de asistencia modificados — revisión sugerida.`}>
                      <ul className="space-y-0.5 text-xs">
                        {asistenciasEditadas.slice(0, 5).map((a) => (
                          <li key={a.id} className="truncate">{a.nombre} ({fmtFecha(a.fecha)}) — {a.editadoPor}: "{a.motivoEdicion}"</li>
                        ))}
                      </ul>
                    </AlertCard>
                  )}
                </div>
              )}
            </Panel>

            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <Panel title="Últimos movimientos de herramientas">
                {(() => {
                  const ultimos = [...remitos].filter((r) => r.herramientaIds?.length > 0).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).slice(0, 6);
                  return ultimos.length === 0 ? (
                    <div className="text-xs text-slate-400">Sin movimientos todavía.</div>
                  ) : (
                    <div className="divide-y divide-stone-100">
                      {ultimos.map((r) => (
                        <div key={r.id} className="flex items-center justify-between py-1.5 text-xs">
                          <span className="flex items-center gap-1.5 text-slate-700">{r.origen} <ArrowRightLeft size={11} className="text-slate-400" /> {r.destino}</span>
                          <Badge estado={r.estado === "En tránsito" ? "En Obra" : "Disponible"} />
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </Panel>

              <Panel title="Compras de materiales recientes">
                {(() => {
                  const ultimos = pedidosMateriales.filter((p) => !obraIdsPapelera.has(p.obraId)).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).slice(0, 6);
                  return ultimos.length === 0 ? (
                    <div className="text-xs text-slate-400">Sin pedidos todavía.</div>
                  ) : (
                    <div className="divide-y divide-stone-100">
                      {ultimos.map((p) => (
                        <div key={p.id} className="flex items-center justify-between py-1.5 text-xs">
                          <span className="text-slate-700">{obras.find((o) => o.id === p.obraId)?.nombre} — {p.proveedor || "sin proveedor"}</span>
                          <span className="flex items-center gap-2">
                            <span className="font-mono text-slate-600">{fmtARS(p.total)}</span>
                            <Badge estado={p.estado} />
                          </span>
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </Panel>
            </div>

            <Panel title="Asignación de personal">
              <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
                {Object.values(cuadrillasPorObra).map(({ obra, empresa, gruposTantero }) => (
                  <button key={obra.id} onClick={() => { setTab("obras"); abrirObra(obra); }} className="rounded-md border border-stone-200 bg-white p-3 text-left hover:bg-stone-50">
                    <div className="truncate text-xs font-semibold text-slate-700">{obra.nombre}</div>
                    <div className="mt-1 font-mono text-lg font-bold text-slate-900">{empresa.length + gruposTantero.reduce((s, g) => s + (g.integrantes?.length || 0), 0)}</div>
                    <div className="text-[10px] text-slate-400">personas afectadas</div>
                  </button>
                ))}
                <div className="rounded-md border border-dashed border-stone-300 p-3">
                  <div className="text-xs font-semibold text-slate-500">Sin asignar</div>
                  <div className="mt-1 font-mono text-lg font-bold text-slate-700">{personalSinAsignar.length}</div>
                  <div className="text-[10px] text-slate-400">disponibles</div>
                </div>
                <div className="rounded-md border border-dashed border-stone-300 p-3">
                  <div className="text-xs font-semibold text-slate-500">Centro General</div>
                  <div className="mt-1 font-mono text-lg font-bold text-slate-700">{personalCentroGeneral.length}</div>
                  <div className="text-[10px] text-slate-400">Gerencia/RRHH/Logística</div>
                </div>
              </div>
            </Panel>

            {obras.filter((o) => o.estado === "En curso").length > 0 && (
              <Panel title="Ir a una obra">
                <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                  {obras.filter((o) => o.estado === "En curso").map((o) => (
                    <button key={o.id} onClick={() => { setTab("obras"); abrirObra(o); }} className="flex items-center justify-between rounded-md border border-stone-200 bg-white px-3 py-2 text-left text-sm hover:bg-stone-50">
                      <span className="font-medium text-slate-800">{o.nombre}</span>
                      <ArrowRightLeft size={14} className="text-slate-400" />
                    </button>
                  ))}
                </div>
              </Panel>
            )}
          </div>
        )}

        {tab === "obras" && !viewingObraId && (
          <div className="space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Obras</h2>
              <button onClick={() => setShowObraForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={16} /> Nueva obra
              </button>
            </div>

            <div className="flex flex-wrap gap-2">
              {[
                { estado: "En curso", label: "Activas", activeCls: "border-amber-500 bg-amber-500 text-slate-900", idleCls: "border-amber-300 text-amber-700 hover:bg-amber-50" },
                { estado: "Finalizada", label: "Finalizadas", activeCls: "border-emerald-600 bg-emerald-600 text-white", idleCls: "border-emerald-300 text-emerald-700 hover:bg-emerald-50" },
                { estado: "Pausada", label: "Pausadas", activeCls: "border-rose-600 bg-rose-600 text-white", idleCls: "border-rose-300 text-rose-700 hover:bg-rose-50" },
                { estado: "Papelera", label: "Papelera", activeCls: "border-slate-500 bg-slate-500 text-white", idleCls: "border-slate-300 text-slate-600 hover:bg-slate-50" },
              ].map((f) => {
                const cantidad = obras.filter((o) => o.estado === f.estado).length;
                const activo = filtroObrasEstado === f.estado;
                return (
                  <button
                    key={f.estado}
                    onClick={() => setFiltroObrasEstado(f.estado)}
                    className={`flex items-center gap-1.5 rounded-full border-2 bg-white px-3 py-1.5 text-xs font-bold uppercase tracking-wide ${activo ? f.activeCls : f.idleCls}`}
                  >
                    {f.label}
                    <span className={`flex h-5 min-w-[1.25rem] items-center justify-center rounded-full px-1 text-[10px] ${activo ? "bg-white/25" : "bg-stone-100"}`}>{cantidad}</span>
                  </button>
                );
              })}
            </div>

            {showObraForm && (
              <Panel title="Añadir obra" action={<button onClick={() => setShowObraForm(false)}><X size={16} /></button>}>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitNuevaObra}>
                  <Field label="Nombre de la obra"><input name="nombre" required className={inputCls} /></Field>
                  <Field label="Cliente">
                    {clientes.length === 0 ? (
                      <div className="rounded-md border border-dashed border-stone-300 px-2 py-2 text-xs text-slate-500">
                        No hay clientes cargados — andá primero a "Clientes/Proveedores" y dalo de alta.
                      </div>
                    ) : (
                      <select name="clienteId" required className={inputCls}>
                        <option value="">-- Elegir --</option>
                        {clientes.map((c) => <option key={c.id} value={c.id}>{c.razonSocial}</option>)}
                      </select>
                    )}
                  </Field>
                  <Field label={`Presupuesto (ARS)${resumenObraImportado ? " — se completa con el Excel" : ""}`}>
                    <MoneyInput name="presupuesto" placeholder={resumenObraImportado ? String(resumenObraImportado.precioTotalConIva) : "0,00"} className={inputCls} />
                  </Field>
                  <Field label="Fecha de inicio"><input name="inicio" type="date" required className={inputCls} /></Field>
                  <Field label="Duración (meses)"><input name="meses" type="number" min="1" required className={inputCls} /></Field>
                  <Field label="Encargado de obra">
                    <select name="encargadoId" className={inputCls}>
                      <option value="">Sin asignar</option>
                      {personal.map((p) => <option key={p.id} value={p.id}>{nombreCompletoDe(p)}</option>)}
                    </select>
                  </Field>
                  <Field label="Día de cierre semanal">
                    <select name="diaCierre" defaultValue="Viernes" className={inputCls}>
                      {DIAS_SEMANA.map((d) => <option key={d}>{d}</option>)}
                    </select>
                  </Field>
                  <Field label="Hora de cierre">
                    <input name="horaCierre" type="time" defaultValue="18:00" className={inputCls} />
                  </Field>

                  <div className="md:col-span-3 rounded-md border border-dashed border-amber-300 bg-amber-50 p-4">
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-amber-800">Presupuesto de la obra (opcional acá, se puede importar después)</div>
                    <label className="flex w-fit cursor-pointer items-center gap-1.5 rounded-md bg-white px-3 py-2 text-sm font-semibold text-slate-700 shadow-sm hover:bg-stone-50">
                      <Upload size={16} /> Subir Planilla Interna (.xlsx)
                      <input ref={obraFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleExcelUploadNuevaObra} className="hidden" />
                    </label>
                    {archivoObraNombre && (
                      <div className="mt-2 flex items-center gap-2 text-xs text-slate-600">
                        <FileSpreadsheet size={13} /> {archivoObraNombre}
                        <button type="button" onClick={quitarExcelNuevaObra} className="text-slate-400 hover:text-rose-600"><X size={13} /></button>
                      </div>
                    )}
                    {resumenObraImportado && (
                      <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-4">
                        <div className="rounded-md border border-stone-200 bg-white p-2">
                          <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Mano de Obra</div>
                          <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalManoObra)}</div>
                        </div>
                        <div className="rounded-md border border-stone-200 bg-white p-2">
                          <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Equipos</div>
                          <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalEquipos)}</div>
                        </div>
                        {resumenObraImportado.totalHerramientas > 0 && (
                          <div className="rounded-md border border-stone-200 bg-white p-2">
                            <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Herramientas</div>
                            <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalHerramientas)}</div>
                          </div>
                        )}
                        <div className="rounded-md border border-stone-200 bg-white p-2">
                          <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Materiales</div>
                          <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalMateriales)}</div>
                        </div>
                        <div className="rounded-md border border-amber-300 bg-white p-2 sm:col-span-2">
                          <div className="text-[9px] font-semibold uppercase tracking-wide text-amber-700">Total con IVA</div>
                          <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.precioTotalConIva)}</div>
                        </div>
                        <div className="text-[11px] text-slate-500 sm:col-span-2 sm:self-center">{itemsObraImportados.length} ítem(s) de Equipos/Herramientas/Materiales detectados.</div>
                      </div>
                    )}
                  </div>

                  <div className="flex items-end">
                    <button disabled={creandoObra} className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700 disabled:opacity-50">
                      {creandoObra ? "Guardando..." : "Guardar"}
                    </button>
                  </div>
                </form>
                <div className="mt-3 text-[11px] text-slate-400">
                  El encargado de obra aprueba la recepción de los remitos que le lleguen a esta obra. Gerencia siempre puede aprobar cualquier remito, sea cual sea el encargado. El día/hora de cierre dispara el aviso de auditoría semanal de herramientas.
                </div>
              </Panel>
            )}

            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              {obras.filter((o) => o.estado === filtroObrasEstado).length === 0 && (
                <div className="text-sm text-slate-400 md:col-span-2">No hay obras en este estado.</div>
              )}
              {obras.filter((o) => o.estado === filtroObrasEstado).map((o) => {
                const encargado = personal.find((p) => p.id === o.encargadoId);
                const gentePropia = personal.filter((p) => obraActualDe(p)?.id === o.id).length;
                if (o.estado === "Papelera") {
                  const horasRestantes = horasRestantesPapelera(o);
                  return (
                    <div key={o.id} className="rounded-lg border border-slate-300 bg-slate-50 p-5 shadow-sm">
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <div className="font-bold text-slate-600">{o.nombre}</div>
                          <div className="text-sm text-slate-400">{o.cliente}</div>
                        </div>
                        <Badge estado="Papelera" />
                      </div>
                      <div className="mt-3 text-xs text-slate-500">
                        Se borra definitivamente {horasRestantes > 0 ? `en ${horasRestantes}hs` : "en cualquier momento"} — sus gastos, ingresos y pedidos ya dejaron de contar en los balances, y las herramientas volvieron a Oficina.
                      </div>
                      <button onClick={() => restaurarObra(o)} className="mt-3 flex items-center gap-1.5 rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">
                        <RefreshCw size={13} /> Restaurar obra
                      </button>
                    </div>
                  );
                }
                return (
                  <div key={o.id} className="rounded-lg border border-stone-200 bg-white p-5 shadow-sm">
                    {editandoObraId === o.id ? (
                      <form onSubmit={(e) => guardarEdicionObra(e, o)} className="space-y-3">
                        <div className="text-sm font-semibold text-slate-900">Editar {o.nombre}</div>
                        <Field label="Nombre"><input name="nombre" defaultValue={o.nombre} required className={inputCls} /></Field>
                        <Field label="Cliente">
                          <select name="clienteId" defaultValue={o.clienteId || ""} className={inputCls}>
                            <option value="">Sin conectar (texto: {o.cliente || "—"})</option>
                            {clientes.map((c) => <option key={c.id} value={c.id}>{c.razonSocial}</option>)}
                          </select>
                        </Field>
                        <Field label="Presupuesto (ARS)"><MoneyInput name="presupuesto" value={o.presupuesto} className={inputCls} /></Field>
                        <Field label="Encargado de obra">
                          <select name="encargadoId" defaultValue={o.encargadoId || ""} className={inputCls}>
                            <option value="">Sin asignar</option>
                            {personal.map((p) => <option key={p.id} value={p.id}>{nombreCompletoDe(p)}</option>)}
                          </select>
                        </Field>

                        {!presupuestoGeneral.some((p) => p.obraId === o.id) && (
                          <div className="rounded-md border border-dashed border-amber-300 bg-amber-50 p-4">
                            <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-amber-800">Todavía no tiene presupuesto — importalo ahora si querés</div>
                            <label className="flex w-fit cursor-pointer items-center gap-1.5 rounded-md bg-white px-3 py-2 text-sm font-semibold text-slate-700 shadow-sm hover:bg-stone-50">
                              <Upload size={16} /> Subir Planilla Interna (.xlsx)
                              <input ref={obraFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleExcelUploadNuevaObra} className="hidden" />
                            </label>
                            {archivoObraNombre && (
                              <div className="mt-2 flex items-center gap-2 text-xs text-slate-600">
                                <FileSpreadsheet size={13} /> {archivoObraNombre}
                                <button type="button" onClick={quitarExcelNuevaObra} className="text-slate-400 hover:text-rose-600"><X size={13} /></button>
                              </div>
                            )}
                            {resumenObraImportado && (
                              <div className="mt-3 grid grid-cols-2 gap-2 sm:grid-cols-4">
                                <div className="rounded-md border border-stone-200 bg-white p-2">
                                  <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Mano de Obra</div>
                                  <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalManoObra)}</div>
                                </div>
                                <div className="rounded-md border border-stone-200 bg-white p-2">
                                  <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Equipos</div>
                                  <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalEquipos)}</div>
                                </div>
                                {resumenObraImportado.totalHerramientas > 0 && (
                                  <div className="rounded-md border border-stone-200 bg-white p-2">
                                    <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Herramientas</div>
                                    <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalHerramientas)}</div>
                                  </div>
                                )}
                                <div className="rounded-md border border-stone-200 bg-white p-2">
                                  <div className="text-[9px] font-semibold uppercase tracking-wide text-slate-500">Materiales</div>
                                  <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.totalMateriales)}</div>
                                </div>
                                <div className="rounded-md border border-amber-300 bg-white p-2 sm:col-span-2">
                                  <div className="text-[9px] font-semibold uppercase tracking-wide text-amber-700">Total con IVA</div>
                                  <div className="font-mono text-sm font-bold text-slate-900">{fmtARS(resumenObraImportado.precioTotalConIva)}</div>
                                </div>
                                <div className="text-[11px] text-slate-500 sm:col-span-2 sm:self-center">{itemsObraImportados.length} ítem(s) de Equipos/Herramientas/Materiales detectados.</div>
                              </div>
                            )}
                          </div>
                        )}

                        <div className="flex gap-2">
                          <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button>
                          <button type="button" onClick={() => { quitarExcelNuevaObra(); setEditandoObraId(null); }} className={btnGhost}>Cancelar</button>
                        </div>
                      </form>
                    ) : (
                      <>
                        <div className="flex items-start justify-between">
                          <div>
                            <button onClick={() => abrirObra(o)} className="text-left font-bold text-slate-900 underline decoration-dotted hover:text-amber-600">{o.nombre}</button>
                            <div className="text-sm text-slate-500">{o.cliente}{!o.clienteId && <span className="ml-1 text-amber-600">(sin conectar)</span>}</div>
                          </div>
                          <div className="flex items-center gap-1.5">
                            <select
                              value={o.estado}
                              onChange={(e) => cambiarEstadoObra(o, e.target.value)}
                              className={`rounded-full border-2 bg-white px-2 py-0.5 text-[10px] font-bold uppercase tracking-widest ${BADGE_STYLES[o.estado] || "border-slate-400 text-slate-500"}`}
                            >
                              {ESTADOS_OBRA.map((s) => <option key={s}>{s}</option>)}
                            </select>
                            <button onClick={() => cancelarObra(o)} title="Cancelar obra (a Papelera)" className="rounded-md border border-slate-300 p-1 text-slate-400 hover:border-rose-300 hover:text-rose-600">
                              <Trash2 size={13} />
                            </button>
                          </div>
                        </div>
                        <div className="mt-3 flex justify-between text-sm">
                          <span className="text-slate-500">Presupuesto</span>
                          <span className="font-mono font-semibold">{fmtARS(o.presupuesto)}</span>
                        </div>
                        <div className="flex justify-between text-sm">
                          <span className="text-slate-500">Duración</span>
                          <span>{o.meses} meses desde {fmtFecha(o.inicio)}</span>
                        </div>
                        <div className="flex justify-between text-sm">
                          <span className="text-slate-500">Encargado</span>
                          <span>{encargado ? nombreCompletoDe(encargado) : <span className="text-slate-400">Sin asignar</span>}</span>
                        </div>
                        {o.estado === "En curso" && (
                          <div className="flex justify-between text-sm">
                            <span className="text-slate-500">Personal afectado</span>
                            <span>{gentePropia}</span>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab === "obras" && viewingObraId && obraSel && (
          <div className="space-y-6">
            <button onClick={() => setViewingObraId(null)} className="text-xs font-semibold text-slate-500 hover:text-slate-800">
              ← Volver a Obras
            </button>

            {(() => {
              const encargadoObra = personal.find((p) => p.id === obraSel.encargadoId);
              const herrAtencionObra = herramientasAtencion.filter((h) => h.ubicacion === obraSel.nombre);
              const cierreObra = obrasEnVentanaCierre.filter((o) => o.id === obraSel.id);
              const aperturaObra = obrasSinAperturaLunes.filter((o) => o.id === obraSel.id);
              const matVencidosObra = materialesVencidos.filter((p) => p.obraId === obraSel.id);
              const matProximosObra = materialesProximos.filter((p) => p.obraId === obraSel.id);
              const pedidosAprobarObra = pedidosPorAprobar.filter((p) => p.obraId === obraSel.id);
              const ocAprobacionObra = ocPendientesAprobacion.filter((o) => o.obraId === obraSel.id);
              const asistEditadasObra = asistenciasEditadas.filter((a) => a.obraId === obraSel.id);
              const totalAlertasObra = herrAtencionObra.length + cierreObra.length + aperturaObra.length + matVencidosObra.length + matProximosObra.length + pedidosAprobarObra.length + ocAprobacionObra.length + asistEditadasObra.length + (hayDesvioAlerta ? 1 : 0);

              return (
                <>
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <div className="text-[11px] font-semibold uppercase tracking-widest text-amber-600">{obraSel.cliente}</div>
                      <h2 className="text-2xl font-bold tracking-tight text-slate-900">{obraSel.nombre}</h2>
                    </div>
                    <div className="flex items-center gap-2">
                      <Badge estado={obraSel.estado} />
                      <button onClick={() => iniciarEdicionObra(obraSel)} className={btnGhost}>
                        <span className="flex items-center gap-1"><Pencil size={13} /> Editar</span>
                      </button>
                      <button onClick={async () => { if (await cancelarObra(obraSel)) setViewingObraId(null); }} className={btnGhostDanger}>
                        <span className="flex items-center gap-1"><Trash2 size={13} /> Cancelar obra</span>
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3 text-sm sm:grid-cols-4">
                    <div className="rounded-lg border border-stone-200 bg-white p-3">
                      <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Encargado</div>
                      <div className="mt-0.5 font-medium text-slate-800">{encargadoObra ? nombreCompletoDe(encargadoObra) : "Sin asignar"}</div>
                    </div>
                    <div className="rounded-lg border border-stone-200 bg-white p-3">
                      <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Duración</div>
                      <div className="mt-0.5 font-medium text-slate-800">{obraSel.meses} meses desde {fmtFecha(obraSel.inicio)}</div>
                    </div>
                    <div className="rounded-lg border border-stone-200 bg-white p-3">
                      <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Cierre semanal</div>
                      <div className="mt-0.5 font-medium text-slate-800">{obraSel.diaCierre || "—"} {obraSel.horaCierre || ""}</div>
                    </div>
                    <div className="rounded-lg border border-stone-200 bg-white p-3">
                      <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Personal afectado</div>
                      <div className="mt-0.5 font-medium text-slate-800">{personal.filter((p) => obraActualDe(p)?.id === obraSel.id).length}</div>
                    </div>
                  </div>

                  <Panel title="Alertas de esta obra">
                    {totalAlertasObra === 0 ? (
                      <div className="flex items-center gap-2 text-sm text-emerald-700"><CheckCircle2 size={16} /> Todo en orden en esta obra.</div>
                    ) : (
                      <div className="grid grid-cols-1 gap-3 lg:grid-cols-2">
                        {hayDesvioAlerta && (
                          <AlertCard tone={desvioPct > DESVIO_DANGER_PCT ? "rose" : "amber"} icon={AlertTriangle} title={`Está ${desvioPct.toFixed(1)}% por encima de lo planificado a la fecha.`} />
                        )}
                        {ocAprobacionObra.length > 0 && (
                          <AlertCard tone="rose" icon={AlertTriangle} title={`${ocAprobacionObra.length} orden(es) de compra esperando aprobación.`} />
                        )}
                        {herrAtencionObra.length > 0 && (
                          <AlertCard tone="amber" icon={AlertTriangle} title={`${herrAtencionObra.length} herramienta(s) en mal estado o rota(s) acá.`}>
                            <ul className="space-y-0.5 text-xs">{herrAtencionObra.map((h) => <li key={h.id} className="truncate">{h.nombre} ({h.numeroSerie}) — {h.estado}</li>)}</ul>
                          </AlertCard>
                        )}
                        {cierreObra.length > 0 && (
                          <AlertCard tone="rose" icon={AlertTriangle} title="Falta menos de 1hs para el cierre — hacé el control de herramientas.">
                            <button onClick={() => abrirAuditoria(obraSel.id, "Cierre")} className="text-xs font-semibold underline hover:no-underline">Hacer control de cierre ahora →</button>
                          </AlertCard>
                        )}
                        {aperturaObra.length > 0 && (
                          <AlertCard tone="amber" icon={AlertTriangle} title="Falta validar el inventario inicial de la semana.">
                            <button onClick={() => abrirAuditoria(obraSel.id, "Apertura")} className="text-xs font-semibold underline hover:no-underline">Hacer control de apertura ahora →</button>
                          </AlertCard>
                        )}
                        {matVencidosObra.length > 0 && (
                          <AlertCard tone="rose" icon={Package} title={`${matVencidosObra.length} pedido${matVencidosObra.length > 1 ? "s" : ""} pendiente${matVencidosObra.length > 1 ? "s" : ""} para esta obra — vencido${matVencidosObra.length > 1 ? "s" : ""}`}>
                            <div className="space-y-1">
                              {matVencidosObra.map((p) => (
                                <button key={p.id} onClick={() => { setTab("materiales"); setVistaMateriales("materiales"); setObraPresupuestoId(p.obraId); }} className="flex w-full items-center justify-between gap-2 rounded-md bg-white/70 px-2 py-1 text-left text-xs hover:bg-white">
                                  <span className="truncate">{p.items.map((it) => it.material).slice(0, 2).join(", ")}{p.items.length > 2 ? "…" : ""}</span>
                                  <span className="flex shrink-0 items-center gap-1 font-semibold"><CalendarDays size={11} /> {fmtFecha(p.fechaNecesaria)}</span>
                                </button>
                              ))}
                            </div>
                          </AlertCard>
                        )}
                        {matProximosObra.length > 0 && (
                          <AlertCard tone="amber" icon={Package} title={`${matProximosObra.length} pedido${matProximosObra.length > 1 ? "s" : ""} pendiente${matProximosObra.length > 1 ? "s" : ""} para esta obra — llega${matProximosObra.length > 1 ? "n" : ""} pronto`}>
                            <div className="space-y-1">
                              {matProximosObra.map((p) => (
                                <button key={p.id} onClick={() => { setTab("materiales"); setVistaMateriales("materiales"); setObraPresupuestoId(p.obraId); }} className="flex w-full items-center justify-between gap-2 rounded-md bg-white/70 px-2 py-1 text-left text-xs hover:bg-white">
                                  <span className="truncate">{p.items.map((it) => it.material).slice(0, 2).join(", ")}{p.items.length > 2 ? "…" : ""}</span>
                                  <span className="flex shrink-0 items-center gap-1 font-semibold"><CalendarDays size={11} /> {fmtFecha(p.fechaNecesaria)}</span>
                                </button>
                              ))}
                            </div>
                          </AlertCard>
                        )}
                        {pedidosAprobarObra.length > 0 && (
                          <AlertCard tone="sky" icon={ShoppingCart} title={`${pedidosAprobarObra.length} pedido${pedidosAprobarObra.length > 1 ? "s" : ""} para esta obra esperando aprobación`}>
                            <div className="space-y-1">
                              {pedidosAprobarObra.map((p) => (
                                <button key={p.id} onClick={() => { setTab("materiales"); setVistaMateriales("materiales"); setObraPresupuestoId(p.obraId); }} className="flex w-full items-center justify-between gap-2 rounded-md bg-white/70 px-2 py-1 text-left text-xs hover:bg-white">
                                  <span className="truncate">{p.items.map((it) => it.material).slice(0, 2).join(", ")}{p.items.length > 2 ? "…" : ""}</span>
                                  <span className="flex shrink-0 items-center gap-1 font-semibold">
                                    {p.fechaNecesaria ? <><CalendarDays size={11} /> {fmtFecha(p.fechaNecesaria)}</> : fmtARS(p.total)}
                                  </span>
                                </button>
                              ))}
                            </div>
                          </AlertCard>
                        )}
                        {asistEditadasObra.length > 0 && (
                          <AlertCard tone="sky" icon={AlertTriangle} title={`${asistEditadasObra.length} registro(s) de asistencia modificados.`} />
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

                  <div className="flex flex-wrap gap-2">
                    <button onClick={() => { setTab("materiales"); setVistaMateriales("materiales"); setObraPresupuestoId(obraSel.id); }} className={btnGhost}>Ver Materiales de esta obra</button>
                    <button onClick={() => { setTab("personal"); }} className={btnGhost}>Ver Personal</button>
                  </div>
                </>
              );
            })()}
          </div>
        )}

        {tab === "personal" && !viewingPerson && (
          <div className="space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Personal/Cuadrillas</h2>
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={() => setShowCostosPanel((v) => !v)}
                  className={`flex items-center gap-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                    showCostosPanel ? "border-slate-400 bg-stone-100 text-slate-700" : "border-stone-300 bg-white text-slate-700 hover:bg-stone-50"
                  }`}
                >
                  <DollarSign size={16} /> Precios Mano de Obra
                </button>
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
                    <FileDown size={16} /> Imprimir Datos Personal
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

            {showCostosPanel && (
              <Panel title="Precios Mano de Obra" action={<button onClick={() => setShowCostosPanel(false)}><X size={16} /></button>}>
                <div className="mb-3 text-xs text-slate-500">
                  El costo por hora de cada categoría varía mes a mes. Al tomar asistencia, Liquidación usa solo el precio del mes en que se trabajó. Si un mes no tiene precio cargado, se usa el último mes anterior que sí lo tenga.
                  {!canEditarCostos && " Solo Gerente y RRHH pueden modificarlo."}
                </div>

                <div className="mb-3 flex flex-wrap gap-2">
                  {aniosCostosDisponibles.map((a) => (
                    <button
                      key={a}
                      onClick={() => setAnioCostos(a)}
                      className={`rounded-md px-3 py-1.5 text-sm font-semibold ${anioCostos === a ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
                    >
                      {a}{a === anioActual ? " (actual)" : ""}
                    </button>
                  ))}
                </div>

                <div className="overflow-x-auto rounded-lg border border-stone-200">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                      <tr>
                        <th className="sticky left-0 bg-stone-50 px-4 py-3">Categoría</th>
                        {mesesCostos.map((mes) => <th key={mes} className="px-2 py-3 text-right capitalize">{nombreMes(mes)}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {CATEGORIAS_PERSONAL.map((categoria) => (
                        <tr key={categoria} className="border-t border-stone-100">
                          <td className="sticky left-0 bg-white px-4 py-2 font-medium text-slate-900">{categoria}</td>
                          {mesesCostos.map((mes) => {
                            const entrada = costosCategoria.find((c) => c.categoria === categoria && c.mes === mes);
                            return (
                              <td key={mes} className="px-2 py-1.5">
                                {canEditarCostos ? (
                                  <MoneyInput
                                    className={`${inputCls} w-full text-right`}
                                    value={entrada?.costoHora ?? 0}
                                    onBlur={(v) => guardarCostoCelda(categoria, mes, v)}
                                  />
                                ) : entrada?.costoHora ? (
                                  <span className="block text-right">{fmtARS(entrada.costoHora)}</span>
                                ) : (
                                  <span className="block text-right text-slate-300">—</span>
                                )}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Panel>
            )}

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

            <div className="space-y-5">
              {personalCentroGeneral.length > 0 && (
                <div className="rounded-lg border border-stone-200 bg-white shadow-sm">
                  <div className="flex items-center gap-1.5 border-b border-stone-100 bg-stone-50 px-3 py-2 text-xs font-bold uppercase tracking-wide text-slate-600">
                    <Landmark size={14} className="text-amber-600" /> Centro de Costos General
                  </div>
                  <div className="divide-y divide-stone-50 p-1">
                    {personalCentroGeneral.map(renderPersonaRow)}
                  </div>
                </div>
              )}

              {Object.values(cuadrillasPorObra).map(({ obra, empresa, gruposTantero }) => (
                <div key={obra.id} className="rounded-lg border border-stone-200 bg-white shadow-sm">
                  <div className="flex items-center gap-1.5 border-b border-stone-100 bg-stone-50 px-3 py-2 text-xs font-bold uppercase tracking-wide text-slate-600">
                    <Building2 size={14} className="text-amber-600" /> {obra.nombre}
                  </div>
                  <div className="p-2">
                    {empresa.length > 0 && (
                      <div className="mb-2">
                        <div className="mb-1 px-1 text-[10px] font-semibold uppercase tracking-wide text-slate-400">Empresa</div>
                        <div className="divide-y divide-stone-50">{empresa.map(renderPersonaRow)}</div>
                      </div>
                    )}
                    {gruposTantero.map((t) => {
                      const integrantes = (t.integrantes || []).map((id) => personal.find((p) => p.id === id)).filter(Boolean);
                      return (
                        <div key={t.id} className="mb-2">
                          <div className="mb-1 px-1 text-[10px] font-semibold uppercase tracking-wide text-slate-400">Tanteros — {t.nombreGrupo}</div>
                          <div className="divide-y divide-stone-50">{integrantes.map(renderPersonaRow)}</div>
                        </div>
                      );
                    })}
                    {empresa.length === 0 && gruposTantero.length === 0 && (
                      <div className="px-2 py-2 text-xs text-slate-400">Sin personal cargado en esta obra todavía.</div>
                    )}
                  </div>
                </div>
              ))}

              <div className="rounded-lg border border-stone-200 bg-white shadow-sm">
                <div className="flex items-center gap-1.5 border-b border-stone-100 bg-stone-50 px-3 py-2 text-xs font-bold uppercase tracking-wide text-slate-600">
                  <Users size={14} className="text-slate-500" /> Sin asignar a ninguna obra
                </div>
                {personalSinAsignar.length === 0 ? (
                  <div className="px-3 py-2 text-xs text-slate-400">Todo el personal está afectado a alguna obra o al Centro General.</div>
                ) : (
                  <div className="divide-y divide-stone-50 p-1">{personalSinAsignar.map(renderPersonaRow)}</div>
                )}
                <div className="border-t border-stone-100 px-3 py-1.5 text-[11px] text-slate-400">
                  Gente que tenemos disponible por si hace falta llamarla, aunque no esté trabajando en este momento.
                </div>
              </div>
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
                      <span className="flex items-center gap-1"><FileDown size={13} /> Imprimir Datos Personal</span>
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
                      {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
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
                          {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                        </select>
                      </Field>
                      <Field label="Precio cerrado (ARS)">
                        <MoneyInput value={tanteroForm.precioTotal} onChange={(v) => setTanteroForm((f) => ({ ...f, precioTotal: v }))} className={inputCls} />
                      </Field>
                      <div className="md:col-span-3">
                        <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Integrantes del grupo</div>
                        {tanterosDisponibles.length === 0 ? (
                          <div className="rounded-md border border-dashed border-stone-300 p-3 text-xs text-slate-500">
                            Todavía no hay nadie marcado como "Tantero" en Personal. Andá a Personal, editá a la persona y elegí "Tantero" en Tipo de trabajador.
                          </div>
                        ) : (
                          <>
                            <select
                              value=""
                              onChange={(e) => { if (e.target.value) toggleIntegranteTantero(Number(e.target.value)); }}
                              className={inputCls}
                            >
                              <option value="">+ Agregar tantero...</option>
                              {tanterosDisponibles
                                .filter((p) => !tanteroForm.integrantes.includes(p.id))
                                .map((p) => <option key={p.id} value={p.id}>{nombreCompletoDe(p)}</option>)}
                            </select>
                            {tanteroForm.integrantes.length > 0 && (
                              <div className="mt-2 flex flex-wrap gap-2">
                                {tanteroForm.integrantes.map((id) => {
                                  const p = personal.find((x) => x.id === id);
                                  return (
                                    <span key={id} className="flex items-center gap-1.5 rounded-full border border-amber-300 bg-amber-50 px-3 py-1 text-xs font-medium text-amber-800">
                                      {p ? nombreCompletoDe(p) : id}
                                      <button type="button" onClick={() => toggleIntegranteTantero(id)} className="text-amber-600 hover:text-amber-900">
                                        <X size={12} />
                                      </button>
                                    </span>
                                  );
                                })}
                              </div>
                            )}
                          </>
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
                                <MoneyInput value={avanceForm.monto} onChange={(v) => setAvanceForm((f) => ({ ...f, monto: v }))} className={inputCls} />
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
                  {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                </select>

                <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                  <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                    <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Total pagado a Personal</div>
                    <div className="mt-1 font-mono text-xl font-bold text-slate-900">{fmtARS(totalHistorico)}</div>
                  </div>
                  <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                    <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Total pagado a Tanteros</div>
                    <div className="mt-1 font-mono text-xl font-bold text-slate-900">{fmtARS(totalHistoricoTanteros)}</div>
                  </div>
                </div>

                {semanasHistorialOrdenadas.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    Todavía no se registraron pagos en esta obra.
                  </div>
                ) : (
                  semanasHistorialOrdenadas.map((semanaKey) => {
                    const { personal: pagosSemana, tanteros: avancesSemana } = semanasHistorial[semanaKey];
                    const totalSemana =
                      pagosSemana.reduce((s, a) => s + (a.montoAbonado || 0), 0) + avancesSemana.reduce((s, av) => s + (av.monto || 0), 0);
                    return (
                      <Panel
                        key={semanaKey}
                        title={`Semana del ${fmtFecha(semanaKey)}`}
                        action={<span className="font-mono text-sm font-bold text-slate-800">{fmtARS(totalSemana)}</span>}
                      >
                        <div className="space-y-4">
                          {pagosSemana.length > 0 && (
                            <div>
                              <div className="mb-1.5 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Personal</div>
                              <div className="space-y-1">
                                {pagosSemana.map((a) => (
                                  <div key={a.id} className="flex items-center justify-between text-sm">
                                    <span className="text-slate-700">{a.nombre} <span className="text-xs text-slate-400">({fmtFecha(a.fecha)} · {a.horas} hs)</span></span>
                                    <span className="font-mono text-slate-800">{fmtARS(a.montoAbonado)}</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}
                          {avancesSemana.length > 0 && (
                            <div className={pagosSemana.length > 0 ? "border-t border-stone-100 pt-3" : ""}>
                              <div className="mb-1.5 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Tanteros</div>
                              <div className="space-y-1">
                                {avancesSemana.map((av) => {
                                  const t = tanteros.find((x) => x.id === av.tanteroId);
                                  return (
                                    <div key={av.id} className="flex items-center justify-between text-sm">
                                      <span className="text-slate-700">{t?.nombreGrupo || "Grupo"} <span className="text-xs text-slate-400">({fmtFecha(av.fecha)} · {av.descripcion || "avance"})</span></span>
                                      <span className="font-mono text-slate-800">{fmtARS(av.monto)}</span>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}
                        </div>
                      </Panel>
                    );
                  })
                )}
              </>
            )}
          </div>
        )}

        {tab === "herramientas" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Herramientas</h2>

            <div className="flex flex-wrap gap-2">
              <button
                onClick={() => setVistaHerramientas("altoValor")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaHerramientas === "altoValor" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Alto Valor / Maquinaria
              </button>
              <button
                onClick={() => setVistaHerramientas("combos")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaHerramientas === "combos" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Caja Herramientas Personal
              </button>
              <button
                onClick={() => setVistaHerramientas("remitos")}
                className={`flex items-center gap-1.5 rounded-md px-3 py-2 text-sm font-semibold ${vistaHerramientas === "remitos" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Remitos
                {remitosPendientes.length > 0 && (
                  <span className="flex h-5 w-5 items-center justify-center rounded-full bg-rose-500 text-[10px] font-bold text-white">{remitosPendientes.length}</span>
                )}
              </button>
              <button
                onClick={() => setVistaHerramientas("auditoria")}
                className={`flex items-center gap-1.5 rounded-md px-3 py-2 text-sm font-semibold ${vistaHerramientas === "auditoria" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Auditoría Semanal
                {(obrasEnVentanaCierre.length + obrasSinAperturaLunes.length) > 0 && (
                  <span className="flex h-5 w-5 items-center justify-center rounded-full bg-rose-500 text-[10px] font-bold text-white">{obrasEnVentanaCierre.length + obrasSinAperturaLunes.length}</span>
                )}
              </button>
            </div>

            {vistaHerramientas === "altoValor" && !viewingHerramienta && (
              <>
                <div className="flex items-center justify-between">
                  <div className="text-xs text-slate-500">Maquinaria y herramientas de alto valor, controladas de forma individual por número de serie.</div>
                  <button onClick={startAddHerramienta} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
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
                  <Panel title={editingHerramientaId ? "Editar herramienta" : "Añadir herramienta"} action={<button onClick={cancelHerrForm}><X size={16} /></button>}>
                    <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitHerrForm}>
                      <Field label="Categoría">
                        <select
                          value={herrForm.categoria}
                          onChange={(e) => setHerrForm((f) => ({ ...f, categoria: e.target.value, nombre: "" }))}
                          className={inputCls}
                        >
                          {CATEGORIAS_HERRAMIENTA.map((c) => <option key={c}>{c}</option>)}
                        </select>
                      </Field>

                      <div>
                        <Field label="Nombre de herramienta">
                          <select value={herrForm.nombre} onChange={(e) => setHerrForm((f) => ({ ...f, nombre: e.target.value }))} className={inputCls}>
                            <option value="">-- Elegir --</option>
                            {nombresDisponiblesParaCategoria.map((c) => <option key={c.id} value={c.nombre}>{c.nombre}</option>)}
                          </select>
                        </Field>
                        <button type="button" onClick={() => setShowAddNombreHerr((v) => !v)} className="mt-1 text-[11px] font-semibold text-amber-700 hover:underline">
                          + Agregar nombre de herramienta
                        </button>
                        {showAddNombreHerr && (
                          <div className="mt-1 flex gap-1">
                            <input value={nuevoNombreHerr} onChange={(e) => setNuevoNombreHerr(e.target.value)} placeholder="Nombre nuevo..." className={inputCls + " flex-1"} />
                            <button type="button" onClick={agregarNombreCatalogo} className={btnGhost}>Agregar</button>
                          </div>
                        )}
                      </div>

                      <div>
                        <Field label="Marca">
                          <select value={herrForm.marca} onChange={(e) => setHerrForm((f) => ({ ...f, marca: e.target.value }))} className={inputCls}>
                            <option value="">Sin especificar</option>
                            {catalogoMarcas.map((m) => <option key={m.id} value={m.nombre}>{m.nombre}</option>)}
                          </select>
                        </Field>
                        <button type="button" onClick={() => setShowAddMarca((v) => !v)} className="mt-1 text-[11px] font-semibold text-amber-700 hover:underline">
                          + Añadir marca
                        </button>
                        {showAddMarca && (
                          <div className="mt-1 flex gap-1">
                            <input value={nuevaMarca} onChange={(e) => setNuevaMarca(e.target.value)} placeholder="Marca nueva..." className={inputCls + " flex-1"} />
                            <button type="button" onClick={agregarMarcaCatalogo} className={btnGhost}>Agregar</button>
                          </div>
                        )}
                      </div>

                      <Field label="¿Viene con maletín?">
                        <select value={herrForm.maletin} onChange={(e) => setHerrForm((f) => ({ ...f, maletin: e.target.value }))} className={inputCls}>
                          {SI_NO.map((s) => <option key={s}>{s}</option>)}
                        </select>
                      </Field>
                      <Field label="¿Tiene accesorios?">
                        <select value={herrForm.accesorios} onChange={(e) => setHerrForm((f) => ({ ...f, accesorios: e.target.value }))} className={inputCls}>
                          {SI_NO.map((s) => <option key={s}>{s}</option>)}
                        </select>
                      </Field>
                      <Field label="Detalle de accesorios">
                        <input value={herrForm.detalleAccesorios} onChange={(e) => setHerrForm((f) => ({ ...f, detalleAccesorios: e.target.value }))} placeholder="Ej: 3 puntas SDS, 1 cincel" className={inputCls} />
                      </Field>
                      <div className="md:col-span-3">
                        <Field label="Observaciones">
                          <textarea value={herrForm.observaciones} onChange={(e) => setHerrForm((f) => ({ ...f, observaciones: e.target.value }))} rows={2} placeholder="Ej: no enciende, revisar batería..." className={inputCls} />
                        </Field>
                      </div>
                      <div className="flex items-end gap-2">
                        <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">
                          {editingHerramientaId ? "Guardar cambios" : "Guardar"}
                        </button>
                        {editingHerramientaId && <button type="button" onClick={cancelHerrForm} className={btnGhost}>Cancelar</button>}
                      </div>
                    </form>
                    {!editingHerramientaId && (
                      <div className="mt-3 text-[11px] text-slate-400">
                        Ubicación inicial: Oficina. Estado inicial: Disponible. El responsable en obra va a ser el capataz asignado (próximamente). El N° de serie se genera solo: letra del tipo + 3 letras de la marca + N° correlativo. Ej: Bosch, Herramienta Eléctrica → <span className="font-mono">E-BOS01</span>.
                      </div>
                    )}
                  </Panel>
                )}

                <div className="flex flex-wrap gap-3">
                  <select className={inputCls} value={filtroHerr.ubicacion} onChange={(e) => setFiltroHerr((f) => ({ ...f, ubicacion: e.target.value }))}>
                    <option>Todas</option>
                    <option>Oficina</option>
                    {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id}>{o.nombre}</option>)}
                  </select>
                  <select className={inputCls} value={filtroHerr.estado} onChange={(e) => setFiltroHerr((f) => ({ ...f, estado: e.target.value }))}>
                    <option>Todos</option>
                    {ESTADOS_HERRAMIENTA.map((s) => <option key={s}>{s}</option>)}
                  </select>
                </div>

                <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                  <table className="w-full text-left text-[11px]">
                    <thead className="bg-stone-50 text-[9px] font-semibold uppercase tracking-wide text-slate-500">
                      <tr>
                        <th className="px-2 py-1.5"></th>
                        <th className="px-2 py-1.5">Herramienta</th>
                        <th className="px-2 py-1.5">Ubicación</th>
                        <th className="px-2 py-1.5">Estado</th>
                      </tr>
                    </thead>
                    <tbody>
                      {herramientas
                        .filter((h) => filtroHerr.ubicacion === "Todas" || h.ubicacion === filtroHerr.ubicacion)
                        .filter((h) => filtroHerr.estado === "Todos" || h.estado === filtroHerr.estado)
                        .sort((a, b) => {
                          if (a.ubicacion === "Oficina" && b.ubicacion !== "Oficina") return 1;
                          if (a.ubicacion !== "Oficina" && b.ubicacion === "Oficina") return -1;
                          return (a.ubicacion || "").localeCompare(b.ubicacion || "");
                        })
                        .map((h) => {
                          const infoExtra = [
                            h.marca ? `Marca: ${h.marca}` : null,
                            h.numeroSerie ? `N° serie: ${h.numeroSerie}` : null,
                            h.accesorios === "Sí" && h.detalleAccesorios ? `Accesorios: ${h.detalleAccesorios}` : null,
                            h.observaciones ? `Obs: ${h.observaciones}` : null,
                          ].filter(Boolean).join(" · ");
                          return (
                            <tr key={h.id} className="border-t border-stone-100">
                              <td className="px-2 py-1"><CategoriaHerrIcon categoria={h.categoria} /></td>
                              <td className="px-2 py-1">
                                <button onClick={() => setViewingHerramientaId(h.id)} className="flex items-center gap-1 font-medium text-slate-900 underline decoration-dotted hover:text-amber-600">
                                  {h.nombre}
                                  {h.maletin === "Sí" && <span title="Viene con maletín"><Briefcase size={11} className="text-slate-400" /></span>}
                                  {infoExtra && <span title={infoExtra}><Info size={11} className="text-sky-500" /></span>}
                                </button>
                              </td>
                              <td className="px-2 py-1 text-slate-600">
                                <span className="inline-flex items-center gap-1"><MapPin size={11} className="text-amber-600" />{h.ubicacion}</span>
                              </td>
                              <td className="px-2 py-1">
                                <select
                                  value={h.estado}
                                  onChange={(e) => cambiarEstadoHerramienta(h, e.target.value)}
                                  className={`rounded-full border-2 bg-white px-1.5 py-0.5 text-[9px] font-bold uppercase tracking-widest ${BADGE_STYLES[h.estado] || "border-slate-400 text-slate-500"}`}
                                >
                                  {ESTADOS_HERRAMIENTA.map((s) => <option key={s}>{s}</option>)}
                                </select>
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              </>
            )}

            {vistaHerramientas === "altoValor" && viewingHerramienta && (
              <div className="space-y-4">
                <button onClick={() => setViewingHerramientaId(null)} className="text-xs font-semibold text-slate-500 hover:text-slate-800">
                  ← Volver a Herramientas
                </button>

                <div className="rounded-lg border border-stone-200 bg-white p-5 shadow-sm">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div className="flex items-center gap-2">
                      <span className="flex h-10 w-10 items-center justify-center rounded-full bg-stone-100 text-slate-500">
                        <CategoriaHerrIcon categoria={viewingHerramienta.categoria} size={18} />
                      </span>
                      <div>
                        <div className="text-xl font-bold text-slate-900">{viewingHerramienta.nombre}</div>
                        <div className="text-sm text-slate-500">{viewingHerramienta.categoria}</div>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button onClick={() => startEditHerramienta(viewingHerramienta)} className={btnGhost}>
                        <span className="flex items-center gap-1"><Pencil size={13} /> Editar</span>
                      </button>
                      <select
                        value={viewingHerramienta.estado}
                        onChange={(e) => cambiarEstadoHerramienta(viewingHerramienta, e.target.value)}
                        className={`rounded-full border-2 bg-white px-2 py-0.5 text-[10px] font-bold uppercase tracking-widest ${BADGE_STYLES[viewingHerramienta.estado] || "border-slate-400 text-slate-500"}`}
                      >
                        {ESTADOS_HERRAMIENTA.map((s) => <option key={s}>{s}</option>)}
                      </select>
                    </div>
                  </div>

                  <div className="mt-6 grid grid-cols-1 gap-4 text-sm sm:grid-cols-2 md:grid-cols-3">
                    <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Marca</div><div className="text-slate-800">{viewingHerramienta.marca || "—"}</div></div>
                    <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">N° de serie / ID</div><div className="font-mono text-slate-800">{viewingHerramienta.numeroSerie || "—"}</div></div>
                    <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Ubicación</div><div className="text-slate-800">{viewingHerramienta.ubicacion}</div></div>
                    <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">¿Viene con maletín?</div><div className="text-slate-800">{viewingHerramienta.maletin || "No"}</div></div>
                    <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">¿Tiene accesorios?</div><div className="text-slate-800">{viewingHerramienta.accesorios || "No"}</div></div>
                    {viewingHerramienta.accesorios === "Sí" && (
                      <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Detalle de accesorios</div><div className="text-slate-800">{viewingHerramienta.detalleAccesorios || "—"}</div></div>
                    )}
                  </div>

                  {viewingHerramienta.observaciones && (
                    <div className="mt-4 rounded-md border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800">
                      <div className="mb-1 flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide"><AlertTriangle size={13} /> Observaciones</div>
                      {viewingHerramienta.observaciones}
                    </div>
                  )}

                  {(() => {
                    const historial = remitos
                      .filter((r) => r.herramientaIds.includes(viewingHerramienta.id))
                      .sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));
                    return (
                      <div className="mt-5 border-t border-stone-100 pt-4">
                        <div className="mb-2 flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                          <ArrowRightLeft size={13} /> Historial de movimientos y reparaciones
                        </div>
                        {historial.length === 0 ? (
                          <div className="text-xs text-slate-400">Todavía no tiene remitos registrados.</div>
                        ) : (
                          <div className="space-y-1.5">
                            {historial.map((r) => (
                              <div key={r.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md bg-stone-50 px-3 py-1.5 text-xs">
                                <span className="flex items-center gap-1.5 text-slate-700">
                                  {r.origen} <ArrowRightLeft size={11} className="text-slate-400" /> {r.destino}
                                  {r.destinoEsTaller && <span title="Reparación"><Wrench size={11} className="text-orange-600" /></span>}
                                </span>
                                <span className="flex items-center gap-2 text-slate-400">
                                  {fmtFecha(r.fecha)}
                                  <Badge estado={r.estado === "En tránsito" ? "En Obra" : "Disponible"} />
                                </span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>
              </div>
            )}

            {vistaHerramientas === "combos" && (
              <>
                <div className="flex items-center justify-between">
                  <div className="text-xs text-slate-500">Cajas de herramientas manuales chicas, armadas por rubro. Primero se asignan a una obra y después a un operario de esa obra.</div>
                  <button onClick={() => setShowComboForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                    <Plus size={16} /> Nueva caja
                  </button>
                </div>

                {showComboForm && (
                  <Panel title="Armar caja de herramientas" action={<button onClick={() => setShowComboForm(false)}><X size={16} /></button>}>
                    <form className="space-y-4" onSubmit={submitComboForm}>
                      <Field label="Rubro de la caja">
                        <select value={comboForm.tipo} onChange={(e) => setComboForm((f) => ({ ...f, tipo: e.target.value }))} className={inputCls}>
                          {TIPOS_CAJA.map((t) => <option key={t}>{t}</option>)}
                        </select>
                      </Field>
                      <div className="text-[11px] text-slate-400">
                        Se va a llamar <span className="font-semibold text-slate-600">"Caja {comboForm.tipo} {generarNumeroCaja(comboForm.tipo)}"</span>. Todavía no tiene a nadie asignado — eso se hace después de guardarla.
                      </div>

                      <div>
                        <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Herramientas de la caja</div>
                        <div className="flex flex-wrap items-end gap-2">
                          <div className="flex-1 min-w-[160px]">
                            <select
                              value={comboItemDraft.nombre}
                              onChange={(e) => setComboItemDraft((d) => ({ ...d, nombre: e.target.value }))}
                              className={inputCls + " w-full"}
                            >
                              <option value="">-- Elegir --</option>
                              {catalogoChicas.map((c) => <option key={c.id} value={c.nombre}>{c.nombre}</option>)}
                            </select>
                          </div>
                          <input
                            type="number"
                            min="1"
                            value={comboItemDraft.cantidad}
                            onChange={(e) => setComboItemDraft((d) => ({ ...d, cantidad: e.target.value }))}
                            className={inputCls + " w-20"}
                          />
                          <button type="button" onClick={agregarItemCombo} className={btnGhost}>+ Agregar</button>
                        </div>
                        <button type="button" onClick={() => setShowAddChica((v) => !v)} className="mt-1 text-[11px] font-semibold text-amber-700 hover:underline">
                          + Agregar nombre de herramienta
                        </button>
                        {showAddChica && (
                          <div className="mt-1 flex gap-1">
                            <input value={nuevaChica} onChange={(e) => setNuevaChica(e.target.value)} placeholder="Ej: Serrucho" className={inputCls + " flex-1"} />
                            <button type="button" onClick={agregarChicaCatalogo} className={btnGhost}>Agregar</button>
                          </div>
                        )}
                        {comboForm.items.length > 0 && (
                          <div className="mt-2 flex flex-wrap gap-2">
                            {comboForm.items.map((it, idx) => (
                              <span key={idx} className="flex items-center gap-1.5 rounded-full border border-amber-300 bg-amber-50 px-3 py-1 text-xs font-medium text-amber-800">
                                {it.nombre} {it.cantidad > 1 ? `x${it.cantidad}` : ""}
                                <button type="button" onClick={() => quitarItemComboDraft(idx)} className="text-amber-600 hover:text-amber-900"><X size={12} /></button>
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar caja</button>
                    </form>
                  </Panel>
                )}

                {combosHerramientas.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    Todavía no hay cajas armadas.
                  </div>
                ) : (
                  <div className="space-y-4">
                    {combosHerramientas.map((combo) => {
                      const persona = personal.find((p) => p.id === combo.personaId);
                      const obraCaja = obras.find((o) => o.id === combo.obraId);
                      return (
                        <div key={combo.id} className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <div>
                              <span className="font-semibold text-slate-900">{nombreCaja(combo)}</span>
                              <span className="ml-2"><Badge estado={obraCaja ? "En Obra" : "Disponible"} /></span>
                            </div>
                            <span className="text-xs text-slate-400">Armada el {fmtFecha(combo.fecha)}</span>
                          </div>

                          <div className="mb-3 flex flex-wrap items-center gap-2 text-sm">
                            {!obraCaja ? (
                              asignandoObraCajaId === combo.id ? (
                                <>
                                  <select value={obraParaAsignar} onChange={(e) => setObraParaAsignar(e.target.value)} className={inputCls}>
                                    <option value="">-- Elegir obra --</option>
                                    {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                                  </select>
                                  <button onClick={() => confirmarAsignarObraCaja(combo)} className={btnGhost}>Confirmar</button>
                                  <button onClick={() => { setAsignandoObraCajaId(null); setObraParaAsignar(""); }} className="text-xs text-slate-400 hover:underline">Cancelar</button>
                                </>
                              ) : (
                                <>
                                  <span className="text-slate-400">Sin asignar — en depósito/oficina</span>
                                  <button onClick={() => setAsignandoObraCajaId(combo.id)} className={btnGhost}>Asignar a obra</button>
                                </>
                              )
                            ) : !persona ? (
                              asignandoCajaId === combo.id ? (
                                <>
                                  <span className="text-slate-600">{obraCaja.nombre} →</span>
                                  <select value={personaParaAsignar} onChange={(e) => setPersonaParaAsignar(e.target.value)} className={inputCls}>
                                    <option value="">-- Elegir operario --</option>
                                    {personal.map((p) => <option key={p.id} value={p.id}>{nombreCompletoDe(p)}</option>)}
                                  </select>
                                  <button onClick={() => confirmarAsignarCaja(combo)} className={btnGhost}>Confirmar</button>
                                  <button onClick={() => { setAsignandoCajaId(null); setPersonaParaAsignar(""); }} className="text-xs text-slate-400 hover:underline">Cancelar</button>
                                </>
                              ) : (
                                <>
                                  <span className="text-slate-600">En <span className="font-medium text-slate-900">{obraCaja.nombre}</span>, sin operario asignado</span>
                                  <button onClick={() => setAsignandoCajaId(combo.id)} className={btnGhost}>Asignar a operario</button>
                                  <button onClick={() => devolverCaja(combo)} className={btnGhostDanger}>Devolver a depósito</button>
                                </>
                              )
                            ) : (
                              <>
                                <span className="text-slate-600">{obraCaja.nombre} — Asignada a <span className="font-medium text-slate-900">{nombreCompletoDe(persona)}</span></span>
                                <button onClick={() => devolverCaja(combo)} className={btnGhost}>Devolver a depósito</button>
                              </>
                            )}
                          </div>

                          <div className="divide-y divide-stone-50 border-t border-stone-100 pt-1">
                            {combo.items.map((it, idx) => (
                              <div key={idx} className="flex flex-wrap items-center justify-between gap-2 py-1.5 text-sm">
                                <span className="text-slate-700">{it.nombre} {it.cantidad > 1 ? `x${it.cantidad}` : ""}</span>
                                <div className="flex items-center gap-2">
                                  <Badge estado={it.estado} />
                                  <select
                                    value={it.estado}
                                    onChange={(e) => actualizarItemCombo(combo, idx, e.target.value)}
                                    className="rounded-md border border-stone-300 bg-white px-1.5 py-1 text-xs text-slate-600"
                                  >
                                    {ESTADOS_ITEM_COMBO.map((s) => <option key={s}>{s}</option>)}
                                  </select>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
                <div className="text-[11px] text-slate-400">
                  "Roto" = requiere devolución física para reposición. "Perdido" = se gestiona descuento o reposición con la persona. Las cajas se asignan/devuelven directo por ahora; el traslado formal con remito es para las herramientas de Alto Valor, en la pestaña "Remitos".
                </div>
              </>
            )}

            {vistaHerramientas === "remitos" && (
              <>
                <div className="flex items-center justify-between">
                  <div className="text-xs text-slate-500">Traslados de herramientas de Alto Valor entre Oficina, obras y talleres — con aprobación de salida y de recepción.</div>
                  <button onClick={() => setShowRemitoForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                    <Plus size={16} /> Nuevo remito
                  </button>
                </div>

                {showRemitoForm && (
                  <Panel title="Generar remito de salida" action={<button onClick={() => setShowRemitoForm(false)}><X size={16} /></button>}>
                    <form className="space-y-4" onSubmit={submitRemitoForm}>
                      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                        <Field label="Origen">
                          <select
                            value={remitoForm.origen}
                            onChange={(e) => setRemitoForm((f) => ({ ...f, origen: e.target.value, herramientaIds: [] }))}
                            className={inputCls}
                          >
                            <option>Oficina</option>
                            {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id}>{o.nombre}</option>)}
                          </select>
                        </Field>
                        <Field label="Destino">
                          <select value={remitoForm.destino} onChange={(e) => setRemitoForm((f) => ({ ...f, destino: e.target.value }))} className={inputCls}>
                            <option value="">-- Elegir --</option>
                            <option>Oficina</option>
                            <optgroup label="Obras">
                              {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id}>{o.nombre}</option>)}
                            </optgroup>
                            {talleres.length > 0 && (
                              <optgroup label="Talleres de reparación">
                                {talleres.map((t) => <option key={t.id}>{t.razonSocial}</option>)}
                              </optgroup>
                            )}
                          </select>
                        </Field>
                      </div>

                      <div>
                        <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Herramientas en "{remitoForm.origen}" para enviar</div>
                        {herramientasEnOrigenRemito.length === 0 ? (
                          <div className="rounded-md border border-dashed border-stone-300 p-3 text-xs text-slate-500">No hay herramientas registradas en esa ubicación.</div>
                        ) : (
                          <div className="max-h-52 space-y-1 overflow-y-auto rounded-md border border-stone-200 p-2">
                            {herramientasEnOrigenRemito.map((h) => (
                              <label key={h.id} className="flex items-center gap-2 rounded px-1 py-1 text-sm hover:bg-stone-50">
                                <input type="checkbox" checked={remitoForm.herramientaIds.includes(h.id)} onChange={() => toggleHerramientaRemito(h.id)} className="h-3.5 w-3.5" />
                                <CategoriaHerrIcon categoria={h.categoria} />
                                {h.nombre} <span className="font-mono text-xs text-slate-400">({h.numeroSerie || "s/n"})</span>
                              </label>
                            ))}
                          </div>
                        )}
                      </div>

                      <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Confirmar salida</button>
                    </form>
                    <div className="mt-3 text-[11px] text-slate-400">
                      Al confirmar, el remito queda "En tránsito". La herramienta cambia de ubicación recién cuando el destino confirma la recepción.
                    </div>
                  </Panel>
                )}

                <div>
                  <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Pendientes de recepción</div>
                  {remitosPendientes.length === 0 ? (
                    <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-6 text-center text-sm text-slate-500">No hay remitos en tránsito.</div>
                  ) : (
                    <div className="space-y-3">
                      {remitosPendientes.map((r) => (
                        <div key={r.id} className="rounded-lg border border-amber-300 bg-amber-50 p-4">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <span className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                              {r.origen} <ArrowRightLeft size={14} className="text-amber-600" /> {r.destino}
                            </span>
                            <span className="text-xs text-slate-500">Salió el {fmtFecha(r.fecha)} ({r.creadoPor})</span>
                          </div>
                          <div className="mt-1 text-xs text-slate-600">
                            {r.herramientaIds.map((id) => herramientas.find((h) => h.id === id)?.nombre).filter(Boolean).join(", ")}
                          </div>
                          <button onClick={() => confirmarRecepcionRemito(r)} className="mt-2 flex items-center gap-1 rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">
                            <Check size={13} /> Confirmar recepción en {r.destino}
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                {remitosCompletados.length > 0 && (
                  <div>
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Historial de remitos recibidos</div>
                    <div className="space-y-2">
                      {[...remitosCompletados].sort((a, b) => fechaLocal(b.fechaRecepcion) - fechaLocal(a.fechaRecepcion)).map((r) => (
                        <div key={r.id} className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-stone-200 bg-white px-4 py-2 text-sm">
                          <span className="flex items-center gap-2 text-slate-700">
                            {r.origen} <ArrowRightLeft size={13} className="text-slate-400" /> {r.destino}
                          </span>
                          <span className="text-xs text-slate-400">Recibido el {fmtFecha(r.fechaRecepcion)} ({r.recibidoPor})</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </>
            )}

            {vistaHerramientas === "auditoria" && (
              <>
                <div className="text-xs text-slate-500">
                  Control semanal de herramientas por obra: cierre (última hora del día de cierre) y apertura (lunes, al arrancar).
                </div>

                {(obrasEnVentanaCierre.length > 0 || obrasSinAperturaLunes.length > 0) && (
                  <div className="space-y-2">
                    {obrasEnVentanaCierre.map((o) => (
                      <div key={`c-${o.id}`} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-rose-300 bg-rose-50 px-4 py-3">
                        <span className="text-sm text-rose-800">Cierre de "{o.nombre}" en menos de 1hs ({o.horaCierre}).</span>
                        <button onClick={() => abrirAuditoria(o.id, "Cierre")} className="rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">Hacer control</button>
                      </div>
                    ))}
                    {obrasSinAperturaLunes.map((o) => (
                      <div key={`a-${o.id}`} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-amber-300 bg-amber-50 px-4 py-3">
                        <span className="text-sm text-amber-800">Falta la apertura de semana de "{o.nombre}".</span>
                        <button onClick={() => abrirAuditoria(o.id, "Apertura")} className="rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">Hacer control</button>
                      </div>
                    ))}
                  </div>
                )}

                <div className="flex items-center justify-end">
                  <button
                    onClick={() => { setShowAuditoriaForm((v) => !v); setPresentesAuditoria([]); setObsAuditoria(""); }}
                    className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400"
                  >
                    <ClipboardCheck size={16} /> Hacer un control manual
                  </button>
                </div>

                {showAuditoriaForm && (
                  <Panel title={`Control de ${tipoAuditoria.toLowerCase()} de herramientas`} action={<button onClick={() => setShowAuditoriaForm(false)}><X size={16} /></button>}>
                    <form className="space-y-4" onSubmit={submitAuditoria}>
                      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                        <Field label="Obra">
                          <select value={obraAuditoriaId} onChange={(e) => { setObraAuditoriaId(Number(e.target.value)); setPresentesAuditoria([]); }} className={inputCls}>
                            {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                          </select>
                        </Field>
                        <Field label="Tipo de control">
                          <select value={tipoAuditoria} onChange={(e) => setTipoAuditoria(e.target.value)} className={inputCls}>
                            <option>Cierre</option>
                            <option>Apertura</option>
                          </select>
                        </Field>
                      </div>

                      <div>
                        <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                          Herramientas registradas en "{obraAuditoriaSel?.nombre}" — tildá las que están físicamente presentes
                        </div>
                        {herramientasDeObraAuditoria.length === 0 ? (
                          <div className="rounded-md border border-dashed border-stone-300 p-3 text-xs text-slate-500">No hay herramientas de Alto Valor registradas en esta obra.</div>
                        ) : (
                          <div className="space-y-1 rounded-md border border-stone-200 p-2">
                            {herramientasDeObraAuditoria.map((h) => (
                              <label key={h.id} className="flex items-center gap-2 rounded px-1 py-1 text-sm hover:bg-stone-50">
                                <input type="checkbox" checked={presentesAuditoria.includes(h.id)} onChange={() => togglePresenteAuditoria(h.id)} className="h-3.5 w-3.5" />
                                <CategoriaHerrIcon categoria={h.categoria} />
                                {h.nombre} <span className="font-mono text-xs text-slate-400">({h.numeroSerie || "s/n"})</span>
                              </label>
                            ))}
                          </div>
                        )}
                      </div>

                      <Field label="Observaciones">
                        <textarea value={obsAuditoria} onChange={(e) => setObsAuditoria(e.target.value)} rows={2} placeholder="Ej: falta el rotomartillo, avisado al encargado..." className={inputCls} />
                      </Field>

                      <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar control</button>
                    </form>
                  </Panel>
                )}

                {auditorias.length > 0 && (
                  <div>
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Historial de controles</div>
                    <div className="space-y-2">
                      {[...auditorias].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).slice(0, 15).map((a) => {
                        const obra = obras.find((o) => o.id === a.obraId);
                        return (
                          <div key={a.id} className="rounded-lg border border-stone-200 bg-white px-4 py-2.5 text-sm">
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <span className="font-medium text-slate-800">{obra?.nombre} — {a.tipo}</span>
                              <span className="text-xs text-slate-400">{fmtFecha(a.fecha)} ({a.realizadoPor})</span>
                            </div>
                            {a.herramientasFaltantes?.length > 0 && (
                              <div className="mt-1 text-xs text-rose-600">
                                Faltantes: {a.herramientasFaltantes.map((id) => herramientas.find((h) => h.id === id)?.nombre).filter(Boolean).join(", ")}
                              </div>
                            )}
                            {a.observaciones && <div className="mt-1 text-xs text-slate-500">{a.observaciones}</div>}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </>
            )}
          </div>
        )}

        {tab === "materiales" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Pedidos de Obra</h2>

            <div className="flex flex-wrap gap-2">
              <button
                onClick={() => { setVistaMateriales("materiales"); setItemManualDraft((d) => ({ ...d, categoria: "Materiales", subcategoria: "", tipo: "" })); }}
                className={`flex items-center gap-1.5 rounded-md px-3 py-2 text-sm font-semibold ${vistaMateriales === "materiales" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Materiales
                {(materialesVencidos.length + materialesProximos.length) > 0 && (
                  <span className="flex h-5 w-5 items-center justify-center rounded-full bg-rose-500 text-[10px] font-bold text-white">{materialesVencidos.length + materialesProximos.length}</span>
                )}
              </button>
              <button
                onClick={() => { setVistaMateriales("equipos"); setItemManualDraft((d) => ({ ...d, categoria: "Equipos", subcategoria: "", tipo: "" })); }}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaMateriales === "equipos" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Equipos y Herramientas
              </button>
              <button
                onClick={() => { setVistaMateriales("epps"); setItemManualDraft((d) => ({ ...d, categoria: "Epps", subcategoria: "", tipo: "" })); }}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaMateriales === "epps" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Epps
              </button>
              <button
                onClick={() => { setVistaMateriales("consumibles"); setItemManualDraft((d) => ({ ...d, categoria: "Consumibles", subcategoria: "", tipo: "" })); }}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaMateriales === "consumibles" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Consumibles
              </button>
              {canVerPreciosPedido && (
                <>
                  <button
                    onClick={() => setVistaMateriales("catalogo")}
                    className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaMateriales === "catalogo" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
                  >
                    Catálogo y precios
                  </button>
                  <button
                    onClick={() => setVistaMateriales("consolidar")}
                    className={`flex items-center gap-1.5 rounded-md px-3 py-2 text-sm font-semibold ${vistaMateriales === "consolidar" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
                  >
                    Consolidar Pedidos
                    {(pedidosSinEnviar.length + remitosMaterialesPendientes.length) > 0 && (
                      <span className="flex h-5 w-5 items-center justify-center rounded-full bg-rose-500 text-[10px] font-bold text-white">{pedidosSinEnviar.length + remitosMaterialesPendientes.length}</span>
                    )}
                  </button>
                  <button
                    onClick={() => setVistaMateriales("stock")}
                    className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaMateriales === "stock" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
                  >
                    Stock
                  </button>
                </>
              )}
            </div>

            {["materiales", "equipos", "epps", "consumibles"].includes(vistaMateriales) && (
              <>
                <div className="flex flex-wrap items-end gap-3">
                  <div className="min-w-[200px]">
                    <Field label="Obra">
                      <select value={obraPresupuestoId} onChange={(e) => setObraPresupuestoId(e.target.value)} className={inputCls}>
                        {["epps", "consumibles"].includes(vistaMateriales) && (
                          <option value="general">Compra general (sin obra — va a depósito)</option>
                        )}
                        {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                      </select>
                    </Field>
                  </div>
                </div>

                {vistaMateriales === "materiales" && !presupuestoGeneral.some((p) => p.obraId === Number(obraPresupuestoId)) && (
                  <div className="rounded-md border border-dashed border-amber-300 bg-amber-50 px-4 py-2 text-xs text-amber-800">
                    Esta obra todavía no tiene presupuesto cargado. Se importa desde "Obras" → editar la obra, con la Planilla Interna para Costeo.
                  </div>
                )}

                {["epps", "consumibles"].includes(vistaMateriales) && obraPresupuestoId === "general" && (
                  <div className="rounded-md border border-dashed border-sky-300 bg-sky-50 px-4 py-2 text-xs text-sky-800">
                    Esto es para comprar en volumen para toda la empresa. Al recibirse, queda en el depósito (pestaña "Stock") — cada obra lo va a ir pidiendo de ahí a medida que lo necesite, y el gasto recién se imputa a su centro de costos en ese momento.
                  </div>
                )}

                {(() => {
                  const esGeneral = ["epps", "consumibles"].includes(vistaMateriales) && obraPresupuestoId === "general";
                  const obraIdActual = esGeneral ? null : Number(obraPresupuestoId);
                  const categoriasActivas = CATEGORIAS_POR_VISTA[vistaMateriales] || [];
                  const lineasObra = presupuestoMateriales.filter((m) => m.obraId === obraIdActual && categoriasActivas.includes(m.categoria));
                  const totalObra = lineasObra.reduce((s, m) => s + (m.total || 0), 0);
                  const pedidosObra = pedidosMateriales.filter((p) => p.obraId === obraIdActual);
                  const pg = presupuestoGeneral.find((p) => p.obraId === obraIdActual);
                  const gastoRealObra = comprasFacturas.filter((c) => c.obraId === obraIdActual).reduce((s, c) => s + (c.monto || 0), 0);
                  const stockEppsOConsumibles = stockMateriales.filter((s) => s.cantidad > 0 && categoriasActivas.includes(s.categoria));
                  const herramientasOrdenadas = [...herramientas].sort((a, b) => {
                    const libreA = a.estado === "Disponible" ? 0 : 1;
                    const libreB = b.estado === "Disponible" ? 0 : 1;
                    return libreA - libreB || a.nombre.localeCompare(b.nombre);
                  });
                  const catalogoEppsOConsumibles = catalogoMateriales
                    .filter((m) => categoriasActivas.includes(m.categoria))
                    .filter((m) => vistaMateriales !== "epps" || filtroEppParte === "Todas" || (m.subcategoria || "Sin parte asignada") === filtroEppParte)
                    .filter((m) => vistaMateriales !== "epps" || filtroEppTipo === "Todos" || (m.tipo || "Sin tipo") === filtroEppTipo)
                    .sort((a, b) => a.nombre.localeCompare(b.nombre));
                  const partesCuerpoEpp = subcategoriasMat.filter((s) => s.categoria === "Epps").map((s) => s.nombre);
                  const tiposEppDeLaParte = filtroEppParte === "Todas" ? [] : tiposMaterial.filter((t) => t.categoria === "Epps" && t.subcategoria === filtroEppParte).map((t) => t.nombre);
                  return (
                    <>
                      {showPedidoForm && (
                        <Panel title="Armar pedido" action={<button onClick={() => setShowPedidoForm(false)}><X size={16} /></button>}>
                          {canVerPreciosPedido ? (
                            <div className="mb-4 grid grid-cols-1 gap-4 sm:grid-cols-2 sm:max-w-md">
                              <div>
                                <Field label="Proveedor">
                                  <select value={pedidoProveedor} onChange={(e) => setPedidoProveedor(e.target.value)} className={inputCls}>
                                    <option value="">Sin especificar</option>
                                    {pedidoProveedor && !proveedores.some((p) => p.razonSocial === pedidoProveedor) && (
                                      <option value={pedidoProveedor}>{pedidoProveedor} (sugerido)</option>
                                    )}
                                    {proveedores.filter((p) => p.esTaller !== "Sí").map((p) => <option key={p.id} value={p.razonSocial}>{p.razonSocial}</option>)}
                                  </select>
                                </Field>
                                {pedidoProveedor && (
                                  <div className="mt-1 text-[11px] text-slate-400">Sugerido según la última compra — cambialo si conseguiste mejor precio.</div>
                                )}
                              </div>
                              <Field label="¿Cuándo lo necesitás?">
                                <input type="date" required value={pedidoFechaNecesaria} onChange={(e) => setPedidoFechaNecesaria(e.target.value)} className={inputCls} />
                              </Field>
                            </div>
                          ) : (
                            <div className="mb-4 max-w-xs">
                              <Field label="¿Cuándo lo necesitás?">
                                <input type="date" required value={pedidoFechaNecesaria} onChange={(e) => setPedidoFechaNecesaria(e.target.value)} className={inputCls} />
                              </Field>
                              <div className="mt-1 text-[11px] text-slate-400">El proveedor y los precios los define Logística cuando el pedido esté aprobado.</div>
                            </div>
                          )}

                          <div className="overflow-x-auto rounded-md border border-stone-200">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-stone-50 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
                                <tr>
                                  <th className="px-2 py-2">Material</th><th className="px-2 py-2">Unidad</th>
                                  <th className="px-2 py-2 text-right">Cantidad</th>
                                  {canVerPreciosPedido && (<><th className="px-2 py-2 text-right">P. Unitario</th><th className="px-2 py-2 text-right">Total</th></>)}
                                  <th className="px-2 py-2"></th>
                                </tr>
                              </thead>
                              <tbody>
                                {pedidoItems.map((it, idx) => (
                                  <tr key={idx} className="border-t border-stone-100">
                                    <td className="px-2 py-1.5 font-medium text-slate-800">
                                      {it.material}
                                      {!it.presupuestoId && <span className="ml-1 rounded bg-sky-100 px-1 text-[9px] font-semibold text-sky-700">MANUAL</span>}
                                    </td>
                                    <td className="px-2 py-1.5">{it.unidad}</td>
                                    <td className="px-2 py-1.5 text-right">
                                      <input type="number" value={it.cantidad} onChange={(e) => actualizarCantidadPedido(idx, "cantidad", Number(e.target.value))} className="w-20 rounded border border-stone-300 px-1.5 py-1 text-right" />
                                    </td>
                                    {canVerPreciosPedido && (
                                      <>
                                        <td className="px-2 py-1.5 text-right">
                                          <MoneyInput value={it.precioUnitario} onChange={(v) => actualizarCantidadPedido(idx, "precioUnitario", v)} className="w-24 rounded border border-stone-300 px-1.5 py-1 text-right" />
                                        </td>
                                        <td className="px-2 py-1.5 text-right font-mono">{fmtARS((Number(it.cantidad) || 0) * (Number(it.precioUnitario) || 0))}</td>
                                      </>
                                    )}
                                    <td className="px-2 py-1.5"><button onClick={() => quitarItemPedido(idx)} className="text-slate-400 hover:text-rose-600"><X size={13} /></button></td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>

                          <div className="mt-3 flex items-end gap-2 rounded-md border border-dashed border-stone-300 p-3">
                            {canVerPreciosPedido && (
                              <div className="w-28">
                                <Field label="Categoría">
                                  <select value={itemManualDraft.categoria} onChange={(e) => setItemManualDraft((d) => ({ ...d, categoria: e.target.value, subcategoria: "", tipo: "" }))} className={inputCls}>
                                    {CATEGORIAS_PEDIDO.map((c) => <option key={c}>{c}</option>)}
                                  </select>
                                </Field>
                              </div>
                            )}
                            {canVerPreciosPedido && itemManualDraft.categoria !== "Consumibles" && (
                              <>
                                <div className="w-32">
                                  <Field label={itemManualDraft.categoria === "Epps" ? "Parte del cuerpo" : "Sub-categoría"}>
                                    <select value={itemManualDraft.subcategoria} onChange={(e) => setItemManualDraft((d) => ({ ...d, subcategoria: e.target.value, tipo: "" }))} className={inputCls}>
                                      <option value="">--</option>
                                      {subcategoriasMat.filter((s) => s.categoria === itemManualDraft.categoria).map((s) => <option key={s.id} value={s.nombre}>{s.nombre}</option>)}
                                    </select>
                                  </Field>
                                </div>
                                <div className="w-32">
                                  <Field label={itemManualDraft.categoria === "Epps" ? "Tipo específico" : "Tipo"}>
                                    <select value={itemManualDraft.tipo} onChange={(e) => setItemManualDraft((d) => ({ ...d, tipo: e.target.value }))} className={inputCls}>
                                      <option value="">--</option>
                                      {tiposMaterial.filter((t) => t.categoria === itemManualDraft.categoria && t.subcategoria === itemManualDraft.subcategoria).map((t) => <option key={t.id} value={t.nombre}>{t.nombre}</option>)}
                                    </select>
                                  </Field>
                                </div>
                              </>
                            )}
                            <div className="flex-1 min-w-[140px]">
                              <Field label="¿Qué querés pedir?">
                                <input value={itemManualDraft.material} onChange={(e) => setItemManualDraft((d) => ({ ...d, material: e.target.value }))} placeholder="Nombre..." className={inputCls} />
                              </Field>
                            </div>
                            <div className="w-20">
                              <Field label="Unidad">
                                <input value={itemManualDraft.unidad} onChange={(e) => setItemManualDraft((d) => ({ ...d, unidad: e.target.value }))} className={inputCls} />
                              </Field>
                            </div>
                            <div className="w-20">
                              <Field label="Cant.">
                                <input type="number" value={itemManualDraft.cantidad} onChange={(e) => setItemManualDraft((d) => ({ ...d, cantidad: Number(e.target.value) }))} className={inputCls} />
                              </Field>
                            </div>
                            {canVerPreciosPedido && (
                              <div className="w-28">
                                <Field label="P. Unitario">
                                  <MoneyInput value={itemManualDraft.precioUnitario} onChange={(v) => setItemManualDraft((d) => ({ ...d, precioUnitario: v }))} className={inputCls} />
                                </Field>
                              </div>
                            )}
                            <button type="button" onClick={agregarItemManualPedido} className={btnGhost}>+ Agregar</button>
                          </div>
                          <div className="mt-1 text-[11px] text-slate-400">
                            {canVerPreciosPedido
                              ? "Esto es para algo que no estaba en el presupuesto original — se agrega igual, a mano. No hace falta el precio ahora: lo carga Logística/Contabilidad cuando llega la compra."
                              : "Esto es para algo que no estaba en el presupuesto original — no hace falta precio ni categoría, eso lo completa Logística."}
                          </div>

                          <div className="mt-4 flex items-center justify-between border-t border-stone-100 pt-3">
                            {canVerPreciosPedido ? (
                              <div className="font-mono text-lg font-bold text-slate-900">
                                Total: {fmtARS(pedidoItems.reduce((s, it) => s + (Number(it.cantidad) || 0) * (Number(it.precioUnitario) || 0), 0))}
                              </div>
                            ) : <div />}
                            <button disabled={enviandoPedido} onClick={confirmarPedido} className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700 disabled:opacity-50">
                              {enviandoPedido ? "Guardando..." : "Confirmar pedido"}
                            </button>
                          </div>
                        </Panel>
                      )}

                      {pg && vistaMateriales === "materiales" && canVerPreciosPedido && (
                        <Panel title="Presupuestado vs. Real">
                          <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
                            <div>
                              <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Mano de Obra (pres.)</div>
                              <div className="font-mono font-semibold text-slate-700">{fmtARS(pg.totalManoObra)}</div>
                            </div>
                            <div>
                              <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Equipos + Materiales (pres.)</div>
                              <div className="font-mono font-semibold text-slate-700">{fmtARS(pg.totalEquipos + pg.totalMateriales)}</div>
                            </div>
                            <div>
                              <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Total presupuestado sin IVA</div>
                              <div className="font-mono font-bold text-slate-900">{fmtARS(pg.precioTotalSinIva)}</div>
                            </div>
                            <div>
                              <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Gastado real hasta hoy</div>
                              <div className={`font-mono font-bold ${gastoRealObra > pg.precioTotalSinIva ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(gastoRealObra)}</div>
                            </div>
                          </div>
                          <div className="mt-3 h-2 overflow-hidden rounded-full bg-stone-100">
                            <div
                              className={`h-full ${gastoRealObra > pg.precioTotalSinIva ? "bg-rose-500" : "bg-emerald-500"}`}
                              style={{ width: `${Math.min(100, (gastoRealObra / (pg.precioTotalSinIva || 1)) * 100)}%` }}
                            />
                          </div>
                          <div className="mt-1 text-[11px] text-slate-400">
                            {gastoRealObra > pg.precioTotalSinIva
                              ? `Superó el presupuesto por ${fmtARS(gastoRealObra - pg.precioTotalSinIva)}.`
                              : `Queda ${fmtARS(pg.precioTotalSinIva - gastoRealObra)} disponible sobre lo presupuestado.`}
                          </div>
                        </Panel>
                      )}

                      {vistaMateriales === "equipos" && (
                        <Panel
                          title="Herramientas y equipos del inventario"
                          action={
                            <button onClick={() => setMostrarListaHerramientas((v) => !v)} className={btnGhost}>
                              {mostrarListaHerramientas ? "Ocultar" : "Ver todas"}
                            </button>
                          }
                        >
                          {!mostrarListaHerramientas ? (
                            <div className="text-xs text-slate-400">Antes de pedir un equipo o herramienta nueva, fijate si ya tenés uno libre en el inventario — te ahorrás la compra, solo hace falta mandarlo por remito.</div>
                          ) : (
                            <div className="overflow-x-auto rounded-md border border-stone-200">
                              <table className="w-full text-left text-sm">
                                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                                  <tr>
                                    <th className="px-2 py-2"></th><th className="px-3 py-2">Herramienta</th><th className="px-3 py-2">Ubicación</th>
                                    <th className="px-3 py-2">Estado</th><th className="px-3 py-2"></th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {herramientasOrdenadas.map((h) => (
                                    <tr key={h.id} className="border-t border-stone-100">
                                      <td className="px-2 py-1.5">
                                        {h.estado === "Disponible" && <span title="Libre, sin obra asignada"><CheckCircle2 size={14} className="text-emerald-600" /></span>}
                                      </td>
                                      <td className="px-3 py-1.5">
                                        <span className="flex items-center gap-1.5 font-medium text-slate-800">
                                          <CategoriaHerrIcon categoria={h.categoria} /> {h.nombre}
                                          <span className="font-mono text-xs text-slate-400">({h.numeroSerie || "s/n"})</span>
                                        </span>
                                      </td>
                                      <td className="px-3 py-1.5 text-slate-600">
                                        <span className="inline-flex items-center gap-1"><MapPin size={11} className="text-amber-600" />{h.ubicacion}</span>
                                      </td>
                                      <td className="px-3 py-1.5"><Badge estado={h.estado} /></td>
                                      <td className="px-3 py-1.5"><button onClick={() => agregarHerramientaAlPedido(h)} className={btnGhost}>+ Agregar al pedido</button></td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </Panel>
                      )}

                      {["epps", "consumibles"].includes(vistaMateriales) && !esGeneral && stockEppsOConsumibles.length > 0 && (
                        <Panel title="Pedir del depósito">
                          <div className="mb-2 text-[11px] text-slate-400">Esto ya está comprado y en stock — al pedirlo, se descuenta acá y el gasto pasa a esta obra.</div>
                          <div className="divide-y divide-stone-100">
                            {stockEppsOConsumibles.map((s) => (
                              <div key={s.id} className="flex flex-wrap items-center justify-between gap-2 py-2 text-sm">
                                <div>
                                  <span className="font-medium text-slate-800">{s.material}</span>
                                  <span className="ml-2 text-xs text-slate-400">{s.cantidad} {s.unidad} disponibles{canVerPreciosPedido ? ` · ${fmtARS(s.precioUnitario)} c/u` : ""}</span>
                                </div>
                                <div className="flex items-center gap-2">
                                  <input
                                    type="number" min="1" max={s.cantidad}
                                    value={cantidadPedirStock[s.id] ?? 1}
                                    onChange={(e) => setCantidadPedirStock((c) => ({ ...c, [s.id]: e.target.value }))}
                                    className="w-20 rounded border border-stone-300 px-1.5 py-1 text-right"
                                  />
                                  <button
                                    onClick={() => { pedirDeStockParaObra(s, cantidadPedirStock[s.id] ?? 1, obraIdActual); setCantidadPedirStock((c) => ({ ...c, [s.id]: 1 })); }}
                                    className={btnGhost}
                                  >
                                    Pedir para esta obra
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        </Panel>
                      )}

                      {["materiales", "equipos"].includes(vistaMateriales) && !showPedidoForm && (
                        <div className="flex items-center justify-end">
                          <button
                            onClick={() => { setItemManualDraft((d) => ({ ...d, categoria: vistaMateriales === "equipos" ? "Equipos" : "Materiales" })); setShowPedidoForm(true); }}
                            className="flex items-center gap-1.5 rounded-md border-2 border-rose-400 bg-rose-50 px-3 py-2 text-sm font-semibold text-rose-700 hover:bg-rose-100"
                          >
                            <AlertTriangle size={16} /> Pedido fuera de presupuesto
                          </button>
                        </div>
                      )}

                      {["epps", "consumibles"].includes(vistaMateriales) ? (
                        <Panel
                          title={`Catálogo de ${vistaMateriales === "epps" ? "Epps" : "Consumibles"}`}
                          action={
                            !showPedidoForm && (
                              <button onClick={() => setShowPedidoForm(true)} className={btnGhost}>+ Cargar uno nuevo</button>
                            )
                          }
                        >
                          {vistaMateriales === "epps" && (
                            <div className="mb-3 flex flex-wrap gap-2">
                              <select
                                value={filtroEppParte}
                                onChange={(e) => { setFiltroEppParte(e.target.value); setFiltroEppTipo("Todos"); }}
                                className="rounded-md border border-stone-300 bg-white px-2 py-1 text-xs"
                              >
                                <option value="Todas">Toda parte del cuerpo</option>
                                {partesCuerpoEpp.map((p) => <option key={p}>{p}</option>)}
                              </select>
                              {filtroEppParte !== "Todas" && tiposEppDeLaParte.length > 0 && (
                                <select
                                  value={filtroEppTipo}
                                  onChange={(e) => setFiltroEppTipo(e.target.value)}
                                  className="rounded-md border border-stone-300 bg-white px-2 py-1 text-xs"
                                >
                                  <option value="Todos">Todo tipo</option>
                                  {tiposEppDeLaParte.map((t) => <option key={t}>{t}</option>)}
                                </select>
                              )}
                            </div>
                          )}
                          {catalogoEppsOConsumibles.length === 0 ? (
                            <div className="text-xs text-slate-400">Todavía no hay {vistaMateriales === "epps" ? "Epps" : "Consumibles"} cargados{filtroEppParte !== "Todas" ? ` para "${filtroEppParte}"` : ""}. Tocá "+ Cargar uno nuevo" para tipearlo — queda guardado para la próxima vez.</div>
                          ) : (
                            <div className="divide-y divide-stone-100">
                              {catalogoEppsOConsumibles.map((it) => (
                                <div key={it.id} className="flex flex-wrap items-center justify-between gap-2 py-2 text-sm">
                                  <div>
                                    <span className="font-medium text-slate-800">{it.nombre}</span>
                                    <span className="ml-2 text-xs text-slate-400">
                                      {vistaMateriales === "epps" && it.subcategoria ? `${it.subcategoria}${it.tipo ? ` · ${it.tipo}` : ""} · ` : ""}
                                      {it.unidad || "und."}
                                    </span>
                                  </div>
                                  <div className="flex items-center gap-3">
                                    {canVerPreciosPedido && <span className="font-mono text-xs text-slate-500">{fmtARS(it.ultimoPrecio)}</span>}
                                    <button onClick={() => agregarCatalogoAlPedido(it)} className={btnGhost}>+ Agregar al pedido</button>
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                        </Panel>
                      ) : lineasObra.length === 0 ? (
                        <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                          Todavía no hay presupuesto importado para esta obra{vistaMateriales === "equipos" ? " (Equipos/Herramientas)" : ""}.
                        </div>
                      ) : (
                        <>
                          <div className="flex flex-wrap items-center justify-between gap-3">
                            {canVerPreciosPedido && (
                              <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                                <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                                  {vistaMateriales === "equipos" ? "Equipos y herramientas importados (sin IVA)" : "Materiales importados (sin IVA)"}
                                </div>
                                <div className="mt-1 font-mono text-xl font-bold text-slate-900">{fmtARS(totalObra)}</div>
                              </div>
                            )}
                            {seleccionPresupuesto.length > 0 && (
                              <button onClick={abrirArmadoPedido} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                                <ShoppingCart size={16} /> Armar pedido con {seleccionPresupuesto.length} ítem(s)
                              </button>
                            )}
                          </div>
                          <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-stone-50 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
                                <tr>
                                  <th className="px-2 py-1.5"></th>
                                  <th className="px-2 py-1.5 text-right">#</th>
                                  <th className="px-2 py-1.5">Rubro</th><th className="px-2 py-1.5">Descripción</th>
                                  <th className="px-2 py-1.5">Unidad</th><th className="px-2 py-1.5 text-right">Cantidad</th>
                                  {canVerPreciosPedido && (<><th className="px-2 py-1.5 text-right">P. Unitario</th><th className="px-2 py-1.5 text-right">Total</th></>)}
                                  <th className="px-2 py-1.5"></th>
                                </tr>
                              </thead>
                              <tbody>
                                {lineasObra.map((m, i) => {
                                  const numero = i + 1;
                                  const yaPedido = !!m.pedidoId;
                                  const seleccionada = seleccionPresupuesto.includes(m.id);
                                  return (
                                    <tr
                                      key={m.id}
                                      onClick={(e) => { if (!yaPedido && !e.target.closest("input, button")) toggleSeleccionPresupuesto(m.id); }}
                                      className={`border-t border-stone-100 leading-tight ${yaPedido ? "opacity-50" : "cursor-pointer hover:bg-amber-50"} ${seleccionada ? "bg-amber-50" : ""}`}
                                    >
                                      <td className="px-2 py-1">
                                        {yaPedido ? (
                                          <span title="Ya está en un pedido"><ShoppingCart size={12} className="text-slate-400" /></span>
                                        ) : (
                                          <input type="checkbox" checked={seleccionada} onChange={() => toggleSeleccionPresupuesto(m.id)} className="h-3.5 w-3.5" />
                                        )}
                                      </td>
                                      <td className="px-2 py-1 text-right font-mono text-slate-400">{numero}</td>
                                      <td className="px-2 py-1 text-slate-600">{m.subcategoria || "—"}</td>
                                      <td className="px-2 py-1 font-medium text-slate-900">{m.material}</td>
                                      <td className="px-2 py-1 text-slate-600">{m.unidad}</td>
                                      <td className="px-2 py-1 text-right font-mono">{m.cantidad}</td>
                                      {canVerPreciosPedido && (
                                        <>
                                          <td className="px-2 py-1 text-right font-mono">{fmtARS(m.precioUnitario)}</td>
                                          <td className="px-2 py-1 text-right font-mono">{fmtARS(m.total)}</td>
                                        </>
                                      )}
                                      <td className="px-2 py-1">{!yaPedido && <button onClick={() => eliminarLineaPresupuesto(m.id)} className="text-slate-400 hover:text-rose-600"><X size={13} /></button>}</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </>
                      )}

                      {pedidosObra.length > 0 && (
                        <div>
                          <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Pedidos de esta obra</div>
                          <div className="space-y-3">
                            {[...pedidosObra].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((p) => {
                              const dias = p.fechaNecesaria ? diasHasta(p.fechaNecesaria) : null;
                              const urgencia = !pedidoEnCurso(p) || dias === null ? "text-slate-400" : dias < 0 ? "font-semibold text-rose-600" : dias <= 2 ? "font-semibold text-amber-700" : "text-slate-500";
                              return (
                                <div key={p.id} className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                                  <div className="flex flex-wrap items-center justify-between gap-2">
                                    <div>
                                      <span className="font-semibold text-slate-900">{canVerPreciosPedido ? (p.proveedor || "Proveedor sin especificar") : `Pedido #${p.id}`}</span>
                                      <span className="ml-2"><Badge estado={p.estado} /></span>
                                      {p.obraId == null && <span className="ml-2 rounded-full border border-sky-300 bg-sky-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-sky-700">Compra general</span>}
                                      {p.items.length > 0 && p.items.every((it) => !it.presupuestoId) && <span className="ml-2 rounded-full border border-rose-300 bg-rose-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-rose-700">Fuera de presupuesto</span>}
                                    </div>
                                    <span className="text-xs text-slate-400">{fmtFecha(p.fecha)} · {p.items.length} ítem(s){canVerPreciosPedido && <> · <span className="font-mono font-semibold text-slate-600">{fmtARS(p.total)}</span></>}</span>
                                  </div>
                                  <div className="mt-2 text-xs text-slate-500">{p.items.map((it) => it.material).join(", ")}</div>
                                  {p.fechaNecesaria && (
                                    <div className={`mt-1 text-xs ${urgencia}`}>
                                      Necesario para el {fmtFecha(p.fechaNecesaria)}
                                      {(p.estado === "Solicitado" || p.estado === "Aprobado") && dias !== null && dias < 0 && " — VENCIDO"}
                                    </div>
                                  )}

                                  {p.estado === "Solicitado" && (
                                    puedeAprobarPedidos ? (
                                      <div className="mt-2 flex gap-2">
                                        <button onClick={() => aprobarPedidoMaterial(p)} className="flex items-center gap-1 rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700">
                                          <Check size={13} /> Aprobar pedido
                                        </button>
                                        <button onClick={() => rechazarPedidoMaterial(p)} className="flex items-center gap-1 rounded-md border border-rose-300 px-3 py-1.5 text-xs font-semibold text-rose-600 hover:bg-rose-50">
                                          <X size={13} /> Rechazar
                                        </button>
                                      </div>
                                    ) : (
                                      <div className="mt-2 text-xs italic text-slate-400">Esperando aprobación de Gerencia.</div>
                                    )
                                  )}
                                  {p.estado === "Aprobado" && (
                                    canVerPreciosPedido ? (
                                      <div className="mt-2 flex flex-wrap gap-2">
                                        <button onClick={() => generarOrdenCompraPDF(p)} className={btnGhost}>
                                          <span className="flex items-center gap-1"><FileDown size={13} /> Orden de Compra (PDF)</span>
                                        </button>
                                        {p.obraId == null ? (
                                          <button onClick={() => recibirPedidoGeneralAStock(p)} className="flex items-center gap-1 rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">
                                            <Package size={13} /> Marcar recibido en depósito
                                          </button>
                                        ) : (
                                          <button onClick={() => abrirCargaFactura(p)} className="flex items-center gap-1 rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">
                                            <Receipt size={13} /> Cargar factura real
                                          </button>
                                        )}
                                      </div>
                                    ) : (
                                      <div className="mt-2 text-xs italic text-slate-400">Aprobado — a la espera de que Logística cargue la orden de compra.</div>
                                    )
                                  )}
                                  {p.estado === "Facturado" && (
                                    <div className="mt-2 text-xs text-slate-500">
                                      Facturado{p.comprobante ? ` — comprobante ${p.comprobante}` : ""}. El gasto ya quedó imputado a esta obra.
                                    </div>
                                  )}

                                  {facturandoPedidoId === p.id && (
                                    <div className="mt-3 rounded-md border border-stone-200 bg-stone-50 p-3">
                                      <div className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Corregí los precios con la factura real</div>
                                      <div className="mb-3 max-w-xs">
                                        <Field label="Proveedor">
                                          <select value={proveedorFacturaDraft} onChange={(e) => setProveedorFacturaDraft(e.target.value)} className={inputCls}>
                                            <option value="">Sin especificar</option>
                                            {proveedorFacturaDraft && !proveedores.some((pr) => pr.razonSocial === proveedorFacturaDraft) && (
                                              <option value={proveedorFacturaDraft}>{proveedorFacturaDraft} (cotizado inicialmente)</option>
                                            )}
                                            {proveedores.filter((pr) => pr.esTaller !== "Sí").map((pr) => <option key={pr.id} value={pr.razonSocial}>{pr.razonSocial}</option>)}
                                          </select>
                                        </Field>
                                      </div>
                                      <div className="space-y-1.5">
                                        {itemsFacturaDraft.map((it, idx) => {
                                          const esEquipo = ["Equipos", "Herramientas"].includes(it.categoria);
                                          return (
                                            <div key={idx} className="flex flex-wrap items-center justify-between gap-2 text-xs">
                                              <span className="text-slate-700">{it.material} <span className="text-slate-400">({it.cantidad} {it.unidad})</span></span>
                                              <div className="flex items-center gap-1">
                                                {esEquipo && (
                                                  <select
                                                    value={it.tipoEquipo || ""}
                                                    onChange={(e) => actualizarTipoEquipoFactura(idx, e.target.value)}
                                                    className="rounded border border-stone-300 px-1 py-1 text-[11px]"
                                                  >
                                                    <option value="">¿Propio o alquilado?</option>
                                                    <option value="Propio">Propio</option>
                                                    <option value="Alquilado">Alquilado</option>
                                                  </select>
                                                )}
                                                <span className="text-slate-400">$</span>
                                                <MoneyInput
                                                  key={`${idx}-${it.tipoEquipo || ""}-${facturaRevision}`}
                                                  value={it.precioUnitario}
                                                  onChange={(v) => actualizarPrecioFactura(idx, v)}
                                                  disabled={esEquipo && it.tipoEquipo === "Propio"}
                                                  className="w-24 rounded border border-stone-300 px-1.5 py-1 text-right"
                                                />
                                                <span className="w-24 text-right font-mono text-slate-600">{fmtARS((Number(it.cantidad) || 0) * (Number(it.precioUnitario) || 0))}</span>
                                              </div>
                                            </div>
                                          );
                                        })}
                                      </div>
                                      {itemsFacturaDraft.some((it) => ["Equipos", "Herramientas"].includes(it.categoria) && !it.tipoEquipo) && (
                                        <div className="mt-1 text-[11px] text-amber-700">Faltan marcar como Propio o Alquilado los equipos/herramientas de arriba.</div>
                                      )}

                                      <div className="mt-3 flex flex-wrap items-end gap-2 rounded-md border border-dashed border-stone-300 p-2">
                                        <div className="w-40">
                                          <Field label="¿Con apuro? Un total y listo">
                                            <MoneyInput value={totalFacturaRapido} onChange={setTotalFacturaRapido} className="w-full rounded border border-stone-300 px-1.5 py-1 text-right text-xs" />
                                          </Field>
                                        </div>
                                        <button type="button" onClick={() => aplicarTotalFactura(totalFacturaRapido)} className={btnGhost}>Repartir entre los ítems</button>
                                        <div className="text-[11px] text-slate-400">Reparte el total a prorrata de cantidad — no toca lo marcado "Propio".</div>
                                      </div>

                                      <div className="mt-2 grid grid-cols-1 gap-2 sm:grid-cols-2">
                                        <input
                                          value={comprobanteDraft}
                                          onChange={(e) => setComprobanteDraft(e.target.value)}
                                          placeholder="N° de comprobante (ej: A-0001-00012345)"
                                          className={inputCls}
                                        />
                                        <div className="flex items-center justify-end font-mono text-sm font-bold text-slate-900">
                                          Total: {fmtARS(itemsFacturaDraft.reduce((s, it) => s + (Number(it.cantidad) || 0) * (Number(it.precioUnitario) || 0), 0))}
                                        </div>
                                      </div>
                                      <div className="mt-2 flex gap-2">
                                        <button disabled={guardandoFactura} onClick={() => confirmarFacturaReal(p)} className="rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700 disabled:opacity-50">
                                          {guardandoFactura ? "Guardando..." : "Confirmar factura"}
                                        </button>
                                        <button onClick={() => setFacturandoPedidoId(null)} className="text-xs text-slate-400 hover:underline">Cancelar</button>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </>
                  );
                })()}
              </>
            )}

            {vistaMateriales === "catalogo" && (
              <>
                <Panel title="Sub-categorías" action={null}>
                  <div className="flex flex-wrap items-end gap-2">
                    <div className="w-40">
                      <Field label="Categoría">
                        <select value={categoriaParaSubcat} onChange={(e) => setCategoriaParaSubcat(e.target.value)} className={inputCls}>
                          {CATEGORIAS_PEDIDO.map((c) => <option key={c}>{c}</option>)}
                        </select>
                      </Field>
                    </div>
                    <div className="flex-1 min-w-[160px]">
                      <Field label="Nueva sub-categoría">
                        <input value={nuevaSubcategoria} onChange={(e) => setNuevaSubcategoria(e.target.value)} placeholder="Ej: Electricidad, Plomería..." className={inputCls} />
                      </Field>
                    </div>
                    <button onClick={agregarSubcategoria} className={btnGhost}>+ Agregar</button>
                  </div>
                  <div className="mt-3 space-y-2">
                    {CATEGORIAS_PEDIDO.map((cat) => {
                      const subs = subcategoriasMat.filter((s) => s.categoria === cat);
                      if (subs.length === 0) return null;
                      return (
                        <div key={cat} className="flex flex-wrap items-center gap-2">
                          <span className="text-xs font-semibold text-slate-500">{cat}:</span>
                          {subs.map((s) => (
                            <span key={s.id} className="rounded-full border border-stone-300 bg-stone-50 px-3 py-1 text-xs text-slate-700">{s.nombre}</span>
                          ))}
                        </div>
                      );
                    })}
                  </div>
                </Panel>

                <Panel title="Tipos (dentro de cada sub-categoría)" action={null}>
                  <div className="flex flex-wrap items-end gap-2">
                    <div className="w-40">
                      <Field label="Categoría">
                        <select
                          value={categoriaParaTipo}
                          onChange={(e) => { setCategoriaParaTipo(e.target.value); setSubcategoriaParaTipo(""); }}
                          className={inputCls}
                        >
                          {CATEGORIAS_PEDIDO.map((c) => <option key={c}>{c}</option>)}
                        </select>
                      </Field>
                    </div>
                    <div className="w-44">
                      <Field label="Sub-categoría">
                        <select value={subcategoriaParaTipo} onChange={(e) => setSubcategoriaParaTipo(e.target.value)} className={inputCls}>
                          <option value="">-- Elegir --</option>
                          {subcategoriasMat.filter((s) => s.categoria === categoriaParaTipo).map((s) => <option key={s.id} value={s.nombre}>{s.nombre}</option>)}
                        </select>
                      </Field>
                    </div>
                    <div className="flex-1 min-w-[160px]">
                      <Field label="Nuevo tipo">
                        <input value={nuevoTipo} onChange={(e) => setNuevoTipo(e.target.value)} placeholder="Ej: Hierros Nervurados, Daisa..." className={inputCls} />
                      </Field>
                    </div>
                    <button onClick={agregarTipoMaterial} className={btnGhost}>+ Agregar</button>
                  </div>
                  {subcategoriasMat.filter((s) => s.categoria === categoriaParaTipo).length === 0 && (
                    <div className="mt-2 text-[11px] text-slate-400">Primero creá una sub-categoría de "{categoriaParaTipo}" arriba, para poder agregarle tipos.</div>
                  )}
                  <div className="mt-3 space-y-2">
                    {subcategoriasMat.map((sub) => {
                      const tipos = tiposMaterial.filter((t) => t.categoria === sub.categoria && t.subcategoria === sub.nombre);
                      if (tipos.length === 0) return null;
                      return (
                        <div key={sub.id} className="flex flex-wrap items-center gap-2">
                          <span className="text-xs font-semibold text-slate-500">{sub.categoria} → {sub.nombre}:</span>
                          {tipos.map((t) => (
                            <span key={t.id} className="rounded-full border border-amber-300 bg-amber-50 px-3 py-1 text-xs text-amber-800">{t.nombre}</span>
                          ))}
                        </div>
                      );
                    })}
                  </div>
                </Panel>

                <div className="text-xs text-slate-500">
                  Se completa solo con lo que vas importando — el precio que ves es el último cargado para cada material.
                </div>

                {catalogoMateriales.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    Todavía no hay materiales en el catálogo. Importá un presupuesto para empezar a llenarlo.
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                        <tr>
                          <th className="px-4 py-3">Categoría</th><th className="px-4 py-3">Sub-categoría</th><th className="px-4 py-3">Tipo</th><th className="px-4 py-3">Material</th>
                          <th className="px-4 py-3">Unidad</th><th className="px-4 py-3 text-right">Último precio</th><th className="px-4 py-3">Último proveedor</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...catalogoMateriales].sort((a, b) => a.nombre.localeCompare(b.nombre)).map((m) => (
                          <tr key={m.id} className="border-t border-stone-100">
                            <td className="px-4 py-2.5 text-slate-600">{m.categoria}</td>
                            <td className="px-4 py-2.5 text-slate-600">{m.subcategoria || "—"}</td>
                            <td className="px-4 py-2.5 text-slate-600">{m.tipo || "—"}</td>
                            <td className="px-4 py-2.5 font-medium text-slate-900">{m.nombre}</td>
                            <td className="px-4 py-2.5 text-slate-600">{m.unidad}</td>
                            <td className="px-4 py-2.5 text-right font-mono">{fmtARS(m.ultimoPrecio)}</td>
                            <td className="px-4 py-2.5 text-slate-500">{m.ultimoProveedor || <span className="text-slate-300">Sin datos aún</span>}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </>
            )}

            {vistaMateriales === "consolidar" && (
              <>
                <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
                  Pedidos pendientes de todas las obras, agrupados por proveedor — así logística compra todo junto (ej: todo el cemento) en un solo viaje. El proveedor se sugiere solo según la última vez que se compró ese material; lo podés cambiar antes de generar los remitos.
                </div>

                {Object.keys(gruposPorProveedor).length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    No hay pedidos pendientes de enviar en este momento.
                  </div>
                ) : (
                  <div className="space-y-4">
                    {Object.entries(gruposPorProveedor).map(([proveedor, pedidosGrupo]) => {
                      const totalGrupo = pedidosGrupo.reduce((s, p) => s + p.total, 0);
                      const obrasInvolucradas = [...new Set(pedidosGrupo.map((p) => obras.find((o) => o.id === p.obraId)?.nombre).filter(Boolean))];
                      return (
                        <div key={proveedor} className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <div className="flex items-center gap-2">
                              <Truck size={16} className="text-amber-600" />
                              <span className="font-semibold text-slate-900">{proveedor}</span>
                              {proveedor === "Sin proveedor asignado" && <Badge estado="Rota" />}
                            </div>
                            <span className="text-xs text-slate-400">{obrasInvolucradas.length} obra(s) · <span className="font-mono font-semibold text-slate-600">{fmtARS(totalGrupo)}</span></span>
                          </div>

                          <div className="space-y-2">
                            {pedidosGrupo.map((p) => {
                              const obra = obras.find((o) => o.id === p.obraId);
                              return (
                                <div key={p.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md bg-stone-50 px-3 py-2 text-sm">
                                  <div>
                                    <span className="font-medium text-slate-800">{obra?.nombre}</span>
                                    <span className="ml-2 text-xs text-slate-500">{p.items.map((it) => it.material).join(", ")}</span>
                                  </div>
                                  <div className="flex items-center gap-2">
                                    <span className="font-mono text-xs text-slate-600">{fmtARS(p.total)}</span>
                                    {!p.proveedor && (
                                      <select
                                        value=""
                                        onChange={(e) => e.target.value && asignarProveedorPedido(p, e.target.value)}
                                        className="rounded-md border border-stone-300 bg-white px-1.5 py-1 text-xs text-slate-600"
                                      >
                                        <option value="">Asignar proveedor...</option>
                                        {proveedores.filter((pr) => pr.esTaller !== "Sí").map((pr) => <option key={pr.id} value={pr.razonSocial}>{pr.razonSocial}</option>)}
                                      </select>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>

                          <button
                            onClick={() => generarRemitosDelGrupo(proveedor === "Sin proveedor asignado" ? pedidosGrupo[0]?.proveedor : proveedor, pedidosGrupo)}
                            disabled={proveedor === "Sin proveedor asignado"}
                            className="mt-3 flex items-center gap-1 rounded-md bg-slate-900 px-3 py-2 text-sm font-semibold text-white hover:bg-slate-700 disabled:cursor-not-allowed disabled:opacity-40"
                          >
                            <ArrowRightLeft size={14} /> Generar remitos por obra ({pedidosGrupo.length})
                          </button>
                          {proveedor === "Sin proveedor asignado" && (
                            <div className="mt-1 text-[11px] text-slate-400">Asigná un proveedor a cada pedido de este grupo antes de generar los remitos.</div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}

                {remitosMaterialesPendientes.length > 0 && (
                  <div>
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Remitos de materiales en tránsito</div>
                    <div className="space-y-3">
                      {remitosMaterialesPendientes.map((r) => (
                        <div key={r.id} className="rounded-lg border border-amber-300 bg-amber-50 p-4">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <span className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                              {r.origen} <ArrowRightLeft size={14} className="text-amber-600" /> {r.destino}
                            </span>
                            <span className="text-xs text-slate-500">Salió el {fmtFecha(r.fecha)} ({r.creadoPor})</span>
                          </div>
                          <div className="mt-1 text-xs text-slate-600">{r.materialItems.map((it) => it.material).join(", ")}</div>

                          {recibiendoRemitoId !== r.id ? (
                            <button onClick={() => confirmarRecepcionRemito(r)} className="mt-2 flex items-center gap-1 rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">
                              <Check size={13} /> Confirmar recepción en {r.destino}
                            </button>
                          ) : (
                            <div className="mt-3 rounded-md border border-amber-200 bg-white p-3">
                              <div className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                                ¿Cuánto entra a "{r.destino}"? Lo que no entra queda en Stock general.
                              </div>
                              <div className="space-y-2">
                                {r.materialItems.map((it, idx) => {
                                  const aObra = cantidadesRecepcion[idx] ?? it.cantidad;
                                  const aStock = Math.max(0, it.cantidad - aObra);
                                  return (
                                    <div key={idx} className="flex flex-wrap items-center justify-between gap-2 text-xs">
                                      <span className="text-slate-700">{it.material} <span className="text-slate-400">(compradas: {it.cantidad} {it.unidad})</span></span>
                                      <div className="flex items-center gap-2">
                                        <span className="text-slate-400">A la obra:</span>
                                        <input
                                          type="number"
                                          min="0"
                                          max={it.cantidad}
                                          value={aObra}
                                          onChange={(e) => actualizarCantidadRecepcion(idx, e.target.value, it.cantidad)}
                                          className="w-20 rounded border border-stone-300 px-1.5 py-1 text-right"
                                        />
                                        {aStock > 0 && <span className="text-amber-700">→ {aStock} {it.unidad} a stock</span>}
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                              <div className="mt-3 flex gap-2">
                                <button disabled={guardandoRecepcion} onClick={() => confirmarRecepcionMaterialConDivision(r)} className="rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700 disabled:opacity-50">
                                  {guardandoRecepcion ? "Guardando..." : "Confirmar recepción"}
                                </button>
                                <button onClick={() => setRecibiendoRemitoId(null)} className="text-xs text-slate-400 hover:underline">Cancelar</button>
                              </div>
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {remitosMaterialesCompletados.length > 0 && (
                  <div>
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Historial de remitos de materiales recibidos</div>
                    <div className="space-y-2">
                      {[...remitosMaterialesCompletados].sort((a, b) => fechaLocal(b.fechaRecepcion) - fechaLocal(a.fechaRecepcion)).map((r) => (
                        <div key={r.id} className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-stone-200 bg-white px-4 py-2 text-sm">
                          <span className="flex items-center gap-2 text-slate-700">
                            {r.origen} <ArrowRightLeft size={13} className="text-slate-400" /> {r.destino}
                          </span>
                          <span className="text-xs text-slate-400">Recibido el {fmtFecha(r.fechaRecepcion)} ({r.recibidoPor})</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </>
            )}

            {vistaMateriales === "stock" && (
              <>
                <div className="rounded-md border border-stone-200 bg-white px-4 py-2 text-xs text-slate-500">
                  Lo que compraste pero no entró completo a una obra queda acá — el gasto ya se hizo y quedó anotado como "general" hasta que lo asignes a alguna obra.
                </div>

                {stockMateriales.filter((s) => s.cantidad > 0).length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    No hay materiales en stock general por ahora.
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                        <tr>
                          <th className="px-4 py-3">Material</th><th className="px-4 py-3">Categoría</th><th className="px-4 py-3 text-right">Cantidad</th>
                          <th className="px-4 py-3 text-right">Precio unitario</th><th className="px-4 py-3 text-right">Valor en stock</th><th className="px-4 py-3"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {stockMateriales.filter((s) => s.cantidad > 0).map((s) => (
                          <Fragment key={s.id}>
                            <tr className="border-t border-stone-100">
                              <td className="px-4 py-2.5 font-medium text-slate-900">{s.material}</td>
                              <td className="px-4 py-2.5 text-slate-600">{s.categoria}{s.subcategoria ? ` — ${s.subcategoria}` : ""}</td>
                              <td className="px-4 py-2.5 text-right font-mono">{s.cantidad} {s.unidad}</td>
                              <td className="px-4 py-2.5 text-right font-mono">{fmtARS(s.precioUnitario)}</td>
                              <td className="px-4 py-2.5 text-right font-mono">{fmtARS(s.cantidad * s.precioUnitario)}</td>
                              <td className="px-4 py-2.5">
                                <button onClick={() => { setAsignandoStockId(s.id); setObraParaStock(""); setCantidadParaStock(s.cantidad); }} className={btnGhost}>Asignar a obra</button>
                              </td>
                            </tr>
                            {asignandoStockId === s.id && (
                              <tr className="border-t border-stone-100 bg-stone-50">
                                <td colSpan={6} className="px-4 py-3">
                                  <div className="flex flex-wrap items-end gap-2">
                                    <div className="w-48">
                                      <Field label="Obra destino">
                                        <select value={obraParaStock} onChange={(e) => setObraParaStock(e.target.value)} className={inputCls}>
                                          <option value="">-- Elegir --</option>
                                          {obras.filter((o) => o.estado === "En curso").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                                        </select>
                                      </Field>
                                    </div>
                                    <div className="w-28">
                                      <Field label="Cantidad">
                                        <input type="number" min="1" max={s.cantidad} value={cantidadParaStock} onChange={(e) => setCantidadParaStock(e.target.value)} className={inputCls} />
                                      </Field>
                                    </div>
                                    <button onClick={() => asignarStockAObra(s)} className="rounded-md bg-slate-900 px-3 py-2 text-xs font-semibold text-white hover:bg-slate-700">Confirmar</button>
                                    <button onClick={() => setAsignandoStockId(null)} className="text-xs text-slate-400 hover:underline">Cancelar</button>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </Fragment>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </>
            )}
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
                  <Field label="Monto estimado (ARS)"><MoneyInput name="montoEstimado" className={inputCls} /></Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="space-y-3">
              {ordenesCompra.filter((oc) => !obraIdsPapelera.has(oc.obraId)).map((oc) => {
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
                    <select name="obraId" className={inputCls}>{obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}</select>
                  </Field>
                  <Field label="Orden de compra (opcional)">
                    <select name="ordenCompraId" className={inputCls}>
                      <option value="">Sin orden asociada</option>
                      {ordenesCompra.filter((oc) => !obraIdsPapelera.has(oc.obraId)).map((oc) => <option key={oc.id} value={oc.id}>#{oc.id} · {oc.proveedor}</option>)}
                    </select>
                  </Field>
                  <Field label="Proveedor"><input name="proveedor" required className={inputCls} /></Field>
                  <Field label="Categoría">
                    <select name="categoria" className={inputCls}>{CATEGORIAS_GASTO.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <Field label="Monto (ARS)"><MoneyInput name="monto" className={inputCls} /></Field>
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
                  {comprasFacturas.filter((c) => !obraIdsPapelera.has(c.obraId)).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((c) => {
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
                    <select name="obraId" className={inputCls}>{obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}</select>
                  </Field>
                  <Field label="Concepto"><input name="concepto" required placeholder="Ej: certificado de avance 3" className={inputCls} /></Field>
                  <Field label="Monto (ARS)"><MoneyInput name="monto" className={inputCls} /></Field>
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
                  {ingresos.filter((i) => !obraIdsPapelera.has(i.obraId)).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((i) => {
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

        {tab === "proveedores" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Clientes/Proveedores</h2>

            <div className="flex flex-wrap gap-2">
              <button
                onClick={() => setVistaClientesProveedores("clientes")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaClientesProveedores === "clientes" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Clientes
              </button>
              <button
                onClick={() => setVistaClientesProveedores("proveedores")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaClientesProveedores === "proveedores" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Proveedores
              </button>
            </div>

            {vistaClientesProveedores === "clientes" && (
              <>
                <div className="flex items-center justify-between">
                  <div className="text-xs text-slate-500">El saldo es lo acordado en el presupuesto de sus obras menos lo que ya ingresó por esa obra.</div>
                  <button onClick={() => setShowClienteForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                    <Plus size={16} /> Nuevo cliente
                  </button>
                </div>

                {showClienteForm && (
                  <Panel title="Añadir cliente" action={<button onClick={() => setShowClienteForm(false)}><X size={16} /></button>}>
                    <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitClienteForm}>
                      <Field label="Razón social / Nombre">
                        <input value={clienteForm.razonSocial} onChange={(e) => setClienteForm((f) => ({ ...f, razonSocial: e.target.value }))} required className={inputCls} />
                      </Field>
                      <Field label="CUIT / DNI">
                        <input value={clienteForm.cuit} onChange={(e) => setClienteForm((f) => ({ ...f, cuit: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="Domicilio">
                        <input value={clienteForm.domicilio} onChange={(e) => setClienteForm((f) => ({ ...f, domicilio: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="Contacto">
                        <input value={clienteForm.contacto} onChange={(e) => setClienteForm((f) => ({ ...f, contacto: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="Teléfono">
                        <input value={clienteForm.telefono} onChange={(e) => setClienteForm((f) => ({ ...f, telefono: e.target.value }))} className={inputCls} />
                      </Field>
                      <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                    </form>
                  </Panel>
                )}

                {clientes.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">Todavía no hay clientes cargados.</div>
                ) : (
                  <div className="space-y-3">
                    {clientes.map((cli) => {
                      const { obrasCliente, totalAcordado, totalCobrado, saldo } = balanceCliente(cli);
                      return (
                        <div key={cli.id} className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <div>
                              <div className="font-semibold text-slate-900">{cli.razonSocial}</div>
                              <div className="text-xs text-slate-500">{cli.contacto}{cli.contacto && cli.telefono ? " · " : ""}{cli.telefono}</div>
                            </div>
                            <div className="text-right">
                              <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Saldo — te debe</div>
                              <div className={`font-mono text-lg font-bold ${saldo > 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(saldo)}</div>
                            </div>
                          </div>
                          <div className="mt-2 flex flex-wrap gap-4 text-xs text-slate-500">
                            <span>Acordado: <span className="font-mono text-slate-700">{fmtARS(totalAcordado)}</span></span>
                            <span>Cobrado: <span className="font-mono text-slate-700">{fmtARS(totalCobrado)}</span></span>
                            <span>Obras: {obrasCliente.map((o) => o.nombre).join(", ") || "—"}</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </>
            )}

            {vistaClientesProveedores === "proveedores" && (
              <>
                <div className="flex items-center justify-between">
                  <div className="text-xs text-slate-500">Los talleres de reparación también se cargan acá — así aparecen como destino posible en los remitos de Herramientas. El saldo es lo facturado menos lo ya pagado.</div>
                  <button onClick={() => setShowProveedorForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400">
                    <Plus size={16} /> Nuevo proveedor
                  </button>
                </div>

                {showProveedorForm && (
                  <Panel title="Añadir proveedor" action={<button onClick={() => setShowProveedorForm(false)}><X size={16} /></button>}>
                    <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitProveedorForm}>
                      <Field label="Razón social">
                        <input value={proveedorForm.razonSocial} onChange={(e) => setProveedorForm((f) => ({ ...f, razonSocial: e.target.value }))} required className={inputCls} />
                      </Field>
                      <Field label="CUIT">
                        <input value={proveedorForm.cuit} onChange={(e) => setProveedorForm((f) => ({ ...f, cuit: e.target.value }))} placeholder="30-12345678-9" className={inputCls} />
                      </Field>
                      <Field label="Domicilio">
                        <input value={proveedorForm.domicilio} onChange={(e) => setProveedorForm((f) => ({ ...f, domicilio: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="Contacto">
                        <input value={proveedorForm.contacto} onChange={(e) => setProveedorForm((f) => ({ ...f, contacto: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="Teléfono">
                        <input value={proveedorForm.telefono} onChange={(e) => setProveedorForm((f) => ({ ...f, telefono: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="¿Es taller de reparación?">
                        <select value={proveedorForm.esTaller} onChange={(e) => setProveedorForm((f) => ({ ...f, esTaller: e.target.value }))} className={inputCls}>
                          {SI_NO.map((s) => <option key={s}>{s}</option>)}
                        </select>
                      </Field>
                      <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                    </form>
                  </Panel>
                )}

                {proveedores.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">Todavía no hay proveedores cargados.</div>
                ) : (
                  <div className="space-y-3">
                    {proveedores.map((p) => {
                      const { totalFacturado, totalPagado, saldo, facturasPendientes } = balanceProveedor(p);
                      return (
                        <div key={p.id} className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <div>
                              <span className="font-semibold text-slate-900">{p.razonSocial}</span>
                              {p.esTaller === "Sí" && <span className="ml-2"><Badge estado="En Reparación" /></span>}
                              <div className="text-xs text-slate-500">{p.contacto}{p.contacto && p.telefono ? " · " : ""}{p.telefono}</div>
                            </div>
                            <div className="text-right">
                              <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Saldo — le debés</div>
                              <div className={`font-mono text-lg font-bold ${saldo > 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(saldo)}</div>
                            </div>
                          </div>
                          <div className="mt-2 flex flex-wrap gap-4 text-xs text-slate-500">
                            <span>Facturado: <span className="font-mono text-slate-700">{fmtARS(totalFacturado)}</span></span>
                            <span>Pagado: <span className="font-mono text-slate-700">{fmtARS(totalPagado)}</span></span>
                          </div>
                          {facturasPendientes.length > 0 && (
                            <div className="mt-3 space-y-1 border-t border-stone-100 pt-2">
                              {facturasPendientes.map((f) => (
                                <div key={f.id} className="flex flex-wrap items-center justify-between gap-2 text-xs">
                                  <span className="text-slate-600">{fmtFecha(f.fecha)} — {f.comprobante || "sin comprobante"} — <span className="font-mono">{fmtARS(f.monto)}</span></span>
                                  <button onClick={() => marcarFacturaPagada(f)} className={btnGhost}>Marcar pagada</button>
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
          </div>
        )}

        {tab === "calendario" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Calendario Corporativo</h2>

            <Panel title="Feriados y días especiales" action={
              <button onClick={() => setShowFeriadoForm((v) => !v)} className="flex items-center gap-1 rounded-md bg-amber-500 px-3 py-1.5 text-xs font-semibold text-slate-900 hover:bg-amber-400">
                <Plus size={14} /> Agregar
              </button>
            }>
              {showFeriadoForm && (
                <form className="mb-4 grid grid-cols-1 gap-3 rounded-md border border-stone-200 bg-stone-50 p-3 md:grid-cols-3" onSubmit={submitFeriadoForm}>
                  <Field label="Fecha">
                    <input type="date" required value={feriadoForm.fecha} onChange={(e) => setFeriadoForm((f) => ({ ...f, fecha: e.target.value }))} className={inputCls} />
                  </Field>
                  <div className="md:col-span-2">
                    <Field label="Descripción">
                      <input required value={feriadoForm.descripcion} onChange={(e) => setFeriadoForm((f) => ({ ...f, descripcion: e.target.value }))} placeholder="Ej: Feriado nacional, cierre de balance..." className={inputCls} />
                    </Field>
                  </div>
                  <div className="md:col-span-3">
                    <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button>
                  </div>
                </form>
              )}
              {feriados.length === 0 ? (
                <div className="text-sm text-slate-400">Sin feriados cargados.</div>
              ) : (
                <div className="space-y-1">
                  {[...feriados].sort((a, b) => fechaLocal(a.fecha) - fechaLocal(b.fecha)).map((f) => (
                    <div key={f.id} className="flex items-center justify-between rounded-md px-2 py-1.5 text-sm hover:bg-stone-50">
                      <span className="flex items-center gap-2">
                        <CalendarDays size={14} className="text-amber-600" />
                        <span className="font-mono text-slate-500">{fmtFecha(f.fecha)}</span>
                        <span className="text-slate-800">{f.descripcion}</span>
                      </span>
                      <button onClick={() => eliminarFeriado(f)} className="text-slate-400 hover:text-rose-600"><X size={14} /></button>
                    </div>
                  ))}
                </div>
              )}
              <div className="mt-3 text-[11px] text-slate-400">
                Los días marcados acá quedan excluidos de las alarmas de auditoría semanal (cierre/apertura) de todas las obras.
              </div>
            </Panel>

            <Panel title="Ficha horaria por obra">
              <div className="space-y-3">
                {obras.filter((o) => o.estado !== "Papelera").map((o) => (
                  <div key={o.id} className="rounded-lg border border-stone-200 p-4">
                    {editandoHorarioObraId === o.id ? (
                      <form className="space-y-3" onSubmit={guardarHorarioObra}>
                        <div className="font-semibold text-slate-900">{o.nombre}</div>
                        <div>
                          <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Días laborables</div>
                          <div className="flex flex-wrap gap-2">
                            {DIAS_SEMANA.map((d) => (
                              <button
                                type="button"
                                key={d}
                                onClick={() => toggleDiaLaborable(d)}
                                className={`rounded-full border px-3 py-1 text-xs font-medium ${
                                  horarioForm.diasLaborables.includes(d) ? "border-amber-500 bg-amber-50 text-amber-800" : "border-stone-300 bg-white text-slate-600 hover:bg-stone-50"
                                }`}
                              >
                                {d}
                              </button>
                            ))}
                          </div>
                        </div>
                        <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
                          <Field label="Hora de apertura">
                            <input type="time" value={horarioForm.horaApertura} onChange={(e) => setHorarioForm((f) => ({ ...f, horaApertura: e.target.value }))} className={inputCls} />
                          </Field>
                          <Field label="Día de cierre semanal">
                            <select value={horarioForm.diaCierre} onChange={(e) => setHorarioForm((f) => ({ ...f, diaCierre: e.target.value }))} className={inputCls}>
                              {DIAS_SEMANA.map((d) => <option key={d}>{d}</option>)}
                            </select>
                          </Field>
                          <Field label="Hora de cierre">
                            <input type="time" value={horarioForm.horaCierre} onChange={(e) => setHorarioForm((f) => ({ ...f, horaCierre: e.target.value }))} className={inputCls} />
                          </Field>
                        </div>
                        <div className="flex gap-2">
                          <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar horario</button>
                          <button type="button" onClick={() => setEditandoHorarioObraId(null)} className={btnGhost}>Cancelar</button>
                        </div>
                      </form>
                    ) : (
                      <div className="flex flex-wrap items-center justify-between gap-3">
                        <div>
                          <div className="font-semibold text-slate-900">{o.nombre}</div>
                          <div className="mt-1 text-sm text-slate-600">
                            {(o.diasLaborables || []).join(", ") || <span className="text-slate-400">Sin días configurados</span>}
                          </div>
                          <div className="text-xs text-slate-400">
                            Apertura {o.horaApertura || "—"} · Cierre {o.diaCierre || "—"} {o.horaCierre || ""}
                          </div>
                        </div>
                        <button onClick={() => abrirEditarHorario(o)} className={btnGhost}>
                          <span className="flex items-center gap-1"><Pencil size={13} /> Editar horario</span>
                        </button>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </Panel>
          </div>
        )}

      </main>
    </div>
  );
}
