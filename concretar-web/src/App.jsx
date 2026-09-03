import { useState, useEffect, useRef, Fragment } from "react";
import { createClient } from "@supabase/supabase-js";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Building2, Users, ClipboardCheck, Wrench,
  ShoppingCart, Receipt, Plus, MapPin, TrendingUp, X, AlertTriangle, CheckCircle2,
  Database, Loader2, RefreshCw, DollarSign, Check, Menu, FileDown, ShieldCheck, Shield,
  Printer, HardHat, Zap, PaintRoller, Droplet, Hammer, Flame, Wallet,
  Landmark, Smartphone, Banknote, Briefcase, Info, Pencil, Truck, ArrowRightLeft, CalendarDays, CalendarClock, Package, Upload, FileSpreadsheet, Trash2, Camera, ChevronLeft, ChevronRight
} from "lucide-react";

// Paleta oficial del Manual de Marca (Grupo Concretar S.A.S)
const BRAND = {
  navy900: "#021d34", // Tono principal
  navy700: "#153f59",
  navy400: "#3a5c66",
  font: "'Poppins', ui-sans-serif, system-ui, sans-serif",
};

const fmtARS = (n) =>
  new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n || 0);

// Color propio de cada obra: se asigna una vez (al crearla) rotando esta
// paleta, y de ahí en más viaja con la obra — se usa en la tarjeta de Obras,
// en Planificación, y para identificarla de un vistazo en Gastos y Facturas
// y en Cuentas. Elegidos para no pisar los colores de estado (ámbar/verde/
// rojo/gris) que ya se usan en toda la app.
const PALETA_OBRA = ["#8b5cf6", "#14b8a6", "#d946ef", "#f97316", "#84cc16", "#ec4899", "#06b6d4", "#a855f7"];
// Lo que no tiene obra asignada ("General") usa este gris fijo, para que se
// lea como "sin obra" y no se confunda con una obra más.
const COLOR_GENERAL = "#78716c";
function colorDeObra(obra) {
  if (!obra) return COLOR_GENERAL;
  return obra.color || PALETA_OBRA[obra.id % PALETA_OBRA.length];
}

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

// Selector de mes en español fijo (Enero, Febrero, ...) en vez del <input type="month">
// nativo, cuyo idioma depende de la configuración del celular/navegador y a veces
// termina mostrando el mes en inglés ("September" en vez de "Septiembre").
function shiftMes(clave, delta) {
  const [y, m] = clave.split("-").map(Number);
  const d = new Date(y, m - 1 + delta, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
function nombreMesClave(clave) {
  const [y, m] = clave.split("-").map(Number);
  const nombre = new Date(y, m - 1, 1).toLocaleDateString("es-AR", { month: "long", year: "numeric" });
  return nombre.charAt(0).toUpperCase() + nombre.slice(1);
}
function MesPicker({ value, onChange, className = "" }) {
  return (
    <div className={`flex items-center gap-1 ${className}`}>
      <button
        type="button"
        onClick={() => onChange(shiftMes(value, -1))}
        className="rounded-md border border-stone-300 bg-white p-1.5 text-slate-600 hover:bg-stone-50"
      >
        <ChevronLeft size={14} />
      </button>
      <span className="min-w-[132px] text-center text-sm font-medium text-slate-700">{nombreMesClave(value)}</span>
      <button
        type="button"
        onClick={() => onChange(shiftMes(value, 1))}
        className="rounded-md border border-stone-300 bg-white p-1.5 text-slate-600 hover:bg-stone-50"
      >
        <ChevronRight size={14} />
      </button>
    </div>
  );
}
function fechaMasDias(dias) {
  const d = new Date();
  d.setDate(d.getDate() + dias);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function nombreComercial(p) {
  return (p?.nombreFantasia && p.nombreFantasia.trim()) || p?.razonSocial || "";
}

const ESTADOS_HERRAMIENTA = ["Disponible", "En Obra", "En Reparación", "Mal Estado", "Rota"];
const ESTADOS_OBRA = ["En curso", "Pendiente de cobro", "Pausada", "Finalizada"];
const ESTADOS_ITEM_COMBO = ["Entregado", "Roto", "Perdido", "Devuelto"];
const TIPOS_CAJA = ["Electricista", "Civil", "Pintor", "Metalúrgico"];
const DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const DIAS_SEMANA_JS = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
const ESTADOS_OC = ["Pendiente", "Requiere aprobación", "Aprobada", "Recibida"];
const ESTADOS_PEDIDO_MATERIAL = ["Solicitado", "Aprobado", "Rechazado", "Facturado", "Recibido"];
const CATEGORIAS_GASTO = ["Materiales", "Equipos y Herramientas", "Epps", "Consumibles", "Combustible"];
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
// Categorías de convenio UOCRA (CCT 76/75) — el resto de CATEGORIAS_PERSONAL es
// personal de estructura (Gerente, RRHH, etc.), no se liquida con este régimen.
const CATEGORIAS_CONVENIO_UOCRA = ["Oficial Especializado", "Oficial", "Medio Oficial", "Ayudante"];
// "En blanco": su liquidación se cierra sí o sí por la calculadora formal UOCRA
// (pasa por Contaduría). "En negro": se le puede pagar en mano desde "Pendientes
// de pago". "Tantero": sigue igual, se liquida por grupo/avance, no por asistencia.
const TIPOS_TRABAJADOR = ["Tantero", "En blanco", "En negro"];
const ASEGURADO_POR = ["No", "ART", "Seg. Accidentes"];
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
const FORMAS_PAGO = ["Banco", "Mercado Pago", "Efectivo", "Cuenta corriente", "eCheq"];
const MEDIOS_BANCARIOS = ["Débito/Transferencia", "Crédito"];
const TIPOS_FACTURA = ["Sin factura", "A", "B", "C"];

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
  Rechazada: "border-rose-600 text-rose-700",
  Facturado: "border-indigo-600 text-indigo-700",
  "Requiere aprobación": "border-rose-600 text-rose-700",
  Aprobada: "border-amber-600 text-amber-700",
  Recibida: "border-emerald-600 text-emerald-700",
  Pagada: "border-emerald-600 text-emerald-700",
  Pagado: "border-emerald-600 text-emerald-700",
  Cobrado: "border-emerald-600 text-emerald-700",
  Blanco: "border-sky-600 text-sky-700",
  Negro: "border-slate-600 text-slate-700",
  "En curso": "border-amber-600 text-amber-700",
  "Pendiente de cobro": "border-sky-600 text-sky-700",
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
// Botón amarillo de "agregar/cargar" — se repetía a mano en cada pestaña.
const btnPrimary = "flex items-center gap-1 rounded-md bg-amber-500 px-3 py-2 text-sm font-semibold text-slate-900 hover:bg-amber-400";

// Campo de plata: mientras escribís va agregando puntos de miles en vivo
// (1 -> 10 -> 100 -> 1.000), como en una caja registradora — y Borrar saca un
// dígito genuino cada vez (antes quedaba "pegado" en el ",00" fijo del final,
// así que Borrar no achicaba el número).
// Funciona tanto con formularios controlados (value+onChange) como con FormData (name).
function MoneyInput({ name, value, onChange, onBlur, className, placeholder, required, disabled }) {
  const [raw, setRaw] = useState(() => {
    if (value === undefined || value === null || value === "") return "";
    const n = Number(value);
    return Number.isInteger(n) ? String(n) : n.toFixed(2).replace(".", ",");
  });
  // "raw" guarda como mucho una coma de centavos, sin separadores de miles, tal cual
  // se va tipeando — así el Backspace siempre borra un carácter real y nunca un
  // ",00" fijo que el usuario no escribió.
  function handleChange(e) {
    // El punto se descarta junto con el resto de la puntuación: es el separador de
    // miles que ya puso el propio formateo, no algo que haya tipeado el usuario.
    let v = e.target.value.replace(/[^\d,]/g, "");
    const primeraComa = v.indexOf(",");
    if (primeraComa !== -1) v = v.slice(0, primeraComa + 1) + v.slice(primeraComa + 1).replace(/,/g, "");
    let [entero, decimales] = v.split(",");
    entero = (entero || "").replace(/^0+(?=\d)/, "");
    if (decimales !== undefined) decimales = decimales.slice(0, 2);
    const nuevoRaw = decimales !== undefined ? `${entero},${decimales}` : (v.includes(",") ? `${entero},` : entero);
    setRaw(nuevoRaw);
    if (onChange) onChange(parseFloat(`${entero || "0"}.${decimales || "0"}`) || 0);
  }
  const [enteroRaw, decimalesRaw] = raw.split(",");
  const num = raw === "" ? 0 : parseFloat(`${enteroRaw || "0"}.${decimalesRaw || "0"}`) || 0;
  const display = raw === "" ? "" : `${new Intl.NumberFormat("es-AR", { maximumFractionDigits: 0 }).format(Number(enteroRaw || "0"))}${raw.includes(",") ? `,${decimalesRaw ?? ""}` : ""}`;
  return (
    <>
      <input
        type="text"
        inputMode="decimal"
        value={display}
        onChange={handleChange}
        onBlur={() => onBlur && onBlur(num)}
        placeholder={placeholder || "0"}
        disabled={disabled}
        className={`${className || inputCls} ${disabled ? "cursor-not-allowed bg-stone-100 text-slate-400" : ""}`}
      />
      {name && <input type="hidden" name={name} value={num} required={required} />}
    </>
  );
}
// Buscador de proveedor: tocás, aparece la lista de proveedores cargados (por nombre
// de fantasía) y filtra a medida que escribís. Si lo que escribiste no coincide con
// ninguno, aparece la opción de darlo de alta al toque sin salir del formulario.
function ProveedorPicker({ name = "proveedor", proveedores, onCrearProveedor }) {
  const [texto, setTexto] = useState("");
  const [seleccionado, setSeleccionado] = useState("");
  const [abierto, setAbierto] = useState(false);
  const [creando, setCreando] = useState(false);
  const cajaRef = useRef(null);

  useEffect(() => {
    function alHacerClickFuera(e) {
      if (cajaRef.current && !cajaRef.current.contains(e.target)) setAbierto(false);
    }
    document.addEventListener("mousedown", alHacerClickFuera);
    return () => document.removeEventListener("mousedown", alHacerClickFuera);
  }, []);

  const busqueda = texto.trim().toLowerCase();
  const opciones = busqueda ? proveedores.filter((p) => nombreComercial(p).toLowerCase().includes(busqueda)) : proveedores;
  // Comparamos ignorando mayúsculas para no obligar a tipear el nombre exacto,
  // pero guardamos el nombre TAL COMO está cargado el proveedor (no lo que
  // escribió el usuario) — si no, una compra podía quedar con otra
  // capitalización y dejar de coincidir con el proveedor real en Cuentas.
  const coincidenciaExacta = proveedores.find((p) => nombreComercial(p).toLowerCase() === busqueda);

  function elegir(p) {
    setTexto(nombreComercial(p));
    setSeleccionado(nombreComercial(p));
    setAbierto(false);
  }

  async function agregarNuevo() {
    const nombreNuevo = texto.trim();
    if (!nombreNuevo || !onCrearProveedor) return;
    setCreando(true);
    const creado = await onCrearProveedor(nombreNuevo);
    setCreando(false);
    if (creado) {
      setTexto(nombreComercial(creado));
      setSeleccionado(nombreComercial(creado));
    }
    setAbierto(false);
  }

  return (
    <div className="relative" ref={cajaRef}>
      <input
        value={texto}
        onChange={(e) => { setTexto(e.target.value); setSeleccionado(""); setAbierto(true); }}
        onFocus={() => setAbierto(true)}
        placeholder="Buscar proveedor…"
        autoComplete="off"
        className={inputCls}
      />
      <input type="hidden" name={name} value={seleccionado || (coincidenciaExacta ? nombreComercial(coincidenciaExacta) : "")} />
      {abierto && (
        <div className="absolute z-20 mt-1 max-h-52 w-full overflow-y-auto rounded-md border border-stone-200 bg-white py-1 shadow-lg">
          {opciones.length === 0 && !busqueda && (
            <div className="px-3 py-2 text-xs text-slate-400">Todavía no hay proveedores cargados.</div>
          )}
          {opciones.map((p) => (
            <button
              type="button"
              key={p.id}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => elegir(p)}
              className="block w-full px-3 py-1.5 text-left text-sm text-slate-700 hover:bg-amber-50"
            >
              {nombreComercial(p)}
            </button>
          ))}
          {busqueda && !coincidenciaExacta && onCrearProveedor && (
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={agregarNuevo}
              disabled={creando}
              className="block w-full border-t border-stone-100 px-3 py-1.5 text-left text-sm font-semibold text-amber-700 hover:bg-amber-50"
            >
              {creando ? "Agregando…" : `+ Agregar "${texto.trim()}" como proveedor nuevo`}
            </button>
          )}
        </div>
      )}
    </div>
  );
}
// Tabla de movimientos de Cuentas (ingresos, egresos y pases entre cuentas),
// reutilizada tanto para el mes actual como para cada mes anterior colapsado.
function TablaMovimientos({ items, obras, onEditar }) {
  if (items.length === 0) {
    return <div className="rounded-lg border border-dashed border-stone-300 bg-white px-3 py-4 text-center text-xs text-slate-400">Todavía no hay movimientos.</div>;
  }
  // Solo gastos/facturas y cobros de socios llevan tipo de factura — el resto
  // (ingresos, transferencias manuales, préstamos, avances) no tiene ese dato.
  const tieneFactura = (origen) => origen === "compras_facturas" || origen === "cobros_socios";
  return (
    <>
      {/* Celular: tarjetas apiladas, sin scroll horizontal. */}
      <div className="space-y-1.5 sm:hidden">
        {items.map((m) => {
          const obra = obras.find((o) => o.id === m.obraId);
          return (
            <div key={m.id} className="rounded-lg border border-stone-200 p-2.5 text-xs shadow-sm" style={{ backgroundColor: `${colorDeObra(obra)}0d` }}>
              <div className="flex items-start justify-between gap-2">
                <span className="font-medium text-slate-900">{m.detalle}</span>
                <span className={`whitespace-nowrap font-mono font-semibold ${m.monto < 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(m.monto)}</span>
              </div>
              <div className="mt-1.5 flex flex-wrap items-center gap-x-2 gap-y-1 text-slate-500">
                <span>{fmtFecha(m.fecha)}</span>
                <span className={`rounded-full border px-1.5 py-0.5 text-[9px] font-bold uppercase tracking-wide ${m.tipo === "Ingreso" ? "border-emerald-300 bg-emerald-50 text-emerald-700" : "border-rose-300 bg-rose-50 text-rose-700"}`}>
                  {m.tipo}
                </span>
                <Badge estado={m.formalidad || "Blanco"} />
                <span className="flex items-center gap-1"><CuentaIcon cuenta={m.cuenta} />{m.cuenta || "—"}</span>
                <span className="flex items-center gap-1"><ObraDot obra={obra} />{obra?.nombre || "General"}</span>
                {m.estado && <Badge estado={m.estado} />}
                {tieneFactura(m.origen) && (
                  <span className={`rounded-full border px-1.5 py-0.5 text-[9px] font-bold uppercase tracking-wide ${(!m.tipoFactura || m.tipoFactura === "Sin factura") ? "border-amber-300 bg-amber-50 text-amber-700" : "border-emerald-300 bg-emerald-50 text-emerald-700"}`}>
                    {(!m.tipoFactura || m.tipoFactura === "Sin factura") ? "S/F" : m.tipoFactura}
                  </span>
                )}
              </div>
              {m.origen && onEditar && (
                <div className="mt-1.5 border-t border-stone-100 pt-1.5 text-right">
                  <button onClick={() => onEditar(m)} className={btnGhost}>
                    <span className="flex items-center gap-1"><Pencil size={12} /> Editar</span>
                  </button>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Tablet/PC: tabla completa. */}
      <div className="hidden overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm sm:block">
        <table className="w-full text-left text-xs">
          <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            <tr>
              <th className="px-2 py-1.5">Fecha</th><th className="px-2 py-1.5">Tipo</th><th className="px-2 py-1.5">Obra</th>
              <th className="px-2 py-1.5">Detalle</th><th className="px-2 py-1.5">Formalidad</th><th className="px-2 py-1.5">Cuenta</th>
              <th className="px-2 py-1.5 text-right">Monto</th><th className="px-2 py-1.5">Estado</th><th className="px-2 py-1.5">Factura</th><th className="px-2 py-1.5"></th>
            </tr>
          </thead>
          <tbody>
            {items.map((m) => {
              const obra = obras.find((o) => o.id === m.obraId);
              return (
                <tr key={m.id} className="border-t border-stone-100" style={{ backgroundColor: `${colorDeObra(obra)}0d` }}>
                  <td className="px-2 py-1 text-slate-600">{fmtFecha(m.fecha)}</td>
                  <td className="px-2 py-1">
                    <span className={`rounded-full border px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide ${m.tipo === "Ingreso" ? "border-emerald-300 bg-emerald-50 text-emerald-700" : "border-rose-300 bg-rose-50 text-rose-700"}`}>
                      {m.tipo}
                    </span>
                  </td>
                  <td className="px-2 py-1 text-slate-600"><span className="flex items-center gap-1.5"><ObraDot obra={obra} />{obra?.nombre || "General"}</span></td>
                  <td className="px-2 py-1 font-medium text-slate-900">{m.detalle}</td>
                  <td className="px-2 py-1"><Badge estado={m.formalidad || "Blanco"} /></td>
                  <td className="px-2 py-1 text-slate-600"><span className="flex items-center gap-1"><CuentaIcon cuenta={m.cuenta} />{m.cuenta || "—"}</span></td>
                  <td className={`px-2 py-1 text-right font-mono font-semibold ${m.monto < 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(m.monto)}</td>
                  <td className="px-2 py-1">{m.estado && <Badge estado={m.estado} />}</td>
                  <td className="px-2 py-1">
                    {tieneFactura(m.origen) && (
                      (!m.tipoFactura || m.tipoFactura === "Sin factura") ? (
                        <span className="rounded-full border border-amber-300 bg-amber-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-amber-700">S/F</span>
                      ) : (
                        <span className="rounded-full border border-emerald-300 bg-emerald-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-emerald-700">{m.tipoFactura}</span>
                      )
                    )}
                  </td>
                  <td className="px-2 py-1">
                    {m.origen && onEditar && (
                      <button onClick={() => onEditar(m)} className={btnGhost}>
                        <span className="flex items-center gap-1"><Pencil size={12} /> Editar</span>
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </>
  );
}

// Historial de gastos de una obra puntual (pestaña Obras → detalle): junta Gastos y
// Facturas con la mano de obra (tanteros, personal en negro pagado y personal en blanco
// ya liquidado), ordenado de más reciente a más antiguo — mismo criterio de "gasto" que
// usa el Balance por obra de Cuentas.
function HistorialGastosObra({ items }) {
  if (items.length === 0) {
    return <div className="rounded-lg border border-dashed border-stone-300 bg-white px-3 py-4 text-center text-xs text-slate-400">Todavía no hay gastos cargados para esta obra.</div>;
  }
  const colorTipo = {
    "Gastos y Facturas": "border-sky-300 bg-sky-50 text-sky-700",
    Tantero: "border-amber-300 bg-amber-50 text-amber-700",
    Personal: "border-slate-300 bg-slate-50 text-slate-700",
    "Personal (blanco)": "border-emerald-300 bg-emerald-50 text-emerald-700",
  };
  return (
    <>
      {/* Celular: tarjetas apiladas. */}
      <div className="space-y-1.5 sm:hidden">
        {items.map((g) => (
          <div key={g.id} className="rounded-lg border border-stone-200 bg-white p-2.5 text-xs shadow-sm">
            <div className="flex items-start justify-between gap-2">
              <span className="font-medium text-slate-900">{g.detalle}</span>
              <span className="whitespace-nowrap font-mono font-semibold text-rose-600">{fmtARS(g.monto)}</span>
            </div>
            <div className="mt-1.5 flex flex-wrap items-center gap-x-2 gap-y-1 text-slate-500">
              <span>{fmtFecha(g.fecha)}</span>
              <span className={`rounded-full border px-1.5 py-0.5 text-[9px] font-bold uppercase tracking-wide ${colorTipo[g.tipo] || "border-slate-300 bg-slate-50 text-slate-700"}`}>{g.tipo}</span>
            </div>
          </div>
        ))}
      </div>

      {/* Tablet/PC: tabla completa. */}
      <div className="hidden overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm sm:block">
        <table className="w-full text-left text-xs">
          <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            <tr>
              <th className="px-2 py-1.5">Fecha</th>
              <th className="px-2 py-1.5">Tipo</th>
              <th className="px-2 py-1.5">Detalle</th>
              <th className="px-2 py-1.5 text-right">Monto</th>
            </tr>
          </thead>
          <tbody>
            {items.map((g) => (
              <tr key={g.id} className="border-t border-stone-100">
                <td className="px-2 py-1 text-slate-600">{fmtFecha(g.fecha)}</td>
                <td className="px-2 py-1">
                  <span className={`rounded-full border px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide ${colorTipo[g.tipo] || "border-slate-300 bg-slate-50 text-slate-700"}`}>{g.tipo}</span>
                </td>
                <td className="px-2 py-1 font-medium text-slate-900">{g.detalle}</td>
                <td className="px-2 py-1 text-right font-mono font-semibold text-rose-600">{fmtARS(g.monto)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

// Resumen de balance por obra en Cuentas: una fila por obra en curso, con lo
// presupuestado (si se importó el Excel de presupuesto), lo cobrado/gastado
// real y la ganancia estimada. Las filas salen solas de "obras" — no hay nada
// hardcodeado, así que a medida que se cargan obras esto se va completando.
function ObraDot({ obra, size = 8 }) {
  return <span className="inline-block shrink-0 rounded-full" style={{ width: size, height: size, backgroundColor: colorDeObra(obra) }} />;
}
function ResumenObrasCuentas({ items }) {
  if (items.length === 0) {
    return <div className="rounded-lg border border-dashed border-stone-300 bg-white px-3 py-4 text-center text-xs text-slate-400">Todavía no hay obras cargadas.</div>;
  }
  const celda = (v) => (v === null || v === undefined ? "—" : fmtARS(v));
  return (
    <>
      {/* Celular: una tarjeta por obra, mismo formato que "Balance de la obra" en Obras. */}
      <div className="space-y-2 sm:hidden">
        {items.map((r) => (
          <div key={r.obra.id} className="rounded-lg border border-stone-200 bg-white p-3 text-xs shadow-sm" style={{ backgroundColor: `${colorDeObra(r.obra)}0d` }}>
            <div className="flex items-center justify-between gap-2">
              <span className="flex items-center gap-1.5 font-semibold text-slate-900"><ObraDot obra={r.obra} />{r.obra.nombre}</span>
              <Badge estado={r.obra.estado} />
            </div>
            <div className="mt-2 grid grid-cols-2 gap-x-3 gap-y-1.5">
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Precio de obra</div><div className="font-mono font-semibold text-slate-800">{celda(r.precioObra)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Falta cobrar</div><div className="font-mono font-semibold text-slate-800">{celda(r.faltaCobrar)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Presup. M.O.</div><div className="font-mono text-slate-700">{celda(r.presupuestadoManoObra)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Pagado M.O.</div><div className="font-mono text-slate-700">{celda(r.gastadoManoObra)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Presup. Eq. y Mat.</div><div className="font-mono text-slate-700">{celda(r.presupuestadoEqYMat)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Gastado Eq. y Mat.</div><div className="font-mono text-slate-700">{celda(r.gastadoEqYMat)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Gastado total</div><div className="font-mono text-slate-700">{celda(r.gastado)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Dinero en caja</div><div className={`font-mono font-semibold ${r.dineroEnCaja < 0 ? "text-rose-600" : "text-emerald-700"}`}>{celda(r.dineroEnCaja)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Ganancia estimada</div><div className={`font-mono font-semibold ${r.gananciaEstimada !== null && r.gananciaEstimada < 0 ? "text-rose-600" : "text-emerald-700"}`}>{celda(r.gananciaEstimada)}</div></div>
              <div><div className="text-[10px] uppercase tracking-wide text-slate-400">% Ganancia</div><div className="font-mono text-slate-700">{r.porcentajeGanancia === null ? "—" : `${Math.round(r.porcentajeGanancia * 100)}%`}</div></div>
            </div>
          </div>
        ))}
      </div>

      {/* Tablet/PC: tabla completa, mismos campos que la tarjeta de celular. */}
      <div className="hidden overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm sm:block">
        <table className="w-full text-left text-xs">
          <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            <tr>
              <th className="px-2 py-1.5">Obra</th>
              <th className="px-2 py-1.5 text-right">Precio de obra</th>
              <th className="px-2 py-1.5 text-right">Falta cobrar</th>
              <th className="px-2 py-1.5 text-right">Presup. M.O.</th>
              <th className="px-2 py-1.5 text-right">Pagado M.O.</th>
              <th className="px-2 py-1.5 text-right">Presup. Eq. y Mat.</th>
              <th className="px-2 py-1.5 text-right">Gastado Eq. y Mat.</th>
              <th className="px-2 py-1.5 text-right">Gastado total</th>
              <th className="px-2 py-1.5 text-right">Dinero en caja</th>
              <th className="px-2 py-1.5 text-right">Ganancia estimada</th>
              <th className="px-2 py-1.5 text-right">% Ganancia</th>
              <th className="px-2 py-1.5">Estado</th>
            </tr>
          </thead>
          <tbody>
            {items.map((r) => (
              <tr key={r.obra.id} className="border-t border-stone-100" style={{ backgroundColor: `${colorDeObra(r.obra)}0d` }}>
                <td className="px-2 py-1 font-medium text-slate-900"><span className="flex items-center gap-1.5"><ObraDot obra={r.obra} />{r.obra.nombre}</span></td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{celda(r.precioObra)}</td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{celda(r.faltaCobrar)}</td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{celda(r.presupuestadoManoObra)}</td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{celda(r.gastadoManoObra)}</td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{celda(r.presupuestadoEqYMat)}</td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{celda(r.gastadoEqYMat)}</td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{celda(r.gastado)}</td>
                <td className={`px-2 py-1 text-right font-mono font-semibold ${r.dineroEnCaja < 0 ? "text-rose-600" : "text-emerald-700"}`}>{celda(r.dineroEnCaja)}</td>
                <td className={`px-2 py-1 text-right font-mono font-semibold ${r.gananciaEstimada !== null && r.gananciaEstimada < 0 ? "text-rose-600" : "text-emerald-700"}`}>{celda(r.gananciaEstimada)}</td>
                <td className="px-2 py-1 text-right font-mono text-slate-700">{r.porcentajeGanancia === null ? "—" : `${Math.round(r.porcentajeGanancia * 100)}%`}</td>
                <td className="px-2 py-1"><Badge estado={r.obra.estado} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

// ---------- Planificación (Gantt de etapas por obra) en la pestaña Obras ----------
function estadoEtapa(etapa, hoy) {
  const avance = etapa.avance || 0;
  const fin = fechaLocal(etapa.fin);
  const inicio = fechaLocal(etapa.inicio);
  if (avance >= 100) return "Finalizada";
  if (fin && fin < hoy) return "Atrasada";
  if (inicio && inicio <= hoy) return "En curso";
  return "Pendiente";
}
const ETAPA_COLOR_BARRA = { Finalizada: "#10b981", "En curso": "#f59e0b", Atrasada: "#f43f5e", Pendiente: "#d6d3d1" };
const ETAPA_COLOR_TRACK = { Finalizada: "#d1fae5", "En curso": "#fef3c7", Atrasada: "#fecdd3", Pendiente: "#f5f5f4" };
const MESES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
function lunesOAntes(d) {
  const dia = d.getDay();
  const r = new Date(d);
  r.setDate(r.getDate() - (dia === 0 ? 6 : dia - 1));
  r.setHours(0, 0, 0, 0);
  return r;
}
function domingoODespues(d) {
  const dia = d.getDay();
  const r = new Date(d);
  r.setDate(r.getDate() + (dia === 0 ? 0 : 7 - dia));
  r.setHours(23, 59, 59, 999);
  return r;
}

function FormEtapa({ inicial, onGuardar, onCancelar }) {
  const [nombre, setNombre] = useState(inicial?.nombre || "");
  const [inicio, setInicio] = useState(inicial?.inicio || hoyISO());
  const [fin, setFin] = useState(inicial?.fin || hoyISO());
  const [avance, setAvance] = useState(inicial?.avance ?? 0);
  return (
    <form
      onSubmit={(e) => { e.preventDefault(); if (!nombre.trim()) return; onGuardar({ nombre: nombre.trim(), inicio, fin, avance }); }}
      className="grid grid-cols-1 gap-3 rounded-md border border-dashed border-stone-300 bg-stone-50 p-3 sm:grid-cols-5"
    >
      <div className="sm:col-span-2">
        <Field label="Nombre de la etapa"><input value={nombre} onChange={(e) => setNombre(e.target.value)} required className={inputCls} /></Field>
      </div>
      <Field label="Inicio"><input type="date" value={inicio} onChange={(e) => setInicio(e.target.value)} required className={inputCls} /></Field>
      <Field label="Fin"><input type="date" value={fin} onChange={(e) => setFin(e.target.value)} required className={inputCls} /></Field>
      <Field label="Avance (%)"><input type="number" min="0" max="100" value={avance} onChange={(e) => setAvance(e.target.value)} className={inputCls} /></Field>
      <div className="flex items-end gap-2 sm:col-span-5">
        <button className="rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">Guardar</button>
        <button type="button" onClick={onCancelar} className={btnGhost}>Cancelar</button>
      </div>
    </form>
  );
}

// Vista de Planificación de la pestaña Obras: todas las obras juntas, con sus
// etapas cargadas a mano en un Gantt semanal (lunes a domingo). Todavía no
// depende de nada más — cuando conectemos más adelante un botón "Ver
// planificación" desde el detalle de cada obra, esto ya queda listo.
function PlanificacionObras({ obras, etapas, agregandoEtapaObraId, setAgregandoEtapaObraId, editandoEtapaId, setEditandoEtapaId, onAgregarEtapa, onGuardarEdicionEtapa, onEliminarEtapa }) {
  const hoy = fechaLocal(hoyISO());
  const fechas = etapas.flatMap((e) => [fechaLocal(e.inicio), fechaLocal(e.fin)]).filter(Boolean);
  const minFecha = fechas.length ? new Date(Math.min(...fechas)) : hoy;
  const maxFecha = fechas.length ? new Date(Math.max(...fechas)) : hoy;
  const weekStart = lunesOAntes(new Date(Math.min(minFecha.getTime(), hoy.getTime() - 14 * 86400000)));
  const weekEnd = domingoODespues(new Date(Math.max(maxFecha.getTime(), hoy.getTime() + 30 * 86400000)));
  const totalMs = weekEnd - weekStart;
  const pct = (d) => ((d - weekStart) / totalMs) * 100;
  const semanas = [];
  for (let cur = new Date(weekStart); cur <= weekEnd; cur.setDate(cur.getDate() + 7)) semanas.push(new Date(cur));
  const anchoSemana = 100 / semanas.length;
  const mesDeSemana = (w) => new Date(w.getTime() + 3 * 86400000).getMonth();
  const bandas = [];
  {
    let i = 0;
    while (i < semanas.length) {
      const m = mesDeSemana(semanas[i]);
      let j = i;
      while (j < semanas.length && mesDeSemana(semanas[j]) === m) j++;
      bandas.push({ mes: m, izq: i * anchoSemana, ancho: (j - i) * anchoSemana });
      i = j;
    }
  }
  const minWidthPx = Math.max(720, Math.round(semanas.length * 34 + 220));

  const etapasPorObra = new Map();
  etapas.forEach((e) => {
    if (!etapasPorObra.has(e.obraId)) etapasPorObra.set(e.obraId, []);
    etapasPorObra.get(e.obraId).push(e);
  });

  return (
    <Panel title="Planificación — todas las obras">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
        <div className="flex flex-wrap gap-4 text-xs text-slate-600">
          {["Finalizada", "En curso", "Atrasada", "Pendiente"].map((estado) => (
            <span key={estado} className="flex items-center gap-1.5">
              <span className="inline-block h-2 w-2 rounded-full" style={{ backgroundColor: ETAPA_COLOR_BARRA[estado] }} />
              {estado}
            </span>
          ))}
        </div>
        <div className="text-[11px] text-slate-400">Semanas de lunes a domingo — el número es el día del mes en que arranca cada semana.</div>
      </div>

      {obras.length === 0 ? (
        <div className="rounded-lg border border-dashed border-stone-300 bg-white px-3 py-6 text-center text-xs text-slate-400">Todavía no hay obras cargadas.</div>
      ) : (
        <div className="overflow-x-auto">
          <div style={{ minWidth: minWidthPx }}>
            <div className="flex">
              <div className="w-[220px] shrink-0" />
              <div className="relative h-6 flex-1">
                {bandas.map((b, i) => (
                  <div
                    key={i}
                    className="absolute top-0 bottom-0 border-l border-stone-200 pl-1.5 text-[10px] font-semibold uppercase tracking-wide text-slate-400"
                    style={{ left: `${b.izq}%`, width: `${b.ancho}%` }}
                  >
                    {MESES_CORTO[b.mes]}
                  </div>
                ))}
              </div>
            </div>
            <div className="flex">
              <div className="w-[220px] shrink-0" />
              <div className="relative h-6 flex-1 border-b border-stone-200">
                {semanas.map((s, i) => {
                  const esHoy = hoy >= s && hoy < new Date(s.getTime() + 7 * 86400000);
                  return (
                    <div
                      key={i}
                      className={`absolute top-0 bottom-0 flex items-center justify-center border-l border-stone-100 text-[10px] font-medium ${esHoy ? "bg-rose-50 font-bold text-rose-600" : "text-slate-400"}`}
                      style={{ left: `${i * anchoSemana}%`, width: `${anchoSemana}%` }}
                    >
                      {String(s.getDate()).padStart(2, "0")}
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="relative mt-1">
              <div className="pointer-events-none absolute inset-y-0" style={{ left: 220, right: 0 }}>
                <div className="absolute top-0 bottom-0 border-l-2 border-dashed border-rose-500" style={{ left: `${pct(hoy)}%` }} />
              </div>

              {obras.map((obra) => {
                const etapasDeObra = etapasPorObra.get(obra.id) || [];
                return (
                  <div key={obra.id}>
                    <div className="flex items-center gap-3 py-2" style={{ backgroundColor: `${colorDeObra(obra)}17` }}>
                      <div className="flex w-[220px] shrink-0 items-center gap-2 pl-2 pr-2">
                        <span className="h-full min-h-[1.5rem] w-1 self-stretch rounded" style={{ backgroundColor: colorDeObra(obra) }} />
                        <div className="min-w-0">
                          <div className="truncate text-sm font-bold text-slate-900">{obra.nombre}</div>
                          <div className="text-[11px] text-slate-400">{etapasDeObra.length} etapa{etapasDeObra.length === 1 ? "" : "s"}</div>
                        </div>
                      </div>
                      <div className="flex-1 pr-2 text-right">
                        <button
                          onClick={() => setAgregandoEtapaObraId(agregandoEtapaObraId === obra.id ? null : obra.id)}
                          className={btnGhost}
                        >
                          <span className="flex items-center gap-1"><Plus size={12} /> Agregar etapa</span>
                        </button>
                      </div>
                    </div>

                    {etapasDeObra.map((etapa) => {
                      if (editandoEtapaId === etapa.id) {
                        return (
                          <div key={etapa.id} className="py-2 pl-2 pr-2">
                            <FormEtapa inicial={etapa} onGuardar={(datos) => onGuardarEdicionEtapa(etapa, datos)} onCancelar={() => setEditandoEtapaId(null)} />
                          </div>
                        );
                      }
                      const estado = estadoEtapa(etapa, hoy);
                      const left = pct(fechaLocal(etapa.inicio));
                      const width = Math.max(pct(fechaLocal(etapa.fin)) - left, 0.5);
                      return (
                        <div key={etapa.id} className="flex items-center gap-3 border-t border-stone-100 py-2">
                          <div className="flex w-[220px] shrink-0 items-start justify-between gap-1 pl-4 pr-2">
                            <div className="min-w-0">
                              <div className="flex items-start gap-1 text-xs font-semibold text-slate-800">
                                {estado === "Atrasada" && <AlertTriangle size={11} className="mt-0.5 shrink-0 text-rose-500" />}
                                <span>{etapa.nombre}</span>
                              </div>
                              <div className="text-[10px] text-slate-400">{fmtFecha(etapa.inicio)} → {fmtFecha(etapa.fin)} · {etapa.avance || 0}%</div>
                            </div>
                            <div className="flex shrink-0 items-center gap-0.5 text-slate-400">
                              <button onClick={() => setEditandoEtapaId(etapa.id)} title="Editar etapa" className="rounded p-1 hover:bg-stone-100 hover:text-slate-600"><Pencil size={12} /></button>
                              <button onClick={() => onEliminarEtapa(etapa)} title="Eliminar etapa" className="rounded p-1 hover:bg-stone-100 hover:text-rose-500"><Trash2 size={12} /></button>
                            </div>
                          </div>
                          <div className="relative h-5 flex-1">
                            <div className="absolute top-0 h-5 rounded" style={{ left: `${left}%`, width: `${width}%`, backgroundColor: ETAPA_COLOR_TRACK[estado] }}>
                              <div className="h-full rounded" style={{ width: `${Math.min(etapa.avance || 0, 100)}%`, backgroundColor: ETAPA_COLOR_BARRA[estado] }} />
                            </div>
                          </div>
                        </div>
                      );
                    })}

                    {agregandoEtapaObraId === obra.id && (
                      <div className="py-2 pl-2 pr-2">
                        <FormEtapa onGuardar={(datos) => onAgregarEtapa(obra.id, datos)} onCancelar={() => setAgregandoEtapaObraId(null)} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </Panel>
  );
}

// Préstamos de inversores/banco: el interés corre día a día a la tasa anual
// pactada, desde que se recibió hasta hoy (o hasta la fecha de devolución si ya
// se pagó) — son funciones puras así las puede usar también la tabla de abajo.
function diasTranscurridosDesde(fechaStr, hastaStr) {
  return Math.max(0, Math.round((fechaLocal(hastaStr) - fechaLocal(fechaStr)) / 86400000));
}
// Recalcula el estado de un préstamo pisando los pagos parciales que se hayan
// hecho: cada pago primero cubre el interés acumulado desde el pago anterior
// (o desde el alta) y lo que sobra amortiza capital — así el interés de ahí en
// adelante corre sobre el saldo de capital que quede, no sobre el capital original.
function calcularEstadoPrestamo(p, pagos, hastaOverride) {
  const tasa = (p.tasaAnualPct || 0) / 100;
  const pagosDelPrestamo = pagos.filter((pg) => pg.prestamoId === p.id).sort((a, b) => fechaLocal(a.fecha) - fechaLocal(b.fecha));
  let saldoCapital = p.capital || 0;
  let fechaCorte = p.fecha;
  for (const pago of pagosDelPrestamo) {
    const dias = diasTranscurridosDesde(fechaCorte, pago.fecha);
    const interesDelPeriodo = saldoCapital * tasa * (dias / 365);
    const aCapital = Math.max(0, (pago.monto || 0) - interesDelPeriodo);
    saldoCapital = Math.max(0, saldoCapital - aCapital);
    fechaCorte = pago.fecha;
  }
  // hastaOverride sirve para proyectar cuánto habría que devolver en una fecha
  // futura (ej: la fecha estimada de devolución), sin tocar el cálculo normal
  // "a hoy" que usa el resto de la app.
  const hasta = hastaOverride || (p.estado === "Pagado" ? (p.fechaPago || fechaCorte) : hoyISO());
  const interesAcumulado = saldoCapital * tasa * (diasTranscurridosDesde(fechaCorte, hasta) / 365);
  const totalPagado = pagosDelPrestamo.reduce((s, pg) => s + (pg.monto || 0), 0);
  return { saldoCapital, interesAcumulado, totalADevolver: saldoCapital + interesAcumulado, totalPagado, pagos: pagosDelPrestamo, fechaCorte, dias: diasTranscurridosDesde(fechaCorte, hasta) };
}

function TablaPrestamos({ items, pagos, onEditar, onRegistrarPago, onEliminar }) {
  if (items.length === 0) {
    return <div className="rounded-lg border border-dashed border-stone-300 bg-white px-3 py-4 text-center text-xs text-slate-400">Todavía no hay préstamos cargados.</div>;
  }
  return (
    <>
      {/* Celular: una tarjeta por préstamo. */}
      <div className="space-y-2 sm:hidden">
        {items.map((p) => {
          const estado = calcularEstadoPrestamo(p, pagos);
          const proyeccion = (p.estado !== "Pagado" && p.fechaEstimadaDevolucion)
            ? calcularEstadoPrestamo(p, pagos, p.fechaEstimadaDevolucion).totalADevolver
            : null;
          return (
            <div key={p.id} className="rounded-lg border border-stone-200 bg-white p-3 text-xs shadow-sm">
              <div className="flex items-center justify-between gap-2">
                <span className="font-semibold text-slate-900">{p.acreedor}</span>
                <Badge estado={p.estado} />
              </div>
              <div className="mt-2 grid grid-cols-2 gap-x-3 gap-y-1.5">
                <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Capital original</div><div className="font-mono font-semibold text-slate-800">{fmtARS(p.capital)}</div></div>
                <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Tasa anual</div><div className="font-mono text-slate-700">{p.tasaAnualPct}%</div></div>
                {estado.totalPagado > 0 && (
                  <>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Pagado hasta ahora</div><div className="font-mono text-emerald-700">{fmtARS(estado.totalPagado)}</div></div>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Saldo de capital</div><div className="font-mono text-slate-700">{fmtARS(estado.saldoCapital)}</div></div>
                  </>
                )}
                <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Días {estado.totalPagado > 0 ? "desde el último pago" : "transcurridos"}</div><div className="font-mono text-slate-700">{estado.dias}</div></div>
                <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Interés acumulado</div><div className="font-mono text-amber-700">{fmtARS(estado.interesAcumulado)}</div></div>
                <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Total a devolver</div><div className="font-mono font-semibold text-rose-600">{fmtARS(estado.totalADevolver)}</div></div>
                <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Fecha estimada</div><div className="text-slate-700">{fmtFecha(p.fechaEstimadaDevolucion)}</div></div>
                {proyeccion !== null && (
                  <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Monto a devolver en fecha estimada</div><div className="font-mono font-semibold text-slate-800">{fmtARS(proyeccion)}</div></div>
                )}
              </div>
              {estado.pagos.length > 0 && (
                <div className="mt-2 border-t border-stone-100 pt-2">
                  <div className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-slate-400">Pagos registrados</div>
                  <div className="space-y-1">
                    {estado.pagos.map((pg) => (
                      <div key={pg.id} className="flex items-center justify-between text-[11px] text-slate-600">
                        <span>{fmtFecha(pg.fecha)} · {pg.cuenta}</span>
                        <span className="font-mono">{fmtARS(pg.monto)}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              <div className="mt-2 flex flex-wrap items-center justify-end gap-1 border-t border-stone-100 pt-2">
                {p.estado !== "Pagado" && (
                  <button onClick={() => onRegistrarPago(p)} className={btnGhost}>Registrar pago</button>
                )}
                <button type="button" onClick={() => onEditar(p)} title="Editar préstamo" className="rounded-md border border-transparent p-1 text-slate-400 hover:border-stone-300 hover:bg-stone-100 hover:text-slate-700">
                  <Pencil size={14} />
                </button>
                <BotonEliminar onClick={() => onEliminar(p)} title="Eliminar préstamo" />
              </div>
            </div>
          );
        })}
      </div>

      {/* Tablet/PC: tabla completa. */}
      <div className="hidden overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm sm:block">
        <table className="w-full text-left text-xs">
          <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            <tr>
              <th className="px-2 py-1.5">Acreedor</th>
              <th className="px-2 py-1.5 text-right">Capital original</th>
              <th className="px-2 py-1.5 text-right">Pagado</th>
              <th className="px-2 py-1.5 text-right">Saldo capital</th>
              <th className="px-2 py-1.5 text-right">Tasa anual</th>
              <th className="px-2 py-1.5 text-right">Días</th>
              <th className="px-2 py-1.5 text-right">Interés acumulado</th>
              <th className="px-2 py-1.5 text-right">Total a devolver</th>
              <th className="px-2 py-1.5">Fecha estimada</th>
              <th className="px-2 py-1.5 text-right">Monto a devolver en fecha estimada</th>
              <th className="px-2 py-1.5">Estado</th>
              <th className="px-2 py-1.5"></th>
            </tr>
          </thead>
          <tbody>
            {items.map((p) => {
              const estado = calcularEstadoPrestamo(p, pagos);
              const proyeccion = (p.estado !== "Pagado" && p.fechaEstimadaDevolucion)
                ? calcularEstadoPrestamo(p, pagos, p.fechaEstimadaDevolucion).totalADevolver
                : null;
              return (
                <tr key={p.id} className="border-t border-stone-100">
                  <td className="px-2 py-1 font-medium text-slate-900">
                    {p.acreedor}
                    {estado.pagos.length > 0 && <div className="font-normal text-slate-400">{estado.pagos.length} pago{estado.pagos.length > 1 ? "s" : ""} registrado{estado.pagos.length > 1 ? "s" : ""}</div>}
                  </td>
                  <td className="px-2 py-1 text-right font-mono text-slate-700">{fmtARS(p.capital)}</td>
                  <td className="px-2 py-1 text-right font-mono text-emerald-700">{estado.totalPagado > 0 ? fmtARS(estado.totalPagado) : "—"}</td>
                  <td className="px-2 py-1 text-right font-mono text-slate-700">{fmtARS(estado.saldoCapital)}</td>
                  <td className="px-2 py-1 text-right font-mono text-slate-700">{p.tasaAnualPct}%</td>
                  <td className="px-2 py-1 text-right font-mono text-slate-700">{estado.dias}</td>
                  <td className="px-2 py-1 text-right font-mono text-amber-700">{fmtARS(estado.interesAcumulado)}</td>
                  <td className="px-2 py-1 text-right font-mono font-semibold text-rose-600">{fmtARS(estado.totalADevolver)}</td>
                  <td className="px-2 py-1 text-slate-600">{fmtFecha(p.fechaEstimadaDevolucion)}</td>
                  <td className="px-2 py-1 text-right font-mono font-semibold text-slate-800">{proyeccion !== null ? fmtARS(proyeccion) : "—"}</td>
                  <td className="px-2 py-1"><Badge estado={p.estado} /></td>
                  <td className="px-2 py-1">
                    <div className="flex flex-nowrap items-center justify-end gap-1">
                      {p.estado !== "Pagado" && (
                        <button onClick={() => onRegistrarPago(p)} className={`${btnGhost} whitespace-nowrap`}>Registrar pago</button>
                      )}
                      <button type="button" onClick={() => onEditar(p)} title="Editar préstamo" className="rounded-md border border-transparent p-1 text-slate-400 hover:border-stone-300 hover:bg-stone-100 hover:text-slate-700">
                        <Pencil size={14} />
                      </button>
                      <BotonEliminar onClick={() => onEliminar(p)} title="Eliminar préstamo" />
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </>
  );
}

// Modal para corregir un préstamo ya cargado (fecha, acreedor, capital, tasa,
// cuenta, formalidad, fecha estimada de devolución) por si se cometió un error al alta.
function ModalEditarPrestamo({ prestamo, onClose, onGuardar }) {
  const [form, setForm] = useState(prestamo || {});
  if (!prestamo) return null;

  function guardar(e) {
    e.preventDefault();
    onGuardar(prestamo.id, {
      fecha: form.fecha,
      acreedor: form.acreedor,
      capital: Number(form.capital) || 0,
      tasaAnualPct: Number(form.tasaAnualPct) || 0,
      cuenta: form.cuenta,
      formalidad: form.formalidad,
      fechaEstimadaDevolucion: form.fechaEstimadaDevolucion || null,
    });
  }

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto bg-black/50 p-4 sm:items-center" onClick={onClose}>
      <div className="w-full max-w-2xl rounded-lg bg-white p-5 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-900">Editar préstamo</h3>
          <button onClick={onClose}><X size={18} /></button>
        </div>
        <form onSubmit={guardar} className="grid grid-cols-1 gap-4 md:grid-cols-3">
          <Field label="Fecha">
            <input type="date" value={form.fecha} onChange={(e) => setForm((f) => ({ ...f, fecha: e.target.value }))} required className={inputCls} />
          </Field>
          <Field label="Acreedor (inversor / banco)">
            <input value={form.acreedor} onChange={(e) => setForm((f) => ({ ...f, acreedor: e.target.value }))} required className={inputCls} />
          </Field>
          <Field label="Capital ($)">
            <MoneyInput value={form.capital} onChange={(v) => setForm((f) => ({ ...f, capital: v }))} className={inputCls} />
          </Field>
          <Field label="Tasa anual (%)">
            <input
              type="number" min="0" step="0.01"
              value={form.tasaAnualPct}
              onChange={(e) => setForm((f) => ({ ...f, tasaAnualPct: e.target.value }))}
              required
              className={inputCls}
            />
          </Field>
          <Field label="Cuenta donde entró">
            <select value={form.cuenta} onChange={(e) => setForm((f) => ({ ...f, cuenta: e.target.value }))} className={inputCls}>
              {CUENTAS.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="Formalidad">
            <select value={form.formalidad} onChange={(e) => setForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
              {FORMALIDADES.map((f) => <option key={f}>{f}</option>)}
            </select>
          </Field>
          <Field label="Fecha estimada de devolución">
            <input type="date" value={form.fechaEstimadaDevolucion || ""} onChange={(e) => setForm((f) => ({ ...f, fechaEstimadaDevolucion: e.target.value }))} className={inputCls} />
          </Field>
          <div className="flex items-end gap-2 md:col-span-3">
            <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button>
            <button type="button" onClick={onClose} className={btnGhost}>Cancelar</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// Modal para registrar una devolución parcial de un préstamo. El pago cubre
// primero el interés acumulado hasta esa fecha y lo que sobra amortiza
// capital — de ahí en más el interés corre sobre el saldo que quede.
function ModalPagoPrestamo({ prestamo, pagos, onClose, onGuardar }) {
  const [fecha, setFecha] = useState(hoyISO());
  const [monto, setMonto] = useState(0);
  const [cuenta, setCuenta] = useState(prestamo?.cuenta);
  // MoneyInput solo lee su prop "value" al montarse (para no pelearle al usuario
  // mientras tipea) — cambiar esta key fuerza que lo tome de nuevo cuando lo
  // llenamos nosotros con el botón de acceso rápido, en vez de tipeándolo.
  const [montoResetKey, setMontoResetKey] = useState(0);
  if (!prestamo) return null;

  const estadoActual = calcularEstadoPrestamo(prestamo, pagos);
  const dias = diasTranscurridosDesde(estadoActual.fechaCorte, fecha);
  const interesAlPagar = estadoActual.saldoCapital * ((prestamo.tasaAnualPct || 0) / 100) * (dias / 365);
  const totalAEstaFecha = estadoActual.saldoCapital + interesAlPagar;
  const aCapital = Math.max(0, (Number(monto) || 0) - interesAlPagar);
  const saldoRestante = Math.max(0, estadoActual.saldoCapital - aCapital);

  function guardar(e) {
    e.preventDefault();
    onGuardar(prestamo.id, { fecha, monto: Number(monto) || 0, cuenta });
  }

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto bg-black/50 p-4 sm:items-center" onClick={onClose}>
      <div className="w-full max-w-lg rounded-lg bg-white p-5 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-900">Registrar pago — {prestamo.acreedor}</h3>
          <button onClick={onClose}><X size={18} /></button>
        </div>
        <div className="mb-3 text-xs text-slate-500">
          Saldo de capital actual: <span className="font-mono font-semibold text-slate-700">{fmtARS(estadoActual.saldoCapital)}</span>
        </div>
        <form onSubmit={guardar} className="grid grid-cols-1 gap-4 md:grid-cols-2">
          <Field label="Fecha del pago">
            <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} required className={inputCls} />
          </Field>
          <Field label="Cuenta de la que sale">
            <select value={cuenta} onChange={(e) => setCuenta(e.target.value)} className={inputCls}>
              {CUENTAS.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <div className="md:col-span-2">
            <Field label="Monto a devolver ($)">
              <MoneyInput key={montoResetKey} value={monto} onChange={setMonto} className={inputCls} />
            </Field>
            <button
              type="button"
              onClick={() => { setMonto(totalAEstaFecha); setMontoResetKey((k) => k + 1); }}
              className={`${btnGhost} mt-1.5`}
            >
              Cargar el total a esta fecha ({fmtARS(totalAEstaFecha)})
            </button>
          </div>
          <div className="rounded-md border border-dashed border-stone-300 bg-stone-50 p-3 text-xs text-slate-600 md:col-span-2">
            <div className="flex items-center justify-between"><span>Interés acumulado hasta esta fecha</span><span className="font-mono">{fmtARS(interesAlPagar)}</span></div>
            <div className="flex items-center justify-between"><span>Va a amortizar capital</span><span className="font-mono">{fmtARS(aCapital)}</span></div>
            <div className="mt-1 flex items-center justify-between border-t border-stone-200 pt-1 font-semibold text-slate-800"><span>Saldo de capital restante</span><span className="font-mono">{fmtARS(saldoRestante)}</span></div>
          </div>
          <div className="flex items-end gap-2 md:col-span-2">
            <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button>
            <button type="button" onClick={onClose} className={btnGhost}>Cancelar</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// Modal para editar un gasto/factura o un cobro de socio ya cargado — accesible
// desde su propia tabla y desde el ledger de Movimientos en Cuentas. Sirve sobre
// todo para agregar la factura (tipo A/B/C, o el archivo) cuando al cargarlo
// todavía no se tenía.
function ModalEditarMovimiento({ editando, comprasFacturas, cobrosSocios, ingresos, movimientosManual, obras, onClose, onGuardarCompra, onGuardarCobro, onGuardarIngreso, onGuardarManual, onEliminar }) {
  const registroInicial =
    editando.origen === "compras_facturas" ? comprasFacturas.find((c) => c.id === editando.origenId)
    : editando.origen === "cobros_socios" ? cobrosSocios.find((c) => c.id === editando.origenId)
    : editando.origen === "ingresos" ? ingresos.find((i) => i.id === editando.origenId)
    : movimientosManual.find((m) => m.id === editando.origenId);
  const [form, setForm] = useState(registroInicial || {});
  if (!registroInicial) return null;

  function guardar(e) {
    e.preventDefault();
    if (editando.origen === "compras_facturas") {
      onGuardarCompra(editando.origenId, {
        obraId: form.obraId,
        proveedor: form.proveedor,
        categoria: form.categoria,
        descripcion: form.descripcion || "",
        monto: Number(form.monto) || 0,
        formalidad: form.formalidad,
        tipoFactura: form.tipoFactura,
        archivo: form.archivo,
        nombreArchivo: form.nombreArchivo,
        tipoArchivo: form.tipoArchivo,
      });
    } else if (editando.origen === "cobros_socios") {
      onGuardarCobro(editando.origenId, {
        socio: form.socio,
        fecha: form.fecha,
        monto: Number(form.monto) || 0,
        cuenta: form.cuenta,
        medioBancario: form.cuenta === "Banco" ? form.medioBancario : null,
        formalidad: form.formalidad,
        tipoFactura: form.tipoFactura,
        archivo: form.archivo,
        nombreArchivo: form.nombreArchivo,
        tipoArchivo: form.tipoArchivo,
        observaciones: form.observaciones,
      });
    } else if (editando.origen === "ingresos") {
      onGuardarIngreso(editando.origenId, {
        obraId: form.obraId,
        concepto: form.concepto,
        monto: Number(form.monto) || 0,
        formalidad: form.formalidad,
        cuenta: form.cuenta,
        medioBancario: form.cuenta === "Banco" ? form.medioBancario : null,
        estado: form.estado,
        fechaCobroEstimada: form.estado === "Pendiente" ? form.fechaCobroEstimada : null,
        archivo: form.archivo,
        nombreArchivo: form.nombreArchivo,
        tipoArchivo: form.tipoArchivo,
      });
    } else {
      onGuardarManual(editando.origenId, {
        fecha: form.fecha,
        detalle: form.detalle || "",
        formalidad: form.formalidad,
        cuentaOrigen: form.cuentaOrigen,
        cuentaDestino: form.cuentaDestino,
        monto: Number(form.monto) || 0,
      });
    }
  }

  const titulo = {
    compras_facturas: "Editar gasto / factura",
    cobros_socios: "Editar cobro",
    ingresos: "Editar ingreso",
    movimientos_cuenta: "Editar transferencia entre cuentas",
  }[editando.origen];

  return (
    <div className="fixed inset-0 z-40 flex items-start justify-center overflow-y-auto bg-black/50 p-4 sm:items-center" onClick={onClose}>
      <div className="w-full max-w-2xl rounded-lg bg-white p-5 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between">
          <h3 className="text-lg font-bold text-slate-900">{titulo}</h3>
          <button onClick={onClose}><X size={18} /></button>
        </div>
        <form onSubmit={guardar} className="grid grid-cols-1 gap-4 md:grid-cols-2">
          {editando.origen === "compras_facturas" && (
            <>
              <Field label="Obra">
                <select value={form.obraId ?? ""} onChange={(e) => setForm((f) => ({ ...f, obraId: e.target.value ? Number(e.target.value) : null }))} className={inputCls}>
                  <option value="">General (sin obra específica)</option>
                  {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                </select>
              </Field>
              <Field label="Proveedor - Nombre de fantasía">
                <input
                  value={form.proveedor || ""}
                  onChange={(e) => setForm((f) => ({ ...f, proveedor: e.target.value }))}
                  required
                  className={inputCls}
                />
              </Field>
              <Field label="Categoría">
                <select value={form.categoria} onChange={(e) => setForm((f) => ({ ...f, categoria: e.target.value }))} className={inputCls}>
                  {CATEGORIAS_GASTO.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
              <Field label="Descripción">
                <input
                  value={form.descripcion || ""}
                  onChange={(e) => setForm((f) => ({ ...f, descripcion: e.target.value }))}
                  placeholder="Ej: Nafta camioneta, cemento para losa..."
                  className={inputCls}
                />
              </Field>
              <Field label="Precio final ($)">
                <MoneyInput value={form.monto} onChange={(v) => setForm((f) => ({ ...f, monto: v }))} className={inputCls} />
              </Field>
              <Field label="Formalidad">
                <select value={form.formalidad} onChange={(e) => setForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                  {FORMALIDADES.map((x) => <option key={x}>{x}</option>)}
                </select>
              </Field>
              <Field label="Factura">
                <select value={form.tipoFactura || "Sin factura"} onChange={(e) => setForm((f) => ({ ...f, tipoFactura: e.target.value }))} className={inputCls}>
                  {TIPOS_FACTURA.map((t) => <option key={t}>{t}</option>)}
                </select>
              </Field>
            </>
          )}
          {editando.origen === "cobros_socios" && (
            <>
              <Field label="Socio">
                <select value={form.socio} onChange={(e) => setForm((f) => ({ ...f, socio: e.target.value }))} className={inputCls}>
                  {["Ricardo", "Pablo"].map((s) => <option key={s}>{s}</option>)}
                </select>
              </Field>
              <Field label="Fecha">
                <input type="date" value={form.fecha} onChange={(e) => setForm((f) => ({ ...f, fecha: e.target.value }))} required className={inputCls} />
              </Field>
              <Field label="Monto ($)">
                <MoneyInput value={form.monto} onChange={(v) => setForm((f) => ({ ...f, monto: v }))} className={inputCls} />
              </Field>
              <Field label="Cuenta">
                <select value={form.cuenta} onChange={(e) => setForm((f) => ({ ...f, cuenta: e.target.value }))} className={inputCls}>
                  {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
              {form.cuenta === "Banco" && (
                <Field label="Medio">
                  <select value={form.medioBancario || "Transferencia"} onChange={(e) => setForm((f) => ({ ...f, medioBancario: e.target.value }))} className={inputCls}>
                    <option value="Transferencia">Transferencia</option>
                    <option value="eCheq">eCheq</option>
                  </select>
                </Field>
              )}
              <Field label="Formalidad">
                <select value={form.formalidad} onChange={(e) => setForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                  {FORMALIDADES.map((x) => <option key={x}>{x}</option>)}
                </select>
              </Field>
              <Field label="Factura">
                <select value={form.tipoFactura || "Sin factura"} onChange={(e) => setForm((f) => ({ ...f, tipoFactura: e.target.value }))} className={inputCls}>
                  {TIPOS_FACTURA.map((t) => <option key={t}>{t}</option>)}
                </select>
              </Field>
              <Field label="Observaciones">
                <input value={form.observaciones || ""} onChange={(e) => setForm((f) => ({ ...f, observaciones: e.target.value }))} className={inputCls} />
              </Field>
            </>
          )}
          {editando.origen === "ingresos" && (
            <>
              <Field label="Obra">
                <select value={form.obraId ?? ""} onChange={(e) => setForm((f) => ({ ...f, obraId: e.target.value ? Number(e.target.value) : null }))} className={inputCls}>
                  <option value="">General (sin obra específica)</option>
                  {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                </select>
              </Field>
              <Field label="Concepto">
                <input value={form.concepto || ""} onChange={(e) => setForm((f) => ({ ...f, concepto: e.target.value }))} required className={inputCls} />
              </Field>
              <Field label="Monto ($)">
                <MoneyInput value={form.monto} onChange={(v) => setForm((f) => ({ ...f, monto: v }))} className={inputCls} />
              </Field>
              <Field label="Cuenta">
                <select value={form.cuenta} onChange={(e) => setForm((f) => ({ ...f, cuenta: e.target.value }))} className={inputCls}>
                  {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
              {form.cuenta === "Banco" && (
                <Field label="Medio">
                  <select value={form.medioBancario || "Transferencia"} onChange={(e) => setForm((f) => ({ ...f, medioBancario: e.target.value }))} className={inputCls}>
                    <option value="Transferencia">Transferencia</option>
                    <option value="eCheq">eCheq</option>
                  </select>
                </Field>
              )}
              <Field label="Formalidad">
                <select value={form.formalidad} onChange={(e) => setForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                  {FORMALIDADES.map((x) => <option key={x}>{x}</option>)}
                </select>
              </Field>
              <Field label="Estado">
                <select value={form.estado} onChange={(e) => setForm((f) => ({ ...f, estado: e.target.value }))} className={inputCls}>
                  <option value="Cobrado">Cobrado</option>
                  <option value="Pendiente">Pendiente</option>
                </select>
              </Field>
              {form.estado === "Pendiente" && (
                <Field label="Fecha estimada de cobro">
                  <input type="date" value={form.fechaCobroEstimada || ""} onChange={(e) => setForm((f) => ({ ...f, fechaCobroEstimada: e.target.value }))} className={inputCls} />
                </Field>
              )}
            </>
          )}
          {editando.origen === "movimientos_cuenta" && (
            <>
              <Field label="Fecha">
                <input type="date" value={form.fecha} onChange={(e) => setForm((f) => ({ ...f, fecha: e.target.value }))} required className={inputCls} />
              </Field>
              <Field label="Detalle">
                <input value={form.detalle || ""} onChange={(e) => setForm((f) => ({ ...f, detalle: e.target.value }))} placeholder="Ej: Pase de efectivo a banco" className={inputCls} />
              </Field>
              <Field label="Formalidad">
                <select value={form.formalidad} onChange={(e) => setForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                  {FORMALIDADES.map((x) => <option key={x}>{x}</option>)}
                </select>
              </Field>
              <Field label="Monto ($)">
                <MoneyInput value={form.monto} onChange={(v) => setForm((f) => ({ ...f, monto: v }))} className={inputCls} />
              </Field>
              <Field label="Cuenta donde sale">
                <select value={form.cuentaOrigen} onChange={(e) => setForm((f) => ({ ...f, cuentaOrigen: e.target.value }))} className={inputCls}>
                  {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
              <Field label="Cuenta que recibe">
                <select value={form.cuentaDestino} onChange={(e) => setForm((f) => ({ ...f, cuentaDestino: e.target.value }))} className={inputCls}>
                  {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
            </>
          )}
          {editando.origen !== "movimientos_cuenta" && (
            <div className="md:col-span-2">
              <ArchivoInput
                label="Factura / comprobante (PDF o foto)"
                value={form.archivo}
                nombreArchivo={form.nombreArchivo}
                onChange={(archivo, nombreArchivo, tipoArchivo) => setForm((f) => ({ ...f, archivo, nombreArchivo, tipoArchivo }))}
              />
            </div>
          )}
          <div className="flex flex-wrap items-center gap-2 md:col-span-2">
            <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button>
            <button type="button" onClick={onClose} className={btnGhost}>Cancelar</button>
            <button type="button" onClick={() => onEliminar(editando)} className={`${btnGhostDanger} ml-auto`}>Eliminar</button>
          </div>
        </form>
      </div>
    </div>
  );
}
const btnGhostDanger = "rounded-md border border-rose-300 px-2.5 py-1 text-xs font-semibold text-rose-600 hover:bg-rose-50";

// Icono de tacho de basura para mandar un registro a la Papelera — se repite al
// costado de cada fila/tarjeta en las secciones que soportan Papelera.
function BotonEliminar({ onClick, title = "Eliminar" }) {
  return (
    <button
      type="button"
      onClick={onClick}
      title={title}
      className="rounded-md border border-transparent p-1 text-slate-400 hover:border-rose-300 hover:bg-rose-50 hover:text-rose-600"
    >
      <Trash2 size={14} />
    </button>
  );
}

// Días que le quedan a un registro de la Papelera antes de que el cron lo
// borre solo (7 días desde que se eliminó).
function diasParaPurgar(eliminadoEn) {
  if (!eliminadoEn) return 7;
  const vencimiento = new Date(eliminadoEn).getTime() + 7 * 24 * 60 * 60 * 1000;
  return Math.max(0, Math.ceil((vencimiento - Date.now()) / 86400000));
}

// Un bloque de la Papelera por tipo de registro (herramientas, personal, gastos,
// etc.) — mismo formato para los diez, solo cambia cómo se arma el nombre/detalle.
function SeccionPapelera({ titulo, items, nombreDe, detalleDe, onRestaurar }) {
  if (items.length === 0) return null;
  return (
    <div>
      <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500">{titulo} ({items.length})</div>
      <div className="space-y-1.5">
        {items.map((x) => {
          const dias = diasParaPurgar(x.eliminadoEn);
          return (
            <div key={x.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-stone-200 px-2.5 py-1.5 text-sm">
              <div>
                <div className="font-medium text-slate-800">{nombreDe(x)}</div>
                {detalleDe && <div className="text-xs text-slate-400">{detalleDe(x)}</div>}
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs text-slate-400">{dias === 0 ? "se borra hoy" : `se borra en ${dias} día(s)`}</span>
                <button onClick={() => onRestaurar(x)} className={btnGhost}>Restaurar</button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// Campo numérico (admite decimales) que solo guarda al perder el foco — para
// porcentajes/factores que se editan poco (Liquidación formal UOCRA).
function PctField({ label, value, onSave, suffix = "%", confirmar = false }) {
  const [v, setV] = useState(value ?? 0);
  useEffect(() => setV(value ?? 0), [value]);
  return (
    <div className="flex items-center gap-2 text-sm">
      <span className="min-w-0 flex-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
        {confirmar ? `${label} (a confirmar)` : label}
      </span>
      <div className="flex shrink-0 items-center gap-1">
        <input
          type="number" step="0.01" value={v}
          onChange={(e) => setV(e.target.value)}
          onBlur={() => onSave(Number(v) || 0)}
          className={`${inputCls} w-16 px-1.5 py-1 text-right`}
        />
        <span className="shrink-0 text-xs text-slate-400">{suffix}</span>
      </div>
    </div>
  );
}

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

// Sube un PDF o una foto (recibo, comprobante) y lo guarda como data URL —
// mismo mecanismo que PhotoInput, pero sin forzar una vista previa de imagen.
// Tiene un botón de "Tomar foto" aparte (con capture="environment") además del
// de "Subir archivo", así en el celular abre directo la cámara en vez de tener
// que ir a buscar la foto a la galería.
function ArchivoInput({ label, value, nombreArchivo, onChange }) {
  const esImagen = value && !nombreArchivo?.toLowerCase().endsWith(".pdf");
  const inputFotoRef = useRef(null);
  const inputArchivoRef = useRef(null);
  async function manejarArchivo(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const dataUrl = await readFileAsDataURL(file);
      onChange(dataUrl, file.name, file.type);
    } catch {
      alert("No se pudo leer el archivo.");
    }
    e.target.value = "";
  }
  return (
    <div className="flex flex-col gap-1 text-sm">
      <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <div className="flex items-center gap-3">
        {value ? (
          esImagen ? (
            <img src={value} alt={label} className="h-14 w-14 rounded-md border border-stone-300 object-cover" />
          ) : (
            <a href={value} target="_blank" rel="noreferrer" className="flex h-14 w-14 flex-col items-center justify-center rounded-md border border-stone-300 text-center text-[9px] text-slate-500 hover:bg-stone-50">
              <FileDown size={16} className="text-slate-400" />
              PDF
            </a>
          )
        ) : (
          <div className="flex h-14 w-14 items-center justify-center rounded-md border border-dashed border-stone-300 text-center text-[9px] text-slate-400">Sin archivo</div>
        )}
        <div className="flex flex-col gap-1.5">
          {nombreArchivo && <span className="max-w-[180px] truncate text-xs text-slate-500">{nombreArchivo}</span>}
          <div className="flex flex-wrap gap-2">
            <button type="button" onClick={() => inputFotoRef.current?.click()} className={btnGhost}>
              <span className="flex items-center gap-1"><Camera size={13} /> Tomar foto</span>
            </button>
            <button type="button" onClick={() => inputArchivoRef.current?.click()} className={btnGhost}>
              <span className="flex items-center gap-1"><Upload size={13} /> Subir archivo</span>
            </button>
          </div>
          <input ref={inputFotoRef} type="file" accept="image/*" capture="environment" onChange={manejarArchivo} className="hidden" />
          <input ref={inputArchivoRef} type="file" accept="application/pdf,image/*" onChange={manejarArchivo} className="hidden" />
          {value && (
            <button type="button" onClick={() => onChange(null, null, null)} className="text-left text-xs text-rose-600 hover:underline">
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
  // El contenido scrollea dentro de <main>, no en el body — al abrir un formulario
  // de edición desde algo que quedó más abajo en la lista, lo subimos a la vista.
  const mainRef = useRef(null);
  function scrollContenidoArriba() {
    mainRef.current?.scrollTo({ top: 0, behavior: "smooth" });
  }

  const DEMO_OBRAS = [
    { id: 1, nombre: "Edificio Belgrano 450", cliente: "Consorcio Belgrano SA", clienteId: 1, presupuesto: 85000000, meses: 10, inicio: "2026-02-01", estado: "En curso", encargadoId: 4, diasLaborables: ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"], horaApertura: "08:00", diaCierre: "Viernes", horaCierre: "18:00", color: PALETA_OBRA[0] },
    { id: 2, nombre: "Casa Quinta Yerba Buena", cliente: "Fam. Ledesma", clienteId: 2, presupuesto: 32000000, meses: 6, inicio: "2026-05-01", estado: "En curso", encargadoId: null, diasLaborables: ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"], horaApertura: "07:30", diaCierre: "Viernes", horaCierre: "17:30", color: PALETA_OBRA[1] },
  ];
  const DEMO_ETAPAS_OBRA = [
    { id: 1, obraId: 1, nombre: "Movimiento de suelos", inicio: "2026-02-01", fin: "2026-02-20", avance: 100 },
    { id: 2, obraId: 1, nombre: "Fundaciones", inicio: "2026-02-15", fin: "2026-03-30", avance: 60 },
    { id: 3, obraId: 2, nombre: "Cimientos", inicio: "2026-05-01", fin: "2026-05-20", avance: 100 },
  ];
  const DEMO_PERSONAL = [
    { id: 1, nombre: "Facundo", apellido: "C", dni: "", telefono: "", categoria: "Ayudante", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "En negro", aseguradoPor: "ART" },
    { id: 2, nombre: "Eduardo", apellido: "Sr", dni: "", telefono: "", categoria: "Oficial", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "En blanco", aseguradoPor: "Seg. Accidentes" },
    { id: 3, nombre: "Daniel", apellido: "Tello", dni: "", telefono: "", categoria: "Oficial Especializado", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "En blanco", aseguradoPor: "ART" },
    { id: 4, nombre: "Pablo", apellido: "Robles", dni: "", telefono: "", categoria: "Gerente", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "En negro", aseguradoPor: "No" },
    { id: 5, nombre: "Pepito", apellido: "Chespirito", dni: "", telefono: "", categoria: "Ayudante", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "En negro", aseguradoPor: "No" },
    { id: 6, nombre: "Emi", apellido: "Perez", dni: "", telefono: "", categoria: "Logística", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "En negro", aseguradoPor: "ART" },
    { id: 7, nombre: "Mario", apellido: "González", dni: "", telefono: "", categoria: "Oficial Especializado", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "Eléctrico", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Tantero", aseguradoPor: "Seg. Accidentes" },
    { id: 8, nombre: "Raúl", apellido: "Medina", dni: "", telefono: "", categoria: "Oficial", costoHora: null, direccion: "", fechaNacimiento: "", estado: "Activo", fotoPersona: null, dniFrente: null, dniDorso: null, manoHabil: "Diestro", tipoSangre: "", tarjetaIeric: "No", observaciones: "", especialidad: "Eléctrico", tallePantalon: "", talleCamisa: "", talleGuantes: "", talleCalzado: "", tipoTrabajador: "Tantero", aseguradoPor: "No" },
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
  // Básico de convenio UOCRA CCT 76/75, Zona A (San Juan) — referencia julio 2026.
  // No es lo que la empresa paga (eso es costosCategoria/DEMO_COSTOS): es el piso
  // legal que usa la calculadora de Liquidación formal.
  const DEMO_BASICOS_CONVENIO = [
    { id: 1, categoria: "Oficial Especializado", mes: "2026-07", basicoHora: 6666 },
    { id: 2, categoria: "Oficial", mes: "2026-07", basicoHora: 5703 },
    { id: 3, categoria: "Medio Oficial", mes: "2026-07", basicoHora: 5270 },
    { id: 4, categoria: "Ayudante", mes: "2026-07", basicoHora: 4851 },
  ];
  const DEMO_CONFIG_LIQUIDACION = [
    {
      id: 1, presentismoPct: 20, horaExtra50Pct: 50, horaExtra100Pct: 100, antiguedadPctAnio: 0,
      aporteJubilacionPct: 11, aporteObraSocialPct: 3, aportePamiPct: 3, aporteSindicalPct: 2,
      contribObraSocialPct: 6, contribEmpresariaPct: 2, contribJubilacionPct: 0,
      fondoCesePrimerAnioPct: 12, fondoCesePosteriorPct: 8, iericMontoFijo: 0,
    },
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
    { id: 1, razonSocial: "Corralón San Martín", cuit: "30-12345678-9", domicilio: "Ruta 40 km 12, San Juan", contacto: "Marcos Díaz", telefono: "264-4000001", esTaller: "No", diaPago: 10 },
    { id: 2, razonSocial: "Electromecánica Ríos", cuit: "30-98765432-1", domicilio: "Av. Libertador 850, San Juan", contacto: "Ríos Hnos.", telefono: "264-4000002", esTaller: "Sí" },
  ];
  const DEMO_CLIENTES = [
    { id: 1, razonSocial: "Consorcio Belgrano SA", nombreFantasia: "Edificio Belgrano", cuit: "30-55566677-8", domicilio: "Belgrano 450, San Juan", contacto: "Adm. Belgrano", telefono: "264-4100001" },
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
    { id: 10, fecha: "2026-08-20", obraId: 1, ordenCompraId: null, proveedor: "Aberturas del Norte", categoria: "Materiales", monto: 2000000, estado: "Pendiente", formalidad: "Blanco", formaPago: "eCheq", fechaPagoEcheq: "2026-10-15", cuenta: "Banco" },
    { id: 11, fecha: "2026-08-10", obraId: 1, ordenCompraId: null, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 900000, estado: "Pendiente", formalidad: "Blanco", formaPago: "Cuenta corriente", cuenta: null },
    { id: 12, fecha: "2026-08-22", obraId: 2, ordenCompraId: null, proveedor: "Corralón San Martín", categoria: "Materiales", monto: 350000, estado: "Pendiente", formalidad: "Blanco", formaPago: "Cuenta corriente", cuenta: null },
  ];

  const DEMO_INGRESOS = [
    { id: 1, fecha: "2026-02-05", obraId: 1, concepto: "Anticipo certificado 1", monto: 20000000, formalidad: "Blanco", cuenta: "Banco", estado: "Cobrado" },
    { id: 2, fecha: "2026-04-10", obraId: 1, concepto: "Certificado de avance 2", monto: 18000000, formalidad: "Blanco", cuenta: "Banco", estado: "Cobrado" },
    { id: 3, fecha: "2026-05-15", obraId: 1, concepto: "Adicional acordado con el cliente", monto: 6000000, formalidad: "Negro", cuenta: "Efectivo", estado: "Cobrado" },
    { id: 4, fecha: "2026-05-01", obraId: 2, concepto: "Anticipo Fam. Ledesma", monto: 12000000, formalidad: "Blanco", cuenta: "Mercado Pago", estado: "Cobrado" },
    { id: 5, fecha: "2026-06-20", obraId: 2, concepto: "Pago en mano acordado", monto: 4000000, formalidad: "Negro", cuenta: "Efectivo", estado: "Cobrado" },
    { id: 6, fecha: "2026-09-20", obraId: 1, concepto: "Certificado de avance 3", monto: 15000000, formalidad: "Blanco", cuenta: "Banco", medioBancario: "eCheq", estado: "Pendiente", fechaCobroEstimada: "2026-10-05" },
  ];

  const DEMO_PRESTAMOS = [
    { id: 1, fecha: "2026-06-01", acreedor: "Inversor Juan Pérez", capital: 5000000, tasaAnualPct: 60, cuenta: "Banco", formalidad: "Blanco", fechaEstimadaDevolucion: "2026-12-01", estado: "Vigente", fechaPago: null, montoPagado: null },
    { id: 2, fecha: "2026-01-15", acreedor: "Banco San Juan", capital: 2000000, tasaAnualPct: 40, cuenta: "Banco", formalidad: "Blanco", fechaEstimadaDevolucion: "2026-04-15", estado: "Pagado", fechaPago: "2026-04-10", montoPagado: 2186301.37 },
  ];
  const DEMO_PRESTAMOS_PAGOS = [
    { id: 1, prestamoId: 1, fecha: "2026-07-15", monto: 1000000, cuenta: "Banco" },
  ];

  const DEMO_COBROS_SOCIOS = [
    { id: 1, fecha: "2026-07-15", socio: "Ricardo", monto: 1500000, cuenta: "Banco", medioBancario: "Transferencia", formalidad: "Blanco", archivo: null, nombreArchivo: null, tipoArchivo: null, observaciones: "" },
    { id: 2, fecha: "2026-08-01", socio: "Pablo", monto: 1200000, cuenta: "Efectivo", medioBancario: null, formalidad: "Negro", archivo: null, nombreArchivo: null, tipoArchivo: null, observaciones: "" },
  ];

  const DEMO_TANTEROS = [
    { id: 1, nombreGrupo: "Mario Electricista", obraId: 1, integrantes: [7, 8], precioTotal: 12000000, formalidad: "Blanco" },
  ];
  const DEMO_AVANCES_TANTEROS = [
    { id: 1, tanteroId: 1, fecha: "2026-06-01", monto: 4000000, descripcion: "1er avance — cableado planta baja", cuenta: "Banco", formalidad: "Blanco" },
    { id: 2, tanteroId: 1, fecha: "2026-07-10", monto: 3000000, descripcion: "2do avance — tablero principal", cuenta: "Banco", formalidad: "Blanco" },
  ];

  const [obras, setObras] = useState(isSupabaseConfigured ? [] : DEMO_OBRAS);
  const [selectedObraId, setSelectedObraId] = useState(1);
  const [personalRaw, setPersonal] = useState(isSupabaseConfigured ? [] : DEMO_PERSONAL);
  const [costosCategoria, setCostosCategoria] = useState(isSupabaseConfigured ? [] : DEMO_COSTOS);
  const [basicosConvenio, setBasicosConvenio] = useState(isSupabaseConfigured ? [] : DEMO_BASICOS_CONVENIO);
  const [configLiquidacion, setConfigLiquidacion] = useState(isSupabaseConfigured ? [] : DEMO_CONFIG_LIQUIDACION);
  const [liquidacionesFormales, setLiquidacionesFormales] = useState(isSupabaseConfigured ? [] : []);
  const [recibosLiquidacion, setRecibosLiquidacion] = useState(isSupabaseConfigured ? [] : []);
  const [asistencia, setAsistencia] = useState(isSupabaseConfigured ? [] : DEMO_ASISTENCIA);
  const [herramientasRaw, setHerramientas] = useState(isSupabaseConfigured ? [] : DEMO_HERRAMIENTAS);
  const [combosHerramientas, setCombosHerramientas] = useState(isSupabaseConfigured ? [] : DEMO_COMBOS);
  const [catalogoNombresHerr, setCatalogoNombresHerr] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_NOMBRES);
  const [catalogoMarcas, setCatalogoMarcas] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_MARCAS);
  const [catalogoChicas, setCatalogoChicas] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_CHICAS);
  const [proveedoresRaw, setProveedores] = useState(isSupabaseConfigured ? [] : DEMO_PROVEEDORES);
  const [clientesRaw, setClientes] = useState(isSupabaseConfigured ? [] : DEMO_CLIENTES);
  const [remitos, setRemitos] = useState(isSupabaseConfigured ? [] : DEMO_REMITOS);
  const [auditorias, setAuditorias] = useState(isSupabaseConfigured ? [] : DEMO_AUDITORIAS);
  const [feriados, setFeriados] = useState(isSupabaseConfigured ? [] : DEMO_FERIADOS);
  const [subcategoriasMat, setSubcategoriasMat] = useState(isSupabaseConfigured ? [] : DEMO_SUBCATEGORIAS_MAT);
  const [tiposMaterial, setTiposMaterial] = useState(isSupabaseConfigured ? [] : DEMO_TIPOS_MATERIAL);
  const [catalogoMateriales, setCatalogoMateriales] = useState(isSupabaseConfigured ? [] : DEMO_CATALOGO_MATERIALES);
  const [presupuestoGeneral, setPresupuestoGeneral] = useState(isSupabaseConfigured ? [] : DEMO_PRESUPUESTO_GENERAL);
  const [presupuestoMateriales, setPresupuestoMateriales] = useState(isSupabaseConfigured ? [] : DEMO_PRESUPUESTO_MATERIALES);
  const [pedidosMaterialesRaw, setPedidosMateriales] = useState(isSupabaseConfigured ? [] : DEMO_PEDIDOS_MATERIALES);
  const [stockMateriales, setStockMateriales] = useState(isSupabaseConfigured ? [] : DEMO_STOCK_MATERIALES);
  const [ordenesCompraRaw, setOrdenesCompra] = useState(isSupabaseConfigured ? [] : DEMO_OC);
  const [comprasFacturasRaw, setComprasFacturas] = useState(isSupabaseConfigured ? [] : DEMO_FACTURAS);
  const [ingresosRaw, setIngresos] = useState(isSupabaseConfigured ? [] : DEMO_INGRESOS);
  // Ajustes manuales de cuentas — pases de una cuenta a otra u otras correcciones que
  // no son ni una compra ni un ingreso de obra, para no ensuciar esas dos pestañas.
  const [movimientosManual, setMovimientosManual] = useState([]);
  // Plata contada a mano (caja física / resumen bancario) para comparar contra lo
  // que el sistema calcula y detectar errores de carga.
  const [dineroReal, setDineroReal] = useState([]);
  // Préstamos de inversores o bancos: el capital entra a una cuenta como plata real,
  // pero es una deuda, no un ingreso — el interés se calcula solo, día a día, hasta
  // que se marca como devuelto. Siempre "General" (sin obra), como pidió el usuario.
  const [prestamosRaw, setPrestamos] = useState(isSupabaseConfigured ? [] : DEMO_PRESTAMOS);
  // Devoluciones parciales de cada préstamo — cada una recalcula cuánto capital
  // queda pendiente y, de ahí en más, el interés corre sobre ese saldo.
  const [prestamosPagos, setPrestamosPagos] = useState(isSupabaseConfigured ? [] : DEMO_PRESTAMOS_PAGOS);
  // Retiros de los socios (Ricardo y Pablo) — plata real que sale de la caja de la
  // empresa, separada de Gastos/Facturas para poder ver el historial de cada uno.
  const [cobrosSociosRaw, setCobrosSocios] = useState(isSupabaseConfigured ? [] : DEMO_COBROS_SOCIOS);
  const [tanteros, setTanteros] = useState(isSupabaseConfigured ? [] : DEMO_TANTEROS);
  const [avancesTanteros, setAvancesTanteros] = useState(isSupabaseConfigured ? [] : DEMO_AVANCES_TANTEROS);
  // Etapas de la Planificación (Gantt) de cada obra.
  const [etapasObraRaw, setEtapasObra] = useState(isSupabaseConfigured ? [] : DEMO_ETAPAS_OBRA);

  // Papelera general: un registro "eliminado" (con fecha en eliminadoEn) deja de
  // contar en toda la app —listados, selectores y balances— hasta que se
  // restaura desde la pestaña Papelera o se purga solo a los 7 días. Estas son
  // las versiones "activas" (sin los eliminados) que usa el resto de la app;
  // los *Raw de arriba son los que trae/guarda la base, con todo incluido.
  const personal = personalRaw.filter((p) => !p.eliminadoEn);
  const herramientas = herramientasRaw.filter((h) => !h.eliminadoEn);
  const proveedores = proveedoresRaw.filter((p) => !p.eliminadoEn);
  const clientes = clientesRaw.filter((c) => !c.eliminadoEn);
  const pedidosMateriales = pedidosMaterialesRaw.filter((p) => !p.eliminadoEn);
  const ordenesCompra = ordenesCompraRaw.filter((o) => !o.eliminadoEn);
  const comprasFacturas = comprasFacturasRaw.filter((c) => !c.eliminadoEn);
  const ingresos = ingresosRaw.filter((i) => !i.eliminadoEn);
  const prestamos = prestamosRaw.filter((p) => !p.eliminadoEn);
  const cobrosSocios = cobrosSociosRaw.filter((c) => !c.eliminadoEn);
  const etapasObra = etapasObraRaw.filter((e) => !e.eliminadoEn);

  // Lo que está en la Papelera ahora mismo, para la pestaña "Papelera".
  const personalPapelera = personalRaw.filter((p) => p.eliminadoEn);
  const herramientasPapelera = herramientasRaw.filter((h) => h.eliminadoEn);
  const proveedoresPapelera = proveedoresRaw.filter((p) => p.eliminadoEn);
  const clientesPapelera = clientesRaw.filter((c) => c.eliminadoEn);
  const pedidosMaterialesPapelera = pedidosMaterialesRaw.filter((p) => p.eliminadoEn);
  const ordenesCompraPapelera = ordenesCompraRaw.filter((o) => o.eliminadoEn);
  const comprasFacturasPapelera = comprasFacturasRaw.filter((c) => c.eliminadoEn);
  const ingresosPapelera = ingresosRaw.filter((i) => i.eliminadoEn);
  const prestamosPapelera = prestamosRaw.filter((p) => p.eliminadoEn);
  const cobrosSociosPapelera = cobrosSociosRaw.filter((c) => c.eliminadoEn);
  const etapasObraPapelera = etapasObraRaw.filter((e) => e.eliminadoEn);

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
        const [o, p, cc, a, h, oc, cf, ing, tt, av, ch, cn, cm, cch, pv, rm, au, fer, cli, sm, tm, cma, pma, ped, pg, stk, bc, cl, lf, rl, mm, dr, pr, cs, pp, eo] = await Promise.all([
          sbSelect("obras"), sbSelect("personal"), sbSelect("costos_categoria"), sbSelect("asistencia"),
          sbSelect("herramientas"), sbSelect("ordenes_compra"), sbSelect("compras_facturas"), sbSelect("ingresos"),
          sbSelect("tanteros"), sbSelect("avances_tanteros"), sbSelect("combos_herramientas"),
          sbSelect("catalogo_nombres_herramienta"), sbSelect("catalogo_marcas"), sbSelect("catalogo_herramientas_chicas"),
          sbSelect("proveedores"), sbSelect("remitos"), sbSelect("auditorias_herramientas"), sbSelect("feriados"), sbSelect("clientes"),
          sbSelect("subcategorias_material"), sbSelect("tipos_material"), sbSelect("catalogo_materiales"), sbSelect("presupuesto_materiales"),
          sbSelect("pedidos_materiales"), sbSelect("presupuesto_general"), sbSelect("stock_materiales"),
          sbSelect("basicos_convenio"), sbSelect("config_liquidacion"), sbSelect("liquidaciones_formales"), sbSelect("recibos_liquidacion"),
          sbSelect("movimientos_cuenta"), sbSelect("dinero_real_cuentas"), sbSelect("prestamos"), sbSelect("cobros_socios"),
          sbSelect("prestamos_pagos"), sbSelect("etapas_obra"),
        ]);
        setObras(o);
        setPersonal(p);
        setCostosCategoria(cc);
        setBasicosConvenio(bc);
        setLiquidacionesFormales(lf);
        setRecibosLiquidacion(rl);
        setConfigLiquidacion(cl);
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
        setMovimientosManual(mm);
        setDineroReal(dr);
        setPrestamos(pr);
        setCobrosSocios(cs);
        setPrestamosPagos(pp);
        setEtapasObra(eo);
        if (o[0]) setSelectedObraId(o[0].id);
      } catch (err) {
        setDbError(err.message);
      } finally {
        setDbLoading(false);
      }
    })();
  }, [reloadKey]);

  // useRef (no una variable suelta) para que el contador sobreviva entre renders —
  // si no, cada nuevo render lo reiniciaba a 200 y dos altas separadas en modo demo
  // terminaban con el mismo id.
  const nextIdRef = useRef(200);
  const genId = () => nextIdRef.current++;

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

  // Papelera general (herramientas, personal, gastos/facturas, ingresos,
  // proveedores, clientes, préstamos, cobros de socios, órdenes de compra y
  // pedidos de obra): en vez de borrar directo, se marca con la fecha y
  // desaparece de toda la app (listados, selectores y balances) pero se puede
  // restaurar durante 7 días desde la pestaña "Papelera" — pasado ese plazo se
  // borra solo. Al estar en la papelera deja de contar en cualquier saldo.
  function moverAPapelera(table, id, setter, nombre) {
    if (!window.confirm(`¿Eliminar "${nombre}"? Se puede restaurar desde la Papelera durante 7 días; pasado ese plazo se borra definitivamente.`)) return;
    updateRecord(table, id, { eliminadoEn: new Date().toISOString() }, setter);
  }
  function restaurarDePapelera(table, id, setter) {
    updateRecord(table, id, { eliminadoEn: null }, setter);
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

  // ---------- Factores UOCRA / Régimen de la Construcción (Liquidación formal) ----------
  function guardarBasicoConvenioCelda(categoria, mes, valor) {
    const existente = basicosConvenio.find((c) => c.categoria === categoria && c.mes === mes);
    if (existente) {
      if (existente.basicoHora === valor) return;
      updateRecord("basicos_convenio", existente.id, { basicoHora: valor }, setBasicosConvenio);
    } else {
      if (!valor) return;
      addRecord("basicos_convenio", { categoria, mes, basicoHora: valor }, setBasicosConvenio);
    }
  }
  const cfgLiq = configLiquidacion[0] || DEMO_CONFIG_LIQUIDACION[0];
  function actualizarConfigLiquidacion(campo, valor) {
    if (!cfgLiq?.id) return;
    updateRecord("config_liquidacion", cfgLiq.id, { [campo]: valor }, setConfigLiquidacion);
  }

  const [viewingPersonId, setViewingPersonId] = useState(null);
  const [modoSeleccionPdf, setModoSeleccionPdf] = useState(false);
  const [showCostosPanel, setShowCostosPanel] = useState(false);
  const [showSegurosPanel, setShowSegurosPanel] = useState(false);
  const [seleccionadosPdf, setSeleccionadosPdf] = useState([]);
  const toggleSeleccionPdf = (id) =>
    setSeleccionadosPdf((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  const viewingPerson = personal.find((p) => p.id === viewingPersonId) || null;
  // Historial de pagos de esta persona: semanas en negro ya pagadas (Pendientes de pago)
  // + quincenas en blanco ya liquidadas (Liquidación formal, con el costo real cargado).
  const historialPagosPersona = viewingPerson
    ? (() => {
        const nombreCompleto = nombreCompletoDe(viewingPerson);
        const negroPorSemana = {};
        asistencia
          .filter((a) => a.nombre === nombreCompleto && a.estadoPago === "Pagado")
          .forEach((a) => {
            const key = claveSemana(a.fecha);
            if (!negroPorSemana[key]) negroPorSemana[key] = { fecha: key, tipo: "Negro", monto: 0 };
            negroPorSemana[key].monto += a.montoAbonado || 0;
          });
        const negro = Object.values(negroPorSemana).map((n) => ({ ...n, periodo: `Semana del ${fmtFecha(n.fecha)}` }));
        const blanco = liquidacionesFormales
          .filter((l) => l.nombre === nombreCompleto && l.costoRealBlanco != null)
          .map((l) => ({
            fecha: rangoQuincena(l.mes, l.quincena).desde,
            tipo: "Blanco",
            monto: l.costoRealBlanco,
            periodo: etiquetaQuincena(l.mes, l.quincena),
          }));
        return [...negro, ...blanco].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));
      })()
    : [];

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
  // - En blanco / en negro: la obra de su registro de asistencia más reciente.
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
    tipoTrabajador: "En negro", aseguradoPor: "No",
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
      tipoTrabajador: p.tipoTrabajador === "Empresa" ? "En negro" : (p.tipoTrabajador || "En negro"),
      aseguradoPor: p.aseguradoPor || "No",
    });
    setEditingPersonalId(p.id);
    setShowPersonalForm(true);
    setViewingPersonId(null);
    scrollContenidoArriba();
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
      aseguradoPor: personalForm.aseguradoPor,
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
    { id: "asistencia", label: "Asistencia", icon: ClipboardCheck },
    { id: "herramientas", label: "Herramientas", icon: Wrench },
    { id: "ingresos", label: "Ingresos", icon: TrendingUp },
    { id: "facturas", label: "Gastos y Facturas", icon: Receipt },
    { id: "personal", label: "Personal/Cuadrillas", icon: Users },
    { id: "cuentas", label: "Cuentas", icon: Landmark },
    { id: "cobros_socios", label: "Cobros Ricardo y Pablo", icon: Banknote },
    { id: "liquidacion", label: "Salario Personal", icon: Wallet },
    { id: "proveedores", label: "Clientes/Proveedores", icon: Truck },
    { id: "calendario", label: "Calendario", icon: CalendarDays },
    // Todavía usamos poco estas dos secciones — quedan al final del menú y resaltadas en amarillo.
    { id: "materiales", label: "Pedidos de Obra", icon: Package },
    { id: "ordenes", label: "Órdenes de Compra", icon: ShoppingCart },
    { id: "papelera", label: "Papelera", icon: Trash2 },
  ];
  const NAV_DESTACADOS = ["materiales", "ordenes"];

  // ---------- Dashboard calculations ----------
  const obraSel = obras.find((o) => o.id === selectedObraId) || obras[0] || null;
  const startDate = obraSel ? fechaLocal(obraSel.inicio) : new Date();
  const meses = obraSel ? Array.from({ length: obraSel.meses }, (_, i) => new Date(startDate.getFullYear(), startDate.getMonth() + i, 1)) : [];
  const gastosObra = obraSel ? comprasFacturas.filter((c) => c.obraId === obraSel.id).sort((a, b) => fechaLocal(a.fecha) - fechaLocal(b.fecha)) : [];

  const chartData = meses.map((d, i) => {
    const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const real = gastosObra.filter((g) => fechaLocal(g.fecha) <= monthEnd).reduce((s, g) => s + g.monto, 0)
      + costoManoDeObraDeObra(obraSel.id, monthEnd);
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
  // Los pedidos de una obra pasan a Órdenes de Compra apenas se solicitan (ver
  // confirmarPedido) — se cuentan ahí (ocPendientesAprobacion), no acá, para no
  // duplicar la alerta. Solo quedan acá las compras generales, que no tienen obra.
  const pedidosPorAprobar = pedidosMateriales.filter((p) => p.estado === "Solicitado" && p.obraId == null && !obraIdsPapelera.has(p.obraId));

  const totalAlertas =
    herramientasAtencion.length + herramientasReparadasRecientes.length + ocPendientesAprobacion.length +
    (hayDesvioAlerta ? 1 : 0) + asistenciasEditadas.length + obrasEnVentanaCierre.length + obrasSinAperturaLunes.length +
    materialesVencidos.length + materialesProximos.length + pedidosPorAprobar.length + personalSinObra5Dias.length;

  // ---------- Forms state ----------
  const [showObraForm, setShowObraForm] = useState(false);
  const [filtroObrasEstado, setFiltroObrasEstado] = useState("En curso");
  // Obras: alterna entre la lista de tarjetas y la Planificación (Gantt) de todas las obras.
  const [vistaObras, setVistaObras] = useState("lista");
  const [agregandoEtapaObraId, setAgregandoEtapaObraId] = useState(null);
  const [editandoEtapaId, setEditandoEtapaId] = useState(null);
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
      cliente: nombreComercial(clientes.find((c) => c.id === Number(f.get("clienteId")))) || "",
      presupuesto: Number(f.get("presupuesto")) || resumenObraImportado?.precioTotalConIva || 0,
      meses: Number(f.get("meses")) || 1,
      inicio: f.get("inicio"),
      estado: "En curso",
      encargadoId: f.get("encargadoId") ? Number(f.get("encargadoId")) : null,
      diaCierre: f.get("diaCierre"),
      horaCierre: f.get("horaCierre"),
      color: PALETA_OBRA[obras.length % PALETA_OBRA.length],
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

  // ---------- Planificación (etapas de obra / Gantt) ----------
  async function agregarEtapa(obraId, datos) {
    const nueva = await addRecord("etapas_obra", {
      obraId,
      nombre: datos.nombre,
      inicio: datos.inicio,
      fin: datos.fin,
      avance: Number(datos.avance) || 0,
    }, setEtapasObra);
    if (nueva) setAgregandoEtapaObraId(null);
  }
  function guardarEdicionEtapa(etapa, datos) {
    updateRecord("etapas_obra", etapa.id, {
      nombre: datos.nombre,
      inicio: datos.inicio,
      fin: datos.fin,
      avance: Number(datos.avance) || 0,
    }, setEtapasObra);
    setEditandoEtapaId(null);
  }
  function eliminarEtapa(etapa) {
    moverAPapelera("etapas_obra", etapa.id, setEtapasObra, etapa.nombre);
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
    scrollContenidoArriba();
  }
  async function guardarEdicionObra(e, obra) {
    e.preventDefault();
    const f = new FormData(e.target);
    const nuevoClienteId = f.get("clienteId") ? Number(f.get("clienteId")) : null;
    await updateRecord("obras", obra.id, {
      nombre: f.get("nombre"),
      clienteId: nuevoClienteId,
      cliente: nombreComercial(clientes.find((c) => c.id === nuevoClienteId)) || obra.cliente,
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
  const [facturaFormaPago, setFacturaFormaPago] = useState("Efectivo");
  const [facturaMedioBancario, setFacturaMedioBancario] = useState("Débito/Transferencia");
  const [facturaPlazoEcheq, setFacturaPlazoEcheq] = useState("30");
  const [facturaArchivo, setFacturaArchivo] = useState(null);
  const [facturaNombreArchivo, setFacturaNombreArchivo] = useState(null);
  const [facturaTipoArchivo, setFacturaTipoArchivo] = useState(null);
  // ---------- PDF mensual para el contador (Gastos y Facturas + Cobros de socios) ----------
  const mesReporteContador = hoyISO().slice(0, 7);
  function generarPdfContadores() {
    const gastosDelMes = comprasFacturas
      .filter((c) => c.fecha?.slice(0, 7) === mesReporteContador && !obraIdsPapelera.has(c.obraId))
      .sort((a, b) => fechaLocal(a.fecha) - fechaLocal(b.fecha));
    const cobrosDelMes = cobrosSocios
      .filter((c) => c.fecha?.slice(0, 7) === mesReporteContador)
      .sort((a, b) => fechaLocal(a.fecha) - fechaLocal(b.fecha));

    if (gastosDelMes.length === 0 && cobrosDelMes.length === 0) {
      alert("No hay gastos ni cobros cargados en ese mes.");
      return;
    }

    const doc = new jsPDF({ unit: "mm", format: "a4" });
    doc.setFillColor(2, 29, 52);
    doc.rect(0, 0, 210, 28, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.text("COMPROBANTES PARA EL CONTADOR", 14, 14);
    doc.setFontSize(10);
    doc.text(nombreMesCuentas(mesReporteContador), 14, 21);
    doc.setTextColor(20, 20, 20);

    let y = 36;
    const totalGastos = gastosDelMes.reduce((s, c) => s + (c.monto || 0), 0);
    if (gastosDelMes.length > 0) {
      doc.setFontSize(12);
      doc.text("Gastos y Facturas", 14, y);
      autoTable(doc, {
        startY: y + 4,
        head: [["Fecha", "Proveedor", "Categoría", "Obra", "Forma de pago", "Formalidad", "Monto", "Estado"]],
        body: gastosDelMes.map((c) => {
          const obra = obras.find((o) => o.id === c.obraId);
          return [
            fmtFecha(c.fecha), c.proveedor, c.categoria, obra?.nombre || "General",
            (c.formaPago || "—") + (c.medioBancario ? ` (${c.medioBancario})` : ""),
            c.formalidad, fmtARS(c.monto), c.estado,
          ];
        }),
        foot: [["", "", "", "", "", "TOTAL", fmtARS(totalGastos), ""]],
        headStyles: { fillColor: [2, 29, 52] },
        footStyles: { fillColor: [245, 245, 244], textColor: [20, 20, 20], fontStyle: "bold" },
        styles: { fontSize: 8 },
      });
      y = (doc.lastAutoTable.finalY || y) + 12;
    }

    const totalCobros = cobrosDelMes.reduce((s, c) => s + (c.monto || 0), 0);
    if (cobrosDelMes.length > 0) {
      doc.setFontSize(12);
      doc.text("Cobros Ricardo y Pablo", 14, y);
      autoTable(doc, {
        startY: y + 4,
        head: [["Fecha", "Socio", "Cuenta", "Formalidad", "Monto", "Comprobante", "Observaciones"]],
        body: cobrosDelMes.map((c) => [
          fmtFecha(c.fecha), c.socio, c.cuenta + (c.medioBancario ? ` (${c.medioBancario})` : ""),
          c.formalidad, fmtARS(c.monto), c.archivo ? "Sí" : "No", c.observaciones || "",
        ]),
        foot: [["", "", "", "TOTAL", fmtARS(totalCobros), "", ""]],
        headStyles: { fillColor: [2, 29, 52] },
        footStyles: { fillColor: [245, 245, 244], textColor: [20, 20, 20], fontStyle: "bold" },
        styles: { fontSize: 8 },
      });
      y = doc.lastAutoTable.finalY || y;
    }

    doc.setFontSize(10);
    doc.setFont(undefined, "bold");
    doc.text(`Total egresos del mes: ${fmtARS(totalGastos + totalCobros)}`, 14, y + 10);
    doc.setFont(undefined, "normal");
    doc.setFontSize(8);
    doc.setTextColor(120, 120, 120);
    doc.text("Generado desde Concretar App — Gastos y Facturas.", 14, y + 18);

    doc.save(`Comprobantes_${mesReporteContador}.pdf`);
  }
  const [showIngresoForm, setShowIngresoForm] = useState(false);
  const [ingresoCuenta, setIngresoCuenta] = useState(CUENTAS[0]);
  const [ingresoMedioBancario, setIngresoMedioBancario] = useState("Transferencia");
  const [ingresoEstado, setIngresoEstado] = useState("Cobrado");
  const [ingresoArchivo, setIngresoArchivo] = useState(null);
  const [ingresoNombreArchivo, setIngresoNombreArchivo] = useState(null);
  const [ingresoTipoArchivo, setIngresoTipoArchivo] = useState(null);
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
    scrollContenidoArriba();
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
  // "Empresa" es el valor legado de antes de separar en blanco/negro — se trata como "En negro".
  function tipoTrabajadorDe(nombreCompleto) {
    const t = personal.find((p) => nombreCompletoDe(p) === nombreCompleto)?.tipoTrabajador;
    return t === "Empresa" ? "En negro" : (t || "En negro");
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
  function basicoConvenioDeCategoria(categoria, fechaStr) {
    const mes = (fechaStr || hoyISO()).slice(0, 7);
    const exacto = basicosConvenio.find((c) => c.categoria === categoria && c.mes === mes);
    if (exacto) return exacto.basicoHora || 0;
    const anteriores = basicosConvenio
      .filter((c) => c.categoria === categoria && c.mes && c.mes <= mes)
      .sort((a, b) => b.mes.localeCompare(a.mes));
    return anteriores[0]?.basicoHora || 0;
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

  // El personal "En blanco" se liquida por quincena (1 al 15 / 16 a fin de mes), no por semana.
  function quincenaDeFecha(fechaStr) {
    return Number(fechaStr.slice(8, 10)) <= 15 ? 1 : 2;
  }
  function rangoQuincena(mes, quincena) {
    const [y, m] = mes.split("-").map(Number);
    const ultimoDia = new Date(y, m, 0).getDate();
    return quincena === 1
      ? { desde: `${mes}-01`, hasta: `${mes}-15` }
      : { desde: `${mes}-16`, hasta: `${mes}-${String(ultimoDia).padStart(2, "0")}` };
  }
  function etiquetaQuincena(mes, quincena) {
    const { desde, hasta } = rangoQuincena(mes, quincena);
    return `${quincena === 1 ? "1ra" : "2da"} quincena de ${nombreMes(mes)} (días ${Number(desde.slice(8, 10))} al ${Number(hasta.slice(8, 10))})`;
  }

  // El personal "En blanco" no se paga en mano acá — su liquidación se cierra sí o sí
  // por la calculadora formal UOCRA (pestaña "Liquidación formal"), pasando por Contaduría.
  const pendientesTodasObras = asistencia.filter(
    (a) => a.estadoPago !== "Pagado" && a.estado !== "Ausente" && (a.horas || 0) > 0 && tipoTrabajadorDe(a.nombre) !== "En blanco"
  );

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

  // Personal "En blanco": queda pendiente por quincena (no por semana) hasta que se carga
  // el costo real del recibo del contador en "Liquidación formal". Hay 5 días desde el
  // cierre de la quincena para pagarles — pasado ese plazo se marca como vencido.
  const asistenciaEnBlanco = asistencia.filter(
    (a) => a.estado !== "Ausente" && (a.horas || 0) > 0 && CATEGORIAS_CONVENIO_UOCRA.includes(categoriaDe(a.nombre)) && tipoTrabajadorDe(a.nombre) === "En blanco"
  );
  const gruposEnBlanco = {}; // "obraId|mes|quincena" -> { obraId, mes, quincena, horas, personas:Set }
  asistenciaEnBlanco.forEach((a) => {
    const mes = a.fecha.slice(0, 7);
    const quincena = quincenaDeFecha(a.fecha);
    const key = `${a.obraId}|${mes}|${quincena}`;
    if (!gruposEnBlanco[key]) gruposEnBlanco[key] = { obraId: a.obraId, mes, quincena, horas: 0, personas: new Set() };
    gruposEnBlanco[key].horas += a.horas || 0;
    gruposEnBlanco[key].personas.add(a.nombre);
  });
  const pendientesEnBlanco = Object.values(gruposEnBlanco)
    .map((g) => {
      const personasArr = [...g.personas].sort();
      // Cerrado solo cuando TODAS las personas de esa obra+quincena ya tienen costo real cargado.
      const cerrado = personasArr.every((nombre) => {
        const l = liquidacionesFormales.find((x) => x.obraId === g.obraId && x.mes === g.mes && x.quincena === g.quincena && x.nombre === nombre);
        return l?.costoRealBlanco != null;
      });
      const { hasta } = rangoQuincena(g.mes, g.quincena);
      const diasDesdeCierre = Math.round((fechaLocal(hoyISO()) - fechaLocal(hasta)) / 86400000);
      return { ...g, personas: personasArr, hasta, diasDesdeCierre, cerrado, vencido: diasDesdeCierre > 5 };
    })
    .filter((g) => !g.cerrado)
    .sort((a, b) => fechaLocal(b.hasta) - fechaLocal(a.hasta));

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

  // ---------- Liquidación formal (UOCRA CCT 76/75 Zona A — San Juan, Ley 22.250) ----------
  // Una fila por persona (no un total agrupado): "Hs. presentes" viene de Asistencia y no
  // se edita acá; "Hs. en recibo" es lo declarado en blanco (por defecto la mitad de las
  // presentes) y se puede tocar a mano — la diferencia queda como horas en negro. Cada
  // campo se guarda solo, sin un botón de "guardar" aparte.
  const [obraFormalId, setObraFormalId] = useState(obras[0]?.id ?? "");
  const [mesFormal, setMesFormal] = useState(hoyISO().slice(0, 7));
  const [quincenaFormal, setQuincenaFormal] = useState(quincenaDeFecha(hoyISO()));
  const [showFactoresLiquidacion, setShowFactoresLiquidacion] = useState(false);

  function esFeriado(fechaStr) {
    return feriados.some((f) => f.fecha === fechaStr);
  }

  const rangoFormalActual = rangoQuincena(mesFormal, quincenaFormal);
  const asistenciaDelMesFormal = asistencia.filter(
    (a) => a.obraId === Number(obraFormalId) && a.fecha >= rangoFormalActual.desde && a.fecha <= rangoFormalActual.hasta
  );
  // Solo entra acá el personal marcado "En blanco" — el resto (en negro) se cierra
  // pagando en mano desde "Pendientes de pago", y los tanteros no pasan por asistencia.
  const nombresDelMesFormal = [...new Set(asistenciaDelMesFormal.map((a) => a.nombre))]
    .filter((nombre) => CATEGORIAS_CONVENIO_UOCRA.includes(categoriaDe(nombre)) && tipoTrabajadorDe(nombre) === "En blanco")
    .sort();

  function registroFormalDe(nombre) {
    return liquidacionesFormales.find(
      (l) => l.obraId === Number(obraFormalId) && l.mes === mesFormal && l.quincena === quincenaFormal && l.nombre === nombre
    );
  }
  function guardarCampoFormal(nombre, patch) {
    const existente = registroFormalDe(nombre);
    if (existente) {
      updateRecord("liquidaciones_formales", existente.id, patch, setLiquidacionesFormales);
    } else {
      addRecord(
        "liquidaciones_formales",
        { obraId: Number(obraFormalId), mes: mesFormal, quincena: quincenaFormal, nombre, horasRecibo: null, presentismo: false, ...patch },
        setLiquidacionesFormales
      );
    }
  }
  const actualizarHorasRecibo = (nombre, valor) => guardarCampoFormal(nombre, { horasRecibo: valor });
  const actualizarPresentismoFormal = (nombre, valor) => guardarCampoFormal(nombre, { presentismo: valor });
  const guardarCostoRealPersona = (nombre, valor) => guardarCampoFormal(nombre, { costoRealBlanco: valor, fechaConfirmacion: hoyISO() });

  const filasFormal = nombresDelMesFormal.map((nombre) => {
    const categoria = categoriaDe(nombre) || CATEGORIAS_PERSONAL[0];
    const registrosAsistencia = asistenciaDelMesFormal.filter((a) => a.nombre === nombre);
    const horasReales = registrosAsistencia.reduce((s, a) => s + (a.horas || 0), 0);
    const diasTrabajados = new Set(registrosAsistencia.map((a) => a.fecha)).size;
    const registro = registroFormalDe(nombre);
    const horasRecibo = registro?.horasRecibo ?? Math.round((horasReales / 2) * 100) / 100;
    const presentismo = registro?.presentismo ?? false;
    const horasNegro = Math.max(0, horasReales - horasRecibo);

    const basicoHora = basicoConvenioDeCategoria(categoria, `${mesFormal}-01`);
    // Lo que no se declara en blanco se sigue pagando informal, al valor que la empresa
    // ya paga hoy por hora (costosCategoria) — así el Gerente ve el gasto total aproximado.
    const costoHoraInformal = costoHoraDeCategoria(categoria, `${mesFormal}-01`);
    const montoBasico = horasRecibo * basicoHora;
    const montoPresentismo = presentismo ? montoBasico * ((cfgLiq.presentismoPct || 0) / 100) : 0;
    const bruto = montoBasico + montoPresentismo;
    const aportesPct = (cfgLiq.aporteJubilacionPct || 0) + (cfgLiq.aporteObraSocialPct || 0) + (cfgLiq.aportePamiPct || 0) + (cfgLiq.aporteSindicalPct || 0);
    const aportes = bruto * (aportesPct / 100);
    const neto = bruto - aportes;
    const contribPct = (cfgLiq.contribObraSocialPct || 0) + (cfgLiq.contribEmpresariaPct || 0) + (cfgLiq.contribJubilacionPct || 0);
    const contribuciones = bruto * (contribPct / 100);
    // Sin dato de antigüedad por persona, se usa siempre el Fondo de Cese "posterior" (8%).
    const fondoCese = bruto * ((cfgLiq.fondoCesePosteriorPct || 0) / 100);
    const costoEmpresa = bruto + contribuciones + fondoCese + (cfgLiq.iericMontoFijo || 0);
    const costoNegro = horasNegro * costoHoraInformal;
    const costoRealBlanco = registro?.costoRealBlanco ?? null;

    return {
      nombre, categoria, horasReales, diasTrabajados, basicoHora,
      horasRecibo, horasNegro, presentismo,
      montoBasico, montoPresentismo, bruto, aportes, neto, contribuciones, fondoCese, costoEmpresa, costoNegro,
      costoRealBlanco, cerrado: costoRealBlanco != null,
    };
  });
  const totalesFormal = filasFormal.reduce(
    (acc, f) => ({
      bruto: acc.bruto + f.bruto, aportes: acc.aportes + f.aportes, neto: acc.neto + f.neto, costoEmpresa: acc.costoEmpresa + f.costoEmpresa,
      horasRecibo: acc.horasRecibo + f.horasRecibo, horasNegro: acc.horasNegro + f.horasNegro, costoNegro: acc.costoNegro + f.costoNegro,
    }),
    { bruto: 0, aportes: 0, neto: 0, costoEmpresa: 0, horasRecibo: 0, horasNegro: 0, costoNegro: 0 }
  );
  const gastoAproximadoTotal = totalesFormal.costoEmpresa + totalesFormal.costoNegro;

  // El recibo (PDF o foto) es uno solo por obra+quincena — se guarda para consulta.
  const reciboFormalActual = recibosLiquidacion.find(
    (r) => r.obraId === Number(obraFormalId) && r.mes === mesFormal && r.quincena === quincenaFormal
  );
  function guardarReciboFormal(dataUrl, nombreArchivo, tipoArchivo) {
    if (!dataUrl) {
      if (reciboFormalActual) deleteRecord("recibos_liquidacion", reciboFormalActual.id, setRecibosLiquidacion);
      return;
    }
    const patch = { archivo: dataUrl, nombreArchivo, tipoArchivo, fechaCarga: hoyISO() };
    if (reciboFormalActual) {
      updateRecord("recibos_liquidacion", reciboFormalActual.id, patch, setRecibosLiquidacion);
    } else {
      addRecord("recibos_liquidacion", { obraId: Number(obraFormalId), mes: mesFormal, quincena: quincenaFormal, ...patch }, setRecibosLiquidacion);
    }
  }

  function generarPdfHorasReales() {
    const obra = obras.find((o) => o.id === Number(obraFormalId));
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    doc.setFillColor(2, 29, 52);
    doc.rect(0, 0, 210, 28, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.text("HORAS TRABAJADAS", 14, 14);
    doc.setFontSize(10);
    doc.text(`${obra?.nombre || "Obra"} · ${etiquetaQuincena(mesFormal, quincenaFormal)}`, 14, 21);
    doc.setTextColor(20, 20, 20);
    autoTable(doc, {
      startY: 36,
      head: [["Persona", "Categoría", "Días trabajados", "Horas presentes", "Horas en recibo", "Horas en negro"]],
      body: filasFormal.map((f) => [f.nombre, f.categoria, String(f.diasTrabajados), String(f.horasReales), String(f.horasRecibo), String(f.horasNegro)]),
      foot: [["", "", "", "TOTAL", String(totalesFormal.horasRecibo), String(totalesFormal.horasNegro)]],
      headStyles: { fillColor: [2, 29, 52] },
      footStyles: { fillColor: [245, 245, 244], textColor: [20, 20, 20], fontStyle: "bold" },
      styles: { fontSize: 9 },
    });
    const finalY = doc.lastAutoTable.finalY || 100;
    doc.setFontSize(8);
    doc.setTextColor(120, 120, 120);
    doc.text(
      [
        "\"Horas en recibo\" es lo que hay que liquidar formalmente; \"Horas en negro\" es el resto de las horas presentes.",
        "Horas presentes según asistencia registrada en obra, sin ajustes.",
      ],
      14, finalY + 8
    );
    doc.save(`Horas_${(obra?.nombre || "obra").replace(/\s+/g, "_")}_${mesFormal}_Q${quincenaFormal}.pdf`);
  }

  function generarPdfLiquidacionFormal() {
    const obra = obras.find((o) => o.id === Number(obraFormalId));
    const doc = new jsPDF({ unit: "mm", format: "a4" });
    doc.setFillColor(2, 29, 52);
    doc.rect(0, 0, 210, 28, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(15);
    doc.text("LIQUIDACIÓN (simulación UOCRA — Zona A / San Juan)", 14, 14);
    doc.setFontSize(10);
    doc.text(`${obra?.nombre || "Obra"} · ${etiquetaQuincena(mesFormal, quincenaFormal)}`, 14, 21);
    doc.setTextColor(20, 20, 20);
    autoTable(doc, {
      startY: 36,
      head: [["Persona", "Categ.", "Hs. en recibo", "Hs. en negro", "Bruto", "Aportes", "Neto", "Costo empresa", "Costo aprox. (negro)", "Costo real"]],
      body: filasFormal.map((f) => [
        f.nombre, f.categoria, String(f.horasRecibo), String(f.horasNegro),
        fmtARS(f.bruto), fmtARS(f.aportes), fmtARS(f.neto), fmtARS(f.costoEmpresa), fmtARS(f.costoNegro),
        f.costoRealBlanco != null ? fmtARS(f.costoRealBlanco) : "—",
      ]),
      foot: [["", "", "", "TOTAL", fmtARS(totalesFormal.bruto), fmtARS(totalesFormal.aportes), fmtARS(totalesFormal.neto), fmtARS(totalesFormal.costoEmpresa), fmtARS(totalesFormal.costoNegro), ""]],
      headStyles: { fillColor: [2, 29, 52] },
      footStyles: { fillColor: [245, 245, 244], textColor: [20, 20, 20], fontStyle: "bold" },
      styles: { fontSize: 7.5 },
    });
    const finalY = doc.lastAutoTable.finalY || 100;
    doc.setFontSize(11);
    doc.setTextColor(20, 20, 20);
    doc.setFont(undefined, "bold");
    doc.text(`Gasto aproximado total (blanco + negro): ${fmtARS(gastoAproximadoTotal)}`, 14, finalY + 12);
    doc.setFont(undefined, "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text(
      [
        "Simulación interna, no reemplaza el recibo oficial. Factores de referencia: UOCRA CCT 76/75 (Zona A / San Juan) y",
        "Fondo de Cese Laboral Ley 22.250, sujetos a la paritaria vigente — el costo en negro se estima al valor que hoy paga la empresa por hora.",
        "Verificar con Contaduría antes de usarla como pago real; una vez liquidado, cargar el monto real del recibo en la app.",
      ],
      14, finalY + 18
    );
    doc.save(`Liquidacion_${(obra?.nombre || "obra").replace(/\s+/g, "_")}_${mesFormal}_Q${quincenaFormal}.pdf`);
  }

  // ---------- Tanteros (mano de obra por precio cerrado) ----------
  const [showTanteroForm, setShowTanteroForm] = useState(false);
  const emptyTanteroForm = { nombreGrupo: "", obraId: obras[0]?.id ?? "", precioTotal: "", integrantes: [], formalidad: FORMALIDADES[0] };
  const [tanteroForm, setTanteroForm] = useState(emptyTanteroForm);
  const [avanceAbiertoId, setAvanceAbiertoId] = useState(null);
  const [editandoAvanceId, setEditandoAvanceId] = useState(null);
  // La cuenta arranca vacía a propósito: si no se elige, el avance queda "sin
  // asignar" y no se descuenta de ningún saldo — mejor eso que adivinar mal de
  // dónde salió la plata. La formalidad no se elige acá: la define el grupo de
  // tanteros (un tantero es en blanco o en negro siempre, no pago por pago).
  const emptyAvanceForm = { fecha: hoyISO(), monto: "", descripcion: "", cuenta: "" };
  const [avanceForm, setAvanceForm] = useState(emptyAvanceForm);

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
      formalidad: tanteroForm.formalidad,
    }, setTanteros);
    setTanteroForm(emptyTanteroForm);
    setShowTanteroForm(false);
  }

  function pagadoDeTantero(tanteroId) {
    return avancesTanteros.filter((a) => a.tanteroId === tanteroId).reduce((s, a) => s + (a.monto || 0), 0);
  }

  function submitAvanceForm(e, tanteroId) {
    e.preventDefault();
    // La formalidad siempre es la del grupo — no se re-elige pago por pago.
    const formalidadGrupo = tanteros.find((t) => t.id === tanteroId)?.formalidad || null;
    const patch = {
      fecha: avanceForm.fecha,
      monto: Number(avanceForm.monto) || 0,
      descripcion: avanceForm.descripcion,
      cuenta: avanceForm.cuenta || null,
      formalidad: formalidadGrupo,
    };
    if (editandoAvanceId) {
      updateRecord("avances_tanteros", editandoAvanceId, patch, setAvancesTanteros);
    } else {
      addRecord("avances_tanteros", { tanteroId, ...patch }, setAvancesTanteros);
    }
    setAvanceForm(emptyAvanceForm);
    setAvanceAbiertoId(null);
    setEditandoAvanceId(null);
  }
  function iniciarEdicionAvance(avance) {
    setAvanceForm({
      fecha: avance.fecha,
      monto: avance.monto || "",
      descripcion: avance.descripcion || "",
      cuenta: avance.cuenta || "",
    });
    setEditandoAvanceId(avance.id);
    setAvanceAbiertoId(avance.tanteroId);
  }

  // ---------- Resumen de Cuentas (blanco/negro x efectivo/banco/MP) ----------
  const canVerFinanzas = ROLES_FINANZAS.includes(currentRole);
  // El capataz arma el pedido de obra sin ver precios ni proveedor — eso es cosa de Logística.
  const canVerPreciosPedido = ROLES_VEN_PRECIOS_PEDIDO.includes(currentRole);

  function saldoCuenta(cuenta, formalidad) {
    // Un ingreso "Pendiente" (todavía no cobrado) no suma al saldo hasta que se cobra.
    const totalIngresos = ingresos
      .filter((i) => i.cuenta === cuenta && i.formalidad === formalidad && i.estado !== "Pendiente" && !obraIdsPapelera.has(i.obraId))
      .reduce((s, i) => s + (i.monto || 0), 0);
    const totalEgresos = comprasFacturas
      .filter((c) => c.cuenta === cuenta && c.formalidad === formalidad && !obraIdsPapelera.has(c.obraId))
      .reduce((s, c) => s + (c.monto || 0), 0);
    // Cada transferencia manual resta en la cuenta de origen y suma en la de destino,
    // siempre dentro de la misma formalidad (blanco y negro nunca se mezclan).
    const totalManual = movimientosManual
      .filter((m) => m.formalidad === formalidad)
      .reduce((s, m) => s + (m.cuentaOrigen === cuenta ? -(m.monto || 0) : 0) + (m.cuentaDestino === cuenta ? (m.monto || 0) : 0), 0);
    // El capital de un préstamo entra a la cuenta como plata real (no es ganancia,
    // pero sí caja); cada devolución (parcial o la final) sale de la cuenta que se
    // eligió en ese pago — no necesariamente la misma en la que entró el capital.
    const totalPrestamos = prestamos
      .filter((p) => p.cuenta === cuenta && p.formalidad === formalidad)
      .reduce((s, p) => s + (p.capital || 0), 0);
    const totalPagosPrestamos = prestamosPagos
      .filter((pg) => pg.cuenta === cuenta && prestamos.find((p) => p.id === pg.prestamoId)?.formalidad === formalidad)
      .reduce((s, pg) => s + (pg.monto || 0), 0);
    // Un cobro de Ricardo o Pablo es plata real que sale de la caja de la empresa.
    const totalCobrosSocios = cobrosSocios
      .filter((c) => c.cuenta === cuenta && c.formalidad === formalidad)
      .reduce((s, c) => s + (c.monto || 0), 0);
    // Un avance pagado a un tantero es plata real que sale de la cuenta elegida.
    const totalAvancesTanteros = avancesTanteros
      .filter((a) => a.cuenta === cuenta && a.formalidad === formalidad)
      .reduce((s, a) => s + (a.monto || 0), 0);
    return totalIngresos - totalEgresos + totalManual + totalPrestamos - totalPagosPrestamos - totalCobrosSocios - totalAvancesTanteros;
  }

  const totalBlanco = FORMALIDADES[0] && CUENTAS.reduce((s, c) => s + saldoCuenta(c, "Blanco"), 0);
  const totalNegro = CUENTAS.reduce((s, c) => s + saldoCuenta(c, "Negro"), 0);

  // ---------- Préstamos (inversores/banco) ----------
  const emptyPrestamoForm = { fecha: hoyISO(), acreedor: "", capital: 0, tasaAnualPct: "", cuenta: CUENTAS[0], formalidad: FORMALIDADES[0], fechaEstimadaDevolucion: "" };
  const [prestamoForm, setPrestamoForm] = useState(emptyPrestamoForm);
  const [showPrestamoForm, setShowPrestamoForm] = useState(false);
  function submitPrestamoForm(e) {
    e.preventDefault();
    addRecord("prestamos", {
      ...prestamoForm,
      capital: Number(prestamoForm.capital) || 0,
      tasaAnualPct: Number(prestamoForm.tasaAnualPct) || 0,
      estado: "Vigente",
      fechaPago: null,
      montoPagado: null,
    }, setPrestamos);
    setPrestamoForm(emptyPrestamoForm);
    setShowPrestamoForm(false);
  }
  const [editandoPrestamoId, setEditandoPrestamoId] = useState(null);
  function guardarEdicionPrestamo(id, patch) {
    updateRecord("prestamos", id, patch, setPrestamos);
    setEditandoPrestamoId(null);
  }
  const [pagandoPrestamoId, setPagandoPrestamoId] = useState(null);
  // Un pago se auto-marca como devolución total (y desaparece de "Próximos
  // pagos/ingresos") en cuanto cubre el 100% del saldo de capital + interés —
  // no hace falta un botón aparte para "marcar devuelto".
  function guardarPagoPrestamo(prestamoId, pago) {
    addRecord("prestamos_pagos", { prestamoId, ...pago }, setPrestamosPagos);
    const p = prestamos.find((x) => x.id === prestamoId);
    const estadoConEsePago = calcularEstadoPrestamo(p, [...prestamosPagos, { prestamoId, ...pago }]);
    if (estadoConEsePago.saldoCapital < 1) {
      updateRecord("prestamos", prestamoId, { estado: "Pagado", fechaPago: pago.fecha, montoPagado: estadoConEsePago.totalPagado }, setPrestamos);
    }
    setPagandoPrestamoId(null);
  }

  // ---------- Cobros Ricardo y Pablo (retiros de los socios) ----------
  const SOCIOS = ["Ricardo", "Pablo"];
  const emptyCobroSocioForm = { fecha: hoyISO(), socio: SOCIOS[0], monto: 0, cuenta: CUENTAS[0], medioBancario: "Transferencia", formalidad: FORMALIDADES[0], tipoFactura: "Sin factura", archivo: null, nombreArchivo: null, tipoArchivo: null, observaciones: "" };
  const [cobroSocioForm, setCobroSocioForm] = useState(emptyCobroSocioForm);
  const [showCobroSocioForm, setShowCobroSocioForm] = useState(false);
  const [filtroSocio, setFiltroSocio] = useState("Todos");
  function submitCobroSocioForm(e) {
    e.preventDefault();
    addRecord("cobros_socios", {
      ...cobroSocioForm,
      monto: Number(cobroSocioForm.monto) || 0,
      medioBancario: cobroSocioForm.cuenta === "Banco" ? cobroSocioForm.medioBancario : null,
    }, setCobrosSocios);
    setCobroSocioForm(emptyCobroSocioForm);
    setShowCobroSocioForm(false);
  }
  function totalCobradoPorSocio(socio) {
    return cobrosSocios.filter((c) => c.socio === socio).reduce((s, c) => s + (c.monto || 0), 0);
  }
  const cobrosSociosFiltrados = cobrosSocios
    .filter((c) => filtroSocio === "Todos" || c.socio === filtroSocio)
    .sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));

  // "Registrar juntos": carga un solo total y lo parte a la mitad para cada socio,
  // pero queda guardado como dos cobros separados (uno por socio) en el historial —
  // y cada uno con SU PROPIA factura, porque aunque sea un solo retiro conjunto,
  // cada socio le factura a Concretar por separado.
  const emptyFacturaSocio = { tipoFactura: "Sin factura", archivo: null, nombreArchivo: null, tipoArchivo: null };
  const emptyCobroJuntosForm = {
    fecha: hoyISO(), monto: 0, cuenta: CUENTAS[0], medioBancario: "Transferencia", formalidad: FORMALIDADES[0], observaciones: "",
    facturas: { Ricardo: { ...emptyFacturaSocio }, Pablo: { ...emptyFacturaSocio } },
  };
  const [cobroJuntosForm, setCobroJuntosForm] = useState(emptyCobroJuntosForm);
  const [showCobroJuntosForm, setShowCobroJuntosForm] = useState(false);
  function setFacturaSocioJuntos(socio, patch) {
    setCobroJuntosForm((f) => ({ ...f, facturas: { ...f.facturas, [socio]: { ...f.facturas[socio], ...patch } } }));
  }
  async function submitCobroJuntosForm(e) {
    e.preventDefault();
    const total = Number(cobroJuntosForm.monto) || 0;
    const mitad = total / 2;
    for (const socio of SOCIOS) {
      const factura = cobroJuntosForm.facturas[socio];
      await addRecord("cobros_socios", {
        fecha: cobroJuntosForm.fecha,
        monto: mitad,
        cuenta: cobroJuntosForm.cuenta,
        medioBancario: cobroJuntosForm.cuenta === "Banco" ? cobroJuntosForm.medioBancario : null,
        formalidad: cobroJuntosForm.formalidad,
        observaciones: cobroJuntosForm.observaciones,
        socio,
        tipoFactura: factura.tipoFactura,
        archivo: factura.archivo,
        nombreArchivo: factura.nombreArchivo,
        tipoArchivo: factura.tipoArchivo,
      }, setCobrosSocios);
    }
    setCobroJuntosForm(emptyCobroJuntosForm);
    setShowCobroJuntosForm(false);
  }

  // Editar un gasto o un cobro de socio ya cargado (ej: para agregarle la factura
  // cuando todavía no la tenías al momento de cargarlo) — accesible tanto desde su
  // propia tabla como desde el ledger de Movimientos en Cuentas.
  const [editandoMovimiento, setEditandoMovimiento] = useState(null);
  function guardarEdicionCompra(id, patch) {
    updateRecord("compras_facturas", id, patch, setComprasFacturas);
    setEditandoMovimiento(null);
  }
  function guardarEdicionCobro(id, patch) {
    updateRecord("cobros_socios", id, patch, setCobrosSocios);
    setEditandoMovimiento(null);
  }
  function guardarEdicionIngreso(id, patch) {
    updateRecord("ingresos", id, patch, setIngresos);
    setEditandoMovimiento(null);
  }
  function guardarEdicionManual(id, patch) {
    updateRecord("movimientos_cuenta", id, patch, setMovimientosManual);
    setEditandoMovimiento(null);
  }
  function eliminarMovimiento(editando) {
    if (editando.origen === "compras_facturas") {
      const c = comprasFacturas.find((x) => x.id === editando.origenId);
      moverAPapelera("compras_facturas", editando.origenId, setComprasFacturas, `${c?.proveedor || "Gasto"} — ${fmtARS(c?.monto)}`);
    } else if (editando.origen === "cobros_socios") {
      const c = cobrosSocios.find((x) => x.id === editando.origenId);
      moverAPapelera("cobros_socios", editando.origenId, setCobrosSocios, `Cobro — ${c?.socio}`);
    } else if (editando.origen === "ingresos") {
      const i = ingresos.find((x) => x.id === editando.origenId);
      moverAPapelera("ingresos", editando.origenId, setIngresos, `${i?.concepto || "Ingreso"} — ${fmtARS(i?.monto)}`);
    } else {
      deleteRecord("movimientos_cuenta", editando.origenId, setMovimientosManual);
    }
    setEditandoMovimiento(null);
  }
  const [showMovimientoForm, setShowMovimientoForm] = useState(false);
  // Un movimiento por cada ingreso (+), compra/factura (-) y cada lado de una
  // transferencia manual (- en origen, + en destino); sumados por cuenta y
  // formalidad dan exactamente los saldos de arriba.
  const movimientosCuentas = [
    ...ingresos.filter((i) => !obraIdsPapelera.has(i.obraId)).map((i) => ({
      id: `ing-${i.id}`, fecha: i.fecha, tipo: "Ingreso", obraId: i.obraId, detalle: i.concepto, formalidad: i.formalidad, cuenta: i.cuenta, monto: i.monto || 0, estado: i.estado === "Pendiente" ? "Pendiente" : null,
      origen: "ingresos", origenId: i.id,
    })),
    ...comprasFacturas.filter((c) => !obraIdsPapelera.has(c.obraId)).map((c) => ({
      id: `egr-${c.id}`, fecha: c.fecha, tipo: "Egreso", obraId: c.obraId, detalle: c.proveedor, formalidad: c.formalidad, cuenta: c.cuenta, monto: -(c.monto || 0), estado: c.estado,
      origen: "compras_facturas", origenId: c.id, tipoFactura: c.tipoFactura,
    })),
    ...movimientosManual.flatMap((m) => [
      { id: `man-${m.id}-sale`, fecha: m.fecha, tipo: "Egreso", obraId: null, detalle: m.detalle || `Pase a ${m.cuentaDestino}`, formalidad: m.formalidad, cuenta: m.cuentaOrigen, monto: -(m.monto || 0), estado: null, origen: "movimientos_cuenta", origenId: m.id },
      { id: `man-${m.id}-recibe`, fecha: m.fecha, tipo: "Ingreso", obraId: null, detalle: m.detalle || `Pase desde ${m.cuentaOrigen}`, formalidad: m.formalidad, cuenta: m.cuentaDestino, monto: m.monto || 0, estado: null, origen: "movimientos_cuenta", origenId: m.id },
    ]),
    ...prestamos.map((p) => ({
      id: `prestamo-alta-${p.id}`, fecha: p.fecha, tipo: "Ingreso", obraId: null, detalle: `Préstamo recibido — ${p.acreedor}`, formalidad: p.formalidad, cuenta: p.cuenta, monto: p.capital || 0, estado: null,
    })),
    ...prestamosPagos.map((pg) => {
      const p = prestamos.find((x) => x.id === pg.prestamoId);
      return {
        id: `prestamo-pago-${pg.id}`, fecha: pg.fecha, tipo: "Egreso", obraId: null, detalle: `Devolución préstamo — ${p?.acreedor || "?"}`, formalidad: p?.formalidad, cuenta: pg.cuenta, monto: -(pg.monto || 0), estado: "Pagada",
      };
    }),
    ...cobrosSocios.map((c) => ({
      id: `cobro-socio-${c.id}`, fecha: c.fecha, tipo: "Egreso", obraId: null, detalle: `Cobro — ${c.socio}`, formalidad: c.formalidad, cuenta: c.cuenta, monto: -(c.monto || 0), estado: "Pagada",
      origen: "cobros_socios", origenId: c.id, tipoFactura: c.tipoFactura,
    })),
    ...avancesTanteros.flatMap((a) => {
      const t = tanteros.find((x) => x.id === a.tanteroId);
      if (!t || obraIdsPapelera.has(t.obraId)) return [];
      return [{
        id: `avance-tantero-${a.id}`, fecha: a.fecha, tipo: "Egreso", obraId: t.obraId, detalle: `Avance tantero — ${t.nombreGrupo}`, formalidad: a.formalidad, cuenta: a.cuenta, monto: -(a.monto || 0), estado: "Pagada",
      }];
    }),
  ].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));

  // Agrupados por mes — el mes actual siempre a la vista, los anteriores quedan
  // colapsados en pestañas desplegables para no alargar la pantalla.
  const claveMesCuentas = (fechaStr) => fechaStr ? fechaStr.slice(0, 7) : "";
  const nombreMesCuentas = (clave) => {
    const [y, m] = clave.split("-").map(Number);
    const nombre = new Date(y, m - 1, 1).toLocaleDateString("es-AR", { month: "long", year: "numeric" });
    return nombre.charAt(0).toUpperCase() + nombre.slice(1);
  };
  const mesActualClave = claveMesCuentas(hoyISO());
  const gruposMovimientos = [];
  for (const m of movimientosCuentas) {
    const clave = claveMesCuentas(m.fecha);
    let grupo = gruposMovimientos.find((g) => g.clave === clave);
    if (!grupo) { grupo = { clave, items: [] }; gruposMovimientos.push(grupo); }
    grupo.items.push(m);
  }
  const movimientosMesActual = gruposMovimientos.find((g) => g.clave === mesActualClave)?.items || [];
  const gruposMovimientosAnteriores = gruposMovimientos.filter((g) => g.clave !== mesActualClave);

  // ---------- Resumen por obra (balance de cada obra en curso) ----------
  // Sale de lo que ya tenemos cargado: precio acordado (obra.presupuesto), lo
  // presupuestado por rubro si se importó un Excel de presupuesto (presupuestoGeneral),
  // y lo efectivamente cobrado/gastado según Ingresos, Gastos/Facturas y mano de obra
  // (personal en negro, personal en blanco y tanteros) de esa obra.
  // Personal en negro: lo ya pagado usa el monto real abonado; lo todavía pendiente de
  // pago se estima con el costo por hora vigente, para que el gasto de la obra no se
  // quede atrasado esperando a que se liquide.
  // Personal en blanco: usa el costo real cargado por Contaduría en "Liquidación formal"
  // cuando ya está confirmado; si esa quincena todavía no se confirmó, se estima con las
  // mismas fórmulas (UOCRA) que usa esa pantalla.
  // "hasta" (Date) es opcional — sirve para cortar el cálculo a una fecha puntual (ej: fin
  // de un mes en la curva de inversión de la pestaña Obras). Sin ese parámetro da el total
  // acumulado a hoy, que es lo que usa el Balance por obra de Cuentas.
  function costoManoDeObraDeObra(obraId, hasta) {
    const dentroDePlazo = (fechaStr) => !hasta || fechaLocal(fechaStr) <= hasta;
    const asistenciaObra = asistencia.filter((a) => a.obraId === obraId && a.estado !== "Ausente" && (a.horas || 0) > 0 && dentroDePlazo(a.fecha));

    let costoNegro = 0;
    asistenciaObra
      .filter((a) => tipoTrabajadorDe(a.nombre) !== "En blanco")
      .forEach((a) => { costoNegro += a.estadoPago === "Pagado" ? (a.montoAbonado || 0) : montoDe(a); });

    let costoBlanco = 0;
    const gruposBlanco = {}; // "mes|quincena|nombre" -> horas trabajadas
    asistenciaObra
      .filter((a) => tipoTrabajadorDe(a.nombre) === "En blanco")
      .forEach((a) => {
        const mes = a.fecha.slice(0, 7);
        const quincena = quincenaDeFecha(a.fecha);
        const key = `${mes}|${quincena}|${a.nombre}`;
        if (!gruposBlanco[key]) gruposBlanco[key] = { mes, quincena, nombre: a.nombre, horas: 0 };
        gruposBlanco[key].horas += a.horas || 0;
      });
    Object.values(gruposBlanco).forEach((g) => {
      const registro = liquidacionesFormales.find((l) => l.obraId === obraId && l.mes === g.mes && l.quincena === g.quincena && l.nombre === g.nombre);
      if (registro?.costoRealBlanco != null) { costoBlanco += registro.costoRealBlanco; return; }
      const categoria = categoriaDe(g.nombre) || CATEGORIAS_PERSONAL[0];
      const horasRecibo = registro?.horasRecibo ?? Math.round((g.horas / 2) * 100) / 100;
      const presentismo = registro?.presentismo ?? false;
      const horasNegroDeBlanco = Math.max(0, g.horas - horasRecibo);
      const basicoHora = basicoConvenioDeCategoria(categoria, `${g.mes}-01`);
      const costoHoraInformal = costoHoraDeCategoria(categoria, `${g.mes}-01`);
      const montoBasico = horasRecibo * basicoHora;
      const montoPresentismo = presentismo ? montoBasico * ((cfgLiq.presentismoPct || 0) / 100) : 0;
      const bruto = montoBasico + montoPresentismo;
      const contribPct = (cfgLiq.contribObraSocialPct || 0) + (cfgLiq.contribEmpresariaPct || 0) + (cfgLiq.contribJubilacionPct || 0);
      const contribuciones = bruto * (contribPct / 100);
      const fondoCese = bruto * ((cfgLiq.fondoCesePosteriorPct || 0) / 100);
      const costoEmpresa = bruto + contribuciones + fondoCese + (cfgLiq.iericMontoFijo || 0);
      costoBlanco += costoEmpresa + horasNegroDeBlanco * costoHoraInformal;
    });

    const tanterosDeObra = new Set(tanteros.filter((t) => t.obraId === obraId).map((t) => t.id));
    const costoTanteros = avancesTanteros
      .filter((av) => tanterosDeObra.has(av.tanteroId) && dentroDePlazo(av.fecha))
      .reduce((s, av) => s + (av.monto || 0), 0);

    return costoNegro + costoBlanco + costoTanteros;
  }
  function fechaFinEstimada(obra) {
    if (!obra.inicio || !obra.meses) return null;
    const [y, m, d] = obra.inicio.split("-").map(Number);
    const fecha = new Date(y, m - 1 + obra.meses, d);
    const mm = String(fecha.getMonth() + 1).padStart(2, "0");
    const dd = String(fecha.getDate()).padStart(2, "0");
    return `${fecha.getFullYear()}-${mm}-${dd}`;
  }
  const resumenPorObra = obras
    .filter((o) => o.estado !== "Papelera")
    .map((o) => {
      const pg = presupuestoGeneral.find((p) => p.obraId === o.id);
      const presupuestadoManoObra = pg?.totalManoObra ?? null;
      const presupuestadoEqYMat = pg ? (pg.totalEquipos || 0) + (pg.totalMateriales || 0) + (pg.totalHerramientas || 0) : null;
      const precioObra = o.presupuesto || pg?.precioTotalConIva || 0;
      const cobrado = ingresos.filter((i) => i.obraId === o.id && i.estado !== "Pendiente").reduce((s, i) => s + (i.monto || 0), 0);
      const gastadoEqYMat = comprasFacturas.filter((c) => c.obraId === o.id).reduce((s, c) => s + (c.monto || 0), 0);
      const gastadoManoObra = costoManoDeObraDeObra(o.id);
      const gastado = gastadoEqYMat + gastadoManoObra;
      const disponibleEqYMat = presupuestadoEqYMat !== null ? presupuestadoEqYMat - gastadoEqYMat : null;
      const costoPresupuestado = pg ? presupuestadoManoObra + presupuestadoEqYMat : null;
      const gananciaEstimada = costoPresupuestado !== null ? precioObra - costoPresupuestado : null;
      return {
        obra: o,
        precioObra,
        cobrado,
        faltaCobrar: precioObra - cobrado,
        presupuestadoManoObra,
        presupuestadoEqYMat,
        gastadoEqYMat,
        gastadoManoObra,
        gastado,
        disponibleEqYMat,
        gananciaEstimada,
        porcentajeGanancia: gananciaEstimada !== null && precioObra ? gananciaEstimada / precioObra : null,
        dineroEnCaja: cobrado - gastado,
        inicio: o.inicio,
        finEstimado: fechaFinEstimada(o),
      };
    });

  // ---------- Próximos pagos/ingresos ----------
  const [showProximos, setShowProximos] = useState(false);
  const [mesProximosSeleccionado, setMesProximosSeleccionado] = useState(null);
  const prestamosPorDevolver = prestamos
    .filter((p) => p.estado !== "Pagado")
    .sort((a, b) => fechaLocal(a.fechaEstimadaDevolucion || a.fecha) - fechaLocal(b.fechaEstimadaDevolucion || b.fecha));
  // eCheqs de salida: los que nosotros libramos al pagarle a un proveedor.
  const echeqsSalida = comprasFacturas
    .filter((c) => c.formaPago === "eCheq" && c.estado === "Pendiente" && !obraIdsPapelera.has(c.obraId))
    .sort((a, b) => fechaLocal(a.fechaPagoEcheq) - fechaLocal(b.fechaPagoEcheq));
  // eCheqs de entrada: los que recibimos de un cliente y todavía no cobramos.
  const echeqsEntrada = ingresos
    .filter((i) => i.medioBancario === "eCheq" && i.estado === "Pendiente" && !obraIdsPapelera.has(i.obraId))
    .sort((a, b) => fechaLocal(a.fechaCobroEstimada || a.fecha) - fechaLocal(b.fechaCobroEstimada || b.fecha));
  // Las cuentas corrientes se pagan de una sola vez por proveedor, no factura por
  // factura — se agrupan y suman, y la fecha de vencimiento es del proveedor (cada
  // uno tiene la suya), no de cada compra.
  // El día de pago es mensual y recurrente (ej: "el 10 de cada mes"), así que la fecha
  // de vencimiento se recalcula sola todos los meses en vez de tener que cargarla a mano.
  function proximaFechaPago(diaPago) {
    if (!diaPago) return null;
    const hoy = new Date();
    let anio = hoy.getFullYear();
    let mes = hoy.getMonth();
    if (diaPago < hoy.getDate()) mes += 1;
    const ultimoDiaDelMes = new Date(anio, mes + 1, 0).getDate();
    const dia = Math.min(diaPago, ultimoDiaDelMes);
    const fecha = new Date(anio, mes, dia);
    const mm = String(fecha.getMonth() + 1).padStart(2, "0");
    const dd = String(fecha.getDate()).padStart(2, "0");
    return `${fecha.getFullYear()}-${mm}-${dd}`;
  }
  const cuentasCorrientesPorProveedor = (() => {
    const grupos = {};
    comprasFacturas
      .filter((c) => c.formaPago === "Cuenta corriente" && c.estado === "Pendiente" && !obraIdsPapelera.has(c.obraId))
      .forEach((c) => {
        if (!grupos[c.proveedor]) grupos[c.proveedor] = { proveedor: c.proveedor, monto: 0, cantidad: 0 };
        grupos[c.proveedor].monto += c.monto || 0;
        grupos[c.proveedor].cantidad += 1;
      });
    return Object.values(grupos)
      .map((g) => {
        const prov = proveedores.find((p) => nombreComercial(p) === g.proveedor);
        return {
          ...g,
          proveedorId: prov?.id ?? null,
          diaPago: prov?.diaPago || null,
          fechaVencimiento: prov?.diaPago ? proximaFechaPago(prov.diaPago) : (prov?.fechaVencimientoCC || null),
        };
      })
      .sort((a, b) => {
        if (!a.fechaVencimiento && !b.fechaVencimiento) return 0;
        if (!a.fechaVencimiento) return 1;
        if (!b.fechaVencimiento) return -1;
        return fechaLocal(a.fechaVencimiento) - fechaLocal(b.fechaVencimiento);
      });
  })();
  const ingresosPendientes = ingresos
    .filter((i) => i.estado === "Pendiente" && !obraIdsPapelera.has(i.obraId))
    .sort((a, b) => fechaLocal(a.fechaCobroEstimada || a.fecha) - fechaLocal(b.fechaCobroEstimada || b.fecha));
  function actualizarDiaPago(proveedorId, dia) {
    updateRecord("proveedores", proveedorId, { diaPago: dia ? Number(dia) : null }, setProveedores);
  }
  function marcarIngresoCobrado(ingreso) {
    updateRecord("ingresos", ingreso.id, { estado: "Cobrado" }, setIngresos);
  }

  // Obras "En curso": lo que todavía queda disponible para gastar (mano de obra + Eq. y
  // Mat. presupuestados, sin contar lo que ya se pasó de presupuesto) es plata que muy
  // probablemente se termine gastando antes de que termine la obra — se reparte en
  // partes iguales entre el mes actual y los meses que quedan de obra, como mínimo el
  // actual y el siguiente (si queda 1 mes o menos, igual se reparte en 2), para tener una
  // proyección aproximada de en qué mes vamos a necesitar esa plata.
  // Las obras "Pendiente de cobro" quedan afuera a propósito: ya están prácticamente sin
  // gastos por delante (entregadas, o falta algo mínimo) y sólo resta cobrarlas, así que no
  // tiene sentido seguir proyectándoles gasto disponible en los balances mensuales.
  function mesesRestantesDeObra(finEstimado) {
    if (!finEstimado) return 1;
    return Math.max(1, monthsBetween(fechaLocal(hoyISO()), fechaLocal(finEstimado)) + 1);
  }
  function sumarMesesAClave(clave, n) {
    const [y, m] = clave.split("-").map(Number);
    const d = new Date(y, m - 1 + n, 1);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }
  const obrasDisponibleProyectado = resumenPorObra
    .filter((r) => r.obra.estado === "En curso")
    .map((r) => {
      const disponibleManoObra = r.presupuestadoManoObra !== null ? Math.max(0, r.presupuestadoManoObra - r.gastadoManoObra) : 0;
      const disponibleEqYMat = r.disponibleEqYMat !== null ? Math.max(0, r.disponibleEqYMat) : 0;
      const disponibleTotal = disponibleManoObra + disponibleEqYMat;
      const mesesParaRepartir = Math.max(2, mesesRestantesDeObra(r.finEstimado));
      const montoPorMes = disponibleTotal / mesesParaRepartir;
      const meses = Array.from({ length: mesesParaRepartir }, (_, i) => sumarMesesAClave(mesActualClave, i));
      return { obra: r.obra, disponibleTotal, montoPorMes, meses };
    })
    .filter((x) => x.disponibleTotal > 0);

  // Agrupa todo lo de arriba (préstamos, cheques, cuentas corrientes de obras/proveedores,
  // ingresos y el disponible proyectado de obras) mes a mes, para la pantalla completa de
  // "Próximos pagos e ingresos" — lo que no tiene fecha estimada cargada queda en un grupo
  // aparte al final.
  const perteneceAMesProximos = (fecha, clave) => (clave === "sin-fecha" ? !fecha : !!fecha && claveMesCuentas(fecha) === clave);
  const gruposMesesProximos = (() => {
    const grupos = {};
    const agregar = (fecha, monto, flujo) => {
      const clave = fecha ? claveMesCuentas(fecha) : "sin-fecha";
      if (!grupos[clave]) grupos[clave] = { clave, ingresos: 0, egresos: 0 };
      grupos[clave][flujo === "ingreso" ? "ingresos" : "egresos"] += monto || 0;
    };
    prestamosPorDevolver.forEach((p) => agregar(p.fechaEstimadaDevolucion, calcularEstadoPrestamo(p, prestamosPagos).totalADevolver, "egreso"));
    echeqsSalida.forEach((c) => agregar(c.fechaPagoEcheq, c.monto, "egreso"));
    echeqsEntrada.forEach((i) => agregar(i.fechaCobroEstimada || i.fecha, i.monto, "ingreso"));
    cuentasCorrientesPorProveedor.forEach((g) => agregar(g.fechaVencimiento, g.monto, "egreso"));
    ingresosPendientes.forEach((i) => agregar(i.fechaCobroEstimada || i.fecha, i.monto, "ingreso"));
    obrasDisponibleProyectado.forEach((o) => o.meses.forEach((clave) => agregar(`${clave}-01`, o.montoPorMes, "egreso")));
    return Object.values(grupos).sort((a, b) => (a.clave === "sin-fecha" ? 1 : b.clave === "sin-fecha" ? -1 : a.clave.localeCompare(b.clave)));
  })();
  // El acumulado arranca de la plata que hay hoy en las cuentas (Blanco + Negro) y le va
  // sumando el total (ingreso - egreso) de cada mes en orden — así se ve en qué mes, si
  // se cumplen estas fechas estimadas, la empresa se quedaría sin plata (acumulado en rojo).
  // Lo "sin fecha" no entra en la cuenta porque no se sabe cuándo va a pasar.
  const saldoActualTotal = totalBlanco + totalNegro;
  let acumuladoProximosRunning = saldoActualTotal;
  const gruposMesesProximosConAcumulado = gruposMesesProximos.map((m) => {
    if (m.clave === "sin-fecha") return { ...m, acumulado: null };
    acumuladoProximosRunning += m.ingresos - m.egresos;
    return { ...m, acumulado: acumuladoProximosRunning };
  });

  // ---------- Dinero real (arqueo de caja) ----------
  // La plata física no distingue blanco de negro, por eso "Dinero real" se compara
  // contra el Total (Blanco + Negro) de cada cuenta, no contra cada columna por separado.
  function dineroRealDe(cuenta) {
    return dineroReal.find((d) => d.cuenta === cuenta)?.monto ?? null;
  }
  function actualizarDineroReal(cuenta, monto) {
    const existente = dineroReal.find((d) => d.cuenta === cuenta);
    if (existente) {
      updateRecord("dinero_real_cuentas", existente.id, { monto, actualizado: hoyISO() }, setDineroReal);
    } else {
      addRecord("dinero_real_cuentas", { cuenta, monto, actualizado: hoyISO() }, setDineroReal);
    }
  }
  // Un "Error de cálculo" nunca es transferencia real entre nuestras cuentas: usamos
  // "Ajuste" como cuenta puente (no forma parte de CUENTAS, así que no aparece en el
  // resumen) solo para poder reutilizar el mecanismo de movimientos y que la cuenta
  // real quede en el número contado a mano. Banco y Mercado Pago son siempre blancos
  // (no hay plata en negro ahí); en Efectivo, que mezcla las dos, el error se carga
  // como negro.
  async function arreglarCaja() {
    let corregidas = 0;
    for (const cuenta of CUENTAS) {
      const real = dineroRealDe(cuenta);
      if (real === null) continue;
      const calculado = saldoCuenta(cuenta, "Blanco") + saldoCuenta(cuenta, "Negro");
      const diferencia = real - calculado;
      if (Math.abs(diferencia) < 1) continue;
      const formalidadAjuste = cuenta === "Banco" || cuenta === "Mercado Pago" ? "Blanco" : "Negro";
      await addRecord("movimientos_cuenta", {
        fecha: hoyISO(),
        detalle: "Error de cálculo",
        formalidad: formalidadAjuste,
        cuentaOrigen: diferencia > 0 ? "Ajuste" : cuenta,
        cuentaDestino: diferencia > 0 ? cuenta : "Ajuste",
        monto: Math.abs(diferencia),
      }, setMovimientosManual);
      corregidas++;
    }
    alert(corregidas === 0 ? "No hay diferencias entre lo calculado y el dinero real cargado." : `Se corrigieron ${corregidas} cuenta(s) con un movimiento "Error de cálculo".`);
  }

  const [filtroHerr, setFiltroHerr] = useState({ ubicacion: "Todas", estado: "Todos" });

  // N° de serie automático: letra de tipo + 3 letras de marca + N° correlativo (ej: E-BOS01)
  function generarNumeroSerie(categoria, marca) {
    const letraTipo = LETRA_TIPO_HERRAMIENTA[categoria] || "X";
    const letrasMarca = (marca || "GEN").replace(/[^a-zA-Zñ]/g, "").toUpperCase().padEnd(3, "X").slice(0, 3);
    const prefijo = `${letraTipo}-${letrasMarca}`;
    // Cuenta sobre TODAS las herramientas (incluidas las que están en la Papelera,
    // que siguen existiendo en la base hasta que se purguen a los 7 días) para no
    // repetirle el mismo número de serie a una herramienta nueva.
    const existentes = herramientasRaw.filter((h) => (h.numeroSerie || "").startsWith(prefijo));
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
    scrollContenidoArriba();
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
  const emptyProveedorForm = { razonSocial: "", nombreFantasia: "", cuit: "", domicilio: "", contacto: "", telefono: "", email: "", cbu: "", numeroCuenta: "", esTaller: "No", diaPago: "" };
  const [proveedorForm, setProveedorForm] = useState(emptyProveedorForm);
  const [showProveedorForm, setShowProveedorForm] = useState(false);
  const [editandoProveedorId, setEditandoProveedorId] = useState(null);
  const talleres = proveedores.filter((p) => p.esTaller === "Sí");

  function submitProveedorForm(e) {
    e.preventDefault();
    const datos = { ...proveedorForm, diaPago: proveedorForm.diaPago ? Number(proveedorForm.diaPago) : null };
    if (editandoProveedorId) {
      updateRecord("proveedores", editandoProveedorId, datos, setProveedores);
    } else {
      addRecord("proveedores", datos, setProveedores);
    }
    setProveedorForm(emptyProveedorForm);
    setEditandoProveedorId(null);
    setShowProveedorForm(false);
  }
  function editarProveedor(p) {
    setProveedorForm({
      razonSocial: p.razonSocial || "",
      nombreFantasia: p.nombreFantasia || "",
      cuit: p.cuit || "",
      diaPago: p.diaPago || "",
      domicilio: p.domicilio || "",
      contacto: p.contacto || "",
      telefono: p.telefono || "",
      email: p.email || "",
      cbu: p.cbu || "",
      numeroCuenta: p.numeroCuenta || "",
      esTaller: p.esTaller || "No",
    });
    setEditandoProveedorId(p.id);
    setShowProveedorForm(true);
  }
  // El nombre de fantasía es el que se usa para elegir el proveedor en Compras y en
  // Órdenes de Compra; si todavía no se cargó, se usa la razón social.
  async function crearProveedorRapido(nombre) {
    return await addRecord("proveedores", { ...emptyProveedorForm, razonSocial: nombre, nombreFantasia: nombre, diaPago: null }, setProveedores);
  }
  function balanceProveedor(prov) {
    const facturas = comprasFacturas.filter((c) => c.proveedor === nombreComercial(prov) && !obraIdsPapelera.has(c.obraId));
    const totalFacturado = facturas.reduce((s, c) => s + (c.monto || 0), 0);
    const totalPagado = facturas.filter((c) => c.estado === "Pagada").reduce((s, c) => s + (c.monto || 0), 0);
    return { totalFacturado, totalPagado, saldo: totalFacturado - totalPagado, facturasPendientes: facturas.filter((c) => c.estado !== "Pagada") };
  }
  function marcarFacturaPagada(factura) {
    updateRecord("compras_facturas", factura.id, { estado: "Pagada" }, setComprasFacturas);
  }
  // Un eCheq queda "Pendiente" hasta su fecha de pago; llegado ese día se acredita solo,
  // sin que nadie tenga que entrar a marcarlo a mano.
  useEffect(() => {
    if (dbLoading) return;
    const hoy = hoyISO();
    comprasFacturas
      .filter((c) => (c.formaPago === "eCheq" || c.medioBancario === "eCheq") && c.estado === "Pendiente" && c.fechaPagoEcheq && c.fechaPagoEcheq <= hoy)
      .forEach((c) => updateRecord("compras_facturas", c.id, { estado: "Pagada" }, setComprasFacturas));
  }, [dbLoading, comprasFacturas]);

  // ---------- Clientes ----------
  const emptyClienteForm = { razonSocial: "", nombreFantasia: "", cuit: "", domicilio: "", contacto: "", telefono: "", email: "", cbu: "", numeroCuenta: "" };
  const [clienteForm, setClienteForm] = useState(emptyClienteForm);
  const [showClienteForm, setShowClienteForm] = useState(false);
  const [editandoClienteId, setEditandoClienteId] = useState(null);
  const [vistaClientesProveedores, setVistaClientesProveedores] = useState("clientes");

  function submitClienteForm(e) {
    e.preventDefault();
    if (editandoClienteId) {
      updateRecord("clientes", editandoClienteId, { ...clienteForm }, setClientes);
    } else {
      addRecord("clientes", { ...clienteForm }, setClientes);
    }
    setClienteForm(emptyClienteForm);
    setEditandoClienteId(null);
    setShowClienteForm(false);
  }
  function editarCliente(cli) {
    setClienteForm({
      razonSocial: cli.razonSocial || "",
      nombreFantasia: cli.nombreFantasia || "",
      cuit: cli.cuit || "",
      domicilio: cli.domicilio || "",
      contacto: cli.contacto || "",
      telefono: cli.telefono || "",
      email: cli.email || "",
      cbu: cli.cbu || "",
      numeroCuenta: cli.numeroCuenta || "",
    });
    setEditandoClienteId(cli.id);
    setShowClienteForm(true);
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
    scrollContenidoArriba();
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
      // Todo pedido armado en una obra pasa automáticamente a Órdenes de Compra —
      // ahí lo aprueba, rechaza o modifica Gerencia (las compras generales, sin obra,
      // se siguen aprobando directo en Pedidos de Obra).
      if (pedidoCreado.obraId != null) {
        await addRecord("ordenes_compra", {
          fecha: hoyISO(),
          obraId: pedidoCreado.obraId,
          proveedor: pedidoProveedor || "Sin especificar",
          item: itemsFinales.map((it) => it.material).join(", "),
          montoEstimado: totalPedido,
          estado: "Requiere aprobación",
          pedidoId: pedidoCreado.id,
        }, setOrdenesCompra);
      }
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
  async function aprobarPedidoMaterial(pedido) {
    await updateRecord("pedidos_materiales", pedido.id, { estado: "Aprobado", aprobadoPor: currentRole, fechaAprobacion: hoyISO() }, setPedidosMateriales);
  }
  async function rechazarPedidoMaterial(pedido, observaciones = "") {
    await updateRecord("pedidos_materiales", pedido.id, { estado: "Rechazado", aprobadoPor: currentRole, fechaAprobacion: hoyISO(), observaciones }, setPedidosMateriales);
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

  // Las órdenes armadas desde un pedido de obra (oc.pedidoId) reflejan la decisión
  // también en el pedido, para que Pedidos de Obra muestre el mismo estado.
  async function aprobarOC(oc) {
    await updateRecord("ordenes_compra", oc.id, { estado: "Aprobada" }, setOrdenesCompra);
    if (oc.pedidoId) {
      const pedido = pedidosMateriales.find((p) => p.id === oc.pedidoId);
      if (pedido) await aprobarPedidoMaterial(pedido);
    }
  }
  async function rechazarOC(oc, observaciones) {
    await updateRecord("ordenes_compra", oc.id, { estado: "Rechazada", observaciones }, setOrdenesCompra);
    if (oc.pedidoId) {
      const pedido = pedidosMateriales.find((p) => p.id === oc.pedidoId);
      if (pedido) await rechazarPedidoMaterial(pedido, observaciones);
    }
  }
  const recibirOC = (id) => updateRecord("ordenes_compra", id, { estado: "Recibida" }, setOrdenesCompra);

  const [editandoOcId, setEditandoOcId] = useState(null);
  const [ocEditDraft, setOcEditDraft] = useState(null);
  const [rechazandoOcId, setRechazandoOcId] = useState(null);
  const [observacionesRechazoOc, setObservacionesRechazoOc] = useState("");
  function iniciarEdicionOc(oc) {
    setEditandoOcId(oc.id);
    setOcEditDraft({ proveedor: oc.proveedor || "", item: oc.item || "", montoEstimado: oc.montoEstimado || 0 });
    scrollContenidoArriba();
  }
  function cancelarEdicionOc() {
    setEditandoOcId(null);
    setOcEditDraft(null);
  }
  async function guardarEdicionOc(oc) {
    await updateRecord("ordenes_compra", oc.id, { ...ocEditDraft }, setOrdenesCompra);
    cancelarEdicionOc();
  }
  function iniciarRechazoOc(oc) {
    setRechazandoOcId(oc.id);
    setObservacionesRechazoOc("");
    scrollContenidoArriba();
  }
  function cancelarRechazoOc() {
    setRechazandoOcId(null);
    setObservacionesRechazoOc("");
  }
  async function confirmarRechazoOc(oc) {
    if (!observacionesRechazoOc.trim()) {
      alert("Escribí el motivo del rechazo en observaciones, para que Logística sepa por qué.");
      return;
    }
    await rechazarOC(oc, observacionesRechazoOc.trim());
    cancelarRechazoOc();
  }

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
                tab === item.id
                  ? "text-white"
                  : NAV_DESTACADOS.includes(item.id)
                  ? "text-red-400 hover:bg-white/5 hover:text-red-300"
                  : "text-slate-300 hover:bg-white/5 hover:text-white"
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
      <main ref={mainRef} className="mt-12 flex-1 overflow-y-auto p-4 md:mt-0 md:p-8">
        {tab === "general" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">General</h2>

            <Panel title="Alertas de todas las obras">
              {totalAlertas === 0 ? (
                <div className="flex items-center gap-2 text-sm text-emerald-700"><CheckCircle2 size={16} /> Todo en orden, sin pendientes críticos.</div>
              ) : (
                <div className="grid grid-cols-1 gap-3 lg:grid-cols-2 xl:grid-cols-3">
                  {ocPendientesAprobacion.length > 0 && (
                    <AlertCard tone="rose" icon={AlertTriangle} title={`${ocPendientesAprobacion.length} orden(es) de compra esperando aprobación.`} />
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
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex rounded-full border border-stone-300 bg-white p-0.5">
                  <button
                    onClick={() => setVistaObras("lista")}
                    className={`flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-bold ${vistaObras === "lista" ? "bg-amber-500 text-slate-900" : "text-slate-500 hover:bg-stone-50"}`}
                  >
                    <LayoutDashboard size={13} /> Lista
                  </button>
                  <button
                    onClick={() => setVistaObras("planificacion")}
                    className={`flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-bold ${vistaObras === "planificacion" ? "bg-amber-500 text-slate-900" : "text-slate-500 hover:bg-stone-50"}`}
                  >
                    <CalendarClock size={13} /> Planificación
                  </button>
                </div>
                <button onClick={() => setShowObraForm((v) => !v)} className={btnPrimary}>
                  <Plus size={16} /> Nueva obra
                </button>
              </div>
            </div>

            {vistaObras === "planificacion" ? (
              <PlanificacionObras
                obras={obras.filter((o) => o.estado !== "Papelera")}
                etapas={etapasObra}
                agregandoEtapaObraId={agregandoEtapaObraId}
                setAgregandoEtapaObraId={setAgregandoEtapaObraId}
                editandoEtapaId={editandoEtapaId}
                setEditandoEtapaId={setEditandoEtapaId}
                onAgregarEtapa={agregarEtapa}
                onGuardarEdicionEtapa={guardarEdicionEtapa}
                onEliminarEtapa={eliminarEtapa}
              />
            ) : (
              <>
            <div className="flex flex-wrap gap-2">
              {[
                { estado: "En curso", label: "Activas", activeCls: "border-amber-500 bg-amber-500 text-slate-900", idleCls: "border-amber-300 text-amber-700 hover:bg-amber-50" },
                { estado: "Pendiente de cobro", label: "Pendientes de cobro", activeCls: "border-sky-600 bg-sky-600 text-white", idleCls: "border-sky-300 text-sky-700 hover:bg-sky-50" },
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
                        {clientes.map((c) => <option key={c.id} value={c.id}>{nombreComercial(c)}{c.nombreFantasia && c.nombreFantasia.trim() && c.nombreFantasia.trim() !== c.razonSocial ? ` (${c.razonSocial})` : ""}</option>)}
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
                  <div
                    key={o.id}
                    onClick={editandoObraId === o.id ? undefined : () => abrirObra(o)}
                    className={`relative overflow-hidden rounded-lg border border-stone-200 p-5 pl-6 shadow-sm ${editandoObraId === o.id ? "" : "cursor-pointer hover:border-amber-300 hover:shadow-md"}`}
                    style={{ backgroundColor: `${colorDeObra(o)}0d` }}
                  >
                    <div className="absolute inset-y-0 left-0 w-1.5" style={{ backgroundColor: colorDeObra(o) }} />
                    {editandoObraId === o.id ? (
                      <form onSubmit={(e) => guardarEdicionObra(e, o)} className="space-y-3">
                        <div className="text-sm font-semibold text-slate-900">Editar {o.nombre}</div>
                        <Field label="Nombre"><input name="nombre" defaultValue={o.nombre} required className={inputCls} /></Field>
                        <Field label="Cliente">
                          <select name="clienteId" defaultValue={o.clienteId || ""} className={inputCls}>
                            <option value="">Sin conectar (texto: {o.cliente || "—"})</option>
                            {clientes.map((c) => <option key={c.id} value={c.id}>{nombreComercial(c)}{c.nombreFantasia && c.nombreFantasia.trim() && c.nombreFantasia.trim() !== c.razonSocial ? ` (${c.razonSocial})` : ""}</option>)}
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
                            <button onClick={() => abrirObra(o)} className="flex items-center gap-1.5 text-left font-bold text-slate-900 underline decoration-dotted hover:text-amber-600">
                              <ObraDot obra={o} size={9} />{o.nombre}
                            </button>
                            <div className="text-sm text-slate-500">{o.cliente}{!o.clienteId && <span className="ml-1 text-amber-600">(sin conectar)</span>}</div>
                          </div>
                          <div className="flex items-center gap-1.5" onClick={(e) => e.stopPropagation()}>
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
              </>
            )}
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

              // Historial de gastos de esta obra — mismo criterio que el "Gastado" del
              // Balance por obra (Gastos y Facturas + mano de obra), pero acá se ve cada
              // movimiento individual, del más reciente al más antiguo.
              const historialGastosObra = [];
              gastosObra.forEach((g) => {
                historialGastosObra.push({
                  id: `gasto-${g.id}`,
                  fecha: g.fecha,
                  tipo: "Gastos y Facturas",
                  detalle: [g.categoria, g.proveedor, g.descripcion].filter(Boolean).join(" — "),
                  monto: g.monto || 0,
                });
              });
              const tanterosDeObraSel = tanteros.filter((t) => t.obraId === obraSel.id);
              avancesTanteros
                .filter((av) => tanterosDeObraSel.some((t) => t.id === av.tanteroId))
                .forEach((av) => {
                  const t = tanterosDeObraSel.find((x) => x.id === av.tanteroId);
                  historialGastosObra.push({
                    id: `tantero-${av.id}`,
                    fecha: av.fecha,
                    tipo: "Tantero",
                    detalle: [t?.nombreGrupo, av.descripcion].filter(Boolean).join(" — "),
                    monto: av.monto || 0,
                  });
                });
              const pagosNegroObraSel = {}; // "fechaPago|nombre" -> monto abonado
              asistencia
                .filter((a) => a.obraId === obraSel.id && a.estadoPago === "Pagado" && tipoTrabajadorDe(a.nombre) !== "En blanco")
                .forEach((a) => {
                  const key = `${a.fechaPago}|${a.nombre}`;
                  if (!pagosNegroObraSel[key]) pagosNegroObraSel[key] = { fecha: a.fechaPago, nombre: a.nombre, monto: 0 };
                  pagosNegroObraSel[key].monto += a.montoAbonado || 0;
                });
              Object.values(pagosNegroObraSel).forEach((p) => {
                historialGastosObra.push({ id: `negro-${p.fecha}-${p.nombre}`, fecha: p.fecha, tipo: "Personal", detalle: p.nombre, monto: p.monto });
              });
              liquidacionesFormales
                .filter((l) => l.obraId === obraSel.id && l.costoRealBlanco != null)
                .forEach((l) => {
                  historialGastosObra.push({
                    id: `blanco-${l.id}`,
                    fecha: l.fechaConfirmacion || rangoQuincena(l.mes, l.quincena).hasta,
                    tipo: "Personal (blanco)",
                    detalle: `${l.nombre} — ${etiquetaQuincena(l.mes, l.quincena)}`,
                    monto: l.costoRealBlanco,
                  });
                });
              historialGastosObra.sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha));
              const resumenObraSel = resumenPorObra.find((r) => r.obra.id === obraSel.id);

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

                  <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                    <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Herramientas en uso</div>
                    <div className="mt-1 font-mono text-lg font-bold text-slate-900">{herramientasEnUso}</div>
                  </div>

                  <Panel title="Balance de la obra">
                    {(() => {
                      const celda = (v) => (v === null || v === undefined ? "—" : fmtARS(v));
                      const r = resumenObraSel;
                      const item = (label, val, tono) => (
                        <div>
                          <div className="text-[10px] uppercase tracking-wide text-slate-400">{label}</div>
                          <div className={`font-mono text-xs font-semibold ${tono || "text-slate-800"}`}>{celda(val)}</div>
                        </div>
                      );
                      return (
                        <div className="grid grid-cols-2 gap-x-3 gap-y-2 sm:grid-cols-4">
                          {item("Precio de obra", r.precioObra)}
                          {item("Falta cobrar", r.faltaCobrar)}
                          {item("Presup. M.O.", r.presupuestadoManoObra)}
                          {item("Pagado M.O.", r.gastadoManoObra)}
                          {item("Presup. Eq. y Mat.", r.presupuestadoEqYMat)}
                          {item("Gastado Eq. y Mat.", r.gastadoEqYMat)}
                          {item("Gastado total", r.gastado)}
                          {item("Dinero en caja", r.dineroEnCaja, r.dineroEnCaja < 0 ? "text-rose-600" : "text-emerald-700")}
                          {item("Ganancia estimada", r.gananciaEstimada, r.gananciaEstimada !== null && r.gananciaEstimada < 0 ? "text-rose-600" : "text-emerald-700")}
                          <div>
                            <div className="text-[10px] uppercase tracking-wide text-slate-400">% Ganancia</div>
                            <div className="font-mono text-xs font-semibold text-slate-800">{r.porcentajeGanancia === null ? "—" : `${Math.round(r.porcentajeGanancia * 100)}%`}</div>
                          </div>
                        </div>
                      );
                    })()}
                  </Panel>

                  <Panel title="Historial de gastos">
                    <HistorialGastosObra items={historialGastosObra} />
                  </Panel>

                  <div className="flex flex-wrap gap-2">
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
                <button
                  onClick={() => setShowSegurosPanel((v) => !v)}
                  className={`flex items-center gap-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                    showSegurosPanel ? "border-slate-400 bg-stone-100 text-slate-700" : "border-stone-300 bg-white text-slate-700 hover:bg-stone-50"
                  }`}
                >
                  <Shield size={16} /> Seguros/ART
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
                    className={btnPrimary}
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
                  <table className="w-full text-left text-xs">
                    <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                      <tr>
                        <th className="sticky left-0 bg-stone-50 px-3 py-1.5">Categoría</th>
                        {mesesCostos.map((mes) => <th key={mes} className="px-2 py-1.5 text-right capitalize">{nombreMes(mes)}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {CATEGORIAS_PERSONAL.map((categoria) => (
                        <tr key={categoria} className="border-t border-stone-100">
                          <td className="sticky left-0 bg-white px-3 py-1 font-medium text-slate-900">{categoria}</td>
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

            {showSegurosPanel && (
              <Panel title="Seguros/ART" action={<button onClick={() => setShowSegurosPanel(false)}><X size={16} /></button>}>
                {(() => {
                  const activos = personal.filter((p) => p.estado === "Activo");
                  const sinSeguro = activos.filter((p) => (p.aseguradoPor || "No") === "No");
                  const conArt = activos.filter((p) => p.aseguradoPor === "ART");
                  const conAccidentes = activos.filter((p) => p.aseguradoPor === "Seg. Accidentes");
                  return (
                    <>
                      <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-3">
                        <div className="rounded-lg border border-stone-200 bg-white p-3">
                          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">ART</div>
                          <div className="mt-1 font-mono text-lg font-bold text-slate-900">{conArt.length}</div>
                        </div>
                        <div className="rounded-lg border border-stone-200 bg-white p-3">
                          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Seg. Accidentes</div>
                          <div className="mt-1 font-mono text-lg font-bold text-slate-900">{conAccidentes.length}</div>
                        </div>
                        <div className={`rounded-lg border p-3 ${sinSeguro.length > 0 ? "border-rose-300 bg-rose-50" : "border-stone-200 bg-white"}`}>
                          <div className={`text-[11px] font-semibold uppercase tracking-wide ${sinSeguro.length > 0 ? "text-rose-700" : "text-slate-500"}`}>Sin seguro</div>
                          <div className={`mt-1 font-mono text-lg font-bold ${sinSeguro.length > 0 ? "text-rose-700" : "text-slate-900"}`}>{sinSeguro.length}</div>
                        </div>
                      </div>
                      {activos.length === 0 ? (
                        <div className="text-xs text-slate-400">No hay personal activo cargado.</div>
                      ) : (
                        <div className="overflow-x-auto rounded-lg border border-stone-200">
                          <table className="w-full text-left text-xs">
                            <thead className="bg-stone-50 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
                              <tr>
                                <th className="px-2 py-1.5">Persona</th>
                                <th className="px-2 py-1.5">Categoría</th>
                                <th className="px-2 py-1.5">Asegurado por</th>
                              </tr>
                            </thead>
                            <tbody>
                              {[...activos]
                                .sort((a, b) => {
                                  const orden = { No: 0, ART: 1, "Seg. Accidentes": 2 };
                                  const da = orden[a.aseguradoPor || "No"] ?? 0;
                                  const db = orden[b.aseguradoPor || "No"] ?? 0;
                                  return da - db || nombreCompletoDe(a).localeCompare(nombreCompletoDe(b));
                                })
                                .map((p) => (
                                  <tr key={p.id} className="border-t border-stone-100">
                                    <td className="px-2 py-1 font-medium text-slate-900">{nombreCompletoDe(p)}</td>
                                    <td className="px-2 py-1 text-slate-600">{p.categoria}</td>
                                    <td className="px-2 py-1">
                                      <span
                                        className={`inline-block rounded-full border-2 px-2 py-0.5 text-[10px] font-bold uppercase tracking-widest ${
                                          (p.aseguradoPor || "No") === "No"
                                            ? "border-rose-600 text-rose-700"
                                            : p.aseguradoPor === "ART"
                                            ? "border-sky-600 text-sky-700"
                                            : "border-emerald-600 text-emerald-700"
                                        }`}
                                      >
                                        {p.aseguradoPor || "No"}
                                      </span>
                                    </td>
                                  </tr>
                                ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </>
                  );
                })()}
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
                  <Field label="Asegurado por">
                    <select value={personalForm.aseguradoPor} onChange={(e) => pf("aseguradoPor")(e.target.value)} className={inputCls}>
                      {ASEGURADO_POR.map((a) => <option key={a}>{a}</option>)}
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
                        moverAPapelera("personal", viewingPerson.id, setPersonal, nombreCompletoDe(viewingPerson));
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
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Tipo de trabajador</div><div className="text-slate-800">{viewingPerson.tipoTrabajador === "Empresa" ? "En negro" : (viewingPerson.tipoTrabajador || "En negro")}</div></div>
                <div><div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Asegurado por</div><div className={(viewingPerson.aseguradoPor || "No") === "No" ? "font-semibold text-rose-600" : "text-slate-800"}>{viewingPerson.aseguradoPor || "No"}</div></div>
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

              {historialPagosPersona.length > 0 && (
                <div className="mt-4">
                  <div className="mb-1.5 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Historial de pagos</div>
                  <div className="overflow-x-auto rounded-lg border border-stone-200">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-stone-50 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
                        <tr>
                          <th className="px-2 py-1.5">Período</th>
                          <th className="px-2 py-1.5">Estado</th>
                          <th className="px-2 py-1.5 text-right">Ganó</th>
                        </tr>
                      </thead>
                      <tbody>
                        {historialPagosPersona.map((h, i) => (
                          <tr key={i} className="border-t border-stone-100">
                            <td className="px-2 py-1 text-slate-700">{h.periodo}</td>
                            <td className="px-2 py-1"><Badge estado={h.tipo} /></td>
                            <td className="px-2 py-1 text-right font-mono font-semibold text-slate-800">{fmtARS(h.monto)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {tab === "asistencia" && (
          <div className="space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Asistencia</h2>
              <button onClick={abrirCargaAsistencia} className={btnPrimary}>
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

              </Panel>
            )}

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
              </Panel>
            )}

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-xs">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-2 py-1.5">Fecha</th><th className="px-2 py-1.5">Nombre</th><th className="px-2 py-1.5">Centro de costo</th><th className="px-2 py-1.5">Hs</th><th className="px-2 py-1.5">Estado</th><th className="px-2 py-1.5">Cargado por</th><th className="px-2 py-1.5"></th></tr>
                </thead>
                <tbody>
                  {[...asistencia].sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((a) => {
                    const obra = obras.find((o) => o.id === a.obraId);
                    return (
                      <tr key={a.id} className="border-t border-stone-100">
                        <td className="px-2 py-1 text-slate-600">{fmtFecha(a.fecha)}</td>
                        <td className="px-2 py-1 font-medium text-slate-900">
                          <span className="flex items-center gap-1.5">
                            {a.nombre}
                            {a.editado && <span title={`Editado por ${a.editadoPor}: ${a.motivoEdicion}`}><AlertTriangle size={12} className="text-sky-500" /></span>}
                          </span>
                        </td>
                        <td className="px-2 py-1 text-slate-600">{obra?.nombre}</td>
                        <td className="px-2 py-1 font-mono text-slate-700">{a.horas}</td>
                        <td className="px-2 py-1"><Badge estado={a.estado} /></td>
                        <td className="px-2 py-1 text-slate-500">{a.cargadoPor}</td>
                        <td className="px-2 py-1"><button onClick={() => startEditAsistencia(a)} className={btnGhost}>Editar</button></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
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
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Salario Personal</h2>

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
              <button
                onClick={() => setVistaLiquidacion("formal")}
                className={`rounded-md px-3 py-2 text-sm font-semibold ${vistaLiquidacion === "formal" ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
              >
                Liquidación formal (UOCRA)
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

                {pendientesEnBlanco.length > 0 && (
                  <div>
                    <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Personal en blanco — pendiente de liquidación formal</div>
                    <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-stone-50 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
                          <tr>
                            <th className="px-2 py-1.5">Obra</th>
                            <th className="px-2 py-1.5">Quincena</th>
                            <th className="px-2 py-1.5">Personas</th>
                            <th className="px-2 py-1.5 text-right">Hs.</th>
                            <th className="px-2 py-1.5">Estado</th>
                            <th className="px-2 py-1.5"></th>
                          </tr>
                        </thead>
                        <tbody>
                          {pendientesEnBlanco.map((g) => {
                            const obra = obras.find((o) => o.id === g.obraId);
                            const diasParaVencer = 5 - g.diasDesdeCierre;
                            return (
                              <tr key={`${g.obraId}|${g.mes}|${g.quincena}`} className={`border-t border-stone-100 ${g.vencido ? "bg-rose-50" : ""}`}>
                                <td className="px-2 py-1 font-medium text-slate-900">{obra?.nombre || "Obra"}</td>
                                <td className="px-2 py-1 text-slate-600">{etiquetaQuincena(g.mes, g.quincena)}</td>
                                <td className="px-2 py-1 text-slate-600">{g.personas.join(", ")}</td>
                                <td className="px-2 py-1 text-right font-mono text-slate-700">{g.horas}</td>
                                <td className={`px-2 py-1 font-semibold ${g.vencido ? "text-rose-600" : g.diasDesdeCierre < 0 ? "text-slate-400" : "text-amber-700"}`}>
                                  {g.vencido
                                    ? `Vencido (${g.diasDesdeCierre} d)`
                                    : g.diasDesdeCierre < 0
                                    ? "En curso"
                                    : diasParaVencer === 0
                                    ? "Vence hoy"
                                    : `Vence en ${diasParaVencer} d`}
                                </td>
                                <td className="px-2 py-1">
                                  <button
                                    onClick={() => { setVistaLiquidacion("formal"); setObraFormalId(g.obraId); setMesFormal(g.mes); setQuincenaFormal(g.quincena); }}
                                    className={btnGhost}
                                  >
                                    Ir a Liquidación formal
                                  </button>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {pendientesEnBlanco.length > 0 && semanasOrdenadas.length > 0 && (
                  <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Personal en negro — pago en mano</div>
                )}

                {semanasOrdenadas.length === 0 ? (
                  pendientesEnBlanco.length === 0 && (
                    <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                      No hay días pendientes de pago en ninguna obra. 🎉
                    </div>
                  )
                ) : (
                  semanasOrdenadas.map((semanaKey) => {
                    const obrasDeSemana = gruposSemana[semanaKey];
                    const filasSemana = [];
                    Object.keys(obrasDeSemana).forEach((obraId) => {
                      const obra = obras.find((o) => o.id === Number(obraId));
                      const trabajadores = obrasDeSemana[obraId];
                      Object.keys(trabajadores).forEach((nombre) => {
                        filasSemana.push({ obraId, obraNombre: obra?.nombre || "Obra", nombre, ...trabajadores[nombre] });
                      });
                    });
                    const totalSemana = filasSemana.reduce((s, f) => s + f.monto, 0);
                    return (
                      <Panel
                        key={semanaKey}
                        title={`Semana del ${fmtFecha(semanaKey)}`}
                        action={<span className="font-mono text-sm font-bold text-slate-800">{fmtARS(totalSemana)}</span>}
                      >
                        <div className="overflow-x-auto rounded-md border border-stone-200">
                          <table className="w-full text-left text-xs">
                            <thead className="bg-stone-50 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
                              <tr>
                                <th className="px-2 py-1.5"></th>
                                <th className="px-2 py-1.5">Persona</th>
                                <th className="px-2 py-1.5">Obra</th>
                                <th className="px-2 py-1.5 text-right">Hs.</th>
                                <th className="px-2 py-1.5 text-right">Monto</th>
                              </tr>
                            </thead>
                            <tbody>
                              {filasSemana.map((f) => {
                                const key = `${semanaKey}|${f.obraId}|${f.nombre}`;
                                const seleccionado = seleccionLiquidacion.includes(key);
                                return (
                                  <tr
                                    key={key}
                                    onClick={() => toggleSeleccionLiquidacion(key)}
                                    className={`cursor-pointer border-t border-stone-100 ${seleccionado ? "bg-amber-50" : "hover:bg-stone-50"}`}
                                  >
                                    <td className="px-2 py-1">
                                      <input type="checkbox" checked={seleccionado} onChange={() => toggleSeleccionLiquidacion(key)} className="h-3.5 w-3.5" />
                                    </td>
                                    <td className="px-2 py-1 font-medium text-slate-900">{f.nombre}</td>
                                    <td className="px-2 py-1 text-slate-600">{f.obraNombre}</td>
                                    <td className="px-2 py-1 text-right font-mono text-slate-500">{f.horas}</td>
                                    <td className="px-2 py-1 text-right font-mono text-slate-800">{fmtARS(f.monto)}</td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </Panel>
                    );
                  })
                )}
              </>
            )}

            {vistaLiquidacion === "tanteros" && (
              <>
                <div className="flex items-center justify-end">
                  <button onClick={() => setShowTanteroForm((v) => !v)} className={btnPrimary}>
                    <Plus size={16} /> Nuevo grupo
                  </button>
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
                      <Field label="Formalidad del grupo">
                        <select value={tanteroForm.formalidad} onChange={(e) => setTanteroForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                          {FORMALIDADES.map((f) => <option key={f}>{f}</option>)}
                        </select>
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
                            <button
                              onClick={() => {
                                if (avanceAbiertoId === t.id) {
                                  setAvanceAbiertoId(null);
                                  setEditandoAvanceId(null);
                                  setAvanceForm(emptyAvanceForm);
                                } else {
                                  setAvanceAbiertoId(t.id);
                                  setEditandoAvanceId(null);
                                  setAvanceForm(emptyAvanceForm);
                                }
                              }}
                              className={btnGhost}
                            >
                              {avanceAbiertoId === t.id ? "Cancelar" : "Cargar avance"}
                            </button>
                          </div>

                          <div className="mt-3 grid grid-cols-2 gap-3 text-sm sm:grid-cols-4">
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
                            <div>
                              <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Formalidad del grupo</div>
                              <select
                                value={t.formalidad || ""}
                                onChange={(e) => updateRecord("tanteros", t.id, { formalidad: e.target.value }, setTanteros)}
                                className={`${inputCls} ${!t.formalidad ? "border-amber-400" : ""}`}
                              >
                                <option value="" disabled>Elegir…</option>
                                {FORMALIDADES.map((f) => <option key={f}>{f}</option>)}
                              </select>
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
                                <MoneyInput key={editandoAvanceId || "nuevo"} value={avanceForm.monto} onChange={(v) => setAvanceForm((f) => ({ ...f, monto: v }))} className={inputCls} />
                              </Field>
                              <Field label="Cuenta">
                                <select value={avanceForm.cuenta} onChange={(e) => setAvanceForm((f) => ({ ...f, cuenta: e.target.value }))} className={inputCls}>
                                  <option value="">Sin definir</option>
                                  {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                                </select>
                              </Field>
                              <div className="md:col-span-4">
                                <Field label="Descripción">
                                  <input value={avanceForm.descripcion} onChange={(e) => setAvanceForm((f) => ({ ...f, descripcion: e.target.value }))} placeholder="Ej: 3er avance" className={inputCls} />
                                </Field>
                              </div>
                              <div className="flex items-center gap-2 md:col-span-4">
                                <button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">
                                  {editandoAvanceId ? "Guardar cambios" : "Guardar avance"}
                                </button>
                                {editandoAvanceId && (
                                  <button
                                    type="button"
                                    onClick={() => { setEditandoAvanceId(null); setAvanceForm(emptyAvanceForm); }}
                                    className={btnGhost}
                                  >
                                    Cancelar edición
                                  </button>
                                )}
                              </div>
                            </form>
                          )}

                          {avancesGrupo.length > 0 && (
                            <div className="mt-4 space-y-1 border-t border-stone-100 pt-3">
                              {avancesGrupo.map((a) => (
                                <div key={a.id} className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
                                  <span>{fmtFecha(a.fecha)} — {a.descripcion || "Avance"}</span>
                                  <div className="flex items-center gap-2">
                                    <span className="font-mono text-slate-700">{fmtARS(a.monto)}</span>
                                    <button onClick={() => iniciarEdicionAvance(a)} title="Editar avance" className="text-slate-400 hover:text-slate-700"><Pencil size={13} /></button>
                                    <button onClick={() => deleteRecord("avances_tanteros", a.id, setAvancesTanteros)} title="Eliminar avance" className="text-slate-400 hover:text-rose-600"><Trash2 size={13} /></button>
                                  </div>
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

            {vistaLiquidacion === "formal" && (
              <>
                <div className="rounded-md border border-sky-300 bg-sky-50 px-4 py-3 text-xs text-sky-800">
                  Simulación de liquidación formal (UOCRA CCT 76/75, Zona A / San Juan — Ley 22.250), por quincena — así trabaja el
                  personal en blanco. Las horas parten de lo cargado en Asistencia pero son 100% editables, por si se declaran menos
                  horas "en blanco" que las reales. Esta quincena queda como pendiente en "Pendientes de pago" hasta que cargues acá el
                  costo real del recibo del contador.
                </div>

                <div className="flex flex-wrap items-end gap-3">
                  <div className="min-w-[200px]">
                    <Field label="Obra">
                      <select value={obraFormalId} onChange={(e) => setObraFormalId(e.target.value)} className={inputCls}>
                        {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                      </select>
                    </Field>
                  </div>
                  <div>
                    <Field label="Mes">
                      <MesPicker value={mesFormal} onChange={setMesFormal} />
                    </Field>
                  </div>
                  <div>
                    <Field label="Quincena">
                      <div className="flex gap-1">
                        <button
                          type="button" onClick={() => setQuincenaFormal(1)}
                          className={`rounded-md px-3 py-2 text-sm font-semibold ${quincenaFormal === 1 ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
                        >
                          1 al 15
                        </button>
                        <button
                          type="button" onClick={() => setQuincenaFormal(2)}
                          className={`rounded-md px-3 py-2 text-sm font-semibold ${quincenaFormal === 2 ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
                        >
                          16 al fin
                        </button>
                      </div>
                    </Field>
                  </div>
                  <button onClick={() => setShowFactoresLiquidacion((v) => !v)} className={btnGhost}>
                    {showFactoresLiquidacion ? "Ocultar factores" : "Ver / editar factores UOCRA"}
                  </button>
                </div>

                {showFactoresLiquidacion && (
                  <Panel title="Factores UOCRA / Régimen de la Construcción — San Juan (Zona A)" action={<button onClick={() => setShowFactoresLiquidacion(false)}><X size={16} /></button>}>
                    <div className="mb-3 text-xs text-slate-500">
                      Valores de referencia a agosto de 2026. Se actualizan acá cada vez que haya un nuevo acuerdo paritario UOCRA.
                      Los marcados <span className="font-semibold text-amber-700">(a confirmar)</span> no se pudieron verificar con precisión — pedile el valor vigente al contador.
                      {!canEditarCostos && " Solo Gerente y RRHH pueden modificarlos."}
                    </div>

                    <div className="mb-5">
                      <div className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-slate-500">Básico de convenio por hora, por categoría y mes</div>
                      <div className="mb-2 flex flex-wrap gap-2">
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
                        <table className="w-full text-left text-xs">
                          <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                            <tr>
                              <th className="sticky left-0 bg-stone-50 px-3 py-1.5">Categoría</th>
                              {mesesCostos.map((mes) => <th key={mes} className="px-2 py-1.5 text-right capitalize">{nombreMes(mes)}</th>)}
                            </tr>
                          </thead>
                          <tbody>
                            {CATEGORIAS_CONVENIO_UOCRA.map((categoria) => (
                              <tr key={categoria} className="border-t border-stone-100">
                                <td className="sticky left-0 bg-white px-3 py-1 font-medium text-slate-900">{categoria}</td>
                                {mesesCostos.map((mes) => {
                                  const entrada = basicosConvenio.find((c) => c.categoria === categoria && c.mes === mes);
                                  return (
                                    <td key={mes} className="px-2 py-1.5">
                                      {canEditarCostos ? (
                                        <MoneyInput
                                          className={`${inputCls} w-full text-right`}
                                          value={entrada?.basicoHora ?? 0}
                                          onBlur={(v) => guardarBasicoConvenioCelda(categoria, mes, v)}
                                        />
                                      ) : entrada?.basicoHora ? (
                                        <span className="block text-right">{fmtARS(entrada.basicoHora)}</span>
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
                    </div>

                    <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
                      <PctField label="Presentismo" value={cfgLiq.presentismoPct} onSave={(v) => actualizarConfigLiquidacion("presentismoPct", v)} />
                      <PctField label="Aporte jubilación (trabajador)" value={cfgLiq.aporteJubilacionPct} onSave={(v) => actualizarConfigLiquidacion("aporteJubilacionPct", v)} />
                      <PctField label="Aporte obra social (trabajador)" value={cfgLiq.aporteObraSocialPct} onSave={(v) => actualizarConfigLiquidacion("aporteObraSocialPct", v)} />
                      <PctField label="Aporte PAMI / Ley 19032" value={cfgLiq.aportePamiPct} onSave={(v) => actualizarConfigLiquidacion("aportePamiPct", v)} />
                      <PctField label="Cuota sindical UOCRA" value={cfgLiq.aporteSindicalPct} onSave={(v) => actualizarConfigLiquidacion("aporteSindicalPct", v)} />
                      <PctField label="Contrib. obra social (patronal)" value={cfgLiq.contribObraSocialPct} onSave={(v) => actualizarConfigLiquidacion("contribObraSocialPct", v)} />
                      <PctField label="Contrib. empresaria al gremio" value={cfgLiq.contribEmpresariaPct} onSave={(v) => actualizarConfigLiquidacion("contribEmpresariaPct", v)} />
                      <PctField label="Contrib. jubilatoria (patronal)" value={cfgLiq.contribJubilacionPct} onSave={(v) => actualizarConfigLiquidacion("contribJubilacionPct", v)} confirmar />
                      <PctField label="Fondo de Cese Laboral — 1er año (referencia)" value={cfgLiq.fondoCesePrimerAnioPct} onSave={(v) => actualizarConfigLiquidacion("fondoCesePrimerAnioPct", v)} />
                      <PctField label="Fondo de Cese Laboral — después (usado)" value={cfgLiq.fondoCesePosteriorPct} onSave={(v) => actualizarConfigLiquidacion("fondoCesePosteriorPct", v)} />
                      <div className="flex items-center gap-2 text-sm">
                        <span className="min-w-0 flex-1 text-[11px] font-semibold uppercase tracking-wide text-slate-500">IERIC (a confirmar)</span>
                        <MoneyInput className={`${inputCls} w-24 px-1.5 py-1 text-right`} value={cfgLiq.iericMontoFijo ?? 0} onBlur={(v) => actualizarConfigLiquidacion("iericMontoFijo", v)} />
                      </div>
                    </div>
                  </Panel>
                )}

                {filasFormal.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    No hay asistencia de personal "En blanco" cargada para esta obra en la {etiquetaQuincena(mesFormal, quincenaFormal)}.
                  </div>
                ) : (
                  <>
                    <div className="grid grid-cols-1 gap-4 sm:grid-cols-3">
                      <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                        <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Costo en recibo (aprox.)</div>
                        <div className="mt-1 font-mono text-lg font-bold text-slate-900">{fmtARS(totalesFormal.costoEmpresa)}</div>
                        <div className="text-[11px] text-slate-400">{totalesFormal.horasRecibo} hs a liquidar</div>
                      </div>
                      <div className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                        <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Costo en negro (aprox.)</div>
                        <div className="mt-1 font-mono text-lg font-bold text-slate-900">{fmtARS(totalesFormal.costoNegro)}</div>
                        <div className="text-[11px] text-slate-400">{totalesFormal.horasNegro} hs pagadas informal</div>
                      </div>
                      <div className="rounded-lg border-2 border-amber-400 bg-amber-50 p-4 shadow-sm">
                        <div className="text-[11px] font-semibold uppercase tracking-wide text-amber-800">Gasto aproximado total</div>
                        <div className="mt-1 font-mono text-lg font-bold text-amber-900">{fmtARS(gastoAproximadoTotal)}</div>
                        <div className="text-[11px] text-amber-700">Recibo + negro — estimado, no el recibo real</div>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-end justify-between gap-4">
                      <div className="flex flex-wrap gap-2">
                        <button onClick={generarPdfHorasReales} className={btnGhost}>
                          <span className="flex items-center gap-1"><FileDown size={13} /> PDF horas (recibo / negro) para el contador</span>
                        </button>
                        <button onClick={generarPdfLiquidacionFormal} className="flex items-center gap-1 rounded-md bg-slate-900 px-3 py-2 text-xs font-semibold text-white hover:bg-slate-700">
                          <FileDown size={13} /> PDF liquidación simulada
                        </button>
                      </div>
                      <ArchivoInput
                        label="Recibo del contador (PDF o foto)"
                        value={reciboFormalActual?.archivo}
                        nombreArchivo={reciboFormalActual?.nombreArchivo}
                        onChange={guardarReciboFormal}
                      />
                    </div>

                    <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-stone-50 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
                          <tr>
                            <th className="px-2 py-2">Persona</th>
                            <th className="px-2 py-2">Categoría</th>
                            <th className="px-2 py-2 text-right">Hs. presentes</th>
                            <th className="px-2 py-2 text-right">Hs. en recibo</th>
                            <th className="px-2 py-2 text-right">Hs. en negro</th>
                            <th className="px-2 py-2 text-center">Presentismo</th>
                            <th className="px-2 py-2 text-right">Bruto</th>
                            <th className="px-2 py-2 text-right">Neto</th>
                            <th className="px-2 py-2 text-right">Costo empresa</th>
                            <th className="px-2 py-2 text-right">Costo negro (aprox.)</th>
                            <th className="px-2 py-2 text-right">Costo real (recibo)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filasFormal.map((f) => {
                            const diffReal = f.costoRealBlanco != null ? f.costoRealBlanco - f.costoEmpresa : null;
                            return (
                              <tr key={f.nombre} className="border-t border-stone-100">
                                <td className="px-2 py-1.5 font-medium text-slate-900">{f.nombre}</td>
                                <td className="px-2 py-1.5 text-slate-600">{f.categoria}</td>
                                <td className="px-2 py-1.5 text-right font-mono text-slate-400">{f.horasReales}</td>
                                <td className="px-2 py-1.5 text-right">
                                  <input
                                    key={`${f.nombre}-${f.horasRecibo}`}
                                    type="number" defaultValue={f.horasRecibo}
                                    onBlur={(e) => actualizarHorasRecibo(f.nombre, Number(e.target.value))}
                                    className="w-16 rounded border border-stone-300 px-1 py-0.5 text-right"
                                  />
                                </td>
                                <td className="px-2 py-1.5 text-right font-mono text-slate-500">{f.horasNegro}</td>
                                <td className="px-2 py-1.5 text-center">
                                  <input type="checkbox" checked={f.presentismo} onChange={(e) => actualizarPresentismoFormal(f.nombre, e.target.checked)} />
                                </td>
                                <td className="px-2 py-1.5 text-right font-mono">{fmtARS(f.bruto)}</td>
                                <td className="px-2 py-1.5 text-right font-mono font-semibold">{fmtARS(f.neto)}</td>
                                <td className="px-2 py-1.5 text-right font-mono text-slate-500">{fmtARS(f.costoEmpresa)}</td>
                                <td className="px-2 py-1.5 text-right font-mono text-slate-500">{fmtARS(f.costoNegro)}</td>
                                <td className="px-2 py-1.5 text-right">
                                  <MoneyInput
                                    key={`${f.nombre}-${f.costoRealBlanco ?? 0}`}
                                    value={f.costoRealBlanco ?? 0}
                                    onBlur={(v) => guardarCostoRealPersona(f.nombre, v)}
                                    className="w-24 rounded border border-stone-300 px-1.5 py-1 text-right"
                                  />
                                  {diffReal != null && (
                                    <div className={`mt-0.5 text-[10px] ${diffReal > 0 ? "text-rose-600" : "text-emerald-700"}`}>
                                      {diffReal > 0 ? "+" : ""}{fmtARS(diffReal)} vs. aprox.
                                    </div>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="border-t-2 border-stone-200 font-semibold">
                            <td className="px-2 py-2" colSpan={6}>Total</td>
                            <td className="px-2 py-2 text-right font-mono">{fmtARS(totalesFormal.bruto)}</td>
                            <td className="px-2 py-2 text-right font-mono">{fmtARS(totalesFormal.neto)}</td>
                            <td className="px-2 py-2 text-right font-mono">{fmtARS(totalesFormal.costoEmpresa)}</td>
                            <td className="px-2 py-2 text-right font-mono">{fmtARS(totalesFormal.costoNegro)}</td>
                            <td className="px-2 py-2"></td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </>
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
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="text-xs text-slate-500">Maquinaria y herramientas de alto valor, controladas de forma individual por número de serie.</div>
                  <button onClick={startAddHerramienta} className={btnPrimary}>
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
                      <BotonEliminar
                        onClick={() => { moverAPapelera("herramientas", viewingHerramienta.id, setHerramientas, viewingHerramienta.nombre); setViewingHerramientaId(null); }}
                        title="Eliminar herramienta"
                      />
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
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="text-xs text-slate-500">Cajas de herramientas manuales chicas, armadas por rubro. Primero se asignan a una obra y después a un operario de esa obra.</div>
                  <button onClick={() => setShowComboForm((v) => !v)} className={btnPrimary}>
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
                        Se va a llamar <span className="font-semibold text-slate-600">"Caja {comboForm.tipo} {generarNumeroCaja(comboForm.tipo)}"</span>
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
              </>
            )}

            {vistaHerramientas === "remitos" && (
              <>
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="text-xs text-slate-500">Traslados de herramientas de Alto Valor entre Oficina, obras y talleres — con aprobación de salida y de recepción.</div>
                  <button onClick={() => setShowRemitoForm((v) => !v)} className={btnPrimary}>
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
                    className={btnPrimary}
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
                              <table className="w-full text-left text-xs">
                                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                                  <tr>
                                    <th className="px-2 py-1.5"></th><th className="px-2 py-1.5">Herramienta</th><th className="px-2 py-1.5">Ubicación</th>
                                    <th className="px-2 py-1.5">Estado</th><th className="px-2 py-1.5"></th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {herramientasOrdenadas.map((h) => (
                                    <tr key={h.id} className="border-t border-stone-100">
                                      <td className="px-2 py-1">
                                        {h.estado === "Disponible" && <span title="Libre, sin obra asignada"><CheckCircle2 size={14} className="text-emerald-600" /></span>}
                                      </td>
                                      <td className="px-2 py-1">
                                        <span className="flex items-center gap-1.5 font-medium text-slate-800">
                                          <CategoriaHerrIcon categoria={h.categoria} /> {h.nombre}
                                          <span className="font-mono text-xs text-slate-400">({h.numeroSerie || "s/n"})</span>
                                        </span>
                                      </td>
                                      <td className="px-2 py-1 text-slate-600">
                                        <span className="inline-flex items-center gap-1"><MapPin size={11} className="text-amber-600" />{h.ubicacion}</span>
                                      </td>
                                      <td className="px-2 py-1"><Badge estado={h.estado} /></td>
                                      <td className="px-2 py-1"><button onClick={() => agregarHerramientaAlPedido(h)} className={btnGhost}>+ Agregar al pedido</button></td>
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
                              <button onClick={abrirArmadoPedido} className={btnPrimary}>
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
                                    <span className="flex items-center gap-2 text-xs text-slate-400">
                                      {fmtFecha(p.fecha)} · {p.items.length} ítem(s){canVerPreciosPedido && <> · <span className="font-mono font-semibold text-slate-600">{fmtARS(p.total)}</span></>}
                                      <BotonEliminar onClick={() => moverAPapelera("pedidos_materiales", p.id, setPedidosMateriales, `Pedido #${p.id}`)} title="Eliminar pedido" />
                                    </span>
                                  </div>
                                  <div className="mt-2 text-xs text-slate-500">{p.items.map((it) => it.material).join(", ")}</div>
                                  {p.fechaNecesaria && (
                                    <div className={`mt-1 text-xs ${urgencia}`}>
                                      Necesario para el {fmtFecha(p.fechaNecesaria)}
                                      {(p.estado === "Solicitado" || p.estado === "Aprobado") && dias !== null && dias < 0 && " — VENCIDO"}
                                    </div>
                                  )}

                                  {p.estado === "Solicitado" && (
                                    p.obraId != null ? (
                                      <div className="mt-2 flex flex-wrap items-center gap-2 text-xs italic text-slate-400">
                                        <ShoppingCart size={13} /> Pasó automáticamente a Órdenes de Compra — ahí lo aprueba, rechaza o modifica Gerencia.
                                        {puedeAprobarPedidos && (
                                          <button onClick={() => setTab("ordenes")} className="not-italic font-semibold text-amber-700 hover:underline">Ir a Órdenes de Compra</button>
                                        )}
                                      </div>
                                    ) : puedeAprobarPedidos ? (
                                      <div className="mt-2 flex gap-2">
                                        <button onClick={() => aprobarPedidoMaterial(p)} className="flex items-center gap-1 rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700">
                                          <Check size={13} /> Aprobar pedido
                                        </button>
                                        <button
                                          onClick={() => rechazarPedidoMaterial(p, window.prompt("Motivo del rechazo (opcional):") || "")}
                                          className="flex items-center gap-1 rounded-md border border-rose-300 px-3 py-1.5 text-xs font-semibold text-rose-600 hover:bg-rose-50"
                                        >
                                          <X size={13} /> Rechazar
                                        </button>
                                      </div>
                                    ) : (
                                      <div className="mt-2 text-xs italic text-slate-400">Esperando aprobación de Gerencia.</div>
                                    )
                                  )}
                                  {p.estado === "Rechazado" && p.observaciones && (
                                    <div className="mt-2 rounded-md border border-rose-200 bg-rose-50 px-2 py-1 text-xs text-rose-700">Rechazado — Motivo: {p.observaciones}</div>
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

                {catalogoMateriales.length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    Todavía no hay materiales en el catálogo. Importá un presupuesto para empezar a llenarlo.
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                        <tr>
                          <th className="px-2 py-1.5">Categoría</th><th className="px-2 py-1.5">Sub-categoría</th><th className="px-2 py-1.5">Tipo</th><th className="px-2 py-1.5">Material</th>
                          <th className="px-2 py-1.5">Unidad</th><th className="px-2 py-1.5 text-right">Último precio</th><th className="px-2 py-1.5">Último proveedor</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...catalogoMateriales].sort((a, b) => a.nombre.localeCompare(b.nombre)).map((m) => (
                          <tr key={m.id} className="border-t border-stone-100">
                            <td className="px-2 py-1 text-slate-600">{m.categoria}</td>
                            <td className="px-2 py-1 text-slate-600">{m.subcategoria || "—"}</td>
                            <td className="px-2 py-1 text-slate-600">{m.tipo || "—"}</td>
                            <td className="px-2 py-1 font-medium text-slate-900">{m.nombre}</td>
                            <td className="px-2 py-1 text-slate-600">{m.unidad}</td>
                            <td className="px-2 py-1 text-right font-mono">{fmtARS(m.ultimoPrecio)}</td>
                            <td className="px-2 py-1 text-slate-500">{m.ultimoProveedor || <span className="text-slate-300">Sin datos aún</span>}</td>
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
                {stockMateriales.filter((s) => s.cantidad > 0).length === 0 ? (
                  <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">
                    No hay materiales en stock general por ahora.
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                        <tr>
                          <th className="px-2 py-1.5">Material</th><th className="px-2 py-1.5">Categoría</th><th className="px-2 py-1.5 text-right">Cantidad</th>
                          <th className="px-2 py-1.5 text-right">Precio unitario</th><th className="px-2 py-1.5 text-right">Valor en stock</th><th className="px-2 py-1.5"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {stockMateriales.filter((s) => s.cantidad > 0).map((s) => (
                          <Fragment key={s.id}>
                            <tr className="border-t border-stone-100">
                              <td className="px-2 py-1 font-medium text-slate-900">{s.material}</td>
                              <td className="px-2 py-1 text-slate-600">{s.categoria}{s.subcategoria ? ` — ${s.subcategoria}` : ""}</td>
                              <td className="px-2 py-1 text-right font-mono">{s.cantidad} {s.unidad}</td>
                              <td className="px-2 py-1 text-right font-mono">{fmtARS(s.precioUnitario)}</td>
                              <td className="px-2 py-1 text-right font-mono">{fmtARS(s.cantidad * s.precioUnitario)}</td>
                              <td className="px-2 py-1">
                                <button onClick={() => { setAsignandoStockId(s.id); setObraParaStock(""); setCantidadParaStock(s.cantidad); }} className={btnGhost}>Asignar a obra</button>
                              </td>
                            </tr>
                            {asignandoStockId === s.id && (
                              <tr className="border-t border-stone-100 bg-stone-50">
                                <td colSpan={6} className="px-2 py-2">
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
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Órdenes de Compra</h2>
              <button onClick={() => setShowOcForm((v) => !v)} className={btnPrimary}>
                <Plus size={16} /> Nueva orden
              </button>
            </div>

            {showOcForm && (
              <Panel title="Añadir orden de compra" action={<button onClick={() => setShowOcForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    if (!f.get("proveedor")) { alert("Elegí un proveedor o escribí uno nuevo."); return; }
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
                  <Field label="Proveedor">
                    <ProveedorPicker proveedores={proveedores} onCrearProveedor={crearProveedorRapido} />
                  </Field>
                  <Field label="Ítems / detalle"><input name="item" className={inputCls} /></Field>
                  <Field label="Monto estimado ($)">
                    <MoneyInput name="montoEstimado" className={inputCls} />
                    <div className="mt-1 text-[11px] text-slate-400">Precio final, con IVA incluido.</div>
                  </Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="space-y-3">
              {ordenesCompra.filter((oc) => !obraIdsPapelera.has(oc.obraId)).map((oc) => {
                const obra = obras.find((o) => o.id === oc.obraId);
                const pendienteDecision = oc.estado === "Pendiente" || oc.estado === "Requiere aprobación";
                const enEdicion = editandoOcId === oc.id;
                const enRechazo = rechazandoOcId === oc.id;
                return (
                  <div key={oc.id} className="rounded-lg border border-stone-200 bg-white p-4 shadow-sm">
                    {enEdicion ? (
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-4">
                        <Field label="Proveedor">
                          <input value={ocEditDraft.proveedor} onChange={(e) => setOcEditDraft((d) => ({ ...d, proveedor: e.target.value }))} className={inputCls} />
                        </Field>
                        <div className="sm:col-span-2">
                          <Field label="Ítems / detalle">
                            <input value={ocEditDraft.item} onChange={(e) => setOcEditDraft((d) => ({ ...d, item: e.target.value }))} className={inputCls} />
                          </Field>
                        </div>
                        <Field label="Monto estimado">
                          <MoneyInput value={ocEditDraft.montoEstimado} onChange={(v) => setOcEditDraft((d) => ({ ...d, montoEstimado: v }))} className={inputCls} />
                        </Field>
                        <div className="flex gap-2 sm:col-span-4">
                          <button onClick={() => guardarEdicionOc(oc)} className="rounded-md bg-slate-900 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-700">Guardar</button>
                          <button onClick={cancelarEdicionOc} className={btnGhost}>Cancelar</button>
                        </div>
                      </div>
                    ) : (
                      <div className="flex flex-wrap items-center justify-between gap-3">
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="font-semibold text-slate-900">{oc.proveedor}</span>
                            {oc.pedidoId && (
                              <span className="rounded-full border border-sky-300 bg-sky-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-sky-700">Desde Pedidos de Obra</span>
                            )}
                          </div>
                          <div className="text-sm text-slate-500">{oc.item}</div>
                          <div className="mt-1 text-xs text-slate-400">{obra?.nombre} · {fmtFecha(oc.fecha)}</div>
                          {oc.estado === "Rechazada" && oc.observaciones && (
                            <div className="mt-1.5 rounded-md border border-rose-200 bg-rose-50 px-2 py-1 text-xs text-rose-700">Motivo: {oc.observaciones}</div>
                          )}
                        </div>
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="font-mono text-sm font-semibold text-slate-800">{fmtARS(oc.montoEstimado)}</span>
                          <Badge estado={oc.estado} />
                          {pendienteDecision && (
                            <>
                              <button onClick={() => aprobarOC(oc)} className="flex items-center gap-1 rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700">
                                <Check size={13} /> Aprobar
                              </button>
                              <button onClick={() => iniciarRechazoOc(oc)} className="flex items-center gap-1 rounded-md border border-rose-300 px-3 py-1.5 text-xs font-semibold text-rose-600 hover:bg-rose-50">
                                <X size={13} /> Rechazar
                              </button>
                              <button onClick={() => iniciarEdicionOc(oc)} className={btnGhost}>
                                <span className="flex items-center gap-1"><Pencil size={13} /> Modificar</span>
                              </button>
                            </>
                          )}
                          {oc.estado === "Aprobada" && (
                            <button onClick={() => recibirOC(oc.id)} className={btnGhost}>Marcar recibida</button>
                          )}
                          <BotonEliminar onClick={() => moverAPapelera("ordenes_compra", oc.id, setOrdenesCompra, `${oc.proveedor} — ${oc.item}`)} title="Eliminar orden de compra" />
                        </div>
                      </div>
                    )}
                    {enRechazo && (
                      <div className="mt-3 rounded-md border border-rose-200 bg-rose-50 p-3">
                        <Field label="Observaciones (obligatorio) — para que Logística sepa por qué">
                          <textarea
                            value={observacionesRechazoOc}
                            onChange={(e) => setObservacionesRechazoOc(e.target.value)}
                            rows={2}
                            placeholder="Ej: proveedor muy caro, no era urgente, pedir otra cotización..."
                            className={inputCls}
                          />
                        </Field>
                        <div className="mt-2 flex gap-2">
                          <button onClick={() => confirmarRechazoOc(oc)} className="rounded-md bg-rose-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-rose-700">Confirmar rechazo</button>
                          <button onClick={cancelarRechazoOc} className={btnGhost}>Cancelar</button>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab === "facturas" && !canVerFinanzas && (
          <div className="flex flex-col items-center justify-center gap-3 rounded-lg border-2 border-dashed border-stone-300 bg-white p-10 text-center">
            <Receipt size={28} className="text-slate-400" />
            <div className="text-sm font-semibold text-slate-700">Sección restringida</div>
            <div className="max-w-sm text-xs text-slate-500">Solo Gerente y Contador pueden ver y cargar gastos y facturas.</div>
          </div>
        )}

        {tab === "facturas" && canVerFinanzas && (
          <div className="space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Gastos y Facturas</h2>
              <div className="flex flex-wrap items-center gap-2">
                <button onClick={generarPdfContadores} className="flex items-center gap-1 rounded-md border border-stone-300 bg-white px-3 py-2 text-sm font-semibold text-slate-600 hover:bg-stone-50">
                  <FileDown size={16} /> PDF para el contador ({nombreMesCuentas(mesReporteContador)})
                </button>
                <button onClick={() => setShowFacturaForm((v) => !v)} className={btnPrimary}>
                  <Plus size={16} /> Cargar gasto
                </button>
              </div>
            </div>

            {showFacturaForm && (
              <Panel title="Cargar gasto / factura" action={<button onClick={() => setShowFacturaForm(false)}><X size={16} /></button>}>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    if (!f.get("proveedor")) { alert("Elegí un proveedor o escribí uno nuevo."); return; }
                    const formaPago = f.get("formaPago");
                    const medioBancario = formaPago === "Banco" ? f.get("medioBancario") : null;
                    const fechaPagoEcheq = formaPago === "eCheq" ? f.get("fechaPagoEcheq") : null;
                    addRecord("compras_facturas", {
                      fecha: f.get("fecha"),
                      obraId: f.get("obraId") ? Number(f.get("obraId")) : null,
                      ordenCompraId: null,
                      proveedor: f.get("proveedor"),
                      categoria: f.get("categoria"),
                      descripcion: f.get("descripcion") || "",
                      monto: Number(f.get("monto")) || 0,
                      comprobante: "",
                      tipoFactura: f.get("tipoFactura"),
                      formalidad: f.get("formalidad"),
                      formaPago,
                      medioBancario,
                      fechaPagoEcheq,
                      // Efectivo, Mercado Pago, débito/transferencia y crédito se acreditan al toque.
                      // El eCheq queda pendiente hasta su fecha de pago y la cuenta corriente hasta que la saldemos a mano.
                      // El eCheq se cobra a través del banco, por eso cuenta a la balanza de Banco.
                      cuenta: formaPago === "Cuenta corriente" ? null : formaPago === "eCheq" ? "Banco" : formaPago,
                      estado: (formaPago === "eCheq" || formaPago === "Cuenta corriente") ? "Pendiente" : "Pagada",
                      archivo: facturaArchivo,
                      nombreArchivo: facturaNombreArchivo,
                      tipoArchivo: facturaTipoArchivo,
                    }, setComprasFacturas);
                    e.target.reset();
                    setFacturaFormaPago("Efectivo");
                    setFacturaMedioBancario("Débito/Transferencia");
                    setFacturaPlazoEcheq("30");
                    setFacturaArchivo(null);
                    setFacturaNombreArchivo(null);
                    setFacturaTipoArchivo(null);
                    setShowFacturaForm(false);
                  }}
                >
                  <Field label="Fecha"><input name="fecha" type="date" defaultValue={hoyISO()} required className={inputCls} /></Field>
                  <Field label="Obra">
                    <select name="obraId" className={inputCls}>
                      <option value="">General (sin obra específica)</option>
                      {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                    </select>
                  </Field>
                  <Field label="Proveedor - Nombre de fantasía">
                    <ProveedorPicker proveedores={proveedores} onCrearProveedor={crearProveedorRapido} />
                  </Field>
                  <Field label="Categoría">
                    <select name="categoria" className={inputCls}>{CATEGORIAS_GASTO.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <Field label="Descripción">
                    <input name="descripcion" placeholder="Ej: Nafta camioneta, cemento para losa..." className={inputCls} />
                  </Field>
                  <Field label="Precio final ($)">
                    <MoneyInput name="monto" className={inputCls} />
                    <div className="mt-1 text-[11px] text-slate-400">Con IVA incluido.</div>
                  </Field>
                  <Field label="Formalidad">
                    <select name="formalidad" className={inputCls}>{FORMALIDADES.map((f) => <option key={f}>{f}</option>)}</select>
                  </Field>
                  <Field label="Factura">
                    <select name="tipoFactura" defaultValue="Sin factura" className={inputCls}>{TIPOS_FACTURA.map((t) => <option key={t}>{t}</option>)}</select>
                  </Field>
                  <Field label="Forma de pago">
                    <select name="formaPago" value={facturaFormaPago} onChange={(e) => setFacturaFormaPago(e.target.value)} className={inputCls}>
                      {FORMAS_PAGO.map((fp) => <option key={fp}>{fp}</option>)}
                    </select>
                  </Field>
                  {facturaFormaPago === "Banco" && (
                    <Field label="Medio">
                      <select name="medioBancario" value={facturaMedioBancario} onChange={(e) => setFacturaMedioBancario(e.target.value)} className={inputCls}>
                        {MEDIOS_BANCARIOS.map((m) => <option key={m}>{m}</option>)}
                      </select>
                    </Field>
                  )}
                  {facturaFormaPago === "eCheq" && (
                    <>
                      <Field label="Plazo">
                        <select value={facturaPlazoEcheq} onChange={(e) => setFacturaPlazoEcheq(e.target.value)} className={inputCls}>
                          <option value="30">30 días</option>
                          <option value="60">60 días</option>
                          <option value="90">90 días</option>
                          <option value="personalizado">Otro (elegir fecha)</option>
                        </select>
                      </Field>
                      <Field label="Fecha de pago">
                        <input
                          key={facturaPlazoEcheq}
                          name="fechaPagoEcheq"
                          type="date"
                          defaultValue={fechaMasDias(facturaPlazoEcheq === "personalizado" ? 30 : Number(facturaPlazoEcheq))}
                          required
                          className={inputCls}
                        />
                        <div className="mt-1 text-[11px] text-slate-400">Queda "Pendiente" hasta esta fecha — ese día pasa solo a "Pagado".</div>
                      </Field>
                    </>
                  )}
                  {facturaFormaPago === "Cuenta corriente" && (
                    <div className="flex items-center text-[11px] text-slate-400">Queda "Pendiente" hasta que la paguemos — ahí la marcás como "Pagada" desde la tabla.</div>
                  )}
                  <div className="md:col-span-2">
                    <ArchivoInput
                      label="Factura / comprobante (opcional acá, se puede agregar después)"
                      value={facturaArchivo}
                      nombreArchivo={facturaNombreArchivo}
                      onChange={(archivo, nombreArchivo, tipoArchivo) => { setFacturaArchivo(archivo); setFacturaNombreArchivo(nombreArchivo); setFacturaTipoArchivo(tipoArchivo); }}
                    />
                  </div>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-xs">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-2 py-1.5">Fecha</th><th className="px-2 py-1.5">Obra</th><th className="px-2 py-1.5">Proveedor</th><th className="px-2 py-1.5">Categoría</th><th className="px-2 py-1.5">Descripción</th><th className="px-2 py-1.5">Formalidad</th><th className="px-2 py-1.5">Forma de pago</th><th className="px-2 py-1.5">Factura</th><th className="px-2 py-1.5">Monto</th><th className="px-2 py-1.5">Estado</th><th className="px-2 py-1.5"></th></tr>
                </thead>
                <tbody>
                  {comprasFacturas.filter((c) => !obraIdsPapelera.has(c.obraId)).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((c) => {
                    const obra = obras.find((o) => o.id === c.obraId);
                    return (
                      <tr key={c.id} className="border-t border-stone-100" style={{ backgroundColor: `${colorDeObra(obra)}0d` }}>
                        <td className="px-2 py-1 text-slate-600">{fmtFecha(c.fecha)}</td>
                        <td className="px-2 py-1 text-slate-600"><span className="flex items-center gap-1.5"><ObraDot obra={obra} />{obra?.nombre || "General"}</span></td>
                        <td className="px-2 py-1 font-medium text-slate-900">{c.proveedor}</td>
                        <td className="px-2 py-1 text-slate-600">{c.categoria}</td>
                        <td className="px-2 py-1 text-slate-500">{c.descripcion || "—"}</td>
                        <td className="px-2 py-1"><Badge estado={c.formalidad || "Blanco"} /></td>
                        <td className="px-2 py-1 text-slate-600">
                          <span className="flex items-center gap-1"><CuentaIcon cuenta={c.formaPago === "eCheq" ? "Banco" : c.formaPago} />{c.formaPago || c.cuenta || "—"}{c.medioBancario ? ` · ${c.medioBancario}` : ""}</span>
                          {(c.formaPago === "eCheq" || c.medioBancario === "eCheq") && c.estado === "Pendiente" && (
                            <div className="text-[10px] text-slate-400">Cobra el {fmtFecha(c.fechaPagoEcheq)}</div>
                          )}
                        </td>
                        <td className="px-2 py-1">
                          {(!c.tipoFactura || c.tipoFactura === "Sin factura") ? (
                            <span className="rounded-full border border-amber-300 bg-amber-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-amber-700">S/F</span>
                          ) : (
                            <span className="rounded-full border border-emerald-300 bg-emerald-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-emerald-700">{c.tipoFactura}</span>
                          )}
                        </td>
                        <td className="px-2 py-1 text-right font-mono font-semibold text-slate-800">{fmtARS(c.monto)}</td>
                        <td className="px-2 py-1"><Badge estado={c.estado} /></td>
                        <td className="px-2 py-1">
                          <div className="flex flex-wrap items-center gap-1.5">
                            {c.estado === "Pendiente" && (
                              <button onClick={() => marcarFacturaPagada(c)} className={btnGhost}>Marcar pagada</button>
                            )}
                            <button onClick={() => setEditandoMovimiento({ origen: "compras_facturas", origenId: c.id })} className={btnGhost}>
                              <span className="flex items-center gap-1"><Pencil size={12} /> Editar</span>
                            </button>
                            <BotonEliminar onClick={() => moverAPapelera("compras_facturas", c.id, setComprasFacturas, `${c.proveedor} — ${fmtARS(c.monto)}`)} title="Eliminar gasto" />
                          </div>
                        </td>
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
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Ingresos</h2>
              <button onClick={() => setShowIngresoForm((v) => !v)} className={btnPrimary}>
                <Plus size={16} /> Cargar ingreso
              </button>
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
                      obraId: f.get("obraId") ? Number(f.get("obraId")) : null,
                      concepto: f.get("concepto"),
                      monto: Number(f.get("monto")) || 0,
                      formalidad: f.get("formalidad"),
                      cuenta: f.get("cuenta"),
                      medioBancario: f.get("cuenta") === "Banco" ? f.get("medioBancario") : null,
                      estado: f.get("estado"),
                      fechaCobroEstimada: f.get("estado") === "Pendiente" ? f.get("fechaCobroEstimada") : null,
                      archivo: ingresoArchivo,
                      nombreArchivo: ingresoNombreArchivo,
                      tipoArchivo: ingresoTipoArchivo,
                    }, setIngresos);
                    e.target.reset();
                    setIngresoCuenta(CUENTAS[0]);
                    setIngresoMedioBancario("Transferencia");
                    setIngresoEstado("Cobrado");
                    setIngresoArchivo(null);
                    setIngresoNombreArchivo(null);
                    setIngresoTipoArchivo(null);
                    setShowIngresoForm(false);
                  }}
                >
                  <Field label="Fecha"><input name="fecha" type="date" defaultValue={hoyISO()} required className={inputCls} /></Field>
                  <Field label="Obra">
                    <select name="obraId" className={inputCls}>
                      <option value="">General (sin obra específica)</option>
                      {obras.filter((o) => o.estado !== "Papelera").map((o) => <option key={o.id} value={o.id}>{o.nombre}</option>)}
                    </select>
                  </Field>
                  <Field label="Concepto"><input name="concepto" required placeholder="Ej: certificado de avance 3" className={inputCls} /></Field>
                  <Field label="Monto (ARS)"><MoneyInput name="monto" className={inputCls} /></Field>
                  <Field label="Formalidad">
                    <select name="formalidad" className={inputCls}>{FORMALIDADES.map((f) => <option key={f}>{f}</option>)}</select>
                  </Field>
                  <Field label="Cuenta">
                    <select name="cuenta" value={ingresoCuenta} onChange={(e) => setIngresoCuenta(e.target.value)} className={inputCls}>{CUENTAS.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  {ingresoCuenta === "Banco" && (
                    <Field label="Medio">
                      <select name="medioBancario" value={ingresoMedioBancario} onChange={(e) => setIngresoMedioBancario(e.target.value)} className={inputCls}>
                        <option value="Transferencia">Transferencia</option>
                        <option value="eCheq">eCheq</option>
                      </select>
                    </Field>
                  )}
                  <Field label="Estado">
                    <select name="estado" value={ingresoEstado} onChange={(e) => setIngresoEstado(e.target.value)} className={inputCls}>
                      <option value="Cobrado">Cobrado</option>
                      <option value="Pendiente">Pendiente (todavía no lo cobramos)</option>
                    </select>
                  </Field>
                  {ingresoEstado === "Pendiente" && (
                    <Field label="Fecha estimada de cobro">
                      <input name="fechaCobroEstimada" type="date" defaultValue={hoyISO()} required className={inputCls} />
                    </Field>
                  )}
                  <div className="md:col-span-2">
                    <ArchivoInput
                      label="Factura / comprobante (opcional)"
                      value={ingresoArchivo}
                      nombreArchivo={ingresoNombreArchivo}
                      onChange={(archivo, nombreArchivo, tipoArchivo) => { setIngresoArchivo(archivo); setIngresoNombreArchivo(nombreArchivo); setIngresoTipoArchivo(tipoArchivo); }}
                    />
                  </div>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm">
              <table className="w-full text-left text-xs">
                <thead className="bg-stone-50 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                  <tr><th className="px-2 py-1.5">Fecha</th><th className="px-2 py-1.5">Obra</th><th className="px-2 py-1.5">Concepto</th><th className="px-2 py-1.5">Formalidad</th><th className="px-2 py-1.5">Cuenta</th><th className="px-2 py-1.5">Monto</th><th className="px-2 py-1.5">Estado</th><th className="px-2 py-1.5"></th></tr>
                </thead>
                <tbody>
                  {ingresos.filter((i) => !obraIdsPapelera.has(i.obraId)).sort((a, b) => fechaLocal(b.fecha) - fechaLocal(a.fecha)).map((i) => {
                    const obra = obras.find((o) => o.id === i.obraId);
                    return (
                      <tr key={i.id} className="border-t border-stone-100">
                        <td className="px-2 py-1 text-slate-600">{fmtFecha(i.fecha)}</td>
                        <td className="px-2 py-1 text-slate-600">{obra?.nombre || "General"}</td>
                        <td className="px-2 py-1 font-medium text-slate-900">
                          <span className="flex items-center gap-1.5">
                            {i.concepto}
                            {i.archivo && (
                              <a href={i.archivo} target="_blank" rel="noreferrer" title={i.nombreArchivo || "Ver comprobante"} className="text-slate-400 hover:text-slate-700">
                                <FileDown size={13} />
                              </a>
                            )}
                          </span>
                        </td>
                        <td className="px-2 py-1"><Badge estado={i.formalidad || "Blanco"} /></td>
                        <td className="px-2 py-1 text-slate-600">
                          <span className="flex items-center gap-1"><CuentaIcon cuenta={i.cuenta} />{i.cuenta || "—"}{i.medioBancario ? ` · ${i.medioBancario}` : ""}</span>
                        </td>
                        <td className="px-2 py-1 text-right font-mono font-semibold text-emerald-700">{fmtARS(i.monto)}</td>
                        <td className="px-2 py-1">
                          <Badge estado={i.estado === "Pendiente" ? "Pendiente" : "Cobrado"} />
                          {i.estado === "Pendiente" && i.fechaCobroEstimada && (
                            <div className="mt-0.5 text-[10px] text-slate-400">Cobra el {fmtFecha(i.fechaCobroEstimada)}</div>
                          )}
                        </td>
                        <td className="px-2 py-1">
                          <div className="flex flex-wrap items-center gap-1.5">
                            {i.estado === "Pendiente" && (
                              <button onClick={() => marcarIngresoCobrado(i)} className={btnGhost}>Marcar cobrado</button>
                            )}
                            <BotonEliminar onClick={() => moverAPapelera("ingresos", i.id, setIngresos, `${i.concepto} — ${fmtARS(i.monto)}`)} title="Eliminar ingreso" />
                          </div>
                        </td>
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

        {tab === "cuentas" && canVerFinanzas && !showProximos && (
          <div className="space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Cuentas</h2>
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={() => setShowProximos(true)}
                  className="flex items-center gap-1 rounded-md border border-stone-300 bg-white px-3 py-2 text-sm font-semibold text-slate-600 hover:bg-stone-50"
                >
                  <CalendarClock size={16} /> Próximos pagos/ingresos
                </button>
                <button
                  onClick={arreglarCaja}
                  className="flex items-center gap-1 rounded-md border border-stone-300 bg-white px-3 py-2 text-sm font-semibold text-slate-600 hover:bg-stone-50"
                >
                  <Wrench size={16} /> Arreglo de caja
                </button>
                <button
                  onClick={() => setShowPrestamoForm((v) => !v)}
                  className="flex items-center gap-1 rounded-md border border-stone-300 bg-white px-3 py-2 text-sm font-semibold text-slate-600 hover:bg-stone-50"
                >
                  <Landmark size={16} /> Agregar préstamo
                </button>
                <button
                  onClick={() => setShowMovimientoForm((v) => !v)}
                  className={btnPrimary}
                >
                  <Plus size={16} /> Agregar movimiento
                </button>
              </div>
            </div>

            {showPrestamoForm && (
              <Panel title="Agregar préstamo" action={<button onClick={() => setShowPrestamoForm(false)}><X size={16} /></button>}>
                <div className="mb-3 text-xs text-slate-500">Plata de un inversor o del banco. El capital entra a la cuenta que elijas como plata real, pero es una deuda — el interés corre solo, día a día, hasta que lo marques como devuelto. Queda siempre en "General", sin obra asociada.</div>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitPrestamoForm}>
                  <Field label="Fecha">
                    <input type="date" value={prestamoForm.fecha} onChange={(e) => setPrestamoForm((f) => ({ ...f, fecha: e.target.value }))} required className={inputCls} />
                  </Field>
                  <Field label="Acreedor (inversor / banco)">
                    <input value={prestamoForm.acreedor} onChange={(e) => setPrestamoForm((f) => ({ ...f, acreedor: e.target.value }))} required placeholder="Ej: Juan Pérez, Banco San Juan" className={inputCls} />
                  </Field>
                  <Field label="Capital ($)">
                    <MoneyInput value={prestamoForm.capital} onChange={(v) => setPrestamoForm((f) => ({ ...f, capital: v }))} className={inputCls} />
                  </Field>
                  <Field label="Tasa anual (%)">
                    <input
                      type="number" min="0" step="0.01" placeholder="Ej: 60"
                      value={prestamoForm.tasaAnualPct}
                      onChange={(e) => setPrestamoForm((f) => ({ ...f, tasaAnualPct: e.target.value }))}
                      required
                      className={inputCls}
                    />
                  </Field>
                  <Field label="Cuenta donde entra">
                    <select value={prestamoForm.cuenta} onChange={(e) => setPrestamoForm((f) => ({ ...f, cuenta: e.target.value }))} className={inputCls}>
                      {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label="Formalidad">
                    <select value={prestamoForm.formalidad} onChange={(e) => setPrestamoForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                      {FORMALIDADES.map((f) => <option key={f}>{f}</option>)}
                    </select>
                  </Field>
                  <Field label="Fecha estimada de devolución">
                    <input type="date" value={prestamoForm.fechaEstimadaDevolucion} onChange={(e) => setPrestamoForm((f) => ({ ...f, fechaEstimadaDevolucion: e.target.value }))} className={inputCls} />
                  </Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            {showMovimientoForm && (
              <Panel title="Agregar movimiento" action={<button onClick={() => setShowMovimientoForm(false)}><X size={16} /></button>}>
                <div className="mb-3 text-xs text-slate-500">Pase de dinero entre cuentas (ej: sacar efectivo y depositarlo en el banco). Siempre dentro de la misma formalidad — blanco y negro nunca se mezclan.</div>
                <form
                  className="grid grid-cols-1 gap-4 md:grid-cols-3"
                  onSubmit={(e) => {
                    e.preventDefault();
                    const f = new FormData(e.target);
                    const cuentaOrigen = f.get("cuentaOrigen");
                    const cuentaDestino = f.get("cuentaDestino");
                    if (cuentaOrigen === cuentaDestino) { alert("La cuenta de origen y la de destino no pueden ser la misma."); return; }
                    addRecord("movimientos_cuenta", {
                      fecha: f.get("fecha"),
                      detalle: f.get("detalle"),
                      formalidad: f.get("formalidad"),
                      cuentaOrigen,
                      cuentaDestino,
                      monto: Number(f.get("monto")) || 0,
                    }, setMovimientosManual);
                    e.target.reset();
                    setShowMovimientoForm(false);
                  }}
                >
                  <Field label="Fecha"><input name="fecha" type="date" defaultValue={hoyISO()} required className={inputCls} /></Field>
                  <Field label="Detalle"><input name="detalle" required placeholder="Ej: Pase de efectivo a banco" className={inputCls} /></Field>
                  <Field label="Formalidad">
                    <select name="formalidad" className={inputCls}>{FORMALIDADES.map((f) => <option key={f}>{f}</option>)}</select>
                  </Field>
                  <Field label="Cuenta donde sale">
                    <select name="cuentaOrigen" className={inputCls}>{CUENTAS.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <Field label="Cuenta que recibe">
                    <select name="cuentaDestino" defaultValue={CUENTAS[1]} className={inputCls}>{CUENTAS.map((c) => <option key={c}>{c}</option>)}</select>
                  </Field>
                  <Field label="Monto ($)"><MoneyInput name="monto" className={inputCls} /></Field>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            {/* Celular: tarjetas apiladas, sin scroll horizontal. */}
            <div className="space-y-2 sm:hidden">
              {CUENTAS.map((cuenta) => {
                const saldoBlanco = saldoCuenta(cuenta, "Blanco");
                const saldoNegro = saldoCuenta(cuenta, "Negro");
                const total = saldoBlanco + saldoNegro;
                const real = dineroRealDe(cuenta);
                const diferencia = real === null ? null : real - total;
                return (
                  <div key={cuenta} className="rounded-lg border border-stone-200 bg-white p-3 shadow-sm">
                    <div className="flex items-center justify-between gap-2">
                      <span className="flex items-center gap-1.5 font-semibold text-slate-800"><CuentaIcon cuenta={cuenta} size={15} />{cuenta}</span>
                      <span className={`font-mono text-base font-bold ${total < 0 ? "text-rose-600" : "text-slate-900"}`}>{fmtARS(total)}</span>
                    </div>
                    <div className="mt-1.5 flex flex-wrap gap-x-4 gap-y-0.5 text-xs text-slate-500">
                      <span>Blanco: <span className="font-mono text-slate-700">{fmtARS(saldoBlanco)}</span></span>
                      <span>Negro: <span className="font-mono text-slate-700">{fmtARS(saldoNegro)}</span></span>
                    </div>
                    <div className="mt-2 flex items-center justify-between gap-2 border-t border-stone-100 pt-2">
                      <div className="flex items-center gap-1.5">
                        <span className="text-xs text-slate-500">Real:</span>
                        <MoneyInput value={real ?? 0} onBlur={(v) => actualizarDineroReal(cuenta, v)} className="w-24 rounded-md border border-stone-300 px-1.5 py-1 text-right text-xs" />
                      </div>
                      <span className={`text-xs font-semibold ${diferencia === null || Math.abs(diferencia) < 1 ? "text-slate-400" : "text-rose-600"}`}>
                        {diferencia === null ? "Sin dato" : `Dif: ${fmtARS(diferencia)}`}
                      </span>
                    </div>
                  </div>
                );
              })}
              <div className="rounded-lg border-2 border-stone-300 bg-stone-50 p-3">
                <div className="flex items-center justify-between gap-2">
                  <span className="font-bold text-slate-900">Total</span>
                  <span className={`font-mono text-base font-bold ${(totalBlanco + totalNegro) < 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(totalBlanco + totalNegro)}</span>
                </div>
                <div className="mt-1.5 flex flex-wrap gap-x-4 gap-y-0.5 text-xs text-slate-500">
                  <span>Blanco: <span className="font-mono text-slate-700">{fmtARS(totalBlanco)}</span></span>
                  <span>Negro: <span className="font-mono text-slate-700">{fmtARS(totalNegro)}</span></span>
                </div>
              </div>
            </div>

            {/* Tablet/PC: tabla completa. */}
            <div className="hidden overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm sm:inline-block">
              <table className="text-left text-sm">
                <thead className="bg-stone-50 text-xs font-semibold uppercase tracking-wide text-slate-500">
                  <tr>
                    <th className="px-4 py-2">Cuenta</th><th className="px-4 py-2 text-right">Blanco</th><th className="px-4 py-2 text-right">Negro</th>
                    <th className="px-4 py-2 text-right">Total</th><th className="px-4 py-2 text-right">Dinero real</th><th className="px-4 py-2 text-right">Diferencia</th>
                  </tr>
                </thead>
                <tbody>
                  {CUENTAS.map((cuenta) => {
                    const saldoBlanco = saldoCuenta(cuenta, "Blanco");
                    const saldoNegro = saldoCuenta(cuenta, "Negro");
                    const total = saldoBlanco + saldoNegro;
                    const real = dineroRealDe(cuenta);
                    const diferencia = real === null ? null : real - total;
                    return (
                      <tr key={cuenta} className="border-t border-stone-100">
                        <td className="px-4 py-2"><span className="flex items-center gap-2 font-medium text-slate-800"><CuentaIcon cuenta={cuenta} size={16} />{cuenta}</span></td>
                        <td className={`px-4 py-2 text-right font-mono ${saldoBlanco < 0 ? "text-rose-600" : "text-slate-700"}`}>{fmtARS(saldoBlanco)}</td>
                        <td className={`px-4 py-2 text-right font-mono ${saldoNegro < 0 ? "text-rose-600" : "text-slate-700"}`}>{fmtARS(saldoNegro)}</td>
                        <td className={`px-4 py-2 text-right font-mono font-semibold ${total < 0 ? "text-rose-600" : "text-slate-900"}`}>{fmtARS(total)}</td>
                        <td className="px-4 py-2 text-right">
                          <MoneyInput value={real ?? 0} onBlur={(v) => actualizarDineroReal(cuenta, v)} className="w-32 rounded-md border border-stone-300 px-2 py-1 text-right text-sm" />
                        </td>
                        <td className={`px-4 py-2 text-right font-mono font-semibold ${diferencia === null || Math.abs(diferencia) < 1 ? "text-slate-400" : "text-rose-600"}`}>
                          {diferencia === null ? "—" : fmtARS(diferencia)}
                        </td>
                      </tr>
                    );
                  })}
                  <tr className="border-t-2 border-stone-300 bg-stone-50">
                    <td className="px-4 py-2.5 font-bold text-slate-900">Total</td>
                    <td className={`px-4 py-2.5 text-right font-mono font-bold ${totalBlanco < 0 ? "text-rose-600" : "text-slate-900"}`}>{fmtARS(totalBlanco)}</td>
                    <td className={`px-4 py-2.5 text-right font-mono font-bold ${totalNegro < 0 ? "text-rose-600" : "text-slate-900"}`}>{fmtARS(totalNegro)}</td>
                    <td className={`px-4 py-2.5 text-right font-mono font-bold ${(totalBlanco + totalNegro) < 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(totalBlanco + totalNegro)}</td>
                    <td className="px-4 py-2.5"></td>
                    <td className="px-4 py-2.5"></td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div>
              <h3 className="mb-1.5 text-sm font-semibold uppercase tracking-wide text-slate-500">Balance por obra</h3>
              <ResumenObrasCuentas items={resumenPorObra} />
            </div>

            <div>
              <h3 className="mb-1.5 text-sm font-semibold uppercase tracking-wide text-slate-500">Movimientos — {nombreMesCuentas(mesActualClave)}</h3>
              <TablaMovimientos items={movimientosMesActual} obras={obras} onEditar={(m) => setEditandoMovimiento({ origen: m.origen, origenId: m.origenId })} />
            </div>

            {gruposMovimientosAnteriores.map((g) => (
              <details key={g.clave} className="rounded-lg border border-stone-200 bg-white">
                <summary className="cursor-pointer select-none rounded-lg px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-stone-50">
                  {nombreMesCuentas(g.clave)} <span className="font-normal text-slate-400">({g.items.length})</span>
                </summary>
                <div className="border-t border-stone-100 p-3">
                  <TablaMovimientos items={g.items} obras={obras} onEditar={(m) => setEditandoMovimiento({ origen: m.origen, origenId: m.origenId })} />
                </div>
              </details>
            ))}

            <div>
              <h3 className="mb-1.5 text-sm font-semibold uppercase tracking-wide text-slate-500">Préstamos</h3>
              <TablaPrestamos
                items={prestamos}
                pagos={prestamosPagos}
                onEditar={(p) => setEditandoPrestamoId(p.id)}
                onRegistrarPago={(p) => setPagandoPrestamoId(p.id)}
                onEliminar={(p) => moverAPapelera("prestamos", p.id, setPrestamos, p.acreedor)}
              />
            </div>
          </div>
        )}

        {tab === "cuentas" && canVerFinanzas && showProximos && (
          <div className="space-y-4">
            <button
              onClick={() => { setShowProximos(false); setMesProximosSeleccionado(null); }}
              className="text-xs font-semibold text-slate-500 hover:text-slate-800"
            >
              ← Volver a Cuentas
            </button>
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Próximos pagos e ingresos</h2>

            {mesProximosSeleccionado === null ? (
              gruposMesesProximosConAcumulado.length === 0 ? (
                <div className="rounded-lg border border-dashed border-stone-300 bg-white px-3 py-6 text-center text-sm text-slate-400">
                  No hay pagos ni ingresos pendientes proyectados.
                </div>
              ) : (
                <>
                  {/* Celular: tarjetas compactas. */}
                  <div className="space-y-1.5 sm:hidden">
                    <div className="rounded-lg border-2 border-stone-300 bg-stone-50 p-2.5">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-700">Saldo actual</span>
                        <span className="font-mono text-sm font-bold text-slate-900">{fmtARS(saldoActualTotal)}</span>
                      </div>
                    </div>
                    {gruposMesesProximosConAcumulado.map((m) => {
                      const totalMes = m.ingresos - m.egresos;
                      const clickable = m.clave !== "sin-fecha";
                      return (
                        <button
                          key={m.clave}
                          onClick={() => clickable && setMesProximosSeleccionado(m.clave)}
                          disabled={!clickable}
                          className={`w-full rounded-lg border border-stone-200 bg-white p-2.5 text-left text-xs shadow-sm ${clickable ? "hover:border-amber-300" : ""}`}
                        >
                          <div className="font-semibold text-slate-900">{m.clave === "sin-fecha" ? "Sin fecha estimada" : nombreMesCuentas(m.clave)}</div>
                          <div className="mt-1 grid grid-cols-3 gap-x-2 gap-y-0.5">
                            <div><div className="text-[9px] uppercase tracking-wide text-slate-400">Ingreso</div><div className="font-mono text-emerald-700">{fmtARS(m.ingresos)}</div></div>
                            <div><div className="text-[9px] uppercase tracking-wide text-slate-400">Egreso</div><div className="font-mono text-rose-600">{fmtARS(m.egresos)}</div></div>
                            <div><div className="text-[9px] uppercase tracking-wide text-slate-400">Total mes</div><div className={`font-mono font-semibold ${totalMes < 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(totalMes)}</div></div>
                          </div>
                          <div className="mt-1.5 flex items-center justify-between rounded-md border-2 border-stone-300 bg-stone-50 px-2 py-1.5">
                            <span className="text-[10px] font-bold uppercase tracking-wide text-slate-600">Acumulado</span>
                            <span className={`font-mono text-base font-bold ${m.acumulado === null ? "text-slate-400" : m.acumulado < 0 ? "text-rose-600" : "text-emerald-700"}`}>
                              {m.acumulado === null ? "—" : fmtARS(m.acumulado)}
                            </span>
                          </div>
                        </button>
                      );
                    })}
                  </div>

                  {/* Tablet/PC: tabla tipo Excel. */}
                  <div className="hidden overflow-x-auto rounded-lg border border-stone-200 bg-white shadow-sm sm:block">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-stone-50 text-xs font-semibold uppercase tracking-wide text-slate-500">
                        <tr>
                          <th className="px-4 py-2">Mes</th>
                          <th className="px-4 py-2 text-right">Ingreso</th>
                          <th className="px-4 py-2 text-right">Egreso</th>
                          <th className="px-4 py-2 text-right">Total del mes</th>
                          <th className="border-l-2 border-stone-300 bg-stone-100 px-4 py-2 text-right font-bold text-slate-700">Acumulado</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr className="border-t-2 border-stone-300 bg-stone-50">
                          <td className="px-4 py-2 font-bold text-slate-900">Saldo actual</td>
                          <td className="px-4 py-2 text-right text-slate-400">—</td>
                          <td className="px-4 py-2 text-right text-slate-400">—</td>
                          <td className="px-4 py-2 text-right text-slate-400">—</td>
                          <td className="border-l-2 border-stone-300 bg-stone-100 px-4 py-2 text-right font-mono text-base font-bold text-slate-900">{fmtARS(saldoActualTotal)}</td>
                        </tr>
                        {gruposMesesProximosConAcumulado.map((m) => {
                          const totalMes = m.ingresos - m.egresos;
                          const clickable = m.clave !== "sin-fecha";
                          return (
                            <tr
                              key={m.clave}
                              onClick={() => clickable && setMesProximosSeleccionado(m.clave)}
                              className={`border-t border-stone-100 ${clickable ? "cursor-pointer hover:bg-stone-50" : ""}`}
                            >
                              <td className="px-4 py-2 font-medium text-slate-900">{m.clave === "sin-fecha" ? "Sin fecha estimada" : nombreMesCuentas(m.clave)}</td>
                              <td className="px-4 py-2 text-right font-mono text-emerald-700">{fmtARS(m.ingresos)}</td>
                              <td className="px-4 py-2 text-right font-mono text-rose-600">{fmtARS(m.egresos)}</td>
                              <td className={`px-4 py-2 text-right font-mono font-semibold ${totalMes < 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(totalMes)}</td>
                              <td className={`border-l-2 border-stone-300 bg-stone-100/60 px-4 py-2 text-right font-mono text-base font-bold ${m.acumulado === null ? "text-slate-400" : m.acumulado < 0 ? "text-rose-600" : "text-emerald-700"}`}>
                                {m.acumulado === null ? "—" : fmtARS(m.acumulado)}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </>
              )
            ) : (
              (() => {
                const claveMes = mesProximosSeleccionado;
                const prestamosDelMes = prestamosPorDevolver.filter((p) => perteneceAMesProximos(p.fechaEstimadaDevolucion, claveMes));
                const echeqsSalidaDelMes = echeqsSalida.filter((c) => perteneceAMesProximos(c.fechaPagoEcheq, claveMes));
                const echeqsEntradaDelMes = echeqsEntrada.filter((i) => perteneceAMesProximos(i.fechaCobroEstimada || i.fecha, claveMes));
                const cuentasCorrientesDelMes = cuentasCorrientesPorProveedor.filter((g) => perteneceAMesProximos(g.fechaVencimiento, claveMes));
                const ingresosDelMes = ingresosPendientes.filter((i) => perteneceAMesProximos(i.fechaCobroEstimada || i.fecha, claveMes));
                const obrasDisponibleDelMes = obrasDisponibleProyectado.filter((o) => o.meses.includes(claveMes));
                return (
                  <>
                    <button onClick={() => setMesProximosSeleccionado(null)} className="text-xs font-semibold text-slate-500 hover:text-slate-800">
                      ← Volver a los meses
                    </button>
                    <h3 className="text-lg font-bold text-slate-900">{claveMes === "sin-fecha" ? "Sin fecha estimada" : nombreMesCuentas(claveMes)}</h3>

                    <div className="space-y-3">
                      <div>
                        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Devoluciones de préstamos</div>
                        {prestamosDelMes.length === 0 ? (
                          <div className="text-xs text-slate-400">No hay préstamos por devolver este mes.</div>
                        ) : (
                          <div className="space-y-1.5">
                            {prestamosDelMes.map((p) => {
                              const dias = p.fechaEstimadaDevolucion ? diasHasta(p.fechaEstimadaDevolucion) : null;
                              const total = calcularEstadoPrestamo(p, prestamosPagos).totalADevolver;
                              return (
                                <div key={p.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-stone-200 px-2.5 py-1.5 text-sm">
                                  <span className="font-medium text-slate-800">{p.acreedor}</span>
                                  <span className={`text-xs ${dias === null ? "text-slate-400" : dias < 0 ? "font-semibold text-rose-600" : dias <= 3 ? "font-semibold text-amber-700" : "text-slate-500"}`}>
                                    {p.fechaEstimadaDevolucion ? `${fmtFecha(p.fechaEstimadaDevolucion)}${dias < 0 ? ` — vencido hace ${Math.abs(dias)} día(s)` : dias === 0 ? " — hoy" : ` — en ${dias} día(s)`}` : "Sin fecha estimada"}
                                  </span>
                                  <span className="font-mono font-semibold text-rose-600">{fmtARS(total)}</span>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>

                      <div>
                        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Cheques (eCheqs)</div>
                        <div className="space-y-2.5">
                          <div>
                            <div className="mb-1 text-[11px] text-slate-400">De salida (a pagar)</div>
                            {echeqsSalidaDelMes.length === 0 ? (
                              <div className="text-xs text-slate-400">No hay eCheqs por pagar este mes.</div>
                            ) : (
                              <div className="space-y-1.5">
                                {echeqsSalidaDelMes.map((c) => {
                                  const dias = diasHasta(c.fechaPagoEcheq);
                                  return (
                                    <div key={c.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-stone-200 px-2.5 py-1.5 text-sm">
                                      <span className="font-medium text-slate-800">{c.proveedor}</span>
                                      <span className={`text-xs ${dias < 0 ? "font-semibold text-rose-600" : dias <= 3 ? "font-semibold text-amber-700" : "text-slate-500"}`}>
                                        {fmtFecha(c.fechaPagoEcheq)}{dias < 0 ? ` — vencido hace ${Math.abs(dias)} día(s)` : dias === 0 ? " — hoy" : ` — en ${dias} día(s)`}
                                      </span>
                                      <span className="font-mono font-semibold text-rose-600">{fmtARS(c.monto)}</span>
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                          <div>
                            <div className="mb-1 text-[11px] text-slate-400">De entrada (a cobrar)</div>
                            {echeqsEntradaDelMes.length === 0 ? (
                              <div className="text-xs text-slate-400">No hay eCheqs por cobrar este mes.</div>
                            ) : (
                              <div className="space-y-1.5">
                                {echeqsEntradaDelMes.map((i) => {
                                  const fechaEstimada = i.fechaCobroEstimada || i.fecha;
                                  const dias = diasHasta(fechaEstimada);
                                  return (
                                    <div key={i.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-stone-200 px-2.5 py-1.5 text-sm">
                                      <span className="font-medium text-slate-800">{i.concepto}</span>
                                      <span className={`text-xs ${dias < 0 ? "font-semibold text-rose-600" : dias <= 3 ? "font-semibold text-amber-700" : "text-slate-500"}`}>
                                        {fmtFecha(fechaEstimada)}{dias < 0 ? ` — vencido hace ${Math.abs(dias)} día(s)` : dias === 0 ? " — hoy" : ` — en ${dias} día(s)`}
                                      </span>
                                      <span className="font-mono font-semibold text-emerald-700">{fmtARS(i.monto)}</span>
                                      <button onClick={() => marcarIngresoCobrado(i)} className={btnGhost}>Marcar cobrado</button>
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>

                      <div>
                        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Obras (cuenta corriente con proveedores)</div>
                        {cuentasCorrientesDelMes.length === 0 ? (
                          <div className="text-xs text-slate-400">No hay saldo de cuenta corriente venciendo este mes.</div>
                        ) : (
                          <div className="space-y-1.5">
                            {cuentasCorrientesDelMes.map((g) => {
                              const dias = g.fechaVencimiento ? diasHasta(g.fechaVencimiento) : null;
                              return (
                                <div key={g.proveedor} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-stone-200 px-2.5 py-1.5 text-sm">
                                  <span className="font-medium text-slate-800">{g.proveedor} <span className="font-normal text-slate-400">({g.cantidad} compra{g.cantidad > 1 ? "s" : ""})</span></span>
                                  <div className="flex items-center gap-2">
                                    <span className="text-xs text-slate-500">Día de pago (de cada mes):</span>
                                    <input
                                      type="number"
                                      min="1"
                                      max="31"
                                      value={g.diaPago || ""}
                                      onChange={(e) => actualizarDiaPago(g.proveedorId, e.target.value)}
                                      disabled={!g.proveedorId}
                                      className="w-16 rounded-md border border-stone-300 px-2 py-1 text-xs"
                                    />
                                    {g.fechaVencimiento && (
                                      <span className="text-xs text-slate-500">({fmtFecha(g.fechaVencimiento)})</span>
                                    )}
                                    {dias !== null && (
                                      <span className={`text-xs ${dias < 0 ? "font-semibold text-rose-600" : dias <= 3 ? "font-semibold text-amber-700" : "text-slate-500"}`}>
                                        {dias < 0 ? `vencido hace ${Math.abs(dias)} día(s)` : dias === 0 ? "hoy" : `en ${dias} día(s)`}
                                      </span>
                                    )}
                                  </div>
                                  <span className="font-mono font-semibold text-rose-600">{fmtARS(g.monto)}</span>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>

                      <div>
                        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Obras — presupuesto disponible por gastar (estimado)</div>
                        {obrasDisponibleDelMes.length === 0 ? (
                          <div className="text-xs text-slate-400">No hay obras con presupuesto disponible proyectado para este mes.</div>
                        ) : (
                          <div className="space-y-1.5">
                            {obrasDisponibleDelMes.map((o) => (
                              <div key={o.obra.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-stone-200 px-2.5 py-1.5 text-sm">
                                <span className="font-medium text-slate-800">{o.obra.nombre}</span>
                                <span className="text-xs text-slate-500">1/{o.meses.length} del disponible ({fmtARS(o.disponibleTotal)})</span>
                                <span className="font-mono font-semibold text-rose-600">{fmtARS(o.montoPorMes)}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>

                      <div>
                        <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Ingresos pendientes de cobro</div>
                        {ingresosDelMes.length === 0 ? (
                          <div className="text-xs text-slate-400">No hay ingresos pendientes de cobro este mes.</div>
                        ) : (
                          <div className="space-y-1.5">
                            {ingresosDelMes.map((i) => {
                              const fechaEstimada = i.fechaCobroEstimada || i.fecha;
                              const dias = diasHasta(fechaEstimada);
                              return (
                                <div key={i.id} className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-stone-200 px-2.5 py-1.5 text-sm">
                                  <span className="font-medium text-slate-800">{i.concepto}</span>
                                  <span className={`text-xs ${dias < 0 ? "font-semibold text-rose-600" : dias <= 3 ? "font-semibold text-amber-700" : "text-slate-500"}`}>
                                    {fmtFecha(fechaEstimada)}{dias < 0 ? ` — vencido hace ${Math.abs(dias)} día(s)` : dias === 0 ? " — hoy" : ` — en ${dias} día(s)`}
                                  </span>
                                  <span className="font-mono font-semibold text-emerald-700">{fmtARS(i.monto)}</span>
                                  <button onClick={() => marcarIngresoCobrado(i)} className={btnGhost}>Marcar cobrado</button>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    </div>
                  </>
                );
              })()
            )}
          </div>
        )}

        {tab === "cobros_socios" && canVerFinanzas && (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-2xl font-bold tracking-tight text-slate-900">Cobros Ricardo y Pablo</h2>
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={() => setShowCobroJuntosForm((v) => !v)}
                  className="flex items-center gap-1 rounded-md border border-stone-300 bg-white px-3 py-2 text-sm font-semibold text-slate-600 hover:bg-stone-50"
                >
                  <Users size={16} /> Registrar juntos
                </button>
                <button
                  onClick={() => setShowCobroSocioForm((v) => !v)}
                  className={btnPrimary}
                >
                  <Plus size={16} /> Registrar cobro
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              {SOCIOS.map((s) => (
                <div key={s} className="rounded-lg border border-stone-200 bg-white p-3">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">Total cobrado — {s}</div>
                  <div className="font-mono text-lg font-bold text-slate-900">{fmtARS(totalCobradoPorSocio(s))}</div>
                </div>
              ))}
            </div>

            {showCobroJuntosForm && (
              <Panel title="Registrar cobro conjunto (mitad y mitad)" action={<button onClick={() => setShowCobroJuntosForm(false)}><X size={16} /></button>}>
                <div className="mb-3 text-xs text-slate-500">Cargás el total y se guarda como dos cobros separados en el historial, uno para Ricardo y otro para Pablo, cada uno por la mitad — y cada uno con su propia factura a nombre de Concretar.</div>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitCobroJuntosForm}>
                  <Field label="Fecha">
                    <input type="date" value={cobroJuntosForm.fecha} onChange={(e) => setCobroJuntosForm((f) => ({ ...f, fecha: e.target.value }))} required className={inputCls} />
                  </Field>
                  <Field label="Monto total ($)">
                    <MoneyInput value={cobroJuntosForm.monto} onChange={(v) => setCobroJuntosForm((f) => ({ ...f, monto: v }))} className={inputCls} />
                  </Field>
                  <Field label="Mitad para cada uno">
                    <input value={fmtARS((Number(cobroJuntosForm.monto) || 0) / 2)} disabled className={`${inputCls} cursor-not-allowed bg-stone-100 text-slate-500`} />
                  </Field>
                  <Field label="Cuenta de la que sale">
                    <select value={cobroJuntosForm.cuenta} onChange={(e) => setCobroJuntosForm((f) => ({ ...f, cuenta: e.target.value }))} className={inputCls}>
                      {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  {cobroJuntosForm.cuenta === "Banco" && (
                    <Field label="Medio">
                      <select value={cobroJuntosForm.medioBancario} onChange={(e) => setCobroJuntosForm((f) => ({ ...f, medioBancario: e.target.value }))} className={inputCls}>
                        <option value="Transferencia">Transferencia</option>
                        <option value="eCheq">eCheq</option>
                      </select>
                    </Field>
                  )}
                  <Field label="Formalidad">
                    <select value={cobroJuntosForm.formalidad} onChange={(e) => setCobroJuntosForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                      {FORMALIDADES.map((f) => <option key={f}>{f}</option>)}
                    </select>
                  </Field>
                  <Field label="Observaciones">
                    <input value={cobroJuntosForm.observaciones} onChange={(e) => setCobroJuntosForm((f) => ({ ...f, observaciones: e.target.value }))} placeholder="Opcional" className={inputCls} />
                  </Field>
                  <div className="md:col-span-3 grid grid-cols-1 gap-4 rounded-md border border-dashed border-stone-300 p-3 sm:grid-cols-2">
                    {SOCIOS.map((socio) => (
                      <div key={socio} className="space-y-2">
                        <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Factura de {socio}</div>
                        <Field label="Tipo de factura">
                          <select
                            value={cobroJuntosForm.facturas[socio].tipoFactura}
                            onChange={(e) => setFacturaSocioJuntos(socio, { tipoFactura: e.target.value })}
                            className={inputCls}
                          >
                            {TIPOS_FACTURA.map((t) => <option key={t}>{t}</option>)}
                          </select>
                        </Field>
                        <ArchivoInput
                          label={`Factura / comprobante de ${socio} (PDF o foto)`}
                          value={cobroJuntosForm.facturas[socio].archivo}
                          nombreArchivo={cobroJuntosForm.facturas[socio].nombreArchivo}
                          onChange={(archivo, nombreArchivo, tipoArchivo) => setFacturaSocioJuntos(socio, { archivo, nombreArchivo, tipoArchivo })}
                        />
                      </div>
                    ))}
                  </div>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            {showCobroSocioForm && (
              <Panel title="Registrar cobro" action={<button onClick={() => setShowCobroSocioForm(false)}><X size={16} /></button>}>
                <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitCobroSocioForm}>
                  <Field label="Socio">
                    <select value={cobroSocioForm.socio} onChange={(e) => setCobroSocioForm((f) => ({ ...f, socio: e.target.value }))} className={inputCls}>
                      {SOCIOS.map((s) => <option key={s}>{s}</option>)}
                    </select>
                  </Field>
                  <Field label="Fecha">
                    <input type="date" value={cobroSocioForm.fecha} onChange={(e) => setCobroSocioForm((f) => ({ ...f, fecha: e.target.value }))} required className={inputCls} />
                  </Field>
                  <Field label="Monto ($)">
                    <MoneyInput value={cobroSocioForm.monto} onChange={(v) => setCobroSocioForm((f) => ({ ...f, monto: v }))} className={inputCls} />
                  </Field>
                  <Field label="Cuenta de la que sale">
                    <select value={cobroSocioForm.cuenta} onChange={(e) => setCobroSocioForm((f) => ({ ...f, cuenta: e.target.value }))} className={inputCls}>
                      {CUENTAS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  {cobroSocioForm.cuenta === "Banco" && (
                    <Field label="Medio">
                      <select value={cobroSocioForm.medioBancario} onChange={(e) => setCobroSocioForm((f) => ({ ...f, medioBancario: e.target.value }))} className={inputCls}>
                        <option value="Transferencia">Transferencia</option>
                        <option value="eCheq">eCheq</option>
                      </select>
                    </Field>
                  )}
                  <Field label="Formalidad">
                    <select value={cobroSocioForm.formalidad} onChange={(e) => setCobroSocioForm((f) => ({ ...f, formalidad: e.target.value }))} className={inputCls}>
                      {FORMALIDADES.map((f) => <option key={f}>{f}</option>)}
                    </select>
                  </Field>
                  <Field label="Factura">
                    <select value={cobroSocioForm.tipoFactura} onChange={(e) => setCobroSocioForm((f) => ({ ...f, tipoFactura: e.target.value }))} className={inputCls}>
                      {TIPOS_FACTURA.map((t) => <option key={t}>{t}</option>)}
                    </select>
                  </Field>
                  <Field label="Observaciones">
                    <input value={cobroSocioForm.observaciones} onChange={(e) => setCobroSocioForm((f) => ({ ...f, observaciones: e.target.value }))} placeholder="Opcional" className={inputCls} />
                  </Field>
                  <div className="md:col-span-2">
                    <ArchivoInput
                      label="Factura / comprobante (PDF o foto)"
                      value={cobroSocioForm.archivo}
                      nombreArchivo={cobroSocioForm.nombreArchivo}
                      onChange={(archivo, nombreArchivo, tipoArchivo) => setCobroSocioForm((f) => ({ ...f, archivo, nombreArchivo, tipoArchivo }))}
                    />
                  </div>
                  <div className="flex items-end"><button className="rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-700">Guardar</button></div>
                </form>
              </Panel>
            )}

            <div className="flex flex-wrap gap-2">
              {["Todos", ...SOCIOS].map((s) => (
                <button
                  key={s}
                  onClick={() => setFiltroSocio(s)}
                  className={`rounded-md px-3 py-2 text-sm font-semibold ${filtroSocio === s ? "bg-amber-500 text-slate-900" : "border border-stone-300 bg-white text-slate-600 hover:bg-stone-50"}`}
                >
                  {s}
                </button>
              ))}
            </div>

            {cobrosSociosFiltrados.length === 0 ? (
              <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">Todavía no hay cobros cargados.</div>
            ) : (
              <div className="space-y-2">
                {cobrosSociosFiltrados.map((c) => (
                  <div key={c.id} className="rounded-lg border border-stone-200 bg-white p-3 shadow-sm">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <div className="flex items-center gap-2">
                        <span className="font-semibold text-slate-900">{c.socio}</span>
                        <span className="text-xs text-slate-500">{fmtFecha(c.fecha)}</span>
                        <Badge estado={c.formalidad || "Blanco"} />
                        {(!c.tipoFactura || c.tipoFactura === "Sin factura") ? (
                          <span className="rounded-full border border-amber-300 bg-amber-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-amber-700">S/F</span>
                        ) : (
                          <span className="rounded-full border border-emerald-300 bg-emerald-50 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-emerald-700">{c.tipoFactura}</span>
                        )}
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="font-mono text-base font-bold text-rose-600">{fmtARS(c.monto)}</span>
                        <button onClick={() => setEditandoMovimiento({ origen: "cobros_socios", origenId: c.id })} className={btnGhost}>
                          <span className="flex items-center gap-1"><Pencil size={12} /> Editar</span>
                        </button>
                        <BotonEliminar onClick={() => moverAPapelera("cobros_socios", c.id, setCobrosSocios, `Cobro de ${c.socio}`)} title="Eliminar cobro" />
                      </div>
                    </div>
                    <div className="mt-1.5 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-500">
                      <span className="flex items-center gap-1"><CuentaIcon cuenta={c.cuenta} />{c.cuenta}{c.medioBancario ? ` · ${c.medioBancario}` : ""}</span>
                      {c.observaciones && <span>{c.observaciones}</span>}
                      {c.archivo && (
                        <a href={c.archivo} target="_blank" rel="noreferrer" className="flex items-center gap-1 text-slate-600 hover:underline">
                          <FileDown size={13} /> {c.nombreArchivo || "Ver comprobante"}
                        </a>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
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
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="text-xs text-slate-500">El saldo es lo acordado en el presupuesto de sus obras menos lo que ya ingresó por esa obra.</div>
                  <button
                    onClick={() => {
                      if (showClienteForm) { setShowClienteForm(false); return; }
                      setClienteForm(emptyClienteForm);
                      setEditandoClienteId(null);
                      setShowClienteForm(true);
                    }}
                    className={btnPrimary}
                  >
                    <Plus size={16} /> Nuevo cliente
                  </button>
                </div>

                {showClienteForm && (
                  <Panel title={editandoClienteId ? "Modificar cliente" : "Añadir cliente"} action={<button onClick={() => { setShowClienteForm(false); setEditandoClienteId(null); }}><X size={16} /></button>}>
                    <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitClienteForm}>
                      <Field label="Razón social / Nombre">
                        <input value={clienteForm.razonSocial} onChange={(e) => setClienteForm((f) => ({ ...f, razonSocial: e.target.value }))} required className={inputCls} />
                      </Field>
                      <Field label="Nombre de fantasía">
                        <input value={clienteForm.nombreFantasia} onChange={(e) => setClienteForm((f) => ({ ...f, nombreFantasia: e.target.value }))} className={inputCls} />
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
                      <Field label="Email">
                        <input type="email" value={clienteForm.email} onChange={(e) => setClienteForm((f) => ({ ...f, email: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="CBU">
                        <input value={clienteForm.cbu} onChange={(e) => setClienteForm((f) => ({ ...f, cbu: e.target.value }))} placeholder="22 dígitos" className={inputCls} />
                      </Field>
                      <Field label="Número de cuenta">
                        <input value={clienteForm.numeroCuenta} onChange={(e) => setClienteForm((f) => ({ ...f, numeroCuenta: e.target.value }))} className={inputCls} />
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
                              {cli.nombreFantasia && cli.nombreFantasia.trim() && cli.nombreFantasia.trim() !== cli.razonSocial && (
                                <div className="text-xs text-slate-400">{cli.nombreFantasia}</div>
                              )}
                              <div className="text-xs text-slate-500">{cli.contacto}{cli.contacto && cli.telefono ? " · " : ""}{cli.telefono}</div>
                              {cli.email && <div className="text-xs text-slate-500">{cli.email}</div>}
                            </div>
                            <div className="flex items-center gap-3">
                              <div className="text-right">
                                <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Saldo — te debe</div>
                                <div className={`font-mono text-lg font-bold ${saldo > 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(saldo)}</div>
                              </div>
                              <button onClick={() => editarCliente(cli)} className={btnGhost}>
                                <span className="flex items-center gap-1"><Pencil size={13} /> Modificar</span>
                              </button>
                              <BotonEliminar onClick={() => moverAPapelera("clientes", cli.id, setClientes, nombreComercial(cli))} title="Eliminar cliente" />
                            </div>
                          </div>
                          <div className="mt-2 flex flex-wrap gap-4 text-xs text-slate-500">
                            <span>Acordado: <span className="font-mono text-slate-700">{fmtARS(totalAcordado)}</span></span>
                            <span>Cobrado: <span className="font-mono text-slate-700">{fmtARS(totalCobrado)}</span></span>
                            <span>Obras: {obrasCliente.map((o) => o.nombre).join(", ") || "—"}</span>
                            {cli.cbu && <span>CBU: <span className="font-mono text-slate-700">{cli.cbu}</span></span>}
                            {cli.numeroCuenta && <span>Cuenta: <span className="font-mono text-slate-700">{cli.numeroCuenta}</span></span>}
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
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="text-xs text-slate-500">Los talleres de reparación también se cargan acá — así aparecen como destino posible en los remitos de Herramientas. El saldo es lo facturado menos lo ya pagado.</div>
                  <button
                    onClick={() => {
                      if (showProveedorForm) { setShowProveedorForm(false); return; }
                      setProveedorForm(emptyProveedorForm);
                      setEditandoProveedorId(null);
                      setShowProveedorForm(true);
                    }}
                    className={btnPrimary}
                  >
                    <Plus size={16} /> Nuevo proveedor
                  </button>
                </div>

                {showProveedorForm && (
                  <Panel title={editandoProveedorId ? "Modificar proveedor" : "Añadir proveedor"} action={<button onClick={() => { setShowProveedorForm(false); setEditandoProveedorId(null); }}><X size={16} /></button>}>
                    <form className="grid grid-cols-1 gap-4 md:grid-cols-3" onSubmit={submitProveedorForm}>
                      <Field label="Razón social">
                        <input value={proveedorForm.razonSocial} onChange={(e) => setProveedorForm((f) => ({ ...f, razonSocial: e.target.value }))} required className={inputCls} />
                      </Field>
                      <Field label="Nombre de fantasía">
                        <input value={proveedorForm.nombreFantasia} onChange={(e) => setProveedorForm((f) => ({ ...f, nombreFantasia: e.target.value }))} placeholder="Con este nombre se elige en Compras" className={inputCls} />
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
                      <Field label="Email">
                        <input type="email" value={proveedorForm.email} onChange={(e) => setProveedorForm((f) => ({ ...f, email: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="CBU">
                        <input value={proveedorForm.cbu} onChange={(e) => setProveedorForm((f) => ({ ...f, cbu: e.target.value }))} placeholder="22 dígitos" className={inputCls} />
                      </Field>
                      <Field label="Número de cuenta">
                        <input value={proveedorForm.numeroCuenta} onChange={(e) => setProveedorForm((f) => ({ ...f, numeroCuenta: e.target.value }))} className={inputCls} />
                      </Field>
                      <Field label="Día de pago (de cada mes)">
                        <input
                          type="number"
                          min="1"
                          max="31"
                          placeholder="Ej: 10"
                          value={proveedorForm.diaPago}
                          onChange={(e) => setProveedorForm((f) => ({ ...f, diaPago: e.target.value }))}
                          className={inputCls}
                        />
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
                              <span className="font-semibold text-slate-900">{nombreComercial(p)}</span>
                              {p.esTaller === "Sí" && <span className="ml-2"><Badge estado="En Reparación" /></span>}
                              {p.nombreFantasia && p.nombreFantasia.trim() && p.nombreFantasia.trim() !== p.razonSocial && (
                                <div className="text-xs text-slate-400">Razón social: {p.razonSocial}</div>
                              )}
                              <div className="text-xs text-slate-500">{p.contacto}{p.contacto && p.telefono ? " · " : ""}{p.telefono}</div>
                              {p.email && <div className="text-xs text-slate-500">{p.email}</div>}
                            </div>
                            <div className="flex items-center gap-3">
                              <div className="text-right">
                                <div className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">Saldo — le debés</div>
                                <div className={`font-mono text-lg font-bold ${saldo > 0 ? "text-rose-600" : "text-emerald-700"}`}>{fmtARS(saldo)}</div>
                              </div>
                              <button onClick={() => editarProveedor(p)} className={btnGhost}>
                                <span className="flex items-center gap-1"><Pencil size={13} /> Modificar</span>
                              </button>
                              <BotonEliminar onClick={() => moverAPapelera("proveedores", p.id, setProveedores, nombreComercial(p))} title="Eliminar proveedor" />
                            </div>
                          </div>
                          <div className="mt-2 flex flex-wrap gap-4 text-xs text-slate-500">
                            <span>Facturado: <span className="font-mono text-slate-700">{fmtARS(totalFacturado)}</span></span>
                            <span>Pagado: <span className="font-mono text-slate-700">{fmtARS(totalPagado)}</span></span>
                            <span>Día de pago: <span className="font-mono text-slate-700">{p.diaPago ? `${p.diaPago} de cada mes` : "sin definir"}</span></span>
                            {p.cbu && <span>CBU: <span className="font-mono text-slate-700">{p.cbu}</span></span>}
                            {p.numeroCuenta && <span>Cuenta: <span className="font-mono text-slate-700">{p.numeroCuenta}</span></span>}
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

        {tab === "papelera" && (
          <div className="space-y-6">
            <h2 className="text-2xl font-bold tracking-tight text-slate-900">Papelera</h2>
            {(herramientasPapelera.length + personalPapelera.length + proveedoresPapelera.length + clientesPapelera.length
              + ordenesCompraPapelera.length + pedidosMaterialesPapelera.length + etapasObraPapelera.length
              + (canVerFinanzas ? comprasFacturasPapelera.length + ingresosPapelera.length + prestamosPapelera.length + cobrosSociosPapelera.length : 0)
            ) === 0 ? (
              <div className="rounded-lg border-2 border-dashed border-stone-300 bg-white p-8 text-center text-sm text-slate-500">La Papelera está vacía.</div>
            ) : (
              <div className="space-y-5">
                <SeccionPapelera
                  titulo="Herramientas"
                  items={herramientasPapelera}
                  nombreDe={(h) => h.nombre}
                  detalleDe={(h) => h.categoria}
                  onRestaurar={(h) => restaurarDePapelera("herramientas", h.id, setHerramientas)}
                />
                <SeccionPapelera
                  titulo="Personal"
                  items={personalPapelera}
                  nombreDe={(p) => nombreCompletoDe(p)}
                  detalleDe={(p) => p.categoria}
                  onRestaurar={(p) => restaurarDePapelera("personal", p.id, setPersonal)}
                />
                <SeccionPapelera
                  titulo="Proveedores"
                  items={proveedoresPapelera}
                  nombreDe={(p) => nombreComercial(p)}
                  detalleDe={(p) => p.contacto}
                  onRestaurar={(p) => restaurarDePapelera("proveedores", p.id, setProveedores)}
                />
                <SeccionPapelera
                  titulo="Clientes"
                  items={clientesPapelera}
                  nombreDe={(c) => nombreComercial(c)}
                  detalleDe={(c) => c.contacto}
                  onRestaurar={(c) => restaurarDePapelera("clientes", c.id, setClientes)}
                />
                {canVerFinanzas && (
                  <>
                    <SeccionPapelera
                      titulo="Gastos y Facturas"
                      items={comprasFacturasPapelera}
                      nombreDe={(c) => `${c.proveedor} — ${fmtARS(c.monto)}`}
                      detalleDe={(c) => `${fmtFecha(c.fecha)} · ${c.categoria}`}
                      onRestaurar={(c) => restaurarDePapelera("compras_facturas", c.id, setComprasFacturas)}
                    />
                    <SeccionPapelera
                      titulo="Ingresos"
                      items={ingresosPapelera}
                      nombreDe={(i) => `${i.concepto} — ${fmtARS(i.monto)}`}
                      detalleDe={(i) => fmtFecha(i.fecha)}
                      onRestaurar={(i) => restaurarDePapelera("ingresos", i.id, setIngresos)}
                    />
                    <SeccionPapelera
                      titulo="Préstamos"
                      items={prestamosPapelera}
                      nombreDe={(p) => p.acreedor}
                      detalleDe={(p) => `Capital: ${fmtARS(p.capital)}`}
                      onRestaurar={(p) => restaurarDePapelera("prestamos", p.id, setPrestamos)}
                    />
                    <SeccionPapelera
                      titulo="Cobros Ricardo y Pablo"
                      items={cobrosSociosPapelera}
                      nombreDe={(c) => `Cobro de ${c.socio} — ${fmtARS(c.monto)}`}
                      detalleDe={(c) => fmtFecha(c.fecha)}
                      onRestaurar={(c) => restaurarDePapelera("cobros_socios", c.id, setCobrosSocios)}
                    />
                  </>
                )}
                <SeccionPapelera
                  titulo="Órdenes de Compra"
                  items={ordenesCompraPapelera}
                  nombreDe={(oc) => `${oc.proveedor} — ${oc.item}`}
                  detalleDe={(oc) => fmtFecha(oc.fecha)}
                  onRestaurar={(oc) => restaurarDePapelera("ordenes_compra", oc.id, setOrdenesCompra)}
                />
                <SeccionPapelera
                  titulo="Pedidos de Obra"
                  items={pedidosMaterialesPapelera}
                  nombreDe={(p) => `Pedido #${p.id}`}
                  detalleDe={(p) => `${fmtFecha(p.fecha)} · ${p.items?.length || 0} ítem(s)`}
                  onRestaurar={(p) => restaurarDePapelera("pedidos_materiales", p.id, setPedidosMateriales)}
                />
                <SeccionPapelera
                  titulo="Etapas de Planificación"
                  items={etapasObraPapelera}
                  nombreDe={(e) => e.nombre}
                  detalleDe={(e) => `${obras.find((o) => o.id === e.obraId)?.nombre || "—"} · ${fmtFecha(e.inicio)} → ${fmtFecha(e.fin)}`}
                  onRestaurar={(e) => restaurarDePapelera("etapas_obra", e.id, setEtapasObra)}
                />
              </div>
            )}
          </div>
        )}

      </main>

      {editandoMovimiento && (
        <ModalEditarMovimiento
          key={`${editandoMovimiento.origen}-${editandoMovimiento.origenId}`}
          editando={editandoMovimiento}
          comprasFacturas={comprasFacturas}
          cobrosSocios={cobrosSocios}
          ingresos={ingresos}
          movimientosManual={movimientosManual}
          obras={obras}
          onClose={() => setEditandoMovimiento(null)}
          onGuardarCompra={guardarEdicionCompra}
          onGuardarCobro={guardarEdicionCobro}
          onGuardarIngreso={guardarEdicionIngreso}
          onGuardarManual={guardarEdicionManual}
          onEliminar={eliminarMovimiento}
        />
      )}

      {editandoPrestamoId && (
        <ModalEditarPrestamo
          key={editandoPrestamoId}
          prestamo={prestamos.find((p) => p.id === editandoPrestamoId)}
          onClose={() => setEditandoPrestamoId(null)}
          onGuardar={guardarEdicionPrestamo}
        />
      )}

      {pagandoPrestamoId && (
        <ModalPagoPrestamo
          key={pagandoPrestamoId}
          prestamo={prestamos.find((p) => p.id === pagandoPrestamoId)}
          pagos={prestamosPagos}
          onClose={() => setPagandoPrestamoId(null)}
          onGuardar={guardarPagoPrestamo}
        />
      )}
    </div>
  );
}
