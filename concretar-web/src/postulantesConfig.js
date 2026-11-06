// Listas compartidas entre el formulario público de postulación (postulacion.jsx) y la
// sección Postulantes de la app.

// Rubros de la construcción que puede manejar un postulante (y, para "Pasar a Personal",
// a qué especialidad de Personal corresponde cada uno).
export const RUBROS_POSTULANTE = [
  "Albañilería",
  "Hormigón / Estructura",
  "Electricidad",
  "Plomería / Gas",
  "Pintura",
  "Herrería / Soldadura",
  "Carpintería",
  "Durlock / Yesería",
  "Techos / Impermeabilización",
  "Cerámicos / Revestimientos",
  "Maquinista",
  "Chofer",
  "Otro",
];
export const ESPECIALIDAD_DE_RUBRO = {
  "Albañilería": "Civil",
  "Hormigón / Estructura": "Civil",
  "Electricidad": "Eléctrico",
  "Plomería / Gas": "Plomería",
  "Pintura": "Pintor",
  "Herrería / Soldadura": "Hierrero",
  "Carpintería": "Carpintero",
  "Durlock / Yesería": "Civil",
  "Techos / Impermeabilización": "Civil",
  "Cerámicos / Revestimientos": "Civil",
};
// Categoría dentro de cada rubro (las mismas de convenio que usa Personal, más Capataz).
export const CATEGORIAS_POSTULANTE = ["Ayudante", "Medio Oficial", "Oficial", "Oficial Especializado", "Capataz"];
export const MODALIDADES_POSTULANTE = ["Tiempo completo", "Por obra", "Por tanto"];
export const ESTADOS_POSTULANTE = ["Nuevo", "Contactado", "Apto", "No apto", "Contratado"];
// Orden de las categorías para ordenar "de mayor a menor" oficio.
export const RANGO_CATEGORIA = { Ayudante: 1, "Medio Oficial": 2, Oficial: 3, "Oficial Especializado": 4, Capataz: 5 };
