// ── CATÁLOGOS (extraídos del Excel MVP) ─────────────────────────────────────

export const EMPRESAS = ["Clean Sea", "Parana Logistica", "Terra Mare"];

export const BASES_POR_EMPRESA = {
  "Clean Sea":        ["Bahia Blanca", "San Lorenzo", "San Fernando", "Quequen", "Punta Loyola", "Parana Ports", "Galpon"],
  "Parana Logistica": ["Golondrina de Mar", "Atlantic Dama", "Parana Ports"],
  "Terra Mare":       ["San Fernando"],
};

export const AREAS_POR_EMPRESA = {
  "Clean Sea":        ["Operaciones", "Tecnica", "Administracion", "Comercial", "RRHH", "HSE", "Viveres", "Otros"],
  "Parana Logistica": ["Operaciones", "Tecnica", "Administracion", "Comercial", "RRHH", "Viveres", "Otros"],
  "Terra Mare":       ["Operaciones", "Tecnica", "Administracion", "Comercial", "RRHH", "HSE", "Viveres", "Otros"],
};

export const SUBAREA_TECNICA = {
  "Clean Sea":        ["Machi", "Clean Sea I", "Clean Sea II", "Clean Sea III", "Clean Sea n", "Fi / Fi"],
  "Parana Logistica": ["Sala de Maquinas", "Cubierta", "Habitabilidad", "Seguridad e Higiene", "Propulsion y gobierno", "Sistemas electricos"],
  "Terra Mare":       ["Sala de Maquinas", "Cubierta", "Habitabilidad", "Seguridad e Higiene", "Propulsion y gobierno", "Sistemas electricos"],
};

export const DETALLE_TECNICO = {
  "Sala de Maquinas": ["Motor principal", "Motores auxiliares / generadores", "Compresores", "Bombas (sentina, lastre, combustible)", "Intercambiadores de calor", "Filtros y purificadores", "Tuberías y válvulas"],
  "Cubierta":         ["Pintura y anticorrosivo", "Chapa y soldadura", "Grúa / winches / cabrestantes", "Ancla y cadenas", "Manguerotes y portillas", "Arenado y preparación de superficie", "Anodos de sacrificio", "Otro"],
  "Habitabilidad":    ["Carpintería y camarotes", "Electrodomésticos", "Lavarropas / lavandería", "Cámara de frío y cocina", "Vajilla y utensilios", "Colchones y ropa de cama", "Sanitarios y plomería", "RIB", "Otro"],
  "Seguridad e Higiene": ["Inspecciones PNA / certificaciones", "Armamento de salvamento (balsas, chalecos)", "Trajes de inmersión", "Cartas náuticas y publicaciones", "EPP general", "Botiquín y medicamentos", "Equipo contraincendio", "Otro"],
  "Propulsion y gobierno": ["Piloto automático", "Timón y servo", "Hélice y eje", "Bocina y prensaestopa", "Thruster de proa", "Otro"],
  "Sistemas electricos": ["Tableros y protecciones", "Generadores de emergencia", "Baterías", "Iluminación", "Cableado y terminales", "Equipos de navegación (radar, GPS, VHF, AIS)", "Otro"],
};

export const TIPOS_REQUISICION = [
  "Material Mantenimiento",
  "Material Urgencia",
  "Servicios Mantenimiento",
  "Servicios Urgencia",
  "Viveres",
  "Otros",
];

export const URGENCIA_OPTIONS = ["Critica", "Alta", "Normal"];

export const PLAZO_PAGO_OPTIONS = ["Contado", "30 días", "60 días", "90 días", "A convenir"];

export const CATEGORIAS_RECHAZO = [
  "Descripción incompleta",
  "Sin stock / duplicado",
  "Fuera de presupuesto",
  "Proveedor no disponible",
  "Otro",
];

export const STATUS_LABELS = {
  pendiente_revision: "Pendiente revisión",
  en_revision:        "En revisión",
  aprobado_cotizar:   "Aprobado p/ cotizar",
  rechazado:          "Rechazado",
  en_compra:          "En compra",
  entregado_parcial:  "Entrega parcial",
  entregado:          "Entregado",
  cerrado:            "Cerrado",
};

export const STATUS_COLOR = {
  pendiente_revision: "amber",
  en_revision:        "blue",
  aprobado_cotizar:   "teal",
  rechazado:          "red",
  en_compra:          "purple",
  entregado_parcial:  "orange",
  entregado:          "green",
  cerrado:            "gray",
};

export const URGENCIA_COLOR = {
  Critica: "red",
  Alta:    "amber",
  Normal:  "green",
};
