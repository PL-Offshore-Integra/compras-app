import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  EMPRESAS, BASES_POR_EMPRESA, AREAS_POR_EMPRESA, SUBAREA_TECNICA,
  DETALLE_TECNICO, TIPOS_REQUISICION, URGENCIA_OPTIONS, PLAZO_PAGO_OPTIONS,
  CATEGORIAS_RECHAZO, STATUS_LABELS, STATUS_COLOR, URGENCIA_COLOR
} from "./lib/catalogos";

// ─── DEMO MODE (sin Supabase) ───────────────────────────────────────────────
// En producción esto se reemplaza por los imports de ./lib/supabase
const DEMO_MODE = false;

let _demoData = [
  {
    id: "1", nro_solicitud: 1, titulo: "Compra Bujías Motor Principal",
    empresa: "Parana Logistica", base_buque: "Golondrina de Mar",
    area: "Tecnica", subarea: "Sala de Maquinas", detalle_tecnico: "Motor principal",
    tipo_requisicion: "Material Urgencia", urgencia: "Critica",
    solicitado_por: "Juan Pérez", fecha_necesaria: "2025-04-20",
    costo_estimado: 85000, moneda_estimada: "ARS",
    busco_alternativas: false, observaciones: "Urgente, parada de buque",
    status: "pendiente_revision", veces_devuelto: 0,
    created_at: new Date(Date.now() - 2 * 86400000).toISOString(),
    requisicion_items: [
      { id: "i1", nro_linea: 1, descripcion: "Bujía NGK BPR6ES", cantidad: 8, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "Bulonería del Puerto", cantidad_aprobada: null },
      { id: "i2", nro_linea: 2, descripcion: "Filtro aceite W712", cantidad: 2, unidad: "Uni", stock_disponible: 1, proveedor_sugerido: "", cantidad_aprobada: null },
    ],
    requisicion_historial: [
      { id: "h1", fecha: new Date(Date.now() - 2 * 86400000).toISOString(), evento: "Requisición creada", usuario: "Juan Pérez", status_nuevo: "pendiente_revision" }
    ]
  },
  {
    id: "2", nro_solicitud: 2, titulo: "Víveres Semana 18",
    empresa: "Parana Logistica", base_buque: "Atlantic Dama",
    area: "Viveres", subarea: null, detalle_tecnico: null,
    tipo_requisicion: "Viveres", urgencia: "Normal",
    solicitado_por: "María González", fecha_necesaria: "2025-04-25",
    costo_estimado: 320000, moneda_estimada: "ARS",
    busco_alternativas: true, observaciones: "",
    status: "aprobado_cotizar", veces_devuelto: 0,
    revisado_por: "Comprador", fecha_revision: new Date(Date.now() - 1 * 86400000).toISOString(),
    fecha_aprobacion: new Date(Date.now() - 86400000).toISOString(),
    created_at: new Date(Date.now() - 3 * 86400000).toISOString(),
    requisicion_items: [
      { id: "i3", nro_linea: 1, descripcion: "Arroz largo fino 1kg", cantidad: 20, unidad: "Kg", stock_disponible: 5, proveedor_sugerido: "Distribuidora Sur", cantidad_aprobada: 20 },
      { id: "i4", nro_linea: 2, descripcion: "Aceite girasol 1.5L", cantidad: 12, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "", cantidad_aprobada: 12 },
    ],
    requisicion_historial: [
      { id: "h2", fecha: new Date(Date.now() - 3 * 86400000).toISOString(), evento: "Requisición creada", usuario: "María González", status_nuevo: "pendiente_revision" },
      { id: "h3", fecha: new Date(Date.now() - 1 * 86400000).toISOString(), evento: "Aprobado para cotizar", usuario: "Comprador", status_nuevo: "aprobado_cotizar" },
    ]
  },
  {
    id: "3", nro_solicitud: 3, titulo: "Pintura Antifouling Casco",
    empresa: "Terra Mare", base_buque: "San Fernando",
    area: "Tecnica", subarea: "Cubierta", detalle_tecnico: "Pintura y anticorrosivo",
    tipo_requisicion: "Material Mantenimiento", urgencia: "Alta",
    solicitado_por: "Carlos Ruiz", fecha_necesaria: "2025-05-01",
    costo_estimado: 450000, moneda_estimada: "ARS",
    busco_alternativas: false, observaciones: "Varada programada semana 20",
    status: "en_compra", veces_devuelto: 1,
    revisado_por: "Comprador", fecha_revision: new Date(Date.now() - 5 * 86400000).toISOString(),
    fecha_aprobacion: new Date(Date.now() - 4 * 86400000).toISOString(),
    nro_oc: "OC-0003", proveedor_elegido: "Pinturas Marinas SA",
    motivo_proveedor: "Mejor precio y stock disponible inmediato",
    costo_real: 487000, moneda_real: "ARS", plazo_pago: "30 días",
    fecha_entrega_prom: "2025-04-28",
    created_at: new Date(Date.now() - 7 * 86400000).toISOString(),
    dias_solicitud_revision: 2, dias_revision_aprobacion: 1,
    desvio_presupuestario: 8.2,
    requisicion_items: [
      { id: "i5", nro_linea: 1, descripcion: "Antifouling Hempel 20L", cantidad: 6, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "", cantidad_aprobada: 6, precio_unitario: 72000, subtotal: 432000 },
      { id: "i6", nro_linea: 2, descripcion: "Primer epoxi 4L", cantidad: 2, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "", cantidad_aprobada: 2, precio_unitario: 27500, subtotal: 55000 },
    ],
    requisicion_historial: [
      { id: "h4", fecha: new Date(Date.now() - 7 * 86400000).toISOString(), evento: "Requisición creada", usuario: "Carlos Ruiz", status_nuevo: "pendiente_revision" },
      { id: "h5", fecha: new Date(Date.now() - 6 * 86400000).toISOString(), evento: "Devuelto: Descripción incompleta", usuario: "Comprador", status_nuevo: "rechazado" },
      { id: "h6", fecha: new Date(Date.now() - 5 * 86400000).toISOString(), evento: "Re-ingresado por solicitante", usuario: "Carlos Ruiz", status_nuevo: "pendiente_revision" },
      { id: "h7", fecha: new Date(Date.now() - 4 * 86400000).toISOString(), evento: "OC emitida — OC-0003", usuario: "Comprador", status_nuevo: "en_compra" },
    ]
  }
];
let _demoOcCounter = 4;

const demoApi = {
  async getRequisiciones(filtros = {}) {
    await delay(300);
    let data = [..._demoData];
    if (filtros.status) data = data.filter(r => r.status === filtros.status);
    if (filtros.empresa) data = data.filter(r => r.empresa === filtros.empresa);
    return data.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  },
  async getRequisicion(id) {
    await delay(200);
    return _demoData.find(r => r.id === id) || null;
  },
  async crearRequisicion(req, items, usuario) {
    await delay(400);
    const nueva = {
      ...req, id: String(Date.now()), nro_solicitud: _demoData.length + 1,
      status: "pendiente_revision", veces_devuelto: 0,
      created_at: new Date().toISOString(),
      requisicion_items: items.map((it, i) => ({ ...it, id: `i${Date.now()}${i}`, nro_linea: i + 1 })),
      requisicion_historial: [{ id: `h${Date.now()}`, fecha: new Date().toISOString(), evento: "Requisición creada", usuario, status_nuevo: "pendiente_revision" }]
    };
    _demoData.unshift(nueva);
    return nueva;
  },
  async actualizarRequisicion(id, cambios, usuario, evento) {
    await delay(300);
    const idx = _demoData.findIndex(r => r.id === id);
    if (idx === -1) throw new Error("No encontrado");
    const anterior = _demoData[idx];
    const actualizado = { ...anterior, ...cambios, updated_at: new Date().toISOString() };
    if (evento) {
      actualizado.requisicion_historial = [...(anterior.requisicion_historial || []), {
        id: `h${Date.now()}`, fecha: new Date().toISOString(),
        evento, usuario, status_anterior: anterior.status, status_nuevo: cambios.status || anterior.status
      }];
    }
    _demoData[idx] = actualizado;
    return actualizado;
  },
  async actualizarItems(reqId, items) {
    await delay(200);
    const idx = _demoData.findIndex(r => r.id === reqId);
    if (idx >= 0) _demoData[idx].requisicion_items = items.map((it, i) => ({ ...it, id: it.id || `i${Date.now()}${i}`, nro_linea: i + 1 }));
  },
  async getProveedores() {
    await delay(100);
    return [
      { id: "p1", nombre: "Bulonería del Puerto", rubro: "Ferretería" },
      { id: "p2", nombre: "Pinturas Marinas SA", rubro: "Pintura" },
      { id: "p3", nombre: "Distribuidora Sur", rubro: "Víveres" },
      { id: "p4", nombre: "Combustibles Parana", rubro: "Combustible" },
      { id: "p5", nombre: "TechMar Equipos", rubro: "Electrónica" },
    ];
  },
  async crearProveedor(prov) {
    await delay(200);
    const nuevo = { ...prov, id: `p${Date.now()}`, activo: true };
    return nuevo;
  },
  nextOcNum() { return `OC-${String(_demoOcCounter++).padStart(4, "0")}`; }
};

const delay = ms => new Promise(r => setTimeout(r, ms));
import { getRequisiciones, getRequisicion, crearRequisicion, actualizarRequisicion, actualizarItems, getProveedores, crearProveedor } from "./lib/supabase";
const api = DEMO_MODE ? demoApi : { getRequisiciones, getRequisicion, crearRequisicion, actualizarRequisicion, actualizarItems, getProveedores, crearProveedor, nextOcNum: () => `OC-${Date.now()}` };

// ─── HELPERS ────────────────────────────────────────────────────────────────
const fmt = (n, cur = "ARS") =>
  n != null ? new Intl.NumberFormat("es-AR", { style: "currency", currency: cur, maximumFractionDigits: 0 }).format(n) : "—";

const fmtDate = d => d ? new Date(d).toLocaleDateString("es-AR") : "—";

const diffDays = (a, b) => {
  if (!a || !b) return null;
  return Math.round(Math.abs(new Date(b) - new Date(a)) / 86400000);
};

// ─── CSS ────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0d1117;--surface:#161b22;--surface2:#21262d;--surface3:#30363d;
  --border:#30363d;--border2:#484f58;
  --text:#e6edf3;--muted:#8b949e;--muted2:#6e7681;
  --accent:#58a6ff;--accent2:#3fb950;--warn:#d29922;--danger:#f85149;
  --purple:#bc8cff;--teal:#39d353;--orange:#f0883e;
  --mono:'DM Mono',monospace;--sans:'DM Sans',sans-serif;
  --r:6px;--r2:10px;
}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.5;min-height:100vh}
.app{display:flex;min-height:100vh}

/* SIDEBAR */
.sidebar{width:240px;min-width:240px;background:var(--surface);border-right:1px solid var(--border);display:flex;flex-direction:column;padding:0}
.sidebar-header{padding:20px 18px 16px;border-bottom:1px solid var(--border)}
.sidebar-logo{font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:2px;color:var(--accent);text-transform:uppercase}
.sidebar-sub{font-size:11px;color:var(--muted2);margin-top:2px}
.nav-section{padding:12px 8px 4px;font-family:var(--mono);font-size:10px;letter-spacing:1.5px;color:var(--muted2);text-transform:uppercase;padding-left:14px}
.nav-item{display:flex;align-items:center;gap:10px;padding:8px 14px;font-size:13px;font-weight:400;cursor:pointer;color:var(--muted);border-left:2px solid transparent;border-radius:0 var(--r) var(--r) 0;margin:1px 6px 1px 0;transition:all .15s}
.nav-item:hover{color:var(--text);background:var(--surface2)}
.nav-item.active{color:var(--accent);border-left-color:var(--accent);background:rgba(88,166,255,.08);font-weight:500}
.nav-badge{margin-left:auto;background:var(--warn);color:#000;font-family:var(--mono);font-size:10px;font-weight:500;padding:1px 6px;border-radius:10px;min-width:18px;text-align:center}
.nav-badge.red{background:var(--danger);color:#fff}

/* MAIN */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0}
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:14px 28px;display:flex;align-items:center;justify-content:space-between;gap:12px}
.topbar-left{display:flex;align-items:center;gap:12px}
.topbar-title{font-family:var(--mono);font-size:12px;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase}
.topbar-right{display:flex;align-items:center;gap:8px}
.content{flex:1;overflow-y:auto;padding:24px 28px}

/* DEMO BADGE */
.demo-banner{background:rgba(210,153,34,.12);border:1px solid rgba(210,153,34,.3);border-radius:var(--r);padding:10px 16px;margin-bottom:20px;font-size:12px;color:var(--warn);display:flex;align-items:center;gap:8px}

/* STATS */
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}
.stat{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:16px 18px;transition:border-color .2s}
.stat:hover{border-color:var(--border2)}
.stat-label{font-size:11px;color:var(--muted);letter-spacing:.5px;margin-bottom:6px;text-transform:uppercase;font-family:var(--mono)}
.stat-value{font-family:var(--mono);font-size:26px;font-weight:500}
.stat-sub{font-size:11px;color:var(--muted2);margin-top:4px}
.v-blue{color:var(--accent)}.v-green{color:var(--accent2)}.v-amber{color:var(--warn)}.v-red{color:var(--danger)}.v-purple{color:var(--purple)}.v-teal{color:var(--teal)}.v-gray{color:var(--muted)}

/* CARD */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:20px;margin-bottom:16px}
.card-title{font-family:var(--mono);font-size:10px;letter-spacing:2px;color:var(--muted);text-transform:uppercase;margin-bottom:16px;display:flex;align-items:center;justify-content:space-between}

/* FILTERS */
.filters{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px;align-items:center}
.filter-chip{font-family:var(--mono);font-size:11px;padding:5px 12px;border-radius:20px;border:1px solid var(--border);cursor:pointer;color:var(--muted);transition:all .15s;white-space:nowrap}
.filter-chip:hover{border-color:var(--border2);color:var(--text)}
.filter-chip.active{border-color:var(--accent);color:var(--accent);background:rgba(88,166,255,.08)}
.search-input{background:var(--surface2);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:13px;padding:6px 12px;outline:none;transition:border-color .15s;min-width:200px}
.search-input:focus{border-color:var(--accent)}
.search-input::placeholder{color:var(--muted2)}

/* TABLE */
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{font-family:var(--mono);font-size:10px;letter-spacing:1px;color:var(--muted);text-transform:uppercase;padding:8px 12px;text-align:left;border-bottom:1px solid var(--border);white-space:nowrap}
td{padding:11px 12px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(255,255,255,.02)}
.tr-clickable{cursor:pointer}

/* BADGES */
.badge{display:inline-flex;align-items:center;font-family:var(--mono);font-size:10px;font-weight:500;padding:3px 8px;border-radius:4px;letter-spacing:.5px;white-space:nowrap}
.b-amber{background:rgba(210,153,34,.15);color:var(--warn)}
.b-blue{background:rgba(88,166,255,.15);color:var(--accent)}
.b-teal{background:rgba(57,211,83,.15);color:var(--teal)}
.b-red{background:rgba(248,81,73,.15);color:var(--danger)}
.b-purple{background:rgba(188,140,255,.15);color:var(--purple)}
.b-orange{background:rgba(240,136,62,.15);color:var(--orange)}
.b-green{background:rgba(63,185,80,.15);color:var(--accent2)}
.b-gray{background:rgba(139,148,158,.15);color:var(--muted)}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:6px;font-family:var(--mono);font-size:11px;font-weight:500;letter-spacing:.5px;padding:7px 14px;border-radius:var(--r);border:1px solid transparent;cursor:pointer;transition:all .15s;white-space:nowrap;text-transform:uppercase}
.btn-primary{background:var(--accent);color:#0d1117;border-color:var(--accent)}
.btn-primary:hover{background:#79c0ff}
.btn-success{background:var(--accent2);color:#0d1117;border-color:var(--accent2)}
.btn-success:hover{background:#56d364}
.btn-danger{background:transparent;color:var(--danger);border-color:var(--danger)}
.btn-danger:hover{background:rgba(248,81,73,.1)}
.btn-ghost{background:transparent;color:var(--muted);border-color:var(--border)}
.btn-ghost:hover{color:var(--text);border-color:var(--border2)}
.btn-warn{background:transparent;color:var(--warn);border-color:var(--warn)}
.btn-warn:hover{background:rgba(210,153,34,.1)}
.btn-sm{padding:4px 10px;font-size:10px}
.btn-xs{padding:3px 8px;font-size:10px}
.btn:disabled{opacity:.4;cursor:not-allowed}

/* MODAL */
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.75);display:flex;align-items:flex-start;justify-content:center;z-index:100;padding:20px;overflow-y:auto;animation:fadeIn .15s}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);width:100%;max-width:820px;margin:auto;animation:slideUp .2s}
.modal-header{display:flex;justify-content:space-between;align-items:flex-start;padding:20px 24px;border-bottom:1px solid var(--border)}
.modal-title{font-family:var(--mono);font-size:13px;letter-spacing:1px}
.modal-body{padding:24px}
.modal-footer{padding:16px 24px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px}
.modal-close{background:none;border:none;color:var(--muted);font-size:18px;cursor:pointer;padding:2px;line-height:1}
.modal-close:hover{color:var(--text)}

@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes slideUp{from{transform:translateY(16px);opacity:0}to{transform:translateY(0);opacity:1}}

/* FORM */
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.form-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}
.form-full{grid-column:1/-1}
.fg{display:flex;flex-direction:column;gap:5px}
.fg label{font-size:11px;color:var(--muted);letter-spacing:.5px;text-transform:uppercase;font-family:var(--mono)}
.fg input,.fg select,.fg textarea{background:var(--surface2);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:13px;padding:8px 10px;outline:none;transition:border-color .15s}
.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--accent)}
.fg select option{background:var(--surface2)}
.fg textarea{resize:vertical;min-height:70px}
.fg input[type=checkbox]{width:16px;height:16px;accent-color:var(--accent)}
.checkbox-row{display:flex;align-items:center;gap:8px;padding:8px 0}
.checkbox-row label{font-size:13px;color:var(--text);letter-spacing:0;text-transform:none;font-family:var(--sans);cursor:pointer}
.form-section{font-family:var(--mono);font-size:10px;letter-spacing:2px;color:var(--accent);text-transform:uppercase;margin:20px 0 12px;padding-bottom:6px;border-bottom:1px solid var(--border)}
.field-hint{font-size:11px;color:var(--muted2);margin-top:2px}

/* ITEMS TABLE */
.items-edit th{font-size:10px}
.items-edit td{padding:6px 8px}
.items-edit input{background:var(--surface2);border:1px solid var(--border);border-radius:4px;color:var(--text);font-family:var(--mono);font-size:12px;padding:4px 7px;width:100%;outline:none}
.items-edit input:focus{border-color:var(--accent)}

/* TIMELINE */
.timeline{padding:0;list-style:none}
.tl-item{display:flex;gap:12px;padding-bottom:16px;position:relative}
.tl-item:last-child{padding-bottom:0}
.tl-item:not(:last-child)::before{content:'';position:absolute;left:11px;top:24px;bottom:0;width:1px;background:var(--border)}
.tl-dot{width:24px;height:24px;border-radius:50%;background:var(--surface2);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:11px;flex-shrink:0;margin-top:2px;z-index:1}
.tl-dot.created{border-color:var(--accent);color:var(--accent)}
.tl-dot.approved{border-color:var(--accent2);color:var(--accent2)}
.tl-dot.rejected{border-color:var(--danger);color:var(--danger)}
.tl-dot.updated{border-color:var(--warn);color:var(--warn)}
.tl-content{flex:1;min-width:0}
.tl-evento{font-size:13px;font-weight:500}
.tl-meta{font-size:11px;color:var(--muted);margin-top:2px;display:flex;gap:12px}

/* KPI BARS */
.kpi-bar-wrap{margin-bottom:10px}
.kpi-bar-label{display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px}
.kpi-bar-track{height:6px;background:var(--surface2);border-radius:3px;overflow:hidden}
.kpi-bar-fill{height:100%;border-radius:3px;transition:width .5s}

/* UPLOAD */
.upload-zone{border:2px dashed var(--border);border-radius:var(--r2);padding:40px;text-align:center;cursor:pointer;transition:all .2s}
.upload-zone:hover,.upload-zone.drag{border-color:var(--accent);background:rgba(88,166,255,.04)}
.upload-icon{font-size:36px;margin-bottom:10px}
.upload-title{font-family:var(--mono);font-size:13px;margin-bottom:6px;color:var(--text)}
.upload-sub{font-size:12px;color:var(--muted)}

/* URGENCY INDICATOR */
.urg-dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-right:6px;flex-shrink:0}
.urg-Critica{background:var(--danger)}
.urg-Alta{background:var(--warn)}
.urg-Normal{background:var(--accent2)}

/* NOTIF */
.notif{position:fixed;bottom:24px;right:24px;background:var(--surface2);border:1px solid var(--border);border-left-width:3px;border-radius:var(--r2);padding:12px 16px;font-size:13px;animation:slideUp .2s;z-index:300;max-width:340px;display:flex;align-items:center;gap:10px}
.notif-close{margin-left:auto;background:none;border:none;color:var(--muted);cursor:pointer;font-size:16px}
.n-green{border-left-color:var(--accent2)}.n-red{border-left-color:var(--danger)}.n-amber{border-left-color:var(--warn)}.n-blue{border-left-color:var(--accent)}

/* MISC */
.divider{height:1px;background:var(--border);margin:16px 0}
.text-mono{font-family:var(--mono)}
.text-muted{color:var(--muted)}
.text-right{text-align:right}
.flex{display:flex}.flex-gap{display:flex;gap:8px;align-items:center}.flex-between{display:flex;justify-content:space-between;align-items:center}
.mt4{margin-top:4px}.mt8{margin-top:8px}.mt12{margin-top:12px}.mt16{margin-top:16px}.mt20{margin-top:20px}
.mb8{margin-bottom:8px}.mb12{margin-bottom:12px}.mb16{margin-bottom:16px}
.w100{width:100%}
.info-box{background:var(--surface2);border:1px solid var(--border);border-radius:var(--r);padding:12px 14px;font-size:12px;color:var(--muted)}
.info-box.accent{border-left:3px solid var(--accent);border-left-width:3px}
.tag{display:inline-block;font-family:var(--mono);font-size:10px;padding:2px 7px;background:var(--surface2);border:1px solid var(--border);border-radius:4px;color:var(--muted2)}
.spin{animation:spin 1s linear infinite}
@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
.empty-state{text-align:center;padding:48px 24px;color:var(--muted)}
.empty-icon{font-size:32px;margin-bottom:10px}
.loading{display:flex;align-items:center;justify-content:center;padding:60px;color:var(--muted);gap:10px;font-size:13px}
`;

// ─── SUB-COMPONENTS ─────────────────────────────────────────────────────────

function Notif({ msg, onClose }) {
  if (!msg) return null;
  const cls = { success: "n-green", error: "n-red", warn: "n-amber", info: "n-blue" }[msg.type] || "n-blue";
  return (
    <div className={`notif ${cls}`}>
      <span>{msg.text}</span>
      <button className="notif-close" onClick={onClose}>✕</button>
    </div>
  );
}

function Badge({ status }) {
  const color = STATUS_COLOR[status] || "gray";
  return <span className={`badge b-${color}`}>{STATUS_LABELS[status] || status}</span>;
}

function UrgBadge({ urgencia }) {
  const color = { Critica: "b-red", Alta: "b-amber", Normal: "b-green" }[urgencia] || "b-gray";
  return <span className={`badge ${color}`}><span className={`urg-dot urg-${urgencia}`} />{urgencia}</span>;
}

function KpiBar({ label, value, max, color = "var(--accent)", suffix = "" }) {
  const pct = max ? Math.min(100, (value / max) * 100) : 0;
  return (
    <div className="kpi-bar-wrap">
      <div className="kpi-bar-label">
        <span className="text-muted">{label}</span>
        <span className="text-mono">{value != null ? `${value}${suffix}` : "—"}</span>
      </div>
      <div className="kpi-bar-track">
        <div className="kpi-bar-fill" style={{ width: `${pct}%`, background: color }} />
      </div>
    </div>
  );
}

function Timeline({ historial }) {
  if (!historial?.length) return <div className="text-muted" style={{ fontSize: 12, padding: "8px 0" }}>Sin historial</div>;
  const iconMap = ev => {
    if (ev.includes("creada") || ev.includes("ingresado")) return { icon: "◎", cls: "created" };
    if (ev.includes("probado") || ev.includes("OC")) return { icon: "✓", cls: "approved" };
    if (ev.includes("echazado") || ev.includes("evuelto")) return { icon: "✗", cls: "rejected" };
    return { icon: "·", cls: "updated" };
  };
  return (
    <ul className="timeline">
      {[...historial].sort((a, b) => new Date(a.fecha) - new Date(b.fecha)).map((h, i) => {
        const { icon, cls } = iconMap(h.evento);
        return (
          <li key={i} className="tl-item">
            <div className={`tl-dot ${cls}`}>{icon}</div>
            <div className="tl-content">
              <div className="tl-evento">{h.evento}</div>
              <div className="tl-meta">
                <span>{fmtDate(h.fecha)}</span>
                <span>{h.usuario}</span>
                {h.detalle && <span style={{ color: "var(--muted2)" }}>{h.detalle}</span>}
              </div>
            </div>
          </li>
        );
      })}
    </ul>
  );
}

// ─── FORM HELPERS ────────────────────────────────────────────────────────────
function FG({ label, hint, children }) {
  return (
    <div className="fg">
      {label && <label>{label}</label>}
      {children}
      {hint && <div className="field-hint">{hint}</div>}
    </div>
  );
}

function Select({ value, onChange, options, placeholder, disabled }) {
  return (
    <select value={value || ""} onChange={e => onChange(e.target.value)} disabled={disabled}>
      {placeholder && <option value="">{placeholder}</option>}
      {options.map(o => <option key={o.value || o} value={o.value || o}>{o.label || o}</option>)}
    </select>
  );
}

// ─── REQUISICION FORM (nueva / editar) ──────────────────────────────────────
function RequisicionForm({ initial, onSave, onCancel, usuario }) {
  const blankItem = () => ({ id: `tmp${Date.now()}${Math.random()}`, descripcion: "", cantidad: 1, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "", cantidad_aprobada: null });

  const [form, setForm] = useState({
    titulo: "", empresa: "Parana Logistica", base_buque: "", area: "",
    subarea: "", detalle_tecnico: "", tipo_requisicion: "", urgencia: "Normal",
    solicitado_por: usuario || "", fecha_necesaria: "", costo_estimado: "",
    moneda_estimada: "ARS", busco_alternativas: false, observaciones: "",
    ...(initial || {})
  });
  const [items, setItems] = useState(initial?.requisicion_items?.length ? initial.requisicion_items : [blankItem()]);
  const [saving, setSaving] = useState(false);

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const setItem = (i, k, v) => {
    const its = [...items];
    its[i] = { ...its[i], [k]: v };
    if (k === "cantidad" || k === "precio_unitario") {
      its[i].subtotal = (parseFloat(its[i].cantidad) || 0) * (parseFloat(its[i].precio_unitario) || 0);
    }
    setItems(its);
  };

  const bases = BASES_POR_EMPRESA[form.empresa] || [];
  const areas = AREAS_POR_EMPRESA[form.empresa] || [];
  const subareas = SUBAREA_TECNICA[form.empresa] || [];
  const detalles = DETALLE_TECNICO[form.subarea] || [];

  const handleSubmit = async () => {
    if (!form.titulo || !form.empresa || !form.base_buque || !form.area || !form.urgencia || !form.solicitado_por) {
      alert("Completá los campos obligatorios: Título, Empresa, Base/Buque, Área, Urgencia, Solicitado por");
      return;
    }
    if (!items.some(i => i.descripcion.trim())) {
      alert("Ingresá al menos un ítem con descripción");
      return;
    }
    setSaving(true);
    try {
      const cleanItems = items.filter(i => i.descripcion.trim()).map(({ id: _id, ...rest }) => rest);
      await onSave({ ...form, costo_estimado: form.costo_estimado ? parseFloat(form.costo_estimado) : null }, cleanItems);
    } finally { setSaving(false); }
  };

  return (
    <div>
      <div className="form-section">Datos de la Requisición</div>
      <div className="form-grid">
        <FG label="Título *"><input value={form.titulo} onChange={e => set("titulo", e.target.value)} placeholder="Ej: Compra Bujías Motor Principal" /></FG>
        <FG label="Tipo de Requisición"><Select value={form.tipo_requisicion} onChange={v => set("tipo_requisicion", v)} options={TIPOS_REQUISICION} placeholder="Seleccionar..." /></FG>
      </div>
      <div className="form-grid-3">
        <FG label="Empresa *"><Select value={form.empresa} onChange={v => { set("empresa", v); set("base_buque", ""); set("area", ""); }} options={EMPRESAS} /></FG>
        <FG label="Base / Buque *"><Select value={form.base_buque} onChange={v => set("base_buque", v)} options={bases} placeholder="Seleccionar..." /></FG>
        <FG label="Área *"><Select value={form.area} onChange={v => { set("area", v); set("subarea", ""); }} options={areas} placeholder="Seleccionar..." /></FG>
      </div>
      {form.area === "Tecnica" && (
        <div className="form-grid">
          <FG label="Sub-área Técnica"><Select value={form.subarea} onChange={v => { set("subarea", v); set("detalle_tecnico", ""); }} options={subareas} placeholder="Seleccionar..." /></FG>
          {detalles.length > 0 && <FG label="Detalle Técnico"><Select value={form.detalle_tecnico} onChange={v => set("detalle_tecnico", v)} options={detalles} placeholder="Seleccionar..." /></FG>}
        </div>
      )}
      <div className="form-grid">
        <FG label="Solicitado por *"><input value={form.solicitado_por} onChange={e => set("solicitado_por", e.target.value)} /></FG>
        <FG label="Fecha en que se necesita"><input type="date" value={form.fecha_necesaria} onChange={e => set("fecha_necesaria", e.target.value)} /></FG>
      </div>
      <div className="form-grid-3">
        <FG label="Urgencia *"><Select value={form.urgencia} onChange={v => set("urgencia", v)} options={URGENCIA_OPTIONS} /></FG>
        <FG label="Costo estimado"><input type="number" value={form.costo_estimado} onChange={e => set("costo_estimado", e.target.value)} placeholder="0" /></FG>
        <FG label="Moneda"><Select value={form.moneda_estimada} onChange={v => set("moneda_estimada", v)} options={["ARS", "USD"]} /></FG>
      </div>
      <div className="form-grid">
        <div className="checkbox-row">
          <input type="checkbox" id="alt" checked={form.busco_alternativas} onChange={e => set("busco_alternativas", e.target.checked)} />
          <label htmlFor="alt">Ya busqué alternativas / presupuestos previos</label>
        </div>
      </div>
      <FG label="Observaciones"><textarea value={form.observaciones} onChange={e => set("observaciones", e.target.value)} placeholder="Contexto adicional, referencias, etc." /></FG>

      <div className="form-section mt20">Ítems solicitados</div>
      <div className="table-wrap">
        <table className="items-edit">
          <thead>
            <tr>
              <th style={{ width: "35%" }}>Descripción *</th>
              <th style={{ width: "8%" }}>Cant.</th>
              <th style={{ width: "8%" }}>Unid.</th>
              <th style={{ width: "10%" }}>Stock disp.</th>
              <th style={{ width: "20%" }}>Proveedor sugerido</th>
              <th style={{ width: "12%" }}>Proyecto</th>
              <th style={{ width: "4%" }}></th>
            </tr>
          </thead>
          <tbody>
            {items.map((it, i) => (
              <tr key={it.id || i}>
                <td><input value={it.descripcion} onChange={e => setItem(i, "descripcion", e.target.value)} placeholder="Descripción del ítem" /></td>
                <td><input type="number" value={it.cantidad} onChange={e => setItem(i, "cantidad", e.target.value)} style={{ width: 60 }} /></td>
                <td><input value={it.unidad} onChange={e => setItem(i, "unidad", e.target.value)} style={{ width: 55 }} /></td>
                <td><input type="number" value={it.stock_disponible} onChange={e => setItem(i, "stock_disponible", e.target.value)} style={{ width: 65 }} /></td>
                <td><input value={it.proveedor_sugerido} onChange={e => setItem(i, "proveedor_sugerido", e.target.value)} /></td>
                <td><input value={it.proyecto || ""} onChange={e => setItem(i, "proyecto", e.target.value)} /></td>
                <td><button className="btn btn-ghost btn-xs" onClick={() => setItems(items.filter((_, j) => j !== i))}>✕</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button className="btn btn-ghost btn-sm mt8" onClick={() => setItems([...items, blankItem()])}>+ Agregar ítem</button>

      <div className="modal-footer" style={{ marginTop: 24, padding: "16px 0 0", borderTop: "1px solid var(--border)" }}>
        <button className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button className="btn btn-primary" onClick={handleSubmit} disabled={saving}>
          {saving ? "Guardando..." : (initial ? "Guardar cambios" : "Crear Requisición")}
        </button>
      </div>
    </div>
  );
}

// ─── MODAL DETALLE / GESTIÓN COMPRADOR ──────────────────────────────────────
function RequisicionModal({ req, proveedores, onClose, onUpdate, usuario }) {
  const [tab, setTab] = useState("detalle");
  const [gestion, setGestion] = useState({
    revisado_por: req.revisado_por || usuario,
    motivo_rechazo: req.motivo_rechazo || "",
    categoria_rechazo: req.categoria_rechazo || "",
    nro_oc: req.nro_oc || "",
    proveedor_elegido: req.proveedor_elegido || "",
    motivo_proveedor: req.motivo_proveedor || "",
    costo_real: req.costo_real || "",
    moneda_real: req.moneda_real || "ARS",
    plazo_pago: req.plazo_pago || "",
    fecha_entrega_prom: req.fecha_entrega_prom || "",
    fecha_entrega_real: req.fecha_entrega_real || "",
  });
  const [items, setItems] = useState(req.requisicion_items || []);
  const [saving, setSaving] = useState(false);

  const setG = (k, v) => setGestion(g => ({ ...g, [k]: v }));

  const setItemAp = (i, k, v) => {
    const its = [...items];
    its[i] = { ...its[i], [k]: v };
    if (k === "cantidad_aprobada" || k === "precio_unitario") {
      its[i].subtotal = (parseFloat(its[i].cantidad_aprobada ?? its[i].cantidad) || 0) * (parseFloat(its[i].precio_unitario) || 0);
    }
    setItems(its);
  };

  const totalReal = items.reduce((s, i) => s + (parseFloat(i.subtotal) || 0), 0);

  const doAction = async (nuevoStatus, evento, extraCambios = {}) => {
    setSaving(true);
    try {
      const ahora = new Date().toISOString();
      const cambios = {
        ...extraCambios,
        status: nuevoStatus,
        revisado_por: gestion.revisado_por,
        ...(nuevoStatus === "en_revision" && !req.fecha_revision ? { fecha_revision: ahora } : {}),
        ...(["aprobado_cotizar", "en_compra"].includes(nuevoStatus) ? { fecha_aprobacion: ahora } : {}),
        ...(nuevoStatus === "rechazado" ? {
          motivo_rechazo: gestion.motivo_rechazo,
          categoria_rechazo: gestion.categoria_rechazo,
          veces_devuelto: (req.veces_devuelto || 0) + 1,
        } : {}),
        ...(nuevoStatus === "en_compra" ? {
          nro_oc: gestion.nro_oc || api.nextOcNum?.() || req.nro_oc,
          proveedor_elegido: gestion.proveedor_elegido,
          motivo_proveedor: gestion.motivo_proveedor,
          costo_real: gestion.costo_real ? parseFloat(gestion.costo_real) : null,
          moneda_real: gestion.moneda_real,
          plazo_pago: gestion.plazo_pago,
          fecha_entrega_prom: gestion.fecha_entrega_prom || null,
        } : {}),
        ...(nuevoStatus === "entregado" ? { fecha_entrega_real: gestion.fecha_entrega_real || new Date().toISOString().slice(0, 10) } : {}),
      };
      await api.actualizarItems(req.id, items);
      const updated = await api.actualizarRequisicion(req.id, cambios, usuario, evento);
      onUpdate(updated);
      onClose();
    } finally { setSaving(false); }
  };

  const canReview = ["pendiente_revision"].includes(req.status);
  const canApprove = ["en_revision", "pendiente_revision"].includes(req.status);
  const canEmitOC = req.status === "aprobado_cotizar";
  const canDeliver = req.status === "en_compra";
  const canReject = ["pendiente_revision", "en_revision", "aprobado_cotizar"].includes(req.status);

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="modal-header">
          <div>
            <div className="modal-title text-mono">REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}</div>
            <div className="flex-gap mt4">
              <Badge status={req.status} />
              <UrgBadge urgencia={req.urgencia} />
              <span className="tag">{req.empresa}</span>
              <span className="tag">{req.base_buque}</span>
              {req.nro_oc && <span className="tag" style={{ color: "var(--accent)" }}>{req.nro_oc}</span>}
            </div>
          </div>
          <button className="modal-close" onClick={onClose}>✕</button>
        </div>

        <div className="modal-body" style={{ paddingBottom: 0 }}>
          <div className="flex-gap mb16" style={{ borderBottom: "1px solid var(--border)", paddingBottom: 0, gap: 0 }}>
            {["detalle", "gestion", "historial"].map(t => (
              <button key={t} onClick={() => setTab(t)} className="btn btn-ghost btn-sm" style={{
                borderRadius: "4px 4px 0 0", borderBottom: "none",
                borderColor: tab === t ? "var(--border)" : "transparent",
                borderBottomColor: tab === t ? "var(--surface)" : "transparent",
                color: tab === t ? "var(--accent)" : "var(--muted)",
                marginBottom: -1, background: tab === t ? "var(--surface)" : "transparent",
              }}>
                {{ detalle: "Detalle", gestion: "Gestión Comprador", historial: "Historial" }[t]}
              </button>
            ))}
          </div>

          {/* TAB DETALLE */}
          {tab === "detalle" && (
            <div>
              <div className="form-grid mb12">
                <div className="info-box"><div className="text-muted mb8" style={{ fontSize: 11, textTransform: "uppercase", fontFamily: "var(--mono)", letterSpacing: 1 }}>Área</div>{req.area}{req.subarea ? ` › ${req.subarea}` : ""}{req.detalle_tecnico ? ` › ${req.detalle_tecnico}` : ""}</div>
                <div className="info-box"><div className="text-muted mb8" style={{ fontSize: 11, textTransform: "uppercase", fontFamily: "var(--mono)", letterSpacing: 1 }}>Fechas</div>
                  <div>Solicitado: <strong>{fmtDate(req.created_at)}</strong></div>
                  {req.fecha_necesaria && <div>Necesario: <strong>{fmtDate(req.fecha_necesaria)}</strong></div>}
                </div>
              </div>
              <div className="form-grid mb12">
                <div className="info-box"><div className="text-muted mb8" style={{ fontSize: 11, textTransform: "uppercase", fontFamily: "var(--mono)", letterSpacing: 1 }}>Solicitante</div>{req.solicitado_por}</div>
                <div className="info-box">
                  <div className="text-muted mb8" style={{ fontSize: 11, textTransform: "uppercase", fontFamily: "var(--mono)", letterSpacing: 1 }}>Presupuesto estimado</div>
                  {req.costo_estimado ? <strong>{fmt(req.costo_estimado, req.moneda_estimada)}</strong> : <span className="text-muted">No especificado</span>}
                  {req.busco_alternativas && <div className="mt4" style={{ fontSize: 11, color: "var(--accent2)" }}>✓ Buscó alternativas previas</div>}
                </div>
              </div>
              {req.observaciones && <div className="info-box mb12">{req.observaciones}</div>}

              <div className="form-section">Ítems {canApprove ? "(editables)" : ""}</div>
              <div className="table-wrap">
                <table className={canApprove ? "items-edit" : ""}>
                  <thead>
                    <tr>
                      <th>#</th><th>Descripción</th><th>Cant. solicitada</th><th>Unid.</th>
                      <th>Stock</th><th>Cant. aprobada</th><th>Precio unit.</th><th>Subtotal</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((it, i) => (
                      <tr key={i}>
                        <td className="text-muted text-mono">{it.nro_linea}</td>
                        <td>{it.descripcion}</td>
                        <td className="text-mono">{it.cantidad}</td>
                        <td className="text-muted">{it.unidad}</td>
                        <td className="text-mono">{it.stock_disponible || 0}</td>
                        <td>
                          {canApprove
                            ? <input type="number" value={it.cantidad_aprobada ?? ""} onChange={e => setItemAp(i, "cantidad_aprobada", e.target.value)} placeholder={it.cantidad} style={{ width: 70 }} />
                            : <span className="text-mono">{it.cantidad_aprobada ?? <span className="text-muted">—</span>}</span>}
                        </td>
                        <td>
                          {canEmitOC
                            ? <input type="number" value={it.precio_unitario ?? ""} onChange={e => setItemAp(i, "precio_unitario", e.target.value)} style={{ width: 90 }} />
                            : <span className="text-mono">{it.precio_unitario ? fmt(it.precio_unitario) : "—"}</span>}
                        </td>
                        <td className="text-mono">{it.subtotal ? fmt(it.subtotal) : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                  {totalReal > 0 && (
                    <tfoot>
                      <tr>
                        <td colSpan={7} className="text-right text-mono" style={{ fontWeight: 600, paddingTop: 10 }}>TOTAL</td>
                        <td className="text-mono" style={{ fontWeight: 700, color: "var(--accent)" }}>{fmt(totalReal, gestion.moneda_real)}</td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          )}

          {/* TAB GESTIÓN */}
          {tab === "gestion" && (
            <div>
              {canReject && (
                <>
                  <div className="form-section">Rechazo / Devolución</div>
                  <div className="form-grid">
                    <FG label="Categoría de rechazo"><Select value={gestion.categoria_rechazo} onChange={v => setG("categoria_rechazo", v)} options={CATEGORIAS_RECHAZO} placeholder="Seleccionar..." /></FG>
                    <FG label="Motivo detallado"><input value={gestion.motivo_rechazo} onChange={e => setG("motivo_rechazo", e.target.value)} placeholder="Explicar qué falta o qué está mal" /></FG>
                  </div>
                </>
              )}

              {(canEmitOC || req.status === "en_compra") && (
                <>
                  <div className="form-section">Orden de Compra</div>
                  <div className="form-grid">
                    <FG label="N° OC"><input value={gestion.nro_oc} onChange={e => setG("nro_oc", e.target.value)} placeholder="OC-0001" /></FG>
                    <FG label="Proveedor elegido">
                      <select value={gestion.proveedor_elegido} onChange={e => setG("proveedor_elegido", e.target.value)}>
                        <option value="">Seleccionar o escribir...</option>
                        {proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}
                      </select>
                    </FG>
                  </div>
                  <FG label="Justificación elección proveedor" hint="¿Por qué este proveedor? (precio, disponibilidad, calidad, relación comercial)">
                    <textarea value={gestion.motivo_proveedor} onChange={e => setG("motivo_proveedor", e.target.value)} />
                  </FG>
                  <div className="form-grid-3">
                    <FG label="Costo real"><input type="number" value={gestion.costo_real} onChange={e => setG("costo_real", e.target.value)} /></FG>
                    <FG label="Moneda"><Select value={gestion.moneda_real} onChange={v => setG("moneda_real", v)} options={["ARS", "USD"]} /></FG>
                    <FG label="Plazo de pago"><Select value={gestion.plazo_pago} onChange={v => setG("plazo_pago", v)} options={PLAZO_PAGO_OPTIONS} placeholder="Seleccionar..." /></FG>
                  </div>
                  <div className="form-grid">
                    <FG label="Fecha entrega prometida"><input type="date" value={gestion.fecha_entrega_prom} onChange={e => setG("fecha_entrega_prom", e.target.value)} /></FG>
                    {req.status === "en_compra" && <FG label="Fecha entrega real"><input type="date" value={gestion.fecha_entrega_real} onChange={e => setG("fecha_entrega_real", e.target.value)} /></FG>}
                  </div>

                  {req.costo_estimado && gestion.costo_real && (
                    <div className="info-box accent mt8">
                      Desvío presupuestario: <strong style={{ color: parseFloat(gestion.costo_real) > req.costo_estimado ? "var(--danger)" : "var(--accent2)" }}>
                        {((parseFloat(gestion.costo_real) - req.costo_estimado) / req.costo_estimado * 100).toFixed(1)}%
                      </strong>
                      {" "}({fmt(req.costo_estimado, req.moneda_estimada)} estimado → {fmt(parseFloat(gestion.costo_real), gestion.moneda_real)} real)
                    </div>
                  )}
                </>
              )}

              {req.status === "en_compra" && (
                <>
                  <div className="form-section mt16">Recepción</div>
                  <FG label="Fecha de entrega real"><input type="date" value={gestion.fecha_entrega_real} onChange={e => setG("fecha_entrega_real", e.target.value)} /></FG>
                </>
              )}

              {["cerrado", "entregado"].includes(req.status) && (
                <div className="info-box mt8">Esta requisición está cerrada. Solo lectura.</div>
              )}
            </div>
          )}

          {/* TAB HISTORIAL */}
          {tab === "historial" && (
            <div>
              <Timeline historial={req.requisicion_historial} />
              {req.veces_devuelto > 0 && (
                <div className="info-box mt16" style={{ borderLeft: "3px solid var(--warn)" }}>
                  ⚠ Esta requisición fue devuelta <strong>{req.veces_devuelto} vez{req.veces_devuelto > 1 ? "ces" : ""}</strong>
                  {req.categoria_rechazo && ` — Categoría: ${req.categoria_rechazo}`}
                </div>
              )}
              {req.dias_solicitud_revision != null && (
                <div className="mt16">
                  <KpiBar label="Solicitud → Revisión" value={req.dias_solicitud_revision} max={10} color="var(--accent)" suffix=" días" />
                  <KpiBar label="Revisión → Aprobación" value={req.dias_revision_aprobacion} max={5} color="var(--accent2)" suffix=" días" />
                  <KpiBar label="Aprobación → Entrega" value={req.dias_aprobacion_entrega} max={30} color="var(--warn)" suffix=" días" />
                </div>
              )}
            </div>
          )}
        </div>

        <div className="modal-footer">
          <button className="btn btn-ghost btn-sm" onClick={onClose}>Cerrar</button>
          {canReject && <button className="btn btn-danger btn-sm" onClick={() => doAction("rechazado", `Devuelto — ${gestion.categoria_rechazo || "Sin categoría"}`)} disabled={saving || !gestion.motivo_rechazo}>Rechazar / Devolver</button>}
          {canReview && <button className="btn btn-warn btn-sm" onClick={() => doAction("en_revision", "Marcado en revisión")} disabled={saving}>Tomar revisión</button>}
          {canApprove && <button className="btn btn-primary btn-sm" onClick={() => doAction("aprobado_cotizar", "Aprobado para cotizar")} disabled={saving}>Aprobar para cotizar</button>}
          {canEmitOC && <button className="btn btn-success btn-sm" onClick={() => doAction("en_compra", `OC emitida — ${gestion.nro_oc || "pendiente"}`, { nro_oc: gestion.nro_oc || api.nextOcNum?.() })} disabled={saving || !gestion.proveedor_elegido}>Emitir OC</button>}
          {canDeliver && <button className="btn btn-success btn-sm" onClick={() => doAction("entregado", "Entrega confirmada")} disabled={saving}>Confirmar entrega</button>}
        </div>
      </div>
    </div>
  );
}

// ─── PAGE: LISTA DE REQUISICIONES ───────────────────────────────────────────
function PageRequisiciones({ onNew, notify, usuario }) {
  const [reqs, setReqs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filter, setFilter] = useState("all");
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(null);
  const [proveedores, setProveedores] = useState([]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [data, provs] = await Promise.all([api.getRequisiciones(), api.getProveedores()]);
      setReqs(data);
      setProveedores(provs);
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  const filtered = reqs.filter(r => {
    if (filter !== "all" && r.status !== filter) return false;
    if (search) {
      const s = search.toLowerCase();
      return r.titulo?.toLowerCase().includes(s) || r.solicitado_por?.toLowerCase().includes(s) ||
        r.empresa?.toLowerCase().includes(s) || r.base_buque?.toLowerCase().includes(s) ||
        r.nro_oc?.toLowerCase().includes(s);
    }
    return true;
  });

  const counts = {};
  reqs.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
  const pendientes = (counts.pendiente_revision || 0) + (counts.en_revision || 0);

  const handleUpdate = (updated) => {
    setReqs(prev => prev.map(r => r.id === updated.id ? { ...r, ...updated } : r));
    notify("Requisición actualizada", "success");
  };

  const filterOptions = [
    { k: "all", label: "Todas" },
    { k: "pendiente_revision", label: "Pendientes" },
    { k: "en_revision", label: "En revisión" },
    { k: "aprobado_cotizar", label: "Aprobadas" },
    { k: "en_compra", label: "En compra" },
    { k: "entregado", label: "Entregadas" },
    { k: "rechazado", label: "Rechazadas" },
  ];

  return (
    <div>
      <div className="stats">
        <div className="stat"><div className="stat-label">Total</div><div className="stat-value v-blue">{reqs.length}</div></div>
        <div className="stat"><div className="stat-label">Por revisar</div><div className="stat-value v-amber">{pendientes}</div><div className="stat-sub">Requieren acción</div></div>
        <div className="stat"><div className="stat-label">En compra</div><div className="stat-value v-purple">{counts.en_compra || 0}</div></div>
        <div className="stat"><div className="stat-label">Entregadas</div><div className="stat-value v-green">{(counts.entregado || 0) + (counts.cerrado || 0)}</div></div>
      </div>

      <div className="card">
        <div className="flex-between mb12">
          <div className="filters" style={{ margin: 0 }}>
            {filterOptions.map(f => (
              <span key={f.k} className={`filter-chip ${filter === f.k ? "active" : ""}`} onClick={() => setFilter(f.k)}>
                {f.label}{f.k !== "all" && counts[f.k] ? ` (${counts[f.k]})` : ""}
              </span>
            ))}
          </div>
          <div className="flex-gap">
            <input className="search-input" placeholder="Buscar..." value={search} onChange={e => setSearch(e.target.value)} />
            <button className="btn btn-primary" onClick={onNew}>+ Nueva</button>
          </div>
        </div>

        {loading ? (
          <div className="loading"><span className="spin">◌</span> Cargando...</div>
        ) : filtered.length === 0 ? (
          <div className="empty-state"><div className="empty-icon">📭</div><div>Sin requisiciones</div></div>
        ) : (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>N°</th><th>Título</th><th>Empresa / Buque</th><th>Área</th>
                  <th>Urgencia</th><th>Solicitado</th><th>Fecha</th>
                  <th>Estimado</th><th>Real</th><th>Status</th><th></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map(r => (
                  <tr key={r.id} className="tr-clickable" onClick={() => setSelected(r)}>
                    <td className="text-mono" style={{ color: "var(--accent)", fontSize: 12 }}>
                      {String(r.nro_solicitud).padStart(4, "0")}
                      {r.nro_oc && <div style={{ fontSize: 10, color: "var(--muted)" }}>{r.nro_oc}</div>}
                    </td>
                    <td style={{ maxWidth: 200 }}>
                      <div style={{ fontWeight: 500 }}>{r.titulo}</div>
                      {r.veces_devuelto > 0 && <span style={{ fontSize: 10, color: "var(--warn)" }}>↩ {r.veces_devuelto}x devuelto</span>}
                    </td>
                    <td style={{ fontSize: 12 }}><div>{r.empresa}</div><div className="text-muted">{r.base_buque}</div></td>
                    <td style={{ fontSize: 12 }}>{r.area}{r.subarea ? <div className="text-muted">{r.subarea}</div> : null}</td>
                    <td><UrgBadge urgencia={r.urgencia} /></td>
                    <td style={{ fontSize: 12 }}>{r.solicitado_por}</td>
                    <td className="text-mono text-muted" style={{ fontSize: 11 }}>{fmtDate(r.created_at)}</td>
                    <td className="text-mono" style={{ fontSize: 12 }}>{r.costo_estimado ? fmt(r.costo_estimado, r.moneda_estimada) : <span className="text-muted">—</span>}</td>
                    <td className="text-mono" style={{ fontSize: 12 }}>{r.costo_real ? fmt(r.costo_real, r.moneda_real) : <span className="text-muted">—</span>}</td>
                    <td><Badge status={r.status} /></td>
                    <td><button className="btn btn-ghost btn-xs" onClick={e => { e.stopPropagation(); setSelected(r); }}>→</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {selected && (
        <RequisicionModal
          req={selected}
          proveedores={proveedores}
          usuario={usuario}
          onClose={() => setSelected(null)}
          onUpdate={r => { handleUpdate(r); setSelected(null); }}
        />
      )}
    </div>
  );
}

// ─── PAGE: NUEVA REQUISICIÓN ─────────────────────────────────────────────────
function PageNueva({ onSaved, onCancel, notify, usuario, fromExcel }) {
  const fileRef = useRef();
  const [excelItems, setExcelItems] = useState(fromExcel || null);
  const [drag, setDrag] = useState(false);

  const handleFile = async (file) => {
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const items = rows.map(r => ({
        descripcion: String(r["Descripcion"] || r["Descripción"] || r["descripcion"] || r["Item"] || Object.values(r)[2] || "").trim(),
        cantidad: parseFloat(r["Cant."] || r["Cantidad"] || r["cantidad"] || 1) || 1,
        unidad: String(r["Unid."] || r["Unidad"] || r["unidad"] || "Uni"),
        stock_disponible: parseFloat(r["Stock"] || 0) || 0,
        proveedor_sugerido: String(r["Proveedor Sugerido"] || r["proveedor_sugerido"] || ""),
        proyecto: String(r["Proyecto"] || r["proyecto"] || ""),
      })).filter(i => i.descripcion);
      if (!items.length) { notify("No se encontraron ítems en el Excel", "error"); return; }
      setExcelItems(items);
      notify(`${items.length} ítems importados del Excel`, "success");
    } catch { notify("Error al leer el Excel", "error"); }
  };

  const handleSave = async (form, items) => {
    const created = await api.crearRequisicion(form, items, usuario);
    notify(`Requisición REQ-${String(created.nro_solicitud).padStart(4, "0")} creada`, "success");
    onSaved();
  };

  if (excelItems !== null) {
    return (
      <div className="card">
        <div className="card-title">Nueva Requisición — {excelItems.length} ítems desde Excel</div>
        <RequisicionForm
          initial={{ requisicion_items: excelItems }}
          onSave={handleSave}
          onCancel={onCancel}
          usuario={usuario}
        />
      </div>
    );
  }

  return (
    <div>
      <div className="card" style={{ marginBottom: 16 }}>
        <div className="card-title">Importar desde Excel</div>
        <div
          className={`upload-zone ${drag ? "drag" : ""}`}
          onDragOver={e => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={e => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]); }}
          onClick={() => fileRef.current.click()}
        >
          <div className="upload-icon">📥</div>
          <div className="upload-title">Arrastrá el Excel de requisición acá</div>
          <div className="upload-sub">o hacé click para seleccionar · .xlsx / .xls</div>
          <div className="upload-sub mt4" style={{ fontSize: 11, color: "var(--muted2)" }}>Columnas: N° · Cant. · Unid. · Descripción · Stock · Proveedor Sugerido</div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
      </div>

      <div className="card">
        <div className="card-title">O cargar manualmente</div>
        <RequisicionForm onSave={handleSave} onCancel={onCancel} usuario={usuario} />
      </div>
    </div>
  );
}

// ─── PAGE: KPIs / REPORTES ───────────────────────────────────────────────────
function PageReportes() {
  const [reqs, setReqs] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    api.getRequisiciones().then(d => { setReqs(d); setLoading(false); });
  }, []);

  if (loading) return <div className="loading"><span className="spin">◌</span> Cargando reportes...</div>;

  const total = reqs.length;
  const urgencias = reqs.filter(r => r.urgencia === "Critica").length;
  const pctUrgencias = total ? Math.round(urgencias / total * 100) : 0;
  const rechazadas = reqs.filter(r => r.status === "rechazado").length;
  const conIdaVuelta = reqs.filter(r => r.veces_devuelto > 0).length;
  const avgRevision = reqs.filter(r => r.dias_solicitud_revision != null).reduce((s, r) => s + r.dias_solicitud_revision, 0) / (reqs.filter(r => r.dias_solicitud_revision != null).length || 1);

  // Por solicitante
  const bySol = {};
  reqs.forEach(r => {
    if (!bySol[r.solicitado_por]) bySol[r.solicitado_por] = { total: 0, urgentes: 0, devueltas: 0 };
    bySol[r.solicitado_por].total++;
    if (r.urgencia === "Critica") bySol[r.solicitado_por].urgentes++;
    if (r.veces_devuelto > 0) bySol[r.solicitado_por].devueltas++;
  });

  // Por categoría rechazo
  const byRechazo = {};
  reqs.filter(r => r.categoria_rechazo).forEach(r => {
    byRechazo[r.categoria_rechazo] = (byRechazo[r.categoria_rechazo] || 0) + 1;
  });

  // Por proveedor
  const byProv = {};
  reqs.filter(r => r.proveedor_elegido).forEach(r => {
    if (!byProv[r.proveedor_elegido]) byProv[r.proveedor_elegido] = { count: 0, total: 0, cur: r.moneda_real || "ARS" };
    byProv[r.proveedor_elegido].count++;
    byProv[r.proveedor_elegido].total += r.costo_real || 0;
  });

  return (
    <div>
      <div className="stats">
        <div className="stat"><div className="stat-label">Total requisiciones</div><div className="stat-value v-blue">{total}</div></div>
        <div className="stat"><div className="stat-label">% Urgencias críticas</div><div className="stat-value v-red">{pctUrgencias}%</div><div className="stat-sub">{urgencias} de {total}</div></div>
        <div className="stat"><div className="stat-label">Con ida y vuelta</div><div className="stat-value v-amber">{conIdaVuelta}</div><div className="stat-sub">Lean: calidad del pedido</div></div>
        <div className="stat"><div className="stat-label">Avg días revisión</div><div className="stat-value v-green">{avgRevision ? avgRevision.toFixed(1) : "—"}</div><div className="stat-sub">Velocidad compras</div></div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div className="card">
          <div className="card-title">Planificación por solicitante</div>
          <table>
            <thead><tr><th>Solicitante</th><th>Pedidos</th><th>Críticos</th><th>% Criticos</th><th>Devueltos</th></tr></thead>
            <tbody>
              {Object.entries(bySol).sort((a, b) => b[1].total - a[1].total).map(([sol, d]) => (
                <tr key={sol}>
                  <td>{sol}</td>
                  <td className="text-mono">{d.total}</td>
                  <td className="text-mono" style={{ color: d.urgentes > 0 ? "var(--danger)" : "inherit" }}>{d.urgentes}</td>
                  <td>
                    <div className="flex-gap">
                      <div style={{ width: `${d.total ? d.urgentes / d.total * 60 : 0}px`, height: 5, background: "var(--danger)", borderRadius: 3 }} />
                      <span className="text-mono" style={{ fontSize: 11 }}>{d.total ? Math.round(d.urgentes / d.total * 100) : 0}%</span>
                    </div>
                  </td>
                  <td className="text-mono" style={{ color: d.devueltas > 0 ? "var(--warn)" : "inherit" }}>{d.devueltas}</td>
                </tr>
              ))}
              {!Object.keys(bySol).length && <tr><td colSpan={5} className="text-muted" style={{ textAlign: "center", padding: 20 }}>Sin datos</td></tr>}
            </tbody>
          </table>
        </div>

        <div className="card">
          <div className="card-title">Motivos de rechazo / devolución</div>
          {Object.entries(byRechazo).length === 0
            ? <div className="empty-state" style={{ padding: 20 }}><div>Sin rechazos registrados</div></div>
            : Object.entries(byRechazo).sort((a, b) => b[1] - a[1]).map(([cat, n]) => (
              <KpiBar key={cat} label={cat} value={n} max={Math.max(...Object.values(byRechazo))} color="var(--danger)" />
            ))
          }
        </div>

        <div className="card">
          <div className="card-title">Proveedores utilizados</div>
          <table>
            <thead><tr><th>Proveedor</th><th>OCs</th><th>Comprado</th></tr></thead>
            <tbody>
              {Object.entries(byProv).sort((a, b) => b[1].total - a[1].total).map(([prov, d]) => (
                <tr key={prov}>
                  <td>{prov}</td>
                  <td className="text-mono">{d.count}</td>
                  <td className="text-mono">{fmt(d.total, d.cur)}</td>
                </tr>
              ))}
              {!Object.keys(byProv).length && <tr><td colSpan={3} className="text-muted" style={{ textAlign: "center", padding: 20 }}>Sin datos</td></tr>}
            </tbody>
          </table>
        </div>

        <div className="card">
          <div className="card-title">Velocidad del departamento</div>
          <KpiBar label="Promedio solicitud → revisión" value={avgRevision.toFixed(1)} max={10} color="var(--accent)" suffix=" días" />
          <div className="info-box mt12 accent">
            <div style={{ fontSize: 12, lineHeight: 1.6 }}>
              <strong>Objetivo Lean:</strong><br />
              · Solicitud → Revisión: &lt; 1 día<br />
              · Revisión → Aprobación: &lt; 2 días<br />
              · % Urgencias críticas: &lt; 20%<br />
              · Devoluciones por calidad: &lt; 10%
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── PAGE: PROVEEDORES ───────────────────────────────────────────────────────
function PageProveedores({ notify }) {
  const [provs, setProvs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(false);
  const [form, setForm] = useState({ nombre: "", rubro: "", contacto: "", email: "", telefono: "", notas: "" });

  useEffect(() => { api.getProveedores().then(d => { setProvs(d); setLoading(false); }); }, []);

  const handleSave = async () => {
    if (!form.nombre) return;
    const nuevo = await api.crearProveedor(form);
    setProvs(p => [...p, nuevo]);
    setModal(false);
    setForm({ nombre: "", rubro: "", contacto: "", email: "", telefono: "", notas: "" });
    notify("Proveedor agregado", "success");
  };

  return (
    <div>
      <div className="card">
        <div className="card-title">
          Maestro de Proveedores
          <button className="btn btn-primary btn-sm" onClick={() => setModal(true)}>+ Agregar</button>
        </div>
        {loading ? <div className="loading"><span className="spin">◌</span></div> : (
          <div className="table-wrap">
            <table>
              <thead><tr><th>Nombre</th><th>Rubro</th><th>Contacto</th><th>Email</th><th>Teléfono</th></tr></thead>
              <tbody>
                {provs.map(p => (
                  <tr key={p.id}>
                    <td style={{ fontWeight: 500 }}>{p.nombre}</td>
                    <td className="text-muted">{p.rubro || "—"}</td>
                    <td>{p.contacto || "—"}</td>
                    <td className="text-mono" style={{ fontSize: 12 }}>{p.email || "—"}</td>
                    <td className="text-mono" style={{ fontSize: 12 }}>{p.telefono || "—"}</td>
                  </tr>
                ))}
                {!provs.length && <tr><td colSpan={5}><div className="empty-state">Sin proveedores</div></td></tr>}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {modal && (
        <div className="overlay" onClick={e => e.target === e.currentTarget && setModal(false)}>
          <div className="modal">
            <div className="modal-header">
              <div className="modal-title">Nuevo Proveedor</div>
              <button className="modal-close" onClick={() => setModal(false)}>✕</button>
            </div>
            <div className="modal-body">
              <div className="form-grid">
                <FG label="Nombre *"><input value={form.nombre} onChange={e => setForm(f => ({ ...f, nombre: e.target.value }))} /></FG>
                <FG label="Rubro"><input value={form.rubro} onChange={e => setForm(f => ({ ...f, rubro: e.target.value }))} /></FG>
                <FG label="Contacto"><input value={form.contacto} onChange={e => setForm(f => ({ ...f, contacto: e.target.value }))} /></FG>
                <FG label="Email"><input value={form.email} onChange={e => setForm(f => ({ ...f, email: e.target.value }))} /></FG>
                <FG label="Teléfono"><input value={form.telefono} onChange={e => setForm(f => ({ ...f, telefono: e.target.value }))} /></FG>
              </div>
              <FG label="Notas"><textarea value={form.notas} onChange={e => setForm(f => ({ ...f, notas: e.target.value }))} /></FG>
            </div>
            <div className="modal-footer">
              <button className="btn btn-ghost" onClick={() => setModal(false)}>Cancelar</button>
              <button className="btn btn-primary" onClick={handleSave}>Guardar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  ROOT APP
// ═══════════════════════════════════════════════════════════════════════════
const USUARIO = "Comprador";

export default function App() {
  const [page, setPage] = useState("requisiciones");
  const [notif, setNotif] = useState(null);
  const [reqs, setReqs] = useState([]);

  const notify = useCallback((text, type = "info") => {
    setNotif({ text, type });
    setTimeout(() => setNotif(null), 4000);
  }, []);

  // Cargar conteo para badge
  useEffect(() => {
    api.getRequisiciones().then(d => setReqs(d));
  }, [page]);

  const pendientes = reqs.filter(r => ["pendiente_revision", "en_revision"].includes(r.status)).length;

  const nav = [
    { id: "requisiciones", icon: "📋", label: "Requisiciones", badge: pendientes || null },
    { id: "nueva", icon: "✚", label: "Nueva Requisición" },
    { id: "reportes", icon: "📊", label: "KPIs & Reportes" },
    { id: "proveedores", icon: "🏭", label: "Proveedores" },
  ];

  const titles = {
    requisiciones: "REQUISICIONES DE COMPRA",
    nueva: "NUEVA REQUISICIÓN",
    reportes: "KPIs & REPORTES LEAN",
    proveedores: "MAESTRO DE PROVEEDORES",
  };

  return (
    <>
      <style>{CSS}</style>
      <div className="app">
        <nav className="sidebar">
          <div className="sidebar-header">
            <div className="sidebar-logo">◈ Compras</div>
            <div className="sidebar-sub">Terra Mare · Parana Logística</div>
          </div>
          <div className="nav-section">Módulos</div>
          {nav.map(n => (
            <div key={n.id} className={`nav-item ${page === n.id ? "active" : ""}`} onClick={() => setPage(n.id)}>
              <span>{n.icon}</span>
              <span>{n.label}</span>
              {n.badge ? <span className={`nav-badge ${n.badge > 3 ? "red" : ""}`}>{n.badge}</span> : null}
            </div>
          ))}
          <div style={{ flex: 1 }} />
          <div style={{ padding: "16px", fontSize: 11, color: "var(--muted2)", fontFamily: "var(--mono)", borderTop: "1px solid var(--border)" }}>
            v1.0 MVP<br />
            <span style={{ color: "var(--border2)" }}>Xubio API: pendiente</span><br />
            <span style={{ color: "var(--border2)" }}>Power Automate: pendiente</span>
          </div>
        </nav>

        <div className="main">
          <div className="topbar">
            <div className="topbar-left">
              <div className="topbar-title">{titles[page]}</div>
            </div>
            <div className="topbar-right">
              <span style={{ fontSize: 12, color: "var(--muted)" }}>👤 {USUARIO}</span>
            </div>
          </div>
          <div className="content">
            <div className="demo-banner">
              
            </div>

            {page === "requisiciones" && (
              <PageRequisiciones
                onNew={() => setPage("nueva")}
                notify={notify}
                usuario={USUARIO}
              />
            )}
            {page === "nueva" && (
              <PageNueva
                onSaved={() => setPage("requisiciones")}
                onCancel={() => setPage("requisiciones")}
                notify={notify}
                usuario={USUARIO}
              />
            )}
            {page === "reportes" && <PageReportes />}
            {page === "proveedores" && <PageProveedores notify={notify} />}
          </div>
        </div>
      </div>
      <Notif msg={notif} onClose={() => setNotif(null)} />
    </>
  );
}
