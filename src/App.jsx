import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  EMPRESAS, BASES_POR_EMPRESA, AREAS_POR_EMPRESA, SUBAREA_TECNICA,
  DETALLE_TECNICO, TIPOS_REQUISICION, URGENCIA_OPTIONS, PLAZO_PAGO_OPTIONS,
  URGENCIA_COLOR, STATUS_COLOR
} from "./lib/catalogos";
import { supabase } from "./lib/supabase";

// ─── CONSTANTS ───────────────────────────────────────────────────────────────
const USUARIO = "Comprador";

const CATEGORIAS_RECHAZO = [
  "Precio fuera de presupuesto",
  "Necesidad cancelada",
  "Descripción incompleta",
  "Duplicado",
  "Sin proveedor disponible",
  "Otro",
];

const TRACKER_STATUS = {
  en_cotizacion: { label: "En cotización", color: "b-amber" },
  oc_emitida:    { label: "OC Emitida",    color: "b-blue" },
  en_transito:   { label: "En tránsito",   color: "b-purple" },
  entregado:     { label: "Entregado",     color: "b-green" },
  archivado:     { label: "Archivado",     color: "b-gray" },
};

const GRUPOS_OPCIONES = ["A", "B", "C", "D", "E"];

// ─── UTILS ───────────────────────────────────────────────────────────────────
const fmt = (n, cur = "ARS") =>
  n != null ? new Intl.NumberFormat("es-AR", { style: "currency", currency: cur, maximumFractionDigits: 0 }).format(n) : "—";
const fmtDate = d => d ? new Date(d).toLocaleDateString("es-AR") : "—";

// ─── API ─────────────────────────────────────────────────────────────────────
const api = {
  async getRequisiciones(filtros = {}) {
    let q = supabase.from("requisiciones").select("*, requisicion_items(*), requisicion_historial(*)").order("created_at", { ascending: false });
    if (filtros.status) q = q.eq("status", filtros.status);
    if (filtros.empresa) q = q.eq("empresa", filtros.empresa);
    if (filtros.statuses) q = q.in("status", filtros.statuses);
    const { data, error } = await q;
    if (error) throw error;
    return data || [];
  },
  async getRequisicion(id) {
    const { data, error } = await supabase.from("requisiciones").select("*, requisicion_items(*), requisicion_historial(*)").eq("id", id).single();
    if (error) throw error;
    return data;
  },
  async crearRequisicion(req, items) {
    const { data: nueva, error } = await supabase.from("requisiciones").insert([{ ...req, status: "pendiente_revision" }]).select().single();
    if (error) throw error;
    if (items?.length) {
      await supabase.from("requisicion_items").insert(items.map((it, i) => ({ ...it, requisicion_id: nueva.id, nro_linea: i + 1 })));
    }
    await supabase.from("requisicion_historial").insert([{ requisicion_id: nueva.id, evento: "Requisición creada", usuario: USUARIO, status_nuevo: "pendiente_revision" }]);
    return nueva;
  },
  async actualizarRequisicion(id, cambios, evento, detalle) {
    const { data, error } = await supabase.from("requisiciones").update({ ...cambios, updated_at: new Date().toISOString() }).eq("id", id).select().single();
    if (error) throw error;
    if (evento) await supabase.from("requisicion_historial").insert([{ requisicion_id: id, evento, usuario: USUARIO, detalle, status_nuevo: cambios.status }]);
    return data;
  },
  async actualizarItems(reqId, items) {
    await supabase.from("requisicion_items").delete().eq("requisicion_id", reqId);
    if (items?.length) await supabase.from("requisicion_items").insert(items.map((it, i) => ({ ...it, requisicion_id: reqId, nro_linea: i + 1 })));
  },
  async getTrackerLineas(filtros = {}) {
    let q = supabase.from("tracker_lineas").select("*, requisiciones(nro_solicitud, titulo, empresa, base_buque, area, subarea, urgencia, solicitado_por, fecha_necesaria, costo_estimado, moneda_estimada, tipo_requisicion, observaciones)").order("created_at", { ascending: false });
    if (filtros.status) q = q.eq("status", filtros.status);
    if (filtros.statuses) q = q.in("status", filtros.statuses);
    if (filtros.proveedor) q = q.eq("proveedor_elegido", filtros.proveedor);
    const { data, error } = await q;
    if (error) throw error;
    return data || [];
  },
  async crearTrackerLineas(requisicionId, lineas) {
    const { error } = await supabase.from("tracker_lineas").insert(lineas.map(l => ({ ...l, requisicion_id: requisicionId })));
    if (error) throw error;
  },
  async actualizarTrackerLinea(id, cambios) {
    const { data, error } = await supabase.from("tracker_lineas").update({ ...cambios, updated_at: new Date().toISOString() }).eq("id", id).select().single();
    if (error) throw error;
    return data;
  },
  async getProveedores() {
    const { data, error } = await supabase.from("proveedores").select("*").eq("activo", true).order("nombre");
    if (error) throw error;
    return data || [];
  },
  async crearProveedor(prov) {
    const { data, error } = await supabase.from("proveedores").insert([prov]).select().single();
    if (error) throw error;
    return data;
  },
  async actualizarProveedor(id, cambios) {
    const { data, error } = await supabase.from("proveedores").update(cambios).eq("id", id).select().single();
    if (error) throw error;
    return data;
  },
  nextOcNum() { return `OC-${String(Date.now()).slice(-4)}`; }
};

// ─── CSS — TERRA MARE BRAND ───────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --tm-navy:   #213363;
  --tm-blue:   #235C96;
  --tm-mid:    #6381A7;
  --tm-light:  #A5B5CC;
  --tm-white:  #FFFFFF;
  --tm-black:  #000000;
  --bg:        #F0F4F8;
  --surface:   #FFFFFF;
  --surface2:  #F5F7FA;
  --surface3:  #EAF0F6;
  --border:    #D6E0ED;
  --border2:   #B0C4D8;
  --text:      #213363;
  --muted:     #6381A7;
  --muted2:    #8FA3BC;
  --accent:    #235C96;
  --accent2:   #1E7E4A;
  --warn:      #B07D0A;
  --danger:    #C0392B;
  --purple:    #6B4FA0;
  --teal:      #1A7A6E;
  --orange:    #C05621;
  --mono:      'DM Mono', monospace;
  --sans:      'Montserrat', sans-serif;
  --r:6px;--r2:10px;
}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.5;min-height:100vh}
.app{display:flex;min-height:100vh}

/* SIDEBAR */
.sidebar{width:235px;min-width:235px;background:var(--tm-navy);border-right:none;display:flex;flex-direction:column;box-shadow:2px 0 8px rgba(33,51,99,.15)}
.sidebar-header{padding:0 0 0;border-bottom:1px solid rgba(255,255,255,.1)}
.sidebar-logo-wrap{padding:20px 18px 16px;display:flex;align-items:center;gap:12px}
.sidebar-logo-img{width:36px;height:36px;object-fit:contain;border-radius:50%}
.sidebar-logo-text{display:flex;flex-direction:column}
.sidebar-logo-main{font-family:var(--sans);font-size:13px;font-weight:700;color:#fff;letter-spacing:2px;text-transform:uppercase;line-height:1.2}
.sidebar-logo-sub{font-size:9px;color:rgba(255,255,255,.5);margin-top:2px;letter-spacing:.5px}
.nav-section{padding:12px 18px 4px;font-family:var(--mono);font-size:9px;letter-spacing:2px;color:rgba(255,255,255,.35);text-transform:uppercase}
.ni{display:flex;align-items:center;gap:9px;padding:7px 18px;font-size:12px;font-weight:500;cursor:pointer;color:rgba(255,255,255,.6);border-left:3px solid transparent;transition:all .12s;user-select:none}
.ni:hover{color:#fff;background:rgba(255,255,255,.06)}
.ni.active{color:#fff;border-left-color:var(--tm-light);background:rgba(255,255,255,.1);font-weight:600}
.ni.sub{padding-left:32px;font-size:11px;font-weight:400}
.ni.sub.active{font-weight:600}
.ni-icon{font-size:13px;width:16px;text-align:center;flex-shrink:0}
.ni-badge{margin-left:auto;background:var(--danger);color:#fff;font-family:var(--mono);font-size:9px;font-weight:700;padding:1px 6px;border-radius:10px;min-width:18px;text-align:center}
.ni-badge.amber{background:#B07D0A;color:#fff}
.ni-badge.gray{background:rgba(255,255,255,.2);color:rgba(255,255,255,.7)}

/* MAIN */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0}
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 1px 3px rgba(33,51,99,.06)}
.topbar-title{font-family:var(--sans);font-size:12px;font-weight:600;letter-spacing:1px;color:var(--tm-navy);text-transform:uppercase}
.content{flex:1;overflow-y:auto;padding:24px 28px;background:var(--bg)}

/* CARDS */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:20px;margin-bottom:16px;box-shadow:0 1px 4px rgba(33,51,99,.06)}
.card-title{font-family:var(--sans);font-size:10px;font-weight:600;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between}

/* STATS */
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}
.stat{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:16px 18px;box-shadow:0 1px 4px rgba(33,51,99,.06)}
.stat-label{font-size:10px;color:var(--muted);font-weight:600;letter-spacing:.5px;margin-bottom:6px;text-transform:uppercase}
.stat-value{font-family:var(--mono);font-size:28px;font-weight:600}
.va{color:var(--tm-blue)}.vg{color:var(--accent2)}.vr{color:var(--danger)}.vp{color:var(--purple)}.vm{color:var(--warn)}.vgr{color:var(--muted)}

/* TABLE */
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{font-family:var(--sans);font-size:10px;font-weight:600;letter-spacing:.5px;color:var(--muted);text-transform:uppercase;padding:9px 12px;text-align:left;border-bottom:2px solid var(--border);white-space:nowrap;background:var(--surface2)}
td{padding:11px 12px;border-bottom:1px solid var(--border);vertical-align:middle;color:var(--text)}
tr:last-child td{border-bottom:none}
tr.click:hover td{background:var(--surface3);cursor:pointer}

/* TRACKER TABLE */
.tracker-table{width:100%;border-collapse:collapse;font-size:12px}
.tracker-table th{font-family:var(--sans);font-size:10px;font-weight:600;letter-spacing:.5px;color:var(--muted);text-transform:uppercase;padding:10px 12px;text-align:left;border-bottom:2px solid var(--border);white-space:nowrap;background:var(--surface2);position:sticky;top:0;z-index:2}
.tracker-table th.sortable{cursor:pointer;user-select:none}
.tracker-table th.sortable:hover{color:var(--tm-navy)}
.tracker-table td{padding:11px 12px;border-bottom:1px solid var(--border);vertical-align:middle}
.tracker-table tr:hover td{background:var(--surface3);cursor:pointer}
.tracker-table tr:last-child td{border-bottom:none}
.filter-row{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center}
.filter-input{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:11px;padding:6px 10px;outline:none;transition:border-color .15s;min-width:130px}
.filter-input:focus{border-color:var(--tm-blue)}
.filter-select{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:11px;padding:6px 10px;outline:none;cursor:pointer;min-width:130px}

/* BADGES */
.badge{display:inline-flex;align-items:center;font-family:var(--mono);font-size:9px;font-weight:600;padding:3px 8px;border-radius:4px;white-space:nowrap;letter-spacing:.3px}
.b-amber{background:#FEF3C7;color:#92400E;border:1px solid #FDE68A}
.b-blue{background:#DBEAFE;color:#1E40AF;border:1px solid #BFDBFE}
.b-teal{background:#D1FAE5;color:#065F46;border:1px solid #A7F3D0}
.b-red{background:#FEE2E2;color:#991B1B;border:1px solid #FECACA}
.b-purple{background:#EDE9FE;color:#4C1D95;border:1px solid #DDD6FE}
.b-orange{background:#FFEDD5;color:#9A3412;border:1px solid #FED7AA}
.b-green{background:#D1FAE5;color:#065F46;border:1px solid #A7F3D0}
.b-gray{background:#F3F4F6;color:#6B7280;border:1px solid #E5E7EB}
.urgdot{width:6px;height:6px;border-radius:50%;display:inline-block;margin-right:4px;flex-shrink:0}

/* BUTTONS */
.btn{display:inline-flex;align-items:center;gap:6px;font-family:var(--sans);font-size:11px;font-weight:600;letter-spacing:.3px;padding:7px 14px;border-radius:var(--r);border:1px solid transparent;cursor:pointer;transition:all .15s;white-space:nowrap;text-transform:uppercase}
.btn-primary{background:var(--tm-blue);color:#fff;border-color:var(--tm-blue)}.btn-primary:hover{background:var(--tm-navy)}
.btn-success{background:var(--accent2);color:#fff}.btn-success:hover{background:#145E37}
.btn-danger{background:transparent;color:var(--danger);border-color:var(--danger)}.btn-danger:hover{background:#FEE2E2}
.btn-ghost{background:transparent;color:var(--muted);border-color:var(--border)}.btn-ghost:hover{color:var(--text);border-color:var(--border2);background:var(--surface2)}
.btn-warn{background:transparent;color:var(--warn);border-color:#FDE68A}.btn-warn:hover{background:#FEF3C7}
.btn-cond{background:transparent;color:var(--purple);border-color:#DDD6FE}.btn-cond:hover{background:#EDE9FE}
.btn-sm{padding:4px 10px;font-size:10px}
.btn:disabled{opacity:.4;cursor:not-allowed}

/* MODAL */
.overlay{position:fixed;inset:0;background:rgba(33,51,99,.5);display:flex;align-items:flex-start;justify-content:center;z-index:100;padding:20px;overflow-y:auto;animation:fadeIn .15s}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:12px;width:100%;max-width:860px;margin:auto;animation:slideUp .2s;box-shadow:0 8px 32px rgba(33,51,99,.18)}
.modal-lg{max-width:960px}
.mhdr{display:flex;justify-content:space-between;align-items:flex-start;padding:18px 22px;border-bottom:1px solid var(--border);background:var(--surface2);border-radius:12px 12px 0 0}
.mtitle{font-family:var(--sans);font-size:13px;font-weight:700;letter-spacing:.5px;color:var(--tm-navy)}
.mbody{padding:22px}
.mftr{padding:14px 22px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px;background:var(--surface2);border-radius:0 0 12px 12px}
.mclose{background:none;border:none;color:var(--muted);font-size:20px;cursor:pointer;line-height:1}
.mclose:hover{color:var(--tm-navy)}

@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes slideUp{from{transform:translateY(14px);opacity:0}to{transform:translateY(0);opacity:1}}

/* FORM */
.fg{display:flex;flex-direction:column;gap:5px}
.fg label{font-size:10px;color:var(--tm-navy);letter-spacing:.5px;text-transform:uppercase;font-weight:600;font-family:var(--sans)}
.fg input,.fg select,.fg textarea{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:13px;padding:8px 10px;outline:none;transition:border-color .15s}
.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--tm-blue);box-shadow:0 0 0 3px rgba(35,92,150,.1)}
.fg select option{background:var(--surface)}
.fg textarea{resize:vertical;min-height:65px}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.form-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}
.form-section{font-family:var(--sans);font-size:10px;font-weight:700;letter-spacing:1.5px;color:var(--tm-blue);text-transform:uppercase;margin:18px 0 12px;padding-bottom:6px;border-bottom:2px solid var(--tm-light)}
.checkbox-row{display:flex;align-items:center;gap:9px;padding:7px 0}
.checkbox-row label{font-size:13px;color:var(--text);cursor:pointer;font-family:var(--sans);font-weight:400}
.fg input[type=checkbox]{width:15px;height:15px;accent-color:var(--tm-blue)}

/* ITEMS TABLE */
.items-edit th{font-size:9px;background:var(--surface2)}
.items-edit td{padding:5px 8px}
.items-edit input,.items-edit select{background:var(--surface);border:1px solid var(--border);border-radius:4px;color:var(--text);font-family:var(--mono);font-size:11px;padding:4px 7px;width:100%;outline:none}
.items-edit input:focus,.items-edit select:focus{border-color:var(--tm-blue)}

/* TIMELINE */
.tl{list-style:none}
.tl-item{display:flex;gap:12px;padding-bottom:14px;position:relative}
.tl-item:not(:last-child)::before{content:'';position:absolute;left:10px;top:22px;bottom:0;width:1px;background:var(--border)}
.tl-dot{width:22px;height:22px;border-radius:50%;background:var(--surface2);border:2px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:10px;flex-shrink:0;z-index:1}
.tl-dot.c{border-color:var(--tm-blue);color:var(--tm-blue);background:#DBEAFE}
.tl-dot.a{border-color:var(--accent2);color:var(--accent2);background:#D1FAE5}
.tl-dot.r{border-color:var(--danger);color:var(--danger);background:#FEE2E2}
.tl-dot.u{border-color:var(--warn);color:var(--warn);background:#FEF3C7}
.tl-ev{font-size:13px;font-weight:600;color:var(--tm-navy)}.tl-meta{font-size:11px;color:var(--muted);margin-top:2px}

/* INBOX */
.inbox-header{display:flex;align-items:center;gap:12px;margin-bottom:16px;padding-bottom:12px;border-bottom:2px solid var(--border)}
.inbox-company{font-family:var(--sans);font-size:14px;font-weight:700;color:var(--tm-navy);letter-spacing:.5px}
.req-row{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:16px 18px;margin-bottom:10px;cursor:pointer;transition:all .15s;box-shadow:0 1px 3px rgba(33,51,99,.05)}
.req-row:hover{border-color:var(--tm-blue);box-shadow:0 2px 8px rgba(35,92,150,.12)}
.req-row.unread{border-left:4px solid var(--tm-blue)}
.req-row.devuelto{border-left:4px solid var(--warn)}
.req-title{font-weight:600;font-size:14px;margin-bottom:6px;color:var(--tm-navy)}
.req-meta{display:flex;gap:14px;font-size:11px;color:var(--muted);flex-wrap:wrap;align-items:center}

/* NOTIF */
.notif{position:fixed;bottom:20px;right:20px;background:var(--surface);border:1px solid var(--border);border-left-width:3px;border-radius:var(--r2);padding:12px 16px;font-size:13px;animation:slideUp .2s;z-index:300;max-width:340px;display:flex;align-items:center;gap:10px;box-shadow:0 4px 16px rgba(33,51,99,.15);color:var(--text)}
.n-green{border-left-color:var(--accent2)}.n-red{border-left-color:var(--danger)}.n-amber{border-left-color:var(--warn)}.n-blue{border-left-color:var(--tm-blue)}

/* MISC */
.tag{display:inline-block;font-family:var(--mono);font-size:9px;padding:2px 7px;background:var(--surface2);border:1px solid var(--border);border-radius:4px;color:var(--muted);font-weight:500}
.info-box{background:var(--surface2);border:1px solid var(--border);border-radius:var(--r);padding:12px 14px;font-size:13px;color:var(--text)}
.info-box.accent{border-left:3px solid var(--tm-blue)}
.info-box.warn{border-left:3px solid var(--warn);background:#FFFBEB}
.flex{display:flex}.flex-gap{display:flex;gap:8px;align-items:center}.flex-between{display:flex;justify-content:space-between;align-items:center}
.mt8{margin-top:8px}.mt12{margin-top:12px}.mt16{margin-top:16px}
.mb8{margin-bottom:8px}.mb12{margin-bottom:12px}
.text-mono{font-family:var(--mono)}.text-muted{color:var(--muted)}.text-right{text-align:right}
.empty-state{text-align:center;padding:48px 20px;color:var(--muted);font-size:13px}
.loading{display:flex;align-items:center;justify-content:center;padding:48px;color:var(--muted);gap:10px;font-size:13px}
.spin{animation:spin 1s linear infinite}
@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
.kbar{margin-bottom:10px}.kbar-lbl{display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px;color:var(--text)}.kbar-track{height:6px;background:var(--surface2);border-radius:3px;overflow:hidden;border:1px solid var(--border)}.kbar-fill{height:100%;border-radius:3px}
.tabs-row{display:flex;gap:0;border-bottom:2px solid var(--border);margin-bottom:18px}
.tab{font-family:var(--sans);font-size:11px;font-weight:600;padding:9px 16px;cursor:pointer;color:var(--muted);border-bottom:2px solid transparent;transition:all .12s;text-transform:uppercase;letter-spacing:.5px;margin-bottom:-2px}
.tab.active{color:var(--tm-blue);border-bottom-color:var(--tm-blue)}
.grupo-chip{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:6px;font-family:var(--mono);font-size:12px;font-weight:700;background:#DBEAFE;color:var(--tm-blue);border:1px solid #BFDBFE;flex-shrink:0}
.detail-items{background:var(--surface2);border-radius:var(--r);padding:10px 12px;margin-top:10px;font-size:11px;color:var(--muted);border:1px solid var(--border)}
.detail-item-row{display:flex;gap:8px;padding:2px 0}

/* CRM proveedores */
.prov-tabs{display:flex;border-bottom:2px solid var(--border);margin-bottom:16px}
.prov-tab{font-family:var(--sans);font-size:11px;font-weight:600;padding:9px 16px;cursor:pointer;color:var(--muted);border-bottom:2px solid transparent;margin-bottom:-2px;text-transform:uppercase;letter-spacing:.5px;transition:all .12s}
.prov-tab.active{color:var(--tm-blue);border-bottom-color:var(--tm-blue)}
.historial-row{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:10px 14px;margin-bottom:8px;font-size:12px}
.historial-row:hover{border-color:var(--border2)}
`;

// ─── COMPONENTS ──────────────────────────────────────────────────────────────

function Notif({ msg, onClose }) {
  if (!msg) return null;
  const cls = { success: "n-green", error: "n-red", warn: "n-amber", info: "n-blue" }[msg.type] || "n-blue";
  return <div className={`notif ${cls}`}><span>{msg.text}</span><button onClick={onClose} style={{ marginLeft: "auto", background: "none", border: "none", color: "var(--muted)", cursor: "pointer" }}>✕</button></div>;
}

function UrgBadge({ urgencia }) {
  const color = { Critica: "b-red", Alta: "b-amber", Normal: "b-green" }[urgencia] || "b-gray";
  return <span className={`badge ${color}`}><span className="urgdot" style={{ background: { Critica: "var(--danger)", Alta: "var(--warn)", Normal: "var(--accent2)" }[urgencia] }} />{urgencia}</span>;
}

function TrackerBadge({ status }) {
  const s = TRACKER_STATUS[status] || { label: status, color: "b-gray" };
  return <span className={`badge ${s.color}`}>{s.label}</span>;
}

function FG({ label, hint, children, full }) {
  return <div className="fg" style={full ? { gridColumn: "1/-1" } : {}}>
    {label && <label>{label}</label>}
    {children}
    {hint && <div style={{ fontSize: 10, color: "var(--muted2)", marginTop: 2 }}>{hint}</div>}
  </div>;
}

function Timeline({ historial }) {
  if (!historial?.length) return <div className="text-muted" style={{ fontSize: 11 }}>Sin historial</div>;
  const icon = ev => {
    if (ev.includes("creada") || ev.includes("ingresado")) return { i: "◎", c: "c" };
    if (ev.includes("probado") || ev.includes("OC") || ev.includes("Tracker")) return { i: "✓", c: "a" };
    if (ev.includes("echazado") || ev.includes("evuelto")) return { i: "✗", c: "r" };
    return { i: "·", c: "u" };
  };
  return <ul className="tl">{[...historial].sort((a, b) => new Date(a.fecha) - new Date(b.fecha)).map((h, i) => {
    const { i: ic, c } = icon(h.evento);
    return <li key={i} className="tl-item">
      <div className={`tl-dot ${c}`}>{ic}</div>
      <div><div className="tl-ev">{h.evento}</div><div className="tl-meta">{fmtDate(h.fecha)} · {h.usuario}{h.detalle ? ` · ${h.detalle}` : ""}</div></div>
    </li>;
  })}</ul>;
}

// ─── MODAL: APROBAR CONDICIONAL (editar ítems antes de aprobar) ──────────────
function AprobarCondicionalModal({ req, proveedores, onClose, onSave }) {
  const blank = () => ({ id: `tmp${Date.now()}${Math.random()}`, descripcion: "", cantidad: 1, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "", proyecto: "" });
  const [items, setItems] = useState(
    req.requisicion_items?.length ? req.requisicion_items.map(it => ({ ...it })) : [blank()]
  );
  const [nota, setNota] = useState("");
  const [saving, setSaving] = useState(false);

  const setItem = (i, k, v) => { const its = [...items]; its[i] = { ...its[i], [k]: v }; setItems(its); };

  const handleSave = async () => {
    setSaving(true);
    try {
      // Actualizar ítems
      await api.actualizarItems(req.id, items.filter(it => it.descripcion?.trim()).map(({ id: _id, requisicion_id: _rid, ...rest }) => rest));
      // Aprobar con nota
      const updated = await api.actualizarRequisicion(
        req.id,
        { status: "aprobado_cotizar", revisado_por: USUARIO, fecha_revision: new Date().toISOString(), fecha_aprobacion: new Date().toISOString() },
        `Aprobado con modificaciones${nota ? ` — ${nota}` : ""}`,
        nota || null
      );
      // Fetch fresh para pasar al consolidar
      const fresh = await api.getRequisicion(req.id);
      onSave(fresh);
    } finally { setSaving(false); }
  };

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="mhdr">
          <div>
            <div className="mtitle">APROBAR CON MODIFICACIONES</div>
            <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}</div>
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          <div className="info-box warn mb12" style={{ fontSize: 11 }}>
            Podés editar la descripción, cantidad, unidad o proveedor sugerido de cada ítem antes de aprobar.
            Los cambios quedan registrados en el historial.
          </div>
          <div className="table-wrap">
            <table className="items-edit">
              <thead><tr><th style={{ width: "35%" }}>Descripción</th><th>Cant.</th><th>Unid.</th><th>Proveedor sugerido</th><th></th></tr></thead>
              <tbody>
                {items.map((it, i) => <tr key={it.id || i}>
                  <td><input value={it.descripcion || ""} onChange={e => setItem(i, "descripcion", e.target.value)} /></td>
                  <td><input type="number" value={it.cantidad} onChange={e => setItem(i, "cantidad", e.target.value)} style={{ width: 55 }} /></td>
                  <td><input value={it.unidad || ""} onChange={e => setItem(i, "unidad", e.target.value)} style={{ width: 55 }} /></td>
                  <td>
                    <select value={it.proveedor_sugerido || ""} onChange={e => setItem(i, "proveedor_sugerido", e.target.value)}>
                      <option value="">Sin sugerencia</option>
                      {proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}
                    </select>
                  </td>
                  <td><button className="btn btn-ghost btn-sm" onClick={() => setItems(items.filter((_, j) => j !== i))}>✕</button></td>
                </tr>)}
              </tbody>
            </table>
          </div>
          <button className="btn btn-ghost btn-sm mt8" onClick={() => setItems([...items, blank()])}>+ Agregar ítem</button>
          <div className="mt12">
            <FG label="Nota para el solicitante (opcional)">
              <textarea value={nota} onChange={e => setNota(e.target.value)} placeholder="Ej: Se ajustó la cantidad del ítem 1 según disponibilidad..." />
            </FG>
          </div>
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? "..." : "Aprobar con cambios → Tracker"}</button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: VER REQUISICIÓN (desde inbox) ────────────────────────────────────
function ReqModal({ req: initialReq, proveedores, onClose, onUpdate, onMoverTracker, onRechazar }) {
  const [req, setReq] = useState(initialReq);
  const [tab, setTab] = useState("detalle");
  const [saving, setSaving] = useState(false);
  const [showAprobarCond, setShowAprobarCond] = useState(false);

  const canAprobar = ["pendiente_revision", "en_revision"].includes(req.status);
  const canRechazar = ["pendiente_revision", "en_revision"].includes(req.status);

  const handleAprobar = async () => {
    setSaving(true);
    try {
      const updated = await api.actualizarRequisicion(
        req.id,
        { status: "aprobado_cotizar", revisado_por: USUARIO, fecha_revision: new Date().toISOString(), fecha_aprobacion: new Date().toISOString() },
        "Aprobado para cotizar"
      );
      onUpdate(updated);
      onMoverTracker(updated);
    } finally { setSaving(false); }
  };

  if (showAprobarCond) {
    return <AprobarCondicionalModal
      req={req}
      proveedores={proveedores}
      onClose={() => setShowAprobarCond(false)}
      onSave={(fresh) => {
        setShowAprobarCond(false);
        onUpdate(fresh);
        onMoverTracker(fresh);
      }}
    />;
  }

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="mhdr">
          <div>
            <div className="mtitle">REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}</div>
            <div className="flex-gap mt8">
              <UrgBadge urgencia={req.urgencia} />
              <span className="tag">{req.empresa}</span>
              <span className="tag">{req.base_buque}</span>
              <span className="tag">{req.area}</span>
              {req.veces_devuelto > 0 && <span className="badge b-orange">↩ Devuelta {req.veces_devuelto}x</span>}
            </div>
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody" style={{ paddingBottom: 0 }}>
          <div className="tabs-row">
            {["detalle", "historial"].map(t => <div key={t} className={`tab ${tab === t ? "active" : ""}`} onClick={() => setTab(t)}>{t === "detalle" ? "Detalle" : "Historial"}</div>)}
          </div>

          {tab === "detalle" && <>
            <div className="form-grid mb12">
              <div className="info-box"><div style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", marginBottom: 4 }}>SOLICITANTE</div>{req.solicitado_por}</div>
              <div className="info-box"><div style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", marginBottom: 4 }}>FECHA NECESARIA</div>{fmtDate(req.fecha_necesaria) || "No especificada"}</div>
              <div className="info-box"><div style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", marginBottom: 4 }}>PRESUPUESTO ESTIMADO</div><strong>{req.costo_estimado ? fmt(req.costo_estimado, req.moneda_estimada) : "No especificado"}</strong></div>
              <div className="info-box"><div style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", marginBottom: 4 }}>TIPO</div>{req.tipo_requisicion || "—"}</div>
            </div>
            {req.veces_devuelto > 0 && req.motivo_rechazo_categoria && (
              <div className="info-box warn mb12">
                <strong>Devuelta anteriormente:</strong> {req.motivo_rechazo_categoria}
                {req.motivo_rechazo_texto && <span> — {req.motivo_rechazo_texto}</span>}
              </div>
            )}
            {req.observaciones && <div className="info-box mb12">{req.observaciones}</div>}
            <div style={{ fontFamily: "var(--mono)", fontSize: 9, letterSpacing: 2, color: "var(--muted)", textTransform: "uppercase", marginBottom: 8 }}>Ítems solicitados</div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>#</th><th>Descripción</th><th>Cant.</th><th>Unid.</th><th>Stock</th><th>Proveedor sugerido</th></tr></thead>
                <tbody>
                  {(req.requisicion_items || []).map((it, i) => <tr key={i}>
                    <td className="text-mono text-muted">{it.nro_linea}</td>
                    <td>{it.descripcion}</td>
                    <td className="text-mono">{it.cantidad}</td>
                    <td className="text-muted">{it.unidad}</td>
                    <td className="text-mono">{it.stock_disponible || 0}</td>
                    <td className="text-muted">{it.proveedor_sugerido || "—"}</td>
                  </tr>)}
                </tbody>
              </table>
            </div>
          </>}

          {tab === "historial" && <Timeline historial={req.requisicion_historial} />}
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cerrar</button>
          {canRechazar && <button className="btn btn-danger" onClick={() => { onClose(); onRechazar(req); }} disabled={saving}>Rechazar</button>}
          {canAprobar && <button className="btn btn-cond" onClick={() => setShowAprobarCond(true)} disabled={saving}>Aprobar condicional</button>}
          {canAprobar && <button className="btn btn-primary" onClick={handleAprobar} disabled={saving}>{saving ? "..." : "Aprobar → Tracker"}</button>}
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: CONSOLIDAR EN TRACKER ────────────────────────────────────────────
function ConsolidarModal({ req, onClose, onSave }) {
  const items = req.requisicion_items || [];
  const [asignaciones, setAsignaciones] = useState(items.map(() => "A"));
  const [saving, setSaving] = useState(false);

  const grupos = [...new Set(asignaciones)].sort();

  const handleSave = async () => {
    setSaving(true);
    try {
      const lineas = grupos.map(g => {
        const itemsGrupo = items.filter((_, i) => asignaciones[i] === g);
        return {
          grupo: g,
          descripcion: `Grupo ${g} — REQ-${String(req.nro_solicitud).padStart(4, "0")}`,
          items_detalle: itemsGrupo,
          status: "en_cotizacion",
        };
      });
      await api.crearTrackerLineas(req.id, lineas);
      await api.actualizarRequisicion(req.id, { status: "en_compra" }, `Movido al Tracker (${grupos.length} grupo${grupos.length > 1 ? "s" : ""})`);
      onSave();
    } finally { setSaving(false); }
  };

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="mhdr">
          <div>
            <div className="mtitle">CONSOLIDAR EN TRACKER</div>
            <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}</div>
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          <div className="info-box accent mb12" style={{ fontSize: 11 }}>
            Asigná cada ítem a un grupo (A, B, C...). Ítems del mismo grupo se consolidan en una sola línea del Tracker.
          </div>
          <table className="items-edit">
            <thead><tr><th style={{ width: "50%" }}>Ítem</th><th>Cant.</th><th>Grupo</th></tr></thead>
            <tbody>
              {items.map((it, i) => <tr key={i}>
                <td>{it.descripcion}</td>
                <td className="text-mono">{it.cantidad} {it.unidad}</td>
                <td>
                  <select value={asignaciones[i]} onChange={e => { const a = [...asignaciones]; a[i] = e.target.value; setAsignaciones(a); }} style={{ width: 60 }}>
                    {GRUPOS_OPCIONES.map(g => <option key={g}>{g}</option>)}
                  </select>
                </td>
              </tr>)}
            </tbody>
          </table>
          {grupos.length > 0 && <div className="mt12">
            <div style={{ fontFamily: "var(--mono)", fontSize: 9, letterSpacing: 2, color: "var(--muted)", textTransform: "uppercase", marginBottom: 8 }}>Preview de líneas en el Tracker</div>
            {grupos.map(g => {
              const its = items.filter((_, i) => asignaciones[i] === g);
              return <div key={g} className="info-box mb8">
                <div className="flex-gap mb8"><div className="grupo-chip">{g}</div><strong style={{ fontSize: 12 }}>{its.length} ítem{its.length > 1 ? "s" : ""}</strong></div>
                {its.map((it, i) => <div key={i} style={{ fontSize: 11, color: "var(--muted)", paddingLeft: 8 }}>· {it.descripcion} × {it.cantidad} {it.unidad}</div>)}
              </div>;
            })}
          </div>}
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? "Guardando..." : `Crear ${grupos.length} línea${grupos.length > 1 ? "s" : ""} en Tracker`}</button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: RECHAZAR ─────────────────────────────────────────────────────────
// E2: status → "rechazado_devuelto" para que vuelva al inbox del solicitante,
//     se registra veces_devuelto y motivo para métricas
function RechazarModal({ req, onClose, onSave }) {
  const [categoria, setCategoria] = useState("");
  const [texto, setTexto] = useState("");
  const [devolver, setDevolver] = useState(true); // true = vuelve al inbox, false = rechazado definitivo
  const [saving, setSaving] = useState(false);

  const handleSave = async () => {
    if (!categoria) return alert("Seleccioná una categoría de rechazo");
    setSaving(true);
    try {
      // Si devolver=true → vuelve al inbox como pendiente_revision con veces_devuelto++
      // Si devolver=false → queda como rechazado definitivo en archivo
      const nuevoStatus = devolver ? "pendiente_revision" : "rechazado";
      const evento = devolver ? `Devuelta al solicitante — ${categoria}` : `Rechazado definitivamente — ${categoria}`;
      const updated = await api.actualizarRequisicion(req.id, {
        status: nuevoStatus,
        motivo_rechazo_categoria: categoria,
        motivo_rechazo_texto: texto,
        veces_devuelto: (req.veces_devuelto || 0) + 1,
      }, evento, texto || null);
      onSave(updated, devolver);
    } finally { setSaving(false); }
  };

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 520 }}>
        <div className="mhdr">
          <div className="mtitle">RECHAZAR / DEVOLVER REQUISICIÓN</div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 14 }}>
            REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}
          </div>
          <FG label="Categoría de rechazo *" full>
            <select value={categoria} onChange={e => setCategoria(e.target.value)}>
              <option value="">Seleccionar motivo...</option>
              {CATEGORIAS_RECHAZO.map(c => <option key={c}>{c}</option>)}
            </select>
          </FG>
          <div className="mt12">
            <FG label="Detalle adicional (opcional)">
              <textarea value={texto} onChange={e => setTexto(e.target.value)} placeholder="Explicación adicional para el solicitante..." />
            </FG>
          </div>
          <div className="mt12" style={{ background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: "var(--r)", padding: "12px 14px" }}>
            <div style={{ fontFamily: "var(--sans)", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: "var(--muted)", marginBottom: 10 }}>¿Qué hacemos con esta requisición?</div>
            <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer", marginBottom: 10 }}>
              <input type="radio" name="devolver" checked={devolver} onChange={() => setDevolver(true)} style={{ marginTop: 2, accentColor: "var(--warn)" }} />
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: "var(--warn)" }}>↩ Devolver para corrección</div>
                <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 2 }}>Vuelve al inbox del solicitante marcada como "devuelta". Queda registrado el rechazo para métricas.</div>
              </div>
            </label>
            <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer" }}>
              <input type="radio" name="devolver" checked={!devolver} onChange={() => setDevolver(false)} style={{ marginTop: 2, accentColor: "var(--danger)" }} />
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, color: "var(--danger)" }}>✕ Rechazar definitivamente</div>
                <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 2 }}>Pasa al Archivo / Rechazados. No vuelve al inbox.</div>
              </div>
            </label>
          </div>
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className={`btn ${devolver ? "btn-warn" : "btn-danger"}`} onClick={handleSave} disabled={saving || !categoria}>
            {saving ? "..." : devolver ? "↩ Devolver al solicitante" : "✕ Rechazar definitivamente"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: EDITAR LÍNEA TRACKER ─────────────────────────────────────────────
function TrackerLineaModal({ linea, proveedores, onClose, onSave }) {
  const emptyCotiz = () => ({ proveedor: "", precio: "", moneda: "ARS", plazo: "" });
  const initCotiz = () => {
    const c = linea.cotizaciones || {};
    return [
      c.c1 || emptyCotiz(),
      c.c2 || emptyCotiz(),
      c.c3 || emptyCotiz(),
    ];
  };

  const [form, setForm] = useState({
    descripcion: linea.descripcion || "",
    proveedor_elegido: linea.proveedor_elegido || "",
    motivo_proveedor: linea.motivo_proveedor || "",
    nro_oc: linea.nro_oc || "",
    costo_real: linea.costo_real || "",
    moneda_real: linea.moneda_real || "ARS",
    plazo_pago: linea.plazo_pago || "",
    fecha_entrega_prom: linea.fecha_entrega_prom || "",
    fecha_entrega_real: linea.fecha_entrega_real || "",
    status: linea.status || "en_cotizacion",
    notas: linea.notas || "",
  });
  const [cotiz, setCotiz] = useState(initCotiz());
  const [adjuntos, setAdjuntos] = useState(linea.cotizaciones?.adjuntos || []);
  const [uploading, setUploading] = useState(false);
  const [showDetail, setShowDetail] = useState(false);
  const [saving, setSaving] = useState(false);
  const fileRef = useRef();

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const setCotizField = (idx, k, v) => {
    const next = cotiz.map((c, i) => i === idx ? { ...c, [k]: v } : c);
    setCotiz(next);
    // Si es la cotización elegida (idx=0), auto-completar campos principales
    if (idx === 0) {
      if (k === "proveedor") set("proveedor_elegido", v);
      if (k === "precio") set("costo_real", v);
      if (k === "moneda") set("moneda_real", v);
      if (k === "plazo") set("plazo_pago", v);
    }
  };

  const buildPayload = (overrides = {}) => {
    const f = { ...form, ...overrides };
    return {
      descripcion:        f.descripcion || null,
      proveedor_elegido:  f.proveedor_elegido || null,
      motivo_proveedor:   f.motivo_proveedor || null,
      nro_oc:             f.nro_oc || null,
      costo_real:         f.costo_real !== "" && f.costo_real != null ? parseFloat(f.costo_real) : null,
      moneda_real:        f.moneda_real || "ARS",
      plazo_pago:         f.plazo_pago || null,
      fecha_entrega_prom: f.fecha_entrega_prom || null,
      fecha_entrega_real: f.fecha_entrega_real || null,
      status:             f.status || "en_cotizacion",
      notas:              f.notas || null,
      cotizaciones: {
        c1: cotiz[0],
        c2: cotiz[1],
        c3: cotiz[2],
        adjuntos,
      },
    };
  };

  const handleSave = async (extraCambios = {}) => {
    setSaving(true);
    try {
      const payload = buildPayload(extraCambios);
      const updated = await api.actualizarTrackerLinea(linea.id, payload);
      onSave(updated);
    } catch (e) {
      console.error("Error guardando línea tracker:", e);
      alert("Error al guardar. Revisá la consola.");
    } finally { setSaving(false); }
  };

  const handleConfirmarEntrega = async () => {
    setSaving(true);
    try {
      const payload = buildPayload({
        status: "entregado",
        fecha_entrega_real: form.fecha_entrega_real || new Date().toISOString().split("T")[0],
      });
      const updated = await api.actualizarTrackerLinea(linea.id, payload);
      onSave(updated);
    } catch (e) {
      console.error("Error confirmando entrega:", e);
      alert("Error al confirmar entrega.");
    } finally { setSaving(false); }
  };

  const handleUpload = async (files) => {
    if (!files?.length) return;
    setUploading(true);
    try {
      const nuevos = [];
      for (const file of Array.from(files)) {
        const path = `${linea.id}/${Date.now()}_${file.name}`;
        const { error } = await supabase.storage.from("cotizaciones").upload(path, file, { upsert: true });
        if (error) throw error;
        const { data: urlData } = supabase.storage.from("cotizaciones").getPublicUrl(path);
        nuevos.push({ nombre: file.name, url: urlData.publicUrl, path });
      }
      setAdjuntos(prev => [...prev, ...nuevos]);
    } catch (e) {
      console.error("Error subiendo archivo:", e);
      alert("Error al subir el archivo. Revisá la consola.");
    } finally { setUploading(false); }
  };

  const handleDeleteAdjunto = async (adj) => {
    await supabase.storage.from("cotizaciones").remove([adj.path]);
    setAdjuntos(prev => prev.filter(a => a.path !== adj.path));
  };

  const req = linea.requisiciones;
  const itemsDetalle = linea.items_detalle || [];
  const costoEstimado = req?.costo_estimado;
  const monedaEstimada = req?.moneda_estimada || "ARS";

  const COTIZ_LABELS = ["Cotización elegida", "Cotización 2", "Cotización 3"];
  const COTIZ_STYLES = [
    { border: "2px solid var(--accent2)", background: "#F0FDF4" },
    { border: "1px solid var(--border)", background: "var(--surface2)" },
    { border: "1px solid var(--border)", background: "var(--surface2)" },
  ];

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="mhdr">
          <div>
            <div className="flex-gap">
              <div className="grupo-chip">{linea.grupo}</div>
              <div className="mtitle">{form.descripcion}</div>
            </div>
            {req && <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>
              REQ-{String(req.nro_solicitud).padStart(4, "0")} · {req.empresa} · {req.base_buque}
              {costoEstimado && <span style={{ marginLeft: 8, color: "var(--warn)" }}>Estimado: {fmt(costoEstimado, monedaEstimada)}</span>}
            </div>}
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          {itemsDetalle.length > 0 && <div className="mb12">
            <button className="btn btn-ghost btn-sm" onClick={() => setShowDetail(!showDetail)}>
              {showDetail ? "▲" : "▼"} Ver ítems originales ({itemsDetalle.length})
            </button>
            {showDetail && <div className="detail-items mt8">
              {itemsDetalle.map((it, i) => <div key={i} className="detail-item-row">
                <span className="text-muted">·</span><span>{it.descripcion}</span>
                <span className="text-muted">×{it.cantidad} {it.unidad}</span>
              </div>)}
            </div>}
          </div>}

          <div className="form-section">Estado</div>
          <div className="form-grid">
            <FG label="Status">
              <select value={form.status} onChange={e => set("status", e.target.value)}>
                {Object.entries(TRACKER_STATUS).filter(([k]) => k !== "archivado").map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
              </select>
            </FG>
            <FG label="Descripción consolidada">
              <input value={form.descripcion} onChange={e => set("descripcion", e.target.value)} />
            </FG>
          </div>

          {/* COTIZACIONES — 3 columnas */}
          <div className="form-section">Cotizaciones</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
            {cotiz.map((c, i) => (
              <div key={i} style={{ borderRadius: "var(--r2)", padding: "12px 14px", ...COTIZ_STYLES[i] }}>
                <div style={{ fontFamily: "var(--sans)", fontSize: 10, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", color: i === 0 ? "var(--accent2)" : "var(--muted)", marginBottom: 10 }}>
                  {i === 0 && <span style={{ marginRight: 4 }}>⭐</span>}{COTIZ_LABELS[i]}
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  <FG label="Proveedor">
                    <select value={c.proveedor} onChange={e => setCotizField(i, "proveedor", e.target.value)} style={{ fontSize: 12 }}>
                      <option value="">Seleccionar...</option>
                      {proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}
                    </select>
                  </FG>
                  <FG label="Precio">
                    <input type="number" value={c.precio} onChange={e => setCotizField(i, "precio", e.target.value)} style={{ fontSize: 12 }} />
                  </FG>
                  <FG label="Moneda">
                    <select value={c.moneda} onChange={e => setCotizField(i, "moneda", e.target.value)} style={{ fontSize: 12 }}>
                      <option>ARS</option><option>USD</option>
                    </select>
                  </FG>
                  <FG label="Plazo de pago">
                    <select value={c.plazo} onChange={e => setCotizField(i, "plazo", e.target.value)} style={{ fontSize: 12 }}>
                      <option value="">Seleccionar...</option>
                      {PLAZO_PAGO_OPTIONS.map(p => <option key={p}>{p}</option>)}
                    </select>
                  </FG>
                </div>
              </div>
            ))}
          </div>

          {/* Comparación vs estimado */}
          {costoEstimado && form.costo_real && (
            <div className={`info-box mb12 ${parseFloat(form.costo_real) > costoEstimado ? "warn" : "accent"}`} style={{ fontSize: 11 }}>
              {parseFloat(form.costo_real) > costoEstimado
                ? `⚠ Costo real (${fmt(parseFloat(form.costo_real), form.moneda_real)}) supera el estimado (${fmt(costoEstimado, monedaEstimada)}) en ${fmt(parseFloat(form.costo_real) - costoEstimado, form.moneda_real)}`
                : `✓ Costo real (${fmt(parseFloat(form.costo_real), form.moneda_real)}) dentro del presupuesto estimado (${fmt(costoEstimado, monedaEstimada)})`
              }
            </div>
          )}

          <div className="form-section">Justificación y OC</div>
          <div className="form-grid">
            <FG label="Proveedor elegido">
              <select value={form.proveedor_elegido} onChange={e => set("proveedor_elegido", e.target.value)}>
                <option value="">Seleccionar proveedor...</option>
                {proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}
              </select>
            </FG>
            <FG label="N° OC">
              <input value={form.nro_oc} onChange={e => set("nro_oc", e.target.value)} placeholder="OC-0001" />
            </FG>
          </div>
          <FG label="¿Por qué elegiste este proveedor sobre los demás?">
            <textarea value={form.motivo_proveedor} onChange={e => set("motivo_proveedor", e.target.value)} placeholder="Ej: Mejor precio, plazo de entrega más corto, proveedor habitual con buena performance..." />
          </FG>

          <div className="form-section">Precio y entrega</div>
          <div className="form-grid-3">
            <FG label="Costo real"><input type="number" value={form.costo_real} onChange={e => set("costo_real", e.target.value)} /></FG>
            <FG label="Moneda"><select value={form.moneda_real} onChange={e => set("moneda_real", e.target.value)}><option>ARS</option><option>USD</option></select></FG>
            <FG label="Plazo de pago">
              <select value={form.plazo_pago} onChange={e => set("plazo_pago", e.target.value)}>
                <option value="">Seleccionar...</option>
                {PLAZO_PAGO_OPTIONS.map(p => <option key={p}>{p}</option>)}
              </select>
            </FG>
            <FG label="Entrega prometida"><input type="date" value={form.fecha_entrega_prom} onChange={e => set("fecha_entrega_prom", e.target.value)} /></FG>
            <FG label="Entrega real"><input type="date" value={form.fecha_entrega_real} onChange={e => set("fecha_entrega_real", e.target.value)} /></FG>
          </div>
          <FG label="Notas"><textarea value={form.notas} onChange={e => set("notas", e.target.value)} /></FG>

          {/* ADJUNTOS */}
          <div className="form-section">Presupuestos adjuntos</div>
          <input ref={fileRef} type="file" multiple accept=".pdf,.jpg,.jpeg,.png,.xlsx,.xls,.doc,.docx" style={{ display: "none" }} onChange={e => handleUpload(e.target.files)} />
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 10 }}>
            <button className="btn btn-ghost btn-sm" onClick={() => fileRef.current.click()} disabled={uploading}>
              {uploading ? "⏳ Subiendo..." : "📎 Adjuntar archivo"}
            </button>
            <span style={{ fontSize: 10, color: "var(--muted2)" }}>PDF, imagen o Excel</span>
          </div>
          {adjuntos.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
              {adjuntos.map((adj, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: "var(--r)", padding: "6px 10px" }}>
                  <span style={{ fontSize: 14 }}>📄</span>
                  <a href={adj.url} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: "var(--accent)", textDecoration: "none", flex: 1 }}>{adj.nombre}</a>
                  <button onClick={() => handleDeleteAdjunto(adj)} style={{ background: "none", border: "none", color: "var(--muted2)", cursor: "pointer", fontSize: 14 }}>✕</button>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="btn btn-success btn-sm" onClick={handleConfirmarEntrega} disabled={saving}>✓ Confirmar entrega</button>
          <button className="btn btn-primary" onClick={() => handleSave()} disabled={saving}>{saving ? "Guardando..." : "Guardar"}</button>
        </div>
      </div>
    </div>
  );
}

// ─── PAGE: INBOX ──────────────────────────────────────────────────────────────
function PageInbox({ empresa, notify, onNeedRefresh }) {
  const [reqs, setReqs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [consolidando, setConsolidando] = useState(null);
  const [rechazando, setRechazando] = useState(null);
  const [proveedores, setProveedores] = useState([]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [data, provs] = await Promise.all([
        api.getRequisiciones({ empresa, statuses: ["pendiente_revision", "en_revision"] }),
        api.getProveedores()
      ]);
      setReqs(data);
      setProveedores(provs);
    } finally { setLoading(false); }
  }, [empresa]);

  useEffect(() => { load(); }, [load]);

  const handleUpdate = (updated) => {
    setReqs(prev =>
      prev.map(r => r.id === updated.id ? { ...r, ...updated } : r)
        .filter(r => ["pendiente_revision", "en_revision"].includes(r.status))
    );
  };

  const handleMoverTracker = async (req) => {
    setSelected(null);
    try {
      const fresh = await api.getRequisicion(req.id);
      setConsolidando(fresh);
    } catch {
      setConsolidando(req);
    }
  };

  const handleConsolidado = () => {
    setConsolidando(null);
    notify("Requisición movida al Tracker", "success");
    load();
    onNeedRefresh();
  };

  // E2 FIX: si devolver=true la req vuelve al inbox (re-aparece), si devolver=false sale
  const handleRechazado = (updated, devolver) => {
    setRechazando(null);
    if (devolver) {
      // Vuelve al inbox con badge naranja
      setReqs(prev => prev.map(r => r.id === updated.id ? { ...r, ...updated } : r));
      notify("Requisición devuelta al solicitante para corrección", "warn");
    } else {
      setReqs(prev => prev.filter(r => r.id !== updated.id));
      notify("Requisición rechazada definitivamente", "warn");
    }
    onNeedRefresh();
  };

  return (
    <div>
      <div className="inbox-header">
        <div className="inbox-company">📥 {empresa}</div>
        <span className="ni-badge" style={{ position: "static" }}>{reqs.length}</span>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span> Cargando...</div> :
        reqs.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}>📭</div>Sin requisiciones pendientes</div> :
        reqs.map(r => (
          <div key={r.id} className={`req-row ${r.veces_devuelto > 0 ? "devuelto" : "unread"}`} onClick={() => setSelected(r)}>
            <div className="flex-between mb8">
              <div className="flex-gap">
                <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(r.nro_solicitud).padStart(4, "0")}</span>
                <UrgBadge urgencia={r.urgencia} />
                {r.veces_devuelto > 0 && <span className="badge b-orange">↩ Devuelta {r.veces_devuelto}x</span>}
              </div>
              <span style={{ fontSize: 10, color: "var(--muted)" }}>{fmtDate(r.created_at)}</span>
            </div>
            <div className="req-title">{r.titulo}</div>
            <div className="req-meta">
              <span>{r.base_buque}</span>
              <span>·</span>
              <span>{r.area}{r.subarea ? ` › ${r.subarea}` : ""}</span>
              <span>·</span>
              <span>{r.solicitado_por}</span>
              {r.fecha_necesaria && <><span>·</span><span style={{ color: "var(--warn)" }}>Necesario: {fmtDate(r.fecha_necesaria)}</span></>}
              {r.costo_estimado && <><span>·</span><span>{fmt(r.costo_estimado, r.moneda_estimada)}</span></>}
              {r.veces_devuelto > 0 && r.motivo_rechazo_categoria && <><span>·</span><span style={{ color: "var(--warn)", fontSize: 10 }}>Motivo anterior: {r.motivo_rechazo_categoria}</span></>}
            </div>
          </div>
        ))
      }

      {selected && <ReqModal req={selected} proveedores={proveedores} onClose={() => setSelected(null)} onUpdate={handleUpdate} onMoverTracker={handleMoverTracker} onRechazar={r => { setSelected(null); setRechazando(r); }} />}
      {consolidando && <ConsolidarModal req={consolidando} onClose={() => setConsolidando(null)} onSave={handleConsolidado} />}
      {rechazando && <RechazarModal req={rechazando} onClose={() => setRechazando(null)} onSave={handleRechazado} />}
    </div>
  );
}

// ─── PAGE: TRACKER — E4: vista tabla con filtros ──────────────────────────────
function PageTracker({ statusFilter, notify, onNeedRefresh }) {
  const [lineas, setLineas] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [proveedores, setProveedores] = useState([]);

  // Filtros
  const [filtros, setFiltros] = useState({ empresa: "", proveedor: "", busqueda: "" });
  const [sortCol, setSortCol] = useState("created_at");
  const [sortDir, setSortDir] = useState("desc");

  const statusMap = {
    cotizacion: ["en_cotizacion"],
    oc_emitida: ["oc_emitida"],
    en_transito: ["en_transito"],
  };

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [data, provs] = await Promise.all([
        api.getTrackerLineas({ statuses: statusMap[statusFilter] || ["en_cotizacion"] }),
        api.getProveedores()
      ]);
      setLineas(data);
      setProveedores(provs);
    } finally { setLoading(false); }
  }, [statusFilter]);

  useEffect(() => { load(); }, [load]);

  const handleSave = (updated) => {
    setSelected(null);
    notify("Línea actualizada", "success");
    // E10/E13 FIX: recargar vista y contadores sidebar
    load();
    onNeedRefresh?.();
  };

  const handleSort = (col) => {
    if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  };

  // Filtrado y ordenamiento
  const lineasFiltradas = lineas
    .filter(l => {
      const req = l.requisiciones;
      if (filtros.empresa && req?.empresa !== filtros.empresa) return false;
      if (filtros.proveedor && l.proveedor_elegido !== filtros.proveedor) return false;
      if (filtros.busqueda) {
        const q = filtros.busqueda.toLowerCase();
        const match = (
          l.descripcion?.toLowerCase().includes(q) ||
          req?.nro_solicitud?.toString().includes(q) ||
          req?.base_buque?.toLowerCase().includes(q) ||
          l.proveedor_elegido?.toLowerCase().includes(q) ||
          l.nro_oc?.toLowerCase().includes(q)
        );
        if (!match) return false;
      }
      return true;
    })
    .sort((a, b) => {
      let va, vb;
      const ra = a.requisiciones, rb = b.requisiciones;
      switch (sortCol) {
        case "nro": va = ra?.nro_solicitud || 0; vb = rb?.nro_solicitud || 0; break;
        case "descripcion": va = a.descripcion || ""; vb = b.descripcion || ""; break;
        case "empresa": va = ra?.empresa || ""; vb = rb?.empresa || ""; break;
        case "buque": va = ra?.base_buque || ""; vb = rb?.base_buque || ""; break;
        case "proveedor": va = a.proveedor_elegido || ""; vb = b.proveedor_elegido || ""; break;
        case "costo": va = a.costo_real || 0; vb = b.costo_real || 0; break;
        case "entrega": va = a.fecha_entrega_prom || ""; vb = b.fecha_entrega_prom || ""; break;
        default: va = a.created_at || ""; vb = b.created_at || "";
      }
      const cmp = typeof va === "number" ? va - vb : String(va).localeCompare(String(vb));
      return sortDir === "asc" ? cmp : -cmp;
    });

  const SortIcon = ({ col }) => sortCol === col ? (sortDir === "asc" ? " ▲" : " ▼") : " ·";

  const empresasDisponibles = [...new Set(lineas.map(l => l.requisiciones?.empresa).filter(Boolean))];
  const proveedoresDisponibles = [...new Set(lineas.map(l => l.proveedor_elegido).filter(Boolean))];

  return (
    <div>
      {/* Fila de filtros */}
      <div className="filter-row">
        <input
          className="filter-input"
          placeholder="🔍  Buscar..."
          value={filtros.busqueda}
          onChange={e => setFiltros(f => ({ ...f, busqueda: e.target.value }))}
        />
        <select className="filter-select" value={filtros.empresa} onChange={e => setFiltros(f => ({ ...f, empresa: e.target.value }))}>
          <option value="">Todas las empresas</option>
          {empresasDisponibles.map(e => <option key={e}>{e}</option>)}
        </select>
        <select className="filter-select" value={filtros.proveedor} onChange={e => setFiltros(f => ({ ...f, proveedor: e.target.value }))}>
          <option value="">Todos los proveedores</option>
          {proveedoresDisponibles.map(p => <option key={p}>{p}</option>)}
        </select>
        {(filtros.empresa || filtros.proveedor || filtros.busqueda) && (
          <button className="btn btn-ghost btn-sm" onClick={() => setFiltros({ empresa: "", proveedor: "", busqueda: "" })}>✕ Limpiar</button>
        )}
        <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--muted)", fontFamily: "var(--mono)" }}>
          {lineasFiltradas.length} de {lineas.length}
        </span>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        lineas.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}>📋</div>Sin líneas en este estado</div> :
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div className="table-wrap">
            <table className="tracker-table">
              <thead>
                <tr>
                  <th className="sortable" onClick={() => handleSort("nro")}>REQ<SortIcon col="nro" /></th>
                  <th className="sortable" onClick={() => handleSort("descripcion")}>Descripción<SortIcon col="descripcion" /></th>
                  <th className="sortable" onClick={() => handleSort("empresa")}>Empresa<SortIcon col="empresa" /></th>
                  <th className="sortable" onClick={() => handleSort("buque")}>Base/Buque<SortIcon col="buque" /></th>
                  <th>Urgencia</th>
                  <th className="sortable" onClick={() => handleSort("proveedor")}>Proveedor<SortIcon col="proveedor" /></th>
                  <th>OC</th>
                  <th className="sortable" onClick={() => handleSort("costo")}>Costo<SortIcon col="costo" /></th>
                  <th className="sortable" onClick={() => handleSort("entrega")}>Entrega prom.<SortIcon col="entrega" /></th>
                  <th>Ítems</th>
                </tr>
              </thead>
              <tbody>
                {lineasFiltradas.length === 0 ? (
                  <tr><td colSpan={10} style={{ textAlign: "center", padding: 32, color: "var(--muted)" }}>Sin resultados para los filtros aplicados</td></tr>
                ) : lineasFiltradas.map(l => {
                  const req = l.requisiciones;
                  const items = l.items_detalle || [];
                  return (
                    <tr key={l.id} onClick={() => setSelected(l)}>
                      <td>
                        <div className="flex-gap">
                          <div className="grupo-chip">{l.grupo}</div>
                          {req && <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>
                            REQ-{String(req.nro_solicitud).padStart(4, "0")}
                          </span>}
                        </div>
                      </td>
                      <td>
                        <div style={{ fontWeight: 600, fontSize: 13, color: "var(--tm-navy)", maxWidth: 220 }}>{l.descripcion}</div>
                      </td>
                      <td><span className="tag">{req?.empresa || "—"}</span></td>
                      <td style={{ color: "var(--muted)", fontSize: 12 }}>{req?.base_buque || "—"}</td>
                      <td>{req ? <UrgBadge urgencia={req.urgencia} /> : "—"}</td>
                      <td style={{ fontSize: 12 }}>{l.proveedor_elegido || <span style={{ color: "var(--muted2)" }}>Sin asignar</span>}</td>
                      <td>
                        {l.nro_oc
                          ? <span className="text-mono" style={{ fontSize: 11, color: "var(--accent2)" }}>{l.nro_oc}</span>
                          : <span style={{ color: "var(--muted2)", fontSize: 11 }}>—</span>}
                      </td>
                      <td className="text-mono" style={{ fontSize: 12 }}>
                        {l.costo_real ? fmt(l.costo_real, l.moneda_real) : <span style={{ color: "var(--muted2)" }}>—</span>}
                      </td>
                      <td style={{ fontSize: 12 }}>
                        {l.fecha_entrega_prom
                          ? <span style={{ color: "var(--warn)" }}>{fmtDate(l.fecha_entrega_prom)}</span>
                          : <span style={{ color: "var(--muted2)" }}>—</span>}
                      </td>
                      <td>
                        <span style={{ fontSize: 11, color: "var(--muted)", fontFamily: "var(--mono)" }}>{items.length} ítem{items.length !== 1 ? "s" : ""}</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      }

      {selected && <TrackerLineaModal linea={selected} proveedores={proveedores} onClose={() => setSelected(null)} onSave={handleSave} />}
    </div>
  );
}

// ─── PAGE: TRACKER GENERAL — todas las líneas, todos los estados ─────────────
function PageTrackerGeneral({ notify, onNeedRefresh }) {
  const [lineas, setLineas] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [proveedores, setProveedores] = useState([]);
  const [filtros, setFiltros] = useState({ empresa: "", status: "", proveedor: "", busqueda: "" });
  const [sortCol, setSortCol] = useState("created_at");
  const [sortDir, setSortDir] = useState("desc");

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [data, provs] = await Promise.all([
        api.getTrackerLineas({ statuses: ["en_cotizacion", "oc_emitida", "en_transito", "entregado"] }),
        api.getProveedores()
      ]);
      setLineas(data);
      setProveedores(provs);
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  // E13 FIX: después de guardar recargamos Y actualizamos contadores del sidebar
  const handleSave = (updated) => {
    setSelected(null);
    notify("Línea actualizada", "success");
    load();
    onNeedRefresh(); // esto dispara loadCounts() en el root
  };

  const handleSort = (col) => {
    if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc");
    else { setSortCol(col); setSortDir("asc"); }
  };

  const lineasFiltradas = lineas
    .filter(l => {
      const req = l.requisiciones;
      if (filtros.empresa && req?.empresa !== filtros.empresa) return false;
      if (filtros.status && l.status !== filtros.status) return false;
      if (filtros.proveedor && l.proveedor_elegido !== filtros.proveedor) return false;
      if (filtros.busqueda) {
        const q = filtros.busqueda.toLowerCase();
        if (!(
          l.descripcion?.toLowerCase().includes(q) ||
          req?.nro_solicitud?.toString().includes(q) ||
          req?.base_buque?.toLowerCase().includes(q) ||
          req?.solicitado_por?.toLowerCase().includes(q) ||
          l.proveedor_elegido?.toLowerCase().includes(q) ||
          l.nro_oc?.toLowerCase().includes(q)
        )) return false;
      }
      return true;
    })
    .sort((a, b) => {
      let va, vb;
      const ra = a.requisiciones, rb = b.requisiciones;
      switch (sortCol) {
        case "nro":        va = ra?.nro_solicitud || 0;    vb = rb?.nro_solicitud || 0; break;
        case "descripcion":va = a.descripcion || "";        vb = b.descripcion || ""; break;
        case "empresa":    va = ra?.empresa || "";          vb = rb?.empresa || ""; break;
        case "buque":      va = ra?.base_buque || "";       vb = rb?.base_buque || ""; break;
        case "solicitante":va = ra?.solicitado_por || "";   vb = rb?.solicitado_por || ""; break;
        case "status":     va = a.status || "";             vb = b.status || ""; break;
        case "proveedor":  va = a.proveedor_elegido || "";  vb = b.proveedor_elegido || ""; break;
        case "costo":      va = a.costo_real || 0;          vb = b.costo_real || 0; break;
        case "entrega":    va = a.fecha_entrega_prom || ""; vb = b.fecha_entrega_prom || ""; break;
        case "entrega_real":va = a.fecha_entrega_real || "";vb = b.fecha_entrega_real || ""; break;
        default:           va = a.created_at || "";         vb = b.created_at || "";
      }
      const cmp = typeof va === "number" ? va - vb : String(va).localeCompare(String(vb));
      return sortDir === "asc" ? cmp : -cmp;
    });

  const SortIcon = ({ col }) => sortCol === col ? (sortDir === "asc" ? " ▲" : " ▼") : " ·";
  const empresasDisponibles = [...new Set(lineas.map(l => l.requisiciones?.empresa).filter(Boolean))];
  const proveedoresDisponibles = [...new Set(lineas.map(l => l.proveedor_elegido).filter(Boolean))];

  const totalARS = lineas.filter(l => l.costo_real && (l.moneda_real === "ARS" || !l.moneda_real)).reduce((a, l) => a + l.costo_real, 0);
  const totalUSD = lineas.filter(l => l.costo_real && l.moneda_real === "USD").reduce((a, l) => a + l.costo_real, 0);
  const enCurso = lineas.filter(l => ["en_cotizacion","oc_emitida","en_transito"].includes(l.status)).length;
  const entregadas = lineas.filter(l => l.status === "entregado").length;

  // Exportar a Excel — todas las columnas
  const handleExport = () => {
    const rows = lineasFiltradas.map(l => {
      const req = l.requisiciones;
      const items = (l.items_detalle || []).map(i => `${i.descripcion} x${i.cantidad}`).join(" | ");
      return {
        "REQ": req ? `REQ-${String(req.nro_solicitud).padStart(4,"0")}` : "",
        "Grupo": l.grupo || "",
        "Descripción consolidada": l.descripcion || "",
        "Empresa": req?.empresa || "",
        "Base/Buque": req?.base_buque || "",
        "Área": req?.area || "",
        "Solicitante": req?.solicitado_por || "",
        "Fecha necesaria": req?.fecha_necesaria ? fmtDate(req.fecha_necesaria) : "",
        "Urgencia": req?.urgencia || "",
        "Estado": TRACKER_STATUS[l.status]?.label || l.status || "",
        "Proveedor": l.proveedor_elegido || "",
        "N° OC": l.nro_oc || "",
        "Justificación proveedor": l.motivo_proveedor || "",
        "Costo real": l.costo_real || "",
        "Costo estimado": req?.costo_estimado || "",
        "Moneda": l.moneda_real || "",
        "Plazo de pago": l.plazo_pago || "",
        "Entrega prometida": l.fecha_entrega_prom ? fmtDate(l.fecha_entrega_prom) : "",
        "Entrega real": l.fecha_entrega_real ? fmtDate(l.fecha_entrega_real) : "",
        "Ítems originales": items,
        "Notas": l.notas || "",
        "Fecha creación": l.created_at ? fmtDate(l.created_at) : "",
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tracker");
    XLSX.writeFile(wb, `tracker_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  return (
    <div>
      {/* Stats */}
      <div className="stats" style={{ gridTemplateColumns: "repeat(4,1fr)", marginBottom: 16 }}>
        <div className="stat"><div className="stat-label">Total líneas</div><div className="stat-value va">{lineas.length}</div></div>
        <div className="stat"><div className="stat-label">En curso</div><div className="stat-value vm">{enCurso}</div></div>
        <div className="stat"><div className="stat-label">Entregadas</div><div className="stat-value vg">{entregadas}</div></div>
        <div className="stat">
          <div className="stat-label">Comprometido</div>
          <div style={{ marginTop: 4 }}>
            {totalARS > 0 && <div className="text-mono" style={{ fontSize: 13, fontWeight: 700, color: "var(--accent)" }}>{fmt(totalARS, "ARS")}</div>}
            {totalUSD > 0 && <div className="text-mono" style={{ fontSize: 13, fontWeight: 700, color: "var(--accent2)" }}>{fmt(totalUSD, "USD")}</div>}
            {totalARS === 0 && totalUSD === 0 && <div style={{ color: "var(--muted2)", fontSize: 12 }}>—</div>}
          </div>
        </div>
      </div>

      {/* Filtros + exportar */}
      <div className="filter-row">
        <input className="filter-input" placeholder="🔍  Buscar..." value={filtros.busqueda} onChange={e => setFiltros(f => ({ ...f, busqueda: e.target.value }))} />
        <select className="filter-select" value={filtros.status} onChange={e => setFiltros(f => ({ ...f, status: e.target.value }))}>
          <option value="">Todos los estados</option>
          {Object.entries(TRACKER_STATUS).filter(([k]) => k !== "archivado").map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
        </select>
        <select className="filter-select" value={filtros.empresa} onChange={e => setFiltros(f => ({ ...f, empresa: e.target.value }))}>
          <option value="">Todas las empresas</option>
          {empresasDisponibles.map(e => <option key={e}>{e}</option>)}
        </select>
        <select className="filter-select" value={filtros.proveedor} onChange={e => setFiltros(f => ({ ...f, proveedor: e.target.value }))}>
          <option value="">Todos los proveedores</option>
          {proveedoresDisponibles.map(p => <option key={p}>{p}</option>)}
        </select>
        {(filtros.empresa || filtros.status || filtros.proveedor || filtros.busqueda) && (
          <button className="btn btn-ghost btn-sm" onClick={() => setFiltros({ empresa: "", status: "", proveedor: "", busqueda: "" })}>✕ Limpiar</button>
        )}
        <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--muted)", fontFamily: "var(--mono)" }}>{lineasFiltradas.length} de {lineas.length}</span>
        <button className="btn btn-ghost btn-sm" onClick={handleExport} style={{ marginLeft: 8 }}>↓ Excel</button>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        lineas.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}>📋</div>Sin líneas en el tracker</div> :
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div className="table-wrap">
            <table className="tracker-table">
              <thead>
                <tr>
                  <th className="sortable" onClick={() => handleSort("nro")}>REQ<SortIcon col="nro" /></th>
                  <th className="sortable" onClick={() => handleSort("descripcion")}>Descripción<SortIcon col="descripcion" /></th>
                  <th className="sortable" onClick={() => handleSort("empresa")}>Empresa<SortIcon col="empresa" /></th>
                  <th className="sortable" onClick={() => handleSort("buque")}>Base/Buque<SortIcon col="buque" /></th>
                  <th className="sortable" onClick={() => handleSort("solicitante")}>Solicitante<SortIcon col="solicitante" /></th>
                  <th>Área</th>
                  <th>Urgencia</th>
                  <th>Fecha nec.</th>
                  <th className="sortable" onClick={() => handleSort("status")}>Estado<SortIcon col="status" /></th>
                  <th className="sortable" onClick={() => handleSort("proveedor")}>Proveedor<SortIcon col="proveedor" /></th>
                  <th>OC</th>
                  <th className="sortable" onClick={() => handleSort("costo")}>Costo real<SortIcon col="costo" /></th>
                  <th>Costo est.</th>
                  <th>Moneda</th>
                  <th>Plazo pago</th>
                  <th className="sortable" onClick={() => handleSort("entrega")}>Entrega prom.<SortIcon col="entrega" /></th>
                  <th className="sortable" onClick={() => handleSort("entrega_real")}>Entrega real<SortIcon col="entrega_real" /></th>
                  <th>Ítems</th>
                  <th>Notas</th>
                </tr>
              </thead>
              <tbody>
                {lineasFiltradas.length === 0 ? (
                  <tr><td colSpan={18} style={{ textAlign: "center", padding: 32, color: "var(--muted)" }}>Sin resultados</td></tr>
                ) : lineasFiltradas.map(l => {
                  const req = l.requisiciones;
                  const items = l.items_detalle || [];
                  const entregadaTarde = l.fecha_entrega_prom && l.fecha_entrega_real && new Date(l.fecha_entrega_real) > new Date(l.fecha_entrega_prom);
                  return (
                    <tr key={l.id} onClick={() => setSelected(l)}>
                      <td>
                        <div className="flex-gap">
                          <div className="grupo-chip">{l.grupo}</div>
                          {req && <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(req.nro_solicitud).padStart(4,"0")}</span>}
                        </div>
                      </td>
                      <td><div style={{ fontWeight: 600, fontSize: 12, color: "var(--tm-navy)", maxWidth: 180 }}>{l.descripcion}</div></td>
                      <td><span className="tag">{req?.empresa || "—"}</span></td>
                      <td style={{ color: "var(--muted)", fontSize: 12 }}>{req?.base_buque || "—"}</td>
                      <td style={{ fontSize: 12, color: "var(--muted)" }}>{req?.solicitado_por || "—"}</td>
                      <td style={{ fontSize: 11, color: "var(--muted)" }}>{req?.area || "—"}{req?.subarea ? ` › ${req.subarea}` : ""}</td>
                      <td>{req ? <UrgBadge urgencia={req.urgencia} /> : "—"}</td>
                      <td style={{ fontSize: 11, color: "var(--muted)", whiteSpace: "nowrap" }}>{req?.fecha_necesaria ? fmtDate(req.fecha_necesaria) : "—"}</td>
                      <td><TrackerBadge status={l.status} /></td>
                      <td style={{ fontSize: 12 }}>{l.proveedor_elegido || <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                      <td>{l.nro_oc ? <span className="text-mono" style={{ fontSize: 11, color: "var(--accent2)" }}>{l.nro_oc}</span> : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                      <td className="text-mono" style={{ fontSize: 12, whiteSpace: "nowrap" }}>{l.costo_real != null ? new Intl.NumberFormat("es-AR", { maximumFractionDigits: 0 }).format(l.costo_real) : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                      <td className="text-mono" style={{ fontSize: 12, whiteSpace: "nowrap", color: "var(--muted)" }}>{req?.costo_estimado != null ? new Intl.NumberFormat("es-AR", { maximumFractionDigits: 0 }).format(req.costo_estimado) : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                      <td style={{ fontSize: 11, color: "var(--muted)" }}>{l.moneda_real || "—"}</td>
                      <td style={{ fontSize: 11, color: "var(--muted)" }}>{l.plazo_pago || "—"}</td>
                      <td style={{ fontSize: 12, color: "var(--warn)", whiteSpace: "nowrap" }}>{l.fecha_entrega_prom ? fmtDate(l.fecha_entrega_prom) : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                      <td style={{ fontSize: 12, whiteSpace: "nowrap", color: entregadaTarde ? "var(--danger)" : l.fecha_entrega_real ? "var(--accent2)" : "var(--muted2)" }}>
                        {l.fecha_entrega_real ? fmtDate(l.fecha_entrega_real) : "—"}
                        {entregadaTarde && <span title="Entrega tardía" style={{ marginLeft: 4 }}>⚠</span>}
                      </td>
                      <td>
                        {items.length > 0
                          ? <span style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", cursor: "help" }} title={items.map(i => `${i.descripcion} ×${i.cantidad}`).join("\n")}>{items.length} ítem{items.length !== 1 ? "s" : ""}</span>
                          : <span style={{ color: "var(--muted2)", fontSize: 11 }}>—</span>}
                      </td>
                      <td style={{ fontSize: 11, color: "var(--muted)", maxWidth: 140 }}>
                        {l.notas ? <span title={l.notas}>{l.notas.length > 40 ? l.notas.slice(0,40) + "…" : l.notas}</span> : <span style={{ color: "var(--muted2)" }}>—</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      }
      {selected && <TrackerLineaModal linea={selected} proveedores={proveedores} onClose={() => setSelected(null)} onSave={handleSave} />}
    </div>
  );
}

// ─── PAGE: ARCHIVO ────────────────────────────────────────────────────────────
function PageArchivo({ tipo }) {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [proveedores, setProveedores] = useState([]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      if (tipo === "entregados") {
        const [lineas, provs] = await Promise.all([api.getTrackerLineas({ status: "entregado" }), api.getProveedores()]);
        setData(lineas);
        setProveedores(provs);
      } else {
        const reqs = await api.getRequisiciones({ status: "rechazado" });
        setData(reqs);
      }
    } finally { setLoading(false); }
  }, [tipo]);

  useEffect(() => { load(); }, [load]);

  if (tipo === "rechazados") return (
    <div>
      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        data.length === 0 ? <div className="empty-state">Sin requisiciones rechazadas definitivamente</div> :
        data.map(r => <div key={r.id} className="req-row">
          <div className="flex-between mb8">
            <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(r.nro_solicitud).padStart(4, "0")}</span>
            <span style={{ fontSize: 10, color: "var(--muted)" }}>{fmtDate(r.updated_at)}</span>
          </div>
          <div className="req-title">{r.titulo}</div>
          <div className="req-meta">
            <span>{r.empresa} · {r.base_buque}</span>
            {r.motivo_rechazo_categoria && <><span>·</span><span style={{ color: "var(--danger)" }}>{r.motivo_rechazo_categoria}</span></>}
            {r.motivo_rechazo_texto && <><span>·</span><span>{r.motivo_rechazo_texto}</span></>}
            {r.veces_devuelto > 0 && <><span>·</span><span className="badge b-orange">↩ {r.veces_devuelto}x devuelta antes</span></>}
          </div>
        </div>)
      }
    </div>
  );

  return (
    <div>
      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        data.length === 0 ? <div className="empty-state">Sin entregas registradas</div> :
        data.map(l => {
          const req = l.requisiciones;
          return <div key={l.id} className="tracker-row" style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r2)", padding: "14px 18px", marginBottom: 10, cursor: "pointer", transition: "all .15s" }} onClick={() => setSelected(l)}>
            <div className="flex-gap" style={{ marginBottom: 6 }}>
              <div className="grupo-chip">{l.grupo}</div>
              <span style={{ fontWeight: 600, fontSize: 14, color: "var(--tm-navy)" }}>{l.descripcion}</span>
              <TrackerBadge status="entregado" />
            </div>
            <div style={{ fontSize: 12, color: "var(--muted)" }}>
              {req && <span>{req.empresa} · {req.base_buque}</span>}
              {l.proveedor_elegido && <span> · {l.proveedor_elegido}</span>}
              {l.nro_oc && <span className="text-mono" style={{ color: "var(--accent2)" }}> · {l.nro_oc}</span>}
              {l.fecha_entrega_real && <span> · Entregado: {fmtDate(l.fecha_entrega_real)}</span>}
              {l.costo_real && <span className="text-mono"> · {fmt(l.costo_real, l.moneda_real)}</span>}
            </div>
          </div>;
        })
      }
      {selected && <TrackerLineaModal linea={selected} proveedores={proveedores} onClose={() => setSelected(null)} onSave={() => { setSelected(null); load(); }} />}
    </div>
  );
}

// ─── PAGE: NUEVA REQUISICIÓN ──────────────────────────────────────────────────
function PageNueva({ onSaved, onCancel, notify }) {
  const fileRef = useRef();
  const [excelItems, setExcelItems] = useState(null);
  const [drag, setDrag] = useState(false);
  const [proveedores, setProveedores] = useState([]);

  useEffect(() => { api.getProveedores().then(setProveedores); }, []);

  const handleFile = async (file) => {
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const items = rows.map(r => ({
        descripcion: String(r["Descripcion"] || r["Descripción"] || r["descripcion"] || r["Item"] || Object.values(r)[2] || "").trim(),
        cantidad: parseFloat(r["Cant."] || r["Cantidad"] || 1) || 1,
        unidad: String(r["Unid."] || r["Unidad"] || "Uni"),
        stock_disponible: parseFloat(r["Stock"] || 0) || 0,
        proveedor_sugerido: String(r["Proveedor Sugerido"] || ""),
        proyecto: String(r["Proyecto"] || ""),
      })).filter(i => i.descripcion);
      if (!items.length) { notify("No se encontraron ítems en el Excel", "error"); return; }
      setExcelItems(items);
      notify(`${items.length} ítems importados`, "success");
    } catch { notify("Error al leer el Excel", "error"); }
  };

  const handleSave = async (form, items) => {
    await api.crearRequisicion(form, items);
    notify("Requisición creada", "success");
    onSaved();
  };

  return (
    <div>
      {!excelItems && <div className="card mb12">
        <div className="card-title">Importar desde Excel</div>
        <div style={{ border: "2px dashed var(--border)", borderRadius: "var(--r2)", padding: 32, textAlign: "center", cursor: "pointer", transition: "all .2s" }}
          onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
          onDrop={e => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]); }}
          onClick={() => fileRef.current.click()}>
          <div style={{ fontSize: 28, marginBottom: 8 }}>📥</div>
          <div style={{ fontFamily: "var(--mono)", fontSize: 12, color: "var(--muted)", marginBottom: 4 }}>Arrastrá el Excel o hacé click</div>
          <div style={{ fontSize: 10, color: "var(--muted2)" }}>Columnas: N° · Cant. · Unid. · Descripción · Stock · Proveedor Sugerido</div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={e => handleFile(e.target.files[0])} />
      </div>}

      <div className="card">
        <div className="card-title">{excelItems ? `${excelItems.length} ítems desde Excel — completar datos` : "Cargar manualmente"}</div>
        <ReqForm initial={excelItems ? { requisicion_items: excelItems } : null} proveedores={proveedores} onSave={handleSave} onCancel={onCancel} />
      </div>
    </div>
  );
}

// ─── FORM: REQUISICIÓN ───────────────────────────────────────────────────────
// E5: proveedor_sugerido ahora es <select> con los proveedores de la BD
function ReqForm({ initial, proveedores = [], onSave, onCancel }) {
  const blank = () => ({ id: `tmp${Date.now()}${Math.random()}`, descripcion: "", cantidad: 1, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "", proyecto: "" });
  const [form, setForm] = useState({ titulo: "", empresa: "Parana Logistica", base_buque: "", area: "", subarea: "", detalle_tecnico: "", tipo_requisicion: "", urgencia: "Normal", solicitado_por: "", fecha_necesaria: "", costo_estimado: "", moneda_estimada: "ARS", busco_alternativas: false, observaciones: "", ...(initial || {}) });
  const [items, setItems] = useState(initial?.requisicion_items?.length ? initial.requisicion_items : [blank()]);
  const [saving, setSaving] = useState(false);
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const setItem = (i, k, v) => { const its = [...items]; its[i] = { ...its[i], [k]: v }; setItems(its); };
  const bases = BASES_POR_EMPRESA[form.empresa] || [];
  const areas = AREAS_POR_EMPRESA[form.empresa] || [];
  const subareas = SUBAREA_TECNICA[form.empresa] || [];
  const detalles = DETALLE_TECNICO[form.subarea] || [];

  const handleSubmit = async () => {
    if (!form.titulo || !form.empresa || !form.base_buque || !form.area || !form.solicitado_por) return alert("Completá: Título, Empresa, Base/Buque, Área, Solicitado por");
    if (!items.some(i => i.descripcion.trim())) return alert("Agregá al menos un ítem");
    setSaving(true);
    try {
      const cleanItems = items.filter(i => i.descripcion.trim()).map(({ id: _id, ...rest }) => rest);
      await onSave({ ...form, costo_estimado: form.costo_estimado ? parseFloat(form.costo_estimado) : null }, cleanItems);
    } finally { setSaving(false); }
  };

  return (
    <div>
      <div className="form-section">Datos</div>
      <div className="form-grid">
        <FG label="Título *"><input value={form.titulo} onChange={e => set("titulo", e.target.value)} placeholder="Ej: Compra Bujías Motor Principal" /></FG>
        <FG label="Tipo"><select value={form.tipo_requisicion} onChange={e => set("tipo_requisicion", e.target.value)}><option value="">Seleccionar...</option>{TIPOS_REQUISICION.map(t => <option key={t}>{t}</option>)}</select></FG>
      </div>
      <div className="form-grid-3">
        <FG label="Empresa *"><select value={form.empresa} onChange={e => { set("empresa", e.target.value); set("base_buque", ""); set("area", ""); }}>{EMPRESAS.map(e => <option key={e}>{e}</option>)}</select></FG>
        <FG label="Base / Buque *"><select value={form.base_buque} onChange={e => set("base_buque", e.target.value)}><option value="">Seleccionar...</option>{bases.map(b => <option key={b}>{b}</option>)}</select></FG>
        <FG label="Área *"><select value={form.area} onChange={e => { set("area", e.target.value); set("subarea", ""); }}><option value="">Seleccionar...</option>{areas.map(a => <option key={a}>{a}</option>)}</select></FG>
      </div>
      {form.area === "Tecnica" && <div className="form-grid">
        <FG label="Sub-área"><select value={form.subarea} onChange={e => { set("subarea", e.target.value); set("detalle_tecnico", ""); }}><option value="">Seleccionar...</option>{subareas.map(s => <option key={s}>{s}</option>)}</select></FG>
        {detalles.length > 0 && <FG label="Detalle técnico"><select value={form.detalle_tecnico} onChange={e => set("detalle_tecnico", e.target.value)}><option value="">Seleccionar...</option>{detalles.map(d => <option key={d}>{d}</option>)}</select></FG>}
      </div>}
      <div className="form-grid">
        <FG label="Solicitado por *"><input value={form.solicitado_por} onChange={e => set("solicitado_por", e.target.value)} /></FG>
        <FG label="Fecha necesaria"><input type="date" value={form.fecha_necesaria} onChange={e => set("fecha_necesaria", e.target.value)} /></FG>
      </div>
      <div className="form-grid-3">
        <FG label="Urgencia *"><select value={form.urgencia} onChange={e => set("urgencia", e.target.value)}>{URGENCIA_OPTIONS.map(u => <option key={u}>{u}</option>)}</select></FG>
        <FG label="Costo estimado"><input type="number" value={form.costo_estimado} onChange={e => set("costo_estimado", e.target.value)} /></FG>
        <FG label="Moneda"><select value={form.moneda_estimada} onChange={e => set("moneda_estimada", e.target.value)}><option>ARS</option><option>USD</option></select></FG>
      </div>
      <div className="checkbox-row"><input type="checkbox" id="alt" checked={form.busco_alternativas} onChange={e => set("busco_alternativas", e.target.checked)} /><label htmlFor="alt">Ya busqué alternativas / presupuestos previos</label></div>
      <FG label="Observaciones"><textarea value={form.observaciones} onChange={e => set("observaciones", e.target.value)} /></FG>

      <div className="form-section mt16">Ítems</div>
      <div className="table-wrap">
        <table className="items-edit">
          <thead><tr><th style={{ width: "35%" }}>Descripción *</th><th>Cant.</th><th>Unid.</th><th>Stock</th><th>Proveedor sugerido</th><th>Proyecto</th><th></th></tr></thead>
          <tbody>
            {items.map((it, i) => <tr key={it.id || i}>
              <td><input value={it.descripcion} onChange={e => setItem(i, "descripcion", e.target.value)} /></td>
              <td><input type="number" value={it.cantidad} onChange={e => setItem(i, "cantidad", e.target.value)} style={{ width: 55 }} /></td>
              <td><input value={it.unidad} onChange={e => setItem(i, "unidad", e.target.value)} style={{ width: 50 }} /></td>
              <td><input type="number" value={it.stock_disponible} onChange={e => setItem(i, "stock_disponible", e.target.value)} style={{ width: 55 }} /></td>
              <td>
                {/* E5 FIX: select en lugar de input libre */}
                <select value={it.proveedor_sugerido || ""} onChange={e => setItem(i, "proveedor_sugerido", e.target.value)}>
                  <option value="">Sin sugerencia</option>
                  {proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}
                </select>
              </td>
              <td><input value={it.proyecto || ""} onChange={e => setItem(i, "proyecto", e.target.value)} /></td>
              <td><button className="btn btn-ghost btn-sm" onClick={() => setItems(items.filter((_, j) => j !== i))}>✕</button></td>
            </tr>)}
          </tbody>
        </table>
      </div>
      <button className="btn btn-ghost btn-sm mt8" onClick={() => setItems([...items, blank()])}>+ Agregar ítem</button>

      <div className="flex-gap mt16" style={{ justifyContent: "flex-end", borderTop: "1px solid var(--border)", paddingTop: 14 }}>
        <button className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button className="btn btn-primary" onClick={handleSubmit} disabled={saving}>{saving ? "Guardando..." : "Crear Requisición"}</button>
      </div>
    </div>
  );
}

// ─── PAGE: KPIs ──────────────────────────────────────────────────────────────
function PageKPIs() {
  const [reqs, setReqs] = useState([]);
  const [loading, setLoading] = useState(true);
  useEffect(() => { api.getRequisiciones().then(d => { setReqs(d); setLoading(false); }); }, []);
  if (loading) return <div className="loading"><span className="spin">◌</span></div>;
  const total = reqs.length;
  const urgentes = reqs.filter(r => r.urgencia === "Critica").length;
  const rechazadas = reqs.filter(r => r.status === "rechazado").length;
  const conIV = reqs.filter(r => r.veces_devuelto > 0).length;
  const byEmpresa = {};
  reqs.forEach(r => { byEmpresa[r.empresa] = (byEmpresa[r.empresa] || 0) + 1; });
  const byRechazo = {};
  reqs.filter(r => r.motivo_rechazo_categoria).forEach(r => { byRechazo[r.motivo_rechazo_categoria] = (byRechazo[r.motivo_rechazo_categoria] || 0) + 1; });
  const bySol = {};
  reqs.forEach(r => { if (!bySol[r.solicitado_por]) bySol[r.solicitado_por] = { total: 0, criticas: 0, devueltas: 0 }; bySol[r.solicitado_por].total++; if (r.urgencia === "Critica") bySol[r.solicitado_por].criticas++; if (r.veces_devuelto > 0) bySol[r.solicitado_por].devueltas++; });
  return (
    <div>
      <div className="stats">
        <div className="stat"><div className="stat-label">Total</div><div className="stat-value va">{total}</div></div>
        <div className="stat"><div className="stat-label">% Críticas</div><div className="stat-value vr">{total ? Math.round(urgentes / total * 100) : 0}%</div></div>
        <div className="stat"><div className="stat-label">Devueltas</div><div className="stat-value vm">{conIV}</div></div>
        <div className="stat"><div className="stat-label">Rechazadas</div><div className="stat-value vgr">{rechazadas}</div></div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <div className="card">
          <div className="card-title">Por solicitante</div>
          <table><thead><tr><th>Solicitante</th><th>Total</th><th>Críticas</th><th>Devueltas</th></tr></thead>
            <tbody>{Object.entries(bySol).sort((a, b) => b[1].total - a[1].total).map(([s, d]) => <tr key={s}><td>{s}</td><td className="text-mono">{d.total}</td><td style={{ color: d.criticas > 0 ? "var(--danger)" : "inherit", fontFamily: "var(--mono)" }}>{d.criticas}</td><td style={{ color: d.devueltas > 0 ? "var(--warn)" : "inherit", fontFamily: "var(--mono)" }}>{d.devueltas}</td></tr>)}</tbody>
          </table>
        </div>
        <div className="card">
          <div className="card-title">Motivos de rechazo / devolución</div>
          {Object.keys(byRechazo).length === 0 ? <div className="text-muted" style={{ fontSize: 12 }}>Sin rechazos registrados</div> :
            Object.entries(byRechazo).sort((a, b) => b[1] - a[1]).map(([cat, n]) => <div key={cat} className="kbar">
              <div className="kbar-lbl"><span style={{ color: "var(--muted)" }}>{cat}</span><span className="text-mono">{n}</span></div>
              <div className="kbar-track"><div className="kbar-fill" style={{ width: `${n / Math.max(...Object.values(byRechazo)) * 100}%`, background: "var(--danger)" }} /></div>
            </div>)}
        </div>
        <div className="card">
          <div className="card-title">Por empresa</div>
          {Object.entries(byEmpresa).map(([e, n]) => <div key={e} className="kbar">
            <div className="kbar-lbl"><span style={{ color: "var(--muted)" }}>{e}</span><span className="text-mono">{n}</span></div>
            <div className="kbar-track"><div className="kbar-fill" style={{ width: `${n / total * 100}%`, background: "var(--accent)" }} /></div>
          </div>)}
        </div>
        <div className="card">
          <div className="card-title">Objetivos Lean</div>
          <div className="info-box accent" style={{ fontSize: 11, lineHeight: 1.8 }}>
            Urgencias críticas: objetivo &lt;20% — actual {total ? Math.round(urgentes / total * 100) : 0}%<br />
            Devoluciones por calidad: objetivo &lt;10% — actual {total ? Math.round(conIV / total * 100) : 0}%<br />
            Tiempo solicitud → revisión: objetivo &lt;1 día
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── PAGE: PROVEEDORES — CRM con historial, palabras clave y catálogo ────────
function PageProveedores({ notify }) {
  const [provs, setProvs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(false);
  const [selected, setSelected] = useState(null);
  const [historial, setHistorial] = useState([]);
  const [histLoading, setHistLoading] = useState(false);

  // form nuevo proveedor — incluye palabras_clave e items_catalogo
  const [form, setForm] = useState({
    nombre: "", rubro: "", contacto: "", email: "", telefono: "", notas: "",
    palabras_clave: "", // string separado por comas
  });
  const [itemsCatalogo, setItemsCatalogo] = useState([]); // array de { descripcion, unidad, precio_ref, moneda }
  const blankItem = () => ({ id: `tmp${Date.now()}${Math.random()}`, descripcion: "", unidad: "Uni", precio_ref: "", moneda: "ARS" });

  useEffect(() => { api.getProveedores().then(d => { setProvs(d); setLoading(false); }); }, []);

  const handleSave = async () => {
    if (!form.nombre) return;
    const payload = {
      ...form,
      activo: true,
      palabras_clave: form.palabras_clave || null,
      items_catalogo: itemsCatalogo.filter(i => i.descripcion.trim()).map(({ id: _id, ...rest }) => ({
        ...rest,
        precio_ref: rest.precio_ref ? parseFloat(rest.precio_ref) : null,
      })),
    };
    const nuevo = await api.crearProveedor(payload);
    setProvs(p => [...p, nuevo]);
    setModal(false);
    setForm({ nombre: "", rubro: "", contacto: "", email: "", telefono: "", notas: "", palabras_clave: "" });
    setItemsCatalogo([]);
    notify("Proveedor agregado", "success");
  };

  const handleSelectProveedor = async (prov) => {
    setSelected(prov);
    setHistLoading(true);
    try {
      const lineas = await api.getTrackerLineas({ proveedor: prov.nombre });
      setHistorial(lineas.filter(l => l.costo_real || l.nro_oc));
    } finally { setHistLoading(false); }
  };

  const totalARS = historial.filter(l => l.moneda_real === "ARS" || !l.moneda_real).reduce((a, l) => a + (l.costo_real || 0), 0);
  const totalUSD = historial.filter(l => l.moneda_real === "USD").reduce((a, l) => a + (l.costo_real || 0), 0);

  const setCI = (i, k, v) => { const its = [...itemsCatalogo]; its[i] = { ...its[i], [k]: v }; setItemsCatalogo(its); };

  return (
    <div>
      {selected ? (
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 18 }}>
            <button className="btn btn-ghost btn-sm" onClick={() => setSelected(null)}>← Volver</button>
            <div style={{ fontFamily: "var(--sans)", fontSize: 16, fontWeight: 700, color: "var(--tm-navy)" }}>{selected.nombre}</div>
            {selected.rubro && <span className="tag">{selected.rubro}</span>}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 18 }}>
            <div className="card" style={{ margin: 0 }}>
              <div className="card-title">Datos de contacto</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {selected.contacto && <div><span style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", display: "block", marginBottom: 2 }}>CONTACTO</span>{selected.contacto}</div>}
                {selected.email && <div><span style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", display: "block", marginBottom: 2 }}>EMAIL</span><span style={{ fontFamily: "var(--mono)", fontSize: 12 }}>{selected.email}</span></div>}
                {selected.telefono && <div><span style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", display: "block", marginBottom: 2 }}>TELÉFONO</span><span style={{ fontFamily: "var(--mono)", fontSize: 12 }}>{selected.telefono}</span></div>}
                {selected.palabras_clave && <div>
                  <span style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", display: "block", marginBottom: 6 }}>PALABRAS CLAVE</span>
                  <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                    {selected.palabras_clave.split(",").map(k => k.trim()).filter(Boolean).map((k, i) => (
                      <span key={i} className="tag" style={{ background: "#EEF2FF", borderColor: "#C7D2FE", color: "var(--tm-blue)" }}>{k}</span>
                    ))}
                  </div>
                </div>}
                {selected.notas && <div className="info-box" style={{ fontSize: 12 }}>{selected.notas}</div>}
              </div>
            </div>
            <div className="card" style={{ margin: 0 }}>
              <div className="card-title">Resumen de compras</div>
              {histLoading ? <div className="loading" style={{ padding: 20 }}><span className="spin">◌</span></div> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  <div className="flex-between">
                    <span style={{ fontSize: 12, color: "var(--muted)" }}>Órdenes registradas</span>
                    <span style={{ fontFamily: "var(--mono)", fontSize: 18, fontWeight: 700, color: "var(--tm-blue)" }}>{historial.length}</span>
                  </div>
                  {totalARS > 0 && <div className="flex-between">
                    <span style={{ fontSize: 12, color: "var(--muted)" }}>Total ARS</span>
                    <span style={{ fontFamily: "var(--mono)", fontSize: 14, fontWeight: 600, color: "var(--accent2)" }}>{fmt(totalARS, "ARS")}</span>
                  </div>}
                  {totalUSD > 0 && <div className="flex-between">
                    <span style={{ fontSize: 12, color: "var(--muted)" }}>Total USD</span>
                    <span style={{ fontFamily: "var(--mono)", fontSize: 14, fontWeight: 600, color: "var(--accent2)" }}>{fmt(totalUSD, "USD")}</span>
                  </div>}
                  {historial.length === 0 && <div style={{ fontSize: 12, color: "var(--muted2)" }}>Sin compras registradas aún</div>}
                </div>
              )}
            </div>
          </div>

          {/* Catálogo de ítems del proveedor */}
          {selected.items_catalogo?.length > 0 && (
            <div className="card" style={{ marginBottom: 14 }}>
              <div className="card-title">Catálogo de productos</div>
              <table>
                <thead><tr><th>Descripción</th><th>Unidad</th><th>Precio referencia</th><th>Moneda</th></tr></thead>
                <tbody>
                  {selected.items_catalogo.map((it, i) => (
                    <tr key={i}>
                      <td style={{ fontWeight: 500 }}>{it.descripcion}</td>
                      <td className="text-muted">{it.unidad || "—"}</td>
                      <td className="text-mono">{it.precio_ref ? fmt(it.precio_ref, it.moneda || "ARS") : "—"}</td>
                      <td className="text-muted">{it.moneda || "ARS"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Historial de compras */}
          <div className="card">
            <div className="card-title">Historial de compras</div>
            {histLoading ? <div className="loading"><span className="spin">◌</span></div> :
              historial.length === 0 ? <div className="empty-state" style={{ padding: 24 }}>Sin compras registradas para este proveedor</div> :
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr><th>REQ</th><th>Descripción</th><th>Ítems</th><th>OC</th><th>Precio</th><th>Entrega prom.</th><th>Entrega real</th><th>Estado</th></tr>
                  </thead>
                  <tbody>
                    {historial.map(l => {
                      const req = l.requisiciones;
                      const items = l.items_detalle || [];
                      return (
                        <tr key={l.id}>
                          <td>{req && <div><div className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(req.nro_solicitud).padStart(4, "0")}</div><div style={{ fontSize: 10, color: "var(--muted)" }}>{req.empresa}</div></div>}</td>
                          <td><div style={{ fontWeight: 600, fontSize: 12 }}>{l.descripcion}</div>{req?.base_buque && <div style={{ fontSize: 10, color: "var(--muted)" }}>{req.base_buque}</div>}</td>
                          <td>{items.length > 0 ? <div style={{ fontSize: 10, color: "var(--muted)", maxWidth: 160 }}>{items.slice(0, 2).map((it, i) => <div key={i}>· {it.descripcion} ×{it.cantidad}</div>)}{items.length > 2 && <div style={{ color: "var(--muted2)" }}>+{items.length - 2} más</div>}</div> : "—"}</td>
                          <td>{l.nro_oc ? <span className="text-mono" style={{ fontSize: 11, color: "var(--accent2)" }}>{l.nro_oc}</span> : "—"}</td>
                          <td>{l.costo_real ? <span className="text-mono" style={{ fontWeight: 600 }}>{fmt(l.costo_real, l.moneda_real || "ARS")}</span> : "—"}</td>
                          <td style={{ fontSize: 12, color: "var(--muted)" }}>{fmtDate(l.fecha_entrega_prom)}</td>
                          <td style={{ fontSize: 12, color: l.fecha_entrega_real ? "var(--accent2)" : "var(--muted2)" }}>{l.fecha_entrega_real ? fmtDate(l.fecha_entrega_real) : "Pendiente"}</td>
                          <td><TrackerBadge status={l.status} /></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            }
          </div>
        </div>
      ) : (
        <div className="card">
          <div className="card-title">
            Maestro de proveedores
            <button className="btn btn-primary btn-sm" onClick={() => setModal(true)}>+ Agregar</button>
          </div>
          {loading ? <div className="loading"><span className="spin">◌</span></div> :
            <table>
              <thead><tr><th>Nombre</th><th>Rubro</th><th>Palabras clave</th><th>Contacto</th><th>Email</th><th>Tel.</th><th></th></tr></thead>
              <tbody>
                {provs.map(p => (
                  <tr key={p.id} className="click" onClick={() => handleSelectProveedor(p)}>
                    <td style={{ fontWeight: 600 }}>{p.nombre}</td>
                    <td className="text-muted">{p.rubro || "—"}</td>
                    <td>
                      {p.palabras_clave ? (
                        <div style={{ display: "flex", gap: 3, flexWrap: "wrap" }}>
                          {p.palabras_clave.split(",").map(k => k.trim()).filter(Boolean).slice(0, 4).map((k, i) => (
                            <span key={i} className="tag">{k}</span>
                          ))}
                        </div>
                      ) : <span style={{ color: "var(--muted2)", fontSize: 11 }}>—</span>}
                    </td>
                    <td>{p.contacto || "—"}</td>
                    <td className="text-mono" style={{ fontSize: 11 }}>{p.email || "—"}</td>
                    <td className="text-mono" style={{ fontSize: 11 }}>{p.telefono || "—"}</td>
                    <td><span style={{ fontSize: 11, color: "var(--tm-blue)" }}>Ver →</span></td>
                  </tr>
                ))}
                {!provs.length && <tr><td colSpan={7}><div className="empty-state">Sin proveedores</div></td></tr>}
              </tbody>
            </table>
          }
        </div>
      )}

      {/* Modal nuevo proveedor */}
      {modal && <div className="overlay" onClick={e => e.target === e.currentTarget && setModal(false)}>
        <div className="modal modal-lg">
          <div className="mhdr"><div className="mtitle">Nuevo Proveedor</div><button className="mclose" onClick={() => setModal(false)}>✕</button></div>
          <div className="mbody">
            <div className="form-section">Datos generales</div>
            <div className="form-grid">
              <FG label="Nombre *"><input value={form.nombre} onChange={e => setForm(f => ({ ...f, nombre: e.target.value }))} /></FG>
              <FG label="Rubro"><input value={form.rubro} onChange={e => setForm(f => ({ ...f, rubro: e.target.value }))} placeholder="Ej: Repuestos navales, Catering..." /></FG>
              <FG label="Contacto"><input value={form.contacto} onChange={e => setForm(f => ({ ...f, contacto: e.target.value }))} /></FG>
              <FG label="Email"><input value={form.email} onChange={e => setForm(f => ({ ...f, email: e.target.value }))} /></FG>
              <FG label="Teléfono"><input value={form.telefono} onChange={e => setForm(f => ({ ...f, telefono: e.target.value }))} /></FG>
            </div>
            <FG label="Palabras clave" hint="Separadas por coma. Ej: bujías, filtros, repuestos motor, lubricantes">
              <input
                value={form.palabras_clave}
                onChange={e => setForm(f => ({ ...f, palabras_clave: e.target.value }))}
                placeholder="bujías, filtros, aceite, repuestos..."
              />
            </FG>
            {form.palabras_clave && (
              <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginTop: 6, marginBottom: 12 }}>
                {form.palabras_clave.split(",").map(k => k.trim()).filter(Boolean).map((k, i) => (
                  <span key={i} className="tag" style={{ background: "#EEF2FF", borderColor: "#C7D2FE", color: "var(--tm-blue)" }}>{k}</span>
                ))}
              </div>
            )}
            <FG label="Notas"><textarea value={form.notas} onChange={e => setForm(f => ({ ...f, notas: e.target.value }))} /></FG>

            <div className="form-section">Catálogo de productos / servicios</div>
            <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 12 }}>
              Agregá los productos o servicios que ofrece este proveedor con precio de referencia. Esto permite sugerir proveedores automáticamente.
            </div>
            <div className="table-wrap">
              <table className="items-edit">
                <thead><tr><th style={{ width: "40%" }}>Descripción</th><th>Unidad</th><th>Precio ref.</th><th>Moneda</th><th></th></tr></thead>
                <tbody>
                  {itemsCatalogo.map((it, i) => (
                    <tr key={it.id}>
                      <td><input value={it.descripcion} onChange={e => setCI(i, "descripcion", e.target.value)} placeholder="Ej: Bujía NGK BPR6ES" /></td>
                      <td><input value={it.unidad} onChange={e => setCI(i, "unidad", e.target.value)} style={{ width: 60 }} /></td>
                      <td><input type="number" value={it.precio_ref} onChange={e => setCI(i, "precio_ref", e.target.value)} style={{ width: 90 }} /></td>
                      <td>
                        <select value={it.moneda} onChange={e => setCI(i, "moneda", e.target.value)} style={{ width: 65 }}>
                          <option>ARS</option><option>USD</option>
                        </select>
                      </td>
                      <td><button className="btn btn-ghost btn-sm" onClick={() => setItemsCatalogo(itemsCatalogo.filter((_, j) => j !== i))}>✕</button></td>
                    </tr>
                  ))}
                  {itemsCatalogo.length === 0 && (
                    <tr><td colSpan={5} style={{ textAlign: "center", padding: "12px", color: "var(--muted2)", fontSize: 11 }}>Sin ítems — usá el botón para agregar</td></tr>
                  )}
                </tbody>
              </table>
            </div>
            <button className="btn btn-ghost btn-sm mt8" onClick={() => setItemsCatalogo([...itemsCatalogo, blankItem()])}>+ Agregar ítem al catálogo</button>
          </div>
          <div className="mftr">
            <button className="btn btn-ghost" onClick={() => { setModal(false); setItemsCatalogo([]); }}>Cancelar</button>
            <button className="btn btn-primary" onClick={handleSave}>Guardar proveedor</button>
          </div>
        </div>
      </div>}
    </div>
  );
}

// ─── API VÍVERES ──────────────────────────────────────────────────────────────
const apiViveres = {
  async getCatalogo() {
    const { data, error } = await supabase.from("viveres_catalogo").select("*").eq("activo", true).order("categoria").order("subcategoria").order("descripcion");
    if (error) throw error;
    return data || [];
  },
  async getParametros() {
    const { data, error } = await supabase.from("viveres_parametros_dieta").select("*");
    if (error) throw error;
    return data || [];
  },
  async getPedidos(filtros = {}) {
    let q = supabase.from("viveres_pedidos").select("*, viveres_pedido_items(*)").order("created_at", { ascending: false });
    if (filtros.empresa) q = q.eq("empresa", filtros.empresa);
    if (filtros.base_buque) q = q.eq("base_buque", filtros.base_buque);
    const { data, error } = await q;
    if (error) throw error;
    return data || [];
  },
  async crearPedido(pedido, items) {
    const payload = {
      ...pedido,
      fecha_pedido: pedido.fecha_pedido || null,
      fecha_necesaria: pedido.fecha_necesaria || null,
    };
    const { data: nuevo, error } = await supabase.from("viveres_pedidos").insert([payload]).select().single();
    if (error) throw error;
    if (items?.length) {
      await supabase.from("viveres_pedido_items").insert(items.map(it => ({ ...it, pedido_id: nuevo.id })));
    }
    return nuevo;
  },
  async actualizarPedido(id, cambios) {
    const { data, error } = await supabase.from("viveres_pedidos").update({ ...cambios, updated_at: new Date().toISOString() }).eq("id", id).select().single();
    if (error) throw error;
    return data;
  },
  async actualizarItems(pedidoId, items) {
    await supabase.from("viveres_pedido_items").delete().eq("pedido_id", pedidoId);
    if (items?.length) await supabase.from("viveres_pedido_items").insert(items.map(it => ({ ...it, pedido_id: pedidoId })));
  },
};

// ─── MAPA CATEGORÍA → GRUPO DIETA ────────────────────────────────────────────
const CATEGORIA_A_GRUPO_DIETA = {
  "Carnicería": "Carniceria",
  "Fiambrería": "Fiambreria",
  "Pescadería": "Pescaderia",
  "Verdulería": "Verduras",
  "Lacteos": "Frutas", // sin parámetro específico
};

const TEMP_COLOR = {
  "Seco": { bg: "#FEF9C3", color: "#92400E", border: "#FDE68A", dot: "#EAB308" },
  "Refrigerado": { bg: "#DBEAFE", color: "#1E40AF", border: "#BFDBFE", dot: "#3B82F6" },
  "Congelado": { bg: "#EDE9FE", color: "#4C1D95", border: "#DDD6FE", dot: "#8B5CF6" },
  "Congelados": { bg: "#EDE9FE", color: "#4C1D95", border: "#DDD6FE", dot: "#8B5CF6" },
};

// ─── PAGE: VÍVERES — NUEVO PEDIDO ────────────────────────────────────────────
function PageViveresNuevo({ notify, onSaved, onCancel }) {
  const [step, setStep] = useState(1); // 1: cabecera, 2: ítems
  const [catalogo, setCatalogo] = useState([]);
  const [parametros, setParametros] = useState([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  const [cabecera, setCabecera] = useState({
    empresa: "Parana Logistica",
    base_buque: "",
    pax: 12,
    dias: 15,
    fecha_pedido: new Date().toISOString().split("T")[0],
    fecha_necesaria: "",
    solicitado_por: "",
    observaciones: "",
  });

  // items: { catalogo_id, descripcion, categoria, subcategoria, temperatura, unidad, volumen_peso, stock_actual, cantidad_pedida }
  const [items, setItems] = useState([]);
  const [filtroCateg, setFiltroCateg] = useState("");
  const [filtroTemp, setFiltroTemp] = useState("");
  const [busqueda, setBusqueda] = useState("");

  useEffect(() => {
    Promise.all([apiViveres.getCatalogo(), apiViveres.getParametros()]).then(([cat, par]) => {
      setCatalogo(cat);
      // Inicializar items desde catálogo con stock=0 y cantidad=0
      setItems(cat.map(c => ({
        catalogo_id: c.id,
        descripcion: c.descripcion,
        categoria: c.categoria,
        subcategoria: c.subcategoria || "",
        temperatura: c.temperatura || "",
        unidad: c.unidad || "Unidad",
        volumen_peso: c.volumen_peso || 1,
        stock_actual: 0,
        cantidad_pedida: 0,
      })));
      setParametros(par);
      setLoading(false);
    });
  }, []);

  const [itemsManuales, setItemsManuales] = useState([]);
  const blankManual = () => ({ id: `manual_${Date.now()}_${Math.random()}`, descripcion: "", categoria: "Almacén", temperatura: "Seco", unidad: "Unidad", stock_actual: 0, cantidad_pedida: 0, catalogo_id: null, volumen_peso: 1 });

  const setCab = (k, v) => setCabecera(c => ({ ...c, [k]: v }));
  const setItem = (id, k, v) => setItems(prev => prev.map(it => it.catalogo_id === id ? { ...it, [k]: parseFloat(v) || 0 } : it));

  const paxDias = (cabecera.pax || 0) * (cabecera.dias || 0);

  // Calcular kg/persona/día por grupo de dieta
  const calcDieta = () => {
    const grupos = {};
    items.forEach(it => {
      if (!it.cantidad_pedida && !it.stock_actual) return;
      const total = (it.stock_actual || 0) + (it.cantidad_pedida || 0);
      const kgTotal = total * (it.volumen_peso || 1);
      const kgPaxDia = paxDias > 0 ? kgTotal / paxDias : 0;
      const cat = it.categoria;
      if (!grupos[cat]) grupos[cat] = 0;
      grupos[cat] += kgPaxDia;
    });
    return grupos;
  };

  const dietaActual = calcDieta();

  const getDietaParam = (categoria) => {
    const grupo = CATEGORIA_A_GRUPO_DIETA[categoria] || categoria;
    return parametros.find(p => p.grupo === grupo || p.grupo === categoria);
  };

  const getDietaStatus = (categoria) => {
    const param = getDietaParam(categoria);
    const val = dietaActual[categoria] || 0;
    if (!param) return "gray";
    if (val === 0) return "yellow";
    if (val < param.min) return "red";
    if (val > param.max) return "red";
    return "green";
  };

  // Filtrar ítems para mostrar
  const itemsFiltrados = items.filter(it => {
    if (filtroCateg && it.categoria !== filtroCateg) return false;
    if (filtroTemp && it.temperatura !== filtroTemp) return false;
    if (busqueda && !it.descripcion.toLowerCase().includes(busqueda.toLowerCase())) return false;
    return true;
  });

  const categorias = [...new Set(catalogo.map(c => c.categoria))].sort();
  const temperaturas = [...new Set(catalogo.map(c => c.temperatura).filter(Boolean))];

  // Items con cantidad pedida > 0 (catálogo + manuales)
  const itemsConPedido = [
    ...items.filter(it => it.cantidad_pedida > 0),
    ...itemsManuales.filter(it => it.cantidad_pedida > 0 && it.descripcion.trim()),
  ];

  const handleGuardar = async (status = "borrador") => {
    if (!cabecera.base_buque || !cabecera.solicitado_por) {
      alert("Completá Base/Buque y Solicitado por");
      return;
    }
    setSaving(true);
    try {
      const itemsAGuardar = [
        ...items.filter(it => it.cantidad_pedida > 0 || it.stock_actual > 0),
        ...itemsManuales.filter(it => it.descripcion.trim() && (it.cantidad_pedida > 0 || it.stock_actual > 0)),
      ];
      const pedido = await apiViveres.crearPedido({ ...cabecera, status }, itemsAGuardar);

      // Si se envía, crear requisición automáticamente
      if (status === "enviado") {
        const reqItems = itemsConPedido.map(it => ({
          descripcion: it.descripcion,
          cantidad: it.cantidad_pedida,
          unidad: it.unidad,
          stock_disponible: it.stock_actual,
          proveedor_sugerido: "",
        }));
        const req = await api.crearRequisicion({
          titulo: `Víveres ${cabecera.base_buque} — ${cabecera.pax} PAX × ${cabecera.dias} días`,
          empresa: cabecera.empresa,
          base_buque: cabecera.base_buque,
          area: "Catering",
          tipo_requisicion: "Víveres",
          urgencia: "Normal",
          solicitado_por: cabecera.solicitado_por,
          fecha_necesaria: cabecera.fecha_necesaria || null,
          observaciones: cabecera.observaciones || null,
        }, reqItems);
        await apiViveres.actualizarPedido(pedido.id, { requisicion_id: req.id });
        notify("Pedido enviado al Inbox del comprador", "success");
      } else {
        notify("Borrador guardado", "info");
      }
      onSaved();
    } finally { setSaving(false); }
  };

  if (loading) return <div className="loading"><span className="spin">◌</span> Cargando catálogo...</div>;

  return (
    <div>
      {/* STEP 1: CABECERA */}
      {step === 1 && (
        <div className="card">
          <div className="card-title">Datos del pedido</div>
          <div className="form-grid-3">
            <FG label="Empresa *">
              <select value={cabecera.empresa} onChange={e => setCab("empresa", e.target.value)}>
                {EMPRESAS.map(e => <option key={e}>{e}</option>)}
              </select>
            </FG>
            <FG label="Base / Buque *">
              <select value={cabecera.base_buque} onChange={e => setCab("base_buque", e.target.value)}>
                <option value="">Seleccionar...</option>
                {(BASES_POR_EMPRESA[cabecera.empresa] || []).map(b => <option key={b}>{b}</option>)}
              </select>
            </FG>
            <FG label="Solicitado por *">
              <input value={cabecera.solicitado_por} onChange={e => setCab("solicitado_por", e.target.value)} placeholder="Nombre del cocinero/encargado" />
            </FG>
          </div>
          <div className="form-grid">
            <FG label="PAX (personas a bordo)">
              <input type="number" value={cabecera.pax} onChange={e => setCab("pax", parseInt(e.target.value) || 0)} min={1} />
            </FG>
            <FG label="Días de navegación">
              <input type="number" value={cabecera.dias} onChange={e => setCab("dias", parseInt(e.target.value) || 0)} min={1} />
            </FG>
            <FG label="Fecha del pedido">
              <input type="date" value={cabecera.fecha_pedido} onChange={e => setCab("fecha_pedido", e.target.value)} />
            </FG>
            <FG label="Fecha necesaria">
              <input type="date" value={cabecera.fecha_necesaria} onChange={e => setCab("fecha_necesaria", e.target.value)} />
            </FG>
          </div>
          <FG label="Observaciones">
            <textarea value={cabecera.observaciones} onChange={e => setCab("observaciones", e.target.value)} placeholder="Notas adicionales para el comprador..." />
          </FG>

          {cabecera.pax > 0 && cabecera.dias > 0 && (
            <div className="info-box accent mt12" style={{ fontSize: 12 }}>
              Total: <strong>{cabecera.pax} PAX × {cabecera.dias} días = {paxDias} raciones</strong>
            </div>
          )}

          <div className="flex-gap mt16" style={{ justifyContent: "flex-end", borderTop: "1px solid var(--border)", paddingTop: 14 }}>
            <button className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
            <button className="btn btn-primary" onClick={() => {
              if (!cabecera.base_buque || !cabecera.solicitado_por) { alert("Completá Base/Buque y Solicitado por"); return; }
              setStep(2);
            }}>Continuar → Cargar ítems</button>
          </div>
        </div>
      )}

      {/* STEP 2: ÍTEMS */}
      {step === 2 && (
        <div>
          {/* Header con datos y semáforo de dieta */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
            <div className="card" style={{ margin: 0 }}>
              <div className="card-title">Datos del pedido</div>
              <div style={{ fontSize: 13, color: "var(--tm-navy)", fontWeight: 600 }}>{cabecera.base_buque}</div>
              <div style={{ fontSize: 12, color: "var(--muted)", marginTop: 4 }}>
                {cabecera.empresa} · {cabecera.pax} PAX · {cabecera.dias} días · <strong>{paxDias} raciones</strong>
              </div>
              <div style={{ fontSize: 12, color: "var(--muted)", marginTop: 2 }}>Por: {cabecera.solicitado_por}</div>
              <button className="btn btn-ghost btn-sm mt8" onClick={() => setStep(1)}>← Editar datos</button>
            </div>
            <div className="card" style={{ margin: 0 }}>
              <div className="card-title">Control de dieta — kg/persona/día</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
                {parametros.map(p => {
                  const catKey = Object.entries(CATEGORIA_A_GRUPO_DIETA).find(([, v]) => v === p.grupo)?.[0] || p.grupo;
                  const val = dietaActual[catKey] || 0;
                  const status = val === 0 ? "yellow" : val < p.min ? "red" : val > p.max ? "red" : "green";
                  const colors = { green: { bg: "#D1FAE5", color: "#065F46" }, red: { bg: "#FEE2E2", color: "#991B1B" }, yellow: { bg: "#FEF9C3", color: "#92400E" } };
                  return (
                    <div key={p.grupo} style={{ background: colors[status].bg, borderRadius: "var(--r)", padding: "5px 8px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 10, color: colors[status].color, fontWeight: 600 }}>{p.grupo}</span>
                      <span style={{ fontFamily: "var(--mono)", fontSize: 11, color: colors[status].color }}>
                        {val.toFixed(2)} / {p.max} {p.unidad_medida}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Pestañas por categoría */}
          <div style={{ display: "flex", gap: 0, borderBottom: "2px solid var(--border)", marginBottom: 14, overflowX: "auto" }}>
            <div className={`tab ${filtroCateg === "" ? "active" : ""}`} onClick={() => setFiltroCateg("")} style={{ whiteSpace: "nowrap" }}>Todos</div>
            {categorias.map(cat => {
              const cantCat = items.filter(it => it.categoria === cat && it.cantidad_pedida > 0).length;
              return (
                <div key={cat} className={`tab ${filtroCateg === cat ? "active" : ""}`} onClick={() => setFiltroCateg(cat)} style={{ whiteSpace: "nowrap" }}>
                  {cat}
                  {cantCat > 0 && <span style={{ marginLeft: 6, background: "var(--accent2)", color: "#fff", fontSize: 9, fontWeight: 700, padding: "1px 5px", borderRadius: 8, fontFamily: "var(--mono)" }}>{cantCat}</span>}
                </div>
              );
            })}
            <div
              className={`tab ${filtroCateg === "__manual__" ? "active" : ""}`}
              onClick={() => setFiltroCateg("__manual__")}
              style={{ whiteSpace: "nowrap", color: filtroCateg === "__manual__" ? "var(--purple)" : undefined, borderBottomColor: filtroCateg === "__manual__" ? "var(--purple)" : undefined }}
            >
              ✏️ Ingreso manual
              {itemsManuales.filter(it => it.cantidad_pedida > 0 && it.descripcion.trim()).length > 0 && (
                <span style={{ marginLeft: 6, background: "var(--purple)", color: "#fff", fontSize: 9, fontWeight: 700, padding: "1px 5px", borderRadius: 8, fontFamily: "var(--mono)" }}>
                  {itemsManuales.filter(it => it.cantidad_pedida > 0 && it.descripcion.trim()).length}
                </span>
              )}
            </div>
          </div>

          {/* Vista Ingreso Manual */}
          {filtroCateg === "__manual__" ? (
            <div>
              <div className="info-box accent mb12" style={{ fontSize: 11 }}>
                Agregá ítems que no están en el catálogo. Completá todos los campos y luego se incluirán en el pedido al enviar.
              </div>
              <div className="card" style={{ padding: 0, overflow: "hidden", marginBottom: 80 }}>
                <div className="table-wrap">
                  <table className="tracker-table">
                    <thead>
                      <tr>
                        <th>Temperatura</th>
                        <th>Categoría</th>
                        <th style={{ width: "30%" }}>Descripción</th>
                        <th>Unidad</th>
                        <th style={{ width: 100 }}>Stock actual</th>
                        <th style={{ width: 120 }}>Cantidad pedida</th>
                        <th></th>
                      </tr>
                    </thead>
                    <tbody>
                      {itemsManuales.map((it, i) => (
                        <tr key={it.id}>
                          <td>
                            <select value={it.temperatura} onChange={e => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], temperatura: e.target.value }; setItemsManuales(arr); }}
                              style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r)", fontFamily: "var(--sans)", fontSize: 11, padding: "4px 8px", outline: "none" }}>
                              <option>Seco</option><option>Refrigerado</option><option>Congelado</option>
                            </select>
                          </td>
                          <td>
                            <select value={it.categoria} onChange={e => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], categoria: e.target.value }; setItemsManuales(arr); }}
                              style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r)", fontFamily: "var(--sans)", fontSize: 11, padding: "4px 8px", outline: "none", minWidth: 100 }}>
                              {categorias.map(c => <option key={c}>{c}</option>)}
                              <option>Otro</option>
                            </select>
                          </td>
                          <td>
                            <input value={it.descripcion} onChange={e => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], descripcion: e.target.value }; setItemsManuales(arr); }}
                              placeholder="Descripción del ítem..."
                              style={{ width: "100%", background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r)", fontFamily: "var(--sans)", fontSize: 12, padding: "4px 8px", outline: "none" }} />
                          </td>
                          <td>
                            <input value={it.unidad} onChange={e => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], unidad: e.target.value }; setItemsManuales(arr); }}
                              placeholder="Uni, Kg..."
                              style={{ width: 70, background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12, padding: "4px 8px", outline: "none" }} />
                          </td>
                          <td>
                            <input type="number" min={0} value={it.stock_actual || ""} placeholder="0"
                              onChange={e => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], stock_actual: parseFloat(e.target.value) || 0 }; setItemsManuales(arr); }}
                              style={{ width: 80, background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12, padding: "4px 8px", outline: "none", textAlign: "right" }} />
                          </td>
                          <td>
                            <input type="number" min={0} value={it.cantidad_pedida || ""} placeholder="0"
                              onChange={e => { const arr = [...itemsManuales]; arr[i] = { ...arr[i], cantidad_pedida: parseFloat(e.target.value) || 0 }; setItemsManuales(arr); }}
                              style={{ width: 90, background: it.cantidad_pedida > 0 ? "#DCFCE7" : "var(--surface)", border: `1px solid ${it.cantidad_pedida > 0 ? "#86EFAC" : "var(--border)"}`, borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12, padding: "4px 8px", outline: "none", textAlign: "right", fontWeight: it.cantidad_pedida > 0 ? 700 : 400 }} />
                          </td>
                          <td>
                            <button className="btn btn-ghost btn-sm" onClick={() => setItemsManuales(itemsManuales.filter((_, j) => j !== i))}>✕</button>
                          </td>
                        </tr>
                      ))}
                      {itemsManuales.length === 0 && (
                        <tr><td colSpan={7} style={{ textAlign: "center", padding: 24, color: "var(--muted2)", fontSize: 12 }}>Sin ítems manuales — usá el botón para agregar</td></tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
              <button className="btn btn-ghost btn-sm" onClick={() => setItemsManuales([...itemsManuales, blankManual()])}>+ Agregar ítem manual</button>
            </div>
          ) : (
            <div>
          {/* Filtro búsqueda y temperatura */}
          <div className="filter-row" style={{ marginBottom: 12 }}>
            <input className="filter-input" placeholder="🔍 Buscar ítem..." value={busqueda} onChange={e => setBusqueda(e.target.value)} style={{ minWidth: 200 }} />
            <select className="filter-select" value={filtroTemp} onChange={e => setFiltroTemp(e.target.value)}>
              <option value="">Todas las temperaturas</option>
              {temperaturas.map(t => <option key={t}>{t}</option>)}
            </select>
            {(filtroTemp || busqueda) && (
              <button className="btn btn-ghost btn-sm" onClick={() => { setFiltroTemp(""); setBusqueda(""); }}>✕ Limpiar</button>
            )}
            <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>
              {itemsFiltrados.length} ítems visibles
            </span>
          </div>

          {/* Tabla de ítems */}
          <div className="card" style={{ padding: 0, overflow: "hidden", marginBottom: 80 }}>
            <div className="table-wrap">
              <table className="tracker-table">
                <thead>
                  <tr>
                    <th>Temp.</th>
                    <th>Categoría</th>
                    <th>Descripción</th>
                    <th>Unidad</th>
                    <th style={{ width: 100 }}>Stock actual</th>
                    <th style={{ width: 120 }}>Cantidad pedida</th>
                    <th>Total</th>
                    <th>kg/PAX/día</th>
                  </tr>
                </thead>
                <tbody>
                  {itemsFiltrados.length === 0 ? (
                    <tr><td colSpan={8} style={{ textAlign: "center", padding: 32, color: "var(--muted)" }}>Sin ítems</td></tr>
                  ) : itemsFiltrados.map(it => {
                    const tc = TEMP_COLOR[it.temperatura] || { bg: "#F3F4F6", color: "#6B7280", border: "#E5E7EB", dot: "#9CA3AF" };
                    const total = (it.stock_actual || 0) + (it.cantidad_pedida || 0);
                    const kgPaxDia = paxDias > 0 ? (total * (it.volumen_peso || 1)) / paxDias : 0;
                    return (
                      <tr key={it.catalogo_id} style={{ background: it.cantidad_pedida > 0 ? "#F0FDF4" : "inherit" }}>
                        <td>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 10, fontWeight: 600, color: tc.color, background: tc.bg, border: `1px solid ${tc.border}`, borderRadius: 4, padding: "2px 6px" }}>
                            <span style={{ width: 6, height: 6, borderRadius: "50%", background: tc.dot, display: "inline-block" }} />
                            {it.temperatura}
                          </span>
                        </td>
                        <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.categoria}</td>
                        <td style={{ fontWeight: it.cantidad_pedida > 0 ? 600 : 400, fontSize: 12 }}>{it.descripcion}</td>
                        <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.unidad}</td>
                        <td>
                          <input type="number" min={0} value={it.stock_actual || ""} placeholder="0"
                            onChange={e => setItem(it.catalogo_id, "stock_actual", e.target.value)}
                            style={{ width: 80, background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12, padding: "4px 8px", outline: "none", textAlign: "right" }} />
                        </td>
                        <td>
                          <input type="number" min={0} value={it.cantidad_pedida || ""} placeholder="0"
                            onChange={e => setItem(it.catalogo_id, "cantidad_pedida", e.target.value)}
                            style={{ width: 90, background: it.cantidad_pedida > 0 ? "#DCFCE7" : "var(--surface)", border: `1px solid ${it.cantidad_pedida > 0 ? "#86EFAC" : "var(--border)"}`, borderRadius: "var(--r)", fontFamily: "var(--mono)", fontSize: 12, padding: "4px 8px", outline: "none", textAlign: "right", fontWeight: it.cantidad_pedida > 0 ? 700 : 400 }} />
                        </td>
                        <td className="text-mono" style={{ fontSize: 11, color: total > 0 ? "var(--tm-navy)" : "var(--muted2)" }}>{total > 0 ? total : "—"}</td>
                        <td className="text-mono" style={{ fontSize: 11, color: kgPaxDia > 0 ? "var(--accent)" : "var(--muted2)" }}>{kgPaxDia > 0 ? kgPaxDia.toFixed(3) : "—"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
            </div>
          )}

          {/* CARRITO FLOTANTE */}
          <div style={{ position: "fixed", bottom: 0, left: 235, right: 0, background: "var(--tm-navy)", borderTop: "2px solid rgba(255,255,255,.15)", padding: "12px 28px", display: "flex", alignItems: "center", gap: 16, zIndex: 50, boxShadow: "0 -4px 16px rgba(33,51,99,.2)" }}>
            <div style={{ flex: 1 }}>
              {itemsConPedido.length === 0 ? (
                <span style={{ fontSize: 12, color: "rgba(255,255,255,.5)" }}>Sin ítems seleccionados — completá cantidades pedidas</span>
              ) : (
                <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
                  {/* Resumen por categoría */}
                  {[...new Set(itemsConPedido.map(it => it.categoria))].map(cat => {
                    const count = itemsConPedido.filter(it => it.categoria === cat).length;
                    return (
                      <div key={cat} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                        <span style={{ fontSize: 10, color: "rgba(255,255,255,.5)" }}>{cat}</span>
                        <span style={{ fontSize: 11, fontFamily: "var(--mono)", fontWeight: 700, color: "#fff", background: "rgba(255,255,255,.15)", borderRadius: 4, padding: "1px 6px" }}>{count}</span>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "var(--mono)" }}>
                {itemsConPedido.length} ítem{itemsConPedido.length !== 1 ? "s" : ""}
              </div>
              <button className="btn btn-ghost" onClick={() => setStep(1)} style={{ color: "rgba(255,255,255,.7)", borderColor: "rgba(255,255,255,.2)" }}>← Volver</button>
              <button className="btn" onClick={() => handleGuardar("borrador")} disabled={saving} style={{ background: "rgba(255,255,255,.15)", color: "#fff", borderColor: "rgba(255,255,255,.2)" }}>Guardar borrador</button>
              <button className="btn btn-success" onClick={() => handleGuardar("enviado")} disabled={saving || itemsConPedido.length === 0}>
                {saving ? "Enviando..." : `✓ Enviar al comprador`}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── PAGE: VÍVERES — HISTORIAL ────────────────────────────────────────────────
function PageViveresHistorial({ notify, onNuevo }) {
  const [pedidos, setPedidos] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);

  useEffect(() => {
    apiViveres.getPedidos().then(d => { setPedidos(d); setLoading(false); });
  }, []);

  const STATUS_PEDIDO = {
    borrador: { label: "Borrador", color: "b-gray" },
    enviado:  { label: "Enviado al comprador", color: "b-blue" },
    aprobado: { label: "Aprobado", color: "b-green" },
    rechazado:{ label: "Rechazado", color: "b-red" },
  };

  return (
    <div>
      <div className="flex-between mb12">
        <div style={{ fontSize: 13, color: "var(--muted)" }}>{pedidos.length} pedidos registrados</div>
        <button className="btn btn-primary btn-sm" onClick={onNuevo}>+ Nuevo pedido</button>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        pedidos.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}>🚢</div>Sin pedidos de víveres aún</div> :
        pedidos.map(p => {
          const s = STATUS_PEDIDO[p.status] || { label: p.status, color: "b-gray" };
          const itemsCantidad = (p.viveres_pedido_items || []).filter(it => it.cantidad_pedida > 0).length;
          return (
            <div key={p.id} className="req-row" onClick={() => setSelected(p)}>
              <div className="flex-between mb8">
                <div className="flex-gap">
                  <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>{fmtDate(p.fecha_pedido)}</span>
                  <span className={`badge ${s.color}`}>{s.label}</span>
                </div>
                <span style={{ fontSize: 10, color: "var(--muted)" }}>{p.empresa}</span>
              </div>
              <div className="req-title">{p.base_buque} — {p.pax} PAX × {p.dias} días</div>
              <div className="req-meta">
                <span>{p.solicitado_por}</span>
                <span>·</span>
                <span>{itemsCantidad} ítems pedidos</span>
                {p.fecha_necesaria && <><span>·</span><span style={{ color: "var(--warn)" }}>Necesario: {fmtDate(p.fecha_necesaria)}</span></>}
              </div>
            </div>
          );
        })
      }

      {selected && (
        <div className="overlay" onClick={e => e.target === e.currentTarget && setSelected(null)}>
          <div className="modal modal-lg">
            <div className="mhdr">
              <div>
                <div className="mtitle">Pedido — {selected.base_buque}</div>
                <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>
                  {selected.empresa} · {selected.pax} PAX · {selected.dias} días · Por: {selected.solicitado_por}
                </div>
              </div>
              <button className="mclose" onClick={() => setSelected(null)}>✕</button>
            </div>
            <div className="mbody">
              <div className="table-wrap">
                <table>
                  <thead><tr><th>Categoría</th><th>Descripción</th><th>Unidad</th><th>Stock</th><th>Pedido</th></tr></thead>
                  <tbody>
                    {(selected.viveres_pedido_items || []).filter(it => it.cantidad_pedida > 0 || it.stock_actual > 0).map((it, i) => (
                      <tr key={i}>
                        <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.categoria}</td>
                        <td style={{ fontWeight: 500 }}>{it.descripcion}</td>
                        <td style={{ fontSize: 11, color: "var(--muted)" }}>{it.unidad}</td>
                        <td className="text-mono">{it.stock_actual || 0}</td>
                        <td className="text-mono" style={{ fontWeight: 700, color: "var(--accent2)" }}>{it.cantidad_pedida || 0}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div className="mftr">
              <button className="btn btn-ghost" onClick={() => setSelected(null)}>Cerrar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── PAGE: VÍVERES — CATÁLOGO ────────────────────────────────────────────────
function PageViveresCatalogo({ notify }) {
  const [catalogo, setCatalogo] = useState([]);
  const [loading, setLoading] = useState(true);
  const [busqueda, setBusqueda] = useState("");
  const [filtroCateg, setFiltroCateg] = useState("");
  const [modal, setModal] = useState(false);
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState({ codigo: "", categoria: "Almacén", subcategoria: "", temperatura: "Seco", descripcion: "", unidad: "Unidad", volumen_peso: "" });

  useEffect(() => {
    apiViveres.getCatalogo().then(d => { setCatalogo(d); setLoading(false); });
  }, []);

  const categorias = [...new Set(catalogo.map(c => c.categoria))].sort();
  const filtrado = catalogo.filter(c => {
    if (filtroCateg && c.categoria !== filtroCateg) return false;
    if (busqueda && !c.descripcion.toLowerCase().includes(busqueda.toLowerCase()) && !c.codigo?.toLowerCase().includes(busqueda.toLowerCase())) return false;
    return true;
  });

  const handleGuardar = async () => {
    if (!form.descripcion.trim()) return alert("La descripción es obligatoria");
    setSaving(true);
    try {
      const { data, error } = await supabase.from("viveres_catalogo").insert([{
        ...form,
        volumen_peso: form.volumen_peso ? parseFloat(form.volumen_peso) : null,
        activo: true,
      }]).select().single();
      if (error) throw error;
      setCatalogo(prev => [...prev, data]);
      setModal(false);
      setForm({ codigo: "", categoria: "Almacén", subcategoria: "", temperatura: "Seco", descripcion: "", unidad: "Unidad", volumen_peso: "" });
      notify("Ítem agregado al catálogo", "success");
    } catch (e) {
      console.error(e);
      alert("Error al guardar: " + e.message);
    } finally { setSaving(false); }
  };

  const setF = (k, v) => setForm(f => ({ ...f, [k]: v }));

  return (
    <div>
      <div className="filter-row mb12">
        <input className="filter-input" placeholder="🔍 Buscar por descripción o código..." value={busqueda} onChange={e => setBusqueda(e.target.value)} style={{ minWidth: 250 }} />
        <select className="filter-select" value={filtroCateg} onChange={e => setFiltroCateg(e.target.value)}>
          <option value="">Todas las categorías</option>
          {categorias.map(c => <option key={c}>{c}</option>)}
        </select>
        {(busqueda || filtroCateg) && <button className="btn btn-ghost btn-sm" onClick={() => { setBusqueda(""); setFiltroCateg(""); }}>✕ Limpiar</button>}
        <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>{filtrado.length} de {catalogo.length}</span>
        <button className="btn btn-primary btn-sm" onClick={() => setModal(true)}>+ Agregar ítem</button>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div className="table-wrap">
            <table>
              <thead>
                <tr><th>Código</th><th>Categoría</th><th>Subcategoría</th><th>Temp.</th><th>Descripción</th><th>Unidad</th></tr>
              </thead>
              <tbody>
                {filtrado.map(c => {
                  const tc = TEMP_COLOR[c.temperatura] || { bg: "#F3F4F6", color: "#6B7280", border: "#E5E7EB", dot: "#9CA3AF" };
                  return (
                    <tr key={c.id}>
                      <td className="text-mono" style={{ fontSize: 10, color: "var(--muted)" }}>{c.codigo || "—"}</td>
                      <td style={{ fontSize: 12, color: "var(--muted)" }}>{c.categoria}</td>
                      <td style={{ fontSize: 11, color: "var(--muted2)" }}>{c.subcategoria || "—"}</td>
                      <td>
                        <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 10, fontWeight: 600, color: tc.color, background: tc.bg, border: `1px solid ${tc.border}`, borderRadius: 4, padding: "2px 6px" }}>
                          <span style={{ width: 5, height: 5, borderRadius: "50%", background: tc.dot, display: "inline-block" }} />
                          {c.temperatura}
                        </span>
                      </td>
                      <td style={{ fontWeight: 500, fontSize: 12 }}>{c.descripcion}</td>
                      <td style={{ fontSize: 11, color: "var(--muted)" }}>{c.unidad || "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      }

      {/* Modal nuevo ítem catálogo */}
      {modal && (
        <div className="overlay" onClick={e => e.target === e.currentTarget && setModal(false)}>
          <div className="modal" style={{ maxWidth: 560 }}>
            <div className="mhdr">
              <div className="mtitle">Agregar ítem al catálogo</div>
              <button className="mclose" onClick={() => setModal(false)}>✕</button>
            </div>
            <div className="mbody">
              <div className="form-grid">
                <FG label="Código (opcional)">
                  <input value={form.codigo} onChange={e => setF("codigo", e.target.value)} placeholder="Ej: NAV001" />
                </FG>
                <FG label="Temperatura *">
                  <select value={form.temperatura} onChange={e => setF("temperatura", e.target.value)}>
                    <option>Seco</option>
                    <option>Refrigerado</option>
                    <option>Congelado</option>
                  </select>
                </FG>
                <FG label="Categoría *">
                  <select value={form.categoria} onChange={e => setF("categoria", e.target.value)}>
                    {categorias.map(c => <option key={c}>{c}</option>)}
                    <option>Otro</option>
                  </select>
                </FG>
                <FG label="Subcategoría">
                  <input value={form.subcategoria} onChange={e => setF("subcategoria", e.target.value)} placeholder="Ej: Aceite/Aceto/Vinagre" />
                </FG>
              </div>
              <FG label="Descripción *" full>
                <input value={form.descripcion} onChange={e => setF("descripcion", e.target.value)} placeholder="Nombre completo del producto" />
              </FG>
              <div className="form-grid mt12">
                <FG label="Unidad">
                  <select value={form.unidad} onChange={e => setF("unidad", e.target.value)}>
                    {["Unidad", "Kg", "Litros", "Caja", "Bolsa", "Atados", "Cajon", "Ristra"].map(u => <option key={u}>{u}</option>)}
                  </select>
                </FG>
                <FG label="Volumen/Peso por unidad" hint="En kg o litros, para el cálculo de dieta">
                  <input type="number" value={form.volumen_peso} onChange={e => setF("volumen_peso", e.target.value)} placeholder="Ej: 1, 0.5, 2.5" />
                </FG>
              </div>
            </div>
            <div className="mftr">
              <button className="btn btn-ghost" onClick={() => setModal(false)}>Cancelar</button>
              <button className="btn btn-primary" onClick={handleGuardar} disabled={saving}>{saving ? "Guardando..." : "Agregar al catálogo"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
export default function App() {
  const [page, setPage] = useState("inbox-parana");
  const [notif, setNotif] = useState(null);
  const [counts, setCounts] = useState({ parana: 0, cleansea: 0, terramare: 0, cotizacion: 0, oc: 0, transito: 0 });
  const [refreshKey, setRefreshKey] = useState(0);

  const notify = useCallback((text, type = "info") => {
    setNotif({ text, type });
    setTimeout(() => setNotif(null), 4000);
  }, []);

  const loadCounts = useCallback(async () => {
    try {
      const [reqs, tracker] = await Promise.all([
        api.getRequisiciones({ statuses: ["pendiente_revision", "en_revision"] }),
        api.getTrackerLineas({ statuses: ["en_cotizacion", "oc_emitida", "en_transito"] })
      ]);
      setCounts({
        parana: reqs.filter(r => r.empresa === "Parana Logistica").length,
        cleansea: reqs.filter(r => r.empresa === "Clean Sea").length,
        terramare: reqs.filter(r => r.empresa === "Terra Mare").length,
        cotizacion: tracker.filter(l => l.status === "en_cotizacion").length,
        oc: tracker.filter(l => l.status === "oc_emitida").length,
        transito: tracker.filter(l => l.status === "en_transito").length,
      });
    } catch (e) { console.error(e); }
  }, []);

  useEffect(() => { loadCounts(); }, [loadCounts, refreshKey]);

  const pageTitles = {
    "inbox-parana": "INBOX — PARANA LOGÍSTICA",
    "inbox-cleansea": "INBOX — CLEAN SEA",
    "inbox-terramare": "INBOX — TERRA MARE",
    "tracker": "TRACKER — TODAS LAS COMPRAS",
    "tracker-cotizacion": "TRACKER — EN COTIZACIÓN",
    "tracker-oc": "TRACKER — OC EMITIDA",
    "tracker-transito": "TRACKER — EN TRÁNSITO",
    "archivo-entregados": "ARCHIVO — ENTREGADOS",
    "archivo-rechazados": "ARCHIVO — RECHAZADOS",
    "viveres-nuevo": "VÍVERES — NUEVO PEDIDO",
    "viveres-historial": "VÍVERES — HISTORIAL",
    "viveres-catalogo": "VÍVERES — CATÁLOGO",
    "nueva": "NUEVA REQUISICIÓN",
    "kpis": "KPIs & REPORTES",
    "proveedores": "PROVEEDORES",
  };

  const NI = ({ id, icon, label, badge, badgeColor, sub }) => (
    <div className={`ni ${sub ? "sub" : ""} ${page === id ? "active" : ""}`} onClick={() => setPage(id)}>
      <span className="ni-icon">{icon}</span>
      <span>{label}</span>
      {badge > 0 && <span className={`ni-badge ${badgeColor || ""}`}>{badge}</span>}
    </div>
  );

  return (
    <>
      <style>{CSS}</style>
      <div className="app">
        <nav className="sidebar">
          <div className="sidebar-header">
            <div className="sidebar-logo-wrap">
              <img src="/logo-terramare.png" alt="Terra Mare" className="sidebar-logo-img" />
              <div className="sidebar-logo-text">
                <div className="sidebar-logo-main">Compras</div>
                <div className="sidebar-logo-sub">oil & gas support services</div>
              </div>
            </div>
          </div>

          <div className="nav-section">Inbox</div>
          <NI id="inbox-parana" icon="📥" label="Parana Logística" sub />
          <NI id="inbox-cleansea" icon="📥" label="Clean Sea" sub />
          <NI id="inbox-terramare" icon="📥" label="Terra Mare" sub />

          <div className="nav-section">Tracker Compras</div>
          <NI id="tracker" icon="📊" label="Tracker" sub={false} />
          <NI id="tracker-cotizacion" icon="🔍" label="En cotización" sub />
          <NI id="tracker-oc" icon="📄" label="OC Emitida" sub />
          <NI id="tracker-transito" icon="🚚" label="En tránsito" sub />

          <div className="nav-section">Archivo</div>
          <NI id="archivo-entregados" icon="✓" label="Entregados" sub />
          <NI id="archivo-rechazados" icon="✗" label="Rechazados" sub />

          <div className="nav-section">Víveres</div>
          <NI id="viveres-nuevo" icon="🛒" label="Nuevo Pedido" sub />
          <NI id="viveres-historial" icon="📋" label="Historial" sub />
          <NI id="viveres-catalogo" icon="📦" label="Catálogo" sub />

          <div className="nav-section">Gestión</div>
          <NI id="nueva" icon="✚" label="Nueva Requisición" />
          <NI id="kpis" icon="📊" label="KPIs & Reportes" />
          <NI id="proveedores" icon="🏭" label="Proveedores" />

          <div style={{ flex: 1 }} />
          <div style={{ padding: "14px 18px", borderTop: "1px solid rgba(255,255,255,.1)" }}>
            <div style={{ fontSize: 9, color: "rgba(255,255,255,.3)", fontFamily: "var(--mono)", letterSpacing: 1 }}>
              SISTEMA DE COMPRAS v2.1
            </div>
          </div>
        </nav>

        <div className="main">
          <div className="topbar">
            <div className="topbar-title">{pageTitles[page] || page}</div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 28, height: 28, borderRadius: "50%", background: "#DBEAFE", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: "var(--tm-blue)", fontWeight: 700 }}>
                {USUARIO.charAt(0)}
              </div>
              <span style={{ fontSize: 12, color: "var(--muted)", fontWeight: 500 }}>{USUARIO}</span>
            </div>
          </div>
          <div className="content">
            {page === "inbox-parana" && <PageInbox empresa="Parana Logistica" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "inbox-cleansea" && <PageInbox empresa="Clean Sea" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "inbox-terramare" && <PageInbox empresa="Terra Mare" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "tracker" && <PageTrackerGeneral key={`tg-${refreshKey}`} notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "tracker-cotizacion" && <PageTracker key={`tc-${refreshKey}`} statusFilter="cotizacion" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "tracker-oc" && <PageTracker key={`to-${refreshKey}`} statusFilter="oc_emitida" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "tracker-transito" && <PageTracker key={`tt-${refreshKey}`} statusFilter="en_transito" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "archivo-entregados" && <PageArchivo tipo="entregados" />}
            {page === "archivo-rechazados" && <PageArchivo tipo="rechazados" />}
            {page === "nueva" && <PageNueva onSaved={() => { setPage("inbox-parana"); loadCounts(); }} onCancel={() => setPage("inbox-parana")} notify={notify} />}
            {page === "viveres-nuevo" && <PageViveresNuevo notify={notify} onSaved={() => setPage("viveres-historial")} onCancel={() => setPage("viveres-historial")} />}
            {page === "viveres-historial" && <PageViveresHistorial notify={notify} onNuevo={() => setPage("viveres-nuevo")} />}
            {page === "viveres-catalogo" && <PageViveresCatalogo notify={notify} />}
            {page === "kpis" && <PageKPIs />}
            {page === "proveedores" && <PageProveedores notify={notify} />}
          </div>
        </div>
      </div>
      <Notif msg={notif} onClose={() => setNotif(null)} />
    </>
  );
}
