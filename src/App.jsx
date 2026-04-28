import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  EMPRESAS, BASES_POR_EMPRESA, AREAS_POR_EMPRESA, SUBAREA_TECNICA,
  DETALLE_TECNICO, TIPOS_REQUISICION, URGENCIA_OPTIONS, PLAZO_PAGO_OPTIONS,
  STATUS_LABELS, CATEGORIAS_RECHAZO
} from "./lib/catalogos";
import { supabase } from "./lib/supabase";

const USUARIO = "Comprador";
const PORTAL_URL = "https://erp-portal-fawn.vercel.app";

const TRACKER_STATUS = {
  en_cotizacion: { label: "En cotización", color: "b-amber" },
  oc_emitida:    { label: "OC Emitida",    color: "b-blue" },
  en_transito:   { label: "En tránsito",   color: "b-purple" },
  entregado:     { label: "Entregado",     color: "b-green" },
  archivado:     { label: "Archivado",     color: "b-gray" },
};

const GRUPOS_OPCIONES = ["A", "B", "C", "D", "E"];

const fmt = (n, cur = "ARS") =>
  n != null ? new Intl.NumberFormat("es-AR", { style: "currency", currency: cur, maximumFractionDigits: 0 }).format(n) : "—";
const fmtDate = d => d ? new Date(d).toLocaleDateString("es-AR") : "—";

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
    const { data: nueva, error } = await supabase.from("requisiciones").insert([{ ...req, status: "pendiente_aprobacion" }]).select().single();
    if (error) throw error;
    if (items?.length) {
      await supabase.from("requisicion_items").insert(items.map((it, i) => ({ ...it, requisicion_id: nueva.id, nro_linea: i + 1 })));
    }
    await supabase.from("requisicion_historial").insert([{ requisicion_id: nueva.id, evento: "Requisición creada", usuario: USUARIO, status_nuevo: "pendiente_aprobacion" }]);
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
    let q = supabase.from("tracker_lineas").select("*, requisiciones(nro_solicitud, titulo, empresa, base_buque, area, subarea, urgencia, solicitado_por, fecha_necesaria, tipo_requisicion, observaciones)").order("created_at", { ascending: false });
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
  async subirAdjunto(file, path) {
    const { error } = await supabase.storage.from("cotizaciones").upload(path, file, { upsert: true });
    if (error) throw error;
    const { data } = supabase.storage.from("cotizaciones").getPublicUrl(path);
    return data.publicUrl;
  },
};

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --navy:#213363;--blue:#235C96;--mid:#6381A7;--light:#A5B5CC;
  --bg:#F0F4F8;--surface:#FFF;--surface2:#F5F7FA;--surface3:#EAF0F6;
  --border:#D6E0ED;--border2:#B0C4D8;
  --text:#213363;--muted:#6381A7;--muted2:#8FA3BC;
  --accent:#235C96;--accent2:#1E7E4A;--warn:#B07D0A;--danger:#C0392B;
  --purple:#6B4FA0;--teal:#1A7A6E;--orange:#C05621;
  --mono:'DM Mono',monospace;--sans:'Montserrat',sans-serif;--r:6px;--r2:10px;
}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.5;min-height:100vh}
.app{display:flex;min-height:100vh}
.sidebar{width:235px;min-width:235px;background:var(--navy);display:flex;flex-direction:column;box-shadow:2px 0 8px rgba(33,51,99,.15)}
.sidebar-header{border-bottom:1px solid rgba(255,255,255,.1)}
.sidebar-logo-wrap{padding:20px 18px 16px;display:flex;align-items:center;gap:12px}
.sidebar-logo-img{width:36px;height:36px;object-fit:cover;border-radius:50%;border:2px solid rgba(255,255,255,.2)}
.sidebar-logo-main{font-size:13px;font-weight:700;color:#fff;letter-spacing:2px;text-transform:uppercase}
.sidebar-logo-sub{font-size:9px;color:rgba(255,255,255,.5);margin-top:2px;letter-spacing:.5px}
.nav-section{padding:12px 18px 4px;font-family:var(--mono);font-size:9px;letter-spacing:2px;color:rgba(255,255,255,.35);text-transform:uppercase}
.ni{display:flex;align-items:center;gap:9px;padding:7px 18px;font-size:12px;font-weight:500;cursor:pointer;color:rgba(255,255,255,.6);border-left:3px solid transparent;transition:all .12s;user-select:none}
.ni:hover{color:#fff;background:rgba(255,255,255,.06)}
.ni.active{color:#fff;border-left-color:var(--light);background:rgba(255,255,255,.1);font-weight:600}
.ni.sub{padding-left:32px;font-size:11px;font-weight:400}
.ni.sub.active{font-weight:600}
.ni.back{color:rgba(255,255,255,.4);font-size:11px;border-top:1px solid rgba(255,255,255,.08);margin-top:4px}
.ni.back:hover{color:rgba(255,255,255,.8)}
.ni-icon{font-size:13px;width:16px;text-align:center;flex-shrink:0}
.ni-badge{margin-left:auto;background:var(--danger);color:#fff;font-family:var(--mono);font-size:9px;font-weight:700;padding:1px 6px;border-radius:10px;min-width:18px;text-align:center}
.ni-badge.amber{background:var(--warn)}
.ni-badge.gray{background:rgba(255,255,255,.2);color:rgba(255,255,255,.7)}
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;min-width:0}
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:13px 28px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 1px 3px rgba(33,51,99,.06)}
.topbar-title{font-size:12px;font-weight:600;letter-spacing:1px;color:var(--navy);text-transform:uppercase}
.content{flex:1;overflow-y:auto;padding:24px 28px;background:var(--bg)}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:20px;margin-bottom:16px;box-shadow:0 1px 4px rgba(33,51,99,.06)}
.card-title{font-size:10px;font-weight:600;letter-spacing:1.5px;color:var(--muted);text-transform:uppercase;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}
.stat{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:16px 18px}
.stat-label{font-size:10px;color:var(--muted);font-weight:600;letter-spacing:.5px;margin-bottom:6px;text-transform:uppercase}
.stat-value{font-family:var(--mono);font-size:28px;font-weight:600}
.va{color:var(--blue)}.vg{color:var(--accent2)}.vr{color:var(--danger)}.vp{color:var(--purple)}.vm{color:var(--warn)}.vgr{color:var(--muted)}
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
th{font-size:10px;font-weight:600;letter-spacing:.5px;color:var(--muted);text-transform:uppercase;padding:9px 12px;text-align:left;border-bottom:2px solid var(--border);white-space:nowrap;background:var(--surface2)}
td{padding:11px 12px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr.click:hover td{background:var(--surface3);cursor:pointer}
.tracker-table{width:100%;border-collapse:collapse;font-size:12px}
.tracker-table th{font-size:10px;font-weight:600;letter-spacing:.5px;color:var(--muted);text-transform:uppercase;padding:10px 12px;text-align:left;border-bottom:2px solid var(--border);white-space:nowrap;background:var(--surface2);position:sticky;top:0;z-index:2}
.tracker-table th.sortable{cursor:pointer;user-select:none}.tracker-table th.sortable:hover{color:var(--navy)}
.tracker-table td{padding:11px 12px;border-bottom:1px solid var(--border);vertical-align:middle}
.tracker-table tr:hover td{background:var(--surface3);cursor:pointer}
.tracker-table tr:last-child td{border-bottom:none}
.filter-row{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center}
.filter-input{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:11px;padding:6px 10px;outline:none;min-width:130px}
.filter-select{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:11px;padding:6px 10px;outline:none;cursor:pointer;min-width:130px}
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
.btn{display:inline-flex;align-items:center;gap:6px;font-family:var(--sans);font-size:11px;font-weight:600;letter-spacing:.3px;padding:7px 14px;border-radius:var(--r);border:1px solid transparent;cursor:pointer;transition:all .15s;white-space:nowrap;text-transform:uppercase}
.btn-primary{background:var(--blue);color:#fff}.btn-primary:hover{background:var(--navy)}
.btn-success{background:var(--accent2);color:#fff}.btn-success:hover{background:#145E37}
.btn-danger{background:transparent;color:var(--danger);border-color:var(--danger)}.btn-danger:hover{background:#FEE2E2}
.btn-ghost{background:transparent;color:var(--muted);border-color:var(--border)}.btn-ghost:hover{color:var(--text);background:var(--surface2)}
.btn-warn{background:transparent;color:var(--warn);border-color:#FDE68A}.btn-warn:hover{background:#FEF3C7}
.btn-cond{background:transparent;color:var(--purple);border-color:#DDD6FE}.btn-cond:hover{background:#EDE9FE}
.btn-confirm{background:transparent;color:var(--orange);border-color:#FED7AA}.btn-confirm:hover{background:#FFEDD5}
.btn-sm{padding:4px 10px;font-size:10px}
.btn:disabled{opacity:.4;cursor:not-allowed}
.overlay{position:fixed;inset:0;background:rgba(33,51,99,.5);display:flex;align-items:flex-start;justify-content:center;z-index:100;padding:20px;overflow-y:auto;animation:fadeIn .15s}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:12px;width:100%;max-width:860px;margin:auto;animation:slideUp .2s;box-shadow:0 8px 32px rgba(33,51,99,.18)}
.modal-lg{max-width:960px}
.mhdr{display:flex;justify-content:space-between;align-items:flex-start;padding:18px 22px;border-bottom:1px solid var(--border);background:var(--surface2);border-radius:12px 12px 0 0}
.mtitle{font-size:13px;font-weight:700;letter-spacing:.5px;color:var(--navy)}
.mbody{padding:22px}
.mftr{padding:14px 22px;border-top:1px solid var(--border);display:flex;justify-content:flex-end;gap:8px;background:var(--surface2);border-radius:0 0 12px 12px}
.mclose{background:none;border:none;color:var(--muted);font-size:20px;cursor:pointer}
.mclose:hover{color:var(--navy)}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes slideUp{from{transform:translateY(14px);opacity:0}to{transform:translateY(0);opacity:1}}
.fg{display:flex;flex-direction:column;gap:5px}
.fg label{font-size:10px;color:var(--navy);letter-spacing:.5px;text-transform:uppercase;font-weight:600}
.fg input,.fg select,.fg textarea{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);color:var(--text);font-family:var(--sans);font-size:13px;padding:8px 10px;outline:none;transition:border-color .15s}
.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--blue)}
.fg textarea{resize:vertical;min-height:65px}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
.form-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}
.form-section{font-size:10px;font-weight:700;letter-spacing:1.5px;color:var(--blue);text-transform:uppercase;margin:18px 0 12px;padding-bottom:6px;border-bottom:2px solid var(--light)}
.items-edit th{font-size:9px;background:var(--surface2)}
.items-edit td{padding:5px 8px}
.items-edit input,.items-edit select{background:var(--surface);border:1px solid var(--border);border-radius:4px;color:var(--text);font-family:var(--mono);font-size:11px;padding:4px 7px;width:100%;outline:none}
.items-edit input:focus,.items-edit select:focus{border-color:var(--blue)}
.tl{list-style:none}
.tl-item{display:flex;gap:12px;padding-bottom:14px;position:relative}
.tl-item:not(:last-child)::before{content:'';position:absolute;left:10px;top:22px;bottom:0;width:1px;background:var(--border)}
.tl-dot{width:22px;height:22px;border-radius:50%;background:var(--surface2);border:2px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:10px;flex-shrink:0;z-index:1}
.tl-dot.c{border-color:var(--blue);color:var(--blue);background:#DBEAFE}
.tl-dot.a{border-color:var(--accent2);color:var(--accent2);background:#D1FAE5}
.tl-dot.r{border-color:var(--danger);color:var(--danger);background:#FEE2E2}
.tl-dot.u{border-color:var(--warn);color:var(--warn);background:#FEF3C7}
.tl-ev{font-size:13px;font-weight:600;color:var(--navy)}.tl-meta{font-size:11px;color:var(--muted);margin-top:2px}
.req-row{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:16px 18px;margin-bottom:10px;cursor:pointer;transition:all .15s}
.req-row:hover{border-color:var(--blue);box-shadow:0 2px 8px rgba(35,92,150,.12)}
.req-row.unread{border-left:4px solid var(--blue)}
.req-row.devuelto{border-left:4px solid var(--warn)}
.req-row.pendiente-confirmacion{border-left:4px solid var(--orange)}
.req-title{font-weight:600;font-size:14px;margin-bottom:6px;color:var(--navy)}
.req-meta{display:flex;gap:14px;font-size:11px;color:var(--muted);flex-wrap:wrap;align-items:center}
.notif{position:fixed;bottom:20px;right:20px;background:var(--surface);border:1px solid var(--border);border-left-width:3px;border-radius:var(--r2);padding:12px 16px;font-size:13px;animation:slideUp .2s;z-index:300;max-width:340px;display:flex;align-items:center;gap:10px;box-shadow:0 4px 16px rgba(33,51,99,.15)}
.n-green{border-left-color:var(--accent2)}.n-red{border-left-color:var(--danger)}.n-amber{border-left-color:var(--warn)}.n-blue{border-left-color:var(--blue)}
.info-box{background:var(--surface2);border:1px solid var(--border);border-radius:var(--r);padding:12px 14px;font-size:13px}
.info-box.accent{border-left:3px solid var(--blue)}
.info-box.warn{border-left:3px solid var(--warn);background:#FFFBEB}
.info-box.danger{border-left:3px solid var(--danger);background:#FEF2F2}
.info-box.orange{border-left:3px solid var(--orange);background:#FFF7ED}
.flex-gap{display:flex;gap:8px;align-items:center}.flex-between{display:flex;justify-content:space-between;align-items:center}
.mt8{margin-top:8px}.mt12{margin-top:12px}.mt16{margin-top:16px}
.mb8{margin-bottom:8px}.mb12{margin-bottom:12px}
.text-mono{font-family:var(--mono)}.text-muted{color:var(--muted)}
.empty-state{text-align:center;padding:48px 20px;color:var(--muted);font-size:13px}
.loading{display:flex;align-items:center;justify-content:center;padding:48px;color:var(--muted);gap:10px;font-size:13px}
.spin{animation:spin 1s linear infinite}
@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
.kbar{margin-bottom:10px}.kbar-lbl{display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px}.kbar-track{height:6px;background:var(--surface2);border-radius:3px;overflow:hidden;border:1px solid var(--border)}.kbar-fill{height:100%;border-radius:3px}
.tabs-row{display:flex;gap:0;border-bottom:2px solid var(--border);margin-bottom:18px;overflow-x:auto}
.tab{font-size:11px;font-weight:600;padding:9px 16px;cursor:pointer;color:var(--muted);border-bottom:2px solid transparent;transition:all .12s;text-transform:uppercase;letter-spacing:.5px;margin-bottom:-2px;white-space:nowrap}
.tab.active{color:var(--blue);border-bottom-color:var(--blue)}
.grupo-chip{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:6px;font-family:var(--mono);font-size:12px;font-weight:700;background:#DBEAFE;color:var(--blue);border:1px solid #BFDBFE;flex-shrink:0}
.tag{display:inline-block;font-family:var(--mono);font-size:9px;padding:2px 7px;background:var(--surface2);border:1px solid var(--border);border-radius:4px;color:var(--muted)}
.tracker-simple-row{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:14px 16px;margin-bottom:8px}
.tracker-simple-row.vencida{border-left:4px solid var(--danger)}
.tracker-simple-row.en-curso{border-left:4px solid var(--warn)}
.tracker-simple-row.entregado{border-left:4px solid var(--accent2)}
`;

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

function StatusBadge({ status }) {
  const colorMap = {
    borrador: "b-gray", pendiente_aprobacion: "b-amber", aprobado_cotizar: "b-blue",
    en_cotizacion: "b-teal", pendiente_confirmacion: "b-orange", aprobado: "b-green",
    rechazado: "b-red", en_compra: "b-purple", entregado: "b-green", cerrado: "b-gray",
  };
  return <span className={`badge ${colorMap[status] || "b-gray"}`}>{STATUS_LABELS[status] || status}</span>;
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
    if (ev.includes("creada")) return { i: "◎", c: "c" };
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

// ─── MODAL: APROBAR CONDICIONAL ───────────────────────────────────────────────
function AprobarCondicionalModal({ req, proveedores, onClose, onSave }) {
  const blank = () => ({ id: `tmp${Date.now()}${Math.random()}`, descripcion: "", cantidad: 1, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "" });
  const [items, setItems] = useState(req.requisicion_items?.length ? req.requisicion_items.map(it => ({ ...it })) : [blank()]);
  const [nota, setNota] = useState("");
  const [saving, setSaving] = useState(false);
  const setItem = (i, k, v) => { const its = [...items]; its[i] = { ...its[i], [k]: v }; setItems(its); };

  const handleSave = async () => {
    setSaving(true);
    try {
      await api.actualizarItems(req.id, items.filter(it => it.descripcion?.trim()).map(({ id: _id, requisicion_id: _rid, ...rest }) => rest));
      await api.actualizarRequisicion(req.id, { status: "aprobado_cotizar", revisado_por: USUARIO, fecha_aprobacion: new Date().toISOString() }, `Aprobado con modificaciones${nota ? ` — ${nota}` : ""}`, nota || null);
      onSave(await api.getRequisicion(req.id));
    } finally { setSaving(false); }
  };

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="mhdr">
          <div><div className="mtitle">APROBAR CON MODIFICACIONES</div><div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}</div></div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          <div className="info-box warn mb12" style={{ fontSize: 11 }}>Editá ítems antes de aprobar. Los cambios quedan registrados.</div>
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
          <div className="mt12"><FG label="Nota (opcional)"><textarea value={nota} onChange={e => setNota(e.target.value)} placeholder="Ej: Se ajustó cantidad del ítem 1..." /></FG></div>
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? "..." : "Aprobar con cambios → Tracker"}</button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: CONFIRMAR VALOR ───────────────────────────────────────────────────
function ConfirmarValorModal({ req, onClose, onSave }) {
  const [aprobado, setAprobado] = useState(true);
  const [nota, setNota] = useState("");
  const [saving, setSaving] = useState(false);

  const handleSave = async () => {
    setSaving(true);
    try {
      const nuevoStatus = aprobado ? "aprobado" : "rechazado";
      await api.actualizarRequisicion(req.id, { status: nuevoStatus }, aprobado ? `Valor confirmado y aprobado${nota ? ` — ${nota}` : ""}` : `Rechazado por valor${nota ? ` — ${nota}` : ""}`, nota || null);
      onSave();
    } finally { setSaving(false); }
  };

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 540 }}>
        <div className="mhdr"><div className="mtitle">CONFIRMACIÓN DE VALOR</div><button className="mclose" onClick={onClose}>✕</button></div>
        <div className="mbody">
          <div className="info-box orange mb12" style={{ fontSize: 12 }}>El comprador solicita confirmación del valor cotizado antes de proceder.</div>
          <div style={{ fontSize: 13, color: "var(--muted)", marginBottom: 14 }}>REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}</div>
          <div style={{ background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: "var(--r)", padding: "12px 14px", marginBottom: 14 }}>
            <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer", marginBottom: 12 }}>
              <input type="radio" checked={aprobado} onChange={() => setAprobado(true)} style={{ marginTop: 2, accentColor: "var(--accent2)" }} />
              <div><div style={{ fontSize: 13, fontWeight: 600, color: "var(--accent2)" }}>✓ Confirmar y aprobar</div><div style={{ fontSize: 11, color: "var(--muted)", marginTop: 2 }}>El valor es aceptable. Se procede con la compra.</div></div>
            </label>
            <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer" }}>
              <input type="radio" checked={!aprobado} onChange={() => setAprobado(false)} style={{ marginTop: 2, accentColor: "var(--danger)" }} />
              <div><div style={{ fontSize: 13, fontWeight: 600, color: "var(--danger)" }}>✕ Rechazar por valor</div><div style={{ fontSize: 11, color: "var(--muted)", marginTop: 2 }}>El valor excede lo aceptable.</div></div>
            </label>
          </div>
          <FG label="Comentario (opcional)"><textarea value={nota} onChange={e => setNota(e.target.value)} placeholder="Explicación adicional..." /></FG>
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className={`btn ${aprobado ? "btn-success" : "btn-danger"}`} onClick={handleSave} disabled={saving}>
            {saving ? "..." : aprobado ? "✓ Confirmar y aprobar" : "✕ Rechazar"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: VER REQUISICIÓN ───────────────────────────────────────────────────
function ReqModal({ req: initialReq, proveedores, onClose, onUpdate, onMoverTracker, onRechazar }) {
  const [req, setReq] = useState(initialReq);
  const [tab, setTab] = useState("detalle");
  const [saving, setSaving] = useState(false);
  const [showAprobarCond, setShowAprobarCond] = useState(false);
  const [showConfirmarValor, setShowConfirmarValor] = useState(false);
  const remitoRef = useRef();
  const [uploadingRemito, setUploadingRemito] = useState(false);

  const esParaAprobar = req.status === "pendiente_aprobacion";
  const esParaCotizar = req.status === "aprobado_cotizar";
  const esParaConfirmar = req.status === "pendiente_confirmacion";

  const handleAprobar = async () => {
    setSaving(true);
    try {
      await api.actualizarRequisicion(req.id, { status: "aprobado_cotizar", revisado_por: USUARIO, fecha_aprobacion: new Date().toISOString() }, "Aprobado para cotizar");
      const fresh = await api.getRequisicion(req.id);
      onUpdate(fresh); onMoverTracker(fresh);
    } finally { setSaving(false); }
  };

  const handleSolicitarConfirmacion = async () => {
    setSaving(true);
    try {
      await api.actualizarRequisicion(req.id, { status: "pendiente_confirmacion" }, "Pendiente confirmación de valor");
      onUpdate({ ...req, status: "pendiente_confirmacion" }); onClose();
    } finally { setSaving(false); }
  };

  const handleUploadRemito = async (file) => {
    setUploadingRemito(true);
    try {
      const url = await api.subirAdjunto(file, `remitos/${req.id}/${Date.now()}_${file.name}`);
      const updated = await api.actualizarRequisicion(req.id, { status: "entregado", remito_url: url }, "Remito adjuntado — entregado");
      setReq(updated); onUpdate(updated);
    } catch (e) { alert("Error al subir remito: " + e.message); }
    finally { setUploadingRemito(false); }
  };

  if (showAprobarCond) return <AprobarCondicionalModal req={req} proveedores={proveedores} onClose={() => setShowAprobarCond(false)} onSave={(fresh) => { setShowAprobarCond(false); onUpdate(fresh); onMoverTracker(fresh); }} />;
  if (showConfirmarValor) return <ConfirmarValorModal req={req} onClose={() => setShowConfirmarValor(false)} onSave={() => { setShowConfirmarValor(false); onClose(); onUpdate({ ...req, status: "aprobado" }); }} />;

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="mhdr">
          <div>
            <div className="mtitle">REQ-{String(req.nro_solicitud).padStart(4, "0")} — {req.titulo}</div>
            <div className="flex-gap mt8">
              <StatusBadge status={req.status} /><UrgBadge urgencia={req.urgencia} />
              <span className="tag">{req.base_buque}</span><span className="tag">{req.area}</span>
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
              <div className="info-box"><div style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", marginBottom: 4 }}>TIPO</div>{req.tipo_requisicion || "—"}</div>
              <div className="info-box"><div style={{ fontSize: 10, color: "var(--muted)", fontFamily: "var(--mono)", marginBottom: 4 }}>SUB-ÁREA</div>{req.subarea || "—"}</div>
            </div>
            {req.observaciones && <div className="info-box mb12">{req.observaciones}</div>}
            {req.remito_url && <div className="info-box accent mb12" style={{ fontSize: 12 }}>
              📎 Remito adjunto: <a href={req.remito_url} target="_blank" rel="noreferrer" style={{ color: "var(--blue)" }}>Ver remito firmado</a>
            </div>}
            <div style={{ fontFamily: "var(--mono)", fontSize: 9, letterSpacing: 2, color: "var(--muted)", textTransform: "uppercase", marginBottom: 8 }}>Ítems</div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>#</th><th>Descripción</th><th>Cant.</th><th>Unid.</th><th>Proveedor sugerido</th></tr></thead>
                <tbody>
                  {(req.requisicion_items || []).map((it, i) => <tr key={i}>
                    <td className="text-mono text-muted">{it.nro_linea}</td>
                    <td>{it.descripcion}</td>
                    <td className="text-mono">{it.cantidad}</td>
                    <td className="text-muted">{it.unidad}</td>
                    <td className="text-muted">{it.proveedor_sugerido || "—"}</td>
                  </tr>)}
                </tbody>
              </table>
            </div>
            {(req.status === "en_compra" || req.status === "aprobado") && !req.remito_url && <>
              <div className="form-section">Adjuntar remito firmado</div>
              <div className="info-box warn mb12" style={{ fontSize: 11 }}>Adjuntá el remito firmado para cerrar la requisición.</div>
              <input ref={remitoRef} type="file" accept=".pdf,.jpg,.jpeg,.png" style={{ display: "none" }} onChange={e => handleUploadRemito(e.target.files[0])} />
              <button className="btn btn-ghost btn-sm" onClick={() => remitoRef.current.click()} disabled={uploadingRemito}>{uploadingRemito ? "⏳ Subiendo..." : "📎 Adjuntar remito"}</button>
            </>}
          </>}
          {tab === "historial" && <Timeline historial={req.requisicion_historial} />}
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cerrar</button>
          {(esParaAprobar || esParaCotizar) && <button className="btn btn-danger" onClick={() => { onClose(); onRechazar(req); }}>Rechazar</button>}
          {(esParaAprobar || esParaCotizar) && <button className="btn btn-cond" onClick={() => setShowAprobarCond(true)}>Aprobar condicional</button>}
          {esParaCotizar && <button className="btn btn-confirm" onClick={handleSolicitarConfirmacion} disabled={saving}>Solicitar conf. valor</button>}
          {(esParaAprobar || esParaCotizar) && <button className="btn btn-primary" onClick={handleAprobar} disabled={saving}>{saving ? "..." : "Aprobar → Tracker"}</button>}
          {esParaConfirmar && <button className="btn btn-primary" onClick={() => setShowConfirmarValor(true)}>Confirmar valor</button>}
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: CONSOLIDAR ────────────────────────────────────────────────────────
function ConsolidarModal({ req, onClose, onSave }) {
  const items = req.requisicion_items || [];
  const [asignaciones, setAsignaciones] = useState(items.map(() => "A"));
  const [saving, setSaving] = useState(false);
  const grupos = [...new Set(asignaciones)].sort();

  const handleSave = async () => {
    setSaving(true);
    try {
      await api.crearTrackerLineas(req.id, grupos.map(g => ({
        grupo: g, descripcion: `Grupo ${g} — REQ-${String(req.nro_solicitud).padStart(4, "0")}`,
        items_detalle: items.filter((_, i) => asignaciones[i] === g), status: "en_cotizacion",
      })));
      await api.actualizarRequisicion(req.id, { status: "en_compra" }, `Movido al Tracker (${grupos.length} grupo${grupos.length > 1 ? "s" : ""})`);
      onSave();
    } finally { setSaving(false); }
  };

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="mhdr"><div className="mtitle">CONSOLIDAR EN TRACKER</div><button className="mclose" onClick={onClose}>✕</button></div>
        <div className="mbody">
          <div className="info-box accent mb12" style={{ fontSize: 11 }}>Asigná cada ítem a un grupo. Ítems del mismo grupo se consolidan en una línea.</div>
          <table className="items-edit">
            <thead><tr><th style={{ width: "50%" }}>Ítem</th><th>Cant.</th><th>Grupo</th></tr></thead>
            <tbody>
              {items.map((it, i) => <tr key={i}>
                <td>{it.descripcion}</td>
                <td className="text-mono">{it.cantidad} {it.unidad}</td>
                <td><select value={asignaciones[i]} onChange={e => { const a = [...asignaciones]; a[i] = e.target.value; setAsignaciones(a); }} style={{ width: 60 }}>{GRUPOS_OPCIONES.map(g => <option key={g}>{g}</option>)}</select></td>
              </tr>)}
            </tbody>
          </table>
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving ? "Guardando..." : `Crear ${grupos.length} línea${grupos.length > 1 ? "s" : ""}`}</button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: RECHAZAR ─────────────────────────────────────────────────────────
function RechazarModal({ req, onClose, onSave }) {
  const [categoria, setCategoria] = useState("");
  const [texto, setTexto] = useState("");
  const [devolver, setDevolver] = useState(true);
  const [saving, setSaving] = useState(false);

  const handleSave = async () => {
    if (!categoria) return alert("Seleccioná una categoría");
    setSaving(true);
    try {
      const updated = await api.actualizarRequisicion(req.id, {
        status: devolver ? "pendiente_aprobacion" : "rechazado",
        motivo_rechazo_categoria: categoria, motivo_rechazo_texto: texto,
        veces_devuelto: (req.veces_devuelto || 0) + 1,
      }, devolver ? `Devuelta — ${categoria}` : `Rechazada — ${categoria}`, texto || null);
      onSave(updated, devolver);
    } finally { setSaving(false); }
  };

  return (
    <div className="overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 520 }}>
        <div className="mhdr"><div className="mtitle">RECHAZAR / DEVOLVER</div><button className="mclose" onClick={onClose}>✕</button></div>
        <div className="mbody">
          <FG label="Categoría *">
            <select value={categoria} onChange={e => setCategoria(e.target.value)}>
              <option value="">Seleccionar...</option>
              {CATEGORIAS_RECHAZO.map(c => <option key={c}>{c}</option>)}
            </select>
          </FG>
          <div className="mt12"><FG label="Detalle adicional"><textarea value={texto} onChange={e => setTexto(e.target.value)} /></FG></div>
          <div className="mt12" style={{ background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: "var(--r)", padding: "12px 14px" }}>
            <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer", marginBottom: 10 }}>
              <input type="radio" checked={devolver} onChange={() => setDevolver(true)} style={{ marginTop: 2, accentColor: "var(--warn)" }} />
              <div><div style={{ fontSize: 13, fontWeight: 600, color: "var(--warn)" }}>↩ Devolver para corrección</div></div>
            </label>
            <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer" }}>
              <input type="radio" checked={!devolver} onChange={() => setDevolver(false)} style={{ marginTop: 2, accentColor: "var(--danger)" }} />
              <div><div style={{ fontSize: 13, fontWeight: 600, color: "var(--danger)" }}>✕ Rechazar definitivamente</div></div>
            </label>
          </div>
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className={`btn ${devolver ? "btn-warn" : "btn-danger"}`} onClick={handleSave} disabled={saving || !categoria}>
            {saving ? "..." : devolver ? "↩ Devolver" : "✕ Rechazar"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL: TRACKER LÍNEA ─────────────────────────────────────────────────────
function TrackerLineaModal({ linea, proveedores, onClose, onSave }) {
  const emptyCotiz = () => ({ proveedor: "", precio: "", moneda: "ARS", plazo: "" });
  const initCotiz = () => { const c = linea.cotizaciones || {}; return [c.c1 || emptyCotiz(), c.c2 || emptyCotiz(), c.c3 || emptyCotiz()]; };
  const [form, setForm] = useState({
    descripcion: linea.descripcion || "", proveedor_elegido: linea.proveedor_elegido || "",
    motivo_proveedor: linea.motivo_proveedor || "", nro_oc: linea.nro_oc || "",
    costo_real: linea.costo_real || "", moneda_real: linea.moneda_real || "ARS",
    plazo_pago: linea.plazo_pago || "", fecha_entrega_prom: linea.fecha_entrega_prom || "",
    fecha_entrega_real: linea.fecha_entrega_real || "", status: linea.status || "en_cotizacion",
    notas: linea.notas || "", nro_remito: linea.nro_remito || "",
  });
  const [cotiz, setCotiz] = useState(initCotiz());
  const [adjuntos, setAdjuntos] = useState(linea.cotizaciones?.adjuntos || []);
  const [uploading, setUploading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [showDetail, setShowDetail] = useState(false);
  const fileRef = useRef();
  const remitoRef = useRef();

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const setCotizField = (idx, k, v) => {
    const next = cotiz.map((c, i) => i === idx ? { ...c, [k]: v } : c);
    setCotiz(next);
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
      descripcion: f.descripcion || null, proveedor_elegido: f.proveedor_elegido || null,
      motivo_proveedor: f.motivo_proveedor || null, nro_oc: f.nro_oc || null,
      costo_real: f.costo_real !== "" && f.costo_real != null ? parseFloat(f.costo_real) : null,
      moneda_real: f.moneda_real || "ARS", plazo_pago: f.plazo_pago || null,
      fecha_entrega_prom: f.fecha_entrega_prom || null, fecha_entrega_real: f.fecha_entrega_real || null,
      status: f.status || "en_cotizacion", notas: f.notas || null, nro_remito: f.nro_remito || null,
      cotizaciones: { c1: cotiz[0], c2: cotiz[1], c3: cotiz[2], adjuntos },
    };
  };

  const handleSave = async (extra = {}) => {
    setSaving(true);
    try { onSave(await api.actualizarTrackerLinea(linea.id, buildPayload(extra))); }
    catch (e) { alert("Error: " + e.message); }
    finally { setSaving(false); }
  };

  const handleEntrega = async () => {
    setSaving(true);
    try { onSave(await api.actualizarTrackerLinea(linea.id, buildPayload({ status: "entregado", fecha_entrega_real: form.fecha_entrega_real || new Date().toISOString().split("T")[0] }))); }
    catch (e) { alert("Error."); }
    finally { setSaving(false); }
  };

  const handleUpload = async (files) => {
    if (!files?.length) return;
    setUploading(true);
    try {
      const nuevos = [];
      for (const file of Array.from(files)) {
        const path = `${linea.id}/${Date.now()}_${file.name}`;
        const url = await api.subirAdjunto(file, path);
        nuevos.push({ nombre: file.name, url, path });
      }
      setAdjuntos(prev => [...prev, ...nuevos]);
    } catch (e) { alert("Error al subir archivo."); }
    finally { setUploading(false); }
  };

  const handleUploadRemito = async (file) => {
    setUploading(true);
    try {
      const path = `remitos/${linea.id}/${Date.now()}_${file.name}`;
      const url = await api.subirAdjunto(file, path);
      set("nro_remito", form.nro_remito || file.name);
      setAdjuntos(prev => [...prev, { nombre: `REMITO: ${file.name}`, url, path }]);
    } catch (e) { alert("Error al subir remito."); }
    finally { setUploading(false); }
  };

  const req = linea.requisiciones;
  const itemsDetalle = linea.items_detalle || [];
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
            <div className="flex-gap"><div className="grupo-chip">{linea.grupo}</div><div className="mtitle">{form.descripcion}</div></div>
            {req && <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>REQ-{String(req.nro_solicitud).padStart(4, "0")} · {req.base_buque}</div>}
          </div>
          <button className="mclose" onClick={onClose}>✕</button>
        </div>
        <div className="mbody">
          {itemsDetalle.length > 0 && <div className="mb12">
            <button className="btn btn-ghost btn-sm" onClick={() => setShowDetail(!showDetail)}>{showDetail ? "▲" : "▼"} Ver ítems ({itemsDetalle.length})</button>
            {showDetail && <div style={{ background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: "var(--r)", padding: "10px 12px", marginTop: 8, fontSize: 11, color: "var(--muted)" }}>
              {itemsDetalle.map((it, i) => <div key={i} style={{ padding: "2px 0" }}>· {it.descripcion} × {it.cantidad} {it.unidad}</div>)}
            </div>}
          </div>}

          <div className="form-section">Estado</div>
          <div className="form-grid">
            <FG label="Status"><select value={form.status} onChange={e => set("status", e.target.value)}>{Object.entries(TRACKER_STATUS).filter(([k]) => k !== "archivado").map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></FG>
            <FG label="Descripción"><input value={form.descripcion} onChange={e => set("descripcion", e.target.value)} /></FG>
          </div>

          <div className="form-section">Cotizaciones</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
            {cotiz.map((c, i) => (
              <div key={i} style={{ borderRadius: "var(--r2)", padding: "12px 14px", ...COTIZ_STYLES[i] }}>
                <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", color: i === 0 ? "var(--accent2)" : "var(--muted)", marginBottom: 10 }}>{i === 0 && "⭐ "}{COTIZ_LABELS[i]}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  <FG label="Proveedor"><select value={c.proveedor} onChange={e => setCotizField(i, "proveedor", e.target.value)} style={{ fontSize: 12 }}><option value="">Seleccionar...</option>{proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}</select></FG>
                  <FG label="Precio"><input type="number" value={c.precio} onChange={e => setCotizField(i, "precio", e.target.value)} style={{ fontSize: 12 }} /></FG>
                  <FG label="Moneda"><select value={c.moneda} onChange={e => setCotizField(i, "moneda", e.target.value)} style={{ fontSize: 12 }}><option>ARS</option><option>USD</option></select></FG>
                  <FG label="Plazo"><select value={c.plazo} onChange={e => setCotizField(i, "plazo", e.target.value)} style={{ fontSize: 12 }}><option value="">—</option>{PLAZO_PAGO_OPTIONS.map(p => <option key={p}>{p}</option>)}</select></FG>
                </div>
              </div>
            ))}
          </div>

          <div className="form-section">OC y entrega</div>
          <div className="form-grid">
            <FG label="Proveedor elegido"><select value={form.proveedor_elegido} onChange={e => set("proveedor_elegido", e.target.value)}><option value="">Seleccionar...</option>{proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}</select></FG>
            <FG label="N° OC"><input value={form.nro_oc} onChange={e => set("nro_oc", e.target.value)} placeholder="OC-0001" /></FG>
          </div>
          <FG label="¿Por qué este proveedor?"><textarea value={form.motivo_proveedor} onChange={e => set("motivo_proveedor", e.target.value)} /></FG>
          <div className="form-grid-3 mt12">
            <FG label="Costo real"><input type="number" value={form.costo_real} onChange={e => set("costo_real", e.target.value)} /></FG>
            <FG label="Moneda"><select value={form.moneda_real} onChange={e => set("moneda_real", e.target.value)}><option>ARS</option><option>USD</option></select></FG>
            <FG label="Plazo de pago"><select value={form.plazo_pago} onChange={e => set("plazo_pago", e.target.value)}><option value="">—</option>{PLAZO_PAGO_OPTIONS.map(p => <option key={p}>{p}</option>)}</select></FG>
            <FG label="Entrega prometida"><input type="date" value={form.fecha_entrega_prom} onChange={e => set("fecha_entrega_prom", e.target.value)} /></FG>
            <FG label="Entrega real"><input type="date" value={form.fecha_entrega_real} onChange={e => set("fecha_entrega_real", e.target.value)} /></FG>
            <FG label="N° Remito"><input value={form.nro_remito} onChange={e => set("nro_remito", e.target.value)} placeholder="0001-00001234" /></FG>
          </div>
          <FG label="Notas"><textarea value={form.notas} onChange={e => set("notas", e.target.value)} /></FG>

          <div className="form-section">Adjuntos (presupuestos y remito)</div>
          <input ref={fileRef} type="file" multiple accept=".pdf,.jpg,.jpeg,.png,.xlsx,.xls" style={{ display: "none" }} onChange={e => handleUpload(e.target.files)} />
          <input ref={remitoRef} type="file" accept=".pdf,.jpg,.jpeg,.png" style={{ display: "none" }} onChange={e => handleUploadRemito(e.target.files[0])} />
          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <button className="btn btn-ghost btn-sm" onClick={() => fileRef.current.click()} disabled={uploading}>📎 Adjuntar presupuesto</button>
            <button className="btn btn-ghost btn-sm" onClick={() => remitoRef.current.click()} disabled={uploading}>📄 Adjuntar remito firmado</button>
          </div>
          {adjuntos.length > 0 && <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
            {adjuntos.map((adj, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: "var(--r)", padding: "6px 10px" }}>
                <span>📄</span>
                <a href={adj.url} target="_blank" rel="noreferrer" style={{ fontSize: 12, color: "var(--accent)", flex: 1 }}>{adj.nombre}</a>
                <button onClick={async () => { await supabase.storage.from("cotizaciones").remove([adj.path]); setAdjuntos(prev => prev.filter(a => a.path !== adj.path)); }} style={{ background: "none", border: "none", color: "var(--muted2)", cursor: "pointer" }}>✕</button>
              </div>
            ))}
          </div>}
        </div>
        <div className="mftr">
          <button className="btn btn-ghost" onClick={onClose}>Cancelar</button>
          <button className="btn btn-success btn-sm" onClick={handleEntrega} disabled={saving}>✓ Confirmar entrega</button>
          <button className="btn btn-primary" onClick={() => handleSave()} disabled={saving}>{saving ? "Guardando..." : "Guardar"}</button>
        </div>
      </div>
    </div>
  );
}

// ─── PAGE: INBOX ──────────────────────────────────────────────────────────────
function PageInbox({ statuses, titulo, notify, onNeedRefresh }) {
  const [reqs, setReqs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [consolidando, setConsolidando] = useState(null);
  const [rechazando, setRechazando] = useState(null);
  const [proveedores, setProveedores] = useState([]);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [data, provs] = await Promise.all([api.getRequisiciones({ empresa: "Parana Logistica", statuses }), api.getProveedores()]);
      setReqs(data); setProveedores(provs);
    } finally { setLoading(false); }
  }, [statuses.join(",")]);

  useEffect(() => { load(); }, [load]);

  const handleUpdate = (updated) => setReqs(prev => prev.map(r => r.id === updated.id ? { ...r, ...updated } : r).filter(r => statuses.includes(r.status)));

  const handleMoverTracker = async (req) => {
    setSelected(null);
    setConsolidando(await api.getRequisicion(req.id));
  };

  const handleRechazado = (updated, devolver) => {
    setRechazando(null);
    if (devolver) { setReqs(prev => prev.map(r => r.id === updated.id ? { ...r, ...updated } : r)); notify("Devuelta al solicitante", "warn"); }
    else { setReqs(prev => prev.filter(r => r.id !== updated.id)); notify("Rechazada definitivamente", "warn"); }
    onNeedRefresh();
  };

  const rowClass = (r) => {
    if (r.status === "pendiente_confirmacion") return "req-row pendiente-confirmacion";
    if (r.veces_devuelto > 0) return "req-row devuelto";
    return "req-row unread";
  };

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16, paddingBottom: 12, borderBottom: "2px solid var(--border)" }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: "var(--navy)" }}>📥 {titulo}</div>
        <span className="ni-badge" style={{ position: "static" }}>{reqs.length}</span>
      </div>
      {loading ? <div className="loading"><span className="spin">◌</span> Cargando...</div> :
        reqs.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}>📭</div>Sin requisiciones pendientes</div> :
        reqs.map(r => (
          <div key={r.id} className={rowClass(r)} onClick={() => setSelected(r)}>
            <div className="flex-between mb8">
              <div className="flex-gap">
                <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(r.nro_solicitud).padStart(4, "0")}</span>
                <StatusBadge status={r.status} /><UrgBadge urgencia={r.urgencia} />
                {r.veces_devuelto > 0 && <span className="badge b-orange">↩ {r.veces_devuelto}x</span>}
              </div>
              <span style={{ fontSize: 10, color: "var(--muted)" }}>{fmtDate(r.created_at)}</span>
            </div>
            <div className="req-title">{r.titulo}</div>
            <div className="req-meta">
              <span>{r.base_buque}</span><span>·</span>
              <span>{r.area}{r.subarea ? ` › ${r.subarea}` : ""}</span><span>·</span>
              <span>{r.solicitado_por}</span>
              {r.fecha_necesaria && <><span>·</span><span style={{ color: "var(--warn)" }}>Nec: {fmtDate(r.fecha_necesaria)}</span></>}
            </div>
          </div>
        ))
      }
      {selected && <ReqModal req={selected} proveedores={proveedores} onClose={() => setSelected(null)} onUpdate={handleUpdate} onMoverTracker={handleMoverTracker} onRechazar={r => { setSelected(null); setRechazando(r); }} />}
      {consolidando && <ConsolidarModal req={consolidando} onClose={() => setConsolidando(null)} onSave={() => { setConsolidando(null); notify("Movida al Tracker", "success"); load(); onNeedRefresh(); }} />}
      {rechazando && <RechazarModal req={rechazando} onClose={() => setRechazando(null)} onSave={handleRechazado} />}
    </div>
  );
}

// ─── PAGE: TRACKER COMPLETO ───────────────────────────────────────────────────
function PageTrackerGeneral({ notify, onNeedRefresh }) {
  const [lineas, setLineas] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [proveedores, setProveedores] = useState([]);
  const [filtros, setFiltros] = useState({ status: "", proveedor: "", busqueda: "" });
  const [sortCol, setSortCol] = useState("created_at");
  const [sortDir, setSortDir] = useState("desc");

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [data, provs] = await Promise.all([api.getTrackerLineas({ statuses: ["en_cotizacion", "oc_emitida", "en_transito", "entregado"] }), api.getProveedores()]);
      setLineas(data); setProveedores(provs);
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, [load]);

  const handleSave = (updated) => { setSelected(null); notify("Línea actualizada", "success"); load(); onNeedRefresh?.(); };
  const handleSort = (col) => { if (sortCol === col) setSortDir(d => d === "asc" ? "desc" : "asc"); else { setSortCol(col); setSortDir("asc"); } };

  const lineasFiltradas = lineas.filter(l => {
    const req = l.requisiciones;
    if (filtros.status && l.status !== filtros.status) return false;
    if (filtros.proveedor && l.proveedor_elegido !== filtros.proveedor) return false;
    if (filtros.busqueda) {
      const q = filtros.busqueda.toLowerCase();
      if (!(l.descripcion?.toLowerCase().includes(q) || req?.nro_solicitud?.toString().includes(q) || req?.base_buque?.toLowerCase().includes(q) || l.proveedor_elegido?.toLowerCase().includes(q) || l.nro_oc?.toLowerCase().includes(q))) return false;
    }
    return true;
  }).sort((a, b) => {
    let va, vb;
    const ra = a.requisiciones, rb = b.requisiciones;
    switch (sortCol) {
      case "nro": va = ra?.nro_solicitud || 0; vb = rb?.nro_solicitud || 0; break;
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
  const proveedoresDisponibles = [...new Set(lineas.map(l => l.proveedor_elegido).filter(Boolean))];
  const totalARS = lineas.filter(l => l.costo_real && (l.moneda_real === "ARS" || !l.moneda_real)).reduce((a, l) => a + l.costo_real, 0);
  const totalUSD = lineas.filter(l => l.costo_real && l.moneda_real === "USD").reduce((a, l) => a + l.costo_real, 0);

  const handleExport = () => {
    const rows = lineasFiltradas.map(l => {
      const req = l.requisiciones;
      return {
        "REQ": req ? `REQ-${String(req.nro_solicitud).padStart(4, "0")}` : "", "Grupo": l.grupo || "",
        "Descripción": l.descripcion || "", "Base/Buque": req?.base_buque || "",
        "Área": req?.area || "", "Solicitante": req?.solicitado_por || "",
        "Urgencia": req?.urgencia || "", "Estado": TRACKER_STATUS[l.status]?.label || l.status || "",
        "Proveedor": l.proveedor_elegido || "", "N° OC": l.nro_oc || "",
        "Costo real": l.costo_real || "", "Moneda": l.moneda_real || "",
        "Plazo pago": l.plazo_pago || "", "N° Remito": l.nro_remito || "",
        "Entrega prometida": l.fecha_entrega_prom ? fmtDate(l.fecha_entrega_prom) : "",
        "Entrega real": l.fecha_entrega_real ? fmtDate(l.fecha_entrega_real) : "",
        "Notas": l.notas || "",
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tracker");
    XLSX.writeFile(wb, `tracker_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div>
      <div className="stats">
        <div className="stat"><div className="stat-label">Total líneas</div><div className="stat-value va">{lineas.length}</div></div>
        <div className="stat"><div className="stat-label">En curso</div><div className="stat-value vm">{lineas.filter(l => ["en_cotizacion", "oc_emitida", "en_transito"].includes(l.status)).length}</div></div>
        <div className="stat"><div className="stat-label">Entregadas</div><div className="stat-value vg">{lineas.filter(l => l.status === "entregado").length}</div></div>
        <div className="stat">
          <div className="stat-label">Comprometido</div>
          <div style={{ marginTop: 4 }}>
            {totalARS > 0 && <div className="text-mono" style={{ fontSize: 13, fontWeight: 700, color: "var(--accent)" }}>{fmt(totalARS, "ARS")}</div>}
            {totalUSD > 0 && <div className="text-mono" style={{ fontSize: 13, fontWeight: 700, color: "var(--accent2)" }}>{fmt(totalUSD, "USD")}</div>}
            {totalARS === 0 && totalUSD === 0 && <div style={{ color: "var(--muted2)", fontSize: 12 }}>—</div>}
          </div>
        </div>
      </div>

      <div className="filter-row">
        <input className="filter-input" placeholder="🔍 Buscar..." value={filtros.busqueda} onChange={e => setFiltros(f => ({ ...f, busqueda: e.target.value }))} />
        <select className="filter-select" value={filtros.status} onChange={e => setFiltros(f => ({ ...f, status: e.target.value }))}>
          <option value="">Todos los estados</option>
          {Object.entries(TRACKER_STATUS).filter(([k]) => k !== "archivado").map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
        </select>
        <select className="filter-select" value={filtros.proveedor} onChange={e => setFiltros(f => ({ ...f, proveedor: e.target.value }))}>
          <option value="">Todos los proveedores</option>
          {proveedoresDisponibles.map(p => <option key={p}>{p}</option>)}
        </select>
        {(filtros.status || filtros.proveedor || filtros.busqueda) && <button className="btn btn-ghost btn-sm" onClick={() => setFiltros({ status: "", proveedor: "", busqueda: "" })}>✕ Limpiar</button>}
        <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--muted)", fontFamily: "var(--mono)" }}>{lineasFiltradas.length} de {lineas.length}</span>
        <button className="btn btn-ghost btn-sm" onClick={handleExport}>↓ Excel</button>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        lineas.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}>📋</div>Sin líneas</div> :
        <div className="card" style={{ padding: 0, overflow: "hidden" }}>
          <div className="table-wrap">
            <table className="tracker-table">
              <thead>
                <tr>
                  <th className="sortable" onClick={() => handleSort("nro")}>REQ<SortIcon col="nro" /></th>
                  <th>Descripción</th>
                  <th className="sortable" onClick={() => handleSort("buque")}>Base/Buque<SortIcon col="buque" /></th>
                  <th>Urgencia</th>
                  <th>Estado</th>
                  <th className="sortable" onClick={() => handleSort("proveedor")}>Proveedor<SortIcon col="proveedor" /></th>
                  <th>OC</th>
                  <th>Remito</th>
                  <th className="sortable" onClick={() => handleSort("costo")}>Costo<SortIcon col="costo" /></th>
                  <th className="sortable" onClick={() => handleSort("entrega")}>Entrega<SortIcon col="entrega" /></th>
                </tr>
              </thead>
              <tbody>
                {lineasFiltradas.length === 0
                  ? <tr><td colSpan={10} style={{ textAlign: "center", padding: 32, color: "var(--muted)" }}>Sin resultados</td></tr>
                  : lineasFiltradas.map(l => {
                      const req = l.requisiciones;
                      return (
                        <tr key={l.id} onClick={() => setSelected(l)}>
                          <td><div className="flex-gap"><div className="grupo-chip">{l.grupo}</div>{req && <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(req.nro_solicitud).padStart(4, "0")}</span>}</div></td>
                          <td><div style={{ fontWeight: 600, fontSize: 12, maxWidth: 200 }}>{l.descripcion}</div></td>
                          <td style={{ fontSize: 12, color: "var(--muted)" }}>{req?.base_buque || "—"}</td>
                          <td>{req ? <UrgBadge urgencia={req.urgencia} /> : "—"}</td>
                          <td><TrackerBadge status={l.status} /></td>
                          <td style={{ fontSize: 12 }}>{l.proveedor_elegido || <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                          <td>{l.nro_oc ? <span className="text-mono" style={{ fontSize: 11, color: "var(--accent2)" }}>{l.nro_oc}</span> : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                          <td>{l.nro_remito ? <span className="text-mono" style={{ fontSize: 11, color: "var(--accent2)" }}>{l.nro_remito}</span> : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                          <td className="text-mono" style={{ fontSize: 12 }}>{l.costo_real ? fmt(l.costo_real, l.moneda_real) : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                          <td style={{ fontSize: 12, color: "var(--warn)" }}>{l.fecha_entrega_prom ? fmtDate(l.fecha_entrega_prom) : <span style={{ color: "var(--muted2)" }}>—</span>}</td>
                        </tr>
                      );
                    })
                }
              </tbody>
            </table>
          </div>
        </div>
      }
      {selected && <TrackerLineaModal linea={selected} proveedores={proveedores} onClose={() => setSelected(null)} onSave={handleSave} />}
    </div>
  );
}

// ─── PAGE: TRACKER SIMPLIFICADO (embarcados) ──────────────────────────────────
function PageTrackerSimple() {
  const [lineas, setLineas] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filtroBase, setFiltroBase] = useState("");
  const [busqueda, setBusqueda] = useState("");

  useEffect(() => {
    api.getTrackerLineas({ statuses: ["en_cotizacion", "oc_emitida", "en_transito", "entregado"] }).then(d => { setLineas(d); setLoading(false); });
  }, []);

  const bases = [...new Set(lineas.map(l => l.requisiciones?.base_buque).filter(Boolean))].sort();
  const filtradas = lineas.filter(l => {
    const req = l.requisiciones;
    if (filtroBase && req?.base_buque !== filtroBase) return false;
    if (busqueda && !l.descripcion?.toLowerCase().includes(busqueda.toLowerCase()) && !req?.nro_solicitud?.toString().includes(busqueda)) return false;
    return true;
  });

  const rowClass = (l) => {
    if (l.status === "entregado") return "tracker-simple-row entregado";
    return "tracker-simple-row en-curso";
  };

  return (
    <div>
      <div className="info-box accent mb16" style={{ fontSize: 11 }}>
        Vista de seguimiento — mostrás el estado de tus pedidos sin valores ni cotizaciones.
      </div>

      <div className="filter-row">
        <input className="filter-input" placeholder="🔍 Buscar pedido..." value={busqueda} onChange={e => setBusqueda(e.target.value)} />
        <select className="filter-select" value={filtroBase} onChange={e => setFiltroBase(e.target.value)}>
          <option value="">Todos los barcos</option>
          {bases.map(b => <option key={b}>{b}</option>)}
        </select>
        {(filtroBase || busqueda) && <button className="btn btn-ghost btn-sm" onClick={() => { setFiltroBase(""); setBusqueda(""); }}>✕ Limpiar</button>}
        <span style={{ marginLeft: "auto", fontFamily: "var(--mono)", fontSize: 11, color: "var(--muted)" }}>{filtradas.length} pedidos</span>
      </div>

      {loading ? <div className="loading"><span className="spin">◌</span> Cargando...</div> :
        filtradas.length === 0 ? <div className="empty-state"><div style={{ fontSize: 28, marginBottom: 8 }}>📋</div>Sin pedidos</div> :
        filtradas.map(l => {
          const req = l.requisiciones;
          return (
            <div key={l.id} className={rowClass(l)}>
              <div className="flex-between mb8">
                <div className="flex-gap">
                  {req && <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(req.nro_solicitud).padStart(4, "0")}</span>}
                  <TrackerBadge status={l.status} />
                </div>
                <span style={{ fontSize: 10, color: "var(--muted)" }}>{req?.base_buque}</span>
              </div>
              <div style={{ fontWeight: 600, fontSize: 13, color: "var(--navy)", marginBottom: 4 }}>{l.descripcion}</div>
              <div className="req-meta">
                {req?.solicitado_por && <span>{req.solicitado_por}</span>}
                {req?.area && <><span>·</span><span>{req.area}{req.subarea ? ` › ${req.subarea}` : ""}</span></>}
                {l.fecha_entrega_prom && <><span>·</span><span style={{ color: "var(--warn)" }}>Entrega est: {fmtDate(l.fecha_entrega_prom)}</span></>}
                {l.fecha_entrega_real && <><span>·</span><span style={{ color: "var(--accent2)" }}>Entregado: {fmtDate(l.fecha_entrega_real)}</span></>}
                {l.nro_remito && <><span>·</span><span className="text-mono" style={{ color: "var(--accent2)" }}>Remito: {l.nro_remito}</span></>}
              </div>
            </div>
          );
        })
      }
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
        setData(lineas); setProveedores(provs);
      } else {
        setData(await api.getRequisiciones({ status: "rechazado" }));
      }
    } finally { setLoading(false); }
  }, [tipo]);

  useEffect(() => { load(); }, [load]);

  if (tipo === "rechazados") return (
    <div>
      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        data.length === 0 ? <div className="empty-state">Sin rechazadas</div> :
        data.map(r => <div key={r.id} className="req-row">
          <div className="flex-between mb8">
            <span className="text-mono" style={{ fontSize: 11, color: "var(--accent)" }}>REQ-{String(r.nro_solicitud).padStart(4, "0")}</span>
            <span style={{ fontSize: 10, color: "var(--muted)" }}>{fmtDate(r.updated_at)}</span>
          </div>
          <div className="req-title">{r.titulo}</div>
          <div className="req-meta"><span>{r.base_buque}</span>{r.motivo_rechazo_categoria && <><span>·</span><span style={{ color: "var(--danger)" }}>{r.motivo_rechazo_categoria}</span></>}</div>
        </div>)
      }
    </div>
  );

  return (
    <div>
      {loading ? <div className="loading"><span className="spin">◌</span></div> :
        data.length === 0 ? <div className="empty-state">Sin entregas</div> :
        data.map(l => {
          const req = l.requisiciones;
          return <div key={l.id} className="req-row" onClick={() => setSelected(l)}>
            <div className="flex-gap mb8"><div className="grupo-chip">{l.grupo}</div><span style={{ fontWeight: 600, fontSize: 14 }}>{l.descripcion}</span><TrackerBadge status="entregado" /></div>
            <div style={{ fontSize: 12, color: "var(--muted)" }}>
              {req && <span>{req.base_buque}</span>}
              {l.proveedor_elegido && <span> · {l.proveedor_elegido}</span>}
              {l.nro_oc && <span className="text-mono" style={{ color: "var(--accent2)" }}> · {l.nro_oc}</span>}
              {l.nro_remito && <span> · Remito: {l.nro_remito}</span>}
              {l.fecha_entrega_real && <span> · {fmtDate(l.fecha_entrega_real)}</span>}
            </div>
          </div>;
        })
      }
      {selected && <TrackerLineaModal linea={selected} proveedores={proveedores} onClose={() => setSelected(null)} onSave={() => { setSelected(null); load(); }} />}
    </div>
  );
}

// ─── FORM: REQUISICIÓN ────────────────────────────────────────────────────────
function ReqForm({ proveedores = [], onSave, onCancel }) {
  const blank = () => ({ id: `tmp${Date.now()}${Math.random()}`, descripcion: "", cantidad: 1, unidad: "Uni", stock_disponible: 0, proveedor_sugerido: "" });
  const [form, setForm] = useState({
    titulo: "", empresa: "Parana Logistica", base_buque: "", area: "", subarea: "",
    detalle_tecnico: "", tipo_requisicion: "", urgencia: "Normal",
    solicitado_por: "", fecha_necesaria: "", observaciones: "",
  });
  const [items, setItems] = useState([blank()]);
  const [saving, setSaving] = useState(false);
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const setItem = (i, k, v) => { const its = [...items]; its[i] = { ...its[i], [k]: v }; setItems(its); };
  const bases = BASES_POR_EMPRESA["Parana Logistica"] || [];
  const areas = AREAS_POR_EMPRESA["Parana Logistica"] || [];
  const subareas = SUBAREA_TECNICA["Parana Logistica"] || [];
  const detalles = DETALLE_TECNICO[form.subarea] || [];

  const handleSubmit = async () => {
    if (!form.titulo || !form.base_buque || !form.area || !form.solicitado_por) return alert("Completá: Título, Base/Buque, Área, Solicitado por");
    if (!items.some(i => i.descripcion.trim())) return alert("Agregá al menos un ítem");
    setSaving(true);
    try {
      const cleanItems = items.filter(i => i.descripcion.trim()).map(({ id: _id, ...rest }) => rest);
      await onSave({ ...form }, cleanItems);
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
        <FG label="Empresa"><input value="Parana Logistica" readOnly style={{ background: "var(--surface2)", color: "var(--muted)" }} /></FG>
        <FG label="Base / Buque *"><select value={form.base_buque} onChange={e => set("base_buque", e.target.value)}><option value="">Seleccionar...</option>{bases.map(b => <option key={b}>{b}</option>)}</select></FG>
        <FG label="Área *"><select value={form.area} onChange={e => { set("area", e.target.value); set("subarea", ""); }}><option value="">Seleccionar...</option>{areas.map(a => <option key={a}>{a}</option>)}</select></FG>
      </div>
      <div className="form-grid">
        <FG label="Sub-área *">
          <select value={form.subarea} onChange={e => { set("subarea", e.target.value); set("detalle_tecnico", ""); }}>
            <option value="">Seleccionar...</option>
            {subareas.map(s => <option key={s}>{s}</option>)}
          </select>
        </FG>
        {detalles.length > 0 && <FG label="Detalle técnico"><select value={form.detalle_tecnico} onChange={e => set("detalle_tecnico", e.target.value)}><option value="">Seleccionar...</option>{detalles.map(d => <option key={d}>{d}</option>)}</select></FG>}
      </div>
      <div className="form-grid">
        <FG label="Solicitado por *"><input value={form.solicitado_por} onChange={e => set("solicitado_por", e.target.value)} /></FG>
        <FG label="Fecha necesaria"><input type="date" value={form.fecha_necesaria} onChange={e => set("fecha_necesaria", e.target.value)} /></FG>
      </div>
      <FG label="Urgencia *"><select value={form.urgencia} onChange={e => set("urgencia", e.target.value)}>{URGENCIA_OPTIONS.map(u => <option key={u}>{u}</option>)}</select></FG>
      <FG label="Observaciones" full><textarea value={form.observaciones} onChange={e => set("observaciones", e.target.value)} style={{ marginTop: 8 }} /></FG>

      <div className="form-section mt16">Ítems</div>
      <div className="table-wrap">
        <table className="items-edit">
          <thead><tr><th style={{ width: "40%" }}>Descripción *</th><th>Cant.</th><th>Unid.</th><th>Proveedor sugerido</th><th></th></tr></thead>
          <tbody>
            {items.map((it, i) => <tr key={it.id || i}>
              <td><input value={it.descripcion} onChange={e => setItem(i, "descripcion", e.target.value)} /></td>
              <td><input type="number" value={it.cantidad} onChange={e => setItem(i, "cantidad", e.target.value)} style={{ width: 55 }} /></td>
              <td><input value={it.unidad} onChange={e => setItem(i, "unidad", e.target.value)} style={{ width: 50 }} /></td>
              <td><select value={it.proveedor_sugerido || ""} onChange={e => setItem(i, "proveedor_sugerido", e.target.value)}><option value="">Sin sugerencia</option>{proveedores.map(p => <option key={p.id} value={p.nombre}>{p.nombre}</option>)}</select></td>
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

function PageNueva({ onSaved, onCancel, notify }) {
  const [proveedores, setProveedores] = useState([]);
  useEffect(() => { api.getProveedores().then(setProveedores); }, []);
  const handleSave = async (form, items) => {
    await api.crearRequisicion(form, items);
    notify("Requisición creada — pendiente de aprobación", "success");
    onSaved();
  };
  return (
    <div className="card">
      <div className="card-title">Nueva Requisición</div>
      <ReqForm proveedores={proveedores} onSave={handleSave} onCancel={onCancel} />
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
  const bySol = {};
  reqs.forEach(r => { if (!bySol[r.solicitado_por]) bySol[r.solicitado_por] = { total: 0, criticas: 0, devueltas: 0 }; bySol[r.solicitado_por].total++; if (r.urgencia === "Critica") bySol[r.solicitado_por].criticas++; if (r.veces_devuelto > 0) bySol[r.solicitado_por].devueltas++; });
  const byRechazo = {};
  reqs.filter(r => r.motivo_rechazo_categoria).forEach(r => { byRechazo[r.motivo_rechazo_categoria] = (byRechazo[r.motivo_rechazo_categoria] || 0) + 1; });
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
          <div className="card-title">Motivos de rechazo</div>
          {Object.keys(byRechazo).length === 0 ? <div style={{ fontSize: 12, color: "var(--muted)" }}>Sin rechazos</div> :
            Object.entries(byRechazo).sort((a, b) => b[1] - a[1]).map(([cat, n]) => <div key={cat} className="kbar">
              <div className="kbar-lbl"><span style={{ color: "var(--muted)" }}>{cat}</span><span className="text-mono">{n}</span></div>
              <div className="kbar-track"><div className="kbar-fill" style={{ width: `${n / Math.max(...Object.values(byRechazo)) * 100}%`, background: "var(--danger)" }} /></div>
            </div>)}
        </div>
      </div>
    </div>
  );
}

// ─── PAGE: PROVEEDORES ────────────────────────────────────────────────────────
function PageProveedores({ notify }) {
  const [provs, setProvs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(false);
  const [selected, setSelected] = useState(null);
  const [historial, setHistorial] = useState([]);
  const [form, setForm] = useState({ nombre: "", rubro: "", contacto: "", email: "", telefono: "", notas: "", palabras_clave: "" });

  useEffect(() => { api.getProveedores().then(d => { setProvs(d); setLoading(false); }); }, []);

  const handleSave = async () => {
    if (!form.nombre) return;
    const nuevo = await api.crearProveedor({ ...form, activo: true });
    setProvs(p => [...p, nuevo]);
    setModal(false);
    setForm({ nombre: "", rubro: "", contacto: "", email: "", telefono: "", notas: "", palabras_clave: "" });
    notify("Proveedor agregado", "success");
  };

  const handleSelect = async (prov) => {
    setSelected(prov);
    const lineas = await api.getTrackerLineas({ proveedor: prov.nombre });
    setHistorial(lineas.filter(l => l.costo_real || l.nro_oc));
  };

  return (
    <div>
      {selected ? (
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 18 }}>
            <button className="btn btn-ghost btn-sm" onClick={() => setSelected(null)}>← Volver</button>
            <div style={{ fontSize: 16, fontWeight: 700 }}>{selected.nombre}</div>
            {selected.rubro && <span className="tag">{selected.rubro}</span>}
          </div>
          <div className="card">
            <div className="card-title">Historial de compras</div>
            {historial.length === 0 ? <div style={{ fontSize: 12, color: "var(--muted)" }}>Sin compras registradas</div> :
              <table>
                <thead><tr><th>REQ</th><th>Descripción</th><th>OC</th><th>Precio</th><th>Entrega</th></tr></thead>
                <tbody>
                  {historial.map(l => <tr key={l.id}>
                    <td className="text-mono" style={{ fontSize: 11 }}>{l.requisiciones ? `REQ-${String(l.requisiciones.nro_solicitud).padStart(4, "0")}` : "—"}</td>
                    <td style={{ fontWeight: 500, fontSize: 12 }}>{l.descripcion}</td>
                    <td className="text-mono" style={{ fontSize: 11, color: "var(--accent2)" }}>{l.nro_oc || "—"}</td>
                    <td className="text-mono" style={{ fontSize: 12 }}>{l.costo_real ? fmt(l.costo_real, l.moneda_real) : "—"}</td>
                    <td style={{ fontSize: 11, color: "var(--muted)" }}>{fmtDate(l.fecha_entrega_real)}</td>
                  </tr>)}
                </tbody>
              </table>
            }
          </div>
        </div>
      ) : (
        <div className="card">
          <div className="card-title">Maestro de proveedores <button className="btn btn-primary btn-sm" onClick={() => setModal(true)}>+ Agregar</button></div>
          {loading ? <div className="loading"><span className="spin">◌</span></div> :
            <table>
              <thead><tr><th>Nombre</th><th>Rubro</th><th>Contacto</th><th>Email</th><th></th></tr></thead>
              <tbody>
                {provs.map(p => <tr key={p.id} className="click" onClick={() => handleSelect(p)}>
                  <td style={{ fontWeight: 600 }}>{p.nombre}</td>
                  <td className="text-muted">{p.rubro || "—"}</td>
                  <td>{p.contacto || "—"}</td>
                  <td className="text-mono" style={{ fontSize: 11 }}>{p.email || "—"}</td>
                  <td><span style={{ fontSize: 11, color: "var(--blue)" }}>Ver →</span></td>
                </tr>)}
                {!provs.length && <tr><td colSpan={5}><div className="empty-state">Sin proveedores</div></td></tr>}
              </tbody>
            </table>
          }
        </div>
      )}
      {modal && <div className="overlay" onClick={e => e.target === e.currentTarget && setModal(false)}>
        <div className="modal" style={{ maxWidth: 560 }}>
          <div className="mhdr"><div className="mtitle">Nuevo Proveedor</div><button className="mclose" onClick={() => setModal(false)}>✕</button></div>
          <div className="mbody">
            <div className="form-grid">
              <FG label="Nombre *"><input value={form.nombre} onChange={e => setForm(f => ({ ...f, nombre: e.target.value }))} /></FG>
              <FG label="Rubro"><input value={form.rubro} onChange={e => setForm(f => ({ ...f, rubro: e.target.value }))} /></FG>
              <FG label="Contacto"><input value={form.contacto} onChange={e => setForm(f => ({ ...f, contacto: e.target.value }))} /></FG>
              <FG label="Email"><input value={form.email} onChange={e => setForm(f => ({ ...f, email: e.target.value }))} /></FG>
              <FG label="Teléfono"><input value={form.telefono} onChange={e => setForm(f => ({ ...f, telefono: e.target.value }))} /></FG>
            </div>
            <FG label="Palabras clave" hint="Separadas por coma"><input value={form.palabras_clave} onChange={e => setForm(f => ({ ...f, palabras_clave: e.target.value }))} /></FG>
            <FG label="Notas"><textarea value={form.notas} onChange={e => setForm(f => ({ ...f, notas: e.target.value }))} /></FG>
          </div>
          <div className="mftr">
            <button className="btn btn-ghost" onClick={() => setModal(false)}>Cancelar</button>
            <button className="btn btn-primary" onClick={handleSave}>Guardar</button>
          </div>
        </div>
      </div>}
    </div>
  );
}

// ─── ROOT APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("inbox-aprobacion");
  const [notif, setNotif] = useState(null);
  const [counts, setCounts] = useState({ aprobacion: 0, cotizar: 0, confirmacion: 0, tracker: 0 });
  const [refreshKey, setRefreshKey] = useState(0);

  const notify = useCallback((text, type = "info") => {
    setNotif({ text, type }); setTimeout(() => setNotif(null), 4000);
  }, []);

  const loadCounts = useCallback(async () => {
    try {
      const [reqs, tracker] = await Promise.all([
        api.getRequisiciones({ empresa: "Parana Logistica", statuses: ["pendiente_aprobacion", "aprobado_cotizar", "pendiente_confirmacion"] }),
        api.getTrackerLineas({ statuses: ["en_cotizacion", "oc_emitida", "en_transito"] })
      ]);
      setCounts({
        aprobacion: reqs.filter(r => r.status === "pendiente_aprobacion").length,
        cotizar: reqs.filter(r => r.status === "aprobado_cotizar").length,
        confirmacion: reqs.filter(r => r.status === "pendiente_confirmacion").length,
        tracker: tracker.length,
      });
    } catch (e) { console.error(e); }
  }, []);

  useEffect(() => { loadCounts(); }, [loadCounts, refreshKey]);

  const pageTitles = {
    "inbox-aprobacion": "PENDIENTES DE APROBACIÓN",
    "inbox-cotizar": "APROBADOS PARA COTIZAR",
    "inbox-confirmacion": "PENDIENTES CONFIRMACIÓN DE VALOR",
    "tracker": "TRACKER — COMPRAS EN CURSO",
    "tracker-simple": "SEGUIMIENTO DE PEDIDOS",
    "archivo-entregados": "ARCHIVO — ENTREGADOS",
    "archivo-rechazados": "ARCHIVO — RECHAZADOS",
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
              <img src="/pL.png" alt="Parana Logística" className="sidebar-logo-img" />
              <div>
                <div className="sidebar-logo-main">Compras Técnicas</div>
                <div className="sidebar-logo-sub">Parana Logística</div>
              </div>
            </div>
          </div>

          <div className="nav-section">Inbox</div>
          <NI id="inbox-aprobacion" icon="⏳" label="Pend. aprobación" badge={counts.aprobacion} />
          <NI id="inbox-cotizar" icon="📥" label="Para cotizar" badge={counts.cotizar} badgeColor="amber" />
          <NI id="inbox-confirmacion" icon="🔁" label="Conf. de valor" badge={counts.confirmacion} badgeColor="amber" />

          <div className="nav-section">Tracker</div>
          <NI id="tracker" icon="📊" label="Compras en curso" badge={counts.tracker} badgeColor="gray" />
          <NI id="tracker-simple" icon="👁" label="Seguimiento" sub />

          <div className="nav-section">Archivo</div>
          <NI id="archivo-entregados" icon="✓" label="Entregados" sub />
          <NI id="archivo-rechazados" icon="✗" label="Rechazados" sub />

          <div className="nav-section">Gestión</div>
          <NI id="nueva" icon="✚" label="Nueva Requisición" />
          <NI id="kpis" icon="📈" label="KPIs & Reportes" />
          <NI id="proveedores" icon="🏭" label="Proveedores" />

          <div style={{ flex: 1 }} />
          <div style={{ padding: "12px 18px", borderTop: "1px solid rgba(255,255,255,.1)" }}>
            <div className="ni back" onClick={() => window.open(PORTAL_URL, "_self")}>
              <span className="ni-icon">←</span>
              <span>Volver al portal</span>
            </div>
            <div style={{ fontSize: 9, color: "rgba(255,255,255,.3)", fontFamily: "var(--mono)", letterSpacing: 1, marginTop: 8 }}>COMPRAS TÉCNICAS v3.1</div>
          </div>
        </nav>

        <div className="main">
          <div className="topbar">
            <div className="topbar-title">{pageTitles[page] || page}</div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 28, height: 28, borderRadius: "50%", background: "#DBEAFE", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: "var(--blue)", fontWeight: 700 }}>C</div>
              <span style={{ fontSize: 12, color: "var(--muted)", fontWeight: 500 }}>{USUARIO}</span>
            </div>
          </div>
          <div className="content">
            {page === "inbox-aprobacion" && <PageInbox statuses={["pendiente_aprobacion"]} titulo="Pendientes de aprobación" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "inbox-cotizar" && <PageInbox statuses={["aprobado_cotizar"]} titulo="Aprobados para cotizar" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "inbox-confirmacion" && <PageInbox statuses={["pendiente_confirmacion"]} titulo="Pendientes confirmación de valor" notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "tracker" && <PageTrackerGeneral key={`tg-${refreshKey}`} notify={notify} onNeedRefresh={() => { setRefreshKey(k => k + 1); loadCounts(); }} />}
            {page === "tracker-simple" && <PageTrackerSimple />}
            {page === "archivo-entregados" && <PageArchivo tipo="entregados" />}
            {page === "archivo-rechazados" && <PageArchivo tipo="rechazados" />}
            {page === "nueva" && <PageNueva onSaved={() => { setPage("inbox-aprobacion"); loadCounts(); }} onCancel={() => setPage("inbox-aprobacion")} notify={notify} />}
            {page === "kpis" && <PageKPIs />}
            {page === "proveedores" && <PageProveedores notify={notify} />}
          </div>
        </div>
      </div>
      <Notif msg={notif} onClose={() => setNotif(null)} />
    </>
  );
}
