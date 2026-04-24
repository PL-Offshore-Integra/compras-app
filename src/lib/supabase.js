import { createClient } from "@supabase/supabase-js";

// Estas variables se configuran en .env.local (ver README)
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ── REQUISICIONES ────────────────────────────────────────────────────────────

export async function getRequisiciones(filtros = {}) {
  let q = supabase
    .from("requisiciones")
    .select("*, requisicion_items(*)")
    .order("created_at", { ascending: false });

  if (filtros.status)   q = q.eq("status", filtros.status);
  if (filtros.empresa)  q = q.eq("empresa", filtros.empresa);
  if (filtros.urgencia) q = q.eq("urgencia", filtros.urgencia);

  const { data, error } = await q;
  if (error) throw error;
  return data;
}

export async function getRequisicion(id) {
  const { data, error } = await supabase
    .from("requisiciones")
    .select("*, requisicion_items(*), requisicion_historial(*)")
    .eq("id", id)
    .single();
  if (error) throw error;
  return data;
}

export async function crearRequisicion(req, items, usuario) {
  // Insertar cabecera
  const { data: nueva, error } = await supabase
    .from("requisiciones")
    .insert([{ ...req, status: "pendiente_revision" }])
    .select()
    .single();
  if (error) throw error;

  // Insertar ítems
  if (items?.length) {
    const { error: errItems } = await supabase
      .from("requisicion_items")
      .insert(items.map((it, i) => ({ ...it, requisicion_id: nueva.id, nro_linea: i + 1 })));
    if (errItems) throw errItems;
  }

  // Historial
  await agregarHistorial(nueva.id, "Requisición creada", usuario, null, "pendiente_revision");
  return nueva;
}

export async function actualizarRequisicion(id, cambios, usuario, evento) {
  const { data: anterior } = await supabase.from("requisiciones").select("status").eq("id", id).single();

  const { data, error } = await supabase
    .from("requisiciones")
    .update(cambios)
    .eq("id", id)
    .select()
    .single();
  if (error) throw error;

  if (evento) {
    await agregarHistorial(id, evento, usuario, anterior?.status, cambios.status || anterior?.status);
  }
  return data;
}

export async function actualizarItems(requisicionId, items) {
  // Borrar existentes y reinsertar (simple para MVP)
  await supabase.from("requisicion_items").delete().eq("requisicion_id", requisicionId);
  if (items?.length) {
    const { error } = await supabase
      .from("requisicion_items")
      .insert(items.map((it, i) => ({ ...it, requisicion_id: requisicionId, nro_linea: i + 1 })));
    if (error) throw error;
  }
}

// ── HISTORIAL ────────────────────────────────────────────────────────────────

export async function agregarHistorial(requisicionId, evento, usuario, statusAnterior, statusNuevo, detalle) {
  await supabase.from("requisicion_historial").insert([{
    requisicion_id: requisicionId,
    evento,
    usuario: usuario || "Sistema",
    status_anterior: statusAnterior,
    status_nuevo: statusNuevo,
    detalle,
  }]);
}

// ── PROVEEDORES ──────────────────────────────────────────────────────────────

export async function getProveedores() {
  const { data, error } = await supabase
    .from("proveedores")
    .select("*")
    .eq("activo", true)
    .order("nombre");
  if (error) throw error;
  return data || [];
}

export async function crearProveedor(prov) {
  const { data, error } = await supabase.from("proveedores").insert([prov]).select().single();
  if (error) throw error;
  return data;
}

// ── KPIs ─────────────────────────────────────────────────────────────────────

export async function getKPIs() {
  const { data, error } = await supabase.from("vw_kpis").select("*").single();
  if (error) throw error;
  return data;
}

export async function getKPIsPorSolicitante() {
  const { data, error } = await supabase
    .from("requisiciones")
    .select("solicitado_por, urgencia, status, veces_devuelto, desvio_presupuestario, dias_solicitud_revision");
  if (error) throw error;
  return data || [];
}

export async function getKPIsPorProveedor() {
  const { data, error } = await supabase
    .from("requisiciones")
    .select("proveedor_elegido, costo_real, moneda_real, dias_aprobacion_entrega, fecha_entrega_prom, fecha_entrega_real")
    .not("proveedor_elegido", "is", null);
  if (error) throw error;
  return data || [];
}
