-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Detalle de la hoja ( HRN ): una fila por bulto/LPN, agrupada por fecha_reporte.
-- ATENCIÓN: sin login, acceso abierto. Al activar el login, cambiar "to anon, authenticated" por "to authenticated".
create table if not exists hrn_detalle (
  id bigint generated always as identity primary key,
  fecha_reporte date not null,
  pedido_cliente text,
  nro_doc_referencia text,
  tipo_documento text,
  lpn text,
  fecha_recepcion timestamp,
  ubicacion_recepcion text,
  cuenta_procedencia text,
  descripcion_cuenta text,
  nodo text,
  bultos numeric,
  observacion text,
  fecha_ruta timestamp,
  flag_carga_stage numeric,
  fecha_carga_stage timestamp,
  flag_terminar_carga numeric,
  fecha_terminar_carga timestamp,
  fecha_despacho timestamp,
  flag_reprogramado numeric,
  contador_reprogramado numeric,
  motivo_pedido text,
  alto numeric,
  ancho numeric,
  largo numeric,
  volumen numeric,
  peso numeric,
  roll_contenedor text,
  codigo_zona text,
  zona text,
  orden_clasificacion text,
  estado_contenedor text,
  placa_real text,
  fecha_pedido timestamp,
  fecha_creacion timestamp,
  ubicacion_movimiento text,
  fecha_movimiento timestamp,
  usuario_movimiento text,
  proveedor text,
  nodo2 text,
  cant_pedidos text,
  created_at timestamptz not null default now()
);
create index if not exists hrn_detalle_fecha_nodo on hrn_detalle (fecha_reporte, nodo);
create index if not exists hrn_detalle_pedido on hrn_detalle (pedido_cliente);
create index if not exists hrn_detalle_lpn on hrn_detalle (lpn);
alter table hrn_detalle enable row level security;
drop policy if exists "auth hrn_detalle" on hrn_detalle;
create policy "auth hrn_detalle" on hrn_detalle for all to anon, authenticated using (true) with check (true);
