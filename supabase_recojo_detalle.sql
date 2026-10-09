-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Detalle de recojos (devoluciones): una fila por pedido recogido, agrupada por fecha_reporte.
-- ATENCIÓN: sin login, acceso abierto. Al activar el login, cambiar "to anon, authenticated" por "to authenticated".
create table if not exists recojo_detalle (
  id bigint generated always as identity primary key,
  fecha_reporte date not null,
  nombre_cuenta text,
  cuenta text,
  nro_pedido text,
  nro_referencia text,
  cliente_final text,
  fecha_pedido date,
  bultos numeric,
  motivo_devolucion text,
  nodo text,
  fecha_solicitud_cx date,
  fecha_solicitud_nodo date,
  fecha_salida_nodo date,
  fecha_recojo_agencia date,
  fecha_llegada_ctd date,
  lt_retorno numeric,
  medio text,
  proveedor text,
  voucher text,
  guia text,
  clave text,
  estatus text,
  created_at timestamptz not null default now()
);
create index if not exists recojo_detalle_fecha_nodo on recojo_detalle (fecha_reporte, nodo);
create index if not exists recojo_detalle_pedido on recojo_detalle (nro_pedido);
create index if not exists recojo_detalle_guia on recojo_detalle (guia);
alter table recojo_detalle enable row level security;
drop policy if exists "auth recojo_detalle" on recojo_detalle;
create policy "auth recojo_detalle" on recojo_detalle for all to anon, authenticated using (true) with check (true);
