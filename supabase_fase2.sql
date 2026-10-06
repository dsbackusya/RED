-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Fase 2 de KPIs: metas mensuales + vista agregada del HRN.
-- ATENCIÓN: sin login, acceso abierto. Al activar el login, cambiar "to anon, authenticated" por "to authenticated".

-- 1) Metas mensuales de costo
create table if not exists metas (
  mes text primary key check (mes ~ '^\d{4}-\d{2}$'),
  monto numeric not null check (monto > 0),
  updated_at timestamptz not null default now()
);
alter table metas enable row level security;
drop policy if exists "auth metas" on metas;
create policy "auth metas" on metas for all to anon, authenticated using (true) with check (true);

-- 2) HRN agregado por fecha, nodo y cuenta: la app descarga estas filas en vez del detalle completo
create or replace view hrn_grupos with (security_invoker = true) as
select fecha_reporte, nodo, descripcion_cuenta,
       count(distinct pedido_cliente)::int as pedidos,
       coalesce(sum(bultos), 0) as bultos,
       coalesce(sum(peso), 0) as peso,
       coalesce(sum(volumen), 0) as volumen
from hrn_detalle
group by fecha_reporte, nodo, descripcion_cuenta;
grant select on hrn_grupos to anon, authenticated;

notify pgrst, 'reload schema';
