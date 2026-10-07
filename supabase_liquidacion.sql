-- Ejecutar en Supabase > SQL Editor (se puede repetir sin problema)
-- Liquidaciones: cada archivo descargado queda registrado y sus registros pasan a "Enviado".
-- ATENCIÓN: sin login, acceso abierto. Al activar el login, cambiar "to anon, authenticated" por "to authenticated".

-- 1) Liquidaciones generadas (un archivo Excel = una liquidación)
create table if not exists liquidaciones (
  id uuid primary key default gen_random_uuid(),
  numero bigint generated always as identity unique,
  fecha date not null default current_date,
  modo_pago text,
  total numeric not null default 0,
  items int not null default 0,
  anulada boolean not null default false,
  creado timestamptz not null default now()
);
alter table liquidaciones enable row level security;
drop policy if exists "auth liquidaciones" on liquidaciones;
create policy "auth liquidaciones" on liquidaciones for all to anon, authenticated using (true) with check (true);

-- 2) Cada registro enviado apunta a su liquidación
alter table envios add column if not exists liquidacion_id uuid references liquidaciones(id) on delete set null;
create index if not exists envios_liquidacion_idx on envios (liquidacion_id);

-- 3) Datos fijos del formato (liquidador, área, centro de costo, límite por archivo...)
create table if not exists parametros (
  clave text primary key,
  valor text not null,
  updated_at timestamptz not null default now()
);
alter table parametros enable row level security;
drop policy if exists "auth parametros" on parametros;
create policy "auth parametros" on parametros for all to anon, authenticated using (true) with check (true);

notify pgrst, 'reload schema';
